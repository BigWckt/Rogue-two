#!/usr/bin/env python3
"""
script_pages_jaunes.py — Scraping Pages Jaunes
===============================================
Collecte des entreprises sur Pages Jaunes par activité et ville,
via Playwright (headless Chromium) avec bypass Cloudflare Turnstile.

Technique : warmup sur une URL PJ bidon, attente adaptative du titre
(pas de sleep fixe), puis scraping des résultats page par page.

Ref : README_scraper_pj.md — issu du scraper en production (3 329 fiches).

Usage :
  python script_pages_jaunes.py --ville Paris --activite "boulangerie" --nb-max 100
  python script_pages_jaunes.py --config params.json
"""

import argparse
import csv
import json
import os
import random
import re
import subprocess
import sys
import time
import unicodedata
from datetime import date
from urllib.parse import quote, urlparse

import pandas as pd

from batch_io import read_current_batch, check_enseigne_excluded
from matrix_display import (
    GREEN, RED, BOLD, RESET,
    matrix_banner, matrix_section, matrix_kv, matrix_separator,
    matrix_step, matrix_ok, matrix_fail, matrix_warn,
    morpheus_says, ask_filename, matrix_decode, set_quiet,
    matrix_rain_fullscreen,
)

# ── User-Agent pool ──────────────────────────────────────────────────────────

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:125.0) Gecko/20100101 Firefox/125.0",
]

# ── Colonnes de sortie (cohérentes avec script_lba_lbb.py) ───────────────────

OUTPUT_COLUMNS = [
    "Nom de l'entreprise",
    "Adresse",
    "Ville",
    "Code Postal",
    "Ville de recherche",
    "Téléphone",
    "Site web",
    "Catégorie",
    "URL fiche PJ",
    "Date de collecte",
]

# ── Exclusions automatiques ──────────────────────────────────────────────────
# Mots-clés d'exclusion par catégorie (insensible casse + accents).
# Étendre cette constante pour ajouter de nouvelles exclusions.

EXCLUSIONS_PJ = {
    "Pharmacies": [
        "pharmacie", "pharmacien", "parapharmacie",
    ],
    "Centres d'hébergement": [
        "centre d'hebergement", "centre d hebergement", "hebergement d'urgence",
        "chrs", "cada", "huda", "centre d'accueil",
    ],
    "Foyers jeunes": [
        "foyer de jeunes", "foyer jeunes travailleurs", "fjt",
        "foyer pour jeunes", "foyer d'accueil",
    ],
    "Centres de planification": [
        "centre de planification", "planification familiale", "cpef",
        "planning familial",
    ],
}

# ── Config ────────────────────────────────────────────────────────────────────

PJ_BASE_URL = "https://www.pagesjaunes.fr"
SAVE_EVERY = 50
MAX_EMPTY_PAGES = 3
DELAY_AFTER_NAV = 12       # secondes d'attente après chaque navigation
DELAY_SAME_KEYWORD = 5    # secondes entre deux pages du même mot-clé
DELAY_BETWEEN_KEYWORDS = 8  # secondes entre deux mots-clés différents
CF_POLL_INTERVAL = 5       # secondes entre chaque vérif du titre CF
CF_TIMEOUT = 90            # timeout max pour le challenge CF


# ── Helpers ───────────────────────────────────────────────────────────────────

def extract_cp_ville(adresse: str) -> tuple[str, str]:
    """Extrait code postal et ville depuis une adresse texte."""
    if not adresse:
        return "", ""
    m = re.search(r"(\d{5})\s+([A-Za-zÀ-ÿ\s\-']+)", adresse)
    if m:
        return m.group(1), m.group(2).strip().title()
    return "", ""


def validate_phone(raw: str) -> str:
    """Valide un numéro de téléphone français. Retourne le numéro nettoyé ou ''."""
    digits = re.sub(r"\D", "", raw)
    if digits and re.match(r"^0[1-9]\d{8}$", digits):
        return digits
    return ""


def strip_accents(s: str) -> str:
    """Supprime les accents d'une chaîne (é→e, è→e, etc.)."""
    return unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode()


def apply_exclusions(fiches: list[dict]) -> tuple[list[dict], dict[str, int]]:
    """
    Filtre les fiches selon EXCLUSIONS_PJ (nom + catégorie, insensible casse/accents).
    Retourne (fiches_filtrées, compteurs_par_catégorie).
    """
    counts = {cat: 0 for cat in EXCLUSIONS_PJ}
    filtered = []
    for fiche in fiches:
        nom = strip_accents(fiche.get("Nom de l'entreprise", "").lower())
        cat = strip_accents(fiche.get("Catégorie", "").lower())
        text = f"{nom} {cat}"
        excluded = False
        for category, keywords in EXCLUSIONS_PJ.items():
            for kw in keywords:
                if kw in text:
                    counts[category] += 1
                    excluded = True
                    break
            if excluded:
                break
        if not excluded:
            filtered.append(fiche)
    return filtered, counts


def apply_exclusions_enseignes(fiches: list[dict]) -> tuple[list[dict], dict[str, int]]:
    """
    Filtre les enseignes nationales (MBT/TG uniquement).
    Retourne (fiches_filtrées, compteurs_par_enseigne).
    """
    counts: dict[str, int] = {}
    filtered = []
    for fiche in fiches:
        nom = fiche.get("Nom de l'entreprise", "")
        matched = check_enseigne_excluded(nom)
        if matched:
            counts[matched] = counts.get(matched, 0) + 1
        else:
            filtered.append(fiche)
    return filtered, counts


def find_chromium():
    """Retourne None — Playwright détecte automatiquement le binaire Chromium."""
    return None


def build_search_url(activite: str, ville: str, page: int) -> str:
    return (
        f"{PJ_BASE_URL}/annuaire/chercherlespros"
        f"?quoiqui={quote(activite)}&ou={quote(ville)}&page={page}"
    )


# ── Attente adaptative Cloudflare ────────────────────────────────────────────

def wait_for_pj_content(page, timeout_s: int = CF_TIMEOUT) -> bool:
    """
    Attend que le challenge Cloudflare se résolve.
    Vérifie toutes les 5s si le titre ne contient plus "un instant" / "cloudflare".
    Retourne True si le contenu est chargé, False si timeout.
    """
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        title = (page.title() or "").lower()
        if "un instant" not in title and "cloudflare" not in title:
            return True
        time.sleep(CF_POLL_INTERVAL)
    return False


# ── Révélation téléphone (clic « Afficher le N° ») ─────────────────────────

REVEAL_SELECTORS = [
    "button.btn_tel",
    "button[class*='btn_tel']",
    "[class*='btn_tel']",
    "[class*='click-to-call']",
    "[class*='bi-click-to']",
    "button[class*='num']",
    "a[class*='num']",
    "[data-pjlb*='tel']",
    "[data-pjlb*='click']",
    "[class*='show-phone']",
    "[class*='phone-reveal']",
    "[class*='bi-phone-number'] button",
    ".bi-fantomas button",
    ".bi-fantomas a[href^='tel:']",
    ".bi-fantomas a",
]

PHONE_READ_SELECTORS = [
    ".bi-fantomas .number-contact",
    "[class*='number-contact']",
    "[class*='phone-number']",
    ".bi-fantomas [class*='phone']",
    ".bi-fantomas [class*='tel']",
    "a[href^='tel:']",
]

REVEAL_THROTTLE_MAX = 50
REVEAL_THROTTLE_PAUSE = 10


def _read_phone_from_bloc(bloc) -> str:
    """Tente de lire le téléphone depuis le bloc avec plusieurs sélecteurs."""
    for sel in PHONE_READ_SELECTORS:
        tel_el = bloc.query_selector(sel)
        if not tel_el:
            continue
        raw = ""
        if sel.startswith("a[href^='tel:']"):
            raw = (tel_el.get_attribute("href") or "").replace("tel:", "")
        else:
            raw = tel_el.inner_text().strip()
        raw = re.sub(r"^T[ée]l\s*:\s*", "", raw)
        validated = validate_phone(raw)
        if validated:
            return validated

    # Fallback : le bouton btn_tel est souvent HORS du <li> ; chercher dans
    # les frères et l'ancêtre .bi-bloc. Après clic, son texte peut contenir
    # le numéro révélé.
    try:
        btn_handle = bloc.evaluate_handle(
            "el => el.nextElementSibling?.querySelector('.btn_tel, [class*=\"btn_tel\"]') || "
            "el.closest('[class*=\"bi-bloc\"]')?.querySelector('.btn_tel, [class*=\"btn_tel\"]') || "
            "el.parentElement?.querySelector('.btn_tel, [class*=\"btn_tel\"]')"
        )
        if btn_handle:
            el = btn_handle.as_element()
            if el:
                raw = el.inner_text().strip()
                raw = re.sub(r"^T[ée]l\s*:\s*", "", raw)
                validated = validate_phone(raw)
                if validated:
                    return validated
    except Exception:
        pass

    return ""


def _try_reveal_phone(bloc, click_state: dict, debug: bool = False) -> str:
    """
    Cherche et clique un bouton de révélation téléphone, puis re-lit le numéro.
    click_state suit le throttling {count, window_start}.
    Le bouton btn_tel est souvent HORS du <li> scrappé — on étend la recherche
    aux frères et à l'ancêtre .bi-bloc.
    """
    btn = None
    found_scope = "bloc"

    # Étape 1 : chercher dans le bloc lui-même
    for sel in REVEAL_SELECTORS:
        btn = bloc.query_selector(sel)
        if btn:
            break

    # Étape 2 : étendre la recherche hors du <li>
    if not btn:
        try:
            btn_handle = bloc.evaluate_handle(
                "el => el.nextElementSibling?.querySelector('.btn_tel, [class*=\"btn_tel\"], [class*=\"click-to-call\"]') || "
                "el.querySelector('.btn_tel, [class*=\"btn_tel\"], [class*=\"click-to-call\"]') || "
                "el.closest('[class*=\"bi-bloc\"]')?.querySelector('.btn_tel, [class*=\"btn_tel\"], [class*=\"click-to-call\"]') || "
                "el.parentElement?.querySelector('.btn_tel, [class*=\"btn_tel\"], [class*=\"click-to-call\"]')"
            )
            if btn_handle:
                el = btn_handle.as_element()
                if el:
                    btn = el
                    found_scope = "sibling/ancestor"
        except Exception:
            pass

    if debug:
        print(f"      🔍 Bouton reveal : {'trouvé' if btn else 'non trouvé'} (scope: {found_scope})")
        if btn:
            try:
                txt = btn.inner_text().strip()[:40]
                print(f"      🔍 Bouton texte avant clic : {txt!r}")
            except Exception:
                pass

    if not btn:
        return ""

    try:
        btn.scroll_into_view_if_needed(timeout=3000)
        time.sleep(random.uniform(0.3, 0.5))
    except Exception:
        pass

    now = time.time()
    if now - click_state.get("window_start", 0) >= 60:
        click_state["window_start"] = now
        click_state["count"] = 0
    click_state["count"] = click_state.get("count", 0) + 1
    if click_state["count"] > REVEAL_THROTTLE_MAX:
        time.sleep(REVEAL_THROTTLE_PAUSE)
        click_state["window_start"] = time.time()
        click_state["count"] = 1

    try:
        btn.click(timeout=3000)
    except Exception:
        return ""

    xhr_resolved = False
    try:
        bloc.wait_for_selector(
            ".bi-fantomas-display .number-contact, .bi-fantomas .number-contact",
            timeout=6000,
            state="visible",
        )
        xhr_resolved = True
    except Exception:
        if debug:
            print("      ⏱  Timeout 6s — AJAX phone_number lent ou non déclenché")

    if debug and xhr_resolved:
        print("      ✓ .number-contact présent dans le DOM après clic")

    # Re-lecture standard (incluant le fallback btn_tel dans _read_phone_from_bloc)
    phone = _read_phone_from_bloc(bloc)
    if phone:
        return phone

    # Le texte du bouton peut avoir été remplacé par le numéro révélé
    try:
        raw = btn.inner_text().strip()
        raw = re.sub(r"^T[ée]l\s*:\s*", "", raw)
        validated = validate_phone(raw)
        if validated:
            if debug:
                print(f"      🔍 Numéro lu depuis le bouton : {validated}")
            return validated
    except Exception:
        pass

    return ""


# ── Parsing d'un bloc entreprise ─────────────────────────────────────────────

def parse_bloc(bloc, click_state: dict | None = None,
               debug: bool = False) -> dict | None:
    """
    Extrait les données d'un bloc entreprise PJ.
    Sélecteurs CSS issus de README_scraper_pj.md (structure DOM PJ vérifiée).
    Cascade : lecture statique → clic de révélation.
    """
    today = date.today().isoformat()

    # Nom : li [class*='bi-denomination'] h3
    name_el = bloc.query_selector("[class*='bi-denomination'] h3")
    nom = name_el.inner_text().strip() if name_el else ""
    if not nom:
        return None

    # URL fiche
    fiche_url = ""
    link_el = bloc.query_selector("[class*='bi-denomination']")
    if link_el and link_el.evaluate("el => el.tagName") == "A":
        href = link_el.get_attribute("href") or ""
        if href.startswith("/"):
            fiche_url = PJ_BASE_URL + href
        elif href.startswith("http"):
            fiche_url = href
    # Fallback: chercher un lien vers /pros/
    if not fiche_url:
        link_el2 = bloc.query_selector("a[href*='/pros/']")
        if link_el2:
            href = link_el2.get_attribute("href") or ""
            if href.startswith("/"):
                fiche_url = PJ_BASE_URL + href
            elif href.startswith("http"):
                fiche_url = href

    # Adresse : li .bi-address
    addr_el = bloc.query_selector(".bi-address")
    adresse = addr_el.inner_text().strip() if addr_el else ""
    cp, ville = extract_cp_ville(adresse)

    # ── Téléphone — Étape A : lecture statique ──
    tel = ""
    phone_source = "vide"
    tel = _read_phone_from_bloc(bloc)
    if tel:
        phone_source = "direct"

    # ── Téléphone — Étape B : clic de révélation ──
    if not tel and click_state is not None:
        try:
            tel = _try_reveal_phone(bloc, click_state, debug=debug)
            if tel:
                phone_source = "via clic"
        except Exception:
            pass

    if debug:
        tag = f"[{phone_source}]"
        print(f"      {GREEN}☎{RESET} {tag:<12s} {tel or 'VIDE'} — {nom[:50]}")

    # Site web (cascade de sélecteurs)
    site = ""
    for wsel in [
        "a[class*='site-internet']", "a[class*='bi-website']", "a.bi-website",
        "a[class*='site']", "a[class*='web']",
        "a[href*='http'][class*='url']",
        "a[data-pjlb*='site']", "a[data-pjlb*='web']",
        "[class*='website'] a",
    ]:
        try:
            web_el = bloc.query_selector(wsel)
            if web_el:
                href = web_el.get_attribute("href") or ""
                if href.startswith("http") and "pagesjaunes.fr" not in href:
                    site = href
                    break
        except Exception:
            continue

    # Catégorie (cascade de sélecteurs)
    categorie = ""
    for csel in [
        ".bi-activite", "[class*='bi-activite']",
        "[class*='activite']", "[class*='rubrique']",
        "[class*='category']", "[class*='metier']",
    ]:
        try:
            cat_el = bloc.query_selector(csel)
            if cat_el:
                text = cat_el.inner_text().strip()
                if text and len(text) < 200:
                    categorie = text
                    break
        except Exception:
            continue

    return {
        "Nom de l'entreprise": nom,
        "Adresse": adresse,
        "Ville": ville,
        "Code Postal": cp,
        "Téléphone": tel,
        "Site web": site,
        "Catégorie": categorie,
        "URL fiche PJ": fiche_url,
        "Date de collecte": today,
        "_phone_source": phone_source,
    }


# ── Sauvegarde intermédiaire ─────────────────────────────────────────────────

def save_intermediate_csv(fiches: list[dict], filepath: str):
    """Sauvegarde CSV intermédiaire pour reprise."""
    try:
        os.makedirs(os.path.dirname(filepath) or ".", exist_ok=True)
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS)
            writer.writeheader()
            writer.writerows(fiches)
    except Exception as e:
        matrix_warn(f"Erreur sauvegarde intermédiaire : {e}")


# ── Scraping principal ───────────────────────────────────────────────────────

def scrape_pages_jaunes(activite: str, ville: str, nb_max: int,
                        output_dir: str, proxy_url: str | None = None,
                        debug: bool = False,
                        deep_phone: bool = False) -> tuple[list[dict], dict]:
    """
    Scrape Pages Jaunes via Playwright avec bypass Cloudflare.
    Retourne (fiches, stats).
    proxy_url : URL proxy explicite (ex: http://user:pass@host:port), None = sortie directe.
    debug : log détaillé par fiche (source téléphone, etc.).
    deep_phone : fallback page de détail pour les fiches sans téléphone.
    """
    from playwright.sync_api import sync_playwright

    stats = {
        "pages_parcourues": 0,
        "fiches_brutes": 0,
        "doublons": 0,
        "blocages_cf": 0,
        "phones_direct": 0,
        "phones_click": 0,
        "phones_deep": 0,
        "phones_missing": 0,
    }

    exec_path = None
    matrix_step("Chromium : auto-détecté par Playwright")

    # Backup CSV path
    today_str = date.today().strftime("%Y%m%d")
    ville_slug = ville.lower().replace(" ", "_").replace("-", "_")
    backup_path = os.path.join(
        output_dir, f"pj_results_{ville_slug}_{today_str}_backup.csv"
    )

    fiches: list[dict] = []
    seen_keys: set[tuple[str, str]] = set()
    consecutive_empty = 0
    click_state: dict = {"count": 0, "window_start": time.time()}

    with sync_playwright() as p:
        launch_args = [
            "--no-sandbox",
            "--disable-blink-features=AutomationControlled",
            "--disable-dev-shm-usage",
            "--ignore-certificate-errors",
        ]

        # Par défaut : sortie directe, ignorer les proxy système
        if not proxy_url:
            launch_args.append("--no-proxy-server")

        launch_kwargs: dict = {
            "headless": True,
            "args": launch_args,
        }

        # Proxy uniquement si fourni explicitement via --proxy
        if proxy_url:
            parsed = urlparse(proxy_url)
            proxy_cfg: dict = {
                "server": f"{parsed.scheme}://{parsed.hostname}:{parsed.port}",
            }
            if parsed.username:
                proxy_cfg["username"] = parsed.username
            if parsed.password:
                proxy_cfg["password"] = parsed.password
            launch_kwargs["proxy"] = proxy_cfg
            matrix_ok(f"Proxy explicite : {proxy_cfg['server']}")
        else:
            matrix_ok("Sortie directe (pas de proxy)")

        browser = p.chromium.launch(**launch_kwargs)
        ua = random.choice(USER_AGENTS)
        context = browser.new_context(
            user_agent=ua,
            locale="fr-FR",
        )
        # Anti-detection appliqué au context (hérité par toutes les pages)
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        page = context.new_page()
        page.set_default_timeout(30_000)

        # ── Warmup CF ────────
        matrix_step("Warmup Cloudflare (fleuriste Paris 20)...")
        try:
            warmup_url = (
                f"{PJ_BASE_URL}/annuaire/chercherlespros"
                "?quoiqui=fleuriste&ou=Paris+20"
            )
            page.goto(warmup_url, wait_until="domcontentloaded", timeout=30_000)
            time.sleep(DELAY_AFTER_NAV)
            if wait_for_pj_content(page, timeout_s=CF_TIMEOUT):
                matrix_ok("Cookie Cloudflare établi")
            else:
                matrix_warn("Timeout CF — le scraping risque d'échouer")
                stats["blocages_cf"] += 1
        except Exception as e:
            matrix_warn(f"Warmup CF : {e}")

        # ── Boucle de pagination ──
        page_num = 1
        while len(fiches) < nb_max:
            url = build_search_url(activite, ville, page_num)
            matrix_step(f"Scraping page {page_num}...")

            try:
                if page_num > 1:
                    time.sleep(DELAY_SAME_KEYWORD)

                page.goto(url, wait_until="domcontentloaded", timeout=30_000)
                time.sleep(DELAY_AFTER_NAV)

                if not wait_for_pj_content(page, timeout_s=CF_TIMEOUT):
                    matrix_warn(f"Timeout CF ({CF_TIMEOUT}s)")
                    consecutive_empty += 1
                    stats["blocages_cf"] += 1
                    if consecutive_empty >= MAX_EMPTY_PAGES:
                        matrix_fail(f"{MAX_EMPTY_PAGES} pages bloquées consécutives — arrêt")
                        break
                    page_num += 1
                    continue

            except Exception as e:
                matrix_fail(f"{e}")
                consecutive_empty += 1
                if consecutive_empty >= MAX_EMPTY_PAGES:
                    matrix_fail(f"{MAX_EMPTY_PAGES} pages échouées consécutives — arrêt")
                    stats["blocages_cf"] = consecutive_empty
                    break
                page_num += 1
                continue

            # Extraire les blocs entreprise
            blocs = page.query_selector_all("li:has(.bi-content)")
            if not blocs:
                blocs = page.query_selector_all(
                    "article.bi-bloc, div[class*='bi-bloc'], li[class*='bi-bloc']"
                )
            if not blocs:
                blocs = page.query_selector_all("[data-bi-name]")

            if not blocs:
                consecutive_empty += 1
                matrix_warn(f"0 fiches (vide {consecutive_empty}/{MAX_EMPTY_PAGES})")
                if consecutive_empty >= MAX_EMPTY_PAGES:
                    matrix_fail(f"{MAX_EMPTY_PAGES} pages vides consécutives — arrêt (blocage CF probable)")
                    stats["blocages_cf"] = consecutive_empty
                    break
                page_num += 1
                continue

            consecutive_empty = 0
            page_count = 0

            for bloc in blocs:
                if len(fiches) >= nb_max:
                    break
                fiche = parse_bloc(bloc, click_state=click_state, debug=debug)
                if fiche is None:
                    continue

                stats["fiches_brutes"] += 1

                # Comptage source téléphone
                src = fiche.pop("_phone_source", "vide")
                if src == "direct":
                    stats["phones_direct"] += 1
                elif src == "via clic":
                    stats["phones_click"] += 1
                else:
                    stats["phones_missing"] += 1

                # Déduplication sur (nom, code postal)
                key = (
                    fiche["Nom de l'entreprise"].lower(),
                    fiche["Code Postal"],
                )
                if key in seen_keys:
                    stats["doublons"] += 1
                    continue
                seen_keys.add(key)

                fiches.append(fiche)
                page_count += 1
                matrix_decode(fiche.get("Nom de l'entreprise", ""), prefix=f"    {GREEN}PJ ▸{RESET} ")

            stats["pages_parcourues"] = page_num
            matrix_ok(f"{page_count} fiches ({len(fiches)} total)")

            # Sauvegarde intermédiaire
            if len(fiches) % SAVE_EVERY < page_count:
                save_intermediate_csv(fiches, backup_path)
                matrix_step(f"Sauvegarde intermédiaire ({len(fiches)} fiches)")

            # Pagination
            has_next = page.query_selector("a#pagination-next")
            if not has_next:
                matrix_ok("Dernière page atteinte")
                break

            page_num += 1

        # ── Étape C : deep-phone (page de détail) ──
        if deep_phone:
            no_phone = [f for f in fiches if not f.get("Téléphone") and f.get("URL fiche PJ")]
            if no_phone:
                matrix_section(f"Deep-phone : {len(no_phone)} fiches sans téléphone")
                for i, fiche in enumerate(no_phone, 1):
                    url = fiche["URL fiche PJ"]
                    nom = fiche.get("Nom de l'entreprise", "")[:40]
                    print(f"    {GREEN}[{i}/{len(no_phone)}]{RESET} {nom}...", end=" ", flush=True)
                    try:
                        page.goto(url, wait_until="domcontentloaded", timeout=30_000)
                        time.sleep(random.uniform(2.5, 4.0))
                        if not wait_for_pj_content(page, timeout_s=30):
                            print(f"{RED}CF{RESET}")
                            continue
                        # Tenter lecture statique sur la page détail
                        tel = ""
                        for sel in PHONE_READ_SELECTORS:
                            el = page.query_selector(sel)
                            if not el:
                                continue
                            raw = el.inner_text().strip()
                            raw = re.sub(r"^T[ée]l\s*:\s*", "", raw)
                            tel = validate_phone(raw)
                            if tel:
                                break
                        # Tenter clic de révélation sur page détail
                        if not tel:
                            for sel in REVEAL_SELECTORS:
                                btn = page.query_selector(sel)
                                if btn:
                                    try:
                                        btn.click(timeout=3000)
                                        time.sleep(random.uniform(1.5, 2.5))
                                    except Exception:
                                        continue
                                    for rsel in PHONE_READ_SELECTORS:
                                        el = page.query_selector(rsel)
                                        if not el:
                                            continue
                                        raw = el.inner_text().strip()
                                        raw = re.sub(r"^T[ée]l\s*:\s*", "", raw)
                                        tel = validate_phone(raw)
                                        if tel:
                                            break
                                    if tel:
                                        break
                        if tel:
                            fiche["Téléphone"] = tel
                            stats["phones_deep"] += 1
                            stats["phones_missing"] -= 1
                            print(f"{GREEN}✓ {tel}{RESET}")
                        else:
                            print(f"{RED}✗{RESET}")
                    except Exception as e:
                        print(f"{RED}ERR{RESET}")
                        if debug:
                            print(f"        {e}")

        browser.close()

    return fiches, stats


# ── Stats téléphone ─────────────────────────────────────────────────────────

def _show_phone_stats(st: dict, n_fiches: int):
    """Affiche les stats de capture téléphone."""
    direct = st.get("phones_direct", 0)
    click = st.get("phones_click", 0)
    deep = st.get("phones_deep", 0)
    missing = st.get("phones_missing", 0)
    total_with = direct + click + deep
    pct = f"{total_with / n_fiches * 100:.0f}%" if n_fiches else "0%"
    matrix_kv("Téléphones", f"{total_with}/{n_fiches} ({pct})")
    if click or deep:
        matrix_kv("  dont direct (DOM)", str(direct))
        matrix_kv("  dont via clic révélation", str(click))
        if deep:
            matrix_kv("  dont via page détail", str(deep))
    if missing:
        matrix_kv("  sans téléphone", str(missing))


# ── Export CSV ───────────────────────────────────────────────────────────────

def export_csv(fiches: list[dict], csv_path: str):
    """Exporte en CSV uniquement."""
    os.makedirs(os.path.dirname(csv_path) or ".", exist_ok=True)
    df = pd.DataFrame(fiches, columns=OUTPUT_COLUMNS)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Scraping Pages Jaunes — collecte entreprises par activité et ville",
    )
    parser.add_argument("--ville", type=str, help="Nom de la ville")
    parser.add_argument("--activite", type=str, help="Type de commerce (ex: boulangerie)")
    parser.add_argument("--nb-max", type=int, default=100,
                        help="Nombre max de résultats (défaut: 100)")
    parser.add_argument("--output", type=str, default=".",
                        help="Répertoire de sortie (défaut: racine repo)")
    parser.add_argument("--config", type=str, help="Fichier JSON de paramètres")
    parser.add_argument("--quiet", action="store_true",
                        help="Mode silencieux (désactive les animations)")
    parser.add_argument("--batch", type=str, default=None,
                        help="Chemin de batch forcé (ex: Prosp/SB/batch_2)")
    parser.add_argument("--proxy", type=str, default=None,
                        help="Proxy explicite (ex: http://user:pass@host:port)")
    parser.add_argument("--debug", action="store_true",
                        help="Mode debug : log détaillé par fiche (source téléphone, etc.)")
    parser.add_argument("--deep-phone", action="store_true",
                        help="Fallback page de détail pour les fiches sans téléphone (lent)")
    return parser.parse_args()


def interactive_menu(args) -> tuple[list[str], str, int, str, str | None]:
    """
    Mode interactif complet (aucun argument CLI).
    Retourne (villes, activite, nb_max, output_dir, output_name).
    """
    print()
    matrix_section("Configuration interactive")

    # 1. Villes
    raw = input(f"    {GREEN}▸{RESET} Ville(s) (séparées par virgules) : ").strip()
    if not raw:
        matrix_fail("Aucune ville saisie")
        sys.exit(1)
    villes = [v.strip() for v in raw.split(",") if v.strip()]

    # 2. Mot-clé de recherche
    activite = input(f"    {GREEN}▸{RESET} Mot-clé de recherche (ex: boulangerie, plombier) : ").strip()
    if not activite:
        matrix_fail("Aucun mot-clé saisi")
        sys.exit(1)

    # 3. Nombre max de fiches
    raw = input(f"    {GREEN}▸{RESET} Nombre max de fiches par ville (défaut 100) : ").strip()
    nb_max = 100
    if raw:
        try:
            nb_max = int(raw)
        except ValueError:
            matrix_warn(f"Valeur invalide « {raw} » — défaut 100")

    # 4. Nom de fichier
    raw = input(f"    {GREEN}▸{RESET} Nom du fichier de sortie (sans extension) [défaut: auto] : ").strip()
    output_name = os.path.basename(raw) if raw else None

    # 5. Résolution output dir
    if args.batch:
        output_dir = args.batch
    else:
        batch = read_current_batch()
        if batch:
            output_dir = batch["BATCH_DIR"]
            matrix_ok(f"Batch actif détecté : {output_dir} ({batch.get('SECTEUR', '')})")
        else:
            output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)

    return villes, activite, nb_max, output_dir, output_name


def load_config(args) -> tuple[list[str], str, int, str]:
    """Retourne (villes, activite, nb_max, output_dir)."""
    if args.config:
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
        if "villes" in cfg:
            villes = cfg["villes"]
        elif "ville" in cfg:
            villes = [cfg["ville"]]
        else:
            matrix_fail("Clé 'ville' ou 'villes' manquante dans le JSON")
            sys.exit(1)
        activite = cfg["type_commerce"]
        nb_max = cfg.get("nb_max_resultats", 100)
        output_dir = args.output  # toujours CLI, jamais le JSON
    elif args.ville and args.activite:
        villes = [args.ville]
        activite = args.activite
        nb_max = args.nb_max
        output_dir = args.output
    else:
        matrix_fail("Spécifiez --ville et --activite, ou --config")
        sys.exit(1)

    # --batch > .current_batch > --output
    if args.batch:
        output_dir = args.batch
    else:
        batch = read_current_batch()
        if batch:
            output_dir = batch["BATCH_DIR"]
            matrix_ok(f"Batch actif détecté : {output_dir} ({batch.get('SECTEUR','')})")
    os.makedirs(output_dir, exist_ok=True)
    return villes, activite, nb_max, output_dir


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    if args.quiet:
        set_quiet(True)
    if not args.config and not args.ville:
        villes, activite, nb_max, output_dir, output_name = interactive_menu(args)
    else:
        villes, activite, nb_max, output_dir = load_config(args)
        output_name = None

    multi = len(villes) > 1

    # ── Bannière Matrix ──
    matrix_banner("SCRAPING PAGES JAUNES")

    matrix_kv("Villes", ", ".join(villes))
    matrix_kv("Activité", activite)
    matrix_kv("Max fiches/ville", str(nb_max))

    # ── Nom de fichier ──
    today_str = date.today().strftime("%Y%m%d")
    if output_name:
        filename = output_name
    else:
        if multi:
            default_name = f"pj_results_multi_{today_str}"
        else:
            ville_slug = villes[0].lower().replace(" ", "_").replace("-", "_")
            default_name = f"pj_results_{ville_slug}_{today_str}"
        filename = ask_filename(default_name)

    # ── Détection secteur pour filtre enseignes ──
    _batch_info = read_current_batch()
    _secteur = (_batch_info or {}).get("SECTEUR", "").upper()
    _apply_enseigne_filter = _secteur in ("MBT", "TG")
    if _apply_enseigne_filter:
        matrix_step(f"Filtre enseignes nationales actif (secteur {_secteur})")

    # ── Boucle sur les villes ──
    all_city_results: dict[str, list[dict]] = {}
    all_city_stats: dict[str, dict] = {}

    for ville in villes:
        matrix_section(f"Infiltration Pages Jaunes — {ville}")

        try:
            fiches, stats = scrape_pages_jaunes(
                activite, ville, nb_max, output_dir, proxy_url=args.proxy,
                debug=args.debug, deep_phone=getattr(args, "deep_phone", False),
            )

            # Appliquer les exclusions catégorielles
            fiches, excl_counts = apply_exclusions(fiches)
            total_excl = sum(excl_counts.values())
            if total_excl:
                matrix_step("Exclusions appliquées :")
                for cat_name, cnt in excl_counts.items():
                    print(f"      - {cat_name:<26s}: {cnt} fiche{'s' if cnt != 1 else ''}")
                print(f"      Total exclus : {total_excl} fiche{'s' if total_excl != 1 else ''}")

            # Appliquer le filtre enseignes nationales (MBT/TG uniquement)
            total_enseigne_excl = 0
            if _apply_enseigne_filter:
                fiches, enseigne_counts = apply_exclusions_enseignes(fiches)
                total_enseigne_excl = sum(enseigne_counts.values())
                if total_enseigne_excl:
                    matrix_step(f"Exclusions enseignes nationales : {total_enseigne_excl} fiches")
                    top = sorted(enseigne_counts.items(), key=lambda x: -x[1])[:10]
                    detail = ", ".join(f"{k.title()} ×{v}" for k, v in top)
                    if len(enseigne_counts) > 10:
                        detail += ", ..."
                    print(f"      {detail}")

            for f in fiches:
                f["Ville de recherche"] = ville

            all_city_results[ville] = fiches
            all_city_stats[ville] = stats
            all_city_stats[ville]["exclusions"] = total_excl
            all_city_stats[ville]["exclusions_enseignes"] = total_enseigne_excl
            matrix_ok(f"{len(fiches)} fiches pour {ville} ({total_excl + total_enseigne_excl} exclus)")

            # Cleanup backup per-city
            ville_slug = ville.lower().replace(" ", "_").replace("-", "_")
            backup_path = os.path.join(
                output_dir, f"pj_results_{ville_slug}_{today_str}_backup.csv",
            )
            if os.path.exists(backup_path):
                os.remove(backup_path)

        except Exception as e:
            matrix_fail(f"Erreur pour {ville} : {e}")
            all_city_results[ville] = []
            all_city_stats[ville] = {
                "pages_parcourues": 0, "fiches_brutes": 0,
                "doublons": 0, "blocages_cf": 0,
            }
            continue

    if not any(all_city_results.values()):
        matrix_warn("Aucune fiche collectée pour aucune ville.")
        sys.exit(0)

    # ── Export CSV ──
    csv_path = os.path.join(output_dir, f"{filename}.csv")

    if multi:
        # Consolidé : dédup globale sur (nom, code postal)
        all_fiches: list[dict] = []
        for fiches in all_city_results.values():
            all_fiches.extend(fiches)
        seen_keys: set[tuple[str, str]] = set()
        consolidated: list[dict] = []
        for f in all_fiches:
            key = (f.get("Nom de l'entreprise", "").lower(), f.get("Code Postal", ""))
            if key in seen_keys:
                continue
            seen_keys.add(key)
            consolidated.append(f)

        matrix_step("Export CSV consolidé...")
        export_csv(consolidated, csv_path)
        matrix_ok(f"Fichier : {csv_path}")

        # ── Synthèse ──
        matrix_section("RÉSULTATS — Données aspirées de la Matrice")
        matrix_kv("Recherche", f'"{activite}"')
        print()
        for ville, st in all_city_stats.items():
            n = len(all_city_results.get(ville, []))
            print(f"    {GREEN}▸{RESET} {ville}: {BOLD}{n}{RESET} fiches "
                  f"({st.get('pages_parcourues', 0)} pages, "
                  f"{st.get('blocages_cf', 0)} blocages CF)")
        matrix_separator()
        matrix_kv("Total consolidé (dédup nom+CP)", f"{len(consolidated)} fiches")
        # Agréger les stats téléphone
        agg_st = {"phones_direct": 0, "phones_click": 0, "phones_deep": 0, "phones_missing": 0}
        for st in all_city_stats.values():
            for k in agg_st:
                agg_st[k] += st.get(k, 0)
        _show_phone_stats(agg_st, len(consolidated))
        matrix_separator()
        matrix_kv("Fichier", os.path.basename(csv_path))
    else:
        ville = villes[0]
        fiches = all_city_results.get(ville, [])
        if not fiches:
            matrix_warn("Aucune fiche collectée.")
            sys.exit(0)

        matrix_step("Export CSV...")
        export_csv(fiches, csv_path)
        matrix_ok(f"Fichier : {csv_path}")

        st = all_city_stats[ville]
        matrix_section(f"RÉSULTATS — Pages Jaunes — {ville}")
        matrix_kv("Recherche", f'"{activite}"')
        matrix_kv("Fiches collectées", str(st["fiches_brutes"]))
        matrix_kv("Doublons supprimés", str(st["doublons"]))
        matrix_kv("Fiches retenues", str(len(fiches)))
        matrix_kv("Pages parcourues", str(st["pages_parcourues"]))
        matrix_kv("Blocages CF détectés", str(st["blocages_cf"]))
        matrix_separator()
        _show_phone_stats(st, len(fiches))
        matrix_separator()
        matrix_kv("Fichier", os.path.basename(csv_path))

    # ── Clôture Matrix ──
    morpheus_says()

    # ── Enchaînement pipeline ──
    batch = read_current_batch()
    if batch and not args.quiet:
        print()
        matrix_section("Enchaînement pipeline")
        print(f"    {GREEN}1.{RESET} Enrichissement + Croisement (SIRENE — gratuit)")
        print(f"    {GREEN}2.{RESET} Enrichissement + Croisement (PAPPERS — nécessite clé API)")
        print(f"    {GREEN}3.{RESET} Enrichissement uniquement (SIRENE)")
        print(f"    {GREEN}4.{RESET} Enrichissement uniquement (PAPPERS)")
        print(f"    {GREEN}5.{RESET} Terminer")
        print()
        choice = input(f"    {GREEN}▸{RESET} Choix [défaut 1] : ").strip() or "1"
        script_dir = os.path.dirname(os.path.abspath(__file__))

        if choice in ("2", "4"):
            pappers_key = os.environ.get("PAPPERS_API_KEY", "").strip()
            if not pappers_key:
                matrix_fail("Variable PAPPERS_API_KEY non définie.")
                matrix_step("Lance : export PAPPERS_API_KEY=\"ta_clé\" (Linux/Mac)")
                matrix_step("  ou : $env:PAPPERS_API_KEY = \"ta_clé\" (PowerShell)")
                print()
                matrix_step("Retour au menu...")
                alt = input(f"    {GREEN}▸{RESET} Utiliser SIRENE à la place ? (O/n) : ").strip().lower()
                if alt == "n":
                    pass  # quitter
                else:
                    choice = "1" if choice == "2" else "3"

        if choice in ("1", "3"):
            matrix_rain_fullscreen(duration=4, next_title="PHASE 2 — ENRICHISSEMENT SIRET (SIRENE)")
            cmd = [sys.executable, os.path.join(script_dir, "script_enrichissement.py"),
                   "--input", csv_path]
            if choice == "1":
                cmd.append("--chain")
            subprocess.run(cmd)
        elif choice in ("2", "4"):
            matrix_rain_fullscreen(duration=4, next_title="PHASE 2 — ENRICHISSEMENT SIRET (PAPPERS)")
            cmd = [sys.executable, os.path.join(script_dir, "script_enrichissement_pappers.py"),
                   "--input", csv_path]
            subprocess.run(cmd)
            if choice == "2":
                matrix_rain_fullscreen(duration=4, next_title="PHASE 3 — CROISEMENT FINAL")
                cmd = [sys.executable, os.path.join(script_dir, "script_comparaison.py")]
                subprocess.run(cmd)


if __name__ == "__main__":
    main()
