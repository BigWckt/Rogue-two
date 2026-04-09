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
import glob as _glob
import json
import os
import random
import re
import sys
import time
from datetime import date
from urllib.parse import quote, urlparse

import pandas as pd

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

EXCEL_COLUMNS = [
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


def find_chromium() -> str | None:
    """Cherche le vrai Chromium (pas headless shell) installé par Playwright."""
    # Préférer le vrai Chrome (pas le headless shell) — requis pour le bypass CF
    chrome_paths = sorted(_glob.glob(
        "/root/.cache/ms-playwright/chromium-*/chrome-linux/chrome"
    ))
    if chrome_paths:
        return chrome_paths[-1]
    # Fallback headless shell (moins fiable pour CF)
    hs_paths = sorted(_glob.glob(
        "/root/.cache/ms-playwright/chromium_headless_shell-*/chrome-headless-shell-linux64/chrome-headless-shell"
    ))
    if hs_paths:
        print("  ⚠️  Seul le headless shell est dispo — le bypass CF peut échouer")
        return hs_paths[-1]
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


# ── Parsing d'un bloc entreprise ─────────────────────────────────────────────

def parse_bloc(bloc) -> dict | None:
    """
    Extrait les données d'un bloc entreprise PJ.
    Sélecteurs CSS issus de README_scraper_pj.md (structure DOM PJ vérifiée).
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

    # Téléphone : li .bi-fantomas .number-contact
    # Note : .bi-fantomas est un div FRÈRE de .bi-content, pas imbriqué dedans
    tel = ""
    tel_el = bloc.query_selector(".bi-fantomas .number-contact")
    if tel_el:
        raw_tel = tel_el.inner_text().strip()
        # Nettoyer le préfixe "Tél : "
        raw_tel = re.sub(r"^T[ée]l\s*:\s*", "", raw_tel)
        tel = validate_phone(raw_tel)

    # Site web
    site = ""
    web_el = bloc.query_selector(
        "a[class*='site'], a[class*='web'], "
        "a[href*='http'][class*='url'], a.bi-website"
    )
    if web_el:
        site = web_el.get_attribute("href") or ""

    # Catégorie
    cat_el = bloc.query_selector(
        ".bi-activite, [class*='activite'], [class*='rubrique']"
    )
    categorie = cat_el.inner_text().strip() if cat_el else ""

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
    }


# ── Sauvegarde intermédiaire ─────────────────────────────────────────────────

def save_intermediate_csv(fiches: list[dict], filepath: str):
    """Sauvegarde CSV intermédiaire pour reprise."""
    try:
        with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=EXCEL_COLUMNS)
            writer.writeheader()
            writer.writerows(fiches)
    except Exception as e:
        print(f"  ⚠️  Erreur sauvegarde intermédiaire : {e}")


# ── Scraping principal ───────────────────────────────────────────────────────

def scrape_pages_jaunes(activite: str, ville: str, nb_max: int,
                        output_dir: str, proxy_url: str | None = None) -> tuple[list[dict], dict]:
    """
    Scrape Pages Jaunes via Playwright avec bypass Cloudflare.
    Retourne (fiches, stats).
    proxy_url : URL proxy explicite (ex: http://user:pass@host:port), None = sortie directe.
    """
    from playwright.sync_api import sync_playwright

    stats = {
        "pages_parcourues": 0,
        "fiches_brutes": 0,
        "doublons": 0,
        "blocages_cf": 0,
    }

    exec_path = find_chromium()
    if not exec_path:
        print("❌ Aucun binaire Chromium trouvé.")
        print("   Installer avec : playwright install chromium")
        sys.exit(1)

    print(f"  🌐 Chromium : {exec_path}")

    # Backup CSV path
    today_str = date.today().strftime("%Y%m%d")
    ville_slug = ville.lower().replace(" ", "_").replace("-", "_")
    backup_path = os.path.join(
        output_dir, f"pj_results_{ville_slug}_{today_str}_backup.csv"
    )

    fiches: list[dict] = []
    seen_keys: set[tuple[str, str]] = set()
    consecutive_empty = 0

    with sync_playwright() as p:
        launch_args = [
            "--no-sandbox",
            "--disable-blink-features=AutomationControlled",
            "--disable-dev-shm-usage",
            "--ignore-certificate-errors",
        ]

        # Par défaut : sortie directe, ignorer les proxy système
        # (HTTPS_PROXY / HTTP_PROXY causent ERR_INVALID_AUTH_CREDENTIALS)
        if not proxy_url:
            launch_args.append("--no-proxy-server")

        launch_kwargs: dict = {
            "headless": True,
            "executable_path": exec_path,
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
            print(f"  🔌 Proxy explicite : {proxy_cfg['server']}")
        else:
            print("  🔌 Sortie directe (pas de proxy)")

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

        # ── Warmup CF : navigation bidon pour établir le cookie CF ────────
        print("  📡 Warmup CF (fleuriste Paris 20)…", end=" ", flush=True)
        try:
            warmup_url = (
                f"{PJ_BASE_URL}/annuaire/chercherlespros"
                "?quoiqui=fleuriste&ou=Paris+20"
            )
            page.goto(warmup_url, wait_until="domcontentloaded", timeout=30_000)
            time.sleep(DELAY_AFTER_NAV)
            if wait_for_pj_content(page, timeout_s=CF_TIMEOUT):
                print("✅ Cookie CF établi")
            else:
                print("⚠️  Timeout CF — le scraping risque d'échouer")
                stats["blocages_cf"] += 1
        except Exception as e:
            print(f"⚠️  {e}")

        # ── Boucle de pagination ──────────────────────────────────────────
        page_num = 1
        while len(fiches) < nb_max:
            url = build_search_url(activite, ville, page_num)
            print(f"  📡 Page {page_num}…", end=" ", flush=True)

            try:
                # Délai entre pages (5s même mot-clé)
                if page_num > 1:
                    time.sleep(DELAY_SAME_KEYWORD)

                page.goto(url, wait_until="domcontentloaded", timeout=30_000)
                time.sleep(DELAY_AFTER_NAV)

                # Attente adaptative : vérifier le titre toutes les 5s
                if not wait_for_pj_content(page, timeout_s=CF_TIMEOUT):
                    print(f"⚠️  Timeout CF ({CF_TIMEOUT}s)")
                    consecutive_empty += 1
                    stats["blocages_cf"] += 1
                    if consecutive_empty >= MAX_EMPTY_PAGES:
                        print(f"  🛑 {MAX_EMPTY_PAGES} pages bloquées consécutives — arrêt")
                        break
                    page_num += 1
                    continue

            except Exception as e:
                print(f"❌ {e}")
                consecutive_empty += 1
                if consecutive_empty >= MAX_EMPTY_PAGES:
                    print(f"  🛑 {MAX_EMPTY_PAGES} pages échouées consécutives — arrêt")
                    stats["blocages_cf"] = consecutive_empty
                    break
                page_num += 1
                continue

            # Extraire les blocs entreprise (<li> contenant .bi-content)
            blocs = page.query_selector_all("li:has(.bi-content)")
            if not blocs:
                # Fallbacks
                blocs = page.query_selector_all(
                    "article.bi-bloc, div[class*='bi-bloc'], li[class*='bi-bloc']"
                )
            if not blocs:
                blocs = page.query_selector_all("[data-bi-name]")

            if not blocs:
                consecutive_empty += 1
                print(f"0 fiches (vide {consecutive_empty}/{MAX_EMPTY_PAGES})")
                if consecutive_empty >= MAX_EMPTY_PAGES:
                    print(f"  🛑 {MAX_EMPTY_PAGES} pages vides consécutives — arrêt (blocage CF probable)")
                    stats["blocages_cf"] = consecutive_empty
                    break
                page_num += 1
                continue

            consecutive_empty = 0
            page_count = 0

            for bloc in blocs:
                if len(fiches) >= nb_max:
                    break
                fiche = parse_bloc(bloc)
                if fiche is None:
                    continue

                stats["fiches_brutes"] += 1

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

            stats["pages_parcourues"] = page_num
            print(f"✅ {page_count} fiches ({len(fiches)} total)")

            # Sauvegarde intermédiaire toutes les SAVE_EVERY fiches
            if len(fiches) % SAVE_EVERY < page_count:
                save_intermediate_csv(fiches, backup_path)
                print(f"  💾 Sauvegarde intermédiaire ({len(fiches)} fiches)")

            # Pagination : a#pagination-next
            has_next = page.query_selector("a#pagination-next")
            if not has_next:
                print("  ℹ️  Dernière page atteinte")
                break

            page_num += 1

        browser.close()

    return fiches, stats


# ── Export ────────────────────────────────────────────────────────────────────

def export_results(fiches: list[dict], output_dir: str,
                   ville: str) -> tuple[str, str]:
    today_str = date.today().strftime("%Y%m%d")
    ville_slug = ville.lower().replace(" ", "_").replace("-", "_")
    base_name = f"pj_results_{ville_slug}_{today_str}"
    xlsx_path = os.path.join(output_dir, f"{base_name}.xlsx")
    csv_path = os.path.join(output_dir, f"{base_name}.csv")

    df = pd.DataFrame(fiches, columns=EXCEL_COLUMNS)
    df.to_excel(xlsx_path, index=False, engine="openpyxl")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    return xlsx_path, csv_path


# ── Synthèse console ─────────────────────────────────────────────────────────

def print_synthese(ville: str, activite: str, fiches: list[dict],
                   stats: dict, xlsx_path: str, csv_path: str):
    n_total = len(fiches)
    print()
    print("══════════════════════════════════════════════════")
    print(f"  RÉSULTATS — Pages Jaunes — {ville}")
    print(f"  Recherche : \"{activite}\"")
    print()
    print(f"  Fiches collectées      : {stats['fiches_brutes']}")
    print(f"  Doublons supprimés     : {stats['doublons']}")
    print(f"  Fiches retenues        : {n_total}")
    print(f"  Pages parcourues       : {stats['pages_parcourues']}")
    print(f"  Blocages CF détectés   : {stats['blocages_cf']}")
    print()
    print(f"  📁 Fichier : {os.path.basename(xlsx_path)}")
    print(f"  📁 Backup  : {os.path.basename(csv_path)}")
    print("══════════════════════════════════════════════════")


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
                        help="Répertoire de sortie")
    parser.add_argument("--config", type=str, help="Fichier JSON de paramètres")
    parser.add_argument("--proxy", type=str, default=None,
                        help="Proxy explicite (ex: http://user:pass@host:port)")
    return parser.parse_args()


def load_config(args) -> tuple[list[str], str, int, str]:
    """Retourne (villes, activite, nb_max, output_dir)."""
    if args.config:
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
        # Support "ville" (string) ou "villes" (array)
        if "villes" in cfg:
            villes = cfg["villes"]
        elif "ville" in cfg:
            villes = [cfg["ville"]]
        else:
            print("❌ Clé 'ville' ou 'villes' manquante dans le JSON")
            sys.exit(1)
        activite = cfg["type_commerce"]
        nb_max = cfg.get("nb_max_resultats", 100)
        output_dir = cfg.get("output", ".")
    elif args.ville and args.activite:
        villes = [args.ville]
        activite = args.activite
        nb_max = args.nb_max
        output_dir = args.output
    else:
        print("❌ Spécifiez --ville et --activite, ou --config")
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)
    return villes, activite, nb_max, output_dir


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    villes, activite, nb_max, output_dir = load_config(args)
    multi = len(villes) > 1

    print("══════════════════════════════════════════════════")
    print("  SCRAPING PAGES JAUNES")
    print(f"  Villes   : {', '.join(villes)}")
    print(f"  Activité : {activite}")
    print(f"  Max      : {nb_max} fiches par ville")
    print("══════════════════════════════════════════════════")

    all_city_results: dict[str, list[dict]] = {}
    all_city_stats: dict[str, dict] = {}
    today_str = date.today().strftime("%Y%m%d")

    for ville in villes:
        print(f"\n{'═' * 50}")
        print(f"  📍 {ville}")
        print(f"{'═' * 50}")

        try:
            fiches, stats = scrape_pages_jaunes(
                activite, ville, nb_max, output_dir, proxy_url=args.proxy,
            )

            # Ajouter "Ville de recherche"
            for f in fiches:
                f["Ville de recherche"] = ville

            all_city_results[ville] = fiches
            all_city_stats[ville] = stats
            print(f"  ✅ {len(fiches)} fiches pour {ville}")

            # Cleanup backup per-city
            ville_slug = ville.lower().replace(" ", "_").replace("-", "_")
            backup_path = os.path.join(
                output_dir, f"pj_results_{ville_slug}_{today_str}_backup.csv",
            )
            if os.path.exists(backup_path):
                os.remove(backup_path)

        except Exception as e:
            print(f"  ❌ Erreur pour {ville} : {e}")
            all_city_results[ville] = []
            all_city_stats[ville] = {
                "pages_parcourues": 0, "fiches_brutes": 0,
                "doublons": 0, "blocages_cf": 0,
            }
            continue

    if not any(all_city_results.values()):
        print("\n⚠️  Aucune fiche collectée pour aucune ville.")
        sys.exit(0)

    if multi:
        xlsx_path = os.path.join(output_dir, f"pj_results_multi_{today_str}.xlsx")
        csv_path = xlsx_path.replace(".xlsx", ".csv")

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

        columns = EXCEL_COLUMNS
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            for ville, fiches in all_city_results.items():
                if fiches:
                    df = pd.DataFrame(fiches, columns=columns)
                    df.to_excel(writer, sheet_name=ville[:31], index=False)
            df_all = pd.DataFrame(consolidated, columns=columns)
            df_all.to_excel(writer, sheet_name="Consolidé", index=False)

        df_all.to_csv(csv_path, index=False, encoding="utf-8-sig")

        print(f"\n── Export multi-villes ──────────────────────────")
        print(f"  ✅ {xlsx_path}")
        print(f"  ✅ {csv_path}")
        print()
        print("══════════════════════════════════════════════════")
        print("  RÉSULTATS MULTI-VILLES — Pages Jaunes")
        print(f"  Recherche : \"{activite}\"")
        print()
        for ville, st in all_city_stats.items():
            n = len(all_city_results.get(ville, []))
            print(f"  📍 {ville}: {n} fiches "
                  f"({st.get('pages_parcourues', 0)} pages, "
                  f"{st.get('blocages_cf', 0)} blocages CF)")
        print("  ──────────────────────────────────────────────")
        print(f"  Total consolidé (dédup nom+CP) : {len(consolidated)} fiches")
        print(f"  📁 Fichier : {os.path.basename(xlsx_path)}")
        print("══════════════════════════════════════════════════")
    else:
        ville = villes[0]
        fiches = all_city_results.get(ville, [])
        if not fiches:
            print("\n⚠️  Aucune fiche collectée.")
            sys.exit(0)
        xlsx_path, csv_path = export_results(fiches, output_dir, ville)
        print(f"\n── Export ──────────────────────────────────────")
        print(f"  ✅ {xlsx_path}")
        print(f"  ✅ {csv_path}")
        print_synthese(ville, activite, fiches, all_city_stats[ville],
                       xlsx_path, csv_path)


if __name__ == "__main__":
    main()
