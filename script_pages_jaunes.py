#!/usr/bin/env python3
"""
script_pages_jaunes.py — Scraping Pages Jaunes
===============================================
Collecte des entreprises sur Pages Jaunes par activité et ville,
via Playwright (headless Chromium) pour contourner la protection Cloudflare.

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
from urllib.parse import quote

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


# ── Helpers ───────────────────────────────────────────────────────────────────

def random_delay(min_s=1.0, max_s=3.0):
    time.sleep(random.uniform(min_s, max_s))


def extract_cp_ville(adresse: str) -> tuple[str, str]:
    """Extrait code postal et ville depuis une adresse texte."""
    if not adresse:
        return "", ""
    m = re.search(r"(\d{5})\s+([A-Za-zÀ-ÿ\s\-']+)", adresse)
    if m:
        return m.group(1), m.group(2).strip().title()
    return "", ""


def find_chromium() -> str | None:
    """Cherche un binaire Chromium installé par Playwright."""
    # Headless shell (newer installs)
    hs_paths = sorted(_glob.glob(
        "/root/.cache/ms-playwright/chromium_headless_shell-*/chrome-headless-shell-linux64/chrome-headless-shell"
    ))
    if hs_paths:
        return hs_paths[-1]
    # Full Chrome (older installs)
    chrome_paths = sorted(_glob.glob(
        "/root/.cache/ms-playwright/chromium-*/chrome-linux/chrome"
    ))
    if chrome_paths:
        return chrome_paths[-1]
    return None


def detect_proxy() -> dict | None:
    """Détecte proxy depuis les variables d'environnement."""
    proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
    if proxy_url:
        return {"server": proxy_url}
    return None


def build_search_url(activite: str, ville: str, page: int) -> str:
    return (
        f"{PJ_BASE_URL}/annuaire/chercherlespros"
        f"?quoiqui={quote(activite)}&ou={quote(ville)}&page={page}"
    )


# ── Parsing d'un bloc entreprise ─────────────────────────────────────────────

def parse_bloc(bloc) -> dict | None:
    """Extrait les données d'un bloc entreprise PJ."""
    today = date.today().isoformat()

    # Nom
    name_el = bloc.query_selector(
        "a.bi-denomination, span.bi-denomination, "
        "[class*='denomination'], h3 a, h2 a"
    )
    nom = name_el.inner_text().strip() if name_el else ""
    if not nom:
        return None

    # URL fiche
    fiche_url = ""
    link_el = bloc.query_selector("a.bi-denomination, a[href*='/pros/']")
    if link_el:
        href = link_el.get_attribute("href") or ""
        if href.startswith("/"):
            fiche_url = PJ_BASE_URL + href
        elif href.startswith("http"):
            fiche_url = href

    # Adresse
    addr_el = bloc.query_selector(
        "[class*='adresse'], address, [class*='address'], .bi-adresse"
    )
    adresse = addr_el.inner_text().strip() if addr_el else ""
    cp, ville = extract_cp_ville(adresse)

    # Téléphone
    tel = ""
    tel_el = bloc.query_selector(
        "[class*='tel'], [class*='phone'], a[href^='tel:']"
    )
    if tel_el:
        href = tel_el.get_attribute("href") or ""
        if href.startswith("tel:"):
            tel = href.replace("tel:", "").strip()
        else:
            tel = tel_el.inner_text().strip()
    # Fallback: numéro affiché (fantomas pattern from Rogue-two)
    if not tel:
        tel_el2 = bloc.query_selector("[id*='fantomas'], .number-contact")
        if tel_el2:
            tel = tel_el2.inner_text().strip()

    # Nettoyage téléphone
    tel = re.sub(r"[^\d+\s]", "", tel).strip()

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
                        output_dir: str) -> tuple[list[dict], dict]:
    """
    Scrape Pages Jaunes via Playwright.
    Retourne (fiches, stats).
    """
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

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
        ]

        launch_kwargs: dict = {
            "headless": True,
            "executable_path": exec_path,
            "args": launch_args,
        }

        proxy = detect_proxy()
        if proxy:
            launch_kwargs["proxy"] = proxy
            print(f"  🔌 Proxy détecté : {proxy['server'][:50]}…")

        browser = p.chromium.launch(**launch_kwargs)
        ua = random.choice(USER_AGENTS)
        context = browser.new_context(
            user_agent=ua,
            locale="fr-FR",
        )
        page = context.new_page()
        page.set_default_timeout(20_000)

        # Anti-detection
        page.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )

        # Warmup homepage (CF cookies)
        print("  📡 Warmup sur pagesjaunes.fr…", end=" ", flush=True)
        try:
            page.goto(PJ_BASE_URL, wait_until="domcontentloaded", timeout=30_000)
            random_delay(2, 4)
            print("✅")
        except Exception as e:
            print(f"⚠️  {e}")
            # Continue anyway — CF cookies might not be needed

        # Pagination loop
        page_num = 1
        while len(fiches) < nb_max:
            url = build_search_url(activite, ville, page_num)
            print(f"  📡 Page {page_num}…", end=" ", flush=True)

            try:
                # Rotate UA every few pages
                if page_num > 1 and page_num % 3 == 0:
                    context.close()
                    ua = random.choice(USER_AGENTS)
                    context = browser.new_context(user_agent=ua, locale="fr-FR")
                    page = context.new_page()
                    page.set_default_timeout(20_000)
                    page.add_init_script(
                        "Object.defineProperty(navigator, 'webdriver', "
                        "{get: () => undefined})"
                    )

                random_delay(1, 3)
                page.goto(url, wait_until="domcontentloaded", timeout=20_000)

                # Wait for company blocks
                try:
                    page.wait_for_selector(
                        "article.bi-bloc, div[class*='bi-bloc'], "
                        "li[class*='bi-bloc'], [data-bi-name]",
                        timeout=8_000,
                    )
                except PWTimeout:
                    pass  # May be empty page or CF block

            except Exception as e:
                print(f"❌ {e}")
                consecutive_empty += 1
                if consecutive_empty >= MAX_EMPTY_PAGES:
                    print(f"  🛑 {MAX_EMPTY_PAGES} pages vides consécutives — arrêt (blocage CF probable)")
                    stats["blocages_cf"] = consecutive_empty
                    break
                page_num += 1
                continue

            # Extract company blocks
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

                # Déduplication
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

            # Sauvegarde intermédiaire
            if len(fiches) % SAVE_EVERY < page_count:
                save_intermediate_csv(fiches, backup_path)
                print(f"  💾 Sauvegarde intermédiaire ({len(fiches)} fiches)")

            # Check for next page
            has_next = page.query_selector(
                "a[rel='next'], a.pagination-next, li.next a"
            )
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
    return parser.parse_args()


def load_config(args) -> tuple[str, str, int, str]:
    """Retourne (ville, activite, nb_max, output_dir)."""
    if args.config:
        with open(args.config, encoding="utf-8") as f:
            cfg = json.load(f)
        ville = cfg["ville"]
        activite = cfg["type_commerce"]
        nb_max = cfg.get("nb_max_resultats", 100)
        output_dir = cfg.get("output", ".")
    elif args.ville and args.activite:
        ville = args.ville
        activite = args.activite
        nb_max = args.nb_max
        output_dir = args.output
    else:
        print("❌ Spécifiez --ville et --activite, ou --config")
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)
    return ville, activite, nb_max, output_dir


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    ville, activite, nb_max, output_dir = load_config(args)

    print("══════════════════════════════════════════════════")
    print("  SCRAPING PAGES JAUNES")
    print(f"  Ville    : {ville}")
    print(f"  Activité : {activite}")
    print(f"  Max      : {nb_max} fiches")
    print("══════════════════════════════════════════════════")

    # Scraping
    fiches, stats = scrape_pages_jaunes(activite, ville, nb_max, output_dir)

    if not fiches:
        print("\n⚠️  Aucune fiche collectée.")
        print("   Causes possibles : blocage Cloudflare, proxy, ou aucun résultat PJ.")
        sys.exit(0)

    # Export
    print("\n── Export ──────────────────────────────────────")
    xlsx_path, csv_path = export_results(fiches, output_dir, ville)
    print(f"  ✅ {xlsx_path}")
    print(f"  ✅ {csv_path}")

    # Cleanup backup
    today_str = date.today().strftime("%Y%m%d")
    ville_slug = ville.lower().replace(" ", "_").replace("-", "_")
    backup_path = os.path.join(
        output_dir, f"pj_results_{ville_slug}_{today_str}_backup.csv"
    )
    if os.path.exists(backup_path):
        os.remove(backup_path)

    # Synthèse
    print_synthese(ville, activite, fiches, stats, xlsx_path, csv_path)


if __name__ == "__main__":
    main()
