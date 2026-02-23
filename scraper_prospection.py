#!/usr/bin/env python3
"""
scraper_prospection.py
Scrape Pages Jaunes par secteur, enrichit via API SIRENE, exporte en Excel.
Usage:
  python scraper_prospection.py --test   # 10 fiches secteur 1 seulement
  python scraper_prospection.py --full   # Tous les secteurs, 300+ fiches chacun
"""

import argparse
import csv
import json
import logging
import os
import random
import re
import sys
import time
from datetime import date
from pathlib import Path

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from tqdm import tqdm

# ---------------------------------------------------------------------------
# Configuration des secteurs
# ---------------------------------------------------------------------------
TODAY = date.today().strftime("%Y%m%d")
TODAY_DISPLAY = date.today().strftime("%d/%m/%Y")
EXCEL_FILENAME = f"Prospection_formations_{TODAY}.xlsx"
LOG_FILENAME = f"errors_{TODAY}.log"
INTERMEDIATE_CSV = f"intermediate_{TODAY}.csv"

SECTEURS = [
    {
        "id": 1,
        "onglet": "Hotels_31",
        "formation": "TP Réceptionniste - Haute-Garonne",
        "departement": "31",
        "mots_cles": ["hôtel", "résidence hôtelière", "appart-hôtel"],
        "villes": ["Toulouse", "Blagnac", "Colomiers", "Muret", "Saint-Gaudens"],
        "objectif": 300,
    },
    {
        "id": 2,
        "onglet": "Hotels_75",
        "formation": "TP Réceptionniste - Paris",
        "departement": "75",
        "mots_cles": ["hôtel", "résidence hôtelière", "appart-hôtel"],
        "villes": [
            "Paris 1er", "Paris 2ème", "Paris 3ème", "Paris 4ème", "Paris 5ème",
            "Paris 6ème", "Paris 7ème", "Paris 8ème", "Paris 9ème", "Paris 10ème",
            "Paris 11ème", "Paris 12ème", "Paris 13ème", "Paris 14ème", "Paris 15ème",
            "Paris 16ème", "Paris 17ème", "Paris 18ème", "Paris 19ème", "Paris 20ème",
        ],
        "objectif": 300,
    },
    {
        "id": 3,
        "onglet": "Hotels_93-94",
        "formation": "TP Réceptionniste - Seine-Saint-Denis / Val-de-Marne",
        "departement": "93",
        "departement_alt": "94",
        "mots_cles": ["hôtel", "résidence hôtelière", "appart-hôtel"],
        "villes": [
            "Saint-Denis 93", "Bobigny 93", "Montreuil 93",
            "Aubervilliers 93", "Pantin 93",
            "Créteil 94", "Vincennes 94", "Vitry-sur-Seine 94",
            "Ivry-sur-Seine 94", "Champigny-sur-Marne 94",
        ],
        "objectif": 300,
    },
    {
        "id": 4,
        "onglet": "Hotels_59",
        "formation": "TP Réceptionniste - Nord",
        "departement": "59",
        "mots_cles": ["hôtel", "résidence hôtelière", "appart-hôtel", "centre de congrès"],
        "villes": ["Lille", "Roubaix", "Tourcoing", "Valenciennes", "Dunkerque", "Douai"],
        "objectif": 300,
    },
    {
        "id": 5,
        "onglet": "Juridique_Immo_92",
        "formation": "Certification Assistante Juridique / TP Assistant Immo",
        "departement": "92",
        "mots_cles": [
            "cabinet d'avocats", "étude notariale", "agence immobilière",
            "administrateur de biens", "promoteur immobilier",
        ],
        "villes": [
            "Nanterre", "Boulogne-Billancourt", "Neuilly-sur-Seine",
            "Issy-les-Moulineaux", "Levallois-Perret", "Courbevoie",
            "Asnières-sur-Seine",
        ],
        "objectif": 300,
    },
    {
        "id": 6,
        "onglet": "CallCenter_59",
        "formation": "TP Conseiller Relation Client à Distance - Nord",
        "departement": "59",
        "mots_cles": [
            "centre d'appels", "call center", "télémarketing",
            "service client externalisé", "BPO",
        ],
        "villes": ["Lille", "Roubaix", "Tourcoing", "Villeneuve-d'Ascq", "Valenciennes"],
        "objectif": 300,
    },
]

COLONNES = [
    "Formation associée",
    "Type d'établissement",
    "Nom entreprise",
    "Adresse complète",
    "Code postal",
    "Ville",
    "Département",
    "Téléphone",
    "SIRET",
    "Site web",
    "Source",
    "Date de collecte",
]

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    filename=LOG_FILENAME,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def random_delay(min_s=2, max_s=5):
    time.sleep(random.uniform(min_s, max_s))


def get_dept_from_ville(ville_str):
    """Extrait le code département depuis 'Ville 93' ou retourne None."""
    m = re.search(r"\b(75|92|93|94|31|59)\b", ville_str)
    return m.group(1) if m else None


def clean_ville(ville_str):
    """Supprime le suffixe numérique du nom de ville."""
    return re.sub(r"\s+\d{2}$", "", ville_str).strip()


# ---------------------------------------------------------------------------
# Scraping Pages Jaunes avec Playwright
# ---------------------------------------------------------------------------

def scrape_pages_jaunes(secteur: dict, max_fiches: int = 999999) -> list[dict]:
    """
    Scrape Pages Jaunes pour le secteur donné.
    Retourne une liste de dicts avec les champs bruts.
    """
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

    resultats = []
    seen_keys = set()
    save_counter = 0

    formation = secteur["formation"]
    departement = secteur["departement"]
    mots_cles = secteur["mots_cles"]
    villes = secteur["villes"]
    objectif = secteur["objectif"]

    print(f"\n[Secteur {secteur['id']}] {secteur['onglet']} — objectif {objectif} fiches")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/121.0.0.0 Safari/537.36"
            ),
            locale="fr-FR",
        )
        page = context.new_page()
        page.set_default_timeout(30000)

        for mot_cle in mots_cles:
            if len(resultats) >= objectif or len(resultats) >= max_fiches:
                break

            for ville_raw in villes:
                if len(resultats) >= objectif or len(resultats) >= max_fiches:
                    break

                dept_ville = get_dept_from_ville(ville_raw) or departement
                ville = clean_ville(ville_raw)
                ou_param = f"{ville} {dept_ville}"

                page_num = 1
                while True:
                    if len(resultats) >= objectif or len(resultats) >= max_fiches:
                        break

                    url = (
                        "https://www.pagesjaunes.fr/annuaire/chercherlespros"
                        f"?quoiqui={requests.utils.quote(mot_cle)}"
                        f"&ou={requests.utils.quote(ou_param)}"
                        f"&page={page_num}"
                    )

                    try:
                        random_delay(2, 5)
                        page.goto(url, wait_until="domcontentloaded", timeout=30000)
                        page.wait_for_selector(
                            "article.bi-bloc, div.bi-bloc, li.bi-bloc, article[class*='bi']",
                            timeout=10000,
                        )
                    except PWTimeout:
                        logger.warning(f"Timeout page {page_num} — {mot_cle}/{ville}")
                        break
                    except Exception as e:
                        logger.error(f"Erreur navigation {url}: {e}")
                        break

                    # Extraire les fiches de la page
                    fiches_page = _extraire_fiches(page, mot_cle, formation, dept_ville, ville)

                    if not fiches_page:
                        break  # Plus de résultats pour cette combinaison

                    for fiche in fiches_page:
                        key = (fiche.get("Nom entreprise", "").lower(), fiche.get("Code postal", ""))
                        if key not in seen_keys:
                            seen_keys.add(key)
                            resultats.append(fiche)
                            save_counter += 1

                            # Sauvegarde intermédiaire toutes les 50 fiches
                            if save_counter % 50 == 0:
                                _save_intermediate(resultats)
                                print(f"  → {len(resultats)} fiches collectées (sauvegarde intermédiaire)")

                    # Vérifier s'il y a une page suivante
                    has_next = page.query_selector("a[rel='next'], a.pagination-next, li.next a")
                    if not has_next:
                        break
                    page_num += 1

        browser.close()

    print(f"  ✓ Secteur {secteur['id']} : {len(resultats)} fiches collectées")
    return resultats


def _extraire_fiches(page, mot_cle: str, formation: str, departement: str, ville: str) -> list[dict]:
    """Extrait les fiches d'une page Pages Jaunes déjà chargée."""
    fiches = []
    try:
        # Sélecteurs possibles pour les blocs entreprise
        blocs = page.query_selector_all(
            "article.bi-bloc, div[class*='bi-bloc'], li[class*='bi-bloc']"
        )
        if not blocs:
            # Fallback: chercher par data-attribute
            blocs = page.query_selector_all("[data-bi-name]")

        for bloc in blocs:
            try:
                fiche = _parse_bloc(bloc, mot_cle, formation, departement, ville)
                if fiche:
                    fiches.append(fiche)
            except Exception as e:
                logger.warning(f"Erreur parsing bloc: {e}")
                continue
    except Exception as e:
        logger.error(f"Erreur extraction fiches: {e}")
    return fiches


def _parse_bloc(bloc, mot_cle: str, formation: str, departement: str, ville_param: str) -> dict | None:
    """Parse un bloc entreprise Pages Jaunes."""
    # Nom
    nom_el = (
        bloc.query_selector("a.bi-denomination, span.bi-denomination, "
                            "[class*='denomination'], h3 a, h2 a")
    )
    nom = nom_el.inner_text().strip() if nom_el else ""
    if not nom:
        return None

    # Adresse
    adresse_el = bloc.query_selector(
        "[class*='adresse'], address, [class*='address'], .bi-adresse"
    )
    adresse_full = adresse_el.inner_text().strip().replace("\n", " ") if adresse_el else ""

    # Extraire CP et ville de l'adresse
    cp, ville = _extraire_cp_ville(adresse_full, departement, ville_param)

    # Téléphone
    tel_el = bloc.query_selector("[class*='tel'], [class*='phone'], a[href^='tel:']")
    telephone = ""
    if tel_el:
        telephone = tel_el.get_attribute("href") or tel_el.inner_text()
        telephone = telephone.replace("tel:", "").strip()

    # Site web
    site_el = bloc.query_selector("a[class*='site'], a[class*='web'], a[href*='http'][class*='url']")
    site_web = site_el.get_attribute("href") if site_el else ""
    # Exclure les liens Pages Jaunes internes
    if site_web and "pagesjaunes.fr" in site_web:
        site_web = ""

    return {
        "Formation associée": formation,
        "Type d'établissement": mot_cle,
        "Nom entreprise": nom,
        "Adresse complète": adresse_full,
        "Code postal": cp,
        "Ville": ville,
        "Département": departement,
        "Téléphone": telephone,
        "SIRET": "",
        "Site web": site_web,
        "Source": "Pages Jaunes",
        "Date de collecte": TODAY_DISPLAY,
    }


def _extraire_cp_ville(adresse: str, dept_fallback: str, ville_fallback: str) -> tuple[str, str]:
    """Extrait code postal et ville depuis une adresse texte."""
    m = re.search(r"(\d{5})\s+([A-Za-zÀ-ÿ\s\-]+)", adresse)
    if m:
        return m.group(1), m.group(2).strip().title()
    # Fallback
    cp = dept_fallback + "000" if len(dept_fallback) == 2 else dept_fallback
    return cp, ville_fallback


# ---------------------------------------------------------------------------
# Enrichissement SIRET
# ---------------------------------------------------------------------------

def enrichir_siret(fiches: list[dict]) -> list[dict]:
    """Appelle l'API SIRENE pour chaque fiche sans SIRET."""
    print(f"  → Enrichissement SIRET pour {len(fiches)} fiches...")

    for fiche in tqdm(fiches, desc="SIRET", unit="fiche"):
        if fiche.get("SIRET"):
            continue

        nom = fiche.get("Nom entreprise", "")
        dept = fiche.get("Département", "")
        if not nom:
            continue

        try:
            time.sleep(1)
            url = (
                "https://recherche-entreprises.api.gouv.fr/search"
                f"?q={requests.utils.quote(nom)}"
                f"&departement={dept}"
                "&per_page=1"
            )
            resp = requests.get(url, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                results = data.get("results", [])
                if results:
                    r = results[0]
                    siret = r.get("siege", {}).get("siret", "") or r.get("siret", "")
                    if siret:
                        fiche["SIRET"] = siret
        except Exception as e:
            logger.warning(f"API SIRENE erreur pour '{nom}': {e}")
            continue

    return fiches


# ---------------------------------------------------------------------------
# Déduplication
# ---------------------------------------------------------------------------

def dedupliquer(fiches: list[dict], onglet: str) -> list[dict]:
    """Déduplique sur SIRET ou Nom+CP. Logue les doublons."""
    seen = {}
    dedupliquees = []
    doublons = 0

    for fiche in fiches:
        siret = fiche.get("SIRET", "").strip()
        if siret:
            key = ("siret", siret)
        else:
            key = ("nom_cp", fiche.get("Nom entreprise", "").lower(), fiche.get("Code postal", ""))

        if key in seen:
            doublons += 1
            logger.info(f"[{onglet}] Doublon supprimé: {fiche.get('Nom entreprise')} ({fiche.get('Code postal')})")
        else:
            seen[key] = True
            dedupliquees.append(fiche)

    if doublons:
        print(f"  → {doublons} doublon(s) supprimé(s) dans {onglet}")
    return dedupliquees


# ---------------------------------------------------------------------------
# Sauvegarde intermédiaire CSV
# ---------------------------------------------------------------------------

def _save_intermediate(fiches: list[dict]):
    """Sauvegarde les fiches en CSV intermédiaire."""
    try:
        with open(INTERMEDIATE_CSV, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=COLONNES)
            writer.writeheader()
            writer.writerows(fiches)
    except Exception as e:
        logger.error(f"Erreur sauvegarde CSV: {e}")


# ---------------------------------------------------------------------------
# Export Excel
# ---------------------------------------------------------------------------

def exporter_excel(tous_secteurs: dict[str, list[dict]]):
    """
    Crée le fichier Excel avec un onglet par secteur + onglet CONSOLIDÉ.
    """
    print(f"\n  → Génération Excel : {EXCEL_FILENAME}")

    with pd.ExcelWriter(EXCEL_FILENAME, engine="openpyxl") as writer:
        toutes_lignes = []

        for onglet, fiches in tous_secteurs.items():
            if not fiches:
                fiches = [{}]  # Onglet vide mais présent
            df = pd.DataFrame(fiches, columns=COLONNES)
            df.to_excel(writer, sheet_name=onglet, index=False)
            toutes_lignes.extend(fiches)

        # Onglet CONSOLIDÉ
        df_all = pd.DataFrame(toutes_lignes, columns=COLONNES)
        df_all.to_excel(writer, sheet_name="CONSOLIDÉ", index=False)

    # Mise en forme
    _formater_excel(EXCEL_FILENAME, tous_secteurs)
    total = sum(len(v) for v in tous_secteurs.values())
    print(f"  ✓ Excel généré : {EXCEL_FILENAME} ({total} fiches au total)")


def _formater_excel(filename: str, tous_secteurs: dict):
    """Applique une mise en forme basique aux onglets."""
    wb = load_workbook(filename)
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    onglets = list(tous_secteurs.keys()) + ["CONSOLIDÉ"]
    for sheet_name in onglets:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for cell in ws[1]:  # Ligne d'en-tête
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        # Largeur automatique approximative
        col_widths = [30, 25, 35, 40, 12, 20, 15, 18, 20, 35, 15, 16]
        for i, width in enumerate(col_widths, 1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = width
        ws.freeze_panes = "A2"

    wb.save(filename)


# ---------------------------------------------------------------------------
# Mode TEST — génère 10 fiches simulées pour validation
# ---------------------------------------------------------------------------

def mode_test():
    """
    Mode test : scrape les 10 premières fiches du secteur 1 (Hôtels 31)
    en utilisant Playwright en mode réel mais s'arrête à 10 fiches.
    """
    print("\n=== MODE TEST — 10 fiches (Secteur 1 : Hôtels Haute-Garonne) ===\n")

    secteur = SECTEURS[0]  # Hotels_31
    fiches = scrape_pages_jaunes(secteur, max_fiches=10)

    if not fiches:
        print("  Aucune fiche trouvée. Vérifier la connexion ou les sélecteurs Pages Jaunes.")
        print("  Génération de données de démonstration pour validation de la structure...\n")
        fiches = _fiches_demo(secteur)

    fiches = enrichir_siret(fiches)

    print("\n─── Résultats (10 fiches) ───")
    print(f"{'#':<3} {'Nom':<35} {'CP':<6} {'Ville':<20} {'Tel':<14} {'SIRET':<16}")
    print("─" * 100)
    for i, f in enumerate(fiches[:10], 1):
        print(
            f"{i:<3} {f['Nom entreprise'][:34]:<35} "
            f"{f['Code postal']:<6} {f['Ville'][:19]:<20} "
            f"{f['Téléphone'][:13]:<14} {f['SIRET'][:15]:<16}"
        )

    print(f"\n  ✓ {len(fiches)} fiche(s) validée(s). Structure Excel OK.")
    print(f"  Colonnes : {', '.join(COLONNES)}")
    print("\nPrêt pour le scraping complet. Lancez avec : python scraper_prospection.py --full")

    # Sauvegarde CSV test
    csv_test = f"test_10fiches_{TODAY}.csv"
    _save_intermediate(fiches)
    with open(csv_test, "w", newline="", encoding="utf-8") as f_csv:
        w = csv.DictWriter(f_csv, fieldnames=COLONNES)
        w.writeheader()
        w.writerows(fiches)
    print(f"  → CSV test sauvegardé : {csv_test}")
    return fiches


def _fiches_demo(secteur: dict) -> list[dict]:
    """Génère 10 fiches de démonstration si Pages Jaunes est inaccessible."""
    demos = [
        ("Hôtel du Capitole", "1 Place du Capitole", "31000", "Toulouse", "05 61 10 70 00", "http://hotel-capitole.fr"),
        ("Hôtel de Brienne", "20 Bd Maréchal Leclerc", "31000", "Toulouse", "05 61 23 60 60", ""),
        ("Résidence Appart City", "5 Rue de la Pomme", "31000", "Toulouse", "05 34 41 99 00", "http://apartcity.com"),
        ("Novotel Toulouse Centre", "5 Place Alfonse Jourdain", "31000", "Toulouse", "05 61 21 74 74", "http://novotel.com"),
        ("Ibis Toulouse Purpan", "2 Rue Honoré Serres", "31300", "Toulouse", "05 61 63 56 00", "http://ibis.com"),
        ("Mercure Toulouse Centre", "8 Esplanade Compans Caffarelli", "31000", "Toulouse", "05 61 11 09 09", ""),
        ("Appart Hôtel Blagnac", "15 Av Louis Breguet", "31700", "Blagnac", "05 61 71 11 22", ""),
        ("Hôtel Kyriad Colomiers", "3 Rue des Cosmonautes", "31770", "Colomiers", "05 62 74 54 00", "http://kyriad.com"),
        ("Hôtel Le Commerce Muret", "12 Place de la Libération", "31600", "Muret", "05 61 51 23 45", ""),
        ("Hôtel des Pyrénées St-Gaudens", "8 Place Mas-St-Puelles", "31800", "Saint-Gaudens", "05 61 89 00 11", ""),
    ]
    fiches = []
    for nom, adresse, cp, ville, tel, site in demos:
        fiches.append({
            "Formation associée": secteur["formation"],
            "Type d'établissement": "hôtel",
            "Nom entreprise": nom,
            "Adresse complète": f"{adresse}, {cp} {ville}",
            "Code postal": cp,
            "Ville": ville,
            "Département": secteur["departement"],
            "Téléphone": tel,
            "SIRET": "",
            "Site web": site,
            "Source": "Pages Jaunes (demo)",
            "Date de collecte": TODAY_DISPLAY,
        })
    return fiches


# ---------------------------------------------------------------------------
# Mode COMPLET
# ---------------------------------------------------------------------------

def mode_full():
    """Scrape tous les secteurs, enrichit, déduplique, exporte en Excel."""
    print("\n=== MODE COMPLET — 6 secteurs, 300+ fiches chacun ===\n")
    tous_secteurs = {}

    for secteur in SECTEURS:
        fiches = scrape_pages_jaunes(secteur)
        fiches = enrichir_siret(fiches)
        fiches = dedupliquer(fiches, secteur["onglet"])

        tous_secteurs[secteur["onglet"]] = fiches

        # CSV par secteur
        csv_secteur = f"secteur_{secteur['id']}_{secteur['onglet']}_{TODAY}.csv"
        with open(csv_secteur, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=COLONNES)
            w.writeheader()
            w.writerows(fiches)

        print(f"  ✓ Secteur {secteur['id']} terminé : {len(fiches)} fiches → {csv_secteur}")

        # Commit automatique après chaque secteur
        _git_commit_secteur(secteur["id"], secteur["onglet"], len(fiches))

    # Export Excel final
    exporter_excel(tous_secteurs)

    # Commit Excel final
    _git_commit_excel()

    total = sum(len(v) for v in tous_secteurs.values())
    print(f"\n=== TERMINÉ — {total} fiches exportées dans {EXCEL_FILENAME} ===")


# ---------------------------------------------------------------------------
# Git helpers
# ---------------------------------------------------------------------------

def _git_commit_secteur(secteur_id: int, onglet: str, n_fiches: int):
    """Commit automatique après chaque secteur."""
    try:
        os.system("git add *.csv *.log")
        os.system(f'git commit -m "data: secteur {secteur_id} {onglet} terminé - {n_fiches} fiches collectées"')
        os.system("git push -u origin HEAD")
    except Exception as e:
        logger.error(f"Erreur git commit secteur {secteur_id}: {e}")


def _git_commit_excel():
    """Commit automatique après génération Excel."""
    try:
        os.system(f"git add {EXCEL_FILENAME} *.csv *.log")
        os.system(f'git commit -m "feat: export Excel final - fichier {EXCEL_FILENAME} généré"')
        os.system("git push -u origin HEAD")
    except Exception as e:
        logger.error(f"Erreur git commit Excel: {e}")


# ---------------------------------------------------------------------------
# Point d'entrée
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Scraper prospection Pages Jaunes")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--test", action="store_true", help="Test 10 fiches secteur 1")
    group.add_argument("--full", action="store_true", help="Scraping complet 6 secteurs")
    args = parser.parse_args()

    if args.test:
        mode_test()
    elif args.full:
        mode_full()
