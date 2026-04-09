#!/usr/bin/env python3
"""
script_lba_lbb.py — Prospection LBA + LBB
==========================================
Interroge La Bonne Alternance (offres actives) et La Bonne Boîte (potentiel
de recrutement) pour qualifier des leads entreprises par ville, codes NAF et
rayon.

Usage :
  python script_lba_lbb.py --ville Paris --naf 10.71A 10.71B --rayon 30
  python script_lba_lbb.py --config params.json
"""

import argparse
import json
import os
import re
import sys
import time
from datetime import date

import pandas as pd
import requests

# ── Table NAF → ROME ─────────────────────────────────────────────────────────

NAF_TO_ROME = {
    # Métiers de bouche
    "10.13A": ["D1103"], "10.13B": ["D1103"],
    "10.71A": ["D1102"], "10.71B": ["D1102"],
    "10.71C": ["D1102", "D1104"], "10.71D": ["D1104"],
    "56.10A": ["G1602", "G1603"], "56.10B": ["G1603"],
    "56.10C": ["G1603"], "56.21Z": ["G1602"],
    # Hôtellerie
    "55.10Z": ["G1703", "G1501"],
    # Commerce alimentaire
    "47.11B": ["D1106"], "47.11C": ["D1106"],
    "47.11D": ["D1504", "D1106"], "47.11F": ["D1504"],
    "47.22Z": ["D1103"], "47.24Z": ["D1106"],
    # Automobile
    "45.11Z": ["D1404"], "45.20A": ["I1604"], "45.20B": ["I1604"],
    # Immobilier
    "68.31Z": ["C1504", "C1501"], "68.32A": ["C1502"],
    # Comptabilité / Juridique
    "69.20Z": ["M1203"], "69.10Z": ["K1903"],
    # Coiffure / Esthétique
    "96.02A": ["D1202"], "96.02B": ["D1208"],
    # Pharmacie
    "47.73Z": ["J1307"],
    # BTP
    "43.21A": ["F1602"], "43.22A": ["F1603"],
    "43.31Z": ["F1611"], "43.34Z": ["F1606"],
    # Informatique
    "62.01Z": ["M1805", "M1802"],
}

# ── Fallback géocodage ────────────────────────────────────────────────────────

VILLES_FALLBACK = {
    "paris":       (48.8566, 2.3522),
    "lyon":        (45.7640, 4.8357),
    "marseille":   (43.2965, 5.3698),
    "toulouse":    (43.6047, 1.4442),
    "bordeaux":    (44.8378, -0.5792),
    "lille":       (50.6292, 3.0573),
    "nice":        (43.7102, 7.2620),
    "nantes":      (47.2184, -1.5536),
    "strasbourg":  (48.5734, 7.7521),
    "montpellier": (43.6108, 3.8767),
    "rennes":      (48.1173, -1.6778),
    "toulon":      (43.1242, 5.9280),
    "grenoble":    (45.1885, 5.7245),
    "dijon":       (47.3220, 5.0415),
    "angers":      (47.4784, -0.5632),
    "reims":       (49.2583, 4.0317),
    "le mans":     (48.0061, 0.1996),
    "aix-en-provence": (43.5297, 5.4474),
    "clermont-ferrand": (45.7772, 3.0870),
    "tours":       (47.3941, 0.6848),
    "boulogne-billancourt": (48.8397, 2.2399),
    "paris 20":    (48.8637, 2.3985),
    "paris 20e":   (48.8637, 2.3985),
}

# ── API config ────────────────────────────────────────────────────────────────

LBA_BASE_URL = "https://api.apprentissage.beta.gouv.fr/api/job/v1"
LBA_TOKEN = (
    "eyJhbGciOiJIUzI1NiJ9."
    "eyJfaWQiOiI2OGM4MzJkZDgxZGY5MmFiYTc2MDNhNzAiLCJhcGlfa2V5IjoieVFaYkpiZElCN1VydDNvb2"
    "w3aTRqN0lSRUhSQ25KSk5ya0djclpxZ1E0bz0iLCJvcmdhbmlzYXRpb24iOiJTa2lsbGFuZHlvdSIsImVt"
    "YWlsIjoiam9mZnJleS5sYWRtaXJhdWx0QHNraWxsYW5keW91LmNvbSIsImlzcyI6ImFwaSIsImlhdCI6MTc"
    "3MzMyMzg1NSwiZXhwIjoxNzg5NDg2Njk2fQ."
    "5P4D4gmCTezPIvycQEtKnPZBjPvcx8BHG1bF8r_Z4ts"
)
CALLER = "skill-and-you-prospection"
API_DELAY = 0.8
MAX_RETRIES = 3
REQUEST_TIMEOUT = 30

# ── Colonnes de sortie ───────────────────────────────────────────────────────

EXCEL_COLUMNS = [
    "Nom de l'entreprise", "SIRET", "Code NAF", "Adresse", "Ville",
    "Code Postal", "Ville de recherche", "Source", "Score LBB",
    "Offres actives", "Date de collecte",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def normalize_siret(raw) -> str:
    if raw is None or str(raw).strip() in ("", "nan", "None"):
        return ""
    try:
        return str(int(float(str(raw).strip()))).zfill(14)
    except (ValueError, OverflowError):
        return str(raw).strip()


def http_get(url, params=None, headers=None):
    """GET avec retry + backoff exponentiel."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(
                url, params=params, headers=headers, timeout=REQUEST_TIMEOUT,
            )
            if resp.status_code in (429, 500, 502, 503, 504):
                wait = 2 ** attempt
                print(f"  ⚠️  HTTP {resp.status_code} — retry {attempt}/{MAX_RETRIES} dans {wait}s")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            ct = resp.headers.get("content-type", "")
            if "html" in ct:
                print(f"  ❌ Réponse HTML inattendue — {url}")
                return None
            return resp.json()
        except requests.exceptions.Timeout:
            wait = 2 ** attempt
            print(f"  ⚠️  Timeout — retry {attempt}/{MAX_RETRIES} dans {wait}s")
            time.sleep(wait)
        except requests.exceptions.ConnectionError as e:
            wait = 2 ** attempt
            print(f"  ⚠️  Connexion : {e} — retry {attempt}/{MAX_RETRIES} dans {wait}s")
            time.sleep(wait)
        except requests.exceptions.HTTPError as e:
            print(f"  ❌ HTTP {e.response.status_code} — {url}")
            return None
        except ValueError:
            print(f"  ❌ Réponse non-JSON — {url}")
            return None
    print(f"  ❌ Échec après {MAX_RETRIES} tentatives — {url}")
    return None


# ── Géocodage ─────────────────────────────────────────────────────────────────

def geocode(ville: str) -> tuple[float, float]:
    """Retourne (latitude, longitude) pour la ville donnée."""
    print(f"📡 Géocodage de « {ville} »…", end=" ", flush=True)
    url = "https://api-adresse.data.gouv.fr/search/"
    data = http_get(url, params={"q": ville, "limit": 1, "type": "municipality"})
    if data and data.get("features"):
        coords = data["features"][0]["geometry"]["coordinates"]
        lon, lat = coords[0], coords[1]
        print(f"✅ ({lat:.4f}, {lon:.4f})")
        return lat, lon

    # Fallback table locale
    key = ville.lower().strip()
    if key in VILLES_FALLBACK:
        lat, lon = VILLES_FALLBACK[key]
        print(f"✅ fallback ({lat:.4f}, {lon:.4f})")
        return lat, lon

    print("❌ Ville introuvable")
    sys.exit(1)


# ── Résolution NAF → ROME ────────────────────────────────────────────────────

def resolve_rome_codes(naf_codes: list[str]) -> list[str]:
    """Convertit une liste de codes NAF en codes ROME uniques."""
    rome_set = set()
    for naf in naf_codes:
        if naf in NAF_TO_ROME:
            rome_set.update(NAF_TO_ROME[naf])
        else:
            print(f"  ⚠️  Code NAF « {naf} » inconnu — ignoré")
    return sorted(rome_set)


# ── Parsing réponses API ─────────────────────────────────────────────────────

def _parse_address_parts(full_address: str) -> tuple[str, str, str]:
    """Extrait (adresse, code_postal, ville) depuis une adresse complète."""
    if not full_address:
        return "", "", ""
    m = re.search(r'(\d{5})\s+(.+)$', full_address.strip())
    if m:
        return (
            full_address[:m.start()].strip().rstrip(",").strip(),
            m.group(1),
            m.group(2).strip(),
        )
    return full_address, "", ""


def _extract_lba_company(item: dict) -> dict | None:
    """Extrait les champs d'une entreprise depuis une offre LBA (jobs/matchas)."""
    # Essaye les deux structures possibles
    wp = item.get("workplace") or {}
    company = item.get("company") or {}
    place = item.get("place") or {}
    apply_info = item.get("apply") or {}

    siret = normalize_siret(
        wp.get("siret") or company.get("siret") or ""
    )
    if not siret:
        return None

    name = (
        wp.get("name") or wp.get("legal_name") or wp.get("brand")
        or company.get("name") or ""
    )
    naf_obj = (wp.get("domain") or {}).get("naf") or {}
    naf = naf_obj.get("code") or company.get("naf") or ""

    # Adresse
    full_addr = (
        (wp.get("location") or {}).get("address")
        or place.get("fullAddress") or ""
    )
    addr, cp, ville = _parse_address_parts(full_addr)
    ville = place.get("city") or ville
    cp = place.get("zipCode") or cp

    return {
        "Nom de l'entreprise": name,
        "SIRET": siret,
        "Code NAF": naf,
        "Adresse": addr,
        "Ville": ville,
        "Code Postal": cp,
    }


def _extract_lbb_company(item: dict) -> dict | None:
    """Extrait les champs d'une entreprise depuis un résultat LBB (lbaCompanies)."""
    company = item.get("company") or {}
    place = item.get("place") or {}
    wp = item.get("workplace") or {}

    siret = normalize_siret(
        company.get("siret") or wp.get("siret") or ""
    )
    if not siret:
        return None

    name = company.get("name") or wp.get("name") or ""
    naf_obj = (wp.get("domain") or {}).get("naf") or {}
    naf = company.get("naf") or naf_obj.get("code") or ""

    full_addr = (
        place.get("fullAddress")
        or (wp.get("location") or {}).get("address") or ""
    )
    addr, cp, ville = _parse_address_parts(full_addr)
    ville = place.get("city") or ville
    cp = place.get("zipCode") or cp

    score = company.get("score") or item.get("hiring_potential") or ""

    return {
        "Nom de l'entreprise": name,
        "SIRET": siret,
        "Code NAF": naf,
        "Adresse": addr,
        "Ville": ville,
        "Code Postal": cp,
        "Score LBB": str(score) if score else "",
    }


# ── Appels API ────────────────────────────────────────────────────────────────

def fetch_search(rome_codes: list[str], lat: float, lon: float, radius: int) -> tuple[list[dict], list[dict]]:
    """
    Appelle /search pour chaque code ROME.
    L'API retourne { jobs: [...], recruiters: [...] }.
    - jobs = offres LBA actives
    - recruiters = entreprises LBB (potentiel recrutement)
    Retourne (lba_rows, lbb_rows) brutes.
    """
    lba_rows = []
    lbb_rows = []
    today = date.today().isoformat()
    headers = {"Authorization": f"Bearer {LBA_TOKEN}"}

    for rome in rome_codes:
        print(f"  📡 /search ROME={rome}…", end=" ", flush=True)
        data = http_get(
            f"{LBA_BASE_URL}/search",
            params={
                "romes": rome,
                "latitude": lat,
                "longitude": lon,
                "radius": radius,
            },
            headers=headers,
        )
        if data is None:
            print("aucune donnée")
            time.sleep(API_DELAY)
            continue

        jobs = data.get("jobs") or []
        recruiters = data.get("recruiters") or []

        lba_count = 0
        for item in jobs:
            row = _extract_lba_company(item)
            if row:
                row["Source"] = "LBA"
                row["Offres actives"] = 1
                row["Score LBB"] = ""
                row["Date de collecte"] = today
                lba_rows.append(row)
                lba_count += 1

        lbb_count = 0
        for item in recruiters:
            row = _extract_lbb_company(item)
            if row:
                row["Source"] = "LBB"
                row["Offres actives"] = ""
                row.setdefault("Score LBB", "")
                row["Date de collecte"] = today
                lbb_rows.append(row)
                lbb_count += 1

        print(f"✅ {lba_count} LBA + {lbb_count} LBB")
        time.sleep(API_DELAY)

    return lba_rows, lbb_rows


# ── Déduplication ─────────────────────────────────────────────────────────────

def deduplicate(lba_rows: list[dict], lbb_rows: list[dict]) -> tuple[list[dict], int]:
    """
    Déduplique par SIRET. LBA prioritaire sur LBB.
    Retourne (rows_uniques, nb_doublons).
    """
    by_siret: dict[str, dict] = {}
    doublons = 0

    # LBA en premier (prioritaire)
    for row in lba_rows:
        siret = row.get("SIRET", "")
        if not siret:
            continue
        if siret in by_siret:
            # Même source, on cumule les offres
            existing = by_siret[siret]
            existing["Offres actives"] = (
                (existing.get("Offres actives") or 0) + (row.get("Offres actives") or 0)
            )
        else:
            by_siret[siret] = row.copy()

    # LBB ensuite
    for row in lbb_rows:
        siret = row.get("SIRET", "")
        if not siret:
            continue
        if siret in by_siret:
            # Doublon LBA+LBB
            existing = by_siret[siret]
            if existing["Source"] == "LBA":
                existing["Source"] = "LBA + LBB"
                existing["Score LBB"] = row.get("Score LBB") or existing.get("Score LBB", "")
                doublons += 1
            # Combler les champs vides
            for col in EXCEL_COLUMNS:
                if not str(existing.get(col, "")).strip() and str(row.get(col, "")).strip():
                    existing[col] = row[col]
        else:
            by_siret[siret] = row.copy()

    return list(by_siret.values()), doublons


# ── Export ────────────────────────────────────────────────────────────────────

def save_backup_csv(rows: list[dict], filepath: str):
    """Sauvegarde intermédiaire CSV."""
    df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    df.to_csv(filepath, index=False, encoding="utf-8-sig")


def export_results(rows: list[dict], output_dir: str, ville: str):
    """Exporte Excel + CSV backup."""
    today_str = date.today().strftime("%Y%m%d")
    ville_slug = ville.lower().replace(" ", "_").replace("-", "_")
    base_name = f"lba_lbb_results_{ville_slug}_{today_str}"
    xlsx_path = os.path.join(output_dir, f"{base_name}.xlsx")
    csv_path = os.path.join(output_dir, f"{base_name}.csv")

    df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    # SIRET en string
    df["SIRET"] = df["SIRET"].astype(str)

    df.to_excel(xlsx_path, index=False, engine="openpyxl")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    return xlsx_path, csv_path


# ── Synthèse console ─────────────────────────────────────────────────────────

def print_synthese(ville: str, rayon: int, naf_codes: list[str],
                   rome_codes: list[str], rows: list[dict],
                   n_lba: int, n_lbb: int, n_doublons: int,
                   xlsx_path: str, csv_path: str):
    n_total = len(rows)
    print()
    print("══════════════════════════════════════════════════")
    print(f"  RÉSULTATS — {ville} (rayon {rayon} km)")
    print(f"  Codes NAF : {', '.join(naf_codes)} → ROME : {', '.join(rome_codes)}")
    print()
    print(f"  LBA (offres actives)     : {n_lba} entreprises")
    print(f"  LBB (potentiel)          : {n_lbb} entreprises")
    print(f"  Doublons LBA+LBB         : {n_doublons} entreprises")
    print("  ──────────────────────────────────────────────")
    print(f"  Total unique (par SIRET) : {n_total} entreprises")
    print()
    print(f"  📁 Fichier : {os.path.basename(xlsx_path)}")
    print(f"  📁 Backup  : {os.path.basename(csv_path)}")
    print("══════════════════════════════════════════════════")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Prospection LBA + LBB — collecte entreprises par ville et codes NAF",
    )
    parser.add_argument("--ville", type=str, help="Nom de la ville")
    parser.add_argument("--naf", nargs="+", help="Codes NAF (ex: 10.71A 10.71B)")
    parser.add_argument("--rayon", type=int, default=30, help="Rayon en km (défaut: 30)")
    parser.add_argument("--output", type=str, default=".", help="Répertoire de sortie")
    parser.add_argument("--config", type=str, help="Fichier JSON de paramètres")
    return parser.parse_args()


def load_config(args) -> tuple[list[str], list[str], int, str]:
    """Retourne (villes, naf_codes, rayon, output_dir)."""
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
        naf_codes = cfg["codes_naf"]
        rayon = cfg.get("rayon_km", 30)
        output_dir = cfg.get("output", ".")
    elif args.ville and args.naf:
        villes = [args.ville]
        naf_codes = args.naf
        rayon = args.rayon
        output_dir = args.output
    else:
        print("❌ Spécifiez --ville et --naf, ou --config")
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)
    return villes, naf_codes, rayon, output_dir


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    villes, naf_codes, rayon, output_dir = load_config(args)
    multi = len(villes) > 1

    print("══════════════════════════════════════════════════")
    print("  PROSPECTION LBA + LBB")
    print(f"  Villes : {', '.join(villes)}")
    print(f"  NAF    : {', '.join(naf_codes)}")
    print(f"  Rayon  : {rayon} km")
    print("══════════════════════════════════════════════════")

    # 1. Résolution NAF → ROME
    rome_codes = resolve_rome_codes(naf_codes)
    if not rome_codes:
        print("❌ Aucun code ROME trouvé pour les NAF fournis")
        sys.exit(1)
    print(f"\n✅ Codes ROME : {', '.join(rome_codes)}")

    # 2. Boucle sur les villes
    all_city_results: dict[str, list[dict]] = {}
    city_stats: dict[str, dict] = {}
    today_str = date.today().strftime("%Y%m%d")

    for ville in villes:
        print(f"\n{'═' * 50}")
        print(f"  📍 {ville}")
        print(f"{'═' * 50}")

        try:
            lat, lon = geocode(ville)

            effective_radius = rayon
            print(f"\n── Collecte via /search (rayon {effective_radius} km) ──")
            lba_rows, lbb_rows = fetch_search(rome_codes, lat, lon, effective_radius)

            # Rayon 0 : si aucun résultat, retry avec 1 km
            if effective_radius == 0 and not lba_rows and not lbb_rows:
                print("  ⚠️  Rayon 0 km : aucun résultat — retry avec 1 km")
                lba_rows, lbb_rows = fetch_search(rome_codes, lat, lon, 1)

            # Ajouter "Ville de recherche"
            for row in lba_rows + lbb_rows:
                row["Ville de recherche"] = ville

            # Backup CSV après chaque ville
            ville_slug = ville.lower().replace(" ", "_").replace("-", "_")
            backup_path = os.path.join(
                output_dir, f"lba_lbb_{ville_slug}_{today_str}_backup.csv",
            )
            save_backup_csv(lba_rows + lbb_rows, backup_path)

            # Stats brutes
            lba_sirets = {r["SIRET"] for r in lba_rows if r.get("SIRET")}
            lbb_sirets = {r["SIRET"] for r in lbb_rows if r.get("SIRET")}

            # Déduplication intra-ville
            rows, n_doublons = deduplicate(lba_rows, lbb_rows)
            for row in rows:
                row.setdefault("Ville de recherche", ville)

            city_stats[ville] = {
                "n_lba": len(lba_sirets),
                "n_lbb": len(lbb_sirets - lba_sirets),
                "n_doublons": n_doublons,
                "n_unique": len(rows),
            }
            all_city_results[ville] = rows
            print(f"  ✅ {len(rows)} entreprises uniques pour {ville}")

            # Nettoyage backup
            if os.path.exists(backup_path):
                os.remove(backup_path)

        except Exception as e:
            print(f"  ❌ Erreur pour {ville} : {e}")
            all_city_results[ville] = []
            city_stats[ville] = {
                "n_lba": 0, "n_lbb": 0, "n_doublons": 0, "n_unique": 0,
            }
            continue

    # 3. Vérification globale
    if not any(all_city_results.values()):
        print("\n⚠️  Aucune entreprise collectée pour aucune ville.")
        sys.exit(0)

    # 4. Export
    if multi:
        xlsx_path = os.path.join(output_dir, f"lba_lbb_results_multi_{today_str}.xlsx")
        csv_path = xlsx_path.replace(".xlsx", ".csv")

        # Consolidé : dédup globale par SIRET
        all_rows = []
        for rows in all_city_results.values():
            all_rows.extend(rows)
        seen_sirets: set[str] = set()
        consolidated: list[dict] = []
        for row in all_rows:
            siret = row.get("SIRET", "")
            if siret and siret in seen_sirets:
                continue
            if siret:
                seen_sirets.add(siret)
            consolidated.append(row)

        columns = EXCEL_COLUMNS
        with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
            for ville, rows in all_city_results.items():
                if rows:
                    df = pd.DataFrame(rows, columns=columns)
                    df["SIRET"] = df["SIRET"].astype(str)
                    df.to_excel(writer, sheet_name=ville[:31], index=False)
            df_all = pd.DataFrame(consolidated, columns=columns)
            df_all["SIRET"] = df_all["SIRET"].astype(str)
            df_all.to_excel(writer, sheet_name="Consolidé", index=False)

        df_all.to_csv(csv_path, index=False, encoding="utf-8-sig")

        print(f"\n── Export multi-villes ──────────────────────────")
        print(f"  ✅ {xlsx_path}")
        print(f"  ✅ {csv_path}")
        print()
        print("══════════════════════════════════════════════════")
        print("  RÉSULTATS MULTI-VILLES — LBA + LBB")
        print(f"  NAF : {', '.join(naf_codes)} → ROME : {', '.join(rome_codes)}")
        print()
        for ville, st in city_stats.items():
            print(f"  📍 {ville}: {st['n_unique']} uniques "
                  f"({st['n_lba']} LBA, {st['n_lbb']} LBB, {st['n_doublons']} doublons)")
        print("  ──────────────────────────────────────────────")
        print(f"  Total consolidé (dédup SIRET) : {len(consolidated)} entreprises")
        print(f"  📁 Fichier : {os.path.basename(xlsx_path)}")
        print("══════════════════════════════════════════════════")
    else:
        # Mode mono-ville (rétrocompatible)
        ville = villes[0]
        rows = all_city_results.get(ville, [])
        if not rows:
            print("\n⚠️  Aucune entreprise collectée.")
            sys.exit(0)
        xlsx_path, csv_path = export_results(rows, output_dir, ville)
        st = city_stats[ville]
        print_synthese(
            ville, rayon, naf_codes, rome_codes, rows,
            st["n_lba"], st["n_lbb"], st["n_doublons"],
            xlsx_path, csv_path,
        )


if __name__ == "__main__":
    main()
