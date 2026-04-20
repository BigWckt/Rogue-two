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

from batch_io import (
    detect_secteur, find_prosp_root, next_batch_number,
    write_current_batch, SECTEUR_PREFIXES,
)
from matrix_display import (
    GREEN, RED, BOLD, RESET,
    matrix_banner, matrix_section, matrix_kv, matrix_separator,
    matrix_step, matrix_ok, matrix_fail, matrix_warn,
    morpheus_says, ask_filename, matrix_decode, set_quiet,
    matrix_rain_fullscreen,
)

# ── Table NAF → ROME ─────────────────────────────────────────────────────────

NAF_TO_ROME = {
    # Métiers de bouche
    "10.13A": ["D1103"], "10.13B": ["D1103"],
    "10.71A": ["D1102"], "10.71B": ["D1102"],
    "10.71C": ["D1102", "D1104"], "10.71D": ["D1104"],
    "56.10A": ["G1602", "G1603"], "56.10B": ["G1603"],
    "56.10C": ["G1603"], "56.21Z": ["G1602"],
    "50.10A": ["G1204"],  # Transports maritimes (inclus à la demande dans mbt_restaurants)
    # Hôtellerie
    "55.10Z": ["G1703", "G1501"],
    # Commerce alimentaire
    "47.11A": ["D1505"],
    "47.11B": ["D1106"], "47.11C": ["D1106"],
    "47.11D": ["D1504", "D1106"], "47.11E": ["D1505"], "47.11F": ["D1504"],
    "47.19A": ["D1502"], "47.19B": ["D1502"],
    "47.21Z": ["D1106"], "47.22Z": ["D1103"],
    "47.23Z": ["D1106"], "47.24Z": ["D1106"], "47.25Z": ["D1106"],
    # Automobile
    "45.11Z": ["D1404"], "45.20A": ["I1604"], "45.20B": ["I1604"],
    "45.31Z": ["G1404", "I1604"], "45.32Z": ["G1404"],
    "45.40Z": ["I1607"],
    # Immobilier
    "68.31Z": ["C1504", "C1501"], "68.32A": ["C1502"],
    # Comptabilité / Juridique
    "69.20Z": ["M1203"], "69.10Z": ["K1903"],
    # Coiffure / Esthétique
    "96.02A": ["D1202"], "96.02B": ["D1208"],
    # Pharmacie
    "47.73Z": ["J1307"],
    # Fleuriste
    "47.76Z": ["D1209"],
    # BTP
    "43.21A": ["F1602"], "43.21B": ["F1602"],
    "43.22A": ["F1603"], "43.22B": ["F1603"],
    "43.29A": ["F1601"], "43.29B": ["F1601"],
    "43.31Z": ["F1611"], "43.32B": ["F1607"],
    "43.33Z": ["F1603"], "43.34Z": ["F1606"],
    "43.99B": ["F1611"], "43.99C": ["F1606"], "43.99D": ["F1612"],
    "41.20A": ["F1201"], "41.20B": ["F1201"],
    "74.90A": ["F1106"], "71.12B": ["F1106"],
    "33.15Z": ["I1604"], "36.00Z": ["F1603"],
    "16.23Z": ["F1607"],
    "31.01Z": ["H2206"], "31.02Z": ["H2206"],
    # Informatique
    "62.01Z": ["M1805", "M1802"],
    # ─── Secteur Santé (NAF 86.xx) ───
    "86.10A": ["J1306", "J1501", "J1502", "J1506", "J1507"],
    "86.10Z": ["J1306", "J1501", "J1502", "J1506", "J1507"],
    "86.21Z": ["J1102"],
    "86.22A": ["J1103"],
    "86.22B": ["J1104"],
    "86.22C": ["J1102", "J1104"],
    "86.23Z": ["J1303"],
    "86.90A": ["J1502"],
    "86.90B": ["J1404"],
    "86.90D": ["J1411"],
    "86.90E": ["J1404", "J1409"],
    "86.90F": ["J1403"],
    # ─── Action sociale avec hébergement (NAF 87.xx) ───
    "87.10A": ["J1501", "J1506", "K1301", "K1302"],
    "87.10C": ["J1501", "K1301"],
    "87.30A": ["K1302", "K1207"],
    "87.30B": ["K1207", "K1301"],
    # ─── Action sociale sans hébergement (NAF 88.xx) ───
    "88.10A": ["K1302", "K1304"],
    "88.91A": ["K1303", "G1202"],
    # ─── Administration publique ───
    "84.11Z": ["K1404", "K1602"],
    # ─── Tertiaire / Grande distribution (NAF 47.xx) ───
    "47.30Z": ["G1502"],
    "47.41Z": ["D1211"], "47.42Z": ["D1211"], "47.43Z": ["D1211"],
    "47.51Z": ["D1212"], "47.52A": ["D1212"], "47.52B": ["D1212"],
    "47.53Z": ["D1212"], "47.54Z": ["D1212"],
    "47.59A": ["D1212"], "47.59B": ["D1212"],
    "47.64Z": ["D1212"], "47.65Z": ["D1212"],
    "47.71Z": ["D1212"], "47.72A": ["D1212"],
    "47.79Z": ["D1212"],
}

# ── Profils sectoriels Skill & You ──────────────────────────────────────────

PROFILS_SKY = {
    # ─── Secteur SB (Services à la personne / Santé / Beauté) ───
    "sb_sante":           ["86.10A", "86.10Z", "86.21Z", "86.22A", "86.22B", "86.22C", "86.23Z", "86.90B"],
    "sb_accueil_enfants": ["88.91A", "84.11Z"],
    "sb_fleuriste":       ["47.76Z"],
    "sb_service_personne":["88.10A", "87.10A", "87.10C", "87.30A"],
    "sb_coiffure":        ["96.02A"],
    "sb_esthetique":      ["96.02B"],

    # ─── Secteur BTPM (Bâtiment / Mécanique) ───
    "btpm_batiment":      ["43.99B", "74.90A", "43.99C", "41.20A", "41.20B", "43.33Z", "43.32B", "71.12B", "84.11Z"],
    "btpm_electricite":   ["43.21A", "43.21B", "43.99C", "43.31Z", "43.22B", "43.29B", "36.00Z", "33.15Z", "45.11Z", "41.20A", "84.11Z"],
    "btpm_plomberie":     ["43.22B", "43.22A", "43.21A", "36.00Z", "43.34Z", "41.20B"],
    "btpm_chauffage_clim":["43.22B", "43.22A"],
    "btpm_menuiserie":    ["84.11Z", "43.32B", "16.23Z", "31.02Z", "31.01Z", "43.29A", "41.20A", "43.99D"],
    "btpm_meca_carrosserie":["45.20A", "45.20B", "45.32Z", "45.11Z", "45.31Z", "84.11Z"],
    "btpm_moto":          ["45.40Z"],

    # ─── Secteur TG (Tertiaire / Grand commerce) ───
    "tg_immobilier":      ["68.31Z", "68.32A"],
    "tg_stations_service":["47.30Z"],
    "tg_informatique":    ["47.41Z", "47.42Z", "47.43Z"],
    "tg_vetements_chaussures":["47.51Z", "47.52A", "47.52B", "47.53Z", "47.54Z", "47.59A", "47.59B", "47.71Z", "47.72A"],
    "tg_sport":           ["47.64Z", "47.65Z"],
    "tg_occasions":       ["47.79Z"],

    # ─── Secteur MBT (Métiers de Bouche / Tertiaire) ───
    "mbt_boulangeries":          ["10.71A", "10.71B", "10.71C", "10.71D"],
    "mbt_grande_distribution":   ["47.11A", "47.11B", "47.11C", "47.11D", "47.11E", "47.11F",
                                  "47.19A", "47.19B", "47.21Z", "47.22Z", "47.23Z", "47.24Z", "47.25Z"],
    "mbt_restaurants":           ["56.10A", "50.10A"],
}

# ── Avertissements de cohérence sémantique NAF ───────────────────────────────
# Codes NAF dont la signification réelle peut surprendre l'utilisateur.
NAF_SEMANTIC_WARNINGS = {
    "50.10A": "⚠ 50.10A = 'Transports maritimes et côtiers de passagers' (pas 'Restaurant'). "
              "Inclus dans mbt_restaurants à la demande explicite.",
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

OUTPUT_COLUMNS = [
    "Nom de l'entreprise", "SIRET", "Code NAF", "Adresse", "Ville",
    "Code Postal", "Téléphone", "Ville de recherche", "Source", "Score LBB",
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


def normalize_phone(raw: str) -> str:
    """Nettoie et valide un numéro de téléphone français. Retourne 10 chiffres ou ''."""
    if not raw:
        return ""
    cleaned = re.sub(r"[\s.\-\(\)]+", "", raw.strip())
    if cleaned.startswith("+33"):
        cleaned = "0" + cleaned[3:]
    digits = re.sub(r"\D", "", cleaned)
    if digits and re.match(r"^0[1-9]\d{8}$", digits):
        return digits
    return ""


def extract_phone(entry: dict) -> str:
    """Cherche le téléphone dans plusieurs chemins possibles d'un item API."""
    candidates = [
        (entry.get("apply") or {}).get("phone"),
        ((entry.get("workplace") or {}).get("contact") or {}).get("phone"),
        (entry.get("contact") or {}).get("phone"),
        entry.get("phone"),
        entry.get("telephone"),
    ]
    for c in candidates:
        if c and str(c).strip():
            return normalize_phone(str(c))
    return ""


def http_get(url, params=None, headers=None):
    """GET avec retry + backoff exponentiel."""
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(
                url, params=params, headers=headers, timeout=REQUEST_TIMEOUT,
            )
            if resp.status_code in (429, 500, 502, 503, 504):
                wait = 2 ** attempt
                matrix_warn(f"HTTP {resp.status_code} — retry {attempt}/{MAX_RETRIES} dans {wait}s")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            ct = resp.headers.get("content-type", "")
            if "html" in ct:
                matrix_fail(f"Réponse HTML inattendue — {url}")
                return None
            return resp.json()
        except requests.exceptions.Timeout:
            wait = 2 ** attempt
            matrix_warn(f"Timeout — retry {attempt}/{MAX_RETRIES} dans {wait}s")
            time.sleep(wait)
        except requests.exceptions.ConnectionError as e:
            wait = 2 ** attempt
            matrix_warn(f"Connexion : {e} — retry {attempt}/{MAX_RETRIES} dans {wait}s")
            time.sleep(wait)
        except requests.exceptions.HTTPError as e:
            matrix_fail(f"HTTP {e.response.status_code} — {url}")
            return None
        except ValueError:
            matrix_fail(f"Réponse non-JSON — {url}")
            return None
    matrix_fail(f"Échec après {MAX_RETRIES} tentatives — {url}")
    return None


# ── Géocodage ─────────────────────────────────────────────────────────────────

def geocode(ville: str) -> tuple[float, float]:
    """Retourne (latitude, longitude) pour la ville donnée."""
    matrix_step(f"Géocodage de « {ville} »…")
    url = "https://api-adresse.data.gouv.fr/search/"
    data = http_get(url, params={"q": ville, "limit": 1, "type": "municipality"})
    if data and data.get("features"):
        coords = data["features"][0]["geometry"]["coordinates"]
        lon, lat = coords[0], coords[1]
        matrix_ok(f"Coordonnées : ({lat:.4f}, {lon:.4f})")
        return lat, lon

    # Fallback table locale
    key = ville.lower().strip()
    if key in VILLES_FALLBACK:
        lat, lon = VILLES_FALLBACK[key]
        matrix_ok(f"Fallback local : ({lat:.4f}, {lon:.4f})")
        return lat, lon

    matrix_fail("Ville introuvable")
    sys.exit(1)


# ── Résolution NAF → ROME ────────────────────────────────────────────────────

def resolve_rome_codes(naf_codes: list[str]) -> list[str]:
    """Convertit une liste de codes NAF en codes ROME uniques."""
    rome_set = set()
    for naf in naf_codes:
        if naf in NAF_TO_ROME:
            rome_set.update(NAF_TO_ROME[naf])
        else:
            matrix_warn(f"Code NAF « {naf} » inconnu — ignoré")
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
    wp = item.get("workplace") or {}
    company = item.get("company") or {}
    place = item.get("place") or {}

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

    full_addr = (
        (wp.get("location") or {}).get("address")
        or place.get("fullAddress") or ""
    )
    addr, cp, ville = _parse_address_parts(full_addr)
    ville = place.get("city") or ville
    cp = place.get("zipCode") or cp

    phone = extract_phone(item)

    return {
        "Nom de l'entreprise": name,
        "SIRET": siret,
        "Code NAF": naf,
        "Adresse": addr,
        "Ville": ville,
        "Code Postal": cp,
        "Téléphone": phone,
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

    phone = extract_phone(item)

    return {
        "Nom de l'entreprise": name,
        "SIRET": siret,
        "Code NAF": naf,
        "Adresse": addr,
        "Ville": ville,
        "Code Postal": cp,
        "Téléphone": phone,
        "Score LBB": str(score) if score else "",
    }


# ── Appels API ────────────────────────────────────────────────────────────────

def fetch_search(rome_codes: list[str], lat: float, lon: float, radius: int) -> tuple[list[dict], list[dict]]:
    """
    Appelle /search pour chaque code ROME.
    Retourne (lba_rows, lbb_rows) brutes.
    """
    lba_rows = []
    lbb_rows = []
    today = date.today().isoformat()
    headers = {"Authorization": f"Bearer {LBA_TOKEN}"}

    for rome in rome_codes:
        matrix_step(f"Interrogation API /search ROME={rome}...")
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
            matrix_warn("Aucune donnée reçue")
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
                matrix_decode(row["Nom de l'entreprise"], prefix=f"    {GREEN}LBA ▸{RESET} ")

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
                if lbb_count <= 3:
                    matrix_decode(row["Nom de l'entreprise"], prefix=f"    {GREEN}LBB ▸{RESET} ")

        matrix_ok(f"Décodage : {lba_count} LBA + {lbb_count} LBB")
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
            existing = by_siret[siret]
            if existing["Source"] == "LBA":
                existing["Source"] = "LBA + LBB"
                existing["Score LBB"] = row.get("Score LBB") or existing.get("Score LBB", "")
                doublons += 1
            for col in OUTPUT_COLUMNS:
                if not str(existing.get(col, "")).strip() and str(row.get(col, "")).strip():
                    existing[col] = row[col]
        else:
            by_siret[siret] = row.copy()

    return list(by_siret.values()), doublons


# ── Export CSV ───────────────────────────────────────────────────────────────

def save_backup_csv(rows: list[dict], filepath: str):
    """Sauvegarde intermédiaire CSV."""
    os.makedirs(os.path.dirname(filepath) or ".", exist_ok=True)
    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    df.to_csv(filepath, index=False, encoding="utf-8-sig")


def export_csv(rows: list[dict], csv_path: str):
    """Exporte en CSV uniquement. SIRET en string."""
    os.makedirs(os.path.dirname(csv_path) or ".", exist_ok=True)
    df = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    df["SIRET"] = df["SIRET"].astype(str)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Prospection LBA + LBB — collecte entreprises par ville et codes NAF",
    )
    parser.add_argument("--ville", type=str, help="Nom de la ville")
    parser.add_argument("--naf", nargs="+", help="Codes NAF (ex: 10.71A 10.71B)")
    parser.add_argument("--rayon", type=int, default=30, help="Rayon en km (défaut: 30)")
    parser.add_argument("--output", type=str, default=".", help="Répertoire de sortie (défaut: racine repo)")
    parser.add_argument("--config", type=str, help="Fichier JSON de paramètres")
    parser.add_argument("--quiet", action="store_true",
                        help="Mode silencieux (désactive les animations)")
    parser.add_argument("--batch", type=str, default=None,
                        help="Chemin de batch forcé (ex: Prosp/SB/batch_2)")
    return parser.parse_args()


def _resolve_profiles(profile_names: list[str]) -> tuple[list[str], str]:
    """Résout un ou plusieurs noms de profils en liste de NAF dédupliquée."""
    combined: list[str] = []
    for name in profile_names:
        if name not in PROFILS_SKY:
            matrix_fail(f"Profil inconnu : « {name} »")
            matrix_step("Profils disponibles : " + ", ".join(sorted(PROFILS_SKY)))
            sys.exit(1)
        combined.extend(PROFILS_SKY[name])
    naf_codes = sorted(set(combined))
    label = ", ".join(profile_names)
    return naf_codes, label


def interactive_profile_menu() -> tuple[list[str], str]:
    """Menu interactif pour choisir un ou plusieurs profils Skill & You."""
    print()
    matrix_section("Sélection du profil sectoriel Skill & You")
    names = sorted(PROFILS_SKY.keys())
    for i, name in enumerate(names, 1):
        naf_list = PROFILS_SKY[name]
        print(f"    {GREEN}{i:>2}.{RESET} {name:<30s} ({len(naf_list)} codes NAF)")
    print()
    user_input = input(f"    Choix (numéro ou numéros séparés par virgules) > ").strip()
    if not user_input:
        matrix_fail("Aucun profil sélectionné")
        sys.exit(1)

    selected: list[str] = []
    for part in user_input.split(","):
        part = part.strip()
        try:
            idx = int(part) - 1
            if 0 <= idx < len(names):
                selected.append(names[idx])
            else:
                matrix_fail(f"Index hors limites : {part}")
                sys.exit(1)
        except ValueError:
            # Peut-être un nom de profil directement
            if part in PROFILS_SKY:
                selected.append(part)
            else:
                matrix_fail(f"Profil inconnu : « {part} »")
                sys.exit(1)

    return _resolve_profiles(selected)


def _resolve_batch_output(args, profile_names: list[str], batch_name: str | None = None,
                          naf_codes: list[str] | None = None) -> tuple[str, str | None, str | None]:
    """
    Détermine le dossier de sortie batch.
    Retourne (output_dir, secteur, batch_path).
    """
    from datetime import date as _date

    # --batch CLI force un chemin explicite
    if args.batch:
        os.makedirs(args.batch, exist_ok=True)
        return args.batch, None, args.batch

    # Pas de profil = pas de batch auto
    if not profile_names:
        os.makedirs(args.output, exist_ok=True)
        return args.output, None, None

    # Détecter le secteur
    secteur = detect_secteur(profile_names)
    if not secteur:
        known = list(SECTEUR_PREFIXES.values())
        print(f"\n    Profils multi-secteurs détectés. Dans quel secteur ranger ce batch ?")
        for i, s in enumerate(known, 1):
            print(f"      {GREEN}{i}.{RESET} {s}")
        choice = input("    Choix > ").strip()
        try:
            secteur = known[int(choice) - 1]
        except (ValueError, IndexError):
            secteur = choice.upper()

    # Chercher Prosp/ en remontant
    prosp_root = find_prosp_root()
    if prosp_root:
        secteur_dir = os.path.join(prosp_root, secteur)
    else:
        matrix_warn("Dossier Prosp/ introuvable — création dans le dossier courant.")
        secteur_dir = os.path.join(".", secteur)

    # Numéro de batch + nom personnalisable
    n = next_batch_number(secteur_dir)
    default_batch_name = f"batch_{n}"

    print(f"\n    {GREEN}▸{RESET} Secteur détecté : {BOLD}{secteur}{RESET}")
    print(f"    {GREEN}▸{RESET} Dossier parent  : {BOLD}{secteur_dir}/{RESET}")

    if batch_name:
        dir_name = batch_name
        matrix_ok(f"Nom de batch (config) : {dir_name}")
    else:
        raw = input(f"    {GREEN}▸{RESET} Nom du dossier de sortie (défaut: {default_batch_name}) : ").strip()
        dir_name = raw if raw else default_batch_name

    batch_dir = os.path.join(secteur_dir, dir_name)

    os.makedirs(batch_dir, exist_ok=True)

    # Écrire .current_batch (avec profils et NAF attendus)
    profils_str = ",".join(profile_names) if profile_names else ""
    naf_str = ",".join(naf_codes) if naf_codes else ""
    write_current_batch(secteur, batch_dir, _date.today().isoformat(),
                        profils=profils_str, naf_attendus=naf_str)
    matrix_ok(f".current_batch écrit — batch actif : {batch_dir}")

    return batch_dir, secteur, batch_dir


def interactive_menu(args) -> tuple[list[str], list[str], int, str, str | None, str | None]:
    """
    Mode interactif complet (aucun argument CLI).
    Retourne (villes, naf_codes, rayon, output_dir, profile_label, output_name).
    """
    print()
    matrix_section("Configuration interactive")

    # 1. Villes
    raw = input(f"    {GREEN}▸{RESET} Ville(s) (séparées par virgules) : ").strip()
    if not raw:
        matrix_fail("Aucune ville saisie")
        sys.exit(1)
    villes = [v.strip() for v in raw.split(",") if v.strip()]

    # 2. Rayon
    raw = input(f"    {GREEN}▸{RESET} Rayon (km) [10, 30, 60, 100] (défaut 30) : ").strip()
    rayon = 30
    if raw:
        try:
            rayon = int(raw)
        except ValueError:
            matrix_warn(f"Rayon invalide « {raw} » — défaut 30 km")

    # 3. Profils groupés par secteur
    print()
    matrix_section("Sélection du profil sectoriel Skill & You")

    sectors_order = ["SB", "BTPM", "MBT", "TG"]
    sector_groups: dict[str, list[str]] = {s: [] for s in sectors_order}
    for name in sorted(PROFILS_SKY.keys()):
        for prefix, label in SECTEUR_PREFIXES.items():
            if name.lower().startswith(f"{prefix}_"):
                sector_groups[label].append(name)
                break

    idx = 1
    idx_map: dict[int, str] = {}
    for sector in sectors_order:
        profiles = sector_groups[sector]
        if not profiles:
            continue
        print(f"\n    {BOLD}── {sector} ──{RESET}")
        for name in profiles:
            naf_list = PROFILS_SKY[name]
            print(f"    {GREEN}{idx:>2}.{RESET} {name:<30s} ({len(naf_list)} codes NAF)")
            idx_map[idx] = name
            idx += 1

    print()
    raw = input(f"    {GREEN}▸{RESET} Choix du/des profils (numéros séparés par virgules) : ").strip()
    if not raw:
        matrix_fail("Aucun profil sélectionné")
        sys.exit(1)

    selected: list[str] = []
    for part in raw.split(","):
        part = part.strip()
        try:
            i = int(part)
            if i in idx_map:
                selected.append(idx_map[i])
            else:
                matrix_fail(f"Numéro hors limites : {part}")
                sys.exit(1)
        except ValueError:
            if part in PROFILS_SKY:
                selected.append(part)
            else:
                matrix_fail(f"Profil inconnu : « {part} »")
                sys.exit(1)

    naf_codes, profile_label = _resolve_profiles(selected)
    matrix_ok(f"Profil(s) : {profile_label}")
    matrix_ok(f"NAF : {', '.join(naf_codes)}")

    # 4. Nom de fichier
    raw = input(f"\n    {GREEN}▸{RESET} Nom du fichier de sortie (sans extension) [défaut: auto] : ").strip()
    output_name = raw if raw else None

    # 5. Résolution batch
    output_dir, _, _ = _resolve_batch_output(args, selected, naf_codes=naf_codes)

    return villes, naf_codes, rayon, output_dir, profile_label, output_name


def load_config(args) -> tuple[list[str], list[str], int, str, str | None]:
    """Retourne (villes, naf_codes, rayon, output_dir, profile_label)."""
    profile_label = None
    profile_names: list[str] = []
    batch_name = None

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

        # Résolution NAF : profil > profils > codes_naf > menu interactif
        if "profil" in cfg:
            profile_names = [cfg["profil"]]
            naf_codes, profile_label = _resolve_profiles(profile_names)
        elif "profils" in cfg:
            profile_names = cfg["profils"]
            naf_codes, profile_label = _resolve_profiles(profile_names)
        elif "codes_naf" in cfg:
            naf_codes = cfg["codes_naf"]
        else:
            naf_codes, profile_label = interactive_profile_menu()

        rayon = cfg.get("rayon_km", 30)
        batch_name = cfg.get("batch_name")
    elif args.ville:
        villes = [args.ville]
        rayon = args.rayon
        if args.naf:
            naf_codes = args.naf
        else:
            naf_codes, profile_label = interactive_profile_menu()
    else:
        matrix_fail("Spécifiez --ville (+ --naf ou profil), ou --config")
        sys.exit(1)

    output_dir, _, _ = _resolve_batch_output(args, profile_names, batch_name=batch_name,
                                             naf_codes=naf_codes)
    return villes, naf_codes, rayon, output_dir, profile_label


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    if args.quiet:
        set_quiet(True)
    if not args.config and not args.ville:
        villes, naf_codes, rayon, output_dir, profile_label, output_name = interactive_menu(args)
    else:
        villes, naf_codes, rayon, output_dir, profile_label = load_config(args)
        output_name = None

    multi = len(villes) > 1

    # ── Bannière Matrix ──
    matrix_banner("PROSPECTION LBA + LBB")

    if profile_label:
        matrix_step(f"Profil chargé : {profile_label}")
        matrix_step(f"Codes NAF injectés : {', '.join(naf_codes)}")
    matrix_kv("Villes", ", ".join(villes))
    matrix_kv("Codes NAF", ", ".join(naf_codes))
    matrix_kv("Rayon", f"{rayon} km")

    # ── Nom de fichier ──
    today_str = date.today().strftime("%Y%m%d")
    if output_name:
        filename = output_name
    else:
        if multi:
            default_name = f"lba_lbb_results_multi_{today_str}"
        else:
            ville_slug = villes[0].lower().replace(" ", "_").replace("-", "_")
            default_name = f"lba_lbb_results_{ville_slug}_{today_str}"
        filename = ask_filename(default_name)

    # ── Vérification de cohérence NAF → ROME ──
    matrix_section("Vérification NAF → ROME")
    missing_naf = [naf for naf in naf_codes if naf not in NAF_TO_ROME]
    if missing_naf:
        for naf in missing_naf:
            matrix_warn(f"Code NAF {naf} absent de la table NAF_TO_ROME — sera ignoré.")
            print(f"      Ajoutez-le à la table pour activer ce mapping.")

    # Avertissements sémantiques : NAF présents mais dont le sens peut surprendre
    semantic_alerts = [naf for naf in naf_codes if naf in NAF_SEMANTIC_WARNINGS]
    if semantic_alerts:
        for naf in semantic_alerts:
            matrix_warn(NAF_SEMANTIC_WARNINGS[naf])
        confirm = input("    Continuer quand même ? (O/n) > ").strip().lower()
        if confirm == "n":
            matrix_fail("Abandonné par l'utilisateur.")
            sys.exit(0)

    # Résolution NAF → ROME
    rome_codes = resolve_rome_codes(naf_codes)
    if not rome_codes:
        matrix_fail("Aucun code ROME trouvé pour les NAF fournis")
        sys.exit(1)

    # Afficher le mapping effectif
    matrix_ok(f"Codes ROME ciblés : {', '.join(rome_codes)}")
    effective_naf = [naf for naf in naf_codes if naf in NAF_TO_ROME]
    if effective_naf:
        matrix_step("Mapping effectif :")
        for naf in effective_naf:
            print(f"      {naf} → {', '.join(NAF_TO_ROME[naf])}")

    # ── Boucle sur les villes ──
    all_city_results: dict[str, list[dict]] = {}
    city_stats: dict[str, dict] = {}

    for ville in villes:
        matrix_section(f"Connexion à la Matrice — {ville}")

        try:
            lat, lon = geocode(ville)

            effective_radius = rayon
            matrix_step(f"Collecte via /search (rayon {effective_radius} km)...")
            lba_rows, lbb_rows = fetch_search(rome_codes, lat, lon, effective_radius)

            # Rayon 0 : si aucun résultat, retry avec 1 km
            if effective_radius == 0 and not lba_rows and not lbb_rows:
                matrix_warn("Rayon 0 km : aucun résultat — retry avec 1 km")
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
            matrix_step("Déduplication par SIRET...")
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
            matrix_ok(f"{len(rows)} entreprises uniques pour {ville}")

            # Nettoyage backup
            if os.path.exists(backup_path):
                os.remove(backup_path)

        except Exception as e:
            matrix_fail(f"Erreur pour {ville} : {e}")
            all_city_results[ville] = []
            city_stats[ville] = {
                "n_lba": 0, "n_lbb": 0, "n_doublons": 0, "n_unique": 0,
            }
            continue

    # ── Vérification globale ──
    if not any(all_city_results.values()):
        matrix_warn("Aucune entreprise collectée pour aucune ville.")
        sys.exit(0)

    # ── Export CSV ──
    csv_path = os.path.join(output_dir, f"{filename}.csv")

    if multi:
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

        matrix_step("Export CSV consolidé...")
        export_csv(consolidated, csv_path)
        matrix_ok(f"Fichier : {csv_path}")

        # ── Synthèse ──
        matrix_section("RÉSULTATS — Données extraites de la Matrice")
        matrix_kv("NAF", f"{', '.join(naf_codes)} → ROME : {', '.join(rome_codes)}")
        print()
        for ville, st in city_stats.items():
            print(f"    {GREEN}▸{RESET} {ville}: {BOLD}{st['n_unique']}{RESET} uniques "
                  f"({st['n_lba']} LBA, {st['n_lbb']} LBB, {st['n_doublons']} doublons)")
        matrix_separator()
        n_phones = sum(1 for r in consolidated if r.get("Téléphone"))
        pct_phones = f"{n_phones / len(consolidated) * 100:.0f}%" if consolidated else "0%"
        matrix_kv("Total consolidé (dédup SIRET)", f"{len(consolidated)} entreprises")
        matrix_kv("Téléphones extraits", f"{n_phones} / {len(consolidated)} entreprises ({pct_phones})")
        matrix_kv("Fichier", os.path.basename(csv_path))
    else:
        ville = villes[0]
        rows = all_city_results.get(ville, [])
        if not rows:
            matrix_warn("Aucune entreprise collectée.")
            sys.exit(0)

        matrix_step("Export CSV...")
        export_csv(rows, csv_path)
        matrix_ok(f"Fichier : {csv_path}")

        st = city_stats[ville]
        matrix_section(f"RÉSULTATS — {ville} (rayon {rayon} km)")
        matrix_kv("Codes NAF → ROME", f"{', '.join(naf_codes)} → {', '.join(rome_codes)}")
        matrix_kv("LBA (offres actives)", f"{st['n_lba']} entreprises")
        matrix_kv("LBB (potentiel)", f"{st['n_lbb']} entreprises")
        matrix_kv("Doublons LBA+LBB", f"{st['n_doublons']} entreprises")
        matrix_separator()
        n_phones = sum(1 for r in rows if r.get("Téléphone"))
        pct_phones = f"{n_phones / len(rows) * 100:.0f}%" if rows else "0%"
        matrix_kv("Total unique (par SIRET)", f"{len(rows)} entreprises")
        matrix_kv("Téléphones extraits", f"{n_phones} / {len(rows)} entreprises ({pct_phones})")
        matrix_kv("Fichier", os.path.basename(csv_path))

    # ── Clôture Matrix ──
    morpheus_says()


if __name__ == "__main__":
    main()
