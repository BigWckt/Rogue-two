"""
Scraper d'entreprises en alternance/apprentissage
Sources : La Bonne Boite (LBB) + La Bonne Alternance (LBA)
Zone : 50km autour de Paris 1er (48.8603, 2.3477)
ROME : D1102 (Boulangerie-viennoiserie), D1104 (Pâtisserie, confiserie, chocolaterie, glacerie)
"""

import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import time
import sys

# ── Configuration ────────────────────────────────────────────────────────────

LATITUDE = 48.8603
LONGITUDE = 2.3477
DISTANCE_KM = 50
ROME_CODES = ["D1102", "D1104"]
MAX_TOTAL = 100

LBB_ENDPOINT = "https://labonneboite.pole-emploi.fr/api/v1/company/"
LBB_CLIENT_ID = "PAR_n8n_a2b06e9b91bb3285ed7afe43f67b260289b4fb31bcbaf45cd0ea863fcb81600e"
LBB_KEY = "ba1cbbc148015125f83c70b7425bdfaf0a28235d805d111b4dc21bbe5488f608"

LBA_ENDPOINT = "https://api.labonnealternance.apprentissage.beta.gouv.fr/api/V1/jobs"
LBA_TOKEN = (
    "eyJhbGciOiJIUzI1NiJ9."
    "eyJfaWQiOiI2OGM4MzJkZDgxZGY5MmFiYTc2MDNhNzAiLCJhcGlfa2V5IjoieVFaYkpiZElCN1VydDNvb2"
    "w3aTRqN0lSRUhSQ25KSk5ya0djclpxZ1E0bz0iLCJvcmdhbmlzYXRpb24iOiJTa2lsbGFuZHlvdSIsImVt"
    "YWlsIjoiam9mZnJleS5sYWRtaXJhdWx0QHNraWxsYW5keW91LmNvbSIsImlzcyI6ImFwaSIsImlhdCI6MTc"
    "3MzMyMzg1NSwiZXhwIjoxNzg5NDg2Njk2fQ."
    "5P4D4gmCTezPIvycQEtKnPZBjPvcx8BHG1bF8r_Z4ts"
)

OUTPUT_FILE = "entreprises_alternance_paris.xlsx"

EXCEL_COLUMNS = [
    "Source",
    "SIRET",
    "Raison sociale",
    "Adresse",
    "Code postal",
    "Ville",
    "Téléphone",
    "Email",
    "Site web",
    "Code NAF",
    "Code ROME",
    "Distance Paris 1er (km)",
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def get_safe(d, *keys, default=None):
    """Accès sécurisé dans un dict imbriqué."""
    for k in keys:
        if not isinstance(d, dict):
            return default
        d = d.get(k, default)
        if d is None:
            return default
    return d


def http_get(url, params=None, headers=None, timeout=20):
    """GET avec gestion des erreurs HTTP courantes."""
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=timeout)
        if resp.status_code == 429:
            print("  [WARN] 429 Too Many Requests — pause 10s puis réessai…")
            time.sleep(10)
            resp = requests.get(url, params=params, headers=headers, timeout=timeout)
        if resp.status_code == 401:
            print(f"  [ERROR] 401 Unauthorized — vérifier les credentials. URL: {url}")
            return None
        if resp.status_code == 403:
            print(f"  [ERROR] 403 Forbidden — accès refusé. URL: {url}")
            return None
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.Timeout:
        print(f"  [ERROR] Timeout après {timeout}s sur {url}")
        return None
    except requests.exceptions.ConnectionError as e:
        print(f"  [ERROR] Connexion impossible : {e}")
        return None
    except requests.exceptions.HTTPError as e:
        print(f"  [ERROR] HTTP {e.response.status_code} sur {url}")
        return None
    except ValueError:
        print(f"  [ERROR] Réponse non-JSON sur {url}")
        return None


# ── La Bonne Boite ────────────────────────────────────────────────────────────

def fetch_lbb(rome_code: str) -> list[dict]:
    """
    Interroge l'API LBB pour un code ROME.
    Teste d'abord avec le préfixe lettre (D1102), puis sans (1102) en cas d'échec.
    """
    base_params = {
        "latitude": LATITUDE,
        "longitude": LONGITUDE,
        "distance": DISTANCE_KM,
        "rome_codes": rome_code,
        "client_id": LBB_CLIENT_ID,
        "key": LBB_KEY,
    }

    print(f"\n[LBB] Requête ROME={rome_code} …")
    data = http_get(LBB_ENDPOINT, params=base_params)

    # Si échec ou résultats vides, réessayer sans préfixe lettre
    if data is None or (isinstance(data, dict) and data.get("companies") is None):
        rome_numeric = rome_code.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
        print(f"  [LBB] Réessai avec ROME numérique={rome_numeric}")
        base_params["rome_codes"] = rome_numeric
        data = http_get(LBB_ENDPOINT, params=base_params)

    if data is None:
        print(f"  [LBB] Aucune donnée reçue pour ROME={rome_code}")
        return []

    companies = data.get("companies") or data.get("results") or []
    if not isinstance(companies, list):
        # Parfois la réponse est directement une liste
        companies = data if isinstance(data, list) else []

    print(f"  [LBB] {len(companies)} entreprise(s) brute(s) pour ROME={rome_code}")

    results = []
    for c in companies:
        siret = str(get_safe(c, "siret") or "").strip()
        if not siret:
            continue
        results.append({
            "Source": "LBB",
            "SIRET": siret,
            "Raison sociale": get_safe(c, "name") or get_safe(c, "raison_sociale") or "",
            "Adresse": get_safe(c, "address") or get_safe(c, "street") or "",
            "Code postal": str(get_safe(c, "zipCode") or get_safe(c, "zip_code") or ""),
            "Ville": get_safe(c, "city") or get_safe(c, "commune") or "",
            "Téléphone": get_safe(c, "phone") or "",
            "Email": get_safe(c, "email") or "",
            "Site web": get_safe(c, "website") or get_safe(c, "url") or "",
            "Code NAF": get_safe(c, "naf") or get_safe(c, "code_naf") or "",
            "Code ROME": rome_code,
            "Distance Paris 1er (km)": get_safe(c, "distance") or "",
        })
    return results


# ── La Bonne Alternance ───────────────────────────────────────────────────────

def fetch_lba(rome_code: str) -> list[dict]:
    """
    Interroge l'API LBA pour un code ROME.
    Extrait les employeurs des offres et dédoublonne par SIRET.
    """
    params = {
        "romes": rome_code,
        "latitude": LATITUDE,
        "longitude": LONGITUDE,
        "radius": DISTANCE_KM,
        "caller": "Skillandyou_scraper",
    }
    headers = {"Authorization": f"Bearer {LBA_TOKEN}"}

    print(f"\n[LBA] Requête ROME={rome_code} …")
    data = http_get(LBA_ENDPOINT, params=params, headers=headers)

    if data is None:
        print(f"  [LBA] Aucune donnée reçue pour ROME={rome_code}")
        return []

    # L'API LBA retourne plusieurs collections d'offres
    # Clés possibles : lbaCompanies, peJobs, lbbCompanies, matchas, …
    raw_offers = []
    for key in ("lbaCompanies", "lbbCompanies", "peJobs", "matchas", "jobs", "results"):
        block = data.get(key)
        if isinstance(block, list):
            raw_offers.extend(block)
        elif isinstance(block, dict):
            inner = block.get("results") or block.get("jobs") or []
            if isinstance(inner, list):
                raw_offers.extend(inner)

    print(f"  [LBA] {len(raw_offers)} offre(s) brute(s) pour ROME={rome_code}")

    seen: dict[str, dict] = {}
    for offer in raw_offers:
        # L'employeur peut être à la racine ou dans un sous-objet
        employer = offer.get("company") or offer.get("employer") or offer

        siret = (
            str(get_safe(employer, "siret") or get_safe(offer, "siret") or "").strip()
        )
        if not siret:
            continue
        if siret in seen:
            continue

        # Adresse : peut être un objet ou une chaîne
        address_obj = employer.get("address") or offer.get("address") or {}
        if isinstance(address_obj, str):
            adresse = address_obj
            code_postal = ""
            ville = ""
        else:
            adresse = address_obj.get("street") or address_obj.get("label") or ""
            code_postal = str(address_obj.get("zipCode") or address_obj.get("zip_code") or "")
            ville = address_obj.get("city") or address_obj.get("commune") or ""

        # Fallback champs plats
        if not code_postal:
            code_postal = str(
                get_safe(employer, "zipCode") or get_safe(employer, "zip_code")
                or get_safe(offer, "zipCode") or ""
            )
        if not ville:
            ville = get_safe(employer, "city") or get_safe(offer, "city") or ""

        distance = (
            get_safe(offer, "distance")
            or get_safe(employer, "distance")
            or ""
        )

        seen[siret] = {
            "Source": "LBA",
            "SIRET": siret,
            "Raison sociale": (
                get_safe(employer, "name") or get_safe(offer, "company", "name") or ""
            ),
            "Adresse": adresse,
            "Code postal": code_postal,
            "Ville": ville,
            "Téléphone": get_safe(employer, "phone") or get_safe(offer, "phone") or "",
            "Email": get_safe(employer, "email") or get_safe(offer, "email") or "",
            "Site web": (
                get_safe(employer, "website") or get_safe(employer, "url")
                or get_safe(offer, "url") or ""
            ),
            "Code NAF": get_safe(employer, "naf") or get_safe(offer, "naf") or "",
            "Code ROME": rome_code,
            "Distance Paris 1er (km)": distance,
        }

    print(f"  [LBA] {len(seen)} entreprise(s) unique(s) après dédup SIRET pour ROME={rome_code}")
    return list(seen.values())


# ── Déduplication et fusion des sources ──────────────────────────────────────

def merge_and_deduplicate(lbb_rows: list[dict], lba_rows: list[dict]) -> list[dict]:
    """
    Fusionne LBB et LBA.
    - Un SIRET présent dans les deux sources → Source = "LBB + LBA"
    - Prend les champs LBB en priorité (complète avec LBA si champ vide)
    """
    by_siret: dict[str, dict] = {}

    for row in lbb_rows:
        s = row["SIRET"]
        if s not in by_siret:
            by_siret[s] = row.copy()
        else:
            # Compléter les champs vides
            for col in EXCEL_COLUMNS:
                if not by_siret[s].get(col) and row.get(col):
                    by_siret[s][col] = row[col]

    for row in lba_rows:
        s = row["SIRET"]
        if s not in by_siret:
            by_siret[s] = row.copy()
        else:
            by_siret[s]["Source"] = "LBB + LBA"
            # Compléter les champs vides avec les données LBA
            for col in EXCEL_COLUMNS:
                if not by_siret[s].get(col) and row.get(col):
                    by_siret[s][col] = row[col]

    return list(by_siret.values())


# ── Export Excel ──────────────────────────────────────────────────────────────

def export_excel(rows: list[dict], filepath: str) -> None:
    """Génère le fichier Excel formaté."""

    df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)
    df = df.head(MAX_TOTAL)

    # Statistiques
    total = len(df)
    nb_lbb = (df["Source"] == "LBB").sum()
    nb_lba = (df["Source"] == "LBA").sum()
    nb_both = (df["Source"] == "LBB + LBA").sum()
    nb_no_phone = df["Téléphone"].apply(lambda x: not str(x).strip()).sum()

    print(f"\n── Synthèse ──────────────────────────────────────")
    print(f"  Total entreprises exportées : {total}")
    print(f"  LBB uniquement             : {nb_lbb}")
    print(f"  LBA uniquement             : {nb_lba}")
    print(f"  LBB + LBA                  : {nb_both}")
    print(f"  Lignes sans téléphone      : {nb_no_phone}")
    print(f"──────────────────────────────────────────────────\n")

    # Écriture de base avec pandas
    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Résultats_Scrap", index=False)

    # Formatage avec openpyxl
    wb = load_workbook(filepath)
    ws = wb["Résultats_Scrap"]

    header_fill = PatternFill("solid", fgColor="0066FF")
    alt_fill = PatternFill("solid", fgColor="E8F0FF")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    summary_fill = PatternFill("solid", fgColor="002080")

    header_font = Font(bold=True, color="FFFFFF")
    summary_font = Font(bold=True, color="FFFFFF")

    # En-têtes
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Lignes alternées
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
        fill = white_fill if row_idx % 2 == 0 else alt_fill
        for cell in row:
            cell.fill = fill
            cell.alignment = Alignment(vertical="center")

    # Ligne de synthèse
    summary_row = ws.max_row + 2
    summary_data = [
        ("Total entreprises", total),
        ("dont LBB", nb_lbb),
        ("dont LBA", nb_lba),
        ("dont LBB+LBA", nb_both),
        ("Sans téléphone", nb_no_phone),
    ]
    for i, (label, value) in enumerate(summary_data):
        label_cell = ws.cell(row=summary_row + i, column=1, value=label)
        value_cell = ws.cell(row=summary_row + i, column=2, value=value)
        for cell in (label_cell, value_cell):
            cell.fill = summary_fill
            cell.font = summary_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Auto-dimensionnement des colonnes
    for col_idx, col in enumerate(ws.columns, start=1):
        max_len = 0
        col_letter = get_column_letter(col_idx)
        for cell in col:
            try:
                cell_len = len(str(cell.value)) if cell.value is not None else 0
                if cell_len > max_len:
                    max_len = cell_len
            except Exception:
                pass
        adjusted_width = min(max_len + 4, 50)
        ws.column_dimensions[col_letter].width = adjusted_width

    # Figer la première ligne
    ws.freeze_panes = "A2"

    wb.save(filepath)
    print(f"[OK] Fichier Excel exporté : {filepath}")


# ── Point d'entrée ────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  SCRAPER ALTERNANCE — Paris 1er, 50 km")
    print("  ROME : D1102 (Boulangerie) + D1104 (Pâtisserie)")
    print("=" * 60)

    all_lbb: list[dict] = []
    all_lba: list[dict] = []

    # ── LBB ──
    for rome in ROME_CODES:
        rows = fetch_lbb(rome)
        all_lbb.extend(rows)

    # ── LBA ──
    for rome in ROME_CODES:
        rows = fetch_lba(rome)
        all_lba.extend(rows)

    print(f"\n[INFO] LBB total brut (toutes ROME) : {len(all_lbb)}")
    print(f"[INFO] LBA total brut (toutes ROME) : {len(all_lba)}")

    # ── Fusion & déduplication ──
    merged = merge_and_deduplicate(all_lbb, all_lba)
    print(f"[INFO] Entreprises uniques après dédup SIRET : {len(merged)}")

    if not merged:
        print("[WARN] Aucune entreprise trouvée. Vérifier les credentials et la disponibilité des APIs.")
        sys.exit(0)

    # ── Vérification des lignes sans téléphone ──
    no_phone_count = sum(1 for r in merged if not str(r.get("Téléphone", "")).strip())
    if no_phone_count:
        print(f"[INFO] {no_phone_count} entreprise(s) sans numéro de téléphone (lignes conservées)")

    # ── Export ──
    export_excel(merged, OUTPUT_FILE)


if __name__ == "__main__":
    main()
