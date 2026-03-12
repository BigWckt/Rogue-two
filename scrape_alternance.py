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

LBB_ENDPOINT = "https://api.francetravail.io/partenaire/labonneboite/v2/company/"
LBB_OAUTH_URL = "https://entreprise.francetravail.fr/connexion/oauth2/access_token?realm=/partenaire"
LBB_CLIENT_ID = "PAR_claude_5c1b4a8fdb5c52468da374fa0c6be6a50d5dd153527a93c20dffd430edb18fef"
LBB_CLIENT_SECRET = "2d987a879cc64773446a3bad979782be29c5eda7380497a47f1de4133a6a4f93"

LBA_ENDPOINT = "https://api.apprentissage.beta.gouv.fr/api/job/v1/search"
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
            print(f"  [ERROR] 403 Forbidden (insufficient_scope) — URL: {url}")
            print("  [HINT]  L'application n'a pas les droits sur cette API LBB.")
            print("          → Aller sur https://francetravail.io → Mon espace → Mes applications")
            print("          → Vérifier la souscription à 'La Bonne Boite' et activer le scope.")
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

_lbb_token_cache: dict = {}


def _get_lbb_token() -> str | None:
    """
    Obtient un Bearer token OAuth2 pour l'API LBB via client_credentials.
    Met en cache le token jusqu'à expiration.
    """
    import time as _time
    now = _time.time()
    if _lbb_token_cache.get("token") and _lbb_token_cache.get("expires_at", 0) > now + 30:
        return _lbb_token_cache["token"]

    print("  [LBB] Obtention du token OAuth2 …")
    try:
        # Scope requis pour l'API La Bonne Boite v2 sur francetravail.io
        scope = f"api_labonneboitev2 application_{LBB_CLIENT_ID}"
        resp = requests.post(
            LBB_OAUTH_URL,
            data={
                "grant_type": "client_credentials",
                "client_id": LBB_CLIENT_ID,
                "client_secret": LBB_CLIENT_SECRET,
                "scope": scope,
            },
            timeout=15,
        )
        if resp.status_code == 401:
            print("  [LBB][ERROR] 401 Unauthorized — credentials invalides pour OAuth2")
            return None
        if resp.status_code == 400:
            err = resp.json().get("error_description", "")
            print(f"  [LBB][ERROR] 400 Bad Request token — {err}")
            print("  [HINT]  Vérifier que l'application est bien souscrite à l'API LBB")
            print("          sur https://francetravail.io et que le scope est autorisé.")
            # Fallback : tenter avec seulement le scope application
            fallback_scope = f"application_{LBB_CLIENT_ID}"
            resp2 = requests.post(LBB_OAUTH_URL, data={
                "grant_type": "client_credentials",
                "client_id": LBB_CLIENT_ID,
                "client_secret": LBB_CLIENT_SECRET,
                "scope": fallback_scope,
            }, timeout=15)
            if resp2.status_code != 200:
                return None
            resp = resp2
        resp.raise_for_status()
        token_data = resp.json()
        token = token_data.get("access_token")
        expires_in = int(token_data.get("expires_in", 1800))
        _lbb_token_cache["token"] = token
        _lbb_token_cache["expires_at"] = now + expires_in
        print(f"  [LBB] Token obtenu (expire dans {expires_in}s)")
        return token
    except requests.exceptions.HTTPError as e:
        print(f"  [LBB][ERROR] HTTP {e.response.status_code} lors de l'obtention du token")
        return None
    except Exception as e:
        print(f"  [LBB][ERROR] Impossible d'obtenir le token OAuth2 : {e}")
        return None


def fetch_lbb(rome_code: str) -> list[dict]:
    """
    Interroge l'API LBB (francetravail.io) pour un code ROME.
    Auth : Bearer token OAuth2.
    Teste d'abord avec le préfixe lettre (D1102), puis sans (1102) en cas d'échec.
    """
    token = _get_lbb_token()
    if not token:
        print(f"  [LBB] Token indisponible — skip ROME={rome_code}")
        return []

    headers = {"Authorization": f"Bearer {token}"}
    base_params = {
        "latitude": LATITUDE,
        "longitude": LONGITUDE,
        "distance": DISTANCE_KM,
        "rome_codes": rome_code,
    }

    print(f"\n[LBB] Requête ROME={rome_code} …")
    data = http_get(LBB_ENDPOINT, params=base_params, headers=headers)

    # Si échec ou résultats vides, réessayer sans préfixe lettre
    if data is None or (isinstance(data, dict) and data.get("companies") is None):
        rome_numeric = rome_code.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz")
        print(f"  [LBB] Réessai avec ROME numérique={rome_numeric}")
        base_params["rome_codes"] = rome_numeric
        data = http_get(LBB_ENDPOINT, params=base_params, headers=headers)

    if data is None:
        print(f"  [LBB] Aucune donnée reçue pour ROME={rome_code}")
        return []

    companies = data.get("companies") or data.get("results") or []
    if not isinstance(companies, list):
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

import re as _re

def _parse_address(full_address: str) -> tuple[str, str, str]:
    """
    Parse une adresse complète du type '2 RUE CATULLE MENDES 75017 PARIS'
    en (rue, code_postal, ville).
    """
    if not full_address:
        return "", "", ""
    # Cherche un code postal 5 chiffres
    m = _re.search(r'(\d{5})\s+(.+)$', full_address.strip())
    if m:
        code_postal = m.group(1)
        ville = m.group(2).strip()
        rue = full_address[:m.start()].strip().rstrip(',').strip()
        return rue, code_postal, ville
    return full_address, "", ""


def _extract_company(item: dict, rome_code: str) -> dict | None:
    """
    Extrait les données entreprise d'un objet job ou recruiter.
    Retourne None si pas de SIRET.
    """
    wp = item.get("workplace") or {}
    apply = item.get("apply") or {}
    domain = wp.get("domain") or {}
    naf = domain.get("naf") or {}

    siret = str(wp.get("siret") or "").strip()
    if not siret:
        return None

    full_address = (wp.get("location") or {}).get("address") or ""
    rue, code_postal, ville = _parse_address(full_address)

    # Code ROME : depuis l'offre si dispo, sinon depuis le paramètre de recherche
    offer = item.get("offer") or {}
    rome = (offer.get("rome_codes") or [rome_code])[0]

    return {
        "Source": "LBA",
        "SIRET": siret,
        "Raison sociale": wp.get("name") or wp.get("legal_name") or wp.get("brand") or "",
        "Adresse": rue,
        "Code postal": code_postal,
        "Ville": ville,
        "Téléphone": apply.get("phone") or "",
        "Email": "",                          # champ absent du schéma API v1
        "Site web": wp.get("website") or "",
        "Code NAF": naf.get("code") or "",
        "Code ROME": rome,
        "Distance Paris 1er (km)": "",        # non fourni par cette API
    }


def fetch_lba(rome_code: str) -> list[dict]:
    """
    Interroge l'API LBA (nouvelle version) pour un code ROME.
    Combine jobs[] (offres) et recruiters[] (entreprises susceptibles de recruter).
    Dédoublonne par SIRET.
    """
    params = {
        "romes": rome_code,
        "latitude": LATITUDE,
        "longitude": LONGITUDE,
        "radius": DISTANCE_KM,
    }
    headers = {"Authorization": f"Bearer {LBA_TOKEN}"}

    print(f"\n[LBA] Requête ROME={rome_code} …")
    data = http_get(LBA_ENDPOINT, params=params, headers=headers)

    if data is None:
        print(f"  [LBA] Aucune donnée reçue pour ROME={rome_code}")
        return []

    jobs = data.get("jobs") or []
    recruiters = data.get("recruiters") or []
    print(f"  [LBA] {len(jobs)} offre(s) + {len(recruiters)} recruteur(s) bruts pour ROME={rome_code}")

    seen: dict[str, dict] = {}
    for item in list(jobs) + list(recruiters):
        company = _extract_company(item, rome_code)
        if company is None:
            continue
        siret = company["SIRET"]
        if siret not in seen:
            seen[siret] = company

    print(f"  [LBA] {len(seen)} entreprise(s) unique(s) après dédup SIRET pour ROME={rome_code}")
    return list(seen.values())


# ── Déduplication et fusion des sources ──────────────────────────────────────

def merge_and_deduplicate(lbb_rows: list[dict], lba_rows: list[dict]) -> list[dict]:
    """
    Fusionne LBB et LBA.
    - Un SIRET présent dans les deux sources → Source = "LBB + LBA"
    - Prend les champs LBB en priorité (complète avec LBA si champ vide)
    - Les doublons intra-LBB ou intra-LBA (même SIRET sur plusieurs ROME) sont
      consolidés sans changer la source.
    """
    # ── Dédup interne LBB (même SIRET sur D1102 et D1104) ──
    lbb_by_siret: dict[str, dict] = {}
    for row in lbb_rows:
        s = row["SIRET"]
        if s not in lbb_by_siret:
            lbb_by_siret[s] = row.copy()
        else:
            for col in EXCEL_COLUMNS:
                if not lbb_by_siret[s].get(col) and row.get(col):
                    lbb_by_siret[s][col] = row[col]

    # ── Dédup interne LBA (même SIRET sur D1102 et D1104) ──
    lba_by_siret: dict[str, dict] = {}
    for row in lba_rows:
        s = row["SIRET"]
        if s not in lba_by_siret:
            lba_by_siret[s] = row.copy()
        else:
            for col in EXCEL_COLUMNS:
                if not lba_by_siret[s].get(col) and row.get(col):
                    lba_by_siret[s][col] = row[col]

    # ── Fusion inter-sources ──
    by_siret: dict[str, dict] = dict(lbb_by_siret)

    for s, row in lba_by_siret.items():
        if s not in by_siret:
            by_siret[s] = row.copy()
        else:
            # SIRET présent dans LBB et LBA → fusion
            by_siret[s]["Source"] = "LBB + LBA"
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
