#!/usr/bin/env python3
"""
script_enrichissement.py — Enrichissement SIRET via API recherche-entreprises
=============================================================================
Pour chaque entreprise issue de Pages Jaunes (sans SIRET), interroge l'API
recherche-entreprises.api.gouv.fr pour retrouver le SIRET à partir du nom
et du code postal / ville.

Usage :
  python script_enrichissement.py --input pj_results_paris_20260402.xlsx
  python script_enrichissement.py --input pj_results_paris_20260402.xlsx --resume
"""

import argparse
import csv
import os
import re
import sys
import time
from datetime import date
from urllib.parse import quote

import pandas as pd
import requests
from rapidfuzz import fuzz

# ── Config ────────────────────────────────────────────────────────────────────

API_URL = "https://recherche-entreprises.api.gouv.fr/search"
API_DELAY = 1.5
API_TIMEOUT = 15
MAX_RETRIES = 3
SIMILARITY_THRESHOLD = 80
CHECKPOINT_EVERY = 25

# ── Colonnes de sortie (cohérentes avec script_lba_lbb.py) ───────────────────

ENRICHMENT_COLUMNS = [
    "Nom de l'entreprise",
    "Adresse",
    "Ville",
    "Code Postal",
    "Téléphone",
    "Site web",
    "Catégorie",
    "URL fiche PJ",
    "Date de collecte",
    "SIRET",
    "Code NAF",
    "Statut enrichissement",
    "Score similarité",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def normalize_siret(raw) -> str:
    if raw is None or str(raw).strip() in ("", "nan", "None"):
        return ""
    try:
        return str(int(float(str(raw).strip()))).zfill(14)
    except (ValueError, OverflowError):
        return str(raw).strip()


USER_AGENT = "SkillAndYou-Prospection/1.0 (enrichissement SIRET)"


def http_get(url, params=None):
    """GET avec retry + backoff exponentiel + respect Retry-After."""
    headers = {"User-Agent": USER_AGENT}
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(
                url, params=params, headers=headers, timeout=API_TIMEOUT,
            )
            if resp.status_code == 429:
                wait = int(resp.headers.get("Retry-After", 2 ** attempt))
                print(f"  ⚠️  HTTP 429 — retry {attempt}/{MAX_RETRIES} dans {wait}s")
                time.sleep(wait)
                continue
            if resp.status_code in (500, 502, 503, 504):
                wait = 2 ** attempt
                print(f"  ⚠️  HTTP {resp.status_code} — retry {attempt}/{MAX_RETRIES} dans {wait}s")
                time.sleep(wait)
                continue
            if resp.status_code == 400:
                return {"results": []}  # Bad request — not retryable
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.Timeout:
            wait = 2 ** attempt
            print(f"  ⚠️  Timeout — retry {attempt}/{MAX_RETRIES} dans {wait}s")
            time.sleep(wait)
        except requests.exceptions.ConnectionError as e:
            wait = 2 ** attempt
            print(f"  ⚠️  Connexion : {e} — retry {attempt}/{MAX_RETRIES}")
            time.sleep(wait)
        except requests.exceptions.HTTPError as e:
            print(f"  ❌ HTTP {e.response.status_code}")
            return None
        except ValueError:
            print(f"  ❌ Réponse non-JSON")
            return None
    return None


def clean_name(name: str) -> str:
    """Normalise un nom d'entreprise pour la comparaison."""
    name = name.upper().strip()
    # Retirer les formes juridiques courantes
    for suffix in ["SARL", "SAS", "SA", "EURL", "SCI", "SASU", "EI"]:
        name = re.sub(rf"\b{suffix}\b", "", name)
    return re.sub(r"\s+", " ", name).strip()


# ── Recherche SIRET ──────────────────────────────────────────────────────────

def _best_match(candidates: list[dict], nom_original: str) -> dict | None:
    """Trouve le meilleur match par similarité de nom parmi les candidats API."""
    nom_clean = clean_name(nom_original)
    best_score = 0
    best_result = None

    for r in candidates:
        api_names = []
        for field in ["nom_complet", "nom_raison_sociale"]:
            v = r.get(field) or ""
            if v:
                api_names.append(v)
        siege = r.get("siege") or {}
        enseignes = siege.get("liste_enseignes") or []
        if enseignes:
            api_names.append(enseignes[0])

        for api_name in api_names:
            score = fuzz.ratio(nom_clean, clean_name(api_name))
            if score > best_score:
                best_score = score
                best_result = r

    if best_result is None:
        return None

    score_rounded = round(best_score)
    if best_score >= SIMILARITY_THRESHOLD:
        siege = best_result.get("siege") or {}
        return {
            "siret": normalize_siret(siege.get("siret") or ""),
            "naf": siege.get("activite_principale") or "",
            "score": score_rounded,
            "status": "SIRET trouvé",
        }
    return {
        "siret": "",
        "naf": "",
        "score": score_rounded,
        "status": f"Exclu — similarité {score_rounded}% < {SIMILARITY_THRESHOLD}%",
    }


def _filter_by_cp(results: list[dict], code_postal: str) -> list[dict]:
    """Filtre les résultats dont le code postal du siège correspond."""
    if not code_postal:
        return []
    return [
        r for r in results
        if (r.get("siege") or {}).get("code_postal", "") == code_postal
    ]


def _filter_by_commune(results: list[dict], ville: str) -> list[dict]:
    """Filtre les résultats dont la commune du siège correspond (insensible casse)."""
    if not ville:
        return []
    ville_lower = ville.lower().strip()
    return [
        r for r in results
        if ville_lower in ((r.get("siege") or {}).get("libelle_commune") or "").lower()
    ]


def search_siret(nom: str, code_postal: str, ville: str) -> dict:
    """
    1 seul appel API (q=NOM, per_page=10, minimal=true), puis filtrage local :
      - Étape 1 : filtrer par code_postal exact
      - Étape 2 (si vide) : filtrer par commune
      - Étape 3 (si vide) : garder tous les résultats
    Fallback : si 0 résultat, 2ᵉ appel avec nom nettoyé. Max 2 appels.
    """
    nom_clean = clean_name(nom)
    if not nom_clean:
        return {"siret": "", "naf": "", "score": 0, "status": "Exclu — nom vide"}

    # ── Appel 1 : nom brut ────────────────────────────────────────────────
    data = http_get(API_URL, params={
        "q": nom, "per_page": "10", "minimal": "false",
    })
    if data is None:
        return {"siret": "", "naf": "", "score": 0, "status": "Exclu — erreur API"}

    results = data.get("results") or []

    # ── Appel 2 (fallback) : nom nettoyé si 0 résultat ───────────────────
    if not results and nom != nom_clean:
        time.sleep(API_DELAY)
        data = http_get(API_URL, params={
            "q": nom_clean, "per_page": "10", "minimal": "false",
        })
        if data is None:
            return {"siret": "", "naf": "", "score": 0, "status": "Exclu — erreur API"}
        results = data.get("results") or []

    if not results:
        return {"siret": "", "naf": "", "score": 0, "status": "Exclu — SIRET introuvable"}

    # ── Filtrage local en cascade ─────────────────────────────────────────
    subset = _filter_by_cp(results, code_postal)
    if not subset:
        subset = _filter_by_commune(results, ville)
    if not subset:
        subset = results

    # ── Meilleur match par similarité ─────────────────────────────────────
    match = _best_match(subset, nom)
    if match:
        return match

    return {"siret": "", "naf": "", "score": 0, "status": "Exclu — SIRET introuvable"}


# ── Checkpoint / Resume ──────────────────────────────────────────────────────

def save_checkpoint(rows: list[dict], filepath: str):
    """Sauvegarde checkpoint CSV."""
    try:
        df = pd.DataFrame(rows, columns=ENRICHMENT_COLUMNS)
        df.to_csv(filepath, index=False, encoding="utf-8-sig")
    except Exception as e:
        print(f"  ⚠️  Erreur checkpoint : {e}")


def load_checkpoint(filepath: str) -> list[dict]:
    """Charge un checkpoint CSV existant."""
    if not os.path.exists(filepath):
        return []
    try:
        df = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str)
        df = df.fillna("")
        return df.to_dict("records")
    except Exception:
        return []


# ── Export ────────────────────────────────────────────────────────────────────

def export_results(rows: list[dict], output_path: str):
    """Exporte Excel avec onglet principal + onglet Exclus."""
    found = [r for r in rows if r.get("Statut enrichissement") == "SIRET trouvé"]
    excluded = [r for r in rows if r.get("Statut enrichissement", "").startswith("Exclu")]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_found = pd.DataFrame(found, columns=ENRICHMENT_COLUMNS)
        df_found["SIRET"] = df_found["SIRET"].astype(str)
        df_found.to_excel(writer, sheet_name="Entreprises", index=False)

        if excluded:
            df_excl = pd.DataFrame(excluded, columns=ENRICHMENT_COLUMNS)
            df_excl.to_excel(writer, sheet_name="Exclus", index=False)


# ── Synthèse console ─────────────────────────────────────────────────────────

def print_synthese(input_file: str, rows: list[dict], output_path: str,
                   n_errors: int):
    n_total = len(rows)
    n_found = sum(1 for r in rows if r.get("Statut enrichissement") == "SIRET trouvé")
    n_excluded = sum(1 for r in rows if r.get("Statut enrichissement", "").startswith("Exclu"))
    pct = f"{n_found / n_total * 100:.0f}%" if n_total else "0%"

    # Extraire ville du nom de fichier
    ville = os.path.basename(input_file).replace("pj_results_", "").split("_")[0].title()

    print()
    print("══════════════════════════════════════════════════")
    print(f"  ENRICHISSEMENT SIRET — {ville}")
    print(f"  Fichier source : {os.path.basename(input_file)}")
    print()
    print(f"  Entreprises traitées     : {n_total}")
    print(f"  SIRET trouvés            : {n_found} ({pct})")
    print(f"  Exclus (< {SIMILARITY_THRESHOLD}% similarité): {n_excluded}")
    print(f"  Erreurs API              : {n_errors}")
    print("  ──────────────────────────────────────────────")
    print(f"  Fichier enrichi : {os.path.basename(output_path)}")
    print(f"  Onglet principal : {n_found} entreprises avec SIRET")
    print(f"  Onglet Exclus : {n_excluded} entreprises")
    print("══════════════════════════════════════════════════")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Enrichissement SIRET — recherche via API recherche-entreprises",
    )
    parser.add_argument("--input", required=True, help="Fichier Excel PJ source")
    parser.add_argument("--output", type=str, default=None,
                        help="Répertoire de sortie (défaut: même dossier)")
    parser.add_argument("--resume", action="store_true",
                        help="Reprendre depuis le dernier checkpoint")
    return parser.parse_args()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    input_file = args.input

    if not os.path.exists(input_file):
        print(f"❌ Fichier introuvable : {input_file}")
        sys.exit(1)

    # Déterminer le répertoire de sortie
    output_dir = args.output or os.path.dirname(input_file) or "."
    os.makedirs(output_dir, exist_ok=True)

    # Noms de fichiers de sortie
    base = os.path.splitext(os.path.basename(input_file))[0]
    output_path = os.path.join(output_dir, f"{base}_enrichi.xlsx")
    checkpoint_path = os.path.join(output_dir, f"{base}_checkpoint.csv")
    csv_backup_path = os.path.join(output_dir, f"{base}_enrichi.csv")

    # Charger le fichier source (onglet Consolidé si multi-villes, sinon sheet 0)
    print(f"📂 Chargement de {input_file}…")
    try:
        df = pd.read_excel(input_file, sheet_name="Consolidé", dtype=str)
        print("  (onglet Consolidé détecté)")
    except (ValueError, KeyError):
        df = pd.read_excel(input_file, dtype=str)
    df = df.fillna("")
    fiches = df.to_dict("records")
    print(f"  {len(fiches)} entreprises chargées")

    # Resume : charger le checkpoint
    already_done: dict[str, dict] = {}
    if args.resume:
        checkpoint_rows = load_checkpoint(checkpoint_path)
        if checkpoint_rows:
            for r in checkpoint_rows:
                key = (r.get("Nom de l'entreprise", "").lower(), r.get("Code Postal", ""))
                already_done[key] = r
            print(f"  ♻️  Reprise : {len(already_done)} entreprises déjà traitées")

    print()
    print("══════════════════════════════════════════════════")
    print("  ENRICHISSEMENT SIRET")
    print(f"  Source : {os.path.basename(input_file)}")
    print(f"  Entreprises : {len(fiches)}")
    print("══════════════════════════════════════════════════")
    print()

    # Enrichissement
    enriched_rows: list[dict] = []
    n_errors = 0

    for i, fiche in enumerate(fiches, 1):
        nom = fiche.get("Nom de l'entreprise", "")
        cp = fiche.get("Code Postal", "")
        ville = fiche.get("Ville", "")

        # Check si déjà fait (resume)
        key = (nom.lower(), cp)
        if key in already_done:
            enriched_rows.append(already_done[key])
            continue

        print(f"  [{i}/{len(fiches)}] 🔍 {nom[:50]}…", end=" ", flush=True)

        time.sleep(API_DELAY)
        result = search_siret(nom, cp, ville)

        # Construire la ligne enrichie
        row = {**fiche}
        row["SIRET"] = result["siret"]
        row["Code NAF"] = result["naf"]
        row["Statut enrichissement"] = result["status"]
        row["Score similarité"] = str(result["score"]) if result["score"] else ""

        if result["status"] == "SIRET trouvé":
            print(f"✅ {result['siret']} ({result['score']}%)")
        elif result["status"].startswith("Exclu — erreur"):
            print(f"⚠️  Erreur API")
            n_errors += 1
        else:
            print(f"❌ {result['status']}")

        enriched_rows.append(row)

        # Checkpoint
        if len(enriched_rows) % CHECKPOINT_EVERY == 0:
            save_checkpoint(enriched_rows, checkpoint_path)
            print(f"  💾 Checkpoint ({len(enriched_rows)} entreprises)")

    # Sauvegarde finale
    save_checkpoint(enriched_rows, checkpoint_path)

    # Export Excel
    print(f"\n── Export ──────────────────────────────────────")
    export_results(enriched_rows, output_path)
    print(f"  ✅ {output_path}")

    # CSV backup
    df_all = pd.DataFrame(enriched_rows, columns=ENRICHMENT_COLUMNS)
    df_all["SIRET"] = df_all["SIRET"].astype(str)
    df_all.to_csv(csv_backup_path, index=False, encoding="utf-8-sig")
    print(f"  ✅ {csv_backup_path}")

    # Cleanup checkpoint
    if os.path.exists(checkpoint_path):
        os.remove(checkpoint_path)

    # Synthèse
    print_synthese(input_file, enriched_rows, output_path, n_errors)


if __name__ == "__main__":
    main()
