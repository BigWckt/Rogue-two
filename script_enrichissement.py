#!/usr/bin/env python3
"""
script_enrichissement.py — Enrichissement SIRET via API recherche-entreprises
=============================================================================
Pour chaque entreprise issue de Pages Jaunes (sans SIRET), interroge l'API
recherche-entreprises.api.gouv.fr pour retrouver le SIRET à partir du nom
et du code postal / ville.

Usage :
  python script_enrichissement.py --input pj_results_multi_20260413.csv
  python script_enrichissement.py --input pj_results_multi_20260413.csv --resume
"""

import argparse
import os
import re
import glob
import subprocess
import sys
import time
from datetime import date

import pandas as pd
import requests
from rapidfuzz import fuzz

from batch_io import read_current_batch, check_naf_coherence, update_current_batch, get_naf_label
from matrix_display import (
    GREEN, RED, BOLD, RESET,
    matrix_banner, matrix_section, matrix_kv, matrix_separator,
    matrix_step, matrix_ok, matrix_fail, matrix_warn,
    morpheus_says, ask_filename, matrix_decode, set_quiet,
    matrix_rain_fullscreen,
)

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
    "Ville de recherche",
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
                matrix_warn(f"HTTP 429 — retry {attempt}/{MAX_RETRIES} dans {wait}s")
                time.sleep(wait)
                continue
            if resp.status_code in (500, 502, 503, 504):
                wait = 2 ** attempt
                matrix_warn(f"HTTP {resp.status_code} — retry {attempt}/{MAX_RETRIES} dans {wait}s")
                time.sleep(wait)
                continue
            if resp.status_code == 400:
                return {"results": []}  # Bad request — not retryable
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.Timeout:
            wait = 2 ** attempt
            matrix_warn(f"Timeout — retry {attempt}/{MAX_RETRIES} dans {wait}s")
            time.sleep(wait)
        except requests.exceptions.ConnectionError as e:
            wait = 2 ** attempt
            matrix_warn(f"Connexion : {e} — retry {attempt}/{MAX_RETRIES}")
            time.sleep(wait)
        except requests.exceptions.HTTPError as e:
            matrix_fail(f"HTTP {e.response.status_code}")
            return None
        except ValueError:
            matrix_fail("Réponse non-JSON")
            return None
    return None


def clean_name(name: str) -> str:
    """Normalise un nom d'entreprise pour la comparaison."""
    name = name.upper().strip()
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
    if not code_postal:
        return []
    return [
        r for r in results
        if (r.get("siege") or {}).get("code_postal", "") == code_postal
    ]


def _filter_by_commune(results: list[dict], ville: str) -> list[dict]:
    if not ville:
        return []
    ville_lower = ville.lower().strip()
    return [
        r for r in results
        if ville_lower in ((r.get("siege") or {}).get("libelle_commune") or "").lower()
    ]


def search_siret(nom: str, code_postal: str, ville: str) -> dict:
    """
    1 seul appel API (q=NOM, per_page=10, minimal=false), puis filtrage local :
      - Étape 1 : filtrer par code_postal exact
      - Étape 2 (si vide) : filtrer par commune
      - Étape 3 (si vide) : garder tous les résultats
    Fallback : si 0 résultat, 2e appel avec nom nettoyé. Max 2 appels.
    """
    nom_clean = clean_name(nom)
    if not nom_clean:
        return {"siret": "", "naf": "", "score": 0, "status": "Exclu — nom vide"}

    data = http_get(API_URL, params={
        "q": nom, "per_page": "10", "minimal": "false",
    })
    if data is None:
        return {"siret": "", "naf": "", "score": 0, "status": "Exclu — erreur API"}

    results = data.get("results") or []

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

    subset = _filter_by_cp(results, code_postal)
    if not subset:
        subset = _filter_by_commune(results, ville)
    if not subset:
        subset = results

    match = _best_match(subset, nom)
    if match:
        return match

    return {"siret": "", "naf": "", "score": 0, "status": "Exclu — SIRET introuvable"}


# ── Checkpoint / Resume ──────────────────────────────────────────────────────

def save_checkpoint(rows: list[dict], filepath: str):
    try:
        os.makedirs(os.path.dirname(filepath) or ".", exist_ok=True)
        df = pd.DataFrame(rows, columns=ENRICHMENT_COLUMNS)
        df.to_csv(filepath, index=False, encoding="utf-8-sig")
    except Exception as e:
        matrix_warn(f"Erreur checkpoint : {e}")


def load_checkpoint(filepath: str) -> list[dict]:
    if not os.path.exists(filepath):
        return []
    try:
        df = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str)
        df = df.fillna("")
        return df.to_dict("records")
    except Exception:
        return []


# ── Export CSV ───────────────────────────────────────────────────────────────

def export_csv(rows: list[dict], csv_path: str):
    """Exporte en CSV. SIRET en string."""
    os.makedirs(os.path.dirname(csv_path) or ".", exist_ok=True)
    df = pd.DataFrame(rows, columns=ENRICHMENT_COLUMNS)
    df["SIRET"] = df["SIRET"].astype(str)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Enrichissement SIRET — recherche via API recherche-entreprises",
    )
    parser.add_argument("--input", type=str, default=None, help="Fichier CSV PJ source")
    parser.add_argument("--output", type=str, default=".",
                        help="Répertoire de sortie (défaut: racine repo)")
    parser.add_argument("--resume", action="store_true",
                        help="Reprendre depuis le dernier checkpoint")
    parser.add_argument("--quiet", action="store_true",
                        help="Mode silencieux (désactive les animations)")
    parser.add_argument("--batch", type=str, default=None,
                        help="Chemin de batch forcé (ex: Prosp/SB/batch_2)")
    parser.add_argument("--chain", action="store_true",
                        help="Enchaîner automatiquement avec le croisement final")
    return parser.parse_args()


# ── Main ──────────────────────────────────────────────────────────────────────

def _interactive_select_input(search_dir: str) -> str:
    """Détecte les fichiers PJ dans le dossier et laisse l'utilisateur choisir."""
    candidates = sorted([
        f for f in os.listdir(search_dir)
        if f.startswith("pj_") and f.endswith(".csv") and "_enrichi" not in f
           and "_checkpoint" not in f and "_backup" not in f
    ])
    if not candidates:
        matrix_fail(f"Aucun fichier PJ détecté dans {search_dir}")
        sys.exit(1)

    print()
    matrix_section("Fichiers PJ détectés")
    for i, f in enumerate(candidates, 1):
        print(f"    {GREEN}{i}.{RESET} {f}")
    print()
    raw = input(f"    {GREEN}▸{RESET} Choix du fichier à enrichir (numéro) : ").strip()
    try:
        idx = int(raw) - 1
        if 0 <= idx < len(candidates):
            return os.path.join(search_dir, candidates[idx])
    except ValueError:
        pass
    matrix_fail("Choix invalide")
    sys.exit(1)


def _interactive_select_engine() -> str:
    """Laisse l'utilisateur choisir SIRENE ou PAPPERS."""
    print()
    matrix_section("Moteur d'enrichissement")
    print(f"    {GREEN}1.{RESET} SIRENE (API recherche-entreprises.api.gouv.fr — gratuit)")
    print(f"    {GREEN}2.{RESET} PAPPERS (API pappers.fr — nécessite clé API)")
    print()
    raw = input(f"    {GREEN}▸{RESET} Choix [défaut 1] : ").strip() or "1"
    return raw


def main():
    args = parse_args()
    if args.quiet:
        set_quiet(True)

    # Résolution du dossier de sortie : --batch > .current_batch > --output
    if args.batch:
        output_dir = args.batch
    else:
        batch_info = read_current_batch()
        if batch_info:
            output_dir = batch_info["BATCH_DIR"]
        else:
            output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)

    # Mode interactif si --input non fourni
    if not args.input:
        input_file = _interactive_select_input(output_dir)
        engine = _interactive_select_engine()
        if engine == "2":
            pappers_key = os.environ.get("PAPPERS_API_KEY", "").strip()
            if not pappers_key:
                matrix_fail("Variable PAPPERS_API_KEY non définie.")
                matrix_step("Lance : export PAPPERS_API_KEY=\"ta_clé\" (Linux/Mac)")
                matrix_step("  ou : $env:PAPPERS_API_KEY = \"ta_clé\" (PowerShell)")
                sys.exit(1)
            script_dir = os.path.dirname(os.path.abspath(__file__))
            cmd = [sys.executable, os.path.join(script_dir, "script_enrichissement_pappers.py"),
                   "--input", input_file]
            subprocess.run(cmd)
            sys.exit(0)
    else:
        input_file = args.input

    if not os.path.exists(input_file):
        matrix_fail(f"Fichier introuvable : {input_file}")
        sys.exit(1)

    # ── Bannière Matrix ──
    matrix_banner("ENRICHISSEMENT SIRET")

    # ── Chargement du fichier source (CSV) ──
    matrix_step(f"Chargement de {input_file}...")
    df = pd.read_csv(input_file, encoding="utf-8-sig", dtype=str).fillna("")
    fiches = df.to_dict("records")
    matrix_ok(f"{len(fiches)} entreprises chargées")

    matrix_kv("Source", os.path.basename(input_file))
    matrix_kv("Entreprises", str(len(fiches)))

    # ── Nom de fichier interactif ──
    base = os.path.splitext(os.path.basename(input_file))[0]
    default_name = f"{base}_enrichi"
    filename = ask_filename(default_name)

    csv_path = os.path.join(output_dir, f"{filename}.csv")
    checkpoint_path = os.path.join(output_dir, f"{base}_checkpoint.csv")

    # Resume
    already_done: dict[str, dict] = {}
    if args.resume:
        checkpoint_rows = load_checkpoint(checkpoint_path)
        if checkpoint_rows:
            for r in checkpoint_rows:
                key = (r.get("Nom de l'entreprise", "").lower(), r.get("Code Postal", ""))
                already_done[key] = r
            matrix_ok(f"Reprise : {len(already_done)} entreprises déjà traitées")

    # ── Enrichissement ──
    matrix_section("Décryptage des identités SIRET")
    enriched_rows: list[dict] = []
    n_errors = 0

    for i, fiche in enumerate(fiches, 1):
        nom = fiche.get("Nom de l'entreprise", "")
        cp = fiche.get("Code Postal", "")
        ville = fiche.get("Ville", "")

        key = (nom.lower(), cp)
        if key in already_done:
            enriched_rows.append(already_done[key])
            continue

        print(f"    {GREEN}[{i}/{len(fiches)}]{RESET} {nom[:50]}...", end=" ", flush=True)

        time.sleep(API_DELAY)
        result = search_siret(nom, cp, ville)

        row = {**fiche}
        row["SIRET"] = result["siret"]
        row["Code NAF"] = result["naf"]
        row["Statut enrichissement"] = result["status"]
        row["Score similarité"] = str(result["score"]) if result["score"] else ""

        if result["status"] == "SIRET trouvé":
            print(f"{GREEN}✓{RESET}")
            matrix_decode(f"{nom[:50]} → {result['siret']} ({result['score']}%)", prefix=f"    {GREEN}SIRET ▸{RESET} ")
        elif result["status"].startswith("Exclu — erreur"):
            print(f"{RED}! Erreur API{RESET}")
            n_errors += 1
        else:
            print(f"{RED}✗{RESET} {result['status']}")

        enriched_rows.append(row)

        if len(enriched_rows) % CHECKPOINT_EVERY == 0:
            save_checkpoint(enriched_rows, checkpoint_path)
            matrix_step(f"Checkpoint ({len(enriched_rows)} entreprises)")

    # Sauvegarde finale
    save_checkpoint(enriched_rows, checkpoint_path)

    # ── Filtre cohérence NAF ──
    batch_info = read_current_batch()
    naf_attendus_str = (batch_info or {}).get("NAF_ATTENDUS", "")
    naf_attendus = [n.strip() for n in naf_attendus_str.split(",") if n.strip()] if naf_attendus_str else []
    naf_stats = {"coherent": 0, "hors_profil": 0, "indisponible": 0, "detail": {}}

    if naf_attendus:
        matrix_section("Filtre cohérence NAF")
        for row in enriched_rows:
            if row.get("Statut enrichissement") != "SIRET trouvé":
                continue
            naf = row.get("Code NAF", "").strip()
            if not naf:
                naf_stats["indisponible"] += 1
                continue
            if check_naf_coherence(naf, naf_attendus):
                naf_stats["coherent"] += 1
            else:
                naf_stats["hors_profil"] += 1
                naf_stats["detail"][naf] = naf_stats["detail"].get(naf, 0) + 1
                row["Statut enrichissement"] = f"Exclu — NAF hors profil ({naf})"

        matrix_kv("NAF cohérent avec le profil", str(naf_stats["coherent"]))
        matrix_kv("NAF hors profil (exclus)", str(naf_stats["hors_profil"]))
        if naf_stats["detail"]:
            shown = 0
            for naf_code, cnt in sorted(naf_stats["detail"].items(), key=lambda x: -x[1]):
                label = get_naf_label(naf_code)
                suffix = f" — {label}" if label else ""
                print(f"      dont : {naf_code}{suffix} ×{cnt}")
                shown += 1
                if shown >= 10:
                    remaining = len(naf_stats["detail"]) - 10
                    if remaining > 0:
                        print(f"      ... et {remaining} autres codes NAF exclus")
                    break
        matrix_kv("NAF non disponible (conservés)", str(naf_stats["indisponible"]))

    # ── Export CSV ──
    matrix_section("Export des données décryptées")
    export_csv(enriched_rows, csv_path)
    update_current_batch(pj_enriched_file=os.path.abspath(csv_path))
    matrix_ok(f"Fichier : {csv_path}")

    # Cleanup checkpoint
    if os.path.exists(checkpoint_path):
        os.remove(checkpoint_path)

    # ── Synthèse ──
    n_total = len(enriched_rows)
    n_found = sum(1 for r in enriched_rows if r.get("Statut enrichissement") == "SIRET trouvé")
    n_excluded = sum(1 for r in enriched_rows if r.get("Statut enrichissement", "").startswith("Exclu"))
    pct = f"{n_found / n_total * 100:.0f}%" if n_total else "0%"

    matrix_section("RÉSULTATS — Identités extraites de la Matrice")
    matrix_kv("Entreprises traitées", str(n_total))
    matrix_kv("SIRET trouvés", f"{n_found} ({pct})")
    matrix_kv(f"Exclus (< {SIMILARITY_THRESHOLD}% similarité)", str(n_excluded))
    matrix_kv("Erreurs API", str(n_errors))
    matrix_separator()
    matrix_kv("Fichier enrichi", os.path.basename(csv_path))

    # ── Clôture Matrix ──
    morpheus_says()

    # ── Enchaînement pipeline vers croisement ──
    if args.chain:
        lba_files = [f for f in glob.glob(os.path.join(output_dir, "lba_lbb_*.csv"))
                     if not f.endswith("_backup.csv")]
        if lba_files:
            lba_file = max(lba_files, key=os.path.getmtime)
            if not args.quiet:
                matrix_rain_fullscreen(duration=4, next_title="PHASE 3 — CROISEMENT FINAL")
            script_dir = os.path.dirname(os.path.abspath(__file__))
            cmd = [sys.executable, os.path.join(script_dir, "script_comparaison.py"),
                   "--pj", csv_path, "--lba", lba_file]
            subprocess.run(cmd)
        else:
            matrix_warn("Aucun fichier LBA/LBB trouvé dans le batch — croisement impossible")


if __name__ == "__main__":
    main()
