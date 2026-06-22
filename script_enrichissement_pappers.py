#!/usr/bin/env python3
"""
script_enrichissement_pappers.py — Enrichissement via API Pappers (expérimental)
================================================================================
Enrichit un fichier CSV d'entreprises via l'API Pappers (données complémentaires :
dirigeant, effectifs, chiffre d'affaires, forme juridique). Peut aussi comparer
le taux d'enrichissement Pappers vs SIRENE sur les mêmes données.

Usage :
  python script_enrichissement_pappers.py --input pj_results_multi_20260414.csv
  python script_enrichissement_pappers.py --input pj_results_multi_20260414_enrichi.csv --compare
"""

import argparse
import json
import os
import re
import sys
import time
import unicodedata
from datetime import date

import pandas as pd
import requests
from rapidfuzz import fuzz

from batch_io import read_current_batch, update_current_batch, score_match
from matrix_display import (
    GREEN, RED, BOLD, RESET,
    matrix_banner, matrix_section, matrix_kv, matrix_separator,
    matrix_step, matrix_ok, matrix_fail, matrix_warn,
    morpheus_says, ask_filename, matrix_decode, set_quiet,
)

# ── Config ────────────────────────────────────────────────────────────────────

PAPPERS_BASE_URL = "https://api.pappers.fr/v2"
API_DELAY = 0.5
API_TIMEOUT = 15
MAX_RETRIES = 2
SIMILARITY_THRESHOLD = 80
CHECKPOINT_EVERY = 25

# ── Colonnes de sortie ────────────────────────────────────────────────────────

OUTPUT_COLUMNS = [
    "Nom de l'entreprise",
    "Adresse",
    "Ville",
    "Code Postal",
    "Ville de recherche",
    "Téléphone",
    "Site web",
    "Catégorie",
    "Date de collecte",
    "SIRET",
    "SIREN",
    "Code NAF",
    "Effectifs",
    "Dirigeant",
    "Chiffre d'affaires",
    "Date de création",
    "Forme juridique",
    "Statut enrichissement",
    "Score similarité",
    "Source enrichissement",
]


# ── Helpers ───────────────────────────────────────────────────────────────────

def _get_api_key() -> str:
    key = os.environ.get("PAPPERS_API_KEY", "").strip()
    if not key:
        matrix_fail("Variable PAPPERS_API_KEY non définie.")
        matrix_step("Lance : export PAPPERS_API_KEY=\"ta_clé\" (Linux/Mac)")
        matrix_step("  ou : $env:PAPPERS_API_KEY = \"ta_clé\" (PowerShell)")
        sys.exit(1)
    return key


def normalize_siret(raw) -> str:
    if raw is None or str(raw).strip() in ("", "nan", "None"):
        return ""
    try:
        return str(int(float(str(raw).strip()))).zfill(14)
    except (ValueError, OverflowError):
        return str(raw).strip()


STOP_WORDS = {
    "sarl", "sas", "eurl", "sa", "srl", "sci", "scp", "snc", "sasu", "ei", "selarl",
    "boulangerie", "patisserie", "patissier", "cabinet", "agence", "societe", "ste",
    "ets", "etablissement", "maison", "le", "la", "les", "du", "de", "des",
    "au", "aux", "et", "en",
}


def strip_accents(s: str) -> str:
    return unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode()


def clean_name(name: str) -> str:
    """Nettoyage léger : accents + formes juridiques (pour fuzz.ratio)."""
    name = strip_accents(name).upper().strip()
    for w in ["SARL", "SAS", "EURL", "SCI", "SASU", "EI", "SELARL",
              "SRL", "SCP", "SNC", "CABINET", "AGENCE", "SOCIETE", "STE",
              "ETS", "ETABLISSEMENT"]:
        name = re.sub(rf"\b{w}\b", "", name)
    return re.sub(r"\s+", " ", name).strip()


def clean_name_aggressive(name: str) -> str:
    """Nettoyage agressif : supprime tous les stop-words sectoriels pour comparaison mots-clés."""
    name = strip_accents(name).lower()
    name = re.sub(r"['\-]", " ", name)
    name = re.sub(r"[^a-z0-9 ]", "", name)
    words = [w for w in name.split() if w not in STOP_WORDS and len(w) > 1]
    return " ".join(words)


def _score_candidate(nom_query: str, cp_query: str, ville_query: str,
                     candidate: dict) -> tuple[int, str]:
    """
    Score multi-critères : nom (50%) + code postal (+30) + ville (+20).
    Délègue à batch_io.score_match() (logique partagée avec l'enrichissement
    SIRENE). Itère sur tous les noms exposés par Pappers et garde le meilleur.
    Retourne (score_final, meilleur_nom_trouvé).
    """
    siege = candidate.get("siege") or {}

    # Collecter tous les noms disponibles
    api_names = []
    for field in ["nom_entreprise", "denomination", "nom_complet", "sigle", "nom"]:
        v = candidate.get(field)
        if v:
            api_names.append(str(v))
    enseigne = siege.get("enseigne") or ""
    if enseigne:
        api_names.append(enseigne)
    if not api_names:
        return 0, ""

    cp_api = str(siege.get("code_postal") or "").strip()
    ville_api = str(siege.get("ville") or "").strip()

    best_score = -1
    best_api_name = api_names[0]
    for api_name in api_names:
        score, _ = score_match(
            nom_query, api_name, cp_query, cp_api, ville_query, ville_api,
        )
        if score > best_score:
            best_score = score
            best_api_name = api_name

    return int(best_score), best_api_name


def http_get(url, params=None):
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(url, params=params, timeout=API_TIMEOUT)
            if resp.status_code == 429:
                wait = 2 ** attempt
                matrix_warn(f"HTTP 429 — retry {attempt}/{MAX_RETRIES} dans {wait}s")
                time.sleep(wait)
                continue
            if resp.status_code in (500, 502, 503, 504):
                wait = 2 ** attempt
                matrix_warn(f"HTTP {resp.status_code} — retry {attempt}/{MAX_RETRIES} dans {wait}s")
                time.sleep(wait)
                continue
            if resp.status_code == 401:
                matrix_fail("Clé API Pappers invalide ou expirée (HTTP 401)")
                return None
            if resp.status_code == 404:
                return None
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


# ── Extraction données Pappers ────────────────────────────────────────────────

def _extract_company_data(company: dict) -> dict:
    """Extrait les champs utiles d'un résultat Pappers."""
    siege = company.get("siege") or {}

    siret = normalize_siret(siege.get("siret") or company.get("siret") or "")
    siren = str(company.get("siren") or "").strip()
    naf = (siege.get("code_naf") or company.get("code_naf")
           or siege.get("activite_principale") or "")
    effectifs = str(company.get("effectif") or company.get("tranche_effectif") or "")

    representants = company.get("representants") or company.get("dirigeants") or []
    dirigeant = ""
    if representants:
        rep = representants[0]
        dirigeant = (rep.get("nom_complet")
                     or f"{rep.get('prenom', '')} {rep.get('nom', '')}".strip())

    finances = company.get("finances") or company.get("comptes") or []
    ca = ""
    if finances:
        ca = str(finances[0].get("chiffre_affaires") or "")

    date_creation = company.get("date_creation") or ""
    forme_juridique = company.get("forme_juridique") or ""

    adresse = siege.get("adresse_ligne_1") or ""
    cp = siege.get("code_postal") or ""
    ville = siege.get("ville") or ""

    return {
        "SIRET": siret,
        "SIREN": siren,
        "Code NAF": naf,
        "Effectifs": effectifs,
        "Dirigeant": dirigeant,
        "Chiffre d'affaires": ca,
        "Date de création": date_creation,
        "Forme juridique": forme_juridique,
        "Adresse Pappers": adresse,
        "Code Postal Pappers": cp,
        "Ville Pappers": ville,
    }


# ── Recherche Pappers ─────────────────────────────────────────────────────────

def search_by_siret(siret: str, api_key: str) -> dict | None:
    """Recherche directe par SIRET."""
    data = http_get(f"{PAPPERS_BASE_URL}/entreprise", params={
        "siret": siret,
        "api_token": api_key,
    })
    if not data:
        return None
    return _extract_company_data(data)


_first_response_printed = False

SCORE_THRESHOLD_HIGH = 80   # confiance normale
SCORE_THRESHOLD_LOW  = 60   # confiance moyenne (SIRET accepté avec flag)
SCORE_SINGLE_CP      = 50   # résultat unique avec bon CP → accepté directement


def search_by_name(nom: str, code_postal: str, ville: str,
                   api_key: str) -> dict | None:
    """
    Recherche par nom + code postal. Scoring multi-critères :
      score = nom×0.5 + bonus_CP(+30) + bonus_ville(+20) → max 100
    Seuils : ≥80 confiance normale, ≥60 confiance moyenne, résultat unique avec CP → ≥50.
    Retourne None si aucun candidat ne passe les seuils.
    Définit _score, _status, _best_candidate_name sur le résultat.
    """
    params = {"q": nom, "api_token": api_key, "par_page": "5"}
    if code_postal:
        params["code_postal"] = code_postal

    data = http_get(f"{PAPPERS_BASE_URL}/recherche", params=params)
    if not data:
        return None

    global _first_response_printed
    if not _first_response_printed:
        _first_response_printed = True
        matrix_step("Structure de la première réponse Pappers :")
        print(json.dumps(data, indent=2, ensure_ascii=False)[:2000])
        print("...")

    results = data.get("resultats") or data.get("results") or []
    if not results:
        return None

    best_score = 0
    best_result = None
    best_name = ""

    for r in results:
        score, api_name = _score_candidate(nom, code_postal, ville, r)
        if score > best_score:
            best_score = score
            best_result = r
            best_name = api_name

    if best_result is None:
        return None

    # Cas spécial : résultat unique avec bon CP → seuil abaissé
    threshold = SCORE_THRESHOLD_LOW
    if len(results) == 1 and code_postal:
        cp_api = str((best_result.get("siege") or {}).get("code_postal") or "")
        if cp_api == code_postal:
            threshold = SCORE_SINGLE_CP

    if best_score < threshold:
        # Retourner quand même le meilleur candidat pour affichage console
        return {
            "_status": "Non trouvé",
            "_score": 0,
            "_best_candidate_name": best_name,
            "_best_candidate_score": best_score,
        }

    extracted = _extract_company_data(best_result)
    extracted["_score"] = best_score
    extracted["_best_candidate_name"] = best_name
    if best_score >= SCORE_THRESHOLD_HIGH:
        extracted["_status"] = "Trouvé par nom"
    else:
        extracted["_status"] = "SIRET trouvé (confiance moyenne)"
    return extracted


def enrich_one(nom: str, code_postal: str, ville: str, siret_existing: str,
               api_key: str) -> dict:
    """Enrichit une entreprise via Pappers. Retourne les données ou un status."""
    # Si on a déjà un SIRET, chercher directement
    if siret_existing:
        result = search_by_siret(siret_existing, api_key)
        if result:
            result["_score"] = 100
            result["_status"] = "SIRET trouvé (direct)"
            return result

    # Recherche par nom + CP + ville
    result = search_by_name(nom, code_postal, ville, api_key)
    if result and result.get("_status") != "Non trouvé":
        return result

    # Garder le meilleur candidat pour affichage si échec
    best_candidate = result  # peut contenir _best_candidate_name/_best_candidate_score

    # Retry sans code postal si le premier essai a échoué
    if code_postal:
        time.sleep(API_DELAY)
        result2 = search_by_name(nom, "", ville, api_key)
        if result2 and result2.get("_status") != "Non trouvé":
            result2["_status"] = result2["_status"].replace("Trouvé par nom",
                                                             "Trouvé par nom (sans CP)")
            return result2
        # Garder le meilleur candidat du retry si meilleur score
        if result2 and result2.get("_best_candidate_score", 0) > (
                best_candidate or {}).get("_best_candidate_score", 0):
            best_candidate = result2

    return {
        "_status": "Non trouvé",
        "_score": 0,
        "_best_candidate_name": (best_candidate or {}).get("_best_candidate_name", ""),
        "_best_candidate_score": (best_candidate or {}).get("_best_candidate_score", 0),
    }


# ── Checkpoint ────────────────────────────────────────────────────────────────

def save_checkpoint(rows: list[dict], filepath: str):
    try:
        os.makedirs(os.path.dirname(filepath) or ".", exist_ok=True)
        df = pd.DataFrame(rows)
        df.to_csv(filepath, index=False, encoding="utf-8-sig")
    except Exception as e:
        matrix_warn(f"Erreur checkpoint : {e}")


def load_checkpoint(filepath: str) -> list[dict]:
    if not os.path.exists(filepath):
        return []
    try:
        df = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str).fillna("")
        return df.to_dict("records")
    except Exception:
        return []


# ── Mode compare ──────────────────────────────────────────────────────────────

def run_compare(fiches: list[dict], api_key: str, output_dir: str, filename: str):
    """Compare SIRENE vs Pappers sur un fichier déjà enrichi."""
    matrix_section("Comparaison SIRENE vs Pappers")

    stats = {
        "identiques": 0,
        "differents": 0,
        "pappers_seul": 0,
        "sirene_seul": 0,
        "non_trouves": 0,
    }

    enriched_rows: list[dict] = []

    for i, fiche in enumerate(fiches, 1):
        nom = fiche.get("Nom de l'entreprise", "")
        cp = fiche.get("Code Postal", "")
        ville = fiche.get("Ville", "")
        siret_sirene = normalize_siret(fiche.get("SIRET", ""))

        print(f"    {GREEN}[{i}/{len(fiches)}]{RESET} {nom[:50]}...", end=" ", flush=True)

        time.sleep(API_DELAY)
        result = enrich_one(nom, cp, ville, siret_sirene, api_key)

        siret_pappers = result.get("SIRET", "")
        row = {**fiche}

        if siret_sirene and siret_pappers:
            if siret_sirene == siret_pappers:
                stats["identiques"] += 1
                row["Comparaison"] = "Identique"
                print(f"{GREEN}={RESET}")
            else:
                stats["differents"] += 1
                row["Comparaison"] = f"Différent (Pappers: {siret_pappers})"
                print(f"{RED}≠{RESET} Pappers: {siret_pappers}")
        elif siret_pappers and not siret_sirene:
            stats["pappers_seul"] += 1
            row["Comparaison"] = "Pappers seul"
            row["SIRET Pappers"] = siret_pappers
            print(f"{GREEN}+ Pappers{RESET}")
            matrix_decode(f"{nom[:40]} → {siret_pappers}", prefix=f"    {GREEN}PAPPERS ▸{RESET} ")
        elif siret_sirene and not siret_pappers:
            stats["sirene_seul"] += 1
            row["Comparaison"] = "SIRENE seul"
            best = result.get("_best_candidate_name", "")
            best_s = result.get("_best_candidate_score", 0)
            if best:
                print(f"SIRENE seul  {RED}[Pappers: \"{best[:35]}\" ({best_s}pts)]{RESET}")
            else:
                print(f"SIRENE seul")
        else:
            stats["non_trouves"] += 1
            row["Comparaison"] = "Non trouvé"
            best = result.get("_best_candidate_name", "")
            best_s = result.get("_best_candidate_score", 0)
            if best:
                print(f"{RED}✗{RESET}")
                print(f"             Meilleur candidat : \"{best[:45]}\" ({best_s}pts)")
            else:
                print(f"{RED}✗{RESET}")

        # Ajouter données Pappers si trouvées
        for col in ["Effectifs", "Dirigeant", "Chiffre d'affaires",
                    "Date de création", "Forme juridique"]:
            row[f"{col} (Pappers)"] = result.get(col, "")

        enriched_rows.append(row)

        if len(enriched_rows) % CHECKPOINT_EVERY == 0:
            save_checkpoint(enriched_rows, os.path.join(output_dir, "_pappers_compare_checkpoint.csv"))

    # Export
    csv_path = os.path.join(output_dir, f"{filename}.csv")
    os.makedirs(os.path.dirname(csv_path) or ".", exist_ok=True)
    df = pd.DataFrame(enriched_rows)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    update_current_batch(pj_enriched_file=os.path.abspath(csv_path))
    matrix_ok(f"Fichier : {csv_path}")

    # Récap
    matrix_section("Comparaison SIRENE vs Pappers")
    matrix_kv("SIRET identiques", str(stats["identiques"]))
    matrix_kv("SIRET différents", str(stats["differents"]))
    matrix_kv("Trouvés par Pappers seul", str(stats["pappers_seul"]))
    matrix_kv("Trouvés par SIRENE seul", str(stats["sirene_seul"]))
    matrix_kv("Non trouvés par les deux", str(stats["non_trouves"]))
    matrix_separator()
    total = len(fiches)
    taux_sirene = stats["identiques"] + stats["differents"] + stats["sirene_seul"]
    taux_pappers = stats["identiques"] + stats["differents"] + stats["pappers_seul"]
    matrix_kv("Taux SIRENE", f"{taux_sirene}/{total} ({taux_sirene*100//total}%)" if total else "0")
    matrix_kv("Taux Pappers", f"{taux_pappers}/{total} ({taux_pappers*100//total}%)" if total else "0")

    return csv_path


# ── Mode enrichissement standard ──────────────────────────────────────────────

def run_enrich(fiches: list[dict], api_key: str, output_dir: str, filename: str,
               resume_data: dict):
    """Enrichissement Pappers standard (sans SIRET existant)."""
    matrix_section("Enrichissement Pappers")

    enriched_rows: list[dict] = []
    n_found = 0
    n_errors = 0
    checkpoint_path = os.path.join(output_dir, "_pappers_enrich_checkpoint.csv")

    for i, fiche in enumerate(fiches, 1):
        nom = fiche.get("Nom de l'entreprise", "")
        cp = fiche.get("Code Postal", "")
        ville = fiche.get("Ville", "")

        key = (nom.lower(), cp)
        if key in resume_data:
            enriched_rows.append(resume_data[key])
            continue

        siret_existing = normalize_siret(fiche.get("SIRET", ""))

        print(f"    {GREEN}[{i}/{len(fiches)}]{RESET} {nom[:50]}...", end=" ", flush=True)

        time.sleep(API_DELAY)
        result = enrich_one(nom, cp, ville, siret_existing, api_key)

        row = {**fiche}
        status = result.get("_status", "Non trouvé")
        score = result.get("_score", 0)

        if status != "Non trouvé":
            n_found += 1
            row["SIRET"] = result.get("SIRET", "") or row.get("SIRET", "")
            row["SIREN"] = result.get("SIREN", "")
            row["Code NAF"] = result.get("Code NAF", "") or row.get("Code NAF", "")
            row["Effectifs"] = result.get("Effectifs", "")
            row["Dirigeant"] = result.get("Dirigeant", "")
            row["Chiffre d'affaires"] = result.get("Chiffre d'affaires", "")
            row["Date de création"] = result.get("Date de création", "")
            row["Forme juridique"] = result.get("Forme juridique", "")
            row["Statut enrichissement"] = status
            row["Score similarité"] = str(score) if score else ""
            row["Source enrichissement"] = "Pappers"
            label = f"({score}pts)" if score < SCORE_THRESHOLD_HIGH else f"({score}%)"
            print(f"{GREEN}✓{RESET}")
            matrix_decode(f"{nom[:40]} → {row['SIRET']} {label}",
                          prefix=f"    {GREEN}PAPPERS ▸{RESET} ")
        else:
            row.setdefault("SIRET", "")
            row.setdefault("SIREN", "")
            row["Statut enrichissement"] = status
            row["Score similarité"] = ""
            row["Source enrichissement"] = ""
            best = result.get("_best_candidate_name", "")
            best_s = result.get("_best_candidate_score", 0)
            print(f"{RED}✗{RESET}")
            if best:
                print(f"             Meilleur candidat : \"{best[:45]}\" ({best_s}pts)")

        enriched_rows.append(row)

        if len(enriched_rows) % CHECKPOINT_EVERY == 0:
            save_checkpoint(enriched_rows, checkpoint_path)
            matrix_step(f"Checkpoint ({len(enriched_rows)} entreprises)")

    # Export
    csv_path = os.path.join(output_dir, f"{filename}.csv")
    os.makedirs(os.path.dirname(csv_path) or ".", exist_ok=True)
    df = pd.DataFrame(enriched_rows)
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    update_current_batch(pj_enriched_file=os.path.abspath(csv_path))
    matrix_ok(f"Fichier : {csv_path}")

    # Cleanup checkpoint
    if os.path.exists(checkpoint_path):
        os.remove(checkpoint_path)

    # Synthèse
    n_total = len(enriched_rows)
    pct = f"{n_found / n_total * 100:.0f}%" if n_total else "0%"

    matrix_section("RÉSULTATS — Enrichissement Pappers")
    matrix_kv("Entreprises traitées", str(n_total))
    matrix_kv("Trouvés via Pappers", f"{n_found} ({pct})")
    matrix_kv("Non trouvés", str(n_total - n_found))
    matrix_separator()
    matrix_kv("Fichier enrichi", os.path.basename(csv_path))

    return csv_path


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Enrichissement Pappers — données complémentaires (dirigeant, effectifs, CA)",
    )
    parser.add_argument("--input", required=True, help="Fichier CSV source")
    parser.add_argument("--output", type=str, default=".",
                        help="Répertoire de sortie (défaut: racine repo)")
    parser.add_argument("--compare", action="store_true",
                        help="Mode comparaison SIRENE vs Pappers")
    parser.add_argument("--resume", action="store_true",
                        help="Reprendre depuis le dernier checkpoint")
    parser.add_argument("--quiet", action="store_true",
                        help="Mode silencieux (désactive les animations)")
    parser.add_argument("--batch", type=str, default=None,
                        help="Chemin de batch forcé (ex: Prosp/SB/batch_2)")
    return parser.parse_args()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()
    if args.quiet:
        set_quiet(True)

    api_key = _get_api_key()
    input_file = args.input

    if not os.path.exists(input_file):
        matrix_fail(f"Fichier introuvable : {input_file}")
        sys.exit(1)

    # Résolution du dossier de sortie
    if args.batch:
        output_dir = args.batch
    else:
        batch = read_current_batch()
        if batch:
            output_dir = batch["BATCH_DIR"]
        else:
            output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)

    # Bannière
    matrix_banner("ENRICHISSEMENT PAPPERS (expérimental)")

    # Chargement
    matrix_step(f"Chargement de {input_file}...")
    df = pd.read_csv(input_file, encoding="utf-8-sig", dtype=str).fillna("")
    fiches = df.to_dict("records")
    matrix_ok(f"{len(fiches)} entreprises chargées")
    matrix_kv("Source", os.path.basename(input_file))
    matrix_kv("Mode", "Comparaison SIRENE vs Pappers" if args.compare else "Enrichissement Pappers")

    # Nom de fichier
    base = os.path.splitext(os.path.basename(input_file))[0]
    if args.compare:
        default_name = f"{base}_compare_pappers"
    else:
        default_name = f"{base}_enrichi_pappers"
    filename = ask_filename(default_name)

    if args.compare:
        csv_path = run_compare(fiches, api_key, output_dir, filename)
    else:
        resume_data: dict = {}
        if args.resume:
            checkpoint_path = os.path.join(output_dir, "_pappers_enrich_checkpoint.csv")
            checkpoint_rows = load_checkpoint(checkpoint_path)
            if checkpoint_rows:
                for r in checkpoint_rows:
                    key = (r.get("Nom de l'entreprise", "").lower(),
                           r.get("Code Postal", ""))
                    resume_data[key] = r
                matrix_ok(f"Reprise : {len(resume_data)} entreprises déjà traitées")
        csv_path = run_enrich(fiches, api_key, output_dir, filename, resume_data)

    # Journal
    _update_journal(args, fiches, csv_path)

    # Clôture
    morpheus_says()


def _update_journal(args, fiches, csv_path):
    """Append une section datée dans JOURNAL.md."""
    today = date.today().isoformat()
    mode = "Comparaison" if args.compare else "Enrichissement"
    entry = f"""
---

## {today} — script_enrichissement_pappers.py ({mode})

### Paramètres
- Fichier source : `{os.path.basename(args.input)}`
- Mode : {mode}
- Entreprises traitées : {len(fiches)}

### Fichier produit
- `{os.path.basename(csv_path)}`

### État
**Exécuté** ✓
"""

    journal_path = "JOURNAL.md"
    batch = read_current_batch()
    if batch:
        batch_dir = batch.get("BATCH_DIR", ".")
        alt_path = os.path.join(batch_dir, "JOURNAL.md")
        if os.path.exists(alt_path):
            journal_path = alt_path

    if not os.path.exists(journal_path):
        journal_path = "JOURNAL.md"

    try:
        with open(journal_path, "a", encoding="utf-8") as f:
            f.write(entry)
        matrix_ok(f"JOURNAL.md mis à jour")
    except Exception as e:
        matrix_warn(f"Impossible de mettre à jour JOURNAL.md : {e}")


if __name__ == "__main__":
    main()
