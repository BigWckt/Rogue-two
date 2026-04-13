#!/usr/bin/env python3
"""
script_comparaison.py — Croisement PJ enrichi x LBA/LBB
========================================================
Croise les résultats Pages Jaunes enrichis (avec SIRET) et les résultats
LBA/LBB sur la base du SIRET, attribue un niveau de priorité (Haute /
Moyenne / Basse), et produit le fichier CSV final prêt pour import HubSpot.

Usage :
  python script_comparaison.py --pj pj_enrichi.csv --lba lba_lbb.csv
  python script_comparaison.py --lba lba_lbb.csv  # sans PJ
"""

import argparse
import os
import sys
from datetime import date

import pandas as pd

from matrix_display import (
    GREEN, RED, BOLD, RESET,
    matrix_banner, matrix_section, matrix_kv, matrix_separator,
    matrix_step, matrix_ok, matrix_fail, matrix_warn,
    morpheus_says, ask_filename,
)

# ── Colonnes de sortie (noms HubSpot) ────────────────────────────────────────

OUTPUT_COLUMNS = [
    "Nom de l'entreprise",
    "SIRET",
    "Code NAF",
    "Adresse",
    "Ville",
    "Code Postal",
    "Téléphone",
    "Site web",
    "Priorité",
    "Source",
    "Score LBB",
    "Offres actives",
    "Date de collecte",
]

PRIORITY_ORDER = {"Haute": 0, "Moyenne": 1, "Basse": 2}


# ── Helpers ───────────────────────────────────────────────────────────────────

def normalize_siret(raw) -> str:
    if raw is None or str(raw).strip() in ("", "nan", "None"):
        return ""
    try:
        return str(int(float(str(raw).strip()))).zfill(14)
    except (ValueError, OverflowError):
        return str(raw).strip()


def load_lba_lbb(filepath: str) -> dict[str, dict]:
    """Charge le fichier CSV LBA/LBB et retourne un dict SIRET -> row."""
    df = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str).fillna("")
    by_siret: dict[str, dict] = {}
    for _, row in df.iterrows():
        siret = normalize_siret(row.get("SIRET", ""))
        if not siret:
            continue
        rec = row.to_dict()
        rec["SIRET"] = siret
        if siret in by_siret:
            existing_src = by_siret[siret].get("Source", "")
            if "LBA" not in existing_src and "LBA" in rec.get("Source", ""):
                by_siret[siret] = rec
        else:
            by_siret[siret] = rec
    return by_siret


def load_pj_enrichi(filepath: str) -> dict[str, dict]:
    """Charge le fichier CSV PJ enrichi et retourne un dict SIRET -> row."""
    df = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str).fillna("")
    by_siret: dict[str, dict] = {}
    for _, row in df.iterrows():
        siret = normalize_siret(row.get("SIRET", ""))
        if not siret:
            continue
        rec = row.to_dict()
        rec["SIRET"] = siret
        if siret not in by_siret:
            by_siret[siret] = rec
    return by_siret


# ── Croisement ────────────────────────────────────────────────────────────────

def merge_and_score(pj_data: dict[str, dict],
                    lba_data: dict[str, dict]) -> tuple[list[dict], dict]:
    """
    Croise PJ et LBA/LBB par SIRET, attribue la priorité.
    Retourne (leads, stats).
    """
    today = date.today().isoformat()
    leads: list[dict] = []
    stats = {
        "pj_count": len(pj_data),
        "lba_count": len(lba_data),
        "matches": 0,
        "haute": 0,
        "moyenne": 0,
        "basse": 0,
    }

    all_sirets = set(pj_data.keys()) | set(lba_data.keys())

    for siret in all_sirets:
        pj = pj_data.get(siret)
        lba = lba_data.get(siret)

        lead: dict = {col: "" for col in OUTPUT_COLUMNS}
        lead["SIRET"] = siret
        lead["Date de collecte"] = today

        if pj and lba:
            stats["matches"] += 1
            lead["Nom de l'entreprise"] = pj.get("Nom de l'entreprise") or lba.get("Nom de l'entreprise", "")
            lead["Code NAF"] = pj.get("Code NAF") or lba.get("Code NAF", "")
            lead["Adresse"] = pj.get("Adresse") or lba.get("Adresse", "")
            lead["Ville"] = pj.get("Ville") or lba.get("Ville", "")
            lead["Code Postal"] = pj.get("Code Postal") or lba.get("Code Postal", "")
            lead["Téléphone"] = pj.get("Téléphone", "")
            lead["Site web"] = pj.get("Site web", "")
            lead["Score LBB"] = lba.get("Score LBB", "")
            lead["Offres actives"] = lba.get("Offres actives", "")

            lba_source = lba.get("Source", "")
            if "LBA" in lba_source:
                lead["Priorité"] = "Haute"
                lead["Source"] = "PJ + LBA" if "LBB" not in lba_source else "PJ + LBA + LBB"
            else:
                lead["Priorité"] = "Moyenne"
                lead["Source"] = "PJ + LBB"

        elif lba:
            lead["Nom de l'entreprise"] = lba.get("Nom de l'entreprise", "")
            lead["Code NAF"] = lba.get("Code NAF", "")
            lead["Adresse"] = lba.get("Adresse", "")
            lead["Ville"] = lba.get("Ville", "")
            lead["Code Postal"] = lba.get("Code Postal", "")
            lead["Score LBB"] = lba.get("Score LBB", "")
            lead["Offres actives"] = lba.get("Offres actives", "")

            lba_source = lba.get("Source", "")
            if "LBA" in lba_source:
                lead["Priorité"] = "Haute"
                lead["Source"] = lba_source
            else:
                lead["Priorité"] = "Moyenne"
                lead["Source"] = lba_source

        elif pj:
            lead["Nom de l'entreprise"] = pj.get("Nom de l'entreprise", "")
            lead["Code NAF"] = pj.get("Code NAF", "")
            lead["Adresse"] = pj.get("Adresse", "")
            lead["Ville"] = pj.get("Ville", "")
            lead["Code Postal"] = pj.get("Code Postal", "")
            lead["Téléphone"] = pj.get("Téléphone", "")
            lead["Site web"] = pj.get("Site web", "")
            lead["Priorité"] = "Basse"
            lead["Source"] = "Pages Jaunes"

        if lead["Priorité"] == "Haute":
            stats["haute"] += 1
        elif lead["Priorité"] == "Moyenne":
            stats["moyenne"] += 1
        elif lead["Priorité"] == "Basse":
            stats["basse"] += 1

        leads.append(lead)

    # Tri : Haute -> Moyenne -> Basse
    leads.sort(key=lambda r: PRIORITY_ORDER.get(r.get("Priorité", ""), 9))

    return leads, stats


# ── Export CSV ───────────────────────────────────────────────────────────────

def export_csv(leads: list[dict], csv_path: str):
    """Exporte un CSV unique avec colonne Priorité. SIRET en string."""
    df = pd.DataFrame(leads, columns=OUTPUT_COLUMNS)
    df["SIRET"] = df["SIRET"].astype(str)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")


# ── CLI ───────────────────────────────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="Croisement PJ enrichi x LBA/LBB — scoring et export HubSpot",
    )
    parser.add_argument("--pj", type=str, default=None,
                        help="Fichier PJ enrichi (.csv)")
    parser.add_argument("--lba", type=str, default=None,
                        help="Fichier LBA/LBB (.csv)")
    parser.add_argument("--output", type=str, default=".",
                        help="Répertoire de sortie (défaut: racine repo)")
    return parser.parse_args()


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    args = parse_args()

    if not args.pj and not args.lba:
        matrix_fail("Spécifiez au moins --pj ou --lba")
        sys.exit(1)

    os.makedirs(args.output, exist_ok=True)

    # ── Bannière Matrix ──
    matrix_banner("CROISEMENT FINAL — LEADS QUALIFIÉS")

    # Charger les données
    pj_data: dict[str, dict] = {}
    lba_data: dict[str, dict] = {}

    if args.pj:
        if not os.path.exists(args.pj):
            matrix_warn(f"Fichier PJ introuvable : {args.pj} — continué sans PJ")
        else:
            matrix_step(f"Chargement PJ enrichi : {args.pj}")
            pj_data = load_pj_enrichi(args.pj)
            matrix_ok(f"{len(pj_data)} entreprises avec SIRET")

    if args.lba:
        if not os.path.exists(args.lba):
            matrix_warn(f"Fichier LBA/LBB introuvable : {args.lba} — continué sans LBA/LBB")
        else:
            matrix_step(f"Chargement LBA/LBB : {args.lba}")
            lba_data = load_lba_lbb(args.lba)
            matrix_ok(f"{len(lba_data)} entreprises avec SIRET")

    if not pj_data and not lba_data:
        matrix_fail("Aucune donnée chargée")
        sys.exit(1)

    # ── Nom de fichier interactif ──
    today_str = date.today().strftime("%Y%m%d")
    # Déduire la ville depuis le nom de fichier
    ville = "multi"
    ref_file = args.lba or args.pj or ""
    base = os.path.basename(ref_file).lower()
    for prefix in ["lba_lbb_results_", "pj_results_"]:
        if prefix in base:
            parts = base.replace(prefix, "").split("_")
            if parts:
                ville = parts[0]
                break

    default_name = f"leads_qualifies_{ville}_{today_str}"
    filename = ask_filename(default_name)

    # ── Croisement ──
    matrix_section("Croisement par SIRET — fusion des réalités")
    leads, stats = merge_and_score(pj_data, lba_data)

    if not leads:
        matrix_warn("Aucun lead généré")
        sys.exit(0)

    # ── Export CSV ──
    csv_path = os.path.join(args.output, f"{filename}.csv")
    matrix_step("Export CSV final...")
    export_csv(leads, csv_path)
    matrix_ok(f"Fichier : {csv_path}")

    # ── Synthèse ──
    matrix_section("RÉSULTATS — La Matrice a livré ses secrets")
    if stats["pj_count"]:
        matrix_kv("PJ enrichi", f"{stats['pj_count']} entreprises (avec SIRET)")
    matrix_kv("LBA/LBB", f"{stats['lba_count']} entreprises")
    if stats["pj_count"]:
        matrix_kv("Correspondances PJ <> LBA/LBB", f"{stats['matches']}")
    matrix_separator()
    matrix_kv("Priorité Haute (LBA)", f"{stats['haute']} entreprises")
    matrix_kv("Priorité Moyenne (LBB)", f"{stats['moyenne']} entreprises")
    matrix_kv("Priorité Basse (PJ seul)", f"{stats['basse']} entreprises")
    matrix_separator()
    matrix_kv("Total leads qualifiés", f"{len(leads)} entreprises")
    matrix_kv("Fichier", os.path.basename(csv_path))

    # ── Clôture Matrix ──
    morpheus_says()


if __name__ == "__main__":
    main()
