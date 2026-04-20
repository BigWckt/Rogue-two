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
import re
import sys
import unicodedata
from datetime import date

import pandas as pd
from rapidfuzz import fuzz

from batch_io import read_current_batch, check_naf_coherence
from matrix_display import (
    GREEN, RED, BOLD, RESET,
    matrix_banner, matrix_section, matrix_kv, matrix_separator,
    matrix_step, matrix_ok, matrix_fail, matrix_warn,
    morpheus_says, ask_filename, matrix_decode, set_quiet,
    matrix_rain_fullscreen,
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
    "Source téléphone",
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


def normalize_cp(raw) -> str:
    """Normalise un code postal : 5 chiffres, padded à gauche."""
    if raw is None:
        return ""
    digits = re.sub(r"\D", "", str(raw).strip())
    if not digits:
        return ""
    return digits.zfill(5)


def strip_accents(s: str) -> str:
    """Supprime les accents (é→e, è→e, etc.)."""
    return unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode()


def clean_name(name: str) -> str:
    """Normalise un nom d'entreprise pour la comparaison fuzzy."""
    name = strip_accents(name).upper().strip()
    for suffix in ["SARL", "SAS", "SA", "EURL", "SCI", "SASU", "EI", "SELARL"]:
        name = re.sub(rf"\b{suffix}\b", "", name)
    return re.sub(r"\s+", " ", name).strip()


PHONE_FALLBACK_THRESHOLD = 85
DEDUP_PHONE_NAME_THRESHOLD = 70
DEDUP_FUZZY_NAME_THRESHOLD = 90


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


def load_pj_all_fiches(filepath: str) -> list[dict]:
    """Charge TOUTES les fiches PJ enrichies (avec ou sans SIRET) pour le fallback téléphone."""
    df = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str).fillna("")
    fiches = []
    for _, row in df.iterrows():
        rec = row.to_dict()
        tel = rec.get("Téléphone", "").strip()
        if not tel:
            continue
        rec["_cp_norm"] = normalize_cp(rec.get("Code Postal", ""))
        rec["_name_clean"] = clean_name(rec.get("Nom de l'entreprise", ""))
        fiches.append(rec)
    return fiches


# ── Fallback téléphone par nom + CP ─────────────────────────────────────────

def _build_pj_by_cp(pj_fiches: list[dict]) -> dict[str, list[dict]]:
    """Index les fiches PJ par code postal normalisé."""
    by_cp: dict[str, list[dict]] = {}
    for fiche in pj_fiches:
        cp = fiche["_cp_norm"]
        if cp:
            by_cp.setdefault(cp, []).append(fiche)
    return by_cp


def enrich_phones_fallback(leads: list[dict], pj_fiches: list[dict]) -> dict:
    """
    Étape 2 : pour chaque lead sans téléphone, tente un matching fuzzy
    nom + CP exact dans les fiches PJ.
    Retourne les stats d'enrichissement.
    """
    pj_by_cp = _build_pj_by_cp(pj_fiches)

    stats = {"already": 0, "added": 0, "missing": 0}

    for lead in leads:
        tel = lead.get("Téléphone", "").strip()
        if tel:
            stats["already"] += 1
            continue

        lead_cp = normalize_cp(lead.get("Code Postal", ""))
        lead_name = clean_name(lead.get("Nom de l'entreprise", ""))
        if not lead_cp or not lead_name:
            stats["missing"] += 1
            continue

        candidates = pj_by_cp.get(lead_cp, [])
        if not candidates:
            stats["missing"] += 1
            continue

        best_score = 0
        best_fiche = None
        for fiche in candidates:
            score = fuzz.token_sort_ratio(lead_name, fiche["_name_clean"])
            if score > best_score:
                best_score = score
                best_fiche = fiche

        if best_fiche and best_score >= PHONE_FALLBACK_THRESHOLD:
            lead["Téléphone"] = best_fiche.get("Téléphone", "")
            lead["Source téléphone"] = "PJ fallback nom+CP"
            stats["added"] += 1
        else:
            stats["missing"] += 1

    return stats


# ── Déduplication renforcée ───────────────────────────────────────────────────

def _normalize_phone_9(raw: str) -> str:
    """Normalise un téléphone sur 9 chiffres (sans le 0 initial)."""
    digits = re.sub(r"\D", "", str(raw).strip())
    if digits.startswith("33") and len(digits) >= 11:
        digits = "0" + digits[2:]
    if digits.startswith("0") and len(digits) == 10:
        return digits[1:]
    return ""


def _count_filled(row: dict) -> int:
    """Compte le nombre de colonnes remplies (heuristique de richesse)."""
    return sum(1 for v in row.values() if str(v).strip())


def _merge_missing(target: dict, source: dict):
    """Copie les champs vides de source vers target."""
    for k, v in source.items():
        if not str(target.get(k, "")).strip() and str(v).strip():
            target[k] = v


def dedup_by_siret(leads: list[dict]) -> tuple[list[dict], int]:
    """Étape 1 : déduplique par SIRET. Priorité au téléphone rempli, puis à la priorité."""
    by_siret: dict[str, list[int]] = {}
    for i, lead in enumerate(leads):
        siret = lead.get("SIRET", "").strip()
        if siret:
            by_siret.setdefault(siret, []).append(i)

    to_remove: set[int] = set()

    for siret, indices in by_siret.items():
        if len(indices) < 2:
            continue
        group = [(i, leads[i]) for i in indices]
        group.sort(key=lambda x: (
            0 if x[1].get("Téléphone", "").strip() else 1,
            PRIORITY_ORDER.get(x[1].get("Priorité", ""), 9),
            -_count_filled(x[1]),
        ))
        keeper_idx, keeper = group[0]
        for idx, row in group[1:]:
            _merge_missing(keeper, row)
            to_remove.add(idx)

    n_fused = len(to_remove)
    if to_remove:
        leads = [lead for i, lead in enumerate(leads) if i not in to_remove]
    return leads, n_fused


def dedup_by_phone(leads: list[dict]) -> tuple[list[dict], int]:
    """Étape 2 : déduplique par téléphone (même 9 chiffres + nom similaire ≥ 70%)."""
    by_phone: dict[str, list[int]] = {}
    for i, lead in enumerate(leads):
        phone = _normalize_phone_9(lead.get("Téléphone", ""))
        if phone:
            by_phone.setdefault(phone, []).append(i)

    to_remove: set[int] = set()

    for phone, indices in by_phone.items():
        if len(indices) < 2:
            continue
        remaining = [i for i in indices if i not in to_remove]
        if len(remaining) < 2:
            continue

        for j in range(len(remaining)):
            if remaining[j] in to_remove:
                continue
            for k in range(j + 1, len(remaining)):
                if remaining[k] in to_remove:
                    continue
                name_a = clean_name(leads[remaining[j]].get("Nom de l'entreprise", ""))
                name_b = clean_name(leads[remaining[k]].get("Nom de l'entreprise", ""))
                if fuzz.token_sort_ratio(name_a, name_b) >= DEDUP_PHONE_NAME_THRESHOLD:
                    a, b = remaining[j], remaining[k]
                    a_score = (
                        0 if leads[a].get("Téléphone", "").strip() else 1,
                        PRIORITY_ORDER.get(leads[a].get("Priorité", ""), 9),
                        -_count_filled(leads[a]),
                    )
                    b_score = (
                        0 if leads[b].get("Téléphone", "").strip() else 1,
                        PRIORITY_ORDER.get(leads[b].get("Priorité", ""), 9),
                        -_count_filled(leads[b]),
                    )
                    if a_score <= b_score:
                        _merge_missing(leads[a], leads[b])
                        to_remove.add(b)
                    else:
                        _merge_missing(leads[b], leads[a])
                        to_remove.add(a)

    n_fused = len(to_remove)
    if to_remove:
        leads = [lead for i, lead in enumerate(leads) if i not in to_remove]
    return leads, n_fused


def dedup_by_fuzzy_name(leads: list[dict]) -> tuple[list[dict], int, list[str]]:
    """Étape 3 : déduplique par nom fuzzy + même département (2 premiers chiffres CP)."""
    by_dept: dict[str, list[int]] = {}
    for i, lead in enumerate(leads):
        cp = normalize_cp(lead.get("Code Postal", ""))
        dept = cp[:2] if len(cp) >= 2 else ""
        if dept:
            by_dept.setdefault(dept, []).append(i)

    to_remove: set[int] = set()
    details: list[str] = []

    for dept, indices in by_dept.items():
        if len(indices) < 2:
            continue
        names_clean = {i: clean_name(leads[i].get("Nom de l'entreprise", ""))
                       for i in indices}

        for j in range(len(indices)):
            if indices[j] in to_remove:
                continue
            for k in range(j + 1, len(indices)):
                if indices[k] in to_remove:
                    continue
                a, b = indices[j], indices[k]
                score = fuzz.token_sort_ratio(names_clean[a], names_clean[b])
                if score >= DEDUP_FUZZY_NAME_THRESHOLD:
                    a_filled = _count_filled(leads[a])
                    b_filled = _count_filled(leads[b])
                    name_a = leads[a].get("Nom de l'entreprise", "")
                    cp_a = leads[a].get("Code Postal", "")
                    name_b = leads[b].get("Nom de l'entreprise", "")
                    cp_b = leads[b].get("Code Postal", "")
                    if a_filled >= b_filled:
                        _merge_missing(leads[a], leads[b])
                        to_remove.add(b)
                    else:
                        _merge_missing(leads[b], leads[a])
                        to_remove.add(a)
                    details.append(
                        f'"{name_a}" ({cp_a}) ≈ "{name_b}" ({cp_b}) → fusionné ({score}%)'
                    )

    n_fused = len(to_remove)
    if to_remove:
        leads = [lead for i, lead in enumerate(leads) if i not in to_remove]
    return leads, n_fused, details


def run_dedup_renforcee(leads: list[dict]) -> tuple[list[dict], dict]:
    """Exécute les 3 étapes de déduplication renforcée."""
    matrix_section("Déduplication renforcée")

    leads, n_siret = dedup_by_siret(leads)
    if n_siret:
        matrix_ok(f"Doublons SIRET fusionnés : {n_siret}")

    leads, n_phone = dedup_by_phone(leads)
    if n_phone:
        matrix_ok(f"Doublons téléphone fusionnés : {n_phone}")

    leads, n_fuzzy, fuzzy_details = dedup_by_fuzzy_name(leads)
    if n_fuzzy:
        matrix_ok(f"Doublons nom fuzzy détectés : {n_fuzzy}")
        for detail in fuzzy_details[:10]:
            print(f"      {detail}")
        if len(fuzzy_details) > 10:
            print(f"      ... et {len(fuzzy_details) - 10} autres")

    total_removed = n_siret + n_phone + n_fuzzy
    dedup_stats = {
        "siret": n_siret,
        "phone": n_phone,
        "fuzzy": n_fuzzy,
        "total": total_removed,
        "remaining": len(leads),
    }

    if total_removed:
        matrix_separator()
        matrix_kv("Doublons SIRET", f"{n_siret} lignes fusionnées")
        matrix_kv("Doublons téléphone", f"{n_phone} lignes fusionnées")
        matrix_kv("Doublons nom fuzzy", f"{n_fuzzy} lignes fusionnées")
        matrix_separator()
        matrix_kv("Total supprimé", f"{total_removed} lignes")
        matrix_kv("Leads uniques restants", str(len(leads)))
    else:
        matrix_ok("Aucun doublon détecté")

    return leads, dedup_stats


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
            lead["Site web"] = pj.get("Site web", "")
            lead["Score LBB"] = lba.get("Score LBB", "")
            lead["Offres actives"] = lba.get("Offres actives", "")

            # Téléphone : PJ prioritaire, fallback LBA/LBB
            pj_tel = pj.get("Téléphone", "").strip()
            lba_tel = lba.get("Téléphone", "").strip()
            if pj_tel:
                lead["Téléphone"] = pj_tel
                lead["Source téléphone"] = "PJ direct"
            elif lba_tel:
                lead["Téléphone"] = lba_tel
                lba_source_raw = lba.get("Source", "")
                lead["Source téléphone"] = "LBA" if "LBA" in lba_source_raw else "LBB"

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

            lba_tel = lba.get("Téléphone", "").strip()
            if lba_tel:
                lead["Téléphone"] = lba_tel
                lba_source_raw = lba.get("Source", "")
                lead["Source téléphone"] = "LBA" if "LBA" in lba_source_raw else "LBB"

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
            lead["Site web"] = pj.get("Site web", "")
            lead["Priorité"] = "Basse"
            lead["Source"] = "Pages Jaunes"

            pj_tel = pj.get("Téléphone", "").strip()
            if pj_tel:
                lead["Téléphone"] = pj_tel
                lead["Source téléphone"] = "PJ direct"

        if lead["Priorité"] == "Haute":
            stats["haute"] += 1
            matrix_decode(lead["Nom de l'entreprise"], prefix=f"    {GREEN}★ HAUTE ▸{RESET} ")
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

    if not args.pj and not args.lba:
        matrix_fail("Spécifiez au moins --pj ou --lba")
        sys.exit(1)

    # Résolution du dossier de sortie : --batch > .current_batch > --output
    if args.batch:
        output_dir = args.batch
    else:
        batch = read_current_batch()
        if batch:
            output_dir = batch["BATCH_DIR"]
        else:
            output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)

    # ── Bannière Matrix ──
    matrix_banner("CROISEMENT FINAL — LEADS QUALIFIÉS")

    # Charger les données
    pj_data: dict[str, dict] = {}
    lba_data: dict[str, dict] = {}
    pj_all_fiches: list[dict] = []

    if args.pj:
        if not os.path.exists(args.pj):
            matrix_warn(f"Fichier PJ introuvable : {args.pj} — continué sans PJ")
        else:
            matrix_step(f"Chargement PJ enrichi : {args.pj}")
            pj_data = load_pj_enrichi(args.pj)
            matrix_ok(f"{len(pj_data)} entreprises avec SIRET")
            # Charger toutes les fiches PJ (avec ou sans SIRET) pour le fallback téléphone
            pj_all_fiches = load_pj_all_fiches(args.pj)
            matrix_ok(f"{len(pj_all_fiches)} fiches PJ avec téléphone (pour fallback)")

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

    # ── Filtre cohérence NAF ──
    batch_info = read_current_batch()
    naf_attendus_str = (batch_info or {}).get("NAF_ATTENDUS", "")
    naf_attendus = [n.strip() for n in naf_attendus_str.split(",") if n.strip()] if naf_attendus_str else []
    naf_filter_stats = {"lba_excluded": 0, "pj_excluded": 0, "detail": {}}

    if naf_attendus:
        for siret in list(lba_data.keys()):
            naf = lba_data[siret].get("Code NAF", "").strip()
            if naf and not check_naf_coherence(naf, naf_attendus):
                naf_filter_stats["lba_excluded"] += 1
                naf_filter_stats["detail"][naf] = naf_filter_stats["detail"].get(naf, 0) + 1
                del lba_data[siret]

        for siret in list(pj_data.keys()):
            naf = pj_data[siret].get("Code NAF", "").strip()
            if naf and not check_naf_coherence(naf, naf_attendus):
                naf_filter_stats["pj_excluded"] += 1
                naf_filter_stats["detail"][naf] = naf_filter_stats["detail"].get(naf, 0) + 1
                del pj_data[siret]

        total_excluded = naf_filter_stats["lba_excluded"] + naf_filter_stats["pj_excluded"]
        if total_excluded:
            matrix_section("Filtre cohérence NAF")
            matrix_kv("LBA/LBB exclus (NAF hors profil)", str(naf_filter_stats["lba_excluded"]))
            matrix_kv("PJ exclus (NAF hors profil)", str(naf_filter_stats["pj_excluded"]))
            for naf_code, cnt in sorted(naf_filter_stats["detail"].items(), key=lambda x: -x[1]):
                print(f"      dont : {naf_code} ×{cnt}")

    # ── Étape 1 : Croisement par SIRET ──
    matrix_section("Étape 1 — Croisement par SIRET")
    leads, stats = merge_and_score(pj_data, lba_data)

    if not leads:
        matrix_warn("Aucun lead généré")
        sys.exit(0)

    # ── Étape 2 : Enrichissement téléphone par fallback nom + CP ──
    phone_stats = {"already": 0, "added": 0, "missing": 0}
    if pj_all_fiches:
        matrix_section("Étape 2 — Enrichissement téléphone (fallback nom+CP)")
        phone_stats = enrich_phones_fallback(leads, pj_all_fiches)
        matrix_ok(f"Téléphones ajoutés via PJ : {phone_stats['added']}")

    # ── Étape 3 : Déduplication renforcée ──
    leads, dedup_stats = run_dedup_renforcee(leads)

    # ── Export CSV ──
    csv_path = os.path.join(output_dir, f"{filename}.csv")
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

    # ── Stats téléphone ──
    if pj_all_fiches:
        matrix_separator()
        n_total_tel = phone_stats["already"] + phone_stats["added"]
        pct = f"{n_total_tel / len(leads) * 100:.0f}%" if leads else "0%"
        matrix_section("Enrichissement téléphone par fallback nom+CP")
        matrix_kv("Téléphones déjà présents", f"{phone_stats['already']}")
        matrix_kv("Téléphones ajoutés via PJ", f"{phone_stats['added']}")
        matrix_kv("Toujours sans téléphone", f"{phone_stats['missing']}")
        matrix_kv("Total avec téléphone (final)", f"{n_total_tel} / {len(leads)} ({pct})")

    # ── Stats déduplication ──
    if dedup_stats["total"]:
        matrix_separator()
        matrix_section("Déduplication renforcée (récap)")
        matrix_kv("Doublons SIRET", f"{dedup_stats['siret']} lignes fusionnées")
        matrix_kv("Doublons téléphone", f"{dedup_stats['phone']} lignes fusionnées")
        matrix_kv("Doublons nom fuzzy", f"{dedup_stats['fuzzy']} lignes fusionnées")
        matrix_separator()
        matrix_kv("Total supprimé", f"{dedup_stats['total']} lignes")
        matrix_kv("Leads uniques restants", str(dedup_stats['remaining']))

    matrix_separator()
    matrix_kv("Fichier", os.path.basename(csv_path))

    # ── Clôture Matrix ──
    morpheus_says()


if __name__ == "__main__":
    main()
