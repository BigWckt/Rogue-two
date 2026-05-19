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

from batch_io import read_current_batch, check_naf_coherence, check_enseigne_excluded, check_btpm_excluded, get_naf_label
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
    os.makedirs(os.path.dirname(csv_path) or ".", exist_ok=True)
    df = pd.DataFrame(leads, columns=OUTPUT_COLUMNS)
    df["SIRET"] = df["SIRET"].astype(str)
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")


# ── Détection souple de fichiers ─────────────────────────────────────────────

def _find_csv_candidates(directory: str, keywords: list[str]) -> list[str]:
    """Cherche les CSV dont le nom contient l'un des mots-clés (insensible casse)."""
    if not os.path.isdir(directory):
        return []
    candidates = []
    for f in os.listdir(directory):
        if not f.lower().endswith(".csv"):
            continue
        basename_lower = f.lower()
        if "_backup" in basename_lower:
            continue
        if any(kw in basename_lower for kw in keywords):
            candidates.append(os.path.join(directory, f))
    return sorted(candidates, key=lambda p: os.path.getmtime(p), reverse=True)


def _select_file_interactive(candidates: list[str], label: str) -> str | None:
    """Propose les fichiers candidats à l'utilisateur."""
    if not candidates:
        return None
    if len(candidates) == 1:
        basename = os.path.basename(candidates[0])
        matrix_ok(f"Fichier {label} détecté : {basename}")
        choice = input(f"    Utiliser ce fichier ? (O/n) : ").strip().lower()
        if choice == "n":
            return None
        return candidates[0]
    matrix_ok(f"Fichiers {label} détectés :")
    for i, f in enumerate(candidates, 1):
        print(f"    {GREEN}{i}.{RESET} {os.path.basename(f)}")
    print()
    raw = input(f"    {GREEN}▸{RESET} Choix (numéro) : ").strip()
    try:
        idx = int(raw) - 1
        if 0 <= idx < len(candidates):
            return candidates[idx]
    except ValueError:
        pass
    matrix_warn("Choix invalide")
    return None


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

def _resolve_file(arg_value: str | None, batch_info: dict | None,
                   batch_key: str, output_dir: str,
                   keywords: list[str], label: str) -> str | None:
    """Cascade : --arg > .current_batch key > auto-detection."""
    if arg_value:
        return arg_value

    if batch_info and batch_key in batch_info:
        path = batch_info[batch_key]
        if os.path.isfile(path):
            matrix_ok(f"{label} (via .current_batch) : {os.path.basename(path)}")
            return path
        else:
            matrix_warn(f"{label} (via .current_batch) introuvable : {path}")

    candidates = _find_csv_candidates(output_dir, keywords)
    if candidates:
        return _select_file_interactive(candidates, label)

    matrix_warn(f"Aucun fichier {label} détecté dans le batch")
    return None


def _check_naf_coherence_warning(leads_data: dict[str, dict], source_label: str,
                                  naf_attendus: list[str], batch_info: dict | None,
                                  quiet: bool) -> bool:
    """Vérifie si >80% des NAF sont hors profil. Retourne True si OK, False si abandon."""
    if not naf_attendus or not leads_data:
        return True
    n_total = 0
    n_mismatch = 0
    for rec in leads_data.values():
        naf = rec.get("Code NAF", "").strip()
        if not naf:
            continue
        n_total += 1
        if not check_naf_coherence(naf, naf_attendus):
            n_mismatch += 1
    if n_total == 0:
        return True
    pct = n_mismatch / n_total
    if pct > 0.8:
        profils = (batch_info or {}).get("PROFILS", "inconnu")
        matrix_warn(
            f"ATTENTION : {n_mismatch}/{n_total} ({pct*100:.0f}%) des entreprises {source_label} "
            f"ont un NAF hors profil ({profils})"
        )
        if not quiet:
            choice = input("    Continuer quand même ? (O/n) > ").strip().lower()
            if choice == "n":
                return False
    return True


def main():
    args = parse_args()
    if args.quiet:
        set_quiet(True)

    # Résolution du dossier de sortie : --batch > .current_batch > --output
    batch_info = read_current_batch()
    if args.batch:
        output_dir = args.batch
    elif batch_info:
        output_dir = batch_info["BATCH_DIR"]
    else:
        output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)

    # ── Bannière Matrix ──
    matrix_banner("CROISEMENT FINAL — LEADS QUALIFIÉS")

    # ── Bannière de configuration ──
    if batch_info:
        matrix_section("Configuration du batch actif")
        matrix_kv("Secteur", batch_info.get("SECTEUR", "—"))
        matrix_kv("Batch", batch_info.get("BATCH_DIR", "—"))
        matrix_kv("Date", batch_info.get("DATE", "—"))
        profils = batch_info.get("PROFILS", "")
        if profils:
            matrix_kv("Profils", profils.replace(",", ", "))
        naf_str = batch_info.get("NAF_ATTENDUS", "")
        if naf_str:
            codes = [c.strip() for c in naf_str.split(",") if c.strip()]
            labeled = []
            for c in codes[:8]:
                lbl = get_naf_label(c)
                labeled.append(f"{c} ({lbl})" if lbl else c)
            display = ", ".join(labeled)
            if len(codes) > 8:
                display += f", ... (+{len(codes) - 8})"
            matrix_kv("NAF attendus", display)
        pj_file = batch_info.get("PJ_ENRICHED_FILE", "")
        if pj_file:
            matrix_kv("PJ enrichi", os.path.basename(pj_file))
        lba_file = batch_info.get("LBA_LBB_FILE", "")
        if lba_file:
            matrix_kv("LBA/LBB", os.path.basename(lba_file))

    # ── Résolution des fichiers (cascade) ──
    args.pj = _resolve_file(args.pj, batch_info, "PJ_ENRICHED_FILE",
                             output_dir, ["enrichi"], "PJ enrichi")
    args.lba = _resolve_file(args.lba, batch_info, "LBA_LBB_FILE",
                              output_dir, ["lba", "lbb"], "LBA/LBB")

    if not args.pj and not args.lba:
        matrix_fail("Aucun fichier détecté. Spécifiez --pj et/ou --lba")
        sys.exit(1)

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

    # ── Garde cohérence NAF ──
    naf_attendus_str = (batch_info or {}).get("NAF_ATTENDUS", "")
    naf_attendus = [n.strip() for n in naf_attendus_str.split(",") if n.strip()] if naf_attendus_str else []

    if naf_attendus:
        if lba_data:
            ok = _check_naf_coherence_warning(lba_data, "LBA/LBB", naf_attendus,
                                               batch_info, args.quiet)
            if not ok:
                matrix_fail("Abandonné par l'utilisateur (incohérence NAF)")
                sys.exit(0)
        if pj_data:
            ok = _check_naf_coherence_warning(pj_data, "PJ", naf_attendus,
                                               batch_info, args.quiet)
            if not ok:
                matrix_fail("Abandonné par l'utilisateur (incohérence NAF)")
                sys.exit(0)

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
            shown = 0
            for naf_code, cnt in sorted(naf_filter_stats["detail"].items(), key=lambda x: -x[1]):
                label = get_naf_label(naf_code)
                suffix = f" — {label}" if label else ""
                print(f"      dont : {naf_code}{suffix} ×{cnt}")
                shown += 1
                if shown >= 10:
                    remaining = len(naf_filter_stats["detail"]) - 10
                    if remaining > 0:
                        print(f"      ... et {remaining} autres codes NAF exclus")
                    break

    # ── Filtre enseignes nationales (MBT / TG) ──
    secteur = (batch_info or {}).get("SECTEUR", "").upper()
    enseigne_stats = {"lba_excluded": 0, "pj_excluded": 0, "detail": {}}
    if secteur in ("MBT", "TG"):
        for siret in list(lba_data.keys()):
            nom = lba_data[siret].get("Nom de l'entreprise", "")
            matched = check_enseigne_excluded(nom)
            if matched:
                enseigne_stats["lba_excluded"] += 1
                enseigne_stats["detail"][matched] = enseigne_stats["detail"].get(matched, 0) + 1
                del lba_data[siret]

        for siret in list(pj_data.keys()):
            nom = pj_data[siret].get("Nom de l'entreprise", "")
            matched = check_enseigne_excluded(nom)
            if matched:
                enseigne_stats["pj_excluded"] += 1
                enseigne_stats["detail"][matched] = enseigne_stats["detail"].get(matched, 0) + 1
                del pj_data[siret]

        total_enseigne = enseigne_stats["lba_excluded"] + enseigne_stats["pj_excluded"]
        if total_enseigne:
            matrix_section("Exclusions enseignes nationales")
            matrix_kv("LBA/LBB exclus", str(enseigne_stats["lba_excluded"]))
            matrix_kv("PJ exclus", str(enseigne_stats["pj_excluded"]))
            top = sorted(enseigne_stats["detail"].items(), key=lambda x: -x[1])[:10]
            detail = ", ".join(f"{k.title()} ×{v}" for k, v in top)
            if len(enseigne_stats["detail"]) > 10:
                detail += ", ..."
            print(f"      {detail}")

    # ── Filtre BTPM (admin / culte / cabinets / loueurs / asso) ──
    if secteur == "BTPM":
        btpm_stats = {"lba_excluded": 0, "pj_excluded": 0, "by_cat": {}}

        for siret in list(lba_data.keys()):
            nom = lba_data[siret].get("Nom de l'entreprise", "")
            motif, cat = check_btpm_excluded(nom)
            if motif:
                btpm_stats["lba_excluded"] += 1
                btpm_stats["by_cat"][cat] = btpm_stats["by_cat"].get(cat, 0) + 1
                del lba_data[siret]

        for siret in list(pj_data.keys()):
            nom = pj_data[siret].get("Nom de l'entreprise", "")
            motif, cat = check_btpm_excluded(nom)
            if motif:
                btpm_stats["pj_excluded"] += 1
                btpm_stats["by_cat"][cat] = btpm_stats["by_cat"].get(cat, 0) + 1
                del pj_data[siret]

        total_btpm = btpm_stats["lba_excluded"] + btpm_stats["pj_excluded"]
        if total_btpm:
            matrix_section("Exclusions BTPM (admin/culte/cabinet/loc/asso)")
            matrix_kv("LBA/LBB exclus", str(btpm_stats["lba_excluded"]))
            matrix_kv("PJ exclus", str(btpm_stats["pj_excluded"]))
            cats = sorted(btpm_stats["by_cat"].items(), key=lambda x: -x[1])
            detail_btpm = ", ".join(f"{c} ×{v}" for c, v in cats)
            print(f"      {detail_btpm}")

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
