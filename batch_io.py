#!/usr/bin/env python3
"""
batch_io.py — Gestion du batch actif (.current_batch)
======================================================
Module partagé entre les 4 scripts de prospection.
Écrit et lit le fichier .current_batch qui trace le dossier actif.
"""

import os
import re
import unicodedata

CURRENT_BATCH_FILE = ".current_batch"

# ── Secteurs ─────────────────────────────────────────────────────────────────

SECTEUR_PREFIXES = {
    "sb":   "SB",
    "btpm": "BTPM",
    "mbt":  "MBT",
    "tg":   "TG",
}


def detect_secteur(profile_names: list[str]) -> str | None:
    """Déduit le secteur depuis les noms de profil (ex: ['sb_sante'] -> 'SB')."""
    secteurs = set()
    for name in profile_names:
        for prefix, label in SECTEUR_PREFIXES.items():
            if name.lower().startswith(f"{prefix}_"):
                secteurs.add(label)
                break
    if len(secteurs) == 1:
        return secteurs.pop()
    return None  # Ambigu ou inconnu


# ── Recherche du dossier Prosp/ ───────────────────────────────────────────────

def find_prosp_root(start: str = ".", max_levels: int = 3) -> str | None:
    """Remonte jusqu'à max_levels niveaux pour trouver un dossier Prosp/."""
    current = os.path.abspath(start)
    for _ in range(max_levels + 1):
        candidate = os.path.join(current, "Prosp")
        if os.path.isdir(candidate):
            return candidate
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    return None


# ── Numérotation de batch ─────────────────────────────────────────────────────

def next_batch_number(secteur_dir: str) -> int:
    """Retourne le prochain numéro de batch dans secteur_dir."""
    if not os.path.isdir(secteur_dir):
        return 1
    existing = []
    for name in os.listdir(secteur_dir):
        m = re.match(r"^batch_(\d+)$", name)
        if m and os.path.isdir(os.path.join(secteur_dir, name)):
            existing.append(int(m.group(1)))
    return max(existing) + 1 if existing else 1


# ── Lecture / écriture de .current_batch ─────────────────────────────────────

def write_current_batch(secteur: str, batch_dir: str, date_str: str,
                        root: str = ".", profils: str = "",
                        naf_attendus: str = ""):
    """Écrit le fichier .current_batch à la racine fournie."""
    path = os.path.join(root, CURRENT_BATCH_FILE)
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"SECTEUR={secteur}\n")
        f.write(f"BATCH_DIR={batch_dir}\n")
        f.write(f"DATE={date_str}\n")
        if profils:
            f.write(f"PROFILS={profils}\n")
        if naf_attendus:
            f.write(f"NAF_ATTENDUS={naf_attendus}\n")
    return path


def read_current_batch(root: str = ".") -> dict | None:
    """
    Lit .current_batch en remontant depuis root jusqu'à 3 niveaux.
    Retourne {'SECTEUR': ..., 'BATCH_DIR': ..., 'DATE': ...} ou None.
    """
    current = os.path.abspath(root)
    for _ in range(4):
        path = os.path.join(current, CURRENT_BATCH_FILE)
        if os.path.isfile(path):
            data = {}
            with open(path, encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if "=" in line:
                        k, v = line.split("=", 1)
                        data[k.strip()] = v.strip()
            if "BATCH_DIR" in data:
                return data
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    return None


def _normalize_naf(code: str) -> str:
    """Normalise un code NAF : insère le point après les 2 premiers chiffres si absent."""
    code = code.strip().upper()
    if len(code) >= 4 and "." not in code and code[:2].isdigit():
        return code[:2] + "." + code[2:]
    return code


def match_naf(code_profil: str, code_entreprise: str) -> bool:
    """
    Compare un code NAF du profil avec un code NAF d'entreprise.
    Normalise les deux codes (insertion du point si absent : 1071C → 10.71C).
    - Si code_profil a une lettre finale (ex: '86.10Z') → match exact uniquement
    - Si code_profil n'a pas de lettre (ex: '86.10') → match toutes les sous-classes
      qui commencent par ce préfixe ('86.10Z', '86.10A', '86.10B', etc.)
    """
    code_profil = _normalize_naf(code_profil)
    code_entreprise = _normalize_naf(code_entreprise)
    if not code_profil or not code_entreprise:
        return False
    if code_profil == code_entreprise:
        return True
    if len(code_profil) >= 4 and code_profil[-1].isdigit():
        return code_entreprise.startswith(code_profil)
    return False


def check_naf_coherence(naf: str, naf_attendus: list[str]) -> bool:
    """
    Vérifie si un NAF est cohérent avec la liste attendue.
    Utilise match_naf() pour supporter les codes classe (sans lettre).
    Retourne True si cohérent, False si hors profil.
    """
    if not naf or not naf_attendus:
        return True
    naf = naf.strip()
    for attendu in naf_attendus:
        if match_naf(attendu, naf):
            return True
    return False


# ── Exclusions enseignes nationales (MBT / TG) ──────────────────────────────

EXCLUSIONS_ENSEIGNES = [
    # Discount / Hyper
    "lidl", "aldi", "auchan", "monoprix",
    "picard", "picard surgeles", "picard surgelés",
    # Hyper Carrefour (mais PAS Market/City/Express/Contact)
    "carrefour hypermarche", "carrefour hypermarché",
    # Fast food / Restauration chaîne
    "mcdonald", "mcdonalds", "mcdonald's",
    "burger king", "kfc", "subway", "domino's", "dominos",
    "pizza hut", "quick", "o'tacos", "otacos", "chamas tacos",
    "nabab kebab", "nabab",
    # Stations
    "esso",
    # Informatique
    "apple store",
    # Vêtements
    "zara", "h&m", "uniqlo", "kiabi", "celio", "primark",
    # Sport
    "decathlon",
    # Bijouterie luxe
    "cartier", "van cleef", "pandora", "swarovski",
    # Ameublement / Bricolage
    "ikea", "maisons du monde", "leroy merlin", "castorama",
    "brico depot", "brico dépôt",
    # Immobilier
    "foncia",
    # Boulangerie chaîne
    "paul",
]

_ENSEIGNES_WORD_BOUNDARY = {"kebab"}


def _strip_accents_bio(s: str) -> str:
    return unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode()


def check_enseigne_excluded(name: str) -> str | None:
    """
    Vérifie si un nom d'entreprise matche une enseigne exclue.
    Retourne le mot-clé matché ou None.
    """
    normalized = _strip_accents_bio(name.lower())
    for keyword in EXCLUSIONS_ENSEIGNES:
        norm_kw = _strip_accents_bio(keyword)
        if norm_kw in normalized:
            return keyword
    for keyword in _ENSEIGNES_WORD_BOUNDARY:
        if re.search(rf"\b{keyword}\b", normalized):
            return keyword
    return None
