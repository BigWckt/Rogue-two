#!/usr/bin/env python3
"""
batch_io.py — Gestion du batch actif (.current_batch)
======================================================
Module partagé entre les 4 scripts de prospection.
Écrit et lit le fichier .current_batch qui trace le dossier actif.
"""

import os
import re

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


def check_naf_coherence(naf: str, naf_attendus: list[str]) -> bool:
    """
    Vérifie si un NAF est cohérent avec la liste attendue.
    Match exact d'abord, puis match sur préfixe 4 caractères (avant la lettre).
    Retourne True si cohérent, False si hors profil.
    """
    if not naf or not naf_attendus:
        return True
    naf = naf.strip()
    if naf in naf_attendus:
        return True
    naf_prefix = naf[:4] if len(naf) >= 4 else naf
    for attendu in naf_attendus:
        if attendu[:4] == naf_prefix:
            return True
    return False
