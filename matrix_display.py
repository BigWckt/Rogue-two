#!/usr/bin/env python3
"""
matrix_display.py — Affichage theme Matrix
===========================================
Banniere, pluie de caracteres, couleurs ANSI et citations de Morpheus.
Partage entre les 4 scripts de prospection.
"""

import random
import shutil

# ── ANSI Colors ──────────────────────────────────────────────────────────────

GREEN = "\033[92m"
DIM_GREEN = "\033[2;32m"
RED = "\033[91m"
BOLD = "\033[1m"
RESET = "\033[0m"

# ── Matrix characters ───────────────────────────────────────────────────────

MATRIX_CHARS = "ﾊﾐﾋｰｳｼﾅﾓﾆｻﾜﾂｵﾘｱﾎﾃﾏｹﾒｴｶｷﾑﾕﾗｾﾈｽﾀﾇﾍ0123456789"

# ── Morpheus quotes ─────────────────────────────────────────────────────────

MORPHEUS_QUOTES = [
    "I'm trying to free your mind, Neo. But I can only show you the door.",
    "There is a difference between knowing the path and walking the path.",
    "Free your mind.",
    "What is real? How do you define 'real'?",
    "The Matrix is everywhere. It is all around us.",
    "You have to let it all go, Neo. Fear, doubt, and disbelief.",
    "Remember, all I'm offering is the truth. Nothing more.",
    "Welcome to the real world.",
    "The body cannot live without the mind.",
    "I didn't say it would be easy, Neo. I just said it would be the truth.",
]


def matrix_rain(lines: int = 7):
    """Mini-effet pluie Matrix dans le terminal."""
    try:
        width = min(shutil.get_terminal_size().columns, 60)
    except Exception:
        width = 50
    for _ in range(lines):
        line = "".join(random.choice(MATRIX_CHARS) for _ in range(width))
        print(f"{DIM_GREEN}{line}{RESET}")


def matrix_banner(script_title: str):
    """Banniere Matrix d'ouverture."""
    print()
    matrix_rain(5)
    print(f"{GREEN}{BOLD}  > WAKE UP, NEO...{RESET}")
    print(f"{GREEN}  > The Matrix has you...{RESET}")
    print(f"{GREEN}  > Follow the white rabbit.{RESET}")
    print()
    print(f"{GREEN}  ╔{'═' * 52}╗{RESET}")
    print(f"{GREEN}  ║  {BOLD}{script_title:^48s}{RESET}{GREEN}  ║{RESET}")
    print(f"{GREEN}  ╚{'═' * 52}╝{RESET}")
    print()


def matrix_section(title: str):
    """En-tete de section."""
    padding = max(0, 46 - len(title))
    print(f"\n{GREEN}  ── {title} {'─' * padding}{RESET}")


def matrix_kv(label: str, value):
    """Ligne cle/valeur lisible."""
    print(f"    {GREEN}▸{RESET} {label:<30s} : {BOLD}{value}{RESET}")


def matrix_separator():
    """Separateur vert."""
    print(f"    {GREEN}{'─' * 46}{RESET}")


def matrix_step(message: str):
    """Message d'etape facon Matrix : [+] Connexion a la Matrice..."""
    print(f"  {GREEN}[+]{RESET} {message}")


def matrix_ok(message: str):
    """Message de succes."""
    print(f"  {GREEN}[✓]{RESET} {message}")


def matrix_fail(message: str):
    """Message d'erreur."""
    print(f"  {RED}[✗]{RESET} {message}")


def matrix_warn(message: str):
    """Message d'avertissement."""
    print(f"  {RED}[!]{RESET} {message}")


def morpheus_says():
    """Citation aleatoire de Morpheus pour cloturer le script."""
    quote = random.choice(MORPHEUS_QUOTES)
    print()
    print(f"{GREEN}{BOLD}  ┌{'─' * 52}┐{RESET}")
    print(f"{GREEN}{BOLD}  │  MORPHEUS says:{RESET}")
    print(f"{GREEN}  │  \"{quote}\"{RESET}")
    print(f"{GREEN}{BOLD}  └{'─' * 52}┘{RESET}")
    matrix_rain(3)


def ask_filename(default: str) -> str:
    """Demande interactivement le nom du fichier de sortie (sans extension)."""
    print(f"{GREEN}  ┌{'─' * 52}┐{RESET}")
    print(f"{GREEN}  │  Nom du fichier de sortie (sans extension .csv)    │{RESET}")
    print(f"{GREEN}  └{'─' * 52}┘{RESET}")
    user_input = input(f"    [{default}] > ").strip()
    return user_input if user_input else default
