#!/usr/bin/env python3
"""
matrix_display.py — Affichage theme Matrix
===========================================
Banniere, pluie de caracteres, couleurs ANSI et citations de Morpheus.
Partage entre les 4 scripts de prospection.
"""

import os
import random
import shutil
import sys
import time

# ── ANSI Colors ──────────────────────────────────────────────────────────────

GREEN = "\033[92m"
DIM_GREEN = "\033[2;32m"
RED = "\033[91m"
BOLD = "\033[1m"
RESET = "\033[0m"

# ── Matrix characters ───────────────────────────────────────────────────────

MATRIX_CHARS = "ﾊﾐﾋｰｳｼﾅﾓﾆｻﾜﾂｵﾘｱﾎﾃﾏｹﾒｴｶｷﾑﾕﾗｾﾈｽﾀﾇﾍ0123456789"

# Katakana pleine largeur pour le décodage (meilleur rendu)
KATAKANA = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"

# Fallback ASCII si le terminal ne supporte pas UTF-8
ASCII_FALLBACK = "!@#$%^&*<>{}[]|~0123456789abcdef"

# ── Mode silencieux (global, activé par --quiet dans chaque script) ────────

QUIET = False


def set_quiet(value: bool):
    """Active/désactive le mode silencieux."""
    global QUIET
    QUIET = value


def _decode_chars() -> str:
    """Retourne les caractères de décodage adaptés au terminal."""
    try:
        "ア".encode(sys.stdout.encoding or "utf-8")
        return KATAKANA
    except (UnicodeEncodeError, LookupError):
        return ASCII_FALLBACK


def matrix_decode(text: str, prefix: str = "  [+] ", steps: int = 4, delay: float = 0.06):
    """Affiche 'text' en décodant progressivement des katakana vers les vrais caractères."""
    if QUIET:
        print(f"{prefix}{text}")
        return

    chars = _decode_chars()
    length = len(text)
    for step in range(steps):
        revealed = int(length * (step + 1) / steps)
        line = ""
        for i, char in enumerate(text):
            if char == " ":
                line += " "
            elif i < revealed:
                line += char
            else:
                line += random.choice(chars)
        sys.stdout.write(f"\r{GREEN}{prefix}{RESET}{line}")
        sys.stdout.flush()
        time.sleep(delay)
    sys.stdout.write(f"\r{GREEN}{prefix}{RESET}{text}\n")
    sys.stdout.flush()


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
    """Mini-effet pluie Matrix dans le terminal (quelques lignes)."""
    try:
        width = min(shutil.get_terminal_size().columns, 60)
    except Exception:
        width = 50
    for _ in range(lines):
        line = "".join(random.choice(MATRIX_CHARS) for _ in range(width))
        print(f"{DIM_GREEN}{line}{RESET}")


def matrix_rain_fullscreen(duration: float = 4.0, next_title: str = ""):
    """
    Pluie plein écran pendant `duration` secondes, façon intro Matrix.
    Affiche ensuite le titre du prochain script si fourni.
    Désactivable via variable d'environnement NO_MATRIX_RAIN=1.
    """
    if QUIET or os.environ.get("NO_MATRIX_RAIN", "0") == "1":
        if next_title:
            _print_phase_title(next_title)
        return

    # Activer les ANSI codes sur Windows (PowerShell récent)
    if os.name == "nt":
        os.system("")

    BRIGHT_GREEN = "\033[1;92m"
    chars = _decode_chars() + "0123456789"

    try:
        size = shutil.get_terminal_size()
        cols = size.columns
        rows = max(size.lines - 2, 10)
    except Exception:
        cols, rows = 80, 24

    # Effacer l'écran
    os.system("cls" if os.name == "nt" else "clear")

    density = 0.6
    n_drops = int(cols * density)
    drops = [random.randint(0, rows) for _ in range(n_drops)]
    col_positions = [int(i / density) for i in range(n_drops)]

    start = time.time()
    while time.time() - start < duration:
        # Construire la frame
        frame: list[list[str]] = [[" "] * cols for _ in range(rows)]
        for i, drop_row in enumerate(drops):
            col = col_positions[i]
            if col >= cols:
                continue
            # Tête de goutte (vert clair)
            if 0 <= drop_row < rows:
                frame[drop_row][col] = f"{BRIGHT_GREEN}{random.choice(chars)}{RESET}"
            # Traîne (vert foncé)
            for trail in range(1, 8):
                r = drop_row - trail
                if 0 <= r < rows:
                    frame[r][col] = f"{GREEN}{random.choice(chars)}{RESET}"

        # Afficher la frame en une seule écriture
        sys.stdout.write("\033[H")
        sys.stdout.write("\n".join("".join(row) for row in frame) + "\n")
        sys.stdout.flush()

        # Avancer les gouttes
        drops = [
            (d + 1) if (d + 1) < rows + 8 else random.randint(-10, 0)
            for d in drops
        ]
        time.sleep(0.08)

    # Nettoyer et afficher le titre de la prochaine phase
    os.system("cls" if os.name == "nt" else "clear")
    sys.stdout.write(RESET)
    sys.stdout.flush()
    if next_title:
        _print_phase_title(next_title)


def _print_phase_title(title: str):
    """Affiche un titre de phase encadré après la pluie."""
    bar = "═" * 51
    print()
    print(f"{GREEN}{BOLD}  {bar}{RESET}")
    print(f"{GREEN}{BOLD}  {title:^51s}{RESET}")
    print(f"{GREEN}{BOLD}  {bar}{RESET}")
    print()


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
    name = user_input if user_input else default
    return os.path.basename(name)
