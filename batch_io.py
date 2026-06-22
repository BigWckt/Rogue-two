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

from rapidfuzz import fuzz

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
                        naf_attendus: str = "",
                        pj_enriched_file: str = "",
                        lba_lbb_file: str = ""):
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
        if pj_enriched_file:
            f.write(f"PJ_ENRICHED_FILE={pj_enriched_file}\n")
        if lba_lbb_file:
            f.write(f"LBA_LBB_FILE={lba_lbb_file}\n")
    return path


def update_current_batch(**kwargs):
    """
    Met à jour des clés dans .current_batch sans écraser les autres.
    Cherche le fichier en remontant (comme read_current_batch).
    """
    current = os.path.abspath(".")
    path = None
    for _ in range(4):
        candidate = os.path.join(current, CURRENT_BATCH_FILE)
        if os.path.isfile(candidate):
            path = candidate
            break
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent

    if not path:
        return None

    data = {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if "=" in line:
                k, v = line.split("=", 1)
                data[k.strip()] = v.strip()

    for k, v in kwargs.items():
        key = k.upper()
        data[key] = str(v)

    with open(path, "w", encoding="utf-8") as f:
        for k, v in data.items():
            f.write(f"{k}={v}\n")

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


# ── Matching multi-critères entreprise (partagé SIRENE + Pappers) ───────────
# Score composite nom (0-50) + CP (+30) + ville (+20). Sector-agnostic :
# ne compare que deux entreprises, sans aucune notion de secteur ou profil.

SEUIL_MATCH = 60   # seuil d'acceptation du meilleur candidat (sur 100)

_NORM_NAME_SUFFIXES = [
    "SARL", "SAS", "SASU", "SA", "EURL", "SCI", "SELARL", "SRL",
    "SCP", "SNC", "STE", "SOCIETE", "ETS", "ETABLISSEMENT", "EI",
]


def norm_name(name: str) -> str:
    """
    Normalise un nom d'entreprise pour la comparaison floue :
    accents supprimés, majuscules, formes juridiques retirées, espaces collés.
    """
    s = unicodedata.normalize("NFKD", str(name)).encode("ASCII", "ignore").decode()
    s = s.upper().strip()
    for suffix in _NORM_NAME_SUFFIXES:
        s = re.sub(rf"\b{suffix}\b", "", s)
    return re.sub(r"\s+", " ", s).strip()


def score_match(nom_recherche: str, nom_candidat: str,
                cp_recherche: str, cp_candidat: str,
                ville_recherche: str, ville_candidat: str) -> tuple[float, str]:
    """
    Score composite de correspondance entreprise (0-100) + détail audit.
    - Similarité nom (norm_name + max ratio/token_sort_ratio) : 0-50 pts
    - CP exact (5 chiffres identiques) : +30 pts
    - Ville fuzzy (≥ 80 après normalisation) : +20 pts
    Retourne (score_total, detail) où detail = "nom 41 + CP 30 + ville 20".
    100% sector-agnostic : ne fait que comparer deux entreprises.
    """
    n1, n2 = norm_name(nom_recherche), norm_name(nom_candidat)
    if n1 and n2:
        name_sim = max(fuzz.ratio(n1, n2), fuzz.token_sort_ratio(n1, n2))
    else:
        name_sim = 0
    name_pts = round(name_sim * 0.5)

    cp_r = str(cp_recherche or "").strip()
    cp_c = str(cp_candidat or "").strip()
    cp_pts = 30 if (cp_r and cp_c and cp_r == cp_c) else 0

    ville_pts = 0
    if ville_recherche and ville_candidat:
        v1 = unicodedata.normalize("NFKD", str(ville_recherche)).encode("ASCII", "ignore").decode().upper().strip()
        v2 = unicodedata.normalize("NFKD", str(ville_candidat)).encode("ASCII", "ignore").decode().upper().strip()
        # max de ratio/partial/token_set : tolère "PARIS"/"PARIS 15",
        # "LILLE"/"LILLE CEDEX", ordre des mots, etc.
        if v1 and v2:
            ville_sim = max(
                fuzz.ratio(v1, v2),
                fuzz.partial_ratio(v1, v2),
                fuzz.token_set_ratio(v1, v2),
            )
            if ville_sim >= 80:
                ville_pts = 20

    total = name_pts + cp_pts + ville_pts
    detail = f"nom {name_pts} + CP {cp_pts} + ville {ville_pts}"
    return total, detail


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


# ── Labels NAF (divisions, classes, sous-classes) ────────────────────────────

NAF_LABELS = {
    # ─── Divisions (2 chiffres) ───
    "10": "Industrie alimentaire",
    "16": "Travail du bois",
    "31": "Fabrication de meubles",
    "33": "Réparation / installation de machines",
    "36": "Captage, traitement, distribution d'eau",
    "41": "Construction de bâtiments",
    "43": "Travaux de construction spécialisés",
    "45": "Commerce / réparation automobile",
    "46": "Commerce de gros",
    "47": "Commerce de détail",
    "55": "Hébergement",
    "56": "Restauration",
    "62": "Programmation / conseil informatique",
    "68": "Activités immobilières",
    "69": "Activités juridiques et comptables",
    "71": "Architecture et ingénierie",
    "74": "Autres activités spécialisées",
    "84": "Administration publique",
    "86": "Activités pour la santé humaine",
    "87": "Hébergement médico-social",
    "88": "Action sociale sans hébergement",
    "95": "Réparation d'ordinateurs et biens personnels",
    "96": "Autres services personnels",
    # ─── Classes (XX.XX) ───
    "10.13": "Préparation de produits à base de viande",
    "10.71": "Fabrication pain / pâtisserie fraîche",
    "16.23": "Fabrication de charpentes et menuiseries",
    "31.01": "Fabrication de meubles de bureau",
    "31.02": "Fabrication de meubles de cuisine",
    "33.15": "Réparation de navires / machines",
    "36.00": "Captage, traitement, distribution d'eau",
    "41.20": "Construction de bâtiments",
    "43.21": "Installation électrique",
    "43.22": "Plomberie, chauffage, climatisation",
    "43.29": "Autres travaux d'installation",
    "43.31": "Travaux de plâtrerie",
    "43.32": "Travaux de menuiserie",
    "43.33": "Travaux de revêtement des sols",
    "43.34": "Peinture et vitrerie",
    "43.99": "Autres travaux de construction spécialisés",
    "45.11": "Commerce de voitures",
    "45.20": "Entretien / réparation automobile",
    "45.31": "Commerce de gros d'équipements auto",
    "45.32": "Commerce de détail d'équipements auto",
    "45.40": "Commerce / réparation de motocycles",
    "46.19": "Intermédiaires du commerce en gros",
    "46.72": "Commerce de gros métaux / minerais",
    "46.73": "Commerce de gros bois / matériaux construction",
    "46.74": "Commerce de gros quincaillerie / fournitures",
    "47.11": "Commerce de détail alimentaire (non spécialisé)",
    "47.19": "Autre commerce de détail non spécialisé",
    "47.21": "Commerce de détail fruits et légumes",
    "47.22": "Commerce de détail viandes",
    "47.23": "Commerce de détail poissons",
    "47.24": "Commerce de détail pain / pâtisserie",
    "47.25": "Commerce de détail boissons",
    "47.29": "Autres commerces alimentaires spécialisés",
    "47.30": "Commerce carburants en magasin",
    "47.41": "Commerce ordinateurs / logiciels",
    "47.42": "Commerce matériel télécom",
    "47.43": "Commerce matériel audio / vidéo",
    "47.51": "Commerce de détail textiles",
    "47.52": "Commerce de détail quincaillerie / peintures",
    "47.53": "Commerce de détail tapis / revêtements",
    "47.54": "Commerce d'appareils électroménagers",
    "47.59": "Commerce meubles / éclairage / autres",
    "47.64": "Commerce articles de sport",
    "47.65": "Commerce jeux et jouets",
    "47.71": "Commerce de détail habillement",
    "47.72": "Commerce de détail chaussures",
    "47.73": "Commerce pharmacie",
    "47.76": "Commerce fleurs / plantes / graines",
    "47.77": "Commerce horlogerie / bijouterie",
    "47.79": "Commerce de détail d'occasion",
    "55.10": "Hôtels et hébergement similaire",
    "56.10": "Restaurants et services de restauration",
    "56.21": "Services des traiteurs",
    "62.01": "Programmation informatique",
    "68.31": "Agences immobilières",
    "68.32": "Administration de biens immobiliers",
    "69.10": "Activités juridiques",
    "69.20": "Comptabilité / audit / conseil fiscal",
    "71.12": "Activités d'ingénierie",
    "74.90": "Autres activités spécialisées n.c.a.",
    "84.11": "Administration publique générale",
    "86.10": "Activités hospitalières",
    "86.21": "Activité des médecins généralistes",
    "86.22": "Activité des médecins spécialistes",
    "86.23": "Pratique dentaire",
    "86.90": "Autres activités pour la santé humaine",
    "87.10": "Hébergement médicalisé",
    "87.30": "Hébergement social pour personnes âgées / handicapées",
    "88.10": "Action sociale sans hébergement pour personnes âgées",
    "88.91": "Action sociale sans hébergement pour jeunes enfants",
    "95.25": "Réparation horlogerie / bijouterie",
    "96.02": "Coiffure et soins de beauté",
    # ─── Sous-classes (XX.XXY) ───
    "10.13A": "Préparation industrielle de viande de boucherie",
    "10.13B": "Charcuterie",
    "10.71A": "Fabrication industrielle de pain",
    "10.71B": "Cuisson de produits de boulangerie",
    "10.71C": "Boulangerie et boulangerie-pâtisserie",
    "10.71D": "Pâtisserie",
    "41.20A": "Construction de maisons individuelles",
    "41.20B": "Construction d'autres bâtiments",
    "43.21A": "Travaux d'installation électrique dans tous locaux",
    "43.21B": "Travaux d'installation électrique sur la voie publique",
    "43.22A": "Travaux d'installation d'eau et de gaz",
    "43.22B": "Travaux d'installation d'équipements thermiques",
    "43.29A": "Travaux d'isolation",
    "43.29B": "Autres travaux d'installation n.c.a.",
    "43.32B": "Travaux de menuiserie métallique et serrurerie",
    "43.99B": "Travaux de montage de structures métalliques",
    "43.99C": "Travaux de maçonnerie générale et gros œuvre",
    "43.99D": "Autres travaux spécialisés de construction",
    "45.20A": "Entretien et réparation de véhicules légers",
    "45.20B": "Entretien et réparation d'autres véhicules",
    "46.19A": "Centrales d'achat non alimentaires",
    "46.73A": "Commerce de gros de bois et matériaux",
    "46.73B": "Commerce de gros d'appareils sanitaires",
    "46.74A": "Commerce de gros de quincaillerie",
    "46.74B": "Commerce de gros de fournitures",
    "47.11B": "Commerce d'alimentation générale",
    "47.11C": "Supérettes",
    "47.11D": "Supermarchés",
    "47.11F": "Hypermarchés",
    "47.19A": "Grands magasins",
    "47.19B": "Autres commerces de détail non spécialisés",
    "47.21Z": "Commerce de détail fruits et légumes (magasin)",
    "47.22Z": "Commerce de détail viandes (magasin)",
    "47.23Z": "Commerce de détail poissons (magasin)",
    "47.24Z": "Commerce de détail pain / pâtisserie / confiserie",
    "47.25Z": "Commerce de détail boissons (magasin)",
    "47.29Z": "Autres commerces alimentaires spécialisés (magasin)",
    "47.52A": "Commerce de détail quincaillerie",
    "47.52B": "Commerce de détail bricolage",
    "47.59A": "Commerce de détail meubles",
    "47.59B": "Commerce de détail autres équipements du foyer",
    "47.72A": "Commerce de détail chaussures",
    "56.10A": "Restauration traditionnelle",
    "56.10B": "Cafétérias et autres libres-services",
    "56.10C": "Restauration rapide",
    "56.21Z": "Services de traiteurs",
    "55.10Z": "Hôtels et hébergement similaire",
    "62.01Z": "Programmation informatique",
    "68.31Z": "Agences immobilières",
    "68.32A": "Administration d'immeubles",
    "69.10Z": "Activités juridiques",
    "69.20Z": "Comptabilité et audit",
    "71.12B": "Ingénierie et études techniques",
    "74.90A": "Activité des économistes de la construction",
    "84.11Z": "Administration publique générale",
    "86.10Z": "Activités hospitalières",
    "86.21Z": "Activité des médecins généralistes",
    "86.22A": "Activités de radiodiagnostic et radiothérapie",
    "86.22B": "Activités chirurgicales",
    "86.22C": "Autres activités des médecins spécialistes",
    "86.23Z": "Pratique dentaire",
    "86.90A": "Ambulances",
    "86.90B": "Laboratoires d'analyses médicales",
    "86.90D": "Activités des infirmiers",
    "86.90E": "Activités des professionnels de la rééducation",
    "86.90F": "Activités de santé humaine non classées ailleurs",
    "87.10A": "Hébergement médicalisé pour personnes âgées",
    "87.10C": "Hébergement médicalisé pour enfants handicapés",
    "87.30A": "Hébergement social pour personnes âgées",
    "87.30B": "Hébergement social pour handicapés mentaux / malades mentaux",
    "88.10A": "Aide à domicile",
    "88.91A": "Accueil de jeunes enfants",
    "95.25Z": "Réparation d'articles d'horlogerie et de bijouterie",
    "96.02A": "Coiffure",
    "96.02B": "Soins de beauté",
}


def get_naf_label(code: str) -> str:
    """
    Retourne le label NAF avec cascade : sous-classe → classe → division.
    Ex: '10.71C' → 'Boulangerie et boulangerie-pâtisserie'
        '10.71'  → 'Fabrication pain / pâtisserie fraîche'
        '10.99X' → 'Industrie alimentaire' (fallback division)
    """
    code = _normalize_naf(code)
    if not code:
        return ""
    if code in NAF_LABELS:
        return NAF_LABELS[code]
    classe = code[:5] if len(code) >= 5 else code
    if classe in NAF_LABELS:
        return NAF_LABELS[classe]
    division = code[:2]
    if division in NAF_LABELS:
        return NAF_LABELS[division]
    return ""


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


# ── Exclusions BTPM (administrations, cultes, cabinets, loueurs, asso) ─────
# Liste évolutive — ajouter des motifs au fil des batches.
# Matching insensible casse + accents (normalisation NFD → ASCII).

EXCLUSIONS_BTPM = {
    "Administrations": [
        "ministere", "prefecture", "sous-prefecture",
        "mairie", "hotel de ville",
        "conseil departemental", "conseil regional", "conseil general",
        "direction departementale", "ddt ", "dreal", "ddpp", "ddcs",
        "tresor public", "dgfip", "centre des finances",
        "service des impots",
        "pole emploi", "france travail",
        "gendarmerie", "commissariat",
        "tribunal", "cour d appel", "cour de cassation",
    ],
    "Organismes sociaux": [
        "cpam", "cnav", "carsat", "urssaf",
    ],
    "Lieux de culte": [
        "mosquee", "eglise", "temple bouddhiste",
        "synagogue", "cathedrale", "basilique", "chapelle",
        "paroisse", "couvent", "monastere", "presbytere",
    ],
    "Professions intellectuelles": [
        "cabinet d architecture", "cabinet d architecte",
        "cabinet d avocats", "cabinet d avocat",
        "cabinet comptable", "expert comptable", "expert-comptable",
        "etude notariale", "office notarial",
        "notaire", "huissier",
        "bureau d etudes", "bureau d etude",
    ],
    "Location véhicules": [
        "location de voiture", "location de vehicule",
        "location automobile",
        "hertz", "europcar", "sixt",
        "enterprise rent", "rent a car", "ucar", "ada location",
    ],
    "Enseignement supérieur": [
        "universite", "rectorat", "inspection academique",
    ],
    "Associations humanitaires": [
        "droits humains", "droits de l homme",
        "humanitaire", "croix rouge", "croix-rouge",
        "amnesty", "amnistie",
        "secours populaire", "restos du coeur",
        "agir ensemble",
        "medecins sans frontieres", "medecins du monde",
    ],
}

_BTPM_WORD_BOUNDARY = {
    "auto ecole": "Mobilité non garage",
    "auto-ecole": "Mobilité non garage",
    "taxi": "Mobilité non garage",
    "vtc": "Mobilité non garage",
    "ambulance": "Mobilité non garage",
    "avis": "Location véhicules",
    "caf": "Organismes sociaux",
}


def check_btpm_excluded(name: str) -> tuple[str, str] | tuple[None, None]:
    """
    Vérifie si un nom d'entreprise matche une exclusion BTPM.
    Retourne (motif, catégorie) ou (None, None).
    """
    n = _strip_accents_bio(name.lower())
    n = re.sub(r"[^a-z0-9 ]", " ", n)
    n = re.sub(r"\s+", " ", n).strip()

    for category, keywords in EXCLUSIONS_BTPM.items():
        for kw in keywords:
            if kw in n:
                return kw, category

    for kw, category in _BTPM_WORD_BOUNDARY.items():
        pattern = re.sub(r"[^a-z0-9 ]", " ", kw)
        if re.search(rf"\b{re.escape(pattern)}\b", n):
            return kw, category

    return None, None
