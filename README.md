# Rogue-two

Pipeline de prospection d'entreprises françaises pour la recherche d'alternants.
Croise des sources publiques (Pages Jaunes, SIRENE, Pappers, La Bonne Alternance,
La Bonne Boîte) pour produire des fichiers de leads enrichis et filtrés par profil
métier.

## Pipeline

```
                    ┌─────────────────────────────────────────────────────────┐
                    │                       Phase 1                            │
                    │   script_pages_jaunes.py  →  CSV PJ brut (par ville)     │
                    │   (Playwright headful, anti-CF, anti-CMP, click reveal)  │
                    └─────────────────────────────────────────────────────────┘
                                              ↓
                    ┌─────────────────────────────────────────────────────────┐
                    │                       Phase 2                            │
                    │   script_enrichissement.py        (API SIRENE, gratuit) │
                    │   ou                                                     │
                    │   script_enrichissement_pappers.py (API Pappers, payant)│
                    │   → CSV PJ enrichi (SIRET, NAF officiel)                 │
                    └─────────────────────────────────────────────────────────┘
                                              ↓
                    ┌─────────────────────────────────────────────────────────┐
                    │                       Phase 3                            │
                    │   script_lba_lbb.py  →  CSV LBA/LBB (par ville/profil)  │
                    │   (API La Bonne Alternance + La Bonne Boîte)             │
                    └─────────────────────────────────────────────────────────┘
                                              ↓
                    ┌─────────────────────────────────────────────────────────┐
                    │                       Phase 4                            │
                    │   script_comparaison.py  →  CSV final croisé             │
                    │   (matching SIRET / fuzzy nom / téléphone)               │
                    └─────────────────────────────────────────────────────────┘
```

Le chaînage est automatique via le fichier `.current_batch` : chaque script y écrit
le chemin de son CSV de sortie, le suivant le lit.

## Scripts

| Fichier | Rôle |
|---|---|
| `script_pages_jaunes.py` | Scraping Pages Jaunes (Playwright, headful) |
| `script_enrichissement.py` | Enrichissement SIRET via API SIRENE |
| `script_enrichissement_pappers.py` | Enrichissement via API Pappers |
| `script_lba_lbb.py` | Prospection La Bonne Alternance + La Bonne Boîte |
| `script_comparaison.py` | Croisement PJ enrichi × LBA/LBB |
| `batch_io.py` | Module partagé : `.current_batch`, NAF, exclusions |
| `matrix_display.py` | Affichage console Matrix (ANSI, bannière, Morpheus) |
| `inspect_pj_dom.py` | Diagnostic DOM Pages Jaunes (headful) |
| `diag_phone_reveal.py` | Diagnostic du clic « Afficher le N° » |

## Installation

```bash
pip install playwright pandas rapidfuzz requests openpyxl
playwright install chromium
```

Python 3.11+ requis. Testé sur Windows (le scraping PJ nécessite une IP résidentielle
pour passer Cloudflare Turnstile — pas exécutable depuis WSL/serveur).

## Usage typique

```bash
# 1. Scraping PJ
python script_pages_jaunes.py --ville Lyon --activite restaurant --nb-max 200

# 2. Enrichissement SIRET (auto-détecté via .current_batch)
python script_enrichissement.py

# 3. Prospection LBA/LBB
python script_lba_lbb.py --ville Lyon --profil mbt_restaurants

# 4. Croisement final
python script_comparaison.py
```

Variables d'environnement utiles :
- `PAPPERS_API_KEY` : clé Pappers (obligatoire pour `script_enrichissement_pappers.py`)
- `HTTP_PROXY` / `HTTPS_PROXY` : proxy optionnel pour `script_pages_jaunes.py`

## Secteurs supportés

| Code | Domaine | Profils principaux |
|---|---|---|
| `SB` | Services aux bâtiments | nettoyage, espaces verts, sécurité |
| `BTPM` | BTP & menuiserie | maçonnerie, menuiserie, plomberie |
| `MBT` | Métallurgie / Bois / Textile | métallurgie, boulangerie, restauration |
| `TG` | Travaux généraux | toutes activités hors filtres MBT/BTPM |

25 profils prédéfinis dans `script_lba_lbb.py` (`PROFILS_SKY`).

## Conventions

Voir `.claude/skills/rogue-two-conventions/SKILL.md` pour les conventions
de développement (encodage CSV, normalisation NAF, anti-détection PJ,
gestion des SIRET, etc.).

## Documentation

- `JOURNAL.md` — journal de développement détaillé (problèmes, fixes, validations)
- `.claude/skills/rogue-two-conventions/SKILL.md` — conventions projet
