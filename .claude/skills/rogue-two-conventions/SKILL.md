---
name: rogue-two-conventions
description: >
  Conventions projet Rogue-two — pipeline de prospection entreprises
  (Pages Jaunes → Enrichissement SIRENE/Pappers → LBA/LBB → Croisement).
globs:
  - "*.py"
  - ".current_batch"
  - "JOURNAL.md"
---

## 1 — Environnement

- **Python 3.11+**, exécution sur le poste Windows de l'utilisateur (pas de git CLI installé).
- Dépendances : `playwright`, `pandas`, `rapidfuzz`, `requests`, `openpyxl`.
- **Playwright** : Chromium headful, résolution manuelle du CAPTCHA Cloudflare Turnstile.
  `find_chromium()` doit retourner `None` — jamais de chemin codé en dur.
- **Encodage CSV** : toujours `utf-8-sig` (Excel FR).
- **SIRET** : toujours `str`. Appliquer `df['SIRET'] = df['SIRET'].astype(str)` systématiquement.
  Ne JAMAIS convertir en `int`/`float` (notation scientifique).
- Secteurs : `SB` (services aux bâtiments), `BTPM` (BTP/menuiserie), `MBT` (métallurgie/bois/textile), `TG` (travaux généraux).
  Préfixes définis dans `batch_io.SECTEUR_PREFIXES`.

## 2 — Structure des scripts

| Script | Rôle | Module partagé |
|---|---|---|
| `script_pages_jaunes.py` | Scraping PJ (Playwright) | `batch_io.py` |
| `script_enrichissement.py` | Enrichissement SIRET via API SIRENE | `batch_io.py` |
| `script_enrichissement_pappers.py` | Enrichissement via API Pappers | `batch_io.py` |
| `script_lba_lbb.py` | Prospection LBA + LBB | `batch_io.py` |
| `script_comparaison.py` | Croisement PJ enrichi × LBA/LBB | `batch_io.py` |
| `batch_io.py` | Gestion `.current_batch`, NAF, exclusions | — |
| `matrix_display.py` | Affichage Matrix (ANSI, bannière, Morpheus) | — |
| `inspect_pj_dom.py` | Inspection DOM PJ (diagnostic) | — |

- Pipeline : PJ scraping → Enrichissement (SIRENE ou Pappers) → Comparaison (croisement).
- `.current_batch` : fichier INI-like traçant le batch actif (SECTEUR, BATCH_DIR, DATE, PROFILS, NAF_ATTENDUS, PJ_ENRICHED_FILE, LBA_LBB_FILE).
- Résolution cascade des fichiers : `--arg` CLI > clé `.current_batch` > auto-détection par mot-clé dans le dossier batch.

## 3 — Règles d'édition du code

- **Plan-before-code** : toujours proposer un plan et attendre validation AVANT de modifier du code.
- Ne PAS lancer `git` — l'utilisateur pousse via l'interface web GitHub (drag-and-drop).
  Proposer les commandes git en texte si nécessaire.
- Ne PAS modifier `JOURNAL.md` sans demande explicite.
- Corrections de l'utilisateur basées sur des évidences (logs, scripts qui marchent) : faire confiance, ne pas argumenter.
- Thème Matrix : console verte ANSI, bannière katakana, citations Morpheus. Utiliser `matrix_display.py` pour tout affichage stylisé.
- Imports depuis `batch_io` : `read_current_batch`, `update_current_batch`, `check_naf_coherence`, `get_naf_label`, `check_enseigne_excluded`, `match_naf`, `NAF_LABELS`.
- NAF : normaliser avec `_normalize_naf()` (insère un `.` après les 2 premiers chiffres si absent : `1071C` → `10.71C`).
- `NAF_LABELS` (~110 entrées, 3 niveaux) avec cascade dans `get_naf_label()` : sous-classe → classe → division.
- `EXCLUSIONS_ENSEIGNES` : substring matching pour enseignes nationales, word-boundary pour "kebab".

## 4 — Bugs connus à éviter

- **SIRET en float** : `pandas` convertit les colonnes numériques — toujours forcer `.astype(str)`.
- **NAF sans point** : les fichiers sources omettent parfois le `.` — `_normalize_naf()` corrige.
- **Cloudflare Turnstile** : bloque depuis les IP datacenter. PJ scraping fonctionne UNIQUEMENT depuis une IP résidentielle (poste Windows utilisateur, mode headful).
- **PJ masque les téléphones** sur grandes villes/Paris derrière un bouton "Afficher le N°". Le script utilise une cascade A (lecture statique) → B (clic reveal) → C (page détail en fallback avec `--deep-phone`).
- **Site web PJ** : filtrer les liens internes PJ (`pagesjaunes.fr`, `javascript:`). Cascade de 9 sélecteurs.
- **Catégorie PJ** : guard `len < 200` pour éviter les blocs HTML complets. Cascade de 6 sélecteurs.
- **Playwright version** : ne pas forcer un `executable_path` — laisser `find_chromium()` retourner `None`.

## 5 — Scraping Pages Jaunes

- Constantes clés : `DELAY_AFTER_NAV=12`, `DELAY_SAME_KEYWORD=5`, `CF_TIMEOUT=90`, `SAVE_EVERY=50`.
- Anti-détection : `--disable-blink-features=AutomationControlled`, randomisation des délais.
- `REVEAL_SELECTORS` (12 sélecteurs) pour le bouton d'affichage téléphone.
- `PHONE_READ_SELECTORS` (6 sélecteurs) pour lire le numéro après reveal.
- `REVEAL_THROTTLE_MAX=50` clics avant pause `REVEAL_THROTTLE_PAUSE=10`s.
- Flags CLI : `--debug` (logs détaillés par fiche), `--deep-phone` (visite page détail si téléphone manquant).
- Stats téléphones affichées en synthèse : direct / click / deep / missing.

## 6 — APIs externes

| API | Constante delay | Clé | Notes |
|---|---|---|---|
| SIRENE (recherche-entreprises) | `API_DELAY=1.5` | Publique | `MAX_RETRIES=3`, backoff ×2 |
| Pappers | `API_DELAY=0.5` | `PAPPERS_API_KEY` env var, jamais en dur | Score multi-critères : nom×0.5 + CP + ville. Seuils : HIGH=80, LOW=60 |
| LBA/LBB | — | Bearer `LBA_TOKEN` | `LBA_BASE_URL = "https://api.apprentissage.beta.gouv.fr/api/job/v1"` |

- Fuzzy matching (enrichissement) : `SIMILARITY_THRESHOLD=80` (rapidfuzz).
- Fuzzy dedup (comparaison) : 3 étapes SIRET → téléphone → nom fuzzy.
- Pappers : scoring multi-critères (name×0.5 + CP bonus + city bonus).

## 7 — Livrables attendus

- Fichiers CSV `utf-8-sig` dans le dossier batch (`BATCH_DIR`).
- Nommage : `{SECTEUR}_{batch_number}_{ville}_{source}.csv`.
- `.current_batch` mis à jour après chaque export CSV (`update_current_batch()`).
- Diagnostics NAF verbeux : code + label, cap 10 lignes + "et N autres".
- Config banner au démarrage de chaque script : secteur, profils, NAF attendus (avec labels, cap 8), fichiers résolus.

## 8 — Workflow utilisateur

1. L'utilisateur lance les scripts depuis un terminal Windows (pas WSL).
2. PJ scraping en mode headful — résolution manuelle CAPTCHA Cloudflare.
3. Enrichissement automatique (SIRENE ou Pappers selon le cas).
4. Croisement via `script_comparaison.py`.
5. Les fichiers passent par `.current_batch` pour le chaînage automatique.
6. Pour les commits : l'utilisateur copie les fichiers modifiés via l'interface web GitHub.
7. Toujours proposer un plan avant de coder. Attendre le feu vert.
