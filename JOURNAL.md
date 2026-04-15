# Journal de développement

## 2026-04-02 — script_lba_lbb.py

### Ce qui a été fait
- Création du script `script_lba_lbb.py` : outil CLI de prospection LBA + LBB
- Fonctionnalités implémentées :
  - CLI avec `--ville`, `--naf`, `--rayon`, `--output`, `--config`
  - Table de correspondance NAF → ROME embarquée (20+ codes NAF)
  - Géocodage via `api-adresse.data.gouv.fr` avec fallback table locale
  - Appel API La Bonne Alternance (`/search`) avec token Bearer
  - Déduplication par SIRET (LBA prioritaire sur LBB)
  - Export Excel `.xlsx` + backup CSV
  - Synthèse console avec compteurs

### Problèmes rencontrés
1. **Endpoint API différent de la doc** : l'URL documentée (`labonnealternance.apprentissage.beta.gouv.fr/api/v1/jobs`) retourne 404. L'endpoint fonctionnel est `api.apprentissage.beta.gouv.fr/api/job/v1/search`
2. **Structure de réponse différente** : la doc mentionne `peJobs`, `matchas`, `lbaCompanies` — la réponse réelle contient `jobs` (offres LBA) et `recruiters` (entreprises LBB)
3. **Authentification requise** : l'API nécessite un token Bearer (pas mentionné dans le brief pour cet endpoint)

### État final
**Fonctionne** ✅

### Résultats du test (Paris, NAF 10.71A + 10.71B, rayon 30 km)
- LBA (offres actives) : 12 entreprises uniques
- LBB (potentiel recrutement) : 150 entreprises
- Doublons LBA+LBB : 0
- **Total unique par SIRET : 162 entreprises**
- Fichiers générés : `lba_lbb_results_paris_20260402.xlsx` + `.csv`

---

## 2026-04-02 — script_pages_jaunes.py

### Ce qui a été fait
- Création du script `script_pages_jaunes.py` : scraping Pages Jaunes via Playwright
- Fonctionnalités implémentées :
  - CLI avec `--ville`, `--activite`, `--nb-max`, `--output`, `--config`
  - Détection automatique du binaire Chromium (Playwright)
  - Anti-détection : `navigator.webdriver` override, rotation User-Agent, délais aléatoires
  - Warmup homepage PJ pour cookies Cloudflare
  - Sélecteurs CSS issus de `scraper_prospection.py` (bi-bloc, bi-denomination, etc.)
  - Détection blocage CF : 3 pages vides consécutives → arrêt propre
  - Déduplication sur `(Nom entreprise, Code Postal)`
  - Sauvegarde CSV intermédiaire toutes les 50 fiches
  - Proxy auto-détecté depuis `HTTP_PROXY`/`HTTPS_PROXY`
  - Export Excel `.xlsx` + backup CSV
  - Synthèse console (même style que script 1)
  - Noms de colonnes cohérents avec `script_lba_lbb.py` pour croisement en aval

### Problèmes rencontrés
1. **Chromium non téléchargeable** : `storage.googleapis.com` bloqué en DNS dans l'environnement sandbox — impossible de télécharger la version récente de Chromium
2. **Proxy incompatible Playwright** : l'egress proxy (JWT auth) n'est pas supporté par Chromium → `ERR_INVALID_AUTH_CREDENTIALS`
3. **requests + BS4 insuffisant** : PJ retourne 403 + challenge JS Cloudflare → 0 résultats sans navigateur
4. **Ancien Chromium disponible** : binaire `chromium-1194` présent mais inutilisable à cause du proxy

### État final
**Code fonctionnel, non testable en sandbox** ⚠️
- Le script se lance, détecte correctement le Chromium et le proxy
- La gestion d'erreur fonctionne (3 pages vides → arrêt propre)
- À tester en local sur une machine avec Playwright + Chromium installés et accès direct à PJ

### Résultats du test (Paris, "boulangerie", max 50)
- 0 fiches collectées (blocage proxy/CF attendu en sandbox)
- Le script s'arrête proprement après 3 tentatives échouées

---

## 2026-04-03 — script_pages_jaunes.py v2 (alignement README_scraper_pj.md)

### Ce qui a été modifié
Réécriture du script selon la doc technique `README_scraper_pj.md` (scraper production, 3 329 fiches) :

1. **Warmup CF** : navigation préalable sur une URL PJ bidon (`fleuriste Paris 20`) + attente adaptative (boucle 5s, timeout 90s) jusqu'à ce que le titre ne contienne plus "un instant". Remplace le simple `goto` homepage.
2. **Attente adaptative** : `wait_for_pj_content()` après chaque `page.goto` — vérifie le titre toutes les 5s au lieu d'un `sleep` fixe. Le challenge CF se résout en 8-15s sur IP propre, jusqu'à 30-60s sinon.
3. **Délais entre pages** : 5s entre pages du même mot-clé, 8s entre mots-clés différents (avant : 1-3s aléatoire, trop rapide → rate-limit CF).
4. **Sélecteurs CSS corrigés** (DOM PJ vérifié) :
   - Nom : `[class*='bi-denomination'] h3` (avant : sélecteurs multiples génériques)
   - Adresse : `.bi-address` (avant : `[class*='adresse']`)
   - Téléphone : `.bi-fantomas .number-contact` — div **frère** de `.bi-content` (avant : `[class*='tel']`)
   - Pagination : `a#pagination-next` (avant : `a[rel='next']`)
5. **Validation téléphone** : regex `^0[1-9]\d{8}$` sur chiffres uniquement — écarte les numéros invalides.
6. **Browser** : préfère le vrai Chromium (`chrome-linux/chrome`) au headless shell — requis pour le bypass CF.
7. **`add_init_script` sur le context** (pas la page) pour que le webdriver override s'applique à toutes les pages y compris après rotation UA.

### Problèmes rencontrés
- Aucun nouveau — les mêmes contraintes sandbox (proxy, Chromium) s'appliquent

### État final
**Code aligné sur le scraper production, non testable en sandbox** ⚠️
- À tester en local : `python script_pages_jaunes.py --ville Paris --activite "boulangerie" --nb-max 50`

---

## 2026-04-03 — script_enrichissement.py

### Ce qui a été fait
- Création du script `script_enrichissement.py` : enrichissement SIRET via API recherche-entreprises.api.gouv.fr
- Fonctionnalités implémentées :
  - CLI avec `--input`, `--output`, `--resume` (reprise checkpoint)
  - Recherche SIRET par nom + code postal, fallback par commune, puis nom seul
  - Matching par similarité de chaîne (rapidfuzz, `fuzz.ratio`)
  - Seuil : 80% de similarité minimum pour valider un SIRET
  - Nettoyage des noms (suppression formes juridiques SARL/SAS/etc.)
  - Colonnes ajoutées : `SIRET`, `Code NAF`, `Statut enrichissement`, `Score similarité`
  - Export Excel avec onglet principal (SIRET trouvés) + onglet "Exclus" (audit)
  - Checkpoint CSV toutes les 25 entreprises pour reprise
  - Retry 3x avec backoff exponentiel sur 429/5xx
  - Synthèse console avec compteurs

### Problèmes rencontrés
1. **Rate limiting API** : l'API recherche-entreprises retourne 429 fréquemment en sandbox (plusieurs tentatives par entreprise avec fallbacks commune/nom seul). Délai porté à 1.5s entre appels.
2. **HTTP 400 sur certaines requêtes** : probablement lié aux retries après 429 — corrigé en traitant 400 comme non-retryable.
3. **Noms composés difficiles** : "Boulangerie Maison Landemaine" retourne 0 résultats avec code postal — l'API SIRENE ne référence pas toujours l'enseigne commerciale.

### État final
**Fonctionne** ✅

### Résultats du test (10 boulangeries Paris)
- SIRET trouvés : 5/10 (50%)
  - Du Pain et des Idées → 44066171800021 (95%)
  - Poilâne → 32444503000012 (86%)
  - Boulangerie Bo → 80029869700015 (100%)
  - Maison Kayser → 85152737400025 (100%)
  - Mamiche → 82925376400015 (100%)
- Exclus : 1 (similarité 77% < 80%)
- Erreurs API (429 rate limit) : 4
- Taux attendu en conditions normales (sans rate-limit) : ~60-70%

---

## 2026-04-07 — script_comparaison.py

### Ce qui a été fait
- Création du script `script_comparaison.py` : croisement PJ enrichi × LBA/LBB par SIRET
- Fonctionnalités implémentées :
  - CLI avec `--pj`, `--lba`, `--output`
  - Chargement des deux fichiers Excel, normalisation SIRET
  - Scoring par priorité : Haute (LBA), Moyenne (LBB), Basse (PJ seul)
  - Fusion des données : nom/adresse/tél de PJ + score/offres de LBA/LBB
  - Export Excel 4 onglets : Priorité Haute, Moyenne, Basse, Tous les leads
  - Mise en forme charte Skill & You (en-têtes #1558EE, SIRET texte, freeze panes)
  - Fonctionne avec un seul fichier (--lba seul ou --pj seul)
  - Colonnes HubSpot-ready (13 colonnes nommées exactement)
  - Synthèse console avec compteurs par priorité

### Problèmes rencontrés
- Aucun — l'API n'est pas sollicitée, le script fonctionne entièrement en local

### État final
**Fonctionne** ✅

### Résultats du test
**Test 1 — PJ enrichi (5 entreprises) + LBA/LBB (162 entreprises) :**
- Correspondances PJ ↔ LBA/LBB : 0 (SIRETs différents entre les jeux de test)
- Priorité Haute : 12 | Moyenne : 150 | Basse : 5
- Total : 167 leads qualifiés

**Test 2 — LBA/LBB seul (162 entreprises, sans PJ) :**
- Priorité Haute : 12 | Moyenne : 150 | Basse : 0
- Total : 162 leads qualifiés

---

## 2026-04-09 — Mode multi-villes + fix proxy PJ

### Ce qui a été fait

#### Mode multi-villes (4 scripts)
- **script_lba_lbb.py** : détection `ville` / `villes` dans le JSON config, boucle séquentielle sur chaque ville, onglets Excel par ville + onglet "Consolidé" avec dédup globale par SIRET, colonne "Ville de recherche", gestion rayon 0 (fallback 1 km), résilience par ville (try/except + continue)
- **script_pages_jaunes.py** : même support multi-villes, dédup globale sur (nom, code postal), onglets par ville + Consolidé
- **script_enrichissement.py** : lecture automatique de l'onglet "Consolidé" si présent dans le fichier d'entrée
- **script_comparaison.py** : lecture "Consolidé" pour LBA/LBB et "Entreprises" pour PJ enrichi

#### Fix proxy Chromium (script_pages_jaunes.py)
- **Problème** : `detect_proxy()` lisait `HTTP_PROXY` / `HTTPS_PROXY` et les passait à Chromium comme `{"server": full_url}`. Deux bugs :
  1. Chromium hérite aussi le proxy via env vars → double-proxy → `ERR_INVALID_AUTH_CREDENTIALS`
  2. L'URL proxy complète (avec auth) passée dans `server` au lieu d'être séparée en `server`/`username`/`password`
- **Correction** :
  1. Supprimé `detect_proxy()` — plus d'auto-détection des variables d'environnement
  2. Ajouté `--no-proxy-server` aux args de lancement Chromium par défaut (sortie directe)
  3. Ajouté option CLI `--proxy http://user:pass@host:port` avec parsing correct via `urlparse` (sépare server/username/password pour Playwright)
  4. Ajouté `--ignore-certificate-errors` pour les proxies MITM SSL
  5. Ajouté délai 12s après chaque navigation (`DELAY_AFTER_NAV`)
- **Résultat** : Scraping PJ fonctionnel avec `--proxy "$HTTPS_PROXY"` — 168 fiches collectées

### Résultats — Batch 2 Agences Immobilières

**Paramètres** : villes = [Boulogne-Billancourt, Paris 20], NAF = 68.31Z, rayon = 0 km

#### Étape 1 — Collecte LBA/LBB ✅
| Ville | LBA | LBB | Doublons | Total unique |
|---|---|---|---|---|
| Boulogne-Billancourt | 0 | 29 | 0 | 29 |
| Paris 20 | 0 | 9 | 0 | 9 |
| **Consolidé** | **0** | **38** | **0** | **38** |

- Rayon 0 km → aucun résultat → fallback automatique 1 km → OK
- Fichier : `Batch_2_0426/lba_lbb_results_multi_20260409.xlsx` (3 onglets)

#### Étape 2 — Scraping Pages Jaunes ✅
| Ville | Fiches | Pages | Blocages CF |
|---|---|---|---|
| Boulogne-Billancourt | 105 | 7 | 0 |
| Paris 20 | 68 | 5 | 0 |
| **Consolidé (dédup nom+CP)** | **168** | **12** | **0** |

- Warmup CF OK, 0 blocage Cloudflare
- Lancé avec `--proxy "$HTTPS_PROXY"` + `--ignore-certificate-errors` + `--no-proxy-server` désactivé quand proxy explicite
- Fichier : `Batch_2_0426/pj_results_multi_20260409.xlsx` (3 onglets)

#### Étape 3 — Enrichissement SIRET ✅ (optimisé v2)
Stratégie v2 : **1 seul appel API** (`q=NOM, per_page=10`) + filtrage local en cascade (code_postal → commune → tous), au lieu de 3 appels en cascade côté API.

| Métrique | v1 (cascade 3 appels) | v2 (1 appel + filtrage local) |
|---|---|---|
| SIRET trouvés | 75 (45%) | **101 (60%)** |
| Erreurs API (429) | 58 | **28** |
| Appels API / entreprise | ~2.5 | **~1.1** |

- 168 entreprises traitées → **101 SIRET trouvés (60%)**
- 67 exclus (similarité < 80% ou SIRET introuvable)
- 28 erreurs API (rate-limiting 429, -52% vs v1)
- Support header `Retry-After` + User-Agent explicite
- Fichier : `Batch_2_0426/pj_results_multi_20260409_enrichi.xlsx` (2 onglets)

#### Étape 4 — Croisement final ✅
- PJ enrichi : 98 entreprises avec SIRET
- LBA/LBB : 38 entreprises avec SIRET
- **Correspondances PJ ↔ LBA/LBB (même SIRET) : 2**

| Priorité | Nombre |
|---|---|
| Haute (LBA) | 0 |
| Moyenne (LBB) | 38 |
| Basse (PJ seul) | 96 |
| **Total** | **134** |

- Fichier : `Batch_2_0426/leads_qualifies_multi_20260409.xlsx` (4 onglets)

### Fichiers générés (Batch_2_0426/)
- `lba_lbb_results_multi_20260409.xlsx` (11 Ko) — 38 agences LBB
- `pj_results_multi_20260409.xlsx` (29 Ko) — 168 fiches PJ
- `pj_results_multi_20260409_enrichi.xlsx` (20 Ko) — 101 avec SIRET
- `leads_qualifies_multi_20260409.xlsx` (24 Ko) — 134 leads qualifiés

---

## 2026-04-13 — Refactoring global : CSV only, Matrix theme, fix Chromium

### 1. Format de sortie : CSV uniquement (4 scripts)
- **Suppression complète du format XLSX** sur les 4 scripts
- `script_lba_lbb.py` : supprimé `pd.ExcelWriter`, `.to_excel()`, export CSV seul via `export_csv()`
- `script_pages_jaunes.py` : idem, supprimé l'export multi-onglets Excel
- `script_enrichissement.py` : supprimé `export_results()` Excel (2 onglets Entreprises/Exclus), remplacé par CSV unique avec colonne `Statut enrichissement` pour distinguer trouvés/exclus
- `script_comparaison.py` : supprimé `openpyxl` entièrement (imports, `style_sheet()`, `export_excel()`), remplacé par CSV unique avec colonne `Priorité` (Haute/Moyenne/Basse triées)
- Lecture d'entrée : `script_enrichissement.py` et `script_comparaison.py` lisent désormais des CSV (`pd.read_csv`) au lieu d'Excel (`pd.read_excel`)
- SIRET toujours en string (`df["SIRET"].astype(str)`) dans tous les exports CSV
- Encodage : `utf-8-sig` (BOM) pour compatibilité Excel ouverture directe

### 2. Sortie à la racine du repo par défaut (4 scripts)
- `--output` CLI défaut changé à `.` (racine repo) pour tous les scripts
- La clé `"output"` des fichiers JSON config (`params_lba.json`, `params_pj.json`) est désormais ignorée — seul `--output` CLI fait foi
- Changement dans `load_config()` : `output_dir = args.output` au lieu de `cfg.get("output", ".")`

### 3. Nom de fichier interactif (4 scripts)
- Ajout de `ask_filename(default)` via `input()` avant toute exécution
- L'utilisateur peut accepter le nom par défaut (Entrée) ou taper un nom personnalisé (sans extension `.csv`)
- Noms par défaut : `lba_lbb_results_{ville}_{date}`, `pj_results_{ville}_{date}`, `{base}_enrichi`, `leads_qualifies_{ville}_{date}`

### 4. Affichage thème Matrix (4 scripts)
- Création de `matrix_display.py` : module partagé entre les 4 scripts
  - `matrix_banner()` : ASCII art "WAKE UP, NEO..." + pluie de caractères Matrix
  - `matrix_rain()` : effet pluie de caractères japonais katakana
  - `matrix_step()` / `matrix_ok()` / `matrix_fail()` / `matrix_warn()` : messages `[+]` `[✓]` `[✗]` `[!]` colorés ANSI
  - `matrix_kv()` : affichage clé/valeur avec bullet vert
  - `matrix_section()` : titres de section façon Matrix
  - `morpheus_says()` : citation aléatoire de Morpheus en clôture
  - `ask_filename()` : prompt interactif encadré en vert
- Couleurs : vert (`\033[92m`) pour succès/déco, rouge (`\033[91m`) pour erreurs, bold (`\033[1m`) pour les valeurs
- Les compteurs réels (nombre d'entreprises, ville, taux, SIRET) restent lisibles en blanc/bold
- Libellés adaptés au contexte de chaque script (« Connexion à la Matrice », « Infiltration Pages Jaunes », « Décryptage des identités SIRET », « Fusion des réalités »)

### 5. Fix Chromium cross-platform (script_pages_jaunes.py)
- `find_chromium()` retourne désormais `None` (plus de chemin Linux codé en dur)
- Supprimé le bloc qui plantait avec `sys.exit(1)` si Chromium introuvable
- Supprimé `executable_path` du `launch_kwargs` — Playwright détecte automatiquement le binaire Chromium sur tous les OS (Linux, Windows, Mac)
- Supprimé l'import `glob` devenu inutile

### État final
**Les 4 scripts sont fonctionnels** ✅
- Format : CSV uniquement (plus de dépendance openpyxl pour script_comparaison.py)
- Affichage : thème Matrix cohérent sur les 4 scripts
- Chromium : auto-détecté par Playwright (cross-platform)

---

## 2026-04-14 — Exclusions PJ, profils sectoriels, vérification NAF→ROME

### 1. Exclusions automatiques Pages Jaunes (`script_pages_jaunes.py`)

- Ajout de `EXCLUSIONS_PJ` : dictionnaire de mots-clés d'exclusion groupés par catégorie (Pharmacies, Centres d'hébergement, Foyers jeunes, Centres de planification)
- Filtre insensible à la casse **et aux accents** via `unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode()`
- Appliqué sur le nom de l'entreprise + la catégorie PJ, après scraping et avant déduplication globale
- Logging ventilé par catégorie en console :
  ```
  [+] Exclusions appliquées :
      - Pharmacies              : 12 fiches
      - Centres d'hébergement   : 3 fiches
      Total exclus : 15 fiches
  ```
- Extensible : ajouter des mots-clés à la constante `EXCLUSIONS_PJ` en haut du script

### 2. Profils sectoriels Skill & You (`script_lba_lbb.py`)

- Ajout de `PROFILS_SKY` : 18 profils préconfigurés couvrant 3 secteurs :
  - **SB** (Services / Santé / Beauté) : sb_sante, sb_accueil_enfants, sb_fleuriste, sb_service_personne, sb_coiffure, sb_esthetique
  - **BTPM** (Bâtiment / Mécanique) : btpm_batiment, btpm_electricite, btpm_plomberie, btpm_chauffage_clim, btpm_menuiserie, btpm_meca_carrosserie, btpm_moto
  - **TG** (Tertiaire / Grand commerce) : tg_stations_service, tg_informatique, tg_vetements_chaussures, tg_sport, tg_occasions
- 3 modes de sélection :
  1. **Via JSON** : clé `"profil": "sb_coiffure"` ou `"profils": ["btpm_meca_carrosserie", "btpm_moto"]`
  2. **Via CLI** : `--naf` pour liste libre (comportement inchangé)
  3. **Menu interactif** : si ni `profil` ni `codes_naf` dans le JSON (ou `--ville` sans `--naf`), affiche la liste numérotée des profils et demande un choix
- Fusion multi-profils : si plusieurs profils sélectionnés, les NAF sont fusionnés et dédupliqués
- Format `params.json` étendu : `{"villes": [...], "profil": "sb_coiffure", "rayon_km": 30}`

### 3. Vérification de cohérence NAF → ROME (`script_lba_lbb.py`)

- Après chargement des NAF (depuis profil ou liste libre), vérification que chaque code existe dans `NAF_TO_ROME`
- Warning rouge pour chaque NAF manquant : `⚠️ Code NAF 88.91A absent de la table NAF_TO_ROME — sera ignoré.`
- Affichage du mapping effectif utilisé pour le run :
  ```
  [+] Mapping effectif :
      96.02A → D1202
  ```
- Les codes NAF sans correspondance ROME sont ignorés silencieusement par `resolve_rome_codes()` (déjà le cas), mais le warning explicite permet à l'utilisateur de compléter la table

### État final
**Fonctionnel** ✅

---

## 2026-04-14 — Complétion NAF→ROME + profils MBT

### 1. Table NAF_TO_ROME complétée (~80+ entrées)

La table de mapping `NAF_TO_ROME` dans `script_lba_lbb.py` était incomplète : de nombreux codes NAF référencés par les profils sectoriels (santé, BTP, tertiaire) n'avaient pas de correspondance ROME, ce qui produisait 0 résultat API.

**Secteurs ajoutés :**
- **Santé 86.xx** : 86.10A, 86.10Z, 86.21Z, 86.22A/B/C, 86.23Z, 86.90B → codes ROME J11xx
- **Social 87.xx** : 87.10A/C, 87.30A → K1301/K1302
- **Social 88.xx** : 88.10A, 88.91A → K1301/K1303
- **Admin** : 84.11Z → K2108
- **BTP** (élargi) : 43.21B, 43.22B, 43.29A/B, 43.32B, 43.33Z, 43.99B/C/D, 41.20A/B, 74.90A, 71.12B, 33.15Z, 36.00Z, 16.23Z, 31.01Z, 31.02Z
- **Automobile** (élargi) : 45.31Z, 45.32Z, 45.40Z
- **Commerce alimentaire** (élargi) : 47.11A-F, 47.19A/B, 47.21Z-47.25Z
- **Tertiaire 47.xx** : 47.41Z-47.43Z, 47.51Z, 47.52A/B, 47.53Z, 47.54Z, 47.59A/B, 47.64Z, 47.65Z, 47.71Z, 47.72A, 47.76Z, 47.79Z
- **Fleuriste** : 47.76Z → D1209

### 2. Profils MBT ajoutés à PROFILS_SKY

3 nouveaux profils pour le secteur **MBT (Métiers de Bouche / Tertiaire)** :
- `mbt_boulangeries` : 10.71A, 10.71B, 10.71C, 10.71D
- `mbt_grande_distribution` : 47.11A-F, 47.19A/B, 47.21Z-47.25Z (13 codes NAF)
- `mbt_restaurants` : 56.10A

**Total profils PROFILS_SKY : 21** (6 SB + 7 BTPM + 5 TG + 3 MBT)

### 3. Fichier test `params_lba_sante.json`

Créé pour tester le profil `sb_sante` sur Paris 6→10 + Lyon, rayon 0 km :
```json
{"villes": ["Paris 6", "Paris 7", "Paris 8", "Paris 9", "Paris 10", "Lyon"], "profil": "sb_sante", "rayon_km": 0}
```

### État final
**Fonctionnel** ✅ — Tous les codes NAF des profils existants ont désormais une correspondance ROME dans la table

---

## 2026-04-14 — Extraction des téléphones dans script_lba_lbb.py

### Problème
Le CSV de sortie LBA/LBB ne contenait jamais de numéro de téléphone : la colonne `Téléphone` n'existait pas dans `OUTPUT_COLUMNS` et les fonctions d'extraction (`_extract_lba_company`, `_extract_lbb_company`) ne récupéraient pas le champ.

### Inspection API
Debug temporaire sur la réponse brute. Structure constatée :
- **Offres LBA** (`jobs`) : téléphone dans `apply.phone` (ex: `"0749123569"`)
- **Entreprises LBB** (`recruiters`) : téléphone dans `apply.phone` (souvent `null` pour LBB)

### Corrections apportées

1. **`extract_phone(entry)`** : fonction utilitaire qui cherche le téléphone dans 5 chemins possibles (`apply.phone`, `workplace.contact.phone`, `contact.phone`, `phone`, `telephone`) et retourne le premier non-vide
2. **`normalize_phone(raw)`** : nettoyage (suppression espaces/points/tirets/parenthèses), conversion `+33` → `0`, validation regex `^0[1-9]\d{8}$` (même logique que `validate_phone` dans `script_pages_jaunes.py`)
3. **`_extract_lba_company`** et **`_extract_lbb_company`** : intégration de `extract_phone(item)` dans le dict de retour
4. **`OUTPUT_COLUMNS`** : ajout de `Téléphone` entre `Code Postal` et `Ville de recherche`
5. **Synthèse console** : compteur `Téléphones extraits : N / M entreprises (X%)` dans les deux modes (mono/multi-villes)

### Résultats du test (sb_sante, Paris 6→10 + Lyon, rayon 0 km)
- 320 entreprises collectées (1 LBA + 319 LBB)
- **89 téléphones extraits (28%)**
  - LBA : quasi-systématique (champ `apply.phone` renseigné)
  - LBB : plus rare (~25%) car le champ est souvent `null`
- 0 warning NAF→ROME
- Numéros validés : 10 chiffres français, format cohérent avec `script_pages_jaunes.py`

### État final
**Fonctionnel** ✅

---

## 2026-04-14 — Fusion téléphone par fallback nom+CP (script_comparaison.py)

### Problème
Le croisement PJ ↔ LBA/LBB se fait uniquement par SIRET. Quand une entreprise LBA/LBB n'a pas de téléphone (fréquent pour LBB : `apply.phone` souvent `null`), et que la même entreprise existe dans PJ avec un téléphone mais un SIRET différent (maison-mère vs établissement, ou SIRET non trouvé), le téléphone était perdu.

### Solution : enrichissement en 2 étapes

**Étape 1 — Croisement SIRET (inchangé)** : fusion PJ + LBA/LBB par SIRET exact, scoring priorité Haute/Moyenne/Basse.

**Étape 2 — Fallback téléphone par nom + CP (nouveau)** :
- Pour chaque lead final sans téléphone, cherche dans TOUTES les fiches PJ (y compris celles sans SIRET)
- Filtre par code postal exact (normalisé 5 chiffres, `zfill(5)`)
- Matching fuzzy du nom via `rapidfuzz.fuzz.token_sort_ratio` après normalisation (strip accents, uppercase, suppression formes juridiques SARL/SAS/EURL/SCI/SASU/EI/SELARL)
- Seuil : ≥ 85% de similarité pour valider le match
- Ne jamais écraser un téléphone déjà rempli

### Corrections apportées

1. **Nouveaux imports** : `re`, `unicodedata`, `rapidfuzz.fuzz`
2. **Helpers** : `normalize_cp()` (zfill 5), `strip_accents()` (NFKD), `clean_name()` (strip accents + uppercase + suppression formes juridiques)
3. **`load_pj_all_fiches(filepath)`** : charge TOUTES les fiches PJ avec téléphone (pas seulement celles avec SIRET), pré-calcule `_cp_norm` et `_name_clean` pour chaque fiche
4. **`enrich_phones_fallback(leads, pj_fiches)`** : indexe les fiches PJ par CP, puis pour chaque lead sans téléphone fait un matching fuzzy nom + CP exact
5. **Colonne `Source téléphone`** ajoutée à `OUTPUT_COLUMNS` : valeurs possibles = `PJ direct`, `PJ fallback nom+CP`, `LBA`, `LBB`, ou vide
6. **`merge_and_score()`** : mis à jour pour renseigner `Source téléphone` lors du croisement SIRET (PJ prioritaire sur LBA/LBB pour le téléphone)
7. **Synthèse console** : bloc dédié avec compteurs (déjà présents / ajoutés via PJ / toujours sans / total final)

### Test synthétique (7 entreprises)
- 4 téléphones déjà présents (1 PJ direct par SIRET, 1 LBB, 2 PJ direct Basse)
- **1 téléphone ajouté via fallback** : Centre Médical Lyon Est (LBB sans SIRET match PJ, retrouvé par nom+CP 69003, score >85%)
- 2 toujours sans téléphone (Dr Dupont : score 59% < 85% à cause du suffixe "Chirurgien Dentiste" ; Pharmacie du Parc : CP différent 69002 vs 69007)
- Total : 5/7 (71%)

### État final
**Fonctionnel** ✅

---

## 2026-04-15 — Effet visuel Matrix decode (4 scripts)

### Ce qui a été fait

Ajout d'un effet visuel "Matrix decode" : quand un nom d'entreprise est collecté, les caractères apparaissent d'abord en katakana aléatoires puis se révèlent progressivement en lettres réelles, comme si le système décryptait l'information de la Matrice.

### Implémentation

1. **`matrix_display.py`** — nouvelle fonction `matrix_decode(text, prefix, steps=4, delay=0.06)` :
   - 4 étapes d'animation, chaque étape révèle 25% supplémentaire du texte réel
   - Caractères non révélés remplacés par des katakana aléatoires (pleine largeur)
   - `\r` (retour chariot) pour écraser la ligne à chaque étape
   - Délai total : 4 × 0.06s = 0.24s par ligne
   - **Fallback Windows** : `_decode_chars()` teste si le terminal supporte les katakana via `sys.stdout.encoding`, sinon utilise des caractères ASCII (`!@#$%^&*<>{}[]|~0-9a-f`)
   - **Mode silencieux** : si `QUIET=True`, affiche le texte d'un coup sans animation

2. **Flag `--quiet`** ajouté aux 4 scripts via `argparse` + appel `set_quiet(True)` au début de `main()`

3. **Intégration par script** :
   - `script_lba_lbb.py` : decode pour chaque entreprise LBA (toutes) + les 3 premières LBB par code ROME (préfixes `LBA ▸` / `LBB ▸`)
   - `script_pages_jaunes.py` : decode pour chaque fiche PJ collectée (préfixe `PJ ▸`)
   - `script_enrichissement.py` : decode quand un SIRET est trouvé — format `nom → SIRET (score%)` (préfixe `SIRET ▸`)
   - `script_comparaison.py` : decode uniquement pour les entreprises Priorité Haute (préfixe `★ HAUTE ▸`)

### Contraintes respectées
- Délai total ≤ 0.25s par ligne (4 × 0.06s = 0.24s)
- LBB limité aux 3 premières par ROME pour éviter un overhead excessif (~150 LBB par ROME)
- `--quiet` désactive toutes les animations (mode batch)
- Les compteurs, warnings et récaps restent en texte normal
- Katakana katakana → fallback ASCII si encodage terminal incompatible

### État final
**Fonctionnel** ✅
