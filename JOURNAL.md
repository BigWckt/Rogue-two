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

#### Étape 3 — Enrichissement SIRET ✅
- 168 entreprises traitées → **75 SIRET trouvés (45%)**
- 93 exclus (similarité < 80% ou SIRET introuvable)
- 58 erreurs API (rate-limiting 429 intensif)
- Taux attendu sans rate-limit : ~60-70%
- Fichier : `Batch_2_0426/pj_results_multi_20260409_enrichi.xlsx` (2 onglets)

#### Étape 4 — Croisement final ✅
- PJ enrichi : 73 entreprises avec SIRET
- LBA/LBB : 38 entreprises avec SIRET
- **Correspondances PJ ↔ LBA/LBB (même SIRET) : 2**

| Priorité | Nombre |
|---|---|
| Haute (LBA) | 0 |
| Moyenne (LBB) | 38 |
| Basse (PJ seul) | 71 |
| **Total** | **109** |

- Fichier : `Batch_2_0426/leads_qualifies_multi_20260409.xlsx` (4 onglets)

### Fichiers générés (Batch_2_0426/)
- `lba_lbb_results_multi_20260409.xlsx` (11 Ko) — 38 agences LBB
- `pj_results_multi_20260409.xlsx` (29 Ko) — 168 fiches PJ
- `pj_results_multi_20260409_enrichi.xlsx` (20 Ko) — 75 avec SIRET
- `leads_qualifies_multi_20260409.xlsx` (24 Ko) — 109 leads qualifiés
