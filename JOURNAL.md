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
