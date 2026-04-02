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
