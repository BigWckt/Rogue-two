# 🎯 Brief — `script_lba_lbb.py`

## Contexte projet

Tu travailles sur un outil de prospection commerciale pour **Skill & You** (organisme de formation en alternance). L'outil croise deux sources de données pour qualifier des leads entreprises :

1. **Pages Jaunes** → base large d'entreprises par secteur (script séparé, pas ton sujet ici)
2. **La Bonne Alternance (LBA) + La Bonne Boîte (LBB)** → signaux de recrutement en alternance (**c'est ton script**)

Le croisement se fait ensuite sur le **SIRET** pour attribuer un niveau de priorité :
- SIRET présent dans LBA (offre active) → **Haute**
- SIRET présent dans LBB (potentiel recrutement) → **Moyenne**
- SIRET uniquement dans Pages Jaunes → **Basse**

---

## Ce que tu dois produire

Un script Python `script_lba_lbb.py` qui :

1. Prend en entrée : **ville**, **codes NAF**, **rayon en km** (défaut 30)
2. Convertit les codes NAF en **codes ROME** (table de correspondance embarquée)
3. Géocode la ville en **latitude/longitude** (API `api-adresse.data.gouv.fr` ou fallback table locale)
4. Interroge l'**API LBA** (entreprises avec offres actives)
5. Interroge l'**API LBB** (entreprises à potentiel de recrutement)
6. Déduplique par **SIRET** (LBA prioritaire sur LBB si doublon)
7. Exporte un fichier **Excel `.xlsx`** intermédiaire + **backup CSV**
8. Affiche une **synthèse console** claire (compteurs LBA / LBB / doublons)

---

## Référence existante — Rogue-two

Tu as déjà un projet `Rogue-two` qui fait exactement ce scraping LBA/LBB. Voici la structure que tu avais utilisée et qui fonctionne :

### Scripts existants
- `scrap_mbt.py` — Scraping Paris 50 km (LBA + LBB)
- `scrap_mbt_villes.py` — Scraping multi-villes (10 villes, 300 lignes max par ville)

### Codes ROME utilisés (MBT — métiers de bouche)
`D1101, D1102, D1103, D1104, D1106, G1603`

### Codes NAF utilisés pour LBB
`10.71A, 10.71B, 10.71C, 10.71D`

### Résultats obtenus
- 3 000 entreprises uniques (déduplication par SIRET)
- Export : `mbt_villes.xlsx` — 1 onglet par ville + onglet Synthèse
- 1 185 contacts avec téléphone déjà présents

**→ Réutilise la logique de ces scripts (appels API, parsing des réponses, gestion des erreurs) en la rendant paramétrique.**

---

## API LBA/LBB — Détails techniques

### Base URL
```
https://labonnealternance.apprentissage.beta.gouv.fr/api/v1
```

### Endpoint `/jobs` (LBA — offres actives)
```
GET /api/v1/jobs?romes=CODE1,CODE2&latitude=LAT&longitude=LON&radius=R&caller=CALLER
```

Retourne un JSON avec plusieurs clés :
- `peJobs` → offres France Travail
- `matchas` → offres LBA natives
- `lbaCompanies` → entreprises LBB (potentiel recrutement)

> **Important** : l'endpoint `/jobs` retourne AUSSI les résultats LBB dans `lbaCompanies`. Tu peux donc tout récupérer en un seul appel par code ROME, ou utiliser l'endpoint `/company` séparément.

### Endpoint `/company` (LBB — potentiel recrutement)
```
GET /api/v1/company?romes=CODE1&latitude=LAT&longitude=LON&radius=R&caller=CALLER
```

### Paramètres
| Paramètre | Type | Obligatoire | Notes |
|-----------|------|-------------|-------|
| `romes` | string | Oui | Codes ROME séparés par virgules (max 20) |
| `latitude` | float | Oui | Latitude en degrés décimaux |
| `longitude` | float | Oui | Longitude en degrés décimaux |
| `radius` | int | Non | Rayon en km (valeurs acceptées : 10, 30, 60, 100 — défaut 30) |
| `caller` | string | Oui | Identifiant du service appelant → utiliser `"skill-and-you-prospection"` |

### Structure des réponses (champs à extraire)
Pour chaque entreprise retournée, extraire :

| Champ cible | Chemin dans le JSON | Notes |
|---|---|---|
| `Nom de l'entreprise` | `company.name` ou `workplace.name` | |
| `SIRET` | `company.siret` ou `workplace.siret` | Clé de déduplication |
| `Adresse` | `place.fullAddress` ou `workplace.location.address` | |
| `Ville` | `place.city` | |
| `Code Postal` | `place.zipCode` | |
| `Code NAF` | `company.naf` ou `workplace.domain.naf` | |
| `Score LBB` | `company.score` ou `hiring_potential` | Seulement LBB |
| `Offres actives` | Compteur d'offres LBA par SIRET | Seulement LBA |
| `Source` | `LBA` / `LBB` / `LBA + LBB` | Selon la provenance |

> ⚠️ La structure JSON a pu évoluer entre les versions de l'API. **Inspecte la réponse réelle** (`print(json.dumps(data, indent=2))`) pour adapter le parsing si les chemins ci-dessus ne correspondent pas exactement.

---

## Table de correspondance NAF → ROME

Embarque une table statique couvrant au minimum les secteurs suivants. Plusieurs codes ROME par NAF est normal — tous doivent être interrogés.

```python
NAF_TO_ROME = {
    # Métiers de bouche
    "10.13A": ["D1103"], "10.13B": ["D1103"],
    "10.71A": ["D1102"], "10.71B": ["D1102"],
    "10.71C": ["D1102", "D1104"], "10.71D": ["D1104"],
    "56.10A": ["G1602", "G1603"], "56.10B": ["G1603"],
    "56.10C": ["G1603"], "56.21Z": ["G1602"],
    # Hôtellerie
    "55.10Z": ["G1703", "G1501"],
    # Commerce alimentaire
    "47.11B": ["D1106"], "47.11C": ["D1106"],
    "47.11D": ["D1504", "D1106"], "47.11F": ["D1504"],
    "47.22Z": ["D1103"], "47.24Z": ["D1106"],
    # Automobile
    "45.11Z": ["D1404"], "45.20A": ["I1604"], "45.20B": ["I1604"],
    # Immobilier
    "68.31Z": ["C1504", "C1501"], "68.32A": ["C1502"],
    # Comptabilité / Juridique
    "69.20Z": ["M1203"], "69.10Z": ["K1903"],
    # Coiffure / Esthétique
    "96.02A": ["D1202"], "96.02B": ["D1208"],
    # Pharmacie
    "47.73Z": ["J1307"],
    # BTP
    "43.21A": ["F1602"], "43.22A": ["F1603"],
    "43.31Z": ["F1611"], "43.34Z": ["F1606"],
    # Informatique
    "62.01Z": ["M1805", "M1802"],
}
```

Si un code NAF fourni en entrée n'est pas dans la table → **log un warning** et continue avec les codes connus. Ne pas planter.

---

## Géocodage

1. **Priorité** : API `https://api-adresse.data.gouv.fr/search/?q=VILLE&limit=1&type=municipality`
2. **Fallback** : table locale des villes principales (Paris, Lyon, Marseille, Toulouse, Bordeaux, Lille, Nice, Nantes, Strasbourg, Montpellier, Rennes, Toulon, Grenoble, Dijon, etc.)

Extraire latitude/longitude depuis `features[0].geometry.coordinates` (attention : l'API retourne `[lon, lat]`, pas `[lat, lon]`).

---

## Déduplication

Logique par SIRET :
- Si un SIRET apparaît dans LBA ET LBB → garder une seule ligne, `Source` = `"LBA + LBB"`, conserver le score LBB et compter les offres LBA
- Si un SIRET apparaît plusieurs fois dans la même source (plusieurs codes ROME) → garder une seule ligne
- Exclure les entreprises sans SIRET

---

## Format de sortie Excel

**Fichier** : `lba_lbb_results_{ville}_{YYYYMMDD}.xlsx`

**Colonnes** (noms exacts — seront réutilisés par le script de comparaison en aval) :

| Colonne | Type |
|---|---|
| `Nom de l'entreprise` | string |
| `SIRET` | string (pas de conversion numérique !) |
| `Code NAF` | string |
| `Adresse` | string |
| `Ville` | string |
| `Code Postal` | string |
| `Source` | string (`LBA` / `LBB` / `LBA + LBB`) |
| `Score LBB` | string ou vide |
| `Offres actives` | int ou vide |
| `Date de collecte` | string `YYYY-MM-DD` |

**Backup CSV** : même nom, extension `.csv`

---

## Interface CLI

```bash
# Usage standard
python script_lba_lbb.py --ville Paris --naf 10.71A 10.71B --rayon 30

# Avec répertoire de sortie
python script_lba_lbb.py --ville Lyon --naf 45.20A --rayon 50 --output ./resultats

# Depuis un fichier JSON de paramètres (pour intégration future avec l'interface web)
python script_lba_lbb.py --config params.json
```

Format du JSON `params.json` :
```json
{
  "ville": "Paris",
  "codes_naf": ["10.71A", "10.71B"],
  "rayon_km": 30
}
```

---

## Robustesse attendue

- **Retry** : 3 tentatives avec backoff exponentiel sur erreur réseau ou HTTP 429/5xx
- **Délai** : 0.8s minimum entre chaque appel API
- **Timeout** : 30s par requête
- **Logs** : affichage console structuré avec emojis (📡 appel, ✅ succès, ❌ erreur, ⚠️ warning)
- **Sauvegarde intermédiaire** : si le script plante en cours de route, les résultats déjà collectés sont sauvés dans un CSV de backup
- **Gestion proxy** : détecter `HTTP_PROXY` / `HTTPS_PROXY` dans l'environnement et les passer à `requests`

---

## Synthèse console attendue en fin d'exécution

```
══════════════════════════════════════════════════
  RÉSULTATS — Paris (rayon 30 km)
  Codes NAF : 10.71A, 10.71B → ROME : D1102, D1104
  
  LBA (offres actives)     : 47 entreprises
  LBB (potentiel)          : 183 entreprises
  Doublons LBA+LBB         : 12 entreprises
  ──────────────────────────────────────────────
  Total unique (par SIRET) : 218 entreprises
  
  📁 Fichier : lba_lbb_results_paris_20260401.xlsx
  📁 Backup  : lba_lbb_results_paris_20260401.csv
══════════════════════════════════════════════════
```

---

## Ce script NE fait PAS

- Pas de scraping Pages Jaunes (script séparé)
- Pas d'enrichissement SIRET (script séparé, pour les entreprises PJ sans SIRET)
- Pas de croisement PJ ↔ LBA/LBB (script de comparaison séparé)
- Pas de push GitHub (étape finale après tous les scripts)

---

## Dépendances

```
requests
pandas
openpyxl
```

---

## Documentation de suivi

À la fin du développement de ce script, ajoute une section dans un fichier `JOURNAL.md` (à créer si inexistant) avec :
- La date
- Ce qui a été fait
- Les problèmes rencontrés (structure API différente de la doc, rate limiting, etc.)
- L'état final (fonctionne / fonctionne partiellement / bloqué)
- Le nombre de résultats obtenus sur un test réel (ville + codes NAF)

Ce journal servira de base pour un rapport non-technique destiné à la direction.
