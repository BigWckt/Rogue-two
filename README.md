# Rogue-two — Enrichissement téléphonique MBT

## Vue d'ensemble

Pipeline de scraping et d'enrichissement téléphonique pour les **Métiers de Bouche** (MBT). Le projet collecte des entreprises (boulangeries, boucheries, pâtisseries...) via l'API **La Bonne Alternance** sur 10 villes de France, puis enrichit les contacts sans numéro de téléphone via l'annuaire **118000.fr**.

---

## Architecture du projet

```
Rogue-two/
├── scrap_mbt.py                   # Scraping Paris 50km (LBA + LBB)
├── scrap_mbt_villes.py            # Scraping multi-villes (10 villes)
├── mbt_villes.xlsx                # Données brutes (3 000 entreprises, 10 onglets + Synthèse)
│
├── enrichissement_google.py       # Script principal d'enrichissement 118000.fr
├── enrichissement_checkpoint.csv  # Checkpoint (reprise possible)
├── enrichissement_google_resultats.xlsx  # Résultats détaillés enrichissement
├── mbt_villes_enrichi.xlsx        # mbt_villes avec téléphones injectés
├── contacts_1815_final.xlsx       # Fichier final — 1815 contacts, enrichis en tête
│
├── monitor_enrichissement.sh      # Monitor : check toutes les 5min, relance auto
├── watchdog.sh                    # Watchdog : relance le monitor si arrêté
├── monitor_enrichissement.log     # Logs du monitor + récaps par paliers de 100
├── enrichissement_run.log         # Logs détaillés du script d'enrichissement
│
├── test_pappers.py                # Test exploratoire API Pappers (abandonné)
├── test_pappers_enrichissement.py # Test enrichissement via Pappers
├── enrich_phones.py               # Prototype enrichissement (version antérieure)
└── scrape_alternance.py           # Scraping entreprises alternance Paris
```

---

## Étape 1 — Scraping LBA / LBB

### Scripts
- `scrap_mbt.py` — Scraping autour de Paris (50 km), sans limite de lignes
- `scrap_mbt_villes.py` — Scraping 10 villes, 300 lignes max par ville

### Villes couvertes
| Ville | Rayon |
|-------|-------|
| Paris (75001) | 50 km |
| Toulouse (31000) | 30 km |
| Lille (59000) | 30 km |
| Lyon (69001) | 30 km |
| Marseille (13001) | 30 km |
| Montpellier (34000) | 30 km |
| Nice (06000) | 30 km |
| Bordeaux (33000) | 30 km |
| Perpignan (66000) | 30 km |
| Toulon (83000) | 30 km |

### Codes ROME ciblés (LBA)
`D1101, D1102, D1103, D1104, D1106, G1603`

### Codes NAF ciblés (LBB)
`10.71A, 10.71B, 10.71C, 10.71D`

### Résultat
- **3 000 entreprises** uniques (déduplication par SIRET)
- Export : `mbt_villes.xlsx` — 1 onglet par ville + onglet `Synthèse`
- **1 185 contacts avec téléphone** déjà présents
- **1 815 contacts sans téléphone** → traités en étape 2

---

## Étape 2 — Enrichissement téléphonique

### Pourquoi 118000.fr et pas Google ?

| Critère | Google | 118000.fr |
|---------|--------|-----------|
| Anti-scraping | Très agressif (CAPTCHA) | Modéré |
| Structure HTML | Variable / JS dynamique | Stable et parsable |
| Données | Générales | Annuaire pro vérifié |
| Fiabilité sur 1 800 req. | Impossible sans proxy | Fonctionnel |

### Stratégie de recherche (`enrichissement_google.py`)

Pour chaque entreprise sans téléphone :
1. Requête `https://www.118000.fr/search?who={nom}&where={cp}`
2. Parcours des cartes résultats — ne retient que la carte dont le CP correspond
   - Correspondance exacte du CP d'abord
   - Puis même département (2 premiers chiffres) si aucun exact
3. Si aucun résultat, retente avec le nom de la ville
4. Validation du numéro : format 10 chiffres, exclusion des numéros surtaxés et du numéro propre 118000

### Paramètres
```python
DELAY_BETWEEN    = 1.5s  # + jitter aléatoire 0-1s (anti rate-limit)
CHECKPOINT_EVERY = 50    # sauvegarde checkpoint tous les 50 enregistrements
TEST_MODE        = False  # True = 25 lignes (validation)
```

### Résultats finaux

| Métrique | Valeur |
|----------|--------|
| Contacts à traiter | 1 815 |
| Numéros trouvés | **458** |
| Taux d'enrichissement | **~25%** |
| Source | 118000.fr |
| Durée totale | ~4h (avec relances) |

### Fichier de sortie
`contacts_1815_final.xlsx` — 1 815 lignes, **les 458 enrichis en tête**, colonnes identiques à `mbt_villes.xlsx`.

---

## Étape 3 — Infrastructure de monitoring

### Problème rencontré
Le script `enrichissement_google.py` s'arrêtait régulièrement (~toutes les 100 contacts) à cause de déconnexions réseau ou de rate-limiting 118000.fr. Le checkpointing permet la reprise sans perte.

### Architecture de relance automatique

```
watchdog.sh  (toutes les 60s)
    └── vérifie si monitor_enrichissement.sh tourne
         └── sinon, le relance

monitor_enrichissement.sh  (toutes les 5min)
    ├── vérifie si enrichissement_google.py tourne
    ├── si arrêté ET < 1815 lignes traitées → relance auto
    ├── récap dans le log tous les 100 contacts
    └── si 1815 lignes traitées → génère contacts_1815_final.xlsx + git push
```

### Format des récaps dans `monitor_enrichissement.log`
```
======================================
  RECAP — 1000 contacts traités
  Numéros trouvés : 271
  Taux : 27.1%
  Derniers numéros trouvés :
    BOUCHERIE X -> 0561234567
    BOULANGERIE Y -> 0437891234
======================================
```

---

## Exploration Pappers (abandonnée)

`test_pappers.py` documente pourquoi l'API Pappers a été écartée :
- Téléphone / Email / Site web : **non disponibles** (données INSEE/INPI uniquement)
- Seules les données légales sont accessibles (SIRET, NAF, dirigeants, statut)
- Coût : 4 crédits/SIRET × 1 815 = 7 260 crédits (quota mensuel : 2 000)

---

## Utilisation

### Lancer l'enrichissement
```bash
cd /home/user/Rogue-two
python3 enrichissement_google.py
```

### Relancer avec reprise automatique depuis le checkpoint
```bash
# Le script détecte automatiquement enrichissement_checkpoint.csv
python3 enrichissement_google.py
```

### Lancer le monitoring avec relance auto
```bash
nohup bash monitor_enrichissement.sh > /dev/null 2>&1 &
nohup bash watchdog.sh > /dev/null 2>&1 &
```

### Régénérer le fichier final depuis le checkpoint
```python
import pandas as pd

src = pd.read_excel('mbt_villes.xlsx')
ckpt = pd.read_csv('enrichissement_checkpoint.csv')
ckpt = ckpt.rename(columns={'Source': 'Source_118000'})

sans_tel = src[src['Téléphone'].isna() | (src['Téléphone'].astype(str).str.strip().isin(['','nan']))].copy()
ckpt['SIRET'] = ckpt['SIRET'].astype(str)
sans_tel['SIRET'] = sans_tel['SIRET'].astype(str)

merged = sans_tel.merge(ckpt[['SIRET','Téléphone_trouvé','Source_118000']], on='SIRET', how='left')
mask = merged['Téléphone_trouvé'].notna() & (merged['Téléphone_trouvé'] != '')
merged.loc[mask, 'Téléphone'] = merged.loc[mask, 'Téléphone_trouvé']
merged['_enrichi'] = mask
merged = merged.sort_values('_enrichi', ascending=False).drop(columns=['_enrichi','Téléphone_trouvé','Source_118000'])
merged.to_excel('contacts_1815_final.xlsx', index=False)
```

---

## Dépendances

```
requests
pandas
beautifulsoup4
openpyxl
```

```bash
pip install requests pandas beautifulsoup4 openpyxl
```
