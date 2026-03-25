#!/bin/bash
# Monitoring toutes les 5 minutes — relance auto si script arrêté, génère le fichier final quand terminé

LOG=/home/user/Rogue-two/enrichissement_run.log
CKPT=/home/user/Rogue-two/enrichissement_checkpoint.csv
MONITOR_LOG=/home/user/Rogue-two/monitor_enrichissement.log

echo "[$(date '+%H:%M:%S')] Monitoring démarré" >> "$MONITOR_LOG"

LAST_RECAP=0

while true; do
    RUNNING=$(ps aux | grep "[e]nrichissement_google.py" | wc -l)
    PROGRESS=$(tail -3 "$LOG" 2>/dev/null | grep -oP '\[\s*\d+/1815\]' | tail -1 | tr -d '[]' | xargs)
    ENRICHIS=$(tail -n +2 "$CKPT" 2>/dev/null | awk -F',' '$3 != ""' | wc -l)
    # Ligne actuelle dans le run total (checkpoint = lignes traitées depuis le début)
    TOTAL_TRAITES=$(tail -n +2 "$CKPT" 2>/dev/null | wc -l)
    MILESTONE=$(( (TOTAL_TRAITES / 100) * 100 ))

    # Récap tous les 100 contacts
    if [ "$MILESTONE" -gt "$LAST_RECAP" ] && [ "$MILESTONE" -gt 0 ]; then
        LAST_RECAP=$MILESTONE
        NUMS=$(tail -n +2 "$CKPT" 2>/dev/null | awk -F',' '$3 != ""' | awk -F',' '{print $2, $3}' | head -20)
        echo "" >> "$MONITOR_LOG"
        echo "======================================" >> "$MONITOR_LOG"
        echo "  RECAP — ${MILESTONE} contacts traités" >> "$MONITOR_LOG"
        echo "  Numéros trouvés : ${ENRICHIS}" >> "$MONITOR_LOG"
        echo "  Taux : $(echo "scale=1; $ENRICHIS * 100 / $TOTAL_TRAITES" | bc)%" >> "$MONITOR_LOG"
        echo "  Derniers numéros trouvés :" >> "$MONITOR_LOG"
        tail -n +2 "$CKPT" | awk -F',' '$3 != ""' | awk -F',' '{print "    " $2 " -> " $3}' >> "$MONITOR_LOG"
        echo "======================================" >> "$MONITOR_LOG"
        echo "" >> "$MONITOR_LOG"
    fi

    if [ "$RUNNING" -gt 0 ]; then
        echo "[$(date '+%H:%M:%S')] En cours — $PROGRESS | Enrichis: $ENRICHIS / $TOTAL_TRAITES traités" >> "$MONITOR_LOG"
    elif [ "$TOTAL_TRAITES" -ge 1815 ]; then
        # Toutes les lignes traitées = script vraiment terminé
        echo "[$(date '+%H:%M:%S')] Script TERMINÉ — $PROGRESS | Enrichis: $ENRICHIS" >> "$MONITOR_LOG"
        echo "DONE" >> "$MONITOR_LOG"

        # Générer le fichier final
        python3 << 'PYEOF'
import pandas as pd

src = pd.read_excel('/home/user/Rogue-two/mbt_villes.xlsx')
ckpt = pd.read_csv('/home/user/Rogue-two/enrichissement_checkpoint.csv')
ckpt = ckpt.rename(columns={'Source': 'Source_118000'})

sans_tel = src[src['Téléphone'].isna() | (src['Téléphone'].astype(str).str.strip().isin(['','nan']))].copy()
ckpt['SIRET'] = ckpt['SIRET'].astype(str)
sans_tel['SIRET'] = sans_tel['SIRET'].astype(str)

merged = sans_tel.merge(ckpt[['SIRET','Téléphone_trouvé','Source_118000']], on='SIRET', how='left')
mask = merged['Téléphone_trouvé'].notna() & (merged['Téléphone_trouvé'] != '')
merged.loc[mask, 'Téléphone'] = merged.loc[mask, 'Téléphone_trouvé']
merged['_enrichi'] = mask
merged = merged.sort_values('_enrichi', ascending=False).drop(columns=['_enrichi','Téléphone_trouvé','Source_118000'])
merged.to_excel('/home/user/Rogue-two/contacts_1815_final.xlsx', index=False)
print(f'{mask.sum()} enrichis sur {len(merged)} lignes')
PYEOF

        ENRICHIS_FINAL=$(tail -n +2 "$CKPT" | awk -F',' '$3 != ""' | wc -l)
        cd /home/user/Rogue-two
        git add -f contacts_1815_final.xlsx
        git commit -m "feat: contacts_1815_final.xlsx — ${ENRICHIS_FINAL} enrichis sur 1815

https://claude.ai/code/session_01XzkQisFxLhzSw3n843fcRb"
        git push -u origin claude/test-lbba-GgdLH
        echo "[$(date '+%H:%M:%S')] Fichier final poussé — ${ENRICHIS_FINAL} enrichis" >> "$MONITOR_LOG"
        break
    else
        # Script arrêté mais pas encore fini — relance auto
        echo "[$(date '+%H:%M:%S')] ⚠ ALERTE : script arrêté à $TOTAL_TRAITES/1815 — relance automatique..." >> "$MONITOR_LOG"
        cd /home/user/Rogue-two
        nohup python3 enrichissement_google.py >> enrichissement_run.log 2>&1 &
        echo "[$(date '+%H:%M:%S')] Script relancé (PID $!)" >> "$MONITOR_LOG"
    fi

    sleep 300
done
