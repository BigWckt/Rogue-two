#!/bin/bash
# Watchdog — relance monitor si arrêté
while true; do
    if ! ps aux | grep "[m]onitor_enrichissement" > /dev/null; then
        echo "[$(date '+%H:%M:%S')] Watchdog: monitor relancé" >> /home/user/Rogue-two/monitor_enrichissement.log
        nohup bash /home/user/Rogue-two/monitor_enrichissement.sh > /dev/null 2>&1 &
    fi
    sleep 60
done
