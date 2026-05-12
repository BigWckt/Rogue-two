#!/usr/bin/env python3
"""
diag_phone_reveal.py — Diagnostic "Afficher le N°" sur Pages Jaunes
====================================================================
Script standalone pour Windows (headful).
Détermine pourquoi le clic programmatique ne révèle pas le numéro.

Usage :
  python diag_phone_reveal.py

Prérequis : playwright install chromium
Sorties   : /tmp/pj_diag/*.png + *.html + rapport console
"""

import difflib
import os
import re
import time

from playwright.sync_api import sync_playwright

# ── Config ──────────────────────────────────────────────────────────────────
PJ_URL = "https://www.pagesjaunes.fr/annuaire/chercherlespros?quoiqui=restaurant&ou=Lyon"
WARMUP_URL = "https://www.pagesjaunes.fr/annuaire/chercherlespros?quoiqui=fleuriste&ou=Paris+20"
CF_TIMEOUT = 90
OUT_DIR = os.path.join(os.environ.get("TEMP", "/tmp"), "pj_diag")

REVEAL_SELECTORS = [
    "button.btn_tel",
    "button[class*='btn_tel']",
    "[class*='btn_tel']",
    "[class*='click-to-call']",
    "[class*='bi-click-to']",
]

PHONE_READ_SELECTORS = [
    ".bi-fantomas .number-contact",
    "[class*='number-contact']",
    "[class*='phone-number']",
    ".bi-fantomas [class*='phone']",
    ".bi-fantomas [class*='tel']",
    "a[href^='tel:']",
]

BLOC_SELECTOR = "li:has(.bi-content)"


def wait_cf(page, timeout_s=CF_TIMEOUT):
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        title = (page.title() or "").lower()
        if "un instant" not in title and "cloudflare" not in title:
            return True
        time.sleep(5)
    return False


def find_btn_in_bloc(bloc):
    """Cherche le bouton reveal dans le bloc (scope <li>)."""
    for sel in REVEAL_SELECTORS:
        btn = bloc.query_selector(sel)
        if btn:
            return btn, "bloc"
    return None, None


def find_btn_outside_bloc(bloc):
    """Cherche le bouton reveal hors du <li> : sibling, parent, .bi-bloc."""
    try:
        handle = bloc.evaluate_handle(
            "el => el.nextElementSibling?.querySelector('.btn_tel, [class*=\"btn_tel\"]') || "
            "el.closest('[class*=\"bi-bloc\"]')?.querySelector('.btn_tel, [class*=\"btn_tel\"]') || "
            "el.parentElement?.querySelector('.btn_tel, [class*=\"btn_tel\"]')"
        )
        if handle:
            el = handle.as_element()
            if el:
                return el, "sibling/ancestor"
    except Exception:
        pass
    return None, None


def dump_bloc_context(bloc, label, out_dir):
    """Sauvegarde le HTML du bloc + les 2 frères (prev/next)."""
    html = bloc.evaluate(
        """el => {
            let prev = el.previousElementSibling
                ? '<PREV_SIBLING>\\n' + el.previousElementSibling.outerHTML + '\\n</PREV_SIBLING>\\n'
                : '';
            let next = el.nextElementSibling
                ? '<NEXT_SIBLING>\\n' + el.nextElementSibling.outerHTML + '\\n</NEXT_SIBLING>\\n'
                : '';
            return prev + '<BLOC>\\n' + el.outerHTML + '\\n</BLOC>\\n' + next;
        }"""
    )
    path = os.path.join(out_dir, f"bloc_context_{label}.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  📄 Sauvegardé : {path} ({len(html)} chars)")
    return html


def read_phone_after_click(bloc, page):
    """Tente de lire le numéro dans : bloc, btn_tel, modale, page entière."""
    results = {}

    # 1) Sélecteurs classiques dans le bloc
    for sel in PHONE_READ_SELECTORS:
        el = bloc.query_selector(sel)
        if el:
            txt = el.inner_text().strip()
            if txt and len(txt) < 30:
                results[f"bloc → {sel}"] = txt

    # 2) Texte du btn_tel (peut avoir changé)
    btn, _ = find_btn_in_bloc(bloc)
    if not btn:
        btn, _ = find_btn_outside_bloc(bloc)
    if btn:
        try:
            txt = btn.inner_text().strip()
            results["btn_tel.innerText"] = txt
        except Exception:
            pass

    # 3) Modale / overlay (scope page entière)
    modal_selectors = [
        ".modal [class*='phone']",
        ".modal [class*='tel']",
        ".modal a[href^='tel:']",
        "[class*='overlay'] [class*='phone']",
        "[class*='overlay'] [class*='tel']",
        "[class*='popin'] [class*='phone']",
        "[class*='popin'] [class*='tel']",
        "[class*='dialog'] [class*='phone']",
        "[role='dialog'] a[href^='tel:']",
        "[class*='modal'] a[href^='tel:']",
    ]
    for sel in modal_selectors:
        el = page.query_selector(sel)
        if el:
            txt = el.inner_text().strip()
            if txt and len(txt) < 30:
                results[f"page → {sel}"] = txt

    # 4) Tout a[href^='tel:'] dans la page
    tel_links = page.query_selector_all("a[href^='tel:']")
    for i, link in enumerate(tel_links[:5]):
        href = link.get_attribute("href") or ""
        results[f"tel_link_{i}"] = href

    return results


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    print(f"=== Diagnostic PJ Phone Reveal ===")
    print(f"Sorties → {OUT_DIR}\n")

    xhr_log = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-first-run",
                "--no-default-browser-check",
            ],
        )
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/125.0.0.0 Safari/537.36"
            ),
            locale="fr-FR",
        )
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        page = context.new_page()
        page.set_default_timeout(30_000)

        # ── Warmup CF ──
        print("1) Warmup Cloudflare...")
        page.goto(WARMUP_URL, wait_until="domcontentloaded", timeout=30_000)
        time.sleep(12)
        if wait_cf(page):
            print("   ✅ CF passé")
        else:
            print("   ⚠️  CF timeout — on continue quand même")
            input("   Résous le CAPTCHA manuellement puis appuie Entrée...")

        # ── Navigation vers la page restaurant Lyon ──
        print("\n2) Navigation vers la page listing...")
        page.goto(PJ_URL, wait_until="domcontentloaded", timeout=30_000)
        time.sleep(12)
        if not wait_cf(page):
            print("   ⚠️  CF timeout")
            input("   Résous le CAPTCHA puis Entrée...")

        print(f"   Titre : {page.title()}")

        # ── Trouver les blocs ──
        blocs = page.query_selector_all(BLOC_SELECTOR)
        print(f"   Blocs trouvés : {len(blocs)}")
        if not blocs:
            print("   ❌ Aucun bloc — la page est peut-être bloquée CF.")
            page.screenshot(path=os.path.join(OUT_DIR, "no_blocs.png"))
            browser.close()
            return

        # ══════════════════════════════════════════════════════════════════════
        #  FICHE 1 — Clic programmatique
        # ══════════════════════════════════════════════════════════════════════
        bloc = blocs[0]
        nom_el = bloc.query_selector("[class*='bi-denomination'] h3")
        nom = nom_el.inner_text().strip() if nom_el else "(inconnu)"
        print(f"\n3) Fiche 1 : {nom}")

        # Sauvegarder contexte AVANT clic
        html_before = dump_bloc_context(bloc, "avant_clic", OUT_DIR)
        page.screenshot(path=os.path.join(OUT_DIR, "avant_clic.png"))
        print(f"  📸 Screenshot avant_clic.png")

        # Sauvegarder HTML page complète AVANT
        with open(os.path.join(OUT_DIR, "page_avant.html"), "w", encoding="utf-8") as f:
            f.write(page.content())

        # Chercher le bouton
        btn, scope = find_btn_in_bloc(bloc)
        if not btn:
            btn, scope = find_btn_outside_bloc(bloc)
        print(f"  🔍 Bouton : {'trouvé' if btn else 'NON TROUVÉ'} (scope: {scope})")
        if btn:
            try:
                print(f"  🔍 Texte bouton : {btn.inner_text().strip()!r}")
                print(f"  🔍 Tag bouton   : {btn.evaluate('el => el.tagName')}")
                print(f"  🔍 Classes      : {btn.get_attribute('class')}")
            except Exception as e:
                print(f"  ⚠️  Erreur lecture bouton : {e}")

        if not btn:
            print("  ❌ Aucun bouton trouvé — arrêt diagnostic fiche 1.")
        else:
            # Activer le logging XHR
            def on_response(response):
                url = response.url.lower()
                if any(kw in url for kw in ["tel", "phone", "num", "contact", "click", "reveal"]):
                    xhr_log.append({
                        "url": response.url,
                        "status": response.status,
                        "time": time.time(),
                    })
            page.on("response", on_response)

            # Clic programmatique
            print("  🖱️  Clic programmatique...")
            try:
                btn.scroll_into_view_if_needed(timeout=3000)
                time.sleep(0.5)
                btn.click(timeout=5000)
                print("  ✅ Clic effectué")
            except Exception as e:
                print(f"  ❌ Clic échoué : {e}")

            # Attente 5 secondes
            print("  ⏳ Attente 5s...")
            time.sleep(5)

            # Sauvegarder contexte APRÈS clic
            html_after = dump_bloc_context(bloc, "apres_clic_prog", OUT_DIR)
            page.screenshot(path=os.path.join(OUT_DIR, "apres_clic_prog.png"))
            print(f"  📸 Screenshot apres_clic_prog.png")

            # Sauvegarder HTML page complète APRÈS
            with open(os.path.join(OUT_DIR, "page_apres_prog.html"), "w", encoding="utf-8") as f:
                f.write(page.content())

            # XHR log
            print(f"\n  📡 Requêtes XHR interceptées : {len(xhr_log)}")
            for r in xhr_log:
                print(f"     {r['status']} {r['url'][:120]}")

            # Différences DOM
            print("\n  📊 Différences DOM bloc (avant → après clic prog) :")
            before_lines = html_before.splitlines()
            after_lines = html_after.splitlines()
            diff = list(difflib.unified_diff(before_lines, after_lines, n=2))
            if diff:
                for line in diff[:40]:
                    print(f"     {line}")
                if len(diff) > 40:
                    print(f"     ... (+{len(diff) - 40} lignes)")
            else:
                print("     ⚠️  AUCUNE DIFFÉRENCE — le DOM bloc n'a pas changé !")

            # Lecture téléphone post-clic
            print("\n  📞 Lecture téléphone post-clic (tous scopes) :")
            results = read_phone_after_click(bloc, page)
            if results:
                for k, v in results.items():
                    print(f"     {k} = {v!r}")
            else:
                print("     ❌ Aucun numéro trouvé nulle part")

            # Chercher toute modale ou overlay apparue
            print("\n  🪟 Recherche modale/overlay sur la page :")
            for sel in [
                "[class*='modal']:not([style*='display: none'])",
                "[class*='overlay']:not([style*='display: none'])",
                "[class*='popin']:not([style*='display: none'])",
                "[role='dialog']",
                "[class*='lightbox']",
            ]:
                el = page.query_selector(sel)
                if el:
                    vis = el.is_visible()
                    txt = el.inner_text().strip()[:100]
                    print(f"     ✅ {sel} — visible={vis} — texte={txt!r}")

            page.remove_listener("response", on_response)

        # ══════════════════════════════════════════════════════════════════════
        #  FICHE 2 — Clic MANUEL (test H3)
        # ══════════════════════════════════════════════════════════════════════
        if len(blocs) >= 2:
            bloc2 = blocs[1]
            nom2_el = bloc2.query_selector("[class*='bi-denomination'] h3")
            nom2 = nom2_el.inner_text().strip() if nom2_el else "(inconnu)"
            print(f"\n{'='*60}")
            print(f"4) Fiche 2 : {nom2} — TEST CLIC MANUEL (H3)")
            print(f"{'='*60}")

            # Sauvegarder AVANT
            dump_bloc_context(bloc2, "fiche2_avant", OUT_DIR)

            btn2, scope2 = find_btn_in_bloc(bloc2)
            if not btn2:
                btn2, scope2 = find_btn_outside_bloc(bloc2)

            if btn2:
                try:
                    btn2.scroll_into_view_if_needed(timeout=3000)
                except Exception:
                    pass
                print(f"  🔍 Bouton trouvé ({scope2})")
                print(f"  🔍 Texte : {btn2.inner_text().strip()!r}")
            else:
                print("  ⚠️  Pas de bouton trouvé pour fiche 2")

            print("\n  👆 CLIQUE MANUELLEMENT sur le bouton 'Afficher le N°'")
            print("     de la 2e fiche dans le navigateur,")
            print("     puis reviens ici et appuie Entrée.")
            input("     >>> Entrée après clic manuel... ")

            time.sleep(2)

            # Sauvegarder APRÈS clic manuel
            html_manual = dump_bloc_context(bloc2, "fiche2_apres_manuel", OUT_DIR)
            page.screenshot(path=os.path.join(OUT_DIR, "apres_clic_manuel.png"))
            print(f"  📸 Screenshot apres_clic_manuel.png")

            with open(os.path.join(OUT_DIR, "page_apres_manuel.html"), "w", encoding="utf-8") as f:
                f.write(page.content())

            # Lecture téléphone post-clic manuel
            print("\n  📞 Lecture téléphone post-clic manuel :")
            results_manual = read_phone_after_click(bloc2, page)
            if results_manual:
                for k, v in results_manual.items():
                    print(f"     {k} = {v!r}")
            else:
                print("     ❌ Aucun numéro trouvé")

            # Comparaison prog vs manuel
            print("\n  📊 Comparaison clic prog (fiche1) vs clic manuel (fiche2) :")
            if results and results_manual:
                print("     Prog   :", results)
                print("     Manuel :", results_manual)
            elif not results and results_manual:
                print("     ⚠️  Prog = rien, Manuel = résultat → H3 CONFIRMÉE (détection bot)")
            elif results and not results_manual:
                print("     Prog = résultat, Manuel = rien → résultat inattendu")
            else:
                print("     Les deux = rien → problème indépendant du mode de clic")

        # ══════════════════════════════════════════════════════════════════════
        #  SYNTHÈSE
        # ══════════════════════════════════════════════════════════════════════
        print(f"\n{'='*60}")
        print("SYNTHÈSE DIAGNOSTIC")
        print(f"{'='*60}")
        print(f"Fichiers de sortie : {OUT_DIR}")
        print("  avant_clic.png / apres_clic_prog.png / apres_clic_manuel.png")
        print("  bloc_context_*.html / page_avant.html / page_apres_*.html")
        print(f"  XHR interceptées : {len(xhr_log)}")
        print()
        print("Prochaine étape : copier-coller la sortie console ci-dessus")
        print("et envoyer les fichiers .html pertinents pour analyse.")

        input("\nAppuie Entrée pour fermer le navigateur...")
        browser.close()


if __name__ == "__main__":
    main()
