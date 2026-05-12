#!/usr/bin/env python3
"""
inspect_pj_dom.py — Inspection DOM Pages Jaunes
================================================
Script autonome à lancer sur ta machine Windows (avec Playwright installé)
pour identifier les sélecteurs exacts du bouton de révélation téléphone,
du site web, et de la catégorie.

Usage :
  python inspect_pj_dom.py
  python inspect_pj_dom.py --url "https://www.pagesjaunes.fr/annuaire/chercherlespros?quoiqui=restaurant&ou=Lyon"
  python inspect_pj_dom.py --headful   # pour voir le navigateur

Sortie : fichiers HTML + screenshots dans ./pj_inspect/
"""

import argparse
import json
import os
import random
import re
import time

from playwright.sync_api import sync_playwright

PJ_BASE = "https://www.pagesjaunes.fr"
WARMUP_URL = f"{PJ_BASE}/annuaire/chercherlespros?quoiqui=fleuriste&ou=Paris+20"
DEFAULT_TARGET = f"{PJ_BASE}/annuaire/chercherlespros?quoiqui=agence+immobiliere&ou=Paris+3eme"

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
]

PHONE_SELECTORS = [
    ".bi-fantomas .number-contact",
    "[class*='number-contact']",
    "[class*='phone']",
    "[class*='tel']",
    "a[href^='tel:']",
    "[class*='click-to-call']",
    "[class*='bi-click-to']",
    "button[class*='num']",
    "a[class*='num']",
    "[data-pjlb*='tel']",
    "[data-pjlb*='click']",
    ".bi-fantomas button",
    ".bi-fantomas a",
    ".bi-fantomas",
]

WEBSITE_SELECTORS = [
    "a[class*='site']",
    "a[class*='web']",
    "a.bi-website",
    "a[href*='http'][class*='url']",
    "[class*='website']",
    "[class*='site-internet']",
    "a[data-pjlb*='site']",
    "a[data-pjlb*='web']",
]

CATEGORY_SELECTORS = [
    ".bi-activite",
    "[class*='activite']",
    "[class*='rubrique']",
    "[class*='category']",
    "[class*='metier']",
]

BLOC_SELECTORS = [
    "li:has(.bi-content)",
    "article.bi-bloc, div[class*='bi-bloc'], li[class*='bi-bloc']",
    "[data-bi-name]",
    ".bi-listItem",
    "li.bi",
    "ul.bi-list > li",
    "#listResults li",
]


def wait_cf(page, timeout_s=90):
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        t = (page.title() or "").lower()
        if "un instant" not in t and "cloudflare" not in t and "just a moment" not in t:
            return True
        time.sleep(5)
    return False


def inspect_bloc(bloc, index, out_dir):
    print(f"\n{'='*60}")
    print(f"  BLOC {index}")
    print(f"{'='*60}")

    outer = bloc.evaluate("el => el.outerHTML")
    path = os.path.join(out_dir, f"bloc_{index}.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(outer)
    print(f"  HTML saved: {path} ({len(outer)} chars)")

    # Name
    for nsel in ["[class*='bi-denomination'] h3", "h3", "[class*='denomination']"]:
        nel = bloc.query_selector(nsel)
        if nel:
            print(f"  NOM ({nsel}): {nel.inner_text().strip()[:60]}")
            break

    # Phone selectors
    print(f"\n  --- PHONE SELECTORS (AVANT clic) ---")
    for psel in PHONE_SELECTORS:
        try:
            el = bloc.query_selector(psel)
            if el:
                text = el.inner_text().strip()[:100]
                href = el.get_attribute("href") or ""
                cls = el.get_attribute("class") or ""
                tag = el.evaluate("el => el.tagName")
                dpjlb = el.get_attribute("data-pjlb") or ""
                print(f"    HIT  '{psel}':")
                print(f"         <{tag}> class='{cls[:80]}'")
                print(f"         text='{text[:60]}'")
                if href:
                    print(f"         href='{href[:80]}'")
                if dpjlb:
                    print(f"         data-pjlb='{dpjlb[:60]}'")
        except Exception as e:
            print(f"    ERR  '{psel}': {e}")

    # Try to find ANY clickable element in the phone area
    print(f"\n  --- BUTTONS/LINKS in .bi-fantomas ---")
    fantomas = bloc.query_selector(".bi-fantomas")
    if fantomas:
        all_children = fantomas.evaluate("""el => {
            const result = [];
            const walk = (node, depth) => {
                if (node.nodeType === 1) {
                    result.push({
                        tag: node.tagName,
                        cls: String(node.className || '').substring(0, 100),
                        text: (node.textContent || '').trim().substring(0, 80),
                        href: (node.getAttribute('href') || '').substring(0, 80),
                        dpjlb: node.getAttribute('data-pjlb') || '',
                        onclick: node.getAttribute('onclick') || '',
                        depth: depth,
                    });
                    for (const child of node.children) walk(child, depth + 1);
                }
            };
            walk(el, 0);
            return result;
        }""")
        for c in all_children:
            indent = "    " + "  " * c['depth']
            extras = ""
            if c['href']:
                extras += f" href='{c['href']}'"
            if c['dpjlb']:
                extras += f" data-pjlb='{c['dpjlb']}'"
            if c['onclick']:
                extras += f" onclick='{c['onclick'][:40]}'"
            print(f"{indent}<{c['tag']}> .{c['cls']}{extras}")
            if c['text']:
                print(f"{indent}  text: '{c['text'][:60]}'")
    else:
        print("    .bi-fantomas NOT FOUND — checking full bloc for phone-related elements")
        all_els = bloc.evaluate("""el => {
            const result = [];
            const walk = (node, depth) => {
                if (node.nodeType === 1) {
                    const cls = String(node.className || '').toLowerCase();
                    const text = (node.textContent || '').trim().toLowerCase();
                    const href = node.getAttribute('href') || '';
                    if (cls.includes('phone') || cls.includes('tel') || cls.includes('num') ||
                        cls.includes('call') || cls.includes('fantom') ||
                        text.includes('afficher') || text.includes('numéro') ||
                        text.includes('tel') || href.startsWith('tel:')) {
                        result.push({
                            tag: node.tagName,
                            cls: String(node.className || '').substring(0, 100),
                            text: (node.textContent || '').trim().substring(0, 80),
                            href: href.substring(0, 80),
                            depth: depth,
                        });
                    }
                    for (const child of node.children) walk(child, depth + 1);
                }
            };
            walk(el, 0);
            return result;
        }""")
        for c in all_els:
            print(f"    <{c['tag']}> .{c['cls'][:60]} text='{c['text'][:40]}'")

    # Website selectors
    print(f"\n  --- WEBSITE SELECTORS ---")
    for wsel in WEBSITE_SELECTORS:
        try:
            el = bloc.query_selector(wsel)
            if el:
                href = el.get_attribute("href") or ""
                cls = el.get_attribute("class") or ""
                print(f"    HIT  '{wsel}': class='{cls[:60]}' href='{href[:80]}'")
        except:
            pass

    # Category selectors
    print(f"\n  --- CATEGORY SELECTORS ---")
    for csel in CATEGORY_SELECTORS:
        try:
            el = bloc.query_selector(csel)
            if el:
                text = el.inner_text().strip()[:60]
                cls = el.get_attribute("class") or ""
                print(f"    HIT  '{csel}': class='{cls[:60]}' text='{text}'")
        except:
            pass

    # Screenshot
    try:
        bloc.screenshot(path=os.path.join(out_dir, f"bloc_{index}_before_click.png"))
    except:
        pass

    return outer


def try_click_and_reread(bloc, index, out_dir):
    """Attempt to click every button-like element in .bi-fantomas and check what changes."""
    print(f"\n  --- TENTATIVES DE CLIC (bloc {index}) ---")
    fantomas = bloc.query_selector(".bi-fantomas")
    target = fantomas or bloc

    clickables = target.query_selector_all("button, a, [role='button'], [onclick]")
    if not clickables:
        print("    Aucun élément cliquable trouvé")
        return

    for j, btn in enumerate(clickables):
        try:
            tag = btn.evaluate("el => el.tagName")
            cls = btn.get_attribute("class") or ""
            text = btn.inner_text().strip()[:40]
            print(f"\n    Clic #{j+1}: <{tag}> .{cls[:60]} text='{text}'")

            btn.click(timeout=3000)
            time.sleep(2)

            # Re-read phone selectors
            for psel in PHONE_SELECTORS[:6]:
                el = bloc.query_selector(psel)
                if el:
                    new_text = el.inner_text().strip()[:60]
                    href = el.get_attribute("href") or ""
                    print(f"      APRÈS CLIC '{psel}': text='{new_text}' href='{href[:60]}'")

            # Screenshot after click
            try:
                bloc.screenshot(path=os.path.join(out_dir, f"bloc_{index}_after_click_{j+1}.png"))
                print(f"      Screenshot: bloc_{index}_after_click_{j+1}.png")
            except:
                pass

        except Exception as e:
            print(f"      ERR clic #{j+1}: {e}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default=DEFAULT_TARGET)
    parser.add_argument("--headful", action="store_true")
    parser.add_argument("--nb-blocs", type=int, default=3)
    args = parser.parse_args()

    out_dir = os.path.join(".", "pj_inspect")
    os.makedirs(out_dir, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=not args.headful,
            args=[
                "--no-sandbox",
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
            ],
        )
        ctx = browser.new_context(
            user_agent=random.choice(USER_AGENTS),
            locale="fr-FR",
        )
        ctx.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        page = ctx.new_page()
        page.set_default_timeout(30000)

        # Warmup CF
        print("=== WARMUP CF ===")
        page.goto(WARMUP_URL, wait_until="domcontentloaded", timeout=30000)
        time.sleep(12)
        if wait_cf(page):
            print(f"CF passé. Title: '{page.title()}'")
        else:
            print(f"CF TIMEOUT. Title: '{page.title()}'")
            page.screenshot(path=os.path.join(out_dir, "cf_timeout.png"))
            print("Si tu es en headful, résous le CAPTCHA manuellement puis appuie sur Entrée ici.")
            input("Appuie sur Entrée pour continuer...")
            if not wait_cf(page, timeout_s=30):
                print("Toujours bloqué par CF. Abandon.")
                browser.close()
                return

        # Target page
        print(f"\n=== TARGET: {args.url} ===")
        page.goto(args.url, wait_until="domcontentloaded", timeout=30000)
        time.sleep(12)
        if not wait_cf(page, timeout_s=60):
            print(f"CF TIMEOUT sur la page cible. Title: '{page.title()}'")
            page.screenshot(path=os.path.join(out_dir, "target_cf_timeout.png"))
            browser.close()
            return

        print(f"Title: '{page.title()}'")
        page.screenshot(path=os.path.join(out_dir, "listing_full.png"), full_page=False)

        # Find blocs
        blocs = []
        for sel in BLOC_SELECTORS:
            blocs = page.query_selector_all(sel)
            if blocs:
                print(f"\nBlocs trouvés avec '{sel}': {len(blocs)}")
                break

        if not blocs:
            print("\nAUCUN BLOC TROUVÉ. Dump de la page :")
            html = page.content()
            with open(os.path.join(out_dir, "full_page.html"), "w", encoding="utf-8") as f:
                f.write(html)
            print(f"  Page complète sauvée dans {out_dir}/full_page.html")
            browser.close()
            return

        # Inspect each bloc
        for i, bloc in enumerate(blocs[:args.nb_blocs], 1):
            inspect_bloc(bloc, i, out_dir)
            try_click_and_reread(bloc, i, out_dir)

        # Also inspect a detail page for the first fiche
        print(f"\n{'='*60}")
        print(f"  INSPECTION PAGE DÉTAIL (1ère fiche)")
        print(f"{'='*60}")
        first_bloc = blocs[0]
        link = first_bloc.query_selector("a[href*='/pros/']")
        if not link:
            link = first_bloc.query_selector("[class*='bi-denomination']")
        if link:
            href = link.get_attribute("href") or ""
            if href.startswith("/"):
                href = PJ_BASE + href
            if href.startswith("http"):
                print(f"  Navigation vers: {href}")
                page.goto(href, wait_until="domcontentloaded", timeout=30000)
                time.sleep(8)
                if wait_cf(page, timeout_s=30):
                    page.screenshot(path=os.path.join(out_dir, "detail_page.png"), full_page=False)
                    html = page.content()
                    with open(os.path.join(out_dir, "detail_page.html"), "w", encoding="utf-8") as f:
                        f.write(html)
                    print(f"  Page détail sauvée dans {out_dir}/detail_page.html")

                    # Search for phone on detail page
                    print(f"\n  --- PHONE sur page détail ---")
                    for psel in PHONE_SELECTORS + ["[itemprop='telephone']", ".coord-numero"]:
                        try:
                            el = page.query_selector(psel)
                            if el:
                                text = el.inner_text().strip()[:60]
                                href_attr = el.get_attribute("href") or ""
                                cls = el.get_attribute("class") or ""
                                print(f"    HIT  '{psel}': text='{text}' href='{href_attr[:60]}' class='{cls[:60]}'")
                        except:
                            pass
                else:
                    print("  CF timeout sur page détail")
            else:
                print(f"  Pas de lien exploitable: href='{href}'")
        else:
            print("  Pas de lien trouvé dans le premier bloc")

        browser.close()

    print(f"\n=== TERMINÉ ===")
    print(f"Fichiers dans ./{out_dir}/")
    print("Envoie-moi le contenu de bloc_1.html + bloc_2.html + les screenshots")
    print("pour que je puisse identifier les sélecteurs exacts.")


if __name__ == "__main__":
    main()
