#!/usr/bin/env python3
"""
Probe d'une page produit PrestaShop spécifique (URL directe).

Pour aller plus loin que probe_prestashop.py qui ne fait qu'effleurer une
fiche produit, ce script :
  - Récupère le HTML complet
  - Cherche TOUTES les sources possibles de données (JSON-LD, JS globals,
    data-attributes, sélecteurs HTML, microdata)
  - Pour les variantes : tente 6 méthodes différentes d'extraction
  - Dumpe le HTML brut dans un fichier pour inspection manuelle

Usage :
    python3 probe_prestashop_page.py https://promolinge.com/.../687-1125-...html
"""

from __future__ import annotations

import gzip
import html
import json
import re
import ssl
import sys
import urllib.request
import urllib.error


UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def _ssl_ctx():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def fetch(url: str) -> tuple[int, str]:
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "gzip, identity",
    })
    try:
        with urllib.request.urlopen(req, timeout=30, context=_ssl_ctx()) as resp:
            raw = resp.read()
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                try:
                    raw = gzip.decompress(raw)
                except Exception:
                    pass
            return resp.status, raw.decode("utf-8", errors="replace")
    except Exception as e:
        print(f"[err] {type(e).__name__}: {e}")
        return 0, ""


def section(title: str) -> None:
    print(f"\n{'─' * 78}\n  {title}\n{'─' * 78}")


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage : python3 probe_prestashop_page.py <url_produit>")
        return 2
    url = sys.argv[1]
    print("=" * 78)
    print(f"  Inspection : {url}")
    print("=" * 78)

    s, body = fetch(url)
    if s != 200 or not body:
        print(f"  HTTP {s} - abort")
        return 1
    print(f"  HTTP {s}, {len(body)} chars")

    # Sauvegarde le HTML brut pour inspection ultérieure
    fn = re.sub(r"[^a-zA-Z0-9._-]", "_", url.split("//")[-1])[:100] + ".html"
    with open(fn, "w", encoding="utf-8") as f:
        f.write(body)
    print(f"  HTML brut sauvegardé : {fn}")

    # ============================================================== 1. JSON-LD
    section("1. JSON-LD (tous les blocs)")
    blocks = re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        body, flags=re.DOTALL | re.IGNORECASE,
    )
    print(f"  {len(blocks)} bloc(s) trouvé(s)")
    for i, blk in enumerate(blocks):
        try:
            data = json.loads(blk.strip())
        except json.JSONDecodeError as e:
            print(f"  [{i}] invalid JSON: {e}")
            continue
        types_found = []
        candidates = data if isinstance(data, list) else (
            data.get("@graph", []) if isinstance(data, dict) and "@graph" in data else [data]
        )
        for c in candidates:
            if isinstance(c, dict):
                t = c.get("@type")
                if t:
                    types_found.append(str(t))
        print(f"  [{i}] @type(s) : {types_found}")
        # Si Product, on dump complet
        for c in candidates:
            if isinstance(c, dict):
                t = c.get("@type")
                if t == "Product" or (isinstance(t, list) and "Product" in t):
                    print(f"      → Product complet :")
                    print(f"      {json.dumps(c, ensure_ascii=False, indent=2)[:1500]}")

    # ============================================================== 2. JS globals
    section("2. Variables JavaScript globales")
    js_patterns = [
        (r'var\s+prestashop\s*=\s*({[\s\S]+?})\s*;', "var prestashop = {...}"),
        (r'window\.prestashop\s*=\s*({[\s\S]+?})\s*;', "window.prestashop = {...}"),
        (r'var\s+combinations\s*=\s*({[\s\S]+?})\s*;', "var combinations = {...}"),
        (r'var\s+combinationsFromController\s*=\s*({[\s\S]+?})\s*;', "var combinationsFromController = {...}"),
        (r'"combinations"\s*:\s*({[\s\S]+?})\s*,\s*"', "key combinations dans JSON inline"),
        (r'prestashop\.product\s*=\s*({[\s\S]+?})\s*;', "prestashop.product = {...}"),
        (r'data-product\s*=\s*["\']({[^"\']+})["\']', "data-product attribute (JSON encodé)"),
        (r'data-combinations\s*=\s*["\']([^"\']+)["\']', "data-combinations attribute"),
    ]
    for pat, label in js_patterns:
        m = re.search(pat, body)
        if m:
            raw = m.group(1)
            # Si encodé HTML
            if "&quot;" in raw or "&#" in raw:
                raw = html.unescape(raw)
            preview = raw[:400]
            print(f"  [OK] {label}")
            print(f"       len={len(m.group(1))} preview: {preview!r}")
            # Tentative de parse JSON
            try:
                parsed = json.loads(raw[:50000])  # cap pour ne pas exploser
                if isinstance(parsed, dict):
                    print(f"       → dict, {len(parsed)} clés : {list(parsed.keys())[:15]}")
                    # Cherche les sous-objets combinations / variants
                    for k in ("combinations", "attributes", "variants", "product"):
                        if k in parsed:
                            v = parsed[k]
                            if isinstance(v, dict):
                                print(f"       → .{k} = dict de {len(v)} éléments")
                            elif isinstance(v, list):
                                print(f"       → .{k} = list de {len(v)} éléments")
                elif isinstance(parsed, list):
                    print(f"       → list, {len(parsed)} éléments")
            except json.JSONDecodeError as e:
                print(f"       parse json fail: {e}")

    # ============================================================== 3. Form variants
    section("3. Form add-to-cart et selects de variantes")
    # Cherche les <select> ou <input type=radio> dans le form
    form_match = re.search(
        r'<form[^>]+id=["\']add-to-cart-or-refresh["\'][^>]*>(.*?)</form>',
        body, flags=re.DOTALL | re.IGNORECASE,
    )
    if not form_match:
        form_match = re.search(
            r'<form[^>]+action="[^"]*"[^>]+id="[^"]*product[^"]*"[^>]*>(.*?)</form>',
            body, flags=re.DOTALL | re.IGNORECASE,
        )
    if form_match:
        form_html = form_match.group(1)
        print(f"  Form trouvé ({len(form_html)} chars)")
        # Extract <select>
        selects = re.findall(
            r'<select[^>]+(?:name|id)=["\']([^"\']+)["\'][^>]*>(.*?)</select>',
            form_html, flags=re.DOTALL,
        )
        for sname, sbody in selects:
            options = re.findall(r'<option[^>]*value=["\']([^"\']+)["\']([^>]*)>(.*?)</option>',
                                 sbody, flags=re.DOTALL)
            print(f"  <select name={sname!r}> : {len(options)} options")
            for val, attrs, label in options[:5]:
                label_clean = re.sub(r"<[^>]+>", "", label).strip()
                print(f"      - value={val!r} label={label_clean!r}")
        # Extract radio/checkbox groups
        radios = re.findall(
            r'<input[^>]+type=["\'](?:radio|checkbox)["\'][^>]+name=["\']([^"\']+)["\'][^>]+value=["\']([^"\']+)["\']',
            form_html,
        )
        if radios:
            from collections import Counter
            radio_groups = Counter(name for name, _ in radios)
            print(f"  Groupes radio/checkbox :")
            for name, count in radio_groups.items():
                print(f"      {name!r} : {count} options")
        # Cherche aussi des spans/li avec data-*
        attr_items = re.findall(
            r'<li[^>]*class="[^"]*input-container[^"]*"[^>]*>.*?<span[^>]*>(.*?)</span>',
            form_html, flags=re.DOTALL,
        )
        if attr_items:
            print(f"  Attributs en li.input-container : {len(attr_items)}")
            for x in attr_items[:5]:
                print(f"      {re.sub(r'[<][^>]+[>]', '', x).strip()[:80]!r}")
    else:
        print(f"  Pas de form add-to-cart trouvé")

    # ============================================================== 4. Microdata
    section("4. Microdata (itemprop)")
    microdata_patterns = [
        (r'<meta[^>]+itemprop="price"[^>]+content="([^"]+)"', "price"),
        (r'<meta[^>]+itemprop="priceCurrency"[^>]+content="([^"]+)"', "priceCurrency"),
        (r'<link[^>]+itemprop="availability"[^>]+href="([^"]+)"', "availability"),
        (r'<meta[^>]+itemprop="sku"[^>]+content="([^"]+)"', "sku"),
    ]
    for pat, label in microdata_patterns:
        m = re.search(pat, body)
        if m:
            print(f"  [OK] {label} = {m.group(1)!r}")

    # ============================================================== 5. Description
    section("5. Description : sélecteurs candidats")
    desc_patterns = [
        (r'<div[^>]+itemprop="description"[^>]*>(.*?)</div>', "[itemprop=description]"),
        (r'<div[^>]+id="description"[^>]*>(.*?)</div>(?=\s*<(?:div|section|footer))', "#description"),
        (r'<div[^>]+class="[^"]*product-description[^"]*"[^>]*>(.*?)</div>(?=\s*<(?:div|section|footer))', ".product-description"),
        (r'<section[^>]+class="[^"]*product-description[^"]*"[^>]*>(.*?)</section>', "section.product-description"),
    ]
    for pat, label in desc_patterns:
        m = re.search(pat, body, flags=re.DOTALL | re.IGNORECASE)
        if m:
            content = m.group(1)
            cleaned = re.sub(r"<[^>]+>", " ", content)
            cleaned = re.sub(r"\s+", " ", cleaned).strip()
            print(f"  [OK] {label} ({len(content)} chars HTML, {len(cleaned)} cleaned)")
            print(f"       preview : {cleaned[:200]!r}")

    # ============================================================== 6. Images
    section("6. Images produit")
    img_patterns = [
        (r'data-image-large-src="([^"]+)"', "data-image-large-src"),
        (r'<a[^>]+class="[^"]*thumb[^"]*"[^>]+data-image-large-src="([^"]+)"', "a.thumb data-image-large-src"),
        (r'<img[^>]+class="[^"]*js-qv-product-cover[^"]*"[^>]+src="([^"]+)"', ".js-qv-product-cover"),
        (r'<img[^>]+class="[^"]*product-cover[^"]*"[^>]+src="([^"]+)"', ".product-cover"),
        (r'<a[^>]+class="[^"]*thumb[^"]*"[^>]+href="([^"]+)"', "a.thumb href"),
    ]
    seen = set()
    for pat, label in img_patterns:
        urls = re.findall(pat, body)
        new = [u for u in urls if u not in seen]
        seen.update(new)
        if urls:
            print(f"  [OK] {label} : {len(urls)} matches ({len(new)} nouvelles)")
            for u in new[:5]:
                print(f"        {u}")

    print(f"\n  Total images uniques : {len(seen)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
