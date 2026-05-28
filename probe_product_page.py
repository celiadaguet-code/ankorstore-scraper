#!/usr/bin/env python3
"""
Probe d'une page produit WooCommerce (pour les sites où la Store API est cassée).

On regarde ce qu'on peut extraire du HTML :
  1) JSON-LD Product schema  (le plus fiable)
  2) data-product_variations sur le form.variations_form  (pour les variants)
  3) Sélecteurs WooCommerce standards (titre, prix, images, SKU, description)

Usage :
    python3 probe_product_page.py https://comptoir-saveurs.fr/produit/cafe-melange-de-noel/

ZÉRO dépendance — stdlib uniquement.
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


def fetch_html(url: str) -> str:
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, identity",
    })
    with urllib.request.urlopen(req, timeout=30, context=_ssl_ctx()) as resp:
        raw = resp.read()
        if resp.headers.get("Content-Encoding", "").lower() == "gzip":
            try:
                raw = gzip.decompress(raw)
            except Exception:
                pass
        return raw.decode("utf-8", errors="replace")


def section(title: str) -> None:
    print(f"\n{'-' * 78}\n  {title}\n{'-' * 78}")


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage : python3 probe_product_page.py <product_url>")
        return 2
    url = sys.argv[1]
    print("=" * 78)
    print(f"  Inspection de : {url}")
    print("=" * 78)
    print("  Fetching HTML...")
    try:
        html_text = fetch_html(url)
    except Exception as e:
        print(f"  ERROR: {type(e).__name__}: {e}")
        return 1
    print(f"  HTML reçu : {len(html_text)} chars")

    # -------------------------------------------------------------- JSON-LD
    section("1) JSON-LD Product schema")
    ld_blocks = re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        html_text, flags=re.DOTALL | re.IGNORECASE,
    )
    print(f"  {len(ld_blocks)} bloc(s) JSON-LD trouvé(s)")
    product_ld = None
    for i, block in enumerate(ld_blocks):
        try:
            data = json.loads(block.strip())
        except json.JSONDecodeError as e:
            print(f"  [{i}] JSON invalide : {e}")
            continue
        # Peut être un dict (avec @graph parfois) ou une liste
        candidates = []
        if isinstance(data, list):
            candidates = data
        elif isinstance(data, dict):
            if "@graph" in data and isinstance(data["@graph"], list):
                candidates = data["@graph"]
            else:
                candidates = [data]
        for c in candidates:
            if isinstance(c, dict):
                ctype = c.get("@type")
                if ctype == "Product" or (isinstance(ctype, list) and "Product" in ctype):
                    product_ld = c
                    break
        if product_ld:
            break

    if product_ld:
        print("  >> Schema Product TROUVÉ. Clés disponibles :")
        print(f"     {list(product_ld.keys())}")
        for k in ("name", "sku", "description", "image", "brand", "category"):
            if k in product_ld:
                v = product_ld[k]
                if isinstance(v, (dict, list)):
                    print(f"     .{k} = (preview) {json.dumps(v, ensure_ascii=False)[:200]}")
                else:
                    print(f"     .{k} = {str(v)[:200]!r}")
        # Offers (prix, stock)
        offers = product_ld.get("offers")
        if offers:
            print(f"     .offers = (preview) {json.dumps(offers, ensure_ascii=False)[:300]}")
        # AggregateOffer pour produits variables
        if isinstance(offers, dict):
            print(f"     .offers keys : {list(offers.keys())}")
    else:
        print("  >> Pas de schema Product trouvé en JSON-LD")

    # -------------------------------------------------- data-product_variations
    section("2) data-product_variations (form.variations_form)")
    # L'attribut peut être entouré de " ou '
    patterns = [
        r'data-product_variations\s*=\s*"([^"]+)"',
        r"data-product_variations\s*=\s*'([^']+)'",
        r'data-product_variations\s*=\s*&quot;(.+?)&quot;',
    ]
    found = None
    for p in patterns:
        m = re.search(p, html_text, flags=re.DOTALL)
        if m:
            found = m.group(1)
            print(f"  >> Match trouvé via pattern : {p[:50]}...")
            break

    if found:
        # Décoder entités HTML
        decoded = html.unescape(found)
        try:
            variations = json.loads(decoded)
            print(f"  >> {len(variations)} variations parsées")
            if variations:
                v0 = variations[0]
                print(f"     [0] keys : {list(v0.keys())}")
                for k in ("variation_id", "sku", "display_price", "display_regular_price",
                          "is_in_stock", "is_purchasable", "max_qty", "min_qty",
                          "attributes", "image", "variation_description"):
                    if k in v0:
                        v = v0[k]
                        if isinstance(v, (dict, list)):
                            print(f"     [0].{k} = (preview) {json.dumps(v, ensure_ascii=False)[:250]}")
                        else:
                            print(f"     [0].{k} = {str(v)[:200]!r}")
        except json.JSONDecodeError as e:
            print(f"  >> Trouvé mais parse JSON échoué : {e}")
            print(f"     preview : {decoded[:300]!r}")
    else:
        if "variations_form" in html_text:
            print("  >> form.variations_form présent mais pas d'attribut data-product_variations directement.")
            print("  >> Peut-être chargé en AJAX. Recherche des indices...")
        else:
            print("  >> Pas de form variations_form (produit simple, ou thème custom)")

    # ---------------------------------------------- Sélecteurs HTML classiques
    section("3) Sélecteurs HTML WooCommerce standards")

    def find_first(patterns):
        for p in patterns:
            m = re.search(p, html_text, flags=re.DOTALL | re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return None

    # Titre H1
    title = find_first([
        r'<h1[^>]*class="[^"]*product_title[^"]*"[^>]*>(.*?)</h1>',
        r'<h1[^>]*class="[^"]*entry-title[^"]*"[^>]*>(.*?)</h1>',
    ])
    print(f"  H1 titre        : {(title or '')[:120]!r}")

    # SKU
    sku = find_first([
        r'<span[^>]*class="[^"]*sku[^"]*"[^>]*>(.*?)</span>',
        r'class="sku_wrapper">[^<]*<span[^>]*>(.*?)</span>',
    ])
    print(f"  SKU             : {sku!r}")

    # Price
    price = find_first([
        r'<p[^>]*class="[^"]*price[^"]*"[^>]*>(.*?)</p>',
        r'<span[^>]*class="[^"]*price[^"]*"[^>]*>(.*?)</span>',
    ])
    # Nettoie le prix pour l'aperçu
    if price:
        price_clean = re.sub(r"<[^>]+>", " ", price)
        price_clean = re.sub(r"\s+", " ", price_clean).strip()
    else:
        price_clean = None
    print(f"  Prix (brut)     : {(price_clean or '')[:150]!r}")

    # Description courte
    short = find_first([
        r'<div[^>]*class="[^"]*woocommerce-product-details__short-description[^"]*"[^>]*>(.*?)</div>',
        r'<div[^>]*itemprop="description"[^>]*>(.*?)</div>',
    ])
    if short:
        short_clean = re.sub(r"<[^>]+>", " ", short)
        short_clean = re.sub(r"\s+", " ", short_clean).strip()
    else:
        short_clean = None
    print(f"  Short desc      : {(short_clean or '')[:200]!r}")

    # Images de galerie
    gallery = re.findall(
        r'data-large_image="([^"]+)"',
        html_text,
    )
    if not gallery:
        gallery = re.findall(
            r'<img[^>]+class="[^"]*wp-post-image[^"]*"[^>]+src="([^"]+)"',
            html_text,
        )
    print(f"  Images galerie  : {len(gallery)} trouvées")
    for u in gallery[:5]:
        print(f"      - {u}")

    # Catégories
    cats = re.findall(
        r'<a[^>]+rel="tag"[^>]*>(.*?)</a>',
        html_text,
    )
    cats = [re.sub(r"<[^>]+>", "", c).strip() for c in cats]
    print(f"  Tags/cats (rel=tag) : {cats[:10]}")

    # Catégories produit (posted_in)
    posted_in = re.search(
        r'class="posted_in">[^<]*(.+?)</span>',
        html_text, flags=re.DOTALL,
    )
    if posted_in:
        pi_clean = re.sub(r"<[^>]+>", " ", posted_in.group(1))
        pi_clean = re.sub(r"\s+", " ", pi_clean).strip()
        print(f"  Catégories      : {pi_clean[:200]!r}")

    print("\n" + "=" * 78)
    print("  Recommandation :")
    if product_ld and found:
        print("  -> JSON-LD + data-product_variations : SUFFISANT pour faire le scraper")
    elif product_ld:
        print("  -> JSON-LD seul : OK pour produits simples ; pour les variants il faudra")
        print("     parser le HTML (sélecteurs CSS)")
    else:
        print("  -> Aucune source structurée : scraping HTML pur, fragile")
    print("=" * 78)
    return 0


if __name__ == "__main__":
    sys.exit(main())
