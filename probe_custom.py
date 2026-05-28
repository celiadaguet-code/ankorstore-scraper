#!/usr/bin/env python3
"""
Probe pour sites e-commerce custom (sans CMS connu).

Pour chaque URL marque :
  1) Détection signatures de TOUS les CMS courants (Shopify, WP/Woo, Presta, Wix,
     Squarespace, Magento, OpenCart, BigCommerce, etc.) — peut-être que ce n'est
     pas si "custom" que ça
  2) Sitemap.xml + structure (URLs produits déductibles ?)
  3) Inspection d'une page produit (heuristique) pour voir :
       - JSON-LD Product schema (universel pour le SEO)
       - Open Graph meta (product:price:amount/currency)
       - Microdata (itemprop="price", "name", "description", "sku")
       - Variables JS globales notables

ZÉRO dépendance, stdlib only.

Usage :
    python3 probe_custom.py https://www.brasserieco-hop.fr/ https://lessenteursdelanature.fr/
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
from urllib.parse import urlparse


UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


def fetch(url: str, accept: str = "text/html,*/*"):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": accept,
        "Accept-Encoding": "gzip, identity",
    })
    try:
        with urllib.request.urlopen(req, timeout=25, context=ctx) as resp:
            raw = resp.read()
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                try:
                    raw = gzip.decompress(raw)
                except Exception:
                    pass
            return resp.status, dict(resp.headers), raw.decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        return e.code, dict(e.headers or {}), body
    except Exception as e:
        return 0, {}, f"ERROR: {type(e).__name__}: {e}"


def section(t):
    print(f"\n{'-'*78}\n  {t}\n{'-'*78}")


def detect_known_cms(body: str, hdrs: dict) -> dict:
    """Détecte signatures de tous les CMS courants."""
    sigs: dict[str, list[str]] = {}
    body_lower = body.lower()
    server = (hdrs.get("server") or hdrs.get("Server") or "").lower()
    powered = (hdrs.get("x-powered-by") or "").lower()

    def add(cms, sig):
        sigs.setdefault(cms, []).append(sig)

    # Shopify
    if "cdn.shopify.com" in body_lower or "shopify-section" in body_lower:
        add("shopify", "cdn.shopify.com / shopify-section")
    if "x-shopid" in hdrs or "x-shopify-stage" in hdrs:
        add("shopify", "header X-ShopId")
    if "myshopify.com" in body_lower:
        add("shopify", "myshopify.com")

    # WooCommerce / WordPress
    if "wp-content" in body_lower:
        add("wordpress_or_woo", "wp-content")
    if "/wp-json/" in body_lower:
        add("wordpress_or_woo", "/wp-json/")
    if re.search(r'\bwoocommerce\b', body, re.IGNORECASE):
        add("wordpress_or_woo", "woocommerce keyword")

    # PrestaShop
    if re.search(r'\bvar\s+prestashop\s*=', body):
        add("prestashop", "var prestashop=")
    if "data-prestashop-version" in body:
        add("prestashop", "data-prestashop-version")

    # Wix
    if "parastorage.com" in body_lower:
        add("wix", "parastorage.com")
    if re.search(r'<meta[^>]+name=["\']generator["\'][^>]+content=["\'][^"\']*[Ww]ix', body):
        add("wix", "meta generator Wix")

    # Squarespace
    if "SQUARESPACE_CONTEXT" in body or "squarespace-cdn" in body_lower:
        add("squarespace", "SQUARESPACE_CONTEXT or squarespace-cdn")

    # Magento
    if "/static/version" in body_lower or "mage/cookies" in body_lower:
        add("magento", "/static/version / mage/cookies")
    if "magento" in powered:
        add("magento", f"X-Powered-By: {powered}")

    # OpenCart
    if "catalog/view/theme" in body_lower:
        add("opencart", "catalog/view/theme")

    # BigCommerce
    if "bigcommerce.com" in body_lower or "stencil-cdn" in body_lower:
        add("bigcommerce", "stencil-cdn")

    # Webflow
    if "webflow.com" in body_lower:
        add("webflow", "webflow.com")

    return sigs


def probe(url: str) -> None:
    print(f"\n{'='*78}\n  {url}\n{'='*78}")
    parsed = urlparse(url if "://" in url else f"https://{url}")
    base = f"{parsed.scheme}://{parsed.netloc}"
    is_root = (not parsed.path or parsed.path == "/")

    # MODE DIRECT : URL non-racine → on inspecte cette URL directement comme
    # une page produit (utile quand le sitemap est inutile ou trompeur).
    if not is_root:
        section(f"[Mode direct] Inspection directe de {url}")
        inspect_product(url)
        # Et on regarde aussi les signatures dans cette page
        section("[Signatures CMS sur cette page]")
        s, hdrs, body = fetch(url)
        if s == 200:
            sigs = detect_known_cms(body, hdrs)
            if sigs:
                for cms, marks in sigs.items():
                    print(f"  → {cms.upper()} : {marks}")
            else:
                print("  → Aucun CMS connu détecté.")
        return

    section("[1] Détection CMS (parmi les connus)")
    s, hdrs, body = fetch(url)
    print(f"  HTTP {s}, bytes={len(body)}")
    if s != 200:
        print("  Home inaccessible — skip")
        return
    sigs = detect_known_cms(body, hdrs)
    if sigs:
        for cms, marks in sigs.items():
            print(f"  → {cms.upper()} : {marks}")
    else:
        print("  → Aucun CMS connu détecté. C'est probablement du custom / sur-mesure.")

    # Server header
    server = hdrs.get("server", hdrs.get("Server", ""))
    if server:
        print(f"  Server header : {server}")
    powered = hdrs.get("x-powered-by", "")
    if powered:
        print(f"  X-Powered-By : {powered}")

    section("[2] Sitemap.xml")
    s, _, body_sm = fetch(f"{base}/sitemap.xml", accept="application/xml,*/*")
    if s != 200:
        print(f"  /sitemap.xml inaccessible (status={s})")
        # Tente /sitemap_index.xml
        s, _, body_sm = fetch(f"{base}/sitemap_index.xml", accept="application/xml,*/*")
        if s == 200:
            print(f"  /sitemap_index.xml OK")
    if s != 200:
        print(f"  Pas de sitemap accessible")
        return
    locs = re.findall(r"<loc>([^<]+)</loc>", body_sm)
    is_index = "<sitemapindex" in body_sm[:300]
    print(f"  {len(locs)} URLs (type: {'index' if is_index else 'urlset'})")
    if is_index:
        print(f"  Sub-sitemaps détectés :")
        for sub in locs[:10]:
            print(f"      {sub}")
        # Tente de récupérer les URLs produits via le sub-sitemap qui contient 'product'
        product_sub = next((u for u in locs if "product" in u.lower()), None)
        if product_sub:
            s2, _, body2 = fetch(product_sub)
            if s2 == 200:
                product_locs = re.findall(r"<loc>([^<]+)</loc>", body2)
                print(f"\n  {len(product_locs)} URLs dans {product_sub}")
                for u in product_locs[:5]:
                    print(f"      {u}")
                if product_locs:
                    return inspect_product(product_locs[0])
    else:
        # Sitemap direct : montre quelques URLs
        print(f"  Échantillon :")
        for u in locs[:10]:
            print(f"      {u}")
        # Détecte un pattern produit
        candidates = [u for u in locs if re.search(
            r"/(?:product|produit|shop|store|boutique|p)/", u, re.IGNORECASE,
        ) or re.search(r"/[a-z0-9][a-z0-9\-_]{3,}\.html?$", u, re.IGNORECASE)]
        if candidates:
            print(f"\n  {len(candidates)} URLs avec pattern produit-like")
            return inspect_product(candidates[0])
        # Last resort : inspecte juste la 2e URL (la 1ère est souvent la home)
        if len(locs) > 1:
            return inspect_product(locs[1])


def inspect_product(url: str) -> None:
    """Inspecte une page produit candidate."""
    section(f"[3] Inspection page produit : {url}")
    s, _, body = fetch(url)
    if s != 200 or not body:
        print(f"  HTTP {s} — skip")
        return
    print(f"  HTTP 200, {len(body)} chars")

    # Save HTML
    fn = re.sub(r"[^a-zA-Z0-9._-]", "_", url.split("//")[-1])[:80] + ".html"
    with open(f"debug_custom_{fn}", "w", encoding="utf-8") as f:
        f.write(body)
    print(f"  HTML dumpé : debug_custom_{fn}")

    # JSON-LD
    print("\n  >>> JSON-LD")
    blocks = re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        body, flags=re.DOTALL | re.IGNORECASE,
    )
    print(f"  {len(blocks)} bloc(s) JSON-LD")
    product_ld = None
    for blk in blocks:
        try:
            d = json.loads(blk.strip())
        except json.JSONDecodeError:
            continue
        candidates = d if isinstance(d, list) else (
            d.get("@graph", []) if isinstance(d, dict) and "@graph" in d else [d]
        )
        for c in candidates:
            if isinstance(c, dict):
                t = c.get("@type")
                types_str = (t if isinstance(t, str) else json.dumps(t))
                print(f"      @type: {types_str}")
                if t == "Product" or (isinstance(t, list) and "Product" in t):
                    product_ld = c
                    break
    if product_ld:
        print(f"  [OK] Product trouvé. Clés : {list(product_ld.keys())}")
        for k in ("name", "sku", "description", "image", "offers", "brand", "category"):
            v = product_ld.get(k)
            if v:
                preview = json.dumps(v, ensure_ascii=False)[:250] if isinstance(v, (dict, list)) else str(v)[:250]
                print(f"      .{k} = {preview!r}")
    else:
        print(f"  [KO] Pas de JSON-LD Product")

    # OG meta
    print("\n  >>> Open Graph meta")
    for prop in ("og:title", "og:description", "og:image",
                 "product:price:amount", "product:price:currency",
                 "og:price:amount", "og:price:currency",
                 "twitter:title", "twitter:description"):
        m = re.search(
            r'<meta[^>]+(?:property|name)=["\']' + re.escape(prop) + r'["\'][^>]+content=["\']([^"\']+)["\']',
            body, flags=re.IGNORECASE,
        )
        if m:
            print(f"  [OK] {prop} = {m.group(1)[:200]!r}")

    # Microdata
    print("\n  >>> Microdata itemprop")
    for prop in ("price", "priceCurrency", "name", "description", "sku", "availability"):
        for pat in (
            rf'<meta[^>]+itemprop=["\']{prop}["\'][^>]+content=["\']([^"\']+)["\']',
            rf'<[^>]+itemprop=["\']{prop}["\'][^>]*>([^<]+)<',
        ):
            m = re.search(pat, body, flags=re.IGNORECASE)
            if m:
                print(f"  [OK] {prop} = {m.group(1)[:200]!r}")
                break


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage : python3 probe_custom.py <url1> [<url2> ...]")
        return 2
    for u in sys.argv[1:]:
        try:
            probe(u)
        except Exception as e:
            print(f"\n[FATAL] {u} : {type(e).__name__}: {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
