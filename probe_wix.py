#!/usr/bin/env python3
"""
Probe Wix — détection des sources de données exploitables pour scraper.

Pour chaque URL marque, on regarde :
  1) Signatures Wix dans la home (parastorage.com, meta generator, etc.)
  2) Sitemap (les URLs produit)
  3) Tentatives d'extraction sur une page produit :
       - JSON-LD Schema.org Product (s'il est server-rendered)
       - Variables JS globales (window.publicData, viewerModel, etc.)
       - Endpoints internes _api/wix-ecommerce-storefront-web
       - Body brut (pour confirmer si JS-rendered ou pré-rendu)

Usage :
    python3 probe_wix.py https://www.labobio-paris.com/
    python3 probe_wix.py url1 url2 url3

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
from urllib.parse import urlparse


UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
TIMEOUT = 25


def _ssl_ctx():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def fetch(url: str) -> tuple[int, dict, str]:
    headers = {
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, identity",
    }
    req = urllib.request.Request(url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT, context=_ssl_ctx()) as resp:
            raw = resp.read()
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                try:
                    raw = gzip.decompress(raw)
                except Exception:
                    pass
            hdrs = {k.lower(): v for k, v in resp.headers.items()}
            return resp.status, hdrs, raw.decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        return e.code, dict(e.headers or {}), body
    except Exception as e:
        print(f"    [err] {type(e).__name__}: {e}")
        return 0, {}, ""


def section(title: str) -> None:
    print(f"\n{'-' * 78}\n  {title}\n{'-' * 78}")


def detect_wix_signatures(body: str) -> dict:
    sig = {}
    if "parastorage.com" in body:
        sig["parastorage_cdn"] = True
    m = re.search(
        r'<meta[^>]+name=["\']generator["\'][^>]+content=["\']([^"\']*[Ww]ix[^"\']*)["\']',
        body,
    )
    if m:
        sig["meta_generator"] = m.group(1)
    if "window.__INITIAL_STATE__" in body or "window.__PUBLIC_DATA__" in body:
        sig["wix_window_state"] = True
    if "wix-ecommerce-storefront" in body:
        sig["wix_ecommerce_storefront"] = True
    if "wix-stores" in body.lower() or "wixstores" in body.lower():
        sig["wix_stores"] = True
    if "_api/wix-" in body or "_api/stores-" in body:
        sig["wix_api_paths"] = True
    return sig


def find_product_urls(sitemap_url: str) -> list[str]:
    s, h, body = fetch(sitemap_url)
    if s != 200 or "<loc>" not in body:
        return []
    locs = re.findall(r"<loc>([^<]+)</loc>", body)
    # Wix uses /product-page/{slug} for Wix Stores
    products = [u for u in locs if "/product-page/" in u or "/products/" in u]
    if products:
        return products
    # If no clear pattern, return all
    return locs


def inspect_product_page(url: str) -> None:
    section(f"Inspection produit : {url}")
    s, h, body = fetch(url)
    if s != 200:
        print(f"  HTTP {s} - skip")
        return
    print(f"  HTTP 200, body = {len(body)} chars")

    # Save HTML for inspection
    fn = re.sub(r"[^a-zA-Z0-9._-]", "_", url.split("//")[-1])[:100] + ".html"
    with open(fn, "w", encoding="utf-8") as f:
        f.write(body)
    print(f"  HTML sauvegardé : {fn}")

    # 1) JSON-LD
    print("\n  >>> JSON-LD Product")
    blocks = re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        body, flags=re.DOTALL | re.IGNORECASE,
    )
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
                if t == "Product" or (isinstance(t, list) and "Product" in t):
                    product_ld = c
                    break
        if product_ld:
            break
    if product_ld:
        print(f"  [OK] Product trouvé, clés : {list(product_ld.keys())}")
        for k in ("name", "sku", "description", "image", "offers", "category", "brand"):
            v = product_ld.get(k)
            if v:
                vstr = json.dumps(v, ensure_ascii=False)[:180] if isinstance(v, (dict, list)) else str(v)[:180]
                print(f"      .{k} = {vstr!r}")
    else:
        print(f"  [KO] Pas de JSON-LD Product ({len(blocks)} blocs JSON-LD au total)")
        # Voir les @type présents
        types_seen = set()
        for blk in blocks:
            try:
                d = json.loads(blk.strip())
                cands = d if isinstance(d, list) else (d.get("@graph", []) if isinstance(d, dict) and "@graph" in d else [d])
                for c in cands:
                    if isinstance(c, dict):
                        t = c.get("@type")
                        if t:
                            types_seen.add(str(t))
            except Exception:
                pass
        print(f"       @type vus : {types_seen}")

    # 2) Variables JS globales Wix
    print("\n  >>> Variables JS Wix")
    for pat, label in (
        (r'window\.__PUBLIC_DATA__\s*=\s*({[\s\S]+?})\s*;', "window.__PUBLIC_DATA__"),
        (r'window\.__INITIAL_STATE__\s*=\s*({[\s\S]+?})\s*</script', "window.__INITIAL_STATE__"),
        (r'window\.viewerModel\s*=\s*({[\s\S]+?})\s*;', "window.viewerModel"),
        (r'window\.commonConfig\s*=\s*({[\s\S]+?})\s*;', "window.commonConfig"),
    ):
        m = re.search(pat, body)
        if m:
            content = m.group(1)
            print(f"  [OK] {label} ({len(content)} chars)")
            # Check for product-related keys
            for keyword in ("product", "sku", "price", "image"):
                if f'"{keyword}"' in content[:5000]:
                    print(f"       contient {keyword!r} dans les 5000 premiers chars")

    # 3) Liens vers les endpoints _api/...
    print("\n  >>> Endpoints _api/ référencés dans la page")
    api_urls = set(re.findall(r'(/_api/[a-z0-9/\-_]+)', body))
    print(f"  {len(api_urls)} URLs _api uniques :")
    for u in list(api_urls)[:10]:
        print(f"      {u}")

    # 4) Open Graph meta (souvent présent sur Wix pour SEO)
    print("\n  >>> Open Graph meta")
    for prop in ("og:title", "og:description", "og:image", "og:price:amount",
                 "product:price:amount", "product:price:currency"):
        m = re.search(
            r'<meta[^>]+(?:property|name)=["\']' + re.escape(prop) + r'["\'][^>]+content=["\']([^"\']+)["\']',
            body, flags=re.IGNORECASE,
        )
        if m:
            print(f"  [OK] {prop} = {m.group(1)[:150]!r}")

    # 5) Confirmer si la page est server-rendered ou JS-only
    print("\n  >>> Body server-rendered ou JS-only ?")
    body_text = re.sub(r"<script[\s\S]+?</script>", " ", body)
    body_text = re.sub(r"<style[\s\S]+?</style>", " ", body_text)
    body_text = re.sub(r"<[^>]+>", " ", body_text)
    body_text = re.sub(r"\s+", " ", body_text).strip()
    print(f"      Texte visible (hors scripts/styles) : {len(body_text)} chars")
    if len(body_text) > 1000:
        print(f"      Preview (chars 0-300) : {body_text[:300]!r}")
        print(f"      [INFO] Page contient du texte → server-rendered partiel")
    else:
        print(f"      [WARN] Quasi-vide → JS-rendered, headless browser nécessaire")


def probe(brand_url: str) -> None:
    print("\n" + "=" * 78)
    print(f"  {brand_url}")
    print("=" * 78)

    parsed = urlparse(brand_url if "://" in brand_url else f"https://{brand_url}")
    base = f"{parsed.scheme}://{parsed.netloc}"

    # 1. Home + signatures
    section("[1] Détection Wix dans la home")
    s, h, body = fetch(brand_url)
    print(f"      HTTP {s}, bytes={len(body)}")
    if s != 200:
        print(f"      Home inaccessible — skip")
        return
    sigs = detect_wix_signatures(body)
    if sigs:
        print(f"      Signatures Wix détectées :")
        for k, v in sigs.items():
            print(f"          {k}: {v}")
    else:
        print(f"      AUCUNE signature Wix trouvée (le site n'est peut-être pas Wix)")

    # 2. Sitemap
    section("[2] Sitemap")
    for sm_url in (f"{base}/sitemap.xml", f"{base}/sitemap-index.xml"):
        s, h, body_sm = fetch(sm_url)
        if s != 200 or "<" not in body_sm[:100]:
            continue
        print(f"      [OK] {sm_url}")
        locs = re.findall(r"<loc>([^<]+)</loc>", body_sm)
        print(f"      {len(locs)} URLs dans le sitemap")

        # Collecte les vraies URLs produits (depuis les sub-sitemaps si index)
        product_urls: list[str] = []
        if "<sitemapindex" in body_sm[:300]:
            print(f"      C'est un index → exploration des sub-sitemaps :")
            for sub in locs:
                s2, _, body2 = fetch(sub)
                if s2 == 200:
                    sub_locs = re.findall(r"<loc>([^<]+)</loc>", body2)
                    print(f"          {sub}: {len(sub_locs)} URLs")
                    # Si c'est un store-products-sitemap, on capture les produits
                    if "product" in sub.lower():
                        product_urls.extend(sub_locs)
        else:
            # Sitemap direct, on filtre
            product_urls = [u for u in locs if "/product-page/" in u or "/products/" in u]

        if product_urls:
            print(f"\n      [OK] {len(product_urls)} URLs produits collectées")
            print(f"      Échantillon :")
            for u in product_urls[:5]:
                print(f"          {u}")
            # 3. Inspecte la première page produit
            inspect_product_page(product_urls[0])
        else:
            print(f"      Aucune URL produit trouvée")
        return


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage : python3 probe_wix.py <url1> [<url2> ...]")
        return 2
    for url in sys.argv[1:]:
        try:
            probe(url)
        except Exception as e:
            print(f"\n[FATAL] {url} : {type(e).__name__}: {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
