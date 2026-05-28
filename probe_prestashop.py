#!/usr/bin/env python3
"""
Probe PrestaShop — détection et inspection complète d'un catalogue Presta.

Pour chaque URL marque, on regarde :
  1) Signatures PrestaShop dans la home (var prestashop=..., generator meta, etc.)
  2) Sitemap (sitemap.xml, /fr/sitemap.xml, ou robots.txt)
  3) Liste des URLs produits trouvées
  4) Inspection d'UNE page produit :
       - JSON-LD Product schema
       - JS globals (prestashop.combinations, productGalleryImages)
       - Sélecteurs HTML classiques (h1.product-detail, .product-prices, ...)
       - Combinations (variantes)

Usage :
    python3 probe_prestashop.py https://lagazellemarrakchia.com/
    python3 probe_prestashop.py https://lagazellemarrakchia.com/ https://www.moulin-garrigue.fr/

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


USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
TIMEOUT = 20


# --- SSL ---------------------------------------------------------------

def _ssl_ctx():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


_SSL_OK = None


def _ssl_works() -> bool:
    global _SSL_OK
    if _SSL_OK is not None:
        return _SSL_OK
    try:
        req = urllib.request.Request("https://www.google.com/", headers={"User-Agent": USER_AGENT})
        urllib.request.urlopen(req, timeout=5).close()
        _SSL_OK = True
    except Exception:
        _SSL_OK = False
        print("[INFO] SSL non vérifié pour ce run (bundle CA absent).", flush=True)
    return _SSL_OK


# --- HTTP --------------------------------------------------------------

def fetch(url: str, accept: str = "text/html") -> tuple[int, dict, str]:
    """Retourne (status, headers, body_text). Body vide en cas d'erreur."""
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": accept,
        "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, identity",
    }
    req = urllib.request.Request(url, headers=headers)
    ctx = ssl.create_default_context() if _ssl_works() else _ssl_ctx()
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT, context=ctx) as resp:
            raw = resp.read()
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                try:
                    raw = gzip.decompress(raw)
                except Exception:
                    pass
            text = raw.decode("utf-8", errors="replace")
            hdrs = {k.lower(): v for k, v in resp.headers.items()}
            return resp.status, hdrs, text
    except urllib.error.HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        return e.code, dict(e.headers or {}), body
    except Exception as e:
        print(f"    [err] {type(e).__name__}: {e}")
        return 0, {}, ""


# --- Détection PrestaShop ---------------------------------------------

def detect_prestashop(home_html: str, hdrs: dict) -> dict:
    """Cherche les signatures PrestaShop dans la home."""
    signals = {}
    # Meta generator
    m = re.search(
        r'<meta[^>]+name=["\']generator["\'][^>]+content=["\']([^"\']*PrestaShop[^"\']*)["\']',
        home_html, flags=re.IGNORECASE,
    )
    if m:
        signals["meta_generator"] = m.group(1)
    # Variable JS globale prestashop
    if re.search(r'\bvar\s+prestashop\s*=', home_html):
        signals["js_var_prestashop"] = True
    if re.search(r'window\.prestashop\b', home_html):
        signals["js_window_prestashop"] = True
    # Header X-Powered-By
    xpb = hdrs.get("x-powered-by") or hdrs.get("X-Powered-By")
    if xpb and "prestashop" in xpb.lower():
        signals["x_powered_by"] = xpb
    # Path /themes/ caractéristique
    if "/themes/" in home_html and "prestashop" in home_html.lower():
        signals["themes_path"] = True
    # body id="checkout" (PS specific)
    if re.search(r'<body[^>]+id=["\']checkout["\']', home_html):
        signals["body_checkout"] = True
    # data-prestashop-version
    m = re.search(r'data-prestashop-version=["\']([^"\']+)["\']', home_html)
    if m:
        signals["prestashop_version"] = m.group(1)
    return signals


# --- Sitemap -----------------------------------------------------------

def find_sitemap(base: str) -> tuple[str | None, list[str]]:
    """Cherche un sitemap accessible. Retourne (url_sitemap, list_des_urls_produits)."""
    candidates = [
        f"{base}/sitemap.xml",
        f"{base}/fr/sitemap.xml",
        f"{base}/1_fr_0_sitemap.xml",
        f"{base}/sitemap_index.xml",
        f"{base}/sitemap-index.xml",
    ]
    # Aussi check robots.txt
    s, h, body = fetch(f"{base}/robots.txt", "text/plain")
    if s == 200 and body:
        for line in body.splitlines():
            if line.lower().startswith("sitemap:"):
                sm_url = line.split(":", 1)[1].strip()
                if sm_url and sm_url not in candidates:
                    candidates.insert(0, sm_url)

    for url in candidates:
        s, h, body = fetch(url, "application/xml")
        if s != 200 or not body.startswith(("<?xml", "<urlset", "<sitemapindex")):
            continue
        locs = re.findall(r"<loc>([^<]+)</loc>", body)
        # Si c'est un index, on suit les sub-sitemaps qui contiennent "product"
        if "<sitemapindex" in body[:200]:
            sub_locs = []
            for sub in locs:
                if "product" in sub.lower() or "catalog" in sub.lower():
                    s2, _, body2 = fetch(sub, "application/xml")
                    if s2 == 200:
                        sub_locs.extend(re.findall(r"<loc>([^<]+)</loc>", body2))
            return url, sub_locs
        return url, locs
    return None, []


def filter_product_urls(urls: list[str], domain: str) -> list[str]:
    """Garde uniquement les URLs qui ressemblent à des fiches produits Presta."""
    out = []
    for u in urls:
        if domain not in u:
            continue
        # Patterns produits Presta : /{lang}/{cat}/{id}-{slug}.html ou /{cat}/{id}_{slug}
        if re.search(r"/\d+-[a-z0-9\-_]+\.html?$", u, re.IGNORECASE):
            out.append(u)
        elif re.search(r"\.html?$", u) and u.count("/") >= 4:
            # Heuristique : URL profonde se terminant en .html
            out.append(u)
    return out


# --- Inspection page produit ------------------------------------------

def inspect_product(url: str) -> None:
    print(f"\n  >>> Inspection produit : {url}")
    s, h, body = fetch(url)
    if s != 200 or not body:
        print(f"    HTTP {s} - skip")
        return
    print(f"    HTTP {s}, bytes={len(body)}")

    # JSON-LD
    ld_blocks = re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        body, flags=re.DOTALL | re.IGNORECASE,
    )
    product_ld = None
    for blk in ld_blocks:
        try:
            data = json.loads(blk.strip())
        except json.JSONDecodeError:
            continue
        candidates = data if isinstance(data, list) else (
            data.get("@graph", []) if isinstance(data, dict) and "@graph" in data else [data]
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
        print(f"    [OK ] JSON-LD Product schema trouvé. Clés : {list(product_ld.keys())}")
        for k in ("name", "sku", "description", "image", "offers", "category", "brand"):
            if k in product_ld:
                v = product_ld[k]
                preview = json.dumps(v, ensure_ascii=False)[:180] if isinstance(v, (dict, list)) else str(v)[:180]
                print(f"          .{k} = {preview!r}")
    else:
        print(f"    [KO ] Pas de JSON-LD Product")

    # Variantes (combinations PrestaShop)
    # PrestaShop 1.7+ : `var combinations = {...}` ou inline JSON dans la page
    print(f"\n    >>> Recherche des combinations (variantes Presta)")
    patterns = [
        (r'var\s+combinations\s*=\s*({[\s\S]+?});', "var combinations = {...};"),
        (r'data-product-combinations\s*=\s*["\']([^"\']+)["\']', "data-product-combinations attribut"),
        (r'"combinations"\s*:\s*({[\s\S]+?})\s*,\s*"', "key 'combinations' dans JSON inline"),
        (r'prestashop\.combinations\s*=\s*({[\s\S]+?});', "prestashop.combinations"),
    ]
    found_combinations = False
    for pat, label in patterns:
        m = re.search(pat, body)
        if m:
            print(f"    [OK ] {label}")
            raw = m.group(1)
            try:
                decoded = html.unescape(raw) if "&quot;" in raw else raw
                combos = json.loads(decoded)
                if isinstance(combos, dict):
                    print(f"          Combinations : {len(combos)} entries")
                    first_key = next(iter(combos), None)
                    if first_key:
                        print(f"          Première: id={first_key}, valeur preview = "
                              f"{json.dumps(combos[first_key], ensure_ascii=False)[:200]}")
                found_combinations = True
                break
            except json.JSONDecodeError as e:
                print(f"          parse failed: {e}")
    if not found_combinations:
        # Cherche un sélecteur de variantes (qui prouve qu'il y en a)
        has_form = re.search(r'<form[^>]+id=["\']add-to-cart-or-refresh["\']', body)
        has_attr_select = re.search(r'class="[^"]*product-variants[^"]*"', body)
        if has_form or has_attr_select:
            print(f"    [WARN] Form add-to-cart présent mais combinations JSON pas trouvé")
        else:
            print(f"    [INFO] Produit simple (pas de variantes apparentes)")

    # Sélecteurs HTML classiques Presta
    print(f"\n    >>> Sélecteurs HTML PrestaShop")
    selectors = [
        (r'<h1[^>]+class="[^"]*h1[^"]*"[^>]*>(.*?)</h1>', "h1.h1 (titre)"),
        (r'<h1[^>]+class="[^"]*product[^"]*"[^>]*>(.*?)</h1>', "h1.product (titre)"),
        (r'<div[^>]+itemprop="description"[^>]*>(.*?)</div>', "div[itemprop=description]"),
        (r'<div[^>]+class="[^"]*product-description[^"]*"[^>]*>(.*?)</div>', "div.product-description"),
        (r'<span[^>]+itemprop="price"[^>]*>(.*?)</span>', "span[itemprop=price]"),
        (r'<meta[^>]+itemprop="price"[^>]+content=["\']([^"\']+)["\']', "meta itemprop=price"),
    ]
    for pat, label in selectors:
        m = re.search(pat, body, re.DOTALL | re.IGNORECASE)
        if m:
            content = re.sub(r"<[^>]+>", " ", m.group(1)).strip()[:150]
            print(f"    [OK ] {label} : {content!r}")

    # Images
    print(f"\n    >>> Images")
    img_patterns = [
        r'<a[^>]+class="[^"]*thumb[^"]*"[^>]+href="([^"]+)"',
        r'<img[^>]+class="[^"]*product-cover[^"]*"[^>]+src="([^"]+)"',
        r'<img[^>]+id="[^"]*bigpic[^"]*"[^>]+src="([^"]+)"',
    ]
    imgs = set()
    for pat in img_patterns:
        for u in re.findall(pat, body):
            imgs.add(u)
    print(f"    {len(imgs)} URL images uniques trouvées")
    for u in list(imgs)[:5]:
        print(f"        - {u}")


# --- Main --------------------------------------------------------------

def probe(brand_url: str) -> None:
    print("\n" + "=" * 78)
    print(f"  {brand_url}")
    print("=" * 78)

    parsed = urlparse(brand_url if "://" in brand_url else f"https://{brand_url}")
    base = f"{parsed.scheme}://{parsed.netloc}"
    domain = parsed.netloc

    # 1. Home page + détection PrestaShop
    print("\n  [1] Détection PrestaShop dans la home")
    s, h, body = fetch(brand_url)
    print(f"      HTTP {s}, bytes={len(body)}")
    if s != 200:
        print(f"      Home inaccessible — skip")
        return
    signals = detect_prestashop(body, h)
    if signals:
        print(f"      Signatures détectées :")
        for k, v in signals.items():
            print(f"          {k}: {v}")
    else:
        print(f"      AUCUNE signature PrestaShop trouvée (le site n'est peut-être pas Presta)")

    # 2. Sitemap
    print(f"\n  [2] Recherche du sitemap")
    sm_url, urls = find_sitemap(base)
    if sm_url:
        print(f"      [OK ] sitemap : {sm_url}")
        print(f"      {len(urls)} URLs au total")
        product_urls = filter_product_urls(urls, domain)
        print(f"      {len(product_urls)} URLs filtrées comme fiches produits")
        if product_urls:
            for u in product_urls[:5]:
                print(f"          - {u}")
    else:
        print(f"      [KO ] aucun sitemap trouvé")
        product_urls = []

    # 3. Inspection d'une page produit (la première de la liste)
    print(f"\n  [3] Inspection d'une page produit (échantillon)")
    if product_urls:
        inspect_product(product_urls[0])
    else:
        print(f"      Aucun produit à inspecter — il faudra fournir manuellement une URL produit")


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage : python3 probe_prestashop.py <url1> [<url2> ...]")
        return 2
    for url in sys.argv[1:]:
        try:
            probe(url)
        except Exception as e:
            print(f"\n[FATAL] {url} : {type(e).__name__}: {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
