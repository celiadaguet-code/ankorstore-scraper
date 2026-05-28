#!/usr/bin/env python3
"""Probe SumUp Store — plateforme e-commerce hosted (boutiques sur {marque}.sumupstore.com).

Usage : python3 probe_sumup.py https://boos-drinks.sumupstore.com/ ...
"""
from __future__ import annotations
import gzip, html, json, re, ssl, sys, urllib.request, urllib.error
from urllib.parse import urlparse


UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


def fetch(url):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "gzip, identity",
    })
    try:
        with urllib.request.urlopen(req, timeout=25, context=ctx) as resp:
            raw = resp.read()
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                raw = gzip.decompress(raw)
            return resp.status, dict(resp.headers), raw.decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers or {}), ""
    except Exception as e:
        return 0, {}, f"ERROR: {type(e).__name__}: {e}"


def section(t):
    print(f"\n{'-'*78}\n  {t}\n{'-'*78}")


def probe(url):
    print(f"\n{'='*78}\n  {url}\n{'='*78}")
    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}"

    section("[1] Home : signatures SumUp")
    s, hdrs, body = fetch(url)
    print(f"  HTTP {s}, bytes={len(body)}")
    if s != 200:
        print("  inaccessible")
        return

    # Server / Powered-By
    for k in ("Server", "X-Powered-By", "X-Frame-Options"):
        if k in hdrs or k.lower() in hdrs:
            v = hdrs.get(k) or hdrs.get(k.lower())
            print(f"  header {k}: {v}")

    # Signatures
    for kw in ["sumup", "shop.sumup", "sumupstore", "merchant.sumup",
               "products.sumup", "shop-cdn.sumup", "go.sumup"]:
        n = body.lower().count(kw)
        if n > 0:
            print(f"  sig '{kw}' : {n} occurrences")

    # Meta generator
    m = re.search(r'<meta[^>]+name=["\']generator["\'][^>]+content=["\']([^"\']+)["\']', body, re.I)
    if m:
        print(f"  meta generator: {m.group(1)}")

    section("[2] Sitemap.xml")
    for sm_url in (f"{base}/sitemap.xml", f"{base}/sitemap_index.xml"):
        s2, _, body_sm = fetch(sm_url)
        if s2 != 200 or "<" not in body_sm[:100]:
            continue
        print(f"  [OK] {sm_url}")
        locs = re.findall(r"<loc>([^<]+)</loc>", body_sm)
        print(f"  {len(locs)} URLs")
        for u in locs[:8]:
            print(f"      {u}")
        # Si sitemap index, suit le sub-sitemap PRODUCTS pour avoir les URLs produits
        is_index = "<sitemapindex" in body_sm[:300] or any(
            "sitemap.products" in u or "sitemap-products" in u or "products-sitemap" in u
            for u in locs
        )
        if is_index:
            product_sub = next((u for u in locs if "product" in u.lower()), None)
            if product_sub:
                print(f"\n  Sub-sitemap products détecté : {product_sub}")
                s3, _, body3 = fetch(product_sub)
                if s3 == 200:
                    sub_locs = re.findall(r"<loc>([^<]+)</loc>", body3)
                    print(f"  {len(sub_locs)} URLs produits dans le sub-sitemap")
                    for u in sub_locs[:5]:
                        print(f"      {u}")
                    if sub_locs:
                        return inspect(sub_locs[0])
            return
        # Sitemap direct
        product_candidate = next(
            (u for u in locs
             if u != url and u != base + "/"
             and re.search(r"/[a-z0-9\-_]{3,}", urlparse(u).path)),
            None,
        )
        if product_candidate:
            return inspect(product_candidate)
        return

    section("[2bis] Pas de sitemap → exploration de la home pour URLs produits")
    # Cherche des liens internes
    internal_links = set()
    for m in re.finditer(r'href=["\']([^"\']+)["\']', body):
        href = m.group(1)
        if href.startswith("/"):
            internal_links.add(base + href)
        elif href.startswith(base):
            internal_links.add(href)
    print(f"  {len(internal_links)} liens internes")
    # Trie pour montrer URLs profondes (probablement produits)
    deep_links = [u for u in internal_links
                  if urlparse(u).path.strip("/").count("/") >= 1
                  or len(urlparse(u).path.strip("/")) > 4]
    deep_links = list(deep_links)[:10]
    for u in deep_links:
        print(f"      {u}")
    if deep_links:
        inspect(deep_links[0])


def inspect(url):
    section(f"[3] Inspection page : {url}")
    s, _, body = fetch(url)
    if s != 200 or not body:
        print(f"  HTTP {s}")
        return
    print(f"  HTTP {s}, bytes={len(body)}")

    # Save HTML
    fn = re.sub(r"[^a-zA-Z0-9._-]", "_", url.split("//")[-1])[:80] + ".html"
    out = f"debug_sumup_{fn}"
    with open(out, "w", encoding="utf-8") as f:
        f.write(body)
    print(f"  HTML dumpé : {out}")

    # JSON-LD
    blocks = re.findall(
        r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
        body, flags=re.DOTALL | re.IGNORECASE,
    )
    print(f"\n  >>> JSON-LD ({len(blocks)} blocs)")
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
                print(f"      @type: {t}")
                if t == "Product" or (isinstance(t, list) and "Product" in t):
                    print(f"      Clés Product : {list(c.keys())}")
                    for k in ("name", "sku", "description", "image", "offers", "brand"):
                        if k in c:
                            v = c[k]
                            preview = json.dumps(v, ensure_ascii=False)[:200] if isinstance(v, (dict, list)) else str(v)[:200]
                            print(f"        .{k} = {preview!r}")

    # OG
    print(f"\n  >>> Open Graph meta")
    for prop in ("og:title", "og:description", "og:image",
                 "product:price:amount", "product:price:currency"):
        m = re.search(r'<meta[^>]+(?:property|name)=["\']' + re.escape(prop) + r'["\'][^>]+content=["\']([^"\']+)["\']', body, re.I)
        if m:
            print(f"  [OK] {prop} = {m.group(1)[:150]!r}")

    # Microdata
    print(f"\n  >>> Microdata itemprop")
    for prop in ("price", "priceCurrency", "name", "description", "sku"):
        for pat in (
            rf'<[^>]*itemprop=["\']{prop}["\'][^>]*content=["\']([^"\']+)["\']',
            rf'<[^>]*content=["\']([^"\']+)["\'][^>]*itemprop=["\']{prop}["\']',
        ):
            m = re.search(pat, body, re.I)
            if m:
                print(f"  [OK] {prop}={m.group(1)[:150]!r}")
                break

    # Cherche les variables JS / data attributes notables
    print(f"\n  >>> Variables JS / data attributes")
    for pat, label in [
        (r'window\.__INITIAL_STATE__\s*=\s*({[\s\S]{0,30000}?})\s*;', "window.__INITIAL_STATE__"),
        (r'window\.__APP_DATA__\s*=\s*({[\s\S]{0,30000}?})\s*;', "window.__APP_DATA__"),
        (r'window\.app\s*=\s*({[\s\S]{0,30000}?})\s*;', "window.app"),
        (r'data-product=["\']({[^"\']+})["\']', "data-product attribute"),
    ]:
        m = re.search(pat, body)
        if m:
            print(f"  [OK] {label} ({len(m.group(1))} chars)")
            try:
                data = json.loads(m.group(1))
                if isinstance(data, dict):
                    print(f"        clés : {list(data.keys())[:15]}")
            except Exception:
                pass


def main():
    if len(sys.argv) < 2:
        print("Usage : python3 probe_sumup.py <url1> [...]")
        return 2
    for u in sys.argv[1:]:
        try:
            probe(u)
        except Exception as e:
            print(f"\n[FATAL] {u}: {type(e).__name__}: {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
