#!/usr/bin/env python3
"""
Probe Squarespace.

Pour chaque URL marque :
  1) Détection signatures Squarespace (Static.SQUARESPACE_CONTEXT, squarespace-cdn, etc.)
  2) Sitemap.xml → liste des URLs (produits sont sous /shop/ ou /store/ ou /products/)
  3) Sur UNE page produit : tester ?format=json (trick natif Squarespace)
  4) Vérifier la structure du JSON renvoyé (items, variants, attributes, etc.)

ZÉRO dépendance, stdlib only.

Usage :
    python3 probe_squarespace.py https://www.maisonfelicia.com/ https://www.floreenterre.com/
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
from urllib.parse import urlparse, urljoin


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
            return resp.status, resp.headers.get("Content-Type", ""), raw.decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        return e.code, "", body
    except Exception as e:
        return 0, "", f"ERROR: {type(e).__name__}: {e}"


def section(t):
    print(f"\n{'-'*78}\n  {t}\n{'-'*78}")


def detect_sqsp(body: str) -> dict:
    sigs = {}
    if "Static.SQUARESPACE_CONTEXT" in body or "SQUARESPACE_CONTEXT" in body:
        sigs["SQUARESPACE_CONTEXT"] = True
    if "squarespace-cdn.com" in body or "images.squarespace-cdn" in body:
        sigs["squarespace-cdn"] = True
    if "static.squarespace.com" in body:
        sigs["static.squarespace.com"] = True
    if re.search(r'<meta[^>]+name=["\']generator["\'][^>]+content=["\'][^"\']*Squarespace[^"\']*["\']', body, re.IGNORECASE):
        sigs["meta_generator"] = True
    return sigs


def probe(url: str) -> None:
    print(f"\n{'='*78}\n  {url}\n{'='*78}")
    parsed = urlparse(url if "://" in url else f"https://{url}")
    base = f"{parsed.scheme}://{parsed.netloc}"
    domain = parsed.netloc

    # 1. Home + signatures
    section("[1] Détection Squarespace dans la home")
    s, ct, body = fetch(url)
    print(f"  HTTP {s}, bytes={len(body)}, ct={ct}")
    if s != 200:
        print("  Home inaccessible — skip")
        return
    sigs = detect_sqsp(body)
    if sigs:
        for k, v in sigs.items():
            print(f"  signature : {k}: {v}")
    else:
        print("  AUCUNE signature Squarespace détectée")

    # 2. Sitemap
    section("[2] Sitemap.xml")
    s, ct, body_sm = fetch(f"{base}/sitemap.xml", accept="application/xml,*/*")
    if s != 200:
        print(f"  /sitemap.xml inaccessible (status={s})")
        return
    locs = re.findall(r"<loc>([^<]+)</loc>", body_sm)
    print(f"  {len(locs)} URLs dans sitemap")
    # Filtre URLs produits
    PRODUCT_PATTERNS = ("/shop/", "/store/", "/products/", "/product/", "/boutique/", "/produits/")
    product_urls = [u for u in locs if any(p in u for p in PRODUCT_PATTERNS)]
    print(f"  {len(product_urls)} URLs produits (filtre {PRODUCT_PATTERNS})")
    for u in product_urls[:5]:
        print(f"      {u}")

    if not product_urls:
        # Si pas de produits trouvés avec ces patterns, montre les premières URLs
        # pour comprendre la structure
        print(f"\n  Aucun match. 10 premières URLs du sitemap pour aide au diagnostic :")
        for u in locs[:10]:
            print(f"      {u}")
        return

    # 3. Test ?format=json sur une page produit
    target = product_urls[0]
    section(f"[3] Test ?format=json sur {target}")
    json_url = target + ("&" if "?" in target else "?") + "format=json"
    print(f"  GET {json_url}")
    s, ct, body_json = fetch(json_url, accept="application/json,*/*")
    print(f"  HTTP {s}, ct={ct}, bytes={len(body_json)}")
    if "json" not in ct.lower() and s == 200:
        # Squarespace renvoie parfois JSON même sans bon content-type
        if body_json.strip().startswith("{"):
            print(f"  body commence par {{ → on tente parse JSON quand même")
    if s == 200:
        try:
            data = json.loads(body_json)
        except json.JSONDecodeError as e:
            print(f"  [KO] parse JSON échoué : {e}")
            print(f"  Preview : {body_json[:500]!r}")
            return
        print(f"  [OK] JSON valide, {len(data)} clés top-level")
        print(f"  Clés : {list(data.keys())[:30]}")

        # On cherche les sous-objets utiles
        for k in ("item", "collection", "items", "websiteSettings"):
            if k in data:
                v = data[k]
                if isinstance(v, dict):
                    print(f"\n  .{k} (dict, {len(v)} clés) : {list(v.keys())[:20]}")
                    # Si c'est l'item produit, dump les clés intéressantes
                    if k == "item":
                        for p_key in ("title", "fullUrl", "description", "tags",
                                       "categories", "structuredContent",
                                       "variants", "items"):
                            if p_key in v:
                                pv = v[p_key]
                                preview = json.dumps(pv, ensure_ascii=False)[:200] if isinstance(pv, (dict, list)) else str(pv)[:200]
                                print(f"    .{p_key} = {preview!r}")
                elif isinstance(v, list):
                    print(f"\n  .{k} (list, {len(v)} éléments)")
                    if v and isinstance(v[0], dict):
                        print(f"    [0] keys : {list(v[0].keys())[:20]}")

        # Cherche les variants partout
        section("[3b] Recherche de variants dans le JSON")
        text = json.dumps(data, ensure_ascii=False)
        if '"variants"' in text:
            # Trouve l'array variants
            m = re.search(r'"variants"\s*:\s*\[([^\]]{0,3000})\]', text, flags=re.DOTALL)
            if m:
                # Parse first variant pour voir sa structure
                try:
                    variants_list = json.loads("[" + m.group(1) + "]")
                    print(f"  {len(variants_list)} variants trouvés")
                    if variants_list:
                        first = variants_list[0]
                        if isinstance(first, dict):
                            print(f"  Clés d'un variant : {list(first.keys())}")
                            print(f"  Exemple : {json.dumps(first, ensure_ascii=False)[:500]}")
                except Exception as e:
                    print(f"  Parse variants échoué : {e}")


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage : python3 probe_squarespace.py <url1> [<url2> ...]")
        return 2
    for url in sys.argv[1:]:
        try:
            probe(url)
        except Exception as e:
            print(f"\n[FATAL] {url} : {type(e).__name__}: {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
