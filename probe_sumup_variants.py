#!/usr/bin/env python3
"""Probe : où sont les variants dans un produit SumUp Store ?

Usage : python3 probe_sumup_variants.py https://nayana-shop.sumupstore.com/article/collier-50-creation
"""
from __future__ import annotations
import gzip, html, json, re, ssl, sys, urllib.request

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


def fetch(url):
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept-Encoding": "gzip, identity",
    })
    with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:
        raw = resp.read()
        if resp.headers.get("Content-Encoding", "").lower() == "gzip":
            raw = gzip.decompress(raw)
        return raw.decode("utf-8", errors="replace")


def main():
    if len(sys.argv) < 2:
        print("Usage : python3 probe_sumup_variants.py <url_produit_sumup>")
        return 1
    url = sys.argv[1]
    print(f"Fetching {url}...")
    body = fetch(url)
    print(f"Taille : {len(body)} chars")

    # Save full HTML
    with open("sumup_variant_dump.html", "w", encoding="utf-8") as f:
        f.write(body)
    print("Dumpé : sumup_variant_dump.html\n")

    # Recherche keywords variants
    print("=== Mentions de variants ===")
    for kw in ["variant", "variation", "options", "choices", "color", "couleur",
               "selectionId", "os-theme-product-variant", "os-theme-product-option"]:
        n = body.lower().count(kw.lower())
        if n > 0:
            print(f"  '{kw}' : {n}")

    # Cherche les data-selectors qui contiennent variant/option
    print("\n=== data-selector 'os-theme-product-*' uniques ===")
    selectors = set(re.findall(r'data-selector=["\'](os-theme-[a-z\-]+)["\']', body))
    for s in sorted(selectors):
        n = body.count(f'data-selector="{s}"') + body.count(f"data-selector='{s}'")
        print(f"  {s} ({n} occurrences)")

    # Cherche les blocs JSON inline
    print("\n=== Gros blocs JSON inline (scripts type=application/json) ===")
    for m in re.finditer(
        r'<script[^>]*type=["\']application/json["\'][^>]*(?:id=["\']([^"\']*)["\'])?[^>]*>([\s\S]*?)</script>',
        body, re.IGNORECASE,
    ):
        sid = m.group(1) or "(no id)"
        content = m.group(2).strip()
        print(f"\n  id={sid} ({len(content)} chars)")
        if not content:
            continue
        try:
            data = json.loads(content)
            if isinstance(data, dict):
                print(f"    clés : {list(data.keys())[:20]}")
            elif isinstance(data, list):
                print(f"    list de {len(data)}")
            # Cherche variants/options
            text = json.dumps(data, ensure_ascii=False) if isinstance(data, (dict, list)) else content
            for kw in ["variant", "option", "color", "couleur"]:
                if f'"{kw}' in text.lower()[:5000]:
                    idx = text.lower().find(f'"{kw}')
                    print(f"    contient '{kw}' à ~{idx} : ...{text[max(0,idx-20):idx+250]!r}...")
                    break
        except json.JSONDecodeError:
            pass

    # Cherche les <select> ou <input radio> de variantes
    print("\n=== Selects / radios dans la page ===")
    selects = re.findall(r'<select[^>]+name=["\']([^"\']+)["\'][^>]*>(.*?)</select>', body, re.DOTALL)
    print(f"  {len(selects)} <select>")
    for name, content in selects[:5]:
        options = re.findall(r'<option[^>]*value=["\']([^"\']+)["\']', content)
        print(f"    name={name!r} : {len(options)} options : {options[:6]}")

    # Cherche __NEXT_DATA__ (Next.js — souvent contient tout)
    print("\n=== __NEXT_DATA__ ===")
    m = re.search(r'<script[^>]*id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>', body, re.DOTALL)
    if m:
        content = m.group(1)
        print(f"  Trouvé ({len(content)} chars)")
        try:
            data = json.loads(content)
            # Navigue dans props.pageProps généralement
            print(f"  clés top : {list(data.keys())}")
            if "props" in data:
                pp = data["props"]
                if isinstance(pp, dict):
                    print(f"  .props.{list(pp.keys())[:10]}")
                    page_props = pp.get("pageProps")
                    if isinstance(page_props, dict):
                        print(f"  .props.pageProps clés : {list(page_props.keys())[:15]}")
                        # Cherche product
                        for k, v in page_props.items():
                            if k.lower() in ("product", "article", "item") and isinstance(v, dict):
                                print(f"\n  → .props.pageProps.{k} (Product/Article/Item) :")
                                for sk in list(v.keys())[:30]:
                                    print(f"      .{sk}")
                                # Variants ?
                                for vk in ("variants", "options", "selections"):
                                    if vk in v:
                                        vv = v[vk]
                                        print(f"\n  >>> .{vk} = (type {type(vv).__name__})")
                                        if isinstance(vv, list):
                                            print(f"      {len(vv)} elements")
                                            if vv:
                                                print(f"      [0] = {json.dumps(vv[0], ensure_ascii=False)[:500]}")
                                        elif isinstance(vv, dict):
                                            print(f"      keys: {list(vv.keys())[:15]}")
        except json.JSONDecodeError as e:
            print(f"  JSON parse fail: {e}")
    else:
        print("  Pas de __NEXT_DATA__")


if __name__ == "__main__":
    sys.exit(main())
