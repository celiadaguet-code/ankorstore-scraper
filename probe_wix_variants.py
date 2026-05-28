#!/usr/bin/env python3
"""Probe : où sont planquées les variantes Wix dans le HTML d'une page produit ?

Usage : python3 probe_wix_variants.py https://www.savage.dog/accessoires/harnais-chien-pinkylicious
"""
from __future__ import annotations
import gzip, json, re, ssl, sys, urllib.request

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")


def fetch(url: str) -> str:
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "gzip, identity",
    })
    with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:
        raw = resp.read()
        if resp.headers.get("Content-Encoding", "").lower() == "gzip":
            raw = gzip.decompress(raw)
        return raw.decode("utf-8", errors="replace")


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage : python3 probe_wix_variants.py <url_produit_wix>")
        return 2
    url = sys.argv[1]
    print(f"Fetching {url}...")
    body = fetch(url)
    print(f"HTML : {len(body)} chars")

    # Save full HTML for inspection
    fn = "wix_variant_dump.html"
    with open(fn, "w", encoding="utf-8") as f:
        f.write(body)
    print(f"HTML dumpé : {fn}")

    # 1. Recherche des mots-clés "variant" / "options" / "choices"
    print("\n=== Mentions de mots-clés variant ===")
    for kw in ["variant", "choices", "options", "selectionId", "optionType",
               "productOptions", "inventoryItem", "Taille", "Size",
               "XS", '"S"', '"M"', '"L"']:
        n = body.count(kw)
        if n > 0:
            print(f"  '{kw}' : {n} occurrences")

    # 2. Scripts JSON inline (type application/json ou type ld+json)
    print("\n=== Scripts <script type=...> ===")
    for m in re.finditer(
        r'<script[^>]*type=["\']([^"\']+)["\'][^>]*(?:id=["\']([^"\']+)["\'])?[^>]*>(.*?)</script>',
        body, flags=re.DOTALL | re.IGNORECASE,
    ):
        stype = m.group(1)
        sid = m.group(2) or ""
        content = m.group(3).strip()
        if not content:
            continue
        # Affiche les scripts JSON-like avec leur ID
        if "json" in stype.lower() or stype in ("application/x-mathjax-config",):
            print(f"\n  [{stype}] id={sid!r} taille={len(content)}")
            # Si small, dump complet ; si gros, recherche keywords
            preview = content[:200]
            print(f"    preview : {preview!r}")
            # Cherche les keywords variant dans CE script
            for kw in ["variant", "options", "choices", "XS", "Taille", "Size"]:
                if kw in content:
                    idx = content.find(kw)
                    ctx_around = content[max(0,idx-50):idx+200]
                    print(f"    >>> contient '{kw}' : ...{ctx_around!r}...")
                    break

    # 3. window.* assignments contenant productOptions/variants
    print("\n=== window.* / var avec mention 'variant' ou 'options' ===")
    for m in re.finditer(
        r'(?:window\.|var\s+)(\w+)\s*=\s*({[\s\S]{0,50000}?});',
        body,
    ):
        varname, jsondata = m.group(1), m.group(2)
        if any(kw in jsondata for kw in ["variant", "productOptions", "choices",
                                          "XS", "Taille"]):
            print(f"\n  var/window {varname} ({len(jsondata)} chars)")
            # Cherche les keywords
            for kw in ["variant", "productOptions", "choices", "XS"]:
                if kw in jsondata:
                    idx = jsondata.find(kw)
                    print(f"    >>> '{kw}' à offset {idx}, contexte :")
                    print(f"    {jsondata[max(0,idx-100):idx+300]!r}")

    # 4. URLs _api/ qui peuvent renvoyer le détail produit complet
    print("\n=== URLs _api/ référencées ===")
    apis = set(re.findall(r'(/_api/[a-zA-Z0-9/\-_.]+)', body))
    for a in sorted(apis):
        print(f"  {a}")

    # 5. URLs externes avec wix.com / wixapps / parastorage
    print("\n=== URLs Wix API externes ===")
    wix_apis = set(re.findall(
        r'(https?://[a-z0-9-]+\.wix(?:apps)?\.com/[^\s"\'<>]+)',
        body,
    ))
    for a in list(wix_apis)[:15]:
        print(f"  {a}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
