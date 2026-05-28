#!/usr/bin/env python3
"""Diag : où est planquée la description dans le JSON Squarespace ?

Fetch ?format=json sur une URL produit et dump les champs candidats.
"""
import gzip, json, re, ssl, sys, urllib.request

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
    with urllib.request.urlopen(req, timeout=25, context=ctx) as resp:
        raw = resp.read()
        if resp.headers.get("Content-Encoding", "").lower() == "gzip":
            raw = gzip.decompress(raw)
        return raw.decode("utf-8", errors="replace")


def main():
    if len(sys.argv) < 2:
        print("Usage : python3 diag_squarespace_desc.py <url_produit>")
        return 1
    url = sys.argv[1]
    json_url = url + ("&" if "?" in url else "?") + "format=json"
    print(f"Fetching {json_url}...")
    body = fetch(json_url)
    print(f"Reçu : {len(body)} chars")

    try:
        data = json.loads(body)
    except json.JSONDecodeError as e:
        print(f"JSON parse failed: {e}")
        return 1

    item = data.get("item", {})
    if not isinstance(item, dict):
        print("Pas d'objet 'item' dans le JSON")
        return 1

    print(f"\nClés de item : {list(item.keys())}")

    # Dump tous les champs qui pourraient être de la description
    candidates = [
        "body", "excerpt", "description", "summary",
        "structuredContent",  # objet potentiellement contenant description
    ]
    for k in candidates:
        v = item.get(k)
        if v is None or v == "":
            continue
        if isinstance(v, (dict, list)):
            print(f"\n  item.{k} = (dict/list, {len(v)} keys/items)")
            if isinstance(v, dict):
                # Print sub-keys
                for sk in v.keys():
                    sv = v[sk]
                    preview = (str(sv)[:200] if isinstance(sv, str)
                               else f"({type(sv).__name__}, len={len(sv) if hasattr(sv, '__len__') else '?'})")
                    print(f"      .{sk} = {preview}")
        else:
            preview = str(v)[:500]
            print(f"\n  item.{k} ({len(str(v))} chars) = {preview!r}")

    # Bonus : dump full item JSON (truncated) pour inspection
    print(f"\n\n--- item full preview ---")
    print(json.dumps(item, ensure_ascii=False, indent=2)[:3000])


if __name__ == "__main__":
    sys.exit(main())
