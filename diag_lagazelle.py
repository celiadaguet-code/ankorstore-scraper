#!/usr/bin/env python3
"""Diag : reproduit EXACTEMENT ce que fait http_get_json du scraper.

Pour comprendre pourquoi le scraper reçoit un body de 0 octets sur lagazelle.
"""
import urllib.request, gzip, ssl

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

URL = "https://lagazellemarrakchia.com/accueil/136-protege-passeport-safar.html"


def test(label, headers):
    print(f"\n=== {label} ===")
    print(f"  Headers : {headers}")
    req = urllib.request.Request(URL, headers=headers, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=20, context=ctx) as resp:
            raw = resp.read()
            print(f"  Status               : {resp.status}")
            print(f"  Content-Type         : {resp.headers.get('Content-Type')}")
            print(f"  Content-Encoding     : {resp.headers.get('Content-Encoding')}")
            print(f"  Content-Length hdr   : {resp.headers.get('Content-Length')}")
            print(f"  raw bytes received   : {len(raw)}")
            if not raw:
                print(f"  -> BODY VIDE !")
                return
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                try:
                    decoded = gzip.decompress(raw)
                    print(f"  After gunzip         : {len(decoded)} bytes")
                    text = decoded.decode("utf-8", errors="replace")
                    print(f"  has 'data-product='  : {'data-product=' in text}")
                    print(f"  has 'application/ld' : {'application/ld+json' in text}")
                except Exception as e:
                    print(f"  gunzip FAIL: {e}")
            else:
                text = raw.decode("utf-8", errors="replace")
                print(f"  has 'data-product='  : {'data-product=' in text}")
                print(f"  has 'application/ld' : {'application/ld+json' in text}")
    except Exception as e:
        print(f"  ERROR: {type(e).__name__}: {e}")


UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"

# Test 1: probe minimal (ce qui marchait avant)
test("PROBE (qui marchait)", {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Encoding": "gzip, identity",
})

# Test 2: scraper actuel (qui foire)
test("SCRAPER (qui foire)", {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/json;q=0.9,*/*;q=0.5",
    "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, identity",
})

# Test 3: identique au probe mais avec Accept-Language en plus
test("PROBE + Accept-Language", {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, identity",
})

# Test 4: identique au scraper mais SANS Accept-Language
test("SCRAPER - Accept-Language", {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/json;q=0.9,*/*;q=0.5",
    "Accept-Encoding": "gzip, identity",
})
