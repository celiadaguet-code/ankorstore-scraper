#!/usr/bin/env python3
"""
Test des endpoints possibles pour récupérer une variation WooCommerce.

On essaie 6 méthodes sur 1 variation connue (cosmella.fr / variation #5174 / parent #5153) :
  A) /wp-json/wc/store/v1/products/{id}                                — direct ID
  B) /wp-json/wc/store/v1/products?type=variation&include=...          — listing filtré
  C) /wp-json/wc/store/v1/products/{parent_id}                         — parent endpoint
  D) /wp-json/wc/store/products/{id}                                   — legacy
  E) Scraper la page parent et extraire data-product_variations        — HTML scraping
  F) /wp-json/wp/v2/product_variation/{id}                             — WP REST (souvent auth)

ZÉRO dépendance.
"""

from __future__ import annotations

import gzip
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

# Cible : cosmella.fr, variation 5174 = "Recharge parfumée - Bélier" (parent 5153)
BASE = "https://cosmella.fr"
PARENT_ID = 5153
VARIATION_ID = 5174
PARENT_PERMALINK = "https://cosmella.fr/boutique/recharge-parfumee-pour-objets-souvenirs/"


def _ssl_ctx():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def fetch(url: str, accept: str = "application/json"):
    print(f"\n>>> {url}")
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": accept,
        "Accept-Encoding": "gzip, identity",
    })
    try:
        with urllib.request.urlopen(req, timeout=20, context=_ssl_ctx()) as resp:
            raw = resp.read()
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                raw = gzip.decompress(raw)
            text = raw.decode("utf-8", errors="replace")
            ct = resp.headers.get("Content-Type", "")
            status = resp.status
            print(f"    HTTP {status}   bytes={len(raw)}   ct={ct}")
            return status, ct, text
    except urllib.error.HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")
        except Exception:
            body = ""
        print(f"    HTTP {e.code} {e.reason}   body[:200]={body[:200]!r}")
        return e.code, "", body
    except Exception as e:
        print(f"    ERROR: {type(e).__name__}: {e}")
        return 0, "", ""


def show_json_preview(text: str, max_keys: int = 12) -> None:
    try:
        data = json.loads(text)
    except Exception:
        print("    (réponse non-JSON)")
        print(f"    body[:300]={text[:300]!r}")
        return
    if isinstance(data, dict):
        keys = list(data.keys())[:max_keys]
        print(f"    JSON dict avec {len(data)} clés. Premières : {keys}")
        for k in ("id", "name", "sku", "price", "prices", "stock_quantity", "is_in_stock", "attributes"):
            if k in data:
                v = data[k]
                if isinstance(v, (dict, list)):
                    print(f"    .{k} = (preview) {json.dumps(v, ensure_ascii=False)[:200]}")
                else:
                    print(f"    .{k} = {v!r}")
    elif isinstance(data, list):
        print(f"    JSON list de {len(data)} éléments")
        if data:
            first = data[0]
            if isinstance(first, dict):
                print(f"    [0] keys = {list(first.keys())[:max_keys]}")


def main() -> int:
    print("=" * 78)
    print(f"  Cible : variation #{VARIATION_ID} (parent #{PARENT_ID}) sur {BASE}")
    print("=" * 78)

    # A) Direct ID
    print("\n[A] Direct ID dans Store API v1")
    s, ct, t = fetch(f"{BASE}/wp-json/wc/store/v1/products/{VARIATION_ID}")
    if s == 200 and "json" in ct.lower():
        show_json_preview(t)

    # B) Listing avec type=variation
    print("\n[B] Listing avec ?type=variation&include=...")
    s, ct, t = fetch(f"{BASE}/wp-json/wc/store/v1/products?type=variation&include={VARIATION_ID}")
    if s == 200 and "json" in ct.lower():
        show_json_preview(t)

    # B-bis) Listing avec catalog_visibility ou autre flag
    print("\n[B'] Listing ?include avec per_page=1")
    s, ct, t = fetch(f"{BASE}/wp-json/wc/store/v1/products?include={VARIATION_ID}&per_page=1&orderby=include&catalog_visibility=hidden")
    if s == 200 and "json" in ct.lower():
        show_json_preview(t)

    # C) Parent endpoint — vérifie si les variations sont incluses en mode détaillé
    print("\n[C] Parent endpoint single product")
    s, ct, t = fetch(f"{BASE}/wp-json/wc/store/v1/products/{PARENT_ID}")
    if s == 200 and "json" in ct.lower():
        try:
            data = json.loads(t)
            if isinstance(data, dict):
                variations = data.get("variations") or []
                print(f"    .variations = {len(variations)} elements")
                if variations:
                    print(f"    .variations[0] keys = {list(variations[0].keys()) if isinstance(variations[0], dict) else type(variations[0])}")
                    print(f"    .variations[0] preview = {json.dumps(variations[0], ensure_ascii=False)[:400]}")
        except Exception:
            pass

    # D) Legacy endpoint
    print("\n[D] Legacy /wp-json/wc/store/products/{id}")
    s, ct, t = fetch(f"{BASE}/wp-json/wc/store/products/{VARIATION_ID}")
    if s == 200 and "json" in ct.lower():
        show_json_preview(t)

    # E) HTML scraping de la page parent — chercher data-product_variations
    print("\n[E] Page parent en HTML, extraction data-product_variations")
    s, ct, t = fetch(PARENT_PERMALINK, accept="text/html,application/xhtml+xml")
    if s == 200:
        # Cherche l'attribut data-product_variations sur form.variations_form
        # Souvent encodé en HTML entities, on récupère puis on décode
        m = re.search(
            r'data-product_variations\s*=\s*(?:"|&quot;)(.+?)(?:"|&quot;)\s*(?:>|\s)',
            t,
            flags=re.DOTALL,
        )
        if m:
            raw_attr = m.group(1)
            # Décoder les entités HTML les plus courantes
            decoded = (raw_attr
                       .replace("&quot;", '"')
                       .replace("&#039;", "'")
                       .replace("&amp;", "&")
                       .replace("&lt;", "<")
                       .replace("&gt;", ">"))
            try:
                arr = json.loads(decoded)
                print(f"    -> trouvé ! Liste de {len(arr)} variations")
                if arr:
                    v0 = arr[0]
                    print(f"       [0] keys: {list(v0.keys())[:20]}")
                    for k in ("variation_id", "sku", "display_price", "display_regular_price",
                              "is_in_stock", "min_qty", "max_qty", "attributes", "image"):
                        if k in v0:
                            v = v0[k]
                            if isinstance(v, (dict, list)):
                                print(f"       [0].{k} = (preview) {json.dumps(v, ensure_ascii=False)[:200]}")
                            else:
                                print(f"       [0].{k} = {v!r}")
            except Exception as e:
                print(f"    -> trouvé mais parse échoué : {e}")
                print(f"    preview du raw : {decoded[:300]!r}")
        else:
            print("    -> attribut data-product_variations introuvable dans le HTML")
            # Cherche d'autres patterns possibles
            if "variations_form" in t:
                print("    (form variations_form trouvé, structure inhabituelle)")
            else:
                print("    (pas de variations_form non plus — thème non-standard)")

    # F) WP REST product_variation
    print("\n[F] /wp-json/wp/v2/product_variation/{id}")
    s, ct, t = fetch(f"{BASE}/wp-json/wp/v2/product_variation/{VARIATION_ID}")
    if s == 200 and "json" in ct.lower():
        show_json_preview(t)

    print("\n" + "=" * 78)
    print(" Termé. La méthode qui donne le plus d'info (prix variation + sku + stock)")
    print(" sera celle qu'on utilise dans le vrai scraper.")
    print("=" * 78)
    return 0


if __name__ == "__main__":
    sys.exit(main())
