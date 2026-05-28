#!/usr/bin/env python3
"""
Dump complet d'un catalogue WooCommerce via Store API v1.

Récupère TOUS les produits d'une boutique (paginé), et, pour les produits
variables, fetche aussi les détails de chaque variation. Tout est sauvegardé
en JSON brut localement pour inspection.

Usage :
    python3 dump_catalog.py https://cosmella.fr
    python3 dump_catalog.py https://cosmella.fr --max 20    # limite pour test

Output :
    catalog_<domaine>.json   — JSON complet avec produits + variations
    catalog_<domaine>.summary.txt — résumé lisible

ZÉRO dépendance — stdlib uniquement.
"""

from __future__ import annotations

import gzip
import json
import re
import ssl
import sys
import time
import urllib.request
import urllib.error
from urllib.parse import urlparse

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
TIMEOUT = 30
CRAWL_DELAY = 1.2  # secondes entre requêtes (cosmella.fr demande 3s mais c'est du legacy,
                   # on reste poli avec 1.2s pour ne pas trop ralentir)


# --- SSL fallback macOS Homebrew ------------------------------------------

def _build_ssl_context(verify: bool):
    if verify:
        return ssl.create_default_context()
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


_SSL_OK: bool | None = None


def _ssl_works() -> bool:
    global _SSL_OK
    if _SSL_OK is not None:
        return _SSL_OK
    try:
        req = urllib.request.Request("https://www.google.com/", headers={"User-Agent": USER_AGENT})
        urllib.request.urlopen(req, timeout=5, context=_build_ssl_context(True)).close()
        _SSL_OK = True
    except Exception:
        _SSL_OK = False
        print("[INFO] SSL non vérifié pour ce run (bundle CA absent).", flush=True)
    return _SSL_OK


def _http_get(url: str) -> tuple[int, dict, bytes]:
    """Retourne (status, headers_dict, body_bytes). Lève Exception en cas d'erreur réseau."""
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "application/json",
        "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, identity",
    }
    req = urllib.request.Request(url, headers=headers, method="GET")
    ctx = _build_ssl_context(verify=_ssl_works())
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT, context=ctx) as resp:
            raw = resp.read()
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                try:
                    raw = gzip.decompress(raw)
                except Exception:
                    pass
            hdrs = {k: v for k, v in resp.headers.items()}
            return resp.status, hdrs, raw
    except urllib.error.HTTPError as e:
        raw = b""
        try:
            raw = e.read()
        except Exception:
            pass
        return e.code, dict(e.headers or {}), raw


def fetch_json(url: str, retries: int = 3) -> tuple[int, dict, list | dict | None]:
    """Retourne (status, headers, parsed_json_or_None). Retry sur 5xx / réseau."""
    for attempt in range(retries):
        try:
            status, hdrs, raw = _http_get(url)
            if 500 <= status < 600 and attempt < retries - 1:
                wait = 2 ** attempt
                print(f"   ! HTTP {status}, retry dans {wait}s...", flush=True)
                time.sleep(wait)
                continue
            try:
                data = json.loads(raw.decode("utf-8", errors="replace")) if raw else None
            except json.JSONDecodeError:
                data = None
            return status, hdrs, data
        except Exception as e:
            if attempt < retries - 1:
                wait = 2 ** attempt
                print(f"   ! erreur réseau {type(e).__name__}: {e}, retry dans {wait}s...", flush=True)
                time.sleep(wait)
                continue
            raise
    return 0, {}, None


# --- Logique principale ---------------------------------------------------

def safe_filename(s: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]", "_", s)


def dump(brand_url: str, max_products: int | None = None) -> None:
    parsed = urlparse(brand_url if "://" in brand_url else f"https://{brand_url}")
    base = f"{parsed.scheme}://{parsed.netloc}"
    domain = parsed.netloc.replace("www.", "")
    out_path = f"catalog_{safe_filename(domain)}.json"
    summary_path = f"catalog_{safe_filename(domain)}.summary.txt"

    print(f"\n>>> Dump du catalogue {base}\n")

    # Étape 1 : récupère le total
    print("[1/3] Récupération du total via X-WP-Total...")
    status, hdrs, _ = fetch_json(f"{base}/wp-json/wc/store/v1/products?per_page=1")
    total = int(hdrs.get("X-WP-Total") or hdrs.get("x-wp-total") or 0)
    print(f"      {total} produits déclarés par le serveur")
    if max_products:
        total = min(total, max_products)
        print(f"      (limité à {total} pour ce dump)")

    # Étape 2 : tout pull en pages de 100
    per_page = 100
    page = 1
    products: list[dict] = []
    while len(products) < total:
        url = f"{base}/wp-json/wc/store/v1/products?per_page={per_page}&page={page}"
        print(f"[2/3] Page {page} (per_page={per_page})...", end=" ", flush=True)
        status, hdrs, data = fetch_json(url)
        if not isinstance(data, list):
            print(f"FAIL (status={status})")
            break
        products.extend(data)
        print(f"+{len(data)} produits (total cumulé : {len(products)})")
        if len(data) < per_page:
            break
        page += 1
        time.sleep(CRAWL_DELAY)
        if max_products and len(products) >= max_products:
            break

    if max_products:
        products = products[:max_products]

    # Étape 3 : fetch détails des variations pour chaque produit variable
    variation_details: dict[int, dict] = {}  # variation_id -> objet variation complet
    variable_products = [p for p in products if p.get("type") == "variable"]
    print(f"\n[3/3] Détails des variations : {len(variable_products)} produits variables à explorer")

    # Pour chaque produit variable, on collecte les variation IDs
    variation_ids: list[int] = []
    for p in variable_products:
        for v in (p.get("variations") or []):
            vid = v.get("id") if isinstance(v, dict) else v
            if isinstance(vid, int):
                variation_ids.append(vid)
    variation_ids = list(dict.fromkeys(variation_ids))  # dédoublonne en gardant l'ordre
    print(f"      {len(variation_ids)} IDs de variations uniques à fetcher")

    # Fetch par batchs avec ?include=
    BATCH = 25
    for i in range(0, len(variation_ids), BATCH):
        batch = variation_ids[i:i + BATCH]
        url = f"{base}/wp-json/wc/store/v1/products?include={','.join(map(str, batch))}&per_page={len(batch)}"
        print(f"      batch {i // BATCH + 1}/{(len(variation_ids) + BATCH - 1) // BATCH}...", end=" ", flush=True)
        status, hdrs, data = fetch_json(url)
        if isinstance(data, list):
            for obj in data:
                if isinstance(obj, dict) and "id" in obj:
                    variation_details[obj["id"]] = obj
            print(f"+{len(data)} variations")
        else:
            print(f"FAIL (status={status})")
        time.sleep(CRAWL_DELAY)

    # Sauvegarde
    payload = {
        "brand_url": brand_url,
        "base": base,
        "total_declared": total,
        "products": products,
        "variations": variation_details,
    }
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
    print(f"\n[OK] JSON brut sauvegardé : {out_path}  ({len(products)} produits, {len(variation_details)} variations)")

    # Résumé lisible
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write(f"Catalogue {base}\n")
        f.write(f"=========================================\n")
        f.write(f"Produits dumped : {len(products)} / {total} déclarés\n")
        f.write(f"  - simples     : {sum(1 for p in products if p.get('type') == 'simple')}\n")
        f.write(f"  - variables   : {sum(1 for p in products if p.get('type') == 'variable')}\n")
        f.write(f"  - autres      : {sum(1 for p in products if p.get('type') not in ('simple', 'variable'))}\n")
        f.write(f"Variations fetchées : {len(variation_details)}\n\n")

        # Mots-clés ateliers ?
        ATELIER_KW = ("atelier", "workshop", "consultation", "rendez-vous", "rdv",
                      "gift card", "chèque cadeau", "bon cadeau")
        n_atelier = sum(1 for p in products if any(kw in (p.get("name") or "").lower() for kw in ATELIER_KW))
        f.write(f"Produits avec mot-clé 'atelier/workshop/etc.' dans le nom : {n_atelier}\n\n")

        # Echantillons
        f.write("=== 5 PREMIERS PRODUITS SIMPLES ===\n")
        for p in [x for x in products if x.get("type") == "simple"][:5]:
            f.write(f"  #{p.get('id')} {p.get('name')!r}\n")
            f.write(f"    sku={p.get('sku')}, price={(p.get('prices') or {}).get('price')}, "
                    f"stock={p.get('is_in_stock')}, n_imgs={len(p.get('images') or [])}\n")
            cats = [c.get("name") for c in (p.get("categories") or [])]
            f.write(f"    categories={cats}\n")
        f.write("\n=== 5 PREMIERS PRODUITS VARIABLES ===\n")
        for p in [x for x in products if x.get("type") == "variable"][:5]:
            f.write(f"  #{p.get('id')} {p.get('name')!r}  n_variations={len(p.get('variations') or [])}\n")
            attrs = p.get("attributes") or []
            f.write(f"    attributes (names): {[a.get('name') for a in attrs if isinstance(a, dict)]}\n")
            for v in (p.get("variations") or [])[:3]:
                vid = v.get("id") if isinstance(v, dict) else v
                vd = variation_details.get(vid)
                if vd:
                    v_attrs = [(a.get("name"), a.get("value")) for a in (vd.get("attributes") or []) if isinstance(a, dict)]
                    f.write(f"    -> variation #{vid}  sku={vd.get('sku')}  "
                            f"price={(vd.get('prices') or {}).get('price')}  attrs={v_attrs}\n")
                else:
                    f.write(f"    -> variation #{vid}  (détails non récupérés)\n")
    print(f"[OK] Résumé lisible : {summary_path}")


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage : python3 dump_catalog.py <url> [--max N]")
        return 2
    url = sys.argv[1]
    max_p = None
    if "--max" in sys.argv:
        i = sys.argv.index("--max")
        max_p = int(sys.argv[i + 1])
    dump(url, max_products=max_p)
    return 0


if __name__ == "__main__":
    sys.exit(main())
