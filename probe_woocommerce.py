#!/usr/bin/env python3
"""
Probe WooCommerce — détection de la meilleure méthode de scraping pour une marque.

Usage :
    python3 probe_woocommerce.py https://cosmella.fr
    python3 probe_woocommerce.py https://cosmella.fr https://autre-marque.fr ...

Le script teste, pour chaque URL :
  1) Store API v1            /wp-json/wc/store/v1/products  (la meilleure)
  2) Store API legacy        /wp-json/wc/store/products
  3) WP REST API root        /wp-json/                       (discovery)
  4) Sitemap index           /sitemap_index.xml              (fallback URLs produits)
  5) Sitemap alternatif      /sitemap.xml
  6) robots.txt              /robots.txt                     (info)

Il affiche un rapport clair par marque + un récap final, et écrit un fichier
"probe_report.json" à côté du script avec les résultats bruts (utile pour la
construction du scraper).

ZÉRO dépendance — utilise uniquement la stdlib Python 3.
"""

from __future__ import annotations

import gzip
import io
import json
import re
import ssl
import sys
import time
import urllib.request
import urllib.error
from dataclasses import dataclass, field, asdict
from typing import Any
from urllib.parse import urlparse

USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)
TIMEOUT = 15
REQUEST_GAP = 0.8  # secondes entre 2 requêtes (politesse / anti rate-limit)


# --- HTTP helper (stdlib only) ---------------------------------------------

# --- Contexte SSL : tente d'abord avec vérif, sinon bascule sans vérif ----
#
# Python 3.14 installé via Homebrew sur macOS n'a pas accès au bundle CA système
# par défaut, ce qui fait échouer la vérif SSL. Pour un probe en read-only sur
# des sites publics, on accepte de désactiver la vérif si nécessaire.
def _build_ssl_context(verify: bool):
    if verify:
        return ssl.create_default_context()
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


# Détecte UNE fois si la vérif SSL fonctionne ; sinon on bascule pour toute la session
_SSL_VERIFY_OK: bool | None = None


def _ssl_works() -> bool:
    global _SSL_VERIFY_OK
    if _SSL_VERIFY_OK is not None:
        return _SSL_VERIFY_OK
    try:
        req = urllib.request.Request(
            "https://www.google.com/",
            headers={"User-Agent": USER_AGENT},
        )
        urllib.request.urlopen(req, timeout=5, context=_build_ssl_context(True)).close()
        _SSL_VERIFY_OK = True
    except Exception:
        _SSL_VERIFY_OK = False
        print(
            "[INFO] Vérification SSL indisponible (bundle CA absent). "
            "Bascule en mode sans vérification SSL pour ce run.\n"
            "       Pour corriger proprement : /Applications/Python\\ 3.x/Install\\ Certificates.command "
            "ou bien `python3 -m pip install --break-system-packages certifi`.",
            flush=True,
        )
    return _SSL_VERIFY_OK


def fetch(url: str, accept_json: bool = False) -> dict[str, Any]:
    """GET stdlib, retourne dict normalisé. Gère SSL macOS Homebrew."""
    headers = {
        "User-Agent": USER_AGENT,
        "Accept": "application/json,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
        "Accept-Encoding": "gzip, identity",
    }
    if accept_json:
        headers["Accept"] = "application/json"

    out: dict[str, Any] = {
        "url": url,
        "ok": False,
        "status": None,
        "content_type": None,
        "bytes": 0,
        "body_preview": None,
        "json": None,
        "headers": {},
        "error": None,
        "final_url": None,
    }

    ctx = _build_ssl_context(verify=_ssl_works())

    req = urllib.request.Request(url, headers=headers, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT, context=ctx) as resp:
            raw = resp.read()
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                try:
                    raw = gzip.decompress(raw)
                except Exception:
                    pass
            out["status"] = resp.status
            out["content_type"] = resp.headers.get("Content-Type", "")
            out["bytes"] = len(raw)
            out["ok"] = 200 <= resp.status < 300
            out["final_url"] = resp.geturl()
            for h in ("X-WP-Total", "X-WP-TotalPages", "Server", "X-Powered-By", "CF-RAY"):
                v = resp.headers.get(h) or resp.headers.get(h.lower())
                if v:
                    out["headers"][h] = v
            try:
                text = raw.decode("utf-8", errors="replace")
            except Exception:
                text = ""
            out["body_preview"] = text[:1500]
            ct = (out["content_type"] or "").lower()
            if accept_json and out["ok"] and "json" in ct:
                try:
                    out["json"] = json.loads(text)
                except Exception as e:
                    out["error"] = f"JSON decode: {e}"
    except urllib.error.HTTPError as e:
        out["status"] = e.code
        try:
            raw = e.read()
            if e.headers.get("Content-Encoding", "").lower() == "gzip":
                try:
                    raw = gzip.decompress(raw)
                except Exception:
                    pass
            out["body_preview"] = raw[:500].decode("utf-8", errors="replace")
            out["bytes"] = len(raw)
        except Exception:
            pass
        out["content_type"] = e.headers.get("Content-Type", "") if e.headers else ""
        out["error"] = f"HTTP {e.code} {e.reason}"
    except urllib.error.URLError as e:
        out["error"] = f"URL error: {e.reason}"
    except TimeoutError:
        out["error"] = f"Timeout >{TIMEOUT}s"
    except Exception as e:
        out["error"] = f"{type(e).__name__}: {e}"
    return out


# --- Helpers ----------------------------------------------------------------

def normalize_base(url: str) -> str:
    """Retourne la racine du domaine (sans path, sans trailing slash)."""
    parsed = urlparse(url if "://" in url else f"https://{url}")
    return f"{parsed.scheme}://{parsed.netloc}"


def fetch_full_text(url: str) -> str:
    """Fetch et retourne le body texte complet (sans cap), pour parser les sitemaps."""
    headers = {"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, identity"}
    req = urllib.request.Request(url, headers=headers)
    ctx = _build_ssl_context(verify=_ssl_works())
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT, context=ctx) as resp:
            raw = resp.read()
            if resp.headers.get("Content-Encoding", "").lower() == "gzip":
                try:
                    raw = gzip.decompress(raw)
                except Exception:
                    pass
            return raw.decode("utf-8", errors="replace")
    except Exception:
        return ""


# --- Result dataclass -------------------------------------------------------

@dataclass
class ProbeResult:
    brand_url: str
    base: str
    tier1_store_v1: dict = field(default_factory=dict)
    tier1_v1_products_sample: list = field(default_factory=list)
    tier1_v1_estimated_total: int | None = None
    tier2_store_legacy: dict = field(default_factory=dict)
    wp_json_root: dict = field(default_factory=dict)
    sitemap_index: dict = field(default_factory=dict)
    sitemap_root: dict = field(default_factory=dict)
    product_sitemap_url: str | None = None
    product_urls_count: int | None = None
    product_urls_sample: list = field(default_factory=list)
    robots: dict = field(default_factory=dict)
    recommended_tier: str | None = None
    notes: list = field(default_factory=list)


# --- Tests individuels ------------------------------------------------------

def probe_store_v1(base: str, result: ProbeResult) -> None:
    url = f"{base}/wp-json/wc/store/v1/products?per_page=3"
    r = fetch(url, accept_json=True)
    # Total via header
    if r.get("headers", {}).get("X-WP-Total"):
        try:
            result.tier1_v1_estimated_total = int(r["headers"]["X-WP-Total"])
        except ValueError:
            pass
    json_payload = r.pop("json", None)
    result.tier1_store_v1 = r
    if r.get("ok") and isinstance(json_payload, list):
        for p in json_payload[:3]:
            result.tier1_v1_products_sample.append({
                "id": p.get("id"),
                "name": p.get("name"),
                "type": p.get("type"),
                "sku": p.get("sku"),
                "price": (p.get("prices") or {}).get("price"),
                "currency": (p.get("prices") or {}).get("currency_code"),
                "is_in_stock": p.get("is_in_stock"),
                "has_options": bool(p.get("variations")),
                "n_variations": len(p.get("variations") or []),
                "n_images": len(p.get("images") or []),
                "permalink": p.get("permalink"),
            })
    time.sleep(REQUEST_GAP)


def probe_store_legacy(base: str, result: ProbeResult) -> None:
    url = f"{base}/wp-json/wc/store/products?per_page=1"
    r = fetch(url, accept_json=True)
    json_payload = r.pop("json", None)
    result.tier2_store_legacy = r
    if isinstance(json_payload, list):
        result.tier2_store_legacy["json_was_list"] = True
        result.tier2_store_legacy["sample_first_keys"] = list(json_payload[0].keys())[:15] if json_payload else []
    elif json_payload is not None:
        result.tier2_store_legacy["json_was_list"] = False
    time.sleep(REQUEST_GAP)


def probe_wp_root(base: str, result: ProbeResult) -> None:
    r = fetch(f"{base}/wp-json/", accept_json=True)
    routes = []
    if r.get("json") and isinstance(r["json"], dict):
        all_routes = r["json"].get("routes") or {}
        for route in all_routes:
            if any(k in route for k in ("/wc/", "/wp/v2/product", "/store/")):
                routes.append(route)
    result.wp_json_root = {
        "status": r.get("status"),
        "ok": r.get("ok"),
        "error": r.get("error"),
        "product_related_routes": routes[:30],
        "n_product_related_routes": len(routes),
    }
    time.sleep(REQUEST_GAP)


def probe_sitemap_index(base: str, result: ProbeResult) -> None:
    url = f"{base}/sitemap_index.xml"
    r = fetch(url)
    body_preview = r.get("body_preview") or ""
    r_clean = {k: v for k, v in r.items() if k not in ("body_preview", "json")}
    result.sitemap_index = r_clean
    if r.get("ok") and body_preview:
        # Si le preview a été tronqué, refetch complet
        full = body_preview
        if r.get("bytes", 0) > 1500:
            full = fetch_full_text(url) or body_preview
        m = re.findall(r"<loc>([^<]+)</loc>", full)
        product_sm = next((u for u in m if "product" in u.lower()), None)
        result.sitemap_index["sub_sitemaps_preview"] = m[:10]
        if product_sm:
            result.product_sitemap_url = product_sm
            full_p = fetch_full_text(product_sm)
            locs = re.findall(r"<loc>([^<]+)</loc>", full_p)
            result.product_urls_count = len(locs)
            result.product_urls_sample = locs[:5]
            time.sleep(REQUEST_GAP)
    time.sleep(REQUEST_GAP)


def probe_sitemap_root(base: str, result: ProbeResult) -> None:
    url = f"{base}/sitemap.xml"
    r = fetch(url)
    body_preview = r.get("body_preview") or ""
    r_clean = {k: v for k, v in r.items() if k not in ("body_preview", "json")}
    result.sitemap_root = r_clean
    if r.get("ok") and body_preview and not result.product_sitemap_url:
        # Cherche un loc qui ressemble à un sub-sitemap product
        m = re.findall(r"<loc>([^<]+)</loc>", body_preview)
        product_sm = next((u for u in m if "product" in u.lower()), None)
        if product_sm:
            result.product_sitemap_url = product_sm
            full_p = fetch_full_text(product_sm)
            locs = re.findall(r"<loc>([^<]+)</loc>", full_p)
            result.product_urls_count = len(locs)
            result.product_urls_sample = locs[:5]
            time.sleep(REQUEST_GAP)
        elif m:
            # Parfois sitemap.xml liste directement les URLs produits
            result.sitemap_root["locs_preview"] = m[:10]
    time.sleep(REQUEST_GAP)


def probe_robots(base: str, result: ProbeResult) -> None:
    url = f"{base}/robots.txt"
    r = fetch(url)
    result.robots = {
        "status": r.get("status"),
        "ok": r.get("ok"),
        "error": r.get("error"),
        "preview": (r.get("body_preview") or "")[:400],
    }
    time.sleep(REQUEST_GAP)


# --- Recommandation ---------------------------------------------------------

def compute_recommendation(r: ProbeResult) -> None:
    if r.tier1_store_v1.get("ok") and r.tier1_v1_products_sample:
        r.recommended_tier = "TIER 1 — Store API v1 (idéal, JSON natif)"
        return
    if r.tier2_store_legacy.get("ok") and r.tier2_store_legacy.get("json_was_list"):
        r.recommended_tier = "TIER 1bis — Store API legacy (fonctionnel)"
        return
    if r.product_urls_count and r.product_urls_count > 0:
        r.recommended_tier = (
            f"TIER 2 — Sitemap + scrape produit page par page "
            f"(~{r.product_urls_count} produits détectés)"
        )
        return
    if r.sitemap_index.get("ok") or r.sitemap_root.get("ok"):
        r.recommended_tier = "TIER 2 — Sitemap accessible mais pas de product-sitemap clair"
        return
    r.recommended_tier = "TIER 3 — Aucune API ni sitemap : scraping HTML pur"


# --- Print rapport ----------------------------------------------------------

def print_report(r: ProbeResult) -> None:
    print()
    print("=" * 78)
    print(f" {r.brand_url}")
    print("=" * 78)

    # Niveau 1
    t1 = r.tier1_store_v1
    badge1 = "OK " if (t1.get("ok") and r.tier1_v1_products_sample) else "KO "
    print(f"\n [{badge1}] TIER 1     Store API v1   {t1.get('url', '')}")
    print(f"          HTTP={t1.get('status')}   bytes={t1.get('bytes')}   ct={t1.get('content_type')}")
    if t1.get("error"):
        print(f"          erreur: {t1['error']}")
    if r.tier1_v1_estimated_total is not None:
        print(f"          total produits estimé (X-WP-Total) : {r.tier1_v1_estimated_total}")
    for p in r.tier1_v1_products_sample:
        var_info = f"{p['n_variations']} variants" if p["has_options"] else "simple"
        name = (p.get("name") or "")[:60]
        print(f"          - #{p['id']} [{var_info}] {name!r}  prix={p['price']} {p['currency']}  imgs={p['n_images']}  stock={p['is_in_stock']}")

    # Niveau 1bis
    t2 = r.tier2_store_legacy
    badge2 = "OK " if (t2.get("ok") and t2.get("json_was_list")) else "KO "
    print(f"\n [{badge2}] TIER 1bis  Store API legacy   {t2.get('url', '')}")
    print(f"          HTTP={t2.get('status')}   bytes={t2.get('bytes')}")
    if t2.get("sample_first_keys"):
        print(f"          clés produit : {t2['sample_first_keys']}")

    # WP root
    wp = r.wp_json_root
    print(f"\n          WP-JSON discovery   status={wp.get('status')}   routes_produits_pertinentes={wp.get('n_product_related_routes')}")
    if wp.get("product_related_routes"):
        for route in wp["product_related_routes"][:6]:
            print(f"              - {route}")

    # Sitemap
    sm = r.sitemap_index
    badge3 = "OK " if sm.get("ok") else "KO "
    print(f"\n [{badge3}] Sitemap index   status={sm.get('status')}")
    if sm.get("sub_sitemaps_preview"):
        for s in sm["sub_sitemaps_preview"][:6]:
            print(f"              - {s}")
    if r.product_sitemap_url:
        print(f"          -> product sitemap detected : {r.product_sitemap_url}")
        print(f"          -> {r.product_urls_count} URLs produits ; échantillon :")
        for u in r.product_urls_sample:
            print(f"              - {u}")

    # robots
    rob = r.robots
    print(f"\n          robots.txt   status={rob.get('status')}")
    if rob.get("preview"):
        print("          preview :")
        for line in (rob["preview"] or "").splitlines()[:6]:
            print(f"              {line}")

    # Reco
    print(f"\n  >> RECOMMANDATION : {r.recommended_tier}")
    if r.notes:
        print("  notes :")
        for n in r.notes:
            print(f"   - {n}")
    print()


# --- Main -------------------------------------------------------------------

def probe(url: str) -> ProbeResult:
    base = normalize_base(url)
    r = ProbeResult(brand_url=url, base=base)
    probe_store_v1(base, r)
    probe_store_legacy(base, r)
    probe_wp_root(base, r)
    probe_sitemap_index(base, r)
    if not r.product_sitemap_url:
        probe_sitemap_root(base, r)
    probe_robots(base, r)
    compute_recommendation(r)
    return r


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage : python3 probe_woocommerce.py <url1> [<url2> ...]")
        print("Ex.    python3 probe_woocommerce.py https://cosmella.fr")
        return 2
    urls = sys.argv[1:]
    all_results: list[dict] = []
    for u in urls:
        try:
            r = probe(u)
        except Exception as e:
            print(f"[FATAL] {u} -> {type(e).__name__}: {e}")
            continue
        print_report(r)
        all_results.append(asdict(r))

    # Récap
    print("\n" + "=" * 78)
    print(" RÉCAP")
    print("=" * 78)
    for d in all_results:
        print(f" - {d['brand_url']}  ->  {d['recommended_tier']}")
    print()

    # Dump JSON
    out_path = "probe_report.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2, default=str)
    print(f"[Rapport JSON brut écrit : {out_path}]")
    return 0


if __name__ == "__main__":
    sys.exit(main())
