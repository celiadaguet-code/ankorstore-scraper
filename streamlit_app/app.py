"""
Ankorstore Catalog Scraper — Streamlit Web App
================================================

App web qui wrappe scrap_ankorstore.py pour usage interne Ankorstore.

Fonctionnalités v1 :
- Authentification par mot de passe partagé (st.secrets["app_password"])
- Identification AE par prénom (session)
- Scrape d'une URL marque → xlsx Ankorstore prêt à importer
- Téléchargement direct du xlsx
- (Phase 2) Historique dans Google Sheet partagé + xlsx archivés sur Drive

Pour lancer en local :
    streamlit run app.py

Pour déployer sur Streamlit Cloud :
    Voir DEPLOY.md
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import streamlit as st

# On ajoute le dossier parent au path pour importer scrap_ankorstore
APP_DIR = Path(__file__).parent
ROOT_DIR = APP_DIR.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

# NOTE : on n'importe PAS scrap_ankorstore ici (au top du fichier) pour éviter
# de bouffer ~200MB de RAM au démarrage (regex pré-compilées + 3000 lignes).
# Sur Streamlit Cloud Community (1 GB RAM), l'import au démarrage peut faire
# crasher l'app avant même que le serveur web ne démarre. On l'importe paresseusement
# dans _run_scrape() au moment du clic utilisateur.
#
# Import léger du module Sheets/Drive (juste quelques fonctions, pas de cache lourd).
try:
    from sheets_drive import (
        is_gcp_configured,
        upload_xlsx_to_drive,
        append_history_row,
        get_recent_scrapes,
        find_existing_scrape,
    )
    SHEETS_DRIVE_AVAILABLE = True
except ImportError:
    SHEETS_DRIVE_AVAILABLE = False
    is_gcp_configured = lambda: False
    upload_xlsx_to_drive = lambda *a, **kw: None
    append_history_row = lambda *a, **kw: False
    get_recent_scrapes = lambda *a, **kw: []
    find_existing_scrape = lambda *a, **kw: None


# ----------------------------------------------------------------------------
# Page config
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Ankorstore Catalog Scraper",
    page_icon="🌿",
    layout="centered",
    initial_sidebar_state="expanded",
)


# ----------------------------------------------------------------------------
# Identité visuelle Ankorstore
# ----------------------------------------------------------------------------
# Palette Tonal (surfaces) + Energetic (accents) + typo Poppins (Google Fonts).
# Note : GT Walsheim est une typo privée Ankorstore (Grilli Type), non dispo
# librement. Poppins est la typo secondaire officielle, dispo Google Fonts.
ANKORSTORE_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');

/* Override Poppins sur les éléments texte uniquement (pas les icônes Material) */
.stApp, .stApp p, .stApp label, .stApp button, .stApp input, .stApp textarea,
.stApp select, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
.stApp [data-testid="stMarkdownContainer"],
.stApp [data-testid="stMetricValue"],
.stApp [data-testid="stMetricLabel"] {
    font-family: 'Poppins', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* Préserve les fonts d'icônes Material (sinon les icônes type "keyboard_double_arrow_left"
   s'affichent en texte brut au survol) */
.material-symbols-outlined,
.material-icons,
.material-icons-outlined,
span[class*="material"],
[data-testid*="Icon"] *,
[class*="emotion-cache"] [class*="Icon"] *,
button[kind="header"] * {
    font-family: 'Material Symbols Outlined', 'Material Icons',
                 'Material Symbols Rounded' !important;
}

/* Titres en bold Poppins */
h1, h2, h3, h4, h5, h6 {
    font-family: 'Poppins', sans-serif !important;
    font-weight: 700 !important;
    color: #14060A !important;
    letter-spacing: -0.02em;
}

/* Bouton primaire — Moss (Tonal palette Ankorstore) */
.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"],
.stFormSubmitButton > button[kind="primary"] {
    background-color: #567570 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 0.6rem 1.2rem !important;
    transition: background-color 0.15s ease;
}
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover,
.stFormSubmitButton > button[kind="primary"]:hover {
    background-color: #455E5A !important;
    color: #FFFFFF !important;
}

/* Boutons secondaires (warm tone) */
.stButton > button[kind="secondary"] {
    background-color: #FBF8F3 !important;
    color: #14060A !important;
    border: 1.5px solid #14060A !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
}
.stButton > button[kind="secondary"]:hover {
    background-color: #F3E0C6 !important;
    border-color: #14060A !important;
}

/* Inputs : bordure douce, focus rouge Ankorstore */
.stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stSelectbox > div > div {
    border-radius: 8px !important;
    border: 1.5px solid #EEEEEE !important;
}
.stTextInput > div > div > input:focus {
    border-color: #567570 !important;
    box-shadow: 0 0 0 1px #567570 !important;
}

/* Sidebar : fond Sand + arrondi */
[data-testid="stSidebar"] {
    background-color: #F3E0C6 !important;
}

/* Metrics : style card */
[data-testid="stMetric"] {
    background-color: #FFFFFF;
    padding: 1rem;
    border-radius: 12px;
    border: 1px solid #EEEEEE;
}
[data-testid="stMetricValue"] {
    font-family: 'Poppins', sans-serif !important;
    font-weight: 700 !important;
    color: #14060A !important;
    font-size: 1.6rem !important;
}
[data-testid="stMetricLabel"] {
    color: #777272 !important;
    font-size: 0.85rem !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

/* Alerts (info, success, warning, error) : arrondi + bordure */
.stAlert {
    border-radius: 12px !important;
    border-left-width: 4px !important;
}

/* Expanders : warm tone */
.streamlit-expanderHeader {
    background-color: #FBF8F3 !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
}

/* Titre principal : un peu plus serré */
h1 {
    margin-bottom: 0.4rem !important;
}

/* Divider plus discret */
hr {
    border-color: #EEEEEE !important;
}
</style>
"""
st.markdown(ANKORSTORE_CSS, unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# Authentication
# ----------------------------------------------------------------------------
def _get_app_password() -> str:
    """Renvoie le password attendu (st.secrets), ou '' si non configuré."""
    try:
        return st.secrets.get("app_password", "")
    except FileNotFoundError:
        # Pas de fichier secrets.toml → mode local sans auth
        return ""


def check_password() -> bool:
    """Affiche un écran de login si pas encore authentifié dans la session."""
    if st.session_state.get("authenticated"):
        return True

    expected = _get_app_password()

    # Cas mode dev local (pas de secrets configurés) → bypass avec warning
    if not expected:
        st.warning(
            "⚠️ **Mode local sans authentification.** "
            "Aucun `app_password` n'est configuré dans `.streamlit/secrets.toml`. "
            "C'est OK pour tester en local, mais ne pas déployer en prod sans password."
        )
        st.session_state.authenticated = True
        return True

    # Cas normal : écran de login
    st.title("🔒 Accès Ankorstore")
    st.caption("Outil interne Ankorstore — entrer le mot de passe partagé.")

    with st.form("login_form", clear_on_submit=False):
        pwd = st.text_input("Mot de passe", type="password")
        submit = st.form_submit_button("Se connecter")

    if submit:
        if pwd == expected:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Mot de passe incorrect.")

    return False


if not check_password():
    st.stop()


# ----------------------------------------------------------------------------
# Sidebar : prénom AE + paramètres avancés
# ----------------------------------------------------------------------------
with st.sidebar:
    st.subheader("👤 Identification")
    ae_name = st.text_input(
        "Ton prénom",
        value=st.session_state.get("ae_name", ""),
        placeholder="Célia",
        help="Mémorisé pour la session, et tracé dans l'historique d'équipe.",
    )
    if ae_name:
        st.session_state.ae_name = ae_name.strip()

    st.divider()
    st.subheader("⚙️ Paramètres avancés")
    cms_choice = st.selectbox(
        "CMS",
        ["auto", "woocommerce", "prestashop", "wix", "squarespace", "sumup", "custom"],
        index=0,
        help="Auto-détection par défaut. Force un CMS si la détection se trompe.",
    )
    max_products = st.number_input(
        "Limite produits (debug)",
        min_value=0,
        value=0,
        help="0 = pas de limite. Mets une petite valeur (ex: 5) pour tester rapidement.",
    )
    speed_choice = st.select_slider(
        "Vitesse",
        options=["🐢 Lent", "🚶 Normal", "🐇 Rapide"],
        value="🚶 Normal",
        help=(
            "Nombre de pages produit récupérées en parallèle.\n\n"
            "• **🐢 Lent (1)** : pour les serveurs qui plantent sous charge "
            "(ex: Promolinge, petits Prestashop).\n"
            "• **🚶 Normal (2)** : recommandé par défaut, équilibre rapidité/stabilité.\n"
            "• **🐇 Rapide (4)** : pour les gros serveurs (Wix, Squarespace), "
            "scrape ~2× plus vite."
        ),
    )
    concurrency = {"🐢 Lent": 1, "🚶 Normal": 2, "🐇 Rapide": 4}[speed_choice]

    st.divider()
    st.caption(f"Connecté en tant que : **{ae_name or 'anonyme'}**")
    if st.button("Se déconnecter", use_container_width=True):
        st.session_state.clear()
        st.rerun()


# ----------------------------------------------------------------------------
# Main UI
# ----------------------------------------------------------------------------
st.title("📦 Scrapeur catalogue → Ankorstore")
st.caption(
    "Colle l'URL d'un site marque, l'outil scrape le catalogue et génère "
    "un .xlsx prêt à importer dans Ankorstore. "
    "CMS supportés : WooCommerce, PrestaShop, Wix, Squarespace, SumUp Store, sites custom."
)

if not ae_name:
    st.warning("⚠️ Renseigne ton prénom dans la barre latérale avant de scraper.")
    st.stop()

# Statut Sheets/Drive (badge discret)
if is_gcp_configured():
    st.caption("☁️ Connecté au Google Sheet d'équipe — chaque scrape est tracé.")
else:
    st.caption(
        "🔌 *Pas de Google Sheet configuré — les scrapes ne sont pas tracés "
        "dans l'historique d'équipe (mode local).*"
    )

with st.form("scrape_form"):
    brand_url = st.text_input(
        "URL de la marque",
        placeholder="https://cosmella.fr",
        help="L'URL racine du site (sans /shop ou /produits).",
    )
    submitted = st.form_submit_button(
        "🚀 Scraper le catalogue", type="primary", use_container_width=True
    )

# Panneau "Historique récent d'équipe" (sous le formulaire)
if is_gcp_configured() and not submitted:
    with st.expander("📜 Historique récent (équipe)", expanded=False):
        recent = get_recent_scrapes(limit=15)
        if not recent:
            st.caption("Aucun scrape encore dans l'historique.")
        else:
            for row in recent:
                domain = row.get("Domaine", "?")
                ae = row.get("AE", "?")
                date = row.get("Date", "?")
                status = row.get("Statut", "?")
                link = row.get("Lien xlsx", "")
                n_prod = row.get("Nb produits", "?")
                emoji = {"success": "✅", "partial": "⚠️", "failed": "❌"}.get(status, "•")
                if link:
                    st.markdown(
                        f"{emoji} **{domain}** — {n_prod} produits — "
                        f"{ae}, {date} — [📥 xlsx]({link})"
                    )
                else:
                    st.markdown(
                        f"{emoji} **{domain}** — {n_prod} produits — {ae}, {date}"
                    )

# ----------------------------------------------------------------------------
# Lance un scrape si le formulaire vient d'être soumis
# ----------------------------------------------------------------------------
def _run_scrape(url: str) -> None:
    """Exécute le scrape et stocke le résultat en session_state."""
    # Import paresseux : ne charge scrap_ankorstore qu'au 1er scrape (économie RAM)
    try:
        from scrap_ankorstore import process_brand, BrandReport
    except ImportError as e:
        st.session_state.last_error = (
            f"Impossible d'importer scrap_ankorstore.py : {e}"
        )
        st.session_state.last_result = None
        return

    output_dir = ROOT_DIR / "outputs"
    output_dir.mkdir(exist_ok=True)

    status_box = st.empty()
    status_box.info(
        f"🕐 Scraping de **{url}** en cours... "
        "(peut prendre 30s à 3min selon la marque)"
    )

    try:
        with st.spinner("Détection CMS et récupération du catalogue..."):
            report = process_brand(
                url,
                output_dir,
                max_products=int(max_products) if max_products > 0 else None,
                cms=cms_choice,
                concurrency=int(concurrency),
            )
    except Exception as e:
        status_box.empty()
        st.session_state.last_error = f"{type(e).__name__}: {e}"
        st.session_state.last_result = None
        return

    status_box.empty()

    # Charge le xlsx en mémoire pour persistance (le fichier disque peut bouger)
    xlsx_bytes = None
    filename = None
    if report.output_file and Path(report.output_file).exists():
        with open(report.output_file, "rb") as f:
            xlsx_bytes = f.read()
        filename = Path(report.output_file).name

    # Lit le .log pour extraire les warnings (affichage UI lisible)
    warnings_list: list[str] = []
    if report.output_file:
        log_file = Path(report.output_file).parent / f"{report.domain}.log"
        if log_file.exists():
            try:
                with open(log_file, encoding="utf-8") as f:
                    for line in f:
                        if " WARNING " in line:
                            # Strip timestamp + level pour ne garder que le message utile
                            parts = line.split(" WARNING ", 1)
                            msg = parts[1].strip() if len(parts) > 1 else line.strip()
                            warnings_list.append(msg)
            except Exception:
                pass

    # Upload Drive + Sheet (UNE seule fois, ici dans le scrape)
    drive_link = ""
    sheet_ok = False
    if is_gcp_configured() and report.output_file and report.status != "failed":
        with st.spinner("Sauvegarde dans l'historique d'équipe..."):
            upload_result = upload_xlsx_to_drive(Path(report.output_file), report.domain)
            if upload_result:
                _, drive_link = upload_result
            sheet_ok = append_history_row(
                ae_name=ae_name,
                brand_url=url,
                brand_domain=report.domain,
                cms=cms_choice,
                status=report.status,
                # Nombre de produits réellement gardés (= total − exclus)
                n_products=report.n_products_total - report.n_products_filtered,
                n_variants=report.n_variants_total,
                n_warnings=report.n_warnings,
                duration_s=report.duration_s,
                xlsx_link=drive_link,
                notes="",
            )

    # Stocke tout en session_state pour persister entre les reruns
    st.session_state.last_result = {
        "status": report.status,
        "domain": report.domain,
        "url": url,
        "n_products_total": report.n_products_total,
        "n_products_filtered": report.n_products_filtered,
        "n_variants_total": report.n_variants_total,
        "n_warnings": report.n_warnings,
        "duration_s": report.duration_s,
        "error": report.error,
        "xlsx_bytes": xlsx_bytes,
        "filename": filename,
        "drive_link": drive_link,
        "sheet_ok": sheet_ok,
        "filtered_out_items": getattr(report, "filtered_out_items", []),
        "warnings_list": warnings_list,
    }
    st.session_state.last_error = None


# Si le formulaire vient d'être soumis avec une URL valide, lance le scrape
if submitted and brand_url.strip():
    url = brand_url.strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url

    # Anti-duplication : check si déjà scrapée par l'équipe
    existing = find_existing_scrape(url)
    if existing and not st.session_state.get("force_rescrape"):
        when = existing.get("Date", "?")
        who = existing.get("AE", "?")
        link = existing.get("Lien xlsx", "")
        st.info(
            f"💡 **{existing.get('Domaine', '?')}** a déjà été scrapée le **{when}** par **{who}**. "
            + (f"[Télécharger le xlsx existant]({link})" if link else "")
        )
        if st.button("Scraper quand même (rafraîchir le catalogue)", type="secondary"):
            st.session_state.force_rescrape = True
            st.rerun()
        st.stop()
    else:
        # Reset le flag, et lance le scrape
        st.session_state.force_rescrape = False
        _run_scrape(url)
elif submitted and not brand_url.strip():
    st.error("URL vide — colle une URL.")


# ----------------------------------------------------------------------------
# Affiche le dernier résultat (persiste entre les reruns grâce à session_state)
# ----------------------------------------------------------------------------
if st.session_state.get("last_error"):
    st.error(f"❌ Crash inattendu : `{st.session_state.last_error}`")
    st.info(
        "Cause probable : URL mal formée, site bloque le scraping, ou bug du parser. "
        "Réessaie ou essaie de forcer un CMS dans la sidebar."
    )

result = st.session_state.get("last_result")
if result:
    st.divider()
    st.subheader(f"📊 Dernier scrape : {result['domain']}")

    if result["status"] == "failed":
        st.error("❌ Échec du scraping")
        st.code(result["error"] or "Erreur inconnue")
        st.info(
            "Causes fréquentes : URL invalide, site bloque le scraping, "
            "ou CMS non détecté. Essaie de forcer un CMS dans la sidebar."
        )
    else:
        if result["status"] == "partial":
            st.warning(f"⚠️ Scraping partiel (warnings : {result['n_warnings']})")
        else:
            st.success("✅ Scraping réussi !")

        n_total = result["n_products_total"]
        n_excluded = result["n_products_filtered"]  # = filtered OUT (nom trompeur côté scraper)
        n_kept = n_total - n_excluded

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Produits trouvés", n_total)
        col2.metric(
            "Produits gardés",
            n_kept,
            delta=f"-{n_excluded} exclus" if n_excluded > 0 else None,
            delta_color="off",
            help="Produits gardés = trouvés − exclus (ateliers, cartes cadeaux, consultations, etc.)",
        )
        col3.metric(
            "Lignes (variants)",
            result["n_variants_total"],
            help="1 ligne = 1 variante (teinte, contenance, taille…). Format Ankorstore.",
        )
        col4.metric("Durée", f"{result['duration_s']}s")

        if result["error"] and result["status"] != "success":
            with st.expander("Détails warnings/erreurs"):
                st.code(result["error"])

        # Liste détaillée des warnings (descriptions trop courtes, prix manquants, etc.)
        warnings_list = result.get("warnings_list") or []
        if warnings_list:
            with st.expander(
                f"⚠️ Détails — {len(warnings_list)} warning(s) à vérifier dans le xlsx",
                expanded=False,
            ):
                st.caption(
                    "Ces produits ont été extraits mais ont un champ incomplet "
                    "(description trop courte, prix manquant, etc.). "
                    "Les cellules concernées sont **surlignées en jaune** dans le xlsx — "
                    "tu peux les compléter manuellement avant l'import Ankorstore."
                )
                for w in warnings_list:
                    st.markdown(f"• {w}")

        # Liste des produits exclus par les filtres (ateliers, cartes cadeaux…)
        excluded_items = result.get("filtered_out_items") or []
        if excluded_items:
            with st.expander(
                f"📋 Détails — {len(excluded_items)} produit(s) exclu(s) par les filtres"
            ):
                st.caption(
                    "Ces produits ont été exclus automatiquement car ils ne sont pas "
                    "vendables sur Ankorstore (ateliers, cartes cadeaux, consultations, "
                    "alcool sans flag, etc.). Si un produit est exclu à tort, dis-le moi "
                    "pour ajuster les filtres."
                )
                # Groupe par raison pour lisibilité
                by_reason: dict[str, list[str]] = {}
                for name, reason in excluded_items:
                    by_reason.setdefault(reason, []).append(name)
                for reason, names in by_reason.items():
                    st.markdown(f"**{reason}** ({len(names)})")
                    for n in names:
                        st.markdown(f"&nbsp;&nbsp;&nbsp;• {n}", unsafe_allow_html=True)

        # Bouton télécharger (persistant grâce aux bytes en session_state)
        if result["xlsx_bytes"]:
            st.download_button(
                label=f"⬇️ Télécharger {result['filename']}",
                data=result["xlsx_bytes"],
                file_name=result["filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key=f"download_{result['domain']}",
            )
            st.caption(
                "Fichier prêt à importer dans Ankorstore — "
                "1 ligne = 1 variante. Les champs obligatoires vides sont surlignés en jaune."
            )

        # Récap Sheets/Drive
        if result["drive_link"]:
            st.caption(f"☁️ Archivé sur Drive : [voir le fichier]({result['drive_link']})")
        if result["sheet_ok"]:
            st.caption("✏️ Ligne ajoutée dans le Sheet historique d'équipe.")

    # Bouton "Nouveau scrape" pour clear l'état
    if st.button("🔄 Nouveau scrape (vider le résultat)", use_container_width=True):
        st.session_state.last_result = None
        st.session_state.last_error = None
        st.session_state.force_rescrape = False
        st.rerun()
