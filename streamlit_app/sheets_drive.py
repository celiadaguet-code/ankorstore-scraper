"""
sheets_drive.py
================

Branchement Google Sheets (historique d'équipe) + Google Drive (archive xlsx)
pour l'app Streamlit.

Activation : il suffit de configurer les 3 secrets dans .streamlit/secrets.toml
ou dans le dashboard Streamlit Cloud :
  - [gcp_service_account] (le JSON du service account, aplati en TOML)
  - gsheet_history_id (l'ID du Sheet)
  - gdrive_folder_id (l'ID du dossier Drive)

Si ces secrets ne sont PAS configurés, l'app fonctionne en mode dégradé
(téléchargement xlsx local uniquement, pas d'historique partagé).
"""
from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path
from typing import Any

import streamlit as st


# Colonnes du Sheet historique (ordre exact, à créer dans la 1ère ligne du Sheet)
HISTORY_HEADERS = [
    "Date",
    "AE",
    "URL marque",
    "Domaine",
    "CMS",
    "Statut",
    "Nb produits",
    "Nb variants",
    "Warnings",
    "Durée (s)",
    "Lien xlsx",
    "Notes",
]


def is_gcp_configured() -> bool:
    """Vérifie si tous les secrets GCP sont configurés."""
    try:
        return bool(
            st.secrets.get("gcp_service_account")
            and st.secrets.get("gsheet_history_id")
            and st.secrets.get("gdrive_folder_id")
        )
    except (FileNotFoundError, KeyError):
        return False


@st.cache_resource(show_spinner=False)
def _get_credentials():
    """Crée les credentials Google une seule fois par session."""
    from google.oauth2.service_account import Credentials

    info = dict(st.secrets["gcp_service_account"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
    ]
    return Credentials.from_service_account_info(info, scopes=scopes)


@st.cache_resource(show_spinner=False)
def _get_drive_service():
    from googleapiclient.discovery import build
    return build("drive", "v3", credentials=_get_credentials())


@st.cache_resource(show_spinner=False)
def _get_gspread_client():
    import gspread
    return gspread.authorize(_get_credentials())


# ----------------------------------------------------------------------------
# Drive : upload xlsx
# ----------------------------------------------------------------------------
def upload_xlsx_to_drive(local_path: Path, brand_domain: str) -> tuple[str, str] | None:
    """
    Upload un xlsx dans le dossier Drive configuré.
    Retourne (file_id, web_view_link) ou None en cas d'erreur.
    """
    if not is_gcp_configured():
        return None

    try:
        from googleapiclient.http import MediaFileUpload

        folder_id = st.secrets["gdrive_folder_id"]
        service = _get_drive_service()

        # Préfixe la date pour éviter les écrasements
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        remote_name = f"{timestamp}__{brand_domain}.xlsx"

        file_metadata = {
            "name": remote_name,
            "parents": [folder_id],
        }
        media = MediaFileUpload(
            str(local_path),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=False,
        )
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,
        ).execute()

        file_id = file.get("id")
        view_link = file.get("webViewLink", "")

        # Rend le fichier lisible par toute personne avec le lien (équipe AE)
        try:
            service.permissions().create(
                fileId=file_id,
                body={"type": "anyone", "role": "reader"},
                supportsAllDrives=True,
            ).execute()
        except Exception:
            pass  # Non bloquant si la permission échoue

        return file_id, view_link

    except Exception as e:
        st.warning(f"⚠️ Upload Drive échoué : {e}")
        return None


# ----------------------------------------------------------------------------
# Sheets : ensure headers + append history row
# ----------------------------------------------------------------------------
def _ensure_history_headers(worksheet) -> None:
    """Crée la ligne d'entête si le Sheet est vide."""
    try:
        first_row = worksheet.row_values(1)
    except Exception:
        first_row = []
    if not first_row:
        worksheet.update("A1", [HISTORY_HEADERS], value_input_option="USER_ENTERED")
        try:
            worksheet.format("A1:L1", {"textFormat": {"bold": True}})
        except Exception:
            pass


def append_history_row(
    *,
    ae_name: str,
    brand_url: str,
    brand_domain: str,
    cms: str,
    status: str,
    n_products: int,
    n_variants: int,
    n_warnings: int,
    duration_s: int,
    xlsx_link: str,
    notes: str = "",
) -> bool:
    """
    Ajoute une ligne dans le Sheet historique.
    Retourne True si succès, False sinon.
    """
    if not is_gcp_configured():
        return False

    try:
        sheet_id = st.secrets["gsheet_history_id"]
        gc = _get_gspread_client()
        sh = gc.open_by_key(sheet_id)
        ws = sh.sheet1

        _ensure_history_headers(ws)

        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ae_name,
            brand_url,
            brand_domain,
            cms,
            status,
            n_products,
            n_variants,
            n_warnings,
            duration_s,
            xlsx_link,
            notes,
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        return True

    except Exception as e:
        st.warning(f"⚠️ Écriture Sheet échouée : {e}")
        return False


# ----------------------------------------------------------------------------
# Sheets : lecture historique (pour anti-duplication / récap)
# ----------------------------------------------------------------------------
def get_recent_scrapes(limit: int = 20) -> list[dict[str, Any]]:
    """Lit les N derniers scrapes pour afficher dans l'app."""
    if not is_gcp_configured():
        return []

    try:
        sheet_id = st.secrets["gsheet_history_id"]
        gc = _get_gspread_client()
        sh = gc.open_by_key(sheet_id)
        ws = sh.sheet1
        all_records = ws.get_all_records()
        # Les plus récents d'abord (les nouvelles lignes sont append en bas)
        return list(reversed(all_records))[:limit]
    except Exception:
        return []


def find_existing_scrape(brand_url: str) -> dict[str, Any] | None:
    """Cherche si une URL marque a déjà été scrapée récemment."""
    if not is_gcp_configured():
        return None

    try:
        recent = get_recent_scrapes(limit=200)
        # Normalise l'URL pour la comparaison (sans trailing slash)
        target = brand_url.rstrip("/").lower()
        for row in recent:
            stored = str(row.get("URL marque", "")).rstrip("/").lower()
            if stored == target and row.get("Statut") == "success":
                return row
        return None
    except Exception:
        return None
