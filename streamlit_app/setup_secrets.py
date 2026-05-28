#!/usr/bin/env python3
"""
setup_secrets.py — Wizard interactif pour créer .streamlit/secrets.toml
========================================================================

Lance ce script une seule fois après avoir :
  1. Téléchargé le JSON du service account GCP (gcp_service_account.json)
  2. Récupéré l'ID du Sheet historique et l'ID du dossier Drive

Le wizard te demande :
  - Le password d'app (que les AE utiliseront pour se connecter)
  - L'ID du Sheet (avec valeur par défaut, ré-appuie sur Entrée pour valider)
  - L'ID du dossier Drive (idem)
  - Le chemin du JSON GCP (auto-détecté dans le dossier)

Sortie : .streamlit/secrets.toml prêt pour Streamlit (local + cloud).

Usage :
    cd streamlit_app/
    python3 setup_secrets.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path


# Valeurs par défaut (modifiées par Célia le 28 mai 2026, ré-appuie sur Entrée
# pour les conserver, ou tape une nouvelle valeur).
DEFAULT_SHEET_ID = "133RFUlvGqY34yE30j82QbKSYolKC0xuE22P4hjJM-ZY"
DEFAULT_DRIVE_ID = "1Hg7CyCDRTuon_xc39L36_ZJnuXjgK_eO"


def prompt(question: str, default: str = "") -> str:
    """Demande une valeur à l'utilisateur, avec un default optionnel."""
    suffix = f" [{default}]" if default else ""
    answer = input(f"{question}{suffix}: ").strip()
    return answer or default


def find_gcp_json() -> Path | None:
    """Cherche un fichier JSON de service account dans le dossier courant."""
    here = Path(__file__).parent
    candidates = sorted(here.glob("*.json"))
    # Privilégie ceux dont le nom évoque un service account
    for c in candidates:
        if "service" in c.name.lower() or "credential" in c.name.lower() or "gcp" in c.name.lower():
            return c
    return candidates[0] if candidates else None


def escape_toml_value(v) -> str:
    """Échappe une valeur pour la mettre dans un string TOML triple-quoté."""
    s = str(v)
    # Échappe les backslash et les triple quotes
    s = s.replace("\\", "\\\\")
    s = s.replace('"""', '\\"\\"\\"')
    return s


def main() -> int:
    print("=" * 70)
    print("  Setup wizard — Ankorstore Catalog Scraper secrets")
    print("=" * 70)
    print()

    here = Path(__file__).parent
    secrets_dir = here / ".streamlit"
    secrets_dir.mkdir(exist_ok=True)
    secrets_file = secrets_dir / "secrets.toml"

    # Avertit si secrets.toml existe déjà
    if secrets_file.exists():
        print(f"⚠️  Un fichier existe déjà : {secrets_file}")
        confirm = input("    L'écraser ? (oui/non) [non]: ").strip().lower()
        if confirm not in ("oui", "o", "yes", "y"):
            print("Abandonné. Aucune modification.")
            return 0

    # 1. Password
    print("\n[1/4] Password d'app (les AE le tapent pour se connecter)")
    print("      Conseil : ~16 caractères, mélange lettres/chiffres/symboles.")
    password = prompt("      App password", "AnkorScrape2026!")

    # 2. Sheet ID
    print("\n[2/4] ID du Google Sheet historique")
    sheet_id = prompt("      Sheet ID", DEFAULT_SHEET_ID)

    # 3. Drive Folder ID
    print("\n[3/4] ID du dossier Drive d'archive")
    drive_id = prompt("      Drive Folder ID", DEFAULT_DRIVE_ID)

    # 4. GCP JSON
    print("\n[4/4] Fichier JSON du service account GCP")
    detected = find_gcp_json()
    if detected:
        print(f"      Auto-détecté : {detected.name}")
    json_path_str = prompt(
        "      Chemin du JSON (relatif ou absolu)",
        str(detected) if detected else "",
    )
    json_path = Path(json_path_str).expanduser().resolve()
    if not json_path.exists():
        print(f"\n❌ Fichier introuvable : {json_path}")
        return 1

    try:
        with open(json_path, encoding="utf-8") as f:
            gcp_data = json.load(f)
    except Exception as e:
        print(f"\n❌ Lecture JSON impossible : {e}")
        return 1

    if not isinstance(gcp_data, dict) or "client_email" not in gcp_data:
        print(f"\n❌ Le JSON n'a pas l'air d'un service account GCP (pas de client_email).")
        return 1

    print(f"      ✅ Service account : {gcp_data['client_email']}")

    # Construit le secrets.toml
    lines: list[str] = []
    lines.append("# =============================================================================")
    lines.append("# Ankorstore Catalog Scraper — Streamlit secrets")
    lines.append("# Généré par setup_secrets.py. NE PAS COMMITTER (filtré par .gitignore).")
    lines.append("# =============================================================================\n")

    lines.append("# Password partagé avec l'équipe AE")
    lines.append(f'app_password = "{password}"\n')

    lines.append("# IDs Google Sheet (historique) + Drive (archive xlsx)")
    lines.append(f'gsheet_history_id = "{sheet_id}"')
    lines.append(f'gdrive_folder_id = "{drive_id}"\n')

    lines.append("# Credentials Service Account Google (extraits du JSON)")
    lines.append("[gcp_service_account]")
    for key, value in gcp_data.items():
        if isinstance(value, str):
            # Strings : utilise triple-quote pour gérer les newlines de private_key
            escaped = escape_toml_value(value)
            if "\n" in value:
                lines.append(f'{key} = """{escaped}"""')
            else:
                lines.append(f'{key} = "{escaped}"')
        elif isinstance(value, bool):
            lines.append(f'{key} = {str(value).lower()}')
        elif isinstance(value, (int, float)):
            lines.append(f'{key} = {value}')
        elif isinstance(value, list):
            items = ", ".join(f'"{escape_toml_value(x)}"' for x in value)
            lines.append(f'{key} = [{items}]')

    content = "\n".join(lines) + "\n"

    with open(secrets_file, "w", encoding="utf-8") as f:
        f.write(content)

    print()
    print("=" * 70)
    print(f"  ✅ Fichier créé : {secrets_file}")
    print("=" * 70)
    print()
    print("Prochaines étapes :")
    print("  1. Lance l'app : double-clic sur run_local.command")
    print("  2. Tu verras un écran de login (entre le password choisi)")
    print("  3. Une fois loggée, scrape une URL test — la ligne doit apparaître")
    print(f"     dans le Sheet et le xlsx dans le dossier Drive.")
    print()
    print("⚠️  Tu peux maintenant SUPPRIMER le fichier JSON original :")
    print(f"     rm {json_path}")
    print("     (son contenu est maintenant dans secrets.toml, on n'en a plus besoin)")
    print()

    return 0


if __name__ == "__main__":
    sys.exit(main())
