#!/usr/bin/env bash
# ============================================================================
# Lance l'app Streamlit en local (Mac)
# ============================================================================
# Usage : double-clic sur ce fichier dans Finder
# (Première fois : clic droit → Ouvrir, pour passer le warning Gatekeeper)
# ============================================================================

set -e
cd "$(dirname "$0")"

ROOT_DIR="$(cd .. && pwd)"
VENV_DIR="$ROOT_DIR/venv"
PORT=8501

# Si un ancien serveur Streamlit traîne sur le port 8501, on le tue
EXISTING_PID=$(lsof -ti:${PORT} 2>/dev/null || true)
if [ -n "$EXISTING_PID" ]; then
    echo ">>> Arrêt de l'ancien serveur Streamlit (PID $EXISTING_PID)..."
    kill -9 $EXISTING_PID 2>/dev/null || true
    sleep 1
fi

# Crée le venv si inexistant
if [ ! -d "$VENV_DIR" ]; then
    echo ">>> Création du venv local (1ère fois, ~30s) ..."
    python3 -m venv "$VENV_DIR"
fi

# Active le venv
source "$VENV_DIR/bin/activate"

# Installe / met à jour les dépendances
echo ">>> Vérification des dépendances ..."
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

# Lance Streamlit
echo ""
echo "============================================================"
echo "  L'app va s'ouvrir dans ton navigateur sur localhost:${PORT}"
echo "  Pour arrêter : Ctrl+C dans ce terminal"
echo "============================================================"
echo ""
streamlit run app.py --server.port ${PORT}
