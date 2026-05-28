# Guide de déploiement — Ankorstore Catalog Scraper

Ce guide te fait passer du code local à l'app web partagée avec ton équipe en 4 étapes. Compte ~45 min au total la 1ère fois.

> **Vue d'ensemble** : 1) tester en local sans Sheets/Drive ; 2) créer un Service Account Google + un Sheet + un dossier Drive ; 3) tester en local avec Sheets/Drive ; 4) déployer sur Streamlit Cloud.

---

## Étape 1 — Test local sans Sheets/Drive (5 min)

C'est juste pour valider que ton Mac fait tourner l'app.

1. Ouvre le Finder dans `/Users/celia.daguet/Documents/Claude/Projects/Scrapping catalogues/streamlit_app/`
2. Double-clic sur **`run_local.command`**
   - Première fois : clic droit → Ouvrir, puis confirme le warning Gatekeeper macOS
3. Un terminal s'ouvre, créé un venv et installe les dépendances (~1 min la 1ère fois)
4. Ton navigateur s'ouvre automatiquement sur `http://localhost:8501`
5. Tu vois un warning jaune "Mode local sans authentification" — c'est normal, pas de password configuré
6. Tape ton prénom dans la barre latérale
7. Colle `https://cosmella.fr` dans le champ URL
8. Clique **🚀 Scraper le catalogue**
9. Attends ~30-60s, tu dois voir les métriques + bouton **⬇️ Télécharger**

✅ Étape validée si le xlsx se télécharge bien.

Pour arrêter l'app : `Ctrl+C` dans le terminal, puis ferme la fenêtre.

---

## Étape 2 — Créer Service Account Google + Sheet + Drive (15 min)

Cette étape branche l'historique d'équipe. Tu fais ça une seule fois.

### 2.1 — Créer un projet Google Cloud

1. Va sur [console.cloud.google.com](https://console.cloud.google.com)
2. En haut, clique sur le sélecteur de projet → **Nouveau projet**
3. Nom : `ankorstore-scraper`. Crée le projet.
4. Une fois créé, sélectionne-le dans le sélecteur

### 2.2 — Activer les APIs nécessaires

1. Dans la barre de recherche du haut, tape **"Google Sheets API"** → clique dessus → **Activer**
2. Pareil pour **"Google Drive API"** → **Activer**

### 2.3 — Créer un Service Account

1. Menu ☰ → **IAM & Admin** → **Service Accounts**
2. **Créer un compte de service**
3. Nom : `scraper-bot`. Description : `Bot scraping catalogues AE`. **Créer**
4. Étape "Permissions" : skip (pas besoin de rôle GCP). **Continuer**
5. Étape "Accès utilisateurs" : skip. **Terminer**

### 2.4 — Télécharger la clé JSON

1. Dans la liste, clique sur le service account `scraper-bot@...`
2. Onglet **Clés** → **Ajouter une clé** → **Créer une nouvelle clé** → **JSON** → **Créer**
3. Un fichier JSON se télécharge. ⚠️ **Garde-le bien**, on ne peut pas le re-télécharger.
4. Renomme-le `gcp_service_account.json` et stocke-le dans `streamlit_app/.streamlit/` (ne PAS commit ce fichier, il est déjà dans le .gitignore).

### 2.5 — Récupérer l'email du service account

Ouvre le JSON, copie la valeur du champ `client_email`. Ça ressemble à :
```
scraper-bot@ankorstore-scraper-1234.iam.gserviceaccount.com
```

### 2.6 — Créer le Google Sheet historique

1. Va sur [sheets.new](https://sheets.new) → crée un Sheet
2. Renomme-le `Ankorstore Scraping — Historique équipe`
3. Clique **Partager** (en haut à droite) → colle l'email du service account → donne-lui le rôle **Éditeur** → décoche "Notifier" → **Partager**
4. Dans l'URL du Sheet, copie l'ID (entre `/d/` et `/edit`) :
   - URL : `https://docs.google.com/spreadsheets/d/`**`1AbC...XyZ`**`/edit`
   - L'ID c'est `1AbC...XyZ`

### 2.7 — Créer le dossier Drive d'archive

1. Va sur [drive.google.com](https://drive.google.com) → **Nouveau** → **Dossier**
2. Nom : `Catalogues scrapés Ankorstore`
3. Clic droit sur le dossier → **Partager** → colle l'email du service account → **Éditeur** → **Partager**
4. Re-clic droit → **Copier le lien** : `https://drive.google.com/drive/folders/`**`1XyZ...AbC`**
5. L'ID du dossier c'est `1XyZ...AbC`

---

## Étape 3 — Configurer les secrets locaux et tester (5 min)

1. Dans `streamlit_app/.streamlit/`, copie `secrets.toml.example` en `secrets.toml`
2. Édite `secrets.toml` et remplace :

   ```toml
   app_password = "AnkorScrape2026!"  # change pour quelque chose à toi

   gsheet_history_id = "1AbC...XyZ"   # ID du Sheet (étape 2.6)
   gdrive_folder_id = "1XyZ...AbC"    # ID du dossier Drive (étape 2.7)
   ```

3. Pour la partie `[gcp_service_account]`, ouvre `gcp_service_account.json` (téléchargé étape 2.4) et copie-colle son contenu dans `secrets.toml` au bon format TOML.

   Astuce : la commande suivante fait la conversion automatique :

   ```bash
   cd streamlit_app/.streamlit/
   python3 -c "
   import json
   with open('gcp_service_account.json') as f: d = json.load(f)
   print('[gcp_service_account]')
   for k, v in d.items():
       v = str(v).replace('\\\\n', '\\n').replace('\"', '\\\\\"')
       print(f'{k} = \"{v}\"')
   " >> secrets.toml
   ```

4. Relance l'app : double-clic `run_local.command`
5. Cette fois tu dois voir :
   - Un écran de login (entre le password choisi)
   - Une fois loggé : `☁️ Connecté au Google Sheet d'équipe`
6. Scrape une URL test. Va voir ton Sheet et ton dossier Drive : une nouvelle ligne et un nouveau xlsx doivent apparaître.

✅ Étape validée si la ligne et le xlsx apparaissent dans Sheet/Drive.

---

## Étape 4 — Déployer sur Streamlit Cloud (15 min)

### 4.1 — Pré-requis

Tu as besoin d'un compte **GitHub** (gratuit) — pour héberger le code. Si tu n'en as pas : [github.com/signup](https://github.com/signup).

### 4.2 — Créer un repo GitHub privé

1. Va sur [github.com/new](https://github.com/new)
2. Nom : `ankorstore-scraper`
3. **Private** ✅ (important : code interne)
4. Crée le repo (sans README ni .gitignore, GitHub a tendance à ajouter du bruit)
5. GitHub te montre les commandes à exécuter

### 4.3 — Push le code

Dans un terminal sur ton Mac :

```bash
cd "/Users/celia.daguet/Documents/Claude/Projects/Scrapping catalogues"
git init
git add scrap_ankorstore.py ankorstore_template.xlsx streamlit_app/
git commit -m "Initial commit - scraper + Streamlit app"
git branch -M main
git remote add origin https://github.com/<TON_USER>/ankorstore-scraper.git
git push -u origin main
```

⚠️ **Vérifie** que `.streamlit/secrets.toml` et `gcp_service_account.json` ne sont PAS dans le commit (le `.gitignore` doit les filtrer). Pour vérifier :

```bash
git ls-files | grep -E "secrets.toml|gcp_service"
# (doit ne rien renvoyer)
```

### 4.4 — Déployer sur Streamlit Cloud

1. Va sur [streamlit.io/cloud](https://streamlit.io/cloud) → **Sign in with GitHub**
2. Autorise Streamlit à accéder à tes repos
3. **New app** → sélectionne ton repo `ankorstore-scraper`
4. **Main file path** : `streamlit_app/app.py`
5. **Advanced settings** → **Python version** : 3.11
6. Clique **Deploy!**

Premier deploy : ~3-5 min (install des dépendances).

### 4.5 — Configurer les secrets sur Streamlit Cloud

1. Une fois l'app déployée, va dans **Settings** → **Secrets**
2. Copie-colle le **contenu complet de ton `secrets.toml` local**
3. **Save**

L'app va se redémarrer. Une fois prête, tu auras une URL type :
`https://ankorstore-scraper-<random>.streamlit.app`

### 4.6 — Partager avec l'équipe

1. Renomme l'URL si tu veux (Settings → Change URL → `ankorstore-scraper.streamlit.app`)
2. Poste dans le Slack équipe AE :
   > Hey, voici un outil pour scraper les catalogues marques et générer le xlsx Ankorstore directement importable :
   > 🔗 https://ankorstore-scraper.streamlit.app
   > 🔑 Password : `[le password choisi]`
   > Le Sheet d'historique de qui a scrapé quoi : [lien Sheet]

---

## Mise à jour du scraper (futur)

Quand je (ou toi) on améliore le scraper :

```bash
cd "/Users/celia.daguet/Documents/Claude/Projects/Scrapping catalogues"
git add scrap_ankorstore.py streamlit_app/
git commit -m "Fix: amélioration parser Prestashop"
git push
```

→ Streamlit Cloud détecte le push et redéploie automatiquement (~2 min). Tes AE rechargent leur page et ont la nouvelle version.

---

## Coûts

- **GCP** : 0€ (les APIs Sheets/Drive sont gratuites au volume qu'on fait)
- **GitHub privé** : 0€
- **Streamlit Cloud Community** : 0€ (suffisant tant qu'on a <1Go RAM utilisé, ce qui est notre cas)

Total : **0€/mois**.

**Limite Streamlit Community** : l'app s'endort après 7 jours sans activité (réveil ~30s au prochain accès). Si gênant, upgrade Streamlit Cloud Pro à 20$/mois pour rester always-on. Pas urgent à mon avis.

---

## Troubleshooting

**L'app dit "ModuleNotFoundError: scrap_ankorstore"** → vérifie que `scrap_ankorstore.py` est bien à la racine du repo (pas dans `streamlit_app/`).

**Streamlit Cloud build fail** → regarde les logs Cloud, souvent un package manque dans `requirements.txt`.

**"Permission denied" sur Sheet/Drive** → le service account n'a pas été ajouté en Éditeur. Refais l'étape 2.6 / 2.7.

**Password partagé compromis** → change-le dans Streamlit Cloud Settings → Secrets → save. Tous les utilisateurs sont automatiquement déconnectés.

**Une marque ne se scrape pas bien** → augmente `Concurrence` dans la sidebar ou force un CMS spécifique. Si ça persiste, c'est un bug du scraper à corriger.
