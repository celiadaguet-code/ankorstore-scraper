# Ankorstore Catalog Scraper

> **Outil interne Ankorstore** — Scrape automatiquement le catalogue produits d'un site marque et génère un `.xlsx` prêt à importer dans Ankorstore.

🔗 **App en prod** : https://ankorstore-scraper.streamlit.app
📊 **Historique d'équipe** : [Google Sheet](https://docs.google.com/spreadsheets/d/133RFUlvGqY34yE30j82QbKSYolKC0xuE22P4hjJM-ZY)

---

## 🎯 Le problème résolu

Quand un AE rencontre une marque prospect, importer son catalogue dans Ankorstore prend habituellement **2 à 3 heures à la main** (copier nom, prix, description, images, variants, etc. pour 50-500 produits).

Cet outil fait le même travail en **30 secondes à 3 minutes** par marque, et génère un xlsx **directement importable** dans Ankorstore (1 ligne = 1 variante, champs obligatoires surlignés en jaune si manquants).

---

## 🚀 Utilisation (pour les AE)

1. Aller sur **https://ankorstore-scraper.streamlit.app**
2. Entrer le password partagé : `AnkorScrape2026!`
3. Renseigner son prénom (mémorisé pour la session)
4. Coller l'URL d'une marque prospect (ex : `https://cosmella.fr`)
5. Cliquer **🚀 Scraper le catalogue**
6. Attendre 30 secondes à 3 minutes selon la taille du catalogue
7. **Télécharger le xlsx** → l'importer dans Ankorstore

L'app vérifie automatiquement si quelqu'un de l'équipe a déjà scrapé cette marque récemment (anti-duplication). Chaque scrape laisse une trace dans le Google Sheet partagé.

---

## 📊 CMS et plateformes supportés

| Plateforme | Statut | Méthode |
|---|---|---|
| **WooCommerce** | ✅ | Store API v1 + fallback sitemap HTML |
| **PrestaShop** | ✅ | Attribut `data-product` JSON + fallback HTML legacy |
| **Wix** | ✅ | `productItems` JSON + JSON-LD + upscale images |
| **Squarespace** | ✅ | Trick `?format=json` + extraction variants |
| **SumUp Store** | ✅ | React Server Components flight data |
| **Sites custom** | ⚠️ | Cascade JSON-LD → Open Graph → microdata (best-effort) |
| **Shopify** | ➖ Skippé | Intégration native Ankorstore déjà existante |
| **Etsy / Facebook Shop** | ❌ | Non viable (anti-scraping fort) |

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────┐
│  Streamlit Cloud Community (hébergement)    │
│  - Auth password partagé                    │
│  - UI Ankorstore (Moss + Poppins)           │
│  - Session state (résultats persistés)      │
└─────────────────┬───────────────────────────┘
                  │  (import paresseux pour limiter la RAM)
                  ▼
┌─────────────────────────────────────────────┐
│  scrap_ankorstore.py  (~3000 lignes)        │
│  - Détection CMS automatique                │
│  - 6 scrapers spécialisés                   │
│  - Filtres anti-bruit (alcool, ateliers,    │
│    cartes cadeaux, consultations…)          │
│  - Mapping vers template xlsx Ankorstore    │
│  - Highlight jaune des champs obligatoires  │
└──┬──────────────────────┬───────────────────┘
   │                      │
   ▼                      ▼
┌──────────────┐    ┌──────────────────────┐
│ Sites marques│    │ Google Workspace     │
│ (scraping    │    │ - Sheet historique   │
│  HTTP/HTML)  │    │ - Drive Partagé      │
└──────────────┘    │   (archive xlsx)     │
                    └──────────────────────┘
```

---

## 🛠️ Stack technique

- **Python 3.11** (compatibilité Mac / Linux / Streamlit Cloud)
- **Streamlit 1.58** (interface web — pas de framework JS à maintenir)
- **openpyxl** (manipulation du template xlsx)
- **gspread + google-api-python-client** (historique partagé)
- **urllib + ThreadPoolExecutor** (scraping HTTP, sans dépendance type `requests`)
- **Hébergement** : Streamlit Cloud Community (gratuit, plan 1 GB RAM)
- **Repo** : GitHub privé (le code redevient public temporairement à chaque update pour permettre à Streamlit Cloud de re-pull)

---

## 📁 Structure du projet

```
ankorstore-scraper/
├── README.md                    # Ce fichier
├── scrap_ankorstore.py          # Le scraper (~3000 lignes, 6 CMS)
├── ankorstore_template.xlsx     # Template Ankorstore officiel
├── probe_*.py                   # Scripts de diagnostic par CMS
├── diag_*.py                    # Diagnostics ponctuels
├── streamlit_app/
│   ├── app.py                   # UI Streamlit
│   ├── sheets_drive.py          # Module Google Sheets + Drive
│   ├── setup_secrets.py         # Wizard de configuration initiale
│   ├── requirements.txt         # Dépendances Python
│   ├── DEPLOY.md                # Guide de déploiement complet
│   ├── run_local.command        # Lancement local (Mac)
│   ├── run_local.bat            # Lancement local (Windows)
│   └── .streamlit/
│       ├── config.toml          # Thème Ankorstore (Moss + Poppins)
│       └── secrets.toml.example # Template de config secrets
└── .gitignore                   # Filtre secrets, venv, outputs, debug...
```

---

## 🔧 Installation locale (pour développeurs)

```bash
git clone https://github.com/celiadaguet-code/ankorstore-scraper.git
cd ankorstore-scraper/streamlit_app
./run_local.command   # Mac (double-clic depuis Finder marche aussi)
# OU
run_local.bat         # Windows
```

Le script crée un venv Python local, installe les dépendances et lance Streamlit sur http://localhost:8501.

Pour la configuration des secrets (Google Sheet + Drive + password), voir [`streamlit_app/DEPLOY.md`](streamlit_app/DEPLOY.md).

---

## 🔐 Sécurité

- **Repo privé** sur GitHub (passé public uniquement le temps des déploiements)
- **Authentification password** sur l'app (suffisant pour usage interne — le scraping cible des sites publics, aucune donnée client Ankorstore en jeu)
- **Service Account Google** dédié, avec accès limité aux 2 ressources strictement nécessaires (1 Sheet + 1 dossier Drive)
- **Secrets jamais commit** (`.gitignore` strict avec wildcards type `*service_account*.json`)
- **Rotation clé GCP** : recommandée tous les 6-12 mois

---

## 🐛 Quand un scrape ne marche pas bien

L'app affiche systématiquement :
- ⚠️ **Warnings** détaillés (descriptions trop courtes, champs manquants, etc.) avec le nom du produit concerné
- 📋 **Produits exclus** par les filtres automatiques (ateliers, alcool, cartes cadeaux…) avec la raison
- 🟡 **Cellules surlignées en jaune** dans le xlsx pour les champs obligatoires à compléter à la main

Si une marque passe vraiment mal, **notifier Célia** avec l'URL de la marque concernée et un screenshot des warnings.

---

## 🛣️ Roadmap (idées d'évolution)

- [ ] Stats d'usage par AE (qui scrape combien, taux de conversion)
- [ ] Détection des erreurs récurrentes (warnings communs entre marques)
- [ ] Bouton "Pousser dans Ankorstore directement" (au lieu de télécharger le xlsx puis l'importer manuellement)
- [ ] Webhook Slack pour notifier les nouveaux scrapes dans le canal

---

## 📞 Maintenance

| Action | Procédure |
|---|---|
| **Mettre à jour le code** | 1. Repasser le repo en public sur GitHub → 2. Push → 3. Streamlit Cloud redéploie auto (~2-3 min) → 4. Repasser en privé |
| **Changer le password** | Streamlit Cloud → app → Settings → Secrets → modifier `app_password` → Save |
| **Voir les logs de prod** | Streamlit Cloud → app → "Manage app" → panneau logs |
| **Rotater la clé GCP** | GCP Console → IAM → Comptes de service → `scraper-bot` → Clés → Créer nouvelle + supprimer ancienne → relancer `setup_secrets.py` → MAJ Secrets Streamlit Cloud |

---

## 👤 Auteur

**Célia Daguet** — Team Lead AE, Ankorstore (mai 2026)

Projet construit avec assistance Claude (Anthropic) en 2 jours intensifs : du scraper Python jusqu'au déploiement web sécurisé pour l'équipe AE.
