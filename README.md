# Ankorstore Catalog Scraper

> **Internal Ankorstore tool** — Automatically scrapes a brand's product catalog and generates an `.xlsx` file ready to import into Ankorstore.

🔗 **Production app**: https://ankorstore-scraper.streamlit.app
📊 **Team history**: [Google Sheet](https://docs.google.com/spreadsheets/d/133RFUlvGqY34yE30j82QbKSYolKC0xuE22P4hjJM-ZY)

---

## 🎯 The problem we're solving

When an AE meets a prospect brand, importing their catalog into Ankorstore usually takes **2 to 3 hours of manual work** (copying name, price, description, images, variants, etc. for 50-500 products).

This tool does the same job in **30 seconds to 3 minutes** per brand, and generates an xlsx file **ready to import** into Ankorstore (1 row = 1 variant, mandatory fields highlighted in yellow when missing).

---

## 🚀 How to use it (for AEs)

1. Go to **https://ankorstore-scraper.streamlit.app**
2. Enter the shared password: `AnkorScrape2026!`
3. Enter your first name (saved for the session)
4. Paste the URL of a prospect brand (e.g. `https://cosmella.fr`)
5. Click **🚀 Scrape catalog**
6. Wait 30 seconds to 3 minutes depending on catalog size
7. **Download the xlsx** → import it into Ankorstore

The app automatically checks if a teammate has already scraped this brand recently (anti-duplication). Every scrape is logged in the shared Google Sheet.

---

## 📊 Supported CMS and platforms

| Platform | Status | Method |
|---|---|---|
| **WooCommerce** | ✅ | Store API v1 + sitemap HTML fallback |
| **PrestaShop** | ✅ | `data-product` JSON attribute + legacy HTML fallback |
| **Shopify** | ✅ | Paginated `/products.json` + sitemap fallback for restricted stores |
| **Wix** | ✅ | `productItems` JSON + JSON-LD + image upscaling |
| **Squarespace** | ✅ | `?format=json` trick + variants extraction |
| **SumUp Store** | ✅ | React Server Components flight data |
| **Custom sites** | ⚠️ | Cascade JSON-LD → Open Graph → microdata (best-effort), automatic fallback when no known CMS is detected |
| **Etsy / Facebook Shop** | ❌ | Not viable (strong anti-scraping) |

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────┐
│  Streamlit Cloud Community (hosting)        │
│  - Shared password auth                     │
│  - Ankorstore UI (Moss + Poppins)           │
│  - Session state (persisted results)        │
└─────────────────┬───────────────────────────┘
                  │  (lazy import to limit RAM)
                  ▼
┌─────────────────────────────────────────────┐
│  scrap_ankorstore.py  (~3300 lines)         │
│  - Automatic CMS detection                  │
│  - 7 specialized scrapers (Woo, Presta,     │
│    Shopify, Wix, Squarespace, SumUp,        │
│    custom best-effort)                      │
│  - Noise filters (alcohol, workshops,       │
│    gift cards, consultations…)              │
│  - Maps to Ankorstore xlsx template         │
│  - Yellow-highlights mandatory empty fields │
└──┬──────────────────────┬───────────────────┘
   │                      │
   ▼                      ▼
┌──────────────┐    ┌──────────────────────┐
│ Brand sites  │    │ Google Workspace     │
│ (HTTP/HTML   │    │ - History Sheet      │
│  scraping)   │    │ - Shared Drive       │
└──────────────┘    │   (xlsx archive)     │
                    └──────────────────────┘
```

---

## 🛠️ Tech stack

- **Python 3.11** (Mac / Linux / Streamlit Cloud compatible)
- **Streamlit 1.58** (web UI — no JS framework to maintain)
- **openpyxl** (xlsx template manipulation)
- **gspread + google-api-python-client** (shared history)
- **urllib + ThreadPoolExecutor** (HTTP scraping, no `requests`-style dependency)
- **Hosting**: Streamlit Cloud Community (free, 1 GB RAM plan)
- **Repo**: private GitHub (briefly made public for each update to allow Streamlit Cloud to re-pull)

---

## 📁 Project structure

```
ankorstore-scraper/
├── README.md                    # This file
├── scrap_ankorstore.py          # The scraper (~3300 lines, 7 CMS incl. Shopify)
├── ankorstore_template.xlsx     # Official Ankorstore template
├── probe_*.py                   # Per-CMS diagnostic scripts
├── diag_*.py                    # One-off diagnostics
├── streamlit_app/
│   ├── app.py                   # Streamlit UI
│   ├── sheets_drive.py          # Google Sheets + Drive module
│   ├── setup_secrets.py         # Initial config wizard
│   ├── requirements.txt         # Python dependencies
│   ├── DEPLOY.md                # Full deployment guide
│   ├── run_local.command        # Local launcher (Mac)
│   ├── run_local.bat            # Local launcher (Windows)
│   └── .streamlit/
│       ├── config.toml          # Ankorstore theme (Moss + Poppins)
│       └── secrets.toml.example # Secrets config template
└── .gitignore                   # Filters secrets, venv, outputs, debug…
```

---

## 🔧 Local install (for developers)

```bash
git clone https://github.com/celiadaguet-code/ankorstore-scraper.git
cd ankorstore-scraper/streamlit_app
./run_local.command   # Mac (double-clicking from Finder also works)
# OR
run_local.bat         # Windows
```

The script creates a local Python venv, installs dependencies and launches Streamlit on http://localhost:8501.

For secrets setup (Google Sheet + Drive + password), see [`streamlit_app/DEPLOY.md`](streamlit_app/DEPLOY.md).

---

## 🔐 Security

- **Private GitHub repo** (only made public briefly during deployments)
- **Password authentication** on the app (sufficient for internal use — scraping targets public sites, no Ankorstore customer data involved)
- **Dedicated Google Service Account** with access limited to the 2 strictly required resources (1 Sheet + 1 Drive folder)
- **Secrets never committed** (strict `.gitignore` with wildcards like `*service_account*.json`)
- **GCP key rotation**: recommended every 6–12 months

---

## 🐛 When a scrape doesn't work well

The app always shows:
- ⚠️ **Detailed warnings** (descriptions too short, missing fields, etc.) with the affected product name
- 📋 **Products excluded** by automatic filters (workshops, alcohol, gift cards…) with the reason
- 🟡 **Cells highlighted in yellow** in the xlsx for mandatory fields that need to be filled manually

If a brand really doesn't scrape well, **notify Célia** with the brand URL and a screenshot of the warnings.

---

## 🛣️ Roadmap (future ideas)

- [ ] Per-AE usage stats (who scrapes how much, conversion rate)
- [ ] Detection of recurring errors (warnings common across brands)
- [ ] "Push directly into Ankorstore" button (instead of downloading the xlsx and importing it manually)
- [ ] Slack webhook to notify new scrapes in the channel

---

## 📞 Maintenance

| Action | Procedure |
|---|---|
| **Update the code** | 1. Make the repo public on GitHub → 2. Push → 3. Streamlit Cloud auto-redeploys (~2-3 min) → 4. Make it private again |
| **Change the password** | Streamlit Cloud → app → Settings → Secrets → edit `app_password` → Save |
| **View production logs** | Streamlit Cloud → app → "Manage app" → logs panel |
| **Rotate the GCP key** | GCP Console → IAM → Service accounts → `scraper-bot` → Keys → Create new + delete old → re-run `setup_secrets.py` → Update Streamlit Cloud Secrets |

---

## 👤 Author

**Célia Daguet** — Team Lead AE, Ankorstore (May 2026)

Project built with Claude (Anthropic) assistance over 2 intense days: from Python scraper to secure web deployment for the AE team.
