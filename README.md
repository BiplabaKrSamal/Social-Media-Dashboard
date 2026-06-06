<div align="center">



<br/>

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Excel](https://img.shields.io/badge/Google_Sheets-Ready-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)](https://sheets.google.com)
[![Power BI](https://img.shields.io/badge/Power_BI-Interactive_HTML-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)](./LRZ_PowerBI_Dashboard.html)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Analytics-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](#linkedin)
[![Instagram](https://img.shields.io/badge/Instagram-Insights-E1306C?style=for-the-badge&logo=instagram&logoColor=white)](#instagram)
[![ArtStation](https://img.shields.io/badge/ArtStation-Stats-13AFF0?style=for-the-badge&logo=artstation&logoColor=white)](#artstation)

<br/>

**A production-grade social media analytics pipeline and dashboard system for [Little Red Zombies](https://www.littleredzombies.com/) — AAA 3D character outsourcing studio.**

*Drop in a CSV export → run one command → every chart, metric, and insight auto-updates.*

---

[**📊 View Dashboard**](#-dashboard-preview) · [**⚡ Quick Start**](#-quick-start-60-seconds) · [**📁 Data Sources**](#-platform-data-sources) · [**🏗 Architecture**](#-architecture)

</div>

---

## 🎯 What This Does — In 10 Seconds

```
You export a CSV from LinkedIn.    →   Drop it in data/
You export a CSV from Instagram.   →   Drop it in data/
You fill 1 row in a spreadsheet.   →   (ArtStation)

python lrz_etl.py

✅  28 months of data written.  0 formula errors.  Dashboard live.
```

That's it. Every chart, KPI, trend line, and priority flag in the dashboard **auto-updates** from the raw data sheet via live formulas.

---

## 📊 Dashboard Preview

<table>
<tr>
<td width="50%">

**📋 5-Sheet Google Sheets Workbook**
- KPI snapshot with MoM growth badges
- Platform trend table (6-month rolling)  
- Top 10 posts log ranked by impressions
- Content mix breakdown
- Priority action flags (🔴🟡🟢)
- `⬇ Import Zone` — live ETL status sheet

</td>
<td width="50%">

**🖥 Interactive Power BI–Style HTML Dashboard**
- 4 drill-down views (Overview / Trends / Posts / Scorecard)
- Platform filter sidebar (toggle on/off per platform)
- Time period selector (6mo / 12mo / All Time)
- Health scoring (0–100) vs industry benchmarks
- Animated charts via Chart.js (no install needed)

</td>
</tr>
</table>

| Sheet | Purpose |
|-------|---------|
| `📊 Dashboard` | Executive overview — KPIs, trends, top posts, content mix |
| `⬇ Import Zone` | Live ETL run status + step-by-step export instructions |
| `📋 Raw Data` | 28 months of monthly data, all platforms, live formulas |
| `🎯 Platform Deep Dive` | Per-platform health profiles, benchmarks, action items |
| `📌 Assumptions` | Full data transparency — source, confidence %, methodology |
| `📆 Content Calendar` | Pre-built content planning template for next 8 weeks |

---

## ⚡ Quick Start — 60 Seconds

```bash
# 1. Clone
git clone https://github.com/YOUR_USERNAME/lrz-social-dashboard.git
cd lrz-social-dashboard

# 2. Install (3 packages)
pip install openpyxl pandas python-dateutil

# 3. Drop your exports into data/
#    See 👇 Platform Data Sources for export instructions

# 4. Run
python lrz_etl.py

# ✅ Dashboard updated. Open LRZ_Social_Media_Dashboard.xlsx
```

> **No LinkedIn/Instagram access?** The repo ships with **sample CSVs** in `data/templates/` — the pipeline runs fully out of the box on demo data.

---

## 📁 Platform Data Sources

### LinkedIn — 2 files, 30 seconds to export

| File | Where to get it |
|------|----------------|
| `data/linkedin_followers.csv` | `linkedin.com/company/[YOUR-PAGE]/admin/analytics/followers/` → **Export** |
| `data/linkedin_updates.csv` | `linkedin.com/company/[YOUR-PAGE]/admin/analytics/updates/` → **Export** |

**Columns used:** `Date` · `Total followers` · `Impressions (organic)` · `Reactions (organic)` · `Comments (organic)` · `Shares (organic)`

---

### Instagram — 1 file, 60 seconds to export

| File | Where to get it |
|------|----------------|
| `data/instagram_insights.csv` | `business.facebook.com` → Insights → Overview → **Export Data** → Type: Account → CSV |

**Columns used:** `Period` · `Impressions` · `Accounts reached` · `Likes` · `Comments` · `Shares` · `Saves` · `Follows`

> ℹ️ Instagram exports weekly data. The pipeline **automatically aggregates to monthly**.

---

### ArtStation — manual template (2 min/month)

ArtStation has no CSV export. Fill one row per month:

```bash
# Open and edit
data/artstation_manual.csv
```

```csv
Month,    Followers, Total Views, Total Likes, Total Comments, Posts Published
Jan 2024, 3200,      6800,        285,         38,             2
Feb 2024, 3320,      7200,        302,         41,             2
...
```

> Lines starting with `#` are comments — ignored by the parser.

---

### X/Twitter & Facebook

No reliable free CSV export exists for either platform.

- Both are **estimated** from follower-based impression multipliers and RivalIQ 2024 industry benchmarks
- To upgrade to live data: connect **Sprout Social**, **Buffer Analytics**, or **Hootsuite** — then add a parser following the `ParseResult` pattern in `lrz_etl.py`

---

## 🏗 Architecture

```
 ┌─────────────────────────────────────────────────────────────┐
 │                    PLATFORM EXPORTS (CSV)                    │
 │  linkedin_followers.csv  │  instagram_insights.csv          │
 │  linkedin_updates.csv    │  artstation_manual.csv           │
 └──────────────┬──────────────────────────┬────────────────────┘
                │                          │
                ▼                          ▼
 ┌──────────────────────────────────────────────────────────────┐
 │                      lrz_etl.py                              │
 │                                                              │
 │   parse_linkedin()    ──►  ParseResult(linkedin)             │
 │   parse_instagram()   ──►  ParseResult(instagram)            │
 │   parse_artstation()  ──►  ParseResult(artstation)           │
 │   estimated_twitter() ──►  ParseResult(twitter)              │
 │   estimated_facebook()──►  ParseResult(facebook)             │
 │                              │                               │
 │   Normalise to monthly schema (YYYY-MM key)                  │
 │   Aggregate: impressions · engagements · followers · posts   │
 └──────────────────────────────┬───────────────────────────────┘
                                │
                                ▼
 ┌──────────────────────────────────────────────────────────────┐
 │           LRZ_Social_Media_Dashboard.xlsx                    │
 │                                                              │
 │   ⬇ Import Zone   ←  run status · errors · instructions     │
 │   📋 Raw Data      ←  live data written here                 │
 │                                                              │
 │   📊 Dashboard  ──────────────────────────────►  auto        │
 │   🎯 Deep Dive  ──── reference Raw Data via ──►  updates     │
 │   📌 Assumptions     live Excel formulas     ►  on open      │
 │   📆 Calendar   ──────────────────────────────►              │
 └──────────────────────────────────────────────────────────────┘
                                │
                                ▼
            lrz_etl_log.json  (audit trail, every run)
```

---

## 📐 Data Model

Every platform is normalised to this monthly schema:

| Field | Type | Source | Formula in XLSX |
|-------|------|--------|----------------|
| `month` | `YYYY-MM` | All | — |
| `followers` | `int` | LinkedIn/IG/AS | — |
| `mom_change` | `int` | Derived | `=followers[m] - followers[m-1]` |
| `mom_pct` | `float` | Derived | `=IFERROR(mom_change/followers[m-1], 0)` |
| `posts` | `int` | All | — |
| `impressions` | `int` | LinkedIn/IG/AS | — |
| `engagements` | `int` | All | `likes + comments + shares (+ saves for IG)` |
| `eng_rate` | `float` | Derived | `=IFERROR(engagements/impressions, 0)` |
| `avg_likes` | `int` | Derived | `likes × 0.75` |
| `reach` | `int` | IG native, others estimated | — |

---

## 📊 Key Metrics (Apr 2026 Snapshot)

| Platform | Followers | MoM Growth | Eng Rate | vs Benchmark | Priority |
|----------|-----------|------------|----------|--------------|----------|
| 🔵 LinkedIn | **8,664** | +6.2% | 4.1% | ✅ +2.1pp | `HIGH` |
| 🔴 Instagram | 3,124 | +1.9% | 3.8% | ✅ +0.8pp | `MEDIUM` |
| 🐦 X/Twitter | 1,284 | +0.8% | 1.2% | 🔴 -0.8pp | `HIGH RISK` |
| 🔷 ArtStation | **4,820** | +4.5% | **5.6%** | ✅ +1.6pp | `HEALTHY` |
| 📘 Facebook | 2,218 | +0.4% | 0.8% | 🔴 -0.7pp | `DEPRIORITISE` |

> **Top finding:** ArtStation has the highest engagement rate (5.6%). Systematic cross-posting ArtStation → LinkedIn adds ~38% total reach per project at zero extra content cost.

---

## 📂 Repository Structure

```
lrz-social-dashboard/
│
├── 📊 LRZ_Social_Media_Dashboard.xlsx   ← Open this in Google Sheets / Excel
├── 🖥  LRZ_PowerBI_Dashboard.html        ← Open this in any browser
│
├── lrz_etl.py                            ← Main ETL pipeline (run this monthly)
├── requirements.txt
├── .gitignore
│
├── data/
│   ├── linkedin_followers.csv            ← Drop your export here
│   ├── linkedin_updates.csv              ← Drop your export here
│   ├── instagram_insights.csv            ← Drop your export here
│   ├── artstation_manual.csv             ← Fill this monthly
│   └── templates/                        ← Empty starter templates
│       ├── linkedin_followers_template.csv
│       ├── linkedin_updates_template.csv
│       ├── instagram_insights_template.csv
│       └── artstation_manual_template.csv
│
├── docs/
│   ├── EXPORT_GUIDE.md                   ← Platform-by-platform export walkthrough
│   └── DATA_DICTIONARY.md                ← Every metric defined
│
├── examples/
│   └── lrz_etl_log.json                  ← Sample run log
│
└── README.md
```

---

## 🔁 Monthly Refresh Workflow

```
First week of each month:

1. LinkedIn  →  Export 2 CSVs  →  place in data/       (2 min)
2. Instagram →  Export 1 CSV   →  place in data/       (2 min)
3. ArtStation → Fill 1 new row in artstation_manual.csv (1 min)
4. Terminal  →  python lrz_etl.py                      (5 sec)

Done. Every KPI, chart, and alert in the dashboard is now current.
```

**Optional: automate with cron**
```bash
# Runs every Monday at 08:00
0 8 * * 1 cd /path/to/repo && python lrz_etl.py >> lrz_etl_cron.log 2>&1
```

---

## ⚙️ Assumptions & Data Confidence

| Platform | Followers | Impressions | Eng Rate | Confidence |
|----------|-----------|-------------|----------|------------|
| LinkedIn | ⚠️ Estimated | ✅ Live export | ✅ Calculated | 60–95% |
| Instagram | ✅ Real (3,124 verified) | ✅ Live export | ✅ Calculated | 95%+ |
| ArtStation | ⚠️ Manual entry | ⚠️ Manual entry | ✅ Calculated | 70% |
| X/Twitter | 🔴 Modelled | 🔴 Modelled | 🔴 Modelled | ~40% |
| Facebook | 🔴 Modelled | 🔴 Modelled | 🔴 Modelled | ~40% |

Full methodology in `📌 Assumptions` sheet and [`docs/DATA_DICTIONARY.md`](docs/DATA_DICTIONARY.md).

---

## 🛠 Tech Stack

| Layer | Tool | Why |
|-------|------|-----|
| ETL / data wrangling | `pandas` + `python-dateutil` | Handles messy date formats, weekly→monthly rollup |
| XLSX authoring | `openpyxl` | Full styling control — colours, formulas, freeze panes |
| Dashboard (HTML) | `Chart.js 4.4` | Zero-install, works offline, full interactivity |
| Data storage | `.xlsx` | Client-native, opens in Google Sheets or Excel |
| Audit log | `JSON` | Machine-readable run history per platform |

---

## 📦 Installation

```bash
pip install -r requirements.txt
```

```
openpyxl>=3.1.0
pandas>=2.0.0
python-dateutil>=2.8.0
```

**Python 3.10+ required.** No database, no cloud services, no API keys needed.

---

## 🤝 Contributing

1. Fork the repo
2. Add a new platform parser in `lrz_etl.py` following the `ParseResult` pattern
3. Add sample CSV in `data/templates/`
4. PR with `[PLATFORM]` prefix in title

See [`CONTRIBUTING.md`](CONTRIBUTING.md) for the full guide.

---

## 📄 License

MIT License — use freely, credit appreciated.

---

<div align="center">

**Built for [Little Red Zombies](https://www.littleredzombies.com/) · Hyderabad, India · May 2026**

*AAA 3D character art outsourcing · Naughty Dog · Obsidian · Rocksteady · Fatshark*

<br/>

[![Made with Python](https://img.shields.io/badge/Made%20with-Python-1f425f?style=flat-square&logo=python)](https://python.org)
[![Open in Sheets](https://img.shields.io/badge/Open%20in-Google%20Sheets-34A853?style=flat-square&logo=google-sheets&logoColor=white)](https://sheets.google.com)
[![No API Keys](https://img.shields.io/badge/No%20API%20Keys-Required-success?style=flat-square)](.)
[![Zero Dependencies on Cloud](https://img.shields.io/badge/Works-Offline-blue?style=flat-square)](.)

</div>
