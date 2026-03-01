# CLAUDE.md — LA 2028 Olympic Games Strategic Analysis
## SportsFanatics Consulting Agency | San Antonio, TX

---

## 🏟️ Project Overview

**Client:** SportsFanatics (fictitious sports consulting agency)
**Project:** Strategic Playbook for the 2028 Los Angeles Olympic Games
**Goal:** Deliver data-driven recommendations supporting three stakeholder groups:
- 🏃 **Athletes** — performance optimization, climate, venue readiness
- 🏙️ **The City of LA** — economic impact, infrastructure, legacy planning
- 🌍 **National Olympic Committees (NOCs)** — medal forecasting, investment strategy, geopolitical context

**Primary Analyst:** Shabeeb (Senior Statistical Programmer, ICON Plc)
**Deliverables:** Strategic consulting report, PowerPoint deck, data visualizations, and optional interactive dashboard

---

## 🗂️ Project Structure

```
la2028-olympics/
│
├── CLAUDE.md                   # This file — project instructions for Claude Code
├── README.md                   # Project summary for stakeholders
│
├── data/
│   ├── raw/                    # Original, unmodified source data
│   ├── processed/              # Cleaned and transformed datasets
│   └── external/               # Scraped or downloaded external datasets
│
├── notebooks/
│   ├── 01_data_profiling.ipynb         # Initial data exploration and QC
│   ├── 02_athlete_analysis.ipynb       # Pillar 1 — Athlete Edge
│   ├── 03_city_analysis.ipynb          # Pillar 2 — City Playbook
│   ├── 04_noc_analysis.ipynb           # Pillar 3 — NOC Intelligence
│   ├── 05_medal_forecasting.ipynb      # Predictive modeling
│   └── 06_visualizations.ipynb         # Final chart compilation
│
├── src/
│   ├── data_loader.py          # Functions for loading and validating data
│   ├── cleaning.py             # Data cleaning and standardization utilities
│   ├── analysis.py             # Core analytical functions
│   ├── visualizations.py       # Reusable chart/plot functions
│   ├── forecasting.py          # Medal prediction models
│   └── utils.py                # General helper functions
│
├── outputs/
│   ├── figures/                # Exported charts (PNG, HTML)
│   ├── tables/                 # Summary tables (CSV, Excel)
│   ├── report/                 # Final written report (Markdown or PDF)
│   └── presentation/           # PowerPoint deck assets
│
├── dashboard/
│   └── app.py                  # Optional Streamlit interactive dashboard
│
├── requirements.txt            # Python dependencies
└── .gitignore
```

---

## 🛠️ Tech Stack

| Purpose | Tool |
|---|---|
| Core Language | Python 3.11+ |
| Data Wrangling | pandas, numpy |
| Statistical Analysis | scipy, statsmodels |
| Machine Learning / Forecasting | scikit-learn |
| Interactive Visualization | Plotly, Plotly Express |
| Static Visualization | Matplotlib, Seaborn |
| Geospatial / Maps | Plotly choropleth, geopandas (if needed) |
| Notebooks | Jupyter Lab |
| Dashboard (optional) | Streamlit |
| Reporting | Markdown → PDF via nbconvert or Quarto |
| Presentation | python-pptx or export to PowerPoint manually |

---

## 📦 Environment Setup

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # Mac/Linux
venv\Scripts\activate     # Windows

# Install dependencies
pip install -r requirements.txt
```

### requirements.txt (baseline)
```
pandas>=2.0
numpy>=1.24
scipy>=1.11
statsmodels>=0.14
scikit-learn>=1.3
plotly>=5.18
matplotlib>=3.7
seaborn>=0.13
jupyter>=1.0
streamlit>=1.28
openpyxl>=3.1
python-pptx>=0.6
requests>=2.31
beautifulsoup4>=4.12
geopandas>=0.14
nbconvert>=7.0
```

---

## 📐 Analysis Framework

### Pillar 1 — The Athlete Edge
- Historical performance trends by sport, gender, and country
- Climate analysis: LA July/August heat, humidity, venue conditions
- Sport-by-sport venue mapping across LA
- Athlete participation growth by sport (including new 2028 sports)
- Performance delta: home country advantage analysis

### Pillar 2 — The City Playbook
- Economic impact modeling (tourism, revenue, infrastructure spend)
- Host city benchmarking: LA 1984, Atlanta 1996, Sydney 2000, Athens 2004, Rio 2016, Tokyo 2020, Paris 2024
- Legacy risk assessment framework
- LA-specific logistics: venue geography, transport, accommodation
- Sustainability and cost containment recommendations

### Pillar 3 — The NOC Intelligence Report
- Medal forecasting model by country and sport
- GDP vs. medal count correlation analysis
- Emerging nations on the rise (underdog tracker)
- Geopolitical context: US-China competition, Russia/Belarus exclusions
- New sports opportunity map: flag football, cricket, lacrosse, squash
- NOC investment allocation recommendations

---

## 🎨 Visualization Standards

- **Color Palette:** Use Olympic ring colors as primary theme
  - Blue: `#0085C7`, Yellow: `#F4C300`, Black: `#000000`, Green: `#009F6B`, Red: `#DF0024`
  - Neutral background: `#F8F9FA`
- **Chart Library:** Plotly as default for all interactive charts
- **Static exports:** Save as PNG at 300 DPI for report/presentation use
- **Chart titles:** Always include chart title, axis labels, data source, and year
- **Figure size standard:** 1200x700px for widescreen, 800x600px for square
- **Font:** Arial or Open Sans for consistency across all outputs

### Chart Naming Convention
```
{pillar}_{chart_type}_{subject}_{year}.html
# Example: noc_choropleth_medal_count_2024.html
#          athlete_bar_topCountries_historical.html
#          city_line_economicImpact_hostCities.html
```

---

## 🔢 Coding Standards

- **PEP 8** compliance throughout
- All functions must have docstrings with parameters and return types
- Use `logging` instead of `print()` for pipeline outputs
- Never hardcode file paths — use `pathlib.Path` and config variables
- All data transformations must be reproducible and version-tracked
- Raw data files are **never modified** — always write to `processed/`
- Commit meaningful checkpoints with descriptive messages

### Function Template
```python
def function_name(param1: type, param2: type) -> return_type:
    """
    Brief description of what this function does.

    Args:
        param1 (type): Description of param1.
        param2 (type): Description of param2.

    Returns:
        return_type: Description of what is returned.
    """
    pass
```

---

## 📊 Data Guidelines

raw data is located here: https://docs.google.com/spreadsheets/d/1pBJWvCsQFxzsywuDVdgeKtsrWCdnHa1IP8dCbMo5ywQ/edit?usp=sharing

### On First Receipt of Any Dataset
1. Run `01_data_profiling.ipynb` before any analysis
2. Document: shape, dtypes, null counts, duplicate rows, value ranges
3. Flag anomalies or data quality issues in a `data_quality_notes.md` file
4. Never assume data is clean — validate every key column

### Key Columns to Standardize
- Country names → map to ISO 3166-1 alpha-3 codes (e.g., USA, GBR, CHN)
- Sport names → standardize to official IOC sport nomenclature
- Year → always integer, Games edition clearly labeled
- Medal types → capitalize: Gold, Silver, Bronze

---

## 🤖 Instructions for Claude Code

- **Always start** a new analysis task by reviewing this CLAUDE.md file
- **Before writing code**, state the analytical goal and the expected output
- **Prefer modular functions** over inline scripts — build `src/` utilities progressively
- **When building visualizations**, always use the color palette and naming conventions above
- **For forecasting**, explain model choice, assumptions, and limitations before presenting results
- **When uncertain** about data interpretation, flag it explicitly rather than assuming
- **After completing each notebook**, summarize: what was found, what was visualized, what is recommended
- **Error handling**: wrap all file I/O and API calls in try/except blocks
- **Output exports**: always save final figures to `outputs/figures/` and tables to `outputs/tables/`
- **Progress updates**: at each major milestone, print a summary of what was completed and what comes next

### Workflow Order
```
1. Data Profiling → 2. Cleaning → 3. Athlete Analysis →
4. City Analysis → 5. NOC Analysis → 6. Forecasting →
7. Visualization Compilation → 8. Report → 9. Presentation
```

---

## 📅 Milestone Tracker

| # | Milestone | Status |
|---|---|---|
| 1 | Project setup and CLAUDE.md finalized | ✅ Done |
| 2 | Data profiling and quality assessment | ✅ Done |
| 3 | Pillar 1 — Athlete Edge analysis | ✅ Done |
| 4 | Pillar 2 — City Playbook analysis | ✅ Done |
| 5 | Pillar 3 — NOC Intelligence analysis | ✅ Done |
| 6 | Medal forecasting model | ✅ Done |
| 7 | Final visualization compilation | ✅ Done |
| 8 | Written strategic report | ✅ Done |
| 9 | PowerPoint consulting deck | ✅ Done |
| 10 | Optional Streamlit dashboard | ✅ Done |

---

## 📝 Notes & Decisions Log

*Use this section to track key analytical decisions made during the project.*

| Date | Decision | Rationale |
|---|---|---|
| 2026-03-01 | Selected Python + Plotly as primary stack | Unified environment, consulting-grade visuals |
| 2026-03-01 | Three-pillar framework adopted | Stakeholder-aligned structure for recommendations |

Notes:

To run daashboard app:

cd "/Volumes/D Drive/Data analysis/Olympic data analysis"
/opt/homebrew/bin/python3.11 -m streamlit run dashboard/app.py
---

*Last updated: March 2026 | Analyst: Shabeeb | Project: LA28 Olympic Strategic Playbook*