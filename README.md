# Exchange Rate Determination — Zambia ZMW/USD Analysis

A data-driven executive presentation and Power BI report analysing Zambia's exchange rate policy, currency depreciation trends, bid-ask spread dynamics, and inflation differentials against the United States (2015–2023).

---

## Contents

| File | Description |
|------|-------------|
| `Analysis.pbix` | Power BI report — source of all data and DAX measures |
| `Exchange_Rate_Determination_Zambia.pptx` | Generated 8-slide executive presentation |
| `generate_presentation.py` | Python script that builds the PPTX from the PBI data |
| `boz_logo_transparent.png` | Bank of Zambia logo used on the cover slide |

---

## Presentation Overview

**`Exchange_Rate_Determination_Zambia.pptx`** is an 8-slide boardroom-quality presentation built with a navy blue and white corporate theme.

| Slide | Title | Key Elements |
|-------|-------|-------------|
| 1 | Cover | BOZ logo, headline KPI strip, metadata bar |
| 2 | Agenda | 4 colour-coded section cards |
| 3 | Executive Summary | 5 KPI cards, key findings, sparkline chart |
| 4 | Exchange Rate Policy | Rate vs policy bounds area chart, PPP target |
| 5 | Depreciation & Bid-Ask Analysis | YoY depreciation bar chart, bid-ask trend |
| 6 | Inflation Analysis | US vs Zambia grouped bar chart, differential line |
| 7 | Policy Verdict & Recommendations | Gauge, verdict banner, 4 action items |
| 8 | Conclusion | Summary statistics, closing remarks |

---

## Key Metrics (2023)

| Metric | Value |
|--------|-------|
| Avg Selling Rate (ZMW/USD) | 20.89 |
| PPP Optimal Target Rate | 19.50 |
| Upper Policy Bound | 22.98 |
| YoY Depreciation | +21.3% |
| Cumulative Depreciation (2015–2023) | +142.9% |
| Bid-Ask Spread | 3.1% |
| Zambia Avg Inflation | 13.5% |
| US Avg Inflation | 4.1% |
| Inflation Differential | 9.4 pp |
| Policy Verdict | Depreciation Pressure |

---

## Power BI Report Structure

**Pages**
1. Overview — navigation and branding
2. Exchange Rate Policy — selling rate, policy bounds, PPP target, variance
3. Depreciation & Bid-Ask Rate Analysis — YoY depreciation, spread trends
4. Inflation & Policy Verdict — US vs Zambia CPI, differential, verdict

**DAX Measures**
- `Avg Selling Rate (USD)` — average ZMW/USD exchange rate
- `Bid-Ask Spread %` — percentage spread between bid and ask prices
- `Inflation Differential` — Zambia CPI minus US CPI
- `Lower Policy Bound` — minimum acceptable rate threshold
- `Optimal PPP Target Rate` — purchasing power parity implied rate
- `Policy Variance %` — deviation from the upper policy bound
- `Policy Verdict` — text classification of current rate alignment
- `US Avg Inflation %` — average US CPI rate
- `Upper Policy Bound` — maximum acceptable rate threshold
- `YoY Depreciation %` — year-on-year ZMW depreciation percentage
- `Zambia Avg Inflation %` — average Zambia CPI rate

**Data Tables**
- `DateTable` — year dimension for filtering
- `Inflation US` — US annual inflation by year
- `Inflation Zambia` — Zambia annual inflation by year
- `Measures` — DAX-only calculated measures table

---

## Regenerating the Presentation

Requirements: Python 3.x, `python-pptx`, `matplotlib`, `Pillow`

```bash
pip install python-pptx matplotlib Pillow
python generate_presentation.py
```

Output: `Exchange_Rate_Determination_Zambia.pptx` (widescreen 13.33" × 7.5")

---

## Design Specifications

- **Colour palette:** Navy `#09152F` · Mid-navy `#122855` · Accent blue `#1E88E5` · Gold `#F5A623`
- **Slide size:** 13.33 × 7.5 inches (16:9 widescreen)
- **Charts:** matplotlib at 150 DPI, embedded as transparent PNG
- **Typography:** Calibri throughout
- **KPI cards:** 2.5pt thick rounded-border cards with horizontal accent separator

---

## Analysis Period

2015 – 2023 · Currency pair: ZMW/USD · Source: Bank of Zambia / IMF data via Power BI model

---

*Prepared by FSi Outsourcing — Analytics & Advisory · k.banda@fsioutsourcing.com · May 2026 · Confidential*
