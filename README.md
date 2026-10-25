# Sourcing Command Center
### GEP × Aura Everyday Solutions — Case Competition Demo

---

## Files

| File | Purpose |
|------|---------|
| `dataset.xlsx` | Source data (4 sheets, 8 000 spend rows) |
| `build_data.py` | Python pipeline — run once, produces `data.js` |
| `data.js` | Generated JSON embedded as a JS constant (auto-generated, ~770 KB) |
| `index.html` | Single-page front end — open directly in a browser |
| `README.md` | This file |

---

## Quick start

```bash
# 1. Install dependencies (one-time)
pip install pandas openpyxl numpy

# 2. Build the data file
python3 build_data.py

# 3. Open the app
#    Option A — direct file open (most features work)
open index.html

#    Option B — local server (recommended, avoids any CORS edge cases)
python3 -m http.server 8000
# then visit http://localhost:8000
```

---

## Pipeline summary (`build_data.py`)

### Step 1 — Load & clean
- Reads all 4 Excel sheets with pandas/openpyxl
- **Risk Scorecard**: Excel formulas are not evaluated; Risk Score is recomputed as the sum of the four component columns (Product Quality + Service + Compliance + Expense). Risk Tier recomputed: ≤7 Low, ≤13 Medium, else High.
- **Supplier ID cleaning**:
  - Strip suffix after `-` (10 IDs: e.g. SUP1328-L → SUP1328) — *ASSUMPTION: suffixes are location/lot codes, not separate legal entities*
  - Remap SUP4204 → SUP2404 — *ASSUMPTION: data-entry typo confirmed by Vendor Master*
  - Result: 267 raw → 256 canonical IDs
- Canonical supplier names come from Vendor Master (22 IDs had multiple name spellings in Spend Data)

### Step 2 — Checkpoints (hard assertions)
| Metric | Expected | Actual |
|--------|----------|--------|
| Baseline spend | ~$340.2M | $340,233,937 |
| Projected spend | ~$391.9M | $391,942,017 |
| Increase | ~$51.7M | $51,708,080 |
| Savings ceiling | ~$47.0M | $47,033,042 |
| Unique suppliers | 256 | 256 |

Pipeline stops with an error if any checkpoint fails.

### Step 3 — Per-supplier aggregation (27 fields)
Key computed fields:
- `off_contract_pct` — share of spend where Contract Status ∈ {Expired, Expiring Q2 2026, Expiring Q3 2026, Month-to-Month} — *ASSUMPTION: these four statuses represent "at-risk" contract coverage*
- `price_index` — spend-weighted average unit price / overall material-weighted median. Values >1 mean this supplier prices above average.
- `total_dollar_gap` — total spend-dollars paid above the Aura internal median price across all materials

### Step 4 — Segmentation (Kraljic-style)
*ASSUMPTION: thresholds are configurable constants in the script*

| Axis | Default | Rationale |
|------|---------|-----------|
| Spend | 75th percentile of projected spend (~$1.84M) | Top quarter of suppliers |
| Risk | Score ≥ 12 | Mid-point of the 4–20 range |

| Segment | Count | Strategy |
|---------|-------|---------|
| Leverage | 48 | Competitive bid, volume bundling |
| Strategic | 16 | Joint cost reduction, index-linked LTA |
| Bottleneck | 88 | Secure supply, qualify second source |
| Routine | 104 | Consolidate, standardize, automate |

### Step 5 — Price fairness (per material, per supplier)
- Benchmark = Aura internal median unit price across all suppliers for same Material ID
- **NOT an external market price** — labeled clearly in the UI
- Verdicts: Fair (≤median), Negotiate (≤1.25× median), Push back (>1.25× median OR in top 10% of price range)

### Step 6 — Savings plan
*ASSUMPTION: capture rates represent expected negotiation success, not guaranteed outcomes*

| Segment | Capture rate | Rationale |
|---------|-------------|-----------|
| Leverage | 85% | High power, low risk → strong negotiating position |
| Strategic | 60% | Relationship-sensitive, partial savings only |
| Bottleneck | 50% | Supply security takes priority over price |
| Routine | 75% | Consolidation yields reliable savings |

**Base case savings: ~$34.3M** (vs $47.0M ceiling and $51.7M gap)

---

## Front-end summary (`index.html`)

### Tabs
1. **Dashboard** — Live Kraljic matrix + KPI strip. Drag sliders to move thresholds; dots recolor live.
2. **Battle Card** — Click any dot or search by name. Shows spend, risk, contract mix, price benchmarks, and recommended lever. Approve/Edit/Reject persisted to `localStorage`.
3. **Savings Tracker** — Running total, progress bar, table of decisions. "Load demo decisions" pre-approves top 10 Leverage suppliers.
4. **Risk Simulator** — *Labeled illustrative.* Three scenario buttons flag affected suppliers and suggest responses.
5. **Methodology** — Full FACT/ESTIMATE/ASSUMPTION table, data-cleaning log, and "Ask the Data" keyword filter.

### What is and isn't AI
- **Rules/statistics** (not AI): segmentation, price verdicts, savings math, scenario flags, talking-point templates
- **Optional LLM** (not required for demo): the "Ask the Data" box falls back to a keyword filter if no API key is present — the demo works fully offline

---

## 60-second demo script (click path)

> **Goal:** Show judges that Aura can capture $34M+ of the $51.7M cost increase through a data-driven procurement strategy.

1. **Open** index.html → point to KPI strip: "$51.7M increase, $47M ceiling, base case $34.3M captured"
2. **Dashboard tab** → gesture at the Kraljic matrix: "256 suppliers, four segments, thresholds are live"
3. **Drag risk slider** to 14 → "Watch the Strategic quadrant grow as we raise the risk bar"  → drag back to 12
4. **Click a Leverage dot** (large blue) → Battle Card opens
5. **Battle Card** → walk through risk bars, contract mix, price verdict table ("Push back on MAT-XXX")
6. **Click Approve** → savings counter ticks up in header
7. **Savings Tracker tab** → click "Load demo decisions" → counter fills to ~$14M with 10 approvals
8. **Risk Simulator tab** → click "Palm Oil Price Spike +15%" → dots pulse red, panel lists affected Natural Oils suppliers
9. **Methodology tab** → point to ASSUMPTION rows: "Every judgment call is labeled, numbers are traceable to source rows"
10. **Close:** "One command to rebuild, one file to present — and every number is auditable."

---

## Key assumptions (all labeled in UI)

| # | Assumption | Impact if wrong |
|---|-----------|----------------|
| A1 | Dash suffixes in Supplier IDs are location codes, not separate entities | Would overcount suppliers |
| A2 | SUP4204 is a typo for SUP2404 | One supplier's spend misattributed |
| A3 | Expiring Q2/Q3 2026 + Expired + M2M = "off-contract" | Off-contract % shifts |
| A4 | 75th-percentile spend = high-impact threshold | Segment counts change |
| A5 | Risk ≥ 12 = high risk | Segment counts change |
| A6 | Capture rates (85/60/50/75%) are negotiation targets | Base-case savings change |
| A7 | Internal median price = fair-price benchmark | Verdict distribution changes |
| A8 | Scenarios in Risk Simulator are illustrative, not live data | — |
