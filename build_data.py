"""
build_data.py — Sourcing Command Center data pipeline
Run once to produce data.js which is loaded by index.html

Usage:  python3 build_data.py
Output: data.js
"""

import json
import pandas as pd
import numpy as np
from collections import defaultdict

# ─── CONFIGURABLE CONSTANTS (labeled ASSUMPTION where applicable) ─────────────
SPEND_THRESHOLD_PERCENTILE = 75          # ASSUMPTION: Kraljic spend axis cutoff
RISK_THRESHOLD              = 12         # ASSUMPTION: Kraljic risk axis cutoff

# Savings capture rates by segment (ASSUMPTION)
CAPTURE_RATES = {
    "Leverage":   0.85,
    "Strategic":  0.60,
    "Bottleneck": 0.50,
    "Routine":    0.75,
}

# Price verdict thresholds (ASSUMPTION)
NEGOTIATE_MULTIPLIER = 1.25   # supplier avg > 1.0× but <= 1.25× median → Negotiate
TOP_RANGE_PCT        = 0.10   # top 10% of [min,max] range → Push back

# Contract statuses that count as "off-contract"
OFF_CONTRACT_STATUSES = {
    "Expired",
    "Expiring Q2 2026",
    "Expiring Q3 2026",
    "Month-to-Month",
}

INPUT_FILE  = "dataset.xlsx"
OUTPUT_FILE = "data.js"

cleaning_log = []


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — LOAD & CLEAN
# ═══════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("STEP 1: Loading and cleaning data")
print("=" * 60)

# ── 1a. Spend Data ────────────────────────────────────────────────────────────
spend_raw = pd.read_excel(INPUT_FILE, sheet_name="Spend Data")
print(f"  Spend rows loaded: {len(spend_raw):,}")

# ── 1b. Risk Scorecard ────────────────────────────────────────────────────────
# Leading Unnamed:0 column → drop it; Risk Score and Risk Tier are formula
# strings in the raw file; recompute from the 4 component columns.
rs_raw = pd.read_excel(INPUT_FILE, sheet_name="Risk Scorecard")
rs_raw = rs_raw.drop(columns=["Unnamed: 0"])
rs_raw.columns = [
    "supplier_name", "supplier_id",
    "risk_score_raw", "risk_tier_raw",
    "quality_risk", "service_risk", "compliance_risk", "expense_risk",
]
# Recompute risk score from components (handles formula-string issue)
rs_raw["risk_score"] = (
    rs_raw["quality_risk"].astype(int)
    + rs_raw["service_risk"].astype(int)
    + rs_raw["compliance_risk"].astype(int)
    + rs_raw["expense_risk"].astype(int)
)
# Recompute risk tier
def compute_risk_tier(score):
    if score <= 7:
        return "Low"
    elif score <= 13:
        return "Medium"
    else:
        return "High"

rs_raw["risk_tier"] = rs_raw["risk_score"].apply(compute_risk_tier)
print(f"  Risk Scorecard rows: {len(rs_raw)}")

# ── 1c. Vendor Master ─────────────────────────────────────────────────────────
vm_raw = pd.read_excel(INPUT_FILE, sheet_name="Vendor Master")
vm_raw = vm_raw.drop(columns=["Unnamed: 0"])
vm_raw.columns = [
    "supplier_name", "supplier_id", "supplier_tier",
    "region", "risk_score_vm", "diversity",
]
print(f"  Vendor Master rows: {len(vm_raw)}")

# ── 1d. Savings Targets ───────────────────────────────────────────────────────
st_raw = pd.read_excel(INPUT_FILE, sheet_name="Savings Targets")
st_raw = st_raw.drop(columns=["Unnamed: 0"])
st_raw.columns = ["spend_category", "sub_category", "savings_pct"]
print(f"  Savings Targets rows: {len(st_raw)}")

# ── 1e. Clean Supplier IDs ────────────────────────────────────────────────────
def clean_supplier_id(sid):
    """Strip dash suffixes (e.g. SUP1328-L → SUP1328)."""
    s = str(sid).strip()
    if "-" in s:
        s = s.split("-")[0]
    return s

# Apply to spend data
dash_before = spend_raw["Supplier ID"].nunique()
spend_raw["Supplier ID"] = spend_raw["Supplier ID"].apply(clean_supplier_id)
dash_after = spend_raw["Supplier ID"].nunique()
cleaning_log.append(
    f"Stripped dash suffixes from Supplier IDs: {dash_before} → {dash_after} unique IDs in spend data"
)
print(f"  After dash-strip: {dash_before} → {dash_after} unique supplier IDs")

# Fix typo SUP4204 → SUP2404
typo_count = (spend_raw["Supplier ID"] == "SUP4204").sum()
spend_raw["Supplier ID"] = spend_raw["Supplier ID"].replace("SUP4204", "SUP2404")
cleaning_log.append(
    f"Remapped typo SUP4204 → SUP2404 ({typo_count} rows affected)"
)
print(f"  Remapped SUP4204 → SUP2404 ({typo_count} rows)")

# Apply same cleaning to Risk Scorecard and Vendor Master
rs_raw["supplier_id"] = rs_raw["supplier_id"].apply(clean_supplier_id)
vm_raw["supplier_id"] = vm_raw["supplier_id"].apply(clean_supplier_id)

# ── 1f. Build canonical name map from Vendor Master ──────────────────────────
# Vendor Master is the source of truth for supplier names
id_to_name  = dict(zip(vm_raw["supplier_id"], vm_raw["supplier_name"]))
id_to_tier  = dict(zip(vm_raw["supplier_id"], vm_raw["supplier_tier"]))
id_to_div   = dict(zip(vm_raw["supplier_id"], vm_raw["diversity"]))

# Build Risk Scorecard lookup
id_to_risk = {}
for _, row in rs_raw.iterrows():
    id_to_risk[row["supplier_id"]] = {
        "risk_score":       int(row["risk_score"]),
        "risk_tier":        row["risk_tier"],
        "quality_risk":     int(row["quality_risk"]),
        "service_risk":     int(row["service_risk"]),
        "compliance_risk":  int(row["compliance_risk"]),
        "expense_risk":     int(row["expense_risk"]),
    }

# Build savings target lookup (sub_category → savings_pct)
savings_lookup = dict(zip(st_raw["sub_category"], st_raw["savings_pct"]))

print(f"  Unique cleaned IDs in spend data: {spend_raw['Supplier ID'].nunique()}")
cleaning_log.append(
    f"Canonical supplier names sourced from Vendor Master ({len(vm_raw)} entries)"
)

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — CHECKPOINTS
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 2: Checkpoints")
print("=" * 60)

total_baseline  = spend_raw["Spend (USD)"].sum()
total_projected = spend_raw["Projected Spend (USD)"].sum()
total_increase  = total_projected - total_baseline

# Savings ceiling: for each transaction, look up its sub-category savings %
# then multiply by that row's Projected Spend (USD)
spend_raw["_savings_pct"] = spend_raw["Sub-Category"].map(savings_lookup).fillna(0)
spend_raw["_savings_row"]  = spend_raw["Projected Spend (USD)"] * spend_raw["_savings_pct"]
savings_ceiling = spend_raw["_savings_row"].sum()

unique_suppliers = spend_raw["Supplier ID"].nunique()

print(f"  Total baseline spend:  ${total_baseline:,.0f}  (~${total_baseline/1e6:.1f}M)")
print(f"  Total projected spend: ${total_projected:,.0f}  (~${total_projected/1e6:.1f}M)")
print(f"  Increase:              ${total_increase:,.0f}  (~${total_increase/1e6:.1f}M)")
print(f"  Savings ceiling:       ${savings_ceiling:,.0f}  (~${savings_ceiling/1e6:.1f}M)")
print(f"  Unique supplier IDs:   {unique_suppliers}")

assert abs(total_baseline  - 340_233_937) < 1e4, f"Baseline mismatch: {total_baseline}"
assert abs(total_projected - 391_942_017) < 1e4, f"Projected mismatch: {total_projected}"
assert unique_suppliers == 256, f"Expected 256 suppliers, got {unique_suppliers}"
print("  ✓ All checkpoints passed")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — PER-SUPPLIER AGGREGATION
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 3: Per-supplier aggregation")
print("=" * 60)

# ── 3a. Material-level median prices (global, across all suppliers) ───────────
# Weighted average unit price per material × supplier, then median across suppliers
mat_group = spend_raw.groupby("Material ID")
mat_stats = {}
for mat_id, grp in mat_group:
    prices  = grp["Unit Price"].values
    qtys    = grp["Quantity"].values
    mat_stats[mat_id] = {
        "description":    grp["Material Description"].iloc[0],
        "median_price":   float(np.median(prices)),
        "p90_price":      float(np.percentile(prices, 90)),
        "price_range_low":  float(prices.min()),
        "price_range_high": float(prices.max()),
        "purchase_count": int(len(grp)),
    }

# ── 3b. Supplier-level aggregation ───────────────────────────────────────────
supplier_records = []
supplier_ids = sorted(spend_raw["Supplier ID"].unique())

for sup_id in supplier_ids:
    grp = spend_raw[spend_raw["Supplier ID"] == sup_id]

    # ── Basic spend ──────────────────────────────────────────────────────────
    baseline_spend  = float(grp["Spend (USD)"].sum())
    projected_spend = float(grp["Projected Spend (USD)"].sum())
    savings_ceil_sup= float(grp["_savings_row"].sum())

    # ── Risk data ────────────────────────────────────────────────────────────
    risk_data = id_to_risk.get(sup_id, {
        "risk_score": 0, "risk_tier": "Unknown",
        "quality_risk": 0, "service_risk": 0,
        "compliance_risk": 0, "expense_risk": 0,
    })

    # ── Main category (by spend) ─────────────────────────────────────────────
    cat_spend = grp.groupby("Spend Category")["Spend (USD)"].sum()
    main_category = str(cat_spend.idxmax()) if not cat_spend.empty else "Unknown"

    # ── Main region (by spend) ───────────────────────────────────────────────
    reg_spend = grp.groupby("Region")["Spend (USD)"].sum()
    main_region = str(reg_spend.idxmax()) if not reg_spend.empty else "Unknown"

    # ── Off-contract % ───────────────────────────────────────────────────────
    off_contract_spend = grp[grp["Contract Status"].isin(OFF_CONTRACT_STATUSES)]["Spend (USD)"].sum()
    off_contract_pct   = float(off_contract_spend / baseline_spend) if baseline_spend > 0 else 0.0

    # ── Contract status mix ──────────────────────────────────────────────────
    cs_mix = grp.groupby("Contract Status")["Spend (USD)"].sum().to_dict()
    cs_mix = {k: float(v) for k, v in cs_mix.items()}

    # ── Payment terms mix ────────────────────────────────────────────────────
    pt_mix = grp.groupby("Payment Terms")["Spend (USD)"].sum().to_dict()
    pt_mix = {k: float(v) for k, v in pt_mix.items()}

    # ── Price index ──────────────────────────────────────────────────────────
    # Weighted avg unit price this supplier / weighted avg unit price same materials globally
    sup_mat_grp = grp.groupby("Material ID").apply(
        lambda x: pd.Series({
            "sup_avg_price": float(np.average(x["Unit Price"], weights=x["Quantity"])),
            "total_qty":     float(x["Quantity"].sum()),
            "total_spend":   float(x["Spend (USD)"].sum()),
        }),
        include_groups=False,
    ).reset_index()

    # Global weighted avg for those same materials
    sup_weighted_sum = 0.0
    glob_weighted_sum = 0.0
    total_qty_all = 0.0
    for _, mrow in sup_mat_grp.iterrows():
        mid = mrow["Material ID"]
        qty = mrow["total_qty"]
        sup_weighted_sum  += mrow["sup_avg_price"] * qty
        # Global median as reference
        glob_weighted_sum += mat_stats[mid]["median_price"] * qty
        total_qty_all += qty

    if glob_weighted_sum > 0:
        price_index = float(sup_weighted_sum / glob_weighted_sum)
    else:
        price_index = 1.0

    # ── Per-material price items ──────────────────────────────────────────────
    price_items = []
    total_dollar_gap = 0.0

    for _, mrow in sup_mat_grp.iterrows():
        mid       = mrow["Material ID"]
        ms        = mat_stats[mid]
        sup_avg   = mrow["sup_avg_price"]
        med_price = ms["median_price"]
        rng_low   = ms["price_range_low"]
        rng_high  = ms["price_range_high"]
        total_qty = mrow["total_qty"]

        # Price verdict
        top_10_threshold = rng_high - TOP_RANGE_PCT * (rng_high - rng_low)
        if sup_avg <= med_price:
            verdict = "Fair"
        elif sup_avg <= NEGOTIATE_MULTIPLIER * med_price and sup_avg < top_10_threshold:
            verdict = "Negotiate"
        else:
            verdict = "Push back"

        dollar_gap = max(0.0, (sup_avg - med_price) * total_qty)
        total_dollar_gap += dollar_gap

        price_items.append({
            "material_id":          mid,
            "material_description": ms["description"],
            "spend":                round(mrow["total_spend"], 2),
            "supplier_avg_price":   round(sup_avg, 4),
            "median_price":         round(med_price, 4),
            "purchase_count":       ms["purchase_count"],
            "verdict":              verdict,
            "dollar_gap":           round(dollar_gap, 2),
        })

    # Sort by spend descending, keep top 8 for battle card
    price_items.sort(key=lambda x: x["spend"], reverse=True)
    top_price_items = price_items[:8]

    # ── Countries ────────────────────────────────────────────────────────────
    countries = sorted(grp["Country"].dropna().unique().tolist())

    supplier_records.append({
        "id":               sup_id,
        "name":             id_to_name.get(sup_id, grp["Supplier Name"].iloc[0]),
        "baseline_spend":   round(baseline_spend, 2),
        "projected_spend":  round(projected_spend, 2),
        "spend_share_pct":  round(baseline_spend / total_baseline * 100, 4),
        "savings_ceiling":  round(savings_ceil_sup, 2),
        "risk_score":       risk_data["risk_score"],
        "risk_tier":        risk_data["risk_tier"],
        "quality_risk":     risk_data["quality_risk"],
        "service_risk":     risk_data["service_risk"],
        "compliance_risk":  risk_data["compliance_risk"],
        "expense_risk":     risk_data["expense_risk"],
        "main_category":    main_category,
        "main_region":      main_region,
        "supplier_tier":    id_to_tier.get(sup_id, "Unknown"),
        "diversity":        id_to_div.get(sup_id, None),
        "off_contract_pct": round(off_contract_pct, 4),
        "contract_status_mix": cs_mix,
        "payment_terms_mix":   pt_mix,
        "price_index":      round(price_index, 4),
        "total_dollar_gap": round(total_dollar_gap, 2),
        "countries":        countries,
        "price_items":      top_price_items,
    })

print(f"  Built records for {len(supplier_records)} suppliers")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — SEGMENTATION (KRALJIC)
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 4: Segmentation (Kraljic matrix)")
print("=" * 60)

projected_spends = [r["projected_spend"] for r in supplier_records]
spend_threshold  = float(np.percentile(projected_spends, SPEND_THRESHOLD_PERCENTILE))
print(f"  Spend threshold ({SPEND_THRESHOLD_PERCENTILE}th percentile): ${spend_threshold:,.0f}")
print(f"  Risk threshold: {RISK_THRESHOLD}")

segment_counts = defaultdict(int)
for rec in supplier_records:
    high_spend = rec["projected_spend"] >= spend_threshold
    high_risk  = rec["risk_score"]      >= RISK_THRESHOLD

    if high_spend and not high_risk:
        seg = "Leverage"
    elif high_spend and high_risk:
        seg = "Strategic"
    elif not high_spend and high_risk:
        seg = "Bottleneck"
    else:
        seg = "Routine"

    rec["segment"] = seg
    # Savings plan
    cr = CAPTURE_RATES[seg]
    rec["savings_low"]    = round(rec["savings_ceiling"] * 0.5, 2)
    rec["savings_high"]   = round(rec["savings_ceiling"] * 1.0, 2)
    rec["savings_target"] = round(rec["savings_ceiling"] * cr,  2)
    segment_counts[seg]  += 1

print(f"  Segment counts: {dict(segment_counts)}")
print(f"  (Expected approximately: Leverage ~48, Strategic ~16, Bottleneck ~88, Routine ~104)")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5 — SAVINGS PLAN TOTALS
# ═══════════════════════════════════════════════════════════════════════════════

base_case_savings = sum(r["savings_target"] for r in supplier_records)
print(f"\n  Base case total savings: ${base_case_savings:,.0f}  (~${base_case_savings/1e6:.1f}M)")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6 — SAVINGS TARGETS AND SUB-CATEGORY SPEND
# ═══════════════════════════════════════════════════════════════════════════════

subcat_spend = {}
for subcat, grp in spend_raw.groupby("Sub-Category"):
    subcat_spend[subcat] = {
        "baseline":  round(float(grp["Spend (USD)"].sum()), 2),
        "projected": round(float(grp["Projected Spend (USD)"].sum()), 2),
    }

# ═══════════════════════════════════════════════════════════════════════════════
# ASSEMBLE AND WRITE OUTPUT
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("Writing data.js")
print("=" * 60)

app_data = {
    "meta": {
        "total_baseline_spend":  round(total_baseline, 2),
        "total_projected_spend": round(total_projected, 2),
        "total_increase":        round(total_increase, 2),
        "savings_ceiling":       round(savings_ceiling, 2),
        "supplier_count":        256,
        "spend_threshold":       round(spend_threshold, 2),
        "risk_threshold":        RISK_THRESHOLD,
        "segment_counts":        dict(segment_counts),
        "capture_rates":         CAPTURE_RATES,
        "base_case_savings":     round(base_case_savings, 2),
        "cleaning_log":          cleaning_log,
    },
    "suppliers": supplier_records,
    "materials": mat_stats,
    "savings_targets": {k: round(float(v), 6) for k, v in savings_lookup.items()},
    "subcategory_spend": subcat_spend,
}

js_output = (
    "// Auto-generated by build_data.py — do not edit manually\n"
    "// Run: python3 build_data.py\n"
    "const APP_DATA = "
    + json.dumps(app_data, indent=2, ensure_ascii=False)
    + ";\n"
)

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    f.write(js_output)

size_kb = len(js_output.encode()) / 1024
print(f"  Written {OUTPUT_FILE}  ({size_kb:.0f} KB)")
print("\nDone.")
