"""
build_data.py — Sourcing Command Center data pipeline
Reads Edited_Aura_Savings_Model.xlsx as source of truth.

Usage:  python3 build_data.py
Output: data.js
"""

import json
import pandas as pd
import numpy as np
from collections import defaultdict

# ─── CONSTANTS ────────────────────────────────────────────────────────────────
INPUT_FILE  = "Edited_Aura_Savings_Model.xlsx"
OUTPUT_FILE = "data.js"

# Thresholds (from Inputs sheet, confirmed)
SPEND_THRESHOLD  = 1_800_000   # Leverage / Negotiate vs Tail boundary
RISK_HIGH_MIN    = 14          # risk >= 14 → High risk (overrides spend)
RISK_LOW_MAX     = 7           # risk <= 7  → Low tier (Leverage if high spend)

# Price verdict thresholds
NEGOTIATE_MULTIPLIER = 1.25    # supplier avg > mat avg but <= 1.25× → Negotiate
TOP_RANGE_PCT        = 0.10    # top 10% of [min,max] range → Push back

# Contract windows that are renegotiable (= "off-contract")
OFF_CONTRACT_WINDOWS = {"Open now", "Renews this year"}

cleaning_log = []


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1 — LOAD SHEETS
# ═══════════════════════════════════════════════════════════════════════════════

print("=" * 60)
print("STEP 1: Loading sheets from", INPUT_FILE)
print("=" * 60)

calc_raw = pd.read_excel(INPUT_FILE, sheet_name="Calc")
print(f"  Calc rows loaded: {len(calc_raw):,}")

raw_data = pd.read_excel(INPUT_FILE, sheet_name="Raw_Data")
print(f"  Raw_Data rows loaded: {len(raw_data):,}")

vm_raw = pd.read_excel(INPUT_FILE, sheet_name="Vendor_Master")
print(f"  Vendor_Master rows: {len(vm_raw)}")

rs_raw = pd.read_excel(INPUT_FILE, sheet_name="Risk_Scorecard")
print(f"  Risk_Scorecard rows: {len(rs_raw)}")

sm_raw = pd.read_excel(INPUT_FILE, sheet_name="3_Savings_Model", header=None)
print(f"  3_Savings_Model loaded")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2 — SUP4204 FIX AND JOIN CALC WITH RAW_DATA
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 2: SUP4204 fix and enriching Calc with Raw_Data")
print("=" * 60)

# Fix SUP4204 → SUP2404 in Calc (12 rows with #N/A segment)
typo_count = (calc_raw["Clean supplier ID"] == "SUP4204").sum()
calc_raw["Clean supplier ID"] = calc_raw["Clean supplier ID"].replace("SUP4204", "SUP2404")
cleaning_log.append(f"Remapped SUP4204 → SUP2404 in Calc ({typo_count} rows)")
print(f"  Remapped SUP4204 → SUP2404 ({typo_count} rows in Calc)")

# Join Calc with Raw_Data on Transaction ID to get Country, Payment Terms,
# Material ID (canonical), Material Description, Spend Category
join_cols = ["Transaction ID", "Country", "Payment Terms",
             "Material ID", "Material Description", "Spend Category"]
rd_lookup = raw_data[join_cols].drop_duplicates("Transaction ID").set_index("Transaction ID")

calc = calc_raw.join(rd_lookup, on="Transaction ID", how="left")
unmatched = calc["Country"].isna().sum()
if unmatched > 0:
    print(f"  Warning: {unmatched} Calc rows have no Raw_Data match")
print(f"  Joined Calc with Raw_Data: {len(calc):,} rows")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3 — BUILD LOOKUP TABLES
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 3: Building lookup tables")
print("=" * 60)

# Vendor Master lookups
vm_raw.columns = ["supplier_name", "supplier_id", "supplier_tier",
                  "region", "risk_score_vm", "diversity"]
id_to_name = dict(zip(vm_raw["supplier_id"], vm_raw["supplier_name"]))
id_to_tier = dict(zip(vm_raw["supplier_id"], vm_raw["supplier_tier"]))
id_to_region = dict(zip(vm_raw["supplier_id"], vm_raw["region"]))
id_to_div  = dict(zip(vm_raw["supplier_id"], vm_raw["diversity"]))

# Risk Scorecard lookups
rs_raw.columns = [
    "supplier_name", "supplier_id", "risk_score_raw", "risk_tier_raw",
    "quality_risk", "service_risk", "compliance_risk", "expense_risk",
]
rs_raw["risk_score"] = (
    rs_raw["quality_risk"].astype(int)
    + rs_raw["service_risk"].astype(int)
    + rs_raw["compliance_risk"].astype(int)
    + rs_raw["expense_risk"].astype(int)
)

def compute_risk_tier(score):
    if score <= 7:   return "Low"
    elif score <= 13: return "Medium"
    else:             return "High"

rs_raw["risk_tier"] = rs_raw["risk_score"].apply(compute_risk_tier)

id_to_risk = {}
for _, row in rs_raw.iterrows():
    id_to_risk[row["supplier_id"]] = {
        "risk_score":      int(row["risk_score"]),
        "risk_tier":       row["risk_tier"],
        "quality_risk":    int(row["quality_risk"]),
        "service_risk":    int(row["service_risk"]),
        "compliance_risk": int(row["compliance_risk"]),
        "expense_risk":    int(row["expense_risk"]),
    }

print(f"  Vendor Master: {len(id_to_name)} suppliers")
print(f"  Risk Scorecard: {len(id_to_risk)} suppliers")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4 — CHECKPOINTS
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 4: Checkpoints")
print("=" * 60)

total_baseline  = calc["Last-year spend"].sum()
total_projected = calc["Projected spend"].sum()
total_pool      = calc["Case savings pool"].sum()
total_l1_3      = calc["Levers 1-3 savings"].sum()
unique_suppliers = calc["Clean supplier ID"].nunique()

print(f"  Unique suppliers in Calc:  {unique_suppliers}")
print(f"  LY spend:                  ${total_baseline:,.0f}  (~${total_baseline/1e6:.2f}M)")
print(f"  Projected spend:           ${total_projected:,.0f}  (~${total_projected/1e6:.2f}M)")
print(f"  Case savings pool:         ${total_pool:,.0f}  (~${total_pool/1e6:.2f}M)")
print(f"  Levers 1-3 savings:        ${total_l1_3:,.0f}  (~${total_l1_3/1e6:.2f}M)")

assert unique_suppliers == 256, f"Expected 256 suppliers, got {unique_suppliers}"
assert abs(total_baseline  - 340_233_937) < 1e5, f"LY spend mismatch: {total_baseline}"
assert abs(total_projected - 391_942_017) < 1e5, f"Projected spend mismatch: {total_projected}"
assert abs(total_pool      -  47_000_000) < 1e6, f"Pool mismatch: {total_pool}"
assert abs(total_l1_3      -  35_900_000) < 2e5, f"L1-3 mismatch: {total_l1_3}"
print("  ✓ All checkpoints passed")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5 — MATERIAL STATS (using Material avg price from Calc)
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 5: Material stats")
print("=" * 60)

mat_stats = {}
for mat_id, grp in calc.groupby("Material ID"):
    if pd.isna(mat_id):
        continue
    # Material avg price is the Excel-computed reference price (same for all rows)
    mat_avg_price = float(grp["Material avg price"].iloc[0])
    mat_low_price = float(grp["Material lowest price"].iloc[0])
    prices = grp["Unit price"].values

    desc_col = grp["Material Description"].dropna()
    description = desc_col.iloc[0] if not desc_col.empty else str(mat_id)

    mat_stats[str(mat_id)] = {
        "description":      str(description),
        "avg_price":        mat_avg_price,
        "lowest_price":     mat_low_price,
        "price_range_low":  float(prices.min()),
        "price_range_high": float(prices.max()),
        "purchase_count":   int(len(grp)),
    }

print(f"  Built stats for {len(mat_stats)} materials")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6 — SEGMENTATION
# ═══════════════════════════════════════════════════════════════════════════════

def compute_segment(proj_spend, risk_score):
    """3-condition segmentation matching the Excel model thresholds."""
    if risk_score >= RISK_HIGH_MIN:
        return "High risk"
    if proj_spend >= SPEND_THRESHOLD and risk_score <= RISK_LOW_MAX:
        return "Leverage"
    if proj_spend >= SPEND_THRESHOLD and risk_score > RISK_LOW_MAX:
        return "Negotiate"
    return "Tail"


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 7 — PER-SUPPLIER AGGREGATION
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 7: Per-supplier aggregation")
print("=" * 60)

supplier_records = []
supplier_ids = sorted(calc["Clean supplier ID"].unique())

for sup_id in supplier_ids:
    grp = calc[calc["Clean supplier ID"] == sup_id]

    # ── Spend ────────────────────────────────────────────────────────────────
    baseline_spend  = float(grp["Last-year spend"].sum())
    projected_spend = float(grp["Projected spend"].sum())
    savings_pool    = float(grp["Case savings pool"].sum())
    savings_l1_3    = float(grp["Levers 1-3 savings"].sum())

    # ── Risk data (from Risk Scorecard) ─────────────────────────────────────
    risk_data = id_to_risk.get(sup_id, {
        "risk_score": int(grp["Risk score"].median()),
        "risk_tier":  compute_risk_tier(int(grp["Risk score"].median())),
        "quality_risk": 0, "service_risk": 0,
        "compliance_risk": 0, "expense_risk": 0,
    })

    # ── Segment ──────────────────────────────────────────────────────────────
    segment = compute_segment(projected_spend, risk_data["risk_score"])

    # ── Main category (by spend) ─────────────────────────────────────────────
    cat_spend = grp.groupby("Category")["Projected spend"].sum()
    main_category = str(cat_spend.idxmax()) if not cat_spend.empty else "Unknown"

    # ── Region (from Vendor Master, authoritative per supplier) ──────────────
    main_region = str(id_to_region.get(sup_id, "Unknown"))

    # ── Off-contract % (contract window is Open now or Renews this year) ────
    off_contract_spend = grp[
        grp["Contract window"].isin(OFF_CONTRACT_WINDOWS)
    ]["Projected spend"].sum()
    off_contract_pct = float(off_contract_spend / projected_spend) if projected_spend > 0 else 0.0

    # ── Contract window mix (by projected spend) ─────────────────────────────
    cs_mix = grp.groupby("Contract window")["Projected spend"].sum().to_dict()
    cs_mix = {k: float(v) for k, v in cs_mix.items()}

    # ── Payment terms mix (by projected spend, from Raw_Data join) ───────────
    if "Payment Terms" in grp.columns:
        pt_mix = grp.groupby("Payment Terms")["Projected spend"].sum().dropna().to_dict()
        pt_mix = {str(k): float(v) for k, v in pt_mix.items() if pd.notna(k)}
    else:
        pt_mix = {}

    # ── Per-material price items ──────────────────────────────────────────────
    price_items = []
    total_dollar_gap = 0.0

    for mat_id_val, mgrp in grp.groupby("Material ID"):
        if pd.isna(mat_id_val):
            continue
        mid_str = str(mat_id_val)
        ms = mat_stats.get(mid_str)
        if ms is None:
            continue

        # Weighted avg unit price for this supplier × material
        quantities = mgrp["Quantity"].values
        prices     = mgrp["Unit price"].values
        total_qty  = float(quantities.sum())
        total_proj = float(mgrp["Projected spend"].sum())

        if total_qty > 0:
            sup_avg_price = float(np.average(prices, weights=quantities))
        else:
            sup_avg_price = float(prices.mean()) if len(prices) > 0 else 0.0

        mat_avg  = ms["avg_price"]
        rng_low  = ms["price_range_low"]
        rng_high = ms["price_range_high"]

        # Verdict
        top_threshold = rng_high - TOP_RANGE_PCT * (rng_high - rng_low)
        if sup_avg_price <= mat_avg:
            verdict = "Fair"
        elif sup_avg_price <= NEGOTIATE_MULTIPLIER * mat_avg and sup_avg_price < top_threshold:
            verdict = "Negotiate"
        else:
            verdict = "Push back"

        dollar_gap = max(0.0, (sup_avg_price - mat_avg) * total_qty)
        total_dollar_gap += dollar_gap

        desc_col = mgrp["Material Description"].dropna()
        mat_desc = desc_col.iloc[0] if not desc_col.empty else ms["description"]

        price_items.append({
            "material_id":          mid_str,
            "material_description": str(mat_desc),
            "spend":                round(total_proj, 2),
            "supplier_avg_price":   round(sup_avg_price, 4),
            "avg_price":            round(mat_avg, 4),
            "purchase_count":       int(len(mgrp)),
            "verdict":              verdict,
            "dollar_gap":           round(dollar_gap, 2),
        })

    price_items.sort(key=lambda x: x["spend"], reverse=True)
    top_price_items = price_items[:8]

    # ── Price index (supplier avg / material avg, qty-weighted) ──────────────
    sup_price_sum  = sum(pi["supplier_avg_price"] * pi["purchase_count"] for pi in price_items)
    glob_price_sum = sum(pi["avg_price"]          * pi["purchase_count"] for pi in price_items)
    total_px_count = sum(pi["purchase_count"] for pi in price_items)
    if glob_price_sum > 0:
        price_index = float(sup_price_sum / glob_price_sum)
    else:
        price_index = 1.0

    # ── Countries ────────────────────────────────────────────────────────────
    if "Country" in grp.columns:
        countries = sorted(grp["Country"].dropna().unique().tolist())
    else:
        countries = []

    # ── Spend category breakdown (for concentration panel) ───────────────────
    if "Spend Category" in grp.columns:
        sc_mix = grp.groupby("Spend Category")["Projected spend"].sum().to_dict()
        sc_mix = {str(k): float(v) for k, v in sc_mix.items() if pd.notna(k)}
    else:
        sc_mix = {}

    supplier_records.append({
        "id":                   sup_id,
        "name":                 id_to_name.get(sup_id, sup_id),
        "baseline_spend":       round(baseline_spend, 2),
        "projected_spend":      round(projected_spend, 2),
        "spend_share_pct":      round(baseline_spend / total_baseline * 100, 4),
        "savings_ceiling":      round(savings_pool, 2),
        "savings_target":       round(savings_l1_3, 2),
        "savings_low":          round(savings_l1_3 * 0.75, 2),
        "savings_high":         round(savings_l1_3 * 1.25, 2),
        "risk_score":           risk_data["risk_score"],
        "risk_tier":            risk_data["risk_tier"],
        "quality_risk":         risk_data["quality_risk"],
        "service_risk":         risk_data["service_risk"],
        "compliance_risk":      risk_data["compliance_risk"],
        "expense_risk":         risk_data["expense_risk"],
        "segment":              segment,
        "main_category":        main_category,
        "main_region":          main_region,
        "supplier_tier":        str(id_to_tier.get(sup_id, "Unknown")),
        "diversity":            id_to_div.get(sup_id),
        "off_contract_pct":     round(off_contract_pct, 4),
        "contract_status_mix":  cs_mix,
        "payment_terms_mix":    pt_mix,
        "spend_category_mix":   sc_mix,
        "price_index":          round(price_index, 4),
        "total_dollar_gap":     round(total_dollar_gap, 2),
        "countries":            countries,
        "price_items":          top_price_items,
    })

print(f"  Built records for {len(supplier_records)} suppliers")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 8 — SEGMENT COUNT PASS/FAIL CHECK
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 8: Segment count PASS/FAIL check")
print("=" * 60)

segment_counts = defaultdict(int)
for rec in supplier_records:
    segment_counts[rec["segment"]] += 1

EXPECTED_COUNTS = {
    "Leverage":  13,
    "Negotiate": 51,
    "Tail":      163,
    "High risk": 29,
}

all_pass = True
for seg, expected in EXPECTED_COUNTS.items():
    actual = segment_counts.get(seg, 0)
    status = "PASS" if actual == expected else "FAIL"
    if status == "FAIL":
        all_pass = False
    print(f"  {status}  {seg:12s}: expected {expected:3d}, got {actual:3d}")

if all_pass:
    print("  ✓ All segment counts match")
else:
    print("  ✗ Segment count mismatch — check thresholds")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 9 — READ HEADLINE NUMBERS FROM 3_Savings_Model
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("STEP 9: Reading headline numbers from 3_Savings_Model")
print("=" * 60)

# Row indices (0-based) confirmed from sheet inspection:
#   Row 11: Lever 1 total row  → col 5 = L1 savings
#   Row 22: Lever 2 total row  → col 5 = L2 savings
#   Row 33: Lever 3 total row  → col 5 = L3 savings
#   Row 41: Lever 4 savings    → col 1
#   Row 48: Lever 5 savings    → col 1
#   Row 58: Official target    → col 1
#   Row 63: Sensitivity Low    → col 2
#   Row 64: Sensitivity Base   → col 2
#   Row 65: Sensitivity High   → col 2

L1_savings     = float(sm_raw.iloc[11, 5])
L2_savings     = float(sm_raw.iloc[22, 5])
L3_savings     = float(sm_raw.iloc[33, 5])
L4_savings     = float(sm_raw.iloc[41, 1])
L5_savings     = float(sm_raw.iloc[48, 1])
official_target = float(sm_raw.iloc[58, 1])
sens_low       = float(sm_raw.iloc[63, 2])
sens_base      = float(sm_raw.iloc[64, 2])
sens_high      = float(sm_raw.iloc[65, 2])

# Realization rates from Lever 1 table (rows 5-10, col 4)
real_rates = {
    "Open now":          {"Other": float(sm_raw.iloc[5, 4]),  "High risk": float(sm_raw.iloc[6, 4])},
    "Renews this year":  {"Other": float(sm_raw.iloc[7, 4]),  "High risk": float(sm_raw.iloc[8, 4])},
    "Locked (active)":   {"Other": float(sm_raw.iloc[9, 4]),  "High risk": float(sm_raw.iloc[10, 4])},
}

print(f"  L1: ${L1_savings/1e6:.2f}M  L2: ${L2_savings/1e6:.2f}M  L3: ${L3_savings/1e6:.2f}M")
print(f"  L4: ${L4_savings/1e6:.3f}M  L5: ${L5_savings/1e6:.3f}M")
print(f"  Official target: ${official_target/1e6:.3f}M")
print(f"  Sensitivity: low=${sens_low/1e6:.1f}M  base=${sens_base/1e6:.1f}M  high=${sens_high/1e6:.1f}M")
print(f"  Realization rates: {real_rates}")


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 10 — SUB-CATEGORY SPEND
# ═══════════════════════════════════════════════════════════════════════════════

subcat_spend = {}
for subcat, grp in calc.groupby("Sub-category"):
    if pd.isna(subcat):
        continue
    subcat_spend[str(subcat)] = {
        "baseline":  round(float(grp["Last-year spend"].sum()), 2),
        "projected": round(float(grp["Projected spend"].sum()), 2),
    }


# ═══════════════════════════════════════════════════════════════════════════════
# ASSEMBLE AND WRITE OUTPUT
# ═══════════════════════════════════════════════════════════════════════════════

print("\n" + "=" * 60)
print("Writing data.js")
print("=" * 60)

app_data = {
    "meta": {
        # Spend totals
        "total_baseline_spend":  round(total_baseline, 2),
        "total_projected_spend": round(total_projected, 2),
        "total_increase":        round(total_projected - total_baseline, 2),
        # Savings
        "savings_ceiling":       round(total_pool, 2),     # Case savings pool
        "base_case_savings":     round(total_l1_3, 2),     # Levers 1-3
        "official_target":       round(official_target, 2),# All 5 levers
        "l1_savings":            round(L1_savings, 2),
        "l2_savings":            round(L2_savings, 2),
        "l3_savings":            round(L3_savings, 2),
        "l4_savings":            round(L4_savings, 2),
        "l5_savings":            round(L5_savings, 2),
        # Sensitivity
        "sensitivity_low":       round(sens_low, 2),
        "sensitivity_base":      round(sens_base, 2),
        "sensitivity_high":      round(sens_high, 2),
        # Realization rates
        "realization_rates":     real_rates,
        # Supplier/segment metadata
        "supplier_count":        256,
        "spend_threshold":       SPEND_THRESHOLD,
        "risk_threshold":        RISK_HIGH_MIN,
        "risk_low_max":          RISK_LOW_MAX,
        "segment_counts":        dict(segment_counts),
        # Data quality log
        "cleaning_log":          cleaning_log,
    },
    "suppliers":        supplier_records,
    "materials":        mat_stats,
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
