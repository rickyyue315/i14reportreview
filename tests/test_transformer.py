"""Regression tests for transformer.py calculation logic.

Each test targets one of the six calculation issues identified in code review:
  #1 Combined Overstock Cap (X) computed inconsistently depending on dataset
  #2 Combined increase_only mode: Recommend vs Safety Qty Change mismatch
  #3 Overstock Review RP Type not normalized (case/whitespace)
  #4 OOS allow_both mode: Recommendation vs Safety Qty Increase mismatch for "other"
  #5 Carry SKU dashboard SKU count vs pivot Grand_Total mismatch
  #6 Combined mtd_days=None crashes
"""

import pandas as pd
import pytest

from transformer import (
    _add_combined_recommendation_columns,
    _add_recommendation_columns,
    _add_oos_recommendation_columns,
    _add_nd_recommendation_columns,
    transform_nd,
    transform_carry,
    _build_carry_dashboard,
)


def _base_row(**kw):
    row = {
        "Buyer": "B1",
        "Brand": "BrandA",
        "SKU": "1001",
        "Site Code": "S01",
        "Item_Shop": "Some Product-S01",
        "Shop Stock Status": "Normal",
        "Adj. Safety Stock": 5,
        "Shop Monthly Sold Qty (P90D)": 10,
        "Sold Qty (CM - MTD)": 8.0,
        "Shop Nature": "Local",
        "RP Type": "RF",
        "Shop Stock Qty (CM)": 0,
        "Supply Source": "Warehouse",
        "OM": "OM1",
        "Adj. MCH1": "MchA",
        "Adj. SST": "SstA",
    }
    row.update(kw)
    return row


def _df(rows):
    return pd.DataFrame([_base_row(**r) for r in rows])


_RATES_O = {"default": 0.5, "tourist": 0.3, "mix": 0.4}
_RATES_P = {"default": 0.5, "tourist": 0.5, "mix": 0.5}


# ── Fix #1 ────────────────────────────────────────────────────────────────────

def test_fix1_combined_x_computed_without_overstock_status_rows():
    """X must be computed for every row even when no row has Overstock status."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal",
         "Shop Monthly Sold Qty (P90D)": 10, "Adj. Safety Stock": 5},
        {"SKU": "2", "Shop Stock Status": "OOS (0 Stock)",
         "Shop Monthly Sold Qty (P90D)": 20, "Adj. Safety Stock": 5},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    # floor(10*0.5)=5 capped at adj_ss 5 -> 5 ; floor(20*0.5)=10 capped at 5 -> 5
    assert out["Overstock Cap (X)"].tolist() == [5, 5]


def test_fix1_combined_x_independent_of_overstock_row_presence():
    """Shared rows must yield identical X whether or not an Overstock row is present."""
    base = _df([
        {"SKU": "1", "Shop Stock Status": "Normal",
         "Shop Monthly Sold Qty (P90D)": 10, "Adj. Safety Stock": 8},
        {"SKU": "2", "Shop Stock Status": "OOS (0 Stock)",
         "Shop Monthly Sold Qty (P90D)": 20, "Adj. Safety Stock": 8},
    ])
    with_os = pd.concat([base, _df([
        {"SKU": "3", "Shop Stock Status": "Overstock (>2*Monthly Sold Qty & >2pcs)",
         "Shop Monthly Sold Qty (P90D)": 100, "Adj. Safety Stock": 8},
    ])], ignore_index=True)

    out_base = _add_combined_recommendation_columns(
        base, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    out_with = _add_combined_recommendation_columns(
        with_os, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out_base["Overstock Cap (X)"].tolist() == out_with["Overstock Cap (X)"].tolist()[:2]
    assert out_base["Recommend Safety Stock"].tolist() == out_with["Recommend Safety Stock"].tolist()[:2]


# ── Fix #2 ────────────────────────────────────────────────────────────────────

def test_fix2_combined_increase_only_recommend_synced_with_change():
    """In increase_only mode, when combined < adj_ss, Recommend must stay at adj_ss
    (consistent with Safety Qty Change == 0), not show the lowered combined value."""
    df = _df([
        {"SKU": "1",
         "Shop Stock Status": "Overstock (>2*Monthly Sold Qty & >2pcs)",
         "Shop Monthly Sold Qty (P90D)": 4, "Adj. Safety Stock": 10,
         "Sold Qty (CM - MTD)": 1.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    # X=floor(4*0.5)=2->clip3, Z=10 -> Z>X -> combined=max(Y,3)=3 ; delta=-7->clip0
    assert out["Safety Qty Change"].tolist()[0] == 0
    # Recommended must equal current adj_ss (no decrease applied)
    assert out["Recommend Safety Stock"].tolist()[0] == 10


# ── Fix #3 ────────────────────────────────────────────────────────────────────

def test_fix3_overstock_rp_type_normalized():
    """RP Type with whitespace/case must still be recognized for P90D 0 Sales rows."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": " rf "},
        {"SKU": "2", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "nd"},
    ])
    out = _add_recommendation_columns(df, rates=_RATES_O)
    assert out["Recommendation of new Safety Qty / ND"].tolist() == ["ND", "0"]


# ── Fix #4 ────────────────────────────────────────────────────────────────────

def test_fix4_oos_allow_both_other_recommend_consistent():
    """In allow_both mode, for an 'other' RP type row, Recommendation must equal
    adj_ss + Safety Qty Increase (other rows are frozen, so both should equal adj_ss)."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "ST",
         "Adj. Safety Stock": 10, "Sold Qty (CM - MTD)": 2.0,
         "Shop Monthly Sold Qty (P90D)": 4},
    ])
    out = _add_oos_recommendation_columns(
        df, _RATES_O, mtd_days=10, month_total_days=30,
        safety_mode="allow_both", oos_calc_mode="mtd",
    )
    rec = int(out["Recommendation of new Safety Qty"].tolist()[0])
    inc = int(out["Safety Qty Increase"].tolist()[0])
    assert inc == 0
    assert rec == 10  # frozen at adj_ss, not the lowered raw rec (3)


# ── Fix #5 ────────────────────────────────────────────────────────────────────

def test_fix5_carry_dashboard_counts_only_nd_rf():
    """Dashboard SKU total must match pivot scope (ND+RF only), excluding other RP types."""
    df = _df([
        {"SKU": "1", "RP Type": "RF", "Shop Stock Qty (CM)": 5, "Shop Stock Status": "Normal"},
        {"SKU": "2", "RP Type": "ND", "Shop Stock Qty (CM)": 5, "Shop Stock Status": "Normal"},
        {"SKU": "3", "RP Type": "XX", "Shop Stock Qty (CM)": 5, "Shop Stock Status": "Normal"},
    ])
    dash = _build_carry_dashboard(df)
    total = dash[dash["Site Code"] == "總計"].iloc[0]
    assert int(total["SKU 總數"]) == 2          # only RF + ND
    assert int(total["RF SKU 總數"]) == 1
    assert int(total["ND SKU 總數"]) == 1
    assert int(total["SKU 總數"]) == int(total["RF SKU 總數"]) + int(total["ND SKU 總數"])


def test_fix5_carry_dashboard_matches_pivot_via_transform():
    df = _df([
        {"SKU": "1", "RP Type": "RF", "Shop Stock Qty (CM)": 5, "Shop Stock Status": "Normal"},
        {"SKU": "2", "RP Type": "ND", "Shop Stock Qty (CM)": 5, "Shop Stock Status": "Normal"},
        {"SKU": "3", "RP Type": "XX", "Shop Stock Qty (CM)": 5, "Shop Stock Status": "Normal"},
    ])
    summary_left, summary_right, dashboard_df, raw_df, warnings = transform_carry(df)
    dash_total = dashboard_df[dashboard_df["Site Code"] == "總計"].iloc[0]
    # pivot Grand_Total over detail rows (exclude subtotal rows containing '合計')
    detail = summary_left[~summary_left["Site Code"].astype(str).str.contains("合計")]
    pivot_grand = int(detail["Grand_Total"].sum())
    assert int(dash_total["SKU 總數"]) == pivot_grand


# ── Fix #6 ────────────────────────────────────────────────────────────────────

def test_fix6_combined_mtd_days_none_no_crash():
    """mtd_days=None must not raise; Projected Monthly Sold Qty falls back to 0."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "Sold Qty (CM - MTD)": 8.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=None, month_total_days=None, safety_mode="increase_only"
    )
    assert out["Projected Monthly Sold Qty"].tolist()[0] == 0


# ── SKU 精簡 Review (ND SKU) ───────────────────────────────────────────────────

def test_nd_warehouse_all_nd_rows_recommend_full_transfer():
    """Every RP Type=ND + Warehouse row must recommend transferring all stock out."""
    df = _df([
        {"SKU": "1", "RP Type": "ND", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 7, "Shop Stock Status": "Normal"},
        {"SKU": "2", "RP Type": "ND", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 12, "Shop Stock Status": "Overstock (P90D 0 Sales)"},
    ])
    out = _add_nd_recommendation_columns(df)
    assert out["Recommendation of new Safety Qty / ND"].tolist() == ["ND", "ND"]
    assert out["After new Safety Qty / ND can transfer out of QTY"].tolist() == [7, 12]


def test_nd_vendor_marked_stay_at_shop():
    """ND + Vendor stock must be marked '留在原鋪' with transfer=0."""
    df = _df([
        {"SKU": "1", "RP Type": "ND", "Supply Source": "Vendor", "Shop Stock Qty (CM)": 9, "Shop Stock Status": "Normal"},
    ])
    out = _add_nd_recommendation_columns(df)
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "留在原鋪"
    assert out["After new Safety Qty / ND can transfer out of QTY"].tolist()[0] == 0


def test_nd_rp_type_normalized():
    """RP Type with whitespace/case must still be recognized as ND."""
    df = _df([
        {"SKU": "1", "RP Type": " nd ", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 5, "Shop Stock Status": "Normal"},
        {"SKU": "2", "RP Type": "Nd", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 3, "Shop Stock Status": "Normal"},
    ])
    out = _add_nd_recommendation_columns(df)
    assert out["Recommendation of new Safety Qty / ND"].tolist() == ["ND", "ND"]
    assert out["After new Safety Qty / ND can transfer out of QTY"].tolist() == [5, 3]


def test_nd_supply_source_normalized():
    """Supply Source with case/whitespace must still be recognized."""
    df = _df([
        {"SKU": "1", "RP Type": "ND", "Supply Source": " warehouse ", "Shop Stock Qty (CM)": 6, "Shop Stock Status": "Normal"},
        {"SKU": "2", "RP Type": "ND", "Supply Source": "VENDOR", "Shop Stock Qty (CM)": 4, "Shop Stock Status": "Normal"},
    ])
    out = _add_nd_recommendation_columns(df)
    assert out["Recommendation of new Safety Qty / ND"].tolist() == ["ND", "留在原鋪"]
    assert out["After new Safety Qty / ND can transfer out of QTY"].tolist() == [6, 0]


def test_nd_zero_stock_transfer_zero():
    """ND Warehouse with Shop Stock Qty=0 must yield transfer=0 (row kept)."""
    df = _df([
        {"SKU": "1", "RP Type": "ND", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 0, "Shop Stock Status": "Normal"},
    ])
    out = _add_nd_recommendation_columns(df)
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND"
    assert out["After new Safety Qty / ND can transfer out of QTY"].tolist()[0] == 0


def test_nd_non_nd_left_blank():
    """Non-ND rows must keep recommendation/transfer columns blank."""
    df = _df([
        {"SKU": "1", "RP Type": "RF", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 9, "Shop Stock Status": "Normal"},
        {"SKU": "2", "RP Type": "ST", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 4, "Shop Stock Status": "Normal"},
    ])
    out = _add_nd_recommendation_columns(df)
    assert out["Recommendation of new Safety Qty / ND"].tolist() == ["", ""]
    assert out["After new Safety Qty / ND can transfer out of QTY"].tolist() == ["", ""]


def test_nd_categories_by_supply_source_x_status():
    """Summary must break ND SKUs into Supply Source × Shop Stock Status categories."""
    df = _df([
        {"SKU": "1", "RP Type": "ND", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 5, "Shop Stock Status": "Overstock (P90D 0 Sales)"},
        {"SKU": "2", "RP Type": "ND", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 3, "Shop Stock Status": "Normal"},
        {"SKU": "3", "RP Type": "ND", "Supply Source": "Vendor", "Shop Stock Qty (CM)": 2, "Shop Stock Status": "Normal"},
        {"SKU": "4", "RP Type": "RF", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 8, "Shop Stock Status": "Normal"},
    ])
    summary, raw, categories, warnings = transform_nd(df)
    # Two-layer labels: "Warehouse - Status" and "Vendor - Status".
    assert set(categories) == {
        "Warehouse - Overstock (P90D 0 Sales)",
        "Warehouse - Normal",
        "Vendor - Normal",
    }
    total_row = summary[summary["Buyer"] == "總計"].iloc[0]
    assert int(total_row["總計"]) == 3  # 3 ND rows (RF excluded)
    assert int(total_row["Warehouse - Normal"]) == 1
    assert int(total_row["Warehouse - Overstock (P90D 0 Sales)"]) == 1
    assert int(total_row["Vendor - Normal"]) == 1
    # Transfer total = 5 + 3 only (Vendor row contributes 0, RF excluded).
    assert int(total_row["After new Safety Qty / ND can transfer out of QTY"]) == 8
