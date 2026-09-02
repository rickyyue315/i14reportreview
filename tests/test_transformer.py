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
    _clean_dataframe,
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
    # Shop Monthly Sold Qty (P90D) 已是月均: floor(10*0.5)=5 ; floor(20*0.5)=10→cap at Adj.SS=5
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
    # P90D 為 90 日總量: X=floor(4/3*0.5)=floor(0.67)=0→clip3→3, Z=10>X → combined=max(Y,3)=3 ; delta=-7→clip0
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
    """ND + Vendor stock must be marked '另有安排' with transfer=0."""
    df = _df([
        {"SKU": "1", "RP Type": "ND", "Supply Source": "Vendor", "Shop Stock Qty (CM)": 9, "Shop Stock Status": "Normal"},
    ])
    out = _add_nd_recommendation_columns(df)
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "另有安排"
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
    assert out["Recommendation of new Safety Qty / ND"].tolist() == ["ND", "另有安排"]
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


# ── Fix #7 ────────────────────────────────────────────────────────────────────

def test_fix7_combined_nd_safety_forced_to_zero():
    """ND items must have Recommend Safety Stock=0, Safety Qty Change=-adj_ss, rec col='0'."""
    df = _df([
        {"SKU": "1", "RP Type": "ND", "Shop Stock Status": "Normal",
         "Adj. Safety Stock": 10, "Shop Monthly Sold Qty (P90D)": 4,
         "Sold Qty (CM - MTD)": 2.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    assert out["Recommend Safety Stock"].tolist()[0] == 0
    assert out["Safety Qty Change"].tolist()[0] == -10
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "0"


def test_fix7_combined_p90d0_safety_forced_to_zero():
    """P90D 0 Sales items must have Recommend Safety Stock=0, rec col='ND'."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    assert out["Recommend Safety Stock"].tolist()[0] == 0
    assert out["Safety Qty Change"].tolist()[0] == -5
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND"


def test_fix7_combined_non_nd_non_p90d0_unaffected():
    """Normal/OOS rows (non-ND, non-P90D0) must retain existing calculation."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 10, "Shop Monthly Sold Qty (P90D)": 30,
         "Sold Qty (CM - MTD)": 8.0},
        {"SKU": "2", "Shop Stock Status": "OOS (0 Stock)", "RP Type": "RF",
         "Adj. Safety Stock": 2, "Shop Monthly Sold Qty (P90D)": 10,
         "Sold Qty (CM - MTD)": 0.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    # Both rows should have the normal combined calculation (positive values)
    assert out["Recommend Safety Stock"].tolist()[0] > 0
    assert out["Recommend Safety Stock"].tolist()[1] > 0
    # "Recommendation of new Safety Qty / ND" must be blank (not ND/0)
    assert out["Recommendation of new Safety Qty / ND"].tolist() == ["", ""]


# ── Fix #8: Shop Nature 倍率精確比對 ─────────────────────────────────────────
# 「Non Tourist」含 'tourist' 子字串，舊版用 str.contains 會把 Non Tourist 誤套
# Tourist 倍率；此組測試鎖定精確比對行為。

_NATURES = ["Tourist", "Non Tourist", "Mix", "Other"]


def _nature_df(natures):
    return _df([
        {"SKU": str(i), "Shop Nature": n,
         "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 999, "Shop Monthly Sold Qty (P90D)": 100,
         "Sold Qty (CM - MTD)": 10.0}
        for i, n in enumerate(natures)
    ])


def test_fix8_combined_overstock_cap_non_tourist_uses_default_rate():
    """Overstock Cap (X): Non Tourist 必須用 default(Non Tourist) 倍率，非 Tourist 倍率."""
    out = _add_combined_recommendation_columns(
        _nature_df(_NATURES), _RATES_O, _RATES_P, mtd_days=15, month_total_days=30,
        safety_mode="allow_both"
    )
    # X = floor(100 * rate): Tourist 30 / Non Tourist 50 / Mix 40 / Other 50
    assert out["Overstock Cap (X)"].tolist() == [30, 50, 40, 50]


def test_fix8_combined_oos_floor_non_tourist_uses_default_rate():
    """OOS Floor (Y): Non Tourist 必須用 default 倍率（用互異倍率區別）."""
    rates_oos = {"default": 0.2, "tourist": 0.6, "mix": 0.4}
    out = _add_combined_recommendation_columns(
        _nature_df(_NATURES), _RATES_O, rates_oos, mtd_days=15, month_total_days=30,
        safety_mode="allow_both"
    )
    # Projected（混合）= 10 + 0.5*100 = 60 ; Y = ceil(60*rate): Tourist 36 / Non Tourist 12 / Mix 24 / Other 12
    assert out["OOS Floor (Y)"].tolist() == [36, 12, 24, 12]


def test_fix8_oos_review_rate_non_tourist_uses_default_rate():
    """OOS Review 建議值同樣必須精確比對 Shop Nature."""
    rates = {"default": 0.2, "tourist": 0.6, "mix": 0.4}
    df = _df([
        {"SKU": str(i), "Shop Nature": n, "RP Type": "RF",
         "Adj. Safety Stock": 999, "Sold Qty (CM - MTD)": 10.0}
        for i, n in enumerate(_NATURES)
    ])
    out = _add_oos_recommendation_columns(
        df, rates, mtd_days=15, month_total_days=30,
        safety_mode="allow_both", oos_calc_mode="mtd"
    )
    # Projected = 20 ; rec = ceil(20*rate): Tourist 12 / Non Tourist 4 / Mix 8 / Other 4
    assert out["Recommendation of new Safety Qty"].tolist() == [12, 4, 8, 4]


def test_fix8_clean_dataframe_shop_nature_exact_filter():
    """Shop Nature 篩選必須精確比對：選 Tourist 不應帶入 Non Tourist 列."""
    df = _nature_df(_NATURES + [""])  # 加一列空白 Shop Nature

    only_tourist = _clean_dataframe(df, shop_natures=["Tourist"])
    assert only_tourist["Shop Nature"].str.lower().tolist() == ["tourist"]

    only_non_tourist = _clean_dataframe(df, shop_natures=["Non Tourist"])
    # Non Tourist + Other（非 Tourist/Mix 之值隨 Non Tourist 保留，空白除外）
    assert sorted(only_non_tourist["Shop Nature"].str.lower().tolist()) == ["non tourist", "other"]

    all_selected = _clean_dataframe(df, shop_natures=["Tourist", "Mix", "Non Tourist"])
    assert sorted(all_selected["Shop Nature"].str.lower().tolist()) == ["mix", "non tourist", "other", "tourist"]


# ── Fix #9: Projected 月銷售加權混合（方案 B）────────────────────────────────
# Projected = MTD實績 + 剩餘日數 × P90D月均（= w×MTD線性外推 + (1−w)×P90D），
# P90D 已是月均（Past 90 Days ÷ 90 × 30）直接使用；月初避免頭幾日外推失真。

def test_fix9_projected_blend_early_month_prefers_p90d():
    """月初 MTD 只有 2 日時以 P90D 為主，而非 2 日線性外推."""
    df = _df([
        {"SKU": "1", "Sold Qty (CM - MTD)": 10.0, "Shop Monthly Sold Qty (P90D)": 30},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=2, month_total_days=30, safety_mode="allow_both"
    )
    # Projected = 10 + (28/30)*30 = 38（純線性外推會是 150）
    assert out["Projected Monthly Sold Qty"].tolist()[0] == 38
    # Y = ceil(38*0.5) = 19（純線性外推會是 75）
    assert out["OOS Floor (Y)"].tolist()[0] == 19


def test_fix9_projected_blend_converges_to_mtd_at_month_end():
    """月尾（MTD 日數 = 當月日數）收斂至 MTD 實績，P90D 權重歸零."""
    df = _df([
        {"SKU": "1", "Sold Qty (CM - MTD)": 8.0, "Shop Monthly Sold Qty (P90D)": 100},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=30, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Projected Monthly Sold Qty"].tolist()[0] == 8


def test_fix9_projected_blend_nan_p90d_row_falls_back_to_linear():
    """個別列 P90D 為 NaN 時，該列退回純 MTD 線性外推，其他列正常混合."""
    df = _df([
        {"SKU": "1", "Sold Qty (CM - MTD)": 6.0, "Shop Monthly Sold Qty (P90D)": 15},
        {"SKU": "2", "Sold Qty (CM - MTD)": 6.0, "Shop Monthly Sold Qty (P90D)": None},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    # Row1: 6 + 0.5*15 = 13.5 → ceil 14 ; Row2 NaN → 純線性 6/15*30 = 12
    assert out["Projected Monthly Sold Qty"].tolist() == [14, 12]


def test_fix9_projected_missing_p90d_column_pure_linear():
    """整個 P90D 欄位缺失時維持舊行為：純 MTD 線性外推."""
    df = _df([{"SKU": "1", "Sold Qty (CM - MTD)": 6.0, "Shop Monthly Sold Qty (P90D)": 999}])
    df = df.drop(columns=["Shop Monthly Sold Qty (P90D)"])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Projected Monthly Sold Qty"].tolist()[0] == 12
