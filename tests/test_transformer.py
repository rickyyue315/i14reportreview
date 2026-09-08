"""Regression tests for transformer.py calculation logic.

Each test targets one of the six calculation issues identified in code review:
  #1 Combined Overstock Cap (X) computed inconsistently depending on dataset
  #2 Combined increase_only mode: Recommend vs Safety Qty Change mismatch
  #3 Overstock Review RP Type not normalized (case/whitespace)
  #4 OOS allow_both mode: Recommendation vs Safety Qty Increase mismatch for "other"
  #5 Carry SKU dashboard SKU count vs pivot Grand_Total mismatch
  #6 Combined mtd_days=None crashes
"""

import datetime

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
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
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
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
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


# ── Fix #10: New SKU（Launch Date 90 天內）保護 ─────────────────────────────
# 新品未出單屬上架時間問題：排除「P90D 0 Sales → Safety=0/ND」override，
# 改走正常 X/Y 計算，並於 ND 建議欄標記 New SKU (3 Month)（ND 列維持「0」）。


def _launch_date(days_ago):
    return (datetime.date.today() - datetime.timedelta(days=days_ago)).strftime("%Y/%m/%d")


def test_fix10_new_sku_p90d0_excluded_from_override_and_marked():
    """Launch Date 90 天內的 P90D 0 Sales 列不被強制 0/ND，走正常計算並標記."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0, "Launch Date": _launch_date(10)},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    # 正常計算: X=clip(floor(0*0.5),3)=3(cap Z=5) ; Y=clip(ceil(0),3)=3 → combined=3
    assert out["Recommend Safety Stock"].tolist()[0] == 3
    assert out["Safety Qty Change"].tolist()[0] == -2
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "New SKU (3 Month)"


def test_fix10_old_sku_p90d0_still_forced_zero_nd():
    """Launch Date 超過 90 天的 P90D 0 Sales 列維持原 override：0 / ND."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0, "Launch Date": _launch_date(200)},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Recommend Safety Stock"].tolist()[0] == 0
    assert out["Safety Qty Change"].tolist()[0] == -5
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND"


def test_fix10_new_sku_nd_row_keeps_zero():
    """ND 列即使為新品仍維持 Safety=0 與「0」建議（ND 規則優先於新品標記）."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "ND",
         "Adj. Safety Stock": 7, "Shop Monthly Sold Qty (P90D)": 10,
         "Sold Qty (CM - MTD)": 8.0, "Launch Date": _launch_date(10)},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Recommend Safety Stock"].tolist()[0] == 0
    assert out["Safety Qty Change"].tolist()[0] == -7
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "0"


def test_fix10_invalid_launch_date_treated_as_old_sku():
    """Launch Date 無法解析（或缺欄位）時視為非新品，維持原 override 行為."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0, "Launch Date": "not-a-date"},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Recommend Safety Stock"].tolist()[0] == 0
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND"


def test_fix10_new_sku_normal_rows_marked_but_calc_unchanged():
    """非 P90D0 的新品列同樣獲得標記，X/Y/建議值與舊邏輯一致."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 10,
         "Sold Qty (CM - MTD)": 8.0, "Launch Date": _launch_date(10)},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    # X=floor(10*0.5)=5(cap5) ; projected=8+0.5*10=13 → Y=ceil(13*0.5)=7 ; Z=5<=X → combined=max(7,5)=7
    assert out["Recommend Safety Stock"].tolist()[0] == 7
    assert out["Safety Qty Change"].tolist()[0] == 2
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "New SKU (3 Month)"


def test_fix10_ref_date_shifts_new_sku_boundary():
    """ref_date 傳入時，新品 90 天邊界以參考日期起計（預設為今天）."""
    ref = datetime.date(2026, 8, 3)
    launch = (ref - datetime.timedelta(days=85)).strftime("%Y/%m/%d")
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0, "Launch Date": launch},
    ])
    out_ref = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30,
        safety_mode="allow_both", ref_date=ref,
    )
    assert out_ref["Recommendation of new Safety Qty / ND"].tolist()[0] == "New SKU (3 Month)"
    out_default = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out_default["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND"


def test_fix10_ref_date_invalid_raises():
    """無法解析的 ref_date 直接報錯（避免靜默用錯基準）."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 10,
         "Sold Qty (CM - MTD)": 8.0, "Launch Date": _launch_date(10)},
    ])
    with pytest.raises(ValueError):
        _add_combined_recommendation_columns(
            df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30,
            safety_mode="allow_both", ref_date="not-a-date",
        )


# ── Fix #11: Recommend Safety Stock Turnday 欄位 ────────────────────────────
# 與來源「Turnday (P90D Qty)」同一準則（P90D 日均）：
# = Recommend Safety Stock ÷ (Shop Monthly Sold Qty (P90D) ÷ 30)；
# P90D 月均 = 0 留空（無周轉意義），Recommend = 0 顯示 0。


def test_fix11_recommend_turnday_calculation():
    """建議值周轉日數 = Recommend ÷ (P90D月均 ÷ 30)."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 10,
         "Sold Qty (CM - MTD)": 8.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    # Recommend=7（見 fix10 同組態）→ 7 ÷ (10/30) = 21.0
    assert out["Recommend Safety Stock"].tolist()[0] == 7
    assert out["Recommend Safety Stock Turnday"].tolist()[0] == 21.0


def test_fix11_recommend_turnday_rounding():
    """小數四捨五入至 1 位."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 99, "Shop Monthly Sold Qty (P90D)": 9,
         "Sold Qty (CM - MTD)": 0.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    # X=floor(9*0.5)=4(cap99) ; projected=0+0.5*9=4.5 → Y=ceil(2.25)=3 ; Z=99>X → combined=max(3,4)=4
    # 4 ÷ (9/30) = 13.333… → 13.3
    assert out["Recommend Safety Stock"].tolist()[0] == 4
    assert out["Safety Qty Change"].tolist()[0] == -95
    assert out["Recommend Safety Stock Turnday"].tolist()[0] == 13.3


def test_fix11_recommend_turnday_blank_when_no_p90d_sales():
    """P90D 月均 = 0 → 無周轉意義，留空（NaN）."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0, "Launch Date": _launch_date(200)},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert pd.isna(out["Recommend Safety Stock Turnday"].tolist()[0])


def test_fix11_recommend_turnday_zero_for_nd_rows():
    """ND 列 Recommend=0 → Turnday 0（零覆蓋），P90D 有銷售時."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "ND",
         "Adj. Safety Stock": 7, "Shop Monthly Sold Qty (P90D)": 10,
         "Sold Qty (CM - MTD)": 8.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Recommend Safety Stock"].tolist()[0] == 0
    assert out["Recommend Safety Stock Turnday"].tolist()[0] == 0.0


def test_fix11_recommend_turnday_column_after_recommend_safety():
    """RAW 白名單欄序：Recommend Safety Stock Turnday 緊隨 Recommend Safety Stock."""
    from transformer import COMBINED_RAW_COLS
    assert COMBINED_RAW_COLS.index("Recommend Safety Stock Turnday") == COMBINED_RAW_COLS.index("Recommend Safety Stock") + 1


# ── Fix #12: 只可增加模式 = 全表不得出現扣減 ────────────────────────────────
# increase_only 下連 ND / P90D 0 Sales 的「歸零 + ND」退場 override 也不套用，
# 一律保持原有 Safety；ND 欄僅保留 New SKU 資訊標記。退場建議只在可加可減出現。


def test_fix12_increase_only_nd_row_keeps_current_safety():
    """只可增加：ND 列保持原有 Safety（不歸零、無負數、無「0」建議）."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "ND",
         "Adj. Safety Stock": 7, "Shop Monthly Sold Qty (P90D)": 10,
         "Sold Qty (CM - MTD)": 8.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    assert out["Recommend Safety Stock"].tolist()[0] == 7
    assert out["Safety Qty Change"].tolist()[0] == 0
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == ""


def test_fix12_increase_only_p90d0_old_sku_keeps_current_safety():
    """只可增加：P90D 0 Sales 舊品保持原有 Safety（不歸零、無 ND 建議）."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0, "Launch Date": _launch_date(200)},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    assert out["Recommend Safety Stock"].tolist()[0] == 5
    assert out["Safety Qty Change"].tolist()[0] == 0
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == ""


def test_fix12_increase_only_new_sku_p90d0_row_kept_and_marked():
    """只可增加：新品 P90D 0 Sales 保持原有 Safety 並標記 New SKU."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0, "Launch Date": _launch_date(10)},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    assert out["Recommend Safety Stock"].tolist()[0] == 5
    assert out["Safety Qty Change"].tolist()[0] == 0
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "New SKU (3 Month)"


def test_fix12_increase_only_global_no_negative_change():
    """只可增加：混合資料全表 Safety Qty Change >= 0 且 Recommend >= 現時."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 10, "Sold Qty (CM - MTD)": 8.0},
        {"SKU": "2", "Shop Stock Status": "Normal", "RP Type": "ND",
         "Adj. Safety Stock": 9, "Shop Monthly Sold Qty (P90D)": 3, "Sold Qty (CM - MTD)": 0.0},
        {"SKU": "3", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 4, "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
         "Launch Date": _launch_date(200)},
        {"SKU": "4", "Shop Stock Status": "OOS (0 Stock)", "RP Type": "RF",
         "Adj. Safety Stock": 2, "Shop Monthly Sold Qty (P90D)": 6, "Sold Qty (CM - MTD)": 2.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    assert (out["Safety Qty Change"] >= 0).all()
    assert (out["Recommend Safety Stock"] >= out["Adj. Safety Stock"]).all()


# ── Fix #13: ND 建議加註 MAKE UP / Private Label + RAW 加回 Adj. MCH1 ──────
# 「ND」轉換建議（P90D 0 Sales 舊品）依 Adj. MCH1 / Adj. SST 加註：
# 02 - MAKE UP → ND(02 - MAKE UP)；03 - Private Label 系 → ND(Private Label)；
# 兩者皆是 → ND(Private Label - MAKE UP)。「0」列（已是 ND）不加註。


def _p90d0_df(mch1, sst):
    return _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0, "Launch Date": _launch_date(200),
         "Adj. MCH1": mch1, "Adj. SST": sst},
    ])


def test_fix13_nd_annotate_makeup_only():
    out = _add_combined_recommendation_columns(
        _p90d0_df("02 - MAKE UP", "01 - Local"), _RATES_O, _RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND(02 - MAKE UP)"


def test_fix13_nd_annotate_private_label_major_and_health():
    for sst in ("03 - Private Label - (1)Major", "03 - Private Label - (2)Health"):
        out = _add_combined_recommendation_columns(
            _p90d0_df("01 - SKIN CARE", sst), _RATES_O, _RATES_P,
            mtd_days=15, month_total_days=30, safety_mode="allow_both"
        )
        assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND(Private Label)"


def test_fix13_nd_annotate_both_private_label_and_makeup():
    out = _add_combined_recommendation_columns(
        _p90d0_df("02 - MAKE UP", "03 - Private Label - (1)Major"), _RATES_O, _RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND(Private Label - MAKE UP)"


def test_fix13_nd_annotate_no_match_stays_plain_nd():
    out = _add_combined_recommendation_columns(
        _p90d0_df("03 - FRAGRANCE", "05 - Para-Prem(1)Major"), _RATES_O, _RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND"


def test_fix13_nd_row_zero_not_annotated():
    """已是 ND 的列顯示「0」，即使 MAKE UP + Private Label 也不加註."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "ND",
         "Adj. Safety Stock": 7, "Shop Monthly Sold Qty (P90D)": 10,
         "Sold Qty (CM - MTD)": 8.0,
         "Adj. MCH1": "02 - MAKE UP", "Adj. SST": "03 - Private Label - (1)Major"},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "0"


def test_fix13_nd_annotate_missing_columns_falls_back():
    """缺 Adj. MCH1 / Adj. SST 欄位時維持純「ND」建議."""
    df = _p90d0_df("02 - MAKE UP", "03 - Private Label - (1)Major")
    df = df.drop(columns=["Adj. MCH1", "Adj. SST"])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Recommendation of new Safety Qty / ND"].tolist()[0] == "ND"


def test_fix13_combined_raw_whitelist_contains_adj_mch1():
    from transformer import COMBINED_RAW_COLS
    assert "Adj. MCH1" in COMBINED_RAW_COLS
    assert COMBINED_RAW_COLS.index("Adj. MCH1") == COMBINED_RAW_COLS.index("Adj. SST") + 1


# ── Fix #14: Safety Qty Change 分類欄（增加 / 減少 / 不變）──────────────────


def test_fix14_change_category_allow_both_all_three():
    """可加可減：同一批資料可出現 增加 / 減少 / 不變 三種分類."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 10, "Sold Qty (CM - MTD)": 8.0},
        {"SKU": "2", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 4, "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
         "Launch Date": _launch_date(200)},
        {"SKU": "3", "Shop Stock Status": "Normal", "RP Type": "ND",
         "Adj. Safety Stock": 0, "Shop Monthly Sold Qty (P90D)": 6, "Sold Qty (CM - MTD)": 2.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    assert out["Safety Qty Change 分類"].tolist() == ["增加", "減少", "不變"]


def test_fix14_change_category_increase_only_never_decrease():
    """只可增加：分類只會是 增加 / 不變."""
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 10, "Sold Qty (CM - MTD)": 8.0},
        {"SKU": "2", "Shop Stock Status": "Normal", "RP Type": "ND",
         "Adj. Safety Stock": 9, "Shop Monthly Sold Qty (P90D)": 3, "Sold Qty (CM - MTD)": 0.0},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    cats = out["Safety Qty Change 分類"].tolist()
    assert cats == ["增加", "不變"]
    assert "減少" not in cats


def test_fix14_change_category_column_after_change():
    """RAW 白名單欄序：Safety Qty Change 分類 緊隨 Safety Qty Change."""
    from transformer import COMBINED_RAW_COLS
    assert COMBINED_RAW_COLS.index("Safety Qty Change 分類") == COMBINED_RAW_COLS.index("Safety Qty Change") + 1


# ── Fix #15: SUMMARY sheet — Safety Qty Change 分類 / ND 建議分類 統計 ────────

from transformer import _build_combined_classification_summary  # noqa: E402


def test_fix15_change_summary_counts_and_totals():
    """可加可減：增加 / 減少 / 不變 三類行數、SKU 數、Change 合計正確；
    同一 SKU 跨分類時總計 SKU 數用全表 nunique（不虛增）."""
    raw = _add_combined_recommendation_columns(
        _df([
            {"SKU": "1", "Site Code": "S01", "Shop Stock Status": "Normal", "RP Type": "RF",
             "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 10, "Sold Qty (CM - MTD)": 8.0},
            {"SKU": "2", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
             "Adj. Safety Stock": 4, "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
             "Launch Date": _launch_date(200)},
            {"SKU": "3", "Shop Stock Status": "Normal", "RP Type": "ND",
             "Adj. Safety Stock": 0, "Shop Monthly Sold Qty (P90D)": 6, "Sold Qty (CM - MTD)": 2.0},
            {"SKU": "1", "Site Code": "S02", "Shop Stock Status": "Normal", "RP Type": "RF",
             "Adj. Safety Stock": 3, "Shop Monthly Sold Qty (P90D)": 10, "Sold Qty (CM - MTD)": 0.0},
        ]),
        _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    change_sum, _ = _build_combined_classification_summary(raw)

    def _row(label):
        return change_sum[change_sum["分類"] == label].iloc[0]

    assert _row("增加")["行數"] == 1 and _row("增加")["SKU 數"] == 1
    assert _row("增加")["Safety Qty Change 合計"] == 2
    assert _row("減少")["行數"] == 1 and _row("減少")["SKU 數"] == 1
    assert _row("減少")["Safety Qty Change 合計"] == -4
    assert _row("不變")["行數"] == 2 and _row("不變")["SKU 數"] == 2
    assert _row("不變")["Safety Qty Change 合計"] == 0

    total = change_sum[change_sum["分類"] == "總計"].iloc[0]
    assert total["行數"] == 4
    assert total["SKU 數"] == 3  # SKU "1" 跨 增加/不變 兩類，只計一次
    assert total["Safety Qty Change 合計"] == -2


def test_fix15_change_summary_increase_only_no_decrease():
    """只可增加：減少 類行數恆為 0（列仍輸出，值為 0）."""
    raw = _add_combined_recommendation_columns(
        _df([
            {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF",
             "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 10, "Sold Qty (CM - MTD)": 8.0},
            {"SKU": "2", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
             "Adj. Safety Stock": 4, "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
             "Launch Date": _launch_date(200)},
            {"SKU": "3", "Shop Stock Status": "Normal", "RP Type": "ND",
             "Adj. Safety Stock": 9, "Shop Monthly Sold Qty (P90D)": 3, "Sold Qty (CM - MTD)": 0.0},
        ]),
        _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="increase_only"
    )
    change_sum, _ = _build_combined_classification_summary(raw)
    dec = change_sum[change_sum["分類"] == "減少"].iloc[0]
    assert dec["行數"] == 0 and dec["SKU 數"] == 0 and dec["Safety Qty Change 合計"] == 0
    total = change_sum[change_sum["分類"] == "總計"].iloc[0]
    assert total["行數"] == 3
    assert total["Safety Qty Change 合計"] >= 0


def test_fix15_nd_summary_all_categories():
    """可加可減：ND 建議分類涵蓋 空白 / New SKU / 0 / ND / 三種加註，順序固定."""
    raw = _add_combined_recommendation_columns(
        _df([
            {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF"},
            {"SKU": "2", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
             "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
             "Launch Date": _launch_date(10)},
            {"SKU": "3", "Shop Stock Status": "Normal", "RP Type": "ND"},
            {"SKU": "4", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
             "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
             "Launch Date": _launch_date(200),
             "Adj. MCH1": "02 - MAKE UP", "Adj. SST": "01 - Local"},
            {"SKU": "5", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
             "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
             "Launch Date": _launch_date(200),
             "Adj. MCH1": "01 - SKIN CARE", "Adj. SST": "03 - Private Label - (1)Major"},
            {"SKU": "6", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
             "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
             "Launch Date": _launch_date(200),
             "Adj. MCH1": "02 - MAKE UP", "Adj. SST": "03 - Private Label - (2)Health"},
            {"SKU": "7", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
             "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
             "Launch Date": _launch_date(200)},
        ]),
        _RATES_O, _RATES_P, mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    _, nd_sum = _build_combined_classification_summary(raw)

    assert nd_sum["建議分類"].tolist() == [
        "(空白) 維持 Safety Stock",
        "New SKU (3 Month)",
        "0（已是 ND）",
        "ND",
        "ND(02 - MAKE UP)",
        "ND(Private Label)",
        "ND(Private Label - MAKE UP)",
        "總計",
    ]
    counts = dict(zip(nd_sum["建議分類"], nd_sum["行數"]))
    assert all(counts[k] == 1 for k in counts if k != "總計")
    assert counts["總計"] == 7
    total = nd_sum[nd_sum["建議分類"] == "總計"].iloc[0]
    assert total["SKU 數"] == 7


def test_fix15_excel_contains_summary_sheet():
    """輸出 Excel 含 SUMMARY sheet（位於 By Shop 之後），兩個區塊標題與表頭齊全."""
    from io import BytesIO

    import openpyxl
    from transformer import transform_combined, generate_excel_combined

    df = _df([
        {"SKU": "1", "Shop Stock Status": "Normal", "RP Type": "RF"},
        {"SKU": "2", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
         "Launch Date": _launch_date(200)},
        {"SKU": "3", "Shop Stock Status": "Normal", "RP Type": "ND"},
    ])
    s, r, c, _ = transform_combined(
        df, rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both"
    )
    buf = generate_excel_combined(s, r, c)

    wb = openpyxl.load_workbook(BytesIO(buf.getvalue()))
    assert wb.sheetnames == ["RAW", "Shop Stock Status", "By Shop", "店舖問題 SKU", "SUMMARY", "Safety Stock 數據分析摘要"]

    ws = wb["SUMMARY"]
    col_a = [ws.cell(row=i, column=1).value for i in range(1, ws.max_row + 1)]
    assert "① Safety Qty Change 分類（增加 / 減少 / 不變）" in col_a
    assert "② Recommendation of new Safety Qty / ND 分類" in col_a

    all_values = [ws.cell(row=i, column=j).value
                  for i in range(1, ws.max_row + 1) for j in range(1, ws.max_column + 1)]
    for expected in ("行數", "SKU 數", "Safety Qty Change 合計", "增加", "減少", "不變",
                     "ND", "0（已是 ND）", "(空白) 維持 Safety Stock", "總計"):
        assert expected in all_values
