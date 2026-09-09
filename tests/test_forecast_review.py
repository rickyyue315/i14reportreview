"""Validation tests for Forecast Review mode (v1.38.0).

SKU 級建議訂貨量（第二維度防缺貨）：三情境訊號、爆升四類、缺貨還原取代式、
Warehouse／Vendor 位置分流、Pending 別名、缺欄守衛、缺口與上限封頂、
全 ND 零訂量、店舖覆蓋、Excel 六表。
"""
import warnings
warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
import pytest
from openpyxl import load_workbook
from transformer import (
    transform_forecast,
    generate_excel_forecast,
    _FORECAST_MAIN_COLUMNS,
    _FORECAST_COVERAGE_COLUMNS,
    _FORECAST_BUYER_COLUMNS,
    _FC_CLASS_NEW, _FC_CLASS_ACCEL, _FC_CLASS_SURGE, _FC_CLASS_STABLE,
    _FC_STATUS_ORDER, _FC_STATUS_ORDER_CAPPED, _FC_STATUS_ENOUGH, _FC_STATUS_ND,
    _FC_TARGET_D001, _FC_TARGET_SHOP,
    _FC_RISK_OOS, _FC_RISK_7D, _FC_RISK_TARGET, _FC_RISK_OK,
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1, "SKU Desc": "Test Product",
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Mix", "Vendor MC": "V1", "Item Status": "Active",
    "Launch Date": "2024-01-01", "Item Class": "C1", "RP Type": "RF",
    "Supply Source": "Warehouse", "Shop Stock Status": "Normal",
    "Adj. Safety Stock": 3.0,
    "Shop Monthly Sold Qty (P90D)": 60.0,
    "Sold Qty (P90D)": 1800.0,
    "Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 55.0, "Sold Qty (CM - MTD)": 16.0,
    "Shop Stock Qty (CM)": 20.0,
    "D001 Stock Qty (CM)": 100.0, "D001 Net Stock Qty (CM)": 80.0,
    "D001 Blocked Qty (CM)": 0.0, "D001 QI Qty (CM)": 0.0,
    "D001 On Order PO Qty (CM)": 10.0,
    "Shop On Order PO Qty (CM)": 0.0,
    "Shop Pending Receive (±7 days)": 5.0,
}

MTD_DAYS = 8
MONTH_TOTAL = 30


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def _run(df, **kw):
    params = dict(mtd_days=MTD_DAYS, month_total_days=MONTH_TOTAL)
    params.update(kw)
    return transform_forecast(df, **params)


def _main_row(main, sku="1"):
    hit = main[main["SKU"] == sku]
    assert len(hit) == 1
    return hit.iloc[0]


# ── 1. 輸入檢核 ─────────────────────────────────────────────────────────────
def test_missing_required_column_stops():
    df = _df([{}]).drop(columns=["Sold Qty (M-1)"])
    with pytest.raises(ValueError, match="缺少必需欄位"):
        _run(df)


def test_missing_common_column_stops():
    df = _df([{}]).drop(columns=["Buyer"])
    with pytest.raises(ValueError, match="缺少必要欄位"):
        _run(df)


# ── 2. 逐列訊號 ─────────────────────────────────────────────────────────────
def test_row_signals():
    # M-2=50, M-1=55, MTD=16 (8日, 月化60), p90d月均=60
    # 保守=60; 標準=max(60,55)=60; 樂觀=max(60*0.8, 60, 55)=60
    df = _df([{}])
    main, _, _, raw, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 保守情境"] == 60
    assert row["FC 標準情境"] == 60
    assert row["FC 樂觀情境"] == 60
    assert raw["FC MTD 月化"].iloc[0] == pytest.approx(60.0)
    assert raw["FC 保守情境"].iloc[0] == pytest.approx(60.0)
    assert raw["FC 標準情境"].iloc[0] == pytest.approx(60.0)
    assert raw["FC 樂觀情境"].iloc[0] == pytest.approx(60.0)


def test_optimistic_applies_damp_then_max():
    # MTD=200 (8日, 月化750), p90d=60, M-1=55 → 樂觀=max(750*0.8, 60, 55)=600
    df = _df([{"Sold Qty (CM - MTD)": 200.0}])
    main, _, _, raw, _, _ = _run(df, damp=0.8)
    assert _main_row(main)["FC 樂觀情境"] == 600
    assert raw["FC 樂觀情境"].iloc[0] == pytest.approx(600.0)


def test_optimistic_never_below_standard():
    # MTD=0 → 月化0; 樂觀=max(0, 60, 55)=60=標準
    df = _df([{"Sold Qty (CM - MTD)": 0.0}])
    main, _, _, _, _, _ = _run(df, damp=0.5)
    row = _main_row(main)
    assert row["FC 樂觀情境"] >= row["FC 標準情境"]


def test_p90d_monthly_fallback():
    df = _df([{}]).drop(columns=["Shop Monthly Sold Qty (P90D)"])
    main, _, _, _, _, warns = _run(df)
    # Sold Qty (P90D)=1800 → 1800/90*30.5=610
    assert _main_row(main)["FC 保守情境"] == 610
    assert any("Shop Monthly Sold Qty (P90D)" in w for w in warns)


def test_invalid_mtd_falls_back():
    df = _df([{"Sold Qty (CM - MTD)": 200.0}])
    main, _, _, _, _, warns = _run(df, mtd_days=0, month_total_days=30)
    assert _main_row(main)["FC 樂觀情境"] == 60  # max(M-1=55, p90d=60)
    assert any("MTD" in w for w in warns)


def test_negative_sales_clamped_with_warning():
    df = _df([{"Sold Qty (M-1)": -5.0}])
    main, _, _, _, _, warns = _run(df)
    assert _main_row(main)["FC 標準情境"] == 60  # M-1 按 0 → max(60, 0)
    assert any("負數" in w for w in warns)


def test_negative_stock_clamped_with_warning():
    df = _df([{"Shop Stock Qty (CM)": -3.0}])
    main, _, _, _, _, warns = _run(df)
    assert _main_row(main)["FC 店舖庫存合計"] == 0
    assert any("Shop Stock Qty (CM)" in w for w in warns)


# ── 3. 爆升四類 ─────────────────────────────────────────────────────────────
def test_class_new_volume():
    # M-2=0, M-1=0, MTD=100 → 新起量（用樂觀）
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 100.0, "Shop Monthly Sold Qty (P90D)": 10.0}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_NEW
    assert row["FC 月預估 (採用)"] == row["FC 樂觀情境"]


def test_class_accelerating():
    # M-2=100, M-1=120 (≥1.15×100), MTD月化=140 (≥1.15×120) → 加速中
    df = _df([{"Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 120.0,
               "Sold Qty (CM - MTD)": 37.0,  # 37/8*30=138.75
               "Shop Monthly Sold Qty (P90D)": 100.0}])
    main, _, _, _, _, _ = _run(df)
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_ACCEL


def test_class_surge():
    # M-2=50, M-1=100, MTD月化=60 → max(100,60)/50=2.0 ≥ 1.5 → 爆升
    df = _df([{"Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 100.0,
               "Sold Qty (CM - MTD)": 16.0,
               "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, _, _, _ = _run(df)
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_SURGE


def test_surge_boundary_exact():
    # max(75,60)/50=1.5 恰達門檻 → 爆升；74 則平穩
    kw = {"Sold Qty (M-2)": 50.0, "Sold Qty (CM - MTD)": 16.0,
          "Shop Monthly Sold Qty (P90D)": 60.0}
    m1 = _run(_df([dict(kw, **{"Sold Qty (M-1)": 75.0})]))[0]
    m2 = _run(_df([dict(kw, **{"Sold Qty (M-1)": 74.0})]))[0]
    assert _main_row(m1)["FC 爆升分類"] == _FC_CLASS_SURGE
    assert _main_row(m2)["FC 爆升分類"] == _FC_CLASS_STABLE


def test_class_stable_uses_standard():
    df = _df([{}])  # M-2=50, M-1=55, 月化60 → 平穩
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_STABLE
    assert row["FC 月預估 (採用)"] == row["FC 標準情境"]


def test_custom_surge_params():
    df = _df([{"Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 60.0,
               "Sold Qty (CM - MTD)": 16.0, "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, _, _, _ = _run(df, surge_mult=1.1)
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_SURGE


# ── 4. 缺貨還原（取代式） ────────────────────────────────────────────────────
def test_oos_replacement_picks_mtd_when_higher():
    # 缺貨列：庫存 0, MTD=76（月化285→樂觀228）, p90d月均=60, M-1=55
    # M-2=50, M-1=55 → 平穩（max(55,285)/50 雖 ≥1.5……實為爆升，見下）
    # 此例 MTD 月化 285 令 SKU 判爆升 → 用樂觀還原：max(228, 76)=228
    df = _df([{"Shop Stock Qty (CM)": 0.0, "Sold Qty (CM - MTD)": 76.0,
               "Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 55.0,
               "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, raw, _, _ = _run(df)
    assert bool(raw["FC 缺貨還原列"].iloc[0]) is True
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_SURGE
    assert _main_row(main)["FC 月預估 (採用)"] == 228


def test_oos_std_replacement_picks_mtd_when_higher():
    # 平穩 SKU 的缺貨列：標準情境訊號 60 vs MTD 實績 76 → 取 76
    # M-2=200, M-1=55, MTD=76（月化285）→ max(55,285)/200=1.425 <1.5 → 平穩
    df = _df([{"Shop Stock Qty (CM)": 0.0, "Sold Qty (CM - MTD)": 76.0,
               "Sold Qty (M-2)": 200.0, "Sold Qty (M-1)": 55.0,
               "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, raw, _, _ = _run(df)
    assert bool(raw["FC 缺貨還原列"].iloc[0]) is True
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_STABLE
    assert _main_row(main)["FC 月預估 (採用)"] == 76


def test_oos_mtd_lower_than_signal_uses_signal():
    # 缺貨列：MTD=10, p90d=60 → max(60, 10)=60
    df = _df([{"Shop Stock Qty (CM)": 0.0, "Sold Qty (CM - MTD)": 10.0}])
    main, _, _, _, _, _ = _run(df)
    assert _main_row(main)["FC 月預估 (採用)"] == 60


def test_non_oos_row_uses_signal_only():
    # 非缺貨列：訊號不受 MTD 實績影響。M-2=200, M-1=55, MTD=76 → 平穩，用標準 60
    df = _df([{"Shop Stock Qty (CM)": 5.0, "Sold Qty (CM - MTD)": 76.0,
               "Sold Qty (M-2)": 200.0}])
    main, _, _, raw, _, _ = _run(df)
    assert bool(raw["FC 缺貨還原列"].iloc[0]) is False
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_STABLE
    assert _main_row(main)["FC 月預估 (採用)"] == 60


def test_oos_nd_row_excluded_from_replacement():
    # 缺貨但 ND：不還原，取情境訊號；且全 ND → 訂量 0（爆升仍判類，但不用於訂量）
    df = _df([{"Shop Stock Qty (CM)": 0.0, "Sold Qty (CM - MTD)": 76.0,
               "RP Type": "ND"}])
    main, _, _, raw, _, _ = _run(df)
    assert bool(raw["FC 缺貨還原列"].iloc[0]) is False
    assert _main_row(main)["FC 月預估 (採用)"] == _main_row(main)["FC 樂觀情境"]
    assert _main_row(main)["FC 建議訂量"] == 0
    assert _main_row(main)["FC SKU 狀態"] == _FC_STATUS_ND


# ── 5. 位置分流 ─────────────────────────────────────────────────────────────
def test_warehouse_position():
    # 店存20 ＋ D001 Net 80 ＋ D001 PO 10 ＋ 在途5 ＝ 115
    df = _df([{}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 位置合計"] == 115
    assert row["FC 建議訂貨對象"] == _FC_TARGET_D001


def test_vendor_position():
    # 店存20 ＋ Shop PO 7 ＋ 在途5 ＝ 32（不計 D001）
    df = _df([{"Supply Source": "Vendor",
               "Shop On Order PO Qty (CM)": 7.0,
               "D001 Net Stock Qty (CM)": 80.0,
               "D001 On Order PO Qty (CM)": 10.0}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 位置合計"] == 32
    assert row["FC 建議訂貨對象"] == _FC_TARGET_SHOP


def test_pending_alias_old_name():
    df = _df([{}]).rename(columns={"Shop Pending Receive (±7 days)": "Shop Pending Receive"})
    main, _, _, _, stats, _ = _run(df)
    assert stats["pending_col"] == "Shop Pending Receive"
    assert _main_row(main)["FC 位置合計"] == 115


def test_pending_missing_warns_zero():
    df = _df([{}]).drop(columns=["Shop Pending Receive (±7 days)"])
    main, _, _, _, stats, warns = _run(df)
    assert stats["pending_col"] is None
    assert _main_row(main)["FC 位置合計"] == 110  # 20+80+10
    assert any("Pending Receive" in w for w in warns)


def test_d001_po_missing_warns():
    df = _df([{}]).drop(columns=["D001 On Order PO Qty (CM)"])
    main, _, _, _, _, warns = _run(df)
    assert _main_row(main)["FC 位置合計"] == 105  # 20+80+5
    assert any("D001 On Order PO" in w for w in warns)


def test_d001_first_value_not_summed():
    # 同 SKU 各列 D001 值相同（實務常態）：取首值，不加總
    df = _df([
        {"Site Code": "S01", "Shop Stock Qty (CM)": 20.0},
        {"Site Code": "S02", "Shop Stock Qty (CM)": 30.0,
         "Item_Shop": "Prod-S02-Shop2"},
    ])
    main, _, _, _, stats, _ = _run(df)
    row = _main_row(main)
    assert row["FC D001 庫存 (Net)"] == 80
    assert row["FC D001 PO"] == 10
    assert row["FC 位置合計"] == 50 + 80 + 10 + 10  # 店存50＋Net80＋PO10＋在途10


def test_d001_inconsistent_warns():
    df = _df([
        {"Site Code": "S01"},
        {"Site Code": "S02", "D001 Net Stock Qty (CM)": 5.0,
         "Item_Shop": "Prod-S02-Shop2"},
    ])
    main, _, _, _, _, warns = _run(df)
    assert _main_row(main)["FC D001 庫存 (Net)"] == 80  # 取首個非空值
    assert any("D001" in w and "不一致" in w for w in warns)


def test_supply_missing_defaults_warehouse():
    df = _df([{}]).drop(columns=["Supply Source"])
    main, _, _, _, _, warns = _run(df)
    assert _main_row(main)["FC 建議訂貨對象"] == _FC_TARGET_D001
    assert any("Supply Source" in w for w in warns)


def test_rp_missing_defaults_rf():
    df = _df([{}]).drop(columns=["RP Type"])
    main, _, _, _, _, warns = _run(df)
    assert _main_row(main)["FC RP 概況"] == "全 RF"
    assert any("RP Type" in w for w in warns)


def test_multi_supply_mode_warning():
    df = _df([{"Site Code": "S01", "Supply Source": "Warehouse"},
              {"SKU": 1, "Site Code": "S02", "Supply Source": "Vendor",
               "Item_Shop": "Prod-S02-Shop2"}])
    main, _, _, _, stats, warns = _run(df)
    assert stats["multi_supply_skus"] == 1
    assert any("Supply Source" in w for w in warns)


# ── 6. 缺口與上限 ───────────────────────────────────────────────────────────
def test_gap_and_cap():
    # 平穩 SKU: 月預估=60; 目標=ceil(60/30.5*30)=60; 位置=115 → 缺口0 → 倉位足夠, 訂量0
    df = _df([{}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 目標覆蓋量"] == 60
    assert row["FC 缺口"] == 0
    assert row["FC 建議訂量"] == 0
    assert row["FC SKU 狀態"] == _FC_STATUS_ENOUGH


def test_order_equals_gap_when_within_cap():
    # 新起量: 月化285 → 樂觀228; 位置=115; 目標=ceil(228/30.5*30)=225; 缺口=110; 上限=ceil(228*2)-115=341 → 訂110
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
               "Shop Stock Qty (CM)": 20.0}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_NEW
    assert row["FC 月預估 (採用)"] == 228
    assert row["FC 缺口"] == 110
    assert row["FC 建議訂量"] == 110
    assert row["FC SKU 狀態"] == _FC_STATUS_ORDER


def test_cap_binds():
    # cap_months=0.5: 上限=ceil(228*0.5)-115=114-115<0 → 0 → 訂0, 狀態已封頂（缺口110>上限0）
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
               "Shop Stock Qty (CM)": 20.0}])
    main, _, _, _, _, _ = _run(df, cap_months=0.5)
    row = _main_row(main)
    assert row["FC 上限空間"] == 0
    assert row["FC 建議訂量"] == 0
    assert row["FC SKU 狀態"] == _FC_STATUS_ORDER_CAPPED


# ── 7. 店舖覆蓋 ─────────────────────────────────────────────────────────────
def test_coverage_rows_and_risk_flags():
    # 建議訂量>0 的 SKU 才有覆蓋列；S02 缺貨(庫存0)→已缺貨；S01 店存20+在途5, 日銷=MTD76/8=9.5 → 2.6日 → ≤7日
    df = _df([
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "Shop Stock Qty (CM)": 20.0, "Site Code": "S01"},
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "Shop Stock Qty (CM)": 0.0, "Site Code": "S02",
         "Shop Pending Receive (±7 days)": 12.0,
         "Item_Shop": "Prod-S02-Shop2"},
    ])
    main, cov, _, _, stats, _ = _run(df)
    assert stats["coverage_rows"] == 2
    by_site = {r["Site Code"]: r for _, r in cov.iterrows()}
    assert by_site["S02"]["FC 風險旗"] == _FC_RISK_OOS
    assert by_site["S01"]["FC 風險旗"] == _FC_RISK_7D
    # 排序：已缺貨在前
    assert list(cov["Site Code"]) == ["S02", "S01"]
    # 主表風險店數=2（兩店可用日數皆 ≤7）
    assert _main_row(main)["FC 缺貨風險店數"] == 2
    assert _main_row(main)["FC 現已缺貨店數"] == 1
    assert _main_row(main)["FC 鋪貨店數"] == 2


def test_coverage_target_flag():
    # 目標30日：可用日數 10日 → ≤目標天數
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
               "Shop Stock Qty (CM)": 83.0,  # (83+12)/9.5=10.0日
               "Shop Pending Receive (±7 days)": 12.0}])
    _, cov, _, _, _, _ = _run(df)
    assert cov["FC 風險旗"].iloc[0] == _FC_RISK_TARGET


def test_coverage_daily_fallback_p90d():
    # 覆蓋表日銷退回公式：MTD=0 的列 → 日銷＝P90D÷30.5
    # 構造訂量>0：A 店缺貨高 MTD（拉高月預估致缺口），B 店 MTD=0、P90D 高、有小量庫存
    df = _df([
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "Sold Qty (P90D)": 76.0, "Shop Stock Qty (CM)": 0.0,
         "Shop Pending Receive (±7 days)": 0.0, "Site Code": "S01"},
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 0.0, "Shop Monthly Sold Qty (P90D)": 0.0,
         "Sold Qty (P90D)": 1800.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Pending Receive (±7 days)": 0.0, "Site Code": "S02",
         "Item_Shop": "Prod-S02-Shop2"},
    ])
    main, cov, _, _, _, _ = _run(df)
    assert _main_row(main)["FC 建議訂量"] > 0
    by_site = {r["Site Code"]: r for _, r in cov.iterrows()}
    # 覆蓋表店舖日銷取 2 位小數
    assert by_site["S02"]["FC 店舖日銷"] == pytest.approx(round(1800.0 / 30.5, 2), rel=1e-6)
    # (5+0)/59.02 ≈ 0.08 日 → ≤7日
    assert by_site["S02"]["FC 風險旗"] == _FC_RISK_7D


def test_coverage_empty_when_no_order():
    df = _df([{}])  # 倉位足夠 → 訂量0 → 覆蓋空表
    _, cov, _, _, stats, _ = _run(df)
    assert cov.empty
    assert list(cov.columns) == _FORECAST_COVERAGE_COLUMNS
    assert stats["coverage_rows"] == 0


def test_coverage_ok_flag():
    # S01 缺貨高 MTD（拉高月預估致缺口>0）；S02 零 MTD 但 P90D 有量 → 日銷>0；
    # S02 高庫存 → 可用日數大 → 足夠（S02 店存須計入位置，故月預估須夠大蓋過位置）
    df = _df([
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 760.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "Sold Qty (P90D)": 760.0, "Shop Stock Qty (CM)": 0.0,
         "Shop Pending Receive (±7 days)": 0.0, "Site Code": "S01",
         "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0,
         "D001 Stock Qty (CM)": 0.0},
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 0.0, "Shop Monthly Sold Qty (P90D)": 0.0,
         "Sold Qty (P90D)": 300.0,
         "Shop Stock Qty (CM)": 500.0, "Site Code": "S02",
         "Shop Pending Receive (±7 days)": 0.0,
         "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0,
         "D001 Stock Qty (CM)": 0.0,
         "Item_Shop": "Prod-S02-Shop2"},
    ])
    main, cov, _, _, _, _ = _run(df)
    assert _main_row(main)["FC 建議訂量"] > 0
    by_site = {r["Site Code"]: r for _, r in cov.iterrows()}
    # S02 日銷=300/30.5≈9.84, 可用=(5000)/9.84≈508日 → 足夠
    assert by_site["S02"]["FC 風險旗"] == _FC_RISK_OK


# ── 8. Buyer 彙總與 stats ───────────────────────────────────────────────────
def test_buyer_summary():
    df = _df([
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "Shop Stock Qty (CM)": 20.0},
        {"SKU": 2, "Buyer": "B2", "Item_Shop": "P2-S01",
         "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "Shop Stock Qty (CM)": 20.0, "Supply Source": "Vendor",
         "Shop On Order PO Qty (CM)": 3.0},
    ])
    main, _, buyer, _, stats, _ = _run(df)
    assert list(buyer.columns) == _FORECAST_BUYER_COLUMNS
    b1 = buyer[buyer["Buyer"] == "B1"].iloc[0]
    b2 = buyer[buyer["Buyer"] == "B2"].iloc[0]
    assert b1["FC 建議訂貨 SKU 數"] == 1
    assert b1["FC 建議訂量合計 (D001 PO)"] == _main_row(main, "1")["FC 建議訂量"]
    assert b2["FC 建議訂量合計 (Shop PO)"] == _main_row(main, "2")["FC 建議訂量"]
    assert stats["total_order"] == stats["total_order_d001"] + stats["total_order_shop"]
    assert stats["total_skus"] == 2


def test_main_sort_and_columns():
    df = _df([
        {"SKU": 2, "Buyer": "B1", "Brand": "B", "Item_Shop": "P2-S01"},
        {"SKU": 1, "Buyer": "B1", "Brand": "A"},
    ])
    main, _, _, _, _, _ = _run(df)
    assert list(main.columns) == _FORECAST_MAIN_COLUMNS
    assert list(main["SKU"]) == ["1", "2"]


def test_stats_params_echo():
    _, _, _, _, stats, _ = _run(_df([{}]), target_days=45, cap_months=3.0,
                                surge_mult=2.0, surge_min_base=5, accel_mult=1.2, damp=0.9)
    assert stats["target_days"] == 45
    assert stats["cap_months"] == 3.0
    assert stats["surge_mult"] == 2.0
    assert stats["surge_min_base"] == 5
    assert stats["accel_mult"] == 1.2
    assert stats["damp"] == 0.9
    assert stats["mtd_days"] == MTD_DAYS


def test_param_clamp_warns():
    _, _, _, _, stats, warns = _run(_df([{}]), target_days=999, damp="bad")
    assert stats["target_days"] == 90
    assert stats["damp"] == 0.8
    assert len([w for w in warns if "目標覆蓋天數" in w]) == 1
    assert len([w for w in warns if "衰減係數" in w]) == 1


# ── 9. Excel 輸出 ───────────────────────────────────────────────────────────
def test_excel_sheets_and_guide_first():
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
               "Shop Stock Qty (CM)": 20.0}])
    main, cov, buyer, raw, stats, _ = _run(df)
    buf = generate_excel_forecast(main, cov, buyer, raw, stats)
    wb = load_workbook(buf)
    assert wb.sheetnames == ["說明", "Forecast 主表", "店舖覆蓋", "按 Buyer 彙總", "RAW", "Forecast 摘要"]
    assert wb.active.title == "說明"
    ws = wb["Forecast 主表"]
    headers = [c.value for c in ws[2]]
    assert headers == _FORECAST_MAIN_COLUMNS
    # 建議訂量欄有醒目底色
    order_idx = _FORECAST_MAIN_COLUMNS.index("FC 建議訂量") + 1
    assert ws.cell(row=3, column=order_idx).fill.start_color.rgb != "00000000"
    ws_cov = wb["店舖覆蓋"]
    assert [c.value for c in ws_cov[2]] == [c for c in _FORECAST_COVERAGE_COLUMNS if c in cov.columns]


def test_excel_empty_main():
    df = _df([{}])
    main, cov, buyer, raw, stats, _ = _run(df)
    buf = generate_excel_forecast(main.iloc[0:0], cov.iloc[0:0], buyer.iloc[0:0], raw.iloc[0:0], stats)
    wb = load_workbook(buf)
    assert wb.sheetnames == ["說明", "Forecast 主表", "店舖覆蓋", "按 Buyer 彙總", "RAW", "Forecast 摘要"]
    assert wb["Forecast 主表"].cell(row=2, column=1).value == "無符合資料"
