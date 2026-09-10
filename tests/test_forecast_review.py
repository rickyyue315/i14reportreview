"""Validation tests for Forecast Review mode (v1.39.0, v2 demand model).

SKU 級建議訂貨量（v2 加權混合需求模型）：缺貨還原（M-2／M-1／CM）、開賣日折算、
MTD 可信度加權、加權標準情境、趨勢外推樂觀情境、SKU 聚合（非逐店 max 相加）、
爆升四類、位置分流、Pending 別名、缺欄守衛、缺口與上限封頂、全 ND 零訂量、
店舖覆蓋、Excel 六表。
"""
import warnings
warnings.filterwarnings("ignore")
from datetime import date
import numpy as np
import pandas as pd
import pytest
from openpyxl import load_workbook
from transformer import (
    transform_forecast,
    generate_excel_forecast,
    _forecast_raw_display_columns,
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
MTD_DATE = date(2026, 9, 8)  # M-1＝2026-08（31 日）、M-2＝2026-07（31 日）


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


# ── 2. v2 情境（加權混合／MTD 可信度／趨勢外推） ─────────────────────────────
def test_scenarios_weighted_blend():
    # B=60, M-1=55, M-2=50, MTD=16（8日→月化60）
    # MTD 調整月化=8/30×60+22/30×60=60；標準=0.4×60+0.3×55+0.3×60=58.5→58
    # g=55/50−1=0.1；樂觀=max(60×1.08, 58.5)=64.8→64
    df = _df([{}])
    main, _, _, raw, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 保守情境"] == 60
    assert row["FC 標準情境"] == 58
    assert row["FC 樂觀情境"] == 64
    assert row["FC MTD 調整月化"] == pytest.approx(60.0)
    assert row["FC 趨勢增速"] == pytest.approx(0.1)
    assert raw["FC MTD 月化"].iloc[0] == pytest.approx(60.0)
    # 三情境欄在 RAW 為 SKU 級數值廣播
    assert raw["FC 保守情境"].iloc[0] == 60
    assert raw["FC 標準情境"].iloc[0] == 58
    assert raw["FC 樂觀情境"].iloc[0] == 64


def test_optimistic_trend_extrapolation():
    # MTD=200（月化750）→ MTD 調整月化=8/30×750+22/30×60=244
    # 標準=0.4×60+0.3×55+0.3×244=113.7→113；爆升 → 樂觀=max(244×1.08, 113.7)=263.52→263
    df = _df([{"Sold Qty (CM - MTD)": 200.0}])
    main, _, _, _, _, _ = _run(df, damp=0.8)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_SURGE
    assert row["FC 標準情境"] == 113
    assert row["FC 樂觀情境"] == 263
    assert row["FC 月預估 (採用)"] == 263


def test_optimistic_never_below_standard():
    df = _df([{"Sold Qty (CM - MTD)": 0.0}])
    main, _, _, _, _, _ = _run(df, damp=0.5)
    row = _main_row(main)
    assert row["FC 樂觀情境"] >= row["FC 標準情境"]


def test_p90d_monthly_fallback():
    df = _df([{}]).drop(columns=["Shop Monthly Sold Qty (P90D)"])
    main, _, _, _, _, warns = _run(df)
    # Sold Qty (P90D)=1800 → 1800/90×30.5=610；保守=610
    assert _main_row(main)["FC 保守情境"] == 610
    assert any("Shop Monthly Sold Qty (P90D)" in w for w in warns)


def test_invalid_mtd_falls_back():
    df = _df([{"Sold Qty (CM - MTD)": 200.0}])
    main, _, _, raw, _, warns = _run(df, mtd_days=0, month_total_days=30)
    row = _main_row(main)
    # MTD 無效：MTD 月化退回 max(M-1=55, B=60)=60
    assert row["FC MTD 調整月化"] == pytest.approx(60.0)
    assert raw["FC MTD 月化"].iloc[0] == pytest.approx(0.0)
    assert row["FC 月預估 (採用)"] == row["FC 標準情境"]
    assert any("MTD" in w for w in warns)


def test_mtd_credibility_increases_with_elapsed_days():
    # 同一日均銷（10 件／日）：已過日數越多，MTD 調整月化越貼近 MTD 月化 300
    outs = []
    for days in (1, 8, 15, 30):
        main, _, _, _, _, _ = _run(
            _df([{"Sold Qty (CM - MTD)": 10.0 * days}]), mtd_days=days
        )
        outs.append(_main_row(main)["FC MTD 調整月化"])
    assert outs == pytest.approx([68.0, 124.0, 180.0, 300.0], abs=0.2)
    assert outs == sorted(outs)


def test_negative_sales_clamped_with_warning():
    df = _df([{"Sold Qty (M-1)": -5.0}])
    main, _, _, _, _, warns = _run(df)
    # M-1 按 0 → 標準=0.4×60+0.3×0+0.3×60=42
    assert _main_row(main)["FC 標準情境"] == 42
    assert any("負數" in w for w in warns)


def test_negative_stock_clamped_with_warning():
    df = _df([{"Shop Stock Qty (CM)": -3.0}])
    main, _, _, _, _, warns = _run(df)
    assert _main_row(main)["FC 店舖庫存合計"] == 0
    assert any("Shop Stock Qty (CM)" in w for w in warns)


# ── 3. SKU 聚合與爆升四類 ────────────────────────────────────────────────────
def test_sku_aggregate_not_row_max():
    # 兩店趨勢相反（v1 逐店 max 相加會得 100+100=200；v2 先聚合：標準＝138）
    df = _df([
        {"Site Code": "S01", "Shop Monthly Sold Qty (P90D)": 100.0,
         "Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 50.0, "Sold Qty (CM - MTD)": 0.0},
        {"Site Code": "S02", "Item_Shop": "Prod-S02-Shop2",
         "Shop Monthly Sold Qty (P90D)": 50.0,
         "Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 100.0, "Sold Qty (CM - MTD)": 0.0},
    ])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    # B=150, S1=150, S2=100；MTD 月化=0 → 調整=22/30×150=110
    # 標準=0.4×150+0.3×150+0.3×110=138；g=150/100−1=0.5
    assert row["FC 標準情境"] == 138
    assert row["FC 標準情境"] != 200
    assert row["FC M-2 合計"] == pytest.approx(100.0)
    assert row["FC M-1 合計"] == pytest.approx(150.0)
    assert row["FC 爆升分類"] == _FC_CLASS_SURGE
    assert row["FC 趨勢增速"] == pytest.approx(0.5)
    assert row["FC 月預估 (採用)"] == row["FC 樂觀情境"]


def test_trend_growth_zero_when_declining():
    df = _df([{"Sold Qty (M-2)": 200.0, "Sold Qty (M-1)": 100.0,
               "Sold Qty (CM - MTD)": 16.0}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_STABLE
    assert row["FC 趨勢增速"] == pytest.approx(0.0)
    assert row["FC 月預估 (採用)"] == row["FC 標準情境"]


def test_trend_growth_cap_applied():
    df = _df([{"Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 100.0,
               "Sold Qty (CM - MTD)": 16.0, "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, _, stats, _ = _run(df, trend_cap=0.2)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_SURGE
    assert row["FC 趨勢增速"] == pytest.approx(0.2)
    # MTD 調整=60；標準=72；樂觀=max(max(100,60)×1.16, 72)=116
    assert row["FC 樂觀情境"] == 116
    assert stats["trend_cap"] == 0.2
    # 預設上限 1.0：樂觀=100×1.8=180
    main2, _, _, _, _, _ = _run(df)
    assert _main_row(main2)["FC 樂觀情境"] == 180


def test_class_new_volume():
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 100.0, "Shop Monthly Sold Qty (P90D)": 10.0}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_NEW
    assert row["FC 月預估 (採用)"] == row["FC 樂觀情境"]


def test_class_accelerating():
    df = _df([{"Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 120.0,
               "Sold Qty (CM - MTD)": 37.0,  # 37/8×30=138.75
               "Shop Monthly Sold Qty (P90D)": 100.0}])
    main, _, _, _, _, _ = _run(df)
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_ACCEL


def test_class_surge():
    df = _df([{"Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 100.0,
               "Sold Qty (CM - MTD)": 16.0, "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, _, _, _ = _run(df)
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_SURGE


def test_surge_boundary_exact():
    kw = {"Sold Qty (M-2)": 50.0, "Sold Qty (CM - MTD)": 16.0,
          "Shop Monthly Sold Qty (P90D)": 60.0}
    m1 = _run(_df([dict(kw, **{"Sold Qty (M-1)": 75.0})]))[0]
    m2 = _run(_df([dict(kw, **{"Sold Qty (M-1)": 74.0})]))[0]
    assert _main_row(m1)["FC 爆升分類"] == _FC_CLASS_SURGE
    assert _main_row(m2)["FC 爆升分類"] == _FC_CLASS_STABLE


def test_class_stable_uses_standard():
    df = _df([{}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_STABLE
    assert row["FC 月預估 (採用)"] == row["FC 標準情境"] == 58


def test_custom_surge_params():
    df = _df([{"Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 60.0,
               "Sold Qty (CM - MTD)": 16.0, "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, _, _, _ = _run(df, surge_mult=1.1)
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_SURGE


# ── 4. 缺貨還原（取代式；M-2／M-1／CM） ──────────────────────────────────────
def test_oos_cm_restored_and_used():
    # 缺貨列 CM：MTD 月化 285 vs 月均 60 → 還原 285
    # MTD 調整=8/30×285+22/30×60=120；標準=76.5；爆升 → 樂觀=120×1.08=129.6→129
    df = _df([{"Shop Stock Qty (CM)": 0.0, "Sold Qty (CM - MTD)": 76.0,
               "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, raw, _, _ = _run(df)
    assert bool(raw["FC 缺貨還原列"].iloc[0]) is True
    assert raw["FC MTD 還原月化"].iloc[0] == pytest.approx(285.0)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_SURGE
    assert row["FC 月預估 (採用)"] == 129


def test_oos_std_replacement_picks_mtd_when_higher():
    # 平穩 SKU 的缺貨列：還原 285 → 標準=76.5→76
    df = _df([{"Shop Stock Qty (CM)": 0.0, "Sold Qty (CM - MTD)": 76.0,
               "Sold Qty (M-2)": 200.0, "Sold Qty (M-1)": 55.0,
               "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, raw, _, _ = _run(df)
    assert bool(raw["FC 缺貨還原列"].iloc[0]) is True
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_STABLE
    assert _main_row(main)["FC 月預估 (採用)"] == 76


def test_oos_cm_mtd_lower_uses_baseline():
    # 缺貨列：MTD 月化 37.5 vs 月均 60 → 還原 60；標準=58
    df = _df([{"Shop Stock Qty (CM)": 0.0, "Sold Qty (CM - MTD)": 10.0}])
    main, _, _, raw, _, _ = _run(df)
    assert raw["FC MTD 還原月化"].iloc[0] == pytest.approx(60.0)
    assert _main_row(main)["FC 月預估 (採用)"] == 58


def test_non_oos_row_not_restored():
    df = _df([{"Shop Stock Qty (CM)": 5.0, "Sold Qty (CM - MTD)": 76.0,
               "Sold Qty (M-2)": 200.0}])
    main, _, _, raw, _, _ = _run(df)
    assert bool(raw["FC 缺貨還原列"].iloc[0]) is False
    # 非缺貨：MTD 還原月化＝原始月化 285（未被壓低或抬高）
    assert raw["FC MTD 還原月化"].iloc[0] == pytest.approx(285.0)
    assert _main_row(main)["FC 爆升分類"] == _FC_CLASS_STABLE
    assert _main_row(main)["FC 月預估 (採用)"] == 76


def test_oos_nd_row_excluded_from_restore():
    df = _df([{"Shop Stock Qty (CM)": 0.0, "Sold Qty (CM - MTD)": 76.0,
               "RP Type": "ND"}])
    main, _, _, raw, _, _ = _run(df)
    assert bool(raw["FC 缺貨還原列"].iloc[0]) is False
    assert raw["FC MTD 還原月化"].iloc[0] == pytest.approx(285.0)
    row = _main_row(main)
    assert row["FC 月預估 (採用)"] == row["FC 樂觀情境"]
    assert row["FC 建議訂量"] == 0
    assert row["FC SKU 狀態"] == _FC_STATUS_ND


def test_oos_m1_restored_to_baseline():
    df = _df([{"Shop Stock Qty (M-1)": 0.0, "Shop Stock Qty (M-2)": 5.0,
               "Sold Qty (M-1)": 10.0,
               "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, raw, stats, _ = _run(df)
    assert raw["FC M-1 還原值"].iloc[0] == pytest.approx(60.0)
    assert _main_row(main)["FC M-1 合計"] == pytest.approx(60.0)
    assert stats["restored_rows_m1"] == 1
    assert stats["m2m1_stock_missing"] is False
    assert stats["m1_lm_proxy"] is False


def test_oos_m2_restored_to_baseline():
    df = _df([{"Shop Stock Qty (M-2)": 0.0, "Sold Qty (M-2)": 10.0,
               "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, raw, stats, _ = _run(df)
    assert raw["FC M-2 還原值"].iloc[0] == pytest.approx(60.0)
    assert _main_row(main)["FC M-2 合計"] == pytest.approx(60.0)
    assert stats["restored_rows_m2"] == 1


def test_oos_m1_restored_via_lm_proxy():
    # 來源缺 Shop Stock Qty (M-1)，以 Shop Stock Qty (LM) 代理
    df = _df([{"Shop Stock Qty (LM)": 0.0, "Sold Qty (M-1)": 10.0,
               "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, raw, stats, warns = _run(df)
    assert raw["FC M-1 還原值"].iloc[0] == pytest.approx(60.0)
    assert _main_row(main)["FC M-1 合計"] == pytest.approx(60.0)
    assert stats["m1_lm_proxy"] is True
    assert stats["m1_stock_missing"] is True
    assert stats["restored_rows_m1"] == 1
    assert any("Shop Stock Qty (LM)" in w for w in warns)


def test_missing_stock_columns_skip_restore_with_warning():
    # BASE 無 Shop Stock Qty (M-2)／(M-1)：即使引擎 Is OOS 全真亦不還原
    df = _df([{"Sold Qty (M-1)": 10.0}])
    _, _, _, raw, stats, warns = _run(df)
    assert raw["FC M-1 還原值"].iloc[0] == pytest.approx(10.0)
    assert raw["FC M-2 還原值"].iloc[0] == pytest.approx(50.0)
    assert stats["m2m1_stock_missing"] is True
    assert stats["restored_rows_m1"] == 0
    assert any("Shop Stock Qty (M-1)" in w for w in warns)
    assert any("Shop Stock Qty (M-2)" in w for w in warns)


def test_nd_not_restored_m1():
    df = _df([{"RP Type": "ND", "Shop Stock Qty (M-1)": 0.0,
               "Sold Qty (M-1)": 10.0}])
    _, _, _, raw, stats, _ = _run(df)
    assert raw["FC M-1 還原值"].iloc[0] == pytest.approx(10.0)
    assert stats["restored_rows_m1"] == 0


# ── 5. 位置分流 ─────────────────────────────────────────────────────────────
def test_warehouse_position():
    df = _df([{}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 位置合計"] == 115
    assert row["FC 建議訂貨對象"] == _FC_TARGET_D001


def test_vendor_position():
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
    assert _main_row(main)["FC 位置合計"] == 110
    assert any("Pending Receive" in w for w in warns)


def test_d001_po_missing_warns():
    df = _df([{}]).drop(columns=["D001 On Order PO Qty (CM)"])
    main, _, _, _, _, warns = _run(df)
    assert _main_row(main)["FC 位置合計"] == 105
    assert any("D001 On Order PO" in w for w in warns)


def test_d001_first_value_not_summed():
    df = _df([
        {"Site Code": "S01", "Shop Stock Qty (CM)": 20.0},
        {"Site Code": "S02", "Shop Stock Qty (CM)": 30.0,
         "Item_Shop": "Prod-S02-Shop2"},
    ])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC D001 庫存 (Net)"] == 80
    assert row["FC D001 PO"] == 10
    assert row["FC 位置合計"] == 50 + 80 + 10 + 10


def test_d001_inconsistent_warns():
    df = _df([
        {"Site Code": "S01"},
        {"Site Code": "S02", "D001 Net Stock Qty (CM)": 5.0,
         "Item_Shop": "Prod-S02-Shop2"},
    ])
    main, _, _, _, _, warns = _run(df)
    assert _main_row(main)["FC D001 庫存 (Net)"] == 80
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
    # 平穩：月預估 58；目標=ceil(58/30.5×30)=58；位置 115 → 缺口 0 → 倉位足夠
    df = _df([{}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 目標覆蓋量"] == 58
    assert row["FC 缺口"] == 0
    assert row["FC 建議訂量"] == 0
    assert row["FC SKU 狀態"] == _FC_STATUS_ENOUGH


def test_order_equals_gap_when_within_cap():
    # 新起量：月化 285 → 調整 120；樂觀 120；目標=ceil(120/30.5×30)=119
    # 位置=20（店存；D001／在途設 0）→ 缺口 99；上限=ceil(240)−20=220 → 訂 99
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 60.0,
               "Shop Stock Qty (CM)": 20.0,
               "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0,
               "Shop Pending Receive (±7 days)": 0.0}])
    main, _, _, _, _, _ = _run(df)
    row = _main_row(main)
    assert row["FC 爆升分類"] == _FC_CLASS_NEW
    assert row["FC 月預估 (採用)"] == 120
    assert row["FC 缺口"] == 99
    assert row["FC 建議訂量"] == 99
    assert row["FC SKU 狀態"] == _FC_STATUS_ORDER


def test_cap_binds():
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 60.0,
               "Shop Stock Qty (CM)": 20.0,
               "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0,
               "Shop Pending Receive (±7 days)": 0.0}])
    main, _, _, _, _, _ = _run(df, cap_months=0.5)
    row = _main_row(main)
    assert row["FC 上限空間"] == 40
    assert row["FC 建議訂量"] == 40
    assert row["FC SKU 狀態"] == _FC_STATUS_ORDER_CAPPED


# ── 7. 店舖覆蓋 ─────────────────────────────────────────────────────────────
def test_coverage_rows_and_risk_flags():
    # S02 缺貨(庫存0)→已缺貨；S01 店存20+在途5, 日銷=76/8=9.5 → 2.6日 → ≤7日
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
    assert list(cov["Site Code"]) == ["S02", "S01"]
    assert _main_row(main)["FC 缺貨風險店數"] == 2
    assert _main_row(main)["FC 現已缺貨店數"] == 1
    assert _main_row(main)["FC 鋪貨店數"] == 2


def test_coverage_target_flag():
    # 目標 30 日：可用日數 8.5 日 → ≤目標天數
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
               "Shop Stock Qty (CM)": 69.0,
               "Shop Pending Receive (±7 days)": 12.0,
               "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0}])
    main, cov, _, _, _, _ = _run(df)
    assert _main_row(main)["FC 建議訂量"] == 1  # 保證覆蓋表非空
    assert cov["FC 風險旗"].iloc[0] == _FC_RISK_TARGET


def test_coverage_daily_fallback_p90d():
    # MTD=0 的列 → 日銷＝P90D÷30.5
    df = _df([
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "Sold Qty (P90D)": 76.0, "Shop Stock Qty (CM)": 0.0,
         "Shop Pending Receive (±7 days)": 0.0, "Site Code": "S01",
         "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0},
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 0.0, "Shop Monthly Sold Qty (P90D)": 0.0,
         "Sold Qty (P90D)": 1800.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Pending Receive (±7 days)": 0.0, "Site Code": "S02",
         "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0,
         "Item_Shop": "Prod-S02-Shop2"},
    ])
    main, cov, _, _, _, _ = _run(df)
    assert _main_row(main)["FC 建議訂量"] > 0
    by_site = {r["Site Code"]: r for _, r in cov.iterrows()}
    assert by_site["S02"]["FC 店舖日銷"] == pytest.approx(round(1800.0 / 30.5, 2), rel=1e-6)
    assert by_site["S02"]["FC 風險旗"] == _FC_RISK_7D


def test_coverage_empty_when_no_order():
    df = _df([{}])
    _, cov, _, _, stats, _ = _run(df)
    assert cov.empty
    assert list(cov.columns) == _FORECAST_COVERAGE_COLUMNS
    assert stats["coverage_rows"] == 0


def test_coverage_ok_flag():
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
    # S02 日銷=300/30.5≈9.84, 可用=5000/9.84≈508日 → 足夠
    assert by_site["S02"]["FC 風險旗"] == _FC_RISK_OK


# ── 8. Buyer 彙總與 stats ───────────────────────────────────────────────────
def test_buyer_summary():
    df = _df([
        {"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "Shop Stock Qty (CM)": 20.0,
         "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0},
        {"SKU": 2, "Buyer": "B2", "Item_Shop": "P2-S01",
         "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "Shop Stock Qty (CM)": 20.0, "Supply Source": "Vendor",
         "Shop On Order PO Qty (CM)": 3.0,
         "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0},
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
    assert _FORECAST_MAIN_COLUMNS.index("FC MTD 調整月化") == _FORECAST_MAIN_COLUMNS.index("FC 樂觀情境") + 1
    assert _FORECAST_MAIN_COLUMNS.index("FC 趨勢增速") == _FORECAST_MAIN_COLUMNS.index("FC MTD 調整月化") + 1


def test_stats_params_echo():
    _, _, _, _, stats, _ = _run(_df([{}]), target_days=45, cap_months=3.0,
                                surge_mult=2.0, surge_min_base=5, accel_mult=1.2, damp=0.9,
                                w_p90d=0.5, w_m1=0.25, w_mtd=0.25,
                                trend_cap=1.5, mtd_date=MTD_DATE)
    assert stats["target_days"] == 45
    assert stats["cap_months"] == 3.0
    assert stats["surge_mult"] == 2.0
    assert stats["surge_min_base"] == 5
    assert stats["accel_mult"] == 1.2
    assert stats["damp"] == 0.9
    assert stats["model"] == "v2"
    assert stats["w_p90d"] == pytest.approx(0.5)
    assert stats["w_m1"] == pytest.approx(0.25)
    assert stats["w_mtd"] == pytest.approx(0.25)
    assert stats["trend_cap"] == 1.5
    assert stats["mtd_date"] == "2026-09-08"
    assert stats["mtd_days"] == MTD_DAYS
    assert stats["launch_prorated_skus"] == 0


def test_stats_v2_fields():
    _, _, _, _, stats, _ = _run(_df([{}]))
    assert stats["model"] == "v2"
    assert stats["w_p90d"] == pytest.approx(0.4)
    assert stats["w_m1"] == pytest.approx(0.3)
    assert stats["w_mtd"] == pytest.approx(0.3)
    assert stats["trend_cap"] == 1.0
    assert stats["mtd_date"] is None
    assert stats["restored_rows_m2"] == 0
    assert stats["restored_rows_m1"] == 0
    assert stats["restored_rows_cm"] == 0
    assert stats["m2m1_stock_missing"] is True
    assert stats["m2_stock_missing"] is True
    assert stats["m1_stock_missing"] is True
    assert stats["m1_lm_proxy"] is False


def test_param_clamp_warns():
    _, _, _, _, stats, warns = _run(_df([{}]), target_days=999, damp="bad",
                                    w_p90d=9, trend_cap=99)
    assert stats["target_days"] == 90
    assert stats["damp"] == 0.8
    assert stats["trend_cap"] == 3.0
    # w_p90d 箝制至 1.0，與 0.3／0.3 正規化 → 0.625
    assert stats["w_p90d"] == pytest.approx(0.625)
    assert len([w for w in warns if "目標覆蓋天數" in w]) == 1
    assert len([w for w in warns if "衰減係數" in w]) == 1
    assert len([w for w in warns if "月均基準權重" in w]) == 1
    assert len([w for w in warns if "趨勢增速上限" in w]) == 1


def test_weight_normalization_and_zero_fallback():
    _, _, _, _, stats, _ = _run(_df([{}]), w_p90d=0.2, w_m1=0.3, w_mtd=0.5)
    assert (stats["w_p90d"], stats["w_m1"], stats["w_mtd"]) == pytest.approx((0.2, 0.3, 0.5))
    _, _, _, _, stats0, warns = _run(_df([{}]), w_p90d=0, w_m1=0, w_mtd=0)
    assert (stats0["w_p90d"], stats0["w_m1"], stats0["w_mtd"]) == pytest.approx((0.4, 0.3, 0.3))
    assert any("權重合計為 0" in w for w in warns)


# ── 9. 開賣日折算 ───────────────────────────────────────────────────────────
def test_launch_proration_m1():
    # 2026-08-20 開賣：M-1（8 月）有效 12 日 → factor=31/12；M-1 120 → 310
    df = _df([{"Launch Date": "2026-08-20",
               "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 120.0,
               "Sold Qty (CM - MTD)": 16.0, "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, _, stats, _ = _run(df, mtd_date=MTD_DATE)
    row = _main_row(main)
    assert stats["launch_prorated_skus"] == 1
    assert row["FC M-1 合計"] == pytest.approx(310.0)
    assert row["FC 爆升分類"] == _FC_CLASS_SURGE
    assert row["FC 月預估 (採用)"] == row["FC 樂觀情境"]


def test_launch_proration_cm_denominator():
    # 2026-09-05 開賣、MTD 截止 09-08：有效 4 日 → 分母 max(4,7)=7
    df = _df([{"Launch Date": "2026-09-05",
               "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 35.0, "Shop Monthly Sold Qty (P90D)": 10.0}])
    _, _, _, raw, stats, _ = _run(df, mtd_date=MTD_DATE)
    assert raw["FC MTD 月化"].iloc[0] == pytest.approx(150.0)  # 35/7×30
    assert stats["launch_prorated_skus"] == 1


def test_launch_proration_min_active_days_7():
    df = _df([{"Launch Date": "2026-09-07",
               "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 14.0, "Shop Monthly Sold Qty (P90D)": 10.0}])
    _, _, _, raw, _, _ = _run(df, mtd_date=MTD_DATE)
    assert raw["FC MTD 月化"].iloc[0] == pytest.approx(60.0)  # 14/7×30


def test_no_mtd_date_no_proration():
    df = _df([{"Launch Date": "2026-08-20",
               "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 120.0,
               "Sold Qty (CM - MTD)": 16.0, "Shop Monthly Sold Qty (P90D)": 60.0}])
    main, _, _, _, stats, _ = _run(df)
    assert stats["launch_prorated_skus"] == 0
    assert _main_row(main)["FC M-1 合計"] == pytest.approx(120.0)


def test_raw_display_columns_include_restored():
    df = _df([{}])
    _, _, _, raw, _, _ = _run(df)
    cols = _forecast_raw_display_columns(raw)
    for c in ("FC M-2 還原值", "FC M-1 還原值", "FC MTD 還原月化"):
        assert c in cols


# ── 10. Excel 輸出 ──────────────────────────────────────────────────────────
def test_excel_sheets_and_guide_first():
    df = _df([{"Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
               "Sold Qty (CM - MTD)": 76.0, "Shop Monthly Sold Qty (P90D)": 10.0,
               "Shop Stock Qty (CM)": 20.0,
               "D001 Net Stock Qty (CM)": 0.0, "D001 On Order PO Qty (CM)": 0.0,
               "Shop Pending Receive (±7 days)": 0.0}])
    main, cov, buyer, raw, stats, _ = _run(df)
    buf = generate_excel_forecast(main, cov, buyer, raw, stats)
    wb = load_workbook(buf)
    assert wb.sheetnames == ["說明", "Forecast 主表", "店舖覆蓋", "按 Buyer 彙總", "RAW", "Forecast 摘要"]
    assert wb.active.title == "說明"
    ws = wb["Forecast 主表"]
    headers = [c.value for c in ws[2]]
    assert headers == _FORECAST_MAIN_COLUMNS
    order_idx = _FORECAST_MAIN_COLUMNS.index("FC 建議訂量") + 1
    assert ws.cell(row=3, column=order_idx).fill.start_color.rgb != "00000000"
    ws_cov = wb["店舖覆蓋"]
    assert [c.value for c in ws_cov[2]] == [c for c in _FORECAST_COVERAGE_COLUMNS if c in cov.columns]
    # 說明表照錄 v2 模型資訊
    guide_values = [str(c.value) for row in wb["說明"].iter_rows() for c in row if c.value is not None]
    assert any("模型權重" == v for v in guide_values)
    assert any("模型版本" == v for v in guide_values)


def test_excel_empty_main():
    df = _df([{}])
    main, cov, buyer, raw, stats, _ = _run(df)
    buf = generate_excel_forecast(main.iloc[0:0], cov.iloc[0:0], buyer.iloc[0:0], raw.iloc[0:0], stats)
    wb = load_workbook(buf)
    assert wb.sheetnames == ["說明", "Forecast 主表", "店舖覆蓋", "按 Buyer 彙總", "RAW", "Forecast 摘要"]
    assert wb["Forecast 主表"].cell(row=2, column=1).value == "無符合資料"
