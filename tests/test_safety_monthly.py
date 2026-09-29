"""Validation tests for Safety Stock 月度變化分析.

覆蓋：MONTH／YEAR 必要欄位、月份解析、環比事件、Safety 原始值分類、
RP 轉變、最新月份月化（Sold 與 Net 同時）、維度樞紐、Excel 工作表結構。
"""
import datetime
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from transformer import (
    transform_safety_monthly,
    generate_excel_safety_monthly,
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1,
    "Item_Shop": "Prod-S01-Shop", "Site Code": "S01",
    "Shop Nature": "Non Tourist", "RP Type": "RF",
    "Shop Stock Status": "Normal",
    "Adj. SST": "SST1", "Adj. MCH1": "01 - SKIN CARE",
    "Adj. Safety Stock": 5.0,
    "Sold Qty (CM - MTD)": 10.0, "Net Sales (CM - MTD)": 100.0,
    "Shop Stock Qty (CM)": 20.0,
    "Month": 5, "Year": 2026,
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def _run(rows, ref_date=None, **kw):
    if ref_date is None:
        ref_date = datetime.date(2026, 6, 30)
    return transform_safety_monthly(_df(rows), ref_date=ref_date, **kw)


# ── 1. 必要欄位 ──────────────────────────────────────────────────────
def test_missing_month_year_rejected():
    df = _df([{"SKU": 1, "Month": 5}, {"SKU": 1, "Month": 6}]).drop(columns=["Year"])
    try:
        transform_safety_monthly(df, ref_date=datetime.date(2026, 6, 30))
    except ValueError as e:
        assert "MONTH" in str(e) and "YEAR" in str(e)
    else:
        raise AssertionError("缺少 YEAR 時應拋出 ValueError")
    df2 = _df([{"SKU": 1}, {"SKU": 2}]).drop(columns=["Month"])
    try:
        transform_safety_monthly(df2, ref_date=datetime.date(2026, 6, 30))
    except ValueError as e:
        assert "MONTH" in str(e)
    else:
        raise AssertionError("缺少 MONTH 時應拋出 ValueError")


def test_missing_required_columns_rejected():
    for col in ("Site Code", "Adj. Safety Stock", "RP Type",
                "Sold Qty (CM - MTD)", "Net Sales (CM - MTD)", "Shop Stock Qty (CM)"):
        df = _df([{"SKU": 1, "Month": 5}, {"SKU": 1, "Month": 6}]).drop(columns=[col])
        try:
            transform_safety_monthly(df, ref_date=datetime.date(2026, 6, 30))
        except ValueError as e:
            assert col in str(e), col
        else:
            raise AssertionError(f"缺少 {col} 時應拋出 ValueError")


def test_single_month_rejected():
    df = _df([{"SKU": 1, "Month": 5}, {"SKU": 2, "Month": 5}])
    try:
        transform_safety_monthly(df, ref_date=datetime.date(2026, 5, 31))
    except ValueError as e:
        assert "2 個" in str(e)
    else:
        raise AssertionError("只有 1 個月份時應拋出 ValueError")


def test_invalid_month_rows_skipped_with_warning():
    change, *_rest, stats, warnings = _run([
        {"SKU": 1, "Month": 5},
        {"SKU": 1, "Month": 6},
        {"SKU": 2, "Month": 13},  # 無效月份
        {"SKU": 2, "Month": 6},
    ])
    assert any("MONTH" in w for w in warnings)
    assert stats["ym_count"] == 2


# ── 2. 環比事件 ──────────────────────────────────────────────────────
def test_mom_event_and_safety_classification():
    change, *_rest, stats, _ = _run([
        {"SKU": 1, "Month": 5, "Adj. Safety Stock": 5.0},
        {"SKU": 1, "Month": 6, "Adj. Safety Stock": 8.0},
        {"SKU": 2, "Month": 5, "Adj. Safety Stock": 5.0},
        {"SKU": 2, "Month": 6, "Adj. Safety Stock": 3.0},
        {"SKU": 3, "Month": 5, "Adj. Safety Stock": 5.0},
        {"SKU": 3, "Month": 6, "Adj. Safety Stock": 5.0},
    ])
    assert len(change) == 3
    by_sku = dict(zip(change["SKU"].astype(str), change["Safety變化分類"]))
    assert by_sku == {"1": "增加", "2": "減少", "3": "不變"}
    assert stats["safety_inc"] == 1 and stats["safety_dec"] == 1 and stats["safety_same"] == 1
    assert stats["safety_delta_total"] == 1  # (8-5)+(3-5)+0


def test_rp_transition_recorded():
    change, *_rest, _s, _w = _run([
        {"SKU": 1, "Month": 5, "RP Type": "RF"},
        {"SKU": 1, "Month": 6, "RP Type": "ND"},
    ])
    assert change.iloc[0]["RP轉變"] == "RF→ND"


def test_new_combo_has_no_event():
    # SKU 4 只有 6 月資料，無上月可比，不產生事件
    change, *_rest, stats, _ = _run([
        {"SKU": 1, "Month": 5}, {"SKU": 1, "Month": 6},
        {"SKU": 4, "Month": 6},
    ])
    assert set(change["SKU"].astype(str)) == {"1"}
    assert stats["total_skus"] == 2


# ── 3. 月化 ──────────────────────────────────────────────────────────
def test_latest_month_annualized_both_columns():
    # 6 月共 30 日，截止日 6 月 15 日 → 係數 2.0
    change, *_rest, stats, _ = _run(
        [
            {"SKU": 1, "Month": 5, "Sold Qty (CM - MTD)": 30.0, "Net Sales (CM - MTD)": 300.0},
            {"SKU": 1, "Month": 6, "Sold Qty (CM - MTD)": 10.0, "Net Sales (CM - MTD)": 100.0},
        ],
        ref_date=datetime.date(2026, 6, 15),
    )
    assert stats["annual_factor"] == 2.0
    row = change.iloc[0]
    assert row["本月 Sold月化"] == 20
    assert row["本月 Net月化"] == 200.0
    # 環比以月化基準：20-30=-10；200-300=-100
    assert row["Sold變化量(月化基準)"] == -10
    assert row["Net變化量(月化基準)"] == -100.0


def test_full_month_factor_one():
    change, *_rest, stats, _ = _run(
        [
            {"SKU": 1, "Month": 5, "Sold Qty (CM - MTD)": 30.0},
            {"SKU": 1, "Month": 6, "Sold Qty (CM - MTD)": 10.0},
        ],
        ref_date=datetime.date(2026, 6, 30),
    )
    assert stats["annual_factor"] == 1.0
    assert change.iloc[0]["本月 Sold月化"] == 10


# ── 4. 樞紐與 Excel ──────────────────────────────────────────────────
def test_pivots_three_blocks_four_metrics():
    change, mch1, sst, cross, shop, raw, stats, _ = _run([
        {"SKU": 1, "Month": 5, "Adj. MCH1": "01 - SKIN CARE", "Adj. SST": "SST1",
         "Sold Qty (CM - MTD)": 100.0, "Net Sales (CM - MTD)": 1000.0,
         "Shop Stock Qty (CM)": 50.0, "Adj. Safety Stock": 10.0},
        {"SKU": 1, "Month": 6, "Adj. MCH1": "01 - SKIN CARE", "Adj. SST": "SST1",
         "Sold Qty (CM - MTD)": 150.0, "Net Sales (CM - MTD)": 1500.0,
         "Shop Stock Qty (CM)": 40.0, "Adj. Safety Stock": 12.0},
    ])
    expected_items = [
        "Sold Qty 每月合計", "Sold Qty 月環比%", "Sold Qty 月環比",
        "Net Sales 每月合計", "Net Sales 月環比%", "Net Sales 月環比",
        "Net Stock Qty 每月合計", "Net Stock Qty 月環比%", "Net Stock Qty 月環比",
        "Safety Stock Qty 每月合計", "Safety Stock Qty 月環比%", "Safety Stock Qty 月環比",
    ]
    assert list(dict.fromkeys(mch1["項目"])) == expected_items
    assert list(mch1.columns)[:3] == ["Adj. MCH1", "項目", "2026-05"]
    sold_total = mch1[mch1["項目"] == "Sold Qty 每月合計"].iloc[0]
    assert sold_total["2026-05"] == 100 and sold_total["2026-06"] == 150
    assert sold_total["全部合計"] == 250
    sold_pct = mch1[mch1["項目"] == "Sold Qty 月環比%"].iloc[0]
    assert sold_pct["2026-05"] is None or pd.isna(sold_pct["2026-05"])
    assert sold_pct["2026-06"] == "+50.0%"
    sold_abs = mch1[mch1["項目"] == "Sold Qty 月環比"].iloc[0]
    assert sold_abs["2026-06"] == 50
    net_pct = mch1[mch1["項目"] == "Net Sales 月環比%"].iloc[0]
    assert net_pct["2026-06"] == "+50.0%"
    stock_abs = mch1[mch1["項目"] == "Net Stock Qty 月環比"].iloc[0]
    assert stock_abs["2026-06"] == -10
    safety_abs = mch1[mch1["項目"] == "Safety Stock Qty 月環比"].iloc[0]
    assert safety_abs["2026-06"] == 2
    # 總計列：每月合計加總；月環比以總計計算
    total_pct = mch1[(mch1["Adj. MCH1"] == "總計") & (mch1["項目"] == "Sold Qty 月環比%")].iloc[0]
    assert total_pct["2026-06"] == "+50.0%"
    assert set(sst["Adj. SST"]) >= {"SST1", "總計"}
    assert len(cross) == 2 * 12  # 1 組合＋總計 × 12 項目列
    assert set(shop["Site Code"]) == {"S01"}
    assert "YEAR-MONTH" in raw.columns


def test_mom_pct_zero_base_blank():
    change, mch1, *_rest, stats, _ = _run([
        {"SKU": 1, "Month": 5, "Sold Qty (CM - MTD)": 0.0},
        {"SKU": 1, "Month": 6, "Sold Qty (CM - MTD)": 10.0},
    ])
    sold_pct = mch1[mch1["項目"] == "Sold Qty 月環比%"].iloc[0]
    assert sold_pct["2026-06"] is None or pd.isna(sold_pct["2026-06"])
    sold_abs = mch1[mch1["項目"] == "Sold Qty 月環比"].iloc[0]
    assert sold_abs["2026-06"] == 10


def test_excel_eight_sheets_and_guide_first():
    change, mch1, sst, cross, shop, raw, stats, _ = _run([
        {"SKU": 1, "Month": 5}, {"SKU": 1, "Month": 6},
    ])
    buf = generate_excel_safety_monthly(change, mch1, sst, cross, shop, raw, stats)
    wb = load_workbook(BytesIO(buf.getvalue()))
    assert wb.sheetnames == ["說明", "SKU月度變化", "By Adj. MCH1", "By Adj. SST",
                             "By MCH1×SST", "By Shop", "RP Type 變化", "月度摘要", "RAW"]
    assert wb.active.title == "說明"


# ── 5. RP Type 轉變統計 ──────────────────────────────────────────────
def test_rp_summary_transition_counts():
    change, _m, _s, _c, _sh, _r, stats, _ = _run([
        {"SKU": 1, "Month": 5, "RP Type": "RF"},
        {"SKU": 1, "Month": 6, "RP Type": "ND"},
        {"SKU": 2, "Month": 5, "RP Type": "ND"},
        {"SKU": 2, "Month": 6, "RP Type": "RF"},
        {"SKU": 3, "Month": 5, "RP Type": "RF"},
        {"SKU": 3, "Month": 6, "RP Type": "RF"},
        {"SKU": 4, "Month": 5, "RP Type": "ND"},
        {"SKU": 4, "Month": 6, "RP Type": "ND"},
    ])
    assert stats["rp_rf_to_nd"] == 1 and stats["rp_nd_to_rf"] == 1
    assert stats["rp_rf_to_nd_skus"] == 1 and stats["rp_nd_to_rf_skus"] == 1
    rp = stats["rp_summary"]
    assert list(rp["RP轉變"]) == ["RF→RF", "ND→ND", "RF→ND", "ND→RF", "總計"]
    rf_nd = rp[rp["RP轉變"] == "RF→ND"].iloc[0]
    assert rf_nd["2026-05→2026-06 事件數"] == 1
    assert rf_nd["2026-05→2026-06 SKU數"] == 1
    assert rf_nd["總計 事件數"] == 1
    total = rp[rp["RP轉變"] == "總計"].iloc[0]
    assert total["總計 事件數"] == 4 and total["總計 SKU數"] == 4


def test_rp_summary_empty_change():
    from transformer import _build_safm_rp_summary
    rp = _build_safm_rp_summary(pd.DataFrame())
    assert list(rp.columns) == ["RP轉變"] and rp.empty
