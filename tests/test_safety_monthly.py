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
    by_sku = dict(zip(change["SKU"].astype(str), change["Safety Stock Qty 變化分類"]))
    assert by_sku == {"1": "增加", "2": "減少", "3": "不變"}
    assert stats["safety_inc"] == 1 and stats["safety_dec"] == 1 and stats["safety_same"] == 1
    assert stats["safety_delta_total"] == 1  # (8-5)+(3-5)+0


def test_rp_transition_recorded():
    change, *_rest, _s, _w = _run([
        {"SKU": 1, "Month": 5, "RP Type": "RF"},
        {"SKU": 1, "Month": 6, "RP Type": "ND"},
    ])
    assert change.iloc[0]["RP Type 轉變"] == "RF→ND"


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
    assert row["本月 Sold Qty 月化"] == 20
    assert row["本月 Net Sales 月化"] == 200.0
    # 環比以月化基準：20-30=-10；200-300=-100
    assert row["Sold Qty 變化量(月化基準)"] == -10
    assert row["Net Sales 變化量(月化基準)"] == -100.0


def test_full_month_factor_one():
    change, *_rest, stats, _ = _run(
        [
            {"SKU": 1, "Month": 5, "Sold Qty (CM - MTD)": 30.0},
            {"SKU": 1, "Month": 6, "Sold Qty (CM - MTD)": 10.0},
        ],
        ref_date=datetime.date(2026, 6, 30),
    )
    assert stats["annual_factor"] == 1.0
    assert change.iloc[0]["本月 Sold Qty 月化"] == 10


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
    assert list(shop.columns) == [
        "Site Code", "Shop Name", "YEAR-MONTH",
        "SKU數", "Safety Stock Qty 合計",
        "Sold Qty 原始合計", "Sold Qty 月化合計",
        "Net Sales 原始合計", "Net Sales 月化合計",
        "Net Stock Qty 合計",
        "Sold Qty 環比變化", "Net Sales 環比變化", "Safety Stock Qty 環比變化",
        "RF→ND SKU數", "ND→RF SKU數",
    ]
    # 樞紐表尾附 RF→ND／ND→RF SKU數（此資料全為 RF→RF，故為 0）
    assert list(mch1.columns)[-2:] == ["RF→ND SKU數", "ND→RF SKU數"]
    assert (mch1["RF→ND SKU數"] == 0).all() and (mch1["ND→RF SKU數"] == 0).all()
    assert list(sst.columns)[-2:] == ["RF→ND SKU數", "ND→RF SKU數"]
    assert list(cross.columns)[-2:] == ["RF→ND SKU數", "ND→RF SKU數"]
    assert "YEAR-MONTH" in raw.columns


def test_excel_pivot_side_by_side_layout():
    # Excel 樞紐工作表照報告設計：四項指標區塊左右並排，上下三段，總計列置段末
    change, mch1, sst, cross, shop, raw, stats, _ = _run([
        {"SKU": 1, "Month": 5, "Adj. MCH1": "01 - SKIN CARE", "Adj. SST": "SST1",
         "Sold Qty (CM - MTD)": 100.0, "Net Sales (CM - MTD)": 1000.0,
         "Shop Stock Qty (CM)": 50.0, "Adj. Safety Stock": 10.0},
        {"SKU": 1, "Month": 6, "Adj. MCH1": "01 - SKIN CARE", "Adj. SST": "SST1",
         "Sold Qty (CM - MTD)": 150.0, "Net Sales (CM - MTD)": 1500.0,
         "Shop Stock Qty (CM)": 40.0, "Adj. Safety Stock": 12.0},
    ])
    buf = generate_excel_safety_monthly(change, mch1, sst, cross, shop, raw, stats)
    wb = load_workbook(BytesIO(buf.getvalue()))
    ws = wb["By Adj. SST"]
    # 表頭：維度欄一個，其後四個區塊（項目＋各月＋全部合計）
    assert [ws.cell(1, c).value for c in range(1, 6)] == [
        "Adj. SST", "項目", "2026-05", "2026-06", "全部合計"]
    assert ws.cell(1, 6).value == "項目" and ws.cell(1, 9).value == "全部合計"
    assert ws.cell(1, 10).value == "項目" and ws.cell(1, 13).value == "全部合計"
    assert ws.cell(1, 14).value == "項目" and ws.cell(1, 17).value == "全部合計"
    # 附加統計欄（RF→ND／ND→RF SKU數）置於四個區塊之後
    assert ws.cell(1, 18).value == "RF→ND SKU數" and ws.cell(1, 19).value == "ND→RF SKU數"
    assert ws.cell(2, 18).value == 0 and ws.cell(2, 19).value == 0
    # 第一段（每月合計）：一列一維度，四區塊並排
    assert ws.cell(2, 1).value == "SST1"
    assert [ws.cell(2, c).value for c in (2, 6, 10, 14)] == [
        "Sold Qty 每月合計", "Net Sales 每月合計",
        "Net Stock Qty 每月合計", "Safety Stock Qty 每月合計",
    ]
    assert (ws.cell(2, 3).value, ws.cell(2, 4).value, ws.cell(2, 5).value) == (100, 150, 250)
    assert (ws.cell(2, 7).value, ws.cell(2, 8).value, ws.cell(2, 9).value) == (1000, 1500, 2500)
    assert (ws.cell(2, 11).value, ws.cell(2, 12).value, ws.cell(2, 13).value) == (50, 40, 90)
    assert (ws.cell(2, 15).value, ws.cell(2, 16).value, ws.cell(2, 17).value) == (10, 12, 22)
    # 總計列置段末（粗體＋灰底），其後為月環比% 段（首月與全部合計留空）
    assert ws.cell(3, 1).value == "總計"
    assert ws.cell(3, 2).font.bold
    assert str(ws.cell(3, 2).fill.start_color.rgb).endswith("D9D9D9")
    assert ws.cell(4, 2).value == "Sold Qty 月環比%"
    assert ws.cell(4, 3).value is None
    assert ws.cell(4, 4).value == "+50.0%"
    assert ws.cell(4, 5).value is None
    assert ws.cell(5, 1).value == "總計"
    # 第三段（月環比）絕對值：Sold +50、Net Sales +500、Net Stock −10、Safety +2
    assert ws.cell(6, 2).value == "Sold Qty 月環比"
    assert ws.cell(6, 4).value == 50
    assert ws.cell(6, 8).value == 500
    assert ws.cell(6, 12).value == -10
    assert ws.cell(6, 16).value == 2
    assert ws.cell(7, 1).value == "總計"
    # 凍結維度欄＋表頭列
    assert ws.freeze_panes == "B2"
    # By MCH1×SST：兩欄維度，區塊自 C 欄起，凍結兩欄
    wc = wb["By MCH1×SST"]
    assert [wc.cell(1, c).value for c in range(1, 4)] == ["Adj. MCH1", "Adj. SST", "項目"]
    assert wc.freeze_panes == "C2"
    assert wc.cell(2, 3).value == "Sold Qty 每月合計"
    assert wc.cell(2, 1).value == "01 - SKIN CARE" and wc.cell(2, 2).value == "SST1"


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
    assert list(rp["RP Type 轉變"]) == ["RF→RF", "ND→ND", "RF→ND", "ND→RF", "總計"]
    rf_nd = rp[rp["RP Type 轉變"] == "RF→ND"].iloc[0]
    assert rf_nd["2026-05→2026-06 事件數"] == 1
    assert rf_nd["2026-05→2026-06 SKU數"] == 1
    assert rf_nd["總計 事件數"] == 1
    total = rp[rp["RP Type 轉變"] == "總計"].iloc[0]
    assert total["總計 事件數"] == 4 and total["總計 SKU數"] == 4


def test_rp_summary_empty_change():
    from transformer import _build_safm_rp_summary
    rp = _build_safm_rp_summary(pd.DataFrame())
    assert list(rp.columns) == ["RP Type 轉變"] and rp.empty


# ── 6. 首尾月份對比 ──────────────────────────────────────────────────
def test_rp_endpoints_first_vs_last():
    from transformer import _build_safm_rp_endpoints
    change, _m, _s, _c, _sh, _r, stats, _ = _run([
        {"SKU": 1, "Month": 5, "RP Type": "RF", "Brand": "B"},
        {"SKU": 1, "Month": 6, "RP Type": "RF", "Brand": "B"},
        {"SKU": 1, "Month": 9, "RP Type": "ND", "Brand": "B"},
        {"SKU": 2, "Month": 5, "RP Type": "ND"},
        {"SKU": 2, "Month": 9, "RP Type": "RF"},
        {"SKU": 3, "Month": 5, "RP Type": "RF"},
        {"SKU": 3, "Month": 9, "RP Type": "RF"},
        {"SKU": 4, "Month": 9, "RP Type": "RF"},  # 缺最早月份，不納入
    ])
    info = stats["rp_end_info"]
    assert (info["first"], info["last"]) == ("2026-05", "2026-09")
    assert info["comparable"] == 3 and info["excluded"] == 1
    assert info["rf_to_nd"] == 1 and info["nd_to_rf"] == 1
    detail = stats["rp_endpoints"]
    assert list(detail["首尾轉變"]) == ["RF→ND", "ND→RF", "RF→RF"]
    assert list(detail["是否轉變"]) == ["是", "是", "否"]
    assert set(detail["SKU"].astype(str)) == {"1", "2", "3"}


def test_rp_endpoints_value_comparison():
    # 首尾對比附最早／最新月份 Sold Qty 月化、Net Sales 月化與 Safety Stock Qty 對比
    change, _m, _s, _c, _sh, _r, stats, _ = _run([
        {"SKU": 1, "Month": 5, "Sold Qty (CM - MTD)": 100.0, "Net Sales (CM - MTD)": 1000.0,
         "Adj. Safety Stock": 10.0},
        {"SKU": 1, "Month": 9, "Sold Qty (CM - MTD)": 150.0, "Net Sales (CM - MTD)": 800.0,
         "Adj. Safety Stock": 4.0},
        {"SKU": 2, "Month": 9, "Sold Qty (CM - MTD)": 50.0},  # 缺最早月份，不納入
    ])
    detail = stats["rp_endpoints"]
    assert list(detail.columns)[-11:] == [
        "最早 Sold Qty 月化", "最新 Sold Qty 月化", "Sold Qty 變化量(月化基準)", "Sold Qty 變化率",
        "最早 Net Sales 月化", "最新 Net Sales 月化", "Net Sales 變化量(月化基準)", "Net Sales 變化率",
        "最早 Safety Stock Qty", "最新 Safety Stock Qty", "Safety Stock Qty 變化量",
    ]
    row = detail[detail["SKU"].astype(str) == "1"].iloc[0]
    assert (row["最早 Sold Qty 月化"], row["最新 Sold Qty 月化"]) == (100, 150)
    assert row["Sold Qty 變化量(月化基準)"] == 50
    assert row["Sold Qty 變化率"] == 0.5
    assert (row["最早 Net Sales 月化"], row["最新 Net Sales 月化"]) == (1000.0, 800.0)
    assert row["Net Sales 變化量(月化基準)"] == -200.0
    assert row["Net Sales 變化率"] == round(-200.0 / 1000.0, 4)
    assert (row["最早 Safety Stock Qty"], row["最新 Safety Stock Qty"]) == (10, 4)
    assert row["Safety Stock Qty 變化量"] == -6


def test_rp_endpoints_zero_base_rate_blank():
    # 最早月份為 0 時變化率留空（與 SKU月度變化同一規則），變化量照計
    change, _m, _s, _c, _sh, _r, stats, _ = _run([
        {"SKU": 1, "Month": 5, "Sold Qty (CM - MTD)": 0.0, "Net Sales (CM - MTD)": 0.0},
        {"SKU": 1, "Month": 9, "Sold Qty (CM - MTD)": 30.0, "Net Sales (CM - MTD)": 300.0},
    ])
    detail = stats["rp_endpoints"]
    row = detail.iloc[0]
    assert row["Sold Qty 變化率"] is None or pd.isna(row["Sold Qty 變化率"])
    assert row["Net Sales 變化率"] is None or pd.isna(row["Net Sales 變化率"])
    assert row["Sold Qty 變化量(月化基準)"] == 30
    assert row["Net Sales 變化量(月化基準)"] == 300.0


def test_rp_sku_counts_in_pivots_and_shop():
    # 樞紐與 By Shop 附 RF→ND SKU數／ND→RF SKU數（整個期間不重複 SKU；總計列為全檔不重複）
    change, mch1, sst, cross, shop, raw, stats, _ = _run([
        # SST1／01 - SKIN CARE：SKU1、SKU2 皆 RF→ND
        {"SKU": 1, "Month": 5, "Adj. SST": "SST1", "Adj. MCH1": "01 - SKIN CARE", "RP Type": "RF"},
        {"SKU": 1, "Month": 6, "Adj. SST": "SST1", "Adj. MCH1": "01 - SKIN CARE", "RP Type": "ND"},
        {"SKU": 2, "Month": 5, "Adj. SST": "SST1", "Adj. MCH1": "01 - SKIN CARE", "RP Type": "RF"},
        {"SKU": 2, "Month": 6, "Adj. SST": "SST1", "Adj. MCH1": "01 - SKIN CARE", "RP Type": "ND"},
        # SST2／02 - MAKE UP：SKU3 ND→RF
        {"SKU": 3, "Month": 5, "Adj. SST": "SST2", "Adj. MCH1": "02 - MAKE UP", "RP Type": "ND"},
        {"SKU": 3, "Month": 6, "Adj. SST": "SST2", "Adj. MCH1": "02 - MAKE UP", "RP Type": "RF"},
        # SST1：SKU4 RF→RF 不計入
        {"SKU": 4, "Month": 5, "Adj. SST": "SST1", "Adj. MCH1": "01 - SKIN CARE", "RP Type": "RF"},
        {"SKU": 4, "Month": 6, "Adj. SST": "SST1", "Adj. MCH1": "01 - SKIN CARE", "RP Type": "RF"},
    ])
    sst1 = sst[sst["Adj. SST"] == "SST1"].iloc[0]
    assert (sst1["RF→ND SKU數"], sst1["ND→RF SKU數"]) == (2, 0)
    sst2 = sst[sst["Adj. SST"] == "SST2"].iloc[0]
    assert (sst2["RF→ND SKU數"], sst2["ND→RF SKU數"]) == (0, 1)
    sst_total = sst[sst["Adj. SST"] == "總計"].iloc[0]
    assert (sst_total["RF→ND SKU數"], sst_total["ND→RF SKU數"]) == (2, 1)
    m01 = mch1[mch1["Adj. MCH1"] == "01 - SKIN CARE"].iloc[0]
    assert (m01["RF→ND SKU數"], m01["ND→RF SKU數"]) == (2, 0)
    cr = cross[(cross["Adj. MCH1"] == "01 - SKIN CARE") & (cross["Adj. SST"] == "SST1")].iloc[0]
    assert (cr["RF→ND SKU數"], cr["ND→RF SKU數"]) == (2, 0)
    cross_total = cross[cross["Adj. MCH1"] == "總計"].iloc[0]
    assert (cross_total["RF→ND SKU數"], cross_total["ND→RF SKU數"]) == (2, 1)
    # By Shop：同店各月份列重複顯示同一值
    assert (shop["RF→ND SKU數"] == 2).all() and (shop["ND→RF SKU數"] == 1).all()


def test_rp_endpoints_empty():
    from transformer import _build_safm_rp_endpoints
    detail, info = _build_safm_rp_endpoints(pd.DataFrame(), [])
    assert detail.empty and info["comparable"] == 0
