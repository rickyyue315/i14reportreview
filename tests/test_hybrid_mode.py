"""Validation tests for Safety Stock x NetStock Review mode (調貨反饋 Safety 混合, v1.10.0).

同步執行 Safety Stock Review + Stock Checking：
- Safety 段與 transform_combined 同公式，調貨結果不修改 Recommend Safety Stock
- 調貨段與 transform_stock_checking 同規則（SKU 級加總 TurnDay、Nature 供應門檻、14 日目標、NT→Mix）
- 反饋：調貨後庫存 = 原始 + 實際調入 − 轉出；調貨前後 Safety 缺口 = max(0, Recommend − 庫存)；調貨後 TurnDay
- Shop Nature：Other/空白/大小寫差異一律正規化為 Non Tourist
"""
import warnings
warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from transformer import (
    transform_hybrid,
    generate_excel_hybrid,
    transform_combined,
    transform_stock_checking,
    validate_hybrid_input,
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1, "SKU Desc": "Test Product",
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Tourist", "Vendor MC": "V1", "Item Status": "A",
    "Adj. SST": "SST1",
    "Launch Date": "2024-01-01", "Item Class": "C1", "RP Type": "RF",
    "Shop Stock Status": "Normal", "Adj. Safety Stock": 3.0,
    "Sold Qty (P90D)": 90.0, "Excess Stock Qty": 2.0,
    "Shop Stock Qty (CM)": 5.0,
    "Shop Monthly Sold Qty (P90D)": 30.5,
    "Sold Qty (CM - MTD)": 8.0,
    "D001 Stock Qty (CM)": 200.0, "D001 Net Stock Qty (CM)": 99.0,
    "Turnday (P90D Qty)": 13.0,
}

_RATES_O = {"default": 0.5, "tourist": 0.3, "mix": 0.4}
_RATES_P = {"default": 0.5, "tourist": 0.5, "mix": 0.5}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


# ──────────────────────────────────────────────────────────────────────────────
# 1. 輸入檢核與基礎呼叫
# ──────────────────────────────────────────────────────────────────────────────
def test_missing_required_column_rejected():
    df = pd.DataFrame([{"SKU": 1, "Shop Stock Status": "Normal"}])
    missing = validate_hybrid_input(df)
    assert "Shop Nature" in missing
    assert "Sold Qty (P90D)" in missing
    try:
        transform_hybrid(df, rates_overstock=_RATES_O, rates_oos=_RATES_P, mtd_days=15, month_total_days=30)
        raise AssertionError("應 raise ValueError")
    except ValueError as e:
        assert "Shop Nature" in str(e)


def test_empty_after_clean_raises():
    df = pd.DataFrame([{"Buyer": np.nan, "Brand": "A", "SKU": 1, "Shop Stock Status": "Normal",
                        "Shop Nature": "Tourist", "Site Code": "S01", "Shop Stock Qty (CM)": 5.0,
                        "Sold Qty (P90D)": 90.0, "Adj. Safety Stock": 3.0,
                        "Shop Monthly Sold Qty (P90D)": 30.5, "Sold Qty (CM - MTD)": 8.0}])
    try:
        transform_hybrid(df, rates_overstock=_RATES_O, rates_oos=_RATES_P, mtd_days=15, month_total_days=30)
        raise AssertionError("應 raise ValueError")
    except ValueError:
        pass


def test_default_args_produce_results():
    df = _df([
        {"SKU": 100, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 100, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, safety_summary, sc_summary, hybrid_summary, stats, warnings = transform_hybrid(df)
    assert stats["need_rows"] == 1
    assert len(warnings) == 0
    assert "調貨後庫存" in raw.columns
    assert not safety_summary.empty
    assert not hybrid_summary.empty


# ──────────────────────────────────────────────────────────────────────────────
# 2. 與既有模式等價（調貨欄位 / Safety 欄位）
# ──────────────────────────────────────────────────────────────────────────────
def test_transfer_columns_equivalent_to_stock_checking():
    df = _df([
        {"SKU": 200, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 720.0, "Shop Stock Qty (CM)": 92.0,
         "Shop Monthly Sold Qty (P90D)": 240.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 200, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 19.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 200, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 22.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    sc_raw, _, sc_stats, _ = transform_stock_checking(df)
    hy_raw, _, _, _, hy_stats, _ = transform_hybrid(df)

    cmp_cols = ["需要調貨", "調入需求量 (補到14日)", "實際調入量", "調入缺口",
                "可轉出貨量 (Mix)", "可轉出貨量 (Non Tourist)", "轉出量", "調貨來源 (Site Code)",
                "Tourist 加總 TurnDay", "Mix 加總 TurnDay", "Non Tourist 加總 TurnDay"]
    for c in cmp_cols:
        assert list(hy_raw[c]) == list(sc_raw[c]), c
    assert hy_stats["nt_transfer"] == sc_stats["nt_transfer"]
    assert hy_stats["mix_transfer"] == sc_stats["mix_transfer"]


def test_safety_columns_equivalent_to_combined():
    df = _df([
        {"SKU": 1, "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 0,
         "Sold Qty (CM - MTD)": 0.0, "Shop Stock Qty (CM)": 3, "Sold Qty (P90D)": 0.0},
        {"SKU": 2, "Shop Stock Status": "Normal", "RP Type": "RF",
         "Adj. Safety Stock": 10, "Shop Monthly Sold Qty (P90D)": 30.5,
         "Sold Qty (CM - MTD)": 8.0, "Shop Stock Qty (CM)": 20, "Sold Qty (P90D)": 90.0},
    ])
    c_sum, c_raw, _, _ = transform_combined(df, rates_overstock=_RATES_O, rates_oos=_RATES_P,
                                             mtd_days=15, month_total_days=30, safety_mode="allow_both")
    hy_raw, hy_safety_sum, _, _, _, _ = transform_hybrid(df, rates_overstock=_RATES_O, rates_oos=_RATES_P,
                                                         mtd_days=15, month_total_days=30, safety_mode="allow_both")
    for c in ["Recommend Safety Stock", "Safety Qty Change", "Overstock Cap (X)", "OOS Floor (Y)",
              "Projected Monthly Sold Qty", "Recommend Safety Stock Turnday",
              "Recommendation of new Safety Qty / ND"]:
        a = hy_raw[c].to_numpy(dtype=object)
        b = c_raw[c].to_numpy(dtype=object)
        if c == "Recommendation of new Safety Qty / ND":
            assert list(a) == list(b), c
        else:
            an = pd.to_numeric(pd.Series(a), errors="coerce").to_numpy()
            bn = pd.to_numeric(pd.Series(b), errors="coerce").to_numpy()
            assert np.allclose(an, bn, equal_nan=True), c
    assert list(hy_safety_sum["總計"]) == list(c_sum["總計"])


# ──────────────────────────────────────────────────────────────────────────────
# 3. 調貨反饋公式（不改寫 Safety，只反映調貨後狀態）
# ──────────────────────────────────────────────────────────────────────────────
def test_feedback_formulas():
    df = _df([
        {"SKU": 300, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 300, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 300, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, _, stats, _ = transform_hybrid(df)
    rec = raw["Recommend Safety Stock"].astype(int)
    cm = raw["Shop Stock Qty (CM)"].astype(int)
    actual = raw["實際調入量"].astype(int)
    transfer = raw["轉出量"].astype(int)
    after = cm + actual - transfer
    assert (raw["調貨後庫存"] == after.values).all()
    expect_before = np.maximum(0, rec.values - cm.values)
    expect_after = np.maximum(0, rec.values - after.values)
    assert (raw["調貨前 Safety 缺口"].values == expect_before).all()
    assert (raw["調貨後 Safety 缺口"].values == expect_after).all()
    assert stats["safety_gap_before"] == int(expect_before.sum())
    assert stats["safety_gap_after"] == int(expect_after.sum())
    assert stats["safety_gap_improved"] == int((expect_before - expect_after).sum())


def test_feedback_does_not_alter_recommendation_or_transfer():
    # 來源店轉出後低於 Safety → 調貨後缺口為正 → 但轉出量與 Recommend 不變（決策：只標記不改寫）
    df = _df([
        {"SKU": 400, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Adj. Safety Stock": 3,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 400, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5,
         "Adj. Safety Stock": 30},
        {"SKU": 400, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
         "Shop Monthly Sold Qty (P90D)": 30.5,
         "Adj. Safety Stock": 30},
    ])
    hy_raw, _, _, _, _, _ = transform_hybrid(df)
    sc_raw, _, _, _ = transform_stock_checking(df)
    nt_row_hy = hy_raw[hy_raw["Site Code"] == "N01"].iloc[0]
    nt_row_sc = sc_raw[sc_raw["Site Code"] == "N01"].iloc[0]
    assert int(nt_row_hy["轉出量"]) == int(nt_row_sc["轉出量"]) == 9
    # NT 店庫存 30 - 9 = 21 < Recommend 30 → 調貨後缺口為正，但 Recommend/轉出不變
    assert int(nt_row_hy["Recommend Safety Stock"]) == 30
    assert int(nt_row_hy["調貨後 Safety 缺口"]) == 9
    assert int(nt_row_hy["調貨後庫存"]) == 21


def test_feedback_turnday_after():
    df = _df([
        {"SKU": 500, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 500, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, _, _, _ = transform_hybrid(df)
    n_row = raw[raw["Site Code"] == "N01"].iloc[0]
    # 調貨後庫存 21 / 月銷 30.5 × 30.5 = 21.0
    assert abs(float(n_row["調貨後 TurnDay"]) - 21.0) < 1e-9
    # 月銷 0 → 留空
    t_row = raw[raw["Site Code"] == "T01"].iloc[0]
    assert pd.notna(t_row["調貨後 TurnDay"])


def test_feedback_turnday_blank_when_no_monthly_sales():
    df = _df([
        {"SKU": 600, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 0.0,
         "Turnday (P90D Qty)": 0.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 600, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, _, _, _ = transform_hybrid(df)
    t_row = raw[raw["Site Code"] == "T01"].iloc[0]
    assert pd.isna(t_row["調貨後 TurnDay"])


# ──────────────────────────────────────────────────────────────────────────────
# 4. Shop Nature 正規化（Other/空白/大小寫 → Non Tourist）
# ──────────────────────────────────────────────────────────────────────────────
def test_other_nature_mapped_to_non_tourist():
    df = _df([
        {"SKU": 1, "Site Code": "O01", "Shop Nature": "Other",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 1, "Site Code": "O02", "Shop Nature": "  non tourist  ",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 10.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 2, "Site Code": "M01", "Shop Nature": "MIX",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 10.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, _, stats, warnings_list = transform_hybrid(df)
    assert (raw.loc[raw["Site Code"] == "O01", "Shop Nature"] == "Non Tourist").all()
    assert (raw.loc[raw["Site Code"] == "O02", "Shop Nature"] == "Non Tourist").all()
    assert (raw.loc[raw["Site Code"] == "M01", "Shop Nature"] == "Mix").all()
    # 只有 Other 屬「無法辨識」；「 non tourist 」與 MIX 是標準值的大小寫/空格變體（計為標準）
    assert stats["mapped_other_rows"] == 1
    assert any("Shop Nature" in w for w in warnings_list)


def test_uppercase_tourist_normalized():
    df = _df([
        {"SKU": 1, "Site Code": "T01", "Shop Nature": "TOURIST",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 1, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, _, stats, _ = transform_hybrid(df)
    assert (raw["Shop Nature"] == "Tourist").any()
    assert stats["need_rows"] == 1


def test_blank_nature_mapped_to_non_tourist():
    df = _df([
        {"SKU": 1, "Site Code": "X01", "Shop Nature": np.nan,
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, _, stats, warnings_list = transform_hybrid(df)
    assert (raw["Shop Nature"] == "Non Tourist").all()
    assert stats["mapped_other_rows"] == 1
    assert any("Shop Nature" in w for w in warnings_list)


# ──────────────────────────────────────────────────────────────────────────────
# 5. 觸發與 D001 fallback
# ──────────────────────────────────────────────────────────────────────────────
def test_trigger_and_aggregate_turnday():
    df = _df([
        {"SKU": 700, "Site Code": "TA", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 10.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 700, "Site Code": "TB", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 20.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Turnday (P90D Qty)": 20.0, "D001 Net Stock Qty (CM)": 10.0},
    ])
    raw, _, _, _, stats, _ = transform_hybrid(df)
    # 加總 TurnDay = 30/61×30.5 = 15 ≥ 14 → 無須調貨
    assert stats["need_rows"] == 0
    assert (raw["需要調貨"] == "").all()


def test_d001_column_fallback():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Turnday (P90D Qty)": 13.0, "D001 Stock Qty (CM)": 50.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)"])
    raw, _, _, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 1


def test_missing_d001_warns():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Turnday (P90D Qty)": 13.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)"])
    raw, _, _, _, stats, warnings_list = transform_hybrid(df)
    assert stats["need_rows"] == 1
    assert any("D001" in w for w in warnings_list)


# ──────────────────────────────────────────────────────────────────────────────
# 6. Hybrid Summary / Safety Summary / stats
# ──────────────────────────────────────────────────────────────────────────────
def test_hybrid_summary_rows_and_improvement():
    df = _df([
        {"SKU": 800, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Adj. Safety Stock": 3,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 800, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5,
         "Adj. Safety Stock": 30},
        {"SKU": 801, "Site Code": "T02", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 0.0,
         "Adj. Safety Stock": 3,
         "Turnday (P90D Qty)": 0.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 801, "Site Code": "N02", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5,
         "Adj. Safety Stock": 30},
    ])
    raw, safety_summary, sc_summary, hybrid_summary, stats, _ = transform_hybrid(df)
    assert len(hybrid_summary) == 2
    row800 = hybrid_summary[hybrid_summary["SKU"] == "800"].iloc[0]
    assert int(row800["需調貨 Tourist 列數"]) == 1
    assert int(row800["實際調入量"]) == 9
    # 調貨前後庫存：T01 47 → 56；N01 30 → 21
    assert int(row800["調貨前庫存"]) == 77
    assert int(row800["調貨後庫存"]) == 77
    # NT 店 Adj. Safety Stock=30（increase_only 下 Recommend 仍 30），轉出 9 → 調貨後庫存 21 → 缺口 +9
    assert int(row800["改善件數"]) == -9
    assert not safety_summary.empty
    assert (safety_summary["Buyer"] == "總計").any()
    assert (safety_summary["總計"] >= 0).all()
    assert stats["need_rows"] == 1
    assert stats["need_skus"] == 1
    assert stats["gap_skus_after"] >= 0


def test_no_candidates_warns_but_completes():
    df = _df([
        {"SKU": 1, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, safety_summary, sc_summary, hybrid_summary, stats, warnings_list = transform_hybrid(df)
    assert stats["need_rows"] == 0
    assert any("沒有符合調貨條件" in w for w in warnings_list)
    buf = generate_excel_hybrid(raw, safety_summary, sc_summary, hybrid_summary, stats)
    assert isinstance(buf, BytesIO)


# ──────────────────────────────────────────────────────────────────────────────
# 7. Excel 契約（8 工作表）
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_excel_sheets_and_headers():
    df = _df([
        {"SKU": 100, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 100, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, safety_summary, sc_summary, hybrid_summary, stats, _ = transform_hybrid(df)
    buf = generate_excel_hybrid(raw, safety_summary, sc_summary, hybrid_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == [
        "RAW", "Shop Stock Status", "By Shop", "SUMMARY",
        "Safety Stock 數據分析摘要", "調貨 Summary", "調貨分析摘要",
        "Hybrid Summary（調貨反饋 Safety）",
    ]
    ws_raw = wb["RAW"]
    headers = [c.value for c in next(ws_raw.iter_rows(min_row=1, max_row=1))]
    for col in ["Recommend Safety Stock", "需要調貨", "調貨後庫存", "調貨前 Safety 缺口",
                "調貨後 Safety 缺口", "調貨後 TurnDay"]:
        assert col in headers
    ws_hyb = wb["Hybrid Summary（調貨反饋 Safety）"]
    hyb_headers = [c.value for c in next(ws_hyb.iter_rows(min_row=1, max_row=1))]
    for col in ["SKU", "調入需求量", "調貨前 Safety 缺口", "調貨後 Safety 缺口", "改善件數", "仍有缺口"]:
        assert col in hyb_headers


def test_generate_excel_dashboard_matches_stats():
    df = _df([
        {"SKU": 100, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 100, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, safety_summary, sc_summary, hybrid_summary, stats, _ = transform_hybrid(df)
    buf = generate_excel_hybrid(raw, safety_summary, sc_summary, hybrid_summary, stats)
    wb = load_workbook(buf, read_only=True)
    ws = wb["調貨分析摘要"]
    vals = {
        ws.cell(row=3, column=c).value: ws.cell(row=4, column=c).value
        for c in range(1, 5)
    } | {
        ws.cell(row=7, column=c).value: ws.cell(row=8, column=c).value
        for c in range(1, 5)
    }
    assert vals["需要調貨列數"] == stats["need_rows"]
    assert vals["需調貨 SKU 數"] == stats["need_skus"]
    assert vals["調入需求量"] == stats["total_demand"]
    assert vals["實際調入量"] == stats["total_actual"]
    assert vals["Non Tourist 轉出量"] == stats["nt_transfer"]
    assert vals["Mix 轉出量"] == stats["mix_transfer"]
    assert vals["無法足額缺口"] == stats["total_gap"]
    assert vals["總明細行數"] == len(raw)


def test_excel_hybrid_sheet_freeze_and_autofilter():
    df = _df([
        {"SKU": 100, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 100, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, safety_summary, sc_summary, hybrid_summary, stats, _ = transform_hybrid(df)
    buf = generate_excel_hybrid(raw, safety_summary, sc_summary, hybrid_summary, stats)
    wb = load_workbook(buf)
    ws = wb["Hybrid Summary（調貨反饋 Safety）"]
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref is not None


# ──────────────────────────────────────────────────────────────────────────────
# 8. Review 修正回歸：header canonicalization / 逐列 fallback / 數值品質
# ──────────────────────────────────────────────────────────────────────────────
def test_header_whitespace_canonicalized():
    """帶前後空格的必要欄位：先驗證通過、後續也能正確運算（不再 KeyError）。"""
    df = _df([
        {"SKU": 100, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 100, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    df = df.rename(columns={c: c + " " for c in df.columns})
    assert validate_hybrid_input(df) == []
    raw, _, _, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 1
    assert "調貨後庫存" in raw.columns


def test_duplicate_canonical_header_rejected():
    df = _df([{"SKU": 1}])
    df["Buyer "] = df["Buyer"]
    try:
        transform_hybrid(df, rates_overstock=_RATES_O, rates_oos=_RATES_P, mtd_days=15, month_total_days=30)
        raise AssertionError("應 raise ValueError")
    except ValueError as e:
        assert "重複" in str(e) or "Buyer" in str(e)


def test_d001_row_level_fallback():
    """D001 Net 單列空白 → 逐列以 D001 Stock 回補，不再誤觸發調貨。"""
    df = _df([
        {"SKU": 1, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5,
         "D001 Net Stock Qty (CM)": None, "D001 Stock Qty (CM)": 200.0},
        {"SKU": 1, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, _, stats, warnings_list = transform_hybrid(df)
    assert stats["need_rows"] == 0
    t_row = raw[raw["Site Code"] == "T01"].iloc[0]
    assert int(t_row["D001 Effective Stock Qty (CM)"]) == 200
    assert int(t_row["D001 Stock Qty (CM)"]) == 200
    assert any("D001 Net Stock Qty (CM)" in w and "空白" in w for w in warnings_list)


def test_monthly_blank_row_level_fallback():
    """Shop Monthly Sold Qty (P90D) 單列空白 → 以 Sold Qty (P90D) 換算，不誤判 inf。"""
    df = _df([
        {"SKU": 2, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": None,
         "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 2, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, _, stats, warnings_list = transform_hybrid(df)
    assert stats["need_rows"] == 1  # 加總 TurnDay = 47/122×30.5 ≈ 11.8 < 14 → 正確觸發（修復前會誤判 inf 不觸發）
    t_row = raw[raw["Site Code"] == "T01"].iloc[0]
    assert abs(float(t_row["Shop Monthly Sold Qty (P90D)"]) - 122.0) < 1e-9
    assert any("空白" in w and "換算" in w for w in warnings_list)


def test_non_integer_stock_rejected():
    """Shop Stock Qty (CM) 出現小數 → 明確拒絕，避免取整後守恆失真。"""
    df = _df([
        {"SKU": 1, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 0.5,
         "Shop Monthly Sold Qty (P90D)": 30.5, "D001 Net Stock Qty (CM)": 10.0},
    ])
    try:
        transform_hybrid(df, rates_overstock=_RATES_O, rates_oos=_RATES_P, mtd_days=15, month_total_days=30)
        raise AssertionError("應 raise ValueError")
    except ValueError as e:
        assert "整數" in str(e)


def test_sku_desc_all_blank_derived():
    """SKU Desc 欄位存在但全欄空白 → 仍以 Item_Shop 補正，摘要不塌陷。"""
    df = _df([
        {"SKU": 1, "Shop Nature": "Tourist", "Site Code": "T01",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0, "D001 Net Stock Qty (CM)": 10.0},
    ])
    df["SKU Desc"] = None
    raw, safety_summary, _, _, _, _ = transform_hybrid(df)
    assert len(safety_summary) >= 2  # 1 分類列 + 總計列
    assert (safety_summary["Buyer"] == "總計").any()


def test_hybrid_excel_progress_monotonic():
    """Excel 生成全程 progress 不得倒退（修復 _write_stock_checking_summary_sheet 固定區間）。"""
    df = _df([
        {"SKU": 100, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 100, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, safety_summary, sc_summary, hybrid_summary, stats, _ = transform_hybrid(df)
    events = []
    generate_excel_hybrid(raw, safety_summary, sc_summary, hybrid_summary, stats,
                          progress_callback=lambda p, _m: events.append(float(p)))
    assert all(b >= a for a, b in zip(events, events[1:])), events


def test_raw_headers_include_effective_d001():
    """RAW 需含 D001 Effective 稽核欄位（fallback 時可追查實際用值）。"""
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5,
         "D001 Net Stock Qty (CM)": 10.0},
    ])
    raw, safety_summary, sc_summary, hybrid_summary, stats, _ = transform_hybrid(df)
    buf = generate_excel_hybrid(raw, safety_summary, sc_summary, hybrid_summary, stats)
    wb = load_workbook(buf, read_only=True)
    headers = [c.value for c in next(wb["RAW"].iter_rows(min_row=1, max_row=1))]
    assert "D001 Effective Stock Qty (CM)" in headers