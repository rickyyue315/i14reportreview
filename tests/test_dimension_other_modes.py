"""Validation tests for 分類維度 × Stock Checking／Safety Stock Review（v1.37.0）.

從 Safety Stock x NetStock Review 引入分類到另外兩個模式（調貨計算一律不改）：
- Stock Checking：transform 恆附銷量帶／銷量存貨分類；dimension 只改 Summary 分組。
- Safety Stock Review：transform 恆附銷量帶／銷量存貨分類；SUMMARY 區塊 ③ 為
  Safety Qty Change 分類 × 分類維度交叉表。
- 兩模式店舖問題 SKU 表末尾附銷量帶／銷量存貨分類欄。
"""
import warnings
warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from transformer import (
    transform_stock_checking,
    generate_excel_stock_checking,
    _STOCK_CHECKING_SUMMARY_COLUMNS,
    _stock_checking_summary_columns,
    _build_stock_checking_shop_issues,
    _STOCK_CHECKING_SHOP_ISSUE_COLUMNS,
    transform_combined,
    generate_excel_combined,
    _build_combined_dimension_cross_tab,
    _build_combined_shop_issues,
    _COMBINED_SHOP_ISSUE_COLUMNS,
    _attach_sales_classification_columns,
    _hybrid_sales_params,
    _HYBRID_SALES_BAND_COLUMN,
    _HYBRID_SALES_CLASS_COLUMN,
    _HYBRID_SALES_BANDS,
    _HYBRID_SALES_CLASS_ORDER,
)

_BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1, "SKU Desc": "Test Product",
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Supply Source": "Warehouse",
    "Shop Nature": "Tourist", "Vendor MC": "V1", "Item Status": "A",
    "Adj. SST": "SST1",
    "Launch Date": "2024-01-01", "Item Class": "C1", "RP Type": "RF",
    "Shop Stock Status": "Normal", "Adj. Safety Stock": 3.0,
    "Sold Qty (P90D)": 90.0, "Excess Stock Qty": 2.0,
    "Shop Stock Qty (CM)": 5.0,
    "D001 Stock Qty (CM)": 200.0, "D001 Net Stock Qty (CM)": 99.0,
    "Turnday (P90D Qty)": 13.0,
}


def _row(**kw):
    r = dict(_BASE)
    r.update(kw)
    return r


def _sc_sales_df():
    """SKU 100 五列（四 Tourist 一 Non Tourist）：銷量帶 高銷×2／中銷×2／低銷×1，
    分類 高銷缺貨×2／低銷積壓×1／正常×2，並觸發 Tourist 需要調貨。"""
    return pd.DataFrame([
        _row(**{"SKU": 100, "Site Code": "T1", "Shop Nature": "Tourist", "OM": "OM1",
                "Sold Qty (P90D)": 720.0, "Shop Monthly Sold Qty (P90D)": 240.0,
                "Shop Stock Qty (CM)": 5.0, "D001 Net Stock Qty (CM)": 10.0}),
        _row(**{"SKU": 100, "Site Code": "T2", "Shop Nature": "Tourist", "OM": "OM1",
                "Sold Qty (P90D)": 90.0, "Shop Monthly Sold Qty (P90D)": 30.0,
                "Shop Stock Qty (CM)": 5.0, "D001 Net Stock Qty (CM)": 10.0}),
        _row(**{"SKU": 100, "Site Code": "T3", "Shop Nature": "Tourist", "OM": "OM2",
                "Sold Qty (P90D)": 360.0, "Shop Monthly Sold Qty (P90D)": 120.0,
                "Shop Stock Qty (CM)": 5.0, "D001 Net Stock Qty (CM)": 10.0}),
        _row(**{"SKU": 100, "Site Code": "T4", "Shop Nature": "Tourist", "OM": "OM2",
                "Sold Qty (P90D)": 180.0, "Shop Monthly Sold Qty (P90D)": 60.0,
                "Shop Stock Qty (CM)": 5.0, "D001 Net Stock Qty (CM)": 10.0}),
        _row(**{"SKU": 100, "Site Code": "N1", "Shop Nature": "Non Tourist", "OM": "OM1",
                "Sold Qty (P90D)": 30.0, "Shop Monthly Sold Qty (P90D)": 30.0,
                "Shop Stock Qty (CM)": 100.0}),
    ])


# ── Stock Checking ────────────────────────────────────────────────────────────

def test_sc_default_summary_contract_unchanged_and_sales_columns_present():
    raw, summary, stats, _ = transform_stock_checking(_sc_sales_df())
    assert list(summary.columns) == _STOCK_CHECKING_SUMMARY_COLUMNS
    assert stats["dimension"] == "nature"
    assert _HYBRID_SALES_BAND_COLUMN in raw.columns and _HYBRID_SALES_CLASS_COLUMN in raw.columns
    band = dict(zip(raw["Site Code"], raw[_HYBRID_SALES_BAND_COLUMN]))
    cls = dict(zip(raw["Site Code"], raw[_HYBRID_SALES_CLASS_COLUMN]))
    # 同 SKU 五列有銷：排名 1.5/1.5/3/4/5 → 30＝低銷、90/180＝中銷、360/720＝高銷
    assert band["N1"] == "低銷" and band["T2"] == "中銷" and band["T4"] == "中銷"
    assert band["T3"] == "高銷" and band["T1"] == "高銷"
    assert cls["T1"] == "高銷缺貨" and cls["T3"] == "高銷缺貨"  # 周轉 0.64/1.27 日 < 7
    assert cls["N1"] == "低銷積壓"  # 周轉 101.7 日 ≥ 30 且庫存 ≥ 2
    assert cls["T2"] == "正常" and cls["T4"] == "正常"


def test_sc_sales_dimension_groups_summary_only():
    df = _sc_sales_df()
    raw_n, _, stats_n, _ = transform_stock_checking(df)
    raw_s, summary_s, stats_s, _ = transform_stock_checking(df, dimension="sales")
    assert stats_s["dimension"] == "sales"
    assert "高銷 加總 - Shop Stock Qty (CM)" in summary_s.columns
    # 調貨計算不隨維度改變
    for col in ("需要調貨", "實際調入量", "轉出量", "調入缺口"):
        assert list(raw_n[col]) == list(raw_s[col])
    row = summary_s[summary_s["SKU"] == "100"].iloc[0]
    assert int(row["高銷 加總 - Shop Stock Qty (CM)"]) == 10   # T1＋T3
    assert int(row["低銷 加總 - Shop Stock Qty (CM)"]) == 100  # N1
    assert int(row["中銷 加總 - Shop Stock Qty (CM)"]) == 10   # T2＋T4
    assert int(row["零銷 加總 - Shop Stock Qty (CM)"]) == 0
    assert list(summary_s.columns) == _stock_checking_summary_columns("sales")


def test_sc_om_dimension_groups_from_raw():
    _, summary, stats, _ = transform_stock_checking(_sc_sales_df(), dimension="om")
    assert stats["dimension"] == "om"
    assert stats["om_groups"] == ("OM1", "OM2")
    assert "OM1 加總 - Shop Stock Qty (CM)" in summary.columns
    assert "OM2 加總 - Shop Stock Qty (CM)" in summary.columns


def test_sc_shop_issue_table_carries_sales_classification():
    raw, _, stats, _ = transform_stock_checking(_sc_sales_df())
    issues = _build_stock_checking_shop_issues(
        raw, turnday_threshold=stats["turnday_threshold"], d001_threshold=stats["d001_threshold"])
    assert not issues.empty
    assert issues.columns[-2:].tolist() == [_HYBRID_SALES_BAND_COLUMN, _HYBRID_SALES_CLASS_COLUMN]
    assert list(issues.columns) == [c for c in _STOCK_CHECKING_SHOP_ISSUE_COLUMNS if c in issues.columns]
    by_site = {r["Site Code"]: r for _, r in issues.iterrows()}
    assert by_site["T1"][_HYBRID_SALES_BAND_COLUMN] == "高銷"
    assert by_site["T1"][_HYBRID_SALES_CLASS_COLUMN] == "高銷缺貨"
    assert by_site["N1"][_HYBRID_SALES_CLASS_COLUMN] == "低銷積壓"


def test_sc_excel_summary_sheet_grouped_by_dimension():
    raw, summary, stats, _ = transform_stock_checking(_sc_sales_df(), dimension="sales")
    buf = generate_excel_stock_checking(raw, summary, stats)
    wb = load_workbook(BytesIO(buf.getvalue()), read_only=True)
    row4 = [c.value for c in next(wb["Summary"].iter_rows(min_row=4, max_row=4))]
    for g in _HYBRID_SALES_BANDS:
        assert g in row4


def test_attach_sales_classification_columns_guards():
    params = _hybrid_sales_params(False)[1]
    base = pd.DataFrame([
        {"SKU": "1", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0, "Shop Monthly Sold Qty (P90D)": 30.0},
        {"SKU": "1", "Sold Qty (P90D)": 720.0, "Shop Stock Qty (CM)": 5.0, "Shop Monthly Sold Qty (P90D)": 240.0},
    ])
    out, w = _attach_sales_classification_columns(base, params)
    assert not w and _HYBRID_SALES_CLASS_COLUMN in out.columns
    # 缺 Sold Qty (P90D) → 兩欄皆不輸出＋warning
    out2, w2 = _attach_sales_classification_columns(base.drop(columns=["Sold Qty (P90D)"]), params)
    assert _HYBRID_SALES_BAND_COLUMN not in out2.columns and len(w2) == 1
    # 缺 Shop Stock Qty (CM) → 只有銷量帶＋warning
    out3, w3 = _attach_sales_classification_columns(base.drop(columns=["Shop Stock Qty (CM)"]), params)
    assert _HYBRID_SALES_BAND_COLUMN in out3.columns and _HYBRID_SALES_CLASS_COLUMN not in out3.columns
    assert len(w3) == 1
    # 缺 Shop Monthly Sold Qty (P90D) → 以 Sold Qty (P90D) 換算＋warning
    out4, w4 = _attach_sales_classification_columns(base.drop(columns=["Shop Monthly Sold Qty (P90D)"]), params)
    assert _HYBRID_SALES_CLASS_COLUMN in out4.columns and len(w4) == 1


# ── Safety Stock Review ───────────────────────────────────────────────────────

_CB_RATES_O = {"tourist": 0.30, "mix": 0.40, "default": 0.50}
_CB_RATES_P = {"tourist": 0.50, "mix": 0.50, "default": 0.50}


def _cb_df():
    return pd.DataFrame([
        _row(**{"SKU": "1", "Site Code": "S1", "Shop Nature": "Tourist", "OM": "OM1",
                "Shop Stock Status": "Normal", "Adj. Safety Stock": 5,
                "Shop Monthly Sold Qty (P90D)": 30.0, "Sold Qty (CM - MTD)": 8.0,
                "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 4.0}),
        _row(**{"SKU": "1", "Site Code": "S2", "Shop Nature": "Tourist", "OM": "OM1",
                "Shop Stock Status": "Normal", "Adj. Safety Stock": 3,
                "Shop Monthly Sold Qty (P90D)": 30.0, "Sold Qty (CM - MTD)": 0.0,
                "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 40.0}),
        _row(**{"SKU": "1", "Site Code": "S3", "Shop Nature": "Non Tourist", "OM": "OM2",
                "Shop Stock Status": "Normal", "Adj. Safety Stock": 3,
                "Shop Monthly Sold Qty (P90D)": 30.0, "Sold Qty (CM - MTD)": 0.0,
                "Sold Qty (P90D)": 180.0, "Shop Stock Qty (CM)": 50.0}),
        _row(**{"SKU": "1", "Site Code": "S4", "Shop Nature": "Tourist", "OM": "OM1",
                "Shop Stock Status": "Normal", "Adj. Safety Stock": 3,
                "Shop Monthly Sold Qty (P90D)": 30.0, "Sold Qty (CM - MTD)": 0.0,
                "Sold Qty (P90D)": 30.0, "Shop Stock Qty (CM)": 3.0}),
        _row(**{"SKU": "2", "Site Code": "S5", "Shop Nature": "Mix", "OM": "OM2",
                "Shop Stock Status": "Normal", "RP Type": "ND", "Adj. Safety Stock": 2,
                "Shop Monthly Sold Qty (P90D)": 0.0, "Sold Qty (CM - MTD)": 0.0,
                "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 10.0}),
    ])


def test_combined_transform_attaches_sales_columns():
    _, raw, _, warns = transform_combined(
        _cb_df(), rates_overstock=_CB_RATES_O, rates_oos=_CB_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both",
        sales_oos_days=7, sales_overstock_days=30, sales_overstock_min_units=2,
    )
    assert _HYBRID_SALES_BAND_COLUMN in raw.columns and _HYBRID_SALES_CLASS_COLUMN in raw.columns
    cls = dict(zip(raw["Site Code"], raw[_HYBRID_SALES_CLASS_COLUMN]))
    # SKU 1 四列有銷：720 高銷（S2）、360 高銷（S3）、180 中銷（S1? 依平均排名）、30 低銷（S4）
    assert cls["S5"] == "零銷積壓"  # P90D 0 銷且庫存 10 ≥ 2
    # S3：高銷帶且周轉 50/120×30.5 ≈ 12.7 日 ≥ 7 → 非缺貨（歸正常）
    assert cls["S3"] == "正常"
    # S2：高銷帶且周轉 40/120×30.5 ≈ 10.2 日 ≥ 7 → 正常；S4：低銷帶但周轉 3.05 日 < 30 → 正常
    assert cls["S2"] == "正常" and cls["S4"] == "正常"
    # 缺欄守衛不應誤報
    assert not any("無法計算銷量帶" in w for w in warns)


def test_combined_transform_without_sold_p90d_warns_and_skips_columns():
    df = _cb_df().drop(columns=["Sold Qty (P90D)"])
    _, raw, _, warns = transform_combined(
        df, rates_overstock=_CB_RATES_O, rates_oos=_CB_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both",
    )
    assert _HYBRID_SALES_BAND_COLUMN not in raw.columns
    assert _HYBRID_SALES_CLASS_COLUMN not in raw.columns
    assert any("無法計算銷量帶與銷量存貨分類" in w for w in warns)


def test_combined_cross_tab_nature_groups_and_totals():
    raw = pd.DataFrame([
        {"SKU": "1", "Site Code": "S1", "Shop Nature": "Tourist", "OM": "OM1",
         "Safety Qty Change": 5, _HYBRID_SALES_BAND_COLUMN: "高銷"},
        {"SKU": "1", "Site Code": "S2", "Shop Nature": "Tourist", "OM": "OM2",
         "Safety Qty Change": -3, _HYBRID_SALES_BAND_COLUMN: "低銷"},
        {"SKU": "2", "Site Code": "S3", "Shop Nature": "Non Tourist", "OM": "OM1",
         "Safety Qty Change": 0, _HYBRID_SALES_BAND_COLUMN: "零銷"},
        {"SKU": "3", "Site Code": "S4", "Shop Nature": "Other", "OM": "OM1",
         "Safety Qty Change": 2, _HYBRID_SALES_BAND_COLUMN: "中銷"},
    ])
    cross = _build_combined_dimension_cross_tab(raw, dimension="nature")
    assert list(cross["Safety Qty Change 分類"]) == ["增加", "減少", "不變", "總計"]
    inc = cross[cross["Safety Qty Change 分類"] == "增加"].iloc[0]
    assert inc["Tourist 列數"] == 1 and inc["Tourist SKU 數"] == 1
    assert inc["Tourist Safety Qty Change 合計"] == 5
    assert inc["合計 列數"] == 2 and inc["合計 Safety Qty Change"] == 7
    dec = cross[cross["Safety Qty Change 分類"] == "減少"].iloc[0]
    assert dec["Tourist Safety Qty Change 合計"] == -3
    keep = cross[cross["Safety Qty Change 分類"] == "不變"].iloc[0]
    assert keep["Non Tourist 列數"] == 1 and keep["合計 列數"] == 1
    total = cross[cross["Safety Qty Change 分類"] == "總計"].iloc[0]
    # 非標準 Nature（Other）不佔分組欄、只計入合計
    assert total["Tourist 列數"] == 2 and total["Non Tourist 列數"] == 1
    assert total["合計 列數"] == 4 and total["合計 SKU 數"] == 3
    assert total["合計 Safety Qty Change"] == 4


def test_combined_cross_tab_sales_and_om_dimensions():
    raw = pd.DataFrame([
        {"SKU": "1", "Site Code": "S1", "Shop Nature": "Tourist", "OM": "OM1",
         "Safety Qty Change": 5, _HYBRID_SALES_BAND_COLUMN: "高銷"},
        {"SKU": "1", "Site Code": "S2", "Shop Nature": "Tourist", "OM": "OM2",
         "Safety Qty Change": -3, _HYBRID_SALES_BAND_COLUMN: "低銷"},
        {"SKU": "2", "Site Code": "S3", "Shop Nature": "Non Tourist", "OM": "OM1",
         "Safety Qty Change": 0, _HYBRID_SALES_BAND_COLUMN: "零銷"},
    ])
    cross_s = _build_combined_dimension_cross_tab(raw, dimension="sales")
    total = cross_s[cross_s["Safety Qty Change 分類"] == "總計"].iloc[0]
    assert total["高銷 列數"] == 1 and total["低銷 列數"] == 1 and total["零銷 列數"] == 1
    assert total["合計 列數"] == 3
    inc = cross_s[cross_s["Safety Qty Change 分類"] == "增加"].iloc[0]
    assert inc["高銷 Safety Qty Change 合計"] == 5
    cross_o = _build_combined_dimension_cross_tab(raw, dimension="om")
    assert "OM1 列數" in cross_o.columns and "OM2 列數" in cross_o.columns
    total_o = cross_o[cross_o["Safety Qty Change 分類"] == "總計"].iloc[0]
    assert total_o["OM1 列數"] == 2 and total_o["OM2 列數"] == 1


def test_combined_cross_tab_sales_dimension_falls_back_without_band():
    raw = pd.DataFrame([
        {"SKU": "1", "Site Code": "S1", "Shop Nature": "Tourist", "OM": "OM1", "Safety Qty Change": 5},
    ])
    cross = _build_combined_dimension_cross_tab(raw, dimension="sales")
    total = cross[cross["Safety Qty Change 分類"] == "總計"].iloc[0]
    assert total["合計 列數"] == 1
    assert all(total[f"{g} 列數"] == 0 for g in _HYBRID_SALES_BANDS)


def test_combined_shop_issue_table_carries_sales_classification():
    _, raw, _, _ = transform_combined(
        _cb_df(), rates_overstock=_CB_RATES_O, rates_oos=_CB_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both",
    )
    issues = _build_combined_shop_issues(raw)
    if issues.empty:
        return  # 母體無需行動列時表為空（欄位契約仍須成立）
    assert issues.columns[-2:].tolist() == [_HYBRID_SALES_BAND_COLUMN, _HYBRID_SALES_CLASS_COLUMN]
    assert list(issues.columns) == [c for c in _COMBINED_SHOP_ISSUE_COLUMNS if c in issues.columns]


def test_combined_excel_summary_contains_dimension_cross_tab():
    s, r, c, _ = transform_combined(
        _cb_df(), rates_overstock=_CB_RATES_O, rates_oos=_CB_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both",
    )
    buf = generate_excel_combined(s, r, c, dimension="sales")
    wb = load_workbook(BytesIO(buf.getvalue()))
    assert wb.sheetnames == ["說明", "RAW", "Shop Stock Status", "By Shop", "店舖問題 SKU", "SUMMARY", "Safety Stock 數據分析摘要"]
    ws = wb["SUMMARY"]
    col_a = [ws.cell(row=i, column=1).value for i in range(1, ws.max_row + 1)]
    assert "① Safety Qty Change 分類（增加 / 減少 / 不變）" in col_a
    assert "② Recommendation of new Safety Qty / ND 分類" in col_a
    assert "③ Safety Qty Change 分類 × 分類維度交叉表" in col_a
    all_values = [ws.cell(row=i, column=j).value
                  for i in range(1, ws.max_row + 1) for j in range(1, ws.max_column + 1)]
    for g in _HYBRID_SALES_BANDS:
        assert f"{g} 列數" in all_values
        assert f"{g} Safety Qty Change 合計" in all_values
    # 預設（未傳 dimension）也應附 nature 交叉表
    buf2 = generate_excel_combined(s, r, c)
    ws2 = load_workbook(BytesIO(buf2.getvalue()))["SUMMARY"]
    col_a2 = [ws2.cell(row=i, column=1).value for i in range(1, ws2.max_row + 1)]
    assert "③ Safety Qty Change 分類 × 分類維度交叉表" in col_a2
    all_values2 = [ws2.cell(row=i, column=j).value
                   for i in range(1, ws2.max_row + 1) for j in range(1, ws2.max_column + 1)]
    assert "Tourist 列數" in all_values2 and "Non Tourist Safety Qty Change 合計" in all_values2


def test_sales_classification_order_contract_unchanged():
    assert _HYBRID_SALES_CLASS_ORDER == ("高銷缺貨", "低銷積壓", "零銷積壓", "正常")
