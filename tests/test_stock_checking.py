"""Validation tests for Stock Checking mode (調貨檢查, v1.9.0).

D001 倉庫少貨（低於門檻）且遊客區（Tourist）店舖缺貨時的檢查：
觸發：Tourist 店 SKU，Turnday (P90D Qty) ∈ [0, 門檻) 且 D001 Net Stock Qty (CM) < 門檻 且補到 14 日周轉的需求 > 0。
需求：調入需求量 = ceil(P90D / 90 × 14) − Shop Stock Qty (CM)。
供應：Mix / Non Tourist 來源店可轉出 = max(0, Shop Stock Qty (CM) − P90D / 90 × 14)；
無銷售但有存貨全數可釋放；RF 與 ND 店皆可。
分配：按 SKU，Non Tourist 優先 → Mix 補足；需求量大→小依序分配。
輸出：RAW 單一工作表（原始欄位 + 調貨計算欄位）。
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
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1, "SKU Desc": "Test Product",
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Tourist", "Vendor MC": "V1", "Item Status": "A",
    "Launch Date": "2024-01-01", "Item Class": "C1", "RP Type": "RF",
    "Shop Stock Status": "Normal", "Adj. Safety Stock": 3.0,
    "Sold Qty (P90D)": 90.0, "Excess Stock Qty": 2.0,
    "Shop Stock Qty (CM)": 5.0,
    "D001 Stock Qty (CM)": 200.0, "D001 Net Stock Qty (CM)": 99.0,
    "Turnday (P90D Qty)": 13.0,
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


# ──────────────────────────────────────────────────────────────────────────────
# 1. 觸發條件
# ──────────────────────────────────────────────────────────────────────────────
def test_trigger_condition():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},  # 符合
        {"SKU": 2, "Site Code": "S02", "Turnday (P90D Qty)": 14.0, "D001 Net Stock Qty (CM)": 99.0},  # Turnday = 門檻 → 不觸發
        {"SKU": 3, "Site Code": "S03", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 100.0},  # D001 = 門檻 → 不觸發
        {"SKU": 4, "Site Code": "S04", "Shop Nature": "Mix"},  # 非 Tourist → 不觸發
        {"SKU": 5, "Site Code": "S05", "Shop Nature": "Non Tourist"},  # 非 Tourist → 不觸發
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert list(raw.loc[raw["需要調貨"] == "是", "Site Code"]) == ["S01"]
    assert list(raw.loc[raw["需要調貨"] == "是", "SKU"]) == ["1"]
    assert stats["need_rows"] == 1
    assert stats["need_skus"] == 1


def test_turnday_invalid_flag_not_trigger():
    # Turnday = -1 表示 P90D=0（無銷售），即使 D001 低也不觸發
    df = _df([
        {"SKU": 1, "Sold Qty (P90D)": 0.0, "Turnday (P90D Qty)": 0.0,
         "Shop Stock Qty (CM)": 5.0, "D001 Net Stock Qty (CM)": 50.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["need_rows"] == 0
    assert (raw["需要調貨"] == "").all()


def test_demand_positive_required():
    # 庫存已達 14 日理想存量 → 需求 = 0 → 不觸發
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 99.0,
         "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["need_rows"] == 0


# ──────────────────────────────────────────────────────────────────────────────
# 2. 調入需求量（補到 14 日周轉）
# ──────────────────────────────────────────────────────────────────────────────
def test_demand_calculation():
    # P90D=90 → 日均 1 → 14 日理想 = 14；庫存 5 → 需求 = 14 − 5 = 9
    df = _df([
        {"SKU": 1, "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    row = raw.iloc[0]
    assert int(row["調入需求量 (補到14日)"]) == 9
    assert stats["total_demand"] == 9


def test_demand_ceil():
    # P90D=89 → 日均 0.9889 → 14 日理想 = 13.844 → ceil = 14；庫存 0 → 需求 = 14
    df = _df([
        {"SKU": 1, "Sold Qty (P90D)": 89.0, "Shop Stock Qty (CM)": 0.0,
         "Turnday (P90D Qty)": 5.0, "D001 Net Stock Qty (CM)": 10.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert int(raw.iloc[0]["調入需求量 (補到14日)"]) == 14


# ──────────────────────────────────────────────────────────────────────────────
# 3. 可轉出貨量與分配（Non Tourist 優先 → Mix 補足）
# ──────────────────────────────────────────────────────────────────────────────
def test_nt_priority_allocation():
    # Tourist 需求 9（P90D=360 日均 4 → 14 日理想 56，庫存 47；D001=10）
    # NT 有充足可轉出（P90D=90 日均 1 → 理想 14，庫存 30 → 可轉出 16）
    # Mix 也有可轉出（理想 14，庫存 25 → 可轉出 11）
    df = _df([
        {"SKU": 100, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 100, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0},
        {"SKU": 100, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    need = raw[raw["需要調貨"] == "是"].iloc[0]
    nt_row = raw[raw["Site Code"] == "N01"].iloc[0]
    mix_row = raw[raw["Site Code"] == "M01"].iloc[0]
    assert int(need["調入需求量 (補到14日)"]) == 9
    assert int(need["實際調入量"]) == 9
    assert int(need["調入缺口"]) == 0
    assert int(nt_row["轉出量"]) == 9
    assert int(mix_row["轉出量"]) == 0
    assert stats["nt_transfer"] == 9
    assert stats["mix_transfer"] == 0


def test_mix_supplement_when_nt_short():
    # Tourist 需求 20；NT 可轉出 5；Mix 可轉出 8 → 實際調入 13，缺口 7
    df = _df([
        {"SKU": 200, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 720.0, "Shop Stock Qty (CM)": 92.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 200, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 19.0},
        {"SKU": 200, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 22.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    need = raw[raw["需要調貨"] == "是"].iloc[0]
    nt_row = raw[raw["Site Code"] == "N01"].iloc[0]
    mix_row = raw[raw["Site Code"] == "M01"].iloc[0]
    assert int(need["調入需求量 (補到14日)"]) == 20
    assert int(nt_row["轉出量"]) == 5
    assert int(mix_row["轉出量"]) == 8
    assert int(need["實際調入量"]) == 13
    assert int(need["調入缺口"]) == 7
    assert stats["nt_transfer"] == 5
    assert stats["mix_transfer"] == 8
    assert stats["total_gap"] == 7


def test_no_sales_with_stock_fully_releasable():
    # 無銷售但有存貨 → 全數可轉出（不限 14 日標準）
    df = _df([
        {"SKU": 1, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 180.0, "Shop Stock Qty (CM)": 14.0,
         "Turnday (P90D Qty)": 7.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 1, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 10.0},
    ])
    raw, _, _, _ = transform_stock_checking(df)
    nt_row = raw[raw["Site Code"] == "N01"].iloc[0]
    assert int(nt_row["可轉出貨量 (Non Tourist)"]) == 10


def test_nd_source_shop_participates():
    # ND 來源店亦可轉出
    df = _df([
        {"SKU": 1, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 1, "Site Code": "N01", "Shop Nature": "Non Tourist", "RP Type": "ND",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["nt_transfer"] == 9
    assert int(raw[raw["Site Code"] == "N01"].iloc[0]["轉出量"]) == 9


# ──────────────────────────────────────────────────────────────────────────────
# 4. 自訂門檻、欄位容錯
# ──────────────────────────────────────────────────────────────────────────────
def test_custom_thresholds():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Turnday (P90D Qty)": 15.0, "D001 Net Stock Qty (CM)": 200.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df, turnday_threshold=20, d001_threshold=250)
    assert stats["need_rows"] == 1
    raw2, _, stats2, _ = transform_stock_checking(df, turnday_threshold=14, d001_threshold=250)
    assert stats2["need_rows"] == 0


def test_d001_column_fallback():
    # 缺少 D001 Net Stock Qty (CM) 時退用 D001 Stock Qty (CM)
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Turnday (P90D Qty)": 13.0,
         "D001 Net Stock Qty (CM)": np.nan, "D001 Stock Qty (CM)": 50.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)"])
    raw, _, stats, _ = transform_stock_checking(df, d001_threshold=100)
    assert stats["need_rows"] == 1
    assert int(raw.iloc[0]["調入需求量 (補到14日)"]) == 9


def test_missing_d001_warns_and_treated_as_zero():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Turnday (P90D Qty)": 13.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)"])
    raw, _, stats, warnings_list = transform_stock_checking(df)
    assert stats["need_rows"] == 1
    assert any("D001" in w for w in warnings_list)


# ──────────────────────────────────────────────────────────────────────────────
# 5. SKU Summary（參考 By SKUxShop_Stock Checking_Output 格式）
# ──────────────────────────────────────────────────────────────────────────────
def test_summary_one_row_per_sku_and_nature_sums():
    # SKU 300：Tourist 1 店、NT 2 店、Mix 1 店
    df = _df([
        {"SKU": 300, "Brand": "B2", "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 300, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Shop Monthly Sold Qty (P90D)": 30.0, "Shop Stock Qty (CM)": 20.0},
        {"SKU": 300, "Site Code": "N02", "Shop Nature": "Non Tourist",
         "Shop Monthly Sold Qty (P90D)": 50.0, "Shop Stock Qty (CM)": 30.0},
        {"SKU": 300, "Site Code": "M01", "Shop Nature": "Mix",
         "Shop Monthly Sold Qty (P90D)": 40.0, "Shop Stock Qty (CM)": 25.0},
        # 另一個 SKU 確認彙總分列
        {"SKU": 301, "Brand": "B2", "Site Code": "T02", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.0,
         "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 50.0},
    ])
    raw, summary, _, _ = transform_stock_checking(df)
    assert len(summary) == 2
    row300 = summary[summary["SKU"] == "300"].iloc[0]
    assert row300["Brand"] == "B2"
    assert row300["D001 Net Stock Qty (CM)"] == 10.0
    assert row300["Tourist 加總 - Shop Stock Qty (CM)"] == 47.0
    assert row300["Tourist 加總 - Shop Monthly Sold Qty (P90D)"] == 120.0
    # TurnDay = 47 / 120 × 30.5
    assert abs(row300["Tourist TurnDay"] - round(47.0 / 120.0 * 30.5, 2)) < 0.01
    assert row300["Mix 加總 - Shop Stock Qty (CM)"] == 25.0
    assert row300["Non Tourist 加總 - Shop Stock Qty (CM)"] == 50.0  # 20 + 30
    assert row300["Non Tourist 加總 - Shop Monthly Sold Qty (P90D)"] == 80.0  # 30 + 50
    assert "由Tourist 店舖預計調出數量" not in summary.columns


def test_summary_transfer_prompt_and_transfer_out_sums():
    # SKU 400：Tourist 需求 9（P90D=360 日均 4 → 理想 56，庫存 47）
    # NT 可轉出 4、Mix 可轉出 20 → NT 調出 4、Mix 調出 5、缺口 0
    df = _df([
        {"SKU": 400, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 400, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Shop Monthly Sold Qty (P90D)": 30.0, "Shop Stock Qty (CM)": 18.0},
        {"SKU": 400, "Site Code": "M01", "Shop Nature": "Mix",
         "Shop Monthly Sold Qty (P90D)": 40.0, "Shop Stock Qty (CM)": 26.0},
        # 無需調貨的 SKU 對照
        {"SKU": 401, "Site Code": "M02", "Shop Nature": "Mix",
         "Shop Monthly Sold Qty (P90D)": 40.0, "Shop Stock Qty (CM)": 10.0},
    ])
    raw, summary, _, _ = transform_stock_checking(df)
    row400 = summary[summary["SKU"] == "400"].iloc[0]
    row401 = summary[summary["SKU"] == "401"].iloc[0]
    assert row400["強制調貨提示"] == "需要調貨"
    assert int(row400["由Non Tourist 店舖預計調出數量"]) == 4
    assert int(row400["由Mix 店舖預計調出數量"]) == 5
    assert row401["強制調貨提示"] == "" or pd.isna(row401["強制調貨提示"])
    assert int(row401["由Non Tourist 店舖預計調出數量"]) == 0
    assert int(row401["由Mix 店舖預計調出數量"]) == 0


def test_summary_sorted_by_brand_sku():
    df = _df([
        {"SKU": 2, "Brand": "BB", "Site Code": "M01", "Shop Nature": "Mix", "Shop Monthly Sold Qty (P90D)": 10.0},
        {"SKU": 1, "Brand": "AA", "Site Code": "M02", "Shop Nature": "Mix", "Shop Monthly Sold Qty (P90D)": 10.0},
    ])
    _, summary, _, _ = transform_stock_checking(df)
    assert list(summary["Brand"]) == ["AA", "BB"]


# ──────────────────────────────────────────────────────────────────────────────
# 6. Excel 輸出冒煙測試
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_excel():
    df = _df([
        {"SKU": 100, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 100, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Shop Monthly Sold Qty (P90D)": 30.0, "Shop Stock Qty (CM)": 30.0},
    ])
    raw, summary, _, _ = transform_stock_checking(df)
    buf = generate_excel_stock_checking(raw, summary)
    assert isinstance(buf, BytesIO)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == ["Summary", "RAW"]
    ws = wb["Summary"]
    headers = [c.value for c in next(ws.iter_rows(min_row=5, max_row=5))]
    for col in ["Brand", "SKU", "強制調貨提示", "由Non Tourist 店舖預計調出數量", "由Mix 店舖預計調出數量"]:
        assert col in headers
    assert ws.max_row >= 6
    ws_raw = wb["RAW"]
    raw_headers = [c.value for c in next(ws_raw.iter_rows(min_row=1, max_row=1))]
    for col in ["需要調貨", "調入需求量 (補到14日)", "實際調入量", "可轉出貨量 (Non Tourist)", "調貨來源 (Site Code)"]:
        assert col in raw_headers