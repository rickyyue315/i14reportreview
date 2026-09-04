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
    "Adj. SST": "SST1",
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
    # 無 Shop Monthly Sold Qty (P90D) 欄位 → 以 P90D/90×30.5 換算月銷
    # P90D=90 → 月銷 30.5 → Tourist 加總 TurnDay = 庫存加總 / 30.5 × 30.5 = 庫存加總
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},  # 符合（TurnDay 5 < 14，D001 99 < 100）
        {"SKU": 2, "Site Code": "S02", "Shop Stock Qty (CM)": 14.0, "D001 Net Stock Qty (CM)": 99.0},  # 加總 TurnDay = 14（門檻邊界）→ 不觸發
        {"SKU": 3, "Site Code": "S03", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 100.0},  # D001 = 門檻 → 不觸發
        {"SKU": 4, "Site Code": "S04", "Shop Nature": "Mix"},  # 非 Tourist → 不觸發
        {"SKU": 5, "Site Code": "S05", "Shop Nature": "Non Tourist"},  # 非 Tourist → 不觸發
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert list(raw.loc[raw["需要調貨"] == "是", "Site Code"]) == ["S01"]
    assert list(raw.loc[raw["需要調貨"] == "是", "SKU"]) == ["1"]
    assert stats["need_rows"] == 1
    assert stats["need_skus"] == 1


def test_aggregate_tourist_turnday_gates_trigger():
    # 多間 Tourist 店合併判定：加總 TurnDay = 加總庫存 ÷ 加總月銷 × 30.5
    # 店 A 個別 TurnDay = 10 < 14，但兩店加總 30/61×30.5 = 15 ≥ 14 → 整個 SKU 無須調貨
    df = _df([
        {"SKU": 500, "Site Code": "TA", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 10.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 500, "Site Code": "TB", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 20.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Turnday (P90D Qty)": 20.0, "D001 Net Stock Qty (CM)": 10.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["need_rows"] == 0
    assert (raw["需要調貨"] == "").all()
    assert (raw["Tourist 加總 TurnDay"] == 15.0).all()


def test_aggregate_tourist_turnday_below_threshold_flags_all_short_rows():
    # 加總 TurnDay = 10/61×30.5 = 5 < 14 且 D001 < 100 → 兩間缺貨 Tourist 店均標示
    df = _df([
        {"SKU": 501, "Site Code": "TA", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Turnday (P90D Qty)": 5.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 501, "Site Code": "TB", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Turnday (P90D Qty)": 5.0, "D001 Net Stock Qty (CM)": 10.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["need_rows"] == 2
    assert stats["need_skus"] == 1
    assert (raw.loc[raw["Shop Nature"] == "Tourist", "需要調貨"] == "是").all()
    assert (raw["調入需求量 (補到14日)"] == 9).all()


def test_turnday_invalid_flag_not_trigger():
    # Turnday = -1 表示 P90D=0（無銷售），即使 D001 低也不觸發
    df = _df([
        {"SKU": 1, "Sold Qty (P90D)": 0.0, "Turnday (P90D Qty)": 0.0,
         "Shop Stock Qty (CM)": 5.0, "D001 Net Stock Qty (CM)": 50.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["need_rows"] == 0
    assert (raw["需要調貨"] == "").all()


def test_both_sources_below_threshold_no_flag():
    # Mix 加總 TurnDay = 25/61×30.5 = 12.5 < 14、NT 加總 = 25/61×30.5 = 12.5 < 14
    # → 理論上無貨可調，即使 Tourist 加總 TurnDay 5 < 14 且 D001 < 100 亦不標示
    # Mix/NT 個別店有 14 日超額（20 − 14 = 6）但被 Nature 級門檻歸零
    df = _df([
        {"SKU": 600, "Site Code": "TA", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Turnday (P90D Qty)": 5.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 600, "Site Code": "MA", "Shop Nature": "Mix",
         "Shop Monthly Sold Qty (P90D)": 30.5, "Shop Stock Qty (CM)": 20.0},
        {"SKU": 600, "Site Code": "MB", "Shop Nature": "Mix",
         "Shop Monthly Sold Qty (P90D)": 30.5, "Shop Stock Qty (CM)": 5.0},
        {"SKU": 600, "Site Code": "NA", "Shop Nature": "Non Tourist",
         "Shop Monthly Sold Qty (P90D)": 30.5, "Shop Stock Qty (CM)": 20.0},
        {"SKU": 600, "Site Code": "NB", "Shop Nature": "Non Tourist",
         "Shop Monthly Sold Qty (P90D)": 30.5, "Shop Stock Qty (CM)": 5.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["need_rows"] == 0
    assert (raw["需要調貨"] == "").all()
    assert (raw["可轉出貨量 (Mix)"] == 0).all()
    assert (raw["可轉出貨量 (Non Tourist)"] == 0).all()
    assert (raw["Mix 加總 TurnDay"] == 12.5).all()
    assert (raw["Non Tourist 加總 TurnDay"] == 12.5).all()


def test_single_source_below_threshold_gated_other_donates():
    # Mix 加總 TurnDay 12.5 < 14 → Mix 可轉出量歸零（個別店超額 6 件也不調出）
    # NT 加總 TurnDay 30 ≥ 14 → 正常供應；Tourist 需求 9 全由 NT 調出
    df = _df([
        {"SKU": 700, "Site Code": "TA", "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 700, "Site Code": "MA", "Shop Nature": "Mix",
         "Shop Monthly Sold Qty (P90D)": 30.5, "Shop Stock Qty (CM)": 20.0},
        {"SKU": 700, "Site Code": "MB", "Shop Nature": "Mix",
         "Shop Monthly Sold Qty (P90D)": 30.5, "Shop Stock Qty (CM)": 5.0},
        {"SKU": 700, "Site Code": "NA", "Shop Nature": "Non Tourist",
         "Shop Monthly Sold Qty (P90D)": 30.5, "Shop Stock Qty (CM)": 30.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["need_rows"] == 1
    assert int(stats["nt_transfer"]) == 9
    assert int(stats["mix_transfer"]) == 0
    mix_rows = raw[raw["Shop Nature"] == "Mix"]
    assert (mix_rows["可轉出貨量 (Mix)"] == 0).all()
    assert (mix_rows["轉出量"] == 0).all()
    nt_row = raw[raw["Site Code"] == "NA"].iloc[0]
    assert int(nt_row["轉出量"]) == 9


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
    # 預設月銷 30.5（無 Shop Monthly Sold Qty 欄位時由 P90D 換算）→ 加總 TurnDay = 庫存 = 5
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Turnday (P90D Qty)": 15.0, "D001 Net Stock Qty (CM)": 200.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df, turnday_threshold=20, d001_threshold=250)
    assert stats["need_rows"] == 1
    raw2, _, stats2, _ = transform_stock_checking(df, turnday_threshold=3, d001_threshold=250)
    assert stats2["need_rows"] == 0  # 加總 TurnDay 5 ≥ 3 → 無須調貨


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
    assert row300["Tourist 加總 - Shop Stock Qty (CM)"] == 47
    assert row300["Tourist 加總 - Shop Monthly Sold Qty (P90D)"] == 120
    # TurnDay 取整：47 / 120 × 30.5 = 11.95 → 12
    assert row300["Tourist TurnDay"] == 12
    assert row300["Mix 加總 - Shop Stock Qty (CM)"] == 25
    assert row300["Non Tourist 加總 - Shop Stock Qty (CM)"] == 50  # 20 + 30
    assert row300["Non Tourist 加總 - Shop Monthly Sold Qty (P90D)"] == 80  # 30 + 50
    assert "由Tourist 店舖預計調出數量" not in summary.columns
    # Buyer / Adj. SST 在 SKU Desc 之後（同 SKU 首個非空值）
    assert list(summary.columns).index("Buyer") == list(summary.columns).index("SKU Desc") + 1
    assert list(summary.columns).index("Adj. SST") == list(summary.columns).index("SKU Desc") + 2
    assert row300["Buyer"] == "B1"
    assert row300["Adj. SST"] == "SST1"
    # 佔比：庫存合計 122、月銷合計 240
    assert abs(row300["Tourist 加總 - Shop Stock Qty (CM) %"] - round(47 / 122, 4)) < 1e-9
    assert abs(row300["Tourist 加總 - Shop Monthly Sold Qty (P90D) %"] - 0.5) < 1e-9
    assert abs(row300["Mix 加總 - Shop Stock Qty (CM) %"] - round(25 / 122, 4)) < 1e-9
    assert abs(row300["Non Tourist 加總 - Shop Monthly Sold Qty (P90D) %"] - round(80 / 240, 4)) < 1e-9


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


def test_summary_stock_skew_hint():
    # NT 庫存佔比 70%、銷量佔比 30.3%（差 39.7 點 ≥ 20）→ 提示；Mix/Tourist 庫存佔比 < 40% 不提示
    df = _df([
        {"SKU": 700, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 150.0, "Shop Stock Qty (CM)": 10.0,
         "Shop Monthly Sold Qty (P90D)": 50.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 700, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 70.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 700, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 60.0, "Shop Stock Qty (CM)": 20.0,
         "Shop Monthly Sold Qty (P90D)": 20.0},
    ])
    _, summary, _, _ = transform_stock_checking(df)
    row = summary[summary["SKU"] == "700"].iloc[0]
    hint = str(row["存貨偏重提示"])
    assert hint.startswith("存貨偏重 Non Tourist（庫存 70.0% / 銷量 30.3%）")
    assert "Mix" not in hint and "Tourist（" not in hint.replace("Non Tourist（", "")


def test_summary_stock_skew_hint_none_when_balanced():
    # NT 庫存 50% 但銷量 45%（差 5 點 < 20）；其餘 Nature 庫存佔比 < 40% → 不提示
    df = _df([
        {"SKU": 701, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 15.0,
         "Shop Monthly Sold Qty (P90D)": 27.5, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 701, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 50.0,
         "Shop Monthly Sold Qty (P90D)": 27.5},
        {"SKU": 701, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 35.0,
         "Shop Monthly Sold Qty (P90D)": 5.5},
    ])
    _, summary, _, _ = transform_stock_checking(df)
    row = summary[summary["SKU"] == "701"].iloc[0]
    assert str(row["存貨偏重提示"]) in ("", "nan")


def test_summary_stock_skew_hint_multiple_natures():
    # NT 庫存 50% / 銷量 10%（差 40）；Mix 庫存 40% / 銷量 10%（差 30）→ 雙提示，差距大者在前
    df = _df([
        {"SKU": 702, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 10.0,
         "Shop Monthly Sold Qty (P90D)": 80.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 702, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 50.0,
         "Shop Monthly Sold Qty (P90D)": 10.0},
        {"SKU": 702, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 40.0,
         "Shop Monthly Sold Qty (P90D)": 10.0},
    ])
    _, summary, _, _ = transform_stock_checking(df)
    hint = str(summary[summary["SKU"] == "702"].iloc[0]["存貨偏重提示"])
    assert hint.startswith("存貨偏重 Non Tourist")
    assert "存貨偏重 Mix" in hint
    assert hint.index("Non Tourist") < hint.index("Mix")


def test_summary_stock_skew_hint_zero_sales():
    # 全 SKU 月銷 0 → 銷量佔比以 0 計：NT 庫存佔比 60% ≥ 40% 且差 60 點 → 提示（銷量 0.0%）
    df = _df([
        {"SKU": 703, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 60.0,
         "Shop Monthly Sold Qty (P90D)": 0.0},
        {"SKU": 703, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 40.0,
         "Shop Monthly Sold Qty (P90D)": 0.0},
    ])
    _, summary, _, _ = transform_stock_checking(df)
    hint = str(summary[summary["SKU"] == "703"].iloc[0]["存貨偏重提示"])
    assert hint.startswith("存貨偏重 Non Tourist（庫存 60.0% / 銷量 0.0%）")
    assert "存貨偏重 Mix（庫存 40.0% / 銷量 0.0%）" in hint  # 40% ≥ 40% 且差 40 點 ≥ 20 → 同樣入提示


def test_generate_excel_summary_skew_not_highlighted_by_default():
    # Stock Checking Summary 維持原樣（橘底加強顯示只用於 Safety Stock x NetStock Review）
    df = _df([
        {"SKU": 700, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 150.0, "Shop Stock Qty (CM)": 10.0,
         "Shop Monthly Sold Qty (P90D)": 50.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 700, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 70.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 700, "Site Code": "M01", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 60.0, "Shop Stock Qty (CM)": 20.0,
         "Shop Monthly Sold Qty (P90D)": 20.0},
    ])
    raw, summary, _, _ = transform_stock_checking(df)
    assert str(summary.iloc[0]["存貨偏重提示"]).startswith("存貨偏重 Non Tourist")
    buf = generate_excel_stock_checking(raw, summary)
    ws = load_workbook(buf)["Summary"]
    assert not str(ws.cell(row=6, column=26).fill.start_color.rgb).endswith("FFC000")
    assert not str(ws.cell(row=6, column=21).fill.start_color.rgb).endswith("FFC000")


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
    assert wb.sheetnames == ["Summary", "RAW", "調貨分析摘要"]
    ws = wb["Summary"]
    headers = [c.value for c in next(ws.iter_rows(min_row=5, max_row=5))]
    for col in ["Brand", "SKU", "Buyer", "Adj. SST", "強制調貨提示",
                "由Non Tourist 店舖預計調出數量", "由Mix 店舖預計調出數量", "存貨偏重提示"]:
        assert col in headers
    # Buyer / Adj. SST 緊隨 SKU Desc 之後
    assert headers.index("Buyer") == headers.index("SKU Desc") + 1
    assert headers.index("Adj. SST") == headers.index("SKU Desc") + 2
    # 共 26 欄（Z = 存貨偏重提示）；Q-V 為佔比欄（%），其餘數值欄取整
    assert len(headers) == 26
    assert ws.max_row >= 6
    wb2 = load_workbook(buf, data_only=False)
    ws2 = wb2["Summary"]
    # H,I 加總取整格式；J TurnDay 為 ROUND 活公式；Q 佔比為 % 格式
    assert ws2.cell(row=6, column=8).number_format == "#,##0"
    assert ws2.cell(row=6, column=9).number_format == "#,##0"
    assert str(ws2.cell(row=6, column=10).value).startswith("=ROUND(")
    assert ws2.cell(row=6, column=10).number_format == "#,##0"
    assert ws2.cell(row=6, column=17).number_format == "0.00%"
    # 群組標題
    assert ws2.cell(row=3, column=8).value == "Shop Nature"
    assert ws2.cell(row=3, column=17).value == "各 Nature 佔比 (%)"
    ws_raw = wb["RAW"]
    raw_headers = [c.value for c in next(ws_raw.iter_rows(min_row=1, max_row=1))]
    for col in ["需要調貨", "調入需求量 (補到14日)", "實際調入量", "可轉出貨量 (Non Tourist)", "調貨來源 (Site Code)"]:
        assert col in raw_headers
    ws_dash = wb["調貨分析摘要"]
    title = ws_dash.cell(row=1, column=1).value
    assert title and "Transfer Analysis Dashboard" in str(title)
    r1 = [ws_dash.cell(row=3, column=c).value for c in range(1, 5)]
    r2 = [ws_dash.cell(row=7, column=c).value for c in range(1, 5)]
    for label in ["需要調貨列數", "需調貨 SKU 數", "調入需求量", "實際調入量"]:
        assert label in r1
    for label in ["Non Tourist 轉出量", "Mix 轉出量", "無法足額缺口", "總明細行數"]:
        assert label in r2


def test_generate_excel_dashboard_matches_stats_no_stats_arg():
    # 不傳 stats 時，Dashboard 指標由 raw 匯總（與有 stats 時一致）
    df = _df([
        {"SKU": 100, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0,
         "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 100, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Shop Monthly Sold Qty (P90D)": 30.0, "Shop Stock Qty (CM)": 30.0},
    ])
    raw, summary, stats, _ = transform_stock_checking(df)
    buf_no_stats = generate_excel_stock_checking(raw, summary)
    buf_with_stats = generate_excel_stock_checking(raw, summary, stats)
    for buf in (buf_no_stats, buf_with_stats):
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


# ──────────────────────────────────────────────────────────────────────────────
# Buyer 需調貨 SKU 統計
# ──────────────────────────────────────────────────────────────────────────────
def test_buyer_need_skus_stats():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Buyer": "B1", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},
        {"SKU": 2, "Site Code": "S02", "Buyer": "B2", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},
        {"SKU": 3, "Site Code": "S03", "Buyer": "B1", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},
        {"SKU": 4, "Site Code": "S04", "Buyer": "B3", "Shop Stock Qty (CM)": 20.0, "D001 Net Stock Qty (CM)": 99.0},  # TurnDay 20 ≥ 14 → 不觸發
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["need_skus"] == 3
    assert stats["buyer_need_skus"] == [("B1", 2), ("B2", 1)]


def test_buyer_need_skus_stats_empty_buyer():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Buyer": "", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["buyer_need_skus"] == [("(未填 Buyer)", 1)]


def test_buyer_need_skus_stats_none_triggered():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Shop Stock Qty (CM)": 20.0, "D001 Net Stock Qty (CM)": 99.0},
    ])
    raw, _, stats, _ = transform_stock_checking(df)
    assert stats["need_skus"] == 0
    assert stats["buyer_need_skus"] == []


def test_generate_excel_buyer_stats_table():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Buyer": "B1", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},
        {"SKU": 2, "Site Code": "S02", "Buyer": "B2", "Turnday (P90D Qty)": 13.0, "D001 Net Stock Qty (CM)": 99.0},
    ])
    raw, summary, stats, _ = transform_stock_checking(df)
    for buf in (generate_excel_stock_checking(raw, summary, stats),
                generate_excel_stock_checking(raw, summary)):  # 不傳 stats 亦由 raw 匯總
        wb = load_workbook(buf)
        ws = wb["調貨分析摘要"]
        assert "Buyer 需調貨 SKU 統計" in str(ws.cell(row=11, column=1).value)
        assert ws.cell(row=12, column=1).value == "Buyer"
        assert ws.cell(row=12, column=2).value == "需調貨 SKU 數"
        rows = {ws.cell(row=r, column=1).value: ws.cell(row=r, column=2).value for r in (13, 14)}
        assert rows == {"B1": 1, "B2": 1}
        assert ws.cell(row=15, column=1).value == "總計"
        assert ws.cell(row=15, column=2).value == 2
        assert ws.cell(row=13, column=3).number_format == "0.0%"


# ──────────────────────────────────────────────────────────────────────────────
# 重複列與自我調貨防護（同店同 SKU 資料異常）
# ──────────────────────────────────────────────────────────────────────────────
def test_duplicate_site_sku_conflicting_nature_warned():
    df = _df([
        {"SKU": 600, "Site Code": "S01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 600, "Site Code": "S01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, stats, warns = transform_stock_checking(df)
    assert any("Site Code × SKU" in w and "重複" in w for w in warns)
    assert any("Shop Nature" in w and "不一致" in w for w in warns)


def test_self_transfer_guarded_and_no_site_both_in_and_out():
    # S01 同時為需求列（Tourist）與唯一來源列（Non Tourist）：修復後同店來源排除，
    # 需求全數列缺口，不得出現自我調貨或同店同時調入與調出
    df = _df([
        {"SKU": 600, "Site Code": "S01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 600, "Site Code": "S01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 600, "Site Code": "S02", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "D001 Net Stock Qty (CM)": 10.0},
    ])
    raw, _, stats, warns = transform_stock_checking(df)
    t01 = raw[(raw["Site Code"] == "S01") & (raw["Shop Nature"] == "Tourist")].iloc[0]
    t02 = raw[raw["Site Code"] == "S02"].iloc[0]
    n01 = raw[(raw["Site Code"] == "S01") & (raw["Shop Nature"] == "Non Tourist")].iloc[0]
    assert int(t01["實際調入量"]) == 0 and int(t01["調入缺口"]) == 9
    assert str(t01["調貨來源 (Site Code)"]) == ""
    assert int(t02["實際調入量"]) == 9 and str(t02["調貨來源 (Site Code)"]) == "S01:9"
    assert int(n01["轉出量"]) == 9
    g = raw.groupby(["SKU", "Site Code"])[["實際調入量", "轉出量"]].sum()
    assert not ((g["實際調入量"] > 0) & (g["轉出量"] > 0)).any()
    assert int(raw["實際調入量"].sum()) == int(raw["轉出量"].sum())