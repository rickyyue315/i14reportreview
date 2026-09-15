"""Validation tests for Safety Stock x NetStock Review「建議退D001」option (return_to_d001).

接收目標為 D001（Warehouse，可無限接收貨量），故不跑分配：
- 「建議退D001 量」＝該列可轉出量（池級門檻通過的 14 日超額；雙向時含 Tourist 欄）
  扣除「轉出量」（不限OM 分配實績）後的剩餘，取整、下限 0——調貨優先、剩餘全數退回
- 池級門檻不通過的 Nature 池一律不建議退貨；零銷有存在所屬池健康時全數庫存計為可退
- ND 列可作供應方，同樣適用；缺「RP Type」欄、OM 限制與分類維度均不影響本欄
- 本欄只作退倉建議，不改變調貨分配、Safety Review 與建議調整量（主表欄序契約不動）
- 開啟時：RAW 新增欄、Excel 另加「退D001 建議」工作表（9→10 表）、
  指標卡另加退D001 概況一排；關閉時輸出欄位與數值與現行完全一致（9 表）
"""
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from transformer import (
    transform_hybrid,
    generate_excel_hybrid,
    _hybrid_dashboard_metrics,
    _hybrid_stats_from_raw,
    _hybrid_raw_display_columns,
    _build_hybrid_return_d001,
    _build_hybrid_diagnosis,
    _HYBRID_RETURN_D001_COLUMN,
    _HYBRID_RETURN_D001_SHEET,
    _HYBRID_RETURN_D001_COLUMNS,
    _HYBRID_DIAG_COLUMNS,
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1, "SKU Desc": "Test Product",
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Tourist", "Vendor MC": "V1", "Item Status": "A",
    "Adj. SST": "SST1",
    "Launch Date": "2024-01-01", "Item Class": "C1", "RP Type": "RF",
    "Shop Stock Status": "Normal", "Adj. Safety Stock": 3.0,
    "Sold Qty (P90D)": 90.0, "Sold Qty (M-2)": 30.0, "Sold Qty (M-1)": 30.0,
    "Excess Stock Qty": 2.0,
    "Shop Stock Qty (CM)": 5.0,
    "Shop Monthly Sold Qty (P90D)": 30.5,
    "Sold Qty (CM - MTD)": 8.0,
    "D001 Stock Qty (CM)": 200.0, "D001 Net Stock Qty (CM)": 99.0,
    "Turnday (P90D Qty)": 13.0,
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


# 標準三列場景（SKU 300）：T01 需調入 9；N01 可轉出 16、轉出 9；M01 可轉出 11、轉出 0
_SCENARIO_A = [
    {"SKU": 300, "Site Code": "T01", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
     "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 300, "Site Code": "N01", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30},
    {"SKU": 300, "Site Code": "M01", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 14},
]

# 零銷有存（單列 SKU 400，NT 池月銷 0 → 池 TurnDay 無限大 → 池健康）：全數庫存可退
_SCENARIO_ZERO_SALE = [
    {"SKU": 400, "Site Code": "N02", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 20.0,
     "Shop Monthly Sold Qty (P90D)": 0.0, "Adj. Safety Stock": 0},
]

# 池級門檻不通過（NT 池兩列：一列有 14 日超額，但池加總 TurnDay 2.73 日 < 14）
_SCENARIO_GATE = [
    {"SKU": 500, "Site Code": "N03", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 5},
    {"SKU": 500, "Site Code": "N04", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 900.0, "Shop Stock Qty (CM)": 0.0,
     "Shop Monthly Sold Qty (P90D)": 305.0, "Adj. Safety Stock": 0},
]

# 雙向：同 SKU 兩列——N31 (NT) 缺貨 9；T32 (Tourist) 有 86 件超額（轉出 9 → 退 77；
# 同 Nature 互調仍被「需求 Nature 自身除外」守衛排除，故供應列用 Tourist、需求列用 NT）
_SCENARIO_BIDIR = [
    {"SKU": 600, "Site Code": "N31", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
     "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 600, "Site Code": "T32", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 100.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30},
]


def _get(raw, site, col):
    return raw.loc[raw["Site Code"] == site, col].iloc[0]


def test_return_column_math_and_stats():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_A), return_to_d001=True)
    assert _HYBRID_RETURN_D001_COLUMN in raw.columns
    assert int(_get(raw, "N01", "可轉出貨量 (Non Tourist)")) == 16
    assert int(_get(raw, "N01", "轉出量")) == 9
    assert int(_get(raw, "N01", _HYBRID_RETURN_D001_COLUMN)) == 7  # 16 − 已轉出 9
    assert int(_get(raw, "M01", _HYBRID_RETURN_D001_COLUMN)) == 11  # 全數剩餘（未被取用）
    assert int(_get(raw, "T01", _HYBRID_RETURN_D001_COLUMN)) == 0  # 需求列無超額
    assert stats["return_to_d001"] is True
    assert stats["return_rows"] == 2
    assert stats["return_skus"] == 1
    assert stats["return_units"] == 18


def test_off_keeps_existing_contract():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, warn = transform_hybrid(_df(_SCENARIO_A))
    assert _HYBRID_RETURN_D001_COLUMN not in raw.columns
    assert stats["return_to_d001"] is False
    assert stats["return_rows"] == 0 and stats["return_skus"] == 0 and stats["return_units"] == 0
    assert not any("退D001" in w for w in warn)
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == [
        "說明", "RAW", "Summary", "調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)",
        "店舖問題 SKU", "Review 摘要", "簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)",
    ]


def test_zero_sale_full_stock_return_without_transfer_demand():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_ZERO_SALE), return_to_d001=True)
    # 無任何調貨需求列（flagged SKU 為空）仍可退——退D001 不以有調貨需求為前提
    assert int(_get(raw, "N02", _HYBRID_RETURN_D001_COLUMN)) == 20
    assert stats["return_rows"] == 1 and stats["return_skus"] == 1 and stats["return_units"] == 20


def test_pool_gate_blocks_return():
    raw, *_rest, stats, warn = transform_hybrid(_df(_SCENARIO_GATE), return_to_d001=True)
    # N03 逐列有 14 日超額（30 − 14 = 16），惟 NT 池加總 TurnDay 2.73 日 < 門檻 → 一律不退
    assert int(_get(raw, "N03", _HYBRID_RETURN_D001_COLUMN)) == 0
    assert int(_get(raw, "N04", _HYBRID_RETURN_D001_COLUMN)) == 0
    assert stats["return_units"] == 0
    assert any("沒有可退回 D001" in w for w in warn)


def test_bidirectional_includes_tourist_pool():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_BIDIR), bidirectional=True, return_to_d001=True)
    assert int(_get(raw, "T32", "可轉出貨量 (Tourist)")) == 86
    assert int(_get(raw, "T32", "轉出量")) == 9
    assert int(_get(raw, "T32", _HYBRID_RETURN_D001_COLUMN)) == 77  # 86 − 已轉出 9
    assert stats["return_units"] == 77
    # 單向：Tourist 列不作供應方，亦不建議退貨
    raw1, *_rest1, stats1, warn1 = transform_hybrid(_df(_SCENARIO_BIDIR), return_to_d001=True)
    assert "可轉出貨量 (Tourist)" not in raw1.columns
    assert stats1["return_units"] == 0
    assert int(_get(raw1, "T32", _HYBRID_RETURN_D001_COLUMN)) == 0


def test_nd_rows_can_return():
    rows = [dict(r) for r in _SCENARIO_A]
    for r in rows:
        r["RP Type"] = "ND" if r["Site Code"] == "N01" else "RF"
    raw, *_rest, stats, _warn = transform_hybrid(_df(rows), return_to_d001=True)
    assert str(_get(raw, "N01", "RP Type")) == "ND"
    assert int(_get(raw, "N01", _HYBRID_RETURN_D001_COLUMN)) == 7
    assert stats["return_units"] == 18


def test_return_independent_of_rp_column_and_dimension():
    # 缺「RP Type」欄（全列按 RF）與 rp 維度均不影響退D001 欄
    df_no_rp = _df(_SCENARIO_A).drop(columns=["RP Type"])
    raw, *_rest, stats, _warn = transform_hybrid(df_no_rp, return_to_d001=True)
    assert int(_get(raw, "N01", _HYBRID_RETURN_D001_COLUMN)) == 7
    raw_rp, *_rest_rp, stats_rp, _warn2 = transform_hybrid(_df(_SCENARIO_A), dimension="rp", return_to_d001=True)
    assert int(_get(raw_rp, "N01", _HYBRID_RETURN_D001_COLUMN)) == 7
    assert int(_get(raw_rp, "M01", _HYBRID_RETURN_D001_COLUMN)) == 11


def test_transfer_and_safety_untouched_by_switch():
    on = transform_hybrid(_df(_SCENARIO_A), return_to_d001=True)
    off = transform_hybrid(_df(_SCENARIO_A))
    for col in ["實際調入量", "調入缺口", "轉出量", "調貨來源 (Site Code)", "調貨後庫存",
                "事件類型", "Safety Review", "建議 Safety 設定值", "Safety 建議調整量"]:
        assert pd.Series(on[0][col]).equals(pd.Series(off[0][col])), col
    # 主表欄序契約不動（退D001 不進兩張 Review 主表）
    assert _HYBRID_RETURN_D001_COLUMN not in _HYBRID_DIAG_COLUMNS
    assert list(on[2].columns) == list(off[2].columns)


def test_raw_display_order_and_builder():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_A), return_to_d001=True)
    ordered = _hybrid_raw_display_columns(raw)
    assert ordered.index(_HYBRID_RETURN_D001_COLUMN) == ordered.index("轉出量") + 1
    ret = _build_hybrid_return_d001(raw)
    assert list(ret.columns) == _HYBRID_RETURN_D001_COLUMNS
    assert list(ret["Site Code"]) == ["M01", "N01"]  # 按 Site Code 排序
    assert list(ret["可轉出貨量"]) == [11, 16]
    assert list(ret["轉出量"]) == [0, 9]
    assert list(ret[_HYBRID_RETURN_D001_COLUMN]) == [11, 7]
    # 關閉或全 0 時輸出空表
    raw_off = transform_hybrid(_df(_SCENARIO_A))[0]
    assert _build_hybrid_return_d001(raw_off).empty
    raw_gate, *_rest_g, _stats_g, _warn_g = transform_hybrid(_df(_SCENARIO_GATE), return_to_d001=True)
    assert _build_hybrid_return_d001(raw_gate).empty


def test_excel_ten_sheets_and_headers():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _warn = transform_hybrid(
        _df(_SCENARIO_A), return_to_d001=True)
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == [
        "說明", "RAW", "Summary", "調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)",
        "店舖問題 SKU", _HYBRID_RETURN_D001_SHEET, "Review 摘要",
        "簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)",
    ]
    ws = wb[_HYBRID_RETURN_D001_SHEET]
    headers = [c.value for c in next(ws.iter_rows(min_row=3, max_row=3))]
    assert headers == _HYBRID_RETURN_D001_COLUMNS
    # 明細兩列（M01 / N01）＋標題列列數
    data_rows = [r for r in ws.iter_rows(min_row=4, min_col=5, max_col=5) if r[0].value]
    assert len(data_rows) == 2


def test_excel_empty_return_sheet_still_present_when_on():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _warn = transform_hybrid(
        _df(_SCENARIO_GATE), return_to_d001=True)
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert _HYBRID_RETURN_D001_SHEET in wb.sheetnames
    ws = wb[_HYBRID_RETURN_D001_SHEET]
    texts = [c.value for row in ws.iter_rows(max_row=5) for c in row]
    assert any("無符合資料" in str(t) for t in texts if t)


def test_dashboard_metrics_and_stats_fallback():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_A), return_to_d001=True)
    rows = _hybrid_dashboard_metrics(raw, stats)
    ret_row = rows[-1]
    assert [name for name, *_ in ret_row] == ["退D001 列數", "退D001 SKU 數", "退D001 合計件數", "已轉出保留店舖件數"]
    assert [v for _, v, *_ in ret_row] == [2, 1, 18, 9]
    # stats 缺失的回退路徑：由 RAW 欄位存在推定
    fb = _hybrid_stats_from_raw(raw)
    assert fb["return_to_d001"] is True
    assert fb["return_units"] == 18 and fb["return_rows"] == 2
    raw_off = transform_hybrid(_df(_SCENARIO_A))[0]
    fb_off = _hybrid_stats_from_raw(raw_off)
    assert fb_off["return_to_d001"] is False and fb_off["return_units"] == 0


def test_guide_sheet_mentions_return():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _warn = transform_hybrid(
        _df(_SCENARIO_A), return_to_d001=True)
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf)
    ws = wb["說明"]
    col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    col_b = [ws.cell(row=r, column=2).value for r in range(1, ws.max_row + 1)]
    pairs = dict(zip(col_a, col_b))
    assert pairs["建議退D001"].startswith("開啟")
    assert "退D001 建議" in col_a
    logic_rows = [b for a, b in zip(col_a, col_b) if a == "•"]
    assert any("建議退D001" in str(b) for b in logic_rows)
