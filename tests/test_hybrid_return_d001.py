"""Validation tests for Safety Stock x NetStock Review「建議退D001」option (return_to_d001).

接收目標為 D001（Warehouse，可無限接收貨量），故不跑分配。退貨基準為 Safety Stock (現時)
（非調貨段的 14 日線，與池級門檻脫鉤，逐列判定）：
- 中銷／低銷／零銷列：可退＝max(0, 庫存 − Safety)
- 高銷列：原則上不退，僅當 Safety > 0 且庫存 ≥ 3×Safety（高於 Safety 200%）時
  可退＝庫存 − 2×Safety（退至 2 倍 Safety）
- 調貨優先：建議退D001 量＝可退 − 轉出量（不限OM 分配實績），「需要調貨」列一律不退
- 缺「RP Type」欄、OM 限制與分類維度均不影響本欄
- 本欄只作退倉建議，不改變調貨分配、Safety Review 與建議調整量（主表欄序契約不動）
- 開啟時：RAW 新增欄、Excel 另加「退D001 建議」工作表（9→10 表）、
  指標卡另加退D001 概況一排；關閉時輸出欄位與數值與現行完全一致（9 表）
"""
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from transformer import (
    transform_hybrid,
    generate_excel_hybrid,
    _hybrid_dashboard_metrics,
    _hybrid_stats_from_raw,
    _hybrid_raw_display_columns,
    _build_hybrid_return_d001,
    _build_hybrid_diagnosis,
    _HYBRID_RETURN_D001_COLUMN,
    _HYBRID_RETURN_D001_GROSS_COLUMN,
    _HYBRID_RETURN_D001_SHEET,
    _HYBRID_RETURN_D001_COLUMNS,
    _HYBRID_DIAG_COLUMNS,
    _display_safety_header,
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


# 標準三列場景（SKU 300，3 間有銷店舖 → 全中銷）：
# T01 需要調貨（一律不退）；N01 庫存 30 − Safety 10 = 20、已轉出 9 → 退 11；
# M01 庫存 25 − Safety 14 = 11 → 退 11
_SCENARIO_A = [
    {"SKU": 300, "Site Code": "T01", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
     "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 300, "Site Code": "N01", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
    {"SKU": 300, "Site Code": "M01", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 14},
]

# 零銷（P90D=0 → 零銷帶）：N02 Safety 0 → 全數退 20；N05 Safety 5 → 退 15
_SCENARIO_ZERO_SALE = [
    {"SKU": 400, "Site Code": "N02", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 20.0,
     "Shop Monthly Sold Qty (P90D)": 0.0, "Adj. Safety Stock": 0},
    {"SKU": 401, "Site Code": "N05", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 20.0,
     "Shop Monthly Sold Qty (P90D)": 0.0, "Adj. Safety Stock": 5},
]

# 高銷例外（SKU 700，5 間有銷店舖：H1/H2 高銷，其餘中銷）：
# H1 庫存 18 = 3×Safety 6（高於 Safety 200%）→ 退至 2×Safety：18 − 12 = 6；
# H2 庫存 15 < 18 → 不退；中銷 N1 退 6、N2 Safety 高於庫存不退、M1 退 6
_SCENARIO_HIGH = [
    {"SKU": 700, "Site Code": "H1", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 900.0, "Shop Stock Qty (CM)": 18.0,
     "Shop Monthly Sold Qty (P90D)": 305.0, "Adj. Safety Stock": 6},
    {"SKU": 700, "Site Code": "H2", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 800.0, "Shop Stock Qty (CM)": 15.0,
     "Shop Monthly Sold Qty (P90D)": 271.0, "Adj. Safety Stock": 6},
    {"SKU": 700, "Site Code": "N1", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 12.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 6},
    {"SKU": 700, "Site Code": "N2", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 12.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 20},
    {"SKU": 700, "Site Code": "M1", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 9.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 3},
]

# 高銷且 Safety 0：門檻不成立（Safety > 0 才生效），高銷列一律不退
_SCENARIO_HIGH_ZERO_SAFETY = [
    {"SKU": 702, "Site Code": "H3", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 900.0, "Shop Stock Qty (CM)": 50.0,
     "Shop Monthly Sold Qty (P90D)": 305.0, "Adj. Safety Stock": 0},
    {"SKU": 702, "Site Code": "N6", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 12.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 6},
    {"SKU": 702, "Site Code": "N7", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 80.0, "Shop Stock Qty (CM)": 10.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 5},
    {"SKU": 702, "Site Code": "M2", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 70.0, "Shop Stock Qty (CM)": 9.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 3},
]

# 池級門檻脫鉤：NT 池加總 TurnDay 2.73 日 < 門檻（舊規則整池不退），新規則逐列判定仍退
_SCENARIO_STARVING_POOL = [
    {"SKU": 500, "Site Code": "N03", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 5},
    {"SKU": 500, "Site Code": "N04", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 900.0, "Shop Stock Qty (CM)": 0.0,
     "Shop Monthly Sold Qty (P90D)": 305.0, "Adj. Safety Stock": 0},
]

# 單列（中銷，庫存 10 = Safety 10 → 無可退）——Excel 空表契約用
_SCENARIO_EMPTY = [
    {"SKU": 800, "Site Code": "T81", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 10.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
]

# 雙向：N31 (NT) 為需要調貨列（不退）；T32 (Tourist) 中銷庫存 100 − Safety 30 = 70
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
    # 中銷列：可退＝庫存 − Safety；調貨優先——先扣已轉出
    assert int(_get(raw, "N01", "Adj. Safety Stock")) == 10
    assert int(_get(raw, "N01", "轉出量")) == 9
    assert int(_get(raw, "N01", _HYBRID_RETURN_D001_COLUMN)) == 11  # (30 − 10) − 9
    assert int(_get(raw, "M01", _HYBRID_RETURN_D001_COLUMN)) == 11  # 25 − 14
    # 需要調貨列一律不退
    assert int(_get(raw, "T01", _HYBRID_RETURN_D001_COLUMN)) == 0
    assert stats["return_to_d001"] is True
    assert stats["return_rows"] == 2
    assert stats["return_skus"] == 1
    assert stats["return_units"] == 22


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
        "店舖問題 SKU", "Review 摘要", "Safety 設定建議 (全列)", "簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)",
    ]


def test_zero_sale_band_returns_above_safety():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_ZERO_SALE), return_to_d001=True)
    # 零銷帶：Safety 0 → 全數退；Safety 5 → 退 15（保留 Safety 5）
    assert int(_get(raw, "N02", _HYBRID_RETURN_D001_COLUMN)) == 20
    assert int(_get(raw, "N05", _HYBRID_RETURN_D001_COLUMN)) == 15
    assert stats["return_rows"] == 2 and stats["return_skus"] == 2 and stats["return_units"] == 35


def test_high_band_exception_threshold():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_HIGH), return_to_d001=True)
    band = "銷量帶"
    assert str(_get(raw, "H1", band)) == "高銷" and str(_get(raw, "H2", band)) == "高銷"
    # 高銷：庫存 18 = 3×Safety 6（高於 200%）→ 退至 2×Safety = 12，退 6 件
    assert int(_get(raw, "H1", _HYBRID_RETURN_D001_COLUMN)) == 6
    # 高銷：庫存 15 < 3×Safety 6 → 不退
    assert int(_get(raw, "H2", _HYBRID_RETURN_D001_COLUMN)) == 0
    # 中銷：退至 Safety
    assert int(_get(raw, "N1", _HYBRID_RETURN_D001_COLUMN)) == 6   # 12 − 6
    assert int(_get(raw, "N2", _HYBRID_RETURN_D001_COLUMN)) == 0   # Safety 20 高於庫存
    assert int(_get(raw, "M1", _HYBRID_RETURN_D001_COLUMN)) == 6   # 9 − 3
    assert stats["return_rows"] == 3 and stats["return_skus"] == 1 and stats["return_units"] == 18


def test_high_band_zero_safety_never_returns():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_HIGH_ZERO_SAFETY), return_to_d001=True)
    assert str(_get(raw, "H3", "銷量帶")) == "高銷"
    assert int(_get(raw, "H3", _HYBRID_RETURN_D001_COLUMN)) == 0  # Safety 0 → 門檻不成立
    # 其餘中銷列照常退至 Safety
    assert int(_get(raw, "N6", _HYBRID_RETURN_D001_COLUMN)) == 6
    assert int(_get(raw, "N7", _HYBRID_RETURN_D001_COLUMN)) == 5
    assert int(_get(raw, "M2", _HYBRID_RETURN_D001_COLUMN)) == 6
    assert stats["return_units"] == 17


def test_return_decoupled_from_pool_gate():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_STARVING_POOL), return_to_d001=True)
    # NT 池加總 TurnDay 2.73 日 < 門檻（池級門檻不再適用於退D001），逐列判定仍退 25
    assert int(_get(raw, "N03", _HYBRID_RETURN_D001_COLUMN)) == 25  # 30 − 5
    assert int(_get(raw, "N04", _HYBRID_RETURN_D001_COLUMN)) == 0   # 零庫存
    assert stats["return_units"] == 25


def test_demand_rows_never_return():
    raw, *_rest, stats, _warn = transform_hybrid(_df(_SCENARIO_BIDIR), bidirectional=True, return_to_d001=True)
    # N31 為需要調貨列（雙向），庫存 47 高於 Safety 30 仍一律不退
    assert str(_get(raw, "N31", "需要調貨")) == "是"
    assert int(_get(raw, "N31", _HYBRID_RETURN_D001_COLUMN)) == 0
    # T32 中銷：70 − 已轉出 9 = 61
    assert int(_get(raw, "T32", _HYBRID_RETURN_D001_COLUMN)) == 61
    assert stats["return_units"] == 61


def test_all_natures_return_regardless_of_direction():
    # 單向：Tourist 列不作調貨供應方，但退D001 按銷量帶逐列判定，與 Nature／方向無關
    raw1, *_rest1, stats1, warn1 = transform_hybrid(_df(_SCENARIO_BIDIR), return_to_d001=True)
    assert "可轉出貨量 (Tourist)" not in raw1.columns
    assert int(_get(raw1, "N31", _HYBRID_RETURN_D001_COLUMN)) == 17  # 47 − 30（非需求列）
    assert int(_get(raw1, "T32", _HYBRID_RETURN_D001_COLUMN)) == 70  # 100 − 30
    assert stats1["return_units"] == 87


def test_nd_rows_can_return():
    rows = [dict(r) for r in _SCENARIO_A]
    for r in rows:
        r["RP Type"] = "ND" if r["Site Code"] == "N01" else "RF"
    raw, *_rest, stats, _warn = transform_hybrid(_df(rows), return_to_d001=True)
    assert str(_get(raw, "N01", "RP Type")) == "ND"
    assert int(_get(raw, "N01", _HYBRID_RETURN_D001_COLUMN)) == 11
    assert stats["return_units"] == 22


def test_return_independent_of_rp_column_and_dimension():
    # 缺「RP Type」欄（全列按 RF）與 rp 維度均不影響退D001 欄
    df_no_rp = _df(_SCENARIO_A).drop(columns=["RP Type"])
    raw, *_rest, stats, _warn = transform_hybrid(df_no_rp, return_to_d001=True)
    assert int(_get(raw, "N01", _HYBRID_RETURN_D001_COLUMN)) == 11
    raw_rp, *_rest_rp, stats_rp, _warn2 = transform_hybrid(_df(_SCENARIO_A), dimension="rp", return_to_d001=True)
    assert int(_get(raw_rp, "N01", _HYBRID_RETURN_D001_COLUMN)) == 11
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
    assert _HYBRID_RETURN_D001_COLUMN in ordered
    ret = _build_hybrid_return_d001(raw)
    assert list(ret.columns) == _HYBRID_RETURN_D001_COLUMNS
    # Safety Qty（來源欄 Adj. Safety Stock，顯示名「Safety Stock (現時)」）緊隨店舖庫存之後
    assert ret.columns.get_loc("Adj. Safety Stock") == ret.columns.get_loc("Shop Stock Qty (CM)") + 1
    assert float(ret.loc[ret["Site Code"] == "N01", "Adj. Safety Stock"].iloc[0]) == 10.0
    assert list(ret["Site Code"]) == ["M01", "N01"]  # 按 Site Code 排序
    assert list(ret[_HYBRID_RETURN_D001_GROSS_COLUMN]) == [11, 20]  # 可退額度（未扣轉出）
    assert list(ret["轉出量"]) == [0, 9]
    assert list(ret[_HYBRID_RETURN_D001_COLUMN]) == [11, 11]
    # 關閉或全 0 時輸出空表
    raw_off = transform_hybrid(_df(_SCENARIO_A))[0]
    assert _build_hybrid_return_d001(raw_off).empty
    raw_empty, *_rest_e, _stats_e, _warn_e = transform_hybrid(_df(_SCENARIO_EMPTY), return_to_d001=True)
    assert _build_hybrid_return_d001(raw_empty).empty


def test_excel_ten_sheets_and_headers():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _warn = transform_hybrid(
        _df(_SCENARIO_A), return_to_d001=True)
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == [
        "說明", "RAW", "Summary", "調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)",
        "店舖問題 SKU", _HYBRID_RETURN_D001_SHEET, "Review 摘要",
        "Safety 設定建議 (全列)",
        "簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)",
    ]
    ws = wb[_HYBRID_RETURN_D001_SHEET]
    headers = [c.value for c in next(ws.iter_rows(min_row=3, max_row=3))]
    # 表頭經 _display_safety_header 轉顯示名（與店舖問題 SKU 同一慣例）
    assert headers == [_display_safety_header(c) for c in _HYBRID_RETURN_D001_COLUMNS]
    assert headers[10] == "Safety Stock (現時)"
    # 明細兩列（M01 / N01）
    data_rows = [r for r in ws.iter_rows(min_row=4, min_col=5, max_col=5) if r[0].value]
    assert len(data_rows) == 2


def test_excel_empty_return_sheet_still_present_when_on():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _warn = transform_hybrid(
        _df(_SCENARIO_EMPTY), return_to_d001=True)
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
    assert [v for _, v, *_ in ret_row] == [2, 1, 22, 9]
    # stats 缺失的回退路徑：由 RAW 欄位存在推定
    fb = _hybrid_stats_from_raw(raw)
    assert fb["return_to_d001"] is True
    assert fb["return_units"] == 22 and fb["return_rows"] == 2
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
