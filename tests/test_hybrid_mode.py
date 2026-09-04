"""Validation tests for Safety Stock x NetStock Review mode (調貨事件診斷 Safety 設定, v1.12.0).

以調貨事件為入口、14 日周轉線為基準診斷 Safety Stock 設定：
- 調貨段與 transform_stock_checking 同規則（SKU 級加總 TurnDay、Nature 供應門檻、14 日目標、NT→Mix）
- 需要調入列：Adj. Safety Stock < 14日周轉所需量 → Safety 設定不足
- 已標示 SKU 的 Mix/NT 有貨列：Adj. Safety Stock > 所需量 → Safety 設定過高
- 建議 Safety 設定值 = 所需量；Safety 建議調整量 = 建議值 − 現有設定（正=增加、負=減少）
- Shop Nature：Other/空白/大小寫差異一律正規化為 Non Tourist
- 輸出 3 工作表：RAW / 調貨診斷 Safety Stock / 診斷摘要
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
    transform_stock_checking,
    validate_hybrid_input,
    _hybrid_dashboard_metrics,
    _hybrid_stats_from_raw,
    _HYBRID_DIAG_COLUMNS,
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


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


# 標準三列場景（SKU 300）：T01 需調入且設定不足；N01 實際轉出且設定過高；M01 可轉出未使用且設定正常
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


def _get(raw, site, col):
    return raw.loc[raw["Site Code"] == site, col].iloc[0]


# ──────────────────────────────────────────────────────────────────────────────
# 1. 輸入檢核與基礎呼叫
# ──────────────────────────────────────────────────────────────────────────────
def test_missing_required_column_rejected():
    df = pd.DataFrame([{"SKU": 1, "Shop Stock Status": "Normal"}])
    missing = validate_hybrid_input(df)
    assert "Shop Nature" in missing
    assert "Sold Qty (P90D)" in missing
    assert "Adj. Safety Stock" in missing
    assert "Sold Qty (CM - MTD)" not in validate_hybrid_input(_df([{"SKU": 1}]).drop(columns=["Sold Qty (CM - MTD)"]))
    try:
        transform_hybrid(df)
        raise AssertionError("應 raise ValueError")
    except ValueError as e:
        assert "Shop Nature" in str(e)


def test_empty_after_clean_raises():
    df = pd.DataFrame([{"Buyer": np.nan, "Brand": "A", "SKU": 1, "Shop Stock Status": "Normal",
                        "Shop Nature": "Tourist", "Site Code": "S01", "Shop Stock Qty (CM)": 5.0,
                        "Sold Qty (P90D)": 90.0, "Adj. Safety Stock": 3.0,
                        "Shop Monthly Sold Qty (P90D)": 30.5}])
    try:
        transform_hybrid(df)
        raise AssertionError("應 raise ValueError")
    except ValueError:
        pass


def test_default_args_produce_results():
    raw, diagnosis, diag_summary, stats, warns = transform_hybrid(_df(_SCENARIO_A))
    assert stats["need_rows"] == 1
    assert len(warns) == 0
    assert "事件類型" in raw.columns and "Safety 診斷" in raw.columns
    assert len(diagnosis) == 3
    assert not diag_summary.empty


# ──────────────────────────────────────────────────────────────────────────────
# 2. 調貨段與 Stock Checking 等價
# ──────────────────────────────────────────────────────────────────────────────
def test_transfer_columns_equivalent_to_stock_checking():
    df = _df(_SCENARIO_A)
    sc_raw, _, sc_stats, _ = transform_stock_checking(df)
    hy_raw, _, _, hy_stats, _ = transform_hybrid(df)
    cmp_cols = ["需要調貨", "調入需求量 (補到14日)", "實際調入量", "調入缺口",
                "可轉出貨量 (Mix)", "可轉出貨量 (Non Tourist)", "轉出量", "調貨來源 (Site Code)",
                "Tourist 加總 TurnDay", "Mix 加總 TurnDay", "Non Tourist 加總 TurnDay"]
    for c in cmp_cols:
        assert list(hy_raw[c]) == list(sc_raw[c]), c
    for k in ("nt_transfer", "mix_transfer", "total_demand", "total_actual", "total_gap"):
        assert hy_stats[k] == sc_stats[k], k


# ──────────────────────────────────────────────────────────────────────────────
# 3. 診斷規則（14 日周轉線基準）
# ──────────────────────────────────────────────────────────────────────────────
def test_tourist_under_setting_diagnosis():
    raw, diagnosis, _, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    assert int(_get(raw, "T01", "14日周轉所需量")) == 56  # ceil(360/90×14)
    assert _get(raw, "T01", "事件類型") == "需要調入"
    assert _get(raw, "T01", "Safety 診斷") == "Safety 設定不足"
    assert int(_get(raw, "T01", "建議 Safety 設定值")) == 56
    assert float(_get(raw, "T01", "Safety 建議調整量")) == 26
    assert stats["under_rows"] == 1 and stats["under_skus"] == 1 and stats["under_inc_total"] == 26


def test_donor_over_setting_diagnosis():
    raw, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    assert _get(raw, "N01", "事件類型") == "實際轉出"
    assert _get(raw, "N01", "Safety 診斷") == "Safety 設定過高"  # Adj 30 > 所需 14
    assert int(_get(raw, "N01", "建議 Safety 設定值")) == 14
    assert float(_get(raw, "N01", "Safety 建議調整量")) == -16
    assert stats["over_rows"] == 1 and stats["over_dec_total"] == -16


def test_donor_boundary_equal_is_normal():
    raw, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    assert _get(raw, "M01", "事件類型") == "可轉出未使用"  # 可轉出 11、轉出 0
    assert _get(raw, "M01", "Safety 診斷") == "設定正常—過剩非設定因素"  # Adj 14 = 所需 14
    assert pd.isna(_get(raw, "M01", "建議 Safety 設定值"))
    assert pd.isna(_get(raw, "M01", "Safety 建議調整量"))
    assert stats["normal_rows"] == 1


def test_tourist_normal_setting_when_covered():
    rows = [dict(r) for r in _SCENARIO_A]
    rows[0].update({"SKU": 500, "Site Code": "T03", "Adj. Safety Stock": 60})
    rows[1].update({"SKU": 500, "Site Code": "N03", "Adj. Safety Stock": 10})
    raw, _, _, _, _ = transform_hybrid(_df([rows[0], rows[1]]))
    assert _get(raw, "T03", "Safety 診斷") == "設定正常—缺貨非設定因素"  # 60 ≥ 56
    assert pd.isna(_get(raw, "T03", "Safety 建議調整量"))
    assert _get(raw, "N03", "Safety 診斷") == "設定正常—過剩非設定因素"  # 10 ≤ 14


def test_no_sale_donor_suggests_zero():
    rows = [
        {"SKU": 400, "Site Code": "T02", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 180.0, "Shop Stock Qty (CM)": 0.0,
         "Shop Monthly Sold Qty (P90D)": 61.0, "Adj. Safety Stock": 10,
         "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 400, "Site Code": "N02", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 0.0, "Shop Stock Qty (CM)": 20.0,
         "Shop Monthly Sold Qty (P90D)": 0.0, "Adj. Safety Stock": 5},
    ]
    raw, _, _, stats, _ = transform_hybrid(_df(rows))
    assert _get(raw, "N02", "事件類型") == "實際轉出"
    assert int(_get(raw, "N02", "14日周轉所需量")) == 0
    assert _get(raw, "N02", "Safety 診斷") == "Safety 設定過高"
    assert int(_get(raw, "N02", "建議 Safety 設定值")) == 0
    assert float(_get(raw, "N02", "Safety 建議調整量")) == -5
    assert stats["total_gap"] == 8  # 需求 28 − NT 全數可轉出 20


def test_blank_adj_safety_treated_as_zero_with_warning():
    rows = [
        {"SKU": 500, "Site Code": "T05", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": None,
         "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 500, "Site Code": "N05", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30},
    ]
    raw, _, _, _, warns = transform_hybrid(_df(rows))
    assert _get(raw, "T05", "Safety 診斷") == "Safety 設定不足"  # 0 < 14
    assert int(_get(raw, "T05", "建議 Safety 設定值")) == 14
    assert any("Adj. Safety Stock" in w for w in warns)


def test_non_event_rows_blank():
    raw, _, _, _, _ = transform_hybrid(_df(_SCENARIO_A))
    t02 = _df([{"SKU": 999, "Site Code": "Z01", "Shop Nature": "Non Tourist",
                "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
                "Shop Monthly Sold Qty (P90D)": 30.5}])
    raw2, diag2, _, _, _ = transform_hybrid(pd.concat([_df(_SCENARIO_A), t02], ignore_index=True))
    assert _get(raw2, "Z01", "事件類型") == ""
    assert pd.isna(_get(raw2, "Z01", "14日周轉所需量"))
    assert pd.isna(_get(raw2, "Z01", "調貨後庫存"))
    assert "Z01" not in set(diag2["Site Code"])


def test_unflagged_sku_surplus_not_in_diagnosis():
    extra = _df([{"SKU": 600, "Site Code": "M04", "Shop Nature": "Mix",
                  "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 40.0,
                  "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30}])
    raw, diagnosis, _, _, _ = transform_hybrid(pd.concat([_df(_SCENARIO_A), extra], ignore_index=True))
    assert int(_get(raw, "M04", "可轉出貨量 (Mix)")) == 26  # RAW 仍計算可轉出
    assert "M04" not in set(diagnosis["Site Code"])  # 但 SKU 600 未標示調貨 → 不入診斷主表


# ──────────────────────────────────────────────────────────────────────────────
# 4. 上下文欄位（調貨後庫存 / TurnDay）
# ──────────────────────────────────────────────────────────────────────────────
def test_stock_after_and_turnday_context():
    raw, _, _, _, _ = transform_hybrid(_df(_SCENARIO_A))
    assert int(_get(raw, "T01", "調貨後庫存")) == 56  # 47 + 9
    assert int(_get(raw, "N01", "調貨後庫存")) == 21  # 30 − 9
    assert abs(float(_get(raw, "N01", "調貨後 TurnDay")) - 21.0) < 1e-9
    assert pd.isna(_get(raw, "M01", "建議 Safety 設定值"))  # 正常列不給建議


# ──────────────────────────────────────────────────────────────────────────────
# 5. Shop Nature 正規化
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
    raw, _, _, stats, warns = transform_hybrid(df)
    assert (raw.loc[raw["Site Code"] == "O01", "Shop Nature"] == "Non Tourist").all()
    assert (raw.loc[raw["Site Code"] == "O02", "Shop Nature"] == "Non Tourist").all()
    assert (raw.loc[raw["Site Code"] == "M01", "Shop Nature"] == "Mix").all()
    assert stats["mapped_other_rows"] == 1
    assert any("Shop Nature" in w for w in warns)


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
    raw, _, _, stats, _ = transform_hybrid(df)
    assert (raw["Shop Nature"] == "Tourist").any()
    assert stats["need_rows"] == 1


def test_blank_nature_mapped_to_non_tourist():
    df = _df([
        {"SKU": 1, "Site Code": "X01", "Shop Nature": np.nan,
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, stats, warns = transform_hybrid(df)
    assert (raw["Shop Nature"] == "Non Tourist").all()
    assert stats["mapped_other_rows"] == 1
    assert any("Shop Nature" in w for w in warns)


# ──────────────────────────────────────────────────────────────────────────────
# 6. 觸發與 D001 fallback
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
    raw, diagnosis, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 0
    assert (raw["需要調貨"] == "").all()
    assert diagnosis.empty


def test_d001_column_fallback():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Turnday (P90D Qty)": 13.0, "D001 Stock Qty (CM)": 50.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)"])
    _, _, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 1


def test_missing_d001_warns():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Turnday (P90D Qty)": 13.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)"])
    _, _, _, stats, warns = transform_hybrid(df)
    assert stats["need_rows"] == 1
    assert any("D001" in w for w in warns)


# ──────────────────────────────────────────────────────────────────────────────
# 7. 診斷主表 / 交叉表 / stats
# ──────────────────────────────────────────────────────────────────────────────
def test_diagnosis_table_rows_and_order():
    raw, diagnosis, _, _, _ = transform_hybrid(_df(_SCENARIO_A))
    assert list(diagnosis["Site Code"]) == ["T01", "N01", "M01"]  # 需要調入 → 實際轉出 → 可轉出未使用
    assert list(diagnosis["事件類型"]) == ["需要調入", "實際轉出", "可轉出未使用"]
    assert int(diagnosis.loc[diagnosis["Site Code"] == "M01", "可轉出貨量"].iloc[0]) == 11
    for c in _HYBRID_DIAG_COLUMNS:
        assert c in diagnosis.columns


def test_diag_summary_cross_and_totals():
    raw, diagnosis, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    assert list(diag_summary["Safety 診斷"])[:1] == ["Safety 設定不足"]
    assert (diag_summary["Safety 診斷"] == "總計").any()
    under = diag_summary[diag_summary["Safety 診斷"] == "Safety 設定不足"].iloc[0]
    assert int(under["Tourist 列數"]) == 1 and int(under["合計 列數"]) == 1
    assert int(under["合計 建議調整量"]) == 26
    over = diag_summary[diag_summary["Safety 診斷"] == "Safety 設定過高"].iloc[0]
    assert int(over["Non Tourist 列數"]) == 1 and int(over["合計 建議調整量"]) == -16
    total = diag_summary[diag_summary["Safety 診斷"] == "總計"].iloc[0]
    assert int(total["合計 列數"]) == len(diagnosis)


def test_stats_fallback_from_raw_matches():
    raw, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    fb = _hybrid_stats_from_raw(raw)
    for k in ("need_rows", "need_skus", "total_demand", "total_actual", "total_gap",
              "nt_transfer", "mix_transfer", "donor_rows", "under_rows", "under_inc_total",
              "over_rows", "over_dec_total", "normal_rows", "total_raw_rows"):
        assert fb[k] == stats[k], k
    assert _hybrid_dashboard_metrics(raw, stats) == _hybrid_dashboard_metrics(raw, None)


def test_no_candidates_warns_but_completes():
    df = _df([
        {"SKU": 1, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, diagnosis, diag_summary, stats, warns = transform_hybrid(df)
    assert stats["need_rows"] == 0
    assert diagnosis.empty and diag_summary.empty
    assert any("沒有符合調貨條件" in w for w in warns)
    buf = generate_excel_hybrid(raw, diagnosis, diag_summary, stats)
    assert isinstance(buf, BytesIO)


# ──────────────────────────────────────────────────────────────────────────────
# 8. Excel 契約（3 工作表）
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_excel_sheets_and_headers():
    raw, diagnosis, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    buf = generate_excel_hybrid(raw, diagnosis, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == ["RAW", "調貨診斷 Safety Stock", "診斷摘要"]
    raw_headers = [c.value for c in next(wb["RAW"].iter_rows(min_row=1, max_row=1))]
    for col in ["事件類型", "Safety 診斷", "14日周轉所需量", "建議 Safety 設定值",
                "Safety 建議調整量", "調貨後庫存", "調貨後 TurnDay", "需要調貨"]:
        assert col in raw_headers
    assert "D001 Effective Stock Qty (CM)" in raw_headers
    diag_headers = [c.value for c in next(wb["調貨診斷 Safety Stock"].iter_rows(min_row=1, max_row=1))]
    for col in ["Brand", "SKU", "Site Code", "Shop Nature", "事件類型", "Safety 診斷",
                "建議 Safety 設定值", "Safety 建議調整量", "可轉出貨量", "調貨後庫存"]:
        assert col in diag_headers


def test_generate_excel_dashboard_matches_stats():
    raw, diagnosis, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    buf = generate_excel_hybrid(raw, diagnosis, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    ws = wb["診斷摘要"]
    vals = {}
    for header_row, value_row in ((3, 4), (7, 8), (11, 12)):
        for c in range(1, 5):
            vals[ws.cell(row=header_row, column=c).value] = ws.cell(row=value_row, column=c).value
    assert vals["需調貨 SKU 數"] == stats["need_skus"]
    assert vals["需要調入列數"] == stats["need_rows"]
    assert vals["轉出來源列數"] == stats["donor_rows"]
    assert vals["調入需求量"] == stats["total_demand"]
    assert vals["實際調入量"] == stats["total_actual"]
    assert vals["無法足額缺口"] == stats["total_gap"]
    assert vals["Non Tourist 轉出量"] == stats["nt_transfer"]
    assert vals["Mix 轉出量"] == stats["mix_transfer"]
    assert vals["Safety 設定不足列數"] == stats["under_rows"]
    assert vals["不足建議增加合計"] == stats["under_inc_total"]
    assert vals["Safety 設定過高列數"] == stats["over_rows"]
    assert vals["過高建議減少合計"] == stats["over_dec_total"]


def test_excel_diagnosis_sheet_freeze_and_autofilter():
    raw, diagnosis, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    buf = generate_excel_hybrid(raw, diagnosis, diag_summary, stats)
    wb = load_workbook(buf)
    ws = wb["調貨診斷 Safety Stock"]
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref is not None


def test_hybrid_excel_progress_monotonic():
    raw, diagnosis, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    events = []
    generate_excel_hybrid(raw, diagnosis, diag_summary, stats,
                          progress_callback=lambda p, _m: events.append(float(p)))
    assert all(b >= a for a, b in zip(events, events[1:])), events


# ──────────────────────────────────────────────────────────────────────────────
# 9. 入檔品質回歸（canonicalization / fallback / 數值檢查）
# ──────────────────────────────────────────────────────────────────────────────
def test_header_whitespace_canonicalized():
    df = _df(_SCENARIO_A)
    df = df.rename(columns={c: c + " " for c in df.columns})
    assert validate_hybrid_input(df) == []
    raw, diagnosis, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 1
    assert len(diagnosis) == 3


def test_duplicate_canonical_header_rejected():
    df = _df([{"SKU": 1}])
    df["Buyer "] = df["Buyer"]
    try:
        transform_hybrid(df)
        raise AssertionError("應 raise ValueError")
    except ValueError as e:
        assert "重複" in str(e) or "Buyer" in str(e)


def test_d001_row_level_fallback():
    df = _df([
        {"SKU": 1, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5,
         "D001 Net Stock Qty (CM)": None, "D001 Stock Qty (CM)": 200.0},
        {"SKU": 1, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, stats, warns = transform_hybrid(df)
    assert stats["need_rows"] == 0
    assert int(_get(raw, "T01", "D001 Effective Stock Qty (CM)")) == 200
    assert any("D001 Net Stock Qty (CM)" in w and "空白" in w for w in warns)


def test_monthly_blank_row_level_fallback():
    df = _df([
        {"SKU": 2, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": None,
         "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 2, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, stats, warns = transform_hybrid(df)
    assert stats["need_rows"] == 1
    assert abs(float(_get(raw, "T01", "Shop Monthly Sold Qty (P90D)")) - 122.0) < 1e-9
    assert any("空白" in w and "換算" in w for w in warns)


def test_non_integer_stock_rejected():
    df = _df([
        {"SKU": 1, "Site Code": "T01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 0.5,
         "Shop Monthly Sold Qty (P90D)": 30.5, "D001 Net Stock Qty (CM)": 10.0},
    ])
    try:
        transform_hybrid(df)
        raise AssertionError("應 raise ValueError")
    except ValueError as e:
        assert "整數" in str(e)


def test_sku_desc_all_blank_derived():
    df = _df([{**r, "Item_Shop": f"Prod-{r['Site Code']}-ShopHK"} for r in _SCENARIO_A])
    df["SKU Desc"] = None
    raw, diagnosis, _, _, _ = transform_hybrid(df)
    assert not diagnosis.empty
    assert (diagnosis["SKU Desc"].fillna("").astype(str).str.rstrip("-") == "Prod").all()
