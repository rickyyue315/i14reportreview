"""Validation tests for Safety Stock x NetStock Review mode (調貨事件 Review Safety 設定, v1.12.0/v1.13.0).

以調貨事件為入口 Review Safety Stock 設定；Review 基準為**按 Nature 的 Safety 基準天數**（預設 7/9/10，可調）：
- 調貨段與 transform_stock_checking 同規則（SKU 級加總 TurnDay、Nature 供應門檻、14 日目標、NT→Mix）；
  調貨段恆用 14 日，**不受 safety_days 影響**
- 需要調入列：Adj. Safety Stock < Safety 周轉所需量 → Safety 設定不足
- 已標示 SKU 的 Mix/NT 有貨列：Adj. Safety Stock > 所需量 → Safety 設定過高
- 建議 Safety 設定值 = 所需量；Safety 建議調整量 = 建議值 − 現有設定（正=增加、負=減少）
- 調貨分兩變體：不限OM（契約欄位）與同OM（來源店與需求店同 OM，欄位加「(同OM)」後綴）；
  OM 限制只影響分配，Safety Review 與建議調整兩變體相同
- Shop Nature：Other/空白/大小寫差異一律正規化為 Non Tourist
- 輸出 7 工作表：RAW / Summary / 調貨 Review Safety Stock (不限OM) / 調貨 Review Safety Stock (同OM) / Review 摘要 / 簡易調貨明細 (不限OM) / 簡易調貨明細 (同OM)
- 「Review 摘要」的三個統計區塊同母體（事件類型非空的列）：① Safety Review × Shop Nature 交叉表、
  ② Safety 調整方向分類（增加／減少／不變）、③ Safety 調整明細（SKU 級彙總）；②③ 於 v1.16.0 新增，
  由 RAW 即時彙算，不受 OM 限制影響，且不改動 transform_hybrid／generate_excel_hybrid 簽名
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
    _hybrid_raw_display_columns,
    _HYBRID_DIAG_COLUMNS,
    _HYBRID_TRANSFER_DETAIL_COLUMNS,
    _HYBRID_SAFETY_CHANGE_COLUMNS,
    _HYBRID_SAFETY_SKU_COLUMNS,
    _HYBRID_SAFETY_LINE_COLUMN,
    _HYBRID_SAFETY_DAYS_COLUMN,
    _HYBRID_SAFETY_DAYS_DEFAULT,
    _HYBRID_SAFETY_DAYS_MIN,
    _HYBRID_SAFETY_DAYS_MAX,
    _hybrid_safety_days,
    _hybrid_safety_days_label,
    _build_hybrid_transfer_detail,
    _build_hybrid_safety_change_summary,
    _build_hybrid_safety_sku_detail,
    _build_hybrid_diag_summary,
    _hybrid_allocate_transfers,
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


# 標準三列場景（SKU 300，全部同 OM）：T01 需調入且設定不足；N01 實際轉出且設定過高；M01 可轉出未使用且設定正常
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

# 跨 OM 場景（SKU 310）：T11 (Ivy) 需調入；N11 (Max) NT 有貨但異 OM；M11 (Ivy) Mix 有貨同 OM
_SCENARIO_OM = [
    {"SKU": 310, "Site Code": "T11", "OM": "Ivy", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
     "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 310, "Site Code": "N11", "OM": "Max", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30},
    {"SKU": 310, "Site Code": "M11", "OM": "Ivy", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 14},
]


def _get(raw, site, col):
    return raw.loc[raw["Site Code"] == site, col].iloc[0]


# 三個 Nature 統一設為合法上限 10 日（舊 _SD14 的 14 日已超出 7–10 建議範圍）
_SD10 = {"Tourist": 10, "Mix": 10, "Non Tourist": 10}


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
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, warns = transform_hybrid(_df(_SCENARIO_A))
    assert stats["need_rows"] == 1
    assert len(warns) == 0
    assert "事件類型" in raw.columns and "Safety Review" in raw.columns
    assert "轉出量 (同OM)" in raw.columns
    assert len(diagnosis) == 3 and len(diagnosis_om) == 3
    assert not diag_summary.empty
    assert not sc_summary.empty and "存貨偏重提示" in sc_summary.columns


# ──────────────────────────────────────────────────────────────────────────────
# 2. 調貨段與 Stock Checking 等價（不限OM＝契約欄位）
# ──────────────────────────────────────────────────────────────────────────────
def test_transfer_columns_equivalent_to_stock_checking():
    df = _df(_SCENARIO_A)
    sc_raw, sc_sum, sc_stats, _ = transform_stock_checking(df)
    hy_raw, hy_sc, _, _, _, hy_stats, _ = transform_hybrid(df)
    cmp_cols = ["需要調貨", "調入需求量 (補到14日)", "實際調入量", "調入缺口",
                "可轉出貨量 (Mix)", "可轉出貨量 (Non Tourist)", "轉出量", "調貨來源 (Site Code)",
                "Tourist 加總 TurnDay", "Mix 加總 TurnDay", "Non Tourist 加總 TurnDay"]
    for c in cmp_cols:
        assert list(hy_raw[c]) == list(sc_raw[c]), c
    for k in ("nt_transfer", "mix_transfer", "total_demand", "total_actual", "total_gap"):
        assert hy_stats[k] == sc_stats[k], k
    # Summary（含存貨偏重提示）與 Stock Checking 完全同格式同值
    assert list(hy_sc.columns) == list(sc_sum.columns)
    for c in ("強制調貨提示", "存貨偏重提示", "Non Tourist 加總 - Shop Stock Qty (CM)"):
        assert list(hy_sc[c].fillna("")) == list(sc_sum[c].fillna("")), c


# ──────────────────────────────────────────────────────────────────────────────
# 3. Review 規則（Safety 基準天數；_SD10 為合法上限 10 日口徑）
# ──────────────────────────────────────────────────────────────────────────────
def test_tourist_under_setting_diagnosis():
    raw, sc_summary, diagnosis, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A10), safety_days=_SD10)
    assert int(_get(raw, "T01", "Safety 周轉所需量")) == 40  # ceil(360/90×10)
    assert _get(raw, "T01", "事件類型") == "需要調入"
    assert _get(raw, "T01", "Safety Review") == "Safety 設定不足"
    assert int(_get(raw, "T01", "建議 Safety 設定值")) == 40
    assert float(_get(raw, "T01", "Safety 建議調整量")) == 10
    assert stats["under_rows"] == 1 and stats["under_skus"] == 1 and stats["under_inc_total"] == 10


def test_donor_over_setting_diagnosis():
    raw, sc_summary, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A10), safety_days=_SD10)
    assert _get(raw, "N01", "事件類型") == "實際轉出"
    assert _get(raw, "N01", "Safety Review") == "Safety 設定過高"  # Adj 12 > 所需 10
    assert int(_get(raw, "N01", "建議 Safety 設定值")) == 10
    assert float(_get(raw, "N01", "Safety 建議調整量")) == -2
    assert stats["over_rows"] == 1 and stats["over_dec_total"] == -2


# 10 日口徑專用場景：Tourist 需調入判不足；N01 判過高；M01 以 Adj 10 對齊 Mix 所需量作邊界正常列
_SCENARIO_A10 = [
    {"SKU": 300, "Site Code": "T01", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
     "Turnday (P90D Qty)": 10.0, "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 300, "Site Code": "N01", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 12},
    {"SKU": 300, "Site Code": "M01", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
]


def test_donor_boundary_equal_is_normal():
    rows = [dict(r) for r in _SCENARIO_A10]
    raw, sc_summary, _, _, _, stats, _ = transform_hybrid(_df(rows), safety_days=_SD10)
    assert _get(raw, "M01", "事件類型") == "可轉出未使用"  # 可轉出 11、轉出 0
    assert _get(raw, "M01", "Safety Review") == "設定正常—過剩非設定因素"  # Adj 10 = 所需 10
    assert pd.isna(_get(raw, "M01", "建議 Safety 設定值"))
    assert pd.isna(_get(raw, "M01", "Safety 建議調整量"))
    assert stats["normal_rows"] == 1


def test_tourist_normal_setting_when_covered():
    rows = [dict(r) for r in _SCENARIO_A]
    rows[0].update({"SKU": 500, "Site Code": "T03", "Adj. Safety Stock": 60})
    rows[1].update({"SKU": 500, "Site Code": "N03", "Adj. Safety Stock": 10})
    raw, sc_summary, _, _, _, _, _ = transform_hybrid(_df([rows[0], rows[1]]))
    assert _get(raw, "T03", "Safety Review") == "設定正常—缺貨非設定因素"  # 預設 Tourist 7 日：60 ≥ 28
    assert pd.isna(_get(raw, "T03", "Safety 建議調整量"))
    assert _get(raw, "N03", "Safety Review") == "設定正常—過剩非設定因素"  # 預設 Non Tourist 10 日：10 ≤ 10


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
    raw, sc_summary, _, _, _, stats, _ = transform_hybrid(_df(rows))
    assert _get(raw, "N02", "事件類型") == "實際轉出"
    assert int(_get(raw, "N02", "Safety 周轉所需量")) == 0
    assert _get(raw, "N02", "Safety Review") == "Safety 設定過高"
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
    raw, sc_summary, _, _, _, _, warns = transform_hybrid(_df(rows), safety_days=_SD10)
    assert _get(raw, "T05", "Safety Review") == "Safety 設定不足"  # 0 < 10
    assert int(_get(raw, "T05", "建議 Safety 設定值")) == 10
    assert any("Adj. Safety Stock" in w for w in warns)


def test_non_event_rows_blank():
    t02 = _df([{"SKU": 999, "Site Code": "Z01", "Shop Nature": "Non Tourist",
                "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
                "Shop Monthly Sold Qty (P90D)": 30.5}])
    raw2, sc_summary, diag2, _, _, _, _ = transform_hybrid(pd.concat([_df(_SCENARIO_A), t02], ignore_index=True))
    assert _get(raw2, "Z01", "事件類型") == ""
    assert pd.isna(_get(raw2, "Z01", "Safety 周轉所需量"))
    assert pd.isna(_get(raw2, "Z01", "調貨後庫存"))
    assert pd.isna(_get(raw2, "Z01", "調貨後庫存 (同OM)"))
    assert "Z01" not in set(diag2["Site Code"])


def test_unflagged_sku_surplus_not_in_diagnosis():
    extra = _df([{"SKU": 600, "Site Code": "M04", "Shop Nature": "Mix",
                  "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 40.0,
                  "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30}])
    raw, sc_summary, diagnosis, diagnosis_om, _, _, _ = transform_hybrid(pd.concat([_df(_SCENARIO_A), extra], ignore_index=True))
    assert int(_get(raw, "M04", "可轉出貨量 (Mix)")) == 26  # RAW 仍計算可轉出
    assert "M04" not in set(diagnosis["Site Code"])  # 但 SKU 600 未標示調貨 → 不入 Review 主表
    assert "M04" not in set(diagnosis_om["Site Code"])


# ──────────────────────────────────────────────────────────────────────────────
# 4. 同OM調貨變體
# ──────────────────────────────────────────────────────────────────────────────
def test_same_om_donor_switch():
    raw, sc_summary, diagnosis, diagnosis_om, _, stats, _ = transform_hybrid(_df(_SCENARIO_OM))
    # 不限OM：NT 優先 → N11 (Max) 轉出；同OM：只可由同 OM 的 M11 (Ivy) 轉出
    assert _get(raw, "N11", "轉出量") == 9 and _get(raw, "M11", "轉出量") == 0
    assert _get(raw, "N11", "轉出量 (同OM)") == 0 and _get(raw, "M11", "轉出量 (同OM)") == 9
    assert _get(raw, "T11", "調貨來源 (同OM)") == "M11:9"
    assert int(_get(raw, "T11", "實際調入量 (同OM)")) == 9 and int(_get(raw, "T11", "調入缺口 (同OM)")) == 0
    assert stats["nt_transfer"] == 9 and stats["mix_transfer"] == 0
    assert stats["nt_transfer_om"] == 0 and stats["mix_transfer_om"] == 9
    assert stats["donated_rows_om"] == 1
    # 事件類型兩版不同：不限OM N11=實際轉出、M11=可轉出未使用；同OM 相反
    assert _get(diagnosis, "N11", "事件類型") == "實際轉出"
    assert _get(diagnosis, "M11", "事件類型") == "可轉出未使用"
    assert _get(diagnosis_om, "N11", "事件類型") == "可轉出未使用"
    assert _get(diagnosis_om, "M11", "事件類型") == "實際轉出"
    # 同OM主表以契約欄名輸出該變體數值
    assert int(_get(diagnosis_om, "M11", "轉出量")) == 9
    assert int(_get(diagnosis_om, "M11", "調貨後庫存")) == 16


def test_same_om_gap_when_no_same_om_donor():
    rows = [
        {"SKU": 320, "Site Code": "T21", "OM": "Ivy", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
         "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 320, "Site Code": "N21", "OM": "Max", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30},
    ]
    raw, sc_summary, _, _, _, stats, _ = transform_hybrid(_df(rows))
    assert int(_get(raw, "T21", "實際調入量")) == 9 and int(_get(raw, "T21", "調入缺口")) == 0
    assert int(_get(raw, "T21", "實際調入量 (同OM)")) == 0
    assert int(_get(raw, "T21", "調入缺口 (同OM)")) == 9
    assert stats["total_actual_om"] == 0 and stats["total_gap_om"] == 9
    assert stats["donated_rows_om"] == 0


def test_om_restriction_does_not_alter_safety_diagnosis():
    raw, sc_summary, diagnosis, diagnosis_om, _, _, _ = transform_hybrid(_df(_SCENARIO_OM))
    diag_cols = ["Safety Review", "建議 Safety 設定值", "Safety 建議調整量", "Safety 周轉所需量"]
    a = diagnosis.set_index("Site Code")[diag_cols].sort_index()
    b = diagnosis_om.set_index("Site Code")[diag_cols].sort_index()
    pd.testing.assert_frame_equal(a, b)
    # 入表列集合相同（來源列以可轉出貨量>0 判定，與 OM 分配無關）
    assert sorted(diagnosis["Site Code"]) == sorted(diagnosis_om["Site Code"])


def test_missing_om_column_warns_and_equals_free():
    df = _df(_SCENARIO_A).drop(columns=["OM"])
    raw, sc_summary, diagnosis, diagnosis_om, _, stats, warns = transform_hybrid(df)
    assert any("OM" in w for w in warns)
    assert (raw["OM"] == "(未填OM)").all()  # 缺欄時 RAW／主表補「(未填OM)」群組
    assert list(diagnosis["轉出量"]) == list(diagnosis_om["轉出量"])
    assert stats["nt_transfer_om"] == stats["nt_transfer"]


def test_blank_om_grouped_as_unfilled():
    rows = [dict(r) for r in _SCENARIO_OM]
    rows[2]["OM"] = None  # M11 OM 空白 → (未填OM)，與 T11 (Ivy) 不同組 → 同OM無來源
    raw, sc_summary, _, _, _, stats, warns = transform_hybrid(_df(rows))
    assert any("OM" in w and "空白" in w for w in warns)
    assert int(_get(raw, "T11", "實際調入量 (同OM)")) == 0
    assert int(_get(raw, "M11", "轉出量 (同OM)")) == 0
    assert stats["total_gap_om"] == 9


# ──────────────────────────────────────────────────────────────────────────────
# 5. 上下文欄位（調貨後庫存 / TurnDay）
# ──────────────────────────────────────────────────────────────────────────────
def test_stock_after_and_turnday_context():
    raw, sc_summary, _, _, _, _, _ = transform_hybrid(_df(_SCENARIO_A10), safety_days=_SD10)
    assert int(_get(raw, "T01", "調貨後庫存")) == 56  # 47 + 9
    assert int(_get(raw, "N01", "調貨後庫存")) == 21  # 30 − 9
    assert abs(float(_get(raw, "N01", "調貨後 TurnDay")) - 21.0) < 1e-9
    # 全同 OM 時兩變體相同
    assert int(_get(raw, "T01", "調貨後庫存 (同OM)")) == 56
    assert int(_get(raw, "N01", "調貨後庫存 (同OM)")) == 21
    assert pd.isna(_get(raw, "M01", "建議 Safety 設定值"))  # M01 邊界正常列不給建議


# ──────────────────────────────────────────────────────────────────────────────
# 6. Shop Nature 正規化
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
    raw, sc_summary, _, _, _, stats, warns = transform_hybrid(df)
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
    raw, sc_summary, _, _, _, stats, _ = transform_hybrid(df)
    assert (raw["Shop Nature"] == "Tourist").any()
    assert stats["need_rows"] == 1


def test_blank_nature_mapped_to_non_tourist():
    df = _df([
        {"SKU": 1, "Site Code": "X01", "Shop Nature": np.nan,
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, sc_summary, _, _, _, stats, warns = transform_hybrid(df)
    assert (raw["Shop Nature"] == "Non Tourist").all()
    assert stats["mapped_other_rows"] == 1
    assert any("Shop Nature" in w for w in warns)


# ──────────────────────────────────────────────────────────────────────────────
# 7. 觸發與 D001 fallback
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
    raw, sc_summary, diagnosis, diagnosis_om, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 0
    assert (raw["需要調貨"] == "").all()
    assert diagnosis.empty and diagnosis_om.empty


def test_d001_column_fallback():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Turnday (P90D Qty)": 13.0, "D001 Stock Qty (CM)": 50.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)"])
    _, sc_summary, _, _, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 1


def test_missing_d001_warns():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Turnday (P90D Qty)": 13.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)"])
    _, sc_summary, _, _, _, stats, warns = transform_hybrid(df)
    assert stats["need_rows"] == 1
    assert any("D001" in w for w in warns)


# ──────────────────────────────────────────────────────────────────────────────
# 8. Review 主表 / 交叉表 / stats
# ──────────────────────────────────────────────────────────────────────────────
def test_diagnosis_table_rows_and_order():
    raw, sc_summary, diagnosis, _, _, _, _ = transform_hybrid(_df(_SCENARIO_A))
    assert list(diagnosis["Site Code"]) == ["T01", "N01", "M01"]  # 需要調入 → 實際轉出 → 可轉出未使用
    assert list(diagnosis["事件類型"]) == ["需要調入", "實際轉出", "可轉出未使用"]
    assert list(diagnosis["OM"]) == ["OM1", "OM1", "OM1"]  # 主表含 OM 欄供同OM對照
    assert int(diagnosis.loc[diagnosis["Site Code"] == "M01", "可轉出貨量"].iloc[0]) == 11
    for c in _HYBRID_DIAG_COLUMNS:
        assert c in diagnosis.columns


def test_diag_summary_cross_and_totals():
    raw, sc_summary, diagnosis, _, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A10), safety_days=_SD10)
    assert list(diag_summary["Safety Review"])[:1] == ["Safety 設定不足"]
    assert (diag_summary["Safety Review"] == "總計").any()
    under = diag_summary[diag_summary["Safety Review"] == "Safety 設定不足"].iloc[0]
    assert int(under["Tourist 列數"]) == 1 and int(under["合計 列數"]) == 1
    assert int(under["合計 建議調整量"]) == 10
    over = diag_summary[diag_summary["Safety Review"] == "Safety 設定過高"].iloc[0]
    assert int(over["Non Tourist 列數"]) == 1 and int(over["合計 建議調整量"]) == -2
    total = diag_summary[diag_summary["Safety Review"] == "總計"].iloc[0]
    assert int(total["合計 列數"]) == len(diagnosis)


def test_stats_fallback_from_raw_matches():
    raw, sc_summary, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_OM))
    fb = _hybrid_stats_from_raw(raw)
    for k in ("need_rows", "need_skus", "total_demand", "total_actual", "total_gap",
              "nt_transfer", "mix_transfer", "total_actual_om", "total_gap_om",
              "nt_transfer_om", "mix_transfer_om",
              "donor_rows", "under_rows", "under_inc_total",
              "over_rows", "over_dec_total", "normal_rows", "total_raw_rows"):
        assert fb[k] == stats[k], k
    assert _hybrid_dashboard_metrics(raw, stats) == _hybrid_dashboard_metrics(raw, None)


def test_no_candidates_warns_but_completes():
    df = _df([
        {"SKU": 1, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, warns = transform_hybrid(df)
    assert stats["need_rows"] == 0
    assert diagnosis.empty and diagnosis_om.empty and diag_summary.empty
    assert any("沒有符合調貨條件" in w for w in warns)
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    assert isinstance(buf, BytesIO)


# ──────────────────────────────────────────────────────────────────────────────
# 9. Excel 契約（7 工作表）
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_excel_sheets_and_headers():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == [
        "RAW", "Summary", "調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)", "Review 摘要",
        "簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)",
    ]
    raw_headers = [c.value for c in next(wb["RAW"].iter_rows(min_row=1, max_row=1))]
    for col in ["事件類型", "Safety Review", "Safety 周轉所需量", "建議 Safety 設定值",
                "Safety 建議調整量", "調貨後庫存", "調貨後 TurnDay",
                "實際調入量 (同OM)", "轉出量 (同OM)", "調貨來源 (同OM)",
                "調貨後庫存 (同OM)", "調貨後 TurnDay (同OM)", "需要調貨",
                "D001 Effective Stock Qty (CM)"]:
        assert col in raw_headers
    for sheet in ("調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)"):
        headers = [c.value for c in next(wb[sheet].iter_rows(min_row=1, max_row=1))]
        for col in ["Brand", "SKU", "Site Code", "OM", "Shop Nature", "事件類型", "Safety Review",
                    "建議 Safety 設定值", "Safety 建議調整量", "可轉出貨量", "轉出量", "調貨後庫存"]:
            assert col in headers, (sheet, col)
    sc_headers = [c.value for c in next(wb["Summary"].iter_rows(min_row=5, max_row=5))]
    assert len(sc_headers) == 26
    for col in ["Brand", "SKU", "強制調貨提示", "由Non Tourist 店舖預計調出數量", "存貨偏重提示"]:
        assert col in sc_headers


def test_generate_excel_summary_skew_highlight():
    # 偏重列：Z 提示欄與偏重 Nature 庫存佔比欄（Tourist=Q）以醒目橘底 FFC000 加強顯示
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_BIDIR_B1), bidirectional=True)
    assert str(sc_summary.iloc[0]["存貨偏重提示"]).startswith("存貨偏重 Tourist")
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    ws = load_workbook(buf)["Summary"]
    assert str(ws.cell(row=6, column=26).value).startswith("存貨偏重 Tourist")
    assert str(ws.cell(row=6, column=26).fill.start_color.rgb).endswith("FFC000")
    assert ws.cell(row=6, column=26).font.bold
    assert str(ws.cell(row=6, column=17).fill.start_color.rgb).endswith("FFC000")  # Tourist 庫存 %
    assert not str(ws.cell(row=6, column=19).fill.start_color.rgb).endswith("FFC000")  # Mix 未偏重
    assert not str(ws.cell(row=6, column=21).fill.start_color.rgb).endswith("FFC000")  # NT 未偏重


def test_hybrid_raw_display_columns_helper():
    df = pd.DataFrame(columns=["Buyer", "SKU", "Item_Shop", "Shop Stock Value (CM)", "事件類型", "SKU Desc"])
    cols = _hybrid_raw_display_columns(df)
    assert "Item_Shop" not in cols and "Shop Stock Value (CM)" not in cols
    assert cols == ["Buyer", "SKU", "SKU Desc", "Shop Name", "事件類型"]  # SKU Desc 去重；Shop Name 為衍生欄


def test_generate_excel_raw_view_columns():
    df = _df(_SCENARIO_A)
    df["Shop Stock Value (CM)"] = 999.0
    df["Excess Stock Value"] = 1.0
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(df)
    assert "Item_Shop" in raw.columns  # RAW dataframe 保留原欄，只在工作表層面移除
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    ws = load_workbook(buf, read_only=True)["RAW"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    assert "Item_Shop" not in headers
    assert not any("value" in str(h).lower() for h in headers)
    idx = {h: i for i, h in enumerate(headers)}
    # 可讀欄序：識別 → 庫存與銷售 → SKU 級判定 → Review 結果 → 調貨 (不限OM) → 調貨 (同OM)
    assert idx["Site Code"] < idx["Shop Stock Qty (CM)"] < idx["需要調貨"]
    assert idx["需要調貨"] < idx["事件類型"] < idx["Safety Review"] < idx["建議 Safety 設定值"] < idx["Safety 建議調整量"]
    assert idx["Safety 建議調整量"] < idx["實際調入量"] < idx["調貨來源 (Site Code)"] < idx["調貨後庫存"]
    assert idx["調貨後 TurnDay"] < idx["實際調入量 (同OM)"] < idx["調貨後 TurnDay (同OM)"]
    assert len(headers) == len(set(headers))


def test_generate_excel_dashboard_matches_stats():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_OM))
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    ws = wb["Review 摘要"]
    vals = {}
    for header_row, value_row in ((3, 4), (7, 8), (11, 12), (15, 16)):
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
    assert vals["實際調入量 (同OM)"] == stats["total_actual_om"]
    assert vals["無法足額缺口 (同OM)"] == stats["total_gap_om"]
    assert vals["Non Tourist 轉出量 (同OM)"] == stats["nt_transfer_om"]
    assert vals["Mix 轉出量 (同OM)"] == stats["mix_transfer_om"]


def test_excel_diagnosis_sheet_freeze_and_autofilter():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf)
    for sheet in ("調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)"):
        ws = wb[sheet]
        assert ws.freeze_panes == "A2"
        assert ws.auto_filter.ref is not None


def test_hybrid_excel_progress_monotonic():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    events = []
    generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats,
                          progress_callback=lambda p, _m: events.append(float(p)))
    assert all(b >= a for a, b in zip(events, events[1:])), events


# ──────────────────────────────────────────────────────────────────────────────
# 10. 入檔品質回歸（canonicalization / fallback / 數值檢查）
# ──────────────────────────────────────────────────────────────────────────────
def test_header_whitespace_canonicalized():
    df = _df(_SCENARIO_A)
    df = df.rename(columns={c: c + " " for c in df.columns})
    assert validate_hybrid_input(df) == []
    raw, sc_summary, diagnosis, diagnosis_om, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 1
    assert len(diagnosis) == 3 and len(diagnosis_om) == 3


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
    raw, sc_summary, _, _, _, stats, warns = transform_hybrid(df)
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
    raw, sc_summary, _, _, _, stats, warns = transform_hybrid(df)
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
    raw, sc_summary, diagnosis, _, _, _, _ = transform_hybrid(df)
    assert not diagnosis.empty
    assert (diagnosis["SKU Desc"].fillna("").astype(str).str.rstrip("-") == "Prod").all()


# ──────────────────────────────────────────────────────────────────────────────
# 11. 重複列與自我調貨防護（同店同 SKU 資料異常）
# ──────────────────────────────────────────────────────────────────────────────
# S01 對 SKU 500 同時出現 Tourist 與 Non Tourist 列（重複列＋Nature 不一致的異常輸入）。
# 修復前：S01-Tourist 由 S01-Non Tourist 調入（自我調貨），S01 同一 SKU 同時調入與調出。
# 修復後：同店來源一律排除——S01-Tourist 缺口 9 件（無其他來源），S02 正常由 S01 調入 9 件。
_SCENARIO_DUP = [
    {"SKU": 500, "Site Code": "S01", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 500, "Site Code": "S01", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5},
    {"SKU": 500, "Site Code": "S02", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "D001 Net Stock Qty (CM)": 10.0},
]


def _pick(raw, site, nat):
    return raw[(raw["Site Code"] == site) & (raw["Shop Nature"] == nat)].iloc[0]


def test_duplicate_site_sku_conflicting_nature_warned():
    raw, _, _, _, _, _, warns = transform_hybrid(_df(_SCENARIO_DUP))
    assert any("Site Code × SKU" in w and "重複" in w for w in warns)
    assert any("Shop Nature" in w and "不一致" in w for w in warns)


def test_duplicate_same_nature_warns_dup_only():
    rows = [
        {"SKU": 510, "Site Code": "S01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 510, "Site Code": "S01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0, "D001 Net Stock Qty (CM)": 10.0},
    ]
    raw, _, _, _, _, _, warns = transform_hybrid(_df(rows))
    assert any("Site Code × SKU" in w and "重複" in w for w in warns)
    assert not any("不一致" in w for w in warns)


def test_self_transfer_guarded_and_no_site_both_in_and_out():
    raw, _, _, _, _, _, _ = transform_hybrid(_df(_SCENARIO_DUP))
    t01 = _pick(raw, "S01", "Tourist")
    t02 = raw[raw["Site Code"] == "S02"].iloc[0]
    n01 = _pick(raw, "S01", "Non Tourist")
    # 自我調貨被排除：S01-Tourist 無其他來源 → 全數缺口
    assert int(t01["實際調入量"]) == 0 and int(t01["調入缺口"]) == 9
    assert str(t01["調貨來源 (Site Code)"]) == ""
    assert int(t01["實際調入量 (同OM)"]) == 0 and int(t01["調入缺口 (同OM)"]) == 9
    # S02 正常由 S01 調入
    assert int(t02["實際調入量"]) == 9 and str(t02["調貨來源 (Site Code)"]) == "S01:9"
    assert int(t02["實際調入量 (同OM)"]) == 9 and str(t02["調貨來源 (同OM)"]) == "S01:9"
    assert int(n01["轉出量"]) == 9 and int(n01["轉出量 (同OM)"]) == 9
    # 店級：同一 SKU + 同一店舖不得同時調入與調出（兩變體）
    g = raw.groupby(["SKU", "Site Code"])[["實際調入量", "轉出量", "實際調入量 (同OM)", "轉出量 (同OM)"]].sum()
    assert not ((g["實際調入量"] > 0) & (g["轉出量"] > 0)).any()
    assert not ((g["實際調入量 (同OM)"] > 0) & (g["轉出量 (同OM)"] > 0)).any()
    # 列級與來源欄：任何需求列的調貨來源都不含自己店號
    for _, r in raw.iterrows():
        for col in ("調貨來源 (Site Code)", "調貨來源 (同OM)"):
            src_sites = {p.split(":")[0] for p in str(r[col]).split("; ") if p}
            assert r["Site Code"] not in src_sites
    # 守恆：總調入 = 總轉出
    assert int(raw["實際調入量"].sum()) == int(raw["轉出量"].sum())
    assert int(raw["實際調入量 (同OM)"].sum()) == int(raw["轉出量 (同OM)"].sum())


# ──────────────────────────────────────────────────────────────────────────────
# 12. 允許雙向調貨（bidirectional=True，v1.14.0）
# ──────────────────────────────────────────────────────────────────────────────
# B1（SKU 800）：存貨偏重遊客區 → Mix 缺貨由 Tourist 逆向轉出
_BIDIR_B1 = [
    {"SKU": 800, "Site Code": "M01", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 3},
    {"SKU": 800, "Site Code": "T01", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30},
]


def test_bidir_mix_short_tourist_donates():
    raw, sc_summary, diagnosis, _, _, stats, _ = transform_hybrid(
        _df(_BIDIR_B1), bidirectional=True, safety_days=_SD10)
    assert _get(raw, "M01", "需要調貨") == "是"
    assert int(_get(raw, "M01", "調入需求量 (補到14日)")) == 9
    assert int(_get(raw, "M01", "實際調入量")) == 9 and int(_get(raw, "M01", "調入缺口")) == 0
    assert _get(raw, "M01", "調貨來源 (Site Code)") == "T01:9"
    assert int(_get(raw, "T01", "可轉出貨量 (Tourist)")) == 16
    assert int(_get(raw, "T01", "轉出量")) == 9
    assert int(_get(raw, "T01", "調貨後庫存")) == 21
    assert stats["tourist_transfer"] == 9 and stats["nt_transfer"] == 0 and stats["mix_transfer"] == 0
    # Review 對稱擴展：Mix 調入列判不足、Tourist 轉出列判過高
    assert _get(raw, "M01", "事件類型") == "需要調入"
    assert _get(raw, "M01", "Safety Review") == "Safety 設定不足"
    assert _get(raw, "T01", "事件類型") == "實際轉出"
    assert _get(raw, "T01", "Safety Review") == "Safety 設定過高"
    assert float(_get(raw, "T01", "Safety 建議調整量")) == -20  # 10 日口徑：建議 10 − 現有 30
    # 主表「可轉出貨量」含 Tourist 欄
    assert int(_get(diagnosis, "T01", "可轉出貨量")) == 16
    # 存貨偏重提示指向 Tourist（使用者場景）
    hint = str(sc_summary.iloc[0]["存貨偏重提示"])
    assert hint.startswith("存貨偏重 Tourist")
    assert int(sc_summary.iloc[0]["由Tourist 店舖預計調出數量"]) == 9


def test_bidir_nt_short_source_order_mix_before_tourist():
    # B2（SKU 801）：NT 缺貨 → 來源固定 NT→Mix→Tourist（NT 自身除外）：Mix 6 件先、Tourist 補 3 件
    rows = [
        {"SKU": 801, "Site Code": "N01", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 3},
        {"SKU": 801, "Site Code": "M02", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 20.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 801, "Site Code": "T02", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ]
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(rows), bidirectional=True)
    assert _get(raw, "N01", "需要調貨") == "是"
    assert int(_get(raw, "N01", "實際調入量")) == 9
    assert _get(raw, "N01", "調貨來源 (Site Code)") == "M02:6; T02:3"
    assert stats["mix_transfer"] == 6 and stats["tourist_transfer"] == 3
    assert int(raw["實際調入量"].sum()) == int(raw["轉出量"].sum())  # 守恆


def test_bidir_two_short_natures_compete_by_demand_desc():
    # B3（SKU 802）：Tourist 與 Mix 同缺、僅 NT 有貨 6 件 → 需求大者（Mix 23）先獲 6 件，Tourist 9 全列缺口
    rows = [
        {"SKU": 802, "Site Code": "T03", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
         "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 802, "Site Code": "M03", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 180.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 61.0, "Adj. Safety Stock": 3,
         "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 802, "Site Code": "N03", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 20.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 14,
         "D001 Net Stock Qty (CM)": 10.0},
    ]
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(rows), bidirectional=True)
    assert stats["need_rows"] == 2
    assert int(_get(raw, "M03", "調入需求量 (補到14日)")) == 23
    assert int(_get(raw, "M03", "實際調入量")) == 6 and int(_get(raw, "M03", "調入缺口")) == 17
    assert int(_get(raw, "T03", "實際調入量")) == 0 and int(_get(raw, "T03", "調入缺口")) == 9
    assert _get(raw, "M03", "調貨來源 (Site Code)") == "N03:6"
    assert stats["total_actual"] == 6 and stats["total_gap"] == 26
    assert stats["nt_transfer"] == 6 and stats["tourist_transfer"] == 0
    # 缺貨 Nature 池低於門檻 → 其超額列可轉出歸零（不得互調）
    assert int(_get(raw, "T03", "可轉出貨量 (Tourist)")) == 0


def test_bidir_self_site_guard_and_om_restriction_on_tourist_donor():
    # B4（SKU 803）：需求 Mix S06 (Ivy)；Tourist 來源 S06（同店，Max）與 S07（Max）
    # 不限OM：同店排除 → S07 轉出 9；同OM：S07 異 OM → 全數缺口
    rows = [
        {"SKU": 803, "Site Code": "S06", "OM": "Ivy", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 3},
        {"SKU": 803, "Site Code": "S06", "OM": "Max", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 803, "Site Code": "S07", "OM": "Max", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ]
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(rows), bidirectional=True)
    mix_rows = raw[raw["Shop Nature"] == "Mix"]
    assert len(mix_rows) == 1 and str(mix_rows.iloc[0]["需要調貨"]) == "是"
    assert int(mix_rows.iloc[0]["實際調入量"]) == 9
    assert str(mix_rows.iloc[0]["調貨來源 (Site Code)"]) == "S07:9"
    assert int(mix_rows.iloc[0]["實際調入量 (同OM)"]) == 0
    assert int(mix_rows.iloc[0]["調入缺口 (同OM)"]) == 9
    assert stats["tourist_transfer"] == 9 and stats["tourist_transfer_om"] == 0
    t07 = raw[raw["Site Code"] == "S07"].iloc[0]
    assert int(t07["轉出量"]) == 9 and int(t07["轉出量 (同OM)"]) == 0


def test_bidir_all_pools_below_threshold_no_flag():
    # B5（SKU 804）：三 Nature 加總 TurnDay 皆 < 門檻 → 無貨可調，不標示＋warning
    rows = [
        {"SKU": 804, "Site Code": "T08", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 804, "Site Code": "M08", "Shop Nature": "Mix",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
        {"SKU": 804, "Site Code": "N08", "Shop Nature": "Non Tourist",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ]
    raw, _, diagnosis, _, _, stats, warns = transform_hybrid(_df(rows), bidirectional=True)
    assert stats["need_rows"] == 0
    assert (raw["需要調貨"] == "").all()
    assert diagnosis.empty
    assert any("沒有符合調貨條件" in w and "Nature" in w for w in warns)


def test_bidir_off_column_contract_and_single_direction_unchanged():
    raw_off, sc_off, _, _, _, stats_off, _ = transform_hybrid(_df(_BIDIR_B1))
    assert "可轉出貨量 (Tourist)" not in raw_off.columns
    assert "由Tourist 店舖預計調出數量" not in sc_off.columns
    assert stats_off["bidirectional"] is False and stats_off["tourist_transfer"] == 0
    assert stats_off["need_rows"] == 0  # 單向：Mix 缺貨不標示
    assert len(_hybrid_dashboard_metrics(raw_off, stats_off)) == 4
    # 單向場景（_SCENARIO_A）開雙向：調貨契約欄數值與單向完全一致（Tourist 缺貨時來源仍是 NT→Mix）
    raw_a_off = transform_hybrid(_df(_SCENARIO_A))[0]
    raw_on, sc_on, _, _, _, stats_on, _ = transform_hybrid(_df(_SCENARIO_A), bidirectional=True)
    cmp_cols = ["需要調貨", "調入需求量 (補到14日)", "實際調入量", "調入缺口",
                "可轉出貨量 (Mix)", "可轉出貨量 (Non Tourist)", "轉出量", "調貨來源 (Site Code)"]
    for c in cmp_cols:
        assert list(raw_on[c]) == list(raw_a_off[c]), c
    assert (raw_on["可轉出貨量 (Tourist)"] == 0).all()  # Tourist 池缺貨 → 歸零
    assert stats_on["tourist_transfer"] == 0
    assert int(sc_on.iloc[0]["由Tourist 店舖預計調出數量"]) == 0


def test_bidir_on_dashboard_and_stats_fallback():
    raw, sc_summary, _, _, _, stats, _ = transform_hybrid(_df(_BIDIR_B1), bidirectional=True)
    assert stats["bidirectional"] is True
    rows = _hybrid_dashboard_metrics(raw, stats)
    assert len(rows) == 5
    assert rows[4][0][0] == "Tourist 轉出量" and rows[4][0][1] == 9
    assert rows[4][1][0] == "Tourist 轉出量 (同OM)" and rows[4][1][1] == 9
    fb = _hybrid_stats_from_raw(raw)
    for k in ("need_rows", "need_skus", "total_demand", "total_actual", "total_gap",
              "nt_transfer", "mix_transfer", "tourist_transfer", "total_actual_om",
              "nt_transfer_om", "tourist_transfer_om", "donor_rows", "under_rows",
              "over_rows", "normal_rows"):
        assert fb[k] == stats[k], k
    assert fb["bidirectional"] is True
    assert _hybrid_dashboard_metrics(raw, stats) == _hybrid_dashboard_metrics(raw, None)


def test_bidir_excel_headers_and_dashboard_rows():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(
        _df(_BIDIR_B1), bidirectional=True)
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == [
        "RAW", "Summary", "調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)", "Review 摘要",
        "簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)",
    ]
    raw_headers = [c.value for c in next(wb["RAW"].iter_rows(min_row=1, max_row=1))]
    assert "可轉出貨量 (Tourist)" in raw_headers
    sc_headers = [c.value for c in next(wb["Summary"].iter_rows(min_row=5, max_row=5))]
    assert len(sc_headers) == 27
    assert sc_headers[26] == "由Tourist 店舖預計調出數量"
    # 單向輸出 A–Z 契約欄序不動
    off_args = transform_hybrid(_df(_BIDIR_B1))
    buf_off = generate_excel_hybrid(*off_args[:5], off_args[5])
    wb_off = load_workbook(buf_off, read_only=True)
    off_headers = [c.value for c in next(wb_off["Summary"].iter_rows(min_row=5, max_row=5))]
    assert len(off_headers) == 26
    assert off_headers == sc_headers[:26]
    ws = wb["Review 摘要"]
    assert ws.cell(row=19, column=1).value == "Tourist 轉出量"
    assert ws.cell(row=20, column=1).value == stats["tourist_transfer"]
    assert ws.cell(row=19, column=2).value == "Tourist 轉出量 (同OM)"
    assert ws.cell(row=23, column=1).value == "Safety Review × Shop Nature 交叉表"
    assert ws.cell(row=24, column=1).value == "Safety Review"


# ──────────────────────────────────────────────────────────────────────────────
# 13. 簡易調貨明細（一列一筆來源店→接收店，來源字串拆行；v1.15.0）
# ──────────────────────────────────────────────────────────────────────────────
# 多來源場景（SKU 820）：T21 (Ivy) 需求 20；NT N21 (Max) 可轉出 6；Mix M21 (Ivy) 可轉出 11
# 不限OM → 來源「N21:6; M21:11」（拆 2 行、缺口 3）；同OM → N21 異 OM 排除，只剩「M21:11」
_SCENARIO_MULTI = [
    {"SKU": 820, "Site Code": "T21", "OM": "Ivy", "Item_Shop": "Prod-T21-ShopT",
     "Shop Nature": "Tourist", "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 36.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
     "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 820, "Site Code": "N21", "OM": "Max", "Item_Shop": "Prod-N21-ShopN",
     "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 20.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30},
    {"SKU": 820, "Site Code": "M21", "OM": "Ivy", "Item_Shop": "Prod-M21-ShopM",
     "Shop Nature": "Mix", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 14},
]

# 無貨可調場景（SKU 830）：NT 池剛過門檻但超額為 0 → 需要調貨=是、來源空、只有缺口
_SCENARIO_NO_SUPPLY = [
    {"SKU": 830, "Site Code": "T31", "Item_Shop": "Prod-T31-ShopT",
     "Shop Nature": "Tourist", "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
     "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 830, "Site Code": "N31", "Item_Shop": "Prod-N31-ShopN",
     "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 14.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30},
]


def test_detail_columns_and_row_splitting():
    raw, *_ = transform_hybrid(_df(_SCENARIO_MULTI))
    assert _get(raw, "T21", "調貨來源 (Site Code)") == "N21:6; M21:11"
    det = _build_hybrid_transfer_detail(raw)
    assert list(det.columns) == _HYBRID_TRANSFER_DETAIL_COLUMNS
    assert len(det) == 2
    assert int(det["調貨數量"].sum()) == int(_get(raw, "T21", "實際調入量")) == 17
    r0 = det.iloc[0]  # 按 Supply Site Code 排序：M21 先於 N21
    assert r0["Supply Site Code(調貨來源)"] == "M21" and int(r0["調貨數量"]) == 11
    assert r0["Target Site Code(接收店舖)"] == "T21"
    assert r0["Target Shop Name"] == "ShopT" and r0["Target OM"] == "Ivy"
    assert r0["Supply Shop Name"] == "ShopM" and r0["Supply OM"] == "Ivy"
    r1 = det.iloc[1]
    assert r1["Supply Site Code(調貨來源)"] == "N21" and int(r1["調貨數量"]) == 6
    assert r1["Supply Shop Name"] == "ShopN" and r1["Supply OM"] == "Max"
    assert det["Brand"].eq("A").all() and det["SKU"].astype(str).eq("820").all()
    assert det["SKU Desc"].eq("Test Product").all()


def test_detail_om_variant():
    raw, *_ = transform_hybrid(_df(_SCENARIO_MULTI))
    assert _get(raw, "T21", "調貨來源 (同OM)") == "M21:11"
    det_om = _build_hybrid_transfer_detail(raw, variant="om")
    assert len(det_om) == 1
    r = det_om.iloc[0]
    assert r["Supply Site Code(調貨來源)"] == "M21" and r["Supply OM"] == "Ivy"
    assert int(r["調貨數量"]) == int(_get(raw, "T21", "實際調入量 (同OM)")) == 11


def test_detail_gap_only_excluded():
    raw, *_ = transform_hybrid(_df(_SCENARIO_NO_SUPPLY))
    assert _get(raw, "T31", "需要調貨") == "是"
    assert _get(raw, "T31", "調貨來源 (Site Code)") == ""
    assert int(_get(raw, "T31", "調入缺口")) == 9
    det = _build_hybrid_transfer_detail(raw)
    assert det.empty and list(det.columns) == _HYBRID_TRANSFER_DETAIL_COLUMNS


def test_detail_merges_duplicate_supply_site():
    raw = pd.DataFrame([{
        "Brand": "A", "SKU": 900, "SKU Desc": "X", "Site Code": "T01",
        "Shop Name": "接收店", "OM": "Ivy", "需要調貨": "是",
        "調貨來源 (Site Code)": "HB30:2; HB30:3; HB49:4", "實際調入量": 9,
    }])
    det = _build_hybrid_transfer_detail(raw)
    assert len(det) == 2
    hb30 = det[det["Supply Site Code(調貨來源)"] == "HB30"]
    assert int(hb30.iloc[0]["調貨數量"]) == 5  # 重複段合併加總
    assert hb30.iloc[0]["Supply Shop Name"] == "" and hb30.iloc[0]["Supply OM"] == ""
    assert det.iloc[0]["Target Shop Name"] == "接收店"


def test_detail_empty_inputs():
    assert list(_build_hybrid_transfer_detail(None).columns) == _HYBRID_TRANSFER_DETAIL_COLUMNS
    assert _build_hybrid_transfer_detail(pd.DataFrame()).empty
    assert _build_hybrid_transfer_detail(pd.DataFrame([{"Brand": "A", "SKU": 1}])).empty


def test_detail_bidirectional_tourist_supply():
    raw, *_ = transform_hybrid(_df(_BIDIR_B1), bidirectional=True)
    det = _build_hybrid_transfer_detail(raw)
    assert len(det) == 1
    r = det.iloc[0]
    assert r["Target Site Code(接收店舖)"] == "M01"
    assert r["Supply Site Code(調貨來源)"] == "T01"
    assert int(r["調貨數量"]) == 9


def test_detail_excel_sheets():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_MULTI))
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf)
    ws = wb["簡易調貨明細 (不限OM)"]
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    assert headers == _HYBRID_TRANSFER_DETAIL_COLUMNS
    assert ws.freeze_panes == "A2"
    assert ws.auto_filter.ref == "A1:J3"
    assert ws.cell(row=2, column=10).number_format == "#,##0"
    assert int(ws.cell(row=2, column=10).value) == 11  # M21:11（排序後首行）
    ws_om = wb["簡易調貨明細 (同OM)"]
    assert ws_om.auto_filter.ref == "A1:J2"
    assert int(ws_om.cell(row=2, column=10).value) == 11
    # 無實際調貨 → 兩表寫「無符合資料」
    raw2, sc2, d2, d2o, c2, s2, _ = transform_hybrid(_df(_SCENARIO_NO_SUPPLY))
    buf2 = generate_excel_hybrid(raw2, sc2, d2, d2o, c2, s2)
    wb2 = load_workbook(buf2)
    assert wb2.sheetnames[-2:] == ["簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)"]
    assert wb2["簡易調貨明細 (不限OM)"].cell(row=2, column=1).value == "無符合資料"
    assert wb2["簡易調貨明細 (同OM)"].cell(row=2, column=1).value == "無符合資料"


# ──────────────────────────────────────────────────────────────────────────────
# 14. Safety 調整統計（② 調整方向分類 ＋ ③ SKU 級調整明細；v1.16.0）
# ──────────────────────────────────────────────────────────────────────────────
# 兩個 SKU 的對照場景：SKU 301 為 SKU 300 的等價複製（10 日口徑：各 1 列不足 +10、1 列過高 −2、1 列正常）
_SAFETY_TWO_SKU = _SCENARIO_A10 + [
    {"SKU": 301, "Site Code": "T12", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
     "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
     "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 301, "Site Code": "N12", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 12},
    {"SKU": 301, "Site Code": "M12", "Shop Nature": "Mix",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
]


def _cat_rows(frame):
    return {str(r["分類"]): r for _, r in frame.iterrows()}


def test_safety_change_summary_matches_diag_stats():
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A10), safety_days=_SD10)
    frame = _build_hybrid_safety_change_summary(raw)
    assert list(frame.columns) == _HYBRID_SAFETY_CHANGE_COLUMNS
    assert list(frame["分類"]) == ["增加", "減少", "不變", "總計"]
    cat = _cat_rows(frame)
    assert int(cat["增加"]["列數"]) == stats["under_rows"] == 1
    assert int(cat["減少"]["列數"]) == stats["over_rows"] == 1
    assert int(cat["不變"]["列數"]) == stats["normal_rows"] == 1
    assert int(cat["增加"]["Safety 建議調整量合計"]) == stats["under_inc_total"] == 10
    assert int(cat["減少"]["Safety 建議調整量合計"]) == stats["over_dec_total"] == -2
    assert int(cat["不變"]["Safety 建議調整量合計"]) == 0
    total = cat["總計"]
    assert int(total["列數"]) == stats["under_rows"] + stats["over_rows"] + stats["normal_rows"]
    assert int(total["Safety 建議調整量合計"]) == stats["under_inc_total"] + stats["over_dec_total"] == 8
    assert all(str(r["說明"]) for _, r in frame.iterrows())


def test_safety_change_summary_total_sku_dedup():
    raw, _, _, _, _, _, _ = transform_hybrid(_df(_SAFETY_TWO_SKU), safety_days=_SD10)
    frame = _build_hybrid_safety_change_summary(raw)
    cat = _cat_rows(frame)
    for label in ("增加", "減少", "不變"):
        assert int(cat[label]["列數"]) == 2 and int(cat[label]["SKU 數"]) == 2
    assert int(cat["總計"]["列數"]) == 6
    # 三類各 2 個 SKU，總計應為母體不重複 SKU 數（2）而非相加（6）
    assert int(cat["總計"]["SKU 數"]) == 2 < 6


def test_safety_sku_detail_contract_and_grouping():
    raw, _, _, _, _, _, _ = transform_hybrid(_df(_SAFETY_TWO_SKU), safety_days=_SD10)
    detail = _build_hybrid_safety_sku_detail(raw)
    assert list(detail.columns) == _HYBRID_SAFETY_SKU_COLUMNS
    assert [str(v) for v in detail["SKU"]] == ["300", "301"]  # 按 Brand / SKU 排序
    assert detail["Brand"].eq("A").all() and detail["Buyer"].eq("B1").all()
    assert detail["SKU Desc"].eq("Test Product").all()
    for _, r in detail.iterrows():
        assert int(r["Review 列數"]) == 3
        assert int(r["需增加店舖數"]) == 1 and int(r["需減少店舖數"]) == 1
        assert int(r["設定正常店舖數"]) == 1
        # 現有／建議合計只累加調整列（30 + 12 → 40 + 10），正常列 M01/M12 不計
        assert int(r["現有 Safety 設定合計（調整列）"]) == 42
        assert int(r["建議 Safety 設定合計（調整列）"]) == 50
        assert int(r["淨調整量"]) == 8


def test_safety_sku_detail_sums_tie_to_raw_and_change_table():
    raw, _, _, _, diag_summary, _, _ = transform_hybrid(_df(_SAFETY_TWO_SKU))
    detail = _build_hybrid_safety_sku_detail(raw)
    chg = pd.to_numeric(raw["Safety 建議調整量"], errors="coerce").fillna(0)
    evt = raw["事件類型"].fillna("").astype(str) != ""
    per_sku = chg[evt].groupby(raw.loc[evt, "SKU"]).sum()
    assert list(detail["淨調整量"]) == [int(v) for _, v in per_sku.items()]
    assert int(detail["淨調整量"].sum()) == int(
        _build_hybrid_safety_change_summary(raw).iloc[-1]["Safety 建議調整量合計"])
    assert int(detail["Review 列數"].sum()) == int(diag_summary.iloc[-1]["合計 列數"])
    assert len(detail) == int(diag_summary.iloc[-1]["合計 SKU 數"])


def test_safety_sku_detail_counts_shops_not_rows():
    raw = pd.DataFrame([{
        "Brand": "A", "SKU": 1, "SKU Desc": "X", "Buyer": "B1", "Site Code": "T01",
        "Shop Nature": "Tourist", "事件類型": "需要調入", "Safety Review": "Safety 設定不足",
        "Adj. Safety Stock": 10.0, "建議 Safety 設定值": 24.0, "Safety 建議調整量": 14.0,
    }, {
        "Brand": "A", "SKU": 1, "SKU Desc": "X", "Buyer": "B1", "Site Code": "T01",
        "Shop Nature": "Tourist", "事件類型": "需要調入", "Safety Review": "Safety 設定不足",
        "Adj. Safety Stock": 10.0, "建議 Safety 設定值": 24.0, "Safety 建議調整量": 14.0,
    }, {
        "Brand": "A", "SKU": 1, "SKU Desc": "X", "Buyer": "B1", "Site Code": "N01",
        "Shop Nature": "Non Tourist", "事件類型": "可轉出未使用",
        "Safety Review": "設定正常—過剩非設定因素",
        "Adj. Safety Stock": 20.0, "建議 Safety 設定值": None, "Safety 建議調整量": None,
    }])
    r = _build_hybrid_safety_sku_detail(raw).iloc[0]
    assert int(r["Review 列數"]) == 3  # 重複列仍逐列計
    assert int(r["需增加店舖數"]) == 1  # 但店舖數按 Site Code 不重複
    assert int(r["設定正常店舖數"]) == 1
    assert int(r["現有 Safety 設定合計（調整列）"]) == 20 and int(r["建議 Safety 設定合計（調整列）"]) == 48
    assert int(r["淨調整量"]) == 28


def test_safety_stats_are_om_independent():
    raw, *_ = transform_hybrid(_df(_SCENARIO_OM))
    with_om = _build_hybrid_safety_sku_detail(raw)
    without = _build_hybrid_safety_sku_detail(
        raw.drop(columns=[c for c in raw.columns if c.endswith(" (同OM)")]))
    pd.testing.assert_frame_equal(with_om, without)
    chg_with = _build_hybrid_safety_change_summary(raw)
    chg_without = _build_hybrid_safety_change_summary(
        raw.drop(columns=[c for c in raw.columns if c.endswith(" (同OM)")]))
    pd.testing.assert_frame_equal(chg_with, chg_without)


def test_safety_stats_bidirectional_population_matches_cross():
    raw, _, _, _, diag_summary, stats, _ = transform_hybrid(_df(_BIDIR_B1), bidirectional=True)
    frame = _build_hybrid_safety_change_summary(raw)
    detail = _build_hybrid_safety_sku_detail(raw)
    assert int(frame.iloc[-1]["列數"]) == int(diag_summary.iloc[-1]["合計 列數"]) == stats["donor_rows"] + stats["need_rows"]
    assert len(detail) == int(diag_summary.iloc[-1]["合計 SKU 數"])
    assert int(frame.iloc[-1]["Safety 建議調整量合計"]) == int(diag_summary.iloc[-1]["合計 建議調整量"])


def test_safety_stats_empty_inputs():
    for frame, cols in ((_build_hybrid_safety_change_summary(None), _HYBRID_SAFETY_CHANGE_COLUMNS),
                        (_build_hybrid_safety_change_summary(pd.DataFrame()), _HYBRID_SAFETY_CHANGE_COLUMNS),
                        (_build_hybrid_safety_change_summary(pd.DataFrame([{"Brand": "A", "SKU": 1}])),
                         _HYBRID_SAFETY_CHANGE_COLUMNS),
                        (_build_hybrid_safety_sku_detail(None), _HYBRID_SAFETY_SKU_COLUMNS),
                        (_build_hybrid_safety_sku_detail(pd.DataFrame([{"Brand": "A", "SKU": 1}])),
                         _HYBRID_SAFETY_SKU_COLUMNS)):
        assert frame.empty and list(frame.columns) == cols


def test_excel_safety_blocks_appended_after_cross():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(
        _df(_SAFETY_TWO_SKU), safety_days=_SD10)
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf)
    assert wb.sheetnames == [
        "RAW", "Summary", "調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)", "Review 摘要",
        "簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)",
    ]  # 7 表契約不變，新統計只在「Review 摘要」內延伸
    ws = wb["Review 摘要"]
    # ① 交叉表位置完全未變（單向：標題 19 / 表頭 20 / 資料 21–25）
    assert ws.cell(row=19, column=1).value == "Safety Review × Shop Nature 交叉表"
    assert ws.cell(row=20, column=1).value == "Safety Review"
    assert ws.cell(row=25, column=1).value == "總計"
    assert ws.cell(row=27, column=1).value == "Safety 調整方向分類（增加 / 減少 / 不變）"
    assert [ws.cell(row=r, column=1).value for r in range(28, 33)] == \
        ["分類", "增加", "減少", "不變", "總計"]
    assert int(ws.cell(row=29, column=4).value) == 20  # +10 × 2 SKU
    assert int(ws.cell(row=30, column=4).value) == -4  # −2 × 2 SKU
    assert int(ws.cell(row=32, column=4).value) == 16
    assert ws.cell(row=32, column=5).value == "SKU 數為全表不重複 SKU（同一 SKU 可跨分類）；合計為淨調整量"
    assert ws.cell(row=34, column=1).value == "Safety 調整明細（SKU 級彙總）"
    headers = [c.value for c in next(ws.iter_rows(min_row=35, max_row=35, min_col=1, max_col=11))]
    assert headers == _HYBRID_SAFETY_SKU_COLUMNS
    assert [ws.cell(row=r, column=2).value for r in (36, 37)] == ["300", "301"]
    assert int(ws.cell(row=36, column=11).value) == 8
    assert ws.cell(row=36, column=11).number_format == "#,##0"


def test_excel_safety_blocks_bidirectional_position():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(
        _df(_BIDIR_B1), bidirectional=True)
    ws = load_workbook(generate_excel_hybrid(
        raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats))["Review 摘要"]
    assert ws.cell(row=23, column=1).value == "Safety Review × Shop Nature 交叉表"
    assert ws.cell(row=24, column=1).value == "Safety Review"
    assert ws.cell(row=31, column=1).value == "Safety 調整方向分類（增加 / 減少 / 不變）"
    assert ws.cell(row=32, column=1).value == "分類"
    assert ws.cell(row=38, column=1).value == "Safety 調整明細（SKU 級彙總）"
    assert ws.cell(row=39, column=1).value == "Brand"


def test_excel_safety_blocks_empty_population():
    solo = _df([{"SKU": 999, "Site Code": "Z01", "Shop Nature": "Non Tourist",
                 "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
                 "Shop Monthly Sold Qty (P90D)": 30.5}])
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(solo)
    assert diag_summary.empty
    ws = load_workbook(generate_excel_hybrid(
        raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats))["Review 摘要"]
    assert ws.cell(row=20, column=1).value == "無符合資料"


# ──────────────────────────────────────────────────────────────────────────────
# 15. Safety 基準天數按 Nature（與調貨 14 日線脫鉤；v1.17.0）
# ──────────────────────────────────────────────────────────────────────────────
_TRANSFER_COLS = ["調入需求量 (補到14日)", "實際調入量", "調入缺口", "轉出量",
                  "調貨來源 (Site Code)", "可轉出貨量 (Mix)", "可轉出貨量 (Non Tourist)",
                  "實際調入量 (同OM)", "轉出量 (同OM)", "調貨後庫存", "調貨後 TurnDay", "需要調貨"]


def test_safety_days_defaults_and_range():
    assert _HYBRID_SAFETY_DAYS_DEFAULT == {"Tourist": 7, "Mix": 9, "Non Tourist": 10}
    assert (_HYBRID_SAFETY_DAYS_MIN, _HYBRID_SAFETY_DAYS_MAX) == (7, 10)
    assert _hybrid_safety_days_label(_HYBRID_SAFETY_DAYS_DEFAULT) == \
        "Tourist 7 日 / Mix 9 日 / Non Tourist 10 日"


def test_safety_days_normalizer_ignores_invalid():
    assert _hybrid_safety_days(None) == (_HYBRID_SAFETY_DAYS_DEFAULT, [])
    assert _hybrid_safety_days({}) == (_HYBRID_SAFETY_DAYS_DEFAULT, [])
    assert _hybrid_safety_days({"Mix": 10}) == ({"Tourist": 7, "Mix": 10, "Non Tourist": 10}, [])
    assert _hybrid_safety_days({" Tourist ": "8"}) == ({"Tourist": 8, "Mix": 9, "Non Tourist": 10}, [])
    # 未知 Nature、非數值、非正數一律忽略（退回預設，不拋錯）
    assert _hybrid_safety_days({"VIP": 5, "Mix": None, "Tourist": 0, "Non Tourist": -3,
                                "Bad": "x"}) == (_HYBRID_SAFETY_DAYS_DEFAULT, [])


def test_safety_days_out_of_range_clamped_with_warning():
    days, warns = _hybrid_safety_days({"Tourist": 4, "Mix": 11, "Non Tourist": 10})
    assert days == {"Tourist": 7, "Mix": 10, "Non Tourist": 10}
    assert any("Tourist" in w and "箝制" in w for w in warns)
    assert any("Mix" in w and "箝制" in w for w in warns)
    # 箝制 warning 會併入 transform_hybrid 的 warnings
    _, _, _, _, _, _, hybrid_warns = transform_hybrid(
        _df(_SCENARIO_A), safety_days={"Tourist": 4})
    assert any("箝制" in w for w in hybrid_warns)


def test_safety_days_non_integer_rejected():
    for bad in (7.9, "9.5", True, np.inf):
        try:
            _hybrid_safety_days({"Mix": bad})
            raise AssertionError("應拒絕非整數 Safety 基準天數")
        except ValueError:
            pass


def test_safety_line_follows_nature_days():
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A))  # 預設 7 / 9 / 10
    assert stats["safety_days"] == {"Tourist": 7, "Mix": 9, "Non Tourist": 10}
    # T01 Tourist：ceil(360/90 × 7) = 28，現有 30 → 不再判不足（14 日口徑下會判不足）
    assert int(_get(raw, "T01", _HYBRID_SAFETY_DAYS_COLUMN)) == 7
    assert int(_get(raw, "T01", _HYBRID_SAFETY_LINE_COLUMN)) == 28
    assert _get(raw, "T01", "Safety Review") == "設定正常—缺貨非設定因素"
    assert pd.isna(_get(raw, "T01", "Safety 建議調整量"))
    # N01 Non Tourist：ceil(90/90 × 10) = 10，現有 30 → 過高 −20
    assert int(_get(raw, "N01", _HYBRID_SAFETY_DAYS_COLUMN)) == 10
    assert int(_get(raw, "N01", _HYBRID_SAFETY_LINE_COLUMN)) == 10
    assert int(_get(raw, "N01", "建議 Safety 設定值")) == 10
    assert float(_get(raw, "N01", "Safety 建議調整量")) == -20
    # M01 Mix：ceil(90/90 × 9) = 9，現有 14 → 原本「等於 14 邊界＝正常」，現改為過高 −5
    assert int(_get(raw, "M01", _HYBRID_SAFETY_DAYS_COLUMN)) == 9
    assert _get(raw, "M01", "Safety Review") == "Safety 設定過高"
    assert float(_get(raw, "M01", "Safety 建議調整量")) == -5
    assert stats["under_rows"] == 0 and stats["over_rows"] == 2 and stats["normal_rows"] == 1
    assert stats["over_dec_total"] == -25 and stats["under_inc_total"] == 0


def test_transfers_unchanged_by_safety_days():
    """Safety 基準天數只能改 Review，調貨段的數字必須逐一相同。"""
    raw_a, _, diag_a, diag_om_a, _, stats_a, _ = transform_hybrid(_df(_SCENARIO_A))
    raw_b, _, diag_b, diag_om_b, _, stats_b, _ = transform_hybrid(_df(_SCENARIO_A), safety_days=_SD10)
    for c in _TRANSFER_COLS:
        assert list(raw_a[c]) == list(raw_b[c]), c
        assert c not in diag_a.columns or list(diag_a[c]) == list(diag_b[c]), c
    for key in ("need_rows", "need_skus", "total_demand", "total_actual", "total_gap",
                "nt_transfer", "mix_transfer", "total_actual_om", "total_gap_om",
                "nt_transfer_om", "mix_transfer_om", "donor_rows", "donated_rows",
                "donor_unused_rows", "total_raw_rows"):
        assert stats_a[key] == stats_b[key], key
    # 同OM 主表把「轉出量 (同OM)」改名為「轉出量」（_build_hybrid_diagnosis variant="om"）
    assert list(diag_om_a["轉出量"]) == list(diag_om_b["轉出量"])
    assert list(diag_om_a["實際調入量"]) == list(diag_om_b["實際調入量"])


def test_safety_days_echoed_in_raw_and_recovered_by_fallback():
    raw, _, diagnosis, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A), safety_days={"Mix": 8})
    assert stats["safety_days"] == {"Tourist": 7, "Mix": 8, "Non Tourist": 10}
    assert raw[_HYBRID_SAFETY_DAYS_COLUMN].dtype.kind in "iu"
    for site, day in (("T01", 7), ("M01", 8), ("N01", 10)):
        assert int(_get(raw, site, _HYBRID_SAFETY_DAYS_COLUMN)) == day
        assert int(_get(diagnosis, site, _HYBRID_SAFETY_DAYS_COLUMN)) == day
    assert _hybrid_stats_from_raw(raw)["safety_days"] == stats["safety_days"]
    assert _hybrid_dashboard_metrics(raw, stats) == _hybrid_dashboard_metrics(raw, None)


def test_main_table_contract_names_safety_columns():
    assert _HYBRID_SAFETY_LINE_COLUMN in _HYBRID_DIAG_COLUMNS
    assert _HYBRID_SAFETY_DAYS_COLUMN in _HYBRID_DIAG_COLUMNS
    assert "14日周轉所需量" not in _HYBRID_DIAG_COLUMNS
    raw, _, diagnosis, _, _, _, _ = transform_hybrid(_df(_SCENARIO_A))
    assert list(diagnosis.columns) == _HYBRID_DIAG_COLUMNS  # 主表 25 欄契約完整
    display = _hybrid_raw_display_columns(raw)
    assert _HYBRID_SAFETY_LINE_COLUMN in display and _HYBRID_SAFETY_DAYS_COLUMN in display
    assert display.index(_HYBRID_SAFETY_DAYS_COLUMN) < display.index(_HYBRID_SAFETY_LINE_COLUMN)


def test_excel_summary_title_shows_applied_basis():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(
        _df(_SCENARIO_A), safety_days={"Tourist": 10, "Mix": 8, "Non Tourist": 7})
    ws = load_workbook(generate_excel_hybrid(
        raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats))["Review 摘要"]
    assert "Safety 基準 Tourist 10 日 / Mix 8 日 / Non Tourist 7 日" in str(ws.cell(row=1, column=1).value)


# ──────────────────────────────────────────────────────────────────────────────
# 16. 輸入驗證邊界（負數／非有限／月銷 0／D001 parity／OM 全空白／必要欄位）
# ──────────────────────────────────────────────────────────────────────────────
def _expect_hybrid_value_error(df, column):
    try:
        transform_hybrid(df)
        raise AssertionError(f"應拒絕欄位：{column}")
    except ValueError as e:
        assert column in str(e)


def test_hybrid_does_not_require_unused_shop_stock_status():
    df = _df(_SCENARIO_A).drop(columns=["Shop Stock Status"])
    assert validate_hybrid_input(df) == []
    raw, _, _, _, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 1 and not raw.empty


def test_negative_stock_and_safety_rejected():
    for column in ("Shop Stock Qty (CM)", "Adj. Safety Stock"):
        df = _df(_SCENARIO_A)
        df.loc[df["Site Code"] == "T01", column] = -1
        _expect_hybrid_value_error(df, column)


def test_nonfinite_stock_sales_and_safety_rejected():
    for column in ("Shop Stock Qty (CM)", "Sold Qty (P90D)",
                   "Shop Monthly Sold Qty (P90D)", "Adj. Safety Stock"):
        df = _df(_SCENARIO_A)
        df[column] = df[column].astype(float)
        df.loc[df["Site Code"] == "T01", column] = np.inf
        _expect_hybrid_value_error(df, column)


def test_negative_sales_treated_as_zero_with_warning():
    df = _df(_SCENARIO_A)
    df.loc[df["Site Code"] == "N01", "Sold Qty (P90D)"] = -10
    raw, _, _, _, _, _, warns = transform_hybrid(df)
    assert any("Sold Qty (P90D)" in w and "負數" in w for w in warns)
    # N01 由可轉出 16 件變為全數釋放 30 件（負銷售不再放大成超額以外的供應）
    assert int(_get(raw, "N01", "可轉出貨量 (Non Tourist)")) == 30
    assert int(_get(raw, "N01", _HYBRID_SAFETY_LINE_COLUMN)) == 0


def test_monthly_zero_with_positive_p90d_recovers():
    df = _df(_SCENARIO_A)
    df.loc[df["Site Code"] == "T01", "Shop Monthly Sold Qty (P90D)"] = 0
    raw, _, _, _, _, stats, warns = transform_hybrid(df)
    assert any("為 0 但 P90D 有銷售" in w for w in warns)
    assert stats["need_rows"] == 1  # T01 不再因 TurnDay=inf 被擋掉
    assert np.isfinite(float(_get(raw, "T01", "Tourist 加總 TurnDay")))


def test_d001_effective_value_shown_in_summary():
    df = _df(_SCENARIO_A)
    df.loc[df["Site Code"] == "T01", "D001 Net Stock Qty (CM)"] = np.nan
    df.loc[df["Site Code"] == "T01", "D001 Stock Qty (CM)"] = 200
    raw, summary, _, _, _, _, _ = transform_hybrid(df)
    assert int(_get(raw, "T01", "D001 Effective Stock Qty (CM)")) == 200
    row = summary.loc[summary["SKU"] == "300"].iloc[0]
    assert int(row["D001 Net Stock Qty (CM)"]) == 200


def test_om_all_blank_warns():
    df = _df(_SCENARIO_A)
    df["OM"] = None
    _, _, _, _, _, _, warns = transform_hybrid(df)
    assert any("OM" in w and "全欄空白" in w for w in warns)


# ──────────────────────────────────────────────────────────────────────────────
# 17. 審查修正回歸（v1.16.x）：Effective 月銷稽核欄／交叉表四捨五入／
#     需求 Nature 自身除外守衛／重複列 warning 補自我調貨守衛說明
# ──────────────────────────────────────────────────────────────────────────────
EFF_MS = "Effective Shop Monthly Sold Qty (P90D)"


def test_effective_monthly_sold_column_added_and_original_kept():
    df = _df(_SCENARIO_A)
    df.loc[df["Site Code"] == "T01", "Shop Monthly Sold Qty (P90D)"] = 0
    raw, _, _, _, _, _, _ = transform_hybrid(df)
    # 回補值（T01：120×30.5=3660 → 122.0）存 Effective 欄；原欄維持覆寫後值（契約）
    assert float(_get(raw, "T01", EFF_MS)) == 122.0
    assert float(_get(raw, "T01", "Shop Monthly Sold Qty (P90D)")) == 122.0
    # 未回補列（N01 月銷 30.5 正常）Effective 欄與原值相同
    assert float(_get(raw, "N01", EFF_MS)) == 30.5
    # RAW 顯示欄序含 Effective 欄
    assert EFF_MS in _hybrid_raw_display_columns(raw)


def test_cross_table_sums_round_not_truncate():
    # 小數調整量：cross 表合計須四捨五入（-0.5→-1、+0.5→0），與分類表／stats 同口徑
    rows = [
        {"SKU": 300, "Site Code": "T01", "Shop Nature": "Tourist", "OM": "OM1",
         "Sold Qty (P90D)": 270.0, "Shop Stock Qty (CM)": 20.0,
         "Shop Monthly Sold Qty (P90D)": 91.5, "Adj. Safety Stock": 28.5,
         "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 300, "Site Code": "M01", "Shop Nature": "Mix", "OM": "OM1",
         "Sold Qty (P90D)": 81.0, "Shop Stock Qty (CM)": 20.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10.5,
         "D001 Net Stock Qty (CM)": 10.0},
    ]
    raw, _, _, _, cross, _, _ = transform_hybrid(_df(rows), safety_days=_SD10)
    # T01：所需量 ceil(270/90×10)=30、現有 28.5 → 調整量 +1.5；M01：所需量 ceil(81/90×10)=9、現有 10.5 → −1.5
    assert float(_get(raw, "T01", "Safety 建議調整量")) == 1.5
    assert float(_get(raw, "M01", "Safety 建議調整量")) == -1.5
    row_under = cross[cross["Safety Review"] == "Safety 設定不足"].iloc[0]
    row_over = cross[cross["Safety Review"] == "Safety 設定過高"].iloc[0]
    # 舊 int() 截斷會得 +1／−1；四捨五入（round-half-even）應得 +2／−2，與分類表／stats 同口徑
    assert int(row_under["Tourist 建議調整量合計"]) == 2 and int(row_under["合計 建議調整量"]) == 2
    assert int(row_over["Mix 建議調整量合計"]) == -2 and int(row_over["合計 建議調整量"]) == -2
    total = cross[cross["Safety Review"] == "總計"].iloc[0]
    assert int(total["合計 建議調整量"]) == 0


def test_nature_self_exclusion_guard_blocks_same_nature_source():
    # 顯式防呆驗證：即使來源池被餵入同 Nature 超額量，需求列也不得由同 Nature 調入
    sku = np.array(["9", "9"], dtype=object)
    site = np.array(["A", "B"], dtype=object)
    nature_arr = np.array(["Mix", "Mix"], dtype=object)
    need = np.array([True, False])
    demand = np.array([5.0, 0.0])
    a, g, d, desc = _hybrid_allocate_transfers(
        2, sku, site, need, demand, np.zeros(2), np.array([9.0, 9.0]),
        nature_arr=nature_arr)
    assert int(a.sum()) == 0 and int(g.sum()) == 5 and int(d.sum()) == 0 and desc[0] == ""
    # 回退防呆：不傳 nature_arr 時維持舊行為（池資料自行把關）
    a2, g2, d2, desc2 = _hybrid_allocate_transfers(
        2, sku, site, need, demand, np.zeros(2), np.array([9.0, 9.0]))
    assert int(a2.sum()) == 5 and int(d2.sum()) == 5 and "B:5" in desc2[0]


def test_dup_rows_diff_site_warns_self_transfer_guard_gap():
    # 同 SKU 同 Site Code 重複列 → warning 須說明自我調貨守衛以 Site Code 比對的破口
    rows = [
        {"SKU": 300, "Site Code": "T01", "Shop Nature": "Tourist", "OM": "OM1",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0, "Adj. Safety Stock": 30,
         "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 300, "Site Code": "T01", "Shop Nature": "Tourist", "OM": "OM1",
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 3,
         "D001 Net Stock Qty (CM)": 10.0},
    ]
    _, _, _, _, _, _, warns = transform_hybrid(_df(rows))
    dup_warn = [w for w in warns if "重複列" in w]
    assert dup_warn and "自我調貨守衛" in dup_warn[0]

