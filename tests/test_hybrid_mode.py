"""Validation tests for Safety Stock x NetStock Review mode (調貨事件 Review Safety 設定, v1.12.0/v1.13.0).

以調貨事件為入口 Review Safety Stock 設定；Review 基準為**按 Nature 的 Safety 基準天數**（預設 7/9/10，可調）：
- 調貨段與 transform_stock_checking 同規則（SKU 級加總 TurnDay、Nature 供應門檻、14 日目標、NT→Mix）；
  調貨段恆用 14 日，**不受 safety_days 影響**
- 需要調入列：Adj. Safety Stock < Safety 周轉所需量 → Safety 設定不足
- 已標示 SKU 的 Mix/NT 有貨列：Adj. Safety Stock > 生效基準（所需量經下調補底）→ Safety 設定過高
- 建議 Safety 設定值 = 所需量（不足列）／生效基準（過高列，下調補底 3 件：現有 ≥ 3 件者建議最低 3、
  現有 1–2 件不強制上調、周轉所需量 = 0 的歸零退場建議不設保底；見第 19 節）；
  Safety 建議調整量 = 建議值 − 現有設定（正=增加、負=減少）
- 調貨分兩變體：不限OM（契約欄位）與同OM（來源店與需求店同 OM，欄位加「(同OM)」後綴）；
  OM 限制只影響分配，Safety Review 與建議調整兩變體相同
- Shop Nature：Other/空白/大小寫差異一律正規化為 Non Tourist
- 輸出 8 工作表：RAW / Summary / 調貨 Review Safety Stock (不限OM) / 調貨 Review Safety Stock (同OM) / 店舖問題 SKU / Review 摘要 / 簡易調貨明細 (不限OM) / 簡易調貨明細 (同OM)
- 「Review 摘要」的三個統計區塊同母體（事件類型非空的列）：① Safety Review × Shop Nature 交叉表、
  ② Safety 調整方向分類（增加／減少／不變）、③ Safety 調整明細（SKU 級彙總）；②③ 於 v1.16.0 新增，
  由 RAW 即時彙算，不受 OM 限制影響，且不改動 transform_hybrid／generate_excel_hybrid 簽名
- 分類維度（v1.20.0）：nature（預設，需求方 Tourist-RF）／rp（需求方任何 Nature 的 RF 列，
  取消同 Nature 守衛與池級歸零）／both（計算同 nature，Summary／交叉表按 6 組顯示）；
  ND 列一律不作需求方（仍可作供應）；TurnDay 與 Safety 基準天數恆按 Shop Nature
"""
import warnings
warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
import pytest
from io import BytesIO
from openpyxl import load_workbook
from modes.framework import _transfer_input_consistency_warnings
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
    _HYBRID_SUGGEST_TURNDAY_COLUMN,
    _build_hybrid_shop_issues,
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
    _hybrid_dimension,
    _hybrid_dimension_groups,
    _hybrid_om_groups_from_values,
    _normalize_hybrid_om,
    _hybrid_skew_thresholds,
    _hybrid_skew_hint_text,
    _HYBRID_RP_GROUPS,
    _HYBRID_COMBINED_GROUPS,
    _hybrid_diag_summary_columns,
    _stock_checking_summary_columns,
    _STOCK_CHECKING_SUMMARY_COLUMNS,
    _normalize_hybrid_rp_type,
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
    assert all("Past 90 Days 0 Sales" not in w for w in warns)
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
    for c in ("強制調貨提示", "存貨偏重提示", "Non Tourist 加總 - Shop Stock Qty (CM)",
              "Tourist 涉及店舖數", "Mix 涉及店舖數", "Non Tourist 涉及店舖數"):
        assert list(hy_sc[c].fillna("")) == list(sc_sum[c].fillna("")), c
    assert int(hy_sc.iloc[0]["Tourist 涉及店舖數"]) == 1
    assert int(hy_sc.iloc[0]["Non Tourist 涉及店舖數"]) == 1


# ──────────────────────────────────────────────────────────────────────────────
# 3. Review 規則（Safety 基準天數；_SD10 為合法上限 10 日基準）
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


# 10 日基準專用場景：Tourist 需調入判不足；N01 判過高；M01 以 Adj 10 對齊 Mix 所需量作邊界正常列
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
    # 所需量 = 0（零銷售歸零屬退場）不設保底 3（v1.29.0 補底規則明確排除）
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
    # 顯示名一律為 Safety Stock (現時)（＝ 來源檔案的 Adj. Safety Stock），故警告亦用顯示名
    assert any("Safety Stock (現時)" in w for w in warns)


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
    assert (raw["OM"] == "空白").all()  # 缺欄時 RAW／主表補「空白」群組
    assert list(diagnosis["轉出量"]) == list(diagnosis_om["轉出量"])
    assert stats["nt_transfer_om"] == stats["nt_transfer"]


def test_blank_om_grouped_as_unfilled():
    rows = [dict(r) for r in _SCENARIO_OM]
    rows[2]["OM"] = None  # M11 OM 空白 → 空白組，與 T11 (Ivy) 不同組 → 同OM無來源
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
# 9. Excel 契約（9 工作表）
# ──────────────────────────────────────────────────────────────────────────────
def test_guide_sheet_first_with_params_echo():
    """「說明」導讀表：首位且為開啟預設頁；參數按 stats 還原列出；導讀與實用提示區塊齊備。"""
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(
        _df(_SCENARIO_A),
        bidirectional=True, nd_priority=True, sales_rebalance=True,
        safety_days={"Tourist": 8, "Mix": 9, "Non Tourist": 10},
        d001_threshold=120, turnday_threshold=13,
    )
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf)
    assert wb.sheetnames[0] == "說明"
    assert wb.active.title == "說明"
    ws = wb["說明"]
    col_a = [ws.cell(row=r, column=1).value for r in range(1, ws.max_row + 1)]
    col_b = [ws.cell(row=r, column=2).value for r in range(1, ws.max_row + 1)]
    # ① 參數列出（按本次 stats 還原）
    pairs = dict(zip(col_a, col_b))
    assert "① 本次執行參數（供核對）" in col_a
    assert pairs["分類維度"].startswith("Shop Nature")
    assert pairs["調貨門檻"].startswith("13 日")
    assert pairs["D001 庫存門檻"].startswith("120 件")
    assert pairs["Safety 基準周轉天數"] == "Tourist 8 日／Mix 9 日／Non Tourist 10 日"
    assert pairs["允許雙向調貨"].startswith("開啟")
    assert pairs["ND 優先轉出"].startswith("開啟")
    assert pairs["銷量再平衡"].startswith("開啟")
    assert "缺貨線 7 日" in pairs["銷量分類門檻"]
    # ②③④ 區塊與導讀列齊備
    for block in ("② 核心邏輯（這份報告怎樣算出來）", "③ 工作表導讀（先看哪一頁、應留意什麼）", "④ 實用提示"):
        assert block in col_a
    assert "說明（本頁）" in col_a and "店舖問題 SKU" in col_a and "簡易調貨明細 (不限OM)／(同OM)" in col_a


def test_generate_excel_sheets_and_headers():
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == [
        "說明", "RAW", "Summary", "調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)", "店舖問題 SKU", "Review 摘要",
        "Safety 設定建議 (全列)",
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
    assert len(sc_headers) == 29
    for col in ["Brand", "SKU", "強制調貨提示", "由Non Tourist 店舖預計調出數量", "存貨偏重提示", "涉及店舖數"]:
        assert col in sc_headers
    assert sc_headers.count("涉及店舖數") == 3


def test_generate_excel_summary_skew_highlight():
    # 偏重列：AC 提示欄與偏重 Nature 庫存佔比欄（Tourist=T）以醒目橘底 FFC000 加強顯示
    raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_BIDIR_B1), bidirectional=True)
    assert str(sc_summary.iloc[0]["存貨偏重提示"]).startswith("存貨偏重 Tourist")
    buf = generate_excel_hybrid(raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats)
    ws = load_workbook(buf)["Summary"]
    assert str(ws.cell(row=6, column=29).value).startswith("存貨偏重 Tourist")
    assert str(ws.cell(row=6, column=29).fill.start_color.rgb).endswith("FFC000")
    assert ws.cell(row=6, column=29).font.bold
    assert str(ws.cell(row=6, column=20).fill.start_color.rgb).endswith("FFC000")  # Tourist 庫存 %
    assert not str(ws.cell(row=6, column=22).fill.start_color.rgb).endswith("FFC000")  # Mix 未偏重
    assert not str(ws.cell(row=6, column=24).fill.start_color.rgb).endswith("FFC000")  # NT 未偏重


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


@pytest.mark.parametrize("column", ["Site Code", "SKU"])
@pytest.mark.parametrize("value", [None, np.nan, pd.NA, "", "   "])
def test_missing_transfer_identifier_rejected(column, value):
    df = _df(_SCENARIO_A).astype({column: object})
    df.loc[1, column] = value
    with pytest.raises(ValueError, match=column):
        transform_hybrid(df)


@pytest.mark.parametrize("value", [300.1, "300.9", np.inf, -np.inf])
def test_invalid_numeric_sku_rejected_before_truncation(value):
    df = _df(_SCENARIO_A).astype({"SKU": object})
    df.loc[1, "SKU"] = value
    with pytest.raises(ValueError, match="SKU"):
        transform_hybrid(df)


@pytest.mark.parametrize("duplicate_pos", [0, 1])
def test_duplicate_demand_or_supply_rejected(duplicate_pos):
    df = _df(_SCENARIO_A)
    df = pd.concat([df, df.iloc[[duplicate_pos]]], ignore_index=True)
    with pytest.raises(ValueError, match="Site Code × SKU"):
        transform_hybrid(df)


def test_duplicate_normalized_transfer_key_rejected():
    df = _df(_SCENARIO_A).astype({"SKU": object})
    duplicate = df.iloc[[1]].copy()
    duplicate["Site Code"] = " N01 "
    duplicate["SKU"] = "300.0"
    with pytest.raises(ValueError, match="Site Code × SKU"):
        transform_hybrid(pd.concat([df, duplicate], ignore_index=True))


def test_transfer_identifiers_trimmed_without_mutating_input():
    df = _df(_SCENARIO_A).astype({"SKU": object})
    df["Site Code"] = " " + df["Site Code"] + " "
    df["SKU"] = " 300 "
    before = df.copy(deep=True)
    raw, _, diagnosis, *_ = transform_hybrid(df)
    assert raw["SKU"].eq("300").all()
    assert raw["Site Code"].tolist() == ["T01", "N01", "M01"]
    assert _get(raw, "T01", "調貨來源 (Site Code)") == "N01:9"
    assert set(diagnosis["Site Code"]) == {"T01", "N01", "M01"}
    pd.testing.assert_frame_equal(df, before)


def test_same_om_blank_display_label_cannot_match_real_om():
    rows = [dict(r) for r in _SCENARIO_A[:2]]
    rows[0]["OM"] = "空白"
    rows[1]["OM"] = None
    raw, _, _, diagnosis_om, _, stats, _ = transform_hybrid(_df(rows))
    assert _get(raw, "T01", "實際調入量") == 9
    assert _get(raw, "T01", "實際調入量 (同OM)") == 0
    assert _get(raw, "T01", "調入缺口 (同OM)") == 9
    assert _get(raw, "N01", "轉出量 (同OM)") == 0
    assert diagnosis_om["實際調入量"].sum() == 0
    assert stats["total_actual_om"] == 0


@pytest.mark.parametrize("dimension", ["nature", "rp", "both", "om", "nature_om", "sales", "sales_nature"])
@pytest.mark.parametrize("bidirectional", [False, True])
@pytest.mark.parametrize("sales_rebalance", [False, True])
@pytest.mark.parametrize("nd_priority", [False, True])
def test_transfer_inventory_invariants(dimension, bidirectional, sales_rebalance, nd_priority):
    rng = np.random.default_rng(20261001)
    rows = []
    for sku in range(1, 5):
        for pos in range(12):
            sold = int(rng.integers(0, 361))
            rows.append({
                "SKU": sku, "Site Code": f"S{pos:02}",
                "Shop Nature": ["Tourist", "Mix", "Non Tourist"][pos % 3],
                "Sold Qty (P90D)": sold,
                "Shop Monthly Sold Qty (P90D)": sold / 90 * 30.5,
                "Shop Stock Qty (CM)": int(rng.integers(0, 81)),
                "RP Type": "ND" if pos % 4 == 0 else "RF",
                "OM": "Ivy" if pos % 2 else "Max",
                "D001 Net Stock Qty (CM)": 0,
            })
    raw, *_ = transform_hybrid(
        _df(rows), dimension=dimension, bidirectional=bidirectional,
        sales_rebalance=sales_rebalance, nd_priority=nd_priority,
    )
    by_key = raw.set_index(["SKU", "Site Code"])
    for actual, donated, gap, desc in [
        ("實際調入量", "轉出量", "調入缺口", "調貨來源 (Site Code)"),
        ("實際調入量 (同OM)", "轉出量 (同OM)", "調入缺口 (同OM)", "調貨來源 (同OM)"),
    ]:
        assert (raw[actual] + raw[gap]).equals(raw["調入需求量 (補到14日)"])
        totals = raw.groupby("SKU")[[actual, donated]].sum()
        assert (totals[actual] == totals[donated]).all()
        after = raw["Shop Stock Qty (CM)"] + raw[actual] - raw[donated]
        assert after.ge(0).all()
        assert not ((raw[actual] > 0) & (raw[donated] > 0)).any()
        donors = raw[donated] > 0
        reserve = np.ceil(raw["Sold Qty (P90D)"] / 90 * 14)
        assert (after[donors] >= reserve[donors]).all()
        assert raw.loc[raw["RP Type"] == "ND", actual].eq(0).all()
        detail = _build_hybrid_transfer_detail(raw, variant="om" if actual.endswith("(同OM)") else "free")
        assert detail["調貨數量"].sum() == raw[actual].sum()
        for _, target in raw.iterrows():
            described = 0
            for part in filter(None, target[desc].split("; ")):
                site, qty = part.rsplit(":", 1)
                source = by_key.loc[(target["SKU"], site)]
                assert site != target["Site Code"]
                if dimension != "rp":
                    assert source["Shop Nature"] != target["Shop Nature"]
                if actual.endswith("(同OM)"):
                    assert source["OM"] == target["OM"]
                described += int(qty)
            assert described == target[actual]


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
# Hybrid 入口拒絕重複店舖×SKU；分配層仍獨立保留自我調貨守衛。
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
    warns = _transfer_input_consistency_warnings(_df(_SCENARIO_DUP))
    assert any("Site Code × SKU" in w and "重複" in w for w in warns)
    assert any("Shop Nature" in w and "不一致" in w for w in warns)
    with pytest.raises(ValueError, match="Site Code × SKU"):
        transform_hybrid(_df(_SCENARIO_DUP))


def test_duplicate_same_nature_warns_dup_only():
    rows = [
        {"SKU": 510, "Site Code": "S01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0, "D001 Net Stock Qty (CM)": 10.0},
        {"SKU": 510, "Site Code": "S01", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 47.0,
         "Shop Monthly Sold Qty (P90D)": 120.0, "D001 Net Stock Qty (CM)": 10.0},
    ]
    warns = _transfer_input_consistency_warnings(_df(rows))
    assert any("Site Code × SKU" in w and "重複" in w for w in warns)
    assert not any("不一致" in w for w in warns)
    with pytest.raises(ValueError, match="Site Code × SKU"):
        transform_hybrid(_df(rows))


@pytest.mark.parametrize("same_om", [False, True])
def test_self_transfer_guarded_and_no_site_both_in_and_out(same_om):
    actual, gap, donated, desc = _hybrid_allocate_transfers(
        3, np.array(["500"] * 3), np.array(["S01", "S01", "S02"]),
        np.array([True, False, True]), np.array([9, 0, 9]),
        np.array([0, 16, 0]), np.zeros(3),
        om_arr=np.array(["OM1"] * 3) if same_om else None,
        nature_arr=np.array(["Tourist", "Non Tourist", "Tourist"]),
    )
    assert actual.tolist() == [0, 0, 9]
    assert gap.tolist() == [9, 0, 0]
    assert donated.tolist() == [0, 9, 0]
    assert desc.tolist() == ["", "", "S01:9"]
    assert actual.sum() == donated.sum()


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
    assert float(_get(raw, "T01", "Safety 建議調整量")) == -20  # 10 日基準：建議 10 − 現有 30
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
    with pytest.raises(ValueError, match="Site Code × SKU"):
        transform_hybrid(_df(rows), bidirectional=True)
    for same_om in (False, True):
        actual, gap, donated, desc = _hybrid_allocate_transfers(
            3, np.array(["803"] * 3), np.array(["S06", "S06", "S07"]),
            np.array([True, False, False]), np.array([9, 0, 0]),
            np.zeros(3), np.zeros(3), tourist_avail0=np.array([0, 16, 16]),
            nature_arr=np.array(["Mix", "Tourist", "Tourist"]),
            om_arr=np.array(["Ivy", "Max", "Max"]) if same_om else None,
        )
        assert actual[0] == (0 if same_om else 9)
        assert gap[0] == (9 if same_om else 0)
        assert donated.tolist() == ([0, 0, 0] if same_om else [0, 0, 9])
        assert desc[0] == ("" if same_om else "S07:9")


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
        "說明", "RAW", "Summary", "調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)", "店舖問題 SKU", "Review 摘要",
        "Safety 設定建議 (全列)",
        "簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)",
    ]
    raw_headers = [c.value for c in next(wb["RAW"].iter_rows(min_row=1, max_row=1))]
    assert "可轉出貨量 (Tourist)" in raw_headers
    sc_headers = [c.value for c in next(wb["Summary"].iter_rows(min_row=5, max_row=5))]
    assert len(sc_headers) == 30
    assert sc_headers[29] == "由Tourist 店舖預計調出數量"
    # 單向輸出 A–AC 契約欄序不動
    off_args = transform_hybrid(_df(_BIDIR_B1))
    buf_off = generate_excel_hybrid(*off_args[:5], off_args[5])
    wb_off = load_workbook(buf_off, read_only=True)
    off_headers = [c.value for c in next(wb_off["Summary"].iter_rows(min_row=5, max_row=5))]
    assert len(off_headers) == 29
    assert off_headers == sc_headers[:29]
    ws = wb["店舖問題 SKU"]
    assert [c.value for c in next(ws.iter_rows(min_row=3, max_row=3))][:11] == [
        "Site Code", "Shop Name", "OM", "Shop Nature",
        "Brand", "SKU", "SKU Desc", "Buyer",
        "問題類型", "問題說明", "建議方法",
    ]
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
    rows = [dict(r) for r in _SCENARIO_MULTI]
    rows[2]["RP Type"] = "ND"  # M21 轉 ND → 仍可作供應來源，RP Type 應帶出
    raw, *_ = transform_hybrid(_df(rows))
    assert _get(raw, "T21", "調貨來源 (Site Code)") == "N21:6; M21:11"
    det = _build_hybrid_transfer_detail(raw)
    assert list(det.columns) == _HYBRID_TRANSFER_DETAIL_COLUMNS
    assert _HYBRID_TRANSFER_DETAIL_COLUMNS.index("調貨數量") == _HYBRID_TRANSFER_DETAIL_COLUMNS.index("Target Site Code(接收店舖)") - 1
    assert _HYBRID_TRANSFER_DETAIL_COLUMNS.index("Target RP Type") == _HYBRID_TRANSFER_DETAIL_COLUMNS.index("Target Shop Name") + 1
    assert _HYBRID_TRANSFER_DETAIL_COLUMNS.index("Supply RP Type") == _HYBRID_TRANSFER_DETAIL_COLUMNS.index("Supply Shop Name") + 1
    assert len(det) == 2
    assert int(det["調貨數量"].sum()) == int(_get(raw, "T21", "實際調入量")) == 17
    r0 = det.iloc[0]  # 按 Supply Site Code 排序：M21 先於 N21
    assert r0["Supply Site Code(調貨來源)"] == "M21" and int(r0["調貨數量"]) == 11
    assert r0["Target Site Code(接收店舖)"] == "T21"
    assert r0["Target Shop Name"] == "ShopT" and r0["Target OM"] == "Ivy"
    assert r0["Target RP Type"] == "RF"
    assert r0["Supply Shop Name"] == "ShopM" and r0["Supply OM"] == "Ivy"
    assert r0["Supply RP Type"] == "ND"
    r1 = det.iloc[1]
    assert r1["Supply Site Code(調貨來源)"] == "N21" and int(r1["調貨數量"]) == 6
    assert r1["Supply Shop Name"] == "ShopN" and r1["Supply OM"] == "Max"
    assert r1["Supply RP Type"] == "RF"
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
    assert hb30.iloc[0]["Target RP Type"] == "" and hb30.iloc[0]["Supply RP Type"] == ""  # 缺 RP Type 欄留空
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
    assert ws.auto_filter.ref == "A1:L3"
    assert ws.cell(row=2, column=8).number_format == "#,##0"
    assert int(ws.cell(row=2, column=8).value) == 11  # M21:11（排序後首行）
    assert ws.cell(row=2, column=6).value == "RF"  # Supply RP Type（M21）
    assert ws.cell(row=2, column=11).value == "RF"  # Target RP Type
    ws_om = wb["簡易調貨明細 (同OM)"]
    assert ws_om.auto_filter.ref == "A1:L2"
    assert int(ws_om.cell(row=2, column=8).value) == 11
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
# 兩個 SKU 的對照場景：SKU 301 為 SKU 300 的等價複製（10 日基準：各 1 列不足 +10、1 列過高 −2、1 列正常）
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
        "說明", "RAW", "Summary", "調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)", "店舖問題 SKU", "Review 摘要",
        "Safety 設定建議 (全列)",
        "簡易調貨明細 (不限OM)", "簡易調貨明細 (同OM)",
    ]  # 10 表契約（首位為「說明」導讀表），新統計只在「Review 摘要」內延伸，另加「Safety 設定建議 (全列)」工作表
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
    assert ws.cell(row=20, column=1).value == "無符合資料（①–④ 統計區塊以調貨事件列為母體）"
    # ⑤ 銷量 × 存貨分類母體為全部列，無調貨事件時仍恆寫入
    assert str(ws.cell(row=22, column=1).value).startswith("銷量 × 存貨分類交叉表")
    assert ws.cell(row=23, column=1).value == "銷量存貨分類"
    assert [ws.cell(row=r, column=1).value for r in range(24, 29)] == \
        ["高銷缺貨", "低銷積壓", "零銷積壓", "正常", "總計"]
    assert int(ws.cell(row=27, column=5).value) == 1  # 正常列數
    assert int(ws.cell(row=28, column=5).value) == 1  # 總計列數


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
    # T01 Tourist：ceil(360/90 × 7) = 28，現有 30 → 不再判不足（14 日基準下會判不足）
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
    assert _HYBRID_SUGGEST_TURNDAY_COLUMN in _HYBRID_DIAG_COLUMNS
    assert _HYBRID_DIAG_COLUMNS.index(_HYBRID_SUGGEST_TURNDAY_COLUMN) == _HYBRID_DIAG_COLUMNS.index("建議 Safety 設定值") + 1
    assert "14日周轉所需量" not in _HYBRID_DIAG_COLUMNS
    raw, _, diagnosis, _, _, _, _ = transform_hybrid(_df(_SCENARIO_A))
    assert list(diagnosis.columns) == _HYBRID_DIAG_COLUMNS  # 主表 29 欄契約完整（含 RP Type、建議值 Turnday 與銷量帶／銷量存貨分類）
    display = _hybrid_raw_display_columns(raw)
    assert _HYBRID_SAFETY_LINE_COLUMN in display and _HYBRID_SAFETY_DAYS_COLUMN in display
    assert _HYBRID_SUGGEST_TURNDAY_COLUMN in display
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
    # 小數調整量：cross 表合計須四捨五入（-0.5→-1、+0.5→0），與分類表／stats 同一準則
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
    # 舊 int() 截斷會得 +1／−1；四捨五入（round-half-even）應得 +2／−2，與分類表／stats 同一準則
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


def test_conflicting_duplicate_stock_rejected_before_transfer():
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
    with pytest.raises(ValueError, match="Site Code × SKU"):
        transform_hybrid(_df(rows))


# ──────────────────────────────────────────────────────────────────────────────
# 18. 分類維度（v1.20.0）：Shop Nature（預設）／RP Type／Shop Nature＋RP Type
#     ND 列一律不作需求方（仍可作供應）；rp 維度需求＝任何 Nature 的 RF 列，
#     取消同 Nature 守衛與池級歸零；both 維度計算同 nature、僅分組顯示不同
# ──────────────────────────────────────────────────────────────────────────────
def _rp_rows(base_rows):
    return [dict(r) for r in base_rows]


def test_dimension_normalization_and_groups():
    assert _hybrid_dimension("RP") == "rp"
    assert _hybrid_dimension("Both") == "both"
    assert _hybrid_dimension(None) == "nature"
    assert _hybrid_dimension("bogus") == "nature"
    assert _hybrid_dimension_groups("nature") == ("Tourist", "Mix", "Non Tourist")
    assert _hybrid_dimension_groups("rp") == _HYBRID_RP_GROUPS
    assert _hybrid_dimension_groups("both") == _HYBRID_COMBINED_GROUPS
    assert _HYBRID_COMBINED_GROUPS == (
        "Tourist-RF", "Tourist-ND", "Mix-RF", "Mix-ND", "Non Tourist-RF", "Non Tourist-ND")


def test_normalize_hybrid_rp_type_rules():
    df = pd.DataFrame({"RP Type": ["rf", " ND ", "Other", None, "RF"]})
    out, mapped, present = _normalize_hybrid_rp_type(df)
    assert present and mapped == 2  # Other 與 None 歸 ND
    assert list(out["RP Type"]) == ["RF", "ND", "ND", "ND", "RF"]
    df2 = pd.DataFrame({"X": [1]})
    _, mapped2, present2 = _normalize_hybrid_rp_type(df2)
    assert not present2 and mapped2 == 0


def test_nd_tourist_row_never_demand_in_default_mode():
    rows = _rp_rows(_SCENARIO_A)
    rows[0]["RP Type"] = "ND"  # T01 Tourist ND → 不作需求方
    raw, _, _, _, _, stats, warns = transform_hybrid(_df(rows))
    assert stats["need_rows"] == 0
    assert _get(raw, "T01", "需要調貨") == ""
    assert int(_get(raw, "T01", "調入需求量 (補到14日)")) == 0
    assert any("沒有符合調貨條件" in w for w in warns)


def test_nd_row_can_still_supply_in_default_mode():
    rows = _rp_rows(_SCENARIO_A)
    rows[1]["RP Type"] = "ND"  # N01 Non Tourist ND → 仍可作供應來源
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(rows))
    assert stats["need_rows"] == 1
    assert _get(raw, "N01", "事件類型") == "實際轉出"
    assert int(_get(raw, "N01", "轉出量")) == 9  # T01 需求 ceil(360/90×14)−47 = 9
    assert "N01:9" in _get(raw, "T01", "調貨來源 (Site Code)")


def test_unknown_rp_normalized_to_nd_with_warning():
    rows = _rp_rows(_SCENARIO_A)
    rows[0]["RP Type"] = "Other"
    raw, _, _, _, _, stats, warns = transform_hybrid(_df(rows))
    assert any("RP Type" in w and "ND" in w for w in warns)
    assert stats["need_rows"] == 0  # Other→ND→不作需求
    assert _get(raw, "T01", "RP Type") == "ND"


def test_missing_rp_type_column_falls_back_to_nature():
    df = _df(_SCENARIO_A).drop(columns=["RP Type"])
    raw, _, diagnosis, _, _, stats, warns = transform_hybrid(df, dimension="rp")
    assert any("缺少「RP Type」" in w and "退回" in w for w in warns)
    assert stats["dimension"] == "nature"
    assert stats["need_rows"] == 1  # 全部列按 RF 處理，維持舊行為
    assert "RP Type" not in raw.columns  # RAW 不虛構欄位
    assert "RP Type" not in diagnosis.columns and len(diagnosis.columns) == 28  # 27＋建議值 Turnday（v1.40.0）


# Mix 池整體缺貨（加總 TurnDay < 14）但 M02 有超額——rp 維度允許同 Nature 店間再平衡
_SCENARIO_RP = [
    {"SKU": 400, "Site Code": "M01", "Shop Nature": "Mix", "RP Type": "RF",
     "Sold Qty (P90D)": 360.0, "Shop Stock Qty (CM)": 20.0,
     "Shop Monthly Sold Qty (P90D)": 160.0, "Adj. Safety Stock": 20,
     "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 400, "Site Code": "M02", "Shop Nature": "Mix", "RP Type": "ND",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 60.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30,
     "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 400, "Site Code": "N01", "Shop Nature": "Non Tourist", "RP Type": "RF",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 20.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 20,
     "D001 Net Stock Qty (CM)": 10.0},
]


def test_rp_mode_allows_same_nature_rebalance_and_all_rf_demand():
    # nature 維度：無 Tourist 列 → 無需求
    _, _, _, _, _, stats_n, _ = transform_hybrid(_df(_SCENARIO_RP))
    assert stats_n["need_rows"] == 0
    # rp 維度：M01（Mix-RF，池 TurnDay 12.8 < 14）轉為需求方，
    # 由 N01（NT-RF 超額 6）與同 Nature 的 M02（ND 超額 46）補足 36 件
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_RP), dimension="rp")
    assert stats["dimension"] == "rp"
    assert stats["need_rows"] == 1
    assert _get(raw, "M01", "需要調貨") == "是"
    assert int(_get(raw, "M01", "調入需求量 (補到14日)")) == 36
    assert int(_get(raw, "M01", "實際調入量")) == 36 and int(_get(raw, "M01", "調入缺口")) == 0
    src = _get(raw, "M01", "調貨來源 (Site Code)")
    assert "N01:6" in src and "M02:30" in src  # NT 優先、同 Nature（Mix-ND）補足
    # rp 維度不作池級歸零：Mix 池低於門檻時 M02 仍有可轉出量
    assert int(_get(raw, "M02", "可轉出貨量 (Mix)")) == 46


def test_rp_mode_nd_rows_never_demand():
    rows = _rp_rows(_SCENARIO_RP)
    rows[0]["RP Type"] = "ND"  # M01 轉 ND → rp 維度下不作需求
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(rows), dimension="rp")
    assert stats["need_rows"] == 0
    assert _get(raw, "M01", "需要調貨") == ""


def test_rp_dimension_switches_summary_and_cross_columns():
    raw, sc, _, _, cross, stats, _ = transform_hybrid(_df(_SCENARIO_RP), dimension="rp")
    expected = [c for c in _stock_checking_summary_columns("rp") if c in sc.columns]
    assert list(sc.columns) == expected
    for c in ("RF 加總 - Shop Stock Qty (CM)", "ND TurnDay", "RF 涉及店舖數", "ND 涉及店舖數",
              "RF 加總 - Shop Monthly Sold Qty (P90D) %",
              "由RF 店舖預計調出數量", "由ND 店舖預計調出數量"):
        assert c in sc.columns, c
    assert "Tourist 加總 - Shop Stock Qty (CM)" not in sc.columns
    row = sc.iloc[0]
    assert int(row["RF 涉及店舖數"]) == 2 and int(row["ND 涉及店舖數"]) == 1
    assert list(cross.columns) == _hybrid_diag_summary_columns("rp")
    assert cross["Safety Review"].tolist() == [
        "Safety 設定不足", "Safety 設定過高", "設定正常—缺貨非設定因素",
        "設定正常—過剩非設定因素", "總計"]
    total = cross[cross["Safety Review"] == "總計"].iloc[0]
    assert int(total["RF 列數"]) == 2 and int(total["ND 列數"]) == 1  # M01/N01 RF、M02 ND
    assert int(total["合計 列數"]) == 3
    # 由ND 轉出 30（M02）、由RF 轉出 6（N01）
    row = sc.iloc[0]
    assert int(row["由ND 店舖預計調出數量"]) == 30 and int(row["由RF 店舖預計調出數量"]) == 6


def test_both_dimension_same_numbers_with_six_group_columns():
    df = _df(_SCENARIO_A)
    raw_n, sc_n, _, _, cross_n, stats_n, _ = transform_hybrid(df)
    raw_b, sc_b, _, _, cross_b, stats_b, _ = transform_hybrid(df, dimension="both")
    assert stats_b["dimension"] == "both"
    # 計算與預設完全相同（需求方皆為 Tourist-RF）
    for c in ("需要調貨", "實際調入量", "轉出量", "調入缺口", "事件類型"):
        assert list(raw_n[c]) == list(raw_b[c]), c
    expected_b = [c for c in _stock_checking_summary_columns("both") if c in sc_b.columns]
    assert list(sc_b.columns) == expected_b
    assert "Tourist-RF 加總 - Shop Stock Qty (CM)" in sc_b.columns
    assert "Non Tourist-ND TurnDay" in sc_b.columns
    assert "Tourist-RF 涉及店舖數" in sc_b.columns
    assert int(sc_b.iloc[0]["Tourist-RF 涉及店舖數"]) == 1
    assert list(cross_b.columns) == _hybrid_diag_summary_columns("both")
    assert list(cross_b.columns)[1:4] == ["Tourist-RF 列數", "Tourist-RF SKU 數", "Tourist-RF 建議調整量合計"]
    t_n = cross_n.iloc[-1]
    t_b = cross_b.iloc[-1]
    assert int(t_b["合計 列數"]) == int(t_n["合計 列數"]) == 3
    assert int(t_b["Tourist-RF 列數"]) == 1 and int(t_b["Tourist-ND 列數"]) == 0


def test_nature_dimension_summary_contract_columns_unchanged():
    assert _stock_checking_summary_columns("nature") == _STOCK_CHECKING_SUMMARY_COLUMNS
    assert _stock_checking_summary_columns("nature", bidirectional=True) == (
        _STOCK_CHECKING_SUMMARY_COLUMNS + ["由Tourist 店舖預計調出數量"])


def test_excel_dimension_layouts():
    # rp：Summary 群組標籤「RP Type」、G=2 時強制調貨提示＝第 20 欄、偏重＝第 23 欄
    raw, sc, diag, diag_om, cross, stats, _ = transform_hybrid(_df(_SCENARIO_RP), dimension="rp")
    wb = load_workbook(generate_excel_hybrid(raw, sc, diag, diag_om, cross, stats))
    ws = wb["Summary"]
    assert ws.cell(row=3, column=8).value == "RP Type"
    assert ws.cell(row=5, column=20).value == "強制調貨提示"
    assert ws.cell(row=5, column=23).value == "存貨偏重提示"
    assert ws.cell(row=5, column=11).value == "涉及店舖數"
    assert ws.cell(row=4, column=12).value == "ND"
    assert str(wb["Review 摘要"].cell(row=19, column=1).value) == "Safety Review × RP Type 交叉表"
    # both：G=6 → 提示第 44 欄、偏重第 51 欄
    raw2, sc2, diag2, diag2_om, cross2, stats2, _ = transform_hybrid(_df(_SCENARIO_A), dimension="both")
    ws2 = load_workbook(generate_excel_hybrid(raw2, sc2, diag2, diag2_om, cross2, stats2))["Summary"]
    assert ws2.cell(row=3, column=8).value == "Shop Nature＋RP Type"
    assert ws2.cell(row=5, column=44).value == "強制調貨提示"
    assert ws2.cell(row=5, column=51).value == "存貨偏重提示"
    # nature 預設：A–AC 契約欄序（提示第 26 欄、偏重第 29 欄）
    raw3, sc3, diag3, diag3_om, cross3, stats3, _ = transform_hybrid(_df(_SCENARIO_A))
    ws3 = load_workbook(generate_excel_hybrid(raw3, sc3, diag3, diag3_om, cross3, stats3))["Summary"]
    assert ws3.cell(row=3, column=8).value == "Shop Nature"
    assert ws3.cell(row=5, column=26).value == "強制調貨提示"
    assert ws3.cell(row=5, column=29).value == "存貨偏重提示"


# ──────────────────────────────────────────────────────────────────────────────
# 19. 下調建議補底 3 件（SAP MOQ：Safety 1–2 件難以補貨；v1.29.0）
# 規則：過高列所需量 ≥ 1 且現有 ≥ 3 件 → 建議最低 3 件（現有 3 件補底後相等 → 設定正常）；
#       現有 1–2 件 → 維持原所需量（不強制上調到 3）；所需量 = 0（零銷售歸零）不設保底；
#       不足列（TurnDay 驅動上調）一律照原所需量；「Safety 周轉所需量」欄維持原值作稽核。
# 預設基準天數：Tourist 7 / Mix 9 / Non Tourist 10。
# ──────────────────────────────────────────────────────────────────────────────
_MIN3_ROWS = [
    {"SKU": 600, "Site Code": "T06", "Shop Nature": "Tourist",
     "Sold Qty (P90D)": 18.0, "Shop Stock Qty (CM)": 0.0,
     "Shop Monthly Sold Qty (P90D)": 6.1, "Adj. Safety Stock": 0},
    # NT 池加總 TurnDay = 90 ÷ 91.5 × 30.5 = 30.0 ≥ 14 → 有貨可調
    {"SKU": 600, "Site Code": "N06", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 9.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},   # 所需 1 → 補底 3
    {"SKU": 600, "Site Code": "N07", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 10.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 3},    # 所需 2 → 補底後 = 現值 → 正常
    {"SKU": 600, "Site Code": "N08", "Shop Nature": "Non Tourist",
     "Sold Qty (P90D)": 9.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 2},    # 現有 1–2 件 → 維持所需 1
]


def _min3_raw_and_stats():
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(_MIN3_ROWS))
    return raw, stats


def test_min3_floor_applies_to_over_rows():
    raw, _ = _min3_raw_and_stats()
    assert int(_get(raw, "N06", "Safety 周轉所需量")) == 1  # 稽核欄維持原所需量
    assert _get(raw, "N06", "Safety Review") == "Safety 設定過高"
    assert int(_get(raw, "N06", "建議 Safety 設定值")) == 3  # 10 → 補底 3（而非 1）
    assert float(_get(raw, "N06", "Safety 建議調整量")) == -7


def test_min3_floor_equal_current_becomes_normal():
    raw, stats = _min3_raw_and_stats()
    assert int(_get(raw, "N07", "Safety 周轉所需量")) == 2
    assert _get(raw, "N07", "Safety Review") == "設定正常—過剩非設定因素"  # 補底 3 = 現值，不出現 0 調整量假動作
    assert pd.isna(_get(raw, "N07", "建議 Safety 設定值"))
    assert pd.isna(_get(raw, "N07", "Safety 建議調整量"))
    assert stats["normal_rows"] == 1


def test_min3_floor_never_forces_low_current_up():
    raw, _ = _min3_raw_and_stats()
    assert _get(raw, "N08", "Safety Review") == "Safety 設定過高"
    assert int(_get(raw, "N08", "建議 Safety 設定值")) == 1  # 現有 2 件 → 維持原所需量，不強制調到 3
    assert float(_get(raw, "N08", "Safety 建議調整量")) == -1


def test_min3_floor_not_applied_to_increase_rows():
    raw, _ = _min3_raw_and_stats()
    assert int(_get(raw, "T06", "Safety 周轉所需量")) == 2  # ceil(18/90×7)
    assert _get(raw, "T06", "Safety Review") == "Safety 設定不足"
    assert int(_get(raw, "T06", "建議 Safety 設定值")) == 2  # TurnDay 驅動上調照常（目標 1–2 件不合保底）
    assert float(_get(raw, "T06", "Safety 建議調整量")) == 2


def test_min3_suggest_turnday_values():
    """建議 Safety 設定值 Turnday = 建議值 ÷ (P90D 月均 ÷ 30)；非建議列留空。"""
    raw, _ = _min3_raw_and_stats()
    # T06：建議 2 件、月均 6.1 → 2 ÷ (6.1/30) = 9.836 → 9.8
    assert float(_get(raw, "T06", _HYBRID_SUGGEST_TURNDAY_COLUMN)) == 9.8
    # N06：建議 3 件（保底）、月均 30.5 → 2.951 → 3.0
    assert float(_get(raw, "N06", _HYBRID_SUGGEST_TURNDAY_COLUMN)) == 3.0
    # N07：補底後改判設定正常、無建議值 → 留空
    assert pd.isna(_get(raw, "N07", _HYBRID_SUGGEST_TURNDAY_COLUMN))


def test_min3_suggest_turnday_in_diag_tables():
    """兩張 Review 主表（不限OM／同OM）都帶出建議值 Turnday，數值與 RAW 相同；
    Excel 主表以 1 位小數格式顯示（不得四捨五入成整數）。"""
    raw, sc, diag, diag_om, ds, stats, _ = transform_hybrid(_df(_MIN3_ROWS))
    assert list(diag.columns) == _HYBRID_DIAG_COLUMNS
    n06 = diag[diag["Site Code"] == "N06"].iloc[0]
    assert float(n06[_HYBRID_SUGGEST_TURNDAY_COLUMN]) == 3.0
    n06_om = diag_om[diag_om["Site Code"] == "N06"].iloc[0]
    assert float(n06_om[_HYBRID_SUGGEST_TURNDAY_COLUMN]) == 3.0
    td_col = _HYBRID_DIAG_COLUMNS.index(_HYBRID_SUGGEST_TURNDAY_COLUMN) + 1
    site_col = _HYBRID_DIAG_COLUMNS.index("Site Code") + 1
    wb = load_workbook(generate_excel_hybrid(raw, sc, diag, diag_om, ds, stats))
    for sheet in ("調貨 Review Safety Stock (不限OM)", "調貨 Review Safety Stock (同OM)"):
        ws = wb[sheet]
        for site, expected in (("N06", 3.0), ("T06", 9.8)):
            excel_row = next(
                r for r in range(2, ws.max_row + 1)
                if ws.cell(row=r, column=site_col).value == site
            )
            cell = ws.cell(row=excel_row, column=td_col)
            assert float(cell.value) == expected, (sheet, site, cell.value)
            assert cell.number_format == "0.0", (sheet, site, cell.number_format)


def test_min3_internal_floor_flag_hidden_from_outputs():
    """`_hybrid_safety_floor3` 屬內部旗標：只在計算層與店舖問題使用，不得寫入 Raw／主表顯示。"""
    raw, sc, diag, diag_om, ds, stats, _ = transform_hybrid(_df(_MIN3_ROWS))
    assert "_hybrid_safety_floor3" in raw.columns
    assert bool(_get(raw, "N06", "_hybrid_safety_floor3")) is True
    assert bool(_get(raw, "N08", "_hybrid_safety_floor3")) is False
    assert "_hybrid_safety_floor3" not in _hybrid_raw_display_columns(raw)
    wb = load_workbook(generate_excel_hybrid(raw, sc, diag, diag_om, ds, stats), read_only=True)
    raw_headers = [c.value for c in next(wb["RAW"].iter_rows(min_row=1, max_row=1))]
    assert "_hybrid_safety_floor3" not in raw_headers


def test_min3_shop_issue_advice_turnday_and_floor_note():
    """建議方法：調整量以「增加／減少 X 件」標明方向；所有建議列附建議周轉；受保底影響列（建議 3 > 所需量 ≥ 1）加註保底原因。"""
    raw, _ = _min3_raw_and_stats()
    issues = _build_hybrid_shop_issues(raw)
    n06 = issues[issues["Site Code"] == "N06"].iloc[0]
    assert "減少 7 件" in str(n06["建議方法"])
    assert "建議周轉約 3 日" in str(n06["建議方法"])
    assert "為保障自動補貨，最低建議設定為 3 件（SAP MOQ：Safety 1–2 件難以補貨）" in str(n06["建議方法"])
    # N08 現有 1–2 件不強制上調（建議 1 件＝原所需量）→ 有周轉、無保底句
    n08 = issues[issues["Site Code"] == "N08"].iloc[0]
    assert "減少 1 件" in str(n08["建議方法"])
    assert "建議周轉約 1 日" in str(n08["建議方法"])
    assert "為保障自動補貨" not in str(n08["建議方法"])
    # T06 不足列照原所需量上調 → 有周轉、無保底句
    t06 = issues[issues["Site Code"] == "T06"].iloc[0]
    assert "增加 2 件" in str(t06["建議方法"])
    assert "建議周轉約 9.8 日" in str(t06["建議方法"])
    assert "為保障自動補貨" not in str(t06["建議方法"])


def test_min3_stats_population_matches_masks():
    # 過高列＝N06＋N08（N07 補底後轉正常）、不足列＝T06；合計與調整量一致
    raw, stats = _min3_raw_and_stats()
    assert stats["under_rows"] == 1 and stats["over_rows"] == 2
    assert stats["under_inc_total"] == 2 and stats["over_dec_total"] == -8


def test_min3_stats_and_safety_tables_consistent():
    raw, stats = _min3_raw_and_stats()
    frame = _build_hybrid_safety_change_summary(raw)
    cat = _cat_rows(frame)
    assert int(cat["增加"]["列數"]) == 1 and int(cat["增加"]["Safety 建議調整量合計"]) == stats["under_inc_total"] == 2
    assert int(cat["減少"]["列數"]) == 2 and int(cat["減少"]["Safety 建議調整量合計"]) == stats["over_dec_total"] == -8
    assert int(cat["不變"]["列數"]) == 1 and int(cat["不變"]["Safety 建議調整量合計"]) == 0
    detail = _build_hybrid_safety_sku_detail(raw)
    assert len(detail) == 1
    r = detail.iloc[0]
    assert int(r["Review 列數"]) == 4
    assert int(r["需增加店舖數"]) == 1 and int(r["需減少店舖數"]) == 2 and int(r["設定正常店舖數"]) == 1
    # 調整列（T06/N06/N08）：現有 0+10+2、建議 2+3+1
    assert int(r["現有 Safety 設定合計（調整列）"]) == 12
    assert int(r["建議 Safety 設定合計（調整列）"]) == 6
    assert int(r["淨調整量"]) == -6


# ──────────────────────────────────────────────────────────────────────────────
# 20. ND 優先轉出（nd_priority=True，v1.30.0）
# 規則：**同一 Shop Nature 池內** RP Type＝ND 的來源列先於 RF 列取用；池間順序仍為
#       Non Tourist → Mix → Tourist（雙向時才含 Tourist），展開為
#       NT-ND → NT-RF → Mix-ND → Mix-RF → Tourist-ND → Tourist-RF；ND／RF 兩段各自維持
#       來源檔列序。轉出資格（池級歸零、逐列 14 日超額、自我調貨／同 Nature／同OM 守衛）
#       一律不放寬；ND 列仍不作需求方；「ND」的定義＝正規化後非 RF（含無法辨識列）。
#       因 donor_mask 與分配順序無關，Safety Review 與建議調整量逐列不變。預設關閉＝零變動。
# ──────────────────────────────────────────────────────────────────────────────
# T21 需調入 14 件；NT 池列序 N22(RF, 超額 16) → N21(ND, 超額 6)；M21(Mix-ND, 超額 11)
_NDPRIO_ROWS = [
    {"SKU": 700, "Site Code": "T21", "Shop Nature": "Tourist", "RP Type": "RF",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 0.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 20},
    {"SKU": 700, "Site Code": "N22", "Shop Nature": "Non Tourist", "RP Type": "RF",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
    {"SKU": 700, "Site Code": "N21", "Shop Nature": "Non Tourist", "RP Type": "ND",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 20.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
    {"SKU": 700, "Site Code": "M21", "Shop Nature": "Mix", "RP Type": "ND",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
]

# 決策 1 反例：ND 只在 Mix、RF 在 Non Tourist → 仍先取 Non Tourist-RF（不跨池搶 ND）
_NDPRIO_POOL_ROWS = [
    {"SKU": 710, "Site Code": "T31", "Shop Nature": "Tourist", "RP Type": "RF",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 0.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 20,
     "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 710, "Site Code": "N31", "Shop Nature": "Non Tourist", "RP Type": "RF",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
    {"SKU": 710, "Site Code": "M31", "Shop Nature": "Mix", "RP Type": "ND",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 25.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
]

# 「以 RP Type 分類」維度：需求 Tourist-RF，來源 Mix-RF 先列、Mix-ND 後列
_NDPRIO_RP_ROWS = [
    {"SKU": 720, "Site Code": "T51", "Shop Nature": "Tourist", "RP Type": "RF",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 0.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 20,
     "D001 Net Stock Qty (CM)": 10.0},
    {"SKU": 720, "Site Code": "M52", "Shop Nature": "Mix", "RP Type": "RF",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
    {"SKU": 720, "Site Code": "M51", "Shop Nature": "Mix", "RP Type": "ND",
     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 18.0,
     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 10},
]

# T21 的 D001 Net 需 < 門檻 100（BASE 99 已符合）
_SAFETY_COLS = ["需要調貨", "調入需求量 (補到14日)", "可轉出貨量 (Non Tourist)",
                "可轉出貨量 (Mix)", "Safety 周轉所需量", "Safety Review",
                "建議 Safety 設定值", "Safety 建議調整量"]


def _ndprio(nd_priority=False, **kw):
    return transform_hybrid(_df(_NDPRIO_ROWS), nd_priority=nd_priority, **kw)


def test_nd_priority_off_equals_current_contract():
    # 預設關閉時，不傳參數與顯式傳 False 結果相同，且按來源檔列序取用（RF 列 N22 先）
    raw_default, _, _, _, _, stats_default, _ = transform_hybrid(_df(_NDPRIO_ROWS))
    raw_off, _, _, _, _, stats_off, _ = _ndprio()
    for col in ["轉出量", "調貨來源 (Site Code)", "實際調入量", "調入缺口", "轉出量 (同OM)", "調貨來源 (同OM)"]:
        assert raw_default[col].tolist() == raw_off[col].tolist()
    assert _get(raw_off, "T21", "調貨來源 (Site Code)") == "N22:14"
    assert int(_get(raw_off, "N22", "轉出量")) == 14 and int(_get(raw_off, "N21", "轉出量")) == 0
    assert stats_off["nd_priority"] is False and stats_off["nd_transfer"] == 0
    assert stats_default["nd_priority"] is False


def test_nd_priority_prefers_nd_within_same_pool():
    raw, _, _, _, _, stats, _ = _ndprio(nd_priority=True)
    assert _get(raw, "T21", "調貨來源 (Site Code)") == "N21:6; N22:8"
    assert int(_get(raw, "N21", "轉出量")) == 6 and int(_get(raw, "N22", "轉出量")) == 8
    assert _get(raw, "T21", "調貨來源 (同OM)") == "N21:6; N22:8"  # 兩變體同規則（本場景全同 OM）
    assert stats["nd_priority"] is True and stats["nd_transfer"] == 6 and stats["nd_transfer_om"] == 6
    # Mix-ND 仍排在 Non Tourist-RF 之後 → 未被取用
    assert int(_get(raw, "M21", "轉出量")) == 0


def test_nd_priority_does_not_cross_pool_order():
    _, _, _, _, _, stats_off, _ = transform_hybrid(_df(_NDPRIO_POOL_ROWS))
    raw, _, _, _, _, stats_on, _ = transform_hybrid(_df(_NDPRIO_POOL_ROWS), nd_priority=True)
    # Non Tourist 只有 RF、Mix 才有 ND → 仍由 Non Tourist-RF 轉出，ND 不跨池前置
    assert _get(raw, "T31", "調貨來源 (Site Code)") == "N31:14"
    assert int(_get(raw, "M31", "轉出量")) == 0
    assert stats_on["nd_transfer"] == stats_off["nd_transfer"] == 0


def test_nd_priority_keeps_safety_review_unchanged():
    raw_off, _, _, _, _, stats_off, _ = _ndprio()
    raw_on, _, _, _, _, stats_on, _ = _ndprio(nd_priority=True)
    for col in _SAFETY_COLS:
        # equals 將 NaN 視為相等（建議值／調整量對非事件列留空）
        assert raw_off[col].reset_index(drop=True).equals(raw_on[col].reset_index(drop=True)), col
    # 池級合計與調入／缺口不變（池間順序未改）
    for key in ("nt_transfer", "mix_transfer", "total_actual", "total_gap",
                "nt_transfer_om", "mix_transfer_om", "need_rows", "donor_rows"):
        assert stats_off[key] == stats_on[key], key
    assert stats_off["nt_transfer"] == stats_on["nt_transfer"] == 14
    for key in ("under_rows", "over_rows", "normal_rows", "under_inc_total", "over_dec_total"):
        assert stats_off[key] == stats_on[key], key


def test_nd_priority_changes_event_label_and_stock_after():
    raw_off, _, _, _, _, _, _ = _ndprio()
    raw_on, _, _, _, _, _, _ = _ndprio(nd_priority=True)
    assert _get(raw_off, "N21", "事件類型") == "可轉出未使用"
    assert _get(raw_on, "N21", "事件類型") == "實際轉出"
    assert int(_get(raw_off, "N21", "調貨後庫存")) == 20
    assert int(_get(raw_on, "N21", "調貨後庫存")) == 14
    assert _get(raw_on, "M21", "事件類型") == _get(raw_off, "M21", "事件類型") == "可轉出未使用"
    assert _get(raw_on, "T21", "事件類型") == "需要調入"


def test_nd_priority_missing_rp_type_column_is_noop_with_warning():
    df = _df(_NDPRIO_ROWS).drop(columns=["RP Type"])
    raw_off, _, _, _, _, stats_off, _ = transform_hybrid(df)
    raw_on, _, _, _, _, stats_on, warns = transform_hybrid(df, nd_priority=True)
    assert any("ND 優先轉出" in w and "RP Type" in w for w in warns), warns
    assert raw_on["調貨來源 (Site Code)"].tolist() == raw_off["調貨來源 (Site Code)"].tolist()
    assert stats_on["nd_transfer"] == 0 and stats_off["nd_transfer"] == 0


def test_nd_priority_includes_unrecognized_rp_type_as_nd():
    rows = [dict(r) for r in _NDPRIO_ROWS]
    rows[2] = {**rows[2], "RP Type": ""}  # N21 空白 → 正規化為 ND，仍列入優先轉出
    raw, _, _, _, _, stats, warns = transform_hybrid(_df(rows), nd_priority=True)
    assert stats["mapped_rp_rows"] == 1
    assert _get(raw, "T21", "調貨來源 (Site Code)") == "N21:6; N22:8"
    assert any("RP Type" in w and "ND" in w for w in warns)


def test_nd_priority_applies_in_rp_dimension():
    _, _, _, _, _, stats_off, _ = transform_hybrid(_df(_NDPRIO_RP_ROWS), dimension="rp")
    raw_on, _, _, _, _, stats_on, _ = transform_hybrid(
        _df(_NDPRIO_RP_ROWS), dimension="rp", nd_priority=True)
    assert stats_off["dimension"] == stats_on["dimension"] == "rp"
    assert stats_off["nd_transfer"] == 0
    assert stats_on["nd_transfer"] == 4
    # ND 列仍不作需求方（M51 為 Mix-ND，缺貨也不會標示需要調貨）
    assert _get(raw_on, "M51", "需要調貨") != "是"


def test_nd_priority_dashboard_card_row():
    raw_off, _, _, _, _, stats_off, _ = _ndprio()
    raw_on, _, _, _, _, stats_on, _ = _ndprio(nd_priority=True)
    assert len(_hybrid_dashboard_metrics(raw_off, stats_off)) == 4
    rows = _hybrid_dashboard_metrics(raw_on, stats_on)
    assert len(rows) == 5
    assert [c[0] for c in rows[4]] == ["ND 來源轉出量", "ND 來源轉出量 (同OM)"]
    assert rows[4][0][1] == 6 and rows[4][1][1] == 6
    # stats 缺失時無法由 raw 反推開關狀態 → 不顯示 ND 卡，off 路徑兩準則一致
    assert len(_hybrid_dashboard_metrics(raw_on, None)) == 4
    assert _hybrid_dashboard_metrics(raw_off, stats_off) == _hybrid_dashboard_metrics(raw_off, None)
    # 回退準則仍算出 ND 轉出量（供其他彙總使用）
    assert _hybrid_stats_from_raw(raw_on)["nd_transfer"] == 6
    assert _hybrid_stats_from_raw(raw_off)["nd_transfer"] == 0


def test_nd_priority_bidir_and_nd_share_fifth_row():
    raw, _, _, _, _, stats, _ = transform_hybrid(
        _df(_BIDIR_B1), bidirectional=True, nd_priority=True)
    rows = _hybrid_dashboard_metrics(raw, stats)
    assert len(rows) == 5
    assert [c[0] for c in rows[4]][:2] == ["Tourist 轉出量", "Tourist 轉出量 (同OM)"]


def test_nd_priority_excel_metric_layout():
    raw_off, sc, dg, dg_om, cs, st_off, _ = _ndprio()
    ws_off = load_workbook(generate_excel_hybrid(raw_off, sc, dg, dg_om, cs, st_off))["Review 摘要"]
    assert ws_off.cell(row=19, column=1).value == "Safety Review × Shop Nature 交叉表"
    raw_on, sc2, dg2, dg2o, cs2, st_on, _ = _ndprio(nd_priority=True)
    ws_on = load_workbook(generate_excel_hybrid(raw_on, sc2, dg2, dg2o, cs2, st_on))["Review 摘要"]
    assert ws_on.cell(row=19, column=1).value == "ND 來源轉出量"
    assert ws_on.cell(row=19, column=2).value == "ND 來源轉出量 (同OM)"
    assert int(ws_on.cell(row=20, column=1).value) == 6
    assert ws_on.cell(row=23, column=1).value == "Safety Review × Shop Nature 交叉表"


def test_allocate_nd_priority_within_pool_only():
    sku = np.array(["900"] * 7, dtype=object)
    site = np.array(["T01", "N01", "N02", "M01", "M02", "T02", "T03"], dtype=object)
    nature = np.array(
        ["Tourist", "Non Tourist", "Non Tourist", "Mix", "Mix", "Tourist", "Tourist"], dtype=object)
    need = np.array([True] + [False] * 6)
    demand = np.array([20, 0, 0, 0, 0, 0, 0], dtype=float)
    nt = np.array([0, 10, 5, 0, 0, 0, 0], dtype=float)
    mix = np.array([0, 0, 0, 8, 3, 0, 0], dtype=float)
    tou = np.array([0, 0, 0, 0, 0, 2, 4], dtype=float)
    nd = np.array([False, False, True, False, True, True, False])
    _, gap_off, _, desc_off = _hybrid_allocate_transfers(
        7, sku, site, need, demand, nt, mix, tourist_avail0=tou, nature_arr=nature)
    assert desc_off[0] == "N01:10; N02:5; M01:5" and gap_off[0] == 0
    _, gap_on, donate_on, desc_on = _hybrid_allocate_transfers(
        7, sku, site, need, demand, nt, mix, tourist_avail0=tou, nature_arr=nature, nd_priority=nd)
    assert desc_on[0] == "N02:5; N01:10; M02:3; M01:2" and gap_on[0] == 0
    assert donate_on.tolist() == [0, 10, 5, 2, 3, 0, 0]
    # Tourist 來源列與需求列同 Nature → 兩個模式皆一律排除
    assert donate_on[5] == donate_on[6] == 0


def test_allocate_nd_priority_applies_to_tourist_pool():
    sku = np.array(["910"] * 3, dtype=object)
    site = np.array(["M01", "T02", "T01"], dtype=object)
    nature = np.array(["Mix", "Tourist", "Tourist"], dtype=object)
    need = np.array([True, False, False])
    demand = np.array([10, 0, 0], dtype=float)
    zero = np.zeros(3, dtype=float)
    tou = np.array([0, 7, 3], dtype=float)
    nd = np.array([False, False, True])
    _, _, _, desc_off = _hybrid_allocate_transfers(
        3, sku, site, need, demand, zero, zero, tourist_avail0=tou, nature_arr=nature)
    assert desc_off[0] == "T02:7; T01:3"
    _, _, _, desc_on = _hybrid_allocate_transfers(
        3, sku, site, need, demand, zero, zero, tourist_avail0=tou, nature_arr=nature, nd_priority=nd)
    assert desc_on[0] == "T01:3; T02:7"


# ──────────────────────────────────────────────────────────────────────────────
# 21. OM 分類維度：以 OM 分類（om）／以 Shop Nature＋OM 分類（nature_om）
#     計算與 nature 完全相同（需求 Tourist-RF，池級歸零與同 Nature 守衛照舊），
#     僅 Summary／交叉表分組不同；OM 名單為動態（排序後不重複值，空白自成「空白」組，大小寫不敏感）；
#     缺 OM 欄或全欄空白時退回 nature；存貨偏重門檻按維度下調
# ──────────────────────────────────────────────────────────────────────────────
def test_om_dimension_normalization_and_dynamic_groups():
    assert _hybrid_dimension("OM") == "om"
    assert _hybrid_dimension("Nature_OM") == "nature_om"
    assert _hybrid_dimension("bogus") == "nature"
    assert _hybrid_dimension_groups("om", om_groups=("B", "A")) == ("A", "B")
    assert _hybrid_dimension_groups("nature_om", om_groups=("B", "A")) == (
        "Tourist-A", "Tourist-B", "Mix-A", "Mix-B", "Non Tourist-A", "Non Tourist-B")
    # om_groups 缺省時退回單一群組（呼叫端應傳入實際名單）
    assert _hybrid_dimension_groups("om") == ("空白",)
    assert _hybrid_dimension_groups("nature_om") == (
        "Tourist-空白", "Mix-空白", "Non Tourist-空白")


def test_om_groups_from_values_rules():
    assert _hybrid_om_groups_from_values(["B", " A ", "B", None, ""]) == ("A", "B", "空白")
    assert _hybrid_om_groups_from_values(pd.Series(["Ivy", "Max", "Ivy"])) == ("Ivy", "Max")
    assert _hybrid_om_groups_from_values([]) == ("空白",)


def test_om_skew_thresholds_and_hint_text():
    assert _hybrid_skew_thresholds("nature") == (0.40, 0.20)
    assert _hybrid_skew_thresholds("rp") == (0.40, 0.20)
    assert _hybrid_skew_thresholds("both") == (0.40, 0.20)
    assert _hybrid_skew_thresholds("om") == (0.25, 0.12)
    assert _hybrid_skew_thresholds("nature_om") == (0.12, 0.08)
    assert "25%" in _hybrid_skew_hint_text("om") and "12" in _hybrid_skew_hint_text("om")
    assert "12%" in _hybrid_skew_hint_text("nature_om") and "8" in _hybrid_skew_hint_text("nature_om")


def test_om_dimension_same_numbers_as_nature():
    df = _df(_SCENARIO_OM)
    raw_n, sc_n, _, _, cross_n, stats_n, _ = transform_hybrid(df)
    raw_o, sc_o, _, _, cross_o, stats_o, _ = transform_hybrid(df, dimension="om")
    assert stats_o["dimension"] == "om"
    assert stats_o["om_groups"] == ["Ivy", "Max"]
    # 計算與預設完全相同（需求方皆為 Tourist-RF；NaN 視為相等）
    for c in ("需要調貨", "實際調入量", "轉出量", "調入缺口", "事件類型",
              "Safety Review", "建議 Safety 設定值", "Safety 建議調整量"):
        for x, y in zip(list(raw_n[c]), list(raw_o[c])):
            if pd.isna(x) and pd.isna(y):
                continue
            assert x == y, c
    assert stats_o["need_rows"] == stats_n["need_rows"] == 1
    # Summary 按 OM 分組：Ivy 兩店、Max 一店
    expected = [c for c in _stock_checking_summary_columns("om", om_groups=("Ivy", "Max")) if c in sc_o.columns]
    assert list(sc_o.columns) == expected
    assert "Tourist 加總 - Shop Stock Qty (CM)" not in sc_o.columns
    row = sc_o.iloc[0]
    assert int(row["Ivy 涉及店舖數"]) == 2 and int(row["Max 涉及店舖數"]) == 1
    assert int(row["由Max 店舖預計調出數量"]) == int(raw_o.loc[raw_o["OM"] == "Max", "轉出量"].sum())
    assert int(row["由Ivy 店舖預計調出數量"]) == int(raw_o.loc[raw_o["OM"] == "Ivy", "轉出量"].sum())
    # 交叉表按 OM 分組，合計與 nature 一致
    assert list(cross_o.columns) == _hybrid_diag_summary_columns("om", om_groups=("Ivy", "Max"))
    t_o = cross_o[cross_o["Safety Review"] == "總計"].iloc[0]
    t_n = cross_n[cross_n["Safety Review"] == "總計"].iloc[0]
    assert int(t_o["合計 列數"]) == int(t_n["合計 列數"]) == 3
    assert int(t_o["Ivy 列數"]) == 2 and int(t_o["Max 列數"]) == 1


def test_nature_om_dimension_cartesian_columns():
    df = _df(_SCENARIO_OM)
    raw_n, _, _, _, cross_n, stats_n, _ = transform_hybrid(df)
    raw_c, sc_c, _, _, cross_c, stats_c, _ = transform_hybrid(df, dimension="nature_om")
    assert stats_c["dimension"] == "nature_om"
    assert stats_c["om_groups"] == ["Ivy", "Max"]
    for c in ("需要調貨", "實際調入量", "轉出量", "調入缺口", "事件類型"):
        assert list(raw_n[c]) == list(raw_c[c]), c
    assert stats_c["need_rows"] == stats_n["need_rows"] == 1
    groups = ("Tourist-Ivy", "Tourist-Max", "Mix-Ivy", "Mix-Max",
              "Non Tourist-Ivy", "Non Tourist-Max")
    expected = [c for c in _stock_checking_summary_columns("nature_om", om_groups=("Ivy", "Max")) if c in sc_c.columns]
    assert list(sc_c.columns) == expected
    assert "Tourist-Ivy 加總 - Shop Stock Qty (CM)" in sc_c.columns
    assert "Non Tourist-Max TurnDay" in sc_c.columns
    assert "Tourist-Ivy 涉及店舖數" in sc_c.columns
    row = sc_c.iloc[0]
    assert int(row["Tourist-Ivy 涉及店舖數"]) == 1
    assert int(row["Mix-Ivy 涉及店舖數"]) == 1
    assert int(row["Non Tourist-Max 涉及店舖數"]) == 1
    assert int(row["Tourist-Max 涉及店舖數"]) == 0
    assert list(cross_c.columns) == _hybrid_diag_summary_columns("nature_om", om_groups=("Ivy", "Max"))
    assert list(cross_c.columns)[1:4] == ["Tourist-Ivy 列數", "Tourist-Ivy SKU 數", "Tourist-Ivy 建議調整量合計"]
    t_c = cross_c[cross_c["Safety Review"] == "總計"].iloc[0]
    t_n = cross_n[cross_n["Safety Review"] == "總計"].iloc[0]
    assert int(t_c["合計 列數"]) == int(t_n["合計 列數"]) == 3
    assert int(t_c["Tourist-Ivy 列數"]) == 1 and int(t_c["Mix-Ivy 列數"]) == 1
    assert int(t_c["Non Tourist-Max 列數"]) == 1


def test_missing_om_column_falls_back_to_nature():
    df = _df(_SCENARIO_OM).drop(columns=["OM"])
    _, _, _, _, _, stats_o, warns_o = transform_hybrid(df, dimension="om")
    assert any("OM" in w and "退回" in w for w in warns_o)
    assert stats_o["dimension"] == "nature"
    _, _, _, _, _, stats_c, warns_c = transform_hybrid(df, dimension="nature_om")
    assert any("OM" in w and "退回" in w for w in warns_c)
    assert stats_c["dimension"] == "nature"


def test_all_blank_om_falls_back_to_nature():
    df = _df(_SCENARIO_OM)
    df["OM"] = ""
    _, _, _, _, _, stats, warns = transform_hybrid(df, dimension="om")
    assert any("OM" in w and "退回" in w for w in warns)
    assert stats["dimension"] == "nature"


def test_om_dimension_survives_missing_rp_type():
    # OM 分組不依賴 RP Type：缺 RP 欄時 om 維度不退回（全部列按 RF 處理）
    df = _df(_SCENARIO_OM).drop(columns=["RP Type"])
    raw, _, _, _, _, stats, warns = transform_hybrid(df, dimension="om")
    assert stats["dimension"] == "om"
    assert stats["need_rows"] == 1
    assert "RP Type" not in raw.columns


def test_om_skew_hint_uses_lower_threshold():
    # O1 庫存佔比 50%（≥25%）且較銷量佔比高約 47 個百分點（≥12pp）→ om 維度提示；
    # 原 40%／20pp 門檻下 nature 維度（全 Tourist，庫存與銷量佔比皆 100%）不提示
    rows = [
        {"SKU": 700, "Site Code": "S01", "OM": "O1", "Shop Nature": "Tourist",
         "Sold Qty (P90D)": 18.0, "Shop Stock Qty (CM)": 60.0,
         "Shop Monthly Sold Qty (P90D)": 6.1, "Adj. Safety Stock": 5,
         "D001 Net Stock Qty (CM)": 500.0},
    ]
    for i, om in enumerate(["O2", "O3", "O4", "O5", "O6", "O7"], start=2):
        rows.append({"SKU": 700, "Site Code": f"S0{i}", "OM": om, "Shop Nature": "Tourist",
                     "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 10.0,
                     "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 5,
                     "D001 Net Stock Qty (CM)": 500.0})
    _, sc_o, _, _, _, stats_o, _ = transform_hybrid(_df(rows), dimension="om")
    assert stats_o["dimension"] == "om"
    hint = str(sc_o.iloc[0]["存貨偏重提示"])
    assert "存貨偏重 O1" in hint
    _, sc_n, _, _, _, _, _ = transform_hybrid(_df(rows))
    assert str(sc_n.iloc[0]["存貨偏重提示"]) == ""


def test_excel_om_layouts():
    # om（G=2）：群組標籤「OM」、提示第 20 欄、偏重第 23 欄
    raw, sc, diag, diag_om, cross, stats, _ = transform_hybrid(_df(_SCENARIO_OM), dimension="om")
    wb = load_workbook(generate_excel_hybrid(raw, sc, diag, diag_om, cross, stats))
    ws = wb["Summary"]
    assert ws.cell(row=3, column=8).value == "OM"
    assert ws.cell(row=5, column=20).value == "強制調貨提示"
    assert ws.cell(row=5, column=23).value == "存貨偏重提示"
    assert ws.cell(row=4, column=8).value == "Ivy"
    assert ws.cell(row=4, column=12).value == "Max"
    assert str(wb["Review 摘要"].cell(row=19, column=1).value) == "Safety Review × OM 交叉表"
    # nature_om（_SCENARIO_A 全 OM1，G=3）：標籤「Shop Nature＋OM」、提示第 26 欄、偏重第 30 欄
    raw2, sc2, diag2, diag2_om, cross2, stats2, _ = transform_hybrid(_df(_SCENARIO_A), dimension="nature_om")
    ws2 = load_workbook(generate_excel_hybrid(raw2, sc2, diag2, diag2_om, cross2, stats2))["Summary"]
    assert ws2.cell(row=3, column=8).value == "Shop Nature＋OM"
    assert ws2.cell(row=5, column=26).value == "強制調貨提示"
    assert ws2.cell(row=5, column=30).value == "存貨偏重提示"
    assert str(load_workbook(generate_excel_hybrid(raw2, sc2, diag2, diag2_om, cross2, stats2))["Review 摘要"].cell(row=19, column=1).value) == "Safety Review × Shop Nature＋OM 交叉表"


def test_excel_om_fallback_rebuilds_from_raw():
    # generate_excel_hybrid 未傳預建表時，由 stats om_groups（或 raw）重建動態欄位
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_OM), dimension="nature_om")
    assert stats["om_groups"] == ["Ivy", "Max"]
    wb = load_workbook(generate_excel_hybrid(raw, None, None, None, None, stats))
    ws = wb["Summary"]
    assert ws.cell(row=3, column=8).value == "Shop Nature＋OM"
    assert ws.cell(row=4, column=8).value == "Tourist-Ivy"
    # stats 缺 om_groups 時由 raw 即時推導
    stats2 = dict(stats)
    del stats2["om_groups"]
    wb2 = load_workbook(generate_excel_hybrid(raw, None, None, None, None, stats2))
    assert wb2["Summary"].cell(row=4, column=8).value == "Tourist-Ivy"

# ──────────────────────────────────────────────────────────────────────────────
# 22. OM 空白成組與大小寫不敏感（v1.32.0）：空白 OM 自成「空白」群組
#     （部分空白不退回）；OM 大小寫不敏感（Ivy＝ivy，顯示沿用首見原值，
#     同OM調貨跨大小寫可互調）；真實 OM 名恰為「空白」時不誤退回
# ──────────────────────────────────────────────────────────────────────────────
def test_normalize_hybrid_om_blank_group_and_case_insensitive():
    disp, groups = _normalize_hybrid_om(pd.Series(["Ivy", "ivy", " IVY ", None, ""]))
    assert list(disp) == ["Ivy", "Ivy", "Ivy", "空白", "空白"]
    assert groups == ("Ivy", "空白")


def test_mixed_blank_om_forms_blank_group_without_fallback():
    rows = [dict(r) for r in _SCENARIO_OM]
    rows[2]["OM"] = None  # M11 OM 空白 → 空白組
    raw, sc, _, _, cross, stats, warns = transform_hybrid(_df(rows), dimension="om")
    assert stats["dimension"] == "om"  # 部分空白不退回
    assert stats["om_groups"] == ["Ivy", "Max", "空白"]
    assert any("空白" in w for w in warns)
    expected = [c for c in _stock_checking_summary_columns("om", om_groups=("Ivy", "Max", "空白")) if c in sc.columns]
    assert list(sc.columns) == expected
    assert int(sc.iloc[0]["空白 涉及店舖數"]) == 1
    total = cross[cross["Safety Review"] == "總計"].iloc[0]
    assert int(total["空白 列數"]) == 1 and int(total["合計 列數"]) == 3


def test_om_case_insensitive_grouping_and_transfer():
    rows = [dict(r) for r in _SCENARIO_OM]
    rows[2]["OM"] = "ivy"  # M11 與 T11 (Ivy) 同組（大小寫不敏感）
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(rows), dimension="om")
    assert stats["dimension"] == "om"
    assert stats["om_groups"] == ["Ivy", "Max"]  # 顯示沿用首見原值
    assert _get(raw, "M11", "OM") == "Ivy"
    # 同OM調貨跨大小寫成功：T11 需求 9 件由同組 M11（Mix 超額 11 件）全數補足
    assert int(_get(raw, "T11", "實際調入量 (同OM)")) == 9
    assert int(_get(raw, "T11", "調入缺口 (同OM)")) == 0


def test_literal_blank_label_om_does_not_fallback():
    # 真實 OM 名恰為「空白」且來源非空 → 不可誤判為全欄空白而退回
    rows = [dict(r) for r in _SCENARIO_A]
    for r in rows:
        r["OM"] = "空白"
    _, _, _, _, _, stats, _ = transform_hybrid(_df(rows), dimension="om")
    assert stats["dimension"] == "om"
    assert stats["om_groups"] == ["空白"]


# ──────────────────────────────────────────────────────────────────────────────
# 22. 全列預防性 Safety 設定建議（母體為全部列；減少日後調貨）
# ──────────────────────────────────────────────────────────────────────────────
def test_full_safety_covers_all_rows_with_same_basis():
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    for col in ["全列 Safety 周轉所需量", "全列 Safety Review",
                "全列建議 Safety 設定值", "全列 Safety 建議調整量",
                "全列建議 Safety 設定值 Turnday"]:
        assert col in raw.columns
    # 全列所需量與事件版同口徑（預設 Tourist 7 / Mix 9 / Non Tourist 10）
    assert int(_get(raw, "T01", "全列 Safety 周轉所需量")) == 28
    assert int(_get(raw, "N01", "全列 Safety 周轉所需量")) == 10
    # 全列不設調貨事件門檻：T01 現有 30 > 28 → 全列過高（事件版因屬需要調入列而判設定正常）；
    # N01 現有 30 > 10 → 全列過高
    assert _get(raw, "T01", "全列 Safety Review") == "Safety 設定過高"
    assert float(_get(raw, "T01", "全列 Safety 建議調整量")) == -2
    assert _get(raw, "N01", "全列 Safety Review") == "Safety 設定過高"
    assert float(_get(raw, "N01", "全列 Safety 建議調整量")) == -20
    assert stats["full_action_rows"] == stats["full_under_rows"] + stats["full_over_rows"]
    assert stats["full_normal_rows"] == stats["total_raw_rows"] - stats["full_action_rows"]


def test_full_safety_nd_never_increases():
    rows = [dict(r) for r in _SCENARIO_A]
    rows[0]["RP Type"] = "ND"
    rows[0]["Adj. Safety Stock"] = 0
    raw, _, _, _, _, stats, _ = transform_hybrid(_df(rows))
    # T01 ND 且現有 0 < 所需 28 → 仍不作上調（不可補貨）
    assert _get(raw, "T01", "全列 Safety Review") == "設定正常"
    assert pd.isna(_get(raw, "T01", "全列建議 Safety 設定值"))


def test_full_safety_builders_and_excel_sheet():
    from transformer import (
        _build_hybrid_full_safety_change_summary,
        _build_hybrid_full_safety_sku_detail,
        _build_hybrid_full_safety_detail,
    )
    raw, sc, diag, diag_om, ds, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    change = _build_hybrid_full_safety_change_summary(raw)
    assert list(change["分類"]) == ["增加", "減少", "不變", "總計"]
    assert int(change.loc[change["分類"] == "總計", "列數"].iloc[0]) == len(raw)
    sku_detail = _build_hybrid_full_safety_sku_detail(raw)
    assert len(sku_detail) == 1 and int(sku_detail.iloc[0]["全列數"]) == len(raw)
    detail = _build_hybrid_full_safety_detail(raw)
    assert set(detail["Site Code"]) == set(raw.loc[raw["全列 Safety 建議調整量"].fillna(0) != 0, "Site Code"])
    buf = generate_excel_hybrid(raw, sc, diag, diag_om, ds, stats)
    wb = load_workbook(buf, read_only=True)
    assert "Safety 設定建議 (全列)" in wb.sheetnames
    assert "全列 Safety Review" in [c.value for c in next(wb["RAW"].iter_rows(min_row=1, max_row=1))]


@pytest.mark.parametrize("variant", ["free", "om"])
@pytest.mark.parametrize("reverse_rows", [False, True])
def test_detail_source_rp_type_is_sku_specific_in_builder_and_excel(variant, reverse_rows):
    rows = []
    for sku, source_rp in [(820, {"N21": "RF", "M21": "ND"}),
                           (821, {"N21": "ND", "M21": "RF"})]:
        for original in _SCENARIO_MULTI:
            row = dict(original, SKU=sku)
            row["RP Type"] = source_rp.get(row["Site Code"], "RF")
            rows.append(row)
    if reverse_rows:
        rows.reverse()
    raw, sc, diag, diag_om, ds, stats, _ = transform_hybrid(_df(rows))
    before = raw.copy(deep=True)
    detail = _build_hybrid_transfer_detail(raw, variant=variant)
    expected = {
        ("820", "M21"): "ND", ("821", "M21"): "RF",
    }
    if variant == "free":
        expected.update({("820", "N21"): "RF", ("821", "N21"): "ND"})
    actual = {
        (str(row["SKU"]), row["Supply Site Code(調貨來源)"]): row["Supply RP Type"]
        for _, row in detail.iterrows()
    }
    assert actual == expected
    assert list(detail.columns) == _HYBRID_TRANSFER_DETAIL_COLUMNS
    assert detail["Target RP Type"].eq("RF").all()
    assert detail["Target Shop Name"].eq("ShopT").all()
    assert detail["Target OM"].eq("Ivy").all()
    for site, name, om in [("M21", "ShopM", "Ivy"), ("N21", "ShopN", "Max")]:
        supply = detail.loc[detail["Supply Site Code(調貨來源)"] == site]
        assert supply["Supply Shop Name"].eq(name).all()
        assert supply["Supply OM"].eq(om).all()
    qty_col = "實際調入量 (同OM)" if variant == "om" else "實際調入量"
    assert detail["調貨數量"].sum() == raw[qty_col].sum()
    pd.testing.assert_frame_equal(raw, before)

    wb = load_workbook(generate_excel_hybrid(raw, sc, diag, diag_om, ds, stats))
    ws = wb["簡易調貨明細 (同OM)" if variant == "om" else "簡易調貨明細 (不限OM)"]
    assert [cell.value for cell in ws[1]] == _HYBRID_TRANSFER_DETAIL_COLUMNS
    exported = {
        (str(row[1]), row[3]): row[5]
        for row in ws.iter_rows(min_row=2, values_only=True)
    }
    assert exported == expected


@pytest.mark.parametrize("target_nature", ["Tourist", "Mix", "Non Tourist"])
@pytest.mark.parametrize("bidirectional", [False, True])
def test_rp_guide_and_dashboard_match_transfer_rules(target_nature, bidirectional):
    rows = [dict(row) for row in _SCENARIO_RP]
    rows[0]["Shop Nature"] = target_nature
    raw, sc, diag, diag_om, ds, stats, _ = transform_hybrid(
        _df(rows), dimension="rp", bidirectional=bidirectional,
    )
    assert _get(raw, "M01", "需要調貨") == "是"
    assert _get(raw, "M01", "實際調入量") == 36
    assert _get(raw, "M02", "轉出量") == 30
    assert _get(raw, "M02", "可轉出貨量 (Mix)") == 46
    before = raw.copy(deep=True)
    need_metric = _hybrid_dashboard_metrics(raw, stats)[0][1]
    assert need_metric[1] == 1
    assert "RF" in need_metric[2] and "任何 Nature 皆可調入" in need_metric[2]

    wb = load_workbook(generate_excel_hybrid(raw, sc, diag, diag_om, ds, stats))
    ws = wb["說明"]
    pairs = {ws.cell(row, 1).value: ws.cell(row, 2).value for row in range(1, ws.max_row + 1)}
    assert "任何 Nature 的 RF 列" in pairs["分類維度"]
    assert "允許同 Nature 調貨" in pairs["分類維度"]
    assert "不影響調貨" not in pairs["分類維度"]
    if bidirectional:
        assert pairs["允許雙向調貨"].startswith("開啟")
        assert "Tourist 店舖亦可轉出" in pairs["允許雙向調貨"]
    else:
        assert pairs["允許雙向調貨"].startswith("關閉")
        assert "任何 Nature 的 RF 列調入" in pairs["允許雙向調貨"]
        assert "不含 Tourist" in pairs["允許雙向調貨"]
    bullets = [ws.cell(row, 2).value for row in range(1, ws.max_row + 1)
               if ws.cell(row, 1).value == "•"]
    transfer_logic = next(text for text in bullets if text.startswith("調貨段"))
    assert "不套用池級歸零" in transfer_logic
    assert "允許同 Nature 調貨" in transfer_logic
    assert "Tourist 僅在雙向調貨開啟時可作來源" in transfer_logic
    assert "缺貨池的列一律不可轉出" not in transfer_logic
    assert not any("分類維度只改變" in text for text in bullets)
    pd.testing.assert_frame_equal(raw, before)


@pytest.mark.parametrize("bidirectional", [False, True])
def test_rp_guide_tourist_supply_requires_bidirectional(bidirectional):
    rows = [dict(_SCENARIO_RP[0]), {
        "SKU": 400, "Site Code": "T02", "Shop Nature": "Tourist", "RP Type": "ND",
        "Sold Qty (P90D)": 90.0, "Shop Monthly Sold Qty (P90D)": 30.5,
        "Shop Stock Qty (CM)": 60.0, "Adj. Safety Stock": 30,
        "D001 Net Stock Qty (CM)": 10.0,
    }]
    raw, sc, diag, diag_om, ds, stats, _ = transform_hybrid(
        _df(rows), dimension="rp", bidirectional=bidirectional,
    )
    assert _get(raw, "T02", "轉出量") == (36 if bidirectional else 0)
    assert _get(raw, "M01", "實際調入量") == (36 if bidirectional else 0)
    wb = load_workbook(generate_excel_hybrid(raw, sc, diag, diag_om, ds, stats))
    assert any("Tourist 僅在雙向調貨開啟時可作來源" in str(cell.value)
               for row in wb["說明"].iter_rows() for cell in row)


def test_return_d001_guide_warns_about_same_om_switched_donor():
    rows = [dict(row) for row in _SCENARIO_OM]
    rows[-1]["Adj. Safety Stock"] = 0
    raw, sc, diag, diag_om, ds, stats, _ = transform_hybrid(_df(rows), return_to_d001=True)
    assert _get(raw, "M11", "轉出量") == 0
    assert _get(raw, "M11", "轉出量 (同OM)") == 9
    assert _get(raw, "M11", "建議退D001 量") == 25
    assert (_get(raw, "M11", "轉出量 (同OM)") + _get(raw, "M11", "建議退D001 量")
            > _get(raw, "M11", "Shop Stock Qty (CM)"))
    before = raw.copy(deep=True)
    wb = load_workbook(generate_excel_hybrid(raw, sc, diag, diag_om, ds, stats))
    warnings = [str(cell.value) for row in wb["說明"].iter_rows() for cell in row
                if "不能直接與「同OM」調貨合併執行" in str(cell.value)]
    assert len(warnings) == 1
    assert "只扣除「不限OM」轉出量" in warnings[0]
    assert "超過現有庫存" in warnings[0]
    assert "逐店重新核對可退量" in warnings[0]
    pd.testing.assert_frame_equal(raw, before)
