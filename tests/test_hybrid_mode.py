"""Validation tests for Safety Stock x NetStock Review mode (調貨事件診斷 Safety 設定, v1.12.0/v1.13.0).

以調貨事件為入口、14 日周轉線為基準診斷 Safety Stock 設定：
- 調貨段與 transform_stock_checking 同規則（SKU 級加總 TurnDay、Nature 供應門檻、14 日目標、NT→Mix）
- 需要調入列：Adj. Safety Stock < 14日周轉所需量 → Safety 設定不足
- 已標示 SKU 的 Mix/NT 有貨列：Adj. Safety Stock > 所需量 → Safety 設定過高
- 建議 Safety 設定值 = 所需量；Safety 建議調整量 = 建議值 − 現有設定（正=增加、負=減少）
- 調貨分兩變體：不限OM（契約欄位）與同OM（來源店與需求店同 OM，欄位加「(同OM)」後綴）；
  OM 限制只影響分配，Safety 診斷與建議調整兩變體相同
- Shop Nature：Other/空白/大小寫差異一律正規化為 Non Tourist
- 輸出 4 工作表：RAW / 調貨診斷 Safety Stock (不限OM) / 調貨診斷 Safety Stock (同OM) / 診斷摘要
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
    raw, diagnosis, diagnosis_om, diag_summary, stats, warns = transform_hybrid(_df(_SCENARIO_A))
    assert stats["need_rows"] == 1
    assert len(warns) == 0
    assert "事件類型" in raw.columns and "Safety 診斷" in raw.columns
    assert "轉出量 (同OM)" in raw.columns
    assert len(diagnosis) == 3 and len(diagnosis_om) == 3
    assert not diag_summary.empty


# ──────────────────────────────────────────────────────────────────────────────
# 2. 調貨段與 Stock Checking 等價（不限OM＝契約欄位）
# ──────────────────────────────────────────────────────────────────────────────
def test_transfer_columns_equivalent_to_stock_checking():
    df = _df(_SCENARIO_A)
    sc_raw, _, sc_stats, _ = transform_stock_checking(df)
    hy_raw, _, _, _, hy_stats, _ = transform_hybrid(df)
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
    raw, diagnosis, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    assert int(_get(raw, "T01", "14日周轉所需量")) == 56  # ceil(360/90×14)
    assert _get(raw, "T01", "事件類型") == "需要調入"
    assert _get(raw, "T01", "Safety 診斷") == "Safety 設定不足"
    assert int(_get(raw, "T01", "建議 Safety 設定值")) == 56
    assert float(_get(raw, "T01", "Safety 建議調整量")) == 26
    assert stats["under_rows"] == 1 and stats["under_skus"] == 1 and stats["under_inc_total"] == 26


def test_donor_over_setting_diagnosis():
    raw, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    assert _get(raw, "N01", "事件類型") == "實際轉出"
    assert _get(raw, "N01", "Safety 診斷") == "Safety 設定過高"  # Adj 30 > 所需 14
    assert int(_get(raw, "N01", "建議 Safety 設定值")) == 14
    assert float(_get(raw, "N01", "Safety 建議調整量")) == -16
    assert stats["over_rows"] == 1 and stats["over_dec_total"] == -16


def test_donor_boundary_equal_is_normal():
    raw, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    assert _get(raw, "M01", "事件類型") == "可轉出未使用"  # 可轉出 11、轉出 0
    assert _get(raw, "M01", "Safety 診斷") == "設定正常—過剩非設定因素"  # Adj 14 = 所需 14
    assert pd.isna(_get(raw, "M01", "建議 Safety 設定值"))
    assert pd.isna(_get(raw, "M01", "Safety 建議調整量"))
    assert stats["normal_rows"] == 1


def test_tourist_normal_setting_when_covered():
    rows = [dict(r) for r in _SCENARIO_A]
    rows[0].update({"SKU": 500, "Site Code": "T03", "Adj. Safety Stock": 60})
    rows[1].update({"SKU": 500, "Site Code": "N03", "Adj. Safety Stock": 10})
    raw, _, _, _, _, _ = transform_hybrid(_df([rows[0], rows[1]]))
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
    raw, _, _, _, stats, _ = transform_hybrid(_df(rows))
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
    raw, _, _, _, _, warns = transform_hybrid(_df(rows))
    assert _get(raw, "T05", "Safety 診斷") == "Safety 設定不足"  # 0 < 14
    assert int(_get(raw, "T05", "建議 Safety 設定值")) == 14
    assert any("Adj. Safety Stock" in w for w in warns)


def test_non_event_rows_blank():
    t02 = _df([{"SKU": 999, "Site Code": "Z01", "Shop Nature": "Non Tourist",
                "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
                "Shop Monthly Sold Qty (P90D)": 30.5}])
    raw2, diag2, _, _, _, _ = transform_hybrid(pd.concat([_df(_SCENARIO_A), t02], ignore_index=True))
    assert _get(raw2, "Z01", "事件類型") == ""
    assert pd.isna(_get(raw2, "Z01", "14日周轉所需量"))
    assert pd.isna(_get(raw2, "Z01", "調貨後庫存"))
    assert pd.isna(_get(raw2, "Z01", "調貨後庫存 (同OM)"))
    assert "Z01" not in set(diag2["Site Code"])


def test_unflagged_sku_surplus_not_in_diagnosis():
    extra = _df([{"SKU": 600, "Site Code": "M04", "Shop Nature": "Mix",
                  "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 40.0,
                  "Shop Monthly Sold Qty (P90D)": 30.5, "Adj. Safety Stock": 30}])
    raw, diagnosis, diagnosis_om, _, _, _ = transform_hybrid(pd.concat([_df(_SCENARIO_A), extra], ignore_index=True))
    assert int(_get(raw, "M04", "可轉出貨量 (Mix)")) == 26  # RAW 仍計算可轉出
    assert "M04" not in set(diagnosis["Site Code"])  # 但 SKU 600 未標示調貨 → 不入診斷主表
    assert "M04" not in set(diagnosis_om["Site Code"])


# ──────────────────────────────────────────────────────────────────────────────
# 4. 同OM調貨變體
# ──────────────────────────────────────────────────────────────────────────────
def test_same_om_donor_switch():
    raw, diagnosis, diagnosis_om, _, stats, _ = transform_hybrid(_df(_SCENARIO_OM))
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
    raw, _, _, _, stats, _ = transform_hybrid(_df(rows))
    assert int(_get(raw, "T21", "實際調入量")) == 9 and int(_get(raw, "T21", "調入缺口")) == 0
    assert int(_get(raw, "T21", "實際調入量 (同OM)")) == 0
    assert int(_get(raw, "T21", "調入缺口 (同OM)")) == 9
    assert stats["total_actual_om"] == 0 and stats["total_gap_om"] == 9
    assert stats["donated_rows_om"] == 0


def test_om_restriction_does_not_alter_safety_diagnosis():
    raw, diagnosis, diagnosis_om, _, _, _ = transform_hybrid(_df(_SCENARIO_OM))
    diag_cols = ["Safety 診斷", "建議 Safety 設定值", "Safety 建議調整量", "14日周轉所需量"]
    a = diagnosis.set_index("Site Code")[diag_cols].sort_index()
    b = diagnosis_om.set_index("Site Code")[diag_cols].sort_index()
    pd.testing.assert_frame_equal(a, b)
    # 入表列集合相同（來源列以可轉出貨量>0 判定，與 OM 分配無關）
    assert sorted(diagnosis["Site Code"]) == sorted(diagnosis_om["Site Code"])


def test_missing_om_column_warns_and_equals_free():
    df = _df(_SCENARIO_A).drop(columns=["OM"])
    raw, diagnosis, diagnosis_om, _, stats, warns = transform_hybrid(df)
    assert any("OM" in w for w in warns)
    assert (raw["OM"] == "(未填OM)").all()  # 缺欄時 RAW／主表補「(未填OM)」群組
    assert list(diagnosis["轉出量"]) == list(diagnosis_om["轉出量"])
    assert stats["nt_transfer_om"] == stats["nt_transfer"]


def test_blank_om_grouped_as_unfilled():
    rows = [dict(r) for r in _SCENARIO_OM]
    rows[2]["OM"] = None  # M11 OM 空白 → (未填OM)，與 T11 (Ivy) 不同組 → 同OM無來源
    raw, _, _, _, stats, warns = transform_hybrid(_df(rows))
    assert any("OM" in w and "空白" in w for w in warns)
    assert int(_get(raw, "T11", "實際調入量 (同OM)")) == 0
    assert int(_get(raw, "M11", "轉出量 (同OM)")) == 0
    assert stats["total_gap_om"] == 9


# ──────────────────────────────────────────────────────────────────────────────
# 5. 上下文欄位（調貨後庫存 / TurnDay）
# ──────────────────────────────────────────────────────────────────────────────
def test_stock_after_and_turnday_context():
    raw, _, _, _, _, _ = transform_hybrid(_df(_SCENARIO_A))
    assert int(_get(raw, "T01", "調貨後庫存")) == 56  # 47 + 9
    assert int(_get(raw, "N01", "調貨後庫存")) == 21  # 30 − 9
    assert abs(float(_get(raw, "N01", "調貨後 TurnDay")) - 21.0) < 1e-9
    # 全同 OM 時兩變體相同
    assert int(_get(raw, "T01", "調貨後庫存 (同OM)")) == 56
    assert int(_get(raw, "N01", "調貨後庫存 (同OM)")) == 21
    assert pd.isna(_get(raw, "M01", "建議 Safety 設定值"))  # 正常列不給建議


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
    raw, _, _, _, stats, warns = transform_hybrid(df)
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
    raw, _, _, _, stats, _ = transform_hybrid(df)
    assert (raw["Shop Nature"] == "Tourist").any()
    assert stats["need_rows"] == 1


def test_blank_nature_mapped_to_non_tourist():
    df = _df([
        {"SKU": 1, "Site Code": "X01", "Shop Nature": np.nan,
         "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 30.0,
         "Shop Monthly Sold Qty (P90D)": 30.5},
    ])
    raw, _, _, _, stats, warns = transform_hybrid(df)
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
    raw, diagnosis, diagnosis_om, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 0
    assert (raw["需要調貨"] == "").all()
    assert diagnosis.empty and diagnosis_om.empty


def test_d001_column_fallback():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Turnday (P90D Qty)": 13.0, "D001 Stock Qty (CM)": 50.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)"])
    _, _, _, _, stats, _ = transform_hybrid(df)
    assert stats["need_rows"] == 1


def test_missing_d001_warns():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Sold Qty (P90D)": 90.0, "Shop Stock Qty (CM)": 5.0,
         "Turnday (P90D Qty)": 13.0},
    ]).drop(columns=["D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)"])
    _, _, _, _, stats, warns = transform_hybrid(df)
    assert stats["need_rows"] == 1
    assert any("D001" in w for w in warns)


# ──────────────────────────────────────────────────────────────────────────────
# 8. 診斷主表 / 交叉表 / stats
# ──────────────────────────────────────────────────────────────────────────────
def test_diagnosis_table_rows_and_order():
    raw, diagnosis, _, _, _, _ = transform_hybrid(_df(_SCENARIO_A))
    assert list(diagnosis["Site Code"]) == ["T01", "N01", "M01"]  # 需要調入 → 實際轉出 → 可轉出未使用
    assert list(diagnosis["事件類型"]) == ["需要調入", "實際轉出", "可轉出未使用"]
    assert list(diagnosis["OM"]) == ["OM1", "OM1", "OM1"]  # 主表含 OM 欄供同OM對照
    assert int(diagnosis.loc[diagnosis["Site Code"] == "M01", "可轉出貨量"].iloc[0]) == 11
    for c in _HYBRID_DIAG_COLUMNS:
        assert c in diagnosis.columns


def test_diag_summary_cross_and_totals():
    raw, diagnosis, _, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
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
    raw, _, _, _, stats, _ = transform_hybrid(_df(_SCENARIO_OM))
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
    raw, diagnosis, diagnosis_om, diag_summary, stats, warns = transform_hybrid(df)
    assert stats["need_rows"] == 0
    assert diagnosis.empty and diagnosis_om.empty and diag_summary.empty
    assert any("沒有符合調貨條件" in w for w in warns)
    buf = generate_excel_hybrid(raw, diagnosis, diagnosis_om, diag_summary, stats)
    assert isinstance(buf, BytesIO)


# ──────────────────────────────────────────────────────────────────────────────
# 9. Excel 契約（4 工作表）
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_excel_sheets_and_headers():
    raw, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    buf = generate_excel_hybrid(raw, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    assert wb.sheetnames == [
        "RAW", "調貨診斷 Safety Stock (不限OM)", "調貨診斷 Safety Stock (同OM)", "診斷摘要",
    ]
    raw_headers = [c.value for c in next(wb["RAW"].iter_rows(min_row=1, max_row=1))]
    for col in ["事件類型", "Safety 診斷", "14日周轉所需量", "建議 Safety 設定值",
                "Safety 建議調整量", "調貨後庫存", "調貨後 TurnDay",
                "實際調入量 (同OM)", "轉出量 (同OM)", "調貨來源 (同OM)",
                "調貨後庫存 (同OM)", "調貨後 TurnDay (同OM)", "需要調貨",
                "D001 Effective Stock Qty (CM)"]:
        assert col in raw_headers
    for sheet in ("調貨診斷 Safety Stock (不限OM)", "調貨診斷 Safety Stock (同OM)"):
        headers = [c.value for c in next(wb[sheet].iter_rows(min_row=1, max_row=1))]
        for col in ["Brand", "SKU", "Site Code", "OM", "Shop Nature", "事件類型", "Safety 診斷",
                    "建議 Safety 設定值", "Safety 建議調整量", "可轉出貨量", "轉出量", "調貨後庫存"]:
            assert col in headers, (sheet, col)


def test_generate_excel_dashboard_matches_stats():
    raw, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_OM))
    buf = generate_excel_hybrid(raw, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf, read_only=True)
    ws = wb["診斷摘要"]
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
    raw, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    buf = generate_excel_hybrid(raw, diagnosis, diagnosis_om, diag_summary, stats)
    wb = load_workbook(buf)
    for sheet in ("調貨診斷 Safety Stock (不限OM)", "調貨診斷 Safety Stock (同OM)"):
        ws = wb[sheet]
        assert ws.freeze_panes == "A2"
        assert ws.auto_filter.ref is not None


def test_hybrid_excel_progress_monotonic():
    raw, diagnosis, diagnosis_om, diag_summary, stats, _ = transform_hybrid(_df(_SCENARIO_A))
    events = []
    generate_excel_hybrid(raw, diagnosis, diagnosis_om, diag_summary, stats,
                          progress_callback=lambda p, _m: events.append(float(p)))
    assert all(b >= a for a, b in zip(events, events[1:])), events


# ──────────────────────────────────────────────────────────────────────────────
# 10. 入檔品質回歸（canonicalization / fallback / 數值檢查）
# ──────────────────────────────────────────────────────────────────────────────
def test_header_whitespace_canonicalized():
    df = _df(_SCENARIO_A)
    df = df.rename(columns={c: c + " " for c in df.columns})
    assert validate_hybrid_input(df) == []
    raw, diagnosis, diagnosis_om, _, stats, _ = transform_hybrid(df)
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
    raw, _, _, _, stats, warns = transform_hybrid(df)
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
    raw, _, _, _, stats, warns = transform_hybrid(df)
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
    raw, diagnosis, _, _, _, _ = transform_hybrid(df)
    assert not diagnosis.empty
    assert (diagnosis["SKU Desc"].fillna("").astype(str).str.rstrip("-") == "Prod").all()
