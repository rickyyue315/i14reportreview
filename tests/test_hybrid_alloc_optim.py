"""Validation tests for Safety Stock x NetStock Review — transfer allocation improvements.

涵蓋兩項優化（各為預設關閉的開關，關閉時輸出與現行完全一致）：

1. `urgent_first`（缺貨深度優先）— 同 SKU 內需求列的分配次序由「需求量由大→小」
   改為「店舖周轉日數由淺至深 → 需求量由大→小 → 來源檔列序」。
   解決供給不足時 0 庫存的斷貨店舖反被快滿店舖蓋過而分文未得的問題。
2. `pool_nd_exclude`（池健康度排除 ND 死庫存）— 需求側門檻改以只計 RF 列的
   SKU 級加總 TurnDay 判定，讓 ND 死庫存不再掩蓋該池真實缺貨的 RF 店舖；
   供給側（可轉出池級門檻）維持計全部列，讓 ND 死庫存仍可清走。
"""
import warnings
warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from transformer import (
    transform_hybrid,
    generate_excel_hybrid,
    _hybrid_allocate_transfers,
    _hybrid_raw_display_columns,
    _hybrid_stats_from_raw,
    _HYBRID_RF_TURND_SUFFIX,
)

_RF_SUFFIX = _HYBRID_RF_TURND_SUFFIX
_RF_COLS = tuple(f"{n}{_RF_SUFFIX}" for n in ("Tourist", "Mix", "Non Tourist"))


def _row(sku, nature, site, om, stock, sold, safety, rp="RF", d001=0.0):
    return {
        "Buyer": "B1", "Brand": "BR", "SKU": sku, "SKU Desc": f"Desc {sku}",
        "Shop Nature": nature, "Site Code": site, "OM": om,
        "Shop Stock Qty (CM)": stock, "Sold Qty (P90D)": sold,
        "Adj. Safety Stock": safety,
        "Shop Monthly Sold Qty (P90D)": sold / 90 * 30.5,
        "RP Type": rp, "D001 Net Stock Qty (CM)": d001,
    }


def _df(rows):
    return pd.DataFrame(rows)


def _get(raw, site, col):
    return raw[raw["Site Code"] == site].iloc[0][col]


# ── 場景 ──────────────────────────────────────────────────────────────────
# 供給不足：NT 池健康（庫存 20／月銷 30.5 → TurnDay 20 ≥ 14），超額 6 件；總需求 8 > 供給 6。
# T_A 周轉 5.6 日、需求 6；T_B 周轉 0 日（斷貨）、需求 2。
_SCENE_SHORT = [
    _row("SKU-Q", "Non Tourist", "N1", "OM1", 20, 90, 10),
    _row("SKU-Q", "Tourist", "T_A", "OM1", 4, 64, 7),
    _row("SKU-Q", "Tourist", "T_B", "OM1", 0, 9, 7),
]

# 同周轉日數（T_A／T_B 均 0 庫存、需求皆 4）→ 深度相同，須與現行逐欄一致。
_SCENE_SAME_DEPTH = [
    _row("SKU-R", "Non Tourist", "N1", "OM1", 20, 90, 10),
    _row("SKU-R", "Tourist", "T_A", "OM1", 0, 20, 7),
    _row("SKU-R", "Tourist", "T_B", "OM1", 0, 20, 7),
]

# ND 死庫存掩蓋真實缺貨：Tourist-ND 持 300 件（零銷售、永不補貨）灌高 Tourist 池
# TurnDay 至 300，令同 SKU 的 0 庫存 RF 店舖不被標示；Non Tourist 店有 26 件可用超額。
_SCENE_ND_MASK = [
    _row("SKU-L", "Tourist", "T_ND", "OM1", 300, 0, 7, rp="ND"),
    _row("SKU-L", "Tourist", "T_RF", "OM1", 0, 90, 7),
    _row("SKU-L", "Non Tourist", "N1", "OM1", 40, 90, 10),
]


# ── 1. urgent_first：關閉時維持現行契約 ───────────────────────────────────
def test_urgent_first_off_preserves_current_order():
    raw, *_ = transform_hybrid(_df(_SCENE_SHORT))
    assert _get(raw, "T_A", "需要調貨") == "是"
    assert int(_get(raw, "T_A", "實際調入量")) == 6 and int(_get(raw, "T_A", "調入缺口")) == 0
    assert int(_get(raw, "T_B", "實際調入量")) == 0 and int(_get(raw, "T_B", "調入缺口")) == 2
    assert int(raw["實際調入量"].sum()) == 6


def test_urgent_first_default_matches_explicit_false():
    a, *_ = transform_hybrid(_df(_SCENE_SHORT))
    b, *_ = transform_hybrid(_df(_SCENE_SHORT), urgent_first=False)
    assert a.equals(b)


# ── 2. urgent_first：開啟後最深缺貨先補 ───────────────────────────────────
def test_urgent_first_prioritises_deepest_shortage():
    raw, *_ = transform_hybrid(_df(_SCENE_SHORT), urgent_first=True)
    assert int(_get(raw, "T_B", "實際調入量")) == 2 and int(_get(raw, "T_B", "調入缺口")) == 0
    assert int(_get(raw, "T_A", "實際調入量")) == 4 and int(_get(raw, "T_A", "調入缺口")) == 2
    assert _get(raw, "T_B", "調貨來源 (Site Code)") == "N1:2"
    assert _get(raw, "T_A", "調貨來源 (Site Code)") == "N1:4"


def test_urgent_first_conserves_totals_and_source():
    off, *_ = transform_hybrid(_df(_SCENE_SHORT))
    on, *_ = transform_hybrid(_df(_SCENE_SHORT), urgent_first=True)
    # 總調入量與來源店轉出量守恆（只改分配次序，不改總量）
    assert int(on["實際調入量"].sum()) == int(off["實際調入量"].sum()) == 6
    assert int(_get(on, "N1", "轉出量")) == int(_get(off, "N1", "轉出量")) == 6


def test_urgent_first_leaves_trigger_and_safety_unchanged():
    off, *_ = transform_hybrid(_df(_SCENE_SHORT))
    on, *_ = transform_hybrid(_df(_SCENE_SHORT), urgent_first=True)
    for col in ("需要調貨", "調入需求量 (補到14日)", "可轉出貨量 (Non Tourist)",
                "Safety Review", "建議 Safety 設定值", "Safety 建議調整量"):
        assert off[col].astype(str).equals(on[col].astype(str)), col


def test_urgent_first_same_depth_equals_current():
    """同周轉日數時以需求量為次要鍵，行為與現行完全一致。"""
    off, *_ = transform_hybrid(_df(_SCENE_SAME_DEPTH))
    on, *_ = transform_hybrid(_df(_SCENE_SAME_DEPTH), urgent_first=True)
    assert off["實際調入量"].equals(on["實際調入量"])
    assert off["調貨來源 (Site Code)"].equals(on["調貨來源 (Site Code)"])


def test_urgent_first_applies_to_same_om_variant():
    off, *_ = transform_hybrid(_df(_SCENE_SHORT))
    on, *_ = transform_hybrid(_df(_SCENE_SHORT), urgent_first=True)
    assert int(_get(off, "T_B", "實際調入量 (同OM)")) == 0
    assert int(_get(on, "T_B", "實際調入量 (同OM)")) == 2


def test_urgent_first_recorded_in_stats():
    _, _, _, _, _, stats, _ = transform_hybrid(
        _df(_SCENE_SHORT), urgent_first=True)
    assert stats["urgent_first"] is True
    _, _, _, _, _, stats_off, _ = transform_hybrid(_df(_SCENE_SHORT))
    assert stats_off["urgent_first"] is False


# ── 3. 分配器層級：urgency 排序鍵 ─────────────────────────────────────────
def _alloc(urgency, demand=(0.0, 9.0, 2.0, 0.0), nt=(0.0, 0.0, 0.0, 5.0)):
    return _hybrid_allocate_transfers(
        4,
        np.array(["S"] * 4, dtype=object),
        np.array(["S0", "S1", "S2", "S3"], dtype=object),
        np.array([False, True, True, False]),
        np.array(demand, dtype=float),
        np.array(nt, dtype=float),
        np.zeros(4),
        urgency_arr=urgency,
    )


def test_allocator_urgency_orders_by_shallow_turnday_first():
    actual, gap, donate, desc = _alloc(np.array([0.0, 8.0, 1.0, 0.0]))
    assert int(actual[2]) == 2 and int(gap[2]) == 0
    assert int(actual[1]) == 3 and int(gap[1]) == 6
    assert desc[2] == "S3:2" and desc[1] == "S3:3"
    assert int(donate[3]) == 5


def test_allocator_urgency_none_keeps_volume_order():
    actual, gap, _, _ = _alloc(None)
    assert int(actual[1]) == 5 and int(gap[1]) == 4
    assert int(actual[2]) == 0 and int(gap[2]) == 2


def test_allocator_nonfinite_urgency_sorts_last():
    """周轉日數 inf（＝月銷 0）視為最不急，排在最後而非最前。"""
    actual, _, _, _ = _alloc(np.array([0.0, np.inf, 3.0, 0.0]))
    assert int(actual[2]) == 2 and int(actual[1]) == 3


def test_allocator_urgency_ties_break_by_volume_then_row_order():
    """深度相同 → 需求量為次要鍵（同需求則依來源檔列序）。"""
    actual, _, _, _ = _alloc(np.array([0.0, 5.0, 5.0, 0.0]), demand=(0.0, 4.0, 4.0, 0.0))
    assert int(actual[1]) == 4 and int(actual[2]) == 1


# ── 4. pool_nd_exclude：關閉時維持現行契約 ────────────────────────────────
def test_pool_nd_exclude_off_keeps_shortage_masked():
    raw, *_ = transform_hybrid(_df(_SCENE_ND_MASK))
    assert float(_get(raw, "T_RF", "Tourist 加總 TurnDay")) == 300.0
    assert _get(raw, "T_RF", "需要調貨") != "是"
    assert not any(c in raw.columns for c in _RF_COLS)


def test_pool_nd_exclude_default_matches_explicit_false():
    a, *_ = transform_hybrid(_df(_SCENE_ND_MASK))
    b, *_ = transform_hybrid(_df(_SCENE_ND_MASK), pool_nd_exclude=False)
    cols = [c for c in a.columns]
    assert a[cols].equals(b[cols])


# ── 5. pool_nd_exclude：開啟後釋放被掩蓋的缺貨 ────────────────────────────
def test_pool_nd_exclude_surfaces_masked_shortage():
    raw, *_ = transform_hybrid(_df(_SCENE_ND_MASK), pool_nd_exclude=True)
    assert _get(raw, "T_RF", "需要調貨") == "是"
    assert int(_get(raw, "T_RF", "實際調入量")) == 14
    assert _get(raw, "T_RF", "調貨後庫存") == 14


def test_pool_nd_exclude_writes_rf_turnday_columns():
    raw, *_ = transform_hybrid(_df(_SCENE_ND_MASK), pool_nd_exclude=True)
    assert all(c in raw.columns for c in _RF_COLS)
    # Tourist 排除 ND 後只剩 T_RF（0 庫存、月銷 30.5）→ TurnDay 0
    assert float(_get(raw, "T_RF", f"Tourist{_RF_SUFFIX}")) == 0.0
    # Non Tourist 全為 RF → RF 值與全列值相同
    assert float(_get(raw, "N1", f"Non Tourist{_RF_SUFFIX}")) == float(
        _get(raw, "N1", "Non Tourist 加總 TurnDay"))


def test_pool_nd_exclude_preserves_contract_columns():
    """原有全列「加總 TurnDay」三欄為資料契約，值必須完全不變。"""
    off, *_ = transform_hybrid(_df(_SCENE_ND_MASK))
    on, *_ = transform_hybrid(_df(_SCENE_ND_MASK), pool_nd_exclude=True)
    for label in ("Tourist", "Mix", "Non Tourist"):
        assert off[f"{label} 加總 TurnDay"].equals(on[f"{label} 加總 TurnDay"])
    assert float(_get(on, "T_RF", "Tourist 加總 TurnDay")) == 300.0


def test_pool_nd_exclude_keeps_nd_stock_drainable():
    """供給側維持計全部列：ND 死庫存仍可清走，不會被誤判為池不足而鎖死。"""
    raw, *_ = transform_hybrid(_df(_SCENE_ND_MASK), pool_nd_exclude=True)
    assert int(_get(raw, "N1", "可轉出貨量 (Non Tourist)")) == 26
    # 純 ND 死庫存的來源店仍可被取用（Non Tourist 池以全列口徑 = 40 ≥ 14）
    nd_only, *_ = transform_hybrid(
        _df([
            _row("SKU-M", "Non Tourist", "N_ND", "OM1", 100, 0, 10, rp="ND"),
            _row("SKU-M", "Tourist", "T1", "OM1", 0, 90, 7),
        ]),
        pool_nd_exclude=True,
    )
    assert int(_get(nd_only, "N_ND", "轉出量")) == 14


def test_pool_nd_exclude_only_expands_trigger():
    """觸發面只擴大不收縮：原本標示的列不會變成不標示。"""
    off, *_ = transform_hybrid(_df(_SCENE_ND_MASK))
    on, *_ = transform_hybrid(_df(_SCENE_ND_MASK), pool_nd_exclude=True)
    off_need = set(off.loc[off["需要調貨"] == "是", "Site Code"])
    on_need = set(on.loc[on["需要調貨"] == "是", "Site Code"])
    assert off_need <= on_need


def test_pool_nd_exclude_bidirectional_and_rp_dimension():
    """雙向與 rp 維度的需求側門檻同樣改用 RF 準則。"""
    rows = [
        _row("SKU-B", "Tourist", "T_ND", "OM1", 300, 0, 7, rp="ND"),
        _row("SKU-B", "Tourist", "T_RF", "OM1", 0, 90, 7),
        _row("SKU-B", "Non Tourist", "N1", "OM1", 40, 90, 10),
    ]
    for kwargs in ({"bidirectional": True}, {"dimension": "rp"}):
        off, *_ = transform_hybrid(_df(rows), **kwargs)
        on, *_ = transform_hybrid(_df(rows), pool_nd_exclude=True, **kwargs)
        assert _get(off, "T_RF", "需要調貨") != "是", kwargs
        assert _get(on, "T_RF", "需要調貨") == "是", kwargs


def test_pool_nd_exclude_requires_rp_column():
    rows = [{k: v for k, v in r.items() if k != "RP Type"} for r in _SCENE_ND_MASK]
    off, *_, warns_off = transform_hybrid(_df(rows), pool_nd_exclude=True)
    on, *_, warns_on = transform_hybrid(
        _df(rows), pool_nd_exclude=False)
    assert off["需要調貨"].equals(on["需要調貨"])
    assert not any(c in off.columns for c in _RF_COLS)
    assert any("排除 ND 死庫存" in w and "無法套用" in w for w in warns_off)


def test_pool_nd_exclude_recorded_in_stats():
    _, _, _, _, _, stats, _ = transform_hybrid(
        _df(_SCENE_ND_MASK), pool_nd_exclude=True)
    assert stats["pool_nd_exclude"] is True
    _, _, _, _, _, stats_off, _ = transform_hybrid(_df(_SCENE_ND_MASK))
    assert stats_off["pool_nd_exclude"] is False


# ── 6. RAW 顯示欄序 ──────────────────────────────────────────────────────
def test_rf_columns_display_position():
    raw, *_ = transform_hybrid(_df(_SCENE_ND_MASK), pool_nd_exclude=True)
    disp = _hybrid_raw_display_columns(raw)
    for label in ("Tourist", "Mix", "Non Tourist"):
        assert disp.index(f"{label} 加總 TurnDay") < disp.index(f"{label}{_RF_SUFFIX}")
        assert disp.index(f"{label}{_RF_SUFFIX}") < disp.index("需要調貨")


def test_raw_display_unchanged_when_off():
    raw, *_ = transform_hybrid(_df(_SCENE_ND_MASK))
    disp = _hybrid_raw_display_columns(raw)
    assert not any(_RF_SUFFIX in c for c in disp)
    assert disp.index("Non Tourist 加總 TurnDay") < disp.index("需要調貨")


# ── 7. 說明頁與回退路徑 ──────────────────────────────────────────────────
def test_guide_sheet_lists_both_switches():
    raw, sc, diag, diag_om, ds, stats, _ = transform_hybrid(
        _df(_SCENE_SHORT + [
            _row("SKU-Q", "Tourist", "T_ND", "OM1", 300, 0, 7, rp="ND"),
        ]),
        urgent_first=True, pool_nd_exclude=True,
    )
    ws = load_workbook(generate_excel_hybrid(
        raw, sc, diag, diag_om, ds, stats))["說明"]
    text = "\n".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
    assert "需求排序（缺貨深度優先）" in text
    assert "池健康度排除 ND 死庫存" in text
    assert "缺貨深度優先，本份已開啟" in text
    # 說明頁文字含換行，以去除空白的正規化文字比對
    flat = text.replace("\n", "").replace("\r", "")
    assert "池健康度排除 ND 死庫存（本份已開啟）" in flat
    assert "供給側（可轉出池級門檻）維持計全部列" in flat


def test_stats_from_raw_omits_switch_state():
    """兩個開關無法由 RAW 反推，回退路徑不列（與 nd_priority 準則一致）。"""
    raw, *_ = transform_hybrid(_df(_SCENE_SHORT), urgent_first=True,
                              pool_nd_exclude=True)
    fb = _hybrid_stats_from_raw(raw)
    assert "urgent_first" not in fb
    assert "pool_nd_exclude" not in fb
