# -*- coding: utf-8 -*-
"""Safety Stock Review 周轉天數基準（可切換選項）測試。

覆蓋：基準計算、計算模式互動、新品保護、ND／P90D 0 Sales 退場、
最低 3 件保底、基準正規化、欄位契約與舊呼叫相容。
"""

import pandas as pd

from transformer import (
    COMBINED_RAW_COLS,
    _add_combined_recommendation_columns,
    _build_combined_shop_issues,
    transform_combined,
)
from modes.safety_review import (
    _normalize_ssr_safety_basis,
    _SSR_SAFETY_BASIS_RATE,
    _SSR_SAFETY_BASIS_TURNDAY,
)


def _base_row(**kw):
    row = {
        "Buyer": "B1",
        "Brand": "BrandA",
        "SKU": "1001",
        "Site Code": "S01",
        "Item_Shop": "Some Product-S01",
        "Shop Stock Status": "Normal",
        "Adj. Safety Stock": 5,
        "Shop Monthly Sold Qty (P90D)": 30,
        "Sold Qty (P90D)": 90,
        "Sold Qty (CM - MTD)": 8.0,
        "Shop Nature": "Tourist",
        "RP Type": "RF",
        "Shop Stock Qty (CM)": 0,
        "Supply Source": "Warehouse",
        "OM": "OM1",
        "Adj. MCH1": "MchA",
        "Adj. SST": "SstA",
    }
    row.update(kw)
    return row


def _df(rows):
    return pd.DataFrame([_base_row(**r) for r in rows])


_RATES_O = {"default": 0.5, "tourist": 0.3, "mix": 0.4}
_RATES_P = {"default": 0.5, "tourist": 0.5, "mix": 0.5}
_DAYS = {"Tourist": 7, "Mix": 9, "Non Tourist": 10}


def _turn(df, **kw):
    params = dict(
        rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30,
        safety_mode=kw.pop("safety_mode", "allow_both"),
        safety_basis="turnday", safety_days=_DAYS,
    )
    params.update(kw)
    return _add_combined_recommendation_columns(df, **params)


def test_normalize_basis():
    assert _normalize_ssr_safety_basis("turnday") == _SSR_SAFETY_BASIS_TURNDAY
    assert _normalize_ssr_safety_basis("TURNDAY") == _SSR_SAFETY_BASIS_TURNDAY
    assert _normalize_ssr_safety_basis("周轉天數") == _SSR_SAFETY_BASIS_TURNDAY
    assert _normalize_ssr_safety_basis("rate") == _SSR_SAFETY_BASIS_RATE
    assert _normalize_ssr_safety_basis("unknown-value") == _SSR_SAFETY_BASIS_RATE
    assert _normalize_ssr_safety_basis(None) == _SSR_SAFETY_BASIS_RATE


def test_raw_columns_include_turnday_fields():
    assert "Safety 基準天數" in COMBINED_RAW_COLS
    assert "Safety 周轉所需量" in COMBINED_RAW_COLS
    assert COMBINED_RAW_COLS.index("Safety 基準天數") > COMBINED_RAW_COLS.index("OOS Floor (Y)")
    assert COMBINED_RAW_COLS.index("Safety 周轉所需量") > COMBINED_RAW_COLS.index("Safety 基準天數")


def test_turnday_line_per_nature():
    df = _df([
        {"SKU": "1", "Shop Nature": "Tourist", "Sold Qty (P90D)": 90},
        {"SKU": "2", "Shop Nature": "Mix", "Sold Qty (P90D)": 90},
        {"SKU": "3", "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 90},
        {"SKU": "4", "Shop Nature": "Other", "Sold Qty (P90D)": 90},
    ])
    out = _turn(df)
    # ceil(90/90*days)：Tourist 7、Mix 9、Non Tourist 10、Other 歸 Non Tourist 10
    assert out["Safety 基準天數"].tolist() == [7, 9, 10, 10]
    assert out["Safety 周轉所需量"].tolist() == [7, 9, 10, 10]


def test_turnday_recommend_allow_both():
    df = _df([
        {"SKU": "1", "Shop Nature": "Tourist", "Sold Qty (P90D)": 90, "Adj. Safety Stock": 2},
        {"SKU": "2", "Shop Nature": "Tourist", "Sold Qty (P90D)": 90, "Adj. Safety Stock": 10},
    ])
    out = _turn(df, safety_mode="allow_both")
    assert out["Recommend Safety Stock"].tolist() == [7, 7]
    assert out["Safety Qty Change"].tolist() == [5, -3]


def test_turnday_increase_only_never_decreases():
    df = _df([
        {"SKU": "1", "Shop Nature": "Tourist", "Sold Qty (P90D)": 90, "Adj. Safety Stock": 10},
        {"SKU": "2", "Shop Nature": "Tourist", "Sold Qty (P90D)": 90, "Adj. Safety Stock": 2},
    ])
    out = _turn(df, safety_mode="increase_only")
    assert out["Recommend Safety Stock"].tolist() == [10, 7]
    assert out["Safety Qty Change"].tolist() == [0, 5]


def test_turnday_floor3():
    df = _df([
        {"SKU": "1", "Shop Nature": "Tourist", "Sold Qty (P90D)": 0,
         "Shop Monthly Sold Qty (P90D)": 6, "Adj. Safety Stock": 0},
    ])
    out = _turn(df, safety_mode="allow_both")
    assert out["Safety 周轉所需量"].tolist() == [3]
    assert out["Recommend Safety Stock"].tolist() == [3]
    assert bool(out["_safety_floor3"].tolist()[0]) is True


def test_turnday_negative_p90d_as_zero():
    df = _df([
        {"SKU": "1", "Shop Nature": "Tourist", "Sold Qty (P90D)": -50, "Adj. Safety Stock": 0},
    ])
    out = _turn(df, safety_mode="allow_both")
    assert out["Safety 周轉所需量"].tolist() == [3]
    assert out["Recommend Safety Stock"].tolist() == [3]


def test_turnday_nd_override_allow_both():
    df = _df([
        {"SKU": "1", "RP Type": "ND", "Sold Qty (P90D)": 90, "Adj. Safety Stock": 5},
    ])
    out = _turn(df, safety_mode="allow_both")
    assert out["Recommend Safety Stock"].tolist() == [0]
    assert out["Safety Qty Change"].tolist() == [-5]
    assert out["Recommendation of new Safety Qty / ND"].tolist() == ["0"]


def test_turnday_nd_no_override_increase_only():
    df = _df([
        {"SKU": "1", "RP Type": "ND", "Sold Qty (P90D)": 90, "Adj. Safety Stock": 5},
    ])
    out = _turn(df, safety_mode="increase_only")
    # 現有 5 < 周轉所需 7：照常上調；退場歸零不套用
    assert out["Recommend Safety Stock"].tolist() == [7]
    assert out["Recommendation of new Safety Qty / ND"].tolist() == [""]


def test_turnday_p90d0_override_with_new_sku_protection():
    old = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Sold Qty (P90D)": 0, "Adj. Safety Stock": 4, "Launch Date": "2020-01-01"},
    ])
    out_old = _turn(old, safety_mode="allow_both")
    assert out_old["Recommend Safety Stock"].tolist() == [0]
    assert out_old["Recommendation of new Safety Qty / ND"].tolist()[0].startswith("ND")

    new = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "RF",
         "Sold Qty (P90D)": 0, "Adj. Safety Stock": 4, "Launch Date": "2099-01-01"},
    ])
    out_new = _turn(new, safety_mode="allow_both", ref_date="2099-02-01")
    assert out_new["Recommend Safety Stock"].tolist() == [3]
    assert out_new["Recommendation of new Safety Qty / ND"].tolist() == ["New SKU (3 Month)"]


def test_turnday_custom_days():
    df = _df([
        {"SKU": "1", "Shop Nature": "Tourist", "Sold Qty (P90D)": 90, "Adj. Safety Stock": 0},
    ])
    out = _turn(df, safety_days={"Tourist": 10, "Mix": 9, "Non Tourist": 10})
    assert out["Safety 基準天數"].tolist() == [10]
    assert out["Safety 周轉所需量"].tolist() == [10]
    assert out["Recommend Safety Stock"].tolist() == [10]


def test_rate_mode_ignores_turnday_line():
    # 月銷倍率基準下，建議值仍按 X/Y；周轉所需量欄僅供對照
    df = _df([
        {"SKU": "1", "Shop Nature": "Tourist", "Shop Monthly Sold Qty (P90D)": 10,
         "Sold Qty (P90D)": 900, "Sold Qty (CM - MTD)": 5.0, "Adj. Safety Stock": 5},
    ])
    out = _add_combined_recommendation_columns(
        df, _RATES_O, _RATES_P, mtd_days=15, month_total_days=30,
        safety_mode="allow_both", safety_basis="rate", safety_days=_DAYS,
    )
    assert "Safety 周轉所需量" in out.columns
    # X = floor(10*0.3)=3（上限 min(3,5)=3）；Y = ceil((5+0.5*10)*0.5)=ceil(5)=5
    assert out["Overstock Cap (X)"].tolist() == [3]
    assert out["OOS Floor (Y)"].tolist() == [5]
    assert out["Recommend Safety Stock"].tolist() == [5]


def test_transform_combined_backward_compat_and_clamp_warning():
    df = _df([
        {"SKU": "1", "Shop Nature": "Tourist", "Sold Qty (P90D)": 90, "Adj. Safety Stock": 2},
    ])
    # 舊呼叫（無 safety_basis）維持月銷倍率行為
    s, raw, _, warnings = transform_combined(
        df, rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both",
    )
    assert "Safety 基準天數" in raw.columns
    assert "Safety 周轉所需量" in raw.columns
    assert raw["Safety 周轉所需量"].tolist() == [7]

    # 超出範圍的基準天數箝制進 7–10 並列 warning
    _, raw2, _, warnings2 = transform_combined(
        df, rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both",
        safety_basis="turnday", safety_days={"Tourist": 20, "Mix": 9, "Non Tourist": 10},
    )
    assert raw2["Safety 基準天數"].tolist() == [10]
    assert any("箝制" in str(w) for w in warnings2)


def test_turnday_shop_issues_floor_note():
    df = _df([
        {"SKU": "1", "Site Code": "S01", "Shop Nature": "Tourist", "Sold Qty (P90D)": 0,
         "Shop Monthly Sold Qty (P90D)": 6, "Adj. Safety Stock": 0},
    ])
    _, raw, _, _ = transform_combined(
        df, rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both",
        safety_basis="turnday", safety_days=_DAYS,
    )
    issues = _build_combined_shop_issues(raw)
    assert len(issues) == 1
    assert "增加 3 件" in str(issues["建議方法"].iloc[0])
    assert "為保障自動補貨" in str(issues["建議方法"].iloc[0])
