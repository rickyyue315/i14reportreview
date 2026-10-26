"""LLM（M-2）納入 90 日有貨準則的整合驗證測試（v1.93.0）。

規則：來源檔含 `Shop Stock Qty (LLM)` 欄時（2026-10-26 起來源報表），
P90D 0 Sales 相關篩選另須 LLM > 0——M-2 無貨（LLM = 0）的列 90 日窗口內
非全程有貨，不算真坐貨，予以排除；來源檔 Shop Stock Status 不改自身邏輯，
此類列仍標 P90D 0 Sales，排除列數列 warning。舊檔（無 LLM 欄）行為不變。

涵蓋五處：
1. framework `_hybrid_p90d_zero_sales_mask`（hybrid／Stock Checking 轉 ND 判定）
2. overstock_p90d `_prepare_filtered`（Shop回覆／個別店舖）
3. p90d `transform_p90d_zero_sales`（Past 90 Days 0 Sales）
4. forecast 缺 `Shop Stock Qty (M-2)` 時以 LLM 代理缺貨還原
5. shop_simplify「過去90日有無存貨」＝ LLM／LM／CM 任一 > 0
"""
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import pytest

from transformer import (
    _hybrid_p90d_zero_sales_mask,
    transform_overstock_p90d,
    transform_overstock_p90d_shop,
    transform_p90d_zero_sales,
    transform_forecast,
    transform_shop_simplify,
    _SHOP_SIMPLIFY_STOCK_COLUMN,
    _SHOP_SIMPLIFY_STOCK_YES,
    _SHOP_SIMPLIFY_STOCK_NO,
)


# ── 共用基底 ──────────────────────────────────────────────────────────────────
def _op90d_row(**kw):
    r = {
        "Buyer": "126 - Mona Chan", "Brand": "A", "SKU": 111186903002.0,
        "Item_Shop": "(JDU) 111186903002-HAIR TREAMENT OIL, 01, 40PCHB97-新旺角文華",
        "Site Code": "HB97", "OM": "OM1", "Shop Nature": "Tourist",
        "Adj. MCH1": "01 - Normal", "Shop Stock Status": "Overstock (P90D 0 Sales)",
        "Shop Stock Qty (LM)": 9.0, "Shop Stock Qty (CM)": 5.0,
        "Shop Stock Qty (LLM)": 11.0,
    }
    r.update(kw)
    return r


def _p90d_row(**kw):
    r = {
        "Buyer": "126 - Mona Chan", "Brand": "A", "SKU": 111186903002.0,
        "Item_Shop": "(JDU) 111186903002-HAIR TREAMENT OIL, 01, 40PCHB97-新旺角文華",
        "Site Code": "HB97", "Shop Stock Status": "Overstock (P90D 0 Sales)",
        "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0, "Sold Qty (CM - MTD)": 0.0,
        "Sold Qty (P90D)": 0.0,
        "Shop Stock Qty (LM)": 9.0, "Shop Stock Qty (CM)": 5.0,
        "Shop Stock Qty (LLM)": 11.0,
    }
    r.update(kw)
    return r


def _hy_row(**kw):
    r = {
        "Buyer": "126 - Mona Chan", "Brand": "A", "SKU": 1, "Item_Shop": "",
        "Site Code": "S01", "Shop Nature": "Tourist", "RP Type": "RF",
        "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
        "Sold Qty (CM - MTD)": 0.0, "Sold Qty (P90D)": 0.0,
        "Shop Stock Qty (CM)": 5.0, "Shop Stock Qty (LM)": 4.0,
        "Shop Stock Qty (LLM)": 3.0, "Shop Monthly Sold Qty (P90D)": 0.0,
        "Adj. Safety Stock": 3.0,
    }
    r.update(kw)
    return r


# ── 1+2. P90D 0 Sales 篩選（Shop回覆／個別店舖）────────────────────────────────
def test_op90d_llm_zero_excluded_with_warning():
    df = pd.DataFrame([
        _op90d_row(SKU=1),  # LLM > 0 → 符合
        _op90d_row(SKU=2, **{"Shop Stock Qty (LLM)": 0.0}),  # LLM = 0 → 排除
        _op90d_row(SKU=3, **{"Shop Stock Qty (LLM)": np.nan}),  # 空白視為 0 → 排除
        _op90d_row(SKU=4, **{"Shop Stock Qty (LLM)": -1.0}),  # 負數 → 排除
    ])
    raw, _, _, _, _, warns = transform_overstock_p90d(df)
    assert sorted(str(x) for x in raw["SKU"]) == ["1"]
    assert any("Shop Stock Qty (LLM) 為 0" in w and "3 列" in w for w in warns)


def test_op90d_shop_same_rule():
    df = pd.DataFrame([
        _op90d_row(SKU=1),
        _op90d_row(SKU=2, **{"Shop Stock Qty (LLM)": 0.0}),
    ])
    raw, _, _, _, _, warns = transform_overstock_p90d_shop(df)
    assert [str(x) for x in raw["SKU"]] == ["1"]
    assert any("LLM" in w for w in warns)


def test_op90d_cm_zero_excluded_with_warning():
    # SKU1 情境：LLM > 0、LM > 0、CM = 0（可能已調走）→ 不計算 90 日 0 銷售
    df = pd.DataFrame([
        _op90d_row(SKU=1),
        _op90d_row(SKU=2, **{"Shop Stock Qty (CM)": 0.0}),
        _op90d_row(SKU=3, **{"Shop Stock Qty (CM)": np.nan}),
    ])
    raw, _, _, _, _, warns = transform_overstock_p90d(df)
    assert [str(x) for x in raw["SKU"]] == ["1"]
    assert any("Shop Stock Qty (CM) 為 0" in w and "2 列" in w for w in warns)


def test_op90d_lm_zero_excluded_silently():
    # SKU2 情境：LLM > 0、LM = 0、CM > 0（LM 期間調走、CM 再派貨）→ 不計算；
    # LM 為 0 依既有行為不列 warning（非 LLM／CM 排除原因）
    df = pd.DataFrame([
        _op90d_row(SKU=1, **{"Shop Stock Qty (LM)": 0.0}),
    ])
    raw, _, _, _, _, warns = transform_overstock_p90d(df)
    assert raw.empty
    # 只會有「沒有符合資料」的空結果提示，不會有 CM／LLM 排除原因 warning
    assert not any("已排除" in w for w in warns)


def test_op90d_cm_takes_precedence_over_llm_in_warnings():
    # CM 與 LLM 同時為 0：只計入 CM 排除（兩類 warning 互斥不重複）
    df = pd.DataFrame([
        _op90d_row(SKU=1, **{"Shop Stock Qty (CM)": 0.0, "Shop Stock Qty (LLM)": 0.0}),
    ])
    raw, _, _, _, _, warns = transform_overstock_p90d(df)
    assert raw.empty
    assert any("Shop Stock Qty (CM) 為 0" in w and "1 列" in w for w in warns)
    assert not any("Shop Stock Qty (LLM) 為 0" in w for w in warns)


def test_op90d_without_llm_column_unchanged():
    # 舊檔（無 LLM 欄）：LM > 0 即納入，無 LLM warning
    df = pd.DataFrame([_op90d_row(SKU=1)]).drop(columns=["Shop Stock Qty (LLM)"])
    raw, _, _, _, _, warns = transform_overstock_p90d(df)
    assert [str(x) for x in raw["SKU"]] == ["1"]
    assert not any("LLM" in w for w in warns)


# ── 3. Past 90 Days 0 Sales 模式 ──────────────────────────────────────────────
def test_p90d_llm_zero_excluded_with_warning():
    df = pd.DataFrame([
        _p90d_row(SKU=1),
        _p90d_row(SKU=2, **{"Shop Stock Qty (LLM)": 0.0}),
        _p90d_row(SKU=3, **{"Shop Stock Qty (LLM)": np.nan}),
    ])
    raw, _, warns = transform_p90d_zero_sales(df)
    assert [str(x) for x in raw["SKU"]] == ["1"]
    assert any("LLM" in w and "2 列" in w for w in warns)


def test_p90d_without_llm_column_unchanged():
    df = pd.DataFrame([_p90d_row(SKU=1)]).drop(columns=["Shop Stock Qty (LLM)"])
    raw, _, warns = transform_p90d_zero_sales(df)
    assert [str(x) for x in raw["SKU"]] == ["1"]
    assert not any("LLM" in w for w in warns)


# ── 4. hybrid／Stock Checking 轉 ND 判定 ─────────────────────────────────────
def test_hybrid_mask_llm_zero_excluded():
    df = pd.DataFrame([
        _hy_row(SKU=1),
        _hy_row(SKU=2, **{"Shop Stock Qty (LLM)": 0.0}),
    ])
    mask, warns = _hybrid_p90d_zero_sales_mask(df)
    assert mask.tolist() == [True, False]
    assert warns == []


def test_hybrid_mask_without_llm_column_unchanged():
    df = pd.DataFrame([_hy_row(SKU=1)]).drop(columns=["Shop Stock Qty (LLM)"])
    mask, warns = _hybrid_p90d_zero_sales_mask(df)
    assert mask.tolist() == [True]
    assert warns == []


# ── 5. Forecast Review：缺 M-2 欄時以 LLM 代理缺貨還原 ───────────────────────
def _fc_row(**kw):
    r = {
        "Buyer": "126 - Mona Chan", "Brand": "A", "SKU": 1, "Site Code": "HB97",
        "Shop Stock Status": "Normal", "RP Type": "RF", "Shop Nature": "Tourist",
        "Sold Qty (M-2)": 10.0, "Sold Qty (M-1)": 10.0, "Sold Qty (CM - MTD)": 3.0,
        "Sold Qty (P90D)": 30.0, "Shop Monthly Sold Qty (P90D)": 30.0,
        "Shop Stock Qty (CM)": 5.0, "Adj. Safety Stock": 3.0,
        "Shop Stock Qty (LLM)": 0.0,
    }
    r.update(kw)
    return r


def test_forecast_m2_restored_via_llm_proxy():
    # 無 Shop Stock Qty (M-2) 欄、有 LLM = 0（M-2 缺貨）→ M-2 期還原
    # 還原值 = max(該期實績, 月均基準 Shop Monthly Sold Qty (P90D)) = max(10, 30) = 30
    df = pd.DataFrame([_fc_row()])
    _, _, _, raw, stats, warns = transform_forecast(df)
    assert raw["FC M-2 還原值"].iloc[0] == pytest.approx(30.0)
    assert stats["m2_llm_proxy"] is True
    assert stats["restored_rows_m2"] == 1
    assert any("M-2 期缺貨還原改用「Shop Stock Qty (LLM)」代理" in w for w in warns)


def test_forecast_m2_llm_positive_no_restore():
    # LLM > 0（M-2 有貨）→ 不還原，還原值＝原 M-2 實績
    df = pd.DataFrame([_fc_row(**{"Shop Stock Qty (LLM)": 8.0})])
    _, _, _, raw, stats, warns = transform_forecast(df)
    assert raw["FC M-2 還原值"].iloc[0] == pytest.approx(10.0)
    assert stats["m2_llm_proxy"] is True
    assert stats["restored_rows_m2"] == 0
    # 代理 warning 仍會列出（LLM 正在代替 M-2 欄），但 LLM > 0 故不作缺貨還原
    assert any("M-2 期缺貨還原改用「Shop Stock Qty (LLM)」代理" in w for w in warns)


def test_forecast_without_llm_and_m2_skips_period():
    df = pd.DataFrame([_fc_row()]).drop(columns=["Shop Stock Qty (LLM)"])
    _, _, _, raw, stats, warns = transform_forecast(df)
    assert raw["FC M-2 還原值"].iloc[0] == pytest.approx(10.0)
    assert stats["m2_llm_proxy"] is False
    assert stats["restored_rows_m2"] == 0
    assert any("缺少「Shop Stock Qty (M-2)」欄，M-2 期缺貨還原略過" in w for w in warns)


# ── 6. 店舖精簡SKU模式：過去90日有無存貨 ─────────────────────────────────────
def _sl_row(**kw):
    r = {
        "Buyer": "126 - Mona Chan", "Brand": "A", "SKU": 111186903002.0,
        "Item_Shop": "(JDU) 111186903002-HAIR TREAMENT OIL, 01, 40PCHB97-新旺角文華",
        "Site Code": "HB97", "Shop Stock Status": "Normal",
        "Sold Qty (P90D)": 0.0,
        "Shop Stock Qty (LM)": 0.0, "Shop Stock Qty (CM)": 0.0,
        "Shop Stock Qty (LLM)": 0.0,
    }
    r.update(kw)
    return r


def test_shop_simplify_llm_any_positive_is_in_stock():
    raw, _ = transform_shop_simplify(pd.DataFrame([
        _sl_row(**{"Shop Stock Qty (LLM)": 3.0}),  # 只有 LLM 有貨 → 有存貨
        _sl_row(**{"Shop Stock Qty (LLM)": 0.0, "Shop Stock Qty (LM)": 2.0}),  # LM 有貨
        _sl_row(**{"Shop Stock Qty (LLM)": 1.0, "Shop Stock Qty (LM)": 2.0, "Shop Stock Qty (CM)": 3.0}),
    ]))
    got = raw[_SHOP_SIMPLIFY_STOCK_COLUMN].tolist()
    assert got == [_SHOP_SIMPLIFY_STOCK_YES] * 3


def test_shop_simplify_all_zero_no_stock():
    raw, warns = transform_shop_simplify(pd.DataFrame([
        _sl_row(),  # LLM／LM／CM 三欄皆 0 → 無存貨
    ]))
    assert raw[_SHOP_SIMPLIFY_STOCK_COLUMN].tolist() == [_SHOP_SIMPLIFY_STOCK_NO]
    # 三個庫存欄都在：不會有缺欄提示
    assert not any("缺少「Shop Stock Qty" in w for w in warns)


def test_shop_simplify_old_format_without_llm_unchanged():
    # 舊檔：無 LLM 欄，LM 或 CM > 0 即有存貨（原 60 日準則行為）
    df = pd.DataFrame([
        _sl_row(**{"Shop Stock Qty (LM)": 2.0}),
        _sl_row(**{"Shop Stock Qty (CM)": 1.0}),
        _sl_row(),
    ]).drop(columns=["Shop Stock Qty (LLM)"])
    raw, warns = transform_shop_simplify(df)
    got = raw[_SHOP_SIMPLIFY_STOCK_COLUMN].tolist()
    assert got == [_SHOP_SIMPLIFY_STOCK_YES, _SHOP_SIMPLIFY_STOCK_YES, _SHOP_SIMPLIFY_STOCK_NO]
    # 舊檔缺 LLM 欄：列提示說明以 LM／CM 判定（行為同前，僅警示）
    assert any("缺少「Shop Stock Qty (LLM)」" in w for w in warns)


def test_shop_simplify_llm_only_still_works():
    # 新檔只保留 LLM（LM／CM 同時缺失）：以 LLM 判定並列 warning
    df = pd.DataFrame([
        _sl_row(**{"Shop Stock Qty (LM)": None, "Shop Stock Qty (CM)": None,
                   "Shop Stock Qty (LLM)": 5.0}),
    ]).drop(columns=["Shop Stock Qty (LM)", "Shop Stock Qty (CM)"])
    raw, warns = transform_shop_simplify(df)
    assert raw[_SHOP_SIMPLIFY_STOCK_COLUMN].tolist() == [_SHOP_SIMPLIFY_STOCK_YES]
    assert any("缺少「Shop Stock Qty (LM)」" in w and "缺少「Shop Stock Qty (CM)」" not in w for w in warns)


def test_shop_simplify_all_three_stock_columns_missing_raises():
    df = pd.DataFrame([_sl_row()]).drop(
        columns=["Shop Stock Qty (LLM)", "Shop Stock Qty (LM)", "Shop Stock Qty (CM)"]
    )
    with pytest.raises(ValueError, match="過去90日有無存貨"):
        transform_shop_simplify(df)
