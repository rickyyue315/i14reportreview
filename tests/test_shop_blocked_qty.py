"""Shop Blocked Qty (CM) 自動扣減驗證（v1.110.0）。

來源檔（2026-11-24 起）新增 `Shop Blocked Qty (CM)`＝店舖已安排退貨、系統尚未扣減的件數：
- 所有模式計算時由 `Shop Stock Qty (CM)` 扣走（下限 0）；
- `Shop Stock Value (CM)` 按「扣減後件數 ÷ 原件數」同一比例調減；
- 原欄位保留供對帳；指定店舖變體（外層與內層各清理一次）不得重複扣減。
"""
import warnings
warnings.filterwarnings("ignore")

import pandas as pd
import pytest

from transformer import (
    _SHOP_BLOCKED_QTY_COLUMN,
    _SHOP_STOCK_QTY_COLUMN,
    _SHOP_STOCK_VALUE_COLUMN,
    _clean_dataframe,
    shop_blocked_qty_summary,
    transform_p90d_zero_sales,
    transform_raw_convert,
    transform_shop_simplify_site,
)

BASE = {
    "Buyer": "127 - Grace Wong", "Brand": "A", "SKU": 1, "SKU Desc": "Test Product",
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Mix", "Vendor MC": "V1", "Item Status": "Active",
    "Launch Date": "2024-01-01", "Item Class": "D", "RP Type": "RF",
    "Shop Stock Status": "Normal", "Adj. Safety Stock": 3.0,
    "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0, "Sold Qty (CM - MTD)": 0.0,
    "Sold Qty (P90D)": 0.0, "Shop Stock Qty (LM)": 5.0,
    "Shop Stock Qty (CM)": 5.0, "Shop Stock Value (CM)": 500.0,
    "Shop Blocked Qty (CM)": 0,
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def test_clean_deducts_qty_and_adjusts_value_proportionally():
    df = _df([
        {"SKU": 1, "Shop Stock Qty (CM)": 3.0, "Shop Blocked Qty (CM)": 2, "Shop Stock Value (CM)": 443.85},
        {"SKU": 2, "Shop Stock Qty (CM)": 3.0, "Shop Blocked Qty (CM)": 3, "Shop Stock Value (CM)": 443.85},
        {"SKU": 3, "Shop Stock Qty (CM)": 5.0, "Shop Blocked Qty (CM)": 0, "Shop Stock Value (CM)": 739.75},
    ])
    out = _clean_dataframe(df)

    assert out[_SHOP_STOCK_QTY_COLUMN].tolist() == [1.0, 0.0, 5.0]
    assert out[_SHOP_STOCK_VALUE_COLUMN].tolist() == pytest.approx([147.95, 0.0, 739.75])
    # 待退件數原樣保留，供與 PowerBI 對帳
    assert out[_SHOP_BLOCKED_QTY_COLUMN].tolist() == [2, 3, 0]
    # 不得改動呼叫端持有的原始資料（session 內的 df_input）
    assert df[_SHOP_STOCK_QTY_COLUMN].tolist() == [3.0, 3.0, 5.0]
    assert df[_SHOP_STOCK_VALUE_COLUMN].tolist() == [443.85, 443.85, 739.75]


def test_clean_without_blocked_column_is_unchanged():
    df = _df([{"SKU": 1}]).drop(columns=[_SHOP_BLOCKED_QTY_COLUMN])
    out = _clean_dataframe(df)
    assert out[_SHOP_STOCK_QTY_COLUMN].iloc[0] == 5.0
    assert out[_SHOP_STOCK_VALUE_COLUMN].iloc[0] == 500.0


def test_clean_with_all_zero_blocked_is_unchanged():
    df = _df([{"SKU": 1}])
    out = _clean_dataframe(df)
    assert out[_SHOP_STOCK_QTY_COLUMN].iloc[0] == 5.0
    assert out[_SHOP_STOCK_VALUE_COLUMN].iloc[0] == 500.0


def test_clean_missing_value_column_only_deducts_qty():
    df = _df([{"SKU": 1, "Shop Blocked Qty (CM)": 2}]).drop(columns=[_SHOP_STOCK_VALUE_COLUMN])
    out = _clean_dataframe(df)
    assert out[_SHOP_STOCK_QTY_COLUMN].iloc[0] == 3.0


def test_blocked_greater_than_stock_clamped_and_flagged():
    df = _df([{
        "SKU": 1, "Shop Stock Qty (CM)": 2.0,
        "Shop Blocked Qty (CM)": 5, "Shop Stock Value (CM)": 100.0,
    }])
    info = shop_blocked_qty_summary(df)
    assert info["available"] is True
    assert info["rows"] == 1
    assert info["blocked_qty"] == 5.0
    assert info["over_rows"] == 1
    assert info["value_available"] is True

    out = _clean_dataframe(df)
    assert out[_SHOP_STOCK_QTY_COLUMN].iloc[0] == 0.0
    assert out[_SHOP_STOCK_VALUE_COLUMN].iloc[0] == 0.0


def test_summary_counts_rows_qty_and_value_deduction():
    df = _df([
        {"SKU": 1, "Shop Stock Qty (CM)": 3.0, "Shop Blocked Qty (CM)": 2, "Shop Stock Value (CM)": 443.85},
        {"SKU": 2, "Shop Stock Qty (CM)": 3.0, "Shop Blocked Qty (CM)": 3, "Shop Stock Value (CM)": 443.85},
        {"SKU": 3, "Shop Blocked Qty (CM)": 0},
    ])
    info = shop_blocked_qty_summary(df)
    assert info["rows"] == 2
    assert info["blocked_qty"] == 5.0
    assert info["value_deduction"] == pytest.approx(295.9 + 443.85)
    assert info["over_rows"] == 0


def test_summary_without_blocked_column_not_available():
    df = _df([{"SKU": 1}]).drop(columns=[_SHOP_BLOCKED_QTY_COLUMN])
    info = shop_blocked_qty_summary(df)
    assert info["available"] is False
    assert info["rows"] == 0


def test_summary_excludes_total_and_footer_rows():
    df = _df([
        {"SKU": 1, "Shop Stock Qty (CM)": 3.0, "Shop Blocked Qty (CM)": 2, "Shop Stock Value (CM)": 300.0},
        # 尾部 Total 列與明細重複，不得計入
        {"SKU": 2, "Item_Shop": "Total", "Shop Stock Qty (CM)": 3.0,
         "Shop Blocked Qty (CM)": 2, "Shop Stock Value (CM)": 300.0},
        # Applied filters 註腳列
        {"SKU": 3, "Item_Shop": "Applied filters: xxx", "Shop Stock Qty (CM)": 99.0,
         "Shop Blocked Qty (CM)": 99, "Shop Stock Value (CM)": 999.0},
    ])
    info = shop_blocked_qty_summary(df)
    assert info["rows"] == 1
    assert info["blocked_qty"] == 2.0
    assert info["value_deduction"] == pytest.approx(200.0)


def test_p90d_zero_sales_excludes_fully_blocked_row():
    df = _df([
        {"SKU": 1, "Item_Shop": "A-S01-ShopHK"},
        {"SKU": 2, "Item_Shop": "B-S01-ShopHK", "Shop Stock Qty (CM)": 5.0, "Shop Blocked Qty (CM)": 5},
    ])
    raw, _, _ = transform_p90d_zero_sales(df)
    assert list(raw["SKU"]) == ["1"]


def test_raw_convert_outputs_netted_values_and_keeps_blocked_column():
    df = _df([{"SKU": 1, "Shop Stock Qty (CM)": 3.0, "Shop Blocked Qty (CM)": 2,
               "Shop Stock Value (CM)": 300.0}])
    raw, _ = transform_raw_convert(df)
    assert raw[_SHOP_STOCK_QTY_COLUMN].iloc[0] == 1.0
    assert raw[_SHOP_STOCK_VALUE_COLUMN].iloc[0] == 100.0
    assert raw[_SHOP_BLOCKED_QTY_COLUMN].iloc[0] == 2


def test_shop_simplify_site_deducts_only_once():
    df = _df([
        {"SKU": 1, "Shop Stock Qty (CM)": 3.0, "Shop Blocked Qty (CM)": 2},
        {"SKU": 2, "Shop Stock Qty (CM)": 3.0, "Shop Blocked Qty (CM)": 0},
    ])
    raw, _ = transform_shop_simplify_site(df, "S01")
    net = dict(zip(raw["SKU"].astype(str), raw[_SHOP_STOCK_QTY_COLUMN].astype(float)))
    # 外層已扣一次（3-2=1）；內層若再扣會變 0
    assert net == {"1": 1.0, "2": 3.0}
