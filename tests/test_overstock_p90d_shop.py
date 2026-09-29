"""Validation tests for P90D 0 Sales Review(個別店舖).

個別店舖版與店舖回覆版篩選準則相同，差異：
- 分類固定以 MAKE UP 分類
- 明細工作表為「店舖90日零銷售ALL SKU」，列出各店舖全部 SKU（不設列數上限），
  包括全部 SKU，不設「店舖回覆」欄
"""
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from transformer import (
    transform_overstock_p90d_shop,
    generate_excel_overstock_p90d_shop,
    _build_op90d_by_sku,
    _build_op90d_by_buyer,
)
from modes.overstock_p90d import (
    _OP90D_ALL_SKU_SHEET,
    _OP90D_ALL_SKU_COLUMNS,
    _OP90D_TOP10_ALL_SHEET,
    _OP90D_TOP10_ALL_COLUMNS,
    _build_op90d_all_sku,
    _build_op90d_top10_all,
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1,
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Mix", "RP Type": "RF", "Item Status": "正常",
    "Adj. SST": "SST1", "Vendor MC": "V1", "Adj. MCH1": "01 - SKIN CARE",
    "Shop Stock Status": "Overstock (P90D 0 Sales)",
    "Supply Source": "Warehouse",
    "Item Class": "D",
    "Adj. Safety Stock": 3.0,
    "Shop Stock Qty (CM)": 5.0, "Shop Stock Qty (LM)": 4.0,
    "Sold Qty (P90D)": 0.0, "Sold Qty (M-2)": 0.0,
    "Sold Qty (M-1)": 0.0, "Sold Qty (CM - MTD)": 0.0,
    "Excess Stock Qty": 2.0, "Shop On Order PO Qty (CM)": 0.0,
    "Shop Stock Value (CM)": 500.0, "Excess Stock Value": 100.0,
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def test_shop_sheet_constant_renamed():
    assert _OP90D_ALL_SKU_SHEET == "店舖90日零銷售ALL SKU"
    assert "Top 10" not in _OP90D_ALL_SKU_SHEET
    # 相容別名與新名一致
    assert _OP90D_TOP10_ALL_SHEET == _OP90D_ALL_SKU_SHEET
    assert _OP90D_TOP10_ALL_COLUMNS == _OP90D_ALL_SKU_COLUMNS


def test_shop_filter_same_as_reply_mode():
    df = _df([
        {"SKU": 1},
        {"SKU": 2, "Shop Stock Status": "Normal"},
        {"SKU": 3, "Shop Stock Qty (LM)": 0.0},
    ])
    raw, _, _, _, _, _ = transform_overstock_p90d_shop(df)
    assert set(raw["SKU"].astype(str)) == {"1"}


def test_shop_all_sku_includes_makeup_no_reply():
    rows = [{"SKU": i, "Shop Stock Qty (CM)": float(100 + i),
             "Adj. MCH1": "02 - MAKE UP"} for i in range(1, 6)]
    rows += [{"SKU": i, "Shop Stock Qty (CM)": float(i),
              "Adj. MCH1": "01 - SKIN CARE"} for i in range(6, 12)]
    df = _df(rows)
    _, _, all_sku, _, _, _ = transform_overstock_p90d_shop(df)
    assert len(all_sku) == 11  # 不設列數上限，全部輸出
    assert set(all_sku["MAKE UP 分類"]) == {"MAKE UP", "Non-MAKE UP"}
    assert list(all_sku["Shop Stock Qty (CM)"])[:5] == [105.0, 104.0, 103.0, 102.0, 101.0]
    assert list(all_sku["店內庫存排名"]) == list(range(1, 12))
    assert "店舖回覆" not in all_sku.columns
    assert list(all_sku.columns) == [c for c in _OP90D_ALL_SKU_COLUMNS if c in all_sku.columns]


def test_shop_all_sku_no_row_limit():
    rows = [{"SKU": i, "Shop Stock Qty (CM)": float(i)} for i in range(1, 13)]
    df = _df(rows)
    raw, _, _, _, _, _ = transform_overstock_p90d_shop(df)
    all_sku = _build_op90d_all_sku(raw)
    assert len(all_sku) == 12
    assert list(all_sku["店內庫存排名"]) == list(range(1, 13))
    assert list(all_sku["Shop Stock Qty (CM)"]) == [float(i) for i in range(12, 0, -1)]
    assert "店舖回覆" not in all_sku.columns
    # 相容別名同樣輸出全部
    assert len(_build_op90d_top10_all(raw)) == 12


def test_shop_all_sku_rank_restarts_per_shop():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Shop Stock Qty (CM)": 9.0},
        {"SKU": 2, "Site Code": "S01", "Shop Stock Qty (CM)": 3.0},
        {"SKU": 3, "Site Code": "S02", "Item_Shop": "P-S02-KL", "Shop Stock Qty (CM)": 7.0},
    ])
    _, _, all_sku, _, _, _ = transform_overstock_p90d_shop(df)
    assert len(all_sku) == 3
    s01 = all_sku[all_sku["Site Code"] == "S01"]
    assert list(s01["店內庫存排名"]) == [1, 2]
    assert list(s01["Shop Stock Qty (CM)"]) == [9.0, 3.0]
    s02 = all_sku[all_sku["Site Code"] == "S02"]
    assert list(s02["店內庫存排名"]) == [1]


def test_shop_by_shop_fixed_makeup_columns():
    df = _df([
        {"SKU": 1, "Adj. MCH1": "02 - MAKE UP", "Shop Stock Qty (CM)": 10.0},
        {"SKU": 2, "Adj. MCH1": "01 - SKIN CARE", "Shop Stock Qty (CM)": 5.0},
    ])
    _, by_shop, _, _, _, _ = transform_overstock_p90d_shop(df)
    assert list(by_shop.columns) == [
        "Site Code", "Shop Name", "OM", "Shop Nature", "P90D 0 Sales SKU 總數",
        "Shop Stock Qty (CM) 總數",
        "MAKE UP SKU 數", "MAKE UP Shop Stock Qty (CM) 總數",
        "Non-MAKE UP SKU 數", "Non-MAKE UP Shop Stock Qty (CM) 總數",
    ]
    s01 = by_shop[by_shop["Site Code"] == "S01"].iloc[0]
    assert int(s01["MAKE UP SKU 數"]) == 1
    assert int(s01["Non-MAKE UP SKU 數"]) == 1


def test_shop_excel_sheet_name_and_no_reply_header():
    df = _df([
        {"SKU": 1, "Adj. MCH1": "02 - MAKE UP"},
        {"SKU": 2},
    ])
    raw, by_shop, all_sku, by_nature, by_om, _ = transform_overstock_p90d_shop(df)
    by_sku = _build_op90d_by_sku(raw)
    by_buyer = _build_op90d_by_buyer(raw)
    buf = generate_excel_overstock_p90d_shop(by_shop, all_sku, by_nature, by_om, None, by_sku=by_sku, by_buyer=by_buyer)
    wb = load_workbook(BytesIO(buf.getvalue()), data_only=True)
    assert wb.sheetnames == ["By Shop", _OP90D_ALL_SKU_SHEET, "By Nature", "By OM", "By SKU", "By Buyer"]
    hdr = [c.value for c in wb[_OP90D_ALL_SKU_SHEET][1]]
    assert "店舖回覆" not in hdr
    assert "MAKE UP 分類" in hdr


def test_shop_empty_result():
    df = _df([{"SKU": 1, "Shop Stock Status": "Normal"}])
    raw, by_shop, all_sku, by_nature, by_om, warnings = transform_overstock_p90d_shop(df)
    assert raw.empty and all_sku.empty
    assert any("P90D 0 Sales" in w for w in warnings)
    buf = generate_excel_overstock_p90d_shop(by_shop, all_sku, by_nature, by_om, None)
    wb = load_workbook(BytesIO(buf.getvalue()), data_only=True)
    assert _OP90D_ALL_SKU_SHEET in wb.sheetnames
