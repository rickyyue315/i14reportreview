"""Validation tests for 銷售趨勢 Review (segment 7)."""
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from transformer import (
    transform_sales_trend,
    generate_excel_sales_trend,
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1,
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Mix",
    "Shop Stock Status": "Normal",
    "Shop Stock Qty (CM)": 20.0,
    "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 100.0,
    "Sold Qty (CM - MTD)": 50.0, "Sold Qty (P90D)": 300.0,
    "Shop Monthly Sold Qty (P90D)": 100.0,
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def test_up_flat_down_classification():
    df = _df([
        {"SKU": 1, "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 130.0},  # +30% → 上升
        {"SKU": 2, "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 110.0},  # +10% → 持平
        {"SKU": 3, "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 70.0},  # -30% → 下跌
        {"SKU": 4, "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
         "Sold Qty (CM - MTD)": 0.0},  # 全零 → 排除
    ])
    raw, by_sku, _, _ = transform_sales_trend(df)
    assert set(raw["SKU"].astype(str)) == {"1", "2", "3"}
    got = dict(zip(by_sku[by_sku["Buyer"] != "總計"]["SKU"].astype(str), by_sku[by_sku["Buyer"] != "總計"]["趨勢分類"]))
    assert got == {"1": "上升", "2": "持平", "3": "下跌"}
    assert int(by_sku[by_sku["SKU"] == "1"]["變動率 %"].iloc[0]) == 30
    assert int(by_sku[by_sku["SKU"] == "3"]["變動率 %"].iloc[0]) == -30


def test_zero_base_up():
    df = _df([
        {"SKU": 1, "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 25.0},
    ])
    _, by_sku, _, _ = transform_sales_trend(df)
    row = by_sku[by_sku["Buyer"] != "總計"].iloc[0]
    assert row["趨勢分類"] == "上升"
    assert pd.isna(row["變動率 %"])


def test_sort_up_first():
    df = _df([
        {"SKU": 1, "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 50.0},
        {"SKU": 2, "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 200.0},
    ])
    _, by_sku, _, _ = transform_sales_trend(df)
    data = by_sku[by_sku["Buyer"] != "總計"]
    assert list(data["SKU"].astype(str)) == ["2", "1"]


def test_missing_required_rejected():
    df = pd.DataFrame([{"Buyer": "B1", "Brand": "A", "SKU": 1, "Shop Stock Status": "Normal",
                        "Site Code": "S01", "Sold Qty (M-2)": 5.0, "Sold Qty (M-1)": 6.0}])
    try:
        transform_sales_trend(df)
    except ValueError as e:
        assert "Sold Qty (CM - MTD)" in str(e)
    else:
        raise AssertionError("應拒絕缺少 CM-MTD 欄的輸入")


def test_by_shop_counts():
    df = _df([
        {"SKU": 1, "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 150.0},
        {"SKU": 2, "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 150.0},
        {"SKU": 1, "Site Code": "S02", "Item_Shop": "Prod-S02-ShopKL",
         "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 150.0},
    ])
    _, _, by_shop, _ = transform_sales_trend(df)
    s01 = by_shop[by_shop["Site Code"] == "S01"].iloc[0]
    assert int(s01["涉及 SKU 數"]) == 2
    assert int(s01["上升 SKU 數"]) == 2


def test_excel_sheets_and_skip_raw():
    df = _df([{"SKU": 1, "Sold Qty (M-2)": 100.0, "Sold Qty (M-1)": 150.0}])
    raw, by_sku, by_shop, _ = transform_sales_trend(df)
    buf = generate_excel_sales_trend(raw, by_sku, by_shop, include_raw=True)
    wb = load_workbook(BytesIO(buf.getvalue()), data_only=True)
    assert wb.sheetnames == ["RAW", "By SKU", "By Shop"]
    buf2 = generate_excel_sales_trend(raw, by_sku, by_shop, include_raw=False)
    wb2 = load_workbook(BytesIO(buf2.getvalue()), data_only=True)
    assert wb2["RAW"]["A2"].value is not None


def test_empty_result_warning():
    df = _df([{"SKU": 1, "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0, "Sold Qty (CM - MTD)": 0.0}])
    raw, _, _, warnings = transform_sales_trend(df)
    assert raw.empty
    assert any("沒有符合" in w for w in warnings)
