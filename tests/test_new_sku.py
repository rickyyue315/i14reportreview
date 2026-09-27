"""Validation tests for 新品追蹤 Review (segment 5)."""
import datetime
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from transformer import (
    transform_new_sku,
    generate_excel_new_sku,
)

REF = datetime.date(2026, 9, 26)
BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1,
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Mix", "RP Type": "RF",
    "Shop Stock Status": "Normal",
    "Launch Date": "2026-09-01",
    "Shop Stock Qty (CM)": 5.0,
    "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
    "Sold Qty (CM - MTD)": 0.0, "Sold Qty (P90D)": 0.0,
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def test_new_window_and_status():
    df = _df([
        {"SKU": 1, "Launch Date": "2026-09-01", "Sold Qty (P90D)": 6.0},  # 25 天 → 新品已起量
        {"SKU": 2, "Launch Date": "2026-06-01"},  # 超過 90 天 → 排除
        {"SKU": 3, "Launch Date": "2026-09-10"},  # 16 天 → 新品未起量
        {"SKU": 4, "Launch Date": "2026-10-01"},  # 未來 → 排除
    ])
    raw, by_sku, _, _ = transform_new_sku(df, ref_date=REF)
    assert set(raw["SKU"].astype(str)) == {"1", "3"}
    got = dict(zip(raw["SKU"].astype(str), raw["起量狀態"]))
    assert got == {"1": "已起量", "3": "未起量"}
    assert int(raw[raw["SKU"] == "1"]["上市日數"].iloc[0]) == 25
    assert len(by_sku[by_sku["Buyer"] != "總計"]) == 2


def test_unparseable_and_future_warnings():
    df = _df([
        {"SKU": 1, "Launch Date": "not-a-date"},
        {"SKU": 2, "Launch Date": "2026-12-01"},
    ])
    raw, _, _, warnings = transform_new_sku(df, ref_date=REF)
    assert raw.empty
    assert any("無法解析" in w for w in warnings)
    assert any("晚於參考日期" in w for w in warnings)


def test_missing_required_rejected():
    df = pd.DataFrame([{"Buyer": "B1", "Brand": "A", "SKU": 1, "Shop Stock Status": "Normal", "Site Code": "S01"}])
    try:
        transform_new_sku(df)
    except ValueError as e:
        assert "Launch Date" in str(e)
    else:
        raise AssertionError("應拒絕缺少 Launch Date 的輸入")


def test_by_shop_counts():
    df = _df([
        {"SKU": 1, "Launch Date": "2026-09-01", "Sold Qty (P90D)": 4.0},
        {"SKU": 2, "Launch Date": "2026-09-05"},
        {"SKU": 1, "Site Code": "S02", "Item_Shop": "Prod-S02-ShopKL", "Launch Date": "2026-09-01"},
    ])
    _, _, by_shop, _ = transform_new_sku(df, ref_date=REF)
    s01 = by_shop[by_shop["Site Code"] == "S01"].iloc[0]
    assert int(s01["新品 SKU 數"]) == 2
    assert int(s01["已起量 SKU 數"]) == 1
    assert int(s01["未起量 SKU 數"]) == 1


def test_ref_date_string_accepted():
    df = _df([{"SKU": 1, "Launch Date": "2026-09-01"}])
    raw, _, _, _ = transform_new_sku(df, ref_date="2026-09-26")
    assert len(raw) == 1


def test_excel_sheets_and_skip_raw():
    df = _df([{"SKU": 1, "Launch Date": "2026-09-01"}])
    raw, by_sku, by_shop, _ = transform_new_sku(df, ref_date=REF)
    buf = generate_excel_new_sku(raw, by_sku, by_shop, include_raw=True)
    wb = load_workbook(BytesIO(buf.getvalue()), data_only=True)
    assert wb.sheetnames == ["RAW", "By SKU", "By Shop"]
    buf2 = generate_excel_new_sku(raw, by_sku, by_shop, include_raw=False)
    wb2 = load_workbook(BytesIO(buf2.getvalue()), data_only=True)
    assert wb2["RAW"]["A2"].value is not None


def test_empty_result_warning():
    df = _df([{"SKU": 1, "Launch Date": "2020-01-01"}])
    raw, _, _, warnings = transform_new_sku(df, ref_date=REF)
    assert raw.empty
    assert any("沒有符合" in w for w in warnings)
