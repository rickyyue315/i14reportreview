"""Validation tests for 績效儀表板 Review (segment 8)."""
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from transformer import (
    transform_perf_dashboard,
    generate_excel_perf_dashboard,
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1,
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Mix", "RP Type": "RF",
    "Shop Stock Status": "Normal",
    "Shop Stock Qty (CM)": 5.0, "Sold Qty (P90D)": 30.0,
    "Net Sales (M-2)": 100.0, "Net Sales (M-1)": 100.0, "Net Sales (CM - MTD)": 100.0,
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def test_three_dimensions_and_abc():
    df = _df([
        {"SKU": 1, "Buyer": "B1", "Brand": "A", "Net Sales (CM - MTD)": 800.0},
        {"SKU": 2, "Buyer": "B2", "Brand": "B", "Net Sales (CM - MTD)": 150.0},
        {"SKU": 3, "Buyer": "B3", "Brand": "C", "Net Sales (CM - MTD)": 50.0},
    ])
    by_buyer, by_brand, by_om, raw, sales_col, _ = transform_perf_dashboard(df)
    assert sales_col == "Net Sales (CM - MTD)"
    assert list(by_buyer[by_buyer["Buyer"] != "總計"]["Buyer"]) == ["B1", "B2", "B3"]
    assert list(by_buyer[by_buyer["Buyer"] != "總計"]["ABC"]) == ["A", "B", "C"]
    assert len(by_brand[by_brand["Brand"] != "總計"]) == 3
    assert len(by_om[by_om["OM"] != "總計"]) == 1
    assert len(raw) == 3


def test_oos_overstock_rates():
    df = _df([
        {"SKU": 1, "Shop Stock Qty (CM)": 0.0},
        {"SKU": 2, "Shop Stock Status": "Overstock (P90D 0 Sales)"},
        {"SKU": 3},
        {"SKU": 4},
    ])
    by_buyer, _, _, _, _, _ = transform_perf_dashboard(df)
    total = by_buyer[by_buyer["Buyer"] == "總計"].iloc[0]
    assert abs(float(total["缺貨率"]) - 0.25) < 1e-9
    assert abs(float(total["過剩率"]) - 0.25) < 1e-9


def test_sales_fallback_and_missing_rejected():
    df = _df([{"SKU": 1}]).drop(columns=["Net Sales (CM - MTD)"])
    _, _, _, _, sales_col, warnings = transform_perf_dashboard(df)
    assert sales_col == "Net Sales (M-1)"
    assert any("CM - MTD" in w for w in warnings)
    df2 = _df([{"SKU": 1}]).drop(columns=["Net Sales (CM - MTD)", "Net Sales (M-1)", "Net Sales (M-2)"])
    try:
        transform_perf_dashboard(df2)
    except ValueError as e:
        assert "銷售金額" in str(e)
    else:
        raise AssertionError("應拒絕缺少銷售金額欄的輸入")


def test_missing_site_code_rejected():
    df = pd.DataFrame([{"Buyer": "B1", "Brand": "A", "SKU": 1, "Shop Stock Status": "Normal",
                        "Net Sales (CM - MTD)": 10.0}])
    try:
        transform_perf_dashboard(df)
    except ValueError as e:
        assert "Site Code" in str(e)
    else:
        raise AssertionError("應拒絕缺少 Site Code 的輸入")


def test_missing_om_single_group():
    df = _df([{"SKU": 1}, {"SKU": 2}]).drop(columns=["OM"])
    _, _, by_om, _, _, warnings = transform_perf_dashboard(df)
    assert list(by_om[by_om["OM"] != "總計"]["OM"]) == ["（空白）"]
    assert any("OM" in w for w in warnings)


def test_excel_sheets_and_skip_raw():
    df = _df([{"SKU": 1}])
    by_buyer, by_brand, by_om, raw, _, _ = transform_perf_dashboard(df)
    buf = generate_excel_perf_dashboard(by_buyer, by_brand, by_om, raw, include_raw=True)
    wb = load_workbook(BytesIO(buf.getvalue()), data_only=True)
    assert wb.sheetnames == ["By Buyer", "By Brand", "By OM", "RAW"]
    buf2 = generate_excel_perf_dashboard(by_buyer, by_brand, by_om, raw, include_raw=False)
    wb2 = load_workbook(BytesIO(buf2.getvalue()), data_only=True)
    assert wb2["RAW"]["A2"].value is not None


def test_percent_format():
    df = _df([
        {"SKU": 1, "Buyer": "B1", "Net Sales (CM - MTD)": 800.0},
        {"SKU": 2, "Buyer": "B2", "Net Sales (CM - MTD)": 200.0},
    ])
    by_buyer, _, _, _, _, _ = transform_perf_dashboard(df)
    buf = generate_excel_perf_dashboard(by_buyer, by_buyer, by_buyer, _df([{"SKU": 1}])[:0])
    wb = load_workbook(BytesIO(buf.getvalue()), data_only=True)
    ws = wb["By Buyer"]
    assert ws["C2"].number_format == "0.00%"
    assert abs(ws["C2"].value - 0.8) < 1e-9
