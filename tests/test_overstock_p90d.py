"""Validation tests for Overstock (P90D 0 Sales) Review.

覆蓋：狀態篩選、MAKE UP 分類、店舖統計、Top 10、Nature／OM 分佈、
不輸出 Value／COGS／GP 欄位、Excel 四頁結構。
"""
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from transformer import (
    transform_overstock_p90d,
    generate_excel_overstock_p90d,
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1,
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Mix", "RP Type": "RF", "Item Status": "正常",
    "Adj. SST": "SST1", "Vendor MC": "V1", "Adj. MCH1": "01 - SKIN CARE",
    "Shop Stock Status": "Overstock (P90D 0 Sales)",
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


def _no_banned(df):
    import re
    for c in df.columns:
        lowered = str(c).lower()
        assert "value" not in lowered, c
        assert "cogs" not in lowered, c
        assert re.search(r"\bgp\b", lowered) is None, c


# ── 1. 狀態篩選 ──────────────────────────────────────────────────────────
def test_filter_only_overstock_p90d_status():
    df = _df([
        {"SKU": 1},  # 符合
        {"SKU": 2, "Shop Stock Status": " Overstock (P90D 0 Sales) "},  # 前後空格 → 符合
        {"SKU": 3, "Shop Stock Status": "Normal"},  # 排除
        {"SKU": 4, "Shop Stock Status": "Overstock (>2*Monthly Sold Qty & >2pcs)"},  # 排除
        {"SKU": 5, "Shop Stock Status": None},  # 排除
    ])
    raw, _, _, _, _, _ = transform_overstock_p90d(df)
    assert set(raw["SKU"].astype(str)) == {"1", "2"}


def test_empty_result_warning():
    df = _df([{"SKU": 1, "Shop Stock Status": "Normal"}])
    raw, by_shop, top10, by_nature, by_om, warnings = transform_overstock_p90d(df)
    assert raw.empty
    assert by_shop.empty and top10.empty and by_nature.empty and by_om.empty
    assert any("Overstock (P90D 0 Sales)" in w for w in warnings)


def test_missing_required_columns_rejected():
    for col in ("Site Code", "Shop Stock Qty (CM)", "Adj. MCH1"):
        df = _df([{"SKU": 1}]).drop(columns=[col])
        try:
            transform_overstock_p90d(df)
        except ValueError as e:
            assert col in str(e)
        else:
            raise AssertionError(f"缺少 {col} 時應拋出 ValueError")


# ── 2. MAKE UP 分類 ──────────────────────────────────────────────────────
def test_makeup_classification():
    df = _df([
        {"SKU": 1, "Adj. MCH1": "02 - MAKE UP"},
        {"SKU": 2, "Adj. MCH1": " 02 - MAKE UP "},
        {"SKU": 3, "Adj. MCH1": "01 - SKIN CARE"},
        {"SKU": 4, "Adj. MCH1": None},
    ])
    raw, _, _, _, _, _ = transform_overstock_p90d(df)
    by_sku = dict(zip(raw["SKU"].astype(str), raw["MAKE UP 分類"]))
    assert by_sku == {"1": "MAKE UP", "2": "MAKE UP", "3": "Non-MAKE UP", "4": "Non-MAKE UP"}


# ── 3. By Shop ───────────────────────────────────────────────────────────
def test_by_shop_counts_and_totals():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Adj. MCH1": "02 - MAKE UP", "Shop Stock Qty (CM)": 10.0},
        {"SKU": 1, "Site Code": "S01", "Adj. MCH1": "02 - MAKE UP", "Shop Stock Qty (CM)": 10.0},  # 同 SKU 重複列
        {"SKU": 2, "Site Code": "S01", "Adj. MCH1": "01 - SKIN CARE", "Shop Stock Qty (CM)": 5.0},
        {"SKU": 3, "Site Code": "S02", "Item_Shop": "Prod-S02-ShopKL",
         "OM": "OM2", "Shop Nature": "Tourist", "Adj. MCH1": "02 - MAKE UP", "Shop Stock Qty (CM)": 7.0},
    ])
    _, by_shop, _, _, _, _ = transform_overstock_p90d(df)
    s01 = by_shop[by_shop["Site Code"] == "S01"].iloc[0]
    assert int(s01["Overstock SKU 總數"]) == 2  # SKU 1、2（已刪除重複）
    assert int(s01["MAKE UP SKU 數"]) == 1
    assert int(s01["Non-MAKE UP SKU 數"]) == 1
    assert int(s01["Shop Stock Qty (CM) 總數"]) == 25  # 未刪除重複加總
    total = by_shop[by_shop["Site Code"] == "總計"].iloc[0]
    assert int(total["Overstock SKU 總數"]) == 3
    assert int(total["MAKE UP SKU 數"]) == 2
    assert int(total["Non-MAKE UP SKU 數"]) == 1
    sites = list(by_shop["Site Code"])
    assert sites[-1] == "總計"
    _no_banned(by_shop)


# ── 4. Top 10 ────────────────────────────────────────────────────────────
def test_top10_per_shop_limit_and_rank():
    rows = [{"SKU": i, "Shop Stock Qty (CM)": float(i)} for i in range(1, 13)]
    df = _df(rows)
    _, _, top10, _, _, _ = transform_overstock_p90d(df)
    assert len(top10) == 10
    assert list(top10["店內排名"]) == list(range(1, 11))
    assert list(top10["Shop Stock Qty (CM)"]) == [12.0, 11.0, 10.0, 9.0, 8.0, 7.0, 6.0, 5.0, 4.0, 3.0]
    assert (top10["店舖回覆"] == "").all()
    assert "MAKE UP 分類" in top10.columns and "Adj. MCH1" in top10.columns
    _no_banned(top10)


def test_top10_multi_shop():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Shop Stock Qty (CM)": 9.0},
        {"SKU": 2, "Site Code": "S02", "Item_Shop": "P-S02-KL", "Shop Stock Qty (CM)": 3.0},
    ])
    _, _, top10, _, _, _ = transform_overstock_p90d(df)
    assert len(top10) == 2
    assert set(top10["Site Code"]) == {"S01", "S02"}


# ── 5. Nature／OM 分佈 ───────────────────────────────────────────────────
def test_nature_and_om_distribution():
    df = _df([
        {"SKU": 1, "Site Code": "S01", "Shop Nature": "Tourist", "OM": "OM1",
         "Adj. MCH1": "02 - MAKE UP", "Shop Stock Qty (CM)": 6.0},
        {"SKU": 2, "Site Code": "S02", "Item_Shop": "P-S02-KL", "Shop Nature": "Mix", "OM": "OM1",
         "Shop Stock Qty (CM)": 4.0},
        {"SKU": 3, "Site Code": "S03", "Item_Shop": "P-S03-NT", "Shop Nature": "Non Tourist", "OM": "OM2",
         "Shop Stock Qty (CM)": 2.0},
    ])
    _, _, _, by_nature, by_om, _ = transform_overstock_p90d(df)
    assert set(by_nature[by_nature["Shop Nature"] != "總計"]["Shop Nature"]) == {"Tourist", "Mix", "Non Tourist"}
    n_total = by_nature[by_nature["Shop Nature"] == "總計"].iloc[0]
    assert int(n_total["Overstock 行數"]) == 3
    assert int(n_total["MAKE UP SKU 數"]) == 1
    assert int(n_total["Non-MAKE UP SKU 數"]) == 2
    assert int(n_total["涉及店舖數"]) == 3
    assert set(by_om[by_om["OM"] != "總計"]["OM"]) == {"OM1", "OM2"}
    o_total = by_om[by_om["OM"] == "總計"].iloc[0]
    assert int(o_total["Overstock 行數"]) == 3
    assert int(o_total["Shop Stock Qty (CM) 總數"]) == 12
    _no_banned(by_nature)
    _no_banned(by_om)


# ── 6. Excel 結構 ────────────────────────────────────────────────────────
def test_excel_four_sheets_no_banned_headers():
    df = _df([
        {"SKU": 1, "Adj. MCH1": "02 - MAKE UP"},
        {"SKU": 2},
    ])
    _, by_shop, top10, by_nature, by_om, _ = transform_overstock_p90d(df)
    buf = generate_excel_overstock_p90d(by_shop, top10, by_nature, by_om, None)
    wb = load_workbook(BytesIO(buf.getvalue()), data_only=True)
    assert wb.sheetnames == ["By Shop", "Top 10 店舖回覆", "By Nature", "By OM"]
    import re
    for ws in wb.worksheets:
        for cell in ws[1]:
            if cell.value:
                lowered = str(cell.value).lower()
                assert "value" not in lowered, cell.value
                assert "cogs" not in lowered, cell.value
                assert re.search(r"\bgp\b", lowered) is None, cell.value


def test_excel_empty_result_still_four_sheets():
    df = _df([{"SKU": 1, "Shop Stock Status": "Normal"}])
    _, by_shop, top10, by_nature, by_om, _ = transform_overstock_p90d(df)
    buf = generate_excel_overstock_p90d(by_shop, top10, by_nature, by_om, None)
    wb = load_workbook(BytesIO(buf.getvalue()), data_only=True)
    assert wb.sheetnames == ["By Shop", "Top 10 店舖回覆", "By Nature", "By OM"]
