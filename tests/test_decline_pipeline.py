"""Validation tests for Sales Decline Detection functions."""
import warnings
warnings.filterwarnings("ignore")
import pandas as pd
from io import BytesIO
from transformer import (
    _add_decline_columns, _build_decline_summary, transform_decline,
    generate_excel_decline, DECLINE_CATEGORIES, _classify_severity,
)

BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": "1", "Site Code": "S01",
    "Item_Shop": "Prod-S01", "Shop Stock Status": "Normal",
    "Adj. Safety Stock": 5, "Shop Monthly Sold Qty (P90D)": 10,
    "Sold Qty (CM - MTD)": 5.0, "Shop Nature": "Local",
    "RP Type": "RF", "Shop Stock Qty (CM)": 3, "Adj. MCH1": "CatX",
    "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 60, "OM": "OM1",
}


def _row(**kw):
    r = dict(BASE); r.update(kw); return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def test_classify_severity():
    assert _classify_severity(5) == "None"
    assert _classify_severity(15) == "Mild"
    assert _classify_severity(30) == "Moderate"
    assert _classify_severity(60) == "Severe"
    print("OK _classify_severity")


def test_decline_detection():
    rows = [
        # r0: 100->60 (40%), 60->40 proj (33%) -> Moderate
        {"SKU": "1", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 60, "Sold Qty (CM - MTD)": 20.0},
        # r1: 100->90 (10%), 90->60 proj (33%) -> Moderate (max=33%)
        {"SKU": "2", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 90, "Sold Qty (CM - MTD)": 30.0},
        # r2: 50->20 (60%), 20->10 proj (50%) -> Severe (max=60%)
        {"SKU": "3", "Sold Qty (M-2)": 50, "Sold Qty (M-1)": 20, "Sold Qty (CM - MTD)": 5.0},
        # r3: 200->200 (0%), 200->400 proj -> not declining
        {"SKU": "4", "Sold Qty (M-2)": 200, "Sold Qty (M-1)": 200, "Sold Qty (CM - MTD)": 200.0},
        # r4: 30->10 (67%), 10->4 proj (60%) -> Severe
        {"SKU": "5", "Sold Qty (M-2)": 30, "Sold Qty (M-1)": 10, "Sold Qty (CM - MTD)": 2.0},
        # r5: M-2=0 -> insufficient data
        {"SKU": "6", "Sold Qty (M-2)": 0, "Sold Qty (M-1)": 50, "Sold Qty (CM - MTD)": 10.0},
        # r6: all zeros -> not declining
        {"SKU": "7", "Sold Qty (M-2)": 0, "Sold Qty (M-1)": 0, "Sold Qty (CM - MTD)": 0.0},
        # r7: 1000->100 (90% drop), but 100->1000 proj (increase) -> not declining
        {"SKU": "8", "Sold Qty (M-2)": 1000, "Sold Qty (M-1)": 100, "Sold Qty (CM - MTD)": 500.0},
    ]
    df = _df(rows)
    raw = _add_decline_columns(df, mtd_days=15, month_total_days=30)

    # Row 0: Moderate
    assert raw.iloc[0]["Is Declining"] == True
    assert raw.iloc[0]["Decline Severity"] == "Moderate", f"Got {raw.iloc[0]['Decline Severity']}"
    print(f"OK r0: severity={raw.iloc[0]['Decline Severity']}")

    # Row 1: Moderate
    assert raw.iloc[1]["Is Declining"] == True
    assert raw.iloc[1]["Decline Severity"] == "Moderate", f"Got {raw.iloc[1]['Decline Severity']}"
    print(f"OK r1: severity={raw.iloc[1]['Decline Severity']}")

    # Row 2: Severe
    assert raw.iloc[2]["Is Declining"] == True
    assert raw.iloc[2]["Decline Severity"] == "Severe", f"Got {raw.iloc[2]['Decline Severity']}"
    print(f"OK r2: severity={raw.iloc[2]['Decline Severity']}")

    # Row 3: not declining (flat)
    assert raw.iloc[3]["Is Declining"] == False
    print("OK r3: not declining (flat)")

    # Row 4: Severe
    assert raw.iloc[4]["Is Declining"] == True
    assert raw.iloc[4]["Decline Severity"] == "Severe", f"Got {raw.iloc[4]['Decline Severity']}"
    print(f"OK r4: severity={raw.iloc[4]['Decline Severity']}")

    # Row 5: insufficient data
    assert raw.iloc[5]["Is Declining"] == False
    assert raw.iloc[5]["Decline Severity"] == "Insufficient Data"
    print("OK r5: insufficient data")

    # Row 6: no data
    assert raw.iloc[6]["Is Declining"] == False
    assert raw.iloc[6]["Decline Severity"] == "None"
    print("OK r6: no sales")

    # Row 7: MTD increased, not declining
    assert raw.iloc[7]["Is Declining"] == False
    print("OK r7: MTD increased (not declining)")


def test_build_decline_summary():
    rows = [
        {"SKU": "1", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 60, "Sold Qty (CM - MTD)": 20.0},
        {"SKU": "2", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 90, "Sold Qty (CM - MTD)": 30.0},
        {"SKU": "3", "Sold Qty (M-2)": 50, "Sold Qty (M-1)": 20, "Sold Qty (CM - MTD)": 5.0},
        {"SKU": "5", "Sold Qty (M-2)": 30, "Sold Qty (M-1)": 10, "Sold Qty (CM - MTD)": 2.0},
    ]
    df = _df(rows)
    raw = _add_decline_columns(df, mtd_days=15, month_total_days=30)

    site_summary = _build_decline_summary(raw, ["Site Code"])
    assert len(site_summary) > 1
    total = site_summary[site_summary["Site Code"] == "總計"]
    assert len(total) == 1
    assert int(total["衰退行數"].values[0]) == 4
    assert int(total["中度衰退"].values[0]) == 2
    assert int(total["重度衰退"].values[0]) == 2
    assert int(total["輕度衰退"].values[0]) == 0
    print("OK _build_decline_summary (site)")

    brand_summary = _build_decline_summary(raw, ["Brand"])
    assert len(brand_summary) > 0
    cols = [c.encode("ascii", "replace").decode() for c in brand_summary.columns]
    print(f"OK brand summary: {cols}")


def test_transform_decline():
    rows = [
        {"SKU": "1", "Adj. MCH1": "CatA", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 60, "Sold Qty (CM - MTD)": 20.0},
        {"SKU": "2", "Adj. MCH1": "CatA", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 90, "Sold Qty (CM - MTD)": 30.0},
        {"SKU": "3", "Adj. MCH1": "CatB", "Sold Qty (M-2)": 50, "Sold Qty (M-1)": 20, "Sold Qty (CM - MTD)": 5.0},
        {"SKU": "4", "Adj. MCH1": "CatB", "Sold Qty (M-2)": 30, "Sold Qty (M-1)": 10, "Sold Qty (CM - MTD)": 2.0},
        {"SKU": "5", "Adj. MCH1": "CatA", "Sold Qty (M-2)": 200, "Sold Qty (M-1)": 200, "Sold Qty (CM - MTD)": 200.0},
    ]
    df = _df(rows)
    site, brand, buyer, cat, raw, cats, warns = transform_decline(
        df, mtd_days=15, month_total_days=30
    )

    assert cats == ["Mild", "Moderate", "Severe"]
    assert len(site) > 0
    assert len(brand) > 0
    assert len(buyer) > 0
    assert len(cat) > 0
    total_site = site[site["Site Code"] == "總計"]
    assert int(total_site["衰退行數"].values[0]) == 4
    print("OK transform_decline (all 4 dimensions)")


def test_generate_excel_decline():
    rows = [
        {"SKU": "1", "Adj. MCH1": "CatA", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 60, "Sold Qty (CM - MTD)": 20.0},
        {"SKU": "2", "Adj. MCH1": "CatB", "Sold Qty (M-2)": 50, "Sold Qty (M-1)": 20, "Sold Qty (CM - MTD)": 5.0},
        {"SKU": "3", "Adj. MCH1": "CatA", "Sold Qty (M-2)": 200, "Sold Qty (M-1)": 200, "Sold Qty (CM - MTD)": 200.0},
    ]
    df = _df(rows)
    site, brand, buyer, cat, raw, cats, _ = transform_decline(
        df, mtd_days=15, month_total_days=30
    )
    buf = generate_excel_decline(site, brand, buyer, cat, raw, cats)
    assert buf is not None
    assert isinstance(buf, BytesIO)
    data = buf.getvalue()
    assert len(data) > 100
    assert data[:2] == b"PK"
    print("OK generate_excel_decline (valid xlsx)")


def test_missing_mch1():
    """When Adj. MCH1 is missing, category summary should be empty."""
    df = _df([
        {"SKU": "1", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 60, "Sold Qty (CM - MTD)": 20.0},
    ])
    df = df.drop(columns=["Adj. MCH1"], errors="ignore")
    site, brand, buyer, cat, raw, cats, warns = transform_decline(
        df, mtd_days=15, month_total_days=30
    )
    assert cat.empty
    assert any("Adj. MCH1" in w for w in warns)
    print("OK missing Adj. MCH1 handled gracefully")


def test_no_declining_items():
    """When no items are declining, summaries should be empty DataFrames."""
    rows = [
        {"SKU": "1", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 100, "Sold Qty (CM - MTD)": 100.0},
        {"SKU": "2", "Sold Qty (M-2)": 50, "Sold Qty (M-1)": 60, "Sold Qty (CM - MTD)": 70.0},
    ]
    df = _df(rows)
    site, brand, buyer, cat, raw, cats, warns = transform_decline(
        df, mtd_days=15, month_total_days=30
    )
    assert site.empty
    print("OK no declining items: empty summaries")


def test_projected_mtd_calculation():
    """Verify projected MTD calculation."""
    df = _df([
        {"SKU": "1", "Sold Qty (M-2)": 100, "Sold Qty (M-1)": 50, "Sold Qty (CM - MTD)": 25.0},
    ])
    raw = _add_decline_columns(df, mtd_days=15, month_total_days=30)
    # projected = ceil(25 / 15 * 30) = ceil(50) = 50
    assert int(raw.iloc[0]["Projected MTD Sold Qty"]) == 50
    print("OK projected MTD: 25/15*30 = ceil(50) = 50")


if __name__ == "__main__":
    test_classify_severity()
    test_decline_detection()
    test_build_decline_summary()
    test_transform_decline()
    test_generate_excel_decline()
    test_missing_mch1()
    test_no_declining_items()
    test_projected_mtd_calculation()
    print("\nALL DECLINE VALIDATION TESTS PASSED")
