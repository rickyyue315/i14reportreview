"""Validation tests for Hybrid Past 90 Days 0 Sales ND hint (v1.26.0, ref_date v1.27.0).

口徑：M-2／M-1／CM-MTD／P90D 皆 ≤ 0 且有庫存（LM 欄存在時須雙 >0），
距參考日期 90 天內新品、需要調貨列、ND 列除外；僅店舖問題表提示，不改主表。
參考日期預設為今天，建議傳入 MTD 截止日期。
"""
import warnings
warnings.filterwarnings("ignore")
import datetime
import pandas as pd
from transformer import (
    transform_hybrid,
    _build_hybrid_shop_issues,
    _hybrid_p90d_zero_sales_mask,
    _normalize_ref_date,
    _HYBRID_SHOP_ISSUE_COLUMNS,
    _HYBRID_P90D_ZERO_SALES_ANNOTATION,
)

HY_BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": 1, "SKU Desc": "Test Product",
    "Item_Shop": "Prod-S01-ShopHK", "Site Code": "S01", "OM": "OM1",
    "Shop Nature": "Non Tourist", "Vendor MC": "V1", "Item Status": "A",
    "Adj. SST": "SST1",
    "Launch Date": "2020-01-01", "Item Class": "C1", "RP Type": "RF",
    "Shop Stock Status": "Normal", "Adj. Safety Stock": 5.0,
    "Sold Qty (P90D)": 0.0, "Sold Qty (M-2)": 0.0, "Sold Qty (M-1)": 0.0,
    "Sold Qty (CM - MTD)": 0.0, "Shop Stock Qty (CM)": 8.0,
    "Shop Stock Qty (LM)": 8.0, "Shop Monthly Sold Qty (P90D)": 0.0,
    "D001 Net Stock Qty (CM)": 200.0,
}


def _hy_row(**kw):
    r = dict(HY_BASE)
    r.update(kw)
    return r


def test_mask_four_period_rule():
    df = pd.DataFrame([
        _hy_row(**{"SKU": 1, "Site Code": "S01"}),
        _hy_row(**{"SKU": 2, "Site Code": "S02", "Sold Qty (M-2)": 3.0}),
        _hy_row(**{"SKU": 3, "Site Code": "S03", "Sold Qty (M-1)": 1.0}),
        _hy_row(**{"SKU": 4, "Site Code": "S04", "Sold Qty (CM - MTD)": 2.0}),
        _hy_row(**{"SKU": 5, "Site Code": "S05", "Sold Qty (P90D)": 5.0}),
        _hy_row(**{"SKU": 6, "Site Code": "S06", "Shop Stock Qty (CM)": 0.0}),
    ])
    mask, warns = _hybrid_p90d_zero_sales_mask(df)
    assert warns == []
    assert list(df.loc[mask, "SKU"]) == [1]


def test_mask_lm_rule_and_nan_negative():
    import numpy as np
    df = pd.DataFrame([
        _hy_row(**{"SKU": 1, "Site Code": "S01", "Shop Stock Qty (LM)": 0.0}),
        _hy_row(**{"SKU": 2, "Site Code": "S02", "Sold Qty (M-2)": np.nan,
                   "Sold Qty (M-1)": -5.0, "Shop Stock Qty (LM)": 3.0,
                   "Shop Stock Qty (CM)": 3.0}),
    ])
    mask, _ = _hybrid_p90d_zero_sales_mask(df)
    assert list(df.loc[mask, "SKU"]) == [2]
    df_nolm = pd.DataFrame([_hy_row(**{"SKU": 7, "Site Code": "S07"})]).drop(columns=["Shop Stock Qty (LM)"])
    mask_nolm, _ = _hybrid_p90d_zero_sales_mask(df_nolm)
    assert bool(mask_nolm.iloc[0]) is True


def test_mask_missing_sales_column_warns():
    df = pd.DataFrame([_hy_row(**{"SKU": 1})]).drop(columns=["Sold Qty (M-2)"])
    mask, warns = _hybrid_p90d_zero_sales_mask(df)
    assert not bool(mask.any())
    assert len(warns) == 1 and "Sold Qty (M-2)" in warns[0]


def test_mask_excludes_new_sku_nd_and_need():
    new_date = (datetime.date.today() - datetime.timedelta(days=30)).isoformat()
    df = pd.DataFrame([
        _hy_row(**{"SKU": 1, "Site Code": "S01", "Launch Date": new_date}),
        _hy_row(**{"SKU": 2, "Site Code": "S02", "RP Type": "ND"}),
        _hy_row(**{"SKU": 3, "Site Code": "S03", "Launch Date": "not-a-date"}),
    ])
    mask, _ = _hybrid_p90d_zero_sales_mask(df)
    assert list(df.loc[mask, "SKU"]) == [3]
    df_need = pd.DataFrame([_hy_row(**{"SKU": 4, "Site Code": "S04", "需要調貨": "是"})])
    mask_need, _ = _hybrid_p90d_zero_sales_mask(df_need)
    assert not bool(mask_need.any())


def test_shop_issues_nd_only_row():
    df = pd.DataFrame([
        _hy_row(**{"SKU": 900, "Site Code": "N90", "Shop Nature": "Non Tourist"}),
    ])
    raw, _, _, _, _, _, warns = transform_hybrid(df)
    issues = _build_hybrid_shop_issues(raw)
    assert list(issues.columns[:11]) == list(_HYBRID_SHOP_ISSUE_COLUMNS[:11])
    assert len(issues) == 1
    row = issues.iloc[0]
    assert row["問題類型"] == "ND 退場"
    assert "轉為 ND" in str(row["建議方法"])
    assert any("Past 90 Days 0 Sales" in str(w) for w in warns)


def test_shop_issues_overlap_keeps_original_with_annotation():
    df = pd.DataFrame([
        _hy_row(**{"SKU": 500, "Site Code": "T50", "Shop Nature": "Tourist",
                   "Sold Qty (P90D)": 360.0, "Sold Qty (M-2)": 100.0,
                   "Sold Qty (M-1)": 100.0, "Sold Qty (CM - MTD)": 10.0,
                   "Shop Stock Qty (CM)": 5.0, "Shop Monthly Sold Qty (P90D)": 120.0,
                   "Adj. Safety Stock": 2.0, "D001 Net Stock Qty (CM)": 10.0}),
        _hy_row(**{"SKU": 500, "Site Code": "N50", "Shop Nature": "Non Tourist",
                   "Sold Qty (P90D)": 0.0, "Sold Qty (M-2)": 0.0,
                   "Sold Qty (M-1)": 0.0, "Sold Qty (CM - MTD)": 0.0,
                   "Shop Stock Qty (CM)": 30.0, "Shop Monthly Sold Qty (P90D)": 30.5,
                   "Adj. Safety Stock": 12.0}),
        _hy_row(**{"SKU": 500, "Site Code": "M50", "Shop Nature": "Mix",
                   "Sold Qty (P90D)": 90.0, "Sold Qty (M-2)": 30.0,
                   "Sold Qty (M-1)": 30.0, "Sold Qty (CM - MTD)": 5.0,
                   "Shop Stock Qty (CM)": 25.0, "Shop Monthly Sold Qty (P90D)": 30.5,
                   "Adj. Safety Stock": 10.0}),
    ])
    raw, _, _, _, _, _, _ = transform_hybrid(df)
    issues = _build_hybrid_shop_issues(raw)
    n50 = issues[issues["Site Code"] == "N50"]
    assert len(n50) == 1
    assert n50["問題類型"].iloc[0] == "Safety 設定過高"
    assert _HYBRID_P90D_ZERO_SALES_ANNOTATION in str(n50["建議方法"].iloc[0])
    assert (issues["Site Code"] == "N50").sum() == 1


def test_ref_date_normalization():
    assert _normalize_ref_date(None) == datetime.date.today()
    assert _normalize_ref_date(datetime.date(2026, 8, 3)) == datetime.date(2026, 8, 3)
    assert _normalize_ref_date(datetime.datetime(2026, 8, 3, 12, 30)) == datetime.date(2026, 8, 3)
    assert _normalize_ref_date("2026-08-03") == datetime.date(2026, 8, 3)
    try:
        _normalize_ref_date("not-a-date")
        raise AssertionError("應 raise ValueError")
    except ValueError:
        pass


def test_ref_date_shifts_new_sku_boundary():
    ref = datetime.date(2026, 8, 3)
    df = pd.DataFrame([
        _hy_row(**{"SKU": 1, "Site Code": "S01", "Launch Date": "2026-05-10"}),
        _hy_row(**{"SKU": 2, "Site Code": "S02", "Launch Date": "2026-05-01"}),
    ])
    mask_default, _ = _hybrid_p90d_zero_sales_mask(df, ref_date=ref)
    assert list(df.loc[mask_default, "SKU"]) == [2]
    mask_today, _ = _hybrid_p90d_zero_sales_mask(df)
    assert list(df.loc[mask_today, "SKU"]) == [1, 2]
    mask_str, _ = _hybrid_p90d_zero_sales_mask(df, ref_date="2026-08-03")
    assert list(df.loc[mask_str, "SKU"]) == [2]


def test_transform_hybrid_ref_date_stats_and_shop_issues():
    ref = datetime.date(2026, 8, 3)
    df = pd.DataFrame([
        _hy_row(**{"SKU": 1, "Site Code": "S01", "Launch Date": "2026-05-10"}),
        _hy_row(**{"SKU": 2, "Site Code": "S02", "Launch Date": "2026-05-01"}),
    ])
    raw, _, _, _, _, stats, _ = transform_hybrid(df, ref_date=ref)
    assert stats["ref_date"] == "2026-08-03"
    issues = _build_hybrid_shop_issues(raw, ref_date=ref)
    assert list(issues["SKU"]) == ["2"]
    raw_default, _, _, _, _, stats_default, _ = transform_hybrid(df)
    assert stats_default["ref_date"] == datetime.date.today().isoformat()
