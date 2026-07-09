import pandas as pd
import numpy as np
import unicodedata
import re
import gc
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import datetime
from concurrent.futures import ThreadPoolExecutor

REQUIRED_COLUMNS = ["Buyer", "Brand", "SKU", "Shop Stock Status"]

RAW_COLUMN_ORDER_NAMES = [
    "Item_Shop",
    "SKU",
    "Brand",
    "SKU Desc",
    "Buyer",
    "Adj. SST",
    "Vendor MC",
    "MCH2",
    "Item Status",
    "Launch Date",
    "Item Class",
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "RP Type",
    "建議新RP Type",
    "Buyer 回覆 ND/RF",
    "保留RF原因",
    "Shop Stock Status",
    "Adj. Safety Stock",
    "Shop Monthly Sold Qty (P90D)",
    "Safety VS Monthly Sold Qty",
    "Projected Monthly Sold Qty",
    "Recommendation of new Safety Qty",
    "Safety Qty Increase",
    "Recommendation of new Safety Qty / ND",
    "After new Safety Qty / ND can transfer out of QTY",
    "Overstock Cap (X)",
    "OOS Floor (Y)",
    "Recommend Safety Stock",
    "Safety Qty Change",
    "Buyer reply",
    "D001 Net Stock Qty (CM)",
    "Shop Stock Qty (CM)",
    "Sold Qty (P90D)",
    "Turnday (P90D Qty)",
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Net Sales (M-2)",
    "Net Sales (M-1)",
    "Net Sales (CM - MTD)",
    "Excess Stock Qty",
    "Excess Stock Value",
    "Shop On Order PO Qty (CM)",
    "Shop Pending Receive",
    "Is OOS (M-2)",
    "Is OOS (M-1)",
    "Is OOS (CM)",
    "OOS Periods Count",
    "OOS Pattern",
    "Max Sales Loss (M-2)",
    "Max Sales Loss (P90D)",
    "Total Sales Loss (M-2)",
    "Total Sales Loss (P90D)",
    "Is Stock Impacted",
    "Turnday (Days)",
    "Mix Turnday Band",
    "Mix Excess Stock",
    "Non Tourist Turnday Band",
]


def _resolve_column_name(name: str, available: set[str]) -> str | None:
    stripped = name.strip()
    for col in available:
        if col.strip() == stripped:
            return col
    return None


def _reorder_raw_columns(df: pd.DataFrame) -> pd.DataFrame:
    available = set(df.columns)
    ordered: list[str] = []
    for name in RAW_COLUMN_ORDER_NAMES:
        resolved = _resolve_column_name(name, available)
        if resolved:
            ordered.append(resolved)
    trailing = [c for c in df.columns if c not in ordered]
    return df[ordered + trailing]


def validate_input(df: pd.DataFrame) -> list[str]:
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    return missing


def _safe_sku(x):
    if pd.notna(x):
        try:
            return str(int(float(x)))
        except (ValueError, TypeError):
            return str(x)
    return x


def _safe_numeric(series: pd.Series, default: float = 0.0) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(default)


def _safe_default_series(value, index):
    if len(index) == 0:
        return pd.Series(dtype='object')
    return pd.Series(value, index=index)


def _ensure_list(val):
    if not isinstance(val, list):
        return [val]
    return val


def _extract_sku_desc(item_shop, site_code) -> str:
    if pd.isna(item_shop) or pd.isna(site_code):
        return ""

    item_text = str(item_shop).strip()
    site_text = str(site_code).strip()
    if not item_text or not site_text:
        return ""

    item_text = re.sub(r'^\([^)]+\)\s*', '', item_text)
    item_text = re.sub(r'^\d+\s*-\s*', '', item_text)

    marker = f"{site_text}-"
    idx = item_text.rfind(marker)
    if idx == -1:
        return ""

    return item_text[:idx].strip().strip(",，;；:： ")


def _compute_raw_extra_columns_impl(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    shop_name = [
        _extract_shop_name(item_shop.iloc[i], site_code.iloc[i])
        for i in range(len(df))
    ]
    sku_desc = [
        _extract_sku_desc(item_shop.iloc[i], site_code.iloc[i])
        for i in range(len(df))
    ]
    df["Shop Name"] = shop_name
    df["SKU Desc"] = sku_desc
    return df


try:
    import streamlit as st
    _compute_raw_extra_columns_impl = st.cache_data(_compute_raw_extra_columns_impl)
except ImportError:
    pass


def _compute_raw_extra_columns(df: pd.DataFrame) -> pd.DataFrame:
    return _compute_raw_extra_columns_impl(df)


def _extract_shop_name(item_shop, site_code) -> str:
    if pd.isna(item_shop) or pd.isna(site_code):
        return ""

    item_text = str(item_shop).strip()
    site_text = str(site_code).strip()
    if not item_text or not site_text:
        return ""

    marker = f"{site_text}-"
    idx = item_text.rfind(marker)
    if idx == -1:
        pattern = rf"{re.escape(site_text)}-(.+)$"
        match = re.search(pattern, item_text)
        if not match:
            return ""
        shop_text = match.group(1)
    else:
        shop_text = item_text[idx + len(marker):]

    return shop_text.strip().strip(",，;；:：")


def _build_by_shop_summary(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()

    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    shop_name = [
        _extract_shop_name(item_shop.iloc[i], site_code.iloc[i])
        for i in range(len(df))
    ]

    df["Site Code"] = site_code
    df["Shop Name"] = shop_name
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)

    status_col = df["Shop Stock Status"].fillna("")
    rec_col = df["Recommendation of new Safety Qty / ND"].fillna("").astype(str)
    rec_num = pd.to_numeric(rec_col, errors="coerce")
    transfer_col = "After new Safety Qty / ND can transfer out of QTY"
    transfer_num = pd.to_numeric(df.get(transfer_col, _safe_default_series(0, df.index)), errors="coerce").fillna(0)

    group_cols = ["Site Code", "Shop Name", "OM"]
    df["_is_0_sales"] = status_col == "Overstock (P90D 0 Sales)"
    df["_is_2x"] = status_col == "Overstock (>2*Monthly Sold Qty & >2pcs)"
    df["_is_nd_change"] = df["_is_0_sales"] & (rec_col == "ND")
    df["_is_safety_change"] = df["_is_2x"] & rec_num.notna()
    df["_sku_0_sales"] = df["SKU"].where(df["_is_0_sales"])
    df["_sku_2x"] = df["SKU"].where(df["_is_2x"])
    df["_transfer_num"] = transfer_num

    by_shop = (
        df.groupby(group_cols, dropna=False, observed=True)
        .agg(
            更改NDRF行數=("_is_nd_change", "sum"),
            更改Safety行數=("_is_safety_change", "sum"),
            預計總轉出貨量=("_transfer_num", "sum"),
            **{
                "P90D 0 Sales SKU數量": ("_sku_0_sales", "nunique"),
                ">2*Monthly Sold Qty & >2pcs數量": ("_sku_2x", "nunique"),
            },
        )
        .reset_index()
    )

    by_shop["更改NDRF行數"] = by_shop["更改NDRF行數"].fillna(0).astype(int)
    by_shop["更改Safety行數"] = by_shop["更改Safety行數"].fillna(0).astype(int)
    by_shop["預計總轉出貨量"] = by_shop["預計總轉出貨量"].fillna(0).astype(int)
    by_shop["P90D 0 Sales SKU數量"] = by_shop["P90D 0 Sales SKU數量"].fillna(0).astype(int)
    by_shop[">2*Monthly Sold Qty & >2pcs數量"] = by_shop[">2*Monthly Sold Qty & >2pcs數量"].fillna(0).astype(int)

    if not by_shop.empty:
        total_row = pd.DataFrame(
            {
                "Site Code": ["總計"],
                "Shop Name": [""],
                "OM": [""],
                "更改NDRF行數": [int(by_shop["更改NDRF行數"].sum())],
                "更改Safety行數": [int(by_shop["更改Safety行數"].sum())],
                "預計總轉出貨量": [int(by_shop["預計總轉出貨量"].sum())],
                "P90D 0 Sales SKU數量": [int(by_shop["P90D 0 Sales SKU數量"].sum())],
                ">2*Monthly Sold Qty & >2pcs數量": [int(by_shop[">2*Monthly Sold Qty & >2pcs數量"].sum())],
            }
        )
        by_shop = pd.concat([by_shop.sort_values(["Site Code", "Shop Name", "OM"], kind="stable"), total_row], ignore_index=True)

    return by_shop


def _add_recommendation_columns(df: pd.DataFrame, rates: dict | None = None) -> pd.DataFrame:
    status_col = df["Shop Stock Status"].fillna("")

    adj_ss = _safe_numeric(df.get("Adj. Safety Stock", _safe_default_series(0, df.index)))
    monthly_sold = _safe_numeric(df.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, df.index)))

    # Normalize RP Type (strip + upper) so values like " rf " or "nd" are recognized,
    # consistent with the OOS / Carry review paths.
    rp_type = (
        df.get("RP Type", _safe_default_series("", df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    cond_overstock_2x = (status_col == "Overstock (>2*Monthly Sold Qty & >2pcs)")

    cond_overstock_0 = (
        (status_col == "Overstock (P90D 0 Sales)") & rp_type.isin(["ND", "RF"])
    )

    rec_values = np.full(len(df), "", dtype=object)

    if cond_overstock_2x.any():
        shop_nature = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.lower()
        _rates = rates if rates is not None else {}
        _default_rate = _rates.get("default", 0.50)
        _tourist_rate = _rates.get("tourist", 0.30)
        _mix_rate = _rates.get("mix", 0.40)
        rate = np.full(len(df), _default_rate)
        rate[shop_nature.str.contains("tourist", na=False)] = _tourist_rate
        rate[shop_nature.str.contains("mix", na=False)] = _mix_rate
        rec_2x = np.floor(monthly_sold[cond_overstock_2x] * rate[cond_overstock_2x]).clip(lower=3).astype(int)
        # If new safety qty > Adj. Safety Stock, cap at Adj. Safety Stock
        adj_ss_2x = adj_ss[cond_overstock_2x].values
        over_adj = rec_2x > adj_ss_2x
        rec_2x[over_adj] = adj_ss_2x[over_adj].astype(int)
        rec_values[cond_overstock_2x.values] = rec_2x.astype(str)

    if cond_overstock_0.any():
        is_rf = cond_overstock_0 & (rp_type == "RF")
        is_nd = cond_overstock_0 & (rp_type == "ND")
        rec_values[is_rf.values] = "ND"
        rec_values[is_nd.values] = "0"
    df["Recommendation of new Safety Qty / ND"] = rec_values

    shop_stock = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, df.index)))

    rec = df["Recommendation of new Safety Qty / ND"]

    # Numeric recommendation (from >2x): transfer = shop_stock - rec_num
    rec_num = pd.to_numeric(rec, errors="coerce")
    mask_rec_num = rec_num.notna() & (shop_stock > rec_num)

    # ND recommendation: transfer = all shop stock
    mask_nd = (rec == "ND") & (shop_stock > 0)

    transfer_values = np.full(len(df), "", dtype=object)
    if mask_rec_num.any():
        transfer_values[mask_rec_num.values] = (shop_stock[mask_rec_num] - rec_num[mask_rec_num]).astype(int).values
    if mask_nd.any():
        transfer_values[mask_nd.values] = shop_stock[mask_nd].astype(int).values

    mask_rec_zero = rec_num.notna() & (shop_stock <= rec_num)
    if mask_rec_zero.any():
        transfer_values[mask_rec_zero.values] = 0

    mask_nd_zero = (rec == "ND") & (shop_stock == 0)
    if mask_nd_zero.any():
        transfer_values[mask_nd_zero.values] = 0

    df["After new Safety Qty / ND can transfer out of QTY"] = transfer_values

    df["Buyer reply"] = ""

    return df


def _clean_dataframe(df_input: pd.DataFrame, shop_natures: list[str] | None = None, progress_callback=None) -> pd.DataFrame:
    df = df_input.dropna(subset=["Buyer"])
    if progress_callback:
        progress_callback(0.1, "清理資料中...")
    _footer_mask = df["Buyer"].astype(str).str.strip().str.lower().str.startswith(("total", "applied filters"))
    if "Item_Shop" in df.columns:
        _item_shop_mask = df["Item_Shop"].astype(str).str.strip().str.lower().str.startswith("total")
        _footer_mask = _footer_mask | _item_shop_mask
    df = df[~_footer_mask]
    df = df[df["Buyer"].astype(str).str.strip() != "117 - Marketing"]
    if shop_natures is not None and "Shop Nature" in df.columns:
        sn_col = df["Shop Nature"].fillna("").astype(str).str.strip().str.lower()
        mask = _safe_default_series(False, df.index)
        if "Tourist" in shop_natures:
            mask |= sn_col.str.contains("tourist", na=False)
        if "Mix" in shop_natures:
            mask |= sn_col.str.contains("mix", na=False)
        if "Non Tourist" in shop_natures:
            is_tourist_or_mix = sn_col.str.contains("tourist|mix", na=False, regex=True)
            is_filled = sn_col.notna() & (sn_col.str.strip() != "")
            mask |= is_filled & ~is_tourist_or_mix
        df = df[mask]
    return df


# ─── Safety Stock Review Functions ─────────────────────────────────────────────


def _add_combined_recommendation_columns(
    df, rates_overstock, rates_oos, mtd_days, month_total_days, safety_mode="increase_only"
):
    df = df.copy()

    status_col = df["Shop Stock Status"].fillna("")
    adj_ss = _safe_numeric(df.get("Adj. Safety Stock", _safe_default_series(0, df.index)))
    monthly_sold = _safe_numeric(df.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, df.index)))
    sold_qty = _safe_numeric(df.get("Sold Qty (CM - MTD)", _safe_default_series(0.0, df.index)))
    shop_nature = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.lower()

    cond_2x = (status_col == "Overstock (>2*Monthly Sold Qty & >2pcs)")

    # Overstock cap X
    def _compute_x(monthly_sold_arr, rate_arr, adj_ss_arr):
        x = np.floor(monthly_sold_arr * rate_arr).clip(min=3).astype(int)
        over_adj = x > adj_ss_arr
        x[over_adj] = adj_ss_arr[over_adj].astype(int)
        return x

    _os_default = rates_overstock.get("default", 0.50)
    _os_tourist = rates_overstock.get("tourist", 0.30)
    _os_mix = rates_overstock.get("mix", 0.40)

    os_rate_arr = np.full(len(df), _os_default)
    os_rate_arr[shop_nature.str.contains("tourist", na=False)] = _os_tourist
    os_rate_arr[shop_nature.str.contains("mix", na=False)] = _os_mix

    # X is the upper bound for "how much safety stock is reasonable"; it applies to
    # every row (so the combined recommendation = max(Y, Z if Z <= X else X) is well
    # defined per row), independent of whether any row currently has Overstock status.
    x_values = _compute_x(monthly_sold.values, os_rate_arr, adj_ss.values)

    # OOS floor Y
    if mtd_days and mtd_days > 0:
        projected = (sold_qty / mtd_days) * month_total_days
    else:
        projected = _safe_default_series(0.0, df.index)

    df["Projected Monthly Sold Qty"] = np.ceil(projected).astype(int)

    _oos_default = rates_oos.get("default", 0.50)
    _oos_tourist = rates_oos.get("tourist", 0.50)
    _oos_mix = rates_oos.get("mix", 0.50)

    oos_rate_arr = np.full(len(df), _oos_default)
    oos_rate_arr[shop_nature.str.contains("tourist", na=False)] = _oos_tourist
    oos_rate_arr[shop_nature.str.contains("mix", na=False)] = _oos_mix

    y_values = np.ceil(projected.values * oos_rate_arr).clip(min=3).astype(int)

    # Z = current Adj. Safety Stock
    z_values = adj_ss.values.astype(int)

    # Combined = max(Y, Z if Z <= X else X)
    combined = np.where(z_values <= x_values, np.maximum(y_values, z_values), np.maximum(y_values, x_values))

    df["Overstock Cap (X)"] = x_values
    df["OOS Floor (Y)"] = y_values

    delta = combined - z_values
    if safety_mode == "increase_only":
        # Only allow increases: clamp negative deltas to 0 and keep the recommended
        # value consistent with that (no decrease applied), so Recommend == adj_ss
        # whenever a row would otherwise be lowered.
        delta = delta.clip(min=0)
        recommend_safety = z_values + delta
    else:
        recommend_safety = combined
    df["Recommend Safety Stock"] = recommend_safety
    df["Safety Qty Change"] = delta

    df["Buyer reply"] = ""

    return df


def _build_combined_by_shop_summary(raw):
    df = raw.copy()

    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    shop_name = [
        _extract_shop_name(item_shop.iloc[i], site_code.iloc[i])
        for i in range(len(df))
    ]

    df["Site Code"] = site_code
    df["Shop Name"] = shop_name
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)

    status_col = df["Shop Stock Status"].fillna("")
    safety_change = pd.to_numeric(df.get("Safety Qty Change", _safe_default_series(0, df.index)), errors="coerce").fillna(0)

    group_cols = ["Site Code", "Shop Name", "OM"]
    df["_is_overstock"] = status_col.str.startswith("Overstock", na=False)
    df["_is_oos"] = (status_col == "OOS (0 Stock)") | (status_col == "Potential OOS (<7 Turnday)")
    df["_sku_overstock"] = df["SKU"].where(df["_is_overstock"])
    df["_sku_oos"] = df["SKU"].where(df["_is_oos"])
    df["_safety_increase"] = (safety_change > 0).astype(int)
    df["_safety_decrease"] = (safety_change < 0).astype(int)
    df["_safety_change"] = safety_change

    by_shop = (
        df.groupby(group_cols, dropna=False, observed=True)
        .agg(
            Overstock_SKU数量=("_sku_overstock", "nunique"),
            OOS_SKU数量=("_sku_oos", "nunique"),
            需增加Safety_SKU数量=("_safety_increase", "sum"),
            需减少Safety_SKU数量=("_safety_decrease", "sum"),
            安全库存调整总量=("_safety_change", "sum"),
            Combined_SKU总数=("SKU", "nunique"),
        )
        .reset_index()
    )

    for col in by_shop.columns:
        if col not in group_cols:
            by_shop[col] = by_shop[col].fillna(0).astype(int)

    if not by_shop.empty:
        total_row = pd.DataFrame(
            {
                "Site Code": ["總計"],
                "Shop Name": [""],
                "OM": [""],
                "Overstock_SKU数量": [int(by_shop["Overstock_SKU数量"].sum())],
                "OOS_SKU数量": [int(by_shop["OOS_SKU数量"].sum())],
                "需增加Safety_SKU数量": [int(by_shop["需增加Safety_SKU数量"].sum())],
                "需减少Safety_SKU数量": [int(by_shop["需减少Safety_SKU数量"].sum())],
                "安全库存调整总量": [int(by_shop["安全库存调整总量"].sum())],
                "Combined_SKU总数": [int(by_shop["Combined_SKU总数"].sum())],
            }
        )
        by_shop = pd.concat([by_shop.sort_values(["Site Code", "Shop Name", "OM"], kind="stable"), total_row], ignore_index=True)

    return by_shop


def transform_combined(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
    rates_overstock: dict | None = None,
    rates_oos: dict | None = None,
    mtd_days: int | None = None,
    month_total_days: int | None = None,
    safety_mode: str = "increase_only",
) -> tuple[pd.DataFrame, pd.DataFrame, list[str], list[str]]:
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    warnings = []

    _optional_cols = ["Adj. Safety Stock", "Sold Qty (CM - MTD)", "Shop Monthly Sold Qty (P90D)", "Shop Nature"]
    for _col in _optional_cols:
        if _col not in df_input.columns:
            warnings.append(f"缺少可選欄位「{_col}」，建議值將可能不準確")

    if progress_callback:
        progress_callback(0.05, "驗證通過")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據（非空白或總計列）")

    df["SKU"] = df["SKU"].apply(_safe_sku)

    if "SKU Desc" not in df.columns:
        df = _compute_raw_extra_columns(df)

    raw = _add_combined_recommendation_columns(
        df, rates_overstock, rates_oos, mtd_days, month_total_days, safety_mode=safety_mode
    )

    status_col = raw["Shop Stock Status"].fillna("")
    safety_change = pd.to_numeric(raw.get("Safety Qty Change", _safe_default_series(0, raw.index)), errors="coerce").fillna(0)

    categories = ["Overstock", "OOS", "Balanced (Safety Adjusted)"]

    group_cols = ["Buyer", "Brand", "SKU", "SKU Desc"]

    if progress_callback:
        progress_callback(0.3, "分類統計中...")

    overstock_mask = status_col.str.startswith("Overstock", na=False)
    oos_mask = (status_col == "OOS (0 Stock)") | (status_col == "Potential OOS (<7 Turnday)")
    balanced_mask = (safety_change != 0) & ~overstock_mask & ~oos_mask

    counts = {}
    counts["Overstock"] = (
        raw[overstock_mask].groupby(group_cols, observed=True).size().rename("Overstock")
    )
    counts["OOS"] = (
        raw[oos_mask].groupby(group_cols, observed=True).size().rename("OOS")
    )
    counts["Balanced (Safety Adjusted)"] = (
        raw[balanced_mask].groupby(group_cols, observed=True).size().rename("Balanced (Safety Adjusted)")
    )

    if progress_callback:
        progress_callback(0.5, "合併彙總中...")

    summary = pd.concat(list(counts.values()), axis=1).fillna(0).astype(int).reset_index()
    summary["SKU"] = summary["SKU"].apply(_safe_sku)
    summary["總計"] = summary[categories].sum(axis=1)

    safety_change_per_group = (
        raw.groupby(group_cols, observed=True)["Safety Qty Change"]
        .sum()
        .astype(int)
    )
    summary = summary.merge(safety_change_per_group, on=group_cols, how="left")
    summary["Safety Qty Change"] = summary["Safety Qty Change"].fillna(0).astype(int)

    safety_inc_per_group = (
        raw[raw["Safety Qty Change"] > 0]
        .groupby(group_cols, observed=True)
        .size()
        .rename("增加Safety Qty的店舖數量")
    )
    safety_dec_per_group = (
        raw[raw["Safety Qty Change"] < 0]
        .groupby(group_cols, observed=True)
        .size()
        .rename("減少Safety Qty的店舖數量")
    )
    safety_unchanged_per_group = (
        raw[raw["Safety Qty Change"] == 0]
        .groupby(group_cols, observed=True)
        .size()
        .rename("Safety Qty不變的店舖數量")
    )

    for _series in [safety_inc_per_group, safety_dec_per_group, safety_unchanged_per_group]:
        summary = summary.merge(_series, on=group_cols, how="left")
    for _col in ["增加Safety Qty的店舖數量", "減少Safety Qty的店舖數量", "Safety Qty不變的店舖數量"]:
        summary[_col] = summary[_col].fillna(0).astype(int)

    buyer_order = summary.groupby("Buyer")["總計"].sum().sort_values(ascending=False).index.tolist()
    summary["_buyer_key"] = summary["Buyer"].map({b: i for i, b in enumerate(buyer_order)})
    summary = summary.sort_values(["_buyer_key", "Brand", "SKU"], ascending=[True, True, True]).drop(columns=["_buyer_key"]).reset_index(drop=True)

    total_cat = int(summary["總計"].sum())
    total_row = pd.DataFrame(
        {
            "Buyer": ["總計"],
            "Brand": [None],
            "SKU": [None],
            **{cat: [int(summary[cat].sum())] for cat in categories},
            "總計": [total_cat],
            "Safety Qty Change": [int(summary["Safety Qty Change"].sum())],
            "增加Safety Qty的店舖數量": [int(summary["增加Safety Qty的店舖數量"].sum())],
            "減少Safety Qty的店舖數量": [int(summary["減少Safety Qty的店舖數量"].sum())],
            "Safety Qty不變的店舖數量": [int(summary["Safety Qty不變的店舖數量"].sum())],
        }
    )
    summary = pd.concat([summary, total_row], ignore_index=True)

    if total_cat == 0:
        warnings.append("未偵測到任何符合條件的 Safety Stock Review 資料。")

    if progress_callback:
        progress_callback(0.7, "轉換完成")

    return summary, raw, categories, warnings


def _write_raw_sheet(ws, raw_df, progress_callback=None, progress_start=0.0, progress_end=0.3, columns=None):
    raw_df = _compute_raw_extra_columns(raw_df)
    raw_df = _reorder_raw_columns(raw_df)

    # 選擇性只輸出指定欄位（節省記憶體）
    if columns is not None:
        avail = [c for c in columns if c in raw_df.columns]
        raw_df = raw_df[avail]

    raw_cols = list(raw_df.columns)
    _display_names = {"Adj. Safety Stock": "Safety Stock (現時)"}
    header_row = [_display_names.get(c, c) for c in raw_cols]
    ws.append(header_row)

    total_raw = len(raw_df)
    _batch = []
    _batch_size = 500
    # 使用 itertuples 避免 to_numpy() 的記憶體峰值
    for i, row in enumerate(raw_df.itertuples(index=False)):
        _batch.append([
            None if pd.isna(v) else (round(float(v), 1) if isinstance(v, (int, float)) and not isinstance(v, bool) else v)
            for v in row
        ])
        if len(_batch) >= _batch_size or i == total_raw - 1:
            for _r in _batch:
                ws.append(_r)
            _batch.clear()
        if progress_callback and (i % 1000 == 0 or i == total_raw - 1):
            p = progress_start + (progress_end - progress_start) * ((i + 1) / total_raw)
            progress_callback(p, f"寫入 RAW ({i+1}/{total_raw})...")

    header_fill_raw = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    ncols = len(raw_cols)
    for col_idx in range(1, ncols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        cell.fill = header_fill_raw
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    ws.auto_filter.ref = ws.dimensions
    _auto_width(ws)


def _write_metric_cards(ws, title, row1_metrics, row2_metrics, thin_border, desc_font, progress_callback=None, progress_pct=0.95, row3_metrics=None):
    ws.merge_cells("A1:D1")
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(bold=True, size=14, color="1E40AF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    row_offsets = [(3, row1_metrics), (7, row2_metrics)]
    if row3_metrics:
        row_offsets.append((11, row3_metrics))

    for row_offset, metrics in row_offsets:
        for i, (name, value, desc, color) in enumerate(metrics):
            col = i + 1
            tc = ws.cell(row=row_offset, column=col, value=name)
            tc.font = Font(bold=True, size=11, color="FFFFFF")
            tc.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            tc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            tc.border = thin_border
            vc = ws.cell(row=row_offset + 1, column=col, value=value)
            vc.font = Font(bold=True, size=20, color=color)
            vc.alignment = Alignment(horizontal="center", vertical="center")
            vc.number_format = "#,##0"
            vc.border = thin_border
            dc = ws.cell(row=row_offset + 2, column=col, value=desc)
            dc.font = desc_font
            dc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            dc.border = thin_border

    if progress_callback:
        progress_callback(progress_pct, "調整格式中...")
    _auto_width(ws)


def generate_excel_combined(
    summary_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    categories: list[str],
    progress_callback=None,
) -> BytesIO:
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    desc_font = Font(size=9, color="6B7280", italic=True)

    # ── Sheet 1: RAW ──
    ws_raw = wb.active
    ws_raw.title = "RAW"

    _write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.15)

    if progress_callback:
        progress_callback(0.2, "建立摘要工作表...")

    # ── Sheet 2: Shop Stock Status ──
    ws_summary = wb.create_sheet(title="Shop Stock Status")

    headers = ["Buyer", "Brand", "SKU", "SKU Desc"] + categories + ["總計", "Safety Qty Change", "增加Safety Qty的店舖數量", "減少Safety Qty的店舖數量", "Safety Qty不變的店舖數量"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    total_summary = len(summary_df)
    for row_idx, row_data in summary_df.iterrows():
        excel_row = 2 + row_idx
        is_total = row_data["Buyer"] == "總計"
        for col_idx, h in enumerate(headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws_summary.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_total:
                cell.font = bold_font
            if isinstance(val, (int, float)) and col_idx >= 4 and h != "SKU":
                cell.number_format = "#,##0"
        if progress_callback and (row_idx % 500 == 0 or row_idx == total_summary - 1):
            progress_callback(0.4 + 0.2 * ((row_idx + 1) / max(total_summary, 1)), f"寫入摘要 ({row_idx+1}/{total_summary})...")

    _auto_width(ws_summary)

    if progress_callback:
        progress_callback(0.65, "建立 By Shop 分析工作表...")

    # ── Sheet 3: By Shop ──
    by_shop_df = _build_combined_by_shop_summary(raw_df)

    ws_by_shop = wb.create_sheet(title="By Shop")
    by_shop_headers = [
        "Site Code",
        "Shop Name",
        "OM",
        "Overstock_SKU数量",
        "OOS_SKU数量",
        "需增加Safety_SKU数量",
        "需减少Safety_SKU数量",
        "安全库存调整总量",
        "Combined_SKU总数",
    ]

    for col_idx, h in enumerate(by_shop_headers, 1):
        cell = ws_by_shop.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, row_data in by_shop_df.iterrows():
        excel_row = row_idx + 2
        is_total = row_data.get("Site Code") == "總計"
        for col_idx, h in enumerate(by_shop_headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws_by_shop.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_total:
                cell.font = bold_font
            if isinstance(val, (int, float)) and col_idx >= 4:
                cell.number_format = "#,##0"
        if progress_callback and (row_idx % 500 == 0 or row_idx == len(by_shop_df) - 1):
            progress_callback(0.7 + 0.15 * ((row_idx + 1) / max(len(by_shop_df), 1)), f"寫入 By Shop ({row_idx+1}/{len(by_shop_df)})...")

    _auto_width(ws_by_shop)

    if progress_callback:
        progress_callback(0.88, "建立分析摘要工作表...")

    # ── Sheet 4: Safety Stock 數據分析摘要 ──
    ws_dashboard = wb.create_sheet(title="Safety Stock 數據分析摘要")

    _total_row = summary_df[summary_df["Buyer"] == "總計"]
    if _total_row.empty:
        raise ValueError("摘要資料中未找到總計列，無法生成分析看板")
    total_combined = int(_total_row["總計"].values[0])
    total_buyers = len(summary_df[summary_df["Buyer"] != "總計"]["Buyer"].dropna().unique())
    total_brands = len(summary_df[summary_df["Buyer"] != "總計"]["Brand"].dropna().unique())
    total_skus = raw_df["SKU"].dropna().nunique()

    cat_overstock = int(_total_row[categories[0]].values[0]) if categories[0] in _total_row.columns else 0
    cat_oos = int(_total_row[categories[1]].values[0]) if categories[1] in _total_row.columns else 0
    cat_balanced = int(_total_row[categories[2]].values[0]) if categories[2] in _total_row.columns else 0

    safety_change_total = int(pd.to_numeric(raw_df.get("Safety Qty Change", _safe_default_series(0, raw_df.index)), errors="coerce").fillna(0).sum())

    row1_metrics = [
        ("總 Safety Stock 審查行數", total_combined, "所有 Safety Stock 審查之總計行數", "EF4444"),
        ("影響採購人員數 (Buyers)", total_buyers, "位採購人員名下有需審查品項", "3B82F6"),
        ("影響品牌總數 (Brands)", total_brands, "個品牌包含需審查商品", "8B5CF6"),
        ("影響SKU總數", total_skus, "個包含需審查商品(SKUs)", "10B981"),
    ]

    row2_metrics = [
        ("Overstock 分類行數", cat_overstock, "庫存過剩分類行數", "F59E0B"),
        ("OOS 分類行數", cat_oos, "缺貨分類行數", "EC4899"),
        ("Safety 調整行數", cat_balanced, "Balanced (Safety Adjusted) 行數", "06B6D4"),
        ("Safety 調整總量", safety_change_total, "Safety Qty Change 總和", "8B5CF6"),
    ]

    _write_metric_cards(ws_dashboard, "Safety Stock 數據分析摘要 (Safety Stock Analysis Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback)

    if progress_callback:
        progress_callback(0.97, "正在儲存 Excel 檔案...")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf


def transform(df_input: pd.DataFrame, progress_callback=None, shop_natures=None, rates=None):
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"\u7f3a\u5c11\u5fc5\u8981\u6b04\u4f4d: {', '.join(missing)}")

    warnings = []

    _optional_cols = ["Adj. Safety Stock", "Shop Stock Qty (CM)", "Shop Monthly Sold Qty (P90D)"]
    for _col in _optional_cols:
        if _col not in df_input.columns:
            warnings.append(f"缺少可選欄位「{_col}」，建議值與轉移數量將全為 0")

    if progress_callback:
        progress_callback(0.05, "\u9a57\u8b49\u901a\u904e")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據（非空白或總計列）")

    df["SKU"] = df["SKU"].apply(_safe_sku)

    if "SKU Desc" not in df.columns:
        df = _compute_raw_extra_columns(df)

    status_col = df["Shop Stock Status"].fillna("")

    categories = [
        "P90D 0 Sales",
        "Overstock (>2*Monthly Sold Qty & >2pcs)",
    ]

    _STATUS_VALUE = {
        "P90D 0 Sales": "Overstock (P90D 0 Sales)",
        "Overstock (>2*Monthly Sold Qty & >2pcs)": "Overstock (>2*Monthly Sold Qty & >2pcs)",
    }

    group_cols = ["Buyer", "Brand", "SKU", "SKU Desc"]

    if progress_callback:
        progress_callback(0.2, "\u5206\u985e\u7d71\u8a08\u4e2d...")

    counts = {}
    for i, cat in enumerate(categories):
        mask = (status_col == _STATUS_VALUE.get(cat, cat))
        counts[cat] = (
            df[mask]
            .groupby(group_cols, observed=True)
            .size()
            .rename(cat)
        )
        if progress_callback:
            progress_callback(0.2 + 0.3 * (i + 1) / len(categories), f"\u7d71\u8a08 {cat}...")

    if progress_callback:
        progress_callback(0.5, "\u8a08\u7b97\u5eab\u5b58\u5efa\u8b70\u4e2d...")

    raw = _add_recommendation_columns(df, rates=rates)

    transfer_col = "After new Safety Qty / ND can transfer out of QTY"
    raw["_transfer_num"] = pd.to_numeric(raw[transfer_col], errors="coerce").fillna(0)
    transfer_summary = (
        raw.groupby(group_cols, observed=True)["_transfer_num"]
        .sum()
        .astype(int)
        .rename(transfer_col)
    )

    if progress_callback:
        progress_callback(0.55, "\u6392\u5e8f\u5f59\u7e3d\u4e2d...")

    summary = pd.concat(list(counts.values()), axis=1).fillna(0).astype(int).reset_index()

    summary["SKU"] = summary["SKU"].apply(_safe_sku)
    summary["\u7e3d\u8a08"] = summary[categories].sum(axis=1)

    summary = summary.merge(transfer_summary, on=group_cols, how="left")
    summary[transfer_col] = summary[transfer_col].fillna(0).astype(int)

    buyer_order = summary.groupby("Buyer")["總計"].sum().sort_values(ascending=False).index.tolist()
    summary["_buyer_key"] = summary["Buyer"].map({b: i for i, b in enumerate(buyer_order)})
    summary = summary.sort_values(["_buyer_key", "Brand", "SKU", "SKU Desc"], ascending=[True, True, True, True]).drop(columns=["_buyer_key"]).reset_index(drop=True)

    transfer_total = int(raw["_transfer_num"].sum())
    del raw["_transfer_num"]

    total_overstock = int(summary["總計"].sum())
    total_row = pd.DataFrame(
        {
            "Buyer": ["總計"],
            "Brand": [None],
            "SKU": [None],
            "SKU Desc": [None],
            **{cat: [int(summary[cat].sum())] for cat in categories},
            "\u7e3d\u8a08": [total_overstock],
            transfer_col: [transfer_total],
        }
    )
    summary = pd.concat([summary, total_row], ignore_index=True)

    if total_overstock == 0:
        warnings.append(
            "\u672a\u5075\u6e2c\u5230\u4efb\u4f55\u7b26\u5408\u689d\u4ef6\u7684\u5eab\u5b58\u7570\u5e38\u8cc7\u6599\u3002\u8acb\u78ba\u8a8d\u4f86\u6e90\u6a94\u6848\u4e2d\u7684\u300cShop Stock Status\u300d\u6b04\u4f4d\u503c\u662f\u5426\u5305\u542b\u4ee5\u4e0b\u5206\u985e\uff1a"
            + ", ".join(categories)
        )

    if progress_callback:
        progress_callback(0.7, "\u8f49\u63db\u5b8c\u6210")

    return summary, raw, categories, warnings


def generate_excel(summary_df: pd.DataFrame, raw_df: pd.DataFrame, categories: list[str], progress_callback=None) -> BytesIO:
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    ws_raw = wb.active
    ws_raw.title = "RAW"

    _write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.3)

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    if progress_callback:
        progress_callback(0.35, "建立摘要工作表...")

    ws_summary = wb.create_sheet(title="Shop Stock Status")

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)

    ws_summary["A3"] = "\u884c\u6578 - Site Code"
    ws_summary["D3"] = "Shop Stock Status"

    transfer_col = "After new Safety Qty / ND can transfer out of QTY"
    headers = ["Buyer", "Brand", "SKU", "SKU Desc"] + categories + ["總計"] + [transfer_col]
    for col_idx, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    data_start = 5
    total_summary = len(summary_df)

    for row_idx, row_data in summary_df.iterrows():
        excel_row = data_start + row_idx
        is_total = row_data["Buyer"] == "\u7e3d\u8a08"

        for col_idx, h in enumerate(headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws_summary.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

            if is_total:
                cell.font = bold_font

            if isinstance(val, (int, float)) and col_idx >= 4 and h != "SKU":
                cell.number_format = "#,##0"

        if progress_callback and (row_idx % 500 == 0 or row_idx == total_summary - 1):
            progress_callback(0.6 + 0.3 * ((row_idx + 1) / total_summary), f"\u5beb\u5165\u6458\u8981 ({row_idx+1}/{total_summary})...")

    _auto_width(ws_summary)

    if progress_callback:
        progress_callback(0.72, "\u5efa\u7acb By Shop \u5206\u6790\u5de5\u4f5c\u8868...")

    by_shop_df = _build_by_shop_summary(raw_df)

    ws_by_shop = wb.create_sheet(title="By Shop")
    by_shop_headers = [
        "Site Code",
        "Shop Name",
        "OM",
        "更改NDRF行數",
        "更改Safety行數",
        "預計總轉出貨量",
        "P90D 0 Sales SKU數量",
        ">2*Monthly Sold Qty & >2pcs數量",
    ]

    for col_idx, h in enumerate(by_shop_headers, 1):
        cell = ws_by_shop.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, row_data in by_shop_df.iterrows():
        excel_row = row_idx + 2
        is_total = row_data.get("Site Code") == "總計"
        for col_idx, h in enumerate(by_shop_headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws_by_shop.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_total:
                cell.font = bold_font
            if isinstance(val, (int, float)) and col_idx >= 4:
                cell.number_format = "#,##0"

        if progress_callback and (row_idx % 500 == 0 or row_idx == len(by_shop_df) - 1):
            progress_callback(0.82 + 0.08 * ((row_idx + 1) / max(len(by_shop_df), 1)), f"\u5beb\u5165 By Shop ({row_idx+1}/{len(by_shop_df)})...")

    _auto_width(ws_by_shop)

    if progress_callback:
        progress_callback(0.92, "\u5efa\u7acb\u5206\u6790\u6458\u8981\u5de5\u4f5c\u8868...")

    ws_dashboard = wb.create_sheet(title="異常庫存數據分析摘要")

    _total_row = summary_df[summary_df["Buyer"] == "總計"]
    if _total_row.empty:
        raise ValueError("摘要資料中未找到總計列，無法生成分析看板")
    total_overstock = int(_total_row["總計"].values[0])
    total_buyers = len(summary_df[summary_df["Buyer"] != "總計"]["Buyer"].dropna().unique())
    total_brands = len(summary_df[summary_df["Buyer"] != "總計"]["Brand"].dropna().unique())
    total_skus = raw_df["SKU"].dropna().nunique()
    cat1_val = int(_total_row[categories[0]].values[0])
    cat2_val = int(_total_row[categories[1]].values[0])

    transfer_col = _resolve_column_name("After new Safety Qty / ND can transfer out of QTY", set(raw_df.columns))
    transfer_qty_total = 0
    if transfer_col:
        transfer_qty_total = int(pd.to_numeric(raw_df[transfer_col], errors="coerce").fillna(0).sum())

    row1_metrics = [
        ("總 Overstock 行數", total_overstock, "所有庫存異常之總計行數", "EF4444"),
        ("影響採購人員數 (Buyers)", total_buyers, "位採購人員名下有異常品項", "3B82F6"),
        ("影響品牌總數 (Brands)", total_brands, "個品牌包含異常庫存商品", "8B5CF6"),
        ("影響SKU總數", total_skus, "個包含異常庫存商品(SKUs)", "10B981"),
    ]

    row2_metrics = [
        ("P90D 0 Sales", cat1_val, "「過去 90 天內零銷售」之異常行數", "F59E0B"),
        ("Overstock (>2*Monthly Sold Qty & >2pcs)", cat2_val, "「大於月均銷售量兩倍且庫存大於 2 件」之異常行數", "EC4899"),
        ("調整後的預計可退倉總數", transfer_qty_total, "After new Safety Qty / ND can transfer out of TOTAL QTY", "06B6D4"),
    ]

    desc_font = Font(size=9, color="6B7280", italic=True)
    _write_metric_cards(ws_dashboard, "異常庫存數據分析摘要 (Analysis Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback)

    if progress_callback:
        progress_callback(0.97, "正在儲存 Excel 檔案...")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "\u5b8c\u6210!")

    return buf


# ─── ND SKU Review (SKU 精簡 Review) Functions ─────────────────────────────────


def _add_nd_recommendation_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Any RP Type == ND is recommended to transfer all stock out (full 退倉),
    # regardless of Shop Stock Status (relaxes Overstock Review which only
    # 退倉 for Overstock remarks). Normalize RP Type (strip + upper) so values
    # like " nd " or "Nd" are recognized, consistent with the Overstock path.
    rp_type = (
        df.get("RP Type", _safe_default_series("", df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    shop_stock = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, df.index)))

    # Only Warehouse stock can be returned to warehouse (退倉); Vendor stock must
    # stay at the shop (另有安排) and contributes 0 to transfer quantity.
    supply_source = (
        df.get("Supply Source", _safe_default_series("", df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )

    is_nd = rp_type == "ND"
    is_warehouse = is_nd & (supply_source == "warehouse")
    is_vendor = is_nd & (supply_source == "vendor")

    rec_values = np.full(len(df), "", dtype=object)
    rec_values[is_warehouse.values] = "ND"
    # Vendor stock cannot be returned — mark it so buyers see why it is not 退倉.
    rec_values[is_vendor.values] = "另有安排"

    transfer_values = np.full(len(df), "", dtype=object)
    # Warehouse: transfer all shop stock; Vendor or zero stock: transfer 0.
    transfer_values[is_warehouse.values] = shop_stock[is_warehouse.values].astype(int).values
    transfer_values[is_vendor.values] = 0

    df = df.copy()
    df["Recommendation of new Safety Qty / ND"] = rec_values
    df["After new Safety Qty / ND can transfer out of QTY"] = transfer_values
    df["Buyer reply"] = ""

    return df


def _build_nd_by_shop_summary(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()

    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    shop_name = [
        _extract_shop_name(item_shop.iloc[i], site_code.iloc[i])
        for i in range(len(df))
    ]

    df["Site Code"] = site_code
    df["Shop Name"] = shop_name
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)

    rp_type = (
        df.get("RP Type", _safe_default_series("", df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    supply_source = (
        df.get("Supply Source", _safe_default_series("", df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )
    transfer_col = "After new Safety Qty / ND can transfer out of QTY"
    transfer_num = pd.to_numeric(df.get(transfer_col, _safe_default_series(0, df.index)), errors="coerce").fillna(0)

    group_cols = ["Site Code", "Shop Name", "OM"]
    df["_is_nd"] = rp_type == "ND"
    df["_is_nd_warehouse"] = df["_is_nd"] & (supply_source == "warehouse")
    df["_is_nd_vendor"] = df["_is_nd"] & (supply_source == "vendor")
    df["_sku_nd"] = df["SKU"].where(df["_is_nd"])
    df["_sku_nd_warehouse"] = df["SKU"].where(df["_is_nd_warehouse"])
    df["_transfer_num"] = transfer_num

    by_shop = (
        df.groupby(group_cols, dropna=False, observed=True)
        .agg(
            ND行數=("_is_nd", "sum"),
            Warehouse退倉行數=("_is_nd_warehouse", "sum"),
            Vendor另有安排行數=("_is_nd_vendor", "sum"),
            預計總退倉量=("_transfer_num", "sum"),
            ND_SKU數量=("_sku_nd", "nunique"),
            Warehouse_SKU數量=("_sku_nd_warehouse", "nunique"),
        )
        .reset_index()
    )

    for c in ["ND行數", "Warehouse退倉行數", "Vendor另有安排行數", "預計總退倉量", "ND_SKU數量", "Warehouse_SKU數量"]:
        by_shop[c] = by_shop[c].fillna(0).astype(int)

    if not by_shop.empty:
        total_row = pd.DataFrame(
            {
                "Site Code": ["總計"],
                "Shop Name": [""],
                "OM": [""],
                "ND行數": [int(by_shop["ND行數"].sum())],
                "Warehouse退倉行數": [int(by_shop["Warehouse退倉行數"].sum())],
                "Vendor另有安排行數": [int(by_shop["Vendor另有安排行數"].sum())],
                "預計總退倉量": [int(by_shop["預計總退倉量"].sum())],
                "ND_SKU數量": [int(by_shop["ND_SKU數量"].sum())],
                "Warehouse_SKU數量": [int(by_shop["Warehouse_SKU數量"].sum())],
            }
        )
        by_shop = pd.concat([by_shop.sort_values(["Site Code", "Shop Name", "OM"], kind="stable"), total_row], ignore_index=True)

    return by_shop


def transform_nd(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str], list[str]]:
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    warnings = []

    _optional_cols = ["RP Type", "Shop Stock Qty (CM)", "Shop Stock Status", "Supply Source"]
    for _col in _optional_cols:
        if _col not in df_input.columns:
            warnings.append(f"缺少可選欄位「{_col}」，退倉建議將可能為 0")

    if progress_callback:
        progress_callback(0.05, "驗證通過")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據（非空白或總計列）")

    df["SKU"] = df["SKU"].apply(_safe_sku)

    if "SKU Desc" not in df.columns:
        df = _compute_raw_extra_columns(df)

    if "RP Type" not in df.columns:
        raise ValueError("來源檔案缺少「RP Type」欄位，SKU 精簡 Review 必須依賴此欄位辨識 ND SKU")

    if progress_callback:
        progress_callback(0.2, "計算退倉建議中...")

    raw = _add_nd_recommendation_columns(df)

    # Two-layer dynamic categories: every ND row is classified by its Supply Source
    # (Warehouse / Vendor) crossed with its actual Shop Stock Status, so the summary
    # breaks down ND SKUs by both where the stock comes from and its real status.
    rp_type = (
        raw.get("RP Type", _safe_default_series("", raw.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    supply_source = (
        raw.get("Supply Source", _safe_default_series("", raw.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )
    status_col = raw["Shop Stock Status"].fillna("")

    nd_mask = rp_type == "ND"
    # Build composite category label "SupplySource - Status" (title-cased source).
    _source_title = supply_source.where(supply_source == "", supply_source.str.title())
    nd_source_status = np.where(
        nd_mask,
        np.where(
            supply_source == "",
            "Unknown - " + status_col.values,
            _source_title + " - " + status_col.values,
        ),
        "",
    )
    # Unique non-empty categories among ND rows.
    nd_cats = sorted({c for c in nd_source_status if c})
    categories = nd_cats

    if progress_callback:
        progress_callback(0.4, "分類統計中...")

    group_cols = ["Buyer", "Brand", "SKU", "SKU Desc"]

    counts = {}
    _nd_cat_series = pd.Series(nd_source_status, index=raw.index)
    for cat in categories:
        mask = nd_mask & (_nd_cat_series == cat)
        counts[cat] = (
            raw[mask]
            .groupby(group_cols, observed=True)
            .size()
            .rename(cat)
        )

    transfer_col = "After new Safety Qty / ND can transfer out of QTY"
    raw["_transfer_num"] = pd.to_numeric(raw[transfer_col], errors="coerce").fillna(0)
    transfer_summary = (
        raw.groupby(group_cols, observed=True)["_transfer_num"]
        .sum()
        .astype(int)
        .rename(transfer_col)
    )

    if progress_callback:
        progress_callback(0.55, "彙總排序中...")

    if counts:
        summary = pd.concat(list(counts.values()), axis=1).fillna(0).astype(int).reset_index()
    else:
        summary = pd.DataFrame(columns=group_cols)

    summary["SKU"] = summary["SKU"].apply(_safe_sku)
    summary["總計"] = summary[categories].sum(axis=1) if categories else 0

    summary = summary.merge(transfer_summary, on=group_cols, how="left")
    summary[transfer_col] = summary[transfer_col].fillna(0).astype(int)

    buyer_order = summary.groupby("Buyer")["總計"].sum().sort_values(ascending=False).index.tolist()
    summary["_buyer_key"] = summary["Buyer"].map({b: i for i, b in enumerate(buyer_order)})
    summary = summary.sort_values(["_buyer_key", "Brand", "SKU", "SKU Desc"], ascending=[True, True, True, True]).drop(columns=["_buyer_key"]).reset_index(drop=True)

    transfer_total = int(raw["_transfer_num"].sum())
    del raw["_transfer_num"]

    total_nd = int(summary["總計"].sum())
    total_row = pd.DataFrame(
        {
            "Buyer": ["總計"],
            "Brand": [None],
            "SKU": [None],
            "SKU Desc": [None],
            **{cat: [int(summary[cat].sum())] for cat in categories},
            "總計": [total_nd],
            transfer_col: [transfer_total],
        }
    )
    summary = pd.concat([summary, total_row], ignore_index=True)

    if total_nd == 0:
        warnings.append("未偵測到任何 RP Type 為 ND 的 SKU。請確認來源檔案中的「RP Type」欄位是否包含 ND 值。")

    if progress_callback:
        progress_callback(0.7, "轉換完成")

    return summary, raw, categories, warnings


def generate_excel_nd(
    summary_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    categories: list[str],
    progress_callback=None,
) -> BytesIO:
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    ws_raw = wb.active
    ws_raw.title = "RAW"

    _write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.3)

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    if progress_callback:
        progress_callback(0.35, "建立摘要工作表...")

    ws_summary = wb.create_sheet(title="Shop Stock Status")

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)

    ws_summary["A3"] = "行數 - Site Code"
    ws_summary["D3"] = "ND SKU (by Shop Stock Status)"

    transfer_col = "After new Safety Qty / ND can transfer out of QTY"
    headers = ["Buyer", "Brand", "SKU", "SKU Desc"] + categories + ["總計"] + [transfer_col]
    for col_idx, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    data_start = 5
    total_summary = len(summary_df)

    for row_idx, row_data in summary_df.iterrows():
        excel_row = data_start + row_idx
        is_total = row_data["Buyer"] == "總計"

        for col_idx, h in enumerate(headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws_summary.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

            if is_total:
                cell.font = bold_font

            if isinstance(val, (int, float)) and col_idx >= 4 and h != "SKU":
                cell.number_format = "#,##0"

        if progress_callback and (row_idx % 500 == 0 or row_idx == total_summary - 1):
            progress_callback(0.6 + 0.3 * ((row_idx + 1) / total_summary), f"寫入摘要 ({row_idx+1}/{total_summary})...")

    _auto_width(ws_summary)

    if progress_callback:
        progress_callback(0.72, "建立 By Shop 分析工作表...")

    by_shop_df = _build_nd_by_shop_summary(raw_df)

    ws_by_shop = wb.create_sheet(title="By Shop")
    by_shop_headers = [
        "Site Code",
        "Shop Name",
        "OM",
        "ND行數",
        "Warehouse退倉行數",
        "Vendor另有安排行數",
        "預計總退倉量",
        "ND_SKU數量",
        "Warehouse_SKU數量",
    ]

    for col_idx, h in enumerate(by_shop_headers, 1):
        cell = ws_by_shop.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, row_data in by_shop_df.iterrows():
        excel_row = row_idx + 2
        is_total = row_data.get("Site Code") == "總計"
        for col_idx, h in enumerate(by_shop_headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws_by_shop.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_total:
                cell.font = bold_font
            if isinstance(val, (int, float)) and col_idx >= 4:
                cell.number_format = "#,##0"

        if progress_callback and (row_idx % 500 == 0 or row_idx == len(by_shop_df) - 1):
            progress_callback(0.82 + 0.08 * ((row_idx + 1) / max(len(by_shop_df), 1)), f"寫入 By Shop ({row_idx+1}/{len(by_shop_df)})...")

    _auto_width(ws_by_shop)

    if progress_callback:
        progress_callback(0.92, "建立分析摘要工作表...")

    ws_dashboard = wb.create_sheet(title="ND SKU數據分析摘要")

    _total_row = summary_df[summary_df["Buyer"] == "總計"]
    if _total_row.empty:
        raise ValueError("摘要資料中未找到總計列，無法生成分析看板")
    total_nd = int(_total_row["總計"].values[0])

    # Warehouse (可退倉) vs Vendor (另有安排) breakdown from the raw data.
    _rp = raw_df.get("RP Type", _safe_default_series("", raw_df.index)).fillna("").astype(str).str.strip().str.upper()
    _ss = raw_df.get("Supply Source", _safe_default_series("", raw_df.index)).fillna("").astype(str).str.strip().str.lower()
    _nd_mask = (_rp == "ND")
    _nd_wh_mask = _nd_mask & (_ss == "warehouse")
    _nd_vendor_mask = _nd_mask & (_ss == "vendor")
    _nd_wh = int(_nd_wh_mask.sum())
    _nd_vendor = int(_nd_vendor_mask.sum())

    # Buyers / Brands / SKU totals must count ONLY affected ND rows (raw_df keeps all
    # cleaned rows incl. non-ND, so counting over raw_df would over-count).
    _nd_df = raw_df[_nd_mask]
    total_buyers = _nd_df["Buyer"].dropna().nunique() if len(_nd_df) else 0
    total_brands = _nd_df["Brand"].dropna().nunique() if len(_nd_df) else 0
    total_skus = _nd_df["SKU"].dropna().nunique() if len(_nd_df) else 0

    transfer_col_resolved = _resolve_column_name("After new Safety Qty / ND can transfer out of QTY", set(raw_df.columns))
    transfer_qty_total = 0
    if transfer_col_resolved:
        transfer_qty_total = int(pd.to_numeric(raw_df[transfer_col_resolved], errors="coerce").fillna(0).sum())

    # 另有安排 total qty = sum of Shop Stock Qty over ND+Vendor rows (cannot be returned).
    _shop_stock_col = _resolve_column_name("Shop Stock Qty (CM)", set(raw_df.columns))
    other_arranged_qty_total = 0
    if _shop_stock_col:
        other_arranged_qty_total = int(_safe_numeric(raw_df.loc[_nd_vendor_mask, _shop_stock_col]).sum())

    row1_metrics = [
        ("總 ND SKU 行數", total_nd, "所有 ND SKU 總計行數", "EF4444"),
        ("Warehouse 可退倉行數", _nd_wh, "Supply Source=Warehouse 可退倉之 ND 行數", "10B981"),
        ("Vendor 另有安排行數", _nd_vendor, "Supply Source=Vendor 不可退倉、另有安排之 ND 行數", "F59E0B"),
        ("影響採購人員數 (Buyers)", total_buyers, "位採購人員名下有 ND SKU", "3B82F6"),
    ]

    row2_metrics = [
        ("影響品牌總數 (Brands)", total_brands, "個品牌包含 ND SKU", "8B5CF6"),
        ("影響SKU總數", total_skus, "個包含 ND SKU(SKUs)", "10B981"),
        ("預計可退倉總數量", transfer_qty_total, "Warehouse 可退倉數量總和", "06B6D4"),
        ("預計另有安排總數量", other_arranged_qty_total, "Vendor 另有安排之庫存數量總和", "F97316"),
    ]

    desc_font = Font(size=9, color="6B7280", italic=True)
    _write_metric_cards(ws_dashboard, "ND SKU 數據分析摘要 (ND SKU Analysis Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback)

    if progress_callback:
        progress_callback(0.97, "正在儲存 Excel 檔案...")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf


def _auto_width(ws, min_width=8, max_width=50, sample_rows=200):
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0.0
        scan_end = min(ws.max_row, 1 + sample_rows)
        for row_idx in range(1, scan_end + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                s = str(cell.value)
                val_len = sum(2.0 if unicodedata.east_asian_width(c) in ("W", "F") else 1.0 for c in s)
                if val_len > max_len:
                    max_len = val_len
        adjusted = min(max(max_len + 2, min_width), max_width)
        ws.column_dimensions[col_letter].width = adjusted


# ─── OOS Review Functions ──────────────────────────────────────────────────────

OOS_CATEGORIES = [
    "OOS (0 Stock)",
    "Potential OOS (<7 Turnday)",
]

SECONDARY_CATEGORY = "Normal (Safety Qty Insufficient)"

EXCLUDED_STATUSES = ["No Stock (non-Active SKU / ND Shop)"]


def _add_oos_recommendation_columns(df, rates, mtd_days, month_total_days, safety_mode="allow_both", oos_calc_mode="mtd"):
    df = df.copy()

    adj_ss = _safe_numeric(df.get("Adj. Safety Stock", _safe_default_series(0.0, df.index)))
    shop_nature = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.lower()
    rp_type = df.get("RP Type", _safe_default_series("", df.index)).fillna("").astype(str).str.strip().str.upper()

    sold_qty = _safe_numeric(df.get("Sold Qty (CM - MTD)", _safe_default_series(0.0, df.index)))
    if mtd_days and mtd_days > 0:
        projected_mtd = (sold_qty / mtd_days) * month_total_days
    else:
        projected_mtd = _safe_default_series(0.0, df.index)
    df["Projected Monthly Sold Qty"] = np.ceil(projected_mtd).astype(int)
    projected_mtd_arr = projected_mtd.values

    if oos_calc_mode == "p90d":
        p90d_monthly = _safe_numeric(df.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0.0, df.index)))
        projected = p90d_monthly.values
    else:
        projected = projected_mtd_arr

    _default_rate = rates.get("default", 0.50)
    _tourist_rate = rates.get("tourist", 0.30)
    _mix_rate = rates.get("mix", 0.40)

    rate_arr = np.full(len(df), _default_rate)
    rate_arr[shop_nature.str.contains("tourist", na=False)] = _tourist_rate
    rate_arr[shop_nature.str.contains("mix", na=False)] = _mix_rate

    is_rf = (rp_type == "RF")
    is_nd = (rp_type == "ND")
    is_other = ~is_rf & ~is_nd

    adj_ss_arr = adj_ss.values.astype(int)

    recommended = np.zeros(len(df), dtype=int)

    if is_rf.any():
        rec_rf = np.ceil(projected[is_rf] * rate_arr[is_rf]).clip(min=3).astype(int)
        recommended[is_rf] = rec_rf

    if is_other.any():
        rec_other = np.ceil(projected[is_other] * rate_arr[is_other]).clip(min=3).astype(int)
        ad_other = adj_ss_arr[is_other]
        over_adj = rec_other > ad_other
        rec_other[over_adj] = ad_other[over_adj]
        recommended[is_other] = rec_other

    df["Recommendation of new Safety Qty"] = recommended

    increase = recommended - adj_ss_arr
    increase[is_nd] = 0
    if safety_mode == "increase_only":
        increase[is_rf] = increase[is_rf].clip(min=0)
        increase[is_other] = increase[is_other].clip(min=0)
    else:
        increase[is_other] = increase[is_other].clip(min=0)
    df["Safety Qty Increase"] = increase

    # Keep Recommendation consistent with the (possibly clipped) increase so the two
    # columns never disagree. increase = recommended - adj_ss by construction, so for
    # rows whose increase was not clipped this is a no-op (rec_sync == recommended);
    # for rows that were frozen (increase clipped to 0, e.g. 'other' rows), this
    # surfaces adj_ss instead of the lowered raw recommendation.
    rec_sync = (adj_ss_arr + increase).astype(int)
    df["Recommendation of new Safety Qty"] = rec_sync

    df["Buyer reply"] = ""

    return df


def _build_oos_by_shop_summary(raw):
    df = raw.copy()

    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    shop_name = [
        _extract_shop_name(item_shop.iloc[i], site_code.iloc[i])
        for i in range(len(df))
    ]

    df["Site Code"] = site_code
    df["Shop Name"] = shop_name
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)

    status_col = df["Shop Stock Status"].fillna("")
    safety_increase = pd.to_numeric(df.get("Safety Qty Increase", _safe_default_series(0, df.index)), errors="coerce").fillna(0)

    group_cols = ["Site Code", "Shop Name", "OM"]
    df["_is_oos"] = status_col == "OOS (0 Stock)"
    df["_is_potential_oos"] = status_col == "Potential OOS (<7 Turnday)"
    df["_is_normal_insufficient"] = (status_col == "Normal") & (safety_increase > 0)
    df["_sku_oos"] = df["SKU"].where(df["_is_oos"])
    df["_sku_potential_oos"] = df["SKU"].where(df["_is_potential_oos"])
    df["_sku_normal_insufficient"] = df["SKU"].where(df["_is_normal_insufficient"])
    df["_safety_increase"] = safety_increase

    by_shop = (
        df.groupby(group_cols, dropna=False, observed=True)
        .agg(
            OOS_SKU数量=("_sku_oos", "nunique"),
            Potential_OOS_SKU数量=("_sku_potential_oos", "nunique"),
            建議SafetyQty增加行數=("_is_normal_insufficient", "sum"),
            其他需調整SKU數量=("_sku_normal_insufficient", "nunique"),
            建議SafetyQty增加總量=("_safety_increase", "sum"),
        )
        .reset_index()
    )

    for col in by_shop.columns:
        if col not in group_cols:
            by_shop[col] = by_shop[col].fillna(0).astype(int)

    if not by_shop.empty:
        total_row = pd.DataFrame(
            {
                "Site Code": ["總計"],
                "Shop Name": [""],
                "OM": [""],
                "OOS_SKU数量": [int(by_shop["OOS_SKU数量"].sum())],
                "Potential_OOS_SKU数量": [int(by_shop["Potential_OOS_SKU数量"].sum())],
                "建議SafetyQty增加行數": [int(by_shop["建議SafetyQty增加行數"].sum())],
                "其他需調整SKU數量": [int(by_shop["其他需調整SKU數量"].sum())],
                "建議SafetyQty增加總量": [int(by_shop["建議SafetyQty增加總量"].sum())],
            }
        )
        by_shop = pd.concat([by_shop.sort_values(["Site Code", "Shop Name", "OM"], kind="stable"), total_row], ignore_index=True)

    return by_shop


def transform_oos(df_input, progress_callback=None, shop_natures=None, rates=None, mtd_days=None, month_total_days=None, safety_mode="allow_both", oos_calc_mode="mtd"):
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    warnings = []

    _optional_cols = ["Adj. Safety Stock", "Shop Stock Qty (CM)", "Sold Qty (CM - MTD)", "Shop Stock Status"]
    if oos_calc_mode == "p90d":
        _optional_cols.append("Shop Monthly Sold Qty (P90D)")
    for _col in _optional_cols:
        if _col not in df_input.columns:
            warnings.append(f"缺少可選欄位「{_col}」，建議值與轉移數量將全為 0")

    if progress_callback:
        progress_callback(0.05, "驗證通過")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據（非空白或總計列）")

    status_col = df["Shop Stock Status"].fillna("")
    exclude_mask = status_col.isin(EXCLUDED_STATUSES) | status_col.str.startswith("Overstock", na=False)
    df = df[~exclude_mask]

    if df.empty:
        raise ValueError("過濾 OOS / No Stock / Overstock 後無有效資料列")

    if progress_callback:
        progress_callback(0.2, "計算庫存建議中...")

    df["SKU"] = df["SKU"].apply(_safe_sku)
    raw = _add_oos_recommendation_columns(df, rates, mtd_days, month_total_days, safety_mode=safety_mode, oos_calc_mode=oos_calc_mode)

    if progress_callback:
        progress_callback(0.4, "分類統計中...")

    categories = OOS_CATEGORIES + [SECONDARY_CATEGORY]

    status_col = raw["Shop Stock Status"].fillna("")
    safety_increase = pd.to_numeric(raw.get("Safety Qty Increase", _safe_default_series(0, raw.index)), errors="coerce").fillna(0)

    group_cols = ["Buyer", "Brand", "SKU"]

    counts = {}
    for cat in OOS_CATEGORIES:
        mask = (status_col == cat)
        counts[cat] = (
            raw[mask]
            .groupby(group_cols, observed=True)
            .size()
            .rename(cat)
        )

    normal_mask = (status_col == "Normal") & (safety_increase > 0)
    counts[SECONDARY_CATEGORY] = (
        raw[normal_mask]
        .groupby(group_cols, observed=True)
        .size()
        .rename(SECONDARY_CATEGORY)
    )

    summary = pd.concat(list(counts.values()), axis=1).fillna(0).astype(int).reset_index()
    summary["SKU"] = summary["SKU"].apply(_safe_sku)
    summary["總計"] = summary[categories].sum(axis=1)

    safety_inc_per_group = (
        raw.groupby(group_cols, observed=True)["Safety Qty Increase"]
        .sum()
        .astype(int)
    )
    summary = summary.merge(safety_inc_per_group, on=group_cols, how="left")
    summary["Safety Qty Increase"] = summary["Safety Qty Increase"].fillna(0).astype(int)

    buyer_order = summary.groupby("Buyer")["總計"].sum().sort_values(ascending=False).index.tolist()
    summary["_buyer_key"] = summary["Buyer"].map({b: i for i, b in enumerate(buyer_order)})
    summary = summary.sort_values(["_buyer_key", "Brand", "SKU"], ascending=[True, True, True]).drop(columns=["_buyer_key"]).reset_index(drop=True)

    total_oos = int(summary["總計"].sum())
    total_row = pd.DataFrame(
        {
            "Buyer": ["總計"],
            "Brand": [None],
            "SKU": [None],
            **{cat: [int(summary[cat].sum())] for cat in categories},
            "總計": [total_oos],
            "Safety Qty Increase": [int(summary["Safety Qty Increase"].sum())],
        }
    )
    summary = pd.concat([summary, total_row], ignore_index=True)

    if total_oos == 0:
        warnings.append("未偵測到任何符合條件的 OOS 資料。請確認來源檔案中的「Shop Stock Status」欄位值是否包含以下分類: " + ", ".join(categories))

    if progress_callback:
        progress_callback(0.7, "轉換完成")

    return summary, raw, categories, warnings


def generate_excel_oos(summary_df, raw_df, categories, progress_callback=None):
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    ws_raw = wb.active
    ws_raw.title = "RAW"

    _write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.3)

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    if progress_callback:
        progress_callback(0.35, "建立摘要工作表...")

    ws_summary = wb.create_sheet(title="Shop Stock Status")

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    desc_font = Font(size=9, color="6B7280", italic=True)

    ws_summary["A3"] = "行數 - Site Code"
    ws_summary["D3"] = "Shop Stock Status"

    headers = ["Buyer", "Brand", "SKU"] + categories + ["總計", "Safety Qty Increase"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    data_start = 5
    total_summary = len(summary_df)

    for row_idx, row_data in summary_df.iterrows():
        excel_row = data_start + row_idx
        is_total = row_data["Buyer"] == "總計"

        for col_idx, h in enumerate(headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws_summary.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

            if is_total:
                cell.font = bold_font

            if isinstance(val, (int, float)) and col_idx >= 4 and h != "SKU":
                cell.number_format = "#,##0"

        if progress_callback and (row_idx % 500 == 0 or row_idx == total_summary - 1):
            progress_callback(0.6 + 0.3 * ((row_idx + 1) / total_summary), f"寫入摘要 ({row_idx+1}/{total_summary})...")

    _auto_width(ws_summary)

    if progress_callback:
        progress_callback(0.72, "建立 By Shop 分析工作表...")

    by_shop_df = _build_oos_by_shop_summary(raw_df)

    ws_by_shop = wb.create_sheet(title="By Shop")
    by_shop_headers = [
        "Site Code",
        "Shop Name",
        "OM",
        "OOS_SKU数量",
        "Potential_OOS_SKU数量",
        "建議SafetyQty增加行數",
        "其他需調整SKU數量",
        "建議SafetyQty增加總量",
    ]

    for col_idx, h in enumerate(by_shop_headers, 1):
        cell = ws_by_shop.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, row_data in by_shop_df.iterrows():
        excel_row = row_idx + 2
        is_total = row_data.get("Site Code") == "總計"
        for col_idx, h in enumerate(by_shop_headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws_by_shop.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_total:
                cell.font = bold_font
            if isinstance(val, (int, float)) and col_idx >= 4:
                cell.number_format = "#,##0"

        if progress_callback and (row_idx % 500 == 0 or row_idx == len(by_shop_df) - 1):
            progress_callback(0.82 + 0.08 * ((row_idx + 1) / max(len(by_shop_df), 1)), f"寫入 By Shop ({row_idx+1}/{len(by_shop_df)})...")

    _auto_width(ws_by_shop)

    if progress_callback:
        progress_callback(0.92, "建立分析摘要工作表...")

    ws_dashboard = wb.create_sheet(title="OOS庫存數據分析摘要")

    _total_row = summary_df[summary_df["Buyer"] == "總計"]
    if _total_row.empty:
        raise ValueError("摘要資料中未找到總計列，無法生成分析看板")
    total_oos = int(_total_row["總計"].values[0])
    total_buyers = len(summary_df[summary_df["Buyer"] != "總計"]["Buyer"].dropna().unique())
    total_brands = len(summary_df[summary_df["Buyer"] != "總計"]["Brand"].dropna().unique())
    total_skus = raw_df["SKU"].dropna().nunique()

    cat1_val = int(_total_row[categories[0]].values[0]) if len(categories) > 0 else 0
    cat2_val = int(_total_row[categories[1]].values[0]) if len(categories) > 1 else 0
    cat3_val = int(_total_row[categories[2]].values[0]) if len(categories) > 2 else 0

    safety_increase_total = int(pd.to_numeric(raw_df.get("Safety Qty Increase", _safe_default_series(0, raw_df.index)), errors="coerce").fillna(0).sum())

    desc_font = Font(size=9, color="6B7280", italic=True)

    row1_metrics = [
        ("總 OOS 行數", total_oos, "所有 OOS 異常之總計行數", "EF4444"),
        ("影響採購人員數 (Buyers)", total_buyers, "位採購人員名下有 OOS 品項", "3B82F6"),
        ("影響品牌總數 (Brands)", total_brands, "個品牌包含 OOS 商品", "8B5CF6"),
        ("影響SKU總數", total_skus, "個包含 OOS 商品(SKUs)", "10B981"),
    ]

    row2_metrics = [
        ("OOS (0 Stock)", cat1_val, "現有庫存為 0 之 OOS 行數", "F59E0B"),
        ("Potential OOS (<7 Turnday)", cat2_val, "Turnday 不足 7 之潛在 OOS 行數", "EC4899"),
        ("Normal (Safety Qty Insufficient)", cat3_val, "正常庫存但安全庫存不足行數", "06B6D4"),
        ("建議 Safety Qty 修改", safety_increase_total, "建議 Safety Qty 修改 之總量", "8B5CF6"),
    ]

    _write_metric_cards(ws_dashboard, "OOS庫存數據分析摘要 (OOS Analysis Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback)

    if progress_callback:
        progress_callback(0.97, "正在儲存 Excel 檔案...")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf


# ─── Carry SKU Review Functions ─────────────────────────────────────────────────

CARRY_P90D0_STATUS = "Overstock (P90D 0 Sales)"


def _build_carry_pivot(df: pd.DataFrame, category_col: str) -> pd.DataFrame:
    df = df.copy()

    site_to_name = {}
    if "Item_Shop" in df.columns:
        site_code = df["Site Code"].fillna("").astype(str).str.strip()
        item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
        for i in range(len(df)):
            sc = site_code.iloc[i]
            if sc and sc not in site_to_name:
                name = _extract_shop_name(item_shop.iloc[i], site_code.iloc[i])
                if name:
                    site_to_name[sc] = name
        for sc in site_code.unique():
            if sc and sc not in site_to_name:
                site_to_name[sc] = ""

    status = df["Shop Stock Status"].fillna("")
    df["_status2"] = np.where(status == CARRY_P90D0_STATUS, "P90D 0 Sales", "Other Status")

    rp = df.get("RP Type", _safe_default_series("", df.index)).fillna("").astype(str).str.strip().str.upper()
    df["_rp"] = rp

    df = df[df["_rp"].isin(["ND", "RF"])].copy()

    if df.empty:
        return pd.DataFrame(columns=["Site Code", "Shop Name", category_col, "ND_P90D0", "ND_Other", "ND_Total", "RF_P90D0", "RF_Other", "RF_Total", "Grand_Total"])

    group_cols = ["Site Code", category_col, "_rp", "_status2"]
    pivot = df.groupby(group_cols, dropna=False, observed=True).size().reset_index(name="count")

    pivot_wide = pivot.pivot_table(
        index=["Site Code", category_col],
        columns=["_rp", "_status2"],
        values="count",
        fill_value=0,
    )
    pivot_wide.columns = [f"{r}_{s}" for r, s in pivot_wide.columns]
    pivot_wide = pivot_wide.reset_index()

    for r in ["ND", "RF"]:
        for s in ["P90D 0 Sales", "Other Status"]:
            col = f"{r}_{s}"
            if col not in pivot_wide.columns:
                pivot_wide[col] = 0

    pivot_wide = pivot_wide.rename(columns={
        "ND_P90D 0 Sales": "ND_P90D0",
        "ND_Other Status": "ND_Other",
        "RF_P90D 0 Sales": "RF_P90D0",
        "RF_Other Status": "RF_Other",
    })

    pivot_wide["ND_Total"] = (pivot_wide["ND_P90D0"] + pivot_wide["ND_Other"]).astype(int)
    pivot_wide["RF_Total"] = (pivot_wide["RF_P90D0"] + pivot_wide["RF_Other"]).astype(int)
    pivot_wide["Grand_Total"] = (pivot_wide["ND_Total"] + pivot_wide["RF_Total"]).astype(int)
    for c in ["ND_P90D0", "ND_Other", "RF_P90D0", "RF_Other"]:
        pivot_wide[c] = pivot_wide[c].fillna(0).astype(int)

    pivot_wide.insert(1, "Shop Name", pivot_wide["Site Code"].map(site_to_name).fillna(""))

    pivot_wide = pivot_wide.sort_values(["Site Code", category_col])

    result_rows = []
    for site, group in pivot_wide.groupby("Site Code", sort=False):
        result_rows.append(group)
        subtotal = pd.DataFrame({
            "Site Code": [f"{site} 合計"],
            "Shop Name": [""],
            category_col: [None],
            "ND_P90D0": [int(group["ND_P90D0"].sum())],
            "ND_Other": [int(group["ND_Other"].sum())],
            "ND_Total": [int(group["ND_Total"].sum())],
            "RF_P90D0": [int(group["RF_P90D0"].sum())],
            "RF_Other": [int(group["RF_Other"].sum())],
            "RF_Total": [int(group["RF_Total"].sum())],
            "Grand_Total": [int(group["Grand_Total"].sum())],
        })
        result_rows.append(subtotal)

    result = pd.concat(result_rows, ignore_index=True)
    return result


def _build_carry_dashboard(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    site_to_name = {}
    if "Item_Shop" in df.columns:
        site_code = df["Site Code"].fillna("").astype(str).str.strip()
        item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
        for i in range(len(df)):
            sc = site_code.iloc[i]
            if sc and sc not in site_to_name:
                name = _extract_shop_name(item_shop.iloc[i], site_code.iloc[i])
                if name:
                    site_to_name[sc] = name
        for sc in site_code.unique():
            if sc and sc not in site_to_name:
                site_to_name[sc] = ""

    site_code = df["Site Code"].fillna("").astype(str).str.strip()
    rp = df.get("RP Type", _safe_default_series("", df.index)).fillna("").astype(str).str.strip().str.upper()
    status = df["Shop Stock Status"].fillna("")
    stock_value = _safe_numeric(df.get("Shop Stock Value (CM)", _safe_default_series(0, df.index)))
    shop_stock_qty = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, df.index)))
    suggested_rp = df.get("建議新RP Type", _safe_default_series("", df.index)).fillna("").astype(str).str.strip().str.upper()

    # Keep the dashboard scope identical to _build_carry_pivot (ND + RF only); rows with
    # any other RP Type are not Carry SKUs and would otherwise inflate the SKU total so it
    # no longer matches the pivot Grand_Total.
    carry_mask = rp.isin(["ND", "RF"])
    site_code = site_code[carry_mask]
    rp = rp[carry_mask]
    status = status[carry_mask]
    stock_value = stock_value[carry_mask]
    shop_stock_qty = shop_stock_qty[carry_mask]
    suggested_rp = suggested_rp[carry_mask]
    df = df[carry_mask].copy()

    df["_site"] = site_code
    df["_is_rf"] = (rp == "RF").astype(int)
    df["_is_nd"] = (rp == "ND").astype(int)
    df["_is_p90d0"] = (status == CARRY_P90D0_STATUS).astype(int)
    df["_stock_value"] = stock_value
    df["_p90d0_stock_value"] = df["_stock_value"].where(df["_is_p90d0"].astype(bool), 0.0)
    df["_p90d0_shop_stock_qty"] = shop_stock_qty.where(df["_is_p90d0"].astype(bool), 0)
    df["_suggested_nd"] = (suggested_rp == "ND").astype(int)
    df["_suggested_nd_stock_value"] = df["_stock_value"].where(df["_suggested_nd"].astype(bool), 0.0)
    df["_suggested_nd_stock_qty"] = shop_stock_qty.where(df["_suggested_nd"].astype(bool), 0.0)

    agg = df.groupby("_site", dropna=False, observed=True).agg(
        SKU總數=("_site", "count"),
        RF_SKU總數=("_is_rf", "sum"),
        ND_SKU總數=("_is_nd", "sum"),
        P90D0_Sales_SKU總數=("_is_p90d0", "sum"),
        SKU_Stock_Value_CM=("_stock_value", "sum"),
        P90D0_Sales_Stock_Value_CM=("_p90d0_stock_value", "sum"),
        P90D0_Sales_Shop_Stock_Qty_CM=("_p90d0_shop_stock_qty", "sum"),
        建議ND_SKU總數=("_suggested_nd", "sum"),
        建議ND_Stock_Value_CM=("_suggested_nd_stock_value", "sum"),
        建議ND_Shop_Stock_Qty_CM=("_suggested_nd_stock_qty", "sum"),
    ).reset_index()

    agg.insert(1, "Shop Name", agg["_site"].map(site_to_name).fillna(""))

    agg = agg.rename(columns={
        "_site": "Site Code",
        "SKU總數": "SKU 總數",
        "RF_SKU總數": "RF SKU 總數",
        "ND_SKU總數": "ND SKU 總數",
        "P90D0_Sales_SKU總數": "P90D 0 Sales SKU 總數",
        "SKU_Stock_Value_CM": "SKU Stock Value (CM)",
        "P90D0_Sales_Stock_Value_CM": "P90D 0 Sales Stock Value (CM)",
        "P90D0_Sales_Shop_Stock_Qty_CM": "P90D 0 Sales的店舖貨量",
        "建議ND_SKU總數": "建議ND SKU 總數",
        "建議ND_Stock_Value_CM": "建議ND Stock Value (CM)",
        "建議ND_Shop_Stock_Qty_CM": "建議ND Shop Stock Qty (CM)",
    })

    agg = agg.sort_values("Site Code")

    desired_order = [
        "Site Code", "Shop Name",
        "SKU 總數", "SKU Stock Value (CM)",
        "RF SKU 總數", "ND SKU 總數",
        "P90D 0 Sales SKU 總數", "P90D 0 Sales的店舖貨量",
        "P90D 0 Sales Stock Value (CM)",
        "建議ND SKU 總數", "建議ND Stock Value (CM)", "建議ND Shop Stock Qty (CM)",
        "Shop Name: P90D 0 Sales SKU 總數/SKU總數 (#P90D 0 Sales的店舖貨量)",
    ]
    existing = list(agg.columns)
    ordered = [c for c in desired_order if c in existing]
    ordered += [c for c in existing if c not in ordered]
    agg = agg[ordered]

    agg[desired_order[-1]] = ""

    total = {}
    for col in agg.columns:
        if col == "Site Code":
            total[col] = "總計"
        elif col == "Shop Name":
            total[col] = ""
        elif agg[col].dtype in ("int64", "float64"):
            total[col] = int(round(float(agg[col].sum())))
        else:
            total[col] = None
    total_row = pd.DataFrame([total])
    agg = pd.concat([agg, total_row], ignore_index=True)

    for col in agg.columns:
        if agg[col].dtype == "float64":
            agg[col] = agg[col].astype(int)

    return agg


def transform_carry(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
):
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    warnings = []

    _optional_cols = ["RP Type", "Adj. MCH1", "Adj. SST", "Site Code", "Shop Stock Qty (CM)", "Shop Stock Value (CM)"]
    for _col in _optional_cols:
        if _col not in df_input.columns:
            warnings.append(f"缺少可選欄位「{_col}」，相關統計以 0 計")

    if progress_callback:
        progress_callback(0.05, "驗證通過")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據（非空白或總計列）")

    if "RP Type" in df.columns and "Shop Stock Qty (CM)" in df.columns:
        rp_type = df["RP Type"].fillna("").astype(str).str.strip().str.upper()
        shop_stock_cm = _safe_numeric(df["Shop Stock Qty (CM)"])
        mask_nd_zero = (rp_type == "ND") & (shop_stock_cm == 0)
        df = df[~mask_nd_zero]

    if df.empty:
        raise ValueError("刪除 ND 且 Shop Stock Qty (CM)=0 的資料後無有效資料列")

    if "RP Type" in df.columns:
        rp_type_unique = df["RP Type"].fillna("").astype(str).str.strip().str.upper()
        non_nd_rf = set(rp_type_unique.unique()) - {"ND", "RF", ""}
        if non_nd_rf:
            warnings.append(f"偵測到非 ND/RF 的 RP Type 值: {', '.join(sorted(non_nd_rf))}，這些列將不納入樞紐統計")

    if progress_callback:
        progress_callback(0.2, "正規化 SKU...")

    df["SKU"] = df["SKU"].apply(_safe_sku)

    df["建議新RP Type"] = ""
    df["Buyer 回覆 ND/RF"] = ""
    df["保留RF原因"] = ""

    if "RP Type" in df.columns and "Shop Stock Status" in df.columns:
        rp_type = df["RP Type"].fillna("").astype(str).str.strip().str.upper()
        status = df["Shop Stock Status"].fillna("")
        cond = (rp_type == "RF") & (status == CARRY_P90D0_STATUS)
        df.loc[cond, "建議新RP Type"] = "ND"

    if progress_callback:
        progress_callback(0.3, "建構樞紐分析...")

    raw = df.copy()

    summary_left = _build_carry_pivot(df, "Adj. MCH1") if "Adj. MCH1" in df.columns else pd.DataFrame()
    summary_right = _build_carry_pivot(df, "Adj. SST") if "Adj. SST" in df.columns else pd.DataFrame()

    if progress_callback:
        progress_callback(0.6, "建構分析摘要...")

    dashboard_df = _build_carry_dashboard(df)

    if progress_callback:
        progress_callback(0.8, "轉換完成")

    return summary_left, summary_right, dashboard_df, raw, warnings


def generate_excel_carry(
    summary_left: pd.DataFrame,
    summary_right: pd.DataFrame,
    dashboard_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    progress_callback=None,
) -> BytesIO:
    wb = Workbook()

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    desc_font = Font(size=9, color="6B7280", italic=True)

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    # ── Sheet 1: RAW ──
    ws_raw = wb.active
    ws_raw.title = "RAW"

    _write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.1)

    if progress_callback:
        progress_callback(0.15, "建立 Summary 工作表...")

    # ── Sheet 2: Summary (雙區塊樞紐) ──
    ws_summary = wb.create_sheet(title="Summary")

    _write_summary_headers(ws_summary, header_font, header_fill, thin_border)

    left_cols = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]   # A-J: Site Code, Shop Name, Adj. MCH1, ND_P90D0, ND_Other, ND_Total, RF_P90D0, RF_Other, RF_Total, Grand_Total
    right_cols = [12, 13, 14, 15, 16, 17, 18, 19, 20, 21]  # L-U: Site Code, Shop Name, Adj. SST, ND_P90D0, ND_Other, ND_Total, RF_P90D0, RF_Other, RF_Total, Grand_Total
    detail_cols_left = [4, 5, 7, 8]
    detail_cols_right = [15, 16, 18, 19]

    data_cols = ["ND_P90D0", "ND_Other", "ND_Total", "RF_P90D0", "RF_Other", "RF_Total", "Grand_Total"]
    has_left = summary_left is not None and not summary_left.empty
    has_right = summary_right is not None and not summary_right.empty

    start_row = 7
    left_len = 0
    right_len = 0

    if has_left:
        left_len = len(summary_left)
        _write_pivot_block(ws_summary, summary_left, start_row, left_cols, detail_cols_left, data_cols,
                           thin_border, bold_font, progress_callback, prefix="左區塊")
    if has_right:
        right_len = len(summary_right)
        _write_pivot_block(ws_summary, summary_right, start_row, right_cols, detail_cols_right, data_cols,
                           thin_border, bold_font, progress_callback, prefix="右區塊")

    final_row = start_row + max(left_len, right_len) - 1
    _auto_width(ws_summary)

    if progress_callback:
        progress_callback(0.55, "建立分析摘要工作表...")

    # ── Sheet 3: Carry SKU 分析摘要 ──
    ws_dashboard = wb.create_sheet(title="Carry SKU 分析摘要")

    db_headers = list(dashboard_df.columns)
    last_col_letter = get_column_letter(len(db_headers))
    ws_dashboard.merge_cells(f"A1:{last_col_letter}1")
    title_cell = ws_dashboard.cell(row=1, column=1, value="Carry SKU 分析摘要 (Carry SKU Analysis Dashboard)")
    title_cell.font = Font(bold=True, size=14, color="1E40AF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx, h in enumerate(db_headers, 1):
        cell = ws_dashboard.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, row_data in dashboard_df.iterrows():
        excel_row = 4 + row_idx
        is_total = row_data.get("Site Code") == "總計"
        for col_idx, h in enumerate(db_headers, 1):
            cell = ws_dashboard.cell(row=excel_row, column=col_idx)
            if h == "Shop Name: P90D 0 Sales SKU 總數/SKU總數 (#P90D 0 Sales的店舖貨量)" and not is_total:
                col_b = get_column_letter(db_headers.index("Shop Name") + 1)
                col_c = get_column_letter(db_headers.index("SKU 總數") + 1)
                col_g = get_column_letter(db_headers.index("P90D 0 Sales SKU 總數") + 1)
                col_h = get_column_letter(db_headers.index("P90D 0 Sales的店舖貨量") + 1)
                cell.value = f'=CONCATENATE({col_b}{excel_row},":",{col_g}{excel_row},"/",{col_c}{excel_row}," (#",{col_h}{excel_row},")")'
            else:
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                if isinstance(val, float) and not isinstance(val, bool):
                    val = int(round(val))
                cell.value = val
                if isinstance(val, (int, float)) and col_idx >= 3:
                    cell.number_format = "#,##0"
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_total:
                cell.font = bold_font

    _auto_width(ws_dashboard)

    if progress_callback:
        progress_callback(0.9, "調整格式中...")

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf


def _write_summary_headers(ws, header_font, header_fill, thin_border):
    """Write rows 3–6 of the Summary sheet (dual-block headers)."""

    def _style(cell):
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Row 4 — top labels for the two pivot dimensions.
    row4_left = {4: "RP Type", 5: "Shop Stock Status2"}
    row4_right = {15: "RP Type", 16: "Shop Stock Status2"}
    for col_map in [row4_left, row4_right]:
        for col, val in col_map.items():
            c = ws.cell(row=4, column=col, value=val)
            _style(c)

    # Row 5 — count title and the RP-type/status bands.
    ws.merge_cells("A5:C5")
    ws.merge_cells("L5:N5")
    c = ws.cell(row=5, column=1, value="計數 - SKU")
    _style(c)
    c = ws.cell(row=5, column=12, value="計數 - SKU")
    _style(c)

    row5_left = {4: "ND", 5: "ND", 6: "Shop Stock Status", 7: "RF", 8: "RF"}
    row5_right = {15: "ND", 16: "ND", 17: "Shop Stock Status", 18: "RF", 19: "RF"}
    for col_map in [row5_left, row5_right]:
        for col, val in col_map.items():
            c = ws.cell(row=5, column=col, value=val)
            _style(c)

    # Row 6 — identifier labels and detail bands.
    row6_left = {1: "Site Code", 2: "Shop Name", 3: "Adj. MCH1",
                 4: "P90D 0 Sales", 5: "Other Status", 6: "ND 合計",
                 7: "P90D 0 Sales", 8: "Other Status", 9: "RF 合計", 10: "總計SKU"}
    row6_right = {12: "Site Code", 13: "Shop Name", 14: "Adj. SST",
                  15: "P90D 0 Sales", 16: "Other Status", 17: "ND 合計",
                  18: "P90D 0 Sales", 19: "Other Status", 20: "RF 合計", 21: "總計SKU"}
    for col_map in [row6_left, row6_right]:
        for col, val in col_map.items():
            c = ws.cell(row=6, column=col, value=val)
            _style(c)


def _write_pivot_block(ws, pivot_df, start_row, excel_cols, detail_cols, data_cols, thin_border, bold_font,
                       progress_callback=None, prefix=""):
    """
    Write a pivot block to the summary sheet.
    pivot_df columns: [Site Code, Shop Name, category_col, ND_P90D0, ND_Other, ND_Total, RF_P90D0, RF_Other, RF_Total, Grand_Total]
    excel_cols: list of 10 column indices [site, shop_name, cat, nd_p90d0, nd_other, nd_total, rf_p90d0, rf_other, rf_total, grand]
    detail_cols: list of column indices where 0 should be written as None
    """
    category_col_name = pivot_df.columns[2]

    total_rows = len(pivot_df)
    for row_idx, row_data in pivot_df.iterrows():
        excel_row = start_row + row_idx
        is_subtotal = "合計" in str(row_data["Site Code"])

        # Col 1: Site Code
        site_val = row_data["Site Code"]
        cell = ws.cell(row=excel_row, column=excel_cols[0], value=site_val if site_val else None)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if is_subtotal:
            cell.font = bold_font

        # Col 2: Shop Name
        shop_val = row_data.get("Shop Name", "")
        cell = ws.cell(row=excel_row, column=excel_cols[1], value=str(shop_val) if shop_val else None)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if is_subtotal:
            cell.font = bold_font

        # Col 3: category
        cat_val = row_data.get(category_col_name)
        if pd.isna(cat_val) or cat_val is None:
            cat_val = None
        cell = ws.cell(row=excel_row, column=excel_cols[2], value=str(cat_val) if cat_val is not None else None)
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if is_subtotal:
            cell.font = bold_font

        # Cols 4-10: numeric data
        for di, dc in enumerate(data_cols):
            col_excel = excel_cols[3 + di]
            val = int(row_data.get(dc, 0)) if pd.notna(row_data.get(dc)) else 0

            if col_excel in detail_cols:
                val = val if val > 0 else None

            cell = ws.cell(row=excel_row, column=col_excel, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_subtotal:
                cell.font = bold_font
            if isinstance(val, (int, float)) and val is not None:
                cell.number_format = "#,##0"

        if progress_callback and (row_idx % 500 == 0 or row_idx == total_rows - 1):
            progress_callback(0.15 + 0.4 * ((row_idx + 1) / max(total_rows, 1)),
                            f"寫入 Summary {prefix} ({row_idx+1}/{total_rows})...")


# ─── Stock Impact Analysis Functions ────────────────────────────────────────────
#
# 取代 Sales Decline Detection (v1.5.1)。改以 Shop Stock = 0（缺貨）對 Sold Qty 的
# 因果影響為核心，回答「無貨就無銷售」。Shop Nature 優先級 Tourist > Mix > Non Tourist。

# 三期 Stock→Sold 對應（同期）。每個 tuple: (期別標籤, Stock 欄位, Sold 欄位)
_STOCK_SOLD_PERIODS = [
    ("M-2", "Shop Stock Qty (M-2)", "Sold Qty (M-2)"),
    ("M-1", "Shop Stock Qty (M-1)", "Sold Qty (M-1)"),
    ("CM", "Shop Stock Qty (CM)", "Sold Qty (CM - MTD)"),
]

# Shop Nature 排序權重（越小越優先）
_NATURE_PRIORITY = {"Tourist": 0, "Mix": 1, "Non Tourist": 2}
_NATURE_ORDER = ["Tourist", "Mix", "Non Tourist"]


def _add_launch_info(df: pd.DataFrame, ref_date: datetime.date | None = None) -> pd.DataFrame:
    """從 Launch Date 新增 認證是否New SKU(3個月內) 欄位。

    Args:
        ref_date: 參考日期（預設今天）。
    """
    if "Launch Date" not in df.columns:
        return df  # 無 Launch Date 欄位則跳過

    if ref_date is None:
        ref_date = datetime.date.today()

    df = df.copy()
    launch_raw = df["Launch Date"].fillna("").astype(str).str.strip()
    # 嘗試多種日期格式解析
    launch_dt = pd.to_datetime(launch_raw, errors="coerce")
    # 若全部 NaN，再嘗試其他常見格式
    if launch_dt.isna().all():
        launch_dt = pd.to_datetime(launch_raw, errors="coerce", format="%Y/%m/%d")
    if launch_dt.isna().all():
        launch_dt = pd.to_datetime(launch_raw, errors="coerce", format="%d/%m/%Y")
    if launch_dt.isna().all():
        launch_dt = pd.to_datetime(launch_raw, errors="coerce", format="%Y%m%d")

    # 3個月內（約 90 天）視為 New SKU
    cutoff = pd.Timestamp(ref_date) - pd.Timedelta(days=90)
    df["Is New SKU (3mo)"] = launch_dt.notna() & (launch_dt >= cutoff)
    df["Is New SKU (3mo)"] = df["Is New SKU (3mo)"].map({True: "✓ New SKU", False: ""})

    # 保留 Launch Date 原始值
    df["Launch Date"] = launch_raw
    return df


def _add_stock_impact_columns(df: pd.DataFrame) -> pd.DataFrame:
    """計算每列的缺貨狀態、缺貨模式與銷售損失（雙基準並列）。

    缺貨定義採嚴格標準：Shop Stock Qty = 0。
    損失基準並列：(1) Sold Qty (M-2) 需求基準；(2) Shop Monthly Sold Qty (P90D) 90日總量（/3得月均）。
    """
    df = df.copy()
    idx = df.index

    # ── 缺貨判定（Shop Stock = 0）──────────────────────────────────────────────
    oos_flags = {}
    for label, stock_col, _sold_col in _STOCK_SOLD_PERIODS:
        stock = _safe_numeric(df.get(stock_col, _safe_default_series(0, idx)))
        is_oos = stock <= 0
        df[f"Is OOS ({label})"] = is_oos
        oos_flags[label] = is_oos

    oos_count = sum(oos_flags.values())
    df["OOS Periods Count"] = oos_count.astype(int)

    # ── 缺貨模式分類 ────────────────────────────────────────────────────────────
    pattern = np.full(len(df), "無缺貨", dtype=object)
    # 恢復供貨：M-2 缺貨（或 M-2、M-1 都缺貨）但 CM 有貨 → 後期補回
    recovered = (~oos_flags["CM"]) & (oos_flags["M-2"] | oos_flags["M-1"])
    # 間歇：僅 1 期缺貨（且非恢復供貨的 M-2/M-1 缺貨 → CM 有貨已涵蓋）
    intermittent = (oos_count == 1) & ~recovered
    # 長期：≥2 期缺貨
    long_term = oos_count >= 2
    pattern[recovered.values] = "恢復供貨"
    pattern[intermittent.values] = "間歇缺貨"
    pattern[long_term.values] = "長期缺貨"
    df["OOS Pattern"] = pattern

    # ── 銷售損失（雙基準並列）─────────────────────────────────────────────────
    # 需求基準 1：Sold Qty (M-2)
    m2_sold = _safe_numeric(df.get("Sold Qty (M-2)", _safe_default_series(0, idx)))
    # 需求基準 2：Shop Monthly Sold Qty (P90D)
    p90d = _safe_numeric(df.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, idx)))

    loss_m2_max = np.zeros(len(df))
    loss_p90d_max = np.zeros(len(df))
    loss_m2_sum = np.zeros(len(df))
    loss_p90d_sum = np.zeros(len(df))

    for label, _stock_col, sold_col in _STOCK_SOLD_PERIODS:
        sold = _safe_numeric(df.get(sold_col, _safe_default_series(0, idx)))
        is_oos = oos_flags[label].values
        # 缺貨當期才算損失；損失 = max(0, 需求基準 − 當期實際銷售)
        loss_m2 = np.where(is_oos & (m2_sold.values > 0), np.maximum(0, m2_sold.values - sold.values), 0)
        loss_p90d = np.where(is_oos & (p90d.values > 0), np.maximum(0, p90d.values / 3.0 - sold.values), 0)
        loss_m2_max = np.maximum(loss_m2_max, loss_m2)
        loss_p90d_max = np.maximum(loss_p90d_max, loss_p90d)
        loss_m2_sum = loss_m2_sum + loss_m2
        loss_p90d_sum = loss_p90d_sum + loss_p90d

    df["Max Sales Loss (M-2)"] = loss_m2_max.astype(int)
    df["Max Sales Loss (P90D)"] = loss_p90d_max.astype(int)
    df["Total Sales Loss (M-2)"] = loss_m2_sum.astype(int)
    df["Total Sales Loss (P90D)"] = loss_p90d_sum.astype(int)

    # ── 影響判定：任一期缺貨且有需求證據，且只計算 RF 店舖 ──────────────────────
    # RP Type 正規化（沿用既有模式：strip + upper），只認 "RF" 為有效店舖
    rp_type = (
        df.get("RP Type", _safe_default_series("", idx))
        .fillna("").astype(str).str.strip().str.upper()
    )
    is_rf = rp_type == "RF"
    has_demand = (m2_sold.values > 0) | (p90d.values > 0)
    any_oos = oos_count.values > 0
    df["Is Stock Impacted"] = any_oos & has_demand & is_rf.values

    # ── Turnday 分級與 Mix 可釋放存貨（存貨分佈分析）──────────────────────────────
    # Turnday (P90D Qty) 原值正規化：inf/NaN 代表 P90D=0（無周轉意義）
    turnday_raw = df.get("Turnday (P90D Qty)", _safe_default_series(0, idx))
    turnday = pd.to_numeric(turnday_raw, errors="coerce")
    # 將 inf（P90D=0 但有存貨）與 NaN 設為 -1 旗標，方便後續分類
    turnday_clean = turnday.replace([np.inf, -np.inf], np.nan).fillna(-1)
    df["Turnday (Days)"] = turnday_clean.where(turnday_clean >= 0, np.nan)

    # Mix 店舖 Turnday 分級（Mix 標準：保留 7-14 日）
    shop_nature = df.get("Shop Nature", _safe_default_series("", idx)).fillna("").astype(str)
    is_mix = shop_nature == "Mix"
    cm_stock_val = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, idx)))
    has_cm_stock = cm_stock_val.values > 0

    mix_band = np.full(len(df), "", dtype=object)
    mix_mask = is_mix.values
    # Mix 子集的 Turnday 值
    td = turnday_clean.values
    # 無存貨
    no_stock = ~has_cm_stock
    # 無銷售但有存貨（Turnday = -1 且有存貨）= 純囤積
    no_sale_has_stock = (td < 0) & has_cm_stock
    # 有銷售：依 Turnday 分級
    band_under7 = (td >= 0) & (td < 7)
    band_7_14 = (td >= 7) & (td < 14)
    band_14_30 = (td >= 14) & (td < 30)
    band_over30 = (td >= 30)

    mix_band[mix_mask & no_stock] = "無存貨"
    mix_band[mix_mask & no_sale_has_stock] = "無銷售有存貨"
    mix_band[mix_mask & band_under7] = "<7日(過低)"
    mix_band[mix_mask & band_7_14] = "7-14日(理想)"
    mix_band[mix_mask & band_14_30] = "14-30日(偏高)"
    mix_band[mix_mask & band_over30] = ">30日(過剩)"
    df["Mix Turnday Band"] = mix_band

    # Mix 可釋放存貨：超過 14 日標準的部分；無銷售但有存貨者全數可釋放
    # 理想存貨 = P90D / 90 × 14（P90D為90日總量）；可釋放 = max(0, 實際存貨 − 理想存貨)
    ideal_14d = p90d.values / 90.0 * 14
    excess = np.where(
        mix_mask,
        np.where(no_sale_has_stock, cm_stock_val.values, np.maximum(0, cm_stock_val.values - ideal_14d)),
        0,
    )
    df["Mix Excess Stock"] = excess.astype(int)

    # Non Tourist Turnday 分級（同樣標準，>14日釋放）
    is_nt = shop_nature == "Non Tourist"
    nt_band = np.full(len(df), "", dtype=object)
    nt_mask = is_nt.values
    nt_band[nt_mask & no_stock] = "無存貨"
    nt_band[nt_mask & no_sale_has_stock] = "無銷售有存貨"
    nt_band[nt_mask & band_under7] = "<7日(過低)"
    nt_band[nt_mask & band_7_14] = "7-14日(理想)"
    nt_band[nt_mask & band_14_30] = "14-30日(偏高)"
    nt_band[nt_mask & band_over30] = ">30日(過剩)"
    df["Non Tourist Turnday Band"] = nt_band

    # Non Tourist 可釋放存貨（同樣標準：>14日即可釋放，參考 Mix）
    nt_excess = np.where(
        nt_mask,
        np.where(no_sale_has_stock, cm_stock_val.values, np.maximum(0, cm_stock_val.values - ideal_14d)),
        0,
    )
    df["Non Tourist Excess Stock"] = nt_excess.astype(int)

    # 為「無銷售有存貨」行附上 RP Type（供後續拆分 ND/RF）
    rp_for_band = rp_type.values
    mix_band_with_rp = mix_band.copy()
    mix_band_with_rp[mix_mask & no_sale_has_stock] = np.where(
        rp_for_band[mix_mask & no_sale_has_stock] == "ND",
        "無銷售有存貨(ND)",
        "無銷售有存貨(RF)",
    )
    df["Mix Turnday Band"] = mix_band_with_rp

    nt_band_with_rp = nt_band.copy()
    nt_band_with_rp[nt_mask & no_sale_has_stock] = np.where(
        rp_for_band[nt_mask & no_sale_has_stock] == "ND",
        "無銷售有存貨(ND)",
        "無銷售有存貨(RF)",
    )
    df["Non Tourist Turnday Band"] = nt_band_with_rp

    return df


# 多維度彙總輸出欄位（中文表頭）
_STOCK_IMPACT_SUMMARY_COLS = [
    "缺貨行數", "涉及SKU數",
    "Tourist 缺貨行數", "Mix 缺貨行數", "Non Tourist 缺貨行數",
    "長期缺貨行數", "間歇缺貨行數", "恢復供貨行數",
]


def _build_stock_impact_summary(raw: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    """依 group_cols 彙總缺貨影響。只統計 Is Stock Impacted = True 的列。

    排序優先級：Tourist 缺貨 → 缺貨行數（皆降序）。
    """
    impacted = raw[raw["Is Stock Impacted"]].copy()
    if impacted.empty:
        col_names = group_cols + _STOCK_IMPACT_SUMMARY_COLS
        return pd.DataFrame(columns=col_names)

    impacted["_nature"] = impacted.get("Shop Nature", pd.Series(dtype=object)).fillna("")
    impacted["_pattern"] = impacted.get("OOS Pattern", pd.Series(dtype=object)).fillna("")

    # Vectorized pre-computation: avoid slow lambda in groupby
    impacted["_is_tourist"] = (impacted["_nature"] == "Tourist").astype(int)
    impacted["_is_mix"] = (impacted["_nature"] == "Mix").astype(int)
    impacted["_is_nt"] = (impacted["_nature"] == "Non Tourist").astype(int)
    impacted["_is_long"] = (impacted["_pattern"] == "長期缺貨").astype(int)
    impacted["_is_inter"] = (impacted["_pattern"] == "間歇缺貨").astype(int)
    impacted["_is_recover"] = (impacted["_pattern"] == "恢復供貨").astype(int)

    agg = (
        impacted.groupby(group_cols, dropna=False, observed=True)
        .agg(
            缺貨行數=("_nature", "count"),
            涉及SKU數=("SKU", "nunique"),
            **{
                "Tourist 缺貨行數": ("_is_tourist", "sum"),
                "Mix 缺貨行數": ("_is_mix", "sum"),
                "Non Tourist 缺貨行數": ("_is_nt", "sum"),
                "長期缺貨行數": ("_is_long", "sum"),
                "間歇缺貨行數": ("_is_inter", "sum"),
                "恢復供貨行數": ("_is_recover", "sum"),
            },
        )
        .reset_index()
    )

    for c in _STOCK_IMPACT_SUMMARY_COLS:
        agg[c] = agg[c].fillna(0).astype(int)

    agg = agg.sort_values(
        ["Tourist 缺貨行數", "缺貨行數"],
        ascending=[False, False],
    ).reset_index(drop=True)

    # 總計列（涉及SKU數為整體不重複計數，非各組加總）
    total_data = {col: "" for col in group_cols}
    total_data[group_cols[0]] = "總計"
    for c in _STOCK_IMPACT_SUMMARY_COLS:
        if c == "涉及SKU數":
            total_data[c] = int(impacted["SKU"].nunique())
        else:
            total_data[c] = int(agg[c].sum())
    agg = pd.concat([agg, pd.DataFrame([total_data])], ignore_index=True)

    for c in ["_nature", "_pattern"]:
        if c in agg.columns:
            agg.drop(columns=[c], inplace=True)

    return agg


def _build_sku_allocation_summary(raw: pd.DataFrame) -> pd.DataFrame:
    """SKU 級跨 Nature 調貨建議。

    針對同時存在於多種 Shop Nature 的 SKU：當 Tourist 店舖缺貨（CM Shop Stock=0）
    但同 SKU 在 Non-Tourist / Mix 店舖仍有存貨時，列出可調貨支援機會。
    排序：Tourist 缺貨損失 → 可調貨存貨量（降序）。
    """
    if "SKU" not in raw.columns or "Shop Nature" not in raw.columns:
        return pd.DataFrame()

    # 先找出同時出現在 Tourist 與其他 Nature 的 SKU（向量化替代 apply+set）
    _nature_dedup = raw.dropna(subset=["Shop Nature"]).groupby("SKU")["Shop Nature"].agg("unique")
    _nature_sets = _nature_dedup.apply(lambda arr: set(arr))
    shared_skus = _nature_sets[_nature_sets.apply(lambda s: "Tourist" in s and len(s) >= 2)].index
    if len(shared_skus) == 0:
        return pd.DataFrame()

    sub = raw[raw["SKU"].isin(shared_skus)].copy()
    if sub.empty:
        return pd.DataFrame()

    cm_stock = _safe_numeric(sub.get("Shop Stock Qty (CM)", _safe_default_series(0, sub.index)))
    cm_sold = _safe_numeric(sub.get("Sold Qty (CM - MTD)", _safe_default_series(0, sub.index)))
    m2_sold = _safe_numeric(sub.get("Sold Qty (M-2)", _safe_default_series(0, sub.index)))
    p90d = _safe_numeric(sub.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, sub.index)))
    loss_m2 = _safe_numeric(sub.get("Max Sales Loss (M-2)", _safe_default_series(0, sub.index)))
    loss_p90d = _safe_numeric(sub.get("Max Sales Loss (P90D)", _safe_default_series(0, sub.index)))
    sub["_cm_stock"] = cm_stock
    sub["_cm_sold"] = cm_sold
    sub["_m2_sold"] = m2_sold
    sub["_p90d"] = p90d
    sub["_loss_m2"] = loss_m2
    sub["_loss_p90d"] = loss_p90d

    # Tourist 缺貨行（CM Shop Stock = 0 且有需求）
    tourist_oos = sub[(sub["Shop Nature"] == "Tourist") & (cm_stock <= 0) & ((m2_sold > 0) | (p90d > 0))]
    # 其他 Nature 有存貨的行
    other_stock = sub[(sub["Shop Nature"] != "Tourist") & (cm_stock > 0)]

    if tourist_oos.empty or other_stock.empty:
        return pd.DataFrame()

    # 以 SKU + Brand 聚合 Tourist 缺貨損失
    t_agg = (
        tourist_oos.groupby(["SKU"] + (["Brand"] if "Brand" in tourist_oos.columns else []), dropna=False)
        .agg(
            Tourist缺貨店數=("SKU", "count"),
            Tourist損失_M2=("_loss_m2", "sum"),
            Tourist損失_P90D=("_loss_p90d", "sum"),
            Tourist_M2銷售件數=("_m2_sold", "sum"),
        )
        .reset_index()
    )

    # 其他 Nature 可調貨存貨
    o_agg = (
        other_stock.groupby(["SKU"] + (["Shop Nature"]), dropna=False)
        .agg(
            可調貨存貨件數=("_cm_stock", "sum"),
            閒置店數=("SKU", "count"),
        )
        .reset_index()
    )
    # 把其他 Nature 的存貨攤平成欄（Mix / Non Tourist）
    o_pivot = o_agg.pivot_table(
        index=["SKU"],
        columns="Shop Nature",
        values=["可調貨存貨件數", "閒置店數"],
        fill_value=0,
        aggfunc="sum",
    ).reset_index()
    # 扁平化多層欄名
    flat_cols = []
    for c in o_pivot.columns:
        if isinstance(c, tuple):
            metric, nature = c
            if metric:
                flat_cols.append(f"{nature}{metric}")
            else:
                flat_cols.append(nature)
        else:
            flat_cols.append(c)
    o_pivot.columns = flat_cols

    merged = t_agg.merge(o_pivot, on="SKU", how="left")

    # 確保 Mix/Non Tourist 欄位存在
    for n in ["Mix", "Non Tourist"]:
        for metric in ["可調貨存貨件數", "閒置店數"]:
            col = f"{n}{metric}"
            if col not in merged.columns:
                merged[col] = 0
            merged[col] = merged[col].fillna(0).astype(int)

    merged["可調貨總存貨件數"] = (merged.get("Mix可調貨存貨件數", 0) + merged.get("Non Tourist可調貨存貨件數", 0)).astype(int)
    merged = merged[merged["可調貨總存貨件數"] > 0]

    if merged.empty:
        return pd.DataFrame()

    # 欄位排序
    out_cols = ["SKU"]
    if "Brand" in merged.columns:
        out_cols.append("Brand")
    out_cols += [
        "Tourist缺貨店數", "Tourist_M2銷售件數",
        "Mix可調貨存貨件數", "Mix閒置店數", "Non Tourist可調貨存貨件數", "Non Tourist閒置店數",
        "可調貨總存貨件數",
    ]
    # 排序須在 subset 之前（因排序鍵 Tourist損失_P90D 不列入輸出）
    merged = merged.sort_values(
        ["Tourist損失_P90D", "可調貨總存貨件數"], ascending=[False, False]
    ).reset_index(drop=True)
    merged = merged[out_cols]

    return merged


def _classify_d001_status(d001_cm: float) -> str:
    """依 D001 倉庫 CM 存貨分類缺貨源頭。

    - 「倉庫有貨」(D001 CM > 0): 補貨流程問題，貨在倉庫卻沒送到店
    - 「源頭斷貨」(D001 CM = 0): 倉庫也斷貨，需上游補貨
    """
    return "倉庫有貨" if d001_cm > 0 else "源頭斷貨"


def _build_sku_stock_diagnosis(raw: pd.DataFrame) -> pd.DataFrame:
    """by-SKU 缺貨診斷表（Tourist + Mix + Non Tourist 缺貨 SKU）。

    每列為一個缺貨 SKU，彙總：
    - Tourist / Mix / Non Tourist 各自缺貨店數與銷售
    - D001 倉庫三期貨量 (M-2/M-1/CM) + D001 狀態分類（源頭斷貨 vs 補貨問題）
    - 店舖三期總存貨 (M-2/M-1/CM)
    排序：Tourist 缺貨店數 → Mix 缺貨店數 → Non Tourist 缺貨店數 → Tourist 損失(P90D)（皆降序）。
    """
    if "SKU" not in raw.columns or "Shop Nature" not in raw.columns:
        return pd.DataFrame()

    # 只看 Tourist + Mix + Non Tourist 的缺貨行（CM Shop Stock=0 且有需求）
    focus = raw[
        raw["Shop Nature"].isin(["Tourist", "Mix", "Non Tourist"])
        & (raw["Is Stock Impacted"])
        & (_safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index))) <= 0)
    ].copy()
    if focus.empty:
        return pd.DataFrame()

    # 補上需要的數值欄位
    focus["_loss_p90d"] = _safe_numeric(focus.get("Max Sales Loss (P90D)", _safe_default_series(0, focus.index)))
    focus["_loss_m2"] = _safe_numeric(focus.get("Max Sales Loss (M-2)", _safe_default_series(0, focus.index)))
    focus["_m2_sold"] = _safe_numeric(focus.get("Sold Qty (M-2)", _safe_default_series(0, focus.index)))
    focus["_cm_sold"] = _safe_numeric(focus.get("Sold Qty (CM - MTD)", _safe_default_series(0, focus.index)))
    focus["_p90d"] = _safe_numeric(focus.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, focus.index)))
    for label, col in [("M-2", "Shop Stock Qty (M-2)"), ("M-1", "Shop Stock Qty (M-1)"), ("CM", "Shop Stock Qty (CM)")]:
        focus[f"_shop_stock_{label}"] = _safe_numeric(focus.get(col, _safe_default_series(0, focus.index)))
    for label, col in [("M-2", "D001 Stock Qty (M-2)"), ("M-1", "D001 Stock Qty (M-1)"), ("CM", "D001 Stock Qty (CM)")]:
        focus[f"_d001_{label}"] = _safe_numeric(focus.get(col, _safe_default_series(0, focus.index)))

    # 先取每個 SKU 的 D001 三期值（D001 是 SKU 級共享，first 即可）
    d001_per_sku = (
        focus.groupby("SKU", dropna=False)
        .agg(
            D001_Stock_M2=("_d001_M-2", "first"),
            D001_Stock_M1=("_d001_M-1", "first"),
            D001_Stock_CM=("_d001_CM", "first"),
        )
        .reset_index()
    )

    # 各 Nature 的缺貨店數 + 銷售（只保留缺貨店數，取消缺貨行數）
    nature_aggs = {}
    _DIAG_NATURES = ["Tourist", "Mix", "Non Tourist"]
    for nat in _DIAG_NATURES:
        sub = focus[focus["Shop Nature"] == nat]
        nature_aggs[nat] = (
            sub.groupby("SKU", dropna=False)
            .agg(
                **{
                    f"{nat}缺貨店數": ("Site Code", "nunique") if "Site Code" in sub.columns else ("SKU", "count"),
                    f"{nat}損失_M2": ("_loss_m2", "sum"),
                    f"{nat}損失_P90D": ("_loss_p90d", "sum"),
                    f"{nat}_CM銷售件數": ("_cm_sold", "sum"),
                    f"{nat}_P90D總量(90日)": ("_p90d", "sum"),
                }
            )
            .reset_index()
        )

    # 店舖總存貨（Tourist+Mix+Non Tourist 缺貨店加總）
    shop_stock_per_sku = (
        focus.groupby("SKU", dropna=False)
        .agg(
            店舖存貨總_M2=("_shop_stock_M-2", "sum"),
            店舖存貨總_M1=("_shop_stock_M-1", "sum"),
            店舖存貨總_CM=("_shop_stock_CM", "sum"),
        )
        .reset_index()
    )

    # 品牌資訊（first）+ SKU Desc + Launch Info
    brand_agg_dict = {
        "Brand": ("Brand", "first") if "Brand" in focus.columns else ("SKU", "first"),
        "Buyer": ("Buyer", "first") if "Buyer" in focus.columns else ("SKU", "first"),
    }
    if "SKU Desc" in focus.columns:
        brand_agg_dict["SKU Desc"] = ("SKU Desc", "first")
    if "Launch Date" in focus.columns:
        brand_agg_dict["Launch Date"] = ("Launch Date", "first")
    if "Is New SKU (3mo)" in focus.columns:
        brand_agg_dict["Is New SKU (3mo)"] = ("Is New SKU (3mo)", "first")
    brand_per_sku = (
        focus.groupby("SKU", dropna=False)
        .agg(**brand_agg_dict)
        .reset_index()
    )

    # 各 SKU 的 NDRF 行數統計（從完整 raw 資料計算，分 ND / RF）
    rp_full = (
        raw.get("RP Type", _safe_default_series("", raw.index))
        .fillna("").astype(str).str.strip().str.upper()
    )
    raw_ndrf = raw.assign(_is_nd=(rp_full == "ND"), _is_rf=(rp_full == "RF"))
    ndrf_per_sku = (
        raw_ndrf.groupby("SKU", dropna=False)
        .agg(
            NDRF_ND行数=("_is_nd", "sum"),
            NDRF_RF行数=("_is_rf", "sum"),
        )
        .reset_index()
    )
    for c in ["NDRF_ND行数", "NDRF_RF行数"]:
        ndrf_per_sku[c] = ndrf_per_sku[c].fillna(0).astype(int)

    # 合併
    merged = brand_per_sku.merge(d001_per_sku, on="SKU", how="left")
    merged = merged.merge(ndrf_per_sku, on="SKU", how="left")
    for nat in _DIAG_NATURES:
        merged = merged.merge(nature_aggs[nat], on="SKU", how="left")
    merged = merged.merge(shop_stock_per_sku, on="SKU", how="left")

    # 填補 NDRF NaN
    for c in ["NDRF_ND行数", "NDRF_RF行数"]:
        merged[c] = merged[c].fillna(0).astype(int)

    # 填補缺漏 Nature 的 NaN
    for nat in _DIAG_NATURES:
        for col in nature_aggs[nat].columns:
            if col == "SKU":
                continue
            merged[col] = merged[col].fillna(0).astype(int)

    # D001 狀態分類
    merged["D001 狀態"] = merged["D001_Stock_CM"].apply(_classify_d001_status)

    # 欄位順序
    out_cols = ["SKU"]
    if "SKU Desc" in merged.columns:
        out_cols.append("SKU Desc")
    if "Launch Date" in merged.columns:
        out_cols.append("Launch Date")
    if "Is New SKU (3mo)" in merged.columns:
        out_cols.append("Is New SKU (3mo)")
    out_cols += ["Brand", "Buyer",
                "NDRF_ND行数", "NDRF_RF行数",
                "Tourist缺貨店數", "Tourist_CM銷售件數", "Tourist_P90D總量(90日)",
                "Mix缺貨店數", "Mix_CM銷售件數", "Mix_P90D總量(90日)",
                "Non Tourist缺貨店數", "Non Tourist_CM銷售件數", "Non Tourist_P90D總量(90日)",
                "D001 狀態", "D001_Stock_M2", "D001_Stock_M1", "D001_Stock_CM",
                "店舖存貨總_M2", "店舖存貨總_M1", "店舖存貨總_CM"]
    # 排序
    merged = merged.sort_values(
        ["Tourist缺貨店數", "Mix缺貨店數", "Non Tourist缺貨店數", "Tourist損失_P90D"],
        ascending=[False, False, False, False],
    ).reset_index(drop=True)
    avail = [c for c in out_cols if c in merged.columns]
    merged = merged[avail]

    return merged


# Turnday 分級順序（用於排序；含 RF/ND 拆分）
_TURNDAY_BAND_ORDER = [
    "<7日(過低)", "7-14日(理想)", "14-30日(偏高)", ">30日(過剩)",
    "無銷售有存貨(RF)", "無銷售有存貨(ND)", "無存貨",
]


def _build_inventory_distribution_summary(raw: pd.DataFrame, stock_threshold: int = 300) -> dict:
    """存貨分佈彙總 — 回傳 dict 含 4 個 DataFrame。

    Args:
        stock_threshold: 低存貨 SKU 門檻（D001 + Shop Stock 總和 < 此值視為低存貨，預設 300）。

    回傳:
        {
            "mix_turnday": Mix Turnday 分級表 (B1a),
            "nt_turnday": Non Tourist Turnday 分級表 (B1b),
            "nature_compare": Nature 存貨對比表 (B2),
            "imbalance_sku": 存貨失衡 SKU 清單 (B3, Tourist缺貨+其他過剩),
            "low_stock_allocation": 低存貨 SKU 配置決策表 (B4),
            "transfer_priority": 可調貨優先分析 (B5, Tourist缺貨+同SKU在Mix/NT有可釋放件數),
        }
    """
    result = {
        "mix_turnday": pd.DataFrame(),
        "nt_turnday": pd.DataFrame(),
        "nature_compare": pd.DataFrame(),
        "imbalance_sku": pd.DataFrame(),
        "low_stock_allocation": pd.DataFrame(),
        "transfer_priority": {},
    }

    if raw.empty or "Shop Nature" not in raw.columns:
        return result

    nature = raw["Shop Nature"].fillna("").astype(str)
    cm_stock = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))
    p90d = _safe_numeric(raw.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, raw.index)))
    turnday = pd.to_numeric(raw.get("Turnday (Days)", _safe_default_series(0, raw.index)), errors="coerce")

    # ── B1: Turnday 分級彙總（Mix + Non Tourist）─────────────────────────────────
    def _build_turnday_table(data, band_col, excess_col=None, label_prefix=""):
        """建構單一 Nature 的 Turnday 分級表。"""
        if data.empty or band_col not in data.columns:
            return pd.DataFrame()
        stock = _safe_numeric(data.get("Shop Stock Qty (CM)", _safe_default_series(0, data.index)))
        excess = _safe_numeric(data.get(excess_col, _safe_default_series(0, data.index))) if excess_col else _safe_default_series(0, data.index)
        band = data[band_col].fillna("無存貨")
        tbl = (
            data.assign(_stock=stock, _excess=excess, _band=band)
            .groupby("_band", dropna=False)
            .agg(行數=("_band", "count"), 存貨件數=("_stock", "sum"), 可釋放件數=("_excess", "sum"))
            .reset_index()
            .rename(columns={"_band": "Turnday 區間"})
        )
        tbl["__order"] = tbl["Turnday 區間"].map({b: i for i, b in enumerate(_TURNDAY_BAND_ORDER)}).fillna(99)
        tbl = tbl.sort_values("__order").drop(columns=["__order"]).reset_index(drop=True)
        for c in ["行數", "存貨件數", "可釋放件數"]:
            tbl[c] = tbl[c].fillna(0).astype(int)
        total = pd.DataFrame([{
            "Turnday 區間": "總計",
            "行數": int(tbl["行數"].sum()),
            "存貨件數": int(tbl["存貨件數"].sum()),
            "可釋放件數": int(tbl["可釋放件數"].sum()),
        }])
        return pd.concat([tbl, total], ignore_index=True)

    # B1a: Mix
    mix_mask = nature == "Mix"
    if mix_mask.any() and "Mix Turnday Band" in raw.columns:
        result["mix_turnday"] = _build_turnday_table(raw[mix_mask], "Mix Turnday Band", excess_col="Mix Excess Stock")

    # B1b: Non Tourist
    nt_mask = nature == "Non Tourist"
    if nt_mask.any() and "Non Tourist Turnday Band" in raw.columns:
        result["nt_turnday"] = _build_turnday_table(raw[nt_mask], "Non Tourist Turnday Band", excess_col="Non Tourist Excess Stock")

    # ── B2: Nature 存貨分佈對比 ───────────────────────────────────────────────
    b2_rows = []
    for nat in ["Tourist", "Mix", "Non Tourist"]:
        sub_mask = nature == nat
        if not sub_mask.any():
            continue
        sub_stock = cm_stock[sub_mask]
        sub_td = turnday[sub_mask]
        # 缺貨行數 (CM Stock=0)
        oos_count = int((cm_stock[sub_mask] == 0).sum())
        # 過剩行數 (>30日)
        over30 = int((sub_td > 30).sum())
        b2_rows.append({
            "Shop Nature": nat,
            "總行數": int(sub_mask.sum()),
            "存貨總量(件)": int(sub_stock.sum()),
            "缺貨行數(CM=0)": oos_count,
            "平均Turnday(日)": round(float(sub_td.mean()), 1) if sub_td.notna().any() else 0,
            "過剩行數(>30日)": over30,
        })
    if b2_rows:
        result["nature_compare"] = pd.DataFrame(b2_rows)

    # ── B3: 存貨失衡 SKU 清單 (Tourist缺貨 + 其他Nature過剩) ──────────────────
    # Tourist 缺貨且有需求的 SKU（Is Stock Impacted 子集）
    if "Is Stock Impacted" in raw.columns:
        tourist_oos = raw[(nature == "Tourist") & raw["Is Stock Impacted"] & (cm_stock <= 0)]
    else:
        tourist_oos = raw[(nature == "Tourist") & (cm_stock <= 0) & (p90d > 0)]

    # 其他 Nature 過剩存貨：Mix>14日（Mix Excess Stock>0）或 Non Tourist>30日
    other_over = raw[
        (nature != "Tourist")
        & (
            ((nature == "Mix") & (_safe_numeric(raw.get("Mix Excess Stock", _safe_default_series(0, raw.index))) > 0))
            | ((nature == "Non Tourist") & (turnday > 30) & (cm_stock > 0))
        )
    ]

    if tourist_oos.empty or other_over.empty or "SKU" not in raw.columns:
        result["imbalance_sku"] = pd.DataFrame()  # 無失衡 SKU 時回傳空
    else:
        # Tourist 缺貨損失（by SKU）
        t_loss = _safe_numeric(tourist_oos.get("Max Sales Loss (P90D)", _safe_default_series(0, tourist_oos.index)))
        tourist_oos = tourist_oos.assign(_loss=t_loss)
        t_agg = (
        tourist_oos.groupby(["SKU"] + (["Brand"] if "Brand" in tourist_oos.columns else []), dropna=False)
        .agg(
            Tourist缺貨店數=("SKU", "count"),
            Tourist缺貨損失_P90D=("_loss", "sum"),
        )
        .reset_index()
        )

        # 其他 Nature 過剩可釋放存貨（by SKU × Nature）
        other_over = other_over.assign(_stock=cm_stock[other_over.index])
        o_agg = (
        other_over.groupby(["SKU", "Shop Nature"], dropna=False)
        .agg(
            過剩存貨=("_stock", "sum"),
            過剩店數=("SKU", "count"),
        )
        .reset_index()
        )
        o_pivot_stock = o_agg.pivot_table(index="SKU", columns="Shop Nature", values="過剩存貨", fill_value=0).reset_index()
        o_pivot_shops = o_agg.pivot_table(index="SKU", columns="Shop Nature", values="過剩店數", fill_value=0).reset_index()

        merged = t_agg.merge(o_pivot_stock, on="SKU", how="inner")
        # 重新命名過剩存貨欄
        rename_map = {c: f"{c}過剩存貨" for c in o_pivot_stock.columns if c != "SKU"}
        merged = merged.rename(columns=rename_map)

        # 確保 Mix/Non Tourist 過剩欄存在
        for n in ["Mix", "Non Tourist"]:
            col = f"{n}過剩存貨"
            if col not in merged.columns:
                merged[col] = 0
            merged[col] = merged[col].fillna(0).astype(int)

        merged["其他Nature過剩總存貨"] = (merged["Mix過剩存貨"] + merged["Non Tourist過剩存貨"]).astype(int)
        merged = merged[merged["其他Nature過剩總存貨"] > 0]
        # 可調貨建議量 = min(Tourist 缺貨損失, 其他過剩存貨)
        merged["建議調貨量"] = np.minimum(merged["Tourist缺貨損失_P90D"], merged["其他Nature過剩總存貨"]).astype(int)

        # 附加 SKU 級資訊：SKU Desc / Launch Date / Buyer / D001 Net Stock + Shop Stock 總數
        sku_info_cols = ["SKU"]
        for _c in ["SKU Desc", "Launch Date", "Is New SKU (3mo)", "Buyer"]:
            if _c in raw.columns:
                sku_info_cols.append(_c)
        d001_net = "D001 Net Stock Qty (CM)"
        has_d001_net = d001_net in raw.columns
        has_shop_cm = "Shop Stock Qty (CM)" in raw.columns
        if sku_info_cols:
            agg_dict = {}
            for c in sku_info_cols:
                if c != "SKU":
                    agg_dict[c] = (c, "first")
            if has_d001_net:
                agg_dict["_d001_net"] = (d001_net, "first")
            if has_shop_cm:
                agg_dict["_shop_total"] = ("Shop Stock Qty (CM)", "sum")
            sku_info = (
                raw[sku_info_cols + ([d001_net] if has_d001_net else []) + (["Shop Stock Qty (CM)"] if has_shop_cm else [])]
                .groupby("SKU", dropna=False)
                .agg(**agg_dict)
                .reset_index()
            )
            merged = merged.merge(sku_info, on="SKU", how="left")
            if has_d001_net and has_shop_cm:
                merged["總可用存貨"] = (_safe_numeric(merged.get("_d001_net", _safe_default_series(0, merged.index))) +
                                        _safe_numeric(merged.get("_shop_total", _safe_default_series(0, merged.index)))).astype(int)
            for _c in ["_d001_net", "_shop_total"]:
                if _c in merged.columns:
                    merged.drop(columns=[_c], inplace=True)

        out_cols = ["SKU"]
        if "Brand" in merged.columns:
                out_cols.append("Brand")
        if "SKU Desc" in merged.columns:
            out_cols.append("SKU Desc")
        if "Launch Date" in merged.columns:
            out_cols.append("Launch Date")
        if "Is New SKU (3mo)" in merged.columns:
            out_cols.append("Is New SKU (3mo)")
        if "Buyer" in merged.columns:
            out_cols.append("Buyer")
        if "總可用存貨" in merged.columns:
            out_cols.append("總可用存貨")
        out_cols += ["Tourist缺貨店數", "Mix過剩存貨", "Non Tourist過剩存貨", "其他Nature過剩總存貨", "建議調貨量"]
        merged = merged[[c for c in out_cols if c in merged.columns]]
        merged = merged.sort_values(["Tourist缺貨店數", "建議調貨量"], ascending=[False, False]).reset_index(drop=True)

        # 附註解（計算方式說明）
        merged.attrs["annotation"] = (
            "【計算方式】\n"
            "Mix過剩存貨：Mix 店舖 Shop Stock (CM) > 0 且 Turnday > 14 日的過剩部分 "
            "（超出 P90D/90×14 標準的存貨總和）。\n"
            "Non Tourist過剩存貨：Non Tourist 店舖 Shop Stock (CM) > 0 且 Turnday > 30 日的存貨總和。\n"
            "建議調貨量 = min(Tourist 缺貨店數對應的需求量, Mix過剩存貨 + Non Tourist過剩存貨)。"
        )
        result["imbalance_sku"] = merged

    # ── B4: 低存貨 SKU 配置決策（D001 + Shop Stock < 門檻）─────────────────────
    # 找出總可用存貨不足門檻的 SKU，分析是否該減少 Non Tourist RF 店舖。
    # 同時統計 RF 店「無銷售有存貨」與 ND 店存貨（有/無銷售）的浪費情況。
    if "SKU" in raw.columns:
        d001_cm_col = "D001 Stock Qty (CM)" if "D001 Stock Qty (CM)" in raw.columns else None
        d001_cm = _safe_numeric(raw.get(d001_cm_col or "SKU", _safe_default_series(0, raw.index)))
        if d001_cm_col is None:
            d001_cm = pd.Series(0, index=raw.index)  # 無 D001 欄位時視為 0
        low_df = raw.assign(_d001=d001_cm, _shop_stock=cm_stock, _p90d=p90d)
        # 每個 SKU 的總可用存貨
        sku_avail = (
            low_df.groupby("SKU", dropna=False)
            .agg(總可用存貨=("_d001", "first"), _shop_total=("_shop_stock", "sum"))
            .reset_index()
        )
        sku_avail["總可用存貨"] = sku_avail["總可用存貨"] + sku_avail["_shop_total"]
        low_skus = set(sku_avail[sku_avail["總可用存貨"] < stock_threshold]["SKU"])

        if low_skus:
            low = low_df[low_df["SKU"].isin(low_skus)].copy()
            rp = low.get("RP Type", _safe_default_series("", low.index)).fillna("").astype(str).str.strip().str.upper()

            # Tourist/Mix 缺貨且有需求（RF）— 代表真的不夠貨
            impacted_mask = low.get("Is Stock Impacted", pd.Series([False] * len(low), index=low.index))
            t_oos = int(((low["Shop Nature"] == "Tourist") & (impacted_mask)).sum()) if "Shop Nature" in low.columns else 0
            m_oos = int(((low["Shop Nature"] == "Mix") & (impacted_mask)).sum()) if "Shop Nature" in low.columns else 0

            # Non Tourist RF 店「無銷售有存貨」— 浪費，建議減店
            nt_rf = low[(low["Shop Nature"] == "Non Tourist") & (rp == "RF")] if "Shop Nature" in low.columns else pd.DataFrame()
            nt_waste_mask = (nt_rf["_p90d"] == 0) & (nt_rf["_shop_stock"] > 0)
            nt_waste_rows = int(nt_waste_mask.sum())
            nt_waste_stock = int(nt_rf.loc[nt_waste_mask, "_shop_stock"].sum())

            # ND 店存貨統計（有/無銷售）— 計畫淘汰卻佔存貨
            nd_all = low[rp == "ND"]
            nd_has_stock = nd_all[nd_all["_shop_stock"] > 0]
            nd_stock_total = int(nd_has_stock["_shop_stock"].sum())
            nd_sale = nd_has_stock[nd_has_stock["_p90d"] > 0]
            nd_nosale = nd_has_stock[nd_has_stock["_p90d"] == 0]
            nd_sale_rows = int(len(nd_sale))
            nd_nosale_rows = int(len(nd_nosale))
            nd_nosale_stock = int(nd_nosale["_shop_stock"].sum())

            b4 = pd.DataFrame([{
                "低存貨門檻(件)": stock_threshold,
                "低存貨SKU數": len(low_skus),
                "Tourist缺貨行數": t_oos,
                "Mix缺貨行數": m_oos,
                "NT_RF無銷售有存貨行數": nt_waste_rows,
                "NT_RF無銷售佔用存貨(件)": nt_waste_stock,
                "ND有存貨行數": nd_sale_rows + nd_nosale_rows,
                "ND存貨總量(件)": nd_stock_total,
                "ND有存貨有銷售行數": nd_sale_rows,
                "ND有存貨無銷售行數": nd_nosale_rows,
                "ND無銷售浪費存貨(件)": nd_nosale_stock,
            }])
            result["low_stock_allocation"] = b4

    # ── B5: 可調貨優先分析 (Tourist缺貨 + 同SKU在Mix/NT有可釋放件數) ──────────
    result["transfer_priority"] = _build_transfer_priority_analysis(raw)

    return result


def _build_transfer_priority_analysis(raw: pd.DataFrame) -> dict:
    """可調貨優先分析：Tourist 缺貨行中，同 SKU 在 Mix/Non Tourist 有可釋放件數的行數及 SKU 明細。

    回傳:
        {
            "summary": DataFrame with aggregate metrics (可調貨Tourist行數, 可調貨SKU數),
            "detail": DataFrame with per-SKU transfer priority details,
        }
    """
    result: dict = {
        "summary": pd.DataFrame(),
        "detail": pd.DataFrame(),
    }

    if raw.empty or "Shop Nature" not in raw.columns or "SKU" not in raw.columns:
        return result

    nature = raw["Shop Nature"].fillna("").astype(str)
    cm_stock = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))

    # 1. Tourist OOS rows (Is Stock Impacted + CM stock <= 0)
    is_impacted = raw.get("Is Stock Impacted", pd.Series(False, index=raw.index))
    tourist_oos_mask = (nature == "Tourist") & is_impacted & (cm_stock <= 0)
    if not tourist_oos_mask.any():
        return result

    tourist_oos = raw[tourist_oos_mask].copy()

    # 2. Aggregate Mix/NT releasable stock by SKU（可釋放件數 = Mix Excess Stock + Non Tourist Excess Stock）
    mix_excess = _safe_numeric(raw.get("Mix Excess Stock", _safe_default_series(0, raw.index)))
    nt_excess = _safe_numeric(raw.get("Non Tourist Excess Stock", _safe_default_series(0, raw.index)))

    sku_releasable = (
        raw.assign(_mix_excess=mix_excess, _nt_excess=nt_excess)
        .groupby("SKU", dropna=False)
        .agg(
            Mix可釋放件數=("_mix_excess", "sum"),
            NonTourist可釋放件數=("_nt_excess", "sum"),
        )
        .reset_index()
    )
    sku_releasable["可釋放總件數"] = sku_releasable["Mix可釋放件數"] + sku_releasable["NonTourist可釋放件數"]
    skus_with_releasable = set(sku_releasable[sku_releasable["可釋放總件數"] > 0]["SKU"])

    if not skus_with_releasable:
        return result

    # 3. Filter Tourist OOS to only those SKUs with releasable stock
    transferrable_mask = tourist_oos["SKU"].isin(skus_with_releasable)
    if not transferrable_mask.any():
        return result

    transferrable_oos = tourist_oos[transferrable_mask]

    # 4. Count metrics
    row_count = len(transferrable_oos)
    sku_count = transferrable_oos["SKU"].nunique()

    summary = pd.DataFrame([{
        "可調貨Tourist行數": row_count,
        "可調貨SKU數": sku_count,
    }])
    result["summary"] = summary

    # 5. Build detail table: per SKU
    # Tourist缺失件數 = max(0, Adj. Safety Stock - Shop Stock Qty (CM))
    adj_safety = _safe_numeric(transferrable_oos.get("Adj. Safety Stock", _safe_default_series(0, transferrable_oos.index)))
    shop_stock = _safe_numeric(transferrable_oos.get("Shop Stock Qty (CM)", _safe_default_series(0, transferrable_oos.index)))
    shortage = np.maximum(0, adj_safety.values - shop_stock.values)
    transferrable_oos = transferrable_oos.assign(_shortage=shortage)

    t_agg = (
        transferrable_oos.groupby("SKU", dropna=False)
        .agg(
            Tourist缺貨行數=("SKU", "count"),
            Tourist缺失件數=("_shortage", "sum"),
        )
        .reset_index()
    )

    # Merge with releasable stock info
    detail = t_agg.merge(sku_releasable, on="SKU", how="left")
    detail = detail[detail["可釋放總件數"] > 0]

    # Add SKU-level info: Brand, SKU Desc, Buyer
    sku_info_cols = ["SKU"]
    for _c in ["Brand", "SKU Desc", "Buyer"]:
        if _c in raw.columns:
            sku_info_cols.append(_c)

    sku_info = (
        raw[sku_info_cols].groupby("SKU", dropna=False)
        .agg({c: "first" for c in sku_info_cols if c != "SKU"})
        .reset_index()
    )
    detail = detail.merge(sku_info, on="SKU", how="left")

    # D001 status
    d001_cm_col = "D001 Stock Qty (CM)"
    if d001_cm_col in raw.columns:
        d001_per_sku = (
            raw.groupby("SKU", dropna=False)
            .agg(D001_Stock_CM=(d001_cm_col, "first"))
            .reset_index()
        )
        d001_per_sku["D001狀態"] = d001_per_sku["D001_Stock_CM"].apply(_classify_d001_status)
        detail = detail.merge(d001_per_sku[["SKU", "D001狀態"]], on="SKU", how="left")
    else:
        detail["D001狀態"] = "—"

    # Column ordering
    out_cols = ["SKU"]
    for _c in ["Brand", "SKU Desc", "Buyer"]:
        if _c in detail.columns:
            out_cols.append(_c)
    out_cols += ["Tourist缺貨行數", "Tourist缺失件數", "Mix可釋放件數", "NonTourist可釋放件數", "可釋放總件數", "D001狀態"]

    detail = detail[[c for c in out_cols if c in detail.columns]]

    # Sort: Tourist OOS count desc, then releasable total desc
    detail = detail.sort_values(["Tourist缺貨行數", "可釋放總件數"], ascending=[False, False]).reset_index(drop=True)

    # Cast int columns
    for c in ["Tourist缺貨行數", "Tourist缺失件數", "Mix可釋放件數", "NonTourist可釋放件數", "可釋放總件數"]:
        if c in detail.columns:
            detail[c] = detail[c].fillna(0).astype(int)

    result["detail"] = detail

    return result


def _build_sku_monthly_analysis(raw: pd.DataFrame, stock_threshold: int = 300) -> pd.DataFrame:
    """按 SKU 計算遊客/混合/非遊客的每月銷量及存儲量分析。

    每列為一個 SKU，彙總各 Nature 的月均銷售、當期庫存、店數等。
    低存貨 SKU（D001 + Shop < threshold）排在前面。

    排序：低存貨優先 → Tourist RF 月均銷售降序。
    """
    if raw.empty or "SKU" not in raw.columns or "Shop Nature" not in raw.columns:
        return pd.DataFrame()

    _NATURES = ["Tourist", "Mix", "Non Tourist"]

    nature = raw["Shop Nature"].fillna("").astype(str)
    cm_stock = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))
    p90d = _safe_numeric(raw.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, raw.index)))
    cm_sold = _safe_numeric(raw.get("Sold Qty (CM - MTD)", _safe_default_series(0, raw.index)))
    turnday = pd.to_numeric(raw.get("Turnday (Days)", _safe_default_series(0, raw.index)), errors="coerce")
    rp_type = (
        raw.get("RP Type", _safe_default_series("", raw.index))
        .fillna("").astype(str).str.strip().str.upper()
    )
    d001_col = "D001 Stock Qty (CM)"
    d001_cm = _safe_numeric(raw[d001_col]) if d001_col in raw.columns else _safe_default_series(0, raw.index)

    work = raw[["SKU", "Shop Nature"]].copy()
    work["_nature"] = nature.values
    work["_cm_stock"] = cm_stock.values
    work["_p90d"] = p90d.values
    work["_cm_sold"] = cm_sold.values
    work["_turnday"] = turnday.values
    work["_d001"] = d001_cm.values
    work["_rp"] = rp_type.values
    work["_is_rf"] = (work["_rp"] == "RF").astype(int)

    # SKU 級基本資訊
    sku_base = (
        work.groupby("SKU", dropna=False)
        .agg(
            D001_庫存=("_d001", "first"),
            總店舖庫存=("_cm_stock", "sum"),
            總行數=("SKU", "count"),
            RF行數=("_is_rf", "sum"),
            ND行數=("_rp", lambda x: (x == "ND").sum()),
        )
        .reset_index()
    )
    # 從 raw 取 Brand/Buyer/SKU Desc
    for col, out_name in [("Brand", "Brand"), ("Buyer", "Buyer"), ("SKU Desc", "SKU Desc")]:
        if col in raw.columns:
            col_first = raw.groupby("SKU", dropna=False)[col].first().reset_index()
            sku_base = sku_base.merge(col_first, on="SKU", how="left")
    sku_base["RF行數"] = sku_base["RF行數"].fillna(0).astype(int)
    sku_base["ND行數"] = sku_base["ND行數"].fillna(0).astype(int)
    sku_base["總可用存貨"] = sku_base["D001_庫存"] + sku_base["總店舖庫存"]
    sku_base["低存貨"] = sku_base["總可用存貨"] < stock_threshold

    # SKU × Nature 級聚合
    sku_nat = (
        work.groupby(["SKU", "_nature"], dropna=False, observed=True)
        .agg(
            月均銷售=("_p90d", "sum"),
            當期CM銷售=("_cm_sold", "sum"),
            當期庫存=("_cm_stock", "sum"),
            店數=("SKU", "count"),
            平均Turnday=("_turnday", "mean"),
        )
        .reset_index()
    )
    sku_nat["月均銷售"] = sku_nat["月均銷售"] / 3.0

    # RF 月均銷售（只計 RF 店）
    work_rf = work[work["_rp"] == "RF"]
    if not work_rf.empty:
        sku_nat_rf = (
            work_rf.groupby(["SKU", "_nature"], dropna=False, observed=True)
            .agg(RF月均銷售=("_p90d", "sum"))
            .reset_index()
        )
        sku_nat_rf["RF月均銷售"] = sku_nat_rf["RF月均銷售"] / 3.0
        sku_nat = sku_nat.merge(sku_nat_rf, on=["SKU", "_nature"], how="left")
    else:
        sku_nat["RF月均銷售"] = 0.0

    # 展開各 Nature 為寬表
    result = sku_base.copy()
    for nat in _NATURES:
        nat_data = sku_nat[sku_nat["_nature"] == nat].copy()
        if nat_data.empty:
            result[f"{nat}_月均銷售"] = 0.0
            result[f"{nat}_當期CM銷售"] = 0.0
            result[f"{nat}_當期庫存"] = 0
            result[f"{nat}_店數"] = 0
            result[f"{nat}_平均Turnday"] = 0.0
            result[f"{nat}_RF月均銷售"] = 0.0
            continue
        nat_data = nat_data.drop(columns=["_nature"])
        nat_data.columns = [f"{nat}_{c}" if c != "SKU" else c for c in nat_data.columns]
        result = result.merge(nat_data, on="SKU", how="left")

    # 填補 NaN
    for nat in _NATURES:
        result[f"{nat}_平均Turnday"] = result.get(f"{nat}_平均Turnday", 0).fillna(0).round(1)
        for c in [f"{nat}_月均銷售", f"{nat}_當期CM銷售", f"{nat}_RF月均銷售"]:
            result[c] = result.get(c, 0).fillna(0).round(1)
        for c in [f"{nat}_當期庫存", f"{nat}_店數"]:
            result[c] = result.get(c, 0).fillna(0).astype(int)

    # 排序：低存貨優先 → Tourist RF 月均銷售降序
    result = result.sort_values(
        ["低存貨", "Tourist_RF月均銷售"],
        ascending=[False, False],
    ).reset_index(drop=True)

    # 轉整數
    for c in ["D001_庫存", "總店舖庫存", "總可用存貨", "總行數", "RF行數", "ND行數"]:
        result[c] = result[c].fillna(0).astype(int)
    for nat in _NATURES:
        result[f"{nat}_當期庫存"] = result[f"{nat}_當期庫存"].fillna(0).astype(int)
        result[f"{nat}_店數"] = result[f"{nat}_店數"].fillna(0).astype(int)

    # 欄位順序
    out_cols = ["SKU"]
    for col in ["SKU Desc", "Brand", "Buyer"]:
        if col in result.columns:
            out_cols.append(col)
    out_cols += ["低存貨", "總可用存貨", "D001_庫存", "總店舖庫存", "總行數", "RF行數", "ND行數"]
    for nat in _NATURES:
        out_cols += [
            f"{nat}_月均銷售", f"{nat}_當期CM銷售", f"{nat}_當期庫存",
            f"{nat}_店數", f"{nat}_平均Turnday", f"{nat}_RF月均銷售",
        ]
    out_cols = [c for c in out_cols if c in result.columns]

    return result[out_cols]


_OOS_DETAIL_COLS = [
    "Buyer", "Brand", "SKU", "SKU Desc", "Item Class", "Launch Date", "Is New SKU (3mo)",
    "Site Code", "Shop Name", "OM",
    "Shop Nature", "Supply Source", "NDRF Status",
    "Shop Stock Qty (M-2)", "Shop Stock Qty (M-1)",
    "Shop Stock Qty (CM)", "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
    "Shop Monthly Sold Qty (P90D)", "Is OOS (M-2)", "Is OOS (M-1)", "Is OOS (CM)",
    "OOS Periods Count", "OOS Pattern",
]



# ── 平行化輔助 ──────────────────────────────────────────────────────────────────
# 批次預處理數值欄位，避免各 summary 函式重複呼叫 _safe_numeric
_STOCK_IMPACT_NUMERIC_COLS = [
    "Shop Stock Qty (M-2)", "Shop Stock Qty (M-1)", "Shop Stock Qty (CM)",
    "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
    "Shop Monthly Sold Qty (P90D)", "D001 Stock Qty (CM)",
    "D001 Stock Qty (M-2)", "D001 Stock Qty (M-1)",
    "Max Sales Loss (M-2)", "Max Sales Loss (P90D)",
    "Mix Excess Stock", "Non Tourist Excess Stock",
]

def _preprocess_numeric_columns(raw: pd.DataFrame) -> pd.DataFrame:
    """一次性轉換所有後續會用到的數值欄位。"""
    for _col in _STOCK_IMPACT_NUMERIC_COLS:
        if _col in raw.columns:
            raw[_col] = pd.to_numeric(raw[_col], errors="coerce").fillna(0)
    return raw


def _build_summary_safe(raw: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    """安全的摘要建構：檢查欄位是否存在。"""
    if not group_cols or group_cols[0] not in raw.columns:
        return pd.DataFrame()
    return _build_stock_impact_summary(raw, group_cols)

def transform_stock_impact(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
    warehouse_only: bool = True,
    stock_threshold: int = 300,
    ref_date: datetime.date | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, dict, list[str]]:
    """Stock Impact Analysis 主流程。

    Args:
        warehouse_only: 若 True（預設），只分析 Supply Source = Warehouse 的列。
        ref_date: Launch Date 計算 New SKU 的參考日期（預設今天，建議傳入 MTD 截止日期）。
    Returns:
        (summary_site, summary_brand, summary_buyer, summary_category, sku_allocation, sku_diagnosis, inv_distribution, sku_monthly_analysis, raw_df, warnings)
    """
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    warnings = []

    _optional_cols = [
        "Shop Stock Qty (M-2)", "Shop Stock Qty (M-1)", "Shop Stock Qty (CM)",
        "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
        "Shop Monthly Sold Qty (P90D)", "Supply Source", "Shop Nature", "Adj. MCH1",
    ]
    for _col in _optional_cols:
        if _col not in df_input.columns:
            warnings.append(f"缺少可選欄位「{_col}」，相關統計將受影響")

    if progress_callback:
        progress_callback(0.05, "驗證通過")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據")

    # 預設只分析 Warehouse SKU
    if warehouse_only and "Supply Source" in df.columns:
        before = len(df)
        df = df[df["Supply Source"].astype(str).str.strip() == "Warehouse"].copy()
        if before > 0 and len(df) == 0:
            warnings.append("勾選「只分析 Warehouse SKU」但來源檔案無 Warehouse 列，已自動取消篩選")
            df = _clean_dataframe(df_input, shop_natures=shop_natures)

    df["SKU"] = df["SKU"].apply(_safe_sku)

    if "SKU Desc" not in df.columns:
        df = _compute_raw_extra_columns(df)

    if progress_callback:
        progress_callback(0.2, "計算缺貨影響指標中...")

    raw = _add_stock_impact_columns(df)

    # 新增 Launch Date / Is New SKU 資訊（於所有 Sheet 共用）
    raw = _add_launch_info(raw, ref_date=ref_date)

    # 批次預處理數值欄位（避免各 summary 函式重複轉換）
    raw = _preprocess_numeric_columns(raw)

    # groupby 加速：將 groupby 欄位轉為 category
    _groupby_cols = ["Brand", "Buyer", "Item Class"]
    if "Adj. MCH1" in raw.columns:
        _groupby_cols.append("Adj. MCH1")
    for _gc in _groupby_cols:
        if _gc in raw.columns:
            raw[_gc] = raw[_gc].astype("category")

    # 注意：stock_threshold 不再用於全域過濾，只會在 low_stock_allocation（B4）
    # 中內部使用，分析所有低存貨 SKU 的配置決策。
    if progress_callback:
        progress_callback(0.4, "建構多維度彙總中...")

    # Site 層級已取消（v1.6.2）
    summary_site = pd.DataFrame()

    # 平行化：同時計算 4 個維度的摘要表（彼此獨立）
    _summary_tasks = {}
    if "Brand" in raw.columns:
        _summary_tasks["brand"] = ["Brand"]
    if "Buyer" in raw.columns:
        _summary_tasks["buyer"] = ["Buyer"]
    if "Adj. MCH1" in raw.columns:
        _summary_tasks["category"] = ["Adj. MCH1"]
    else:
        warnings.append("缺少「Adj. MCH1」欄位，無法產出 By Category 分析")
    if "Item Class" in raw.columns:
        _summary_tasks["item_class"] = ["Item Class"]
    else:
        warnings.append("缺少「Item Class」欄位，無法產出 By Item Class 分析")

    if _summary_tasks:
        with ThreadPoolExecutor(max_workers=min(4, len(_summary_tasks))) as _executor:
            _futures = {
                _key: _executor.submit(_build_summary_safe, raw, _cols)
                for _key, _cols in _summary_tasks.items()
            }
            _results = {key: fut.result() for key, fut in _futures.items()}
    else:
        _results = {}

    summary_brand = _results.get("brand", pd.DataFrame())
    summary_buyer = _results.get("buyer", pd.DataFrame())
    summary_category = _results.get("category", pd.DataFrame())
    summary_item_class = _results.get("item_class", pd.DataFrame())

    if progress_callback:
        progress_callback(0.55, "建構 SKU 跨 Nature 調貨分析中...")

    sku_allocation = _build_sku_allocation_summary(raw)

    if progress_callback:
        progress_callback(0.62, "建構 by-SKU 缺貨診斷表中...")

    sku_diagnosis = _build_sku_stock_diagnosis(raw)

    if progress_callback:
        progress_callback(0.66, "建構存貨分佈分析中...")

    inv_distribution = _build_inventory_distribution_summary(raw, stock_threshold=stock_threshold)

    if progress_callback:
        progress_callback(0.72, "建構 SKU 月銷量和存儲量分析中...")

    sku_monthly_analysis = _build_sku_monthly_analysis(raw, stock_threshold=stock_threshold)

    if progress_callback:
        progress_callback(0.75, "轉換完成")

    return summary_site, summary_brand, summary_buyer, summary_category, summary_item_class, sku_allocation, sku_diagnosis, inv_distribution, sku_monthly_analysis, raw, warnings


def _write_oos_detail_sheet(ws, raw_df, header_font, header_fill, thin_border, progress_callback):
    """寫入缺貨影響明細（含 ND/RF 行數統計），按缺貨期數降序。"""
    idx = raw_df.index
    # 擴展篩選：任一期缺貨且有需求證據（不限 RF）
    oos_count = _safe_numeric(raw_df.get("OOS Periods Count", _safe_default_series(0, idx)))
    m2_sold = _safe_numeric(raw_df.get("Sold Qty (M-2)", _safe_default_series(0, idx)))
    p90d = _safe_numeric(raw_df.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, idx)))
    has_oos = oos_count.values > 0
    has_demand = (m2_sold.values > 0) | (p90d.values > 0)

    detail = raw_df[has_oos & has_demand].copy()
    if detail.empty:
        ws.cell(row=1, column=1, value="無缺貨影響資料").font = Font(size=12)
        return

    # NDRF Status（顯示 RF 或 ND）
    rp_type = (
        detail.get("RP Type", _safe_default_series("", detail.index))
        .fillna("").astype(str).str.strip().str.upper()
    )
    detail["NDRF Status"] = rp_type

    # 排序：缺貨期數降序 → M-2 銷售降序
    detail["_oos"] = _safe_numeric(detail.get("OOS Periods Count", _safe_default_series(0, detail.index)))
    detail["_m2"] = _safe_numeric(detail.get("Sold Qty (M-2)", _safe_default_series(0, detail.index)))
    detail = detail.sort_values(["_oos", "_m2"], ascending=[False, False])

    avail_cols = [c for c in _OOS_DETAIL_COLS if c in detail.columns]
    detail = detail[avail_cols]

    # 標題
    ws.cell(row=1, column=1, value="OOS Impact Detail — 缺貨銷售影響明細（按缺貨期數降序，含 NDRF Status）").font = Font(bold=True, size=12, color="1E40AF")

    headers = list(detail.columns)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    total_rows = len(detail)
    for row_idx, row_data in enumerate(detail.itertuples(index=False), start=4):
        for col_idx, val in enumerate(row_data, 1):
            if pd.isna(val):
                val = None
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(val, bool):
                cell.value = "✓" if val else ""
            elif isinstance(val, (int, float)):
                cell.number_format = "#,##0"
        if progress_callback and ((row_idx - 4) % 2000 == 0 or row_idx - 3 == total_rows):
            progress_callback(0.1 + 0.15 * ((row_idx - 3) / max(total_rows, 1)),
                              f"寫入缺貨明細 ({row_idx-3}/{total_rows})...")

    _auto_width(ws)


def _write_stock_impact_summary_sheet(ws, summary_df, sheet_title, header_font, header_fill, thin_border, progress_callback, progress_start, progress_end):
    ws.cell(row=1, column=1, value=sheet_title).font = Font(bold=True, size=12, color="1E40AF")

    headers = list(summary_df.columns)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    total_rows = len(summary_df)
    for row_idx, row_data in summary_df.iterrows():
        excel_row = 4 + row_idx
        is_total = any(str(v) == "總計" for v in row_data.values if isinstance(v, str))
        for col_idx, h in enumerate(headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_total:
                cell.font = Font(bold=True, size=11)
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0"
        if progress_callback and (row_idx % 500 == 0 or row_idx == total_rows - 1):
            p = progress_start + (progress_end - progress_start) * ((row_idx + 1) / max(total_rows, 1))
            progress_callback(p, f"寫入 {sheet_title} ({row_idx+1}/{total_rows})...")

    _auto_width(ws)


def _write_sku_allocation_sheet(ws, alloc_df, header_font, header_fill, thin_border, progress_callback):
    ws.cell(row=1, column=1, value="SKU Allocation — 跨 Nature 調貨建議（Tourist 缺貨 vs Non-Tourist/Mix 可支援存貨）").font = Font(bold=True, size=12, color="1E40AF")

    if alloc_df is None or alloc_df.empty:
        ws.cell(row=3, column=1, value="無可調貨 SKU（無 Tourist 缺貨且同時其他 Nature 有存貨的情況）").font = Font(size=11, italic=True, color="6B7280")
        return

    headers = list(alloc_df.columns)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, row_data in enumerate(alloc_df.itertuples(index=False), start=4):
        for col_idx, val in enumerate(row_data, 1):
            if pd.isna(val):
                val = None
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0"

    if progress_callback:
        progress_callback(0.72, f"SKU 調貨表寫入完成（{len(alloc_df)} SKU）")

    _auto_width(ws)


def _write_sku_stock_diagnosis_sheet(ws, diag_df, header_font, header_fill, thin_border, progress_callback):
    """寫入 by-SKU 缺貨診斷表（Tourist+Mix 缺貨 SKU，含 D001 貨量與店舖貨量）。

    D001 狀態分類用顏色標示：「倉庫有貨」(補貨問題) 用橘色警示、「源頭斷貨」用灰色。
    """
    ws.cell(row=1, column=1, value="SKU Stock Diagnosis — by-SKU 缺貨診斷（Tourist + Mix + Non Tourist 缺貨 SKU｜含 D001 倉庫貨量與店舖貨量）").font = Font(bold=True, size=12, color="1E40AF")

    if diag_df is None or diag_df.empty:
        ws.cell(row=3, column=1, value="無 Tourist/Mix 缺貨 SKU").font = Font(size=11, italic=True, color="6B7280")
        return

    headers = list(diag_df.columns)
    # 找出 D001 狀態欄的索引（用於著色）
    d001_status_col_idx = headers.index("D001 狀態") + 1 if "D001 狀態" in headers else None

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    restock_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")  # 淡橘: 補貨問題
    source_fill = PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid")    # 淡灰: 源頭斷貨

    total_rows = len(diag_df)
    for row_idx, row_data in enumerate(diag_df.itertuples(index=False), start=4):
        d001_status = getattr(row_data, "_asdict")().get("D001 狀態") if d001_status_col_idx else None
        for col_idx, val in enumerate(row_data, 1):
            if pd.isna(val):
                val = None
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0"
            # D001 狀態欄著色
            if d001_status_col_idx and col_idx == d001_status_col_idx:
                if val == "倉庫有貨":
                    cell.fill = restock_fill
                    cell.font = Font(bold=True, size=10, color="92400E")
                elif val == "源頭斷貨":
                    cell.fill = source_fill
                    cell.font = Font(size=10, color="4B5563")
        if progress_callback and ((row_idx - 4) % 500 == 0 or row_idx - 3 == total_rows):
            progress_callback(0.741 + 0.018 * ((row_idx - 3) / max(total_rows, 1)),
                              f"寫入 SKU 缺貨診斷 ({row_idx-3}/{total_rows})...")

    if progress_callback:
        progress_callback(0.759, f"SKU 缺貨診斷表寫入完成（{len(diag_df)} SKU）")

    _auto_width(ws)


def _write_inventory_distribution_sheet(ws, inv_distribution, header_font, header_fill, thin_border, progress_callback):
    """寫入存貨分佈分析 sheet（3 個表垂直排列）。

    B1: Mix + Non Tourist Turnday 分級彙總
    B2: Nature 存貨對比
    B3: 存貨失衡 SKU 清單
    """
    ws.cell(row=1, column=1, value="Inventory Distribution — 存貨分佈分析").font = Font(bold=True, size=14, color="1E40AF")
    ws.cell(row=2, column=1, value="重點：高銷售 Tourist 店無貨 vs Non Tourist/Mix 大量囤積；Non Tourist/Mix 只需保留 7-14 日 Turnday").font = Font(size=10, italic=True, color="6B7280")

    current_row = 4

    def _write_subtable(title, df, start_row):
        ws.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=12, color="1E40AF")
        start_row += 1
        if df is None or df.empty:
            ws.cell(row=start_row, column=1, value="無資料").font = Font(size=11, italic=True, color="6B7280")
            return start_row + 2
        headers = list(df.columns)
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=start_row, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, row_data in enumerate(df.itertuples(index=False), start=start_row + 1):
            is_total = any(str(v) == "總計" for v in row_data if isinstance(v, str))
            for col_idx, val in enumerate(row_data, 1):
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if is_total:
                    cell.font = Font(bold=True, size=11)
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
        return start_row + 1 + len(df) + 2  # header + data + spacing

    # B1a: Mix Turnday 分級
    if progress_callback:
        progress_callback(0.77, "寫入 Mix Turnday 分級...")
    current_row = _write_subtable("📊 Mix 店舖 Turnday 分級（標準：保留 7-14 日）", inv_distribution.get("mix_turnday"), current_row)

    # B1b: Non Tourist Turnday 分級
    if progress_callback:
        progress_callback(0.775, "寫入 Non Tourist Turnday 分級...")
    current_row = _write_subtable("📊 Non Tourist 店舖 Turnday 分級", inv_distribution.get("nt_turnday"), current_row)

    # B2: Nature 對比
    if progress_callback:
        progress_callback(0.785, "寫入 Nature 存貨對比...")
    current_row = _write_subtable("🏪 各 Shop Nature 存貨分佈對比", inv_distribution.get("nature_compare"), current_row)

    # B3: 存貨失衡 SKU
    if progress_callback:
        progress_callback(0.79, "寫入存貨失衡 SKU...")
    imb = inv_distribution.get("imbalance_sku")
    current_row = _write_subtable("🔄 存貨失衡 SKU（Tourist 缺貨 + Non Tourist/Mix 過剩 → 建議調貨）", imb, current_row)

    # B3 註解（計算方式說明）
    if imb is not None and hasattr(imb, "attrs") and "annotation" in imb.attrs:
        for line in imb.attrs["annotation"].split("\n"):
            ws.cell(row=current_row, column=1, value=line).font = Font(size=9, italic=True, color="6B7280")
            current_row += 1
        current_row += 1

    # B4: 低存貨 SKU 配置決策
    if progress_callback:
        progress_callback(0.799, "寫入低存貨 SKU 配置決策...")
    low_alloc = inv_distribution.get("low_stock_allocation")
    if low_alloc is not None and not low_alloc.empty:
        # 轉置為 2 欄（指標/數值）方便閱讀
        low_display = low_alloc.T.reset_index()
        low_display.columns = ["指標", "數值"]
        current_row = _write_subtable("⚖️ 低存貨 SKU 配置決策（D001 + Shop Stock < 門檻 → 是否減少 Non Tourist RF 店舖）", low_display, current_row)
    else:
        ws.cell(row=current_row, column=1, value="⚖️ 低存貨 SKU 配置決策：無資料").font = Font(size=12, color="1E40AF")
        current_row += 3

    # B5: 可調貨優先 SKU 清單（Tourist缺貨 + 同SKU在Mix/NT有可釋放件數）
    if progress_callback:
        progress_callback(0.81, "寫入可調貨優先 SKU 清單...")
    tp = inv_distribution.get("transfer_priority", {})
    tp_detail = tp.get("detail") if isinstance(tp, dict) else pd.DataFrame()
    tp_summary = tp.get("summary") if isinstance(tp, dict) else pd.DataFrame()
    if tp_detail is not None and not tp_detail.empty:
        if tp_summary is not None and not tp_summary.empty:
            s = tp_summary.iloc[0]
            summary_text = f"📋 可調貨Tourist行數: {int(s.get('可調貨Tourist行數', 0)):,} 行｜可調貨SKU數: {int(s.get('可調貨SKU數', 0)):,} 個"
        else:
            summary_text = f"📋 可調貨優先 SKU 清單（{len(tp_detail):,} 個 SKU）"
        current_row = _write_subtable(
            f"🔄 可調貨優先 SKU 清單（Tourist缺貨 + 同SKU在Mix/NT有可釋放件數 → 優先調貨）\n{summary_text}",
            tp_detail, current_row,
        )
    else:
        ws.cell(row=current_row, column=1, value="🔄 可調貨優先 SKU 清單：無符合條件 SKU（所有Tourist缺貨SKU在Mix/NT均無可釋放存貨）").font = Font(size=11, color="1E40AF")
        current_row += 3

    _auto_width(ws)


def _write_stock_impact_dashboard(ws, raw_df, summary_buyer, inv_distribution, thin_border, desc_font, progress_callback):
    """重做為「存貨分佈分析」焦點 Dashboard（v1.6.1）。

    聚焦存貨錯配議題：
    - Tourist 高銷售店缺貨 vs Non Tourist/Mix 大量囤積
    - Non Tourist/Mix 只需保留 7-14 日 Turnday
    """
    impacted = raw_df[raw_df["Is Stock Impacted"]].copy()
    nature = raw_df.get("Shop Nature", pd.Series(dtype=object)).fillna("").astype(str)
    cm_stock = _safe_numeric(raw_df.get("Shop Stock Qty (CM)", _safe_default_series(0, raw_df.index)))
    turnday = pd.to_numeric(raw_df.get("Turnday (Days)", _safe_default_series(0, raw_df.index)), errors="coerce")

    # Tourist 缺貨行數（高銷售店無貨）
    tourist_oos_rows = int(((nature == "Tourist") & (cm_stock <= 0) & raw_df["Is Stock Impacted"]).sum()) if "Is Stock Impacted" in raw_df else 0

    # Tourist 缺貨數量（缺失件數 = max(0, Adj. Safety Stock - Shop Stock Qty (CM))）
    _adj_ss = _safe_numeric(raw_df.get("Adj. Safety Stock", _safe_default_series(0, raw_df.index)))
    _oos_mask = (nature == "Tourist") & (cm_stock <= 0) & (raw_df.get("Is Stock Impacted", pd.Series(False, index=raw_df.index)))
    tourist_oos_qty = int(np.maximum(0, (_adj_ss.values - cm_stock.values))[_oos_mask.values].sum())

    # Non Tourist 可釋放存貨件數（>14日標準，與 Mix 一致）
    nt_excess_total = int(_safe_numeric(raw_df.get("Non Tourist Excess Stock", _safe_default_series(0, raw_df.index))).sum())

    # Mix 可釋放存貨件數（>14日標準，來自 Mix Excess Stock）
    mix_excess_total = int(_safe_numeric(raw_df.get("Mix Excess Stock", _safe_default_series(0, raw_df.index))).sum())

    # 存貨失衡 SKU 數（Tourist 缺貨 + 其他 Nature 過剩）
    imbalance_count = len(inv_distribution.get("imbalance_sku", pd.DataFrame())) if inv_distribution else 0

    # 可調貨優先分析（Tourist缺貨 + 同SKU在Mix/NT有可釋放件數）
    tp = inv_distribution.get("transfer_priority", {}) if inv_distribution else {}
    tp_summary = tp.get("summary", pd.DataFrame()) if isinstance(tp, dict) else pd.DataFrame()
    transferrable_tourist_rows = 0
    transferrable_sku_count = 0
    if tp_summary is not None and not tp_summary.empty:
        s = tp_summary.iloc[0]
        transferrable_tourist_rows = int(s.get("可調貨Tourist行數", 0))
        transferrable_sku_count = int(s.get("可調貨SKU數", 0))

    # 各 Nature 平均 Turnday
    tourist_td = turnday[nature == "Tourist"]
    tourist_avg_td = round(float(tourist_td.mean()), 1) if tourist_td.notna().any() and len(tourist_td) else 0
    mix_td = turnday[nature == "Mix"]
    mix_avg_td = round(float(mix_td.mean()), 1) if mix_td.notna().any() and len(mix_td) else 0

    # Non Tourist 平均 Turnday
    nt_td = turnday[nature == "Non Tourist"]
    nt_avg_td = round(float(nt_td.mean()), 1) if nt_td.notna().any() and len(nt_td) else 0

    # 受影響最嚴重 Buyer
    top_buyer_name = "—"
    top_buyer_loss = 0
    if summary_buyer is not None and not summary_buyer.empty and "Buyer" in summary_buyer.columns:
        buyer_data = summary_buyer[~summary_buyer["Buyer"].isin(["總計", ""])]
        if not buyer_data.empty and "Tourist 缺貨行數" in buyer_data.columns:
            top = buyer_data.sort_values("Tourist 缺貨行數", ascending=False).iloc[0]
            top_buyer_name = str(top["Buyer"])
            top_buyer_loss = int(top.get("Tourist 缺貨行數", 0))

    row1_metrics = [
        ("⭐ Tourist 缺貨行數", tourist_oos_rows, "高銷售店無貨（只計 RF）｜單位：行數", "DC2626"),
        ("Non Tourist 可釋存貨", nt_excess_total, "超過14日標準可釋放｜單位：件數", "F59E0B"),
        ("Mix 可釋放存貨", mix_excess_total, "超過14日標準可釋放｜單位：件數", "F97316"),
        ("存貨失衡 SKU 數", imbalance_count, "Tourist缺貨+其他過剩｜單位：SKU數", "EF4444"),
    ]

    row2_metrics = [
        ("Tourist 平均Turnday", tourist_avg_td, "Tourist 平均存貨天數｜單位：日", "8B5CF6"),
        ("Mix 平均Turnday", mix_avg_td, "Mix 平均存貨天數｜單位：日", "3B82F6"),
        ("Non Tourist 平均Turnday", nt_avg_td, "Non Tourist 平均存貨天數｜單位：日", "06B6D4"),
        ("受影響最嚴重 Buyer", top_buyer_name, f"Tourist 缺貨: {top_buyer_loss:,} 行", "1E40AF"),
    ]

    row3_metrics = [
        ("🔄 可調貨Tourist行數", transferrable_tourist_rows, "Tourist缺貨+同SKU在Mix/NT有可釋放件數｜單位：行數", "059669"),
        ("📦 可調貨SKU數", transferrable_sku_count, "涉及的不重複SKU數量｜單位：SKU數", "0D9488"),
        ("📉 Tourist 缺貨數量", tourist_oos_qty, "缺失件數 = Safety − Shop Net Stock｜單位：件數", "DC2626"),
        ("", "", "", "FFFFFF"),
    ]

    _write_metric_cards(ws, "缺貨銷售影響分析摘要 (Stock Impact Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback, row3_metrics=row3_metrics)

    # Top 5 受影響 Brand（按 Tourist 缺貨）
    if summary_buyer is not None and not summary_buyer.empty and len(summary_buyer) > 1:
        ws.cell(row=14, column=1, value="Top 5 受缺貨影響最嚴重的 Buyer（按 Tourist 缺貨排序）").font = Font(bold=True, size=12, color="1E40AF")
        top_buyers = summary_buyer[summary_buyer["Buyer"] != "總計"].head(5)
        for col_idx, h in enumerate(top_buyers.columns, 1):
            cell = ws.cell(row=15, column=col_idx, value=h)
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, (_, row_data) in enumerate(top_buyers.iterrows()):
            excel_row = 16 + row_idx
            for col_idx, h in enumerate(top_buyers.columns, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"

    _auto_width(ws)


def generate_excel_stock_impact(
    summary_brand: pd.DataFrame,
    summary_buyer: pd.DataFrame,
    summary_category: pd.DataFrame,
    summary_item_class: pd.DataFrame,
    sku_diagnosis: pd.DataFrame,
    inv_distribution: dict,
    raw_df: pd.DataFrame,
    progress_callback=None,
    sku_monthly_analysis: pd.DataFrame | None = None,
) -> BytesIO:
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    desc_font = Font(size=9, color="6B7280", italic=True)

    # Sheet 1: RAW（只輸出分析相關欄位以節省記憶體）
    ws_raw = wb.active
    ws_raw.title = "RAW"
    _write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.08, columns=_OOS_DETAIL_COLS)

    if progress_callback:
        progress_callback(0.1, "建立缺貨影響明細工作表...")

    # Sheet 2: OOS Impact Detail
    ws_detail = wb.create_sheet(title="OOS Impact Detail")
    _write_oos_detail_sheet(ws_detail, raw_df, header_font, header_fill, thin_border, progress_callback)

    # Sheet 3: By Brand
    if progress_callback:
        progress_callback(0.3, "建立 By Brand 分析工作表...")
    ws_brand = wb.create_sheet(title="By Brand")
    if not summary_brand.empty:
        _write_stock_impact_summary_sheet(ws_brand, summary_brand, "By Brand — 缺貨影響分析",
                                          header_font, header_fill, thin_border, progress_callback, 0.30, 0.42)
    else:
        ws_brand.cell(row=1, column=1, value="無 Brand 層級缺貨影響數據").font = Font(size=12)

    # Sheet 4: By Buyer
    if progress_callback:
        progress_callback(0.44, "建立 By Buyer 分析工作表...")
    ws_buyer = wb.create_sheet(title="By Buyer")
    if not summary_buyer.empty:
        _write_stock_impact_summary_sheet(ws_buyer, summary_buyer, "By Buyer — 缺貨影響分析",
                                          header_font, header_fill, thin_border, progress_callback, 0.44, 0.56)
    else:
        ws_buyer.cell(row=1, column=1, value="無 Buyer 層級缺貨影響數據").font = Font(size=12)

    # Sheet 5: By Category
    if progress_callback:
        progress_callback(0.58, "建立 By Category 分析工作表...")
    ws_category = wb.create_sheet(title="By Category")
    if not summary_category.empty:
        _write_stock_impact_summary_sheet(ws_category, summary_category, "By Category (Adj. MCH1) — 缺貨影響分析",
                                          header_font, header_fill, thin_border, progress_callback, 0.58, 0.68)
    else:
        ws_category.cell(row=1, column=1, value="無 Category (Adj. MCH1) 層級缺貨影響數據").font = Font(size=12)

    # Sheet 6: By Item Class
    if progress_callback:
        progress_callback(0.68, "建立 By Item Class 分析工作表...")
    ws_item_class = wb.create_sheet(title="By Item Class")
    if not summary_item_class.empty:
        _write_stock_impact_summary_sheet(ws_item_class, summary_item_class, "By Item Class — 缺貨影響分析",
                                          header_font, header_fill, thin_border, progress_callback, 0.68, 0.70)
    else:
        ws_item_class.cell(row=1, column=1, value="無 Item Class 層級缺貨影響數據").font = Font(size=12)

    # Sheet 6b: SKU Monthly Sales & Storage (按 SKU 月銷量與存儲量)
    if progress_callback:
        progress_callback(0.69, "建立 SKU 月銷量和存儲量工作表...")
    ws_sku_monthly = wb.create_sheet(title="SKU Monthly Sales&Store")
    if sku_monthly_analysis is not None and not sku_monthly_analysis.empty:
        ws_sku_monthly.cell(row=1, column=1, value="SKU 月銷量和存儲量分析（按 Shop Nature）").font = Font(bold=True, size=12, color="1E40AF")
        ws_sku_monthly.cell(row=2, column=1, value="各 SKU 在 Tourist/Mix/Non Tourist 店舖的月均銷售（P90D/3）與當期庫存（CM Shop Stock）。低存貨 SKU 優先顯示。").font = Font(size=10, italic=True, color="6B7280")
        headers = list(sku_monthly_analysis.columns)
        for col_idx, h in enumerate(headers, 1):
            cell = ws_sku_monthly.cell(row=4, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, row_data in enumerate(sku_monthly_analysis.itertuples(index=False), start=5):
            for col_idx, val in enumerate(row_data, 1):
                if pd.isna(val):
                    val = None
                cell = ws_sku_monthly.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, bool):
                    cell.value = "✓ 低存貨" if val else ""
                elif isinstance(val, (int, float)) and val is not None:
                    cell.number_format = "#,##0"
            if progress_callback and (row_idx % 500 == 0 or row_idx - 4 == len(sku_monthly_analysis)):
                progress_callback(0.69 + 0.01 * ((row_idx - 4) / max(len(sku_monthly_analysis), 1)), f"寫入 SKU 月銷量 ({row_idx-4}/{len(sku_monthly_analysis)})...")
        _auto_width(ws_sku_monthly)
    else:
        ws_sku_monthly.cell(row=1, column=1, value="無 SKU 月銷量/存儲量數據").font = Font(size=12)

    # Sheet 7: SKU Stock Diagnosis (by-SKU 缺貨診斷, 含 D001 + 店舖貨量)
    if progress_callback:
        progress_callback(0.70, "建立 by-SKU 缺貨診斷工作表...")
    ws_diag = wb.create_sheet(title="SKU Stock Diagnosis")
    _write_sku_stock_diagnosis_sheet(ws_diag, sku_diagnosis, header_font, header_fill, thin_border, progress_callback)

    # Sheet 8: Inventory Distribution (存貨分佈分析)
    if progress_callback:
        progress_callback(0.76, "建立存貨分佈分析工作表...")
    ws_inv = wb.create_sheet(title="Inventory Distribution")
    _write_inventory_distribution_sheet(ws_inv, inv_distribution, header_font, header_fill, thin_border, progress_callback)

    # Sheet 9: Dashboard
    if progress_callback:
        progress_callback(0.80, "建立缺貨影響分析摘要工作表...")
    ws_dashboard = wb.create_sheet(title="Stock Impact Dashboard")
    _write_stock_impact_dashboard(ws_dashboard, raw_df, summary_buyer, inv_distribution,
                                  thin_border, desc_font, progress_callback)

    if progress_callback:
        progress_callback(0.97, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf




# ─── Stock Impact Analysis (Simplified) Functions ─────────────────────────────
# 簡化版本：隻輸出 RAW 和 SKU Monthly Sales&Store 兩個工作表
# SKU Monthly Sales&Store 不分 Shop Nature，隻顯示 M-2, M-1, MTD 的 D001 存貨、店舖存貨和總銷售量

def _build_sku_monthly_analysis_simplified(raw: pd.DataFrame, stock_threshold: int = 300) -> pd.DataFrame:
    """簡化版 SKU 月銷量和存儲量分析（保留原版所有欄位，新增不分 Nature 的週期存貨/銷售欄位）。

    保留原版 `_build_sku_monthly_analysis` 的所有欄位（含 Tourist/Mix/Non Tourist 分拆），
    再附加 M-2, M-1 的 D001 存貨量、店舖總存貨量和店舖總銷售量（不分 Shop Nature）。
    MTD 期的 D001 和店舖存貨已存在於原版 D001_庫存/總店舖庫存，不再重複。
    """
    if raw.empty or "SKU" not in raw.columns:
        return pd.DataFrame()

    # 先取得原版完整輸出（含 Nature 分拆欄位）
    base = _build_sku_monthly_analysis(raw, stock_threshold=stock_threshold)
    if base.empty:
        return base

    # 合併 raw 的補充欄位（每 SKU 取第一筆）
    _extra_cols = ["Vendor MC", "Adj. MCH1", "Item Status", "Launch Date", "Item Class"]
    _present = [c for c in _extra_cols if c in raw.columns]
    if _present:
        col_first = raw.groupby("SKU", dropna=False)[_present].first().reset_index()
        base = base.merge(col_first, on="SKU", how="left")

    # Is New SKU (3mo) — 原版 _add_launch_info 已產生
    has_new_sku = "Is New SKU (3mo)" in raw.columns and "Is New SKU (3mo)" not in base.columns
    if has_new_sku:
        new_sku_col = raw.groupby("SKU", dropna=False)["Is New SKU (3mo)"].first().reset_index()
        base = base.merge(new_sku_col, on="SKU", how="left")

    # 計算不分 Nature 的週期欄位（只算 M-2, M-1；MTD 已存在於 D001_庫存/總店舖庫存）
    d001_m2 = _safe_numeric(raw["D001 Stock Qty (M-2)"]) if "D001 Stock Qty (M-2)" in raw.columns else _safe_default_series(0, raw.index)
    d001_m1 = _safe_numeric(raw["D001 Stock Qty (M-1)"]) if "D001 Stock Qty (M-1)" in raw.columns else _safe_default_series(0, raw.index)
    shop_m2 = _safe_numeric(raw.get("Shop Stock Qty (M-2)", _safe_default_series(0, raw.index)))
    shop_m1 = _safe_numeric(raw.get("Shop Stock Qty (M-1)", _safe_default_series(0, raw.index)))
    sold_m2 = _safe_numeric(raw.get("Sold Qty (M-2)", _safe_default_series(0, raw.index)))
    sold_m1 = _safe_numeric(raw.get("Sold Qty (M-1)", _safe_default_series(0, raw.index)))
    sold_mtd = _safe_numeric(raw.get("Sold Qty (CM - MTD)", _safe_default_series(0, raw.index)))

    extra_data = pd.DataFrame({
        "SKU": raw["SKU"].fillna("").astype(str).values,
        "_d001_m2": d001_m2.values,
        "_d001_m1": d001_m1.values,
        "_shop_m2": shop_m2.values,
        "_shop_m1": shop_m1.values,
        "_sold_m2": sold_m2.values,
        "_sold_m1": sold_m1.values,
        "_sold_mtd": sold_mtd.values,
    })

    extra_agg = extra_data.groupby("SKU", dropna=False).agg(
        D001_M2=("_d001_m2", "first"),
        D001_M1=("_d001_m1", "first"),
        ShopStock_M2=("_shop_m2", "sum"),
        ShopStock_M1=("_shop_m1", "sum"),
        Sold_M2=("_sold_m2", "sum"),
        Sold_M1=("_sold_m1", "sum"),
        Sold_MTD=("_sold_mtd", "sum"),
    ).reset_index()

    # 合併
    result = base.merge(extra_agg, on="SKU", how="left")

    # 轉整數
    for c in ["D001_M2", "D001_M1",
              "ShopStock_M2", "ShopStock_M1",
              "Sold_M2", "Sold_M1", "Sold_MTD"]:
        if c in result.columns:
            result[c] = result[c].fillna(0).astype(int)

    # 欄位順序
    _nature_cols = []
    for nat in ["Tourist", "Mix", "Non Tourist"]:
        _nature_cols += [
            f"{nat}_月均銷售", f"{nat}_當期CM銷售", f"{nat}_當期庫存",
            f"{nat}_店數", f"{nat}_平均Turnday", f"{nat}_RF月均銷售",
        ]

    _extra_info_cols = ["Vendor MC", "Adj. MCH1", "Item Status", "Launch Date", "Item Class", "Is New SKU (3mo)"]
    _extra_info_cols = [c for c in _extra_info_cols if c in result.columns]

    _new_period_cols = ["D001_M2", "D001_M1", "ShopStock_M2", "ShopStock_M1", "Sold_M2", "Sold_M1", "Sold_MTD"]
    _new_period_cols = [c for c in _new_period_cols if c in result.columns]

    out_cols = ["SKU", "SKU Desc", "Brand", "Buyer"]
    out_cols += _extra_info_cols
    out_cols += ["低存貨", "總可用存貨", "D001_庫存", "總店舖庫存", "總行數", "RF行數", "ND行數"]
    out_cols += _nature_cols
    out_cols += _new_period_cols
    out_cols = [c for c in out_cols if c in result.columns]

    return result[out_cols]


def generate_excel_stock_impact_simplified(
    raw_df: pd.DataFrame,
    progress_callback=None,
    sku_monthly_analysis: pd.DataFrame | None = None,
) -> BytesIO:
    """簡化版 Excel 輸出：隻輸出 RAW 和 SKU Monthly Sales&Store 兩個工作表。

    SKU Monthly Sales&Store 保留原版所有欄位，並新增 M-2, M-1, MTD 的
    D001 存貨量、店舖總存貨量及店舖總銷售量（不分 Shop Nature）。
    """
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    # Sheet 1: RAW
    ws_raw = wb.active
    ws_raw.title = "RAW"
    _write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.3, columns=_OOS_DETAIL_COLS)

    if progress_callback:
        progress_callback(0.4, "建立 SKU Monthly Sales&Store 工作表...")

    # Sheet 2: SKU Monthly Sales&Store
    ws_sku_monthly = wb.create_sheet(title="SKU Monthly Sales&Store")
    if sku_monthly_analysis is not None and not sku_monthly_analysis.empty:
        ws_sku_monthly.cell(row=1, column=1, value="SKU 月銷量和存儲量分析（簡化版，含原版欄位 + 新增 M-2/M-1 不分 Nature 存貨/銷售）").font = Font(bold=True, size=12, color="1E40AF")
        ws_sku_monthly.cell(row=2, column=1, value="保留原版 Tourist/Mix/Non Tourist 所有欄位及新增資訊欄位，再新增 M-2, M-1 的 D001 存貨量、店舖總存貨量及 M-2, M-1, MTD 的店舖總銷售量（不分 Shop Nature）。低存貨 SKU 優先顯示。").font = Font(size=10, italic=True, color="6B7280")
        headers = list(sku_monthly_analysis.columns)
        for col_idx, h in enumerate(headers, 1):
            cell = ws_sku_monthly.cell(row=4, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, row_data in enumerate(sku_monthly_analysis.itertuples(index=False), start=5):
            for col_idx, val in enumerate(row_data, 1):
                if pd.isna(val):
                    val = None
                cell = ws_sku_monthly.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, bool):
                    cell.value = "✓ 低存貨" if val else ""
                elif isinstance(val, (int, float)) and val is not None:
                    cell.number_format = "#,##0"
            if progress_callback and (row_idx % 500 == 0 or row_idx - 4 == len(sku_monthly_analysis)):
                progress_callback(0.5 + 0.3 * ((row_idx - 4) / max(len(sku_monthly_analysis), 1)), f"寫入 SKU 月銷量 ({row_idx-4}/{len(sku_monthly_analysis)})...")
        _auto_width(ws_sku_monthly)
    else:
        ws_sku_monthly.cell(row=1, column=1, value="無 SKU 月銷量/存儲量數據").font = Font(size=12)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf
