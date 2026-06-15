import pandas as pd
import numpy as np
import unicodedata
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

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
    "Combined Safety Stock",
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

    cond_overstock_2x = (status_col == "Overstock (>2*Monthly Sold Qty & >2pcs)")

    cond_overstock_0 = (
        (status_col == "Overstock (P90D 0 Sales)")
        & (df.get("RP Type", _safe_default_series("", df.index)).isin(["ND", "RF"]))
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
        rp = df.get("RP Type", _safe_default_series("", df.index))
        is_rf = cond_overstock_0 & (rp == "RF")
        is_nd = cond_overstock_0 & (rp == "ND")
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
        x = np.floor(monthly_sold_arr * rate_arr).clip(lower=3).astype(int)
        over_adj = x > adj_ss_arr
        x[over_adj] = adj_ss_arr[over_adj].astype(int)
        return x

    _os_default = rates_overstock.get("default", 0.50)
    _os_tourist = rates_overstock.get("tourist", 0.30)
    _os_mix = rates_overstock.get("mix", 0.40)

    os_rate_arr = np.full(len(df), _os_default)
    os_rate_arr[shop_nature.str.contains("tourist", na=False)] = _os_tourist
    os_rate_arr[shop_nature.str.contains("mix", na=False)] = _os_mix

    x_values = np.full(len(df), 0, dtype=int)
    if cond_2x.any():
        x_values = _compute_x(monthly_sold.values, os_rate_arr, adj_ss.values)

    # OOS floor Y
    if mtd_days > 0:
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

    y_values = np.ceil(projected.values * oos_rate_arr).clip(lower=3).astype(int)

    # Z = current Adj. Safety Stock
    z_values = adj_ss.values.astype(int)

    # Combined = max(Y, Z if Z <= X else X)
    combined = np.where(z_values <= x_values, np.maximum(y_values, z_values), np.maximum(y_values, x_values))

    df["Overstock Cap (X)"] = x_values
    df["OOS Floor (Y)"] = y_values
    df["Combined Safety Stock"] = combined

    delta = combined - z_values
    if safety_mode == "increase_only":
        delta = delta.clip(lower=0)
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


def _write_raw_sheet(ws, raw_df, progress_callback=None, progress_start=0.0, progress_end=0.3):
    raw_df = _compute_raw_extra_columns(raw_df)
    raw_df = _reorder_raw_columns(raw_df)

    raw_cols = list(raw_df.columns)
    _display_names = {"Adj. Safety Stock": "Safety Stock (現時)"}
    header_row = [_display_names.get(c, c) for c in raw_cols]
    ws.append(header_row)

    raw_array = raw_df.to_numpy(dtype=object, na_value=None)
    total_raw = len(raw_array)
    _batch = []
    _batch_size = 500
    for i, row_vals in enumerate(raw_array):
        _batch.append([
            round(float(v), 1) if isinstance(v, (int, float)) and not isinstance(v, bool) else v
            for v in row_vals
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


def _write_metric_cards(ws, title, row1_metrics, row2_metrics, thin_border, desc_font, progress_callback=None):
    ws.merge_cells("A1:D1")
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(bold=True, size=14, color="1E40AF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_offset, metrics in [(3, row1_metrics), (7, row2_metrics)]:
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
        progress_callback(0.95, "調整格式中...")
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
        ("總 Safety Stock 審查筆數", total_combined, "所有 Safety Stock 審查之總計筆數", "EF4444"),
        ("影響採購人員數 (Buyers)", total_buyers, "位採購人員名下有需審查品項", "3B82F6"),
        ("影響品牌總數 (Brands)", total_brands, "個品牌包含需審查商品", "8B5CF6"),
        ("影響SKU總數", total_skus, "個包含需審查商品(SKUs)", "10B981"),
    ]

    row2_metrics = [
        ("Overstock 分類筆數", cat_overstock, "庫存過賸分類筆數", "F59E0B"),
        ("OOS 分類筆數", cat_oos, "缺貨分類筆數", "EC4899"),
        ("Safety 調整筆數", cat_balanced, "Balanced (Safety Adjusted) 筆數", "06B6D4"),
        ("Safety 調整總量", safety_change_total, "Safety Qty Change 總和", "8B5CF6"),
    ]

    _write_metric_cards(ws_dashboard, "Safety Stock 數據分析摘要 (Safety Stock Analysis Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback)

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
        "Overstock (P90D 0 Sales)",
        "Overstock (>2*Monthly Sold Qty & >2pcs)",
    ]

    group_cols = ["Buyer", "Brand", "SKU", "SKU Desc"]

    if progress_callback:
        progress_callback(0.2, "\u5206\u985e\u7d71\u8a08\u4e2d...")

    counts = {}
    for i, cat in enumerate(categories):
        mask = (status_col == cat)
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
        ("總 Overstock 筆數", total_overstock, "所有庫存異常之總計筆數", "EF4444"),
        ("影響採購人員數 (Buyers)", total_buyers, "位採購人員名下有異常品項", "3B82F6"),
        ("影響品牌總數 (Brands)", total_brands, "個品牌包含異常庫存商品", "8B5CF6"),
        ("影響SKU總數", total_skus, "個包含異常庫存商品(SKUs)", "10B981"),
    ]

    row2_metrics = [
        ("Overstock (P90D 0 Sales)", cat1_val, "「過去 90 天內零銷售」之異常筆數", "F59E0B"),
        ("Overstock (>2*Monthly Sold Qty & >2pcs)", cat2_val, "「大於月均銷售量兩倍且庫存大於 2 件」之異常筆數", "EC4899"),
        ("New Safety Qty / ND 調整後的預計總數", transfer_qty_total, "After new Safety Qty / ND can transfer out of TOTAL QTY", "06B6D4"),
    ]

    desc_font = Font(size=9, color="6B7280", italic=True)
    _write_metric_cards(ws_dashboard, "異常庫存數據分析摘要 (Analysis Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "\u5b8c\u6210!")

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
        rec_rf = np.ceil(projected[is_rf] * rate_arr[is_rf]).clip(lower=3).astype(int)
        recommended[is_rf] = rec_rf

    if is_other.any():
        rec_other = np.ceil(projected[is_other] * rate_arr[is_other]).clip(lower=3).astype(int)
        ad_other = adj_ss_arr[is_other]
        over_adj = rec_other > ad_other
        rec_other[over_adj] = ad_other[over_adj]
        recommended[is_other] = rec_other

    df["Recommendation of new Safety Qty"] = recommended

    increase = recommended - adj_ss_arr
    increase[is_nd] = 0
    if safety_mode == "increase_only":
        increase[is_rf] = increase[is_rf].clip(lower=0)
        increase[is_other] = increase[is_other].clip(lower=0)
    else:
        increase[is_other] = increase[is_other].clip(lower=0)
    df["Safety Qty Increase"] = increase

    if safety_mode == "increase_only":
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
            建議SafetyQty增加筆數=("_is_normal_insufficient", "sum"),
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
                "建議SafetyQty增加筆數": [int(by_shop["建議SafetyQty增加筆數"].sum())],
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
        "建議SafetyQty增加筆數",
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
        ("總 OOS 筆數", total_oos, "所有 OOS 異常之總計筆數", "EF4444"),
        ("影響採購人員數 (Buyers)", total_buyers, "位採購人員名下有 OOS 品項", "3B82F6"),
        ("影響品牌總數 (Brands)", total_brands, "個品牌包含 OOS 商品", "8B5CF6"),
        ("影響SKU總數", total_skus, "個包含 OOS 商品(SKUs)", "10B981"),
    ]

    row2_metrics = [
        ("OOS (0 Stock)", cat1_val, "現有庫存為 0 之 OOS 筆數", "F59E0B"),
        ("Potential OOS (<7 Turnday)", cat2_val, "Turnday 不足 7 之潛在 OOS 筆數", "EC4899"),
        ("Normal (Safety Qty Insufficient)", cat3_val, "正常庫存但安全庫存不足筆數", "06B6D4"),
        ("建議 Safety Qty 修改", safety_increase_total, "建議 Safety Qty 修改 之總量", "8B5CF6"),
    ]

    _write_metric_cards(ws_dashboard, "OOS庫存數據分析摘要 (OOS Analysis Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback)

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
    df["_status2"] = np.where(status == CARRY_P90D0_STATUS, "Overstock (P90D 0 Sales)", "Other Status")

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
        for s in ["Overstock (P90D 0 Sales)", "Other Status"]:
            col = f"{r}_{s}"
            if col not in pivot_wide.columns:
                pivot_wide[col] = 0

    pivot_wide = pivot_wide.rename(columns={
        "ND_Overstock (P90D 0 Sales)": "ND_P90D0",
        "ND_Other Status": "ND_Other",
        "RF_Overstock (P90D 0 Sales)": "RF_P90D0",
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
                 4: "Overstock (P90D 0 Sales)", 5: "Other Status", 6: "ND 合計",
                 7: "Overstock (P90D 0 Sales)", 8: "Other Status", 9: "RF 合計", 10: "總計SKU"}
    row6_right = {12: "Site Code", 13: "Shop Name", 14: "Adj. SST",
                  15: "Overstock (P90D 0 Sales)", 16: "Other Status", 17: "ND 合計",
                  18: "Overstock (P90D 0 Sales)", 19: "Other Status", 20: "RF 合計", 21: "總計SKU"}
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

