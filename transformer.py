import pandas as pd
import numpy as np
import unicodedata
import re
import gc
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
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


COMBINED_RAW_COLS = [
    "Brand",
    "SKU",
    "SKU Desc",
    "Buyer",
    "Adj. SST",
    "Adj. MCH1",
    "Vendor MC",
    "Item Status",
    "Launch Date",
    "Item Class",
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "RP Type",
    "Adj. Safety Stock",  # 表頭顯示為 Safety Stock (現時)
    "Recommend Safety Stock",
    "Recommend Safety Stock Turnday",
    "Safety Qty Change",
    "Safety Qty Change 分類",
    "Recommendation of new Safety Qty / ND",
    "Buyer reply",
    "Shop Monthly Sold Qty (P90D)",
    "Safety VS Monthly Sold Qty",
    "Projected Monthly Sold Qty",
    "Overstock Cap (X)",
    "OOS Floor (Y)",
    "D001 Net Stock Qty (CM)",
    "Shop Stock Qty (CM)",
    "Sold Qty (P90D)",
    "Turnday (P90D Qty)",
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Supply Source",
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


def _nature_rate_arr(shop_nature, default_rate, tourist_rate, mix_rate):
    # Shop Nature → 每列 rate。必須精確比對：'Non Tourist' 含 'tourist' 子字串，
    # 用 contains 會把 Non Tourist 誤套 Tourist 倍率；其餘值（如 'Other'）歸 default。
    arr = np.full(len(shop_nature), default_rate)
    arr[shop_nature == "tourist"] = tourist_rate
    arr[shop_nature == "mix"] = mix_rate
    return arr


def _parse_launch_dates(launch_raw: pd.Series) -> pd.Series:
    """Launch Date 原始字串 → datetime（自動推斷 + 常見格式重試，無法解析為 NaT）。"""
    launch_dt = pd.to_datetime(launch_raw, errors="coerce")
    if launch_dt.isna().any():
        still_na = launch_dt.isna()
        for fmt in ("%Y/%m/%d", "%d/%m/%Y", "%Y%m%d"):
            if not still_na.any():
                break
            retry = pd.to_datetime(launch_raw[still_na], errors="coerce", format=fmt)
            launch_dt[still_na] = retry
            still_na = launch_dt.isna()
    return launch_dt


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
    # 必須限制緩存大小與壽命 — 無 TTL/無 max_entries 時每次轉換都累積一份 DataFrame 拷貝，
    # 長時間運行的容器記憶體會持續增長且「釋放記憶體」按鈕前的運行間隔內無法自動回收
    _compute_raw_extra_columns_impl = st.cache_data(_compute_raw_extra_columns_impl, max_entries=4, ttl=3600)
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
                "P90D 0 Sales SKU數量": [int(df["_sku_0_sales"].dropna().nunique())],
                ">2*Monthly Sold Qty & >2pcs數量": [int(df["_sku_2x"].dropna().nunique())],
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
        rate = _nature_rate_arr(shop_nature, _default_rate, _tourist_rate, _mix_rate)
        # Shop Monthly Sold Qty (P90D) 已是月均（P90D 總量 ÷ 3），直接乘 rate
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
            # 精確比對：'Non Tourist' 含 'tourist' 子字串，contains 會把 Non Tourist 誤當 Tourist
            mask |= sn_col == "tourist"
        if "Mix" in shop_natures:
            mask |= sn_col == "mix"
        if "Non Tourist" in shop_natures:
            # 非 Tourist/Mix 的其他值（如 'Other'）隨 Non Tourist 一併保留，
            # 與倍率邏輯一致（非 tourist/mix 一律套 default 倍率）
            is_tourist_or_mix = (sn_col == "tourist") | (sn_col == "mix")
            mask |= (sn_col != "") & ~is_tourist_or_mix
        df = df[mask]
    return df


# ─── Safety Stock Review Functions ─────────────────────────────────────────────


def _build_nd_rec_qty_nd(is_nd, is_p90d_0, n):
    arr = np.full(n, "", dtype=object)
    if is_p90d_0.any():
        arr[is_p90d_0.values] = "ND"
    if is_nd.any():
        arr[is_nd.values] = "0"
    return arr


def _annotate_nd_recommendation(nd_rec, df):
    """為「ND」轉換建議依 Adj. MCH1 / Adj. SST 加註類別提示：
    Adj. MCH1 = 02 - MAKE UP → ND(02 - MAKE UP)；
    Adj. SST 為 03 - Private Label 系（(1)Major / (2)Health）→ ND(Private Label)；
    兩者皆是 → ND(Private Label - MAKE UP)。
    只加註「ND」轉換建議；「0」列（本來已是 ND）與空白列不動。
    """
    rec = np.asarray(nd_rec, dtype=object)
    is_nd_suggest = rec == "ND"
    if not is_nd_suggest.any():
        return rec
    mch1 = (
        df.get("Adj. MCH1", _safe_default_series("", df.index))
        .fillna("").astype(str).str.strip().str.lower()
    )
    sst = (
        df.get("Adj. SST", _safe_default_series("", df.index))
        .fillna("").astype(str).str.strip().str.lower()
    )
    is_makeup = (mch1 == "02 - make up").values
    is_pl = sst.str.startswith("03 - private label").values
    rec = rec.copy()
    rec[is_nd_suggest & is_pl & is_makeup] = "ND(Private Label - MAKE UP)"
    rec[is_nd_suggest & is_pl & ~is_makeup] = "ND(Private Label)"
    rec[is_nd_suggest & ~is_pl & is_makeup] = "ND(02 - MAKE UP)"
    return rec


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

    # Overstock cap X — Shop Monthly Sold Qty (P90D) 已是月均，直接乘 rate
    def _compute_x(monthly_sold_arr, rate_arr, adj_ss_arr):
        x = np.floor(monthly_sold_arr * rate_arr).clip(min=3).astype(int)
        over_adj = x > adj_ss_arr
        x[over_adj] = adj_ss_arr[over_adj].astype(int)
        return x

    _os_default = rates_overstock.get("default", 0.50)
    _os_tourist = rates_overstock.get("tourist", 0.30)
    _os_mix = rates_overstock.get("mix", 0.40)

    os_rate_arr = _nature_rate_arr(shop_nature, _os_default, _os_tourist, _os_mix)

    # X is the upper bound for "how much safety stock is reasonable"; it applies to
    # every row (so the combined recommendation = max(Y, Z if Z <= X else X) is well
    # defined per row), independent of whether any row currently has Overstock status.
    x_values = _compute_x(monthly_sold.values, os_rate_arr, adj_ss.values)

    # OOS floor Y — Projected 月銷售採加權混合：
    #   w = mtd_days / month_total_days
    #   Projected = w × MTD線性外推 + (1−w) × P90D月均 = sold_qty + (1−w) × P90D月均
    # （月初 w 小，以穩定的 P90D 為主，避免頭幾日外推失真；月尾收斂至 MTD 實績）
    # Shop Monthly Sold Qty (P90D) 已是月均（Past 90 Days ÷ 90 × 30），直接使用；
    # 欄位整個缺失或個別列缺值（NaN）時，該列退回純 MTD 線性外推。
    if mtd_days and mtd_days > 0:
        w = mtd_days / month_total_days
        linear = (sold_qty / mtd_days) * month_total_days
        if "Shop Monthly Sold Qty (P90D)" in df.columns:
            p90d_raw = pd.to_numeric(df["Shop Monthly Sold Qty (P90D)"], errors="coerce")
            projected = sold_qty + (1.0 - w) * p90d_raw.fillna(0.0)
            projected = projected.where(p90d_raw.notna(), linear)
        else:
            projected = linear
    else:
        projected = _safe_default_series(0.0, df.index)

    df["Projected Monthly Sold Qty"] = np.ceil(projected).astype(int)

    _oos_default = rates_oos.get("default", 0.50)
    _oos_tourist = rates_oos.get("tourist", 0.50)
    _oos_mix = rates_oos.get("mix", 0.50)

    oos_rate_arr = _nature_rate_arr(shop_nature, _oos_default, _oos_tourist, _oos_mix)

    y_values = np.ceil(projected.values * oos_rate_arr).clip(min=3).astype(int)

    # Z = current Adj. Safety Stock
    z_values = adj_ss.values.astype(int)

    # Combined = max(Y, Z if Z <= X else X)
    combined = np.where(z_values <= x_values, np.maximum(y_values, z_values), np.maximum(y_values, x_values))

    df["Overstock Cap (X)"] = x_values
    df["OOS Floor (Y)"] = y_values

    delta = combined - z_values
    if safety_mode == "increase_only":
        # 只可增加：delta 鉗制 ≥ 0，一般調整列只增不減（保底）
        delta = delta.clip(min=0)
        recommend_safety = z_values + delta
    else:
        recommend_safety = combined
    # Override: ND items must have Safety Stock = 0 (ND = no replenishment)
    # Override: Overstock (P90D 0 Sales) items must have Safety Stock = 0 and be recommended ND
    rp_type = (
        df.get("RP Type", _safe_default_series("", df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    is_nd = rp_type == "ND"
    is_p90d_0 = status_col == "Overstock (P90D 0 Sales)"

    # 新品保護：Launch Date 距今 90 天內（≈3 個月）視為 New SKU。新品未出單是上架
    # 時間問題而非滯銷證據，排除「P90D 0 Sales → Safety=0/ND」override，改走正常
    # X/Y 計算；Launch Date 缺欄位、缺值或無法解析時一律視為非新品（維持原行為）。
    if "Launch Date" in df.columns:
        _launch_raw = df["Launch Date"].fillna("").astype(str).str.strip()
        _launch_dt = _parse_launch_dates(_launch_raw)
        is_new_sku = (
            _launch_dt.notna()
            & (_launch_dt >= pd.Timestamp(datetime.date.today()) - pd.Timedelta(days=90))
        ).values
    else:
        is_new_sku = np.zeros(len(df), dtype=bool)
    is_p90d_0_override = is_p90d_0 & ~is_new_sku

    if safety_mode != "increase_only":
        # 「可加可減」才套用退場 override —「只可增加」模式下任何列都不出現扣減：
        # ND / P90D 0 Sales 列保持原有 Safety，退場決策留待可加可減模式
        if is_nd.any() | is_p90d_0_override.any():
            recommend_safety = recommend_safety.copy()
            delta = delta.copy()
        if is_nd.any():
            recommend_safety[is_nd.values] = 0
            delta[is_nd.values] = -z_values[is_nd.values]
        if is_p90d_0_override.any():
            recommend_safety[is_p90d_0_override.values] = 0
            delta[is_p90d_0_override.values] = -z_values[is_p90d_0_override.values]

    df["Recommend Safety Stock"] = recommend_safety
    df["Safety Qty Change"] = delta

    # Safety Qty Change 分類（增加 / 減少 / 不變）—「只可增加」模式下不會出現「減少」
    df["Safety Qty Change 分類"] = np.where(delta > 0, "增加", np.where(delta < 0, "減少", "不變"))

    # 建議 Safety Stock 的周轉日數，與來源「Turnday (P90D Qty)」同口徑（P90D 日均）：
    # = Recommend Safety Stock ÷ (Shop Monthly Sold Qty (P90D) ÷ 30)。
    # P90D 月均 = 0 時無周轉意義（來源為 inf），此處留空；Recommend = 0 顯示 0（零覆蓋）。
    rec_turnday = np.full(len(df), np.nan)
    _monthly_arr = monthly_sold.values
    _monthly_pos = _monthly_arr > 0
    rec_turnday[_monthly_pos] = recommend_safety[_monthly_pos] / (_monthly_arr[_monthly_pos] / 30.0)
    df["Recommend Safety Stock Turnday"] = np.round(rec_turnday, 1)

    if safety_mode == "increase_only":
        # 只可增加：無歸零 / ND 退場建議，ND 欄僅保留 New SKU 資訊標記
        nd_rec = np.full(len(df), "", dtype=object)
        if is_new_sku.any():
            nd_rec[is_new_sku] = "New SKU (3 Month)"
    else:
        nd_rec = _build_nd_rec_qty_nd(is_nd, is_p90d_0_override, len(df))
        # New SKU 標記：ND 列維持「0」（ND 規則不受新品保護影響），其餘新品列標記
        if is_new_sku.any():
            nd_rec = nd_rec.copy()
            nd_rec[is_new_sku & ~is_nd.values] = "New SKU (3 Month)"
        # ND 轉換建議依 Adj. MCH1 / Adj. SST 加註 MAKE UP / Private Label
        nd_rec = _annotate_nd_recommendation(nd_rec, df)
    df["Recommendation of new Safety Qty / ND"] = nd_rec
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
            Overstock_SKU數量=("_sku_overstock", "nunique"),
            OOS_SKU數量=("_sku_oos", "nunique"),
            需增加Safety_SKU數量=("_safety_increase", "sum"),
            需減少Safety_SKU數量=("_safety_decrease", "sum"),
            安全庫存調整總量=("_safety_change", "sum"),
            Combined_SKU總數=("SKU", "nunique"),
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
                "Overstock_SKU數量": [int(df["_sku_overstock"].dropna().nunique())],
                "OOS_SKU數量": [int(df["_sku_oos"].dropna().nunique())],
                "需增加Safety_SKU數量": [int(by_shop["需增加Safety_SKU數量"].sum())],
                "需減少Safety_SKU數量": [int(by_shop["需減少Safety_SKU數量"].sum())],
                "安全庫存調整總量": [int(by_shop["安全庫存調整總量"].sum())],
                "Combined_SKU總數": [int(df["SKU"].dropna().nunique())],
            }
        )
        by_shop = pd.concat([by_shop.sort_values(["Site Code", "Shop Name", "OM"], kind="stable"), total_row], ignore_index=True)

    return by_shop


# ── SUMMARY sheet：兩大分類統計的固定顯示順序與說明 ──────────────────────────

_CHANGE_CAT_DESC = {
    "增加": "建議 Safety Stock 高於現時",
    "減少": "建議 Safety Stock 低於現時（僅「可加可減」模式出現）",
    "不變": "建議 Safety Stock 與現時相同",
}

_ND_REC_DISPLAY = {
    "": "(空白) 維持 Safety Stock",
    "New SKU (3 Month)": "New SKU (3 Month)",
    "0": "0（已是 ND）",
    "ND": "ND",
    "ND(02 - MAKE UP)": "ND(02 - MAKE UP)",
    "ND(Private Label)": "ND(Private Label)",
    "ND(Private Label - MAKE UP)": "ND(Private Label - MAKE UP)",
}

_ND_REC_DESC = {
    "": "無 ND 建議，按建議 Safety Stock 繼續運作",
    "New SKU (3 Month)": "上市 90 天內新品，保留觀察、不套用歸零/ND",
    "0": "RP Type 已是 ND，建議 Safety = 0",
    "ND": "建議轉 ND 停補貨（P90D 0 Sales 舊品）",
    "ND(02 - MAKE UP)": "建議轉 ND — Adj. MCH1 = 02 - MAKE UP",
    "ND(Private Label)": "建議轉 ND — Adj. SST 屬 03 - Private Label 系",
    "ND(Private Label - MAKE UP)": "建議轉 ND — MAKE UP + Private Label",
}


def _build_combined_classification_summary(raw):
    """SUMMARY sheet 資料 — 兩張分類統計表：
    ① Safety Qty Change 分類（增加 / 減少 / 不變）
    ② Recommendation of new Safety Qty / ND 分類（空白 / New SKU / 0 / ND 系列）
    每列統計 行數 與 SKU 數；同一 SKU 可跨分類出現（不同店舖），總計列的
    SKU 數改用全表不重複 SKU 數，避免跨分類加總虛增。
    """
    sku = raw["SKU"] if "SKU" in raw.columns else _safe_default_series("", raw.index)
    change = pd.to_numeric(raw.get("Safety Qty Change", _safe_default_series(0, raw.index)), errors="coerce").fillna(0)
    change_cat = np.where(change > 0, "增加", np.where(change < 0, "減少", "不變"))
    total_sku = int(sku.dropna().nunique())

    change_rows = []
    for label in ("增加", "減少", "不變"):
        m = change_cat == label
        change_rows.append({
            "分類": label,
            "行數": int(m.sum()),
            "SKU 數": int(sku[m].dropna().nunique()),
            "Safety Qty Change 合計": int(change[m].sum()),
            "說明": _CHANGE_CAT_DESC[label],
        })
    change_rows.append({
        "分類": "總計",
        "行數": int(len(raw)),
        "SKU 數": total_sku,
        "Safety Qty Change 合計": int(change.sum()),
        "說明": "SKU 數為全表不重複 SKU（同一 SKU 可跨分類）",
    })
    change_summary = pd.DataFrame(change_rows)

    nd_rec = (
        raw.get("Recommendation of new Safety Qty / ND", _safe_default_series("", raw.index))
        .fillna("").astype(str)
    )
    known = list(_ND_REC_DISPLAY.keys())
    present_order = known + sorted(v for v in nd_rec.unique() if v not in known)
    nd_rows = []
    for val in present_order:
        m = (nd_rec == val).values
        nd_rows.append({
            "建議分類": _ND_REC_DISPLAY.get(val, val),
            "行數": int(m.sum()),
            "SKU 數": int(sku[m].dropna().nunique()),
            "說明": _ND_REC_DESC.get(val, ""),
        })
    nd_rows.append({
        "建議分類": "總計",
        "行數": int(len(raw)),
        "SKU 數": total_sku,
        "說明": "SKU 數為全表不重複 SKU（同一 SKU 可跨分類）",
    })
    nd_summary = pd.DataFrame(nd_rows)

    return change_summary, nd_summary


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

    # 預計算 int 旗標（單次 groupby + 條件聚合，取代 7 次獨立 groupby）
    raw["_is_overstock"] = overstock_mask.astype(int)
    raw["_is_oos"] = oos_mask.astype(int)
    raw["_is_balanced"] = balanced_mask.astype(int)
    raw["_is_inc"] = (safety_change > 0).astype(int)
    raw["_is_dec"] = (safety_change < 0).astype(int)
    raw["_is_unchanged"] = (safety_change == 0).astype(int)

    if progress_callback:
        progress_callback(0.5, "合併彙總中...")

    summary = (
        raw.groupby(group_cols, observed=True)
        .agg(
            Overstock=("_is_overstock", "sum"),
            OOS=("_is_oos", "sum"),
            **{"Balanced (Safety Adjusted)": ("_is_balanced", "sum")},
            **{"Safety Qty Change": ("Safety Qty Change", "sum")},
            **{"增加Safety Qty的店舖數量": ("_is_inc", "sum")},
            **{"減少Safety Qty的店舖數量": ("_is_dec", "sum")},
            **{"Safety Qty不變的店舖數量": ("_is_unchanged", "sum")},
        )
        .fillna(0).astype(int).reset_index()
    )
    summary["SKU"] = summary["SKU"].apply(_safe_sku)
    summary["總計"] = summary[categories].sum(axis=1)

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

    # 清理暫存欄位
    raw.drop(columns=[c for c in ["_is_overstock", "_is_oos", "_is_balanced", "_is_inc", "_is_dec", "_is_unchanged"] if c in raw.columns], inplace=True, errors="ignore")

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


def _write_metric_cards(ws, title, row1_metrics, row2_metrics, thin_border, desc_font, progress_callback=None, progress_pct=0.95, row3_metrics=None, row4_metrics=None, row5_metrics=None):
    ws.merge_cells("A1:D1")
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(bold=True, size=14, color="1E40AF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    row_offsets = [(3, row1_metrics), (7, row2_metrics)]
    if row3_metrics:
        row_offsets.append((11, row3_metrics))
    if row4_metrics:
        row_offsets.append((15, row4_metrics))
    if row5_metrics:
        row_offsets.append((19, row5_metrics))

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


def _write_summary_sheet(ws, change_summary, nd_summary, thin_border):
    """SUMMARY sheet — 兩個區塊：Safety Qty Change 分類 與 ND 建議分類 統計表."""
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    section_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    ws.merge_cells("A1:E1")
    title_cell = ws.cell(row=1, column=1, value="Safety Stock Review SUMMARY（分類統計）")
    title_cell.font = Font(bold=True, size=14, color="1E40AF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    def _write_block(start_row, section_title, block):
        cols = list(block.columns)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=len(cols))
        sc = ws.cell(row=start_row, column=1, value=section_title)
        sc.font = Font(bold=True, size=12, color="1E40AF")
        sc.fill = section_fill
        sc.alignment = Alignment(horizontal="left", vertical="center")

        header_row = start_row + 1
        for col_idx, h in enumerate(cols, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        for r, row_data in enumerate(block.itertuples(index=False), header_row + 1):
            is_total = row_data[0] == "總計"
            for col_idx, v in enumerate(row_data, 1):
                cell = ws.cell(row=r, column=col_idx, value=v)
                cell.border = thin_border
                if is_total:
                    cell.font = bold_font
                    cell.fill = total_fill
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 1:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        return header_row + len(block) + 1

    next_row = _write_block(3, "① Safety Qty Change 分類（增加 / 減少 / 不變）", change_summary)
    _write_block(next_row + 1, "② Recommendation of new Safety Qty / ND 分類", nd_summary)
    _auto_width(ws)


# ─── 店舖問題 SKU（每間店舖需行動列：問題＋建議方法）────────────────────────

# 三張新表欄序（新契約：固定表頭，未經明確要求不得改動）
_SHOP_ISSUE_BASE_COLUMNS = [
    "Site Code", "Shop Name", "OM", "Shop Nature",
    "Brand", "SKU", "SKU Desc", "Buyer",
    "問題類型", "問題說明", "建議方法",
]
_STOCK_CHECKING_SHOP_ISSUE_NUM_COLUMNS = [
    "Supply Source", "D001 Net Stock Qty (CM)",
    "Shop Stock Qty (CM)", "Shop Monthly Sold Qty (P90D)",
    "調入需求量 (補到14日)", "實際調入量", "調入缺口",
    "可轉出貨量 (Mix)", "可轉出貨量 (Non Tourist)", "轉出量",
    "調貨來源 (Site Code)",
]
_COMBINED_SHOP_ISSUE_NUM_COLUMNS = [
    "Shop Stock Qty (CM)", "Shop Monthly Sold Qty (P90D)", "Turnday (P90D Qty)",
    "Adj. Safety Stock", "Recommend Safety Stock", "Safety Qty Change",
    "Overstock Cap (X)", "OOS Floor (Y)",
    "Recommendation of new Safety Qty / ND",
]
_HYBRID_SHOP_ISSUE_NUM_COLUMNS = [
    "Supply Source", "D001 Net Stock Qty (CM)",
    "Shop Stock Qty (CM)", "Shop Monthly Sold Qty (P90D)",
    "Safety 基準天數", "Safety 周轉所需量",
    "Adj. Safety Stock", "建議 Safety 設定值", "Safety 建議調整量",
    "調入需求量 (補到14日)", "實際調入量", "調入缺口",
    "可轉出貨量", "轉出量", "調貨來源 (Site Code)",
    "實際調入量 (同OM)", "調入缺口 (同OM)", "轉出量 (同OM)",
]
_STOCK_CHECKING_SHOP_ISSUE_COLUMNS = _SHOP_ISSUE_BASE_COLUMNS + _STOCK_CHECKING_SHOP_ISSUE_NUM_COLUMNS
_COMBINED_SHOP_ISSUE_COLUMNS = _SHOP_ISSUE_BASE_COLUMNS + _COMBINED_SHOP_ISSUE_NUM_COLUMNS
_HYBRID_SHOP_ISSUE_COLUMNS = _SHOP_ISSUE_BASE_COLUMNS + _HYBRID_SHOP_ISSUE_NUM_COLUMNS

# 店內問題嚴重度（數字越小越優先；缺口最緊急）
_SHOP_ISSUE_SEVERITY = {
    "調入缺口（來源不足）": 0,
    "需要調入": 1,
    "Safety 設定不足": 1,
    "Safety 設定過高": 2,
    "可轉出供應": 3,
    "ND 退場": 3,
}


def _shop_name_lookup(raw: pd.DataFrame) -> dict[str, str]:
    """Site Code → Shop Name 查找表（RAW 已由 Item_Shop 衍生 Shop Name；缺欄時回退解析）。"""
    lookup: dict[str, str] = {}
    if raw is None or raw.empty or "Site Code" not in raw.columns:
        return lookup
    site = raw["Site Code"].fillna("").astype(str).str.strip()
    if "Shop Name" in raw.columns:
        names = raw["Shop Name"].fillna("").astype(str).str.strip()
    elif "Item_Shop" in raw.columns:
        names = pd.Series(
            [_extract_shop_name(raw["Item_Shop"].iloc[i], site.iloc[i]) for i in range(len(raw))],
            index=raw.index,
        )
    else:
        return lookup
    for site_v, name_v in zip(site.to_numpy(), names.to_numpy()):
        if site_v and name_v and site_v not in lookup:
            lookup[site_v] = str(name_v)
    return lookup


def _shop_issue_base_frame(raw: pd.DataFrame, mask: pd.Series) -> pd.DataFrame:
    """新表共用底：篩出需行動列並帶出識別欄（Site Code／Shop Name／OM／Nature／品項）。"""
    sel = raw.loc[mask].copy()
    if sel.empty:
        return sel
    if "Site Code" not in sel.columns:
        sel["Site Code"] = ""
    site = sel["Site Code"].fillna("").astype(str).str.strip()
    sel["Site Code"] = site
    names = _shop_name_lookup(raw)
    if "Shop Name" in sel.columns:
        own = sel["Shop Name"].fillna("").astype(str).str.strip()
        sel["Shop Name"] = [o if o else names.get(s, "") for o, s in zip(own.to_numpy(), site.to_numpy())]
    else:
        sel["Shop Name"] = site.map(names).fillna("")
    if "OM" not in sel.columns:
        sel["OM"] = ""
    else:
        sel["OM"] = sel["OM"].fillna("").astype(str).str.strip()
    if "Shop Nature" not in sel.columns:
        sel["Shop Nature"] = ""
    return sel


def _count_transfer_sources(src: str) -> int:
    """計算調貨來源店舖數量：「HB30:2; HB49:3」→ 2；空白／無法解析 → 0。"""
    if not src:
        return 0
    count = 0
    for part in str(src).split(";"):
        part = part.strip()
        if not part:
            continue
        supply_site, sep, _qty_text = part.rpartition(":")
        if sep and supply_site.strip():
            count += 1
    return count


def _d001_available_qty(raw: pd.DataFrame) -> pd.Series:
    """店舖問題表共用：逐列實際採用的 D001 數量。

    優先序與觸發判定一致 — Stock Checking 用 Net→Stock→0；Hybrid 用逐列
    Net→Stock→0 的 Effective 值（缺欄時退回 Net→0）。非數值視為 0。
    """
    idx = raw.index
    if "D001 Effective Stock Qty (CM)" in raw.columns:
        return pd.to_numeric(raw["D001 Effective Stock Qty (CM)"], errors="coerce").fillna(0.0)
    net_col = _resolve_column_name("D001 Net Stock Qty (CM)", set(raw.columns))
    stock_col = _resolve_column_name("D001 Stock Qty (CM)", set(raw.columns))
    net = pd.to_numeric(raw[net_col], errors="coerce") if net_col else _safe_default_series(np.nan, idx)
    stock = pd.to_numeric(raw[stock_col], errors="coerce") if stock_col else _safe_default_series(np.nan, idx)
    return net.where(net.notna(), stock).fillna(0.0)


def _is_warehouse_supply(raw: pd.DataFrame) -> pd.Series:
    """店舖問題表共用：Supply Source 是否為 Warehouse（大小寫與前後空白不敏感）。

    只有 Warehouse 由 D001 供貨；Vendor 由供應商直送店舖，D001 有無貨與其無關；
    缺欄或空白一律視為非 Warehouse（保守：不作 D001 相關建議）。
    """
    if "Supply Source" not in raw.columns:
        return pd.Series(False, index=raw.index)
    return raw["Supply Source"].fillna("").astype(str).str.strip().str.lower() == "warehouse"


def _d001_gap_advice(d001_qty: float, gap_n: int, is_warehouse: bool) -> str:
    """店舖問題表共用：調入缺口的跟進行動（只在缺口 > 0 時呼叫）。

    - Warehouse＋D001 有貨：可由 D001 補貨。
    - Warehouse＋D001 無貨：D001 亦無貨，請向上游追貨。
    - Vendor：不經 D001，請向供應商追貨。
    """
    if not is_warehouse:
        return "Vendor 貨不經 D001 請向供應商追貨"
    if d001_qty > 0:
        return f"D001 尚有 {d001_qty:g} 件可補貨，缺口部分請由 D001 補貨"
    return "D001 亦無貨，缺口部分請向上游追貨"


def _build_stock_checking_shop_issues(
    raw: pd.DataFrame,
    turnday_threshold: int = 14,
    d001_threshold: int = 100,
) -> pd.DataFrame:
    """Stock Checking 店舖問題表：只收需行動列（需要調入／可轉出供應）。

    - 需要調入：`需要調貨 == 是`；調入缺口 > 0 時問題類型記為「調入缺口（來源不足）」。
    - 可轉出供應：可轉出貨量或轉出量 > 0 且非需求列。
    - 問題說明帶 SKU 級 Tourist 加總 TurnDay 與 D001 觸發原因；建議方法帶數量與行動指引。
    - 排序：Site Code → 嚴重度 → Brand／SKU；門檻僅用於問題說明文字，不重算判定。
    """
    cols = _STOCK_CHECKING_SHOP_ISSUE_COLUMNS
    if raw is None or raw.empty or "需要調貨" not in raw.columns:
        return pd.DataFrame(columns=cols)
    need = raw["需要調貨"].fillna("").astype(str).str.strip() == "是"
    mix_avail = _safe_numeric(raw.get("可轉出貨量 (Mix)", _safe_default_series(0, raw.index)))
    nt_avail = _safe_numeric(raw.get("可轉出貨量 (Non Tourist)", _safe_default_series(0, raw.index)))
    donated = _safe_numeric(raw.get("轉出量", _safe_default_series(0, raw.index)))
    gap = _safe_numeric(raw.get("調入缺口", _safe_default_series(0, raw.index)))
    supply = (~need) & ((mix_avail > 0) | (nt_avail > 0) | (donated > 0))
    mask = need | supply
    if not bool(mask.any()):
        return pd.DataFrame(columns=cols)
    work = _shop_issue_base_frame(raw, mask)
    if "Supply Source" not in work.columns:
        work["Supply Source"] = raw.get("Supply Source", _safe_default_series("", work.index))
    if "D001 Net Stock Qty (CM)" not in work.columns:
        net_col = _resolve_column_name("D001 Net Stock Qty (CM)", set(raw.columns))
        if net_col is not None:
            work["D001 Net Stock Qty (CM)"] = raw[net_col]
    need_w = work["需要調貨"].fillna("").astype(str).str.strip() == "是"
    gap_w = _safe_numeric(work.get("調入缺口", _safe_default_series(0, work.index)))
    ttd_w = pd.to_numeric(work.get("Tourist 加總 TurnDay", _safe_default_series(np.nan, work.index)), errors="coerce")
    d001_qty_w = _d001_available_qty(raw).loc[work.index]
    d001_w = work.get("D001 Net Stock Qty (CM)", _safe_default_series("", work.index))
    warehouse_w = _is_warehouse_supply(raw).loc[work.index]
    demand_w = _safe_numeric(work.get("調入需求量 (補到14日)", _safe_default_series(0, work.index))).astype(int)
    actual_w = _safe_numeric(work.get("實際調入量", _safe_default_series(0, work.index))).astype(int)
    src_w = work.get("調貨來源 (Site Code)", _safe_default_series("", work.index)).fillna("").astype(str).str.strip()
    mix_w = _safe_numeric(work.get("可轉出貨量 (Mix)", _safe_default_series(0, work.index))).astype(int)
    nt_w = _safe_numeric(work.get("可轉出貨量 (Non Tourist)", _safe_default_series(0, work.index))).astype(int)
    out_w = _safe_numeric(work.get("轉出量", _safe_default_series(0, work.index))).astype(int)

    issue, desc, advice, sev = [], [], [], []
    for i in range(len(work)):
        if bool(need_w.iloc[i]):
            ttd_txt = f"{float(ttd_w.iloc[i]):.1f} 日" if pd.notna(ttd_w.iloc[i]) else "未能計算"
            d001_txt = str(d001_w.iloc[i]).strip() or "未能確認"
            gap_n = int(gap_w.iloc[i])
            if gap_n > 0:
                issue.append("調入缺口（來源不足）")
                desc.append(
                    f"Tourist 加總 TurnDay {ttd_txt}低於門檻 {int(turnday_threshold)} 日，"
                    f"D001 僅 {d001_txt} 件（門檻 {int(d001_threshold)} 件），需調入卻無足額來源"
                )
                src_n = _count_transfer_sources(str(src_w.iloc[i]))
                advice.append(
                    f"建議調入 {int(demand_w.iloc[i])} 件（實際可調入 {int(actual_w.iloc[i])} 件，"
                    f"缺口 {gap_n} 件，涉及 {src_n} 間來源店舖）；"
                    f"{_d001_gap_advice(float(d001_qty_w.iloc[i]), gap_n, bool(warehouse_w.iloc[i]))}"
                )
                sev.append(_SHOP_ISSUE_SEVERITY["調入缺口（來源不足）"])
            else:
                issue.append("需要調入")
                desc.append(
                    f"Tourist 加總 TurnDay {ttd_txt}低於門檻 {int(turnday_threshold)} 日，"
                    f"D001 僅 {d001_txt} 件（門檻 {int(d001_threshold)} 件），店舖庫存不足 14 日周轉"
                )
                src_n = _count_transfer_sources(str(src_w.iloc[i]))
                advice.append(
                    f"建議調入 {int(demand_w.iloc[i])} 件（實際可調入 {int(actual_w.iloc[i])} 件，"
                    f"涉及 {src_n} 間來源店舖），詳見「調貨來源 (Site Code)」欄"
                )
                sev.append(_SHOP_ISSUE_SEVERITY["需要調入"])
        else:
            issue.append("可轉出供應")
            desc.append("店舖庫存超過 14 日周轉，可向缺貨店舖供應")
            advice.append(
                f"可轉出 Non Tourist {int(nt_w.iloc[i])} 件、Mix {int(mix_w.iloc[i])} 件"
                f"（實際轉出 {int(out_w.iloc[i])} 件），詳見「調貨來源 (Site Code)」欄"
            )
            sev.append(_SHOP_ISSUE_SEVERITY["可轉出供應"])
    work["問題類型"] = issue
    work["問題說明"] = desc
    work["建議方法"] = advice
    work["_sev"] = sev
    if "Brand" not in work.columns:
        work["Brand"] = ""
    if "SKU" not in work.columns:
        work["SKU"] = ""
    work = work.sort_values(["Site Code", "_sev", "Brand", "SKU"], kind="stable").drop(columns=["_sev"])
    return work[[c for c in cols if c in work.columns]].reset_index(drop=True)


def _build_combined_shop_issues(raw: pd.DataFrame) -> pd.DataFrame:
    """Safety Stock Review 店舖問題表：只收 Safety 需調整與 ND 退場列。

    - Safety 需增加／減少：`Safety Qty Change != 0`。
    - ND 退場：`Recommendation of new Safety Qty / ND` 為 ND 系列或 0（已是 ND）；
      `New SKU (3 Month)` 屬資訊標記，不收錄。
    - Overstock／OOS 只作問題說明上下文，不單獨成列。
    """
    cols = _COMBINED_SHOP_ISSUE_COLUMNS
    if raw is None or raw.empty:
        return pd.DataFrame(columns=cols)
    chg = pd.to_numeric(raw.get("Safety Qty Change", _safe_default_series(np.nan, raw.index)), errors="coerce").fillna(0)
    nd_rec = raw.get("Recommendation of new Safety Qty / ND", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    is_new_sku = nd_rec == "New SKU (3 Month)"
    is_nd_exit = nd_rec.str.startswith("ND") | (nd_rec == "0")
    mask = ((chg != 0) | (is_nd_exit & ~is_new_sku)).fillna(False)
    if not bool(mask.any()):
        return pd.DataFrame(columns=cols)
    work = _shop_issue_base_frame(raw, mask)
    chg_w = pd.to_numeric(work.get("Safety Qty Change", _safe_default_series(0, work.index)), errors="coerce").fillna(0)
    nd_w = work.get("Recommendation of new Safety Qty / ND", _safe_default_series("", work.index)).fillna("").astype(str).str.strip()
    adj_w = pd.to_numeric(work.get("Adj. Safety Stock", _safe_default_series(np.nan, work.index)), errors="coerce")
    rec_w = pd.to_numeric(work.get("Recommend Safety Stock", _safe_default_series(np.nan, work.index)), errors="coerce")
    status_w = work.get("Shop Stock Status", _safe_default_series("", work.index)).fillna("").astype(str).str.strip()

    issue, desc, advice, sev = [], [], [], []
    for i in range(len(work)):
        delta = float(chg_w.iloc[i])
        rec_txt = nd_w.iloc[i]
        adj_txt = f"{float(adj_w.iloc[i]):g}" if pd.notna(adj_w.iloc[i]) else "未能確認"
        rec_val_txt = f"{float(rec_w.iloc[i]):g}" if pd.notna(rec_w.iloc[i]) else "未能確認"
        status_txt = str(status_w.iloc[i]).strip()
        status_suffix = f"，狀態為{status_txt}" if status_txt else ""
        if delta > 0:
            issue.append("Safety 設定不足")
            desc.append(f"現時 Safety {adj_txt} 件低於建議 {rec_val_txt} 件（需增加 {delta:g} 件）{status_suffix}")
            advice.append(f"建議將 Safety 由 {adj_txt} 件調整為 {rec_val_txt} 件（＋{delta:g}），請按建議值補足安全庫存")
            sev.append(_SHOP_ISSUE_SEVERITY["Safety 設定不足"])
        elif delta < 0:
            issue.append("Safety 設定過高")
            desc.append(f"現時 Safety {adj_txt} 件高於建議 {rec_val_txt} 件（可減少 {abs(delta):g} 件）{status_suffix}")
            advice.append(f"建議將 Safety 由 {adj_txt} 件調整為 {rec_val_txt} 件（{delta:g}），請按建議值調低安全庫存")
            sev.append(_SHOP_ISSUE_SEVERITY["Safety 設定過高"])
        elif rec_txt.startswith("ND"):
            issue.append("ND 退場")
            desc.append(f"過去 90 日零銷售，建議轉為 ND 停補貨{status_suffix}（建議：{rec_txt}）")
            advice.append(f"建議轉為 {rec_txt}，現有庫存全數轉出或退倉（現時 Safety {adj_txt} 件）")
            sev.append(_SHOP_ISSUE_SEVERITY["ND 退場"])
        else:
            issue.append("ND 退場")
            desc.append(f"RP Type 已是 ND，建議 Safety 維持 0 件{status_suffix}")
            advice.append("已是 ND，無須補貨；現有庫存請按退倉建議處理")
            sev.append(_SHOP_ISSUE_SEVERITY["ND 退場"])
    work["問題類型"] = issue
    work["問題說明"] = desc
    work["建議方法"] = advice
    work["_sev"] = sev
    if "Brand" not in work.columns:
        work["Brand"] = ""
    if "SKU" not in work.columns:
        work["SKU"] = ""
    work = work.sort_values(["Site Code", "_sev", "Brand", "SKU"], kind="stable").drop(columns=["_sev"])
    return work[[c for c in cols if c in work.columns]].reset_index(drop=True)


def _build_hybrid_shop_issues(raw: pd.DataFrame) -> pd.DataFrame:
    """Safety Stock x NetStock Review 店舖問題表：只收 Safety 設定不足／過高列。

    - 母體與 Safety Review × 分類維度交叉表相同（事件類型非空），但設定正常列一律排除。
    - 主表以不限OM事件為準，另附三個同OM對照欄；調入缺口 > 0 時在建議方法加註。
    """
    cols = _HYBRID_SHOP_ISSUE_COLUMNS
    if raw is None or raw.empty or "事件類型" not in raw.columns:
        return pd.DataFrame(columns=cols)
    diag = raw.get("Safety Review", _safe_default_series("", raw.index)).fillna("").astype(str)
    mask = diag.isin([_HYBRID_DIAG_UNDER, _HYBRID_DIAG_OVER])
    if not bool(mask.any()):
        return pd.DataFrame(columns=cols)
    work = _shop_issue_base_frame(raw, mask)
    if "Supply Source" not in work.columns:
        work["Supply Source"] = raw.get("Supply Source", _safe_default_series("", work.index))
    if "D001 Net Stock Qty (CM)" not in work.columns:
        net_col = _resolve_column_name("D001 Net Stock Qty (CM)", set(raw.columns))
        if net_col is not None:
            work["D001 Net Stock Qty (CM)"] = raw[net_col]
    diag_w = work.get("Safety Review", _safe_default_series("", work.index)).fillna("").astype(str)
    days_w = pd.to_numeric(work.get(_HYBRID_SAFETY_DAYS_COLUMN, _safe_default_series(np.nan, work.index)), errors="coerce")
    line_w = pd.to_numeric(work.get(_HYBRID_SAFETY_LINE_COLUMN, _safe_default_series(np.nan, work.index)), errors="coerce")
    adj_w = pd.to_numeric(work.get("Adj. Safety Stock", _safe_default_series(np.nan, work.index)), errors="coerce")
    sugg_w = pd.to_numeric(work.get("建議 Safety 設定值", _safe_default_series(np.nan, work.index)), errors="coerce")
    chg_w = pd.to_numeric(work.get("Safety 建議調整量", _safe_default_series(np.nan, work.index)), errors="coerce")
    evt_w = work.get("事件類型", _safe_default_series("", work.index)).fillna("").astype(str)
    gap_w = _safe_numeric(work.get("調入缺口", _safe_default_series(0, work.index))).astype(int)
    actual_w = _safe_numeric(work.get("實際調入量", _safe_default_series(0, work.index))).astype(int)
    demand_w = _safe_numeric(work.get("調入需求量 (補到14日)", _safe_default_series(0, work.index))).astype(int)
    out_w = _safe_numeric(work.get("轉出量", _safe_default_series(0, work.index))).astype(int)
    src_w = work.get("調貨來源 (Site Code)", _safe_default_series("", work.index)).fillna("").astype(str).str.strip()
    d001_qty_w = _d001_available_qty(raw).loc[work.index]
    warehouse_w = _is_warehouse_supply(raw).loc[work.index]

    issue, desc, advice, sev = [], [], [], []
    for i in range(len(work)):
        days_txt = f"{float(days_w.iloc[i]):g} 日" if pd.notna(days_w.iloc[i]) else "未能確認"
        line_txt = f"{float(line_w.iloc[i]):g}" if pd.notna(line_w.iloc[i]) else "未能確認"
        adj_txt = f"{float(adj_w.iloc[i]):g}" if pd.notna(adj_w.iloc[i]) else "未能確認"
        sugg_txt = f"{float(sugg_w.iloc[i]):g}" if pd.notna(sugg_w.iloc[i]) else "未能確認"
        delta_txt = f"{float(chg_w.iloc[i]):g}" if pd.notna(chg_w.iloc[i]) else "未能確認"
        evt_txt = str(evt_w.iloc[i]).strip()
        if diag_w.iloc[i] == _HYBRID_DIAG_UNDER:
            issue.append("Safety 設定不足")
            desc.append(f"現時 Safety {adj_txt} 件低於周轉所需 {line_txt} 件（基準 {days_txt}；事件：{evt_txt}）")
            text = f"建議將 Safety 由 {adj_txt} 件調整為 {sugg_txt} 件（{delta_txt}）"
            if int(demand_w.iloc[i]) > 0:
                src_n = _count_transfer_sources(str(src_w.iloc[i]))
                text += f"；另需調入 {int(demand_w.iloc[i])} 件（實際可調入 {int(actual_w.iloc[i])} 件，涉及 {src_n} 間來源店舖）"
            if int(gap_w.iloc[i]) > 0:
                text += (f"；尚有缺口 {int(gap_w.iloc[i])} 件，"
                         f"{_d001_gap_advice(float(d001_qty_w.iloc[i]), int(gap_w.iloc[i]), bool(warehouse_w.iloc[i]))}")
            advice.append(text)
            sev.append(_SHOP_ISSUE_SEVERITY["Safety 設定不足"])
        else:
            issue.append("Safety 設定過高")
            desc.append(f"現時 Safety {adj_txt} 件高於周轉所需 {line_txt} 件（基準 {days_txt}；事件：{evt_txt}）")
            advice.append(
                f"建議將 Safety 由 {adj_txt} 件調整為 {sugg_txt} 件（{delta_txt}），"
                f"多餘存貨可轉出（實際轉出 {int(out_w.iloc[i])} 件），詳見「調貨來源 (Site Code)」欄"
            )
            sev.append(_SHOP_ISSUE_SEVERITY["Safety 設定過高"])
    work["問題類型"] = issue
    work["問題說明"] = desc
    work["建議方法"] = advice
    work["_sev"] = sev
    avail_w = _safe_numeric(work.get("可轉出貨量 (Mix)", _safe_default_series(0, work.index)))
    avail_nt = _safe_numeric(work.get("可轉出貨量 (Non Tourist)", _safe_default_series(0, work.index)))
    avail_t = _safe_numeric(work.get("可轉出貨量 (Tourist)", _safe_default_series(0, work.index)))
    work["可轉出貨量"] = (avail_w + avail_nt + avail_t).astype(int)
    if "Brand" not in work.columns:
        work["Brand"] = ""
    if "SKU" not in work.columns:
        work["SKU"] = ""
    work = work.sort_values(["Site Code", "_sev", "Brand", "SKU"], kind="stable").drop(columns=["_sev"])
    return work[[c for c in cols if c in work.columns]].reset_index(drop=True)


def _write_shop_issue_sheet(ws, issue_df, *, title, progress_callback=None, progress_start=0.0, progress_end=0.0):
    """店舖問題 SKU 工作表：店舖×SKU 逐列（僅含需行動列），問題類型置中、說明／建議靠左換行。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    band_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    _text_keys = ["Brand", "SKU Desc", "Buyer", "問題說明", "建議方法", "調貨來源 (Site Code)"]

    if issue_df is None or issue_df.empty:
        ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12, color="1E40AF")
        ws.cell(row=3, column=1, value="無符合資料").font = Font(size=12)
        _auto_width(ws)
        return
    headers = list(issue_df.columns)
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12, color="1E40AF")
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[3].height = 30
    total = len(issue_df)
    for r_idx, (_, row) in enumerate(issue_df.iterrows()):
        excel_row = r_idx + 4
        is_band = (r_idx % 2 == 1)
        for col_idx, h in enumerate(headers, 1):
            val = row.get(h)
            cell = ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)
            cell.border = thin_border
            cell.alignment = left_wrap if h in _text_keys else center
            if isinstance(val, (int, float)) and not isinstance(val, bool) and h not in _text_keys:
                cell.number_format = "#,##0"
            if is_band:
                cell.fill = band_fill
        if progress_callback and (r_idx % 200 == 0 or r_idx == total - 1):
            p = progress_start + (progress_end - progress_start) * ((r_idx + 1) / max(total, 1))
            progress_callback(p, f"寫入店舖問題 SKU ({r_idx+1}/{total})...")
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{3 + total}"
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

    _write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.15, columns=COMBINED_RAW_COLS)

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
        "Overstock_SKU數量",
        "OOS_SKU數量",
        "需增加Safety_SKU數量",
        "需減少Safety_SKU數量",
        "安全庫存調整總量",
        "Combined_SKU總數",
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
        progress_callback(0.85, "建立店舖問題 SKU 工作表...")

    # ── Sheet 4: 店舖問題 SKU（每間店舖需行動列；空表寫「無符合資料」）──
    ws_shop_issue = wb.create_sheet(title="店舖問題 SKU")
    _write_shop_issue_sheet(
        ws_shop_issue, _build_combined_shop_issues(raw_df),
        title="店舖問題 SKU — 每間店舖需行動列（問題＋建議方法）",
        progress_callback=progress_callback, progress_start=0.85, progress_end=0.86,
    )

    if progress_callback:
        progress_callback(0.86, "建立 SUMMARY 工作表...")

    # ── Sheet 5: SUMMARY（Safety Qty Change 分類 / ND 建議分類 統計）──
    change_summary, nd_summary = _build_combined_classification_summary(raw_df)
    ws_class_summary = wb.create_sheet(title="SUMMARY")
    _write_summary_sheet(ws_class_summary, change_summary, nd_summary, thin_border)

    if progress_callback:
        progress_callback(0.88, "建立分析摘要工作表...")

    # ── Sheet 6: Safety Stock 數據分析摘要 ──
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

    # 預計算 int 旗標（單次 groupby 取代多次 filter+groupby + transfer merge）
    _agg_dict = {}
    for cat in categories:
        mask = (status_col == _STATUS_VALUE.get(cat, cat))
        key = f"_cat_{abs(hash(cat)) % 10_000_000}"
        df[key] = mask.astype(int)
        _agg_dict[cat] = (key, "sum")

    if progress_callback:
        progress_callback(0.5, "計算庫存建議中...")

    raw = _add_recommendation_columns(df, rates=rates)

    transfer_col = "After new Safety Qty / ND can transfer out of QTY"
    raw["_transfer_num"] = pd.to_numeric(raw.get(transfer_col, _safe_default_series(0, raw.index)), errors="coerce").fillna(0)
    _agg_dict[transfer_col] = ("_transfer_num", "sum")

    if progress_callback:
        progress_callback(0.55, "排序彙總中...")

    summary = raw.groupby(group_cols, observed=True).agg(**_agg_dict).fillna(0).astype(int).reset_index()

    summary["SKU"] = summary["SKU"].apply(_safe_sku)
    summary["總計"] = summary[categories].sum(axis=1)

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

    # 清理暫存欄位
    raw.drop(columns=[c for c in raw.columns if c.startswith("_cat_") or c == "_transfer_num"], inplace=True, errors="ignore")

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
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf


def _write_ndrf_generic_pivot(ws, pivot_df, thin_border, header_font, header_fill, bold_font):
    """寫入通用樞紐對比表（Adj. SST / MCH2 共用）。"""
    if pivot_df is None or pivot_df.empty:
        ws.cell(row=1, column=1, value="無樞紐資料").font = Font(size=12)
        return
    headers = list(pivot_df.columns)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    for row_idx, row_data in pivot_df.iterrows():
        excel_row = row_idx + 2
        is_subtotal = "合計" in str(row_data.get("Site Code", ""))
        for col_idx, h in enumerate(headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_subtotal:
                cell.font = bold_font
            if isinstance(val, (int, float)) and val is not None:
                cell.number_format = "#,##0"
    _auto_width(ws)


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

    # 預計算 int 旗標 + 單次 groupby 取代多次 filter+groupby + transfer merge
    transfer_col = "After new Safety Qty / ND can transfer out of QTY"
    raw["_transfer_num"] = pd.to_numeric(raw.get(transfer_col, _safe_default_series(0, raw.index)), errors="coerce").fillna(0)
    _agg_dict = {transfer_col: ("_transfer_num", "sum")}
    _nd_cat_series = pd.Series(nd_source_status, index=raw.index)
    for cat in categories:
        safe_key = f"_cat_{abs(hash(cat)) % 10_000_000}"
        raw[safe_key] = (nd_mask & (_nd_cat_series == cat)).astype(int)
        _agg_dict[cat] = (safe_key, "sum")

    if progress_callback:
        progress_callback(0.55, "彙總排序中...")

    summary = raw.groupby(group_cols, observed=True).agg(**_agg_dict).fillna(0).astype(int).reset_index()

    summary["SKU"] = summary["SKU"].apply(_safe_sku)
    summary["總計"] = summary[categories].sum(axis=1) if categories else 0

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

    # 清理暫存欄位
    raw.drop(columns=[c for c in raw.columns if c.startswith("_cat_") or c == "_transfer_num"], inplace=True, errors="ignore")

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
        # Shop Monthly Sold Qty (P90D) 已是月均，直接作為 projected
        p90d_monthly = _safe_numeric(df.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0.0, df.index)))
        projected = p90d_monthly.values
    else:
        projected = projected_mtd_arr

    # OOS 模式下所有 Shop Nature 預設 rate 均為 50%（與 UI 預設一致）
    _default_rate = rates.get("default", 0.50)
    _tourist_rate = rates.get("tourist", 0.50)
    _mix_rate = rates.get("mix", 0.50)

    rate_arr = _nature_rate_arr(shop_nature, _default_rate, _tourist_rate, _mix_rate)

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
            OOS_SKU數量=("_sku_oos", "nunique"),
            Potential_OOS_SKU數量=("_sku_potential_oos", "nunique"),
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
                "OOS_SKU數量": [int(df["_sku_oos"].dropna().nunique())],
                "Potential_OOS_SKU數量": [int(df["_sku_potential_oos"].dropna().nunique())],
                "建議SafetyQty增加行數": [int(by_shop["建議SafetyQty增加行數"].sum())],
                "其他需調整SKU數量": [int(df["_sku_normal_insufficient"].dropna().nunique())],
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

    # 預計算 int 旗標（單次 groupby 取代多次 filter+groupby）
    _agg_dict = {"Safety Qty Increase": ("Safety Qty Increase", "sum")}
    normal_mask = (status_col == "Normal") & (safety_increase > 0)
    raw["_is_normal_inc"] = normal_mask.astype(int)
    _agg_dict[SECONDARY_CATEGORY] = ("_is_normal_inc", "sum")
    for cat in OOS_CATEGORIES:
        raw[f"_is_{cat.replace(' ', '_').replace('(', '').replace(')', '').replace('/', '_')}"] = (status_col == cat).astype(int)
        _agg_dict[cat] = (f"_is_{cat.replace(' ', '_').replace('(', '').replace(')', '').replace('/', '_')}", "sum")

    _renames = {}
    for cat in OOS_CATEGORIES:
        _key = f"_is_{cat.replace(' ', '_').replace('(', '').replace(')', '').replace('/', '_')}"
        _renames[_key] = cat

    summary = raw.groupby(group_cols, observed=True).agg(**_agg_dict).fillna(0).astype(int).reset_index()
    if _renames:
        summary = summary.rename(columns=_renames)
    summary["SKU"] = summary["SKU"].apply(_safe_sku)
    summary["總計"] = summary[categories].sum(axis=1)

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

    # 清理暫存欄位
    raw.drop(columns=[c for c in raw.columns if c.startswith("_is_")], inplace=True, errors="ignore")

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
        "OOS_SKU數量",
        "Potential_OOS_SKU數量",
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


def _build_site_to_name(df: pd.DataFrame) -> dict:
    """建立 Site Code → Shop Name 查找表（供 Carry SKU / NDRF 相關函式共用）。"""
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
    return site_to_name


def _build_carry_pivot(df: pd.DataFrame, category_col: str, site_to_name: dict | None = None) -> pd.DataFrame:
    df = df.copy()

    if site_to_name is None:
        site_to_name = _build_site_to_name(df)

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


def _build_carry_dashboard(df: pd.DataFrame, site_to_name: dict | None = None) -> pd.DataFrame:
    df = df.copy()

    if site_to_name is None:
        site_to_name = _build_site_to_name(df)

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

    # 建立 site_to_name 一次，供所有 carry 分析函式共用
    _carry_site_to_name = _build_site_to_name(df)

    # 平行化：同時計算 pivot（Adj. MCH1 + Adj. SST）與 dashboard
    with ThreadPoolExecutor(max_workers=3) as _executor:
        f_left = _executor.submit(_build_carry_pivot, df, "Adj. MCH1", _carry_site_to_name) if "Adj. MCH1" in df.columns else None
        f_right = _executor.submit(_build_carry_pivot, df, "Adj. SST", _carry_site_to_name) if "Adj. SST" in df.columns else None
        f_dash = _executor.submit(_build_carry_dashboard, df, _carry_site_to_name)

        summary_left = f_left.result() if f_left else pd.DataFrame()
        summary_right = f_right.result() if f_right else pd.DataFrame()
        dashboard_df = f_dash.result()

    if progress_callback:
        progress_callback(0.6, "建構分析摘要...")

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


# ─── NDRF Month Comparison Functions ─────────────────────────────────────────────


def _build_ndrf_adj_sst_combined_pivot(df_a, df_b, month_a_label, month_b_label):
    """Wrapper — Adj. SST 樞紐對比（委託通用函式）。"""
    return _build_ndrf_category_combined_pivot(df_a, df_b, month_a_label, month_b_label, "Adj. SST")

def _build_ndrf_adj_mch1_combined_pivot(df_a, df_b, month_a_label, month_b_label):
    """Wrapper — Adj. MCH1 樞紐對比（委託通用函式）。"""
    return _build_ndrf_category_combined_pivot(df_a, df_b, month_a_label, month_b_label, "Adj. MCH1")


def transform_ndrf_comparison(
    df_input: pd.DataFrame,
    month_a: str,
    month_b: str,
    progress_callback=None,
    shop_natures: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, list[str]]:
    """對比兩個月份每個店舖的 Carry SKU 數量、RF SKU 數量、ND With Stock SKU 數量，並拆解 Adj. SST 分析。

    Returns:
        comparison_by_shop: 店舖級月份對比表（欄位標明實際月份名稱）
        adj_sst_pivot: 單一 Adj. SST 樞紐對比表（兩月份數據並列）
        raw: 合併後完整 RAW 資料（含 Month 欄位，僅保留 Shop Name）
        warnings: 警告訊息列表
    """
    missing = [c for c in ["Month", "SKU", "RP Type", "Shop Stock Qty (CM)", "Adj. SST", "Shop Stock Status", "Site Code"] if c not in df_input.columns]
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    warnings = []

    available_months = sorted(df_input["Month"].dropna().unique())
    if len(available_months) < 2:
        raise ValueError(f"Month 欄位至少需有 2 個不同月份值，目前僅有: {available_months}")

    if progress_callback:
        progress_callback(0.02, "驗證通過")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)
    if df.empty:
        raise ValueError("過濾後無有效資料列")

    if "SKU Desc" not in df.columns:
        df = _compute_raw_extra_columns(df)

    if "Shop" not in df.columns:
        shop_name_vals = [
            _extract_shop_name(
                df["Item_Shop"].iloc[i] if "Item_Shop" in df.columns else "",
                df["Site Code"].iloc[i] if "Site Code" in df.columns else ""
            )
            for i in range(len(df))
        ]
        df["Shop"] = shop_name_vals

    df = df.copy()
    rp_type = df["RP Type"].fillna("").astype(str).str.strip().str.upper()
    df["_rp"] = rp_type
    shop_stock_cm = _safe_numeric(df["Shop Stock Qty (CM)"])

    non_nd_rf = set(rp_type.unique()) - {"ND", "RF", ""}
    if non_nd_rf:
        warnings.append(f"偵測到非 ND/RF 的 RP Type 值: {', '.join(sorted(non_nd_rf))}，這些列將不納入統計")

    ndrf_mask = rp_type.isin(["ND", "RF"])
    df = df[ndrf_mask].copy()

    shop_stock_value_cm = _safe_numeric(df.get("Shop Stock Value (CM)", _safe_default_series(0, df.index)))

    df["_is_carry"] = ~((rp_type[ndrf_mask] == "ND") & (shop_stock_cm[ndrf_mask] == 0))
    df["_is_rf"] = (rp_type[ndrf_mask] == "RF").astype(int)
    df["_is_nd_with_stock"] = ((rp_type[ndrf_mask] == "ND") & (shop_stock_cm[ndrf_mask] > 0)).astype(int)
    df["_stock_qty"] = shop_stock_cm[ndrf_mask]
    df["_stock_value"] = shop_stock_value_cm[ndrf_mask]

    month_a_df = df[df["Month"].astype(str).str.strip() == month_a].copy()
    month_b_df = df[df["Month"].astype(str).str.strip() == month_b].copy()

    if progress_callback:
        progress_callback(0.2, "建構月份對比表...")

    # 平行化：同時計算 3 個獨立對比表
    with ThreadPoolExecutor(max_workers=3) as _executor:
        f_shop = _executor.submit(_build_ndrf_shop_comparison, month_a_df, month_b_df, month_a, month_b)
        f_sst = _executor.submit(_build_ndrf_adj_sst_combined_pivot, month_a_df, month_b_df, month_a, month_b)
        f_mch1 = _executor.submit(_build_ndrf_adj_mch1_combined_pivot, month_a_df, month_b_df, month_a, month_b)

    comparison_by_shop = f_shop.result()
    adj_sst_pivot = f_sst.result()
    mch2_pivot = f_mch1.result()

    if progress_callback:
        progress_callback(0.6, "合併 RAW 資料...")

    combined_raw = pd.concat([month_a_df, month_b_df], ignore_index=True)
    drop_cols = [c for c in combined_raw.columns if c.startswith("_")]
    combined_raw = combined_raw.drop(columns=drop_cols, errors="ignore")
    if "Shop" in combined_raw.columns and "Shop Name" in combined_raw.columns:
        combined_raw = combined_raw.drop(columns=["Shop"])

    if progress_callback:
        progress_callback(0.7, "轉換完成")

    return comparison_by_shop, adj_sst_pivot, mch2_pivot, combined_raw, warnings


def _build_ndrf_shop_comparison(df_a: pd.DataFrame, df_b: pd.DataFrame, month_a_label: str, month_b_label: str) -> pd.DataFrame:
    """建構店舖級月份對比表，欄位名稱直接使用實際月份標籤（如 Carry (Mar2026)）。"""

    def _agg_shop(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return pd.DataFrame(columns=["Site Code", "Shop Name"])
        site_code = df["Site Code"].fillna("").astype(str).str.strip()
        shop_col = df["Shop Name"].fillna("").astype(str).str.strip() if "Shop Name" in df.columns else (
            df["Shop"].fillna("").astype(str).str.strip() if "Shop" in df.columns else site_code
        )
        agg = df.groupby(["Site Code", shop_col.name], dropna=False, observed=True).agg(
            carry_count=("_is_carry", "sum"),
            rf_count=("_is_rf", "sum"),
            nd_with_stock_count=("_is_nd_with_stock", "sum"),
        ).reset_index()
        agg.columns = ["Site Code", "Shop Name", "carry_count", "rf_count", "nd_with_stock_count"]
        for c in ["carry_count", "rf_count", "nd_with_stock_count"]:
            agg[c] = agg[c].fillna(0).astype(int)
        return agg

    shop_a = _agg_shop(df_a)
    shop_b = _agg_shop(df_b)

    merged = pd.merge(shop_a, shop_b, on=["Site Code", "Shop Name"], how="outer", suffixes=(f"_{month_a_label}", f"_{month_b_label}"))

    carry_a_col = f"carry_count_{month_a_label}"
    carry_b_col = f"carry_count_{month_b_label}"
    rf_a_col = f"rf_count_{month_a_label}"
    rf_b_col = f"rf_count_{month_b_label}"
    nd_a_col = f"nd_with_stock_count_{month_a_label}"
    nd_b_col = f"nd_with_stock_count_{month_b_label}"

    for c in [carry_a_col, carry_b_col, rf_a_col, rf_b_col, nd_a_col, nd_b_col]:
        if c not in merged.columns:
            merged[c] = 0
        merged[c] = merged[c].fillna(0).astype(int)

    merged[f"Carry ({month_a_label})"] = merged[carry_a_col]
    merged[f"Carry ({month_b_label})"] = merged[carry_b_col]
    merged["Carry 變化"] = merged[f"Carry ({month_b_label})"] - merged[f"Carry ({month_a_label})"]

    merged[f"RF ({month_a_label})"] = merged[rf_a_col]
    merged[f"RF ({month_b_label})"] = merged[rf_b_col]
    merged["RF 變化"] = merged[f"RF ({month_b_label})"] - merged[f"RF ({month_a_label})"]

    merged[f"ND With Stock ({month_a_label})"] = merged[nd_a_col]
    merged[f"ND With Stock ({month_b_label})"] = merged[nd_b_col]
    merged["ND With Stock 變化"] = merged[f"ND With Stock ({month_b_label})"] - merged[f"ND With Stock ({month_a_label})"]

    display_cols = [
        "Site Code", "Shop Name",
        f"Carry ({month_a_label})", f"Carry ({month_b_label})", "Carry 變化",
        f"RF ({month_a_label})", f"RF ({month_b_label})", "RF 變化",
        f"ND With Stock ({month_a_label})", f"ND With Stock ({month_b_label})", "ND With Stock 變化",
    ]
    result = merged[display_cols].copy()
    result = result.sort_values("Site Code").reset_index(drop=True)

    total = {}
    for col in result.columns:
        if col == "Site Code":
            total[col] = "總計"
        elif col == "Shop Name":
            total[col] = ""
        elif result[col].dtype in ("int64", "int32", "float64"):
            total[col] = int(result[col].sum())
        else:
            total[col] = None
    total_row = pd.DataFrame([total])
    result = pd.concat([result, total_row], ignore_index=True)

    for col in result.columns:
        if result[col].dtype == "float64":
            result[col] = result[col].astype(int)

    return result


def _build_ndrf_category_combined_pivot(df_a: pd.DataFrame, df_b: pd.DataFrame, month_a_label: str, month_b_label: str, category_col: str) -> pd.DataFrame:
    """建構單一樞紐對比表（依 category_col 分組），兩月份數據並列，含 ND With Stock、RF Total、Carry、Stock Qty、Stock Value 及變化欄位。"""
    def _pivot_for(df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return pd.DataFrame(columns=["Site Code", category_col, "ND_With_Stock", "RF_Total", "Carry", "Stock_Qty", "Stock_Value"])
        df = df.copy()
        rp = df["_rp"]
        stock = df["_stock_qty"]
        df["_cat"] = np.select(
            [(rp == "ND") & (stock > 0), (rp == "ND") & (stock == 0)],
            ["ND_With_Stock", "ND_No_Stock"],
            default="RF"
        )
        count_pivot = df.groupby(["Site Code", category_col, "_cat"], dropna=False, observed=True).size().reset_index(name="count")
        wide = count_pivot.pivot_table(index=["Site Code", category_col], columns="_cat", values="count", fill_value=0).reset_index()
        for c in ["ND_With_Stock", "RF"]:
            if c not in wide.columns:
                wide[c] = 0
        wide["ND_With_Stock"] = wide["ND_With_Stock"].fillna(0).astype(int)
        wide["RF_Total"] = wide["RF"].fillna(0).astype(int)
        wide["Carry"] = (wide["ND_With_Stock"] + wide["RF_Total"]).astype(int)
        stock_agg = df.groupby(["Site Code", category_col], dropna=False, observed=True).agg(
            Stock_Qty=("_stock_qty", "sum"),
            Stock_Value=("_stock_value", "sum"),
        ).reset_index()
        result = wide.merge(stock_agg, on=["Site Code", category_col], how="left")
        for c in ["Stock_Qty", "Stock_Value"]:
            result[c] = result[c].fillna(0).astype(int)
        keep = ["Site Code", category_col, "ND_With_Stock", "RF_Total", "Carry", "Stock_Qty", "Stock_Value"]
        return result[keep].sort_values(["Site Code", category_col])

    pivot_a = _pivot_for(df_a) if not df_a.empty else pd.DataFrame()
    pivot_b = _pivot_for(df_b) if not df_b.empty else pd.DataFrame()

    if pivot_a.empty and pivot_b.empty:
        return pd.DataFrame(columns=["Site Code", "Shop Name", category_col])

    if pivot_a.empty:
        pivot_a = pivot_b[["Site Code", category_col]].copy()
        for c in ["ND_With_Stock", "RF_Total", "Carry", "Stock_Qty", "Stock_Value"]:
            pivot_a[c] = 0
    if pivot_b.empty:
        pivot_b = pivot_a[["Site Code", category_col]].copy()
        for c in ["ND_With_Stock", "RF_Total", "Carry", "Stock_Qty", "Stock_Value"]:
            pivot_b[c] = 0

    site_to_name = {}
    for _src in [df_a, df_b]:
        if _src.empty:
            continue
        sc = _src["Site Code"].fillna("").astype(str).str.strip()
        sn = _src.get("Shop Name", _src.get("Shop", _safe_default_series("", _src.index))).fillna("").astype(str).str.strip()
        for s, n in zip(sc, sn):
            if s and s not in site_to_name:
                site_to_name[s] = n if n else ""

    data_cols = ["ND_With_Stock", "RF_Total", "Carry", "Stock_Qty", "Stock_Value"]
    a_suffix = f"_{month_a_label}"
    b_suffix = f"_{month_b_label}"

    merged = pd.merge(pivot_a, pivot_b, on=["Site Code", category_col], how="outer",
                      suffixes=(a_suffix, b_suffix))

    _all_a = [f"{c}{a_suffix}" for c in data_cols]
    _all_b = [f"{c}{b_suffix}" for c in data_cols]
    for c in _all_a + _all_b:
        if c not in merged.columns:
            merged[c] = 0
        merged[c] = merged[c].fillna(0).astype(int)

    merged = merged.sort_values(["Site Code", category_col]).reset_index(drop=True)

    def _col_name(base: str, month: str) -> str:
        return f"{base.replace('_', ' ')} {month}"
    def _delta_name(base: str) -> str:
        return f"{base.replace('_', ' ')} 變化"

    rename_map = {}
    for dc in data_cols:
        rename_map[f"{dc}{a_suffix}"] = _col_name(dc, month_a_label)
        rename_map[f"{dc}{b_suffix}"] = _col_name(dc, month_b_label)
    merged = merged.rename(columns=rename_map)
    merged = merged.drop(columns=[c for c in merged.columns if c in ["Shop Name", "Shop"] if c in merged.columns], errors="ignore")

    for dc in data_cols:
        a_c = _col_name(dc, month_a_label)
        b_c = _col_name(dc, month_b_label)
        if a_c in merged.columns and b_c in merged.columns:
            merged[_delta_name(dc)] = merged[b_c] - merged[a_c]

    col_order = ["Site Code", category_col]
    for dc in data_cols:
        col_order.append(_col_name(dc, month_a_label))
    for dc in data_cols:
        col_order.append(_col_name(dc, month_b_label))
    for dc in data_cols:
        col_order.append(_delta_name(dc))
    existing = [c for c in col_order if c in merged.columns]
    trailing = [c for c in merged.columns if c not in existing]
    merged = merged[existing + trailing]

    merged.insert(1, "Shop Name", merged["Site Code"].map(site_to_name).fillna(""))

    result_rows = []
    for site, group in merged.groupby("Site Code", sort=False):
        result_rows.append(group)
        sub = {"Site Code": [f"{site} 合計"], "Shop Name": [""], category_col: [None]}
        for dc in data_cols:
            sub[_col_name(dc, month_a_label)] = [int(group[_col_name(dc, month_a_label)].sum())]
            sub[_col_name(dc, month_b_label)] = [int(group[_col_name(dc, month_b_label)].sum())]
            sub[_delta_name(dc)] = [int(group[_delta_name(dc)].sum())]
        result_rows.append(pd.DataFrame(sub))

    result = pd.concat(result_rows, ignore_index=True) if result_rows else merged
    return result


def generate_excel_ndrf_comparison(
    comparison_by_shop: pd.DataFrame,
    adj_sst_pivot: pd.DataFrame,
    mch2_pivot: pd.DataFrame,
    raw_df: pd.DataFrame,
    month_a: str,
    month_b: str,
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

    # ── Sheet 1: Month Comparison (By Shop) ──
    ws_compare = wb.active
    ws_compare.title = "月份對比 (By Shop)"

    headers = list(comparison_by_shop.columns)
    for col_idx, h in enumerate(headers, 1):
        cell = ws_compare.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, row_data in comparison_by_shop.iterrows():
        excel_row = row_idx + 2
        is_total = row_data.get("Site Code") == "總計"
        for col_idx, h in enumerate(headers, 1):
            val = row_data.get(h)
            if pd.isna(val):
                val = None
            cell = ws_compare.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if is_total:
                cell.font = bold_font
            if isinstance(val, (int, float)) and col_idx >= 3:
                cell.number_format = "#,##0"
    _auto_width(ws_compare)

    if progress_callback:
        progress_callback(0.3, "建立 Adj. SST 樞紐對比表...")

    # ── Sheet 2: Adj. SST Comparison Pivot (both months side by side) ──
    ws_pivot = wb.create_sheet(title=f"Adj. SST 樞紐對比")
    _write_ndrf_generic_pivot(ws_pivot, adj_sst_pivot, thin_border, header_font, header_fill, bold_font)

    if progress_callback:
        progress_callback(0.45, "建立 MCH2 樞紐對比表...")

    # ── Sheet 3: Adj. MCH1 Comparison Pivot ──
    ws_mch2 = wb.create_sheet(title="Adj. MCH1 樞紐對比")
    _write_ndrf_generic_pivot(ws_mch2, mch2_pivot, thin_border, header_font, header_fill, bold_font)

    if progress_callback:
        progress_callback(0.6, "建立 RAW 工作表...")

    # ── Sheet 4: RAW ──
    ws_raw = wb.create_sheet(title="RAW")
    _write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.6, progress_end=0.9)

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
    launch_dt = _parse_launch_dates(launch_raw)

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
    損失基準並列：(1) Sold Qty (M-2) 需求基準；(2) Shop Monthly Sold Qty (P90D) 月均銷售。
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
        loss_p90d = np.where(is_oos & (p90d.values > 0), np.maximum(0, p90d.values - sold.values), 0)
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
    # 正規化 Shop Nature（strip）並寫回，確保後續精確比對一致
    shop_nature_raw = df.get("Shop Nature", _safe_default_series("", idx))
    df["Shop Nature"] = shop_nature_raw.fillna("").astype(str).str.strip()
    shop_nature = df["Shop Nature"]
    is_mix = shop_nature == "Mix"
    is_nt = shop_nature == "Non Tourist"
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


def _build_sku_info_lookup(raw: pd.DataFrame) -> pd.DataFrame:
    """建立 SKU 級基本資訊查找表（單一次 groupby，供多個 summary 函式共用）。

    回傳 DataFrame 每列一個 SKU，含 Brand / Buyer / SKU Desc / Launch Date /
    Is New SKU (3mo) / Vendor MC / Adj. MCH1 / Item Status / Item Class 等 first-value 欄位。
    """
    if "SKU" not in raw.columns:
        return pd.DataFrame()
    info_cols = [c for c in [
        "Brand", "Buyer", "SKU Desc", "Launch Date", "Is New SKU (3mo)",
        "Vendor MC", "Adj. MCH1", "Item Status", "Item Class",
    ] if c in raw.columns]
    if not info_cols:
        return raw[["SKU"]].drop_duplicates().reset_index(drop=True)
    agg_dict = {col: (col, "first") for col in info_cols}
    return raw.groupby("SKU", dropna=False).agg(**agg_dict).reset_index()


def _build_sku_allocation_summary(raw: pd.DataFrame, sku_lookup: pd.DataFrame | None = None) -> pd.DataFrame:
    """SKU 級跨 Nature 調貨建議。

    針對同時存在於多種 Shop Nature 的 SKU：當 Tourist 店舖缺貨（CM Shop Stock=0）
    但同 SKU 在 Non-Tourist / Mix 店舖仍有存貨時，列出可調貨支援機會。
    排序：Tourist 缺貨損失 → 可調貨存貨量（降序）。
    
    Args:
        sku_lookup: 預先建好的 SKU 基本資訊查找表，為 None 時即時計算。
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


def _build_sku_stock_diagnosis(raw: pd.DataFrame, sku_lookup: pd.DataFrame | None = None) -> pd.DataFrame:
    """by-SKU 缺貨 Review 表（Tourist + Mix + Non Tourist 缺貨 SKU）。

    每列為一個缺貨 SKU，彙總：
    - Tourist / Mix / Non Tourist 各自缺貨店數與銷售
    - D001 倉庫三期貨量 (M-2/M-1/CM) + D001 狀態分類（源頭斷貨 vs 補貨問題）
    - 店舖三期總存貨 (M-2/M-1/CM)
    排序：Tourist 缺貨店數 → Mix 缺貨店數 → Non Tourist 缺貨店數 → Tourist 損失(P90D)（皆降序）。

    效能優化：原本 6 次獨立 groupby + 5 次 merge → 改為 2 次 groupby（SKU-level ＋ SKU×Nature-level）。

    Args:
        sku_lookup: 預先建好的 SKU 基本資訊查找表，為 None 時即時計算。
    """
    if "SKU" not in raw.columns or "Shop Nature" not in raw.columns:
        return pd.DataFrame()

    # 只看 Tourist + Mix + Non Tourist 的缺貨行（CM Shop Stock=0 且有需求）
    cm_stock_vec = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))
    focus_mask = (
        raw["Shop Nature"].isin(["Tourist", "Mix", "Non Tourist"])
        & raw["Is Stock Impacted"]
        & (cm_stock_vec <= 0)
    )
    if not focus_mask.any():
        return pd.DataFrame()

    # 預先建立 SKU 基本資訊查找表（只建一次，之後所有函式可重用）
    if sku_lookup is None:
        sku_lookup = _build_sku_info_lookup(raw)

    focus = raw.loc[focus_mask].copy()

    # 補上需要的數值欄位
    focus["_loss_p90d"] = _safe_numeric(focus.get("Max Sales Loss (P90D)", pd.Series(dtype=float, index=focus.index)))
    focus["_loss_m2"] = _safe_numeric(focus.get("Max Sales Loss (M-2)", pd.Series(dtype=float, index=focus.index)))
    focus["_cm_sold"] = _safe_numeric(focus.get("Sold Qty (CM - MTD)", pd.Series(dtype=float, index=focus.index)))
    focus["_p90d"] = _safe_numeric(focus.get("Shop Monthly Sold Qty (P90D)", pd.Series(dtype=float, index=focus.index)))
    for label, col in [("M-2", "Shop Stock Qty (M-2)"), ("M-1", "Shop Stock Qty (M-1)"), ("CM", "Shop Stock Qty (CM)")]:
        focus[f"_shop_stock_{label}"] = _safe_numeric(focus.get(col, pd.Series(dtype=float, index=focus.index)))
    for label, col in [("M-2", "D001 Stock Qty (M-2)"), ("M-1", "D001 Stock Qty (M-1)"), ("CM", "D001 Stock Qty (CM)")]:
        focus[f"_d001_{label}"] = _safe_numeric(focus.get(col, pd.Series(dtype=float, index=focus.index)))

    # 判斷 Site Code 可用性（用於 nunique）
    has_site_code = "Site Code" in focus.columns

    # ── 單次 groupby 計算所有 SKU-level 指標（D001 + Shop Stock + Brand/Buyer 已從 lookup 取得）──
    sku_agg = focus.groupby("SKU", dropna=False).agg(
        **{
            "D001_Stock_M2": ("_d001_M-2", "first"),
            "D001_Stock_M1": ("_d001_M-1", "first"),
            "D001_Stock_CM": ("_d001_CM", "first"),
            "店舖存貨總_M2": ("_shop_stock_M-2", "sum"),
            "店舖存貨總_M1": ("_shop_stock_M-1", "sum"),
            "店舖存貨總_CM": ("_shop_stock_CM", "sum"),
        }
    ).reset_index()

    # ── 單次 groupby 計算 SKU×Nature-level 指標（取代 3 次獨立 filter+groupby）──
    _DIAG_NATURES = ["Tourist", "Mix", "Non Tourist"]
    nat_agg_dict = {}
    for nat in _DIAG_NATURES:
        shop_count_col = "Site Code" if has_site_code else "SKU"
        nat_agg_dict[f"{nat}缺貨店數"] = (shop_count_col, "nunique") if has_site_code else (shop_count_col, "count")
        nat_agg_dict[f"{nat}損失_M2"] = ("_loss_m2", "sum")
        nat_agg_dict[f"{nat}損失_P90D"] = ("_loss_p90d", "sum")
        nat_agg_dict[f"{nat}_CM銷售件數"] = ("_cm_sold", "sum")
        nat_agg_dict[f"{nat}_P90D總量(90日)"] = ("_p90d", "sum")

    sku_nat_agg = focus.groupby(["SKU", "Shop Nature"], dropna=False, observed=True).agg(
        **{k: v for k, v in nat_agg_dict.items() if v[0] in focus.columns or has_site_code}
    ).reset_index()

    # 將 Nature 展開為寬表（pivot 比 merge 3 次快，避免重複欄位問題）
    merged = sku_agg.copy()
    for nat in _DIAG_NATURES:
        nat_cols = [c for c in [f"{nat}缺貨店數", f"{nat}損失_M2", f"{nat}損失_P90D",
                                 f"{nat}_CM銷售件數", f"{nat}_P90D總量(90日)"] if c in sku_nat_agg.columns]
        if nat_cols:
            nat_sub = sku_nat_agg[sku_nat_agg["Shop Nature"] == nat][["SKU"] + nat_cols].copy()
            if not nat_sub.empty:
                merged = merged.merge(nat_sub, on="SKU", how="left")

    # 補上 SKU 基本資訊（從 lookup，取代原本的 brand_per_sku groupby）
    lookup_cols = ["SKU"] + [c for c in ["Brand", "Buyer", "SKU Desc", "Launch Date", "Is New SKU (3mo)"] if c in sku_lookup.columns]
    merged = merged.merge(sku_lookup[lookup_cols], on="SKU", how="left")

    # NDRF 行數統計（從完整 raw 計算）
    rp_full = (
        raw.get("RP Type", _safe_default_series("", raw.index))
        .fillna("").astype(str).str.strip().str.upper()
    )
    raw_ndrf = raw.assign(_is_nd=(rp_full == "ND"), _is_rf=(rp_full == "RF"))
    ndrf_per_sku = (
        raw_ndrf.groupby("SKU", dropna=False)
        .agg(
            NDRF_ND行數=("_is_nd", "sum"),
            NDRF_RF行數=("_is_rf", "sum"),
        )
        .reset_index()
    )
    merged = merged.merge(ndrf_per_sku, on="SKU", how="left")
    for c in ["NDRF_ND行數", "NDRF_RF行數"]:
        merged[c] = merged[c].fillna(0).astype(int)

    # 確保各 Nature 欄位存在（當 focus 無該 Nature 時補 0）
    _diag_sort_keys = []
    for nat in _DIAG_NATURES:
        for col in [f"{nat}缺貨店數", f"{nat}損失_M2", f"{nat}損失_P90D",
                     f"{nat}_CM銷售件數", f"{nat}_P90D總量(90日)"]:
            if col not in merged.columns:
                merged[col] = 0
            else:
                merged[col] = merged[col].fillna(0).astype(int)
        _diag_sort_keys.append(f"{nat}缺貨店數")
    # 排序鍵：Tourist 損失 P90D（ Tour-specific，供排序用）
    if "Tourist損失_P90D" not in merged.columns:
        merged["Tourist損失_P90D"] = 0
    _diag_sort_keys.append("Tourist損失_P90D")

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
                "NDRF_ND行數", "NDRF_RF行數",
                "Tourist缺貨店數", "Tourist_CM銷售件數", "Tourist_P90D總量(90日)",
                "Mix缺貨店數", "Mix_CM銷售件數", "Mix_P90D總量(90日)",
                "Non Tourist缺貨店數", "Non Tourist_CM銷售件數", "Non Tourist_P90D總量(90日)",
                "D001 狀態", "D001_Stock_M2", "D001_Stock_M1", "D001_Stock_CM",
                "店舖存貨總_M2", "店舖存貨總_M1", "店舖存貨總_CM"]
    # 排序
    merged = merged.sort_values(
        _diag_sort_keys,
        ascending=[False] * len(_diag_sort_keys),
    ).reset_index(drop=True)
    avail = [c for c in out_cols if c in merged.columns]
    merged = merged[avail]

    return merged


# Turnday 分級順序（用於排序；含 RF/ND 拆分）
_TURNDAY_BAND_ORDER = [
    "<7日(過低)", "7-14日(理想)", "14-30日(偏高)", ">30日(過剩)",
    "無銷售有存貨(RF)", "無銷售有存貨(ND)", "無存貨",
]


def _build_inventory_distribution_summary(raw: pd.DataFrame, stock_threshold: int = 1000,
                                            sku_lookup: pd.DataFrame | None = None) -> dict:
    """存貨分佈彙總 — 回傳 dict 含 4 個 DataFrame。

    Args:
        stock_threshold: 低存貨 SKU 門檻（D001 + Shop Stock 總和 < 此值視為低存貨，預設 300）。
        sku_lookup: 預先建好的 SKU 基本資訊查找表，為 None 時即時計算。

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

    # 其他 Nature 過剩存貨：Mix>14日（Mix Excess Stock>0）或 Non Tourist>14日（Non Tourist Excess Stock>0）
    # 統一到相同 14 日標準，與 Turnday 分級邏輯一致
    mix_excess_vec = _safe_numeric(raw.get("Mix Excess Stock", _safe_default_series(0, raw.index)))
    nt_excess_vec = _safe_numeric(raw.get("Non Tourist Excess Stock", _safe_default_series(0, raw.index)))
    other_over = raw[
        (nature != "Tourist")
        & (
            ((nature == "Mix") & (mix_excess_vec > 0))
            | ((nature == "Non Tourist") & (nt_excess_vec > 0))
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

        # 附加 SKU 級資訊：SKU Desc / Launch Date / Buyer（從 lookup 讀取，避免重複 groupby）
        if sku_lookup is None:
            sku_lookup = _build_sku_info_lookup(raw)
        info_from_lookup = [c for c in ["SKU Desc", "Launch Date", "Is New SKU (3mo)", "Buyer"]
                           if c in sku_lookup.columns and c not in merged.columns]
        if info_from_lookup:
            merged = merged.merge(sku_lookup[["SKU"] + info_from_lookup], on="SKU", how="left")

        # D001 Net Stock + Shop Stock 總數（不在 lookup 中，獨立計算）
        d001_net = "D001 Net Stock Qty (CM)"
        has_d001_net = d001_net in raw.columns
        has_shop_cm = "Shop Stock Qty (CM)" in raw.columns
        if has_d001_net and has_shop_cm:
            extra_agg = (
                raw.groupby("SKU", dropna=False)
                .agg(_d001_net=(d001_net, "first"), _shop_total=("Shop Stock Qty (CM)", "sum"))
                .reset_index()
            )
            merged = merged.merge(extra_agg, on="SKU", how="left")
            merged["總可用存貨"] = (_safe_numeric(merged["_d001_net"]) + _safe_numeric(merged["_shop_total"])).astype(int)
            merged.drop(columns=["_d001_net", "_shop_total"], inplace=True)

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
    result["transfer_priority"] = _build_transfer_priority_analysis(raw, sku_lookup)

    return result


def _build_transfer_priority_analysis(raw: pd.DataFrame, sku_lookup: pd.DataFrame | None = None) -> dict:
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

    # Add SKU-level info: Brand, SKU Desc, Buyer（從 lookup 讀取，避免重複 groupby）
    if sku_lookup is None:
        sku_lookup = _build_sku_info_lookup(raw)
    info_from_lookup = [c for c in ["Brand", "SKU Desc", "Buyer"]
                       if c in sku_lookup.columns and c not in detail.columns]
    if info_from_lookup:
        detail = detail.merge(sku_lookup[["SKU"] + info_from_lookup], on="SKU", how="left")

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


def _build_sku_monthly_analysis(raw: pd.DataFrame, stock_threshold: int = 1000, sku_lookup: pd.DataFrame | None = None) -> pd.DataFrame:
    """按 SKU 計算遊客/混合/非遊客的每月銷量及存儲量分析。

    每列為一個 SKU，彙總各 Nature 的月均銷售、當期庫存、店數等。
    低存貨 SKU（D001 + Shop < threshold）排在前面。

    排序：低存貨優先 → Tourist RF 月均銷售降序。

    效能優化：
    - 使用共用 sku_lookup 取代 3 次獨立 groupby.first() merge
    - ND 行數改用預計算 int 欄位取代 groupby 內的 lambda
    - 補充資訊欄位改用單次 merge 取代逐欄 merge

    Args:
        sku_lookup: 預先建好的 SKU 基本資訊查找表，為 None 時即時計算。
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
    work["_is_nd"] = (work["_rp"] == "ND").astype(int)

    # SKU 級基本資訊 (ND 行數改用預計算 int 取代 lambda)
    sku_base = (
        work.groupby("SKU", dropna=False)
        .agg(
            D001_庫存=("_d001", "first"),
            總店舖庫存=("_cm_stock", "sum"),
            總行數=("SKU", "count"),
            RF行數=("_is_rf", "sum"),
            ND行數=("_is_nd", "sum"),
        )
        .reset_index()
    )
    # 使用共用 lookup 表（單次 merge），取代原本 3 次獨立 groupby.merge
    if sku_lookup is None:
        sku_lookup = _build_sku_info_lookup(raw)
    info_cols = ["SKU"] + [c for c in ["Brand", "Buyer", "SKU Desc"] if c in sku_lookup.columns]
    sku_base = sku_base.merge(sku_lookup[info_cols], on="SKU", how="left")
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
    sku_nat["月均銷售"] = sku_nat["月均銷售"].astype(float)

    # RF 月均銷售（只計 RF 店）
    work_rf = work[work["_rp"] == "RF"]
    if not work_rf.empty:
        sku_nat_rf = (
            work_rf.groupby(["SKU", "_nature"], dropna=False, observed=True)
            .agg(RF月均銷售=("_p90d", "sum"))
            .reset_index()
        )
        sku_nat_rf["RF月均銷售"] = sku_nat_rf["RF月均銷售"].astype(float)
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

    # ── 補充資訊欄位（從 sku_lookup 一次性補齊，避免重複 groupby）──
    extra_from_lookup = [c for c in ["Vendor MC", "Adj. MCH1", "Item Status", "Launch Date", "Item Class", "Is New SKU (3mo)"]
                          if c not in sku_base.columns and c in sku_lookup.columns]
    if extra_from_lookup:
        result = result.merge(sku_lookup[["SKU"] + extra_from_lookup], on="SKU", how="left")

    # ── 不分 Nature 的週期存貨/銷售欄位（M-2, M-1 D001/ShopStock；Sold M-2, M-1, MTD） ──
    d001_m2 = _safe_numeric(raw["D001 Stock Qty (M-2)"]) if "D001 Stock Qty (M-2)" in raw.columns else _safe_default_series(0, raw.index)
    d001_m1 = _safe_numeric(raw["D001 Stock Qty (M-1)"]) if "D001 Stock Qty (M-1)" in raw.columns else _safe_default_series(0, raw.index)
    shop_m2 = _safe_numeric(raw.get("Shop Stock Qty (M-2)", _safe_default_series(0, raw.index)))
    shop_m1 = _safe_numeric(raw.get("Shop Stock Qty (M-1)", _safe_default_series(0, raw.index)))
    sold_m2 = _safe_numeric(raw.get("Sold Qty (M-2)", _safe_default_series(0, raw.index)))
    sold_m1 = _safe_numeric(raw.get("Sold Qty (M-1)", _safe_default_series(0, raw.index)))
    sold_mtd = _safe_numeric(raw.get("Sold Qty (CM - MTD)", _safe_default_series(0, raw.index)))

    per_data = pd.DataFrame({
        "SKU": raw["SKU"].fillna("").astype(str).values,
        "_d001_m2": d001_m2.values, "_d001_m1": d001_m1.values,
        "_shop_m2": shop_m2.values, "_shop_m1": shop_m1.values,
        "_sold_m2": sold_m2.values, "_sold_m1": sold_m1.values,
        "_sold_mtd": sold_mtd.values,
    })
    per_agg = per_data.groupby("SKU", dropna=False).agg(
        D001_M2=("_d001_m2", "first"), D001_M1=("_d001_m1", "first"),
        ShopStock_M2=("_shop_m2", "sum"), ShopStock_M1=("_shop_m1", "sum"),
        Sold_M2=("_sold_m2", "sum"), Sold_M1=("_sold_m1", "sum"),
        Sold_MTD=("_sold_mtd", "sum"),
    ).reset_index()
    per_agg = per_agg.rename(columns={
        "D001_M2": "D001_M-2", "D001_M1": "D001_M-1",
        "ShopStock_M2": "ShopStock_M-2", "ShopStock_M1": "ShopStock_M-1",
        "Sold_M2": "Sold_M-2", "Sold_M1": "Sold_M-1",
    })
    result = result.merge(per_agg, on="SKU", how="left")
    for c in ["D001_M-2", "D001_M-1", "ShopStock_M-2", "ShopStock_M-1",
              "Sold_M-2", "Sold_M-1", "Sold_MTD"]:
        if c in result.columns:
            result[c] = result[c].fillna(0).astype(int)

    # ── 欄位順序 ──
    _extra_info_cols = [c for c in ["Vendor MC", "Adj. MCH1", "Item Status", "Launch Date", "Item Class", "Is New SKU (3mo)"] if c in result.columns]
    _sold_cols = [c for c in ["Sold_M-2", "Sold_M-1", "Sold_MTD"] if c in result.columns]
    _d001_cols = [c for c in ["D001_M-2", "D001_M-1", "ShopStock_M-2", "ShopStock_M-1"] if c in result.columns]

    out_cols = ["SKU"]
    for col in ["SKU Desc", "Brand", "Buyer"]:
        if col in result.columns:
            out_cols.append(col)
    out_cols += _extra_info_cols
    out_cols += ["低存貨", "總可用存貨", "D001_庫存", "總店舖庫存"]
    out_cols += _sold_cols
    out_cols += ["總行數", "RF行數", "ND行數"]
    for nat in _NATURES:
        out_cols += [
            f"{nat}_月均銷售", f"{nat}_當期CM銷售", f"{nat}_當期庫存",
            f"{nat}_店數", f"{nat}_平均Turnday", f"{nat}_RF月均銷售",
        ]
    out_cols += _d001_cols
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

# 對齊 Joseph 檔的 RAW 欄位順序（38 欄，不含 NDRF Status，含 Total Stock 等擴充欄位）
_JOSEPH_RAW_COLS = [
    "Buyer", "Brand", "SKU", "SKU Desc", "Item Class", "Launch Date", "Is New SKU (3mo)",
    "Site Code", "Shop Name", "OM",
    "Shop Nature", "Supply Source",
    "Total Stock Qty (M-2)", "Total Stock Qty (M-1)", "Total Stock Qty (CM)",
    "Shop Stock Qty (M-2)", "Shop Stock Qty (M-1)", "Shop Stock Qty (CM)",
    "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
    "Shop Monthly Sold Qty (P90D)", "Is OOS (M-2)", "Is OOS (M-1)", "Is OOS (CM)",
    "OOS Periods Count", "OOS Pattern",
    "有貨店舖 (M-2)", "有貨店舖 (M-1)", "有貨店舖 (CM)",
    "有賣店舖 (M-2)", "有賣店舖 (M-1)", "有賣店舖 (CM)",
    "Eff Price", "Price Tier",
    "SOH Tier (M-2)", "SOH Tier (M-1)", "SOH Tier (CM)",
]

# SOH Tier 邊界（左閉右開），Joseph 分級：a 0(=0), b 1, c 2, d 3-4, e 5-7, f 8-14, g 15-49, h 50+
_SOH_TIER_BINS = [-1, 0, 1, 2, 4, 7, 14, 49, 10**9]
_SOH_TIER_LABELS = ["a 0", "b 1", "c 2", "d 3-4", "e 5-7", "f 8-14", "g 15-49", "h 50+"]

# Price Tier 邊界（下限值），Joseph 分級
_PRICE_TIER_BOUNDS = [0, 100, 300, 500, 800, 1200]
_PRICE_TIER_LABELS = ["a below 100", "b 100-300", "c 300-500", "d 500-800", "e 800-1200", "f 1200+"]


def _soh_tier_label(qty: int | float) -> str | None:
    """將 Shop Stock Qty 轉為 SOH tier 標籤（a 0 … h 50+），NaN/None → None。"""
    if pd.isna(qty):
        return None
    qty = int(qty)
    for i in range(len(_SOH_TIER_BINS) - 1):
        if _SOH_TIER_BINS[i] < qty <= _SOH_TIER_BINS[i + 1]:
            return _SOH_TIER_LABELS[i]
    return _SOH_TIER_LABELS[-1]


def _price_tier_label(eff_price: float) -> str | None:
    """將 Eff Price 轉為 Price Tier 標籤，NaN → None。"""
    if pd.isna(eff_price) or eff_price < 0:
        return None
    for i in range(len(_PRICE_TIER_BOUNDS) - 1):
        lo = _PRICE_TIER_BOUNDS[i]
        hi = _PRICE_TIER_BOUNDS[i + 1]
        if lo <= eff_price < hi:
            return _PRICE_TIER_LABELS[i]
    return _PRICE_TIER_LABELS[-1]  # ≥1200 → f 1200+


def _enrich_raw_joseph(raw: pd.DataFrame) -> pd.DataFrame:
    """在 raw_df 上計算 Joseph 額外欄位的 copy（不修改原文）。"""
    df = raw.copy()
    idx = df.index

    # 1. Total Stock Qty = D001 Stock Qty + groupby(SKU) sum of Shop Stock Qty
    for period, d001_col, shop_col in [
        ("M-2", "D001 Stock Qty (M-2)", "Shop Stock Qty (M-2)"),
        ("M-1", "D001 Stock Qty (M-1)", "Shop Stock Qty (M-1)"),
        ("CM", "D001 Stock Qty (CM)", "Shop Stock Qty (CM)"),
    ]:
        col = f"Total Stock Qty ({period})"
        d001 = _safe_numeric(df.get(d001_col, _safe_default_series(0, idx)))
        shop = _safe_numeric(df.get(shop_col, _safe_default_series(0, idx)))
        sku_total = shop.groupby(df["SKU"], dropna=False).transform("sum")
        df[col] = d001 + sku_total

    # 2. 有貨店舖 / 有賣店舖（0/1 flag）
    for period, shop_col, sold_col in [
        ("M-2", "Shop Stock Qty (M-2)", "Sold Qty (M-2)"),
        ("M-1", "Shop Stock Qty (M-1)", "Sold Qty (M-1)"),
        ("CM", "Shop Stock Qty (CM)", "Sold Qty (CM - MTD)"),
    ]:
        shop = _safe_numeric(df.get(shop_col, _safe_default_series(0, idx)))
        df[f"有貨店舖 ({period})"] = (shop > 0).astype(int)
        sold = _safe_numeric(df.get(sold_col, _safe_default_series(0, idx)))
        df[f"有賣店舖 ({period})"] = (sold > 0).astype(int)

    # 3. Eff Price (SKU 級，merged 回每行)
    ns2 = _safe_numeric(df.get("Net Sales (M-2)", _safe_default_series(0, idx)))
    ns1 = _safe_numeric(df.get("Net Sales (M-1)", _safe_default_series(0, idx)))
    nsm = _safe_numeric(df.get("Net Sales (CM - MTD)", _safe_default_series(0, idx)))
    sq2 = _safe_numeric(df.get("Sold Qty (M-2)", _safe_default_series(0, idx)))
    sq1 = _safe_numeric(df.get("Sold Qty (M-1)", _safe_default_series(0, idx)))
    sqm = _safe_numeric(df.get("Sold Qty (CM - MTD)", _safe_default_series(0, idx)))
    ep_data = pd.DataFrame({
        "SKU": df["SKU"],
        "_ns": ns2 + ns1 + nsm,
        "_sq": sq2 + sq1 + sqm,
    })
    ep_agg = ep_data.groupby("SKU", dropna=False).agg(
        _ns_sum=("_ns", "sum"), _sq_sum=("_sq", "sum"),
    ).reset_index()
    # 分母 0 → NaN
    ep_agg["Eff Price"] = np.where(ep_agg["_sq_sum"] > 0, ep_agg["_ns_sum"] / ep_agg["_sq_sum"], np.nan)
    df = df.merge(ep_agg[["SKU", "Eff Price"]], on="SKU", how="left")

    # 4. Price Tier
    df["Price Tier"] = df["Eff Price"].apply(_price_tier_label)

    # 5. SOH Tier 逐期
    for period in ["M-2", "M-1", "CM"]:
        shop = _safe_numeric(df.get(f"Shop Stock Qty ({period})", _safe_default_series(0, idx)))
        df[f"SOH Tier ({period})"] = shop.apply(_soh_tier_label)

    return df




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
    stock_threshold: int = 1000,
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
        progress_callback(0.55, "建構 SKU 層級分析中...")

    # 建立 SKU 基本資訊查找表（供 4 個 SKU-level 函式共用，避免重複 groupby）
    _sku_lookup = _build_sku_info_lookup(raw)

    # 平行化：同時計算 4 個 SKU-level 分析（彼此獨立，共用 _sku_lookup）
    with ThreadPoolExecutor(max_workers=4) as _sku_executor:
        f_alloc = _sku_executor.submit(_build_sku_allocation_summary, raw, _sku_lookup)
        f_diag = _sku_executor.submit(_build_sku_stock_diagnosis, raw, _sku_lookup)
        f_inv = _sku_executor.submit(_build_inventory_distribution_summary, raw, stock_threshold)
        f_monthly = _sku_executor.submit(_build_sku_monthly_analysis, raw, stock_threshold, _sku_lookup)

    sku_allocation = f_alloc.result()
    sku_diagnosis = f_diag.result()
    inv_distribution = f_inv.result()
    sku_monthly_analysis = f_monthly.result()

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
    """寫入 by-SKU 缺貨 Review 表（Tourist+Mix 缺貨 SKU，含 D001 貨量與店舖貨量）。

    D001 狀態分類用顏色標示：「倉庫有貨」(補貨問題) 用橘色警示、「源頭斷貨」用灰色。
    """
    ws.cell(row=1, column=1, value="SKU Stock Diagnosis — by-SKU 缺貨 Review（Tourist + Mix + Non Tourist 缺貨 SKU｜含 D001 倉庫貨量與店舖貨量）").font = Font(bold=True, size=12, color="1E40AF")

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
                              f"寫入 SKU 缺貨 Review ({row_idx-3}/{total_rows})...")

    if progress_callback:
        progress_callback(0.759, f"SKU 缺貨 Review 表寫入完成（{len(diag_df)} SKU）")

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


def _build_low_stock_item_class_summary(sku_monthly_analysis: pd.DataFrame) -> pd.DataFrame:
    if sku_monthly_analysis is None or sku_monthly_analysis.empty:
        return pd.DataFrame()
    if "低存貨" not in sku_monthly_analysis.columns:
        return pd.DataFrame()
    df = sku_monthly_analysis[sku_monthly_analysis["低存貨"]].copy()
    if df.empty:
        return pd.DataFrame()
    ic_col = "Item Class"
    if ic_col not in df.columns:
        return pd.DataFrame()
    df[ic_col] = df[ic_col].apply(
        lambda v: "New SKU" if pd.isna(v) or str(v).lower() in ("nan", "none", "nat", "") else str(v)
    )
    counts = df.groupby(ic_col, dropna=False).agg(
        低存貨SKU隻數=("SKU", "nunique")
    ).reset_index()
    counts[ic_col] = counts[ic_col].apply(
        lambda v: "New SKU" if pd.isna(v) or str(v).lower() in ("nan", "none", "nat", "") else str(v)
    )
    all_cats = sorted([str(c) for c in counts[ic_col].unique() if c not in ("New SKU", "總計")])
    all_cats += ["New SKU"] if "New SKU" in counts[ic_col].values else []
    all_cats += ["總計"]
    counts[ic_col] = pd.Categorical(counts[ic_col], categories=all_cats, ordered=True)
    counts = counts.sort_values(ic_col).reset_index(drop=True)
    total_val = int(counts["低存貨SKU隻數"].sum())
    total_row = pd.DataFrame({"Item Class": ["總計"], "低存貨SKU隻數": [total_val]})
    return pd.concat([counts, total_row], ignore_index=True)


def _write_stock_impact_dashboard(ws, raw_df, summary_buyer, inv_distribution, thin_border, desc_font, progress_callback, sku_monthly_analysis=None, stock_threshold=1000):
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
    # 曾缺貨 Tourist 行數（含已恢復，不限 CM 當期缺貨）
    tourist_oos_broader = int(((nature == "Tourist") & raw_df["Is Stock Impacted"]).sum()) if "Is Stock Impacted" in raw_df else 0

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
        ("🕐 曾缺貨Tourist行數(含已恢復)", tourist_oos_broader, "任何一期曾缺貨的Tourist RF行數｜含已恢復供貨", "6366F1"),
    ]

    low_stock_total = 0
    low_stock_item_class = pd.DataFrame()
    if sku_monthly_analysis is not None and not sku_monthly_analysis.empty and "低存貨" in sku_monthly_analysis.columns:
        low_stock_total = int(sku_monthly_analysis["低存貨"].sum())
        low_stock_item_class = _build_low_stock_item_class_summary(sku_monthly_analysis)

    row4_metrics = [
        ("🔴 低存貨 SKU 總數", low_stock_total, f"D001+店舖存貨<{stock_threshold}件｜單位：SKU數", "DC2626"),
    ]

    _write_metric_cards(ws, "缺貨銷售影響分析摘要 (Stock Impact Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback, row3_metrics=row3_metrics, row4_metrics=row4_metrics)

    # Top 5 受影響 Brand（按 Tourist 缺貨）
    if summary_buyer is not None and not summary_buyer.empty and len(summary_buyer) > 1:
        ws.cell(row=19, column=1, value="Top 5 受缺貨影響最嚴重的 Buyer（按 Tourist 缺貨排序）").font = Font(bold=True, size=12, color="1E40AF")
        top_buyers = summary_buyer[summary_buyer["Buyer"] != "總計"].head(5)
        for col_idx, h in enumerate(top_buyers.columns, 1):
            cell = ws.cell(row=20, column=col_idx, value=h)
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, (_, row_data) in enumerate(top_buyers.iterrows()):
            excel_row = 21 + row_idx
            for col_idx, h in enumerate(top_buyers.columns, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"

    # 低存貨 SKU 按 Item Class 分類
    table_start = 27
    if low_stock_item_class is not None and not low_stock_item_class.empty:
        ws.cell(row=table_start, column=1, value="🔴 低存貨 SKU 按 Item Class 分類").font = Font(bold=True, size=12, color="DC2626")
        header_row = table_start + 1
        for col_idx, h in enumerate(low_stock_item_class.columns, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, (_, row_data) in enumerate(low_stock_item_class.iterrows()):
            excel_row = header_row + 1 + row_idx
            is_total = row_data.get("Item Class") == "總計"
            for col_idx, h in enumerate(low_stock_item_class.columns, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
                if is_total:
                    cell.font = Font(bold=True)
    else:
        ws.cell(row=table_start, column=1, value="🔴 低存貨 SKU 按 Item Class 分類：無低存貨 SKU").font = Font(size=12, color="DC2626")

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
    stock_threshold: int = 1000,
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
        ws_sku_monthly.cell(row=2, column=1, value="各 SKU 在 Tourist/Mix/Non Tourist 店舖的月均銷售（Shop Monthly Sold Qty P90D 已是月均）與當期庫存（CM Shop Stock）。低存貨 SKU 優先顯示。").font = Font(size=10, italic=True, color="6B7280")
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

    # Sheet 7: SKU Stock Diagnosis (by-SKU 缺貨 Review, 含 D001 + 店舖貨量)
    if progress_callback:
        progress_callback(0.70, "建立 by-SKU 缺貨 Review 工作表...")
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
                                  thin_border, desc_font, progress_callback, sku_monthly_analysis=sku_monthly_analysis, stock_threshold=stock_threshold)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf


# ─── Safety Stock x NetStock Review (調貨事件 Review Safety 設定) Functions ──────


_HYBRID_REQUIRED = [
    "Buyer", "Brand", "SKU",
    "Shop Nature", "Site Code", "Shop Stock Qty (CM)",
    "Sold Qty (P90D)", "Adj. Safety Stock",
    "Shop Monthly Sold Qty (P90D)",
]

# ── 調貨事件 Review（Safety 基準天數按 Nature；調貨段仍為 14 日）──
_HYBRID_EVENT_NEED_IN = "需要調入"
_HYBRID_EVENT_DONATED = "實際轉出"
_HYBRID_EVENT_DONOR = "可轉出未使用"
_HYBRID_DIAG_UNDER = "Safety 設定不足"
_HYBRID_DIAG_OVER = "Safety 設定過高"
_HYBRID_DIAG_NORMAL_NEED = "設定正常—缺貨非設定因素"
_HYBRID_DIAG_NORMAL_DONOR = "設定正常—過剩非設定因素"
_HYBRID_DIAG_ORDER = [
    _HYBRID_DIAG_UNDER, _HYBRID_DIAG_OVER,
    _HYBRID_DIAG_NORMAL_NEED, _HYBRID_DIAG_NORMAL_DONOR,
]
_HYBRID_EVENT_ORDER = {_HYBRID_EVENT_NEED_IN: 0, _HYBRID_EVENT_DONATED: 1, _HYBRID_EVENT_DONOR: 2}
_HYBRID_DIAG_NATURES = ["Tourist", "Mix", "Non Tourist"]

# ── 分類維度（需求方過濾與彙總分組；TurnDay 池與 Safety 基準天數恆按 Shop Nature）──
# nature：以 Shop Nature 分類（預設；需求方 Tourist-RF）
# rp：以 RP Type 分類（需求方任何 Nature 的 RF 列，供應方維持 Mix／Non Tourist，允許同 Nature 互調）
# both：以 Shop Nature＋RP Type 分類（6 組；計算同 nature 模式，僅分組顯示不同）
_HYBRID_DIMENSIONS = ("nature", "rp", "both")
_HYBRID_DIMENSION_LABELS = {"nature": "Shop Nature", "rp": "RP Type", "both": "Shop Nature＋RP Type"}
_HYBRID_RP_GROUPS = ("RF", "ND")
_HYBRID_COMBINED_GROUPS = tuple(
    f"{n}-{r}" for n in ("Tourist", "Mix", "Non Tourist") for r in _HYBRID_RP_GROUPS
)


def _hybrid_dimension(dimension) -> str:
    """分類維度正規化：非法值退回預設 nature，不拋錯（沿用 _hybrid_safety_days 風格）。"""
    d = str(dimension or "nature").strip().lower()
    return d if d in _HYBRID_DIMENSIONS else "nature"


def _hybrid_dimension_groups(dimension) -> tuple[str, ...]:
    """分類維度對應的彙總分組名稱（Summary／交叉表欄位按此切換）。"""
    d = _hybrid_dimension(dimension)
    if d == "rp":
        return _HYBRID_RP_GROUPS
    if d == "both":
        return _HYBRID_COMBINED_GROUPS
    return tuple(_HYBRID_DIAG_NATURES)

# ── Safety 基準周轉天數（Review 線；與調貨段的 14 日線分開，Safety 不跟隨存貨周轉）──
_HYBRID_SAFETY_DAYS_DEFAULT = {"Tourist": 7, "Mix": 9, "Non Tourist": 10}
_HYBRID_SAFETY_DAYS_MIN = 7
_HYBRID_SAFETY_DAYS_MAX = 10
# Review 欄名（改自「14日周轉所需量」；數值隨所屬 Nature 的基準天數而變）
_HYBRID_SAFETY_LINE_COLUMN = "Safety 周轉所需量"
# 每列實際採用的基準天數（Echo 自參數，供導出檔稽核與回退彙總）
_HYBRID_SAFETY_DAYS_COLUMN = "Safety 基準天數"


def _hybrid_safety_days(safety_days: dict | None = None) -> tuple[dict[str, int], list[str]]:
    """Safety 基準天數 — 以預設值為底，覆寫傳入的整數值。

    未知 Nature、非數值維持忽略（退回預設）；0 以下拒收；範圍外數值箝制進
    7–10 並回報 warning；非整數浮點直接拒絕，避免靜默截斷。
    回傳 (實際天數, warnings)。
    """
    days = dict(_HYBRID_SAFETY_DAYS_DEFAULT)
    warnings: list[str] = []
    for key, val in (safety_days or {}).items():
        name = str(key).strip()
        if name not in days:
            continue
        if isinstance(val, bool):
            raise ValueError(f"Safety 基準天數「{name}」必須為整數，不得為布林值。")
        try:
            numeric = float(val)
        except (TypeError, ValueError):
            continue
        if not np.isfinite(numeric) or not numeric.is_integer():
            raise ValueError(f"Safety 基準天數「{name}」必須為整數。")
        n = int(numeric)
        if n <= 0:
            continue
        if n < _HYBRID_SAFETY_DAYS_MIN or n > _HYBRID_SAFETY_DAYS_MAX:
            clamped = min(max(n, _HYBRID_SAFETY_DAYS_MIN), _HYBRID_SAFETY_DAYS_MAX)
            warnings.append(
                f"Safety 基準天數「{name}」輸入 {n} 日超出建議範圍，已箝制為 {clamped} 日。"
            )
            days[name] = clamped
            continue
        days[name] = n
    return days, warnings


def _hybrid_safety_days_label(days: dict[str, int]) -> str:
    """基準天數的文字描述（供指標卡說明、warning 與 UI 使用）。"""
    return " / ".join(f"{n} {int(days[n])} 日" for n in _HYBRID_DIAG_NATURES)


def _hybrid_numeric_column(
    df: pd.DataFrame,
    column: str,
    index,
    *,
    allow_negative: bool = False,
    nonfinite_as_missing: bool = False,
) -> pd.Series:
    """Hybrid 專用數值欄位清理（只在 hybrid 路徑使用，不動全域 _safe_numeric）。

    空白／非數值保留為 NaN 供各欄 fallback；負數與非有限值依政策處理：
    allow_negative=False 時負數直接拒絕；nonfinite_as_missing=True 時非有限值
    視為缺值（供 D001 fallback），否則拒絕。
    """
    values = pd.to_numeric(df.get(column, _safe_default_series(np.nan, index)), errors="coerce")
    raw_values = values.to_numpy(dtype=float)
    nonfinite_mask = ~np.isnan(raw_values) & ~np.isfinite(raw_values)
    if nonfinite_mask.any():
        if nonfinite_as_missing:
            values = values.mask(pd.Series(nonfinite_mask, index=values.index))
        else:
            raise ValueError(f"「{column}」包含無限大或其他非有限數值，請修正來源資料後再執行。")
    if not allow_negative and bool(((raw_values < 0) & ~np.isnan(raw_values)).any()):
        raise ValueError(f"「{column}」不可為負數，請修正來源資料後再執行。")
    return values

# ── Safety 調整統計（欄名屬新契約，未經要求不得改動）──
_HYBRID_SAFETY_CHANGE_COLUMNS = ["分類", "列數", "SKU 數", "Safety 建議調整量合計", "說明"]
_HYBRID_SAFETY_CHANGE_ORDER = ["增加", "減少", "不變"]
_HYBRID_SAFETY_CHANGE_DESC = {
    "增加": "建議 Safety 設定值高於現有設定（Safety 設定不足）",
    "減少": "建議 Safety 設定值低於現有設定（Safety 設定過高）",
    "不變": "設定正常或四捨五入後無需調整（缺貨／過剩非 Safety 設定因素）",
}
_HYBRID_SAFETY_CHANGE_TOTAL_DESC = "SKU 數為全表不重複 SKU（同一 SKU 可跨分類）；合計為淨調整量"
_HYBRID_SAFETY_SKU_COLUMNS = [
    "Brand", "SKU", "SKU Desc", "Buyer",
    "Review 列數", "需增加店舖數", "需減少店舖數", "設定正常店舖數",
    "現有 Safety 設定合計（調整列）", "建議 Safety 設定合計（調整列）", "淨調整量",
]

# 「調貨 Review Safety Stock」主表欄序（店舖×SKU 逐列；RAW 則輸出全部欄位 + Review 欄）
# RP Type 為正規化後值（RF／ND），來源缺「RP Type」欄時該欄不出現（主表退回 25 欄）
_HYBRID_DIAG_COLUMNS = [
    "Brand", "SKU", "SKU Desc", "Buyer", "Adj. SST", "Site Code", "OM", "Shop Nature", "RP Type",
    "事件類型", "Shop Stock Qty (CM)", "Shop Monthly Sold Qty (P90D)",
    _HYBRID_SAFETY_DAYS_COLUMN, _HYBRID_SAFETY_LINE_COLUMN, "Adj. Safety Stock", "Safety Review",
    "建議 Safety 設定值", "Safety 建議調整量",
    "調入需求量 (補到14日)", "實際調入量", "調入缺口",
    "可轉出貨量", "轉出量", "調貨來源 (Site Code)",
    "調貨後庫存", "調貨後 TurnDay",
]


def validate_hybrid_input(df_input: pd.DataFrame) -> list[str]:
    """混合模式必要欄位（含基礎四欄）：缺任一即回報，由呼叫方 raise。"""
    avail = set(df_input.columns)
    return [name for name in _HYBRID_REQUIRED if _resolve_column_name(name, avail) is None]


def _canonicalize_hybrid_columns(df_input: pd.DataFrame) -> pd.DataFrame:
    """混合模式入口：將所有字串欄名 strip 後統一運算（只偵測不改名的情況，
    例如「Buyer 」帶前後空格時先通過驗證、後續取值卻 KeyError）。"""
    df = df_input.copy()
    mapping: dict[str, str] = {}
    counts: dict[str, int] = {}
    for col in df.columns:
        if isinstance(col, str):
            new = col.strip()
            if new:
                mapping[col] = new
                counts[new] = counts.get(new, 0) + 1
    dup = [k for k, v in counts.items() if v > 1]
    if dup:
        raise ValueError(f"欄位名稱去空白後重複: {', '.join(dup)}")
    if mapping:
        df = df.rename(columns=mapping)
    return df


def _normalize_hybrid_nature(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Shop Nature 正規化（混合模式固定規則）：
    strip + 小寫後 Tourist→Tourist、Mix→Mix；其餘一律（含 Other、空白、大小寫差異）→ Non Tourist。
    回傳 (已正規化 df, 被映射的非標準列數)——映射列數只計「無法辨識」的值
    （小寫去除空白後不在 tourist / mix / non tourist 三者內，如 Other、空白、未知值）。
    """
    df = df.copy()
    raw = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    low = raw.str.lower()
    standard = low.isin(["tourist", "mix", "non tourist"])
    mapped = int((~standard).sum())
    out = pd.Series(
        np.select([low == "mix", low == "tourist"], ["Mix", "Tourist"], default="Non Tourist"),
        index=df.index,
    )
    df["Shop Nature"] = out
    return df, mapped


def _normalize_hybrid_rp_type(df: pd.DataFrame) -> tuple[pd.DataFrame, int, bool]:
    """RP Type 正規化（混合模式固定規則；RF＝可補貨、ND＝不可補貨）：
    strip + 大寫後 RF→RF、ND→ND；其餘（含 Other、空白、大小寫差異）一律→ ND
    （保守歸類為不可補貨：不作調貨需求方，但仍可作供應方）。
    回傳 (已正規化 df, 被映射的非標準列數, 來源是否含「RP Type」欄)——
    欄缺失時不改 df（RAW 不虛構欄位），由呼叫方退回 nature 維度並列 warning。
    """
    if "RP Type" not in df.columns:
        return df, 0, False
    df = df.copy()
    raw = df["RP Type"].fillna("").astype(str).str.strip().str.upper()
    known = raw.isin(["RF", "ND"])
    mapped = int((~known).sum())
    df["RP Type"] = np.where(raw == "RF", "RF", "ND")
    return df, mapped, True


def _transfer_input_consistency_warnings(df: pd.DataFrame) -> list[str]:
    """調貨模式輸入一致性檢核（調貨 Review 與 Stock Checking 共用，只列 warning 不拒絕）：

    - 「Site Code × SKU」重複列：正常輸入粒度為一店×一SKU一列，重複列會令
      同一店舖的庫存與銷量被多次計入；
    - 同一店舖的「Shop Nature」在不同列之間不一致：Shop Nature 應為店舖層級屬性，
      同一店舖不應同時出現 Tourist 與非 Tourist 列；
      調貨分配已排除同店來源，自我調貨（調貨來源＝自己店舖）不會發生。
    - OM 欄存在但全欄空白：同OM分組無法區分店群，結果會等同全部店舖使用同一群組。
    """
    out: list[str] = []
    if "Site Code" not in df.columns or "SKU" not in df.columns:
        return out
    site = df["Site Code"].fillna("").astype(str).str.strip()
    sku = df["SKU"].fillna("").astype(str).str.strip()
    key = site + "\u0000" + sku
    sizes = key.groupby(key).size()
    dup_sizes = sizes[sizes > 1]
    if len(dup_sizes):
        out.append(
            f"偵測到「Site Code × SKU」重複列 {len(dup_sizes)} 組（共 {int(dup_sizes.sum())} 列；"
            "正常輸入應為一店×一SKU一列），重複列會令同一店舖的庫存與銷量被多次計入；"
            "重複列的 Site Code 不同時，自我調貨守衛（以 Site Code 比對）無法辨識同店，"
            "同一實體店舖可能同時出現在調貨來源與接收端，請檢查來源檔案。"
        )
    nature = df.get("Shop Nature", _safe_default_series("", df.index))
    nature_norm = nature.fillna("").astype(str).str.strip().str.lower()
    n_unique = nature_norm.groupby(site).nunique()
    bad_sites = n_unique[n_unique > 1]
    if len(bad_sites):
        shown = "、".join(str(s) for s in list(bad_sites.index[:5]))
        suffix = f" 等 {len(bad_sites)} 間店舖" if len(bad_sites) > 5 else ""
        out.append(
            f"店舖 {shown}{suffix} 的「Shop Nature」在不同列之間不一致；"
            "Shop Nature 應為店舖層級固定屬性，請檢查來源資料。"
        )
    if "OM" in df.columns:
        om = df["OM"].fillna("").astype(str).str.strip()
        if len(df) and bool((om == "").all()):
            out.append("來源檔案的「OM」欄全欄空白，同OM調貨結果會等同全部店舖使用同一群組。")
    return out


def _build_hybrid_diagnosis(raw: pd.DataFrame, variant: str = "free") -> pd.DataFrame:
    """調貨 Review 主表（店舖×SKU 逐列）：只收錄事件類型非空的列，
    按 Brand / SKU / 事件類型 / Site Code 排序；「可轉出貨量」為三 Nature 可轉出欄加總
    （Tourist 欄僅雙向時存在）。
    variant="om" 時改用同OM調貨欄位並重算事件類型（轉出量=0 的來源列改列「可轉出未使用」）；
    Safety Review 與建議調整欄兩變體相同。"""
    if raw is None or raw.empty or "事件類型" not in raw.columns:
        return pd.DataFrame(columns=_HYBRID_DIAG_COLUMNS)
    df = raw[raw["事件類型"].fillna("").astype(str) != ""].copy()
    if df.empty:
        return pd.DataFrame(columns=_HYBRID_DIAG_COLUMNS)
    if variant == "om":
        renames = {
            "實際調入量 (同OM)": "實際調入量", "調入缺口 (同OM)": "調入缺口",
            "轉出量 (同OM)": "轉出量", "調貨來源 (同OM)": "調貨來源 (Site Code)",
            "調貨後庫存 (同OM)": "調貨後庫存", "調貨後 TurnDay (同OM)": "調貨後 TurnDay",
        }
        # 先移除同名的不限OM欄，避免 rename 後重複
        df = df.drop(columns=[v for v in renames.values() if v in df.columns])
        df = df.rename(columns={k: v for k, v in renames.items() if k in df.columns})
        donated = _safe_numeric(df.get("轉出量", _safe_default_series(0, df.index))) > 0
        df["事件類型"] = np.where(
            df["事件類型"] == _HYBRID_EVENT_NEED_IN, _HYBRID_EVENT_NEED_IN,
            np.where(donated, _HYBRID_EVENT_DONATED, _HYBRID_EVENT_DONOR),
        )
    df["可轉出貨量"] = (
        _safe_numeric(df.get("可轉出貨量 (Mix)", _safe_default_series(0, df.index)))
        + _safe_numeric(df.get("可轉出貨量 (Non Tourist)", _safe_default_series(0, df.index)))
        + _safe_numeric(df.get("可轉出貨量 (Tourist)", _safe_default_series(0, df.index)))
    ).astype(int)
    df["_evt_key"] = df["事件類型"].map(_HYBRID_EVENT_ORDER).fillna(9)
    df = df.sort_values(["Brand", "SKU", "_evt_key", "Site Code"], kind="stable")
    ordered = [c for c in _HYBRID_DIAG_COLUMNS if c in df.columns]
    return df[ordered].reset_index(drop=True)


# 簡易調貨明細欄序（一列一筆「來源店舖 → 接收店舖」調貨；RP Type 緊接對應 Shop Name，
# 調貨數量置於 Supply Site Code 前作 Target／Supply 分隔；v1.22.0 起 12 欄）
_HYBRID_TRANSFER_DETAIL_COLUMNS = [
    "Brand", "SKU", "SKU Desc",
    "Target Site Code(接收店舖)", "Target Shop Name", "Target RP Type", "Target OM",
    "調貨數量",
    "Supply Site Code(調貨來源)", "Supply Shop Name", "Supply RP Type", "Supply OM",
]


def _build_hybrid_transfer_detail(raw: pd.DataFrame, variant: str = "free") -> pd.DataFrame:
    """簡易調貨明細：只收錄「需要調貨=是」且實際有調入的列，把調貨來源字串
    （如「HB30:2; HB49:3」）逐段拆行，一列一筆來源店→接收店調貨。
    variant="om" 時改用同OM調貨欄位；同一接收列同一來源店重複出現時合併加總。
    來源店名稱／OM／RP Type 以 RAW 全量建立 Site Code 查找表補回，查不到留空。"""
    cols = _HYBRID_TRANSFER_DETAIL_COLUMNS
    if raw is None or raw.empty or "需要調貨" not in raw.columns:
        return pd.DataFrame(columns=cols)
    if variant == "om":
        src_col, qty_col = "調貨來源 (同OM)", "實際調入量 (同OM)"
    else:
        src_col, qty_col = "調貨來源 (Site Code)", "實際調入量"
    if src_col not in raw.columns or qty_col not in raw.columns:
        return pd.DataFrame(columns=cols)

    need = raw["需要調貨"].fillna("").astype(str).str.strip() == "是"
    src = raw[src_col].fillna("").astype(str).str.strip()
    qty_actual = _safe_numeric(raw[qty_col])
    sel = raw[need & (src != "") & (qty_actual > 0)]
    if sel.empty:
        return pd.DataFrame(columns=cols)

    site_all = raw.get("Site Code", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    if "Shop Name" in raw.columns:
        name_all = raw["Shop Name"].fillna("").astype(str).str.strip()
    elif "Item_Shop" in raw.columns:
        name_all = pd.Series(
            [_extract_shop_name(raw["Item_Shop"].iloc[i], site_all.iloc[i]) for i in range(len(raw))],
            index=raw.index,
        )
    else:
        name_all = _safe_default_series("", raw.index)
    om_all = raw.get("OM", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    rp_all = (
        raw["RP Type"].fillna("").astype(str).str.strip().str.upper()
        if "RP Type" in raw.columns
        else _safe_default_series("", raw.index)
    )

    name_map: dict[str, str] = {}
    om_map: dict[str, str] = {}
    rp_map: dict[str, str] = {}
    for site_v, name_v, om_v, rp_v in zip(
        site_all.to_numpy(), name_all.to_numpy(), om_all.to_numpy(), rp_all.to_numpy()
    ):
        if site_v == "":
            continue
        if name_v and site_v not in name_map:
            name_map[site_v] = str(name_v)
        if om_v and site_v not in om_map:
            om_map[site_v] = str(om_v)
        if rp_v and site_v not in rp_map:
            rp_map[site_v] = str(rp_v)

    rows = []
    for idx_pos, row in sel.iterrows():
        site_v = str(row.get("Site Code", "")).strip()
        target_name = str(row.get("Shop Name", "")).strip() if "Shop Name" in sel.columns else ""
        if not target_name:
            target_name = name_map.get(site_v, "")
        target_om = str(row.get("OM", "")).strip() or om_map.get(site_v, "")
        target_rp = str(row.get("RP Type", "")).strip().upper() or rp_map.get(site_v, "")
        per_supply: dict[str, int] = {}
        for part in str(row[src_col]).split(";"):
            part = part.strip()
            if not part:
                continue
            supply_site, sep, qty_text = part.rpartition(":")
            if not sep:
                continue
            supply_site = supply_site.strip()
            try:
                qty_n = int(float(qty_text.strip()))
            except ValueError:
                continue
            if not supply_site or qty_n <= 0:
                continue
            per_supply[supply_site] = per_supply.get(supply_site, 0) + qty_n
        for supply_site, qty_n in per_supply.items():
            rows.append({
                "Brand": row.get("Brand", ""),
                "SKU": row.get("SKU", ""),
                "SKU Desc": row.get("SKU Desc", ""),
                "Target Site Code(接收店舖)": site_v,
                "Target Shop Name": target_name,
                "Target RP Type": target_rp,
                "Target OM": target_om,
                "調貨數量": qty_n,
                "Supply Site Code(調貨來源)": supply_site,
                "Supply Shop Name": name_map.get(supply_site, ""),
                "Supply RP Type": rp_map.get(supply_site, ""),
                "Supply OM": om_map.get(supply_site, ""),
            })
    if not rows:
        return pd.DataFrame(columns=cols)
    out = pd.DataFrame(rows, columns=cols)
    out = out.sort_values(
        ["Brand", "SKU", "Target Site Code(接收店舖)", "Supply Site Code(調貨來源)"],
        kind="stable",
    ).reset_index(drop=True)
    return out


def _hybrid_diag_summary_columns(dimension: str = "nature") -> list[str]:
    cols = ["Safety Review"]
    for g in _hybrid_dimension_groups(dimension):
        cols += [f"{g} 列數", f"{g} SKU 數", f"{g} 建議調整量合計"]
    cols += ["合計 列數", "合計 SKU 數", "合計 建議調整量"]
    return cols


def _hybrid_group_key_series(df: pd.DataFrame, dimension: str) -> pd.Series:
    """依分類維度產生每列的分組鍵（nature→Shop Nature；rp→RP Type；both→Nature-RP）。
    RP Type 為正規化後值；缺欄時一律按 RF（可補貨）處理，與 transform_hybrid 同口徑。"""
    d = _hybrid_dimension(dimension)
    nature = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    if d == "nature":
        return nature
    rp = (
        df["RP Type"].fillna("").astype(str).str.strip().str.upper()
        if "RP Type" in df.columns
        else _safe_default_series("RF", df.index)
    )
    rp = pd.Series(np.where(rp == "RF", "RF", "ND"), index=df.index)
    if d == "rp":
        return rp
    return nature + "-" + rp


def _build_hybrid_diag_summary(raw: pd.DataFrame, dimension: str = "nature") -> pd.DataFrame:
    """Safety Review × 分類維度（預設 Shop Nature）交叉表（列數 / SKU 數 / 建議調整量合計），含總計列。"""
    cols = _hybrid_diag_summary_columns(dimension)
    groups = _hybrid_dimension_groups(dimension)
    if raw is None or raw.empty or "事件類型" not in raw.columns:
        return pd.DataFrame(columns=cols)
    df = raw[raw["事件類型"].fillna("").astype(str) != ""].copy()
    if df.empty:
        return pd.DataFrame(columns=cols)
    df["_chg"] = pd.to_numeric(
        df.get("Safety 建議調整量", _safe_default_series(np.nan, df.index)), errors="coerce"
    ).fillna(0)
    grp = _hybrid_group_key_series(df, dimension)
    diag = df["Safety Review"].fillna("").astype(str)

    def _agg(sub: pd.DataFrame) -> dict:
        return {
            "列數": int(len(sub)),
            "SKU 數": int(sub["SKU"].nunique()) if "SKU" in sub.columns else 0,
            # 先四捨五入再取整，與「Safety 調整方向分類」／stats 合計口徑一致
            "建議調整量合計": int(round(float(sub["_chg"].sum()))),
        }

    _total_key = {"列數": "合計 列數", "SKU 數": "合計 SKU 數", "建議調整量合計": "合計 建議調整量"}

    rows = []
    for d in _HYBRID_DIAG_ORDER:
        sub_d = df[diag == d]
        row = {"Safety Review": d}
        for g in groups:
            for k, v in _agg(sub_d[grp.loc[sub_d.index] == g]).items():
                row[f"{g} {k}"] = v
        for k, v in _agg(sub_d).items():
            row[_total_key[k]] = v
        rows.append(row)
    total_row = {"Safety Review": "總計"}
    for g in groups:
        for k, v in _agg(df[grp == g]).items():
            total_row[f"{g} {k}"] = v
    for k, v in _agg(df).items():
        total_row[_total_key[k]] = v
    rows.append(total_row)
    return pd.DataFrame(rows, columns=cols)


def _hybrid_safety_event_frame(raw: pd.DataFrame) -> pd.DataFrame:
    """Safety 調整統計母體 — 僅「事件類型」非空的列（與 Safety Review × Shop Nature
    交叉表同母體）。Safety 建議不受 OM 限制影響，故不限OM／同OM 共用同一份統計。"""
    if raw is None or raw.empty or "事件類型" not in raw.columns:
        return pd.DataFrame()
    return raw[raw["事件類型"].fillna("").astype(str) != ""].copy()


def _first_non_blank_value(values) -> str:
    """取群組內首個非空值（SKU Desc／Buyer 每 SKU 應一致，容許個別列空白）。"""
    for v in values:
        text = "" if pd.isna(v) else str(v).strip()
        if text:
            return text
    return ""


def _build_hybrid_safety_change_summary(raw: pd.DataFrame) -> pd.DataFrame:
    """Safety 調整方向分類表 — 按「Safety 建議調整量」（顯示值，已四捨五入至 1 位小數）
    的正負分為 增加／減少／不變，統計列數、SKU 數與調整量合計，含總計列。
    邊界情形：調整量四捨五入後為 0 的列歸「不變」（其「Safety Review」仍可能標示不足／過高）。
    總計列的 SKU 數用母體不重複 SKU，避免同一 SKU 跨分類虛增。"""
    cols = _HYBRID_SAFETY_CHANGE_COLUMNS
    df = _hybrid_safety_event_frame(raw)
    if df.empty:
        return pd.DataFrame(columns=cols)
    chg = pd.to_numeric(
        df.get("Safety 建議調整量", _safe_default_series(np.nan, df.index)), errors="coerce"
    ).fillna(0)
    sku = df.get("SKU", _safe_default_series("", df.index))
    cat = np.where(chg > 0, "增加", np.where(chg < 0, "減少", "不變"))

    rows = []
    for label in _HYBRID_SAFETY_CHANGE_ORDER:
        m = cat == label
        rows.append({
            "分類": label,
            "列數": int(m.sum()),
            "SKU 數": int(sku[m].dropna().nunique()),
            "Safety 建議調整量合計": int(round(float(chg[m].sum()))),
            "說明": _HYBRID_SAFETY_CHANGE_DESC[label],
        })
    rows.append({
        "分類": "總計",
        "列數": int(len(df)),
        "SKU 數": int(sku.dropna().nunique()),
        "Safety 建議調整量合計": int(round(float(chg.sum()))),
        "說明": _HYBRID_SAFETY_CHANGE_TOTAL_DESC,
    })
    return pd.DataFrame(rows, columns=cols)


def _build_hybrid_safety_sku_detail(raw: pd.DataFrame) -> pd.DataFrame:
    """Safety 調整明細（SKU 級彙總）— 每個有調貨事件的 SKU 一列：
    Review 列數、需增加／需減少／設定正常店舖數（按 Site Code 不重複計）、
    調整列（設定不足＋設定過高）的現有與建議 Safety 設定合計，以及淨調整量。
    只有「設定正常」列的 SKU 仍會納入（淨調整量為 0，代表缺貨／過剩非 Safety 設定因素）。"""
    cols = _HYBRID_SAFETY_SKU_COLUMNS
    df = _hybrid_safety_event_frame(raw)
    if df.empty:
        return pd.DataFrame(columns=cols)

    diag = df.get("Safety Review", _safe_default_series("", df.index)).fillna("").astype(str)
    chg = pd.to_numeric(
        df.get("Safety 建議調整量", _safe_default_series(np.nan, df.index)), errors="coerce"
    ).fillna(0)
    adj = pd.to_numeric(
        df.get("Adj. Safety Stock", _safe_default_series(np.nan, df.index)), errors="coerce"
    ).fillna(0)
    # 「建議 Safety 設定值」只落在不足／過高列，故其非空即代表該列需要調整
    sugg = pd.to_numeric(
        df.get("建議 Safety 設定值", _safe_default_series(np.nan, df.index)), errors="coerce"
    )
    work = pd.DataFrame({
        "Brand": df.get("Brand", _safe_default_series("", df.index)).fillna("").astype(str),
        "SKU": df.get("SKU", _safe_default_series("", df.index)),
        "SKU Desc": df.get("SKU Desc", _safe_default_series("", df.index)),
        "Buyer": df.get("Buyer", _safe_default_series("", df.index)),
        "_site": df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip(),
        "_chg": chg,
        "_adj": adj,
        "_sugg": sugg,
        "_under": (diag == _HYBRID_DIAG_UNDER).to_numpy(),
        "_over": (diag == _HYBRID_DIAG_OVER).to_numpy(),
    })

    rows = []
    for (_brand, _sku), grp in work.groupby(["Brand", "SKU"], sort=False, dropna=False):
        under = grp["_under"]
        over = grp["_over"]
        fix = grp[grp["_sugg"].notna()]
        rows.append({
            "Brand": _brand,
            "SKU": _sku,
            "SKU Desc": _first_non_blank_value(grp["SKU Desc"]),
            "Buyer": _first_non_blank_value(grp["Buyer"]),
            "Review 列數": int(len(grp)),
            "需增加店舖數": int(grp.loc[under, "_site"].nunique()),
            "需減少店舖數": int(grp.loc[over, "_site"].nunique()),
            "設定正常店舖數": int(grp.loc[~under & ~over, "_site"].nunique()),
            "現有 Safety 設定合計（調整列）": int(round(float(fix["_adj"].sum()))),
            "建議 Safety 設定合計（調整列）": int(round(float(fix["_sugg"].sum()))),
            "淨調整量": int(round(float(grp["_chg"].sum()))),
        })
    out = pd.DataFrame(rows, columns=cols)
    return out.sort_values(["Brand", "SKU"], kind="stable").reset_index(drop=True)


# ── Buyer 統計（欄名屬新契約，未經要求不得改動）──
_HYBRID_BUYER_COMBINED_TITLE = "Buyer 統計（需調貨＋需調整 Safety）"


def _hybrid_buyer_need_stats(raw: pd.DataFrame, need_series: pd.Series | None = None) -> list[tuple[str, int]]:
    """Buyer 級統計：各 Buyer 涉及多少個「需要調貨」SKU（已刪除重複）。

    與 Stock Checking 的 `_stock_checking_buyer_stats` 同口徑（Buyer 取同 SKU 首個非空值；
    全空歸入「(未填 Buyer)」），供 Safety Stock x NetStock Review 沿用。
    回傳 [(Buyer, 需調貨 SKU 數)]，按數量降序、同數按 Buyer 名排序。
    """
    if raw is None or raw.empty or "SKU" not in raw.columns:
        return []
    if need_series is None:
        need_series = raw.get("需要調貨", _safe_default_series("", raw.index)).astype(str) == "是"
    need_skus = set(raw.loc[need_series, "SKU"].dropna().unique())
    if not need_skus:
        return []
    buyer = (
        raw["Buyer"].fillna("").astype(str).str.strip()
        if "Buyer" in raw.columns
        else _safe_default_series("", raw.index)
    )
    tmp = pd.DataFrame({"_sku": raw["SKU"], "_buyer": buyer})
    tmp = tmp[tmp["_sku"].isin(need_skus)]

    def _first_nonempty(s: pd.Series) -> str:
        s = s[s != ""]
        return s.iloc[0] if len(s) else "(未填 Buyer)"

    per_sku_buyer = tmp.groupby("_sku", sort=False)["_buyer"].agg(_first_nonempty)
    counts = per_sku_buyer.value_counts()
    return sorted(((str(b), int(c)) for b, c in counts.items()), key=lambda kv: (-kv[1], kv[0]))


def _hybrid_buyer_safety_stats(raw: pd.DataFrame) -> list[tuple[str, int, int]]:
    """Buyer 級統計：各 Buyer 涉及多少個「需要修改 Safety」SKU（已刪除重複）。

    母體為 Safety Review 屬「Safety 設定不足」或「Safety 設定過高」的列（即有
    「建議 Safety 設定值」的列，與 Safety 調整方向分類的增加＋減少同母體）。
    Buyer 取同 SKU 首個非空值（與需調貨統計同口徑）；全空歸入「(未填 Buyer)」。
    調整量合計為該 Buyer 名下各需調整 SKU 的「Safety 建議調整量」列合計（四捨五入取整，
    與交叉表／分類表同口徑；正=增加、負=減少）。
    回傳 [(Buyer, 需調整 Safety SKU 數, Safety 建議調整量合計)]，
    按 SKU 數降序、同數按 Buyer 名排序。
    """
    if raw is None or raw.empty or "SKU" not in raw.columns:
        return []
    diag = raw.get("Safety Review", _safe_default_series("", raw.index)).fillna("").astype(str)
    fix_mask = diag.isin([_HYBRID_DIAG_UNDER, _HYBRID_DIAG_OVER])
    if not bool(fix_mask.any()):
        return []
    chg = pd.to_numeric(
        raw.get("Safety 建議調整量", _safe_default_series(np.nan, raw.index)), errors="coerce"
    ).fillna(0.0)
    buyer = (
        raw["Buyer"].fillna("").astype(str).str.strip()
        if "Buyer" in raw.columns
        else _safe_default_series("", raw.index)
    )
    tmp = pd.DataFrame({
        "_sku": raw.loc[fix_mask, "SKU"],
        "_buyer": buyer.loc[fix_mask],
        "_chg": chg.loc[fix_mask],
    })
    tmp = tmp[tmp["_sku"].notna()]
    if tmp.empty:
        return []

    def _first_nonempty(s: pd.Series) -> str:
        s = s[s != ""]
        return s.iloc[0] if len(s) else "(未填 Buyer)"

    per_sku_buyer = tmp.groupby("_sku", sort=False)["_buyer"].agg(_first_nonempty)
    per_sku_chg = tmp.groupby("_sku", sort=False)["_chg"].sum()
    per_buyer_skus: dict[str, int] = {}
    per_buyer_chg: dict[str, float] = {}
    for sku, b in per_sku_buyer.items():
        b = str(b)
        per_buyer_skus[b] = per_buyer_skus.get(b, 0) + 1
        per_buyer_chg[b] = per_buyer_chg.get(b, 0.0) + float(per_sku_chg.loc[sku])
    result = [
        (b, int(per_buyer_skus[b]), int(round(per_buyer_chg[b])))
        for b in per_buyer_skus
    ]
    return sorted(result, key=lambda kv: (-kv[1], kv[0]))


def _hybrid_buyer_combined_stats(
    buyer_need: list[tuple[str, int]] | None,
    buyer_safety: list[tuple[str, int, int]] | None,
) -> list[tuple[str, int, int, int]]:
    """合併 Buyer 統計：將需調貨與需調整 Safety 按 Buyer 外連接為單一列。

    輸入為 `_hybrid_buyer_need_stats` 與 `_hybrid_buyer_safety_stats` 的回傳；
    任一為空或缺席的 Buyer 以 0 補齊。回傳
    [(Buyer, 需調貨 SKU 數, 需調整 Safety SKU 數, Safety 建議調整量合計)]，
    按需調貨數降序、再按需調整數降序、同數按 Buyer 名排序。
    （保留供相容；新合併表含貨量與佔比，請改用 `_hybrid_buyer_combined_detail`。）
    """
    need_map: dict[str, int] = {str(b): int(c) for b, c in (buyer_need or [])}
    safety_map: dict[str, tuple[int, int]] = {
        str(b): (int(c), int(v)) for b, c, v in (buyer_safety or [])
    }
    buyers = sorted(
        set(need_map) | set(safety_map),
        key=lambda b: (-need_map.get(b, 0), -safety_map.get(b, (0, 0))[0], b),
    )
    return [
        (b, int(need_map.get(b, 0)), int(safety_map.get(b, (0, 0))[0]), int(safety_map.get(b, (0, 0))[1]))
        for b in buyers
    ]


def _hybrid_buyer_combined_detail(raw: pd.DataFrame) -> list[tuple[str, int, int, int, int]]:
    """Buyer 合併明細：同一 Buyer 的需調貨與需調整 Safety 並列（已刪除重複）。

    Buyer 取同 SKU 全表首個非空值（需調貨與需調整共用同一歸屬，避免同一 SKU
    在兩母體歸入不同 Buyer）；全空歸入「(未填 Buyer)」。
    需調貨貨量為該 Buyer 名下各需調貨 SKU 的「調入需求量 (補到14日)」列合計
    （取整，與指標卡總需求同口徑）；Safety 調整量為「Safety 建議調整量」列合計
    （取整，正=增加、負=減少）。
    回傳 [(Buyer, 需調貨 SKU 數, 調入需求量合計, 需調整 Safety SKU 數,
    Safety 建議調整量合計)]，按需調貨數降序、再按需調整數降序、同數按 Buyer 名排序。
    """
    if raw is None or raw.empty or "SKU" not in raw.columns:
        return []
    buyer = (
        raw["Buyer"].fillna("").astype(str).str.strip()
        if "Buyer" in raw.columns
        else _safe_default_series("", raw.index)
    )
    sku = raw["SKU"]

    def _first_nonempty(s: pd.Series) -> str:
        s = s[s != ""]
        return s.iloc[0] if len(s) else "(未填 Buyer)"

    per_sku_buyer = pd.DataFrame({"_sku": sku, "_buyer": buyer}).groupby("_sku", sort=False)["_buyer"].agg(_first_nonempty)

    need_mask = raw.get("需要調貨", _safe_default_series("", raw.index)).astype(str) == "是"
    demand = pd.to_numeric(
        raw.get("調入需求量 (補到14日)", _safe_default_series(np.nan, raw.index)), errors="coerce"
    ).fillna(0.0)
    need_skus = set(raw.loc[need_mask, "SKU"].dropna().unique()) if bool(need_mask.any()) else set()
    per_sku_demand = demand.loc[need_mask].groupby(raw.loc[need_mask, "SKU"], sort=False).sum() if need_skus else pd.Series(dtype=float)

    diag = raw.get("Safety Review", _safe_default_series("", raw.index)).fillna("").astype(str)
    fix_mask = diag.isin([_HYBRID_DIAG_UNDER, _HYBRID_DIAG_OVER])
    chg = pd.to_numeric(
        raw.get("Safety 建議調整量", _safe_default_series(np.nan, raw.index)), errors="coerce"
    ).fillna(0.0)
    fix_skus = set(raw.loc[fix_mask, "SKU"].dropna().unique()) if bool(fix_mask.any()) else set()
    per_sku_chg = chg.loc[fix_mask].groupby(raw.loc[fix_mask, "SKU"], sort=False).sum() if fix_skus else pd.Series(dtype=float)

    if not need_skus and not fix_skus:
        return []
    per_buyer: dict[str, list[float]] = {}
    for s in need_skus | fix_skus:
        b = str(per_sku_buyer.get(s, "(未填 Buyer)"))
        rec = per_buyer.setdefault(b, [0, 0.0, 0, 0.0])
        if s in need_skus:
            rec[0] += 1
            rec[1] += float(per_sku_demand.get(s, 0.0))
        if s in fix_skus:
            rec[2] += 1
            rec[3] += float(per_sku_chg.get(s, 0.0))
    result = [
        (b, int(v[0]), int(round(v[1])), int(v[2]), int(round(v[3])))
        for b, v in per_buyer.items()
    ]
    return sorted(result, key=lambda kv: (-kv[1], -kv[3], kv[0]))


def _hybrid_safety_days_from_raw(raw: pd.DataFrame, nature=None) -> dict[str, int]:
    """由 RAW 的「Safety 基準天數」欄還原各 Nature 實際採用的基準天數
    （該欄每列同一 Nature 恆定）；缺欄時退回預設 7 / 9 / 10。"""
    days, _ = _hybrid_safety_days({})
    if raw is None or raw.empty or _HYBRID_SAFETY_DAYS_COLUMN not in raw.columns:
        return days
    if nature is None:
        nature = raw.get("Shop Nature", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    vals = pd.to_numeric(raw[_HYBRID_SAFETY_DAYS_COLUMN], errors="coerce")
    for n in _HYBRID_DIAG_NATURES:
        sub = vals[nature == n].dropna()
        if len(sub):
            days[n] = int(sub.iloc[0])
    return days


def _hybrid_stats_from_raw(raw: pd.DataFrame) -> dict:
    """由 raw 匯總 Review stats（stats 缺失時的回退，與 transform_hybrid 同口徑）。"""
    if raw is None or raw.empty:
        return {}
    need = raw.get("需要調貨", _safe_default_series("", raw.index)).astype(str) == "是"
    evt = raw.get("事件類型", _safe_default_series("", raw.index)).fillna("").astype(str)
    diag = raw.get("Safety Review", _safe_default_series("", raw.index)).fillna("").astype(str)
    chg = pd.to_numeric(raw.get("Safety 建議調整量", _safe_default_series(np.nan, raw.index)), errors="coerce")
    nature = raw.get("Shop Nature", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    bidirectional = "可轉出貨量 (Tourist)" in raw.columns
    under = diag == _HYBRID_DIAG_UNDER
    over = diag == _HYBRID_DIAG_OVER
    donor = evt.isin([_HYBRID_EVENT_DONATED, _HYBRID_EVENT_DONOR])
    return {
        "safety_days": _hybrid_safety_days_from_raw(raw, nature),
        "dimension": "nature",
        "need_rows": int(need.sum()),
        "need_skus": int(raw.loc[need, "SKU"].nunique()) if need.any() else 0,
        "total_demand": int(_safe_numeric(raw.get("調入需求量 (補到14日)", _safe_default_series(0, raw.index))).sum()),
        "total_actual": int(_safe_numeric(raw.get("實際調入量", _safe_default_series(0, raw.index))).sum()),
        "total_gap": int(_safe_numeric(raw.get("調入缺口", _safe_default_series(0, raw.index))).sum()),
        "nt_transfer": int(_safe_numeric(raw.loc[nature == "Non Tourist", "轉出量"]).sum()) if "轉出量" in raw.columns else 0,
        "mix_transfer": int(_safe_numeric(raw.loc[nature == "Mix", "轉出量"]).sum()) if "轉出量" in raw.columns else 0,
        "tourist_transfer": int(_safe_numeric(raw.loc[nature == "Tourist", "轉出量"]).sum()) if bidirectional and "轉出量" in raw.columns else 0,
        "total_actual_om": int(_safe_numeric(raw.get("實際調入量 (同OM)", _safe_default_series(0, raw.index))).sum()),
        "total_gap_om": int(_safe_numeric(raw.get("調入缺口 (同OM)", _safe_default_series(0, raw.index))).sum()),
        "nt_transfer_om": int(_safe_numeric(raw.loc[nature == "Non Tourist", "轉出量 (同OM)"]).sum()) if "轉出量 (同OM)" in raw.columns else 0,
        "mix_transfer_om": int(_safe_numeric(raw.loc[nature == "Mix", "轉出量 (同OM)"]).sum()) if "轉出量 (同OM)" in raw.columns else 0,
        "tourist_transfer_om": int(_safe_numeric(raw.loc[nature == "Tourist", "轉出量 (同OM)"]).sum()) if bidirectional and "轉出量 (同OM)" in raw.columns else 0,
        "bidirectional": bool(bidirectional),
        "donor_rows": int(donor.sum()),
        "under_rows": int(under.sum()),
        "under_skus": int(raw.loc[under, "SKU"].nunique()) if under.any() else 0,
        "under_inc_total": int(round(float(chg[under].fillna(0).sum()))),
        "over_rows": int(over.sum()),
        "over_skus": int(raw.loc[over, "SKU"].nunique()) if over.any() else 0,
        "over_dec_total": int(round(float(chg[over].fillna(0).sum()))),
        "normal_rows": int((evt != "").sum()) - int(under.sum()) - int(over.sum()),
        "total_raw_rows": int(len(raw)),
        "buyer_need_skus": _hybrid_buyer_need_stats(raw, need),
        "buyer_safety_fix": _hybrid_buyer_safety_stats(raw),
    }


def _hybrid_dashboard_metrics(raw: pd.DataFrame, stats: dict | None = None):
    """調貨 Review 指標卡（每排 4 卡，與 Streamlit Hybrid Dashboard 同口徑；
    第 4 排為同OM調貨對照，Safety Review 建議不受 OM 限制影響；
    雙向調貨時另加第 5 排 Tourist 轉出量兩卡）。
    回傳各排指標 list（4 或 5 排）供 _write_metric_cards 使用；stats 缺失時由 raw 自動匯總。"""
    if stats is None:
        stats = _hybrid_stats_from_raw(raw)

    def g(key: str) -> int:
        return int(stats.get(key, 0) or 0)

    bidir = bool(stats.get("bidirectional"))
    row1 = [
        ("需調貨 SKU 數", g("need_skus"), "符合調貨條件的 SKU 總數（已刪除重複）", "f97316"),
        ("需要調入列數", g("need_rows"),
         "符合調貨條件的店舖×SKU 列數（任何 Nature 皆可調入）" if bidir else "符合調貨條件的 Tourist 店舖×SKU 列數", "f59e0b"),
        ("轉出來源列數", g("donor_rows"),
         "實際轉出與可轉出未使用的各 Nature 列數" if bidir else "實際轉出與可轉出未使用的 Mix / Non Tourist 列數", "8b5cf6"),
        ("調入需求量", g("total_demand"), "補到 14 日周轉的總需求（件）", "3b82f6"),
    ]
    row2 = [
        ("實際調入量", g("total_actual"),
         "不限OM：NT→Mix→Tourist 依序補足後的總調入（件）" if bidir else "不限OM：NT 優先、Mix 補足後的總調入（件）", "10b981"),
        ("無法足額缺口", g("total_gap"), "不限OM：來源可轉出量不足的缺口（件）", "ef4444"),
        ("Non Tourist 轉出量", g("nt_transfer"), "不限OM：優先由 Non Tourist 店舖轉出（件）", "8b5cf6"),
        ("Mix 轉出量", g("mix_transfer"), "不限OM：NT 不足時由 Mix 店舖補足（件）", "06b6d4"),
    ]
    row3 = [
        ("Safety 設定不足列數", g("under_rows"), "現有設定低於「Safety 周轉所需量」的調入列數（基準天數按 Nature）", "ef4444"),
        ("不足建議增加合計", g("under_inc_total"), "設定不足列的 Safety 建議調整量合計（正=增加）", "10b981"),
        ("Safety 設定過高列數", g("over_rows"), "現有設定高於「Safety 周轉所需量」的轉出列數（基準天數按 Nature）", "f59e0b"),
        ("過高建議減少合計", g("over_dec_total"), "設定過高列的 Safety 建議調整量合計（負=減少）", "3b82f6"),
    ]
    row4 = [
        ("實際調入量 (同OM)", g("total_actual_om"), "限同 OM 調貨後的總調入（件）", "10b981"),
        ("無法足額缺口 (同OM)", g("total_gap_om"), "限同 OM 調貨後仍不足的缺口（件）", "ef4444"),
        ("Non Tourist 轉出量 (同OM)", g("nt_transfer_om"), "限同 OM：Non Tourist 店舖轉出（件）", "8b5cf6"),
        ("Mix 轉出量 (同OM)", g("mix_transfer_om"), "限同 OM：Mix 店舖轉出（件）", "06b6d4"),
    ]
    rows = [row1, row2, row3, row4]
    if bidir:
        rows.append([
            ("Tourist 轉出量", g("tourist_transfer"), "不限OM：由 Tourist 店舖轉出（件）", "f43f5e"),
            ("Tourist 轉出量 (同OM)", g("tourist_transfer_om"), "限同 OM：Tourist 店舖轉出（件）", "ec4899"),
        ])
    return rows


def _hybrid_allocate_transfers(
    n_rows: int,
    sku_arr: np.ndarray,
    site_arr: np.ndarray,
    need_mask: np.ndarray,
    demand: np.ndarray,
    nt_avail0: np.ndarray,
    mix_avail0: np.ndarray,
    om_arr: np.ndarray | None = None,
    tourist_avail0: np.ndarray | None = None,
    nature_arr: np.ndarray | None = None,
):
    """調貨分配：按 SKU 彙總需求，來源池固定 Non Tourist → Mix → Tourist（tourist_avail0
    為 None 時不設 Tourist 池，即現行單向規則），需求列按需求量大→小依序分配。
    om_arr 非 None 時來源店須與需求列同 OM 才可轉出（同OM調貨）。
    nature_arr 非 None 時來源列與需求列同 Shop Nature 一律排除（「需求 Nature 自身除外」
    契約守衛；現行池級歸零已令同 Nature 來源恆為 0，此守衛為顯式防呆）。
    來源店與需求列同為一店（Site Code 相同）時一律排除，防止重複列資料異常造成自我調貨。
    回傳 (actual_np, gap_np, donate_np, desc_np)。"""
    actual_np = np.zeros(n_rows, dtype=int)
    donate_np = np.zeros(n_rows, dtype=int)
    gap_np = np.zeros(n_rows, dtype=int)
    desc_np = np.full(n_rows, "", dtype=object)
    if not need_mask.any():
        return actual_np, gap_np, donate_np, desc_np
    nt_avail = np.asarray(nt_avail0, dtype=float).copy()
    mix_avail = np.asarray(mix_avail0, dtype=float).copy()
    tourist_avail = (
        np.asarray(tourist_avail0, dtype=float).copy() if tourist_avail0 is not None else None
    )
    for sku in set(sku_arr[need_mask]):
        sku_mask = sku_arr == sku
        need_rows = np.nonzero(sku_mask & need_mask)[0]
        nt_src = [int(p) for p in np.nonzero(sku_mask & (nt_avail > 0))[0]]
        mix_src = [int(p) for p in np.nonzero(sku_mask & (mix_avail > 0))[0]]
        order = np.argsort(-demand[need_rows], kind="stable")
        remain_nt = {p: float(nt_avail[p]) for p in nt_src}
        remain_mix = {p: float(mix_avail[p]) for p in mix_src}
        pools = [remain_nt, remain_mix]
        if tourist_avail is not None:
            tourist_src = [int(p) for p in np.nonzero(sku_mask & (tourist_avail > 0))[0]]
            pools.append({p: float(tourist_avail[p]) for p in tourist_src})
        for n in order:
            row_pos = int(need_rows[n])
            rem = float(demand[row_pos])
            parts: list[str] = []
            donated_actual = 0
            for pool in pools:
                for spos in list(pool.keys()):
                    if rem <= 0:
                        break
                    if site_arr[spos] == site_arr[row_pos]:
                        continue  # 自我調貨守衛：來源店＝需求店時排除（重複列資料異常防護）
                    if nature_arr is not None and nature_arr[spos] == nature_arr[row_pos]:
                        continue  # 需求 Nature 自身除外守衛（雙向契約，顯式防呆）
                    if om_arr is not None and om_arr[spos] != om_arr[row_pos]:
                        continue
                    avail_n = pool[spos]
                    take = min(rem, avail_n)
                    if take > 0:
                        pool[spos] -= take
                        rem -= take
                        donate_np[spos] += int(take)
                        donated_actual += int(take)
                        parts.append(f"{site_arr[spos]}:{int(take)}")
                if rem <= 0:
                    break
            actual_np[row_pos] = donated_actual
            gap_np[row_pos] = int(rem)
            desc_np[row_pos] = "; ".join(parts)
    return actual_np, gap_np, donate_np, desc_np


def transform_hybrid(
    df_input: pd.DataFrame,
    turnday_threshold: int = 14,
    d001_threshold: int = 100,
    bidirectional: bool = False,
    safety_days: dict | None = None,
    dimension: str = "nature",
    progress_callback=None,
):
    """Safety Stock x NetStock Review 核心轉換（調貨事件 Review Safety 設定）。

    - 分類維度（dimension，預設 "nature" 維持現行契約）：
      "nature"＝以 Shop Nature 分類；"rp"＝以 RP Type 分類；"both"＝Shop Nature＋RP Type（6 組）。
      維度決定需求方過濾與 Summary／交叉表分組，**不改變 TurnDay 池與 Safety 基準天數
      （兩者恆按 Shop Nature）**。RP Type 正規化：RF＝可補貨、ND＝不可補貨；
      其他／空白歸 ND 並列 warning。來源缺「RP Type」欄時全部列按 RF 處理
      （ND 排除規則無法套用）、維度退回 nature 並列 warning。
      需求方一律排除 ND 列（ND 不作調入方，仍可作供應方）——三種維度皆適用：
      nature／both 需求方為 Tourist-RF；rp 需求方為任何 Nature 的 RF 列，
      且取消「需求 Nature 自身除外」守衛與池級歸零（允許同 Nature 店間再平衡）。
      供應方三種維度皆維持 Mix／Non Tourist（勾選雙向時另加 Tourist）。
      非法維度值退回 nature，不拋錯。
    - 調貨段：與 Stock Checking 同規則（SKU 級加總 TurnDay、Nature 供應門檻、
      14 日需求/可轉出、NT 優先→Mix 補足、需求量大→小），另疊加本模式特有的
      **資料品質補正**（Shop Nature 正規化、月銷逐列回補、D001 逐列 Net→Stock
      fallback；回補口徑見各段 warning），故同一份輸入在兩個模式可能出現細微差異。
      **調貨段恆以 14 日為基準，不受 safety_days 影響。**
    - Review 段（Safety 基準周轉天數，按 Shop Nature 分別設定）：
      Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 該列 Nature 的基準天數)；
      safety_days 預設 {"Tourist": 7, "Mix": 9, "Non Tourist": 10}（UI 建議範圍 7–10 日）。
      Safety 設定不跟隨存貨周轉（調貨）天數，兩段基準各自獨立。
      Tourist 需調入列：現有 Adj. Safety Stock < 所需量 → Safety 設定不足；
      已標示 SKU 的有貨列（實際轉出或可轉出未使用）：
      現有設定 > 所需量 → Safety 設定過高；等於邊界一律視為設定正常。
      建議 Safety 設定值 = Safety 周轉所需量（僅不足/過高列），
      Safety 建議調整量 = 建議值 − 現有設定（正=增加、負=減少）。
    - bidirectional=True（允許雙向調貨，預設 False 維持單向契約）：任何 Nature 缺貨
      （該 Nature 加總 TurnDay < 門檻 且 D001 < 門檻 且列需求 > 0）且其他任一 Nature
      有貨（加總 TurnDay ≥ 門檻）即標示需要調貨；新增 RAW 欄「可轉出貨量 (Tourist)」，
      Tourist 列亦可作轉出來源；來源池順序固定 Non Tourist → Mix → Tourist，
      且「需求 Nature 自身除外」為分配層的顯式守衛（池級歸零另保證缺貨 Nature 不進入來源池）。
      Review 段對稱擴展：需要調入可出現在任何 Nature 列、實際轉出／可轉出未使用
      可出現在 Tourist 列（事件與診斷文字不變）。
    - 上下文欄位：調貨後庫存 = 原始 + 實際調入 − 轉出；
      調貨後 TurnDay = 調貨後庫存 ÷ 月銷 × 30.5（月銷 0 留空；月銷為回補後值，
      另存「Effective Shop Monthly Sold Qty (P90D)」稽核欄，原值保留於原欄）。
    - OM 變體：分配跑兩遍——不限OM（契約欄位）與同OM（來源店須與需求店同 OM，
      欄位加「(同OM)」後綴）；Safety Review 與建議調整量兩變體相同，不受調貨限制影響。
    - Summary：附 Stock Checking 的 SKU 級彙總（按分類維度分組：nature 三組／rp 兩組／
      both 六組；庫存/月銷加總、TurnDay、佔比、強制調貨提示、存貨偏重提示），
      供查閱哪些 SKU 存貨偏重哪個分組。
    - Buyer 統計：沿用 Stock Checking 的 Buyer 需調貨 SKU 統計（同 SKU 首個非空 Buyer，
      已刪除重複），另加 Buyer 需調整 Safety SKU 統計（Safety 設定不足＋過高母體，
      含各 Buyer 的調整量合計）。
    回傳 (raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, warnings)；
    stats["safety_days"] 為實際採用的各 Nature 基準天數；
    stats["dimension"] 為實際採用的分類維度（rp 欄缺失退回 nature 時亦為 nature）；
    stats["buyer_need_skus"] 為 [(Buyer, 需調貨 SKU 數)]；
    stats["buyer_safety_fix"] 為 [(Buyer, 需調整 Safety SKU 數, Safety 建議調整量合計)]。
    """
    missing = validate_hybrid_input(df_input)
    if missing:
        raise ValueError(f"缺少混合模式必要欄位: {', '.join(missing)}")
    df_input = _canonicalize_hybrid_columns(df_input)
    if progress_callback:
        progress_callback(0.05, "驗證通過")
    warnings: list[str] = []

    df = _clean_dataframe(df_input, shop_natures=None, progress_callback=progress_callback)
    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據（非空白或總計列）")

    df["SKU"] = df["SKU"].apply(_safe_sku)
    # SKU Desc 缺失或全欄空白時亦補正（既有空白欄位會讓摘要 groupby 失效）
    if "SKU Desc" not in df.columns or df["SKU Desc"].fillna("").astype(str).str.strip().eq("").all():
        df = _compute_raw_extra_columns(df)

    if progress_callback:
        progress_callback(0.12, "正規化 Shop Nature...")
    df, mapped_other = _normalize_hybrid_nature(df)
    if mapped_other:
        warnings.append(f"有 {mapped_other} 列 Shop Nature 無法辨識（Other / 空白 / 其他非標準值），已正規化為 Non Tourist。")
    warnings.extend(_transfer_input_consistency_warnings(df))

    # ── 分類維度與 RP Type 正規化（RF＝可補貨、ND＝不可補貨；ND 列不作調貨需求方）──
    dim = _hybrid_dimension(dimension)
    df, mapped_rp, rp_present = _normalize_hybrid_rp_type(df)
    if not rp_present:
        warnings.append(
            "來源檔案缺少「RP Type」欄位，ND 列排除規則無法套用（全部列按可補貨 RF 處理）"
            + ("，分類維度已退回 Shop Nature。" if dim != "nature" else "。")
        )
        dim = "nature"
    elif mapped_rp:
        warnings.append(
            f"有 {mapped_rp} 列 RP Type 無法辨識（非 RF／ND 或空白），已正規化為 ND"
            "（不可補貨：不作調貨需求方，仍可作供應方）。"
        )
    is_rf = (
        (df["RP Type"] == "RF").to_numpy()
        if rp_present
        else np.ones(len(df), dtype=bool)
    )

    idx = df.index
    nature = df["Shop Nature"]

    sold_raw = _hybrid_numeric_column(df, "Sold Qty (P90D)", idx, allow_negative=True)
    n_sold_invalid = int(sold_raw.isna().sum())
    n_sold_negative = int((sold_raw < 0).fillna(False).sum())
    if n_sold_invalid:
        warnings.append(f"有 {n_sold_invalid} 列「Sold Qty (P90D)」空白或非數值，已視為 0。")
    if n_sold_negative:
        warnings.append(f"有 {n_sold_negative} 列「Sold Qty (P90D)」為負數，已按 0 計算。")
    sold_p90d = sold_raw.mask(sold_raw < 0, 0.0).fillna(0.0)

    # 庫存量必須為整數（件數）；空白視為 0，小數、負數、非有限值直接拒絕避免守恆失真
    cm_stock_raw = _hybrid_numeric_column(df, "Shop Stock Qty (CM)", idx)
    nonint = cm_stock_raw.notna() & (cm_stock_raw != np.floor(cm_stock_raw))
    if nonint.any():
        raise ValueError("「Shop Stock Qty (CM)」必須為整數件數，發現小數值，請修正來源資料後再執行。")
    cm_stock = cm_stock_raw.fillna(0.0)
    df["Shop Stock Qty (CM)"] = cm_stock

    # 月銷空白、非數值、負數，或月銷為 0 但 P90D 有正數銷售時，逐列以 P90D 換算
    msold_raw = _hybrid_numeric_column(df, "Shop Monthly Sold Qty (P90D)", idx, allow_negative=True)
    n_ms_negative = int((msold_raw < 0).fillna(False).sum())
    if n_ms_negative:
        warnings.append(f"有 {n_ms_negative} 列「Shop Monthly Sold Qty (P90D)」為負數，已改用 P90D 換算。")
    msold_raw = msold_raw.mask(msold_raw < 0)
    msold_blank = msold_raw.isna()
    msold_zero_conflict = msold_raw.eq(0) & sold_p90d.gt(0)
    n_ms_fallback = int((msold_blank | msold_zero_conflict).sum())
    if n_ms_fallback:
        warnings.append(
            f"有 {n_ms_fallback} 列「Shop Monthly Sold Qty (P90D)」空白或為 0 但 P90D 有銷售，"
            "已逐列以 Sold Qty (P90D) ÷ 90 × 30.5 換算月銷。"
        )
    msold_eff = msold_raw.where(~(msold_blank | msold_zero_conflict), sold_p90d / 90.0 * 30.5)
    # 實際採用月銷另存 Effective 欄（與 D001 Effective 欄同口徑）；覆寫前保留原始輸入供稽核比對
    df["Effective Shop Monthly Sold Qty (P90D)"] = msold_eff
    df["Shop Monthly Sold Qty (P90D)"] = msold_eff

    # D001：逐列 Net 優先、Stock 回補（兩者皆空白或非有限值視為空白），並保留 Effective 欄位供 RAW 稽核
    avail_now = set(df.columns)
    net_d001_col = _resolve_column_name("D001 Net Stock Qty (CM)", avail_now)
    stock_d001_col = _resolve_column_name("D001 Stock Qty (CM)", avail_now)
    if net_d001_col is None:
        if stock_d001_col is not None:
            warnings.append("來源檔案缺少「D001 Net Stock Qty (CM)」，D001 以「D001 Stock Qty (CM)」取代。")
        else:
            warnings.append("來源檔案缺少「D001 Net Stock Qty (CM)」與「D001 Stock Qty (CM)」，D001 一律視為 0（D001 低於門檻的判定將全數成立）。")
    net_d001_s = (
        _hybrid_numeric_column(df, net_d001_col, idx, allow_negative=True, nonfinite_as_missing=True)
        if net_d001_col else _safe_default_series(np.nan, idx)
    )
    stock_d001_s = (
        _hybrid_numeric_column(df, stock_d001_col, idx, allow_negative=True, nonfinite_as_missing=True)
        if stock_d001_col else _safe_default_series(np.nan, idx)
    )
    n_d001_blank = int(net_d001_s.isna().sum()) if net_d001_col else 0
    if n_d001_blank:
        warnings.append(f"有 {n_d001_blank} 列「D001 Net Stock Qty (CM)」空白或非有限值，已逐列改用「D001 Stock Qty (CM)」回補（兩者皆空白視為 0）。")
    eff_d001 = net_d001_s.where(net_d001_s.notna(), stock_d001_s).fillna(0.0)
    df["D001 Effective Stock Qty (CM)"] = eff_d001.astype(int)

    is_mix = (nature == "Mix").to_numpy()
    is_nt = (nature == "Non Tourist").to_numpy()
    is_tourist = (nature == "Tourist").to_numpy()
    cm_stock_val = cm_stock.to_numpy(dtype=float)
    sold_val = sold_p90d.to_numpy(dtype=float)
    msold_val = msold_eff.to_numpy(dtype=float)
    no_sale_has_stock = (sold_val <= 0) & (cm_stock_val > 0)
    ideal_14d = sold_val / 90.0 * 14
    # Safety 基準天數（按 Nature）— 只用於 Review 段；調貨段一律用上面的 14 日
    _safety_days, _safety_day_warnings = _hybrid_safety_days(safety_days)
    warnings.extend(_safety_day_warnings)
    safety_days_np = np.select(
        [is_tourist, is_mix, is_nt],
        [float(_safety_days["Tourist"]), float(_safety_days["Mix"]), float(_safety_days["Non Tourist"])],
        default=float(_safety_days["Non Tourist"]),
    )

    sku_key = df["SKU"]

    def _sku_nature_turnday(nature_mask):
        s_stock = pd.Series(np.where(nature_mask, cm_stock_val, 0.0), index=idx)
        s_msold = pd.Series(np.where(nature_mask, msold_val, 0.0), index=idx)
        agg_stock = s_stock.groupby(sku_key).transform("sum")
        agg_msold = s_msold.groupby(sku_key).transform("sum")
        with np.errstate(divide="ignore", invalid="ignore"):
            return np.where(agg_msold.to_numpy() > 0, agg_stock.to_numpy() / agg_msold.to_numpy() * 30.5, np.inf)

    sku_tourist_turnday = _sku_nature_turnday(is_tourist)
    sku_mix_turnday = _sku_nature_turnday(is_mix)
    sku_nt_turnday = _sku_nature_turnday(is_nt)
    for label, td in (("Tourist", sku_tourist_turnday), ("Mix", sku_mix_turnday), ("Non Tourist", sku_nt_turnday)):
        df[f"{label} 加總 TurnDay"] = np.round(np.where(np.isfinite(td), td, np.nan), 2)

    per_row_excess = np.where(no_sale_has_stock, cm_stock_val, np.maximum(0, cm_stock_val - ideal_14d))
    mix_pool_ok = sku_mix_turnday >= turnday_threshold
    nt_pool_ok = sku_nt_turnday >= turnday_threshold
    if dim == "rp":
        # RP Type 模式：允許同 Nature 店間再平衡——不作池級歸零，可轉出量按逐列 14 日超額計
        mix_excess = np.where(is_mix, per_row_excess, 0)
        nt_excess = np.where(is_nt, per_row_excess, 0)
    else:
        mix_excess = np.where(mix_pool_ok, np.where(is_mix, per_row_excess, 0), 0)
        nt_excess = np.where(nt_pool_ok, np.where(is_nt, per_row_excess, 0), 0)
    df["可轉出貨量 (Mix)"] = mix_excess.astype(int)
    df["可轉出貨量 (Non Tourist)"] = nt_excess.astype(int)

    demand = np.maximum(0, np.ceil(ideal_14d) - cm_stock_val)
    d001_np = eff_d001.to_numpy(dtype=float)
    ttd_np = pd.Series(sku_tourist_turnday, index=idx).to_numpy(dtype=float)
    mtd_np = pd.Series(sku_mix_turnday, index=idx).to_numpy(dtype=float)
    ntd_np = pd.Series(sku_nt_turnday, index=idx).to_numpy(dtype=float)
    # 每列所屬 Nature 的 SKU 級加總 TurnDay（rp 模式需求判定用；仍按 Shop Nature 口徑）
    nature_td_np = np.select([is_tourist, is_mix], [ttd_np, mtd_np], default=ntd_np)
    tourist_pool_ok = sku_tourist_turnday >= turnday_threshold
    if dim == "rp":
        # 需求方＝任何 Nature 的 RF 列（ND 一律不作需求）；有貨可調＝同 SKU 任一列有超額
        # （含同 Nature 店間再平衡，故不以他 Nature 池級門檻為前提）；
        # 供應池維持 Mix／Non Tourist（雙向時另加 Tourist）
        if bidirectional:
            tourist_excess = np.where(is_tourist, per_row_excess, 0)
            df["可轉出貨量 (Tourist)"] = tourist_excess.astype(int)
        else:
            tourist_excess = np.zeros(len(df))
        supply_avail = mix_excess + nt_excess + tourist_excess
        sku_supply = pd.Series(supply_avail, index=idx).groupby(sku_key).transform("sum").to_numpy(dtype=float)
        need_mask = (
            is_rf
            & (nature_td_np < turnday_threshold)
            & (d001_np < d001_threshold)
            & (demand > 0)
            & (sku_supply > 0)
        )
    elif bidirectional:
        # 對稱三向：任何 Nature 缺貨（加總 TurnDay < 門檻）且其他任一 Nature 有貨即為需求方
        # （ND 列一律不作需求）；Tourist 亦可作轉出來源（來源池順序固定 NT → Mix → Tourist）
        tourist_excess = np.where(
            tourist_pool_ok,
            np.where(is_tourist, per_row_excess, 0),
            0,
        )
        df["可轉出貨量 (Tourist)"] = tourist_excess.astype(int)
        need_mask = (
            (
                (is_tourist & (ttd_np < turnday_threshold) & (mix_pool_ok | nt_pool_ok))
                | (is_mix & (mtd_np < turnday_threshold) & (nt_pool_ok | tourist_pool_ok))
                | (is_nt & (ntd_np < turnday_threshold) & (mix_pool_ok | tourist_pool_ok))
            )
            & is_rf
            & (d001_np < d001_threshold)
            & (demand > 0)
        )
    else:
        tourist_excess = np.zeros(len(df))
        # 需求方＝Tourist-RF（ND 列一律不作需求）；both 維度計算與本分支完全相同，僅分組顯示不同
        need_mask = (
            is_tourist
            & is_rf
            & (ttd_np < turnday_threshold)
            & (d001_np < d001_threshold)
            & (demand > 0)
            & (mix_pool_ok | nt_pool_ok)
        )
    need_series = pd.Series(need_mask, index=idx)

    df["需要調貨"] = np.where(need_mask, "是", "")
    df["調入需求量 (補到14日)"] = np.where(need_mask, demand, 0).astype(int)

    site_arr = df.get("Site Code", _safe_default_series("", idx)).fillna("").astype(str).str.strip().to_numpy(dtype=object)
    sku_arr = df["SKU"].to_numpy(dtype=object)
    n_rows = len(df)

    # ── OM 群組（同OM調貨變體用）：欄缺失→全部同群組並 warning；空白→(未填OM) ──
    om_col = _resolve_column_name("OM", set(df.columns))
    if om_col is None:
        warnings.append("來源檔案缺少「OM」欄位，「同OM調貨」無法區分店群，全部店舖視為同一組（結果與不限OM相同）。")
        df["OM"] = "(未填OM)"
        om_arr = np.full(n_rows, "(未填OM)", dtype=object)
    else:
        om_raw = df[om_col].fillna("").astype(str).str.strip()
        n_om_blank = int((om_raw == "").sum())
        if n_om_blank:
            warnings.append(f"有 {n_om_blank} 列「OM」空白，同OM調貨時歸入「(未填OM)」群組。")
        om_arr = om_raw.replace("", "(未填OM)").to_numpy(dtype=object)
        df["OM"] = om_arr

    # ── 調貨分配：不限OM（現行契約欄位）＋ 同OM（來源店與需求店同 OM 才可轉出）──
    # rp 維度取消「需求 Nature 自身除外」守衛（允許同 Nature 店間再平衡）；
    # nature／both 維度維持守衛（池級歸零＋分配層守衛雙重保證）
    nt_avail0 = df["可轉出貨量 (Non Tourist)"].to_numpy(dtype=float)
    mix_avail0 = df["可轉出貨量 (Mix)"].to_numpy(dtype=float)
    tourist_avail0 = df["可轉出貨量 (Tourist)"].to_numpy(dtype=float) if bidirectional else None
    _guard_nature = None if dim == "rp" else nature.to_numpy()
    actual_np, gap_np, donate_np, desc_np = _hybrid_allocate_transfers(
        n_rows, sku_arr, site_arr, need_mask, demand, nt_avail0, mix_avail0,
        tourist_avail0=tourist_avail0, nature_arr=_guard_nature,
    )
    actual_om, gap_om, donate_om, desc_om = _hybrid_allocate_transfers(
        n_rows, sku_arr, site_arr, need_mask, demand, nt_avail0, mix_avail0, om_arr=om_arr,
        tourist_avail0=tourist_avail0, nature_arr=_guard_nature,
    )
    df["實際調入量"] = actual_np
    df["調入缺口"] = gap_np
    df["轉出量"] = donate_np
    df["調貨來源 (Site Code)"] = desc_np
    df["實際調入量 (同OM)"] = actual_om
    df["調入缺口 (同OM)"] = gap_om
    df["轉出量 (同OM)"] = donate_om
    df["調貨來源 (同OM)"] = desc_om

    if not need_mask.any():
        if dim == "rp":
            warnings.append("沒有符合調貨條件的 RF 列（任何 Nature 加總 TurnDay < 門檻 且 D001 < 門檻 且補到 14 日周轉的需求 > 0 且有貨可調；ND 列不作需求方）。")
        elif bidirectional:
            warnings.append("沒有符合調貨條件的 SKU（任何 Nature 的 RF 列加總 TurnDay < 門檻 且 D001 < 門檻 且補到 14 日周轉的需求 > 0 且有貨可調）。")
        else:
            warnings.append("沒有符合調貨條件的 Tourist RF SKU（加總 TurnDay < 門檻 且 D001 < 門檻 且補到 14 日周轉的需求 > 0 且有貨可調；ND 列不作需求方）。")

    if progress_callback:
        progress_callback(0.45, "Review Safety 設定...")
    raw = df

    # ── 調貨事件 Review（Safety 基準天數按 Nature；Review 與建議調整不受 OM 限制影響）──
    safety_line = np.maximum(0, np.ceil(np.maximum(0, sold_val / 90.0 * safety_days_np)))
    adj_s = _hybrid_numeric_column(raw, "Adj. Safety Stock", idx)
    n_adj_blank = int(adj_s.isna().sum())
    if n_adj_blank:
        warnings.append(f"有 {n_adj_blank} 列「Adj. Safety Stock」空白或非數值，Review 時已視為 0。")
    adj_val = adj_s.fillna(0.0).to_numpy(dtype=float)

    flagged_skus = set(sku_arr[need_mask])
    transferable_np = mix_excess + nt_excess + (tourist_excess if bidirectional else 0)
    donor_pool = np.isin(sku_arr, list(flagged_skus)) if flagged_skus else np.zeros(n_rows, dtype=bool)
    if bidirectional:
        # 雙向：轉出來源可以是任何 Nature（含 Tourist），需求列本身不列為來源
        donor_mask = (~need_mask) & donor_pool & ((donate_np > 0) | (transferable_np > 0))
    else:
        donor_mask = (~is_tourist) & donor_pool & ((donate_np > 0) | (transferable_np > 0))
    donated_mask = donor_mask & (donate_np > 0)
    donor_unused_mask = donor_mask & ~donated_mask

    under_mask = need_mask & (adj_val < safety_line)
    over_mask = donor_mask & (adj_val > safety_line)
    in_table = need_mask | donor_mask

    event = np.full(n_rows, "", dtype=object)
    event[need_mask] = _HYBRID_EVENT_NEED_IN
    event[donated_mask] = _HYBRID_EVENT_DONATED
    event[donor_unused_mask] = _HYBRID_EVENT_DONOR

    diag = np.full(n_rows, "", dtype=object)
    diag[need_mask] = np.where(under_mask[need_mask], _HYBRID_DIAG_UNDER, _HYBRID_DIAG_NORMAL_NEED)
    diag[donor_mask] = np.where(over_mask[donor_mask], _HYBRID_DIAG_OVER, _HYBRID_DIAG_NORMAL_DONOR)

    line_col = np.full(n_rows, np.nan)
    sugg = np.full(n_rows, np.nan)
    chg = np.full(n_rows, np.nan)
    line_col[in_table] = safety_line[in_table]
    fix_mask = under_mask | over_mask
    sugg[fix_mask] = safety_line[fix_mask]
    chg[fix_mask] = safety_line[fix_mask] - adj_val[fix_mask]

    stock_after = cm_stock_val + actual_np - donate_np
    td_after = np.full(n_rows, np.nan)
    ms_pos = msold_val > 0
    td_after[ms_pos] = stock_after[ms_pos] / msold_val[ms_pos] * 30.5
    stock_after_om = cm_stock_val + actual_om - donate_om
    td_after_om = np.full(n_rows, np.nan)
    td_after_om[ms_pos] = stock_after_om[ms_pos] / msold_val[ms_pos] * 30.5

    raw[_HYBRID_SAFETY_DAYS_COLUMN] = safety_days_np.astype(int)
    raw[_HYBRID_SAFETY_LINE_COLUMN] = line_col
    raw["事件類型"] = event
    raw["Safety Review"] = diag
    raw["建議 Safety 設定值"] = sugg
    raw["Safety 建議調整量"] = np.round(chg, 1)
    raw["調貨後庫存"] = np.where(in_table, stock_after, np.nan)
    raw["調貨後 TurnDay"] = np.where(in_table, np.round(td_after, 1), np.nan)
    raw["調貨後庫存 (同OM)"] = np.where(in_table, stock_after_om, np.nan)
    raw["調貨後 TurnDay (同OM)"] = np.where(in_table, np.round(td_after_om, 1), np.nan)

    if progress_callback:
        progress_callback(0.55, "建立 Review 表...")
    diagnosis = _build_hybrid_diagnosis(raw)
    diagnosis_om = _build_hybrid_diagnosis(raw, variant="om")
    diag_summary = _build_hybrid_diag_summary(raw, dimension=dim)
    sc_summary = _build_stock_checking_summary(raw, bidirectional=bidirectional, dimension=dim)

    stats = {
        "safety_days": dict(_safety_days),
        "dimension": dim,
        "turnday_threshold": int(turnday_threshold),
        "d001_threshold": int(d001_threshold),
        "need_rows": int(need_series.sum()),
        "need_skus": int(raw.loc[need_series, "SKU"].nunique()) if need_series.any() else 0,
        "total_demand": int(raw.loc[need_series, "調入需求量 (補到14日)"].sum()) if need_series.any() else 0,
        "total_actual": int(raw["實際調入量"].sum()),
        "total_gap": int(raw["調入缺口"].sum()),
        "nt_transfer": int(raw.loc[raw["Shop Nature"] == "Non Tourist", "轉出量"].sum()),
        "mix_transfer": int(raw.loc[raw["Shop Nature"] == "Mix", "轉出量"].sum()),
        "tourist_transfer": int(raw.loc[raw["Shop Nature"] == "Tourist", "轉出量"].sum()) if bidirectional else 0,
        "total_actual_om": int(actual_om.sum()),
        "total_gap_om": int(gap_om.sum()),
        "nt_transfer_om": int(raw.loc[raw["Shop Nature"] == "Non Tourist", "轉出量 (同OM)"].sum()),
        "mix_transfer_om": int(raw.loc[raw["Shop Nature"] == "Mix", "轉出量 (同OM)"].sum()),
        "tourist_transfer_om": int(raw.loc[raw["Shop Nature"] == "Tourist", "轉出量 (同OM)"].sum()) if bidirectional else 0,
        "bidirectional": bool(bidirectional),
        "donated_rows_om": int((donor_mask & (donate_om > 0)).sum()),
        "mapped_other_rows": int(mapped_other),
        "mapped_rp_rows": int(mapped_rp),
        "donor_rows": int(donor_mask.sum()),
        "donated_rows": int(donated_mask.sum()),
        "donor_unused_rows": int(donor_unused_mask.sum()),
        "under_rows": int(under_mask.sum()),
        "under_skus": int(raw.loc[under_mask, "SKU"].nunique()) if under_mask.any() else 0,
        "under_inc_total": int(round(float(np.nansum(chg[under_mask])))) if under_mask.any() else 0,
        "over_rows": int(over_mask.sum()),
        "over_skus": int(raw.loc[over_mask, "SKU"].nunique()) if over_mask.any() else 0,
        "over_dec_total": int(round(float(np.nansum(chg[over_mask])))) if over_mask.any() else 0,
        "normal_rows": int(in_table.sum()) - int(under_mask.sum()) - int(over_mask.sum()),
        "total_raw_rows": int(n_rows),
        "buyer_need_skus": _hybrid_buyer_need_stats(raw, need_series),
        "buyer_safety_fix": _hybrid_buyer_safety_stats(raw),
    }

    if progress_callback:
        progress_callback(0.7, "轉換完成")

    return raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, warnings


def _write_hybrid_diagnosis_sheet(ws, diagnosis: pd.DataFrame, progress_callback=None,
                                  progress_start: float = 0.42, progress_end: float = 0.72):
    """調貨 Review Safety Stock 主表工作表：店舖×SKU 逐列（僅含事件列）。
    progress 區間必須落在對應 generate_excel_* 的下一階段區間內，避免進度倒退。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    band_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    _text_keys = ["Brand", "SKU Desc", "Buyer", "Adj. SST", "Site Code", "事件類型", "Safety Review", "調貨來源 (Site Code)"]
    _decimal_keys = {"調貨後 TurnDay": "0.0", "Safety 建議調整量": "#,##0.##"}

    headers = list(diagnosis.columns)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[1].height = 30

    if diagnosis is None or diagnosis.empty:
        ws.cell(row=2, column=1, value="無符合資料").font = Font(size=12)
        _auto_width(ws)
        return

    total = len(diagnosis)
    for r_idx, (_, row) in enumerate(diagnosis.iterrows()):
        excel_row = r_idx + 2
        is_band = (r_idx % 2 == 1)
        for col_idx, h in enumerate(headers, 1):
            val = row.get(h)
            cell = ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)
            cell.border = thin_border
            if h in _text_keys:
                cell.alignment = left_wrap
            else:
                cell.alignment = center
            if isinstance(val, (int, float)) and not isinstance(val, bool) and h not in _text_keys:
                cell.number_format = _decimal_keys.get(h, "#,##0")
            if is_band:
                cell.fill = band_fill
        if progress_callback and (r_idx % 200 == 0 or r_idx == total - 1):
            p = progress_start + (progress_end - progress_start) * ((r_idx + 1) / max(total, 1))
            progress_callback(p, f"寫入調貨 Review 主表 ({r_idx+1}/{total})...")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{1 + total}"
    _auto_width(ws)


def _write_hybrid_transfer_detail_sheet(ws, detail: pd.DataFrame, progress_callback=None,
                                        progress_start: float = 0.90, progress_end: float = 0.94):
    """簡易調貨明細工作表：一列一筆來源店→接收店調貨（來源字串已拆行）。
    progress 區間必須落在 generate_excel_hybrid 的下一階段區間內，避免進度倒退。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    band_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    _text_keys = ["Brand", "SKU Desc",
                  "Target Site Code(接收店舖)", "Target Shop Name",
                  "Supply Site Code(調貨來源)", "Supply Shop Name"]

    headers = list(_HYBRID_TRANSFER_DETAIL_COLUMNS)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[1].height = 30

    if detail is None or detail.empty:
        ws.cell(row=2, column=1, value="無符合資料").font = Font(size=12)
        _auto_width(ws)
        return

    total = len(detail)
    for r_idx, (_, row) in enumerate(detail.iterrows()):
        excel_row = r_idx + 2
        is_band = (r_idx % 2 == 1)
        for col_idx, h in enumerate(headers, 1):
            val = row.get(h)
            cell = ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)
            cell.border = thin_border
            cell.alignment = left_wrap if h in _text_keys else center
            if isinstance(val, (int, float)) and not isinstance(val, bool) and h not in _text_keys:
                cell.number_format = "#,##0"
            if is_band:
                cell.fill = band_fill
        if progress_callback and (r_idx % 200 == 0 or r_idx == total - 1):
            p = progress_start + (progress_end - progress_start) * ((r_idx + 1) / max(total, 1))
            progress_callback(p, f"寫入簡易調貨明細 ({r_idx+1}/{total})...")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{1 + total}"
    _auto_width(ws)


def _write_hybrid_summary_block(ws, start_row, section_title, block, *, header_fill,
                                header_font, bold_font, section_fill, total_fill,
                                thin_border, center, left_wrap):
    """「Review 摘要」工作表的統計區塊：合併標題列 → 表頭列 → 資料列
    （首欄值為「總計」者加粗體與灰底）。樣式與 Safety Stock Review 的 SUMMARY
    （_write_summary_sheet，transformer.py:971）一致。回傳下一個區塊的標題列號。"""
    headers = list(block.columns)
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=len(headers))
    sc = ws.cell(row=start_row, column=1, value=section_title)
    sc.font = Font(bold=True, size=12, color="1E40AF")
    sc.fill = section_fill
    sc.alignment = Alignment(horizontal="left", vertical="center")

    header_row = start_row + 1
    if block.empty:
        ws.cell(row=header_row, column=1, value="無符合資料").font = Font(size=12)
        return header_row + 2

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border

    for offset, (_, row) in enumerate(block.iterrows(), 1):
        is_total = str(row.get(headers[0])) == "總計"
        for col_idx, h in enumerate(headers, 1):
            val = row.get(h)
            cell = ws.cell(row=header_row + offset, column=col_idx,
                           value=None if (isinstance(val, float) and pd.isna(val)) else val)
            cell.border = thin_border
            cell.alignment = left_wrap if col_idx == 1 else center
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cell.number_format = "#,##0"
            if is_total:
                cell.font = bold_font
                cell.fill = total_fill
    return header_row + len(block) + 2


def _write_hybrid_buyer_combined_table(ws, combined, thin_border, header_fill,
                                        header_font, bold_font, section_fill,
                                        total_fill, center, start_row: int) -> int:
    """Review 摘要的 Buyer 合併統計表（Buyer／需調貨 SKU 數／需調貨 SKU 佔比／
    調入需求量合計／需調整 Safety SKU 數／需調整 Safety SKU 佔比／
    Safety 建議調整量合計＋總計列）。

    combined 為 [(Buyer, 需調貨 SKU 數, 調入需求量合計, 需調整 Safety SKU 數,
    Safety 建議調整量合計)]；舊 4 元組（無貨量）亦相容，貨量以 0 補齊。
    缺席值以 0 補齊；總計列為各欄合計（貨量與調整量為淨額；佔比總計為 100%）。
    空集合時不寫入，回傳原 start_row。回傳下一個區塊的標題列號。
    """
    if not combined:
        return start_row
    norm: list[tuple[str, int, int, int, int]] = []
    for row in combined:
        if len(row) == 5:
            b, n, d, f, v = row
            norm.append((str(b), int(n), int(d), int(f), int(v)))
        else:
            b, n, f, v = row  # type: ignore[misc]
            norm.append((str(b), int(n), 0, int(f), int(v)))
    combined = norm
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
    tc = ws.cell(row=start_row, column=1, value=_HYBRID_BUYER_COMBINED_TITLE)
    tc.font = Font(bold=True, size=12, color="1E40AF")
    tc.fill = section_fill
    tc.alignment = Alignment(horizontal="left", vertical="center")
    hr = start_row + 1
    for col_idx, h in enumerate(
        ("Buyer", "需調貨 SKU 數", "需調貨 SKU 佔比", "調入需求量合計",
         "需調整 Safety SKU 數", "需調整 Safety SKU 佔比", "Safety 建議調整量合計"), 1):
        cell = ws.cell(row=hr, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    total_need = sum(n for _, n, _, _, _ in combined)
    total_demand = sum(d for _, _, d, _, _ in combined)
    total_fix = sum(f for _, _, _, f, _ in combined)
    total_chg = sum(v for _, _, _, _, v in combined)
    r = hr + 1
    for buyer, need_cnt, demand_total, fix_cnt, chg_total in combined:
        c1 = ws.cell(row=r, column=1, value=buyer)
        c1.border = thin_border
        c2 = ws.cell(row=r, column=2, value=need_cnt)
        c2.number_format = "#,##0"
        c2.alignment = center
        c2.border = thin_border
        c3 = ws.cell(row=r, column=3, value=need_cnt / total_need if total_need else None)
        c3.number_format = "0.0%"
        c3.alignment = center
        c3.border = thin_border
        c4 = ws.cell(row=r, column=4, value=demand_total)
        c4.number_format = "#,##0"
        c4.alignment = center
        c4.border = thin_border
        c5 = ws.cell(row=r, column=5, value=fix_cnt)
        c5.number_format = "#,##0"
        c5.alignment = center
        c5.border = thin_border
        c6 = ws.cell(row=r, column=6, value=fix_cnt / total_fix if total_fix else None)
        c6.number_format = "0.0%"
        c6.alignment = center
        c6.border = thin_border
        c7 = ws.cell(row=r, column=7, value=chg_total)
        c7.number_format = "#,##0"
        c7.alignment = center
        c7.border = thin_border
        r += 1
    for col_idx, v in ((1, "總計"), (2, total_need), (3, 1.0 if total_need else None),
                       (4, total_demand), (5, total_fix), (6, 1.0 if total_fix else None),
                       (7, total_chg)):
        cell = ws.cell(row=r, column=col_idx, value=v)
        cell.font = bold_font
        cell.fill = total_fill
        cell.border = thin_border
        cell.alignment = center
        if col_idx in (2, 4, 5, 7):
            cell.number_format = "#,##0"
        elif col_idx in (3, 6):
            cell.number_format = "0.0%"
    return r + 2


def _write_hybrid_diag_summary_sheet(ws, cross: pd.DataFrame, stats: dict | None,
                                     raw: pd.DataFrame, thin_border, desc_font,
                                     progress_callback=None):
    """Review 摘要工作表：指標卡 4 排（第 4 排為同OM調貨對照；雙向調貨時另加第 5 排
    Tourist 轉出量）＋ ① Safety Review × 分類維度交叉表（維度由 stats["dimension"] 推定，
    預設 Shop Nature）＋ ② Safety 調整方向分類
    ＋ ③ Safety 調整明細（SKU 級彙總）＋ ④ Buyer 統計（需調貨＋需調整 Safety 合併表）。
    ②③ 由 raw 即時彙算（Safety 建議不受 OM 限制影響，兩變體共用同一份統計）；
    ④ 與 Stock Checking 的 Buyer 統計同口徑（同 SKU 首個非空 Buyer，已刪除重複）。"""
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    section_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    _rows = _hybrid_dashboard_metrics(raw, stats)
    row5_metrics = _rows[4] if len(_rows) > 4 else None
    _days = (stats or {}).get("safety_days") or _hybrid_safety_days_from_raw(raw)
    _write_metric_cards(
        ws, f"Safety Stock 設定 Review 摘要 (Transfer Diagnosis Dashboard｜"
            f"Safety 基準 {_hybrid_safety_days_label(_days)})",
        _rows[0], _rows[1], thin_border, desc_font, progress_callback,
        progress_pct=0.80, row3_metrics=_rows[2], row4_metrics=_rows[3],
        row5_metrics=row5_metrics,
    )

    table_start = 24 if row5_metrics else 20
    # 三個區塊共用同一母體（事件類型非空），故 ① 為空即代表全無資料
    if cross is None or cross.empty:
        ws.cell(row=table_start, column=1, value="無符合資料").font = Font(size=12)
        _auto_width(ws)
        return

    _block_kw = dict(header_fill=header_fill, header_font=header_font, bold_font=bold_font,
                     section_fill=section_fill, total_fill=total_fill, thin_border=thin_border,
                     center=center, left_wrap=left_wrap)
    _dim = _hybrid_dimension((stats or {}).get("dimension"))
    next_row = _write_hybrid_summary_block(
        ws, table_start - 1, f"Safety Review × {_HYBRID_DIMENSION_LABELS[_dim]} 交叉表", cross, **_block_kw)
    next_row = _write_hybrid_summary_block(
        ws, next_row, "Safety 調整方向分類（增加 / 減少 / 不變）",
        _build_hybrid_safety_change_summary(raw), **_block_kw)
    next_row = _write_hybrid_summary_block(
        ws, next_row, "Safety 調整明細（SKU 級彙總）",
        _build_hybrid_safety_sku_detail(raw), **_block_kw)

    try:
        combined = _hybrid_buyer_combined_detail(raw)
    except Exception:
        buyer_need = (stats or {}).get("buyer_need_skus")
        if buyer_need is None:
            buyer_need = _hybrid_buyer_need_stats(raw)
        buyer_safety = (stats or {}).get("buyer_safety_fix")
        if buyer_safety is None:
            buyer_safety = _hybrid_buyer_safety_stats(raw)
        combined = _hybrid_buyer_combined_stats(buyer_need, buyer_safety)
    _table_kw = dict(thin_border=thin_border, header_fill=header_fill,
                     header_font=header_font, bold_font=bold_font,
                     section_fill=section_fill, total_fill=total_fill,
                     center=center)
    if combined:
        next_row = _write_hybrid_buyer_combined_table(ws, combined, start_row=next_row, **_table_kw)

    if progress_callback:
        progress_callback(0.90, "Review 摘要完成...")
    _auto_width(ws)


# RAW 顯示欄序分組（識別→品項屬性→庫存與銷售→SKU 級判定→Review 結果→調貨 不限OM→調貨 同OM→其他）
_HYBRID_RAW_DISPLAY_GROUPS = [
    ["Buyer", "Brand", "SKU", "SKU Desc", "Site Code", "Shop Name", "OM", "Shop Nature"],
    ["Adj. SST", "Adj. MCH1", "Vendor MC", "MCH2", "Item Status", "Item Class", "RP Type",
     "Shop Stock Status", "Supply Source", "Launch Date"],
    ["Shop Stock Qty (CM)", "Adj. Safety Stock", "Sold Qty (P90D)", "Shop Monthly Sold Qty (P90D)",
     "Turnday (P90D Qty)", "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
     "Net Sales (M-2)", "Net Sales (M-1)", "Net Sales (CM - MTD)", "Excess Stock Qty",
     "Shop On Order PO Qty (CM)", "Shop Pending Receive",
     "Effective Shop Monthly Sold Qty (P90D)",
     "D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)", "D001 Effective Stock Qty (CM)"],
    ["Tourist 加總 TurnDay", "Mix 加總 TurnDay", "Non Tourist 加總 TurnDay",
     "需要調貨", "調入需求量 (補到14日)"],
    ["事件類型", _HYBRID_SAFETY_DAYS_COLUMN, _HYBRID_SAFETY_LINE_COLUMN,
     "Safety Review", "建議 Safety 設定值", "Safety 建議調整量"],
    ["可轉出貨量 (Non Tourist)", "可轉出貨量 (Mix)", "可轉出貨量 (Tourist)",
     "實際調入量", "調入缺口", "轉出量", "調貨來源 (Site Code)", "調貨後庫存", "調貨後 TurnDay"],
    ["實際調入量 (同OM)", "調入缺口 (同OM)", "轉出量 (同OM)", "調貨來源 (同OM)",
     "調貨後庫存 (同OM)", "調貨後 TurnDay (同OM)"],
]


def _hybrid_raw_display_columns(raw: pd.DataFrame) -> list[str]:
    """Safety Stock x NetStock Review RAW 顯示欄序：不顯示「Item_Shop」與所有含「Value」的金額欄，
    其餘欄位（含 _compute_raw_extra_columns 衍生的 Shop Name / SKU Desc）按 _HYBRID_RAW_DISPLAY_GROUPS
    分組排列，未歸組的原始欄位按原序置於最後。"""
    cols = [str(c) for c in raw.columns] + ["Shop Name", "SKU Desc"]
    seen: set[str] = set()
    kept: list[str] = []
    for c in cols:
        if c in seen or c == "Item_Shop" or "value" in c.lower():
            continue
        seen.add(c)
        kept.append(c)
    kept_set = set(kept)
    ordered: list[str] = []
    for group in _HYBRID_RAW_DISPLAY_GROUPS:
        for c in group:
            if c in kept_set:
                ordered.append(c)
                kept_set.discard(c)
    ordered += [c for c in kept if c in kept_set]
    return ordered


def generate_excel_hybrid(
    raw: pd.DataFrame,
    sc_summary: pd.DataFrame | None = None,
    diagnosis: pd.DataFrame | None = None,
    diagnosis_om: pd.DataFrame | None = None,
    diag_summary: pd.DataFrame | None = None,
    stats: dict | None = None,
    progress_callback=None,
) -> BytesIO:
    """Safety Stock x NetStock Review 輸出（8 工作表，順序固定）：
    RAW（重新排列欄序提升可讀性，不顯示 Item_Shop 與 Value 金額欄；含不限OM與同OM兩套調貨結果；
    雙向時另含「可轉出貨量 (Tourist)」）/
    Summary（SKU 級彙總，按 stats["dimension"] 分類維度分組：nature 三組／rp 兩組／both 六組；
    各組加總/TurnDay/涉及店舖數/佔比、強制調貨提示、存貨偏重提示；nature 雙向時另加 AD 欄
    「由Tourist 店舖預計調出數量」）/
    調貨 Review Safety Stock (不限OM) / 調貨 Review Safety Stock (同OM)（店舖×SKU 主表含 RP Type 欄）/
    店舖問題 SKU（每間店舖需行動列：問題＋建議方法）/
  Review 摘要（指標卡 4 排（雙向時 5 排）+ Safety Review × 分類維度交叉表
 + Safety 調整方向分類 + Safety 調整明細（SKU 級彙總）
 + Buyer 統計（需調貨＋需調整 Safety 合併表））/
    簡易調貨明細 (不限OM) / 簡易調貨明細 (同OM)（一列一筆來源店→接收店調貨，來源字串已拆行）。
    雙向與否由 RAW 是否含「可轉出貨量 (Tourist)」欄推定；分類維度由 stats["dimension"] 推定
    （缺失時退回 nature）。
    """
    if diagnosis is None:
        diagnosis = _build_hybrid_diagnosis(raw)
    if diagnosis_om is None:
        diagnosis_om = _build_hybrid_diagnosis(raw, variant="om")
    _dim = _hybrid_dimension((stats or {}).get("dimension"))
    if diag_summary is None:
        diag_summary = _build_hybrid_diag_summary(raw, dimension=_dim)
    _bidirectional = raw is not None and "可轉出貨量 (Tourist)" in raw.columns
    if sc_summary is None:
        sc_summary = _build_stock_checking_summary(raw, bidirectional=_bidirectional, dimension=_dim)
    wb = Workbook()
    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    desc_font = Font(size=9, color="6B7280", italic=True)

    # ── Sheet 1: RAW（重排欄序；不顯示 Item_Shop 與 Value 金額欄）──
    ws_raw = wb.active
    ws_raw.title = "RAW"
    _write_raw_sheet(ws_raw, raw, progress_callback, progress_start=0.0, progress_end=0.24,
                     columns=_hybrid_raw_display_columns(raw))

    if progress_callback:
        progress_callback(0.25, "建立 SKU Summary（存貨偏重提示）...")

    # ── Sheet 2: Summary（SKU 級彙總，按分類維度分組；雙向時另加 AD 欄；偏重提示以醒目橘底加強顯示）──
    ws_sc = wb.create_sheet(title="Summary")
    _write_stock_checking_summary_sheet(ws_sc, sc_summary, progress_callback, progress_start=0.25, progress_end=0.38, bidirectional=_bidirectional, highlight_skew=True, dimension=_dim)

    if progress_callback:
        progress_callback(0.40, "建立調貨 Review 主表 (不限OM)...")

    # ── Sheet 3: 調貨 Review Safety Stock（不限OM）──
    ws_diag = wb.create_sheet(title="調貨 Review Safety Stock (不限OM)")
    _write_hybrid_diagnosis_sheet(ws_diag, diagnosis, progress_callback, progress_start=0.40, progress_end=0.56)

    if progress_callback:
        progress_callback(0.57, "建立調貨 Review 主表 (同OM)...")

    # ── Sheet 4: 調貨 Review Safety Stock（同OM）──
    ws_diag_om = wb.create_sheet(title="調貨 Review Safety Stock (同OM)")
    _write_hybrid_diagnosis_sheet(ws_diag_om, diagnosis_om, progress_callback, progress_start=0.57, progress_end=0.70)

    if progress_callback:
        progress_callback(0.71, "建立店舖問題 SKU...")

    # ── Sheet 5: 店舖問題 SKU（只收 Safety 設定不足／過高；空表寫「無符合資料」）──
    ws_shop_issue = wb.create_sheet(title="店舖問題 SKU")
    _write_shop_issue_sheet(
        ws_shop_issue, _build_hybrid_shop_issues(raw),
        title="店舖問題 SKU — 每間店舖需行動列（問題＋建議方法）",
        progress_callback=progress_callback, progress_start=0.71, progress_end=0.74,
    )

    if progress_callback:
        progress_callback(0.75, "建立 Review 摘要...")

    # ── Sheet 6: Review 摘要 ──
    ws_sum = wb.create_sheet(title="Review 摘要")
    _write_hybrid_diag_summary_sheet(ws_sum, diag_summary, stats, raw, thin_border, desc_font, progress_callback)

    if progress_callback:
        progress_callback(0.90, "建立簡易調貨明細 (不限OM)...")

    # ── Sheet 6: 簡易調貨明細（不限OM）──
    ws_detail = wb.create_sheet(title="簡易調貨明細 (不限OM)")
    _write_hybrid_transfer_detail_sheet(
        ws_detail, _build_hybrid_transfer_detail(raw, variant="free"),
        progress_callback, progress_start=0.90, progress_end=0.93,
    )

    if progress_callback:
        progress_callback(0.93, "建立簡易調貨明細 (同OM)...")

    # ── Sheet 7: 簡易調貨明細（同OM）──
    ws_detail_om = wb.create_sheet(title="簡易調貨明細 (同OM)")
    _write_hybrid_transfer_detail_sheet(
        ws_detail_om, _build_hybrid_transfer_detail(raw, variant="om"),
        progress_callback, progress_start=0.93, progress_end=0.95,
    )

    if progress_callback:
        progress_callback(0.96, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf




# ─── Stock Impact Analysis (Simplified) Functions ─────────────────────────────
# 簡化版本：只輸出 RAW 和 SKU Monthly Sales&Store 兩個工作表
# SKU Monthly Sales&Store 不分 Shop Nature，只顯示 M-2, M-1, MTD 的 D001 存貨、店舖存貨和總銷售量


def _write_stock_impact_dashboard_simplified(ws, raw_df, thin_border, desc_font, progress_callback, sku_monthly_analysis=None, stock_threshold=1000, sku_soh_summary=None):
    """簡化版 Dashboard（含低存貨分析 + 有賣VS有貨比率區間）。"""
    nature = raw_df.get("Shop Nature", pd.Series(dtype=object)).fillna("").astype(str)
    cm_stock = _safe_numeric(raw_df.get("Shop Stock Qty (CM)", _safe_default_series(0, raw_df.index)))
    turnday = pd.to_numeric(raw_df.get("Turnday (Days)", _safe_default_series(0, raw_df.index)), errors="coerce")

    # Tourist 缺貨行數（高銷售店無貨）
    tourist_oos_rows = int(((nature == "Tourist") & (cm_stock <= 0) & raw_df["Is Stock Impacted"]).sum()) if "Is Stock Impacted" in raw_df else 0

    # Tourist 缺貨數量（缺失件數 = max(0, Adj. Safety Stock - Shop Stock Qty (CM))）
    _adj_ss = _safe_numeric(raw_df.get("Adj. Safety Stock", _safe_default_series(0, raw_df.index)))
    _oos_mask = (nature == "Tourist") & (cm_stock <= 0) & (raw_df.get("Is Stock Impacted", pd.Series(False, index=raw_df.index)))
    tourist_oos_qty = int(np.maximum(0, (_adj_ss.values - cm_stock.values))[_oos_mask.values].sum())

    # 曾缺貨 Tourist 行數（含已恢復，不限 CM 當期缺貨）
    tourist_oos_broader = int(((nature == "Tourist") & raw_df["Is Stock Impacted"]).sum()) if "Is Stock Impacted" in raw_df else 0

    # Non Tourist 可釋放存貨件數（>14日標準）
    nt_excess_total = int(_safe_numeric(raw_df.get("Non Tourist Excess Stock", _safe_default_series(0, raw_df.index))).sum())

    # Mix 可釋放存貨件數（>14日標準）
    mix_excess_total = int(_safe_numeric(raw_df.get("Mix Excess Stock", _safe_default_series(0, raw_df.index))).sum())

    # 各 Nature 平均 Turnday
    tourist_td = turnday[nature == "Tourist"]
    tourist_avg_td = round(float(tourist_td.mean()), 1) if tourist_td.notna().any() and len(tourist_td) else 0
    mix_td = turnday[nature == "Mix"]
    mix_avg_td = round(float(mix_td.mean()), 1) if mix_td.notna().any() and len(mix_td) else 0
    nt_td = turnday[nature == "Non Tourist"]
    nt_avg_td = round(float(nt_td.mean()), 1) if nt_td.notna().any() and len(nt_td) else 0

    row1_metrics = [
        ("⭐ Tourist 缺貨行數", tourist_oos_rows, "高銷售店無貨（只計 RF）｜單位：行數", "DC2626"),
        ("Non Tourist 可釋存貨", nt_excess_total, "超過14日標準可釋放｜單位：件數", "F59E0B"),
        ("Mix 可釋放存貨", mix_excess_total, "超過14日標準可釋放｜單位：件數", "F97316"),
        ("📉 Tourist 缺貨數量", tourist_oos_qty, "缺失件數 = Safety − Shop Net Stock｜單位：件數", "DC2626"),
    ]

    row2_metrics = [
        ("Tourist 平均Turnday", tourist_avg_td, "Tourist 平均存貨天數｜單位：日", "8B5CF6"),
        ("Mix 平均Turnday", mix_avg_td, "Mix 平均存貨天數｜單位：日", "3B82F6"),
        ("Non Tourist 平均Turnday", nt_avg_td, "Non Tourist 平均存貨天數｜單位：日", "06B6D4"),
        ("🕐 曾缺貨Tourist行數(含已恢復)", tourist_oos_broader, "任何一期曾缺貨的Tourist RF行數｜含已恢復供貨", "6366F1"),
    ]

    low_stock_total = 0
    low_stock_item_class = pd.DataFrame()
    if sku_monthly_analysis is not None and not sku_monthly_analysis.empty and "低存貨" in sku_monthly_analysis.columns:
        low_stock_total = int(sku_monthly_analysis["低存貨"].sum())
        low_stock_item_class = _build_low_stock_item_class_summary(sku_monthly_analysis)

    # 有賣VS有貨新指標（從 sku_soh_summary 計算）
    cm_lt50 = 0
    cm_lt30 = 0
    cm_no_stock = 0
    has_band_data = sku_soh_summary is not None and not sku_soh_summary.empty
    if has_band_data:
        cm_oh = pd.to_numeric(sku_soh_summary.get("有貨店舖 (CM)", pd.Series(0)), errors="coerce").fillna(0)
        cm_ratio = pd.to_numeric(sku_soh_summary.get("有賣VS有貨 (CM)", pd.Series(np.nan)), errors="coerce")
        cm_lt50 = int(((cm_oh > 0) & (cm_ratio < 0.5)).sum())
        cm_lt30 = int(((cm_oh > 0) & (cm_ratio < 0.3)).sum())
        cm_no_stock = int((cm_oh == 0).sum())

    row3_metrics = [
        ("🔴 低存貨 SKU 總數", low_stock_total, f"D001+店舖存貨<{stock_threshold}件｜單位：SKU數", "DC2626"),
        ("📉 CM 有賣/有貨<50% SKU", cm_lt50, f"有貨店舖>0 且 有賣÷有貨<50%｜單位：SKU數", "F97316"),
        ("🔻 CM 有賣/有貨<30% SKU", cm_lt30, "有貨店舖>0 且 有賣÷有貨<30%｜單位：SKU數", "DC2626"),
        ("⬜ CM 完全無貨 SKU", cm_no_stock, "有貨店舖 (CM)=0｜單位：SKU數", "6B7280"),
    ]

    _write_metric_cards(ws, "缺貨銷售影響分析摘要 (Stock Impact Dashboard - 簡化)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback, row3_metrics=row3_metrics)

    # 有賣VS有貨比率區間分佈表
    if has_band_data:
        band_df = _build_sell_through_band_summary(sku_soh_summary)
        band_start = 15
        ws.cell(row=band_start, column=1, value="📊 有賣VS有貨比率分佈（按 SKU 數）").font = Font(bold=True, size=12, color="1E40AF")
        band_header_row = band_start + 1
        for col_idx, h in enumerate(band_df.columns, 1):
            cell = ws.cell(row=band_header_row, column=col_idx, value=h)
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, (_, row_data) in enumerate(band_df.iterrows()):
            excel_row = band_header_row + 1 + row_idx
            is_total = row_data.get("有賣VS有貨比率區間") == "總計"
            for col_idx, h in enumerate(band_df.columns, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
                if is_total:
                    cell.font = Font(bold=True)
        item_class_start = band_header_row + 1 + len(band_df) + 1  # skip 1 blank row after band table
    else:
        item_class_start = 15

    # 低存貨 SKU 按 Item Class 分類（已下移）
    if low_stock_item_class is not None and not low_stock_item_class.empty:
        ws.cell(row=item_class_start, column=1, value="🔴 低存貨 SKU 按 Item Class 分類").font = Font(bold=True, size=12, color="DC2626")
        header_row = item_class_start + 1
        for col_idx, h in enumerate(low_stock_item_class.columns, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, (_, row_data) in enumerate(low_stock_item_class.iterrows()):
            excel_row = header_row + 1 + row_idx
            is_total = row_data.get("Item Class") == "總計"
            for col_idx, h in enumerate(low_stock_item_class.columns, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
                if is_total:
                    cell.font = Font(bold=True)
    else:
        ws.cell(row=item_class_start, column=1, value="🔴 低存貨 SKU 按 Item Class 分類：無低存貨 SKU").font = Font(size=12, color="DC2626")

    _auto_width(ws)

def _build_sku_soh_tier_summary(raw: pd.DataFrame) -> pd.DataFrame:
    """建立 'New' sheet（Joseph 格式）：每 SKU 一行的 SOH Tier 分佈 + 有賣VS有貨比率。

    回傳 DataFrame 含 52 欄，排序：Total Stock Qty (CM) 降序。
    """
    if raw.empty or "SKU" not in raw.columns:
        return pd.DataFrame()

    # 先 enrich 取得擴充欄位
    df = _enrich_raw_joseph(raw)

    # 基本資訊 lookup
    sku_lookup = _build_sku_info_lookup(df)
    base_cols = ["SKU"]
    for c in ["SKU Desc", "Brand", "Buyer", "Item Class"]:
        if c in sku_lookup.columns:
            base_cols.append(c)
    info = sku_lookup[base_cols].copy()
    info["Status"] = sku_lookup.get("Item Status", "")
    info["Launch Date"] = sku_lookup.get("Launch Date", "")
    info["Is New SKU (3mo)"] = sku_lookup.get("Is New SKU (3mo)", "")

    # SKU 級彙總
    g = df.groupby("SKU", dropna=False)
    # Eff Price / Price Tier / Total Stock Qty (first)
    _first_cols = [c for c in ["Eff Price", "Price Tier"] if c in df.columns]
    for p in ["M-2", "M-1", "CM"]:
        c = f"Total Stock Qty ({p})"
        if c in df.columns:
            _first_cols.append(c)
    sku_first = g[_first_cols].first().reset_index() if _first_cols else pd.DataFrame({"SKU": g.size().index})

    # 有貨/有賣店舖 (sum of 0/1 flags)
    stock_sold = {}
    for p in ["M-2", "M-1", "CM"]:
        stock_sold[f"有貨店舖 ({p})"] = (f"有貨店舖 ({p})", "sum")
        stock_sold[f"有賣店舖 ({p})"] = (f"有賣店舖 ({p})", "sum")
    ss_agg = g.agg(**stock_sold).reset_index()

    # Sold Qty 加總
    sold_cols = {}
    for p in ["M-2", "M-1"]:
        sold_cols[f"Sold Qty ({p})"] = (f"Sold Qty ({p})", "sum")
    sold_cols["Sold Qty (CM - MTD)"] = ("Sold Qty (CM - MTD)", "sum")
    sold_agg = g.agg(**sold_cols).reset_index()

    # SOH Tier 各期分佈 (value_counts → 寬表)
    tier_periods = []
    for p in ["M-2", "M-1", "CM"]:
        tier_col = f"SOH Tier ({p})"
        if tier_col not in df.columns:
            continue
        # count rows per SKU per tier
        vc = df.groupby(["SKU", tier_col], dropna=False).size().unstack(fill_value=0)
        # ensure all tier labels present
        for lbl in _SOH_TIER_LABELS:
            if lbl not in vc.columns:
                vc[lbl] = 0
        vc = vc[_SOH_TIER_LABELS].copy()
        vc.columns = [f"{lbl}" for lbl in _SOH_TIER_LABELS]
        # add total
        vc["總計"] = vc.sum(axis=1).astype(int)
        vc = vc.reset_index()
        vc.columns = ["SKU"] + list(_SOH_TIER_LABELS) + ["總計"]
        tier_periods.append(vc)

    # 合併
    result = info.merge(sku_first, on="SKU", how="left")
    result = result.merge(ss_agg, on="SKU", how="left")
    result = result.merge(sold_agg, on="SKU", how="left")

    # 有賣VS有貨比率 (cap 1, div/0 → NaN)
    for p in ["M-2", "M-1", "CM"]:
        on_hand = result.get(f"有貨店舖 ({p})", pd.Series(0, index=result.index))
        sold = result.get(f"有賣店舖 ({p})", pd.Series(0, index=result.index))
        ratio = np.where(
            on_hand > 0,
            np.minimum(sold / on_hand, 1.0),
            np.nan,
        )
        result[f"有賣VS有貨 ({p})"] = ratio

    # 低存量標記 (< 1000 件)
    cm_stock = pd.to_numeric(result.get("Total Stock Qty (CM)", pd.Series(0, index=result.index)), errors="coerce").fillna(0)
    result["Low Stock (CM<1000)"] = np.where(cm_stock < 1000, "Y", "")

    # 低有賣VS有貨率標記 — 只標有貨（有貨店舖>0）且比率<50%嘅 SKU
    on_hand_cm = pd.to_numeric(result.get("有貨店舖 (CM)", pd.Series(0, index=result.index)), errors="coerce").fillna(0)
    cm_ratio = pd.to_numeric(result.get("有賣VS有貨 (CM)", pd.Series(np.nan, index=result.index)), errors="coerce")
    result["Low 有賣VS有貨 (CM<50%)"] = np.where((on_hand_cm > 0) & (cm_ratio < 0.5), "Y", "")

    # 合併 tier 區塊
    for tp in tier_periods:
        result = result.merge(tp, on="SKU", how="left")

    # 排序
    sort_col = "Total Stock Qty (CM)"
    if sort_col in result.columns:
        result = result.sort_values(sort_col, ascending=False).reset_index(drop=True)

    # 欄位順序 (52+2 columns, matching Joseph + new marker columns)
    out = ["SKU", "SKU Desc", "Brand", "Buyer", "Item Class", "Status",
           "Launch Date", "Is New SKU (3mo)",
           "Low Stock (CM<1000)", "Low 有賣VS有貨 (CM<50%)",
           "Eff Price", "Price Tier"]
    for p in ["M-2", "M-1", "CM"]:
        out.append(f"Total Stock Qty ({p})")
    for p in ["M-2", "M-1", "CM"]:
        out.append(f"有貨店舖 ({p})")
    for p in ["M-2", "M-1", "CM"]:
        out.append(f"有賣店舖 ({p})")
    for p in ["M-2", "M-1", "CM"]:
        out.append(f"有賣VS有貨 ({p})")
    if "Sold Qty (M-2)" in result.columns:
        out.append("Sold Qty (M-2)")
    if "Sold Qty (M-1)" in result.columns:
        out.append("Sold Qty (M-1)")
    if "Sold Qty (CM - MTD)" in result.columns:
        out.append("Sold Qty (CM - MTD)")
    for p, pname in [("M-2", "M-2"), ("M-1", "M-1"), ("CM", "CM")]:
        for lbl in _SOH_TIER_LABELS:
            col = f"{lbl}"
            if col in result.columns:
                out.append(col)
        out.append("總計")

    out = [c for c in out if c in result.columns]
    result = result[out]

    return result


def _build_sell_through_band_summary(sku_soh_summary: pd.DataFrame) -> pd.DataFrame:
    """建立「有賣VS有貨比率區間分佈表」：三期 × 5 區間（含 無貨行 + 總計行）。

    輸入 = _build_sku_soh_tier_summary 的輸出（New sheet df）。
    回傳 6 行 × 4 欄 DataFrame：比率區間／三期各區間 SKU 數。
    """
    if sku_soh_summary.empty:
        return pd.DataFrame()

    required = []
    for p in ["M-2", "M-1", "CM"]:
        required.append(f"有貨店舖 ({p})")
        required.append(f"有賣VS有貨 ({p})")
    if not all(c in sku_soh_summary.columns for c in required):
        return pd.DataFrame()

    bands = [
        ("<30%", lambda oh, r: (oh > 0) & (r < 0.3)),
        ("30-50%", lambda oh, r: (oh > 0) & (r >= 0.3) & (r < 0.5)),
        ("50-70%", lambda oh, r: (oh > 0) & (r >= 0.5) & (r < 0.7)),
        (">=70%", lambda oh, r: (oh > 0) & (r >= 0.7)),
        ("無貨（有貨=0）", lambda oh, r: oh == 0),
    ]
    periods = ["M-2", "M-1", "CM"]

    rows = []
    for label, mask_fn in bands:
        row = {"有賣VS有貨比率區間": label}
        for p in periods:
            oh = pd.to_numeric(sku_soh_summary[f"有貨店舖 ({p})"], errors="coerce").fillna(0)
            r = pd.to_numeric(sku_soh_summary[f"有賣VS有貨 ({p})"], errors="coerce")
            row[f"SKU數 ({p})"] = int(mask_fn(oh, r).sum())
        rows.append(row)

    band_df = pd.DataFrame(rows)
    total_row = {"有賣VS有貨比率區間": "總計"}
    for p in periods:
        total_row[f"SKU數 ({p})"] = band_df[f"SKU數 ({p})"].sum()
    band_df = pd.concat([band_df, pd.DataFrame([total_row])], ignore_index=True)
    return band_df


def _write_new_sheet(ws, df, header_font, header_fill, thin_border, progress_callback):
    """寫入 'New' sheet（Joseph 格式），含 merged headers for SOH tier blocks。

    格式化：
    - Launch Date: 轉為純文字（去除 .0）
    - 有賣VS有貨: 百分比格式 + 條件格式（<30% 紅底白字粗體 / 30-50% 橘底 / 50-70% 淺黃底）
    - 有貨店舖: >70 紅底白字粗體
    """
    if df is None or df.empty:
        ws.cell(row=1, column=1, value="無 SOH Tier 分佈數據").font = Font(size=12)
        return

    col_list = list(df.columns)

    # 找出各特殊欄位的 1-based index
    launch_date_idx = None
    stock_col_idxs = []       # 有貨店舖 (M-2/M-1/CM)
    ratio_col_idxs = []       # 有賣VS有貨 (M-2/M-1/CM)
    total_stock_col_idxs = [] # Total Stock Qty (M-2/M-1/CM)
    marker_col_idxs = []      # Low Stock / Low 有賣VS有貨 標記欄
    for i, c in enumerate(col_list, 1):
        if c == "Launch Date":
            launch_date_idx = i
        elif c.startswith("有貨店舖"):
            stock_col_idxs.append(i)
        elif c.startswith("有賣VS有貨"):
            ratio_col_idxs.append(i)
        elif c.startswith("Total Stock Qty"):
            total_stock_col_idxs.append(i)
        elif c in ("Low Stock (CM<1000)", "Low 有賣VS有貨 (CM<50%)"):
            marker_col_idxs.append(i)

    # SOH tier merged row: find Sold Qty (CM - MTD) position
    sold_idx = None
    for i, c in enumerate(col_list):
        if c == "Sold Qty (CM - MTD)":
            sold_idx = i
            break

    # Row 3: merged labels for SOH tier blocks
    block_width = len(_SOH_TIER_LABELS) + 1  # 9
    start_col = sold_idx + 2 if sold_idx is not None else (len(col_list) - 3 * block_width + 1)
    if start_col < 1:
        start_col = 1
    for bi, pname in enumerate(["M-2", "M-1", "CM"]):
        try:
            c1 = start_col + bi * block_width
            c2 = c1 + block_width - 1
            ws.merge_cells(f"{get_column_letter(c1)}3:{get_column_letter(c2)}3")
            cell = ws.cell(row=3, column=c1, value=f"店舖SOH Tier ({pname})")
            cell.font = Font(bold=True, size=10, color="1E40AF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        except Exception:
            pass

    # Row 4: headers
    for col_idx, h in enumerate(col_list, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Conditional format fills
    red_fill = PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid")
    red_font = Font(bold=True, color="FFFFFF", size=11)
    orange_fill = PatternFill(start_color="F97316", end_color="F97316", fill_type="solid")
    yellow_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    total_rows = len(df)
    for row_idx, row_data in enumerate(df.itertuples(index=False), start=5):
        for col_idx, val in enumerate(row_data, 1):
            if pd.isna(val):
                val = None
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

            # Launch Date: 轉純文字（去 .0）
            if col_idx == launch_date_idx and val is not None:
                try:
                    cell.value = str(int(float(val)))
                except (ValueError, TypeError):
                    cell.value = str(val)

            elif col_idx in marker_col_idxs:
                cell.number_format = "@"
            elif isinstance(val, float):
                # Total Stock Qty: 整數格式（無 .00）
                if col_idx in total_stock_col_idxs:
                    cell.number_format = "#,##0"
                # 有賣VS有貨: 百分比格式 + 條件格式
                elif col_idx in ratio_col_idxs:
                    cell.number_format = "0.0%"
                    if val < 0.3:
                        cell.fill = red_fill
                        cell.font = red_font
                    elif val < 0.5:
                        cell.fill = orange_fill
                        cell.font = Font(bold=True, color="FFFFFF", size=11)
                    elif val < 0.7:
                        cell.fill = yellow_fill
                else:
                    cell.number_format = "#,##0.00"
            elif isinstance(val, int):
                cell.number_format = "#,##0"
                # 有貨店舖: >70 紅底白字粗體
                if col_idx in stock_col_idxs and val > 70:
                    cell.fill = red_fill
                    cell.font = red_font

        if progress_callback and ((row_idx - 5) % 500 == 0 or row_idx - 4 == total_rows):
            progress_callback(0.55 + 0.25 * ((row_idx - 4) / max(total_rows, 1)),
                              f"寫入 New sheet ({row_idx-4}/{total_rows})...")

    _auto_width(ws)


def _build_sku_monthly_analysis_simplified(raw: pd.DataFrame, stock_threshold: int = 1000) -> pd.DataFrame:
    """簡化版 SKU 月銷量和存儲量分析（同原版欄位，含新增不分 Nature 的週期存貨/銷售欄位）。

    直接使用 `_build_sku_monthly_analysis` 的完整輸出（已含補充資訊與週期欄位），
    不需要額外計算。
    """
    if raw.empty or "SKU" not in raw.columns:
        return pd.DataFrame()
    return _build_sku_monthly_analysis(raw, stock_threshold=stock_threshold)


def generate_excel_stock_impact_simplified(
    raw_df: pd.DataFrame,
    progress_callback=None,
    sku_monthly_analysis: pd.DataFrame | None = None,
    stock_threshold: int = 1000,
    sku_soh_summary: pd.DataFrame | None = None,
) -> BytesIO:
    """簡化版 Excel 輸出：輸出 RAW（Joseph 擴充欄位）、SKU Monthly Sales&Store、New、Stock Impact Dashboard 四個工作表。"""
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

    # 先 enrich raw (計算 Total Stock / 有貨有賣 / Eff Price / SOH Tier 等)
    enriched = _enrich_raw_joseph(raw_df)

    # Sheet 1: RAW (Joseph 擴充欄位順序)
    ws_raw = wb.active
    ws_raw.title = "RAW"
    _write_raw_sheet(ws_raw, enriched, progress_callback, progress_start=0.0, progress_end=0.2, columns=_JOSEPH_RAW_COLS)

    if progress_callback:
        progress_callback(0.25, "建立 SKU Monthly Sales&Store 工作表...")

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
                progress_callback(0.2 + 0.35 * ((row_idx - 4) / max(len(sku_monthly_analysis), 1)), f"寫入 SKU 月銷量 ({row_idx-4}/{len(sku_monthly_analysis)})...")
        _auto_width(ws_sku_monthly)
    else:
        ws_sku_monthly.cell(row=1, column=1, value="無 SKU 月銷量/存儲量數據").font = Font(size=12)

    if progress_callback:
        progress_callback(0.58, "建立 New sheet（SOH Tier 分佈）...")

    # Sheet 3: New (SOH Tier 分佈，Joseph 格式)
    ws_new = wb.create_sheet(title="New")
    if sku_soh_summary is not None and not sku_soh_summary.empty:
        new_df = sku_soh_summary
    else:
        new_df = _build_sku_soh_tier_summary(raw_df)
    _write_new_sheet(ws_new, new_df, header_font, header_fill, thin_border, progress_callback)

    if progress_callback:
        progress_callback(0.85, "建立 Stock Impact Dashboard 工作表...")

    # Sheet 4: Stock Impact Dashboard
    ws_dashboard = wb.create_sheet(title="Stock Impact Dashboard")
    if "Is Stock Impacted" in raw_df.columns and not raw_df.empty:
        _write_stock_impact_dashboard_simplified(ws_dashboard, raw_df, thin_border, desc_font, progress_callback, sku_monthly_analysis=sku_monthly_analysis, stock_threshold=stock_threshold, sku_soh_summary=new_df)
    else:
        ws_dashboard.cell(row=1, column=1, value="無足夠數據生成 Dashboard").font = Font(size=12)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf


# ─── Past 90 Days 0 Sales Functions ─────────────────────────────────────────────


_P90D_ZERO_SALES_REQUIRED = [
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Shop Stock Qty (LM)",
    "Shop Stock Qty (CM)",
    "Site Code",
]


_P90D_RAW_OUTPUT_COLUMNS = [
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "Brand",
    "SKU",
    "SKU Desc",
    "Vendor MC",
    "Item Status",
    "Launch Date",
    "Item Class",
    "RP Type",
    "Shop Stock Status",
    "Adj. Safety Stock",  # 顯示名稱: Safety Stock (現時)
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Sold Qty (P90D)",
    "Shop Stock Qty (CM)",
    "Shop Stock Qty (LM)",
    "Excess Stock Qty",
    "Shop On Order PO Qty (CM)",
    "Buyer",
]


def _p90d_raw_columns(raw: pd.DataFrame) -> list[str]:
    """輸出欄位清單：限定 Past 90 Days 0 Sales RAW 只輸出指定欄位與順序。"""
    r = _reorder_raw_columns(_compute_raw_extra_columns(raw))
    available = set(r.columns)
    return [c for c in _P90D_RAW_OUTPUT_COLUMNS if _resolve_column_name(c, available) is not None]


def _build_p90d_zero_sales_by_shop_summary(raw: pd.DataFrame) -> pd.DataFrame:
    if raw is None or raw.empty:
        return pd.DataFrame()
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
    df["_stock_lm"] = _safe_numeric(df.get("Shop Stock Qty (LM)", _safe_default_series(0, df.index)))
    df["_stock_cm"] = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, df.index)))

    by_shop = (
        df.groupby(["Site Code", "Shop Name", "OM"], dropna=False, observed=True)
        .agg(
            **{
                "Past 90 Days 0 Sales SKU 數量": ("SKU", "nunique"),
                "Shop Stock Qty (LM) 總數": ("_stock_lm", "sum"),
                "Shop Stock Qty (CM) 總數": ("_stock_cm", "sum"),
            }
        )
        .reset_index()
    )
    for c in ["Past 90 Days 0 Sales SKU 數量", "Shop Stock Qty (LM) 總數", "Shop Stock Qty (CM) 總數"]:
        by_shop[c] = by_shop[c].fillna(0)
    by_shop["Past 90 Days 0 Sales SKU 數量"] = by_shop["Past 90 Days 0 Sales SKU 數量"].astype(int)

    if not by_shop.empty:
        total_row = pd.DataFrame(
            {
                "Site Code": ["總計"],
                "Shop Name": [""],
                "OM": [""],
                "Past 90 Days 0 Sales SKU 數量": [int(df["SKU"].dropna().nunique())],
                "Shop Stock Qty (LM) 總數": [float(df["_stock_lm"].sum())],
                "Shop Stock Qty (CM) 總數": [float(df["_stock_cm"].sum())],
            }
        )
        by_shop = pd.concat(
            [by_shop.sort_values(["Site Code", "Shop Name", "OM"], kind="stable"), total_row],
            ignore_index=True,
        )
    return by_shop


def transform_p90d_zero_sales(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
):
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    validate_input(df_input)
    avail = set(df_input.columns)
    missing = [name for name in _P90D_ZERO_SALES_REQUIRED if _resolve_column_name(name, avail) is None]
    if missing:
        raise ValueError(f"缺少必需欄位: {', '.join(missing)}")

    if progress_callback:
        progress_callback(0.1, "清理資料中...")
    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if progress_callback:
        progress_callback(0.2, "解像 SKU Desc / Shop Name...")
    df = _compute_raw_extra_columns(df)

    sold_m2 = _safe_numeric(df.get("Sold Qty (M-2)", _safe_default_series(0, df.index)))
    sold_m1 = _safe_numeric(df.get("Sold Qty (M-1)", _safe_default_series(0, df.index)))
    sold_mtd = _safe_numeric(df.get("Sold Qty (CM - MTD)", _safe_default_series(0, df.index)))
    sold_p90d = _safe_numeric(df.get("Sold Qty (P90D)", _safe_default_series(0, df.index)))
    stock_lm = _safe_numeric(df.get("Shop Stock Qty (LM)", _safe_default_series(0, df.index)))
    stock_cm = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, df.index)))

    mask = (
        (sold_m2 <= 0) & (sold_m1 <= 0) & (sold_mtd <= 0)
        & (sold_p90d <= 0)
        & (stock_lm > 0) & (stock_cm > 0)
    ).fillna(False)

    raw = df[mask].copy()
    if not raw.empty and "SKU" in raw.columns:
        raw["SKU"] = raw["SKU"].apply(_safe_sku)

    if progress_callback:
        progress_callback(0.35, "建立 Summary by Shop...")
    summary_by_shop = _build_p90d_zero_sales_by_shop_summary(raw)

    warnings = []
    if raw.empty:
        warnings.append("沒有符合「Past 90 Days 0 Sales（Sold Qty M-2/M-1/CM-MTD 與 Sold Qty (P90D) 皆 ≤ 0，且 Shop Stock Qty LM/CM 皆 > 0）」的資料。")

    if progress_callback:
        progress_callback(0.4, "完成資料轉換...")
    return raw, summary_by_shop, warnings


def generate_excel_p90d_zero_sales(
    raw: pd.DataFrame,
    summary_by_shop: pd.DataFrame,
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

    out_cols = _p90d_raw_columns(raw if raw is not None else pd.DataFrame())

    # ── Sheet 1: RAW ──
    ws_raw = wb.active
    ws_raw.title = "RAW"
    _write_raw_sheet(ws_raw, raw, progress_callback, progress_start=0.0, progress_end=0.55, columns=out_cols)

    # ── Sheet 2: Summary by Shop ──
    if progress_callback:
        progress_callback(0.58, "建立 Summary by Shop...")
    ws_summary = wb.create_sheet(title="Summary by Shop")
    summary_headers = [
        "Site Code",
        "Shop Name",
        "OM",
        "Past 90 Days 0 Sales SKU 數量",
        "Shop Stock Qty (LM) 總數",
        "Shop Stock Qty (CM) 總數",
    ]
    for col_idx, h in enumerate(summary_headers, 1):
        cell = ws_summary.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    if summary_by_shop is not None and not summary_by_shop.empty:
        total_summary = len(summary_by_shop)
        for row_idx, row_data in summary_by_shop.iterrows():
            excel_row = row_idx + 2
            is_total = row_data.get("Site Code") == "總計"
            for col_idx, h in enumerate(summary_headers, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws_summary.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if is_total:
                    cell.font = bold_font
                if isinstance(val, (int, float)) and col_idx >= 4:
                    cell.number_format = "#,##0"
            if progress_callback and (row_idx % 500 == 0 or row_idx == total_summary - 1):
                progress_callback(0.6 + 0.1 * ((row_idx + 1) / max(total_summary, 1)), f"寫入 Summary by Shop ({row_idx+1}/{total_summary})...")
    else:
        ws_summary.cell(row=2, column=1, value="無符合資料").font = Font(size=12)
    _auto_width(ws_summary)

    if progress_callback:
        progress_callback(0.9, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf


# ─── Stock Checking (調貨檢查) Functions ───────────────────────────────────────


_STOCK_CHECKING_REQUIRED = [
    "Shop Nature",
    "Sold Qty (P90D)",
    "Shop Stock Qty (CM)",
    "Site Code",
]


def _stock_checking_new_columns(raw: pd.DataFrame) -> list[str]:
    """Stock Checking 新增的計算欄位清單（存在才輸出，供測試與 Excel 使用）。"""
    cols = [
        "Tourist 加總 TurnDay",
        "Mix 加總 TurnDay",
        "Non Tourist 加總 TurnDay",
        "需要調貨",
        "調入需求量 (補到14日)",
        "實際調入量",
        "調入缺口",
        "可轉出貨量 (Mix)",
        "可轉出貨量 (Non Tourist)",
        "轉出量",
        "調貨來源 (Site Code)",
    ]
    avail = set(raw.columns)
    return [c for c in cols if _resolve_column_name(c, avail) is not None]


def _stock_checking_stats_empty() -> dict:
    return {
        "need_rows": 0,
        "need_skus": 0,
        "total_demand": 0,
        "total_actual": 0,
        "total_gap": 0,
        "nt_transfer": 0,
        "mix_transfer": 0,
        "buyer_need_skus": [],
    }


def _stock_checking_buyer_stats(raw: pd.DataFrame, need_series: pd.Series | None = None) -> list[tuple[str, int]]:
    """Buyer 級統計：各 Buyer 涉及多少個「需要調貨」SKU（已刪除重複）。

    - Buyer 取同 SKU 首個非空值（與 Summary 同口徑）；全空歸入「(未填 Buyer)」
    - 回傳 [(Buyer, 需調貨 SKU 數)]，按數量降序、同數按 Buyer 名排序
    """
    if raw is None or raw.empty or "SKU" not in raw.columns:
        return []
    if need_series is None:
        need_series = raw.get("需要調貨", _safe_default_series("", raw.index)).astype(str) == "是"
    need_skus = set(raw.loc[need_series, "SKU"].dropna().unique())
    if not need_skus:
        return []
    buyer = (
        raw["Buyer"].fillna("").astype(str).str.strip()
        if "Buyer" in raw.columns
        else _safe_default_series("", raw.index)
    )
    tmp = pd.DataFrame({"_sku": raw["SKU"], "_buyer": buyer})
    tmp = tmp[tmp["_sku"].isin(need_skus)]

    def _first_nonempty(s: pd.Series) -> str:
        s = s[s != ""]
        return s.iloc[0] if len(s) else "(未填 Buyer)"

    per_sku_buyer = tmp.groupby("_sku", sort=False)["_buyer"].agg(_first_nonempty)
    counts = per_sku_buyer.value_counts()
    return sorted(((str(b), int(c)) for b, c in counts.items()), key=lambda kv: (-kv[1], kv[0]))


# Stock Checking／混合 nature 維度的 Summary A–AC 契約欄序基準
# （由 _stock_checking_summary_columns("nature") 產生，測試對帳用）
_STOCK_CHECKING_SUMMARY_COLUMNS = [
    "Brand", "SKU", "SKU Desc", "Buyer", "Adj. SST", "Supply Source", "D001 Net Stock Qty (CM)",
    "Tourist 加總 - Shop Stock Qty (CM)", "Tourist 加總 - Shop Monthly Sold Qty (P90D)", "Tourist TurnDay",
    "Tourist 涉及店舖數",
    "Mix 加總 - Shop Stock Qty (CM)", "Mix 加總 - Shop Monthly Sold Qty (P90D)", "Mix TurnDay",
    "Mix 涉及店舖數",
    "Non Tourist 加總 - Shop Stock Qty (CM)", "Non Tourist 加總 - Shop Monthly Sold Qty (P90D)", "Non Tourist TurnDay",
    "Non Tourist 涉及店舖數",
    "Tourist 加總 - Shop Stock Qty (CM) %", "Tourist 加總 - Shop Monthly Sold Qty (P90D) %",
    "Mix 加總 - Shop Stock Qty (CM) %", "Mix 加總 - Shop Monthly Sold Qty (P90D) %",
    "Non Tourist 加總 - Shop Stock Qty (CM) %", "Non Tourist 加總 - Shop Monthly Sold Qty (P90D) %",
    "強制調貨提示", "由Non Tourist 店舖預計調出數量", "由Mix 店舖預計調出數量",
    "存貨偏重提示",
]


def _stock_checking_summary_columns(dimension: str = "nature", bidirectional: bool = False) -> list[str]:
    """Summary 欄序（按分類維度切換分組欄位）：
    nature＝三 Nature（現行契約，含雙向時末尾「由Tourist 店舖預計調出數量」）；
    rp＝RF／ND 兩組；both＝六組（Tourist-RF … Non Tourist-ND）。
    每組加總區為加總庫存／加總月銷／TurnDay／涉及店舖數四欄（店舖數＝同 SKU 同組
    Site Code 不重複計數，空白 Site Code 不計）。"""
    dim = _hybrid_dimension(dimension)
    groups = _hybrid_dimension_groups(dim)
    cols = ["Brand", "SKU", "SKU Desc", "Buyer", "Adj. SST", "Supply Source", "D001 Net Stock Qty (CM)"]
    for gname in groups:
        cols += [
            f"{gname} 加總 - Shop Stock Qty (CM)",
            f"{gname} 加總 - Shop Monthly Sold Qty (P90D)",
            f"{gname} TurnDay",
            f"{gname} 涉及店舖數",
        ]
    for gname in groups:
        cols += [
            f"{gname} 加總 - Shop Stock Qty (CM) %",
            f"{gname} 加總 - Shop Monthly Sold Qty (P90D) %",
        ]
    cols += ["強制調貨提示"]
    if dim == "nature":
        cols += ["由Non Tourist 店舖預計調出數量", "由Mix 店舖預計調出數量"]
    else:
        cols += [f"由{gname} 店舖預計調出數量" for gname in groups]
    cols += ["存貨偏重提示"]
    if dim == "nature" and bidirectional:
        cols += ["由Tourist 店舖預計調出數量"]
    return cols


def _build_stock_checking_summary(
    raw: pd.DataFrame, bidirectional: bool = False, dimension: str = "nature"
) -> pd.DataFrame:
    """SKU 級 Summary（參考 By SKUxShop_Stock Checking_Output 格式）：

    - 每個 SKU 一列：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001 Net Stock Qty (CM)
      （Buyer、Adj. SST 取同 SKU 首個非空值作代表）
    - 按分類維度分組（dimension：nature 三組／rp 兩組 RF／ND／both 六組）加總
      Shop Stock Qty (CM) 與 Shop Monthly Sold Qty (P90D)，
      TurnDay = 加總庫存 ÷ 加總月銷 × 30.5（月銷為 0 時留空）；除佔比外一律取整顯示
    - 各組涉及店舖數：同 SKU 同組的 Site Code 不重複計數（空白 Site Code 不計），
      與加總同母體，置於該組 TurnDay 之後
    - 各組佔比 (%)：該組加總 ÷ 全部分組加總合計（0-1 小數，Excel 以 0.00% 顯示；
      合計為 0 時留空；nature 維度僅計 Tourist / Mix / Non Tourist 精確比對）
    - 強制調貨提示：該 SKU 有任何店舖列被標示「需要調貨」時顯示「需要調貨」
    - 存貨偏重提示：某分組庫存佔比 ≥ 40% 且庫存佔比 − 銷量佔比 ≥ 20 個百分點 →
      「存貨偏重 {分組}（庫存 x% / 銷量 y%）」，多組以；分隔（差距大者在前）
    - 由 X 店舖預計調出數量 = 該 SKU 對應分組來源店的「轉出量」加總；
      nature 維度：Non Tourist / Mix 固定、bidirectional=True 時另加「由Tourist 店舖預計調出數量」
      （置於欄序最後，A–AC 契約欄位不動）；rp／both 維度：每個分組各一欄
    """
    if raw is None or raw.empty:
        return pd.DataFrame()
    dim = _hybrid_dimension(dimension)
    groups = _hybrid_dimension_groups(dim)
    df = raw.copy()
    df["_sku"] = df["SKU"]
    df = df[df["_sku"].notna()]
    if df.empty:
        return pd.DataFrame()
    df["_nature"] = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    if dim == "nature":
        df["_grp"] = df["_nature"]
    else:
        rp = (
            df["RP Type"].fillna("").astype(str).str.strip().str.upper()
            if "RP Type" in df.columns
            else _safe_default_series("RF", df.index)
        )
        df["_grp"] = pd.Series(np.where(rp == "RF", "RF", "ND"), index=df.index)
        if dim == "both":
            df["_grp"] = df["_nature"] + "-" + df["_grp"]
    df["_brand"] = df.get("Brand", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["_stock"] = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, df.index)))
    df["_msold"] = _safe_numeric(df.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, df.index)))
    df["_transfer"] = _safe_numeric(df.get("轉出量", _safe_default_series(0, df.index)))
    df["_need"] = df.get("需要調貨", _safe_default_series("", df.index)).astype(str) == "是"
    df["_site"] = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()

    def _first_text(series):
        s = series.dropna().astype(str).str.strip()
        s = s[s != ""]
        return s.iloc[0] if len(s) else ""

    rows = []
    for sku, g in df.groupby("_sku", sort=False):
        row = {
            "SKU": sku,
            "Brand": _first_text(g["_brand"]),
        }
        if "SKU Desc" in g.columns:
            row["SKU Desc"] = _first_text(g["SKU Desc"])
        # SKU Desc 之後：Buyer、Adj. SST（同 SKU 取首個非空值作代表）
        if "Buyer" in g.columns:
            row["Buyer"] = _first_text(g["Buyer"].fillna("").astype(str))
        if "Adj. SST" in g.columns:
            row["Adj. SST"] = _first_text(g["Adj. SST"].fillna("").astype(str))
        if "Supply Source" in g.columns:
            row["Supply Source"] = _first_text(g["Supply Source"])
        d001_source = (
            "D001 Effective Stock Qty (CM)"
            if "D001 Effective Stock Qty (CM)" in g.columns
            else "D001 Net Stock Qty (CM)"
        )
        if d001_source in g.columns:
            d001 = g[d001_source].dropna()
            row["D001 Net Stock Qty (CM)"] = d001.iloc[0] if len(d001) else None
        _sums: dict[str, tuple[float, float]] = {}
        for grp_name in groups:
            gn = g[g["_grp"] == grp_name]
            stock_sum = float(gn["_stock"].sum())
            msold_sum = float(gn["_msold"].sum())
            _sums[grp_name] = (stock_sum, msold_sum)
            # 整數顯示（% 除外）：庫存本為整數；月銷四捨五入到整數；TurnDay 四捨五入到整數日
            row[f"{grp_name} 加總 - Shop Stock Qty (CM)"] = int(round(stock_sum))
            row[f"{grp_name} 加總 - Shop Monthly Sold Qty (P90D)"] = int(round(msold_sum))
            row[f"{grp_name} TurnDay"] = int(round(stock_sum / msold_sum * 30.5)) if msold_sum > 0 else None
            row[f"{grp_name} 涉及店舖數"] = int(gn.loc[gn["_site"] != "", "_site"].nunique())
            # nature 維度維持現行契約（Tourist 欄僅雙向時存在）；rp／both 每組各一欄
            if dim != "nature" or grp_name != "Tourist" or bidirectional:
                row[f"由{grp_name} 店舖預計調出數量"] = int(gn["_transfer"].sum())
        # 各組佔比（分母 = 全部分組加總合計；合計為 0 留空）
        _total_stock = sum(v[0] for v in _sums.values())
        _total_msold = sum(v[1] for v in _sums.values())
        for grp_name in groups:
            stock_sum, msold_sum = _sums[grp_name]
            row[f"{grp_name} 加總 - Shop Stock Qty (CM) %"] = (
                round(stock_sum / _total_stock, 4) if _total_stock > 0 else None
            )
            row[f"{grp_name} 加總 - Shop Monthly Sold Qty (P90D) %"] = (
                round(msold_sum / _total_msold, 4) if _total_msold > 0 else None
            )
        row["強制調貨提示"] = "需要調貨" if bool(g["_need"].any()) else ""
        # 存貨偏重提示：某分組庫存佔比 ≥ 40% 且（庫存佔比 − 銷量佔比）≥ 20 個百分點
        _hints: list[tuple[float, str]] = []
        if _total_stock > 0:
            for grp_name in groups:
                stock_sum, msold_sum = _sums[grp_name]
                s_share = stock_sum / _total_stock
                v_share = (msold_sum / _total_msold) if _total_msold > 0 else 0.0
                if s_share >= 0.40 and (s_share - v_share) >= 0.20:
                    _hints.append((s_share - v_share,
                                   f"存貨偏重 {grp_name}（庫存 {s_share:.1%} / 銷量 {v_share:.1%}）"))
        row["存貨偏重提示"] = "；".join(h for _, h in sorted(_hints, key=lambda t: -t[0]))
        rows.append(row)

    summary = pd.DataFrame(rows)
    summary = summary.sort_values(["Brand", "SKU"], kind="stable").reset_index(drop=True)
    ordered = _stock_checking_summary_columns(dim, bidirectional=bidirectional)
    ordered = [c for c in ordered if c in summary.columns]
    trailing = [c for c in summary.columns if c not in ordered]
    return summary[ordered + trailing]


def transform_stock_checking(
    df_input: pd.DataFrame,
    turnday_threshold: int = 14,
    d001_threshold: int = 100,
    progress_callback=None,
):
    """Stock Checking（調貨檢查）核心轉換。

    D001 倉庫少貨（低於門檻）且遊客區（Tourist）店舖缺貨時的檢查（SKU 級判定）：

    - 觸發條件（同時符合才標示）：該 SKU 的「Tourist 加總 TurnDay」
      （= 各 Tourist 店 Shop Stock Qty (CM) 加總 ÷ Shop Monthly Sold Qty (P90D) 加總 × 30.5，與 Summary 同口徑）
      < Turnday 門檻，且 D001 < 門檻，且 Mix / Non Tourist 至少一個 Nature 的加總 TurnDay ≥ 門檻（有貨可調）。
      加總 TurnDay ≥ 門檻的 SKU 一律無須調貨；Mix 與 Non Tourist 加總 TurnDay 皆低於門檻時視為無貨可調，不標示。
    - 需求端：已標示 SKU 的 Tourist 店列，補到 14 日周轉的需求 > 0
      → 標示「需要調貨」，調入需求量 = ceil(P90D / 90 × 14) − Shop Stock Qty (CM)；RAW 只為已標示 SKU 計算調貨數量。
    - 供應端：Mix / Non Tourist 店列的可轉出貨量（14 日超額標準，與 Stock Impact 模式一致）：
      可轉出 = max(0, Shop Stock Qty (CM) − P90D / 90 × 14)；「無銷售但有存貨」全數可釋放。
      該 Nature 的 SKU 級加總 TurnDay < 門檻時，該 Nature 視為無貨可調（可轉出量歸零）。RP Type 為 RF 與 ND 皆可作為來源。
    - 分配：按 SKU 彙總需求，Non Tourist 店優先，不足再由 Mix 店補足；需求列按需求量大→小依序分配；
      來源店與需求店同為一店（Site Code 相同）時一律排除，防止重複列資料異常造成自我調貨。
    """
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    validate_input(df_input)
    avail = set(df_input.columns)
    missing = [name for name in _STOCK_CHECKING_REQUIRED if _resolve_column_name(name, avail) is None]
    if missing:
        raise ValueError(f"缺少必需欄位: {', '.join(missing)}")

    if progress_callback:
        progress_callback(0.1, "清理資料中...")
    df = _clean_dataframe(df_input, shop_natures=None, progress_callback=progress_callback)

    if progress_callback:
        progress_callback(0.2, "解像 SKU Desc / Shop Name...")
    df = _compute_raw_extra_columns(df)

    idx = df.index
    warnings: list[str] = []

    # ── 欄位正規化 ──────────────────────────────────────────────────────────
    df["Shop Nature"] = df.get("Shop Nature", _safe_default_series("", idx)).fillna("").astype(str).str.strip()
    nature = df["Shop Nature"]
    if "SKU" in df.columns:
        df["SKU"] = df["SKU"].apply(_safe_sku)
    warnings.extend(_transfer_input_consistency_warnings(df))

    # D001：優先 D001 Net Stock Qty (CM)，退用 D001 Stock Qty (CM)，兩者皆缺視為 0
    d001_col = None
    for cand in ("D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)"):
        if _resolve_column_name(cand, avail):
            d001_col = cand
            break
    if d001_col is None:
        warnings.append("來源檔案缺少「D001 Net Stock Qty (CM)」與「D001 Stock Qty (CM)」，D001 一律視為 0（D001 低於門檻的判定將全數成立）。")
    d001_cm = _safe_numeric(df.get(d001_col, _safe_default_series(0, idx)) if d001_col else _safe_default_series(0, idx))

    sold_p90d = _safe_numeric(df.get("Sold Qty (P90D)", _safe_default_series(0, idx)))
    cm_stock = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, idx)))

    # ── 每列計算欄位 ────────────────────────────────────────────────────────
    is_mix = (nature == "Mix").to_numpy()
    is_nt = (nature == "Non Tourist").to_numpy()
    is_tourist = (nature == "Tourist").to_numpy()
    cm_stock_val = cm_stock.to_numpy(dtype=float)
    sold_val = sold_p90d.to_numpy(dtype=float)
    no_sale_has_stock = (sold_val <= 0) & (cm_stock_val > 0)
    ideal_14d = sold_val / 90.0 * 14

    # ── SKU 級各 Nature 加總 TurnDay（與 Summary 同口徑：加總庫存 ÷ 加總月銷 × 30.5）──
    if _resolve_column_name("Shop Monthly Sold Qty (P90D)", avail) is not None:
        msold_monthly = _safe_numeric(df["Shop Monthly Sold Qty (P90D)"])
    else:
        # 缺月銷欄位時以 Sold Qty (P90D) 換算（P90D 為 90 日總量 → 月銷 = /90 × 30.5）
        msold_monthly = sold_p90d / 90.0 * 30.5
        warnings.append("來源檔案缺少「Shop Monthly Sold Qty (P90D)」，加總 TurnDay 改以 Sold Qty (P90D) 換算。")
    msold_val = msold_monthly.to_numpy(dtype=float)
    sku_key = df["SKU"] if "SKU" in df.columns else _safe_default_series(None, idx)

    def _sku_nature_turnday(nature_mask: np.ndarray) -> np.ndarray:
        s_stock = pd.Series(np.where(nature_mask, cm_stock_val, 0.0), index=df.index)
        s_msold = pd.Series(np.where(nature_mask, msold_val, 0.0), index=df.index)
        agg_stock = s_stock.groupby(sku_key).transform("sum")
        agg_msold = s_msold.groupby(sku_key).transform("sum")
        with np.errstate(divide="ignore", invalid="ignore"):
            return np.where(agg_msold.to_numpy() > 0,
                            agg_stock.to_numpy() / agg_msold.to_numpy() * 30.5,
                            np.inf)

    sku_tourist_turnday = _sku_nature_turnday(is_tourist)
    sku_mix_turnday = _sku_nature_turnday(is_mix)
    sku_nt_turnday = _sku_nature_turnday(is_nt)
    for label, td in (("Tourist", sku_tourist_turnday), ("Mix", sku_mix_turnday), ("Non Tourist", sku_nt_turnday)):
        df[f"{label} 加總 TurnDay"] = np.round(np.where(np.isfinite(td), td, np.nan), 2)

    # 供應門檻：Nature 加總 TurnDay < 門檻 → 該 Nature 理論上無貨可調，可轉出量歸零
    mix_pool_ok = sku_mix_turnday >= turnday_threshold   # inf（無銷售）→ 有存貨可釋放
    nt_pool_ok = sku_nt_turnday >= turnday_threshold
    mix_excess = np.where(mix_pool_ok, np.where(is_mix, np.where(no_sale_has_stock, cm_stock_val, np.maximum(0, cm_stock_val - ideal_14d)), 0), 0)
    nt_excess = np.where(nt_pool_ok, np.where(is_nt, np.where(no_sale_has_stock, cm_stock_val, np.maximum(0, cm_stock_val - ideal_14d)), 0), 0)
    df["可轉出貨量 (Mix)"] = mix_excess.astype(int)
    df["可轉出貨量 (Non Tourist)"] = nt_excess.astype(int)

    demand = np.maximum(0, np.ceil(ideal_14d) - cm_stock_val)
    d001_np = d001_cm.to_numpy(dtype=float)
    ttd_np = pd.Series(sku_tourist_turnday, index=df.index).to_numpy(dtype=float)
    need_mask = (
        is_tourist
        & (ttd_np < turnday_threshold)   # Tourist 加總 TurnDay ≥ 門檻（含月銷 0 → inf、NaN）→ 無須調貨
        & (d001_np < d001_threshold)
        & (demand > 0)
        & (mix_pool_ok | nt_pool_ok)     # Mix 與 Non Tourist 加總 TurnDay 皆低於門檻 → 無貨可調，不標示
    )
    need_series = pd.Series(need_mask, index=df.index)

    df["需要調貨"] = np.where(need_mask, "是", "")
    df["調入需求量 (補到14日)"] = np.where(need_mask, demand, 0).astype(int)
    df["實際調入量"] = 0
    df["調入缺口"] = 0
    df["轉出量"] = 0
    df["調貨來源 (Site Code)"] = ""

    # ── 跨店調貨分配（按 SKU：Non Tourist 優先 → Mix 補足）───────────────────
    if need_mask.any():
        site_arr = df.get("Site Code", _safe_default_series("", idx)).fillna("").astype(str).str.strip().to_numpy(dtype=object)
        sku_arr = df["SKU"].to_numpy(dtype=object)
        nt_avail = df["可轉出貨量 (Non Tourist)"].to_numpy(dtype=float).copy()
        mix_avail = df["可轉出貨量 (Mix)"].to_numpy(dtype=float).copy()
        demand_np = demand.copy()
        donate_np = np.zeros(len(df), dtype=int)
        actual_np = np.zeros(len(df), dtype=int)
        gap_np = np.zeros(len(df), dtype=int)
        desc_np = np.full(len(df), "", dtype=object)

        need_sku_set = set(sku_arr[need_mask])
        for sku in need_sku_set:
            sku_mask = sku_arr == sku
            need_rows = np.nonzero(sku_mask & need_mask)[0]
            nt_src = [int(p) for p in np.nonzero(sku_mask & (nt_avail > 0))[0]]
            mix_src = [int(p) for p in np.nonzero(sku_mask & (mix_avail > 0))[0]]
            order = np.argsort(-demand_np[need_rows], kind="stable")
            remain_nt = {int(p): float(nt_avail[p]) for p in nt_src}
            remain_mix = {int(p): float(mix_avail[p]) for p in mix_src}
            for n in order:
                row_pos = int(need_rows[n])
                rem = float(demand_np[row_pos])
                parts: list[str] = []
                donated_actual = 0
                for pool, _label in ((remain_nt, "Non Tourist"), (remain_mix, "Mix")):
                    for spos in list(pool.keys()):
                        if rem <= 0:
                            break
                        if site_arr[spos] == site_arr[row_pos]:
                            continue  # 自我調貨守衛：來源店＝需求店時排除（重複列資料異常防護）
                        avail = pool[spos]
                        take = min(rem, avail)
                        if take > 0:
                            pool[spos] -= take
                            rem -= take
                            donate_np[spos] += int(take)
                            donated_actual += int(take)
                            parts.append(f"{site_arr[spos]}:{int(take)}")
                    if rem <= 0:
                        break
                actual_np[row_pos] = donated_actual
                gap_np[row_pos] = int(rem)
                desc_np[row_pos] = "; ".join(parts)

        df["實際調入量"] = actual_np
        df["調入缺口"] = gap_np
        df["轉出量"] = donate_np
        df["調貨來源 (Site Code)"] = desc_np

    # ── 統計 ────────────────────────────────────────────────────────────────
    stats = {
        "turnday_threshold": int(turnday_threshold),
        "d001_threshold": int(d001_threshold),
        "need_rows": int(need_series.sum()),
        "need_skus": int(df.loc[need_series, "SKU"].nunique()) if need_series.any() else 0,
        "total_demand": int(df.loc[need_series, "調入需求量 (補到14日)"].sum()) if need_series.any() else 0,
        "total_actual": int(df["實際調入量"].sum()),
        "total_gap": int(df["調入缺口"].sum()),
        "nt_transfer": int(df.loc[df["Shop Nature"] == "Non Tourist", "轉出量"].sum()),
        "mix_transfer": int(df.loc[df["Shop Nature"] == "Mix", "轉出量"].sum()),
        "buyer_need_skus": _stock_checking_buyer_stats(df, need_series),
    }

    if not need_mask.any():
        warnings.append("沒有符合調貨條件的 Tourist SKU（Turnday < 門檻 且 D001 < 門檻 且補到 14 日周轉的需求 > 0）。")

    # ── SKU 級 Summary ──────────────────────────────────────────────────────
    if progress_callback:
        progress_callback(0.35, "建立 SKU Summary...")
    summary = _build_stock_checking_summary(df)

    if progress_callback:
        progress_callback(0.4, "完成資料轉換...")
    return df, summary, stats, warnings


def _write_stock_checking_summary_sheet(ws, summary, progress_callback=None, progress_start: float = 0.05, progress_end: float = 0.30, bidirectional: bool = False, highlight_skew: bool = False, dimension: str = "nature"):
    """Summary 工作表：SKU 級彙總（雙列表頭模仿 By SKUxShop_Stock Checking_Output 樞紐佈局）。

    欄位配置按分類維度分組數 G（nature 3／rp 2／both 6）動態排列：
    - A-G：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001
    - H 起：各分組加總庫存 / 加總月銷 / TurnDay / 涉及店舖數 ×G（整數顯示，TurnDay 活公式四捨五入到整數日）
    - 其後：各分組佔比 (%) ×G（活公式，0.00% 顯示；合計 0 留空）
    - 其後：強制調貨提示 / 各分組預計調出數量（nature 維度＝NT、Mix 兩欄固定順序，
      雙向時「由Tourist」欄附於存貨偏重提示之後，A–AC 契約欄序不動；rp／both 維度＝每組各一欄）
    - 最後：存貨偏重提示

    nature 維度（預設）之欄位號與 v1.22 版一致再各組加一欄涉及店舖數
    （加總區每組 4 欄；佔比 T–Y、提示 Z、調出 AA–AB、偏重 AC、雙向 AD）。
    highlight_skew=True 時，存貨偏重提示非空的列以醒目橘底（FFC000）加強顯示
    提示欄與偏重分組的庫存佔比欄。
    """
    dim = _hybrid_dimension(dimension)
    groups = _hybrid_dimension_groups(dim)
    G = len(groups)
    base_col = 8
    pct_start = base_col + 4 * G
    tip_col = pct_start + 2 * G
    if dim == "nature":
        out_keys = ["由Non Tourist 店舖預計調出數量", "由Mix 店舖預計調出數量"]
        tail_keys = ["由Tourist 店舖預計調出數量"] if bidirectional else []
        block_label, pct_label = "Shop Nature", "各 Nature 佔比 (%)"
    else:
        out_keys = [f"由{g} 店舖預計調出數量" for g in groups]
        tail_keys = []
        block_label = _HYBRID_DIMENSION_LABELS[dim]
        pct_label = f"各 {block_label} 佔比 (%)"
    hint_col = tip_col + 1 + len(out_keys)

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    nature_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    pct_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
    band_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    skew_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    skew_font = Font(bold=True)
    header_font = Font(bold=True, size=11)
    nature_font = Font(bold=True, size=11, color="FFFFFF")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    def _style_merge(row, c1, c2, value, font, fill):
        ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
        for cc in range(c1, c2 + 1):
            cell = ws.cell(row=row, column=cc)
            cell.border = thin_border
            cell.alignment = center
            cell.fill = fill
            cell.font = font
        ws.cell(row=row, column=c1, value=value)

    # Row 3: 分組群組標籤 + 佔比群組
    _style_merge(3, base_col, pct_start - 1, block_label, header_font, PatternFill(fill_type=None))
    _style_merge(3, pct_start, tip_col - 1, pct_label, header_font, PatternFill(fill_type=None))

    # Row 4: 各分組標籤 — 加總區（橫向合併 4 欄）+ 佔比區（橫向合併 2 欄）
    for i, gname in enumerate(groups):
        _style_merge(4, base_col + i * 4, base_col + i * 4 + 3, gname, nature_font, nature_fill)
    for i, gname in enumerate(groups):
        _style_merge(4, pct_start + i * 2, pct_start + i * 2 + 1, gname, nature_font, pct_fill)

    # Row 5: 欄位標題
    _headers = [
        "Brand", "SKU", "SKU Desc", "Buyer", "Adj. SST", "Supply Source", "D001 Net Stock Qty (CM)",
    ]
    for _ in groups:
        _headers += ["加總 - Shop Stock Qty (CM)", "加總 - Shop Monthly Sold Qty (P90D)", "TurnDay", "涉及店舖數"]
    for _ in groups:
        _headers += ["加總 - Shop Stock Qty (CM) %", "加總 - Shop Monthly Sold Qty (P90D) %"]
    _headers += ["強制調貨提示"] + out_keys + ["存貨偏重提示"] + tail_keys
    for col_idx, h in enumerate(_headers, 1):
        cell = ws.cell(row=5, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 28
    ws.row_dimensions[5].height = 45

    if summary is None or summary.empty:
        ws.cell(row=6, column=1, value="無符合資料").font = Font(size=12)
        _auto_width(ws)
        return

    _text_keys = ["Brand", "SKU Desc", "Buyer", "Adj. SST"]
    _center_text_keys = ["SKU", "Supply Source"]

    total = len(summary)
    n_cols = len(_headers)
    _stock_letters = [get_column_letter(base_col + i * 4) for i in range(G)]
    _sold_letters = [get_column_letter(base_col + i * 4 + 1) for i in range(G)]
    for r_idx, (_, row) in enumerate(summary.iterrows()):
        excel_row = r_idx + 6
        is_band = (r_idx % 2 == 1)
        # A-G：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001
        for col_idx, key in enumerate(
            ["Brand", "SKU", "SKU Desc", "Buyer", "Adj. SST", "Supply Source", "D001 Net Stock Qty (CM)"], 1
        ):
            val = row.get(key)
            cell = ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)
            cell.border = thin_border
            if key in _text_keys:
                cell.alignment = left_wrap
            else:
                cell.alignment = center
            if col_idx == 7:
                cell.number_format = "#,##0"
            if is_band:
                cell.fill = band_fill
        # 各分組加總、TurnDay 與涉及店舖數（整數；TurnDay 活公式四捨五入）
        for i, gname in enumerate(groups):
            bcol = base_col + i * 4
            stock = row.get(f"{gname} 加總 - Shop Stock Qty (CM)", 0)
            msold = row.get(f"{gname} 加總 - Shop Monthly Sold Qty (P90D)", 0)
            nshop = row.get(f"{gname} 涉及店舖數", 0)
            c_stock = ws.cell(row=excel_row, column=bcol,
                              value=int(stock) if pd.notna(stock) else None)
            c_sold = ws.cell(row=excel_row, column=bcol + 1,
                             value=int(msold) if pd.notna(msold) else None)
            c_nshop = ws.cell(row=excel_row, column=bcol + 3,
                              value=int(nshop) if pd.notna(nshop) else None)
            for cc in (c_stock, c_sold, c_nshop):
                cc.number_format = "#,##0"
                cc.border = thin_border
                cc.alignment = center
                if is_band:
                    cc.fill = band_fill
            s_letter = get_column_letter(bcol)
            m_letter = get_column_letter(bcol + 1)
            if msold is not None and pd.notna(msold) and float(msold) > 0:
                c_td = ws.cell(row=excel_row, column=bcol + 2,
                               value=f"=ROUND({s_letter}{excel_row}/{m_letter}{excel_row}*30.5,0)")
            else:
                c_td = ws.cell(row=excel_row, column=bcol + 2, value=None)
            c_td.number_format = "#,##0"
            c_td.border = thin_border
            c_td.alignment = center
            if is_band:
                c_td.fill = band_fill
        # 各分組佔比 (%) 活公式（分母 = 全部分組合計；合計 0 留空）
        _stock_total = "(" + "+".join(f"{L}{excel_row}" for L in _stock_letters) + ")"
        _sold_total = "(" + "+".join(f"{L}{excel_row}" for L in _sold_letters) + ")"
        for i, gname in enumerate(groups):
            for j, (num_letter, denom) in enumerate(((_stock_letters[i], _stock_total), (_sold_letters[i], _sold_total))):
                denom_bare = denom.strip("()")
                c_pct = ws.cell(row=excel_row, column=pct_start + i * 2 + j,
                                value=f'=IF({denom_bare}=0,"",{num_letter}{excel_row}/{denom})')
                c_pct.number_format = "0.00%"
                c_pct.border = thin_border
                c_pct.alignment = center
                if is_band:
                    c_pct.fill = band_fill
        # 強制調貨提示
        tip = row.get("強制調貨提示", "")
        cell = ws.cell(row=excel_row, column=tip_col, value=tip if tip else None)
        cell.border = thin_border
        cell.alignment = center
        if is_band:
            cell.fill = band_fill
        # 各分組預計調出數量
        for j, key in enumerate(out_keys):
            val = row.get(key, 0)
            cell = ws.cell(row=excel_row, column=tip_col + 1 + j,
                           value=int(val) if pd.notna(val) else None)
            cell.number_format = "#,##0"
            cell.border = thin_border
            cell.alignment = center
            if is_band:
                cell.fill = band_fill
        # 存貨偏重提示
        hint = row.get("存貨偏重提示", "")
        hint = "" if pd.isna(hint) else str(hint)
        cell = ws.cell(row=excel_row, column=hint_col, value=hint if hint else None)
        cell.border = thin_border
        cell.alignment = left_wrap
        if is_band:
            cell.fill = band_fill
        # 偏重數值以醒目橘底加強顯示（提示欄 + 偏重分組的庫存佔比欄）
        if hint and highlight_skew:
            cell.fill = skew_fill
            cell.font = skew_font
            for i, gname in enumerate(groups):
                if f"存貨偏重 {gname}" in hint:
                    _pc = ws.cell(row=excel_row, column=pct_start + i * 2)
                    _pc.fill = skew_fill
                    _pc.font = skew_font
        # nature 維度雙向：由Tourist 店舖預計調出數量（附於最後，A–AC 契約欄序不動）
        for j, key in enumerate(tail_keys):
            val = row.get(key, 0)
            cell = ws.cell(row=excel_row, column=hint_col + 1 + j,
                           value=int(val) if pd.notna(val) else None)
            cell.number_format = "#,##0"
            cell.border = thin_border
            cell.alignment = center
            if is_band:
                cell.fill = band_fill
        if progress_callback and (r_idx % 200 == 0 or r_idx == total - 1):
            p = progress_start + (progress_end - progress_start) * ((r_idx + 1) / max(total, 1))
            progress_callback(p, f"寫入 Summary ({r_idx+1}/{total})...")

    last_row = 5 + total
    ws.auto_filter.ref = f"A5:{get_column_letter(n_cols)}{last_row}"
    ws.freeze_panes = "A6"
    _auto_width(ws)


def _stock_checking_dashboard_metrics(raw: pd.DataFrame, stats: dict | None = None):
    """Stock Checking Dashboard 指標（與 Streamlit Transfer Analysis Dashboard 同口徑）。

    回傳 (row1_metrics, row2_metrics)，供 _write_metric_cards 使用。
    stats 缺失時（舊呼叫相容）直接由 raw 匯總。
    """
    total_rows = len(raw) if raw is not None else 0
    if stats is None:
        stats = {}
    need_col = raw.get("需要調貨", pd.Series(dtype=object)) if raw is not None else pd.Series(dtype=object)
    need_mask = (need_col.astype(str) == "是")
    need_rows = int(stats.get("need_rows", int(need_mask.sum())))
    if "need_skus" in stats:
        need_skus = int(stats["need_skus"])
    elif raw is not None and "SKU" in raw.columns and need_mask.any():
        need_skus = int(raw.loc[need_mask, "SKU"].nunique())
    else:
        need_skus = 0
    demand_col = "調入需求量 (補到14日)"
    actual_col = "實際調入量"
    gap_col = "調入缺口"
    out_col = "轉出量"
    total_demand = int(stats.get("total_demand", int(_safe_numeric(raw.get(demand_col, _safe_default_series(0, raw.index))).sum()) if raw is not None else 0))
    total_actual = int(stats.get("total_actual", int(_safe_numeric(raw.get(actual_col, _safe_default_series(0, raw.index))).sum()) if raw is not None else 0))
    total_gap = int(stats.get("total_gap", int(_safe_numeric(raw.get(gap_col, _safe_default_series(0, raw.index))).sum()) if raw is not None else 0))
    if "nt_transfer" in stats and "mix_transfer" in stats:
        nt_transfer = int(stats["nt_transfer"])
        mix_transfer = int(stats["mix_transfer"])
    elif raw is not None and "Shop Nature" in raw.columns and out_col in raw.columns:
        _out = _safe_numeric(raw.get(out_col, _safe_default_series(0, raw.index)))
        nt_transfer = int(_out[raw["Shop Nature"] == "Non Tourist"].sum())
        mix_transfer = int(_out[raw["Shop Nature"] == "Mix"].sum())
    else:
        nt_transfer = int(stats.get("nt_transfer", 0))
        mix_transfer = int(stats.get("mix_transfer", 0))

    row1_metrics = [
        ("需要調貨列數", need_rows, "Tourist 加總 TurnDay & D001 皆低於門檻的 SKU", "EF4444"),
        ("需調貨 SKU 數", need_skus, "符合條件的 SKU 總數（已刪除重複）", "F59E0B"),
        ("調入需求量", total_demand, "補到 14 日周轉的總需求（件）", "3B82F6"),
        ("實際調入量", total_actual, "NT 優先、Mix 補足後的總調入（件）", "10B981"),
    ]
    row2_metrics = [
        ("Non Tourist 轉出量", nt_transfer, "優先由 Non Tourist 店舖轉出（件）", "8B5CF6"),
        ("Mix 轉出量", mix_transfer, "NT 不足時由 Mix 店舖補足（件）", "06B6D4"),
        ("無法足額缺口", total_gap, "來源可轉出量不足的缺口（件）", "F97316"),
        ("總明細行數", total_rows, "RAW 完整明細總行數", "64748B"),
    ]
    return row1_metrics, row2_metrics


def _write_stock_checking_buyer_table(ws, buyer_stats, thin_border, start_row: int = 11):
    """Dashboard 指標卡下方的「Buyer 需調貨 SKU 統計」表（Buyer / 需調貨 SKU 數 / 佔比 + 總計）。"""
    if not buyer_stats:
        return
    center = Alignment(horizontal="center", vertical="center")
    header_fill = PatternFill(start_color="0EA5E9", end_color="0EA5E9", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    total_skus = sum(c for _, c in buyer_stats)

    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=3)
    tc = ws.cell(row=start_row, column=1, value="Buyer 需調貨 SKU 統計（按需調貨 SKU 數排序）")
    tc.font = Font(bold=True, size=12, color="1E40AF")
    tc.alignment = Alignment(horizontal="left", vertical="center")

    hr = start_row + 1
    for col_idx, h in enumerate(("Buyer", "需調貨 SKU 數", "佔比"), 1):
        cell = ws.cell(row=hr, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    r = hr + 1
    for buyer, cnt in buyer_stats:
        c1 = ws.cell(row=r, column=1, value=buyer)
        c1.border = thin_border
        c2 = ws.cell(row=r, column=2, value=cnt)
        c2.number_format = "#,##0"
        c2.alignment = center
        c2.border = thin_border
        c3 = ws.cell(row=r, column=3, value=cnt / total_skus if total_skus else None)
        c3.number_format = "0.0%"
        c3.alignment = center
        c3.border = thin_border
        r += 1
    for col_idx, v in ((1, "總計"), (2, total_skus), (3, 1.0 if total_skus else None)):
        cell = ws.cell(row=r, column=col_idx, value=v)
        cell.font = Font(bold=True, size=11)
        cell.fill = total_fill
        cell.border = thin_border
        cell.alignment = center
        if col_idx == 2:
            cell.number_format = "#,##0"
        elif col_idx == 3:
            cell.number_format = "0.0%"
    _auto_width(ws)


# Stock Checking RAW 顯示欄序分組（識別→品項屬性→庫存與銷售→SKU 級判定→調貨執行→其他）
_STOCK_CHECKING_RAW_DISPLAY_GROUPS = [
    ["Buyer", "Brand", "SKU", "SKU Desc", "Site Code", "Shop Name", "OM", "Shop Nature"],
    ["Adj. SST", "Adj. MCH1", "Vendor MC", "MCH2", "Item Status", "Item Class", "RP Type",
     "Shop Stock Status", "Supply Source", "Launch Date"],
    ["Shop Stock Qty (CM)", "Adj. Safety Stock", "Sold Qty (P90D)", "Shop Monthly Sold Qty (P90D)",
     "Turnday (P90D Qty)", "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
     "Net Sales (M-2)", "Net Sales (M-1)", "Net Sales (CM - MTD)", "Excess Stock Qty",
     "Shop On Order PO Qty (CM)", "Shop Pending Receive",
     "D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)", "D001 Effective Stock Qty (CM)"],
    ["Tourist 加總 TurnDay", "Mix 加總 TurnDay", "Non Tourist 加總 TurnDay",
     "需要調貨", "調入需求量 (補到14日)"],
    ["可轉出貨量 (Non Tourist)", "可轉出貨量 (Mix)", "實際調入量", "調入缺口", "轉出量",
     "調貨來源 (Site Code)"],
]


def _stock_checking_raw_display_columns(raw: pd.DataFrame) -> list[str]:
    """Stock Checking RAW 顯示欄序：不顯示「Item_Shop」與所有含「Value」的金額欄，
    其餘欄位（含 _compute_raw_extra_columns 衍生的 Shop Name / SKU Desc）按
    _STOCK_CHECKING_RAW_DISPLAY_GROUPS 分組排列，未歸組的原始欄位按原序置於最後。"""
    cols = [str(c) for c in raw.columns] + ["Shop Name", "SKU Desc"]
    seen: set[str] = set()
    kept: list[str] = []
    for c in cols:
        if c in seen or c == "Item_Shop" or "value" in c.lower():
            continue
        seen.add(c)
        kept.append(c)
    kept_set = set(kept)
    ordered: list[str] = []
    for group in _STOCK_CHECKING_RAW_DISPLAY_GROUPS:
        for c in group:
            if c in kept_set:
                ordered.append(c)
                kept_set.discard(c)
    ordered += [c for c in kept if c in kept_set]
    return ordered


def generate_excel_stock_checking(
    raw: pd.DataFrame,
    summary: pd.DataFrame | None = None,
    stats: dict | None = None,
    progress_callback=None,
) -> BytesIO:
    """Stock Checking 輸出：Summary（SKU 級彙總，偏重提示以醒目橘底加強顯示）
    + RAW（重排欄序提升可讀性，不顯示 Item_Shop 與 Value 金額欄）
    + 店舖問題 SKU（每間店舖需行動列；空表寫「無符合資料」）
    + 調貨分析摘要（Dashboard）。"""
    wb = Workbook()
    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    ws_summary = wb.active
    ws_summary.title = "Summary"
    _write_stock_checking_summary_sheet(ws_summary, summary, progress_callback, highlight_skew=True)

    if progress_callback:
        progress_callback(0.35, "建立 RAW...")
    ws_raw = wb.create_sheet(title="RAW")
    _write_raw_sheet(ws_raw, raw, progress_callback, progress_start=0.35, progress_end=0.80,
                     columns=_stock_checking_raw_display_columns(raw))

    if progress_callback:
        progress_callback(0.82, "建立店舖問題 SKU...")
    ws_shop_issue = wb.create_sheet(title="店舖問題 SKU")
    _write_shop_issue_sheet(
        ws_shop_issue,
        _build_stock_checking_shop_issues(
            raw,
            turnday_threshold=int((stats or {}).get("turnday_threshold", 14)),
            d001_threshold=int((stats or {}).get("d001_threshold", 100)),
        ),
        title="店舖問題 SKU — 每間店舖需行動列（問題＋建議方法）",
        progress_callback=progress_callback, progress_start=0.82, progress_end=0.86,
    )

    if progress_callback:
        progress_callback(0.88, "建立調貨分析摘要工作表...")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    desc_font = Font(size=9, color="6B7280", italic=True)
    ws_dashboard = wb.create_sheet(title="調貨分析摘要")
    row1_metrics, row2_metrics = _stock_checking_dashboard_metrics(raw, stats)
    _write_metric_cards(
        ws_dashboard, "調貨分析摘要 (Transfer Analysis Dashboard)",
        row1_metrics, row2_metrics, thin_border, desc_font, progress_callback,
    )
    buyer_stats = (stats or {}).get("buyer_need_skus")
    if buyer_stats is None:
        buyer_stats = _stock_checking_buyer_stats(raw)
    _write_stock_checking_buyer_table(ws_dashboard, buyer_stats, thin_border)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf

