# -*- coding: utf-8 -*-
"""Safety Stock Review模式：由 transformer.py 拆分而來。
共用 helper 由 transformer（最終為 modes.common／issues／framework）引用；
transformer 以墊片 re-export 本模組，舊有 import 路徑不變。
"""
import datetime
import gc
import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


from modes.common import (
    _ADJ_SAFETY_SOURCE_COL,
    _auto_width,
    _cached_shop_names,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _display_safety_header,
    _extract_shop_name,
    _iter_header_rows,
    _nature_rate_arr,
    _normalize_ref_date,
    _parse_launch_dates,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _cached_safe_sku_series,
    _write_guide_sheet,
    _write_metric_cards,
    _write_raw_placeholder,
    _write_raw_sheet,
    validate_input,
)


from modes.framework import (
    _HYBRID_DIMENSION_LABELS,
    _HYBRID_SAFETY_DAYS_DEFAULT,
    _HYBRID_SALES_OOS_DAYS_DEFAULT,
    _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT,
    _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT,
    _attach_sales_classification_columns,
    _hybrid_dimension,
    _hybrid_dimension_groups,
    _hybrid_group_key_series,
    _hybrid_om_groups_from_values,
    _hybrid_safety_days,
    _hybrid_safety_days_label,
    _hybrid_sales_params,
)


from modes.issues import (
    _build_combined_shop_issues,
    _write_shop_issue_sheet,
)
__all__ = [
    "COMBINED_RAW_COLS",
    "_CHANGE_CAT_DESC",
    "_ND_REC_DESC",
    "_ND_REC_DISPLAY",
    "_SSR_SAFETY_BASIS_RATE",
    "_SSR_SAFETY_BASIS_TURNDAY",
    "_SSR_SAFETY_BASIS_LABELS",
    "_SSR_SAFETY_DAYS_COLUMN",
    "_SSR_SAFETY_LINE_COLUMN",
    "_normalize_ssr_safety_basis",
    "_add_combined_recommendation_columns",
    "_annotate_nd_recommendation",
    "_build_combined_by_shop_summary",
    "_build_combined_classification_summary",
    "_build_combined_dimension_cross_tab",
    "_build_nd_rec_qty_nd",
    "_write_combined_guide_sheet",
    "_write_summary_sheet",
    "generate_excel_combined",
    "transform_combined",
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
    "Safety 基準天數",
    "Safety 周轉所需量",
    "D001 Net Stock Qty (CM)",
    "Shop Stock Qty (CM)",
    "Sold Qty (P90D)",
    "銷量帶",
    "銷量存貨分類",
    "Turnday (P90D Qty)",
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Supply Source",
]


# Safety 計算基準（可切換選項；預設 "rate" 維持現行 X/Y 行為）：
# - "rate"：月銷倍率（Overstock Cap X／OOS Floor Y；建議值 = max(Y, Z if Z ≤ X else X)）
# - "turnday"：周轉天數（Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 該列 Nature 基準天數)，
#   與 Safety Stock x NetStock Review 同口徑；建議值直接取周轉所需量含最低 3 件保底，
#   不設 Overstock 上限）。
_SSR_SAFETY_BASIS_RATE = "rate"
_SSR_SAFETY_BASIS_TURNDAY = "turnday"
_SSR_SAFETY_BASIS_LABELS = {
    _SSR_SAFETY_BASIS_RATE: "月銷倍率 (X/Y)",
    _SSR_SAFETY_BASIS_TURNDAY: "周轉天數 (按 Nature)",
}

# 周轉天數基準欄（與 Safety Stock x NetStock Review 同名同義，供對照與稽核）：
# - Safety 基準天數：該列實際採用的基準天數（按 Shop Nature；預設與 Hybrid 相同）
# - Safety 周轉所需量：ceil(Sold Qty (P90D) ÷ 90 × 基準天數) 並含最低 3 件保底
_SSR_SAFETY_DAYS_COLUMN = "Safety 基準天數"
_SSR_SAFETY_LINE_COLUMN = "Safety 周轉所需量"


def _normalize_ssr_safety_basis(safety_basis) -> str:
    """Safety 計算基準正規化：未知值退回 "rate"（維持現行行為，不拋錯）。"""
    raw = str(safety_basis or "").strip().lower()
    if raw in (_SSR_SAFETY_BASIS_TURNDAY, "turnover", "days", "周轉天數"):
        return _SSR_SAFETY_BASIS_TURNDAY
    return _SSR_SAFETY_BASIS_RATE






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
    df, rates_overstock, rates_oos, mtd_days, month_total_days, safety_mode="increase_only",
    ref_date=None, safety_basis="rate", safety_days=None,
):
    df = df.copy()
    safety_basis = _normalize_ssr_safety_basis(safety_basis)
    _ssr_days, _ = _hybrid_safety_days(safety_days)
    rates_overstock = rates_overstock or {"default": 0.50, "tourist": 0.30, "mix": 0.40}
    rates_oos = rates_oos or {"default": 0.50, "tourist": 0.50, "mix": 0.50}

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

    # y_preclip 保留未套「最低 3 件」保底前的 OOS 底線，供 _safety_floor3 判定用；
    # y_values 為現行輸出值（clip min=3），兩者計算準則完全相同。
    y_preclip = np.ceil(projected.values * oos_rate_arr)
    y_values = y_preclip.clip(min=3).astype(int)

    # Z = current Adj. Safety Stock
    z_values = adj_ss.values.astype(int)

    # ── 周轉天數基準（與 Safety Stock x NetStock Review 同口徑；兩種基準恆並列計算供對照）──
    # Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 該列 Nature 基準天數)，含最低 3 件保底
    # （沿用現行 OOS 底線保底；負數 P90D 按 0 計算；其他 Nature 值歸 Non Tourist 天數）。
    _days_arr = np.where(
        shop_nature.values == "tourist", _ssr_days.get("Tourist", 7),
        np.where(shop_nature.values == "mix", _ssr_days.get("Mix", 9),
                 _ssr_days.get("Non Tourist", 10)),
    ).astype(float)
    _sold_p90d = _safe_numeric(df.get("Sold Qty (P90D)", _safe_default_series(0, df.index))).values
    _sold_p90d = np.maximum(_sold_p90d, 0.0)
    _turn_preclip = np.ceil(_sold_p90d / 90.0 * _days_arr)
    _turn_values = _turn_preclip.clip(min=3).astype(int)

    df["Overstock Cap (X)"] = x_values
    df["OOS Floor (Y)"] = y_values
    df[_SSR_SAFETY_DAYS_COLUMN] = _days_arr.astype(int)
    df[_SSR_SAFETY_LINE_COLUMN] = _turn_values

    if safety_basis == _SSR_SAFETY_BASIS_TURNDAY:
        combined = _turn_values
        combined_pre = _turn_preclip
    else:
        # 月銷倍率基準：建議值 = max(Y, Z if Z ≤ X else X)，其中
        # X = Overstock Cap（庫存上限，RAW「Overstock Cap (X)」欄），
        # Y = OOS Floor（缺貨底線，RAW「OOS Floor (Y)」欄，含最低 3 件保底），
        # Z = 現有 Safety 設定（來源檔案的 Adj. Safety Stock）。
        combined = np.where(z_values <= x_values, np.maximum(y_values, z_values), np.maximum(y_values, x_values))
        # 同一公式、改用未保底的 y_preclip：代表「沒有最低 3 件保底時」的建議值
        combined_pre = np.where(z_values <= x_values, np.maximum(y_preclip, z_values), np.maximum(y_preclip, x_values))

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

    # 新品保護：Launch Date 距參考日期 90 天內（≈3 個月）視為 New SKU。新品未出單是上架
    # 時間問題而非滯銷證據，排除「P90D 0 Sales → Safety=0/ND」override，改走正常
    # 基準計算（月銷倍率 X/Y 或周轉天數）；Launch Date 缺欄位、缺值或無法解析時一律視為非新品（維持原行為）。
    # 參考日期預設為今天，建議傳入 MTD 截止日期（與報表資料基準一致）。
    _ref = _normalize_ref_date(ref_date)
    if "Launch Date" in df.columns:
        _launch_raw = df["Launch Date"].fillna("").astype(str).str.strip()
        _launch_dt = _parse_launch_dates(_launch_raw)
        is_new_sku = (
            _launch_dt.notna()
            & (_launch_dt >= pd.Timestamp(_ref) - pd.Timedelta(days=90))
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

    # 保底影響偵測（供店舖問題 SKU 建議方法加註）：最終建議值停在 3 件，且未套最低
    # 3 件保底時的本來建議值 < 3 → 建議值是因保底而停在 3（保障自動補貨）。
    # 「只可增加」模式的最終值 = max(現有設定, 未保底建議值)，故 recommend_pre 亦取兩者較大。
    recommend_pre = (
        combined_pre if safety_mode != "increase_only" else np.maximum(z_values, combined_pre)
    )
    df["_safety_floor3"] = (recommend_safety == 3) & (recommend_pre < 3)

    # Safety Qty Change 分類（增加 / 減少 / 不變）—「只可增加」模式下不會出現「減少」
    df["Safety Qty Change 分類"] = np.where(delta > 0, "增加", np.where(delta < 0, "減少", "不變"))

    # 建議 Safety Stock 的周轉日數，與來源「Turnday (P90D Qty)」同一準則（P90D 日均）：
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
    # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
    shop_name = _cached_shop_names(item_shop, site_code)

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



def _build_combined_dimension_cross_tab(raw, dimension="nature", om_groups=None):
    """Safety Qty Change 分類 × 分類維度交叉表（SUMMARY 區塊 ③）。

    Safety Stock Review 引入分類維度用：只改分組顯示，不影響任何 Safety 建議計算。
    分組鍵與混合模式同一準則（_hybrid_group_key_series）；om_groups 缺省時由 raw 即時推導；
    RAW 缺「銷量帶」欄時銷量兩個維度退回 Shop Nature 鍵（與 hybrid 同一守衛）。
    每格統計 列數 / SKU 數 / Safety Qty Change 合計；總計列的 SKU 數為全表不重複 SKU
    （同一 SKU 可跨分類，與區塊 ① 同一準則）；非標準 Shop Nature 值不佔分組欄、只計入合計。
    """
    dim = _hybrid_dimension(dimension)
    first_col = "Safety Qty Change 分類"
    if raw is None or raw.empty or "Safety Qty Change" not in raw.columns:
        groups = _hybrid_dimension_groups(dim, om_groups=om_groups)
        cols = [first_col]
        for g in groups:
            cols += [f"{g} 列數", f"{g} SKU 數", f"{g} Safety Qty Change 合計"]
        cols += ["合計 列數", "合計 SKU 數", "合計 Safety Qty Change"]
        return pd.DataFrame(columns=cols)
    if dim in ("om", "nature_om") and om_groups is None:
        om_groups = _hybrid_om_groups_from_values(
            raw["OM"] if "OM" in raw.columns else _safe_default_series("", raw.index)
        )
    groups = _hybrid_dimension_groups(dim, om_groups=om_groups)
    cols = [first_col]
    for g in groups:
        cols += [f"{g} 列數", f"{g} SKU 數", f"{g} Safety Qty Change 合計"]
    cols += ["合計 列數", "合計 SKU 數", "合計 Safety Qty Change"]
    chg = pd.to_numeric(raw["Safety Qty Change"], errors="coerce").fillna(0)
    grp = _hybrid_group_key_series(raw, dim, om_groups=om_groups)
    sku = raw["SKU"] if "SKU" in raw.columns else _safe_default_series("", raw.index)
    total_sku = int(sku.dropna().nunique())
    _total_keys = {
        "列數": "合計 列數",
        "SKU 數": "合計 SKU 數",
        "Safety Qty Change 合計": "合計 Safety Qty Change",
    }

    def _agg(sub_chg: pd.Series, sub_sku: pd.Series) -> dict:
        return {
            "列數": int(len(sub_chg)),
            "SKU 數": int(sub_sku.dropna().nunique()),
            "Safety Qty Change 合計": int(round(float(sub_chg.sum()))),
        }

    def _fill(row: dict, mask: pd.Series, total_row: bool = False) -> dict:
        for g in groups:
            gm = mask & (grp == g).to_numpy()
            for k, v in _agg(chg[gm], sku[gm]).items():
                row[f"{g} {k}"] = v
        agg = _agg(chg[mask], sku[mask])
        if total_row:
            agg["SKU 數"] = total_sku
        for k, v in agg.items():
            row[_total_keys[k]] = v
        return row

    cat_masks = [
        ("增加", chg > 0),
        ("減少", chg < 0),
        ("不變", chg == 0),
    ]
    rows = [_fill({first_col: label}, mask.to_numpy()) for label, mask in cat_masks]
    rows.append(_fill({first_col: "總計"}, np.ones(len(raw), dtype=bool), total_row=True))
    return pd.DataFrame(rows, columns=cols)



def transform_combined(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
    rates_overstock: dict | None = None,
    rates_oos: dict | None = None,
    mtd_days: int | None = None,
    month_total_days: int | None = None,
    safety_mode: str = "increase_only",
    ref_date=None,
    sales_oos_days=None,
    sales_overstock_days=None,
    sales_overstock_min_units=None,
    safety_basis: str = "rate",
    safety_days: dict | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str], list[str]]:
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    warnings = []
    safety_basis = _normalize_ssr_safety_basis(safety_basis)
    _ssr_days, _ssr_day_warnings = _hybrid_safety_days(safety_days)
    warnings.extend(_ssr_day_warnings)

    # Safety 來源欄相容檢查：顯示名一律視為來源名（Safety Stock (現時) = 來源檔案的 Adj. Safety Stock）
    _adj_avail = set(df_input.columns) | {_resolve_column_name(_ADJ_SAFETY_SOURCE_COL, set(df_input.columns))}
    _optional_cols = ["Adj. Safety Stock", "Sold Qty (CM - MTD)", "Shop Monthly Sold Qty (P90D)", "Shop Nature"]
    for _col in _optional_cols:
        _label = _display_safety_header(_col)
        if _resolve_column_name(_col, _adj_avail) is None:
            warnings.append(f"缺少可選欄位「{_label}」，建議值將可能不準確")

    if progress_callback:
        progress_callback(0.05, "驗證通過")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據（非空白或總計列）")

    df["SKU"] = df["SKU"].apply(_safe_sku)

    if "SKU Desc" not in df.columns:
        df = _compute_raw_extra_columns(df)

    raw = _add_combined_recommendation_columns(
        df, rates_overstock, rates_oos, mtd_days, month_total_days, safety_mode=safety_mode,
        ref_date=ref_date, safety_basis=safety_basis, safety_days=_ssr_days,
    )

    # 銷量帶／銷量存貨分類（常開顯示欄；只作分類顯示，不影響 Safety 建議計算）
    _sales_params, _sales_param_warnings = _hybrid_sales_params(
        False, sales_oos_days, sales_overstock_days, sales_overstock_min_units
    )[1:]
    warnings.extend(_sales_param_warnings)
    raw, _sales_col_warnings = _attach_sales_classification_columns(raw, _sales_params)
    warnings.extend(_sales_col_warnings)

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



def _write_summary_sheet(ws, change_summary, nd_summary, thin_border, cross_tab=None):
    """SUMMARY sheet — 區塊：Safety Qty Change 分類、ND 建議分類 統計表，
    另附 Safety Qty Change 分類 × 分類維度交叉表（cross_tab，Safety Stock Review 引入分類維度用）."""
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
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col_idx, h in enumerate(cols, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        # P0：共用對齊物件，避免逐格新建 Alignment；值與格式語意不變
        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
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
                    cell.alignment = center_align
                elif col_idx == 1:
                    cell.alignment = left_align
                else:
                    cell.alignment = left_wrap
        return header_row + len(block) + 1

    next_row = _write_block(3, "① Safety Qty Change 分類（增加 / 減少 / 不變）", change_summary)
    next_row = _write_block(next_row + 1, "② Recommendation of new Safety Qty / ND 分類", nd_summary)
    if cross_tab is not None and not cross_tab.empty:
        _write_block(next_row + 1, "③ Safety Qty Change 分類 × 分類維度交叉表", cross_tab)
    _auto_width(ws)



def generate_excel_combined(
    summary_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    categories: list[str],
    progress_callback=None,
    dimension: str = "nature",
    om_groups=None,
    stats: dict | None = None,
    include_raw: bool = True,
) -> BytesIO:
    """Safety Stock Review 輸出。dimension／om_groups 決定 SUMMARY 區塊 ③
    「Safety Qty Change 分類 × 分類維度交叉表」的分組（只影響顯示，不影響任何建議計算）；
    om_groups 缺省時由 raw 即時推導。stats 只供首頁「說明」導讀表照錄本次執行參數
    （缺失時顯示預設值），不參與任何計算。輸出 7 工作表：說明（首頁導讀）＋
    RAW／Shop Stock Status／By Shop／店舖問題 SKU／SUMMARY／Safety Stock 數據分析摘要。"""
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    _dim = _hybrid_dimension(dimension)
    if _dim in ("om", "nature_om") and not om_groups and raw_df is not None and not raw_df.empty:
        om_groups = _hybrid_om_groups_from_values(
            raw_df["OM"] if "OM" in raw_df.columns else _safe_default_series("", raw_df.index)
        )
    cross_tab = _build_combined_dimension_cross_tab(raw_df, dimension=_dim, om_groups=om_groups)

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    desc_font = Font(size=9, color="6B7280", italic=True)

    # ── Sheet 0: 說明（邏輯、本次執行參數與導讀；置於工作簿首位）──
    ws_guide = wb.active
    ws_guide.title = "說明"
    _write_combined_guide_sheet(ws_guide, stats, _dim)

    # ── Sheet 1: RAW ──
    ws_raw = wb.create_sheet(title="RAW")

    (_write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.15, columns=COMBINED_RAW_COLS) if include_raw else _write_raw_placeholder(ws_raw))

    if progress_callback:
        progress_callback(0.2, "建立摘要工作表...")

    # ── Sheet 2: Shop Stock Status ──
    ws_summary = wb.create_sheet(title="Shop Stock Status")

    headers = ["Buyer", "Brand", "SKU", "SKU Desc"] + categories + ["總計", "Safety Qty Change", "增加Safety Qty的店舖數量", "減少Safety Qty的店舖數量", "Safety Qty不變的店舖數量"]
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    total_summary = len(summary_df)
    # P0：itertuples＋共用對齊取代 iterrows＋逐格新建 Alignment；值與格式語意不變
    center_align = Alignment(horizontal="center", vertical="center")
    for row_idx, _vals in enumerate(_iter_header_rows(summary_df, headers)):
        excel_row = 2 + row_idx
        is_total = _vals[0] == "總計"
        for col_idx, h in enumerate(headers, 1):
            val = _vals[col_idx - 1]
            if pd.isna(val):
                val = None
            cell = ws_summary.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = center_align
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
        cell.alignment = header_align
        cell.border = thin_border

    for row_idx, _vals in enumerate(_iter_header_rows(by_shop_df, by_shop_headers)):
        excel_row = row_idx + 2
        is_total = _vals[0] == "總計"
        for col_idx, h in enumerate(by_shop_headers, 1):
            val = _vals[col_idx - 1]
            if pd.isna(val):
                val = None
            cell = ws_by_shop.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = center_align
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
    _write_summary_sheet(ws_class_summary, change_summary, nd_summary, thin_border, cross_tab=cross_tab)

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









def _write_combined_guide_sheet(ws, stats, dimension):
    """Safety Stock Review「說明」工作表（工作簿首位）：列出本次執行參數、
    講解 Review 邏輯，並逐頁導讀各工作表應留意的部分。"""
    st_ = stats or {}
    dim = _hybrid_dimension(dimension)
    _os = st_.get("rates_overstock") or {}
    _oos = st_.get("rates_oos") or {}
    os_t = int(round(float(_os.get("tourist", 0.30)) * 100))
    os_m = int(round(float(_os.get("mix", 0.40)) * 100))
    os_n = int(round(float(_os.get("default", 0.50)) * 100))
    oos_t = int(round(float(_oos.get("tourist", 0.50)) * 100))
    oos_m = int(round(float(_oos.get("mix", 0.50)) * 100))
    oos_n = int(round(float(_oos.get("default", 0.50)) * 100))
    mode_txt = ("可加可減（允許減少並套用退場建議）" if st_.get("safety_mode") == "allow_both"
                else "只可增加（Safety 只增不減，ND／P90D 0 Sales 退場建議不套用）")
    _basis = _normalize_ssr_safety_basis(st_.get("safety_basis", "rate"))
    _sd = st_.get("safety_days") or dict(_HYBRID_SAFETY_DAYS_DEFAULT)
    _sd_label = _hybrid_safety_days_label({
        "Tourist": int(_sd.get("Tourist", 7)),
        "Mix": int(_sd.get("Mix", 9)),
        "Non Tourist": int(_sd.get("Non Tourist", 10)),
    })
    basis_txt = ("周轉天數 (按 Nature) — 建議值取 Safety 周轉所需量（含最低 3 件保底）"
                 if _basis == _SSR_SAFETY_BASIS_TURNDAY
                 else "月銷倍率 (X/Y) — 建議值 = max(Y, Z if Z ≤ X else X)")

    params = [
        ("分類維度",
         f"{_HYBRID_DIMENSION_LABELS[dim]}（只影響 SUMMARY 區塊 ③ 的分組，不影響 Safety 建議計算）"),
        ("Safety Qty 計算模式", mode_txt),
        ("Safety 計算基準", f"{basis_txt}；基準天數：{_sd_label}"),
        ("Overstock Rate",
         f"Tourist {os_t}%／Mix {os_m}%／Non Tourist {os_n}%（庫存上限 X = floor(月均銷售 × Rate)，上限不超過現有 Safety）"),
        ("OOS Rate",
         f"Tourist {oos_t}%／Mix {oos_m}%／Non Tourist {oos_n}%（缺貨底線 Y = ceil(月銷預估 × Rate)，最低 3 件）"),
        ("Safety 基準天數",
         f"{_sd_label}（Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 基準天數)，最低 3 件；與調貨段無關）"),
        ("MTD 日期",
         f'MTD {st_.get("mtd_days", "—")} 日（當月共 {st_.get("month_total_days", "—")} 日；'
         "月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均，月初以 P90D 為主）"),
        ("參考日期（New SKU 保護）",
         f'{st_.get("ref_date", st_.get("mtd_date", "—"))}（Launch Date 起 90 天內視為新品，預設 MTD 日期）'),
        ("銷量分類門檻",
         f'缺貨線 {st_.get("sales_oos_days", _HYBRID_SALES_OOS_DAYS_DEFAULT)} 日／'
         f'積壓線 {st_.get("sales_overstock_days", _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT)} 日／'
         f'保底 {st_.get("sales_overstock_min_units", _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT)} 件'
         "（銷量帶與銷量存貨分類恆按此計算）"),
    ]

    logic = [
        "Overstock Cap (X) = floor(月均銷售 × Overstock Rate)，上限不超過現有 Safety 設定"
        "（Safety Stock (現時)，即來源檔案的 Adj. Safety Stock）。",
        "OOS Floor (Y) = ceil(月銷預估 × OOS Rate)，最低 3 件；月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均"
        "（月初以 P90D 為主，月尾收斂至 MTD；P90D 缺值時退回純 MTD 線性外推）。",
        "Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × Safety 基準天數)，最低 3 件"
        "（與 Safety Stock x NetStock Review 同口徑；負數 P90D 按 0 計算）。",
        "建議 Safety Stock（月銷倍率基準）= max(Y, Z if Z ≤ X else X)；"
        "建議 Safety Stock（周轉天數基準）= Safety 周轉所需量（Z = 現有 Safety 設定）；"
        "Safety Qty Change = 建議值 − Z（正=增加、負=減少、零=不變）。",
        "Recommend Safety Stock Turnday = 建議值 ÷ (P90D 月均 ÷ 30)（月均 0 留空；"
        "與來源「Turnday (P90D Qty)」同一準則）。",
        "「只可增加」模式全表只增不減（ND／P90D 0 Sales 的歸零與 ND 退場建議不套用）；"
        "「可加可減」模式套用退場建議：ND 列 → 0；P90D 0 Sales 舊品 → ND"
        "（並依 Adj. MCH1／Adj. SST 加註 ND 類別）。",
        "New SKU 保護：Launch Date 距參考日期 90 天內的列排除「P90D 0 Sales → Safety=0/ND」override，"
        "改走正常基準計算並標記 New SKU (3 Month)。",
        "銷量帶（常開顯示欄）：同 SKU 內按 Sold Qty (P90D) 的店舖平均排名分帶——排名 > 75%＝高銷、"
        "≤ 25%＝低銷、其餘＝中銷、P90D ≤ 0＝零銷；並列同值同帶；有銷售店舖少於 4 間時不設高銷／低銷。",
        "銷量存貨分類（常開顯示欄）：高銷且店舖周轉日數 < 缺貨線 → 高銷缺貨；低銷且周轉 ≥ 積壓線且"
        "庫存 ≥ 保底件數 → 低銷積壓；零銷且庫存 ≥ 保底件數 → 零銷積壓；其餘 → 正常（門檻見上方 ①）；"
        "兩欄與 Safety 建議計算互相獨立。",
        "建議值受最低 3 件保底影響時（最終建議值 = 3 件且未套保底前 < 3），"
        "店舖問題 SKU 的建議方法加註保底原因；建議方法亦附「建議周轉約 X 日」。",
    ]

    sheet_rows = [
        ("說明（本頁）", "Review 邏輯、本次參數與導讀", "執行任何行動前，先核對上方 ① 的參數是否符合預期"),
        ("RAW", "全部店舖×SKU 逐列明細與判定欄",
         "留意「Shop Stock Status」「Safety Qty Change」「Recommend Safety Stock Turnday」「銷量帶」「銷量存貨分類」；"
         "「只可增加」模式不會出現「減少」"),
        ("Shop Stock Status", "採購人員×品牌之庫存異常交叉統計", "各 Buyer／品牌的分類分佈概覽"),
        ("By Shop", "各店舖層級彙總", "留意「安全庫存調整總量」與需增／需減 SKU 數較多的店舖"),
        ("店舖問題 SKU", "每間店舖需行動列（問題＋建議方法）",
         "店務同事可先從這一頁著手；含 Safety 需增加／減少、ND 退場建議與調入缺口跟進；建議方法附建議周轉日數"),
        ("SUMMARY", "① Safety Qty Change 分類 ／ ② ND 建議分類 ／ ③ 分類維度交叉表",
         "看調整方向與總量；③ 的分組按上方分類維度切換"),
        ("Safety Stock 數據分析摘要", "指標卡總覽", "留意 Overstock／OOS／調整行數與調整總量"),
    ]

    tips = [
        "「Safety Qty Change」正數＝建議增加、負數＝建議減少、零＝不變；「只可增加」模式下不會出現負數。",
        "「只可增加」模式下 Overstock Rate 不產生壓縮效果；要壓縮庫存請選「可加可減」。",
        "建議值最低 3 件為保障自動補貨的保底；受影響的店舖問題列會加註原因。",
        "分類維度只改變 SUMMARY 區塊 ③ 的分組顯示，不影響任何 Safety 建議計算。",
    ]

    _write_guide_sheet(
        ws, title="Safety Stock Review 說明與導讀",
        subtitle="本頁講解此報告的 Review 邏輯與本次執行參數，並逐頁說明應留意的部分；各項判定以 RAW 逐列欄位為準。",
        params=params, logic=logic, sheet_rows=sheet_rows, tips=tips,
    )

