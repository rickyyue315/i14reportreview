# -*- coding: utf-8 -*-
"""OOS Review模式：由 transformer.py 拆分而來。
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
    _display_safety_header,
    _extract_shop_name,
    _iter_header_rows,
    _nature_rate_arr,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _cached_safe_sku_series,
    _write_metric_cards,
    _write_raw_placeholder,
    _write_raw_sheet,
    validate_input,
)
__all__ = [
    "EXCLUDED_STATUSES",
    "OOS_CATEGORIES",
    "SECONDARY_CATEGORY",
    "_add_oos_recommendation_columns",
    "_build_oos_by_shop_summary",
    "generate_excel_oos",
    "transform_oos",
]


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
    # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
    shop_name = _cached_shop_names(item_shop, site_code)

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

    # Safety 來源欄相容檢查：顯示名一律視為來源名（Safety Stock (現時) = 來源檔案的 Adj. Safety Stock）
    _adj_avail = set(df_input.columns) | {_resolve_column_name(_ADJ_SAFETY_SOURCE_COL, set(df_input.columns))}
    _optional_cols = ["Adj. Safety Stock", "Shop Stock Qty (CM)", "Sold Qty (CM - MTD)", "Shop Stock Status"]
    if oos_calc_mode == "p90d":
        _optional_cols.append("Shop Monthly Sold Qty (P90D)")
    for _col in _optional_cols:
        _label = _display_safety_header(_col)
        if _resolve_column_name(_col, _adj_avail) is None:
            warnings.append(f"缺少可選欄位「{_label}」，建議值與轉移數量將全為 0")

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



def generate_excel_oos(summary_df, raw_df, categories, progress_callback=None, include_raw: bool = True):
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    ws_raw = wb.active
    ws_raw.title = "RAW"

    (_write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.3) if include_raw else _write_raw_placeholder(ws_raw))

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
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    data_start = 5
    total_summary = len(summary_df)
    # P0：itertuples＋共用對齊取代 iterrows＋逐格新建 Alignment；值與格式語意不變
    center_align = Alignment(horizontal="center", vertical="center")

    for row_idx, _vals in enumerate(_iter_header_rows(summary_df, headers)):
        excel_row = data_start + row_idx
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

