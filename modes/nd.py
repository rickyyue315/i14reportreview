# -*- coding: utf-8 -*-
"""SKU 精簡 Review模式：由 transformer.py 拆分而來。
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
    _auto_width,
    _cached_shop_names,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _extract_shop_name,
    _iter_header_rows,
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
    "_add_nd_recommendation_columns",
    "_build_nd_by_shop_summary",
    "generate_excel_nd",
    "transform_nd",
]


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
    # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
    shop_name = _cached_shop_names(item_shop, site_code)

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
    include_raw: bool = True,
) -> BytesIO:
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

    ws_summary["A3"] = "行數 - Site Code"
    ws_summary["D3"] = "ND SKU (by Shop Stock Status)"

    transfer_col = "After new Safety Qty / ND can transfer out of QTY"
    headers = ["Buyer", "Brand", "SKU", "SKU Desc"] + categories + ["總計"] + [transfer_col]
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

