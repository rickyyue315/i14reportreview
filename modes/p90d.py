# -*- coding: utf-8 -*-
"""Past 90 Days 0 Sales 模式：由 transformer.py 拆分而來（試點）。

共用 helper 由 transformer 引用；transformer 尾部…（實為原位置）以墊片
re-export 本模組，保持舊有 import 路徑不變。
"""
import gc
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


from modes.common import (
    _auto_width,
    _cached_shop_names,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _extract_shop_name,
    _iter_header_rows,
    _reorder_raw_columns,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _cached_safe_sku_series,
    _write_raw_placeholder,
    _write_raw_sheet,
    validate_input,
)
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
    # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
    shop_name = _cached_shop_names(item_shop, site_code)
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
    include_raw: bool = True,
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
    (_write_raw_sheet(ws_raw, raw, progress_callback, progress_start=0.0, progress_end=0.55, columns=out_cols) if include_raw else _write_raw_placeholder(ws_raw))

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

    # P0：itertuples＋共用對齊取代 iterrows＋逐格新建 Alignment；值與格式語意不變
    center_align = Alignment(horizontal="center", vertical="center")
    if summary_by_shop is not None and not summary_by_shop.empty:
        total_summary = len(summary_by_shop)
        for row_idx, _vals in enumerate(_iter_header_rows(summary_by_shop, summary_headers)):
            excel_row = row_idx + 2
            is_total = _vals[0] == "總計"
            for col_idx, h in enumerate(summary_headers, 1):
                val = _vals[col_idx - 1]
                if pd.isna(val):
                    val = None
                cell = ws_summary.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = center_align
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
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf
