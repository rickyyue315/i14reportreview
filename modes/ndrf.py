# -*- coding: utf-8 -*-
"""NDRF 月份對比模式：由 transformer.py 拆分而來。
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
    _safe_default_series,
    _safe_numeric,
    _write_raw_placeholder,
    _write_raw_sheet,
)
__all__ = [
    "_build_ndrf_adj_mch1_combined_pivot",
    "_build_ndrf_adj_sst_combined_pivot",
    "_build_ndrf_category_combined_pivot",
    "_build_ndrf_shop_comparison",
    "_write_ndrf_generic_pivot",
    "generate_excel_ndrf_comparison",
    "transform_ndrf_comparison",
]


def _write_ndrf_generic_pivot(ws, pivot_df, thin_border, header_font, header_fill, bold_font):
    """寫入通用樞紐對比表（Adj. SST / MCH2 共用）。"""
    if pivot_df is None or pivot_df.empty:
        ws.cell(row=1, column=1, value="無樞紐資料").font = Font(size=12)
        return
    headers = list(pivot_df.columns)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    # P0：itertuples＋共用對齊取代 iterrows＋逐格新建 Alignment；值與格式語意不變
    center_align = Alignment(horizontal="center", vertical="center")
    for row_idx, _vals in enumerate(_iter_header_rows(pivot_df, headers)):
        excel_row = row_idx + 2
        is_subtotal = "合計" in str(_vals[0] if len(_vals) else "")
        for col_idx, h in enumerate(headers, 1):
            val = _vals[col_idx - 1]
            if pd.isna(val):
                val = None
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = center_align
            if is_subtotal:
                cell.font = bold_font
            if isinstance(val, (int, float)) and val is not None:
                cell.number_format = "#,##0"
    _auto_width(ws)



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
        # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
        if "Item_Shop" in df.columns and "Site Code" in df.columns:
            _shop_item = df["Item_Shop"]
            _shop_site = df["Site Code"].fillna("").astype(str).str.strip()
            df["Shop"] = _cached_shop_names(_shop_item, _shop_site)
        else:
            _item_fallback = df["Item_Shop"] if "Item_Shop" in df.columns else _safe_default_series("", df.index)
            _site_fallback = df["Site Code"] if "Site Code" in df.columns else _safe_default_series("", df.index)
            shop_name_vals = [
                _extract_shop_name(
                    _item_fallback.iloc[i] if hasattr(_item_fallback, "iloc") else "",
                    _site_fallback.iloc[i] if hasattr(_site_fallback, "iloc") else ""
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
        # P0（個別模式）：單次 groupby 取代 size＋pivot＋sum 兩次掃描；計數與合計口徑不變
        try:
            df["_is_ndws"] = (df["_cat"] == "ND_With_Stock").astype(int)
            df["_is_rf"] = (df["_cat"] == "RF").astype(int)
            _agg = df.groupby(["Site Code", category_col], sort=False, observed=True, dropna=False).agg(
                ND_With_Stock=("_is_ndws", "sum"),
                RF_Total=("_is_rf", "sum"),
                Stock_Qty=("_stock_qty", "sum"),
                Stock_Value=("_stock_value", "sum"),
            ).reset_index()
            _agg["ND_With_Stock"] = _agg["ND_With_Stock"].fillna(0).astype(int)
            _agg["RF_Total"] = _agg["RF_Total"].fillna(0).astype(int)
            _agg["Carry"] = (_agg["ND_With_Stock"] + _agg["RF_Total"]).astype(int)
            for c in ["Stock_Qty", "Stock_Value"]:
                _agg[c] = _agg[c].fillna(0).astype(int)
            keep = ["Site Code", category_col, "ND_With_Stock", "RF_Total", "Carry", "Stock_Qty", "Stock_Value"]
            return _agg[keep].sort_values(["Site Code", category_col])
        except Exception:
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
    # P0（個別模式）：去重後建表取代逐列 zip；首見保留語意不變
    try:
        for _src in [df_a, df_b]:
            if _src.empty:
                continue
            sc = _src["Site Code"].fillna("").astype(str).str.strip()
            sn = _src.get("Shop Name", _src.get("Shop", _safe_default_series("", _src.index))).fillna("").astype(str).str.strip()
            _uniq = pd.DataFrame({"_s": sc, "_n": sn}).drop_duplicates(subset=["_s"], keep="first")
            for s, n in zip(_uniq["_s"].to_numpy(), _uniq["_n"].to_numpy()):
                if s and s not in site_to_name:
                    site_to_name[s] = n if n else ""
    except Exception:
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

    # ── Sheet 1: Month Comparison (By Shop) ──
    ws_compare = wb.active
    ws_compare.title = "月份對比 (By Shop)"

    headers = list(comparison_by_shop.columns)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, h in enumerate(headers, 1):
        cell = ws_compare.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # P0：itertuples＋共用對齊取代 iterrows＋逐格新建 Alignment；值與格式語意不變
    center_align = Alignment(horizontal="center", vertical="center")
    for row_idx, _vals in enumerate(_iter_header_rows(comparison_by_shop, headers)):
        excel_row = row_idx + 2
        is_total = _vals[0] == "總計"
        for col_idx, h in enumerate(headers, 1):
            val = _vals[col_idx - 1]
            if pd.isna(val):
                val = None
            cell = ws_compare.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = center_align
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
    (_write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.6, progress_end=0.9) if include_raw else _write_raw_placeholder(ws_raw))

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf

