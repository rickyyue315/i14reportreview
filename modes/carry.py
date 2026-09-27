# -*- coding: utf-8 -*-
"""Carry SKU Review模式：由 transformer.py 拆分而來。
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
    _clean_dataframe,
    _extract_shop_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _write_raw_placeholder,
    _write_raw_sheet,
    validate_input,
)
__all__ = [
    "CARRY_P90D0_STATUS",
    "_build_carry_dashboard",
    "_build_carry_pivot",
    "_build_site_to_name",
    "_write_pivot_block",
    "_write_summary_headers",
    "generate_excel_carry",
    "transform_carry",
]


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
    include_raw: bool = True,
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

    (_write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.1) if include_raw else _write_raw_placeholder(ws_raw))

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

