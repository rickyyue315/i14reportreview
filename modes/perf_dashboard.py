# -*- coding: utf-8 -*-
"""績效儀表板 Review：按 Buyer／Brand／OM 彙總銷售、缺貨與過剩表現，ABC 分級。

只引用 modes.common，不引用其他模式模組，避免循環引用。
"""
import gc
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


from modes.common import (
    _auto_width,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _extract_shop_name,
    _reorder_raw_columns,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _write_raw_placeholder,
    _write_raw_sheet,
    validate_input,
)

__all__ = [
    "transform_perf_dashboard",
    "generate_excel_perf_dashboard",
    "_build_pf_table",
]

_PF_SALES_CANDIDATES = [
    "Net Sales (CM - MTD)",
    "Net Sales (M-1)",
    "Net Sales (M-2)",
]

_PF_RAW_OUTPUT_COLUMNS = [
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "Buyer",
    "Brand",
    "SKU",
    "SKU Desc",
    "RP Type",
    "Shop Stock Status",
    "Shop Stock Qty (CM)",
    "Sold Qty (P90D)",
    "Net Sales (M-2)",
    "Net Sales (M-1)",
    "Net Sales (CM - MTD)",
]


def _pf_sales_column(df: pd.DataFrame) -> tuple[str | None, list[str]]:
    """銷售金額欄解像：CM-MTD → M-1 → M-2 順序取首個存在欄。"""
    warnings: list[str] = []
    for col in _PF_SALES_CANDIDATES:
        if col in df.columns:
            if col != _PF_SALES_CANDIDATES[0]:
                warnings.append(f"來源檔案缺少「{_PF_SALES_CANDIDATES[0]}」，改以「{col}」計算銷售排名。")
            return col, warnings
    return None, warnings


def _pf_raw_columns(raw: pd.DataFrame) -> list[str]:
    r = _reorder_raw_columns(_compute_raw_extra_columns(raw))
    available = set(r.columns)
    return [c for c in _PF_RAW_OUTPUT_COLUMNS if _resolve_column_name(c, available) is not None]


def _pf_abc(sales: pd.Series) -> pd.Series:
    """ABC 分級：按銷售由大至小累計佔比，80% 內為 A，95% 內為 B，其餘為 C；銷售最高者恆為 A。"""
    total = float(sales.sum())
    if total <= 0:
        return pd.Series("C", index=sales.index)
    cum = sales.sort_values(ascending=False).cumsum() / total
    grade = pd.Series("C", index=sales.index)
    grade[cum <= 0.80] = "A"
    grade[(cum > 0.80) & (cum <= 0.95)] = "B"
    grade.loc[sales.idxmax()] = "A"
    return grade


def _build_pf_table(raw: pd.DataFrame, group_col: str, sales_col: str, label: str) -> pd.DataFrame:
    """指定維度績效表：銷售、SKU 數、缺缺貨率、過剩率、ABC 分級，按銷售排序。"""
    headers = [label, "銷售合計", "銷售佔比", "ABC", "SKU 數", "列數", "缺貨列數", "缺貨率", "過剩列數", "過剩率", "庫存合計"]
    if raw is None or raw.empty:
        return pd.DataFrame(columns=headers)
    df = raw.copy()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_sales"] = _safe_numeric(df.get(sales_col, _safe_default_series(0, df.index)))
    if "Shop Stock Qty (CM)" in df.columns:
        df["_oos"] = (_safe_numeric(df["Shop Stock Qty (CM)"]) <= 0).astype(int)
    else:
        df["_oos"] = 0
    status = df.get("Shop Stock Status", _safe_default_series("", df.index)).fillna("").astype(str)
    df["_over"] = status.str.startswith("Overstock").astype(int)
    if "Shop Stock Qty (CM)" in df.columns:
        df["_stock"] = _safe_numeric(df["Shop Stock Qty (CM)"])
    else:
        df["_stock"] = 0
    grp = df.get(group_col, _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    if group_col == "OM":
        grp = grp.mask(grp == "", "（空白）")
    elif (grp == "").all():
        grp = pd.Series("（空白）", index=df.index)
    df["_grp"] = grp.values
    grouped = (
        df.groupby("_grp", dropna=False, observed=True)
        .agg(
            **{
                "銷售合計": ("_sales", "sum"),
                "SKU 數": ("SKU", "nunique"),
                "列數": ("_sales", "size"),
                "缺貨列數": ("_oos", "sum"),
                "過剩列數": ("_over", "sum"),
                "庫存合計": ("_stock", "sum"),
            },
        )
        .reset_index()
    )
    total_sales = float(grouped["銷售合計"].sum())
    total_rows = int(grouped["列數"].sum())
    grouped["銷售佔比"] = grouped["銷售合計"] / total_sales if total_sales > 0 else 0.0
    grouped["缺貨率"] = grouped["缺貨列數"] / grouped["列數"].replace(0, 1)
    grouped["過剩率"] = grouped["過剩列數"] / grouped["列數"].replace(0, 1)
    for c in ["SKU 數", "列數", "缺貨列數", "過剩列數", "庫存合計"]:
        grouped[c] = pd.to_numeric(grouped[c], errors="coerce").fillna(0).round().astype(int)
    grouped["銷售合計"] = pd.to_numeric(grouped["銷售合計"], errors="coerce").fillna(0).round().astype(int)
    grouped = grouped.rename(columns={"_grp": label})
    grouped = grouped.sort_values(["銷售合計", label], ascending=[False, True], kind="stable").reset_index(drop=True)
    grouped["ABC"] = _pf_abc(grouped["銷售合計"]).values
    if not grouped.empty:
        total_row = pd.DataFrame([{
            label: "總計",
            "銷售合計": int(grouped["銷售合計"].sum()),
            "銷售佔比": 1.0 if total_sales > 0 else 0.0,
            "ABC": "",
            "SKU 數": int(df["SKU"].dropna().nunique()),
            "列數": total_rows,
            "缺貨列數": int(grouped["缺貨列數"].sum()),
            "過剩列數": int(grouped["過剩列數"].sum()),
            "缺貨率": (float(grouped["缺貨列數"].sum()) / total_rows) if total_rows else 0.0,
            "過剩率": (float(grouped["過剩列數"].sum()) / total_rows) if total_rows else 0.0,
            "庫存合計": int(grouped["庫存合計"].sum()),
        }])
        grouped = pd.concat([grouped, total_row], ignore_index=True)
    return grouped[headers]


def transform_perf_dashboard(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
):
    """績效儀表板轉換：三維度彙總，輸出 By Buyer、By Brand、By OM、RAW 與提示訊息。"""
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")
    avail = set(df_input.columns)
    if _resolve_column_name("Site Code", avail) is None:
        raise ValueError("缺少必需欄位: Site Code")

    if progress_callback:
        progress_callback(0.1, "清理資料中...")
    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)
    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據")

    if progress_callback:
        progress_callback(0.2, "解像 SKU Desc / Shop Name...")
    df = _compute_raw_extra_columns(df)
    df["SKU"] = df["SKU"].apply(_safe_sku)

    sales_col, sales_warnings = _pf_sales_column(df)
    if sales_col is None:
        raise ValueError(f"缺少銷售金額欄位（需至少一個）: {'、'.join(_PF_SALES_CANDIDATES)}")

    if progress_callback:
        progress_callback(0.3, "建立三維度彙總...")
    by_buyer = _build_pf_table(df, "Buyer", sales_col, "Buyer")
    by_brand = _build_pf_table(df, "Brand", sales_col, "Brand")
    by_om = _build_pf_table(df, "OM", sales_col, "OM")

    warnings = list(sales_warnings)
    for col in ["Shop Stock Qty (CM)", "Sold Qty (P90D)"]:
        if col not in df.columns:
            warnings.append(f"來源檔案缺少「{col}」，相關欄以空值顯示。")
    if "OM" not in df.columns:
        warnings.append("來源檔案缺少「OM」欄，By OM 合併為單一「（空白）」組。")

    if progress_callback:
        progress_callback(0.4, "完成資料轉換...")
    return by_buyer, by_brand, by_om, df, sales_col, warnings


def _write_pf_table(ws, headers, data: pd.DataFrame, progress_callback=None, progress_start=0.0, progress_end=0.1, sheet_name=""):
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    if data is not None and not data.empty:
        total = len(data)
        for row_idx, (_, row_data) in enumerate(data.iterrows()):
            excel_row = row_idx + 2
            is_total = str(row_data.iloc[0]) == "總計"
            for col_idx, h in enumerate(headers, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if is_total:
                    cell.font = bold_font
                if h in ("銷售佔比", "缺貨率", "過剩率") and isinstance(val, (int, float)):
                    cell.number_format = "0.00%"
                elif isinstance(val, (int, float)) and col_idx > 1:
                    cell.number_format = "#,##0"
            if progress_callback and (row_idx % 500 == 0 or row_idx == total - 1):
                p = progress_start + (progress_end - progress_start) * ((row_idx + 1) / max(total, 1))
                progress_callback(p, f"寫入 {sheet_name} ({row_idx+1}/{total})...")
    else:
        ws.cell(row=2, column=1, value="無符合資料").font = Font(size=12)
    _auto_width(ws)


def generate_excel_perf_dashboard(
    by_buyer: pd.DataFrame,
    by_brand: pd.DataFrame,
    by_om: pd.DataFrame,
    raw: pd.DataFrame,
    progress_callback=None,
    include_raw: bool = True,
) -> BytesIO:
    wb = Workbook()
    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    ws_buyer = wb.active
    ws_buyer.title = "By Buyer"
    _write_pf_table(ws_buyer, list(by_buyer.columns) if by_buyer is not None and not by_buyer.empty else
                    ["Buyer", "銷售合計", "銷售佔比", "ABC", "SKU 數", "列數", "缺貨列數", "缺貨率", "過剩列數", "過剩率", "庫存合計"],
                    by_buyer, progress_callback, 0.0, 0.25, "By Buyer")

    if progress_callback:
        progress_callback(0.3, "建立 By Brand...")
    ws_brand = wb.create_sheet(title="By Brand")
    _write_pf_table(ws_brand, list(by_brand.columns) if by_brand is not None and not by_brand.empty else
                    ["Brand", "銷售合計", "銷售佔比", "ABC", "SKU 數", "列數", "缺貨列數", "缺貨率", "過剩列數", "過剩率", "庫存合計"],
                    by_brand, progress_callback, 0.3, 0.5, "By Brand")

    if progress_callback:
        progress_callback(0.55, "建立 By OM...")
    ws_om = wb.create_sheet(title="By OM")
    _write_pf_table(ws_om, list(by_om.columns) if by_om is not None and not by_om.empty else
                    ["OM", "銷售合計", "銷售佔比", "ABC", "SKU 數", "列數", "缺貨列數", "缺貨率", "過剩列數", "過剩率", "庫存合計"],
                    by_om, progress_callback, 0.55, 0.65, "By OM")

    out_cols = _pf_raw_columns(raw if raw is not None else pd.DataFrame())
    ws_raw = wb.create_sheet(title="RAW")
    if include_raw:
        _write_raw_sheet(ws_raw, raw, progress_callback, progress_start=0.65, progress_end=0.9, columns=out_cols)
    else:
        _write_raw_placeholder(ws_raw)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf
