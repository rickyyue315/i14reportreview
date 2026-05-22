import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

REQUIRED_COLUMNS = ["Buyer", "Brand", "SKU", "Shop Stock Status"]

RAW_COLUMN_ORDER_NAMES = [
    "Item_Shop",
    "SKU",
    "Brand",
    "Buyer",
    "Adj. SST",
    "Vendor MC",
    "MCH2",
    "Item Status",
    "Launch Date",
    "Item Class",
    "Site Code",
    "OM",
    "Shop Nature",
    "RP Type",
    "Shop Stock Status",
    "Adj. Safety Stock",
    "Shop Monthly Sold Qty (P90D)",
    "Safety VS Monthly Sold Qty",
    "Recommendation of new Safety Qty / ND",
    "After new Safety Qty / ND can transfer out of QTY",
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
    if name in available:
        return name
    for col in available:
        if col.startswith(name):
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


def _safe_numeric(series: pd.Series, default: float = 0.0) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(default)


def _add_recommendation_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    status_col = df["Shop Stock Status"].fillna("")

    adj_ss = _safe_numeric(df.get("Adj. Safety Stock", pd.Series([0] * len(df))))

    cond_overstock_2x = status_col.str.contains(
        "Overstock (>2*Monthly Sold Qty & >2pcs)", regex=False, na=False
    ) & (adj_ss > 3)

    cond_overstock_0 = status_col.str.contains(
        "Overstock (P90D 0 Sales)", regex=False, na=False
    ) & (adj_ss > 0)

    rec_values = np.full(len(df), "", dtype=object)
    rec_values[cond_overstock_2x.values] = "3"
    rec_values[cond_overstock_0.values] = "ND"
    df["Recommendation of new Safety Qty / ND"] = rec_values

    shop_stock = _safe_numeric(df.get("Shop Stock Qty (CM)", pd.Series([0] * len(df))))

    rec = df["Recommendation of new Safety Qty / ND"]
    mask_3 = (rec == "3") & (shop_stock > 3)
    mask_nd = (rec == "ND") & (shop_stock > 0)

    transfer_values = np.full(len(df), "", dtype=object)
    transfer_values[mask_3.values] = (shop_stock[mask_3] - 3).astype(int).values
    transfer_values[mask_nd.values] = (shop_stock[mask_nd]).astype(int).values
    df["After new Safety Qty / ND can transfer out of QTY"] = transfer_values

    df["Buyer reply"] = ""

    return df


def transform(df_input: pd.DataFrame, progress_callback=None):
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"\u7f3a\u5c11\u5fc5\u8981\u6b04\u4f4d: {', '.join(missing)}")

    if progress_callback:
        progress_callback(0.05, "\u9a57\u8b49\u901a\u904e")

    df = df_input.dropna(subset=["Buyer"]).copy()

    if progress_callback:
        progress_callback(0.1, "\u6e05\u7406\u8cc7\u6599\u4e2d...")

    _footer_mask = df["Buyer"].astype(str).str.strip().str.lower().str.startswith(("total", "applied filters"))
    if "Item_Shop" in df.columns:
        _item_shop_mask = df["Item_Shop"].astype(str).str.strip().str.lower().str.startswith("total")
        _footer_mask = _footer_mask | _item_shop_mask
    df = df[~_footer_mask].copy()

    status_col = df["Shop Stock Status"].fillna("")

    categories = [
        "Overstock (P90D 0 Sales)",
        "Overstock (>2*Monthly Sold Qty & >2pcs)",
    ]

    group_cols = ["Buyer", "Brand", "SKU"]

    if progress_callback:
        progress_callback(0.2, "\u5206\u985e\u7d71\u8a08\u4e2d...")

    counts = {}
    for i, cat in enumerate(categories):
        mask = status_col.str.contains(cat, regex=False, na=False)
        counts[cat] = (
            df[mask]
            .groupby(group_cols, observed=True)
            .size()
            .rename(cat)
        )
        if progress_callback:
            progress_callback(0.2 + 0.3 * (i + 1) / len(categories), f"\u7d71\u8a08 {cat}...")

    if progress_callback:
        progress_callback(0.5, "\u6392\u5e8f\u5f59\u7e3d\u4e2d...")

    summary = pd.concat(counts.values(), axis=1).fillna(0).astype(int).reset_index()

    summary["SKU"] = summary["SKU"].apply(lambda x: str(int(x)) if pd.notna(x) else x)
    summary["\u7e3d\u8a08"] = summary[categories].sum(axis=1)

    buyer_order = summary.groupby("Buyer")["\u7e3d\u8a08"].sum().sort_values(ascending=False).index.tolist()
    summary["_buyer_key"] = summary["Buyer"].map({b: i for i, b in enumerate(buyer_order)})
    summary = summary.sort_values(["_buyer_key", "Brand"], ascending=[True, True]).drop(columns=["_buyer_key"]).reset_index(drop=True)

    total_row = pd.DataFrame(
        {
            "Buyer": ["\u7e3d\u8a08"],
            "Brand": [None],
            "SKU": [None],
            **{cat: [int(summary[cat].sum())] for cat in categories},
            "\u7e3d\u8a08": [int(summary["\u7e3d\u8a08"].sum())],
        }
    )
    summary = pd.concat([summary, total_row], ignore_index=True)

    if progress_callback:
        progress_callback(0.6, "\u8a08\u7b97\u5eab\u5b58\u5efa\u8b70\u4e2d...")

    raw = df.copy()
    raw["SKU"] = raw["SKU"].apply(lambda x: str(int(x)) if pd.notna(x) else x)
    raw = _add_recommendation_columns(raw)

    if progress_callback:
        progress_callback(0.7, "\u8f49\u63db\u5b8c\u6210")

    return summary, raw, categories


def generate_excel(summary_df: pd.DataFrame, raw_df: pd.DataFrame, categories: list[str], progress_callback=None) -> BytesIO:
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "\u5efa\u7acb Excel \u7d50\u69cb...")

    ws_raw = wb.active
    ws_raw.title = "RAW"

    raw_df = _reorder_raw_columns(raw_df)

    raw_cols = list(raw_df.columns)
    ws_raw.append(raw_cols)

    raw_values = raw_df.values.tolist()
    total_raw = len(raw_values)
    for i, row_vals in enumerate(raw_values):
        cleaned = [None if isinstance(v, float) and pd.isna(v) else v for v in row_vals]
        ws_raw.append(cleaned)
        if progress_callback and (i % 1000 == 0 or i == total_raw - 1):
            progress_callback(0.3 * ((i + 1) / total_raw), f"\u5beb\u5165 RAW ({i+1}/{total_raw})...")

    header_fill_raw = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    ncols = len(raw_cols)
    for col_idx in range(1, ncols + 1):
        cell = ws_raw.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        cell.fill = header_fill_raw
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    ws_raw.auto_filter.ref = ws_raw.dimensions

    if progress_callback:
        progress_callback(0.35, "\u5efa\u7acb\u6458\u8981\u5de5\u4f5c\u8868...")

    ws_summary = wb.create_sheet(title="Shop Stock Status")

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)

    ws_summary["A3"] = "\u884c\u6578 - Site Code"
    ws_summary["D3"] = "Shop Stock Status"

    headers = ["Buyer", "Brand", "SKU"] + categories + ["\u7e3d\u8a08"]
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
        progress_callback(0.90, "\u5efa\u7acb\u5206\u6790\u6458\u8981\u5de5\u4f5c\u8868...")

    ws_dashboard = wb.create_sheet(title="\u7570\u5e38\u5eab\u5b58\u6578\u64da\u5206\u6790\u6458\u8981")

    total_overstock = int(summary_df[summary_df["Buyer"] == "\u7e3d\u8a08"]["\u7e3d\u8a08"].values[0])
    total_buyers = len(summary_df[summary_df["Buyer"] != "\u7e3d\u8a08"]["Buyer"].dropna().unique())
    total_brands = len(summary_df[summary_df["Buyer"] != "\u7e3d\u8a08"]["Brand"].dropna().unique())
    total_skus = raw_df["SKU"].dropna().nunique()
    cat1_val = int(summary_df[summary_df["Buyer"] == "\u7e3d\u8a08"][categories[0]].values[0])
    cat2_val = int(summary_df[summary_df["Buyer"] == "\u7e3d\u8a08"][categories[1]].values[0])

    transfer_col = _resolve_column_name("After new Safety Qty / ND can transfer out of QTY", set(raw_df.columns))
    transfer_qty_total = 0
    if transfer_col:
        transfer_qty_total = int(pd.to_numeric(raw_df[transfer_col], errors="coerce").fillna(0).sum())

    ws_dashboard.merge_cells("A1:D1")
    title_cell = ws_dashboard.cell(row=1, column=1, value="\u7570\u5e38\u5eab\u5b58\u6578\u64da\u5206\u6790\u6458\u8981 (Analysis Dashboard)")
    title_cell.font = Font(bold=True, size=14, color="1E40AF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    row1_metrics = [
        ("\u7e3d Overstock \u7b46\u6578", total_overstock, "\u6240\u6709\u5eab\u5b58\u7570\u5e38\u4e4b\u7e3d\u8a08\u7b46\u6578", "EF4444"),
        ("\u5f71\u97ff\u63a1\u8cfc\u4eba\u54e1\u6578 (Buyers)", total_buyers, "\u4f4d\u63a1\u8cfc\u4eba\u54e1\u540d\u4e0b\u6709\u7570\u5e38\u54c1\u9805", "3B82F6"),
        ("\u5f71\u97ff\u54c1\u724c\u7e3d\u6578 (Brands)", total_brands, "\u500b\u54c1\u724c\u5305\u542b\u7570\u5e38\u5eab\u5b58\u5546\u54c1", "8B5CF6"),
        ("\u5f71\u97ffSKU\u7e3d\u6578", total_skus, "\u500b\u5305\u542b\u7570\u5e38\u5eab\u5b58\u5546\u54c1(SKUs)", "10B981"),
    ]

    desc_font = Font(size=9, color="6B7280", italic=True)

    for i, (name, value, desc, color) in enumerate(row1_metrics):
        col = i + 1
        tc = ws_dashboard.cell(row=3, column=col, value=name)
        tc.font = Font(bold=True, size=11, color="FFFFFF")
        tc.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        tc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        tc.border = thin_border

        vc = ws_dashboard.cell(row=4, column=col, value=value)
        vc.font = Font(bold=True, size=20, color=color)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.number_format = "#,##0"
        vc.border = thin_border

        dc = ws_dashboard.cell(row=5, column=col, value=desc)
        dc.font = desc_font
        dc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        dc.border = thin_border

    row2_metrics = [
        ("Overstock (P90D 0 Sales)", cat1_val, "\u300c\u904e\u53bb 90 \u5929\u5167\u96f6\u92b7\u552e\u300d\u4e4b\u7570\u5e38\u7b46\u6578", "F59E0B"),
        ("Overstock (>2*Monthly Sold Qty & >2pcs)", cat2_val, "\u300c\u5927\u65bc\u6708\u5747\u92b7\u552e\u91cf\u5169\u500d\u4e14\u5eab\u5b58\u5927\u65bc 2 \u4ef6\u300d\u4e4b\u7570\u5e38\u7b46\u6578", "EC4899"),
        ("New Safety Qty / ND \u8abf\u6574\u5f8c\u7684\u9810\u8a08\u7e3d\u6578", transfer_qty_total, "After new Safety Qty / ND can transfer out of TOTAL QTY", "06B6D4"),
    ]

    for i, (name, value, desc, color) in enumerate(row2_metrics):
        col = i + 1
        tc = ws_dashboard.cell(row=7, column=col, value=name)
        tc.font = Font(bold=True, size=11, color="FFFFFF")
        tc.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        tc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        tc.border = thin_border

        vc = ws_dashboard.cell(row=8, column=col, value=value)
        vc.font = Font(bold=True, size=20, color=color)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        vc.number_format = "#,##0"
        vc.border = thin_border

        dc = ws_dashboard.cell(row=9, column=col, value=desc)
        dc.font = desc_font
        dc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        dc.border = thin_border

    if progress_callback:
        progress_callback(0.95, "\u8abf\u6574\u683c\u5f0f\u4e2d...")

    _auto_width(ws_dashboard)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "\u5b8c\u6210!")

    return buf


def _auto_width(ws, min_width=8, max_width=50, sample_rows=200):
    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        scan_end = min(ws.max_row, 1 + sample_rows)
        for row_idx in range(1, scan_end + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value is not None:
                val_len = len(str(cell.value))
                if val_len > max_len:
                    max_len = val_len
        adjusted = min(max(max_len + 2, min_width), max_width)
        ws.column_dimensions[col_letter].width = adjusted
