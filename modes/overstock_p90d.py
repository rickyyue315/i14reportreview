# -*- coding: utf-8 -*-
"""Overstock (P90D 0 Sales) Review：按店舖統計零銷售積壓 SKU，分 MAKE UP／Non-MAKE UP。

只引用 modes.common，不引用其他模式模組，避免循環引用。

篩選：``Shop Stock Status`` 去除前後空格後等於 ``Overstock (P90D 0 Sales)`` 的列。
MAKE UP 判定：``Adj. MCH1`` 去除前後空格後等於 ``02 - MAKE UP`` 者為 MAKE UP，
其餘（含空白）一律為 Non-MAKE UP。

輸出四個工作表（不輸出 RAW，所有欄位均不含 Value／COGS／GP）：
1. By Shop — 各店舖 Overstock SKU 總數（已刪除重複）及 MAKE UP／Non-MAKE UP 分項
2. Top 10 店舖回覆 — 各店舖按 Shop Stock Qty (CM) 由大至小取前 10 列，附空白「店舖回覆」欄
3. By Nature — 按 Shop Nature 分佈
4. By OM — 按 OM 分佈
"""

import gc
import re
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


from modes.common import (
    _auto_width,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _display_safety_header,
    _extract_shop_name,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    validate_input,
)

__all__ = [
    "transform_overstock_p90d",
    "generate_excel_overstock_p90d",
    "_build_op90d_by_shop",
    "_build_op90d_top10",
    "_build_op90d_by_nature",
    "_build_op90d_by_om",
    "_OP90D_STATUS",
    "_OP90D_MAKEUP_MCH1",
    "_OP90D_BY_SHOP_COLUMNS",
    "_OP90D_TOP10_COLUMNS",
    "_OP90D_BY_NATURE_COLUMNS",
    "_OP90D_BY_OM_COLUMNS",
]

_OP90D_STATUS = "Overstock (P90D 0 Sales)"
_OP90D_MAKEUP_MCH1 = "02 - MAKE UP"

_OP90D_MAKEUP_LABEL = "MAKE UP"
_OP90D_NON_MAKEUP_LABEL = "Non-MAKE UP"

_OP90D_BY_SHOP_COLUMNS = [
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "Overstock SKU 總數",
    "MAKE UP SKU 數",
    "Non-MAKE UP SKU 數",
    "Shop Stock Qty (CM) 總數",
]

_OP90D_TOP10_COLUMNS = [
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "店內排名",
    "Buyer",
    "Brand",
    "SKU",
    "SKU Desc",
    "Adj. MCH1",
    "MAKE UP 分類",
    "RP Type",
    "Item Status",
    "Shop Stock Status",
    "Shop Stock Qty (CM)",
    "Sold Qty (P90D)",
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Excess Stock Qty",
    "Shop On Order PO Qty (CM)",
    "Adj. Safety Stock",
    "店舖回覆",
]

_OP90D_BY_NATURE_COLUMNS = [
    "Shop Nature",
    "涉及店舖數",
    "Overstock SKU 數",
    "Overstock 行數",
    "MAKE UP SKU 數",
    "Non-MAKE UP SKU 數",
    "Shop Stock Qty (CM) 總數",
]

_OP90D_BY_OM_COLUMNS = [
    "OM",
    "涉及店舖數",
    "Overstock SKU 數",
    "Overstock 行數",
    "MAKE UP SKU 數",
    "Non-MAKE UP SKU 數",
    "Shop Stock Qty (CM) 總數",
]

# 以文字形式輸出的識別欄：Excel 內不套用千分位數值格式
_OP90D_TEXT_COLUMNS = frozenset({
    "Site Code", "Shop Name", "OM", "Shop Nature", "Buyer", "Brand", "SKU",
    "SKU Desc", "Adj. MCH1", "MAKE UP 分類", "RP Type", "Item Status",
    "Shop Stock Status", "店舖回覆",
})


def _is_banned_column(name: object) -> bool:
    """含 Value／COGS／GP 的欄位一律不得輸出（本模式固定白名單已排除，此為防線）。"""
    if not isinstance(name, str):
        return False
    lowered = name.lower()
    if "value" in lowered or "cogs" in lowered:
        return True
    return re.search(r"\bgp\b", lowered) is not None


def _classify_makeup(mch1: pd.Series) -> pd.Series:
    """Adj. MCH1 → MAKE UP 分類（去除前後空格後完全相符者為 MAKE UP）。"""
    txt = mch1.fillna("").astype(str).str.strip()
    return pd.Series(
        [_OP90D_MAKEUP_LABEL if v == _OP90D_MAKEUP_MCH1 else _OP90D_NON_MAKEUP_LABEL for v in txt],
        index=mch1.index,
    )


def _display_group(value: object) -> str:
    """分佈表分組顯示：空白（含缺值）顯示為「空白」。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "空白"
    txt = str(value).strip()
    return txt if txt else "空白"


def _representative_nature(group: pd.DataFrame) -> str:
    """同一店舖出現多個 Shop Nature（資料異常）時，取出現次數最多者為代表值。"""
    vals = group.get("Shop Nature", _safe_default_series("", group.index)).fillna("").astype(str).str.strip()
    vals = vals[vals != ""]
    if vals.empty:
        return ""
    return str(vals.value_counts().index[0])


def _prepare_filtered(df_input: pd.DataFrame, shop_natures: list[str] | None = None,
                      progress_callback=None) -> pd.DataFrame:
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")
    avail = set(df_input.columns)
    for col in ("Site Code", "Shop Stock Qty (CM)", "Adj. MCH1"):
        if _resolve_column_name(col, avail) is None:
            raise ValueError(f"缺少必需欄位: {col}")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)
    if df.empty:
        return df
    df = _compute_raw_extra_columns(df)
    df["SKU"] = df["SKU"].apply(_safe_sku)
    status = df.get("Shop Stock Status", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    raw = df[status == _OP90D_STATUS].copy()
    if raw.empty:
        return raw
    raw["MAKE UP 分類"] = _classify_makeup(raw["Adj. MCH1"])
    raw["_is_makeup"] = (raw["MAKE UP 分類"] == _OP90D_MAKEUP_LABEL).astype(int)
    raw["_stock_cm"] = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))
    return raw


def _build_op90d_by_shop(raw: pd.DataFrame) -> pd.DataFrame:
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_OP90D_BY_SHOP_COLUMNS)
    df = raw.copy()
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    df["Site Code"] = site_code
    df["Shop Name"] = [_extract_shop_name(item_shop.iloc[i], site_code.iloc[i]) for i in range(len(df))]
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)

    rows = []
    for (site, name, om), group in df.groupby(["Site Code", "Shop Name", "OM"], dropna=False, observed=True):
        skus = group["SKU"].dropna().astype(str)
        mk = group[group["_is_makeup"] == 1]["SKU"].dropna().astype(str)
        nm = group[group["_is_makeup"] == 0]["SKU"].dropna().astype(str)
        rows.append({
            "Site Code": site,
            "Shop Name": name,
            "OM": om,
            "Shop Nature": _representative_nature(group),
            "Overstock SKU 總數": int(skus.nunique()),
            "MAKE UP SKU 數": int(mk.nunique()),
            "Non-MAKE UP SKU 數": int(nm.nunique()),
            "Shop Stock Qty (CM) 總數": int(float(group["_stock_cm"].sum())),
        })
    by_shop = pd.DataFrame(rows, columns=_OP90D_BY_SHOP_COLUMNS)
    by_shop = by_shop.sort_values(
        ["Overstock SKU 總數", "Shop Stock Qty (CM) 總數", "Site Code"],
        ascending=[False, False, True],
        kind="stable",
    ).reset_index(drop=True)
    if not by_shop.empty:
        total_row = pd.DataFrame([{
            "Site Code": "總計",
            "Shop Name": "",
            "OM": "",
            "Shop Nature": "",
            "Overstock SKU 總數": int(df["SKU"].dropna().astype(str).nunique()),
            "MAKE UP SKU 數": int(df[df["_is_makeup"] == 1]["SKU"].dropna().astype(str).nunique()),
            "Non-MAKE UP SKU 數": int(df[df["_is_makeup"] == 0]["SKU"].dropna().astype(str).nunique()),
            "Shop Stock Qty (CM) 總數": int(float(df["_stock_cm"].sum())),
        }])
        by_shop = pd.concat([by_shop, total_row], ignore_index=True)
    return by_shop[_OP90D_BY_SHOP_COLUMNS]


def _build_op90d_top10(raw: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_OP90D_TOP10_COLUMNS)
    df = raw.copy()
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    df["Site Code"] = site_code
    df["Shop Name"] = [_extract_shop_name(item_shop.iloc[i], site_code.iloc[i]) for i in range(len(df))]
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["Shop Nature"] = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df = df.sort_values(
        ["Site Code", "_stock_cm", "Brand", "SKU"],
        ascending=[True, False, True, True],
        kind="stable",
    )
    parts = []
    for _site, group in df.groupby("Site Code", dropna=False, observed=True):
        head = group.head(top_n).copy()
        head["店內排名"] = range(1, len(head) + 1)
        parts.append(head)
    top10 = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=list(df.columns) + ["店內排名"])
    top10["店舖回覆"] = ""
    available = set(top10.columns)
    cols = [c for c in _OP90D_TOP10_COLUMNS if _resolve_column_name(c, available) is not None]
    if "MAKE UP 分類" not in cols:
        cols.insert(cols.index("Adj. MCH1") + 1 if "Adj. MCH1" in cols else len(cols), "MAKE UP 分類")
    if "店內排名" not in cols:
        cols.append("店內排名")
    if "店舖回覆" not in cols:
        cols.append("店舖回覆")
    return top10[cols]


def _build_distribution(raw: pd.DataFrame, key: str, columns: list[str]) -> pd.DataFrame:
    if raw is None or raw.empty:
        return pd.DataFrame(columns=columns)
    df = raw.copy()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_group"] = df.get(key, _safe_default_series("", df.index)).map(_display_group)
    rows = []
    for grp, sub in df.groupby("_group", dropna=False, observed=True):
        mk = sub[sub["_is_makeup"] == 1]["SKU"].dropna().astype(str)
        nm = sub[sub["_is_makeup"] == 0]["SKU"].dropna().astype(str)
        rows.append({
            key: grp,
            "涉及店舖數": int(sub["Site Code"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
            "Overstock SKU 數": int(sub["SKU"].dropna().astype(str).nunique()),
            "Overstock 行數": int(len(sub)),
            "MAKE UP SKU 數": int(mk.nunique()),
            "Non-MAKE UP SKU 數": int(nm.nunique()),
            "Shop Stock Qty (CM) 總數": int(float(sub["_stock_cm"].sum())),
        })
    dist = pd.DataFrame(rows, columns=columns)
    dist = dist.sort_values(["Overstock 行數", key], ascending=[False, True], kind="stable").reset_index(drop=True)
    if not dist.empty:
        df_mk = df[df["_is_makeup"] == 1]["SKU"].dropna().astype(str)
        df_nm = df[df["_is_makeup"] == 0]["SKU"].dropna().astype(str)
        total_row = pd.DataFrame([{
            key: "總計",
            "涉及店舖數": int(df["Site Code"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
            "Overstock SKU 數": int(df["SKU"].dropna().astype(str).nunique()),
            "Overstock 行數": int(len(df)),
            "MAKE UP SKU 數": int(df_mk.nunique()),
            "Non-MAKE UP SKU 數": int(df_nm.nunique()),
            "Shop Stock Qty (CM) 總數": int(float(df["_stock_cm"].sum())),
        }])
        dist = pd.concat([dist, total_row], ignore_index=True)
    return dist[columns]


def _build_op90d_by_nature(raw: pd.DataFrame) -> pd.DataFrame:
    return _build_distribution(raw, "Shop Nature", _OP90D_BY_NATURE_COLUMNS)


def _build_op90d_by_om(raw: pd.DataFrame) -> pd.DataFrame:
    return _build_distribution(raw, "OM", _OP90D_BY_OM_COLUMNS)


def transform_overstock_p90d(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
):
    """篩選 Overstock (P90D 0 Sales) 列，輸出店舖統計、Top 10、Nature 及 OM 分佈。"""
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    raw = _prepare_filtered(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if progress_callback:
        progress_callback(0.35, "建立店舖統計...")
    by_shop = _build_op90d_by_shop(raw)

    if progress_callback:
        progress_callback(0.5, "擷取各店最高庫存 SKU...")
    top10 = _build_op90d_top10(raw)

    if progress_callback:
        progress_callback(0.65, "建立 Nature 及 OM 分佈...")
    by_nature = _build_op90d_by_nature(raw)
    by_om = _build_op90d_by_om(raw)

    warnings = []
    if raw.empty:
        warnings.append("沒有符合「Overstock (P90D 0 Sales)」的資料。")

    if progress_callback:
        progress_callback(0.7, "完成資料轉換...")
    return raw, by_shop, top10, by_nature, by_om, warnings


def _write_op90d_table(ws, headers, data: pd.DataFrame, progress_callback=None,
                       progress_start=0.0, progress_end=0.1, sheet_name="",
                       auto_filter: bool = False):
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    display_headers = [_display_safety_header(h) for h in headers]
    for col_idx, h in enumerate(display_headers, 1):
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
                if isinstance(val, (int, float)) and h not in _OP90D_TEXT_COLUMNS:
                    cell.number_format = "#,##0"
            if progress_callback and (row_idx % 500 == 0 or row_idx == total - 1):
                p = progress_start + (progress_end - progress_start) * ((row_idx + 1) / max(total, 1))
                progress_callback(p, f"寫入 {sheet_name} ({row_idx+1}/{total})...")
    else:
        ws.cell(row=2, column=1, value="無符合資料").font = Font(size=12)
    if auto_filter and ws.max_row > 1 and ws.max_column > 1:
        ws.auto_filter.ref = ws.dimensions
    _auto_width(ws)


def generate_excel_overstock_p90d(
    by_shop: pd.DataFrame,
    top10: pd.DataFrame,
    by_nature: pd.DataFrame,
    by_om: pd.DataFrame,
    progress_callback=None,
    include_raw: bool = True,
) -> BytesIO:
    """產生四個工作表的 Excel（無 RAW；include_raw 僅為介面相容而保留，不影響輸出）。"""
    for name, frame in (("By Shop", by_shop), ("Top 10", top10), ("By Nature", by_nature), ("By OM", by_om)):
        if frame is not None and not frame.empty:
            banned = [c for c in frame.columns if _is_banned_column(c)]
            if banned:
                raise ValueError(f"{name} 輸出含不應出現的欄位: {', '.join(banned)}")

    wb = Workbook()
    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    ws_shop = wb.active
    ws_shop.title = "By Shop"
    _write_op90d_table(ws_shop, _OP90D_BY_SHOP_COLUMNS, by_shop,
                       progress_callback, 0.0, 0.25, "By Shop")

    if progress_callback:
        progress_callback(0.3, "建立 Top 10 店舖回覆...")
    ws_top = wb.create_sheet(title="Top 10 店舖回覆")
    _write_op90d_table(ws_top, _OP90D_TOP10_COLUMNS, top10,
                       progress_callback, 0.3, 0.6, "Top 10 店舖回覆", auto_filter=True)

    if progress_callback:
        progress_callback(0.65, "建立 By Nature...")
    ws_nature = wb.create_sheet(title="By Nature")
    _write_op90d_table(ws_nature, _OP90D_BY_NATURE_COLUMNS, by_nature,
                       progress_callback, 0.65, 0.78, "By Nature")

    if progress_callback:
        progress_callback(0.8, "建立 By OM...")
    ws_om = wb.create_sheet(title="By OM")
    _write_op90d_table(ws_om, _OP90D_BY_OM_COLUMNS, by_om,
                       progress_callback, 0.8, 0.9, "By OM")

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf
