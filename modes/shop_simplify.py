# -*- coding: utf-8 -*-
"""店舖精簡SKU模式：把 By SKUxShop 明細轉化為店舖精簡 SKU 回覆表。

每列一店×一 SKU，保留來源檔全部欄位，並在最後附加五個固定欄位：

- 過去90日有無賣過：Sold Qty (P90D) > 0 → 「有」，否則（等於 0、負數或空白）→「無賣」；
- 過去60日有無存貨：Shop Stock Qty (LM) 或 Shop Stock Qty (CM) > 0 →「有存貨」，
  兩者同時為 0（或空白／負數）→「無存貨」；
- SCT建議、Shop 回覆新的NDRF、如果RF的Safety Qty：一律留空，供 SCT／店舖填寫。

本模式不作調貨、Safety 或退倉建議，亦不篩選列（全店全 SKU 輸出）；清理規則沿用
`_clean_dataframe`（剔除合計列與 117 - Marketing），Shop Nature 篩選為選填。
另由 Item_Shop 以 Site Code 解析 SKU Desc 與 Shop Name（SKU Desc 排於 Brand 之後、
Shop Name 排於 Site Code 之後），其餘來源欄序不變；欄名含「Value」的欄位一律刪除
（大小寫敏感，與「RAW 明細轉化 (不含 Value)」同一判定，故 Valuation Class 不受影響）。
輸出單一工作表 Export，欄序與來源檔一致，五個回覆欄固定排最後。
"""
import gc
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from modes.common import (
    _auto_width,
    _cached_safe_sku_series,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _resolve_column_name,
    _safe_numeric,
    validate_input,
)
from modes.raw_convert import _is_raw_convert_value_column

__all__ = [
    "_SHOP_SIMPLIFY_SOLD_COLUMN",
    "_SHOP_SIMPLIFY_STOCK_COLUMN",
    "_SHOP_SIMPLIFY_SCT_COLUMN",
    "_SHOP_SIMPLIFY_NDRF_COLUMN",
    "_SHOP_SIMPLIFY_SAFETY_COLUMN",
    "_SHOP_SIMPLIFY_REPLY_COLUMNS",
    "_SHOP_SIMPLIFY_SOLD_YES",
    "_SHOP_SIMPLIFY_SOLD_NO",
    "_SHOP_SIMPLIFY_STOCK_YES",
    "_SHOP_SIMPLIFY_STOCK_NO",
    "shop_simplify_stats",
    "transform_shop_simplify",
    "generate_excel_shop_simplify",
]


# ── 輸出欄位契約（欄名固定，測試與導出檔依賴，未經明確要求不得改動）──
_SHOP_SIMPLIFY_SOLD_COLUMN = "過去90日有無賣過"
_SHOP_SIMPLIFY_STOCK_COLUMN = "過去60日有無存貨"
_SHOP_SIMPLIFY_SCT_COLUMN = "SCT建議"
_SHOP_SIMPLIFY_NDRF_COLUMN = "Shop 回覆新的NDRF"
_SHOP_SIMPLIFY_SAFETY_COLUMN = "如果RF的Safety Qty"

# 需要由 SCT／店舖填寫的三欄（一律留空；不沿用來源檔舊值）
_SHOP_SIMPLIFY_REPLY_COLUMNS = [
    _SHOP_SIMPLIFY_SCT_COLUMN,
    _SHOP_SIMPLIFY_NDRF_COLUMN,
    _SHOP_SIMPLIFY_SAFETY_COLUMN,
]

# 兩個判定欄的固定取值
_SHOP_SIMPLIFY_SOLD_YES = "有"
_SHOP_SIMPLIFY_SOLD_NO = "無賣"
_SHOP_SIMPLIFY_STOCK_YES = "有存貨"
_SHOP_SIMPLIFY_STOCK_NO = "無存貨"


def shop_simplify_stats(df: pd.DataFrame) -> dict:
    """結果頁指標卡用的統計（不影響輸出檔）：列數、無賣、無存貨、同時無賣且無存貨、店舖數、SKU 數。"""
    stats = {
        "rows": 0, "sold_no": 0, "stock_no": 0, "both_no": 0, "sites": 0, "skus": 0,
    }
    if df is None or len(df) == 0:
        return stats
    sold_no = df.get(_SHOP_SIMPLIFY_SOLD_COLUMN, pd.Series("", index=df.index)).astype(str) == _SHOP_SIMPLIFY_SOLD_NO
    stock_no = df.get(_SHOP_SIMPLIFY_STOCK_COLUMN, pd.Series("", index=df.index)).astype(str) == _SHOP_SIMPLIFY_STOCK_NO
    stats["rows"] = int(len(df))
    stats["sold_no"] = int(sold_no.sum())
    stats["stock_no"] = int(stock_no.sum())
    stats["both_no"] = int((sold_no & stock_no).sum())
    if "Site Code" in df.columns:
        sites = df["Site Code"].dropna().astype(str).str.strip()
        stats["sites"] = int(sites[sites != ""].nunique())
    if "SKU" in df.columns:
        skus = df["SKU"].dropna().astype(str).str.strip()
        stats["skus"] = int(skus[skus != ""].nunique())
    return stats


def _shop_simplify_place_derived_columns(df: pd.DataFrame) -> pd.DataFrame:
    """把解析出的 SKU Desc／Shop Name 移到慣用位置，其餘來源欄序不變。

    SKU Desc 排於 Brand 之後（缺 Brand 時依序退回 SKU／Item_Shop）；
    Shop Name 排於 Site Code 之後（缺 Site Code 時依序退回 OM／Shop Nature／SKU Desc）。
    """
    def _move(frame: pd.DataFrame, column: str, anchors: tuple[str, ...]) -> pd.DataFrame:
        if column not in frame.columns:
            return frame
        cols = [c for c in frame.columns if c != column]
        for anchor in anchors:
            if anchor in cols:
                cols.insert(cols.index(anchor) + 1, column)
                break
        else:
            cols.append(column)
        return frame[cols]

    out = _move(df, "SKU Desc", ("Brand", "SKU", "Item_Shop"))
    out = _move(out, "Shop Name", ("Site Code", "OM", "Shop Nature", "SKU Desc"))
    return out


def _shop_simplify_drop_existing_columns(df: pd.DataFrame) -> pd.DataFrame:
    """移除來源檔已有的五個輸出欄（含前後空格變體），避免重複欄位。

    回傳移除後的 df；是否有欄被移除由呼叫方比對欄數自行判斷。
    """
    avail = set(df.columns)
    drop = []
    for name in [_SHOP_SIMPLIFY_SOLD_COLUMN, _SHOP_SIMPLIFY_STOCK_COLUMN, *_SHOP_SIMPLIFY_REPLY_COLUMNS]:
        resolved = _resolve_column_name(name, avail)
        if resolved is not None:
            drop.append(resolved)
    return df.drop(columns=drop) if drop else df


def transform_shop_simplify(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
):
    """清理來源資料並附加「過去90日有無賣過／過去60日有無存貨」與三個空白回覆欄。

    回傳 (raw_df, warnings)。缺少 Sold Qty (P90D)，或 Shop Stock Qty (LM) 與
    Shop Stock Qty (CM) 同時缺失時拋出 ValueError（無法判定，不應靜默輸出）。
    """
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必需欄位: {', '.join(missing)}")

    avail = set(df_input.columns)
    sold_col = _resolve_column_name("Sold Qty (P90D)", avail)
    lm_col = _resolve_column_name("Shop Stock Qty (LM)", avail)
    cm_col = _resolve_column_name("Shop Stock Qty (CM)", avail)
    if sold_col is None:
        raise ValueError("來源檔案缺少「Sold Qty (P90D)」，無法判定過去90日有無賣過。")
    if lm_col is None and cm_col is None:
        raise ValueError(
            "來源檔案同時缺少「Shop Stock Qty (LM)」與「Shop Stock Qty (CM)」，無法判定過去60日有無存貨。"
        )

    if progress_callback:
        progress_callback(0.15, "清理資料中...")
    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback).copy()

    warnings: list[str] = []
    if lm_col is None:
        warnings.append("來源檔案缺少「Shop Stock Qty (LM)」，過去60日有無存貨只以 Shop Stock Qty (CM) 判定。")
    if cm_col is None:
        warnings.append("來源檔案缺少「Shop Stock Qty (CM)」，過去60日有無存貨只以 Shop Stock Qty (LM) 判定。")

    had_output_columns = any(
        _resolve_column_name(name, avail) is not None
        for name in [_SHOP_SIMPLIFY_SOLD_COLUMN, _SHOP_SIMPLIFY_STOCK_COLUMN, *_SHOP_SIMPLIFY_REPLY_COLUMNS]
    )
    if had_output_columns:
        warnings.append(
            "來源檔已含本模式的輸出欄（過去90日有無賣過／過去60日有無存貨／SCT建議／"
            "Shop 回覆新的NDRF／如果RF的Safety Qty），將以本次計算結果覆蓋，三個回覆欄一律留空。"
        )

    if progress_callback:
        progress_callback(0.35, "判定過去90日有無賣過...")
    df = _shop_simplify_drop_existing_columns(df)

    if "Item_Shop" not in df.columns:
        warnings.append("來源檔案缺少「Item_Shop」，SKU Desc 與 Shop Name 留空。")
    df = _compute_raw_extra_columns(df)
    df = _shop_simplify_place_derived_columns(df)

    value_columns = [c for c in df.columns if _is_raw_convert_value_column(c)]
    if value_columns:
        df = df.drop(columns=value_columns)

    sold = _safe_numeric(df[sold_col]) if sold_col in df.columns else pd.Series(0.0, index=df.index)
    sold_status = np.where(
        sold.to_numpy(dtype=float) > 0,
        _SHOP_SIMPLIFY_SOLD_YES,
        _SHOP_SIMPLIFY_SOLD_NO,
    )

    if progress_callback:
        progress_callback(0.55, "判定過去60日有無存貨...")
    lm_stock = (
        _safe_numeric(df[lm_col]) if lm_col is not None and lm_col in df.columns else pd.Series(0.0, index=df.index)
    )
    cm_stock = (
        _safe_numeric(df[cm_col]) if cm_col is not None and cm_col in df.columns else pd.Series(0.0, index=df.index)
    )
    has_stock = (lm_stock.to_numpy(dtype=float) > 0) | (cm_stock.to_numpy(dtype=float) > 0)
    stock_status = np.where(has_stock, _SHOP_SIMPLIFY_STOCK_YES, _SHOP_SIMPLIFY_STOCK_NO)

    df[_SHOP_SIMPLIFY_SOLD_COLUMN] = sold_status
    df[_SHOP_SIMPLIFY_STOCK_COLUMN] = stock_status
    for name in _SHOP_SIMPLIFY_REPLY_COLUMNS:
        df[name] = ""

    # SKU 統一為文字（與其他模式一致，避免 Excel 以科學記號顯示長編號）
    if "SKU" in df.columns:
        df["SKU"] = _cached_safe_sku_series(df["SKU"])

    if df.empty:
        warnings.append("清理後沒有可輸出的資料列，請確認來源檔案與 Shop Nature 篩選。")

    if progress_callback:
        progress_callback(0.7, "完成資料轉換...")
    return df, warnings


def _shop_simplify_cell(value):
    """Excel 儲存格值正規化：空白字串與各類缺值一律寫成 None（真正空白）。"""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, str):
        return value if value != "" else None
    if isinstance(value, np.bool_):
        return bool(value)
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, np.floating):
        return round(float(value), 1)
    return value


def generate_excel_shop_simplify(raw: pd.DataFrame, progress_callback=None) -> BytesIO:
    """產生單一工作表 Export：來源欄序不變，五個輸出欄固定排最後。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Export"

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    df = raw if raw is not None else pd.DataFrame()
    raw_cols = list(df.columns)
    ws.append([_shop_simplify_cell(c) for c in raw_cols])

    total = len(df)
    for i, row in enumerate(df.itertuples(index=False, name=None)):
        ws.append([_shop_simplify_cell(v) for v in row])
        if progress_callback and (i % 2000 == 0 or i == total - 1):
            progress_callback(0.05 + 0.85 * ((i + 1) / total), f"寫入 Export ({i+1}/{total})...")

    ncols = len(raw_cols)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    for col_idx in range(1, ncols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center")

    # 凍結窗格設於第一個回覆欄之前：向右捲動時其餘欄位保持可見（與目標檔一致）
    n_source_cols = max(ncols - 5, 0)
    if n_source_cols > 0:
        ws.freeze_panes = f"{get_column_letter(n_source_cols + 1)}2"

    ws.auto_filter.ref = ws.dimensions
    _auto_width(ws)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")
    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf
