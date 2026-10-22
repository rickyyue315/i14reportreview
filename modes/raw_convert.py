# -*- coding: utf-8 -*-
"""RAW 明細轉化：純粹把來源 By SKUxShop 轉化為 RAW 工作表，不作任何統計或判定。

- 不計算 OOS／Overstock／P90D 0 Sales／調貨／Safety，不篩選列，不產生摘要工作表；
- 保留來源檔案全部欄位（欄序沿用 RAW 可讀性重排），並由 Item_Shop 解析 SKU Desc 與 Shop Name
  （原 I14 沒有獨立的品項描述與店舖名稱欄位）；
- 清理規則與其他模式一致（剔除合計列、117 - Marketing），Shop Nature 篩選為選填；
- 兩個變體：預設保留全部欄位；「不含 Value」變體刪除所有欄名含「Value」的欄位
  （大小寫敏感，與其他模式 RAW 隱藏金額欄的規則一致，故「Valuation Class」不受影響）。
"""
import gc
from io import BytesIO

import pandas as pd
from openpyxl import Workbook

from modes.common import (
    _clean_dataframe,
    _compute_raw_extra_columns,
    _reorder_raw_columns,
    _safe_sku,
    _write_raw_sheet,
    validate_input,
)

__all__ = [
    "_RAW_CONVERT_NO_VALUE_SUFFIX",
    "_is_raw_convert_value_column",
    "transform_raw_convert",
    "generate_excel_raw_convert",
]


# 「不含 Value」變體的顯示名稱後綴（介面選項與下載檔名共用）
_RAW_CONVERT_NO_VALUE_SUFFIX = " (不含 Value)"


def _is_raw_convert_value_column(column) -> bool:
    """欄名是否含「Value」（大小寫敏感，與其他模式 RAW 隱藏金額欄的規則一致）。

    「Valuation Class」不含「Value」，不會被誤刪。
    """
    return isinstance(column, str) and "Value" in column


def transform_raw_convert(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
    include_value: bool = True,
):
    """清理來源資料並輸出 RAW 明細（不作統計、不篩選列）。

    ``include_value=False`` 時刪除所有欄名含「Value」的欄位（金額欄）。
    回傳 (raw_df, warnings)。缺少共用必要欄位時拋出 ValueError。
    """
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必需欄位: {', '.join(missing)}")

    if progress_callback:
        progress_callback(0.1, "清理資料中...")
    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if progress_callback:
        progress_callback(0.3, "解析 SKU Desc / Shop Name...")
    raw = _compute_raw_extra_columns(df)
    if "SKU" in raw.columns:
        raw["SKU"] = raw["SKU"].apply(_safe_sku)
    raw = _reorder_raw_columns(raw)

    value_columns = [c for c in raw.columns if _is_raw_convert_value_column(c)]
    if not include_value and value_columns:
        raw = raw.drop(columns=value_columns)

    warnings = []
    if not include_value and not value_columns:
        warnings.append("來源檔案沒有欄名含「Value」的欄位，刪除 Value 欄的步驟沒有實際作用。")
    if "Item_Shop" not in df_input.columns:
        warnings.append("來源檔案缺少「Item_Shop」欄位，SKU Desc 與 Shop Name 留空。")
    if raw.empty:
        warnings.append("清理後沒有可轉化的資料列，請確認來源檔案與 Shop Nature 篩選。")

    if progress_callback:
        progress_callback(0.4, "完成資料轉換...")
    return raw, warnings


def generate_excel_raw_convert(raw: pd.DataFrame, progress_callback=None) -> BytesIO:
    """產生單一工作表 RAW 的 Excel（無說明、無統計工作表）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "RAW"
    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    _write_raw_sheet(
        ws,
        raw if raw is not None else pd.DataFrame(),
        progress_callback,
        progress_start=0.0,
        progress_end=0.9,
    )

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")
    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf
