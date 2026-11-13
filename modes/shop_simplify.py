# -*- coding: utf-8 -*-
"""店舖精簡SKU模式：把 By SKUxShop 明細轉化為店舖精簡 SKU 回覆表。

每列一店×一 SKU，保留來源檔全部欄位，並在最後附加六個固定欄位：

- 過去90日有無賣過：Sold Qty (P90D) > 0 → 「有」，否則（等於 0、負數或空白）→「無賣」；
- 過去90日有無存貨：Shop Stock Qty (LLM)、(LM) 或 (CM) 任一 > 0 →「有存貨」
  （LLM 欄自 2026-10-26 起來源報表提供；缺欄時以其餘庫存欄判定），
  三者同時為 0（或空白／負數）→「無存貨」；
- 是否該店的Top 80%：每間店以 Adj. SST 分類為母體，在每個分類內按 Sold Qty (P90D)
  由高至低累計，累計銷量首次達該店該分類總銷量 80% 的 SKU 為「是」
  （同一銷量整組納入或整組排除；分類之間不互相擠佔；缺 Site Code 時全表單一分組、
  缺 Adj. SST 時每店單一分組，均列 warning）；
- SCT建議、Shop 回覆新的NDRF、如果RF的Safety Qty：一律留空，供 SCT／店舖填寫。

本模式不作調貨、Safety 或退倉建議，亦不篩選列（全店全 SKU 輸出）；清理規則沿用
`_clean_dataframe`（剔除合計列與 117 - Marketing），Shop Nature 篩選為選填。
另由 Item_Shop 以 Site Code 解析 SKU Desc 與 Shop Name（SKU Desc 排於 Brand 之後、
Shop Name 排於 Site Code 之後），其餘來源欄序不變；欄名含「Value」的欄位一律刪除
（大小寫敏感，與「RAW 明細轉化 (不含 Value)」同一判定，故 Valuation Class 不受影響）。
輸出經 Brand／SKU（文字）排序，並產生 Summary（Adj. MCH1 × Item Status 計數）與
Export（回覆明細）兩個工作表；Export 套用店舖精簡回覆表範本的欄寬、隱藏欄、
數字格式、凍結窗格與顏色標示——其中 Item Status 為「Discontinued」、RP Type 為
「ND」、過去90日有無存貨為「無存貨」三類值以紅色底標示（比對忽略大小寫與
前後空白；RP Type 的紅色底會蓋過該欄的綠色），兩個模式共用同一匯出引擎。

另有「指定店舖」變體（`transform_shop_simplify_site`）：只輸出指定店舖（如 HA21）
的 SKU 列，同時輸出件數與金額兩組 Top N% 判定（共四欄；N 可選 80%（預設）或 70%）——

- 件數：是否{店舖}的Top N%（件數）／是否同Shop Nature的Top N%（件數），按 Sold Qty (P90D)；
- 金額：是否{店舖}的Top N%（金額）／是否同Shop Nature的Top N%（金額），按 Net Sales (P{N}D)＝
  Net Sales (M-2)＋(M-1)＋(CM - MTD)（負數視為 0），N 由數據截止日（T-1）的 MTD 日數＋
  兩個整月日數決定。

四欄均以 Adj. SST 分類為母體，分類之間不互相擠佔；來源缺 Adj. SST 時退回
全店／全 Nature 單一分組並列 warning，缺三個月 Net Sales 欄時只輸出件數兩欄。
欄名隨比例與視窗改變（如 Top 80%／Net Sales (P73D)），故切換重跑不會留下舊欄。
原模式（多店輸出）只有件數一種口徑，欄名維持「是否該店的Top 80%」不加後綴。

四個判定欄只加在指定店舖的列上，其餘計算、清理、解析與回覆欄留空規則與原模式相同。
"""
import calendar
import datetime
import gc
import re
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from modes.common import (
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
    "_SHOP_SIMPLIFY_EXPORT_SHEET",
    "_SHOP_SIMPLIFY_SUMMARY_SHEET",
    "_SHOP_SIMPLIFY_TOP80_SHARE",
    "_SHOP_SIMPLIFY_TOP_RATIO_OPTIONS",
    "_SHOP_SIMPLIFY_NET_SALES_COLUMNS",
    "_SHOP_SIMPLIFY_VALUE_SUFFIX",
    "_SHOP_SIMPLIFY_SITE_TOP80_TEMPLATE",
    "_SHOP_SIMPLIFY_NATURE_TOP80_TEMPLATE",
    "_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN",
    "_SHOP_SIMPLIFY_TOP80_YES",
    "_SHOP_SIMPLIFY_TOP80_NO",
    "_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN",
    "_SHOP_SIMPLIFY_RED",
    "_SHOP_SIMPLIFY_RED_RULES",
    "_shop_simplify_site_top80_column",
    "_shop_simplify_nature_top_column",
    "_shop_simplify_value_site_top_column",
    "_shop_simplify_value_nature_top_column",
    "_shop_simplify_net_sales_column",
    "_shop_simplify_net_sales_window",
    "_shop_simplify_net_sales_series",
    "_is_shop_simplify_value_top_column",
    "_shop_simplify_top_share",
    "_shop_simplify_ratio_percent",
    "_shop_simplify_top80_mask",
    "_shop_simplify_top80_keys",
    "shop_simplify_stats",
    "shop_simplify_site_stats",
    "transform_shop_simplify",
    "transform_shop_simplify_site",
    "generate_excel_shop_simplify",
]


# ── 輸出欄位契約（欄名固定，測試與匯出檔依賴，未經明確要求不得改動）──
_SHOP_SIMPLIFY_SOLD_COLUMN = "過去90日有無賣過"
_SHOP_SIMPLIFY_STOCK_COLUMN = "過去90日有無存貨"
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

# ── 指定店舖模式的 Top N% 判定（預設 80%，介面可選 70%）──
# 累計銷量達該母體（指定店舖／同 Shop Nature 全部店舖）總銷量的 N% 即屬 Top N%；
# 同一銷量的 SKU 整組納入或整組排除，避免同銷量被隨意切開。
_SHOP_SIMPLIFY_TOP80_SHARE = 0.8
# 介面可選的比例（0.8＝Top 80%、0.7＝Top 70%）；原模式（多店輸出）固定 80%
_SHOP_SIMPLIFY_TOP_RATIO_OPTIONS = (0.8, 0.7)
_SHOP_SIMPLIFY_SITE_TOP80_TEMPLATE = "是否{site}的Top {pct}%"
_SHOP_SIMPLIFY_NATURE_TOP80_TEMPLATE = "是否同Shop Nature的Top {pct}%"
# 指定店舖變體同時輸出件數與金額兩種口徑，故四欄一律加註口徑，避免同名混淆
_SHOP_SIMPLIFY_QTY_SUFFIX = "（件數）"
_SHOP_SIMPLIFY_VALUE_SUFFIX = "（金額）"
_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN = (
    _SHOP_SIMPLIFY_NATURE_TOP80_TEMPLATE.format(pct=80) + _SHOP_SIMPLIFY_QTY_SUFFIX
)
_SHOP_SIMPLIFY_TOP80_YES = "是"
_SHOP_SIMPLIFY_TOP80_NO = "否"
# 金額口徑：來源只有月欄（M-2／M-1／CM-MTD），三者相加即「MTD 日數＋兩個整月」的實際覆蓋日數
# （例如數據截止 11 月 12 日 → 12＋31＋30＝73 日，欄名標示 Net Sales (P73D)）。
_SHOP_SIMPLIFY_NET_SALES_COLUMNS = ("Net Sales (M-2)", "Net Sales (M-1)", "Net Sales (CM - MTD)")
# 原模式（多店輸出）每店各自的 Top 80% 欄名；指定店舖變體改用「是否{店舖}的Top 80%」
_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN = "是否該店的Top 80%"


# ── Excel 輸出格式（對齊店舖精簡回覆表範本）──
_SHOP_SIMPLIFY_EXPORT_SHEET = "Export"
_SHOP_SIMPLIFY_SUMMARY_SHEET = "Summary"

# 固定表頭顏色（使用者範本）
_SHOP_SIMPLIFY_HEADER_BLUE = "D9E2F3"
_SHOP_SIMPLIFY_GREEN = "CCFF99"
_SHOP_SIMPLIFY_PURPLE = "E4DFEC"
_SHOP_SIMPLIFY_NO_SALE_PINK = "F2DCDB"

# 紅色底（醒目紅底；比「無賣」的淡紅 F2DCDB 明顯，故兩者不會混淆）
_SHOP_SIMPLIFY_RED = "FF9999"
# 紅色底判定：欄名 → 觸發值（比對前去除前後空白、忽略大小寫；欄位不存在時略過該條件）
_SHOP_SIMPLIFY_RED_RULES = (
    ("Item Status", "discontinued"),
    ("RP Type", "nd"),
    (_SHOP_SIMPLIFY_STOCK_COLUMN, _SHOP_SIMPLIFY_STOCK_NO.lower()),
)

# 綠色表頭（另加綠色資料格：RP Type／Shop Stock Qty (CM)）
_SHOP_SIMPLIFY_GREEN_HEADERS = {
    "RP Type", "Shop Stock Qty (CM)", _SHOP_SIMPLIFY_NDRF_COLUMN, _SHOP_SIMPLIFY_SAFETY_COLUMN,
}
_SHOP_SIMPLIFY_GREEN_DATA_COLUMNS = {"RP Type", "Shop Stock Qty (CM)"}
# 紫色表頭與資料格（Sold Qty (P90D)）
_SHOP_SIMPLIFY_PURPLE_COLUMNS = {"Sold Qty (P90D)"}

# 隱藏欄（使用者範本；按欄名對應，來源缺欄時自動略過）
_SHOP_SIMPLIFY_HIDDEN_COLUMNS = {
    "Item_Shop", "Valuation Class", "MCH2", "MC", "Shop Nature", "Shop Type",
    "D001 Stock Qty (LM)", "D001 Blocked Qty (CM)", "D001 QI Qty (CM)",
    "Turnday (P90D Qty)", "Net Sales (M-2)", "Excess Stock Qty",
    "D001 On Order PO Qty (CM)", "Shop On Order PO Qty (CM)", "Shop Pending Receive (±7 days)",
}

# 欄寬（使用者範本；未列出的欄沿用 Excel 預設寬度）
_SHOP_SIMPLIFY_COLUMN_WIDTHS = {
    "Item_Shop": 79.7109375,
    "SKU": 13.140625,
    "Valuation Class": 25.0,
    "Brand": 8.7109375,
    "SKU Desc": 46.42578125,
    "Buyer": 18.85546875,
    "Adj. SST": 29.42578125,
    "Vendor MC": 23.140625,
    "Adj. MCH1": 17.28515625,
    "MCH2": 48.28515625,
    "MC": 56.140625,
    "Item Status": 12.7109375,
    "Launch Date": 10.5703125,
    "Item Class": 10.0,
    "Supply Source": 11.140625,
    "Site Code": 10.140625,
    "Shop Name": 10.85546875,
    "OM": 8.85546875,
    "Shop Nature": 11.28515625,
    "Shop Type": 10.0,
    "RP Type": 9.85546875,
    "Adj. Safety Stock": 10.28515625,
    "Shop Monthly Sold Qty (P90D)": 13.140625,
    "Safety VS Monthly Sold Qty": 10.0,
    "D001 Stock Qty (LM)": 10.28515625,
    "D001 Net Stock Qty (CM)": 9.7109375,
    "D001 Blocked Qty (CM)": 11.42578125,
    "D001 QI Qty (CM)": 10.0,
    "Shop Stock Qty (LM)": 9.0,
    "Shop Stock Qty (CM)": 9.0,
    "Sold Qty (P90D)": 10.7109375,
    "Turnday (P90D Qty)": 12.0,
    "Sold Qty (M-2)": 10.5703125,
    "Sold Qty (M-1)": 10.5703125,
    "Sold Qty (CM - MTD)": 10.42578125,
    "Net Sales (M-2)": 11.0,
    "Net Sales (CM - MTD)": 10.42578125,
    "Shop Stock Status": 56.85546875,
    "Excess Stock Qty": 10.42578125,
    "D001 On Order PO Qty (CM)": 12.85546875,
    "Shop On Order PO Qty (CM)": 13.0,
    "Shop Pending Receive (±7 days)": 13.42578125,
    _SHOP_SIMPLIFY_SOLD_COLUMN: 10.85546875,
    _SHOP_SIMPLIFY_STOCK_COLUMN: 10.85546875,
    _SHOP_SIMPLIFY_NATURE_TOP80_COLUMN: 26.0,
    _SHOP_SIMPLIFY_NDRF_COLUMN: 8.42578125,
    _SHOP_SIMPLIFY_SAFETY_COLUMN: 9.7109375,
}

# 數字格式（使用者範本）
_SHOP_SIMPLIFY_NUMBER_FORMATS = {
    "Shop Monthly Sold Qty (P90D)": '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)',
    "Safety VS Monthly Sold Qty": "0%",
    "Turnday (P90D Qty)": "#,##0_);[Red](#,##0)",
    "Sold Qty (M-2)": "#,##0_);[Red](#,##0)",
    "Sold Qty (M-1)": "#,##0_);[Red](#,##0)",
    "Sold Qty (CM - MTD)": "#,##0_);[Red](#,##0)",
    "Net Sales (M-2)": "#,##0_);[Red](#,##0)",
    "Net Sales (M-1)": "#,##0_);[Red](#,##0)",
    "Net Sales (CM - MTD)": "#,##0_);[Red](#,##0)",
}


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


def shop_simplify_site_stats(df: pd.DataFrame, site_code: str, top_ratio=None) -> dict:
    """指定店舖模式的結果頁指標：列數、SKU 數、Adj. SST 分類數、件數與金額兩組 Top N% 命中數、兩項皆非、無賣／無存貨。

    `top_ratio` 只在判定欄不存在時作後備；正常情況由欄名（「…Top 70%」）取回實際比例，
    故同一份結果不論介面當前選哪個比例，統計都反映該檔實際使用的比例。
    金額口徑只在結果檔含「（金額）」判定欄時提供，否則 `value_available` 為 False。
    """
    site = str(site_code or "").strip()
    stats = {
        "rows": 0, "skus": 0, "sst_groups": 0, "sold_no": 0, "stock_no": 0,
        "site_top80": 0, "nature_top80": 0, "neither": 0,
        "value_site_top80": 0, "value_nature_top80": 0, "value_neither": 0,
        "value_available": False, "window": "",
        "site": site, "nature": "", "ratio": _shop_simplify_ratio_percent(top_ratio),
    }
    if df is None or len(df) == 0:
        return stats

    def _flag(column: str) -> pd.Series:
        if column not in df.columns:
            return pd.Series("", index=df.index, dtype=object)
        return df[column].fillna("").astype(str).str.strip()

    sold_no = _flag(_SHOP_SIMPLIFY_SOLD_COLUMN) == _SHOP_SIMPLIFY_SOLD_NO
    stock_no = _flag(_SHOP_SIMPLIFY_STOCK_COLUMN) == _SHOP_SIMPLIFY_STOCK_NO
    nature_columns = [
        c for c in df.columns
        if isinstance(c, str) and c.strip().startswith("是否同Shop Nature的Top ")
    ]
    nature_qty_columns = [c for c in nature_columns if not _is_shop_simplify_value_top_column(c)]
    nature_value_columns = [c for c in nature_columns if _is_shop_simplify_value_top_column(c)]
    site_qty_columns = [
        c for c in df.columns
        if _is_shop_simplify_top80_column(c)
        and c not in nature_columns
        and not _is_shop_simplify_value_top_column(c)
    ]
    site_value_columns = [
        c for c in df.columns
        if _is_shop_simplify_value_top_column(c) and c not in nature_columns
    ]
    site_column = site_qty_columns[0] if len(site_qty_columns) == 1 else _shop_simplify_site_top80_column(site, top_ratio)
    nature_column = nature_qty_columns[0] if len(nature_qty_columns) == 1 else _shop_simplify_nature_top_column(top_ratio)
    site_flag = _flag(site_column)
    nature_flag = _flag(nature_column)
    stats["ratio"] = (
        _shop_simplify_top_column_percent(site_column)
        or _shop_simplify_top_column_percent(nature_column)
        or _shop_simplify_ratio_percent(top_ratio)
    )

    stats["rows"] = int(len(df))
    stats["sold_no"] = int(sold_no.sum())
    stats["stock_no"] = int(stock_no.sum())
    stats["site_top80"] = int((site_flag == _SHOP_SIMPLIFY_TOP80_YES).sum())
    stats["nature_top80"] = int((nature_flag == _SHOP_SIMPLIFY_TOP80_YES).sum())
    stats["neither"] = int(((site_flag == _SHOP_SIMPLIFY_TOP80_NO) & (nature_flag == _SHOP_SIMPLIFY_TOP80_NO)).sum())

    if site_value_columns and nature_value_columns:
        value_site_flag = _flag(site_value_columns[0])
        value_nature_flag = _flag(nature_value_columns[0])
        stats["value_available"] = True
        stats["value_site_top80"] = int((value_site_flag == _SHOP_SIMPLIFY_TOP80_YES).sum())
        stats["value_nature_top80"] = int((value_nature_flag == _SHOP_SIMPLIFY_TOP80_YES).sum())
        stats["value_neither"] = int(
            ((value_site_flag == _SHOP_SIMPLIFY_TOP80_NO) & (value_nature_flag == _SHOP_SIMPLIFY_TOP80_NO)).sum()
        )
    for c in df.columns:
        if isinstance(c, str) and c.strip().startswith("Net Sales (P") and c.strip().endswith("D)"):
            stats["window"] = c.strip()[len("Net Sales ("):-1]
            break

    if "SKU" in df.columns:
        skus = df["SKU"].dropna().astype(str).str.strip()
        stats["skus"] = int(skus[skus != ""].nunique())
    if "Adj. SST" in df.columns:
        sst = df["Adj. SST"].dropna().astype(str).str.strip()
        stats["sst_groups"] = int(sst[sst != ""].nunique())
    if "Shop Nature" in df.columns:
        natures = df["Shop Nature"].dropna().astype(str).str.strip()
        natures = natures[natures != ""]
        if not natures.empty:
            stats["nature"] = str(natures.mode().iloc[0])
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


def _shop_simplify_launch_int(value):
    """Launch Date 統一為 YYYYMMDD 整數（來源多為 float 如 20120520.0）；
    日期物件、YYYY-MM-DD、YYYY/MM/DD 皆可解析；無法解析或空白回傳 NaN。"""
    if value is None:
        return np.nan
    try:
        if pd.isna(value):
            return np.nan
    except (TypeError, ValueError):
        pass
    if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
        return int(value.strftime("%Y%m%d"))
    text = re.sub(r"\.0+$", "", str(value).strip())
    text = text.replace("-", "").replace("/", "")
    if len(text) == 8 and text.isdigit():
        return int(text)
    return np.nan


def _shop_simplify_top_share(ratio=None) -> float:
    """正規化 Top 比例：接受 0.7／70／"70%"；必須在 (0, 1] 之內，否則拋 ValueError。

    未提供時回傳預設 0.8（Top 80%），故舊呼叫端的行為與欄名完全不變。
    """
    if ratio is None:
        return _SHOP_SIMPLIFY_TOP80_SHARE
    value = ratio
    if isinstance(value, str):
        text = value.strip().replace("%", "")
        try:
            value = float(text)
        except (TypeError, ValueError):
            raise ValueError(
                f"Top 比例「{ratio}」無法識別，請提供 0 至 1 的比例（如 0.7）或百分比（如 70）。"
            )
    try:
        value = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"Top 比例「{ratio!r}」不是數值，請提供 0 至 1 的比例（如 0.7）或百分比（如 70）。")
    if value > 1:
        value = value / 100.0
    if not 0 < value <= 1:
        raise ValueError(f"Top 比例必須大於 0 且不超過 100%（收到 {ratio!r}）。")
    return value


def _shop_simplify_ratio_percent(ratio=None) -> int:
    """比例轉整數百分比（0.7 → 70），用於欄名與文案。"""
    return int(round(_shop_simplify_top_share(ratio) * 100))


def _shop_simplify_top_column_percent(name) -> int | None:
    """從欄名取回百分比（「是否HA21的Top 70%」→ 70）；不是 Top 欄時回傳 None。"""
    if not isinstance(name, str):
        return None
    match = re.search(r"Top\s*(\d{1,3})%", name.strip())
    return int(match.group(1)) if match else None


def _shop_simplify_net_sales_window(data_date=None) -> dict:
    """回傳金額母體的日數拆解：MTD 日數＋M-1 月日數＋M-2 月日數。

    來源數據為 T-1，故未指定時以「今天減一日」為數據截止日。例如 11 月 13 日執行 →
    截止 11 月 12 日：MTD 12 日＋10 月 31 日＋9 月 30 日＝73 日，標示 `P73D`。
    """
    if data_date is None:
        data_date = datetime.date.today() - datetime.timedelta(days=1)
    if isinstance(data_date, datetime.datetime):
        data_date = data_date.date()
    if not isinstance(data_date, datetime.date):
        raise ValueError(f"數據截止日「{data_date!r}」不是日期，無法計算 Net Sales 視窗。")
    mtd_days = int(data_date.day)
    m1_end = data_date.replace(day=1) - datetime.timedelta(days=1)
    m1_days = int(calendar.monthrange(m1_end.year, m1_end.month)[1])
    m2_end = m1_end.replace(day=1) - datetime.timedelta(days=1)
    m2_days = int(calendar.monthrange(m2_end.year, m2_end.month)[1])
    total = mtd_days + m1_days + m2_days
    return {
        "data_date": data_date,
        "mtd_days": mtd_days,
        "m1_days": m1_days,
        "m2_days": m2_days,
        "m1": m1_end.strftime("%Y-%m"),
        "m2": m2_end.strftime("%Y-%m"),
        "total_days": total,
        "label": f"P{total}D",
    }


def _shop_simplify_site_top_base(site_code: str, ratio=None) -> str:
    """店舖層 Top N% 的基礎欄名（不含口徑後綴）：HA21、80% →「是否HA21的Top 80%」。"""
    return _SHOP_SIMPLIFY_SITE_TOP80_TEMPLATE.format(
        site=str(site_code or "").strip(),
        pct=_shop_simplify_ratio_percent(ratio),
    )


def _shop_simplify_nature_top_base(ratio=None) -> str:
    """同 Shop Nature 的基礎欄名（不含口徑後綴）：80% →「是否同Shop Nature的Top 80%」。"""
    return _SHOP_SIMPLIFY_NATURE_TOP80_TEMPLATE.format(pct=_shop_simplify_ratio_percent(ratio))


def _shop_simplify_value_site_top_column(site_code: str, ratio=None) -> str:
    """金額口徑的店舖層欄名（80% →「是否HA21的Top 80%（金額）」）。"""
    return _shop_simplify_site_top_base(site_code, ratio) + _SHOP_SIMPLIFY_VALUE_SUFFIX


def _shop_simplify_value_nature_top_column(ratio=None) -> str:
    """金額口徑的同 Nature 欄名（80% →「是否同Shop Nature的Top 80%（金額）」）。"""
    return _shop_simplify_nature_top_base(ratio) + _SHOP_SIMPLIFY_VALUE_SUFFIX


def _shop_simplify_net_sales_column(label: str) -> str:
    """金額資料欄名（label＝P73D →「Net Sales (P73D)」）。"""
    return f"Net Sales ({str(label or '').strip()})"


def _is_shop_simplify_value_top_column(name) -> bool:
    """是否金額口徑的 Top 判定欄（欄名以「（金額）」結尾）。"""
    return isinstance(name, str) and _is_shop_simplify_top80_column(name) and name.strip().endswith(
        _SHOP_SIMPLIFY_VALUE_SUFFIX
    )


def _shop_simplify_site_top80_column(site_code: str, ratio=None) -> str:
    """指定店舖模式的件數店舖層欄名（HA21、80% →「是否HA21的Top 80%（件數）」）。"""
    return _shop_simplify_site_top_base(site_code, ratio) + _SHOP_SIMPLIFY_QTY_SUFFIX


def _shop_simplify_nature_top_column(ratio=None) -> str:
    """件數口徑的同 Shop Nature 欄名（80% →「是否同Shop Nature的Top 80%（件數）」）。"""
    return _shop_simplify_nature_top_base(ratio) + _SHOP_SIMPLIFY_QTY_SUFFIX


def _is_shop_simplify_top80_column(name) -> bool:
    """是否本模式的 Top N% 判定欄（含指定店舖動態欄名、Nature 欄、70% 與金額變體）。"""
    if not isinstance(name, str):
        return False
    text = name.strip()
    return text.startswith("是否") and re.search(r"Top\s*\d{1,3}%", text) is not None


def _shop_simplify_top80_mask(qty: pd.Series, share: float = _SHOP_SIMPLIFY_TOP80_SHARE) -> pd.Series:
    """按銷量由高至低累計，回傳「累計首次達總銷量 N%」的 SKU 布林標記。

    - 以索引為 SKU 鍵（呼叫方先按 SKU 彙總）；
    - `share` 為累計比例（0.8＝Top 80%、0.7＝Top 70%）；
    - 同一銷量的 SKU 整組納入或整組排除，避免同銷量被任意切開；
    - 總銷量為 0（或全為空白／負數）時全部回傳 False。
    """
    share = _shop_simplify_top_share(share)
    values = pd.to_numeric(qty, errors="coerce").fillna(0.0).clip(lower=0.0)
    mask = pd.Series(False, index=values.index, dtype=bool)
    if values.empty:
        return mask
    total = float(values.sum())
    if total <= 0:
        return mask
    ordered = values.sort_values(ascending=False, kind="stable")
    cumulative = ordered.cumsum().to_numpy(dtype=float)
    reached = np.flatnonzero(cumulative >= total * share)
    if reached.size == 0:
        return mask
    cutoff = float(ordered.iloc[int(reached[0])])
    return values >= cutoff


def _shop_simplify_top80_keys(
    groups: pd.Series,
    skus: pd.Series,
    qty: pd.Series,
    share: float = _SHOP_SIMPLIFY_TOP80_SHARE,
) -> set[tuple[str, str]]:
    """按分類（Adj. SST）分組，回傳各組累計 N% 的 (分類, SKU) 集合。

    每個分類獨立累計，分類之間不互相擠佔；同一 SKU 在同一分類內重複出現時先加總。
    """
    frame = pd.DataFrame({
        "group": groups.fillna("").astype(str).str.strip(),
        "sku": skus.fillna("").astype(str).str.strip(),
        "qty": pd.to_numeric(qty, errors="coerce").fillna(0.0).clip(lower=0.0),
    })
    if frame.empty:
        return set()
    totals = frame.groupby(["group", "sku"], sort=False)["qty"].sum()
    keys: set[tuple[str, str]] = set()
    for group, series in totals.groupby(level=0, sort=False):
        mask = _shop_simplify_top80_mask(series.droplevel(0), share=share)
        keys.update((str(group), str(sku)) for sku in mask.index[mask.to_numpy()])
    return keys


def _shop_simplify_net_sales_series(df: pd.DataFrame, net_columns) -> pd.Series:
    """三個月 Net Sales 相加成金額母體；負數（退貨）先 clip 為 0，缺值當 0。

    `net_columns` 為 [(實際欄名, 標準欄名), ...]；實際欄名為 None 時略過該月。
    """
    total = pd.Series(0.0, index=df.index, dtype=float)
    for actual, _standard in net_columns:
        if actual is not None and actual in df.columns:
            total = total.add(_safe_numeric(df[actual]).clip(lower=0.0), fill_value=0.0)
    return total.fillna(0.0)


def _shop_simplify_shop_top80_flags(
    df: pd.DataFrame,
    sold: pd.Series,
    site_col: str | None,
    sst_col: str | None,
    share: float = _SHOP_SIMPLIFY_TOP80_SHARE,
) -> pd.Series:
    """每列是否屬「該店 × 該 Adj. SST 分類」累計 N% 內的 SKU（原模式多店輸出用）。

    以 (Site Code, Adj. SST) 為母體鍵，母體之間互不擠佔；缺 Site Code 時全表視為
    單一店舖（"(全部店舖)"），缺 Adj. SST 時每店單一分組（"(全部)"）。
    """
    index = df.index
    if site_col is not None and site_col in df.columns:
        sites = df[site_col].fillna("").astype(str).str.strip()
        sites = sites.mask(sites == "", "(空白店舖)")
    else:
        sites = pd.Series("(全部店舖)", index=index, dtype=object)
    if sst_col is not None and sst_col in df.columns:
        groups = df[sst_col].fillna("").astype(str).str.strip()
        groups = groups.mask(groups == "", "(空白分類)")
    else:
        groups = pd.Series("(全部)", index=index, dtype=object)
    if "SKU" not in df.columns:
        return pd.Series(False, index=index, dtype=bool)

    # 以不可見分隔符併成單一母體鍵，重用按分類累計 80% 的同一判定
    body_keys = sites.astype(str) + "\u0000" + groups.astype(str)
    skus = _cached_safe_sku_series(df["SKU"]).astype(str).str.strip()
    keys = _shop_simplify_top80_keys(body_keys, skus, sold, share=share)
    return pd.Series(
        [(group, sku) in keys for group, sku in zip(body_keys, skus)],
        index=index,
        dtype=bool,
    )


def _shop_simplify_drop_existing_columns(df: pd.DataFrame) -> pd.DataFrame:
    """移除來源檔已有的輸出欄（含前後空格變體與指定店舖的 Top 80% 欄），避免重複欄位。

    回傳移除後的 df；是否有欄被移除由呼叫方比對欄數自行判斷。
    """
    avail = set(df.columns)
    drop = []
    for name in [_SHOP_SIMPLIFY_SOLD_COLUMN, _SHOP_SIMPLIFY_STOCK_COLUMN, *_SHOP_SIMPLIFY_REPLY_COLUMNS]:
        resolved = _resolve_column_name(name, avail)
        if resolved is not None:
            drop.append(resolved)
    drop.extend(c for c in df.columns if _is_shop_simplify_top80_column(c) and c not in drop)
    # 舊結果的 Net Sales (PxxD) 屬衍生欄，重跑時一併移除，避免與本次視窗並存
    drop.extend(
        c for c in df.columns
        if isinstance(c, str) and re.fullmatch(r"Net Sales \(P\d+D\)", c.strip()) and c not in drop
    )
    return df.drop(columns=drop) if drop else df


def transform_shop_simplify(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
    include_shop_top80: bool = True,
):
    """清理來源資料並附加「過去90日有無賣過／過去90日有無存貨／是否該店的Top 80%」
    與三個空白回覆欄。

    - 「是否該店的Top 80%」以 (Site Code, Adj. SST) 為母體，每間店每個分類各自一條
      80% 線；指定店舖變體另有自己的店舖／同 Nature 判定，故以 `include_shop_top80=False`
      關閉（避免重複欄位）。
    - 回傳 (raw_df, warnings)。缺少 Sold Qty (P90D)，或 Shop Stock Qty (LLM)／(LM)／(CM)
      三個庫存欄同時缺失時拋出 ValueError（無法判定，不應靜默輸出）。
    """
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必需欄位: {', '.join(missing)}")

    avail = set(df_input.columns)
    sold_col = _resolve_column_name("Sold Qty (P90D)", avail)
    llm_col = _resolve_column_name("Shop Stock Qty (LLM)", avail)
    lm_col = _resolve_column_name("Shop Stock Qty (LM)", avail)
    cm_col = _resolve_column_name("Shop Stock Qty (CM)", avail)
    site_col = _resolve_column_name("Site Code", avail)
    sst_col = _resolve_column_name("Adj. SST", avail)
    if sold_col is None:
        raise ValueError("來源檔案缺少「Sold Qty (P90D)」，無法判定過去90日有無賣過。")
    _stock_pairs = [
        ("Shop Stock Qty (LLM)", llm_col),
        ("Shop Stock Qty (LM)", lm_col),
        ("Shop Stock Qty (CM)", cm_col),
    ]
    if all(c is None for _, c in _stock_pairs):
        raise ValueError(
            "來源檔案同時缺少「Shop Stock Qty (LLM)」「Shop Stock Qty (LM)」「Shop Stock Qty (CM)」，"
            "無法判定過去90日有無存貨。"
        )

    if progress_callback:
        progress_callback(0.15, "清理資料中...")
    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback).copy()

    warnings: list[str] = []
    _used_stock_names = [n for n, c in _stock_pairs if c is not None]
    for _name, _col in _stock_pairs:
        if _col is None:
            warnings.append(
                f"來源檔案缺少「{_name}」，過去90日有無存貨以「{'」、「'.join(_used_stock_names)}」判定。"
            )

    had_output_columns = any(
        _resolve_column_name(name, avail) is not None
        for name in [_SHOP_SIMPLIFY_SOLD_COLUMN, _SHOP_SIMPLIFY_STOCK_COLUMN, *_SHOP_SIMPLIFY_REPLY_COLUMNS]
    )
    if had_output_columns:
        warnings.append(
            "來源檔已含本模式的輸出欄（過去90日有無賣過／過去90日有無存貨／"
            "SCT建議／Shop 回覆新的NDRF／如果RF的Safety Qty／Top 80% 判定欄），"
            "將以本次計算結果覆蓋，三個回覆欄一律留空。"
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
        progress_callback(0.55, "判定過去90日有無存貨...")
    llm_stock = (
        _safe_numeric(df[llm_col]) if llm_col is not None and llm_col in df.columns else pd.Series(0.0, index=df.index)
    )
    lm_stock = (
        _safe_numeric(df[lm_col]) if lm_col is not None and lm_col in df.columns else pd.Series(0.0, index=df.index)
    )
    cm_stock = (
        _safe_numeric(df[cm_col]) if cm_col is not None and cm_col in df.columns else pd.Series(0.0, index=df.index)
    )
    has_stock = (
        (llm_stock.to_numpy(dtype=float) > 0)
        | (lm_stock.to_numpy(dtype=float) > 0)
        | (cm_stock.to_numpy(dtype=float) > 0)
    )
    stock_status = np.where(has_stock, _SHOP_SIMPLIFY_STOCK_YES, _SHOP_SIMPLIFY_STOCK_NO)

    df[_SHOP_SIMPLIFY_SOLD_COLUMN] = sold_status
    df[_SHOP_SIMPLIFY_STOCK_COLUMN] = stock_status

    if include_shop_top80:
        if progress_callback:
            progress_callback(0.62, "判定各店的Top 80%...")
        if site_col is None:
            warnings.append(
                f"來源檔案缺少「Site Code」，「{_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN}」未能按店舖分組，"
                "改以全表為單一店舖計算。"
            )
        if sst_col is None:
            warnings.append(
                f"來源檔案缺少「Adj. SST」，「{_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN}」未能按 Adj. SST 分類，"
                "改以每店單一分組計算。"
            )
        shop_top80 = _shop_simplify_shop_top80_flags(df, sold, site_col, sst_col)
        df[_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN] = np.where(
            shop_top80.to_numpy(dtype=bool),
            _SHOP_SIMPLIFY_TOP80_YES,
            _SHOP_SIMPLIFY_TOP80_NO,
        )

    for name in _SHOP_SIMPLIFY_REPLY_COLUMNS:
        df[name] = ""

    # SKU 統一為文字（與其他模式一致，避免 Excel 以科學記號顯示長編號）
    if "SKU" in df.columns:
        df["SKU"] = _cached_safe_sku_series(df["SKU"])

    # Launch Date 統一為 YYYYMMDD 整數（Excel 顯示為數字，與範本一致）
    if "Launch Date" in df.columns:
        df["Launch Date"] = pd.Series(
            [_shop_simplify_launch_int(v) for v in df["Launch Date"]],
            index=df.index,
            dtype=object,
        )

    # 非有限數值（如 0 銷售時 Turnday 的 inf）一律留空，避免 Excel 出現 inf
    for col in df.columns:
        if pd.api.types.is_float_dtype(df[col]):
            arr = df[col].to_numpy(dtype=float)
            if np.isinf(arr).any():
                df[col] = df[col].where(np.isfinite(arr))

    # 排序：Brand、SKU（文字升序），與店舖精簡回覆表範本一致
    if "Brand" in df.columns and "SKU" in df.columns:
        df = df.sort_values(
            ["Brand", "SKU"],
            key=lambda s: s.astype(str).str.strip(),
            kind="stable",
        ).reset_index(drop=True)

    if df.empty:
        warnings.append("清理後沒有可輸出的資料列，請確認來源檔案與 Shop Nature 篩選。")

    if progress_callback:
        progress_callback(0.7, "完成資料轉換...")
    return df, warnings


def transform_shop_simplify_site(
    df_input: pd.DataFrame,
    site_code: str,
    progress_callback=None,
    top_ratio=None,
    include_net_sales: bool = True,
    data_date=None,
):
    """指定店舖模式：只輸出該店舖（Site Code）的 SKU 列，並加兩組（件數／金額）Top N% 判定欄。

    - 件數：「是否{店舖}的Top N%（件數）」與「是否同Shop Nature的Top N%（件數）」，按 Sold Qty (P90D)；
    - 金額：「是否{店舖}的Top N%（金額）」與「是否同Shop Nature的Top N%（金額）」，
      按 Net Sales (P{N}D)＝`Net Sales (M-2)`＋`Net Sales (M-1)`＋`Net Sales (CM - MTD)`
      （負數先 clip 為 0），N 由 MTD 日數＋兩個整月日數決定（例如 11 月 12 日 → 73 日）。

    兩種口徑都按 Adj. SST 分類，在每個分類內累計至該分類 N%，同一銷量／金額整組納入或排除；
    分類之間不互相擠佔。來源缺 Shop Nature 時 Nature 欄留空、缺 Adj. SST 時退回單一分組，
    缺三個月 Net Sales 欄時只輸出件數判定，均列 warning。

    `top_ratio` 為累計比例：0.8（預設）或 0.7；亦接受 70、"70%" 等寫法。欄名隨比例改變，
    故切換比例重跑不會留下舊欄。`data_date` 為數據截止日（T-1），未提供時用「今天減一日」；
    金額欄名會帶實際覆蓋日數（如 `Net Sales (P73D)`）。

    其餘清理、Item_Shop 解析、刪除 Value 欄、三個回覆欄留空與排序規則完全沿用
    `transform_shop_simplify`（以 `include_shop_top80=False` 關閉原模式的
    「是否該店的Top 80%」）。回傳 (raw_df, warnings)。
    """
    site = str(site_code or "").strip()
    if not site:
        raise ValueError("未提供指定店舖（Site Code），無法執行指定店舖模式。")
    share = _shop_simplify_top_share(top_ratio)
    pct = _shop_simplify_ratio_percent(share)
    window = _shop_simplify_net_sales_window(data_date)
    net_column = _shop_simplify_net_sales_column(window["label"])

    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必需欄位: {', '.join(missing)}")

    avail = set(df_input.columns)
    site_col = _resolve_column_name("Site Code", avail)
    if site_col is None:
        raise ValueError("來源檔案缺少「Site Code」，無法指定店舖。")
    sku_col = _resolve_column_name("SKU", avail)
    sold_col = _resolve_column_name("Sold Qty (P90D)", avail)
    if sku_col is None or sold_col is None:
        raise ValueError(f"來源檔案缺少「SKU」或「Sold Qty (P90D)」，無法計算 Top {pct}%。")
    sst_col = _resolve_column_name("Adj. SST", avail)
    nature_col = _resolve_column_name("Shop Nature", avail)

    net_cols = [(_resolve_column_name(name, avail), name) for name in _SHOP_SIMPLIFY_NET_SALES_COLUMNS]
    net_missing = [name for actual, name in net_cols if actual is None]
    net_available = bool(include_net_sales) and not net_missing

    if progress_callback:
        progress_callback(0.02, "清理來源資料（全表）...")
    # 全表只做向量化清理：逐列 Item_Shop 解析留到篩選店舖後才做，
    # 80 萬行時可避免對全表逐列解析與排序，大幅節省時間與記憶體。
    df_clean = _clean_dataframe(df_input, shop_natures=None, progress_callback=progress_callback)

    warnings: list[str] = []
    sites = df_clean[site_col].fillna("").astype(str).str.strip()
    exact = sites == site
    if exact.any():
        matched_site = site
    else:
        case_insensitive = sites.str.upper() == site.upper()
        if not case_insensitive.any():
            raise ValueError(
                f"指定店舖「{site}」在清理後沒有資料列，請確認 Site Code 是否正確，"
                "或來源檔是否只包含部分店舖。"
            )
        matched_site = str(sites[case_insensitive].iloc[0]).strip()
        warnings.append(f"指定店舖「{site}」按大小寫不敏感比對為「{matched_site}」，以下按此店舖計算。")

    shop_raw = df_clean[sites == matched_site].copy()
    if shop_raw.empty:
        raise ValueError(f"指定店舖「{matched_site}」在清理後沒有資料列，無法計算 Top {pct}%。")

    site_top_column = _shop_simplify_site_top80_column(matched_site, share)
    nature_column = _shop_simplify_nature_top_column(share)
    value_site_column = _shop_simplify_value_site_top_column(matched_site, share)
    value_nature_column = _shop_simplify_value_nature_top_column(share)

    if include_net_sales and net_missing:
        warnings.append(
            f"來源檔案缺少「{'」「'.join(net_missing)}」，未能計算金額口徑 Top {pct}%，"
            "只輸出件數判定欄（不會加入 Net Sales 與（金額）欄）。"
        )

    if sst_col is None:
        warnings.append(
            f"來源檔案缺少「Adj. SST」，Top {pct}% 未能按 Adj. SST 分類，"
            "改以全店／全 Nature 單一分組計算。"
        )

    # ── 同 Shop Nature 的 Top N%（件數與金額）：先在清理後全表按 (Adj. SST, SKU) 算好，
    #    稍後映射到指定店舖的列；其他 Nature 不計入母體 ──
    nature_available = False
    nature_value = ""
    nature_keys: set[tuple[str, str]] = set()
    nature_value_keys: set[tuple[str, str]] = set()
    nature_shop_count = 0
    nature_total = 0.0
    nature_net_total = 0.0
    if nature_col is None:
        warnings.append(
            "來源檔案缺少「Shop Nature」，未能判定「"
            f"{nature_column}」，該欄留空。"
        )
    else:
        row_natures = shop_raw[nature_col].fillna("").astype(str).str.strip()
        row_natures = row_natures[row_natures != ""]
        if row_natures.empty:
            warnings.append(
                f"指定店舖「{matched_site}」的 Shop Nature 空白，未能判定「"
                f"{nature_column}」，該欄留空。"
            )
        else:
            nature_value = str(row_natures.mode().iloc[0])
            nature_available = True
            if row_natures.nunique() > 1:
                warnings.append(
                    f"指定店舖「{matched_site}」的 Shop Nature 有多個值"
                    f"（{'、'.join(sorted(row_natures.unique()))}），同 Nature 判定以「{nature_value}」計算。"
                )

            nature_series = df_clean[nature_col].fillna("").astype(str).str.strip()
            nature_df = df_clean[nature_series == nature_value]
            nature_sku = _cached_safe_sku_series(nature_df[sku_col]).astype(str).str.strip()
            nature_qty = _safe_numeric(nature_df[sold_col])
            if sst_col is None:
                nature_sst = pd.Series("(全部)", index=nature_df.index)
            else:
                nature_sst = nature_df[sst_col].fillna("").astype(str).str.strip()
                nature_sst = nature_sst.mask(nature_sst == "", "(空白)")
            nature_keys = _shop_simplify_top80_keys(nature_sst, nature_sku, nature_qty, share=share)
            if net_available:
                nature_net = _shop_simplify_net_sales_series(nature_df, net_cols)
                nature_value_keys = _shop_simplify_top80_keys(
                    nature_sst, nature_sku, nature_net, share=share
                )
                nature_net_total = float(nature_net.sum())
                del nature_net

            nature_sites = nature_df[site_col].fillna("").astype(str).str.strip()
            nature_sites = nature_sites[nature_sites != ""]
            nature_shop_count = int(nature_sites.nunique())
            nature_total = float(nature_qty.sum())
            del nature_df, nature_series, nature_sku, nature_qty, nature_sst, nature_sites

    # 全表清理結果只留到此處，及早釋放（80 萬行可省數 GB）
    del df_clean, sites
    gc.collect()

    def _site_progress(fraction, msg):
        if progress_callback:
            progress_callback(0.2 + 0.65 * float(fraction), msg)

    if progress_callback:
        progress_callback(0.2, f"整理 {matched_site} 的店舖×SKU 明細...")
    shop_df, base_warnings = transform_shop_simplify(
        shop_raw,
        progress_callback=_site_progress if progress_callback else None,
        shop_natures=None,
        include_shop_top80=False,
    )
    warnings = list(base_warnings) + warnings

    if progress_callback:
        progress_callback(0.88, f"計算 {matched_site} 的 Top {pct}%（件數＋金額）...")
    shop_sku = _cached_safe_sku_series(shop_df[sku_col]).astype(str).str.strip()
    shop_qty = _safe_numeric(shop_df[sold_col])
    if sst_col is None:
        shop_sst = pd.Series("(全部)", index=shop_df.index)
    else:
        shop_sst = shop_df[sst_col].fillna("").astype(str).str.strip()
        shop_sst = shop_sst.mask(shop_sst == "", "(空白)")

    shop_keys = _shop_simplify_top80_keys(shop_sst, shop_sku, shop_qty, share=share)
    shop_flags = pd.Series(
        [(group, sku) in shop_keys for group, sku in zip(shop_sst, shop_sku)],
        index=shop_df.index,
        dtype=bool,
    )
    nature_flags = pd.Series(
        [(group, sku) in nature_keys for group, sku in zip(shop_sst, shop_sku)]
        if nature_available else False,
        index=shop_df.index,
        dtype=bool,
    )

    if float(shop_qty.sum()) <= 0:
        warnings.append(
            f"指定店舖「{matched_site}」的 Sold Qty (P90D) 合計為 0，"
            f"「{site_top_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
        )
    if nature_available:
        if nature_shop_count <= 1:
            warnings.append(
                f"同 Shop Nature「{nature_value}」在來源檔只有 1 間店舖（即指定店舖），"
                f"「{nature_column}」與店舖 Top {pct}% 結果相同。"
            )
        if nature_total <= 0:
            warnings.append(
                f"同 Shop Nature「{nature_value}」的 Sold Qty (P90D) 合計為 0，"
                f"「{nature_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
            )

    shop_df[site_top_column] = np.where(shop_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO)
    shop_df[nature_column] = (
        np.where(nature_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO)
        if nature_available else ""
    )

    shop_net = None
    if net_available:
        shop_net = _shop_simplify_net_sales_series(shop_df, net_cols)
        shop_value_keys = _shop_simplify_top80_keys(shop_sst, shop_sku, shop_net, share=share)
        shop_value_flags = pd.Series(
            [(group, sku) in shop_value_keys for group, sku in zip(shop_sst, shop_sku)],
            index=shop_df.index,
            dtype=bool,
        )
        nature_value_flags = pd.Series(
            [(group, sku) in nature_value_keys for group, sku in zip(shop_sst, shop_sku)]
            if nature_available else False,
            index=shop_df.index,
            dtype=bool,
        )
        if float(shop_net.sum()) <= 0:
            warnings.append(
                f"指定店舖「{matched_site}」的 {net_column} 合計為 0，"
                f"「{value_site_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
            )
        if nature_available and nature_net_total <= 0:
            warnings.append(
                f"同 Shop Nature「{nature_value}」的 {net_column} 合計為 0，"
                f"「{value_nature_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
            )

        shop_df[net_column] = shop_net.to_numpy(dtype=float)
        shop_df[value_site_column] = np.where(
            shop_value_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO
        )
        shop_df[value_nature_column] = (
            np.where(nature_value_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO)
            if nature_available else ""
        )

    # 判定欄緊接「過去90日有無存貨」（件數在前、金額在後），三個回覆欄維持最後；
    # Net Sales (PxxD) 緊接 Sold Qty (P90D)，方便與件數對照
    cols = list(shop_df.columns)
    insert_at = cols.index(_SHOP_SIMPLIFY_STOCK_COLUMN) + 1
    flag_columns = [site_top_column, nature_column]
    if net_available:
        flag_columns += [value_site_column, value_nature_column]
    for name in flag_columns:
        cols.remove(name)
        cols.insert(insert_at, name)
        insert_at += 1
    if net_available:
        cols.remove(net_column)
        cols.insert(cols.index(sold_col) + 1, net_column)
    shop_df = shop_df[cols].reset_index(drop=True)

    if progress_callback:
        progress_callback(0.85, f"完成 {matched_site} 的 Top {pct}% 判定...")
    return shop_df, warnings


def _shop_simplify_cell(value):
    """Excel 儲存格值正規化：空白字串與各類缺值一律寫成 None（真正空白）；
    浮點數保留完整精度（不四捨五入），inf／-inf 視為空白。"""
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
    if isinstance(value, (float, np.floating)):
        number = float(value)
        return number if np.isfinite(number) else None
    return value


def _shop_simplify_red_mask(series: pd.Series, target: str) -> np.ndarray:
    """逐列判定是否套紅色底（去除前後空白、忽略大小寫；空白／缺值不命中）。"""
    text = series.fillna("").astype(str).str.strip().str.lower()
    return text.eq(target).to_numpy(dtype=bool)


def _write_shop_simplify_summary(ws, df: pd.DataFrame) -> None:
    """Summary 工作表：Adj. MCH1 × Item Status 的 SKU 列數樞紐（含總計；零值留空）。"""
    ws["A3"] = "計數 - SKU"
    ws["B3"] = "欄標籤"

    index_col = "Adj. MCH1" if "Adj. MCH1" in df.columns else (
        "Item Class" if "Item Class" in df.columns else None
    )
    status_col = "Item Status" if "Item Status" in df.columns else None
    if df is None or df.empty or index_col is None or status_col is None or "SKU" not in df.columns:
        return

    summary_df = df.copy()
    index_values = summary_df[index_col].fillna("").astype(str).str.strip()
    status_values = summary_df[status_col].fillna("").astype(str).str.strip()
    summary_df[index_col] = index_values.mask(index_values == "", "(空白)")
    summary_df[status_col] = status_values.mask(status_values == "", "(空白)")

    table = summary_df.pivot_table(
        index=index_col,
        columns=status_col,
        values="SKU",
        aggfunc="count",
        margins=True,
        margins_name="總計",
        fill_value=0,
    )
    # 空白組排在最後（Excel 樞紐慣例），總計列／欄置末
    statuses = [c for c in table.columns if c != "總計"]
    statuses = [c for c in statuses if c != "(空白)"] + [c for c in statuses if c == "(空白)"]
    total_col = 2 + len(statuses)
    row_order = [v for v in table.index if v not in ("總計", "(空白)")]
    if "(空白)" in table.index:
        row_order.append("(空白)")
    if "總計" in table.index:
        row_order.append("總計")
    table = table.reindex(
        index=row_order,
        columns=statuses + (["總計"] if "總計" in table.columns else []),
    )

    ws.cell(row=4, column=1, value="列標籤")
    for j, status in enumerate(statuses, start=2):
        ws.cell(row=4, column=j, value=str(status))
    ws.cell(row=4, column=total_col, value="總計")

    for i, (idx, row) in enumerate(table.iterrows(), start=5):
        cell = ws.cell(row=i, column=1, value=str(idx))
        cell.alignment = Alignment(horizontal="left")
        for j, status in enumerate(statuses, start=2):
            count = int(row[status])
            if count:
                ws.cell(row=i, column=j, value=count)
        total = int(row["總計"])
        if total:
            ws.cell(row=i, column=total_col, value=total)

    # 欄寬（對齊範本；狀態欄多於 4 個時其餘沿用預設）
    widths = [35.28515625, 9.7109375, 12.7109375, 8.28515625, 10.0, 5.28515625]
    for j, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = width


def generate_excel_shop_simplify(raw: pd.DataFrame, progress_callback=None) -> BytesIO:
    """產生兩個工作表：Summary（Adj. MCH1 × Item Status 計數）與 Export（回覆明細）。

    Export 沿用店舖精簡回覆表範本的格式：Brand／SKU 排序、凍結 F2、指定欄寬與隱藏欄、
    數字格式、綠色／紫色重點欄、「無賣」列的淡紅標示，以及三類紅色底標示
    （Item Status＝Discontinued、RP Type＝ND、過去90日有無存貨＝無存貨）。
    """
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = _SHOP_SIMPLIFY_SUMMARY_SHEET
    ws = wb.create_sheet(title=_SHOP_SIMPLIFY_EXPORT_SHEET)

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    df = raw if raw is not None else pd.DataFrame()
    headers = list(df.columns)
    ws.append([_shop_simplify_cell(c) for c in headers])

    total = len(df)
    for i, row in enumerate(df.itertuples(index=False, name=None)):
        ws.append([_shop_simplify_cell(v) for v in row])
        if progress_callback and (i % 2000 == 0 or i == total - 1):
            progress_callback(0.05 + 0.75 * ((i + 1) / total), f"寫入 Export ({i+1}/{total})...")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    fill_blue = PatternFill(start_color=_SHOP_SIMPLIFY_HEADER_BLUE, end_color=_SHOP_SIMPLIFY_HEADER_BLUE, fill_type="solid")
    fill_green = PatternFill(start_color=_SHOP_SIMPLIFY_GREEN, end_color=_SHOP_SIMPLIFY_GREEN, fill_type="solid")
    fill_purple = PatternFill(start_color=_SHOP_SIMPLIFY_PURPLE, end_color=_SHOP_SIMPLIFY_PURPLE, fill_type="solid")
    fill_pink = PatternFill(start_color=_SHOP_SIMPLIFY_NO_SALE_PINK, end_color=_SHOP_SIMPLIFY_NO_SALE_PINK, fill_type="solid")
    fill_red = PatternFill(start_color=_SHOP_SIMPLIFY_RED, end_color=_SHOP_SIMPLIFY_RED, fill_type="solid")

    # 表頭：粗體、指定底色、細框、置中置頂、自動換行；列高 70.5
    ws.row_dimensions[1].height = 70.5
    for col_idx, name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        if name in _SHOP_SIMPLIFY_GREEN_HEADERS:
            cell.fill = fill_green
        elif name in _SHOP_SIMPLIFY_PURPLE_COLUMNS:
            cell.fill = fill_purple
        else:
            cell.fill = fill_blue
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

        letter = get_column_letter(col_idx)
        width = _SHOP_SIMPLIFY_COLUMN_WIDTHS.get(name)
        if width is None and _is_shop_simplify_top80_column(name):
            # 指定店舖變體的判定欄較長（…Top 80%（件數）／（金額））；原模式的「是否該店的Top 80%」不加闊
            text = str(name).strip()
            is_nature = text.startswith("是否同Shop Nature的Top ")
            if text.endswith(_SHOP_SIMPLIFY_VALUE_SUFFIX) or text.endswith(_SHOP_SIMPLIFY_QTY_SUFFIX):
                width = 26.0 if is_nature else 22.0
            else:
                width = 16.0 if is_nature else 14.0
        if width is None and isinstance(name, str) and re.fullmatch(r"Net Sales \(P\d+D\)", name.strip()):
            # 金額母體欄（Net Sales (P73D)）
            width = 14.0
        if width is not None:
            ws.column_dimensions[letter].width = width
        if name in _SHOP_SIMPLIFY_HIDDEN_COLUMNS:
            ws.column_dimensions[letter].hidden = True

    # 資料格：綠色重點欄、紫色 Sold Qty (P90D)、「無賣」淡紅標示、
    # 以及紅色底重點（Item Status＝Discontinued／RP Type＝ND／過去90日有無存貨＝無存貨）
    green_cols = [i for i, name in enumerate(headers, start=1) if name in _SHOP_SIMPLIFY_GREEN_DATA_COLUMNS]
    purple_cols = [i for i, name in enumerate(headers, start=1) if name in _SHOP_SIMPLIFY_PURPLE_COLUMNS]
    sold_col_idx = headers.index(_SHOP_SIMPLIFY_SOLD_COLUMN) + 1 if _SHOP_SIMPLIFY_SOLD_COLUMN in headers else None
    red_rules = [
        (headers.index(name) + 1, _shop_simplify_red_mask(df[name], target))
        for name, target in _SHOP_SIMPLIFY_RED_RULES
        if name in headers and name in df.columns
    ]
    number_formats = [
        (i, _SHOP_SIMPLIFY_NUMBER_FORMATS[name])
        for i, name in enumerate(headers, start=1)
        if name in _SHOP_SIMPLIFY_NUMBER_FORMATS
    ]
    for row_idx in range(2, total + 2):
        for col_idx in green_cols:
            ws.cell(row=row_idx, column=col_idx).fill = fill_green
        for col_idx in purple_cols:
            ws.cell(row=row_idx, column=col_idx).fill = fill_purple
        if sold_col_idx is not None:
            sold_cell = ws.cell(row=row_idx, column=sold_col_idx)
            if sold_cell.value == _SHOP_SIMPLIFY_SOLD_NO:
                sold_cell.fill = fill_pink
        # 紅色底最後套用，故 RP Type＝ND 會蓋過該欄的綠色
        for col_idx, mask in red_rules:
            if mask[row_idx - 2]:
                ws.cell(row=row_idx, column=col_idx).fill = fill_red
        for col_idx, fmt in number_formats:
            ws.cell(row=row_idx, column=col_idx).number_format = fmt

    # 凍結於 SKU Desc 之後（範本為 F2）；縮放 85%
    if "SKU Desc" in headers:
        ws.freeze_panes = f"{get_column_letter(headers.index('SKU Desc') + 2)}2"
    elif headers:
        ws.freeze_panes = "A2"
    ws.sheet_view.zoomScale = 85
    ws.auto_filter.ref = ws.dimensions

    if progress_callback:
        progress_callback(0.85, "建立 Summary 工作表...")
    _write_shop_simplify_summary(ws_summary, df)

    # 開啟檔案時預設顯示 Export（範本行為）
    wb.active = wb.sheetnames.index(_SHOP_SIMPLIFY_EXPORT_SHEET)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")
    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf
