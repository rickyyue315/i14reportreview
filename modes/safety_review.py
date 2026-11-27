# -*- coding: utf-8 -*-
"""Safety Stock Review模式：由 transformer.py 拆分而來。
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
    _ADJ_SAFETY_SOURCE_COL,
    _auto_width,
    _cached_shop_names,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _display_safety_header,
    _extract_shop_name,
    _extract_sku_desc,
    _footer_row_mask,
    _iter_header_rows,
    _nature_rate_arr,
    _normalize_ref_date,
    _parse_launch_dates,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _cached_safe_sku_series,
    _write_guide_sheet,
    _write_metric_cards,
    _write_raw_placeholder,
    _write_raw_sheet,
    validate_input,
)


from modes.framework import (
    _HYBRID_DIMENSION_LABELS,
    _HYBRID_SAFETY_DAYS_DEFAULT,
    _HYBRID_SALES_OOS_DAYS_DEFAULT,
    _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT,
    _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT,
    _attach_sales_classification_columns,
    _hybrid_dimension,
    _hybrid_dimension_groups,
    _hybrid_group_key_series,
    _hybrid_om_groups_from_values,
    _hybrid_safety_days,
    _hybrid_safety_days_label,
    _hybrid_sales_params,
)


from modes.issues import (
    _build_combined_shop_issues,
    _write_shop_issue_sheet,
)
__all__ = [
    "COMBINED_RAW_COLS",
    "_CHANGE_CAT_DESC",
    "_ND_REC_DESC",
    "_ND_REC_DISPLAY",
    "_SSR_SAFETY_BASIS_RATE",
    "_SSR_SAFETY_BASIS_TURNDAY",
    "_SSR_SAFETY_BASIS_LABELS",
    "_SSR_SAFETY_DAYS_COLUMN",
    "_SSR_PIVOT_LEVEL_COL",
    "_SSR_SAFETY_LINE_COLUMN",
    "_SSR_TARGET_TOTAL_COLUMN",
    "_SSR_TARGET_SHARE_COLUMN",
    "_SSR_TARGET_NOTE_COLUMN",
    "_SSR_TARGET_BASIS_LABEL",
    "_SSR_TARGET_FILE_CURRENT_COLUMN",
    "_SSR_TARGET_FILE_ALIASES",
    "_SSR_TARGET_CURRENT_QTY_COLUMN",
    "_SSR_TARGET_TEMPLATE_COLUMNS",
    "_build_combined_target_template",
    "_normalize_ssr_safety_basis",
    "_ssr_column_key",
    "_ssr_normalize_target_columns",
    "_ssr_read_target_file",
    "_ssr_normalize_target_map",
    "_ssr_largest_remainder",
    "_ssr_target_weight_series",
    "_add_combined_recommendation_columns",
    "_annotate_nd_recommendation",
    "_build_combined_by_shop_summary",
    "_build_combined_classification_summary",
    "_build_combined_dimension_cross_tab",
    "_build_combined_target_summary",
    "_build_combined_impact_summary",
    "_build_combined_issue_pivot",
    "_write_issue_pivot_block",
    "_build_nd_rec_qty_nd",
    "_write_combined_guide_sheet",
    "_write_summary_sheet",
    "generate_excel_combined",
    "transform_combined",
]


COMBINED_RAW_COLS = [
    "Brand",
    "SKU",
    "SKU Desc",
    "Buyer",
    "Adj. SST",
    "Adj. MCH1",
    "Vendor MC",
    "Item Status",
    "Launch Date",
    "Item Class",
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "RP Type",
    "Adj. Safety Stock",  # 表頭顯示為 Safety Stock (現時)
    "Recommend Safety Stock",
    "Recommend Safety Stock Turnday",
    "Target Safety (全線合計)",
    "Safety 分配佔比",
    "Target 分配提示",
    "Safety Qty Change",
    "Safety Qty Change 分類",
    "Recommendation of new Safety Qty / ND",
    "Buyer reply",
    "Shop Monthly Sold Qty (P90D)",
    "Safety VS Monthly Sold Qty",
    "Projected Monthly Sold Qty",
    "Overstock Cap (X)",
    "OOS Floor (Y)",
    "Safety 基準天數",
    "Safety 周轉所需量",
    "D001 Net Stock Qty (CM)",
    "Shop Stock Qty (CM)",
    "Sold Qty (P90D)",
    "銷量帶",
    "銷量存貨分類",
    "Turnday (P90D Qty)",
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Supply Source",
]


# Safety 計算基準（可切換選項；預設 "rate" 維持現行 X/Y 行為）：
# - "rate"：月銷倍率（Overstock Cap X／OOS Floor Y；建議值 = max(Y, Z if Z ≤ X else X)）
# - "turnday"：周轉天數（Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 有效統計日數 × 該列 Nature
#   基準天數)，有效統計日數 = min(90, max(上市至參考日天數（含首尾）, 3))；與 Safety Stock x NetStock
#   Review 同口徑；建議值直接取周轉所需量，有銷量或新品時含最低 3 件保底（舊品零銷量維持 0），
#   不設 Overstock 上限）。
_SSR_SAFETY_BASIS_RATE = "rate"
_SSR_SAFETY_BASIS_TURNDAY = "turnday"
_SSR_SAFETY_BASIS_LABELS = {
    _SSR_SAFETY_BASIS_RATE: "月銷倍率 (X/Y)",
    _SSR_SAFETY_BASIS_TURNDAY: "周轉天數 (按 Nature)",
}

# 周轉天數基準欄（與 Safety Stock x NetStock Review 同名同義，供對照與稽核）：
# - Safety 基準天數：該列實際採用的基準天數（按 Shop Nature；預設與 Hybrid 相同）
# - Safety 周轉所需量：ceil(Sold Qty (P90D) ÷ 有效統計日數 × 基準天數)，有銷量或新品時含最低 3 件保底
#   （有效統計日數：取 min(90, max(上市至參考日天數, 3))，舊品為 90）
_SSR_SAFETY_DAYS_COLUMN = "Safety 基準天數"
_SSR_SAFETY_LINE_COLUMN = "Safety 周轉所需量"


def _normalize_ssr_safety_basis(safety_basis) -> str:
    """Safety 計算基準正規化：未知值退回 "rate"（維持現行行為，不拋錯）。"""
    raw = str(safety_basis or "").strip().lower()
    if raw in (_SSR_SAFETY_BASIS_TURNDAY, "turnover", "days", "周轉天數"):
        return _SSR_SAFETY_BASIS_TURNDAY
    return _SSR_SAFETY_BASIS_RATE


# ── P90D 分母修正（有效統計日數）──
# Sold Qty (P90D) 是「近 90 日累計」，但新品上市不足 90 日時，該累計只涵蓋上市後的
# 少數日子（例如 11 月 13 日上市、數據截至 11 月 15 日 → 只有 3 日）。固定除以 90 會
# 把日均銷量按比例低估（3 日 → 低估 30 倍），令 Safety 周轉所需量全線撞 3 件保底，
# 賣得越快的店被削得越深。改按「實際可售日數」作分母：
#   eff_days = min(90, max(上市至參考日（含首尾）, floor))
# floor 預設 3 日——只防治以 1–2 日尖峰外推，不改變「按實際日數」的口徑；設 14 則
# 額外假設上市初期銷量會回落（保守），設 1 即完全按實際日數。
_SSR_SOLD_DAYS_MAX = 90
_SSR_SOLD_DAYS_FLOOR_DEFAULT = 3


def _sold_days_denominator(df, ref_date=None, floor_days=None) -> np.ndarray:
    """每列 Sold Qty (P90D) 的有效統計日數（即除以多少日換算日均銷量）。

    舊品（上市滿 90 日）／無 Launch Date／無法解析／上市日為未來 → 一律 90，維持原行為；
    新品 → clip(上市至參考日天數（含首尾）, floor_days, 90)。逐列回傳 float 陣列，長度同 df。
    """
    n = len(df)
    days = np.full(n, float(_SSR_SOLD_DAYS_MAX))
    if n == 0 or "Launch Date" not in df.columns:
        return days
    _floor = _SSR_SOLD_DAYS_FLOOR_DEFAULT if floor_days is None else floor_days
    try:
        _floor = max(1.0, float(_floor))
    except (TypeError, ValueError):
        _floor = float(_SSR_SOLD_DAYS_FLOOR_DEFAULT)
    launch = _parse_launch_dates(df["Launch Date"].fillna("").astype(str).str.strip())
    _ref = pd.Timestamp(_normalize_ref_date(ref_date))
    # 可售日數含首尾（上市日當日已可售）：參考日 − 上市日 + 1
    _diff = (_ref - launch).dt.days.to_numpy(dtype=float)
    valid = launch.notna().to_numpy() & (_diff >= 0)
    if valid.any():
        days[valid] = np.clip(_diff[valid] + 1.0, _floor, float(_SSR_SOLD_DAYS_MAX))
    return days




def _build_nd_rec_qty_nd(is_nd, is_p90d_0, n):
    arr = np.full(n, "", dtype=object)
    if is_p90d_0.any():
        arr[is_p90d_0.values] = "ND"
    if is_nd.any():
        arr[is_nd.values] = "0"
    return arr



def _annotate_nd_recommendation(nd_rec, df):
    """為「ND」轉換建議依 Adj. MCH1 / Adj. SST 加註類別提示：
    Adj. MCH1 = 02 - MAKE UP → ND(02 - MAKE UP)；
    Adj. SST 為 03 - Private Label 系（(1)Major / (2)Health）→ ND(Private Label)；
    兩者皆是 → ND(Private Label - MAKE UP)。
    只加註「ND」轉換建議；「0」列（本來已是 ND）與空白列不動。
    """
    rec = np.asarray(nd_rec, dtype=object)
    is_nd_suggest = rec == "ND"
    if not is_nd_suggest.any():
        return rec
    mch1 = (
        df.get("Adj. MCH1", _safe_default_series("", df.index))
        .fillna("").astype(str).str.strip().str.lower()
    )
    sst = (
        df.get("Adj. SST", _safe_default_series("", df.index))
        .fillna("").astype(str).str.strip().str.lower()
    )
    is_makeup = (mch1 == "02 - make up").values
    is_pl = sst.str.startswith("03 - private label").values
    rec = rec.copy()
    rec[is_nd_suggest & is_pl & is_makeup] = "ND(Private Label - MAKE UP)"
    rec[is_nd_suggest & is_pl & ~is_makeup] = "ND(Private Label)"
    rec[is_nd_suggest & ~is_pl & is_makeup] = "ND(02 - MAKE UP)"
    return rec



# ── 指定總量分配（Target Safety）──
# 用途：公司有時會強制大幅增加／減少 Safety（例如配合大堆頭陳列、推廣期），
# 下達的是「這個 SKU 全線合計多少件」，而不是逐店按銷量計出來。此機制只覆寫
# 被指定的 SKU，其餘 SKU 一律沿用原有基準（月銷倍率／周轉天數），互不影響。
#
# 分配基準固定為各半：50% 按 Sold Qty (P90D) 佔比、50% 按現有 Safety 佔比——
# 令銷量強的店分得多，同時保留原有設定的相對結構，避免純按銷量令個別旺店暴增。
# 取整用最大餘數法，保證每個 SKU 的分配總和恆等於指定值（與調貨段 R7 同一準則）。
_SSR_TARGET_TOTAL_COLUMN = "Target Safety (全線合計)"
_SSR_TARGET_SHARE_COLUMN = "Safety 分配佔比"
_SSR_TARGET_NOTE_COLUMN = "Target 分配提示"
_SSR_TARGET_WEIGHT_SOLD = 0.5
_SSR_TARGET_WEIGHT_SAFETY = 0.5
_SSR_TARGET_BASIS_LABEL = "各半（Sold Qty (P90D) 佔比 50% + 現有 Safety 佔比 50%）"

# 上載 Target 清單的欄名對應：一個系統欄位對多個使用者寫法（比對前先經
# `_ssr_column_key` 正規化，故「現時 Safety」「現時safety」「Safety Stock (現時)」
# 都會命中同一欄）。`現時Safety` 只作核對，不參與計算。
_SSR_TARGET_FILE_CURRENT_COLUMN = "現時Safety"
_SSR_TARGET_CURRENT_QTY_COLUMN = "現時Safety Qty"
_SSR_TARGET_TEMPLATE_COLUMNS = (
    "Brand", "SKU", "SKU Desc", "Vendor MC", "Buyer",
    _SSR_TARGET_CURRENT_QTY_COLUMN, _SSR_TARGET_TOTAL_COLUMN,
)
# 來源檔常見的非資料列：PowerBI／Excel 貼出的「Total」合計列與「Applied filters: …」註腳列。
# 使用者忘記刪除時不可讓它變成一個叫「nan」的 SKU（會被當成要分配的貨品）。
_SSR_TARGET_SKIP_LABELS = ("total", "grand total", "subtotal", "總計", "合計", "小計")
_SSR_TARGET_SKIP_NOTE_KEYWORD = "applied filters"


def _ssr_junk_row_mask(series):
    """判斷某欄是否屬「非資料列」：空白／NaN、Total 類合計標籤、或 Applied filters 註腳。

    NaN 經 astype(str) 會變成「nan」字串，故必須以原 series 的 isna() 判斷，
    否則整列空白會被當成一個 SKU。
    """
    text = series.astype(str).str.strip()
    lower = text.str.lower()
    return (
        series.isna()
        | text.eq("")
        | lower.isin(("nan", "none", "<na>", "nat") + _SSR_TARGET_SKIP_LABELS)
        | lower.str.contains(_SSR_TARGET_SKIP_NOTE_KEYWORD, regex=False, na=False)
    )
_SSR_TARGET_FILE_ALIASES = {
    "SKU": ("sku", "貨品編號", "貨號", "item code", "item"),
    "Brand": ("brand", "品牌"),
    "SKU Desc": ("sku desc", "sku description", "description", "貨品描述", "貨品名稱", "品名"),
    "Vendor MC": ("vendor mc", "vendor", "供應商"),
    "Buyer": ("buyer", "採購", "買手"),
    _SSR_TARGET_FILE_CURRENT_COLUMN: (
        "現時 safety", "現時 safety stock", "現時safety", "現時safety stock",
        "現有 safety", "現有safety", "safety stock (現時)", "adj. safety stock",
        "現時安全庫存", "現時safety qty",
    ),
    _SSR_TARGET_TOTAL_COLUMN: (
        "我想輸入的 safety qty", "我想輸入的safety qty", "我想輸入的 safety", "我想輸入的safety",
        "目標 safety qty", "目標safety qty", "目標 safety", "目標safety",
        "target safety qty", "target safety", "target safety (全線合計)",
        "輸入 safety qty", "新 safety qty", "新safety", "建議 safety qty",
    ),
}


def _ssr_normalize_target_map(target_safety):
    """指定總量輸入正規化：接受 {SKU: 總量}、[(SKU, 總量)] 或 DataFrame（首欄 SKU、次欄總量）。

    回傳 dict[str, int]：SKU 以 `_safe_sku` 正規化（與 RAW 的 SKU 同一口徑，數字
    與文字寫法皆可對上），總量轉正整數。空白 SKU、非數值、0 或負數一律略過
    （不視為指定，該 SKU 按原基準計算）。同一 SKU 重複出現時以最後一筆為準。
    """
    if target_safety is None:
        return {}
    pairs = []
    if isinstance(target_safety, pd.DataFrame):
        if target_safety.empty or len(target_safety.columns) < 2:
            return {}
        _cols = list(target_safety.columns)
        pairs = list(zip(target_safety[_cols[0]], target_safety[_cols[1]]))
    elif isinstance(target_safety, dict):
        pairs = list(target_safety.items())
    else:
        for item in target_safety:
            try:
                _k, _v = item
            except (TypeError, ValueError):
                continue
            pairs.append((_k, _v))

    out = {}
    for key, val in pairs:
        if key is None or (isinstance(key, float) and pd.isna(key)):
            continue
        sku = str(key).strip()
        if not sku:
            continue
        try:
            sku = str(_safe_sku(sku)).strip()
        except Exception:
            pass
        try:
            qty = int(float(val))
        except (TypeError, ValueError):
            continue
        if qty <= 0:
            continue
        out[sku] = qty
    return out


def _ssr_column_key(name) -> str:
    """欄名比對鍵：全形轉半形、去空白與常見標點、轉小寫。

    讓「現時 Safety」「現時Safety」「Safety Stock (現時)」等寫法都對得上同一欄，
    避免使用者上載的清單因空格或括號不同而讀不到。
    """
    s = unicodedata.normalize("NFKC", str(name)).strip().lower()
    for ch in " \t\r\n()（）[]【】{}_-／/\\:：.,，、":
        s = s.replace(ch, "")
    return s


def _ssr_normalize_target_columns(df):
    """把上載 Target 清單的欄名對應到系統欄名。

    回傳 (改名後的 DataFrame, 已識別欄位清單, 缺少的必要欄位清單)。
    只保留識別到的欄位；`SKU` 與目標總量為必要，其餘（Brand／SKU Desc／
    Vendor MC／Buyer／現時Safety）只作顯示與核對，缺了不影響計算。
    """
    if df is None or len(df) == 0:
        return pd.DataFrame(), [], ["SKU", _SSR_TARGET_TOTAL_COLUMN]

    key_to_target = {}
    for _target, _aliases in _SSR_TARGET_FILE_ALIASES.items():
        for _alias in _aliases:
            key_to_target.setdefault(_ssr_column_key(_alias), _target)

    mapping = {}
    used = set()
    for col in df.columns:
        _tgt = key_to_target.get(_ssr_column_key(col))
        if _tgt and _tgt not in used:
            mapping[col] = _tgt
            used.add(_tgt)
    out = df[list(mapping)].rename(columns=mapping)
    missing = [c for c in ("SKU", _SSR_TARGET_TOTAL_COLUMN) if c not in out.columns]
    return out, list(out.columns), missing


def _ssr_read_target_file(data, filename=""):
    """讀取上載的 Target 清單，回傳 (正規化 DataFrame, 警告清單)。

    - `.xlsx`：逐個工作表尋找同時具備 `SKU` 與目標總量欄的一頁（不假設次序）；
      `.csv`：先試 UTF-8，失敗再試 Big5（Excel 另存的繁體 CSV 常見）。
    - 容許千分位逗號；略過空白／合計列、目標值空白或 ≤ 0 的列；
      同一 SKU 重複出現時以最後一筆為準並列警告。
    - 找不到必要欄位或無可用列時拋 ValueError，訊息列出已識別的欄位供核對。
    """
    if data is None:
        raise ValueError("沒有收到檔案內容。")
    raw_bytes = data if isinstance(data, (bytes, bytearray)) else bytes(data)
    name = str(filename or "").lower()
    buf = BytesIO(raw_bytes)

    best = None
    candidates = []
    if name.endswith(".csv") or name.endswith(".txt"):
        for enc in ("utf-8", "utf-8-sig", "big5"):
            try:
                buf.seek(0)
                candidates = [pd.read_csv(buf, encoding=enc, header=0)]
                break
            except (UnicodeDecodeError, LookupError):
                continue
            except Exception as exc:
                raise ValueError(f"CSV 讀取失敗：{exc}") from exc
        if not candidates:
            raise ValueError("CSV 編碼無法辨識（已試 UTF-8 與 Big5）。")
    else:
        try:
            book = pd.ExcelFile(BytesIO(raw_bytes))
        except Exception as exc:
            raise ValueError(f"Excel 讀取失敗：{exc}") from exc
        try:
            candidates = [book.parse(sheet) for sheet in book.sheet_names]
        finally:
            book.close()

    first_error = None
    for frame in candidates:
        if frame is None or frame.empty:
            continue
        normalized, cols, missing = _ssr_normalize_target_columns(frame)
        if missing:
            if first_error is None:
                first_error = (cols, missing)
            continue
        best = normalized
        break
    if best is None:
        if first_error is not None:
            cols, missing = first_error
            raise ValueError(
                "找不到必要欄位：" + "、".join(missing)
                + "（已識別：" + ("、".join(str(c) for c in cols) or "無") + "）"
            )
        raise ValueError("檔案沒有可讀的資料列。")

    warnings = []
    _skip_mask = _ssr_junk_row_mask(best["SKU"])
    _dropped = int(_skip_mask.sum())
    if _dropped:
        warnings.append(f"已略過 {_dropped} 列空白、合計（Total）或 Applied filters 註腳列。")
    best = best[~_skip_mask].copy()

    def _num(series):
        if series.dtype == object:
            series = series.astype(str).str.replace(",", "", regex=False).str.strip()
        return pd.to_numeric(series, errors="coerce")

    best["SKU"] = best["SKU"].map(_safe_sku).astype(str).str.strip()
    best[_SSR_TARGET_TOTAL_COLUMN] = _num(best[_SSR_TARGET_TOTAL_COLUMN])
    if _SSR_TARGET_FILE_CURRENT_COLUMN in best.columns:
        best[_SSR_TARGET_FILE_CURRENT_COLUMN] = _num(best[_SSR_TARGET_FILE_CURRENT_COLUMN])

    _invalid = best[_SSR_TARGET_TOTAL_COLUMN].isna() | (best[_SSR_TARGET_TOTAL_COLUMN] <= 0)
    if int(_invalid.sum()):
        warnings.append(
            f"有 {int(_invalid.sum())} 隻 SKU 的「{_SSR_TARGET_TOTAL_COLUMN}」空白、非數值或 ≤ 0，已略過。"
        )
        best = best[~_invalid].copy()

    _dup = best["SKU"].duplicated(keep=False)
    if int(_dup.sum()):
        warnings.append(
            f"有 {int(best.loc[_dup, 'SKU'].nunique())} 隻 SKU 重複出現，已以最後一筆為準。"
        )
        best = best.drop_duplicates(subset="SKU", keep="last")

    if best.empty:
        raise ValueError("清單沒有可用的 SKU 與目標數量（全部為空白、合計列或非數值）。")

    best[_SSR_TARGET_TOTAL_COLUMN] = best[_SSR_TARGET_TOTAL_COLUMN].astype(int)
    return best.reset_index(drop=True), warnings


def _ssr_target_desc(item_shop, site_code) -> str:
    """取 SKU 描述：先用共用 `_extract_sku_desc`（Item_Shop 含 Site Code 時最準），
    取不到時退回「去掉 (品牌) 與前置 SKU、再去掉最後一段（店舖名）」。

    來源的 Item_Shop 有兩種寫法：`{Desc}-{SiteCode}` 與 `(Brand) {SKU}-{Desc}-{Shop}`，
    後者沒有 Site Code，共用解析會回空，故需要這個後備。
    """
    desc = _extract_sku_desc(item_shop, site_code)
    if desc:
        return desc
    if item_shop is None or (isinstance(item_shop, float) and pd.isna(item_shop)):
        return ""
    text = re.sub(r"^\([^)]+\)\s*", "", str(item_shop).strip())
    text = re.sub(r"^\d+\s*-\s*", "", text)
    if "-" in text:
        text = text.rsplit("-", 1)[0]
    return text.strip().strip(",，;；:： ")


def _build_combined_target_template(df_input):
    """由來源資料建立「指定總量分配」填寫表：每個 SKU 一列。

    欄位：Brand／SKU／SKU Desc／Vendor MC／Buyer／現時Safety Qty／
    Target Safety (全線合計)。
    - `現時Safety Qty`＝該 SKU 在全部店舖的 Adj. Safety Stock 合計，
      與分配所用口徑一致（同一份上傳資料）
    - Brand／Vendor MC／Buyer 取該 SKU 首個非空值；SKU Desc 由 Item_Shop 解析
      （來源已有「SKU Desc」欄時直接沿用），只解析不重複 SKU，成本與店舖數無關
    - `Target Safety (全線合計)` 預設 0，由使用者在介面填寫；排序按 Brand、SKU
    - SKU 空白或無法辨識的列不納入；來源缺欄時對應欄留空
    """
    empty = pd.DataFrame(columns=list(_SSR_TARGET_TEMPLATE_COLUMNS))
    empty.attrs["dropped_rows"] = 0
    if df_input is None or len(df_input) == 0 or "SKU" not in df_input.columns:
        return empty

    sku_raw = df_input["SKU"]
    sku_text = sku_raw.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    # Total 合計列與 Applied filters 註腳列常被忘記刪除，一律不可當成 SKU。
    # 先沿用全系統同一準則（`_footer_row_mask` 會看 Column A／Item_Shop／Buyer，
    # 且不會誤刪「Total Care 洗髮水」這類品名），再補上 SKU 本身為空白／標籤的情況。
    _junk = _ssr_junk_row_mask(sku_raw) | _footer_row_mask(df_input)
    dropped_rows = int(_junk.sum())
    # 來源可能用顯示名「Safety Stock (現時)」，沿用全系統同一欄名解析
    _adj_col = _resolve_column_name(_ADJ_SAFETY_SOURCE_COL, set(df_input.columns))
    if _adj_col is not None:
        safety = _safe_numeric(df_input[_adj_col])
    else:
        safety = _safe_default_series(0, df_input.index)
    work = pd.DataFrame({"_sku": sku_text, "_z": safety}, index=df_input.index)
    work = work[~_junk]

    for col in ("Brand", "Buyer", "Vendor MC"):
        if col in df_input.columns:
            series = df_input[col].astype(str).str.strip()
            work[col] = series.where(~series.isin(("", "nan", "None")))

    if "SKU Desc" in df_input.columns:
        series = df_input["SKU Desc"].astype(str).str.strip()
        work["_desc"] = series.where(~series.isin(("", "nan", "None")))

    if work.empty:
        return empty

    totals = work.groupby("_sku")["_z"].sum()
    agg_cols = [c for c in ("Brand", "Buyer", "Vendor MC") if c in work.columns]
    agg = work.groupby("_sku")[agg_cols].first() if agg_cols else pd.DataFrame(index=totals.index)

    if "_desc" in work.columns:
        desc = work.groupby("_sku")["_desc"].first()
    else:
        # 只保留每個 SKU 第一筆有 Item_Shop 的列再解析，避免逐列處理全表
        _site = df_input.get("Site Code", _safe_default_series("", df_input.index)).astype(str).str.strip()
        _item = df_input.get("Item_Shop", _safe_default_series("", df_input.index))
        _pairs = pd.DataFrame({"_sku": sku_text, "_site": _site, "_item": _item}, index=df_input.index)
        _item_text = _pairs["_item"].astype(str).str.strip()
        _pairs = _pairs[_pairs["_sku"].ne("") & ~_item_text.isin(("", "nan", "None"))]
        _pairs = _pairs.drop_duplicates("_sku", keep="first")
        if _pairs.empty:
            desc = pd.Series(dtype=object)
        else:
            desc = pd.Series(
                [_ssr_target_desc(it, st) for it, st in zip(_pairs["_item"], _pairs["_site"])],
                index=_pairs["_sku"].to_numpy(dtype=object),
            )

    out = pd.DataFrame(index=totals.index)
    for col in ("Brand", "Buyer", "Vendor MC"):
        out[col] = agg[col].reindex(out.index).fillna("").astype(str) if col in agg.columns else ""
    out["SKU"] = out.index.astype(str)
    out["SKU Desc"] = desc.reindex(out.index).fillna("").astype(str) if len(desc) else ""
    out[_SSR_TARGET_CURRENT_QTY_COLUMN] = totals.reindex(out.index).fillna(0).round(0).astype(int)
    out[_SSR_TARGET_TOTAL_COLUMN] = 0

    out = out[list(_SSR_TARGET_TEMPLATE_COLUMNS)]
    out = out.sort_values(["Brand", "SKU"], kind="stable").reset_index(drop=True)

    # 註腳文字若落在其他欄（例如 SKU 空白但 SKU Desc 寫「Applied filters: …」），一律剔走
    _note_cols = [c for c in ("Brand", "SKU Desc", "Buyer", "Vendor MC") if c in out.columns]
    if _note_cols and not out.empty:
        _note = (
            out[_note_cols].astype(str)
            .apply(
                lambda s: s.str.lower().str.contains(
                    _SSR_TARGET_SKIP_NOTE_KEYWORD, regex=False, na=False
                )
            )
            .any(axis=1)
        )
        if bool(_note.any()):
            dropped_rows += int(_note.sum())
            out = out[~_note].reset_index(drop=True)

    out.attrs["dropped_rows"] = dropped_rows
    return out


def _ssr_largest_remainder(total, weights) -> np.ndarray:
    """按權重把整數 total 分成整數件（最大餘數法），保證加總等於 total。

    先取 floor，餘量按小數部分由大到小逐件補 1；小數部分相同時依原次序
    （穩定排序），令同一批資料每次分配的結果一致。權重全為 0 或負數時退回均分；
    total ≤ 0 時回傳全 0。與調貨段 R7 同一取整準則。
    """
    w = np.asarray(weights, dtype=float)
    n = len(w)
    if n == 0:
        return np.zeros(0, dtype=int)
    total = int(total)
    if total <= 0:
        return np.zeros(n, dtype=int)
    w = np.where(np.isfinite(w) & (w > 0), w, 0.0)
    _sum = float(w.sum())
    if _sum <= 0:
        w = np.ones(n, dtype=float)
        _sum = float(n)
    raw = total * w / _sum
    base = np.floor(raw).astype(int)
    remainder = total - int(base.sum())
    if remainder > 0:
        order = np.argsort(-(raw - base), kind="stable")
        base[order[:remainder]] += 1
    return base


def _ssr_target_weight_series(sold, safety) -> np.ndarray:
    """各店分配權重：50% 按 Sold Qty (P90D) 佔比 + 50% 按現有 Safety 佔比。

    某一半的合計為 0（例如全線零銷量、或現有 Safety 全為 0）時，該半不參與，
    等同把權重全給另一半；兩半皆為 0 時退回均分，令指定總量仍可全數分配。
    負數與缺值一律按 0 計。回傳權重已正規化為合計 1（n > 0 時）。
    """
    sold_arr = np.asarray(sold, dtype=float)
    safety_arr = np.asarray(safety, dtype=float)
    sold_arr = np.where(np.isfinite(sold_arr) & (sold_arr > 0), sold_arr, 0.0)
    safety_arr = np.where(np.isfinite(safety_arr) & (safety_arr > 0), safety_arr, 0.0)
    _sold_sum = float(sold_arr.sum())
    _safety_sum = float(safety_arr.sum())
    w_sold = sold_arr / _sold_sum if _sold_sum > 0 else np.zeros_like(sold_arr)
    w_safety = safety_arr / _safety_sum if _safety_sum > 0 else np.zeros_like(safety_arr)
    w = _SSR_TARGET_WEIGHT_SOLD * w_sold + _SSR_TARGET_WEIGHT_SAFETY * w_safety
    _w_sum = float(w.sum())
    if _w_sum > 0:
        return w / _w_sum
    return np.ones(len(sold_arr), dtype=float) / max(len(sold_arr), 1)


def _apply_target_safety(df, recommend_safety, delta, target_map, x_values, y_values):
    """把指定 SKU 的建議 Safety 改為「全線合計」由上而下分配。

    只處理 target_map 內、且 RP Type 不是 ND 的列（ND＝不執行自動補貨，
    Safety 必須為 0，故不參與分配）；其餘列原值不動。回傳
    (建議值, 變化量, Target 欄, 分配佔比欄, 提示欄, 統計 dict)。

    提示欄逐列比較分配值與 X／Y：超出庫存上限或低於缺貨底線都會標明，
    令審批人清楚知道這是公司指定值、不是公式結果。
    """
    n = len(df)
    recommend = np.array(recommend_safety, dtype=int, copy=True)
    delta_arr = np.array(delta, dtype=int, copy=True)
    target_col = np.full(n, np.nan)
    share_col = np.full(n, np.nan)
    note_col = np.full(n, "", dtype=object)
    info = {"matched": [], "missing": [], "below_current": 0}

    if not target_map or n == 0:
        info["missing"] = sorted(target_map) if target_map else []
        return recommend, delta_arr, target_col, share_col, note_col, info

    if "SKU" in df.columns:
        sku_arr = df["SKU"].apply(_safe_sku).astype(str).str.strip().to_numpy(dtype=object)
    else:
        sku_arr = np.full(n, "", dtype=object)
    rp = (
        df.get("RP Type", _safe_default_series("", df.index))
        .fillna("").astype(str).str.strip().str.upper().to_numpy()
    )
    is_nd = rp == "ND"
    sold = _safe_numeric(df.get("Sold Qty (P90D)", _safe_default_series(0, df.index))).to_numpy(dtype=float)
    safety = _safe_numeric(df.get("Adj. Safety Stock", _safe_default_series(0, df.index))).to_numpy(dtype=float)
    safety_int = np.rint(safety).astype(int)
    x_arr = np.asarray(x_values, dtype=float)
    y_arr = np.asarray(y_values, dtype=float)

    for key in sorted(target_map):
        total = int(target_map[key])
        rows = np.flatnonzero((sku_arr == key) & ~is_nd)
        if rows.size == 0:
            info["missing"].append(key)
            continue
        weights = _ssr_target_weight_series(sold[rows], safety[rows])
        alloc = _ssr_largest_remainder(total, weights)
        recommend[rows] = alloc
        delta_arr[rows] = alloc - safety_int[rows]
        target_col[rows] = total
        share_col[rows] = np.round(weights * 100.0, 2)
        for pos, ridx in enumerate(rows):
            notes = []
            if alloc[pos] > x_arr[ridx]:
                notes.append(f"超出庫存上限 X={int(x_arr[ridx])}")
            if alloc[pos] < y_arr[ridx]:
                notes.append(f"低於缺貨底線 Y={int(y_arr[ridx])}")
            if notes:
                note_col[ridx] = "；".join(notes)
        info["matched"].append(key)
        info["below_current"] += int((alloc < safety_int[rows]).sum())

    return recommend, delta_arr, target_col, share_col, note_col, info



def _add_combined_recommendation_columns(
    df, rates_overstock, rates_oos, mtd_days, month_total_days, safety_mode="increase_only",
    ref_date=None, safety_basis="rate", safety_days=None, sold_days_floor=None,
    target_safety=None,
):
    df = df.copy()
    safety_basis = _normalize_ssr_safety_basis(safety_basis)
    _ssr_days, _ = _hybrid_safety_days(safety_days)
    rates_overstock = rates_overstock or {"default": 0.50, "tourist": 0.30, "mix": 0.40}
    rates_oos = rates_oos or {"default": 0.50, "tourist": 0.50, "mix": 0.50}

    # Shop Stock Status／Shop Nature 先去前後空格（與 overstock_p90d／清理層同一準則），
    # 帶空格的狀態值／Nature 不會漏觸發歸零 override 或誤套 default 倍率
    status_col = df["Shop Stock Status"].fillna("").astype(str).str.strip()
    adj_ss = _safe_numeric(df.get("Adj. Safety Stock", _safe_default_series(0, df.index)))
    monthly_sold = _safe_numeric(df.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, df.index)))
    sold_qty = _safe_numeric(df.get("Sold Qty (CM - MTD)", _safe_default_series(0.0, df.index)))
    shop_nature = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.strip().str.lower()

    # 新品判定（供 X／Y／周轉所需量保底共用）：Launch Date 距參考日期 90 天內視為新品，
    # 新品即使零銷量仍保留最低 3 件（需有庫存可售，否則新品永無補貨）。
    # 提前計算，令三處保底共用同一新品定義（與後段 P90D 0 Sales 保護一致）。
    _ref_early = _normalize_ref_date(ref_date)
    if "Launch Date" in df.columns:
        _launch_raw_early = df["Launch Date"].fillna("").astype(str).str.strip()
        _launch_dt_early = _parse_launch_dates(_launch_raw_early)
        _is_new_early = (
            _launch_dt_early.notna()
            & (_launch_dt_early >= pd.Timestamp(_ref_early) - pd.Timedelta(days=90))
        ).to_numpy(dtype=bool)
    else:
        _is_new_early = np.zeros(len(df), dtype=bool)

    # Overstock cap X — Shop Monthly Sold Qty (P90D) 已是月均，直接乘 rate。
    # 最低 3 件保底與 OOS 底線 Y 同一準則：只適用於「有銷量或新品」的列；
    # 舊品零銷量（月均 ≤ 0 且非新品）不保底 → X = 0，令「可加可減」下建議值可歸零。
    def _compute_x(monthly_sold_arr, rate_arr, adj_ss_arr, floor_mask):
        x_pre = np.floor(monthly_sold_arr * rate_arr)
        x = np.where(floor_mask, np.maximum(x_pre, 3), np.clip(x_pre, 0, None)).astype(int)
        over_adj = x > adj_ss_arr
        x[over_adj] = adj_ss_arr[over_adj].astype(int)
        return x

    _os_default = rates_overstock.get("default", 0.50)
    _os_tourist = rates_overstock.get("tourist", 0.30)
    _os_mix = rates_overstock.get("mix", 0.40)

    os_rate_arr = _nature_rate_arr(shop_nature, _os_default, _os_tourist, _os_mix)

    # X is the upper bound for "how much safety stock is reasonable"; it applies to
    # every row (so the combined recommendation = max(Y, Z if Z <= X else X) is well
    # defined per row), independent of whether any row currently has Overstock status.
    _x_floor_mask = (monthly_sold.values > 0) | _is_new_early
    x_values = _compute_x(monthly_sold.values, os_rate_arr, adj_ss.values, _x_floor_mask)

    # OOS floor Y — Projected 月銷售採加權混合：
    #   w = mtd_days / month_total_days
    #   Projected = w × MTD線性外推 + (1−w) × P90D月均 = sold_qty + (1−w) × P90D月均
    # （月初 w 小，以穩定的 P90D 為主，避免頭幾日外推失真；月尾收斂至 MTD 實績）
    # Shop Monthly Sold Qty (P90D) 已是月均（Past 90 Days ÷ 90 × 30），直接使用；
    # 欄位整個缺失或個別列缺值（NaN）時，該列退回純 MTD 線性外推。
    if mtd_days and mtd_days > 0:
        w = mtd_days / month_total_days
        linear = (sold_qty / mtd_days) * month_total_days
        if "Shop Monthly Sold Qty (P90D)" in df.columns:
            p90d_raw = pd.to_numeric(df["Shop Monthly Sold Qty (P90D)"], errors="coerce")
            projected = sold_qty + (1.0 - w) * p90d_raw.fillna(0.0)
            projected = projected.where(p90d_raw.notna(), linear)
        else:
            projected = linear
    else:
        projected = _safe_default_series(0.0, df.index)

    # 退貨大於銷售等異常資料可令外推為負：顯示欄一律下限 0（Y 的計算本就以 > 0 判定有銷量）
    df["Projected Monthly Sold Qty"] = np.maximum(np.ceil(projected), 0).astype(int)

    _oos_default = rates_oos.get("default", 0.50)
    _oos_tourist = rates_oos.get("tourist", 0.50)
    _oos_mix = rates_oos.get("mix", 0.50)

    oos_rate_arr = _nature_rate_arr(shop_nature, _oos_default, _oos_tourist, _oos_mix)

    # y_preclip 保留未套「最低 3 件」保底前的 OOS 底線，供 _safety_floor3 判定用；
    # y_values 為現行輸出值（有預估銷量或新品時 clip min=3；舊品零銷量維持 0，不保底），兩者計算準則完全相同。
    # 舊品零銷量不應被保底上調至 3 件（無銷售即無需缺貨保護）。
    y_preclip = np.ceil(projected.values * oos_rate_arr)
    _has_proj = (np.asarray(projected.values, dtype=float) > 0) | _is_new_early
    y_values = np.where(_has_proj, np.maximum(y_preclip, 3), 0).astype(int)

    # Z = current Adj. Safety Stock
    z_values = adj_ss.values.astype(int)

    # ── 周轉天數基準（與 Safety Stock x NetStock Review 同口徑；兩種基準恆並列計算供對照）──
    # Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 有效統計日數 × 該列 Nature 基準天數)；
    # 有效統計日數 = min(90, max(上市至參考日天數, sold_days_floor))，舊品恆為 90
    # （分母修正：新品上市不足 90 日時，固定 ÷90 會按比例低估日均銷量）。
    # 有銷量或 90 天內新品時含最低 3 件保底（沿用現行 OOS 底線保底），
    # 舊品零銷量（P90D ≤ 0）時維持 0，不保底
    # （負數 P90D 按 0 計算；其他 Nature 值歸 Non Tourist 天數）。
    _days_arr = np.where(
        shop_nature.values == "tourist", _ssr_days.get("Tourist", 10),
        np.where(shop_nature.values == "mix", _ssr_days.get("Mix", 9),
                 _ssr_days.get("Non Tourist", 7)),
    ).astype(float)
    _sold_p90d = _safe_numeric(df.get("Sold Qty (P90D)", _safe_default_series(0, df.index))).values
    _sold_p90d = np.maximum(_sold_p90d, 0.0)
    _sold_days = _sold_days_denominator(df, _ref_early, sold_days_floor)
    _turn_preclip = np.ceil(_sold_p90d / _sold_days * _days_arr)
    _has_sold = (_sold_p90d > 0) | _is_new_early
    _turn_values = np.where(_has_sold, np.maximum(_turn_preclip, 3), 0).astype(int)

    df["Overstock Cap (X)"] = x_values
    df["OOS Floor (Y)"] = y_values
    df[_SSR_SAFETY_DAYS_COLUMN] = _days_arr.astype(int)
    df[_SSR_SAFETY_LINE_COLUMN] = _turn_values

    if safety_basis == _SSR_SAFETY_BASIS_TURNDAY:
        combined = _turn_values
        combined_pre = _turn_preclip
    else:
        # 月銷倍率基準：建議值 = max(Y, Z if Z ≤ X else X)，其中
        # X = Overstock Cap（庫存上限，RAW「Overstock Cap (X)」欄），
        # Y = OOS Floor（缺貨底線，RAW「OOS Floor (Y)」欄，含最低 3 件保底），
        # Z = 現有 Safety 設定（來源檔案的 Adj. Safety Stock）。
        combined = np.where(z_values <= x_values, np.maximum(y_values, z_values), np.maximum(y_values, x_values))
        # 同一公式、改用未保底的 y_preclip：代表「沒有最低 3 件保底時」的建議值
        combined_pre = np.where(z_values <= x_values, np.maximum(y_preclip, z_values), np.maximum(y_preclip, x_values))

    delta = combined - z_values
    if safety_mode == "increase_only":
        # 只可增加：delta 鉗制 ≥ 0，一般調整列只增不減（保底）
        delta = delta.clip(min=0)
        recommend_safety = z_values + delta
    else:
        recommend_safety = combined
    # Override: ND items must have Safety Stock = 0 (ND = no replenishment)
    # Override: Overstock (P90D 0 Sales) items must have Safety Stock = 0 and be recommended ND
    rp_type = (
        df.get("RP Type", _safe_default_series("", df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    is_nd = rp_type == "ND"
    is_p90d_0 = status_col == "Overstock (P90D 0 Sales)"

    # 新品保護：Launch Date 距參考日期 90 天內（≈3 個月）視為 New SKU。新品未出單是上架
    # 時間問題而非滯銷證據，排除「P90D 0 Sales → Safety=0/ND」override，改走正常
    # 基準計算（月銷倍率 X/Y 或周轉天數）；Launch Date 缺欄位、缺值或無法解析時一律視為非新品（維持原行為）。
    # 參考日期預設為今天，建議傳入 MTD 截止日期（與報表資料基準一致）。
    # 沿用前段（Y／周轉所需量保底）已算好的新品判定，確保兩處定義一致。
    is_new_sku = _is_new_early
    is_p90d_0_override = is_p90d_0 & ~is_new_sku

    if safety_mode != "increase_only":
        # 「可加可減」才套用退場 override —「只可增加」模式下任何列都不出現扣減：
        # ND / P90D 0 Sales 列保持原有 Safety，退場決策留待可加可減模式
        if is_nd.any() | is_p90d_0_override.any():
            recommend_safety = recommend_safety.copy()
            delta = delta.copy()
        if is_nd.any():
            recommend_safety[is_nd.values] = 0
            delta[is_nd.values] = -z_values[is_nd.values]
        if is_p90d_0_override.any():
            recommend_safety[is_p90d_0_override.values] = 0
            delta[is_p90d_0_override.values] = -z_values[is_p90d_0_override.values]
    elif is_nd.any():
        # 「只可增加」下 ND 列一律不作上調 — ND＝不執行自動補貨，Safety 不能 > 0 件，
        # 故不給出增加建議（一般 ND 列現有設定為 0，維持 0＝不出現任何調整建議，
        # 亦不會進入店舖問題 SKU）；ND 列亦不在此模式歸零，退場決策留待可加可減模式。
        recommend_safety = recommend_safety.copy()
        delta = delta.copy()
        recommend_safety[is_nd.values] = z_values[is_nd.values]
        delta[is_nd.values] = 0

    df["Recommend Safety Stock"] = recommend_safety
    df["Safety Qty Change"] = delta

    # ── 指定總量分配（Target Safety）：只覆寫被指定的 SKU，其餘列不動 ──
    # 分配值直接取代該 SKU 的建議值，是「公司指令」而非公式結果，故不套用
    # Overstock Cap／OOS Floor 上限，亦不受「只可增加」限制（否則指定值永遠達不到）；
    # 落差改為逐列寫入「Target 分配提示」供審批核對。
    _target_mask = None
    if target_safety:
        _target_map = (
            target_safety if isinstance(target_safety, dict)
            else _ssr_normalize_target_map(target_safety)
        )
        if _target_map:
            (recommend_safety, delta,
             _target_col, _share_col, _note_col, _target_info) = _apply_target_safety(
                df, recommend_safety, delta, _target_map, x_values, y_values
            )
            df["Recommend Safety Stock"] = recommend_safety
            df["Safety Qty Change"] = delta
            _target_mask = (~pd.isna(pd.Series(_target_col, index=df.index))).to_numpy()
            # 一個 SKU 都對不上時不加空欄（避免 RAW 出現三欄全空）
            if _target_mask.any():
                df[_SSR_TARGET_TOTAL_COLUMN] = _target_col
                df[_SSR_TARGET_SHARE_COLUMN] = _share_col
                df[_SSR_TARGET_NOTE_COLUMN] = _note_col
            df.attrs["ssr_target_info"] = _target_info

    # 保底影響偵測（供店舖問題 SKU 建議方法加註）：最終建議值停在 3 件，且未套最低
    # 3 件保底時的本來建議值 < 3 → 建議值是因保底而停在 3（保障自動補貨）。
    # 「只可增加」模式的最終值 = max(現有設定, 未保底建議值)，故 recommend_pre 亦取兩者較大。
    recommend_pre = (
        combined_pre if safety_mode != "increase_only" else np.maximum(z_values, combined_pre)
    )
    df["_safety_floor3"] = (recommend_safety == 3) & (recommend_pre < 3)
    if _target_mask is not None and _target_mask.any():
        # 指定總量分配的列，其建議值由公司指定而非保底決定 → 不標示保底影響
        df.loc[_target_mask, "_safety_floor3"] = False

    # Safety Qty Change 分類（增加 / 減少 / 不變）—「只可增加」模式下不會出現「減少」
    # （唯一例外：指定總量分配的列，以公司指定值為準，可能低於現有 Safety）
    df["Safety Qty Change 分類"] = np.where(delta > 0, "增加", np.where(delta < 0, "減少", "不變"))

    # 建議 Safety Stock 的周轉日數，與來源「Turnday (P90D Qty)」同一準則（P90D 日均）：
    # = Recommend Safety Stock ÷ (Shop Monthly Sold Qty (P90D) ÷ 30)。
    # P90D 月均 = 0 時無周轉意義（來源為 inf），此處留空；Recommend = 0 顯示 0（零覆蓋）。
    rec_turnday = np.full(len(df), np.nan)
    _monthly_arr = monthly_sold.values
    _monthly_pos = _monthly_arr > 0
    rec_turnday[_monthly_pos] = recommend_safety[_monthly_pos] / (_monthly_arr[_monthly_pos] / 30.0)
    df["Recommend Safety Stock Turnday"] = np.round(rec_turnday, 1)

    if safety_mode == "increase_only":
        # 只可增加：無歸零 / ND 退場建議，ND 欄僅保留 New SKU 資訊標記
        # （ND 列維持「0」——ND 規則不受新品保護影響，與「可加可減」同一口徑）
        nd_rec = np.full(len(df), "", dtype=object)
        if is_new_sku.any():
            nd_rec[is_new_sku & ~is_nd.values] = "New SKU (3 Month)"
    else:
        nd_rec = _build_nd_rec_qty_nd(is_nd, is_p90d_0_override, len(df))
        # New SKU 標記：ND 列維持「0」（ND 規則不受新品保護影響），其餘新品列標記
        if is_new_sku.any():
            nd_rec = nd_rec.copy()
            nd_rec[is_new_sku & ~is_nd.values] = "New SKU (3 Month)"
        # ND 轉換建議依 Adj. MCH1 / Adj. SST 加註 MAKE UP / Private Label
        nd_rec = _annotate_nd_recommendation(nd_rec, df)
    if _target_mask is not None and _target_mask.any():
        # 指定總量分配的列由公司指定補貨量，不再同時給「轉 ND／歸零」建議
        # （兩者互相矛盾）；New SKU 標記保留。清空後該列在 Summary 歸「（空白）維持」
        nd_rec = np.asarray(nd_rec, dtype=object).copy()
        for _i in np.flatnonzero(_target_mask):
            _val = str(nd_rec[_i])
            if _val in ("0", "ND") or _val.startswith("ND("):
                nd_rec[_i] = ""
    df["Recommendation of new Safety Qty / ND"] = nd_rec
    df["Buyer reply"] = ""

    return df



def _build_combined_by_shop_summary(raw):
    df = raw.copy()

    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
    shop_name = _cached_shop_names(item_shop, site_code)

    df["Site Code"] = site_code
    df["Shop Name"] = shop_name
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)

    status_col = df["Shop Stock Status"].fillna("").astype(str).str.strip()
    safety_change = pd.to_numeric(df.get("Safety Qty Change", _safe_default_series(0, df.index)), errors="coerce").fillna(0)

    group_cols = ["Site Code", "Shop Name", "OM"]
    df["_is_overstock"] = status_col.str.startswith("Overstock", na=False)
    df["_is_oos"] = (status_col == "OOS (0 Stock)") | (status_col == "Potential OOS (<7 Turnday)")
    df["_sku_overstock"] = df["SKU"].where(df["_is_overstock"])
    df["_sku_oos"] = df["SKU"].where(df["_is_oos"])
    df["_safety_increase"] = (safety_change > 0).astype(int)
    df["_safety_decrease"] = (safety_change < 0).astype(int)
    df["_safety_change"] = safety_change

    by_shop = (
        df.groupby(group_cols, dropna=False, observed=True)
        .agg(
            Overstock_SKU數量=("_sku_overstock", "nunique"),
            OOS_SKU數量=("_sku_oos", "nunique"),
            需增加Safety_SKU數量=("_safety_increase", "sum"),
            需減少Safety_SKU數量=("_safety_decrease", "sum"),
            安全庫存調整總量=("_safety_change", "sum"),
            Combined_SKU總數=("SKU", "nunique"),
        )
        .reset_index()
    )

    for col in by_shop.columns:
        if col not in group_cols:
            by_shop[col] = by_shop[col].fillna(0).astype(int)

    if not by_shop.empty:
        total_row = pd.DataFrame(
            {
                "Site Code": ["總計"],
                "Shop Name": [""],
                "OM": [""],
                "Overstock_SKU數量": [int(df["_sku_overstock"].dropna().nunique())],
                "OOS_SKU數量": [int(df["_sku_oos"].dropna().nunique())],
                "需增加Safety_SKU數量": [int(by_shop["需增加Safety_SKU數量"].sum())],
                "需減少Safety_SKU數量": [int(by_shop["需減少Safety_SKU數量"].sum())],
                "安全庫存調整總量": [int(by_shop["安全庫存調整總量"].sum())],
                "Combined_SKU總數": [int(df["SKU"].dropna().nunique())],
            }
        )
        by_shop = pd.concat([by_shop.sort_values(["Site Code", "Shop Name", "OM"], kind="stable"), total_row], ignore_index=True)

    return by_shop



_CHANGE_CAT_DESC = {
    "增加": "建議 Safety Stock 高於現時",
    "減少": "建議 Safety Stock 低於現時（「可加可減」模式，或指定總量分配的列）",
    "不變": "建議 Safety Stock 與現時相同",
}



_ND_REC_DISPLAY = {
    "": "(空白) 維持 Safety Stock",
    "New SKU (3 Month)": "New SKU (3 Month)",
    "0": "0（已是 ND）",
    "ND": "ND",
    "ND(02 - MAKE UP)": "ND(02 - MAKE UP)",
    "ND(Private Label)": "ND(Private Label)",
    "ND(Private Label - MAKE UP)": "ND(Private Label - MAKE UP)",
}



_ND_REC_DESC = {
    "": "無 ND 建議，按建議 Safety Stock 繼續運作",
    "New SKU (3 Month)": "上市 90 天內新品，保留觀察、不套用歸零/ND",
    "0": "RP Type 已是 ND，建議 Safety = 0",
    "ND": "建議轉 ND 停補貨（P90D 0 Sales 舊品）",
    "ND(02 - MAKE UP)": "建議轉 ND — Adj. MCH1 = 02 - MAKE UP",
    "ND(Private Label)": "建議轉 ND — Adj. SST 屬 03 - Private Label 系",
    "ND(Private Label - MAKE UP)": "建議轉 ND — MAKE UP + Private Label",
}



def _build_combined_classification_summary(raw):
    """SUMMARY sheet 資料 — 兩張分類統計表：
    ① Safety Qty Change 分類（增加 / 減少 / 不變）
    ② Recommendation of new Safety Qty / ND 分類（空白 / New SKU / 0 / ND 系列）
    每列統計 行數 與 SKU 數；同一 SKU 可跨分類出現（不同店舖），總計列的
    SKU 數改用全表不重複 SKU 數，避免跨分類加總虛增。
    """
    sku = raw["SKU"] if "SKU" in raw.columns else _safe_default_series("", raw.index)
    change = pd.to_numeric(raw.get("Safety Qty Change", _safe_default_series(0, raw.index)), errors="coerce").fillna(0)
    change_cat = np.where(change > 0, "增加", np.where(change < 0, "減少", "不變"))
    total_sku = int(sku.dropna().nunique())

    change_rows = []
    for label in ("增加", "減少", "不變"):
        m = change_cat == label
        change_rows.append({
            "分類": label,
            "行數": int(m.sum()),
            "SKU 數": int(sku[m].dropna().nunique()),
            "Safety Qty Change 合計": int(change[m].sum()),
            "說明": _CHANGE_CAT_DESC[label],
        })
    change_rows.append({
        "分類": "總計",
        "行數": int(len(raw)),
        "SKU 數": total_sku,
        "Safety Qty Change 合計": int(change.sum()),
        "說明": "SKU 數為全表不重複 SKU（同一 SKU 可跨分類）",
    })
    change_summary = pd.DataFrame(change_rows)

    nd_rec = (
        raw.get("Recommendation of new Safety Qty / ND", _safe_default_series("", raw.index))
        .fillna("").astype(str)
    )
    known = list(_ND_REC_DISPLAY.keys())
    present_order = known + sorted(v for v in nd_rec.unique() if v not in known)
    nd_rows = []
    for val in present_order:
        m = (nd_rec == val).values
        nd_rows.append({
            "建議分類": _ND_REC_DISPLAY.get(val, val),
            "行數": int(m.sum()),
            "SKU 數": int(sku[m].dropna().nunique()),
            "說明": _ND_REC_DESC.get(val, ""),
        })
    nd_rows.append({
        "建議分類": "總計",
        "行數": int(len(raw)),
        "SKU 數": total_sku,
        "說明": "SKU 數為全表不重複 SKU（同一 SKU 可跨分類）",
    })
    nd_summary = pd.DataFrame(nd_rows)

    return change_summary, nd_summary



def _build_combined_target_summary(raw):
    """SUMMARY 區塊 ④：指定總量分配（Target Safety）逐 SKU 對帳。

    只列有指定總量的 SKU；「參與店舖數」只計實際獲分配的非 ND 列（ND 店
    Safety 必須為 0，不參與分配）。「差額」理論上恆為 0（最大餘數法保證），
    保留作稽核；「超出庫存上限／低於缺貨底線店數」反映指定值與 X／Y 的落差，
    讓審批人一眼看出哪些店舖的分配值偏離公式範圍。未啟用時回傳空表。
    """
    columns = [
        "SKU", "Target 全線合計", "已分配", "差額", "參與店舖數",
        "超出庫存上限店數", "低於缺貨底線店數",
    ]
    empty = pd.DataFrame(columns=columns)
    if raw is None or raw.empty or _SSR_TARGET_TOTAL_COLUMN not in raw.columns:
        return empty
    target = pd.to_numeric(raw[_SSR_TARGET_TOTAL_COLUMN], errors="coerce")
    mask = target.notna()
    if not mask.any():
        return empty

    sub = raw.loc[mask]
    rec = pd.to_numeric(sub.get("Recommend Safety Stock", _safe_default_series(0, sub.index)), errors="coerce").fillna(0)
    x_ser = pd.to_numeric(sub.get("Overstock Cap (X)", _safe_default_series(np.nan, sub.index)), errors="coerce")
    y_ser = pd.to_numeric(sub.get("OOS Floor (Y)", _safe_default_series(np.nan, sub.index)), errors="coerce")
    sub = sub.assign(_rec=rec, _x=x_ser, _y=y_ser)

    rows = []
    for sku, part in sub.groupby("SKU", sort=False, dropna=False):
        _target = pd.to_numeric(part[_SSR_TARGET_TOTAL_COLUMN], errors="coerce").max()
        _alloc = float(part["_rec"].sum())
        rows.append({
            "SKU": sku,
            "Target 全線合計": int(_target) if pd.notna(_target) else 0,
            "已分配": int(round(_alloc)),
            "差額": int(round(float(_target) - _alloc)) if pd.notna(_target) else 0,
            "參與店舖數": int(len(part)),
            "超出庫存上限店數": int(((part["_x"].notna()) & (part["_rec"] > part["_x"])).sum()),
            "低於缺貨底線店數": int(((part["_y"].notna()) & (part["_rec"] < part["_y"])).sum()),
        })

    out = pd.DataFrame(rows, columns=columns)
    if out.empty:
        return empty
    total_row = pd.DataFrame([{
        "SKU": "總計",
        "Target 全線合計": int(out["Target 全線合計"].sum()),
        "已分配": int(out["已分配"].sum()),
        "差額": int(out["差額"].sum()),
        "參與店舖數": int(out["參與店舖數"].sum()),
        "超出庫存上限店數": int(out["超出庫存上限店數"].sum()),
        "低於缺貨底線店數": int(out["低於缺貨底線店數"].sum()),
    }])
    return pd.concat([out, total_row], ignore_index=True)



def _build_combined_dimension_cross_tab(raw, dimension="nature", om_groups=None):
    """Safety Qty Change 分類 × 分類維度交叉表（SUMMARY 區塊 ③）。

    Safety Stock Review 引入分類維度用：只改分組顯示，不影響任何 Safety 建議計算。
    分組鍵與混合模式同一準則（_hybrid_group_key_series）；om_groups 未提供時由 raw 即時推導；
    RAW 缺「銷量帶」欄時銷量兩個維度退回 Shop Nature 鍵（與 hybrid 同一守衛）。
    每格統計 列數 / SKU 數 / Safety Qty Change 合計；總計列的 SKU 數為全表不重複 SKU
    （同一 SKU 可跨分類，與區塊 ① 同一準則）；非標準 Shop Nature 值不佔分組欄、只計入合計。
    """
    dim = _hybrid_dimension(dimension)
    first_col = "Safety Qty Change 分類"
    if raw is None or raw.empty or "Safety Qty Change" not in raw.columns:
        groups = _hybrid_dimension_groups(dim, om_groups=om_groups)
        cols = [first_col]
        for g in groups:
            cols += [f"{g} 列數", f"{g} SKU 數", f"{g} Safety Qty Change 合計"]
        cols += ["合計 列數", "合計 SKU 數", "合計 Safety Qty Change"]
        return pd.DataFrame(columns=cols)
    if dim in ("om", "nature_om") and om_groups is None:
        om_groups = _hybrid_om_groups_from_values(
            raw["OM"] if "OM" in raw.columns else _safe_default_series("", raw.index)
        )
    groups = _hybrid_dimension_groups(dim, om_groups=om_groups)
    cols = [first_col]
    for g in groups:
        cols += [f"{g} 列數", f"{g} SKU 數", f"{g} Safety Qty Change 合計"]
    cols += ["合計 列數", "合計 SKU 數", "合計 Safety Qty Change"]
    chg = pd.to_numeric(raw["Safety Qty Change"], errors="coerce").fillna(0)
    grp = _hybrid_group_key_series(raw, dim, om_groups=om_groups)
    sku = raw["SKU"] if "SKU" in raw.columns else _safe_default_series("", raw.index)
    total_sku = int(sku.dropna().nunique())
    _total_keys = {
        "列數": "合計 列數",
        "SKU 數": "合計 SKU 數",
        "Safety Qty Change 合計": "合計 Safety Qty Change",
    }

    def _agg(sub_chg: pd.Series, sub_sku: pd.Series) -> dict:
        return {
            "列數": int(len(sub_chg)),
            "SKU 數": int(sub_sku.dropna().nunique()),
            "Safety Qty Change 合計": int(round(float(sub_chg.sum()))),
        }

    def _fill(row: dict, mask: pd.Series, total_row: bool = False) -> dict:
        for g in groups:
            gm = mask & (grp == g).to_numpy()
            for k, v in _agg(chg[gm], sku[gm]).items():
                row[f"{g} {k}"] = v
        agg = _agg(chg[mask], sku[mask])
        if total_row:
            agg["SKU 數"] = total_sku
        for k, v in agg.items():
            row[_total_keys[k]] = v
        return row

    cat_masks = [
        ("增加", chg > 0),
        ("減少", chg < 0),
        ("不變", chg == 0),
    ]
    rows = [_fill({first_col: label}, mask.to_numpy()) for label, mask in cat_masks]
    rows.append(_fill({first_col: "總計"}, np.ones(len(raw), dtype=bool), total_row=True))
    return pd.DataFrame(rows, columns=cols)



# ── SUMMARY 區塊 ⑤–⑦：影響範圍統計與樞紐分析式統計 ──
# 三個區塊一律以「店舖問題 SKU」工作表同一批需行動列、同一「問題類型」為準，令 SUMMARY
# 與該表（以及使用者自行以該表做成的樞紐分析表）數字一致。問題類型由該表判定：
#   Safety Qty Change > 0 → Safety 設定不足；< 0 → Safety 設定過高；
#   = 0 而有 ND 建議或已是 ND → ND 退場。其餘列（不變且無 ND 建議）不屬任何問題類型。
# ⑤ 以 Brand／Buyer／店舖為單位回答「多少個」有問題；⑥⑦ 仿樞紐分析表版面列出明細：
# 列＝Buyer（小計）＞Brand，或 Site Code／Shop Name；欄＝問題類型 × 量值（行數／
# 加總 - Safety Qty Change），最後一組為總計。
_SSR_ISSUE_TYPES = ("ND 退場", "Safety 設定不足", "Safety 設定過高")
_SSR_ISSUE_COUNT = "行數"
_SSR_ISSUE_SUM = "加總 - Safety Qty Change"
_SSR_PIVOT_TOTAL = "總計"
_SSR_PIVOT_LEVEL_COL = "層級"
_SSR_PIVOT_BLANK = "（空白）"
_SSR_PIVOT_LABELS = {"buyer_brand": ("Buyer / Brand",), "shop": ("Site Code", "Shop Name")}
# SUMMARY 工作表欄寬上限（A 欄為列標籤，放寬少許以容納較長的 Buyer／Brand 名稱）
_SSR_SUMMARY_COL_MAX = 30
_SSR_SUMMARY_LABEL_COL_MAX = 36
# 計算規則寫在區塊標題；每列說明只補該對象的計法，並保持簡短，以免撐闊 ⑥⑦ 的數值欄
_SSR_IMPACT_ENTITIES = (
    ("Brand", "Brand", "不重複 Brand"),
    ("Buyer", "Buyer", "不重複 Buyer"),
    ("Site Code", "店舖", "不重複店舖；以 Site Code 計"),
)
_SSR_IMPACT_COLUMNS = ["統計對象", "總數", *_SSR_ISSUE_TYPES, "無需行動", "說明"]


def _ssr_entity_keys(raw) -> pd.DataFrame:
    """RAW 的 Brand／Buyer／Site Code 鍵（去首尾空白；缺欄為空字串）；不修改 raw。

    一律以位置陣列建立（新的 RangeIndex），不依賴 raw 的 index——重複或非整數 index
    亦不影響對齊。
    """
    def _key(col):
        if col not in raw.columns:
            return np.full(len(raw), "", dtype=object)
        return raw[col].fillna("").astype(str).str.strip().to_numpy(dtype=object)

    return pd.DataFrame({c: _key(c) for c in ("Brand", "Buyer", "Site Code")})


def _build_combined_impact_summary(raw, issues=None):
    """SUMMARY 區塊 ⑤：影響範圍統計（Brand／Buyer／店舖 × 問題類型）。

    每欄為「有任何一列屬該問題類型」的不重複對象數（同一對象可同時計入多欄，故
    各欄相加不等於總數）；「無需行動」＝在「店舖問題 SKU」中沒有任何列的對象數。
    問題類型沿用該表（issues，即 `_build_combined_shop_issues` 的輸出）的判定，
    未傳時即時建立。空白鍵的列不計入對象數；店舖以 Site Code 計（與 By Shop 同一鍵）。
    """
    if issues is None:
        issues = _build_combined_shop_issues(raw)
    keys = _ssr_entity_keys(raw)
    has_issues = issues is not None and not issues.empty and "問題類型" in issues.columns
    types = (
        issues["問題類型"].fillna("").astype(str).str.strip() if has_issues
        else pd.Series([], dtype=object)
    )
    rows = []
    for col, label, desc in _SSR_IMPACT_ENTITIES:
        valid_keys = keys[col][keys[col] != ""]
        total = int(valid_keys.nunique())
        row = {"統計對象": label, "總數": total, "說明": desc, **dict.fromkeys(_SSR_ISSUE_TYPES, 0)}
        touched = 0
        if has_issues and col in issues.columns:
            ikeys = issues[col].fillna("").astype(str).str.strip()
            keep = ikeys != ""
            touched = int(ikeys[keep].nunique())
            for t in _SSR_ISSUE_TYPES:
                row[t] = int(ikeys[keep & (types == t).to_numpy()].nunique())
        row["無需行動"] = total - touched
        rows.append(row)
    return pd.DataFrame(rows, columns=_SSR_IMPACT_COLUMNS)


def _ssr_issue_frame(issues) -> pd.DataFrame:
    """問題類型樞紐的底表：識別欄（去首尾空白，空白顯示為「（空白）」）、問題類型與 Safety Qty Change。

    一律以位置陣列取值（新的 RangeIndex），不依賴 issues 的 index。
    """
    n = len(issues)
    out = pd.DataFrame(index=pd.RangeIndex(n))
    for c in ("Buyer", "Brand", "Site Code", "Shop Name"):
        vals = (
            issues[c].fillna("").astype(str).str.strip().to_numpy(dtype=object)
            if c in issues.columns else np.full(n, "", dtype=object)
        )
        s = pd.Series(vals)
        out[c] = s.mask(s == "", _SSR_PIVOT_BLANK).to_numpy()
    out["問題類型"] = issues["問題類型"].fillna("").astype(str).str.strip().to_numpy()
    out["_chg"] = pd.to_numeric(
        issues.get("Safety Qty Change", _safe_default_series(0, issues.index)), errors="coerce"
    ).fillna(0).to_numpy(dtype=float)
    return out


def _ssr_issue_wide(frame, keys):
    """依 keys 彙總每個問題類型的行數與 Safety Qty Change 加總，另加總計；數值為 Python int。

    keys 為空時回傳只有一列的全表合計。不屬於 `_SSR_ISSUE_TYPES` 的問題類型只計入總計。
    """
    columns = [
        f"{t} {m}" for t in _SSR_ISSUE_TYPES + (_SSR_PIVOT_TOTAL,) for m in (_SSR_ISSUE_COUNT, _SSR_ISSUE_SUM)
    ]
    if frame.empty:
        return pd.DataFrame([[0] * len(columns)], columns=columns)
    keyed = bool(keys)
    if not keyed:
        # 全表合計：借用同一條 groupby 路徑（以固定鍵分組），確保行數與加總的 index 一致
        frame = frame.assign(_all=0)
        keys = ["_all"]
    g = frame.groupby([*keys, "問題類型"], sort=False, observed=True, dropna=False)["_chg"].agg(["size", "sum"])
    n = g["size"].unstack(fill_value=0)
    s = g["sum"].unstack(fill_value=0)
    out = pd.DataFrame(index=n.index)

    def _col(x):
        return x if isinstance(x, pd.Series) else pd.Series(x, index=out.index, dtype=int)

    for t in _SSR_ISSUE_TYPES + (_SSR_PIVOT_TOTAL,):
        if t == _SSR_PIVOT_TOTAL:
            cnt, tot = n.sum(axis=1), s.sum(axis=1)
        else:
            cnt = n[t] if t in n.columns else 0
            tot = s[t] if t in s.columns else 0
        out[f"{t} {_SSR_ISSUE_COUNT}"] = _col(cnt)
        out[f"{t} {_SSR_ISSUE_SUM}"] = np.round(_col(tot).astype(float)).astype(int)
    return out if keyed else out.reset_index(drop=True)


def _build_combined_issue_pivot(issues, by="buyer_brand"):
    """SUMMARY 區塊 ⑥（by="buyer_brand"）／⑦（by="shop"）：「店舖問題 SKU」的樞紐分析。

    issues 為 `_build_combined_shop_issues` 的輸出。by="buyer_brand"：每個 Buyer 一列小計
    （層級＝Buyer），其下按 Brand 分列明細（層級＝Brand），依名稱排序；by="shop"：每個
    Site Code 一列（Shop Name 取該店第一個可解析名稱），依 Site Code 排序。只列出有
    需行動列的對象，末列為總計。回傳扁平欄名的表（「{問題類型} {量值}」），
    Excel 寫入時再還原成兩層表頭。
    """
    label_cols = _SSR_PIVOT_LABELS[by]
    measure_cols = [
        f"{t} {m}" for t in _SSR_ISSUE_TYPES + (_SSR_PIVOT_TOTAL,) for m in (_SSR_ISSUE_COUNT, _SSR_ISSUE_SUM)
    ]
    columns = [*label_cols, _SSR_PIVOT_LEVEL_COL, *measure_cols]
    has_issues = issues is not None and not issues.empty and "問題類型" in issues.columns
    base = _ssr_issue_frame(issues) if has_issues else pd.DataFrame(columns=["問題類型", "_chg"])
    records = []
    if has_issues and by == "buyer_brand":
        brand_w = _ssr_issue_wide(base, ["Buyer", "Brand"]).reset_index()
        for b in _ssr_issue_wide(base, ["Buyer"]).reset_index().sort_values("Buyer").to_dict("records"):
            records.append({label_cols[0]: b["Buyer"], _SSR_PIVOT_LEVEL_COL: "Buyer", **{c: b[c] for c in measure_cols}})
            sub = brand_w[brand_w["Buyer"] == b["Buyer"]].sort_values("Brand")
            for r in sub.to_dict("records"):
                records.append({label_cols[0]: r["Brand"], _SSR_PIVOT_LEVEL_COL: "Brand", **{c: r[c] for c in measure_cols}})
    elif has_issues:
        names = base[base["Shop Name"] != _SSR_PIVOT_BLANK].groupby("Site Code")["Shop Name"].first()
        for r in _ssr_issue_wide(base, ["Site Code"]).reset_index().sort_values("Site Code").to_dict("records"):
            records.append({
                label_cols[0]: r["Site Code"],
                label_cols[1]: names.get(r["Site Code"], _SSR_PIVOT_BLANK),
                _SSR_PIVOT_LEVEL_COL: "店舖",
                **{c: r[c] for c in measure_cols},
            })
    total = _ssr_issue_wide(base, []).iloc[0]
    records.append({
        **dict.fromkeys(label_cols, ""),
        label_cols[0]: _SSR_PIVOT_TOTAL,
        _SSR_PIVOT_LEVEL_COL: _SSR_PIVOT_TOTAL,
        **{c: total[c] for c in measure_cols},
    })
    return pd.DataFrame(records, columns=columns)


def _write_issue_pivot_block(ws, start_row, title, table, by, thin_border):
    """SUMMARY 區塊 ⑥／⑦：仿樞紐分析表版面寫入問題類型樞紐；回傳下一個可用列號。

    兩層表頭（問題類型跨兩欄置中，下為 行數／加總 - Safety Qty Change）；Buyer 小計列
    黃底粗體，其下 Brand 明細列入摺疊群組（Excel 左側可按「−」收合）；總計欄淺黃底、
    總計列淺藍底粗體；沒有列的問題類型留白（與樞紐分析表相同）。店舖表（by="shop"）
    另設自動篩選——Excel 每張工作表只容許一個篩選範圍，故只設在列數較多的店舖表，
    範圍由表頭至最後一間店舖（總計列不在內，篩選時仍會顯示）。
    """
    header_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    subtotal_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    total_col_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    section_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    bold_font = Font(bold=True, size=11)
    base_font = Font(size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    # 「跨越選取範圍置中」：不合併儲存格亦可令問題類型名稱置中於兩欄之上
    group_align = Alignment(horizontal="centerContinuous", vertical="center")
    num_align = Alignment(horizontal="right", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    label_cols = _SSR_PIVOT_LABELS[by]
    n_label = len(label_cols)
    groups = _SSR_ISSUE_TYPES + (_SSR_PIVOT_TOTAL,)
    n_cols = n_label + 2 * len(groups)

    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=n_cols)
    sc = ws.cell(row=start_row, column=1, value=title)
    sc.font = Font(bold=True, size=12, color="1E40AF")
    sc.fill = section_fill
    sc.alignment = Alignment(horizontal="left", vertical="center")

    r_group, r_head = start_row + 1, start_row + 2
    for r in (r_group, r_head):
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = header_fill
            cell.border = thin_border
            cell.font = bold_font
    # 標籤欄表頭：Buyer／Brand 跨兩列置中；店舖表因設有自動篩選，表頭不可合併
    if by == "buyer_brand":
        ws.merge_cells(start_row=r_group, start_column=1, end_row=r_head, end_column=1)
        ws.cell(row=r_group, column=1, value=label_cols[0]).alignment = header_align
    else:
        for i, name in enumerate(label_cols, 1):
            ws.cell(row=r_head, column=i, value=name).alignment = header_align
    for gi, t in enumerate(groups):
        c0 = n_label + 1 + 2 * gi
        for c in (c0, c0 + 1):
            cell = ws.cell(row=r_group, column=c)
            cell.alignment = group_align
            if t == _SSR_PIVOT_TOTAL:
                cell.fill = total_col_fill
        ws.cell(row=r_group, column=c0, value=t)
        for off, m in enumerate((_SSR_ISSUE_COUNT, _SSR_ISSUE_SUM)):
            cell = ws.cell(row=r_head, column=c0 + off, value=m)
            cell.alignment = header_align
            if t == _SSR_PIVOT_TOTAL:
                cell.fill = total_col_fill
    ws.row_dimensions[r_group].height = 20
    ws.row_dimensions[r_head].height = 24

    if by == "buyer_brand":
        # 小計（Buyer）列顯示在群組上方，與樞紐分析表相同
        ws.sheet_properties.outlinePr.summaryBelow = False
    r = r_head + 1
    first_data = r
    for rec in table.to_dict("records"):
        level = rec.get(_SSR_PIVOT_LEVEL_COL)
        is_sub, is_total = level == "Buyer", level == _SSR_PIVOT_TOTAL
        for i, name in enumerate(label_cols, 1):
            cell = ws.cell(row=r, column=i, value=rec.get(name) or None)
            cell.border = thin_border
            cell.font = bold_font if (is_sub or is_total) else base_font
            if level == "Brand":
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            else:
                cell.alignment = left_align
        for gi, t in enumerate(groups):
            # 沒有列的問題類型兩格留白；總計一律顯示
            blank = t != _SSR_PIVOT_TOTAL and not int(rec.get(f"{t} {_SSR_ISSUE_COUNT}", 0) or 0)
            for off, m in enumerate((_SSR_ISSUE_COUNT, _SSR_ISSUE_SUM)):
                cell = ws.cell(
                    row=r, column=n_label + 1 + 2 * gi + off,
                    value=None if blank else int(rec.get(f"{t} {m}", 0) or 0),
                )
                cell.border = thin_border
                cell.number_format = "#,##0"
                cell.alignment = num_align
                cell.font = bold_font if (is_sub or is_total) else base_font
                if t == _SSR_PIVOT_TOTAL:
                    cell.fill = total_col_fill
        if is_sub or is_total:
            row_fill = subtotal_fill if is_sub else header_fill
            for c in range(1, n_cols + 1):
                ws.cell(row=r, column=c).fill = row_fill
        if level == "Brand":
            ws.row_dimensions[r].outlineLevel = 1
        r += 1
    last_data = r - 2
    if by == "shop" and last_data >= first_data:
        ws.auto_filter.ref = f"A{r_head}:{get_column_letter(n_cols)}{last_data}"
    return r  # 末列的下一列（與 _write_block 相同：呼叫端再加一列空白）



def transform_combined(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
    rates_overstock: dict | None = None,
    rates_oos: dict | None = None,
    mtd_days: int | None = None,
    month_total_days: int | None = None,
    safety_mode: str = "increase_only",
    ref_date=None,
    sales_oos_days=None,
    sales_overstock_days=None,
    sales_overstock_min_units=None,
    safety_basis: str = "rate",
    safety_days: dict | None = None,
    sold_days_floor: int | None = None,
    target_safety=None,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str], list[str]]:
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    warnings = []
    safety_basis = _normalize_ssr_safety_basis(safety_basis)
    _ssr_days, _ssr_day_warnings = _hybrid_safety_days(safety_days)
    warnings.extend(_ssr_day_warnings)
    _target_map = _ssr_normalize_target_map(target_safety)

    # Safety 來源欄相容檢查：顯示名一律視為來源名（Safety Stock (現時) = 來源檔案的 Adj. Safety Stock）
    _adj_avail = set(df_input.columns) | {_resolve_column_name(_ADJ_SAFETY_SOURCE_COL, set(df_input.columns))}
    _optional_cols = ["Adj. Safety Stock", "Sold Qty (CM - MTD)", "Shop Monthly Sold Qty (P90D)", "Shop Nature"]
    for _col in _optional_cols:
        _label = _display_safety_header(_col)
        if _resolve_column_name(_col, _adj_avail) is None:
            warnings.append(f"缺少可選欄位「{_label}」，建議值將可能不準確")

    if progress_callback:
        progress_callback(0.05, "驗證通過")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據（非空白或總計列）")

    df["SKU"] = df["SKU"].apply(_safe_sku)

    if "SKU Desc" not in df.columns:
        df = _compute_raw_extra_columns(df)

    raw = _add_combined_recommendation_columns(
        df, rates_overstock, rates_oos, mtd_days, month_total_days, safety_mode=safety_mode,
        ref_date=ref_date, safety_basis=safety_basis, safety_days=_ssr_days,
        sold_days_floor=sold_days_floor, target_safety=_target_map or None,
    )
    _target_info = dict(raw.attrs.get("ssr_target_info") or {})

    if _target_map:
        _matched = _target_info.get("matched") or []
        _target_total = sum(int(_target_map[k]) for k in _matched)
        if _matched:
            warnings.append(
                f"已套用指定總量分配：{len(_matched)} 個 SKU（合計 {_target_total:,} 件），"
                f"分配基準為{_SSR_TARGET_BASIS_LABEL}，分配總和等於指定值；"
                "未指定的 SKU 一律按上方計算基準計算。"
            )
        if _target_info.get("missing"):
            _miss = "、".join(str(k) for k in _target_info["missing"])
            warnings.append(
                f"以下指定 SKU 在本次資料中沒有可分配的非 ND 店舖，未套用指定總量：{_miss}"
                "（請核對 SKU 寫法，或該 SKU 是否全線為 ND／已被 Shop Nature 篩選剔除）。"
            )
        if _target_info.get("below_current"):
            warnings.append(
                f"指定總量分配下有 {_target_info['below_current']} 列的分配值低於現有 Safety；"
                "此類列以公司指定值為準，不受「只可增加」限制，"
                "RAW「Safety Qty Change 分類」會出現「減少」。"
            )

    # 銷量帶／銷量存貨分類（常開顯示欄；只作分類顯示，不影響 Safety 建議計算）
    _sales_params, _sales_param_warnings = _hybrid_sales_params(
        False, sales_oos_days, sales_overstock_days, sales_overstock_min_units
    )[1:]
    warnings.extend(_sales_param_warnings)
    raw, _sales_col_warnings = _attach_sales_classification_columns(raw, _sales_params)
    warnings.extend(_sales_col_warnings)

    status_col = raw["Shop Stock Status"].fillna("")
    safety_change = pd.to_numeric(raw.get("Safety Qty Change", _safe_default_series(0, raw.index)), errors="coerce").fillna(0)

    categories = ["Overstock", "OOS", "Balanced (Safety Adjusted)"]

    group_cols = ["Buyer", "Brand", "SKU", "SKU Desc"]

    if progress_callback:
        progress_callback(0.3, "分類統計中...")

    overstock_mask = status_col.str.startswith("Overstock", na=False)
    oos_mask = (status_col == "OOS (0 Stock)") | (status_col == "Potential OOS (<7 Turnday)")
    balanced_mask = (safety_change != 0) & ~overstock_mask & ~oos_mask

    # 預計算 int 旗標（單次 groupby + 條件聚合，取代 7 次獨立 groupby）
    raw["_is_overstock"] = overstock_mask.astype(int)
    raw["_is_oos"] = oos_mask.astype(int)
    raw["_is_balanced"] = balanced_mask.astype(int)
    raw["_is_inc"] = (safety_change > 0).astype(int)
    raw["_is_dec"] = (safety_change < 0).astype(int)
    raw["_is_unchanged"] = (safety_change == 0).astype(int)

    if progress_callback:
        progress_callback(0.5, "合併彙總中...")

    summary = (
        raw.groupby(group_cols, observed=True)
        .agg(
            Overstock=("_is_overstock", "sum"),
            OOS=("_is_oos", "sum"),
            **{"Balanced (Safety Adjusted)": ("_is_balanced", "sum")},
            **{"Safety Qty Change": ("Safety Qty Change", "sum")},
            **{"增加Safety Qty的店舖數量": ("_is_inc", "sum")},
            **{"減少Safety Qty的店舖數量": ("_is_dec", "sum")},
            **{"Safety Qty不變的店舖數量": ("_is_unchanged", "sum")},
        )
        .fillna(0).astype(int).reset_index()
    )
    summary["SKU"] = summary["SKU"].apply(_safe_sku)
    summary["總計"] = summary[categories].sum(axis=1)

    buyer_order = summary.groupby("Buyer")["總計"].sum().sort_values(ascending=False).index.tolist()
    summary["_buyer_key"] = summary["Buyer"].map({b: i for i, b in enumerate(buyer_order)})
    summary = summary.sort_values(["_buyer_key", "Brand", "SKU"], ascending=[True, True, True]).drop(columns=["_buyer_key"]).reset_index(drop=True)

    total_cat = int(summary["總計"].sum())
    total_row = pd.DataFrame(
        {
            "Buyer": ["總計"],
            "Brand": [None],
            "SKU": [None],
            **{cat: [int(summary[cat].sum())] for cat in categories},
            "總計": [total_cat],
            "Safety Qty Change": [int(summary["Safety Qty Change"].sum())],
            "增加Safety Qty的店舖數量": [int(summary["增加Safety Qty的店舖數量"].sum())],
            "減少Safety Qty的店舖數量": [int(summary["減少Safety Qty的店舖數量"].sum())],
            "Safety Qty不變的店舖數量": [int(summary["Safety Qty不變的店舖數量"].sum())],
        }
    )
    summary = pd.concat([summary, total_row], ignore_index=True)

    # 清理暫存欄位
    raw.drop(columns=[c for c in ["_is_overstock", "_is_oos", "_is_balanced", "_is_inc", "_is_dec", "_is_unchanged"] if c in raw.columns], inplace=True, errors="ignore")

    if total_cat == 0:
        warnings.append("未偵測到任何符合條件的 Safety Stock Review 資料。")

    if progress_callback:
        progress_callback(0.7, "轉換完成")

    return summary, raw, categories, warnings



def _write_summary_sheet(
    ws, change_summary, nd_summary, thin_border, cross_tab=None, target_summary=None,
    impact_summary=None, buyer_brand_pivot=None, shop_pivot=None,
):
    """SUMMARY sheet — 區塊：Safety Qty Change 分類、ND 建議分類 統計表，
    另附 Safety Qty Change 分類 × 分類維度交叉表（cross_tab，Safety Stock Review 引入分類維度用）、
    指定總量分配逐 SKU 對帳（target_summary，只在有指定 SKU 時出現）、
    影響範圍統計（impact_summary）與「店舖問題 SKU」的樞紐分析
    （buyer_brand_pivot＝Buyer＞Brand；shop_pivot＝Site Code／Shop Name）."""
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    section_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    ws.merge_cells("A1:E1")
    title_cell = ws.cell(row=1, column=1, value="Safety Stock Review SUMMARY（分類統計）")
    title_cell.font = Font(bold=True, size=14, color="1E40AF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    def _write_block(start_row, section_title, block):
        cols = list(block.columns)
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=len(cols))
        sc = ws.cell(row=start_row, column=1, value=section_title)
        sc.font = Font(bold=True, size=12, color="1E40AF")
        sc.fill = section_fill
        sc.alignment = Alignment(horizontal="left", vertical="center")

        header_row = start_row + 1
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col_idx, h in enumerate(cols, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        # P0：共用對齊物件，避免逐格新建 Alignment；值與格式語意不變
        center_align = Alignment(horizontal="center", vertical="center")
        left_align = Alignment(horizontal="left", vertical="center")
        left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
        for r, row_data in enumerate(block.itertuples(index=False), header_row + 1):
            is_total = row_data[0] == "總計"
            for col_idx, v in enumerate(row_data, 1):
                cell = ws.cell(row=r, column=col_idx, value=v)
                cell.border = thin_border
                if is_total:
                    cell.font = bold_font
                    cell.fill = total_fill
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    cell.number_format = "#,##0"
                    cell.alignment = center_align
                elif col_idx == 1:
                    cell.alignment = left_align
                else:
                    cell.alignment = left_wrap
        return header_row + len(block) + 1

    next_row = _write_block(3, "① Safety Qty Change 分類（增加 / 減少 / 不變）", change_summary)
    next_row = _write_block(next_row + 1, "② Recommendation of new Safety Qty / ND 分類", nd_summary)
    if cross_tab is not None and not cross_tab.empty:
        next_row = _write_block(next_row + 1, "③ Safety Qty Change 分類 × 分類維度交叉表", cross_tab)
    if target_summary is not None and not target_summary.empty:
        next_row = _write_block(
            next_row + 1,
            "④ 指定總量分配（Target Safety）逐 SKU 對帳"
            f"（分配基準：{_SSR_TARGET_BASIS_LABEL}；差額應為 0）",
            target_summary,
        )
    if impact_summary is not None and not impact_summary.empty:
        next_row = _write_block(
            next_row + 1,
            "⑤ 影響範圍統計（Brand／Buyer／店舖）"
            "— 只計「店舖問題 SKU」的需行動列；有任何一列屬該問題類型即計入，同一對象可同時計入多欄；"
            "「無需行動」＝該表沒有任何列",
            impact_summary,
        )
    if buyer_brand_pivot is not None and not buyer_brand_pivot.empty:
        next_row = _write_issue_pivot_block(
            ws, next_row + 1,
            "⑥ 按 Buyer／Brand 統計 —「店舖問題 SKU」的樞紐分析（列＝Buyer 小計＞Brand；"
            "欄＝問題類型 × 行數／加總 - Safety Qty Change）"
            f"。問題類型：Safety 設定不足＝Safety Qty Change > 0；Safety 設定過高＝< 0；"
            "ND 退場＝Safety Qty Change 為 0 而有 ND 建議或已是 ND。"
            "Buyer 列黃底粗體，其下 Brand 可按左側「−」收合",
            buyer_brand_pivot, "buyer_brand", thin_border,
        )
    if shop_pivot is not None and not shop_pivot.empty:
        _write_issue_pivot_block(
            ws, next_row + 1,
            "⑦ 按店舖統計（Site Code／Shop Name）—「店舖問題 SKU」的樞紐分析"
            "（欄＝問題類型 × 行數／加總 - Safety Qty Change；問題類型定義見 ⑥）。表頭設有自動篩選",
            shop_pivot, "shop", thin_border,
        )
    _auto_width(ws)
    # 各區塊共用欄寬：①② 的「說明」長句與區塊標題會把 A／D／E 欄撐至 47–50，令 ⑥⑦ 的
    # 「行數／加總」欄寬窄懸殊。收窄至上限後，長句改以自動換行顯示（相關格已設 wrap_text）。
    for col_idx in range(1, ws.max_column + 1):
        dim = ws.column_dimensions[get_column_letter(col_idx)]
        cap = _SSR_SUMMARY_LABEL_COL_MAX if col_idx == 1 else _SSR_SUMMARY_COL_MAX
        if dim.width and dim.width > cap:
            dim.width = cap
    _ssr_fit_wrapped_rows(ws)


def _ssr_fit_wrapped_rows(ws, line_pt=15.0):
    """為設有自動換行的長文字格調整列高，確保收窄欄寬後文字仍完整顯示。

    以顯示寬度估算所需行數（全形字元約佔 2 個欄寬單位），列高取該列最大行數；
    只會加大、不會縮小原有列高。合併格（如區塊標題）不計——其文字在合併範圍內延伸。
    """
    merged = set()
    for m in ws.merged_cells.ranges:
        for r in range(m.min_row, m.max_row + 1):
            for c in range(m.min_col, m.max_col + 1):
                if m.max_col > m.min_col or (r, c) != (m.min_row, m.min_col):
                    merged.add((r, c))
    for row in ws.iter_rows():
        lines = 1
        for cell in row:
            if (cell.row, cell.column) in merged or not isinstance(cell.value, str):
                continue
            if not cell.alignment or not cell.alignment.wrap_text:
                continue
            width = ws.column_dimensions[get_column_letter(cell.column)].width or 8.43
            text_w = sum(2.0 if unicodedata.east_asian_width(ch) in ("W", "F") else 1.0 for ch in cell.value)
            lines = max(lines, int(-(-text_w // max(width - 2.0, 1.0))))
        if lines > 1:
            dim = ws.row_dimensions[row[0].row]
            need = lines * line_pt + 3
            if (dim.height or 0) < need:
                dim.height = need



def generate_excel_combined(
    summary_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    categories: list[str],
    progress_callback=None,
    dimension: str = "nature",
    om_groups=None,
    stats: dict | None = None,
    include_raw: bool = True,
) -> BytesIO:
    """Safety Stock Review 輸出。dimension／om_groups 決定 SUMMARY 區塊 ③
    「Safety Qty Change 分類 × 分類維度交叉表」的分組（只影響顯示，不影響任何建議計算）；
    om_groups 未提供時由 raw 即時推導。stats 只供首頁「說明」導讀表照錄本次執行參數
    （缺失時顯示預設值），不參與任何計算。輸出 7 工作表：說明（首頁導讀）＋
    RAW／Shop Stock Status／By Shop／店舖問題 SKU／SUMMARY／Safety Stock 數據分析摘要。"""
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    _dim = _hybrid_dimension(dimension)
    if _dim in ("om", "nature_om") and not om_groups and raw_df is not None and not raw_df.empty:
        om_groups = _hybrid_om_groups_from_values(
            raw_df["OM"] if "OM" in raw_df.columns else _safe_default_series("", raw_df.index)
        )
    cross_tab = _build_combined_dimension_cross_tab(raw_df, dimension=_dim, om_groups=om_groups)

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    desc_font = Font(size=9, color="6B7280", italic=True)

    # ── Sheet 0: 說明（邏輯、本次執行參數與導讀；置於工作簿首位）──
    ws_guide = wb.active
    ws_guide.title = "說明"
    _write_combined_guide_sheet(ws_guide, stats, _dim)

    # ── Sheet 1: RAW ──
    ws_raw = wb.create_sheet(title="RAW")

    (_write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.15, columns=COMBINED_RAW_COLS) if include_raw else _write_raw_placeholder(ws_raw))

    if progress_callback:
        progress_callback(0.2, "建立摘要工作表...")

    # ── Sheet 2: Shop Stock Status ──
    ws_summary = wb.create_sheet(title="Shop Stock Status")

    headers = ["Buyer", "Brand", "SKU", "SKU Desc"] + categories + ["總計", "Safety Qty Change", "增加Safety Qty的店舖數量", "減少Safety Qty的店舖數量", "Safety Qty不變的店舖數量"]
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, h in enumerate(headers, 1):
        cell = ws_summary.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    total_summary = len(summary_df)
    # P0：itertuples＋共用對齊取代 iterrows＋逐格新建 Alignment；值與格式語意不變
    center_align = Alignment(horizontal="center", vertical="center")
    for row_idx, _vals in enumerate(_iter_header_rows(summary_df, headers)):
        excel_row = 2 + row_idx
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
            progress_callback(0.4 + 0.2 * ((row_idx + 1) / max(total_summary, 1)), f"寫入摘要 ({row_idx+1}/{total_summary})...")

    _auto_width(ws_summary)

    if progress_callback:
        progress_callback(0.65, "建立 By Shop 分析工作表...")

    # ── Sheet 3: By Shop ──
    by_shop_df = _build_combined_by_shop_summary(raw_df)

    ws_by_shop = wb.create_sheet(title="By Shop")
    by_shop_headers = [
        "Site Code",
        "Shop Name",
        "OM",
        "Overstock_SKU數量",
        "OOS_SKU數量",
        "需增加Safety_SKU數量",
        "需減少Safety_SKU數量",
        "安全庫存調整總量",
        "Combined_SKU總數",
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
            progress_callback(0.7 + 0.15 * ((row_idx + 1) / max(len(by_shop_df), 1)), f"寫入 By Shop ({row_idx+1}/{len(by_shop_df)})...")

    _auto_width(ws_by_shop)

    if progress_callback:
        progress_callback(0.85, "建立店舖問題 SKU 工作表...")

    # ── Sheet 4: 店舖問題 SKU（每間店舖需行動列；空表寫「無符合資料」）──
    # 同一批需行動列亦是 SUMMARY ⑤–⑦ 的統計來源，只建立一次
    ws_shop_issue = wb.create_sheet(title="店舖問題 SKU")
    shop_issues = _build_combined_shop_issues(raw_df)
    _write_shop_issue_sheet(
        ws_shop_issue, shop_issues,
        title="店舖問題 SKU — 每間店舖需行動列（問題＋建議方法）",
        progress_callback=progress_callback, progress_start=0.85, progress_end=0.86,
    )

    if progress_callback:
        progress_callback(0.86, "建立 SUMMARY 工作表...")

    # ── Sheet 5: SUMMARY（Safety Qty Change 分類 / ND 建議分類 / 影響範圍 / 問題類型樞紐）──
    change_summary, nd_summary = _build_combined_classification_summary(raw_df)
    target_summary = _build_combined_target_summary(raw_df)
    ws_class_summary = wb.create_sheet(title="SUMMARY")
    _write_summary_sheet(
        ws_class_summary, change_summary, nd_summary, thin_border,
        cross_tab=cross_tab, target_summary=target_summary,
        impact_summary=_build_combined_impact_summary(raw_df, issues=shop_issues),
        buyer_brand_pivot=_build_combined_issue_pivot(shop_issues, "buyer_brand"),
        shop_pivot=_build_combined_issue_pivot(shop_issues, "shop"),
    )

    if progress_callback:
        progress_callback(0.88, "建立分析摘要工作表...")

    # ── Sheet 6: Safety Stock 數據分析摘要 ──
    ws_dashboard = wb.create_sheet(title="Safety Stock 數據分析摘要")

    _total_row = summary_df[summary_df["Buyer"] == "總計"]
    if _total_row.empty:
        raise ValueError("摘要資料中未找到總計列，無法生成分析儀表板")
    total_combined = int(_total_row["總計"].values[0])
    total_buyers = len(summary_df[summary_df["Buyer"] != "總計"]["Buyer"].dropna().unique())
    total_brands = len(summary_df[summary_df["Buyer"] != "總計"]["Brand"].dropna().unique())
    total_skus = raw_df["SKU"].dropna().nunique()

    cat_overstock = int(_total_row[categories[0]].values[0]) if categories[0] in _total_row.columns else 0
    cat_oos = int(_total_row[categories[1]].values[0]) if categories[1] in _total_row.columns else 0
    cat_balanced = int(_total_row[categories[2]].values[0]) if categories[2] in _total_row.columns else 0

    safety_change_total = int(pd.to_numeric(raw_df.get("Safety Qty Change", _safe_default_series(0, raw_df.index)), errors="coerce").fillna(0).sum())

    row1_metrics = [
        ("總 Safety Stock 審查行數", total_combined, "所有 Safety Stock 審查之總計行數", "EF4444"),
        ("影響採購人員數 (Buyers)", total_buyers, "位採購人員名下有需審查品項", "3B82F6"),
        ("影響品牌總數 (Brands)", total_brands, "個品牌包含需審查商品", "8B5CF6"),
        ("影響SKU總數", total_skus, "個包含需審查商品(SKUs)", "10B981"),
    ]

    row2_metrics = [
        ("Overstock 分類行數", cat_overstock, "庫存過剩分類行數", "F59E0B"),
        ("OOS 分類行數", cat_oos, "缺貨分類行數", "EC4899"),
        ("Safety 調整行數", cat_balanced, "Balanced (Safety Adjusted) 行數", "06B6D4"),
        ("Safety 調整總量", safety_change_total, "Safety Qty Change 總和", "8B5CF6"),
    ]

    _write_metric_cards(ws_dashboard, "Safety Stock 數據分析摘要 (Safety Stock Analysis Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback)

    if progress_callback:
        progress_callback(0.97, "正在儲存 Excel 檔案...")

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf









def _write_combined_guide_sheet(ws, stats, dimension):
    """Safety Stock Review「說明」工作表（工作簿首位）：列出本次執行參數、
    講解 Review 邏輯，並逐頁導讀各工作表應留意的部分。"""
    st_ = stats or {}
    dim = _hybrid_dimension(dimension)
    _os = st_.get("rates_overstock") or {}
    _oos = st_.get("rates_oos") or {}
    os_t = int(round(float(_os.get("tourist", 0.30)) * 100))
    os_m = int(round(float(_os.get("mix", 0.40)) * 100))
    os_n = int(round(float(_os.get("default", 0.50)) * 100))
    oos_t = int(round(float(_oos.get("tourist", 0.50)) * 100))
    oos_m = int(round(float(_oos.get("mix", 0.50)) * 100))
    oos_n = int(round(float(_oos.get("default", 0.50)) * 100))
    mode_txt = ("可加可減（允許減少並套用退場建議）" if st_.get("safety_mode") == "allow_both"
                else "只可增加（Safety 只增不減，ND 列不作上調、P90D 0 Sales 退場建議不套用）")
    _basis = _normalize_ssr_safety_basis(st_.get("safety_basis", "rate"))
    _sd = st_.get("safety_days") or dict(_HYBRID_SAFETY_DAYS_DEFAULT)
    _sd_label = _hybrid_safety_days_label({
        "Tourist": int(_sd.get("Tourist", 10)),
        "Mix": int(_sd.get("Mix", 9)),
        "Non Tourist": int(_sd.get("Non Tourist", 7)),
    })
    basis_txt = ("周轉天數 (按 Nature) — 建議值取 Safety 周轉所需量（有銷量或新品時最低 3 件，舊品零銷量維持 0）"
                 if _basis == _SSR_SAFETY_BASIS_TURNDAY
                 else "月銷倍率 (X/Y) — 建議值 = max(Y, Z if Z ≤ X else X)")

    _tgt_raw = st_.get("target_safety") or {}
    try:
        _tgt = {str(_k): int(_v) for _k, _v in dict(_tgt_raw).items()}
    except (TypeError, ValueError):
        _tgt = {}
    if _tgt:
        _tgt_list = "、".join(f"{k}={v:,}" for k, v in list(_tgt.items())[:8])
        _tgt_txt = (
            f"（來源：{st_.get('target_safety_source') or '表格填寫'}）"
            f"{len(_tgt)} 隻 SKU 指定全線合計（合計 {sum(_tgt.values()):,} 件）：{_tgt_list}"
            + ("…" if len(_tgt) > 8 else "")
            + f"；分配基準＝{_SSR_TARGET_BASIS_LABEL}，以最大餘數法取整，分配總和等於指定值；"
            "RP Type＝ND 的店舖不參與分配，未指定的 SKU 按上方基準計算。"
        )
    else:
        _tgt_txt = "未啟用（全部 SKU 按上方基準計算）"

    params = [
        ("分類維度",
         f"{_HYBRID_DIMENSION_LABELS[dim]}（只影響 SUMMARY 區塊 ③ 的分組，不影響 Safety 建議計算）"),
        ("Safety Qty 計算模式", mode_txt),
        ("Safety 計算基準", f"{basis_txt}；基準天數：{_sd_label}"),
        ("指定總量分配 (Target Safety)", _tgt_txt),
        ("Overstock Rate",
         f"Tourist {os_t}%／Mix {os_m}%／Non Tourist {os_n}%（庫存上限 X = floor(月均銷售 × Rate)，有銷量或新品時另設最低 3 件；上限不超過現有 Safety）"),
        ("OOS Rate",
         f"Tourist {oos_t}%／Mix {oos_m}%／Non Tourist {oos_n}%（缺貨底線 Y = ceil(月銷預估 × Rate)；有銷量或新品時最低 3 件，舊品零銷量維持 0）"),
        ("Safety 基準天數",
         f"{_sd_label}（Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 有效統計日數 × 基準天數)；"
         "有效統計日數 = min(90, max(上市至參考日天數（含首尾）, 3))，上市滿 90 日者為 90；"
         "有銷量或新品時最低 3 件，舊品零銷量維持 0；與調貨段無關）"),
        ("MTD 日期",
         f'MTD {st_.get("mtd_days", "—")} 日（當月共 {st_.get("month_total_days", "—")} 日；'
         "月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均，月初以 P90D 為主）"),
        ("參考日期（New SKU 保護）",
         f'{st_.get("ref_date", st_.get("mtd_date", "—"))}（Launch Date 起 90 天內視為新品，預設 MTD 日期）'),
        ("銷量分類門檻",
         f'缺貨線 {st_.get("sales_oos_days", _HYBRID_SALES_OOS_DAYS_DEFAULT)} 日／'
         f'積壓線 {st_.get("sales_overstock_days", _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT)} 日／'
         f'保底 {st_.get("sales_overstock_min_units", _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT)} 件'
         "（銷量帶與銷量存貨分類恆按此計算）"),
    ]

    logic = [
        "Overstock Cap (X) = floor(月均銷售 × Overstock Rate)，有銷量或新品時另設最低 3 件（舊品零銷量 X = 0），上限不超過現有 Safety 設定"
        "（Safety Stock (現時)，即來源檔案的 Adj. Safety Stock）。",
        "OOS Floor (Y) = ceil(月銷預估 × OOS Rate)；有銷量或 90 天內上市的新品時最低 3 件，"
        "舊品零銷量維持 0（不為滯銷品項設安全庫存）；月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均"
        "（月初以 P90D 為主，月尾收斂至 MTD；P90D 缺值時退回純 MTD 線性外推）。",
        "Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 有效統計日數 × Safety 基準天數)；"
        "有效統計日數 = min(90, max(上市至參考日天數（含首尾）, 3))（上市未滿 90 日的新品只按實際可售日數"
        "換算日均，避免以 90 日作分母低估銷量；上市日缺值或無法解析時按 90 日）；"
        "有銷量或 90 天內上市的新品時最低 3 件，舊品零銷量維持 0"
        "（與 Safety Stock x NetStock Review 同口徑；負數 P90D 按 0 計算）。",
        "建議 Safety Stock（月銷倍率基準）= max(Y, Z if Z ≤ X else X)；"
        "建議 Safety Stock（周轉天數基準）= Safety 周轉所需量（Z = 現有 Safety 設定）；"
        "Safety Qty Change = 建議值 − Z（正=增加、負=減少、零=不變）。",
        "指定總量分配（Target Safety，選用）：只覆寫被指定的 SKU。各店權重 = 0.5 ×（該店 Sold Qty (P90D) ÷ 同 SKU 合計）"
        "+ 0.5 ×（該店現有 Safety ÷ 同 SKU 合計），再按指定全線合計以最大餘數法分配整數件，"
        "分配總和恆等於指定值；某一半合計為 0 時該半不參與，兩半皆為 0 時退回均分。"
        "RP Type＝ND 的店舖不參與（Safety 必須為 0）；分配值直接取代該 SKU 的建議值，"
        "不受 Overstock Cap／OOS Floor 限制，亦不受「只可增加」限制（否則指定值永遠達不到），"
        "落差逐列寫入「Target 分配提示」。",
        "Recommend Safety Stock Turnday = 建議值 ÷ (P90D 月均 ÷ 30)（月均 0 留空；"
        "與來源「Turnday (P90D Qty)」同一準則）。",
        "「只可增加」模式全表只增不減（指定總量分配的列除外），且 ND 列（RP Type＝ND）一律不作上調"
        "（ND＝不執行自動補貨，Safety 不能 > 0 件，建議值維持現有設定）；"
        "ND／P90D 0 Sales 的歸零與 ND 退場建議不套用；"
        "「可加可減」模式套用退場建議：ND 列 → 0；P90D 0 Sales 舊品 → ND"
        "（並依 Adj. MCH1／Adj. SST 加註 ND 類別）。",
        "New SKU 保護：Launch Date 距參考日期 90 天內的列排除「P90D 0 Sales → Safety=0/ND」override，"
        "改走正常基準計算並標記 New SKU (3 Month)。",
        "銷量帶（常開顯示欄）：同 SKU 內按 Sold Qty (P90D) 的店舖平均排名分帶——排名 > 75%＝高銷、"
        "≤ 25%＝低銷、其餘＝中銷、P90D ≤ 0＝零銷；並列同值同帶；有銷售店舖少於 4 間時不設高銷／低銷。",
        "銷量存貨分類（常開顯示欄）：高銷且店舖周轉日數 < 缺貨線 → 高銷缺貨；低銷且周轉 ≥ 積壓線且"
        "庫存 ≥ 保底件數 → 低銷積壓；零銷且庫存 ≥ 保底件數 → 零銷積壓；其餘 → 正常（門檻見上方 ①）；"
        "兩欄與 Safety 建議計算互相獨立。",
        "建議值受最低 3 件保底影響時（最終建議值 = 3 件且未套保底前 < 3），"
        "店舖問題 SKU 的建議方法加註保底原因；保底只適用於有銷量或 90 天內上市的新品，"
        "舊品零銷量不會因保底而被上調至 3 件；建議方法亦附「建議周轉約 X 日」。",
    ]

    sheet_rows = [
        ("說明（本頁）", "Review 邏輯、本次參數與導讀", "執行任何行動前，先核對上方 ① 的參數是否符合預期"),
        ("RAW", "全部店舖×SKU 逐列明細與判定欄",
         "留意「Shop Stock Status」「Safety Qty Change」「Recommend Safety Stock Turnday」「銷量帶」「銷量存貨分類」；"
         "有指定總量分配時另見「Target Safety (全線合計)」「Safety 分配佔比」「Target 分配提示」；"
         "「只可增加」模式不會出現「減少」（指定總量分配的列除外），且 RP Type＝ND 列不作上調"),
        ("Shop Stock Status", "採購人員×品牌之庫存異常交叉統計", "各 Buyer／品牌的分類分佈概覽"),
        ("By Shop", "各店舖層級彙總", "留意「安全庫存調整總量」與需增／需減 SKU 數較多的店舖"),
        ("店舖問題 SKU", "每間店舖需行動列（問題＋建議方法）",
         "店務同事可先從這一頁著手；含 Safety 需增加／減少、ND 退場建議與調入缺口跟進；建議方法附建議周轉日數"),
        ("SUMMARY", "① Safety Qty Change 分類 ／ ② ND 建議分類 ／ ③ 分類維度交叉表 ／ ④ 指定總量分配對帳 ／ ⑤ 影響範圍統計（Brand／Buyer／店舖）／ ⑥ Buyer＞Brand 樞紐分析 ／ ⑦ 店舖樞紐分析",
         "④ 只在有指定 SKU 時出現；核對「已分配」是否等於「Target 全線合計」（差額應為 0）；"
         "⑤ 一眼看出多少個 Brand／Buyer／店舖有 Safety 設定不足、設定過高或 ND 退場；"
         "⑥⑦ 以「店舖問題 SKU」同一批列按問題類型列出，可按 Buyer 收合 Brand、或按店舖篩選"),
        ("Safety Stock 數據分析摘要", "指標卡總覽", "留意 Overstock／OOS／調整行數與調整總量"),
    ]

    tips = [
        "「Safety Qty Change」正數＝建議增加、負數＝建議減少、零＝不變；「只可增加」模式下不會出現負數"
        "（唯一例外：指定總量分配的列以公司指定值為準，可能低於現有 Safety）。",
        "RP Type＝ND 的列不會出現增加建議（ND 不執行自動補貨，Safety 不能 > 0 件），"
        "「只可增加」模式下建議值維持現有設定；ND 歸零只出現在「可加可減」模式。",
        "「只可增加」模式下 Overstock Rate 不產生壓縮效果；要壓縮庫存請選「可加可減」。",
        "指定總量分配為「公司指令」而非公式結果，故不套用 Overstock Cap／OOS Floor 上限；"
        "分配值超出 X 或低於 Y 的店舖會列在「Target 分配提示」，SUMMARY ④ 亦有逐 SKU 店數統計。"
        "此類 Safety 屬臨時性質（如配合大堆頭陳列），陳列期完結後需自行調回，"
        "否則下一次 Review 會計出「需減少」。",
        "建議值最低 3 件為保障自動補貨的保底，只適用於有銷量或 90 天內上市的新品；"
        "舊品零銷量維持 0（不為滯銷品項設安全庫存）；受影響的店舖問題列會加註原因。",
        "分類維度只改變 SUMMARY 區塊 ③ 的分組顯示，不影響任何 Safety 建議計算。",
    ]

    _write_guide_sheet(
        ws, title="Safety Stock Review 說明與導讀",
        subtitle="本頁講解此報告的 Review 邏輯與本次執行參數，並逐頁說明應留意的部分；各項判定以 RAW 逐列欄位為準。最下方的 ⑤ 為 Safety 計算邏輯圖。",
        params=params, logic=logic, sheet_rows=sheet_rows, tips=tips,
        image="assets/safety_logic_guide.png",
        image_section="⑤ Safety 計算邏輯圖（與本報告建議值同一準則：Z＝Safety Stock (現時)、X0＝Overstock Cap、Y＝OOS Floor）",
        image_width_px=900,
    )

