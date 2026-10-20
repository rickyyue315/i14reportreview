# -*- coding: utf-8 -*-
"""調貨與分類框架（dimension／sales／OM／safety-days）：由 transformer.py 拆分而來的共用核心。
transformer 以墊片 re-export，舊有 import 路徑不變。
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
    _normalize_ref_date,
    _parse_launch_dates,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _standardize_safety_column,
)
__all__ = [
    "_HYBRID_COMBINED_GROUPS",
    "_HYBRID_DIAG_NATURES",
    "_HYBRID_DIAG_NORMAL_DONOR",
    "_HYBRID_DIAG_NORMAL_NEED",
    "_HYBRID_DIAG_ORDER",
    "_HYBRID_DIAG_OVER",
    "_HYBRID_DIAG_UNDER",
    "_HYBRID_DIMENSIONS",
    "_HYBRID_DIMENSION_LABELS",
    "_HYBRID_EVENT_DONATED",
    "_HYBRID_EVENT_DONOR",
    "_HYBRID_EVENT_NEED_IN",
    "_HYBRID_EVENT_ORDER",
    "_HYBRID_OM_BLANK_LABEL",
    "_HYBRID_REQUIRED",
    "_HYBRID_RETURN_D001_COLUMN",
    "_HYBRID_RETURN_D001_COLUMNS",
    "_HYBRID_RETURN_D001_GROSS_COLUMN",
    "_HYBRID_RETURN_D001_GROSS_INTERNAL",
    "_HYBRID_RETURN_D001_SHEET",
    "_HYBRID_RF_TURND_SUFFIX",
    "_HYBRID_FULFILL_COLUMN",
    "_HYBRID_FULFILL_FULL",
    "_HYBRID_FULFILL_PARTIAL",
    "_HYBRID_FULFILL_NONE",
    "_HYBRID_RP_GROUPS",
    "_HYBRID_SAFETY_DAYS_COLUMN",
    "_HYBRID_SAFETY_DAYS_DEFAULT",
    "_HYBRID_SAFETY_DAYS_MAX",
    "_HYBRID_SAFETY_DAYS_MIN",
    "_HYBRID_SAFETY_LINE_COLUMN",
    "_HYBRID_SAFETY_MIN_KEEP",
    "_HYBRID_FULL_SAFETY_LINE_COLUMN",
    "_HYBRID_FULL_SAFETY_REVIEW_COLUMN",
    "_HYBRID_FULL_SAFETY_SUGGEST_COLUMN",
    "_HYBRID_FULL_SAFETY_CHG_COLUMN",
    "_HYBRID_FULL_SAFETY_TURNDAY_COLUMN",
    "_HYBRID_FULL_SAFETY_NORMAL",
    "_HYBRID_FULL_SAFETY_SHEET",
    "_HYBRID_FULL_SAFETY_DETAIL_COLUMNS",
    "_HYBRID_SALES_BANDS",
    "_HYBRID_SALES_BAND_COLUMN",
    "_HYBRID_SALES_CLASS_COLUMN",
    "_HYBRID_SALES_CLASS_ORDER",
    "_HYBRID_SALES_MIN_SOLD_ROWS",
    "_HYBRID_SALES_NATURE_GROUPS",
    "_HYBRID_SALES_OOS_DAYS_DEFAULT",
    "_HYBRID_SALES_OOS_DAYS_MAX",
    "_HYBRID_SALES_OOS_DAYS_MIN",
    "_HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT",
    "_HYBRID_SALES_OVERSTOCK_DAYS_MAX",
    "_HYBRID_SALES_OVERSTOCK_DAYS_MIN",
    "_HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT",
    "_HYBRID_SALES_OVERSTOCK_MIN_UNITS_MAX",
    "_HYBRID_SALES_OVERSTOCK_MIN_UNITS_MIN",
    "_HYBRID_SALES_P_HIGH",
    "_HYBRID_SALES_P_LOW",
    "_HYBRID_SHORTAGE_THRESHOLDS",
    "_HYBRID_SKEW_THRESHOLDS",
    "_HYBRID_SUGGEST_TURNDAY_COLUMN",
    "_canonicalize_hybrid_columns",
    "_hybrid_dimension",
    "_hybrid_dimension_groups",
    "_hybrid_numeric_column",
    "_hybrid_om_groups_from_values",
    "_hybrid_safety_days",
    "_hybrid_safety_days_label",
    "_hybrid_sales_bands",
    "_hybrid_sales_classification",
    "_hybrid_sales_params",
    "_hybrid_shortage_hint_text",
    "_hybrid_shortage_thresholds",
    "_hybrid_skew_hint_text",
    "_hybrid_skew_thresholds",
    "_hybrid_zero_sales_groups",
    "_normalize_hybrid_nature",
    "_normalize_hybrid_om",
    "_normalize_hybrid_rp_type",
    "_transfer_input_consistency_warnings",
    "validate_hybrid_input",
    "_attach_sales_classification_columns",
    "_HYBRID_DIAG_COLUMNS",
    "_HYBRID_RAW_DISPLAY_GROUPS",
    "_hybrid_p90d_zero_sales_mask",
    "_HYBRID_P90D_ZERO_SALES_COLUMNS",
    "_hybrid_group_key_series",
]


_HYBRID_REQUIRED = [
    "Buyer", "Brand", "SKU",
    "Shop Nature", "Site Code", "Shop Stock Qty (CM)",
    "Sold Qty (P90D)", "Adj. Safety Stock",
    "Shop Monthly Sold Qty (P90D)",
]



# ── 調貨事件 Review（Safety 基準天數按 Nature；調貨段仍為 14 日）──
_HYBRID_EVENT_NEED_IN = "需要調入"



_HYBRID_EVENT_DONATED = "實際轉出"



_HYBRID_EVENT_DONOR = "可轉出未使用"



_HYBRID_DIAG_UNDER = "Safety 設定不足"



_HYBRID_DIAG_OVER = "Safety 設定過高"



_HYBRID_DIAG_NORMAL_NEED = "設定正常—缺貨非設定因素"



_HYBRID_DIAG_NORMAL_DONOR = "設定正常—過剩非設定因素"



_HYBRID_DIAG_ORDER = [
    _HYBRID_DIAG_UNDER, _HYBRID_DIAG_OVER,
    _HYBRID_DIAG_NORMAL_NEED, _HYBRID_DIAG_NORMAL_DONOR,
]



_HYBRID_EVENT_ORDER = {_HYBRID_EVENT_NEED_IN: 0, _HYBRID_EVENT_DONATED: 1, _HYBRID_EVENT_DONOR: 2}



_HYBRID_DIAG_NATURES = ["Tourist", "Mix", "Non Tourist"]



# ── 分類維度（需求方過濾與彙總分組；TurnDay 池與 Safety 基準天數恆按 Shop Nature）──
# nature：以 Shop Nature 分類（預設；需求方 Tourist-RF）
# rp：以 RP Type 分類（需求方任何 Nature 的 RF 列，供應方維持 Mix／Non Tourist，允許同 Nature 互調）
# both：以 Shop Nature＋RP Type 分類（6 組；計算同 nature 模式，僅分組顯示不同）
# om：以 OM 分類（動態 OM 名單；計算同 nature 模式，僅分組顯示不同）
# nature_om：以 Shop Nature＋OM 分類（3×N 組；計算同 nature 模式，僅分組顯示不同）
# sales：以銷量分類（高銷／中銷／低銷／零銷 四帶；計算同 nature 模式，僅分組顯示不同）
# sales_nature：以銷量＋Shop Nature 分類（3 Nature×4 帶＝12 組；計算同 nature 模式，僅分組顯示不同）
_HYBRID_DIMENSIONS = ("nature", "rp", "both", "om", "nature_om", "sales", "sales_nature")



_HYBRID_DIMENSION_LABELS = {
    "nature": "Shop Nature",
    "rp": "RP Type",
    "both": "Shop Nature＋RP Type",
    "om": "OM",
    "nature_om": "Shop Nature＋OM",
    "sales": "銷量",
    "sales_nature": "銷量＋Shop Nature",
}



# 存貨偏重提示門檻（庫存佔比下限，庫存佔比 − 銷量佔比差距下限）：
# nature／rp／both 沿用現行契約；om（約 7 組）與 nature_om（3×N 組）因組數多、
# 平均佔比低，沿用原門檻幾乎不會觸發，故分別下調；sales（4 帶）介於兩者之間；
# sales_nature（12 組）組數最多，與 nature_om 同門檻。
_HYBRID_SKEW_THRESHOLDS = {
    "nature": (0.40, 0.20),
    "rp": (0.40, 0.20),
    "both": (0.40, 0.20),
    "om": (0.25, 0.12),
    "nature_om": (0.12, 0.08),
    "sales": (0.30, 0.15),
    "sales_nature": (0.12, 0.08),
}



# 存貨偏低標示門檻（銷量佔比下限，銷量佔比 − 庫存佔比差距下限）：
# 與存貨偏重對稱，但採用較寬鬆門檻——缺貨風險較積壓更需及早發現，
# 故各維度佔比與差距要求均低於偏重側；只作儲存格標示，不新增文字提示欄。
_HYBRID_SHORTAGE_THRESHOLDS = {
    "nature": (0.30, 0.10),
    "rp": (0.30, 0.10),
    "both": (0.30, 0.10),
    "om": (0.20, 0.08),
    "nature_om": (0.10, 0.05),
    "sales": (0.20, 0.10),
    "sales_nature": (0.10, 0.05),
}



_HYBRID_RP_GROUPS = ("RF", "ND")



_HYBRID_COMBINED_GROUPS = tuple(
    f"{n}-{r}" for n in ("Tourist", "Mix", "Non Tourist") for r in _HYBRID_RP_GROUPS
)



# ── 銷量帶／銷量存貨分類（常開顯示欄；「銷量再平衡」調貨通道另受開關控制）──
# 銷量帶：同 SKU 內按 Sold Qty (P90D) 排名分帶——P90D≤0＝零銷；
# 有銷列中排名超過 75%＝高銷、排名在最低 25% 內＝低銷、其餘＝中銷；
# 同 SKU 有銷列數 < _HYBRID_SALES_MIN_SOLD_ROWS 時不產生高銷／低銷（全歸中銷，
# 避免小樣本假區分）。並列以平均排名處理（同值同帶）。
_HYBRID_SALES_BANDS = ("高銷", "中銷", "低銷", "零銷")



_HYBRID_SALES_BAND_COLUMN = "銷量帶"



_HYBRID_SALES_CLASS_COLUMN = "銷量存貨分類"



_HYBRID_SALES_CLASS_ORDER = ("高銷缺貨", "低銷積壓", "零銷積壓", "正常")



_HYBRID_SALES_P_HIGH = 0.75



_HYBRID_SALES_P_LOW = 0.25



_HYBRID_SALES_MIN_SOLD_ROWS = 4



_HYBRID_SALES_OOS_DAYS_DEFAULT = 7



_HYBRID_SALES_OOS_DAYS_MIN = 1



_HYBRID_SALES_OOS_DAYS_MAX = 14



_HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT = 30



_HYBRID_SALES_OVERSTOCK_DAYS_MIN = 15



_HYBRID_SALES_OVERSTOCK_DAYS_MAX = 90



_HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT = 2



_HYBRID_SALES_OVERSTOCK_MIN_UNITS_MIN = 1



_HYBRID_SALES_OVERSTOCK_MIN_UNITS_MAX = 100



# 銷量＋Shop Nature 合併維度的分組（Nature 外層固定順序、銷量帶內層固定順序，如 Tourist-高銷；
# 計算同 nature 模式，僅 Summary／交叉表分組顯示不同）
_HYBRID_SALES_NATURE_GROUPS = tuple(
    f"{n}-{b}" for n in _HYBRID_DIAG_NATURES for b in _HYBRID_SALES_BANDS
)



def _hybrid_sales_params(
    sales_rebalance,
    sales_oos_days=None,
    sales_overstock_days=None,
    sales_overstock_min_units=None,
) -> tuple[bool, dict, list[str]]:
    """銷量分類參數正規化：開關轉 bool；三個門檻以預設為底，
    非整數拒絕、範圍外箝制並列 warning（風格同 _hybrid_safety_days）。
    回傳 (sales_rebalance, {"oos_days", "overstock_days", "overstock_min_units"}, warnings)。"""
    warnings: list[str] = []
    params = {
        "oos_days": _HYBRID_SALES_OOS_DAYS_DEFAULT,
        "overstock_days": _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT,
        "overstock_min_units": _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT,
    }
    spec = {
        "oos_days": (sales_oos_days, _HYBRID_SALES_OOS_DAYS_MIN, _HYBRID_SALES_OOS_DAYS_MAX, "缺貨線（日）"),
        "overstock_days": (sales_overstock_days, _HYBRID_SALES_OVERSTOCK_DAYS_MIN, _HYBRID_SALES_OVERSTOCK_DAYS_MAX, "積壓線（日）"),
        "overstock_min_units": (sales_overstock_min_units, _HYBRID_SALES_OVERSTOCK_MIN_UNITS_MIN, _HYBRID_SALES_OVERSTOCK_MIN_UNITS_MAX, "積壓保底件數（件）"),
    }
    for key, (val, lo, hi, label) in spec.items():
        if val is None:
            continue
        if isinstance(val, bool):
            raise ValueError(f"銷量門檻「{label}」必須為整數，不得為布林值。")
        try:
            numeric = float(val)
        except (TypeError, ValueError):
            continue
        if not np.isfinite(numeric) or not numeric.is_integer():
            raise ValueError(f"銷量門檻「{label}」必須為整數。")
        n = int(numeric)
        if n < lo or n > hi:
            clamped = min(max(n, lo), hi)
            warnings.append(f"銷量門檻「{label}」輸入 {n} 超出建議範圍，已箝制為 {clamped}。")
            n = clamped
        params[key] = n
    return bool(sales_rebalance), params, warnings



def _hybrid_sales_bands(sold_p90d: pd.Series, sku_key: pd.Series) -> pd.Series:
    """同 SKU 內銷量帶（見 _HYBRID_SALES_BANDS 註解）：Sold Qty (P90D)≤0＝零銷；
    有銷列按同 SKU 平均排名分帶（>75%＝高銷、≤25%＝低銷、其餘＝中銷）；
    有銷列數 < _HYBRID_SALES_MIN_SOLD_ROWS 的 SKU 不產生高銷／低銷。"""
    result = pd.Series(_HYBRID_SALES_BANDS[3], index=sold_p90d.index, dtype=object)
    pos_mask = sold_p90d > 0
    if not bool(pos_mask.any()):
        return result
    sub_sold = sold_p90d[pos_mask]
    sub_sku = sku_key[pos_mask]
    ranks = sub_sold.groupby(sub_sku, sort=False).rank(method="average")
    counts = sub_sold.groupby(sub_sku, sort=False).transform("size")
    frac = ranks / counts
    band = pd.Series(_HYBRID_SALES_BANDS[1], index=sub_sold.index, dtype=object)
    band[frac > _HYBRID_SALES_P_HIGH] = _HYBRID_SALES_BANDS[0]
    band[frac <= _HYBRID_SALES_P_LOW] = _HYBRID_SALES_BANDS[2]
    band[counts < _HYBRID_SALES_MIN_SOLD_ROWS] = _HYBRID_SALES_BANDS[1]
    result[sub_sold.index] = band
    return result



def _hybrid_sales_classification(
    band: pd.Series,
    shop_turn_day: np.ndarray,
    cm_stock: np.ndarray,
    params: dict,
) -> pd.Series:
    """銷量存貨分類（逐列，全列覆蓋）：
    高銷 且 店舖 TurnDay < 缺貨線 → 高銷缺貨；
    低銷 且 店舖 TurnDay ≥ 積壓線 且 庫存 ≥ 保底件數 → 低銷積壓；
    零銷 且 庫存 ≥ 保底件數 → 零銷積壓（庫存不足保底的零銷有存列歸正常，
    仍由既有「Past 90 Days 0 Sales 轉 ND」提示涵蓋）；其餘 → 正常。
    邊界：缺貨用「低於」（等於缺貨線為正常），積壓用「不少於」。"""
    band_txt = band.fillna("").astype(str)
    td = np.asarray(shop_turn_day, dtype=float)
    stock = np.asarray(cm_stock, dtype=float)
    cls = pd.Series(_HYBRID_SALES_CLASS_ORDER[3], index=band.index, dtype=object)
    is_high = (band_txt == _HYBRID_SALES_BANDS[0]).to_numpy()
    is_low = (band_txt == _HYBRID_SALES_BANDS[2]).to_numpy()
    is_zero = (band_txt == _HYBRID_SALES_BANDS[3]).to_numpy()
    oos = is_high & (td < float(params["oos_days"]))
    over = is_low & (td >= float(params["overstock_days"])) & (stock >= float(params["overstock_min_units"]))
    zero_over = is_zero & (stock >= float(params["overstock_min_units"]))
    cls[over] = _HYBRID_SALES_CLASS_ORDER[1]
    cls[zero_over] = _HYBRID_SALES_CLASS_ORDER[2]
    cls[oos] = _HYBRID_SALES_CLASS_ORDER[0]
    return cls



def _hybrid_dimension(dimension) -> str:
    """分類維度正規化：非法值退回預設 nature，不拋錯（沿用 _hybrid_safety_days 風格）。"""
    d = str(dimension or "nature").strip().lower()
    return d if d in _HYBRID_DIMENSIONS else "nature"



def _hybrid_dimension_groups(dimension, om_groups=None) -> tuple[str, ...]:
    """分類維度對應的彙總分組名稱（Summary／交叉表欄位按此切換）。

    sales 維度回傳固定四帶（高銷／中銷／低銷／零銷）；
    sales_nature 回傳 Nature×銷量帶 12 組（Nature 外層固定順序、帶內層固定順序）；
    om 維度回傳動態 OM 名單（排序後不重複值）；nature_om 回傳完整笛卡兒積
    （Nature 外層固定順序、OM 內層排序）。om_groups 缺省時退回單一
    「空白」群組，呼叫端應傳入實際名單（transform 內由 OM 欄推導，
    Excel 回退路徑由 stats 或 raw 重建）。
    """
    d = _hybrid_dimension(dimension)
    if d == "rp":
        return _HYBRID_RP_GROUPS
    if d == "sales":
        return _HYBRID_SALES_BANDS
    if d == "sales_nature":
        return _HYBRID_SALES_NATURE_GROUPS
    if d == "both":
        return _HYBRID_COMBINED_GROUPS
    if d == "om":
        return tuple(sorted(om_groups)) if om_groups else (_HYBRID_OM_BLANK_LABEL,)
    if d == "nature_om":
        oms = tuple(sorted(om_groups)) if om_groups else (_HYBRID_OM_BLANK_LABEL,)
        return tuple(f"{n}-{o}" for n in _HYBRID_DIAG_NATURES for o in oms)
    return tuple(_HYBRID_DIAG_NATURES)



def _hybrid_zero_sales_groups(dimension) -> tuple[str, ...]:
    """零銷分組名稱（Summary 紅底標示用）：sales 回傳（零銷,）；
    sales_nature 回傳三 Nature 的「{Nature}-零銷」；其餘維度回傳空元組
    （零銷帶只存在於銷量相關維度）。"""
    d = _hybrid_dimension(dimension)
    zero = _HYBRID_SALES_BANDS[3]
    if d == "sales":
        return (zero,)
    if d == "sales_nature":
        return tuple(f"{n}-{zero}" for n in _HYBRID_DIAG_NATURES)
    return ()



# OM 空白群組顯示名（空白 OM 自成一組，令使用者察覺未安排 OM 的店舖）
_HYBRID_OM_BLANK_LABEL = "空白"



def _normalize_hybrid_om(series: pd.Series) -> tuple[pd.Series, tuple[str, ...]]:
    """OM 正規化（混合模式固定規則）：去除前後空格；空白歸入「空白」群組；
    大小寫不敏感（Ivy＝ivy），顯示沿用首見原值。
    回傳 (顯示值序列, 排序後不重複名單)。"""
    raw = series.fillna("").astype(str).str.strip()
    disp = raw.mask(raw == "", _HYBRID_OM_BLANK_LABEL)
    key = disp.str.lower()
    first: dict[str, str] = {}
    for k, d in zip(key.tolist(), disp.tolist()):
        if k not in first:
            first[k] = d
    canon = pd.Series([first[k] for k in key.tolist()], index=series.index)
    return canon, tuple(sorted(set(first.values())))



def _hybrid_om_groups_from_values(values) -> tuple[str, ...]:
    """由 OM 欄值推導排序後不重複的 OM 名單（去除前後空格，空白歸入「空白」）。
    輸入應為已正規化值（transform_hybrid 輸出）；直接傳入大小寫變體時不合併。
    """
    cleaned: list[str] = []
    for v in values:
        text = "" if pd.isna(v) else str(v).strip()
        cleaned.append(text if text else _HYBRID_OM_BLANK_LABEL)
    uniq = sorted(set(cleaned))
    return tuple(uniq) if uniq else (_HYBRID_OM_BLANK_LABEL,)



def _hybrid_skew_thresholds(dimension) -> tuple[float, float]:
    """存貨偏重提示門檻（庫存佔比下限，差距下限），按分類維度切換。"""
    return _HYBRID_SKEW_THRESHOLDS.get(_hybrid_dimension(dimension), (0.40, 0.20))



def _hybrid_skew_hint_text(dimension) -> str:
    """存貨偏重提示門檻的文字描述（供 UI 與文件使用）。"""
    share, gap = _hybrid_skew_thresholds(dimension)
    return f"某分組庫存佔比 ≥ {share:.0%} 且較銷量佔比高 ≥ {gap * 100:.0f} 個百分點"



def _hybrid_shortage_thresholds(dimension) -> tuple[float, float]:
    """存貨偏低標示門檻（銷量佔比下限，差距下限），按分類維度切換。

    與存貨偏重對稱（銷量佔比 − 庫存佔比），但採用較寬鬆門檻，只作儲存格標示。
    """
    return _HYBRID_SHORTAGE_THRESHOLDS.get(_hybrid_dimension(dimension), (0.30, 0.10))



def _hybrid_shortage_hint_text(dimension) -> str:
    """存貨偏低標示門檻的文字描述（供 UI 與文件使用）。"""
    share, gap = _hybrid_shortage_thresholds(dimension)
    return f"某分組銷量佔比 ≥ {share:.0%} 且較庫存佔比高 ≥ {gap * 100:.0f} 個百分點"



# ── Safety 基準周轉天數（Review 線；與調貨段的 14 日線分開，Safety 不跟隨存貨周轉）──
# 預設按 Shop Nature：本地舖 7 日／混合舖 9 日／遊客舖 10 日（可調範圍 7–10 日）。
_HYBRID_SAFETY_DAYS_DEFAULT = {"Tourist": 10, "Mix": 9, "Non Tourist": 7}



_HYBRID_SAFETY_DAYS_MIN = 7



_HYBRID_SAFETY_DAYS_MAX = 10



# Review 欄名（改自「14日周轉所需量」；數值隨所屬 Nature 的基準天數而變）
_HYBRID_SAFETY_LINE_COLUMN = "Safety 周轉所需量"



# 每列實際採用的基準天數（Echo 自參數，供導出檔稽核與回退彙總）
_HYBRID_SAFETY_DAYS_COLUMN = "Safety 基準天數"



# 下調建議補底件數 — SAP MOQ 設定使 Safety 1–2 件難以補貨（I14 原檔無 MOQ 欄），
# 故「Safety 設定過高」列的下調建議最低補至 3 件；現有設定 ≤ 2 件不強制上調
# （補底不得高於現值）、周轉所需量 = 0 的歸零退場建議不受保底影響。
_HYBRID_SAFETY_MIN_KEEP = 3



# 「可補貨（RF）店舖」加總 TurnDay 欄名後綴（RAW／主表顯示用；僅
# `exclude_nd_pool` 開關啟用時輸出）。ND 店舖永不補貨，其庫存屬死貨，
# 不得充當「該 Nature 的可補貨需求已被庫存抵銷」的依據；需求判定改以本欄
# （只計 RF 列）為準，供給側池門檻則維持計全部列（ND 死貨清走不損補貨能力）。
_HYBRID_RF_TURND_SUFFIX = " 加總 TurnDay (RF)"



# 「調貨補足狀態」欄名與三態取值（RAW／主表；僅 `real_supply_gate` 開啟時輸出，
# 非需求列留空）。以不限OM 分配實績為準：
#   已補足＝調入缺口 0；部分補足＝0 < 實際調入量 < 需求量；無貨可調＝實際調入量 0。
# 現行「調入缺口」只分有／無缺口，無法分辨「完全調不到」與「部分補足」，
# 兩者在店舖問題 SKU、指標卡與 Excel 主表都混作同一類下游行動。
_HYBRID_FULFILL_COLUMN = "調貨補足狀態"
_HYBRID_FULFILL_FULL = "已補足"
_HYBRID_FULFILL_PARTIAL = "部分補足"
_HYBRID_FULFILL_NONE = "無貨可調"



# 建議值周轉欄（RAW／主表；僅不足／過高列有值）：建議值 ÷ (P90D 月均 ÷ 30)，
# 與來源「Turnday (P90D Qty)」及 SSR「Recommend Safety Stock Turnday」同一準則。
_HYBRID_SUGGEST_TURNDAY_COLUMN = "建議 Safety 設定值 Turnday"



# ── 全列預防性 Safety 設定建議（不依賴調貨事件；母體為全部店舖×SKU 列）──
# 用途：協助用家一次過設定適合的 Safety Stock，減少日後調貨次數。
# 計算與事件版 Review 同口徑（Safety 周轉所需量＋下調補底 3 件），惟不設調貨事件門檻：
# 現有 < 所需量 → 不足（ND 列不作上調）；現有 > 生效基準 → 過高；其餘 → 設定正常。
# 欄名屬新契約，未經明確要求不得改動。
_HYBRID_FULL_SAFETY_LINE_COLUMN = "全列 Safety 周轉所需量"



_HYBRID_FULL_SAFETY_REVIEW_COLUMN = "全列 Safety Review"



_HYBRID_FULL_SAFETY_SUGGEST_COLUMN = "全列建議 Safety 設定值"



_HYBRID_FULL_SAFETY_CHG_COLUMN = "全列 Safety 建議調整量"



_HYBRID_FULL_SAFETY_TURNDAY_COLUMN = "全列建議 Safety 設定值 Turnday"



_HYBRID_FULL_SAFETY_NORMAL = "設定正常"



_HYBRID_FULL_SAFETY_SHEET = "Safety 設定建議 (全列)"



_HYBRID_FULL_SAFETY_DETAIL_COLUMNS = [
    "Brand", "SKU", "SKU Desc", "Buyer", "Site Code", "Shop Name", "OM", "Shop Nature",
    "RP Type", "Shop Stock Qty (CM)", "Sold Qty (P90D)", "Adj. Safety Stock",
    _HYBRID_FULL_SAFETY_LINE_COLUMN, _HYBRID_FULL_SAFETY_REVIEW_COLUMN,
    _HYBRID_FULL_SAFETY_SUGGEST_COLUMN, _HYBRID_FULL_SAFETY_TURNDAY_COLUMN,
    _HYBRID_FULL_SAFETY_CHG_COLUMN, "事件類型", "Safety Review",
]



# 建議退D001（選開；return_to_d001）：接收目標為 D001（Warehouse，可無限接收）。
# 退貨基準為 Safety Stock (現時)（非調貨段的 14 日線，與池級門檻脫鉤，逐列判定）：
# 中銷／低銷／零銷列退回超出 Safety 的庫存；高銷列原則上不退，僅當 Safety > 0 且
# 庫存 ≥ 3×Safety（高於 Safety 200%）時退至 2×Safety；調貨優先——先扣「轉出量」
# （不限OM 分配實績），「需要調貨」列一律不退。
_HYBRID_RETURN_D001_COLUMN = "建議退D001 量"



_HYBRID_RETURN_D001_GROSS_COLUMN = "可退D001 貨量"



_HYBRID_RETURN_D001_GROSS_INTERNAL = "_hybrid_return_d001_gross"



_HYBRID_RETURN_D001_SHEET = "退D001 建議"



_HYBRID_RETURN_D001_COLUMNS = [
    "Buyer", "Brand", "SKU", "SKU Desc", "Site Code", "Shop Name", "OM", "Shop Nature",
    "RP Type", "Shop Stock Qty (CM)", "Adj. Safety Stock",
    _HYBRID_RETURN_D001_GROSS_COLUMN, "轉出量",
    _HYBRID_RETURN_D001_COLUMN, _HYBRID_SALES_BAND_COLUMN, _HYBRID_SALES_CLASS_COLUMN,
]



def _hybrid_safety_days(safety_days: dict | None = None) -> tuple[dict[str, int], list[str]]:
    """Safety 基準天數 — 以預設值為底，覆寫傳入的整數值。

    未知 Nature、非數值維持忽略（退回預設）；0 以下拒收；範圍外數值箝制進
    7–10 並回報 warning；非整數浮點直接拒絕，避免靜默截斷。
    回傳 (實際天數, warnings)。
    """
    days = dict(_HYBRID_SAFETY_DAYS_DEFAULT)
    warnings: list[str] = []
    for key, val in (safety_days or {}).items():
        name = str(key).strip()
        if name not in days:
            continue
        if isinstance(val, bool):
            raise ValueError(f"Safety 基準天數「{name}」必須為整數，不得為布林值。")
        try:
            numeric = float(val)
        except (TypeError, ValueError):
            continue
        if not np.isfinite(numeric) or not numeric.is_integer():
            raise ValueError(f"Safety 基準天數「{name}」必須為整數。")
        n = int(numeric)
        if n <= 0:
            continue
        if n < _HYBRID_SAFETY_DAYS_MIN or n > _HYBRID_SAFETY_DAYS_MAX:
            clamped = min(max(n, _HYBRID_SAFETY_DAYS_MIN), _HYBRID_SAFETY_DAYS_MAX)
            warnings.append(
                f"Safety 基準天數「{name}」輸入 {n} 日超出建議範圍，已箝制為 {clamped} 日。"
            )
            days[name] = clamped
            continue
        days[name] = n
    return days, warnings



def _hybrid_safety_days_label(days: dict[str, int]) -> str:
    """基準天數的文字描述（供指標卡說明、warning 與 UI 使用）。"""
    return " / ".join(f"{n} {int(days[n])} 日" for n in _HYBRID_DIAG_NATURES)



def _hybrid_numeric_column(
    df: pd.DataFrame,
    column: str,
    index,
    *,
    allow_negative: bool = False,
    nonfinite_as_missing: bool = False,
) -> pd.Series:
    """Hybrid 專用數值欄位清理（只在 hybrid 路徑使用，不動全域 _safe_numeric）。

    空白／非數值保留為 NaN 供各欄 fallback；負數與非有限值依政策處理：
    allow_negative=False 時負數直接拒絕；nonfinite_as_missing=True 時非有限值
    視為缺值（供 D001 fallback），否則拒絕。
    """
    values = pd.to_numeric(df.get(column, _safe_default_series(np.nan, index)), errors="coerce")
    raw_values = values.to_numpy(dtype=float)
    nonfinite_mask = ~np.isnan(raw_values) & ~np.isfinite(raw_values)
    if nonfinite_mask.any():
        if nonfinite_as_missing:
            values = values.mask(pd.Series(nonfinite_mask, index=values.index))
        else:
            raise ValueError(f"「{column}」包含無限大或其他非有限數值，請修正來源資料後再執行。")
    if not allow_negative and bool(((raw_values < 0) & ~np.isnan(raw_values)).any()):
        raise ValueError(f"「{column}」不可為負數，請修正來源資料後再執行。")
    return values



def validate_hybrid_input(df_input: pd.DataFrame) -> list[str]:
    """混合模式必要欄位（含基礎四欄）：缺任一即回報，由呼叫方 raise。"""
    avail = set(df_input.columns)
    return [name for name in _HYBRID_REQUIRED if _resolve_column_name(name, avail) is None]



def _canonicalize_hybrid_columns(df_input: pd.DataFrame) -> pd.DataFrame:
    """混合模式入口：將所有字串欄名 strip 後統一運算（只偵測不改名的情況，
    例如「Buyer 」帶前後空格時先通過驗證、後續取值卻 KeyError）。"""
    df = _standardize_safety_column(df_input)
    mapping: dict[str, str] = {}
    counts: dict[str, int] = {}
    for col in df.columns:
        if isinstance(col, str):
            new = col.strip()
            if new:
                mapping[col] = new
                counts[new] = counts.get(new, 0) + 1
    dup = [k for k, v in counts.items() if v > 1]
    if dup:
        raise ValueError(f"欄位名稱去空白後重複: {', '.join(dup)}")
    if mapping:
        df = df.rename(columns=mapping)
    return df



def _normalize_hybrid_nature(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Shop Nature 正規化（混合模式固定規則）：
    strip + 小寫後 Tourist→Tourist、Mix→Mix；其餘一律（含 Other、空白、大小寫差異）→ Non Tourist。
    回傳 (已正規化 df, 被映射的非標準列數)——映射列數只計「無法辨識」的值
    （小寫去除空白後不在 tourist / mix / non tourist 三者內，如 Other、空白、未知值）。
    """
    df = df.copy()
    raw = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    low = raw.str.lower()
    standard = low.isin(["tourist", "mix", "non tourist"])
    mapped = int((~standard).sum())
    out = pd.Series(
        np.select([low == "mix", low == "tourist"], ["Mix", "Tourist"], default="Non Tourist"),
        index=df.index,
    )
    df["Shop Nature"] = out
    return df, mapped



def _normalize_hybrid_rp_type(df: pd.DataFrame) -> tuple[pd.DataFrame, int, bool]:
    """RP Type 正規化（混合模式固定規則；RF＝可補貨、ND＝不可補貨）：
    strip + 大寫後 RF→RF、ND→ND；其餘（含 Other、空白、大小寫差異）一律→ ND
    （保守歸類為不可補貨：不作調貨需求方，但仍可作供應方）。
    回傳 (已正規化 df, 被映射的非標準列數, 來源是否含「RP Type」欄)——
    欄缺失時不改 df（RAW 不虛構欄位），由呼叫方退回 nature 維度並列 warning。
    """
    if "RP Type" not in df.columns:
        return df, 0, False
    df = df.copy()
    raw = df["RP Type"].fillna("").astype(str).str.strip().str.upper()
    known = raw.isin(["RF", "ND"])
    mapped = int((~known).sum())
    df["RP Type"] = np.where(raw == "RF", "RF", "ND")
    return df, mapped, True



def _transfer_input_consistency_warnings(df: pd.DataFrame) -> list[str]:
    """調貨模式輸入一致性檢核（調貨 Review 與 Stock Checking 共用，只列 warning 不拒絕）：

    - 「Site Code × SKU」重複列：正常輸入粒度為一店×一SKU一列，重複列會令
      同一店舖的庫存與銷量被多次計入；
    - 同一店舖的「Shop Nature」在不同列之間不一致：Shop Nature 應為店舖層級屬性，
      同一店舖不應同時出現 Tourist 與非 Tourist 列；
      調貨分配已排除同店來源，自我調貨（調貨來源＝自己店舖）不會發生。
    - OM 欄存在但全欄空白：同OM分組無法區分店群，結果會等同全部店舖使用同一群組。
    """
    out: list[str] = []
    if "Site Code" not in df.columns or "SKU" not in df.columns:
        return out
    site = df["Site Code"].fillna("").astype(str).str.strip()
    sku = df["SKU"].fillna("").astype(str).str.strip()
    key = site + "\u0000" + sku
    sizes = key.groupby(key).size()
    dup_sizes = sizes[sizes > 1]
    if len(dup_sizes):
        out.append(
            f"偵測到「Site Code × SKU」重複列 {len(dup_sizes)} 組（共 {int(dup_sizes.sum())} 列；"
            "正常輸入應為一店×一SKU一列），重複列會令同一店舖的庫存與銷量被多次計入；"
            "重複列的 Site Code 不同時，自我調貨守衛（以 Site Code 比對）無法辨識同店，"
            "同一實體店舖可能同時出現在調貨來源與接收端，請檢查來源檔案。"
        )
    nature = df.get("Shop Nature", _safe_default_series("", df.index))
    nature_norm = nature.fillna("").astype(str).str.strip().str.lower()
    n_unique = nature_norm.groupby(site).nunique()
    bad_sites = n_unique[n_unique > 1]
    if len(bad_sites):
        shown = "、".join(str(s) for s in list(bad_sites.index[:5]))
        suffix = f" 等 {len(bad_sites)} 間店舖" if len(bad_sites) > 5 else ""
        out.append(
            f"店舖 {shown}{suffix} 的「Shop Nature」在不同列之間不一致；"
            "Shop Nature 應為店舖層級固定屬性，請檢查來源資料。"
        )
    if "OM" in df.columns:
        om = df["OM"].fillna("").astype(str).str.strip()
        if len(df) and bool((om == "").all()):
            out.append("來源檔案的「OM」欄全欄空白，同OM調貨結果會等同全部店舖使用同一群組。")
    return out




def _attach_sales_classification_columns(
    df: pd.DataFrame,
    sales_params: dict,
    *,
    monthly_sold: pd.Series | None = None,
) -> tuple[pd.DataFrame, list[str]]:
    """為模式 RAW 附加「銷量帶」與「銷量存貨分類」兩欄（Stock Checking／Safety Stock
    Review 引入用；hybrid 在 transform_hybrid 內自行計算，不經此函數）。

    - 銷量帶需「Sold Qty (P90D)」與「SKU」；缺任一欄時兩欄皆不輸出並列 warning。
    - 銷量存貨分類另需「Shop Stock Qty (CM)」與店舖周轉日數（庫存 ÷ 月銷 × 30.5）；
      月銷優先「Shop Monthly Sold Qty (P90D)」，缺欄時以 Sold Qty (P90D) ÷ 90 × 30.5
      換算（與 Stock Checking 同一準則）並列 warning；缺庫存欄時只輸出銷量帶並列 warning。
    monthly_sold 已由呼叫方解像時傳入，避免重複 warning。回傳 (附加後 df, warnings)。
    """
    warnings: list[str] = []
    df = df.copy()
    if "Sold Qty (P90D)" not in df.columns or "SKU" not in df.columns:
        warnings.append("來源檔案缺少「Sold Qty (P90D)」，無法計算銷量帶與銷量存貨分類。")
        return df, warnings
    sold_p90d = _safe_numeric(df["Sold Qty (P90D)"])
    df[_HYBRID_SALES_BAND_COLUMN] = _hybrid_sales_bands(sold_p90d, df["SKU"])
    if "Shop Stock Qty (CM)" not in df.columns:
        warnings.append("來源檔案缺少「Shop Stock Qty (CM)」，無法計算銷量存貨分類（銷量帶已輸出）。")
        return df, warnings
    if monthly_sold is None:
        if "Shop Monthly Sold Qty (P90D)" in df.columns:
            monthly_sold = _safe_numeric(df["Shop Monthly Sold Qty (P90D)"])
        else:
            monthly_sold = sold_p90d / 90.0 * 30.5
            warnings.append(
                "來源檔案缺少「Shop Monthly Sold Qty (P90D)」，銷量存貨分類的店舖周轉日數改以 Sold Qty (P90D) 換算。"
            )
    msold_val = np.asarray(monthly_sold, dtype=float)
    cm_stock_val = _safe_numeric(df["Shop Stock Qty (CM)"]).to_numpy(dtype=float)
    with np.errstate(divide="ignore", invalid="ignore"):
        shop_turn_day = np.where(
            msold_val > 0,
            cm_stock_val / np.where(msold_val > 0, msold_val, 1.0) * 30.5,
            np.inf,
        )
    df[_HYBRID_SALES_CLASS_COLUMN] = _hybrid_sales_classification(
        df[_HYBRID_SALES_BAND_COLUMN], shop_turn_day, cm_stock_val, sales_params,
    )
    return df, warnings



# 「調貨 Review Safety Stock」主表欄序（店舖×SKU 逐列；RAW 則輸出全部欄位 + Review 欄）
# RP Type 為正規化後值（RF／ND），來源缺「RP Type」欄時該欄不出現（主表退回 28 欄）
# 調貨補足狀態不在本契約內：僅 `real_supply_gate` 開啟時由 _build_hybrid_diagnosis
# 插入於「調入需求量 (補到14日)」之後（主表 29→30 欄），故本清單恆為 29 欄基準契約。
_HYBRID_DIAG_COLUMNS = [
    "Brand", "SKU", "SKU Desc", "Buyer", "Adj. SST", "Site Code", "OM", "Shop Nature", "RP Type",
    "事件類型", "Shop Stock Qty (CM)", "Shop Monthly Sold Qty (P90D)",
    _HYBRID_SALES_BAND_COLUMN, _HYBRID_SALES_CLASS_COLUMN,
    _HYBRID_SAFETY_DAYS_COLUMN, _HYBRID_SAFETY_LINE_COLUMN, "Adj. Safety Stock", "Safety Review",
    "建議 Safety 設定值", _HYBRID_SUGGEST_TURNDAY_COLUMN, "Safety 建議調整量",
    "調入需求量 (補到14日)", "實際調入量", "調入缺口",
    "可轉出貨量", "轉出量", "調貨來源 (Site Code)",
    "調貨後庫存", "調貨後 TurnDay",
]



# RAW 顯示欄序分組（識別→品項屬性→庫存與銷售→SKU 級判定→Review 結果→調貨 不限OM→調貨 同OM→其他）
_HYBRID_RAW_DISPLAY_GROUPS = [
    ["Buyer", "Brand", "SKU", "SKU Desc", "Site Code", "Shop Name", "OM", "Shop Nature"],
    ["Adj. SST", "Adj. MCH1", "Vendor MC", "MCH2", "Item Status", "Item Class", "RP Type",
     "Shop Stock Status", "Supply Source", "Launch Date"],
    ["Shop Stock Qty (CM)", "Adj. Safety Stock", "Sold Qty (P90D)", "Shop Monthly Sold Qty (P90D)",
     _HYBRID_SALES_BAND_COLUMN, _HYBRID_SALES_CLASS_COLUMN,
     "Turnday (P90D Qty)", "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
     "Net Sales (M-2)", "Net Sales (M-1)", "Net Sales (CM - MTD)", "Excess Stock Qty",
     "Shop On Order PO Qty (CM)", "Shop Pending Receive",
     "Effective Shop Monthly Sold Qty (P90D)",
     "D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)", "D001 Effective Stock Qty (CM)"],
    ["Tourist 加總 TurnDay", "Mix 加總 TurnDay", "Non Tourist 加總 TurnDay",
     "Tourist 加總 TurnDay (RF)", "Mix 加總 TurnDay (RF)", "Non Tourist 加總 TurnDay (RF)",
     "需要調貨", "調入需求量 (補到14日)", _HYBRID_FULFILL_COLUMN],
    ["事件類型", _HYBRID_SAFETY_DAYS_COLUMN, _HYBRID_SAFETY_LINE_COLUMN,
     "Safety Review", "建議 Safety 設定值", _HYBRID_SUGGEST_TURNDAY_COLUMN, "Safety 建議調整量",
     _HYBRID_FULL_SAFETY_LINE_COLUMN, _HYBRID_FULL_SAFETY_REVIEW_COLUMN,
     _HYBRID_FULL_SAFETY_SUGGEST_COLUMN, _HYBRID_FULL_SAFETY_TURNDAY_COLUMN,
     _HYBRID_FULL_SAFETY_CHG_COLUMN],
    ["可轉出貨量 (Non Tourist)", "可轉出貨量 (Mix)", "可轉出貨量 (Tourist)",
     "實際調入量", "調入缺口", "轉出量", _HYBRID_RETURN_D001_COLUMN,
     "調貨來源 (Site Code)", "調貨後庫存", "調貨後 TurnDay"],
    ["實際調入量 (同OM)", "調入缺口 (同OM)", "轉出量 (同OM)", "調貨來源 (同OM)",
     "調貨後庫存 (同OM)", "調貨後 TurnDay (同OM)"],
]



def _hybrid_p90d_zero_sales_mask(raw: pd.DataFrame, ref_date=None) -> tuple[pd.Series, list[str]]:
    """混合模式零銷售轉 ND 候選遮罩（僅店舖問題表提示用，不改主表）。

    準則與獨立 Past 90 Days 0 Sales 模式一致：四期銷售皆 ≤ 0 且店舖有庫存
    （NaN 視為 0；負數視為 ≤ 0；`Shop Stock Qty (LM)` 欄存在時須 LM／CM 皆 > 0，
    缺欄時僅以 CM > 0 判定）。另作新品保護（`Launch Date` 距參考日期 90 天內排除，
    重用 `_parse_launch_dates`；缺欄／缺值／無法解析視為非新品；參考日期預設為今天，
    建議傳入 MTD 截止日期）與防呆排除
    （`需要調貨=是` 列、非 RF 列不納入；`RP Type` 缺欄時沿用全 RF 處理）。
    四個銷售欄任一缺失時回傳全 False＋一條 warning（不停止、不拋錯）。
    """
    idx = raw.index
    empty = pd.Series(False, index=idx)
    warnings: list[str] = []
    if raw is None or raw.empty:
        return empty, warnings
    avail = set(raw.columns)
    missing = [c for c in _HYBRID_P90D_ZERO_SALES_COLUMNS if _resolve_column_name(c, avail) is None]
    if missing:
        warnings.append(
            f"缺少「{'／'.join(missing)}」欄位，無法判定 Past 90 Days 0 Sales 轉 ND 候選（店舖問題 SKU 不顯示 ND 退場提示）。"
        )
        return empty, warnings
    sold = pd.concat(
        [_safe_numeric(raw[c], _safe_default_series(0, idx)) for c in _HYBRID_P90D_ZERO_SALES_COLUMNS],
        axis=1,
    )
    stock_cm = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, idx)))
    mask = (sold.le(0).all(axis=1) & (stock_cm > 0)).fillna(False)
    if "Shop Stock Qty (LM)" in raw.columns:
        stock_lm = _safe_numeric(raw.get("Shop Stock Qty (LM)", _safe_default_series(0, idx)))
        mask &= (stock_lm > 0).fillna(False)
    if "Launch Date" in raw.columns:
        launch_dt = _parse_launch_dates(raw["Launch Date"].fillna("").astype(str).str.strip())
        _ref = _normalize_ref_date(ref_date)
        is_new_sku = launch_dt.notna() & (
            launch_dt >= pd.Timestamp(_ref) - pd.Timedelta(days=90)
        )
        mask &= ~is_new_sku.fillna(False).to_numpy()
    need = raw.get("需要調貨", _safe_default_series("", idx)).fillna("").astype(str).str.strip() == "是"
    mask &= ~need.fillna(False).to_numpy()
    if "RP Type" in raw.columns:
        rp = raw["RP Type"].fillna("").astype(str).str.strip().str.upper()
        mask &= (rp == "RF").fillna(False).to_numpy()
    return pd.Series(mask.to_numpy(), index=idx), warnings



# ── Past 90 Days 0 Sales 轉 ND 候選（混合模式店舖問題表提示用）──
_HYBRID_P90D_ZERO_SALES_COLUMNS = [
    "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)", "Sold Qty (P90D)",
]



def _hybrid_group_key_series(df: pd.DataFrame, dimension: str, om_groups=None) -> pd.Series:
    """依分類維度產生每列的分組鍵（nature→Shop Nature；rp→RP Type；both→Nature-RP；
    om→OM；nature_om→Nature-OM；sales→銷量帶；sales_nature→Nature-銷量帶）。
    RP Type 為正規化後值；缺欄時一律按 RF（可補貨）處理，與 transform_hybrid 同一準則。
    OM 空白歸入「空白」，與 transform_hybrid 同一準則（om_groups 僅決定欄位集合，
    不影響逐列鍵值）。銷量帶為 transform_hybrid 常開產出欄；缺欄時退回
    Shop Nature 鍵（顯式守衛，正常輸入不會觸發，sales 與 sales_nature 同規則）。"""
    d = _hybrid_dimension(dimension)
    nature = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    if d == "nature":
        return nature
    if d in ("sales", "sales_nature"):
        if _HYBRID_SALES_BAND_COLUMN in df.columns:
            band = df[_HYBRID_SALES_BAND_COLUMN].fillna("").astype(str).str.strip()
            band = band.where(band.isin(_HYBRID_SALES_BANDS), _HYBRID_SALES_BANDS[1])
            return nature + "-" + band if d == "sales_nature" else band
        return nature
    if d in ("om", "nature_om"):
        om = (
            df["OM"].fillna("").astype(str).str.strip()
            if "OM" in df.columns
            else _safe_default_series("", df.index)
        )
        om = pd.Series(np.where(om == "", _HYBRID_OM_BLANK_LABEL, om), index=df.index)
        if d == "om":
            return om
        return nature + "-" + om
    rp = (
        df["RP Type"].fillna("").astype(str).str.strip().str.upper()
        if "RP Type" in df.columns
        else _safe_default_series("RF", df.index)
    )
    rp = pd.Series(np.where(rp == "RF", "RF", "ND"), index=df.index)
    if d == "rp":
        return rp
    return nature + "-" + rp
