# -*- coding: utf-8 -*-
"""Forecast Review模式：由 transformer.py 拆分而來。
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
    _compute_raw_extra_columns,
    _iter_header_rows,
    _parse_launch_dates,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _cached_safe_sku_series,
    _write_metric_cards,
    _write_raw_placeholder,
    _write_raw_sheet,
    validate_input,
)


from modes.framework import (
    _transfer_input_consistency_warnings,
)


from modes.stock_impact import (
    _add_stock_impact_columns,
)
__all__ = [
    "_FC_ACCEL_MULT_DEFAULT",
    "_FC_ACCEL_MULT_MAX",
    "_FC_ACCEL_MULT_MIN",
    "_FC_CAP_MONTHS_DEFAULT",
    "_FC_CAP_MONTHS_MAX",
    "_FC_CAP_MONTHS_MIN",
    "_FC_CLASS_ACCEL",
    "_FC_CLASS_NEW",
    "_FC_CLASS_STABLE",
    "_FC_CLASS_SURGE",
    "_FC_COVERAGE_RISK_DAYS",
    "_FC_DAMP_DEFAULT",
    "_FC_DAMP_MAX",
    "_FC_DAMP_MIN",
    "_FC_LAUNCH_MIN_ACTIVE_DAYS",
    "_FC_RISK_7D",
    "_FC_RISK_OK",
    "_FC_RISK_OOS",
    "_FC_RISK_TARGET",
    "_FC_STATUS_ENOUGH",
    "_FC_STATUS_ND",
    "_FC_STATUS_ORDER",
    "_FC_STATUS_ORDER_CAPPED",
    "_FC_SURGE_COVER_DEFAULT",
    "_FC_SURGE_COVER_MAX",
    "_FC_SURGE_COVER_MIN",
    "_FC_SURGE_MIN_BASE_DEFAULT",
    "_FC_SURGE_MIN_BASE_MAX",
    "_FC_SURGE_MIN_BASE_MIN",
    "_FC_SURGE_MULT_DEFAULT",
    "_FC_SURGE_MULT_MAX",
    "_FC_SURGE_MULT_MIN",
    "_FC_TARGET_D001",
    "_FC_TARGET_DAYS_DEFAULT",
    "_FC_TARGET_DAYS_MAX",
    "_FC_TARGET_DAYS_MIN",
    "_FC_TARGET_SHOP",
    "_FC_TREND_CAP_DEFAULT",
    "_FC_TREND_CAP_MAX",
    "_FC_TREND_CAP_MIN",
    "_FC_WEIGHT_M1_DEFAULT",
    "_FC_WEIGHT_MAX",
    "_FC_WEIGHT_MIN",
    "_FC_WEIGHT_MTD_DEFAULT",
    "_FC_WEIGHT_P90D_DEFAULT",
    "_FORECAST_BUYER_COLUMNS",
    "_FORECAST_COVERAGE_COLUMNS",
    "_FORECAST_D001_PO_COL",
    "_FORECAST_MAIN_COLUMNS",
    "_FORECAST_PENDING_RECEIVE_CANDIDATES",
    "_FORECAST_RAW_DISPLAY_GROUPS",
    "_FORECAST_REQUIRED",
    "_FORECAST_SHOP_PO_COL",
    "_fc_clamp_number",
    "_fc_launch_period_factors",
    "_fc_resolve_first",
    "_fc_rp_series",
    "_fc_supply_source_series",
    "_forecast_main_columns",
    "_forecast_raw_display_columns",
    "_write_forecast_guide_sheet",
    "_write_forecast_sheet",
    "generate_excel_forecast",
    "transform_forecast",
]


# ── Forecast 參數預設值（UI 範圍見 analyzer_page；此處為函數層預設）──
_FC_TARGET_DAYS_DEFAULT = 30



_FC_TARGET_DAYS_MIN = 7



_FC_TARGET_DAYS_MAX = 90



_FC_CAP_MONTHS_DEFAULT = 2.0



_FC_CAP_MONTHS_MIN = 0.5



_FC_CAP_MONTHS_MAX = 6.0



_FC_SURGE_MULT_DEFAULT = 1.5



_FC_SURGE_MULT_MIN = 1.1



_FC_SURGE_MULT_MAX = 5.0



_FC_SURGE_MIN_BASE_DEFAULT = 10



_FC_SURGE_MIN_BASE_MIN = 0



_FC_SURGE_MIN_BASE_MAX = 100



_FC_ACCEL_MULT_DEFAULT = 1.15



_FC_ACCEL_MULT_MIN = 1.0



_FC_ACCEL_MULT_MAX = 3.0



_FC_DAMP_DEFAULT = 0.8



_FC_DAMP_MIN = 0.5



_FC_DAMP_MAX = 1.0



_FC_COVERAGE_RISK_DAYS = 7  # 店舖覆蓋缺貨風險線（日），固定



# v2 需求模型（加權混合＋趨勢外推）：三項權重（會自動正規化）與趨勢增速上限
_FC_WEIGHT_P90D_DEFAULT = 0.4



_FC_WEIGHT_M1_DEFAULT = 0.3



_FC_WEIGHT_MTD_DEFAULT = 0.3



_FC_WEIGHT_MIN = 0.0



_FC_WEIGHT_MAX = 1.0



_FC_TREND_CAP_DEFAULT = 1.0



_FC_TREND_CAP_MIN = 0.2



_FC_TREND_CAP_MAX = 3.0



# 開賣日折算的最少有效日數（防開賣首日極端外推；月長上限令 factor 最多約 4.4）
_FC_LAUNCH_MIN_ACTIVE_DAYS = 7



# 爆升類（新起量／加速中／爆升）目標覆蓋月數：熱銷 SKU 預留較長覆蓋，
# 目標覆蓋量＝ceil(月預估 × 月數)；平穩類沿用目標覆蓋天數；訂量仍受上限月數封頂。
_FC_SURGE_COVER_DEFAULT = 2.0



_FC_SURGE_COVER_MIN = 0.5



_FC_SURGE_COVER_MAX = 6.0



# 本模式嚴格檢核欄（缺任一即停止；共用必要欄位另由 validate_input 檢核）
_FORECAST_REQUIRED = [
    "Site Code",
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Sold Qty (P90D)",
    "Shop Stock Qty (CM)",
]



# Shop Pending Receive 欄名相容：實檔表頭為「Shop Pending Receive (±7 days)」，
# 與系統已知「Shop Pending Receive」字串不同，按序取首個命中。
_FORECAST_PENDING_RECEIVE_CANDIDATES = (
    "Shop Pending Receive (±7 days)",
    "Shop Pending Receive",
)



_FORECAST_D001_PO_COL = "D001 On Order PO Qty (CM)"



_FORECAST_SHOP_PO_COL = "Shop On Order PO Qty (CM)"



# 爆升分類標籤
_FC_CLASS_NEW = "新起量"



_FC_CLASS_ACCEL = "加速中"



_FC_CLASS_SURGE = "爆升"



_FC_CLASS_STABLE = "平穩"



# SKU 狀態標籤
_FC_STATUS_ORDER = "建議訂貨"



_FC_STATUS_ORDER_CAPPED = "建議訂貨（已封頂）"



_FC_STATUS_ENOUGH = "倉位足夠"



_FC_STATUS_ND = "ND 不建議訂貨"



# 建議訂貨對象標籤（資料契約：固定表頭，未經明確要求不得改動）
_FC_TARGET_D001 = "D001 PO"



_FC_TARGET_SHOP = "Shop PO"



# 店舖覆蓋風險旗
_FC_RISK_OOS = "已缺貨"



_FC_RISK_7D = "≤7日"



_FC_RISK_TARGET = "≤目標天數"



_FC_RISK_OK = "足夠"



def _fc_resolve_first(candidates: tuple[str, ...], avail: set[str]) -> str | None:
    """按序取首個命中的欄名（供 Pending Receive 等別名相容用）。"""
    for cand in candidates:
        hit = _resolve_column_name(cand, avail)
        if hit is not None:
            return hit
    return None



def _fc_clamp_number(value, default, lo, hi, warnings, label):
    """Forecast 數值參數正規化：非數值退回預設＋warning；範圍外箝制＋warning。"""
    try:
        v = float(value)
    except (TypeError, ValueError):
        warnings.append(f"「{label}」無法解析為數值，改用預設 {default:g}。")
        return default
    if not np.isfinite(v):
        warnings.append(f"「{label}」為非有限值，改用預設 {default:g}。")
        return default
    if v < lo or v > hi:
        warnings.append(f"「{label}」超出範圍（{lo:g}–{hi:g}），已箝制。")
        return min(max(v, lo), hi)
    return v



def _fc_supply_source_series(df: pd.DataFrame, idx, warnings) -> pd.Series:
    """Supply Source 正規化：缺欄全視為 Warehouse＋warning（與 RP 缺欄全 RF 對稱）。"""
    if "Supply Source" not in df.columns:
        warnings.append("來源檔案缺少「Supply Source」欄，全部列按 Warehouse 處理（位置計 D001 庫存＋D001 PO＋在途）。")
        return pd.Series("Warehouse", index=idx)
    return df["Supply Source"].fillna("").astype(str).str.strip()



def _fc_rp_series(df: pd.DataFrame, idx, warnings) -> pd.Series:
    """RP Type 正規化：缺欄全 RF＋warning（沿用 hybrid 慣例）；其餘值歸 ND。"""
    if "RP Type" not in df.columns:
        warnings.append("來源檔案缺少「RP Type」欄，全部列按 RF 處理。")
        return pd.Series("RF", index=idx)
    rp = df["RP Type"].fillna("").astype(str).str.strip().str.upper()
    return rp.where(rp == "RF", "ND")



def _fc_launch_period_factors(
    df: pd.DataFrame,
    sku_key: pd.Series,
    mtd_date,
    mtd_days,
    month_total_days,
    warnings: list[str],
) -> tuple[dict[str, dict], int]:
    """v2 開賣日折算：按 SKU 最早 Launch Date 計 M-2／M-1 折算因子與本月有效日數。

    回傳 (facts, prorated_skus)。facts 各 dict 以 SKU 為鍵：
    factor_m2／factor_m1（≥1，未折算為 1）與 denom_cm（本月有效日數，未折算為 mtd_days）。
    無 Launch Date 欄、mtd_date 無效或 mtd_days ≤ 0 時全部不折算。
    """
    facts: dict[str, dict] = {"factor_m2": {}, "factor_m1": {}, "denom_cm": {}}
    default_denom = float(mtd_days) if mtd_days and mtd_days > 0 else 0.0
    try:
        mtd_ts = pd.Timestamp(mtd_date).normalize() if mtd_date is not None else pd.NaT
    except (TypeError, ValueError):
        mtd_ts = pd.NaT
    if (
        "Launch Date" not in df.columns
        or pd.isna(mtd_ts)
        or mtd_days is None
        or int(mtd_days) <= 0
    ):
        return facts, 0

    cur = mtd_ts
    cm_start = cur.replace(day=1)
    prev_end = cm_start - pd.Timedelta(days=1)
    prev_start = prev_end.replace(day=1)
    prev2_end = prev_start - pd.Timedelta(days=1)
    prev2_start = prev2_end.replace(day=1)
    mtd_days_int = int(mtd_days)

    launch_raw = df["Launch Date"].fillna("").astype(str).str.strip()
    launch_dt = pd.Series(_parse_launch_dates(launch_raw).values, index=df.index)
    earliest = launch_dt.groupby(sku_key.values).min()

    def _period_factor(launch, p_start, p_end) -> float:
        if pd.isna(launch):
            return 1.0
        l = pd.Timestamp(launch).normalize()
        if l < p_start or l > p_end:
            return 1.0
        active = int((p_end - l).days) + 1
        days = int((p_end - p_start).days) + 1
        return max(1.0, days / max(active, _FC_LAUNCH_MIN_ACTIVE_DAYS))

    def _cm_denom(launch) -> float:
        if pd.isna(launch):
            return default_denom
        l = pd.Timestamp(launch).normalize()
        if l > cur or l <= cm_start:
            return default_denom
        active = min(mtd_days_int, int((cur - l).days) + 1)
        return float(max(active, _FC_LAUNCH_MIN_ACTIVE_DAYS))

    prorated = 0
    for sku_val, launch in earliest.items():
        f_m2 = _period_factor(launch, prev2_start, prev2_end)
        f_m1 = _period_factor(launch, prev_start, prev_end)
        d_cm = _cm_denom(launch)
        facts["factor_m2"][sku_val] = f_m2
        facts["factor_m1"][sku_val] = f_m1
        facts["denom_cm"][sku_val] = d_cm
        if f_m2 != 1.0 or f_m1 != 1.0 or d_cm != default_denom:
            prorated += 1
    return facts, prorated



def transform_forecast(
    df_input: pd.DataFrame,
    mtd_days: int | None = None,
    month_total_days: int | None = None,
    target_days: float = _FC_TARGET_DAYS_DEFAULT,
    cap_months: float = _FC_CAP_MONTHS_DEFAULT,
    surge_mult: float = _FC_SURGE_MULT_DEFAULT,
    surge_min_base: float = _FC_SURGE_MIN_BASE_DEFAULT,
    accel_mult: float = _FC_ACCEL_MULT_DEFAULT,
    damp: float = _FC_DAMP_DEFAULT,
    w_p90d: float = _FC_WEIGHT_P90D_DEFAULT,
    w_m1: float = _FC_WEIGHT_M1_DEFAULT,
    w_mtd: float = _FC_WEIGHT_MTD_DEFAULT,
    trend_cap: float = _FC_TREND_CAP_DEFAULT,
    surge_cover_months: float = _FC_SURGE_COVER_DEFAULT,
    mtd_date=None,
    progress_callback=None,
):
    """Forecast Review 核心轉換（SKU 級建議訂貨量；第二維度防缺貨）。

    v2 需求模型（SKU 級聚合後預估）：
    逐列缺貨還原（M-2／M-1／CM 缺貨列取代式還原至該店月均基準）→ SKU 合計＋
    開賣日折算（Launch Date 落在期間內按有效日數折算）→ MTD 可信度加權
    （MTD 月化與月均基準按已過日數混合）→ 標準情境＝三項加權混合
    （月均基準／M-1／MTD 調整月化）→ 樂觀情境＝趨勢外推（增速上限 trend_cap，
    再乘衰減係數 damp，且不低於標準）→ SKU 級爆升分類（新起量／加速中／爆升
    用樂觀、平穩用標準）→ 位置（Warehouse／Vendor 分流）→ 目標覆蓋與上限封頂
    → SKU 主表＋店舖覆蓋表＋Buyer 彙總＋stats。

    目標覆蓋：平穩類＝ceil(月預估 ÷ 30.5 × target_days)；爆升類（新起量／加速中／
    爆升）＝ceil(月預估 × surge_cover_months)，仍受上限月數 cap_months 封頂。
    damp 為趨勢外推的衰減係數（樂觀＝max(S1, MTD 調整月化) × (1 + damp × g)）。
    MTD 日期無效時 MTD 月化退回 max(M-1, 月均基準) 並列 warning。
    另輸出 4 個月 FC 拆分參考欄：本月 (M)＝MTD 調整月化；M+1＝月預估 (採用)；
    爆升類 M+2／M+3 按趨勢增速 × 衰減係數逐月遞推，平穩類與 M+1 相同。

    回傳 (main_df, coverage_df, buyer_df, raw_df, stats, warnings)。
    stats 另回存全部正規化後參數，供 Excel「說明」表照錄。
    """
    warnings: list[str] = []
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")

    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")
    avail = set(df_input.columns)
    fc_missing = [name for name in _FORECAST_REQUIRED if _resolve_column_name(name, avail) is None]
    if fc_missing:
        raise ValueError(f"缺少必需欄位: {', '.join(fc_missing)}")

    if progress_callback:
        progress_callback(0.1, "清理資料中...")
    df = _clean_dataframe(df_input, shop_natures=None, progress_callback=progress_callback)
    if df.empty:
        raise ValueError("清理後無可用資料列（可能僅含 Total 列或註腳列），無法計算 Forecast。")

    if progress_callback:
        progress_callback(0.2, "解像 SKU Desc / Shop Name...")
    df = _compute_raw_extra_columns(df)

    idx = df.index
    if "SKU" in df.columns:
        df["SKU"] = df["SKU"].apply(_safe_sku)
    warnings.extend(_transfer_input_consistency_warnings(df))

    # ── 參數正規化 ──────────────────────────────────────────────────────────
    target_days = _fc_clamp_number(target_days, _FC_TARGET_DAYS_DEFAULT,
                                   _FC_TARGET_DAYS_MIN, _FC_TARGET_DAYS_MAX, warnings, "目標覆蓋天數")
    cap_months = _fc_clamp_number(cap_months, _FC_CAP_MONTHS_DEFAULT,
                                  _FC_CAP_MONTHS_MIN, _FC_CAP_MONTHS_MAX, warnings, "上限月數")
    surge_mult = _fc_clamp_number(surge_mult, _FC_SURGE_MULT_DEFAULT,
                                  _FC_SURGE_MULT_MIN, _FC_SURGE_MULT_MAX, warnings, "爆升倍數")
    surge_min_base = _fc_clamp_number(surge_min_base, _FC_SURGE_MIN_BASE_DEFAULT,
                                      _FC_SURGE_MIN_BASE_MIN, _FC_SURGE_MIN_BASE_MAX, warnings, "新起量基數")
    accel_mult = _fc_clamp_number(accel_mult, _FC_ACCEL_MULT_DEFAULT,
                                  _FC_ACCEL_MULT_MIN, _FC_ACCEL_MULT_MAX, warnings, "加速倍數")
    damp = _fc_clamp_number(damp, _FC_DAMP_DEFAULT, _FC_DAMP_MIN, _FC_DAMP_MAX, warnings, "衰減係數")
    w_p90d = _fc_clamp_number(w_p90d, _FC_WEIGHT_P90D_DEFAULT, _FC_WEIGHT_MIN, _FC_WEIGHT_MAX,
                              warnings, "月均基準權重")
    w_m1 = _fc_clamp_number(w_m1, _FC_WEIGHT_M1_DEFAULT, _FC_WEIGHT_MIN, _FC_WEIGHT_MAX,
                            warnings, "M-1 權重")
    w_mtd = _fc_clamp_number(w_mtd, _FC_WEIGHT_MTD_DEFAULT, _FC_WEIGHT_MIN, _FC_WEIGHT_MAX,
                             warnings, "MTD 權重")
    w_sum = w_p90d + w_m1 + w_mtd
    if w_sum <= 0:
        warnings.append("三項模型權重合計為 0，已改用預設 0.4／0.3／0.3。")
        w_p90d, w_m1, w_mtd = _FC_WEIGHT_P90D_DEFAULT, _FC_WEIGHT_M1_DEFAULT, _FC_WEIGHT_MTD_DEFAULT
    else:
        w_p90d, w_m1, w_mtd = w_p90d / w_sum, w_m1 / w_sum, w_mtd / w_sum
    trend_cap = _fc_clamp_number(trend_cap, _FC_TREND_CAP_DEFAULT, _FC_TREND_CAP_MIN,
                                 _FC_TREND_CAP_MAX, warnings, "趨勢增速上限")
    surge_cover_months = _fc_clamp_number(surge_cover_months, _FC_SURGE_COVER_DEFAULT,
                                          _FC_SURGE_COVER_MIN, _FC_SURGE_COVER_MAX,
                                          warnings, "爆升目標覆蓋月數")

    # ── 銷售／庫存數值（負銷售按 0＋warning；庫存負數／非有限值按 0＋warning）──
    m2 = _safe_numeric(df.get("Sold Qty (M-2)", _safe_default_series(0, idx)))
    m1 = _safe_numeric(df.get("Sold Qty (M-1)", _safe_default_series(0, idx)))
    mtd = _safe_numeric(df.get("Sold Qty (CM - MTD)", _safe_default_series(0, idx)))
    sold_p90d = _safe_numeric(df.get("Sold Qty (P90D)", _safe_default_series(0, idx)))
    cm_stock_raw = pd.to_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, idx)), errors="coerce")
    if bool(((m2 < 0) | (m1 < 0) | (mtd < 0) | (sold_p90d < 0)).any()):
        warnings.append("部分銷售欄位出現負數，已按 0 處理。")
    m2 = m2.clip(lower=0)
    m1 = m1.clip(lower=0)
    mtd = mtd.clip(lower=0)
    sold_p90d = sold_p90d.clip(lower=0)
    if bool((cm_stock_raw.isna() | (cm_stock_raw < 0)).any()):
        warnings.append("部分「Shop Stock Qty (CM)」為負數或非數值，已按 0 處理。")
    cm_stock = cm_stock_raw.fillna(0.0).clip(lower=0)

    # ── 月均基準 p90d_row：Shop Monthly Sold Qty (P90D)；缺欄退回 Sold Qty (P90D)÷90×30.5 ──
    if _resolve_column_name("Shop Monthly Sold Qty (P90D)", avail) is not None:
        p90d_row = _safe_numeric(df["Shop Monthly Sold Qty (P90D)"]).clip(lower=0)
    else:
        p90d_row = (sold_p90d / 90.0 * 30.5).clip(lower=0)
        warnings.append("來源檔案缺少「Shop Monthly Sold Qty (P90D)」，月均基準改以 Sold Qty (P90D) ÷ 90 × 30.5 換算。")

    # ── v2：缺貨狀態（重用引擎欄；M-2／M-1 缺 stock 欄時不參與還原）──
    if progress_callback:
        progress_callback(0.3, "逐列缺貨還原中...")
    df = _add_stock_impact_columns(df)
    is_oos_cm = df.get("Is OOS (CM)", _safe_default_series(False, idx))
    if not pd.api.types.is_bool_dtype(is_oos_cm):
        is_oos_cm = is_oos_cm.astype(bool)
    oos_cm = is_oos_cm.to_numpy(dtype=bool)
    rp = _fc_rp_series(df, idx, warnings)
    is_rf = (rp == "RF").to_numpy(dtype=bool)
    df["FC 缺貨還原列"] = (oos_cm & is_rf)
    df["FC RP Type (正規)"] = rp.values
    sku_key = df["SKU"] if "SKU" in df.columns else _safe_default_series(None, idx)

    p90d_arr = p90d_row.to_numpy(dtype=float)
    m1_arr = m1.to_numpy(dtype=float)
    m2_arr = m2.to_numpy(dtype=float)
    mtd_arr = mtd.to_numpy(dtype=float)

    mtd_ok = (
        mtd_days is not None and month_total_days is not None
        and mtd_days > 0 and month_total_days > 0
    )
    _denom_default = float(mtd_days) if mtd_ok else 0.0

    # ── v2：開賣日折算（Launch Date 落在 M-2／M-1／本月時按有效日數折算）──
    launch_facts, prorated_skus = _fc_launch_period_factors(
        df, sku_key, mtd_date, mtd_days, month_total_days, warnings
    )
    factor_m2 = sku_key.map(launch_facts["factor_m2"]).fillna(1.0).astype(float)
    factor_m1 = sku_key.map(launch_facts["factor_m1"]).fillna(1.0).astype(float)
    denom_cm = sku_key.map(launch_facts["denom_cm"]).fillna(_denom_default).astype(float)

    # ── v2：逐列缺貨還原（取代式；ND 列不還原；還原上限為該列月均基準）──
    mtd_denom = denom_cm.to_numpy(dtype=float)
    if not mtd_ok:
        warnings.append("MTD 日期無效（或未提供），MTD 月化改以 max(M-1, 月均基準) 計算。")
        mtd_ann_row = np.zeros(len(df))
        r2_cm = np.zeros(len(df))
    else:
        mtd_ann_row = mtd_arr / np.maximum(mtd_denom, 1e-9) * float(month_total_days)
        r2_cm = np.where(
            oos_cm & is_rf & (p90d_arr > 0),
            np.maximum(mtd_ann_row, p90d_arr),
            mtd_ann_row,
        )
    df["FC MTD 月化"] = np.round(mtd_ann_row, 1)
    m2_stock_native = _resolve_column_name("Shop Stock Qty (M-2)", avail)
    m2_stock_col = m2_stock_native
    m2_stock_llm_proxy = False
    if m2_stock_col is None:
        llm_stock_col = _resolve_column_name("Shop Stock Qty (LLM)", avail)
        if llm_stock_col is not None:
            m2_stock_col = llm_stock_col
            m2_stock_llm_proxy = True
    m1_stock_native = _resolve_column_name("Shop Stock Qty (M-1)", avail)
    m1_stock_col = m1_stock_native
    m1_stock_lm_proxy = False
    if m1_stock_col is None:
        lm_stock_col = _resolve_column_name("Shop Stock Qty (LM)", avail)
        if lm_stock_col is not None:
            m1_stock_col = lm_stock_col
            m1_stock_lm_proxy = True
    if m2_stock_col is None:
        warnings.append("來源檔案缺少「Shop Stock Qty (M-2)」欄，M-2 期缺貨還原略過。")
    elif m2_stock_llm_proxy:
        warnings.append("來源檔案缺少「Shop Stock Qty (M-2)」欄，M-2 期缺貨還原改用「Shop Stock Qty (LLM)」代理。")
    if m1_stock_col is None:
        warnings.append("來源檔案缺少「Shop Stock Qty (M-1)」欄，M-1 期缺貨還原略過。")
    elif m1_stock_lm_proxy:
        warnings.append("來源檔案缺少「Shop Stock Qty (M-1)」欄，M-1 期缺貨還原改用「Shop Stock Qty (LM)」代理。")
    if m2_stock_llm_proxy:
        oos_m2 = _safe_numeric(df[m2_stock_col]).to_numpy(dtype=float) <= 0
    else:
        oos_m2 = df.get("Is OOS (M-2)", _safe_default_series(False, idx)).to_numpy(dtype=bool)
    if m1_stock_lm_proxy:
        oos_m1 = _safe_numeric(df[m1_stock_col]).to_numpy(dtype=float) <= 0
    else:
        oos_m1 = df.get("Is OOS (M-1)", _safe_default_series(False, idx)).to_numpy(dtype=bool)
    r2_m2 = m2_arr if m2_stock_col is None else np.where(
        oos_m2 & is_rf & (p90d_arr > 0), np.maximum(m2_arr, p90d_arr), m2_arr)
    r2_m1 = m1_arr if m1_stock_col is None else np.where(
        oos_m1 & is_rf & (p90d_arr > 0), np.maximum(m1_arr, p90d_arr), m1_arr)
    df["FC M-2 還原值"] = np.round(r2_m2, 1)
    df["FC M-1 還原值"] = np.round(r2_m1, 1)
    df["FC MTD 還原月化"] = np.round(r2_cm, 1)
    restored_rows_m2 = int(np.sum(r2_m2 > m2_arr + 1e-9))
    restored_rows_m1 = int(np.sum(r2_m1 > m1_arr + 1e-9))
    restored_rows_cm = int(np.sum(r2_cm > mtd_ann_row + 1e-9))

    # ── Supply Source 正規化＋同 SKU 眾數 ──
    supply_raw = _fc_supply_source_series(df, idx, warnings)

    def _mode_or_first(s: pd.Series) -> str:
        vals = [str(v).strip() for v in s.tolist() if str(v).strip() != ""]
        if not vals:
            return "Warehouse"
        lowered = [v.lower() for v in vals]
        if "warehouse" in lowered:
            return "Warehouse"
        if "vendor" in lowered:
            return "Vendor"
        vc = pd.Series(vals).value_counts()
        return str(vc.index[0])

    sku_supply = sku_key.groupby(sku_key).apply(
        lambda g: _mode_or_first(supply_raw.loc[g.index])
    ) if len(df) else pd.Series(dtype=object)
    n_supply_vals = supply_raw.replace("", np.nan).groupby(sku_key).nunique(dropna=True)
    multi_supply = int(((n_supply_vals.fillna(0)) > 1).sum())
    if multi_supply:
        warnings.append(
            f"偵測到 {multi_supply} 個 SKU 同時混有多個 Supply Source 值，"
            "以眾數（Warehouse 優先）為準計算位置，請檢查來源資料。"
        )
    df["FC Supply Source (SKU)"] = sku_key.map(sku_supply).fillna("Warehouse")
    # 以 Vendor 關鍵字判定（大小寫不敏感；Warehouse 或空白／其他值一律視 Warehouse，
    # 與眾數邏輯一致；「Vendor…」字首值如實檔 Vendor MC 亦視 Vendor 貨）
    _supply_norm = df["FC Supply Source (SKU)"].fillna("").astype(str).str.strip().str.lower()
    is_wh_row = (~_supply_norm.str.startswith("vendor")).to_numpy(dtype=bool)

    # ── 在途／PO 欄（選填欄，缺欄計 0＋warning）──
    pending_col = _fc_resolve_first(_FORECAST_PENDING_RECEIVE_CANDIDATES, avail)
    if pending_col is None:
        warnings.append("來源檔案缺少「Shop Pending Receive (±7 days)」與「Shop Pending Receive」欄，在途貨量一律計 0。")
        pending = np.zeros(len(df))
    else:
        pending = _safe_numeric(df.get(pending_col, _safe_default_series(0, idx))).clip(lower=0).to_numpy(dtype=float)
    if _resolve_column_name(_FORECAST_D001_PO_COL, avail) is None:
        warnings.append(f"來源檔案缺少「{_FORECAST_D001_PO_COL}」欄，D001 在途一律計 0。")
        d001_po = np.zeros(len(df))
    else:
        d001_po = _safe_numeric(df.get(_FORECAST_D001_PO_COL, _safe_default_series(0, idx))).clip(lower=0).to_numpy(dtype=float)
    if _resolve_column_name(_FORECAST_SHOP_PO_COL, avail) is None:
        warnings.append(f"來源檔案缺少「{_FORECAST_SHOP_PO_COL}」欄，店舖 PO 在途一律計 0。")
        shop_po = np.zeros(len(df))
    else:
        shop_po = _safe_numeric(df.get(_FORECAST_SHOP_PO_COL, _safe_default_series(0, idx))).clip(lower=0).to_numpy(dtype=float)
    df["FC 在途 (Pending)"] = pending
    df["FC D001 PO"] = d001_po
    df["FC Shop PO"] = shop_po

    # ── D001 欄值一致性 warning（同 SKU D001 欄應一致；實檔 D001 Net=1 而 Stock=605 另見 §7 風險）──
    d001_net_col = _fc_resolve_first(("D001 Net Stock Qty (CM)", "D001 Stock Qty (CM)"), avail)
    if d001_net_col is None:
        warnings.append("來源檔案缺少「D001 Net Stock Qty (CM)」與「D001 Stock Qty (CM)」，D001 庫存一律視為 0。")
        d001_net = np.zeros(len(df))
        d001_net_is_stock_fallback = False
    else:
        d001_net_all = pd.to_numeric(df.get(d001_net_col, _safe_default_series(0, idx)), errors="coerce")
        d001_net = d001_net_all.fillna(0.0).to_numpy(dtype=float)
        # 負 Net 值保留（沿用 hybrid 的 Net→Stock 回補語意：僅兩者皆空白／非有限值才視為 0）
        d001_net_is_stock_fallback = (d001_net_col == "D001 Stock Qty (CM)")
    df["FC D001 庫存 (Net)"] = d001_net

    if progress_callback:
        progress_callback(0.5, "SKU 級彙總中...")
    work = pd.DataFrame({
        "_sku": sku_key.values,
        "_r2_m2": r2_m2,
        "_r2_m1": r2_m1,
        "_r2_cm": r2_cm,
        "_m2": m2_arr,
        "_m1": m1_arr,
        "_mtd": mtd_arr,
        "_sold_p90d": sold_p90d.to_numpy(dtype=float),
        "_p90d_m": p90d_arr,
        "_stock": cm_stock.to_numpy(dtype=float),
        "_pending": pending,
        "_d001_po": d001_po,
        "_shop_po": shop_po,
        "_d001_net": d001_net,
        "_is_wh": is_wh_row,
        "_is_rf": is_rf,
        "_f_m2": factor_m2.to_numpy(dtype=float),
        "_f_m1": factor_m1.to_numpy(dtype=float),
    })
    grp = work.groupby("_sku", sort=False)
    # P0（個別模式）：單次 agg 取代 9 次同鍵 groupby 掃描；合計／首值／列數口徑不變
    _rf_count_pre = None
    _row_count_pre = None
    try:
        _sku_agg = work.groupby("_sku", sort=False).agg(
            _mtd_sum=("_mtd", "sum"),
            _sold_p90d_sum=("_sold_p90d", "sum"),
            _p90d_m_sum=("_p90d_m", "sum"),
            _r2_m2_sum=("_r2_m2", "sum"),
            _r2_m1_sum=("_r2_m1", "sum"),
            _r2_cm_sum=("_r2_cm", "sum"),
            _f_m2_first=("_f_m2", "first"),
            _f_m1_first=("_f_m1", "first"),
            _is_rf_sum=("_is_rf", "sum"),
            _cnt=("_is_rf", "size"),
        )
        sku_mtd = _sku_agg["_mtd_sum"]
        sku_p90d = _sku_agg["_sold_p90d_sum"]
        B_sku = _sku_agg["_p90d_m_sum"]
        S2 = _sku_agg["_r2_m2_sum"] * _sku_agg["_f_m2_first"]
        S1 = _sku_agg["_r2_m1_sum"] * _sku_agg["_f_m1_first"]
        sku_r2_cm = _sku_agg["_r2_cm_sum"]
        _rf_count_pre = _sku_agg["_is_rf_sum"]
        _row_count_pre = _sku_agg["_cnt"]
    except Exception:
        sku_mtd = grp["_mtd"].sum()
        sku_p90d = grp["_sold_p90d"].sum()
        B_sku = grp["_p90d_m"].sum()
        # v2：SKU 聚合後才做情境；M-2／M-1 先按 SKU 折算，CM 月化已在逐列用有效日數處理
        S2 = grp["_r2_m2"].sum() * grp["_f_m2"].first()
        S1 = grp["_r2_m1"].sum() * grp["_f_m1"].first()
        sku_r2_cm = grp["_r2_cm"].sum()
        _rf_count_pre = None
        _row_count_pre = None
    if mtd_ok:
        sku_mtd_ann = sku_r2_cm
        r_cred = float(mtd_days) / float(month_total_days)
    else:
        sku_mtd_ann = pd.concat([S1, B_sku], axis=1).max(axis=1)
        r_cred = 1.0
    sku_mtd_adj = r_cred * sku_mtd_ann + (1.0 - r_cred) * B_sku

    # ── 情境（SKU 級）：標準＝加權混合；樂觀＝趨勢外推（增速上限＋衰減）──
    s_std = w_p90d * B_sku + w_m1 * S1 + w_mtd * sku_mtd_adj
    m2v = S2.to_numpy(dtype=float)
    m1v = S1.to_numpy(dtype=float)
    mtannv = sku_mtd_ann.to_numpy(dtype=float)
    stdv = s_std.to_numpy(dtype=float)
    adjv = sku_mtd_adj.to_numpy(dtype=float)
    raw_g = np.where(m2v > 0, m1v / np.maximum(m2v, 1e-9) - 1.0, np.where(m1v > 0, trend_cap, 0.0))
    g = np.clip(raw_g, 0.0, trend_cap)
    s_opt = np.maximum(np.maximum(m1v, adjv) * (1.0 + damp * g), stdv)

    # ── 爆升分類（輸入改用還原＋折算後的 S2／S1／MTD 月化）──
    base = np.maximum(m2v, surge_min_base)
    cls = np.full(len(S2), _FC_CLASS_STABLE, dtype=object)
    is_new = (m2v + m1v <= surge_min_base) & (np.maximum(m1v, mtannv) > 0)
    is_accel = (~is_new) & (m1v >= accel_mult * m2v) & (mtannv >= accel_mult * m1v) & (m2v > 0)
    is_surge = (~is_new) & (~is_accel) & (
        np.maximum(m1v, mtannv) / np.maximum(base, 1e-9) >= surge_mult
    ) & (np.maximum(m1v, mtannv) > 0)
    cls[is_new] = _FC_CLASS_NEW
    cls[is_accel] = _FC_CLASS_ACCEL
    cls[is_surge] = _FC_CLASS_SURGE
    use_opt = is_new | is_accel | is_surge

    # ── 月預估（採用）：爆升類取樂觀，其餘取標準；無條件捨去取整 ──
    monthly = np.where(use_opt, s_opt, stdv)
    monthly = np.floor(np.maximum(monthly, 0.0)).astype(int)

    main = pd.DataFrame({
        "SKU": S2.index,
        "FC M-2 合計": S2.to_numpy(dtype=float),
        "FC M-1 合計": S1.to_numpy(dtype=float),
        "FC MTD 合計": sku_mtd.to_numpy(dtype=float),
        "FC MTD 月化": np.round(mtannv, 1),
        "FC P90D 合計": sku_p90d.to_numpy(dtype=float),
        "FC 爆升分類": cls,
        "FC 保守情境": np.floor(np.maximum(B_sku.to_numpy(dtype=float), 0.0)).astype(int),
        "FC 標準情境": np.floor(np.maximum(stdv, 0.0)).astype(int),
        "FC 樂觀情境": np.floor(np.maximum(s_opt, 0.0)).astype(int),
        "FC MTD 調整月化": np.round(adjv, 1),
        "FC 趨勢增速": np.round(g, 2),
    })
    main["FC 月預估 (採用)"] = monthly

    # ── 4 個月 FC 拆分（參考欄，不影響訂量）：本月＝MTD 調整月化；M+1＝月預估 (採用)；
    #    爆升類 M+2／M+3 按趨勢增速 × 衰減係數逐月遞推，平穩類與 M+1 相同（平推）──
    growth_step = 1.0 + damp * g
    fc_m2_ = np.where(use_opt, monthly * growth_step, monthly.astype(float))
    fc_m3_ = np.where(use_opt, monthly * growth_step ** 2, monthly.astype(float))
    main["FC 本月 (M)"] = np.floor(np.maximum(sku_mtd_adj.to_numpy(dtype=float), 0.0)).astype(int)
    main["FC M+1 預估"] = monthly
    main["FC M+2 預估"] = np.floor(np.maximum(fc_m2_, 0.0)).astype(int)
    main["FC M+3 預估"] = np.floor(np.maximum(fc_m3_, 0.0)).astype(int)

    # RAW 三情境欄＝SKU 級數值廣播（v2 情境在 SKU 級計算，逐列供稽核）
    _scen = main.set_index("SKU")[["FC 保守情境", "FC 標準情境", "FC 樂觀情境"]]
    for _c in ("FC 保守情境", "FC 標準情境", "FC 樂觀情境"):
        df[_c] = sku_key.map(_scen[_c]).fillna(0).astype(int)

    # ── 識別欄（Brand／Buyer／SKU Desc 取同 SKU 首個非空值；併最早 Launch Date）──
    # P0：空字串遮成 NaN＋內建 first（Cython 路徑）取代 groupby.apply 自訂函數；
    # 語意等價：first 略過 NaN＝首個非空字串，全空組回填空字串。
    for col in ("Brand", "Buyer", "SKU Desc"):
        if col in df.columns:
            s = df[col].fillna("").astype(str).str.strip()
            s_masked = s.mask(s == "")
            try:
                first_vals = s_masked.groupby(sku_key, sort=False).first()
            except Exception:
                def _first_nonempty(x: pd.Series) -> str:
                    x = x[x != ""]
                    return x.iloc[0] if len(x) else ""

                first_vals = sku_key.groupby(sku_key).apply(lambda g: _first_nonempty(s.loc[g.index]))
            main[col] = (
                first_vals.reindex(main["SKU"]).fillna("").values
            )
        else:
            main[col] = ""
    main["FC Supply Source"] = (
        pd.Series(sku_supply).reindex(pd.Index(main["SKU"])).fillna("Warehouse").values
    )
    if "Launch Date" in df.columns:
        ldt = _parse_launch_dates(df["Launch Date"].fillna("").astype(str).str.strip())
        ldt_s = pd.Series(ldt.values, index=df.index)
        main["FC 最早 Launch Date"] = (
            ldt_s.groupby(sku_key.values).min().reindex(main["SKU"]).dt.strftime("%Y-%m-%d").fillna("").values
        )
    else:
        main["FC 最早 Launch Date"] = ""

    # RP 概況：RF 列數／總列數（全 ND 即無 RF 列）
    # P0（個別模式）：沿用單次 agg 預彙總，避免重複同鍵掃描；口徑不變
    if "_rf_count_pre" in locals() and "_row_count_pre" in locals() and _rf_count_pre is not None and _row_count_pre is not None:
        rf_count = _rf_count_pre
        row_count = _row_count_pre
    else:
        rf_count = grp["_is_rf"].sum()
        row_count = grp["_is_rf"].size()
    main["FC RF 列數"] = rf_count.reindex(main["SKU"]).fillna(0).astype(int).values
    main["FC 總列數"] = row_count.reindex(main["SKU"]).fillna(0).astype(int).values
    main["FC RP 概況"] = np.where(
        main["FC RF 列數"].to_numpy() == 0, "全 ND",
        np.where(main["FC RF 列數"].to_numpy() == main["FC 總列數"].to_numpy(), "全 RF", "RF＋ND"),
    )

    # ── 位置（Warehouse／Vendor 分流；D001 欄取同 SKU 首個非空正值列，實務上同 SKU 各列相同）──
    stock_sum = grp["_stock"].sum().reindex(main["SKU"]).fillna(0).to_numpy(dtype=float)
    pending_sum = grp["_pending"].sum().reindex(main["SKU"]).fillna(0).to_numpy(dtype=float)
    shop_po_sum = grp["_shop_po"].sum().reindex(main["SKU"]).fillna(0).to_numpy(dtype=float)

    def _sku_d001(gidx: pd.Index) -> tuple[float, float]:
        nets = d001_net[gidx.to_numpy()] if hasattr(gidx, "to_numpy") else d001_net[np.asarray(gidx)]
        pos = float(np.sum(nets))
        return pos, float(np.sum(d001_po[np.asarray(gidx)] if hasattr(d001_po, "__len__") else 0.0))

    # D001 欄取同 SKU 首個非空值（實務上同 SKU 各列相同；加總會把同值重複計算）
    d001_net_sum: dict = {}
    d001_po_sum: dict = {}
    if d001_net_col is None:
        for sku_val in work["_sku"].unique():
            d001_net_sum[sku_val] = 0.0
            d001_po_sum[sku_val] = 0.0
    else:
        _d001_frame = pd.DataFrame({
            "_sku": work["_sku"].to_numpy(),
            "_net": d001_net_all.fillna(0.0).to_numpy(dtype=float),
            "_po": d001_po,
            "_seq": np.arange(len(work)),
        }).sort_values("_seq", kind="stable")
        _d001_first = _d001_frame.groupby("_sku", sort=False).first()
        for sku_val in _d001_first.index:
            d001_net_sum[sku_val] = float(_d001_first.loc[sku_val, "_net"])
            d001_po_sum[sku_val] = float(_d001_first.loc[sku_val, "_po"])
    # ── D001 一致性 warning：同 SKU 出現多個不同 D001 Net／PO 值時列示 ──
    if d001_net_col is not None:
        _d001_nuniq = _d001_frame.groupby("_sku", sort=False)[["_net", "_po"]].nunique(dropna=False)
        _d001_multi = _d001_nuniq[(_d001_nuniq["_net"] > 1) | (_d001_nuniq["_po"] > 1)]
        if len(_d001_multi):
            _ex = ", ".join(str(s) for s in list(_d001_multi.index[:5]))
            warnings.append(
                f"偵測到 {len(_d001_multi)} 個 SKU 的 D001 庫存／PO 欄在各列不一致"
                f"（例：{_ex}），已取同 SKU 首個非空值計算，請檢查來源資料。"
            )
    main["_d001_net"] = main["SKU"].map(d001_net_sum).fillna(0.0).to_numpy(dtype=float)
    main["_d001_po"] = main["SKU"].map(d001_po_sum).fillna(0.0).to_numpy(dtype=float)

    is_wh = (
        ~main["FC Supply Source"].fillna("").astype(str).str.strip().str.lower().str.startswith("vendor")
    ).to_numpy(dtype=bool)
    position = np.where(
        is_wh,
        stock_sum + main["_d001_net"].to_numpy(dtype=float) + main["_d001_po"].to_numpy(dtype=float) + pending_sum,
        stock_sum + shop_po_sum + pending_sum,
    )
    main["FC 店舖庫存合計"] = np.floor(stock_sum).astype(int)
    main["FC D001 庫存 (Net)"] = np.floor(main["_d001_net"].to_numpy(dtype=float)).astype(int)
    main["FC D001 PO"] = np.floor(main["_d001_po"].to_numpy(dtype=float)).astype(int)
    main["FC Shop PO 合計"] = np.floor(shop_po_sum).astype(int)
    main["FC 在途合計"] = np.floor(pending_sum).astype(int)
    main["FC 位置合計"] = np.floor(np.maximum(position, 0.0)).astype(int)

    # ── 目標／缺口／上限／建議訂量 ──
    monthly_f = main["FC 月預估 (採用)"].to_numpy(dtype=float)
    pos_f = main["FC 位置合計"].to_numpy(dtype=float)
    # 目標覆蓋：平穩類按目標覆蓋天數；爆升類（新起量／加速中／爆升）按爆升目標覆蓋月數
    cover_days = np.where(use_opt, surge_cover_months * 30.5, float(target_days))
    target = np.ceil(monthly_f / 30.5 * cover_days).astype(int)
    gap = np.maximum(0, target - pos_f).astype(int)
    cap_room = np.maximum(0, np.ceil(monthly_f * cap_months) - pos_f).astype(int)
    order = np.minimum(gap, cap_room)
    all_nd = (main["FC RF 列數"].to_numpy() == 0)
    order = np.where(all_nd, 0, order)
    main["FC 目標覆蓋量"] = target
    main["FC 缺口"] = gap
    main["FC 上限空間"] = cap_room
    main["FC 建議訂量"] = order.astype(int)
    main["FC 建議訂貨對象"] = np.where(is_wh, _FC_TARGET_D001, _FC_TARGET_SHOP)
    main["FC SKU 狀態"] = np.where(
        all_nd, _FC_STATUS_ND,
        np.where((gap > 0) & (gap > cap_room), _FC_STATUS_ORDER_CAPPED,
                 np.where(order > 0, _FC_STATUS_ORDER, _FC_STATUS_ENOUGH)),
    )

    if progress_callback:
        progress_callback(0.65, "計算店舖覆蓋...")

    # ── 風險店數（SKU 級）：可用日數 ≤7 日店數（RF／ND 皆計）、現已缺貨店數、鋪貨店數 ──
    df["_fc_daily"] = np.where(
        (mtd_arr > 0) & (mtd_denom > 0),
        mtd_arr / np.maximum(mtd_denom, 1e-9),
        np.where(sold_p90d.to_numpy(dtype=float) > 0, sold_p90d.to_numpy(dtype=float) / 30.5, np.nan),
    )
    df["_fc_avail"] = cm_stock.to_numpy(dtype=float) + pending
    df["_fc_days"] = df["_fc_avail"] / df["_fc_daily"]
    # 覆蓋風險旗的「目標天數」：爆升類用爆升目標覆蓋月數換算（每 SKU 廣播）
    df["_fc_cover_days"] = (
        sku_key.map(pd.Series(cover_days, index=S2.index)).fillna(float(target_days))
    )
    risk_flag = pd.Series("", index=df.index, dtype=object)
    risk_flag[(cm_stock.to_numpy(dtype=float) <= 0)] = _FC_RISK_OOS
    mask_7 = (risk_flag == "") & df["_fc_days"].notna() & (df["_fc_days"] <= _FC_COVERAGE_RISK_DAYS)
    risk_flag[mask_7] = _FC_RISK_7D
    df["FC 風險旗"] = risk_flag.values
    risk_cnt = risk_flag.isin([_FC_RISK_OOS, _FC_RISK_7D]).groupby(sku_key.values).sum()
    oos_cnt = (cm_stock.to_numpy(dtype=float) <= 0)
    oos_cnt_s = pd.Series(oos_cnt, index=df.index).groupby(sku_key.values).sum()
    site = df["Site Code"].fillna("").astype(str).str.strip() if "Site Code" in df.columns else _safe_default_series("", idx)
    shop_cnt = site[site != ""].groupby(sku_key[site != ""].values).nunique()
    main["FC 缺貨風險店數"] = main["SKU"].map(risk_cnt).fillna(0).astype(int).values
    main["FC 現已缺貨店數"] = main["SKU"].map(oos_cnt_s).fillna(0).astype(int).values
    main["FC 鋪貨店數"] = main["SKU"].map(shop_cnt).fillna(0).astype(int).values

    main = main.sort_values(["Buyer", "Brand", "SKU"], kind="stable").reset_index(drop=True)
    main = main.drop(columns=["_d001_net", "_d001_po"])
    # 契約欄序整理：固定表頭，未經明確要求不得改動
    _fc_text_cols = {
        "Brand", "SKU", "SKU Desc", "Buyer", "FC Supply Source", "FC RP 概況",
        "FC 最早 Launch Date", "FC 爆升分類", "FC 建議訂貨對象", "FC SKU 狀態",
    }
    for _cc in _FORECAST_MAIN_COLUMNS:
        if _cc not in main.columns:
            main[_cc] = "" if _cc in _fc_text_cols else 0
    main = main[_FORECAST_MAIN_COLUMNS]

    # ── 店舖覆蓋表（母體：建議訂量＞0 的 SKU，每店一列）──
    order_skus = set(main.loc[main["FC 建議訂量"] > 0, "SKU"])
    cov = df[df["SKU"].isin(order_skus)].copy()
    if cov.empty:
        coverage = pd.DataFrame(columns=[
            "SKU", "Brand", "Buyer", "Site Code", "Shop Name", "Shop Nature", "RP Type",
            "FC 店舖庫存", "FC 店舖在途", "FC 店舖日銷", "FC 可用日數", "FC 風險旗",
        ])
    else:
        shop_name = cov.get("Shop Name", _safe_default_series("", cov.index))
        try:
            shop_name = shop_name.fillna("").astype(str)
        except Exception:
            shop_name = _safe_default_series("", cov.index).astype(str)
        coverage = pd.DataFrame({
            "SKU": cov["SKU"].values,
            "Brand": cov.get("Brand", _safe_default_series("", cov.index)).fillna("").astype(str).values,
            "Buyer": cov.get("Buyer", _safe_default_series("", cov.index)).fillna("").astype(str).values,
            "Site Code": cov.get("Site Code", _safe_default_series("", cov.index)).fillna("").astype(str).values,
            "Shop Name": shop_name.values,
            "Shop Nature": cov.get("Shop Nature", _safe_default_series("", cov.index)).fillna("").astype(str).values,
            "RP Type": cov.get("RP Type", _safe_default_series("", cov.index)).fillna("").astype(str).values,
            "FC 店舖庫存": np.floor(cov["_fc_avail"].to_numpy(dtype=float) - pending[cov.index.to_numpy()]).astype(int),
            "FC 店舖在途": np.floor(pending[cov.index.to_numpy()]).astype(int),
            "FC 店舖日銷": np.round(cov["_fc_daily"].to_numpy(dtype=float), 2),
            "FC 可用日數": np.round(cov["_fc_days"].to_numpy(dtype=float), 1),
            "FC 風險旗": cov["FC 風險旗"].values,
        })
        # 風險旗第二層：≤目標天數（僅尚未標旗且有日數者；目標天數按 SKU 類別）
        _days = coverage["FC 可用日數"].to_numpy(dtype=float)
        _blank = coverage["FC 風險旗"].to_numpy() == ""
        _cover_days = cov["_fc_cover_days"].to_numpy(dtype=float)
        _le_target = _blank & np.isfinite(_days) & (_days <= _cover_days)
        coverage.loc[_le_target, "FC 風險旗"] = _FC_RISK_TARGET
        coverage.loc[coverage["FC 風險旗"] == "", "FC 風險旗"] = _FC_RISK_OK
        _risk_order = {_FC_RISK_OOS: 0, _FC_RISK_7D: 1, _FC_RISK_TARGET: 2, _FC_RISK_OK: 3}
        coverage["_risk_ord"] = coverage["FC 風險旗"].map(_risk_order).fillna(3).astype(int)
        coverage["_days_ord"] = coverage["FC 可用日數"].fillna(1e12)
        coverage = coverage.sort_values(["SKU", "_risk_ord", "_days_ord"], kind="stable").reset_index(drop=True)
        coverage = coverage.drop(columns=["_risk_ord", "_days_ord"])

    # ── Buyer 彙總（全母體；爆升 SKU 數含新起量／加速中／爆升）──
    buyer_key = main["Buyer"].fillna("").astype(str).str.strip().replace("", "(未填 Buyer)")
    is_order = main["FC 建議訂量"].to_numpy() > 0
    is_surge_cls = main["FC 爆升分類"].isin([_FC_CLASS_NEW, _FC_CLASS_ACCEL, _FC_CLASS_SURGE]).to_numpy()
    is_d001 = main["FC 建議訂貨對象"].to_numpy() == _FC_TARGET_D001
    buyer_df = pd.DataFrame({
        "Buyer": buyer_key.values,
        "_order": np.where(is_order, main["FC 建議訂量"].to_numpy(dtype=float), 0.0),
        "_d001": np.where(is_order & is_d001, main["FC 建議訂量"].to_numpy(dtype=float), 0.0),
        "_shop": np.where(is_order & ~is_d001, main["FC 建議訂量"].to_numpy(dtype=float), 0.0),
        "_surge": is_surge_cls.astype(int),
        "_sku": main["SKU"].values,
    })
    bgrp = buyer_df.groupby("Buyer", sort=False)
    # P0：(s>0).sum() 向量化取代 groupby.apply(lambda)，語意完全等價
    _order_pos = (buyer_df["_order"] > 0).astype(int).groupby(buyer_df["Buyer"], sort=False).sum()
    _order_pos = _order_pos.reindex(bgrp.size().index).fillna(0).astype(int).values
    buyer_out = pd.DataFrame({
        "Buyer": bgrp.size().index,
        "FC SKU 數": bgrp["_sku"].nunique().values,
        "FC 建議訂貨 SKU 數": _order_pos,
        "FC 建議訂量合計 (D001 PO)": bgrp["_d001"].sum().astype(int).values,
        "FC 建議訂量合計 (Shop PO)": bgrp["_shop"].sum().astype(int).values,
        "FC 爆升 SKU 數": bgrp["_surge"].sum().astype(int).values,
    }).sort_values(["FC 建議訂貨 SKU 數", "Buyer"], ascending=[False, True], kind="stable").reset_index(drop=True)

    if progress_callback:
        progress_callback(0.8, "整理輸出...")
    stats = {
        "mtd_days": int(mtd_days) if mtd_days else None,
        "month_total_days": int(month_total_days) if month_total_days else None,
        "target_days": float(target_days),
        "cap_months": float(cap_months),
        "surge_mult": float(surge_mult),
        "surge_min_base": float(surge_min_base),
        "accel_mult": float(accel_mult),
        "damp": float(damp),
        "model": "v2",
        "w_p90d": float(w_p90d),
        "w_m1": float(w_m1),
        "w_mtd": float(w_mtd),
        "trend_cap": float(trend_cap),
        "surge_cover_months": float(surge_cover_months),
        "mtd_date": (
            mtd_date.isoformat() if hasattr(mtd_date, "isoformat")
            else (str(mtd_date) if mtd_date else None)
        ),
        "launch_prorated_skus": int(prorated_skus),
        "restored_rows_m2": int(restored_rows_m2),
        "restored_rows_m1": int(restored_rows_m1),
        "restored_rows_cm": int(restored_rows_cm),
        "m2m1_stock_missing": bool(m2_stock_col is None or m1_stock_native is None),
        "m2_stock_missing": bool(m2_stock_col is None),
        "m1_stock_missing": bool(m1_stock_native is None),
        "m1_lm_proxy": bool(m1_stock_lm_proxy),
        "m2_llm_proxy": bool(m2_stock_llm_proxy),
        "pending_col": pending_col,
        "d001_net_col": d001_net_col,
        "d001_net_is_stock_fallback": bool(d001_net_is_stock_fallback),
        "multi_supply_skus": int(multi_supply),
        "total_skus": int(len(main)),
        "surge_skus": int(is_surge_cls.sum()),
        "order_skus": int(is_order.sum()),
        "total_order": int(main["FC 建議訂量"].sum()),
        "total_order_d001": int(main.loc[is_d001, "FC 建議訂量"].sum()),
        "total_order_shop": int(main.loc[~is_d001, "FC 建議訂量"].sum()),
        "capped_skus": int((main["FC SKU 狀態"] == _FC_STATUS_ORDER_CAPPED).sum()),
        "nd_skus": int((main["FC SKU 狀態"] == _FC_STATUS_ND).sum()),
        "coverage_rows": int(len(coverage)),
        "raw_rows": int(len(df)),
    }

    raw_out = df.drop(columns=["_fc_daily", "_fc_avail", "_fc_days"], errors="ignore")
    if progress_callback:
        progress_callback(1.0, "完成!")
    return main, coverage, buyer_out, raw_out, stats, warnings



# Forecast 主表契約欄序（一 SKU 一列；固定表頭屬資料契約，未經明確要求不得改動）
_FORECAST_MAIN_COLUMNS = [
    "Brand", "SKU", "SKU Desc", "Buyer", "FC Supply Source", "FC RP 概況", "FC 最早 Launch Date",
    "FC M-2 合計", "FC M-1 合計", "FC MTD 合計", "FC MTD 月化", "FC P90D 合計",
    "FC 爆升分類", "FC 保守情境", "FC 標準情境", "FC 樂觀情境",
    "FC MTD 調整月化", "FC 趨勢增速", "FC 月預估 (採用)",
    "FC 本月 (M)", "FC M+1 預估", "FC M+2 預估", "FC M+3 預估",
    "FC 店舖庫存合計", "FC D001 庫存 (Net)", "FC D001 PO", "FC Shop PO 合計", "FC 在途合計",
    "FC 位置合計", "FC 目標覆蓋量", "FC 缺口", "FC 上限空間",
    "FC 建議訂量", "FC 建議訂貨對象", "FC SKU 狀態",
    "FC 缺貨風險店數", "FC 現已缺貨店數", "FC 鋪貨店數",
]



_FORECAST_COVERAGE_COLUMNS = [
    "SKU", "Brand", "Buyer", "Site Code", "Shop Name", "Shop Nature", "RP Type",
    "FC 店舖庫存", "FC 店舖在途", "FC 店舖日銷", "FC 可用日數", "FC 風險旗",
]



_FORECAST_BUYER_COLUMNS = [
    "Buyer", "FC SKU 數", "FC 建議訂貨 SKU 數",
    "FC 建議訂量合計 (D001 PO)", "FC 建議訂量合計 (Shop PO)", "FC 爆升 SKU 數",
]



# Forecast RAW 顯示欄序分組（識別→訊號→位置→SKU 目標→其他原始欄）
_FORECAST_RAW_DISPLAY_GROUPS = [
    ["Buyer", "Brand", "SKU", "SKU Desc", "Site Code", "Shop Name", "OM", "Shop Nature"],
    ["Supply Source", "FC Supply Source (SKU)", "RP Type", "FC RP Type (正規)", "Launch Date",
     "Shop Stock Status"],
    ["Shop Stock Qty (CM)", "Sold Qty (P90D)", "Shop Monthly Sold Qty (P90D)",
     "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
     "FC M-2 還原值", "FC M-1 還原值", "FC MTD 還原月化",
     "FC MTD 月化", "FC 保守情境", "FC 標準情境", "FC 樂觀情境", "FC 缺貨還原列",
     "FC D001 庫存 (Net)", "FC D001 PO", "FC Shop PO", "FC 在途 (Pending)", "FC 風險旗"],
]



def _forecast_raw_display_columns(raw: pd.DataFrame) -> list[str]:
    """Forecast RAW 顯示欄序：不顯示「Item_Shop」與所有含「Value」的金額欄，
    其餘欄位按 _FORECAST_RAW_DISPLAY_GROUPS 分組排列，未歸組的原始欄位按原序置於最後。"""
    cols = [str(c) for c in raw.columns] + ["Shop Name", "SKU Desc"]
    seen: set[str] = set()
    kept: list[str] = []
    for c in cols:
        if c in seen or c == "Item_Shop" or "value" in c.lower():
            continue
        seen.add(c)
        kept.append(c)
    kept_set = set(kept)
    ordered: list[str] = []
    for group in _FORECAST_RAW_DISPLAY_GROUPS:
        for c in group:
            if c in kept_set:
                ordered.append(c)
                kept_set.discard(c)
    ordered += [c for c in kept if c in kept_set]
    return ordered



def _write_forecast_guide_sheet(ws, stats):
    """Forecast「說明」工作表（工作簿首位）：列出本次執行參數、講解模式核心邏輯，
    並逐頁導讀各工作表應留意的部分。內容為靜態說明＋stats 參數照錄，不含計算欄位。"""
    st_ = stats or {}

    def _dw(text):
        return sum(2 if ord(ch) > 0x2E7F else 1 for ch in str(text))

    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 56
    ws.column_dimensions["C"].width = 66
    title_font = Font(size=15, bold=True, color="1A1A2E")
    sub_font = Font(size=9, color="6B7280", italic=True)
    section_fill = PatternFill("solid", start_color="0077A3")
    section_font = Font(size=11, bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="E4EEF5")
    label_font = Font(size=10, bold=True)
    body_font = Font(size=10)
    wrap = Alignment(vertical="top", wrap_text=True)
    r = 1

    def _merged(text, font, height=None, fill=None):
        nonlocal r
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        c = ws.cell(row=r, column=1, value=text)
        c.font = font
        c.alignment = Alignment(vertical="center", wrap_text=True)
        if fill:
            c.fill = fill
        ws.row_dimensions[r].height = height or 18
        r += 1

    def _sec(text):
        _merged(text, section_font, height=20, fill=section_fill)

    def _kv(label, value):
        nonlocal r
        ws.cell(row=r, column=1, value=label).font = label_font
        ws.cell(row=r, column=1).alignment = wrap
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c = ws.cell(row=r, column=2, value=value)
        c.font = body_font
        c.alignment = wrap
        lines = max(1, -(-_dw(value) // 110))
        ws.row_dimensions[r].height = 14 * lines + 6
        r += 1

    def _bullet(text):
        _kv("•", text)

    def _row3(a, b, c):
        nonlocal r
        for col, val, f in ((1, a, label_font), (2, b, body_font), (3, c, body_font)):
            cell = ws.cell(row=r, column=col, value=val)
            cell.font = f
            cell.alignment = wrap
        lines = max(1, -(-_dw(a) // 25), -(-_dw(b) // 54), -(-_dw(c) // 64))
        ws.row_dimensions[r].height = 14 * lines + 6
        r += 1

    _merged("Forecast Review 說明與導讀", title_font, height=26)
    _merged("本頁講解此報告的 Forecast 邏輯與本次執行參數，並逐頁說明應留意的部分；各項判定以主表逐 SKU 欄位為準。",
            sub_font, height=16)

    # ── ① 本次執行參數 ──
    _sec("① 本次執行參數（供核對）")
    _kv("模型版本", "v2（SKU 級聚合：標準＝加權混合、樂觀＝趨勢外推）")
    _kv("MTD 日期", f'{st_.get("mtd_days", "—")} 日（當月共 {st_.get("month_total_days", "—")} 日；MTD 月化＝MTD ÷ 有效日數 × 當月日數，開賣未滿月按開賣後日數）')
    _kv("目標覆蓋天數", f'{st_.get("target_days", 30):g} 日（平穩類目標覆蓋量＝ceil(月預估 ÷ 30.5 × 天數)）')
    _kv("爆升目標覆蓋月數", f'{st_.get("surge_cover_months", 2.0):g} 個月（新起量／加速中／爆升：目標覆蓋量＝ceil(月預估 × 月數)，仍受上限月數封頂）')
    _kv("上限月數", f'{st_.get("cap_months", 2.0):g} 個月（上限空間＝ceil(月預估 × 月數) − 位置；建議訂量＝min(缺口, 上限空間)）')
    _kv("模型權重", f'月均基準 {st_.get("w_p90d", 0.4):g}／M-1 {st_.get("w_m1", 0.3):g}／MTD 調整月化 {st_.get("w_mtd", 0.3):g}（自動正規化；標準情境＝三項加權混合）')
    _kv("趨勢增速上限", f'{st_.get("trend_cap", 1.0):g} 倍（樂觀情境＝max(M-1 折算後, MTD 調整月化) × (1 ＋ 衰減係數 × 增速)）')
    _kv("爆升倍數", f'{st_.get("surge_mult", 1.5):g} 倍（max(M-1, MTD 月化) ÷ max(M-2, 新起量基數) 達此倍數判爆升；輸入為還原＋折算後數值）')
    _kv("新起量基數", f'{st_.get("surge_min_base", 10):g} 件（M-2＋M-1 合計不超過此數且近期有量 → 新起量）')
    _kv("加速倍數", f'{st_.get("accel_mult", 1.15):g} 倍（M-1 ≥ 此倍數 × M-2 且 MTD 月化 ≥ 此倍數 × M-1 → 加速中）')
    _kv("衰減係數", f'{st_.get("damp", 0.8):g}（趨勢外推的衰減；樂觀不低於標準）')
    _kv("缺貨還原列數", f'M-2 {st_.get("restored_rows_m2", 0):g} 列／M-1 {st_.get("restored_rows_m1", 0):g} 列／CM {st_.get("restored_rows_cm", 0):g} 列（缺貨列還原至該店月均；ND 不還原）'
        + ("；M-2 期以 Shop Stock Qty (LLM) 代理" if st_.get("m2_llm_proxy")
           else ("；M-2 期缺 Shop Stock Qty (M-2) 欄不還原" if st_.get("m2_stock_missing") else ""))
        + ("；M-1 期以 Shop Stock Qty (LM) 代理" if st_.get("m1_lm_proxy")
           else ("；M-1 期缺 Shop Stock Qty (M-1) 欄不還原" if st_.get("m1_stock_missing") else "")))
    _kv("開賣折算 SKU 數", f'{st_.get("launch_prorated_skus", 0):g} 個（M-2／M-1 按整月折算；本月按開賣後有效日數，最少 7 日）')
    _kv("在途欄", f'{st_.get("pending_col") or "—"}（Warehouse／Vendor 貨皆計入位置）')
    _kv("D001 庫存欄", f'{st_.get("d001_net_col") or "—"}'
        + ("（來源缺 Net 欄，已退用 D001 Stock Qty (CM)）" if st_.get("d001_net_is_stock_fallback") else ""))

    # ── ② 核心邏輯 ──
    _sec("② 核心邏輯（這份報告怎樣算出來）")
    _bullet("逐列缺貨還原（取代式，非疊加）：M-2／M-1／CM 期缺貨（該期店舖庫存 0）且 RF 的列，"
            "取該期實績與該店月均基準（Shop Monthly Sold Qty (P90D)）孰高；還原上限為該店月均。"
            "ND 列不還原；來源缺 Shop Stock Qty (M-2) 但有 Shop Stock Qty (LLM) 時以 LLM 代理 M-2、缺 (M-1) 但有 Shop Stock Qty (LM) 時以 LM 代理，兩期欄皆缺時該期略過，均列 warning。")
    _bullet("開賣日折算（SKU 級）：最早 Launch Date 落在 M-2／M-1 時，按期內有效日數折算為整月規模"
            "（最少按 7 日計）；落在本月時 MTD 月化以開賣後有效日數為分母。")
    _bullet("標準情境（SKU 級聚合後計算）＝月均基準 × 權重 ＋ M-1 折算後 × 權重 ＋ MTD 調整月化 × 權重；"
            "MTD 調整月化＝按已過日數把 MTD 月化與月均基準混合（月初 MTD 佔比較低）。保守情境＝月均基準（只作參考）。")
    _bullet("樂觀情境＝max(M-1 折算後, MTD 調整月化) × (1 ＋ 衰減係數 × 趨勢增速)，趨勢增速設上限且不低於標準。"
            "爆升四類（新起量／加速中／爆升）用樂觀情境、其餘平穩用標準情境；分類輸入同為還原＋折算後數值。")
    _bullet("4 個月 FC 拆分（參考欄）：本月 (M)＝MTD 調整月化（MTD 日期所在月份）；M+1＝月預估 (採用)；"
            "爆升類 M+2／M+3 按趨勢增速 × 衰減係數逐月遞推，平穩類與 M+1 相同。此 4 欄僅供參考，建議訂量仍按月預估 (採用) 計算。")
    _bullet("位置分流：Warehouse 貨＝店舖庫存合計＋D001 庫存 (Net)＋D001 PO＋在途；"
            "Vendor 貨＝店舖庫存合計＋Shop PO 合計＋在途。"
            "建議訂貨：Warehouse（D001）貨由 D001 向 Vendor 下單補貨（D001 PO）；Vendor 貨由店舖向 Vendor 下單（Shop PO）。")
    _bullet("缺口與上限：目標覆蓋量——平穩類＝ceil(月預估 ÷ 30.5 × 目標覆蓋天數)，爆升類（新起量／加速中／爆升）＝ceil(月預估 × 爆升目標覆蓋月數)；"
            "缺口＝max(0, 目標覆蓋量 − 位置)；上限空間＝max(0, ceil(月預估 × 上限月數) − 位置)；"
            "建議訂量＝min(缺口, 上限空間)；缺口大於上限空間時狀態標「建議訂貨（已封頂）」。全 ND 的 SKU 建議訂量為 0。")
    _bullet("店舖覆蓋：店舖日銷＝MTD ÷ MTD 日數（無 MTD 銷售時退回 P90D ÷ 30.5）；"
            "可用日數＝(店舖庫存＋該店在途) ÷ 店舖日銷；風險旗：已缺貨（庫存 0）／≤7日／≤目標天數（平穩類用目標覆蓋天數、爆升類用爆升目標覆蓋月數）／足夠。")

    # ── ③ 工作表導讀 ──
    _sec("③ 工作表導讀（先看哪一頁、應留意什麼）")
    for col, val in ((1, "工作表"), (2, "內容"), (3, "建議留意")):
        c = ws.cell(row=r, column=col, value=val)
        c.font = label_font
        c.fill = head_fill
        c.alignment = Alignment(vertical="center")
    ws.row_dimensions[r].height = 16
    r += 1
    _row3("說明（本頁）", "Forecast 邏輯、本次參數與導讀", "執行任何行動前，先核對上方 ① 的參數是否符合預期")
    _row3("Forecast 主表", "SKU 級建議訂貨主表（一 SKU 一列）",
          "留意「爆升分類」「FC 月預估 (採用)」「FC 本月 (M)」「FC M+1／M+2／M+3 預估」「FC 建議訂量」「FC 建議訂貨對象」「FC SKU 狀態」；已封頂列代表需求超過上限")
    _row3("店舖覆蓋", "有訂量 SKU 的每店可用日數與風險旗",
          "先跟進「已缺貨」與「≤7日」列；訂貨到店前的空窗期以此表安排")
    _row3("按 Buyer 彙總", "各 Buyer 的建議訂量合計（分 D001 PO／Shop PO）",
          "採購按人落單；爆升 SKU 數看需求熱度分佈")
    _row3("RAW", "全部店舖×SKU 逐列明細與訊號欄",
          "留意「FC M-2 還原值」「FC M-1 還原值」「FC MTD 還原月化」；三情境欄為 SKU 級數值同值廣播，供稽核")
    _row3("Forecast 摘要", "指標卡＋按 Buyer 表", "管理層概覽；建議總訂量看 D001 PO／Shop PO 分項")

    # ── ④ 實用提示 ──
    _sec("④ 實用提示")
    _bullet("「FC 建議訂量」為 SKU 級總量；Warehouse(D001) 貨向 Vendor 下單（替 D001 補貨，D001 PO），"
            "Vendor 貨由店舖向供應商下單（Shop PO）；店舖分配由 Buyer 按店舖覆蓋表安排。")
    _bullet("月初 MTD 日數少時，MTD 調整月化已按已過日數與月均基準混合，樂觀情境再乘趨勢衰減係數；仍建議對照保守情境欄。")
    _bullet("「建議訂貨（已封頂）」代表需求超過上限月數，訂完仍可能不足，下期複查。")



def _forecast_main_columns(main: pd.DataFrame) -> list[str]:
    """Forecast 主表契約欄序：固定表頭，未經明確要求不得改動；缺欄時按契約補空欄。"""
    out = main.copy()
    _fc_text_cols = {
        "Brand", "SKU", "SKU Desc", "Buyer", "FC Supply Source", "FC RP 概況",
        "FC 最早 Launch Date", "FC 爆升分類", "FC 建議訂貨對象", "FC SKU 狀態",
    }
    for c in _FORECAST_MAIN_COLUMNS:
        if c not in out.columns:
            out[c] = "" if c in _fc_text_cols else 0
    return out[_FORECAST_MAIN_COLUMNS]



def _write_forecast_sheet(ws, df: pd.DataFrame, columns: list[str], title: str,
                          highlight_order: bool = False, thin_border=None,
                          progress_callback=None, progress_start=0.0, progress_end=0.1):
    """Forecast 通用表寫入：表頭＋凍結＋自動篩選；建議訂量欄醒目底色加強顯示。"""
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
    tc = ws.cell(row=1, column=1, value=title)
    tc.font = Font(bold=True, size=13, color="1E40AF")
    tc.alignment = Alignment(horizontal="left", vertical="center")
    for col_idx, h in enumerate(columns, 1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill(start_color="0077A3", end_color="0077A3", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if thin_border is not None:
            cell.border = thin_border
    order_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    order_col_idx = columns.index("FC 建議訂量") + 1 if "FC 建議訂量" in columns else None
    order_bold = Font(bold=True)
    total = len(df)
    # P0：itertuples＋預先共用字型取代 iterrows＋逐格新建 Font；值與格式語意不變
    for i, _vals in enumerate(_iter_header_rows(df, columns)):
        r = i + 3
        for col_idx, h in enumerate(columns, 1):
            v = _vals[col_idx - 1]
            cell = ws.cell(row=r, column=col_idx, value=None if pd.isna(v) else v)
            if thin_border is not None:
                cell.border = thin_border
            if highlight_order and col_idx == order_col_idx:
                cell.fill = order_fill
                cell.font = order_bold
        if progress_callback and total and (i % 1000 == 0 or i == total - 1):
            p = progress_start + (progress_end - progress_start) * ((i + 1) / total)
            progress_callback(p, f"寫入 {title} ({i+1}/{total})...")
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = ws.dimensions
    _auto_width(ws)



def generate_excel_forecast(
    main: pd.DataFrame | None = None,
    coverage: pd.DataFrame | None = None,
    buyer: pd.DataFrame | None = None,
    raw: pd.DataFrame | None = None,
    stats: dict | None = None,
    progress_callback=None,
    include_raw: bool = True,
) -> BytesIO:
    """Forecast Review 輸出（6 工作表，順序固定）：
    說明（首頁導讀：本次執行參數列出＋模式核心邏輯＋各工作表應留意的部分）/
    Forecast 主表（SKU 級建議訂貨主表，建議訂量欄醒目底色）/
    店舖覆蓋（有訂量 SKU 的每店可用日數與風險旗）/
    按 Buyer 彙總（建議訂量合計分 D001 PO／Shop PO 兩欄）/
    RAW（重排欄序提升可讀性，不顯示 Item_Shop 與 Value 金額欄；含逐列訊號欄與廣播 SKU 目標欄）/
    Forecast 摘要（指標卡＋按 Buyer 表）。
    開啟 Excel 時預設顯示「說明」頁。"""
    wb = Workbook()
    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    desc_font = Font(size=9, color="6B7280", italic=True)
    st_ = stats or {}

    main_df = pd.DataFrame() if main is None else main.copy()
    if not main_df.empty:
        main_df = _forecast_main_columns(main_df)
    coverage_df = pd.DataFrame() if coverage is None else coverage.copy()
    if not coverage_df.empty:
        coverage_df = coverage_df[[c for c in _FORECAST_COVERAGE_COLUMNS if c in coverage_df.columns]]
    buyer_df = pd.DataFrame() if buyer is None else buyer.copy()
    if not buyer_df.empty:
        buyer_df = buyer_df[[c for c in _FORECAST_BUYER_COLUMNS if c in buyer_df.columns]]

    ws_guide = wb.active
    ws_guide.title = "說明"
    _write_forecast_guide_sheet(ws_guide, st_)

    if progress_callback:
        progress_callback(0.15, "建立 Forecast 主表...")
    ws_main = wb.create_sheet(title="Forecast 主表")
    if main_df.empty:
        ws_main.cell(row=2, column=1, value="無符合資料")
    else:
        _write_forecast_sheet(ws_main, main_df, _FORECAST_MAIN_COLUMNS,
                              "Forecast 主表 — SKU 級建議訂貨（一 SKU 一列，按 Buyer／Brand／SKU 排序）",
                              highlight_order=True, thin_border=thin_border,
                              progress_callback=progress_callback, progress_start=0.15, progress_end=0.45)

    if progress_callback:
        progress_callback(0.5, "建立店舖覆蓋...")
    ws_cov = wb.create_sheet(title="店舖覆蓋")
    if coverage_df.empty:
        ws_cov.cell(row=2, column=1, value="無符合資料（無建議訂量＞0 的 SKU）")
    else:
        _write_forecast_sheet(ws_cov, coverage_df, _FORECAST_COVERAGE_COLUMNS,
                              "店舖覆蓋 — 有訂量 SKU 的每店可用日數與風險旗（按 SKU＋風險＋日數排序）",
                              thin_border=thin_border,
                              progress_callback=progress_callback, progress_start=0.5, progress_end=0.6)

    if progress_callback:
        progress_callback(0.62, "建立按 Buyer 彙總...")
    ws_buyer = wb.create_sheet(title="按 Buyer 彙總")
    if buyer_df.empty:
        ws_buyer.cell(row=2, column=1, value="無符合資料")
    else:
        _write_forecast_sheet(ws_buyer, buyer_df, _FORECAST_BUYER_COLUMNS,
                              "按 Buyer 彙總 — 建議訂量合計（分 D001 PO／Shop PO）",
                              thin_border=thin_border,
                              progress_callback=progress_callback, progress_start=0.62, progress_end=0.68)

    if progress_callback:
        progress_callback(0.7, "建立 RAW...")
    ws_raw = wb.create_sheet(title="RAW")
    (_write_raw_sheet(ws_raw, pd.DataFrame() if raw is None else raw,
                     progress_callback, progress_start=0.7, progress_end=0.85,
                     columns=_forecast_raw_display_columns(pd.DataFrame() if raw is None else raw)) if include_raw else _write_raw_placeholder(ws_raw))

    if progress_callback:
        progress_callback(0.87, "建立 Forecast 摘要...")
    ws_dash = wb.create_sheet(title="Forecast 摘要")
    row1_metrics = [
        ("SKU 數", int(st_.get("total_skus", len(main_df))), "本次計算的 SKU 總數", "3B82F6"),
        ("爆升 SKU 數", int(st_.get("surge_skus", 0)), "新起量／加速中／爆升合計", "F59E0B"),
        ("建議總訂量", int(st_.get("total_order", 0)), "全部 SKU 建議訂量合計（件）", "EF4444"),
        ("其中 D001 PO", int(st_.get("total_order_d001", 0)), "Warehouse(D001) 貨向 Vendor 下單（件）", "8B5CF6"),
    ]
    row2_metrics = [
        ("其中 Shop PO", int(st_.get("total_order_shop", 0)), "Vendor 貨向供應商下單（件）", "06B6D4"),
        ("已封頂 SKU 數", int(st_.get("capped_skus", 0)), "缺口超過上限空間的 SKU", "F97316"),
        ("ND 不訂 SKU 數", int(st_.get("nd_skus", 0)), "全 ND 列的 SKU（建議訂量 0）", "64748B"),
        ("店舖覆蓋列數", int(st_.get("coverage_rows", len(coverage_df))), "有訂量 SKU 的店舖列數", "10B981"),
    ]
    _write_metric_cards(
        ws_dash, "Forecast 摘要 (Forecast Dashboard)",
        row1_metrics, row2_metrics, thin_border, desc_font, progress_callback,
    )
    if not buyer_df.empty:
        _br = 11
        ws_dash.merge_cells(start_row=_br, start_column=1, end_row=_br, end_column=4)
        tc = ws_dash.cell(row=_br, column=1, value="按 Buyer 建議訂量統計")
        tc.font = Font(bold=True, size=12, color="1E40AF")
        tc.alignment = Alignment(horizontal="left", vertical="center")
        for col_idx, h in enumerate(("Buyer", "建議訂貨 SKU 數", "D001 PO (件)", "Shop PO (件)"), 1):
            cell = ws_dash.cell(row=_br + 1, column=col_idx, value=h)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="0EA5E9", end_color="0EA5E9", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        rr = _br + 2
        _buyer_headers = ("Buyer", "FC 建議訂貨 SKU 數",
                          "FC 建議訂量合計 (D001 PO)", "FC 建議訂量合計 (Shop PO)")
        _buyer_center = Alignment(horizontal="center", vertical="center")
        for _bvals in _iter_header_rows(buyer_df, list(_buyer_headers)):
            for col_idx, h in enumerate(_buyer_headers, 1):
                cell = ws_dash.cell(row=rr, column=col_idx, value=_bvals[col_idx - 1])
                cell.border = thin_border
                if col_idx > 1:
                    cell.number_format = "#,##0"
                    cell.alignment = _buyer_center
            rr += 1
        _auto_width(ws_dash)

    wb.active = 0

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf

