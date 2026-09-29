# -*- coding: utf-8 -*-
"""Safety Stock 月度變化分析：追蹤店舖每月 Safety Stock 原始值變化對生意的影響。

只引用 modes.common，不引用其他模式模組，避免循環引用。

輸入：標準 By SKUxShop 欄位＋尾部 ``Month``／``Year``（大小寫皆接受，
此模式恒為必要欄位，缺任一即停止）。主鍵為 ``Site Code × SKU × YEAR-MONTH``，
兼容多店舖檔案；單一店舖（如 HA19）亦可執行。

分析欄固定五項（Safety 只看原始值增減，不作建議值計算）：
- ``Adj. Safety Stock`` 變化（增加／減少／不變）
- ``RP Type`` 轉變（RF→RF、ND→ND、RF→ND、ND→RF）
- ``Sold Qty (CM - MTD)`` 環比變化（最新月份按數據截止日月化後比較）
- ``Net Sales (CM - MTD)`` 環比變化（同上，兩欄同時月化）
- ``Shop Stock Qty (CM)`` 環比變化（時點快照，不月化）

月化規則：除上載當月外，其餘月份的 ``CM - MTD`` 已是足月實績；
上載當月只截至數據截止日（預設昨日）。月化係數＝當月總日數 ÷ 截止日日號，
僅套用於最新 ``YEAR-MONTH`` 的 Sold 與 Net Sales，舊月份係數為 1。

輸出八個工作表：
1. 說明 — 本次執行參數、核心邏輯與工作表導讀（預設開啟頁）
2. SKU月度變化 — 每組 ``Site Code × SKU`` 相鄰月份環比事件（上月缺資料者標示新增組合）
3. By Adj. MCH1 — 各 MCH1 每月合計（月化基準）＋全部合計
4. By Adj. SST — 各 SST 每月合計（月化基準）＋全部合計
5. By MCH1×SST — MCH1 與 SST 組合每月合計
6. By Shop — 每間店舖每月合計＋環比變化
7. 月度摘要 — 指標卡＋每月合計表
8. RAW — 原始明細＋YEAR-MONTH 與月化估算欄
"""

import calendar
import datetime
import gc
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from modes.common import (
    _ADJ_SAFETY_SOURCE_COL,
    _auto_width,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _display_safety_header,
    _normalize_ref_date,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _write_guide_sheet,
    _write_metric_cards,
    _write_raw_placeholder,
    _write_raw_sheet,
    validate_input,
)

__all__ = [
    "transform_safety_monthly",
    "generate_excel_safety_monthly",
    "safety_monthly_dashboard_metrics",
    "_SAFM_CHANGE_COLUMNS",
    "_SAFM_SHOP_COLUMNS",
    "_SAFM_YM_COLUMN",
    "_SAFM_SOLD_ADJ_COLUMN",
    "_SAFM_NET_ADJ_COLUMN",
    "_SAFM_REQUIRED",
    "_SAFM_PIVOT_METRICS",
    "_SAFM_RP_ORDER",
    "_SAFM_RP_ENDPOINT_COLUMNS",
    "_build_safm_rp_summary",
    "_build_safm_rp_endpoints",
]

# 樞紐四項指標（顯示全稱，不用簡稱）：
# Sold Qty＝Sold Qty (CM - MTD) 月化合計；Net Sales＝Net Sales (CM - MTD) 月化合計；
# Net Stock Qty＝店舖存貨 Shop Stock Qty (CM) 合計（時點快照）；Safety Stock Qty＝Adj. Safety Stock 合計。
_SAFM_PIVOT_METRICS = ("Sold Qty", "Net Sales", "Net Stock Qty", "Safety Stock Qty")

# RP Type 轉變固定顯示順序（RF→ND 與 ND→RF 為重點觀察對象，置於後方兩列）
_SAFM_RP_ORDER = ("RF→RF", "ND→ND", "RF→ND", "ND→RF")

_SAFM_METRIC_COLORS = {
    "Sold Qty": "4472C4",
    "Net Sales": "70AD47",
    "Net Stock Qty": "ED7D31",
    "Safety Stock Qty": "A855F7",
}

# 新契約欄名（未經明確要求不得改動）
_SAFM_YM_COLUMN = "YEAR-MONTH"
_SAFM_SOLD_ADJ_COLUMN = "Sold Qty (CM-月化估算)"
_SAFM_NET_ADJ_COLUMN = "Net Sales (CM-月化估算)"

_SAFM_REQUIRED = [
    "Site Code",
    "Adj. Safety Stock",
    "RP Type",
    "Sold Qty (CM - MTD)",
    "Net Sales (CM - MTD)",
    "Shop Stock Qty (CM)",
]

_SAFM_CHANGE_COLUMNS = [
    "Site Code", "Shop Name", "SKU", "SKU Desc", "Brand", "Buyer",
    "Adj. MCH1", "Adj. SST",
    "上月", "本月", "是否相鄰曆月",
    "上月 Adj. Safety Stock", "本月 Adj. Safety Stock",
    "Safety變化量", "Safety變化分類",
    "上月 RP Type", "本月 RP Type", "RP轉變",
    "上月 Sold Qty (CM-MTD)", "本月 Sold Qty (CM-MTD)",
    "上月 Sold月化", "本月 Sold月化",
    "Sold變化量(月化基準)", "Sold變化率",
    "上月 Net Sales (CM-MTD)", "本月 Net Sales (CM-MTD)",
    "上月 Net月化", "本月 Net月化",
    "Net變化量(月化基準)", "Net變化率",
    "上月 Shop Stock Qty (CM)", "本月 Shop Stock Qty (CM)",
    "Stock變化量", "Stock變化率",
]

_SAFM_SHOP_COLUMNS = [
    "Site Code", "Shop Name", "YEAR-MONTH",
    "SKU數", "Safety合計",
    "Sold原始合計", "Sold月化合計",
    "Net原始合計", "Net月化合計",
    "Stock合計",
    "Sold環比變化", "Net環比變化", "Safety環比變化",
]


def _find_month_year_columns(columns) -> tuple[str | None, str | None]:
    """尾部 MONTH／YEAR 欄位尋找（大小寫皆接受，前後空格忽略）。"""
    month_col = None
    year_col = None
    for c in columns:
        if not isinstance(c, str):
            continue
        s = c.strip().lower()
        if s == "month" and month_col is None:
            month_col = c
        elif s == "year" and year_col is None:
            year_col = c
    return month_col, year_col


def _ym_sort_key(year: int, month: int) -> int:
    return int(year) * 100 + int(month)


def _ym_label(year: int, month: int) -> str:
    return f"{int(year)}-{int(month):02d}"


def _is_calendar_adjacent(prev_key: int, curr_key: int) -> bool:
    py, pm = prev_key // 100, prev_key % 100
    cy, cm = curr_key // 100, curr_key % 100
    if cm == 1:
        return py == cy - 1 and pm == 12
    return py == cy and pm == cm - 1


def _change_rate(curr: float, prev: float):
    if prev == 0:
        return None
    try:
        return round((curr - prev) / abs(prev), 4)
    except (TypeError, ValueError, ZeroDivisionError):
        return None


def transform_safety_monthly(
    df_input: pd.DataFrame,
    ref_date=None,
    progress_callback=None,
    shop_natures: list[str] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, dict, list[str]]:
    """Safety Stock 月度變化分析核心轉換。

    Returns:
        change_df: SKU 月度變化主表（每組相鄰月份一列環比事件）
        mch1_pivot: By Adj. MCH1 每月合計
        sst_pivot: By Adj. SST 每月合計
        mch1_sst_pivot: By MCH1×SST 每月合計
        shop_pivot: By Shop 每月合計＋環比
        raw_df: RAW 明細（含 YEAR-MONTH 與月化估算欄）
        stats: 統計與本次執行參數
        warnings: 警告訊息列表
    """
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    missing_global = validate_input(df_input)
    if missing_global:
        raise ValueError(f"缺少必要欄位: {', '.join(missing_global)}")

    month_col, year_col = _find_month_year_columns(df_input.columns)
    missing_my = []
    if month_col is None:
        missing_my.append("MONTH")
    if year_col is None:
        missing_my.append("YEAR")
    if missing_my:
        raise ValueError(
            f"此模式必須包含 MONTH 與 YEAR 欄位，用以分類每個月的變化；"
            f"缺少：{'、'.join(missing_my)}（位於來源檔案尾部，日後此模式恒為必要欄位）。"
        )

    avail = set(df_input.columns)
    missing_req = [
        name for name in _SAFM_REQUIRED
        if _resolve_column_name(name, avail) is None
    ]
    if missing_req:
        raise ValueError(f"缺少必需欄位: {', '.join(missing_req)}")

    warnings: list[str] = []
    df = df_input.copy()

    month_num = pd.to_numeric(df[month_col], errors="coerce")
    year_num = pd.to_numeric(df[year_col], errors="coerce")
    valid_mask = (
        month_num.notna() & year_num.notna()
        & (month_num.round() >= 1) & (month_num.round() <= 12)
        & (year_num.round() >= 1990) & (year_num.round() <= 2100)
    )
    invalid_count = int((~valid_mask).sum())
    if invalid_count:
        warnings.append(
            f"MONTH／YEAR 無法解析或超出範圍（月份須為 1–12 整數）共 {invalid_count} 列，已略過不計。"
        )
    df = df.loc[valid_mask].copy()
    if df.empty:
        raise ValueError("MONTH／YEAR 全數無效，無法分類每個月的變化。")
    df["_safm_month"] = month_num.loc[valid_mask].round().astype(int).to_numpy()
    df["_safm_year"] = year_num.loc[valid_mask].round().astype(int).to_numpy()
    df["_safm_ym_key"] = (df["_safm_year"] * 100 + df["_safm_month"]).astype(int)
    df[_SAFM_YM_COLUMN] = [
        _ym_label(y, m) for y, m in zip(df["_safm_year"], df["_safm_month"])
    ]

    ym_keys = sorted(df["_safm_ym_key"].unique().tolist())
    if len(ym_keys) < 2:
        raise ValueError(
            f"YEAR-MONTH 至少需有 2 個不同月份值方可作環比分析，目前僅有: "
            f"{sorted(df[_SAFM_YM_COLUMN].unique().tolist())}"
        )
    ym_labels = [_ym_label(k // 100, k % 100) for k in ym_keys]
    if progress_callback:
        progress_callback(0.05, "驗證通過")

    df = _clean_dataframe(df, shop_natures=shop_natures, progress_callback=progress_callback)
    if df.empty:
        raise ValueError("過濾後無有效資料列。")
    # _clean_dataframe 可能重建索引，上述 _safm_* 欄位隨列保留；重新對齊 YM 衍生欄
    ym_keys = sorted(df["_safm_ym_key"].unique().tolist())
    if len(ym_keys) < 2:
        raise ValueError("篩選 Shop Nature 後剩餘月份不足 2 個，無法作環比分析。")
    ym_labels = [_ym_label(k // 100, k % 100) for k in ym_keys]

    if "SKU" in df.columns:
        df["SKU"] = df["SKU"].apply(_safe_sku)
    if "SKU Desc" not in df.columns or "Shop Name" not in df.columns:
        df = _compute_raw_extra_columns(df)

    idx = df.index
    safety_col = _resolve_column_name("Adj. Safety Stock", set(df.columns))
    sold_col = _resolve_column_name("Sold Qty (CM - MTD)", set(df.columns))
    net_col = _resolve_column_name("Net Sales (CM - MTD)", set(df.columns))
    stock_col = _resolve_column_name("Shop Stock Qty (CM)", set(df.columns))
    rp_col = _resolve_column_name("RP Type", set(df.columns))

    safety_raw = pd.to_numeric(df[safety_col], errors="coerce")
    blank_safety = int(safety_raw.isna().sum())
    if blank_safety:
        warnings.append(
            f"Adj. Safety Stock 空白或非數值共 {blank_safety} 列，已視為 0（只看原始值增減，不作建議值計算）。"
        )
    safety_val = safety_raw.fillna(0).round().astype(int).to_numpy()

    sold_val = _safe_numeric(df[sold_col]).round().astype(int).to_numpy(dtype=float)
    net_val = pd.to_numeric(df[net_col], errors="coerce").fillna(0).to_numpy(dtype=float)
    stock_val = _safe_numeric(df[stock_col]).round().astype(int).to_numpy(dtype=float)

    rp_raw = (
        df[rp_col].fillna("").astype(str).str.strip().str.upper()
        if rp_col in df.columns else _safe_default_series("", idx).astype(str)
    )
    rp_disp = np.where(rp_raw == "", "空白", rp_raw)
    unusual_rp = sorted({str(v) for v in set(rp_disp.tolist())} - {"RF", "ND", "空白"})
    if unusual_rp:
        shown = "、".join(unusual_rp[:5])
        warnings.append(f"RP Type 出現非 RF／ND 值：{shown}（已按原值記錄轉變，不歸入 RF／ND）。")

    mch1_col = _resolve_column_name("Adj. MCH1", set(df.columns))
    sst_col = _resolve_column_name("Adj. SST", set(df.columns))
    if mch1_col is None:
        warnings.append("來源檔案缺少「Adj. MCH1」，By Adj. MCH1 與 By MCH1×SST 以「空白」列示。")
    if sst_col is None:
        warnings.append("來源檔案缺少「Adj. SST」，By Adj. SST 與 By MCH1×SST 以「空白」列示。")

    # ── 最新月份月化係數 ────────────────────────────────────────────────
    latest_key = int(ym_keys[-1])
    latest_year, latest_month = latest_key // 100, latest_key % 100
    month_total = calendar.monthrange(latest_year, latest_month)[1]
    factor = 1.0
    ref_iso = ""
    if ref_date is None:
        warnings.append("未提供數據截止日，最新月份不作月化（係數為 1）。")
    else:
        try:
            ref_d = _normalize_ref_date(ref_date)
        except ValueError as e:
            raise ValueError(f"數據截止日無效：{e}")
        ref_iso = ref_d.isoformat()
        if ref_d.year == latest_year and ref_d.month == latest_month:
            cutoff_day = min(max(int(ref_d.day), 1), int(month_total))
            if cutoff_day >= month_total:
                factor = 1.0
            else:
                factor = float(month_total) / float(cutoff_day)
        else:
            warnings.append(
                f"數據截止日 {ref_iso} 與最新月份 {_ym_label(latest_year, latest_month)}不一致，"
                "最新月份不作月化（係數為 1）。"
            )
    factor = round(float(factor), 4)

    is_latest = (df["_safm_ym_key"].to_numpy() == latest_key)
    sold_adj = np.where(is_latest, np.round(sold_val * factor), sold_val).astype(int)
    net_adj = np.where(is_latest, np.round(net_val * factor, 2), np.round(net_val, 2))

    df["_safm_safety"] = safety_val
    df["_safm_sold_raw"] = sold_val.astype(int)
    df["_safm_sold_adj"] = sold_adj
    df["_safm_net_raw"] = np.round(net_val, 2)
    df["_safm_net_adj"] = net_adj
    df["_safm_stock"] = stock_val.astype(int)
    df["_safm_rp"] = rp_disp
    df[_SAFM_SOLD_ADJ_COLUMN] = sold_adj
    df[_SAFM_NET_ADJ_COLUMN] = net_adj

    if progress_callback:
        progress_callback(0.25, "建立月度粒度的彙總...")
    site_s = df.get("Site Code", _safe_default_series("", idx)).fillna("").astype(str).str.strip()
    sku_s = df.get("SKU", _safe_default_series("", idx))
    brand_s = df.get("Brand", _safe_default_series("", idx)).fillna("").astype(str).str.strip()
    buyer_s = df.get("Buyer", _safe_default_series("", idx)).fillna("").astype(str).str.strip()
    mch1_s = (
        df[mch1_col].fillna("").astype(str).str.strip()
        if mch1_col in df.columns else _safe_default_series("", idx).astype(str)
    )
    sst_s = (
        df[sst_col].fillna("").astype(str).str.strip()
        if sst_col in df.columns else _safe_default_series("", idx).astype(str)
    )
    shop_name_s = df.get("Shop Name", _safe_default_series("", idx)).fillna("").astype(str)
    sku_desc_s = df.get("SKU Desc", _safe_default_series("", idx)).fillna("").astype(str)

    tmp = pd.DataFrame({
        "_site": site_s, "_sku": sku_s, "_ym": df["_safm_ym_key"],
        "_safety": df["_safm_safety"], "_sold_raw": df["_safm_sold_raw"],
        "_sold_adj": df["_safm_sold_adj"], "_net_raw": df["_safm_net_raw"],
        "_net_adj": df["_safm_net_adj"], "_stock": df["_safm_stock"],
        "_rp": df["_safm_rp"], "_brand": brand_s, "_buyer": buyer_s,
        "_mch1": mch1_s, "_sst": sst_s,
        "_shop": shop_name_s, "_desc": sku_desc_s,
    })

    dup_keys = tmp.duplicated(subset=["_site", "_sku", "_ym"]).sum()
    if int(dup_keys):
        warnings.append(
            f"偵測到「Site Code × SKU × YEAR-MONTH」重複列 {int(dup_keys)} 列；"
            "同一主鍵多列時，Sold／Net／Stock 以加總計，Safety／RP／MCH1／SST 取首個非空值。"
        )

    def _first_nonempty(s: pd.Series) -> str:
        s = s.fillna("").astype(str)
        s = s[s.str.strip() != ""]
        return s.iloc[0] if len(s) else ""

    monthly = tmp.groupby(["_site", "_sku", "_ym"], sort=True).agg(
        safety=("_safety", "first"),
        sold_raw=("_sold_raw", "sum"),
        sold_adj=("_sold_adj", "sum"),
        net_raw=("_net_raw", "sum"),
        net_adj=("_net_adj", "sum"),
        stock=("_stock", "sum"),
        rp=("_rp", "first"),
        brand=("_brand", _first_nonempty),
        buyer=("_buyer", _first_nonempty),
        mch1=("_mch1", _first_nonempty),
        sst=("_sst", _first_nonempty),
        shop=("_shop", _first_nonempty),
        desc=("_desc", _first_nonempty),
    ).reset_index()
    monthly = monthly.sort_values(["_site", "_sku", "_ym"]).reset_index(drop=True)

    if progress_callback:
        progress_callback(0.45, "計算相鄰月份環比事件...")
    change_rows: list[dict] = []
    for (_site, _sku), g in monthly.groupby(["_site", "_sku"], sort=False):
        g = g.sort_values("_ym").reset_index(drop=True)
        if len(g) < 2:
            continue
        for i in range(1, len(g)):
            prev = g.iloc[i - 1]
            curr = g.iloc[i]
            prev_key, curr_key = int(prev["_ym"]), int(curr["_ym"])
            safety_delta = int(curr["safety"]) - int(prev["safety"])
            if safety_delta > 0:
                safety_cat = "增加"
            elif safety_delta < 0:
                safety_cat = "減少"
            else:
                safety_cat = "不變"
            rp_prev, rp_curr = str(prev["rp"]), str(curr["rp"])
            sold_delta = int(curr["sold_adj"]) - int(prev["sold_adj"])
            net_delta = round(float(curr["net_adj"]) - float(prev["net_adj"]), 2)
            stock_delta = int(curr["stock"]) - int(prev["stock"])
            change_rows.append({
                "Site Code": _site, "Shop Name": curr["shop"], "SKU": _sku,
                "SKU Desc": curr["desc"] or prev["desc"],
                "Brand": curr["brand"] or prev["brand"],
                "Buyer": curr["buyer"] or prev["buyer"],
                "Adj. MCH1": curr["mch1"] or prev["mch1"],
                "Adj. SST": curr["sst"] or prev["sst"],
                "上月": _ym_label(prev_key // 100, prev_key % 100),
                "本月": _ym_label(curr_key // 100, curr_key % 100),
                "是否相鄰曆月": "是" if _is_calendar_adjacent(prev_key, curr_key) else "否",
                "上月 Adj. Safety Stock": int(prev["safety"]),
                "本月 Adj. Safety Stock": int(curr["safety"]),
                "Safety變化量": safety_delta, "Safety變化分類": safety_cat,
                "上月 RP Type": rp_prev, "本月 RP Type": rp_curr,
                "RP轉變": f"{rp_prev}→{rp_curr}",
                "上月 Sold Qty (CM-MTD)": int(prev["sold_raw"]),
                "本月 Sold Qty (CM-MTD)": int(curr["sold_raw"]),
                "上月 Sold月化": int(prev["sold_adj"]),
                "本月 Sold月化": int(curr["sold_adj"]),
                "Sold變化量(月化基準)": sold_delta,
                "Sold變化率": _change_rate(float(curr["sold_adj"]), float(prev["sold_adj"])),
                "上月 Net Sales (CM-MTD)": round(float(prev["net_raw"]), 2),
                "本月 Net Sales (CM-MTD)": round(float(curr["net_raw"]), 2),
                "上月 Net月化": round(float(prev["net_adj"]), 2),
                "本月 Net月化": round(float(curr["net_adj"]), 2),
                "Net變化量(月化基準)": net_delta,
                "Net變化率": _change_rate(float(curr["net_adj"]), float(prev["net_adj"])),
                "上月 Shop Stock Qty (CM)": int(prev["stock"]),
                "本月 Shop Stock Qty (CM)": int(curr["stock"]),
                "Stock變化量": stock_delta,
                "Stock變化率": _change_rate(float(curr["stock"]), float(prev["stock"])),
            })
    change_df = pd.DataFrame(change_rows, columns=_SAFM_CHANGE_COLUMNS)
    if not change_df.empty:
        change_df = change_df.sort_values(
            ["Brand", "SKU", "Site Code", "本月"], kind="stable"
        ).reset_index(drop=True)

    if progress_callback:
        progress_callback(0.65, "建立維度樞紐...")
    mch1_pivot = _build_safm_metric_pivot(monthly, "mch1", "Adj. MCH1", ym_keys)
    sst_pivot = _build_safm_metric_pivot(monthly, "sst", "Adj. SST", ym_keys)
    mch1_sst_pivot = _build_safm_metric_pivot(monthly, ["mch1", "sst"], None, ym_keys)
    shop_pivot = _build_safm_shop_pivot(monthly, ym_keys)

    if progress_callback:
        progress_callback(0.8, "整理 RAW 明細...")
    raw_df = df.copy()
    drop_internal = [c for c in raw_df.columns if str(c).startswith("_safm_")]
    raw_df = raw_df.drop(columns=[c for c in drop_internal if c not in ("_safm_ym_key",)], errors="ignore")
    raw_df = raw_df.drop(columns=["_safm_ym_key"], errors="ignore")

    safety_delta_total = int(change_df["Safety變化量"].sum()) if not change_df.empty else 0
    rp_counts = change_df["RP轉變"].value_counts().to_dict() if not change_df.empty else {}
    if not change_df.empty:
        _rf_to_nd = change_df["RP轉變"] == "RF→ND"
        _nd_to_rf = change_df["RP轉變"] == "ND→RF"
        rp_rf_to_nd = int(_rf_to_nd.sum())
        rp_nd_to_rf = int(_nd_to_rf.sum())
        rp_rf_to_nd_skus = int(change_df.loc[_rf_to_nd, "SKU"].nunique())
        rp_nd_to_rf_skus = int(change_df.loc[_nd_to_rf, "SKU"].nunique())
    else:
        rp_rf_to_nd = rp_nd_to_rf = rp_rf_to_nd_skus = rp_nd_to_rf_skus = 0
    rp_summary = _build_safm_rp_summary(change_df)
    rp_endpoints, rp_end_info = _build_safm_rp_endpoints(monthly, ym_keys)
    stats = {
        "ym_list": ym_labels,
        "ym_count": len(ym_labels),
        "latest_ym": ym_labels[-1],
        "month_total_days": int(month_total),
        "ref_date": ref_iso,
        "annual_factor": float(factor),
        "total_groups": int(monthly[["_site", "_sku"]].drop_duplicates().shape[0]),
        "total_skus": int(monthly["_sku"].nunique()),
        "total_shops": int(monthly["_site"].nunique()),
        "total_events": int(len(change_df)),
        "safety_inc": int((change_df["Safety變化分類"] == "增加").sum()) if not change_df.empty else 0,
        "safety_dec": int((change_df["Safety變化分類"] == "減少").sum()) if not change_df.empty else 0,
        "safety_same": int((change_df["Safety變化分類"] == "不變").sum()) if not change_df.empty else 0,
        "safety_delta_total": safety_delta_total,
        "rp_counts": {str(k): int(v) for k, v in rp_counts.items()},
        "rp_rf_to_nd": rp_rf_to_nd,
        "rp_nd_to_rf": rp_nd_to_rf,
        "rp_rf_to_nd_skus": rp_rf_to_nd_skus,
        "rp_nd_to_rf_skus": rp_nd_to_rf_skus,
        "rp_summary": rp_summary,
        "rp_endpoints": rp_endpoints,
        "rp_end_info": rp_end_info,
    }
    if change_df.empty:
        warnings.append("所有組合皆只有單一月份資料，無相鄰月份環比事件可輸出。")

    if progress_callback:
        progress_callback(1.0, "轉換完成")
    gc.collect()
    return change_df, mch1_pivot, sst_pivot, mch1_sst_pivot, shop_pivot, raw_df, stats, warnings


def _blank_dim(value: str) -> str:
    text = "" if value is None else str(value).strip()
    return text if text else "空白"


def _safm_metric_value(g: pd.DataFrame, metric: str):
    """維度單月合計（月化基準）：Net Sales 取 2 位小數，其餘取整。"""
    if metric == "Net Sales":
        return round(float(g["net_adj"].sum()), 2) if not g.empty else 0.0
    if metric == "Sold Qty":
        return int(g["sold_adj"].sum()) if not g.empty else 0
    if metric == "Net Stock Qty":
        return int(g["stock"].sum()) if not g.empty else 0
    return int(g["safety"].sum()) if not g.empty else 0


def _safm_mom_pct(curr: float, prev: float):
    """月環比%（字串）：上月為 0 時回傳 None（留空）；零變動顯示 0.0%。"""
    try:
        if float(prev) == 0:
            return None
        delta = (float(curr) - float(prev)) / abs(float(prev))
    except (TypeError, ValueError, ZeroDivisionError):
        return None
    if delta == 0:
        return "0.0%"
    return f"{delta:+.1%}"


def _safm_dim_cells(dim_val: str, dim_cols: list[str]) -> dict:
    if dim_val == "總計":
        return {c: ("總計" if i == 0 else "") for i, c in enumerate(dim_cols)}
    if len(dim_cols) == 2:
        parts = str(dim_val).split("｜", 1)
        return {
            "Adj. MCH1": _blank_dim(parts[0] if len(parts) > 0 else ""),
            "Adj. SST": _blank_dim(parts[1] if len(parts) > 1 else ""),
        }
    return {dim_cols[0]: _blank_dim(dim_val)}


def _build_safm_metric_pivot(
    monthly: pd.DataFrame, group_keys, header, ym_keys: list[int]
) -> pd.DataFrame:
    """維度樞紐（三段式：每月合計 → 月環比% → 月環比絕對值，四項指標齊列）。

    每個指標連續三列（項目＝「{指標} 每月合計」「{指標} 月環比%」「{指標} 月環比」），
    欄為維度欄＋項目＋各 YEAR-MONTH＋全部合計（環比兩列的全部合計留空）。
    月環比% 為顯示字串（上月為 0 或首月留空）；其餘為數值。
    """
    if monthly is None or monthly.empty:
        return pd.DataFrame()
    df = monthly.copy()
    if isinstance(group_keys, list):
        df["_d1"] = df[group_keys[0]].fillna("").astype(str).str.strip()
        df["_d2"] = df[group_keys[1]].fillna("").astype(str).str.strip()
        df["_dim"] = df["_d1"] + "｜" + df["_d2"]
        dim_cols = ["Adj. MCH1", "Adj. SST"]
    else:
        df["_dim"] = df[group_keys].fillna("").astype(str).str.strip().apply(_blank_dim)
        dim_cols = [header]
    ym_labels = [_ym_label(k // 100, k % 100) for k in ym_keys]
    dim_vals = sorted(df["_dim"].unique().tolist())

    base: dict[tuple[str, int, str], float] = {}
    for dim_val in dim_vals:
        for key in ym_keys:
            gg = df[(df["_dim"] == dim_val) & (df["_ym"] == key)]
            for metric in _SAFM_PIVOT_METRICS:
                base[(dim_val, key, metric)] = _safm_metric_value(gg, metric)

    def _total(dim_val: str | None, key: int, metric: str):
        vals = [base[(d, key, metric)] for d in (dim_vals if dim_val is None else [dim_val])]
        return round(float(sum(vals)), 2) if metric == "Net Sales" else int(sum(vals))

    rows = []
    for metric in _SAFM_PIVOT_METRICS:
        for block in ("每月合計", "月環比%", "月環比"):
            for dim_val in dim_vals + ["總計"]:
                row = _safm_dim_cells(dim_val, dim_cols)
                row["項目"] = f"{metric} {block}"
                period_total = 0.0 if metric == "Net Sales" else 0
                for i, (key, label) in enumerate(zip(ym_keys, ym_labels)):
                    curr = _total(None if dim_val == "總計" else dim_val, key, metric)
                    if block == "每月合計":
                        row[label] = curr
                        period_total += curr
                    elif block == "月環比%":
                        if i == 0:
                            row[label] = None
                        else:
                            prev_key = ym_keys[i - 1]
                            prev = _total(None if dim_val == "總計" else dim_val, prev_key, metric)
                            row[label] = _safm_mom_pct(curr, prev)
                    else:
                        if i == 0:
                            row[label] = None
                        else:
                            prev_key = ym_keys[i - 1]
                            prev = _total(None if dim_val == "總計" else dim_val, prev_key, metric)
                            row[label] = (
                                round(curr - prev, 2)
                                if metric == "Net Sales" else int(curr - prev)
                            )
                if block == "每月合計":
                    row["全部合計"] = (
                        round(period_total, 2)
                        if metric == "Net Sales" else int(period_total)
                    )
                else:
                    row["全部合計"] = None
                rows.append(row)
    cols = dim_cols + ["項目"] + ym_labels + ["全部合計"]
    return pd.DataFrame(rows, columns=cols)


def _build_safm_shop_pivot(monthly: pd.DataFrame, ym_keys: list[int]) -> pd.DataFrame:
    """By Shop：每間店舖每月合計＋環比變化（月化基準）。"""
    if monthly is None or monthly.empty:
        return pd.DataFrame(columns=_SAFM_SHOP_COLUMNS)
    df = monthly.copy()
    ym_labels = {k: _ym_label(k // 100, k % 100) for k in ym_keys}
    rows = []
    for site, g_site in df.groupby("_site", sort=True):
        shop_name = ""
        _names = g_site["shop"].fillna("").astype(str)
        _names = _names[_names.str.strip() != ""]
        if len(_names):
            shop_name = str(_names.iloc[0])
        prev_sold = prev_net = prev_safety = None
        for key in ym_keys:
            gg = g_site[g_site["_ym"] == key]
            sku_n = int(gg["_sku"].nunique()) if not gg.empty else 0
            safety_sum = int(gg["safety"].sum()) if not gg.empty else 0
            sold_raw = int(gg["sold_raw"].sum()) if not gg.empty else 0
            sold_adj = int(gg["sold_adj"].sum()) if not gg.empty else 0
            net_raw = round(float(gg["net_raw"].sum()), 2) if not gg.empty else 0.0
            net_adj = round(float(gg["net_adj"].sum()), 2) if not gg.empty else 0.0
            stock_sum = int(gg["stock"].sum()) if not gg.empty else 0
            if prev_sold is None:
                sold_chg = net_chg = safety_chg = None
            else:
                sold_chg = sold_adj - prev_sold
                net_chg = round(net_adj - prev_net, 2)
                safety_chg = safety_sum - prev_safety
            rows.append({
                "Site Code": site, "Shop Name": shop_name,
                _SAFM_YM_COLUMN: ym_labels[key],
                "SKU數": sku_n, "Safety合計": safety_sum,
                "Sold原始合計": sold_raw, "Sold月化合計": sold_adj,
                "Net原始合計": net_raw, "Net月化合計": net_adj,
                "Stock合計": stock_sum,
                "Sold環比變化": sold_chg, "Net環比變化": net_chg,
                "Safety環比變化": safety_chg,
            })
            prev_sold, prev_net, prev_safety = sold_adj, net_adj, safety_sum
    return pd.DataFrame(rows, columns=_SAFM_SHOP_COLUMNS)


def _build_safm_rp_summary(change_df: pd.DataFrame) -> pd.DataFrame:
    """RP Type 轉變 × 月份配對統計（事件數／SKU數，含總計列）。

    列為四類轉變（RF→RF、ND→ND、RF→ND、ND→RF，未知值附加於後），
    欄為各「上月→本月」配對的事件數／SKU數＋總計；總計列 SKU 數為不重複 SKU。
    """
    if change_df is None or change_df.empty:
        return pd.DataFrame(columns=["RP轉變"])
    df = change_df.copy()
    df["_pair"] = df["上月"].fillna("").astype(str) + "→" + df["本月"].fillna("").astype(str)
    pairs = sorted(df["_pair"].unique().tolist())
    present = df["RP轉變"].fillna("").astype(str).tolist()
    order = [t for t in _SAFM_RP_ORDER if t in set(present)]
    order += sorted(set(present) - set(order) - {""})
    rows = []
    for trans in order:
        g = df[df["RP轉變"] == trans]
        row = {"RP轉變": trans}
        for p in pairs:
            gg = g[g["_pair"] == p]
            row[f"{p} 事件數"] = int(len(gg))
            row[f"{p} SKU數"] = int(gg["SKU"].nunique()) if not gg.empty else 0
        row["總計 事件數"] = int(len(g))
        row["總計 SKU數"] = int(g["SKU"].nunique())
        rows.append(row)
    total = {"RP轉變": "總計"}
    for p in pairs:
        gg = df[df["_pair"] == p]
        total[f"{p} 事件數"] = int(len(gg))
        total[f"{p} SKU數"] = int(gg["SKU"].nunique()) if not gg.empty else 0
    total["總計 事件數"] = int(len(df))
    total["總計 SKU數"] = int(df["SKU"].nunique())
    rows.append(total)
    cols = ["RP轉變"]
    for p in pairs:
        cols += [f"{p} 事件數", f"{p} SKU數"]
    cols += ["總計 事件數", "總計 SKU數"]
    return pd.DataFrame(rows, columns=cols)


_SAFM_RP_ENDPOINT_COLUMNS = [
    "Site Code", "Shop Name", "SKU", "SKU Desc", "Brand", "Buyer",
    "Adj. MCH1", "Adj. SST",
    "最早月份", "最早 RP Type", "最新月份", "最新 RP Type",
    "首尾轉變", "是否轉變",
]


def _build_safm_rp_endpoints(monthly: pd.DataFrame, ym_keys: list[int]
                             ) -> tuple[pd.DataFrame, dict]:
    """首尾月份 RP 對比明細：每組取最早與最新月份各一列比較。

    只收錄最早與最新月份皆有資料的組合；缺任一者不納入並計入 excluded。
    回傳 (明細表, {"first": 首月標籤, "last": 尾月標籤, "comparable": 可比組數,
    "excluded": 未納入組數, "rf_to_nd": 首尾 RF→ND 組數, "nd_to_rf": 首尾 ND→RF 組數})。
    """
    empty_info = {"first": "", "last": "", "comparable": 0, "excluded": 0,
                  "rf_to_nd": 0, "nd_to_rf": 0}
    if monthly is None or monthly.empty or not ym_keys or len(ym_keys) < 2:
        return pd.DataFrame(columns=_SAFM_RP_ENDPOINT_COLUMNS), empty_info
    first_key, last_key = int(ym_keys[0]), int(ym_keys[-1])
    first_label = _ym_label(first_key // 100, first_key % 100)
    last_label = _ym_label(last_key // 100, last_key % 100)
    first_m = monthly[monthly["_ym"] == first_key].set_index(["_site", "_sku"])
    last_m = monthly[monthly["_ym"] == last_key].set_index(["_site", "_sku"])
    common = first_m.index.intersection(last_m.index)
    excluded = int(monthly[["_site", "_sku"]].drop_duplicates().shape[0] - len(common))

    def _pick(latest_val, earliest_val) -> str:
        latest_text = "" if pd.isna(latest_val) else str(latest_val).strip()
        if latest_text:
            return latest_text
        return "" if pd.isna(earliest_val) else str(earliest_val).strip()

    rows = []
    for site, sku in common.tolist():
        prev = first_m.loc[(site, sku)]
        curr = last_m.loc[(site, sku)]
        if isinstance(prev, pd.DataFrame):
            prev = prev.iloc[0]
        if isinstance(curr, pd.DataFrame):
            curr = curr.iloc[0]
        rp_prev, rp_curr = str(prev["rp"]), str(curr["rp"])
        trans = f"{rp_prev}→{rp_curr}"
        rows.append({
            "Site Code": site, "Shop Name": _pick(curr["shop"], prev["shop"]),
            "SKU": sku, "SKU Desc": _pick(curr["desc"], prev["desc"]),
            "Brand": _pick(curr["brand"], prev["brand"]),
            "Buyer": _pick(curr["buyer"], prev["buyer"]),
            "Adj. MCH1": _pick(curr["mch1"], prev["mch1"]),
            "Adj. SST": _pick(curr["sst"], prev["sst"]),
            "最早月份": first_label, "最早 RP Type": rp_prev,
            "最新月份": last_label, "最新 RP Type": rp_curr,
            "首尾轉變": trans, "是否轉變": "是" if rp_prev != rp_curr else "否",
        })
    detail = pd.DataFrame(rows, columns=_SAFM_RP_ENDPOINT_COLUMNS)
    if not detail.empty:
        _order = {t: i for i, t in enumerate(_SAFM_RP_ORDER)}
        detail["_changed"] = (detail["是否轉變"] == "是").astype(int)
        detail["_tord"] = detail["首尾轉變"].map(_order).fillna(99).astype(int)
        detail = detail.sort_values(
            ["_changed", "_tord", "Brand", "SKU", "Site Code"],
            ascending=[False, True, True, True, True], kind="stable",
        ).drop(columns=["_changed", "_tord"]).reset_index(drop=True)
    info = {
        "first": first_label, "last": last_label,
        "comparable": int(len(detail)), "excluded": excluded,
        "rf_to_nd": int((detail["首尾轉變"] == "RF→ND").sum()) if not detail.empty else 0,
        "nd_to_rf": int((detail["首尾轉變"] == "ND→RF").sum()) if not detail.empty else 0,
    }
    return detail, info


def safety_monthly_dashboard_metrics(change_df: pd.DataFrame, stats: dict | None = None):
    """月度摘要指標卡（與 Streamlit 看板及 Excel 摘要同一準則）。"""
    stats = stats or {}
    total_events = int(stats.get("total_events", len(change_df) if change_df is not None else 0))
    row1 = [
        ("環比事件數", total_events, "相鄰月份比較事件總數", "3B82F6"),
        ("Safety 增加事件", int(stats.get("safety_inc", 0)), "原始值增加的事件數", "EF4444"),
        ("Safety 減少事件", int(stats.get("safety_dec", 0)), "原始值減少的事件數", "10B981"),
        ("Safety 不變事件", int(stats.get("safety_same", 0)), "原始值不變的事件數", "64748B"),
    ]
    row2 = [
        ("Safety 淨變化量", int(stats.get("safety_delta_total", 0)), "全部事件變化量合計（件）", "8B5CF6"),
        ("涉及 SKU 數", int(stats.get("total_skus", 0)), "不重複 SKU（已刪除重複）", "F59E0B"),
        ("涉及店舖數", int(stats.get("total_shops", 0)), "不重複店舖數", "06B6D4"),
        ("覆蓋月份數", int(stats.get("ym_count", 0)), "YEAR-MONTH 不重複數", "EC4899"),
    ]
    row3 = [
        ("RF→ND 事件", int(stats.get("rp_rf_to_nd", 0)), "可補貨轉不可補貨事件數", "EF4444"),
        ("ND→RF 事件", int(stats.get("rp_nd_to_rf", 0)), "不可補貨轉可補貨事件數", "10B981"),
        ("RF→ND SKU數", int(stats.get("rp_rf_to_nd_skus", 0)), "曾由 RF 轉 ND 的 SKU（已刪除重複）", "F59E0B"),
        ("ND→RF SKU數", int(stats.get("rp_nd_to_rf_skus", 0)), "曾由 ND 轉 RF 的 SKU（已刪除重複）", "06B6D4"),
    ]
    return row1, row2, row3


def _write_safm_pivot_sheet(ws, df: pd.DataFrame, progress_callback=None,
                            progress_start=0.0, progress_end=0.1):
    """三段式樞紐工作表：項目欄按指標配色分段，總計列加粗灰底。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    total_font = Font(bold=True, size=11)
    center = Alignment(horizontal="center", vertical="center")
    center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if df is None or df.empty:
        ws.cell(row=1, column=1, value="無符合資料").font = Font(size=12)
        return
    headers = list(df.columns)
    item_idx = headers.index("項目") if "項目" in headers else -1
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_wrap
        cell.border = thin_border
    total = len(df)
    for row_idx, row_data in enumerate(df.itertuples(index=False, name=None), 2):
        row_dict = dict(zip(headers, row_data))
        item_text = str(row_dict.get("項目", ""))
        metric = item_text.split(" ")[0] if " " in item_text else ""
        is_total = any(str(row_dict.get(c, "")) == "總計" for c in headers[:2])
        is_money = item_text.startswith("Net Sales")
        is_pct_block = item_text.endswith("月環比%")
        for col_idx, (h, v) in enumerate(zip(headers, row_data), 1):
            if pd.isna(v):
                v = None
            cell = ws.cell(row=row_idx, column=col_idx, value=v)
            cell.border = thin_border
            if col_idx - 1 == item_idx and metric in _SAFM_METRIC_COLORS:
                cell.fill = PatternFill(
                    start_color=_SAFM_METRIC_COLORS[metric],
                    end_color=_SAFM_METRIC_COLORS[metric], fill_type="solid",
                )
                cell.font = Font(bold=True, color="FFFFFF", size=11)
                cell.alignment = center
            elif is_total:
                cell.font = total_font
                cell.fill = total_fill
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    cell.number_format = "#,##0.00" if is_money else "#,##0"
                    cell.alignment = center
                else:
                    cell.alignment = center_wrap
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.number_format = "#,##0.00" if is_money else "#,##0"
                cell.alignment = center
            else:
                cell.alignment = center_wrap
        if progress_callback and (row_idx % 500 == 0 or row_idx == total + 1):
            p = progress_start + (progress_end - progress_start) * (row_idx / max(total + 1, 1))
            progress_callback(p, f"寫入工作表 ({row_idx-1}/{total})...")
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "C2"
    _auto_width(ws)


def _write_safm_block(ws, df: pd.DataFrame, start_row: int = 1,
                      progress_callback=None, progress_start=0.0, progress_end=0.1,
                      total_hint: int | None = None) -> int:
    """於工作表寫入一個表格區塊（表頭＋資料列），回傳下一個可用列號，不設篩選與凍結。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    if df is None or df.empty:
        ws.cell(row=start_row, column=1, value="無符合資料").font = Font(size=12)
        return start_row + 2
    headers = list(df.columns)
    display = [_display_safety_header(h) for h in headers]
    for col_idx, h in enumerate(display, 1):
        cell = ws.cell(row=start_row, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    total = len(df)
    for offset, row_data in enumerate(df.itertuples(index=False, name=None), 1):
        row_idx = start_row + offset
        for col_idx, (h, v) in enumerate(zip(headers, row_data), 1):
            if pd.isna(v):
                v = None
            cell = ws.cell(row=row_idx, column=col_idx, value=v)
            cell.border = thin_border
            if isinstance(v, float):
                if "變化率" in str(h):
                    cell.number_format = "0.00%"
                elif "Net" in str(h):
                    cell.number_format = "#,##0.00"
                else:
                    cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if progress_callback and (offset % 2000 == 0 or offset == total):
            denom = max(total_hint or total, 1)
            p = progress_start + (progress_end - progress_start) * (offset / denom)
            progress_callback(p, f"寫入工作表 ({offset}/{total})...")
    return start_row + total + 1


def _write_safm_table_sheet(ws, df: pd.DataFrame, progress_callback=None,
                            progress_start=0.0, progress_end=0.1):
    """單表工作表（表頭＋資料＋自動篩選＋凍結首列）。"""
    _write_safm_block(ws, df, start_row=1, progress_callback=progress_callback,
                      progress_start=progress_start, progress_end=progress_end)
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    _auto_width(ws)


def _write_safm_rp_sheet(ws, summary_df: pd.DataFrame, detail_df: pd.DataFrame,
                         end_title: str, progress_callback=None,
                         progress_start=0.0, progress_end=0.1):
    """RP Type 變化工作表：上半為轉變 × 月份配對統計，下半為首尾月份對比明細。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    next_row = _write_safm_block(ws, summary_df, start_row=1,
                                 progress_callback=progress_callback,
                                 progress_start=progress_start,
                                 progress_end=(progress_start + progress_end) / 2)
    next_row += 1
    detail_cols = len(detail_df.columns) if detail_df is not None and not detail_df.empty else 1
    ws.merge_cells(start_row=next_row, start_column=1,
                   end_row=next_row, end_column=max(detail_cols, 1))
    title_cell = ws.cell(row=next_row, column=1, value=end_title)
    title_cell.font = Font(bold=True, size=12, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="0077A3", end_color="0077A3", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[next_row].height = 20
    detail_header_row = next_row + 1
    _write_safm_block(ws, detail_df, start_row=detail_header_row,
                      progress_callback=progress_callback,
                      progress_start=(progress_start + progress_end) / 2,
                      progress_end=progress_end)
    if detail_df is not None and not detail_df.empty:
        ws.auto_filter.ref = f"A{detail_header_row}:{ws.cell(row=ws.max_row, column=len(detail_df.columns)).coordinate}"
        ws.freeze_panes = f"A{detail_header_row + 1}"
    else:
        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"
    _auto_width(ws)


def _write_safety_monthly_guide_sheet(ws, stats):
    st_ = stats or {}
    ym_list = st_.get("ym_list") or []
    factor = st_.get("annual_factor", 1.0)
    params = [
        ("MONTH／YEAR", "此模式恒為必要欄位（位於來源檔案尾部），用以分類每個月的變化；缺任一即停止。"),
        ("覆蓋月份", "、".join(ym_list) if ym_list else "—"),
        ("數據截止日", f'{st_.get("ref_date") or "—"}（最新月份 {st_.get("latest_ym") or "—"} 月化係數＝當月總日數 ÷ 截止日日號＝{factor}；舊月份係數為 1）'),
        ("環比方式", "相鄰月份環比（每組 Site Code × SKU 與上一個有資料月份比較；上月缺資料者標示新增組合，不計變化）"),
        ("分析欄", "Adj. Safety Stock（只看原始值增減）／RP Type 轉變／Sold Qty (CM - MTD)／Net Sales (CM - MTD)／Shop Stock Qty (CM)；Sold 與 Net 同時月化，Stock 與 Safety 為快照不月化"),
        ("維度", "Adj. MCH1、Adj. SST、MCH1×SST、Site Code"),
    ]
    logic = [
        "YEAR-MONTH 由 YEAR × 100 ＋ MONTH 排序；MONTH 須為 1–12 整數，無效列略過並列警告。",
        "最新月份月化：Sold月化＝Sold Qty (CM - MTD) × 係數；Net月化＝Net Sales (CM - MTD) × 係數；舊月份月化值等於原始值。",
        "Safety變化量＝本月 Adj. Safety Stock − 上月；正數為增加、負數為減少、零為不變（只看原始值，不作建議值計算）。",
        "RP轉變記錄上月→本月（如 RF→ND）；Sold／Net 變化量以月化基準計算，變化率＝變化量 ÷ 上月絕對值（上月為 0 時留空）。",
        "樞紐四項指標：Sold Qty＝Sold Qty (CM - MTD) 月化合計；Net Sales＝Net Sales (CM - MTD) 月化合計；"
        "Net Stock Qty＝店舖存貨 Shop Stock Qty (CM) 合計（時點快照）；Safety Stock Qty＝Adj. Safety Stock 合計。",
        "樞紐三段式：每月合計 → 月環比%（上月為 0 或首月留空）→ 月環比絕對值；項目欄按指標配色分段。",
    ]
    sheet_rows = [
        ("說明（本頁）", "月化邏輯、本次參數與導讀", "先核對覆蓋月份與月化係數是否符合預期"),
        ("SKU月度變化", "每組相鄰月份環比事件（上月／本月值＋變化量＋變化率）", "篩選 Safety 增加／減少，查看同期 Sold Qty 與 Net Sales 變化"),
        ("By Adj. MCH1", "各 MCH1 三段式樞紐（四項指標 × 每月合計／月環比%／月環比）", "先看每月合計，再看月環比% 找變動最大的品類"),
        ("By Adj. SST", "各 SST 三段式樞紐", "比較不同供應類型每月生意變化"),
        ("By MCH1×SST", "MCH1 與 SST 組合三段式樞紐", "細分品類 × 供應類型"),
        ("By Shop", "每間店舖每月合計＋環比變化", "多店舖檔案時比較各店走勢"),
        ("RP Type 變化", "RP轉變 × 月份配對統計＋首尾月份對比明細", "重點看 RF→ND 與 ND→RF 兩列：哪些月份最多轉換；下半首尾對比看整個期間的最終轉變"),
        ("月度摘要", "指標卡（含 RF→ND／ND→RF 一排）＋每月合計", "管理層概覽"),
        ("RAW", "原始明細＋YEAR-MONTH 與月化估算欄", "核對逐列數值"),
    ]
    tips = [
        "最新月份為非足月時，Sold 與 Net 的環比已按月化基準計算；原始 MTD 值仍保留於主表以供核對。",
        "數據截止日與最新月份不一致時，最新月份不作月化（係數為 1），並列出警告。",
        "上月缺資料的組合不計變化，屬新增組合；變化率分母為 0 時留空。",
        "RF→ND 的 SKU 轉為不可補貨後，留意其後月份 Sold Qty 與 Net Sales 是否受影響；ND→RF 則留意補貨是否恢復。",
    ]
    _write_guide_sheet(
        ws, title="Safety Stock 月度變化分析 說明與導讀",
        subtitle="本頁講解月化邏輯與本次執行參數，並逐頁說明應留意的部分；各項判定以 SKU月度變化逐列欄位為準。",
        params=params, logic=logic, sheet_rows=sheet_rows, tips=tips,
    )


def generate_excel_safety_monthly(
    change_df: pd.DataFrame,
    mch1_pivot: pd.DataFrame | None = None,
    sst_pivot: pd.DataFrame | None = None,
    mch1_sst_pivot: pd.DataFrame | None = None,
    shop_pivot: pd.DataFrame | None = None,
    raw_df: pd.DataFrame | None = None,
    stats: dict | None = None,
    progress_callback=None,
    include_raw: bool = True,
) -> BytesIO:
    """Safety Stock 月度變化分析輸出九個工作表（說明預設開啟）。"""
    wb = Workbook()
    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    ws_guide = wb.active
    ws_guide.title = "說明"
    _write_safety_monthly_guide_sheet(ws_guide, stats)
    try:
        wb.active = wb.sheetnames.index("說明")
    except ValueError:
        pass

    ws_change = wb.create_sheet(title="SKU月度變化")
    _write_safm_table_sheet(ws_change, change_df, progress_callback, 0.05, 0.35)

    if progress_callback:
        progress_callback(0.4, "建立維度樞紐...")
    ws_mch1 = wb.create_sheet(title="By Adj. MCH1")
    _write_safm_pivot_sheet(ws_mch1, mch1_pivot, progress_callback, 0.4, 0.48)
    ws_sst = wb.create_sheet(title="By Adj. SST")
    _write_safm_pivot_sheet(ws_sst, sst_pivot, progress_callback, 0.48, 0.56)
    ws_cross = wb.create_sheet(title="By MCH1×SST")
    _write_safm_pivot_sheet(ws_cross, mch1_sst_pivot, progress_callback, 0.56, 0.64)
    ws_shop = wb.create_sheet(title="By Shop")
    _write_safm_table_sheet(ws_shop, shop_pivot, progress_callback, 0.64, 0.7)

    if progress_callback:
        progress_callback(0.7, "建立 RP Type 變化...")
    ws_rp = wb.create_sheet(title="RP Type 變化")
    _rp_summary = (stats or {}).get("rp_summary")
    if _rp_summary is None:
        _rp_summary = _build_safm_rp_summary(change_df)
    _rp_endpoints = (stats or {}).get("rp_endpoints")
    _rp_end_info = (stats or {}).get("rp_end_info") or {}
    if _rp_end_info.get("first") and _rp_end_info.get("last"):
        _end_title = (
            f'首尾月份對比（{_rp_end_info["first"]}→{_rp_end_info["last"]}，'
            f'共 {_rp_end_info.get("comparable", 0):,} 組可比，'
            f'{_rp_end_info.get("excluded", 0):,} 組缺首尾月份資料未納入；有轉變列置前）'
        )
    else:
        _end_title = "首尾月份對比"
    _write_safm_rp_sheet(ws_rp, _rp_summary, _rp_endpoints, _end_title,
                         progress_callback, 0.7, 0.76)

    if progress_callback:
        progress_callback(0.78, "建立月度摘要...")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    desc_font = Font(size=9, color="6B7280", italic=True)
    ws_dash = wb.create_sheet(title="月度摘要")
    row1, row2, row3 = safety_monthly_dashboard_metrics(change_df, stats)
    _write_metric_cards(
        ws_dash, "Safety Stock 月度變化摘要 (Monthly Safety Review Dashboard)",
        row1, row2, thin_border, desc_font, progress_callback,
        progress_pct=0.82, row3_metrics=row3,
    )
    ym_list = (stats or {}).get("ym_list") or []
    if ym_list and shop_pivot is not None and not shop_pivot.empty:
        start_row = 16
        ws_dash.cell(row=start_row, column=1, value="每月合計（全店舖月化基準）").font = Font(bold=True, size=12, color="1E40AF")
        headers = ["YEAR-MONTH", "SKU 數", "Safety Stock Qty 合計",
                   "Sold Qty 合計 (月化)", "Net Sales 合計 (月化)",
                   "Net Stock Qty 合計", "RF→ND 事件數", "ND→RF 事件數"]
        for col_idx, h in enumerate(headers, 1):
            cell = ws_dash.cell(row=start_row + 1, column=col_idx, value=h)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        agg = shop_pivot.groupby(_SAFM_YM_COLUMN, sort=True).agg(
            SKU數=("SKU數", "sum"), Safety合計=("Safety合計", "sum"),
            Sold月化合計=("Sold月化合計", "sum"), Net月化合計=("Net月化合計", "sum"),
            Stock合計=("Stock合計", "sum"),
        ).reset_index()
        _rp_month = {}
        if change_df is not None and not change_df.empty and "本月" in change_df.columns:
            _rp_ct = change_df.groupby("本月")["RP轉變"].value_counts().unstack(fill_value=0)
            for _ym, _row in _rp_ct.iterrows():
                _rp_month[str(_ym)] = {
                    "RF→ND 事件數": int(_row.get("RF→ND", 0)),
                    "ND→RF 事件數": int(_row.get("ND→RF", 0)),
                }
        order = {label: i for i, label in enumerate(ym_list)}
        agg["_o"] = agg[_SAFM_YM_COLUMN].map(order).fillna(9999)
        agg = agg.sort_values("_o").drop(columns=["_o"])
        for r_idx, (_, r) in enumerate(agg.iterrows(), start_row + 2):
            _ym_label_val = str(r.get(_SAFM_YM_COLUMN, ""))
            _vals = {
                "YEAR-MONTH": r.get(_SAFM_YM_COLUMN),
                "SKU 數": r.get("SKU數"),
                "Safety Stock Qty 合計": r.get("Safety合計"),
                "Sold Qty 合計 (月化)": r.get("Sold月化合計"),
                "Net Sales 合計 (月化)": r.get("Net月化合計"),
                "Net Stock Qty 合計": r.get("Stock合計"),
                "RF→ND 事件數": _rp_month.get(_ym_label_val, {}).get("RF→ND 事件數", 0),
                "ND→RF 事件數": _rp_month.get(_ym_label_val, {}).get("ND→RF 事件數", 0),
            }
            for col_idx, h in enumerate(headers, 1):
                v = _vals.get(h)
                cell = ws_dash.cell(row=r_idx, column=col_idx, value=None if pd.isna(v) else v)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    cell.number_format = "#,##0.00" if h.startswith("Net Sales") else "#,##0"
        _auto_width(ws_dash)

    if progress_callback:
        progress_callback(0.9, "建立 RAW...")
    ws_raw = wb.create_sheet(title="RAW")
    if include_raw:
        _write_raw_sheet(ws_raw, raw_df, progress_callback, 0.9, 0.97)
    else:
        _write_raw_placeholder(ws_raw)

    if progress_callback:
        progress_callback(0.98, "正在儲存 Excel 檔案...")
    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf
