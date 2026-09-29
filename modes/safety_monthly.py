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
]

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


def _build_safm_metric_pivot(
    monthly: pd.DataFrame, group_keys, header, ym_keys: list[int]
) -> pd.DataFrame:
    """維度每月合計樞紐（月化基準）：每列一維度值，各 YM 五項合計＋全部合計。"""
    if monthly is None or monthly.empty:
        return pd.DataFrame()
    df = monthly.copy()
    if isinstance(group_keys, list):
        df["_dim1"] = df[group_keys[0]].fillna("").astype(str).str.strip()
        df["_dim2"] = df[group_keys[1]].fillna("").astype(str).str.strip()
        df["_dim"] = df["_dim1"] + "｜" + df["_dim2"]
        dim_cols = ["Adj. MCH1", "Adj. SST"]
        df["Adj. MCH1"] = df["_dim1"].apply(_blank_dim)
        df["Adj. SST"] = df["_dim2"].apply(_blank_dim)
    else:
        df["_dim"] = df[group_keys].fillna("").astype(str).str.strip().apply(_blank_dim)
        dim_cols = [header]
        df[header] = df["_dim"]
    ym_labels = [_ym_label(k // 100, k % 100) for k in ym_keys]
    rows = []
    for dim_val, g in df.groupby("_dim", sort=True):
        row: dict = {}
        if isinstance(group_keys, list):
            parts = str(dim_val).split("｜", 1)
            row["Adj. MCH1"] = _blank_dim(parts[0] if len(parts) > 0 else "")
            row["Adj. SST"] = _blank_dim(parts[1] if len(parts) > 1 else "")
        else:
            row[header] = _blank_dim(dim_val)
        total_net = 0.0
        total_sold = 0
        for key, label in zip(ym_keys, ym_labels):
            gg = g[g["_ym"] == key]
            net_sum = round(float(gg["net_adj"].sum()), 2) if not gg.empty else 0.0
            sold_sum = int(gg["sold_adj"].sum()) if not gg.empty else 0
            stock_sum = int(gg["stock"].sum()) if not gg.empty else 0
            safety_sum = int(gg["safety"].sum()) if not gg.empty else 0
            sku_n = int(gg["_sku"].nunique()) if not gg.empty else 0
            row[f"Net月化合計({label})"] = net_sum
            row[f"Sold月化合計({label})"] = sold_sum
            row[f"Stock合計({label})"] = stock_sum
            row[f"Safety合計({label})"] = safety_sum
            row[f"SKU數({label})"] = sku_n
            total_net += net_sum
            total_sold += sold_sum
        row["Net月化合計(全部)"] = round(total_net, 2)
        row["Sold月化合計(全部)"] = total_sold
        rows.append(row)
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    if isinstance(group_keys, list):
        out = out.sort_values(["Adj. MCH1", "Adj. SST"], kind="stable").reset_index(drop=True)
        total_row = {"Adj. MCH1": "總計", "Adj. SST": ""}
    else:
        out = out.sort_values(header, kind="stable").reset_index(drop=True)
        total_row = {header: "總計"}
    for c in out.columns:
        if c in dim_cols:
            continue
        try:
            total_row[c] = out[c].sum()
        except TypeError:
            total_row[c] = None
    if "Net月化合計(全部)" in total_row:
        total_row["Net月化合計(全部)"] = round(float(total_row["Net月化合計(全部)"]), 2)
    out = pd.concat([out, pd.DataFrame([total_row])], ignore_index=True)
    return out


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
    return row1, row2


def _write_safm_table_sheet(ws, df: pd.DataFrame, progress_callback=None,
                            progress_start=0.0, progress_end=0.1):
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    if df is None or df.empty:
        ws.cell(row=1, column=1, value="無符合資料").font = Font(size=12)
        return
    headers = list(df.columns)
    display = [_display_safety_header(h) for h in headers]
    for col_idx, h in enumerate(display, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    total = len(df)
    for row_idx, row_data in enumerate(df.itertuples(index=False, name=None), 2):
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
        if progress_callback and (row_idx % 2000 == 0 or row_idx == total + 1):
            p = progress_start + (progress_end - progress_start) * (row_idx / max(total + 1, 1))
            progress_callback(p, f"寫入工作表 ({row_idx-1}/{total})...")
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
        "樞紐合計：Net／Sold 以月化值加總；Stock／Safety 以原始值加總；SKU 數為不重複 SKU。",
    ]
    sheet_rows = [
        ("說明（本頁）", "月化邏輯、本次參數與導讀", "先核對覆蓋月份與月化係數是否符合預期"),
        ("SKU月度變化", "每組相鄰月份環比事件（上月／本月值＋變化量＋變化率）", "篩選 Safety 增加／減少，查看同期 Sold 與 Net 變化"),
        ("By Adj. MCH1", "各 MCH1 每月合計（月化基準）", "比較不同品類每月生意變化"),
        ("By Adj. SST", "各 SST 每月合計（月化基準）", "比較不同供應類型每月生意變化"),
        ("By MCH1×SST", "MCH1 與 SST 組合每月合計", "細分品類 × 供應類型"),
        ("By Shop", "每間店舖每月合計＋環比變化", "多店舖檔案時比較各店走勢"),
        ("月度摘要", "指標卡＋每月合計", "管理層概覽"),
        ("RAW", "原始明細＋YEAR-MONTH 與月化估算欄", "核對逐列數值"),
    ]
    tips = [
        "最新月份為非足月時，Sold 與 Net 的環比已按月化基準計算；原始 MTD 值仍保留於主表以供核對。",
        "數據截止日與最新月份不一致時，最新月份不作月化（係數為 1），並列出警告。",
        "上月缺資料的組合不計變化，屬新增組合；變化率分母為 0 時留空。",
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
    """Safety Stock 月度變化分析輸出八個工作表（說明預設開啟）。"""
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
    _write_safm_table_sheet(ws_mch1, mch1_pivot, progress_callback, 0.4, 0.5)
    ws_sst = wb.create_sheet(title="By Adj. SST")
    _write_safm_table_sheet(ws_sst, sst_pivot, progress_callback, 0.5, 0.6)
    ws_cross = wb.create_sheet(title="By MCH1×SST")
    _write_safm_table_sheet(ws_cross, mch1_sst_pivot, progress_callback, 0.6, 0.7)
    ws_shop = wb.create_sheet(title="By Shop")
    _write_safm_table_sheet(ws_shop, shop_pivot, progress_callback, 0.7, 0.78)

    if progress_callback:
        progress_callback(0.8, "建立月度摘要...")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    desc_font = Font(size=9, color="6B7280", italic=True)
    ws_dash = wb.create_sheet(title="月度摘要")
    row1, row2 = safety_monthly_dashboard_metrics(change_df, stats)
    _write_metric_cards(
        ws_dash, "Safety Stock 月度變化摘要 (Monthly Safety Review Dashboard)",
        row1, row2, thin_border, desc_font, progress_callback,
        progress_pct=0.85,
    )
    ym_list = (stats or {}).get("ym_list") or []
    if ym_list and shop_pivot is not None and not shop_pivot.empty:
        start_row = 12
        ws_dash.cell(row=start_row, column=1, value="每月合計（全店舖月化基準）").font = Font(bold=True, size=12, color="1E40AF")
        headers = ["YEAR-MONTH", "SKU數", "Safety合計", "Sold月化合計", "Net月化合計", "Stock合計"]
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
        order = {label: i for i, label in enumerate(ym_list)}
        agg["_o"] = agg[_SAFM_YM_COLUMN].map(order).fillna(9999)
        agg = agg.sort_values("_o").drop(columns=["_o"])
        for r_idx, (_, r) in enumerate(agg.iterrows(), start_row + 2):
            for col_idx, h in enumerate(headers, 1):
                v = r.get(h)
                cell = ws_dash.cell(row=r_idx, column=col_idx, value=None if pd.isna(v) else v)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    cell.number_format = "#,##0.00" if "Net" in h else "#,##0"
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
