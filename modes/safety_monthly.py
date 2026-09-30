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
僅套用於最新 ``YEAR-MONTH`` 的 Sold Qty 與 Net Sales，舊月份係數為 1。

輸出九個工作表：
1. 說明 — 本次執行參數、核心邏輯與工作表導讀（預設開啟頁）
2. SKU月度變化 — 每組 ``Site Code × SKU`` 相鄰月份環比事件（上月缺資料者標示新增組合）
3. By Adj. MCH1 — 三段式樞紐，四項指標區塊左右並排（每月合計／月環比%／月環比）＋RF→ND／ND→RF SKU數（分月份）
4. By Adj. SST — 各 SST 三段式樞紐（版面同 By Adj. MCH1）
5. By MCH1×SST — MCH1 與 SST 組合三段式樞紐（版面同 By Adj. MCH1）
6. By Shop — 每間店舖每月合計＋環比變化＋RF→ND／ND→RF SKU數（當月）
7. RP Type 變化 — RP Type 轉變 × 月份配對統計＋首尾月份對比明細（含最早／最新月份 Sold Qty 月化、Net Sales 月化與 Safety Stock Qty 對比）
8. 月度摘要 — 指標卡＋每月合計表
9. RAW — 原始明細＋YEAR-MONTH 與月化估算欄
"""

import calendar
import datetime
import gc
import unicodedata
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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
    "_SAFM_COHORT_COLUMNS",
    "_SAFM_SHOP_EFFECT_COLUMNS",
    "_SAFM_T0_COLUMNS",
    "_SAFM_YM_COLUMN",
    "_SAFM_SOLD_ADJ_COLUMN",
    "_SAFM_NET_ADJ_COLUMN",
    "_SAFM_REQUIRED",
    "_SAFM_PIVOT_METRICS",
    "_SAFM_RP_ORDER",
    "_SAFM_RP_ENDPOINT_COLUMNS",
    "_build_safm_rp_summary",
    "_build_safm_rp_endpoints",
    "_build_safm_rp_cohort",
    "_build_safm_shop_effect",
    "_build_safm_t0_track",
]

# 樞紐四項指標（顯示全稱，不用簡稱）：
# Sold Qty＝Sold Qty (CM - MTD) 月化合計；Net Sales＝Net Sales (CM - MTD) 月化合計；
# Net Stock Qty＝店舖存貨 Shop Stock Qty (CM) 合計（時點快照）；Safety Stock Qty＝Adj. Safety Stock 合計。
_SAFM_PIVOT_METRICS = ("Sold Qty", "Net Sales", "Net Stock Qty", "Safety Stock Qty")

# 樞紐與 By Shop 附加的 RP 轉變 SKU 統計（分月份不重複 SKU 數，按轉變的「本月」歸月；
# 樞紐區塊的「全部合計」為整個期間不重複 SKU 數）
_SAFM_RP_SKU_COUNT_METRICS = ("RF→ND SKU數", "ND→RF SKU數")
_SAFM_PIVOT_ALL_METRICS = _SAFM_PIVOT_METRICS + _SAFM_RP_SKU_COUNT_METRICS

# RP Type 轉變固定顯示順序（RF→ND 與 ND→RF 為重點觀察對象，置於後方兩列）
_SAFM_RP_ORDER = ("RF→RF", "ND→ND", "RF→ND", "ND→RF")

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
    "Safety Stock Qty 變化量", "Safety Stock Qty 變化分類",
    "上月 RP Type", "本月 RP Type", "RP Type 轉變",
    "上月 Sold Qty (CM-MTD)", "本月 Sold Qty (CM-MTD)",
    "上月 Sold Qty 月化", "本月 Sold Qty 月化",
    "Sold Qty 變化量(月化基準)", "Sold Qty 變化率",
    "上月 Net Sales (CM-MTD)", "本月 Net Sales (CM-MTD)",
    "上月 Net Sales 月化", "本月 Net Sales 月化",
    "Net Sales 變化量(月化基準)", "Net Sales 變化率",
    "上月 Shop Stock Qty (CM)", "本月 Shop Stock Qty (CM)",
    "Net Stock Qty 變化量", "Net Stock Qty 變化率",
]

_SAFM_SHOP_COLUMNS = [
    "Site Code", "Shop Name", "YEAR-MONTH",
    "SKU數", "Safety Stock Qty 合計",
    "Sold Qty 原始合計", "Sold Qty 月化合計",
    "Net Sales 原始合計", "Net Sales 月化合計",
    "Net Stock Qty 合計",
    "Sold Qty 環比變化", "Net Sales 環比變化", "Safety Stock Qty 環比變化",
    "RF→ND SKU數", "ND→RF SKU數",
]

# ── RP 轉變群組生意對比：按 RP Type 轉變 × 本月分組，隔離轉變群組的生意變化 ──
# 群組口徑為 SKU月度變化事件（Site Code × SKU 相鄰月份環比），避免與持續群組混合。
_SAFM_COHORT_COLUMNS = [
    "RP Type 轉變", "本月",
    "事件數", "SKU數",
    "上月 Sold Qty 月化合計", "本月 Sold Qty 月化合計",
    "Sold Qty 變化量(月化基準)", "Sold Qty 變化率",
    "上月 Net Sales 月化合計", "本月 Net Sales 月化合計",
    "Net Sales 變化量(月化基準)", "Net Sales 變化率",
    "上月 Net Stock Qty 合計", "本月 Net Stock Qty 合計", "Net Stock Qty 變化量",
    "上月 Safety Stock Qty 合計", "本月 Safety Stock Qty 合計", "Safety Stock Qty 變化量",
]

# ── 店舖精簡成效：按店舖 × 本月，驗證減少 SKU 數量後留存 SKU 生意是否提升 ──
# 留存 RF 組＝同一店舖同一 SKU 上月與本月皆為 RF（RF→RF 事件）；
# RF→ND 組上月生意＝被精簡掉的生意基數；全店變化以月度彙總為準（含新增與消失組合）。
_SAFM_SHOP_EFFECT_COLUMNS = [
    "Site Code", "Shop Name", "本月",
    "RF SKU數", "ND SKU數", "總SKU數",
    "RF SKU數環比變化", "總SKU數環比變化",
    "RF→ND SKU數", "ND→RF SKU數", "淨精簡SKU數",
    "RF→ND組上月 Sold 合計(月化)",
    "留存RF組 SKU數",
    "留存RF組上月 Sold 合計(月化)", "留存RF組本月 Sold 合計(月化)",
    "留存RF組 Sold 變化量", "留存RF組 Sold 變化率",
    "留存RF組上月 Net 合計(月化)", "留存RF組本月 Net 合計(月化)",
    "留存RF組 Net 變化量", "留存RF組 Net 變化率",
    "全店 Sold 變化量(月化基準)", "全店 Net 變化量(月化基準)",
]

# ── 轉變前後追蹤（T0 對齊）：以轉變發生的本月為 T0，追蹤 T-1／T0／T+1 ──
# ND 代表不執行自動補貨，賣出／調走／退貨後不再銷售，故 RF→ND 後生意與存貨應逐步去化；
# T+1 取資料內下一個 YEAR-MONTH（非曆月推算），缺資料者標示未有次月資料。
_SAFM_T0_COLUMNS = [
    "轉變類型", "Site Code", "Shop Name", "SKU", "SKU Desc", "Brand", "Buyer",
    "Adj. MCH1", "Adj. SST",
    "T-1月份", "T0月份", "T+1月份",
    "T-1 RP Type", "T0 RP Type", "T+1 RP Type",
    "T-1 Sold Qty 月化", "T0 Sold Qty 月化", "T+1 Sold Qty 月化",
    "T0環比變化量", "T+1環比變化量",
    "T-1 Net Sales 月化", "T0 Net Sales 月化", "T+1 Net Sales 月化",
    "T-1 Net Stock Qty", "T0 Net Stock Qty", "T+1 Net Stock Qty",
    "T-1 Safety Stock Qty", "T0 Safety Stock Qty", "T+1 Safety Stock Qty",
    "去化狀態",
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
            "同一主鍵多列時，Sold Qty／Net Sales／Net Stock Qty 以加總計，"
            "Safety Stock Qty／RP Type／Adj. MCH1／Adj. SST 取首個非空值。"
        )

    # P0：文字欄「首個非空值」改以空字串遮成 NaN＋內建 first 聚合（Cython 路徑）。
    # 自訂 Python 函數會令 pandas 放棄快速路徑、逐組逐欄呼叫，為轉換階段主要瓶頸；
    # 語意等價：first 略過 NaN＝首個非空字串，全空組聚合後回填空字串。
    _safm_text_cols = ["_brand", "_buyer", "_mch1", "_sst", "_shop", "_desc"]
    agg_df = tmp.copy()
    for _c in _safm_text_cols:
        _blank = agg_df[_c].str.strip() == ""
        if _blank.any():
            agg_df.loc[_blank, _c] = None

    monthly = agg_df.groupby(["_site", "_sku", "_ym"], sort=True).agg(
        safety=("_safety", "first"),
        sold_raw=("_sold_raw", "sum"),
        sold_adj=("_sold_adj", "sum"),
        net_raw=("_net_raw", "sum"),
        net_adj=("_net_adj", "sum"),
        stock=("_stock", "sum"),
        rp=("_rp", "first"),
        brand=("_brand", "first"),
        buyer=("_buyer", "first"),
        mch1=("_mch1", "first"),
        sst=("_sst", "first"),
        shop=("_shop", "first"),
        desc=("_desc", "first"),
    ).reset_index()
    monthly[["brand", "buyer", "mch1", "sst", "shop", "desc"]] = (
        monthly[["brand", "buyer", "mch1", "sst", "shop", "desc"]].fillna("")
    )
    monthly = monthly.sort_values(["_site", "_sku", "_ym"]).reset_index(drop=True)

    if progress_callback:
        progress_callback(0.45, "計算相鄰月份環比事件...")
    # P0：monthly 已按 Site Code × SKU × YEAR-MONTH 排序，同一組合的列必然相鄰，
    # 直接以原生 list 相鄰列配對取代 groupby＋逐列 .iloc（每次 .iloc 都會重建 Series）。
    _m_site = monthly["_site"].tolist()
    _m_sku = monthly["_sku"].tolist()
    _m_ym = monthly["_ym"].tolist()
    _m_safety = monthly["safety"].tolist()
    _m_sold_raw = monthly["sold_raw"].tolist()
    _m_sold_adj = monthly["sold_adj"].tolist()
    _m_net_raw = monthly["net_raw"].tolist()
    _m_net_adj = monthly["net_adj"].tolist()
    _m_stock = monthly["stock"].tolist()
    _m_rp = monthly["rp"].tolist()
    _m_brand = monthly["brand"].tolist()
    _m_buyer = monthly["buyer"].tolist()
    _m_mch1 = monthly["mch1"].tolist()
    _m_sst = monthly["sst"].tolist()
    _m_shop = monthly["shop"].tolist()
    _m_desc = monthly["desc"].tolist()

    change_rows: list[dict] = []
    for i in range(1, len(monthly)):
        if _m_site[i] != _m_site[i - 1] or _m_sku[i] != _m_sku[i - 1]:
            continue
        prev_key = int(_m_ym[i - 1])
        curr_key = int(_m_ym[i])
        prev_safety = int(_m_safety[i - 1])
        curr_safety = int(_m_safety[i])
        safety_delta = curr_safety - prev_safety
        if safety_delta > 0:
            safety_cat = "增加"
        elif safety_delta < 0:
            safety_cat = "減少"
        else:
            safety_cat = "不變"
        rp_prev = str(_m_rp[i - 1])
        rp_curr = str(_m_rp[i])
        prev_sold_adj = int(_m_sold_adj[i - 1])
        curr_sold_adj = int(_m_sold_adj[i])
        sold_delta = curr_sold_adj - prev_sold_adj
        # 變化量與變化率沿用未捨入的聚合值（與原實作一致），顯示月化值取 2 位小數
        prev_net_adj = round(float(_m_net_adj[i - 1]), 2)
        curr_net_adj = round(float(_m_net_adj[i]), 2)
        net_delta = round(float(_m_net_adj[i]) - float(_m_net_adj[i - 1]), 2)
        prev_stock = int(_m_stock[i - 1])
        curr_stock = int(_m_stock[i])
        change_rows.append({
            "Site Code": _m_site[i], "Shop Name": _m_shop[i], "SKU": _m_sku[i],
            "SKU Desc": _m_desc[i] or _m_desc[i - 1],
            "Brand": _m_brand[i] or _m_brand[i - 1],
            "Buyer": _m_buyer[i] or _m_buyer[i - 1],
            "Adj. MCH1": _m_mch1[i] or _m_mch1[i - 1],
            "Adj. SST": _m_sst[i] or _m_sst[i - 1],
            "上月": _ym_label(prev_key // 100, prev_key % 100),
            "本月": _ym_label(curr_key // 100, curr_key % 100),
            "是否相鄰曆月": "是" if _is_calendar_adjacent(prev_key, curr_key) else "否",
            "上月 Adj. Safety Stock": prev_safety,
            "本月 Adj. Safety Stock": curr_safety,
            "Safety Stock Qty 變化量": safety_delta, "Safety Stock Qty 變化分類": safety_cat,
            "上月 RP Type": rp_prev, "本月 RP Type": rp_curr,
            "RP Type 轉變": f"{rp_prev}→{rp_curr}",
            "上月 Sold Qty (CM-MTD)": int(_m_sold_raw[i - 1]),
            "本月 Sold Qty (CM-MTD)": int(_m_sold_raw[i]),
            "上月 Sold Qty 月化": prev_sold_adj,
            "本月 Sold Qty 月化": curr_sold_adj,
            "Sold Qty 變化量(月化基準)": sold_delta,
            "Sold Qty 變化率": _change_rate(float(curr_sold_adj), float(prev_sold_adj)),
            "上月 Net Sales (CM-MTD)": round(float(_m_net_raw[i - 1]), 2),
            "本月 Net Sales (CM-MTD)": round(float(_m_net_raw[i]), 2),
            "上月 Net Sales 月化": prev_net_adj,
            "本月 Net Sales 月化": curr_net_adj,
            "Net Sales 變化量(月化基準)": net_delta,
            "Net Sales 變化率": _change_rate(float(_m_net_adj[i]), float(_m_net_adj[i - 1])),
            "上月 Shop Stock Qty (CM)": prev_stock,
            "本月 Shop Stock Qty (CM)": curr_stock,
            "Net Stock Qty 變化量": curr_stock - prev_stock,
            "Net Stock Qty 變化率": _change_rate(float(curr_stock), float(prev_stock)),
        })
    change_df = pd.DataFrame(change_rows, columns=_SAFM_CHANGE_COLUMNS)
    if not change_df.empty:
        change_df = change_df.sort_values(
            ["Brand", "SKU", "Site Code", "本月"], kind="stable"
        ).reset_index(drop=True)

    if progress_callback:
        progress_callback(0.65, "建立維度樞紐...")
    _mch1_counts_raw, _rp_totals = _safm_rp_sku_counts(change_df, ["Adj. MCH1"])
    _sst_counts_raw, _ = _safm_rp_sku_counts(change_df, ["Adj. SST"])
    _cross_counts_raw, _ = _safm_rp_sku_counts(change_df, ["Adj. MCH1", "Adj. SST"])
    _shop_counts, _ = _safm_rp_sku_counts(change_df, ["Site Code"])
    mch1_pivot = _build_safm_metric_pivot(
        monthly, "mch1", "Adj. MCH1", ym_keys,
        rp_counts={_blank_dim(k[0]): v for k, v in _mch1_counts_raw.items()},
        rp_totals=_rp_totals,
    )
    sst_pivot = _build_safm_metric_pivot(
        monthly, "sst", "Adj. SST", ym_keys,
        rp_counts={_blank_dim(k[0]): v for k, v in _sst_counts_raw.items()},
        rp_totals=_rp_totals,
    )
    mch1_sst_pivot = _build_safm_metric_pivot(
        monthly, ["mch1", "sst"], None, ym_keys,
        rp_counts={"｜".join(k): v for k, v in _cross_counts_raw.items()},
        rp_totals=_rp_totals,
    )
    shop_pivot = _build_safm_shop_pivot(monthly, ym_keys, rp_counts=_shop_counts)

    if progress_callback:
        progress_callback(0.8, "整理 RAW 明細...")
    raw_df = df.copy()
    drop_internal = [c for c in raw_df.columns if str(c).startswith("_safm_")]
    raw_df = raw_df.drop(columns=[c for c in drop_internal if c not in ("_safm_ym_key",)], errors="ignore")
    raw_df = raw_df.drop(columns=["_safm_ym_key"], errors="ignore")

    safety_delta_total = int(change_df["Safety Stock Qty 變化量"].sum()) if not change_df.empty else 0
    rp_counts = change_df["RP Type 轉變"].value_counts().to_dict() if not change_df.empty else {}
    if not change_df.empty:
        _rf_to_nd = change_df["RP Type 轉變"] == "RF→ND"
        _nd_to_rf = change_df["RP Type 轉變"] == "ND→RF"
        rp_rf_to_nd = int(_rf_to_nd.sum())
        rp_nd_to_rf = int(_nd_to_rf.sum())
        rp_rf_to_nd_skus = int(change_df.loc[_rf_to_nd, "SKU"].nunique())
        rp_nd_to_rf_skus = int(change_df.loc[_nd_to_rf, "SKU"].nunique())
    else:
        rp_rf_to_nd = rp_nd_to_rf = rp_rf_to_nd_skus = rp_nd_to_rf_skus = 0
    rp_summary = _build_safm_rp_summary(change_df)
    rp_endpoints, rp_end_info = _build_safm_rp_endpoints(monthly, ym_keys)
    rp_cohort = _build_safm_rp_cohort(change_df)
    shop_effect = _build_safm_shop_effect(change_df, monthly, ym_keys, ym_labels)
    t0_track = _build_safm_t0_track(change_df, monthly, ym_keys, ym_labels)
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
        "safety_inc": int((change_df["Safety Stock Qty 變化分類"] == "增加").sum()) if not change_df.empty else 0,
        "safety_dec": int((change_df["Safety Stock Qty 變化分類"] == "減少").sum()) if not change_df.empty else 0,
        "safety_same": int((change_df["Safety Stock Qty 變化分類"] == "不變").sum()) if not change_df.empty else 0,
        "safety_delta_total": safety_delta_total,
        "rp_counts": {str(k): int(v) for k, v in rp_counts.items()},
        "rp_rf_to_nd": rp_rf_to_nd,
        "rp_nd_to_rf": rp_nd_to_rf,
        "rp_rf_to_nd_skus": rp_rf_to_nd_skus,
        "rp_nd_to_rf_skus": rp_nd_to_rf_skus,
        "rp_summary": rp_summary,
        "rp_endpoints": rp_endpoints,
        "rp_end_info": rp_end_info,
        "rp_cohort": rp_cohort,
        "shop_effect": shop_effect,
        "t0_track": t0_track,
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
    monthly: pd.DataFrame, group_keys, header, ym_keys: list[int],
    rp_counts: dict | None = None, rp_totals: dict | None = None,
) -> pd.DataFrame:
    """維度樞紐（三段式：每月合計 → 月環比% → 月環比絕對值，四項指標齊列）。

    每個指標連續三列（項目＝「{指標} 每月合計」「{指標} 月環比%」「{指標} 月環比」），
    欄為維度欄＋項目＋各 YEAR-MONTH＋全部合計（環比兩列的全部合計留空）。
    月環比% 為顯示字串（上月為 0 或首月留空）；其餘為數值。
    表尾另附兩列 RP 轉變 SKU 統計（項目＝「RF→ND SKU數 每月合計」「ND→RF SKU數 每月合計」）：
    各月為該月發生轉變的不重複 SKU 數（按轉變的「本月」歸月），全部合計為整個期間不重複 SKU 數；
    rp_counts 以樞紐的維度值為 key，rp_totals 供總計列使用（全檔口徑）。
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
    _rp_counts = rp_counts or {}
    _rp_totals = rp_totals or {}

    # P0：單次 groupby 聚合取代逐維度×逐月的全表布林掃描（後者接近平方級增長）。
    # 缺組合（某維度某月無資料）以 0 補，與原「空群組回 0」一致。
    _agg = df.groupby(["_dim", "_ym"], sort=False)[
        ["sold_adj", "net_adj", "stock", "safety"]
    ].sum()
    base: dict[tuple[str, int, str], float] = {}
    for (_dim_val, _key), (_sold, _net, _stock, _safety) in zip(
        _agg.index, _agg.itertuples(index=False, name=None)
    ):
        base[(_dim_val, int(_key), "Sold Qty")] = int(_sold)
        base[(_dim_val, int(_key), "Net Sales")] = round(float(_net), 2)
        base[(_dim_val, int(_key), "Net Stock Qty")] = int(_stock)
        base[(_dim_val, int(_key), "Safety Stock Qty")] = int(_safety)

    def _total(dim_val: str | None, key: int, metric: str):
        if dim_val is None:
            vals = [base.get((d, key, metric), 0) for d in dim_vals]
        else:
            vals = [base.get((dim_val, key, metric), 0)]
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
    for rp_metric in _SAFM_RP_SKU_COUNT_METRICS:
        for dim_val in dim_vals + ["總計"]:
            row = _safm_dim_cells(dim_val, dim_cols)
            row["項目"] = f"{rp_metric} 每月合計"
            src = _rp_totals if dim_val == "總計" else _rp_counts.get(dim_val) or {}
            _monthly_src = src.get("monthly") or {}
            for label in ym_labels:
                row[label] = int(_monthly_src.get(label, {}).get(rp_metric, 0))
            row["全部合計"] = int((src.get("total") or {}).get(rp_metric, 0))
            rows.append(row)
    cols = dim_cols + ["項目"] + ym_labels + ["全部合計"]
    return pd.DataFrame(rows, columns=cols)


def _build_safm_shop_pivot(monthly: pd.DataFrame, ym_keys: list[int],
                           rp_counts: dict | None = None) -> pd.DataFrame:
    """By Shop：每間店舖每月合計＋環比變化（月化基準）。

    表尾另附 RF→ND SKU數／ND→RF SKU數：該店舖「當月」發生轉變的不重複 SKU 數
    （按轉變的「本月」歸月；rp_counts 由 _safm_rp_sku_counts 提供）。
    """
    if monthly is None or monthly.empty:
        return pd.DataFrame(columns=_SAFM_SHOP_COLUMNS)
    df = monthly.copy()
    _rp_counts = rp_counts or {}
    ym_labels = {k: _ym_label(k // 100, k % 100) for k in ym_keys}
    rows = []
    for site, g_site in df.groupby("_site", sort=True):
        shop_name = ""
        _names = g_site["shop"].fillna("").astype(str)
        _names = _names[_names.str.strip() != ""]
        if len(_names):
            shop_name = str(_names.iloc[0])
        _shop_monthly = (_rp_counts.get((site,)) or {}).get("monthly") or {}
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
                "SKU數": sku_n, "Safety Stock Qty 合計": safety_sum,
                "Sold Qty 原始合計": sold_raw, "Sold Qty 月化合計": sold_adj,
                "Net Sales 原始合計": net_raw, "Net Sales 月化合計": net_adj,
                "Net Stock Qty 合計": stock_sum,
                "Sold Qty 環比變化": sold_chg, "Net Sales 環比變化": net_chg,
                "Safety Stock Qty 環比變化": safety_chg,
                "RF→ND SKU數": int(_shop_monthly.get(ym_labels[key], {}).get("RF→ND SKU數", 0)),
                "ND→RF SKU數": int(_shop_monthly.get(ym_labels[key], {}).get("ND→RF SKU數", 0)),
            })
            prev_sold, prev_net, prev_safety = sold_adj, net_adj, safety_sum
    return pd.DataFrame(rows, columns=_SAFM_SHOP_COLUMNS)


def _safm_rp_sku_counts(change_df: pd.DataFrame, keys: list[str]) -> tuple[dict, dict]:
    """按維度統計 RP 轉變的不重複 SKU 數（RF→ND／ND→RF），分月份＋整個期間。

    keys 為 change_df 的分組欄（如 ["Adj. MCH1"] 或 ["Adj. MCH1", "Adj. SST"]）。
    回傳 (每維度計數, 全檔總計)，兩者結構相同：
    {"monthly": {"YYYY-MM": {"RF→ND SKU數": n, "ND→RF SKU數": m}, ...},
     "total": {"RF→ND SKU數": n, "ND→RF SKU數": m}}
    monthly 按轉變的「本月」歸月；total 為整個期間不重複 SKU 數（非各月相加，
    同一 SKU 可在不同月份多次轉變；全檔總計亦非各維度相加）。
    """
    empty_total = {"RF→ND SKU數": 0, "ND→RF SKU數": 0}
    if change_df is None or change_df.empty or not all(k in change_df.columns for k in keys):
        return {}, {"monthly": {}, "total": dict(empty_total)}

    def _new_entry() -> dict:
        return {"monthly": {}, "total": dict(empty_total)}

    by_dim: dict[tuple, dict] = {}
    totals = _new_entry()
    for trans, label in (("RF→ND", "RF→ND SKU數"), ("ND→RF", "ND→RF SKU數")):
        sub = change_df.loc[change_df["RP Type 轉變"] == trans]
        if sub.empty:
            continue
        for gkey, n in sub.groupby([*keys, "本月"])["SKU"].nunique().items():
            gkey = gkey if isinstance(gkey, tuple) else (gkey,)
            dims, ym = gkey[:-1], str(gkey[-1])
            by_dim.setdefault(dims, _new_entry())["monthly"].setdefault(ym, {})[label] = int(n)
        for ym, n in sub.groupby("本月")["SKU"].nunique().items():
            totals["monthly"].setdefault(str(ym), {})[label] = int(n)
        for gkey, n in sub.groupby(keys)["SKU"].nunique().items():
            dims = gkey if isinstance(gkey, tuple) else (gkey,)
            by_dim.setdefault(dims, _new_entry())["total"][label] = int(n)
        totals["total"][label] = int(sub["SKU"].nunique())
    return by_dim, totals


def _build_safm_rp_summary(change_df: pd.DataFrame) -> pd.DataFrame:
    """RP Type 轉變 × 月份配對統計（事件數／SKU數，含總計列）。

    列為四類轉變（RF→RF、ND→ND、RF→ND、ND→RF，未知值附加於後），
    欄為各「上月→本月」配對的事件數／SKU數＋總計；總計列 SKU 數為不重複 SKU。
    """
    if change_df is None or change_df.empty:
        return pd.DataFrame(columns=["RP Type 轉變"])
    df = change_df.copy()
    df["_pair"] = df["上月"].fillna("").astype(str) + "→" + df["本月"].fillna("").astype(str)
    pairs = sorted(df["_pair"].unique().tolist())
    present = df["RP Type 轉變"].fillna("").astype(str).tolist()
    order = [t for t in _SAFM_RP_ORDER if t in set(present)]
    order += sorted(set(present) - set(order) - {""})
    rows = []
    for trans in order:
        g = df[df["RP Type 轉變"] == trans]
        row = {"RP Type 轉變": trans}
        for p in pairs:
            gg = g[g["_pair"] == p]
            row[f"{p} 事件數"] = int(len(gg))
            row[f"{p} SKU數"] = int(gg["SKU"].nunique()) if not gg.empty else 0
        row["總計 事件數"] = int(len(g))
        row["總計 SKU數"] = int(g["SKU"].nunique())
        rows.append(row)
    total = {"RP Type 轉變": "總計"}
    for p in pairs:
        gg = df[df["_pair"] == p]
        total[f"{p} 事件數"] = int(len(gg))
        total[f"{p} SKU數"] = int(gg["SKU"].nunique()) if not gg.empty else 0
    total["總計 事件數"] = int(len(df))
    total["總計 SKU數"] = int(df["SKU"].nunique())
    rows.append(total)
    cols = ["RP Type 轉變"]
    for p in pairs:
        cols += [f"{p} 事件數", f"{p} SKU數"]
    cols += ["總計 事件數", "總計 SKU數"]
    return pd.DataFrame(rows, columns=cols)


_SAFM_RP_ENDPOINT_COLUMNS = [
    "Site Code", "Shop Name", "SKU", "SKU Desc", "Brand", "Buyer",
    "Adj. MCH1", "Adj. SST",
    "最早月份", "最早 RP Type", "最新月份", "最新 RP Type",
    "首尾轉變", "是否轉變",
    "最早 Sold Qty 月化", "最新 Sold Qty 月化", "Sold Qty 變化量(月化基準)", "Sold Qty 變化率",
    "最早 Net Sales 月化", "最新 Net Sales 月化", "Net Sales 變化量(月化基準)", "Net Sales 變化率",
    "最早 Safety Stock Qty", "最新 Safety Stock Qty", "Safety Stock Qty 變化量",
]


def _build_safm_rp_endpoints(monthly: pd.DataFrame, ym_keys: list[int]
                             ) -> tuple[pd.DataFrame, dict]:
    """首尾月份 RP 對比明細：每組取最早與最新月份各一列比較。

    只收錄最早與最新月份皆有資料的組合；缺任一者不納入並計入 excluded。
    另附兩個月份的生意與 Safety 對比（Sold Qty／Net Sales 以月化基準，
    Safety Stock Qty 為原始值；變化率＝變化量 ÷ 最早月份絕對值，最早為 0 留空）。
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
    first_m = monthly[monthly["_ym"] == first_key]
    last_m = monthly[monthly["_ym"] == last_key]

    def _pick(latest_val, earliest_val) -> str:
        latest_text = "" if pd.isna(latest_val) else str(latest_val).strip()
        if latest_text:
            return latest_text
        return "" if pd.isna(earliest_val) else str(earliest_val).strip()

    # P0：以鍵值直查（dict）取代逐組 Series.loc；monthly 主鍵唯一，
    # 逐組順序與原 index.intersection 相同（按首月列序）。
    _end_cols = ["shop", "desc", "brand", "buyer", "mch1", "sst", "rp", "sold_adj", "net_adj", "safety"]
    _last = last_m.reset_index(drop=True)
    _last_site = _last["_site"].tolist()
    _last_sku = _last["_sku"].tolist()
    _last_pos: dict[tuple, int] = {}
    for _j in range(len(_last)):
        _last_pos[(_last_site[_j], _last_sku[_j])] = _j
    _last_val = {c: _last[c].tolist() for c in _end_cols}
    _first_site = first_m["_site"].tolist()
    _first_sku = first_m["_sku"].tolist()
    _first_val = {c: first_m[c].tolist() for c in _end_cols}

    rows = []
    for _i in range(len(first_m)):
        _key = (_first_site[_i], _first_sku[_i])
        _j = _last_pos.get(_key)
        if _j is None:
            continue
        site, sku = _key
        rp_prev, rp_curr = str(_first_val["rp"][_i]), str(_last_val["rp"][_j])
        trans = f"{rp_prev}→{rp_curr}"
        sold_prev, sold_curr = int(_first_val["sold_adj"][_i]), int(_last_val["sold_adj"][_j])
        net_prev = round(float(_first_val["net_adj"][_i]), 2)
        net_curr = round(float(_last_val["net_adj"][_j]), 2)
        safety_prev, safety_curr = int(_first_val["safety"][_i]), int(_last_val["safety"][_j])
        rows.append({
            "Site Code": site,
            "Shop Name": _pick(_last_val["shop"][_j], _first_val["shop"][_i]),
            "SKU": sku,
            "SKU Desc": _pick(_last_val["desc"][_j], _first_val["desc"][_i]),
            "Brand": _pick(_last_val["brand"][_j], _first_val["brand"][_i]),
            "Buyer": _pick(_last_val["buyer"][_j], _first_val["buyer"][_i]),
            "Adj. MCH1": _pick(_last_val["mch1"][_j], _first_val["mch1"][_i]),
            "Adj. SST": _pick(_last_val["sst"][_j], _first_val["sst"][_i]),
            "最早月份": first_label, "最早 RP Type": rp_prev,
            "最新月份": last_label, "最新 RP Type": rp_curr,
            "首尾轉變": trans, "是否轉變": "是" if rp_prev != rp_curr else "否",
            "最早 Sold Qty 月化": sold_prev, "最新 Sold Qty 月化": sold_curr,
            "Sold Qty 變化量(月化基準)": sold_curr - sold_prev,
            "Sold Qty 變化率": _change_rate(float(sold_curr), float(sold_prev)),
            "最早 Net Sales 月化": net_prev, "最新 Net Sales 月化": net_curr,
            "Net Sales 變化量(月化基準)": round(net_curr - net_prev, 2),
            "Net Sales 變化率": _change_rate(net_curr, net_prev),
            "最早 Safety Stock Qty": safety_prev, "最新 Safety Stock Qty": safety_curr,
            "Safety Stock Qty 變化量": safety_curr - safety_prev,
        })
    excluded = int(monthly[["_site", "_sku"]].drop_duplicates().shape[0] - len(rows))
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


def _safm_label_to_key(label: str) -> int | None:
    """將 YEAR-MONTH 標籤（如 2026-06）轉為排序鍵 YYYYMM，無法解析時回傳 None。"""
    try:
        parts = str(label).strip().split("-")
        if len(parts) != 2:
            return None
        return int(parts[0]) * 100 + int(parts[1])
    except (TypeError, ValueError):
        return None


def _build_safm_rp_cohort(change_df: pd.DataFrame) -> pd.DataFrame:
    """RP 轉變群組生意對比：按 RP Type 轉變 × 本月分組，隔離各群組的生意變化。

    母體為 SKU月度變化事件（Site Code × SKU 相鄰月份環比），每列一組轉變 × 本月。
    另附每種轉變的全部合計列、每月總計列與全檔總計列，便於對照
    RF→ND 群組與持續 RF 群組的同期生意走勢。
    變化率＝變化量 ÷ 上月合計絕對值，上月合計為 0 時留空。
    """
    if change_df is None or change_df.empty:
        return pd.DataFrame(columns=_SAFM_COHORT_COLUMNS)
    df = change_df.copy()
    if "RP Type 轉變" not in df.columns or "本月" not in df.columns:
        return pd.DataFrame(columns=_SAFM_COHORT_COLUMNS)

    def _agg(group: pd.DataFrame) -> dict:
        prev_sold = int(pd.to_numeric(group["上月 Sold Qty 月化"], errors="coerce").fillna(0).sum())
        curr_sold = int(pd.to_numeric(group["本月 Sold Qty 月化"], errors="coerce").fillna(0).sum())
        prev_net = round(float(pd.to_numeric(group["上月 Net Sales 月化"], errors="coerce").fillna(0).sum()), 2)
        curr_net = round(float(pd.to_numeric(group["本月 Net Sales 月化"], errors="coerce").fillna(0).sum()), 2)
        prev_stock = int(pd.to_numeric(group["上月 Shop Stock Qty (CM)"], errors="coerce").fillna(0).sum())
        curr_stock = int(pd.to_numeric(group["本月 Shop Stock Qty (CM)"], errors="coerce").fillna(0).sum())
        prev_safety = int(pd.to_numeric(group["上月 Adj. Safety Stock"], errors="coerce").fillna(0).sum())
        curr_safety = int(pd.to_numeric(group["本月 Adj. Safety Stock"], errors="coerce").fillna(0).sum())
        return {
            "事件數": int(len(group)),
            "SKU數": int(group["SKU"].nunique()) if "SKU" in group.columns else int(len(group)),
            "上月 Sold Qty 月化合計": prev_sold,
            "本月 Sold Qty 月化合計": curr_sold,
            "Sold Qty 變化量(月化基準)": int(curr_sold - prev_sold),
            "Sold Qty 變化率": _change_rate(float(curr_sold), float(prev_sold)),
            "上月 Net Sales 月化合計": prev_net,
            "本月 Net Sales 月化合計": curr_net,
            "Net Sales 變化量(月化基準)": round(curr_net - prev_net, 2),
            "Net Sales 變化率": _change_rate(float(curr_net), float(prev_net)),
            "上月 Net Stock Qty 合計": prev_stock,
            "本月 Net Stock Qty 合計": curr_stock,
            "Net Stock Qty 變化量": int(curr_stock - prev_stock),
            "上月 Safety Stock Qty 合計": prev_safety,
            "本月 Safety Stock Qty 合計": curr_safety,
            "Safety Stock Qty 變化量": int(curr_safety - prev_safety),
        }

    months = sorted(df["本月"].fillna("").astype(str).unique().tolist())
    present = df["RP Type 轉變"].fillna("").astype(str)
    order = [t for t in _SAFM_RP_ORDER if t in set(present.tolist())]
    order += sorted(set(present.tolist()) - set(order) - {""})
    if "" in set(present.tolist()):
        order.append("")

    rows: list[dict] = []
    for trans in order:
        sub = df[df["RP Type 轉變"].fillna("").astype(str) == trans]
        if sub.empty:
            continue
        for ym in months:
            g = sub[sub["本月"].fillna("").astype(str) == ym]
            if g.empty:
                continue
            row = {"RP Type 轉變": trans, "本月": ym}
            row.update(_agg(g))
            rows.append(row)
        total = {"RP Type 轉變": trans, "本月": "全部合計"}
        total.update(_agg(sub))
        rows.append(total)
    for ym in months:
        g = df[df["本月"].fillna("").astype(str) == ym]
        if g.empty:
            continue
        row = {"RP Type 轉變": "總計", "本月": ym}
        row.update(_agg(g))
        rows.append(row)
    grand = {"RP Type 轉變": "總計", "本月": "全部合計"}
    grand.update(_agg(df))
    rows.append(grand)

    result = pd.DataFrame(rows, columns=_SAFM_COHORT_COLUMNS)
    order_idx = {t: i for i, t in enumerate(order + ["總計"])}
    month_idx = {m: i for i, m in enumerate(months + ["全部合計"])}
    result["_o1"] = result["RP Type 轉變"].map(order_idx).fillna(99).astype(int)
    result["_o2"] = result["本月"].map(month_idx).fillna(99).astype(int)
    result = result.sort_values(["_o1", "_o2"], kind="stable").drop(
        columns=["_o1", "_o2"]).reset_index(drop=True)
    return result


def _build_safm_shop_effect(
    change_df: pd.DataFrame | None,
    monthly: pd.DataFrame | None,
    ym_keys: list[int] | None = None,
    ym_labels: list[str] | None = None,
) -> pd.DataFrame:
    """店舖精簡成效：按店舖 × 本月，驗證減少 SKU 數量後留存 SKU 生意是否提升。

    每列一間店舖一個轉變月份，包含當月 RF／ND／總 SKU 數及其環比變化、
    RF→ND／ND→RF SKU 數與淨精簡數、RF→ND 組被精簡掉的上月生意基數、
    留存 RF 組（RF→RF 事件）上月與本月生意對比，以及全店生意變化。
    全店變化以月度彙總為準（含新增與消失組合），留存組只計可比事件，
    兩者並列即可判斷精簡後其餘 SKU 是否承接生意。
    """
    if monthly is None or monthly.empty:
        return pd.DataFrame(columns=_SAFM_SHOP_EFFECT_COLUMNS)
    ym_keys = list(ym_keys or sorted(monthly["_ym"].unique().tolist()))
    if not ym_keys:
        return pd.DataFrame(columns=_SAFM_SHOP_EFFECT_COLUMNS)
    if not ym_labels:
        ym_labels = [_ym_label(k // 100, k % 100) for k in ym_keys]
    key_to_label = {int(k): str(l) for k, l in zip(ym_keys, ym_labels)}
    label_to_key = {str(l): int(k) for k, l in zip(ym_keys, ym_labels)}

    df_m = monthly.copy()
    sites = sorted(df_m["_site"].fillna("").astype(str).unique().tolist())

    totals: dict[tuple, dict] = {}
    for (site, key), g in df_m.groupby(["_site", "_ym"], sort=False):
        totals[(str(site), int(key))] = {
            "rf": int(((g["rp"] == "RF")).sum()),
            "nd": int(((g["rp"] == "ND")).sum()),
            "total": int(g["_sku"].nunique()),
            "sold": int(pd.to_numeric(g["sold_adj"], errors="coerce").fillna(0).sum()),
            "net": round(float(pd.to_numeric(g["net_adj"], errors="coerce").fillna(0).sum()), 2),
            "shop": str(g["shop"].fillna("").astype(str).loc[g["shop"].fillna("").astype(str).str.strip() != ""].iloc[0])
            if (g["shop"].fillna("").astype(str).str.strip() != "").any() else "",
        }

    by_site_month: dict[tuple, dict] = {}
    if change_df is not None and not change_df.empty and "本月" in change_df.columns:
        cdf = change_df.copy()
        cdf["_site_s"] = cdf["Site Code"].fillna("").astype(str)
        for (site, ym_label), g in cdf.groupby(["_site_s", "本月"], sort=False):
            g_rf_nd = g[g["RP Type 轉變"] == "RF→ND"]
            g_nd_rf = g[g["RP Type 轉變"] == "ND→RF"]
            g_keep = g[g["RP Type 轉變"] == "RF→RF"]
            keep_sold_prev = int(pd.to_numeric(g_keep["上月 Sold Qty 月化"], errors="coerce").fillna(0).sum()) if not g_keep.empty else 0
            keep_sold_curr = int(pd.to_numeric(g_keep["本月 Sold Qty 月化"], errors="coerce").fillna(0).sum()) if not g_keep.empty else 0
            keep_net_prev = round(float(pd.to_numeric(g_keep["上月 Net Sales 月化"], errors="coerce").fillna(0).sum()), 2) if not g_keep.empty else 0.0
            keep_net_curr = round(float(pd.to_numeric(g_keep["本月 Net Sales 月化"], errors="coerce").fillna(0).sum()), 2) if not g_keep.empty else 0.0
            by_site_month[(str(site), str(ym_label))] = {
                "rf_to_nd_skus": int(g_rf_nd["SKU"].nunique()) if not g_rf_nd.empty and "SKU" in g_rf_nd.columns else int(len(g_rf_nd)),
                "nd_to_rf_skus": int(g_nd_rf["SKU"].nunique()) if not g_nd_rf.empty and "SKU" in g_nd_rf.columns else int(len(g_nd_rf)),
                "removed_sold": int(pd.to_numeric(g_rf_nd["上月 Sold Qty 月化"], errors="coerce").fillna(0).sum()) if not g_rf_nd.empty else 0,
                "keep_skus": int(g_keep["SKU"].nunique()) if not g_keep.empty and "SKU" in g_keep.columns else int(len(g_keep)),
                "keep_sold_prev": keep_sold_prev,
                "keep_sold_curr": keep_sold_curr,
                "keep_sold_chg": int(keep_sold_curr - keep_sold_prev),
                "keep_sold_rate": _change_rate(float(keep_sold_curr), float(keep_sold_prev)) if not g_keep.empty else None,
                "keep_net_prev": keep_net_prev,
                "keep_net_curr": keep_net_curr,
                "keep_net_chg": round(keep_net_curr - keep_net_prev, 2),
                "keep_net_rate": _change_rate(float(keep_net_curr), float(keep_net_prev)) if not g_keep.empty else None,
            }

    rows: list[dict] = []
    for site in sites:
        prev_rf = prev_total = prev_sold = prev_net = None
        shop_name = ""
        for key in ym_keys:
            label = key_to_label.get(int(key), "")
            t = totals.get((site, int(key)), {"rf": 0, "nd": 0, "total": 0, "sold": 0, "net": 0.0, "shop": ""})
            if t.get("shop"):
                shop_name = t["shop"]
            info = by_site_month.get((site, label), {})
            rf_to_nd = int(info.get("rf_to_nd_skus", 0))
            nd_to_rf = int(info.get("nd_to_rf_skus", 0))
            if prev_rf is None:
                rf_chg = total_chg = sold_chg = net_chg = None
            else:
                rf_chg = int(t["rf"] - prev_rf)
                total_chg = int(t["total"] - prev_total)
                sold_chg = int(t["sold"] - prev_sold)
                net_chg = round(float(t["net"]) - float(prev_net), 2)
            # 首月無上月可比，轉變計數恆為 0，避免將期初存量誤讀為當月精簡
            if prev_rf is None:
                rf_to_nd = 0
                nd_to_rf = 0
                removed = 0
                keep = {"keep_skus": 0, "keep_sold_prev": 0, "keep_sold_curr": 0,
                        "keep_sold_chg": None, "keep_sold_rate": None,
                        "keep_net_prev": 0.0, "keep_net_curr": 0.0,
                        "keep_net_chg": None, "keep_net_rate": None}
            else:
                removed = int(info.get("removed_sold", 0))
                keep = {
                    "keep_skus": int(info.get("keep_skus", 0)),
                    "keep_sold_prev": int(info.get("keep_sold_prev", 0)),
                    "keep_sold_curr": int(info.get("keep_sold_curr", 0)),
                    "keep_sold_chg": int(info.get("keep_sold_chg", 0)) if info else 0,
                    "keep_sold_rate": info.get("keep_sold_rate"),
                    "keep_net_prev": round(float(info.get("keep_net_prev", 0.0)), 2),
                    "keep_net_curr": round(float(info.get("keep_net_curr", 0.0)), 2),
                    "keep_net_chg": round(float(info.get("keep_net_chg", 0.0)), 2) if info else 0.0,
                    "keep_net_rate": info.get("keep_net_rate"),
                }
                if not info:
                    keep["keep_sold_chg"] = 0
                    keep["keep_net_chg"] = 0.0
            rows.append({
                "Site Code": site, "Shop Name": shop_name or site, "本月": label,
                "RF SKU數": int(t["rf"]), "ND SKU數": int(t["nd"]), "總SKU數": int(t["total"]),
                "RF SKU數環比變化": rf_chg, "總SKU數環比變化": total_chg,
                "RF→ND SKU數": rf_to_nd, "ND→RF SKU數": nd_to_rf,
                "淨精簡SKU數": int(rf_to_nd - nd_to_rf),
                "RF→ND組上月 Sold 合計(月化)": removed,
                "留存RF組 SKU數": keep["keep_skus"],
                "留存RF組上月 Sold 合計(月化)": keep["keep_sold_prev"],
                "留存RF組本月 Sold 合計(月化)": keep["keep_sold_curr"],
                "留存RF組 Sold 變化量": keep["keep_sold_chg"],
                "留存RF組 Sold 變化率": keep["keep_sold_rate"],
                "留存RF組上月 Net 合計(月化)": keep["keep_net_prev"],
                "留存RF組本月 Net 合計(月化)": keep["keep_net_curr"],
                "留存RF組 Net 變化量": keep["keep_net_chg"],
                "留存RF組 Net 變化率": keep["keep_net_rate"],
                "全店 Sold 變化量(月化基準)": sold_chg,
                "全店 Net 變化量(月化基準)": net_chg,
            })
            prev_rf, prev_total, prev_sold, prev_net = t["rf"], t["total"], t["sold"], t["net"]
    void = [k for k in label_to_key if k not in set(ym_labels)]
    _ = void
    return pd.DataFrame(rows, columns=_SAFM_SHOP_EFFECT_COLUMNS)


def _build_safm_t0_track(
    change_df: pd.DataFrame | None,
    monthly: pd.DataFrame | None,
    ym_keys: list[int] | None = None,
    ym_labels: list[str] | None = None,
) -> pd.DataFrame:
    """轉變前後追蹤（T0 對齊）：以轉變發生的本月為 T0，追蹤 T-1／T0／T+1。

    只收錄 RF→ND 與 ND→RF 事件；T+1 為資料內下一個 YEAR-MONTH，
    缺資料者 T+1 各欄留空並於去化狀態標示未有次月資料。
    RF→ND 屬退出自動補貨，預期存貨逐步去化；ND→RF 屬恢復補貨，預期存貨回升。
    """
    if change_df is None or change_df.empty or monthly is None or monthly.empty:
        return pd.DataFrame(columns=_SAFM_T0_COLUMNS)
    ym_keys = list(ym_keys or sorted(monthly["_ym"].unique().tolist()))
    if len(ym_keys) < 2:
        return pd.DataFrame(columns=_SAFM_T0_COLUMNS)
    if not ym_labels:
        ym_labels = [_ym_label(k // 100, k % 100) for k in ym_keys]
    next_label = {str(ym_labels[i]): str(ym_labels[i + 1]) for i in range(len(ym_labels) - 1)}

    lookup: dict[tuple, dict] = {}
    for row in monthly.itertuples(index=False, name=None):
        try:
            cols = list(monthly.columns)
            site = str(row[cols.index("_site")])
            sku = row[cols.index("_sku")]
            key = int(row[cols.index("_ym")])
            lookup[(site, str(sku), key)] = {
                "sold": int(row[cols.index("sold_adj")]),
                "net": round(float(row[cols.index("net_adj")]), 2),
                "stock": int(row[cols.index("stock")]),
                "safety": int(row[cols.index("safety")]),
                "rp": str(row[cols.index("rp")]),
            }
        except (ValueError, TypeError, IndexError):
            continue

    sub = change_df[change_df["RP Type 轉變"].isin(["RF→ND", "ND→RF"])].copy()
    if sub.empty:
        return pd.DataFrame(columns=_SAFM_T0_COLUMNS)
    rows: list[dict] = []
    for r in sub.itertuples(index=False, name=None):
        cols = list(sub.columns)
        try:
            site = str(r[cols.index("Site Code")])
            sku = r[cols.index("SKU")]
            t0_label = str(r[cols.index("本月")])
            t1_label = str(r[cols.index("上月")])
        except (ValueError, IndexError):
            continue
        t0_key = _safm_label_to_key(t0_label)
        if t0_key is None:
            continue
        t1_sold = r[cols.index("上月 Sold Qty 月化")] if "上月 Sold Qty 月化" in cols else None
        t0_sold = r[cols.index("本月 Sold Qty 月化")] if "本月 Sold Qty 月化" in cols else None
        t1_net = r[cols.index("上月 Net Sales 月化")] if "上月 Net Sales 月化" in cols else None
        t0_net = r[cols.index("本月 Net Sales 月化")] if "本月 Net Sales 月化" in cols else None
        t1_stock = r[cols.index("上月 Shop Stock Qty (CM)")] if "上月 Shop Stock Qty (CM)" in cols else None
        t0_stock = r[cols.index("本月 Shop Stock Qty (CM)")] if "本月 Shop Stock Qty (CM)" in cols else None
        t1_safety = r[cols.index("上月 Adj. Safety Stock")] if "上月 Adj. Safety Stock" in cols else None
        t0_safety = r[cols.index("本月 Adj. Safety Stock")] if "本月 Adj. Safety Stock" in cols else None
        try:
            t1_sold = int(t1_sold) if t1_sold is not None and not pd.isna(t1_sold) else None
            t0_sold = int(t0_sold) if t0_sold is not None and not pd.isna(t0_sold) else None
            t1_net = round(float(t1_net), 2) if t1_net is not None and not pd.isna(t1_net) else None
            t0_net = round(float(t0_net), 2) if t0_net is not None and not pd.isna(t0_net) else None
            t1_stock = int(t1_stock) if t1_stock is not None and not pd.isna(t1_stock) else None
            t0_stock = int(t0_stock) if t0_stock is not None and not pd.isna(t0_stock) else None
            t1_safety = int(t1_safety) if t1_safety is not None and not pd.isna(t1_safety) else None
            t0_safety = int(t0_safety) if t0_safety is not None and not pd.isna(t0_safety) else None
        except (TypeError, ValueError):
            pass
        t2_label = next_label.get(t0_label)
        nxt = None
        if t2_label is not None:
            t2_key = _safm_label_to_key(t2_label)
            if t2_key is not None:
                nxt = lookup.get((site, str(sku), int(t2_key)))
        if nxt is None:
            t2_sold = t2_net = t2_stock = t2_safety = t2_rp = None
            t1_chg = None
            status = "未有次月資料"
        else:
            t2_sold, t2_net, t2_stock, t2_safety, t2_rp = nxt["sold"], nxt["net"], nxt["stock"], nxt["safety"], nxt["rp"]
            t1_chg = None
            try:
                if t0_sold is not None and t2_sold is not None:
                    t1_chg = int(t2_sold - t0_sold)
            except (TypeError, ValueError):
                t1_chg = None
            trans = str(r[cols.index("RP Type 轉變")])
            if trans == "RF→ND":
                if t0_stock is not None and t2_stock is not None:
                    if t0_stock == 0 and t2_stock == 0:
                        status = "已無存貨"
                    elif t2_stock == 0:
                        status = "已清空"
                    elif t2_stock < t0_stock:
                        status = "去化中"
                    elif t2_stock == t0_stock:
                        status = "持平滯留"
                    else:
                        status = "存貨回升(異常)"
                else:
                    status = "資料不齊"
            else:
                if t0_stock is not None and t2_stock is not None:
                    if t2_stock > t0_stock:
                        status = "存貨回升"
                    elif t2_stock == t0_stock:
                        status = "持平"
                    else:
                        status = "存貨下降"
                else:
                    status = "資料不齊"
        try:
            t0_chg = int(t0_sold - t1_sold) if t0_sold is not None and t1_sold is not None else None
        except (TypeError, ValueError):
            t0_chg = None
        rows.append({
            "轉變類型": str(r[cols.index("RP Type 轉變")]),
            "Site Code": site,
            "Shop Name": str(r[cols.index("Shop Name")]) if "Shop Name" in cols else "",
            "SKU": sku,
            "SKU Desc": str(r[cols.index("SKU Desc")]) if "SKU Desc" in cols else "",
            "Brand": str(r[cols.index("Brand")]) if "Brand" in cols else "",
            "Buyer": str(r[cols.index("Buyer")]) if "Buyer" in cols else "",
            "Adj. MCH1": str(r[cols.index("Adj. MCH1")]) if "Adj. MCH1" in cols else "",
            "Adj. SST": str(r[cols.index("Adj. SST")]) if "Adj. SST" in cols else "",
            "T-1月份": t1_label, "T0月份": t0_label, "T+1月份": t2_label or "",
            "T-1 RP Type": str(r[cols.index("上月 RP Type")]) if "上月 RP Type" in cols else "",
            "T0 RP Type": str(r[cols.index("本月 RP Type")]) if "本月 RP Type" in cols else "",
            "T+1 RP Type": t2_rp or "",
            "T-1 Sold Qty 月化": t1_sold, "T0 Sold Qty 月化": t0_sold, "T+1 Sold Qty 月化": t2_sold,
            "T0環比變化量": t0_chg, "T+1環比變化量": t1_chg,
            "T-1 Net Sales 月化": t1_net, "T0 Net Sales 月化": t0_net, "T+1 Net Sales 月化": t2_net,
            "T-1 Net Stock Qty": t1_stock, "T0 Net Stock Qty": t0_stock, "T+1 Net Stock Qty": t2_stock,
            "T-1 Safety Stock Qty": t1_safety, "T0 Safety Stock Qty": t0_safety, "T+1 Safety Stock Qty": t2_safety,
            "去化狀態": status,
        })
    detail = pd.DataFrame(rows, columns=_SAFM_T0_COLUMNS)
    if not detail.empty:
        order = {"RF→ND": 0, "ND→RF": 1}
        detail["_o"] = detail["轉變類型"].map(order).fillna(99).astype(int)
        detail = detail.sort_values(
            ["_o", "Brand", "SKU", "Site Code", "T0月份"],
            kind="stable").drop(columns=["_o"]).reset_index(drop=True)
    return detail


def safety_monthly_dashboard_metrics(change_df: pd.DataFrame, stats: dict | None = None):
    """月度摘要指標卡（與 Streamlit 看板及 Excel 摘要同一準則）。"""
    stats = stats or {}
    total_events = int(stats.get("total_events", len(change_df) if change_df is not None else 0))
    row1 = [
        ("環比事件數", total_events, "相鄰月份比較事件總數", "3B82F6"),
        ("Safety Stock Qty 增加事件", int(stats.get("safety_inc", 0)), "原始值增加的事件數", "EF4444"),
        ("Safety Stock Qty 減少事件", int(stats.get("safety_dec", 0)), "原始值減少的事件數", "10B981"),
        ("Safety Stock Qty 不變事件", int(stats.get("safety_same", 0)), "原始值不變的事件數", "64748B"),
    ]
    row2 = [
        ("Safety Stock Qty 淨變化量", int(stats.get("safety_delta_total", 0)), "全部事件變化量合計（件）", "8B5CF6"),
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


_SAFM_PIVOT_BLOCKS = ("每月合計", "月環比%", "月環比")


def _safm_cjk_len(text: str) -> float:
    # 全形（W／F）與模糊寬度（A，如「→」「×」）在 Excel 一律以約 2 個字元寬渲染
    return sum(
        2.0 if unicodedata.east_asian_width(c) in ("W", "F", "A") else 1.0
        for c in str(text)
    )


def _safm_fit_pivot_columns(ws, dim_count: int, ym_count: int, block_count: int | None = None):
    """並排樞紐欄寬：按格式化後的顯示字元數計算（數值含千分位／小數），
    表頭預留自動篩選按鈕寬度，避免出現 ########## 或月份表頭被按鈕遮尾。"""
    if block_count is None:
        block_count = len(_SAFM_PIVOT_METRICS)
    block_width = 1 + ym_count + 1
    span = block_count * block_width
    n_cols = dim_count + span
    for col_idx in range(1, n_cols + 1):
        header = str(ws.cell(row=1, column=col_idx).value or "")
        need = _safm_cjk_len(header) + 4.5
        money = False
        if dim_count < col_idx <= dim_count + span:
            rel = col_idx - dim_count - 1
            offset = rel % block_width
            block_idx = rel // block_width
            money = (
                offset > 0 and block_idx < len(_SAFM_PIVOT_METRICS)
                and _SAFM_PIVOT_METRICS[block_idx] == "Net Sales"
            )
        for row_idx in range(2, ws.max_row + 1):
            v = ws.cell(row=row_idx, column=col_idx).value
            if v is None or isinstance(v, bool):
                continue
            if isinstance(v, (int, float)):
                text = f"{v:,.2f}" if money else f"{v:,.0f}"
            else:
                text = str(v)
            need = max(need, _safm_cjk_len(text) + 2.0)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(need, 8.0), 50.0)


def _write_safm_pivot_sheet(ws, df: pd.DataFrame, progress_callback=None,
                            progress_start=0.0, progress_end=0.1):
    """三段式樞紐工作表（四項指標區塊左右並排）。

    版面：維度欄置最左（只出現一次），其後 Sold Qty／Net Sales／Net Stock Qty／
    Safety Stock Qty 各成一個區塊，每區塊自帶「項目」＋各 YEAR-MONTH＋「全部合計」欄；
    區塊之後為 RF→ND SKU數／ND→RF SKU數 兩個分月份區塊（結構同指標區塊，
    僅每月合計段有值，環比兩段留白）；每段一列一維度，總計列（粗體灰底）置於段末。
    輸入仍為長格式樞紐表（每指標連續三列），於寫出時重排為並排區塊。
    """
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    total_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    total_font = Font(bold=True, size=11)
    left = Alignment(horizontal="left", vertical="center")
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    if df is None or df.empty:
        ws.cell(row=1, column=1, value="無符合資料").font = Font(size=12)
        return
    headers = list(df.columns)
    item_idx = headers.index("項目")
    total_idx = headers.index("全部合計")
    dim_cols = headers[:item_idx]
    ym_cols = headers[item_idx + 1:total_idx]
    block_width = 1 + len(ym_cols) + 1

    def _split_item(text: str) -> tuple[str, str]:
        for block in _SAFM_PIVOT_BLOCKS:
            suffix = f" {block}"
            if text.endswith(suffix):
                return text[: -len(suffix)], block
        return text, ""

    lookup: dict[tuple[tuple, str, str], tuple] = {}
    dim_order: list[tuple] = []
    for row_data in df.itertuples(index=False, name=None):
        dim_key = tuple(row_data[:item_idx])
        metric, block = _split_item(str(row_data[item_idx]))
        lookup[(dim_key, block, metric)] = row_data
        if dim_key not in dim_order:
            dim_order.append(dim_key)

    header_cells = list(dim_cols)
    for _metric in _SAFM_PIVOT_ALL_METRICS:
        header_cells += ["項目"] + list(ym_cols) + ["全部合計"]
    for col_idx, h in enumerate(header_cells, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    total_rows = len(_SAFM_PIVOT_BLOCKS) * len(dim_order)
    written = 0
    row_idx = 1
    for block in _SAFM_PIVOT_BLOCKS:
        for dim_key in dim_order:
            row_idx += 1
            is_total = str(dim_key[0]) == "總計"
            for col_idx, v in enumerate(dim_key, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=None if pd.isna(v) else v)
                cell.border = thin_border
                cell.alignment = left_wrap
                if is_total:
                    cell.font = total_font
                    cell.fill = total_fill
            base_col = len(dim_cols)
            for metric in _SAFM_PIVOT_ALL_METRICS:
                row_data = lookup.get((dim_key, block, metric))
                if row_data is None:
                    # RP 轉變 SKU 統計無環比段：整個區塊留白（保留邊框）
                    for offset in range(block_width):
                        ws.cell(row=row_idx, column=base_col + 1 + offset).border = thin_border
                    base_col += block_width
                    continue
                item_cell = ws.cell(row=row_idx, column=base_col + 1, value=f"{metric} {block}")
                item_cell.border = thin_border
                item_cell.alignment = left
                if is_total:
                    item_cell.font = total_font
                    item_cell.fill = total_fill
                values = (
                    row_data[item_idx + 1:total_idx + 1]
                    if row_data is not None else [None] * (block_width - 1)
                )
                is_money = metric == "Net Sales"
                for offset, v in enumerate(values):
                    cell = ws.cell(
                        row=row_idx, column=base_col + 2 + offset,
                        value=None if pd.isna(v) else v,
                    )
                    cell.border = thin_border
                    cell.alignment = right
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        cell.number_format = "#,##0.00" if is_money else "#,##0"
                    if is_total:
                        cell.font = total_font
                        cell.fill = total_fill
                base_col += block_width
            written += 1
            if progress_callback and (written % 200 == 0 or written == total_rows):
                p = progress_start + (progress_end - progress_start) * (written / max(total_rows, 1))
                progress_callback(p, f"寫入工作表 ({written}/{total_rows})...")

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = f"{get_column_letter(len(dim_cols) + 1)}2"
    _safm_fit_pivot_columns(ws, len(dim_cols), len(ym_cols), len(_SAFM_PIVOT_ALL_METRICS))


def _write_safm_block(ws, df: pd.DataFrame, start_row: int = 1,
                      progress_callback=None, progress_start=0.0, progress_end=0.1,
                      total_hint: int | None = None) -> int:
    """於工作表寫入一個表格區塊（表頭＋資料列），回傳下一個可用列號，不設篩選與凍結。

    P0：逐格 isinstance／pd.isna 分支改為按欄位 dtype 預先分類——float 欄按表頭定
    數字格式（變化率→0.00%、Net→#,##0.00、其他→#,##0）、其餘欄逐格判定，
    Alignment／Border 物件全區塊共用，NaN 於寫出前一次性行轉 None；
    儲存格值、數字格式與對齊與原逐格實作完全一致。
    """
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
    align_center = Alignment(horizontal="center", vertical="center")
    align_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, h in enumerate(display, 1):
        cell = ws.cell(row=start_row, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_wrap
        cell.border = thin_border

    # 欄位級預判定：(數字格式或 None, 對齊)。None＝混合型欄，逐格按值判定。
    col_spec: list[tuple[str | None, Alignment, Alignment]] = []
    for h, dtype in zip(headers, df.dtypes):
        if pd.api.types.is_float_dtype(dtype):
            if "變化率" in str(h):
                fmt = "0.00%"
            elif "Net" in str(h):
                fmt = "#,##0.00"
            else:
                fmt = "#,##0"
            col_spec.append((fmt, align_center, align_wrap))
        elif pd.api.types.is_integer_dtype(dtype):
            col_spec.append(("#,##0", align_center, align_wrap))
        else:
            col_spec.append((None, align_center, align_wrap))

    # NaN／None 於寫出前一次性行轉 None，免去逐格 pd.isna
    values = df.astype(object)
    for _c in df.columns:
        _mask = df[_c].isna()
        if _mask.any():
            values.loc[_mask, _c] = None

    total = len(df)
    for offset, row_data in enumerate(values.itertuples(index=False, name=None), 1):
        row_idx = start_row + offset
        for col_idx, (v, (fmt, align_num, align_wrap_col)) in enumerate(
            zip(row_data, col_spec), 1
        ):
            cell = ws.cell(row=row_idx, column=col_idx, value=v)
            cell.border = thin_border
            if fmt is not None:
                cell.alignment = align_wrap_col if v is None else align_num
                if v is not None:
                    cell.number_format = fmt
            elif v is None or isinstance(v, (str, bool)):
                cell.alignment = align_wrap_col
            elif isinstance(v, float):
                cell.alignment = align_num
                if "變化率" in headers[col_idx - 1]:
                    cell.number_format = "0.00%"
                elif "Net" in headers[col_idx - 1]:
                    cell.number_format = "#,##0.00"
                else:
                    cell.number_format = "#,##0"
            elif isinstance(v, int):
                cell.alignment = align_num
                cell.number_format = "#,##0"
            else:
                cell.alignment = align_wrap_col
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
        ("分析欄", "Adj. Safety Stock（只看原始值增減）／RP Type 轉變／Sold Qty (CM - MTD)／Net Sales (CM - MTD)／Shop Stock Qty (CM)；"
         "Sold Qty 與 Net Sales 同時月化，Shop Stock Qty 與 Adj. Safety Stock 為快照不月化"),
        ("維度", "Adj. MCH1、Adj. SST、MCH1×SST、Site Code；另按 RP Type 轉變群組、店舖精簡成效與 T0 轉變前後追蹤三個生意視角覆核"),
    ]
    logic = [
        "YEAR-MONTH 由 YEAR × 100 ＋ MONTH 排序；MONTH 須為 1–12 整數，無效列略過並列警告。",
        "最新月份月化：Sold Qty 月化＝Sold Qty (CM - MTD) × 係數；Net Sales 月化＝Net Sales (CM - MTD) × 係數；舊月份月化值等於原始值。",
        "Safety Stock Qty 變化量＝本月 Adj. Safety Stock − 上月；正數為增加、負數為減少、零為不變（只看原始值，不作建議值計算）。",
        "RP Type 轉變記錄上月→本月（如 RF→ND）；Sold Qty／Net Sales 變化量以月化基準計算，變化率＝變化量 ÷ 上月絕對值（上月為 0 時留空）。",
        "首尾月份對比另附最早／最新月份的 Sold Qty 月化、Net Sales 月化與 Safety Stock Qty：變化量以月化基準（Safety Stock Qty 為原始值），"
        "變化率＝變化量 ÷ 最早月份絕對值（最早為 0 時留空）。",
        "樞紐四項指標：Sold Qty＝Sold Qty (CM - MTD) 月化合計；Net Sales＝Net Sales (CM - MTD) 月化合計；"
        "Net Stock Qty＝店舖存貨 Shop Stock Qty (CM) 合計（時點快照）；Safety Stock Qty＝Adj. Safety Stock 合計。",
        "樞紐三段式：每月合計 → 月環比%（上月為 0 或首月留空）→ 月環比絕對值；"
        "四項指標區塊左右並排，每區塊各含項目、各月與全部合計欄，每段一列一維度，總計列置於段末。",
        "RF→ND SKU數／ND→RF SKU數：三個樞紐按維度、By Shop 按店舖，分月份統計當月發生該轉變的不重複 SKU 數"
        "（按轉變的「本月」歸月，取自 SKU月度變化事件；樞紐的「全部合計」為整個期間不重複 SKU 數，"
        "總計列為全檔口徑，各維度相加可大於總計——同一 SKU 可在多個維度、多個月份轉變）。",
        "RP 轉變群組生意對比：母體同為 SKU月度變化事件，按 RF→RF／ND→ND／RF→ND／ND→RF 四組各自加總上月與本月生意，"
        "避免轉 ND 群組與持續群組混合；重點比較 RF→ND 群組與 RF→RF 群組的同期變化率。",
        "店舖精簡成效：每店每月列示 RF／ND／總 SKU 數及其環比變化、淨精簡數（RF→ND 減 ND→RF）、"
        "被精簡掉的上月生意基數（RF→ND 組上月合計）與留存 RF 組（RF→RF 事件）上月與本月對比；"
        "留存組上升而總 SKU 下降，即其餘 SKU 承接生意；首月無上月可比，轉變計數恆為 0。",
        "轉變前後追蹤（T0 對齊）：只收 RF→ND 與 ND→RF 事件，以轉變本月為 T0，列示 T-1／T0／T+1 的生意、存貨與 Safety；"
        "T+1 取資料內下一個 YEAR-MONTH，缺資料者標示未有次月資料；RF 代表執行自動補貨，ND 代表不執行自動補貨，"
        "故 RF→ND 後預期存貨逐步去化（已清空／去化中／持平滯留），ND→RF 後預期存貨回升。",
    ]
    sheet_rows = [
        ("說明（本頁）", "月化邏輯、本次參數與導讀", "先核對覆蓋月份與月化係數是否符合預期"),
        ("SKU月度變化", "每組相鄰月份環比事件（上月／本月值＋變化量＋變化率）", "篩選 Safety Stock Qty 變化分類＝增加／減少，查看同期 Sold Qty 與 Net Sales 變化"),
        ("By Adj. MCH1", "各 MCH1 三段式樞紐（四項指標區塊左右並排：每月合計／月環比%／月環比）＋RF→ND／ND→RF SKU數（分月份）", "先看每月合計，再看月環比% 找變動最大的品類"),
        ("By Adj. SST", "各 SST 三段式樞紐（版面同 By Adj. MCH1）", "比較不同供應類型每月生意變化"),
        ("By MCH1×SST", "MCH1 與 SST 組合三段式樞紐（版面同 By Adj. MCH1）", "細分品類 × 供應類型"),
        ("By Shop", "每間店舖每月合計＋環比變化＋RF→ND／ND→RF SKU數（當月）", "多店舖檔案時比較各店走勢"),
        ("RP Type 變化", "RP Type 轉變 × 月份配對統計＋首尾月份對比明細（含 Sold Qty／Net Sales 月化與 Safety Stock Qty 對比）", "重點看 RF→ND 與 ND→RF 兩列：哪些月份最多轉換；下半首尾對比看整個期間的最終轉變，以及生意與 Safety 設定的前後變化"),
        ("RP轉變生意對比", "按 RP Type 轉變 × 本月分組的生意合計與變化（事件數／SKU數／Sold／Net Sales／存貨／Safety）", "直接比較 RF→ND 群組與 RF→RF 群組的同期變化率，判斷轉 ND 對生意有無影響"),
        ("店舖精簡成效", "每店每月 RF／ND／總 SKU 數、淨精簡數、被精簡生意基數與留存 RF 組生意對比", "總 SKU 下降而留存 RF 組生意上升，即精簡後其餘 SKU 承接生意"),
        ("轉變前後追蹤(T0)", "RF→ND 與 ND→RF 事件以轉變本月為 T0 的 T-1／T0／T+1 對比＋去化狀態", "逐列查看轉 ND 後次月是否去化（已清空／去化中）或滯留"),
        ("月度摘要", "指標卡（含 RF→ND／ND→RF 一排）＋每月合計", "管理層概覽"),
        ("RAW", "原始明細＋YEAR-MONTH 與月化估算欄", "核對逐列數值"),
    ]
    tips = [
        "最新月份為非足月時，Sold Qty 與 Net Sales 的環比已按月化基準計算；原始 MTD 值仍保留於主表以供核對。",
        "數據截止日與最新月份不一致時，最新月份不作月化（係數為 1），並列出警告。",
        "上月缺資料的組合不計變化，屬新增組合；變化率分母為 0 時留空。",
        "RF→ND 的 SKU 轉為不可補貨後，留意其後月份 Sold Qty 與 Net Sales 是否受影響；ND→RF 則留意補貨是否恢復。",
        "判斷轉 ND 對生意有無影響時，以 RP轉變生意對比的群組變化率為準，而非全店合計；全店合計同時受持續群組影響。",
        "驗證精簡假設時，以店舖精簡成效的留存 RF 組變化為準；若留存組上升而全店持平或上升，表示其餘 SKU 已承接被精簡 SKU 的生意。",
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
    rp_cohort_df: pd.DataFrame | None = None,
    shop_effect_df: pd.DataFrame | None = None,
    t0_df: pd.DataFrame | None = None,
) -> BytesIO:
    """Safety Stock 月度變化分析輸出十二個工作表（說明預設開啟）。"""
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
        progress_callback(0.76, "建立 RP 轉變生意對比...")
    _rp_cohort = rp_cohort_df
    if _rp_cohort is None:
        _rp_cohort = (stats or {}).get("rp_cohort")
    if _rp_cohort is None:
        _rp_cohort = _build_safm_rp_cohort(change_df)
    ws_cohort = wb.create_sheet(title="RP轉變生意對比")
    _write_safm_table_sheet(ws_cohort, _rp_cohort, progress_callback, 0.76, 0.79)

    if progress_callback:
        progress_callback(0.79, "建立店舖精簡成效...")
    _shop_effect = shop_effect_df
    if _shop_effect is None:
        _shop_effect = (stats or {}).get("shop_effect")
    ws_effect = wb.create_sheet(title="店舖精簡成效")
    _write_safm_table_sheet(ws_effect, _shop_effect, progress_callback, 0.79, 0.82)

    if progress_callback:
        progress_callback(0.82, "建立轉變前後追蹤...")
    _t0 = t0_df
    if _t0 is None:
        _t0 = (stats or {}).get("t0_track")
    ws_t0 = wb.create_sheet(title="轉變前後追蹤(T0)")
    _write_safm_table_sheet(ws_t0, _t0, progress_callback, 0.82, 0.85)

    if progress_callback:
        progress_callback(0.85, "建立月度摘要...")
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
        progress_pct=0.86, row3_metrics=row3,
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
            SKU數=("SKU數", "sum"),
            SafetyStockQty合計=("Safety Stock Qty 合計", "sum"),
            SoldQty月化合計=("Sold Qty 月化合計", "sum"),
            NetSales月化合計=("Net Sales 月化合計", "sum"),
            NetStockQty合計=("Net Stock Qty 合計", "sum"),
        ).reset_index()
        _rp_month = {}
        if change_df is not None and not change_df.empty and "本月" in change_df.columns:
            _rp_ct = change_df.groupby("本月")["RP Type 轉變"].value_counts().unstack(fill_value=0)
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
                "Safety Stock Qty 合計": r.get("SafetyStockQty合計"),
                "Sold Qty 合計 (月化)": r.get("SoldQty月化合計"),
                "Net Sales 合計 (月化)": r.get("NetSales月化合計"),
                "Net Stock Qty 合計": r.get("NetStockQty合計"),
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
