# -*- coding: utf-8 -*-
"""Stock Impact Analysis模式：由 transformer.py 拆分而來。
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
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _cached_safe_sku_series,
    _write_metric_cards,
    _write_raw_placeholder,
    _write_raw_sheet,
    validate_input,
)
__all__ = [
    "_JOSEPH_RAW_COLS",
    "_OOS_DETAIL_COLS",
    "_PRICE_TIER_BOUNDS",
    "_PRICE_TIER_LABELS",
    "_SOH_TIER_BINS",
    "_SOH_TIER_LABELS",
    "_STOCK_IMPACT_NUMERIC_COLS",
    "_STOCK_IMPACT_SUMMARY_COLS",
    "_TURNDAY_BAND_ORDER",
    "_add_launch_info",
    "_build_inventory_distribution_summary",
    "_build_low_stock_item_class_summary",
    "_build_sell_through_band_summary",
    "_build_sku_allocation_summary",
    "_build_sku_soh_tier_summary",
    "_build_sku_stock_diagnosis",
    "_build_stock_impact_summary",
    "_build_summary_safe",
    "_build_transfer_priority_analysis",
    "_classify_d001_status",
    "_enrich_raw_joseph",
    "_preprocess_numeric_columns",
    "_price_tier_label",
    "_soh_tier_label",
    "_write_inventory_distribution_sheet",
    "_write_new_sheet",
    "_write_oos_detail_sheet",
    "_write_sku_stock_diagnosis_sheet",
    "_write_stock_impact_dashboard",
    "_write_stock_impact_dashboard_simplified",
    "_write_stock_impact_summary_sheet",
    "generate_excel_stock_impact",
    "generate_excel_stock_impact_simplified",
    "transform_stock_impact",
    "_add_stock_impact_columns",
    "_build_sku_info_lookup",
    "_build_sku_monthly_analysis",
    "_build_sku_monthly_analysis_simplified",
    "_STOCK_SOLD_PERIODS",
    "_write_sku_allocation_sheet",
    "_NATURE_PRIORITY",
    "_NATURE_ORDER",
]


def _add_launch_info(df: pd.DataFrame, ref_date: datetime.date | None = None) -> pd.DataFrame:
    """從 Launch Date 新增 認證是否New SKU(3個月內) 欄位。

    Args:
        ref_date: 參考日期（預設今天）。
    """
    if "Launch Date" not in df.columns:
        return df  # 無 Launch Date 欄位則跳過

    if ref_date is None:
        ref_date = datetime.date.today()

    df = df.copy()
    launch_raw = df["Launch Date"].fillna("").astype(str).str.strip()
    launch_dt = _parse_launch_dates(launch_raw)

    # 3個月內（約 90 天）視為 New SKU
    cutoff = pd.Timestamp(ref_date) - pd.Timedelta(days=90)
    df["Is New SKU (3mo)"] = launch_dt.notna() & (launch_dt >= cutoff)
    df["Is New SKU (3mo)"] = df["Is New SKU (3mo)"].map({True: "✓ New SKU", False: ""})

    # 保留 Launch Date 原始值
    df["Launch Date"] = launch_raw
    return df



# 多維度彙總輸出欄位（中文表頭）
_STOCK_IMPACT_SUMMARY_COLS = [
    "缺貨行數", "涉及SKU數",
    "Tourist 缺貨行數", "Mix 缺貨行數", "Non Tourist 缺貨行數",
    "長期缺貨行數", "間歇缺貨行數", "恢復供貨行數",
]



def _build_stock_impact_summary(raw: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    """依 group_cols 彙總缺貨影響。只統計 Is Stock Impacted = True 的列。

    排序優先級：Tourist 缺貨 → 缺貨行數（皆降序）。
    """
    impacted = raw[raw["Is Stock Impacted"]].copy()
    if impacted.empty:
        col_names = group_cols + _STOCK_IMPACT_SUMMARY_COLS
        return pd.DataFrame(columns=col_names)

    impacted["_nature"] = impacted.get("Shop Nature", pd.Series(dtype=object)).fillna("")
    impacted["_pattern"] = impacted.get("OOS Pattern", pd.Series(dtype=object)).fillna("")

    # Vectorized pre-computation: avoid slow lambda in groupby
    impacted["_is_tourist"] = (impacted["_nature"] == "Tourist").astype(int)
    impacted["_is_mix"] = (impacted["_nature"] == "Mix").astype(int)
    impacted["_is_nt"] = (impacted["_nature"] == "Non Tourist").astype(int)
    impacted["_is_long"] = (impacted["_pattern"] == "長期缺貨").astype(int)
    impacted["_is_inter"] = (impacted["_pattern"] == "間歇缺貨").astype(int)
    impacted["_is_recover"] = (impacted["_pattern"] == "恢復供貨").astype(int)

    agg = (
        impacted.groupby(group_cols, dropna=False, observed=True)
        .agg(
            缺貨行數=("_nature", "count"),
            涉及SKU數=("SKU", "nunique"),
            **{
                "Tourist 缺貨行數": ("_is_tourist", "sum"),
                "Mix 缺貨行數": ("_is_mix", "sum"),
                "Non Tourist 缺貨行數": ("_is_nt", "sum"),
                "長期缺貨行數": ("_is_long", "sum"),
                "間歇缺貨行數": ("_is_inter", "sum"),
                "恢復供貨行數": ("_is_recover", "sum"),
            },
        )
        .reset_index()
    )

    for c in _STOCK_IMPACT_SUMMARY_COLS:
        agg[c] = agg[c].fillna(0).astype(int)

    agg = agg.sort_values(
        ["Tourist 缺貨行數", "缺貨行數"],
        ascending=[False, False],
    ).reset_index(drop=True)

    # 總計列（涉及SKU數為整體不重複計數，非各組加總）
    total_data = {col: "" for col in group_cols}
    total_data[group_cols[0]] = "總計"
    for c in _STOCK_IMPACT_SUMMARY_COLS:
        if c == "涉及SKU數":
            total_data[c] = int(impacted["SKU"].nunique())
        else:
            total_data[c] = int(agg[c].sum())
    agg = pd.concat([agg, pd.DataFrame([total_data])], ignore_index=True)

    for c in ["_nature", "_pattern"]:
        if c in agg.columns:
            agg.drop(columns=[c], inplace=True)

    return agg



def _build_sku_allocation_summary(raw: pd.DataFrame, sku_lookup: pd.DataFrame | None = None) -> pd.DataFrame:
    """SKU 級跨 Nature 調貨建議。

    針對同時存在於多種 Shop Nature 的 SKU：當 Tourist 店舖缺貨（CM Shop Stock=0）
    但同 SKU 在 Non-Tourist / Mix 店舖仍有存貨時，列出可調貨支援機會。
    排序：Tourist 缺貨損失 → 可調貨存貨量（降序）。
    
    Args:
        sku_lookup: 預先建好的 SKU 基本資訊查找表，為 None 時即時計算。
    """
    if "SKU" not in raw.columns or "Shop Nature" not in raw.columns:
        return pd.DataFrame()

    # 先找出同時出現在 Tourist 與其他 Nature 的 SKU（向量化替代 apply+set）
    _nature_dedup = raw.dropna(subset=["Shop Nature"]).groupby("SKU")["Shop Nature"].agg("unique")
    _nature_sets = _nature_dedup.apply(lambda arr: set(arr))
    shared_skus = _nature_sets[_nature_sets.apply(lambda s: "Tourist" in s and len(s) >= 2)].index
    if len(shared_skus) == 0:
        return pd.DataFrame()

    sub = raw[raw["SKU"].isin(shared_skus)].copy()
    if sub.empty:
        return pd.DataFrame()

    cm_stock = _safe_numeric(sub.get("Shop Stock Qty (CM)", _safe_default_series(0, sub.index)))
    cm_sold = _safe_numeric(sub.get("Sold Qty (CM - MTD)", _safe_default_series(0, sub.index)))
    m2_sold = _safe_numeric(sub.get("Sold Qty (M-2)", _safe_default_series(0, sub.index)))
    p90d = _safe_numeric(sub.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, sub.index)))
    loss_m2 = _safe_numeric(sub.get("Max Sales Loss (M-2)", _safe_default_series(0, sub.index)))
    loss_p90d = _safe_numeric(sub.get("Max Sales Loss (P90D)", _safe_default_series(0, sub.index)))
    sub["_cm_stock"] = cm_stock
    sub["_cm_sold"] = cm_sold
    sub["_m2_sold"] = m2_sold
    sub["_p90d"] = p90d
    sub["_loss_m2"] = loss_m2
    sub["_loss_p90d"] = loss_p90d

    # Tourist 缺貨行（CM Shop Stock = 0 且有需求）
    tourist_oos = sub[(sub["Shop Nature"] == "Tourist") & (cm_stock <= 0) & ((m2_sold > 0) | (p90d > 0))]
    # 其他 Nature 有存貨的行
    other_stock = sub[(sub["Shop Nature"] != "Tourist") & (cm_stock > 0)]

    if tourist_oos.empty or other_stock.empty:
        return pd.DataFrame()

    # 以 SKU + Brand 聚合 Tourist 缺貨損失
    t_agg = (
        tourist_oos.groupby(["SKU"] + (["Brand"] if "Brand" in tourist_oos.columns else []), dropna=False)
        .agg(
            Tourist缺貨店數=("SKU", "count"),
            Tourist損失_M2=("_loss_m2", "sum"),
            Tourist損失_P90D=("_loss_p90d", "sum"),
            Tourist_M2銷售件數=("_m2_sold", "sum"),
        )
        .reset_index()
    )

    # 其他 Nature 可調貨存貨
    o_agg = (
        other_stock.groupby(["SKU"] + (["Shop Nature"]), dropna=False)
        .agg(
            可調貨存貨件數=("_cm_stock", "sum"),
            閒置店數=("SKU", "count"),
        )
        .reset_index()
    )
    # 把其他 Nature 的存貨攤平成欄（Mix / Non Tourist）
    o_pivot = o_agg.pivot_table(
        index=["SKU"],
        columns="Shop Nature",
        values=["可調貨存貨件數", "閒置店數"],
        fill_value=0,
        aggfunc="sum",
    ).reset_index()
    # 扁平化多層欄名
    flat_cols = []
    for c in o_pivot.columns:
        if isinstance(c, tuple):
            metric, nature = c
            if metric:
                flat_cols.append(f"{nature}{metric}")
            else:
                flat_cols.append(nature)
        else:
            flat_cols.append(c)
    o_pivot.columns = flat_cols

    merged = t_agg.merge(o_pivot, on="SKU", how="left")

    # 確保 Mix/Non Tourist 欄位存在
    for n in ["Mix", "Non Tourist"]:
        for metric in ["可調貨存貨件數", "閒置店數"]:
            col = f"{n}{metric}"
            if col not in merged.columns:
                merged[col] = 0
            merged[col] = merged[col].fillna(0).astype(int)

    merged["可調貨總存貨件數"] = (merged.get("Mix可調貨存貨件數", 0) + merged.get("Non Tourist可調貨存貨件數", 0)).astype(int)
    merged = merged[merged["可調貨總存貨件數"] > 0]

    if merged.empty:
        return pd.DataFrame()

    # 欄位排序
    out_cols = ["SKU"]
    if "Brand" in merged.columns:
        out_cols.append("Brand")
    out_cols += [
        "Tourist缺貨店數", "Tourist_M2銷售件數",
        "Mix可調貨存貨件數", "Mix閒置店數", "Non Tourist可調貨存貨件數", "Non Tourist閒置店數",
        "可調貨總存貨件數",
    ]
    # 排序須在 subset 之前（因排序鍵 Tourist損失_P90D 不列入輸出）
    merged = merged.sort_values(
        ["Tourist損失_P90D", "可調貨總存貨件數"], ascending=[False, False]
    ).reset_index(drop=True)
    merged = merged[out_cols]

    return merged



def _classify_d001_status(d001_cm: float) -> str:
    """依 D001 倉庫 CM 存貨分類缺貨源頭。

    - 「倉庫有貨」(D001 CM > 0): 補貨流程問題，貨在倉庫卻沒送到店
    - 「源頭斷貨」(D001 CM = 0): 倉庫也斷貨，需上游補貨
    """
    return "倉庫有貨" if d001_cm > 0 else "源頭斷貨"



def _build_sku_stock_diagnosis(raw: pd.DataFrame, sku_lookup: pd.DataFrame | None = None) -> pd.DataFrame:
    """by-SKU 缺貨 Review 表（Tourist + Mix + Non Tourist 缺貨 SKU）。

    每列為一個缺貨 SKU，彙總：
    - Tourist / Mix / Non Tourist 各自缺貨店數與銷售
    - D001 倉庫三期貨量 (M-2/M-1/CM) + D001 狀態分類（源頭斷貨 vs 補貨問題）
    - 店舖三期總存貨 (M-2/M-1/CM)
    排序：Tourist 缺貨店數 → Mix 缺貨店數 → Non Tourist 缺貨店數 → Tourist 損失(P90D)（皆降序）。

    效能優化：原本 6 次獨立 groupby + 5 次 merge → 改為 2 次 groupby（SKU-level ＋ SKU×Nature-level）。

    Args:
        sku_lookup: 預先建好的 SKU 基本資訊查找表，為 None 時即時計算。
    """
    if "SKU" not in raw.columns or "Shop Nature" not in raw.columns:
        return pd.DataFrame()

    # 只看 Tourist + Mix + Non Tourist 的缺貨行（CM Shop Stock=0 且有需求）
    cm_stock_vec = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))
    focus_mask = (
        raw["Shop Nature"].isin(["Tourist", "Mix", "Non Tourist"])
        & raw["Is Stock Impacted"]
        & (cm_stock_vec <= 0)
    )
    if not focus_mask.any():
        return pd.DataFrame()

    # 預先建立 SKU 基本資訊查找表（只建一次，之後所有函式可重用）
    if sku_lookup is None:
        sku_lookup = _build_sku_info_lookup(raw)

    focus = raw.loc[focus_mask].copy()

    # 補上需要的數值欄位
    focus["_loss_p90d"] = _safe_numeric(focus.get("Max Sales Loss (P90D)", pd.Series(dtype=float, index=focus.index)))
    focus["_loss_m2"] = _safe_numeric(focus.get("Max Sales Loss (M-2)", pd.Series(dtype=float, index=focus.index)))
    focus["_cm_sold"] = _safe_numeric(focus.get("Sold Qty (CM - MTD)", pd.Series(dtype=float, index=focus.index)))
    focus["_p90d"] = _safe_numeric(focus.get("Shop Monthly Sold Qty (P90D)", pd.Series(dtype=float, index=focus.index)))
    for label, col in [("M-2", "Shop Stock Qty (M-2)"), ("M-1", "Shop Stock Qty (M-1)"), ("CM", "Shop Stock Qty (CM)")]:
        focus[f"_shop_stock_{label}"] = _safe_numeric(focus.get(col, pd.Series(dtype=float, index=focus.index)))
    for label, col in [("M-2", "D001 Stock Qty (M-2)"), ("M-1", "D001 Stock Qty (M-1)"), ("CM", "D001 Stock Qty (CM)")]:
        focus[f"_d001_{label}"] = _safe_numeric(focus.get(col, pd.Series(dtype=float, index=focus.index)))

    # 判斷 Site Code 可用性（用於 nunique）
    has_site_code = "Site Code" in focus.columns

    # ── 單次 groupby 計算所有 SKU-level 指標（D001 + Shop Stock + Brand/Buyer 已從 lookup 取得）──
    sku_agg = focus.groupby("SKU", dropna=False).agg(
        **{
            "D001_Stock_M2": ("_d001_M-2", "first"),
            "D001_Stock_M1": ("_d001_M-1", "first"),
            "D001_Stock_CM": ("_d001_CM", "first"),
            "店舖存貨總_M2": ("_shop_stock_M-2", "sum"),
            "店舖存貨總_M1": ("_shop_stock_M-1", "sum"),
            "店舖存貨總_CM": ("_shop_stock_CM", "sum"),
        }
    ).reset_index()

    # ── 單次 groupby 計算 SKU×Nature-level 指標（取代 3 次獨立 filter+groupby）──
    _DIAG_NATURES = ["Tourist", "Mix", "Non Tourist"]
    nat_agg_dict = {}
    for nat in _DIAG_NATURES:
        shop_count_col = "Site Code" if has_site_code else "SKU"
        nat_agg_dict[f"{nat}缺貨店數"] = (shop_count_col, "nunique") if has_site_code else (shop_count_col, "count")
        nat_agg_dict[f"{nat}損失_M2"] = ("_loss_m2", "sum")
        nat_agg_dict[f"{nat}損失_P90D"] = ("_loss_p90d", "sum")
        nat_agg_dict[f"{nat}_CM銷售件數"] = ("_cm_sold", "sum")
        nat_agg_dict[f"{nat}_P90D總量(90日)"] = ("_p90d", "sum")

    sku_nat_agg = focus.groupby(["SKU", "Shop Nature"], dropna=False, observed=True).agg(
        **{k: v for k, v in nat_agg_dict.items() if v[0] in focus.columns or has_site_code}
    ).reset_index()

    # 將 Nature 展開為寬表（pivot 比 merge 3 次快，避免重複欄位問題）
    merged = sku_agg.copy()
    for nat in _DIAG_NATURES:
        nat_cols = [c for c in [f"{nat}缺貨店數", f"{nat}損失_M2", f"{nat}損失_P90D",
                                 f"{nat}_CM銷售件數", f"{nat}_P90D總量(90日)"] if c in sku_nat_agg.columns]
        if nat_cols:
            nat_sub = sku_nat_agg[sku_nat_agg["Shop Nature"] == nat][["SKU"] + nat_cols].copy()
            if not nat_sub.empty:
                merged = merged.merge(nat_sub, on="SKU", how="left")

    # 補上 SKU 基本資訊（從 lookup，取代原本的 brand_per_sku groupby）
    lookup_cols = ["SKU"] + [c for c in ["Brand", "Buyer", "SKU Desc", "Launch Date", "Is New SKU (3mo)"] if c in sku_lookup.columns]
    merged = merged.merge(sku_lookup[lookup_cols], on="SKU", how="left")

    # NDRF 行數統計（從完整 raw 計算）
    rp_full = (
        raw.get("RP Type", _safe_default_series("", raw.index))
        .fillna("").astype(str).str.strip().str.upper()
    )
    raw_ndrf = raw.assign(_is_nd=(rp_full == "ND"), _is_rf=(rp_full == "RF"))
    ndrf_per_sku = (
        raw_ndrf.groupby("SKU", dropna=False)
        .agg(
            NDRF_ND行數=("_is_nd", "sum"),
            NDRF_RF行數=("_is_rf", "sum"),
        )
        .reset_index()
    )
    merged = merged.merge(ndrf_per_sku, on="SKU", how="left")
    for c in ["NDRF_ND行數", "NDRF_RF行數"]:
        merged[c] = merged[c].fillna(0).astype(int)

    # 確保各 Nature 欄位存在（當 focus 無該 Nature 時補 0）
    _diag_sort_keys = []
    for nat in _DIAG_NATURES:
        for col in [f"{nat}缺貨店數", f"{nat}損失_M2", f"{nat}損失_P90D",
                     f"{nat}_CM銷售件數", f"{nat}_P90D總量(90日)"]:
            if col not in merged.columns:
                merged[col] = 0
            else:
                merged[col] = merged[col].fillna(0).astype(int)
        _diag_sort_keys.append(f"{nat}缺貨店數")
    # 排序鍵：Tourist 損失 P90D（ Tour-specific，供排序用）
    if "Tourist損失_P90D" not in merged.columns:
        merged["Tourist損失_P90D"] = 0
    _diag_sort_keys.append("Tourist損失_P90D")

    # D001 狀態分類
    merged["D001 狀態"] = merged["D001_Stock_CM"].apply(_classify_d001_status)

    # 欄位順序
    out_cols = ["SKU"]
    if "SKU Desc" in merged.columns:
        out_cols.append("SKU Desc")
    if "Launch Date" in merged.columns:
        out_cols.append("Launch Date")
    if "Is New SKU (3mo)" in merged.columns:
        out_cols.append("Is New SKU (3mo)")
    out_cols += ["Brand", "Buyer",
                "NDRF_ND行數", "NDRF_RF行數",
                "Tourist缺貨店數", "Tourist_CM銷售件數", "Tourist_P90D總量(90日)",
                "Mix缺貨店數", "Mix_CM銷售件數", "Mix_P90D總量(90日)",
                "Non Tourist缺貨店數", "Non Tourist_CM銷售件數", "Non Tourist_P90D總量(90日)",
                "D001 狀態", "D001_Stock_M2", "D001_Stock_M1", "D001_Stock_CM",
                "店舖存貨總_M2", "店舖存貨總_M1", "店舖存貨總_CM"]
    # 排序
    merged = merged.sort_values(
        _diag_sort_keys,
        ascending=[False] * len(_diag_sort_keys),
    ).reset_index(drop=True)
    avail = [c for c in out_cols if c in merged.columns]
    merged = merged[avail]

    return merged



# Turnday 分級順序（用於排序；含 RF/ND 拆分）
_TURNDAY_BAND_ORDER = [
    "<7日(過低)", "7-14日(理想)", "14-30日(偏高)", ">30日(過剩)",
    "無銷售有存貨(RF)", "無銷售有存貨(ND)", "無存貨",
]



def _build_inventory_distribution_summary(raw: pd.DataFrame, stock_threshold: int = 1000,
                                            sku_lookup: pd.DataFrame | None = None) -> dict:
    """存貨分佈彙總 — 回傳 dict 含 4 個 DataFrame。

    Args:
        stock_threshold: 低存貨 SKU 門檻（D001 + Shop Stock 總和 < 此值視為低存貨，預設 300）。
        sku_lookup: 預先建好的 SKU 基本資訊查找表，為 None 時即時計算。

    回傳:
        {
            "mix_turnday": Mix Turnday 分級表 (B1a),
            "nt_turnday": Non Tourist Turnday 分級表 (B1b),
            "nature_compare": Nature 存貨對比表 (B2),
            "imbalance_sku": 存貨失衡 SKU 清單 (B3, Tourist缺貨+其他過剩),
            "low_stock_allocation": 低存貨 SKU 配置決策表 (B4),
            "transfer_priority": 可調貨優先分析 (B5, Tourist缺貨+同SKU在Mix/NT有可釋放件數),
        }
    """
    result = {
        "mix_turnday": pd.DataFrame(),
        "nt_turnday": pd.DataFrame(),
        "nature_compare": pd.DataFrame(),
        "imbalance_sku": pd.DataFrame(),
        "low_stock_allocation": pd.DataFrame(),
        "transfer_priority": {},
    }

    if raw.empty or "Shop Nature" not in raw.columns:
        return result

    nature = raw["Shop Nature"].fillna("").astype(str)
    cm_stock = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))
    p90d = _safe_numeric(raw.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, raw.index)))
    turnday = pd.to_numeric(raw.get("Turnday (Days)", _safe_default_series(0, raw.index)), errors="coerce")

    # ── B1: Turnday 分級彙總（Mix + Non Tourist）─────────────────────────────────
    def _build_turnday_table(data, band_col, excess_col=None, label_prefix=""):
        """建構單一 Nature 的 Turnday 分級表。"""
        if data.empty or band_col not in data.columns:
            return pd.DataFrame()
        stock = _safe_numeric(data.get("Shop Stock Qty (CM)", _safe_default_series(0, data.index)))
        excess = _safe_numeric(data.get(excess_col, _safe_default_series(0, data.index))) if excess_col else _safe_default_series(0, data.index)
        band = data[band_col].fillna("無存貨")
        tbl = (
            data.assign(_stock=stock, _excess=excess, _band=band)
            .groupby("_band", dropna=False)
            .agg(行數=("_band", "count"), 存貨件數=("_stock", "sum"), 可釋放件數=("_excess", "sum"))
            .reset_index()
            .rename(columns={"_band": "Turnday 區間"})
        )
        tbl["__order"] = tbl["Turnday 區間"].map({b: i for i, b in enumerate(_TURNDAY_BAND_ORDER)}).fillna(99)
        tbl = tbl.sort_values("__order").drop(columns=["__order"]).reset_index(drop=True)
        for c in ["行數", "存貨件數", "可釋放件數"]:
            tbl[c] = tbl[c].fillna(0).astype(int)
        total = pd.DataFrame([{
            "Turnday 區間": "總計",
            "行數": int(tbl["行數"].sum()),
            "存貨件數": int(tbl["存貨件數"].sum()),
            "可釋放件數": int(tbl["可釋放件數"].sum()),
        }])
        return pd.concat([tbl, total], ignore_index=True)

    # B1a: Mix
    mix_mask = nature == "Mix"
    if mix_mask.any() and "Mix Turnday Band" in raw.columns:
        result["mix_turnday"] = _build_turnday_table(raw[mix_mask], "Mix Turnday Band", excess_col="Mix Excess Stock")

    # B1b: Non Tourist
    nt_mask = nature == "Non Tourist"
    if nt_mask.any() and "Non Tourist Turnday Band" in raw.columns:
        result["nt_turnday"] = _build_turnday_table(raw[nt_mask], "Non Tourist Turnday Band", excess_col="Non Tourist Excess Stock")

    # ── B2: Nature 存貨分佈對比 ───────────────────────────────────────────────
    b2_rows = []
    for nat in ["Tourist", "Mix", "Non Tourist"]:
        sub_mask = nature == nat
        if not sub_mask.any():
            continue
        sub_stock = cm_stock[sub_mask]
        sub_td = turnday[sub_mask]
        # 缺貨行數 (CM Stock=0)
        oos_count = int((cm_stock[sub_mask] == 0).sum())
        # 過剩行數 (>30日)
        over30 = int((sub_td > 30).sum())
        b2_rows.append({
            "Shop Nature": nat,
            "總行數": int(sub_mask.sum()),
            "存貨總量(件)": int(sub_stock.sum()),
            "缺貨行數(CM=0)": oos_count,
            "平均Turnday(日)": round(float(sub_td.mean()), 1) if sub_td.notna().any() else 0,
            "過剩行數(>30日)": over30,
        })
    if b2_rows:
        result["nature_compare"] = pd.DataFrame(b2_rows)

    # ── B3: 存貨失衡 SKU 清單 (Tourist缺貨 + 其他Nature過剩) ──────────────────
    # Tourist 缺貨且有需求的 SKU（Is Stock Impacted 子集）
    if "Is Stock Impacted" in raw.columns:
        tourist_oos = raw[(nature == "Tourist") & raw["Is Stock Impacted"] & (cm_stock <= 0)]
    else:
        tourist_oos = raw[(nature == "Tourist") & (cm_stock <= 0) & (p90d > 0)]

    # 其他 Nature 過剩存貨：Mix>14日（Mix Excess Stock>0）或 Non Tourist>14日（Non Tourist Excess Stock>0）
    # 統一到相同 14 日標準，與 Turnday 分級邏輯一致
    mix_excess_vec = _safe_numeric(raw.get("Mix Excess Stock", _safe_default_series(0, raw.index)))
    nt_excess_vec = _safe_numeric(raw.get("Non Tourist Excess Stock", _safe_default_series(0, raw.index)))
    other_over = raw[
        (nature != "Tourist")
        & (
            ((nature == "Mix") & (mix_excess_vec > 0))
            | ((nature == "Non Tourist") & (nt_excess_vec > 0))
        )
    ]

    if tourist_oos.empty or other_over.empty or "SKU" not in raw.columns:
        result["imbalance_sku"] = pd.DataFrame()  # 無失衡 SKU 時回傳空
    else:
        # Tourist 缺貨損失（by SKU）
        t_loss = _safe_numeric(tourist_oos.get("Max Sales Loss (P90D)", _safe_default_series(0, tourist_oos.index)))
        tourist_oos = tourist_oos.assign(_loss=t_loss)
        t_agg = (
        tourist_oos.groupby(["SKU"] + (["Brand"] if "Brand" in tourist_oos.columns else []), dropna=False)
        .agg(
            Tourist缺貨店數=("SKU", "count"),
            Tourist缺貨損失_P90D=("_loss", "sum"),
        )
        .reset_index()
        )

        # 其他 Nature 過剩可釋放存貨（by SKU × Nature）
        other_over = other_over.assign(_stock=cm_stock[other_over.index])
        o_agg = (
        other_over.groupby(["SKU", "Shop Nature"], dropna=False)
        .agg(
            過剩存貨=("_stock", "sum"),
            過剩店數=("SKU", "count"),
        )
        .reset_index()
        )
        o_pivot_stock = o_agg.pivot_table(index="SKU", columns="Shop Nature", values="過剩存貨", fill_value=0).reset_index()
        o_pivot_shops = o_agg.pivot_table(index="SKU", columns="Shop Nature", values="過剩店數", fill_value=0).reset_index()

        merged = t_agg.merge(o_pivot_stock, on="SKU", how="inner")
        # 重新命名過剩存貨欄
        rename_map = {c: f"{c}過剩存貨" for c in o_pivot_stock.columns if c != "SKU"}
        merged = merged.rename(columns=rename_map)

        # 確保 Mix/Non Tourist 過剩欄存在
        for n in ["Mix", "Non Tourist"]:
            col = f"{n}過剩存貨"
            if col not in merged.columns:
                merged[col] = 0
            merged[col] = merged[col].fillna(0).astype(int)

        merged["其他Nature過剩總存貨"] = (merged["Mix過剩存貨"] + merged["Non Tourist過剩存貨"]).astype(int)
        merged = merged[merged["其他Nature過剩總存貨"] > 0]
        # 可調貨建議量 = min(Tourist 缺貨損失, 其他過剩存貨)
        merged["建議調貨量"] = np.minimum(merged["Tourist缺貨損失_P90D"], merged["其他Nature過剩總存貨"]).astype(int)

        # 附加 SKU 級資訊：SKU Desc / Launch Date / Buyer（從 lookup 讀取，避免重複 groupby）
        if sku_lookup is None:
            sku_lookup = _build_sku_info_lookup(raw)
        info_from_lookup = [c for c in ["SKU Desc", "Launch Date", "Is New SKU (3mo)", "Buyer"]
                           if c in sku_lookup.columns and c not in merged.columns]
        if info_from_lookup:
            merged = merged.merge(sku_lookup[["SKU"] + info_from_lookup], on="SKU", how="left")

        # D001 Net Stock + Shop Stock 總數（不在 lookup 中，獨立計算）
        d001_net = "D001 Net Stock Qty (CM)"
        has_d001_net = d001_net in raw.columns
        has_shop_cm = "Shop Stock Qty (CM)" in raw.columns
        if has_d001_net and has_shop_cm:
            extra_agg = (
                raw.groupby("SKU", dropna=False)
                .agg(_d001_net=(d001_net, "first"), _shop_total=("Shop Stock Qty (CM)", "sum"))
                .reset_index()
            )
            merged = merged.merge(extra_agg, on="SKU", how="left")
            merged["總可用存貨"] = (_safe_numeric(merged["_d001_net"]) + _safe_numeric(merged["_shop_total"])).astype(int)
            merged.drop(columns=["_d001_net", "_shop_total"], inplace=True)

        out_cols = ["SKU"]
        if "Brand" in merged.columns:
                out_cols.append("Brand")
        if "SKU Desc" in merged.columns:
            out_cols.append("SKU Desc")
        if "Launch Date" in merged.columns:
            out_cols.append("Launch Date")
        if "Is New SKU (3mo)" in merged.columns:
            out_cols.append("Is New SKU (3mo)")
        if "Buyer" in merged.columns:
            out_cols.append("Buyer")
        if "總可用存貨" in merged.columns:
            out_cols.append("總可用存貨")
        out_cols += ["Tourist缺貨店數", "Mix過剩存貨", "Non Tourist過剩存貨", "其他Nature過剩總存貨", "建議調貨量"]
        merged = merged[[c for c in out_cols if c in merged.columns]]
        merged = merged.sort_values(["Tourist缺貨店數", "建議調貨量"], ascending=[False, False]).reset_index(drop=True)

        # 附註解（計算方式說明）
        merged.attrs["annotation"] = (
            "【計算方式】\n"
            "Mix過剩存貨：Mix 店舖 Shop Stock (CM) > 0 且 Turnday > 14 日的過剩部分 "
            "（超出 P90D/90×14 標準的存貨總和）。\n"
            "Non Tourist過剩存貨：Non Tourist 店舖 Shop Stock (CM) > 0 且 Turnday > 30 日的存貨總和。\n"
            "建議調貨量 = min(Tourist 缺貨店數對應的需求量, Mix過剩存貨 + Non Tourist過剩存貨)。"
        )
        result["imbalance_sku"] = merged

    # ── B4: 低存貨 SKU 配置決策（D001 + Shop Stock < 門檻）─────────────────────
    # 找出總可用存貨不足門檻的 SKU，分析是否該減少 Non Tourist RF 店舖。
    # 同時統計 RF 店「無銷售有存貨」與 ND 店存貨（有/無銷售）的浪費情況。
    if "SKU" in raw.columns:
        d001_cm_col = "D001 Stock Qty (CM)" if "D001 Stock Qty (CM)" in raw.columns else None
        d001_cm = _safe_numeric(raw.get(d001_cm_col or "SKU", _safe_default_series(0, raw.index)))
        if d001_cm_col is None:
            d001_cm = pd.Series(0, index=raw.index)  # 無 D001 欄位時視為 0
        low_df = raw.assign(_d001=d001_cm, _shop_stock=cm_stock, _p90d=p90d)
        # 每個 SKU 的總可用存貨
        sku_avail = (
            low_df.groupby("SKU", dropna=False)
            .agg(總可用存貨=("_d001", "first"), _shop_total=("_shop_stock", "sum"))
            .reset_index()
        )
        sku_avail["總可用存貨"] = sku_avail["總可用存貨"] + sku_avail["_shop_total"]
        low_skus = set(sku_avail[sku_avail["總可用存貨"] < stock_threshold]["SKU"])

        if low_skus:
            low = low_df[low_df["SKU"].isin(low_skus)].copy()
            rp = low.get("RP Type", _safe_default_series("", low.index)).fillna("").astype(str).str.strip().str.upper()

            # Tourist/Mix 缺貨且有需求（RF）— 代表真的不夠貨
            impacted_mask = low.get("Is Stock Impacted", pd.Series([False] * len(low), index=low.index))
            t_oos = int(((low["Shop Nature"] == "Tourist") & (impacted_mask)).sum()) if "Shop Nature" in low.columns else 0
            m_oos = int(((low["Shop Nature"] == "Mix") & (impacted_mask)).sum()) if "Shop Nature" in low.columns else 0

            # Non Tourist RF 店「無銷售有存貨」— 浪費，建議減店
            nt_rf = low[(low["Shop Nature"] == "Non Tourist") & (rp == "RF")] if "Shop Nature" in low.columns else pd.DataFrame()
            nt_waste_mask = (nt_rf["_p90d"] == 0) & (nt_rf["_shop_stock"] > 0)
            nt_waste_rows = int(nt_waste_mask.sum())
            nt_waste_stock = int(nt_rf.loc[nt_waste_mask, "_shop_stock"].sum())

            # ND 店存貨統計（有/無銷售）— 計畫淘汰卻佔存貨
            nd_all = low[rp == "ND"]
            nd_has_stock = nd_all[nd_all["_shop_stock"] > 0]
            nd_stock_total = int(nd_has_stock["_shop_stock"].sum())
            nd_sale = nd_has_stock[nd_has_stock["_p90d"] > 0]
            nd_nosale = nd_has_stock[nd_has_stock["_p90d"] == 0]
            nd_sale_rows = int(len(nd_sale))
            nd_nosale_rows = int(len(nd_nosale))
            nd_nosale_stock = int(nd_nosale["_shop_stock"].sum())

            b4 = pd.DataFrame([{
                "低存貨門檻(件)": stock_threshold,
                "低存貨SKU數": len(low_skus),
                "Tourist缺貨行數": t_oos,
                "Mix缺貨行數": m_oos,
                "NT_RF無銷售有存貨行數": nt_waste_rows,
                "NT_RF無銷售佔用存貨(件)": nt_waste_stock,
                "ND有存貨行數": nd_sale_rows + nd_nosale_rows,
                "ND存貨總量(件)": nd_stock_total,
                "ND有存貨有銷售行數": nd_sale_rows,
                "ND有存貨無銷售行數": nd_nosale_rows,
                "ND無銷售浪費存貨(件)": nd_nosale_stock,
            }])
            result["low_stock_allocation"] = b4

    # ── B5: 可調貨優先分析 (Tourist缺貨 + 同SKU在Mix/NT有可釋放件數) ──────────
    result["transfer_priority"] = _build_transfer_priority_analysis(raw, sku_lookup)

    return result



def _build_transfer_priority_analysis(raw: pd.DataFrame, sku_lookup: pd.DataFrame | None = None) -> dict:
    """可調貨優先分析：Tourist 缺貨行中，同 SKU 在 Mix/Non Tourist 有可釋放件數的行數及 SKU 明細。

    回傳:
        {
            "summary": DataFrame with aggregate metrics (可調貨Tourist行數, 可調貨SKU數),
            "detail": DataFrame with per-SKU transfer priority details,
        }
    """
    result: dict = {
        "summary": pd.DataFrame(),
        "detail": pd.DataFrame(),
    }

    if raw.empty or "Shop Nature" not in raw.columns or "SKU" not in raw.columns:
        return result

    nature = raw["Shop Nature"].fillna("").astype(str)
    cm_stock = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))

    # 1. Tourist OOS rows (Is Stock Impacted + CM stock <= 0)
    is_impacted = raw.get("Is Stock Impacted", pd.Series(False, index=raw.index))
    tourist_oos_mask = (nature == "Tourist") & is_impacted & (cm_stock <= 0)
    if not tourist_oos_mask.any():
        return result

    tourist_oos = raw[tourist_oos_mask].copy()

    # 2. Aggregate Mix/NT releasable stock by SKU（可釋放件數 = Mix Excess Stock + Non Tourist Excess Stock）
    mix_excess = _safe_numeric(raw.get("Mix Excess Stock", _safe_default_series(0, raw.index)))
    nt_excess = _safe_numeric(raw.get("Non Tourist Excess Stock", _safe_default_series(0, raw.index)))

    sku_releasable = (
        raw.assign(_mix_excess=mix_excess, _nt_excess=nt_excess)
        .groupby("SKU", dropna=False)
        .agg(
            Mix可釋放件數=("_mix_excess", "sum"),
            NonTourist可釋放件數=("_nt_excess", "sum"),
        )
        .reset_index()
    )
    sku_releasable["可釋放總件數"] = sku_releasable["Mix可釋放件數"] + sku_releasable["NonTourist可釋放件數"]
    skus_with_releasable = set(sku_releasable[sku_releasable["可釋放總件數"] > 0]["SKU"])

    if not skus_with_releasable:
        return result

    # 3. Filter Tourist OOS to only those SKUs with releasable stock
    transferrable_mask = tourist_oos["SKU"].isin(skus_with_releasable)
    if not transferrable_mask.any():
        return result

    transferrable_oos = tourist_oos[transferrable_mask]

    # 4. Count metrics
    row_count = len(transferrable_oos)
    sku_count = transferrable_oos["SKU"].nunique()

    summary = pd.DataFrame([{
        "可調貨Tourist行數": row_count,
        "可調貨SKU數": sku_count,
    }])
    result["summary"] = summary

    # 5. Build detail table: per SKU
    # Tourist缺失件數 = max(0, Adj. Safety Stock - Shop Stock Qty (CM))
    adj_safety = _safe_numeric(transferrable_oos.get("Adj. Safety Stock", _safe_default_series(0, transferrable_oos.index)))
    shop_stock = _safe_numeric(transferrable_oos.get("Shop Stock Qty (CM)", _safe_default_series(0, transferrable_oos.index)))
    shortage = np.maximum(0, adj_safety.values - shop_stock.values)
    transferrable_oos = transferrable_oos.assign(_shortage=shortage)

    t_agg = (
        transferrable_oos.groupby("SKU", dropna=False)
        .agg(
            Tourist缺貨行數=("SKU", "count"),
            Tourist缺失件數=("_shortage", "sum"),
        )
        .reset_index()
    )

    # Merge with releasable stock info
    detail = t_agg.merge(sku_releasable, on="SKU", how="left")
    detail = detail[detail["可釋放總件數"] > 0]

    # Add SKU-level info: Brand, SKU Desc, Buyer（從 lookup 讀取，避免重複 groupby）
    if sku_lookup is None:
        sku_lookup = _build_sku_info_lookup(raw)
    info_from_lookup = [c for c in ["Brand", "SKU Desc", "Buyer"]
                       if c in sku_lookup.columns and c not in detail.columns]
    if info_from_lookup:
        detail = detail.merge(sku_lookup[["SKU"] + info_from_lookup], on="SKU", how="left")

    # D001 status
    d001_cm_col = "D001 Stock Qty (CM)"
    if d001_cm_col in raw.columns:
        d001_per_sku = (
            raw.groupby("SKU", dropna=False)
            .agg(D001_Stock_CM=(d001_cm_col, "first"))
            .reset_index()
        )
        d001_per_sku["D001狀態"] = d001_per_sku["D001_Stock_CM"].apply(_classify_d001_status)
        detail = detail.merge(d001_per_sku[["SKU", "D001狀態"]], on="SKU", how="left")
    else:
        detail["D001狀態"] = "—"

    # Column ordering
    out_cols = ["SKU"]
    for _c in ["Brand", "SKU Desc", "Buyer"]:
        if _c in detail.columns:
            out_cols.append(_c)
    out_cols += ["Tourist缺貨行數", "Tourist缺失件數", "Mix可釋放件數", "NonTourist可釋放件數", "可釋放總件數", "D001狀態"]

    detail = detail[[c for c in out_cols if c in detail.columns]]

    # Sort: Tourist OOS count desc, then releasable total desc
    detail = detail.sort_values(["Tourist缺貨行數", "可釋放總件數"], ascending=[False, False]).reset_index(drop=True)

    # Cast int columns
    for c in ["Tourist缺貨行數", "Tourist缺失件數", "Mix可釋放件數", "NonTourist可釋放件數", "可釋放總件數"]:
        if c in detail.columns:
            detail[c] = detail[c].fillna(0).astype(int)

    result["detail"] = detail

    return result



_OOS_DETAIL_COLS = [
    "Buyer", "Brand", "SKU", "SKU Desc", "Item Class", "Launch Date", "Is New SKU (3mo)",
    "Site Code", "Shop Name", "OM",
    "Shop Nature", "Supply Source", "NDRF Status",
    "Shop Stock Qty (M-2)", "Shop Stock Qty (M-1)",
    "Shop Stock Qty (CM)", "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
    "Shop Monthly Sold Qty (P90D)", "Is OOS (M-2)", "Is OOS (M-1)", "Is OOS (CM)",
    "OOS Periods Count", "OOS Pattern",
]



# 對齊 Joseph 檔的 RAW 欄位順序（38 欄，不含 NDRF Status，含 Total Stock 等擴充欄位）
_JOSEPH_RAW_COLS = [
    "Buyer", "Brand", "SKU", "SKU Desc", "Item Class", "Launch Date", "Is New SKU (3mo)",
    "Site Code", "Shop Name", "OM",
    "Shop Nature", "Supply Source",
    "Total Stock Qty (M-2)", "Total Stock Qty (M-1)", "Total Stock Qty (CM)",
    "Shop Stock Qty (M-2)", "Shop Stock Qty (M-1)", "Shop Stock Qty (CM)",
    "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
    "Shop Monthly Sold Qty (P90D)", "Is OOS (M-2)", "Is OOS (M-1)", "Is OOS (CM)",
    "OOS Periods Count", "OOS Pattern",
    "有貨店舖 (M-2)", "有貨店舖 (M-1)", "有貨店舖 (CM)",
    "有賣店舖 (M-2)", "有賣店舖 (M-1)", "有賣店舖 (CM)",
    "Eff Price", "Price Tier",
    "SOH Tier (M-2)", "SOH Tier (M-1)", "SOH Tier (CM)",
]



# SOH Tier 邊界（左閉右開），Joseph 分級：a 0(=0), b 1, c 2, d 3-4, e 5-7, f 8-14, g 15-49, h 50+
_SOH_TIER_BINS = [-1, 0, 1, 2, 4, 7, 14, 49, 10**9]



_SOH_TIER_LABELS = ["a 0", "b 1", "c 2", "d 3-4", "e 5-7", "f 8-14", "g 15-49", "h 50+"]



# Price Tier 邊界（下限值），Joseph 分級
_PRICE_TIER_BOUNDS = [0, 100, 300, 500, 800, 1200]



_PRICE_TIER_LABELS = ["a below 100", "b 100-300", "c 300-500", "d 500-800", "e 800-1200", "f 1200+"]



def _soh_tier_label(qty: int | float) -> str | None:
    """將 Shop Stock Qty 轉為 SOH tier 標籤（a 0 … h 50+），NaN/None → None。"""
    if pd.isna(qty):
        return None
    qty = int(qty)
    for i in range(len(_SOH_TIER_BINS) - 1):
        if _SOH_TIER_BINS[i] < qty <= _SOH_TIER_BINS[i + 1]:
            return _SOH_TIER_LABELS[i]
    return _SOH_TIER_LABELS[-1]



def _price_tier_label(eff_price: float) -> str | None:
    """將 Eff Price 轉為 Price Tier 標籤，NaN → None。"""
    if pd.isna(eff_price) or eff_price < 0:
        return None
    for i in range(len(_PRICE_TIER_BOUNDS) - 1):
        lo = _PRICE_TIER_BOUNDS[i]
        hi = _PRICE_TIER_BOUNDS[i + 1]
        if lo <= eff_price < hi:
            return _PRICE_TIER_LABELS[i]
    return _PRICE_TIER_LABELS[-1]  # ≥1200 → f 1200+



def _enrich_raw_joseph(raw: pd.DataFrame) -> pd.DataFrame:
    """在 raw_df 上計算 Joseph 額外欄位的 copy（不修改原文）。"""
    df = raw.copy()
    idx = df.index

    # 1. Total Stock Qty = D001 Stock Qty + groupby(SKU) sum of Shop Stock Qty
    for period, d001_col, shop_col in [
        ("M-2", "D001 Stock Qty (M-2)", "Shop Stock Qty (M-2)"),
        ("M-1", "D001 Stock Qty (M-1)", "Shop Stock Qty (M-1)"),
        ("CM", "D001 Stock Qty (CM)", "Shop Stock Qty (CM)"),
    ]:
        col = f"Total Stock Qty ({period})"
        d001 = _safe_numeric(df.get(d001_col, _safe_default_series(0, idx)))
        shop = _safe_numeric(df.get(shop_col, _safe_default_series(0, idx)))
        sku_total = shop.groupby(df["SKU"], dropna=False).transform("sum")
        df[col] = d001 + sku_total

    # 2. 有貨店舖 / 有賣店舖（0/1 flag）
    for period, shop_col, sold_col in [
        ("M-2", "Shop Stock Qty (M-2)", "Sold Qty (M-2)"),
        ("M-1", "Shop Stock Qty (M-1)", "Sold Qty (M-1)"),
        ("CM", "Shop Stock Qty (CM)", "Sold Qty (CM - MTD)"),
    ]:
        shop = _safe_numeric(df.get(shop_col, _safe_default_series(0, idx)))
        df[f"有貨店舖 ({period})"] = (shop > 0).astype(int)
        sold = _safe_numeric(df.get(sold_col, _safe_default_series(0, idx)))
        df[f"有賣店舖 ({period})"] = (sold > 0).astype(int)

    # 3. Eff Price (SKU 級，merged 回每行)
    ns2 = _safe_numeric(df.get("Net Sales (M-2)", _safe_default_series(0, idx)))
    ns1 = _safe_numeric(df.get("Net Sales (M-1)", _safe_default_series(0, idx)))
    nsm = _safe_numeric(df.get("Net Sales (CM - MTD)", _safe_default_series(0, idx)))
    sq2 = _safe_numeric(df.get("Sold Qty (M-2)", _safe_default_series(0, idx)))
    sq1 = _safe_numeric(df.get("Sold Qty (M-1)", _safe_default_series(0, idx)))
    sqm = _safe_numeric(df.get("Sold Qty (CM - MTD)", _safe_default_series(0, idx)))
    ep_data = pd.DataFrame({
        "SKU": df["SKU"],
        "_ns": ns2 + ns1 + nsm,
        "_sq": sq2 + sq1 + sqm,
    })
    ep_agg = ep_data.groupby("SKU", dropna=False).agg(
        _ns_sum=("_ns", "sum"), _sq_sum=("_sq", "sum"),
    ).reset_index()
    # 分母 0 → NaN
    ep_agg["Eff Price"] = np.where(ep_agg["_sq_sum"] > 0, ep_agg["_ns_sum"] / ep_agg["_sq_sum"], np.nan)
    df = df.merge(ep_agg[["SKU", "Eff Price"]], on="SKU", how="left")

    # 4. Price Tier
    df["Price Tier"] = df["Eff Price"].apply(_price_tier_label)

    # 5. SOH Tier 逐期
    for period in ["M-2", "M-1", "CM"]:
        shop = _safe_numeric(df.get(f"Shop Stock Qty ({period})", _safe_default_series(0, idx)))
        df[f"SOH Tier ({period})"] = shop.apply(_soh_tier_label)

    return df



# ── 平行化輔助 ──────────────────────────────────────────────────────────────────
# 批次預處理數值欄位，避免各 summary 函式重複呼叫 _safe_numeric
_STOCK_IMPACT_NUMERIC_COLS = [
    "Shop Stock Qty (M-2)", "Shop Stock Qty (M-1)", "Shop Stock Qty (CM)",
    "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
    "Shop Monthly Sold Qty (P90D)", "D001 Stock Qty (CM)",
    "D001 Stock Qty (M-2)", "D001 Stock Qty (M-1)",
    "Max Sales Loss (M-2)", "Max Sales Loss (P90D)",
    "Mix Excess Stock", "Non Tourist Excess Stock",
]



def _preprocess_numeric_columns(raw: pd.DataFrame) -> pd.DataFrame:
    """一次性轉換所有後續會用到的數值欄位。"""
    for _col in _STOCK_IMPACT_NUMERIC_COLS:
        if _col in raw.columns:
            raw[_col] = pd.to_numeric(raw[_col], errors="coerce").fillna(0)
    return raw



def _build_summary_safe(raw: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    """安全的摘要建構：檢查欄位是否存在。"""
    if not group_cols or group_cols[0] not in raw.columns:
        return pd.DataFrame()
    return _build_stock_impact_summary(raw, group_cols)



def transform_stock_impact(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
    warehouse_only: bool = True,
    stock_threshold: int = 1000,
    ref_date: datetime.date | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, dict, list[str]]:
    """Stock Impact Analysis 主流程。

    Args:
        warehouse_only: 若 True（預設），只分析 Supply Source = Warehouse 的列。
        ref_date: Launch Date 計算 New SKU 的參考日期（預設今天，建議傳入 MTD 截止日期）。
    Returns:
        (summary_site, summary_brand, summary_buyer, summary_category, sku_allocation, sku_diagnosis, inv_distribution, sku_monthly_analysis, raw_df, warnings)
    """
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")

    warnings = []

    _optional_cols = [
        "Shop Stock Qty (M-2)", "Shop Stock Qty (M-1)", "Shop Stock Qty (CM)",
        "Sold Qty (M-2)", "Sold Qty (M-1)", "Sold Qty (CM - MTD)",
        "Shop Monthly Sold Qty (P90D)", "Supply Source", "Shop Nature", "Adj. MCH1",
    ]
    for _col in _optional_cols:
        if _col not in df_input.columns:
            warnings.append(f"缺少可選欄位「{_col}」，相關統計將受影響")

    if progress_callback:
        progress_callback(0.05, "驗證通過")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據")

    # 預設只分析 Warehouse SKU
    if warehouse_only and "Supply Source" in df.columns:
        before = len(df)
        df = df[df["Supply Source"].astype(str).str.strip() == "Warehouse"].copy()
        if before > 0 and len(df) == 0:
            warnings.append("勾選「只分析 Warehouse SKU」但來源檔案無 Warehouse 列，已自動取消篩選")
            df = _clean_dataframe(df_input, shop_natures=shop_natures)

    df["SKU"] = df["SKU"].apply(_safe_sku)

    if "SKU Desc" not in df.columns:
        df = _compute_raw_extra_columns(df)

    if progress_callback:
        progress_callback(0.2, "計算缺貨影響指標中...")

    raw = _add_stock_impact_columns(df)

    # 新增 Launch Date / Is New SKU 資訊（於所有 Sheet 共用）
    raw = _add_launch_info(raw, ref_date=ref_date)

    # 批次預處理數值欄位（避免各 summary 函式重複轉換）
    raw = _preprocess_numeric_columns(raw)

    # groupby 加速：將 groupby 欄位轉為 category
    _groupby_cols = ["Brand", "Buyer", "Item Class"]
    if "Adj. MCH1" in raw.columns:
        _groupby_cols.append("Adj. MCH1")
    for _gc in _groupby_cols:
        if _gc in raw.columns:
            raw[_gc] = raw[_gc].astype("category")

    # 注意：stock_threshold 不再用於全域過濾，只會在 low_stock_allocation（B4）
    # 中內部使用，分析所有低存貨 SKU 的配置決策。
    if progress_callback:
        progress_callback(0.4, "建構多維度彙總中...")

    # Site 層級已取消（v1.6.2）
    summary_site = pd.DataFrame()

    # 平行化：同時計算 4 個維度的摘要表（彼此獨立）
    _summary_tasks = {}
    if "Brand" in raw.columns:
        _summary_tasks["brand"] = ["Brand"]
    if "Buyer" in raw.columns:
        _summary_tasks["buyer"] = ["Buyer"]
    if "Adj. MCH1" in raw.columns:
        _summary_tasks["category"] = ["Adj. MCH1"]
    else:
        warnings.append("缺少「Adj. MCH1」欄位，無法產出 By Category 分析")
    if "Item Class" in raw.columns:
        _summary_tasks["item_class"] = ["Item Class"]
    else:
        warnings.append("缺少「Item Class」欄位，無法產出 By Item Class 分析")

    if _summary_tasks:
        with ThreadPoolExecutor(max_workers=min(4, len(_summary_tasks))) as _executor:
            _futures = {
                _key: _executor.submit(_build_summary_safe, raw, _cols)
                for _key, _cols in _summary_tasks.items()
            }
            _results = {key: fut.result() for key, fut in _futures.items()}
    else:
        _results = {}

    summary_brand = _results.get("brand", pd.DataFrame())
    summary_buyer = _results.get("buyer", pd.DataFrame())
    summary_category = _results.get("category", pd.DataFrame())
    summary_item_class = _results.get("item_class", pd.DataFrame())

    if progress_callback:
        progress_callback(0.55, "建構 SKU 層級分析中...")

    # 建立 SKU 基本資訊查找表（供 4 個 SKU-level 函式共用，避免重複 groupby）
    _sku_lookup = _build_sku_info_lookup(raw)

    # 平行化：同時計算 4 個 SKU-level 分析（彼此獨立，共用 _sku_lookup）
    with ThreadPoolExecutor(max_workers=4) as _sku_executor:
        f_alloc = _sku_executor.submit(_build_sku_allocation_summary, raw, _sku_lookup)
        f_diag = _sku_executor.submit(_build_sku_stock_diagnosis, raw, _sku_lookup)
        f_inv = _sku_executor.submit(_build_inventory_distribution_summary, raw, stock_threshold)
        f_monthly = _sku_executor.submit(_build_sku_monthly_analysis, raw, stock_threshold, _sku_lookup)

    sku_allocation = f_alloc.result()
    sku_diagnosis = f_diag.result()
    inv_distribution = f_inv.result()
    sku_monthly_analysis = f_monthly.result()

    if progress_callback:
        progress_callback(0.75, "轉換完成")

    return summary_site, summary_brand, summary_buyer, summary_category, summary_item_class, sku_allocation, sku_diagnosis, inv_distribution, sku_monthly_analysis, raw, warnings



def _write_oos_detail_sheet(ws, raw_df, header_font, header_fill, thin_border, progress_callback):
    """寫入缺貨影響明細（含 ND/RF 行數統計），按缺貨期數降序。"""
    idx = raw_df.index
    # 擴展篩選：任一期缺貨且有需求證據（不限 RF）
    oos_count = _safe_numeric(raw_df.get("OOS Periods Count", _safe_default_series(0, idx)))
    m2_sold = _safe_numeric(raw_df.get("Sold Qty (M-2)", _safe_default_series(0, idx)))
    p90d = _safe_numeric(raw_df.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, idx)))
    has_oos = oos_count.values > 0
    has_demand = (m2_sold.values > 0) | (p90d.values > 0)

    detail = raw_df[has_oos & has_demand].copy()
    if detail.empty:
        ws.cell(row=1, column=1, value="無缺貨影響資料").font = Font(size=12)
        return

    # NDRF Status（顯示 RF 或 ND）
    rp_type = (
        detail.get("RP Type", _safe_default_series("", detail.index))
        .fillna("").astype(str).str.strip().str.upper()
    )
    detail["NDRF Status"] = rp_type

    # 排序：缺貨期數降序 → M-2 銷售降序
    detail["_oos"] = _safe_numeric(detail.get("OOS Periods Count", _safe_default_series(0, detail.index)))
    detail["_m2"] = _safe_numeric(detail.get("Sold Qty (M-2)", _safe_default_series(0, detail.index)))
    detail = detail.sort_values(["_oos", "_m2"], ascending=[False, False])

    avail_cols = [c for c in _OOS_DETAIL_COLS if c in detail.columns]
    detail = detail[avail_cols]

    # 標題
    ws.cell(row=1, column=1, value="OOS Impact Detail — 缺貨銷售影響明細（按缺貨期數降序，含 NDRF Status）").font = Font(bold=True, size=12, color="1E40AF")

    headers = list(detail.columns)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    total_rows = len(detail)
    # P0：共用對齊取代逐格新建 Alignment；值與格式語意不變（已使用 itertuples）
    center_align = Alignment(horizontal="center", vertical="center")
    for row_idx, row_data in enumerate(detail.itertuples(index=False), start=4):
        for col_idx, val in enumerate(row_data, 1):
            if pd.isna(val):
                val = None
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = center_align
            if isinstance(val, bool):
                cell.value = "✓" if val else ""
            elif isinstance(val, (int, float)):
                cell.number_format = "#,##0"
        if progress_callback and ((row_idx - 4) % 2000 == 0 or row_idx - 3 == total_rows):
            progress_callback(0.1 + 0.15 * ((row_idx - 3) / max(total_rows, 1)),
                              f"寫入缺貨明細 ({row_idx-3}/{total_rows})...")

    _auto_width(ws)



def _write_stock_impact_summary_sheet(ws, summary_df, sheet_title, header_font, header_fill, thin_border, progress_callback, progress_start, progress_end):
    ws.cell(row=1, column=1, value=sheet_title).font = Font(bold=True, size=12, color="1E40AF")

    headers = list(summary_df.columns)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    total_rows = len(summary_df)
    # P0：itertuples＋共用對齊／字型取代 iterrows＋逐格新建；值與格式語意不變
    center_align = Alignment(horizontal="center", vertical="center")
    total_bold = Font(bold=True, size=11)
    for row_idx, _vals in enumerate(_iter_header_rows(summary_df, headers)):
        excel_row = 4 + row_idx
        is_total = any(str(v) == "總計" for v in _vals if isinstance(v, str))
        for col_idx, h in enumerate(headers, 1):
            val = _vals[col_idx - 1]
            if pd.isna(val):
                val = None
            cell = ws.cell(row=excel_row, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = center_align
            if is_total:
                cell.font = total_bold
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0"
        if progress_callback and (row_idx % 500 == 0 or row_idx == total_rows - 1):
            p = progress_start + (progress_end - progress_start) * ((row_idx + 1) / max(total_rows, 1))
            progress_callback(p, f"寫入 {sheet_title} ({row_idx+1}/{total_rows})...")

    _auto_width(ws)



def _write_sku_stock_diagnosis_sheet(ws, diag_df, header_font, header_fill, thin_border, progress_callback):
    """寫入 by-SKU 缺貨 Review 表（Tourist+Mix 缺貨 SKU，含 D001 貨量與店舖貨量）。

    D001 狀態分類用顏色標示：「倉庫有貨」(補貨問題) 用橘色警示、「源頭斷貨」用灰色。
    """
    ws.cell(row=1, column=1, value="SKU Stock Diagnosis — by-SKU 缺貨 Review（Tourist + Mix + Non Tourist 缺貨 SKU｜含 D001 倉庫貨量與店舖貨量）").font = Font(bold=True, size=12, color="1E40AF")

    if diag_df is None or diag_df.empty:
        ws.cell(row=3, column=1, value="無 Tourist/Mix 缺貨 SKU").font = Font(size=11, italic=True, color="6B7280")
        return

    headers = list(diag_df.columns)
    # 找出 D001 狀態欄的索引（用於著色）
    d001_status_col_idx = headers.index("D001 狀態") + 1 if "D001 狀態" in headers else None

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    restock_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")  # 淡橘: 補貨問題
    source_fill = PatternFill(start_color="E5E7EB", end_color="E5E7EB", fill_type="solid")    # 淡灰: 源頭斷貨

    total_rows = len(diag_df)
    for row_idx, row_data in enumerate(diag_df.itertuples(index=False), start=4):
        d001_status = getattr(row_data, "_asdict")().get("D001 狀態") if d001_status_col_idx else None
        for col_idx, val in enumerate(row_data, 1):
            if pd.isna(val):
                val = None
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0"
            # D001 狀態欄著色
            if d001_status_col_idx and col_idx == d001_status_col_idx:
                if val == "倉庫有貨":
                    cell.fill = restock_fill
                    cell.font = Font(bold=True, size=10, color="92400E")
                elif val == "源頭斷貨":
                    cell.fill = source_fill
                    cell.font = Font(size=10, color="4B5563")
        if progress_callback and ((row_idx - 4) % 500 == 0 or row_idx - 3 == total_rows):
            progress_callback(0.741 + 0.018 * ((row_idx - 3) / max(total_rows, 1)),
                              f"寫入 SKU 缺貨 Review ({row_idx-3}/{total_rows})...")

    if progress_callback:
        progress_callback(0.759, f"SKU 缺貨 Review 表寫入完成（{len(diag_df)} SKU）")

    _auto_width(ws)



def _write_inventory_distribution_sheet(ws, inv_distribution, header_font, header_fill, thin_border, progress_callback):
    """寫入存貨分佈分析 sheet（3 個表垂直排列）。

    B1: Mix + Non Tourist Turnday 分級彙總
    B2: Nature 存貨對比
    B3: 存貨失衡 SKU 清單
    """
    ws.cell(row=1, column=1, value="Inventory Distribution — 存貨分佈分析").font = Font(bold=True, size=14, color="1E40AF")
    ws.cell(row=2, column=1, value="重點：高銷售 Tourist 店無貨 vs Non Tourist/Mix 大量囤積；Non Tourist/Mix 只需保留 7-14 日 Turnday").font = Font(size=10, italic=True, color="6B7280")

    current_row = 4

    def _write_subtable(title, df, start_row):
        ws.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=12, color="1E40AF")
        start_row += 1
        if df is None or df.empty:
            ws.cell(row=start_row, column=1, value="無資料").font = Font(size=11, italic=True, color="6B7280")
            return start_row + 2
        headers = list(df.columns)
        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=start_row, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, row_data in enumerate(df.itertuples(index=False), start=start_row + 1):
            is_total = any(str(v) == "總計" for v in row_data if isinstance(v, str))
            for col_idx, val in enumerate(row_data, 1):
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if is_total:
                    cell.font = Font(bold=True, size=11)
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
        return start_row + 1 + len(df) + 2  # header + data + spacing

    # B1a: Mix Turnday 分級
    if progress_callback:
        progress_callback(0.77, "寫入 Mix Turnday 分級...")
    current_row = _write_subtable("📊 Mix 店舖 Turnday 分級（標準：保留 7-14 日）", inv_distribution.get("mix_turnday"), current_row)

    # B1b: Non Tourist Turnday 分級
    if progress_callback:
        progress_callback(0.775, "寫入 Non Tourist Turnday 分級...")
    current_row = _write_subtable("📊 Non Tourist 店舖 Turnday 分級", inv_distribution.get("nt_turnday"), current_row)

    # B2: Nature 對比
    if progress_callback:
        progress_callback(0.785, "寫入 Nature 存貨對比...")
    current_row = _write_subtable("🏪 各 Shop Nature 存貨分佈對比", inv_distribution.get("nature_compare"), current_row)

    # B3: 存貨失衡 SKU
    if progress_callback:
        progress_callback(0.79, "寫入存貨失衡 SKU...")
    imb = inv_distribution.get("imbalance_sku")
    current_row = _write_subtable("🔄 存貨失衡 SKU（Tourist 缺貨 + Non Tourist/Mix 過剩 → 建議調貨）", imb, current_row)

    # B3 註解（計算方式說明）
    if imb is not None and hasattr(imb, "attrs") and "annotation" in imb.attrs:
        for line in imb.attrs["annotation"].split("\n"):
            ws.cell(row=current_row, column=1, value=line).font = Font(size=9, italic=True, color="6B7280")
            current_row += 1
        current_row += 1

    # B4: 低存貨 SKU 配置決策
    if progress_callback:
        progress_callback(0.799, "寫入低存貨 SKU 配置決策...")
    low_alloc = inv_distribution.get("low_stock_allocation")
    if low_alloc is not None and not low_alloc.empty:
        # 轉置為 2 欄（指標/數值）方便閱讀
        low_display = low_alloc.T.reset_index()
        low_display.columns = ["指標", "數值"]
        current_row = _write_subtable("⚖️ 低存貨 SKU 配置決策（D001 + Shop Stock < 門檻 → 是否減少 Non Tourist RF 店舖）", low_display, current_row)
    else:
        ws.cell(row=current_row, column=1, value="⚖️ 低存貨 SKU 配置決策：無資料").font = Font(size=12, color="1E40AF")
        current_row += 3

    # B5: 可調貨優先 SKU 清單（Tourist缺貨 + 同SKU在Mix/NT有可釋放件數）
    if progress_callback:
        progress_callback(0.81, "寫入可調貨優先 SKU 清單...")
    tp = inv_distribution.get("transfer_priority", {})
    tp_detail = tp.get("detail") if isinstance(tp, dict) else pd.DataFrame()
    tp_summary = tp.get("summary") if isinstance(tp, dict) else pd.DataFrame()
    if tp_detail is not None and not tp_detail.empty:
        if tp_summary is not None and not tp_summary.empty:
            s = tp_summary.iloc[0]
            summary_text = f"📋 可調貨Tourist行數: {int(s.get('可調貨Tourist行數', 0)):,} 行｜可調貨SKU數: {int(s.get('可調貨SKU數', 0)):,} 個"
        else:
            summary_text = f"📋 可調貨優先 SKU 清單（{len(tp_detail):,} 個 SKU）"
        current_row = _write_subtable(
            f"🔄 可調貨優先 SKU 清單（Tourist缺貨 + 同SKU在Mix/NT有可釋放件數 → 優先調貨）\n{summary_text}",
            tp_detail, current_row,
        )
    else:
        ws.cell(row=current_row, column=1, value="🔄 可調貨優先 SKU 清單：無符合條件 SKU（所有Tourist缺貨SKU在Mix/NT均無可釋放存貨）").font = Font(size=11, color="1E40AF")
        current_row += 3

    _auto_width(ws)



def _build_low_stock_item_class_summary(sku_monthly_analysis: pd.DataFrame) -> pd.DataFrame:
    if sku_monthly_analysis is None or sku_monthly_analysis.empty:
        return pd.DataFrame()
    if "低存貨" not in sku_monthly_analysis.columns:
        return pd.DataFrame()
    df = sku_monthly_analysis[sku_monthly_analysis["低存貨"]].copy()
    if df.empty:
        return pd.DataFrame()
    ic_col = "Item Class"
    if ic_col not in df.columns:
        return pd.DataFrame()
    df[ic_col] = df[ic_col].apply(
        lambda v: "New SKU" if pd.isna(v) or str(v).lower() in ("nan", "none", "nat", "") else str(v)
    )
    counts = df.groupby(ic_col, dropna=False).agg(
        低存貨SKU隻數=("SKU", "nunique")
    ).reset_index()
    counts[ic_col] = counts[ic_col].apply(
        lambda v: "New SKU" if pd.isna(v) or str(v).lower() in ("nan", "none", "nat", "") else str(v)
    )
    all_cats = sorted([str(c) for c in counts[ic_col].unique() if c not in ("New SKU", "總計")])
    all_cats += ["New SKU"] if "New SKU" in counts[ic_col].values else []
    all_cats += ["總計"]
    counts[ic_col] = pd.Categorical(counts[ic_col], categories=all_cats, ordered=True)
    counts = counts.sort_values(ic_col).reset_index(drop=True)
    total_val = int(counts["低存貨SKU隻數"].sum())
    total_row = pd.DataFrame({"Item Class": ["總計"], "低存貨SKU隻數": [total_val]})
    return pd.concat([counts, total_row], ignore_index=True)



def _write_stock_impact_dashboard(ws, raw_df, summary_buyer, inv_distribution, thin_border, desc_font, progress_callback, sku_monthly_analysis=None, stock_threshold=1000):
    """重做為「存貨分佈分析」焦點 Dashboard（v1.6.1）。

    聚焦存貨錯配議題：
    - Tourist 高銷售店缺貨 vs Non Tourist/Mix 大量囤積
    - Non Tourist/Mix 只需保留 7-14 日 Turnday
    """
    impacted = raw_df[raw_df["Is Stock Impacted"]].copy()
    nature = raw_df.get("Shop Nature", pd.Series(dtype=object)).fillna("").astype(str)
    cm_stock = _safe_numeric(raw_df.get("Shop Stock Qty (CM)", _safe_default_series(0, raw_df.index)))
    turnday = pd.to_numeric(raw_df.get("Turnday (Days)", _safe_default_series(0, raw_df.index)), errors="coerce")

    # Tourist 缺貨行數（高銷售店無貨）
    tourist_oos_rows = int(((nature == "Tourist") & (cm_stock <= 0) & raw_df["Is Stock Impacted"]).sum()) if "Is Stock Impacted" in raw_df else 0

    # Tourist 缺貨數量（缺失件數 = max(0, Adj. Safety Stock - Shop Stock Qty (CM))）
    _adj_ss = _safe_numeric(raw_df.get("Adj. Safety Stock", _safe_default_series(0, raw_df.index)))
    _oos_mask = (nature == "Tourist") & (cm_stock <= 0) & (raw_df.get("Is Stock Impacted", pd.Series(False, index=raw_df.index)))
    tourist_oos_qty = int(np.maximum(0, (_adj_ss.values - cm_stock.values))[_oos_mask.values].sum())
    # 曾缺貨 Tourist 行數（含已恢復，不限 CM 當期缺貨）
    tourist_oos_broader = int(((nature == "Tourist") & raw_df["Is Stock Impacted"]).sum()) if "Is Stock Impacted" in raw_df else 0

    # Non Tourist 可釋放存貨件數（>14日標準，與 Mix 一致）
    nt_excess_total = int(_safe_numeric(raw_df.get("Non Tourist Excess Stock", _safe_default_series(0, raw_df.index))).sum())

    # Mix 可釋放存貨件數（>14日標準，來自 Mix Excess Stock）
    mix_excess_total = int(_safe_numeric(raw_df.get("Mix Excess Stock", _safe_default_series(0, raw_df.index))).sum())

    # 存貨失衡 SKU 數（Tourist 缺貨 + 其他 Nature 過剩）
    imbalance_count = len(inv_distribution.get("imbalance_sku", pd.DataFrame())) if inv_distribution else 0

    # 可調貨優先分析（Tourist缺貨 + 同SKU在Mix/NT有可釋放件數）
    tp = inv_distribution.get("transfer_priority", {}) if inv_distribution else {}
    tp_summary = tp.get("summary", pd.DataFrame()) if isinstance(tp, dict) else pd.DataFrame()
    transferrable_tourist_rows = 0
    transferrable_sku_count = 0
    if tp_summary is not None and not tp_summary.empty:
        s = tp_summary.iloc[0]
        transferrable_tourist_rows = int(s.get("可調貨Tourist行數", 0))
        transferrable_sku_count = int(s.get("可調貨SKU數", 0))

    # 各 Nature 平均 Turnday
    tourist_td = turnday[nature == "Tourist"]
    tourist_avg_td = round(float(tourist_td.mean()), 1) if tourist_td.notna().any() and len(tourist_td) else 0
    mix_td = turnday[nature == "Mix"]
    mix_avg_td = round(float(mix_td.mean()), 1) if mix_td.notna().any() and len(mix_td) else 0

    # Non Tourist 平均 Turnday
    nt_td = turnday[nature == "Non Tourist"]
    nt_avg_td = round(float(nt_td.mean()), 1) if nt_td.notna().any() and len(nt_td) else 0

    # 受影響最嚴重 Buyer
    top_buyer_name = "—"
    top_buyer_loss = 0
    if summary_buyer is not None and not summary_buyer.empty and "Buyer" in summary_buyer.columns:
        buyer_data = summary_buyer[~summary_buyer["Buyer"].isin(["總計", ""])]
        if not buyer_data.empty and "Tourist 缺貨行數" in buyer_data.columns:
            top = buyer_data.sort_values("Tourist 缺貨行數", ascending=False).iloc[0]
            top_buyer_name = str(top["Buyer"])
            top_buyer_loss = int(top.get("Tourist 缺貨行數", 0))

    row1_metrics = [
        ("⭐ Tourist 缺貨行數", tourist_oos_rows, "高銷售店無貨（只計 RF）｜單位：行數", "DC2626"),
        ("Non Tourist 可釋存貨", nt_excess_total, "超過14日標準可釋放｜單位：件數", "F59E0B"),
        ("Mix 可釋放存貨", mix_excess_total, "超過14日標準可釋放｜單位：件數", "F97316"),
        ("存貨失衡 SKU 數", imbalance_count, "Tourist缺貨+其他過剩｜單位：SKU數", "EF4444"),
    ]

    row2_metrics = [
        ("Tourist 平均Turnday", tourist_avg_td, "Tourist 平均存貨天數｜單位：日", "8B5CF6"),
        ("Mix 平均Turnday", mix_avg_td, "Mix 平均存貨天數｜單位：日", "3B82F6"),
        ("Non Tourist 平均Turnday", nt_avg_td, "Non Tourist 平均存貨天數｜單位：日", "06B6D4"),
        ("受影響最嚴重 Buyer", top_buyer_name, f"Tourist 缺貨: {top_buyer_loss:,} 行", "1E40AF"),
    ]

    row3_metrics = [
        ("🔄 可調貨Tourist行數", transferrable_tourist_rows, "Tourist缺貨+同SKU在Mix/NT有可釋放件數｜單位：行數", "059669"),
        ("📦 可調貨SKU數", transferrable_sku_count, "涉及的不重複SKU數量｜單位：SKU數", "0D9488"),
        ("📉 Tourist 缺貨數量", tourist_oos_qty, "缺失件數 = Safety − Shop Net Stock｜單位：件數", "DC2626"),
        ("🕐 曾缺貨Tourist行數(含已恢復)", tourist_oos_broader, "任何一期曾缺貨的Tourist RF行數｜含已恢復供貨", "6366F1"),
    ]

    low_stock_total = 0
    low_stock_item_class = pd.DataFrame()
    if sku_monthly_analysis is not None and not sku_monthly_analysis.empty and "低存貨" in sku_monthly_analysis.columns:
        low_stock_total = int(sku_monthly_analysis["低存貨"].sum())
        low_stock_item_class = _build_low_stock_item_class_summary(sku_monthly_analysis)

    row4_metrics = [
        ("🔴 低存貨 SKU 總數", low_stock_total, f"D001+店舖存貨<{stock_threshold}件｜單位：SKU數", "DC2626"),
    ]

    _write_metric_cards(ws, "缺貨銷售影響分析摘要 (Stock Impact Dashboard)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback, row3_metrics=row3_metrics, row4_metrics=row4_metrics)

    # Top 5 受影響 Brand（按 Tourist 缺貨）
    if summary_buyer is not None and not summary_buyer.empty and len(summary_buyer) > 1:
        ws.cell(row=19, column=1, value="Top 5 受缺貨影響最嚴重的 Buyer（按 Tourist 缺貨排序）").font = Font(bold=True, size=12, color="1E40AF")
        top_buyers = summary_buyer[summary_buyer["Buyer"] != "總計"].head(5)
        for col_idx, h in enumerate(top_buyers.columns, 1):
            cell = ws.cell(row=20, column=col_idx, value=h)
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, (_, row_data) in enumerate(top_buyers.iterrows()):
            excel_row = 21 + row_idx
            for col_idx, h in enumerate(top_buyers.columns, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"

    # 低存貨 SKU 按 Item Class 分類
    table_start = 27
    if low_stock_item_class is not None and not low_stock_item_class.empty:
        ws.cell(row=table_start, column=1, value="🔴 低存貨 SKU 按 Item Class 分類").font = Font(bold=True, size=12, color="DC2626")
        header_row = table_start + 1
        for col_idx, h in enumerate(low_stock_item_class.columns, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, (_, row_data) in enumerate(low_stock_item_class.iterrows()):
            excel_row = header_row + 1 + row_idx
            is_total = row_data.get("Item Class") == "總計"
            for col_idx, h in enumerate(low_stock_item_class.columns, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
                if is_total:
                    cell.font = Font(bold=True)
    else:
        ws.cell(row=table_start, column=1, value="🔴 低存貨 SKU 按 Item Class 分類：無低存貨 SKU").font = Font(size=12, color="DC2626")

    _auto_width(ws)



def generate_excel_stock_impact(
    summary_brand: pd.DataFrame,
    summary_buyer: pd.DataFrame,
    summary_category: pd.DataFrame,
    summary_item_class: pd.DataFrame,
    sku_diagnosis: pd.DataFrame,
    inv_distribution: dict,
    raw_df: pd.DataFrame,
    progress_callback=None,
    sku_monthly_analysis: pd.DataFrame | None = None,
    stock_threshold: int = 1000,
    include_raw: bool = True,
) -> BytesIO:
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    desc_font = Font(size=9, color="6B7280", italic=True)

    # Sheet 1: RAW（只輸出分析相關欄位以節省記憶體）
    ws_raw = wb.active
    ws_raw.title = "RAW"
    (_write_raw_sheet(ws_raw, raw_df, progress_callback, progress_start=0.0, progress_end=0.08, columns=_OOS_DETAIL_COLS) if include_raw else _write_raw_placeholder(ws_raw))

    if progress_callback:
        progress_callback(0.1, "建立缺貨影響明細工作表...")

    # Sheet 2: OOS Impact Detail
    ws_detail = wb.create_sheet(title="OOS Impact Detail")
    _write_oos_detail_sheet(ws_detail, raw_df, header_font, header_fill, thin_border, progress_callback)

    # Sheet 3: By Brand
    if progress_callback:
        progress_callback(0.3, "建立 By Brand 分析工作表...")
    ws_brand = wb.create_sheet(title="By Brand")
    if not summary_brand.empty:
        _write_stock_impact_summary_sheet(ws_brand, summary_brand, "By Brand — 缺貨影響分析",
                                          header_font, header_fill, thin_border, progress_callback, 0.30, 0.42)
    else:
        ws_brand.cell(row=1, column=1, value="無 Brand 層級缺貨影響數據").font = Font(size=12)

    # Sheet 4: By Buyer
    if progress_callback:
        progress_callback(0.44, "建立 By Buyer 分析工作表...")
    ws_buyer = wb.create_sheet(title="By Buyer")
    if not summary_buyer.empty:
        _write_stock_impact_summary_sheet(ws_buyer, summary_buyer, "By Buyer — 缺貨影響分析",
                                          header_font, header_fill, thin_border, progress_callback, 0.44, 0.56)
    else:
        ws_buyer.cell(row=1, column=1, value="無 Buyer 層級缺貨影響數據").font = Font(size=12)

    # Sheet 5: By Category
    if progress_callback:
        progress_callback(0.58, "建立 By Category 分析工作表...")
    ws_category = wb.create_sheet(title="By Category")
    if not summary_category.empty:
        _write_stock_impact_summary_sheet(ws_category, summary_category, "By Category (Adj. MCH1) — 缺貨影響分析",
                                          header_font, header_fill, thin_border, progress_callback, 0.58, 0.68)
    else:
        ws_category.cell(row=1, column=1, value="無 Category (Adj. MCH1) 層級缺貨影響數據").font = Font(size=12)

    # Sheet 6: By Item Class
    if progress_callback:
        progress_callback(0.68, "建立 By Item Class 分析工作表...")
    ws_item_class = wb.create_sheet(title="By Item Class")
    if not summary_item_class.empty:
        _write_stock_impact_summary_sheet(ws_item_class, summary_item_class, "By Item Class — 缺貨影響分析",
                                          header_font, header_fill, thin_border, progress_callback, 0.68, 0.70)
    else:
        ws_item_class.cell(row=1, column=1, value="無 Item Class 層級缺貨影響數據").font = Font(size=12)

    # Sheet 6b: SKU Monthly Sales & Storage (按 SKU 月銷量與存儲量)
    if progress_callback:
        progress_callback(0.69, "建立 SKU 月銷量和存儲量工作表...")
    ws_sku_monthly = wb.create_sheet(title="SKU Monthly Sales&Store")
    if sku_monthly_analysis is not None and not sku_monthly_analysis.empty:
        ws_sku_monthly.cell(row=1, column=1, value="SKU 月銷量和存儲量分析（按 Shop Nature）").font = Font(bold=True, size=12, color="1E40AF")
        ws_sku_monthly.cell(row=2, column=1, value="各 SKU 在 Tourist/Mix/Non Tourist 店舖的月均銷售（Shop Monthly Sold Qty P90D 已是月均）與當期庫存（CM Shop Stock）。低存貨 SKU 優先顯示。").font = Font(size=10, italic=True, color="6B7280")
        headers = list(sku_monthly_analysis.columns)
        for col_idx, h in enumerate(headers, 1):
            cell = ws_sku_monthly.cell(row=4, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, row_data in enumerate(sku_monthly_analysis.itertuples(index=False), start=5):
            for col_idx, val in enumerate(row_data, 1):
                if pd.isna(val):
                    val = None
                cell = ws_sku_monthly.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, bool):
                    cell.value = "✓ 低存貨" if val else ""
                elif isinstance(val, (int, float)) and val is not None:
                    cell.number_format = "#,##0"
            if progress_callback and (row_idx % 500 == 0 or row_idx - 4 == len(sku_monthly_analysis)):
                progress_callback(0.69 + 0.01 * ((row_idx - 4) / max(len(sku_monthly_analysis), 1)), f"寫入 SKU 月銷量 ({row_idx-4}/{len(sku_monthly_analysis)})...")
        _auto_width(ws_sku_monthly)
    else:
        ws_sku_monthly.cell(row=1, column=1, value="無 SKU 月銷量/存儲量數據").font = Font(size=12)

    # Sheet 7: SKU Stock Diagnosis (by-SKU 缺貨 Review, 含 D001 + 店舖貨量)
    if progress_callback:
        progress_callback(0.70, "建立 by-SKU 缺貨 Review 工作表...")
    ws_diag = wb.create_sheet(title="SKU Stock Diagnosis")
    _write_sku_stock_diagnosis_sheet(ws_diag, sku_diagnosis, header_font, header_fill, thin_border, progress_callback)

    # Sheet 8: Inventory Distribution (存貨分佈分析)
    if progress_callback:
        progress_callback(0.76, "建立存貨分佈分析工作表...")
    ws_inv = wb.create_sheet(title="Inventory Distribution")
    _write_inventory_distribution_sheet(ws_inv, inv_distribution, header_font, header_fill, thin_border, progress_callback)

    # Sheet 9: Dashboard
    if progress_callback:
        progress_callback(0.80, "建立缺貨影響分析摘要工作表...")
    ws_dashboard = wb.create_sheet(title="Stock Impact Dashboard")
    _write_stock_impact_dashboard(ws_dashboard, raw_df, summary_buyer, inv_distribution,
                                  thin_border, desc_font, progress_callback, sku_monthly_analysis=sku_monthly_analysis, stock_threshold=stock_threshold)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf



def _write_stock_impact_dashboard_simplified(ws, raw_df, thin_border, desc_font, progress_callback, sku_monthly_analysis=None, stock_threshold=1000, sku_soh_summary=None):
    """簡化版 Dashboard（含低存貨分析 + 有賣VS有貨比率區間）。"""
    nature = raw_df.get("Shop Nature", pd.Series(dtype=object)).fillna("").astype(str)
    cm_stock = _safe_numeric(raw_df.get("Shop Stock Qty (CM)", _safe_default_series(0, raw_df.index)))
    turnday = pd.to_numeric(raw_df.get("Turnday (Days)", _safe_default_series(0, raw_df.index)), errors="coerce")

    # Tourist 缺貨行數（高銷售店無貨）
    tourist_oos_rows = int(((nature == "Tourist") & (cm_stock <= 0) & raw_df["Is Stock Impacted"]).sum()) if "Is Stock Impacted" in raw_df else 0

    # Tourist 缺貨數量（缺失件數 = max(0, Adj. Safety Stock - Shop Stock Qty (CM))）
    _adj_ss = _safe_numeric(raw_df.get("Adj. Safety Stock", _safe_default_series(0, raw_df.index)))
    _oos_mask = (nature == "Tourist") & (cm_stock <= 0) & (raw_df.get("Is Stock Impacted", pd.Series(False, index=raw_df.index)))
    tourist_oos_qty = int(np.maximum(0, (_adj_ss.values - cm_stock.values))[_oos_mask.values].sum())

    # 曾缺貨 Tourist 行數（含已恢復，不限 CM 當期缺貨）
    tourist_oos_broader = int(((nature == "Tourist") & raw_df["Is Stock Impacted"]).sum()) if "Is Stock Impacted" in raw_df else 0

    # Non Tourist 可釋放存貨件數（>14日標準）
    nt_excess_total = int(_safe_numeric(raw_df.get("Non Tourist Excess Stock", _safe_default_series(0, raw_df.index))).sum())

    # Mix 可釋放存貨件數（>14日標準）
    mix_excess_total = int(_safe_numeric(raw_df.get("Mix Excess Stock", _safe_default_series(0, raw_df.index))).sum())

    # 各 Nature 平均 Turnday
    tourist_td = turnday[nature == "Tourist"]
    tourist_avg_td = round(float(tourist_td.mean()), 1) if tourist_td.notna().any() and len(tourist_td) else 0
    mix_td = turnday[nature == "Mix"]
    mix_avg_td = round(float(mix_td.mean()), 1) if mix_td.notna().any() and len(mix_td) else 0
    nt_td = turnday[nature == "Non Tourist"]
    nt_avg_td = round(float(nt_td.mean()), 1) if nt_td.notna().any() and len(nt_td) else 0

    row1_metrics = [
        ("⭐ Tourist 缺貨行數", tourist_oos_rows, "高銷售店無貨（只計 RF）｜單位：行數", "DC2626"),
        ("Non Tourist 可釋存貨", nt_excess_total, "超過14日標準可釋放｜單位：件數", "F59E0B"),
        ("Mix 可釋放存貨", mix_excess_total, "超過14日標準可釋放｜單位：件數", "F97316"),
        ("📉 Tourist 缺貨數量", tourist_oos_qty, "缺失件數 = Safety − Shop Net Stock｜單位：件數", "DC2626"),
    ]

    row2_metrics = [
        ("Tourist 平均Turnday", tourist_avg_td, "Tourist 平均存貨天數｜單位：日", "8B5CF6"),
        ("Mix 平均Turnday", mix_avg_td, "Mix 平均存貨天數｜單位：日", "3B82F6"),
        ("Non Tourist 平均Turnday", nt_avg_td, "Non Tourist 平均存貨天數｜單位：日", "06B6D4"),
        ("🕐 曾缺貨Tourist行數(含已恢復)", tourist_oos_broader, "任何一期曾缺貨的Tourist RF行數｜含已恢復供貨", "6366F1"),
    ]

    low_stock_total = 0
    low_stock_item_class = pd.DataFrame()
    if sku_monthly_analysis is not None and not sku_monthly_analysis.empty and "低存貨" in sku_monthly_analysis.columns:
        low_stock_total = int(sku_monthly_analysis["低存貨"].sum())
        low_stock_item_class = _build_low_stock_item_class_summary(sku_monthly_analysis)

    # 有賣VS有貨新指標（從 sku_soh_summary 計算）
    cm_lt50 = 0
    cm_lt30 = 0
    cm_no_stock = 0
    has_band_data = sku_soh_summary is not None and not sku_soh_summary.empty
    if has_band_data:
        cm_oh = pd.to_numeric(sku_soh_summary.get("有貨店舖 (CM)", pd.Series(0)), errors="coerce").fillna(0)
        cm_ratio = pd.to_numeric(sku_soh_summary.get("有賣VS有貨 (CM)", pd.Series(np.nan)), errors="coerce")
        cm_lt50 = int(((cm_oh > 0) & (cm_ratio < 0.5)).sum())
        cm_lt30 = int(((cm_oh > 0) & (cm_ratio < 0.3)).sum())
        cm_no_stock = int((cm_oh == 0).sum())

    row3_metrics = [
        ("🔴 低存貨 SKU 總數", low_stock_total, f"D001+店舖存貨<{stock_threshold}件｜單位：SKU數", "DC2626"),
        ("📉 CM 有賣/有貨<50% SKU", cm_lt50, f"有貨店舖>0 且 有賣÷有貨<50%｜單位：SKU數", "F97316"),
        ("🔻 CM 有賣/有貨<30% SKU", cm_lt30, "有貨店舖>0 且 有賣÷有貨<30%｜單位：SKU數", "DC2626"),
        ("⬜ CM 完全無貨 SKU", cm_no_stock, "有貨店舖 (CM)=0｜單位：SKU數", "6B7280"),
    ]

    _write_metric_cards(ws, "缺貨銷售影響分析摘要 (Stock Impact Dashboard - 簡化)", row1_metrics, row2_metrics, thin_border, desc_font, progress_callback, row3_metrics=row3_metrics)

    # 有賣VS有貨比率區間分佈表
    if has_band_data:
        band_df = _build_sell_through_band_summary(sku_soh_summary)
        band_start = 15
        ws.cell(row=band_start, column=1, value="📊 有賣VS有貨比率分佈（按 SKU 數）").font = Font(bold=True, size=12, color="1E40AF")
        band_header_row = band_start + 1
        for col_idx, h in enumerate(band_df.columns, 1):
            cell = ws.cell(row=band_header_row, column=col_idx, value=h)
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, (_, row_data) in enumerate(band_df.iterrows()):
            excel_row = band_header_row + 1 + row_idx
            is_total = row_data.get("有賣VS有貨比率區間") == "總計"
            for col_idx, h in enumerate(band_df.columns, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
                if is_total:
                    cell.font = Font(bold=True)
        item_class_start = band_header_row + 1 + len(band_df) + 1  # skip 1 blank row after band table
    else:
        item_class_start = 15

    # 低存貨 SKU 按 Item Class 分類（已下移）
    if low_stock_item_class is not None and not low_stock_item_class.empty:
        ws.cell(row=item_class_start, column=1, value="🔴 低存貨 SKU 按 Item Class 分類").font = Font(bold=True, size=12, color="DC2626")
        header_row = item_class_start + 1
        for col_idx, h in enumerate(low_stock_item_class.columns, 1):
            cell = ws.cell(row=header_row, column=col_idx, value=h)
            cell.font = Font(bold=True, size=10, color="FFFFFF")
            cell.fill = PatternFill(start_color="C0392B", end_color="C0392B", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, (_, row_data) in enumerate(low_stock_item_class.iterrows()):
            excel_row = header_row + 1 + row_idx
            is_total = row_data.get("Item Class") == "總計"
            for col_idx, h in enumerate(low_stock_item_class.columns, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
                if is_total:
                    cell.font = Font(bold=True)
    else:
        ws.cell(row=item_class_start, column=1, value="🔴 低存貨 SKU 按 Item Class 分類：無低存貨 SKU").font = Font(size=12, color="DC2626")

    _auto_width(ws)



def _build_sku_soh_tier_summary(raw: pd.DataFrame) -> pd.DataFrame:
    """建立 'New' sheet（Joseph 格式）：每 SKU 一行的 SOH Tier 分佈 + 有賣VS有貨比率。

    回傳 DataFrame 含 52 欄，排序：Total Stock Qty (CM) 降序。
    """
    if raw.empty or "SKU" not in raw.columns:
        return pd.DataFrame()

    # 先 enrich 取得擴充欄位
    df = _enrich_raw_joseph(raw)

    # 基本資訊 lookup
    sku_lookup = _build_sku_info_lookup(df)
    base_cols = ["SKU"]
    for c in ["SKU Desc", "Brand", "Buyer", "Item Class"]:
        if c in sku_lookup.columns:
            base_cols.append(c)
    info = sku_lookup[base_cols].copy()
    info["Status"] = sku_lookup.get("Item Status", "")
    info["Launch Date"] = sku_lookup.get("Launch Date", "")
    info["Is New SKU (3mo)"] = sku_lookup.get("Is New SKU (3mo)", "")

    # SKU 級彙總
    g = df.groupby("SKU", dropna=False)
    # Eff Price / Price Tier / Total Stock Qty (first)
    _first_cols = [c for c in ["Eff Price", "Price Tier"] if c in df.columns]
    for p in ["M-2", "M-1", "CM"]:
        c = f"Total Stock Qty ({p})"
        if c in df.columns:
            _first_cols.append(c)
    sku_first = g[_first_cols].first().reset_index() if _first_cols else pd.DataFrame({"SKU": g.size().index})

    # 有貨/有賣店舖 (sum of 0/1 flags)
    stock_sold = {}
    for p in ["M-2", "M-1", "CM"]:
        stock_sold[f"有貨店舖 ({p})"] = (f"有貨店舖 ({p})", "sum")
        stock_sold[f"有賣店舖 ({p})"] = (f"有賣店舖 ({p})", "sum")
    ss_agg = g.agg(**stock_sold).reset_index()

    # Sold Qty 加總
    sold_cols = {}
    for p in ["M-2", "M-1"]:
        sold_cols[f"Sold Qty ({p})"] = (f"Sold Qty ({p})", "sum")
    sold_cols["Sold Qty (CM - MTD)"] = ("Sold Qty (CM - MTD)", "sum")
    sold_agg = g.agg(**sold_cols).reset_index()

    # SOH Tier 各期分佈 (value_counts → 寬表)
    tier_periods = []
    for p in ["M-2", "M-1", "CM"]:
        tier_col = f"SOH Tier ({p})"
        if tier_col not in df.columns:
            continue
        # count rows per SKU per tier
        vc = df.groupby(["SKU", tier_col], dropna=False).size().unstack(fill_value=0)
        # ensure all tier labels present
        for lbl in _SOH_TIER_LABELS:
            if lbl not in vc.columns:
                vc[lbl] = 0
        vc = vc[_SOH_TIER_LABELS].copy()
        vc.columns = [f"{lbl}" for lbl in _SOH_TIER_LABELS]
        # add total
        vc["總計"] = vc.sum(axis=1).astype(int)
        vc = vc.reset_index()
        vc.columns = ["SKU"] + list(_SOH_TIER_LABELS) + ["總計"]
        tier_periods.append(vc)

    # 合併
    result = info.merge(sku_first, on="SKU", how="left")
    result = result.merge(ss_agg, on="SKU", how="left")
    result = result.merge(sold_agg, on="SKU", how="left")

    # 有賣VS有貨比率 (cap 1, div/0 → NaN)
    for p in ["M-2", "M-1", "CM"]:
        on_hand = result.get(f"有貨店舖 ({p})", pd.Series(0, index=result.index))
        sold = result.get(f"有賣店舖 ({p})", pd.Series(0, index=result.index))
        ratio = np.where(
            on_hand > 0,
            np.minimum(sold / on_hand, 1.0),
            np.nan,
        )
        result[f"有賣VS有貨 ({p})"] = ratio

    # 低存量標記 (< 1000 件)
    cm_stock = pd.to_numeric(result.get("Total Stock Qty (CM)", pd.Series(0, index=result.index)), errors="coerce").fillna(0)
    result["Low Stock (CM<1000)"] = np.where(cm_stock < 1000, "Y", "")

    # 低有賣VS有貨率標記 — 只標有貨（有貨店舖>0）且比率<50%嘅 SKU
    on_hand_cm = pd.to_numeric(result.get("有貨店舖 (CM)", pd.Series(0, index=result.index)), errors="coerce").fillna(0)
    cm_ratio = pd.to_numeric(result.get("有賣VS有貨 (CM)", pd.Series(np.nan, index=result.index)), errors="coerce")
    result["Low 有賣VS有貨 (CM<50%)"] = np.where((on_hand_cm > 0) & (cm_ratio < 0.5), "Y", "")

    # 合併 tier 區塊
    for tp in tier_periods:
        result = result.merge(tp, on="SKU", how="left")

    # 排序
    sort_col = "Total Stock Qty (CM)"
    if sort_col in result.columns:
        result = result.sort_values(sort_col, ascending=False).reset_index(drop=True)

    # 欄位順序 (52+2 columns, matching Joseph + new marker columns)
    out = ["SKU", "SKU Desc", "Brand", "Buyer", "Item Class", "Status",
           "Launch Date", "Is New SKU (3mo)",
           "Low Stock (CM<1000)", "Low 有賣VS有貨 (CM<50%)",
           "Eff Price", "Price Tier"]
    for p in ["M-2", "M-1", "CM"]:
        out.append(f"Total Stock Qty ({p})")
    for p in ["M-2", "M-1", "CM"]:
        out.append(f"有貨店舖 ({p})")
    for p in ["M-2", "M-1", "CM"]:
        out.append(f"有賣店舖 ({p})")
    for p in ["M-2", "M-1", "CM"]:
        out.append(f"有賣VS有貨 ({p})")
    if "Sold Qty (M-2)" in result.columns:
        out.append("Sold Qty (M-2)")
    if "Sold Qty (M-1)" in result.columns:
        out.append("Sold Qty (M-1)")
    if "Sold Qty (CM - MTD)" in result.columns:
        out.append("Sold Qty (CM - MTD)")
    for p, pname in [("M-2", "M-2"), ("M-1", "M-1"), ("CM", "CM")]:
        for lbl in _SOH_TIER_LABELS:
            col = f"{lbl}"
            if col in result.columns:
                out.append(col)
        out.append("總計")

    out = [c for c in out if c in result.columns]
    result = result[out]

    return result



def _build_sell_through_band_summary(sku_soh_summary: pd.DataFrame) -> pd.DataFrame:
    """建立「有賣VS有貨比率區間分佈表」：三期 × 5 區間（含 無貨行 + 總計行）。

    輸入 = _build_sku_soh_tier_summary 的輸出（New sheet df）。
    回傳 6 行 × 4 欄 DataFrame：比率區間／三期各區間 SKU 數。
    """
    if sku_soh_summary.empty:
        return pd.DataFrame()

    required = []
    for p in ["M-2", "M-1", "CM"]:
        required.append(f"有貨店舖 ({p})")
        required.append(f"有賣VS有貨 ({p})")
    if not all(c in sku_soh_summary.columns for c in required):
        return pd.DataFrame()

    bands = [
        ("<30%", lambda oh, r: (oh > 0) & (r < 0.3)),
        ("30-50%", lambda oh, r: (oh > 0) & (r >= 0.3) & (r < 0.5)),
        ("50-70%", lambda oh, r: (oh > 0) & (r >= 0.5) & (r < 0.7)),
        (">=70%", lambda oh, r: (oh > 0) & (r >= 0.7)),
        ("無貨（有貨=0）", lambda oh, r: oh == 0),
    ]
    periods = ["M-2", "M-1", "CM"]

    rows = []
    for label, mask_fn in bands:
        row = {"有賣VS有貨比率區間": label}
        for p in periods:
            oh = pd.to_numeric(sku_soh_summary[f"有貨店舖 ({p})"], errors="coerce").fillna(0)
            r = pd.to_numeric(sku_soh_summary[f"有賣VS有貨 ({p})"], errors="coerce")
            row[f"SKU數 ({p})"] = int(mask_fn(oh, r).sum())
        rows.append(row)

    band_df = pd.DataFrame(rows)
    total_row = {"有賣VS有貨比率區間": "總計"}
    for p in periods:
        total_row[f"SKU數 ({p})"] = band_df[f"SKU數 ({p})"].sum()
    band_df = pd.concat([band_df, pd.DataFrame([total_row])], ignore_index=True)
    return band_df



def _write_new_sheet(ws, df, header_font, header_fill, thin_border, progress_callback):
    """寫入 'New' sheet（Joseph 格式），含 merged headers for SOH tier blocks。

    格式化：
    - Launch Date: 轉為純文字（去除 .0）
    - 有賣VS有貨: 百分比格式 + 條件格式（<30% 紅底白字粗體 / 30-50% 橘底 / 50-70% 淺黃底）
    - 有貨店舖: >70 紅底白字粗體
    """
    if df is None or df.empty:
        ws.cell(row=1, column=1, value="無 SOH Tier 分佈數據").font = Font(size=12)
        return

    col_list = list(df.columns)

    # 找出各特殊欄位的 1-based index
    launch_date_idx = None
    stock_col_idxs = []       # 有貨店舖 (M-2/M-1/CM)
    ratio_col_idxs = []       # 有賣VS有貨 (M-2/M-1/CM)
    total_stock_col_idxs = [] # Total Stock Qty (M-2/M-1/CM)
    marker_col_idxs = []      # Low Stock / Low 有賣VS有貨 標記欄
    for i, c in enumerate(col_list, 1):
        if c == "Launch Date":
            launch_date_idx = i
        elif c.startswith("有貨店舖"):
            stock_col_idxs.append(i)
        elif c.startswith("有賣VS有貨"):
            ratio_col_idxs.append(i)
        elif c.startswith("Total Stock Qty"):
            total_stock_col_idxs.append(i)
        elif c in ("Low Stock (CM<1000)", "Low 有賣VS有貨 (CM<50%)"):
            marker_col_idxs.append(i)

    # SOH tier merged row: find Sold Qty (CM - MTD) position
    sold_idx = None
    for i, c in enumerate(col_list):
        if c == "Sold Qty (CM - MTD)":
            sold_idx = i
            break

    # Row 3: merged labels for SOH tier blocks
    block_width = len(_SOH_TIER_LABELS) + 1  # 9
    start_col = sold_idx + 2 if sold_idx is not None else (len(col_list) - 3 * block_width + 1)
    if start_col < 1:
        start_col = 1
    for bi, pname in enumerate(["M-2", "M-1", "CM"]):
        try:
            c1 = start_col + bi * block_width
            c2 = c1 + block_width - 1
            ws.merge_cells(f"{get_column_letter(c1)}3:{get_column_letter(c2)}3")
            cell = ws.cell(row=3, column=c1, value=f"店舖SOH Tier ({pname})")
            cell.font = Font(bold=True, size=10, color="1E40AF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
        except Exception:
            pass

    # Row 4: headers
    for col_idx, h in enumerate(col_list, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Conditional format fills
    red_fill = PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid")
    red_font = Font(bold=True, color="FFFFFF", size=11)
    orange_fill = PatternFill(start_color="F97316", end_color="F97316", fill_type="solid")
    yellow_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

    total_rows = len(df)
    for row_idx, row_data in enumerate(df.itertuples(index=False), start=5):
        for col_idx, val in enumerate(row_data, 1):
            if pd.isna(val):
                val = None
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

            # Launch Date: 轉純文字（去 .0）
            if col_idx == launch_date_idx and val is not None:
                try:
                    cell.value = str(int(float(val)))
                except (ValueError, TypeError):
                    cell.value = str(val)

            elif col_idx in marker_col_idxs:
                cell.number_format = "@"
            elif isinstance(val, float):
                # Total Stock Qty: 整數格式（無 .00）
                if col_idx in total_stock_col_idxs:
                    cell.number_format = "#,##0"
                # 有賣VS有貨: 百分比格式 + 條件格式
                elif col_idx in ratio_col_idxs:
                    cell.number_format = "0.0%"
                    if val < 0.3:
                        cell.fill = red_fill
                        cell.font = red_font
                    elif val < 0.5:
                        cell.fill = orange_fill
                        cell.font = Font(bold=True, color="FFFFFF", size=11)
                    elif val < 0.7:
                        cell.fill = yellow_fill
                else:
                    cell.number_format = "#,##0.00"
            elif isinstance(val, int):
                cell.number_format = "#,##0"
                # 有貨店舖: >70 紅底白字粗體
                if col_idx in stock_col_idxs and val > 70:
                    cell.fill = red_fill
                    cell.font = red_font

        if progress_callback and ((row_idx - 5) % 500 == 0 or row_idx - 4 == total_rows):
            progress_callback(0.55 + 0.25 * ((row_idx - 4) / max(total_rows, 1)),
                              f"寫入 New sheet ({row_idx-4}/{total_rows})...")

    _auto_width(ws)



def generate_excel_stock_impact_simplified(
    raw_df: pd.DataFrame,
    progress_callback=None,
    sku_monthly_analysis: pd.DataFrame | None = None,
    stock_threshold: int = 1000,
    sku_soh_summary: pd.DataFrame | None = None,
    include_raw: bool = True,
) -> BytesIO:
    """簡化版 Excel 輸出：輸出 RAW（Joseph 擴充欄位）、SKU Monthly Sales&Store、New、Stock Impact Dashboard 四個工作表。"""
    wb = Workbook()

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    desc_font = Font(size=9, color="6B7280", italic=True)

    # 先 enrich raw (計算 Total Stock / 有貨有賣 / Eff Price / SOH Tier 等)
    enriched = _enrich_raw_joseph(raw_df)

    # Sheet 1: RAW (Joseph 擴充欄位順序)
    ws_raw = wb.active
    ws_raw.title = "RAW"
    (_write_raw_sheet(ws_raw, enriched, progress_callback, progress_start=0.0, progress_end=0.2, columns=_JOSEPH_RAW_COLS) if include_raw else _write_raw_placeholder(ws_raw))

    if progress_callback:
        progress_callback(0.25, "建立 SKU Monthly Sales&Store 工作表...")

    # Sheet 2: SKU Monthly Sales&Store
    ws_sku_monthly = wb.create_sheet(title="SKU Monthly Sales&Store")
    if sku_monthly_analysis is not None and not sku_monthly_analysis.empty:
        ws_sku_monthly.cell(row=1, column=1, value="SKU 月銷量和存儲量分析（簡化版，含原版欄位 + 新增 M-2/M-1 不分 Nature 存貨/銷售）").font = Font(bold=True, size=12, color="1E40AF")
        ws_sku_monthly.cell(row=2, column=1, value="保留原版 Tourist/Mix/Non Tourist 所有欄位及新增資訊欄位，再新增 M-2, M-1 的 D001 存貨量、店舖總存貨量及 M-2, M-1, MTD 的店舖總銷售量（不分 Shop Nature）。低存貨 SKU 優先顯示。").font = Font(size=10, italic=True, color="6B7280")
        headers = list(sku_monthly_analysis.columns)
        for col_idx, h in enumerate(headers, 1):
            cell = ws_sku_monthly.cell(row=4, column=col_idx, value=h)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        for row_idx, row_data in enumerate(sku_monthly_analysis.itertuples(index=False), start=5):
            for col_idx, val in enumerate(row_data, 1):
                if pd.isna(val):
                    val = None
                cell = ws_sku_monthly.cell(row=row_idx, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(val, bool):
                    cell.value = "✓ 低存貨" if val else ""
                elif isinstance(val, (int, float)) and val is not None:
                    cell.number_format = "#,##0"
            if progress_callback and (row_idx % 500 == 0 or row_idx - 4 == len(sku_monthly_analysis)):
                progress_callback(0.2 + 0.35 * ((row_idx - 4) / max(len(sku_monthly_analysis), 1)), f"寫入 SKU 月銷量 ({row_idx-4}/{len(sku_monthly_analysis)})...")
        _auto_width(ws_sku_monthly)
    else:
        ws_sku_monthly.cell(row=1, column=1, value="無 SKU 月銷量/存儲量數據").font = Font(size=12)

    if progress_callback:
        progress_callback(0.58, "建立 New sheet（SOH Tier 分佈）...")

    # Sheet 3: New (SOH Tier 分佈，Joseph 格式)
    ws_new = wb.create_sheet(title="New")
    if sku_soh_summary is not None and not sku_soh_summary.empty:
        new_df = sku_soh_summary
    else:
        new_df = _build_sku_soh_tier_summary(raw_df)
    _write_new_sheet(ws_new, new_df, header_font, header_fill, thin_border, progress_callback)

    if progress_callback:
        progress_callback(0.85, "建立 Stock Impact Dashboard 工作表...")

    # Sheet 4: Stock Impact Dashboard
    ws_dashboard = wb.create_sheet(title="Stock Impact Dashboard")
    if "Is Stock Impacted" in raw_df.columns and not raw_df.empty:
        _write_stock_impact_dashboard_simplified(ws_dashboard, raw_df, thin_border, desc_font, progress_callback, sku_monthly_analysis=sku_monthly_analysis, stock_threshold=stock_threshold, sku_soh_summary=new_df)
    else:
        ws_dashboard.cell(row=1, column=1, value="無足夠數據生成 Dashboard").font = Font(size=12)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf




def _add_stock_impact_columns(df: pd.DataFrame) -> pd.DataFrame:
    """計算每列的缺貨狀態、缺貨模式與銷售損失（雙基準並列）。

    缺貨定義採嚴格標準：Shop Stock Qty = 0。
    損失基準並列：(1) Sold Qty (M-2) 需求基準；(2) Shop Monthly Sold Qty (P90D) 月均銷售。
    """
    df = df.copy()
    idx = df.index

    # P0（個別模式）：同欄數值轉換快取，避免同一欄重複 pd.to_numeric；轉換口徑不變
    _num_cache: dict[str, pd.Series] = {}
    def _cached_numeric(col: str | None) -> pd.Series:
        if col is None:
            return _safe_numeric(_safe_default_series(0, idx))
        if col in _num_cache:
            return _num_cache[col]
        try:
            s = _safe_numeric(df.get(col, _safe_default_series(0, idx)))
        except Exception:
            s = _safe_numeric(_safe_default_series(0, idx))
        _num_cache[col] = s
        return s

    # ── 缺貨判定（Shop Stock = 0）──────────────────────────────────────────────
    oos_flags = {}
    for label, stock_col, _sold_col in _STOCK_SOLD_PERIODS:
        stock = _cached_numeric(stock_col)
        is_oos = stock <= 0
        df[f"Is OOS ({label})"] = is_oos
        oos_flags[label] = is_oos

    oos_count = sum(oos_flags.values())
    df["OOS Periods Count"] = oos_count.astype(int)

    # ── 缺貨模式分類 ────────────────────────────────────────────────────────────
    pattern = np.full(len(df), "無缺貨", dtype=object)
    # 恢復供貨：M-2 缺貨（或 M-2、M-1 都缺貨）但 CM 有貨 → 後期補回
    recovered = (~oos_flags["CM"]) & (oos_flags["M-2"] | oos_flags["M-1"])
    # 間歇：僅 1 期缺貨（且非恢復供貨的 M-2/M-1 缺貨 → CM 有貨已涵蓋）
    intermittent = (oos_count == 1) & ~recovered
    # 長期：≥2 期缺貨
    long_term = oos_count >= 2
    pattern[recovered.values] = "恢復供貨"
    pattern[intermittent.values] = "間歇缺貨"
    pattern[long_term.values] = "長期缺貨"
    df["OOS Pattern"] = pattern

    # ── 銷售損失（雙基準並列）─────────────────────────────────────────────────
    # 需求基準 1：Sold Qty (M-2)
    m2_sold = _cached_numeric("Sold Qty (M-2)")
    # 需求基準 2：Shop Monthly Sold Qty (P90D)
    p90d = _cached_numeric("Shop Monthly Sold Qty (P90D)")

    loss_m2_max = np.zeros(len(df))
    loss_p90d_max = np.zeros(len(df))
    loss_m2_sum = np.zeros(len(df))
    loss_p90d_sum = np.zeros(len(df))

    for label, _stock_col, sold_col in _STOCK_SOLD_PERIODS:
        sold = _cached_numeric(sold_col)
        is_oos = oos_flags[label].values
        # 缺貨當期才算損失；損失 = max(0, 需求基準 − 當期實際銷售)
        loss_m2 = np.where(is_oos & (m2_sold.values > 0), np.maximum(0, m2_sold.values - sold.values), 0)
        loss_p90d = np.where(is_oos & (p90d.values > 0), np.maximum(0, p90d.values - sold.values), 0)
        loss_m2_max = np.maximum(loss_m2_max, loss_m2)
        loss_p90d_max = np.maximum(loss_p90d_max, loss_p90d)
        loss_m2_sum = loss_m2_sum + loss_m2
        loss_p90d_sum = loss_p90d_sum + loss_p90d

    df["Max Sales Loss (M-2)"] = loss_m2_max.astype(int)
    df["Max Sales Loss (P90D)"] = loss_p90d_max.astype(int)
    df["Total Sales Loss (M-2)"] = loss_m2_sum.astype(int)
    df["Total Sales Loss (P90D)"] = loss_p90d_sum.astype(int)

    # ── 影響判定：任一期缺貨且有需求證據，且只計算 RF 店舖 ──────────────────────
    # RP Type 正規化（沿用既有模式：strip + upper），只認 "RF" 為有效店舖
    rp_type = (
        df.get("RP Type", _safe_default_series("", idx))
        .fillna("").astype(str).str.strip().str.upper()
    )
    is_rf = rp_type == "RF"
    has_demand = (m2_sold.values > 0) | (p90d.values > 0)
    any_oos = oos_count.values > 0
    df["Is Stock Impacted"] = any_oos & has_demand & is_rf.values

    # ── Turnday 分級與 Mix 可釋放存貨（存貨分佈分析）──────────────────────────────
    # Turnday (P90D Qty) 原值正規化：inf/NaN 代表 P90D=0（無周轉意義）
    turnday_raw = df.get("Turnday (P90D Qty)", _safe_default_series(0, idx))
    turnday = pd.to_numeric(turnday_raw, errors="coerce")
    # 將 inf（P90D=0 但有存貨）與 NaN 設為 -1 旗標，方便後續分類
    turnday_clean = turnday.replace([np.inf, -np.inf], np.nan).fillna(-1)
    df["Turnday (Days)"] = turnday_clean.where(turnday_clean >= 0, np.nan)

    # Mix 店舖 Turnday 分級（Mix 標準：保留 7-14 日）
    # 正規化 Shop Nature（strip）並寫回，確保後續精確比對一致
    shop_nature_raw = df.get("Shop Nature", _safe_default_series("", idx))
    df["Shop Nature"] = shop_nature_raw.fillna("").astype(str).str.strip()
    shop_nature = df["Shop Nature"]
    is_mix = shop_nature == "Mix"
    is_nt = shop_nature == "Non Tourist"
    cm_stock_val = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, idx)))
    has_cm_stock = cm_stock_val.values > 0

    mix_band = np.full(len(df), "", dtype=object)
    mix_mask = is_mix.values
    # Mix 子集的 Turnday 值
    td = turnday_clean.values
    # 無存貨
    no_stock = ~has_cm_stock
    # 無銷售但有存貨（Turnday = -1 且有存貨）= 純囤積
    no_sale_has_stock = (td < 0) & has_cm_stock
    # 有銷售：依 Turnday 分級
    band_under7 = (td >= 0) & (td < 7)
    band_7_14 = (td >= 7) & (td < 14)
    band_14_30 = (td >= 14) & (td < 30)
    band_over30 = (td >= 30)

    mix_band[mix_mask & no_stock] = "無存貨"
    mix_band[mix_mask & no_sale_has_stock] = "無銷售有存貨"
    mix_band[mix_mask & band_under7] = "<7日(過低)"
    mix_band[mix_mask & band_7_14] = "7-14日(理想)"
    mix_band[mix_mask & band_14_30] = "14-30日(偏高)"
    mix_band[mix_mask & band_over30] = ">30日(過剩)"
    df["Mix Turnday Band"] = mix_band

    # Mix 可釋放存貨：超過 14 日標準的部分；無銷售但有存貨者全數可釋放
    # 理想存貨 = P90D / 90 × 14（P90D為90日總量）；可釋放 = max(0, 實際存貨 − 理想存貨)
    ideal_14d = p90d.values / 90.0 * 14
    excess = np.where(
        mix_mask,
        np.where(no_sale_has_stock, cm_stock_val.values, np.maximum(0, cm_stock_val.values - ideal_14d)),
        0,
    )
    df["Mix Excess Stock"] = excess.astype(int)

    # Non Tourist Turnday 分級（同樣標準，>14日釋放）
    nt_band = np.full(len(df), "", dtype=object)
    nt_mask = is_nt.values
    nt_band[nt_mask & no_stock] = "無存貨"
    nt_band[nt_mask & no_sale_has_stock] = "無銷售有存貨"
    nt_band[nt_mask & band_under7] = "<7日(過低)"
    nt_band[nt_mask & band_7_14] = "7-14日(理想)"
    nt_band[nt_mask & band_14_30] = "14-30日(偏高)"
    nt_band[nt_mask & band_over30] = ">30日(過剩)"
    df["Non Tourist Turnday Band"] = nt_band

    # Non Tourist 可釋放存貨（同樣標準：>14日即可釋放，參考 Mix）
    nt_excess = np.where(
        nt_mask,
        np.where(no_sale_has_stock, cm_stock_val.values, np.maximum(0, cm_stock_val.values - ideal_14d)),
        0,
    )
    df["Non Tourist Excess Stock"] = nt_excess.astype(int)

    # 為「無銷售有存貨」行附上 RP Type（供後續拆分 ND/RF）
    rp_for_band = rp_type.values
    mix_band_with_rp = mix_band.copy()
    mix_band_with_rp[mix_mask & no_sale_has_stock] = np.where(
        rp_for_band[mix_mask & no_sale_has_stock] == "ND",
        "無銷售有存貨(ND)",
        "無銷售有存貨(RF)",
    )
    df["Mix Turnday Band"] = mix_band_with_rp

    nt_band_with_rp = nt_band.copy()
    nt_band_with_rp[nt_mask & no_sale_has_stock] = np.where(
        rp_for_band[nt_mask & no_sale_has_stock] == "ND",
        "無銷售有存貨(ND)",
        "無銷售有存貨(RF)",
    )
    df["Non Tourist Turnday Band"] = nt_band_with_rp

    return df



def _build_sku_info_lookup(raw: pd.DataFrame) -> pd.DataFrame:
    """建立 SKU 級基本資訊查找表（單一次 groupby，供多個 summary 函式共用）。

    回傳 DataFrame 每列一個 SKU，含 Brand / Buyer / SKU Desc / Launch Date /
    Is New SKU (3mo) / Vendor MC / Adj. MCH1 / Item Status / Item Class 等 first-value 欄位。
    """
    if "SKU" not in raw.columns:
        return pd.DataFrame()
    info_cols = [c for c in [
        "Brand", "Buyer", "SKU Desc", "Launch Date", "Is New SKU (3mo)",
        "Vendor MC", "Adj. MCH1", "Item Status", "Item Class",
    ] if c in raw.columns]
    if not info_cols:
        return raw[["SKU"]].drop_duplicates().reset_index(drop=True)
    agg_dict = {col: (col, "first") for col in info_cols}
    return raw.groupby("SKU", dropna=False).agg(**agg_dict).reset_index()



def _build_sku_monthly_analysis(raw: pd.DataFrame, stock_threshold: int = 1000, sku_lookup: pd.DataFrame | None = None) -> pd.DataFrame:
    """按 SKU 計算遊客/混合/非遊客的每月銷量及存儲量分析。

    每列為一個 SKU，彙總各 Nature 的月均銷售、當期庫存、店數等。
    低存貨 SKU（D001 + Shop < threshold）排在前面。

    排序：低存貨優先 → Tourist RF 月均銷售降序。

    效能優化：
    - 使用共用 sku_lookup 取代 3 次獨立 groupby.first() merge
    - ND 行數改用預計算 int 欄位取代 groupby 內的 lambda
    - 補充資訊欄位改用單次 merge 取代逐欄 merge

    Args:
        sku_lookup: 預先建好的 SKU 基本資訊查找表，為 None 時即時計算。
    """
    if raw.empty or "SKU" not in raw.columns or "Shop Nature" not in raw.columns:
        return pd.DataFrame()

    _NATURES = ["Tourist", "Mix", "Non Tourist"]

    nature = raw["Shop Nature"].fillna("").astype(str)
    cm_stock = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))
    p90d = _safe_numeric(raw.get("Shop Monthly Sold Qty (P90D)", _safe_default_series(0, raw.index)))
    cm_sold = _safe_numeric(raw.get("Sold Qty (CM - MTD)", _safe_default_series(0, raw.index)))
    turnday = pd.to_numeric(raw.get("Turnday (Days)", _safe_default_series(0, raw.index)), errors="coerce")
    rp_type = (
        raw.get("RP Type", _safe_default_series("", raw.index))
        .fillna("").astype(str).str.strip().str.upper()
    )
    d001_col = "D001 Stock Qty (CM)"
    d001_cm = _safe_numeric(raw[d001_col]) if d001_col in raw.columns else _safe_default_series(0, raw.index)

    work = raw[["SKU", "Shop Nature"]].copy()
    work["_nature"] = nature.values
    work["_cm_stock"] = cm_stock.values
    work["_p90d"] = p90d.values
    work["_cm_sold"] = cm_sold.values
    work["_turnday"] = turnday.values
    work["_d001"] = d001_cm.values
    work["_rp"] = rp_type.values
    work["_is_rf"] = (work["_rp"] == "RF").astype(int)
    work["_is_nd"] = (work["_rp"] == "ND").astype(int)

    # SKU 級基本資訊 (ND 行數改用預計算 int 取代 lambda)
    sku_base = (
        work.groupby("SKU", dropna=False)
        .agg(
            D001_庫存=("_d001", "first"),
            總店舖庫存=("_cm_stock", "sum"),
            總行數=("SKU", "count"),
            RF行數=("_is_rf", "sum"),
            ND行數=("_is_nd", "sum"),
        )
        .reset_index()
    )
    # 使用共用 lookup 表（單次 merge），取代原本 3 次獨立 groupby.merge
    if sku_lookup is None:
        sku_lookup = _build_sku_info_lookup(raw)
    info_cols = ["SKU"] + [c for c in ["Brand", "Buyer", "SKU Desc"] if c in sku_lookup.columns]
    sku_base = sku_base.merge(sku_lookup[info_cols], on="SKU", how="left")
    sku_base["RF行數"] = sku_base["RF行數"].fillna(0).astype(int)
    sku_base["ND行數"] = sku_base["ND行數"].fillna(0).astype(int)
    sku_base["總可用存貨"] = sku_base["D001_庫存"] + sku_base["總店舖庫存"]
    sku_base["低存貨"] = sku_base["總可用存貨"] < stock_threshold

    # SKU × Nature 級聚合
    sku_nat = (
        work.groupby(["SKU", "_nature"], dropna=False, observed=True)
        .agg(
            月均銷售=("_p90d", "sum"),
            當期CM銷售=("_cm_sold", "sum"),
            當期庫存=("_cm_stock", "sum"),
            店數=("SKU", "count"),
            平均Turnday=("_turnday", "mean"),
        )
        .reset_index()
    )
    sku_nat["月均銷售"] = sku_nat["月均銷售"].astype(float)

    # RF 月均銷售（只計 RF 店）
    work_rf = work[work["_rp"] == "RF"]
    if not work_rf.empty:
        sku_nat_rf = (
            work_rf.groupby(["SKU", "_nature"], dropna=False, observed=True)
            .agg(RF月均銷售=("_p90d", "sum"))
            .reset_index()
        )
        sku_nat_rf["RF月均銷售"] = sku_nat_rf["RF月均銷售"].astype(float)
        sku_nat = sku_nat.merge(sku_nat_rf, on=["SKU", "_nature"], how="left")
    else:
        sku_nat["RF月均銷售"] = 0.0

    # 展開各 Nature 為寬表
    result = sku_base.copy()
    for nat in _NATURES:
        nat_data = sku_nat[sku_nat["_nature"] == nat].copy()
        if nat_data.empty:
            result[f"{nat}_月均銷售"] = 0.0
            result[f"{nat}_當期CM銷售"] = 0.0
            result[f"{nat}_當期庫存"] = 0
            result[f"{nat}_店數"] = 0
            result[f"{nat}_平均Turnday"] = 0.0
            result[f"{nat}_RF月均銷售"] = 0.0
            continue
        nat_data = nat_data.drop(columns=["_nature"])
        nat_data.columns = [f"{nat}_{c}" if c != "SKU" else c for c in nat_data.columns]
        result = result.merge(nat_data, on="SKU", how="left")

    # 填補 NaN
    for nat in _NATURES:
        result[f"{nat}_平均Turnday"] = result.get(f"{nat}_平均Turnday", 0).fillna(0).round(1)
        for c in [f"{nat}_月均銷售", f"{nat}_當期CM銷售", f"{nat}_RF月均銷售"]:
            result[c] = result.get(c, 0).fillna(0).round(1)
        for c in [f"{nat}_當期庫存", f"{nat}_店數"]:
            result[c] = result.get(c, 0).fillna(0).astype(int)

    # 排序：低存貨優先 → Tourist RF 月均銷售降序
    result = result.sort_values(
        ["低存貨", "Tourist_RF月均銷售"],
        ascending=[False, False],
    ).reset_index(drop=True)

    # 轉整數
    for c in ["D001_庫存", "總店舖庫存", "總可用存貨", "總行數", "RF行數", "ND行數"]:
        result[c] = result[c].fillna(0).astype(int)
    for nat in _NATURES:
        result[f"{nat}_當期庫存"] = result[f"{nat}_當期庫存"].fillna(0).astype(int)
        result[f"{nat}_店數"] = result[f"{nat}_店數"].fillna(0).astype(int)

    # ── 補充資訊欄位（從 sku_lookup 一次性補齊，避免重複 groupby）──
    extra_from_lookup = [c for c in ["Vendor MC", "Adj. MCH1", "Item Status", "Launch Date", "Item Class", "Is New SKU (3mo)"]
                          if c not in sku_base.columns and c in sku_lookup.columns]
    if extra_from_lookup:
        result = result.merge(sku_lookup[["SKU"] + extra_from_lookup], on="SKU", how="left")

    # ── 不分 Nature 的週期存貨/銷售欄位（M-2, M-1 D001/ShopStock；Sold M-2, M-1, MTD） ──
    d001_m2 = _safe_numeric(raw["D001 Stock Qty (M-2)"]) if "D001 Stock Qty (M-2)" in raw.columns else _safe_default_series(0, raw.index)
    d001_m1 = _safe_numeric(raw["D001 Stock Qty (M-1)"]) if "D001 Stock Qty (M-1)" in raw.columns else _safe_default_series(0, raw.index)
    shop_m2 = _safe_numeric(raw.get("Shop Stock Qty (M-2)", _safe_default_series(0, raw.index)))
    shop_m1 = _safe_numeric(raw.get("Shop Stock Qty (M-1)", _safe_default_series(0, raw.index)))
    sold_m2 = _safe_numeric(raw.get("Sold Qty (M-2)", _safe_default_series(0, raw.index)))
    sold_m1 = _safe_numeric(raw.get("Sold Qty (M-1)", _safe_default_series(0, raw.index)))
    sold_mtd = _safe_numeric(raw.get("Sold Qty (CM - MTD)", _safe_default_series(0, raw.index)))

    per_data = pd.DataFrame({
        "SKU": raw["SKU"].fillna("").astype(str).values,
        "_d001_m2": d001_m2.values, "_d001_m1": d001_m1.values,
        "_shop_m2": shop_m2.values, "_shop_m1": shop_m1.values,
        "_sold_m2": sold_m2.values, "_sold_m1": sold_m1.values,
        "_sold_mtd": sold_mtd.values,
    })
    per_agg = per_data.groupby("SKU", dropna=False).agg(
        D001_M2=("_d001_m2", "first"), D001_M1=("_d001_m1", "first"),
        ShopStock_M2=("_shop_m2", "sum"), ShopStock_M1=("_shop_m1", "sum"),
        Sold_M2=("_sold_m2", "sum"), Sold_M1=("_sold_m1", "sum"),
        Sold_MTD=("_sold_mtd", "sum"),
    ).reset_index()
    per_agg = per_agg.rename(columns={
        "D001_M2": "D001_M-2", "D001_M1": "D001_M-1",
        "ShopStock_M2": "ShopStock_M-2", "ShopStock_M1": "ShopStock_M-1",
        "Sold_M2": "Sold_M-2", "Sold_M1": "Sold_M-1",
    })
    result = result.merge(per_agg, on="SKU", how="left")
    for c in ["D001_M-2", "D001_M-1", "ShopStock_M-2", "ShopStock_M-1",
              "Sold_M-2", "Sold_M-1", "Sold_MTD"]:
        if c in result.columns:
            result[c] = result[c].fillna(0).astype(int)

    # ── 欄位順序 ──
    _extra_info_cols = [c for c in ["Vendor MC", "Adj. MCH1", "Item Status", "Launch Date", "Item Class", "Is New SKU (3mo)"] if c in result.columns]
    _sold_cols = [c for c in ["Sold_M-2", "Sold_M-1", "Sold_MTD"] if c in result.columns]
    _d001_cols = [c for c in ["D001_M-2", "D001_M-1", "ShopStock_M-2", "ShopStock_M-1"] if c in result.columns]

    out_cols = ["SKU"]
    for col in ["SKU Desc", "Brand", "Buyer"]:
        if col in result.columns:
            out_cols.append(col)
    out_cols += _extra_info_cols
    out_cols += ["低存貨", "總可用存貨", "D001_庫存", "總店舖庫存"]
    out_cols += _sold_cols
    out_cols += ["總行數", "RF行數", "ND行數"]
    for nat in _NATURES:
        out_cols += [
            f"{nat}_月均銷售", f"{nat}_當期CM銷售", f"{nat}_當期庫存",
            f"{nat}_店數", f"{nat}_平均Turnday", f"{nat}_RF月均銷售",
        ]
    out_cols += _d001_cols
    out_cols = [c for c in out_cols if c in result.columns]

    return result[out_cols]



def _build_sku_monthly_analysis_simplified(raw: pd.DataFrame, stock_threshold: int = 1000) -> pd.DataFrame:
    """簡化版 SKU 月銷量和存儲量分析（同原版欄位，含新增不分 Nature 的週期存貨/銷售欄位）。

    直接使用 `_build_sku_monthly_analysis` 的完整輸出（已含補充資訊與週期欄位），
    不需要額外計算。
    """
    if raw.empty or "SKU" not in raw.columns:
        return pd.DataFrame()
    return _build_sku_monthly_analysis(raw, stock_threshold=stock_threshold)



# 三期 Stock→Sold 對應（同期）。每個 tuple: (期別標籤, Stock 欄位, Sold 欄位)
_STOCK_SOLD_PERIODS = [
    ("M-2", "Shop Stock Qty (M-2)", "Sold Qty (M-2)"),
    ("M-1", "Shop Stock Qty (M-1)", "Sold Qty (M-1)"),
    ("CM", "Shop Stock Qty (CM)", "Sold Qty (CM - MTD)"),
]



def _write_sku_allocation_sheet(ws, alloc_df, header_font, header_fill, thin_border, progress_callback):
    ws.cell(row=1, column=1, value="SKU Allocation — 跨 Nature 調貨建議（Tourist 缺貨 vs Non-Tourist/Mix 可支援存貨）").font = Font(bold=True, size=12, color="1E40AF")

    if alloc_df is None or alloc_df.empty:
        ws.cell(row=3, column=1, value="無可調貨 SKU（無 Tourist 缺貨且同時其他 Nature 有存貨的情況）").font = Font(size=11, italic=True, color="6B7280")
        return

    headers = list(alloc_df.columns)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    for row_idx, row_data in enumerate(alloc_df.itertuples(index=False), start=4):
        for col_idx, val in enumerate(row_data, 1):
            if pd.isna(val):
                val = None
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(val, (int, float)):
                cell.number_format = "#,##0"

    if progress_callback:
        progress_callback(0.72, f"SKU 調貨表寫入完成（{len(alloc_df)} SKU）")

    _auto_width(ws)



# Shop Nature 排序權重（越小越優先）
_NATURE_PRIORITY = {"Tourist": 0, "Mix": 1, "Non Tourist": 2}



_NATURE_ORDER = ["Tourist", "Mix", "Non Tourist"]
