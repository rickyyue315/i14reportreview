# -*- coding: utf-8 -*-
"""Safety Stock x NetStock Review模式：由 transformer.py 拆分而來。
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
    _display_safety_header,
    _extract_shop_name,
    _normalize_ref_date,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _write_metric_cards,
    _write_raw_placeholder,
    _write_raw_sheet,
)


from modes.framework import (
    _HYBRID_DIAG_COLUMNS,
    _HYBRID_DIAG_NATURES,
    _HYBRID_DIAG_NORMAL_DONOR,
    _HYBRID_DIAG_NORMAL_NEED,
    _HYBRID_DIAG_ORDER,
    _HYBRID_DIAG_OVER,
    _HYBRID_DIAG_UNDER,
    _HYBRID_DIMENSION_LABELS,
    _HYBRID_EVENT_DONATED,
    _HYBRID_EVENT_DONOR,
    _HYBRID_EVENT_NEED_IN,
    _HYBRID_EVENT_ORDER,
    _HYBRID_FULL_SAFETY_CHG_COLUMN,
    _HYBRID_FULL_SAFETY_DETAIL_COLUMNS,
    _HYBRID_FULL_SAFETY_LINE_COLUMN,
    _HYBRID_FULL_SAFETY_NORMAL,
    _HYBRID_FULL_SAFETY_REVIEW_COLUMN,
    _HYBRID_FULL_SAFETY_SHEET,
    _HYBRID_FULL_SAFETY_SUGGEST_COLUMN,
    _HYBRID_FULL_SAFETY_TURNDAY_COLUMN,
    _HYBRID_RAW_DISPLAY_GROUPS,
    _HYBRID_RETURN_D001_COLUMN,
    _HYBRID_RETURN_D001_COLUMNS,
    _HYBRID_RETURN_D001_GROSS_COLUMN,
    _HYBRID_RETURN_D001_GROSS_INTERNAL,
    _HYBRID_RETURN_D001_SHEET,
    _HYBRID_SAFETY_DAYS_COLUMN,
    _HYBRID_SAFETY_DAYS_DEFAULT,
    _HYBRID_SAFETY_LINE_COLUMN,
    _HYBRID_SAFETY_MIN_KEEP,
    _HYBRID_SALES_BANDS,
    _HYBRID_SALES_BAND_COLUMN,
    _HYBRID_SALES_CLASS_COLUMN,
    _HYBRID_SALES_CLASS_ORDER,
    _HYBRID_SALES_OOS_DAYS_DEFAULT,
    _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT,
    _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT,
    _HYBRID_SUGGEST_TURNDAY_COLUMN,
    _canonicalize_hybrid_columns,
    _hybrid_dimension,
    _hybrid_dimension_groups,
    _hybrid_group_key_series,
    _hybrid_numeric_column,
    _hybrid_om_groups_from_values,
    _hybrid_p90d_zero_sales_mask,
    _hybrid_safety_days,
    _hybrid_safety_days_label,
    _hybrid_sales_bands,
    _hybrid_sales_classification,
    _hybrid_sales_params,
    _hybrid_shortage_hint_text,
    _normalize_hybrid_nature,
    _normalize_hybrid_om,
    _normalize_hybrid_rp_type,
    _transfer_input_consistency_warnings,
    validate_hybrid_input,
)


from modes.issues import (
    _build_hybrid_shop_issues,
    _shop_issue_base_frame,
    _write_shop_issue_sheet,
)


from modes.stock_checking import (
    _build_stock_checking_summary,
    _write_stock_checking_summary_sheet,
)
__all__ = [
    "_HYBRID_BUYER_COMBINED_TITLE",
    "_HYBRID_SAFETY_CHANGE_COLUMNS",
    "_HYBRID_SAFETY_CHANGE_DESC",
    "_HYBRID_SAFETY_CHANGE_ORDER",
    "_HYBRID_SAFETY_CHANGE_TOTAL_DESC",
    "_HYBRID_SAFETY_SKU_COLUMNS",
    "_HYBRID_SALES_SUMMARY_COLUMNS",
    "_HYBRID_TRANSFER_DETAIL_COLUMNS",
    "_build_hybrid_diag_summary",
    "_build_hybrid_diagnosis",
    "_build_hybrid_full_safety_change_summary",
    "_build_hybrid_full_safety_detail",
    "_build_hybrid_full_safety_sku_detail",
    "_build_hybrid_return_d001",
    "_build_hybrid_safety_change_summary",
    "_build_hybrid_safety_sku_detail",
    "_build_hybrid_sales_summary",
    "_build_hybrid_transfer_detail",
    "_first_non_blank_value",
    "_hybrid_allocate_transfers",
    "_hybrid_buyer_combined_detail",
    "_hybrid_buyer_combined_stats",
    "_hybrid_buyer_need_stats",
    "_hybrid_buyer_safety_stats",
    "_hybrid_dashboard_metrics",
    "_hybrid_diag_summary_columns",
    "_hybrid_raw_display_columns",
    "_hybrid_safety_days_from_raw",
    "_hybrid_safety_event_frame",
    "_hybrid_stats_from_raw",
    "_write_hybrid_buyer_combined_table",
    "_write_hybrid_diag_summary_sheet",
    "_write_hybrid_diagnosis_sheet",
    "_write_hybrid_full_safety_sheet",
    "_write_hybrid_guide_sheet",
    "_write_hybrid_summary_block",
    "_write_hybrid_transfer_detail_sheet",
    "generate_excel_hybrid",
    "transform_hybrid",
]


# ── Safety 調整統計（欄名屬新契約，未經要求不得改動）──
_HYBRID_SAFETY_CHANGE_COLUMNS = ["分類", "列數", "SKU 數", "Safety 建議調整量合計", "說明"]



_HYBRID_SAFETY_CHANGE_ORDER = ["增加", "減少", "不變"]



_HYBRID_SAFETY_CHANGE_DESC = {
    "增加": "建議 Safety 設定值高於現有設定（Safety 設定不足）",
    "減少": "建議 Safety 設定值低於現有設定（Safety 設定過高）",
    "不變": "設定正常或四捨五入後無需調整（缺貨／過剩非 Safety 設定因素）",
}



_HYBRID_SAFETY_CHANGE_TOTAL_DESC = "SKU 數為全表不重複 SKU（同一 SKU 可跨分類）；合計為淨調整量"



_HYBRID_SAFETY_SKU_COLUMNS = [
    "Brand", "SKU", "SKU Desc", "Buyer",
    "Review 列數", "需增加店舖數", "需減少店舖數", "設定正常店舖數",
    "現有 Safety 設定合計（調整列）", "建議 Safety 設定合計（調整列）", "淨調整量",
]









def _build_hybrid_diagnosis(raw: pd.DataFrame, variant: str = "free") -> pd.DataFrame:
    """調貨 Review 主表（店舖×SKU 逐列）：只收錄事件類型非空的列，
    按 Brand / SKU / 事件類型 / Site Code 排序；「可轉出貨量」為三 Nature 可轉出欄加總
    （Tourist 欄僅雙向時存在）。
    variant="om" 時改用同OM調貨欄位並重算事件類型（轉出量=0 的來源列改列「可轉出未使用」）；
    Safety Review 與建議調整欄兩變體相同。"""
    if raw is None or raw.empty or "事件類型" not in raw.columns:
        return pd.DataFrame(columns=_HYBRID_DIAG_COLUMNS)
    df = raw[raw["事件類型"].fillna("").astype(str) != ""].copy()
    if df.empty:
        return pd.DataFrame(columns=_HYBRID_DIAG_COLUMNS)
    if variant == "om":
        renames = {
            "實際調入量 (同OM)": "實際調入量", "調入缺口 (同OM)": "調入缺口",
            "轉出量 (同OM)": "轉出量", "調貨來源 (同OM)": "調貨來源 (Site Code)",
            "調貨後庫存 (同OM)": "調貨後庫存", "調貨後 TurnDay (同OM)": "調貨後 TurnDay",
        }
        # 先移除同名的不限OM欄，避免 rename 後重複
        df = df.drop(columns=[v for v in renames.values() if v in df.columns])
        df = df.rename(columns={k: v for k, v in renames.items() if k in df.columns})
        donated = _safe_numeric(df.get("轉出量", _safe_default_series(0, df.index))) > 0
        df["事件類型"] = np.where(
            df["事件類型"] == _HYBRID_EVENT_NEED_IN, _HYBRID_EVENT_NEED_IN,
            np.where(donated, _HYBRID_EVENT_DONATED, _HYBRID_EVENT_DONOR),
        )
    df["可轉出貨量"] = (
        _safe_numeric(df.get("可轉出貨量 (Mix)", _safe_default_series(0, df.index)))
        + _safe_numeric(df.get("可轉出貨量 (Non Tourist)", _safe_default_series(0, df.index)))
        + _safe_numeric(df.get("可轉出貨量 (Tourist)", _safe_default_series(0, df.index)))
    ).astype(int)
    df["_evt_key"] = df["事件類型"].map(_HYBRID_EVENT_ORDER).fillna(9)
    df = df.sort_values(["Brand", "SKU", "_evt_key", "Site Code"], kind="stable")
    ordered = [c for c in _HYBRID_DIAG_COLUMNS if c in df.columns]
    return df[ordered].reset_index(drop=True)



def _build_hybrid_return_d001(raw: pd.DataFrame) -> pd.DataFrame:
    """退D001 建議明細（店舖×SKU 逐列）：只收錄「建議退D001 量」> 0 的列。

    「可退D001 貨量」＝扣除已轉出前的可退額度（中銷／低銷／零銷列＝庫存 − Safety；
    高銷列＝庫存 − 2×Safety，須庫存 ≥ 3×Safety），「建議退D001 量」＝可退額度 − 轉出量。
    RAW 缺本欄（開關未開）時輸出空表；按 Site Code / Brand / SKU 排序。
    """
    cols = _HYBRID_RETURN_D001_COLUMNS
    if (
        raw is None or raw.empty
        or _HYBRID_RETURN_D001_COLUMN not in raw.columns
        or _HYBRID_RETURN_D001_GROSS_INTERNAL not in raw.columns
    ):
        return pd.DataFrame(columns=cols)
    qty = _safe_numeric(raw[_HYBRID_RETURN_D001_COLUMN])
    sel_mask = qty.gt(0).fillna(False)
    if not bool(sel_mask.any()):
        return pd.DataFrame(columns=cols)
    work = _shop_issue_base_frame(raw, sel_mask)
    if "Brand" not in work.columns:
        work["Brand"] = ""
    if "SKU" not in work.columns:
        work["SKU"] = ""
    if "Buyer" not in work.columns:
        work["Buyer"] = ""
    if "SKU Desc" not in work.columns:
        work["SKU Desc"] = _compute_raw_extra_columns(work)["SKU Desc"]
    gross = _safe_numeric(raw[_HYBRID_RETURN_D001_GROSS_INTERNAL])
    work[_HYBRID_RETURN_D001_GROSS_COLUMN] = gross.reindex(work.index).fillna(0).astype(int)
    work["轉出量"] = _safe_numeric(work.get("轉出量", _safe_default_series(0, work.index))).fillna(0).astype(int)
    work[_HYBRID_RETURN_D001_COLUMN] = qty.reindex(work.index).fillna(0).astype(int)
    work = work.sort_values(["Site Code", "Brand", "SKU"], kind="stable")
    return work[[c for c in cols if c in work.columns]].reset_index(drop=True)



# 簡易調貨明細欄序（一列一筆「來源店舖 → 接收店舖」調貨；RP Type 緊接對應 Shop Name，
# 調貨數量置於 Target Site Code 前作 Supply／Target 分隔，先轉出後接收；v1.22.0 起 12 欄）
_HYBRID_TRANSFER_DETAIL_COLUMNS = [
    "Brand", "SKU", "SKU Desc",
    "Supply Site Code(調貨來源)", "Supply Shop Name", "Supply RP Type", "Supply OM",
    "調貨數量",
    "Target Site Code(接收店舖)", "Target Shop Name", "Target RP Type", "Target OM",
]



def _build_hybrid_transfer_detail(raw: pd.DataFrame, variant: str = "free") -> pd.DataFrame:
    """簡易調貨明細：只收錄「需要調貨=是」且實際有調入的列，把調貨來源字串
    （如「HB30:2; HB49:3」）逐段拆行，一列一筆來源店→接收店調貨。
    variant="om" 時改用同OM調貨欄位；同一接收列同一來源店重複出現時合併加總。
    來源店名稱／OM／RP Type 以 RAW 全量建立 Site Code 查找表補回，查不到留空。"""
    cols = _HYBRID_TRANSFER_DETAIL_COLUMNS
    if raw is None or raw.empty or "需要調貨" not in raw.columns:
        return pd.DataFrame(columns=cols)
    if variant == "om":
        src_col, qty_col = "調貨來源 (同OM)", "實際調入量 (同OM)"
    else:
        src_col, qty_col = "調貨來源 (Site Code)", "實際調入量"
    if src_col not in raw.columns or qty_col not in raw.columns:
        return pd.DataFrame(columns=cols)

    need = raw["需要調貨"].fillna("").astype(str).str.strip() == "是"
    src = raw[src_col].fillna("").astype(str).str.strip()
    qty_actual = _safe_numeric(raw[qty_col])
    sel = raw[need & (src != "") & (qty_actual > 0)]
    if sel.empty:
        return pd.DataFrame(columns=cols)

    site_all = raw.get("Site Code", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    if "Shop Name" in raw.columns:
        name_all = raw["Shop Name"].fillna("").astype(str).str.strip()
    elif "Item_Shop" in raw.columns:
        name_all = pd.Series(
            [_extract_shop_name(raw["Item_Shop"].iloc[i], site_all.iloc[i]) for i in range(len(raw))],
            index=raw.index,
        )
    else:
        name_all = _safe_default_series("", raw.index)
    om_all = raw.get("OM", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    rp_all = (
        raw["RP Type"].fillna("").astype(str).str.strip().str.upper()
        if "RP Type" in raw.columns
        else _safe_default_series("", raw.index)
    )

    name_map: dict[str, str] = {}
    om_map: dict[str, str] = {}
    rp_map: dict[str, str] = {}
    for site_v, name_v, om_v, rp_v in zip(
        site_all.to_numpy(), name_all.to_numpy(), om_all.to_numpy(), rp_all.to_numpy()
    ):
        if site_v == "":
            continue
        if name_v and site_v not in name_map:
            name_map[site_v] = str(name_v)
        if om_v and site_v not in om_map:
            om_map[site_v] = str(om_v)
        if rp_v and site_v not in rp_map:
            rp_map[site_v] = str(rp_v)

    rows = []
    for idx_pos, row in sel.iterrows():
        site_v = str(row.get("Site Code", "")).strip()
        target_name = str(row.get("Shop Name", "")).strip() if "Shop Name" in sel.columns else ""
        if not target_name:
            target_name = name_map.get(site_v, "")
        target_om = str(row.get("OM", "")).strip() or om_map.get(site_v, "")
        target_rp = str(row.get("RP Type", "")).strip().upper() or rp_map.get(site_v, "")
        per_supply: dict[str, int] = {}
        for part in str(row[src_col]).split(";"):
            part = part.strip()
            if not part:
                continue
            supply_site, sep, qty_text = part.rpartition(":")
            if not sep:
                continue
            supply_site = supply_site.strip()
            try:
                qty_n = int(float(qty_text.strip()))
            except ValueError:
                continue
            if not supply_site or qty_n <= 0:
                continue
            per_supply[supply_site] = per_supply.get(supply_site, 0) + qty_n
        for supply_site, qty_n in per_supply.items():
            rows.append({
                "Brand": row.get("Brand", ""),
                "SKU": row.get("SKU", ""),
                "SKU Desc": row.get("SKU Desc", ""),
                "Target Site Code(接收店舖)": site_v,
                "Target Shop Name": target_name,
                "Target RP Type": target_rp,
                "Target OM": target_om,
                "調貨數量": qty_n,
                "Supply Site Code(調貨來源)": supply_site,
                "Supply Shop Name": name_map.get(supply_site, ""),
                "Supply RP Type": rp_map.get(supply_site, ""),
                "Supply OM": om_map.get(supply_site, ""),
            })
    if not rows:
        return pd.DataFrame(columns=cols)
    out = pd.DataFrame(rows, columns=cols)
    out = out.sort_values(
        ["Brand", "SKU", "Supply Site Code(調貨來源)", "Target Site Code(接收店舖)"],
        kind="stable",
    ).reset_index(drop=True)
    return out



def _hybrid_diag_summary_columns(dimension: str = "nature", om_groups=None) -> list[str]:
    cols = ["Safety Review"]
    for g in _hybrid_dimension_groups(dimension, om_groups=om_groups):
        cols += [f"{g} 列數", f"{g} SKU 數", f"{g} 建議調整量合計"]
    cols += ["合計 列數", "合計 SKU 數", "合計 建議調整量"]
    return cols






def _build_hybrid_diag_summary(raw: pd.DataFrame, dimension: str = "nature", om_groups=None) -> pd.DataFrame:
    """Safety Review × 分類維度（預設 Shop Nature）交叉表（列數 / SKU 數 / 建議調整量合計），含總計列。
    om／nature_om 維度由 OM 欄推導動態分組（om_groups 缺省時由 raw 即時推導）。"""
    dim = _hybrid_dimension(dimension)
    if dim in ("om", "nature_om") and om_groups is None and raw is not None and not raw.empty:
        om_groups = _hybrid_om_groups_from_values(
            raw["OM"] if "OM" in raw.columns else _safe_default_series("", raw.index)
        )
    cols = _hybrid_diag_summary_columns(dim, om_groups=om_groups)
    groups = _hybrid_dimension_groups(dim, om_groups=om_groups)
    if raw is None or raw.empty or "事件類型" not in raw.columns:
        return pd.DataFrame(columns=cols)
    df = raw[raw["事件類型"].fillna("").astype(str) != ""].copy()
    if df.empty:
        return pd.DataFrame(columns=cols)
    df["_chg"] = pd.to_numeric(
        df.get("Safety 建議調整量", _safe_default_series(np.nan, df.index)), errors="coerce"
    ).fillna(0)
    grp = _hybrid_group_key_series(df, dim, om_groups=om_groups)
    diag = df["Safety Review"].fillna("").astype(str)

    def _agg(sub: pd.DataFrame) -> dict:
        return {
            "列數": int(len(sub)),
            "SKU 數": int(sub["SKU"].nunique()) if "SKU" in sub.columns else 0,
            # 先四捨五入再取整，與「Safety 調整方向分類」／stats 合計準則一致
            "建議調整量合計": int(round(float(sub["_chg"].sum()))),
        }

    _total_key = {"列數": "合計 列數", "SKU 數": "合計 SKU 數", "建議調整量合計": "合計 建議調整量"}

    rows = []
    for d in _HYBRID_DIAG_ORDER:
        sub_d = df[diag == d]
        row = {"Safety Review": d}
        for g in groups:
            for k, v in _agg(sub_d[grp.loc[sub_d.index] == g]).items():
                row[f"{g} {k}"] = v
        for k, v in _agg(sub_d).items():
            row[_total_key[k]] = v
        rows.append(row)
    total_row = {"Safety Review": "總計"}
    for g in groups:
        for k, v in _agg(df[grp == g]).items():
            total_row[f"{g} {k}"] = v
    for k, v in _agg(df).items():
        total_row[_total_key[k]] = v
    rows.append(total_row)
    return pd.DataFrame(rows, columns=cols)



# 「銷量 × 存貨分類」交叉表欄序（Review 摘要區塊⑤；欄名屬新契約，未經要求不得改動）
_HYBRID_SALES_SUMMARY_COLUMNS = [
    "銷量存貨分類", "Tourist 列數", "Mix 列數", "Non Tourist 列數",
    "總計 列數", "總計 SKU 數", "總計 店舖數",
    "庫存合計（件）", "Sold Qty (P90D) 合計（件）",
]



def _build_hybrid_sales_summary(raw: pd.DataFrame) -> pd.DataFrame:
    """銷量 × 存貨分類交叉表（Review 摘要區塊⑤）——母體為**全部店舖×SKU 列**
    （有別於 ①–④ 的調貨事件列母體），無調貨事件時仍可查閱高銷缺貨與積壓分布。
    列：高銷缺貨／低銷積壓／零銷積壓／正常／總計；各 Nature 列數按 Shop Nature 拆分，
    SKU 數與店舖數為不重複計數（店舖數＝Site Code 不重複，空白不計），
    庫存與 P90D 銷量為件數合計。分類欄為 transform_hybrid 常開產出；缺欄時回傳空表。"""
    cols = _HYBRID_SALES_SUMMARY_COLUMNS
    if raw is None or raw.empty or _HYBRID_SALES_CLASS_COLUMN not in raw.columns:
        return pd.DataFrame(columns=cols)
    cls = raw[_HYBRID_SALES_CLASS_COLUMN].fillna("").astype(str)
    nature = raw.get("Shop Nature", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    stock = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))
    sold = _safe_numeric(raw.get("Sold Qty (P90D)", _safe_default_series(0, raw.index)))
    site = raw.get("Site Code", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    rows = []
    for label in (*_HYBRID_SALES_CLASS_ORDER, "總計"):
        m = pd.Series(True, index=raw.index) if label == "總計" else cls.eq(label)
        sub = raw[m]
        site_sub = site[m]
        rows.append({
            "銷量存貨分類": label,
            "Tourist 列數": int((m & (nature == "Tourist")).sum()),
            "Mix 列數": int((m & (nature == "Mix")).sum()),
            "Non Tourist 列數": int((m & (nature == "Non Tourist")).sum()),
            "總計 列數": int(m.sum()),
            "總計 SKU 數": int(sub["SKU"].nunique()) if "SKU" in sub.columns else 0,
            "總計 店舖數": int(site_sub[site_sub != ""].nunique()),
            "庫存合計（件）": int(stock[m].sum()),
            "Sold Qty (P90D) 合計（件）": int(sold[m].sum()),
        })
    return pd.DataFrame(rows, columns=cols)



def _hybrid_safety_event_frame(raw: pd.DataFrame) -> pd.DataFrame:
    """Safety 調整統計母體 — 僅「事件類型」非空的列（與 Safety Review × Shop Nature
    交叉表同母體）。Safety 建議不受 OM 限制影響，故不限OM／同OM 共用同一份統計。"""
    if raw is None or raw.empty or "事件類型" not in raw.columns:
        return pd.DataFrame()
    return raw[raw["事件類型"].fillna("").astype(str) != ""].copy()



def _first_non_blank_value(values) -> str:
    """取群組內首個非空值（SKU Desc／Buyer 每 SKU 應一致，容許個別列空白）。"""
    for v in values:
        text = "" if pd.isna(v) else str(v).strip()
        if text:
            return text
    return ""



def _build_hybrid_safety_change_summary(raw: pd.DataFrame) -> pd.DataFrame:
    """Safety 調整方向分類表 — 按「Safety 建議調整量」（顯示值，已四捨五入至 1 位小數）
    的正負分為 增加／減少／不變，統計列數、SKU 數與調整量合計，含總計列。
    邊界情形：調整量四捨五入後為 0 的列歸「不變」（其「Safety Review」仍可能標示不足／過高）。
    總計列的 SKU 數用母體不重複 SKU，避免同一 SKU 跨分類虛增。"""
    cols = _HYBRID_SAFETY_CHANGE_COLUMNS
    df = _hybrid_safety_event_frame(raw)
    if df.empty:
        return pd.DataFrame(columns=cols)
    chg = pd.to_numeric(
        df.get("Safety 建議調整量", _safe_default_series(np.nan, df.index)), errors="coerce"
    ).fillna(0)
    sku = df.get("SKU", _safe_default_series("", df.index))
    cat = np.where(chg > 0, "增加", np.where(chg < 0, "減少", "不變"))

    rows = []
    for label in _HYBRID_SAFETY_CHANGE_ORDER:
        m = cat == label
        rows.append({
            "分類": label,
            "列數": int(m.sum()),
            "SKU 數": int(sku[m].dropna().nunique()),
            "Safety 建議調整量合計": int(round(float(chg[m].sum()))),
            "說明": _HYBRID_SAFETY_CHANGE_DESC[label],
        })
    rows.append({
        "分類": "總計",
        "列數": int(len(df)),
        "SKU 數": int(sku.dropna().nunique()),
        "Safety 建議調整量合計": int(round(float(chg.sum()))),
        "說明": _HYBRID_SAFETY_CHANGE_TOTAL_DESC,
    })
    return pd.DataFrame(rows, columns=cols)



def _build_hybrid_safety_sku_detail(raw: pd.DataFrame) -> pd.DataFrame:
    """Safety 調整明細（SKU 級彙總）— 每個有調貨事件的 SKU 一列：
    Review 列數、需增加／需減少／設定正常店舖數（按 Site Code 不重複計）、
    調整列（設定不足＋設定過高）的現有與建議 Safety 設定合計，以及淨調整量。
    只有「設定正常」列的 SKU 仍會納入（淨調整量為 0，代表缺貨／過剩非 Safety 設定因素）。"""
    cols = _HYBRID_SAFETY_SKU_COLUMNS
    df = _hybrid_safety_event_frame(raw)
    if df.empty:
        return pd.DataFrame(columns=cols)

    diag = df.get("Safety Review", _safe_default_series("", df.index)).fillna("").astype(str)
    chg = pd.to_numeric(
        df.get("Safety 建議調整量", _safe_default_series(np.nan, df.index)), errors="coerce"
    ).fillna(0)
    adj = pd.to_numeric(
        df.get("Adj. Safety Stock", _safe_default_series(np.nan, df.index)), errors="coerce"
    ).fillna(0)
    # 「建議 Safety 設定值」只落在不足／過高列，故其非空即代表該列需要調整
    sugg = pd.to_numeric(
        df.get("建議 Safety 設定值", _safe_default_series(np.nan, df.index)), errors="coerce"
    )
    work = pd.DataFrame({
        "Brand": df.get("Brand", _safe_default_series("", df.index)).fillna("").astype(str),
        "SKU": df.get("SKU", _safe_default_series("", df.index)),
        "SKU Desc": df.get("SKU Desc", _safe_default_series("", df.index)),
        "Buyer": df.get("Buyer", _safe_default_series("", df.index)),
        "_site": df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip(),
        "_chg": chg,
        "_adj": adj,
        "_sugg": sugg,
        "_under": (diag == _HYBRID_DIAG_UNDER).to_numpy(),
        "_over": (diag == _HYBRID_DIAG_OVER).to_numpy(),
    })

    rows = []
    for (_brand, _sku), grp in work.groupby(["Brand", "SKU"], sort=False, dropna=False):
        under = grp["_under"]
        over = grp["_over"]
        fix = grp[grp["_sugg"].notna()]
        rows.append({
            "Brand": _brand,
            "SKU": _sku,
            "SKU Desc": _first_non_blank_value(grp["SKU Desc"]),
            "Buyer": _first_non_blank_value(grp["Buyer"]),
            "Review 列數": int(len(grp)),
            "需增加店舖數": int(grp.loc[under, "_site"].nunique()),
            "需減少店舖數": int(grp.loc[over, "_site"].nunique()),
            "設定正常店舖數": int(grp.loc[~under & ~over, "_site"].nunique()),
            "現有 Safety 設定合計（調整列）": int(round(float(fix["_adj"].sum()))),
            "建議 Safety 設定合計（調整列）": int(round(float(fix["_sugg"].sum()))),
            "淨調整量": int(round(float(grp["_chg"].sum()))),
        })
    out = pd.DataFrame(rows, columns=cols)
    return out.sort_values(["Brand", "SKU"], kind="stable").reset_index(drop=True)



# ── 全列預防性 Safety 設定建議（母體為全部店舖×SKU 列；欄名屬新契約）──
_HYBRID_FULL_SAFETY_CHANGE_COLUMNS = ["分類", "列數", "SKU 數", "全列 Safety 建議調整量合計", "說明"]



_HYBRID_FULL_SAFETY_SKU_COLUMNS = [
    "Brand", "SKU", "SKU Desc", "Buyer",
    "全列數", "需增加店舖數", "需減少店舖數", "設定正常店舖數",
    "現有 Safety 設定合計（調整列）", "建議 Safety 設定合計（調整列）", "淨調整量",
]



def _build_hybrid_full_safety_change_summary(raw: pd.DataFrame) -> pd.DataFrame:
    """全列預防性調整方向分類 — 以全部店舖×SKU 列為母體，按「全列 Safety 建議調整量」
    正負分為 增加／減少／不變，統計列數、SKU 數與調整量合計，含總計列。
    既有事件版統計（_build_hybrid_safety_change_summary）維持不變。"""
    cols = _HYBRID_FULL_SAFETY_CHANGE_COLUMNS
    if raw is None or raw.empty or _HYBRID_FULL_SAFETY_CHG_COLUMN not in raw.columns:
        return pd.DataFrame(columns=cols)
    chg = pd.to_numeric(raw[_HYBRID_FULL_SAFETY_CHG_COLUMN], errors="coerce").fillna(0)
    sku = raw.get("SKU", _safe_default_series("", raw.index))
    cat = np.where(chg > 0, "增加", np.where(chg < 0, "減少", "不變"))
    rows = []
    for label in _HYBRID_SAFETY_CHANGE_ORDER:
        m = cat == label
        rows.append({
            "分類": label,
            "列數": int(m.sum()),
            "SKU 數": int(sku[m].dropna().nunique()),
            "全列 Safety 建議調整量合計": int(round(float(chg[m].sum()))),
            "說明": _HYBRID_SAFETY_CHANGE_DESC[label],
        })
    rows.append({
        "分類": "總計",
        "列數": int(len(raw)),
        "SKU 數": int(sku.dropna().nunique()),
        "全列 Safety 建議調整量合計": int(round(float(chg.sum()))),
        "說明": "SKU 數為全表不重複 SKU；合計為淨調整量",
    })
    return pd.DataFrame(rows, columns=cols)



def _build_hybrid_full_safety_sku_detail(raw: pd.DataFrame) -> pd.DataFrame:
    """全列預防性調整明細（SKU 級彙總）— 每個 SKU 一列：全列數、需增加／需減少／
    設定正常店舖數（按 Site Code 不重複計）、調整列的現有與建議合計，以及淨調整量。
    只有設定正常列的 SKU 仍會納入（淨調整量為 0）。"""
    cols = _HYBRID_FULL_SAFETY_SKU_COLUMNS
    if raw is None or raw.empty or _HYBRID_FULL_SAFETY_CHG_COLUMN not in raw.columns:
        return pd.DataFrame(columns=cols)
    review = raw.get(_HYBRID_FULL_SAFETY_REVIEW_COLUMN, _safe_default_series("", raw.index)).fillna("").astype(str)
    chg = pd.to_numeric(raw[_HYBRID_FULL_SAFETY_CHG_COLUMN], errors="coerce").fillna(0)
    adj = pd.to_numeric(
        raw.get("Adj. Safety Stock", _safe_default_series(np.nan, raw.index)), errors="coerce"
    ).fillna(0)
    sugg = pd.to_numeric(raw.get(_HYBRID_FULL_SAFETY_SUGGEST_COLUMN, _safe_default_series(np.nan, raw.index)), errors="coerce")
    work = pd.DataFrame({
        "Brand": raw.get("Brand", _safe_default_series("", raw.index)).fillna("").astype(str),
        "SKU": raw.get("SKU", _safe_default_series("", raw.index)),
        "SKU Desc": raw.get("SKU Desc", _safe_default_series("", raw.index)),
        "Buyer": raw.get("Buyer", _safe_default_series("", raw.index)),
        "_site": raw.get("Site Code", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip(),
        "_chg": chg,
        "_adj": adj,
        "_sugg": sugg,
        "_under": (review == _HYBRID_DIAG_UNDER).to_numpy(),
        "_over": (review == _HYBRID_DIAG_OVER).to_numpy(),
    })
    rows = []
    for (_brand, _sku), grp in work.groupby(["Brand", "SKU"], sort=False, dropna=False):
        under = grp["_under"]
        over = grp["_over"]
        fix = grp[grp["_sugg"].notna()]
        rows.append({
            "Brand": _brand,
            "SKU": _sku,
            "SKU Desc": _first_non_blank_value(grp["SKU Desc"]),
            "Buyer": _first_non_blank_value(grp["Buyer"]),
            "全列數": int(len(grp)),
            "需增加店舖數": int(grp.loc[under, "_site"].nunique()),
            "需減少店舖數": int(grp.loc[over, "_site"].nunique()),
            "設定正常店舖數": int(grp.loc[~under & ~over, "_site"].nunique()),
            "現有 Safety 設定合計（調整列）": int(round(float(fix["_adj"].sum()))),
            "建議 Safety 設定合計（調整列）": int(round(float(fix["_sugg"].sum()))),
            "淨調整量": int(round(float(grp["_chg"].sum()))),
        })
    out = pd.DataFrame(rows, columns=cols)
    return out.sort_values(["Brand", "SKU"], kind="stable").reset_index(drop=True)



def _build_hybrid_full_safety_detail(raw: pd.DataFrame) -> pd.DataFrame:
    """全列預防性設定建議明細（店舖×SKU 逐列）：只收錄「全列 Safety 建議調整量」非 0 的列
    （即需要增加或減少的列），按 Brand / SKU / Site Code 排序。
    RAW 缺全列欄時輸出空表。"""
    cols = _HYBRID_FULL_SAFETY_DETAIL_COLUMNS
    if raw is None or raw.empty or _HYBRID_FULL_SAFETY_CHG_COLUMN not in raw.columns:
        return pd.DataFrame(columns=cols)
    chg = pd.to_numeric(raw[_HYBRID_FULL_SAFETY_CHG_COLUMN], errors="coerce").fillna(0)
    sel_mask = chg.ne(0).fillna(False)
    if not bool(sel_mask.any()):
        return pd.DataFrame(columns=cols)
    work = _shop_issue_base_frame(raw, sel_mask)
    for c in ("Brand", "SKU", "Buyer", "SKU Desc", "Shop Name", "OM", "Shop Nature", "RP Type",
              "Shop Stock Qty (CM)", "Sold Qty (P90D)", "Adj. Safety Stock",
              _HYBRID_FULL_SAFETY_LINE_COLUMN, _HYBRID_FULL_SAFETY_REVIEW_COLUMN,
              _HYBRID_FULL_SAFETY_SUGGEST_COLUMN, _HYBRID_FULL_SAFETY_TURNDAY_COLUMN,
              _HYBRID_FULL_SAFETY_CHG_COLUMN, "事件類型", "Safety Review"):
        if c not in work.columns and c in raw.columns:
            work[c] = raw.loc[work.index, c]
    if "SKU Desc" not in work.columns or work["SKU Desc"].fillna("").astype(str).str.strip().eq("").all():
        try:
            work["SKU Desc"] = _compute_raw_extra_columns(work)["SKU Desc"]
        except Exception:
            pass
    work = work.sort_values(["Brand", "SKU", "Site Code"], kind="stable")
    return work[[c for c in cols if c in work.columns]].reset_index(drop=True)



# ── Buyer 統計（欄名屬新契約，未經要求不得改動）──
_HYBRID_BUYER_COMBINED_TITLE = "Buyer 統計（需調貨＋需調整 Safety）"



def _hybrid_buyer_need_stats(raw: pd.DataFrame, need_series: pd.Series | None = None) -> list[tuple[str, int]]:
    """Buyer 級統計：各 Buyer 涉及多少個「需要調貨」SKU（已刪除重複）。

    與 Stock Checking 的 `_stock_checking_buyer_stats` 同一準則（Buyer 取同 SKU 首個非空值；
    全空歸入「(未填 Buyer)」），供 Safety Stock x NetStock Review 沿用。
    回傳 [(Buyer, 需調貨 SKU 數)]，按數量降序、同數按 Buyer 名排序。
    """
    if raw is None or raw.empty or "SKU" not in raw.columns:
        return []
    if need_series is None:
        need_series = raw.get("需要調貨", _safe_default_series("", raw.index)).astype(str) == "是"
    need_skus = set(raw.loc[need_series, "SKU"].dropna().unique())
    if not need_skus:
        return []
    buyer = (
        raw["Buyer"].fillna("").astype(str).str.strip()
        if "Buyer" in raw.columns
        else _safe_default_series("", raw.index)
    )
    tmp = pd.DataFrame({"_sku": raw["SKU"], "_buyer": buyer})
    tmp = tmp[tmp["_sku"].isin(need_skus)]

    def _first_nonempty(s: pd.Series) -> str:
        s = s[s != ""]
        return s.iloc[0] if len(s) else "(未填 Buyer)"

    per_sku_buyer = tmp.groupby("_sku", sort=False)["_buyer"].agg(_first_nonempty)
    counts = per_sku_buyer.value_counts()
    return sorted(((str(b), int(c)) for b, c in counts.items()), key=lambda kv: (-kv[1], kv[0]))



def _hybrid_buyer_safety_stats(raw: pd.DataFrame) -> list[tuple[str, int, int]]:
    """Buyer 級統計：各 Buyer 涉及多少個「需要修改 Safety」SKU（已刪除重複）。

    母體為 Safety Review 屬「Safety 設定不足」或「Safety 設定過高」的列（即有
    「建議 Safety 設定值」的列，與 Safety 調整方向分類的增加＋減少同母體）。
    Buyer 取同 SKU 首個非空值（與需調貨統計同一準則）；全空歸入「(未填 Buyer)」。
    調整量合計為該 Buyer 名下各需調整 SKU 的「Safety 建議調整量」列合計（四捨五入取整，
    與交叉表／分類表同一準則；正=增加、負=減少）。
    回傳 [(Buyer, 需調整 Safety SKU 數, Safety 建議調整量合計)]，
    按 SKU 數降序、同數按 Buyer 名排序。
    """
    if raw is None or raw.empty or "SKU" not in raw.columns:
        return []
    diag = raw.get("Safety Review", _safe_default_series("", raw.index)).fillna("").astype(str)
    fix_mask = diag.isin([_HYBRID_DIAG_UNDER, _HYBRID_DIAG_OVER])
    if not bool(fix_mask.any()):
        return []
    chg = pd.to_numeric(
        raw.get("Safety 建議調整量", _safe_default_series(np.nan, raw.index)), errors="coerce"
    ).fillna(0.0)
    buyer = (
        raw["Buyer"].fillna("").astype(str).str.strip()
        if "Buyer" in raw.columns
        else _safe_default_series("", raw.index)
    )
    tmp = pd.DataFrame({
        "_sku": raw.loc[fix_mask, "SKU"],
        "_buyer": buyer.loc[fix_mask],
        "_chg": chg.loc[fix_mask],
    })
    tmp = tmp[tmp["_sku"].notna()]
    if tmp.empty:
        return []

    def _first_nonempty(s: pd.Series) -> str:
        s = s[s != ""]
        return s.iloc[0] if len(s) else "(未填 Buyer)"

    per_sku_buyer = tmp.groupby("_sku", sort=False)["_buyer"].agg(_first_nonempty)
    per_sku_chg = tmp.groupby("_sku", sort=False)["_chg"].sum()
    per_buyer_skus: dict[str, int] = {}
    per_buyer_chg: dict[str, float] = {}
    for sku, b in per_sku_buyer.items():
        b = str(b)
        per_buyer_skus[b] = per_buyer_skus.get(b, 0) + 1
        per_buyer_chg[b] = per_buyer_chg.get(b, 0.0) + float(per_sku_chg.loc[sku])
    result = [
        (b, int(per_buyer_skus[b]), int(round(per_buyer_chg[b])))
        for b in per_buyer_skus
    ]
    return sorted(result, key=lambda kv: (-kv[1], kv[0]))



def _hybrid_buyer_combined_stats(
    buyer_need: list[tuple[str, int]] | None,
    buyer_safety: list[tuple[str, int, int]] | None,
) -> list[tuple[str, int, int, int]]:
    """合併 Buyer 統計：將需調貨與需調整 Safety 按 Buyer 外連接為單一列。

    輸入為 `_hybrid_buyer_need_stats` 與 `_hybrid_buyer_safety_stats` 的回傳；
    任一為空或缺席的 Buyer 以 0 補齊。回傳
    [(Buyer, 需調貨 SKU 數, 需調整 Safety SKU 數, Safety 建議調整量合計)]，
    按需調貨數降序、再按需調整數降序、同數按 Buyer 名排序。
    （保留供相容；新合併表含貨量與佔比，請改用 `_hybrid_buyer_combined_detail`。）
    """
    need_map: dict[str, int] = {str(b): int(c) for b, c in (buyer_need or [])}
    safety_map: dict[str, tuple[int, int]] = {
        str(b): (int(c), int(v)) for b, c, v in (buyer_safety or [])
    }
    buyers = sorted(
        set(need_map) | set(safety_map),
        key=lambda b: (-need_map.get(b, 0), -safety_map.get(b, (0, 0))[0], b),
    )
    return [
        (b, int(need_map.get(b, 0)), int(safety_map.get(b, (0, 0))[0]), int(safety_map.get(b, (0, 0))[1]))
        for b in buyers
    ]



def _hybrid_buyer_combined_detail(raw: pd.DataFrame) -> list[tuple[str, int, int, int, int]]:
    """Buyer 合併明細：同一 Buyer 的需調貨與需調整 Safety 並列（已刪除重複）。

    Buyer 取同 SKU 全表首個非空值（需調貨與需調整共用同一歸屬，避免同一 SKU
    在兩母體歸入不同 Buyer）；全空歸入「(未填 Buyer)」。
    需調貨貨量為該 Buyer 名下各需調貨 SKU 的「調入需求量 (補到14日)」列合計
    （取整，與指標卡總需求同一準則）；Safety 調整量為「Safety 建議調整量」列合計
    （取整，正=增加、負=減少）。
    回傳 [(Buyer, 需調貨 SKU 數, 調入需求量合計, 需調整 Safety SKU 數,
    Safety 建議調整量合計)]，按需調貨數降序、再按需調整數降序、同數按 Buyer 名排序。
    """
    if raw is None or raw.empty or "SKU" not in raw.columns:
        return []
    buyer = (
        raw["Buyer"].fillna("").astype(str).str.strip()
        if "Buyer" in raw.columns
        else _safe_default_series("", raw.index)
    )
    sku = raw["SKU"]

    def _first_nonempty(s: pd.Series) -> str:
        s = s[s != ""]
        return s.iloc[0] if len(s) else "(未填 Buyer)"

    per_sku_buyer = pd.DataFrame({"_sku": sku, "_buyer": buyer}).groupby("_sku", sort=False)["_buyer"].agg(_first_nonempty)

    need_mask = raw.get("需要調貨", _safe_default_series("", raw.index)).astype(str) == "是"
    demand = pd.to_numeric(
        raw.get("調入需求量 (補到14日)", _safe_default_series(np.nan, raw.index)), errors="coerce"
    ).fillna(0.0)
    need_skus = set(raw.loc[need_mask, "SKU"].dropna().unique()) if bool(need_mask.any()) else set()
    per_sku_demand = demand.loc[need_mask].groupby(raw.loc[need_mask, "SKU"], sort=False).sum() if need_skus else pd.Series(dtype=float)

    diag = raw.get("Safety Review", _safe_default_series("", raw.index)).fillna("").astype(str)
    fix_mask = diag.isin([_HYBRID_DIAG_UNDER, _HYBRID_DIAG_OVER])
    chg = pd.to_numeric(
        raw.get("Safety 建議調整量", _safe_default_series(np.nan, raw.index)), errors="coerce"
    ).fillna(0.0)
    fix_skus = set(raw.loc[fix_mask, "SKU"].dropna().unique()) if bool(fix_mask.any()) else set()
    per_sku_chg = chg.loc[fix_mask].groupby(raw.loc[fix_mask, "SKU"], sort=False).sum() if fix_skus else pd.Series(dtype=float)

    if not need_skus and not fix_skus:
        return []
    per_buyer: dict[str, list[float]] = {}
    for s in need_skus | fix_skus:
        b = str(per_sku_buyer.get(s, "(未填 Buyer)"))
        rec = per_buyer.setdefault(b, [0, 0.0, 0, 0.0])
        if s in need_skus:
            rec[0] += 1
            rec[1] += float(per_sku_demand.get(s, 0.0))
        if s in fix_skus:
            rec[2] += 1
            rec[3] += float(per_sku_chg.get(s, 0.0))
    result = [
        (b, int(v[0]), int(round(v[1])), int(v[2]), int(round(v[3])))
        for b, v in per_buyer.items()
    ]
    return sorted(result, key=lambda kv: (-kv[1], -kv[3], kv[0]))



def _hybrid_safety_days_from_raw(raw: pd.DataFrame, nature=None) -> dict[str, int]:
    """由 RAW 的「Safety 基準天數」欄還原各 Nature 實際採用的基準天數
    （該欄每列同一 Nature 恆定）；缺欄時退回預設 7 / 9 / 10。"""
    days, _ = _hybrid_safety_days({})
    if raw is None or raw.empty or _HYBRID_SAFETY_DAYS_COLUMN not in raw.columns:
        return days
    if nature is None:
        nature = raw.get("Shop Nature", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    vals = pd.to_numeric(raw[_HYBRID_SAFETY_DAYS_COLUMN], errors="coerce")
    for n in _HYBRID_DIAG_NATURES:
        sub = vals[nature == n].dropna()
        if len(sub):
            days[n] = int(sub.iloc[0])
    return days



def _hybrid_stats_from_raw(raw: pd.DataFrame) -> dict:
    """由 raw 匯總 Review stats（stats 缺失時的回退，與 transform_hybrid 同一準則）。"""
    if raw is None or raw.empty:
        return {}
    need = raw.get("需要調貨", _safe_default_series("", raw.index)).astype(str) == "是"
    evt = raw.get("事件類型", _safe_default_series("", raw.index)).fillna("").astype(str)
    diag = raw.get("Safety Review", _safe_default_series("", raw.index)).fillna("").astype(str)
    chg = pd.to_numeric(raw.get("Safety 建議調整量", _safe_default_series(np.nan, raw.index)), errors="coerce")
    nature = raw.get("Shop Nature", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    bidirectional = "可轉出貨量 (Tourist)" in raw.columns
    rp_in_raw = "RP Type" in raw.columns
    raw_rp = (
        raw["RP Type"].fillna("").astype(str).str.strip()
        if rp_in_raw else _safe_default_series("", raw.index)
    )
    under = diag == _HYBRID_DIAG_UNDER
    over = diag == _HYBRID_DIAG_OVER
    donor = evt.isin([_HYBRID_EVENT_DONATED, _HYBRID_EVENT_DONOR])
    cls_col = raw.get(_HYBRID_SALES_CLASS_COLUMN, _safe_default_series("", raw.index)).fillna("").astype(str)
    _full_chg = pd.to_numeric(raw.get(_HYBRID_FULL_SAFETY_CHG_COLUMN, _safe_default_series(np.nan, raw.index)), errors="coerce")
    _full_review = raw.get(_HYBRID_FULL_SAFETY_REVIEW_COLUMN, _safe_default_series("", raw.index)).fillna("").astype(str)
    _full_under = _full_review == _HYBRID_DIAG_UNDER
    _full_over = _full_review == _HYBRID_DIAG_OVER
    return {
        "safety_days": _hybrid_safety_days_from_raw(raw, nature),
        "dimension": "nature",
        "om_groups": list(_hybrid_om_groups_from_values(
            raw["OM"] if raw is not None and not raw.empty and "OM" in raw.columns
            else _safe_default_series("", raw.index if raw is not None else []))),
        "need_rows": int(need.sum()),
        "need_skus": int(raw.loc[need, "SKU"].nunique()) if need.any() else 0,
        "total_demand": int(_safe_numeric(raw.get("調入需求量 (補到14日)", _safe_default_series(0, raw.index))).sum()),
        "total_actual": int(_safe_numeric(raw.get("實際調入量", _safe_default_series(0, raw.index))).sum()),
        "total_gap": int(_safe_numeric(raw.get("調入缺口", _safe_default_series(0, raw.index))).sum()),
        "nt_transfer": int(_safe_numeric(raw.loc[nature == "Non Tourist", "轉出量"]).sum()) if "轉出量" in raw.columns else 0,
        "mix_transfer": int(_safe_numeric(raw.loc[nature == "Mix", "轉出量"]).sum()) if "轉出量" in raw.columns else 0,
        "tourist_transfer": int(_safe_numeric(raw.loc[nature == "Tourist", "轉出量"]).sum()) if bidirectional and "轉出量" in raw.columns else 0,
        "total_actual_om": int(_safe_numeric(raw.get("實際調入量 (同OM)", _safe_default_series(0, raw.index))).sum()),
        "total_gap_om": int(_safe_numeric(raw.get("調入缺口 (同OM)", _safe_default_series(0, raw.index))).sum()),
        "nt_transfer_om": int(_safe_numeric(raw.loc[nature == "Non Tourist", "轉出量 (同OM)"]).sum()) if "轉出量 (同OM)" in raw.columns else 0,
        "mix_transfer_om": int(_safe_numeric(raw.loc[nature == "Mix", "轉出量 (同OM)"]).sum()) if "轉出量 (同OM)" in raw.columns else 0,
        "tourist_transfer_om": int(_safe_numeric(raw.loc[nature == "Tourist", "轉出量 (同OM)"]).sum()) if bidirectional and "轉出量 (同OM)" in raw.columns else 0,
        "bidirectional": bool(bidirectional),
        # ND 來源轉出量（與 transform_hybrid 同一準則；nd_priority 無法由 raw 反推，故不列）
        "nd_transfer": int(_safe_numeric(raw.loc[raw_rp == "ND", "轉出量"]).sum()) if rp_in_raw and "轉出量" in raw.columns else 0,
        "nd_transfer_om": int(_safe_numeric(raw.loc[raw_rp == "ND", "轉出量 (同OM)"]).sum()) if rp_in_raw and "轉出量 (同OM)" in raw.columns else 0,
        "donor_rows": int(donor.sum()),
        "under_rows": int(under.sum()),
        "under_skus": int(raw.loc[under, "SKU"].nunique()) if under.any() else 0,
        "under_inc_total": int(round(float(chg[under].fillna(0).sum()))),
        "over_rows": int(over.sum()),
        "over_skus": int(raw.loc[over, "SKU"].nunique()) if over.any() else 0,
        "over_dec_total": int(round(float(chg[over].fillna(0).sum()))),
        "normal_rows": int((evt != "").sum()) - int(under.sum()) - int(over.sum()),
        "total_raw_rows": int(len(raw)),
        # 銷量分類列可由 RAW 反推；開關狀態與門檻為參數、無法由 RAW 反推，
        # 故回退路徑不列 sales_rebalance／sales_need_rows（銷量分類概況排自然不顯示，
        # 與 ND 優先轉出的回退準則一致）
        "sales_oos_rows": int(cls_col.eq(_HYBRID_SALES_CLASS_ORDER[0]).sum()),
        "sales_oos_skus": int(
            raw.loc[cls_col.eq(_HYBRID_SALES_CLASS_ORDER[0]), "SKU"].nunique()
        ) if "SKU" in raw.columns else 0,
        "sales_overstock_rows": int(cls_col.isin(_HYBRID_SALES_CLASS_ORDER[1:3]).sum()),
        "sales_overstock_units": int(
            _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))[
                cls_col.isin(_HYBRID_SALES_CLASS_ORDER[1:3]).to_numpy()
            ].sum()
        ),
        # 退D001 建議可由 RAW 欄位存在反推（與 bidirectional 同一準則）
        "return_to_d001": _HYBRID_RETURN_D001_COLUMN in raw.columns,
        "return_rows": int((_safe_numeric(raw.get(_HYBRID_RETURN_D001_COLUMN, _safe_default_series(0, raw.index))) > 0).sum()) if _HYBRID_RETURN_D001_COLUMN in raw.columns else 0,
        "return_skus": int(raw.loc[_safe_numeric(raw[_HYBRID_RETURN_D001_COLUMN]).fillna(0) > 0, "SKU"].nunique()) if (_HYBRID_RETURN_D001_COLUMN in raw.columns and bool((_safe_numeric(raw[_HYBRID_RETURN_D001_COLUMN]).fillna(0) > 0).any())) else 0,
        "return_units": int(_safe_numeric(raw.get(_HYBRID_RETURN_D001_COLUMN, _safe_default_series(0, raw.index))).fillna(0).sum()) if _HYBRID_RETURN_D001_COLUMN in raw.columns else 0,
        "full_under_rows": int(_full_under.sum()),
        "full_under_skus": int(raw.loc[_full_under, "SKU"].nunique()) if _full_under.any() else 0,
        "full_under_inc_total": int(round(float(_full_chg[_full_under].fillna(0).sum()))),
        "full_over_rows": int(_full_over.sum()),
        "full_over_skus": int(raw.loc[_full_over, "SKU"].nunique()) if _full_over.any() else 0,
        "full_over_dec_total": int(round(float(_full_chg[_full_over].fillna(0).sum()))),
        "full_normal_rows": int(len(raw)) - int(_full_under.sum()) - int(_full_over.sum()),
        "full_action_rows": int((_full_under | _full_over).sum()),
        "full_action_skus": int(raw.loc[_full_under | _full_over, "SKU"].nunique()) if (_full_under | _full_over).any() else 0,
        "buyer_need_skus": _hybrid_buyer_need_stats(raw, need),
        "buyer_safety_fix": _hybrid_buyer_safety_stats(raw),
    }



def _hybrid_dashboard_metrics(raw: pd.DataFrame, stats: dict | None = None):
    """調貨 Review 指標卡（每排 4 卡，與 Streamlit Hybrid Dashboard 同一準則；
    第 4 排為同OM調貨對照，Safety Review 建議不受 OM 限制影響；
    第 5 排為轉出來源對照——雙向調貨時含 Tourist 轉出量兩卡，「ND 優先轉出」開啟且 ND 店舖
    有實際轉出時另加 ND 來源轉出量兩卡（該排 2–4 卡）；
    「銷量再平衡」開啟時（stats["sales_rebalance"] 為 True）末尾另加一排銷量分類概況（4 卡；
    無第 5 排時該排為第 5 排，否則為第 6 排）。
    RAW 含「建議退D001 量」欄時（開關可由欄位存在推定）再加一排退D001 建議概況（4 卡）。
    回傳各排指標 list（4–7 排）供 _write_metric_cards 使用；stats 缺失時由 raw 自動匯總。"""
    if stats is None:
        stats = _hybrid_stats_from_raw(raw)

    def g(key: str) -> int:
        return int(stats.get(key, 0) or 0)

    bidir = bool(stats.get("bidirectional"))
    nd_prio = bool(stats.get("nd_priority"))
    row1 = [
        ("需調貨 SKU 數", g("need_skus"), "符合調貨條件的 SKU 總數（已刪除重複）", "f97316"),
        ("需要調入列數", g("need_rows"),
         "符合調貨條件的店舖×SKU 列數（任何 Nature 皆可調入）" if bidir else "符合調貨條件的 Tourist 店舖×SKU 列數", "f59e0b"),
        ("轉出來源列數", g("donor_rows"),
         "實際轉出與可轉出未使用的各 Nature 列數" if bidir else "實際轉出與可轉出未使用的 Mix / Non Tourist 列數", "8b5cf6"),
        ("調入需求量", g("total_demand"), "補到 14 日周轉的總需求（件）", "3b82f6"),
    ]
    row2 = [
        ("實際調入量", g("total_actual"),
         "不限OM：NT→Mix→Tourist 依序補足後的總調入（件）" if bidir else "不限OM：NT 優先、Mix 補足後的總調入（件）", "10b981"),
        ("無法足額缺口", g("total_gap"), "不限OM：來源可轉出量不足的缺口（件）", "ef4444"),
        ("Non Tourist 轉出量", g("nt_transfer"), "不限OM：優先由 Non Tourist 店舖轉出（件）", "8b5cf6"),
        ("Mix 轉出量", g("mix_transfer"), "不限OM：NT 不足時由 Mix 店舖補足（件）", "06b6d4"),
    ]
    row3 = [
        ("Safety 設定不足列數", g("under_rows"), "現有設定低於「Safety 周轉所需量」的調入列數（基準天數按 Nature）", "ef4444"),
        ("不足建議增加合計", g("under_inc_total"), "設定不足列的 Safety 建議調整量合計（正=增加）", "10b981"),
        ("Safety 設定過高列數", g("over_rows"), "現有設定高於生效基準（所需量補底 3 件，零銷售歸零除外）的轉出列數（基準天數按 Nature）", "f59e0b"),
        ("過高建議減少合計", g("over_dec_total"), "設定過高列的 Safety 建議調整量合計（負=減少）", "3b82f6"),
    ]
    row4 = [
        ("實際調入量 (同OM)", g("total_actual_om"), "限同 OM 調貨後的總調入（件）", "10b981"),
        ("無法足額缺口 (同OM)", g("total_gap_om"), "限同 OM 調貨後仍不足的缺口（件）", "ef4444"),
        ("Non Tourist 轉出量 (同OM)", g("nt_transfer_om"), "限同 OM：Non Tourist 店舖轉出（件）", "8b5cf6"),
        ("Mix 轉出量 (同OM)", g("mix_transfer_om"), "限同 OM：Mix 店舖轉出（件）", "06b6d4"),
    ]
    rows = [row1, row2, row3, row4]
    # 第 5 排：轉出來源對照——雙向調貨含 Tourist 轉出量兩卡；「ND 優先轉出」開啟且 ND 店舖
    # 確實轉出時另加 ND 來源轉出量兩卡（該排 2–4 卡）。ND 卡以開關狀態為前提（開關關閉時
    # 即使有 ND 列因列序恰好轉出，也不改動現行版面），stats 缺失的回退路徑無法反推開關狀態，
    # 故該路徑不顯示 ND 卡。
    row5: list = []
    if bidir:
        row5 += [
            ("Tourist 轉出量", g("tourist_transfer"), "不限OM：由 Tourist 店舖轉出（件）", "f43f5e"),
            ("Tourist 轉出量 (同OM)", g("tourist_transfer_om"), "限同 OM：Tourist 店舖轉出（件）", "ec4899"),
        ]
    if nd_prio and (g("nd_transfer") > 0 or g("nd_transfer_om") > 0):
        row5 += [
            ("ND 來源轉出量", g("nd_transfer"), "不限OM：由 RP Type＝ND 店舖轉出（件），店號見「簡易調貨明細」Supply RP Type 欄", "a855f8"),
            ("ND 來源轉出量 (同OM)", g("nd_transfer_om"), "限同 OM：RP Type＝ND 店舖轉出（件）", "7c3aed"),
        ]
    if row5:
        rows.append(row5)
    # 末尾另加銷量分類概況排（僅開關狀態為 True 時）——母體為全部店舖×SKU 列；
    # 開關狀態無法由 RAW 反推，故回退路徑（stats 缺 sales_rebalance）不顯示本排
    if bool(stats.get("sales_rebalance")):
        rows.append([
            ("高銷缺貨列數", g("sales_oos_rows"), "同 SKU 內銷量屬高銷帶且店舖周轉日數低於缺貨線的店舖×SKU 列數", "ef4444"),
            ("高銷缺貨 SKU 數", g("sales_oos_skus"), "高銷缺貨列涉及的 SKU 數（已刪除重複）", "f97316"),
            ("積壓列數", g("sales_overstock_rows"), "低銷／零銷積壓列數（周轉日數達積壓線且庫存達保底件數）", "f59e0b"),
            ("積壓庫存件數", g("sales_overstock_units"), "積壓列的 Shop Stock Qty (CM) 合計（件）", "8b5cf6"),
        ])
    # 末排：退D001 建議（RAW 含「建議退D001 量」欄即顯示——開關狀態可由欄位存在推定，
    # stats 缺失的回退路徑同樣成立）。退回量＝可轉出量扣除已轉出後的剩餘（接收目標 D001）。
    if raw is not None and _HYBRID_RETURN_D001_COLUMN in raw.columns:
        _ret = _safe_numeric(raw[_HYBRID_RETURN_D001_COLUMN]).fillna(0)
        _ret_pos = _ret > 0
        _ret_skus = int(raw.loc[_ret_pos, "SKU"].nunique()) if bool(_ret_pos.any()) else 0
        rows.append([
            ("退D001 列數", int(_ret_pos.sum()), "有退D001 建議的店舖×SKU 列數（可轉出量扣已轉出後有剩餘）", "0d9488"),
            ("退D001 SKU 數", _ret_skus, "退D001 建議涉及的不重複 SKU 數", "14b8a6"),
            ("退D001 合計件數", int(_ret.sum()), "建議退回 D001（Warehouse）的總件數（件）", "059669"),
            ("已轉出保留店舖件數", int(_safe_numeric(raw.get("轉出量", _safe_default_series(0, raw.index))).fillna(0).sum()),
             "同一計算內優先撥予店舖調貨、不退 D001 的件數（件）", "0ea5e9"),
        ])
    return rows



def _hybrid_allocate_transfers(
    n_rows: int,
    sku_arr: np.ndarray,
    site_arr: np.ndarray,
    need_mask: np.ndarray,
    demand: np.ndarray,
    nt_avail0: np.ndarray,
    mix_avail0: np.ndarray,
    om_arr: np.ndarray | None = None,
    tourist_avail0: np.ndarray | None = None,
    nature_arr: np.ndarray | None = None,
    nd_priority: np.ndarray | None = None,
):
    """調貨分配：按 SKU 彙總需求，來源池固定 Non Tourist → Mix → Tourist（tourist_avail0
    為 None 時不設 Tourist 池，即現行單向規則），需求列按需求量大→小依序分配。
    om_arr 非 None 時來源店須與需求列同 OM 才可轉出（同OM調貨）。
    nature_arr 非 None 時來源列與需求列同 Shop Nature 一律排除（「需求 Nature 自身除外」
    契約守衛；現行池級歸零已令同 Nature 來源恆為 0，此守衛為顯式防呆）。
    來源店與需求列同為一店（Site Code 相同）時一律排除，防止重複列資料異常造成自我調貨。
    nd_priority 非 None（bool 陣列，True＝該列 RP Type 為 ND）時，**池內**改依 ND 列先於
    RF 列取用（池間順序與全部轉出資格條件不變）。
    回傳 (actual_np, gap_np, donate_np, desc_np)。"""
    actual_np = np.zeros(n_rows, dtype=int)
    donate_np = np.zeros(n_rows, dtype=int)
    gap_np = np.zeros(n_rows, dtype=int)
    desc_np = np.full(n_rows, "", dtype=object)
    if not need_mask.any():
        return actual_np, gap_np, donate_np, desc_np
    nt_avail = np.asarray(nt_avail0, dtype=float).copy()
    mix_avail = np.asarray(mix_avail0, dtype=float).copy()
    tourist_avail = (
        np.asarray(tourist_avail0, dtype=float).copy() if tourist_avail0 is not None else None
    )
    nd_flag = np.asarray(nd_priority, dtype=bool) if nd_priority is not None else None

    def _nd_first(keys: list[int]) -> list[int]:
        """池內 ND 列前置（ND 群組與 RF 群組各自維持來源檔列序）；池間順序不變。"""
        if nd_flag is None:
            return keys
        return [p for p in keys if nd_flag[p]] + [p for p in keys if not nd_flag[p]]

    for sku in set(sku_arr[need_mask]):
        sku_mask = sku_arr == sku
        need_rows = np.nonzero(sku_mask & need_mask)[0]
        nt_src = [int(p) for p in np.nonzero(sku_mask & (nt_avail > 0))[0]]
        mix_src = [int(p) for p in np.nonzero(sku_mask & (mix_avail > 0))[0]]
        order = np.argsort(-demand[need_rows], kind="stable")
        remain_nt = {p: float(nt_avail[p]) for p in _nd_first(nt_src)}
        remain_mix = {p: float(mix_avail[p]) for p in _nd_first(mix_src)}
        pools = [remain_nt, remain_mix]
        if tourist_avail is not None:
            tourist_src = [int(p) for p in np.nonzero(sku_mask & (tourist_avail > 0))[0]]
            pools.append({p: float(tourist_avail[p]) for p in _nd_first(tourist_src)})
        for n in order:
            row_pos = int(need_rows[n])
            rem = float(demand[row_pos])
            parts: list[str] = []
            donated_actual = 0
            for pool in pools:
                for spos in list(pool.keys()):
                    if rem <= 0:
                        break
                    if site_arr[spos] == site_arr[row_pos]:
                        continue  # 自我調貨守衛：來源店＝需求店時排除（重複列資料異常防護）
                    if nature_arr is not None and nature_arr[spos] == nature_arr[row_pos]:
                        continue  # 需求 Nature 自身除外守衛（雙向契約，顯式防呆）
                    if om_arr is not None and om_arr[spos] != om_arr[row_pos]:
                        continue
                    avail_n = pool[spos]
                    take = min(rem, avail_n)
                    if take > 0:
                        pool[spos] -= take
                        rem -= take
                        donate_np[spos] += int(take)
                        donated_actual += int(take)
                        parts.append(f"{site_arr[spos]}:{int(take)}")
                if rem <= 0:
                    break
            actual_np[row_pos] = donated_actual
            gap_np[row_pos] = int(rem)
            desc_np[row_pos] = "; ".join(parts)
    return actual_np, gap_np, donate_np, desc_np



def transform_hybrid(
    df_input: pd.DataFrame,
    turnday_threshold: int = 14,
    d001_threshold: int = 100,
    bidirectional: bool = False,
    safety_days: dict | None = None,
    dimension: str = "nature",
    nd_priority: bool = False,
    progress_callback=None,
    ref_date=None,
    sales_rebalance: bool = False,
    sales_oos_days: int | None = None,
    sales_overstock_days: int | None = None,
    sales_overstock_min_units: int | None = None,
    return_to_d001: bool = False,
):
    """Safety Stock x NetStock Review 核心轉換（調貨事件 Review Safety 設定）。

    - 分類維度（dimension，預設 "nature" 維持現行契約）：
      "nature"＝以 Shop Nature 分類；"rp"＝以 RP Type 分類；"both"＝Shop Nature＋RP Type（6 組）；
      "om"＝以 OM 分類（動態 OM 名單）；"nature_om"＝Shop Nature＋OM（3×N 組）；
      "sales"＝以銷量分類（高銷／中銷／低銷／零銷 四帶；計算同 nature 模式，僅分組顯示不同）；
      "sales_nature"＝以銷量＋Shop Nature 分類（3 Nature×4 帶＝12 組，如 Tourist-高銷；
      計算同 nature 模式，僅分組顯示不同）。
      維度決定需求方過濾與 Summary／交叉表分組，**不改變 TurnDay 池與 Safety 基準天數
      （兩者恆按 Shop Nature）**。RP Type 正規化：RF＝可補貨、ND＝不可補貨；
      其他／空白歸 ND 並列 warning。來源缺「RP Type」欄時全部列按 RF 處理
      （ND 排除規則無法套用）；rp／both 維度退回 nature 並列 warning，
      om／nature_om／sales／sales_nature 維度不退回（OM 分組不依賴 RP Type，
      銷量帶分組為常開產出欄）。
      需求方一律排除 ND 列（ND 不作調入方，仍可作供應方）——七種維度皆適用：
      nature／both／om／nature_om／sales／sales_nature 需求方為 Tourist-RF
      （後四者計算與 nature 完全相同，僅 Summary／交叉表分組不同）；
      rp 需求方為任何 Nature 的 RF 列，
      且取消「需求 Nature 自身除外」守衛與池級歸零（允許同 Nature 店間再平衡）。
      供應方七種維度皆維持 Mix／Non Tourist（勾選雙向時另加 Tourist）。
      OM 分組：OM 欄去除前後空格、空白歸入「空白」群組、大小寫不敏感（Ivy＝ivy，顯示沿用首見原值）、按字串排序；
      來源缺「OM」欄或全欄空白時 om／nature_om 退回 nature 並列 warning。
      非法維度值退回 nature，不拋錯。
    - 調貨段：與 Stock Checking 同規則（SKU 級加總 TurnDay、Nature 供應門檻、
      14 日需求/可轉出、NT 優先→Mix 補足、需求量大→小），另疊加本模式特有的
      **資料品質補正**（Shop Nature 正規化、月銷逐列回補、D001 逐列 Net→Stock
      fallback；回補準則見各段 warning），故同一份輸入在兩個模式可能出現細微差異。
      **調貨段恆以 14 日為基準，不受 safety_days 影響。**
    - Review 段（Safety 基準周轉天數，按 Shop Nature 分別設定）：
      Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 該列 Nature 的基準天數)；
      safety_days 預設 {"Tourist": 7, "Mix": 9, "Non Tourist": 10}（UI 建議範圍 7–10 日）。
      Safety 設定不跟隨存貨周轉（調貨）天數，兩段基準各自獨立。
      Tourist 需調入列：現有 Adj. Safety Stock < 所需量 → Safety 設定不足；
      已標示 SKU 的有貨列（實際轉出或可轉出未使用）：
      現有設定 > 生效基準 → Safety 設定過高；等於邊界一律視為設定正常。
      生效基準（下調補底，SAP MOQ 理據，見 _HYBRID_SAFETY_MIN_KEEP）：所需量 = 0 的
      歸零退場建議不設保底；所需量 ≥ 1 且現有 ≥ 3 件者補底至 max(所需量, 3)
      （補底不高於現值，現有 3 件而所需 1–2 件者改判設定正常）；現有 1–2 件
      維持原所需量，不強制上調到 3。
      建議 Safety 設定值 = 不足列取所需量、過高列取生效基準（僅不足/過高列），
      建議 Safety 設定值 Turnday = 建議值 ÷ (P90D 月均 ÷ 30)（月均 0／非建議列留空）；
      受保底影響（建議值 3 件 > 所需量 ≥ 1 件）時，店舖問題 SKU 的建議方法加註保底原因；
      Safety 建議調整量 = 建議值 − 現有設定（正=增加、負=減少）；
      「Safety 周轉所需量」欄維持原所需量供稽核，保底只體現在建議值與調整量。
    - bidirectional=True（允許雙向調貨，預設 False 維持單向契約）：任何 Nature 缺貨
      （該 Nature 加總 TurnDay < 門檻 且 D001 < 門檻 且列需求 > 0）且其他任一 Nature
      有貨（加總 TurnDay ≥ 門檻）即標示需要調貨；新增 RAW 欄「可轉出貨量 (Tourist)」，
      Tourist 列亦可作轉出來源；來源池順序固定 Non Tourist → Mix → Tourist，
      且「需求 Nature 自身除外」為分配層的顯式守衛（池級歸零另保證缺貨 Nature 不進入來源池）。
      Review 段對稱擴展：需要調入可出現在任何 Nature 列、實際轉出／可轉出未使用
      可出現在 Tourist 列（事件與診斷文字不變）。
    - nd_priority=True（ND 優先轉出，預設 False 維持現行契約）：**同一 Shop Nature 池內**
      RP Type 為 ND 的來源列先於 RF 列取用，池間順序仍為 Non Tourist → Mix → Tourist
      （雙向時才含 Tourist），展開為 NT-ND → NT-RF → Mix-ND → Mix-RF → Tourist-ND →
      Tourist-RF 六段；ND／RF 兩段各自維持來源檔列序（不按可轉出量重排）。
      轉出資格（池級歸零、逐列 14 日超額、自我調貨／同 Nature／同OM 守衛）一律不放寬，
      ND 列仍不作調貨需求方；「ND」的定義＝正規化後 RP Type 非 RF（含無法辨識而歸入 ND 的列）。
      因 donor_mask 與分配順序無關，本開關**只改調貨分配結果**（轉出量／調貨來源／
      事件類型／調貨後庫存與 TurnDay），不影響 Safety Review、建議調整值與調整量。
      來源檔缺 RP Type 欄時自動失效（全部列按 RF）並列 warning。
    - return_to_d001=True（建議退D001，預設 False 維持現行契約）：新增 RAW 欄
      「建議退D001 量」——接收目標為 D001（Warehouse，可無限接收），故不跑分配。
      退貨基準為 Safety Stock (現時)（非調貨段的 14 日線，與池級門檻脫鉤，逐列判定）：
      中銷／低銷／零銷列可退＝max(0, Shop Stock Qty (CM) − Adj. Safety Stock)；
      高銷列原則上不退，僅當 Safety > 0 且庫存 ≥ 3×Safety（高於 Safety 200%）時
      可退＝庫存 − 2×Safety；「需要調貨」列一律不退；建議退D001 量＝可退 − 轉出量
      （不限OM 分配實績，調貨優先），取整、下限 0。另存內部欄
      「_hybrid_return_d001_gross」（可退額度毛額，RAW 顯示不輸出）供退D001 建議
      清單對照。缺「RP Type」欄、OM 限制與分類維度均不影響本欄；本欄只作退倉建議，
      不改變調貨分配、Safety Review 與建議調整量。
    - 銷量分類（常開，與 sales_rebalance 開關無關）：RAW／主表新增「銷量帶」
      （高銷／中銷／低銷／零銷，同 SKU 內按 Sold Qty (P90D) 平均排名分帶，
      見 _hybrid_sales_bands）與「銷量存貨分類」（高銷缺貨／低銷積壓／零銷積壓／正常，
      店舖 TurnDay = Shop Stock Qty (CM) ÷ 生效月銷 × 30.5，門檻見
      _hybrid_sales_params：缺貨線預設 7 日、積壓線預設 30 日、積壓保底預設 2 件）。
    - sales_rebalance=True（銷量再平衡，預設 False 維持現行契約）：需求方在原判定之外
      **增列**「高銷缺貨」列——豁免所屬 Nature 池加總 TurnDay 門檻（本通道存在目的），
      但仍須 RF、補到 14 日需求 > 0、D001 < 門檻、同 SKU 有可轉出量；單向模式需求方
      仍限 Tourist 列、轉出池仍限 Mix／Non Tourist（跟隨方向框架，雙向時兩端放寬至
      任何 Nature）。供應方轉出量公式不變（逐列 14 日超額＋池級歸零），需求方擴大後
      更多 SKU 有需求列，健康池內積壓列經既有 donor 判定自動入事件表；
      自我調貨／同 Nature／同OM 守衛與 ND 不作需求方全部保留。
      新事件列自動流入 Safety Review、店舖問題 SKU、Buyer 統計與簡易調貨明細。
    - 上下文欄位：調貨後庫存 = 原始 + 實際調入 − 轉出；
      調貨後 TurnDay = 調貨後庫存 ÷ 月銷 × 30.5（月銷 0 留空；月銷為回補後值，
      另存「Effective Shop Monthly Sold Qty (P90D)」稽核欄，原值保留於原欄）。
    - OM 變體：分配跑兩遍——不限OM（契約欄位）與同OM（來源店須與需求店同 OM，
      欄位加「(同OM)」後綴）；Safety Review 與建議調整量兩變體相同，不受調貨限制影響。
    - Summary：附 Stock Checking 的 SKU 級彙總（按分類維度分組：nature 三組／rp 兩組／
      both 六組／om 動態 OM 名單／nature_om 三 Nature × OM／sales 四銷量帶／
      sales_nature 三 Nature × 四帶 12 組；庫存/月銷加總、TurnDay、佔比、
      強制調貨提示、存貨偏重提示（門檻按維度切換）），
      供查閱哪些 SKU 存貨偏重哪個分組。
    - Buyer 統計：沿用 Stock Checking 的 Buyer 需調貨 SKU 統計（同 SKU 首個非空 Buyer，
      已刪除重複），另加 Buyer 需調整 Safety SKU 統計（Safety 設定不足＋過高母體，
      含各 Buyer 的調整量合計）。
    - 零銷售轉 ND 提示：Past 90 Days 0 Sales 候選（四期銷售皆 ≤ 0 且有庫存）於
      「店舖問題 SKU」以 ND 退場提示（距參考日期 90 天內新品除外）；參考日期
      ref_date 預設為今天，建議傳入 MTD 截止日期（與報表資料基準一致）。
    回傳 (raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, warnings)；
    stats["safety_days"] 為實際採用的各 Nature 基準天數；
    stats["dimension"] 為實際採用的分類維度（rp／both 缺 RP 欄、om／nature_om 缺 OM 分組時退回 nature）；
    stats["om_groups"] 為排序後不重複的 OM 名單（供 Excel 重建動態欄位用）；
    stats["nd_priority"] 為 ND 優先轉出的請求狀態（缺 RP Type 欄時不生效，nd_transfer 恆為 0）；
    stats["nd_transfer"]／["nd_transfer_om"] 為 RP Type＝ND 來源列的轉出量合計（不限OM／同OM）；
    stats["buyer_need_skus"] 為 [(Buyer, 需調貨 SKU 數)]；
    stats["buyer_safety_fix"] 為 [(Buyer, 需調整 Safety SKU 數, Safety 建議調整量合計)]；
    stats["sales_rebalance"] 為銷量再平衡開關狀態；stats["sales_oos_days"]／
    ["sales_overstock_days"]／["sales_overstock_min_units"] 為實際採用的銷量門檻；
    stats["sales_oos_rows"]／["sales_oos_skus"] 為高銷缺貨列數／SKU 數（全列母體）；
    stats["sales_overstock_rows"]／["sales_overstock_units"] 為積壓列數（低銷＋零銷）／
    積壓列庫存合計；stats["sales_need_rows"] 為僅由銷量再平衡通道新增的需求列數；
    stats["return_to_d001"] 為建議退D001 開關狀態，stats["return_rows"]／["return_skus"]／
    ["return_units"] 為退D001 建議的列數／SKU 數／合計件數（關閉時為 0）。
    """
    missing = validate_hybrid_input(df_input)
    if missing:
        raise ValueError(f"缺少混合模式必要欄位: {', '.join(_display_safety_header(m) for m in missing)}")
    df_input = _canonicalize_hybrid_columns(df_input)
    if progress_callback:
        progress_callback(0.05, "驗證通過")
    warnings: list[str] = []

    df = _clean_dataframe(df_input, shop_natures=None, progress_callback=progress_callback)
    if df.empty:
        raise ValueError("過濾後無有效資料列，請確認來源檔案中的 Buyer 欄位是否包含有效數據（非空白或總計列）")

    df["SKU"] = df["SKU"].apply(_safe_sku)
    # SKU Desc 缺失或全欄空白時亦補正（既有空白欄位會讓摘要 groupby 失效）
    if "SKU Desc" not in df.columns or df["SKU Desc"].fillna("").astype(str).str.strip().eq("").all():
        df = _compute_raw_extra_columns(df)

    if progress_callback:
        progress_callback(0.12, "正規化 Shop Nature...")
    df, mapped_other = _normalize_hybrid_nature(df)
    if mapped_other:
        warnings.append(f"有 {mapped_other} 列 Shop Nature 無法辨識（Other / 空白 / 其他非標準值），已正規化為 Non Tourist。")
    warnings.extend(_transfer_input_consistency_warnings(df))

    # ── 分類維度與 RP Type 正規化（RF＝可補貨、ND＝不可補貨；ND 列不作調貨需求方）──
    # om／nature_om 維度按 OM 分組，不依賴 RP Type，故缺 RP 欄時不退回（仍按 RF 處理）；
    # 僅 rp／both 維度缺 RP 欄時退回 nature。
    dim = _hybrid_dimension(dimension)
    df, mapped_rp, rp_present = _normalize_hybrid_rp_type(df)
    if not rp_present:
        _rp_fallback = dim in ("rp", "both")
        warnings.append(
            "來源檔案缺少「RP Type」欄位，ND 列排除規則無法套用（全部列按可補貨 RF 處理）"
            + ("，分類維度已退回 Shop Nature。" if _rp_fallback else "。")
        )
        if _rp_fallback:
            dim = "nature"
    elif mapped_rp:
        warnings.append(
            f"有 {mapped_rp} 列 RP Type 無法辨識（非 RF／ND 或空白），已正規化為 ND"
            "（不可補貨：不作調貨需求方，仍可作供應方）。"
        )
    is_rf = (
        (df["RP Type"] == "RF").to_numpy()
        if rp_present
        else np.ones(len(df), dtype=bool)
    )

    idx = df.index
    nature = df["Shop Nature"]

    sold_raw = _hybrid_numeric_column(df, "Sold Qty (P90D)", idx, allow_negative=True)
    n_sold_invalid = int(sold_raw.isna().sum())
    n_sold_negative = int((sold_raw < 0).fillna(False).sum())
    if n_sold_invalid:
        warnings.append(f"有 {n_sold_invalid} 列「Sold Qty (P90D)」空白或非數值，已視為 0。")
    if n_sold_negative:
        warnings.append(f"有 {n_sold_negative} 列「Sold Qty (P90D)」為負數，已按 0 計算。")
    sold_p90d = sold_raw.mask(sold_raw < 0, 0.0).fillna(0.0)

    # 庫存量必須為整數（件數）；空白視為 0，小數、負數、非有限值直接拒絕避免守恆失真
    cm_stock_raw = _hybrid_numeric_column(df, "Shop Stock Qty (CM)", idx)
    nonint = cm_stock_raw.notna() & (cm_stock_raw != np.floor(cm_stock_raw))
    if nonint.any():
        raise ValueError("「Shop Stock Qty (CM)」必須為整數件數，發現小數值，請修正來源資料後再執行。")
    cm_stock = cm_stock_raw.fillna(0.0)
    df["Shop Stock Qty (CM)"] = cm_stock

    # 月銷空白、非數值、負數，或月銷為 0 但 P90D 有正數銷售時，逐列以 P90D 換算
    msold_raw = _hybrid_numeric_column(df, "Shop Monthly Sold Qty (P90D)", idx, allow_negative=True)
    n_ms_negative = int((msold_raw < 0).fillna(False).sum())
    if n_ms_negative:
        warnings.append(f"有 {n_ms_negative} 列「Shop Monthly Sold Qty (P90D)」為負數，已改用 P90D 換算。")
    msold_raw = msold_raw.mask(msold_raw < 0)
    msold_blank = msold_raw.isna()
    msold_zero_conflict = msold_raw.eq(0) & sold_p90d.gt(0)
    n_ms_fallback = int((msold_blank | msold_zero_conflict).sum())
    if n_ms_fallback:
        warnings.append(
            f"有 {n_ms_fallback} 列「Shop Monthly Sold Qty (P90D)」空白或為 0 但 P90D 有銷售，"
            "已逐列以 Sold Qty (P90D) ÷ 90 × 30.5 換算月銷。"
        )
    msold_eff = msold_raw.where(~(msold_blank | msold_zero_conflict), sold_p90d / 90.0 * 30.5)
    # 實際採用月銷另存 Effective 欄（與 D001 Effective 欄同一準則）；覆寫前保留原始輸入供稽核比對
    df["Effective Shop Monthly Sold Qty (P90D)"] = msold_eff
    df["Shop Monthly Sold Qty (P90D)"] = msold_eff

    # ── 銷量分類參數（門檻恆計算並寫入分類欄；sales_rebalance 開關只控制調貨通道）──
    _sales_rebalance, _sales_params, _sales_param_warnings = _hybrid_sales_params(
        sales_rebalance,
        sales_oos_days=sales_oos_days,
        sales_overstock_days=sales_overstock_days,
        sales_overstock_min_units=sales_overstock_min_units,
    )
    warnings.extend(_sales_param_warnings)

    # D001：逐列 Net 優先、Stock 回補（兩者皆空白或非有限值視為空白），並保留 Effective 欄位供 RAW 稽核
    avail_now = set(df.columns)
    net_d001_col = _resolve_column_name("D001 Net Stock Qty (CM)", avail_now)
    stock_d001_col = _resolve_column_name("D001 Stock Qty (CM)", avail_now)
    if net_d001_col is None:
        if stock_d001_col is not None:
            warnings.append("來源檔案缺少「D001 Net Stock Qty (CM)」，D001 以「D001 Stock Qty (CM)」取代。")
        else:
            warnings.append("來源檔案缺少「D001 Net Stock Qty (CM)」與「D001 Stock Qty (CM)」，D001 一律視為 0（D001 低於門檻的判定將全數成立）。")
    net_d001_s = (
        _hybrid_numeric_column(df, net_d001_col, idx, allow_negative=True, nonfinite_as_missing=True)
        if net_d001_col else _safe_default_series(np.nan, idx)
    )
    stock_d001_s = (
        _hybrid_numeric_column(df, stock_d001_col, idx, allow_negative=True, nonfinite_as_missing=True)
        if stock_d001_col else _safe_default_series(np.nan, idx)
    )
    n_d001_blank = int(net_d001_s.isna().sum()) if net_d001_col else 0
    if n_d001_blank:
        warnings.append(f"有 {n_d001_blank} 列「D001 Net Stock Qty (CM)」空白或非有限值，已逐列改用「D001 Stock Qty (CM)」回補（兩者皆空白視為 0）。")
    eff_d001 = net_d001_s.where(net_d001_s.notna(), stock_d001_s).fillna(0.0)
    df["D001 Effective Stock Qty (CM)"] = eff_d001.astype(int)

    is_mix = (nature == "Mix").to_numpy()
    is_nt = (nature == "Non Tourist").to_numpy()
    is_tourist = (nature == "Tourist").to_numpy()
    cm_stock_val = cm_stock.to_numpy(dtype=float)
    sold_val = sold_p90d.to_numpy(dtype=float)
    msold_val = msold_eff.to_numpy(dtype=float)
    no_sale_has_stock = (sold_val <= 0) & (cm_stock_val > 0)
    ideal_14d = sold_val / 90.0 * 14
    # Safety 基準天數（按 Nature）— 只用於 Review 段；調貨段一律用上面的 14 日
    _safety_days, _safety_day_warnings = _hybrid_safety_days(safety_days)
    warnings.extend(_safety_day_warnings)
    safety_days_np = np.select(
        [is_tourist, is_mix, is_nt],
        [float(_safety_days["Tourist"]), float(_safety_days["Mix"]), float(_safety_days["Non Tourist"])],
        default=float(_safety_days["Non Tourist"]),
    )

    sku_key = df["SKU"]

    # ── 銷量帶／銷量存貨分類（常開顯示欄；店舖 TurnDay 與調貨 14 日線、Safety 基準天數互相獨立）──
    with np.errstate(divide="ignore", invalid="ignore"):
        _shop_turn_day = np.where(msold_val > 0, cm_stock_val / np.where(msold_val > 0, msold_val, 1.0) * 30.5, np.inf)
    df[_HYBRID_SALES_BAND_COLUMN] = _hybrid_sales_bands(sold_p90d, sku_key)
    df[_HYBRID_SALES_CLASS_COLUMN] = _hybrid_sales_classification(
        df[_HYBRID_SALES_BAND_COLUMN], _shop_turn_day, cm_stock_val, _sales_params,
    )

    def _sku_nature_turnday(nature_mask):
        s_stock = pd.Series(np.where(nature_mask, cm_stock_val, 0.0), index=idx)
        s_msold = pd.Series(np.where(nature_mask, msold_val, 0.0), index=idx)
        agg_stock = s_stock.groupby(sku_key).transform("sum")
        agg_msold = s_msold.groupby(sku_key).transform("sum")
        with np.errstate(divide="ignore", invalid="ignore"):
            return np.where(agg_msold.to_numpy() > 0, agg_stock.to_numpy() / agg_msold.to_numpy() * 30.5, np.inf)

    sku_tourist_turnday = _sku_nature_turnday(is_tourist)
    sku_mix_turnday = _sku_nature_turnday(is_mix)
    sku_nt_turnday = _sku_nature_turnday(is_nt)
    for label, td in (("Tourist", sku_tourist_turnday), ("Mix", sku_mix_turnday), ("Non Tourist", sku_nt_turnday)):
        df[f"{label} 加總 TurnDay"] = np.round(np.where(np.isfinite(td), td, np.nan), 2)

    per_row_excess = np.where(no_sale_has_stock, cm_stock_val, np.maximum(0, cm_stock_val - ideal_14d))
    mix_pool_ok = sku_mix_turnday >= turnday_threshold
    nt_pool_ok = sku_nt_turnday >= turnday_threshold
    if dim == "rp":
        # RP Type 模式：允許同 Nature 店間再平衡——不作池級歸零，可轉出量按逐列 14 日超額計
        mix_excess = np.where(is_mix, per_row_excess, 0)
        nt_excess = np.where(is_nt, per_row_excess, 0)
    else:
        mix_excess = np.where(mix_pool_ok, np.where(is_mix, per_row_excess, 0), 0)
        nt_excess = np.where(nt_pool_ok, np.where(is_nt, per_row_excess, 0), 0)
    df["可轉出貨量 (Mix)"] = mix_excess.astype(int)
    df["可轉出貨量 (Non Tourist)"] = nt_excess.astype(int)

    demand = np.maximum(0, np.ceil(ideal_14d) - cm_stock_val)
    d001_np = eff_d001.to_numpy(dtype=float)
    ttd_np = pd.Series(sku_tourist_turnday, index=idx).to_numpy(dtype=float)
    mtd_np = pd.Series(sku_mix_turnday, index=idx).to_numpy(dtype=float)
    ntd_np = pd.Series(sku_nt_turnday, index=idx).to_numpy(dtype=float)
    # 每列所屬 Nature 的 SKU 級加總 TurnDay（rp 模式需求判定用；仍按 Shop Nature 準則）
    nature_td_np = np.select([is_tourist, is_mix], [ttd_np, mtd_np], default=ntd_np)
    tourist_pool_ok = sku_tourist_turnday >= turnday_threshold
    if dim == "rp":
        # 需求方＝任何 Nature 的 RF 列（ND 一律不作需求）；有貨可調＝同 SKU 任一列有超額
        # （含同 Nature 店間再平衡，故不以他 Nature 池級門檻為前提）；
        # 供應池維持 Mix／Non Tourist（雙向時另加 Tourist）
        if bidirectional:
            tourist_excess = np.where(is_tourist, per_row_excess, 0)
            df["可轉出貨量 (Tourist)"] = tourist_excess.astype(int)
        else:
            tourist_excess = np.zeros(len(df))
        supply_avail = mix_excess + nt_excess + tourist_excess
        sku_supply = pd.Series(supply_avail, index=idx).groupby(sku_key).transform("sum").to_numpy(dtype=float)
        need_mask = (
            is_rf
            & (nature_td_np < turnday_threshold)
            & (d001_np < d001_threshold)
            & (demand > 0)
            & (sku_supply > 0)
        )
    elif bidirectional:
        # 對稱三向：任何 Nature 缺貨（加總 TurnDay < 門檻）且其他任一 Nature 有貨即為需求方
        # （ND 列一律不作需求）；Tourist 亦可作轉出來源（來源池順序固定 NT → Mix → Tourist）
        tourist_excess = np.where(
            tourist_pool_ok,
            np.where(is_tourist, per_row_excess, 0),
            0,
        )
        df["可轉出貨量 (Tourist)"] = tourist_excess.astype(int)
        need_mask = (
            (
                (is_tourist & (ttd_np < turnday_threshold) & (mix_pool_ok | nt_pool_ok))
                | (is_mix & (mtd_np < turnday_threshold) & (nt_pool_ok | tourist_pool_ok))
                | (is_nt & (ntd_np < turnday_threshold) & (mix_pool_ok | tourist_pool_ok))
            )
            & is_rf
            & (d001_np < d001_threshold)
            & (demand > 0)
        )
    else:
        tourist_excess = np.zeros(len(df))
        # 需求方＝Tourist-RF（ND 列一律不作需求）；both 維度計算與本分支完全相同，僅分組顯示不同
        need_mask = (
            is_tourist
            & is_rf
            & (ttd_np < turnday_threshold)
            & (d001_np < d001_threshold)
            & (demand > 0)
            & (mix_pool_ok | nt_pool_ok)
        )
    # ── 銷量再平衡通道（僅 sales_rebalance=True）：高銷缺貨列增列為需求方 ──
    # 豁免所屬 Nature 池加總 TurnDay 門檻（本通道存在目的），但仍須 RF、
    # 補到 14 日需求 > 0、D001 < 門檻、同 SKU 有可轉出量（供應欄公式與池級歸零不變）；
    # 單向模式需求方仍限 Tourist 列（跟隨方向框架；雙向／rp 維度任何 Nature）。
    # 需求方擴大後更多 SKU 進入 flagged_skus，健康池內積壓列經既有 donor 判定
    # 自動入事件表；自我調貨／同 Nature／同OM 守衛與 ND 不作需求方全部保留。
    need_mask_base = need_mask
    if _sales_rebalance:
        _sales_sku_supply = pd.Series(
            mix_excess + nt_excess + tourist_excess, index=idx
        ).groupby(sku_key).transform("sum").to_numpy(dtype=float)
        _high_short = (df[_HYBRID_SALES_CLASS_COLUMN] == _HYBRID_SALES_CLASS_ORDER[0]).to_numpy()
        _sales_dir_allow = (
            np.ones(len(df), dtype=bool) if (bidirectional or dim == "rp") else is_tourist
        )
        _sales_need = (
            _sales_dir_allow
            & is_rf
            & _high_short
            & (demand > 0)
            & (d001_np < d001_threshold)
            & (_sales_sku_supply > 0)
        )
        need_mask = need_mask | _sales_need

    need_series = pd.Series(need_mask, index=idx)
    _sales_added_rows = int((need_mask & ~need_mask_base).sum()) if _sales_rebalance else 0
    if _sales_added_rows:
        warnings.append(
            f"⚖️ 銷量再平衡：{int(_sales_added_rows)} 列高銷缺貨店舖列新增為調貨需求方"
            "（豁免 Nature 池加總 TurnDay 門檻；D001 門檻、有貨可調與 RF 資格要求不變）。"
        )

    df["需要調貨"] = np.where(need_mask, "是", "")
    df["調入需求量 (補到14日)"] = np.where(need_mask, demand, 0).astype(int)

    site_arr = df.get("Site Code", _safe_default_series("", idx)).fillna("").astype(str).str.strip().to_numpy(dtype=object)
    sku_arr = df["SKU"].to_numpy(dtype=object)
    n_rows = len(df)

    # ── OM 群組（同OM調貨變體與 om／nature_om 分類維度用）：
    # 去除前後空格；空白自成「空白」群組；大小寫不敏感（顯示沿用首見原值）──
    om_col = _resolve_column_name("OM", set(df.columns))
    if om_col is None:
        warnings.append("來源檔案缺少「OM」欄位，「同OM調貨」無法區分店群，全部店舖視為同一組（結果與不限OM相同）。")
        om_src = _safe_default_series("", idx)
        n_om_blank = n_rows
    else:
        om_src = df[om_col]
        n_om_blank = int((om_src.fillna("").astype(str).str.strip() == "").sum())
        if 0 < n_om_blank < n_rows:
            warnings.append(
                f"有 {n_om_blank} 列「OM」空白，已歸入「空白」群組"
                "（同OM調貨與 OM 分類均視為同一 OM，請檢查該等店舖是否已安排 OM）。"
            )
    om_disp, om_groups = _normalize_hybrid_om(om_src)
    df["OM"] = om_disp
    # 同OM調貨比對鍵（大小寫不敏感；顯示值另存 df["OM"] 供 RAW／主表稽核）
    om_arr = om_disp.str.lower().to_numpy(dtype=object)
    # 退回條件按來源是否全空判斷（不以顯示名比對，避免真實 OM 名與「空白」碰撞）
    if dim in ("om", "nature_om") and (om_col is None or n_om_blank == n_rows):
        if om_col is None:
            warnings.append("來源檔案缺少「OM」欄位，無法按 OM 分組，分類維度已退回 Shop Nature。")
        else:
            warnings.append("「OM」欄全欄空白，無法按 OM 分組，分類維度已退回 Shop Nature。")
        dim = "nature"

    # ── 調貨分配：不限OM（現行契約欄位）＋ 同OM（來源店與需求店同 OM 才可轉出）──
    # rp 維度取消「需求 Nature 自身除外」守衛（允許同 Nature 店間再平衡）；
    # nature／both 維度維持守衛（池級歸零＋分配層守衛雙重保證）
    # ND 優先轉出：池內 ND 列前置（ND 的定義＝正規化後非 RF，含無法辨識列）；
    # 缺 RP Type 欄時全部列按 RF，開關自動失效
    _src_nd = (~is_rf) if (nd_priority and rp_present) else None
    _is_nd_src = (df["RP Type"] == "ND") if rp_present else pd.Series(False, index=idx)
    if nd_priority and not rp_present:
        warnings.append(
            "來源檔案缺少「RP Type」欄位（全部列按可補貨 RF 處理），「ND 優先轉出」無法套用，已按現行列序分配。"
        )
    nt_avail0 = df["可轉出貨量 (Non Tourist)"].to_numpy(dtype=float)
    mix_avail0 = df["可轉出貨量 (Mix)"].to_numpy(dtype=float)
    tourist_avail0 = df["可轉出貨量 (Tourist)"].to_numpy(dtype=float) if bidirectional else None
    _guard_nature = None if dim == "rp" else nature.to_numpy()
    actual_np, gap_np, donate_np, desc_np = _hybrid_allocate_transfers(
        n_rows, sku_arr, site_arr, need_mask, demand, nt_avail0, mix_avail0,
        tourist_avail0=tourist_avail0, nature_arr=_guard_nature, nd_priority=_src_nd,
    )
    actual_om, gap_om, donate_om, desc_om = _hybrid_allocate_transfers(
        n_rows, sku_arr, site_arr, need_mask, demand, nt_avail0, mix_avail0, om_arr=om_arr,
        tourist_avail0=tourist_avail0, nature_arr=_guard_nature, nd_priority=_src_nd,
    )
    df["實際調入量"] = actual_np
    df["調入缺口"] = gap_np
    df["轉出量"] = donate_np
    df["調貨來源 (Site Code)"] = desc_np
    df["實際調入量 (同OM)"] = actual_om
    df["調入缺口 (同OM)"] = gap_om
    df["轉出量 (同OM)"] = donate_om
    df["調貨來源 (同OM)"] = desc_om

    if not need_mask.any():
        if dim == "rp":
            warnings.append("沒有符合調貨條件的 RF 列（任何 Nature 加總 TurnDay < 門檻 且 D001 < 門檻 且補到 14 日周轉的需求 > 0 且有貨可調；ND 列不作需求方）。")
        elif bidirectional:
            warnings.append("沒有符合調貨條件的 SKU（任何 Nature 的 RF 列加總 TurnDay < 門檻 且 D001 < 門檻 且補到 14 日周轉的需求 > 0 且有貨可調）。")
        else:
            warnings.append("沒有符合調貨條件的 Tourist RF SKU（加總 TurnDay < 門檻 且 D001 < 門檻 且補到 14 日周轉的需求 > 0 且有貨可調；ND 列不作需求方）。")

    if progress_callback:
        progress_callback(0.45, "Review Safety 設定...")
    raw = df

    # ── 調貨事件 Review（Safety 基準天數按 Nature；Review 與建議調整不受 OM 限制影響）──
    safety_line = np.maximum(0, np.ceil(np.maximum(0, sold_val / 90.0 * safety_days_np)))
    adj_s = _hybrid_numeric_column(raw, "Adj. Safety Stock", idx)
    n_adj_blank = int(adj_s.isna().sum())
    if n_adj_blank:
        warnings.append(f"有 {n_adj_blank} 列「{_display_safety_header('Adj. Safety Stock')}」空白或非數值，Review 時已視為 0。")
    adj_val = adj_s.fillna(0.0).to_numpy(dtype=float)

    flagged_skus = set(sku_arr[need_mask])
    transferable_np = mix_excess + nt_excess + (tourist_excess if bidirectional else 0)
    donor_pool = np.isin(sku_arr, list(flagged_skus)) if flagged_skus else np.zeros(n_rows, dtype=bool)
    if bidirectional:
        # 雙向：轉出來源可以是任何 Nature（含 Tourist），需求列本身不列為來源
        donor_mask = (~need_mask) & donor_pool & ((donate_np > 0) | (transferable_np > 0))
    else:
        donor_mask = (~is_tourist) & donor_pool & ((donate_np > 0) | (transferable_np > 0))
    donated_mask = donor_mask & (donate_np > 0)
    donor_unused_mask = donor_mask & ~donated_mask

    # 下調建議補底（SAP MOQ：Safety 1–2 件難以補貨）—「設定過高」列的生效基準
    # eff_line：所需量 = 0（P90D 零銷售，歸零屬退場）不設保底；
    # 所需量 ≥ 1 且現有 ≥ 3 件 → max(所需量, 3)（補底不得高於現值，
    # 現有 3 件而所需 1–2 件者補底後相等 → 改判設定正常，不出現 0 調整量假動作）；
    # 現有 ≤ 2 件 → 維持原所需量（本身 1–2 件不強制上調到 3）。
    # 「設定不足」列屬 TurnDay 驅動上調，一律照原所需量，不受保底影響。
    eff_line = np.where(
        safety_line <= 0,
        safety_line,
        np.where(
            adj_val >= _HYBRID_SAFETY_MIN_KEEP,
            np.maximum(safety_line, _HYBRID_SAFETY_MIN_KEEP),
            safety_line,
        ),
    )
    under_mask = need_mask & (adj_val < safety_line)
    over_mask = donor_mask & (adj_val > eff_line)
    in_table = need_mask | donor_mask

    event = np.full(n_rows, "", dtype=object)
    event[need_mask] = _HYBRID_EVENT_NEED_IN
    event[donated_mask] = _HYBRID_EVENT_DONATED
    event[donor_unused_mask] = _HYBRID_EVENT_DONOR

    diag = np.full(n_rows, "", dtype=object)
    diag[need_mask] = np.where(under_mask[need_mask], _HYBRID_DIAG_UNDER, _HYBRID_DIAG_NORMAL_NEED)
    diag[donor_mask] = np.where(over_mask[donor_mask], _HYBRID_DIAG_OVER, _HYBRID_DIAG_NORMAL_DONOR)

    line_col = np.full(n_rows, np.nan)
    sugg = np.full(n_rows, np.nan)
    chg = np.full(n_rows, np.nan)
    line_col[in_table] = safety_line[in_table]
    sugg[under_mask] = safety_line[under_mask]
    sugg[over_mask] = eff_line[over_mask]
    fix_mask = under_mask | over_mask
    chg[fix_mask] = sugg[fix_mask] - adj_val[fix_mask]

    stock_after = cm_stock_val + actual_np - donate_np
    td_after = np.full(n_rows, np.nan)
    ms_pos = msold_val > 0
    td_after[ms_pos] = stock_after[ms_pos] / msold_val[ms_pos] * 30.5
    stock_after_om = cm_stock_val + actual_om - donate_om
    td_after_om = np.full(n_rows, np.nan)
    td_after_om[ms_pos] = stock_after_om[ms_pos] / msold_val[ms_pos] * 30.5

    raw[_HYBRID_SAFETY_DAYS_COLUMN] = safety_days_np.astype(int)
    raw[_HYBRID_SAFETY_LINE_COLUMN] = line_col
    raw["事件類型"] = event
    raw["Safety Review"] = diag
    raw["建議 Safety 設定值"] = sugg
    # 建議值周轉日數（與來源「Turnday (P90D Qty)」、SSR「Recommend Safety Stock Turnday」
    # 同一準則）：建議值 ÷ (P90D 月均 ÷ 30)；月均 0 或非建議列留空。
    _sugg_td = np.full(n_rows, np.nan)
    _sugg_td[ms_pos] = sugg[ms_pos] / (msold_val[ms_pos] / 30.0)
    raw[_HYBRID_SUGGEST_TURNDAY_COLUMN] = np.round(_sugg_td, 1)
    raw["Safety 建議調整量"] = np.round(chg, 1)
    # 受下調補底 3 件影響（生效基準 = max(所需量, 3) > 所需量）— 內部旗標，供店舖問題
    # SKU 建議方法加註；於計算層一次判定，避免顯示層反推欄位而與補底規則走樣。
    raw["_hybrid_safety_floor3"] = over_mask & (eff_line > safety_line)
    raw["調貨後庫存"] = np.where(in_table, stock_after, np.nan)
    raw["調貨後 TurnDay"] = np.where(in_table, np.round(td_after, 1), np.nan)
    raw["調貨後庫存 (同OM)"] = np.where(in_table, stock_after_om, np.nan)
    raw["調貨後 TurnDay (同OM)"] = np.where(in_table, np.round(td_after_om, 1), np.nan)

    # ── 全列預防性 Safety 設定建議（不依賴調貨事件；母體為全部列）──
    # 與事件版同口徑（所需量＋下調補底 3 件），惟不設調貨事件門檻，協助用家一次過
    # 設定適合的 Safety Stock，減少日後調貨次數。ND 列不作上調（不可補貨），只可下調。
    _is_nd_full = (raw["RP Type"].to_numpy() == "ND") if "RP Type" in raw.columns else np.zeros(n_rows, dtype=bool)
    full_under = (adj_val < safety_line) & (~_is_nd_full)
    full_over = adj_val > eff_line
    full_review = np.full(n_rows, _HYBRID_FULL_SAFETY_NORMAL, dtype=object)
    full_review[full_under] = _HYBRID_DIAG_UNDER
    full_review[full_over] = _HYBRID_DIAG_OVER
    full_sugg = np.full(n_rows, np.nan)
    full_sugg[full_under] = safety_line[full_under]
    full_sugg[full_over] = eff_line[full_over]
    full_fix = full_under | full_over
    full_chg = np.full(n_rows, np.nan)
    full_chg[full_fix] = full_sugg[full_fix] - adj_val[full_fix]
    _full_sugg_td = np.full(n_rows, np.nan)
    _full_sugg_td[ms_pos] = full_sugg[ms_pos] / (msold_val[ms_pos] / 30.0)
    raw[_HYBRID_FULL_SAFETY_LINE_COLUMN] = safety_line.astype(int)
    raw[_HYBRID_FULL_SAFETY_REVIEW_COLUMN] = full_review
    raw[_HYBRID_FULL_SAFETY_SUGGEST_COLUMN] = full_sugg
    raw[_HYBRID_FULL_SAFETY_TURNDAY_COLUMN] = np.round(_full_sugg_td, 1)
    raw[_HYBRID_FULL_SAFETY_CHG_COLUMN] = np.round(full_chg, 1)

    # ── 建議退D001（選開）：接收目標 D001（Warehouse）可無限接收，故不跑分配。
    # 退貨基準為 Safety Stock (現時)（非調貨段的 14 日線，與池級門檻脫鉤，逐列判定）：
    # 中銷／低銷／零銷列退回超出 Safety 的庫存；高銷列原則上不退，僅當 Safety > 0 且
    # 庫存 ≥ 3×Safety（高於 Safety 200%）時退至 2×Safety；調貨優先——先扣「轉出量」
    # （不限OM 分配實績），「需要調貨」列一律不退；缺 RP Type 欄、OM 與分類維度均不影響。
    _return_units = 0
    if return_to_d001:
        _band_np = df[_HYBRID_SALES_BAND_COLUMN].to_numpy(dtype=object)
        _adj_ret = np.maximum(adj_val, 0.0)
        _is_high_band = _band_np == _HYBRID_SALES_BANDS[0]
        _high_gate = _is_high_band & (_adj_ret > 0) & (cm_stock_val >= 3.0 * _adj_ret)
        # 高銷未達 200% 門檻 → 保留線設為庫存本身（可退歸 0）
        _ret_keep = np.where(_is_high_band & ~_high_gate, cm_stock_val,
                             np.where(_is_high_band, 2.0 * _adj_ret, _adj_ret))
        _ret_gross = np.where(need_mask, 0.0, np.maximum(0.0, cm_stock_val - _ret_keep))
        _ret_net = np.where(need_mask, 0.0, np.maximum(0.0, np.floor(_ret_gross - donate_np)))
        df[_HYBRID_RETURN_D001_GROSS_INTERNAL] = np.floor(_ret_gross).astype(int)
        df[_HYBRID_RETURN_D001_COLUMN] = _ret_net.astype(int)
        _return_units = int(_ret_net.sum())
        if _return_units == 0:
            warnings.append(
                "「建議退D001」已開啟：沒有可退回 D001 的店舖×SKU 列"
                "（僅中銷／低銷／零銷列建議退回超出 Safety Stock (現時) 的庫存；"
                "高銷列須庫存高於 Safety 200% 才退至 2 倍 Safety；需要調貨列一律不退）。"
            )

    if progress_callback:
        progress_callback(0.55, "建立 Review 表...")
    _ref_date = _normalize_ref_date(ref_date)
    _zero_mask, _zero_warnings = _hybrid_p90d_zero_sales_mask(raw, ref_date=_ref_date)
    warnings.extend(_zero_warnings)
    if bool(_zero_mask.any()):
        warnings.append(
            f"偵測到 {int(_zero_mask.sum())} 列符合 Past 90 Days 0 Sales"
            "（M-2／M-1／CM-MTD／P90D 皆零銷售且有庫存），已於「店舖問題 SKU」標示 ND 退場建議（新品已排除）。"
        )
    diagnosis = _build_hybrid_diagnosis(raw)
    diagnosis_om = _build_hybrid_diagnosis(raw, variant="om")
    diag_summary = _build_hybrid_diag_summary(raw, dimension=dim, om_groups=om_groups)
    sc_summary = _build_stock_checking_summary(raw, bidirectional=bidirectional, dimension=dim, om_groups=om_groups)

    stats = {
        "safety_days": dict(_safety_days),
        "dimension": dim,
        "om_groups": list(om_groups),
        "ref_date": _ref_date.isoformat(),
        "turnday_threshold": int(turnday_threshold),
        "d001_threshold": int(d001_threshold),
        "need_rows": int(need_series.sum()),
        "need_skus": int(raw.loc[need_series, "SKU"].nunique()) if need_series.any() else 0,
        "total_demand": int(raw.loc[need_series, "調入需求量 (補到14日)"].sum()) if need_series.any() else 0,
        "total_actual": int(raw["實際調入量"].sum()),
        "total_gap": int(raw["調入缺口"].sum()),
        "nt_transfer": int(raw.loc[raw["Shop Nature"] == "Non Tourist", "轉出量"].sum()),
        "mix_transfer": int(raw.loc[raw["Shop Nature"] == "Mix", "轉出量"].sum()),
        "tourist_transfer": int(raw.loc[raw["Shop Nature"] == "Tourist", "轉出量"].sum()) if bidirectional else 0,
        "total_actual_om": int(actual_om.sum()),
        "total_gap_om": int(gap_om.sum()),
        "nt_transfer_om": int(raw.loc[raw["Shop Nature"] == "Non Tourist", "轉出量 (同OM)"].sum()),
        "mix_transfer_om": int(raw.loc[raw["Shop Nature"] == "Mix", "轉出量 (同OM)"].sum()),
        "tourist_transfer_om": int(raw.loc[raw["Shop Nature"] == "Tourist", "轉出量 (同OM)"].sum()) if bidirectional else 0,
        "bidirectional": bool(bidirectional),
        # 記錄 UI 請求狀態（缺 RP Type 欄時 nd_transfer 恆為 0，指標卡自然不顯示，
        # 不因正規化而與參數快照比對產生不一致的過期警告）
        "nd_priority": bool(nd_priority),
        "nd_transfer": int(raw.loc[_is_nd_src, "轉出量"].sum()) if rp_present else 0,
        "nd_transfer_om": int(raw.loc[_is_nd_src, "轉出量 (同OM)"].sum()) if rp_present else 0,
        "donated_rows_om": int((donor_mask & (donate_om > 0)).sum()),
        "mapped_other_rows": int(mapped_other),
        "mapped_rp_rows": int(mapped_rp),
        "donor_rows": int(donor_mask.sum()),
        "donated_rows": int(donated_mask.sum()),
        "donor_unused_rows": int(donor_unused_mask.sum()),
        "under_rows": int(under_mask.sum()),
        "under_skus": int(raw.loc[under_mask, "SKU"].nunique()) if under_mask.any() else 0,
        "under_inc_total": int(round(float(np.nansum(chg[under_mask])))) if under_mask.any() else 0,
        "over_rows": int(over_mask.sum()),
        "over_skus": int(raw.loc[over_mask, "SKU"].nunique()) if over_mask.any() else 0,
        "over_dec_total": int(round(float(np.nansum(chg[over_mask])))) if over_mask.any() else 0,
        "normal_rows": int(in_table.sum()) - int(under_mask.sum()) - int(over_mask.sum()),
        "total_raw_rows": int(n_rows),
        "full_under_rows": int(full_under.sum()),
        "full_under_skus": int(raw.loc[full_under, "SKU"].nunique()) if full_under.any() else 0,
        "full_under_inc_total": int(round(float(np.nansum(full_chg[full_under])))) if full_under.any() else 0,
        "full_over_rows": int(full_over.sum()),
        "full_over_skus": int(raw.loc[full_over, "SKU"].nunique()) if full_over.any() else 0,
        "full_over_dec_total": int(round(float(np.nansum(full_chg[full_over])))) if full_over.any() else 0,
        "full_normal_rows": int(n_rows) - int(full_under.sum()) - int(full_over.sum()),
        "full_action_rows": int(full_fix.sum()),
        "full_action_skus": int(raw.loc[full_fix, "SKU"].nunique()) if full_fix.any() else 0,
        "sales_rebalance": bool(_sales_rebalance),
        "sales_oos_days": int(_sales_params["oos_days"]),
        "sales_overstock_days": int(_sales_params["overstock_days"]),
        "sales_overstock_min_units": int(_sales_params["overstock_min_units"]),
        "sales_oos_rows": int((raw[_HYBRID_SALES_CLASS_COLUMN] == _HYBRID_SALES_CLASS_ORDER[0]).sum()),
        "sales_oos_skus": int(
            raw.loc[raw[_HYBRID_SALES_CLASS_COLUMN] == _HYBRID_SALES_CLASS_ORDER[0], "SKU"].nunique()
        ),
        "sales_overstock_rows": int(
            raw[_HYBRID_SALES_CLASS_COLUMN].isin(_HYBRID_SALES_CLASS_ORDER[1:3]).sum()
        ),
        "sales_overstock_units": int(
            raw.loc[raw[_HYBRID_SALES_CLASS_COLUMN].isin(_HYBRID_SALES_CLASS_ORDER[1:3]), "Shop Stock Qty (CM)"].sum()
        ),
        "sales_need_rows": int(_sales_added_rows),
        "return_to_d001": bool(return_to_d001),
        "return_rows": int((df[_HYBRID_RETURN_D001_COLUMN] > 0).sum()) if return_to_d001 else 0,
        "return_skus": int(df.loc[df[_HYBRID_RETURN_D001_COLUMN] > 0, "SKU"].nunique()) if (return_to_d001 and (df[_HYBRID_RETURN_D001_COLUMN] > 0).any()) else 0,
        "return_units": int(_return_units),
        "buyer_need_skus": _hybrid_buyer_need_stats(raw, need_series),
        "buyer_safety_fix": _hybrid_buyer_safety_stats(raw),
    }

    if progress_callback:
        progress_callback(0.7, "轉換完成")

    return raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, warnings



def _write_hybrid_diagnosis_sheet(ws, diagnosis: pd.DataFrame, progress_callback=None,
                                  progress_start: float = 0.42, progress_end: float = 0.72):
    """調貨 Review Safety Stock 主表工作表：店舖×SKU 逐列（僅含事件列）。
    progress 區間必須落在對應 generate_excel_* 的下一階段區間內，避免進度倒退。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    band_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    _text_keys = ["Brand", "SKU Desc", "Buyer", "Adj. SST", "Site Code", "事件類型", "Safety Review", "調貨來源 (Site Code)"]
    _decimal_keys = {
        "調貨後 TurnDay": "0.0",
        _HYBRID_SUGGEST_TURNDAY_COLUMN: "0.0",
        "Safety 建議調整量": "#,##0.##",
    }

    headers = list(diagnosis.columns)
    header_labels = [_display_safety_header(h) for h in headers]
    for col_idx, h in enumerate(header_labels, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[1].height = 30

    if diagnosis is None or diagnosis.empty:
        ws.cell(row=2, column=1, value="無符合資料").font = Font(size=12)
        _auto_width(ws)
        return

    total = len(diagnosis)
    for r_idx, (_, row) in enumerate(diagnosis.iterrows()):
        excel_row = r_idx + 2
        is_band = (r_idx % 2 == 1)
        for col_idx, h in enumerate(headers, 1):
            val = row.get(h)
            cell = ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)
            cell.border = thin_border
            if h in _text_keys:
                cell.alignment = left_wrap
            else:
                cell.alignment = center
            if isinstance(val, (int, float)) and not isinstance(val, bool) and h not in _text_keys:
                cell.number_format = _decimal_keys.get(h, "#,##0")
            if is_band:
                cell.fill = band_fill
        if progress_callback and (r_idx % 200 == 0 or r_idx == total - 1):
            p = progress_start + (progress_end - progress_start) * ((r_idx + 1) / max(total, 1))
            progress_callback(p, f"寫入調貨 Review 主表 ({r_idx+1}/{total})...")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{1 + total}"
    _auto_width(ws)



def _write_hybrid_transfer_detail_sheet(ws, detail: pd.DataFrame, progress_callback=None,
                                        progress_start: float = 0.90, progress_end: float = 0.94):
    """簡易調貨明細工作表：一列一筆來源店→接收店調貨（來源字串已拆行）。
    progress 區間必須落在 generate_excel_hybrid 的下一階段區間內，避免進度倒退。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    band_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    _text_keys = ["Brand", "SKU Desc",
                  "Target Site Code(接收店舖)", "Target Shop Name",
                  "Supply Site Code(調貨來源)", "Supply Shop Name"]

    headers = list(_HYBRID_TRANSFER_DETAIL_COLUMNS)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[1].height = 30

    if detail is None or detail.empty:
        ws.cell(row=2, column=1, value="無符合資料").font = Font(size=12)
        _auto_width(ws)
        return

    total = len(detail)
    for r_idx, (_, row) in enumerate(detail.iterrows()):
        excel_row = r_idx + 2
        is_band = (r_idx % 2 == 1)
        for col_idx, h in enumerate(headers, 1):
            val = row.get(h)
            cell = ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)
            cell.border = thin_border
            cell.alignment = left_wrap if h in _text_keys else center
            if isinstance(val, (int, float)) and not isinstance(val, bool) and h not in _text_keys:
                cell.number_format = "#,##0"
            if is_band:
                cell.fill = band_fill
        if progress_callback and (r_idx % 200 == 0 or r_idx == total - 1):
            p = progress_start + (progress_end - progress_start) * ((r_idx + 1) / max(total, 1))
            progress_callback(p, f"寫入簡易調貨明細 ({r_idx+1}/{total})...")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{1 + total}"
    _auto_width(ws)



def _write_hybrid_full_safety_sheet(ws, detail: pd.DataFrame, progress_callback=None,
                                    progress_start: float = 0.75, progress_end: float = 0.79):
    """全列預防性設定建議工作表：店舖×SKU 逐列，只收錄需要調整的列。
    progress 區間必須落在 generate_excel_hybrid 的下一階段區間內，避免進度倒退。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    band_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    _text_keys = ["Brand", "SKU Desc", "Buyer", "Shop Name",
                  _HYBRID_FULL_SAFETY_REVIEW_COLUMN, "事件類型", "Safety Review"]
    _decimal_keys = {
        _HYBRID_FULL_SAFETY_TURNDAY_COLUMN: "0.0",
        _HYBRID_FULL_SAFETY_CHG_COLUMN: "#,##0.##",
    }

    headers = list(_HYBRID_FULL_SAFETY_DETAIL_COLUMNS)
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[1].height = 30

    if detail is None or detail.empty:
        ws.cell(row=2, column=1, value="無符合資料").font = Font(size=12)
        _auto_width(ws)
        return

    total = len(detail)
    for r_idx, (_, row) in enumerate(detail.iterrows()):
        excel_row = r_idx + 2
        is_band = (r_idx % 2 == 1)
        for col_idx, h in enumerate(headers, 1):
            val = row.get(h)
            cell = ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)
            cell.border = thin_border
            cell.alignment = left_wrap if h in _text_keys else center
            if isinstance(val, (int, float)) and not isinstance(val, bool) and h not in _text_keys:
                cell.number_format = _decimal_keys.get(h, "#,##0")
            if is_band:
                cell.fill = band_fill
        if progress_callback and (r_idx % 200 == 0 or r_idx == total - 1):
            p = progress_start + (progress_end - progress_start) * ((r_idx + 1) / max(total, 1))
            progress_callback(p, f"寫入 Safety 設定建議 (全列) ({r_idx+1}/{total})...")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{1 + total}"
    _auto_width(ws)



def _write_hybrid_summary_block(ws, start_row, section_title, block, *, header_fill,
                                header_font, bold_font, section_fill, total_fill,
                                thin_border, center, left_wrap):
    """「Review 摘要」工作表的統計區塊：合併標題列 → 表頭列 → 資料列
    （首欄值為「總計」者加粗體與灰底）。樣式與 Safety Stock Review 的 SUMMARY
    （_write_summary_sheet，transformer.py:971）一致。回傳下一個區塊的標題列號。"""
    headers = list(block.columns)
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=len(headers))
    sc = ws.cell(row=start_row, column=1, value=section_title)
    sc.font = Font(bold=True, size=12, color="1E40AF")
    sc.fill = section_fill
    sc.alignment = Alignment(horizontal="left", vertical="center")

    header_row = start_row + 1
    if block.empty:
        ws.cell(row=header_row, column=1, value="無符合資料").font = Font(size=12)
        return header_row + 2

    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border

    for offset, (_, row) in enumerate(block.iterrows(), 1):
        is_total = str(row.get(headers[0])) == "總計"
        for col_idx, h in enumerate(headers, 1):
            val = row.get(h)
            cell = ws.cell(row=header_row + offset, column=col_idx,
                           value=None if (isinstance(val, float) and pd.isna(val)) else val)
            cell.border = thin_border
            cell.alignment = left_wrap if col_idx == 1 else center
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                cell.number_format = "#,##0"
            if is_total:
                cell.font = bold_font
                cell.fill = total_fill
    return header_row + len(block) + 2



def _write_hybrid_buyer_combined_table(ws, combined, thin_border, header_fill,
                                        header_font, bold_font, section_fill,
                                        total_fill, center, start_row: int) -> int:
    """Review 摘要的 Buyer 合併統計表（Buyer／需調貨 SKU 數／需調貨 SKU 佔比／
    調入需求量合計／需調整 Safety SKU 數／需調整 Safety SKU 佔比／
    Safety 建議調整量合計＋總計列）。

    combined 為 [(Buyer, 需調貨 SKU 數, 調入需求量合計, 需調整 Safety SKU 數,
    Safety 建議調整量合計)]；舊 4 元組（無貨量）亦相容，貨量以 0 補齊。
    缺席值以 0 補齊；總計列為各欄合計（貨量與調整量為淨額；佔比總計為 100%）。
    空集合時不寫入，回傳原 start_row。回傳下一個區塊的標題列號。
    """
    if not combined:
        return start_row
    norm: list[tuple[str, int, int, int, int]] = []
    for row in combined:
        if len(row) == 5:
            b, n, d, f, v = row
            norm.append((str(b), int(n), int(d), int(f), int(v)))
        else:
            b, n, f, v = row  # type: ignore[misc]
            norm.append((str(b), int(n), 0, int(f), int(v)))
    combined = norm
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
    tc = ws.cell(row=start_row, column=1, value=_HYBRID_BUYER_COMBINED_TITLE)
    tc.font = Font(bold=True, size=12, color="1E40AF")
    tc.fill = section_fill
    tc.alignment = Alignment(horizontal="left", vertical="center")
    hr = start_row + 1
    for col_idx, h in enumerate(
        ("Buyer", "需調貨 SKU 數", "需調貨 SKU 佔比", "調入需求量合計",
         "需調整 Safety SKU 數", "需調整 Safety SKU 佔比", "Safety 建議調整量合計"), 1):
        cell = ws.cell(row=hr, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    total_need = sum(n for _, n, _, _, _ in combined)
    total_demand = sum(d for _, _, d, _, _ in combined)
    total_fix = sum(f for _, _, _, f, _ in combined)
    total_chg = sum(v for _, _, _, _, v in combined)
    r = hr + 1
    for buyer, need_cnt, demand_total, fix_cnt, chg_total in combined:
        c1 = ws.cell(row=r, column=1, value=buyer)
        c1.border = thin_border
        c2 = ws.cell(row=r, column=2, value=need_cnt)
        c2.number_format = "#,##0"
        c2.alignment = center
        c2.border = thin_border
        c3 = ws.cell(row=r, column=3, value=need_cnt / total_need if total_need else None)
        c3.number_format = "0.0%"
        c3.alignment = center
        c3.border = thin_border
        c4 = ws.cell(row=r, column=4, value=demand_total)
        c4.number_format = "#,##0"
        c4.alignment = center
        c4.border = thin_border
        c5 = ws.cell(row=r, column=5, value=fix_cnt)
        c5.number_format = "#,##0"
        c5.alignment = center
        c5.border = thin_border
        c6 = ws.cell(row=r, column=6, value=fix_cnt / total_fix if total_fix else None)
        c6.number_format = "0.0%"
        c6.alignment = center
        c6.border = thin_border
        c7 = ws.cell(row=r, column=7, value=chg_total)
        c7.number_format = "#,##0"
        c7.alignment = center
        c7.border = thin_border
        r += 1
    for col_idx, v in ((1, "總計"), (2, total_need), (3, 1.0 if total_need else None),
                       (4, total_demand), (5, total_fix), (6, 1.0 if total_fix else None),
                       (7, total_chg)):
        cell = ws.cell(row=r, column=col_idx, value=v)
        cell.font = bold_font
        cell.fill = total_fill
        cell.border = thin_border
        cell.alignment = center
        if col_idx in (2, 4, 5, 7):
            cell.number_format = "#,##0"
        elif col_idx in (3, 6):
            cell.number_format = "0.0%"
    return r + 2



def _write_hybrid_diag_summary_sheet(ws, cross: pd.DataFrame, stats: dict | None,
                                     raw: pd.DataFrame, thin_border, desc_font,
                                     progress_callback=None):
    """Review 摘要工作表：指標卡 4–6 排（第 4 排為同OM調貨對照；雙向調貨時另加第 5 排
    Tourist 轉出量；銷量再平衡開啟時末尾另加銷量分類概況一排）
    ＋ ① Safety Review × 分類維度交叉表（維度由 stats["dimension"] 推定，
    預設 Shop Nature）＋ ② Safety 調整方向分類
    ＋ ③ Safety 調整明細（SKU 級彙總）＋ ④ Buyer 統計（需調貨＋需調整 Safety 合併表）
    ＋ ⑤ 銷量 × 存貨分類交叉表（母體為全部店舖×SKU 列，不受調貨事件影響）
    ＋ ⑥ 全列 Safety 調整方向分類／明細（母體為全部列的預防性建議）。
    ②③ 由 raw 即時彙算（Safety 建議不受 OM 限制影響，兩變體共用同一份統計）；
    ④ 與 Stock Checking 的 Buyer 統計同一準則（同 SKU 首個非空 Buyer，已刪除重複）。"""
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    section_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    _rows = _hybrid_dashboard_metrics(raw, stats)
    row5_metrics = _rows[4] if len(_rows) > 4 else None
    row6_metrics = _rows[5] if len(_rows) > 5 else None
    row7_metrics = _rows[6] if len(_rows) > 6 else None
    _days = (stats or {}).get("safety_days") or _hybrid_safety_days_from_raw(raw)
    _write_metric_cards(
        ws, f"Safety Stock 設定 Review 摘要 (Transfer Diagnosis Dashboard｜"
            f"Safety 基準 {_hybrid_safety_days_label(_days)})",
        _rows[0], _rows[1], thin_border, desc_font, progress_callback,
        progress_pct=0.80, row3_metrics=_rows[2], row4_metrics=_rows[3],
        row5_metrics=row5_metrics, row6_metrics=row6_metrics, row7_metrics=row7_metrics,
    )

    table_start = 4 * len(_rows) + 4
    _block_kw = dict(header_fill=header_fill, header_font=header_font, bold_font=bold_font,
                     section_fill=section_fill, total_fill=total_fill, thin_border=thin_border,
                     center=center, left_wrap=left_wrap)
    _dim = _hybrid_dimension((stats or {}).get("dimension"))
    # ①–④ 共用同一母體（事件類型非空），① 為空即代表四區塊全無資料；
    # ⑤ 母體為全部列，不受影響，恆在四區塊之後寫入
    if cross is None or cross.empty:
        ws.cell(row=table_start, column=1,
                value="無符合資料（①–④ 統計區塊以調貨事件列為母體）").font = Font(size=12)
        next_row = table_start + 2
    else:
        next_row = _write_hybrid_summary_block(
            ws, table_start - 1, f"Safety Review × {_HYBRID_DIMENSION_LABELS[_dim]} 交叉表", cross, **_block_kw)
        next_row = _write_hybrid_summary_block(
            ws, next_row, "Safety 調整方向分類（增加 / 減少 / 不變）",
            _build_hybrid_safety_change_summary(raw), **_block_kw)
        next_row = _write_hybrid_summary_block(
            ws, next_row, "Safety 調整明細（SKU 級彙總）",
            _build_hybrid_safety_sku_detail(raw), **_block_kw)

        try:
            combined = _hybrid_buyer_combined_detail(raw)
        except Exception:
            buyer_need = (stats or {}).get("buyer_need_skus")
            if buyer_need is None:
                buyer_need = _hybrid_buyer_need_stats(raw)
            buyer_safety = (stats or {}).get("buyer_safety_fix")
            if buyer_safety is None:
                buyer_safety = _hybrid_buyer_safety_stats(raw)
            combined = _hybrid_buyer_combined_stats(buyer_need, buyer_safety)
        _table_kw = dict(thin_border=thin_border, header_fill=header_fill,
                         header_font=header_font, bold_font=bold_font,
                         section_fill=section_fill, total_fill=total_fill,
                         center=center)
        if combined:
            next_row = _write_hybrid_buyer_combined_table(ws, combined, start_row=next_row, **_table_kw)

    _oos_days = int((stats or {}).get("sales_oos_days") or _HYBRID_SALES_OOS_DAYS_DEFAULT)
    _over_days = int((stats or {}).get("sales_overstock_days") or _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT)
    _min_units = int((stats or {}).get("sales_overstock_min_units") or _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT)
    next_row = _write_hybrid_summary_block(
        ws, next_row,
        f"銷量 × 存貨分類交叉表（母體為全部店舖×SKU 列；缺貨線 {_oos_days} 日、"
        f"積壓線 {_over_days} 日、積壓保底 {_min_units} 件）",
        _build_hybrid_sales_summary(raw), **_block_kw)

    # ── ⑥ 全列預防性 Safety 設定建議（母體為全部列；與事件版 ②③ 同口徑、不同母體）──
    next_row = _write_hybrid_summary_block(
        ws, next_row, "全列 Safety 調整方向分類（增加 / 減少 / 不變）",
        _build_hybrid_full_safety_change_summary(raw), **_block_kw)
    _write_hybrid_summary_block(
        ws, next_row, "全列 Safety 調整明細（SKU 級彙總）",
        _build_hybrid_full_safety_sku_detail(raw), **_block_kw)

    if progress_callback:
        progress_callback(0.90, "Review 摘要完成...")
    _auto_width(ws)



def _hybrid_raw_display_columns(raw: pd.DataFrame) -> list[str]:
    """Safety Stock x NetStock Review RAW 顯示欄序：不顯示「Item_Shop」、所有含「Value」的
    金額欄與底線開頭的內部欄（如 `_hybrid_safety_floor3`），其餘欄位（含
    _compute_raw_extra_columns 衍生的 Shop Name / SKU Desc）按 _HYBRID_RAW_DISPLAY_GROUPS
    分組排列，未歸組的原始欄位按原序置於最後。"""
    cols = [str(c) for c in raw.columns] + ["Shop Name", "SKU Desc"]
    seen: set[str] = set()
    kept: list[str] = []
    for c in cols:
        if c in seen or c == "Item_Shop" or "value" in c.lower() or c.startswith("_"):
            continue
        seen.add(c)
        kept.append(c)
    kept_set = set(kept)
    ordered: list[str] = []
    for group in _HYBRID_RAW_DISPLAY_GROUPS:
        for c in group:
            if c in kept_set:
                ordered.append(c)
                kept_set.discard(c)
    ordered += [c for c in kept if c in kept_set]
    return ordered



def _write_hybrid_guide_sheet(ws, stats):
    """「說明」工作表（工作簿首位）：列出本次執行參數、講解模式核心邏輯，
    並逐頁導讀各工作表應留意的部分。內容為靜態說明＋stats 參數照錄，不含計算欄位。"""
    st_ = stats or {}
    dim = _hybrid_dimension(st_.get("dimension"))
    sd = {**_HYBRID_SAFETY_DAYS_DEFAULT, **(st_.get("safety_days") or {})}

    def _dw(text):
        # 顯示寬度估算（全形字元約 2、半形約 1），供換行高度估算
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

    _merged("Safety Stock x NetStock Review 說明與導讀", title_font, height=26)
    _merged("本頁講解此報告的 Review 邏輯與本次執行參數，並逐頁說明應留意的部分；各項判定以 RAW 逐列欄位為準。",
            sub_font, height=16)

    # ── ① 本次執行參數 ──
    _sec("① 本次執行參數（供核對）")
    _kv("分類維度", f"{_HYBRID_DIMENSION_LABELS[dim]}（只影響 Summary／交叉表分組，不影響調貨與 Safety 計算）")
    _kv("調貨門檻", f'{st_.get("turnday_threshold", 14)} 日（SKU 級 Nature 池加總 TurnDay）')
    _kv("D001 庫存門檻", f'{st_.get("d001_threshold", 100)} 件（D001 Net Stock Qty (CM) 低於此值才標示需要調貨）')
    _kv("Safety 基準周轉天數", f'Tourist {sd["Tourist"]} 日／Mix {sd["Mix"]} 日／Non Tourist {sd["Non Tourist"]} 日')
    _kv("允許雙向調貨", "開啟（任何 Nature 缺貨皆可調入，Tourist 店舖亦可轉出）" if st_.get("bidirectional")
        else "關閉（需求限 Tourist，供應限 Mix／Non Tourist）")
    _kv("ND 優先轉出", "開啟（同一 Nature 池內 ND 來源列先於 RF 列取用）" if st_.get("nd_priority")
        else "關閉（池內按來源檔列序取用）")
    _kv("銷量分類門檻",
        f'缺貨線 {st_.get("sales_oos_days", 7)} 日／積壓線 {st_.get("sales_overstock_days", 30)} 日／'
        f'保底 {st_.get("sales_overstock_min_units", 2)} 件（銷量存貨分類恆按此計算）')
    _kv("銷量再平衡", "開啟（高銷缺貨列豁免池門檻加入調貨需求方）" if st_.get("sales_rebalance")
        else "關閉（調貨需求判定按現行規則）")
    _kv("建議退D001", "開啟（中銷／低銷／零銷列退回超出 Safety 的庫存；高銷列須庫存高於 Safety 200% 才退至 2 倍 Safety）" if st_.get("return_to_d001")
        else "關閉（不產生退倉建議）")
    _kv("參考日期", f'{st_.get("ref_date", "—")}（Past 90 Days 0 Sales 轉 ND 提示的新品判定基準，預設昨日）')

    # ── ② 核心邏輯 ──
    _sec("② 核心邏輯（這份報告怎樣算出來）")
    _bullet("調貨段（固定 14 日，與 Safety 無關）：同 SKU 中某 Nature 池的加總 TurnDay（加總庫存 ÷ 加總月銷 × 30.5）"
            "低於調貨門檻，且 D001 低於門檻，該 Nature 內 RP Type＝RF 的缺貨店舖列標示「需要調貨」，"
            "需求量＝補到 14 日周轉；供應來源為 Mix／Non Tourist 池的 14 日超額（雙向時另加 Tourist）；"
            "池級歸零——缺貨池的列一律不可轉出。分配跑兩遍：「不限OM」與「同OM」（來源店須與需求店同 OM）。")
    _bullet("Review 段（Safety 基準天數，按 Shop Nature）：Safety 周轉所需量＝ceil(Sold Qty (P90D) ÷ 90 × 該 Nature 基準天數)。"
            "需要調入列：現有 Safety 設定 < 所需量 → 「Safety 設定不足」；有貨轉出列（實際轉出／可轉出未使用）："
            "現有設定 > 生效基準 → 「Safety 設定過高」；等於邊界屬設定正常。建議值＝不足列取所需量、"
            "過高列取生效基準（下調補底 3 件——SAP MOQ 使 1–2 件難以補貨）。建議 Safety 設定值 Turnday＝建議值 ÷ (P90D 月均 ÷ 30)；"
            "受保底影響（建議值 3 件、未套保底前 < 3）時，店舖問題 SKU 的建議方法加註保底原因。Safety 設定與調貨 14 日線互相獨立。")
    _bullet("全列預防性 Safety 設定建議（母體為全部店舖×SKU 列，不依賴調貨事件）：與 Review 段同口徑"
            "（所需量＋下調補底 3 件），惟不設調貨事件門檻——現有 < 所需量 → 不足（ND 列不作上調，只可下調），"
            "現有 > 生效基準 → 過高，其餘 → 設定正常。RAW「全列」五欄與「Safety 設定建議 (全列)」工作表"
            "供一次過設定適合的 Safety Stock，減少日後調貨次數；事件版 Review 欄維持只覆蓋調貨事件列。")
    _bullet("銷量帶（常開顯示欄）：同 SKU 內按 Sold Qty (P90D) 的店舖平均排名分帶——排名百分比 > 75%＝高銷、"
            "≤ 25%＝低銷、其餘＝中銷、P90D ≤ 0＝零銷；並列同值同帶；有銷售店舖少於 4 間時不設高銷／低銷（全歸中銷）。")
    _bullet("銷量存貨分類（常開顯示欄）：高銷且店舖周轉日數低於缺貨線 → 高銷缺貨；低銷且周轉不少於積壓線且庫存不少於保底件數 → "
            "低銷積壓；零銷且庫存不少於保底件數 → 零銷積壓；其餘 → 正常（門檻見上方 ①）。")
    _bullet("Past 90 Days 0 Sales 轉 ND：M-2／M-1／CM-MTD／P90D 皆零銷售且有庫存的列，於「店舖問題 SKU」"
            "標示 ND 退場建議（距參考日期 90 天內新品除外）。")
    if st_.get("return_to_d001"):
        _bullet("建議退D001（本份已開啟）：接收目標 D001＝Warehouse，可無限接收，不跑分配。"
                "退貨基準為 Safety Stock (現時)（非調貨段的 14 日線，與池級門檻脫鉤）：中銷／低銷／零銷列"
                "退回超出 Safety 的庫存（庫存 − Safety）；高銷列原則上不退，僅當 Safety > 0 且庫存 ≥ 3×Safety"
                "（高於 Safety 200%）時退至 2×Safety。調貨優先——建議退D001 量＝可退額度 − 轉出量（不限OM），"
                "需要調貨列一律不退。本欄只作退倉建議，不影響調貨分配、Safety Review 與建議調整量。")

    # ── ③ 工作表導讀 ──
    _sec("③ 工作表導讀（先看哪一頁、應留意什麼）")
    for col, val in ((1, "工作表"), (2, "內容"), (3, "建議留意")):
        c = ws.cell(row=r, column=col, value=val)
        c.font = label_font
        c.fill = head_fill
        c.alignment = Alignment(vertical="center")
    ws.row_dimensions[r].height = 16
    r += 1
    _row3("說明（本頁）", "Review 邏輯、本次參數與導讀", "執行任何行動前，先核對上方 ① 的參數是否符合預期")
    _row3("RAW", "全部店舖×SKU 逐列明細與判定欄",
          "留意「需要調貨」「事件類型」「Safety Review」「建議 Safety 設定值」「建議 Safety 設定值 Turnday」「Safety 建議調整量」"
          "與「全列」五欄（全列 Safety 周轉所需量／全列 Safety Review／全列建議 Safety 設定值／"
          "全列建議 Safety 設定值 Turnday／全列 Safety 建議調整量）；非事件列事件版 Review 欄留空屬正常，全列欄恆有值")
    _summary_watch = "留意「存貨偏重提示」（橘底）——哪些分組存貨偏重而銷量不高；另留意淺藍底——庫存佔比明顯低於銷量佔比的分組（潛在缺貨風險，只標示儲存格）"
    if dim in ("sales", "sales_nature"):
        _summary_watch += "；零銷組「加總 - Shop Stock Qty (CM) %」大於 0% 以紅底標示——零銷售而仍有存貨，理論上全屬 Overstock"
    _row3("Summary", "SKU 級庫存／銷量分佈（按分類維度分組）", _summary_watch)
    _row3("調貨 Review Safety Stock (不限OM)", "逐列 Safety Review 主表（調貨來源不限 OM）",
          "按「Safety Review」篩選設定不足／過高；行動看「建議 Safety 設定值」「建議 Safety 設定值 Turnday」與「Safety 建議調整量」")
    _row3("調貨 Review Safety Stock (同OM)", "同上，惟調貨來源須與需求店同一 OM",
          "兩套數值可能不同；(同OM) 限制較嚴，按公司規定選用")
    _row3("店舖問題 SKU", "每間店舖需行動事項（問題＋建議方法）",
          "店務同事可先從這一頁著手；建議方法附建議值周轉日數，受 3 件保底影響時加註保底原因；含調入缺口跟進與 0 Sales 轉 ND 建議")
    if st_.get("return_to_d001"):
        _row3("退D001 建議", "超額庫存退回 D001（Warehouse）的店舖×SKU 清單",
              "留意「可退D001 貨量」（庫存 − Safety；高銷為庫存 − 2×Safety）與「轉出量」的差額即建議退回量")
    _row3("Review 摘要", "指標卡＋Safety Review×分類維度交叉表＋Safety 調整方向分類／明細＋銷量×存貨分類交叉表＋全列 Safety 調整方向分類／明細",
          "管理層概覽；「Safety 調整方向分類」看事件列建議增加／減少的總量，「全列」分類看全部列預防性建議總量")
    _row3("Safety 設定建議 (全列)", "全部店舖×SKU 的預防性 Safety 設定建議（只列需調整列）",
          "一次過設定 Safety 用：按「全列 Safety Review」篩選需增加／需減少，行動看「全列建議 Safety 設定值」與「全列 Safety 建議調整量」")
    _row3("簡易調貨明細 (不限OM)／(同OM)", "一列一筆「來源店舖 → 接收店舖」調貨",
          "實際執行調貨時按此表進行：來源店、接收店、調貨數量")

    # ── ④ 實用提示 ──
    _sec("④ 實用提示")
    _bullet("「Safety 建議調整量」正數＝建議上調設定；負數＝建議下調設定（僅不足／過高列有值）；"
            "「全列 Safety 建議調整量」同口徑，惟母體為全部列（需調整列有值），供預防性一次過設定。")
    _bullet("Summary 橘底＝觸發「存貨偏重提示」的分組（庫存佔比高而銷量佔比低）。")
    _bullet(f"Summary 淺藍底＝存貨偏低分組（{_hybrid_shortage_hint_text(dim)}，門檻較偏重側寬鬆）——庫存佔比明顯低於銷量佔比，提示潛在缺貨風險；只標示庫存佔比欄，不另設文字欄。")
    if dim in ("sales", "sales_nature"):
        _bullet("Summary 紅底＝零銷組仍有庫存（「加總 - Shop Stock Qty (CM) %」大於 0%）——零銷售而仍有存貨，"
                "理論上全屬 Overstock，不設佔比門檻一律標示；同格同時觸發橘底時以紅底優先。")
    _bullet("分類維度只改變 Summary／交叉表的分組顯示，不影響調貨與 Safety 計算。")
    _bullet("「可轉出未使用」＝該列有 14 日超額庫存但本次未被取用，可留作後備來源。")
    if st_.get("return_to_d001"):
        _bullet("「建議退D001 量」＝可退D001 貨量 − 轉出量（不限OM 分配實績）；(同OM) 分配轉出較少時，"
                "實際可退量可能更多，執行前請按選用的調貨變體自行核對。")



def generate_excel_hybrid(
    raw: pd.DataFrame,
    sc_summary: pd.DataFrame | None = None,
    diagnosis: pd.DataFrame | None = None,
    diagnosis_om: pd.DataFrame | None = None,
    diag_summary: pd.DataFrame | None = None,
    stats: dict | None = None,
    progress_callback=None,
    include_raw: bool = True,
) -> BytesIO:
    """Safety Stock x NetStock Review 輸出（10 工作表，順序固定；開啟「建議退D001」時 11 表）：
    說明（首頁導讀：本次執行參數列出＋模式核心邏輯＋各工作表應留意的部分）/
    RAW（重新排列欄序提升可讀性，不顯示 Item_Shop 與 Value 金額欄；含不限OM與同OM兩套調貨結果；
    雙向時另含「可轉出貨量 (Tourist)」，開啟「建議退D001」時另含「建議退D001 量」）/
    Summary（SKU 級彙總，按 stats["dimension"] 分類維度分組：nature 三組／rp 兩組／both 六組／
    om 動態 OM 名單／nature_om 三 Nature × OM／sales 四銷量帶／
    sales_nature 三 Nature × 四帶 12 組（動態 OM 名單由 stats["om_groups"] 還原，
    缺失時由 raw 即時推導）；
    各組加總/TurnDay/涉及店舖數/佔比、強制調貨提示、存貨偏重提示（門檻按維度切換）；
    sales／sales_nature 維度另以醒目紅底標示零銷組庫存佔比 > 0%（零銷組仍有庫存，
    零銷售而仍有存貨理論上全屬 Overstock）；
    nature 雙向時另加 AD 欄「由Tourist 店舖預計調出數量」）/
    調貨 Review Safety Stock (不限OM) / 調貨 Review Safety Stock (同OM)（店舖×SKU 主表含 RP Type 欄）/
    店舖問題 SKU（每間店舖需行動列：問題＋建議方法）/
    退D001 建議（「建議退D001 量」欄存在時才輸出：超額庫存扣除已轉出後退回 D001 的店舖×SKU 清單）/
  Review 摘要（指標卡 4–6 排（雙向或 ND 店舖有實際轉出時 5 排；銷量再平衡開啟時末尾另加銷量分類概況一排）
  + Safety Review × 分類維度交叉表
 + Safety 調整方向分類 + Safety 調整明細（SKU 級彙總）
 + Buyer 統計（需調貨＋需調整 Safety 合併表）
 + 銷量 × 存貨分類交叉表（母體為全部店舖×SKU 列））/
    簡易調貨明細 (不限OM) / 簡易調貨明細 (同OM)（一列一筆來源店→接收店調貨，來源字串已拆行）。
    雙向與否由 RAW 是否含「可轉出貨量 (Tourist)」欄推定；分類維度由 stats["dimension"] 推定
    （缺失時退回 nature）。
    """
    if diagnosis is None:
        diagnosis = _build_hybrid_diagnosis(raw)
    if diagnosis_om is None:
        diagnosis_om = _build_hybrid_diagnosis(raw, variant="om")
    _dim = _hybrid_dimension((stats or {}).get("dimension"))
    _om_groups = (stats or {}).get("om_groups")
    if _om_groups is None and raw is not None and not raw.empty:
        _om_groups = _hybrid_om_groups_from_values(
            raw["OM"] if "OM" in raw.columns else _safe_default_series("", raw.index)
        )
    _om_groups = tuple(_om_groups) if _om_groups else None
    _ref_date = None
    try:
        if stats and stats.get("ref_date"):
            _ref_date = _normalize_ref_date(stats.get("ref_date"))
    except ValueError:
        _ref_date = None
    if diag_summary is None:
        diag_summary = _build_hybrid_diag_summary(raw, dimension=_dim, om_groups=_om_groups)
    _bidirectional = raw is not None and "可轉出貨量 (Tourist)" in raw.columns
    if sc_summary is None:
        sc_summary = _build_stock_checking_summary(raw, bidirectional=_bidirectional, dimension=_dim, om_groups=_om_groups)
    wb = Workbook()
    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    desc_font = Font(size=9, color="6B7280", italic=True)

    # ── Sheet 0: 說明（邏輯、本次執行參數與導讀；置於工作簿首位）──
    ws_guide = wb.active
    ws_guide.title = "說明"
    _write_hybrid_guide_sheet(ws_guide, stats)

    # ── Sheet 1: RAW（重排欄序；不顯示 Item_Shop 與 Value 金額欄）──
    ws_raw = wb.create_sheet(title="RAW")
    (_write_raw_sheet(ws_raw, raw, progress_callback, progress_start=0.0, progress_end=0.24,
                     columns=_hybrid_raw_display_columns(raw)) if include_raw else _write_raw_placeholder(ws_raw))

    if progress_callback:
        progress_callback(0.25, "建立 SKU Summary（存貨偏重提示）...")

    # ── Sheet 2: Summary（SKU 級彙總，按分類維度分組；雙向時另加 AD 欄；偏重提示以醒目橘底加強顯示）──
    ws_sc = wb.create_sheet(title="Summary")
    _write_stock_checking_summary_sheet(ws_sc, sc_summary, progress_callback, progress_start=0.25, progress_end=0.38, bidirectional=_bidirectional, highlight_skew=True, dimension=_dim, om_groups=_om_groups, highlight_zero_sales=True)

    if progress_callback:
        progress_callback(0.40, "建立調貨 Review 主表 (不限OM)...")

    # ── Sheet 3: 調貨 Review Safety Stock（不限OM）──
    ws_diag = wb.create_sheet(title="調貨 Review Safety Stock (不限OM)")
    _write_hybrid_diagnosis_sheet(ws_diag, diagnosis, progress_callback, progress_start=0.40, progress_end=0.56)

    if progress_callback:
        progress_callback(0.57, "建立調貨 Review 主表 (同OM)...")

    # ── Sheet 4: 調貨 Review Safety Stock（同OM）──
    ws_diag_om = wb.create_sheet(title="調貨 Review Safety Stock (同OM)")
    _write_hybrid_diagnosis_sheet(ws_diag_om, diagnosis_om, progress_callback, progress_start=0.57, progress_end=0.70)

    if progress_callback:
        progress_callback(0.71, "建立店舖問題 SKU...")

    # ── Sheet 5: 店舖問題 SKU（Safety 設定不足／過高＋ND 退場；空表寫「無符合資料」）──
    ws_shop_issue = wb.create_sheet(title="店舖問題 SKU")
    _write_shop_issue_sheet(
        ws_shop_issue, _build_hybrid_shop_issues(raw, ref_date=_ref_date),
        title="店舖問題 SKU — 每間店舖需行動列（問題＋建議方法）",
        progress_callback=progress_callback, progress_start=0.71, progress_end=0.74,
    )

    # ── Sheet 5.5: 退D001 建議（「建議退D001 量」欄存在時才輸出，工作表 9→10）──
    _has_return_d001 = raw is not None and _HYBRID_RETURN_D001_COLUMN in raw.columns
    if _has_return_d001:
        if progress_callback:
            progress_callback(0.745, "建立退D001 建議...")
        ws_return = wb.create_sheet(title=_HYBRID_RETURN_D001_SHEET)
        _write_shop_issue_sheet(
            ws_return, _build_hybrid_return_d001(raw),
            title="退D001 建議 — 中銷／低銷／零銷退回超出 Safety 的庫存；高銷須庫存高於 Safety 200%（退至 2 倍 Safety）；退回 D001（Warehouse）",
            progress_callback=progress_callback, progress_start=0.745, progress_end=0.75,
        )

    if progress_callback:
        progress_callback(0.75, "建立 Review 摘要...")

    # ── Sheet 6: Review 摘要 ──
    ws_sum = wb.create_sheet(title="Review 摘要")
    _write_hybrid_diag_summary_sheet(ws_sum, diag_summary, stats, raw, thin_border, desc_font, progress_callback)

    if progress_callback:
        progress_callback(0.90, "建立 Safety 設定建議 (全列)...")

    # ── Sheet: Safety 設定建議 (全列)（預防性建議，只列需調整列）──
    ws_full = wb.create_sheet(title=_HYBRID_FULL_SAFETY_SHEET)
    _write_hybrid_full_safety_sheet(
        ws_full, _build_hybrid_full_safety_detail(raw),
        progress_callback, progress_start=0.90, progress_end=0.905,
    )

    if progress_callback:
        progress_callback(0.905, "建立簡易調貨明細 (不限OM)...")

    # ── Sheet 6: 簡易調貨明細（不限OM）──
    ws_detail = wb.create_sheet(title="簡易調貨明細 (不限OM)")
    _write_hybrid_transfer_detail_sheet(
        ws_detail, _build_hybrid_transfer_detail(raw, variant="free"),
        progress_callback, progress_start=0.905, progress_end=0.93,
    )

    if progress_callback:
        progress_callback(0.93, "建立簡易調貨明細 (同OM)...")

    # ── Sheet 7: 簡易調貨明細（同OM）──
    ws_detail_om = wb.create_sheet(title="簡易調貨明細 (同OM)")
    _write_hybrid_transfer_detail_sheet(
        ws_detail_om, _build_hybrid_transfer_detail(raw, variant="om"),
        progress_callback, progress_start=0.93, progress_end=0.95,
    )

    if progress_callback:
        progress_callback(0.96, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.active = 0
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")

    return buf

