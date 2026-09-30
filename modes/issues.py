# -*- coding: utf-8 -*-
"""店舖問題 SKU 共用（Stock Checking／Safety Review／Hybrid 共用）：由 transformer.py 拆分而來的共用核心。
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


from transformer import (
    _HYBRID_P90D_ZERO_SALES_ANNOTATION,
)


from modes.common import (
    _auto_width,
    _cached_shop_names,
    _display_safety_header,
    _extract_shop_name,
    _iter_header_rows,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
)


from modes.framework import (
    _HYBRID_DIAG_OVER,
    _HYBRID_DIAG_UNDER,
    _HYBRID_SAFETY_DAYS_COLUMN,
    _HYBRID_SAFETY_LINE_COLUMN,
    _HYBRID_SUGGEST_TURNDAY_COLUMN,
    _hybrid_p90d_zero_sales_mask,
)
__all__ = [
    "_COMBINED_SHOP_ISSUE_COLUMNS",
    "_COMBINED_SHOP_ISSUE_NUM_COLUMNS",
    "_HYBRID_SHOP_ISSUE_COLUMNS",
    "_HYBRID_SHOP_ISSUE_NUM_COLUMNS",
    "_SAFETY_FLOOR3_HYBRID_NOTE",
    "_SAFETY_FLOOR3_SSR_NOTE",
    "_SHOP_ISSUE_BASE_COLUMNS",
    "_SHOP_ISSUE_SEVERITY",
    "_STOCK_CHECKING_SHOP_ISSUE_COLUMNS",
    "_STOCK_CHECKING_SHOP_ISSUE_NUM_COLUMNS",
    "_build_combined_shop_issues",
    "_build_hybrid_shop_issues",
    "_build_stock_checking_shop_issues",
    "_count_transfer_sources",
    "_d001_available_qty",
    "_d001_gap_advice",
    "_is_warehouse_supply",
    "_safety_delta_text",
    "_shop_issue_base_frame",
    "_shop_name_lookup",
    "_suggest_turnday_clause",
    "_write_shop_issue_sheet",
]


# 三張新表欄序（新契約：固定表頭，未經明確要求不得改動）
_SHOP_ISSUE_BASE_COLUMNS = [
    "Site Code", "Shop Name", "OM", "Shop Nature",
    "Brand", "SKU", "SKU Desc", "Buyer",
    "問題類型", "問題說明", "建議方法",
]



_STOCK_CHECKING_SHOP_ISSUE_NUM_COLUMNS = [
    "Supply Source", "D001 Net Stock Qty (CM)",
    "Shop Stock Qty (CM)", "Shop Monthly Sold Qty (P90D)",
    "調入需求量 (補到14日)", "實際調入量", "調入缺口",
    "可轉出貨量 (Mix)", "可轉出貨量 (Non Tourist)", "轉出量",
    "調貨來源 (Site Code)",
    "銷量帶", "銷量存貨分類",
]



_COMBINED_SHOP_ISSUE_NUM_COLUMNS = [
    "Shop Stock Qty (CM)", "Shop Monthly Sold Qty (P90D)", "Turnday (P90D Qty)",
    "Adj. Safety Stock", "Recommend Safety Stock", "Safety Qty Change",
    "Overstock Cap (X)", "OOS Floor (Y)",
    "Safety 基準天數", "Safety 周轉所需量",
    "Recommendation of new Safety Qty / ND",
    "銷量帶", "銷量存貨分類",
]



_HYBRID_SHOP_ISSUE_NUM_COLUMNS = [
    "Supply Source", "D001 Net Stock Qty (CM)",
    "Shop Stock Qty (CM)", "Shop Monthly Sold Qty (P90D)",
    "Safety 基準天數", "Safety 周轉所需量",
    "Adj. Safety Stock", "建議 Safety 設定值", "Safety 建議調整量",
    "調入需求量 (補到14日)", "實際調入量", "調入缺口",
    "可轉出貨量", "轉出量", "調貨來源 (Site Code)",
    "實際調入量 (同OM)", "調入缺口 (同OM)", "轉出量 (同OM)",
]



_STOCK_CHECKING_SHOP_ISSUE_COLUMNS = _SHOP_ISSUE_BASE_COLUMNS + _STOCK_CHECKING_SHOP_ISSUE_NUM_COLUMNS



_COMBINED_SHOP_ISSUE_COLUMNS = _SHOP_ISSUE_BASE_COLUMNS + _COMBINED_SHOP_ISSUE_NUM_COLUMNS



_HYBRID_SHOP_ISSUE_COLUMNS = _SHOP_ISSUE_BASE_COLUMNS + _HYBRID_SHOP_ISSUE_NUM_COLUMNS



# 店內問題嚴重度（數字越小越優先；缺口最緊急）
_SHOP_ISSUE_SEVERITY = {
    "調入缺口（來源不足）": 0,
    "需要調入": 1,
    "Safety 設定不足": 1,
    "Safety 設定過高": 2,
    "可轉出供應": 3,
    "ND 退場": 3,
}



def _shop_name_lookup(raw: pd.DataFrame) -> dict[str, str]:
    """Site Code → Shop Name 查找表（RAW 已由 Item_Shop 衍生 Shop Name；缺欄時回退解析）。"""
    lookup: dict[str, str] = {}
    if raw is None or raw.empty or "Site Code" not in raw.columns:
        return lookup
    site = raw["Site Code"].fillna("").astype(str).str.strip()
    if "Shop Name" in raw.columns:
        names = raw["Shop Name"].fillna("").astype(str).str.strip()
    elif "Item_Shop" in raw.columns:
        # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
        names = pd.Series(
            _cached_shop_names(raw["Item_Shop"], site),
            index=raw.index,
        )
    else:
        return lookup
    for site_v, name_v in zip(site.to_numpy(), names.to_numpy()):
        if site_v and name_v and site_v not in lookup:
            lookup[site_v] = str(name_v)
    return lookup



def _shop_issue_base_frame(raw: pd.DataFrame, mask: pd.Series) -> pd.DataFrame:
    """新表共用底：篩出需行動列並帶出識別欄（Site Code／Shop Name／OM／Nature／品項）。"""
    sel = raw.loc[mask].copy()
    if sel.empty:
        return sel
    if "Site Code" not in sel.columns:
        sel["Site Code"] = ""
    site = sel["Site Code"].fillna("").astype(str).str.strip()
    sel["Site Code"] = site
    names = _shop_name_lookup(raw)
    if "Shop Name" in sel.columns:
        own = sel["Shop Name"].fillna("").astype(str).str.strip()
        sel["Shop Name"] = [o if o else names.get(s, "") for o, s in zip(own.to_numpy(), site.to_numpy())]
    else:
        sel["Shop Name"] = site.map(names).fillna("")
    if "OM" not in sel.columns:
        sel["OM"] = ""
    else:
        sel["OM"] = sel["OM"].fillna("").astype(str).str.strip()
    if "Shop Nature" not in sel.columns:
        sel["Shop Nature"] = ""
    return sel



def _count_transfer_sources(src: str) -> int:
    """計算調貨來源店舖數量：「HB30:2; HB49:3」→ 2；空白／無法解析 → 0。"""
    if not src:
        return 0
    count = 0
    for part in str(src).split(";"):
        part = part.strip()
        if not part:
            continue
        supply_site, sep, _qty_text = part.rpartition(":")
        if sep and supply_site.strip():
            count += 1
    return count



def _d001_available_qty(raw: pd.DataFrame) -> pd.Series:
    """店舖問題表共用：逐列實際採用的 D001 數量。

    優先序與觸發判定一致 — Stock Checking 用 Net→Stock→0；Hybrid 用逐列
    Net→Stock→0 的 Effective 值（缺欄時退回 Net→0）。非數值視為 0。
    """
    idx = raw.index
    if "D001 Effective Stock Qty (CM)" in raw.columns:
        return pd.to_numeric(raw["D001 Effective Stock Qty (CM)"], errors="coerce").fillna(0.0)
    net_col = _resolve_column_name("D001 Net Stock Qty (CM)", set(raw.columns))
    stock_col = _resolve_column_name("D001 Stock Qty (CM)", set(raw.columns))
    net = pd.to_numeric(raw[net_col], errors="coerce") if net_col else _safe_default_series(np.nan, idx)
    stock = pd.to_numeric(raw[stock_col], errors="coerce") if stock_col else _safe_default_series(np.nan, idx)
    return net.where(net.notna(), stock).fillna(0.0)



def _is_warehouse_supply(raw: pd.DataFrame) -> pd.Series:
    """店舖問題表共用：Supply Source 是否為 Warehouse（大小寫與前後空白不敏感）。

    只有 Warehouse 由 D001 供貨；Vendor 由供應商直送店舖，D001 有無貨與其無關；
    缺欄或空白一律視為非 Warehouse（保守：不作 D001 相關建議）。
    """
    if "Supply Source" not in raw.columns:
        return pd.Series(False, index=raw.index)
    return raw["Supply Source"].fillna("").astype(str).str.strip().str.lower() == "warehouse"



def _d001_gap_advice(d001_qty: float, gap_n: int, is_warehouse: bool) -> str:
    """店舖問題表共用：調入缺口的跟進行動（只在缺口 > 0 時呼叫）。

    - Warehouse＋D001 有貨：可由 D001 補貨。
    - Warehouse＋D001 無貨：D001缺貨。
    - Vendor：不經 D001，請向供應商追貨。
    """
    if not is_warehouse:
        return "Vendor 貨不經 D001 請向供應商追貨"
    if d001_qty > 0:
        return f"D001 尚有 {d001_qty:g} 件可補貨，缺口部分請由 D001 補貨"
    return "D001缺貨"



# 建議方法加註：建議值受保底 3 件影響時，說明保底原因（保障自動補貨之最低設定）。
# SSR 的 3 件來自基準底線 clip(min=3)（月銷倍率 OOS 底線或周轉天數所需量）；Hybrid 的 3 件來自 SAP MOQ 下調補底。
_SAFETY_FLOOR3_SSR_NOTE = "；為保障自動補貨，最低建議設定為 3 件（最低 3 件保底）"



_SAFETY_FLOOR3_HYBRID_NOTE = "；為保障自動補貨，最低建議設定為 3 件（SAP MOQ：Safety 1–2 件難以補貨）"



def _suggest_turnday_clause(turnday) -> str:
    """建議值的周轉子句（供店舖問題 SKU 建議方法用）。

    月均為 0／未能計算（NaN）時回傳空字串；否則回傳「；建議周轉約 X 日」。
    """
    if turnday is None or pd.isna(turnday):
        return ""
    return f"；建議周轉約 {float(turnday):g} 日"



def _safety_delta_text(delta) -> str:
    """Safety 調整量的文字（供店舖問題 SKU 建議方法用）。

    正數回傳「增加 X 件」、負數回傳「減少 X 件」；未能計算（NaN）回傳「未能確認」。
    """
    if delta is None or pd.isna(delta):
        return "未能確認"
    val = float(delta)
    if val > 0:
        return f"增加 {val:g} 件"
    if val < 0:
        return f"減少 {abs(val):g} 件"
    return f"{val:g} 件"



def _build_stock_checking_shop_issues(
    raw: pd.DataFrame,
    turnday_threshold: int = 14,
    d001_threshold: int = 100,
) -> pd.DataFrame:
    """Stock Checking 店舖問題表：只收需行動列（需要調入／可轉出供應）。

    - 需要調入：`需要調貨 == 是`；調入缺口 > 0 時問題類型記為「調入缺口（來源不足）」。
    - 可轉出供應：可轉出貨量或轉出量 > 0 且非需求列。
    - 問題說明帶 SKU 級 Tourist 加總 TurnDay 與 D001 觸發原因；建議方法帶數量與行動指引。
    - 排序：Site Code → 嚴重度 → Brand／SKU；門檻僅用於問題說明文字，不重算判定。
    """
    cols = _STOCK_CHECKING_SHOP_ISSUE_COLUMNS
    if raw is None or raw.empty or "需要調貨" not in raw.columns:
        return pd.DataFrame(columns=cols)
    need = raw["需要調貨"].fillna("").astype(str).str.strip() == "是"
    mix_avail = _safe_numeric(raw.get("可轉出貨量 (Mix)", _safe_default_series(0, raw.index)))
    nt_avail = _safe_numeric(raw.get("可轉出貨量 (Non Tourist)", _safe_default_series(0, raw.index)))
    donated = _safe_numeric(raw.get("轉出量", _safe_default_series(0, raw.index)))
    gap = _safe_numeric(raw.get("調入缺口", _safe_default_series(0, raw.index)))
    supply = (~need) & ((mix_avail > 0) | (nt_avail > 0) | (donated > 0))
    mask = need | supply
    if not bool(mask.any()):
        return pd.DataFrame(columns=cols)
    work = _shop_issue_base_frame(raw, mask)
    if "Supply Source" not in work.columns:
        work["Supply Source"] = raw.get("Supply Source", _safe_default_series("", work.index))
    if "D001 Net Stock Qty (CM)" not in work.columns:
        net_col = _resolve_column_name("D001 Net Stock Qty (CM)", set(raw.columns))
        if net_col is not None:
            work["D001 Net Stock Qty (CM)"] = raw[net_col]
    need_w = work["需要調貨"].fillna("").astype(str).str.strip() == "是"
    gap_w = _safe_numeric(work.get("調入缺口", _safe_default_series(0, work.index)))
    ttd_w = pd.to_numeric(work.get("Tourist 加總 TurnDay", _safe_default_series(np.nan, work.index)), errors="coerce")
    d001_qty_w = _d001_available_qty(raw).loc[work.index]
    d001_w = work.get("D001 Net Stock Qty (CM)", _safe_default_series("", work.index))
    warehouse_w = _is_warehouse_supply(raw).loc[work.index]
    demand_w = _safe_numeric(work.get("調入需求量 (補到14日)", _safe_default_series(0, work.index))).astype(int)
    actual_w = _safe_numeric(work.get("實際調入量", _safe_default_series(0, work.index))).astype(int)
    src_w = work.get("調貨來源 (Site Code)", _safe_default_series("", work.index)).fillna("").astype(str).str.strip()
    mix_w = _safe_numeric(work.get("可轉出貨量 (Mix)", _safe_default_series(0, work.index))).astype(int)
    nt_w = _safe_numeric(work.get("可轉出貨量 (Non Tourist)", _safe_default_series(0, work.index))).astype(int)
    out_w = _safe_numeric(work.get("轉出量", _safe_default_series(0, work.index))).astype(int)

    issue, desc, advice, sev = [], [], [], []
    for i in range(len(work)):
        if bool(need_w.iloc[i]):
            ttd_txt = f"{float(ttd_w.iloc[i]):.1f} 日" if pd.notna(ttd_w.iloc[i]) else "未能計算"
            d001_txt = str(d001_w.iloc[i]).strip() or "未能確認"
            gap_n = int(gap_w.iloc[i])
            if gap_n > 0:
                issue.append("調入缺口（來源不足）")
                desc.append(
                    f"Tourist 加總 TurnDay {ttd_txt}低於門檻 {int(turnday_threshold)} 日，"
                    f"D001 僅 {d001_txt} 件（門檻 {int(d001_threshold)} 件），需調入卻無足額來源"
                )
                src_n = _count_transfer_sources(str(src_w.iloc[i]))
                advice.append(
                    f"建議調入 {int(demand_w.iloc[i])} 件（實際可調入 {int(actual_w.iloc[i])} 件，"
                    f"缺口 {gap_n} 件，涉及 {src_n} 間來源店舖）；"
                    f"{_d001_gap_advice(float(d001_qty_w.iloc[i]), gap_n, bool(warehouse_w.iloc[i]))}"
                )
                sev.append(_SHOP_ISSUE_SEVERITY["調入缺口（來源不足）"])
            else:
                issue.append("需要調入")
                desc.append(
                    f"Tourist 加總 TurnDay {ttd_txt}低於門檻 {int(turnday_threshold)} 日，"
                    f"D001 僅 {d001_txt} 件（門檻 {int(d001_threshold)} 件），店舖庫存不足 14 日周轉"
                )
                src_n = _count_transfer_sources(str(src_w.iloc[i]))
                advice.append(
                    f"建議調入 {int(demand_w.iloc[i])} 件（實際可調入 {int(actual_w.iloc[i])} 件，"
                    f"涉及 {src_n} 間來源店舖），詳見「調貨來源 (Site Code)」欄"
                )
                sev.append(_SHOP_ISSUE_SEVERITY["需要調入"])
        else:
            issue.append("可轉出供應")
            desc.append("店舖庫存超過 14 日周轉，可向缺貨店舖供應")
            advice.append(
                f"可轉出 Non Tourist {int(nt_w.iloc[i])} 件、Mix {int(mix_w.iloc[i])} 件"
                f"（實際轉出 {int(out_w.iloc[i])} 件），詳見「調貨來源 (Site Code)」欄"
            )
            sev.append(_SHOP_ISSUE_SEVERITY["可轉出供應"])
    work["問題類型"] = issue
    work["問題說明"] = desc
    work["建議方法"] = advice
    work["_sev"] = sev
    if "Brand" not in work.columns:
        work["Brand"] = ""
    if "SKU" not in work.columns:
        work["SKU"] = ""
    work = work.sort_values(["Site Code", "_sev", "Brand", "SKU"], kind="stable").drop(columns=["_sev"])
    return work[[c for c in cols if c in work.columns]].reset_index(drop=True)



def _build_combined_shop_issues(raw: pd.DataFrame) -> pd.DataFrame:
    """Safety Stock Review 店舖問題表：只收 Safety 需調整與 ND 退場列。

    - Safety 需增加／減少：`Safety Qty Change != 0`。
    - ND 退場：`Recommendation of new Safety Qty / ND` 為 ND 系列或 0（已是 ND）；
      `New SKU (3 Month)` 屬資訊標記，不收錄。
    - Overstock／OOS 只作問題說明上下文，不單獨成列。
    - 建議方法以「增加／減少 X 件」標明調整方向，並附「建議周轉約 X 日」（`Recommend Safety Stock Turnday`；月均 0 略去）；
      `_safety_floor3` 為 True（建議值受最低 3 件保底影響）時加註保底原因。
    """
    cols = _COMBINED_SHOP_ISSUE_COLUMNS
    if raw is None or raw.empty:
        return pd.DataFrame(columns=cols)
    chg = pd.to_numeric(raw.get("Safety Qty Change", _safe_default_series(np.nan, raw.index)), errors="coerce").fillna(0)
    nd_rec = raw.get("Recommendation of new Safety Qty / ND", _safe_default_series("", raw.index)).fillna("").astype(str).str.strip()
    is_new_sku = nd_rec == "New SKU (3 Month)"
    is_nd_exit = nd_rec.str.startswith("ND") | (nd_rec == "0")
    mask = ((chg != 0) | (is_nd_exit & ~is_new_sku)).fillna(False)
    if not bool(mask.any()):
        return pd.DataFrame(columns=cols)
    work = _shop_issue_base_frame(raw, mask)
    chg_w = pd.to_numeric(work.get("Safety Qty Change", _safe_default_series(0, work.index)), errors="coerce").fillna(0)
    nd_w = work.get("Recommendation of new Safety Qty / ND", _safe_default_series("", work.index)).fillna("").astype(str).str.strip()
    adj_w = pd.to_numeric(work.get("Adj. Safety Stock", _safe_default_series(np.nan, work.index)), errors="coerce")
    rec_w = pd.to_numeric(work.get("Recommend Safety Stock", _safe_default_series(np.nan, work.index)), errors="coerce")
    rec_td_w = pd.to_numeric(
        work.get("Recommend Safety Stock Turnday", _safe_default_series(np.nan, work.index)), errors="coerce"
    )
    floor_w = (
        work.get("_safety_floor3", _safe_default_series(False, work.index))
        .fillna(False)
        .astype(bool)
    )
    status_w = work.get("Shop Stock Status", _safe_default_series("", work.index)).fillna("").astype(str).str.strip()

    issue, desc, advice, sev = [], [], [], []
    for i in range(len(work)):
        delta = float(chg_w.iloc[i])
        rec_txt = nd_w.iloc[i]
        adj_txt = f"{float(adj_w.iloc[i]):g}" if pd.notna(adj_w.iloc[i]) else "未能確認"
        rec_val_txt = f"{float(rec_w.iloc[i]):g}" if pd.notna(rec_w.iloc[i]) else "未能確認"
        td_clause = _suggest_turnday_clause(rec_td_w.iloc[i])
        floor_note = _SAFETY_FLOOR3_SSR_NOTE if bool(floor_w.iloc[i]) else ""
        status_txt = str(status_w.iloc[i]).strip()
        status_suffix = f"，狀態為{status_txt}" if status_txt else ""
        if delta > 0:
            issue.append("Safety 設定不足")
            desc.append(f"現時 Safety {adj_txt} 件低於建議 {rec_val_txt} 件（需增加 {delta:g} 件）{status_suffix}")
            advice.append(
                f"建議將 Safety 由 {adj_txt} 件調整為 {rec_val_txt} 件（{_safety_delta_text(delta)}{td_clause}），"
                f"請按建議值補足安全庫存{floor_note}"
            )
            sev.append(_SHOP_ISSUE_SEVERITY["Safety 設定不足"])
        elif delta < 0:
            issue.append("Safety 設定過高")
            desc.append(f"現時 Safety {adj_txt} 件高於建議 {rec_val_txt} 件（可減少 {abs(delta):g} 件）{status_suffix}")
            advice.append(
                f"建議將 Safety 由 {adj_txt} 件調整為 {rec_val_txt} 件（{_safety_delta_text(delta)}{td_clause}），"
                f"請按建議值調低安全庫存{floor_note}"
            )
            sev.append(_SHOP_ISSUE_SEVERITY["Safety 設定過高"])
        elif rec_txt.startswith("ND"):
            issue.append("ND 退場")
            desc.append(f"過去 90 日零銷售，建議轉為 ND 停補貨{status_suffix}（建議：{rec_txt}）")
            advice.append(f"建議轉為 {rec_txt}，現有庫存全數轉出或退倉（現時 Safety {adj_txt} 件）")
            sev.append(_SHOP_ISSUE_SEVERITY["ND 退場"])
        else:
            issue.append("ND 退場")
            desc.append(f"RP Type 已是 ND，建議 Safety 維持 0 件{status_suffix}")
            advice.append("已是 ND，無須補貨；現有庫存請按退倉建議處理")
            sev.append(_SHOP_ISSUE_SEVERITY["ND 退場"])
    work["問題類型"] = issue
    work["問題說明"] = desc
    work["建議方法"] = advice
    work["_sev"] = sev
    if "Brand" not in work.columns:
        work["Brand"] = ""
    if "SKU" not in work.columns:
        work["SKU"] = ""
    work = work.sort_values(["Site Code", "_sev", "Brand", "SKU"], kind="stable").drop(columns=["_sev"])
    return work[[c for c in cols if c in work.columns]].reset_index(drop=True)



def _build_hybrid_shop_issues(raw: pd.DataFrame, ref_date=None) -> pd.DataFrame:
    """Safety Stock x NetStock Review 店舖問題表：Safety 設定不足／過高＋ND 退場列。

    - 母體為 Safety 需調整列（不足／過高）加上 Past 90 Days 0 Sales 轉 ND 候選
      （四期銷售皆 ≤ 0 且有庫存、距參考日期 90 天內新品與需要調貨列除外，見
      `_hybrid_p90d_zero_sales_mask`；參考日期預設為今天，建議傳入 MTD 截止日期）；設定正常列一律排除。
    - 已屬不足／過高且同時符合零銷售的列保留原問題類型，建議方法末尾加註；
      僅零銷售的列以「ND 退場」收錄（每店×SKU 最多一列）。
    - 主表以不限OM事件為準，另附三個同OM對照欄；調入缺口 > 0 時在建議方法加註。
    - 建議方法以「增加／減少 X 件」標明調整方向，並附「建議周轉約 X 日」（`建議 Safety 設定值 Turnday`；月均 0 略去）；
      `_hybrid_safety_floor3` 為 True（計算層判定受下調補底 3 件影響）時加註保底原因。
    """
    cols = _HYBRID_SHOP_ISSUE_COLUMNS
    if raw is None or raw.empty:
        return pd.DataFrame(columns=cols)
    diag = raw.get("Safety Review", _safe_default_series("", raw.index)).fillna("").astype(str)
    mask = diag.isin([_HYBRID_DIAG_UNDER, _HYBRID_DIAG_OVER])
    zero_mask, _ = _hybrid_p90d_zero_sales_mask(raw, ref_date=ref_date)
    zero_mask = zero_mask.reindex(raw.index, fill_value=False).fillna(False).astype(bool)
    nd_only_mask = zero_mask & ~mask.fillna(False).to_numpy()
    combined = mask.fillna(False).to_numpy() | nd_only_mask.to_numpy()
    if not bool(combined.any()):
        return pd.DataFrame(columns=cols)
    work = _shop_issue_base_frame(raw, pd.Series(combined, index=raw.index))
    if "Supply Source" not in work.columns:
        work["Supply Source"] = raw.get("Supply Source", _safe_default_series("", work.index))
    if "D001 Net Stock Qty (CM)" not in work.columns:
        net_col = _resolve_column_name("D001 Net Stock Qty (CM)", set(raw.columns))
        if net_col is not None:
            work["D001 Net Stock Qty (CM)"] = raw[net_col]
    diag_w = work.get("Safety Review", _safe_default_series("", work.index)).fillna("").astype(str)
    days_w = pd.to_numeric(work.get(_HYBRID_SAFETY_DAYS_COLUMN, _safe_default_series(np.nan, work.index)), errors="coerce")
    line_w = pd.to_numeric(work.get(_HYBRID_SAFETY_LINE_COLUMN, _safe_default_series(np.nan, work.index)), errors="coerce")
    adj_w = pd.to_numeric(work.get("Adj. Safety Stock", _safe_default_series(np.nan, work.index)), errors="coerce")
    sugg_w = pd.to_numeric(work.get("建議 Safety 設定值", _safe_default_series(np.nan, work.index)), errors="coerce")
    sugg_td_w = pd.to_numeric(
        work.get(_HYBRID_SUGGEST_TURNDAY_COLUMN, _safe_default_series(np.nan, work.index)), errors="coerce"
    )
    floor_w = (
        work.get("_hybrid_safety_floor3", _safe_default_series(False, work.index))
        .fillna(False)
        .astype(bool)
    )
    chg_w = pd.to_numeric(work.get("Safety 建議調整量", _safe_default_series(np.nan, work.index)), errors="coerce")
    evt_w = work.get("事件類型", _safe_default_series("", work.index)).fillna("").astype(str)
    gap_w = _safe_numeric(work.get("調入缺口", _safe_default_series(0, work.index))).astype(int)
    actual_w = _safe_numeric(work.get("實際調入量", _safe_default_series(0, work.index))).astype(int)
    demand_w = _safe_numeric(work.get("調入需求量 (補到14日)", _safe_default_series(0, work.index))).astype(int)
    out_w = _safe_numeric(work.get("轉出量", _safe_default_series(0, work.index))).astype(int)
    src_w = work.get("調貨來源 (Site Code)", _safe_default_series("", work.index)).fillna("").astype(str).str.strip()
    d001_qty_w = _d001_available_qty(raw).loc[work.index]
    warehouse_w = _is_warehouse_supply(raw).loc[work.index]
    zero_w = zero_mask.reindex(work.index, fill_value=False).fillna(False).astype(bool)
    stock_cm_w = _safe_numeric(work.get("Shop Stock Qty (CM)", _safe_default_series(0, work.index))).astype(int)

    issue, desc, advice, sev = [], [], [], []
    for i in range(len(work)):
        days_txt = f"{float(days_w.iloc[i]):g} 日" if pd.notna(days_w.iloc[i]) else "未能確認"
        line_txt = f"{float(line_w.iloc[i]):g}" if pd.notna(line_w.iloc[i]) else "未能確認"
        adj_txt = f"{float(adj_w.iloc[i]):g}" if pd.notna(adj_w.iloc[i]) else "未能確認"
        sugg_txt = f"{float(sugg_w.iloc[i]):g}" if pd.notna(sugg_w.iloc[i]) else "未能確認"
        delta_txt = _safety_delta_text(chg_w.iloc[i])
        td_clause = _suggest_turnday_clause(sugg_td_w.iloc[i])
        evt_txt = str(evt_w.iloc[i]).strip()
        is_zero = bool(zero_w.iloc[i])
        if diag_w.iloc[i] == _HYBRID_DIAG_UNDER:
            issue.append("Safety 設定不足")
            desc.append(f"現時 Safety {adj_txt} 件低於周轉所需 {line_txt} 件（基準 {days_txt}；事件：{evt_txt}）")
            text = f"建議將 Safety 由 {adj_txt} 件調整為 {sugg_txt} 件（{delta_txt}{td_clause}）"
            if int(demand_w.iloc[i]) > 0:
                src_n = _count_transfer_sources(str(src_w.iloc[i]))
                text += f"；另需調入 {int(demand_w.iloc[i])} 件（實際可調入 {int(actual_w.iloc[i])} 件，涉及 {src_n} 間來源店舖）"
            if int(gap_w.iloc[i]) > 0:
                text += (f"；尚有缺口 {int(gap_w.iloc[i])} 件，"
                         f"{_d001_gap_advice(float(d001_qty_w.iloc[i]), int(gap_w.iloc[i]), bool(warehouse_w.iloc[i]))}")
            if is_zero:
                text += f"；{_HYBRID_P90D_ZERO_SALES_ANNOTATION}"
            advice.append(text)
            sev.append(_SHOP_ISSUE_SEVERITY["Safety 設定不足"])
        elif diag_w.iloc[i] == _HYBRID_DIAG_OVER:
            issue.append("Safety 設定過高")
            desc.append(f"現時 Safety {adj_txt} 件高於周轉所需 {line_txt} 件（基準 {days_txt}；事件：{evt_txt}）")
            text = (
                f"建議將 Safety 由 {adj_txt} 件調整為 {sugg_txt} 件（{delta_txt}{td_clause}），"
                f"多餘存貨可轉出（實際轉出 {int(out_w.iloc[i])} 件），詳見「調貨來源 (Site Code)」欄"
            )
            if is_zero:
                text += f"；{_HYBRID_P90D_ZERO_SALES_ANNOTATION}"
            # 計算層已判定受下調補底影響（`_hybrid_safety_floor3`）：加註最低建議 3 件
            # 為保障自動補貨的保底設定（SAP MOQ）
            if bool(floor_w.iloc[i]):
                text += _SAFETY_FLOOR3_HYBRID_NOTE
            advice.append(text)
            sev.append(_SHOP_ISSUE_SEVERITY["Safety 設定過高"])
        else:
            stock_txt = f"{int(stock_cm_w.iloc[i])}"
            issue.append("ND 退場")
            desc.append("過去 90 日（M-2／M-1／CM-MTD／P90D）皆零銷售，店舖仍有庫存，建議檢視是否轉 ND 停補貨")
            text = f"建議轉為 ND，現有庫存全數轉出或退倉（現時 Safety {adj_txt} 件；店舖庫存 {stock_txt} 件）"
            if int(gap_w.iloc[i]) > 0:
                text += (f"；尚有缺口 {int(gap_w.iloc[i])} 件，"
                         f"{_d001_gap_advice(float(d001_qty_w.iloc[i]), int(gap_w.iloc[i]), bool(warehouse_w.iloc[i]))}")
            advice.append(text)
            sev.append(_SHOP_ISSUE_SEVERITY["ND 退場"])
    work["問題類型"] = issue
    work["問題說明"] = desc
    work["建議方法"] = advice
    work["_sev"] = sev
    avail_w = _safe_numeric(work.get("可轉出貨量 (Mix)", _safe_default_series(0, work.index)))
    avail_nt = _safe_numeric(work.get("可轉出貨量 (Non Tourist)", _safe_default_series(0, work.index)))
    avail_t = _safe_numeric(work.get("可轉出貨量 (Tourist)", _safe_default_series(0, work.index)))
    work["可轉出貨量"] = (avail_w + avail_nt + avail_t).astype(int)
    if "Brand" not in work.columns:
        work["Brand"] = ""
    if "SKU" not in work.columns:
        work["SKU"] = ""
    work = work.sort_values(["Site Code", "_sev", "Brand", "SKU"], kind="stable").drop(columns=["_sev"])
    return work[[c for c in cols if c in work.columns]].reset_index(drop=True)



def _write_shop_issue_sheet(ws, issue_df, *, title, progress_callback=None, progress_start=0.0, progress_end=0.0):
    """店舖問題 SKU 工作表：店舖×SKU 逐列（僅含需行動列），問題類型置中、說明／建議靠左換行。"""
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    band_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    _text_keys = ["Brand", "SKU Desc", "Buyer", "問題說明", "建議方法", "調貨來源 (Site Code)"]

    if issue_df is None or issue_df.empty:
        ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12, color="1E40AF")
        ws.cell(row=3, column=1, value="無符合資料").font = Font(size=12)
        _auto_width(ws)
        return
    headers = list(issue_df.columns)
    header_labels = [_display_safety_header(h) for h in headers]
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12, color="1E40AF")
    for col_idx, h in enumerate(header_labels, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border
    ws.row_dimensions[3].height = 30
    total = len(issue_df)
    # P0：itertuples＋set 查詢取代 iterrows＋list 查詢；值與格式語意不變
    _text_set = set(_text_keys)
    for r_idx, _vals in enumerate(_iter_header_rows(issue_df, headers)):
        excel_row = r_idx + 4
        is_band = (r_idx % 2 == 1)
        for col_idx, h in enumerate(headers, 1):
            val = _vals[col_idx - 1]
            cell = ws.cell(row=excel_row, column=col_idx, value=None if pd.isna(val) else val)
            cell.border = thin_border
            cell.alignment = left_wrap if h in _text_set else center
            if isinstance(val, (int, float)) and not isinstance(val, bool) and h not in _text_set:
                cell.number_format = "#,##0"
            if is_band:
                cell.fill = band_fill
        if progress_callback and (r_idx % 200 == 0 or r_idx == total - 1):
            p = progress_start + (progress_end - progress_start) * ((r_idx + 1) / max(total, 1))
            progress_callback(p, f"寫入店舖問題 SKU ({r_idx+1}/{total})...")
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{3 + total}"
    _auto_width(ws)

