# -*- coding: utf-8 -*-
"""P90D 0 Sales Review(Shop回覆)：按店舖統計零銷售積壓 SKU，分 MAKE UP／Non-MAKE UP。

只引用 modes.common，不引用其他模式模組，避免循環引用。

篩選：``Shop Stock Status`` 去除前後空格後等於 ``Overstock (P90D 0 Sales)``
（狀態值見 ``_OP90D_STATUS``）且 ``Shop Stock Qty (LM)`` 大於 0 的列；
LM 為 0（或空白視為 0）者不計算。
MAKE UP 判定：``Adj. MCH1`` 去除前後空格後等於 ``02 - MAKE UP`` 者為 MAKE UP，
其餘（含空白）一律為 Non-MAKE UP。

輸出六個工作表（不輸出 RAW，所有欄位均不含 Value／COGS／GP）：
1. By Shop — 各店舖 P90D 0 Sales SKU 總數（已刪除重複）、庫存總數及各分類 SKU 數與庫存總數；
   分類維度可切換：MAKE UP 分類、Item Class、Adj. MCH1、交叉維度（MAKE UP × Item Class、
   MAKE UP × Adj. MCH1：每個店舖／分組按 MAKE UP／Non-MAKE UP 逐列，欄為次分類分項與小計）
2. 店舖90日零銷售Top 10 SKU回覆 — 各店舖按 Shop Stock Qty (CM) 由大至小取前 10 列
   （預設只限 Non-MAKE UP；勾選包括 MAKE UP 時則一併計算），附店內庫存排名及空白「店舖回覆」欄
3. By Nature — 按 Shop Nature 分佈（分組欄隨分類維度變動）
4. By OM — 按 OM 分佈（分組欄隨分類維度變動）
5. By SKU — 按 SKU 統計出現 P90D 0 Sales 的店舖數（已刪除重複），按出現店舖數由大至小排序
6. By Buyer — 按 Buyer 統計出現 P90D 0 Sales 的店舖數（已刪除重複），按出現店舖數由大至小排序

個別店舖版（P90D 0 Sales Review(個別店舖)）：篩選準則相同，分類固定以 MAKE UP 分類；
明細工作表為「店舖90日零銷售ALL SKU」，列出各店舖全部 SKU（包括 MAKE UP 與
Non-MAKE UP，不設列數上限）且不設「店舖回覆」欄，
供個別店舖查閱自身全部零銷售 SKU 情況。
"""

import gc
import re
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


from modes.common import (
    _auto_width,
    _cached_shop_names,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _display_safety_header,
    _extract_shop_name,
    _iter_header_rows,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    _cached_safe_sku_series,
    validate_input,
)

__all__ = [
    "transform_overstock_p90d",
    "generate_excel_overstock_p90d",
    "transform_overstock_p90d_shop",
    "generate_excel_overstock_p90d_shop",
    "_build_op90d_by_shop",
    "_build_op90d_top10",
    "_build_op90d_top10_all",
    "_build_op90d_all_sku",
    "_build_op90d_by_nature",
    "_build_op90d_by_om",
    "_build_op90d_by_sku",
    "_build_op90d_top_no_sales_sku",
    "_build_op90d_by_buyer",
    "_OP90D_STATUS",
    "_OP90D_MAKEUP_MCH1",
    "_OP90D_BREAKDOWN_MAKEUP",
    "_OP90D_BREAKDOWN_ITEM_CLASS",
    "_OP90D_BREAKDOWN_MCH1",
    "_OP90D_BREAKDOWNS",
    "_OP90D_BREAKDOWN_LABELS",
    "_OP90D_BREAKDOWN_MAKEUP_ITEM",
    "_OP90D_BREAKDOWN_MAKEUP_MCH1",
    "_OP90D_ALL_PRIMARIES",
    "_OP90D_BY_SHOP_COLUMNS",
    "_OP90D_TOP10_COLUMNS",
    "_OP90D_TOP10_ALL_SHEET",
    "_OP90D_TOP10_ALL_COLUMNS",
    "_OP90D_ALL_SKU_SHEET",
    "_OP90D_ALL_SKU_COLUMNS",
    "_OP90D_BY_NATURE_COLUMNS",
    "_OP90D_BY_OM_COLUMNS",
    "_OP90D_BY_SKU_COLUMNS",
    "_OP90D_TOP_NO_SALES_SHEET",
    "_OP90D_TOP_NO_SALES_COLUMNS",
    "_OP90D_TOP_NO_SALES_DEFAULT",
    "_OP90D_BY_BUYER_COLUMNS",
]

_OP90D_STATUS = "Overstock (P90D 0 Sales)"
_OP90D_MAKEUP_MCH1 = "02 - MAKE UP"

# 店舖總數（分母）暫存於 raw.attrs，供 By SKU 計算無銷售店舖佔比。
# 母體為清理及 Shop Nature 篩選後的上載資料，而非只有 P90D 0 Sales 的列。
_OP90D_TOTAL_SHOPS_ATTR = "_op90d_total_shops"

_OP90D_MAKEUP_LABEL = "MAKE UP"
_OP90D_NON_MAKEUP_LABEL = "Non-MAKE UP"

# 分類維度：決定 By Shop、By Nature、By OM 的分組欄位（Top 10 不受影響）
_OP90D_BREAKDOWN_MAKEUP = "makeup"
_OP90D_BREAKDOWN_ITEM_CLASS = "item_class"
_OP90D_BREAKDOWN_MCH1 = "mch1"
# 交叉維度：以 MAKE UP／Non-MAKE UP 為主分類逐列交叉單一次分類（Item Class 或 Adj. MCH1）
_OP90D_BREAKDOWN_MAKEUP_ITEM = "makeup_item"
_OP90D_BREAKDOWN_MAKEUP_MCH1 = "makeup_mch1"

_OP90D_BREAKDOWNS = (
    _OP90D_BREAKDOWN_MAKEUP, _OP90D_BREAKDOWN_ITEM_CLASS, _OP90D_BREAKDOWN_MCH1,
    _OP90D_BREAKDOWN_MAKEUP_ITEM, _OP90D_BREAKDOWN_MAKEUP_MCH1,
)

_OP90D_BREAKDOWN_LABELS = {
    _OP90D_BREAKDOWN_MAKEUP: "MAKE UP 分類",
    _OP90D_BREAKDOWN_ITEM_CLASS: "Item Class",
    _OP90D_BREAKDOWN_MCH1: "Adj. MCH1",
    _OP90D_BREAKDOWN_MAKEUP_ITEM: "MAKE UP × Item Class",
    _OP90D_BREAKDOWN_MAKEUP_MCH1: "MAKE UP × Adj. MCH1",
}

# 交叉維度逐列使用的主分類次序（MAKE UP 在前）
_OP90D_ALL_PRIMARIES = (_OP90D_MAKEUP_LABEL, _OP90D_NON_MAKEUP_LABEL)

# Item Class 固定顯示順序（空白組置於最後）
_OP90D_ITEM_CLASS_ORDER = ["A1", "A2", "A3", "B", "C", "D"]

_OP90D_BY_SHOP_COLUMNS = [
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "P90D 0 Sales SKU 總數",
    "Shop Stock Qty (CM) 總數",
    "MAKE UP SKU 數",
    "MAKE UP Shop Stock Qty (CM) 總數",
    "Non-MAKE UP SKU 數",
    "Non-MAKE UP Shop Stock Qty (CM) 總數",
]

# By Shop 固定首尾欄（中間為隨分類維度變動的分組欄）
_OP90D_BY_SHOP_BASE = [
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "P90D 0 Sales SKU 總數",
]
_OP90D_BY_SHOP_TAIL = ["Shop Stock Qty (CM) 總數"]

_OP90D_TOP10_COLUMNS = [
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "店內庫存排名",
    "Buyer",
    "Brand",
    "SKU",
    "SKU Desc",
    "Adj. MCH1",
    "MAKE UP 分類",
    "RP Type",
    "Adj. Safety Stock",
    "Item Status",
    "Supply Source",
    "Shop Stock Qty (LM)",
    "Shop Stock Qty (CM)",
    "Sold Qty (P90D)",
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Shop On Order PO Qty (CM)",
    "店舖回覆",
]

# 個別店舖版 ALL SKU 明細：包含 MAKE UP 與 Non-MAKE UP 全部 SKU，不設「店舖回覆」欄
_OP90D_ALL_SKU_SHEET = "店舖90日零銷售ALL SKU"
_OP90D_ALL_SKU_COLUMNS = [c for c in _OP90D_TOP10_COLUMNS if c != "店舖回覆"]
# 相容別名（舊名含 Top 10，現已改為全部 SKU）
_OP90D_TOP10_ALL_SHEET = _OP90D_ALL_SKU_SHEET
_OP90D_TOP10_ALL_COLUMNS = _OP90D_ALL_SKU_COLUMNS

_OP90D_BY_NATURE_COLUMNS = [
    "Shop Nature",
    "涉及店舖數",
    "P90D 0 Sales SKU 數",
    "P90D 0 Sales 行數",
    "MAKE UP SKU 數",
    "Non-MAKE UP SKU 數",
    "Shop Stock Qty (CM) 總數",
]

_OP90D_BY_OM_COLUMNS = [
    "OM",
    "涉及店舖數",
    "P90D 0 Sales SKU 數",
    "P90D 0 Sales 行數",
    "MAKE UP SKU 數",
    "Non-MAKE UP SKU 數",
    "Shop Stock Qty (CM) 總數",
]

_OP90D_BY_SKU_COLUMNS = [
    "排名",
    "SKU",
    "SKU Desc",
    "Brand",
    "Buyer",
    "MAKE UP 分類",
    "Adj. MCH1",
    "Item Class",
    "Item Status",
    "Launch Date",
    "出現店舖數",
    "店舖總數 (分母)",
    "無銷售店舖佔比",
    "P90D 0 Sales 行數",
    "Shop Stock Qty (CM) 總數",
]

# 「最多店舖無銷售 SKU」摘要表：按無銷售店舖數由大至小取前 N 個 SKU。
# 完整清單仍保留於 By SKU 工作表。
_OP90D_TOP_NO_SALES_SHEET = "最多店舖無銷售SKU"
_OP90D_TOP_NO_SALES_DEFAULT = 20
_OP90D_TOP_NO_SALES_COLUMNS = [
    "排名",
    "SKU",
    "SKU Desc",
    "Brand",
    "Buyer",
    "MAKE UP 分類",
    "Adj. MCH1",
    "Item Class",
    "Item Status",
    "Launch Date",
    "無銷售店舖數",
    "店舖總數 (分母)",
    "無銷售店舖佔比",
    "Shop Stock Qty (CM) 總數",
]

_OP90D_BY_BUYER_COLUMNS = [
    "Buyer",
    "出現店舖數",
    "P90D 0 Sales SKU 數",
    "P90D 0 Sales 行數",
    "Shop Stock Qty (CM) 總數",
]

# 以文字形式輸出的識別欄：Excel 內不套用千分位數值格式
# Launch Date 亦列入——來源常為 YYYYMMDD 數字（如 20260501），套用千分位會誤顯示為 20,260,501
_OP90D_TEXT_COLUMNS = frozenset({
    "Site Code", "Shop Name", "OM", "Shop Nature", "Buyer", "Brand", "SKU",
    "SKU Desc", "Adj. MCH1", "MAKE UP 分類", "RP Type", "Item Status",
    "Supply Source", "店舖回覆", "Item Class", "Launch Date",
})

# 以 0-1 小數儲存、Excel 以百分比顯示的欄位
_OP90D_PCT_COLUMNS = frozenset({"無銷售店舖佔比"})


def _is_banned_column(name: object) -> bool:
    """含 Value／COGS／GP 的欄位一律不得輸出（本模式固定白名單已排除，此為防線）。"""
    if not isinstance(name, str):
        return False
    lowered = name.lower()
    if "value" in lowered or "cogs" in lowered:
        return True
    return re.search(r"\bgp\b", lowered) is not None


def _classify_makeup(mch1: pd.Series) -> pd.Series:
    """Adj. MCH1 → MAKE UP 分類（去除前後空格後完全相符者為 MAKE UP）。"""
    txt = mch1.fillna("").astype(str).str.strip()
    return pd.Series(
        [_OP90D_MAKEUP_LABEL if v == _OP90D_MAKEUP_MCH1 else _OP90D_NON_MAKEUP_LABEL for v in txt],
        index=mch1.index,
    )


def _display_group(value: object) -> str:
    """分佈表分組顯示：空白（含缺值）顯示為「空白」。"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "空白"
    txt = str(value).strip()
    return txt if txt else "空白"


def _representative_nature(group: pd.DataFrame) -> str:
    """同一店舖出現多個 Shop Nature（資料異常）時，取出現次數最多者為代表值。"""
    vals = group.get("Shop Nature", _safe_default_series("", group.index)).fillna("").astype(str).str.strip()
    vals = vals[vals != ""]
    if vals.empty:
        return ""
    return str(vals.value_counts().index[0])


def _prepare_filtered(df_input: pd.DataFrame, shop_natures: list[str] | None = None,
                      progress_callback=None) -> pd.DataFrame:
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必要欄位: {', '.join(missing)}")
    avail = set(df_input.columns)
    for col in ("Site Code", "Shop Stock Qty (CM)", "Shop Stock Qty (LM)", "Adj. MCH1"):
        if _resolve_column_name(col, avail) is None:
            raise ValueError(f"缺少必需欄位: {col}")

    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback)
    # 分母：清理及 Shop Nature 篩選後的店舖總數（非只有 P90D 0 Sales 的店舖）
    _site_all = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    _total_shops = int(_site_all.replace("", pd.NA).dropna().nunique())
    if df.empty:
        df.attrs[_OP90D_TOTAL_SHOPS_ATTR] = _total_shops
        return df
    df = _compute_raw_extra_columns(df)
    df["SKU"] = df["SKU"].apply(_safe_sku)
    status = df.get("Shop Stock Status", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    stock_lm = _safe_numeric(df.get("Shop Stock Qty (LM)", _safe_default_series(0, df.index)))
    raw = df[(status == _OP90D_STATUS) & (stock_lm > 0)].copy()
    raw.attrs[_OP90D_TOTAL_SHOPS_ATTR] = _total_shops
    if raw.empty:
        return raw
    raw["MAKE UP 分類"] = _classify_makeup(raw["Adj. MCH1"])
    raw["_is_makeup"] = (raw["MAKE UP 分類"] == _OP90D_MAKEUP_LABEL).astype(int)
    raw["_stock_cm"] = _safe_numeric(raw.get("Shop Stock Qty (CM)", _safe_default_series(0, raw.index)))
    return raw


def _op90d_group_series(df: pd.DataFrame, breakdown: str) -> pd.Series:
    """逐列分組標籤：makeup 沿用 MAKE UP 分類；其餘以欄位值表示（空白歸入「空白」組）。"""
    if breakdown == _OP90D_BREAKDOWN_MCH1:
        col = df.get("Adj. MCH1", _safe_default_series("", df.index))
    elif breakdown == _OP90D_BREAKDOWN_ITEM_CLASS:
        col = df.get("Item Class", _safe_default_series("", df.index))
    else:
        return df.get(
            "MAKE UP 分類",
            _safe_default_series(_OP90D_NON_MAKEUP_LABEL, df.index),
        ).fillna(_OP90D_NON_MAKEUP_LABEL).astype(str)
    return col.map(_display_group)


def _op90d_ordered_groups(groups: pd.Series, breakdown: str) -> list[str]:
    """分組欄顯示順序：makeup 固定兩組；Item Class 按 A1→D；MCH1 按編號排序，空白組置於最後。"""
    present = [g for g in list(pd.unique(groups.fillna("").astype(str))) if g != ""]
    if breakdown == _OP90D_BREAKDOWN_ITEM_CLASS:
        ordered = [g for g in _OP90D_ITEM_CLASS_ORDER if g in present]
        ordered += sorted([g for g in present if g not in _OP90D_ITEM_CLASS_ORDER and g != "空白"])
        if "空白" in present:
            ordered.append("空白")
        return ordered
    if breakdown == _OP90D_BREAKDOWN_MCH1:
        ordered = sorted([g for g in present if g != "空白"])
        if "空白" in present:
            ordered.append("空白")
        return ordered
    return [_OP90D_MAKEUP_LABEL, _OP90D_NON_MAKEUP_LABEL]


def _op90d_group_header(breakdown: str, group: str) -> str:
    if breakdown == _OP90D_BREAKDOWN_MAKEUP:
        return "MAKE UP SKU 數" if group == _OP90D_MAKEUP_LABEL else "Non-MAKE UP SKU 數"
    return f"{group} SKU 數"


def _op90d_group_stock_header(group: str) -> str:
    """By Shop 各分類庫存欄：各分類 SKU 數欄之後，附同分類庫存總數欄。"""
    return f"{group} Shop Stock Qty (CM) 總數"


def _op90d_cross_sub(breakdown: str) -> str:
    """交叉維度對應的次分類（makeup_item → item_class；makeup_mch1 → mch1）。"""
    return _OP90D_BREAKDOWN_ITEM_CLASS if breakdown == _OP90D_BREAKDOWN_MAKEUP_ITEM else _OP90D_BREAKDOWN_MCH1


def _build_op90d_by_shop(raw: pd.DataFrame, breakdown: str = _OP90D_BREAKDOWN_MAKEUP) -> pd.DataFrame:
    if breakdown not in _OP90D_BREAKDOWNS:
        breakdown = _OP90D_BREAKDOWN_MAKEUP
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_OP90D_BY_SHOP_BASE + _OP90D_BY_SHOP_TAIL)
    if breakdown in (_OP90D_BREAKDOWN_MAKEUP_ITEM, _OP90D_BREAKDOWN_MAKEUP_MCH1):
        return _build_op90d_by_shop_cross(raw, _op90d_cross_sub(breakdown))
    df = raw.copy()
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    df["Site Code"] = site_code
    # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
    df["Shop Name"] = _cached_shop_names(item_shop, site_code)
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_group"] = _op90d_group_series(df, breakdown)
    groups = _op90d_ordered_groups(df["_group"], breakdown)
    group_headers = [_op90d_group_header(breakdown, g) for g in groups]
    group_stock_headers = [_op90d_group_stock_header(g) for g in groups]
    pair_headers = [h for pair in zip(group_headers, group_stock_headers) for h in pair]
    headers = _OP90D_BY_SHOP_BASE + _OP90D_BY_SHOP_TAIL + pair_headers

    rows = []
    # P0（個別模式）：單次 groupby 預彙總取代每店×每分類全表切片；不重複計數與合計口徑不變
    try:
        _pre = df.groupby(["Site Code", "Shop Name", "OM", "_group"], sort=False, observed=True).agg(
            _sku_n=("SKU", lambda s: int(s.dropna().astype(str).nunique())),
            _stock=("_stock_cm", "sum"),
        ).reset_index()
        _pre_map = {(str(r["Site Code"]), str(r["Shop Name"]), str(r["OM"]), str(r["_group"])): (int(r["_sku_n"]), float(r["_stock"])) for _, r in _pre.iterrows()}
        _shop_tot = df.groupby(["Site Code", "Shop Name", "OM"], sort=False, observed=True).agg(
            _sku_tot=("SKU", lambda s: int(s.dropna().astype(str).nunique())),
            _stock_tot=("_stock_cm", "sum"),
        )
        _shop_tot_map = {(str(s), str(n), str(o)): (int(r["_sku_tot"]), float(r["_stock_tot"])) for (s, n, o), r in _shop_tot.iterrows()}
        _use_pre = True
    except Exception:
        _pre_map = {}
        _shop_tot_map = {}
        _use_pre = False
    for (site, name, om), group in df.groupby(["Site Code", "Shop Name", "OM"], dropna=False, observed=True):
        if _use_pre:
            _key3 = (str(site), str(name), str(om))
            _tot_n, _tot_s = _shop_tot_map.get(_key3, (0, 0.0))
            row = {
                "Site Code": site,
                "Shop Name": name,
                "OM": om,
                "Shop Nature": _representative_nature(group),
                "P90D 0 Sales SKU 總數": int(_tot_n),
            }
            for g, h, sh in zip(groups, group_headers, group_stock_headers):
                _n, _s = _pre_map.get((str(site), str(name), str(om), str(g)), (0, 0.0))
                row[h] = int(_n)
                row[sh] = int(float(_s))
            row["Shop Stock Qty (CM) 總數"] = int(float(_tot_s))
        else:
            skus = group["SKU"].dropna().astype(str)
            row = {
                "Site Code": site,
                "Shop Name": name,
                "OM": om,
                "Shop Nature": _representative_nature(group),
                "P90D 0 Sales SKU 總數": int(skus.nunique()),
            }
            for g, h, sh in zip(groups, group_headers, group_stock_headers):
                gslice = group[group["_group"] == g]
                row[h] = int(gslice["SKU"].dropna().astype(str).nunique())
                row[sh] = int(float(gslice["_stock_cm"].sum()))
            row["Shop Stock Qty (CM) 總數"] = int(float(group["_stock_cm"].sum()))
        rows.append(row)
    by_shop = pd.DataFrame(rows, columns=headers)
    by_shop = by_shop.sort_values(
        ["P90D 0 Sales SKU 總數", "Shop Stock Qty (CM) 總數", "Site Code"],
        ascending=[False, False, True],
        kind="stable",
    ).reset_index(drop=True)
    if not by_shop.empty:
        total_row = {
            "Site Code": "總計",
            "Shop Name": "",
            "OM": "",
            "Shop Nature": "",
            "P90D 0 Sales SKU 總數": int(df["SKU"].dropna().astype(str).nunique()),
        }
        # P0（個別模式）：總計列沿用預彙總，避免每分類全表掃描；口徑不變
        try:
            _gtot = df.groupby("_group", sort=False, observed=True).agg(
                _sku_n=("SKU", lambda s: int(s.dropna().astype(str).nunique())),
                _stock=("_stock_cm", "sum"),
            )
            _gtot_map = {str(k): (int(r["_sku_n"]), float(r["_stock"])) for k, r in _gtot.iterrows()}
            for g, h, sh in zip(groups, group_headers, group_stock_headers):
                _n, _s = _gtot_map.get(str(g), (0, 0.0))
                total_row[h] = int(_n)
                total_row[sh] = int(float(_s))
        except Exception:
            for g, h, sh in zip(groups, group_headers, group_stock_headers):
                gslice = df[df["_group"] == g]
                total_row[h] = int(gslice["SKU"].dropna().astype(str).nunique())
                total_row[sh] = int(float(gslice["_stock_cm"].sum()))
        total_row["Shop Stock Qty (CM) 總數"] = int(float(df["_stock_cm"].sum()))
        by_shop = pd.concat([by_shop, pd.DataFrame([total_row])], ignore_index=True)
    return by_shop[headers]


_OP90D_CROSS_PRIMARY_COL = "MAKE UP 分類"
_OP90D_CROSS_SUBTOTAL_COL = "小計 SKU 數"


def _build_op90d_by_shop_cross(raw: pd.DataFrame, sub: str) -> pd.DataFrame:
    """交叉維度的 By Shop（樞紐式長表）：每個店舖按 MAKE UP／Non-MAKE UP 逐列，
    欄為次分類分項 SKU 數與小計（店舖無該主分類資料時不列該列）。"""
    df = raw.copy()
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    df["Site Code"] = site_code
    # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
    df["Shop Name"] = _cached_shop_names(item_shop, site_code)
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_pri"] = _op90d_group_series(df, _OP90D_BREAKDOWN_MAKEUP)
    df["_sec"] = _op90d_group_series(df, sub)
    sec_groups = _op90d_ordered_groups(df["_sec"], sub)
    sec_headers = [_op90d_group_header(sub, g) for g in sec_groups]
    base = ["Site Code", "Shop Name", "OM", "Shop Nature", _OP90D_CROSS_PRIMARY_COL, "P90D 0 Sales SKU 總數"]
    headers = base + sec_headers + [_OP90D_CROSS_SUBTOTAL_COL] + _OP90D_BY_SHOP_TAIL

    rows = []
    for (site, name, om), group in df.groupby(["Site Code", "Shop Name", "OM"], dropna=False, observed=True):
        shop_total = int(group["SKU"].dropna().astype(str).nunique())
        for pri_ord, pri in enumerate(_OP90D_ALL_PRIMARIES):
            psub = group[group["_pri"] == pri]
            if psub.empty:
                continue
            row = {
                "Site Code": site,
                "Shop Name": name,
                "OM": om,
                "Shop Nature": _representative_nature(group),
                _OP90D_CROSS_PRIMARY_COL: pri,
                "P90D 0 Sales SKU 總數": shop_total,
            }
            for g, h in zip(sec_groups, sec_headers):
                row[h] = int(psub[psub["_sec"] == g]["SKU"].dropna().astype(str).nunique())
            row[_OP90D_CROSS_SUBTOTAL_COL] = int(psub["SKU"].dropna().astype(str).nunique())
            row["Shop Stock Qty (CM) 總數"] = int(float(psub["_stock_cm"].sum()))
            row["_shop_total"] = shop_total
            row["_pri_ord"] = pri_ord
            rows.append(row)
    by_shop = pd.DataFrame(rows, columns=headers + ["_shop_total", "_pri_ord"])
    by_shop = by_shop.sort_values(
        ["_shop_total", "Site Code", "_pri_ord"], ascending=[False, True, True], kind="stable",
    ).reset_index(drop=True).drop(columns=["_shop_total", "_pri_ord"])
    if not by_shop.empty:
        shop_total_all = int(df["SKU"].dropna().astype(str).nunique())
        for pri in _OP90D_ALL_PRIMARIES:
            psub = df[df["_pri"] == pri]
            if psub.empty:
                continue
            total_row = {
                "Site Code": "總計",
                "Shop Name": "",
                "OM": "",
                "Shop Nature": "",
                _OP90D_CROSS_PRIMARY_COL: pri,
                "P90D 0 Sales SKU 總數": shop_total_all,
            }
            for g, h in zip(sec_groups, sec_headers):
                total_row[h] = int(psub[psub["_sec"] == g]["SKU"].dropna().astype(str).nunique())
            total_row[_OP90D_CROSS_SUBTOTAL_COL] = int(psub["SKU"].dropna().astype(str).nunique())
            total_row["Shop Stock Qty (CM) 總數"] = int(float(psub["_stock_cm"].sum()))
            by_shop = pd.concat([by_shop, pd.DataFrame([total_row])], ignore_index=True)
    return by_shop[headers]


def _build_op90d_top10(raw: pd.DataFrame, top_n: int = 10, include_makeup: bool = False) -> pd.DataFrame:
    """各店舖按 Shop Stock Qty (CM) 取前 N 列供店舖回覆。

    預設只限 Non-MAKE UP；include_makeup 為 True 時則包括 MAKE UP SKU。
    """
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_OP90D_TOP10_COLUMNS)
    df = raw.copy()
    if not include_makeup:
        df = df[pd.to_numeric(df.get("_is_makeup", 0), errors="coerce").fillna(0).astype(int) == 0]
    if df.empty:
        return pd.DataFrame(columns=_OP90D_TOP10_COLUMNS)
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    df["Site Code"] = site_code
    # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
    df["Shop Name"] = _cached_shop_names(item_shop, site_code)
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["Shop Nature"] = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df = df.sort_values(
        ["Site Code", "_stock_cm", "Brand", "SKU"],
        ascending=[True, False, True, True],
        kind="stable",
    )
    parts = []
    for _site, group in df.groupby("Site Code", dropna=False, observed=True):
        head = group.head(top_n).copy()
        head["店內庫存排名"] = range(1, len(head) + 1)
        parts.append(head)
    top10 = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=list(df.columns) + ["店內庫存排名"])
    top10["店舖回覆"] = ""
    available = set(top10.columns)
    cols = [c for c in _OP90D_TOP10_COLUMNS if _resolve_column_name(c, available) is not None]
    if "MAKE UP 分類" not in cols:
        cols.insert(cols.index("Adj. MCH1") + 1 if "Adj. MCH1" in cols else len(cols), "MAKE UP 分類")
    if "店內庫存排名" not in cols:
        cols.append("店內庫存排名")
    if "店舖回覆" not in cols:
        cols.append("店舖回覆")
    return top10[cols]


def _build_op90d_all_sku(raw: pd.DataFrame) -> pd.DataFrame:
    """個別店舖版 ALL SKU 明細：各店舖全部列按 Shop Stock Qty (CM) 由大至小排序。

    與店舖回覆版 Top 10 的差異：不設列數上限（輸出全部 SKU），固定包括 MAKE UP SKU，
    且不輸出「店舖回覆」欄，供店舖查閱自身全部零銷售 SKU 情況，無須回覆。
    """
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_OP90D_ALL_SKU_COLUMNS)
    df = raw.copy()
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    df["Site Code"] = site_code
    # P0：快取解析（同一店舖×品項只執行一次正則），輸出與逐列解析完全一致
    df["Shop Name"] = _cached_shop_names(item_shop, site_code)
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["Shop Nature"] = df.get("Shop Nature", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df = df.sort_values(
        ["Site Code", "_stock_cm", "Brand", "SKU"],
        ascending=[True, False, True, True],
        kind="stable",
    )
    parts = []
    for _site, group in df.groupby("Site Code", dropna=False, observed=True):
        ranked = group.copy()
        ranked["店內庫存排名"] = range(1, len(ranked) + 1)
        parts.append(ranked)
    all_sku = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=list(df.columns) + ["店內庫存排名"])
    available = set(all_sku.columns)
    cols = [c for c in _OP90D_ALL_SKU_COLUMNS if _resolve_column_name(c, available) is not None]
    if "MAKE UP 分類" not in cols:
        cols.insert(cols.index("Adj. MCH1") + 1 if "Adj. MCH1" in cols else len(cols), "MAKE UP 分類")
    if "店內庫存排名" not in cols:
        cols.append("店內庫存排名")
    return all_sku[cols]


def _build_op90d_top10_all(raw: pd.DataFrame, top_n: int | None = None) -> pd.DataFrame:
    """相容別名：現輸出全部列，``top_n`` 參數已停用而保留。"""
    return _build_op90d_all_sku(raw)


def _build_distribution(raw: pd.DataFrame, key: str, breakdown: str = _OP90D_BREAKDOWN_MAKEUP) -> pd.DataFrame:
    base = [key, "涉及店舖數", "P90D 0 Sales SKU 數", "P90D 0 Sales 行數"]
    tail = ["Shop Stock Qty (CM) 總數"]
    if breakdown not in _OP90D_BREAKDOWNS:
        breakdown = _OP90D_BREAKDOWN_MAKEUP
    if raw is None or raw.empty:
        return pd.DataFrame(columns=base + tail)
    if breakdown in (_OP90D_BREAKDOWN_MAKEUP_ITEM, _OP90D_BREAKDOWN_MAKEUP_MCH1):
        return _build_distribution_cross(raw, key, _op90d_cross_sub(breakdown))
    df = raw.copy()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_dist"] = df.get(key, _safe_default_series("", df.index)).map(_display_group)
    df["_group"] = _op90d_group_series(df, breakdown)
    groups = _op90d_ordered_groups(df["_group"], breakdown)
    group_headers = [_op90d_group_header(breakdown, g) for g in groups]
    columns = base + group_headers + tail
    rows = []
    for grp, sub in df.groupby("_dist", dropna=False, observed=True):
        row = {
            key: grp,
            "涉及店舖數": int(sub["Site Code"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
            "P90D 0 Sales SKU 數": int(sub["SKU"].dropna().astype(str).nunique()),
            "P90D 0 Sales 行數": int(len(sub)),
        }
        for g, h in zip(groups, group_headers):
            row[h] = int(sub[sub["_group"] == g]["SKU"].dropna().astype(str).nunique())
        row["Shop Stock Qty (CM) 總數"] = int(float(sub["_stock_cm"].sum()))
        rows.append(row)
    dist = pd.DataFrame(rows, columns=columns)
    dist = dist.sort_values(["P90D 0 Sales 行數", key], ascending=[False, True], kind="stable").reset_index(drop=True)
    if not dist.empty:
        total_row = {
            key: "總計",
            "涉及店舖數": int(df["Site Code"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
            "P90D 0 Sales SKU 數": int(df["SKU"].dropna().astype(str).nunique()),
            "P90D 0 Sales 行數": int(len(df)),
        }
        for g, h in zip(groups, group_headers):
            total_row[h] = int(df[df["_group"] == g]["SKU"].dropna().astype(str).nunique())
        total_row["Shop Stock Qty (CM) 總數"] = int(float(df["_stock_cm"].sum()))
        dist = pd.concat([dist, pd.DataFrame([total_row])], ignore_index=True)
    return dist[columns]


def _build_distribution_cross(raw: pd.DataFrame, key: str, sub: str) -> pd.DataFrame:
    """交叉維度的分佈表（樞紐式長表）：每個分組按 MAKE UP／Non-MAKE UP 逐列，
    欄為次分類分項 SKU 數與小計（分組無該主分類資料時不列該列）。"""
    tail = ["Shop Stock Qty (CM) 總數"]
    df = raw.copy()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_dist"] = df.get(key, _safe_default_series("", df.index)).map(_display_group)
    df["_pri"] = _op90d_group_series(df, _OP90D_BREAKDOWN_MAKEUP)
    df["_sec"] = _op90d_group_series(df, sub)
    sec_groups = _op90d_ordered_groups(df["_sec"], sub)
    sec_headers = [_op90d_group_header(sub, g) for g in sec_groups]
    base = [key, _OP90D_CROSS_PRIMARY_COL, "涉及店舖數", "P90D 0 Sales 行數"]
    columns = base + sec_headers + [_OP90D_CROSS_SUBTOTAL_COL] + tail

    rows = []
    for grp, sub_df in df.groupby("_dist", dropna=False, observed=True):
        for pri_ord, pri in enumerate(_OP90D_ALL_PRIMARIES):
            psub = sub_df[sub_df["_pri"] == pri]
            if psub.empty:
                continue
            row = {
                key: grp,
                _OP90D_CROSS_PRIMARY_COL: pri,
                "涉及店舖數": int(psub["Site Code"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
                "P90D 0 Sales 行數": int(len(psub)),
            }
            for g, h in zip(sec_groups, sec_headers):
                row[h] = int(psub[psub["_sec"] == g]["SKU"].dropna().astype(str).nunique())
            row[_OP90D_CROSS_SUBTOTAL_COL] = int(psub["SKU"].dropna().astype(str).nunique())
            row["Shop Stock Qty (CM) 總數"] = int(float(psub["_stock_cm"].sum()))
            row["_key_rows"] = int(len(sub_df))
            row["_pri_ord"] = pri_ord
            rows.append(row)
    dist = pd.DataFrame(rows, columns=columns + ["_key_rows", "_pri_ord"])
    dist = dist.sort_values(
        ["_key_rows", key, "_pri_ord"], ascending=[False, True, True], kind="stable",
    ).reset_index(drop=True).drop(columns=["_key_rows", "_pri_ord"])
    if not dist.empty:
        for pri in _OP90D_ALL_PRIMARIES:
            psub = df[df["_pri"] == pri]
            if psub.empty:
                continue
            total_row = {
                key: "總計",
                _OP90D_CROSS_PRIMARY_COL: pri,
                "涉及店舖數": int(psub["Site Code"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
                "P90D 0 Sales 行數": int(len(psub)),
            }
            for g, h in zip(sec_groups, sec_headers):
                total_row[h] = int(psub[psub["_sec"] == g]["SKU"].dropna().astype(str).nunique())
            total_row[_OP90D_CROSS_SUBTOTAL_COL] = int(psub["SKU"].dropna().astype(str).nunique())
            total_row["Shop Stock Qty (CM) 總數"] = int(float(psub["_stock_cm"].sum()))
            dist = pd.concat([dist, pd.DataFrame([total_row])], ignore_index=True)
    return dist[columns]


def _build_op90d_by_nature(raw: pd.DataFrame, breakdown: str = _OP90D_BREAKDOWN_MAKEUP) -> pd.DataFrame:
    return _build_distribution(raw, "Shop Nature", breakdown)


def _build_op90d_by_om(raw: pd.DataFrame, breakdown: str = _OP90D_BREAKDOWN_MAKEUP) -> pd.DataFrame:
    return _build_distribution(raw, "OM", breakdown)


def _op90d_representative(values: pd.Series) -> str:
    """同一 SKU 多列屬性不一致時，取出現次數最多者為代表值（空白不計入）。"""
    vals = values.fillna("").astype(str).str.strip()
    vals = vals[vals != ""]
    if vals.empty:
        return ""
    return str(vals.value_counts().index[0])


def _op90d_representative_value(values: pd.Series):
    """同一 SKU 多列屬性不一致時，取出現次數最多者的原始值（保留日期／數值類型，空白不計入）。"""
    vals = values[~values.isna()]
    if vals.empty:
        return ""
    stripped = vals.astype(str).str.strip()
    vals = vals[stripped != ""]
    if vals.empty:
        return ""
    return vals.value_counts().index[0]


def _op90d_total_shops(raw: pd.DataFrame, total_shops: int | None = None) -> int:
    """回傳店舖總數（分母）：優先使用呼叫端指定值，其次 raw.attrs，最後以 raw 內店舖數推斷。"""
    if total_shops is not None:
        try:
            val = int(total_shops)
            if val > 0:
                return val
        except (TypeError, ValueError):
            pass
    if raw is not None:
        try:
            attr = raw.attrs.get(_OP90D_TOTAL_SHOPS_ATTR)
            if attr is not None and int(attr) > 0:
                return int(attr)
        except (TypeError, ValueError, AttributeError):
            pass
        if not raw.empty:
            site = raw.get("Site Code", _safe_default_series("", raw.index))
            site = site.fillna("").astype(str).str.strip()
            return int(site.replace("", pd.NA).dropna().nunique())
    return 0


def _build_op90d_by_sku(raw: pd.DataFrame, total_shops: int | None = None) -> pd.DataFrame:
    """按 SKU 統計出現 P90D 0 Sales 的店舖數（Site Code 已刪除重複，空白不計）。

    每個 SKU 一列，按出現店舖數由大至小排序，店舖數相同時按庫存總數排序。
    另附「店舖總數 (分母)」與「無銷售店舖佔比」：分母為清理及 Shop Nature 篩選後的
    店舖總數（由 transform 暫存於 raw.attrs，亦可由呼叫端傳入 total_shops 覆寫）；
    佔比為 0-1 小數，Excel 以百分比顯示。
    屬性欄（SKU Desc／Brand／Buyer／MAKE UP 分類／Adj. MCH1／Item Class／Item Status）
    取該 SKU 出現次數最多者為代表值；Launch Date 取出現次數最多者的原始值
    （保留日期類型）；來源缺欄時顯示為空白。
    """
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_OP90D_BY_SKU_COLUMNS)
    df = raw.copy()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    site = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["Site Code"] = site
    if "_stock_cm" not in df.columns:
        df["_stock_cm"] = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, df.index)))
    _total = _op90d_total_shops(raw, total_shops)

    rows = []
    # P0（個別模式）：店舖數／行數／庫存預彙總取代每 SKU 重複字串清洗；代表值維持眾數語意不變
    try:
        _site_clean = site.replace("", pd.NA)
        _site_df = pd.DataFrame({"SKU": df["SKU"].to_numpy(), "Site": _site_clean.to_numpy()})
        _site_n_map = _site_df.groupby("SKU", sort=False, observed=True)["Site"].nunique().astype(int).to_dict()
        _cnt_map = df.groupby("SKU", sort=False, observed=True).size().astype(int).to_dict()
        _stock_map = df.groupby("SKU", sort=False, observed=True)["_stock_cm"].sum().to_dict()
        _use_pre2 = True
    except Exception:
        _site_n_map = {}
        _cnt_map = {}
        _stock_map = {}
        _use_pre2 = False
    for sku, group in df.groupby("SKU", dropna=False, observed=True):
        sku_txt = "" if pd.isna(sku) else str(sku).strip()
        if not sku_txt:
            continue
        if _use_pre2 and sku in _site_n_map:
            site_n = int(_site_n_map.get(sku, 0))
            _rows_n = int(_cnt_map.get(sku, len(group)))
            _stock_tot = float(_stock_map.get(sku, 0.0))
        else:
            site_n = int(
                group["Site Code"].replace("", pd.NA).dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()
            )
            _rows_n = int(len(group))
            _stock_tot = float(group["_stock_cm"].sum())
        rows.append({
            "SKU": sku_txt,
            "SKU Desc": _op90d_representative(group.get("SKU Desc", _safe_default_series("", group.index))),
            "Brand": _op90d_representative(group.get("Brand", _safe_default_series("", group.index))),
            "Buyer": _op90d_representative(group.get("Buyer", _safe_default_series("", group.index))),
            "MAKE UP 分類": _op90d_representative(
                group.get("MAKE UP 分類", _safe_default_series(_OP90D_NON_MAKEUP_LABEL, group.index))
            ),
            "Adj. MCH1": _op90d_representative(group.get("Adj. MCH1", _safe_default_series("", group.index))),
            "Item Class": _op90d_representative(group.get("Item Class", _safe_default_series("", group.index))),
            "Item Status": _op90d_representative(group.get("Item Status", _safe_default_series("", group.index))),
            "Launch Date": _op90d_representative_value(group.get("Launch Date", _safe_default_series("", group.index))),
            "出現店舖數": site_n,
            "店舖總數 (分母)": _total,
            "無銷售店舖佔比": (site_n / _total) if _total > 0 else None,
            "P90D 0 Sales 行數": _rows_n,
            "Shop Stock Qty (CM) 總數": int(_stock_tot),
        })
    by_sku = pd.DataFrame(rows, columns=_OP90D_BY_SKU_COLUMNS)
    if by_sku.empty:
        return by_sku
    by_sku = by_sku.sort_values(
        ["出現店舖數", "Shop Stock Qty (CM) 總數", "SKU"],
        ascending=[False, False, True],
        kind="stable",
    ).reset_index(drop=True)
    total_row = {
        "排名": None,
        "SKU": "總計",
        "SKU Desc": "",
        "Brand": "",
        "Buyer": "",
        "MAKE UP 分類": "",
        "Adj. MCH1": "",
        "Item Class": "",
        "Item Status": "",
        "Launch Date": "",
        "出現店舖數": int(df["Site Code"].replace("", pd.NA).dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
        "店舖總數 (分母)": _total,
        "無銷售店舖佔比": None,
        "P90D 0 Sales 行數": int(len(df)),
        "Shop Stock Qty (CM) 總數": int(float(df["_stock_cm"].sum())),
    }
    by_sku = pd.DataFrame(
        list(by_sku.to_dict("records")) + [total_row],
        columns=_OP90D_BY_SKU_COLUMNS,
    )
    # 排名為整數（可空），總計列留空；避免總計列的 None 令整欄變成小數
    by_sku["排名"] = pd.array(list(range(1, len(by_sku))) + [pd.NA], dtype="Int64")
    return by_sku[_OP90D_BY_SKU_COLUMNS]


def _build_op90d_top_no_sales_sku(
    raw: pd.DataFrame | None = None,
    top_n: int = _OP90D_TOP_NO_SALES_DEFAULT,
    total_shops: int | None = None,
    by_sku: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """最多店舖無銷售的 SKU 排名（取前 N 個）。

    以 By SKU 的「出現店舖數」作為無銷售店舖數，按同一排序取前 N 個；
    另附店舖總數（分母）與無銷售店舖佔比。完整清單仍保留於 By SKU 表。
    ``by_sku`` 可傳入已計算的 By SKU 表，避免重複運算。
    """
    if by_sku is None:
        by_sku = _build_op90d_by_sku(raw, total_shops=total_shops)
    if by_sku is None or by_sku.empty or "SKU" not in by_sku.columns:
        return pd.DataFrame(columns=_OP90D_TOP_NO_SALES_COLUMNS)
    df = by_sku[by_sku["SKU"].astype(str).str.strip() != "總計"].copy()
    if top_n is not None and int(top_n) > 0:
        df = df.head(int(top_n)).copy()
    df["排名"] = range(1, len(df) + 1)
    if "出現店舖數" in df.columns:
        df = df.rename(columns={"出現店舖數": "無銷售店舖數"})
    else:
        df["無銷售店舖數"] = 0
    for col in _OP90D_TOP_NO_SALES_COLUMNS:
        if col not in df.columns:
            df[col] = None
    return df[_OP90D_TOP_NO_SALES_COLUMNS].reset_index(drop=True)


def _build_op90d_by_buyer(raw: pd.DataFrame) -> pd.DataFrame:
    """按 Buyer 統計出現 P90D 0 Sales 的店舖數（Site Code 已刪除重複，空白不計）。

    每個 Buyer 一列，按出現店舖數由大至小排序，店舖數相同時按庫存總數排序。
    Buyer 空白列歸入「空白」組；SKU 數一律已刪除重複。
    """
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_OP90D_BY_BUYER_COLUMNS)
    df = raw.copy()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    buyer = df.get("Buyer", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["_buyer"] = buyer.map(_display_group)
    site = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["Site Code"] = site
    if "_stock_cm" not in df.columns:
        df["_stock_cm"] = _safe_numeric(df.get("Shop Stock Qty (CM)", _safe_default_series(0, df.index)))

    def _site_nunique(series: pd.Series) -> int:
        return int(series.replace("", pd.NA).dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique())

    rows = []
    for b, group in df.groupby("_buyer", dropna=False, observed=True):
        rows.append({
            "Buyer": b,
            "出現店舖數": _site_nunique(group["Site Code"]),
            "P90D 0 Sales SKU 數": int(group["SKU"].dropna().astype(str).nunique()),
            "P90D 0 Sales 行數": int(len(group)),
            "Shop Stock Qty (CM) 總數": int(float(group["_stock_cm"].sum())),
        })
    by_buyer = pd.DataFrame(rows, columns=_OP90D_BY_BUYER_COLUMNS)
    if by_buyer.empty:
        return by_buyer
    by_buyer = by_buyer.sort_values(
        ["出現店舖數", "Shop Stock Qty (CM) 總數", "Buyer"],
        ascending=[False, False, True],
        kind="stable",
    ).reset_index(drop=True)
    total_row = {
        "Buyer": "總計",
        "出現店舖數": _site_nunique(df["Site Code"]),
        "P90D 0 Sales SKU 數": int(df["SKU"].dropna().astype(str).nunique()),
        "P90D 0 Sales 行數": int(len(df)),
        "Shop Stock Qty (CM) 總數": int(float(df["_stock_cm"].sum())),
    }
    by_buyer = pd.concat([by_buyer, pd.DataFrame([total_row])], ignore_index=True)
    return by_buyer[_OP90D_BY_BUYER_COLUMNS]


def transform_overstock_p90d(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
    include_makeup: bool = False,
    breakdown: str = _OP90D_BREAKDOWN_MAKEUP,
):
    """篩選 P90D 0 Sales 列（狀態值見 _OP90D_STATUS，且 Shop Stock Qty (LM) 須大於 0），輸出店舖統計、Top 10、Nature 及 OM 分佈。

    Top 10 預設只限 Non-MAKE UP；include_makeup 為 True 時則包括 MAKE UP SKU。
    breakdown 決定 By Shop、By Nature、By OM 的分組欄位：
    makeup（MAKE UP 分類）、item_class（Item Class）、mch1（Adj. MCH1）、
    makeup_item／makeup_mch1（交叉維度：以 MAKE UP／Non-MAKE UP 為主分類逐列，
    欄為次分類分項 SKU 數與小計）。
    """
    if breakdown not in _OP90D_BREAKDOWNS:
        breakdown = _OP90D_BREAKDOWN_MAKEUP
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    raw = _prepare_filtered(df_input, shop_natures=shop_natures, progress_callback=progress_callback)
    if breakdown in (_OP90D_BREAKDOWN_ITEM_CLASS, _OP90D_BREAKDOWN_MAKEUP_ITEM) and not raw.empty:
        if _resolve_column_name("Item Class", set(raw.columns)) is None:
            raise ValueError("缺少必需欄位: Item Class（以 Item Class 分類時必須提供）")

    if progress_callback:
        progress_callback(0.35, "建立店舖統計...")
    by_shop = _build_op90d_by_shop(raw, breakdown)

    if progress_callback:
        progress_callback(0.5, "擷取各店最高庫存 SKU...")
    top10 = _build_op90d_top10(raw, include_makeup=include_makeup)

    if progress_callback:
        progress_callback(0.65, "建立 Nature 及 OM 分佈...")
    by_nature = _build_op90d_by_nature(raw, breakdown)
    by_om = _build_op90d_by_om(raw, breakdown)

    warnings = []
    if raw.empty:
        warnings.append("沒有符合「P90D 0 Sales（Shop Stock Qty (LM) 大於 0）」的資料。")
    elif breakdown in (_OP90D_BREAKDOWN_MAKEUP_ITEM, _OP90D_BREAKDOWN_MAKEUP_MCH1):
        sub = _op90d_cross_sub(breakdown)
        blank_count = int((_op90d_group_series(raw, sub) == "空白").sum())
        if blank_count:
            label = _OP90D_BREAKDOWN_LABELS[sub]
            warnings.append(f"「{label}」有 {blank_count} 列為空白，已歸入「空白」組。")
    elif breakdown != _OP90D_BREAKDOWN_MAKEUP:
        blank_count = int((_op90d_group_series(raw, breakdown) == "空白").sum())
        if blank_count:
            label = _OP90D_BREAKDOWN_LABELS[breakdown]
            warnings.append(f"「{label}」有 {blank_count} 列為空白，已歸入「空白」組。")

    if progress_callback:
        progress_callback(0.7, "完成資料轉換...")
    return raw, by_shop, top10, by_nature, by_om, warnings


def transform_overstock_p90d_shop(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
):
    """個別店舖版：篩選準則與店舖回覆版相同，分類固定以 MAKE UP 分類統計。

    ALL SKU 明細包括 MAKE UP 與 Non-MAKE UP 全部 SKU（不設列數上限），
    不輸出「店舖回覆」欄，供店舖查閱自身全部零銷售 SKU 情況，無須回覆。
    """
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    raw = _prepare_filtered(df_input, shop_natures=shop_natures, progress_callback=progress_callback)

    if progress_callback:
        progress_callback(0.35, "建立店舖統計...")
    by_shop = _build_op90d_by_shop(raw, _OP90D_BREAKDOWN_MAKEUP)

    if progress_callback:
        progress_callback(0.5, "建立各店全部 SKU 明細...")
    top10 = _build_op90d_all_sku(raw)

    if progress_callback:
        progress_callback(0.65, "建立 Nature 及 OM 分佈...")
    by_nature = _build_op90d_by_nature(raw, _OP90D_BREAKDOWN_MAKEUP)
    by_om = _build_op90d_by_om(raw, _OP90D_BREAKDOWN_MAKEUP)

    warnings = []
    if raw.empty:
        warnings.append("沒有符合「P90D 0 Sales（Shop Stock Qty (LM) 大於 0）」的資料。")

    if progress_callback:
        progress_callback(0.7, "完成資料轉換...")
    return raw, by_shop, top10, by_nature, by_om, warnings


def _write_op90d_table(ws, headers, data: pd.DataFrame, progress_callback=None,
                       progress_start=0.0, progress_end=0.1, sheet_name="",
                       auto_filter: bool = False):
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    display_headers = [_display_safety_header(h) for h in headers]
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for col_idx, h in enumerate(display_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
    # P0：itertuples＋共用對齊取代 iterrows＋逐格新建 Alignment；值與格式語意不變
    center_align = Alignment(horizontal="center", vertical="center")
    if data is not None and not data.empty:
        total = len(data)
        for row_idx, _vals in enumerate(_iter_header_rows(data, headers)):
            excel_row = row_idx + 2
            is_total = any(str(v) == "總計" for v in _vals[:3])
            for col_idx, h in enumerate(headers, 1):
                val = _vals[col_idx - 1]
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = center_align
                if is_total:
                    cell.font = bold_font
                if isinstance(val, (int, float)) and h not in _OP90D_TEXT_COLUMNS:
                    cell.number_format = "0.0%" if h in _OP90D_PCT_COLUMNS else "#,##0"
            if progress_callback and (row_idx % 500 == 0 or row_idx == total - 1):
                p = progress_start + (progress_end - progress_start) * ((row_idx + 1) / max(total, 1))
                progress_callback(p, f"寫入 {sheet_name} ({row_idx+1}/{total})...")
    else:
        ws.cell(row=2, column=1, value="無符合資料").font = Font(size=12)
    if auto_filter and ws.max_row > 1 and ws.max_column > 1:
        ws.auto_filter.ref = ws.dimensions
    _auto_width(ws)


def generate_excel_overstock_p90d(
    by_shop: pd.DataFrame,
    top10: pd.DataFrame,
    by_nature: pd.DataFrame,
    by_om: pd.DataFrame,
    progress_callback=None,
    include_raw: bool = True,
    by_sku: pd.DataFrame | None = None,
    by_buyer: pd.DataFrame | None = None,
    top_sheet_title: str | None = None,
    top_no_sales_sku: pd.DataFrame | None = None,
) -> BytesIO:
    """產生 Excel 工作表（無 RAW；include_raw 僅為介面相容而保留，不影響輸出）。

    ``by_sku``／``by_buyer`` 為 None 時略過對應工作表（保持舊有呼叫相容）；
    傳入統計表時分別另加「By SKU」及「By Buyer」工作表。
    ``top_no_sales_sku`` 傳入時另加「最多店舖無銷售SKU」摘要工作表（取前 N 個 SKU）。
    ``top_sheet_title`` 指定明細工作表名稱，預設為店舖回覆版名稱；
    個別店舖版傳入 ``_OP90D_ALL_SKU_SHEET``。
    """
    top_title = top_sheet_title or "店舖90日零銷售Top 10 SKU回覆"
    top_default_cols = (
        _OP90D_ALL_SKU_COLUMNS if top_title == _OP90D_ALL_SKU_SHEET
        else _OP90D_TOP10_COLUMNS
    )
    for name, frame in (("By Shop", by_shop), (top_title, top10), ("By Nature", by_nature), ("By OM", by_om),
                          ("By SKU", by_sku), ("By Buyer", by_buyer),
                          (_OP90D_TOP_NO_SALES_SHEET, top_no_sales_sku)):
        if frame is not None and not frame.empty:
            banned = [c for c in frame.columns if _is_banned_column(c)]
            if banned:
                raise ValueError(f"{name} 輸出含不應出現的欄位: {', '.join(banned)}")

    wb = Workbook()
    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    ws_shop = wb.active
    ws_shop.title = "By Shop"
    _write_op90d_table(ws_shop,
                       list(by_shop.columns) if by_shop is not None else _OP90D_BY_SHOP_COLUMNS,
                       by_shop, progress_callback, 0.0, 0.2, "By Shop")

    if progress_callback:
        progress_callback(0.24, f"建立{top_title}...")
    ws_top = wb.create_sheet(title=top_title)
    _write_op90d_table(ws_top,
                       list(top10.columns) if top10 is not None else top_default_cols,
                       top10, progress_callback, 0.24, 0.48, top_title, auto_filter=True)

    if top_no_sales_sku is not None:
        if progress_callback:
            progress_callback(0.52, f"建立{_OP90D_TOP_NO_SALES_SHEET}...")
        ws_top_no_sales = wb.create_sheet(title=_OP90D_TOP_NO_SALES_SHEET)
        _write_op90d_table(ws_top_no_sales,
                           list(top_no_sales_sku.columns) if top_no_sales_sku is not None else _OP90D_TOP_NO_SALES_COLUMNS,
                           top_no_sales_sku, progress_callback, 0.52, 0.58,
                           _OP90D_TOP_NO_SALES_SHEET, auto_filter=True)

    if progress_callback:
        progress_callback(0.6, "建立 By Nature...")
    ws_nature = wb.create_sheet(title="By Nature")
    _write_op90d_table(ws_nature,
                       list(by_nature.columns) if by_nature is not None else _OP90D_BY_NATURE_COLUMNS,
                       by_nature, progress_callback, 0.6, 0.7, "By Nature")

    if progress_callback:
        progress_callback(0.72, "建立 By OM...")
    ws_om = wb.create_sheet(title="By OM")
    _write_op90d_table(ws_om,
                       list(by_om.columns) if by_om is not None else _OP90D_BY_OM_COLUMNS,
                       by_om, progress_callback, 0.72, 0.8, "By OM")

    if by_sku is not None:
        if progress_callback:
            progress_callback(0.82, "建立 By SKU...")
        ws_sku = wb.create_sheet(title="By SKU")
        _write_op90d_table(ws_sku,
                           list(by_sku.columns) if by_sku is not None else _OP90D_BY_SKU_COLUMNS,
                           by_sku, progress_callback, 0.82, 0.88, "By SKU", auto_filter=True)

    if by_buyer is not None:
        if progress_callback:
            progress_callback(0.89, "建立 By Buyer...")
        ws_buyer = wb.create_sheet(title="By Buyer")
        _write_op90d_table(ws_buyer,
                           list(by_buyer.columns) if by_buyer is not None else _OP90D_BY_BUYER_COLUMNS,
                           by_buyer, progress_callback, 0.89, 0.94, "By Buyer")

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf


def generate_excel_overstock_p90d_shop(
    by_shop: pd.DataFrame,
    top10: pd.DataFrame,
    by_nature: pd.DataFrame,
    by_om: pd.DataFrame,
    progress_callback=None,
    include_raw: bool = True,
    by_sku: pd.DataFrame | None = None,
    by_buyer: pd.DataFrame | None = None,
) -> BytesIO:
    """個別店舖版 Excel：明細工作表為「店舖90日零銷售ALL SKU」。

    其餘五個工作表與店舖回覆版相同（無 RAW；include_raw 僅為介面相容而保留）。
    """
    return generate_excel_overstock_p90d(
        by_shop, top10, by_nature, by_om,
        progress_callback=progress_callback,
        include_raw=include_raw,
        by_sku=by_sku,
        by_buyer=by_buyer,
        top_sheet_title=_OP90D_ALL_SKU_SHEET,
    )
