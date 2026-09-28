# -*- coding: utf-8 -*-
"""P90D 0 Sales Review(Shop回覆)：按店舖統計零銷售積壓 SKU，分 MAKE UP／Non-MAKE UP。

只引用 modes.common，不引用其他模式模組，避免循環引用。

篩選：``Shop Stock Status`` 去除前後空格後等於 ``Overstock (P90D 0 Sales)``
（狀態值見 ``_OP90D_STATUS``）且 ``Shop Stock Qty (LM)`` 大於 0 的列；
LM 為 0（或空白視為 0）者不計算。
MAKE UP 判定：``Adj. MCH1`` 去除前後空格後等於 ``02 - MAKE UP`` 者為 MAKE UP，
其餘（含空白）一律為 Non-MAKE UP。

輸出四個工作表（不輸出 RAW，所有欄位均不含 Value／COGS／GP）：
1. By Shop — 各店舖 P90D 0 Sales SKU 總數（已刪除重複）及分類分項；
   分類維度可切換：MAKE UP 分類、Item Class、Adj. MCH1、全部（三類合併顯示：
   三節並列，Excel 表頭以合併儲存格分節標示並按節配色；Item Class 分項欄名加前綴）
2. Top 10 店舖回覆 — 各店舖按 Shop Stock Qty (CM) 由大至小取前 10 列
   （預設只限 Non-MAKE UP；勾選包括 MAKE UP 時則一併計算），附店內庫存排名及空白「店舖回覆」欄
3. By Nature — 按 Shop Nature 分佈（分組欄隨分類維度變動）
4. By OM — 按 OM 分佈（分組欄隨分類維度變動）
"""

import gc
import re
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


from modes.common import (
    _auto_width,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _display_safety_header,
    _extract_shop_name,
    _resolve_column_name,
    _safe_default_series,
    _safe_numeric,
    _safe_sku,
    validate_input,
)

__all__ = [
    "transform_overstock_p90d",
    "generate_excel_overstock_p90d",
    "_build_op90d_by_shop",
    "_build_op90d_top10",
    "_build_op90d_by_nature",
    "_build_op90d_by_om",
    "_OP90D_STATUS",
    "_OP90D_MAKEUP_MCH1",
    "_OP90D_BREAKDOWN_MAKEUP",
    "_OP90D_BREAKDOWN_ITEM_CLASS",
    "_OP90D_BREAKDOWN_MCH1",
    "_OP90D_BREAKDOWN_ALL",
    "_OP90D_ALL_SECTIONS",
    "_OP90D_BREAKDOWNS",
    "_OP90D_BREAKDOWN_LABELS",
    "_OP90D_BY_SHOP_COLUMNS",
    "_OP90D_TOP10_COLUMNS",
    "_OP90D_BY_NATURE_COLUMNS",
    "_OP90D_BY_OM_COLUMNS",
]

_OP90D_STATUS = "Overstock (P90D 0 Sales)"
_OP90D_MAKEUP_MCH1 = "02 - MAKE UP"

_OP90D_MAKEUP_LABEL = "MAKE UP"
_OP90D_NON_MAKEUP_LABEL = "Non-MAKE UP"

# 分類維度：決定 By Shop、By Nature、By OM 的分組欄位（Top 10 不受影響）
_OP90D_BREAKDOWN_MAKEUP = "makeup"
_OP90D_BREAKDOWN_ITEM_CLASS = "item_class"
_OP90D_BREAKDOWN_MCH1 = "mch1"
# 合併維度：三個分類以分節並列顯示（Excel 表頭以合併儲存格分節標示）
_OP90D_BREAKDOWN_ALL = "all"

_OP90D_BREAKDOWNS = (_OP90D_BREAKDOWN_MAKEUP, _OP90D_BREAKDOWN_ITEM_CLASS, _OP90D_BREAKDOWN_MCH1, _OP90D_BREAKDOWN_ALL)

_OP90D_BREAKDOWN_LABELS = {
    _OP90D_BREAKDOWN_MAKEUP: "MAKE UP 分類",
    _OP90D_BREAKDOWN_ITEM_CLASS: "Item Class",
    _OP90D_BREAKDOWN_MCH1: "Adj. MCH1",
    _OP90D_BREAKDOWN_ALL: "全部（三類合併顯示）",
}

# 合併維度的分節次序（By Shop／By Nature／By OM 的分組欄排列）
_OP90D_ALL_SECTIONS = (_OP90D_BREAKDOWN_MAKEUP, _OP90D_BREAKDOWN_ITEM_CLASS, _OP90D_BREAKDOWN_MCH1)

# Item Class 固定顯示順序（空白組置於最後）
_OP90D_ITEM_CLASS_ORDER = ["A1", "A2", "A3", "B", "C", "D"]

_OP90D_BY_SHOP_COLUMNS = [
    "Site Code",
    "Shop Name",
    "OM",
    "Shop Nature",
    "P90D 0 Sales SKU 總數",
    "MAKE UP SKU 數",
    "Non-MAKE UP SKU 數",
    "Shop Stock Qty (CM) 總數",
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
    "Item Status",
    "Supply Source",
    "Shop Stock Qty (LM)",
    "Shop Stock Qty (CM)",
    "Sold Qty (P90D)",
    "Sold Qty (M-2)",
    "Sold Qty (M-1)",
    "Sold Qty (CM - MTD)",
    "Shop On Order PO Qty (CM)",
    "Adj. Safety Stock",
    "店舖回覆",
]

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

# 以文字形式輸出的識別欄：Excel 內不套用千分位數值格式
_OP90D_TEXT_COLUMNS = frozenset({
    "Site Code", "Shop Name", "OM", "Shop Nature", "Buyer", "Brand", "SKU",
    "SKU Desc", "Adj. MCH1", "MAKE UP 分類", "RP Type", "Item Status",
    "Supply Source", "店舖回覆",
})

# 合併維度分節標題列的底色（白字粗體），三節各一色以便對讀
_OP90D_SECTION_FILLS = {
    "MAKE UP 分類": "C2255C",
    "Item Class": "0E7490",
    "Adj. MCH1": "15803D",
}


def _op90d_sheet_sections(columns, breakdown: str):
    """合併維度（all）的表頭分節：回傳（節標題, [欄名]）清單；其他維度回傳 None（單行表頭）。

    分項欄名可唯一辨識所屬分節：MAKE UP 兩欄固定；Item Class 分項帶「Item Class 」前綴；
    其餘「… SKU 數」欄為 Adj. MCH1 分項（前後基礎欄排除，避免「P90D 0 Sales SKU 數」誤判）。
    """
    if breakdown != _OP90D_BREAKDOWN_ALL:
        return None
    cols = list(columns)
    base_cols = {"涉及店舖數", "P90D 0 Sales SKU 數", "P90D 0 Sales 行數", "P90D 0 Sales SKU 總數"}
    makeup = [c for c in cols if c in ("MAKE UP SKU 數", "Non-MAKE UP SKU 數")]
    item = [c for c in cols if c.startswith("Item Class ") and c.endswith(" SKU 數")]
    mch1 = [
        c for c in cols
        if c.endswith(" SKU 數") and c not in base_cols and c not in makeup and c not in item
    ]
    sections = []
    if makeup:
        sections.append((_OP90D_BREAKDOWN_LABELS[_OP90D_BREAKDOWN_MAKEUP], makeup))
    if item:
        sections.append((_OP90D_BREAKDOWN_LABELS[_OP90D_BREAKDOWN_ITEM_CLASS], item))
    if mch1:
        sections.append((_OP90D_BREAKDOWN_LABELS[_OP90D_BREAKDOWN_MCH1], mch1))
    return sections or None


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
    if df.empty:
        return df
    df = _compute_raw_extra_columns(df)
    df["SKU"] = df["SKU"].apply(_safe_sku)
    status = df.get("Shop Stock Status", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    stock_lm = _safe_numeric(df.get("Shop Stock Qty (LM)", _safe_default_series(0, df.index)))
    raw = df[(status == _OP90D_STATUS) & (stock_lm > 0)].copy()
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


def _op90d_group_header(breakdown: str, group: str, merged: bool = False) -> str:
    """分組欄表頭；合併維度下 Item Class 分項加前綴，避免與 Adj. MCH1 的「空白 SKU 數」等欄名相撞。"""
    if breakdown == _OP90D_BREAKDOWN_MAKEUP:
        return "MAKE UP SKU 數" if group == _OP90D_MAKEUP_LABEL else "Non-MAKE UP SKU 數"
    if merged and breakdown == _OP90D_BREAKDOWN_ITEM_CLASS:
        return f"Item Class {group} SKU 數"
    return f"{group} SKU 數"


def _build_op90d_by_shop(raw: pd.DataFrame, breakdown: str = _OP90D_BREAKDOWN_MAKEUP) -> pd.DataFrame:
    if breakdown not in _OP90D_BREAKDOWNS:
        breakdown = _OP90D_BREAKDOWN_MAKEUP
    # 合併維度：三個分類依 _OP90D_ALL_SECTIONS 次序分節並列
    sections = _OP90D_ALL_SECTIONS if breakdown == _OP90D_BREAKDOWN_ALL else (breakdown,)
    merged = breakdown == _OP90D_BREAKDOWN_ALL
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_OP90D_BY_SHOP_BASE + _OP90D_BY_SHOP_TAIL)
    df = raw.copy()
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    df["Site Code"] = site_code
    df["Shop Name"] = [_extract_shop_name(item_shop.iloc[i], site_code.iloc[i]) for i in range(len(df))]
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    group_info = []
    for sub in sections:
        df[f"_group_{sub}"] = _op90d_group_series(df, sub)
        groups = _op90d_ordered_groups(df[f"_group_{sub}"], sub)
        group_info.append((sub, groups, [_op90d_group_header(sub, g, merged=merged) for g in groups]))
    headers = _OP90D_BY_SHOP_BASE + [h for _, _, hs in group_info for h in hs] + _OP90D_BY_SHOP_TAIL

    rows = []
    for (site, name, om), group in df.groupby(["Site Code", "Shop Name", "OM"], dropna=False, observed=True):
        skus = group["SKU"].dropna().astype(str)
        row = {
            "Site Code": site,
            "Shop Name": name,
            "OM": om,
            "Shop Nature": _representative_nature(group),
            "P90D 0 Sales SKU 總數": int(skus.nunique()),
        }
        for sub, groups, group_headers in group_info:
            for g, h in zip(groups, group_headers):
                row[h] = int(group[group[f"_group_{sub}"] == g]["SKU"].dropna().astype(str).nunique())
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
        for sub, groups, group_headers in group_info:
            for g, h in zip(groups, group_headers):
                total_row[h] = int(df[df[f"_group_{sub}"] == g]["SKU"].dropna().astype(str).nunique())
        total_row["Shop Stock Qty (CM) 總數"] = int(float(df["_stock_cm"].sum()))
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
    df["Shop Name"] = [_extract_shop_name(item_shop.iloc[i], site_code.iloc[i]) for i in range(len(df))]
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


def _build_distribution(raw: pd.DataFrame, key: str, breakdown: str = _OP90D_BREAKDOWN_MAKEUP) -> pd.DataFrame:
    base = [key, "涉及店舖數", "P90D 0 Sales SKU 數", "P90D 0 Sales 行數"]
    tail = ["Shop Stock Qty (CM) 總數"]
    if breakdown not in _OP90D_BREAKDOWNS:
        breakdown = _OP90D_BREAKDOWN_MAKEUP
    sections = _OP90D_ALL_SECTIONS if breakdown == _OP90D_BREAKDOWN_ALL else (breakdown,)
    merged = breakdown == _OP90D_BREAKDOWN_ALL
    if raw is None or raw.empty:
        return pd.DataFrame(columns=base + tail)
    df = raw.copy()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_dist"] = df.get(key, _safe_default_series("", df.index)).map(_display_group)
    group_info = []
    for sub in sections:
        df[f"_group_{sub}"] = _op90d_group_series(df, sub)
        groups = _op90d_ordered_groups(df[f"_group_{sub}"], sub)
        group_info.append((sub, groups, [_op90d_group_header(sub, g, merged=merged) for g in groups]))
    columns = base + [h for _, _, hs in group_info for h in hs] + tail
    rows = []
    for grp, sub in df.groupby("_dist", dropna=False, observed=True):
        row = {
            key: grp,
            "涉及店舖數": int(sub["Site Code"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
            "P90D 0 Sales SKU 數": int(sub["SKU"].dropna().astype(str).nunique()),
            "P90D 0 Sales 行數": int(len(sub)),
        }
        for sb, groups, group_headers in group_info:
            for g, h in zip(groups, group_headers):
                row[h] = int(sub[sub[f"_group_{sb}"] == g]["SKU"].dropna().astype(str).nunique())
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
        for sb, groups, group_headers in group_info:
            for g, h in zip(groups, group_headers):
                total_row[h] = int(df[df[f"_group_{sb}"] == g]["SKU"].dropna().astype(str).nunique())
        total_row["Shop Stock Qty (CM) 總數"] = int(float(df["_stock_cm"].sum()))
        dist = pd.concat([dist, pd.DataFrame([total_row])], ignore_index=True)
    return dist[columns]


def _build_op90d_by_nature(raw: pd.DataFrame, breakdown: str = _OP90D_BREAKDOWN_MAKEUP) -> pd.DataFrame:
    return _build_distribution(raw, "Shop Nature", breakdown)


def _build_op90d_by_om(raw: pd.DataFrame, breakdown: str = _OP90D_BREAKDOWN_MAKEUP) -> pd.DataFrame:
    return _build_distribution(raw, "OM", breakdown)


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
    all（三類合併顯示：MAKE UP 分類、Item Class、Adj. MCH1 分節並列）。
    """
    if breakdown not in _OP90D_BREAKDOWNS:
        breakdown = _OP90D_BREAKDOWN_MAKEUP
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    raw = _prepare_filtered(df_input, shop_natures=shop_natures, progress_callback=progress_callback)
    if breakdown == _OP90D_BREAKDOWN_ITEM_CLASS and not raw.empty:
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
    elif breakdown == _OP90D_BREAKDOWN_ALL:
        # 合併維度：逐個分類檢查空白（Item Class 欄整欄缺席時全部歸入「空白」組，藉此提示）
        for sub in (_OP90D_BREAKDOWN_ITEM_CLASS, _OP90D_BREAKDOWN_MCH1):
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


def _write_op90d_table(ws, headers, data: pd.DataFrame, progress_callback=None,
                       progress_start=0.0, progress_end=0.1, sheet_name="",
                       auto_filter: bool = False, sections=None):
    """寫出統計表；sections 為（節標題, [欄名]）清單時採兩行表頭：

    第一行為分節標題（水平合佔各節欄位，按節配色）與前後欄（上下合併兩行），
    第二行為分項表頭，資料由第三行開始。
    """
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    bold_font = Font(bold=True, size=11)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    display_headers = [_display_safety_header(h) for h in headers]

    def _style_header(cell, fill: PatternFill):
        cell.font = header_font
        cell.fill = fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # 分節欄位置（1 起始）；按連續段切分合併範圍，保證互不重疊；分節缺欄（空表）時退回單行表頭
    section_spans = []
    if sections:
        pos = {h: i + 1 for i, h in enumerate(headers)}
        for title, cols in sections:
            idxs = sorted(pos[c] for c in cols if c in pos)
            run_start = prev = None
            for i in idxs:
                if run_start is None:
                    run_start = prev = i
                elif i == prev + 1:
                    prev = i
                else:
                    section_spans.append((title, run_start, prev - run_start + 1))
                    run_start = prev = i
            if run_start is not None:
                section_spans.append((title, run_start, prev - run_start + 1))
    covered = {i for _t, start, span in section_spans for i in range(start, start + span)}

    if section_spans:
        for col_idx, h in enumerate(display_headers, 1):
            if col_idx in covered:
                cell = ws.cell(row=2, column=col_idx, value=h)
                _style_header(cell, header_fill)
            else:
                cell = ws.cell(row=1, column=col_idx, value=h)
                _style_header(cell, header_fill)
                ws.merge_cells(start_row=1, start_column=col_idx, end_row=2, end_column=col_idx)
        for title, start, span in section_spans:
            fill = PatternFill(
                start_color=_OP90D_SECTION_FILLS.get(title, "4472C4"),
                end_color=_OP90D_SECTION_FILLS.get(title, "4472C4"),
                fill_type="solid",
            )
            if span > 1:
                ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=start + span - 1)
            cell = ws.cell(row=1, column=start, value=title)
            _style_header(cell, fill)
            for c in range(start + 1, start + span):
                _style_header(ws.cell(row=1, column=c), fill)
        first_data_row = 3
    else:
        for col_idx, h in enumerate(display_headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            _style_header(cell, header_fill)
        first_data_row = 2

    if data is not None and not data.empty:
        total = len(data)
        for row_idx, (_, row_data) in enumerate(data.iterrows()):
            excel_row = first_data_row + row_idx
            is_total = str(row_data.iloc[0]) == "總計"
            for col_idx, h in enumerate(headers, 1):
                val = row_data.get(h)
                if pd.isna(val):
                    val = None
                cell = ws.cell(row=excel_row, column=col_idx, value=val)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if is_total:
                    cell.font = bold_font
                if isinstance(val, (int, float)) and h not in _OP90D_TEXT_COLUMNS:
                    cell.number_format = "#,##0"
            if progress_callback and (row_idx % 500 == 0 or row_idx == total - 1):
                p = progress_start + (progress_end - progress_start) * ((row_idx + 1) / max(total, 1))
                progress_callback(p, f"寫入 {sheet_name} ({row_idx+1}/{total})...")
    else:
        ws.cell(row=first_data_row, column=1, value="無符合資料").font = Font(size=12)
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
    breakdown: str = _OP90D_BREAKDOWN_MAKEUP,
) -> BytesIO:
    """產生四個工作表的 Excel（無 RAW；include_raw 僅為介面相容而保留，不影響輸出）。

    breakdown 為 all（三類合併顯示）時，By Shop／By Nature／By OM 採兩行合併表頭分節。
    """
    for name, frame in (("By Shop", by_shop), ("Top 10", top10), ("By Nature", by_nature), ("By OM", by_om)):
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
                       by_shop, progress_callback, 0.0, 0.25, "By Shop",
                       sections=_op90d_sheet_sections(by_shop.columns if by_shop is not None else [], breakdown))

    if progress_callback:
        progress_callback(0.3, "建立 Top 10 店舖回覆...")
    ws_top = wb.create_sheet(title="Top 10 店舖回覆")
    _write_op90d_table(ws_top,
                       list(top10.columns) if top10 is not None else _OP90D_TOP10_COLUMNS,
                       top10, progress_callback, 0.3, 0.6, "Top 10 店舖回覆", auto_filter=True)

    if progress_callback:
        progress_callback(0.65, "建立 By Nature...")
    ws_nature = wb.create_sheet(title="By Nature")
    _write_op90d_table(ws_nature,
                       list(by_nature.columns) if by_nature is not None else _OP90D_BY_NATURE_COLUMNS,
                       by_nature, progress_callback, 0.65, 0.78, "By Nature",
                       sections=_op90d_sheet_sections(by_nature.columns if by_nature is not None else [], breakdown))

    if progress_callback:
        progress_callback(0.8, "建立 By OM...")
    ws_om = wb.create_sheet(title="By OM")
    _write_op90d_table(ws_om,
                       list(by_om.columns) if by_om is not None else _OP90D_BY_OM_COLUMNS,
                       by_om, progress_callback, 0.8, 0.9, "By OM",
                       sections=_op90d_sheet_sections(by_om.columns if by_om is not None else [], breakdown))

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf
