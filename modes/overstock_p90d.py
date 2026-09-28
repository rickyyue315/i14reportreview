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
   分類維度可切換：MAKE UP 分類、Item Class、Adj. MCH1、三類合併（MAKE UP 為主分類：
   每個主分類下先列小計，再分 Item Class 與 Adj. MCH1 分項；Excel 表頭三行階層並按層配色）
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
    "_OP90D_ALL_PRIMARIES",
    "_OP90D_ALL_SUBS",
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
# 合併維度：以 MAKE UP／Non-MAKE UP 為主分類，其下再分 Item Class 與 Adj. MCH1（階層式）
_OP90D_BREAKDOWN_ALL = "all"

_OP90D_BREAKDOWNS = (_OP90D_BREAKDOWN_MAKEUP, _OP90D_BREAKDOWN_ITEM_CLASS, _OP90D_BREAKDOWN_MCH1, _OP90D_BREAKDOWN_ALL)

_OP90D_BREAKDOWN_LABELS = {
    _OP90D_BREAKDOWN_MAKEUP: "MAKE UP 分類",
    _OP90D_BREAKDOWN_ITEM_CLASS: "Item Class",
    _OP90D_BREAKDOWN_MCH1: "Adj. MCH1",
    _OP90D_BREAKDOWN_ALL: "三類合併（MAKE UP 為主分類）",
}

# 合併維度：主分類次序及其下的次分類
_OP90D_ALL_PRIMARIES = (_OP90D_MAKEUP_LABEL, _OP90D_NON_MAKEUP_LABEL)
_OP90D_ALL_SUBS = (_OP90D_BREAKDOWN_ITEM_CLASS, _OP90D_BREAKDOWN_MCH1)

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

# 階層合併表頭的底色（白字粗體）：主分類兩色，次分類（Item Class／Adj. MCH1）兩色
_OP90D_HIER_FILLS = {
    "MAKE UP": "C2255C",
    "Non-MAKE UP": "7030A0",
    "Item Class": "0E7490",
    "Adj. MCH1": "15803D",
}


def _op90d_header_layout(columns, breakdown: str):
    """合併維度（all）的階層表頭版面；其他維度回傳 None（單行表頭）。

    回傳 (n_rows, row_spans, col_header_row)：
    - n_rows：表頭行數（3：主分類 → 次分類 → 分項）
    - row_spans：{行號: [(起始欄, 跨欄數, 標題, 底色), ...]}（水平合併的階層標題）
    - col_header_row：{欄號: (行號, 底色)}　欄自有表頭所在行（垂直合併至表頭末行）
    欄位辨識：小計欄「{主分類} SKU 數」；分項欄以「{主分類}｜」前綴，Item Class 分項另帶「Item Class 」。
    """
    if breakdown != _OP90D_BREAKDOWN_ALL:
        return None
    cols = list(columns)
    blocks = []
    for pri in _OP90D_ALL_PRIMARIES:
        subtotal = _op90d_hier_subtotal_header(pri)
        block_cols = [c for c in cols if c == subtotal or c.startswith(pri + "｜")]
        if not block_cols:
            continue
        item_cols = [c for c in block_cols if "｜Item Class " in c]
        mch1_cols = [c for c in block_cols if c != subtotal and c not in item_cols]
        blocks.append((pri, subtotal, block_cols, item_cols, mch1_cols))
    if not blocks:
        return None

    n_rows = 3
    row_spans = {1: [], 2: []}
    col_header_row = {}
    in_block = {c for _p, _s, bc, _i, _m in blocks for c in bc}
    for idx, h in enumerate(cols, 1):
        if h not in in_block:
            col_header_row[idx] = (1, None)  # 前後基礎欄：垂直合併全部表頭行
    for pri, subtotal, block_cols, item_cols, mch1_cols in blocks:
        fill = _OP90D_HIER_FILLS[pri]
        start = cols.index(block_cols[0]) + 1
        row_spans[1].append((start, len(block_cols), pri, fill))
        col_header_row[cols.index(subtotal) + 1] = (2, fill)  # 小計欄：垂直合併第 2–3 行
        for sub_cols, sub_label in ((item_cols, "Item Class"), (mch1_cols, "Adj. MCH1")):
            if sub_cols:
                sub_start = cols.index(sub_cols[0]) + 1
                row_spans[2].append((sub_start, len(sub_cols), sub_label, _OP90D_HIER_FILLS[sub_label]))
    return n_rows, row_spans, col_header_row


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


def _op90d_group_header(breakdown: str, group: str) -> str:
    """單一維度的分組欄表頭。"""
    if breakdown == _OP90D_BREAKDOWN_MAKEUP:
        return "MAKE UP SKU 數" if group == _OP90D_MAKEUP_LABEL else "Non-MAKE UP SKU 數"
    return f"{group} SKU 數"


def _op90d_hier_subtotal_header(primary: str) -> str:
    """階層合併：主分類小計欄名。"""
    return f"{primary} SKU 數"


def _op90d_hier_leaf_header(primary: str, sub: str, group: str) -> str:
    """階層合併分項欄名：主分類｜分項；Item Class 分項帶次分類前綴，避免與 Adj. MCH1 的「空白」欄相撞。"""
    if sub == _OP90D_BREAKDOWN_ITEM_CLASS:
        return f"{primary}｜Item Class {group} SKU 數"
    return f"{primary}｜{group} SKU 數"


def _op90d_all_block_spec(df: pd.DataFrame) -> list:
    """階層合併的欄位規格：[(主分類, 小計欄名, [(次分類, groups, 分項欄名清單)]), ...]。

    分組次序以全量資料計算，兩個主分類共用同一套分項欄（無資料者顯示 0）。
    """
    for sub in _OP90D_ALL_SUBS:
        df[f"_group_{sub}"] = _op90d_group_series(df, sub)
    spec = []
    for pri in _OP90D_ALL_PRIMARIES:
        sub_info = []
        for sub in _OP90D_ALL_SUBS:
            groups = _op90d_ordered_groups(df[f"_group_{sub}"], sub)
            sub_info.append((sub, groups, [_op90d_hier_leaf_header(pri, sub, g) for g in groups]))
        spec.append((pri, _op90d_hier_subtotal_header(pri), sub_info))
    return spec


def _build_op90d_by_shop(raw: pd.DataFrame, breakdown: str = _OP90D_BREAKDOWN_MAKEUP) -> pd.DataFrame:
    if breakdown not in _OP90D_BREAKDOWNS:
        breakdown = _OP90D_BREAKDOWN_MAKEUP
    if raw is None or raw.empty:
        return pd.DataFrame(columns=_OP90D_BY_SHOP_BASE + _OP90D_BY_SHOP_TAIL)
    if breakdown == _OP90D_BREAKDOWN_ALL:
        return _build_op90d_by_shop_all(raw)
    df = raw.copy()
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    df["Site Code"] = site_code
    df["Shop Name"] = [_extract_shop_name(item_shop.iloc[i], site_code.iloc[i]) for i in range(len(df))]
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_group"] = _op90d_group_series(df, breakdown)
    groups = _op90d_ordered_groups(df["_group"], breakdown)
    group_headers = [_op90d_group_header(breakdown, g) for g in groups]
    headers = _OP90D_BY_SHOP_BASE + group_headers + _OP90D_BY_SHOP_TAIL

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
        for g, h in zip(groups, group_headers):
            row[h] = int(group[group["_group"] == g]["SKU"].dropna().astype(str).nunique())
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
        for g, h in zip(groups, group_headers):
            total_row[h] = int(df[df["_group"] == g]["SKU"].dropna().astype(str).nunique())
        total_row["Shop Stock Qty (CM) 總數"] = int(float(df["_stock_cm"].sum()))
        by_shop = pd.concat([by_shop, pd.DataFrame([total_row])], ignore_index=True)
    return by_shop[headers]


def _build_op90d_by_shop_all(raw: pd.DataFrame) -> pd.DataFrame:
    """三類合併（MAKE UP 為主分類）：每個主分類下先列小計，再分 Item Class 與 Adj. MCH1 分項。"""
    df = raw.copy()
    site_code = df.get("Site Code", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    item_shop = df.get("Item_Shop", _safe_default_series("", df.index))
    df["Site Code"] = site_code
    df["Shop Name"] = [_extract_shop_name(item_shop.iloc[i], site_code.iloc[i]) for i in range(len(df))]
    df["OM"] = df.get("OM", _safe_default_series("", df.index)).fillna("").astype(str).str.strip()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_pri"] = _op90d_group_series(df, _OP90D_BREAKDOWN_MAKEUP)
    spec = _op90d_all_block_spec(df)
    block_cols = []
    for _pri, subtotal, sub_info in spec:
        block_cols.append(subtotal)
        block_cols.extend(leaf for _, _, leaves in sub_info for leaf in leaves)
    headers = _OP90D_BY_SHOP_BASE + block_cols + _OP90D_BY_SHOP_TAIL

    def _counts(source: pd.DataFrame, row: dict):
        for pri, subtotal, sub_info in spec:
            psub = source[source["_pri"] == pri]
            row[subtotal] = int(psub["SKU"].dropna().astype(str).nunique())
            for sub, groups, leaves in sub_info:
                for g, leaf in zip(groups, leaves):
                    row[leaf] = int(psub[psub[f"_group_{sub}"] == g]["SKU"].dropna().astype(str).nunique())

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
        _counts(group, row)
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
        _counts(df, total_row)
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
    if raw is None or raw.empty:
        return pd.DataFrame(columns=base + tail)
    if breakdown == _OP90D_BREAKDOWN_ALL:
        return _build_distribution_all(raw, key)
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


def _build_distribution_all(raw: pd.DataFrame, key: str) -> pd.DataFrame:
    """三類合併（MAKE UP 為主分類）的分佈表：每個主分類下分 Item Class 與 Adj. MCH1 分項。"""
    base = [key, "涉及店舖數", "P90D 0 Sales SKU 數", "P90D 0 Sales 行數"]
    tail = ["Shop Stock Qty (CM) 總數"]
    df = raw.copy()
    df["SKU"] = df["SKU"].apply(_safe_sku)
    df["_dist"] = df.get(key, _safe_default_series("", df.index)).map(_display_group)
    df["_pri"] = _op90d_group_series(df, _OP90D_BREAKDOWN_MAKEUP)
    spec = _op90d_all_block_spec(df)
    block_cols = []
    for _pri, subtotal, sub_info in spec:
        block_cols.append(subtotal)
        block_cols.extend(leaf for _, _, leaves in sub_info for leaf in leaves)
    columns = base + block_cols + tail

    def _counts(source: pd.DataFrame, row: dict):
        for pri, subtotal, sub_info in spec:
            psub = source[source["_pri"] == pri]
            row[subtotal] = int(psub["SKU"].dropna().astype(str).nunique())
            for sub, groups, leaves in sub_info:
                for g, leaf in zip(groups, leaves):
                    row[leaf] = int(psub[psub[f"_group_{sub}"] == g]["SKU"].dropna().astype(str).nunique())

    rows = []
    for grp, sub in df.groupby("_dist", dropna=False, observed=True):
        row = {
            key: grp,
            "涉及店舖數": int(sub["Site Code"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
            "P90D 0 Sales SKU 數": int(sub["SKU"].dropna().astype(str).nunique()),
            "P90D 0 Sales 行數": int(len(sub)),
        }
        _counts(sub, row)
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
        _counts(df, total_row)
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
    all（三類合併、MAKE UP 為主分類：每個主分類下分 Item Class 與 Adj. MCH1）。
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
                       auto_filter: bool = False, header_layout=None):
    """寫出統計表；header_layout 為 _op90d_header_layout() 回傳值時採多行階層表頭：

    - row_spans 各行的階層標題水平合併並按層配色；
    - col_header_row 指定各欄自有表頭所在行，並垂直合併至表頭末行（前後基礎欄貫穿全部表頭行）；
    - 資料由表頭末行的下一行開始。
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

    def _col_fill(hex_code):
        if not hex_code:
            return header_fill
        return PatternFill(start_color=hex_code, end_color=hex_code, fill_type="solid")

    if header_layout:
        n_rows, row_spans, col_header_row = header_layout
        for col_idx, h in enumerate(display_headers, 1):
            # 未指定的欄（分項欄）自有表頭置於表頭末行，不垂直合併
            r0, fill_hex = col_header_row.get(col_idx, (n_rows, None))
            fill = _col_fill(fill_hex)
            cell = ws.cell(row=r0, column=col_idx, value=h)
            _style_header(cell, fill)
            if n_rows - r0 >= 1:
                for rr in range(r0 + 1, n_rows + 1):
                    _style_header(ws.cell(row=rr, column=col_idx), fill)
                ws.merge_cells(start_row=r0, start_column=col_idx, end_row=n_rows, end_column=col_idx)
        for r, spans in row_spans.items():
            for start, span, title, fill_hex in spans:
                fill = _col_fill(fill_hex)
                if span > 1:
                    ws.merge_cells(start_row=r, start_column=start, end_row=r, end_column=start + span - 1)
                cell = ws.cell(row=r, column=start, value=title)
                _style_header(cell, fill)
                for c in range(start + 1, start + span):
                    _style_header(ws.cell(row=r, column=c), fill)
        first_data_row = n_rows + 1
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

    breakdown 為 all（三類合併、MAKE UP 為主分類）時，By Shop／By Nature／By OM 採三行階層表頭
    （主分類 MAKE UP／Non-MAKE UP → 次分類 Item Class／Adj. MCH1 → 分項）。
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
                       header_layout=_op90d_header_layout(by_shop.columns if by_shop is not None else [], breakdown))

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
                       header_layout=_op90d_header_layout(by_nature.columns if by_nature is not None else [], breakdown))

    if progress_callback:
        progress_callback(0.8, "建立 By OM...")
    ws_om = wb.create_sheet(title="By OM")
    _write_op90d_table(ws_om,
                       list(by_om.columns) if by_om is not None else _OP90D_BY_OM_COLUMNS,
                       by_om, progress_callback, 0.8, 0.9, "By OM",
                       header_layout=_op90d_header_layout(by_om.columns if by_om is not None else [], breakdown))

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")

    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf
