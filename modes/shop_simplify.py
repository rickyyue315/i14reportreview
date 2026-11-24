# -*- coding: utf-8 -*-
"""店舖精簡SKU模式：把 By SKUxShop 明細轉化為店舖精簡 SKU 回覆表。

每列一店×一 SKU，保留來源檔全部欄位，並在最後附加八個固定欄位（另加一個金額計算欄）：

- 過去90日有無賣過：Sold Qty (P90D) > 0 → 「有」，否則（等於 0、負數或空白）→「無賣」；
- 過去90日有無存貨：Shop Stock Qty (LLM)、(LM) 或 (CM) 任一 > 0 →「有存貨」
  （LLM 欄自 2026-10-26 起來源報表提供；缺欄時以其餘庫存欄判定），
  三者同時為 0（或空白／負數）→「無存貨」；
- 是否該店的Top N%（件數）：每間店按所選分組（MCH2 中類（預設）／MC 小類／Adj. SST 陳列分類）
  ，在每個分類內按 Sold Qty (P90D) 由高至低累計，累計銷量首次達該店該分類總銷量 N% 的
  SKU 為「是」（同一銷量整組納入或整組排除；分類之間不互相擠佔；缺 Site Code 時
  全表單一分組；所選分組欄缺失時按 MCH2 → MC → Adj. SST 退回並列 warning，全缺則每店單一分組）；
- 是否該店的Top N%（金額）：同一分組與同一 N%，改按 Net Sales (P{N}D)＝
  Net Sales (M-2)＋(M-1)＋(CM - MTD)（先按原值相加，容許退貨抵銷同品項其他月份銷貨，
  總和下限為 0）累計；
- 該店 件數 vs 該店 金額 Top N%：由上述兩個判定合成四象限——「兩者皆是」／
  「僅店舖件數」／「僅店舖金額」／「兩者皆非」，一眼看出同一間店的件數與金額
  是否同為核心（缺三個月 Net Sales 欄時整欄不加入）；
- SCT建議、Shop 回覆新的NDRF、如果RF的Safety Qty：一律留空，供 SCT／店舖填寫；
  指定店舖變體可另傳 `sct_rules`（16 種 Top 判定情況 → 首選RF／次選RF／建議RF／留空）
  自動填寫 SCT建議，其餘兩個回覆欄維持留空。

N 可選 80%（預設）或 70%；金額欄名隨數據截止日自動標示實際覆蓋日數（如 P73D），
緊接 Sold Qty (P90D) 之後。本模式不作調貨、Safety 或退倉建議，亦不篩選列
（全店全 SKU 輸出）；清理規則沿用 `_clean_dataframe`（剔除合計列與 117 - Marketing），
Shop Nature 篩選為選填。
另由 Item_Shop 以 Site Code 解析 SKU Desc 與 Shop Name（SKU Desc 排於 Brand 之後、
Shop Name 排於 Site Code 之後），其餘來源欄序不變；欄名含「Value」的欄位一律刪除
（大小寫敏感，與「RAW 明細轉化 (不含 Value)」同一判定，故 Valuation Class 不受影響）。
輸出經 Brand／SKU（文字）排序，並產生 Summary（Adj. MCH1 × Item Status 計數）與
Export（回覆明細）兩個工作表；Export 套用店舖精簡回覆表範本的欄寬、隱藏欄、
數字格式、凍結窗格與顏色標示——其中 Item Status 為「Discontinued」、RP Type 為
「ND」、過去90日有無存貨為「無存貨」三類值以紅色底標示（比對忽略大小寫與
前後空白；RP Type 的紅色底會蓋過該欄的綠色），兩個模式共用同一匯出引擎。

另有「指定店舖」變體（`transform_shop_simplify_site`）：只輸出指定店舖（如 HA21）
的 SKU 列，同時輸出件數與金額兩組 Top N% 判定（四欄＋兩欄對比；N 可選 80%（預設）或 70%）——

- 件數：是否{店舖}的Top N%（件數）／是否同Shop Nature的Top N%（件數），按 Sold Qty (P90D)；
- 金額：是否{店舖}的Top N%（金額）／是否同Shop Nature的Top N%（金額），按 Net Sales (P{N}D)＝
  Net Sales (M-2)＋(M-1)＋(CM - MTD)（先按原值相加，容許退貨抵銷同品項其他月份銷貨，
  總和下限為 0），N 由數據截止日（T-1）的 MTD 日數＋
  兩個整月日數決定。
- 對比：{店舖} vs 同Nature Top N%（件數）／（金額），值為「兩者皆是／僅店舖／僅同Nature／
  兩者皆非」，一眼看出該 SKU 在兩邊是否同為核心。

六欄均按所選分組（介面可選 MCH2（預設中類）／MC（小類）／Adj. SST（陳列分類）），
分類之間不互相擠佔；所選分組欄缺失時按 MCH2 → MC → Adj. SST 退回並列 warning，
全缺則退回全店／全 Nature 單一分組；缺三個月 Net Sales 欄時只輸出件數兩欄（兩欄對比）。
欄名隨比例與視窗改變（如 Top 80%／Net Sales (P73D)），故切換重跑不會留下舊欄。
原模式（多店輸出）同樣輸出件數與金額兩個判定欄及一欄「該店 件數 vs 該店 金額 Top N%」
對比，但不設同 Shop Nature 判定，金額欄與判定欄的計算範圍為各店自己的 (Site Code, 所選分組)。

六個判定／對比欄只加在指定店舖的列上，其餘計算、清理、解析與回覆欄留空規則與原模式相同。
匯出時指定店舖變體另設「預設顯示欄」檢視（`site_view=True`）：只顯示識別、庫存、銷量與
Top N% 判定等 23 個固定欄加 7 個動態欄，其餘欄（Shop Name／OM／三個月庫存與銷量明細／
Shop Stock Status／SCT建議 等）一律隱藏，開啟檔案即為 SCT 回覆用的檢視。
"""
import calendar
import datetime
import gc
import itertools
import re
from io import BytesIO

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from modes.common import (
    _cached_safe_sku_series,
    _clean_dataframe,
    _compute_raw_extra_columns,
    _resolve_column_name,
    _safe_numeric,
    validate_input,
)
from modes.raw_convert import _is_raw_convert_value_column

__all__ = [
    "_SHOP_SIMPLIFY_SOLD_COLUMN",
    "_SHOP_SIMPLIFY_STOCK_COLUMN",
    "_SHOP_SIMPLIFY_SCT_COLUMN",
    "_SHOP_SIMPLIFY_SCT_PREFERRED",
    "_SHOP_SIMPLIFY_SCT_SECONDARY",
    "_SHOP_SIMPLIFY_SCT_RF",
    "_SHOP_SIMPLIFY_SCT_TIER_RANK",
    "_SHOP_SIMPLIFY_SCT_TIER_LABELS",
    "_shop_simplify_sct_tier_label",
    "_SHOP_SIMPLIFY_SCT_OPTIONS",
    "_SHOP_SIMPLIFY_SCT_FLAG_LABELS",
    "_SHOP_SIMPLIFY_SCT_CASES",
    "_SHOP_SIMPLIFY_SCT_PRESETS",
    "_SHOP_SIMPLIFY_SCT_DEFAULT_PRESET",
    "_shop_simplify_sct_case_text",
    "_shop_simplify_sct_case_compare",
    "_shop_simplify_sct_series",
    "_SHOP_SIMPLIFY_NDRF_COLUMN",
    "_SHOP_SIMPLIFY_SAFETY_COLUMN",
    "_SHOP_SIMPLIFY_REPLY_COLUMNS",
    "_SHOP_SIMPLIFY_SOLD_YES",
    "_SHOP_SIMPLIFY_SOLD_NO",
    "_SHOP_SIMPLIFY_STOCK_YES",
    "_SHOP_SIMPLIFY_STOCK_NO",
    "_SHOP_SIMPLIFY_EXPORT_SHEET",
    "_SHOP_SIMPLIFY_SUMMARY_SHEET",
    "_SHOP_SIMPLIFY_TOP80_SHARE",
    "_SHOP_SIMPLIFY_TOP80_GROUP_COLUMN",
    "_SHOP_SIMPLIFY_GROUP_OPTIONS",
    "_SHOP_SIMPLIFY_GROUP_FALLBACK_CHAIN",
    "_SHOP_SIMPLIFY_TOP_RATIO_OPTIONS",
    "_SHOP_SIMPLIFY_NET_SALES_COLUMNS",
    "_SHOP_SIMPLIFY_QTY_SUFFIX",
    "_SHOP_SIMPLIFY_VALUE_SUFFIX",
    "_SHOP_SIMPLIFY_COMPARE_BOTH",
    "_SHOP_SIMPLIFY_COMPARE_SHOP_ONLY",
    "_SHOP_SIMPLIFY_COMPARE_NATURE_ONLY",
    "_SHOP_SIMPLIFY_COMPARE_NEITHER",
    "_SHOP_SIMPLIFY_COMPARE_LABELS",
    "_SHOP_SIMPLIFY_COMPARE_TEMPLATE",
    "_SHOP_SIMPLIFY_COMPARE_BOTH_FILL",
    "_SHOP_SIMPLIFY_COMPARE_NEITHER_FILL",
    "_SHOP_SIMPLIFY_MEASURE_COMPARE_QTY_ONLY",
    "_SHOP_SIMPLIFY_MEASURE_COMPARE_VALUE_ONLY",
    "_SHOP_SIMPLIFY_MEASURE_COMPARE_LABELS",
    "_SHOP_SIMPLIFY_SHOP_TOP_TEMPLATE",
    "_SHOP_SIMPLIFY_SHOP_COMPARE_TEMPLATE",
    "_SHOP_SIMPLIFY_SITE_TOP80_TEMPLATE",
    "_SHOP_SIMPLIFY_NATURE_TOP80_TEMPLATE",
    "_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN",
    "_SHOP_SIMPLIFY_TOP80_YES",
    "_SHOP_SIMPLIFY_TOP80_NO",
    "_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN",
    "_SHOP_SIMPLIFY_SHOP_VALUE_TOP80_COLUMN",
    "_SHOP_SIMPLIFY_SHOP_COMPARE80_COLUMN",
    "_SHOP_SIMPLIFY_RED",
    "_SHOP_SIMPLIFY_RED_RULES",
    "_shop_simplify_shop_qty_top_column",
    "_shop_simplify_shop_value_top_column",
    "_shop_simplify_shop_compare_column",
    "_shop_simplify_measure_compare_labels",
    "_is_shop_simplify_measure_compare_column",
    "_shop_simplify_site_top80_column",
    "_shop_simplify_nature_top_column",
    "_shop_simplify_value_site_top_column",
    "_shop_simplify_value_nature_top_column",
    "_shop_simplify_net_sales_column",
    "_shop_simplify_net_sales_window",
    "_shop_simplify_net_sales_series",
    "_shop_simplify_compare_column",
    "_shop_simplify_compare_labels",
    "_is_shop_simplify_compare_column",
    "_is_shop_simplify_value_compare_column",
    "_is_shop_simplify_value_top_column",
    "_shop_simplify_top_share",
    "_shop_simplify_ratio_percent",
    "_shop_simplify_top80_mask",
    "_shop_simplify_top80_keys",
    "_shop_simplify_group_column",
    "_shop_simplify_group_choice",
    "shop_simplify_stats",
    "shop_simplify_site_stats",
    "transform_shop_simplify",
    "transform_shop_simplify_site",
    "generate_excel_shop_simplify",
]


# ── 輸出欄位契約（欄名固定，測試與匯出檔依賴，未經明確要求不得改動）──
_SHOP_SIMPLIFY_SOLD_COLUMN = "過去90日有無賣過"
_SHOP_SIMPLIFY_STOCK_COLUMN = "過去90日有無存貨"
_SHOP_SIMPLIFY_SCT_COLUMN = "SCT建議"
_SHOP_SIMPLIFY_NDRF_COLUMN = "Shop 回覆新的NDRF"
_SHOP_SIMPLIFY_SAFETY_COLUMN = "如果RF的Safety Qty"

# 需要由 SCT／店舖填寫的三欄（一律留空；不沿用來源檔舊值）
_SHOP_SIMPLIFY_REPLY_COLUMNS = [
    _SHOP_SIMPLIFY_SCT_COLUMN,
    _SHOP_SIMPLIFY_NDRF_COLUMN,
    _SHOP_SIMPLIFY_SAFETY_COLUMN,
]

# 兩個判定欄的固定取值
_SHOP_SIMPLIFY_SOLD_YES = "有"
_SHOP_SIMPLIFY_SOLD_NO = "無賣"
_SHOP_SIMPLIFY_STOCK_YES = "有存貨"
_SHOP_SIMPLIFY_STOCK_NO = "無存貨"

# ── 原模式與指定店舖模式的 Top N% 判定（預設 80%，介面可選 70%）──
# 累計銷量／金額達該分組（該店／同 Shop Nature 全部店舖）總量的 N% 即屬 Top N%；
# 同一銷量／金額的 SKU 整組納入或整組排除，避免同值被隨意切開。
_SHOP_SIMPLIFY_TOP80_SHARE = 0.8
# 分組欄（介面可選，預設 MCH2）：MCH2＝中類（如「0101 - FACE & NECK」）、
# MC＝小類（如「0101004 - AMPOULES & SERUM」）、Adj. SST＝陳列分類（舊口徑）。
# 所選欄缺欄時按下列順序退回並列 warning；全缺則單一分組。
_SHOP_SIMPLIFY_GROUP_OPTIONS = ("MCH2", "MC", "Adj. SST")
_SHOP_SIMPLIFY_TOP80_GROUP_COLUMN = "MCH2"
_SHOP_SIMPLIFY_GROUP_FALLBACK_CHAIN = {
    "MCH2": ("MCH2", "MC", "Adj. SST"),
    "MC": ("MC", "Adj. SST"),
    "Adj. SST": ("Adj. SST",),
}
# 介面可選的比例（0.8＝Top 80%、0.7＝Top 70%）；兩個模式共用
_SHOP_SIMPLIFY_TOP_RATIO_OPTIONS = (0.8, 0.7)
_SHOP_SIMPLIFY_SHOP_TOP_TEMPLATE = "是否該店的Top {pct}%"
_SHOP_SIMPLIFY_SITE_TOP80_TEMPLATE = "是否{site}的Top {pct}%"
_SHOP_SIMPLIFY_NATURE_TOP80_TEMPLATE = "是否同Shop Nature的Top {pct}%"
# 兩個模式都同時輸出件數與金額兩種口徑，故判定欄一律加註口徑，避免同名混淆
_SHOP_SIMPLIFY_QTY_SUFFIX = "（件數）"
_SHOP_SIMPLIFY_VALUE_SUFFIX = "（金額）"
_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN = (
    _SHOP_SIMPLIFY_NATURE_TOP80_TEMPLATE.format(pct=80) + _SHOP_SIMPLIFY_QTY_SUFFIX
)
_SHOP_SIMPLIFY_TOP80_YES = "是"
_SHOP_SIMPLIFY_TOP80_NO = "否"
# 店舖 vs 同 Nature 的對比（四象限），讓 SCT 一眼分辨「兩邊都核心」、「只有本店賣得動」、
# 「只有同 Nature 其他店賣得動」與「兩邊都不是核心」
_SHOP_SIMPLIFY_COMPARE_BOTH = "兩者皆是"
_SHOP_SIMPLIFY_COMPARE_SHOP_ONLY = "僅店舖"
_SHOP_SIMPLIFY_COMPARE_NATURE_ONLY = "僅同Nature"
_SHOP_SIMPLIFY_COMPARE_NEITHER = "兩者皆非"
_SHOP_SIMPLIFY_COMPARE_LABELS = (
    _SHOP_SIMPLIFY_COMPARE_BOTH,
    _SHOP_SIMPLIFY_COMPARE_SHOP_ONLY,
    _SHOP_SIMPLIFY_COMPARE_NATURE_ONLY,
    _SHOP_SIMPLIFY_COMPARE_NEITHER,
)
_SHOP_SIMPLIFY_COMPARE_TEMPLATE = "{site} vs 同Nature Top {pct}%"
# 原模式（多店輸出）的「件數 vs 金額」對比：同一間店的兩種口徑是否同為 Top N%
_SHOP_SIMPLIFY_MEASURE_COMPARE_QTY_ONLY = "僅店舖件數"
_SHOP_SIMPLIFY_MEASURE_COMPARE_VALUE_ONLY = "僅店舖金額"
_SHOP_SIMPLIFY_MEASURE_COMPARE_LABELS = (
    _SHOP_SIMPLIFY_COMPARE_BOTH,
    _SHOP_SIMPLIFY_MEASURE_COMPARE_QTY_ONLY,
    _SHOP_SIMPLIFY_MEASURE_COMPARE_VALUE_ONLY,
    _SHOP_SIMPLIFY_COMPARE_NEITHER,
)
_SHOP_SIMPLIFY_SHOP_COMPARE_TEMPLATE = "該店 件數 vs 該店 金額 Top {pct}%"
# 金額口徑：來源只有月欄（M-2／M-1／CM-MTD），三者相加即「MTD 日數＋兩個整月」的實際覆蓋日數
# （例如數據截止 11 月 12 日 → 12＋31＋30＝73 日，欄名標示 Net Sales (P73D)）。
_SHOP_SIMPLIFY_NET_SALES_COLUMNS = ("Net Sales (M-2)", "Net Sales (M-1)", "Net Sales (CM - MTD)")
# 原模式（多店輸出）每店各自的 Top 80% 欄名（預設比例、件數口徑）；指定店舖變體改用「是否{店舖}的Top 80%」
_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN = _SHOP_SIMPLIFY_SHOP_TOP_TEMPLATE.format(pct=80) + _SHOP_SIMPLIFY_QTY_SUFFIX
_SHOP_SIMPLIFY_SHOP_VALUE_TOP80_COLUMN = _SHOP_SIMPLIFY_SHOP_TOP_TEMPLATE.format(pct=80) + _SHOP_SIMPLIFY_VALUE_SUFFIX
_SHOP_SIMPLIFY_SHOP_COMPARE80_COLUMN = _SHOP_SIMPLIFY_SHOP_COMPARE_TEMPLATE.format(pct=80)


# ── SCT建議：依四個 Top N% 判定自動填寫（2026-11-24 初步設定）──
# 四個判定各只有「是／否」，對比欄由兩個判定自動合成，故全部情況＝2×2×2×2＝16 種。
# 每個情況可指定「首選RF」「次選RF」「建議RF」或留空（留空＝由 SCT 自行判斷，例如 ND）。
_SHOP_SIMPLIFY_SCT_PREFERRED = "首選RF"
_SHOP_SIMPLIFY_SCT_SECONDARY = "次選RF"
_SHOP_SIMPLIFY_SCT_RF = "建議RF"
_SHOP_SIMPLIFY_SCT_OPTIONS = (
    "",
    _SHOP_SIMPLIFY_SCT_PREFERRED,
    _SHOP_SIMPLIFY_SCT_SECONDARY,
    _SHOP_SIMPLIFY_SCT_RF,
)
# 情況鍵：四個字母依次為 件數－店舖／件數－同Shop Nature／金額－店舖／金額－同Shop Nature，
# Y＝Top（是）、N＝非 Top（否）；例如 "YNYY"＝只有本店件數不是 Top。
_SHOP_SIMPLIFY_SCT_FLAG_LABELS = (
    "件數－店舖", "件數－同Shop Nature", "金額－店舖", "金額－同Shop Nature",
)
_SHOP_SIMPLIFY_SCT_CASES = tuple("".join(bits) for bits in itertools.product("YN", repeat=4))


def _shop_simplify_sct_case_text(case: str) -> str:
    """情況鍵 → 人看的說明，例如 "YNYY" →「只有本店件數 Top；金額兩邊都是 Top」。"""

    def part(site: str, nature: str, unit: str) -> str:
        if site == "Y" and nature == "Y":
            return f"{unit}兩邊都是 Top"
        if site == "Y":
            return f"只有本店{unit} Top"
        if nature == "Y":
            return f"只有同Nature{unit} Top"
        return f"{unit}兩邊都不是 Top"

    return f"{part(case[0], case[1], '件數')}；{part(case[2], case[3], '金額')}"


def _shop_simplify_sct_case_compare(case: str) -> tuple[str, str]:
    """情況鍵 → (件數對比, 金額對比) 兩個對比標籤。"""

    def label(site: str, nature: str) -> str:
        if site == "Y" and nature == "Y":
            return _SHOP_SIMPLIFY_COMPARE_BOTH
        if site == "Y":
            return _SHOP_SIMPLIFY_COMPARE_SHOP_ONLY
        if nature == "Y":
            return _SHOP_SIMPLIFY_COMPARE_NATURE_ONLY
        return _SHOP_SIMPLIFY_COMPARE_NEITHER

    return label(case[0], case[1]), label(case[2], case[3])


def _shop_simplify_sct_preset(rule) -> dict[str, str]:
    """把「情況鍵 → 建議」的函數展開為 16 個情況的完整字典。"""
    return {case: rule(case) for case in _SHOP_SIMPLIFY_SCT_CASES}


# 四級建議：件數與金額各自判定，取兩者之中「最好」的一級。
# 兩者皆是（兩邊都 Top）＞僅同Nature＞僅店舖＞兩者皆非
_SHOP_SIMPLIFY_SCT_TIER_RANK = {
    _SHOP_SIMPLIFY_COMPARE_BOTH: 3,
    _SHOP_SIMPLIFY_COMPARE_NATURE_ONLY: 2,
    _SHOP_SIMPLIFY_COMPARE_SHOP_ONLY: 1,
    _SHOP_SIMPLIFY_COMPARE_NEITHER: 0,
}
_SHOP_SIMPLIFY_SCT_TIER_LABELS = {
    3: _SHOP_SIMPLIFY_SCT_PREFERRED,
    2: _SHOP_SIMPLIFY_SCT_SECONDARY,
    1: _SHOP_SIMPLIFY_SCT_RF,
    0: "",
}


def _shop_simplify_sct_tier_label(case: str) -> str:
    """四級建議：按（件數對比、金額對比）之中較高的一級給建議。"""
    qty_label, value_label = _shop_simplify_sct_case_compare(case)
    best = max(_SHOP_SIMPLIFY_SCT_TIER_RANK[qty_label], _SHOP_SIMPLIFY_SCT_TIER_RANK[value_label])
    return _SHOP_SIMPLIFY_SCT_TIER_LABELS[best]


# 內建方案（介面可選；亦可自訂 16 種情況逐格指定）
_SHOP_SIMPLIFY_SCT_PRESETS = {
    "三級建議（兩者皆是→首選／同Nature→次選／只有店舖→建議RF／無Top→留空）": _shop_simplify_sct_preset(
        _shop_simplify_sct_tier_label
    ),
    "件數優先（任一件數 Top → 首選RF）": _shop_simplify_sct_preset(
        lambda c: _SHOP_SIMPLIFY_SCT_PREFERRED if "Y" in c[:2] else ""
    ),
    "雙重把關（件數＋金額都有 Top → 首選RF，只中一組 → 次選RF）": _shop_simplify_sct_preset(
        lambda c: (
            _SHOP_SIMPLIFY_SCT_PREFERRED
            if ("Y" in c[:2] and "Y" in c[2:])
            else _SHOP_SIMPLIFY_SCT_SECONDARY
            if ("Y" in c[:2] or "Y" in c[2:])
            else ""
        )
    ),
    "店舖優先（店舖件數 Top → 首選RF，只有店舖金額 Top → 次選RF）": _shop_simplify_sct_preset(
        lambda c: (
            _SHOP_SIMPLIFY_SCT_PREFERRED
            if c[0] == "Y"
            else _SHOP_SIMPLIFY_SCT_SECONDARY
            if c[2] == "Y"
            else ""
        )
    ),
    "命中項數分級（4～3 項 → 首選RF，2 項 → 次選RF）": _shop_simplify_sct_preset(
        lambda c: (
            _SHOP_SIMPLIFY_SCT_PREFERRED
            if c.count("Y") >= 3
            else _SHOP_SIMPLIFY_SCT_SECONDARY
            if c.count("Y") == 2
            else ""
        )
    ),
}
_SHOP_SIMPLIFY_SCT_DEFAULT_PRESET = "三級建議（兩者皆是→首選／同Nature→次選／只有店舖→建議RF／無Top→留空）"


def _shop_simplify_sct_series(df: pd.DataFrame, rules, columns) -> pd.Series:
    """按四個判定欄與規則（情況鍵 → 建議）產生 SCT建議；未涵蓋的情況留空。

    `columns` 次序固定為（件數店舖、件數同Nature、金額店舖、金額同Nature）；
    傳 None 或缺欄時該項一律視為非 Top（金額口徑不可用時即 "??NN"）。
    """
    mapping = {str(k).strip(): str(v or "").strip() for k, v in dict(rules).items()}
    letters = []
    for name in columns:
        if name and name in df.columns:
            flags = _shop_simplify_cell_mask(df[name], _SHOP_SIMPLIFY_TOP80_YES)
        else:
            flags = np.zeros(len(df), dtype=bool)
        letters.append(pd.Series(np.where(flags, "Y", "N"), index=df.index))
    keys = letters[0].str.cat(letters[1:], sep="")
    return keys.map(mapping).fillna("").astype(object)


# ── Excel 輸出格式（對齊店舖精簡回覆表範本）──
_SHOP_SIMPLIFY_EXPORT_SHEET = "Export"
_SHOP_SIMPLIFY_SUMMARY_SHEET = "Summary"

# 固定表頭顏色（使用者範本）
_SHOP_SIMPLIFY_HEADER_BLUE = "D9E2F3"
_SHOP_SIMPLIFY_GREEN = "CCFF99"
_SHOP_SIMPLIFY_PURPLE = "E4DFEC"
_SHOP_SIMPLIFY_NO_SALE_PINK = "F2DCDB"

# 紅色底（醒目紅底；比「無賣」的淡紅 F2DCDB 明顯，故兩者不會混淆）
_SHOP_SIMPLIFY_RED = "FF9999"
# 對比欄底色：兩者皆是＝青綠（核心，優先保留）；兩者皆非＝粉紅（精簡候選）
_SHOP_SIMPLIFY_COMPARE_BOTH_FILL = "99E6CC"
_SHOP_SIMPLIFY_COMPARE_NEITHER_FILL = "FFC7CE"
# 紅色底判定：欄名 → 觸發值（比對前去除前後空白、忽略大小寫；欄位不存在時略過該條件）
_SHOP_SIMPLIFY_RED_RULES = (
    ("Item Status", "discontinued"),
    ("RP Type", "nd"),
    (_SHOP_SIMPLIFY_STOCK_COLUMN, _SHOP_SIMPLIFY_STOCK_NO.lower()),
)

# 綠色表頭（另加綠色資料格：RP Type／Shop Stock Qty (CM)）
_SHOP_SIMPLIFY_GREEN_HEADERS = {
    "RP Type", "Shop Stock Qty (CM)", _SHOP_SIMPLIFY_NDRF_COLUMN, _SHOP_SIMPLIFY_SAFETY_COLUMN,
}
_SHOP_SIMPLIFY_GREEN_DATA_COLUMNS = {"RP Type", "Shop Stock Qty (CM)"}
# 紫色表頭與資料格（Sold Qty (P90D)）
_SHOP_SIMPLIFY_PURPLE_COLUMNS = {"Sold Qty (P90D)"}

# 隱藏欄（使用者範本；按欄名對應，來源缺欄時自動略過）
_SHOP_SIMPLIFY_HIDDEN_COLUMNS = {
    "Item_Shop", "Valuation Class", "MCH2", "MC", "Shop Nature", "Shop Type",
    "D001 Stock Qty (LM)", "D001 Blocked Qty (CM)", "D001 QI Qty (CM)",
    "Turnday (P90D Qty)", "Net Sales (M-2)", "Excess Stock Qty",
    "D001 On Order PO Qty (CM)", "Shop On Order PO Qty (CM)", "Shop Pending Receive (±7 days)",
}

# 指定店舖變體的預設顯示欄（2026-11-18 使用者指定）：匯出檔一打開只顯示下列欄，
# 其餘一律隱藏（使用者仍可在 Excel 內取消隱藏）。欄名固定者列於此；隨店舖／Top 比例／
# 數據截止日改變的欄名（Net Sales (P{N}D)、是否{店舖}的Top N%、同 Nature 判定與對比）
# 由 `_is_shop_simplify_site_visible_column` 以樣式比對，不寫死比例與日數。
# 註：本清單不含 SCT建議，故該欄在指定店舖變體亦會隱藏（與使用者清單一致）。
_SHOP_SIMPLIFY_SITE_VISIBLE_COLUMNS = (
    "SKU", "Brand", "SKU Desc", "Buyer", "Adj. SST", "Vendor MC", "Adj. MCH1",
    "Item Status", "Launch Date", "Item Class", "Supply Source", "Site Code",
    "RP Type", "Adj. Safety Stock", "D001 Net Stock Qty (CM)", "Shop Stock Qty (CM)",
    "Sold Qty (P90D)", "Sold Qty (CM - MTD)", "Net Sales (CM - MTD)",
    _SHOP_SIMPLIFY_SOLD_COLUMN, _SHOP_SIMPLIFY_STOCK_COLUMN,
    _SHOP_SIMPLIFY_NDRF_COLUMN, _SHOP_SIMPLIFY_SAFETY_COLUMN,
)
# 金額計算欄（Net Sales (P73D)／(P78D)…）隨數據截止日改變日數，故以樣式比對
_SHOP_SIMPLIFY_SITE_NET_SALES_RE = re.compile(r"Net Sales \(P\d+D\)")

# 欄寬（使用者範本；未列出的欄沿用 Excel 預設寬度）
_SHOP_SIMPLIFY_COLUMN_WIDTHS = {
    "Item_Shop": 79.7109375,
    "SKU": 13.140625,
    "Valuation Class": 25.0,
    "Brand": 8.7109375,
    "SKU Desc": 46.42578125,
    "Buyer": 18.85546875,
    "Adj. SST": 29.42578125,
    "Vendor MC": 23.140625,
    "Adj. MCH1": 17.28515625,
    "MCH2": 48.28515625,
    "MC": 56.140625,
    "Item Status": 12.7109375,
    "Launch Date": 10.5703125,
    "Item Class": 10.0,
    "Supply Source": 11.140625,
    "Site Code": 10.140625,
    "Shop Name": 10.85546875,
    "OM": 8.85546875,
    "Shop Nature": 11.28515625,
    "Shop Type": 10.0,
    "RP Type": 9.85546875,
    "Adj. Safety Stock": 10.28515625,
    "Shop Monthly Sold Qty (P90D)": 13.140625,
    "Safety VS Monthly Sold Qty": 10.0,
    "D001 Stock Qty (LM)": 10.28515625,
    "D001 Net Stock Qty (CM)": 9.7109375,
    "D001 Blocked Qty (CM)": 11.42578125,
    "D001 QI Qty (CM)": 10.0,
    "Shop Stock Qty (LM)": 9.0,
    "Shop Stock Qty (CM)": 9.0,
    "Sold Qty (P90D)": 10.7109375,
    "Turnday (P90D Qty)": 12.0,
    "Sold Qty (M-2)": 10.5703125,
    "Sold Qty (M-1)": 10.5703125,
    "Sold Qty (CM - MTD)": 10.42578125,
    "Net Sales (M-2)": 11.0,
    "Net Sales (CM - MTD)": 10.42578125,
    "Shop Stock Status": 56.85546875,
    "Excess Stock Qty": 10.42578125,
    "D001 On Order PO Qty (CM)": 12.85546875,
    "Shop On Order PO Qty (CM)": 13.0,
    "Shop Pending Receive (±7 days)": 13.42578125,
    _SHOP_SIMPLIFY_SOLD_COLUMN: 10.85546875,
    _SHOP_SIMPLIFY_STOCK_COLUMN: 10.85546875,
    _SHOP_SIMPLIFY_NATURE_TOP80_COLUMN: 26.0,
    _SHOP_SIMPLIFY_NDRF_COLUMN: 8.42578125,
    _SHOP_SIMPLIFY_SAFETY_COLUMN: 9.7109375,
}

# 數字格式（使用者範本）
_SHOP_SIMPLIFY_NUMBER_FORMATS = {
    "Shop Monthly Sold Qty (P90D)": '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)',
    "Safety VS Monthly Sold Qty": "0%",
    "Turnday (P90D Qty)": "#,##0_);[Red](#,##0)",
    "Sold Qty (M-2)": "#,##0_);[Red](#,##0)",
    "Sold Qty (M-1)": "#,##0_);[Red](#,##0)",
    "Sold Qty (CM - MTD)": "#,##0_);[Red](#,##0)",
    "Net Sales (M-2)": "#,##0_);[Red](#,##0)",
    "Net Sales (M-1)": "#,##0_);[Red](#,##0)",
    "Net Sales (CM - MTD)": "#,##0_);[Red](#,##0)",
}


def shop_simplify_stats(df: pd.DataFrame, group_by=None) -> dict:
    """結果頁指標卡用的統計（不影響輸出檔）：列數、無賣、無存貨、同時無賣且無存貨、店舖數、SKU 數，
    以及件數／金額 Top N% 命中數與「件數 vs 金額」四象限計數。

    比例與金額視窗由結果檔欄名取回（切換側欄設定但未重算時，統計仍反映該檔實際結果）；
    金額口徑只在結果檔含「（金額）」判定欄時提供，否則 `value_available` 為 False。
    """
    stats = {
        "rows": 0, "sold_no": 0, "stock_no": 0, "both_no": 0, "sites": 0, "skus": 0,
        "ratio": _shop_simplify_ratio_percent(None),
        "qty_top80": 0, "value_top80": 0, "value_available": False, "window": "",
        "compare": {}, "group_column": "",
    }
    if df is None or len(df) == 0:
        return stats
    sold_no = df.get(_SHOP_SIMPLIFY_SOLD_COLUMN, pd.Series("", index=df.index)).astype(str) == _SHOP_SIMPLIFY_SOLD_NO
    stock_no = df.get(_SHOP_SIMPLIFY_STOCK_COLUMN, pd.Series("", index=df.index)).astype(str) == _SHOP_SIMPLIFY_STOCK_NO
    stats["rows"] = int(len(df))
    stats["sold_no"] = int(sold_no.sum())
    stats["stock_no"] = int(stock_no.sum())
    stats["both_no"] = int((sold_no & stock_no).sum())
    if "Site Code" in df.columns:
        sites = df["Site Code"].dropna().astype(str).str.strip()
        stats["sites"] = int(sites[sites != ""].nunique())
    if "SKU" in df.columns:
        skus = df["SKU"].dropna().astype(str).str.strip()
        stats["skus"] = int(skus[skus != ""].nunique())

    def _flags(column: str) -> pd.Series:
        return df[column].fillna("").astype(str).str.strip()

    stats["group_column"] = _shop_simplify_group_column(df.columns, group_by)[1]

    # 原模式的判定欄：店舖層（「是否該店的Top N%」），不含同 Nature 欄
    qty_columns = [
        c for c in df.columns
        if _is_shop_simplify_top80_column(c)
        and not _is_shop_simplify_value_top_column(c)
        and not str(c).strip().startswith("是否同Shop Nature")
    ]
    value_columns = [
        c for c in df.columns
        if _is_shop_simplify_value_top_column(c)
        and not str(c).strip().startswith("是否同Shop Nature")
    ]
    if qty_columns:
        qty_column = qty_columns[0]
        stats["ratio"] = _shop_simplify_top_column_percent(qty_column) or stats["ratio"]
        stats["qty_top80"] = int((_flags(qty_column) == _SHOP_SIMPLIFY_TOP80_YES).sum())
    if value_columns:
        stats["value_available"] = True
        stats["value_top80"] = int((_flags(value_columns[0]) == _SHOP_SIMPLIFY_TOP80_YES).sum())

    compare_columns = [c for c in df.columns if _is_shop_simplify_measure_compare_column(c)]
    if compare_columns:
        labels = _flags(compare_columns[0])
        stats["compare"] = {
            label: int((labels == label).sum()) for label in _SHOP_SIMPLIFY_MEASURE_COMPARE_LABELS
        }
    for c in df.columns:
        if isinstance(c, str) and re.fullmatch(r"Net Sales \(P\d+D\)", c.strip()):
            stats["window"] = c.strip()[len("Net Sales ("):-1]
            break
    return stats


def shop_simplify_site_stats(df: pd.DataFrame, site_code: str, top_ratio=None, group_by=None) -> dict:
    """指定店舖模式的結果頁指標：列數、SKU 數、分組數（所選 MCH2／MC／Adj. SST）、
    件數與金額兩組 Top N% 命中數、兩項皆非、無賣／無存貨。

    `top_ratio` 只在判定欄不存在時作後備；正常情況由欄名（「…Top 70%」）取回實際比例，
    故同一份結果不論介面當前選哪個比例，統計都反映該檔實際使用的比例。
    金額口徑只在結果檔含「（金額）」判定欄時提供，否則 `value_available` 為 False。
    """
    site = str(site_code or "").strip()
    stats = {
        "rows": 0, "skus": 0, "group_count": 0, "group_column": "",
        "sold_no": 0, "stock_no": 0,
        "site_top80": 0, "nature_top80": 0, "neither": 0,
        "value_site_top80": 0, "value_nature_top80": 0, "value_neither": 0,
        "value_available": False, "window": "",
        "compare_qty": {}, "compare_value": {},
        "site": site, "nature": "", "ratio": _shop_simplify_ratio_percent(top_ratio),
    }
    if df is None or len(df) == 0:
        return stats

    def _flag(column: str) -> pd.Series:
        if column not in df.columns:
            return pd.Series("", index=df.index, dtype=object)
        return df[column].fillna("").astype(str).str.strip()

    sold_no = _flag(_SHOP_SIMPLIFY_SOLD_COLUMN) == _SHOP_SIMPLIFY_SOLD_NO
    stock_no = _flag(_SHOP_SIMPLIFY_STOCK_COLUMN) == _SHOP_SIMPLIFY_STOCK_NO
    nature_columns = [
        c for c in df.columns
        if isinstance(c, str) and c.strip().startswith("是否同Shop Nature的Top ")
    ]
    nature_qty_columns = [c for c in nature_columns if not _is_shop_simplify_value_top_column(c)]
    nature_value_columns = [c for c in nature_columns if _is_shop_simplify_value_top_column(c)]
    site_qty_columns = [
        c for c in df.columns
        if _is_shop_simplify_top80_column(c)
        and c not in nature_columns
        and not _is_shop_simplify_value_top_column(c)
    ]
    site_value_columns = [
        c for c in df.columns
        if _is_shop_simplify_value_top_column(c) and c not in nature_columns
    ]
    site_column = site_qty_columns[0] if len(site_qty_columns) == 1 else _shop_simplify_site_top80_column(site, top_ratio)
    nature_column = nature_qty_columns[0] if len(nature_qty_columns) == 1 else _shop_simplify_nature_top_column(top_ratio)
    site_flag = _flag(site_column)
    nature_flag = _flag(nature_column)
    stats["ratio"] = (
        _shop_simplify_top_column_percent(site_column)
        or _shop_simplify_top_column_percent(nature_column)
        or _shop_simplify_ratio_percent(top_ratio)
    )

    stats["rows"] = int(len(df))
    stats["sold_no"] = int(sold_no.sum())
    stats["stock_no"] = int(stock_no.sum())
    stats["site_top80"] = int((site_flag == _SHOP_SIMPLIFY_TOP80_YES).sum())
    stats["nature_top80"] = int((nature_flag == _SHOP_SIMPLIFY_TOP80_YES).sum())
    stats["neither"] = int(((site_flag == _SHOP_SIMPLIFY_TOP80_NO) & (nature_flag == _SHOP_SIMPLIFY_TOP80_NO)).sum())

    if site_value_columns and nature_value_columns:
        value_site_flag = _flag(site_value_columns[0])
        value_nature_flag = _flag(nature_value_columns[0])
        stats["value_available"] = True
        stats["value_site_top80"] = int((value_site_flag == _SHOP_SIMPLIFY_TOP80_YES).sum())
        stats["value_nature_top80"] = int((value_nature_flag == _SHOP_SIMPLIFY_TOP80_YES).sum())
        stats["value_neither"] = int(
            ((value_site_flag == _SHOP_SIMPLIFY_TOP80_NO) & (value_nature_flag == _SHOP_SIMPLIFY_TOP80_NO)).sum()
        )

    # 店舖 vs 同 Nature 的四象限計數（件數／金額各一組）
    compare_columns = [c for c in df.columns if _is_shop_simplify_compare_column(c)]
    compare_qty_columns = [c for c in compare_columns if not _is_shop_simplify_value_compare_column(c)]
    compare_value_columns = [c for c in compare_columns if _is_shop_simplify_value_compare_column(c)]
    if compare_qty_columns:
        flags = _flag(compare_qty_columns[0])
        stats["compare_qty"] = {
            label: int((flags == label).sum()) for label in _SHOP_SIMPLIFY_COMPARE_LABELS
        }
    if compare_value_columns:
        flags = _flag(compare_value_columns[0])
        stats["compare_value"] = {
            label: int((flags == label).sum()) for label in _SHOP_SIMPLIFY_COMPARE_LABELS
        }
    for c in df.columns:
        if isinstance(c, str) and c.strip().startswith("Net Sales (P") and c.strip().endswith("D)"):
            stats["window"] = c.strip()[len("Net Sales ("):-1]
            break

    if "SKU" in df.columns:
        skus = df["SKU"].dropna().astype(str).str.strip()
        stats["skus"] = int(skus[skus != ""].nunique())
    group_column, group_label = _shop_simplify_group_column(df.columns, group_by)
    if group_column is not None and group_column in df.columns:
        groups = df[group_column].dropna().astype(str).str.strip()
        groups = groups[groups != ""]
        stats["group_count"] = int(groups.nunique())
        stats["group_column"] = group_label
    if "Shop Nature" in df.columns:
        natures = df["Shop Nature"].dropna().astype(str).str.strip()
        natures = natures[natures != ""]
        if not natures.empty:
            stats["nature"] = str(natures.mode().iloc[0])
    return stats


def _shop_simplify_place_derived_columns(df: pd.DataFrame) -> pd.DataFrame:
    """把解析出的 SKU Desc／Shop Name 移到慣用位置，其餘來源欄序不變。

    SKU Desc 排於 Brand 之後（缺 Brand 時依序退回 SKU／Item_Shop）；
    Shop Name 排於 Site Code 之後（缺 Site Code 時依序退回 OM／Shop Nature／SKU Desc）。
    """
    def _move(frame: pd.DataFrame, column: str, anchors: tuple[str, ...]) -> pd.DataFrame:
        if column not in frame.columns:
            return frame
        cols = [c for c in frame.columns if c != column]
        for anchor in anchors:
            if anchor in cols:
                cols.insert(cols.index(anchor) + 1, column)
                break
        else:
            cols.append(column)
        return frame[cols]

    out = _move(df, "SKU Desc", ("Brand", "SKU", "Item_Shop"))
    out = _move(out, "Shop Name", ("Site Code", "OM", "Shop Nature", "SKU Desc"))
    return out


def _shop_simplify_launch_int(value):
    """Launch Date 統一為 YYYYMMDD 整數（來源多為 float 如 20120520.0）；
    日期物件、YYYY-MM-DD、YYYY/MM/DD 皆可解析；無法解析或空白回傳 NaN。"""
    if value is None:
        return np.nan
    try:
        if pd.isna(value):
            return np.nan
    except (TypeError, ValueError):
        pass
    if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
        return int(value.strftime("%Y%m%d"))
    text = re.sub(r"\.0+$", "", str(value).strip())
    text = text.replace("-", "").replace("/", "")
    if len(text) == 8 and text.isdigit():
        return int(text)
    return np.nan


def _shop_simplify_top_share(ratio=None) -> float:
    """正規化 Top 比例：接受 0.7／70／"70%"；必須在 (0, 1] 之內，否則拋 ValueError。

    未提供時回傳預設 0.8（Top 80%），故舊呼叫端的行為與欄名完全不變。
    """
    if ratio is None:
        return _SHOP_SIMPLIFY_TOP80_SHARE
    value = ratio
    if isinstance(value, str):
        text = value.strip().replace("%", "")
        try:
            value = float(text)
        except (TypeError, ValueError):
            raise ValueError(
                f"Top 比例「{ratio}」無法識別，請提供 0 至 1 的比例（如 0.7）或百分比（如 70）。"
            )
    try:
        value = float(value)
    except (TypeError, ValueError):
        raise ValueError(f"Top 比例「{ratio!r}」不是數值，請提供 0 至 1 的比例（如 0.7）或百分比（如 70）。")
    if value > 1:
        value = value / 100.0
    if not 0 < value <= 1:
        raise ValueError(f"Top 比例必須大於 0 且不超過 100%（收到 {ratio!r}）。")
    return value


def _shop_simplify_ratio_percent(ratio=None) -> int:
    """比例轉整數百分比（0.7 → 70），用於欄名與文案。"""
    return int(round(_shop_simplify_top_share(ratio) * 100))


def _shop_simplify_top_column_percent(name) -> int | None:
    """從欄名取回百分比（「是否HA21的Top 70%」→ 70）；不是 Top 欄時回傳 None。"""
    if not isinstance(name, str):
        return None
    match = re.search(r"Top\s*(\d{1,3})%", name.strip())
    return int(match.group(1)) if match else None


def _shop_simplify_net_sales_window(data_date=None) -> dict:
    """回傳金額覆蓋日數的拆解：MTD 日數＋M-1 月日數＋M-2 月日數。

    來源數據為 T-1，故未指定時以「今天減一日」為數據截止日。例如 11 月 13 日執行 →
    截止 11 月 12 日：MTD 12 日＋10 月 31 日＋9 月 30 日＝73 日，標示 `P73D`。
    """
    if data_date is None:
        data_date = datetime.date.today() - datetime.timedelta(days=1)
    if isinstance(data_date, datetime.datetime):
        data_date = data_date.date()
    if not isinstance(data_date, datetime.date):
        raise ValueError(f"數據截止日「{data_date!r}」不是日期，無法計算 Net Sales 視窗。")
    mtd_days = int(data_date.day)
    m1_end = data_date.replace(day=1) - datetime.timedelta(days=1)
    m1_days = int(calendar.monthrange(m1_end.year, m1_end.month)[1])
    m2_end = m1_end.replace(day=1) - datetime.timedelta(days=1)
    m2_days = int(calendar.monthrange(m2_end.year, m2_end.month)[1])
    total = mtd_days + m1_days + m2_days
    return {
        "data_date": data_date,
        "mtd_days": mtd_days,
        "m1_days": m1_days,
        "m2_days": m2_days,
        "m1": m1_end.strftime("%Y-%m"),
        "m2": m2_end.strftime("%Y-%m"),
        "total_days": total,
        "label": f"P{total}D",
    }


def _shop_simplify_site_top_base(site_code: str, ratio=None) -> str:
    """店舖層 Top N% 的基礎欄名（不含口徑後綴）：HA21、80% →「是否HA21的Top 80%」。"""
    return _SHOP_SIMPLIFY_SITE_TOP80_TEMPLATE.format(
        site=str(site_code or "").strip(),
        pct=_shop_simplify_ratio_percent(ratio),
    )


def _shop_simplify_nature_top_base(ratio=None) -> str:
    """同 Shop Nature 的基礎欄名（不含口徑後綴）：80% →「是否同Shop Nature的Top 80%」。"""
    return _SHOP_SIMPLIFY_NATURE_TOP80_TEMPLATE.format(pct=_shop_simplify_ratio_percent(ratio))


def _shop_simplify_value_site_top_column(site_code: str, ratio=None) -> str:
    """金額口徑的店舖層欄名（80% →「是否HA21的Top 80%（金額）」）。"""
    return _shop_simplify_site_top_base(site_code, ratio) + _SHOP_SIMPLIFY_VALUE_SUFFIX


def _shop_simplify_value_nature_top_column(ratio=None) -> str:
    """金額口徑的同 Nature 欄名（80% →「是否同Shop Nature的Top 80%（金額）」）。"""
    return _shop_simplify_nature_top_base(ratio) + _SHOP_SIMPLIFY_VALUE_SUFFIX


def _shop_simplify_net_sales_column(label: str) -> str:
    """金額資料欄名（label＝P73D →「Net Sales (P73D)」）。"""
    return f"Net Sales ({str(label or '').strip()})"


def _is_shop_simplify_value_top_column(name) -> bool:
    """是否金額口徑的 Top 判定欄（欄名以「（金額）」結尾）。"""
    return isinstance(name, str) and _is_shop_simplify_top80_column(name) and name.strip().endswith(
        _SHOP_SIMPLIFY_VALUE_SUFFIX
    )


def _shop_simplify_shop_qty_top_column(ratio=None) -> str:
    """原模式（多店輸出）的件數判定欄名（80% →「是否該店的Top 80%（件數）」）。"""
    return _SHOP_SIMPLIFY_SHOP_TOP_TEMPLATE.format(pct=_shop_simplify_ratio_percent(ratio)) + _SHOP_SIMPLIFY_QTY_SUFFIX


def _shop_simplify_shop_value_top_column(ratio=None) -> str:
    """原模式（多店輸出）的金額判定欄名（80% →「是否該店的Top 80%（金額）」）。"""
    return _SHOP_SIMPLIFY_SHOP_TOP_TEMPLATE.format(pct=_shop_simplify_ratio_percent(ratio)) + _SHOP_SIMPLIFY_VALUE_SUFFIX


def _shop_simplify_shop_compare_column(ratio=None) -> str:
    """原模式（多店輸出）的件數 vs 金額對比欄名（80% →「該店 件數 vs 該店 金額 Top 80%」）。"""
    return _SHOP_SIMPLIFY_SHOP_COMPARE_TEMPLATE.format(pct=_shop_simplify_ratio_percent(ratio))


def _shop_simplify_measure_compare_labels(
    qty_flags: pd.Series, value_flags: pd.Series, available: bool = True
) -> pd.Series:
    """同一間店的件數與金額兩個判定合成四象限；金額不可用時整欄留空。"""
    if not available:
        return pd.Series("", index=qty_flags.index, dtype=object)
    qty = qty_flags.fillna(False).to_numpy(dtype=bool)
    value = value_flags.fillna(False).to_numpy(dtype=bool)
    labels = np.select(
        [qty & value, qty & ~value, ~qty & value],
        [
            _SHOP_SIMPLIFY_COMPARE_BOTH,
            _SHOP_SIMPLIFY_MEASURE_COMPARE_QTY_ONLY,
            _SHOP_SIMPLIFY_MEASURE_COMPARE_VALUE_ONLY,
        ],
        default=_SHOP_SIMPLIFY_COMPARE_NEITHER,
    )
    return pd.Series(labels, index=qty_flags.index, dtype=object)


def _is_shop_simplify_compare_column(name) -> bool:
    """是否對比欄：店舖 vs 同 Nature（件數／金額），或原模式的件數 vs 金額。"""
    if not isinstance(name, str):
        return False
    text = name.strip()
    return " vs 同Nature Top " in text or _is_shop_simplify_measure_compare_column(text)


def _is_shop_simplify_measure_compare_column(name) -> bool:
    """是否原模式的「件數 vs 金額」對比欄（欄名含「件數 vs … 金額 Top」）。"""
    if not isinstance(name, str):
        return False
    text = name.strip()
    return " 件數 vs " in text and " 金額 Top " in text


def _shop_simplify_site_top80_column(site_code: str, ratio=None) -> str:
    """指定店舖模式的件數店舖層欄名（HA21、80% →「是否HA21的Top 80%（件數）」）。"""
    return _shop_simplify_site_top_base(site_code, ratio) + _SHOP_SIMPLIFY_QTY_SUFFIX


def _shop_simplify_nature_top_column(ratio=None) -> str:
    """件數口徑的同 Shop Nature 欄名（80% →「是否同Shop Nature的Top 80%（件數）」）。"""
    return _shop_simplify_nature_top_base(ratio) + _SHOP_SIMPLIFY_QTY_SUFFIX


def _shop_simplify_compare_column(site_code: str, ratio=None, suffix: str = _SHOP_SIMPLIFY_QTY_SUFFIX) -> str:
    """店舖 vs 同 Nature 對比欄名（80%、件數 →「HA21 vs 同Nature Top 80%（件數）」）。"""
    return (
        _SHOP_SIMPLIFY_COMPARE_TEMPLATE.format(
            site=str(site_code or "").strip(),
            pct=_shop_simplify_ratio_percent(ratio),
        )
        + suffix
    )


def _shop_simplify_compare_labels(shop_flags: pd.Series, nature_flags: pd.Series, available: bool = True) -> pd.Series:
    """把兩組布林旗標合成四象限標籤；Nature 不可用時整欄留空。"""
    if not available:
        return pd.Series("", index=shop_flags.index, dtype=object)
    shop = shop_flags.fillna(False).to_numpy(dtype=bool)
    nature = nature_flags.fillna(False).to_numpy(dtype=bool)
    labels = np.select(
        [shop & nature, shop & ~nature, ~shop & nature],
        [
            _SHOP_SIMPLIFY_COMPARE_BOTH,
            _SHOP_SIMPLIFY_COMPARE_SHOP_ONLY,
            _SHOP_SIMPLIFY_COMPARE_NATURE_ONLY,
        ],
        default=_SHOP_SIMPLIFY_COMPARE_NEITHER,
    )
    return pd.Series(labels, index=shop_flags.index, dtype=object)


def _is_shop_simplify_value_compare_column(name) -> bool:
    return _is_shop_simplify_compare_column(name) and name.strip().endswith(_SHOP_SIMPLIFY_VALUE_SUFFIX)


def _is_shop_simplify_top80_column(name) -> bool:
    """是否本模式的 Top N% 判定欄（含指定店舖動態欄名、Nature 欄、70% 與金額變體）。"""
    if not isinstance(name, str):
        return False
    text = name.strip()
    return text.startswith("是否") and re.search(r"Top\s*\d{1,3}%", text) is not None


def _is_shop_simplify_site_visible_column(name) -> bool:
    """指定店舖變體匯出時的預設顯示欄（清單見 `_SHOP_SIMPLIFY_SITE_VISIBLE_COLUMNS`）。

    動態欄名一併納入：Net Sales (P{N}D)、是否{店舖}的Top N%（件數／金額）、
    是否同Shop Nature的Top N%（件數／金額）、{店舖} vs 同Nature Top N%（件數／金額）。
    原模式的「是否該店的Top N%」與「該店 件數 vs 該店 金額」不在此列（指定店舖變體不輸出）。
    """
    if not isinstance(name, str):
        return False
    text = name.strip()
    if text in _SHOP_SIMPLIFY_SITE_VISIBLE_COLUMNS:
        return True
    if _SHOP_SIMPLIFY_SITE_NET_SALES_RE.fullmatch(text):
        return True
    if _is_shop_simplify_compare_column(text):
        return True
    return _is_shop_simplify_top80_column(text) and not text.startswith("是否該店的")


def _shop_simplify_top80_mask(qty: pd.Series, share: float = _SHOP_SIMPLIFY_TOP80_SHARE, tail_guard: bool = True) -> pd.Series:
    """按數值（銷量或金額）由高至低累計，回傳「累計首次達總量 N%」的 SKU 布林標記。

    - 以索引為 SKU 鍵（呼叫方先按 SKU 彙總）；
    - `share` 為累計比例（0.8＝Top 80%、0.7＝Top 70%）；
    - 同一數值的 SKU 整組納入或整組排除，避免同值被任意切開；
    - 總量為 0（或全為空白／負數）時全部回傳 False；
    - 長尾保護（僅件數口徑，`tail_guard=True`）：若達標點落在僅售 1 件的長尾
      （cutoff ≤ 1）且整組納入會令全部分類 100% 入選（即該分類最低銷量亦為 1），
      則改為僅保留 > cutoff（排除 1 件長尾），避免「每間店每個分類全部 SKU 皆為 Top」
      而喪失精簡意義；但若分類內全部 SKU 同為 1 件（無更高銷量者），排除後會一併清空，
      故此時維持整組納入（唯一／同值 SKU 仍屬該分類核心）。金額口徑（`tail_guard=False`）
      不套用——「cutoff ≤ 1」的門檻單位是件，金額達標點落在 $1 內屬正常分佈，套用反而
      會整組排除低單價品項。
    """
    share = _shop_simplify_top_share(share)
    values = pd.to_numeric(qty, errors="coerce").fillna(0.0).clip(lower=0.0)
    mask = pd.Series(False, index=values.index, dtype=bool)
    if values.empty:
        return mask
    total = float(values.sum())
    if total <= 0:
        return mask
    ordered = values.sort_values(ascending=False, kind="stable")
    cumulative = ordered.cumsum().to_numpy(dtype=float)
    reached = np.flatnonzero(cumulative >= total * share)
    if reached.size == 0:
        return mask
    cutoff = float(ordered.iloc[int(reached[0])])
    if (
        tail_guard
        and cutoff <= 1.0
        and bool((values >= cutoff).all())
        and bool((values > cutoff).any())
    ):
        return values > cutoff
    return values >= cutoff


def _shop_simplify_group_choice(group_by=None) -> str:
    """正規化分組選擇（MCH2／MC／Adj. SST，忽略大小寫與前後空白）；不認識時用預設 MCH2。"""
    if group_by is None:
        return _SHOP_SIMPLIFY_TOP80_GROUP_COLUMN
    text = str(group_by).strip().upper()
    for name in _SHOP_SIMPLIFY_GROUP_OPTIONS:
        if text == name.upper():
            return name
    return _SHOP_SIMPLIFY_TOP80_GROUP_COLUMN


def _shop_simplify_group_column(columns, group_by=None) -> tuple[str | None, str]:
    """回傳 (實際分組欄名, 命中的分組名稱)。

    由所選分組開始按退回鏈找第一個存在的欄位（MCH2 → MC → Adj. SST；選 MC 時
    MC → Adj. SST），全部缺欄時回傳 (None, 所選分組)。
    """
    requested = _shop_simplify_group_choice(group_by)
    avail = set(columns)
    for name in _SHOP_SIMPLIFY_GROUP_FALLBACK_CHAIN[requested]:
        resolved = _resolve_column_name(name, avail)
        if resolved is not None:
            return resolved, name
    return None, requested


def _shop_simplify_blank_sku_keys(skus: pd.Series) -> set[str]:
    """SKU「空白鍵」集合：空字串與字串化的 NaN／None（astype(str) 的產物）。

    空白 SKU 列無品項身分，不得計入 Top N% 分組（會令分母虛增、判定線被推移）；
    呼叫方以 `sku in blank_keys` 排除並列 warning。
    """
    keys = skus.fillna("").astype(str).str.strip()
    return set(keys[keys.str.lower().isin(["", "nan", "none"])].unique())


def _shop_simplify_top80_keys(
    groups: pd.Series,
    skus: pd.Series,
    qty: pd.Series,
    share: float = _SHOP_SIMPLIFY_TOP80_SHARE,
    tail_guard: bool = True,
) -> set[tuple[str, str]]:
    """按所選分組欄（MC；缺欄時退回 Adj. SST）分組，回傳各組累計 N% 的 (分類, SKU) 集合。

    每個分類獨立累計，分類之間不互相擠佔；同一 SKU 在同一分類內重複出現時先加總。
    SKU 空白（NaN／空字串）的列不計入分母（該等列的判定一律為非 Top，
    由呼叫方列 warning 提示檢查來源檔）；`tail_guard` 僅件數口徑使用（見
    _shop_simplify_top80_mask）。
    """
    frame = pd.DataFrame({
        "group": groups.fillna("").astype(str).str.strip(),
        "sku": skus.fillna("").astype(str).str.strip(),
        "qty": pd.to_numeric(qty, errors="coerce").fillna(0.0).clip(lower=0.0),
    })
    if frame.empty:
        return set()
    _blank_keys = _shop_simplify_blank_sku_keys(frame["sku"])
    if _blank_keys:
        frame = frame[~frame["sku"].isin(_blank_keys)]
        if frame.empty:
            return set()
    totals = frame.groupby(["group", "sku"], sort=False)["qty"].sum()
    keys: set[tuple[str, str]] = set()
    for group, series in totals.groupby(level=0, sort=False):
        mask = _shop_simplify_top80_mask(series.droplevel(0), share=share, tail_guard=tail_guard)
        keys.update((str(group), str(sku)) for sku in mask.index[mask.to_numpy()])
    return keys


def _shop_simplify_net_sales_series(df: pd.DataFrame, net_columns) -> pd.Series:
    """三個月 Net Sales 相加成金額計算欄；先按原始值相加，最後才將總和下限為 0，
    令跨月退貨可抵銷原銷貨（避免逐月先 clip 為 0 導致退貨月份被忽略）；缺值當 0。

    `net_columns` 為 [(實際欄名, 標準欄名), ...]；實際欄名為 None 時略過該月。
    """
    total = pd.Series(0.0, index=df.index, dtype=float)
    for actual, _standard in net_columns:
        if actual is not None and actual in df.columns:
            total = total.add(_safe_numeric(df[actual]), fill_value=0.0)
    return total.fillna(0.0).clip(lower=0.0)


def _shop_simplify_shop_top_flags(
    df: pd.DataFrame,
    measure: pd.Series,
    site_col: str | None,
    group_col: str | None,
    share: float = _SHOP_SIMPLIFY_TOP80_SHARE,
    tail_guard: bool = True,
) -> pd.Series:
    """每列是否屬「該店 × 該分組」累計 N% 內的 SKU（原模式多店輸出用）。

    以 (Site Code, 分組值) 為分組鍵（分類欄優先 MC，缺欄時退回 Adj. SST），
    分組之間互不擠佔；缺 Site Code 時全表視為單一店舖（"(全部店舖)"），
    缺分類欄時每店單一分組（"(全部)"）。
    `measure` 可為件數（Sold Qty (P90D)）或金額（Net Sales (P{N}D)），兩者共用同一引擎；
    `tail_guard` 長尾保護僅件數口徑使用（金額口徑傳 False）。
    SKU 空白的列不計入分母，判定一律為非 Top。
    """
    index = df.index
    if site_col is not None and site_col in df.columns:
        sites = df[site_col].fillna("").astype(str).str.strip()
        sites = sites.mask(sites == "", "(空白店舖)")
    else:
        sites = pd.Series("(全部店舖)", index=index, dtype=object)
    if group_col is not None and group_col in df.columns:
        groups = df[group_col].fillna("").astype(str).str.strip()
        groups = groups.mask(groups == "", "(空白分類)")
    else:
        groups = pd.Series("(全部)", index=index, dtype=object)
    if "SKU" not in df.columns:
        return pd.Series(False, index=index, dtype=bool)

    # 以不可見分隔符併成單一分組鍵，重用按分類累計 N% 的同一判定
    body_keys = sites.astype(str) + "\u0000" + groups.astype(str)
    skus = _cached_safe_sku_series(df["SKU"]).astype(str).str.strip()
    keys = _shop_simplify_top80_keys(body_keys, skus, measure, share=share, tail_guard=tail_guard)
    return pd.Series(
        [(group, sku) in keys for group, sku in zip(body_keys, skus)],
        index=index,
        dtype=bool,
    )


def _shop_simplify_shop_top80_flags(
    df: pd.DataFrame,
    sold: pd.Series,
    site_col: str | None,
    group_col: str | None,
    share: float = _SHOP_SIMPLIFY_TOP80_SHARE,
) -> pd.Series:
    """件數口徑的包裝（保留舊呼叫端名稱）。"""
    return _shop_simplify_shop_top_flags(df, sold, site_col, group_col, share=share)


def _shop_simplify_drop_existing_columns(df: pd.DataFrame) -> pd.DataFrame:
    """移除來源檔已有的輸出欄（含前後空格變體與指定店舖的 Top 80% 欄），避免重複欄位。

    回傳移除後的 df；是否有欄被移除由呼叫方比對欄數自行判斷。
    """
    avail = set(df.columns)
    drop = []
    for name in [_SHOP_SIMPLIFY_SOLD_COLUMN, _SHOP_SIMPLIFY_STOCK_COLUMN, *_SHOP_SIMPLIFY_REPLY_COLUMNS]:
        resolved = _resolve_column_name(name, avail)
        if resolved is not None:
            drop.append(resolved)
    drop.extend(c for c in df.columns if _is_shop_simplify_top80_column(c) and c not in drop)
    # 店舖 vs 同 Nature 對比欄同樣屬衍生欄，換比例重跑時一併移除
    drop.extend(c for c in df.columns if _is_shop_simplify_compare_column(c) and c not in drop)
    # 舊結果的 Net Sales (PxxD) 屬衍生欄，重跑時一併移除，避免與本次視窗並存
    drop.extend(
        c for c in df.columns
        if isinstance(c, str) and re.fullmatch(r"Net Sales \(P\d+D\)", c.strip()) and c not in drop
    )
    return df.drop(columns=drop) if drop else df


def transform_shop_simplify(
    df_input: pd.DataFrame,
    progress_callback=None,
    shop_natures: list[str] | None = None,
    include_shop_top80: bool = True,
    top_ratio=None,
    include_net_sales: bool = True,
    data_date=None,
    group_by=None,
):
    """清理來源資料並附加「過去90日有無賣過／過去90日有無存貨」、件數與金額兩組
    「是否該店的Top N%」、一欄「該店 件數 vs 該店 金額 Top N%」對比，以及三個空白回覆欄。

    - 兩個判定都以 (Site Code, 分組值) 為分組，每間店每個分類各自一條 N% 線；
      件數按 Sold Qty (P90D)、金額按 Net Sales (P{N}D)（三個月先按原值相加，
      總和下限為 0）。SKU 空白的列不計入 Top N% 分組（判定為非 Top）並列 warning。
      `group_by` 可選 MCH2（預設，中類）／MC（小類）／Adj. SST（陳列分類）；
      所選欄缺欄時按 MCH2 → MC → Adj. SST 退回並列 warning，全缺則每店單一分組。
      對比值為「兩者皆是／僅店舖件數／僅店舖金額／兩者皆非」。
    - `top_ratio` 為累計比例：0.8（預設）或 0.7；亦接受 70、"70%" 等寫法，欄名隨比例改變。
    - `data_date` 為數據截止日（T-1），未提供時用「今天減一日」，金額欄名帶實際覆蓋日數
      （如 `Net Sales (P73D)`）；缺三個月 Net Sales 欄時只輸出件數判定並列 warning。
    - 指定店舖變體另有自己的店舖／同 Nature 判定，故以 `include_shop_top80=False` 及
      `include_net_sales=False` 關閉，避免重複欄位。
    - 回傳 (raw_df, warnings)。缺少 Sold Qty (P90D)，或 Shop Stock Qty (LLM)／(LM)／(CM)
      三個庫存欄同時缺失時拋出 ValueError（無法判定，不應靜默輸出）。
    """
    if progress_callback:
        progress_callback(0.0, "驗證輸入資料...")
    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必需欄位: {', '.join(missing)}")

    avail = set(df_input.columns)
    sold_col = _resolve_column_name("Sold Qty (P90D)", avail)
    llm_col = _resolve_column_name("Shop Stock Qty (LLM)", avail)
    lm_col = _resolve_column_name("Shop Stock Qty (LM)", avail)
    cm_col = _resolve_column_name("Shop Stock Qty (CM)", avail)
    site_col = _resolve_column_name("Site Code", avail)
    group_requested = _shop_simplify_group_choice(group_by)
    group_col, group_used = _shop_simplify_group_column(avail, group_requested)
    share = _shop_simplify_top_share(top_ratio)
    pct = _shop_simplify_ratio_percent(share)
    window = _shop_simplify_net_sales_window(data_date)
    net_column = _shop_simplify_net_sales_column(window["label"])
    net_cols = [(_resolve_column_name(name, avail), name) for name in _SHOP_SIMPLIFY_NET_SALES_COLUMNS]
    net_missing = [name for actual, name in net_cols if actual is None]
    net_available = bool(include_net_sales) and not net_missing
    qty_column = _shop_simplify_shop_qty_top_column(share)
    value_column = _shop_simplify_shop_value_top_column(share)
    compare_column = _shop_simplify_shop_compare_column(share)
    if sold_col is None:
        raise ValueError("來源檔案缺少「Sold Qty (P90D)」，無法判定過去90日有無賣過。")
    _stock_pairs = [
        ("Shop Stock Qty (LLM)", llm_col),
        ("Shop Stock Qty (LM)", lm_col),
        ("Shop Stock Qty (CM)", cm_col),
    ]
    if all(c is None for _, c in _stock_pairs):
        raise ValueError(
            "來源檔案同時缺少「Shop Stock Qty (LLM)」「Shop Stock Qty (LM)」「Shop Stock Qty (CM)」，"
            "無法判定過去90日有無存貨。"
        )

    if progress_callback:
        progress_callback(0.15, "清理資料中...")
    df = _clean_dataframe(df_input, shop_natures=shop_natures, progress_callback=progress_callback).copy()

    warnings: list[str] = []
    _used_stock_names = [n for n, c in _stock_pairs if c is not None]
    for _name, _col in _stock_pairs:
        if _col is None:
            warnings.append(
                f"來源檔案缺少「{_name}」，過去90日有無存貨以「{'」、「'.join(_used_stock_names)}」判定。"
            )

    if include_net_sales and net_missing:
        warnings.append(
            f"來源檔案缺少「{'」「'.join(net_missing)}」，未能計算金額口徑 Top {pct}%，"
            "只輸出件數判定欄（不會加入 Net Sales 與（金額）欄）。"
        )

    # SKU 空白列不計入 Top N% 分組（分母不含該等列；判定一律為非 Top）
    if "SKU" in df.columns:
        _blank_keys = _shop_simplify_blank_sku_keys(df["SKU"])
        if _blank_keys:
            _blank_sku_rows = int(
                df["SKU"].fillna("").astype(str).str.strip().isin(_blank_keys).sum()
            )
            warnings.append(
                f"有 {_blank_sku_rows} 列「SKU」空白，已不計入 Top {pct}% 分母"
                "（該等列的件數／金額判定一律標示為非 Top），請檢查來源檔。"
            )

    had_output_columns = any(
        _resolve_column_name(name, avail) is not None
        for name in [_SHOP_SIMPLIFY_SOLD_COLUMN, _SHOP_SIMPLIFY_STOCK_COLUMN, *_SHOP_SIMPLIFY_REPLY_COLUMNS]
    )
    if had_output_columns:
        warnings.append(
            "來源檔已含本模式的輸出欄（過去90日有無賣過／過去90日有無存貨／"
            "SCT建議／Shop 回覆新的NDRF／如果RF的Safety Qty／Top N% 判定欄），"
            "將以本次計算結果覆蓋，三個回覆欄一律留空。"
        )

    if progress_callback:
        progress_callback(0.35, "判定過去90日有無賣過...")
    df = _shop_simplify_drop_existing_columns(df)

    if "Item_Shop" not in df.columns:
        warnings.append("來源檔案缺少「Item_Shop」，SKU Desc 與 Shop Name 留空。")
    df = _compute_raw_extra_columns(df)
    df = _shop_simplify_place_derived_columns(df)

    value_columns = [c for c in df.columns if _is_raw_convert_value_column(c)]
    if value_columns:
        df = df.drop(columns=value_columns)

    sold = _safe_numeric(df[sold_col]) if sold_col in df.columns else pd.Series(0.0, index=df.index)
    sold_status = np.where(
        sold.to_numpy(dtype=float) > 0,
        _SHOP_SIMPLIFY_SOLD_YES,
        _SHOP_SIMPLIFY_SOLD_NO,
    )

    if progress_callback:
        progress_callback(0.55, "判定過去90日有無存貨...")
    llm_stock = (
        _safe_numeric(df[llm_col]) if llm_col is not None and llm_col in df.columns else pd.Series(0.0, index=df.index)
    )
    lm_stock = (
        _safe_numeric(df[lm_col]) if lm_col is not None and lm_col in df.columns else pd.Series(0.0, index=df.index)
    )
    cm_stock = (
        _safe_numeric(df[cm_col]) if cm_col is not None and cm_col in df.columns else pd.Series(0.0, index=df.index)
    )
    has_stock = (
        (llm_stock.to_numpy(dtype=float) > 0)
        | (lm_stock.to_numpy(dtype=float) > 0)
        | (cm_stock.to_numpy(dtype=float) > 0)
    )
    stock_status = np.where(has_stock, _SHOP_SIMPLIFY_STOCK_YES, _SHOP_SIMPLIFY_STOCK_NO)

    df[_SHOP_SIMPLIFY_SOLD_COLUMN] = sold_status
    df[_SHOP_SIMPLIFY_STOCK_COLUMN] = stock_status

    net = None
    if net_available:
        net = _shop_simplify_net_sales_series(df, net_cols)
        df[net_column] = net.to_numpy(dtype=float)

    if include_shop_top80:
        if progress_callback:
            progress_callback(0.62, f"判定各店的Top {pct}%（件數／金額）...")
        _flag_text = (
            f"「{qty_column}」與「{value_column}」" if net_available else f"「{qty_column}」"
        )
        if site_col is None:
            warnings.append(
                f"來源檔案缺少「Site Code」，{_flag_text}未能按店舖分組，"
                "改以全表為單一店舖計算。"
            )
        if group_col is None:
            warnings.append(
                f"來源檔案缺少「MCH2」「MC」「Adj. SST」，{_flag_text}未能按品類分組，"
                "改以每店單一分組計算。"
            )
        elif group_used != group_requested:
            warnings.append(
                f"來源檔案缺少「{group_requested}」，{_flag_text}改以「{group_used}」分類計算。"
            )
        qty_flags = _shop_simplify_shop_top_flags(df, sold, site_col, group_col, share=share)
        df[qty_column] = np.where(
            qty_flags.to_numpy(dtype=bool),
            _SHOP_SIMPLIFY_TOP80_YES,
            _SHOP_SIMPLIFY_TOP80_NO,
        )
        if net_available:
            value_flags = _shop_simplify_shop_top_flags(df, net, site_col, group_col, share=share, tail_guard=False)
            df[value_column] = np.where(
                value_flags.to_numpy(dtype=bool),
                _SHOP_SIMPLIFY_TOP80_YES,
                _SHOP_SIMPLIFY_TOP80_NO,
            )
            df[compare_column] = _shop_simplify_measure_compare_labels(qty_flags, value_flags)

            # 逐店金額合計為 0 屬預期內的退貨／零動銷情況，彙總成一條 warning 避免逐店洗版
            if site_col is not None and site_col in df.columns:
                site_totals = df[site_col].fillna("").astype(str).str.strip()
            else:
                site_totals = pd.Series("(全部店舖)", index=df.index, dtype=object)
            totals = pd.DataFrame({"site": site_totals, "value": net}).groupby("site")["value"].sum()
            zero_sites = int((totals <= 0).sum())
            if zero_sites:
                warnings.append(
                    f"有 {zero_sites} 間店的 {net_column} 合計為 0（或全為負數），"
                    f"該等店的「{value_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
                )

    for name in _SHOP_SIMPLIFY_REPLY_COLUMNS:
        df[name] = ""

    # 欄序：判定欄緊接「過去90日有無存貨」（件數 → 金額 → 對比），
    # 金額計算欄緊接 Sold Qty (P90D)，三個回覆欄維持最後
    if include_shop_top80:
        cols = list(df.columns)
        insert_at = cols.index(_SHOP_SIMPLIFY_STOCK_COLUMN) + 1
        flag_columns = [qty_column] + ([value_column, compare_column] if net_available else [])
        for name in flag_columns:
            cols.remove(name)
            cols.insert(insert_at, name)
            insert_at += 1
        if net_available:
            cols.remove(net_column)
            cols.insert(cols.index(sold_col) + 1, net_column)
        df = df[cols]

    # SKU 統一為文字（與其他模式一致，避免 Excel 以科學記號顯示長編號）
    if "SKU" in df.columns:
        df["SKU"] = _cached_safe_sku_series(df["SKU"])

    # Launch Date 統一為 YYYYMMDD 整數（Excel 顯示為數字，與範本一致）
    if "Launch Date" in df.columns:
        df["Launch Date"] = pd.Series(
            [_shop_simplify_launch_int(v) for v in df["Launch Date"]],
            index=df.index,
            dtype=object,
        )

    # 非有限數值（如 0 銷售時 Turnday 的 inf）一律留空，避免 Excel 出現 inf
    for col in df.columns:
        if pd.api.types.is_float_dtype(df[col]):
            arr = df[col].to_numpy(dtype=float)
            if np.isinf(arr).any():
                df[col] = df[col].where(np.isfinite(arr))

    # 排序：Brand、SKU（文字升序），與店舖精簡回覆表範本一致
    if "Brand" in df.columns and "SKU" in df.columns:
        df = df.sort_values(
            ["Brand", "SKU"],
            key=lambda s: s.astype(str).str.strip(),
            kind="stable",
        ).reset_index(drop=True)

    if df.empty:
        warnings.append("清理後沒有可輸出的資料列，請確認來源檔案與 Shop Nature 篩選。")

    if progress_callback:
        progress_callback(0.7, "完成資料轉換...")
    return df, warnings


def transform_shop_simplify_site(
    df_input: pd.DataFrame,
    site_code: str,
    progress_callback=None,
    top_ratio=None,
    include_net_sales: bool = True,
    data_date=None,
    group_by=None,
    sct_rules=None,
):
    """指定店舖模式：只輸出該店舖（Site Code）的 SKU 列，並加兩組（件數／金額）Top N% 判定欄。

    - 件數：「是否{店舖}的Top N%（件數）」與「是否同Shop Nature的Top N%（件數）」，按 Sold Qty (P90D)；
- 金額：「是否{店舖}的Top N%（金額）」與「是否同Shop Nature的Top N%（金額）」，
  按 Net Sales (P{N}D)＝`Net Sales (M-2)`＋`Net Sales (M-1)`＋`Net Sales (CM - MTD)`
  （三個月先按原值相加，總和下限為 0），N 由 MTD 日數＋兩個整月日數決定（例如 11 月 12 日 → 73 日）。
- 對比：`{店舖} vs 同Nature Top N%（件數）`／`（金額）`，值為「兩者皆是／僅店舖／
  僅同Nature／兩者皆非」，直接看兩邊是否同為核心；Nature 不可用時整欄留空。

    兩種口徑都按所選分組（`group_by`：MCH2 中類（預設）／MC 小類／Adj. SST 陳列分類），
    在每個分類內累計至該分類 N%，同一銷量／金額整組納入或排除；分類之間不互相擠佔。
    來源缺 Shop Nature 時 Nature 欄留空；所選分組欄缺失時按 MCH2 → MC → Adj. SST 退回並列
    warning，全缺則退回單一分組；缺三個月 Net Sales 欄時只輸出件數判定，均列 warning。

    `top_ratio` 為累計比例：0.8（預設）或 0.7；亦接受 70、"70%" 等寫法。欄名隨比例改變，
    故切換比例重跑不會留下舊欄。`data_date` 為數據截止日（T-1），未提供時用「今天減一日」；
    金額欄名會帶實際覆蓋日數（如 `Net Sales (P73D)`）。

    `sct_rules` 為 SCT建議的情況表（情況鍵 → 首選RF／次選RF／建議RF／留空），
    提供時按其填寫「SCT建議」欄，未提供則維持留空（原本行為）。

    其餘清理、Item_Shop 解析、刪除 Value 欄、三個回覆欄留空與排序規則完全沿用
    `transform_shop_simplify`（以 `include_shop_top80=False`、`include_net_sales=False` 關閉
    原模式的店舖判定與金額欄，避免重複）。回傳 (raw_df, warnings)。
    """
    site = str(site_code or "").strip()
    if not site:
        raise ValueError("未提供指定店舖（Site Code），無法執行指定店舖模式。")
    share = _shop_simplify_top_share(top_ratio)
    pct = _shop_simplify_ratio_percent(share)
    window = _shop_simplify_net_sales_window(data_date)
    net_column = _shop_simplify_net_sales_column(window["label"])

    missing = validate_input(df_input)
    if missing:
        raise ValueError(f"缺少必需欄位: {', '.join(missing)}")

    avail = set(df_input.columns)
    site_col = _resolve_column_name("Site Code", avail)
    if site_col is None:
        raise ValueError("來源檔案缺少「Site Code」，無法指定店舖。")
    sku_col = _resolve_column_name("SKU", avail)
    sold_col = _resolve_column_name("Sold Qty (P90D)", avail)
    if sku_col is None or sold_col is None:
        raise ValueError(f"來源檔案缺少「SKU」或「Sold Qty (P90D)」，無法計算 Top {pct}%。")
    group_requested = _shop_simplify_group_choice(group_by)
    group_col, group_used = _shop_simplify_group_column(avail, group_requested)
    nature_col = _resolve_column_name("Shop Nature", avail)

    net_cols = [(_resolve_column_name(name, avail), name) for name in _SHOP_SIMPLIFY_NET_SALES_COLUMNS]
    net_missing = [name for actual, name in net_cols if actual is None]
    net_available = bool(include_net_sales) and not net_missing

    if progress_callback:
        progress_callback(0.02, "清理來源資料（全表）...")
    # 全表只做向量化清理：逐列 Item_Shop 解析留到篩選店舖後才做，
    # 80 萬行時可避免對全表逐列解析與排序，大幅節省時間與記憶體。
    df_clean = _clean_dataframe(df_input, shop_natures=None, progress_callback=progress_callback)

    warnings: list[str] = []
    sites = df_clean[site_col].fillna("").astype(str).str.strip()
    exact = sites == site
    if exact.any():
        matched_site = site
    else:
        case_insensitive = sites.str.upper() == site.upper()
        if not case_insensitive.any():
            raise ValueError(
                f"指定店舖「{site}」在清理後沒有資料列，請確認 Site Code 是否正確，"
                "或來源檔是否只包含部分店舖。"
            )
        matched_site = str(sites[case_insensitive].iloc[0]).strip()
        warnings.append(f"指定店舖「{site}」按大小寫不敏感比對為「{matched_site}」，以下按此店舖計算。")

    shop_raw = df_clean[sites == matched_site].copy()
    if shop_raw.empty:
        raise ValueError(f"指定店舖「{matched_site}」在清理後沒有資料列，無法計算 Top {pct}%。")

    site_top_column = _shop_simplify_site_top80_column(matched_site, share)
    nature_column = _shop_simplify_nature_top_column(share)
    value_site_column = _shop_simplify_value_site_top_column(matched_site, share)
    value_nature_column = _shop_simplify_value_nature_top_column(share)

    if include_net_sales and net_missing:
        warnings.append(
            f"來源檔案缺少「{'」「'.join(net_missing)}」，未能計算金額口徑 Top {pct}%，"
            "只輸出件數判定欄（不會加入 Net Sales 與（金額）欄）。"
        )

    # SKU 空白列不計入 Top N% 分組（分母不含該等列；判定一律為非 Top）
    if "SKU" in shop_raw.columns:
        _blank_keys = _shop_simplify_blank_sku_keys(shop_raw["SKU"])
        if _blank_keys:
            _blank_sku_rows = int(
                shop_raw["SKU"].fillna("").astype(str).str.strip().isin(_blank_keys).sum()
            )
            warnings.append(
                f"有 {_blank_sku_rows} 列「SKU」空白，已不計入 Top {pct}% 分母"
                "（該等列的件數／金額判定一律標示為非 Top），請檢查來源檔。"
            )

    if group_col is None:
        warnings.append(
            f"來源檔案缺少「MCH2」「MC」「Adj. SST」，Top {pct}% 未能按品類分組，"
            "改以全店／全 Nature 單一分組計算。"
        )
    elif group_used != group_requested:
        warnings.append(
            f"來源檔案缺少「{group_requested}」，Top {pct}% 改以「{group_used}」分類計算。"
        )

    # ── 同 Shop Nature 的 Top N%（件數與金額）：先在清理後全表按 (分組值, SKU) 算好，
    #    稍後映射到指定店舖的列；其他 Nature 不計入分母 ──
    nature_available = False
    nature_value = ""
    nature_keys: set[tuple[str, str]] = set()
    nature_value_keys: set[tuple[str, str]] = set()
    nature_shop_count = 0
    nature_total = 0.0
    nature_net_total = 0.0
    if nature_col is None:
        warnings.append(
            "來源檔案缺少「Shop Nature」，未能判定「"
            f"{nature_column}」，該欄留空。"
        )
    else:
        row_natures = shop_raw[nature_col].fillna("").astype(str).str.strip()
        row_natures = row_natures[row_natures != ""]
        if row_natures.empty:
            warnings.append(
                f"指定店舖「{matched_site}」的 Shop Nature 空白，未能判定「"
                f"{nature_column}」，該欄留空。"
            )
        else:
            nature_value = str(row_natures.mode().iloc[0])
            nature_available = True
            if row_natures.nunique() > 1:
                warnings.append(
                    f"指定店舖「{matched_site}」的 Shop Nature 有多個值"
                    f"（{'、'.join(sorted(row_natures.unique()))}），同 Nature 判定以「{nature_value}」計算。"
                )

            nature_series = df_clean[nature_col].fillna("").astype(str).str.strip()
            nature_df = df_clean[nature_series == nature_value]
            nature_sku = _cached_safe_sku_series(nature_df[sku_col]).astype(str).str.strip()
            nature_qty = _safe_numeric(nature_df[sold_col])
            if group_col is None:
                nature_group = pd.Series("(全部)", index=nature_df.index)
            else:
                nature_group = nature_df[group_col].fillna("").astype(str).str.strip()
                nature_group = nature_group.mask(nature_group == "", "(空白)")
            nature_keys = _shop_simplify_top80_keys(nature_group, nature_sku, nature_qty, share=share)
            if net_available:
                nature_net = _shop_simplify_net_sales_series(nature_df, net_cols)
                nature_value_keys = _shop_simplify_top80_keys(
                    nature_group, nature_sku, nature_net, share=share, tail_guard=False
                )
                nature_net_total = float(nature_net.sum())
                del nature_net

            nature_sites = nature_df[site_col].fillna("").astype(str).str.strip()
            nature_sites = nature_sites[nature_sites != ""]
            nature_shop_count = int(nature_sites.nunique())
            nature_total = float(nature_qty.sum())
            del nature_df, nature_series, nature_sku, nature_qty, nature_group, nature_sites

    # 全表清理結果只留到此處，及早釋放（80 萬行可省數 GB）
    del df_clean, sites
    gc.collect()

    def _site_progress(fraction, msg):
        if progress_callback:
            progress_callback(0.2 + 0.65 * float(fraction), msg)

    if progress_callback:
        progress_callback(0.2, f"整理 {matched_site} 的店舖×SKU 明細...")
    shop_df, base_warnings = transform_shop_simplify(
        shop_raw,
        progress_callback=_site_progress if progress_callback else None,
        shop_natures=None,
        include_shop_top80=False,
        include_net_sales=False,
    )
    warnings = list(base_warnings) + warnings

    if progress_callback:
        progress_callback(0.88, f"計算 {matched_site} 的 Top {pct}%（件數＋金額）...")
    shop_sku = _cached_safe_sku_series(shop_df[sku_col]).astype(str).str.strip()
    shop_qty = _safe_numeric(shop_df[sold_col])
    if group_col is None:
        shop_group = pd.Series("(全部)", index=shop_df.index)
    else:
        shop_group = shop_df[group_col].fillna("").astype(str).str.strip()
        shop_group = shop_group.mask(shop_group == "", "(空白)")

    shop_keys = _shop_simplify_top80_keys(shop_group, shop_sku, shop_qty, share=share)
    shop_flags = pd.Series(
        [(group, sku) in shop_keys for group, sku in zip(shop_group, shop_sku)],
        index=shop_df.index,
        dtype=bool,
    )
    nature_flags = pd.Series(
        [(group, sku) in nature_keys for group, sku in zip(shop_group, shop_sku)]
        if nature_available else False,
        index=shop_df.index,
        dtype=bool,
    )

    if float(shop_qty.sum()) <= 0:
        warnings.append(
            f"指定店舖「{matched_site}」的 Sold Qty (P90D) 合計為 0，"
            f"「{site_top_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
        )
    if nature_available:
        if nature_shop_count <= 1:
            warnings.append(
                f"同 Shop Nature「{nature_value}」在來源檔只有 1 間店舖（即指定店舖），"
                f"「{nature_column}」與店舖 Top {pct}% 結果相同。"
            )
        if nature_total <= 0:
            warnings.append(
                f"同 Shop Nature「{nature_value}」的 Sold Qty (P90D) 合計為 0，"
                f"「{nature_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
            )

    shop_df[site_top_column] = np.where(shop_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO)
    shop_df[nature_column] = (
        np.where(nature_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO)
        if nature_available else ""
    )
    compare_qty_column = _shop_simplify_compare_column(matched_site, share, _SHOP_SIMPLIFY_QTY_SUFFIX)
    shop_df[compare_qty_column] = _shop_simplify_compare_labels(shop_flags, nature_flags, nature_available)

    shop_net = None
    if net_available:
        shop_net = _shop_simplify_net_sales_series(shop_df, net_cols)
        shop_value_keys = _shop_simplify_top80_keys(shop_group, shop_sku, shop_net, share=share, tail_guard=False)
        shop_value_flags = pd.Series(
            [(group, sku) in shop_value_keys for group, sku in zip(shop_group, shop_sku)],
            index=shop_df.index,
            dtype=bool,
        )
        nature_value_flags = pd.Series(
            [(group, sku) in nature_value_keys for group, sku in zip(shop_group, shop_sku)]
            if nature_available else False,
            index=shop_df.index,
            dtype=bool,
        )
        if float(shop_net.sum()) <= 0:
            warnings.append(
                f"指定店舖「{matched_site}」的 {net_column} 合計為 0，"
                f"「{value_site_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
            )
        if nature_available and nature_net_total <= 0:
            warnings.append(
                f"同 Shop Nature「{nature_value}」的 {net_column} 合計為 0，"
                f"「{value_nature_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
            )

        shop_df[net_column] = shop_net.to_numpy(dtype=float)
        shop_df[value_site_column] = np.where(
            shop_value_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO
        )
        shop_df[value_nature_column] = (
            np.where(nature_value_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO)
            if nature_available else ""
        )
        compare_value_column = _shop_simplify_compare_column(matched_site, share, _SHOP_SIMPLIFY_VALUE_SUFFIX)
        shop_df[compare_value_column] = _shop_simplify_compare_labels(
            shop_value_flags, nature_value_flags, nature_available
        )

    # SCT建議：按四個判定欄與使用者指定的情況表自動填寫（未指定時維持留空）
    if sct_rules:
        _sct_columns = [site_top_column, nature_column, value_site_column, value_nature_column]
        if not net_available:
            # 金額口徑不可用時，金額兩項一律視為非 Top（情況鍵尾兩碼為 N）
            _sct_columns = _sct_columns[:2] + [None, None]
        shop_df[_SHOP_SIMPLIFY_SCT_COLUMN] = _shop_simplify_sct_series(shop_df, sct_rules, _sct_columns)
        _sct_labels = shop_df[_SHOP_SIMPLIFY_SCT_COLUMN].replace("", np.nan).dropna()
        if len(_sct_labels):
            warnings.append(
                f"已按 SCT建議規則填寫「{_SHOP_SIMPLIFY_SCT_COLUMN}」欄，"
                f"共 {len(_sct_labels):,} 列（可按實際情況在 Excel 內修改）。"
            )
        else:
            warnings.append(
                "SCT建議規則沒有任何情況命中，該欄維持空白；請檢查情況表是否有指定建議。"
            )

    # 判定欄緊接「過去90日有無存貨」（件數在前、金額在後），三個回覆欄維持最後；
    # 每個口徑依次為：店舖 Top、同 Nature Top、兩者對比；Net Sales (PxxD) 緊接 Sold Qty (P90D)
    cols = list(shop_df.columns)
    insert_at = cols.index(_SHOP_SIMPLIFY_STOCK_COLUMN) + 1
    flag_columns = [site_top_column, nature_column, compare_qty_column]
    if net_available:
        flag_columns += [value_site_column, value_nature_column, compare_value_column]
    for name in flag_columns:
        cols.remove(name)
        cols.insert(insert_at, name)
        insert_at += 1
    if net_available:
        cols.remove(net_column)
        cols.insert(cols.index(sold_col) + 1, net_column)
    shop_df = shop_df[cols].reset_index(drop=True)

    if progress_callback:
        progress_callback(0.85, f"完成 {matched_site} 的 Top {pct}% 判定...")
    return shop_df, warnings


def _shop_simplify_cell(value):
    """Excel 儲存格值正規化：空白字串與各類缺值一律寫成 None（真正空白）；
    浮點數保留完整精度（不四捨五入），inf／-inf 視為空白。"""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, str):
        return value if value != "" else None
    if isinstance(value, np.bool_):
        return bool(value)
    if isinstance(value, np.integer):
        return int(value)
    if isinstance(value, (float, np.floating)):
        number = float(value)
        return number if np.isfinite(number) else None
    return value


def _shop_simplify_cell_mask(series: pd.Series, target: str) -> np.ndarray:
    """逐列判定值是否等於目標值（去除前後空白、忽略大小寫；空白／缺值不命中）。

    紅色底（Item Status＝Discontinued／RP Type＝ND／無存貨）與 Top N% 判定欄的
    綠色底（＝「是」）共用同一比對規則。
    """
    text = series.fillna("").astype(str).str.strip().str.lower()
    return text.eq(str(target).strip().lower()).to_numpy(dtype=bool)


def _shop_simplify_summary_levels(series: pd.Series) -> pd.Series:
    """Summary 的分組標籤：去除前後空白，空白／缺值一律成「(空白)」。"""
    values = series.fillna("").astype(str).str.strip()
    return values.mask(values == "", "(空白)")


def _shop_simplify_summary_row_order(levels: pd.Series) -> list[str]:
    """分組次序：文字升序，空白組排最後（Excel 樞紐慣例）。"""
    order = sorted({v for v in levels if v != "(空白)"})
    if bool((levels == "(空白)").any()):
        order.append("(空白)")
    return order


def _shop_simplify_summary_counts(
    levels: pd.Series, masks: list[tuple[str, pd.Series]], valid: pd.Series
) -> dict[tuple[str, str], int]:
    """按 (分組標籤, 欄標籤) 統計 SKU 列數（只計 SKU 非空白的列）。"""
    frame = pd.DataFrame({"level": levels.to_numpy()})
    for label, mask in masks:
        frame[label] = (mask.fillna(False) & valid).to_numpy(dtype=int)
    grouped = frame.groupby("level", sort=False).sum()
    return {
        (str(level), str(label)): int(grouped.at[level, label])
        for level in grouped.index
        for label in grouped.columns
    }


def _shop_simplify_summary_block(
    ws, title: str, start_row: int, row_levels: list[str],
    column_labels: list[str], counts: dict[tuple[str, str], int],
) -> tuple[int, int]:
    """寫一個「列標籤 × 欄標籤」的 SKU 計數區塊（零值留空、底部總計列）。

    回傳 (最後一列列號, 用到的最大欄號)。
    """
    ws.cell(row=start_row, column=1, value=title)
    header = start_row + 1
    ws.cell(row=header, column=1, value="列標籤")
    total_col = 2 + len(column_labels)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for j, label in enumerate(column_labels, start=2):
        cell = ws.cell(row=header, column=j, value=label)
        cell.alignment = header_align
    ws.cell(row=header, column=total_col, value="總計").alignment = header_align
    ws.row_dimensions[header].height = 30

    column_totals = dict.fromkeys(column_labels, 0)
    last = header
    for i, level in enumerate(row_levels, start=header + 1):
        cell = ws.cell(row=i, column=1, value=level)
        cell.alignment = Alignment(horizontal="left")
        row_total = 0
        for j, label in enumerate(column_labels, start=2):
            count = counts.get((level, label), 0)
            if count:
                ws.cell(row=i, column=j, value=count)
            row_total += count
            column_totals[label] += count
        if row_total:
            ws.cell(row=i, column=total_col, value=row_total)
        last = i

    total_row = last + 1
    ws.cell(row=total_row, column=1, value="總計")
    for j, label in enumerate(column_labels, start=2):
        if column_totals[label]:
            ws.cell(row=total_row, column=j, value=column_totals[label])
    grand = sum(column_totals.values())
    if grand:
        ws.cell(row=total_row, column=total_col, value=grand)
    return total_row, total_col


def _shop_simplify_summary_top_columns(df: pd.DataFrame) -> list[str]:
    """Summary 要拆分的 Top N% 判定欄，順序為：店舖件數／同Nature件數／店舖金額／同Nature金額。

    欄名直接取自輸出（如「是否HA21的Top 80%（件數）」），故 Top 70%／80% 與店舖代號
    自動跟隨；金額口徑不可用時只有兩欄。
    """
    tops = [c for c in df.columns if _is_shop_simplify_top80_column(c)]

    def key(name: str):
        text = str(name).strip()
        is_nature = text.startswith("是否同Shop Nature")
        is_value = text.endswith(_SHOP_SIMPLIFY_VALUE_SUFFIX)
        return (1 if is_value else 0, 1 if is_nature else 0)

    return sorted(tops, key=key)


def _write_shop_simplify_summary(ws, df: pd.DataFrame) -> None:
    """Summary 工作表：以 Adj. MCH1 為列，逐個維度統計 SKU 列數。

    區塊次序：
    1. × Item Status（原有版面，固定於 A3／A4）；
    1b. Item Status × RP Type（每個狀態再分 RF／ND，另有其他值時加「其他」）；
    2. × RP Type（RF／ND／其他）；
    2b. × SCT建議（僅在該欄有值時輸出：首選RF／次選RF／建議RF）；
    3. × 四個 Top N% 判定欄（店舖件數／同Nature件數／店舖金額／同Nature金額，
       欄名取實際輸出欄，故 Top 70%／80% 與店舖代號自動跟隨）；
    4. RF／ND × 每個 Top N% 判定欄的四象限（店舖件數／同Nature件數／店舖金額／同Nature金額
       各一個區塊，欄為 RF-Top／RF-非Top／ND-Top／ND-非Top；非Top 含「否」及未判定）。
    零值一律留空，每個區塊底部有總計列；來源缺該欄時該區塊自動略過。
    """
    index_col = "Adj. MCH1" if "Adj. MCH1" in df.columns else (
        "Item Class" if "Item Class" in df.columns else None
    )
    if df is None or df.empty or index_col is None or "SKU" not in df.columns:
        return

    levels = _shop_simplify_summary_levels(df[index_col])
    row_order = _shop_simplify_summary_row_order(levels)
    valid = df["SKU"].notna()

    # ── 區塊 1：× Item Status（原有版面，測試與使用者習慣依賴 A3／A4 位置）──
    status_col = "Item Status" if "Item Status" in df.columns else None
    status_labels: list[str] = []
    blocks_end = 2
    max_col = 0
    if status_col is not None:
        ws["A3"] = "計數 - SKU"
        ws["B3"] = "欄標籤"

        summary_df = df.copy()
        summary_df[index_col] = levels
        status_values = _shop_simplify_summary_levels(df[status_col])
        summary_df[status_col] = status_values

        table = summary_df.pivot_table(
            index=index_col,
            columns=status_col,
            values="SKU",
            aggfunc="count",
            margins=True,
            margins_name="總計",
            fill_value=0,
        )
        # 空白組排在最後（Excel 樞紐慣例），總計列／欄置末
        statuses = [c for c in table.columns if c != "總計"]
        statuses = [c for c in statuses if c != "(空白)"] + [c for c in statuses if c == "(空白)"]
        total_col = 2 + len(statuses)
        order = [v for v in table.index if v not in ("總計", "(空白)")]
        if "(空白)" in table.index:
            order.append("(空白)")
        if "總計" in table.index:
            order.append("總計")
        table = table.reindex(
            index=order,
            columns=statuses + (["總計"] if "總計" in table.columns else []),
        )

        ws.cell(row=4, column=1, value="列標籤")
        for j, status in enumerate(statuses, start=2):
            ws.cell(row=4, column=j, value=str(status))
        ws.cell(row=4, column=total_col, value="總計")

        for i, (idx, row) in enumerate(table.iterrows(), start=5):
            cell = ws.cell(row=i, column=1, value=str(idx))
            cell.alignment = Alignment(horizontal="left")
            for j, status in enumerate(statuses, start=2):
                count = int(row[status])
                if count:
                    ws.cell(row=i, column=j, value=count)
            total = int(row["總計"])
            if total:
                ws.cell(row=i, column=total_col, value=total)
        status_labels = [str(s) for s in statuses]
        blocks_end = 4 + len(table)
        max_col = total_col

    # ── 區塊 1b：Item Status × RP Type（每個狀態再分 RF／ND，另有其他值時加「其他」）──
    if status_labels and "RP Type" in df.columns:
        # RF／ND 比對忽略大小寫（與 Export 的紅色底判定一致，"nd" 亦計入 ND）
        rp = _shop_simplify_summary_levels(df["RP Type"]).str.upper()
        rp_other = ~rp.isin(["RF", "ND"])
        rp_labels = ["RF", "ND"] + (["其他"] if bool(rp_other.any()) else [])
        status_levels = _shop_simplify_summary_levels(df[status_col])
        labels = []
        masks = []
        for status in status_labels:
            in_status = status_levels.eq(status)
            for rp_label in rp_labels:
                mask = in_status & (rp_other if rp_label == "其他" else rp.eq(rp_label))
                label = f"{status} - {rp_label}"
                labels.append(label)
                masks.append((label, mask))
        counts = _shop_simplify_summary_counts(levels, masks, valid)
        blocks_end, _ = _shop_simplify_summary_block(
            ws, "計數 - SKU：Item Status × RP Type", blocks_end + 3, row_order, labels, counts
        )
        max_col = max(max_col, 2 + len(labels))

    # ── 區塊 2：× RP Type ──
    if "RP Type" in df.columns:
        rp = _shop_simplify_summary_levels(df["RP Type"]).str.upper()
        labels = ["RF", "ND"]
        masks = [(label, rp.eq(label)) for label in labels]
        other = ~rp.isin(["RF", "ND"])
        if bool(other.any()):
            labels.append("其他")
            masks.append(("其他", other))
        counts = _shop_simplify_summary_counts(levels, masks, valid)
        blocks_end, _ = _shop_simplify_summary_block(
            ws, "計數 - SKU：RP Type", blocks_end + 3, row_order, labels, counts
        )
        max_col = max(max_col, 2 + len(labels))

    # ── 區塊 2b：× SCT建議（僅在該欄有值時輸出，方便核對自動建議的分佈）──
    if _SHOP_SIMPLIFY_SCT_COLUMN in df.columns:
        sct = df[_SHOP_SIMPLIFY_SCT_COLUMN].fillna("").astype(str).str.strip()
        if bool((sct != "").any()):
            known = [v for v in _SHOP_SIMPLIFY_SCT_OPTIONS if v]
            labels = [v for v in known if bool(sct.eq(v).any())]
            labels += sorted({v for v in sct.unique() if v and v not in known})
            masks = [(label, sct.eq(label)) for label in labels]
            counts = _shop_simplify_summary_counts(levels, masks, valid)
            blocks_end, _ = _shop_simplify_summary_block(
                ws,
                f"計數 - SKU：{_SHOP_SIMPLIFY_SCT_COLUMN}",
                blocks_end + 3,
                row_order,
                labels,
                counts,
            )
            max_col = max(max_col, 2 + len(labels))

    # ── 區塊 3：× 四個 Top N% 判定欄 ──
    top_columns = _shop_simplify_summary_top_columns(df)
    for name in top_columns:
        flags = _shop_simplify_summary_levels(df[name])
        labels = ["是", "否"]
        masks = [("是", flags.eq(_SHOP_SIMPLIFY_TOP80_YES)), ("否", flags.eq(_SHOP_SIMPLIFY_TOP80_NO))]
        other = ~flags.isin([_SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO])
        if bool(other.any()):
            labels.append("其他")
            masks.append(("其他", other))
        counts = _shop_simplify_summary_counts(levels, masks, valid)
        blocks_end, _ = _shop_simplify_summary_block(
            ws, f"計數 - SKU：{name}", blocks_end + 3, row_order, labels, counts
        )
        max_col = max(max_col, 2 + len(labels))

    # ── 區塊 4：RF／ND × 各 Top 判定欄的四象限（件數與金額、店舖與同 Nature 各一個）──
    if "RP Type" in df.columns:
        rp = _shop_simplify_summary_levels(df["RP Type"])
        for name in top_columns:
            # Top＝「是」；非Top＝「否」及未判定（空白／其他值），與區塊標題一致
            top = _shop_simplify_summary_levels(df[name]).eq(_SHOP_SIMPLIFY_TOP80_YES)
            labels = ["RF - Top", "RF - 非Top", "ND - Top", "ND - 非Top"]
            masks = [
                ("RF - Top", rp.eq("RF") & top),
                ("RF - 非Top", rp.eq("RF") & ~top),
                ("ND - Top", rp.eq("ND") & top),
                ("ND - 非Top", rp.eq("ND") & ~top),
            ]
            other = ~rp.isin(["RF", "ND"])
            if bool(other.any()):
                labels += ["其他 - Top", "其他 - 非Top"]
                masks += [("其他 - Top", other & top), ("其他 - 非Top", other & ~top)]
            counts = _shop_simplify_summary_counts(levels, masks, valid)
            blocks_end, _ = _shop_simplify_summary_block(
                ws,
                f"計數 - SKU：RF／ND × {name}（非Top 含「否」及未判定）",
                blocks_end + 3,
                row_order,
                labels,
                counts,
            )
            max_col = max(max_col, 2 + len(labels))

    # 欄寬（對齊範本；區塊欄數多於 4 個時其餘用 12）
    widths = [35.28515625, 9.7109375, 12.7109375, 8.28515625, 10.0, 5.28515625]
    while len(widths) < max_col:
        # 交叉區塊的欄名較長（如「Discontinued - ND」），配 15 並自動換行
        widths.append(15.0)
    for j, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = width


def generate_excel_shop_simplify(raw: pd.DataFrame, progress_callback=None, site_view: bool = False) -> BytesIO:
    """產生兩個工作表：Summary（Adj. MCH1 × Item Status 計數）與 Export（回覆明細）。

    Export 沿用店舖精簡回覆表範本的格式：Brand／SKU 排序、凍結 F2、指定欄寬與隱藏欄、
    數字格式、綠色／紫色重點欄、「無賣」列的淡紅標示、三類紅色底標示
    （Item Status＝Discontinued、RP Type＝ND、過去90日有無存貨＝無存貨），
    以及對比欄底色（兩者皆是＝青綠、兩者皆非＝粉紅）。

    `site_view=True`（指定店舖變體）：只顯示 `_SHOP_SIMPLIFY_SITE_VISIBLE_COLUMNS` 所列的
    23 個固定欄，加上 Net Sales (P{N}D) 與六個 Top 判定／對比欄；其餘欄一律隱藏，
    令匯出檔一打開就是回覆表需要的檢視（欄序、顏色、隱藏欄以外的格式完全不變）。
    """
    wb = Workbook()
    ws_summary = wb.active
    ws_summary.title = _SHOP_SIMPLIFY_SUMMARY_SHEET
    ws = wb.create_sheet(title=_SHOP_SIMPLIFY_EXPORT_SHEET)

    if progress_callback:
        progress_callback(0.0, "建立 Excel 結構...")

    df = raw if raw is not None else pd.DataFrame()
    headers = list(df.columns)
    # SCT建議 有任何值（自動規則命中）時保持顯示；全空則沿用「預設顯示欄」檢視把該欄隱藏
    sct_filled = (
        _SHOP_SIMPLIFY_SCT_COLUMN in df.columns
        and bool((df[_SHOP_SIMPLIFY_SCT_COLUMN].fillna("").astype(str).str.strip() != "").any())
    )
    ws.append([_shop_simplify_cell(c) for c in headers])

    total = len(df)
    for i, row in enumerate(df.itertuples(index=False, name=None)):
        ws.append([_shop_simplify_cell(v) for v in row])
        if progress_callback and (i % 2000 == 0 or i == total - 1):
            progress_callback(0.05 + 0.75 * ((i + 1) / total), f"寫入 Export ({i+1}/{total})...")

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    fill_blue = PatternFill(start_color=_SHOP_SIMPLIFY_HEADER_BLUE, end_color=_SHOP_SIMPLIFY_HEADER_BLUE, fill_type="solid")
    fill_green = PatternFill(start_color=_SHOP_SIMPLIFY_GREEN, end_color=_SHOP_SIMPLIFY_GREEN, fill_type="solid")
    fill_purple = PatternFill(start_color=_SHOP_SIMPLIFY_PURPLE, end_color=_SHOP_SIMPLIFY_PURPLE, fill_type="solid")
    fill_pink = PatternFill(start_color=_SHOP_SIMPLIFY_NO_SALE_PINK, end_color=_SHOP_SIMPLIFY_NO_SALE_PINK, fill_type="solid")
    fill_red = PatternFill(start_color=_SHOP_SIMPLIFY_RED, end_color=_SHOP_SIMPLIFY_RED, fill_type="solid")
    fill_compare_both = PatternFill(
        start_color=_SHOP_SIMPLIFY_COMPARE_BOTH_FILL, end_color=_SHOP_SIMPLIFY_COMPARE_BOTH_FILL, fill_type="solid"
    )
    fill_compare_neither = PatternFill(
        start_color=_SHOP_SIMPLIFY_COMPARE_NEITHER_FILL,
        end_color=_SHOP_SIMPLIFY_COMPARE_NEITHER_FILL,
        fill_type="solid",
    )

    # 表頭：粗體、指定底色、細框、置中置頂、自動換行；列高 70.5
    ws.row_dimensions[1].height = 70.5
    for col_idx, name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        if name in _SHOP_SIMPLIFY_GREEN_HEADERS:
            cell.fill = fill_green
        elif name in _SHOP_SIMPLIFY_PURPLE_COLUMNS:
            cell.fill = fill_purple
        else:
            cell.fill = fill_blue
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

        letter = get_column_letter(col_idx)
        width = _SHOP_SIMPLIFY_COLUMN_WIDTHS.get(name)
        if width is None and _is_shop_simplify_compare_column(name):
            # 對比欄名較長（HA21 vs 同Nature Top 80%（件數））
            width = 26.0
        if width is None and _is_shop_simplify_top80_column(name):
            # 指定店舖變體的判定欄較長（…Top 80%（件數）／（金額））；原模式的「是否該店的Top 80%」不加闊
            text = str(name).strip()
            is_nature = text.startswith("是否同Shop Nature的Top ")
            if text.endswith(_SHOP_SIMPLIFY_VALUE_SUFFIX) or text.endswith(_SHOP_SIMPLIFY_QTY_SUFFIX):
                width = 26.0 if is_nature else 22.0
            else:
                width = 16.0 if is_nature else 14.0
        if width is None and isinstance(name, str) and re.fullmatch(r"Net Sales \(P\d+D\)", name.strip()):
            # 金額計算欄（Net Sales (P73D)）
            width = 14.0
        if width is not None:
            ws.column_dimensions[letter].width = width
        if name in _SHOP_SIMPLIFY_HIDDEN_COLUMNS:
            ws.column_dimensions[letter].hidden = True
        elif site_view and not (
            _is_shop_simplify_site_visible_column(name)
            or (sct_filled and name == _SHOP_SIMPLIFY_SCT_COLUMN)
        ):
            # 指定店舖變體：不在預設顯示清單內的欄（Shop Name、OM、三個月庫存／銷量明細、
            # Shop Stock Status 等）一律隱藏，需要時在 Excel 內取消隱藏即可；
            # 例外：SCT建議 一經填寫（自動規則命中）便保持顯示，方便對照
            ws.column_dimensions[letter].hidden = True

    # 資料格：綠色重點欄、紫色 Sold Qty (P90D)、「無賣」淡紅標示、Top 判定欄「是」綠色底、
    # 以及紅色底重點（Item Status＝Discontinued／RP Type＝ND／過去90日有無存貨＝無存貨）
    green_cols = [i for i, name in enumerate(headers, start=1) if name in _SHOP_SIMPLIFY_GREEN_DATA_COLUMNS]
    purple_cols = [i for i, name in enumerate(headers, start=1) if name in _SHOP_SIMPLIFY_PURPLE_COLUMNS]
    sold_col_idx = headers.index(_SHOP_SIMPLIFY_SOLD_COLUMN) + 1 if _SHOP_SIMPLIFY_SOLD_COLUMN in headers else None
    red_rules = [
        (headers.index(name) + 1, _shop_simplify_cell_mask(df[name], target))
        for name, target in _SHOP_SIMPLIFY_RED_RULES
        if name in headers and name in df.columns
    ]
    # Top N% 判定欄（是否{店舖}的Top N%／是否同Shop Nature的Top N%，件數與金額）：
    # 值為「是」時以綠色底標示，方便一眼掃出核心 SKU；「否」不著色
    top_yes_rules = [
        (headers.index(name) + 1, _shop_simplify_cell_mask(df[name], _SHOP_SIMPLIFY_TOP80_YES))
        for name in headers
        if _is_shop_simplify_top80_column(name) and name in df.columns
    ]
    # 對比欄：兩者皆是＝青綠底、兩者皆非＝粉紅底，其餘（僅店舖／僅同Nature）不著色
    compare_rules = []
    for name in headers:
        if not _is_shop_simplify_compare_column(name) or name not in df.columns:
            continue
        text = df[name].fillna("").astype(str).str.strip()
        compare_rules.append((
            headers.index(name) + 1,
            text.eq(_SHOP_SIMPLIFY_COMPARE_BOTH).to_numpy(dtype=bool),
            text.eq(_SHOP_SIMPLIFY_COMPARE_NEITHER).to_numpy(dtype=bool),
        ))
    number_formats = [
        (i, _SHOP_SIMPLIFY_NUMBER_FORMATS[name])
        for i, name in enumerate(headers, start=1)
        if name in _SHOP_SIMPLIFY_NUMBER_FORMATS
    ]
    # 金額計算欄名隨視窗改變（Net Sales (P73D)），按樣式套用與三個月 Net Sales 相同的數字格式
    number_formats += [
        (i, _SHOP_SIMPLIFY_NUMBER_FORMATS["Net Sales (M-2)"])
        for i, name in enumerate(headers, start=1)
        if isinstance(name, str) and re.fullmatch(r"Net Sales \(P\d+D\)", name.strip())
    ]
    for row_idx in range(2, total + 2):
        for col_idx in green_cols:
            ws.cell(row=row_idx, column=col_idx).fill = fill_green
        for col_idx in purple_cols:
            ws.cell(row=row_idx, column=col_idx).fill = fill_purple
        if sold_col_idx is not None:
            sold_cell = ws.cell(row=row_idx, column=sold_col_idx)
            if sold_cell.value == _SHOP_SIMPLIFY_SOLD_NO:
                sold_cell.fill = fill_pink
        for col_idx, mask in top_yes_rules:
            if mask[row_idx - 2]:
                ws.cell(row=row_idx, column=col_idx).fill = fill_green
        for col_idx, both_mask, neither_mask in compare_rules:
            if both_mask[row_idx - 2]:
                ws.cell(row=row_idx, column=col_idx).fill = fill_compare_both
            elif neither_mask[row_idx - 2]:
                ws.cell(row=row_idx, column=col_idx).fill = fill_compare_neither
        # 紅色底最後套用，故 RP Type＝ND 會蓋過該欄的綠色
        for col_idx, mask in red_rules:
            if mask[row_idx - 2]:
                ws.cell(row=row_idx, column=col_idx).fill = fill_red
        for col_idx, fmt in number_formats:
            ws.cell(row=row_idx, column=col_idx).number_format = fmt

    # 凍結於 SKU Desc 之後（範本為 F2）；縮放 85%
    if "SKU Desc" in headers:
        ws.freeze_panes = f"{get_column_letter(headers.index('SKU Desc') + 2)}2"
    elif headers:
        ws.freeze_panes = "A2"
    ws.sheet_view.zoomScale = 85
    ws.auto_filter.ref = ws.dimensions

    if progress_callback:
        progress_callback(0.85, "建立 Summary 工作表...")
    _write_shop_simplify_summary(ws_summary, df)

    # 開啟檔案時預設顯示 Export（範本行為）
    wb.active = wb.sheetnames.index(_SHOP_SIMPLIFY_EXPORT_SHEET)

    if progress_callback:
        progress_callback(0.95, "正在儲存 Excel 檔案...")
    gc.collect()
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)

    if progress_callback:
        progress_callback(1.0, "完成!")
    return buf
