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
  自動填寫 SCT建議，其餘兩個回覆欄維持留空；Buyer＝123 - Cara Tsoi 的 ND 貨品不套用
  情況表，改標示「需由Buyer決定」（該 Buyer 可能有受限制 SKU，例如只限澳門出售）。

N 可選 80%（預設）或 70%；金額欄名隨數據截止日自動標示實際覆蓋日數（如 P73D），
緊接 Sold Qty (P90D) 之後。本模式不作調貨、Safety 或退倉建議，亦不篩選列
（全店全 SKU 輸出）；清理規則沿用 `_clean_dataframe`（剔除合計列與 117 - Marketing），
Shop Nature 篩選為選填。
另由 Item_Shop 以 Site Code 解析 SKU Desc 與 Shop Name（SKU Desc 排於 Brand 之後、
Shop Name 排於 Site Code 之後），其餘來源欄序不變；欄名含「Value」的欄位一律刪除
（大小寫敏感，與「RAW 明細轉化 (不含 Value)」同一判定，故 Valuation Class 不受影響）。
輸出經 Brand／SKU（文字）排序，並產生 說明（SCT建議 情況表與顏色說明）、
Summary（Adj. MCH1 × Item Status 計數）與 Export（回覆明細）三個工作表；Export 套用店舖精簡回覆表範本的欄寬、隱藏欄、
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

`site_code` 亦接受店碼清單（如 ["HA21", "HA47"]）：所選店舖的銷量／金額合併成一個池計算
四條判定（視為同一區域），輸出包含所選店舖的全部 SKU 列——適合新舖銷量未夠、單店計
Top N% 會失真時，借用同區老舖一併做分母。欄名用自訂群組名（`group_label`，如「柴灣區」），
未提供時以「+」串接實際店碼（如 HA21+HA47）；同 Shop Nature 判定按每列自己的 Shop Nature
計算，故群組跨 Nature 亦不會混淆。單一店舖時欄名與行為與舊版完全相同。

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
    _SHOP_BLOCKED_QTY_COLUMN,
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
    "_SHOP_SIMPLIFY_SCT_SKIP_STATUS",
    "_SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER",
    "_SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE",
    "_SHOP_SIMPLIFY_SCT_BUYER_ND_FILL",
    "_shop_simplify_buyer_nd_mask",
    "_shop_simplify_sct_mark_buyer_nd",
    "_SHOP_SIMPLIFY_SCT_BLANK_LABEL",
    "_SHOP_SIMPLIFY_SCT_FILLS",
    "_shop_simplify_sct_skip_discontinued",
    "_shop_simplify_sct_case_text",
    "_shop_simplify_sct_case_compare",
    "_shop_simplify_sct_series",
    "_SHOP_SIMPLIFY_SCT_SHOP_CASES",
    "_SHOP_SIMPLIFY_SCT_SHOP_DEFAULT",
    "_shop_simplify_sct_shop_case_text",
    "_shop_simplify_sct_shop_series",
    "_SHOP_SIMPLIFY_NDRF_COLUMN",
    "_SHOP_SIMPLIFY_SAFETY_COLUMN",
    "_SHOP_SIMPLIFY_REPLY_COLUMNS",
    "_SHOP_SIMPLIFY_SOLD_YES",
    "_SHOP_SIMPLIFY_SOLD_NO",
    "_SHOP_SIMPLIFY_STOCK_YES",
    "_SHOP_SIMPLIFY_STOCK_NO",
    "_SHOP_SIMPLIFY_EXPORT_SHEET",
    "_SHOP_SIMPLIFY_SUMMARY_SHEET",
    "_SHOP_SIMPLIFY_GUIDE_SHEET",
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
    "_SHOP_SIMPLIFY_PRIVATE_LABEL_KEYWORD",
    "_SHOP_SIMPLIFY_PRIVATE_LABEL_FILL",
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
    "_shop_simplify_site_codes",
    "_shop_simplify_site_label",
    "_shop_simplify_site_scope_text",
    "_shop_simplify_site_codes_text",
    "_SHOP_SIMPLIFY_SITE_SEPARATOR",
    "_shop_simplify_nature_top_column",
    "_shop_simplify_value_site_top_column",
    "_shop_simplify_value_nature_top_column",
    "_shop_simplify_net_sales_column",
    "_shop_simplify_now_hkt",
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
# 多店舖群組：預設以「+」串接店碼作為判定欄名（如 HA21+HA47），亦可由介面自訂群組名
_SHOP_SIMPLIFY_SITE_SEPARATOR = "+"
_SHOP_SIMPLIFY_SITE_LABEL_MAX_CODES = 5
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
# Item Status＝Discontinued 的列一律留空，不套用 SCT 規則（由 SCT 另行判斷）
_SHOP_SIMPLIFY_SCT_SKIP_STATUS = "Discontinued"
# Buyer 123 - Cara Tsoi 旗下的 SKU 有機會屬受限制貨品（例如只限澳門出售），
# 其 RP Type＝ND 的列不自動建議轉 RF，改為標示「需由Buyer決定」，由 Buyer 拍板
_SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER = "123 - Cara Tsoi"
_SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE = "需由Buyer決定"
_SHOP_SIMPLIFY_SCT_BUYER_ND_FILL = "FFF2CC"
# Summary 內 SCT建議 分類的標註顏色（與 16 情況設定表一致；留空用灰）
_SHOP_SIMPLIFY_SCT_BLANK_LABEL = "（留空）"
_SHOP_SIMPLIFY_SCT_FILLS = {
    _SHOP_SIMPLIFY_SCT_PREFERRED: "CCFF99",
    _SHOP_SIMPLIFY_SCT_SECONDARY: "FFEB9C",
    _SHOP_SIMPLIFY_SCT_RF: "FFCC99",
    _SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE: _SHOP_SIMPLIFY_SCT_BUYER_ND_FILL,
    _SHOP_SIMPLIFY_SCT_BLANK_LABEL: "D9D9D9",
}


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


def _shop_simplify_sct_skip_discontinued(df: pd.DataFrame, series: pd.Series) -> tuple[pd.Series, int]:
    """Item Status＝Discontinued 的列一律留空，回傳（新序列, 被留空的列數）。

    停售品不論 Top 判定如何都不預設 RF／ND，交由 SCT 按實際情況判斷；
    比對忽略大小寫與前後空白（與 Export 紅色底判定同一規則）。
    """
    if "Item Status" not in df.columns or not len(df):
        return series, 0
    skip = _shop_simplify_cell_mask(df["Item Status"], _SHOP_SIMPLIFY_SCT_SKIP_STATUS)
    count = int(skip.sum())
    if not count:
        return series, 0
    result = series.astype(object).copy()
    result.loc[skip] = ""
    return result, count


def _shop_simplify_buyer_nd_mask(df: pd.DataFrame) -> np.ndarray:
    """Buyer＝123 - Cara Tsoi 且 RP Type＝ND 的列。

    比對忽略大小寫、前後空白與「-」兩側空白（例如「123-Cara Tsoi」亦命中）；
    缺 Buyer 或 RP Type 欄時一律不命中。
    """
    if not len(df) or "Buyer" not in df.columns or "RP Type" not in df.columns:
        return np.zeros(len(df), dtype=bool)
    buyer = df["Buyer"].fillna("").astype(str).str.strip().str.lower()
    buyer = buyer.str.replace(r"\s*-\s*", " - ", regex=True)
    target = re.sub(r"\s*-\s*", " - ", _SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER.strip().lower())
    is_buyer = buyer.eq(target).to_numpy(dtype=bool)
    is_nd = _shop_simplify_cell_mask(df["RP Type"], "ND")
    return is_buyer & is_nd


def _shop_simplify_sct_mark_buyer_nd(df: pd.DataFrame, series: pd.Series) -> tuple[pd.Series, int]:
    """把 Buyer 123 - Cara Tsoi 的 ND 貨品改寫為「需由Buyer決定」，回傳（新序列, 標示列數）。

    這批 SKU 有機會屬受限制貨品（例如只限澳門出售），不宜由 SCT 自行建議轉 RF；
    標示後仍以 Item Status＝Discontinued 的留空規則為準（停售品一律留空）。
    """
    mask = _shop_simplify_buyer_nd_mask(df)
    count = int(mask.sum())
    if not count:
        return series, 0
    result = series.astype(object).copy()
    result.loc[mask] = _SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE
    return result, count


# 原模式（多店輸出）只有「件數 Top」「金額 Top」兩個維度，故全部情況＝2×2＝4 種
_SHOP_SIMPLIFY_SCT_SHOP_CASES = ("YY", "YN", "NY", "NN")
_SHOP_SIMPLIFY_SCT_SHOP_DEFAULT = {
    "YY": _SHOP_SIMPLIFY_SCT_PREFERRED,
    "YN": _SHOP_SIMPLIFY_SCT_SECONDARY,
    "NY": _SHOP_SIMPLIFY_SCT_RF,
    "NN": "",
}


def _shop_simplify_sct_shop_case_text(case: str) -> str:
    """原模式情況鍵（件數、金額）→ 人看的說明。"""
    return {
        "YY": "件數與金額都是 Top",
        "YN": "只有件數 Top",
        "NY": "只有金額 Top",
        "NN": "件數與金額都不是 Top",
    }.get(case, case)


def _shop_simplify_sct_shop_series(df: pd.DataFrame, rules, qty_column, value_column=None) -> pd.Series:
    """原模式：按「是否該店的Top N%（件數／金額）」兩個判定產生 SCT建議。

    金額欄缺失時一律視為非 Top（情況鍵尾碼為 N）；規則未涵蓋的情況留空。
    """
    mapping = {str(k).strip(): str(v or "").strip() for k, v in dict(rules).items()}
    qty = (
        _shop_simplify_cell_mask(df[qty_column], _SHOP_SIMPLIFY_TOP80_YES)
        if qty_column in df.columns
        else np.zeros(len(df), dtype=bool)
    )
    value = (
        _shop_simplify_cell_mask(df[value_column], _SHOP_SIMPLIFY_TOP80_YES)
        if value_column and value_column in df.columns
        else np.zeros(len(df), dtype=bool)
    )
    keys = pd.Series(
        np.char.add(np.where(qty, "Y", "N").astype(str), np.where(value, "Y", "N").astype(str)),
        index=df.index,
    )
    return keys.map(mapping).fillna("").astype(object)


# ── Excel 輸出格式（對齊店舖精簡回覆表範本）──
_SHOP_SIMPLIFY_EXPORT_SHEET = "Export"
_SHOP_SIMPLIFY_SUMMARY_SHEET = "Summary"

# 固定表頭顏色（使用者範本）
_SHOP_SIMPLIFY_HEADER_BLUE = "D9E2F3"
_SHOP_SIMPLIFY_GREEN = "CCFF99"
_SHOP_SIMPLIFY_PURPLE = "E4DFEC"
_SHOP_SIMPLIFY_NO_SALE_PINK = "F2DCDB"

# ── Summary 版面（美化）：顏色分類、格線與欄寬 ──
# 標題深藍、表頭分兩種（計數＝藍、佔比＝灰）、框線淺灰，數值置中、總計列加粗。
_SHOP_SIMPLIFY_SUMMARY_TITLE_COLOUR = "1F3864"
_SHOP_SIMPLIFY_SUMMARY_COUNT_HEADER_FILL = _SHOP_SIMPLIFY_HEADER_BLUE
_SHOP_SIMPLIFY_SUMMARY_PERCENT_HEADER_FILL = "E7E6E6"
# 欄標籤顏色分類（與 Export 同一語言）：綠＝核心（Top／RF）、淡紅＝停售或 ND、灰＝非 Top。
# 個別區塊另有 column_fills（SCT建議 分類）時以該色為先。
_SHOP_SIMPLIFY_SUMMARY_COLUMN_FILLS = {
    "是": _SHOP_SIMPLIFY_GREEN,
    "否": "F2F2F2",
    "RF": _SHOP_SIMPLIFY_GREEN,
    "ND": "FFC7CE",
    "Discontinued": "FFC7CE",
    "RF - Top": _SHOP_SIMPLIFY_GREEN,
    "RF - 非Top": "EAF7DD",
    "ND - Top": "FFC7CE",
    "ND - 非Top": "FBE2E5",
    "其他 - Top": _SHOP_SIMPLIFY_GREEN,
    "其他 - 非Top": "F2F2F2",
}
_SHOP_SIMPLIFY_SUMMARY_GRID_COLOUR = "BFBFBF"
_SHOP_SIMPLIFY_SUMMARY_GRID_STRONG_COLOUR = "808080"
_SHOP_SIMPLIFY_SUMMARY_MIN_WIDTH = 11.0
_SHOP_SIMPLIFY_SUMMARY_MAX_WIDTH = 22.0
_SHOP_SIMPLIFY_SUMMARY_MIN_LABEL_WIDTH = 32.0
_SHOP_SIMPLIFY_SUMMARY_MAX_LABEL_WIDTH = 48.0


def _shop_simplify_summary_text_width(text) -> float:
    """估算顯示寬度（中日韓字元以 1.9 個英文字元計），只用於換算 Summary 欄寬。"""
    return sum(1.9 if ord(ch) > 0x2E80 else 1.0 for ch in str(text))


def _shop_simplify_summary_fill(colour: str) -> PatternFill:
    return PatternFill(start_color=colour, end_color=colour, fill_type="solid")


def _shop_simplify_summary_column_fill(label) -> str | None:
    """欄標籤的分類底色；沒有對應分類時回傳 None（改用區塊表頭底色）。

    交叉區塊的欄名如「Active - ND」「Discontinued - RF」，按結尾的 RF／ND 上色，
    與 `_SHOP_SIMPLIFY_SUMMARY_COLUMN_FILLS` 的單值欄名同一套顏色語言。
    """
    text = str(label).strip()
    colour = _SHOP_SIMPLIFY_SUMMARY_COLUMN_FILLS.get(text)
    if colour:
        return colour
    if text.endswith(" - ND"):
        return _SHOP_SIMPLIFY_SUMMARY_COLUMN_FILLS["ND"]
    if text.endswith(" - RF"):
        return _SHOP_SIMPLIFY_SUMMARY_COLUMN_FILLS["RF"]
    return None


def _shop_simplify_summary_border(strong_top: bool = False) -> Border:
    """格線：四面細框；總計列另加一條較粗的上框線（不填色，與分類底色分開）。"""
    thin = Side(style="thin", color=_SHOP_SIMPLIFY_SUMMARY_GRID_COLOUR)
    top = (
        Side(style="medium", color=_SHOP_SIMPLIFY_SUMMARY_GRID_STRONG_COLOUR)
        if strong_top
        else thin
    )
    return Border(left=thin, right=thin, top=top, bottom=thin)

# 紅色底（醒目紅底；比「無賣」的淡紅 F2DCDB 明顯，故兩者不會混淆）
_SHOP_SIMPLIFY_RED = "FF9999"
# 對比欄底色：兩者皆是＝青綠（核心，優先保留）；兩者皆非＝粉紅（精簡候選）
_SHOP_SIMPLIFY_COMPARE_BOTH_FILL = "99E6CC"
_SHOP_SIMPLIFY_COMPARE_NEITHER_FILL = "FFC7CE"
# 「Adj. SST」含指定字眼（03 - Private Label）＝Private Label 貨品，以醒目橙色底標示
_SHOP_SIMPLIFY_PRIVATE_LABEL_KEYWORD = "03 - private label"
_SHOP_SIMPLIFY_PRIVATE_LABEL_FILL = "FFC000"
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


def shop_simplify_site_stats(
    df: pd.DataFrame, site_code, top_ratio=None, group_by=None, group_label=None
) -> dict:
    """指定店舖模式的結果頁指標：列數、SKU 數、分組數（所選 MCH2／MC／Adj. SST）、
    件數與金額兩組 Top N% 命中數、兩項皆非、無賣／無存貨。

    `site_code` 可為單一店碼或店碼清單；多店舖時 `site` 為判定欄使用的群組名
    （自訂名稱或「+」串接的店碼），`sites` 為結果檔實際出現的店碼清單，
    `site_count` 為店舖數。單一店舖時 `site` 仍是該店碼，既有介面與測試不變。

    `top_ratio` 只在判定欄不存在時作後備；正常情況由欄名（「…Top 70%」）取回實際比例，
    故同一份結果不論介面當前選哪個比例，統計都反映該檔實際使用的比例。
    金額口徑只在結果檔含「（金額）」判定欄時提供，否則 `value_available` 為 False。
    """
    codes = _shop_simplify_site_codes(site_code)
    stats = {
        "rows": 0, "skus": 0, "group_count": 0, "group_column": "",
        "sold_no": 0, "stock_no": 0,
        "site_top80": 0, "nature_top80": 0, "neither": 0,
        "value_site_top80": 0, "value_nature_top80": 0, "value_neither": 0,
        "value_available": False, "window": "",
        "compare_qty": {}, "compare_value": {},
        "site": _shop_simplify_site_label(codes, group_label),
        "sites": list(codes), "site_count": len(codes),
        "nature": "", "ratio": _shop_simplify_ratio_percent(top_ratio),
    }
    if df is None or len(df) == 0:
        return stats

    # 結果檔實際出現的店碼（保持首次出現次序；大小寫視為同一間店），作為標籤與店舖數的依據
    if "Site Code" in df.columns:
        present_sites: list[str] = []
        seen_upper: set[str] = set()
        for value in df["Site Code"].fillna("").astype(str).str.strip().tolist():
            text = str(value).strip()
            if not text or text.upper() in seen_upper:
                continue
            seen_upper.add(text.upper())
            present_sites.append(text)
        if present_sites:
            stats["sites"] = present_sites
            stats["site_count"] = len(present_sites)
            stats["site"] = _shop_simplify_site_label(present_sites, group_label)
    site = stats["site"]

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
            unique_natures = [str(v) for v in dict.fromkeys(natures.tolist()) if str(v).strip()]
            stats["nature"] = "／".join(unique_natures[:3]) + ("…" if len(unique_natures) > 3 else "")
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


# ── 時間：一律以香港時間（HKT）標示 ──
# 部署環境（容器／雲端）的系統時區多為 UTC，直接呼叫 `datetime.now()` 會比香港慢 8 小時，
# 令說明頁的產生時間與「今天」都出現偏差。香港自 1979 年起不設夏令時間，固定 UTC+8 已足夠，
# 亦不依賴 tzdata（Windows 或精簡映像可能未安裝）。
_SHOP_SIMPLIFY_HKT_TZ = datetime.timezone(datetime.timedelta(hours=8), "HKT")


def _shop_simplify_now_hkt() -> datetime.datetime:
    """回傳目前的香港時間（UTC+8，tz-aware）；請勿改用 `datetime.now()`。"""
    return datetime.datetime.now(_SHOP_SIMPLIFY_HKT_TZ)


def _shop_simplify_net_sales_window(data_date=None) -> dict:
    """回傳金額覆蓋日數的拆解：MTD 日數＋M-1 月日數＋M-2 月日數。

    來源數據為 T-1，故未指定時以「今天（香港時間）減一日」為數據截止日。例如 11 月 13 日執行
    → 截止 11 月 12 日：MTD 12 日＋10 月 31 日＋9 月 30 日＝73 日，標示 `P73D`。
    """
    if data_date is None:
        data_date = _shop_simplify_now_hkt().date() - datetime.timedelta(days=1)
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


def _shop_simplify_site_codes(site_code) -> list[str]:
    """把 `site_code`（單一店碼、店碼清單或任何可迭代物件）正規化為店碼清單。

    去除前後空白、略過空白項、按輸入次序去重（大小寫不敏感的重複亦只保留第一個），
    故同一店碼的不同大小寫寫法不會重複計入。未提供時回傳空清單，由呼叫方決定錯誤訊息。
    """
    if site_code is None:
        return []
    if isinstance(site_code, str):
        raw = [site_code]
    else:
        try:
            raw = list(site_code)
        except TypeError:
            raw = [site_code]
    codes: list[str] = []
    seen: set[str] = set()
    for value in raw:
        text = str(value if value is not None else "").strip()
        if not text or text.upper() in seen:
            continue
        seen.add(text.upper())
        codes.append(text)
    return codes


def _shop_simplify_site_label(codes, group_label=None) -> str:
    """判定欄用的「店舖／群組」名稱。

    - 單一店舖：一律用實際店碼（例如 HA21），故既有欄名與統計完全不變；
    - 多店舖：用自訂群組名（例如「柴灣區」），未提供時以「+」串接店碼（例如 HA21+HA47）。
    """
    values = [str(code).strip() for code in (codes or []) if str(code).strip()]
    if not values:
        return ""
    if len(values) == 1:
        return values[0]
    label = str(group_label or "").strip()
    return label or _SHOP_SIMPLIFY_SITE_SEPARATOR.join(values)


def _shop_simplify_site_scope_text(codes, label=None) -> str:
    """warning／說明用的人讀範圍：單店「指定店舖「HA21」」、
    多店「指定店舖群組「柴灣區」（2 間：HA21、HA47）」。
    """
    values = [str(code).strip() for code in (codes or []) if str(code).strip()]
    if not values:
        return "指定店舖"
    if len(values) == 1:
        return f"指定店舖「{values[0]}」"
    name = str(label or "").strip() or _SHOP_SIMPLIFY_SITE_SEPARATOR.join(values)
    return f"指定店舖群組「{name}」（{len(values)} 間：{'、'.join(values)}）"


def _shop_simplify_site_codes_text(codes, limit: int = _SHOP_SIMPLIFY_SITE_LABEL_MAX_CODES) -> str:
    """店碼清單的短標示（去重、保留首次出現次序；超過上限時以「…」收尾），
    供 Summary／說明頁概況使用（輸入可為逐列值，故先自行去重）。"""
    values = [str(code).strip() for code in dict.fromkeys(codes or []) if str(code).strip()]
    if len(values) <= limit:
        return "、".join(values)
    return "、".join(values[:limit]) + "…"


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
    sct_rules=None,
    apply_shop_blocked: bool = True,
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
    - `sct_rules` 為 SCT建議 的 4 情況表（情況鍵 "YY"／"YN"／"NY"／"NN" →
      首選RF／次選RF／建議RF／留空）；提供時按其填寫「SCT建議」欄，否則維持留空。
      Item Status＝Discontinued 的列一律留空，不套用規則（由 SCT 自行判斷）。
    - `apply_shop_blocked` 預設 True：清理時由 `Shop Stock Qty (CM)` 扣走
      `Shop Blocked Qty (CM)` 並按比例調減 `Shop Stock Value (CM)`；
      指定店舖變體已在外層扣減一次，內層呼叫須設 False 避免重複扣減。
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
    df = _clean_dataframe(
        df_input,
        shop_natures=shop_natures,
        progress_callback=progress_callback,
        apply_shop_blocked=apply_shop_blocked,
    ).copy()

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

    # SCT建議：按 4 種情況自動填寫（未指定時維持留空）
    if sct_rules and include_shop_top80 and qty_column in df.columns:
        _sct_series, _ = _shop_simplify_sct_mark_buyer_nd(
            df,
            _shop_simplify_sct_shop_series(
                df, sct_rules, qty_column, value_column if net_available else None
            ),
        )
        _sct_filled, _sct_skipped = _shop_simplify_sct_skip_discontinued(df, _sct_series)
        df[_SHOP_SIMPLIFY_SCT_COLUMN] = _sct_filled
        _sct_buyer_nd = int(
            _shop_simplify_cell_mask(_sct_filled, _SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE).sum()
        )
        _sct_labels = df[_SHOP_SIMPLIFY_SCT_COLUMN].replace("", np.nan).dropna()
        if len(_sct_labels):
            _sct_notes = []
            if _sct_skipped:
                _sct_notes.append(
                    f"Item Status＝{_SHOP_SIMPLIFY_SCT_SKIP_STATUS} 的 {_sct_skipped:,} 列一律留空"
                )
            if _sct_buyer_nd:
                _sct_notes.append(
                    f"Buyer＝{_SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER} 且 RP Type＝ND 的 "
                    f"{_sct_buyer_nd:,} 列標示「{_SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE}」"
                )
            _sct_note = f"（{'；'.join(_sct_notes)}）" if _sct_notes else ""
            warnings.append(
                f"已按 SCT建議規則填寫「{_SHOP_SIMPLIFY_SCT_COLUMN}」欄，"
                f"共 {len(_sct_labels):,} 列{_sct_note}（可按實際情況在 Excel 內修改）。"
            )
        elif _sct_skipped:
            warnings.append(
                f"全部列均為 Item Status＝{_SHOP_SIMPLIFY_SCT_SKIP_STATUS}，"
                f"「{_SHOP_SIMPLIFY_SCT_COLUMN}」欄一律留空，由 SCT 自行判斷。"
            )
        else:
            warnings.append(
                "SCT建議規則沒有任何情況命中，該欄維持空白；請檢查情況表是否有指定建議。"
            )

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
    site_code,
    progress_callback=None,
    top_ratio=None,
    include_net_sales: bool = True,
    data_date=None,
    group_by=None,
    sct_rules=None,
    group_label=None,
):
    """指定店舖模式：只輸出指定店舖的 SKU 列，並加兩組（件數／金額）Top N% 判定欄。

    `site_code` 可為單一店碼（如 "HA21"）或店碼清單（如 ["HA21", "HA47"]）。多於一間店時，
    所選店舖的銷量／金額會合併成一個池（視為同一區域），四條判定欄都用合併後的池計算，
    輸出則包含所選店舖的全部 SKU 列——適合新舖銷量未夠、單店計 Top N% 會失真時使用。

    - 件數：「是否{店舖／群組}的Top N%（件數）」與「是否同Shop Nature的Top N%（件數）」，按 Sold Qty (P90D)；
- 金額：「是否{店舖／群組}的Top N%（金額）」與「是否同Shop Nature的Top N%（金額）」，
  按 Net Sales (P{N}D)＝`Net Sales (M-2)`＋`Net Sales (M-1)`＋`Net Sales (CM - MTD)`
  （三個月先按原值相加，總和下限為 0），N 由 MTD 日數＋兩個整月日數決定（例如 11 月 12 日 → 73 日）。
- 對比：`{店舖／群組} vs 同Nature Top N%（件數）`／`（金額）`，值為「兩者皆是／僅店舖／
  僅同Nature／兩者皆非」，直接看兩邊是否同為核心；Nature 不可用時該列的對比欄留空。

    兩種口徑都按所選分組（`group_by`：MCH2 中類（預設）／MC 小類／Adj. SST 陳列分類），
    在每個分類內累計至該分類 N%，同一銷量／金額整組納入或排除；分類之間不互相擠佔。
    來源缺 Shop Nature 時 Nature 欄留空；所選分組欄缺失時按 MCH2 → MC → Adj. SST 退回並列
    warning，全缺則退回單一分組；缺三個月 Net Sales 欄時只輸出件數判定，均列 warning。

    `top_ratio` 為累計比例：0.8（預設）或 0.7；亦接受 70、"70%" 等寫法。欄名隨比例改變，
    故切換比例重跑不會留下舊欄。`data_date` 為數據截止日（T-1），未提供時用「今天減一日」；
    金額欄名會帶實際覆蓋日數（如 `Net Sales (P73D)`）。

    `group_label` 只在多店舖時生效：提供時用作判定欄名（如「柴灣區」），未提供時以「+」
    串接實際店碼（如 `HA21+HA47`）。單一店舖一律用該店碼，欄名與既有輸出完全相同。
    所選店碼在來源檔找不到時略過並列 warning；全部找不到才拋 ValueError。

    `sct_rules` 為 SCT建議的情況表（情況鍵 → 首選RF／次選RF／建議RF／留空），
    提供時按其填寫「SCT建議」欄，未提供則維持留空（原本行為）；
    Item Status＝Discontinued 的列一律留空，不套用規則（由 SCT 自行判斷）。

    其餘清理、Item_Shop 解析、刪除 Value 欄、三個回覆欄留空與排序規則完全沿用
    `transform_shop_simplify`（以 `include_shop_top80=False`、`include_net_sales=False` 關閉
    原模式的店舖判定與金額欄，避免重複）。回傳 (raw_df, warnings)。
    """
    codes = _shop_simplify_site_codes(site_code)
    if not codes:
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
    # 逐個店碼比對：先按來源檔寫法精確比對，找不到才按大小寫不敏感比對；
    # 仍然找不到的店碼略過並列 warning（新舖或未收錄的店舖不應令整份報表停擺）。
    upper_to_actual: dict[str, str] = {}
    for value in pd.unique(sites):
        text = str(value).strip()
        if text and text.upper() not in upper_to_actual:
            upper_to_actual[text.upper()] = text

    matched_sites: list[str] = []
    missing_sites: list[str] = []
    casing_notes: list[str] = []
    for code in codes:
        actual = upper_to_actual.get(code.upper())
        if actual is None:
            missing_sites.append(code)
            continue
        if actual != code:
            casing_notes.append(f"「{code}」→「{actual}」")
        if actual not in matched_sites:
            matched_sites.append(actual)

    if not matched_sites:
        if len(codes) == 1:
            raise ValueError(
                f"指定店舖「{codes[0]}」在清理後沒有資料列，請確認 Site Code 是否正確，"
                "或來源檔是否只包含部分店舖。"
            )
        raise ValueError(
            f"指定店舖群組（{'、'.join(codes)}）在清理後都沒有資料列，"
            "請確認 Site Code 是否正確，或來源檔是否只包含部分店舖。"
        )

    label = _shop_simplify_site_label(matched_sites, group_label)
    scope_text = _shop_simplify_site_scope_text(matched_sites, label)
    if casing_notes:
        warnings.append(
            f"以下店碼按大小寫不敏感比對：{'、'.join(casing_notes)}，以下按來源檔寫法計算。"
        )
    if missing_sites:
        warnings.append(
            f"以下指定店舖在來源檔沒有資料列，已略過：{'、'.join(missing_sites)}"
            f"（Top {pct}% 只按其餘所選店舖計算）。"
        )

    matched_upper = {site.upper() for site in matched_sites}
    shop_raw = df_clean[sites.str.upper().isin(matched_upper)].copy()
    if shop_raw.empty:
        raise ValueError(f"{scope_text}在清理後沒有資料列，無法計算 Top {pct}%。")

    site_top_column = _shop_simplify_site_top80_column(label, share)
    nature_column = _shop_simplify_nature_top_column(share)
    value_site_column = _shop_simplify_value_site_top_column(label, share)
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
    #    稍後按每列自己的 Shop Nature 映射到所選店舖的列；其他 Nature 不計入分母。
    #    單一店舖時只會有一個 Nature，行為與舊版相同；多店舖時逐個 Nature 各自計算。 ──
    nature_available = False
    nature_values: list[str] = []
    nature_keys_by_value: dict[str, set[tuple[str, str]]] = {}
    nature_value_keys_by_value: dict[str, set[tuple[str, str]]] = {}
    nature_shop_count: dict[str, int] = {}
    nature_total: dict[str, float] = {}
    nature_net_total: dict[str, float] = {}
    if nature_col is None:
        warnings.append(
            "來源檔案缺少「Shop Nature」，未能判定「"
            f"{nature_column}」，該欄留空。"
        )
    else:
        row_natures = shop_raw[nature_col].fillna("").astype(str).str.strip()
        blank_nature_rows = int((row_natures == "").sum())
        nature_values = [str(v) for v in dict.fromkeys(row_natures.tolist()) if str(v).strip()]
        if not nature_values:
            warnings.append(
                f"{scope_text}的 Shop Nature 空白，未能判定「"
                f"{nature_column}」，該欄留空。"
            )
        else:
            nature_available = True
            if len(nature_values) > 1:
                warnings.append(
                    f"{scope_text}跨 {len(nature_values)} 個 Shop Nature"
                    f"（{'、'.join(nature_values)}），「{nature_column}」與對比欄"
                    "按每間店自己的 Shop Nature 計算。"
                )
            if blank_nature_rows:
                warnings.append(
                    f"{scope_text}有 {blank_nature_rows:,} 列的 Shop Nature 空白，"
                    f"該等列的「{nature_column}」與對比欄留空。"
                )

            nature_series = df_clean[nature_col].fillna("").astype(str).str.strip()
            for nature_value in nature_values:
                nature_df = df_clean[nature_series == nature_value]
                nature_sku = _cached_safe_sku_series(nature_df[sku_col]).astype(str).str.strip()
                nature_qty = _safe_numeric(nature_df[sold_col])
                if group_col is None:
                    nature_group = pd.Series("(全部)", index=nature_df.index)
                else:
                    nature_group = nature_df[group_col].fillna("").astype(str).str.strip()
                    nature_group = nature_group.mask(nature_group == "", "(空白)")
                nature_keys_by_value[nature_value] = _shop_simplify_top80_keys(
                    nature_group, nature_sku, nature_qty, share=share
                )
                if net_available:
                    nature_net = _shop_simplify_net_sales_series(nature_df, net_cols)
                    nature_value_keys_by_value[nature_value] = _shop_simplify_top80_keys(
                        nature_group, nature_sku, nature_net, share=share, tail_guard=False
                    )
                    nature_net_total[nature_value] = float(nature_net.sum())
                    del nature_net

                nature_sites = nature_df[site_col].fillna("").astype(str).str.strip()
                nature_sites = nature_sites[nature_sites != ""]
                nature_shop_count[nature_value] = int(nature_sites.nunique())
                nature_total[nature_value] = float(nature_qty.sum())
                del nature_df, nature_sku, nature_qty, nature_group, nature_sites
            del nature_series

    # 全表清理結果只留到此處，及早釋放（80 萬行可省數 GB）
    del df_clean, sites
    gc.collect()

    def _site_progress(fraction, msg):
        if progress_callback:
            progress_callback(0.2 + 0.65 * float(fraction), msg)

    if progress_callback:
        progress_callback(0.2, f"整理 {scope_text} 的店舖×SKU 明細...")
    shop_df, base_warnings = transform_shop_simplify(
        shop_raw,
        progress_callback=_site_progress if progress_callback else None,
        shop_natures=None,
        include_shop_top80=False,
        include_net_sales=False,
        apply_shop_blocked=False,
    )
    warnings = list(base_warnings) + warnings

    if progress_callback:
        progress_callback(0.88, f"計算 {scope_text} 的 Top {pct}%（件數＋金額）...")
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
    # 每列按自己的 Shop Nature 取回同 Nature 判定；Nature 空白或不可用的列留空
    row_nature = (
        shop_df[nature_col].fillna("").astype(str).str.strip()
        if nature_col is not None and nature_col in shop_df.columns
        else pd.Series("", index=shop_df.index, dtype=object)
    )
    if nature_available:
        nature_flags = pd.Series(
            [
                (group, sku) in nature_keys_by_value.get(str(nature), set())
                for nature, group, sku in zip(row_nature, shop_group, shop_sku)
            ],
            index=shop_df.index,
            dtype=bool,
        )
        nature_known = row_nature.isin(nature_values)
    else:
        nature_flags = pd.Series(False, index=shop_df.index, dtype=bool)
        nature_known = pd.Series(False, index=shop_df.index, dtype=bool)

    if float(shop_qty.sum()) <= 0:
        warnings.append(
            f"{scope_text}的 Sold Qty (P90D) 合計為 0，"
            f"「{site_top_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
        )
    if nature_available:
        for nature_value in nature_values:
            if nature_shop_count.get(nature_value, 0) <= 1:
                warnings.append(
                    f"同 Shop Nature「{nature_value}」在來源檔只有 1 間店舖，"
                    f"「{nature_column}」與店舖 Top {pct}% 結果相同。"
                )
            if nature_total.get(nature_value, 0.0) <= 0:
                warnings.append(
                    f"同 Shop Nature「{nature_value}」的 Sold Qty (P90D) 合計為 0，"
                    f"「{nature_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
                )

    shop_df[site_top_column] = np.where(shop_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO)
    shop_df[nature_column] = pd.Series(
        np.where(nature_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO),
        index=shop_df.index,
        dtype=object,
    ).where(nature_known, "")
    compare_qty_column = _shop_simplify_compare_column(label, share, _SHOP_SIMPLIFY_QTY_SUFFIX)
    _compare_qty = _shop_simplify_compare_labels(shop_flags, nature_flags, nature_available)
    shop_df[compare_qty_column] = _compare_qty.where(nature_known, "") if nature_available else _compare_qty

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
            [
                (group, sku) in nature_value_keys_by_value.get(str(nature), set())
                for nature, group, sku in zip(row_nature, shop_group, shop_sku)
            ]
            if nature_available else False,
            index=shop_df.index,
            dtype=bool,
        )
        if float(shop_net.sum()) <= 0:
            warnings.append(
                f"{scope_text}的 {net_column} 合計為 0，"
                f"「{value_site_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
            )
        if nature_available:
            for nature_value in nature_values:
                if nature_net_total.get(nature_value, 0.0) <= 0:
                    warnings.append(
                        f"同 Shop Nature「{nature_value}」的 {net_column} 合計為 0，"
                        f"「{value_nature_column}」全部標示為「{_SHOP_SIMPLIFY_TOP80_NO}」。"
                    )

        shop_df[net_column] = shop_net.to_numpy(dtype=float)
        shop_df[value_site_column] = np.where(
            shop_value_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO
        )
        shop_df[value_nature_column] = pd.Series(
            np.where(nature_value_flags, _SHOP_SIMPLIFY_TOP80_YES, _SHOP_SIMPLIFY_TOP80_NO),
            index=shop_df.index,
            dtype=object,
        ).where(nature_known, "")
        compare_value_column = _shop_simplify_compare_column(label, share, _SHOP_SIMPLIFY_VALUE_SUFFIX)
        _compare_value = _shop_simplify_compare_labels(
            shop_value_flags, nature_value_flags, nature_available
        )
        shop_df[compare_value_column] = (
            _compare_value.where(nature_known, "") if nature_available else _compare_value
        )

    # SCT建議：按四個判定欄與使用者指定的情況表自動填寫（未指定時維持留空）
    if sct_rules:
        _sct_columns = [site_top_column, nature_column, value_site_column, value_nature_column]
        if not net_available:
            # 金額口徑不可用時，金額兩項一律視為非 Top（情況鍵尾兩碼為 N）
            _sct_columns = _sct_columns[:2] + [None, None]
        _sct_series, _ = _shop_simplify_sct_mark_buyer_nd(
            shop_df, _shop_simplify_sct_series(shop_df, sct_rules, _sct_columns)
        )
        _sct_filled, _sct_skipped = _shop_simplify_sct_skip_discontinued(shop_df, _sct_series)
        shop_df[_SHOP_SIMPLIFY_SCT_COLUMN] = _sct_filled
        _sct_buyer_nd = int(
            _shop_simplify_cell_mask(_sct_filled, _SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE).sum()
        )
        _sct_labels = shop_df[_SHOP_SIMPLIFY_SCT_COLUMN].replace("", np.nan).dropna()
        if len(_sct_labels):
            _sct_notes = []
            if _sct_skipped:
                _sct_notes.append(
                    f"Item Status＝{_SHOP_SIMPLIFY_SCT_SKIP_STATUS} 的 {_sct_skipped:,} 列一律留空"
                )
            if _sct_buyer_nd:
                _sct_notes.append(
                    f"Buyer＝{_SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER} 且 RP Type＝ND 的 "
                    f"{_sct_buyer_nd:,} 列標示「{_SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE}」"
                )
            _sct_note = f"（{'；'.join(_sct_notes)}）" if _sct_notes else ""
            warnings.append(
                f"已按 SCT建議規則填寫「{_SHOP_SIMPLIFY_SCT_COLUMN}」欄，"
                f"共 {len(_sct_labels):,} 列{_sct_note}（可按實際情況在 Excel 內修改）。"
            )
        elif _sct_skipped:
            warnings.append(
                f"全部列均為 Item Status＝{_SHOP_SIMPLIFY_SCT_SKIP_STATUS}，"
                f"「{_SHOP_SIMPLIFY_SCT_COLUMN}」欄一律留空，由 SCT 自行判斷。"
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
        progress_callback(0.85, f"完成 {scope_text} 的 Top {pct}% 判定...")
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
    row_fills: dict[str, str] | None = None,
    column_fills: dict[str, str] | None = None,
    as_percent: bool = False,
    widths: dict[int, float] | None = None,
    header_fill: str | None = None,
) -> tuple[int, int]:
    """寫一個「列標籤 × 欄標籤」的 SKU 計數區塊（零值留空、底部總計列）。

    `row_fills`／`column_fills` 為選填的底色（hex，按列標籤／欄標籤），
    用於 SCT建議 分類的顏色標註；沒有對應顏色的標籤改用
    `_SHOP_SIMPLIFY_SUMMARY_COLUMN_FILLS` 的欄標籤分類色，最後才用區塊表頭底色。
    `as_percent=True` 時版面不變，但數值改為佔該區塊總計的百分比（格式 `0.0%`），
    總計列的每一欄為該欄佔比、右下角為 100.0%。

    版面：區塊標題深藍加粗；表頭與資料格四面細框線（零值留空的格亦畫框線），
    數值置中；總計列加粗並在頂部加一條較粗框線。`widths` 會累計每欄最長的
    表頭／列標籤顯示寬度，供區塊全部寫完後一次過設定欄寬。
    回傳 (最後一列列號, 用到的最大欄號)。
    """
    widths = widths if widths is not None else {}
    base_fill = header_fill or _SHOP_SIMPLIFY_SUMMARY_COUNT_HEADER_FILL
    grid = _shop_simplify_summary_border()
    grid_total = _shop_simplify_summary_border(strong_top=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    number_align = Alignment(horizontal="center", vertical="center")

    title_cell = ws.cell(row=start_row, column=1, value=title)
    title_cell.font = Font(bold=True, size=11.5, color=_SHOP_SIMPLIFY_SUMMARY_TITLE_COLOUR)
    ws.row_dimensions[start_row].height = 20

    header = start_row + 1
    total_col = 2 + len(column_labels)

    label_head = ws.cell(row=header, column=1, value="列標籤")
    label_head.font = Font(bold=True)
    label_head.fill = _shop_simplify_summary_fill(base_fill)
    label_head.alignment = header_align
    label_head.border = grid
    widths[1] = max(widths.get(1, 0), _shop_simplify_summary_text_width("列標籤"))

    for j, label in enumerate(column_labels, start=2):
        cell = ws.cell(row=header, column=j, value=label)
        cell.alignment = header_align
        colour = (
            (column_fills or {}).get(label)
            or _shop_simplify_summary_column_fill(label)
            or base_fill
        )
        cell.fill = _shop_simplify_summary_fill(colour)
        cell.font = Font(bold=True)
        cell.border = grid
        widths[j] = max(widths.get(j, 0), _shop_simplify_summary_text_width(label))

    total_head = ws.cell(row=header, column=total_col, value="總計")
    total_head.alignment = header_align
    total_head.font = Font(bold=True)
    total_head.fill = _shop_simplify_summary_fill(base_fill)
    total_head.border = grid
    widths[total_col] = max(widths.get(total_col, 0), _shop_simplify_summary_text_width("總計"))
    ws.row_dimensions[header].height = 30

    grand = sum(
        counts.get((level, label), 0) for level in row_levels for label in column_labels
    )

    def write(r: int, c: int, value: int, border: Border = grid, bold: bool = False) -> None:
        cell = ws.cell(row=r, column=c)
        if as_percent:
            cell.value = value / grand if grand else 0
            cell.number_format = "0.0%"
        else:
            cell.value = value
            cell.number_format = "#,##0"
        cell.alignment = number_align
        cell.border = border
        if bold:
            cell.font = Font(bold=True)

    column_totals = dict.fromkeys(column_labels, 0)
    last = header
    for i, level in enumerate(row_levels, start=header + 1):
        cell = ws.cell(row=i, column=1, value=level)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = grid
        colour = (row_fills or {}).get(level)
        if colour:
            cell.fill = _shop_simplify_summary_fill(colour)
            cell.font = Font(bold=True)
        widths[1] = max(widths.get(1, 0), _shop_simplify_summary_text_width(level))
        row_total = 0
        for j, label in enumerate(column_labels, start=2):
            count = counts.get((level, label), 0)
            if count:
                write(i, j, count)
            else:
                # 零值留空，但仍畫框線，令整個區塊看起來是一張完整的表
                ws.cell(row=i, column=j).border = grid
            row_total += count
            column_totals[label] += count
        if row_total:
            write(i, total_col, row_total)
        else:
            ws.cell(row=i, column=total_col).border = grid
        ws.row_dimensions[i].height = 18
        last = i

    total_row = last + 1
    total_label = ws.cell(row=total_row, column=1, value="總計")
    total_label.font = Font(bold=True)
    total_label.alignment = Alignment(horizontal="left", vertical="center")
    total_label.border = grid_total
    for j, label in enumerate(column_labels, start=2):
        if column_totals[label]:
            write(total_row, j, column_totals[label], border=grid_total, bold=True)
        else:
            ws.cell(row=total_row, column=j).border = grid_total
    if grand:
        write(total_row, total_col, grand, border=grid_total, bold=True)
    else:
        ws.cell(row=total_row, column=total_col).border = grid_total
    ws.row_dimensions[total_row].height = 18
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

    每個計數區塊之後緊接同一區塊的「佔比」版本（標題 `佔比 - SKU：…`）：版面與欄名完全相同，
    數值為佔該區塊總計的百分比（格式 0.0%，分母＝該區塊全部列與欄的總數），方便直接比較分佈。

    區塊次序：
    1. × Item Status（原有版面，固定於 A3／A4）；
    1b. Item Status × RP Type（每個狀態再分 RF／ND，另有其他值時加「其他」）；
    2. × RP Type（RF／ND／其他）；
    2b. × SCT建議（僅在該欄有值時輸出：首選RF／次選RF／建議RF）；
    2c／2d. SCT建議 × Item Status、SCT建議 × RP Type（列 = 各建議值，按值填色）；
    3. × 四個 Top N% 判定欄（店舖件數／同Nature件數／店舖金額／同Nature金額，
       欄名取實際輸出欄，故 Top 70%／80% 與店舖代號自動跟隨）；
    4. RF／ND × 每個 Top N% 判定欄的四象限（店舖件數／同Nature件數／店舖金額／同Nature金額
       各一個區塊，欄為 RF-Top／RF-非Top／ND-Top／ND-非Top；非Top 含「否」及未判定）。
    零值一律留空，每個區塊底部有總計列；來源缺該欄時該區塊自動略過。

    版面（美化）：A1 為標題、A2 為本次輸出概況；每個區塊標題深藍加粗，表頭按分類填色
    （計數＝藍、佔比＝灰；「是／RF」綠、「否／非Top」灰、「Discontinued／ND」淡紅，
    SCT建議 沿用其分類色），全部表頭與資料格加四面細框線（等同格線），數值置中、
    總計列加粗加上框線；欄寬按各欄最長內容自動換算，捲動時左邊的列標籤保持可見。
    """
    index_col = "Adj. MCH1" if "Adj. MCH1" in df.columns else (
        "Item Class" if "Item Class" in df.columns else None
    )
    if df is None or df.empty or index_col is None or "SKU" not in df.columns:
        return

    levels = _shop_simplify_summary_levels(df[index_col])
    row_order = _shop_simplify_summary_row_order(levels)
    valid = df["SKU"].notna()

    # 全表欄寬需求（欄號 → 最長顯示寬度），區塊全部寫完後一次過設定
    widths: dict[int, float] = {}

    # ── 標題列（A1／A2；A3 起才是第一個區塊，沿用既有位置）──
    ws.cell(row=1, column=1, value="店舖精簡SKU模式 — Summary").font = Font(
        bold=True, size=15, color=_SHOP_SIMPLIFY_SUMMARY_TITLE_COLOUR
    )
    ws.row_dimensions[1].height = 26
    overview = [f"共 {len(df):,} 列 SKU"]
    if "Site Code" in df.columns:
        shops = df["Site Code"].dropna().astype(str).str.strip()
        shops = shops[shops != ""]
        if len(shops):
            n_shops = shops.nunique()
            if n_shops == 1:
                overview.append(f"指定店舖 {shops.iloc[0]}")
            else:
                overview.append(
                    f"指定店舖 {n_shops} 間（{_shop_simplify_site_codes_text(shops.tolist())}）"
                )
    tops = _shop_simplify_summary_top_columns(df)
    if tops:
        found = re.search(r"Top\s*(\d{1,3})%", " ".join(str(n) for n in tops))
        if found:
            overview.append(f"Top {found.group(1)}%")
    window = next(
        (c for c in df.columns if re.fullmatch(r"Net Sales \(P\d+D\)", str(c).strip())),
        None,
    )
    if window is not None:
        found = re.search(r"P(\d+)D", str(window))
        overview.append(f"金額視窗 {window}（{found.group(1)} 日）")
    overview.append("每個區塊先列計數、再列佔比（分母＝該區塊總計）")
    info = ws.cell(row=2, column=1, value="｜".join(overview))
    info.font = Font(size=10, color="595959")
    ws.row_dimensions[2].height = 18

    def block_pair(suffix, start, labels, counts, row_fills=None, column_fills=None, rows=None):
        """先寫計數區塊，緊接同一區塊的「佔比」版本（版面相同，分母＝該區塊總計）。

        `rows` 預設用 Adj. MCH1 的列次序；SCT建議 交叉區塊改傳 `sct_row_order`。
        """
        rows = row_order if rows is None else rows
        end, last_col = _shop_simplify_summary_block(
            ws, f"計數 - SKU：{suffix}", start, rows, labels, counts, row_fills, column_fills,
            widths=widths,
        )
        end, last_col = _shop_simplify_summary_block(
            ws, f"佔比 - SKU：{suffix}", end + 3, rows, labels, counts,
            row_fills, column_fills, as_percent=True, widths=widths,
            header_fill=_SHOP_SIMPLIFY_SUMMARY_PERCENT_HEADER_FILL,
        )
        return end, last_col

    # ── 區塊 1：× Item Status（原有版面，測試與使用者習慣依賴 A3／A4 位置）──
    status_col = "Item Status" if "Item Status" in df.columns else None
    status_labels: list[str] = []
    blocks_end = 2
    max_col = 0
    if status_col is not None:
        # A3 標題／B3「欄標籤」與 A4 表頭位置為既有版面（測試與使用者習慣依賴），只加樣式不改位置
        title_cell = ws.cell(row=3, column=1, value="計數 - SKU")
        title_cell.font = Font(bold=True, size=11.5, color=_SHOP_SIMPLIFY_SUMMARY_TITLE_COLOUR)
        ws.row_dimensions[3].height = 20
        legacy_caption = ws.cell(row=3, column=2, value="欄標籤")
        legacy_caption.font = Font(size=9, italic=True, color="808080")

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

        # 表頭（藍底、加粗、細框、置中換行）與資料格（四面細框、數值置中）；
        # 「總計」列加粗並在頂部加較粗框線，與其他列分開
        grid = _shop_simplify_summary_border()
        grid_total = _shop_simplify_summary_border(strong_top=True)
        legacy_base = _SHOP_SIMPLIFY_SUMMARY_COUNT_HEADER_FILL
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        head = ws.cell(row=4, column=1, value="列標籤")
        head.font = Font(bold=True)
        head.fill = _shop_simplify_summary_fill(legacy_base)
        head.alignment = header_align
        head.border = grid
        widths[1] = max(widths.get(1, 0), _shop_simplify_summary_text_width("列標籤"))
        for j, status in enumerate(statuses, start=2):
            cell = ws.cell(row=4, column=j, value=str(status))
            cell.font = Font(bold=True)
            cell.fill = _shop_simplify_summary_fill(
                _shop_simplify_summary_column_fill(status) or legacy_base
            )
            cell.alignment = header_align
            cell.border = grid
            widths[j] = max(widths.get(j, 0), _shop_simplify_summary_text_width(status))
        total_head = ws.cell(row=4, column=total_col, value="總計")
        total_head.font = Font(bold=True)
        total_head.fill = _shop_simplify_summary_fill(legacy_base)
        total_head.alignment = header_align
        total_head.border = grid
        widths[total_col] = max(
            widths.get(total_col, 0), _shop_simplify_summary_text_width("總計")
        )
        ws.row_dimensions[4].height = 30

        for i, (idx, row) in enumerate(table.iterrows(), start=5):
            is_total = str(idx) == "總計"
            cell = ws.cell(row=i, column=1, value=str(idx))
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = grid_total if is_total else grid
            if is_total:
                cell.font = Font(bold=True)
            widths[1] = max(widths.get(1, 0), _shop_simplify_summary_text_width(idx))
            for j, status in enumerate(statuses, start=2):
                count = int(row[status])
                target = ws.cell(row=i, column=j)
                if count:
                    target.value = count
                    target.number_format = "#,##0"
                    target.alignment = Alignment(horizontal="center", vertical="center")
                target.border = grid_total if is_total else grid
                if is_total:
                    target.font = Font(bold=True)
            total = int(row["總計"])
            total_cell = ws.cell(row=i, column=total_col)
            if total:
                total_cell.value = total
                total_cell.number_format = "#,##0"
                total_cell.alignment = Alignment(horizontal="center", vertical="center")
            total_cell.border = grid_total if is_total else grid
            if is_total:
                total_cell.font = Font(bold=True)
            ws.row_dimensions[i].height = 18
        status_labels = [str(s) for s in statuses]
        blocks_end = 4 + len(table)
        max_col = total_col

        # 同一區塊的佔比版本（分母＝區塊 1 的總計）
        status_levels = _shop_simplify_summary_levels(df[status_col])
        status_counts = _shop_simplify_summary_counts(
            levels, [(label, status_levels.eq(label)) for label in status_labels], valid
        )
        blocks_end, _ = _shop_simplify_summary_block(
            ws, "佔比 - SKU", blocks_end + 3, row_order, status_labels, status_counts,
            as_percent=True, widths=widths,
            header_fill=_SHOP_SIMPLIFY_SUMMARY_PERCENT_HEADER_FILL,
        )

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
        blocks_end, _ = block_pair(
            "Item Status × RP Type", blocks_end + 3, labels, counts
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
        blocks_end, _ = block_pair("RP Type", blocks_end + 3, labels, counts)
        max_col = max(max_col, 2 + len(labels))

    # ── 區塊 2b–2d：SCT建議 分佈（僅在該欄有值時輸出，方便核對自動建議的分佈）──
    sct_column = _SHOP_SIMPLIFY_SCT_COLUMN
    if sct_column in df.columns:
        sct = df[sct_column].fillna("").astype(str).str.strip()
        if bool((sct != "").any()):
            known = [v for v in _SHOP_SIMPLIFY_SCT_OPTIONS if v]
            value_labels = [v for v in known if bool(sct.eq(v).any())]
            value_labels += sorted({v for v in sct.unique() if v and v not in known})

            # 2b：× SCT建議（列 = Adj. MCH1，欄 = 各建議值；欄名以顏色標註）
            masks = [(label, sct.eq(label)) for label in value_labels]
            counts = _shop_simplify_summary_counts(levels, masks, valid)
            blocks_end, _ = block_pair(
                sct_column, blocks_end + 3, value_labels, counts,
                column_fills=_SHOP_SIMPLIFY_SCT_FILLS,
            )
            max_col = max(max_col, 2 + len(value_labels))

            # 2c／2d：SCT建議 × Item Status／RP Type（列 = 各建議值，欄 = 分類；列名以顏色標註）
            sct_levels = sct.mask(sct == "", _SHOP_SIMPLIFY_SCT_BLANK_LABEL)
            sct_row_order = list(value_labels)
            if bool(sct_levels.eq(_SHOP_SIMPLIFY_SCT_BLANK_LABEL).any()):
                sct_row_order.append(_SHOP_SIMPLIFY_SCT_BLANK_LABEL)

            if status_labels:
                status_levels = _shop_simplify_summary_levels(df[status_col])
                masks = [(label, status_levels.eq(label)) for label in status_labels]
                counts = _shop_simplify_summary_counts(sct_levels, masks, valid)
                blocks_end, _ = block_pair(
                    f"{sct_column} × Item Status", blocks_end + 3, status_labels, counts,
                    row_fills=_SHOP_SIMPLIFY_SCT_FILLS, rows=sct_row_order,
                )
                max_col = max(max_col, 2 + len(status_labels))

            if "RP Type" in df.columns:
                sct_rp = _shop_simplify_summary_levels(df["RP Type"]).str.upper()
                rp_labels = ["RF", "ND"]
                rp_masks = [("RF", sct_rp.eq("RF")), ("ND", sct_rp.eq("ND"))]
                rp_other = ~sct_rp.isin(["RF", "ND"])
                if bool(rp_other.any()):
                    rp_labels.append("其他")
                    rp_masks.append(("其他", rp_other))
                counts = _shop_simplify_summary_counts(sct_levels, rp_masks, valid)
                blocks_end, _ = block_pair(
                    f"{sct_column} × RP Type", blocks_end + 3, rp_labels, counts,
                    row_fills=_SHOP_SIMPLIFY_SCT_FILLS, rows=sct_row_order,
                )
                max_col = max(max_col, 2 + len(rp_labels))

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
        blocks_end, _ = block_pair(name, blocks_end + 3, labels, counts)
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
            blocks_end, _ = block_pair(
                f"RF／ND × {name}（非Top 含「否」及未判定）", blocks_end + 3, labels, counts
            )
            max_col = max(max_col, 2 + len(labels))

    # 欄寬：A 欄放列標籤（Adj. MCH1）與區塊標題，按其最長標籤換算（32–48 字元）；
    # 其餘按各欄最長表頭／列標籤換算（下限 11、上限 22 字元，長欄名靠自動換行處理）
    ws.column_dimensions["A"].width = min(
        _SHOP_SIMPLIFY_SUMMARY_MAX_LABEL_WIDTH,
        max(_SHOP_SIMPLIFY_SUMMARY_MIN_LABEL_WIDTH, widths.get(1, 0) + 6.0),
    )
    for j in range(2, max_col + 1):
        needed = widths.get(j, 0) + 2.5
        ws.column_dimensions[get_column_letter(j)].width = min(
            _SHOP_SIMPLIFY_SUMMARY_MAX_WIDTH,
            max(_SHOP_SIMPLIFY_SUMMARY_MIN_WIDTH, needed),
        )

    # 檢視：關閉預設格線（各表已用框線畫出格線，避免與框線重疊）、凍結 A 欄、
    # 略微縮小以便一次看到整個區塊；分頁標籤用 Summary 表頭的藍色
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.freeze_panes = "B1"
    ws.sheet_properties.tabColor = _SHOP_SIMPLIFY_SUMMARY_COUNT_HEADER_FILL


_SHOP_SIMPLIFY_GUIDE_SHEET = "說明"


def _shop_simplify_guide_flags(df: pd.DataFrame) -> list[tuple[str, str]]:
    """說明頁情況表用的 Top 判定欄與其簡短標籤。

    次序跟 Export 一致：件數（店舖 → 同 Shop Nature）→ 金額（店舖 → 同 Shop Nature）；
    原模式只有件數與金額兩個判定，故只有「件數」「金額」兩欄。
    """
    flags: list[tuple[str, str]] = []
    for name in _shop_simplify_summary_top_columns(df):
        text = str(name).strip()
        unit = "金額" if text.endswith(_SHOP_SIMPLIFY_VALUE_SUFFIX) else "件數"
        if text.startswith("是否同Shop Nature"):
            label = f"{unit}－同 Shop Nature"
        elif "該店" in text:
            label = unit
        else:
            label = f"{unit}－店舖"
        flags.append((name, label))
    return flags


def _write_shop_simplify_guide(ws, df: pd.DataFrame) -> None:
    """「說明」工作表：SCT建議 的填寫規則（按本檔實際結果）、Top N% 與顏色說明。

    情況表按輸出實際有的 Top 判定欄生成（指定店舖四個判定＝16 種、原模式兩個判定＝4 種），
    並逐個情況列出本檔 SKU 數與實際寫入的 SCT建議，故更換比例、店舖或規則後說明仍然準確。
    """
    thin = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    header_fill = PatternFill(
        start_color=_SHOP_SIMPLIFY_HEADER_BLUE,
        end_color=_SHOP_SIMPLIFY_HEADER_BLUE,
        fill_type="solid",
    )
    title_font = Font(bold=True, size=16, color="1F3864")
    section_font = Font(bold=True, size=12, color="1F3864")
    body_font = Font(size=10.5)
    small_font = Font(size=10)
    note_font = Font(size=9, italic=True, color="595959")

    ws.sheet_view.showGridLines = False
    for letter, width in {
        "A": 2.5, "B": 10.0, "C": 20.0, "D": 24.0, "E": 20.0,
        "F": 20.0, "G": 24.0, "H": 14.0, "I": 18.0,
    }.items():
        ws.column_dimensions[letter].width = width

    def section(r: int, text: str) -> int:
        cell = ws.cell(row=r, column=2, value=text)
        cell.font = section_font
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        ws.row_dimensions[r].height = 24
        return r + 1

    def paragraph(r: int, text: str, font=None, chars_per_line: int = 62) -> int:
        cell = ws.cell(row=r, column=2, value=text)
        cell.font = font or body_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
        lines = max(1, (len(text) + chars_per_line - 1) // chars_per_line)
        ws.row_dimensions[r].height = 16 * lines + 6
        return r + 1

    def table_header(r: int, headers: list[str]) -> int:
        for j, text in enumerate(headers, start=2):
            cell = ws.cell(row=r, column=j, value=text)
            cell.font = Font(bold=True, size=10)
            cell.fill = header_fill
            cell.border = thin
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 32
        return r + 1

    flags = _shop_simplify_guide_flags(df)
    valid = df["SKU"].notna() if "SKU" in df.columns else pd.Series(True, index=df.index)
    sct_series = (
        df[_SHOP_SIMPLIFY_SCT_COLUMN].fillna("").astype(str).str.strip()
        if _SHOP_SIMPLIFY_SCT_COLUMN in df.columns
        else pd.Series("", index=df.index)
    )
    sct_any = bool((sct_series != "").any())

    # ── 標題 ──
    cell = ws.cell(row=1, column=2, value="店舖精簡SKU模式 — 輸出說明")
    cell.font = title_font
    ws.row_dimensions[1].height = 28

    # ── 本次輸出概況（單行）──
    parts = [f"{len(df):,} 列"]
    if "Site Code" in df.columns:
        shops = df["Site Code"].dropna().astype(str).str.strip()
        shops = shops[shops != ""]
        if len(shops):
            n_shops = shops.nunique()
            parts.append(
                f"{n_shops} 間店舖（{shops.iloc[0]}）" if n_shops == 1
                else f"{n_shops} 間店舖（{_shop_simplify_site_codes_text(shops.tolist())}）"
            )
    if flags:
        found = re.search(r"Top\s*(\d{1,3})%", " ".join(str(n) for n, _ in flags))
        if found:
            parts.append(f"Top {found.group(1)}%")
    window = next(
        (c for c in df.columns if re.fullmatch(r"Net Sales \(P\d+D\)", str(c).strip())),
        None,
    )
    if window is not None:
        found = re.search(r"P(\d+)D", str(window))
        parts.append(f"金額視窗 {window}（{found.group(1)} 日）")
    if sct_any:
        filled = int((sct_series != "").sum())
        parts.append(f"SCT建議 已填寫 {filled:,} 列（{filled / max(len(df), 1):.1%}）")
    else:
        parts.append("SCT建議 未填寫（全部留空）")
    info = ws.cell(row=2, column=2, value="本次輸出：" + "｜".join(parts))
    info.font = Font(size=10, color="404040")
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=9)
    ws.row_dimensions[2].height = 20

    row = 4

    # ── 一、SCT建議 是甚麼 ──
    row = section(row, "一、SCT建議 是甚麼")
    row = paragraph(
        row,
        "「SCT建議」按 Top N% 判定自動填寫，目的是在匯出時先為每個 SKU 標示一個 RF 取態，減少逐列人手判斷。"
        "填寫規則由應用程式側欄「SCT建議」的情況表決定（指定店舖模式 16 種、店舖精簡SKU模式 4 種）；"
        "第二節列出的是本檔實際採用的規則與結果。留空的 SKU 代表不預設建議，由 SCT 按實際情況判斷（例如 ND）。"
        f"Buyer＝{_SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER} 的 ND 貨品不套用情況表，一律標示"
        f"「{_SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE}」。",
    )
    row += 1

    # ── 二、情況表 ──
    case_count = 2 ** len(flags) if flags else 0
    row = section(row, "二、SCT建議 的情況表（本檔實際採用）")
    if flags:
        formula = "×".join(["2"] * len(flags))
        row = paragraph(
            row,
            (
                f"四個判定各只有「是／否」，兩欄對比由判定自動合成，故全部情況＝{formula}＝{case_count} 種。"
                if len(flags) == 4
                else f"本模式只有件數與金額兩個判定，對比欄由判定自動合成，故全部情況＝{formula}＝{case_count} 種。"
            )
            + "「本檔 SKU 數」為落在該情況的列數，「本檔 SCT建議」為實際寫入的文字。",
        )
        row = table_header(
            row, ["編號"] + [label for _, label in flags] + ["本檔 SKU 數", "本檔 SCT建議"]
        )

        masks = [_shop_simplify_cell_mask(df[name], _SHOP_SIMPLIFY_TOP80_YES) for name, _ in flags]
        valid_arr = valid.to_numpy(dtype=bool)
        for index, case in enumerate(itertools.product("YN", repeat=len(flags)), start=1):
            mask = valid_arr.copy()
            for letter, hit in zip(case, masks):
                mask &= hit if letter == "Y" else ~hit
            count = int(mask.sum())
            values: list[str] = []
            if sct_any and count:
                raw_values = df.loc[mask, _SHOP_SIMPLIFY_SCT_COLUMN].fillna("").astype(str).str.strip()
                values = list(dict.fromkeys(raw_values.tolist()))
            labels = [v for v in values if v]
            shown = "／".join(labels) if labels else "（留空）"
            if len(labels) > 1:
                shown += "（同一情況有多個值）"
            elif labels and "" in values:
                shown += "（部分留空）"

            first = ws.cell(row=row, column=2, value=index)
            first.alignment = Alignment(horizontal="center")
            for j, letter in enumerate(case, start=3):
                cell = ws.cell(row=row, column=j, value="是" if letter == "Y" else "否")
                cell.alignment = Alignment(horizontal="center")
            count_cell = ws.cell(row=row, column=3 + len(flags), value=count)
            count_cell.alignment = Alignment(horizontal="center")
            count_cell.number_format = "#,##0"
            value_cell = ws.cell(row=row, column=4 + len(flags), value=shown)
            value_cell.alignment = Alignment(horizontal="center")
            for j in range(2, 5 + len(flags)):
                cell = ws.cell(row=row, column=j)
                cell.border = thin
                cell.font = small_font
            ws.row_dimensions[row].height = 20
            row += 1
        row = paragraph(
            row,
            "註：同一情況若曾在 Excel 內手動改成不同值，會以「／」列出；SKU 數只計 SKU 非空白的列。"
            "情況的欄序與 Export 的判定欄一致（件數 → 金額，店舖 → 同 Shop Nature）。"
            f"Item Status＝{_SHOP_SIMPLIFY_SCT_SKIP_STATUS} 的 SKU 不套用情況表，一律留空；"
            f"Buyer＝{_SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER} 的 ND 貨品改標示"
            f"「{_SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE}」，故該情況亦會顯示「（部分留空）」或另一值。",
            font=note_font,
        )
    else:
        row = paragraph(
            row,
            "本檔沒有 Top N% 判定欄，故未能列出情況表；請確認已啟用 Top N% 判定，並提供 Sold Qty (P90D)。",
        )
    row += 1

    # ── 三、SCT建議 的值 ──
    row = section(row, "三、SCT建議 的值與留空")
    _sct_text = sct_series.fillna("").astype(str).str.strip()[valid]
    for label, meaning in (
        (_SHOP_SIMPLIFY_SCT_PREFERRED, "最高優先：資源足夠時先處理這批 SKU。"),
        (_SHOP_SIMPLIFY_SCT_SECONDARY, "次一級：首選處理完後才考慮。"),
        (_SHOP_SIMPLIFY_SCT_RF, "可考慮：條件較弱，有剩餘資源時才處理。"),
        (
            _SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE,
            f"Buyer＝{_SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER} 的 ND 貨品：有機會屬受限制貨品"
            "（例如只限澳門出售），必須由 Buyer 決定，不由 SCT 建議轉 RF。",
        ),
        (
            "",
            f"不預設建議，由 SCT 按實際情況判斷（例如 ND）；Item Status＝{_SHOP_SIMPLIFY_SCT_SKIP_STATUS} 亦一律留空。",
        ),
    ):
        text = f"• {label} — {meaning}" if label else f"• 留空 — {meaning}"
        if len(_sct_text):
            hit = int((_sct_text == label).sum())
            text += f"（本檔：{hit:,} 個 SKU）"
        row = paragraph(row, text)
    row += 1

    # ── 四、Top N% 判定與對比欄 ──
    row = section(row, "四、Top N% 判定與對比欄")
    # 多店舖群組時分母是所選店舖合併後的池，分組說明須同步（單店維持原文）
    _site_count = 0
    if "Site Code" in df.columns:
        _site_count = int(
            df["Site Code"].fillna("").astype(str).str.strip().replace("", np.nan).nunique()
        )
    for text in (
        "• 件數 Top：按 Sold Qty (P90D) 在每個分組內由高至低累加，累計銷量首次達該分組總銷量 N% 的 SKU 屬 Top；"
        "同一銷量的整組會一併納入或排除，分組之間不互相擠佔。",
        "• 金額 Top：同一分組與同一 N%，改按 Net Sales (P{N}D)（即 M-2、M-1 與本月至今相加，總和下限為 0）累計。",
        "• 對比欄：兩者皆是／僅店舖／僅同 Shop Nature／兩者皆非，由兩個判定自動合成，不能獨立填寫。",
        "• 「非 Top」包括「否」及未判定（空白）：判定欄沒有「是／否」時一律當作非 Top，故「非Top」欄會包含未判定的 SKU。",
        (
            "• 分組：指定店舖群組模式以所選店舖合併後的所選分組欄分組（同一 SKU 跨店先加總）；"
            if _site_count > 1
            else "• 分組：指定店舖模式以 (Site Code, 側欄所選分組欄) 分組；"
        )
        + "店舖精簡SKU模式以每間店各自的所選分組欄分組；件數與金額各自計算，互不影響。",
    ):
        row = paragraph(row, text)
    row += 1

    # ── 五、顏色說明 ──
    row = section(row, "五、顏色說明")
    row = table_header(row, ["顏色", "出現位置", "含意"])
    for colour, where, meaning in (
        (_SHOP_SIMPLIFY_PRIVATE_LABEL_FILL, "Adj. SST", "含「03 - Private Label」字眼，屬 Private Label 貨品"),
        (_SHOP_SIMPLIFY_GREEN, "各 Top 判定欄（件數、金額）", "值為「是」，即屬該口徑的 Top N%"),
        (_SHOP_SIMPLIFY_GREEN, "RP Type、Shop Stock Qty (CM) 等表頭", "綠色表頭代表重點欄"),
        (_SHOP_SIMPLIFY_PURPLE, "Sold Qty (P90D)", "重點銷量欄（表頭與資料格均為紫色）"),
        (_SHOP_SIMPLIFY_NO_SALE_PINK, "整列", "過去90日有無賣過＝無賣"),
        (_SHOP_SIMPLIFY_RED, "Item Status／RP Type／過去90日有無存貨", "Discontinued／ND／無存貨"),
        (_SHOP_SIMPLIFY_COMPARE_BOTH_FILL, "對比欄（件數、金額）", "兩者皆是：店舖與同 Shop Nature 均屬 Top"),
        (_SHOP_SIMPLIFY_COMPARE_NEITHER_FILL, "對比欄（件數、金額）", "兩者皆非：兩邊都不是 Top，屬精簡候選"),
        (_SHOP_SIMPLIFY_SCT_FILLS[_SHOP_SIMPLIFY_SCT_PREFERRED], "Summary 的 SCT建議 分類", "首選RF（與 Top 判定「是」同色）"),
        (_SHOP_SIMPLIFY_SCT_FILLS[_SHOP_SIMPLIFY_SCT_SECONDARY], "Summary 的 SCT建議 分類", "次選RF"),
        (_SHOP_SIMPLIFY_SCT_FILLS[_SHOP_SIMPLIFY_SCT_RF], "Summary 的 SCT建議 分類", "建議RF"),
        (
            _SHOP_SIMPLIFY_SCT_BUYER_ND_FILL,
            "Export 的 SCT建議；Summary 的 SCT建議 分類",
            f"{_SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE}（Buyer {_SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER} 的 ND 貨品）",
        ),
        (_SHOP_SIMPLIFY_SCT_FILLS[_SHOP_SIMPLIFY_SCT_BLANK_LABEL], "Summary 的 SCT建議 分類", "留空（未給建議，例如停售品）"),
        (_SHOP_SIMPLIFY_HEADER_BLUE, "其餘表頭", "一般欄"),
    ):
        swatch = ws.cell(row=row, column=2, value=f"#{colour}")
        swatch.fill = PatternFill(start_color=colour, end_color=colour, fill_type="solid")
        swatch.font = Font(size=9, bold=True, color="404040")
        swatch.alignment = Alignment(horizontal="center", vertical="center")
        where_cell = ws.cell(row=row, column=3, value=where)
        where_cell.font = small_font
        where_cell.alignment = Alignment(vertical="center", wrap_text=True)
        meaning_cell = ws.cell(row=row, column=4, value=meaning)
        meaning_cell.font = small_font
        meaning_cell.alignment = Alignment(vertical="center", wrap_text=True)
        for j in range(2, 5):
            ws.cell(row=row, column=j).border = thin
        lines = max((len(where) + 9) // 10, (len(meaning) + 11) // 12, 1)
        ws.row_dimensions[row].height = 15 * lines + 10
        row += 1
    row += 1

    # ── 六、Summary 工作表 ──
    row = section(row, "六、Summary 工作表")
    blocks = ["× Item Status", "Item Status × RP Type（每個狀態再分 RF／ND）", "× RP Type"]
    if sct_any:
        blocks.append("× SCT建議（欄名按建議值填色）")
        blocks.append("SCT建議 × Item Status（列 = 各建議值，按值填色）")
        blocks.append("SCT建議 × RP Type（同上）")
    if flags:
        blocks.append(
            "× 兩個 Top N% 判定欄（是／否）" if len(flags) == 2 else "× 四個 Top N% 判定欄（是／否）"
        )
        blocks.append(
            "兩個 RF／ND × Top N% 四象限"
            if len(flags) == 2
            else "四個 RF／ND × Top N% 四象限（RF-Top／RF-非Top／ND-Top／ND-非Top）"
        )
    block_text = "；".join(f"（{i}）{text}" for i, text in enumerate(blocks, 1))
    row = paragraph(
        row,
        f"Summary 以 Adj. MCH1 為列，逐個區塊統計 SKU 數；零值留空，每個區塊底部有「總計」列。"
        f"每個計數區塊之後緊接同一區塊的「佔比 - SKU：…」版本：欄名與版面相同，數值為佔該區塊"
        f"總計的百分比（格式 0.0%，右下角為 100.0%），方便直接比較分佈。區塊依次為：{block_text}。"
        f"版面：表頭按分類填色——是／RF 綠、Discontinued／ND 淡紅、否／非Top 灰、SCT建議 用其"
        f"分類色；計數區塊藍底、佔比區塊灰底，全表每格有框線，總計列加粗並加上框線。",
    )
    row += 1

    # ── 七、回覆欄與其他注意事項 ──
    row = section(row, "七、回覆欄與其他注意事項")
    notes = [
        "• Shop 回覆新的NDRF、如果RF的Safety Qty 一律留空，供店舖填寫。",
        "• SCT建議 可自行修改；手動只建議用三個值（首選RF／次選RF／建議RF），其他文字會在 Summary 的 SCT建議 區塊另開一欄，以該文字作欄名。",
        f"• SCT建議＝「{_SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE}」的列，是 Buyer {_SHOP_SIMPLIFY_SCT_BUYER_ND_BUYER} 的 ND 貨品，"
        "有機會屬受限制貨品（例如只限澳門出售），請由 Buyer 確認後才改為 RF，不要自行改動。",
    ]
    if len(flags) == 4:
        notes.append(
            "• 指定店舖模式的 Export 一打開只顯示回覆表所需的 30 欄，其餘欄隱藏；需要時在 Excel 選取相鄰欄「取消隱藏」即可。"
        )
    if _SHOP_BLOCKED_QTY_COLUMN in df.columns:
        _blocked_s = pd.to_numeric(df[_SHOP_BLOCKED_QTY_COLUMN], errors="coerce").fillna(0.0)
        _blocked_rows = int((_blocked_s > 0).sum())
        if _blocked_rows:
            notes.append(
                f"• 「Shop Stock Qty (CM)」已扣走「{_SHOP_BLOCKED_QTY_COLUMN}」"
                f"（店舖已安排退貨的件數）：本檔 {_blocked_rows:,} 列、合計 {_blocked_s.sum():,.0f} 件；"
                "扣減後為 0 的列不會再被當作有貨，Shop Stock Value (CM) 亦按同一比例調減"
                "（本模式不匯出 Value 欄）。"
                f"「{_SHOP_BLOCKED_QTY_COLUMN}」欄原樣保留，供與 PowerBI 對帳。"
            )
    notes += [
        "• 顏色、凍結窗格與隱藏欄只影響顯示，不影響篩選、排序與數值。",
        "• 本頁數值按本次輸出的資料計算，換一批資料或改動 Top 比例後會重新產生。",
    ]
    for text in notes:
        row = paragraph(row, text)
    row += 1

    stamp = _shop_simplify_now_hkt().strftime("%Y-%m-%d %H:%M")
    try:
        from version import __version__ as version_text
    except Exception:  # pragma: no cover - 版本檔缺失時不影響匯出
        version_text = ""
    footer = f"本頁由系統自動產生：{stamp}（HKT）" + (f"｜版本 {version_text}" if version_text else "")
    cell = ws.cell(row=row, column=2, value=footer)
    cell.font = note_font
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=9)


def generate_excel_shop_simplify(raw: pd.DataFrame, progress_callback=None, site_view: bool = False) -> BytesIO:
    """產生三個工作表：說明（SCT建議 與顏色說明）、Summary（Adj. MCH1 × Item Status 計數）
    與 Export（回覆明細）。

    說明頁按本檔實際結果列出 SCT建議 的情況表（16 或 4 種）、三個值的分佈、Top N% 與
    顏色說明，並附本次輸出的列數、店舖、比例與金額視窗，方便店舖與 SCT 直接查閱。

    Export 沿用店舖精簡回覆表範本的格式：Brand／SKU 排序、凍結 F2、指定欄寬與隱藏欄、
    數字格式、綠色／紫色重點欄、「無賣」列的淡紅標示、三類紅色底標示
    （Item Status＝Discontinued、RP Type＝ND、過去90日有無存貨＝無存貨），
    以及對比欄底色（兩者皆是＝青綠、兩者皆非＝粉紅）；另 `Adj. SST` 含
    「03 - Private Label」時以醒目橙色底標示。

    `site_view=True`（指定店舖變體）：只顯示 `_SHOP_SIMPLIFY_SITE_VISIBLE_COLUMNS` 所列的
    23 個固定欄，加上 Net Sales (P{N}D) 與六個 Top 判定／對比欄；其餘欄一律隱藏，
    令匯出檔一打開就是回覆表需要的檢視（欄序、顏色、隱藏欄以外的格式完全不變）。
    """
    wb = Workbook()
    # 工作表次序：說明（第一頁，與其他模式一致）→ Summary → Export
    ws_guide = wb.active
    ws_guide.title = _SHOP_SIMPLIFY_GUIDE_SHEET
    ws_summary = wb.create_sheet(title=_SHOP_SIMPLIFY_SUMMARY_SHEET)
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
    fill_private_label = PatternFill(
        start_color=_SHOP_SIMPLIFY_PRIVATE_LABEL_FILL,
        end_color=_SHOP_SIMPLIFY_PRIVATE_LABEL_FILL,
        fill_type="solid",
    )
    fill_buyer_nd = PatternFill(
        start_color=_SHOP_SIMPLIFY_SCT_BUYER_ND_FILL,
        end_color=_SHOP_SIMPLIFY_SCT_BUYER_ND_FILL,
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
    # Adj. SST 含「03 - Private Label」（大小寫、前後空白不拘）→ 醒目橙色底，
    # 一眼分出 Private Label 貨品；該欄沒有其他底色，不會互相覆蓋
    private_label_col = headers.index("Adj. SST") + 1 if "Adj. SST" in headers else None
    private_label_mask = None
    if private_label_col is not None and "Adj. SST" in df.columns:
        _adj_text = df["Adj. SST"].fillna("").astype(str).str.strip().str.lower()
        private_label_mask = _adj_text.str.contains(
            _SHOP_SIMPLIFY_PRIVATE_LABEL_KEYWORD, regex=False
        ).to_numpy(dtype=bool)

    # SCT建議＝「需由Buyer決定」（Buyer 123 - Cara Tsoi 的 ND 貨品）→ 淺黃底，
    # 與紅色底互不重疊（該欄沒有其他底色規則）
    buyer_nd_col = headers.index(_SHOP_SIMPLIFY_SCT_COLUMN) + 1 if _SHOP_SIMPLIFY_SCT_COLUMN in headers else None
    buyer_nd_mask = None
    if buyer_nd_col is not None and _SHOP_SIMPLIFY_SCT_COLUMN in df.columns:
        buyer_nd_mask = _shop_simplify_cell_mask(
            df[_SHOP_SIMPLIFY_SCT_COLUMN], _SHOP_SIMPLIFY_SCT_BUYER_ND_VALUE
        )

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
        if private_label_mask is not None and private_label_mask[row_idx - 2]:
            ws.cell(row=row_idx, column=private_label_col).fill = fill_private_label
        for col_idx, mask in top_yes_rules:
            if mask[row_idx - 2]:
                ws.cell(row=row_idx, column=col_idx).fill = fill_green
        for col_idx, both_mask, neither_mask in compare_rules:
            if both_mask[row_idx - 2]:
                ws.cell(row=row_idx, column=col_idx).fill = fill_compare_both
            elif neither_mask[row_idx - 2]:
                ws.cell(row=row_idx, column=col_idx).fill = fill_compare_neither
        if buyer_nd_mask is not None and buyer_nd_mask[row_idx - 2]:
            ws.cell(row=row_idx, column=buyer_nd_col).fill = fill_buyer_nd
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

    if progress_callback:
        progress_callback(0.90, "建立說明工作表...")
    _write_shop_simplify_guide(ws_guide, df)

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
