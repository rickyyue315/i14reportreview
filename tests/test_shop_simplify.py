"""店舖精簡SKU模式驗證測試。

核心契約：
- 過去90日有無賣過：Sold Qty (P90D) > 0 →「有」，其餘（0／負數／空白）→「無賣」；
- 過去90日有無存貨（v1.93.0 起；前名「過去60日有無存貨」）：Shop Stock Qty (LLM)／(LM)／(CM) 任一 > 0 →「有存貨」，三者同時為 0 →「無存貨」；
- SCT建議、Shop 回覆新的NDRF、如果RF的Safety Qty 一律留空；
- 是否該店的Top 80%：每間店以 Adj. SST 分類為母體，每個分類各自一條 80% 線
  （指定店舖變體改用「是否{店舖}的Top 80%」與「是否同Shop Nature的Top 80%」，不設本欄）；
- 不篩選列，保留來源欄序，六個輸出欄固定排最後；輸出 Summary＋Export 兩個工作表，
  並依 Brand／SKU（文字）排序、套用範本欄寬／隱藏欄／數字格式與顏色標示。
"""
import warnings
warnings.filterwarnings("ignore")
import importlib
from io import BytesIO

import numpy as np
import pandas as pd
import pytest
from openpyxl import load_workbook

from modes import MODE_REGISTRY
from transformer import (
    transform_shop_simplify,
    generate_excel_shop_simplify,
    shop_simplify_stats,
    _SHOP_SIMPLIFY_SOLD_COLUMN,
    _SHOP_SIMPLIFY_STOCK_COLUMN,
    _SHOP_SIMPLIFY_SCT_COLUMN,
    _SHOP_SIMPLIFY_NDRF_COLUMN,
    _SHOP_SIMPLIFY_SAFETY_COLUMN,
    _SHOP_SIMPLIFY_REPLY_COLUMNS,
    _SHOP_SIMPLIFY_SOLD_YES,
    _SHOP_SIMPLIFY_SOLD_NO,
    _SHOP_SIMPLIFY_STOCK_YES,
    _SHOP_SIMPLIFY_STOCK_NO,
    transform_shop_simplify_site,
    shop_simplify_site_stats,
    _SHOP_SIMPLIFY_NATURE_TOP80_COLUMN,
    _SHOP_SIMPLIFY_TOP80_YES,
    _SHOP_SIMPLIFY_TOP80_NO,
    _SHOP_SIMPLIFY_RED,
    _SHOP_SIMPLIFY_SHOP_TOP80_COLUMN,
    _shop_simplify_site_top80_column,
    _shop_simplify_top80_mask,
    _shop_simplify_top80_keys,
)

BASE = {
    "Buyer": "127 - Grace Wong", "Brand": "SUI - SUISSE PROGRAMME", "SKU": 1103775,
    "Item_Shop": "(SUI) 1103775-CAV PRE LIFT ESSENCE WATERHA20-新香港仔",
    "Valuation Class": "Normal Trade Item", "Desc.": "CAV PRE LIFT ESSENCE WATER",
    "Shop Stock Status": "Normal", "Site Code": "HA20", "OM": "OM1",
    "Shop Nature": "Tourist", "RP Type": "RF", "Adj. Safety Stock": 6,
    "Sold Qty (P90D)": 40, "Shop Stock Qty (LM)": 28, "Shop Stock Qty (CM)": 18,
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def _status(df, column):
    return list(df[column])


# ──────────────────────────────────────────────────────────────────────────────
# 1. 註冊與基本行為
# ──────────────────────────────────────────────────────────────────────────────
def test_registry_entry_resolves_to_functions():
    entry = MODE_REGISTRY["店舖精簡SKU模式"]
    mod = importlib.import_module(entry["module"])
    assert callable(getattr(mod, entry["transform"]))
    assert callable(getattr(mod, entry["generate_excel"]))


def test_output_columns_contract_names():
    assert _SHOP_SIMPLIFY_SOLD_COLUMN == "過去90日有無賣過"
    assert _SHOP_SIMPLIFY_STOCK_COLUMN == "過去90日有無存貨"
    assert _SHOP_SIMPLIFY_SCT_COLUMN == "SCT建議"
    assert _SHOP_SIMPLIFY_NDRF_COLUMN == "Shop 回覆新的NDRF"
    assert _SHOP_SIMPLIFY_SAFETY_COLUMN == "如果RF的Safety Qty"
    assert _SHOP_SIMPLIFY_REPLY_COLUMNS == [
        "SCT建議", "Shop 回覆新的NDRF", "如果RF的Safety Qty",
    ]


def test_input_dataframe_not_mutated():
    df = _df([{"SKU": 1}, {"SKU": 2}])
    before = df.copy()
    transform_shop_simplify(df)
    pd.testing.assert_frame_equal(df, before)


# ──────────────────────────────────────────────────────────────────────────────
# 2. 過去90日有無賣過
# ──────────────────────────────────────────────────────────────────────────────
def test_sold_positive_is_yes():
    raw, _ = transform_shop_simplify(_df([{"Sold Qty (P90D)": 1}, {"Sold Qty (P90D)": 293.0}]))
    assert _status(raw, _SHOP_SIMPLIFY_SOLD_COLUMN) == [_SHOP_SIMPLIFY_SOLD_YES] * 2


def test_sold_zero_negative_and_blank_is_no_sale():
    raw, _ = transform_shop_simplify(_df([
        {"Sold Qty (P90D)": 0},
        {"Sold Qty (P90D)": -3},
        {"Sold Qty (P90D)": None},
        {"Sold Qty (P90D)": ""},
    ]))
    assert _status(raw, _SHOP_SIMPLIFY_SOLD_COLUMN) == [_SHOP_SIMPLIFY_SOLD_NO] * 4


# ──────────────────────────────────────────────────────────────────────────────
# 3. 過去90日有無存貨
# ──────────────────────────────────────────────────────────────────────────────
def test_stock_any_positive_is_in_stock():
    raw, _ = transform_shop_simplify(_df([
        {"Shop Stock Qty (LM)": 2, "Shop Stock Qty (CM)": 0},
        {"Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 1},
        {"Shop Stock Qty (LM)": 7, "Shop Stock Qty (CM)": 9},
    ]))
    assert _status(raw, _SHOP_SIMPLIFY_STOCK_COLUMN) == [_SHOP_SIMPLIFY_STOCK_YES] * 3


def test_stock_both_zero_is_no_stock():
    raw, _ = transform_shop_simplify(_df([
        {"Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 0},
        {"Shop Stock Qty (LM)": None, "Shop Stock Qty (CM)": None},
        {"Shop Stock Qty (LM)": -2, "Shop Stock Qty (CM)": 0},
    ]))
    assert _status(raw, _SHOP_SIMPLIFY_STOCK_COLUMN) == [_SHOP_SIMPLIFY_STOCK_NO] * 3


# ──────────────────────────────────────────────────────────────────────────────
# 4. 三個回覆欄一律留空
# ──────────────────────────────────────────────────────────────────────────────
def test_reply_columns_are_blank():
    raw, _ = transform_shop_simplify(_df([{}, {}]))
    for col in _SHOP_SIMPLIFY_REPLY_COLUMNS:
        assert col in raw.columns
        assert list(raw[col]) == ["", ""]


def test_existing_reply_values_are_replaced_by_blank():
    df = _df([{
        "SCT建議": "需要keep",
        "Shop 回覆新的NDRF": "ND",
        "如果RF的Safety Qty": 99,
    }])
    raw, w = transform_shop_simplify(df)
    for col in _SHOP_SIMPLIFY_REPLY_COLUMNS:
        assert list(raw[col]) == [""]
    assert any("一律留空" in x for x in w)


def test_existing_output_columns_are_replaced_and_moved_to_end():
    df = _df([{
        _SHOP_SIMPLIFY_SOLD_COLUMN: "無賣",
        _SHOP_SIMPLIFY_STOCK_COLUMN: "無存貨",
        "SCT建議": "自行決定",
    }])
    raw, _ = transform_shop_simplify(df)
    # 六個輸出欄各只出現一次，且固定排最後
    assert list(raw.columns)[-6:] == [
        _SHOP_SIMPLIFY_SOLD_COLUMN,
        _SHOP_SIMPLIFY_STOCK_COLUMN,
        _SHOP_SIMPLIFY_SHOP_TOP80_COLUMN,
        *_SHOP_SIMPLIFY_REPLY_COLUMNS,
    ]
    assert sum(c == _SHOP_SIMPLIFY_SOLD_COLUMN for c in raw.columns) == 1
    # 舊值被本次計算覆蓋
    assert raw[_SHOP_SIMPLIFY_SOLD_COLUMN].iloc[0] == _SHOP_SIMPLIFY_SOLD_YES
    assert raw[_SHOP_SIMPLIFY_STOCK_COLUMN].iloc[0] == _SHOP_SIMPLIFY_STOCK_YES


# ──────────────────────────────────────────────────────────────────────────────
# 5. 欄序、清理與篩選
# ──────────────────────────────────────────────────────────────────────────────
def test_source_column_order_preserved():
    df = _df([{"SKU": 1}])
    raw, _ = transform_shop_simplify(df)
    cols = list(raw.columns)
    # 解析欄插入慣用位置：SKU Desc 於 Brand 後、Shop Name 於 Site Code 後
    assert cols[:4] == ["Buyer", "Brand", "SKU Desc", "SKU"]
    assert cols[cols.index("Site Code") + 1] == "Shop Name"
    # 其餘來源欄位相對順序不變，六個輸出欄固定排最後
    source_cols = [c for c in cols if c in df.columns and c not in ("SKU Desc", "Shop Name")]
    assert source_cols == list(df.columns)
    assert cols[-6:] == [
        _SHOP_SIMPLIFY_SOLD_COLUMN,
        _SHOP_SIMPLIFY_STOCK_COLUMN,
        _SHOP_SIMPLIFY_SHOP_TOP80_COLUMN,
        *_SHOP_SIMPLIFY_REPLY_COLUMNS,
    ]


def test_value_columns_removed_but_valuation_class_kept():
    df = _df([{
        "Shop Stock Value (LM)": 100.0,
        "Shop Stock Value (CM)": 50.0,
        "Excess Stock Value": 0.0,
        "Valuation Class": "Normal Trade Item",
    }])
    raw, _ = transform_shop_simplify(df)
    for col in ("Shop Stock Value (LM)", "Shop Stock Value (CM)", "Excess Stock Value"):
        assert col not in raw.columns
    # 大小寫敏感：Valuation Class 不含「Value」，必須保留
    assert "Valuation Class" in raw.columns
    assert "SKU Desc" in raw.columns and "Shop Name" in raw.columns
    assert list(raw.columns)[-6] == _SHOP_SIMPLIFY_SOLD_COLUMN

    ws = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))["Export"]
    headers = [c.value for c in ws[1]]
    assert not any(h and "Value" in str(h) for h in headers)
    assert "Valuation Class" in headers


def test_item_shop_parses_desc_and_shop_name():
    df = _df([
        {"SKU": 111186903002.0, "Site Code": "HB97",
         "Item_Shop": "(JDU) 111186903002-HAIR TREAMENT OIL, 01, 40PCHB97-新旺角文華"},
        {"SKU": 111188002003.0, "Site Code": "HB29",
         "Item_Shop": "(JDU) 111188002003-SILKY HAIR OIL, 02, 6PCHB29-重慶站"},
    ])
    raw, _ = transform_shop_simplify(df)
    by_site = raw.set_index("Site Code")
    assert by_site.loc["HB97", "Shop Name"] == "新旺角文華"
    assert by_site.loc["HB97", "SKU Desc"] == "HAIR TREAMENT OIL, 01, 40PC"
    assert by_site.loc["HB29", "Shop Name"] == "重慶站"
    assert by_site.loc["HB29", "SKU Desc"] == "SILKY HAIR OIL, 02, 6PC"


def test_missing_item_shop_warns_and_leaves_derived_columns_blank():
    df = _df([{"SKU": 1}]).drop(columns=["Item_Shop"])
    raw, w = transform_shop_simplify(df)
    assert raw["SKU Desc"].iloc[0] == ""
    assert raw["Shop Name"].iloc[0] == ""
    assert any("Item_Shop" in x for x in w)


def test_sku_normalized_to_text():
    raw, _ = transform_shop_simplify(_df([{"SKU": 1103775.0}]))
    assert raw["SKU"].iloc[0] == "1103775"


def test_footer_and_marketing_rows_removed():
    df = _df([
        {"SKU": 1},
        {"SKU": 2, "Buyer": "Total"},
        {"SKU": 3, "Buyer": "117 - Marketing"},
        {"SKU": 4, "Buyer": "Applied filters: x"},
    ])
    raw, _ = transform_shop_simplify(df)
    assert list(raw["SKU"]) == ["1"]


def test_shop_nature_filter_only_when_selected():
    df = _df([
        {"SKU": 1, "Shop Nature": "Tourist"},
        {"SKU": 2, "Shop Nature": "Mix", "Site Code": "HB29"},
        {"SKU": 3, "Shop Nature": "Non Tourist", "Site Code": "HB50"},
    ])
    raw_all, _ = transform_shop_simplify(df)
    assert len(raw_all) == 3
    raw_mix, _ = transform_shop_simplify(df, shop_natures=["Mix"])
    assert list(raw_mix["SKU"]) == ["2"]


# ──────────────────────────────────────────────────────────────────────────────
# 6. 必要欄位與警告
# ──────────────────────────────────────────────────────────────────────────────
def test_missing_sold_column_raises():
    df = _df([{"SKU": 1}]).drop(columns=["Sold Qty (P90D)"])
    with pytest.raises(ValueError, match="Sold Qty \\(P90D\\)"):
        transform_shop_simplify(df)


def test_missing_both_stock_columns_raises():
    df = _df([{"SKU": 1}]).drop(columns=["Shop Stock Qty (LM)", "Shop Stock Qty (CM)"])
    with pytest.raises(ValueError, match="Shop Stock Qty"):
        transform_shop_simplify(df)


def test_missing_lm_warns_and_uses_cm():
    df = _df([{"SKU": 1, "Shop Stock Qty (CM)": 3}]).drop(columns=["Shop Stock Qty (LM)"])
    raw, w = transform_shop_simplify(df)
    assert raw[_SHOP_SIMPLIFY_STOCK_COLUMN].iloc[0] == _SHOP_SIMPLIFY_STOCK_YES
    assert any("Shop Stock Qty (LM)" in x for x in w)


def test_missing_cm_warns_and_uses_lm():
    df = _df([{"SKU": 1, "Shop Stock Qty (LM)": 0}]).drop(columns=["Shop Stock Qty (CM)"])
    raw, w = transform_shop_simplify(df)
    assert raw[_SHOP_SIMPLIFY_STOCK_COLUMN].iloc[0] == _SHOP_SIMPLIFY_STOCK_NO
    assert any("Shop Stock Qty (CM)" in x for x in w)


def test_empty_after_filter_warns():
    df = _df([{"SKU": 1, "Shop Nature": "Tourist"}])
    raw, w = transform_shop_simplify(df, shop_natures=["Mix"])
    assert raw.empty
    assert any("沒有可輸出" in x for x in w)


# ──────────────────────────────────────────────────────────────────────────────
# 7. 統計與 Excel 輸出
# ──────────────────────────────────────────────────────────────────────────────
def test_stats_counts_sold_stock_both():
    df = _df([
        {"SKU": 1, "Sold Qty (P90D)": 5, "Shop Stock Qty (LM)": 1, "Shop Stock Qty (CM)": 1},
        {"SKU": 2, "Sold Qty (P90D)": 0, "Shop Stock Qty (LM)": 1, "Shop Stock Qty (CM)": 0},
        {"SKU": 3, "Sold Qty (P90D)": 2, "Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 0},
        {"SKU": 4, "Sold Qty (P90D)": 0, "Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 0},
    ])
    raw, _ = transform_shop_simplify(df)
    stats = shop_simplify_stats(raw)
    assert stats["rows"] == 4
    assert stats["sold_no"] == 2
    assert stats["stock_no"] == 2
    assert stats["both_no"] == 1
    assert stats["sites"] == 1
    assert stats["skus"] == 4


# ──────────────────────────────────────────────────────────────────────────────
# 7b. 原模式：每間店各自的 Top 80%（Adj. SST 分組、不設同 Nature 判定）
# ──────────────────────────────────────────────────────────────────────────────
def _top80_flags(raw):
    return dict(zip(zip(raw["Site Code"], raw["SKU"]), raw[_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN]))


def test_shop_top80_uses_each_shops_own_ball_line():
    """HA21 A70/B10/C10/D5/E5（合計 100）→ A、B、C；HA20 A100/B1/C1/D1/E1 → 只有 A。"""
    rows = [{"SKU": s, "Site Code": "HA21", "Adj. SST": "01 - Local", "Sold Qty (P90D)": q}
            for s, q in zip("ABCDE", (70, 10, 10, 5, 5))]
    rows += [{"SKU": s, "Site Code": "HA20", "Adj. SST": "01 - Local", "Sold Qty (P90D)": q}
             for s, q in zip("ABCDE", (100, 1, 1, 1, 1))]
    raw, _ = transform_shop_simplify(_df(rows))
    flags = _top80_flags(raw)
    assert [flags[("HA21", s)] for s in "ABCDE"] == ["是", "是", "是", "否", "否"]
    assert [flags[("HA20", s)] for s in "ABCDE"] == ["是", "否", "否", "否", "否"]


def test_shop_top80_sst_groups_do_not_crowd_each_other():
    """同一店兩個 Adj. SST 各自一條線：Local A100/B5 → 只有 A；Private Label C10/D9 → C、D。
    若不分類（全店合計 124、線 99.2）只會剩 A，故本測試同時驗證分類母體。"""
    rows = [
        {"SKU": "A", "Site Code": "HA21", "Adj. SST": "01 - Local", "Sold Qty (P90D)": 100},
        {"SKU": "B", "Site Code": "HA21", "Adj. SST": "01 - Local", "Sold Qty (P90D)": 5},
        {"SKU": "C", "Site Code": "HA21", "Adj. SST": "03 - Private Label", "Sold Qty (P90D)": 10},
        {"SKU": "D", "Site Code": "HA21", "Adj. SST": "03 - Private Label", "Sold Qty (P90D)": 9},
    ]
    raw, _ = transform_shop_simplify(_df(rows))
    assert list(raw[_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN]) == ["是", "否", "是", "是"]


def test_shop_top80_zero_sales_all_no():
    raw, _ = transform_shop_simplify(_df([
        {"SKU": "A", "Sold Qty (P90D)": 0},
        {"SKU": "B", "Sold Qty (P90D)": -3},
    ]))
    assert list(raw[_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN]) == ["否", "否"]


def test_shop_top80_missing_sst_falls_back_to_single_group_per_shop():
    df = _df([
        {"SKU": "A", "Site Code": "HA21", "Sold Qty (P90D)": 10},
        {"SKU": "B", "Site Code": "HA21", "Sold Qty (P90D)": 1},
    ])
    assert "Adj. SST" not in df.columns
    raw, warnings = transform_shop_simplify(df)
    assert list(raw[_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN]) == ["是", "否"]
    assert any("缺少「Adj. SST」" in x and "Top 80%" in x for x in warnings)


def test_shop_top80_missing_site_code_falls_back_to_whole_table():
    df = _df([
        {"SKU": "A", "Sold Qty (P90D)": 10},
        {"SKU": "B", "Sold Qty (P90D)": 1},
    ]).drop(columns=["Site Code"])
    raw, warnings = transform_shop_simplify(df)
    assert list(raw[_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN]) == ["是", "否"]
    assert any("缺少「Site Code」" in x and "Top 80%" in x for x in warnings)


def test_shop_top80_column_export_width_and_no_nature_column():
    from openpyxl.utils import get_column_letter

    raw, _ = transform_shop_simplify(_df([{"SKU": "A", "Sold Qty (P90D)": 3}]))
    assert _SHOP_SIMPLIFY_NATURE_TOP80_COLUMN not in raw.columns
    assert raw.columns[-4] == _SHOP_SIMPLIFY_SHOP_TOP80_COLUMN
    ws = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))["Export"]
    headers = [c.value for c in ws[1]]
    letter = get_column_letter(headers.index(_SHOP_SIMPLIFY_SHOP_TOP80_COLUMN) + 1)
    assert round(ws.column_dimensions[letter].width, 2) == 14.0


def test_excel_has_summary_and_export_sheets():
    df = _df([
        {"SKU": 1, "Adj. MCH1": "01 - SKIN CARE", "Item Status": "Active"},
        {"SKU": 2, "Adj. MCH1": "01 - SKIN CARE", "Item Status": "Discontinued"},
        {"SKU": 3, "Adj. MCH1": "02 - MAKE UP", "Item Status": "Active"},
    ])
    raw, _ = transform_shop_simplify(df)
    buf = generate_excel_shop_simplify(raw)
    wb = load_workbook(BytesIO(buf.getvalue()))
    assert wb.sheetnames == ["Summary", "Export"]
    assert wb.active.title == "Export"

    ws = wb["Export"]
    headers = [c.value for c in ws[1]]
    assert headers == list(raw.columns)
    assert "SKU Desc" in headers and "Shop Name" in headers
    assert headers[-6:] == [
        "過去90日有無賣過", "過去90日有無存貨", "是否該店的Top 80%",
        "SCT建議", "Shop 回覆新的NDRF", "如果RF的Safety Qty",
    ]
    assert ws.max_row == 1 + len(raw)
    # 來源表頭不改名（Adj. Safety Stock 不轉成 Safety Stock (現時)）
    assert "Adj. Safety Stock" in headers

    summary = wb["Summary"]
    assert summary["A3"].value == "計數 - SKU"
    assert summary["B3"].value == "欄標籤"
    assert summary["A4"].value == "列標籤"
    assert [summary.cell(row=4, column=c).value for c in range(2, 5)] == [
        "Active", "Discontinued", "總計",
    ]
    assert summary["A5"].value == "01 - SKIN CARE"
    assert [summary.cell(row=5, column=c).value for c in range(2, 5)] == [1, 1, 2]
    assert summary["A6"].value == "02 - MAKE UP"
    assert [summary.cell(row=6, column=c).value for c in range(2, 5)] == [1, None, 1]
    assert summary["A7"].value == "總計"
    assert [summary.cell(row=7, column=c).value for c in range(2, 5)] == [2, 1, 3]


def test_summary_blank_groups_sorted_last():
    df = _df([
        {"SKU": 1, "Adj. MCH1": None, "Item Status": "Active"},
        {"SKU": 2, "Adj. MCH1": "01 - SKIN CARE", "Item Status": ""},
        {"SKU": 3, "Adj. MCH1": "01 - SKIN CARE", "Item Status": "Active"},
    ])
    raw, _ = transform_shop_simplify(df)
    summary = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))["Summary"]
    assert summary["A5"].value == "01 - SKIN CARE"
    assert [summary.cell(row=5, column=c).value for c in range(2, 5)] == [1, 1, 2]
    assert summary["A6"].value == "(空白)"
    assert [summary.cell(row=6, column=c).value for c in range(2, 5)] == [1, None, 1]
    assert summary["A7"].value == "總計"
    assert [summary.cell(row=7, column=c).value for c in range(2, 5)] == [2, 1, 3]


def test_output_sorted_by_brand_then_sku_text():
    df = _df([
        {"Brand": "B - B", "SKU": 200},
        {"Brand": "A - A", "SKU": 1000},
        {"Brand": "A - A", "SKU": 999},
    ])
    raw, _ = transform_shop_simplify(df)
    # SKU 以文字排序："1000" < "999"
    assert list(zip(raw["Brand"], raw["SKU"])) == [
        ("A - A", "1000"), ("A - A", "999"), ("B - B", "200"),
    ]


def test_launch_date_converted_to_int():
    raw, _ = transform_shop_simplify(_df([
        {"Launch Date": 20120520.0},
        {"Launch Date": "2020-08-28"},
        {"Launch Date": None},
    ]))
    values = list(raw["Launch Date"])
    assert values[0] == 20120520 and isinstance(values[0], int)
    assert values[1] == 20200828 and isinstance(values[1], int)
    assert pd.isna(values[2])


def test_infinite_turnday_becomes_blank_in_excel():
    raw, _ = transform_shop_simplify(_df([
        {"Turnday (P90D Qty)": np.inf},
        {"Turnday (P90D Qty)": 12.5},
    ]))
    assert pd.isna(raw["Turnday (P90D Qty)"].iloc[0])
    ws = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))["Export"]
    headers = [c.value for c in ws[1]]
    col = headers.index("Turnday (P90D Qty)") + 1
    assert ws.cell(row=2, column=col).value is None


def test_numeric_values_keep_full_precision():
    value = 7.455555555555556
    raw, _ = transform_shop_simplify(_df([{"Shop Monthly Sold Qty (P90D)": value}]))
    ws = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))["Export"]
    headers = [c.value for c in ws[1]]
    written = ws.cell(row=2, column=headers.index("Shop Monthly Sold Qty (P90D)") + 1).value
    assert abs(float(written) - value) < 1e-12


def test_excel_cell_values_and_blank_reply_cells():
    df = _df([
        {"SKU": 1103775, "Sold Qty (P90D)": 40, "Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 0},
    ])
    raw, _ = transform_shop_simplify(df)
    buf = generate_excel_shop_simplify(raw)
    ws = load_workbook(BytesIO(buf.getvalue()))["Export"]
    headers = [c.value for c in ws[1]]
    row = [c.value for c in ws[2]]
    assert row[headers.index(_SHOP_SIMPLIFY_SOLD_COLUMN)] == "有"
    assert row[headers.index(_SHOP_SIMPLIFY_STOCK_COLUMN)] == "無存貨"
    for col in _SHOP_SIMPLIFY_REPLY_COLUMNS:
        assert row[headers.index(col)] is None


def test_excel_format_matches_template():
    df = pd.DataFrame([{
        "Item_Shop": "(SUI) 1103775-CAV PRE LIFT ESSENCE WATERHA20-新香港仔",
        "SKU": 1103775, "Valuation Class": "Normal Trade Item",
        "Brand": "SUI - SUISSE PROGRAMME", "Buyer": "127 - Grace Wong",
        "Shop Stock Status": "Normal", "Site Code": "HA20", "Shop Nature": "Tourist",
        "RP Type": "RF", "Adj. Safety Stock": 6,
        "Adj. MCH1": "01 - SKIN CARE", "Item Status": "Active",
        "Shop Monthly Sold Qty (P90D)": 7.455555555555556,
        "Safety VS Monthly Sold Qty": 0.94,
        "Sold Qty (P90D)": 40, "Shop Stock Qty (LM)": 28, "Shop Stock Qty (CM)": 18,
    }, {
        "Item_Shop": "(SUI) 1106553-THE RICH SERUM 50MLHA20-新香港仔",
        "SKU": 1106553, "Valuation Class": "Normal Trade Item",
        "Brand": "SUI - SUISSE PROGRAMME", "Buyer": "127 - Grace Wong",
        "Shop Stock Status": "Normal", "Site Code": "HA20", "Shop Nature": "Tourist",
        "RP Type": "RF", "Adj. Safety Stock": 6,
        "Adj. MCH1": "01 - SKIN CARE", "Item Status": "Active",
        "Shop Monthly Sold Qty (P90D)": 1.0, "Safety VS Monthly Sold Qty": 0.0,
        "Sold Qty (P90D)": 0, "Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 0,
    }])
    raw, _ = transform_shop_simplify(df)
    ws = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))["Export"]
    headers = [c.value for c in ws[1]]
    col = {name: i + 1 for i, name in enumerate(headers)}

    # 表頭：列高、換行、凍結 F2、縮放 85%
    assert ws.row_dimensions[1].height == 70.5
    assert ws.freeze_panes == "F2"
    assert ws.sheet_view.zoomScale == 85
    header = ws.cell(row=1, column=col["Item_Shop"])
    assert header.alignment.wrap_text is True
    assert header.alignment.horizontal == "center"
    assert header.alignment.vertical == "top"
    assert header.fill.fgColor.rgb[-6:] == "D9E2F3"
    assert ws.cell(row=1, column=col["RP Type"]).fill.fgColor.rgb[-6:] == "CCFF99"
    assert ws.cell(row=1, column=col["Sold Qty (P90D)"]).fill.fgColor.rgb[-6:] == "E4DFEC"
    assert ws.cell(row=1, column=col["Shop 回覆新的NDRF"]).fill.fgColor.rgb[-6:] == "CCFF99"

    # 資料格重點色
    assert ws.cell(row=2, column=col["RP Type"]).fill.fgColor.rgb[-6:] == "CCFF99"
    assert ws.cell(row=2, column=col["Shop Stock Qty (CM)"]).fill.fgColor.rgb[-6:] == "CCFF99"
    assert ws.cell(row=2, column=col["Sold Qty (P90D)"]).fill.fgColor.rgb[-6:] == "E4DFEC"
    # 第二列 Sold Qty = 0 → 無賣 → 淡紅
    sold_cell = ws.cell(row=3, column=col[_SHOP_SIMPLIFY_SOLD_COLUMN])
    assert sold_cell.value == "無賣"
    assert sold_cell.fill.fgColor.rgb[-6:] == "F2DCDB"

    # 欄寬與隱藏欄
    assert ws.column_dimensions["A"].hidden is True
    assert ws.column_dimensions["C"].hidden is True
    assert round(ws.column_dimensions["A"].width, 2) == 79.71
    assert round(ws.column_dimensions["E"].width, 2) == 46.43

    # 數字格式
    assert ws.cell(row=2, column=col["Shop Monthly Sold Qty (P90D)"]).number_format.startswith("_(* #,##0")
    assert ws.cell(row=2, column=col["Safety VS Monthly Sold Qty"]).number_format == "0%"


def test_excel_red_highlight_rules():
    """Item Status＝Discontinued、RP Type＝ND、過去90日有無存貨＝無存貨 → 紅色底；
    比對忽略前後空白與大小寫；RP Type 的紅色底蓋過該欄綠色；「無賣」淡紅不受影響。"""
    df = _df([
        {"SKU": 1, "Item Status": "Discontinued", "RP Type": "ND",
         "Sold Qty (P90D)": 0, "Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 0},
        {"SKU": 2, "Item Status": "Active", "RP Type": "RF",
         "Sold Qty (P90D)": 5, "Shop Stock Qty (LM)": 3, "Shop Stock Qty (CM)": 1},
        {"SKU": 3, "Item Status": " discontinued ", "RP Type": "nd",
         "Sold Qty (P90D)": 1, "Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 2},
    ])
    raw, _ = transform_shop_simplify(df)
    ws = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))["Export"]
    col = {name: i + 1 for i, name in enumerate([c.value for c in ws[1]])}

    def rgb(row, name):
        return str(ws.cell(row=row, column=col[name]).fill.start_color.rgb)[-6:]

    # 第一列：三個條件同時成立
    assert rgb(2, "Item Status") == _SHOP_SIMPLIFY_RED
    assert rgb(2, "RP Type") == _SHOP_SIMPLIFY_RED
    assert ws.cell(row=2, column=col[_SHOP_SIMPLIFY_STOCK_COLUMN]).value == _SHOP_SIMPLIFY_STOCK_NO
    assert rgb(2, _SHOP_SIMPLIFY_STOCK_COLUMN) == _SHOP_SIMPLIFY_RED
    # 「無賣」淡紅與新紅色底各自獨立（同一列可同時出現）
    assert rgb(2, _SHOP_SIMPLIFY_SOLD_COLUMN) == "F2DCDB"

    # 第二列：沒有命中任何條件，RP Type 維持綠色、其餘無填色
    assert ws.cell(row=3, column=col["Item Status"]).fill.fill_type is None
    assert rgb(3, "RP Type") == "CCFF99"
    assert ws.cell(row=3, column=col[_SHOP_SIMPLIFY_STOCK_COLUMN]).fill.fill_type is None

    # 第三列：來源值前後空白／小寫亦命中（Item Status、RP Type 紅底），有存貨不套紅底
    assert rgb(4, "Item Status") == _SHOP_SIMPLIFY_RED
    assert rgb(4, "RP Type") == _SHOP_SIMPLIFY_RED
    assert ws.cell(row=4, column=col[_SHOP_SIMPLIFY_STOCK_COLUMN]).value == _SHOP_SIMPLIFY_STOCK_YES
    assert ws.cell(row=4, column=col[_SHOP_SIMPLIFY_STOCK_COLUMN]).fill.fill_type is None


def test_excel_red_highlight_skips_missing_columns():
    """來源缺 Item Status 欄時該條件自動略過，其餘兩個條件照常標示。"""
    df = _df([
        {"SKU": 1, "RP Type": "ND", "Sold Qty (P90D)": 1,
         "Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 0},
    ])
    raw, _ = transform_shop_simplify(df)
    assert "Item Status" not in raw.columns
    ws = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))["Export"]
    col = {name: i + 1 for i, name in enumerate([c.value for c in ws[1]])}
    assert str(ws.cell(row=2, column=col["RP Type"]).fill.start_color.rgb)[-6:] == _SHOP_SIMPLIFY_RED
    assert str(ws.cell(row=2, column=col[_SHOP_SIMPLIFY_STOCK_COLUMN]).fill.start_color.rgb)[-6:] == _SHOP_SIMPLIFY_RED


def test_excel_freeze_panes_after_sku_desc():
    df = _df([{"SKU": 1}])
    raw, _ = transform_shop_simplify(df)
    buf = generate_excel_shop_simplify(raw)
    ws = load_workbook(BytesIO(buf.getvalue()))["Export"]
    from openpyxl.utils import get_column_letter
    headers = [c.value for c in ws[1]]
    expected = f"{get_column_letter(headers.index('SKU Desc') + 2)}2"
    assert ws.freeze_panes == expected
    assert ws.auto_filter.ref == ws.dimensions


def test_empty_dataframe_still_writes_header():
    raw = pd.DataFrame(columns=list(BASE) + [
        _SHOP_SIMPLIFY_SOLD_COLUMN, _SHOP_SIMPLIFY_STOCK_COLUMN, *_SHOP_SIMPLIFY_REPLY_COLUMNS,
    ])
    buf = generate_excel_shop_simplify(raw)
    wb = load_workbook(BytesIO(buf.getvalue()))
    ws = wb["Export"]
    assert ws.max_row == 1
    assert [c.value for c in ws[1]][-5] == _SHOP_SIMPLIFY_SOLD_COLUMN


# ──────────────────────────────────────────────────────────────────────────────
# 8. 指定店舖模式：只輸出該店、兩個 Top 80% 判定
# ──────────────────────────────────────────────────────────────────────────────
def _site_rows():
    """HA21（Non Tourist）A70／B10／C10／D5／E5；HA20 同 Nature 的 A 賣 1000；
    HA19（Tourist）的 E 賣 9999（不得計入 Non Tourist 的 Top 80% 母體）。"""
    return [
        {"SKU": "A", "Site Code": "HA21", "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 70},
        {"SKU": "B", "Site Code": "HA21", "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 10},
        {"SKU": "C", "Site Code": "HA21", "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 10},
        {"SKU": "D", "Site Code": "HA21", "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 5},
        {"SKU": "E", "Site Code": "HA21", "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 5},
        {"SKU": "A", "Site Code": "HA20", "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 1000},
        {"SKU": "E", "Site Code": "HA19", "Shop Nature": "Tourist", "Sold Qty (P90D)": 9999},
    ]


def test_site_mode_registry_entry_resolves_to_functions():
    entry = MODE_REGISTRY["店舖精簡SKU模式(指定店舖)"]
    mod = importlib.import_module(entry["module"])
    assert callable(getattr(mod, entry["transform"]))
    assert callable(getattr(mod, entry["generate_excel"]))


def test_site_mode_outputs_only_selected_shop():
    raw, _ = transform_shop_simplify_site(_df(_site_rows()), "HA21")
    assert list(raw["Site Code"]) == ["HA21"] * 5
    assert sorted(raw["SKU"].tolist()) == ["A", "B", "C", "D", "E"]
    assert len(raw) == 5


def test_site_top80_uses_cumulative_80_with_tie_group():
    raw, _ = transform_shop_simplify_site(_df(_site_rows()), "HA21")
    flags = dict(zip(raw["SKU"], raw[_shop_simplify_site_top80_column("HA21")]))
    # 70＋10 已達 100 的 80%；同為 10 件的 C 一併納入，5 件的 D／E 排除
    assert flags == {"A": _SHOP_SIMPLIFY_TOP80_YES, "B": _SHOP_SIMPLIFY_TOP80_YES,
                     "C": _SHOP_SIMPLIFY_TOP80_YES, "D": _SHOP_SIMPLIFY_TOP80_NO,
                     "E": _SHOP_SIMPLIFY_TOP80_NO}


def test_nature_top80_aggregates_same_nature_only():
    raw, _ = transform_shop_simplify_site(_df(_site_rows()), "HA21")
    flags = dict(zip(raw["SKU"], raw[_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN]))
    # Non Tourist 合計：A 1070、B 10、C 10、D 5、E 5；80% 門檻 880 → 只有 A
    # HA19（Tourist）的 E 9999 不得計入，否則 E 會變成「是」
    assert flags == {"A": _SHOP_SIMPLIFY_TOP80_YES, "B": _SHOP_SIMPLIFY_TOP80_NO,
                     "C": _SHOP_SIMPLIFY_TOP80_NO, "D": _SHOP_SIMPLIFY_TOP80_NO,
                     "E": _SHOP_SIMPLIFY_TOP80_NO}


def test_top80_mask_zero_and_negative_sales_all_false():
    qty = pd.Series([0.0, 0.0, -5.0, np.nan], index=["A", "B", "C", "D"])
    assert list(_shop_simplify_top80_mask(qty)) == [False] * 4
    assert list(_shop_simplify_top80_mask(pd.Series(dtype=float))) == []


def test_site_top80_column_name_and_order_before_reply_columns():
    raw, _ = transform_shop_simplify_site(_df(_site_rows()), "HA21")
    site_col = _shop_simplify_site_top80_column("HA21")
    assert site_col == "是否HA21的Top 80%"
    cols = list(raw.columns)
    index = cols.index(_SHOP_SIMPLIFY_STOCK_COLUMN)
    assert cols[index + 1:index + 4] == [site_col, _SHOP_SIMPLIFY_NATURE_TOP80_COLUMN, _SHOP_SIMPLIFY_SCT_COLUMN]
    assert cols[-3:] == _SHOP_SIMPLIFY_REPLY_COLUMNS
    for col in _SHOP_SIMPLIFY_REPLY_COLUMNS:
        assert list(raw[col]) == [""] * len(raw)


def test_site_mode_missing_shop_nature_leaves_column_blank_and_warns():
    df = _df(_site_rows()).drop(columns=["Shop Nature"])
    raw, warnings = transform_shop_simplify_site(df, "HA21")
    assert list(raw[_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN]) == [""] * len(raw)
    assert any("缺少「Shop Nature」" in x for x in warnings)


def test_site_mode_blank_shop_nature_leaves_column_blank_and_warns():
    rows = [dict(r, **{"Shop Nature": ""}) for r in _site_rows()]
    raw, warnings = transform_shop_simplify_site(_df(rows), "HA21")
    assert list(raw[_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN]) == [""] * len(raw)
    assert any("Shop Nature 空白" in x for x in warnings)


def test_site_mode_zero_total_sales_all_no_and_warns():
    rows = [dict(r, **{"Sold Qty (P90D)": 0}) for r in _site_rows()]
    raw, warnings = transform_shop_simplify_site(_df(rows), "HA21")
    site_col = _shop_simplify_site_top80_column("HA21")
    assert list(raw[site_col]) == [_SHOP_SIMPLIFY_TOP80_NO] * len(raw)
    assert list(raw[_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN]) == [_SHOP_SIMPLIFY_TOP80_NO] * len(raw)
    assert any("合計為 0" in x for x in warnings)


def test_site_mode_invalid_site_raises():
    with pytest.raises(ValueError, match="HA99"):
        transform_shop_simplify_site(_df(_site_rows()), "HA99")
    with pytest.raises(ValueError, match="未提供指定店舖"):
        transform_shop_simplify_site(_df(_site_rows()), "  ")


def test_site_mode_case_insensitive_site_match():
    raw, warnings = transform_shop_simplify_site(_df(_site_rows()), " ha21 ")
    assert set(raw["Site Code"]) == {"HA21"}
    assert _shop_simplify_site_top80_column("HA21") in raw.columns
    assert any("大小寫不敏感" in x for x in warnings)


def test_site_mode_single_shop_nature_warns_nature_equals_shop():
    df = _df([{"SKU": 1, "Site Code": "HA21", "Shop Nature": "Non Tourist", "Sold Qty (P90D)": 5}])
    _, warnings = transform_shop_simplify_site(df, "HA21")
    assert any("只有 1 間店舖" in x for x in warnings)


def test_site_mode_rerun_replaces_old_top80_columns():
    first, _ = transform_shop_simplify_site(_df(_site_rows()), "HA21")
    second, _ = transform_shop_simplify_site(first, "HA21")
    site_col = _shop_simplify_site_top80_column("HA21")
    assert list(second.columns).count(site_col) == 1
    assert list(second.columns).count(_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN) == 1


def test_site_mode_stats_counts():
    raw, _ = transform_shop_simplify_site(_df(_site_rows()), "HA21")
    stats = shop_simplify_site_stats(raw, "HA21")
    assert stats["rows"] == 5
    assert stats["skus"] == 5
    assert stats["site_top80"] == 3
    assert stats["nature_top80"] == 1
    assert stats["neither"] == 2
    assert stats["sold_no"] == 0
    assert stats["stock_no"] == 0
    assert stats["site"] == "HA21"
    assert stats["nature"] == "Non Tourist"


def test_site_mode_excel_headers_widths_and_summary():
    from openpyxl.utils import get_column_letter

    raw, _ = transform_shop_simplify_site(_df(_site_rows()), "HA21")
    wb = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))
    ws = wb["Export"]
    headers = [c.value for c in ws[1]]
    site_col = _shop_simplify_site_top80_column("HA21")
    assert headers == list(raw.columns)
    assert headers[-5:] == [
        site_col, _SHOP_SIMPLIFY_NATURE_TOP80_COLUMN,
        *_SHOP_SIMPLIFY_REPLY_COLUMNS,
    ]
    # 指定店舖變體不加原模式的「是否該店的Top 80%」，避免兩個同義欄並存
    assert _SHOP_SIMPLIFY_SHOP_TOP80_COLUMN not in headers
    assert round(ws.column_dimensions[get_column_letter(headers.index(site_col) + 1)].width, 2) == 14.0
    assert round(ws.column_dimensions[get_column_letter(headers.index(_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN) + 1)].width, 2) == 16.0
    assert ws.freeze_panes == f"{get_column_letter(headers.index('SKU Desc') + 2)}2"
    assert wb.sheetnames == ["Summary", "Export"]


def test_site_mode_excel_red_highlight_rules():
    """指定店舖變體共用同一匯出引擎：Discontinued／ND／無存貨同樣以紅色底標示。"""
    df = _df([
        {"SKU": "A", "Site Code": "HA21", "Item Status": "Discontinued", "RP Type": "ND",
         "Sold Qty (P90D)": 5, "Shop Stock Qty (LM)": 0, "Shop Stock Qty (CM)": 0},
        {"SKU": "B", "Site Code": "HA21", "Item Status": "Active", "RP Type": "RF",
         "Sold Qty (P90D)": 1, "Shop Stock Qty (LM)": 2, "Shop Stock Qty (CM)": 0},
    ])
    raw, _ = transform_shop_simplify_site(df, "HA21")
    ws = load_workbook(BytesIO(generate_excel_shop_simplify(raw).getvalue()))["Export"]
    col = {name: i + 1 for i, name in enumerate([c.value for c in ws[1]])}

    def rgb(row, name):
        return str(ws.cell(row=row, column=col[name]).fill.start_color.rgb)[-6:]

    assert rgb(2, "Item Status") == _SHOP_SIMPLIFY_RED
    assert rgb(2, "RP Type") == _SHOP_SIMPLIFY_RED
    assert rgb(2, _SHOP_SIMPLIFY_STOCK_COLUMN) == _SHOP_SIMPLIFY_RED
    assert rgb(3, "RP Type") == "CCFF99"
    assert ws.cell(row=3, column=col[_SHOP_SIMPLIFY_STOCK_COLUMN]).fill.fill_type is None


# ──────────────────────────────────────────────────────────────────────────────
# 9. 指定店舖模式：Top 80% 以 Adj. SST 分類為母體
# ──────────────────────────────────────────────────────────────────────────────
def _sst_rows():
    """HA21（Non Tourist）分兩個 Adj. SST；HA20 同 Nature 以跨店銷量影響 Nature 判定：
    - Local：HA21 A70／B10／C10／D5／H1；HA20 A1000、D5000
    - Private Label：HA21 E5／F3／G2；HA20 G1000
    """
    def row(sku, sst, sold, site="HA21"):
        return {"SKU": sku, "Site Code": site, "Shop Nature": "Non Tourist",
                "Adj. SST": sst, "Sold Qty (P90D)": sold}

    return [
        row("A", "01 - Local", 70),
        row("B", "01 - Local", 10),
        row("C", "01 - Local", 10),
        row("D", "01 - Local", 5),
        row("H", "01 - Local", 1),
        row("E", "03 - Private Label", 5),
        row("F", "03 - Private Label", 3),
        row("G", "03 - Private Label", 2),
        row("A", "01 - Local", 1000, site="HA20"),
        row("D", "01 - Local", 5000, site="HA20"),
        row("G", "03 - Private Label", 1000, site="HA20"),
    ]


def test_top80_keys_computed_within_each_adj_sst_group():
    keys = _shop_simplify_top80_keys(
        pd.Series(["L", "L", "L", "P", "P"]),
        pd.Series(["A", "B", "C", "D", "E"]),
        pd.Series([70, 10, 10, 5, 3]),
    )
    # L：A70＋B10 達 80，C 同為 10 件整組納入；P：D5＋E3 剛好 8（100%）全組納入
    assert keys == {("L", "A"), ("L", "B"), ("L", "C"), ("P", "D"), ("P", "E")}


def test_site_top80_isolated_per_adj_sst():
    raw, _ = transform_shop_simplify_site(_df(_sst_rows()), "HA21")
    flags = dict(zip(raw["SKU"], raw[_shop_simplify_site_top80_column("HA21")]))
    # Local：A70＋B10 達 95 的 80%，C 同為 10 件整組納入；D、H 排除
    # Private Label：E5＋F3 達 10 的 80%，G 排除
    assert flags == {
        "A": _SHOP_SIMPLIFY_TOP80_YES, "B": _SHOP_SIMPLIFY_TOP80_YES,
        "C": _SHOP_SIMPLIFY_TOP80_YES, "D": _SHOP_SIMPLIFY_TOP80_NO,
        "E": _SHOP_SIMPLIFY_TOP80_YES, "F": _SHOP_SIMPLIFY_TOP80_YES,
        "G": _SHOP_SIMPLIFY_TOP80_NO, "H": _SHOP_SIMPLIFY_TOP80_NO,
    }


def test_nature_top80_isolated_per_adj_sst_across_shops():
    raw, _ = transform_shop_simplify_site(_df(_sst_rows()), "HA21")
    flags = dict(zip(raw["SKU"], raw[_SHOP_SIMPLIFY_NATURE_TOP80_COLUMN]))
    # Nature Local 合計：D5005＋A1070＋B10＋C10＋H1；80% 門檻 4876.8 → 只有 D
    # Nature Private Label 合計：G1002＋E5＋F3；80% 門檻 808 → 只有 G
    assert flags == {
        "A": _SHOP_SIMPLIFY_TOP80_NO, "B": _SHOP_SIMPLIFY_TOP80_NO,
        "C": _SHOP_SIMPLIFY_TOP80_NO, "D": _SHOP_SIMPLIFY_TOP80_YES,
        "E": _SHOP_SIMPLIFY_TOP80_NO, "F": _SHOP_SIMPLIFY_TOP80_NO,
        "G": _SHOP_SIMPLIFY_TOP80_YES, "H": _SHOP_SIMPLIFY_TOP80_NO,
    }


def test_site_mode_stats_include_adj_sst_group_count():
    raw, _ = transform_shop_simplify_site(_df(_sst_rows()), "HA21")
    stats = shop_simplify_site_stats(raw, "HA21")
    assert stats["rows"] == 8
    assert stats["skus"] == 8
    assert stats["sst_groups"] == 2
    assert stats["site_top80"] == 5
    assert stats["nature_top80"] == 2
    assert stats["neither"] == 1


def test_site_mode_missing_adj_sst_falls_back_and_warns():
    _, warnings = transform_shop_simplify_site(_df(_site_rows()), "HA21")
    assert any("缺少「Adj. SST」" in x for x in warnings)


def test_site_mode_only_parses_selected_shop_rows(monkeypatch):
    """大檔效能契約：全表只做向量化清理，Item_Shop 逐列解析只套用在指定店舖的列。"""
    import modes.shop_simplify as shop_simplify_module

    calls = []
    original = shop_simplify_module._compute_raw_extra_columns

    def counting(df):
        calls.append(len(df))
        return original(df)

    monkeypatch.setattr(shop_simplify_module, "_compute_raw_extra_columns", counting)
    rows = [
        {"SKU": f"X{i}", "Site Code": "HA20", "Shop Nature": "Non Tourist",
         "Adj. SST": "01 - Local", "Sold Qty (P90D)": 1}
        for i in range(50)
    ]
    rows += [
        {"SKU": "A", "Site Code": "HA21", "Shop Nature": "Non Tourist",
         "Adj. SST": "01 - Local", "Sold Qty (P90D)": 10},
        {"SKU": "B", "Site Code": "HA21", "Shop Nature": "Non Tourist",
         "Adj. SST": "01 - Local", "Sold Qty (P90D)": 5},
    ]
    transform_shop_simplify_site(_df(rows), "HA21")
    assert calls == [2]
