# -*- coding: utf-8 -*-
"""Safety Stock Review SUMMARY 區塊 ⑤–⑦（影響範圍統計／問題類型樞紐分析）測試。

覆蓋：與「店舖問題 SKU」同一批列、同一問題類型（數字與該表及使用者自行做成的樞紐分析
一致）、Brand／Buyer／店舖三個層級的對象數與無需行動、只可增加模式、空白鍵、空表與缺欄
守衛、Buyer＞Brand 小計結構、店舖表排序與 Shop Name、SUMMARY 版面（兩層表頭、黃底小計、
摺疊群組、自動篩選、欄寬上限）與端到端輸出。
"""

import datetime
from io import BytesIO

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter

import transformer  # noqa: F401  先載入墊片，避免 modes 之間循環引用
from transformer import _add_combined_recommendation_columns, transform_combined, generate_excel_combined
from modes.issues import _build_combined_shop_issues
from modes.safety_review import (
    _SSR_ISSUE_TYPES,
    _SSR_PIVOT_BLANK,
    _SSR_PIVOT_LEVEL_COL,
    _SSR_PIVOT_TOTAL,
    _build_combined_impact_summary,
    _build_combined_issue_pivot,
    _write_issue_pivot_block,
    _write_summary_sheet,
)

_RATES_O = {"default": 0.5, "tourist": 0.3, "mix": 0.4}
_RATES_P = {"default": 0.5, "tourist": 0.5, "mix": 0.5}
_ND_COL = "Recommendation of new Safety Qty / ND"
_ND, _SHORT, _HIGH = _SSR_ISSUE_TYPES
_MEASURES = ("行數", "加總 - Safety Qty Change")
_GROUPS = (_ND, _SHORT, _HIGH, _SSR_PIVOT_TOTAL)
_MEASURE_COLS = [f"{t} {m}" for t in _GROUPS for m in _MEASURES]


def _base_row(**kw):
    row = {
        "Buyer": "B1",
        "Brand": "BrandA",
        "SKU": "1001",
        "Site Code": "S01",
        "Item_Shop": "Some Product-S01",
        "Shop Stock Status": "Normal",
        "Adj. Safety Stock": 5,
        "Shop Monthly Sold Qty (P90D)": 10,
        "Sold Qty (CM - MTD)": 8.0,
        "Shop Nature": "Local",
        "RP Type": "RF",
        "Shop Stock Qty (CM)": 0,
        "Supply Source": "Warehouse",
        "OM": "OM1",
        "Adj. MCH1": "MchA",
        "Adj. SST": "SstA",
    }
    row.update(kw)
    return row


def _df(rows):
    return pd.DataFrame([_base_row(**r) for r in rows])


def _launch_date(days_ago):
    return (datetime.date.today() - datetime.timedelta(days=days_ago)).strftime("%Y/%m/%d")


def _run(rows, **kw):
    params = dict(
        rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both",
    )
    params.update(kw)
    return _add_combined_recommendation_columns(_df(rows), **params)


def _designed_rows():
    """七列、3 個 Buyer／3 個 Brand／4 間店舖（月均 10、MTD 8、15/30 日 → Y=7、X=5）：
    ① B1/BrandA/SKU 1/S01 增加 +2             → Safety 設定不足
    ② B1/BrandA/SKU 1/S02 不變                → 不在店舖問題 SKU
    ③ B1/BrandA/SKU 2/S01 減少 −3             → Safety 設定過高
    ④ B2/BrandA/SKU 3/S01 增加 +4             → Safety 設定不足
    ⑤ B2/BrandB/SKU 4/S02 P90D 0 舊品 −4      → Safety 設定過高（Safety Qty Change < 0 優先）
    ⑥ B2/BrandB/SKU 5/S03 P90D 0 舊品、現有 0 → ND 退場
    ⑦ B3/BrandC/SKU 6/S04 不變                → 不在店舖問題 SKU
    """
    old_p90d0 = {
        "Shop Stock Status": "Overstock (P90D 0 Sales)",
        "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0,
        "Launch Date": _launch_date(200),
    }
    return [
        {"Buyer": "B1", "Brand": "BrandA", "SKU": "1", "Site Code": "S01", "Adj. Safety Stock": 5},
        {"Buyer": "B1", "Brand": "BrandA", "SKU": "1", "Site Code": "S02", "Adj. Safety Stock": 7},
        {"Buyer": "B1", "Brand": "BrandA", "SKU": "2", "Site Code": "S01", "Adj. Safety Stock": 10},
        {"Buyer": "B2", "Brand": "BrandA", "SKU": "3", "Site Code": "S01", "Adj. Safety Stock": 3},
        {"Buyer": "B2", "Brand": "BrandB", "SKU": "4", "Site Code": "S02", "Adj. Safety Stock": 4, **old_p90d0},
        {"Buyer": "B2", "Brand": "BrandB", "SKU": "5", "Site Code": "S03", "Adj. Safety Stock": 0, **old_p90d0},
        {"Buyer": "B3", "Brand": "BrandC", "SKU": "6", "Site Code": "S04", "Adj. Safety Stock": 7},
    ]


def _vals(rec):
    """樞紐列的數值部分：依 問題類型 → 總計 的次序，各取 (行數, 加總)。"""
    return [rec[c] for c in _MEASURE_COLS]


def _impact_row(imp, label):
    row = imp[imp["統計對象"] == label]
    assert len(row) == 1
    return row.iloc[0]


def _expect(frame):
    """獨立於 builder 的做法：直接以問題類型 groupby 店舖問題 SKU 的列，回傳 8 個數值。"""
    g = frame.groupby("問題類型")["Safety Qty Change"].agg(["size", "sum"])
    out = []
    for t in _SSR_ISSUE_TYPES:
        out += [int(g.loc[t, "size"]), int(round(float(g.loc[t, "sum"])))] if t in g.index else [0, 0]
    return out + [int(g["size"].sum()), int(round(float(g["sum"].sum())))]


# ── 與「店舖問題 SKU」同一口徑 ──

def test_issue_types_follow_shop_issue_sheet():
    """問題類型與「店舖問題 SKU」相同：不足／過高按 Safety Qty Change 正負，其後才是 ND 退場。"""
    issues = _build_combined_shop_issues(_run(_designed_rows()))
    assert len(issues) == 5
    assert dict(zip(issues["SKU"], issues["問題類型"])) == {"1": _SHORT, "2": _HIGH, "3": _SHORT, "4": _HIGH, "5": _ND}


def test_pivot_equals_manual_pivot_of_shop_issue_sheet():
    """樞紐每列（含 Buyer 小計、Brand 明細、店舖、總計）皆等於直接 groupby「店舖問題 SKU」的結果。"""
    rows = []
    for i in range(60):
        dead = i % 11 == 0
        rows.append({
            "Buyer": f"B{i % 3}", "Brand": f"Brand{i % 5}", "SKU": str(100 + i), "Site Code": f"S{i % 7:02d}",
            # 零銷售舊品：i 為 22 的倍數時現有 Safety 0 → ND 退場，其餘 → 設定過高
            "Adj. Safety Stock": 0 if (dead and i % 22 == 0) else (i * 7) % 13,
            "Shop Monthly Sold Qty (P90D)": 0 if dead else 10,
            "Sold Qty (CM - MTD)": 0.0 if dead else 8.0,
            "Launch Date": _launch_date(200),
            **({"Shop Stock Status": "Overstock (P90D 0 Sales)"} if dead else {}),
        })
    issues = _build_combined_shop_issues(_run(rows))
    assert set(issues["問題類型"]) == set(_SSR_ISSUE_TYPES)
    buyer = issues["Buyer"].fillna("").astype(str).str.strip().mask(lambda s: s == "", _SSR_PIVOT_BLANK)
    brand = issues["Brand"].fillna("").astype(str).str.strip().mask(lambda s: s == "", _SSR_PIVOT_BLANK)
    site = issues["Site Code"].fillna("").astype(str).str.strip().mask(lambda s: s == "", _SSR_PIVOT_BLANK)

    tab = _build_combined_issue_pivot(issues, "buyer_brand")
    assert tab[_SSR_PIVOT_LEVEL_COL].value_counts().to_dict() == {
        "Brand": len(pd.DataFrame({"b": buyer, "g": brand}).drop_duplicates()),
        "Buyer": buyer.nunique(),
        "總計": 1,
    }
    cur = None
    for rec in tab.to_dict("records"):
        level, label = rec[_SSR_PIVOT_LEVEL_COL], rec["Buyer / Brand"]
        if level == "Buyer":
            cur = label
            sub = issues[buyer == label]
        elif level == "Brand":
            sub = issues[(buyer == cur) & (brand == label)]
        else:
            sub = issues
        assert _vals(rec) == _expect(sub)

    tab = _build_combined_issue_pivot(issues, "shop")
    assert tab[_SSR_PIVOT_LEVEL_COL].value_counts().to_dict() == {"店舖": site.nunique(), "總計": 1}
    for rec in tab.to_dict("records"):
        sub = issues if rec[_SSR_PIVOT_LEVEL_COL] == "總計" else issues[site == rec["Site Code"]]
        assert _vals(rec) == _expect(sub)
    # 總計列與全表 groupby 一致；三個問題類型相加等於總計（互斥）
    total = tab.iloc[-1]
    assert _vals(total) == _expect(issues)
    assert total[f"{_SSR_PIVOT_TOTAL} 行數"] == sum(total[f"{t} 行數"] for t in _SSR_ISSUE_TYPES)


# ── ⑤ 影響範圍統計 ──

def test_impact_summary_counts_entities_by_issue_type():
    out = _run(_designed_rows())
    issues = _build_combined_shop_issues(out)
    imp = _build_combined_impact_summary(out, issues=issues)

    assert list(imp.columns) == ["統計對象", "總數", _ND, _SHORT, _HIGH, "無需行動", "說明"]
    assert imp["統計對象"].tolist() == ["Brand", "Buyer", "店舖"]
    # 欄序：總數／ND 退場／Safety 設定不足／Safety 設定過高／無需行動
    assert _impact_row(imp, "Brand")[1:6].tolist() == [3, 1, 1, 2, 1]
    assert _impact_row(imp, "Buyer")[1:6].tolist() == [3, 1, 2, 2, 1]
    # S03 只有 ND 退場（現有 Safety 0）；S04 沒有任何需行動列
    assert _impact_row(imp, "店舖")[1:6].tolist() == [4, 1, 1, 2, 1]
    # 未傳 issues 時即時建立，結果相同
    pd.testing.assert_frame_equal(_build_combined_impact_summary(out), imp)
    # 寫入 Excel 走 itertuples：數值格須為 Python int 才會套用千分位與置中
    numeric = imp.drop(columns=["統計對象", "說明"]).itertuples(index=False, name=None)
    assert all(type(v) is int for tup in numeric for v in tup)


def test_impact_summary_agrees_with_pivot_tables():
    """⑤ 各問題類型的對象數＝⑥／⑦ 中該類行數 > 0 的列數；無需行動＝總數 − 有列的對象數。"""
    out = _run(_designed_rows())
    issues = _build_combined_shop_issues(out)
    imp = _build_combined_impact_summary(out, issues=issues)
    bb = _build_combined_issue_pivot(issues, "buyer_brand")
    brands = bb[bb[_SSR_PIVOT_LEVEL_COL] == "Brand"]
    buyers = bb[bb[_SSR_PIVOT_LEVEL_COL] == "Buyer"]
    shops = _build_combined_issue_pivot(issues, "shop")
    shops = shops[shops[_SSR_PIVOT_LEVEL_COL] == "店舖"]
    for t in _SSR_ISSUE_TYPES:
        # ⑥ 的 Brand 按 Buyer 分列（同一 Brand 可多於一列），故以不重複名稱核對
        hit = brands[brands[f"{t} 行數"] > 0]
        assert _impact_row(imp, "Brand")[t] == hit["Buyer / Brand"].nunique()
        assert _impact_row(imp, "Buyer")[t] == int((buyers[f"{t} 行數"] > 0).sum())
        assert _impact_row(imp, "店舖")[t] == int((shops[f"{t} 行數"] > 0).sum())
    assert _impact_row(imp, "Brand")["無需行動"] == 3 - brands["Buyer / Brand"].nunique()
    assert _impact_row(imp, "Buyer")["無需行動"] == 3 - len(buyers)
    assert _impact_row(imp, "店舖")["無需行動"] == 4 - len(shops)


def test_impact_summary_increase_only_has_no_nd_exit():
    """只可增加：退場建議不套用，ND 退場恆為 0；原本減少的列歸無需行動。"""
    out = _run(_designed_rows(), safety_mode="increase_only")
    issues = _build_combined_shop_issues(out)
    assert set(issues["問題類型"]) == {_SHORT}

    imp = _build_combined_impact_summary(out, issues=issues)
    assert _impact_row(imp, "Brand")[1:6].tolist() == [3, 0, 1, 0, 2]
    assert _impact_row(imp, "Buyer")[1:6].tolist() == [3, 0, 2, 0, 1]
    assert _impact_row(imp, "店舖")[1:6].tolist() == [4, 0, 1, 0, 3]


def test_already_nd_rows_count_as_nd_exit_like_issue_sheet():
    """已是 ND、現有 Safety 0 的列在「店舖問題 SKU」列為 ND 退場，⑤⑥⑦ 同樣計入。"""
    out = _run([{"Buyer": "B8", "Brand": "BrandY", "SKU": "7", "Site Code": "S08", "RP Type": "ND", "Adj. Safety Stock": 0}])
    issues = _build_combined_shop_issues(out)
    assert issues["問題類型"].tolist() == [_ND]
    imp = _build_combined_impact_summary(out, issues=issues)
    for label in ("Brand", "Buyer", "店舖"):
        assert _impact_row(imp, label)[1:6].tolist() == [1, 1, 0, 0, 0]
    assert _vals(_build_combined_issue_pivot(issues, "shop").iloc[0]) == [1, 0, 0, 0, 0, 0, 1, 0]


def test_new_sku_mark_is_not_an_issue():
    """New SKU 標記不屬任何問題類型：不變的新品列 → 無需行動，樞紐只有總計 0。"""
    out = _run([{
        "Shop Stock Status": "Overstock (P90D 0 Sales)", "Adj. Safety Stock": 3,
        "Shop Monthly Sold Qty (P90D)": 0, "Sold Qty (CM - MTD)": 0.0, "Launch Date": _launch_date(10),
    }])
    assert out[_ND_COL].tolist() == ["New SKU (3 Month)"]
    issues = _build_combined_shop_issues(out)
    assert issues.empty
    imp = _build_combined_impact_summary(out, issues=issues)
    assert _impact_row(imp, "店舖")[1:6].tolist() == [1, 0, 0, 0, 1]


def test_impact_summary_blank_keys_not_counted_and_whitespace_stripped():
    raw = _run([
        {"Buyer": " B1 ", "Brand": "BrandA", "SKU": "1", "Site Code": "S01", "Adj. Safety Stock": 5},
        {"Buyer": "B1", "Brand": None, "SKU": "2", "Site Code": None, "Adj. Safety Stock": 5},
        {"Buyer": "B2", "Brand": " BrandA ", "SKU": "3", "Site Code": " S02 ", "Adj. Safety Stock": 3},
    ])
    imp = _build_combined_impact_summary(raw, issues=_build_combined_shop_issues(raw))
    # 一列缺 Brand、一列缺 Site Code → 不計入總數；" B1 " 與 " BrandA " 去空白後即同一對象
    assert _impact_row(imp, "Brand")[1:6].tolist() == [1, 0, 1, 0, 0]
    assert _impact_row(imp, "Buyer")[1:6].tolist() == [2, 0, 2, 0, 0]
    assert _impact_row(imp, "店舖")[1:6].tolist() == [2, 0, 2, 0, 0]


def test_impact_summary_empty_and_missing_columns():
    empty = _run(_designed_rows()).iloc[0:0]
    imp = _build_combined_impact_summary(empty, issues=_build_combined_shop_issues(empty))
    assert imp["統計對象"].tolist() == ["Brand", "Buyer", "店舖"]
    assert (imp[["總數", _ND, _SHORT, _HIGH, "無需行動"]] == 0).all().all()

    raw = _run([{"SKU": "1", "Adj. Safety Stock": 5}]).drop(columns=["Brand"])
    imp = _build_combined_impact_summary(raw, issues=_build_combined_shop_issues(raw))
    assert _impact_row(imp, "Brand")["總數"] == 0
    assert _impact_row(imp, "Buyer")[1:6].tolist() == [1, 0, 1, 0, 0]


# ── ⑥⑦ 問題類型樞紐 ──

def test_pivot_buyer_brand_subtotals_and_order():
    tab = _build_combined_issue_pivot(_build_combined_shop_issues(_run(_designed_rows())), "buyer_brand")

    assert list(tab.columns) == ["Buyer / Brand", _SSR_PIVOT_LEVEL_COL, *_MEASURE_COLS]
    # Buyer 依名稱排序，其下 Brand 依名稱排序；沒有需行動列的 B3 不列出；末列為總計
    assert tab[["Buyer / Brand", _SSR_PIVOT_LEVEL_COL]].values.tolist() == [
        ["B1", "Buyer"], ["BrandA", "Brand"],
        ["B2", "Buyer"], ["BrandA", "Brand"], ["BrandB", "Brand"],
        ["總計", "總計"],
    ]
    rows = tab.to_dict("records")
    # Buyer 列＝其下 Brand 列之和；數值序：ND(行數,加總)／不足／過高／總計
    assert _vals(rows[0]) == [0, 0, 1, 2, 1, -3, 2, -1]
    assert _vals(rows[1]) == [0, 0, 1, 2, 1, -3, 2, -1]
    assert _vals(rows[2]) == [1, 0, 1, 4, 1, -4, 3, 0]
    assert _vals(rows[3]) == [0, 0, 1, 4, 0, 0, 1, 4]
    assert _vals(rows[4]) == [1, 0, 0, 0, 1, -4, 2, -4]
    assert _vals(rows[5]) == [1, 0, 2, 6, 2, -7, 5, -1]
    assert all(type(v) is int for rec in rows for v in _vals(rec))


def test_pivot_shop_sorted_by_site_code_with_shop_name():
    rows = _designed_rows()
    for r in rows:
        r["Item_Shop"] = f"Some Product, N/A, PC{r['Site Code']}-{r['Site Code']}店"
    tab = _build_combined_issue_pivot(_build_combined_shop_issues(_run(rows)), "shop")

    assert list(tab.columns) == ["Site Code", "Shop Name", _SSR_PIVOT_LEVEL_COL, *_MEASURE_COLS]
    # 依 Site Code 排序；S04 沒有需行動列不列出
    assert tab[["Site Code", "Shop Name", _SSR_PIVOT_LEVEL_COL]].values.tolist() == [
        ["S01", "S01店", "店舖"], ["S02", "S02店", "店舖"], ["S03", "S03店", "店舖"], ["總計", "", "總計"],
    ]
    recs = tab.to_dict("records")
    # S01：SKU 1 +2、SKU 3 +4（設定不足）與 SKU 2 −3（設定過高）；S02：SKU 4 −4；S03：SKU 5 ND 退場
    assert _vals(recs[0]) == [0, 0, 2, 6, 1, -3, 3, 3]
    assert _vals(recs[1]) == [0, 0, 0, 0, 1, -4, 1, -4]
    assert _vals(recs[2]) == [1, 0, 0, 0, 0, 0, 1, 0]
    assert _vals(recs[3]) == [1, 0, 2, 6, 2, -7, 5, -1]


def test_pivot_blank_labels_and_missing_key_columns():
    """缺 Buyer／Brand 的列以「（空白）」列示，小計仍等於總計；缺欄時同樣不拋錯。"""
    raw = _run([
        {"Buyer": "B1", "Brand": "BrandA", "SKU": "1", "Site Code": "S01", "Adj. Safety Stock": 5},
        {"Buyer": "B1", "Brand": "", "SKU": "2", "Site Code": "S01", "Adj. Safety Stock": 5},
        {"Buyer": None, "Brand": None, "SKU": "3", "Site Code": "S01", "Adj. Safety Stock": 3},
    ])
    issues = _build_combined_shop_issues(raw)
    tab = _build_combined_issue_pivot(issues, "buyer_brand")
    # 名稱排序：「B」在「（」之前
    assert tab[["Buyer / Brand", _SSR_PIVOT_LEVEL_COL]].values.tolist() == [
        ["B1", "Buyer"], ["BrandA", "Brand"], [_SSR_PIVOT_BLANK, "Brand"],
        [_SSR_PIVOT_BLANK, "Buyer"], [_SSR_PIVOT_BLANK, "Brand"],
        ["總計", "總計"],
    ]
    recs = tab.to_dict("records")
    # B1：BrandA +2 與（空白）+2；全空白 Buyer：+4
    assert _vals(recs[0]) == [0, 0, 2, 4, 0, 0, 2, 4]
    assert _vals(recs[1]) == [0, 0, 1, 2, 0, 0, 1, 2]
    assert _vals(recs[2]) == [0, 0, 1, 2, 0, 0, 1, 2]
    assert _vals(recs[3]) == [0, 0, 1, 4, 0, 0, 1, 4]
    assert _vals(recs[4]) == [0, 0, 1, 4, 0, 0, 1, 4]
    assert _vals(recs[5]) == [0, 0, 3, 8, 0, 0, 3, 8]

    no_cols = issues.drop(columns=["Shop Name", "Brand"])
    assert set(_build_combined_issue_pivot(no_cols, "shop")["Shop Name"]) == {_SSR_PIVOT_BLANK, ""}
    assert _SSR_PIVOT_BLANK in _build_combined_issue_pivot(no_cols, "buyer_brand")["Buyer / Brand"].tolist()


def test_pivot_without_issues_is_one_zero_total_row():
    """沒有需行動列（或未提供）時只有一行總計 0，欄位結構不變。"""
    issues = _build_combined_shop_issues(_run([{"SKU": "1", "Adj. Safety Stock": 7}]))  # 不變
    assert issues.empty
    for src in (None, issues):
        tab = _build_combined_issue_pivot(src, "buyer_brand")
        assert len(tab) == 1 and tab.iloc[0, :2].tolist() == ["總計", "總計"] and _vals(tab.iloc[0]) == [0] * 8
        tab = _build_combined_issue_pivot(src, "shop")
        assert len(tab) == 1 and tab.iloc[0, :3].tolist() == ["總計", "", "總計"] and _vals(tab.iloc[0]) == [0] * 8


def test_pivot_ignores_issue_frame_index():
    """issues 帶字串或重複 index 時，統計仍按列位置計算。"""
    issues = _build_combined_shop_issues(_run(_designed_rows()))
    base = _build_combined_issue_pivot(issues, "buyer_brand")
    shuffled = issues.sample(frac=1.0, random_state=3)
    shuffled.index = [f"r{i}" for i in range(len(shuffled))]
    pd.testing.assert_frame_equal(_build_combined_issue_pivot(shuffled, "buyer_brand"), base)
    twice = _build_combined_issue_pivot(pd.concat([shuffled, shuffled]), "shop")
    assert _vals(twice.iloc[-1]) == [2 * v for v in _vals(base.iloc[-1])]


# ── 版面 ──

def _thin():
    return Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))


def _grid(ws):
    return [[ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)] for r in range(1, ws.max_row + 1)]


def _merged(ws):
    return sorted(str(m) for m in ws.merged_cells.ranges)


def _fill(ws, row, cols):
    return {ws.cell(row=row, column=c).fill.fgColor.rgb for c in cols}


def test_write_pivot_block_buyer_brand_layout():
    tab = _build_combined_issue_pivot(_build_combined_shop_issues(_run(_designed_rows())), "buyer_brand")
    wb = Workbook()
    ws = wb.active
    nxt = _write_issue_pivot_block(ws, 5, "⑥ 標題", tab, "buyer_brand", _thin())
    g = _grid(ws)

    assert g[4][0] == "⑥ 標題" and "A5:I5" in _merged(ws)
    # 兩層表頭：問題類型名稱在其兩欄之首（跨越選取範圍置中），下一列為 行數／加總
    assert g[5][:9] == ["Buyer / Brand", _ND, None, _SHORT, None, _HIGH, None, _SSR_PIVOT_TOTAL, None]
    assert g[6][:9] == [None, *_MEASURES * 4]  # A7 屬合併儲存格 A6:A7，讀值為 None
    assert "A6:A7" in _merged(ws)  # 標籤表頭跨兩列
    assert all(ws.cell(row=6, column=c).alignment.horizontal == "centerContinuous" for c in range(2, 10))
    assert all(ws.cell(row=7, column=c).font.bold for c in range(2, 10))
    assert ws.cell(row=6, column=1).font.bold  # 合併格 A6:A7 以左上格格式為準
    # 資料列：沒有列的問題類型留白；總計一律顯示
    assert [r[0] for r in g[7:13]] == ["B1", "BrandA", "B2", "BrandA", "BrandB", "總計"]
    assert [r[1:9] for r in g[7:13]] == [
        [None, None, 1, 2, 1, -3, 2, -1],
        [None, None, 1, 2, 1, -3, 2, -1],
        [1, 0, 1, 4, 1, -4, 3, 0],
        [None, None, 1, 4, None, None, 1, 4],
        [1, 0, None, None, 1, -4, 2, -4],
        [1, 0, 2, 6, 2, -7, 5, -1],
    ]
    # Buyer 小計列：全列黃底粗體；Brand 列：縮排、列入摺疊群組（小計列在群組上方）
    assert _fill(ws, 8, range(1, 10)) == {"00FFFF99"} and all(ws.cell(row=8, column=c).font.bold for c in range(1, 10))
    assert _fill(ws, 10, range(1, 10)) == {"00FFFF99"}
    assert ws.cell(row=9, column=1).alignment.indent == 1 and ws.row_dimensions[9].outlineLevel == 1
    assert [ws.row_dimensions[r].outlineLevel for r in range(8, 14)] == [0, 1, 0, 1, 1, 0]
    assert ws.sheet_properties.outlinePr.summaryBelow is False
    # 總計欄淺黃底（Buyer 小計列整列黃底，涵蓋總計欄）；總計列淺藍底粗體
    assert _fill(ws, 9, (8, 9)) == {"00FFF2CC"} and _fill(ws, 9, (1, 5)) == {"00000000"}
    assert _fill(ws, 13, range(1, 10)) == {"00DDEBF7"}
    assert all(ws.cell(row=13, column=c).font.bold for c in range(1, 10))
    # 數值格右靠並套用千分位
    cell = ws.cell(row=8, column=4)
    assert cell.alignment.horizontal == "right" and cell.number_format == "#,##0"
    assert ws.auto_filter.ref is None  # Buyer／Brand 表不設自動篩選（每張工作表只容許一個）
    assert nxt == 14  # 末列 13 的下一列（呼叫端再加一列空白）


def test_write_pivot_block_shop_layout_with_autofilter():
    tab = _build_combined_issue_pivot(_build_combined_shop_issues(_run(_designed_rows())), "shop")
    wb = Workbook()
    ws = wb.active
    nxt = _write_issue_pivot_block(ws, 3, "⑦ 標題", tab, "shop", _thin())
    g = _grid(ws)
    assert g[2][:10] == ["⑦ 標題", None, None, None, None, None, None, None, None, None]
    assert g[3][:10] == [None, None, _ND, None, _SHORT, None, _HIGH, None, _SSR_PIVOT_TOTAL, None]
    assert g[4][:10] == ["Site Code", "Shop Name", *_MEASURES * 4]
    assert [r[0] for r in g[5:9]] == ["S01", "S02", "S03", "總計"]
    assert _merged(ws) == ["A3:J3"]  # 表頭不合併（設有自動篩選）
    # 自動篩選由表頭（第 5 列）至最後一間店舖（第 8 列）；總計列（第 9 列）不在內
    assert ws.auto_filter.ref == "A5:J8"
    assert all(ws.row_dimensions[r].outlineLevel == 0 for r in range(1, 10))
    assert all(ws.cell(row=6, column=c).alignment.horizontal == "right" for c in range(3, 11))
    assert _fill(ws, 6, (9, 10)) == {"00FFF2CC"} and _fill(ws, 8, (9, 10)) == {"00FFF2CC"}
    assert _fill(ws, 9, range(1, 11)) == {"00DDEBF7"}
    assert nxt == 10


def test_write_pivot_block_without_issues_has_no_autofilter():
    tab = _build_combined_issue_pivot(None, "shop")
    wb = Workbook()
    ws = wb.active
    _write_issue_pivot_block(ws, 1, "⑦ 標題", tab, "shop", _thin())
    assert ws.auto_filter.ref is None
    assert [ws.cell(row=4, column=c).value for c in range(1, 11)] == ["總計", None, None, None, None, None, None, None, 0, 0]


def test_summary_sheet_writes_blocks_five_to_seven_with_capped_widths():
    out = _run(_designed_rows())
    issues = _build_combined_shop_issues(out)
    wb = Workbook()
    ws = wb.active
    _write_summary_sheet(
        ws,
        pd.DataFrame({"分類": ["增加", "總計"], "行數": [1, 2], "說明": ["短", "很長的說明文字" * 8]}),
        pd.DataFrame({"建議分類": ["總計"], "行數": [2], "說明": ["另一段很長的說明文字" * 8]}),
        _thin(),
        impact_summary=_build_combined_impact_summary(out, issues=issues),
        buyer_brand_pivot=_build_combined_issue_pivot(issues, "buyer_brand"),
        shop_pivot=_build_combined_issue_pivot(issues, "shop"),
    )
    g = _grid(ws)
    col_a = [r[0] for r in g]
    assert any(str(t).startswith("⑤ 影響範圍統計（Brand／Buyer／店舖）") for t in col_a)
    assert any(str(t).startswith("⑥ 按 Buyer／Brand 統計") for t in col_a)
    assert any(str(t).startswith("⑦ 按店舖統計（Site Code／Shop Name）") for t in col_a)
    i5 = next(i for i, t in enumerate(col_a) if str(t).startswith("⑤"))
    i6 = next(i for i, t in enumerate(col_a) if str(t).startswith("⑥"))
    i7 = next(i for i, t in enumerate(col_a) if str(t).startswith("⑦"))
    assert i5 < i6 < i7
    h5 = next(r for r in range(1, ws.max_row + 1)
              if ws.cell(row=r, column=1).value == "統計對象" and ws.cell(row=r, column=2).value == "總數")
    assert [ws.cell(row=h5, column=c).value for c in range(1, 8)] == ["統計對象", "總數", _ND, _SHORT, _HIGH, "無需行動", "說明"]
    assert [ws.cell(row=h5 + 3, column=c).value for c in range(2, 7)] == [4, 1, 1, 2, 1]
    # 各區塊之間相隔一列空白
    assert g[i6 - 1][0] is None and g[i7 - 1][0] is None
    # 自動篩選設在店舖表的表頭列
    head_row = int(ws.auto_filter.ref.split(":")[0][1:])
    assert ws.cell(row=head_row, column=1).value == "Site Code"
    # 欄寬上限：A 欄 ≤ 36、其餘 ≤ 30；被收窄的長文字格已設自動換行
    widths = {get_column_letter(c): ws.column_dimensions[get_column_letter(c)].width for c in range(1, ws.max_column + 1)}
    assert all(v is not None and v <= 30 for k, v in widths.items() if k != "A")
    assert widths["A"] <= 36
    long_cells = [c for row in ws.iter_rows() for c in row if isinstance(c.value, str) and len(c.value) > 30 and c.column > 1]
    assert long_cells and all(c.alignment.wrap_text for c in long_cells)
    # 被收窄的長文字格：列高按估算行數加大，文字得以完整顯示
    tallest = max(long_cells, key=lambda c: len(c.value))
    assert ws.row_dimensions[tallest.row].height >= 2 * 15 + 3


def test_summary_sheet_without_new_blocks_still_works():
    """未提供新表時只寫 ①②（舊呼叫方式相容）。"""
    wb = Workbook()
    ws = wb.active
    _write_summary_sheet(
        ws, pd.DataFrame({"分類": ["總計"], "行數": [1]}),
        pd.DataFrame({"建議分類": ["總計"], "行數": [1]}),
        None,
    )
    col_a = [ws.cell(row=i, column=1).value for i in range(1, ws.max_row + 1)]
    assert not any(str(t)[:1] in ("⑤", "⑥", "⑦") for t in col_a)
    assert ws.auto_filter.ref is None


def test_excel_summary_pivots_reconcile_with_shop_issue_sheet():
    """輸出 Excel：SUMMARY ⑥⑦ 的總計等於「店舖問題 SKU」的列數與 Safety Qty Change 合計；說明頁導讀同步。"""
    rows = _designed_rows()
    for r in rows:
        r["Item_Shop"] = f"Some Product, N/A, PC{r['Site Code']}-{r['Site Code']}店"
    s, raw, cats, _ = transform_combined(
        _df(rows), rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="allow_both",
    )
    buf = generate_excel_combined(s, raw, cats)
    wb = load_workbook(BytesIO(buf.getvalue()))
    assert wb.sheetnames == ["說明", "RAW", "Shop Stock Status", "By Shop", "店舖問題 SKU", "SUMMARY", "Safety Stock 數據分析摘要"]

    issues_ws = wb["店舖問題 SKU"]
    header = [c.value for c in issues_ws[3]]
    body = [r for r in issues_ws.iter_rows(min_row=4, values_only=True) if r[0] is not None]
    types = [r[header.index("問題類型")] for r in body]
    chg = [r[header.index("Safety Qty Change")] for r in body]
    assert types.count(_ND) == 1 and types.count(_SHORT) == 2 and types.count(_HIGH) == 2
    assert len(body) == 5 and sum(chg) == -1

    ws = wb["SUMMARY"]
    g = _grid(ws)
    # ⑥：表頭下一列起第六列為總計（A 欄表頭為合併儲存格，讀值為 None）
    j = next(idx for idx, r in enumerate(g) if r[:3] == [None, "行數", "加總 - Safety Qty Change"]) + 1
    assert g[j + 5][1:9] == [1, 0, 2, 6, 2, -7, 5, -1]
    k = next(idx for idx, r in enumerate(g) if r[:3] == ["Site Code", "Shop Name", "行數"]) + 1
    assert [r[:2] for r in g[k:k + 4]] == [["S01", "S01店"], ["S02", "S02店"], ["S03", "S03店"], ["總計", None]]
    assert [r[2:10] for r in g[k:k + 4]] == [
        [None, None, 2, 6, 1, -3, 3, 3],
        [None, None, None, None, 1, -4, 1, -4],
        [1, 0, None, None, None, None, 1, 0],
        [1, 0, 2, 6, 2, -7, 5, -1],
    ]
    # k 同時是店舖表表頭的列號；自動篩選範圍＝表頭至最後一間店舖（總計列的前一列）
    assert ws.auto_filter.ref == f"A{k}:J{k + 3}"

    guide = wb["說明"]
    guide_texts = [str(c.value) for row in guide.iter_rows() for c in row if c.value is not None]
    assert any("⑤ 影響範圍統計" in t and "Buyer＞Brand" in t and "店舖樞紐分析" in t for t in guide_texts)
