# -*- coding: utf-8 -*-
"""Safety Stock Review「指定總量分配」（Target Safety）測試。

覆蓋：輸入正規化、各半權重、最大餘數法取整、只覆寫指定 SKU、ND 店排除、
低於現有 Safety 的處理、ND 建議清理、提示欄、SUMMARY ④ 對帳與端到端 warning。
"""

import numpy as np
import pandas as pd
from openpyxl import Workbook

import transformer  # noqa: F401  先載入墊片，避免 modes 之間循環引用
from transformer import COMBINED_RAW_COLS, _add_combined_recommendation_columns, transform_combined
from modes.safety_review import (
    _SSR_TARGET_BASIS_LABEL,
    _SSR_TARGET_NOTE_COLUMN,
    _SSR_TARGET_SHARE_COLUMN,
    _SSR_TARGET_TOTAL_COLUMN,
    _build_combined_target_summary,
    _ssr_largest_remainder,
    _ssr_normalize_target_map,
    _ssr_target_weight_series,
    _write_summary_sheet,
)

_RATES_O = {"default": 0.5, "tourist": 0.3, "mix": 0.4}
_RATES_P = {"default": 0.5, "tourist": 0.5, "mix": 0.5}


def _base_row(**kw):
    row = {
        "Buyer": "B1",
        "Brand": "BrandA",
        "SKU": "1001",
        "Site Code": "S01",
        "Item_Shop": "Some Product-S01",
        "Shop Stock Status": "Normal",
        "Adj. Safety Stock": 100,
        "Shop Monthly Sold Qty (P90D)": 200,
        "Sold Qty (P90D)": 600,
        "Sold Qty (CM - MTD)": 100.0,
        "Shop Nature": "Tourist",
        "RP Type": "RF",
        "Shop Stock Qty (CM)": 50,
        "Supply Source": "Warehouse",
        "OM": "OM1",
        "Adj. MCH1": "MchA",
        "Adj. SST": "SstA",
    }
    row.update(kw)
    return row


def _df(rows):
    return pd.DataFrame([_base_row(**r) for r in rows])


def _three_shops(**kw):
    """三間店：銷量 600／1200／1800、現有 Safety 各 100 → 權重 25%／33.33%／41.67%。"""
    return _df([
        {"Site Code": "S01", "Item_Shop": "Some Product-S01",
         "Sold Qty (P90D)": 600, "Shop Monthly Sold Qty (P90D)": 200},
        {"Site Code": "S02", "Item_Shop": "Some Product-S02",
         "Sold Qty (P90D)": 1200, "Shop Monthly Sold Qty (P90D)": 400},
        {"Site Code": "S03", "Item_Shop": "Some Product-S03",
         "Sold Qty (P90D)": 1800, "Shop Monthly Sold Qty (P90D)": 600},
    ])


def _run(df, target_safety, **kw):
    params = dict(
        rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="increase_only",
        target_safety=target_safety,
    )
    params.update(kw)
    return _add_combined_recommendation_columns(df, **params)


# ── 基本元件 ──

def test_largest_remainder_sums_exactly():
    assert _ssr_largest_remainder(3000, [0.25, 0.3333333, 0.4166667]).tolist() == [750, 1000, 1250]
    # 除不盡時餘量按小數部分大者先補，總和恆等於指定值
    alloc = _ssr_largest_remainder(10, [1 / 3, 1 / 3, 1 / 3])
    assert alloc.tolist() == [4, 3, 3]
    assert int(alloc.sum()) == 10
    # 權重全 0 → 均分；指定 0 或負數 → 全 0
    assert _ssr_largest_remainder(9, [0, 0, 0]).tolist() == [3, 3, 3]
    assert _ssr_largest_remainder(0, [1, 1]).tolist() == [0, 0]
    assert _ssr_largest_remainder(-5, [1, 1]).tolist() == [0, 0]
    assert _ssr_largest_remainder(5, []).tolist() == []


def test_target_weight_is_half_sold_half_safety():
    w = _ssr_target_weight_series([600, 1200, 1800], [100, 100, 100])
    assert np.allclose(w, [0.25, 1 / 3, 5 / 12], atol=1e-6)
    # 銷量全 0 → 全按現有 Safety 佔比
    assert np.allclose(_ssr_target_weight_series([0, 0], [300, 100]), [0.75, 0.25])
    # 兩半皆 0 → 均分
    assert np.allclose(_ssr_target_weight_series([0, 0, 0], [0, 0, 0]), [1 / 3, 1 / 3, 1 / 3])
    # 負數與缺值按 0 計
    assert np.allclose(_ssr_target_weight_series([-5, 0], [10, 10]), [0.5, 0.5])


def test_normalize_target_map():
    out = _ssr_normalize_target_map({
        "1001": "3000",      # 文字數字
        "1002.0": 0,         # 0 → 不啟用
        " 1003 ": 12.7,      # 前後空白 + 小數 → 無條件捨去
        "": 500,             # 空 SKU → 略過
        "1004": "abc",       # 非數值 → 略過
        "1005": -3,          # 負數 → 略過
    })
    assert out == {"1001": 3000, "1003": 12}
    assert _ssr_normalize_target_map(None) == {}
    assert _ssr_normalize_target_map(pd.DataFrame()) == {}
    # DataFrame／tuple 清單亦可
    df = pd.DataFrame({"SKU": ["1", "2"], "Qty": [10, 20]})
    assert _ssr_normalize_target_map(df) == {"1": 10, "2": 20}
    assert _ssr_normalize_target_map([("3", 30)]) == {"3": 30}


# ── 主流程 ──

def test_target_allocates_exactly_and_only_listed_sku():
    df = _three_shops()
    df = pd.concat([df, _df([{
        "SKU": "2002", "Site Code": "S09", "Item_Shop": "Other-S09",
    }])], ignore_index=True)
    out = _run(df, {"1001": 3000})

    target_rows = out[out["SKU"] == "1001"]
    assert target_rows["Recommend Safety Stock"].tolist() == [750, 1000, 1250]
    assert int(target_rows["Recommend Safety Stock"].sum()) == 3000
    assert target_rows[_SSR_TARGET_TOTAL_COLUMN].tolist() == [3000, 3000, 3000]
    assert target_rows[_SSR_TARGET_SHARE_COLUMN].round(2).tolist() == [25.0, 33.33, 41.67]

    # 未指定的 SKU 完全不受影響（按原公式 Z ≤ X → max(Y, Z)）
    other = out[out["SKU"] == "2002"].iloc[0]
    assert other["Recommend Safety Stock"] == 100
    assert pd.isna(other[_SSR_TARGET_TOTAL_COLUMN])


def test_target_excludes_nd_shops():
    df = _three_shops()
    df.loc[2, "RP Type"] = "ND"
    df.loc[2, "Adj. Safety Stock"] = 0
    out = _run(df, {"1001": 3000})

    nd_row = out[out["Site Code"] == "S03"].iloc[0]
    assert pd.isna(nd_row[_SSR_TARGET_TOTAL_COLUMN])
    assert nd_row["Recommend Safety Stock"] == 0

    rf_rows = out[out["RP Type"] == "RF"]
    assert int(rf_rows["Recommend Safety Stock"].sum()) == 3000


def test_target_missing_sku_is_reported_not_allocated():
    out = _run(_three_shops(), {"9999": 3000})
    assert out["_safety_floor3"].sum() == 0
    info = out.attrs.get("ssr_target_info") or {}
    assert info.get("matched") == []
    assert info.get("missing") == ["9999"]
    assert _SSR_TARGET_TOTAL_COLUMN not in out.columns


def test_target_below_current_safety_shows_decrease():
    df = _three_shops()
    out = _run(df, {"1001": 90})
    # 90 件按 25%／33.33%／41.67% 分配 → 23／30／37（餘數法），全部低於現有 100
    # → 「只可增加」模式下仍以指定值為準，出現「減少」
    assert out["Recommend Safety Stock"].tolist() == [23, 30, 37]
    assert int(out["Recommend Safety Stock"].sum()) == 90
    assert out["Safety Qty Change"].tolist() == [-77, -70, -63]
    assert out["Safety Qty Change 分類"].tolist() == ["減少", "減少", "減少"]
    assert (out.attrs["ssr_target_info"]["below_current"]) == 3


def test_target_note_flags_over_cap_and_under_floor():
    df = _three_shops()
    # X = min(floor(月銷×rate), Z)；Y = ceil(月銷預估×rate)。S03 月銷 600、Z=100
    # → X=100、Y=200；分配 1250 件必然超出上限
    out = _run(df, {"1001": 3000})
    note = out[_SSR_TARGET_NOTE_COLUMN].tolist()
    assert all("超出庫存上限" in n for n in note)
    assert "低於缺貨底線" not in note[0]

    # 指定值偏低時反過來標示低於缺貨底線
    low = _run(_three_shops(), {"1001": 9})
    assert all("低於缺貨底線" in n for n in low[_SSR_TARGET_NOTE_COLUMN].tolist())


def test_target_clears_nd_recommendation():
    df = _df([{
        "Site Code": "S01", "Item_Shop": "Some Product-S01",
        "Shop Stock Status": "Overstock (P90D 0 Sales)",
        "Adj. Safety Stock": 0, "Sold Qty (P90D)": 0,
        "Shop Monthly Sold Qty (P90D)": 0,
    }])
    out = _run(df, {"1001": 500}, safety_mode="allow_both")
    # 指定補貨後不應同時建議轉 ND
    assert out["Recommend Safety Stock"].tolist() == [500]
    assert out["Recommendation of new Safety Qty / ND"].tolist() == [""]


def test_target_summary_block():
    out = _run(_three_shops(), {"1001": 3000})
    summary = _build_combined_target_summary(out)
    assert list(summary.columns) == [
        "SKU", "Target 全線合計", "已分配", "差額", "參與店舖數",
        "超出庫存上限店數", "低於缺貨底線店數",
    ]
    assert summary.iloc[0]["已分配"] == 3000
    assert summary.iloc[0]["差額"] == 0
    assert summary.iloc[0]["參與店舖數"] == 3
    assert summary.iloc[0]["超出庫存上限店數"] == 3
    assert summary.iloc[-1]["SKU"] == "總計"
    assert summary.iloc[-1]["Target 全線合計"] == 3000

    # 未啟用 → 空表
    assert _build_combined_target_summary(_run(_three_shops(), None)).empty


def test_summary_sheet_writes_block_four():
    out = _run(_three_shops(), {"1001": 3000})
    target_summary = _build_combined_target_summary(out)
    wb = Workbook()
    ws = wb.active
    _write_summary_sheet(
        ws, pd.DataFrame({"分類": ["總計"], "行數": [3]}),
        pd.DataFrame({"建議分類": ["總計"], "行數": [3]}),
        None, target_summary=target_summary,
    )
    texts = [str(c.value) for row in ws.iter_rows() for c in row if c.value is not None]
    assert any("④ 指定總量分配" in t for t in texts)
    assert any(_SSR_TARGET_BASIS_LABEL in t for t in texts)


def test_transform_combined_end_to_end():
    df = _three_shops()
    _, raw, _, warnings = transform_combined(
        df, rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="increase_only",
        target_safety={"1001": 3000},
    )
    assert int(raw.loc[raw["SKU"] == "1001", "Recommend Safety Stock"].sum()) == 3000
    assert any("指定總量分配" in str(w) for w in warnings)
    assert any(_SSR_TARGET_BASIS_LABEL in str(w) for w in warnings)
    # 三個新欄位已列入 RAW 契約欄序
    for col in (_SSR_TARGET_TOTAL_COLUMN, _SSR_TARGET_SHARE_COLUMN, _SSR_TARGET_NOTE_COLUMN):
        assert col in COMBINED_RAW_COLS


def test_transform_combined_warns_when_all_shops_nd():
    df = _three_shops()
    df["RP Type"] = "ND"
    df["Adj. Safety Stock"] = 0
    _, raw, _, warnings = transform_combined(
        df, rates_overstock=_RATES_O, rates_oos=_RATES_P,
        mtd_days=15, month_total_days=30, safety_mode="increase_only",
        target_safety={"1001": 3000},
    )
    assert _SSR_TARGET_TOTAL_COLUMN not in raw.columns
    assert any("沒有可分配的非 ND 店舖" in str(w) for w in warnings)


# ── 上載 Target 清單 ──

def _target_csv(rows, header=None):
    import csv
    from io import StringIO

    header = header or ["Brand", "SKU", "SKU Desc", "Vendor MC", "Buyer", "現時Safety", "我想輸入的Safety Qty"]
    buf = StringIO()
    writer = csv.writer(buf)
    writer.writerow(header)
    for r in rows:
        writer.writerow(list(r))
    return buf.getvalue().encode("utf-8")


def test_column_key_normalizes():
    from modes.safety_review import _ssr_column_key
    assert _ssr_column_key("現時 Safety") == _ssr_column_key("現時Safety")
    assert _ssr_column_key("Safety Stock (現時)") == _ssr_column_key("safety stock（現時）")
    assert _ssr_column_key("  SKU  ") == "sku"


def test_read_target_file_csv_aliases_and_cleaning():
    from modes.safety_review import _ssr_read_target_file

    data = _target_csv([
        ("BSK", "1001", "Hydra Serum", "VC1", "Cara", "800", "3,000"),
        ("BSK", "1002", "Repair Cream", "VC1", "Cara", "450", "200"),
        ("BSK", "1002", "Repair Cream", "VC1", "Cara", "450", "250"),   # 重複 SKU → 取最後
        ("BSK", "1003", "Cleansing Oil", "VC2", "Cara", "0", "0"),      # 目標 0 → 略過
        ("", "總計", "", "", "", "1250", "3250"),                        # 合計列 → 略過
    ])
    df, warnings = _ssr_read_target_file(data, "target.csv")

    assert list(df.columns) == [
        "Brand", "SKU", "SKU Desc", "Vendor MC", "Buyer", "現時Safety", "Target Safety (全線合計)",
    ]
    assert df["SKU"].tolist() == ["1001", "1002"]
    assert df["Target Safety (全線合計)"].tolist() == [3000, 250]   # 千分位 + 重複取最後
    assert df["現時Safety"].tolist() == [800, 450]
    assert any("合計列" in w for w in warnings)
    assert any("重複" in w for w in warnings)
    assert any("≤ 0" in w for w in warnings)


def test_read_target_file_xlsx_picks_sheet_with_required_columns():
    from openpyxl import Workbook
    from io import BytesIO
    from modes.safety_review import _ssr_read_target_file

    wb = Workbook()
    cover = wb.active
    cover.title = "封面"
    cover.append(["Target Safety 清單"])
    cover.append(["說明：每隻 SKU 一列"])
    data_ws = wb.create_sheet("Target")
    data_ws.append(["SKU", "我想輸入的Safety Qty"])
    data_ws.append(["2001", 1200])
    data_ws.append(["2002", 800])
    buf = BytesIO()
    wb.save(buf)

    df, warnings = _ssr_read_target_file(buf.getvalue(), "target.xlsx")
    assert df["SKU"].tolist() == ["2001", "2002"]
    assert df["Target Safety (全線合計)"].tolist() == [1200, 800]
    assert warnings == []


def test_read_target_file_missing_columns_reports_identified():
    from modes.safety_review import _ssr_read_target_file
    data = _target_csv([("BSK", "1001", "x", "VC1", "Cara", "800", "")],
                       header=["Brand", "貨品編號", "描述", "Vendor", "Buyer", "現時 Safety", ""])
    try:
        _ssr_read_target_file(data, "bad.csv")
    except ValueError as exc:
        msg = str(exc)
        assert "找不到必要欄位" in msg
        assert "SKU" in msg
        assert "貨品編號" in msg or "已識別" in msg
    else:
        raise AssertionError("應拋出 ValueError")


def test_target_from_file_drives_allocation():
    from modes.safety_review import _ssr_read_target_file

    data = _target_csv([
        ("BSK", "1001", "Hydra Serum", "VC1", "Cara", "400", "3,000"),
    ])
    tdf, _ = _ssr_read_target_file(data, "target.csv")
    target_map = {str(r["SKU"]): int(r["Target Safety (全線合計)"]) for _, r in tdf.iterrows()}

    out = _run(_three_shops(), target_map)
    assert out["Recommend Safety Stock"].tolist() == [750, 1000, 1250]
    assert int(out["Recommend Safety Stock"].sum()) == 3000
    assert out[_SSR_TARGET_TOTAL_COLUMN].tolist() == [3000, 3000, 3000]
