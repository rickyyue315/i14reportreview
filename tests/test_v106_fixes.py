# -*- coding: utf-8 -*-
"""v1.106.0 邏輯修復回歸測試。

涵蓋：
- OOS Review：新品保底 3 件（漏 import 修复）、舊品零銷量不保底、
  Shop Nature 帶空格仍按正確倍率、Projected Monthly 負值下限 0。
- Safety Stock Review：Overstock Cap (X) 保底與 Y 同一準則（舊品零銷量 X=0）、
  Shop Nature 帶空格套正確天數。
- Stock Impact（簡化）：New 表三期 SOH Tier 區塊欄名帶期別（修 merge 碰撞）、
  D001／店舖 M-2／M-1 以 LLM／LM 代理。
- Forecast：D001「首個非空值」（首列空白不再取 0）。
- 店舖精簡：SKU 空白列不計入 Top N% 母體；金額口徑不套件數長尾保護。
- NDRF：缺 Adj. MCH1 不崩潰；Month 數值型可比對；ND 負庫存不計入 RF_Total。
- hybrid：_write_metric_cards 支援第 8 排指標卡。
- OOS／Overstock：儀表板影響面只計有異常的 Buyer／Brand／SKU。
"""
import warnings as _warnings

_warnings.filterwarnings("ignore")
from io import BytesIO
from datetime import date, timedelta

import numpy as np
import pandas as pd
import pytest
from openpyxl import load_workbook

from transformer import (
    _add_oos_recommendation_columns,
    transform_oos,
    generate_excel_oos,
    _add_combined_recommendation_columns,
    _shop_simplify_top80_keys,
    _shop_simplify_top80_mask,
    transform_ndrf_comparison,
    _build_sku_soh_tier_summary,
    _hybrid_dashboard_metrics,
    _hybrid_stats_from_raw,
    _write_metric_cards,
    transform_forecast,
)
# transformer 已完成初始化後才個別取用（此名稱不在 star-export 清單）
from modes.stock_impact import _resolve_period_stock_column  # noqa: E402


_RATES_OOS = {"default": 0.5, "tourist": 0.5, "mix": 0.5}
_RATES_OS = {"default": 0.5, "tourist": 0.3, "mix": 0.4}


def _oos_row(**kw):
    base = {
        "Buyer": "B1", "Brand": "A", "SKU": 1, "Shop Stock Status": "Normal",
        "Shop Nature": " Tourist ", "RP Type": "RF", "Adj. Safety Stock": 5.0,
        "Sold Qty (CM - MTD)": 0.0, "Shop Monthly Sold Qty (P90D)": 0.0,
        "Launch Date": (date.today() - timedelta(days=10)).strftime("%Y%m%d"),
    }
    if "LaunchDate" in kw:
        base["Launch Date"] = kw.pop("LaunchDate")
    base.update(kw)
    return base


class TestOosNewSkuFloor:
    def test_new_sku_zero_sales_keeps_floor3(self):
        # 高 1：_parse_launch_dates 漏 import 曾令新品保底永遠失效（NameError 被吞）
        df = pd.DataFrame([_oos_row()])
        out = _add_oos_recommendation_columns(
            df, _RATES_OOS, 15, 30, safety_mode="allow_both", oos_calc_mode="mtd"
        )
        assert int(out["Recommendation of new Safety Qty"].iloc[0]) == 3
        assert int(out["Safety Qty Increase"].iloc[0]) == -2

    def test_old_zero_sales_stays_zero(self):
        df = pd.DataFrame([_oos_row(
            LaunchDate=(date.today() - timedelta(days=400)).strftime("%Y%m%d")
        )])
        out = _add_oos_recommendation_columns(
            df, _RATES_OOS, 15, 30, safety_mode="allow_both", oos_calc_mode="mtd"
        )
        assert int(out["Recommendation of new Safety Qty"].iloc[0]) == 0

    def test_nature_with_spaces_uses_tourist_rate(self):
        # 中 6：帶空格的「 Tourist 」須套 Tourist 倍率（MTD 外推 16×0.5＝8），
        # 若誤套 default（0.9）會得 15
        df = pd.DataFrame([_oos_row(**{
            "Sold Qty (CM - MTD)": 8.0, "Launch Date": "2024-01-01",
        })])
        out = _add_oos_recommendation_columns(
            df, {"default": 0.9, "tourist": 0.5, "mix": 0.5}, 15, 30,
            safety_mode="allow_both", oos_calc_mode="mtd",
        )
        assert int(out["Recommendation of new Safety Qty"].iloc[0]) == 8

    def test_projected_negative_clamped_to_zero(self):
        df = pd.DataFrame([_oos_row(**{"Sold Qty (CM - MTD)": -3.0, "Launch Date": "2024-01-01"})])
        out = _add_oos_recommendation_columns(
            df, _RATES_OOS, 15, 30, safety_mode="allow_both", oos_calc_mode="mtd"
        )
        assert int(out["Projected Monthly Sold Qty"].iloc[0]) == 0


class TestOosDashboardCounts:
    def test_dashboard_counts_only_abnormal(self):
        rows = [
            _oos_row(SKU=1, **{"Shop Stock Status": "OOS (0 Stock)"}),
            _oos_row(SKU=2, Buyer="B2", Brand="Z", **{"Shop Stock Status": "Normal"}),
        ]
        summary, raw, cats, ws = transform_oos(
            pd.DataFrame(rows), rates=_RATES_OOS, mtd_days=15, month_total_days=30,
        )
        buf = generate_excel_oos(summary, raw, cats)
        wb = load_workbook(BytesIO(buf.getvalue()))
        ws_dash = wb["OOS庫存數據分析摘要"]
        # row1_metrics：名稱列在第 3 行、數值列在第 4 行；只計有異常者 → 皆為 1
        assert ws_dash.cell(row=4, column=2).value == 1  # Buyers
        assert ws_dash.cell(row=4, column=3).value == 1  # Brands
        assert ws_dash.cell(row=4, column=4).value == 1  # SKUs


class TestSafetyReviewX:
    def _run(self, launch=None):
        kw = {} if launch is None else {"LaunchDate": launch}
        df = pd.DataFrame([_oos_row(**kw)])
        return _add_combined_recommendation_columns(
            df, None, None, 15, 30, safety_mode="allow_both",
            ref_date=date.today(), safety_basis="rate",
        )

    def test_old_zero_sales_x_zero(self):
        # 中 7：舊品零銷量 X 不保底（=0），可加可減下建議值歸 0
        out = self._run(launch=(date.today() - timedelta(days=400)).strftime("%Y%m%d"))
        assert int(out["Overstock Cap (X)"].iloc[0]) == 0
        assert int(out["Recommend Safety Stock"].iloc[0]) == 0

    def test_new_zero_sales_x_floor3(self):
        out = self._run()
        assert int(out["Overstock Cap (X)"].iloc[0]) == 3

    def test_nature_spaces_use_tourist_days(self):
        # 中 6：帶空格 Nature 須套 Tourist 天數（turnday 基準）
        df = pd.DataFrame([_oos_row(**{
            "Sold Qty (P90D)": 900.0, "Launch Date": "2024-01-01",
        })])
        out = _add_combined_recommendation_columns(
            df, None, None, 15, 30, safety_mode="allow_both",
            ref_date=date.today(), safety_basis="turnday",
        )
        # Tourist 10 日：ceil(900/90×10)=100；default（Non Tourist 7 日）會是 70
        assert int(out["Safety 周轉所需量"].iloc[0]) == 100


class TestStockImpactTier:
    def _raw(self):
        return pd.DataFrame([
            {
                "Buyer": "B1", "Brand": "A", "SKU": "S1", "Site Code": "S01",
                "Shop Nature": "Tourist", "RP Type": "RF",
                "Shop Stock Qty (LLM)": 5, "Shop Stock Qty (LM)": 3, "Shop Stock Qty (CM)": 9,
                "D001 Stock Qty (LLM)": 120, "D001 Stock Qty (LM)": 110, "D001 Stock Qty (CM)": 100,
                "Sold Qty (M-2)": 1, "Sold Qty (M-1)": 2, "Sold Qty (CM - MTD)": 3,
                "Shop Monthly Sold Qty (P90D)": 2,
            },
            {
                "Buyer": "B1", "Brand": "A", "SKU": "S1", "Site Code": "S02",
                "Shop Nature": "Mix", "RP Type": "RF",
                "Shop Stock Qty (LLM)": 0, "Shop Stock Qty (LM)": 2, "Shop Stock Qty (CM)": 9,
                "D001 Stock Qty (LLM)": 120, "D001 Stock Qty (LM)": 110, "D001 Stock Qty (CM)": 100,
                "Sold Qty (M-2)": 0, "Sold Qty (M-1)": 1, "Sold Qty (CM - MTD)": 1,
                "Shop Monthly Sold Qty (P90D)": 1,
            },
        ])

    def test_resolver_d001_prefix(self):
        df = self._raw()
        assert _resolve_period_stock_column(df, "M-2", prefix="D001 Stock Qty") == "D001 Stock Qty (LLM)"
        assert _resolve_period_stock_column(df, "M-1", prefix="D001 Stock Qty") == "D001 Stock Qty (LM)"

    def test_tier_blocks_differ_by_period_and_d001_proxied(self):
        # 高 2＋中 1／中 2：三期 tier 欄名帶期別；D001／店舖 M-2／M-1 經 LLM／LM 代理
        out = _build_sku_soh_tier_summary(self._raw())
        assert out["a 0 (M-2)"].iloc[0] == 1  # LLM: 5/0 → e 5-7×1、a 0×1
        assert out["e 5-7 (M-2)"].iloc[0] == 1
        assert out["a 0 (M-1)"].iloc[0] == 0  # LM: 3/2 → d 3-4×1、c 2×1
        assert out["c 2 (M-1)"].iloc[0] == 1
        assert out["d 3-4 (M-1)"].iloc[0] == 1
        assert out["總計 (M-2)"].iloc[0] == 2
        # D001 代理：Total Stock (M-2) = 120 + (5+0) = 125
        assert int(out["Total Stock Qty (M-2)"].iloc[0]) == 125
        # 有貨店舖 (M-2)：LLM 5/0 → 1 間
        assert int(out["有貨店舖 (M-2)"].iloc[0]) == 1


class TestForecastD001First:
    def test_first_row_blank_takes_next_non_null(self):
        base = {
            "Buyer": "B1", "Brand": "A", "SKU": 1, "Site Code": "S01",
            "Shop Nature": "Mix", "RP Type": "RF", "Supply Source": "Warehouse",
            "Shop Stock Status": "Normal", "Adj. Safety Stock": 3.0,
            "Shop Monthly Sold Qty (P90D)": 60.0, "Sold Qty (P90D)": 1800.0,
            "Sold Qty (M-2)": 50.0, "Sold Qty (M-1)": 55.0, "Sold Qty (CM - MTD)": 16.0,
            "Shop Stock Qty (CM)": 20.0, "D001 Stock Qty (CM)": 100.0,
            "D001 On Order PO Qty (CM)": 10.0, "Shop On Order PO Qty (CM)": 0.0,
            "Shop Pending Receive (±7 days)": 5.0, "Launch Date": "2024-01-01",
        }
        rows = []
        for site, d001_net in [("S01", None), ("S02", 200.0), ("S03", 200.0), ("S04", 200.0)]:
            r = dict(base)
            r["Site Code"] = site
            r["D001 Net Stock Qty (CM)"] = d001_net
            rows.append(r)
        main, coverage, buyer_out, raw_out, stats, warns = transform_forecast(
            pd.DataFrame(rows), mtd_days=8, month_total_days=30, mtd_date=date(2026, 9, 8),
        )
        # 中 10：首列空白不再令 D001 庫存取 0（正確取首個非空值 200）
        assert int(main["FC D001 庫存 (Net)"].iloc[0]) == 200


class TestShopSimplifyBlankSku:
    def test_blank_sku_excluded_from_population(self):
        groups = pd.Series(["A", "A", "A"])
        skus = pd.Series(["SKU1", "SKU2", None])
        qty = pd.Series([70.0, 10.0, 10.0])
        keys = _shop_simplify_top80_keys(groups, skus, qty, share=0.8)
        # 剔除空白 SKU 後總量 80、門檻 64 → 只有 SKU1 是 Top
        assert ("A", "SKU1") in keys
        assert ("A", "SKU2") not in keys
        assert all(sku != "nan" for _, sku in keys)

    def test_money_tail_no_guard(self):
        values = pd.Series([0.9, 0.5], index=["a", "b"])
        mask = _shop_simplify_top80_mask(values, share=0.8, tail_guard=False)
        assert list(mask) == [True, True]
        mask_qty = _shop_simplify_top80_mask(values, share=0.8, tail_guard=True)
        assert list(mask_qty) == [True, False]


def _ndrf_df():
    return pd.DataFrame([
        {
            "Buyer": "B1", "Brand": "A", "SKU": "S1", "Site Code": "S01",
            "RP Type": "RF", "Shop Stock Qty (CM)": 3, "Adj. SST": "X",
            "Adj. MCH1": "M", "Shop Stock Status": "Normal", "Month": 202601,
        },
        {
            "Buyer": "B1", "Brand": "A", "SKU": "S2", "Site Code": "S01",
            "RP Type": "ND", "Shop Stock Qty (CM)": -5, "Adj. SST": "X",
            "Adj. MCH1": "M", "Shop Stock Status": "Normal", "Month": 202601,
        },
        {
            "Buyer": "B1", "Brand": "A", "SKU": "S1", "Site Code": "S01",
            "RP Type": "RF", "Shop Stock Qty (CM)": 4, "Adj. SST": "X",
            "Adj. MCH1": "M", "Shop Stock Status": "Normal", "Month": 202602,
        },
    ])


class TestNdrf:
    def test_numeric_month_and_negative_nd(self):
        # 中 14b／14c：Month 數值型可比對；ND 負庫存不計入 RF_Total
        comp, sst, mch1, raw, ws = transform_ndrf_comparison(_ndrf_df(), 202601, 202602)
        row = comp[comp["Site Code"] == "S01"].iloc[0]
        assert int(row["RF (202601)"]) == 1
        assert int(row["Carry (202601)"]) == 1  # ND 負庫存不算 Carry（與樞紐同一口徑）
        sst_row = sst[sst["Site Code"] == "S01"].iloc[0]
        assert int(sst_row["RF Total 202601"]) == 1
        assert int(sst_row["ND With Stock 202601"]) == 0

    def test_missing_adj_mch1_no_crash(self):
        df = _ndrf_df().drop(columns=["Adj. MCH1"])
        comp, sst, mch1, raw, ws = transform_ndrf_comparison(df, "202601", "202602")
        assert mch1.empty
        assert any("Adj. MCH1" in w for w in ws)


class TestHybridRow8Metrics:
    def test_write_metric_cards_supports_row8(self):
        from openpyxl import Workbook
        from openpyxl.styles import Font
        wb = Workbook()
        ws = wb.active
        desc_font = Font(size=9, color="6B7280", italic=True)
        thin = None
        cards = [("卡", 1, "說明", "EF4444")]
        _write_metric_cards(
            ws, "測試", cards, cards, thin, desc_font,
            row3_metrics=cards, row4_metrics=cards, row5_metrics=cards,
            row6_metrics=cards, row7_metrics=cards, row8_metrics=cards,
        )
        assert ws.cell(row=31, column=1).value == "卡"
        assert ws.cell(row=32, column=1).value == 1

    def test_dashboard_metrics_return_8_rows_under_full_flags(self):
        # 以 stats 旗標直接驗證排數上限（第 8 排＝退D001 概況）
        raw = pd.DataFrame({
            "SKU": ["S1"], "轉出量": [0],
            "建議退D001 量": [0],
        })
        stats = _hybrid_stats_from_raw(raw)
        stats.update({
            "bidirectional": True, "nd_priority": False,
            "sales_rebalance": True, "real_supply_gate": True,
        })
        rows = _hybrid_dashboard_metrics(raw, stats)
        assert len(rows) == 8
        assert rows[7][0][0] == "退D001 列數"
