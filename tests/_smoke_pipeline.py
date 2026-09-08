"""End-to-end smoke test: every review mode's transform + generate_excel must complete."""
import pathlib
import sys
import warnings

warnings.filterwarnings("ignore")
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))
import pandas as pd

from transformer import (
    transform, generate_excel,
    transform_oos, generate_excel_oos,
    transform_carry, generate_excel_carry,
    transform_combined, generate_excel_combined,
    transform_stock_impact, generate_excel_stock_impact,
    transform_hybrid, generate_excel_hybrid,
    _hybrid_dimension_groups,
)

BASE = {
    "Buyer": "B1",
    "Brand": "BrandA",
    "SKU": "1001",
    "Site Code": "S01",
    "Item_Shop": "Prod-S01",
    "Shop Stock Status": "Normal",
    "Adj. Safety Stock": 5,
    "Shop Monthly Sold Qty (P90D)": 10,
    "Sold Qty (P90D)": 60,
    "Sold Qty (CM - MTD)": 8.0,
    "Shop Nature": "Local",
    "RP Type": "RF",
    "Shop Stock Qty (CM)": 3,
    "OM": "OM1",
    "Adj. MCH1": "MchA",
    "Adj. SST": "SstA",
}


def _row(**kw):
    r = dict(BASE); r.update(kw); return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


def main():
    # Diverse rows exercising every status / RP type path
    df = _df([
        {"SKU": "1", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": " rf "},   # dirty RF
        {"SKU": "2", "Shop Stock Status": "Overstock (P90D 0 Sales)", "RP Type": "nd"},      # dirty ND
        {"SKU": "3", "Shop Stock Status": "Overstock (>2*Monthly Sold Qty & >2pcs)",
         "Shop Monthly Sold Qty (P90D)": 50, "Adj. Safety Stock": 12},
        {"SKU": "4", "Shop Stock Status": "OOS (0 Stock)", "Shop Stock Qty (CM)": 0},
        {"SKU": "5", "Shop Stock Status": "Potential OOS (<7 Turnday)"},
        {"SKU": "6", "Shop Stock Status": "Normal", "RP Type": "RF"},
        {"SKU": "7", "Shop Stock Status": "Normal", "RP Type": "ST"},            # 'other' RP
        {"SKU": "8", "Shop Stock Status": "Normal", "RP Type": "XX"},            # non-ND/RF
    ])

    # ── Overstock Review ──
    s, r, c, w = transform(df, rates={"default": 0.5, "tourist": 0.3, "mix": 0.4})
    generate_excel(s, r, c)
    print("OK Overstock Review   | rows=%d warnings=%d" % (len(s), len(w)))

    # ── OOS Review (MTD, both modes) ──
    for mode in ("increase_only", "allow_both"):
        s, r, c, w = transform_oos(df, rates={"default": 0.5, "tourist": 0.3, "mix": 0.4},
                                   mtd_days=15, month_total_days=30,
                                   safety_mode=mode, oos_calc_mode="mtd")
        generate_excel_oos(s, r, c)
        print("OK OOS Review (%s)   | rows=%d warnings=%d" % (mode, len(s), len(w)))

    # ── Carry SKU Review ──
    sL, sR, dash, r, w = transform_carry(df)
    generate_excel_carry(sL, sR, dash, r)
    print("OK Carry SKU Review   | dash_rows=%d warnings=%d" % (len(dash), len(w)))

    # ── Safety Stock Review (both modes) ──
    ro = {"default": 0.5, "tourist": 0.3, "mix": 0.4}
    po = {"default": 0.5, "tourist": 0.5, "mix": 0.5}
    for mode in ("increase_only", "allow_both"):
        s, r, c, w = transform_combined(df, rates_overstock=ro, rates_oos=po,
                                        mtd_days=15, month_total_days=30, safety_mode=mode)
        generate_excel_combined(s, r, c)
        print("OK Safety Stock (%s)  | rows=%d warnings=%d" % (mode, len(s), len(w)))

    # ── Safety Stock x NetStock Review (調貨事件 Review Safety 設定) ──
    r, h_sc, h_diag, h_diag_om, h_cross, h_stats, h_w = transform_hybrid(df)
    generate_excel_hybrid(r, h_sc, h_diag, h_diag_om, h_cross, h_stats)
    print("OK Hybrid Diagnosis   | rows=%d need=%d diag=%d skew=%d warnings=%d" % (len(r), h_stats["need_rows"], len(h_diag), int((h_sc["存貨偏重提示"].fillna("").astype(str) != "").sum()) if not h_sc.empty else 0, len(h_w)))

    # ── Safety Stock x NetStock Review（允許雙向調貨）──
    r2, h_sc2, h_diag2, h_diag2_om, h_cross2, h_stats2, h_w2 = transform_hybrid(df, bidirectional=True)
    generate_excel_hybrid(r2, h_sc2, h_diag2, h_diag2_om, h_cross2, h_stats2)
    print("OK Hybrid Bidirectional | rows=%d need=%d diag=%d tourist_out=%d warnings=%d" % (len(r2), h_stats2["need_rows"], len(h_diag2), h_stats2["tourist_transfer"], len(h_w2)))

    # ── Safety Stock x NetStock Review（銷量再平衡＋以銷量分類維度）──
    r3, h_sc3, h_diag3, h_diag3_om, h_cross3, h_stats3, h_w3 = transform_hybrid(
        df, sales_rebalance=True, dimension="sales")
    generate_excel_hybrid(r3, h_sc3, h_diag3, h_diag3_om, h_cross3, h_stats3)
    print("OK Hybrid Sales Rebalance | rows=%d need=%d sales_need=%d oos=%d overstock=%d warnings=%d" % (
        len(r3), h_stats3["need_rows"], h_stats3["sales_need_rows"],
        h_stats3["sales_oos_rows"], h_stats3["sales_overstock_rows"], len(h_w3)))

    # ── Safety Stock x NetStock Review（以銷量＋Shop Nature 分類維度）──
    r3b, h_sc3b, h_diag3b, h_diag3b_om, h_cross3b, h_stats3b, h_w3b = transform_hybrid(
        df, sales_rebalance=True, dimension="sales_nature")
    assert h_stats3b["dimension"] == "sales_nature"
    generate_excel_hybrid(r3b, h_sc3b, h_diag3b, h_diag3b_om, h_cross3b, h_stats3b)
    print("OK Hybrid Sales x Nature  | rows=%d need=%d groups=%d warnings=%d" % (
        len(r3b), h_stats3b["need_rows"], len(_hybrid_dimension_groups("sales_nature")), len(h_w3b)))

    # ── Stock Impact Analysis ──
    _, si_brand, si_buyer, si_cat, _, si_alloc, si_diag, si_inv, si_mly, si_raw, si_warns = transform_stock_impact(
        df, warehouse_only=False, stock_threshold=300
    )
    generate_excel_stock_impact(si_brand, si_buyer, si_cat, si_alloc, si_diag, si_inv, si_raw, sku_monthly_analysis=si_mly)
    si_count = int(si_raw["Is Stock Impacted"].sum())
    print("OK Stock Impact Analysis | impacted_rows=%d diagnosis_skus=%d imbalance_skus=%d warnings=%d" % (si_count, len(si_diag), len(si_inv.get("imbalance_sku", [])), len(si_warns)))

    print("\nALL PIPELINES COMPLETED")


if __name__ == "__main__":
    main()
