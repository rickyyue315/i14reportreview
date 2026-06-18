"""End-to-end smoke test: every review mode's transform + generate_excel must complete."""
import warnings
warnings.filterwarnings("ignore")
import pandas as pd

from transformer import (
    transform, generate_excel,
    transform_oos, generate_excel_oos,
    transform_carry, generate_excel_carry,
    transform_combined, generate_excel_combined,
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
        {"SKU": "1", "Shop Stock Status": "P90D 0 Sales", "RP Type": " rf "},   # dirty RF
        {"SKU": "2", "Shop Stock Status": "P90D 0 Sales", "RP Type": "nd"},      # dirty ND
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

    print("\nALL PIPELINES COMPLETED")


if __name__ == "__main__":
    main()
