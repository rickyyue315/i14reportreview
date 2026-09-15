from __future__ import annotations

import pandas as pd


def build_hybrid_params(
    turnday_threshold: int,
    d001_threshold: int,
    safety_days: dict,
    bidirectional: bool,
    dimension: str = "nature",
    ref_date=None,
    nd_priority: bool = False,
    sales_rebalance: bool = False,
    sales_oos_days: int = 7,
    sales_overstock_days: int = 30,
    sales_overstock_min_units: int = 2,
    return_to_d001: bool = False,
) -> dict:
    if ref_date is None:
        ref_text = ""
    else:
        try:
            ref_text = pd.to_datetime(ref_date).date().isoformat()
        except (TypeError, ValueError):
            ref_text = str(ref_date).strip()
    return {
        "turnday_threshold": int(turnday_threshold),
        "d001_threshold": int(d001_threshold),
        "safety_days": {str(key): int(value) for key, value in dict(safety_days).items()},
        "bidirectional": bool(bidirectional),
        "dimension": str(dimension),
        "ref_date": ref_text,
        "nd_priority": bool(nd_priority),
        "sales_rebalance": bool(sales_rebalance),
        "sales_oos_days": int(sales_oos_days),
        "sales_overstock_days": int(sales_overstock_days),
        "sales_overstock_min_units": int(sales_overstock_min_units),
        "return_to_d001": bool(return_to_d001),
    }


def hybrid_params_changed(applied: dict | None, current: dict) -> bool:
    if applied is None:
        return False
    if set(applied) != set(current):
        # 舊快照缺新增鍵時以預設值補齊再比對（向後相容）
        defaults = {
            "ref_date": "",
            "nd_priority": False,
            "sales_rebalance": False,
            "sales_oos_days": 7,
            "sales_overstock_days": 30,
            "sales_overstock_min_units": 2,
            "return_to_d001": False,
        }
        applied = {**{k: applied.get(k, d) for k, d in defaults.items()}, **applied}
        current = {**{k: current.get(k, d) for k, d in defaults.items()}, **current}
    return applied != current


def canonical_site_codes(values) -> list[str]:
    codes = {
        str(value).strip()
        for value in values
        if pd.notna(value) and str(value).strip()
    }
    return sorted(codes)


def filter_by_site_code(df: pd.DataFrame, site_code: str) -> pd.DataFrame:
    if site_code == "全部" or "Site Code" not in df.columns:
        return df.copy()
    codes = df["Site Code"].fillna("").astype(str).str.strip()
    return df.loc[codes == str(site_code).strip()].copy()
