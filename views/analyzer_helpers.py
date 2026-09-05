from __future__ import annotations

import pandas as pd


def build_hybrid_params(
    turnday_threshold: int,
    d001_threshold: int,
    safety_days: dict,
    bidirectional: bool,
) -> dict:
    return {
        "turnday_threshold": int(turnday_threshold),
        "d001_threshold": int(d001_threshold),
        "safety_days": {str(key): int(value) for key, value in dict(safety_days).items()},
        "bidirectional": bool(bidirectional),
    }


def hybrid_params_changed(applied: dict | None, current: dict) -> bool:
    return applied is not None and applied != current


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
