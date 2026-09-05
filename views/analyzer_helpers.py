from __future__ import annotations

import pandas as pd


def build_hybrid_params(
    turnday_threshold: int,
    d001_threshold: int,
    safety_days: dict,
    bidirectional: bool,
    dimension: str = "nature",
    ref_date=None,
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
    }


def hybrid_params_changed(applied: dict | None, current: dict) -> bool:
    if applied is None:
        return False
    if set(applied) != set(current):
        # 舊快照缺 ref_date 鍵時以空值補齊再比對（向後相容）
        applied = {**applied, "ref_date": applied.get("ref_date", "")}
        current = {**current, "ref_date": current.get("ref_date", "")}
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
