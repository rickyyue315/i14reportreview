from __future__ import annotations

from datetime import datetime
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

# 香港時間（UTC+8，無夏令時間）；容器與雲端平台預設為 UTC，顯示前須明確轉換
_HKT = ZoneInfo("Asia/Hong_Kong")

# 頁尾「最後更新」只計程式碼：專案根目錄、views、modes 三個資料夾的 .py
_APP_SOURCE_DIRS = ("", "views", "modes")


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
    urgent_first: bool = False,
    pool_nd_exclude: bool = False,
    real_supply_gate: bool = False,
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
        "urgent_first": bool(urgent_first),
        "pool_nd_exclude": bool(pool_nd_exclude),
        "real_supply_gate": bool(real_supply_gate),
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
            "urgent_first": False,
            "pool_nd_exclude": False,
            "real_supply_gate": False,
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


def pick_source_sheet(sheet_names):
    """選擇來源資料工作表：By SKUxShop 匯出檔的資料表恆為 Export；
    多工作表檔案（例如另加樞紐／摘要的版本）優先讀 Export，避免誤讀第一個工作表。
    沒有 Export 時沿用第一個工作表（與原有行為一致）；空清單回傳 0。"""
    names = list(sheet_names) if sheet_names is not None else []
    if not names:
        return 0
    for name in names:
        if isinstance(name, str) and name.strip() == "Export":
            return name
    return names[0]


def read_source_excel(raw_bytes):
    """讀取來源 Excel：優先 calamine（快 6 至 10 倍），失敗退回 openpyxl；
    工作表按 `pick_source_sheet` 選擇（多工作表時優先 Export）。"""
    last_error = None
    for engine in ("calamine", "openpyxl"):
        try:
            book = pd.ExcelFile(BytesIO(raw_bytes), engine=engine)
            try:
                return book.parse(pick_source_sheet(book.sheet_names))
            finally:
                book.close()
        except Exception as exc:
            last_error = exc
    raise last_error


def app_last_updated_hkt(root=None, now=None) -> str:
    """回傳程式碼最後修改日期時間（HKT），格式 `YYYY-MM-DD HH:MM`。

    以專案根目錄、`views`、`modes` 內 `.py` 檔案最新的修改時間為準：
    容器映像經 `Dockerfile` 的 COPY 保留來源檔案時間，雲端部署則為 checkout 時間，
    兩者都貼近實際更新時間。完全不寫死日期，避免換版後頁尾顯示過期時間。

    `root` 可指定其他專案根目錄（測試用）；讀不到任何 `.py` 時退回 `now`
    （預設目前時間），確保頁尾不會留空。
    """
    base = Path(root) if root is not None else Path(__file__).resolve().parent.parent
    newest = None
    for sub in _APP_SOURCE_DIRS:
        folder = base / sub if sub else base
        for path in folder.glob("*.py"):
            try:
                mtime = path.stat().st_mtime
            except OSError:
                continue
            if newest is None or mtime > newest:
                newest = mtime
    moment = datetime.fromtimestamp(newest, _HKT) if newest is not None else (now or datetime.now(_HKT))
    return moment.strftime("%Y-%m-%d %H:%M")
