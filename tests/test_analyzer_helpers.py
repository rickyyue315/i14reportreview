import ast
import os
from datetime import datetime
from io import BytesIO
from pathlib import Path
from types import SimpleNamespace
from zoneinfo import ZoneInfo

import pandas as pd
from openpyxl import Workbook

from views.analyzer_helpers import (
    app_last_updated_hkt,
    build_hybrid_params,
    canonical_site_codes,
    filter_by_site_code,
    hybrid_params_changed,
    pick_source_sheet,
    read_source_excel,
)


def test_pick_source_sheet_prefers_export():
    assert pick_source_sheet(["工作表1", "Export"]) == "Export"
    assert pick_source_sheet(["Export"]) == "Export"
    assert pick_source_sheet([" Export "]) == " Export "
    assert pick_source_sheet(["工作表1", "工作表2"]) == "工作表1"
    assert pick_source_sheet([]) == 0
    assert pick_source_sheet(None) == 0


def _workbook_bytes(named_sheets):
    wb = Workbook()
    first = True
    for name, rows in named_sheets:
        ws = wb.active if first else wb.create_sheet()
        ws.title = name
        first = False
        for row in rows:
            ws.append(row)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_read_source_excel_reads_export_sheet_when_not_first():
    data = _workbook_bytes([
        ("工作表1", [["樞紐"], [1]]),
        ("Export", [["SKU", "Sold Qty (P90D)"], ["A1", 0]]),
    ])
    df = read_source_excel(data)
    assert list(df.columns) == ["SKU", "Sold Qty (P90D)"]
    assert str(df.iloc[0]["SKU"]) == "A1"


def test_read_source_excel_falls_back_to_first_sheet():
    data = _workbook_bytes([("Summary", [["Market"], ["HK"]])])
    df = read_source_excel(data)
    assert list(df.columns) == ["Market"]
    assert df.iloc[0]["Market"] == "HK"


def test_hybrid_params_snapshot_and_stale_detection():
    params = build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False)
    assert not hybrid_params_changed(params, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False))
    assert hybrid_params_changed(params, build_hybrid_params(
        14, 100, {"Tourist": 8, "Mix": 9, "Non Tourist": 10}, False))
    assert hybrid_params_changed(params, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, True))
    assert hybrid_params_changed(params, build_hybrid_params(
        15, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False))


def test_hybrid_params_dimension_default_and_stale_detection():
    base = build_hybrid_params(14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False)
    assert base["dimension"] == "nature"  # 預設維度（舊呼叫未傳維度仍相容）
    assert base["ref_date"] == ""  # 未傳參考日期時為空字串
    assert not hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "nature"))
    assert hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "rp"))
    assert hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "both"))


def test_hybrid_params_ref_date_stale_detection():
    import datetime
    base = build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "nature",
        datetime.date(2026, 8, 3))
    assert base["ref_date"] == "2026-08-03"
    assert not hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "nature",
        datetime.date(2026, 8, 3)))
    assert hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "nature",
        datetime.date(2026, 8, 4)))
    # 舊快照缺 ref_date 鍵時不誤報（向後相容）
    legacy = {k: v for k, v in base.items() if k != "ref_date"}
    assert not hybrid_params_changed(legacy, dict(legacy))


def test_nd_priority_stale_detection_and_legacy_snapshot():
    base = build_hybrid_params(14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False)
    assert base["nd_priority"] is False  # 舊呼叫未傳開關仍相容
    assert not hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "nature", None, False))
    assert hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "nature", None, True))
    # 舊快照缺 nd_priority 鍵時不誤報（向後相容）
    legacy = {k: v for k, v in base.items() if k != "nd_priority"}
    assert not hybrid_params_changed(legacy, dict(legacy))
    legacy_no_ref = {k: v for k, v in base.items() if k not in {"nd_priority", "ref_date"}}
    assert not hybrid_params_changed(legacy_no_ref, dict(legacy_no_ref))


def test_site_code_options_are_canonical_strings():
    assert canonical_site_codes([1002, "1001", " 1002 ", None, "", " B01 "]) == ["1001", "1002", "B01"]


def test_site_code_filter_handles_numeric_and_string_values():
    frame = pd.DataFrame({"Site Code": [1001, "1002", " 1001 ", None], "SKU": [1, 2, 3, 4]})
    filtered = filter_by_site_code(frame, "1001")
    assert list(filtered["SKU"]) == [1, 3]
    assert len(filter_by_site_code(frame, "全部")) == 4
    assert filter_by_site_code(frame, "missing").empty


def test_top_download_button_respects_disabled_state():
    page = Path(__file__).resolve().parents[1] / "views" / "analyzer_page.py"
    tree = ast.parse(page.read_text(encoding="utf-8"))
    function = next(node for node in tree.body if isinstance(node, ast.FunctionDef) and node.name == "_top_download_button")
    calls = []
    namespace = {
        "datetime": datetime, "ZoneInfo": ZoneInfo,
        "st": SimpleNamespace(download_button=lambda **kwargs: calls.append(kwargs)),
    }
    # 只載入下載函數，避免匯入頁面時執行 Streamlit 上傳與導覽流程。
    exec(compile(ast.Module(body=[function], type_ignores=[]), str(page), "exec"), namespace)
    button = namespace["_top_download_button"]
    for disabled in (False, True):
        button(b"report", "Hybrid", "dl_top_hybrid", disabled=disabled)
        assert calls[-1]["disabled"] is disabled
        assert calls[-1]["data"] == b"report"
    button(b"report", "Other", "dl_other")
    assert calls[-1]["disabled"] is False
    button(None, "Hybrid", "dl_top_hybrid", disabled=True)
    assert len(calls) == 3


def test_hybrid_download_entry_points_share_stale_guard():
    page = Path(__file__).resolve().parents[1] / "views" / "analyzer_page.py"
    tree = ast.parse(page.read_text(encoding="utf-8"))
    calls = [node for node in ast.walk(tree) if isinstance(node, ast.Call)]
    top = next(
        node for node in calls
        if isinstance(node.func, ast.Name) and node.func.id == "_top_download_button"
        and any(isinstance(arg, ast.Constant) and arg.value == "dl_top_hybrid" for arg in node.args)
    )
    bottom = next(
        node for node in calls
        if isinstance(node.func, ast.Attribute) and node.func.attr == "download_button"
        and any(keyword.arg == "label" and isinstance(keyword.value, ast.Constant)
                and "Safety Stock x NetStock Review" in str(keyword.value.value) for keyword in node.keywords)
    )
    for call in (top, bottom):
        disabled = next(keyword.value for keyword in call.keywords if keyword.arg == "disabled")
        assert isinstance(disabled, ast.Name) and disabled.id == "hybrid_is_stale"


def test_app_last_updated_hkt_uses_newest_python_mtime(tmp_path):
    """頁尾時間取根目錄／views／modes 內最新 .py 的修改時間，非程式碼檔案不計。"""
    (tmp_path / "views").mkdir()
    (tmp_path / "modes").mkdir()
    (tmp_path / "version.py").write_text("", encoding="utf-8")
    (tmp_path / "views" / "usage_page.py").write_text("", encoding="utf-8")
    (tmp_path / "modes" / "shop_simplify.py").write_text("", encoding="utf-8")
    (tmp_path / "RAW 明細.xlsx").write_text("", encoding="utf-8")
    os.utime(tmp_path / "version.py", (1_700_000_000, 1_700_000_000))
    os.utime(tmp_path / "modes" / "shop_simplify.py", (1_700_001_800, 1_700_001_800))
    os.utime(tmp_path / "views" / "usage_page.py", (1_700_003_600, 1_700_003_600))
    # 非 .py 檔更新也不影響：時間戳刻意設得最新
    os.utime(tmp_path / "RAW 明細.xlsx", (1_700_100_000, 1_700_100_000))

    assert app_last_updated_hkt(tmp_path) == "2023-11-15 07:13"


def test_app_last_updated_hkt_falls_back_to_now(tmp_path):
    """讀不到任何 .py 時退回傳入的目前時間（HKT），確保頁尾不會留空。"""
    moment = datetime(2026, 11, 24, 10, 46, tzinfo=ZoneInfo("Asia/Hong_Kong"))
    assert app_last_updated_hkt(tmp_path, now=moment) == "2026-11-24 10:46"
