# -*- coding: utf-8 -*-
"""「說明」工作表插入 Safety 計算邏輯圖的測試。

覆蓋：
- 圖片按指定寬度等比縮放、錨點正確；
- 圖檔缺失時靜默略過（報告照常輸出）；
- Safety Stock Review 的說明表確實在 ⑤ 段之後帶圖；
- 未傳圖的說明表維持原樣（不應影響其他模式）。
"""
import os
from io import BytesIO

from openpyxl import Workbook, load_workbook

import transformer  # noqa: F401  （先載入墊片，避免 modes 循環匯入）
from modes.common import _add_guide_image, _write_guide_sheet
from modes.safety_review import _write_combined_guide_sheet

ASSET = os.path.join("assets", "safety_logic_guide.png")


def _roundtrip(wb):
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return load_workbook(buf)


def test_guide_image_scaled_and_anchored():
    wb = Workbook()
    ws = wb.active
    assert _add_guide_image(ws, ASSET, anchor="A5", width_px=900) is True
    assert len(ws._images) == 1

    im = _roundtrip(wb).active._images[0]
    # 顯示尺寸以 EMU 儲存：900 × (2100 ÷ 1750) = 1080
    assert round(im.anchor.ext.cx / 9525) == 900
    assert round(im.anchor.ext.cy / 9525) == 1080
    assert (im.anchor._from.col, im.anchor._from.row) == (0, 4)  # A5（0-based）


def test_guide_image_missing_file_is_silent():
    wb = Workbook()
    ws = wb.active
    assert _add_guide_image(ws, "assets/not_exists.png") is False
    assert len(ws._images) == 0

    # 經 _write_guide_sheet 傳入不存在的圖亦不應拋錯
    _write_guide_sheet(ws, title="T", subtitle="S", params=[("a", "b")],
                       logic=["x"], sheet_rows=[("s", "c", "n")], tips=["t"],
                       image="assets/not_exists.png", image_section="⑤ 圖")
    assert len(ws._images) == 0


def test_safety_guide_sheet_has_figure_after_section_5():
    wb = Workbook()
    ws = wb.active
    ws.title = "說明"
    _write_combined_guide_sheet(ws, {}, None)

    sheet = _roundtrip(wb)["說明"]
    assert len(sheet._images) == 1

    sec5 = [r for r in range(1, sheet.max_row + 1)
            if str(sheet.cell(row=r, column=1).value or "").startswith("⑤")]
    assert sec5, "說明表應有 ⑤ Safety 計算邏輯圖 段"
    # 錨點為 0-based 列索引：等於標題的 1-based 列號 → 即落在標題的下一列
    assert sheet._images[0].anchor._from.row == sec5[0]


def test_guide_sheet_without_image_stays_plain():
    wb = Workbook()
    ws = wb.active
    _write_guide_sheet(ws, title="T", subtitle="S", params=[("a", "b")],
                       logic=["x"], sheet_rows=[("s", "c", "n")], tips=["t"])
    assert len(ws._images) == 0
    assert ws.max_column == 3
