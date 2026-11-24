# -*- coding: utf-8 -*-
"""SCT建議 16 種情況設定表（初步設定用；非應用程式元件）。

四個 Top N% 判定各只有「是／否」，兩個對比欄由判定自動合成，所以全部情況
只有 2×2×2×2＝16 種。本表把 16 種情況逐列列出，讓 SCT 逐格指定「首選RF／
次選RF／建議RF／留空」，並同場對照五個內建方案會怎樣分。

輸出：`SCT建議_16種情況設定表_YYYYMMDD.xlsx`
- 情況一覽：16 種情況＋下拉選單（填好後可直接照此在應用程式「自訂 16 種情況」設定）
- 預設方案：四個內建方案逐格結果（顏色標示）
- 方案說明：各方案的判定邏輯與適用情境
"""
import datetime
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from modes.shop_simplify import (  # noqa: E402
    _SHOP_SIMPLIFY_SCT_CASES,
    _SHOP_SIMPLIFY_SCT_OPTIONS,
    _SHOP_SIMPLIFY_SCT_PREFERRED,
    _SHOP_SIMPLIFY_SCT_PRESETS,
    _SHOP_SIMPLIFY_SCT_RF,
    _SHOP_SIMPLIFY_SCT_SECONDARY,
    _shop_simplify_sct_case_compare,
    _shop_simplify_sct_case_text,
)

BLUE = "D9E2F3"
GREEN = "CCFF99"
YELLOW = "FFEB9C"
ORANGE = "FFCC99"
TEAL = "99E6CC"
PINK = "FFC7CE"
GREY = "F2F2F2"

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def _header(ws, row: int, labels: list[str], widths: list[float]) -> None:
    for col, (label, width) in enumerate(zip(labels, widths), start=1):
        cell = ws.cell(row=row, column=col, value=label)
        cell.font = Font(bold=True)
        cell.fill = _fill(BLUE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[row].height = 42


def _title(ws, text: str, span: int, row: int = 1, size: int = 14) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, size=size)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 26


def _note(ws, text: str, span: int, row: int) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(size=10, color="555555")
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32


def _case_row(ws, row: int, index: int, case: str, extra: str = "") -> None:
    qty_cmp, val_cmp = _shop_simplify_sct_case_compare(case)
    values = [
        index,
        "是" if case[0] == "Y" else "否",
        "是" if case[1] == "Y" else "否",
        qty_cmp,
        "是" if case[2] == "Y" else "否",
        "是" if case[3] == "Y" else "否",
        val_cmp,
        _shop_simplify_sct_case_text(case),
        extra,
    ]
    for col, value in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col, value=value)
        cell.border = BORDER
        if col in (1,):
            cell.alignment = Alignment(horizontal="center")
        elif col in (2, 3, 5, 6):
            cell.alignment = Alignment(horizontal="center")
            if value == "是":
                cell.fill = _fill(GREEN)
        elif col in (4, 7):
            cell.alignment = Alignment(horizontal="center")
            if value in ("兩者皆是",):
                cell.fill = _fill(TEAL)
            elif value in ("兩者皆非",):
                cell.fill = _fill(PINK)
        elif col == 8:
            cell.alignment = Alignment(vertical="center", wrap_text=True)


def build(output_dir: Path) -> Path:
    wb = Workbook()

    # ── 工作表 1：16 種情況＋你的選擇 ──
    ws = wb.active
    ws.title = "情況一覽"
    widths = [6, 14, 20, 13, 14, 20, 13, 44, 16, 24]
    _title(ws, "SCT建議：16 種 Top 判定組合（在此指定每種情況的建議）", len(widths))
    _note(
        ws,
        "四個判定各只有「是／否」，兩個對比欄由判定自動合成，故全部情況＝2×2×2×2＝16 種。"
        "請在「你的SCT建議」欄逐格選擇（下拉：首選RF／次選RF／建議RF，留空＝由 SCT 自行判斷，例如 ND）；"
        "填好後可在應用程式「SCT建議（依四個 Top 判定自動填寫）」的「自訂 16 種情況」照此設定。",
        len(widths), 2,
    )
    headers = [
        "編號", "件數－店舖", "件數－同Shop Nature", "件數對比",
        "金額－店舖", "金額－同Shop Nature", "金額對比", "情況說明", "你的SCT建議", "備註",
    ]
    _header(ws, 3, headers, widths)

    first_row = 4
    for i, case in enumerate(_SHOP_SIMPLIFY_SCT_CASES, start=1):
        _case_row(ws, first_row + i - 1, i, case)
    last_row = first_row + len(_SHOP_SIMPLIFY_SCT_CASES) - 1

    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(v for v in _SHOP_SIMPLIFY_SCT_OPTIONS if v) + '"',
        allow_blank=True,
        showDropDown=False,
        errorTitle="只可選固定值",
        error="請選擇：首選RF／次選RF／建議RF，或留空。",
        promptTitle="SCT建議",
        prompt="按情況選擇建議；留空＝不自動填寫。",
    )
    dv.add(f"I{first_row}:I{last_row}")
    ws.add_data_validation(dv)

    ws.freeze_panes = f"A{first_row}"
    ws.auto_filter.ref = f"A3:J{last_row}"

    # ── 工作表 2：四個內建方案對照 ──
    ws2 = wb.create_sheet("預設方案")
    preset_names = list(_SHOP_SIMPLIFY_SCT_PRESETS)
    widths2 = [6, 14, 20, 13, 44] + [16] * len(preset_names)
    _title(ws2, "四個內建方案：同一組 16 種情況會得到甚麼建議", len(widths2))
    _note(
        ws2,
        "綠＝首選RF、黃＝次選RF、橙＝建議RF、白＝留空。各方案只是起點，"
        "可在「情況一覽」自訂後取代；應用程式側欄可即時切換方案或逐格修改。",
        len(widths2), 2,
    )
    headers2 = ["編號", "件數－店舖", "件數－同Nature", "件數對比", "情況說明"] + preset_names
    _header(ws2, 3, headers2, widths2)
    ws2.row_dimensions[3].height = 104
    for i, case in enumerate(_SHOP_SIMPLIFY_SCT_CASES, start=1):
        row = first_row + i - 1
        qty_cmp, _ = _shop_simplify_sct_case_compare(case)
        values = [
            i,
            "是" if case[0] == "Y" else "否",
            "是" if case[1] == "Y" else "否",
            qty_cmp,
            _shop_simplify_sct_case_text(case),
        ]
        for col, value in enumerate(values, start=1):
            cell = ws2.cell(row=row, column=col, value=value)
            cell.border = BORDER
            if col in (1, 2, 3, 4):
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col in (2, 3) and value == "是":
                cell.fill = _fill(GREEN)
            if col == 4:
                if value == "兩者皆是":
                    cell.fill = _fill(TEAL)
                elif value == "兩者皆非":
                    cell.fill = _fill(PINK)
        for offset, name in enumerate(preset_names):
            label = _SHOP_SIMPLIFY_SCT_PRESETS[name].get(case, "")
            cell = ws2.cell(row=row, column=6 + offset, value=label or None)
            cell.alignment = Alignment(horizontal="center")
            cell.border = BORDER
            if label == _SHOP_SIMPLIFY_SCT_PREFERRED:
                cell.fill = _fill(GREEN)
            elif label == _SHOP_SIMPLIFY_SCT_SECONDARY:
                cell.fill = _fill(YELLOW)
            elif label == _SHOP_SIMPLIFY_SCT_RF:
                cell.fill = _fill(ORANGE)
    ws2.freeze_panes = f"F{first_row}"

    # ── 工作表 3：方案說明 ──
    ws3 = wb.create_sheet("方案說明")
    widths3 = [34, 40, 34, 16, 46]
    _title(ws3, "四個內建方案的判定邏輯", len(widths3))
    _note(
        ws3,
        "件數＝Sold Qty (P90D)；金額＝Net Sales (MTD＋M-1＋M-2)。「件數 Top」指該 SKU 在"
        "本店或同 Shop Nature 的累計銷量達該分類 Top N%（預設 80%）。",
        len(widths3), 2,
    )
    _header(ws3, 3, ["方案", "首選RF 條件", "次選RF 條件", "其餘", "適用情境"], widths3)
    rows = [
        ("三級建議（預設）", "件數或金額任一為「兩者皆是」", "只中「僅同Nature」",
         "只中「僅店舖」→建議RF；四項皆非→留空",
         "按你指定的分配：兩者皆是＝首選、同Nature＝次選、只有本店＝建議RF、無Top 留空由 SCT 判斷"),
        ("件數優先", "件數－店舖 或 件數－同Nature 任一為「是」", "—", "留空",
         "以銷量為主、不想漏掉核心 SKU；最寬鬆，首選RF 數量最多"),
        ("雙重把關", "件數有 Top 且 金額有 Top", "只有件數 Top 或只有金額 Top", "留空",
         "既要賣得多、也要賣得貴；只中一邊的列入次選，留給 SCT 覆核"),
        ("店舖優先", "件數－店舖 為「是」", "件數－店舖非 Top，但 金額－店舖 為「是」", "留空",
         "以本店表現為準；同 Nature 只作參考，避免被其他店拉高"),
        ("命中項數分級", "四項中命中 3 項或以上", "命中 2 項", "留空",
         "想按「命中幾多項」分級，容忍單一維度差異"),
    ]
    for offset, values in enumerate(rows, start=4):
        for col, value in enumerate(values, start=1):
            cell = ws3.cell(row=offset, column=col, value=value)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws3.row_dimensions[offset].height = 40
    ws3.freeze_panes = "A4"

    start = 4 + len(rows) + 2
    ws3.cell(row=start, column=1, value="使用步驟").font = Font(bold=True)
    steps = [
        "1. 在「情況一覽」I 欄逐格選擇 SCT建議（留空＝不自動填寫）。",
        "2. 應用程式 → 店舖精簡SKU模式(指定店舖) →「📝 SCT建議」→ 勾選自動填寫 → 規則來源選「自訂 16 種情況」，照本表設定。",
        "3. 匯出後 Summary 會多一個「計數 - SKU：SCT建議」區塊，可直接核對首選RF／次選RF 的數量。",
        "4. 若只想先用現成方案，直接在側欄選「件數優先／雙重把關／店舖優先／命中項數分級」。",
    ]
    for offset, text in enumerate(steps, start=start + 1):
        ws3.merge_cells(start_row=offset, start_column=1, end_row=offset, end_column=len(widths3))
        cell = ws3.cell(row=offset, column=1, value=text)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws3.row_dimensions[offset].height = 30

    output = output_dir / f"SCT建議_16種情況設定表_{datetime.date.today():%Y%m%d}.xlsx"
    wb.save(output)
    return output


if __name__ == "__main__":
    path = build(ROOT)
    print(f"written {path.name} ({path.stat().st_size // 1024} KB)")
