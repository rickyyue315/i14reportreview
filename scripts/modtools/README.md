# modes 維護工具（modtools）

`transformer.py` 已拆分為 `modes/` 套件（每模式一檔＋`common`／`issues`／`framework` 共用核心，
`transformer.py` 僅保留相容墊片）。本目錄收錄拆分與日常維護工具，一律於倉庫根目錄執行。

## 健康檢查（唯讀，建議優先執行）

```bash
python scripts/modtools/check_modes.py
```

檢查引用缺口、模組循環引用、對外契約（`views`／`tests`／`scripts` 經 `transformer`
引入的名稱）、定義期求值、登記一致性；結束碼 0 為通過。

## 引用重接線

手動搬移函數或增刪跨模組引用後執行，自動重算全部 `modes/*.py` 的引用並分流：

```bash
python scripts/modtools/rewire.py         # 寫入
python scripts/modtools/rewire.py --check # 只報告
```

## 跨模組搬移

```bash
python scripts/modtools/relocate.py <名稱> <來源模組> <目標模組> [...]
```

例：`python scripts/modtools/relocate.py _write_guide_sheet safety_review common`。
搬運後自動更新雙方 `__all__` 與 `moved.json`，完成後執行 `rewire.py` 與完整測試。

## 整模式拆分（如 `transformer.py` 又長大）

```bash
python scripts/modtools/move_mode.py <mode_key>
```

`mode_key` 見 `move_mode.py` 內 `SPECS`（新增拆分對象時先加一筆，含種子函數與註冊資訊）。
搬完先跑 `check_modes.py`，再跑全測試，通過才繼續下一個。

## 標準流程

1. 修改／搬移程式碼。
2. `python scripts/modtools/check_modes.py`（必須 0 錯誤）。
3. `python -m pytest tests/ -q`（必須全過）。
4. 對外名稱有增減時，確認 `transformer` 墊片轉出正常。

## 依賴鐵律

* 模式模組只可引用 `modes.common`／`issues`／`framework`，以及已確立的單向邊
  （現況僅 `hybrid → stock_checking`、`forecast → stock_impact`）。
* 禁止新增模式間循環引用；`check_modes.py` 會攔截。
* 新公開名稱必須列入該模組 `__all__`，否則墊片不會轉出。
* `moved.json` 為名稱歸屬登記檔，由工具自動維護，勿手動編輯。
