# 版本記錄

## v1.5.1 (2026-07-06)

- New: Sales Decline Detection 模式 — 銷售趨勢衰退分析
- New: `transform_decline()`, `_add_decline_columns()`, `_build_decline_summary()`, `generate_excel_decline()`, `_write_decline_detail_sheet()`, `_write_decline_dashboard()` in transformer.py
- New: 比對 Sold Qty (M-2) → (M-1) → 預估 MTD 三期的銷售趨勢，判定連續衰退
- New: 分級警告 — 輕度 (10-20%) / 中度 (20-50%) / 重度 (≥50%)
- New: 多維度分析 — By Site / By Brand / By Buyer / By Category (Adj. MCH1)，以 Tab 切換
- New: Excel 輸出包含 RAW、Decline Warning Detail、By Site、By Brand、By Buyer、By Category、Sales Decline Dashboard 共 7 個工作表
- New: MTD 日期選擇與自動正規化（Sold Qty CM-MTD → 預估全月銷售）
- Update: RAW_COLUMN_ORDER_NAMES 新增 6 個衰退相關欄位

## v1.4.17 (2026-06-24)

- Fix: ND SKU Dashboard 的影響採購人員數/品牌總數/SKU總數改為只計 ND 受影響列（不再計入非 ND 列，避免重複或虛增）
- New: 新增「預計另有安排總數量」指標（Vendor 另有安排之庫存數量總和），與「預計可退倉總數量」並列

## v1.4.16 (2026-06-24)

- Update: SKU 精簡 Review 區分 Supply Source — 只有 Warehouse 貨可退倉，Vendor 貨標記「另有安排」(transfer=0)
- Update: 彙總表分類改為 Supply Source × Shop Stock Status 雙層（如 Warehouse - Overstock、Vendor - Normal）
- Update: By Shop 新增 Warehouse 退倉行數 / Vendor 另有安排行數 / Warehouse SKU 數量欄位
- Update: Dashboard 與 UI metric cards 新增 Warehouse 可退倉 / Vendor 另有安排行數指標

## v1.4.15 (2026-06-24)

- New: SKU 精簡 Review 模式 — 針對 RP Type 為 ND 的 SKU 建議全數退倉（不再限定 Overstock 狀態）
- New: `transform_nd()`、`_add_nd_recommendation_columns()`、`_build_nd_by_shop_summary()`、`generate_excel_nd()` in transformer.py
- New: 彙總表依 ND SKU 的實際 Shop Stock Status 動態分類（如 ND - Overstock / ND - Normal / ND - OOS）
- New: 輸出 RAW、Shop Stock Status、By Shop、ND SKU 數據分析摘要四張工作表
- 保留 RP Type=ND 但 Shop Stock Qty=0 的列（退倉量為 0），資訊透明不漏列

## v1.4.14 (2026-06-23)

- 

## v1.4.13 (2026-06-18)

- 

## v1.4.11 (2026-06-17)

- 

## v1.4.10 (2026-06-15)

- 

## v1.4.9 (2026-06-11)

- 

## v1.4.8 (2026-06-09)

- 

## v1.4.7 (2026-06-08)

- Rename: "Combined Review" mode renamed to "Safety Stock Review" across UI and docs
- Update: Sidebar「分析模式說明」與「核心計算規則」補充 Safety Stock Review 功能描述
- Update: README.md 同步更新模式名稱與說明 

## v1.4.6 (2026-06-04)

- 

## v1.4.5 (2026-06-03)

- 

## v1.4.4 (2026-06-02)

- 

## v1.4.3 (2026-06-01)

- 

## v1.4.2 (2026-05-31)

- 

## v1.4.1 (2026-05-29)

- 

## v1.5.0 (2026-05-28)

- New: Comprehensive README.md with full project documentation
- Update: Sidebar now reflects all three analysis modes (Overstock Review, OOS Review, Combined Review)
- Update: Sidebar adds OOS-specific required fields and calculation rules
- Update: Sidebar操作流程 extended to 6 steps covering mode selection and parameter tuning

- New: OOS Review mode — analyze OOS (0 Stock), Potential OOS (<7 Turnday), and Normal (Safety Qty Insufficient) conditions
- New: Mode selector (Overstock Review / OOS Review / Combined Review) at top of UI
- New: MTD date input for OOS projected monthly sold calculation
- New: OOS-specific output sheets (RAW, Shop Stock Status, By Shop, OOS庫存數據分析摘要)
- New: `transform_oos()`, `_add_oos_recommendation_columns()`, `_build_oos_by_shop_summary()`, `generate_excel_oos()` in transformer.py
- Chore: Session state separation between Overstock and OOS processing
- Chore: Combined Review mode placeholder with "Coming Soon" message

## v1.3.1 (2026-05-28)

- 

## v1.2.1 (2026-05-26)

- 

## v1.1.4 (2026-05-25)

- 

## v1.1.3 (2026-05-24)

- 

## v1.1.2 (2026-05-23)

- Fix: Streamlit deployment on Python 3.14 (use `>=` deps, update zeabur `python_version` to 3.14, upgrade Docker to 3.13-slim)
- Fix: File-switch upload does not reload data when same file name is re-uploaded
- Fix: SKU float-string conversion retains decimal place (e.g. `"123.0"`)
- Fix: Progress bar regression — `_cb_transform` no longer prints before transform
- Fix: Align zeabur `python_version` with Dockerfile base image
- Fix: numpy deprecation warning on `np.full` without explicit dtype
- Fix: Missing optional column warning only shown in dashboard, not on initial transform
- Fix: Empty filter result raises clear error instead of crashing
- Fix: SKU conversion crash on non-numeric values (safe int parsing)
- Fix: Empty file handling with clear error message
- Fix: IndexError when summary total row is missing (both app.py and transformer.py)
- Fix: Redundant session state condition in file loading
- Fix: Column prefix matching ambiguity (sorted + exact-space matching)
- Fix: Overlapping mask safety in transfer quantity calculation (explicit mutual exclusion)
- Fix: Transfer total consistency between dashboard and Excel (read from summary total row)
- Fix: Excel write performance with batched row appends
- Fix: Progress bar duplicate trigger at boundary
- Cleanup: Remove outdated i14-calculator/ directory
- Chore: Pin dependency versions in requirements.txt

## v1.0.1 (2026-05-22)

- 初始版本
- I14 報表計算工具上線
- 支援 SKUxShop 資料上傳、驗證、分類統計、樞紐分析與 Excel 下載
