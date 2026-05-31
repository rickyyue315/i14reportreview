# 版本記錄

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
