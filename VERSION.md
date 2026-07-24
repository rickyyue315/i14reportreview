# 版本記錄

## v1.8.0 (2026-07-24)

- New: 簡化版 Dashboard 新增有賣VS有貨指標 — CM <50% / <30% SKU 數卡 + CM 完全無貨 SKU 卡 + 三期（M-2/M-1/CM）比率區間分佈表（Excel + Streamlit）
- New: `_build_sell_through_band_summary()` in transformer.py — 有賣VS有貨比率區間分佈表（<30%/30-50%/50-70%/≥70%/無貨，三期並排）
- Fix: New sheet `Low 有賣VS有貨 (CM<50%)` 標記口徑 — 無貨 SKU（有貨店舖=0）唔再當 0% 標 Y，只標有貨且比率<50%（實測 21Jul 數據 Y 由 692 → 651）
- Update: 簡化版 Excel + Streamlit Dashboard row3 擴充為 4 卡（低存貨 SKU 總數／<50%／<30%／完全無貨），Item Class 分類表下移
- Update: `generate_excel_stock_impact_simplified()` 新增 `sku_soh_summary` 參數，避免重複 enrich
- Update: `_write_stock_impact_dashboard_simplified()` 新增 `sku_soh_summary` 參數；分佈表放 row 15，Item Class 表移至分佈表之後

## v1.7.9 (2026-07-24)

- Update: 簡化版 Excel 對齊 Joseph 格式 — RAW 擴充至 38 欄（Total Stock Qty/有貨店舖/有賣店舖/Eff Price/Price Tier/SOH Tier）
- New: 簡化版新增「New」sheet（SKU 級 SOH Tier 分佈 + 有賣VS有貨比率），移除 NDRF Status 欄
- Update: 低存貨門檻預設 300 → 1000（完整版 + 簡化版，使用者仍可在 UI 調整）
- Update: Dashboard 低存貨卡文字改為動態顯示實際門檻
- New: `_enrich_raw_joseph()`, `_build_sku_soh_tier_summary()`, `_write_new_sheet()` in transformer.py

## v1.7.8 (2026-07-22)

- New: Stock Impact Dashboard & 簡化版 — 新增低存貨分析：🔴 低存貨 SKU 總數指標卡 + 按 Item Class 分類的低存貨 SKU 隻數表格（空白 Item Class 歸入「New SKU」）
- Excel 完整版: Dashboard 新增第四排指標卡 + Item Class 分類表（在 Top 5 Buyer 之後）
- Excel 簡化版: Dashboard 新增第三排指標卡 + Item Class 分類表
- Streamlit: 兩個 Dashboard 模式均同步顯示低存貨指標卡與 Item Class 分類表

## v1.7.7 (2026-07-21)

- 

## v1.7.6 (2026-07-11)

- 

## v1.7.5 (2026-07-10)

- 

## v1.7.4 (2026-07-09)

- New: NDRF 月份對比模式 — 對比兩個月份各店舖的 Carry SKU、RF SKU、ND With Stock SKU 數量
- New: `transform_ndrf_comparison()`, `generate_excel_ndrf_comparison()` in transformer.py
- New: 支援單一檔案含 Month 欄位，使用者選取兩個月份做對比
- New: 輸出 Excel — 月份對比 (By Shop)、Adj. SST 樞紐表 (x2)、SKU 明細、RAW
- New: SKU 明細包含 SKU Desc、Shop Name、Adj. SST 及兩月份 RP Type/Stock Qty 對照

## v1.7.3 (2026-07-09)

- 

## v1.7.2 (2026-07-08)

- 

## v1.7.0 (2026-07-07)

- Breaking: 取消 Site 層級缺貨影響分析 — Excel 及 UI 移除 By Site
- Simplify: Brand/Buyer/Category 彙總移除 4 個損失欄位（最大銷售損失(M-2)、最大銷售損失(P90D)、總銷售損失(M-2)、總銷售損失(P90D)）
- Simplify: SKU 跨 Nature 調貨建議移除 Tourist損失_M2、Tourist損失_P90D
- Simplify: by-SKU 缺貨診斷移除 Tourist損失_M2、Tourist損失_P90D、Mix損失_M2、Mix損失_P90D
- Fix: Dashboard 「Mix 可釋放總量」重複 → 改為「Tourist 平均Turnday」; 「總缺貨影響行數」→「Mix 平均Turnday」
- Update: Excel sheets 10→9; Streamlit tabs 7→6
- Update: `version.py` 1.6.2 → 1.7.0

## v1.6.2 (2026-07-07)

- Fix: Stock Impact Analysis 在 98%「調整格式中...」crash — 修復 progress bar 倒退問題
- Root cause: `_write_inventory_distribution_sheet` 內部 progress 推進到 0.93，但外層下一個步驟設為 0.82，導致進度倒退（97.2% → 92.8%），可能觸發 Streamlit 內部例外
- Fix: 將 `_write_inventory_distribution_sheet` 內部 progress 限制在 [0.79, 0.819] 範圍
- Fix: 將 `_write_sku_allocation_sheet` progress 從 0.78 修正為 0.76（不得超越外層 0.77）
- Fix: 將 `_write_sku_stock_diagnosis_sheet` 內部 progress 從 [0.74, 0.77] 移位至 [0.771, 0.789]
- Fix: 所有 `generate_excel_*` 函式在 `wb.save()` 前新增「正在儲存 Excel 檔案...」進度提示
- Fix: `_write_metric_cards` 的 progress 值改為可配置參數 `progress_pct`（預設 0.95）
- Update: `version.py` 1.6.1 → 1.6.2

## v1.6.1 (2026-07-07)

- Upgrade: Stock Impact Dashboard 從「缺貨焦點」升級為「存貨分佈分析」
- Why: 高銷售 Tourist 店無貨、Non Tourist 店卻大量囤積；Mix 店舖只需保留 7-14 日 Turnday 存貨。資料驗證 Mix `>30日` 過剩存貨 96,834 件，按 14 日標準可釋放 84,320 件；688 個 SKU 同時 Tourist 缺貨 + Non Tourist 過剩
- New: 每列新增 `Turnday (Days)`、`Mix Turnday Band`（<7日/7-14日/14-30日/>30日/無銷售有存貨）、`Mix Excess Stock`（Mix 可釋放存貨 = max(0, Shop Stock − P90D/30×14)）
- New: `_build_inventory_distribution_summary()` — 3 表存貨分佈彙總（Mix Turnday 分級 / Nature 存貨對比 / 存貨失衡 SKU）
- Upgrade: `_build_sku_allocation_summary()` 改用「過剩存貨」(Mix>14日 / Non Tourist>30日) 而非全部存貨；新增可調貨建議量 = min(Tourist 缺貨損失, 其他過剩存貨)
- Upgrade: Dashboard metric cards 重做為「存貨分佈焦點」（Tourist 缺貨 / Non Tourist 過剩 / Mix 可釋放 / 存貨失衡 SKU 數）
- New: Excel 新增「Inventory Distribution」工作表（Mix Turnday 分級 + Nature 對比 + 失衡 SKU），共 10 個工作表
- New: UI 新增「📊 存貨分佈」tab（優先顯示）
- Update: `version.py` 1.6.0 → 1.6.1

## v1.6.0 (2026-07-07)

- Breaking: **Sales Decline Detection 模式完全移除**，改為全新「Stock Impact Analysis」模式 — 缺貨銷售影響分析
- Why: 舊模式只比對 Sold Qty 三期時間序列，完全未使用任何 Stock 欄位，無法回答「為何衰退」。資料驗證證實 Tourist 店舖 Shop Stock (CM)=0 時，平均 Sold 從 5.58 降至 0.31（無貨就無銷售）
- New: `transform_stock_impact()`, `_add_stock_impact_columns()`, `_build_stock_impact_summary()`, `_build_sku_allocation_summary()`, `generate_excel_stock_impact()`, `_write_oos_detail_sheet()`, `_write_sku_allocation_sheet()`, `_write_stock_impact_dashboard()` in transformer.py
- New: 分析 Shop Stock Qty (M-2/M-1/CM) = 0 對 Sold Qty 的影響，判定缺貨模式（長期缺貨/間歇缺貨/恢復供貨/無缺貨）
- New: 缺貨銷售損失雙基準並列 — 以 Sold Qty (M-2) 為需求基準 vs 以 Shop Monthly Sold Qty (P90D) 90 天月均為基準
- New: Shop Nature 優先級 — Tourist > Mix > Non Tourist，所有彙總按 Tourist 缺貨嚴重度排序
- New: SKU 級跨 Nature 調貨建議 — Tourist 店舖缺貨但 Non-Tourist/Mix 仍有存貨的 SKU，揪出存貨配置不當
- New: 預設只分析 Warehouse SKU（Vendor 無 D001 倉庫概念），提供 checkbox 可關閉篩選
- New: 影響判定只計算 RP Type = RF 的店舖（排除 ND 店舖），符合「只計 RF 店舖」需求；RP Type 正規化（strip + upper）容錯髒資料
- New: Excel 輸出包含 RAW、OOS Impact Detail、By Site、By Brand、By Buyer、By Category、SKU Allocation、SKU Stock Diagnosis、Stock Impact Dashboard 共 9 個工作表
- New: SKU Stock Diagnosis 工作表 — by-SKU 缺貨診斷（Tourist + Mix 缺貨 SKU），含 D001 倉庫三期貨量 (M-2/M-1/CM)、店舖總存貨三期、D001 狀態分類（🟧 倉庫有貨但店舖缺貨 = 補貨問題｜⬜ 源頭斷貨 = 倉庫也斷貨）
- Remove: `transform_decline()`, `_add_decline_columns()`, `_build_decline_summary()`, `generate_excel_decline()`, `_write_decline_*()`, `DECLINE_CATEGORIES`, `_classify_severity`, `_DECLINE_WARNING_COLS` 全數移除
- Update: `version.py` 1.4.17 → 1.6.0

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
