# 版本記錄

## v1.37.0 (2026-09-08)

Safety Stock x NetStock Review 的分類引入 **Stock Checking** 與 **Safety Stock Review** 兩個模式（調貨與 Safety 建議計算一律不變，只加分類顯示）：

- **Stock Checking** — 參數區新增「🗂️ 分類維度」七選一（Shop Nature／RP Type／Shop Nature＋RP Type／OM／Shop Nature＋OM／銷量／銷量＋Shop Nature）與「📈 銷量存貨分類門檻」三項（缺貨線 7 日／積壓線 30 日／保底 2 件，與 hybrid 同一準則可調）。Summary 按所選維度分組（與 hybrid 同一 builder，nature 預設的 A–AC 契約欄序不變）；RAW 恆附「銷量帶」與「銷量存貨分類」顯示欄；「店舖問題 SKU」表末尾新增兩欄。調貨觸發、分配與數量**完全不隨維度改變**；存貨偏重提示門檻按維度切換（與 hybrid 同表）
- **Safety Stock Review** — 參數區新增同一組「🗂️ 分類維度」與門檻輸入。SUMMARY 工作表原 ①② 區塊下方新增 **③ Safety Qty Change 分類 × 分類維度交叉表**（列數／SKU 數／Safety Qty Change 合計，含總計列；總計列 SKU 數為全表不重複 SKU，與 ① 同一準則；非標準 Shop Nature 值如 Other 不佔分組欄、只計入合計）；RAW 恆附「銷量帶」與「銷量存貨分類」；「店舖問題 SKU」表末尾新增兩欄；結果頁新增「🧮 分類交叉表」tab（區塊 ③ 即時數據＋銷量 × 存貨分類分佈，後者母體為全部店舖×SKU 列）。Safety 建議值、倍率、MTD 月銷預估與退場建議**完全不變**；①② 表頭與 RAW 原有欄序不動
- **資料守衛** — 兩模式來源缺「Sold Qty (P90D)」時不輸出銷量兩欄並列 warning；缺「Shop Stock Qty (CM)」時只輸出銷量帶並列 warning；Safety Stock Review 缺「Shop Monthly Sold Qty (P90D)」時店舖周轉日數以 Sold Qty (P90D) ÷ 90 × 30.5 換算（與 Stock Checking 同一準則）並列 warning；OM 兩個維度缺「OM」欄或全欄空白時退回 Shop Nature（與 hybrid 同）
- New: `_attach_sales_classification_columns()`（兩模式共用的銷量帶／銷量存貨分類附加函數，hybrid 不經此函數）；`_build_combined_dimension_cross_tab()`（SUMMARY 區塊 ③，分組鍵與 hybrid 同一 `_hybrid_group_key_series`）；`_sales_threshold_inputs()`／`_DIMENSION_OPTIONS`（UI 共用）
- Change: `transform_stock_checking()` 新增 `dimension`／`sales_*` 參數，stats 記錄 `dimension`／`om_groups`／三項銷量門檻；`generate_excel_stock_checking()` 按 stats 維度寫 Summary；`transform_combined()` 新增 `sales_*` 參數；`generate_excel_combined()` 新增 `dimension`／`om_groups` 參數並寫區塊 ③；`_COMBINED_RAW_COLS` 於 Sold Qty (P90D) 後插入銷量兩欄
- Update: `tests/test_dimension_other_modes.py` 新增 14 項測試（預設契約不變、sales／om 維度分組、調貨計算零變動對照、店舖問題 SKU 兩欄、缺欄守衛、交叉表分組與總計、Excel 區塊 ③）；共 337 項測試通過

## v1.36.0 (2026-09-08)

Safety Stock x NetStock Review「積壓保底件數」預設值由 5 件下調為 **2 件**——低銷／零銷且庫存 2–4 件的店舖×SKU 列，預設門檻下即判為低銷積壓／零銷積壓（此前須達 5 件；庫存 0–1 件仍歸正常）。UI「🧱 積壓保底件數 (件)」預設顯示 2，仍可調 1–100；已明確輸入過參數的執行不受影響（僅無輸入時落到新預設）。**行為變更限銷量存貨分類及其統計**（分類欄、區塊⑤、銷量分類概況指標卡）——調貨 14 日線、Safety 基準天數及再平衡通道的守衛邏輯不變。

- Change: `_HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT` 5→2；`transform_hybrid()` 經 `_hybrid_sales_params()` 採用新預設（簽名不變，未傳值時落到新預設）；`build_hybrid_params()` 預設與 `hybrid_params_changed()` 舊快照補齊值同步 2；Excel「說明」表回退預設同步
- Change: 全站說明同步——README、使用說明頁（兩處）、結果頁側欄 caption 改「保底 2 件」
- Update: `tests/test_hybrid_sales.py` 參數快照相容測試改用新預設 2（`_PARAMS` 純函數測試夾具維持顯式 5 不變，屬固定輸入而非預設斷言）；其餘場景斷言不受影響；共 323 項測試與煙霧測試通過

## v1.35.1 (2026-09-08)

說明字眼優化為香港繁體書面語，避免大陸用語：Safety Stock x NetStock Review「說明」工作表全文、使用說明頁與 README 的相關描述、銷量帶分類說明圖——「回顯」→「列出／照錄」、「關注」→「留意」、「打開 Excel 預設停在此頁」→「開啟 Excel 時預設顯示本頁」、「速查提示」→「實用提示」、「邊界口徑」→「邊界準則」、「尾列／前列店舖」→「銷量最少／最高的店舖」、「由此頁入手」→「從這一頁著手」、「照此表」→「按此表進行」、「口徑」→「準則／定義」、「數據口徑」→「資料基準」、「運行」→「運作／執行」；另修正兩處既有文字（「逐列回顯 Safety 基準天數」→「逐列顯示」、「重點關注 Tourist 店舖」→「重點留意」）。僅文字變更，計算、欄位與版面零變動。

- Change: `_write_hybrid_guide_sheet()` 各區塊標題與導讀文字（③ 導讀標題改「先看哪一頁、應留意什麼」、欄首「建議留意」、④ 改「實用提示」）；`generate_excel_hybrid()` 與 `_write_hybrid_guide_sheet()` docstring 同步
- Change: `scripts/make_sales_band_chart.py` 判定條件表「說明」欄與備註用語，重新生成 `assets/sales_band_guide.png`
- Change: transformer.py／測試檔／圖表腳本的代碼註釋與 docstring 同步掃描（約 50 處「口徑／回顯／運行」等）——僅註釋文字，所有欄名、工作表名、警告訊息、UI 字串與計算邏輯零變動，323 項測試與煙霧測試全數通過驗證
- Change: README（9 工作表段落、「說明」列、Review 基準段）、使用說明頁（輸出 9 工作表描述、Review 基準、圖片 caption）與結果頁 Overstock 模式 caption（「重點留意」）
- Update: `tests/test_hybrid_mode.py` 說明表測試斷言同步新標題；共 323 項測試通過

## v1.35.0 (2026-09-08)

Safety Stock x NetStock Review 輸出 Excel 由 8 個工作表擴為 **9 個工作表**，新增首頁「說明」導讀表——列出本次執行參數（分類維度／調貨門檻／D001 門檻／Safety 基準周轉天數／雙向調貨／ND 優先轉出／銷量分類門檻／銷量再平衡／參考日期）、講解核心邏輯（調貨段 14 日線、Review 段 Safety 基準、銷量帶與銷量存貨分類、Past 90 Days 0 Sales 轉 ND），並逐頁說明各工作表的內容與用家應留意的部分（含實用提示）。開啟 Excel 時預設顯示「說明」頁。**非破壞性**——原 8 個工作表的欄位、數值與排序零變動，僅新增說明頁；使用說明頁另於本模式收合區塊加入「銷量帶分類說明圖」（高銷／中銷／低銷／零銷判定與例子，口徑與說明表一致）。

- New: `_write_hybrid_guide_sheet()`——「說明」工作表撰寫（① 本次執行參數列出：由 `stats` 還原分類維度／門檻／基準天數／開關狀態；② 核心邏輯五點；③ 工作表導讀八列（工作表／內容／建議留意）；④ 實用提示四點；含換行高度估算，全文字面語繁體中文）；`generate_excel_hybrid()` 接入（`wb.active = 0` 預設停於說明頁）
- New: 使用說明頁「🔀 Safety Stock x NetStock Review」收合區塊新增「🖼️ 銷量帶分類說明圖」子區塊（`assets/sales_band_guide.png`，由 `scripts/make_sales_band_chart.py` 生成，規則口徑與 `_hybrid_sales_bands()` 一致）；`MODES` 元組支援第 4 位可選圖片路徑
- Change: 全站說明同步——README／使用說明頁輸出工作表描述 8→9（含首頁「說明」）、Excel 輸出結構表加「說明」列
- Update: `tests/test_hybrid_mode.py` 三處 `sheetnames` 斷言前置「說明」，新增 `test_guide_sheet_first_with_params_echo`（說明表為首位且為開啟預設頁、參數按 stats 還原列出、導讀與實用提示區塊齊備）；共 323 項測試通過

## v1.34.0 (2026-09-08)

Safety Stock x NetStock Review 分類維度由六選一擴為**七選一**，新增「以銷量＋Shop Nature分類」（`sales_nature`）——Summary 與 Review 摘要交叉表改按 Nature×銷量帶 12 組（Tourist-高銷／Tourist-中銷／…／Non Tourist-零銷，Nature 外層固定順序、帶內層固定順序）分欄呈現。**非破壞性**——計算與預設完全相同（需求方 Tourist-RF，池級歸零與同 Nature 守衛照舊；TurnDay 池與 Safety 基準天數恆按 Shop Nature），銷量帶口徑與「以銷量分類」相同（常開產出欄，分帶結果隨本次上傳資料的篩選範圍而變）；舊六維度輸出欄位與數值零變動；缺 RP Type／OM 欄時不退回（銷量帶分組不依賴兩者）。

- New: 分類維度第七選項 `sales_nature`（以銷量＋Shop Nature分類）——`_HYBRID_DIMENSIONS`／`_HYBRID_DIMENSION_LABELS`（「銷量＋Shop Nature」）／新常數 `_HYBRID_SALES_NATURE_GROUPS`（3 Nature×4 帶＝12 組）；`_hybrid_dimension_groups()`／`_hybrid_group_key_series()`／`_build_stock_checking_summary()` 支持合併分組（每列鍵值＝Shop Nature-銷量帶；銷量帶欄缺失時顯式退回 Shop Nature 鍵，正常輸入不會觸發）；Summary 每組各一欄「由{Nature-帶} 店舖預計調出數量」；存貨偏重門檻 sales_nature 為 ≥12% 且差距 ≥8 個百分點（12 組，與 nature_om 同門檻）
- New: 結果頁參數區「🗂️ 分類維度」radio 加「🧮 以銷量＋Shop Nature分類」＋help 文字與維度 caption 說明；存貨偏重門檻完整對照 caption 加 sales_nature 門檻
- Change: 全站說明同步——README（分類維度七選一、銷量＋Shop Nature 維度規則、輸出工作表描述與門檻對照）、使用說明頁（分類維度 ①–⑦ 與輸出 8 工作表描述）
- Update: `tests/test_hybrid_sales.py` 新增 sales_nature 維度測試區（註冊與門檻、Summary／交叉表分組欄位與分組加總對帳、逐列分組鍵值、缺銷量帶欄退回 Shop Nature 鍵、計算與預設零變動、Excel Summary 群組標籤與 Review 摘要交叉表標題）；共 324 項測試通過

## v1.33.0 (2026-09-08)

Safety Stock x NetStock Review 新增「以銷量分類」：RAW／主表恆附**銷量帶**與**銷量存貨分類**兩個顯示欄（同 SKU 內按 Sold Qty (P90D) 平均排名分高銷／中銷／低銷／零銷四帶；以店舖周轉日數與缺貨線／積壓線／保底件數三條可調門檻分類高銷缺貨／低銷積壓／零銷積壓／正常），分類維度由五選一擴為**六選一**（新增「以銷量分類」），並新增預設關閉的「⚖️ 啟用銷量再平衡」開關——高銷缺貨的 RF 店舖列即使所屬 Nature 池不缺貨亦可加入調貨需求方、積壓列經既有轉出判定提供來源。**非破壞性**——再平衡開關未勾選時調貨欄位與數值與 v1.32.0 完全一致（僅新增兩個恆輸出的銷量顯示欄，主表 26→28 欄）；分類維度舊五選項計算零變動；銷量分類與調貨 14 日線及 Safety 基準天數互相獨立。

- New: 銷量分類常數與計算函數——`_HYBRID_SALES_BANDS`／`_HYBRID_SALES_CLASS_ORDER`、`_hybrid_sales_bands()`（同 SKU 有銷列平均排名分帶，有銷列 < 4 不產生高銷／低銷避免小樣本假區分，並列同值同帶）、`_hybrid_sales_classification()`（店舖 TurnDay = 庫存 ÷ 生效月銷 × 30.5；缺貨用「低於」、積壓用「不少於」；零銷有存未達保底歸正常，仍由 Past 90 Days 0 Sales 轉 ND 提示涵蓋）、`_hybrid_sales_params()`（缺貨線預設 7 日可調 1–14、積壓線預設 30 日可調 15–90、保底預設 5 件可調 1–100；非整數拒絕、範圍外箝制列 warning，風格同 `_hybrid_safety_days`）
- New: `transform_hybrid()` 新增 `sales_rebalance`／`sales_oos_days`／`sales_overstock_days`／`sales_overstock_min_units` 關鍵字參數（全部帶預設值，舊呼叫零影響）；銷量帶／銷量存貨分類兩欄恆寫入 RAW 並加入「調貨 Review Safety Stock」主表（`_HYBRID_DIAG_COLUMNS` 26→28 欄，置於 `Shop Monthly Sold Qty (P90D)` 之後）與 RAW 顯示欄序
- New: **銷量再平衡通道**（`sales_rebalance=True`）——需求方在原判定外**增列**高銷缺貨 RF 列：豁免所屬 Nature 池加總 TurnDay 門檻，仍須補到 14 日需求 > 0、D001 < 門檻、同 SKU 有可轉出量；跟隨單向／雙向框架（單向需求仍限 Tourist、轉出池仍限 Mix／Non Tourist）；供應方轉出量公式與池級歸零不變，需求擴大後健康池內積壓列經既有 donor 判定自動入事件表；自我調貨／同 Nature／同OM 守衛與 ND 不作需求方全部保留；stats 新增 `sales_rebalance`／三個門檻／`sales_oos_rows`／`sales_oos_skus`／`sales_overstock_rows`／`sales_overstock_units`／`sales_need_rows`（僅本通道新增需求列數），`_hybrid_stats_from_raw()` 以同口徑補算四個分類統計（開關狀態無法由 RAW 反推，回退路徑不列）
- New: 分類維度第六選項 `sales`（以銷量分類）——`_HYBRID_DIMENSIONS`／`_HYBRID_DIMENSION_LABELS`／`_hybrid_dimension_groups()`／`_hybrid_group_key_series()`／`_build_stock_checking_summary()` 支持銷量帶分組；存貨偏重門檻 sales 為 ≥30% 且差距 ≥15 個百分點；計算與預設完全相同（僅分組顯示不同）；新區塊 `_build_hybrid_sales_summary()`——「銷量 × 存貨分類交叉表」（母體為**全部店舖×SKU 列**，有別於 ①–④ 的事件列母體；各分類 × Shop Nature 列數、總計列數／SKU 數／店舖數、庫存與 P90D 銷量合計，含總計列）掛入「Review 摘要」工作表（⑤，恆寫入——無調貨事件時 ①–④ 顯示「無符合資料（①–④ 統計區塊以調貨事件列為母體）」而 ⑤ 照常輸出）與 Streamlit「Review 摘要」tab
- New: 指標卡末尾另加「銷量分類概況」一排（高銷缺貨列數／高銷缺貨 SKU 數／積壓列數／積壓庫存件數，僅 `sales_rebalance=True` 時，無第 5 排時該排為第 5 排；`_write_metric_cards()` 新增 `row6_metrics`，Review 摘要 table_start 改按排數動態計算 4×排數＋4，舊排版面零變動）；結果頁參數區新增「⚖️ 啟用銷量再平衡」核選（key `hy_sales_reb`）＋缺貨線／積壓線／保底件數三個數值輸入＋開啟時 caption 說明；分類維度 radio 加「📈 以銷量分類」；`build_hybrid_params()` 納入新開關與三參數，`hybrid_params_changed()` 舊快照補齊預設值（向後相容）
- Change: 全站說明同步——README（分類維度六選一、銷量分類與再平衡規則、輸出欄位描述）、使用說明頁（模式說明與核心計算規則）、結果頁模式說明 caption、存貨偏重門檻完整對照 caption 加 sales 門檻
- Update: 新增 `tests/test_hybrid_sales.py` ×25（銷量帶分帶規則與小樣本守衛、分類邊界與保底件數、參數正規化、再平衡通道單向／雙向／ND／池級歸零守衛、sales 維度 Summary／交叉表／門檻、區塊⑤版面與無事件時輸出、銷量分類概況指標卡排、params 向後相容）；`tests/test_hybrid_mode.py` 主表欄數斷言 25→27／26→28 更新、①–④ 空母體提示與區塊⑤恆輸出斷言更新；`tests/_smoke_pipeline.py` 加銷量再平衡＋sales 維度案例；共 317 項測試通過

## v1.32.0 (2026-09-07)

Safety Stock x NetStock Review OM 正規化：空白 OM 自成「空白」群組，OM 大小寫不敏感。RAW「OM」欄空白值改顯示「空白」（舊版顯示「(未填OM)」），大小寫變體合併為一組（顯示沿用首見原值）；舊維度計算零變動。

- New: 新增 `_HYBRID_OM_BLANK_LABEL`（「空白」）與 `_normalize_hybrid_om()`（去空格、空白成組、大小寫不敏感、顯示沿用首見原值）；同OM調貨比對改用小寫鍵（Ivy＝ivy 跨大小寫可互調）；`_hybrid_om_groups_from_values()` 等佔位字串統一改用常數
- Change: om／nature_om 退回條件改按來源是否全空判斷（缺欄或全欄空白才退回），真實 OM 名恰為「空白」時不再誤退回；部分空白只列一次歸組 warning（全欄空白時由退回 warning 涵蓋，不再重複）
- Fix: Streamlit 存貨偏重高亮改由欄位反推群組名（不再以正則解析提示文字，群組名含全形括號時仍可正確高亮，與 Excel 同步）
- Update: 新增第 22 節測試 ×4（正規化規則、空白成組不退回、大小寫合併與跨大小寫同OM調貨、保留字不誤退回）＋既有佔位斷言更新；README／使用說明頁同步

## v1.31.1 (2026-09-07)

Safety Stock x NetStock Review 參數區新增存貨偏重提示門檻說明（隨分類維度即時切換＋完整對照），避免日後忘記各維度門檻。

- New: 結果頁參數區分類維度下方新增「📦 存貨偏重提示門檻」caption——前半為目前所選維度的門檻（`_hybrid_skew_hint_text()`），後半為五個維度的完整對照；計算邏輯與輸出零變動

## v1.31.0 (2026-09-07)

Safety Stock x NetStock Review 分類維度由三選一擴為五選一，新增「以 OM 分類」（`om`）與「以 Shop Nature＋OM 分類」（`nature_om`）。**非破壞性**——兩個新維度計算與預設完全相同（需求方 Tourist-RF，池級歸零與同 Nature 守衛照舊；TurnDay 池與 Safety 基準天數恆按 Shop Nature），僅 Summary 與 Review 摘要交叉表改按 OM 或 Nature×OM 動態分組；舊維度輸出欄位與數值零變動。

- New: `_HYBRID_DIMENSIONS` 新增 `om`／`nature_om`，`_HYBRID_DIMENSION_LABELS` 新增 `om: OM`／`nature_om: Shop Nature＋OM`；新增 `_hybrid_om_groups_from_values()`（OM 欄去除前後空格、空白歸入「(未填OM)」、按字串排序）、`_hybrid_skew_thresholds()`／`_hybrid_skew_hint_text()`（存貨偏重門檻按維度切換：nature／rp／both 沿用 40%／20pp，om 用 25%／12pp，nature_om 用 12%／8pp）；`_hybrid_dimension_groups()` 新增 `om_groups` 參數（om 回傳動態名單，nature_om 回傳完整笛卡兒積 Nature 外層×OM 內層）
- New: `_hybrid_group_key_series()`、`_hybrid_diag_summary_columns()`、`_build_hybrid_diag_summary()`、`_stock_checking_summary_columns()`、`_build_stock_checking_summary()`、`_write_stock_checking_summary_sheet()` 新增 `om_groups` 參數（缺省時由 raw 即時推導；寫入層另由 Summary 欄位反推 `_hybrid_om_groups_from_summary_columns()`）；`transform_hybrid()` stats 新增 `om_groups`（排序後名單），`generate_excel_hybrid()` 由 stats 還原、缺失時由 raw 推導；`_hybrid_stats_from_raw()` 同步補算
- Change: 來源缺「RP Type」欄時僅 rp／both 維度退回 nature（om／nature_om 分組不依賴 RP Type，不退回）；來源缺「OM」欄或全欄空白時 om／nature_om 退回 nature 並列 warning；非法維度值維持退回 nature
- New: 結果頁參數區「🗂️ 分類維度」擴為五選一（預設維持 Shop Nature，新選項置後）＋ OM 維度 caption 說明；存貨偏重提示 caption 改按所選維度顯示門檻；`build_hybrid_params()` 無需改簽名（dimension 為字串透傳，過期偵測自動涵蓋新值）
- Update: 新增 `tests/test_hybrid_mode.py` 第 21 節 OM 維度測試 ×11（正規化與動態分組、OM 名單規則、門檻與提示文字、om／nature_om 與 nature 數值一致性、笛卡兒積欄位、缺 OM 欄與全空白退回、缺 RP 欄不退回、新門檻觸發、Excel 版面與回退重建）；共 288 項測試通過；README／使用說明頁同步

## v1.30.0 (2026-09-07)

Safety Stock x NetStock Review 新增「ND 優先轉出」開關（預設關閉）：同一 Shop Nature 池內，RP Type＝ND（不可補貨）的來源店先於 RF 店轉出。**非破壞性**——未勾選時輸出欄位與數值與 v1.29.0 完全一致。

- New: `_hybrid_allocate_transfers()` 新增 `nd_priority` 參數（bool 陣列，True＝該列為 ND 來源）——池內改依 ND 列先於 RF 列取用（池 dict 建立時以 `_nd_first()` 前置，dict 保序故內層取用即依此順序），池間順序仍為 Non Tourist → Mix → Tourist（雙向時才含 Tourist），展開為 NT-ND → NT-RF → Mix-ND → Mix-RF → Tourist-ND → Tourist-RF 六段；ND／RF 兩段各自維持來源檔列序（不按可轉出貨量重排）；三道守衛（自我調貨／需求 Nature 自身除外／同OM）與需求列「量大→小」排序全部不變
- New: `transform_hybrid()` 新增 `nd_priority: bool = False` 關鍵字參數（置於 `dimension` 之後，全部呼叫端為關鍵字傳故位置相容）；來源旗標 `_src_nd = ~is_rf`（ND 口徑＝正規化後非 RF，含無法辨識而歸入 ND 的列），同時套用於不限OM與同OM兩次分配；缺「RP Type」欄時開關自動失效並列 warning；轉出資格一律不放寬（池級歸零、逐列 14 日超額、零銷售有貨列全數可轉出均照舊），ND 列仍不作調貨需求方
- New: stats 新增 `nd_priority`（UI 請求狀態）、`nd_transfer`／`nd_transfer_om`（RP Type＝ND 來源列的轉出量合計）；`_hybrid_stats_from_raw()` 以同口徑補算兩個轉出量（`nd_priority` 無法由 RAW 反推，故回退路徑不列）
- New: `_hybrid_dashboard_metrics()` 第 5 排改為「轉出來源對照」——雙向時含 Tourist 轉出量兩卡，`nd_priority` 為 True 且 ND 有實際轉出時另加「ND 來源轉出量」／「ND 來源轉出量 (同OM)」兩卡（該排 2–4 卡）；排數上限仍為 5，`_write_metric_cards()` 與 `table_start`（20／24）邏輯零變動；ND 卡以開關狀態為前提，故 RAW 回退路徑不顯示（維持 `_hybrid_dashboard_metrics(raw, stats) == _hybrid_dashboard_metrics(raw, None)` 對現行預設情境成立）
- New: 結果頁參數區新增「🥇 ND 優先轉出」核選（預設關閉，key `hy_nd_priority`）＋開啟時的 caption 說明六段順序與「正常一店×一SKU 資料下池級合計不變、Safety Review 與建議調整量不變」；`build_hybrid_params()` 新增末位 `nd_priority` 參數並寫入快照，`hybrid_params_changed()` 舊快照補齊由只補 `ref_date` 改為同時補 `nd_priority`（避免 session 內誤報過期）
- Change（**設計要點**）: 因 `donor_mask` 只取決於「可轉出貨量 > 0」（轉出量恆 ≤ 可轉出量），轉出列母體與分配順序無關——本開關只改變 `轉出量`／`調貨來源`／`事件類型`／`調貨後庫存`／`調貨後 TurnDay`（含同OM變體）與店舖問題表的涉及來源店舖計數，**`Safety Review`、`建議 Safety 設定值`、`Safety 建議調整量`、①②③④ 統計與交叉表逐列不變**；各 Nature 轉出合計與調入總量在正常一店×一SKU 一列的資料下不變（若同一 Site Code 同時出現在需求與來源端——「Site Code × SKU」重複列異常，系統會列 warning——自我調貨守衛會令池內取用順序影響合計與缺口）
- Change: Stock Checking、Safety Stock Review 與 Overstock Review 輸出零變動（Stock Checking 的分配迴圈為獨立實作，且其輸入契約不要求 `RP Type`，本輪不加開關）；ND 庫存仍計入 Shop Nature 池的加總 TurnDay（有貨可調／存貨偏重判定維持現行口徑）
- Update: 新增 `tests/test_hybrid_mode.py` 第 20 節 ND 優先轉出測試 ×14（關閉＝現行契約、池內 ND 先、不跨池搶 ND、Safety 欄與池級合計不变、事件標籤與調貨後庫存變動、缺 RP Type 欄退化、正規化 ND 列入優先、rp 維度適用、指標卡第 5 排、Excel 版面第 19／23 列、`_hybrid_allocate_transfers` 池內與 Tourist 池單元測試）；`tests/test_analyzer_helpers.py` 新增 `nd_priority` 快照與舊向後相容測試 ×1；共 277 項測試通過；README／使用說明頁／結果頁 caption 同步

## v1.29.0 (2026-09-07)

Safety Stock x NetStock Review「Safety 設定過高」列的下調建議補底 3 件（SAP MOQ：Safety 1–2 件難以補貨，I14 原檔無 MOQ 欄）。

- New: `transform_hybrid()` 過高列改以**生效基準**判定與給建議——所需量 ≥ 1 且現有設定 ≥ 3 件者生效基準 = max(所需量, 3)（補底不高於現值：現有 3 件而所需 1–2 件者補底後相等，改判「設定正常—過剩非設定因素」，不出現調整量 0 的假動作列）；現有設定本身只有 1–2 件者維持原所需量（不強制上調到 3）；所需量 = 0（P90D 零銷售）的歸零建議屬退場，不受保底影響；「Safety 設定不足」列屬 TurnDay 驅動上調，一律照原所需量（含上調至 1–2 件）
- Change: 新增常數 `_HYBRID_SAFETY_MIN_KEEP = 3`；「Safety 周轉所需量」欄維持原所需量作稽核口徑，保底只體現在「建議 Safety 設定值／Safety 建議調整量」；stats（under/over/normal 列數與合計）、① 交叉表、② 調整方向分類、③ SKU 級調整明細、④ Buyer 統計、店舖問題 SKU 建議文字與指標卡說明均由 RAW 欄／新 mask 彙算自動跟隨；Safety Stock Review、Stock Checking 與 Overstock Review 輸出零變動（SSR 下調建議原已經 Overstock 上限 X 保底 3 件，維持現行「計算結果最低 3 件」口徑；Stock Checking 無 Safety 建議欄）
- Update: 新增 `tests/test_hybrid_mode.py` 第 19 節保底測試 ×6（補底至 3／補底後相等改判正常／現有 1–2 件保留原建議／上調不受保底／列集合與合計對帳／②③ 統計一致性）；`test_no_sale_donor_suggests_zero` 加註歸零不受保底；共 263 項測試通過；README／使用說明頁／結果頁 caption 同步

## v1.28.0 (2026-09-07)

「Adj. Safety Stock」於所有報表與頁面統一顯示為「Safety Stock (現時)」，避免誤會為系統建議值。

- Change: 新增名稱契約常數 `_ADJ_SAFETY_SOURCE_COL`／`_ADJ_SAFETY_DISPLAY_COL` 與 `_display_safety_header()`／`_standardize_safety_column()`（transformer 頂部）——來源檔案欄名 `Adj. Safety Stock` 為內部正規欄名（所有運算與 dataframe 讀取沿用），使用者可見輸出（Excel 表頭／Streamlit 表格／警告文字）一律顯示 `Safety Stock (現時)`；`Safety Stock (現時) ＝ 來源檔案的 Adj. Safety Stock` 語意固定
- New: 輸入別名相容——`_clean_dataframe()` 與 `_canonicalize_hybrid_columns()` 入口先正規化：上傳檔只有「Safety Stock (現時)」時自動改回來源名運算，兩名並存時保留來源名；`_resolve_column_name()` 支援雙名互認（缺欄報錯訊息亦用顯示名）
- Update: 表頭顯示名套用至 `_write_raw_sheet`（全部模式 RAW）、`_write_hybrid_diagnosis_sheet`（調貨 Review 主表）、`_write_shop_issue_sheet`（店舖問題 SKU）；`views/analyzer_page.py` 各 RAW／主表／店舖問題預覽以 `rename(columns=_display_safety_header)` 顯示；缺少可選欄位／空白非數值等 warning 改用顯示名
- Update: README／使用說明頁／側欄檢核標準同步標明「Safety Stock (現時) ＝ 來源檔案的 Adj. Safety Stock（並非系統建議值；系統建議值為『建議 Safety 設定值』欄）」；`tests/test_hybrid_mode.py` 空白 warning 斷言改查顯示名；共 257 項測試通過

## v1.27.0 (2026-09-05)

New SKU 判定改以 MTD 截止日期為基準（Safety Stock Review 與 Safety Stock x NetStock Review 同步）。

- Change: Safety Stock Review 沿用既有「📅 MTD 日期」作為新品保護參考日期（`transform_combined()`／`_add_combined_recommendation_columns()` 新增 `ref_date` 參數；預設今天，向後相容）；混合模式新增「📅 參考日期」輸入（預設昨日，`transform_hybrid()`／`_hybrid_p90d_zero_sales_mask()`／`_build_hybrid_shop_issues()` 新增 `ref_date` 參數，`stats["ref_date"]` 回顯，`build_hybrid_params()` 納入過期偵測；`generate_excel_hybrid()` 由 stats 還原）
- Update: 新增 `_normalize_ref_date()`（None 退回今天；date／datetime／Timestamp／字串皆可；非法值報錯）；新增測試 ×6（combined 邊界位移與非法值、hybrid 遮罩位移／stats 回顯／問題表、params 過期偵測與舊快照相容）；共 257 項測試通過；README／使用說明頁同步

## v1.26.0 (2026-09-05)

Safety Stock x NetStock Review 店舖問題表新增 Past 90 Days 0 Sales 轉 ND 提示。

- New: `_build_hybrid_shop_issues()` 另收零銷售轉 ND 候選（M-2／M-1／CM-MTD／P90D 皆零銷售且有庫存，LM 欄存在時須雙＞0；上市 90 天內新品、需要調貨列與 ND 列除外；重疊於 Safety 不足／過高列時保留原問題類型並加註，否則以「ND 退場」收錄，每店×SKU 最多一列）；`transform_hybrid()` 缺銷售欄與候選列數列 warning（不改 7 元組簽名與主表欄序）
- Update: 新增 `_hybrid_p90d_zero_sales_mask()`（NaN 視為 0、負數視為 ≤0，重用 `_parse_launch_dates` 新品保護）；新增 `tests/test_hybrid_p90d_nd.py`（×6：四欄口徑、LM／NaN／負數、缺欄 warning、新品／ND／需求排除、ND 單列、重疊加註不重複計列）；`tests/test_hybrid_mode.py` BASE 補 M-2／M-1 銷售欄（維持零 warning 口徑）；共 251 項測試通過；README／使用說明頁／結果頁 caption 同步

## v1.25.1 (2026-09-05)

店舖問題 SKU 顯示 Supply Source 與 D001 Net Stock Qty (CM)，缺口跟進按供貨來源分支。

- New: Stock Checking 與 Hybrid 店舖問題表新增「Supply Source」與「D001 Net Stock Qty (CM)」欄（置於店舖庫存之前）；D001 數量取與觸發判定一致的實際採用值（Hybrid 用逐列 Effective 值）
- Change: 調入缺口的跟進行動不再一律寫追蹤 D001 補貨 — Warehouse＋D001 有貨才由 D001 補貨；Warehouse＋D001 無貨寫 D001缺貨；Vendor 貨不經 D001 改為向供應商追貨（Vendor 由供應商直送店舖，D001 有無貨與其無關）
- Update: 新增 `_d001_available_qty()`／`_is_warehouse_supply()`／`_d001_gap_advice()` 共用分支；`tests/test_shop_issues.py` 新增 Warehouse／Vendor 分支測試（×1）；共 245 項測試通過；README／使用說明頁同步

## v1.25.0 (2026-09-05)

使用說明頁新增「圖解操作流程」（實際畫面擷圖：先選模式，再上傳檔案）。

- New: `views/usage_page.py` 新增「🖼️ 圖解操作流程」區塊（`st.image` ×2：步驟 1 選擇分析模式、步驟 2 拖曳 By SKUxShop Excel 至上傳區域），系統操作流程同步改為先選模式再上傳（與實際 UI 順序一致）
- New: 新增 `assets/guide/` 實際擷圖（`01_mode_select_annotated.png`／`02_upload_annotated.png` 紅框＋繁體中文標籤）與 `scripts/annotate_guide.py` 標註腳本（Pillow＋微軟正黑體，重跑即可更新標註）
- Update: 共 244 項測試通過；README／使用說明頁同步

## v1.24.1 (2026-09-05)

店舖問題 SKU 的建議方法不再重複列出調貨來源店號，只計涉及來源店舖數量（店號留給「調貨來源 (Site Code)」／「調貨來源 (同OM)」欄）。

- Change: Stock Checking 與 Hybrid 店舖問題表的建議方法改為「涉及 N 間來源店舖」，並指引用家查閱調貨來源欄；新增 `_count_transfer_sources()`（「HB30:2; HB49:3」計 2 間，空白計 0）
- Update: `tests/test_shop_issues.py` 新增去重複測試（×1：店號不出現在建議方法；Safety 表同步斷言）；共 244 項測試通過

## v1.24.0 (2026-09-05)

Stock Checking、Safety Stock Review、Safety Stock x NetStock Review 各新增「店舖問題 SKU」分析（每間店舖需行動列：問題＋建議方法）。

- New: 三模式各加一個「店舖問題 SKU」衍生表（只收需行動列，按店舖分組，店內按嚴重度排序；問題說明與建議方法為數量＋行動指引）——Stock Checking（需要調入／可轉出供應，缺口在同一列加註）、Safety Stock Review（Safety 需增加／減少、ND 退場；New SKU 標記不收錄）、Hybrid（只收 Safety 設定不足／過高，設定正常列排除；主表以不限OM為準，另附三個同OM對照欄）
- New: `transformer.py` 新增 `_build_stock_checking_shop_issues()`／`_build_combined_shop_issues()`／`_build_hybrid_shop_issues()` 與三組欄序常數、共用 `_write_shop_issue_sheet()`（表頭、凍結、篩選、隔列底紋與 Review 主表同式；空表寫「無符合資料」）；`generate_excel_stock_checking()`／`generate_excel_combined()`／`generate_excel_hybrid()` 各加一個「店舖問題 SKU」工作表（Hybrid 置於兩個 Review 主表之後、Review 摘要之前，8 表；另兩模式置於 By Shop／Summary 之後、分析摘要之前）；現有計算、固定表頭與欄序零變動
- New: Streamlit 三模式結果頁各加「🏪 店舖問題 SKU」分頁（即時由 raw 彙算，不新增持久化大表；空表顯示「無符合資料」）
- Update: 新增 `tests/test_shop_issues.py`（×10：只收行動列、數量對帳、按店排序、缺口加註、空輸入、Excel 工作表契約）；既有 Excel 契約同步（Stock Checking 4 表、Safety 6 表、Hybrid 8 表）；共 242 項測試通過；README／使用說明頁同步

## v1.23.0 (2026-09-05)

Summary 每個分類加入涉及店舖數量統計（三種分類維度同步，Stock Checking 亦同步）。

- Change（**破壞性**）: `_build_stock_checking_summary()` 每組加總區新增「涉及店舖數」欄（置於該組 TurnDay 之後；同 SKU 同組 Site Code 不重複計數，空白 Site Code 不計，與加總同母體）——nature 加 3 欄（29 欄，A–AC）、rp 加 2 欄、both 加 6 欄；`_stock_checking_summary_columns()`／`_STOCK_CHECKING_SUMMARY_COLUMNS`／`_write_stock_checking_summary_sheet()` 同步（加總區由每組 3 欄改為 4 欄，佔比／提示／調出欄位自動後移；TurnDay 與佔比活公式不受影響；雙向 Tourist 欄改為 AD）
- Update: Excel 第 5 列按現行格式輸出簡短欄名「涉及店舖數」（各組重複）；第 4 列各分組合併改為 4 欄；Streamlit 結果頁 Summary 標題與使用說明頁同步；README／使用說明頁同步
- Update: `tests/test_stock_checking.py` 同步（店舖數值、空白排除、29 欄契約、欄位位置、橘底欄號）＋ `tests/test_hybrid_mode.py` 同步（三維度欄位切換、Excel 欄號、等價對帳）；共 232 項測試通過；README／使用說明頁同步

## v1.22.0 (2026-09-05)

簡易調貨明細新增 Target／Supply RP Type 欄並調整欄序。

- Update: 欄序改為 12 欄 `Brand / SKU / SKU Desc / Target Site Code(接收店舖) / Target Shop Name / Target RP Type / Target OM / 調貨數量 / Supply Site Code(調貨來源) / Supply Shop Name / Supply RP Type / Supply OM`——RP Type 緊接對應 Shop Name 之後；調貨數量移到 Supply Site Code(調貨來源) 前，分隔 Target 與 Supply 兩組欄位（Streamlit tab 與 Excel 兩工作表同步）
- Update: `_build_hybrid_transfer_detail()` 新增 Site Code→RP Type 查找表（與 Shop Name／OM 同口徑，來源店查不到留空；接收店取自身列值），`_HYBRID_TRANSFER_DETAIL_COLUMNS` 契約同步
- Update: `tests/test_hybrid_mode.py` 同步（欄序契約、ND 來源店 RP Type 帶出、缺 RP Type 欄留空、Excel 表頭／自動篩選 A1:L／調貨數量第 8 欄）；README／使用說明頁同步

## v1.21.0 (2026-09-05)

Safety Stock x NetStock Review Buyer 合併表補回貨量與佔比欄。

- Fix: ④ Buyer 統計合併表補回「調入需求量合計」（各 Buyer 名下需調貨 SKU 的調入需求量列合計，取整，與指標卡總需求對帳）與「需調貨 SKU 佔比」／「需調整 Safety SKU 佔比」（0.0%，總計列 100%）；欄位為 Buyer／需調貨 SKU 數／需調貨 SKU 佔比／調入需求量合計／需調整 Safety SKU 數／需調整 Safety SKU 佔比／Safety 建議調整量合計＋總計列
- Update: 新增 `_hybrid_buyer_combined_detail(raw)`（同 SKU 全表首個非空 Buyer 歸屬，調入需求量與 Safety 調整量同步合計；舊 `_hybrid_buyer_combined_stats` 保留相容）；Excel 與 Streamlit 合併表同步七欄；Buyer 統計跟隨現行分類維度與調貨結果
- Update: `tests/test_hybrid_buyer_stats.py` 同步（×9，含貨量對帳與七欄位置／總計）；使用說明頁同步

## v1.20.0 (2026-09-05)

Safety Stock x NetStock Review 新增分類維度選擇（Shop Nature／RP Type／Shop Nature＋RP Type，預設 Shop Nature），並加入 ND 列不作調貨需求方規則。

- New: 結果頁參數區新增「🗂️ 分類維度」三選一（預設「以 Shop Nature 分類」）——決定調貨需求方過濾與 Summary／Review 摘要交叉表的分組欄位；**TurnDay 池與 Safety 基準周轉天數恆按 Shop Nature（Tourist／Mix／Non Tourist）設定，不隨維度改變**（RP Type 語意＝該店舖可否補貨：RF＝可補貨、ND＝不可補貨）
- Change（**破壞性**）: 調貨需求方一律排除 RP Type＝ND 的店舖列（ND 不作調入方，仍可作供應方轉出）——三種維度皆適用；預設 Shop Nature 模式下 Tourist-ND 缺貨列由「會標示需要調貨」改為「不標示」；RP Type 其他／空白值歸入 ND 並列 warning（保守視為不可補貨）
- New: 「以 RP Type 分類」維度——需求方擴展為任何 Nature 的 RF 列（Mix／Non Tourist 缺貨亦可調入），並取消「需求 Nature 自身除外」分配守衛與池級歸零（允許同 Nature 店間再平衡，可轉出貨量按逐列 14 日超額計）；供應方維持 Mix／Non Tourist（雙向時另加 Tourist）
- New: 「以 Shop Nature＋RP Type 分類」維度——調貨計算與預設模式完全相同（需求方 Tourist-RF），僅 Summary 與交叉表改按 6 組（Tourist-RF／Tourist-ND／Mix-RF／Mix-ND／Non Tourist-RF／Non Tourist-ND）分組呈現
- Change: Summary 欄位按維度切換分組（nature 三組＝現行 A–Z 契約欄序零變動；rp 兩組；both 六組），存貨偏重提示與橘底高亮同步支援新分組名稱；Review 摘要 ① 交叉表標題與欄位按維度切換（「Safety Review × RP Type 交叉表」／「× Shop Nature＋RP Type」）；「調貨 Review Safety Stock」主表新增 `RP Type` 欄（25→26 欄；來源缺「RP Type」欄時主表退回 25 欄）
- Change: 來源缺「RP Type」欄時全部列按 RF 處理（ND 排除規則無法套用）、維度退回 nature 並列 warning；非法維度值退回 nature 不拋錯
- Update: `transform_hybrid()` 新增 `dimension` 參數（位於 `progress_callback` 之前、以關鍵字傳入；回傳仍 7 元組，stats 新增 `dimension`／`mapped_rp_rows`）；新增 `_hybrid_dimension()`／`_hybrid_dimension_groups()`／`_normalize_hybrid_rp_type()`／`_stock_checking_summary_columns()` 與 `_HYBRID_DIMENSIONS`／`_HYBRID_DIMENSION_LABELS`／`_HYBRID_RP_GROUPS`／`_HYBRID_COMBINED_GROUPS` 常數；`_build_hybrid_diag_summary()`／`_build_stock_checking_summary()`／`_write_stock_checking_summary_sheet()` 新增 `dimension` 參數（預設 nature，Stock Checking 輸出零變動）；`generate_excel_hybrid()` 簽名不變（維度由 stats 推定）；`build_hybrid_params()` 納入維度並參與過期偵測
- Update: Streamlit——hybrid 區新增維度 radio（key `hy_dim`）與各維度觸發說明 caption、模式總說明與 Summary／交叉表分頁標題按維度顯示；`_style_inventory_skew()` 改為解析提示文字中的分組名稱（三維度通用）
- Update: tests/test_hybrid_mode.py 新增第 18 節 ×11（維度正規化、RP 正規化規則、預設模式 ND 不作需求／仍可供應、未知 RP 值 warning、缺 RP 欄退回、rp 同 Nature 再平衡與全 RF 需求、rp 維度 ND 永不需求、rp／both 分組欄位切換、both 計算對帳 nature、nature 契約欄序零變動、Excel 三維度欄位配置）；test_analyzer_helpers.py 新增維度偵測 ×1；共 230 項測試通過；README／使用說明頁同步

## v1.19.0 (2026-09-05)

Safety Stock x NetStock Review Buyer 統計合併為單一對照表。

- Change: ④ Buyer 需調貨 SKU 統計與 ⑤ Buyer 需調整 Safety SKU 統計合併為單一 ④ Buyer 統計（需調貨＋需調整 Safety）：同一 Buyer 的需調貨 SKU 數、需調整 Safety SKU 數與 Safety 建議調整量合計並列顯示（外連接，缺席以 0 補齊；按需調貨數降序、再按需調整數降序）
- Update: 新增 `_hybrid_buyer_combined_stats` 與 `_write_hybrid_buyer_combined_table`（欄位 Buyer／需調貨 SKU 數／需調整 Safety SKU 數／Safety 建議調整量合計＋總計列）；`transform_hybrid` stats 維持 `buyer_need_skus` 與 `buyer_safety_fix` 作為組合輸入；Excel Review 摘要由 ①②③④⑤ 五表改為 ①②③④ 四表；Streamlit 看板與分頁同步改為單表
- Update: `tests/test_hybrid_buyer_stats.py` 同步（×8，含外連接與合併表位置／總計）；使用說明頁同步

## v1.18.0 (2026-09-05)

Safety Stock x NetStock Review 新增 Buyer 統計與需要修改 Safety 統計。

- New: 沿用 Stock Checking 的 Buyer 需調貨 SKU 統計（同 SKU 首個非空 Buyer，已刪除重複，按數量降序），於 Hybrid 看板、Review 摘要分頁與 Excel Review 摘要同步顯示
- New: 新增 Buyer 需調整 Safety SKU 統計（母體為 Safety 設定不足＋過高，即有建議 Safety 設定值的列；同 SKU 首個非空 Buyer，已刪除重複；另含各 Buyer 的 Safety 建議調整量合計，淨額正=增加、負=減少），與需調貨統計並列顯示
- Update: `transform_hybrid` stats 新增 `buyer_need_skus` 與 `buyer_safety_fix`（`_hybrid_stats_from_raw` 回退同口徑）；`generate_excel_hybrid` 的 Review 摘要在既有 ①②③ 區塊後另加 ④⑤ Buyer 表（7 工作表數目與既有區塊列位置不變）；Streamlit Hybrid 看板與 Review 摘要分頁同步新增兩表
- Update: 新增 `tests/test_hybrid_buyer_stats.py`（×7：需調貨口徑等價、空白 Buyer 歸組、Safety 計數與金額對帳、空母體、回退一致、Excel 兩表位置與總計、空表不寫入）；共 216 項測試通過；使用說明頁同步

## v1.17.2 (2026-09-05)

Safety Stock x NetStock Review 代碼審查修正（文獻一致性、統計口徑、顯式守衛與稽核欄）。

- Fix: 「調貨段與 Stock Checking 完全同規則」宣稱改正為「同規則＋資料品質補正」——明列三處分岔（Shop Nature 正規化、月銷逐列回補、D001 逐列 Net→Stock fallback），同一份輸入在兩個模式的調貨結果可能有細微差異（本模式補正較嚴謹）；`transform_hybrid` docstring、Streamlit 模式說明與使用說明頁／README 同步
- Fix: 「Review 摘要」① Safety Review × Shop Nature 交叉表的「建議調整量合計」由 `int()` 截斷改為四捨五入（round-half-even），與 ② Safety 調整方向分類／stats 合計同口徑（小數 Safety 設定時不再出現 ±1 差異）
- Change: 「需求 Nature 自身除外」由池級歸零的隱式推論改為 `_hybrid_allocate_transfers` 的顯式守衛（新增選填 `nature_arr` 引數，預設 `None` 維持舊行為；`transform_hybrid` 不限OM／同OM 兩處呼叫傳入）；現行輸出零變化，屬防未來改動之防呆
- New: RAW 新增 `Effective Shop Monthly Sold Qty (P90D)` 稽核欄（與 D001 Effective 欄同口徑）——月銷回補後的實際採用值，覆寫前原始值保留於原欄；加入 RAW 顯示欄序「庫存與銷售」組（Effective Shop… 置於 D001 欄前）
- Change: `Site Code × SKU` 重複列 warning 補說明——重複列 Site Code 不同時，自我調貨守衛（以 Site Code 比對）無法辨識同店，同一實體店舖可能同時出現在調貨來源與接收端；`_build_hybrid_diagnosis` docstring 的「可轉出貨量」加總欄數改正為三 Nature（Tourist 欄僅雙向時存在）
- Update: tests/test_hybrid_mode.py 新增第 17 節回歸測試 ×4（Effective 欄數值與欄序、交叉表 ±1.5 四捨五入對帳、同 Nature 來源守衛直測與向後相容、重複列 warning 文案）；共 209 項測試通過，README／使用說明頁同步
## v1.17.1 (2026-09-04)

Safety Stock x NetStock Review 全面資料品質與結果一致性修正。

- Fix: hybrid 輸入不再要求未參與計算的 `Shop Stock Status`；店舖庫存、銷售、月銷及 Safety 設定拒絕非有限值，店舖庫存與 Safety 設定另拒絕負數，小數庫存仍拒絕；負數銷售按 0 處理並列 warning，避免錯誤釋放供應
- Fix: 月銷空白／非數值／負數，以及月銷為 0 但 P90D 有正數銷售時，按 P90D 換算月銷並列 warning，避免加總 TurnDay 變 `inf` 反而擋掉調貨；Summary D001 改顯示實際採用的有效值，與觸發判定一致
- Change: Safety 基準天數維持 7–10 建議範圍：越界數值箝制並列 warning；非整數浮點直接拒絕，不再靜默截斷；Nature 警告改為店舖層級描述，新增 OM 全欄空白 warning
- Fix: Streamlit 保存執行時參數 snapshot；設定變更後顯示 warning 並停用下載，結果標籤使用最後一次執行的基準；Site Code 選項與篩選統一字串 canonical 化
- Update: tests/test_hybrid_mode.py（`_SD14` 改合法 `_SD10` 並重算期望；新增第 16 節輸入邊界測試；Safety days 箝制／拒絕測試）、新增 tests/test_analyzer_helpers.py；README／使用說明頁同步

## v1.17.0 (2026-09-04)

Safety Stock x NetStock Review：**Safety 設定不再跟隨存貨（調貨）周轉天數**，改為按 Shop Nature 的可調基準天數。

- New: Safety Review 基準由固定 14 日改為**按 Nature 的「Safety 基準周轉天數」** — 結果頁左側新增三個輸入（🧳 Tourist／🔀 Mix／🏪 Non Tourist，範圍 7–10 日），預設 **Tourist 7 / Mix 9 / Non Tourist 10**；`Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 該列 Nature 的基準天數)`
- New: **調貨段完全不受影響** — `調入需求量 (補到14日)`、`實際調入量`、`調入缺口`、`轉出量`、`可轉出貨量 (Mix/Non Tourist/Tourist)`、`調貨後庫存`、`調貨後 TurnDay` 及其同OM變體一律沿用 14 日口徑（測試 `test_transfers_unchanged_by_safety_days` 逐欄比較預設與 14/14/14 兩種設定，要求完全相同）
- Change: **欄位改名（破壞性）** — `14日周轉所需量` → `Safety 周轉所需量`（數值已非固定 14 日）；影響 RAW、兩個「調貨 Review Safety Stock」主表、Review 摘要三個統計區塊與相關測試
- New: RAW 與主表新增 **`Safety 基準天數`** 欄（逐列回顯實際採用的 Nature 基準，供導出檔稽核，並讓 stats 回退路徑可還原基準）；主表由 24 欄改為 25 欄
- Update: `transform_hybrid()` 新增 `safety_days` 參數（**位於 `progress_callback` 之前、以關鍵字傳入**；回傳仍為 7 元組）；新增 `_hybrid_safety_days()`（未知 Nature／非數值／非正數一律退回預設，不拋錯）、`_hybrid_safety_days_label()`、`_hybrid_safety_days_from_raw()` 與 `_HYBRID_SAFETY_DAYS_DEFAULT`／`_MIN`／`_MAX`／`_HYBRID_SAFETY_LINE_COLUMN`／`_HYBRID_SAFETY_DAYS_COLUMN` 常數；`generate_excel_hybrid()` 簽名不變
- Update: 指標卡第 3 排說明與「Review 摘要」工作表標題改為顯示實際採用的基準（如「…｜Safety 基準 Tourist 7 日 / Mix 9 日 / Non Tourist 10 日」）；指標卡文字不再內嵌日數，維持 `stats` 有無時輸出一致
- Update: tests/test_hybrid_mode.py — 既有的 Review 數值測試改傳 `safety_days=_SD14`（三個 Nature 統一 14 日＝v1.16.0 口徑，數值斷言原樣保留）；欄名斷言全域改名；新增第 15 節 ×7（預設值與範圍、正規化無效輸入、Nature 基準逐列判定與 不足/過高/正常 反轉、**調貨欄位不受 safety_days 影響**、基準回顯與回退還原、主表 25 欄契約、Excel 標題含基準）；共 193 項測試通過
- Update: README／使用說明頁同步（模式說明、Review 基準規則第 3 點、輸出欄序與工作表結構表）

## v1.16.0 (2026-09-04)

- New: Safety Stock x NetStock Review 的「Review 摘要」新增 **② Safety 調整方向分類** — 以有調貨事件的店舖×SKU 列為母體（與 ① Safety Review × Shop Nature 交叉表相同），按「Safety 建議調整量」正負分 增加／減少／不變，各列 列數、SKU 數、Safety 建議調整量合計，末列「總計」（SKU 數不重複計算，合計為淨調整量）；四捨五入後為 0 的列歸「不變」；正常情況下三類列數即等於 stats 的 under_rows／over_rows／normal_rows
- New: 同時新增 **③ Safety 調整明細（SKU 級彙總）** — 每個有事件的 SKU 一列：`Brand / SKU / SKU Desc / Buyer / Review 列數 / 需增加店舖數 / 需減少店舖數 / 設定正常店舖數 / 現有 Safety 設定合計（調整列）/ 建議 Safety 設定合計（調整列）/ 淨調整量`；三個店舖數按 Site Code 不重複計算，兩個合計只累加「Safety 設定不足＋過高」列，淨調整量為該 SKU 各列「Safety 建議調整量」加總；只有「設定正常」列的 SKU 仍納入（淨調整量 0）；按 Brand / SKU 排序
- Update: 兩個統計**延伸在同一「Review 摘要」工作表**（① 交叉表下方，依序 ② ③，各區塊隔一空白列）與 Streamlit「📊 Review 摘要（Safety 調整統計）」分頁（原標籤「Review 摘要（Review × Shop Nature）」因內容已擴展而改名）；Excel 仍 7 個工作表、結果頁仍 6 個分頁，① 交叉表與指標卡的列位置零變動
- Update: Safety 建議不受 OM 限制與「允許雙向調貨」影響，故 ②③ 為 不限OM／同OM 共用的一份統計；`transform_hybrid()` 回傳與 `generate_excel_hybrid()` 簽名**皆不變**（②③ 由 RAW 即時彙算），新增 `_build_hybrid_safety_change_summary()`、`_build_hybrid_safety_sku_detail()`、`_hybrid_safety_event_frame()`、`_write_hybrid_summary_block()`（三個區塊共用樣式：合併標題、4472C4 表頭、總計列粗體灰底、數字 #,##0）與 `_HYBRID_SAFETY_CHANGE_*`／`_HYBRID_SAFETY_SKU_COLUMNS` 契約常數
- Update: tests/test_hybrid_mode.py 新增第 14 節「Safety 調整統計」測試 ×11（分類表與 stats 等價 ×2、明細欄序契約與分組、逐列加總對帳、店舖數按 Site Code 不重複、與 OM 限制無關、雙向母體對齊交叉表、空輸入、Excel 區塊位置（單向 ② 於第 27 列／雙向於第 31 列）、空母體僅寫「無符合資料」）；共 186 項測試通過
- Update: README／使用說明頁同步（模式說明、輸出規則第 12 點、Excel 輸出結構表）

## v1.15.0 (2026-09-04)

- New: Safety Stock x NetStock Review 新增「🧾 簡易調貨明細」— 結果頁新 tab（置於 Review 摘要與 RAW 之間，radio 切換**不限OM／同OM**），下載 Excel 同步新增「簡易調貨明細 (不限OM)／(同OM)」兩個工作表（共 7 表，附於最後，既有 5 表零變動）
- New: 簡易版一列一筆「來源店舖 → 接收店舖」調貨，欄位固定 `Brand / SKU / SKU Desc / Target Site Code(接收店舖) / Target Shop Name / Target OM / Supply Site Code(調貨來源) / Supply Shop Name / Supply OM / 調貨數量`；調貨來源字串（如「HB30:2; HB49:3」）按分號逐段拆行，同一接收列同一來源店重複出現時合併加總；只收錄「需要調貨=是」且實際有調入（實際調入量 > 0）的列，無貨可調的缺口列不進簡易版；來源店 Shop Name／OM 以 RAW 全量 Site Code 查找表補回（查不到留空）；按 Brand / SKU / 接收店 / 來源店排序；凍結表頭 + 自動篩選；無資料時顯示「無符合資料」
- Update: `transformer.py` 新增 `_HYBRID_TRANSFER_DETAIL_COLUMNS`、`_build_hybrid_transfer_detail(raw, variant)`、`_write_hybrid_transfer_detail_sheet()`；`generate_excel_hybrid()` 簽名不變、末尾追加兩表（progress 區間 0.90–0.95）
- Update: tests/test_hybrid_mode.py Excel 契約更新為 7 工作表 ×2，新增簡易版測試 ×7（拆行與 10 欄契約、同OM變體、缺口列排除、重複來源合併、空輸入、雙向來源、Excel 兩表格式）；共 175 項測試通過
- Update: README／使用說明頁同步（簡易調貨明細 tab 與兩工作表說明）

## v1.14.3 (2026-09-04)

- Change: Stock Checking 的 Summary 出現存貨偏重數值時，同步以**醒目橘底（FFC000）加粗體**加強顯示（與 Safety Stock x NetStock Review 同式）——Excel「Summary」工作表的「存貨偏重提示」欄（Z）與偏重 Nature 的庫存佔比欄（Q/S/U）底色；Streamlit「📊 Summary (SKU 彙總)」分頁以 `_style_inventory_skew()` Styler 同步高亮（`generate_excel_stock_checking()` 傳 `highlight_skew=True`）
- Change: Stock Checking 的 RAW 工作表優化顯示 — **不顯示「Item_Shop」與所有含「Value」的金額欄**（僅工作表層面移除，RAW dataframe 保留原欄供稽核；SKU Desc / Shop Name 仍由 Item_Shop 衍生顯示）
- New: RAW 欄序按可讀性重排 — 識別（Buyer / Brand / SKU / SKU Desc / Site Code / Shop Name / OM / Shop Nature）→ 品項屬性（Adj. SST / Vendor MC / Item Status / Item Class / RP Type / Shop Stock Status / Supply Source / Launch Date 等）→ 庫存與銷售（店舖庫存 / Safety Stock / P90D 銷售 / 月銷 / D001 三欄）→ SKU 級判定（各 Nature 加總 TurnDay / 需要調貨 / 調入需求量）→ 調貨執行（可轉出貨量 NT/Mix / 實際調入量 / 調入缺口 / 轉出量 / 調貨來源）→ 其他原始欄按原序；新增 `_stock_checking_raw_display_columns()`（transformer.py，Excel 與 Streamlit 共用）；Streamlit「RAW 完整明細」分頁改用共用函數（移除舊 `_base`/`_rest` 排列）
- Update: tests/test_stock_checking.py 偏重底色測試反轉斷言（Stock Checking 現在高亮）、新增欄序測試 ×2（helper 去重／移除斷言與分組順序、Excel RAW 表頭無 Item_Shop 與 Value 欄且分組順序正確）；README／使用說明頁同步；共 168 項測試通過

## v1.14.2 (2026-09-04)

- Change: Safety Stock x NetStock Review 的 RAW 工作表優化顯示 — **不顯示「Item_Shop」與所有含「Value」的金額欄**（僅工作表層面移除，RAW dataframe 保留原欄供稽核；SKU Desc / Shop Name 仍由 Item_Shop 衍生顯示）
- New: RAW 欄序按可讀性重排 — 識別（Buyer / Brand / SKU / SKU Desc / Site Code / Shop Name / OM / Shop Nature）→ 品項屬性（Adj. SST / Vendor MC / Item Status / Item Class / RP Type / Shop Stock Status / Supply Source / Launch Date 等）→ 庫存與銷售（店舖庫存 / Safety Stock / P90D 銷售 / 月銷 / D001 三欄）→ SKU 級判定（各 Nature 加總 TurnDay / 需要調貨 / 調入需求量）→ **Review 結果**（事件類型 / 14日周轉所需量 / Safety Review / 建議 Safety 設定值 / Safety 建議調整量）→ 調貨 (不限OM) → 調貨 (同OM) → 其他原始欄按原序；Streamlit「📄 RAW 完整明細」分頁與 Excel 同口徑
- Update: 新增 `_hybrid_raw_display_columns()`（transformer.py，Excel 與 Streamlit 共用）；`generate_excel_hybrid()` 的 `_write_raw_sheet()` 呼叫加 `columns=` 參數；analyzer_page.py RAW tab 改用共用函數（移除舊 `_hyb_cols`/`_base`/`_rest` 排列）
- Update: tests/test_hybrid_mode.py 新增欄序測試 ×2（helper 去重／移除斷言、Excel RAW 表頭無 Item_Shop 與 Value 欄且分組順序正確）；共 166 項測試通過

## v1.14.1 (2026-09-04)

- New: Safety Stock x NetStock Review 的 Summary 出現存貨偏重數值時，以**醒目橘底（FFC000）加粗體**加強顯示——Excel「Summary」工作表的「存貨偏重提示」欄（Z）與偏重 Nature 的庫存佔比欄（Q/S/U）同步底色；Streamlit「📋 Summary（SKU 存貨偏重）」分頁的偏重提示與偏重 Nature 庫存佔比儲存格同步橘底高亮
- Update: `_write_stock_checking_summary_sheet()` 加 `highlight_skew` 參數（預設 False，僅混合模式傳 True；Stock Checking 輸出零變動）；analyzer_page.py 新增 `_style_inventory_skew()` Styler
- Update: tests/test_hybrid_mode.py 新增偏重底色測試 ×1、test_stock_checking.py 新增預設不底色測試 ×1；共 164 項測試通過

## v1.14.0 (2026-09-04)

- New: Safety Stock x NetStock Review 新增「允許雙向調貨」選項（UI checkbox `hy_bidir`，預設關閉）— 針對「存貨偏重遊客區而其他區域賣得好出現缺貨」場景，調貨由單向（只向遊客區調入）改為**完全對稱三向**：任何 Nature 缺貨（該 Nature 加總 TurnDay < 門檻 且 D001 < 門檻 且個別店補到 14 日需求 > 0）且其他任一 Nature 有貨（加總 TurnDay ≥ 門檻）即標示需要調貨；來源優先序固定 **Non Tourist → Mix → Tourist**（需求 Nature 自身除外）；缺貨 Nature 的超額列一律不可轉出（池級歸零）；自我調貨與同OM守衛不變
- New: 雙向時 RAW 另加「可轉出貨量 (Tourist)」欄；Summary 末尾另加「由Tourist 店舖預計調出數量」欄（AA 欄，A–Z 契約欄序不動）；Review 事件對稱擴展——`需要調入` 可出現在任何 Nature 列、`實際轉出`／`可轉出未使用` 可出現在 Tourist 列（事件與診斷文字不變）；主表「可轉出貨量」含 Tourist；Review 摘要與 Streamlit 看板另加第 5 排指標卡（Tourist 轉出量／Tourist 轉出量 (同OM)）；stats 新增 tourist_transfer / tourist_transfer_om / bidirectional
- Change: 未勾選雙向時輸出欄位、數值與工作表結構與 v1.13.3 完全一致（零契約變動）；雙向只適用於本模式，**Stock Checking 維持單向**（兩模式調貨引擎自此分岔）
- Update: `transform_hybrid()` 加 `bidirectional` 參數；`_hybrid_allocate_transfers()` 加 `tourist_avail0` 來源池；`_build_stock_checking_summary()`／`_write_stock_checking_summary_sheet()` 加 `bidirectional` 參數；`_write_metric_cards()` 加 `row5_metrics`；`generate_excel_hybrid()` 以 RAW 是否含「可轉出貨量 (Tourist)」推定雙向
- Update: tests/test_hybrid_mode.py 新增雙向測試 ×8（逆向轉出、來源序、雙缺競爭、同店／同OM守衛、全池不足不標示、單向契約零變動、第 5 排指標卡與 stats 回退、Excel 27 欄契約）；_smoke_pipeline.py 加雙向一輪；共 162 項測試通過
- Update: README／使用說明頁同步（雙向規則、Stock Checking 單向限制註明）

## v1.13.3 (2026-09-04)

- Change: 全站字眼修正 — 「診斷」一律改為「Review」（「Review」的書面語對應並非診斷）；影響欄位名 `Safety 診斷` → `Safety Review`、工作表名 `調貨診斷 Safety Stock (不限OM)／(同OM)` → `調貨 Review Safety Stock (不限OM)／(同OM)`、`診斷摘要` → `Review 摘要`，以及 Streamlit 介面、進度訊息、README／使用說明頁同義文字
- Update: tests/test_hybrid_mode.py、test_stock_impact_pipeline.py、_smoke_pipeline.py 同步新欄位與工作表名稱；共 154 項測試通過（既有歷史版本記錄保留原用字）

## v1.13.2 (2026-09-04)

- Fix: 調貨分配自我調貨防護 — 來源店與需求店同為一店（Site Code 相同）時一律排除（`_hybrid_allocate_transfers` 與 Stock Checking 分配迴圈）；過去輸入檔若同一店×同一 SKU 出現重複列且 Shop Nature 不一致（如一行 Tourist、一行 Non Tourist），該店會同時成為調入方與調出方，甚至出現「調貨來源＝自己店舖」的自我調貨
- New: 調貨模式輸入一致性檢核（`_transfer_input_consistency_warnings()`，調貨診斷與 Stock Checking 共用，只列 warning 不拒絕）— ①「Site Code × SKU」重複列（組數與列數，正常輸入應為一店×一SKU一列）；②同一店舖的 Shop Nature 在不同列之間不一致（列出店號，最多 5 間）
- Update: tests/test_hybrid_mode.py 新增重複列與自我調貨回歸測試 ×3（warning 檢核、同 Nature 重複不误報 Nature 不一致、店級同時調入+調出＝0／調貨來源不含自己店號／守恆）；test_stock_checking.py 新增 ×2；共 154 項測試通過
- Update: 真實來源檔實測 — 正常輸入（22,732 列）修復前後調入／轉出合計完全相同（不限OM 9,433／同OM 4,784），新檢核無誤報；異常輸入合成案例（同店 Tourist＋Non Tourist 衝突列）修復後同店來源被排除、改列缺口並出現 warning
- Update: README／使用說明頁同步（分配規則與輸入檢核兩節）

## v1.13.1 (2026-09-04)

- New: Safety Stock x NetStock Review 引入 Stock Checking 的 SKU 級 Summary — 於混合模式輸出新增第 2 工作表「Summary」（與 Stock Checking 同格式同值：各 Nature 庫存/月銷加總、TurnDay、佔比、強制調貨提示、存貨偏重提示），供查閱哪些 SKU 存貨偏重哪個 Nature；Excel 由 4 工作表改為 5（RAW／Summary／調貨診斷 (不限OM)／調貨診斷 (同OM)／診斷摘要）
- New: Streamlit 混合模式新增「📋 Summary（SKU 存貨偏重）」分頁 — 先列存貨偏重 SKU 清單，再列全部 SKU Summary；session 鍵 hybrid_sc_summary
- Update: `transform_hybrid()` 回傳改 7 元組 (raw, sc_summary, diagnosis, diagnosis_om, diag_summary, stats, warnings)；`generate_excel_hybrid()` 第二參數為 sc_summary（None 時自動由 raw 建立，全部參數可省略回退）；README／使用說明頁同步
- Update: tests/test_hybrid_mode.py 解包改 7 元組、5 表契約、Summary 與 Stock Checking 等價斷言（含存貨偏重提示欄）；_smoke_pipeline.py 輸出 skew 計數；共 149 項測試通過

## v1.13.0 (2026-09-04)

- New: 調貨診斷 Safety Stock 分「不限OM」與「同OM」兩套 — 公司規定某些情況只允許同 OM 店舖調貨（例 Ivy HBA6 → Ivy HB94），同OM變體來源店須與需求店同 OM 才可轉出；分配抽為共用函數 `_hybrid_allocate_transfers()`（om_arr 參數），跑兩遍；同OM結果以「(同OM)」後綴欄存於 RAW（實際調入量 / 調入缺口 / 轉出量 / 調貨來源 / 調貨後庫存 / 調貨後 TurnDay），診斷主表輸出兩份工作表「調貨診斷 Safety Stock (不限OM)／(同OM)」，事件類型（實際轉出／可轉出未使用）按變體重算；**Safety 診斷與建議調整量兩變體完全相同（OM 限制不影響建議）**；OM 為選填欄位：缺失時同OM退化為不限OM＋warning（RAW／主表補「(未填OM)」欄），空白列歸入「(未填OM)」群組
- New: 診斷主表新增「OM」欄（Site Code 後，共 24 欄）— 同OM對照時可直接核對店群；OM 缺失時以「(未填OM)」補入
- New: 診斷摘要與 Streamlit 看板加第 4 排指標卡（同OM對照）：實際調入量 (同OM) / 無法足額缺口 (同OM) / Non Tourist 轉出量 (同OM) / Mix 轉出量 (同OM)（共 16 卡）；混合模式 Excel 由 3 工作表改為 4（RAW／不限OM 主表／同OM 主表／診斷摘要）；stats 新增 total_actual_om / total_gap_om / nt_transfer_om / mix_transfer_om / donated_rows_om；`transform_hybrid()` 回傳改 6 元組
- New: Stock Checking Summary 新增「存貨偏重提示」欄（Z 欄，25→26 欄）— 個別 Nature 庫存佔比 ≥ 40% 且（庫存佔比 − 銷量佔比）≥ 20 個百分點 →「存貨偏重 {Nature}（庫存 x% / 銷量 y%）」，識別存貨集中但銷量不高的 Nature；全 SKU 月銷 0 時銷量佔比以 0 計；多 Nature 符合以；分隔、差距大者在前；總庫存 0 不提示
- Update: README / 使用說明頁同步；真實來源檔（22,735 列 × 45 欄、275 SKU）實測：不限OM 實際調入 9,433／缺口 8,338 對比同OM 4,784／12,987，存貨偏重提示命中 14 個 SKU，Excel 生成約 12 秒
- Update: tests/test_hybrid_mode.py 新增同OM測試 ×6 與主表 OM 欄斷言；test_stock_checking.py 新增存貨偏重提示測試 ×4 並更新 26 欄 Excel 契約；共 149+ 項測試通過

## v1.13.0 (2026-09-04)

- New: Safety Stock x NetStock Review 調貨分配分「不限OM」與「同OM」兩變體 — 來源店須與需求店同 OM 才可轉出（公司規定某些情況只允許同 OM 店舖調貨，例 Ivy HBA6 → Ivy HB94）；同OM結果以「(同OM)」後綴欄存於 RAW（實際調入量 / 調入缺口 / 轉出量 / 調貨來源 / 調貨後庫存 / 調貨後 TurnDay），診斷主表輸出兩份工作表「調貨診斷 Safety Stock (不限OM)／(同OM)」，事件類型（實際轉出／可轉出未使用）按變體重算；Safety 診斷與建議調整量兩變體完全相同（OM 限制不影響建議）；OM 為選填欄位：缺失時同OM退化為不限OM＋warning，空白列歸入「(未填OM)」群組
- New: 診斷摘要加第 4 排指標卡（同OM對照）：實際調入量 (同OM) / 無法足額缺口 (同OM) / Non Tourist 轉出量 (同OM) / Mix 轉出量 (同OM)；Streamlit 看板同步 4 排 16 卡 + 診斷主表分兩頁籤；stats 新增 total_actual_om / total_gap_om / nt_transfer_om / mix_transfer_om / donated_rows_om；混合模式 Excel 由 3 工作表改為 4（RAW／不限OM 主表／同OM 主表／診斷摘要）
- New: Stock Checking Summary 新增「存貨偏重提示」欄（Z 欄，25→26 欄）— 個別 Nature 庫存佔比 ≥ 40% 且（庫存佔比 − 銷量佔比）≥ 20 個百分點 →「存貨偏重 {Nature}（庫存 x% / 銷量 y%）」；全 SKU 月銷 0 時銷量佔比以 0 計；多 Nature 符合以；分隔、差距大者在前；總庫存 0 不提示
- Update: 調貨分配抽為共用函數 `_hybrid_allocate_transfers()`（om_arr 參數控制同OM限制）；`_build_hybrid_diagnosis()` 加 variant 參數（free／om）；`transform_hybrid()` 回傳改 6 元組 (raw, diagnosis, diagnosis_om, diag_summary, stats, warnings)；`generate_excel_hybrid()` 改 4 表簽名
- Update: tests/test_hybrid_mode.py 新增同OM測試 ×6（來源切換、同OM缺口、診斷兩版一致、OM 缺失／空白、RAW 同OM欄、row4 指標卡對照）；test_stock_checking.py 新增存貨偏重提示測試 ×4 並更新 26 欄契約；共 149 項測試通過

## v1.12.0 (2026-09-04)

- Change: 「Safety Stock x NetStock Review」由「兩報告並排 + 反饋欄」完全轉換為「調貨事件診斷 Safety 設定」格式 — 以 14 日周轉線（14日周轉所需量 = ceil(P90D ÷ 90 × 14)，與調貨邏輯同口徑）診斷：Tourist 需調入列現有 Adj. Safety Stock < 所需量 → 「Safety 設定不足」；已標示調貨 SKU 的 Mix / Non Tourist 有貨列（轉出量 > 0 或可轉出貨量 > 0）現有設定 > 所需量 → 「Safety 設定過高」；等於邊界視為正常；建議 Safety 設定值 = 所需量、Safety 建議調整量 = 建議值 − 現有設定（正=增加、負=減少）
- Change: 移除混合模式內的安全庫存引擎段 — 不再計算 Recommend Safety Stock / X cap / Y floor；UI 刪除 MTD 日期、Overstock/OOS Rate、只可增加/可加可減參數（僅保留調貨兩門檻）；輸入檢核不再要求 Sold Qty (CM - MTD)；Adj. Safety Stock 空白或非數值診斷視為 0 並列 warning
- Change: Excel 輸出由 8 工作表精簡為 3 — RAW（全部欄位 + 診斷欄：14日周轉所需量 / 事件類型 / Safety 診斷 / 建議 Safety 設定值 / Safety 建議調整量 / 調貨後庫存 / 調貨後 TurnDay，非事件列留空）、調貨診斷 Safety Stock（店舖×SKU 主表，按 Brand / SKU / 事件類型 / Site Code 排序）、診斷摘要（指標卡 12 項 + Safety 診斷 × Shop Nature 交叉表含總計列）；刪除 Shop Stock Status / By Shop / SUMMARY / Safety 數據分析摘要 / 調貨 Summary / 調貨分析摘要 / Hybrid Summary
- Change: Streamlit 混合模式看板改為診斷指標卡（3 排 × 4：需調貨 SKU 數、需要調入列數、轉出來源列數、調入需求量、實際調入量、無法足額缺口、NT/Mix 轉出量、設定不足列數、建議增加合計、設定過高列數、建議減少合計）+ 3 個分頁（診斷主表 / 交叉表 / RAW）
- Update: `transform_hybrid()` 簽名精簡（移除 rates / mtd / safety_mode 參數）回傳 (raw, diagnosis, diag_summary, stats, warnings)；新增 `_build_hybrid_diagnosis()`、`_build_hybrid_diag_summary()`、`_hybrid_dashboard_metrics()`、`_hybrid_stats_from_raw()`、`_write_hybrid_diagnosis_sheet()`、`_write_hybrid_diag_summary_sheet()`；刪除 `_hybrid_build_safety_summary()`、`_build_hybrid_summary()`、`_write_hybrid_summary_sheet()`；`generate_excel_hybrid()` 改 3 表
- Update: tests/test_hybrid_mode.py 重寫為診斷契約（判定規則、事件類型、主表範圍、交叉表、3 表 Excel 契約等）；_smoke_pipeline.py 同步；README / 使用說明頁同步


## v1.11.1 (2026-09-04)

- New: Stock Checking Dashboard 加入「Buyer 需調貨 SKU 統計」— 統計各 Buyer 涉及多少個需調貨的去重 SKU（Buyer 取同 SKU 首個非空值、空白歸入「(未填 Buyer)」，按需調貨 SKU 數降序），Streamlit 指標卡下方顯示統計表（含佔比），Excel「調貨分析摘要」工作表指標卡下方同步輸出統計表（含總計列）；`transform_stock_checking` stats 新增 `buyer_need_skus`，呼叫端未帶 stats 時由 raw 自動匯總
- Change: Buyer 需調貨 SKU 統計「佔比」改顯示 1 位小數百分比（Streamlit `%.1f%%`、Excel `0.0%`）
- Change: 全站字眼優化（書面語繁體）—「去重」改為「已刪除重複」（Stock Checking / Hybrid / Past 90 Days 0 Sales 卡標題與描述、Excel 指標說明、使用說明頁）；「Non Tourist 鋪／Mix 鋪」改為「店舖」；「至少一邊」改為「至少一個 Nature」；新增 `AGENTS.md` 記錄語言規範（欄位名稱與 Excel 固定表頭屬資料契約，維持不動）

## v1.11.0 (2026-09-03)

- New: 改為 Streamlit 多頁架構（`st.navigation`）— 新增獨立「📖 使用說明」頁，側邊欄的「分析模式說明」與「核心計算規則」兩大段長文移至該頁，並按模式重組為 11 個收合區塊（每模式一塊，模式說明 🎯 + 核心計算規則 ⚙️ 合併呈現）；側邊欄僅留操作流程、欄位檢核與記憶體釋放按鈕，並加提示導向說明頁
- Change: `app.py` 改為導航入口（頁面設定 + 共用 CSS + `st.navigation`）；分析工具主體移至 `views/analyzer_page.py`，說明頁為 `views/usage_page.py`（內容逐字保留自原側邊欄，僅重新編排）
- Fix: 補齊全頁面 widget `key`（模式選擇、MTD 日期、Shop Nature、Safety 模式、Rate 滑桿、Site Code、月份對比、調貨門檻等 24 處）— 切換至說明頁再返回時，選擇與參數不會被重置；切頁不會清空已產生的分析結果

## v1.10.1 (2026-09-03)

- Fix: 混合模式入檔欄名統一 canonicalize（strip 前後空白，「偵測通過但取值 KeyError」不再發生；去空白後重複欄名明確報錯）
- Fix: D001 改為逐列有效值（`D001 Effective Stock Qty (CM)` = Net 優先、Stock 回補、空白視為 0，附 warning）；RAW 新增 Effective / D001 Stock 稽核欄位
- Fix: Shop Monthly Sold Qty (P90D) 單列空白改以 Sold Qty (P90D) ÷ 90 × 30.5 逐列換算（修復加總 TurnDay 誤判 inf 而不觸發調貨），並列 warning
- Fix: Shop Stock Qty (CM) 要求整數件數，小數值 fail-closed 明確報錯（避免 astype(int) 造成守恆失真）
- Fix: Hybrid UI Dashboard 補足 20 卡（Safety 8 + 調貨 8 + 反饋 4：含影響 Buyers/Brands、Overstock/OOS/調整行數、需要調貨列數、總明細行數）；「改善件數」改名「Safety 缺口改善量」並支援負值紅字顯示
- Fix: Excel 生成 progress bar 全程單調不倒退（`_write_stock_checking_summary_sheet` 區間參數化）
- Fix: SKU Desc 欄存在但全空白時以 Item_Shop 補正，摘要 groupby 不塌陷
- Fix: tests/_smoke_pipeline.py 可直接 `python tests/_smoke_pipeline.py` 執行（self 加入 repo root 至 path）
- Update: tests/test_hybrid_mode.py 新增回歸測試 ×8（header canonicalization、D001/monthly 逐列 fallback、整數檢核、SKU Desc 補正、progress 單調、RAW Effective 欄位）；共 131 項測試通過

## v1.10.0 (2026-09-03)

- New: 新增「Safety Stock x NetStock Review」混合模式 — 同一次執行 Safety Stock Review + Stock Checking，調貨結果只反饋為 Safety 診斷（不改寫 Recommend Safety Stock）
- New: `transform_hybrid()`、`generate_excel_hybrid()`、`validate_hybrid_input()`、`_normalize_hybrid_nature()`、`_hybrid_build_safety_summary()`、`_build_hybrid_summary()`、`_write_hybrid_summary_sheet()` in transformer.py
- New: 反饋欄位 — 調貨後庫存 = Shop Stock Qty (CM) + 實際調入量 − 轉出量；調貨前/後 Safety 缺口 = max(0, Recommend Safety Stock − 庫存)；調貨後 TurnDay = 調貨後庫存 ÷ Shop Monthly Sold Qty (P90D) × 30.5（月銷 0 留空）
- New: 參數合併兩模式 — MTD 日期、Overstock/OOS Rate（Tourist/Mix/Non Tourist）、Safety Qty 模式（只可增加/可加可減）、Tourist 加總 TurnDay 門檻（預設 14）、D001 門檻（預設 100）；調貨目標固定 14 日
- New: Shop Nature 固定三種；Other/空白/大小寫差異一律正規化為 Non Tourist（映射列數列入 warning）；混合必要欄位（基礎四欄 + Shop Nature / Site Code / Shop Stock Qty (CM) / Sold Qty (P90D) / Adj. Safety Stock / Shop Monthly Sold Qty (P90D) / Sold Qty (CM - MTD)）嚴格檢核；D001 沿用 Net→Stock fallback
- New: Excel 8 工作表 — RAW / Shop Stock Status / By Shop / SUMMARY / Safety Stock 數據分析摘要 / Summary（調貨）/ 調貨分析摘要 / Hybrid Summary（調貨反饋 Safety）
- Update: Streamlit 新增模式分支（指標卡 ×3 排 + 4 tabs + 下載，檔名 I14 Safety Stock x NetStock Review_<時間戳>.xlsx）；Sidebar 說明與核心計算規則同步
- Update: tests/test_hybrid_mode.py 新增（×20：輸入檢核、等價性、反饋公式、Nature 正規化、D001 fallback、Hybrid/Safety Summary、Excel 契約）；_smoke_pipeline.py 新增 hybrid 項目；共 123 項測試通過

## v1.9.6 (2026-09-03)

- New: Stock Checking Summary 在 SKU Desc 後新增 **Buyer、Adj. SST** 欄（同 SKU 取首個非空值；共 25 欄 A-Y）
- Change: Summary 除佔比 (%) 外一律取整 — 加總庫存/月銷 `#,##0`（月銷四捨五入）、TurnDay 改活公式 `=ROUND(加總庫存/加總月銷*30.5,0)`（`#,##0`，月銷 0 留空）；summary DataFrame 同步取整，佔比維持 0-1 小數（Excel `0.00%`）
- Style: Summary 可讀性優化 — 文字欄靠左、數值置中、間隔列淺灰底紋、表頭列高（3/4/5 列）+ 自動換行、合併儲存格完整邊框、空白 TurnDay 保留邊框網格；凍結表頭 + 自動篩選維持
- Update: tests/test_stock_checking.py 更新（TurnDay 取整、Buyer/Adj. SST 位置、佔比值、Excel 格式/公式/群組標題斷言），共 103 項測試通過

## v1.9.5 (2026-09-03)

- New: Stock Checking Summary 新增 **各 Nature 佔比 (%)**（6 欄）— 每個 SKU 按 Shop Nature 計算「加總 - Shop Stock Qty (CM)」佔比與「加總 - Shop Monthly Sold Qty (P90D)」佔比分母 = Tourist + Mix + Non Tourist 三者合計（合計 0 留空）；summary DataFrame 存 0-1 小數，Excel 以活公式 `=IF(合計=0,"",該Nature加總/合計)` + `0.00%` 格式顯示
- Change: Summary 工作表由 17 欄擴至 23 欄 — F-N（加總/TurnDay，欄位號不動）+ O-T（佔比，綠色群組標題「各 Nature 佔比 (%)」）+ U-W（強制調貨提示/調出量，由原 O-Q 右移）；Streamlit Summary 分頁同步顯示
- Update: tests 全數通過（103 項）

## v1.9.4 (2026-09-03)

- New: Stock Checking 輸出新增 **調貨分析摘要** 工作表（Dashboard，主題色指標卡，與 Streamlit Transfer Analysis Dashboard 同口徑）— 兩排共 8 卡：需要調貨列數 / 需調貨 SKU 數 / 調入需求量 / 實際調入量 / Non Tourist 轉出量 / Mix 轉出量 / 無法足額缺口 / 總明細行數；沿用 `_write_metric_cards()` 格局（其餘模式的「分析摘要」工作表同式）
- Change: `generate_excel_stock_checking()` 新增 `stats` 參數（由 `transform_stock_checking()` 回傳的統計直傳；缺省時由 RAW 匯總，舊呼叫相容）；工作表順序為 Summary → RAW → 調貨分析摘要
- Update: app.py 傳入 `_stats`、Sidebar 說明與核心計算規則同步；tests/test_stock_checking.py 更新 Excel 期望（×2：三表結構 + Dashboard 指標值校對），共 103 項測試通過

- Change: Stock Checking 新增供應端門檻（SKU 級）— 該 SKU 的 Mix 加總 TurnDay 與 Non Tourist 加總 TurnDay **皆低於門檻**（UI 可調，預設 14 日）時，理論上無貨可調，整個 SKU **無須標示「需要調貨」**；此前只檢查 Tourist 加總 TurnDay 與 D001
- Change: 單一 Nature 的 SKU 級加總 TurnDay < 門檻時，該 Nature 視為無貨可調 — 其店列「可轉出貨量」歸零、不參與調出分配（即使個別店有 14 日超額）；「無銷售但有存貨」的店不受影響（加總月銷 0 → TurnDay 無限 → 視為可釋放）
- New: RAW 新增「Mix 加總 TurnDay」「Non Tourist 加總 TurnDay」欄（SKU 級，與 Summary 同值），連同「Tourist 加總 TurnDay」方便核對標示與供應判定
- Update: Sidebar 說明、核心計算規則、README 同步；tests/test_stock_checking.py 新增供應門檻測試（×2：兩 Nature 皆低於門檻不標示 / 單一 Nature 低於門檻被歸零另一 Nature 正常調出），共 102 項測試通過

## v1.9.2 (2026-09-03)

- Change: Stock Checking 觸發條件改為 **SKU 級判定** — 以 Summary 同口徑的「Tourist 加總 TurnDay」（各 Tourist 店 Shop Stock Qty (CM) 加總 ÷ Shop Monthly Sold Qty (P90D) 加總 × 30.5）< 門檻（UI 可調，預設 14 日）且 D001 Net Stock Qty (CM) < 門檻（UI 可調，預設 100 件），**同時符合才標示「需要調貨」**；此前以每列 Turnday (P90D Qty) 個別判定，會出現「個別店 TurnDay 低但該 SKU Tourist 加總 TurnDay 達標（無須調貨）」卻被標示的情況
- Change: 加總 TurnDay ≥ 門檻的 SKU 一律無須調貨；只有已標示 SKU 才在 RAW 計算調入需求量 / 實際調入量 / 調入缺口與來源分配（個別 Tourist 店補到 14 日周轉的需求 > 0 才標示）
- New: RAW 新增「Tourist 加總 TurnDay」欄（SKU 級，與 Summary 同值），方便核對標示原因；缺「Shop Monthly Sold Qty (P90D)」欄位時以 Sold Qty (P90D) / 90 × 30.5 換算並發 warning
- Update: Sidebar 說明、核心計算規則、README 同步；tests/test_stock_checking.py 更新觸發條件測試並新增加總 TurnDay 判定測試（多 Tourist 店合併判定 / 個別店低但加總達標不標示），共 100 項測試通過

## v1.9.1 (2026-09-03)

- New: Stock Checking 輸出新增 **Summary** 工作表（參考 By SKUxShop_Stock Checking_Output 樞紐佈局，置於 RAW 之前）— 每 SKU 一列（按 Brand / SKU 排序）：Brand / SKU / SKU Desc / Supply Source / D001 Net Stock Qty (CM)；各 Shop Nature（Tourist / Mix / Non Tourist）加總 Shop Stock Qty (CM) 與 Shop Monthly Sold Qty (P90D)，TurnDay 以活公式 = 加總庫存 ÷ 加總月銷 × 30.5（月銷 0 時留空）；強制調貨提示（該 SKU 有 Tourist 店需要調貨時標示「需要調貨」）；由 Non Tourist / Mix 店舖預計調出數量 = 該 SKU 對應 Nature 來源店「轉出量」加總；凍結表頭 + 自動篩選
- New: `transform_stock_checking()` 改為回傳 `(raw, summary, stats, warnings)`；`_build_stock_checking_summary()`、`_write_stock_checking_summary_sheet()` in transformer.py；`generate_excel_stock_checking()` 新增 summary 參數
- New: Streamlit Stock Checking 結果新增「Summary (SKU 彙總)」分頁（Excel Summary 即時數據）
- Update: tests/test_stock_checking.py 新增 Summary 測試（×3：每 SKU 一列與 Nature 加總 / 提示與調出量彙總 / 排序），共 98 項測試通過
- Update: Sidebar 說明與 README 同步 Summary 輸出結構

## v1.9.0 (2026-09-03)

- New: 新增「Stock Checking（調貨檢查）」分析模式 — D001 倉庫少貨且遊客區（Tourist）店舖缺貨時的檢查：上傳 By SKUxShop_Stock Checking 檔案，計算 Tourist 鋪調貨需求並從 Mix / Non Tourist 鋪分配可轉出貨量（Non Tourist 優先）
- New: `transform_stock_checking()`、`generate_excel_stock_checking()`（單一 RAW 工作表）、`_stock_checking_new_columns()` in transformer.py
- New: 觸發條件 — Tourist 店 SKU 的 Turnday (P90D Qty) ≥ 0 且 < 門檻（UI 可調，預設 14 日）且 D001 Net Stock Qty (CM) < 門檻（UI 可調，預設 100 件）且需求 > 0 → 標示「需要調貨」
- New: 調入需求量 = 補到 14 日周轉（ceil(P90D / 90 × 14) − Shop Stock Qty (CM)）；來源可轉出量 = 14 日超額（max(0, 庫存 − P90D/90 × 14)，無銷售有存貨全數可釋放）；RP Type RF 與 ND 店皆可作為來源
- New: 分配規則 — 按 SKU 彙總需求，Non Tourist 店優先轉出，不足由 Mix 店補足；需求量大→小依序分配
- New: RAW 新增欄位 — 需要調貨 / 調入需求量 (補到14日) / 實際調入量 / 調入缺口 / 可轉出貨量 (Mix) / 可轉出貨量 (Non Tourist) / 轉出量 / 調貨來源 (Site Code)
- New: D001 欄位容錯 — 優先 D001 Net Stock Qty (CM)，退用 D001 Stock Qty (CM)，兩者皆缺視為 0 並發 warning
- New: Streamlit 新增模式選項、Turnday / D001 門檻參數、調貨分析摘要 dashboard、RAW 預覽與下載（檔名 By SKUxShop_Stock Checking_Output_<時間戳>.xlsx）
- New: 新增 tests/test_stock_checking.py（×13），涵蓋觸發條件、需求量、NT 優先 / Mix 補足分配、無銷售全釋放、ND 來源店、自訂門檻、D001 容錯與 Excel 輸出

## v1.8.17 (2026-09-02)

- New: Safety Stock Review 輸出 Excel 新增 **SUMMARY** 工作表（位於 By Shop 之後）— 兩個分類統計區塊：① Safety Qty Change 分類（增加 / 減少 / 不變）統計 行數 / SKU 數 / Safety Qty Change 合計；② Recommendation of new Safety Qty / ND 分類統計 行數 / SKU 數，固定順序涵蓋 (空白) 維持 Safety Stock、New SKU (3 Month)、0（已是 ND）、ND、ND(02 - MAKE UP)、ND(Private Label)、ND(Private Label - MAKE UP)，每類附中文說明；總計列 SKU 數採全表不重複 SKU（同一 SKU 可跨分類，避免加總虛增）；「只可增加」模式下減少類仍輸出但恆為 0
- Update: Sidebar 分析模式說明與核心計算規則、README 同步；新增回歸測試 Fix #15（×4），共 82 項測試通過

## v1.8.16 (2026-09-02)

- New: Safety Stock Review RAW 新增「Safety Qty Change 分類」欄（緊隨「Safety Qty Change」之後）— 按最終 Change 值分類：>0「增加」、<0「減少」、=0「不變」，方便直接篩選/樞紐；「只可增加」模式下不會出現「減少」；白名單增至 35 欄
- Update: Sidebar 核心計算規則同步；新增回歸測試 Fix #14（×3），共 78 項測試通過

## v1.8.15 (2026-09-02)

- Update: Safety Stock Review「自定義 Shop Nature Rate」設定區加註解說明，解決用家不知如何選值的問題 — 頂部說明兩組倍率的角色與換算直覺（Overstock Rate 30% ≈ 保留 9 日銷售量；OOS Rate 50% ≈ 覆蓋 15 日）與建議值公式；新增「如何選擇合適數值」指引（減庫存 → 調低 Overstock Rate 需配「可加可減」、減缺貨 → 調高 OOS Rate、按補貨週期長短分配各 Nature 的值），並標注「只可增加」模式下 Overstock Rate 無壓縮效果的陷阱；六個 slider 的 help 同步補上換算日數與設定傾向
- Update: Shop Nature 多選框說明補充 — Non Tourist 選項涵蓋「Other」類型（與倍率歸類一致）、部分選取會剔除其他類型、全取消會無資料

## v1.8.14 (2026-09-02)

- New: Safety Stock Review「Recommendation of new Safety Qty / ND」的 ND 轉換建議依類別加註 — Adj. MCH1 = 02 - MAKE UP → **ND(02 - MAKE UP)**；Adj. SST 為 03 - Private Label 系（(1)Major / (2)Health）→ **ND(Private Label)**；兩者皆是 → **ND(Private Label - MAKE UP)**；比對以 strip + 小寫正規化，Private Label 用「03 - Private Label」前綴比對（涵蓋未來新增子類）；僅加註「ND」轉換建議，「0」列（本來已是 ND）與 New SKU 標記列不受影響；Adj. MCH1 / Adj. SST 缺欄位時維持純「ND」
- Update: Safety Stock Review RAW 白名單加回「Adj. MCH1」欄位（緊隨 Adj. SST 之後，33 → 34 欄），方便 buyer 核對 MAKE UP 標示；新增回歸測試 Fix #13（×7），共 75 項測試通過

## v1.8.13 (2026-09-02)

- Change: Safety Stock Review「只可增加」模式語意收緊 — 全表不得出現任何扣減：ND 列與 P90D 0 Sales 舊品的「歸零 + ND」退場 override 改為只在「可加可減」模式套用；「只可增加」下這些列一律保持原有 Safety（Change = 0），ND 欄僅保留 New SKU (3 Month) 資訊標記。此前「只可增加」下一般列雖只增不減，但 ND / P90D 0 Sales 列仍會出現 Safety 歸零（Change = −Z）的扣減建議，與模式名稱矛盾
- Update: Sidebar 模式說明、核心計算規則與 README 同步；新增回歸測試 Fix #12（×4：ND / P90D0 / 新品 / 全表無負數），4 個退場 override 測試改為鎖定「可加可減」模式，共 68 項測試通過

## v1.8.12 (2026-09-02)

- New: Safety Stock Review RAW 新增「Recommend Safety Stock Turnday」欄位（緊隨「Recommend Safety Stock」之後）— 建議值的周轉日數 = Recommend Safety Stock ÷ (Shop Monthly Sold Qty (P90D) ÷ 30)，與來源「Turnday (P90D Qty)」同口徑（P90D 日均），用家可直接對比現時與建議的周轉日數（如建議後是否脫離 <7 Turnday 的 Potential OOS 區、落入 Mix 7-14 日標準帶）；P90D 月均 = 0 時留空（無周轉意義，來源為 inf），建議值 = 0 顯示 0（零覆蓋）；白名單增至 33 欄
- Update: Sidebar 核心計算規則與 README 同步；新增回歸測試 Fix #11（×5），共 64 項測試通過

## v1.8.11 (2026-09-02)

- New: Safety Stock Review 新品保護 — Launch Date 距今 90 天內（≈3 個月）的列排除「P90D 0 Sales → Safety=0/ND」override（新品未出單屬上架時間問題，過去會被建議歸零並標 ND 判死），改走正常 X/Y 計算，並在「Recommendation of new Safety Qty / ND」欄標記 **New SKU (3 Month)**；RP Type=ND 列仍維持 Safety=0 與「0」建議（ND 規則優先於標記）；Launch Date 缺欄位、缺值或無法解析時視為非新品（維持原行為）。該欄位在 Safety Stock Review 為純輸出（summary 與 Excel 均不讀取），標記不影響任何計算；日期解析抽成共用 `_parse_launch_dates()` 與 Stock Impact 的 `_add_launch_info()` 一致
- Update: Sidebar 核心計算規則與 README 同步新品保護說明；新增回歸測試 Fix #10（×5），共 59 項測試通過

## v1.8.10 (2026-09-02)

- Fix: Shop Nature 倍率對應改為精確比對 — 舊版 `str.contains("tourist")` 會把「Non Tourist」誤套 Tourist 倍率（Non Tourist 店 Overstock 上限被 30% 而非 50% 計算），同時 Shop Nature 篩選選「Non Tourist」時實際會排除 Non Tourist 列、誤留 Other 列；修正範圍含 Safety Stock Review（Overstock Cap / OOS Floor）、Overstock Review、OOS Review 與 `_clean_dataframe` 共 5 處，新增共用 `_nature_rate_arr()`
- Change: Safety Stock Review 的月銷預估（Projected Monthly Sold Qty）改為加權混合 — 舊版以 MTD 線性外推（MTD Sold ÷ MTD日數 × 當月日數），月初只有頭幾日數據時嚴重失真且 increase_only 模式下高估永不回落；新公式 = MTD 實績 + 剩餘日數 × P90D 月均（等價 w×MTD外推 + (1−w)×P90D，w = MTD日數/當月日數；P90D 已是月均直接使用），月初以穩定的 P90D 為主、月尾收斂至 MTD；P90D 欄位缺失或個別列缺值時退回純 MTD 線性外推（舊行為）。OOS Review 不受影響（維持 mtd / p90d 手動切換）
- Update: Safety Stock Review RAW 欄序調整 —「Recommendation of new Safety Qty / ND」移至「Safety Qty Change」與「Buyer reply」之間（原本在「Projected Monthly Sold Qty」之後）
- Update: Sidebar 設定、核心計算規則與 README 同步新公式說明；新增回歸測試 Fix #8（倍率精確比對 ×4）與 Fix #9（混合預估 ×4），共 54 項測試通過

## v1.8.9 (2026-09-01)

- Change: Safety Stock Review RAW 改為固定 32 欄白名單 — Brand / SKU / SKU Desc / Buyer / Adj. SST / Vendor MC / Item Status / Launch Date / Item Class / Site Code / Shop Name / OM / Shop Nature / RP Type / Safety Stock (現時) / Recommend Safety Stock / Safety Qty Change / Buyer reply / Shop Monthly Sold Qty (P90D) / Safety VS Monthly Sold Qty / Projected Monthly Sold Qty / Recommendation of new Safety Qty / ND / Overstock Cap (X) / OOS Floor (Y) / D001 Net Stock Qty (CM) / Shop Stock Qty (CM) / Sold Qty (P90D) / Turnday (P90D Qty) / Sold Qty (M-2) / Sold Qty (M-1) / Sold Qty (CM - MTD) / Supply Source；移除 Item_Shop、MCH2、Net Sales、Excess、PO 等欄位

## v1.8.7 (2026-08-28)

- Change: Past 90 Days 0 Sales 篩選定義新增條件 — Sold Qty (P90D) > 0 的記錄一律排除，避免「過去 90 天其實有銷售」造成混淆；無此欄位的輸入檔視同 0（維持原行為）

## v1.8.6 (2026-08-28)

- Change: Past 90 Days 0 Sales 的 RAW 改為固定欄位白名單（23 欄）— Site Code / Shop Name / OM / Shop Nature / Brand / SKU / SKU Desc / Vendor MC / Item Status / Launch Date / Item Class / RP Type / Shop Stock Status / Safety Stock (現時) / Sold Qty (M-2) / Sold Qty (M-1) / Sold Qty (CM - MTD) / Sold Qty (P90D) / Shop Stock Qty (CM) / Shop Stock Qty (LM) / Excess Stock Qty / Shop On Order PO Qty (CM) / Buyer；輸入缺少選填欄位時自動略過不報錯
- Update: 首頁模式說明同步新欄位定義

## v1.8.5 (2026-08-28)

- Fix: 「🧹 釋放記憶體」按鈕修復 — 過去刪除 session_state 後檔案仍掛在 file_uploader 上，st.rerun() 時會即時重新讀取整個 Excel，導致每按一次記憶體不減反增；現以更換 uploader widget key 徹底釋放已上傳檔案
- Fix: 按鈕新增 `st.cache_data.clear()` — transformer 的 `_compute_raw_extra_columns` 模組級緩存過去無 TTL／無 max_entries（無限累積 DataFrame 拷貝）且按鈕從未清除
- Fix: `_release_os_memory()` — gc 後呼叫 `malloc_trim` 將已釋放記憶體真正歸還作業系統（單靠 gc.collect() RSS 不會下降），令 Memory Usage 顯示如實反映
- Update: transformer.py 緩存加上 `max_entries=4, ttl=3600`，限制長時間運行下的記憶體增長

## v1.8.4 (2026-08-28)

- New: Zeabur 執行優化 — Dockerfile 移除 curl/apt 層，健康檢查改用 Python 內建模組（映像更細、構建更快）
- New: 生產環境運行旗標 — `--server.fileWatcherType=none`（節省記憶體）、`--server.enableWebsocketCompression=false`（降低大數據表傳輸 CPU）、`--browser.gatherUsageStats=false`（停用遙測）
- New: `PYTHONUNBUFFERED=1` — 日誌即時輸出到 Zeabur Logs
- New: requirements.txt 固定版本（streamlit==1.51.0 / pandas==2.3.3 / numpy==2.3.4 / openpyxl==3.1.5），確保 Zeabur 構建可重現
- New: Sidebar 底部新增 Memory Usage 顯示（讀取 /proc/self/status，僅 Linux 容器生效），方便監察 Zeabur 容器資源
- Update: zeabur.yaml 運行命令與 Dockerfile 同步；README 擴充 Zeabur 部署機制與容量建議（≥1GB RAM）

## v1.8.3 (2026-08-28)

- Update: Past 90 Days 0 Sales 定義更新 — Sold Qty (M-2) / (M-1) / (CM - MTD) 三期皆 ≤ 0（0 件或少於 0 件）且 Shop Stock Qty (LM) 與 (CM) 皆 > 0
- Remove: Past 90 Days 0 Sales 模式移除 Excel「Shop SKU」工作表（Site 下拉篩選頁無用）及隱藏的「_SiteList」頁 — 現在只輸出 RAW、Summary by Shop 兩頁
- Update: Streamlit 對應移除 Shop SKU 分頁，Sidebar 說明與摘要卡文字同步新定義

## v1.8.2 (2026-08-27)

- Fix: Past 90 Days 0 Sales 模式 Excel Shop SKU 頁開啟時被 Excel 修復移除公式 — FILTER 動態陣列公式加上 `_xlfn.` 前綴（Excel 檔格式中 2007 後新增函數必需），避免開啟時被當作未定義 UDF 而移除

## v1.8.1 (2026-08-26)

- New: 新增「Past 90 Days 0 Sales」分析模式 — 篩選「過去 90 日零銷售且店舖仍有存貨（Shop Stock Qty LM 與 CM 皆 > 0）」之 SKU×Shop 記錄
- New: `transform_p90d_zero_sales()`, `_build_p90d_zero_sales_by_shop_summary()`, `_p90d_raw_columns()`, `generate_excel_p90d_zero_sales()` in transformer.py
- New: 輸出 3 個工作表 — RAW（移除所有 Value 及 D001 欄位，保留 SKU Desc/Shop Name）、Summary by Shop（各店去重 SKU 數 + LM/CM 店舖庫存總數 + 總計列）、Shop SKU（頂部 Site 下拉篩選，以 FILTER 動態陣列隨選擇更新）
- New: Streamlit 新增模式選項、三頁結果表與下載按鈕、Sidebar 說明

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
