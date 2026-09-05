# 版本記錄

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
