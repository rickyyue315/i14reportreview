# I14 報表轉換與分析系統

自動化 I14 庫存報表轉換與多維度數據分析工具，支援 **Overstock Review**（庫存過賸審查）、**OOS Review**（缺貨審查）、**Carry SKU Review**（在售品項盤點）、**SKU 精簡 Review**（ND SKU 退倉建議）、**Safety Stock Review**（安全庫存審查）、**Safety Stock x NetStock Review**（調貨事件 Review Safety 設定）、**Safety Stock 月度變化分析**（每月 Safety 變化對生意影響）與 **Forecast Review**（SKU 級建議訂貨量）等多種模式。

## 功能概述

| 功能 | 說明 |
|------|------|
| 檔案上傳與驗證 | 上傳 By SKUxShop Excel (.xlsx)，自動檢核必要欄位 |
| 模式選擇 | Overstock Review / OOS Review / Carry SKU Review / SKU 精簡 Review / NDRF 月份對比 / Safety Stock Review / Stock Impact Analysis / Stock Impact Analysis（簡化）/ Past 90 Days 0 Sales / Stock Checking / Safety Stock x NetStock Review / Safety Stock 月度變化分析 / Forecast Review |
| Overstock 分析 | 篩選 P90D 零銷售及超過月均兩倍之庫存過賸品項 |
| OOS 分析 | 篩選缺貨（0 Stock）、潛在缺貨（Turnday < 7）及 Safety Qty 不足 |
| SKU 精簡分析 | 針對 RP Type = ND 的 SKU 建議退倉（Warehouse 可退倉、Vendor 另有安排），依 Supply Source × Shop Stock Status 雙層分類 |
| Stock Checking 分析 | D001 倉庫少貨且 Tourist 店舖 Turnday < 門檻 → 標示調貨需求，從 Mix / Non Tourist 店舖分配可轉出貨量（Non Tourist 優先）；Summary 另以佔比差額標示「存貨偏重」的 Nature（存貨集中但銷量不高，偏重數值以醒目橘底加強顯示，偏低數值以淺藍底標示）；RAW 按可讀性重排欄序，不顯示 Item_Shop 與 Value 金額欄；另附「店舖問題 SKU」（每間店舖需行動列：需要調入／可轉出供應，問題＋建議方法；建議方法只計涉及來源店舖數量，不重複列出店號）；可選擇**分類維度**（七選一，與 Safety Stock x NetStock Review 同一組選項；只改 Summary 分組，調貨計算不變）；RAW／店舖問題 SKU 恆附「銷量帶」與「銷量存貨分類」顯示欄 |
| Safety Stock x NetStock | 沿用 Stock Checking 調貨引擎，以 **Safety 基準周轉天數（按 Nature，預設 Tourist 7 / Mix 9 / Non Tourist 10，可調 7–10）** Review 調貨事件是否因 Safety Stock 設定不足／過高，輸出店舖×SKU Review 主表與建議調整量（下調建議補底 3 件——SAP MOQ 下 Safety 1–2 件難補貨；現有 1–2 件不強制上調、零銷售歸零不受保底）；Safety 基準與調貨的 14 日線互相獨立；調貨分不限OM與同OM兩套對照（OM 限制不影響 Safety 建議）；可選擇**分類維度**（預設 Shop Nature，可改 RP Type、Shop Nature＋RP Type、OM、Shop Nature＋OM、銷量（高銷／中銷／低銷／零銷四帶）或銷量＋Shop Nature（Nature×銷量帶 12 組）；OM／銷量／銷量＋Shop Nature 維度計算與預設相同、僅分組顯示不同；ND 列一律不作調貨需求方、仍可作供應）；可勾選「允許雙向調貨」讓任何 Nature 缺貨皆可調入、遊客區店舖亦可轉出；可勾選「ND 優先轉出」（同一 Shop Nature 池內 ND 店先於 RF 店轉出，池間順序與轉出資格不變，不影響 Safety Review 與建議調整量）；可勾選「啟用銷量再平衡」（高銷缺貨的 RF 店舖列即使所屬 Nature 池不缺貨亦可調入、積壓列提供來源，跟隨單向／雙向框架，守衛不放寬，預設關閉）；可勾選「建議退D001」（退貨基準為 Safety Stock (現時)：中銷／低銷／零銷列退回超出 Safety 的庫存；高銷列原則上不退，僅當庫存 ≥ 3×Safety（高於 Safety 200%）時退至 2×Safety；調貨優先、需要調貨列不退；接收目標 D001 Warehouse 可無限接收，故不跑分配；不影響調貨分配與 Safety 計算；開啟時 Excel 另加「退D001 建議」工作表，預設關閉）；RAW／主表恆附「銷量帶」與「銷量存貨分類」顯示欄（同 SKU 內排名分帶＋店舖周轉日數門檻判定）；「Review 摘要」另提供 **Safety 調整方向分類**（增加／減少／不變）、**SKU 級 Safety 調整明細**與**銷量 × 存貨分類交叉表**（母體為全部列）；另附**全列預防性 Safety 設定建議**（母體為全部店舖×SKU 列，不依賴調貨事件，建議一次過設定，減少日後調貨次數；Excel 另加「Safety 設定建議 (全列)」工作表） |
| 自定義參數 | Shop Nature 過濾、安全庫存倍率、Safety Qty 計算模式、調貨門檻、Safety 基準周轉天數（按 Nature）、分類維度（Shop Nature／RP Type／Shop Nature＋RP Type／OM／Shop Nature＋OM／銷量／銷量＋Shop Nature）、允許雙向調貨、ND 優先轉出、建議退D001、缺貨深度優先、池健康度排除 ND 死庫存、有貨可調改用實際可轉出量 |
| 視覺化看板 | 即時顯示統計指標與交叉分析表 |
| Excel 導出 | 一鍵下載含 RAW、Shop Stock Status、By Shop、分析摘要之美化報表 |
| Forecast 分析 | 以 SKU 級計算建議訂貨量（v2 加權混合：缺貨還原＋開賣折算後，標準情境＝月均基準／M-1／MTD 調整月化按可調權重混合、樂觀情境＝趨勢外推設增速上限；爆升 SKU 用樂觀情境、其餘用標準情境；目標覆蓋量平穩類按目標天數、爆升類按爆升目標覆蓋月數（預設 2 個月）；Warehouse（D001）貨由 D001 向 Vendor 下單補貨（D001 PO）、Vendor 貨由店舖向供應商下單（Shop PO）；建議訂量＝min(缺口, 上限空間)防 Overstock；全 ND 的 SKU 不建議訂貨），另附店舖覆蓋檢視（有訂量 SKU 的每店可用日數與風險旗） |

## 系統架構

```
├── app.py              # 導航入口（頁面設定、共用 CSS、多頁導航）
├── views/
│   ├── analyzer_page.py    # 分析工具主頁（上傳、模式、執行、看板）
│   └── usage_page.py       # 使用說明頁（各模式說明與核心計算規則）
├── transformer.py      # 資料轉換、計算與 Excel 產生邏輯
├── version.py          # 版本號
├── requirements.txt    # Python 依賴
├── Dockerfile          # Docker 容器映像
├── zeabur.yaml         # Zeabur 部署配置
└── .streamlit/
    └── config.toml     # Streamlit 設定
```

## 快速開始

### 本機執行

```bash
pip install -r requirements.txt
streamlit run app.py
```

### Docker

```bash
docker build -t i14-report-review .
docker run -p 8080:8080 i14-report-review
```

### 部署至 Zeabur

專案已預設 `zeabur.yaml` 與 `Dockerfile`，連接 GitHub 倉庫後可直接自動部署。

**Zeabur 部署機制**：倉庫根目錄有 `Dockerfile` 時，Zeabur (zbpack) 會自動以 Docker 構建；`requirements.txt` 已固定版本（`==`），確保每次構建可重現、加快緩存命中。

**運行優化（已內建）**：

| 項目 | 設定 | 說明 |
|------|------|------|
| PORT 綁定 | `--server.port=${PORT:-8080}` | 自動適配 Zeabur 注入的 PORT，監聽 0.0.0.0 |
| 健康檢查 | `/_stcore/health` | Docker HEALTHCHECK 以 Python 內建模組檢查（無需 curl，映像更小） |
| 檔案監視 | `--server.fileWatcherType=none` | 生產環境關閉代碼監視，節省記憶體與 CPU |
| WebSocket 壓縮 | `--server.enableWebsocketCompression=false` | 大數據表傳輸時降低 CPU 負擔 |
| 遙測 | `--browser.gatherUsageStats=false` | 停用對外統計請求，加快載入 |
| 日誌 | `PYTHONUNBUFFERED=1` | 即時輸出到 Zeabur Logs |

**容量建議**：應用會將 DataFrame 與生成的 Excel 保存在 session 記憶體中，建議服務配置 **至少 1GB RAM**（多用戶同時處理大檔案時建議 2GB）。Sidebar 底部會顯示目前行程的 Memory Usage，方便監察容器資源。若服務無故重啟，通常是 OOM（記憶體不足），可升級 plan 或提醒用戶按「🧹 釋放記憶體」。

## 必要欄位

共用必要欄位（依模式檢核，hybrid 模式不要求未參與計算的 `Shop Stock Status`）：

| 欄位 | 說明 | 模式 |
|------|------|------|
| Buyer | 採購人員 | 全部 |
| Brand | 品牌名稱 | 全部 |
| SKU | 品號編碼 | 全部 |
| Shop Stock Status | 庫存狀態明細 | 除 Safety Stock x NetStock Review 外全部 |

OOS Review 建議額外包含：

| 欄位 | 說明 |
|------|------|
| Sold Qty (CM - MTD) | 月度銷售數量 |
| Adj. Safety Stock | 現有安全庫存量（報表內顯示為 Safety Stock (現時)，並非系統建議值；輸入欄名兩名皆接受，語意相同） |
| Shop Nature | 店舖性質（Tourist / Mix / Non Tourist） |

Safety Stock x NetStock Review 另建議包含 `RP Type`（RF＝可補貨／ND＝不可補貨）——ND 列不作調貨需求方、RP Type／混合分類維度與「ND 優先轉出」均依賴此欄；缺失時全部列按 RF 處理並退回 Shop Nature 維度（列 warning，不停止），「ND 優先轉出」同時自動失效。

Forecast Review 額外必要欄位（缺少時停止並提示）：`Site Code／Sold Qty (M-2)／Sold Qty (M-1)／Sold Qty (CM - MTD)／Sold Qty (P90D)／Shop Stock Qty (CM)`；`Shop Stock Qty (M-2)／Shop Stock Qty (M-1)` 為選填（缺欄時該期缺貨還原略過並列 warning）；`D001 On Order PO Qty (CM)／Shop On Order PO Qty (CM)／Shop Pending Receive (±7 days)` 缺欄時計 0 並列 warning。

Safety Stock 月度變化分析額外必要欄位（缺少時停止並提示）：`MONTH／YEAR`（位於來源檔案尾部，日後此模式恒為必要欄位）`／Site Code／Adj. Safety Stock／RP Type／Sold Qty (CM - MTD)／Net Sales (CM - MTD)／Shop Stock Qty (CM)`；`Adj. MCH1／Adj. SST` 為選填（缺欄時對應樞紐以「空白」列示並列 warning）。

## 分析模式

### Overstock Review

針對 `Shop Stock Status` 欄位篩選以下庫存過賸條件：

1. **P90D 0 Sales** — 過去 90 天內零銷售
2. **Overstock (>2\*Monthly Sold Qty & >2pcs)** — 庫存超過月均銷售兩倍且大於 2 件

系統自動計算新的 Safety Qty 建議，並產出可轉出貨量。

### OOS Review

針對 `Shop Stock Status` 欄位篩選以下缺貨條件：

1. **OOS (0 Stock)** — 現有庫存為 0
2. **Potential OOS (<7 Turnday)** — Turnday 不足 7 天
3. **Normal (Safety Qty Insufficient)** — 庫存正常但 Safety Qty 不足以覆蓋預估銷售

支援 MTD 日期選擇，自動推算月度預估銷售量（Projected Monthly Sold Qty）。

### SKU 精簡 Review

針對 RP Type 為 **ND** 的 SKU 建議退倉，區分 Supply Source（只有 Warehouse 貨可退倉）：

1. 篩選 `RP Type = ND` 的所有 SKU（不分 Shop Stock Status）
2. **Warehouse 貨**：建議全數退倉 `transfer = Shop Stock Qty (CM)`
3. **Vendor 貨**：標記「另有安排」，不可退倉 `transfer = 0`
4. ND SKU 依 Supply Source × Shop Stock Status 雙層動態分類（如 Warehouse - Overstock、Vendor - Normal）

輸出包含 RAW、Shop Stock Status、By Shop、ND SKU數據分析摘要共 4 個工作表。

### Safety Stock Review

全面審查各 SKU×Shop 的安全庫存設定，支援兩種計算基準（可切換，預設月銷倍率）：月銷倍率（Overstock Cap／OOS Floor）或周轉天數（Safety 周轉所需量，與 Safety Stock x NetStock Review 同口徑）：

1. **Overstock Cap (X)** — 基於月均銷售與 Overstock Rate 計算庫存上限，上限不超過現有 Safety 設定（Safety Stock (現時)，即來源檔案的 Adj. Safety Stock）
2. **OOS Floor (Y)** — 基於月銷預估與 OOS Rate 計算缺貨底線，最低 3；月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均（加權混合：月初以穩定的 P90D 為主，避免頭幾日線性外推失真，月尾收斂至 MTD；P90D 缺值時退回純 MTD 線性外推）
3. **Safety 周轉所需量** = ceil(Sold Qty (P90D) ÷ 90 × Safety 基準天數)，最低 3（基準天數預設 Tourist 7／Mix 9／Non Tourist 10 日，可調 7–10；負數 P90D 按 0 計算）
3a. **建議 Safety Stock**（月銷倍率基準）= max(Y, Z if Z ≤ X else X)；**建議 Safety Stock**（周轉天數基準）= Safety 周轉所需量；其中 Z = 現有 Safety 設定（Safety Stock (現時)，即來源檔案的 Adj. Safety Stock）
4. **Recommend Safety Stock Turnday** — 建議值的周轉日數 = 建議值 ÷ (P90D 月均 ÷ 30)，與來源「Turnday (P90D Qty)」同一準則；P90D 月均 = 0 時留空。店舖問題 SKU 的建議方法以「增加／減少 X 件」標明調整方向，並會顯示「建議周轉約 X 日」；當最終建議值 = 3 件且未套最低 3 件保底前 < 3（建議值由保底決定）時，建議方法加註「為保障自動補貨，最低建議設定為 3 件（最低 3 件保底）」
5. **Safety Qty Change** = 建議值 − Z（正數=增加、負數=減少、零=不變）
6. **New SKU 保護** — Launch Date 距今 90 天內的列排除「P90D 0 Sales → Safety=0/ND」override，改走正常基準計算，並在「Recommendation of new Safety Qty / ND」標記 New SKU (3 Month)（Launch Date 缺失或無法解析時視為非新品）
7. **計算模式** — 「只可增加」：全表只有增加或保持原有 Safety，且 RP Type＝ND 的列一律不作上調（ND＝不執行自動補貨，Safety 不能 > 0 件，建議值維持現有設定），ND / P90D 0 Sales 的歸零+ND 退場建議不套用（ND 欄僅保留 New SKU 標記）；「可加可減」：允許減少並套用退場建議（ND 列 → 0；P90D 0 Sales 舊品 → ND 建議，並依 Adj. MCH1 / Adj. SST 加註：02 - MAKE UP → ND(02 - MAKE UP)、03 - Private Label 系 → ND(Private Label)、兩者皆是 → ND(Private Label - MAKE UP)）

品項分為 Overstock / OOS / Balanced (Safety Adjusted) 三類，標示調整原因。
輸出包含 **說明（首頁導讀：Review 邏輯＋本次執行參數＋工作表導讀，開啟 Excel 時預設顯示本頁）**、RAW、Shop Stock Status、By Shop、店舖問題 SKU（每間店舖需行動列：Safety 需增加／減少、ND 退場，問題＋建議方法（調整量以「增加／減少 X 件」標明方向；附建議值周轉日數，受最低 3 件保底影響時加註保底原因）；表末另附「銷量帶」「銷量存貨分類」兩欄）、SUMMARY、Safety Stock 數據分析摘要共 **7 個工作表**。
SUMMARY 工作表統計 Safety Qty Change 分類（增加 / 減少 / 不變）與 Recommendation of new Safety Qty / ND 分類（空白維持 / New SKU / 0 已是 ND / ND 系列加註）的行數、SKU 數與調整量合計；另附 **③ Safety Qty Change 分類 × 分類維度交叉表**（見下）。

可選擇 **分類維度**（七選一，選項與 Safety Stock x NetStock Review 相同：Shop Nature／RP Type／Shop Nature＋RP Type／OM／Shop Nature＋OM／銷量／銷量＋Shop Nature；預設 Shop Nature）——本模式所有維度**只改分組顯示**（SUMMARY 區塊 ③），計算基準、倍率、MTD 月銷預估、退場建議等 Safety 計算一律不變。RAW 恆附「銷量帶」與「銷量存貨分類」顯示欄（門檻缺貨線 7 日／積壓線 30 日／保底 2 件，UI 可調，準則與 Safety Stock x NetStock Review 相同；銷量帶按本次上傳資料內同 SKU 排名計算，篩選範圍不同時分帶結果會不同）。來源缺「Sold Qty (P90D)」時不輸出銷量兩欄並列 warning；缺「Shop Stock Qty (CM)」時只輸出銷量帶並列 warning；缺「Shop Monthly Sold Qty (P90D)」時店舖周轉日數以 Sold Qty (P90D) ÷ 90 × 30.5 換算並列 warning；OM 兩個維度缺「OM」欄或全欄空白時退回 Shop Nature。結果頁「🧮 分類交叉表」tab 即時顯示區塊 ③ 與銷量 × 存貨分類分佈（後者母體為全部店舖×SKU 列）。

### Stock Checking（調貨檢查）

上傳 `By SKUxShop_Stock Checking_<日期>.xlsx`（D001 倉庫少貨且遊客區店舖缺貨時的檢查），計算 Tourist 鋪 SKU 的調貨需求，並從 Mix / Non Tourist 鋪分配可轉出貨量（調貨方向固定單向——只向遊客區調入；「允許雙向調貨」只適用於 Safety Stock x NetStock Review 模式）：

1. **觸發條件（SKU 級，同時符合才標示）** — 該 SKU 的「Tourist 加總 TurnDay」（各 Tourist 店 Shop Stock Qty (CM) 加總 ÷ Shop Monthly Sold Qty (P90D) 加總 × 30.5，與 Summary 同一準則）< 門檻（預設 14 日，UI 可調），且 D001 Net Stock Qty (CM) < 門檻（預設 100 件，UI 可調），且 Mix / Non Tourist 至少一個 Nature 的加總 TurnDay ≥ 門檻（有貨可調）；Tourist 加總 TurnDay ≥ 門檻的 SKU 一律無須調貨；Mix 與 Non Tourist 加總 TurnDay 皆低於門檻時視為無貨可調、不標示。只有已標示 SKU 才在 RAW 計算調貨數量（個別 Tourist 店補到 14 日周轉的需求 > 0）
2. **調入需求量** = ceil(P90D / 90 × 14) − Shop Stock Qty (CM)，補到 14 日周轉
3. **可轉出貨量（來源店）** = max(0, Shop Stock Qty (CM) − P90D / 90 × 14)；「無銷售但有存貨」全數可釋放；RP Type 為 RF 與 ND 的店皆可作為來源；該 Nature 的 SKU 級加總 TurnDay < 門檻時視為無貨可調（可轉出量歸零）
4. **分配** — 按 SKU 彙總需求，Non Tourist 店優先轉出，不足由 Mix 店補足；需求量大→小依序分配；來源店與需求店同為一店（Site Code 相同）時一律排除，防止自我調貨
5. **分類維度與銷量分類** — 可選擇 **分類維度**（七選一，選項與 Safety Stock x NetStock Review 相同；預設 Shop Nature）決定 Summary 的分組欄位；**調貨觸發、分配與數量一律按 Shop Nature 準則計算，不隨維度改變**。RAW／店舖問題 SKU 恆附「銷量帶」與「銷量存貨分類」顯示欄（門檻缺貨線 7 日／積壓線 30 日／保底 2 件，UI 可調，準則與 Safety Stock x NetStock Review 相同；銷量帶按本次上傳資料內同 SKU 排名計算，篩選範圍不同時分帶結果會不同）；來源缺「Sold Qty (P90D)」或「Shop Stock Qty (CM)」時對應欄不輸出並列 warning。存貨偏重提示門檻按維度切換（與 Safety Stock x NetStock Review 同表：nature／rp／both 40%／20pp、om 25%／12pp、sales 30%／15pp、nature_om／sales_nature 12%／8pp）

輸出包含 **說明（首頁導讀：調貨邏輯＋本次執行參數＋工作表導讀，開啟 Excel 時預設顯示本頁）**、Summary、RAW、店舖問題 SKU（每間店舖需行動列：需要調入／可轉出供應，問題＋建議方法；表末另附「銷量帶」「銷量存貨分類」兩欄）與調貨分析摘要共 **5 個工作表**：

- **Summary** — SKU 級彙總（每 SKU 一列，按 Brand / SKU 排序；**按所選分類維度分組**，nature 三組為 A–AC 契約欄序且與 Safety Stock x NetStock Review 同格式同值）：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001 Net Stock Qty (CM)；各分組加總 Shop Stock Qty (CM) 與 Shop Monthly Sold Qty (P90D)（取整）、TurnDay = ROUND(加總庫存 ÷ 加總月銷 × 30.5, 0)（取整，月銷 0 留空）、涉及店舖數（同 SKU 同組 Site Code 不重複計數，空白不計，置於該組 TurnDay 之後）、各分組佔比 (%)（庫存/月銷各 N 欄，活公式 + 0.00% 顯示，合計 0 留空）；強制調貨提示（該 SKU 有 Tourist 店需要調貨時標示）、由各分組店舖預計調出數量（nature 維度為 Non Tourist / Mix 兩欄固定順序）；存貨偏重提示（某分組庫存佔比 ≥ 門檻且庫存佔比 − 銷量佔比 ≥ 差距門檻 →「存貨偏重 {分組}（庫存 x% / 銷量 y%）」，多個分組以；分隔、差距大者在前；偏重提示欄與偏重分組的庫存佔比欄以醒目橘底加強顯示，Streamlit 同步）；文字欄靠左、數值置中、間隔列底紋、凍結表頭 + 自動篩選
- **RAW** — 按可讀性重排欄序（識別→品項屬性→庫存與銷售→SKU 級判定→調貨執行→其他原始欄），不顯示「Item_Shop」與所有含「Value」的金額欄；新增：Tourist / Mix / Non Tourist 加總 TurnDay / 需要調貨 / 調入需求量 (補到14日) / 實際調入量 / 調入缺口 / 可轉出貨量 (Mix) / 可轉出貨量 (Non Tourist) / 轉出量 / 調貨來源 (Site Code)
- **店舖問題 SKU** — 每間店舖需行動列（只收需要調入／可轉出供應，缺口在同一列加註），按店舖分組、店內按嚴重度排序，逐列顯示問題說明與建議方法（數量＋行動指引）；另顯示 Supply Source 與 D001 Net Stock Qty (CM)；缺口跟進按供貨來源分支 — Warehouse＋D001 有貨由 D001 補貨、Warehouse＋D001 無貨寫 D001缺貨、Vendor 貨向供應商追貨（Vendor 由供應商直送店舖，不經 D001）
- **調貨分析摘要** — Dashboard 指標卡（與 Streamlit 看板同一準則）：需要調貨列數 / 需調貨 SKU 數 / 調入需求量 / 實際調入量 / Non Tourist 轉出量 / Mix 轉出量 / 無法足額缺口 / 總明細行數

### Safety Stock x NetStock Review（調貨事件 Review Safety 設定）

以「調貨事件」為入口，Review Safety Stock 設定是否為缺貨／過剩的成因（**Safety 基準周轉天數按 Shop Nature 分別設定**，預設 Tourist 7 / Mix 9 / Non Tourist 10 日，可調 7–10；與調貨段的 14 日線**互相獨立**）；不再執行 Safety Stock Review 引擎（無建議值 / Rate / MTD 參數）：

0. **分類維度（UI 七選一，預設「以 Shop Nature 分類」）** — 決定調貨需求方過濾與 Summary／Review 摘要交叉表的分組欄位；**TurnDay 池與 Safety 基準天數恆按 Shop Nature，不隨維度改變**（RP Type 語意＝該店舖可否補貨：RF＝可補貨、ND＝不可補貨）。RP Type 正規化：其他／空白值歸入 ND 並列 warning（保守視為不可補貨）；**需求方一律排除 ND 列（ND 仍可作供應方轉出）——七種維度皆適用**（預設模式下 Tourist-ND 缺貨列不再標示需要調貨，屬 v1.20.0 破壞性變更）：
   - **以 Shop Nature 分類**（預設）：需求方＝Tourist 且 RF；供應方＝Mix／Non Tourist（有貨即可轉出，不論 RF／ND）；其餘規則與現行完全相同（含池級歸零與「需求 Nature 自身除外」守衛）
   - **以 RP Type 分類**：需求方擴展為**任何 Nature 的 RF 列**（該 Nature 加總 TurnDay < 門檻），並取消池級歸零與同 Nature 守衛（允許同 Nature 店間再平衡，可轉出貨量按逐列 14 日超額計；有貨可調＝同 SKU 任一列有超額）；供應方維持 Mix／Non Tourist（雙向時另加 Tourist）；Summary 與交叉表按 RF／ND 兩組分欄
    - **以 Shop Nature＋RP Type 分類**：調貨計算與預設完全相同（需求方 Tourist-RF），僅 Summary 與交叉表改按 6 組（Tourist-RF／Tourist-ND／Mix-RF／Mix-ND／Non Tourist-RF／Non Tourist-ND）分欄呈現
    - **以 OM 分類**：調貨計算與預設完全相同（需求方 Tourist-RF），僅 Summary 與交叉表改按來源 OM 欄的不重複值分組呈現（去除前後空格，空白自成「空白」組，大小寫不敏感，按字串排序）
    - **以 Shop Nature＋OM 分類**：調貨計算與預設完全相同（需求方 Tourist-RF），僅 Summary 與交叉表改按 Nature×OM 組合（如 Tourist-XXX）分欄呈現
    - **以銷量分類**：調貨計算與預設完全相同（需求方 Tourist-RF），僅 Summary 與交叉表改按銷量帶（高銷／中銷／低銷／零銷）分欄呈現；銷量帶為常開產出欄（見調貨段第 1 點子項）
    - **以銷量＋Shop Nature分類**：調貨計算與預設完全相同（需求方 Tourist-RF），僅 Summary 與交叉表改按 Nature×銷量帶 12 組（Tourist-高銷／Tourist-中銷／…／Non Tourist-零銷，Nature 外層固定順序、帶內層固定順序）分欄呈現；銷量帶準則與「以銷量分類」相同（常開產出欄，分帶結果隨本次上傳資料的篩選範圍而變）
    - 來源缺「RP Type」欄時全部列按 RF 處理（ND 排除規則無法套用）；rp／both 維度退回 nature 並列 warning（om／nature_om／sales／sales_nature 維度不退回）；來源缺「OM」欄或全欄空白時 om／nature_om 維度退回 nature 並列 warning；非法維度值退回 nature 不拋錯
1. **調貨段** — 與 Stock Checking 同規則，並疊加本模式特有的資料品質補正（Shop Nature 正規化、月銷空白／為 0 但 P90D 有銷售時逐列回補、D001 逐列 Net→Stock 回補；同檔與 Stock Checking 的結果可能有細微差異）：SKU 級 Tourist 加總 TurnDay < 門檻（預設 14）且 D001 < 門檻（預設 100）且有貨可調 → 標示需要調貨（需求列另須 RP Type＝RF）；補到 14 日周轉，Non Tourist 優先再由 Mix 補足（目標日數固定 14）。分配跑兩遍：**不限OM**（現行契約欄位）與**同OM**（來源店須與需求店同 OM 才可轉出，欄位加「(同OM)」後綴；公司規定某些情況只允許同 OM 店舖調貨，例 Ivy HBA6 → Ivy HB94）；來源店與需求店同為一店（Site Code 相同）時一律排除，防止自我調貨
    - **允許雙向調貨（UI 勾選，預設關閉）** — 針對「存貨偏重遊客區而其他區域缺貨」場景：任何 Nature 缺貨（該 Nature 加總 TurnDay < 門檻 且 D001 < 門檻 且個別店補到 14 日需求 > 0 且該列 RP Type＝RF）且其他任一 Nature 有貨（加總 TurnDay ≥ 門檻）即標示需要調貨；RAW 另加「可轉出貨量 (Tourist)」，遊客區店舖亦可轉出，來源優先序固定 **Non Tourist → Mix → Tourist**，需求 Nature 自身一律排除（分配層顯式守衛＋池級歸零雙重保證；「以 RP Type 分類」維度例外——允許同 Nature 互調）；缺貨 Nature 的超額列一律不可轉出（池級歸零）；未勾選時輸出欄位與數值與單向完全一致。雙向只適用於本模式，Stock Checking 仍為單向（只向遊客區調入）
    - **ND 優先轉出（UI 勾選，預設關閉）** — 針對「不可補貨（ND）店舖的滯留貨先清走」場景：**同一 Shop Nature 池內**，RP Type＝ND 的來源列先於 RF 列取用，池間順序仍為 Non Tourist → Mix → Tourist（雙向時才含 Tourist），展開為 **Non Tourist-ND → Non Tourist-RF → Mix-ND → Mix-RF → Tourist-ND → Tourist-RF** 六段；ND 段與 RF 段各自維持來源檔列序（不按可轉出貨量重排）。**轉出資格一律不放寬**（仍須該池加總 TurnDay ≥ 門檻、該列有 14 日超額（零銷售有貨列為全數庫存）、非自我調貨、非同 Nature（「以 RP Type 分類」維度允許同 Nature 互調）、同OM變體另須同 OM），ND 列仍不作調貨需求方；「ND」的定義為正規化後非 RF（含無法辨識而歸入 ND 的列）。因轉出列母體與分配順序無關，本開關**只改變調貨分配結果**（`轉出量` / `調貨來源` / `事件類型` / `調貨後庫存` / `調貨後 TurnDay`），**不影響 Safety Review、建議 Safety 設定值與調整量；各 Nature 轉出合計與調入總量在正常一店×一SKU 一列的資料下不變（若同一 Site Code 同時出現在需求與來源端——「Site Code × SKU」重複列異常，系統會列 warning——自我調貨守衛會令池內取用順序影響合計與缺口）**；來源檔缺「RP Type」欄時自動失效並列 warning；未勾選時輸出與現行完全一致。Review 摘要第 5 排另加「ND 來源轉出量」／「ND 來源轉出量 (同OM)」兩卡（開啟且有實際 ND 轉出時顯示），店號見「簡易調貨明細」的 `Supply RP Type` 欄。只適用於本模式，Stock Checking 無此開關
    - **缺貨深度優先（UI 勾選，預設關閉）** — 針對「供給不足時，0 庫存的斷貨店舖反被需求量較大的快滿店舖蓋過而分文未得」：同一 SKU 內需求列的分配次序由現行「需求量由大→小」改為 **「該列周轉日數由淺至深 → 需求量由大→小 → 來源檔列序」**，周轉日數＝該列調貨前 `Shop Stock Qty (CM)` ÷ 生效月銷 × 30.5（與店舖問題 SKU、銷量存貨分類同一準則）；月銷 0 的列不會成為需求方。**只改變同 SKU 內需求列的分配次序**：總調入量、來源店轉出量、可轉出量、轉出資格、觸發判定與 Safety Review／建議調整量全部不變，被改變者為 `實際調入量` / `調入缺口` / `調貨來源 (Site Code)` / `調貨後庫存` / `調貨後 TurnDay` 及其同OM 對應欄；不限OM 與同OM 兩次分配使用同一排序鍵；同周轉日數時仍以需求量為次要鍵，故同深度行為與現行一致；未勾選時輸出與現行完全一致。只適用於本模式
    - **池健康度排除 ND 死庫存（UI 勾選，預設關閉）** — 針對「不可補貨（ND）店舖的死庫存灌高池 TurnDay，令該池真實缺貨的 RF 店舖不被標示」：ND 店舖永不補貨、庫存不會再進來，計入 SKU 級加總 TurnDay 會令池看來健康（例如 Tourist-ND 店持 300 件 → Tourist 池 TurnDay 300 → 同 SKU 的 0 庫存 RF 店舖連同可用的 26 件 Non Tourist 供應完全未被觸發）。**需求側門檻**（判定該 Nature 是否缺貨——單向的 Tourist 加總 TurnDay、雙向的各 Nature 加總 TurnDay、rp 維度的所屬 Nature 加總 TurnDay）改以**只計 RP Type＝RF 列**的加總 TurnDay 為準，並另存 `Tourist 加總 TurnDay (RF)` / `Mix 加總 TurnDay (RF)` / `Non Tourist 加總 TurnDay (RF)` 三欄供稽核。**契約不動**：原有全列 `X 加總 TurnDay` 三欄值完全不變（Summary 與店舖問題 SKU 亦維持全列口徑）。**供給側（可轉出池級門檻）刻意維持計全部列（含 ND）**——ND 死庫存正是最該清走的貨，「池看來健康」在此是正確訊號；若一併排除反而會把池判為不足而鎖死，與「ND 優先轉出」的用意相反。因此觸發面只會**擴大不會收縮**，需求、分配與 Safety 計算邏輯一律不變；來源檔缺「RP Type」欄時排除前後恆等，開關自動失效並列 warning；未勾選時輸出與現行完全一致。只適用於本模式
    - **有貨可調改用實際可轉出量（UI 勾選，預設關閉）** — 針對「標示需要調貨但實際調入量 0、缺口＝全部需求」的假陽性事件列：現行「有貨可調」以 Nature 池加總 TurnDay 判定，但 (i) 池內可能沒有任何一列有 14 日超額；(ii) 某 Nature 完全沒有列時其加總 TurnDay 為 `inf` 而令門檻**空過**（該欄於 RAW 顯示為空白，稽核軌跡看不出觸發原因）。兩者在下游與真正「部分補足」的需求列完全無法區分。**開啟後「有貨可調」一律改以 SKU 級 `sum(可轉出貨量) > 0` 判定**，與「以 RP Type 分類」維度及銷量再平衡通道的既有寫法一致（七種維度同一口徑）；此條件**隱含**池級門檻（可轉出量已按池門檻逐列歸零），故池門檻仍然生效，但不再單獨作為「有貨可調」的依據。無需另行扣除需求列自身 Nature 的供給額度——可轉出量已按池門檻歸零，而需求門檻本身要求該列所屬 Nature 的池 TurnDay < 門檻（該池必被歸零、超額恆為 0），故任一需求列的供給額度天然只來自其他 Nature，與分配層「需求 Nature 自身除外」契約一致（雙向模式亦同）。**觸發面只收縮不擴大**（已加測試斷言），總調入量與需求／分配演算法不變；連帶被移除 SKU 的「可轉出未使用」來源列亦不進 Review，Safety Review 與建議調整量只隨事件列母體收縮而變動。同時新增 **`調貨補足狀態`** 欄（僅開啟時輸出，主表 29→30 欄、該欄不在 `_HYBRID_DIAG_COLUMNS` 契約內）：`已補足`（調入缺口 0）／`部分補足`（0 < 實際調入量 < 需求量）／`無貨可調`（實際調入量 0），非需求列留空，以不限OM 分配實績為準——開啟後「無貨可調」只剩真實的供給競爭／守衛阻擋（同 SKU 多個需求列爭用同一批貨、同OM 變體無同 OM 來源、自我調貨守衛）；`店舖問題 SKU` 的缺口說明據此分流（無貨可調者改建議由 D001／供應商補貨或依建議值調整 Safety Stock），指標卡末尾加補足狀態概況一排（已補足／部分補足／無貨可調列數／無貨可調 SKU 數），stats 另記 `real_supply_gate`／`fulfill_full_rows`／`fulfill_partial_rows`／`no_supply_need_rows`／`no_supply_need_skus`；未勾選時輸出與現行完全一致。只適用於本模式
    - **銷量帶與銷量存貨分類（恆輸出的顯示欄）** — `銷量帶`：同 SKU 內按 `Sold Qty (P90D)` 的店舖**平均排名**分帶——排名超過 75%＝高銷、排名在最低 25% 內＝低銷、其餘＝中銷、P90D≤0＝零銷（有銷店舖少於 4 間時不產生高銷／低銷，全歸中銷，避免小樣本假區分；並列同值同帶；按本次上傳資料內同 SKU 排名，篩選範圍不同時結果會不同）。`銷量存貨分類`：以店舖周轉日數（`Shop Stock Qty (CM)` ÷ 生效月銷 × 30.5）與三條 UI 可調門檻判定——高銷且周轉 < 缺貨線（預設 7 日，可調 1–14）→ `高銷缺貨`；低銷且周轉 ≥ 積壓線（預設 30 日，可調 15–90）且庫存 ≥ 積壓保底件數（預設 2 件，可調 1–100）→ `低銷積壓`；零銷且庫存 ≥ 保底件數 → `零銷積壓`；其餘 → `正常`（邊界：缺貨用「低於」、積壓用「不少於」）。兩欄與調貨 14 日線及 Safety 基準天數**互相獨立**，不受任何開關影響恆輸出；零銷有存但未達保底件數的列歸正常，仍由「店舖問題 SKU」的 Past 90 Days 0 Sales 轉 ND 提示涵蓋
    - **銷量再平衡（UI 勾選，預設關閉）** — 針對「店舖銷量高但缺貨、銷量低或 0 銷售卻存著大量存貨」場景：需求方在原判定之外**增列**`高銷缺貨` 列——豁免所屬 Nature 池加總 TurnDay 門檻（本通道存在目的），但仍須 RP Type＝RF、補到 14 日需求 > 0、D001 < 門檻且同 SKU 有可轉出量；積壓列（低銷／零銷積壓）經既有轉出判定提供來源（轉出量公式與池級歸零不變——需求方擴大後更多 SKU 進入調貨，健康池內積壓列自動入事件表）。**跟隨單向／雙向框架**：單向時銷量需求方仍限 Tourist 列、轉出池仍限 Mix／Non Tourist；雙向時兩端放寬至任何 Nature。自我調貨／同 Nature／同OM 守衛與 ND 不作需求方一律不放寬；新事件列自動流入 Safety Review、店舖問題 SKU、Buyer 統計與簡易調貨明細；stats 另記 `sales_need_rows`（僅由本通道新增的需求列數）；開啟時指標卡末尾另加銷量分類概況一排（高銷缺貨列數／SKU 數、積壓列數／積壓庫存件數；無第 5 排時該排為第 5 排）；未勾選時輸出欄位與數值與現行完全一致（新增的兩個銷量顯示欄除外）。只適用於本模式
 2. **Review 基準** — `Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 該列 Nature 的基準天數)`（無銷售列為 0）；基準天數按 Nature 於 UI 分別設定（預設 Tourist 7 / Mix 9 / Non Tourist 10，可調 7–10），與調貨觸發門檻及 14 日調貨線**完全無關** — 改變基準天數不會影響任何調貨欄位（`調入需求量 (補到14日)` / `實際調入量` / `調入缺口` / `轉出量` / `可轉出貨量` / `調貨後庫存` / `調貨後 TurnDay` 及其同OM變體）；RAW 與主表另逐列顯示 `Safety 基準天數` 供稽核
3. **事件列與 Review**：
   - `需要調入`（需要調貨=是 的列；單向只出現在 Tourist-RF，雙向任何 Nature 的 RF，rp 維度任何 Nature 的 RF）：現有 Safety Stock (現時) < 所需量 → `Safety 設定不足`；否則 `設定正常—缺貨非設定因素`
    - `實際轉出`／`可轉出未使用`（已標示調貨 SKU 的有貨列——單向為 Mix / Non Tourist 列，雙向含 Tourist 列，ND 列亦可為來源——轉出量 > 0 或可轉出貨量 > 0）：現有設定 > 生效基準（所需量按下調補底後的值，見下）→ `Safety 設定過高`；否則 `設定正常—過剩非設定因素`；等於邊界一律視為正常；未標示調貨的 SKU 不入 Review
    - **下調補底 3 件（SAP MOQ）** — Safety 1–2 件難以補貨而 I14 原檔無 MOQ 欄，故「設定過高」列的生效基準：所需量 ≥ 1 且現有設定 ≥ 3 件者取 `max(所需量, 3)`（補底不高於現值——現有 3 件而所需 1–2 件者補底後相等，改判設定正常，不出現調整量 0 的假動作列）；現有設定本身只有 1–2 件者維持原所需量（**不強制上調到 3**）；所需量 = 0（P90D 零銷售）的歸零建議屬退場，不受保底影響；「設定不足」列屬 TurnDay 驅動上調，一律照原所需量（含上調至 1–2 件），不受保底影響。建議值受保底影響（建議值 3 件 > 所需量 ≥ 1 件）時，店舖問題 SKU 的建議方法加註「為保障自動補貨，最低建議設定為 3 件（SAP MOQ：Safety 1–2 件難以補貨）」
    - `建議 Safety 設定值` = 不足列取所需量、過高列取生效基準（僅不足／過高列）；`建議 Safety 設定值 Turnday` = 建議值 ÷ (P90D 月均 ÷ 30)（1 位小數；月均 0 或非建議列留空；與來源「Turnday (P90D Qty)」及 Safety Stock Review 的 `Recommend Safety Stock Turnday` 同一準則）；`Safety 建議調整量 = 建議值 − 現有設定`（正=增加、負=減少）；`Safety 周轉所需量` 欄維持原所需量供稽核，保底只體現在建議值與調整量；店舖問題 SKU 的建議方法以「增加／減少 X 件」標明調整方向，並會顯示「建議周轉約 X 日」，受保底影響時另加註保底原因
4. **上下文欄位** — `調貨後庫存 = Shop Stock Qty (CM) + 實際調入量 − 轉出量`；`調貨後 TurnDay = 調貨後庫存 ÷ Shop Monthly Sold Qty (P90D) × 30.5`（月銷 0 留空；月銷為回補後值，實際採用值另存 `Effective Shop Monthly Sold Qty (P90D)` 稽核欄）；同OM變體另存「(同OM)」後綴欄。OM 限制只影響調貨分配與事件標籤（實際轉出／可轉出未使用），**Safety Review 與建議調整量兩變體完全相同**（不受調貨限制影響）
5. **OM 欄位** — 選填；來源檔案缺「OM」欄時「同OM調貨」退化為與不限OM相同（全部店舖同一群組）並列 warning；OM 空白列自成「空白」群組（大小寫不敏感）
6. **Shop Nature** — 固定保留三種；Other / 空白 / 大小寫差異一律正規化為 Non Tourist（映射列數列入 warning）；不做母體篩選 — 調貨供需需要同時看到 Tourist 需求與 Mix/Non Tourist 來源
7. **輸入檢核** — 混合必要欄位嚴格檢核：`Buyer / Brand / SKU / Shop Nature / Site Code / Shop Stock Qty (CM) / Sold Qty (P90D) / Safety Stock (現時) 或 Adj. Safety Stock（兩名皆接受，語意相同：Safety Stock (現時) ＝ 來源檔案的 Adj. Safety Stock；報表內一律顯示 Safety Stock (現時)） / Shop Monthly Sold Qty (P90D)`，不要求未參與計算的 `Shop Stock Status`，缺任一即停止；`RP Type` 為選填欄位（缺失時 ND 排除規則無法套用、全部列按 RF 處理並退回 Shop Nature 維度，列 warning）；欄名前後空格自動正規化。`Shop Stock Qty (CM)` 必須為非負整數件數（小數、負數、非有限值即拒絕）；店舖銷售、月銷及 Safety 設定的非有限值即拒絕，負數銷售按 0 處理並列 warning。月銷空白／非數值／負數，以及月銷為 0 但 P90D 有正數銷售時，逐列以 `Sold Qty (P90D) ÷ 90 × 30.5` 回補並列 warning；D001 逐列以 Net → Stock 回補（兩者皆空白或非有限值視為 0，負 Net 值保留），Summary 顯示實際採用的有效值；Safety Stock (現時) 空白或非數值時，Review 視為 0 並列 warning；「Site Code × SKU」重複列、同店 Shop Nature 不一致、OM 全欄空白均列 warning（正常輸入應為一店×一SKU一列）
8. **RAW 稽核欄位** — 新增 `D001 Effective Stock Qty (CM)`（實際採用的 D001 值）與原始 D001 Stock，以便核查 fallback

輸出包含 **9 個工作表**：說明（首頁導讀——列出本次執行參數（分類維度／調貨門檻／D001 門檻／Safety 基準天數／雙向／ND 優先轉出／銷量分類門檻／銷量再平衡／參考日期）、核心邏輯（調貨段 14 日線、Review 段 Safety 基準天數、銷量帶與銷量存貨分類、Past 90 Days 0 Sales 轉 ND）與工作表導讀（每頁內容與建議留意的部分）＋實用提示；開啟 Excel 時預設顯示本頁）、RAW（按可讀性重排欄序——識別→品項屬性→庫存與銷售（含銷量帶／銷量存貨分類）→SKU 級判定→Review 結果→調貨 (不限OM)→調貨 (同OM)→其他原始欄；不顯示「Item_Shop」與所有含「Value」的金額欄；非事件列 Review 欄留空；雙向時另含「可轉出貨量 (Tourist)」）、Summary（SKU 級彙總，按所選分類維度分組——nature 三組與 Stock Checking 同格式同值（A–AC 契約欄序：各組加總庫存／加總月銷／TurnDay／涉及店舖數，店舖數為同 SKU 同組 Site Code 不重複計數，空白不計）、rp 兩組（RF／ND）、both 六組、om（動態 OM 名單）、nature_om（三 Nature × OM）、sales（高銷／中銷／低銷／零銷 四帶）、sales_nature（三 Nature × 四帶 12 組）：各組庫存/月銷加總、TurnDay、涉及店舖數、佔比、強制調貨提示、存貨偏重提示（門檻按維度切換：nature／rp／both 為庫存佔比 ≥40% 且差距 ≥20 個百分點，om 為 ≥25% 且差距 ≥12 個百分點，nature_om／sales_nature 為 ≥12% 且差距 ≥8 個百分點，sales 為 ≥30% 且差距 ≥15 個百分點），偏重提示欄與偏重分組的庫存佔比欄以醒目橘底加強顯示；以銷量／銷量＋Shop Nature 分類時，零銷組的「加總 - Shop Stock Qty (CM) %」大於 0%（即零銷組仍有庫存）另以醒目紅底標示——零銷售而仍有存貨，理論上全屬 Overstock，不設佔比門檻一律標示，同格與橘底重疊時以紅底優先；nature 雙向時末尾另加「由Tourist 店舖預計調出數量」欄）、調貨 Review Safety Stock (不限OM) 與 調貨 Review Safety Stock (同OM)（店舖×SKU 逐列主表，含 Site Code／OM／Shop Nature／RP Type、建議值 Turnday 欄與銷量帶／銷量存貨分類，按 Brand / SKU / 事件類型 / Site Code 排序；Safety Review 欄兩表相同）、店舖問題 SKU（每間店舖需行動列：Safety 設定不足／過高＋Past 90 Days 0 Sales 轉 ND 候選（M-2／M-1／CM-MTD／P90D 皆零銷售且有庫存，距參考日期 90 天內新品、需要調貨列與 ND 列除外；重疊列保留原問題類型並加註）；按店舖分組、店內按嚴重度排序；主表以不限OM為準，另附 Supply Source、D001 Net Stock Qty (CM) 與三個同OM對照欄；建議方法以「增加／減少 X 件」標明調整方向並附建議值周轉日數，調入缺口與受保底 3 件影響時加註，缺口跟進按供貨來源分支 — Warehouse＋D001 有貨由 D001 補貨、Warehouse＋D001 無貨寫 D001缺貨、Vendor 貨向供應商追貨；參考日期由「📅 參考日期」輸入決定，預設昨日）、Review 摘要（指標卡——第 4 排為同OM調貨對照、雙向時另加第 5 排 Tourist 轉出量、銷量再平衡開啟時末尾另加銷量分類概況一排 + Safety Review × 分類維度交叉表（欄位按維度切換：nature 三組／rp 兩組／both 六組／om 動態 OM 名單／nature_om 三 Nature × OM／sales 四銷量帶／sales_nature 12 組）：列數 / SKU 數 / 建議調整量合計，含總計列；其下另加 **Safety 調整方向分類**（增加 / 減少 / 不變 × 列數 / SKU 數 / Safety 建議調整量合計，含總計列）與 **Safety 調整明細（SKU 級彙總）**：Brand / SKU / SKU Desc / Buyer / Review 列數 / 需增加店舖數 / 需減少店舖數 / 設定正常店舖數 / 現有 Safety 設定合計（調整列）/ 建議 Safety 設定合計（調整列）/ 淨調整量，按 Brand / SKU 排序；兩個 Safety 統計區塊與交叉表同母體（事件類型非空的列），且不受 OM 限制、調貨方向與分類維度影響）；最後為 **銷量 × 存貨分類交叉表**（母體為全部店舖×SKU 列，有別於 ①–④ 的事件列母體：各分類 × Shop Nature 列數、總計列數／SKU 數／店舖數、庫存與 P90D 銷量合計，含總計列；無調貨事件時仍輸出）、簡易調貨明細 (不限OM) 與 簡易調貨明細 (同OM)（一列一筆「來源店舖 → 接收店舖」調貨：Brand / SKU / SKU Desc / Supply Site Code(調貨來源) / Supply Shop Name / Supply RP Type / Supply OM / 調貨數量 / Target Site Code(接收店舖) / Target Shop Name / Target RP Type / Target OM；RP Type 緊接對應 Shop Name，調貨數量置於接收欄前分隔 Supply 與 Target；調貨來源如「HB30:2; HB49:3」會拆成兩行；只收錄實際有調入的列）。

### Safety Stock 月度變化分析

追蹤店舖每月 Safety Stock 原始值變化對生意的影響（主鍵 `Site Code × SKU × YEAR-MONTH`，兼容多店舖檔案）：

1. **YEAR-MONTH** — 由 `YEAR × 100 ＋ MONTH` 排序；`MONTH` 須為 1–12 整數，無效列略過並列警告；至少需 2 個月份方可作環比
2. **月化** — 除上載當月外，其餘月份的 `CM - MTD` 已是足月實績；最新月份按數據截止日月化（係數＝當月總日數 ÷ 截止日日號，僅套用於 `Sold Qty (CM - MTD)` 與 `Net Sales (CM - MTD)` 兩欄；`Shop Stock Qty (CM)` 與 `Adj. Safety Stock` 為快照不月化）
3. **Safety Stock Qty 變化** — 只看原始值增減：`Safety Stock Qty 變化量＝本月 Adj. Safety Stock − 上月`（正數為增加、負數為減少、零為不變，不作建議值計算）
4. **RP Type 轉變** — 記錄上月→本月（如 `RF→ND`）；Sold Qty／Net Sales 變化量以月化基準計算，變化率＝變化量 ÷ 上月絕對值（上月為 0 留空）
5. **維度樞紐（三段式）** — `Adj. MCH1`、`Adj. SST`、`MCH1×SST`、`Site Code`；四項指標全稱：Sold Qty／Net Sales／Net Stock Qty（店舖存貨 `Shop Stock Qty (CM)` 合計）／Safety Stock Qty；四項指標區塊左右並排（每區塊各含項目、各月與全部合計欄），每段一列一維度、總計列置段末：每月合計 → 月環比% → 月環比絕對值；表尾另附 RF→ND SKU數／ND→RF SKU數 兩個分月份區塊（當月發生轉變的不重複 SKU 數，按轉變的「本月」歸月；全部合計為整個期間）
6. **生意視角（三項）** — RP轉變生意對比（按 RP Type 轉變 × 本月分組，隔離 RF→ND 群組與持續群組的生意變化）／店舖精簡成效（每店每月 RF／ND／總 SKU 數、淨精簡數、被精簡生意基數與留存 RF 組對比）／轉變前後追蹤(T0)（RF→ND 與 ND→RF 事件以轉變本月為 T0 的 T-1／T0／T+1 對比與去化狀態）；RF 代表執行自動補貨，ND 代表不執行自動補貨

輸出包含 **12 個工作表**：說明（首頁導讀：月化邏輯＋本次執行參數＋工作表導讀；開啟 Excel 時預設顯示本頁）、SKU月度變化（每組相鄰月份環比事件）、By Adj. MCH1、By Adj. SST、By MCH1×SST、By Shop、RP Type 變化（RP Type 轉變 × 月份配對 事件數／SKU數，重點為 RF→ND 與 ND→RF）、RP轉變生意對比、店舖精簡成效、轉變前後追蹤(T0)、月度摘要（指標卡含 RF→ND／ND→RF 一排＋每月合計）、RAW（含 YEAR-MONTH 與月化估算欄）。

### Forecast Review（SKU 級建議訂貨量）

以 **SKU 級**計算建議訂貨量（v2 加權混合需求模型；第二維度防缺貨），另附**店舖覆蓋檢視**確保店舖不斷貨：

1. **缺貨還原（取代式，逐列；只限 RF 列）** — M-2／M-1／CM 期缺貨（該期庫存 ≤ 0）且月均基準 > 0 的列，取該期實績與該店月均基準（Shop Monthly Sold Qty (P90D)，缺欄時以 Sold Qty (P90D) ÷ 90 × 30.5 換算並列 warning）孰高；還原上限為該店月均，不疊加；CM 期先月化再還原。ND 列不還原；`Shop Stock Qty (M-2)`／`(M-1)` 為選填（M-2 缺欄時該期還原略過；M-1 缺欄但有 `Shop Stock Qty (LM)` 時以 LM 代理，兩者皆缺則略過；均列 warning）
2. **開賣日折算（SKU 級）** — 最早 Launch Date 落在 M-2／M-1 時，按期內有效日數折算為整月規模（`factor = 期間日數 ÷ max(開賣後日數, 7)`）；落在本月時 MTD 月化以開賣後有效日數為分母（最少 7 日）。無 Launch Date 或無效 MTD 日期時不折算
3. **情境（SKU 級聚合後計算）** — MTD 月化＝Sold Qty (CM - MTD) ÷ 有效日數 × 當月日數；MTD 調整月化＝按已過日數把 MTD 月化與月均基準混合（`r = MTD 日數 ÷ 當月日數`；MTD 日期無效時退回 max(M-1, 月均基準)並列 warning）；標準情境＝月均基準 × 權重 ＋ M-1 折算後 × 權重 ＋ MTD 調整月化 × 權重（預設 0.4／0.3／0.3，自動正規化）；保守情境＝月均基準（只作參考欄，不用於訂量）；樂觀情境＝max(M-1 折算後, MTD 調整月化) × (1 ＋ 衰減係數（預設 0.8）× 趨勢增速) 且不低於標準，趨勢增速 `g = max(0, M-1 ÷ M-2 − 1)` 設上限（預設 1.0 倍，可調 0.2–3.0）
4. **爆升分類（SKU 級；輸入為還原＋折算後數值）** — 新起量（M-2＋M-1 合計 ≤ 新起量基數（預設 10 件）且近期有量）／加速中（M-1 ≥ 加速倍數（預設 1.15）× M-2 且 MTD 月化 ≥ 加速倍數 × M-1）／爆升（max(M-1, MTD 月化) ÷ max(M-2, 基數) ≥ 爆升倍數（預設 1.5））用樂觀情境；其餘平穩用標準情境；月預估為採用情境無條件捨去取整
5. **位置分流** — Warehouse 貨＝店舖庫存合計＋D001 庫存 (Net)＋D001 PO＋在途；Vendor 貨＝店舖庫存合計＋Shop PO 合計＋在途（在途＝Shop Pending Receive (±7 days)，實檔別名相容；缺欄計 0 並列 warning）。Supply Source 缺欄全視 Warehouse，RP Type 缺欄全視 RF；同 SKU 多個 Supply Source 以眾數為準並列 warning
6. **建議訂量** — 目標覆蓋量：平穩類＝ceil(月預估 ÷ 30.5 × 目標天數（預設 30 日）)；爆升類（新起量／加速中／爆升）＝ceil(月預估 × 爆升目標覆蓋月數（預設 2.0 個月，可調 0.5–6.0）)；缺口＝max(0, 目標覆蓋量 − 位置）；上限空間＝max(0, ceil(月預估 × 上限月數（預設 2.0 個月）) − 位置）；建議訂量＝min(缺口, 上限空間）；缺口大於上限空間時狀態標「建議訂貨（已封頂）」；全 ND 的 SKU 建議訂量為 0（狀態「ND 不建議訂貨」）。建議訂貨：Warehouse（D001）貨由 D001 向 Vendor 下單補貨（D001 PO）；Vendor 貨由店舖向供應商下單（Shop PO）
7. **店舖覆蓋（母體：建議訂量＞0 的 SKU，每店一列）** — 店舖日銷＝MTD ÷ 有效日數（無 MTD 銷售時退回 P90D ÷ 30.5，皆 0 則可用日數留空）；可用日數＝(店舖庫存＋該店在途) ÷ 店舖日銷（1 位小數）；風險旗：已缺貨（庫存 0）／≤7日／≤目標天數／足夠；按 SKU＋風險＋日數排序。主表另計缺貨風險店數（可用日數 ≤7 日，RF／ND 皆計）／現已缺貨店數／鋪貨店數
8. **輸入檢核** — 共用必要欄位另加本模式嚴格欄（Site Code／三期銷售／P90D／店舖庫存 CM，缺任一即停止）；`Shop Stock Qty (M-2)`／`(M-1)`、D001 PO／Shop PO／在途欄為選填（缺欄計 0、略過該期還原或以 `Shop Stock Qty (LM)` 代理 M-1，並列 warning）；負銷售按 0、負數或非數值庫存按 0（列 warning，不拒絕）

輸出包含 **6 個工作表**：說明（首頁導讀＋本次參數照錄，含 v2 權重與趨勢上限；開啟 Excel 時預設顯示本頁）、Forecast 主表（SKU 級，含 MTD 調整月化與趨勢增速欄；建議訂量欄醒目底色；按 Buyer／Brand／SKU 排序）、店舖覆蓋、按 Buyer 彙總（建議訂量合計分 D001 PO／Shop PO 兩欄）、RAW（重排欄序，不顯示 Item_Shop 與 Value 金額欄；含還原值欄與逐列訊號欄）與 Forecast 摘要（指標卡 8 項＋按 Buyer 表）。

## Excel 輸出結構

| 工作表 | 說明 |
|--------|------|
| 說明 | 首頁導讀（混合模式、Safety Stock Review、Stock Checking 與 Forecast Review）— 列出本次執行參數、講解模式核心邏輯（如調貨段 14 日線／Review 段 Safety 基準／銷量帶與銷量存貨分類／0 Sales 轉 ND），並逐頁說明各工作表內容與建議留意的部分；開啟 Excel 時預設顯示本頁 |
| RAW | 完整明細資料（含所有原始及計算欄位；Stock Checking 與混合模式另重排欄序，不顯示 Item_Shop 與 Value 金額欄） |
| 店舖問題 SKU | 每間店舖需行動列（問題＋建議方法；Stock Checking／Safety Stock Review／混合模式；Safety Stock Review 與混合模式的建議方法附建議值周轉日數，建議值受最低 3 件保底影響時另加註保底原因；Stock Checking 與 Safety Stock Review 表末另附「銷量帶」「銷量存貨分類」兩欄） |
| Shop Stock Status | 採購人員 x 品牌之庫存異常交叉統計表 |
| By Shop | 各店舖層級之彙總分析 |
| SUMMARY | 分類統計 — ① Safety Qty Change 分類（增加/減少/不變）與 ② ND 建議分類，另附 ③ Safety Qty Change 分類 × 分類維度交叉表（列數／SKU 數／Safety Qty Change 合計，含總計列；分組按所選分類維度切換）（Safety Stock Review） |
| 分析摘要 | 數據看板 — 關鍵指標總覽 |
| Summary | SKU 級調貨彙總 — 庫存/月銷加總、TurnDay、涉及店舖數、佔比、調貨提示、存貨偏重提示（門檻按維度切換：nature／rp／both 為分組庫存佔比 ≥40% 且較銷量佔比高 ≥20 個百分點，om 為 ≥25% 且差距 ≥12 個百分點，nature_om／sales_nature 為 ≥12% 且差距 ≥8 個百分點，sales 為 ≥30% 且差距 ≥15 個百分點；店舖數為同 SKU 同組 Site Code 不重複計數，空白不計；混合模式按所選分類維度分組——nature 三組／rp 兩組 RF／ND／both 六組／om 動態 OM 名單／nature_om 三 Nature × OM／sales 四銷量帶／sales_nature 三 Nature × 四帶 12 組；Stock Checking 與混合模式的 Summary 均以醒目橘底加強顯示偏重提示與偏重分組的庫存佔比欄；混合模式另於以銷量／銷量＋Shop Nature 分類時以醒目紅底標示零銷組庫存佔比大於 0% 的儲存格（零銷售而仍有存貨，理論上全屬 Overstock））（Stock Checking 及混合模式均按所選分類維度分組——Stock Checking 於 v1.37.0 起支援七選一維度，調貨計算不隨維度改變） |
| 調貨分析摘要 | 調貨 Dashboard — 需求/調入/轉出/缺口（Stock Checking） |
| 調貨 Review Safety Stock (不限OM)／(同OM) | 店舖×SKU Review 主表兩版（僅混合模式）— 需要調入／實際轉出／可轉出未使用事件列，按 Nature 的 Safety 基準天數（預設 7 / 9 / 10）Review 與建議調整量，逐列含「Safety 基準天數」「Safety 周轉所需量」「建議 Safety 設定值 Turnday」及 RP Type、銷量帶、銷量存貨分類；同OM版來源店須與需求店同 OM |
| Review 摘要 | 混合模式 — 指標卡（第 4 排為同OM調貨對照；第 5 排為轉出來源對照——雙向調貨時含 Tourist 轉出量兩卡，「ND 優先轉出」開啟且 ND 店舖有實際轉出時另加「ND 來源轉出量」／「ND 來源轉出量 (同OM)」兩卡，該排 2–4 卡；「銷量再平衡」開啟時末尾另加銷量分類概況一排 4 卡）+ Safety Review × 分類維度交叉表（欄位按維度切換：nature 三組／rp 兩組／both 六組／om 動態 OM 名單／nature_om 三 Nature × OM／sales 四銷量帶／sales_nature 三 Nature × 四帶 12 組；列數 / SKU 數 / 建議調整量合計，含總計列）+ Safety 調整方向分類（增加 / 減少 / 不變 × 列數 / SKU 數 / Safety 建議調整量合計，含總計列）+ Safety 調整明細（SKU 級彙總：Review 列數 / 需增加店舖數 / 需減少店舖數 / 設定正常店舖數 / 現有與建議 Safety 設定合計（調整列）/ 淨調整量）＋ 銷量 × 存貨分類交叉表（母體為全部店舖×SKU 列：各分類 × Shop Nature 列數、總計列數／SKU 數／店舖數、庫存與 P90D 銷量合計，含總計列；無調貨事件時仍輸出）；①–④ 同母體（事件類型非空的列），不受 OM 限制影響 |
| 簡易調貨明細 (不限OM)／(同OM) | 混合模式 — 一列一筆「來源店舖 → 接收店舖」調貨（Brand / SKU / SKU Desc / Supply Site Code(調貨來源) / Supply Shop Name / Supply RP Type / Supply OM / 調貨數量 / Target Site Code(接收店舖) / Target Shop Name / Target RP Type / Target OM）；來源如「HB30:2; HB49:3」拆成兩行；只收錄實際有調入的列；Streamlit 結果頁另有同名 tab（可切換兩版） |
| Forecast 主表 | Forecast Review — SKU 級建議訂貨主表（一 SKU 一列）：識別／三期銷售＋月化（還原及開賣折算後）／爆升分類／三情境＋MTD 調整月化＋趨勢增速＋採用月預估／位置分項與合計／目標覆蓋量／缺口／上限空間／建議訂量（醒目底色）／建議訂貨對象（D001 PO／Shop PO）／SKU 狀態／風險店數；按 Buyer／Brand／SKU 排序 |
| 店舖覆蓋 | Forecast Review — 有訂量 SKU 的每店可用日數與風險旗（按 SKU＋風險＋日數排序） |
| 按 Buyer 彙總 | Forecast Review — 各 Buyer 的建議訂量合計（分 D001 PO／Shop PO 兩欄） |
| Forecast 摘要 | Forecast Review Dashboard — 指標卡（SKU 數／爆升 SKU 數／建議總訂量／其中 D001 PO／其中 Shop PO／已封頂 SKU 數／ND 不訂 SKU 數／店舖覆蓋列數）＋按 Buyer 表 |
| SKU月度變化 | Safety Stock 月度變化分析 — 每組相鄰月份環比事件（上月／本月值＋變化量＋變化率；Safety Stock Qty 只看原始值；Sold Qty／Net Sales 以月化基準） |
| By Adj. MCH1／By Adj. SST／By MCH1×SST | Safety Stock 月度變化分析 — 三段式樞紐（每月合計／月環比%／月環比絕對值），四項指標全稱區塊左右並排，每區塊各含項目、各月與全部合計欄；每段一列一維度，總計列（粗體灰底）置段末；表尾另附 RF→ND SKU數／ND→RF SKU數 分月份區塊（全部合計為整個期間不重複 SKU） |
| By Shop | Safety Stock 月度變化分析 — 每間店舖每月合計＋環比變化＋RF→ND SKU數／ND→RF SKU數（當月發生轉變的不重複 SKU 數，按轉變的本月歸月） |
| RP Type 變化 | Safety Stock 月度變化分析 — 上半：RP Type 轉變 × 月份配對（事件數／SKU數）；下半：首尾月份對比明細（最早與最新月份皆有資料的組合，有轉變列置前；重點為 RF→ND 與 ND→RF；含最早／最新月份 Sold Qty 月化、Net Sales 月化與 Safety Stock Qty 值＋變化量，變化率以最早月份絕對值為分母，為 0 留空） |
| 月度摘要 | Safety Stock 月度變化分析 Dashboard — 指標卡（環比事件數／Safety Stock Qty 增加／減少／不變／淨變化量／SKU 數／店舖數／月份數／RF→ND 事件／ND→RF 事件／RF→ND SKU數／ND→RF SKU數）＋每月合計（全稱欄＋每月 RF→ND／ND→RF 事件數） |

## 版本

查看 [VERSION.md](VERSION.md) 了解完整版本歷史。

## 技術棧

- **前端/框架**: Streamlit
- **資料處理**: Pandas, NumPy
- **Excel 產生**: OpenPyXL（安裝 lxml 時改用 lxml 寫出器加速存檔）
- **部署**: Docker, Zeabur
