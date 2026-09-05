# I14 報表轉換與分析系統

自動化 I14 庫存報表轉換與多維度數據分析工具，支援 **Overstock Review**（庫存過賸審查）、**OOS Review**（缺貨審查）、**Carry SKU Review**（在售品項盤點）、**SKU 精簡 Review**（ND SKU 退倉建議）、**Safety Stock Review**（安全庫存審查）與 **Safety Stock x NetStock Review**（調貨事件 Review Safety 設定）等多種模式。

## 功能概述

| 功能 | 說明 |
|------|------|
| 檔案上傳與驗證 | 上傳 By SKUxShop Excel (.xlsx)，自動檢核必要欄位 |
| 模式選擇 | Overstock Review / OOS Review / Carry SKU Review / SKU 精簡 Review / NDRF 月份對比 / Safety Stock Review / Stock Impact Analysis / Stock Impact Analysis（簡化）/ Past 90 Days 0 Sales / Stock Checking / Safety Stock x NetStock Review |
| Overstock 分析 | 篩選 P90D 零銷售及超過月均兩倍之庫存過賸品項 |
| OOS 分析 | 篩選缺貨（0 Stock）、潛在缺貨（Turnday < 7）及 Safety Qty 不足 |
| SKU 精簡分析 | 針對 RP Type = ND 的 SKU 建議退倉（Warehouse 可退倉、Vendor 另有安排），依 Supply Source × Shop Stock Status 雙層分類 |
| Stock Checking 分析 | D001 倉庫少貨且 Tourist 店舖 Turnday < 門檻 → 標示調貨需求，從 Mix / Non Tourist 店舖分配可轉出貨量（Non Tourist 優先）；Summary 另以佔比差額標示「存貨偏重」的 Nature（存貨集中但銷量不高，偏重數值以醒目橘底加強顯示）；RAW 按可讀性重排欄序，不顯示 Item_Shop 與 Value 金額欄；另附「店舖問題 SKU」（每間店舖需行動列：需要調入／可轉出供應，問題＋建議方法；建議方法只計涉及來源店舖數量，不重複列出店號） |
| Safety Stock x NetStock | 沿用 Stock Checking 調貨引擎，以 **Safety 基準周轉天數（按 Nature，預設 Tourist 7 / Mix 9 / Non Tourist 10，可調 7–10）** Review 調貨事件是否因 Safety Stock 設定不足／過高，輸出店舖×SKU Review 主表與建議調整量；Safety 基準與調貨的 14 日線互相獨立；調貨分不限OM與同OM兩套對照（OM 限制不影響 Safety 建議）；可選擇**分類維度**（預設 Shop Nature，可改 RP Type 或 Shop Nature＋RP Type；ND 列一律不作調貨需求方、仍可作供應）；可勾選「允許雙向調貨」讓任何 Nature 缺貨皆可調入、遊客區店舖亦可轉出；「Review 摘要」另提供 **Safety 調整方向分類**（增加／減少／不變）與 **SKU 級 Safety 調整明細** |
| 自定義參數 | Shop Nature 過濾、安全庫存倍率、Safety Qty 計算模式、調貨門檻、Safety 基準周轉天數（按 Nature）、分類維度（Shop Nature／RP Type／Shop Nature＋RP Type） |
| 視覺化看板 | 即時顯示統計指標與交叉分析表 |
| Excel 導出 | 一鍵下載含 RAW、Shop Stock Status、By Shop、分析摘要之美化報表 |

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
| Adj. Safety Stock | 現有安全庫存量 |
| Shop Nature | 店舖性質（Tourist / Mix / Non Tourist） |

Safety Stock x NetStock Review 另建議包含 `RP Type`（RF＝可補貨／ND＝不可補貨）——ND 列不作調貨需求方、RP Type／混合分類維度均依賴此欄；缺失時全部列按 RF 處理並退回 Shop Nature 維度（列 warning，不停止）。

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

全面審查各 SKU×Shop 的安全庫存設定，結合 Overstock Cap 與 OOS Floor 計算平衡建議值：

1. **Overstock Cap (X)** — 基於月均銷售與 Overstock Rate 計算庫存上限，上限不超過現有 Adj. Safety Stock
2. **OOS Floor (Y)** — 基於月銷預估與 OOS Rate 計算缺貨底線，最低 3；月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均（加權混合：月初以穩定的 P90D 為主，避免頭幾日線性外推失真，月尾收斂至 MTD；P90D 缺值時退回純 MTD 線性外推）
3. **建議 Safety Stock** = max(Y, Z if Z ≤ X else X)，其中 Z = 現有 Adj. Safety Stock
4. **Recommend Safety Stock Turnday** — 建議值的周轉日數 = 建議值 ÷ (P90D 月均 ÷ 30)，與來源「Turnday (P90D Qty)」同口徑；P90D 月均 = 0 時留空
5. **Safety Qty Change** = 建議值 − Z（正數=增加、負數=減少、零=不變）
6. **New SKU 保護** — Launch Date 距今 90 天內的列排除「P90D 0 Sales → Safety=0/ND」override，改走正常 X/Y 計算，並在「Recommendation of new Safety Qty / ND」標記 New SKU (3 Month)（Launch Date 缺失或無法解析時視為非新品）
7. **計算模式** — 「只可增加」：全表只有增加或保持原有 Safety，ND / P90D 0 Sales 的歸零+ND 退場建議不套用（ND 欄僅保留 New SKU 標記）；「可加可減」：允許減少並套用退場建議（ND 列 → 0；P90D 0 Sales 舊品 → ND 建議，並依 Adj. MCH1 / Adj. SST 加註：02 - MAKE UP → ND(02 - MAKE UP)、03 - Private Label 系 → ND(Private Label)、兩者皆是 → ND(Private Label - MAKE UP)）

品項分為 Overstock / OOS / Balanced (Safety Adjusted) 三類，標示調整原因。
輸出包含 RAW、Shop Stock Status、By Shop、店舖問題 SKU（每間店舖需行動列：Safety 需增加／減少、ND 退場，問題＋建議方法）、SUMMARY、Safety Stock 數據分析摘要共 6 個工作表。
SUMMARY 工作表統計 Safety Qty Change 分類（增加 / 減少 / 不變）與 Recommendation of new Safety Qty / ND 分類（空白維持 / New SKU / 0 已是 ND / ND 系列加註）的行數、SKU 數與調整量合計。

### Stock Checking（調貨檢查）

上傳 `By SKUxShop_Stock Checking_<日期>.xlsx`（D001 倉庫少貨且遊客區店舖缺貨時的檢查），計算 Tourist 鋪 SKU 的調貨需求，並從 Mix / Non Tourist 鋪分配可轉出貨量（調貨方向固定單向——只向遊客區調入；「允許雙向調貨」只適用於 Safety Stock x NetStock Review 模式）：

1. **觸發條件（SKU 級，同時符合才標示）** — 該 SKU 的「Tourist 加總 TurnDay」（各 Tourist 店 Shop Stock Qty (CM) 加總 ÷ Shop Monthly Sold Qty (P90D) 加總 × 30.5，與 Summary 同口徑）< 門檻（預設 14 日，UI 可調），且 D001 Net Stock Qty (CM) < 門檻（預設 100 件，UI 可調），且 Mix / Non Tourist 至少一個 Nature 的加總 TurnDay ≥ 門檻（有貨可調）；Tourist 加總 TurnDay ≥ 門檻的 SKU 一律無須調貨；Mix 與 Non Tourist 加總 TurnDay 皆低於門檻時視為無貨可調、不標示。只有已標示 SKU 才在 RAW 計算調貨數量（個別 Tourist 店補到 14 日周轉的需求 > 0）
2. **調入需求量** = ceil(P90D / 90 × 14) − Shop Stock Qty (CM)，補到 14 日周轉
3. **可轉出貨量（來源店）** = max(0, Shop Stock Qty (CM) − P90D / 90 × 14)；「無銷售但有存貨」全數可釋放；RP Type 為 RF 與 ND 的店皆可作為來源；該 Nature 的 SKU 級加總 TurnDay < 門檻時視為無貨可調（可轉出量歸零）
4. **分配** — 按 SKU 彙總需求，Non Tourist 店優先轉出，不足由 Mix 店補足；需求量大→小依序分配；來源店與需求店同為一店（Site Code 相同）時一律排除，防止自我調貨

輸出包含 Summary、RAW、店舖問題 SKU（每間店舖需行動列：需要調入／可轉出供應，問題＋建議方法）與調貨分析摘要四個工作表：

- **Summary** — SKU 級彙總（每 SKU 一列，按 Brand / SKU 排序）：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001 Net Stock Qty (CM)；各 Shop Nature（Tourist / Mix / Non Tourist）加總 Shop Stock Qty (CM) 與 Shop Monthly Sold Qty (P90D)（取整）、TurnDay = ROUND(加總庫存 ÷ 加總月銷 × 30.5, 0)（取整，月銷 0 留空）、涉及店舖數（同 SKU 同組 Site Code 不重複計數，空白不計，置於該組 TurnDay 之後）、各 Nature 佔比 (%) = 該 Nature 加總 ÷ 三 Nature 合計（庫存/月銷各 3 欄，活公式 + 0.00% 顯示，合計 0 留空）；強制調貨提示（該 SKU 有 Tourist 店需要調貨時標示）、由 Non Tourist / Mix 店舖預計調出數量；存貨偏重提示（某 Nature 庫存佔比 ≥ 40% 且庫存佔比 − 銷量佔比 ≥ 20 個百分點 →「存貨偏重 {Nature}（庫存 x% / 銷量 y%）」，多個 Nature 以；分隔、差距大者在前；偏重提示欄與偏重 Nature 的庫存佔比欄以醒目橘底加強顯示，Streamlit 同步）；文字欄靠左、數值置中、間隔列底紋、凍結表頭 + 自動篩選
- **RAW** — 按可讀性重排欄序（識別→品項屬性→庫存與銷售→SKU 級判定→調貨執行→其他原始欄），不顯示「Item_Shop」與所有含「Value」的金額欄；新增：Tourist / Mix / Non Tourist 加總 TurnDay / 需要調貨 / 調入需求量 (補到14日) / 實際調入量 / 調入缺口 / 可轉出貨量 (Mix) / 可轉出貨量 (Non Tourist) / 轉出量 / 調貨來源 (Site Code)
- **店舖問題 SKU** — 每間店舖需行動列（只收需要調入／可轉出供應，缺口在同一列加註），按店舖分組、店內按嚴重度排序，逐列顯示問題說明與建議方法（數量＋行動指引）；另顯示 Supply Source 與 D001 Net Stock Qty (CM)；缺口跟進按供貨來源分支 — Warehouse＋D001 有貨由 D001 補貨、Warehouse＋D001 無貨寫 D001缺貨、Vendor 貨向供應商追貨（Vendor 由供應商直送店舖，不經 D001）
- **調貨分析摘要** — Dashboard 指標卡（與 Streamlit 看板同口徑）：需要調貨列數 / 需調貨 SKU 數 / 調入需求量 / 實際調入量 / Non Tourist 轉出量 / Mix 轉出量 / 無法足額缺口 / 總明細行數

### Safety Stock x NetStock Review（調貨事件 Review Safety 設定）

以「調貨事件」為入口，Review Safety Stock 設定是否為缺貨／過剩的成因（**Safety 基準周轉天數按 Shop Nature 分別設定**，預設 Tourist 7 / Mix 9 / Non Tourist 10 日，可調 7–10；與調貨段的 14 日線**互相獨立**）；不再執行 Safety Stock Review 引擎（無建議值 / Rate / MTD 參數）：

0. **分類維度（UI 三選一，預設「以 Shop Nature 分類」）** — 決定調貨需求方過濾與 Summary／Review 摘要交叉表的分組欄位；**TurnDay 池與 Safety 基準天數恆按 Shop Nature，不隨維度改變**（RP Type 語意＝該店舖可否補貨：RF＝可補貨、ND＝不可補貨）。RP Type 正規化：其他／空白值歸入 ND 並列 warning（保守視為不可補貨）；**需求方一律排除 ND 列（ND 仍可作供應方轉出）——三種維度皆適用**（預設模式下 Tourist-ND 缺貨列不再標示需要調貨，屬 v1.20.0 破壞性變更）：
   - **以 Shop Nature 分類**（預設）：需求方＝Tourist 且 RF；供應方＝Mix／Non Tourist（有貨即可轉出，不論 RF／ND）；其餘規則與現行完全相同（含池級歸零與「需求 Nature 自身除外」守衛）
   - **以 RP Type 分類**：需求方擴展為**任何 Nature 的 RF 列**（該 Nature 加總 TurnDay < 門檻），並取消池級歸零與同 Nature 守衛（允許同 Nature 店間再平衡，可轉出貨量按逐列 14 日超額計；有貨可調＝同 SKU 任一列有超額）；供應方維持 Mix／Non Tourist（雙向時另加 Tourist）；Summary 與交叉表按 RF／ND 兩組分欄
   - **以 Shop Nature＋RP Type 分類**：調貨計算與預設完全相同（需求方 Tourist-RF），僅 Summary 與交叉表改按 6 組（Tourist-RF／Tourist-ND／Mix-RF／Mix-ND／Non Tourist-RF／Non Tourist-ND）分欄呈現
   - 來源缺「RP Type」欄時全部列按 RF 處理（ND 排除規則無法套用）、維度退回 nature 並列 warning；非法維度值退回 nature 不拋錯
1. **調貨段** — 與 Stock Checking 同規則，並疊加本模式特有的資料品質補正（Shop Nature 正規化、月銷空白／為 0 但 P90D 有銷售時逐列回補、D001 逐列 Net→Stock 回補；同檔與 Stock Checking 的結果可能有細微差異）：SKU 級 Tourist 加總 TurnDay < 門檻（預設 14）且 D001 < 門檻（預設 100）且有貨可調 → 標示需要調貨（需求列另須 RP Type＝RF）；補到 14 日周轉，Non Tourist 優先再由 Mix 補足（目標日數固定 14）。分配跑兩遍：**不限OM**（現行契約欄位）與**同OM**（來源店須與需求店同 OM 才可轉出，欄位加「(同OM)」後綴；公司規定某些情況只允許同 OM 店舖調貨，例 Ivy HBA6 → Ivy HB94）；來源店與需求店同為一店（Site Code 相同）時一律排除，防止自我調貨
   - **允許雙向調貨（UI 勾選，預設關閉）** — 針對「存貨偏重遊客區而其他區域缺貨」場景：任何 Nature 缺貨（該 Nature 加總 TurnDay < 門檻 且 D001 < 門檻 且個別店補到 14 日需求 > 0 且該列 RP Type＝RF）且其他任一 Nature 有貨（加總 TurnDay ≥ 門檻）即標示需要調貨；RAW 另加「可轉出貨量 (Tourist)」，遊客區店舖亦可轉出，來源優先序固定 **Non Tourist → Mix → Tourist**，需求 Nature 自身一律排除（分配層顯式守衛＋池級歸零雙重保證；「以 RP Type 分類」維度例外——允許同 Nature 互調）；缺貨 Nature 的超額列一律不可轉出（池級歸零）；未勾選時輸出欄位與數值與單向完全一致。雙向只適用於本模式，Stock Checking 仍為單向（只向遊客區調入）
 2. **Review 基準** — `Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 該列 Nature 的基準天數)`（無銷售列為 0）；基準天數按 Nature 於 UI 分別設定（預設 Tourist 7 / Mix 9 / Non Tourist 10，可調 7–10），與調貨觸發門檻及 14 日調貨線**完全無關** — 改變基準天數不會影響任何調貨欄位（`調入需求量 (補到14日)` / `實際調入量` / `調入缺口` / `轉出量` / `可轉出貨量` / `調貨後庫存` / `調貨後 TurnDay` 及其同OM變體）；RAW 與主表另逐列回顯 `Safety 基準天數` 供稽核
3. **事件列與 Review**：
   - `需要調入`（需要調貨=是 的列；單向只出現在 Tourist-RF，雙向任何 Nature 的 RF，rp 維度任何 Nature 的 RF）：現有 Adj. Safety Stock < 所需量 → `Safety 設定不足`；否則 `設定正常—缺貨非設定因素`
   - `實際轉出`／`可轉出未使用`（已標示調貨 SKU 的有貨列——單向為 Mix / Non Tourist 列，雙向含 Tourist 列，ND 列亦可為來源——轉出量 > 0 或可轉出貨量 > 0）：現有設定 > 所需量 → `Safety 設定過高`；否則 `設定正常—過剩非設定因素`；等於邊界一律視為正常；未標示調貨的 SKU 不入 Review
   - `建議 Safety 設定值 = Safety 周轉所需量`（僅不足／過高列）；`Safety 建議調整量 = 建議值 − 現有設定`（正=增加、負=減少）
4. **上下文欄位** — `調貨後庫存 = Shop Stock Qty (CM) + 實際調入量 − 轉出量`；`調貨後 TurnDay = 調貨後庫存 ÷ Shop Monthly Sold Qty (P90D) × 30.5`（月銷 0 留空；月銷為回補後值，實際採用值另存 `Effective Shop Monthly Sold Qty (P90D)` 稽核欄）；同OM變體另存「(同OM)」後綴欄。OM 限制只影響調貨分配與事件標籤（實際轉出／可轉出未使用），**Safety Review 與建議調整量兩變體完全相同**（不受調貨限制影響）
5. **OM 欄位** — 選填；來源檔案缺「OM」欄時「同OM調貨」退化為與不限OM相同（全部店舖同一群組）並列 warning；OM 空白列歸入「(未填OM)」群組
6. **Shop Nature** — 固定保留三種；Other / 空白 / 大小寫差異一律正規化為 Non Tourist（映射列數列入 warning）；不做母體篩選 — 調貨供需需要同時看到 Tourist 需求與 Mix/Non Tourist 來源
7. **輸入檢核** — 混合必要欄位嚴格檢核：`Buyer / Brand / SKU / Shop Nature / Site Code / Shop Stock Qty (CM) / Sold Qty (P90D) / Adj. Safety Stock / Shop Monthly Sold Qty (P90D)`，不要求未參與計算的 `Shop Stock Status`，缺任一即停止；`RP Type` 為選填欄位（缺失時 ND 排除規則無法套用、全部列按 RF 處理並退回 Shop Nature 維度，列 warning）；欄名前後空格自動正規化。`Shop Stock Qty (CM)` 必須為非負整數件數（小數、負數、非有限值即拒絕）；店舖銷售、月銷及 Safety 設定的非有限值即拒絕，負數銷售按 0 處理並列 warning。月銷空白／非數值／負數，以及月銷為 0 但 P90D 有正數銷售時，逐列以 `Sold Qty (P90D) ÷ 90 × 30.5` 回補並列 warning；D001 逐列以 Net → Stock 回補（兩者皆空白或非有限值視為 0，負 Net 值保留），Summary 顯示實際採用的有效值；Adj. Safety Stock 空白或非數值時，Review 視為 0 並列 warning；「Site Code × SKU」重複列、同店 Shop Nature 不一致、OM 全欄空白均列 warning（正常輸入應為一店×一SKU一列）
8. **RAW 稽核欄位** — 新增 `D001 Effective Stock Qty (CM)`（實際採用的 D001 值）與原始 D001 Stock，以便核查 fallback

輸出包含 **8 個工作表**：RAW（按可讀性重排欄序——識別→品項屬性→庫存與銷售→SKU 級判定→Review 結果→調貨 (不限OM)→調貨 (同OM)→其他原始欄；不顯示「Item_Shop」與所有含「Value」的金額欄；非事件列 Review 欄留空；雙向時另含「可轉出貨量 (Tourist)」）、Summary（SKU 級彙總，按所選分類維度分組——nature 三組與 Stock Checking 同格式同值（A–AC 契約欄序：各組加總庫存／加總月銷／TurnDay／涉及店舖數，店舖數為同 SKU 同組 Site Code 不重複計數，空白不計）、rp 兩組（RF／ND）、both 六組：各組庫存/月銷加總、TurnDay、涉及店舖數、佔比、強制調貨提示、存貨偏重提示，偏重提示欄與偏重分組的庫存佔比欄以醒目橘底加強顯示；nature 雙向時末尾另加「由Tourist 店舖預計調出數量」欄）、調貨 Review Safety Stock (不限OM) 與 調貨 Review Safety Stock (同OM)（店舖×SKU 逐列主表，含 Site Code／OM／Shop Nature／RP Type，按 Brand / SKU / 事件類型 / Site Code 排序；Safety Review 欄兩表相同）、店舖問題 SKU（每間店舖需行動列：只收 Safety 設定不足／過高，設定正常列排除；按店舖分組、店內按嚴重度排序；主表以不限OM為準，另附 Supply Source、D001 Net Stock Qty (CM) 與三個同OM對照欄；調入缺口在建議方法加註，缺口跟進按供貨來源分支 — Warehouse＋D001 有貨由 D001 補貨、Warehouse＋D001 無貨寫 D001缺貨、Vendor 貨向供應商追貨）、Review 摘要（指標卡——第 4 排為同OM調貨對照、雙向時另加第 5 排 Tourist 轉出量 + Safety Review × 分類維度交叉表（欄位按維度切換：nature 三組／rp 兩組／both 六組）：列數 / SKU 數 / 建議調整量合計，含總計列；其下另加 **Safety 調整方向分類**（增加 / 減少 / 不變 × 列數 / SKU 數 / Safety 建議調整量合計，含總計列）與 **Safety 調整明細（SKU 級彙總）**：Brand / SKU / SKU Desc / Buyer / Review 列數 / 需增加店舖數 / 需減少店舖數 / 設定正常店舖數 / 現有 Safety 設定合計（調整列）/ 建議 Safety 設定合計（調整列）/ 淨調整量，按 Brand / SKU 排序；兩個 Safety 統計區塊與交叉表同母體（事件類型非空的列），且不受 OM 限制、調貨方向與分類維度影響）、簡易調貨明細 (不限OM) 與 簡易調貨明細 (同OM)（一列一筆「來源店舖 → 接收店舖」調貨：Brand / SKU / SKU Desc / Target Site Code(接收店舖) / Target Shop Name / Target RP Type / Target OM / 調貨數量 / Supply Site Code(調貨來源) / Supply Shop Name / Supply RP Type / Supply OM；RP Type 緊接對應 Shop Name，調貨數量置於來源欄前分隔 Target 與 Supply；調貨來源如「HB30:2; HB49:3」會拆成兩行；只收錄實際有調入的列）。

## Excel 輸出結構

| 工作表 | 說明 |
|--------|------|
| RAW | 完整明細資料（含所有原始及計算欄位；Stock Checking 與混合模式另重排欄序，不顯示 Item_Shop 與 Value 金額欄） |
| 店舖問題 SKU | 每間店舖需行動列（問題＋建議方法；Stock Checking／Safety Stock Review／混合模式） |
| Shop Stock Status | 採購人員 x 品牌之庫存異常交叉統計表 |
| By Shop | 各店舖層級之彙總分析 |
| SUMMARY | 分類統計 — Safety Qty Change 分類（增加/減少/不變）與 ND 建議分類（Safety Stock Review） |
| 分析摘要 | 數據看板 — 關鍵指標總覽 |
| Summary | SKU 級調貨彙總 — 庫存/月銷加總、TurnDay、涉及店舖數、佔比、調貨提示、存貨偏重提示（分組庫存佔比 ≥40% 且較銷量佔比高 ≥20 個百分點的分組；店舖數為同 SKU 同組 Site Code 不重複計數，空白不計；混合模式按所選分類維度分組——nature 三組／rp 兩組 RF／ND／both 六組；Stock Checking 與混合模式的 Summary 均以醒目橘底加強顯示偏重提示與偏重分組的庫存佔比欄）（Stock Checking 及混合模式） |
| 調貨分析摘要 | 調貨 Dashboard — 需求/調入/轉出/缺口（Stock Checking） |
| 調貨 Review Safety Stock (不限OM)／(同OM) | 店舖×SKU Review 主表兩版（僅混合模式）— 需要調入／實際轉出／可轉出未使用事件列，按 Nature 的 Safety 基準天數（預設 7 / 9 / 10）Review 與建議調整量，逐列含「Safety 基準天數」與「Safety 周轉所需量」及 RP Type；同OM版來源店須與需求店同 OM |
| Review 摘要 | 混合模式 — 指標卡（第 4 排為同OM調貨對照；雙向調貨時另加第 5 排 Tourist 轉出量）+ Safety Review × 分類維度交叉表（欄位按維度切換：nature 三組／rp 兩組／both 六組；列數 / SKU 數 / 建議調整量合計，含總計列）+ Safety 調整方向分類（增加 / 減少 / 不變 × 列數 / SKU 數 / Safety 建議調整量合計，含總計列）+ Safety 調整明細（SKU 級彙總：Review 列數 / 需增加店舖數 / 需減少店舖數 / 設定正常店舖數 / 現有與建議 Safety 設定合計（調整列）/ 淨調整量）；三個區塊同母體（事件類型非空的列），不受 OM 限制影響 |
| 簡易調貨明細 (不限OM)／(同OM) | 混合模式 — 一列一筆「來源店舖 → 接收店舖」調貨（Brand / SKU / SKU Desc / Target Site Code(接收店舖) / Target Shop Name / Target RP Type / Target OM / 調貨數量 / Supply Site Code(調貨來源) / Supply Shop Name / Supply RP Type / Supply OM）；來源如「HB30:2; HB49:3」拆成兩行；只收錄實際有調入的列；Streamlit 結果頁另有同名 tab（可切換兩版） |

## 版本

查看 [VERSION.md](VERSION.md) 了解完整版本歷史。

## 技術棧

- **前端/框架**: Streamlit
- **資料處理**: Pandas, NumPy
- **Excel 產生**: OpenPyXL
- **部署**: Docker, Zeabur
