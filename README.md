# I14 報表轉換與分析系統

自動化 I14 庫存報表轉換與多維度數據分析工具，支援 **Overstock Review**（庫存過賸審查）、**OOS Review**（缺貨審查）、**Carry SKU Review**（在售品項盤點）、**SKU 精簡 Review**（ND SKU 退倉建議）與 **Safety Stock Review**（安全庫存審查）五種模式。

## 功能概述

| 功能 | 說明 |
|------|------|
| 檔案上傳與驗證 | 上傳 By SKUxShop Excel (.xlsx)，自動檢核必要欄位 |
| 模式選擇 | Overstock Review / OOS Review / Carry SKU Review / SKU 精簡 Review / Safety Stock Review / Stock Checking |
| Overstock 分析 | 篩選 P90D 零銷售及超過月均兩倍之庫存過賸品項 |
| OOS 分析 | 篩選缺貨（0 Stock）、潛在缺貨（Turnday < 7）及 Safety Qty 不足 |
| SKU 精簡分析 | 針對 RP Type = ND 的 SKU 建議退倉（Warehouse 可退倉、Vendor 另有安排），依 Supply Source × Shop Stock Status 雙層分類 |
| Stock Checking 分析 | D001 倉庫少貨且 Tourist 鋪 Turnday < 門檻 → 標示調貨需求，從 Mix / Non Tourist 鋪分配可轉出貨量（Non Tourist 優先） |
| 自定義參數 | Shop Nature 過濾、安全庫存倍率、Safety Qty 計算模式、調貨門檻 |
| 視覺化看板 | 即時顯示統計指標與交叉分析表 |
| Excel 導出 | 一鍵下載含 RAW、Shop Stock Status、By Shop、分析摘要之美化報表 |

## 系統架構

```
├── app.py              # Streamlit 主應用程式
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

| 欄位 | 說明 | 模式 |
|------|------|------|
| Buyer | 採購人員 | 全部 |
| Brand | 品牌名稱 | 全部 |
| SKU | 品號編碼 | 全部 |
| Shop Stock Status | 庫存狀態明細 | 全部 |

OOS Review 建議額外包含：

| 欄位 | 說明 |
|------|------|
| Sold Qty (CM - MTD) | 月度銷售數量 |
| Adj. Safety Stock | 現有安全庫存量 |
| Shop Nature | 店鋪性質（Tourist / Mix / Non Tourist） |

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

輸出包含 RAW、Shop Stock Status、By Shop、ND SKU 數據分析摘要共 4 個工作表。

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
輸出包含 RAW、Shop Stock Status、By Shop、SUMMARY、Safety Stock 數據分析摘要共 5 個工作表。
SUMMARY 工作表統計 Safety Qty Change 分類（增加 / 減少 / 不變）與 Recommendation of new Safety Qty / ND 分類（空白維持 / New SKU / 0 已是 ND / ND 系列加註）的行數、SKU 數與調整量合計。

### Stock Checking（調貨檢查）

上傳 `By SKUxShop_Stock Checking_<日期>.xlsx`（D001 倉庫少貨且遊客區店舖缺貨時的檢查），計算 Tourist 鋪 SKU 的調貨需求，並從 Mix / Non Tourist 鋪分配可轉出貨量：

1. **觸發條件（SKU 級，同時符合才標示）** — 該 SKU 的「Tourist 加總 TurnDay」（各 Tourist 店 Shop Stock Qty (CM) 加總 ÷ Shop Monthly Sold Qty (P90D) 加總 × 30.5，與 Summary 同口徑）< 門檻（預設 14 日，UI 可調），且 D001 Net Stock Qty (CM) < 門檻（預設 100 件，UI 可調），且 Mix / Non Tourist 至少一個 Nature 的加總 TurnDay ≥ 門檻（有貨可調）；Tourist 加總 TurnDay ≥ 門檻的 SKU 一律無須調貨；Mix 與 Non Tourist 加總 TurnDay 皆低於門檻時視為無貨可調、不標示。只有已標示 SKU 才在 RAW 計算調貨數量（個別 Tourist 店補到 14 日周轉的需求 > 0）
2. **調入需求量** = ceil(P90D / 90 × 14) − Shop Stock Qty (CM)，補到 14 日周轉
3. **可轉出貨量（來源店）** = max(0, Shop Stock Qty (CM) − P90D / 90 × 14)；「無銷售但有存貨」全數可釋放；RP Type 為 RF 與 ND 的店皆可作為來源；該 Nature 的 SKU 級加總 TurnDay < 門檻時視為無貨可調（可轉出量歸零）
4. **分配** — 按 SKU 彙總需求，Non Tourist 店優先轉出，不足由 Mix 店補足；需求量大→小依序分配

輸出包含 Summary、RAW 與調貨分析摘要三個工作表：

- **Summary** — SKU 級彙總（每 SKU 一列，按 Brand / SKU 排序）：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001 Net Stock Qty (CM)；各 Shop Nature（Tourist / Mix / Non Tourist）加總 Shop Stock Qty (CM) 與 Shop Monthly Sold Qty (P90D)（取整）、TurnDay = ROUND(加總庫存 ÷ 加總月銷 × 30.5, 0)（取整，月銷 0 留空）、各 Nature 佔比 (%) = 該 Nature 加總 ÷ 三 Nature 合計（庫存/月銷各 3 欄，活公式 + 0.00% 顯示，合計 0 留空）；強制調貨提示（該 SKU 有 Tourist 店需要調貨時標示）、由 Non Tourist / Mix 店舖預計調出數量；文字欄靠左、數值置中、間隔列底紋、凍結表頭 + 自動篩選
- **RAW** — 原始欄位全部保留並新增：Tourist / Mix / Non Tourist 加總 TurnDay / 需要調貨 / 調入需求量 (補到14日) / 實際調入量 / 調入缺口 / 可轉出貨量 (Mix) / 可轉出貨量 (Non Tourist) / 轉出量 / 調貨來源 (Site Code)
- **調貨分析摘要** — Dashboard 指標卡（與 Streamlit 看板同口徑）：需要調貨列數 / 需調貨 SKU 數 / 調入需求量 / 實際調入量 / Non Tourist 轉出量 / Mix 轉出量 / 無法足額缺口 / 總明細行數

## Excel 輸出結構

| 工作表 | 說明 |
|--------|------|
| RAW | 完整明細資料（含所有原始及計算欄位） |
| Shop Stock Status | 採購人員 x 品牌之庫存異常交叉統計表 |
| By Shop | 各店鋪層級之彙總分析 |
| SUMMARY | 分類統計 — Safety Qty Change 分類（增加/減少/不變）與 ND 建議分類（僅 Safety Stock Review） |
| 分析摘要 | 數據看板 — 關鍵指標總覽 |

## 版本

查看 [VERSION.md](VERSION.md) 了解完整版本歷史。

## 技術棧

- **前端/框架**: Streamlit
- **資料處理**: Pandas, NumPy
- **Excel 產生**: OpenPyXL
- **部署**: Docker, Zeabur
