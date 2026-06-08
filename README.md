# I14 報表轉換與分析系統

自動化 I14 庫存報表轉換與多維度數據分析工具，支援 **Overstock Review**（庫存過賸審查）、**OOS Review**（缺貨審查）、**Carry SKU Review**（在售品項盤點）與 **Safety Stock Review**（安全庫存審查）四種模式。

## 功能概述

| 功能 | 說明 |
|------|------|
| 檔案上傳與驗證 | 上傳 By SKUxShop Excel (.xlsx)，自動檢核必要欄位 |
| 模式選擇 | Overstock Review / OOS Review / Carry SKU Review / Safety Stock Review |
| Overstock 分析 | 篩選 P90D 零銷售及超過月均兩倍之庫存過賸品項 |
| OOS 分析 | 篩選缺貨（0 Stock）、潛在缺貨（Turnday < 7）及 Safety Qty 不足 |
| 自定義參數 | Shop Nature 過濾、安全庫存倍率、Safety Qty 計算模式 |
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

專案已預設 `zeabur.yaml`，可直接連接 GitHub 倉庫自動部署。

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

1. **Overstock (P90D 0 Sales)** — 過去 90 天內零銷售
2. **Overstock (>2\*Monthly Sold Qty & >2pcs)** — 庫存超過月均銷售兩倍且大於 2 件

系統自動計算新的 Safety Qty 建議，並產出可轉出貨量。

### OOS Review

針對 `Shop Stock Status` 欄位篩選以下缺貨條件：

1. **OOS (0 Stock)** — 現有庫存為 0
2. **Potential OOS (<7 Turnday)** — Turnday 不足 7 天
3. **Normal (Safety Qty Insufficient)** — 庫存正常但 Safety Qty 不足以覆蓋預估銷售

支援 MTD 日期選擇，自動推算月度預估銷售量（Projected Monthly Sold Qty）。

### Safety Stock Review

全面審查各 SKU×Shop 的安全庫存設定，結合 Overstock Cap 與 OOS Floor 計算平衡建議值：

1. **Overstock Cap (X)** — 基於月均銷售與 Overstock Rate 計算庫存上限，上限不超過現有 Adj. Safety Stock
2. **OOS Floor (Y)** — 基於 MTD 預估銷售與 OOS Rate 計算缺貨底線，最低 3
3. **建議 Safety Stock** = max(Y, Z if Z ≤ X else X)，其中 Z = 現有 Adj. Safety Stock
4. **Safety Qty Change** = 建議值 − Z（正數=增加、負數=減少、零=不變）

品項分為 Overstock / OOS / Balanced (Safety Adjusted) 三類，標示調整原因。
輸出包含 RAW、Shop Stock Status、By Shop、Safety Stock 數據分析摘要共 4 個工作表。

## Excel 輸出結構

| 工作表 | 說明 |
|--------|------|
| RAW | 完整明細資料（含所有原始及計算欄位） |
| Shop Stock Status | 採購人員 x 品牌之庫存異常交叉統計表 |
| By Shop | 各店鋪層級之彙總分析 |
| 分析摘要 | 數據看板 — 關鍵指標總覽 |

## 版本

查看 [VERSION.md](VERSION.md) 了解完整版本歷史。

## 技術棧

- **前端/框架**: Streamlit
- **資料處理**: Pandas, NumPy
- **Excel 產生**: OpenPyXL
- **部署**: Docker, Zeabur
