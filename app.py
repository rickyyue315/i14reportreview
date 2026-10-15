import streamlit as st

# 1. 設置網頁基本資訊（多頁架構：本檔為導航入口，所有頁面共用頁面設定與 CSS）
st.set_page_config(
    page_title="I14 報表轉換與分析系統",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# 2. 注入自定義進階 CSS 樣式以實現淺色專業配色（所有頁面共用）
st.markdown("""
<style>
/* 載入 Google 字型（含等寬數字字型 JetBrains Mono） */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500;600;700&family=Noto+Sans+TC:wght@400;500;600;700&display=swap');

/* 全域字型與背景微調 (淺色專業主題) */
html, body, [class*="css"] {
    font-family: 'Inter', 'Noto Sans TC', -apple-system, BlinkMacSystemFont, sans-serif;
    color: #0f172a;
    -webkit-font-smoothing: antialiased;
}

.stApp {
    background-color: #f8fafc;
}

/* 主容器間距調整 */
.main .block-container {
    padding-top: 1.5rem;
    padding-bottom: 3.5rem;
    max-width: 1340px;
}

/* 側邊欄外觀優化 */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #ffffff 0%, #f1f5f9 100%);
    border-right: 1px solid #e2e8f0;
}
[data-testid="stSidebarNav"] {
    padding-top: 0.75rem;
}
[data-testid="stSidebarNav"] a {
    border-radius: 8px;
    margin: 2px 8px;
    font-weight: 500;
    transition: all 0.18s ease;
}
[data-testid="stSidebarNav"] a[aria-current="page"] {
    background-color: #e0f2fe !important;
    color: #0369a1 !important;
    font-weight: 700 !important;
}

/* 側邊欄品牌卡片與底部狀態卡 */
.sidebar-brand-card {
    background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 100%);
    border: 1px solid #bae6fd;
    border-radius: 12px;
    padding: 1rem 0.9rem;
    text-align: center;
    margin-bottom: 0.5rem;
    box-shadow: 0 2px 6px rgba(2, 132, 199, 0.08);
}
.sidebar-brand-title {
    color: #0369a1;
    margin: 0;
    font-weight: 700;
    font-size: 1.35rem;
    letter-spacing: -0.02em;
}
.sidebar-brand-sub {
    color: #475569;
    font-size: 0.8rem;
    margin: 4px 0 0 0;
    font-weight: 500;
}
.sidebar-footer-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 0.7rem 0.8rem;
    margin-top: 0.6rem;
    color: #475569;
    font-size: 0.75rem;
    text-align: center;
    line-height: 1.6;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03);
}

/* 專業漸層 Header Banner (淺色高質感) */
.header-container {
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, #ffffff 0%, #f0f9ff 55%, #e0f2fe 100%);
    padding: 1.85rem 2.2rem;
    border-radius: 16px;
    color: #0f172a;
    margin-bottom: 1.5rem;
    box-shadow: 0 10px 25px -5px rgba(2, 132, 199, 0.1), 0 4px 10px -2px rgba(15, 23, 42, 0.03);
    border: 1px solid #bae6fd;
    border-left: 6px solid #0284c7;
}
.header-badge {
    display: inline-block;
    background: rgba(2, 132, 199, 0.1);
    color: #0369a1;
    border: 1px solid rgba(2, 132, 199, 0.22);
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.06em;
    padding: 3px 10px;
    border-radius: 999px;
    margin-bottom: 8px;
}
.header-title {
    font-size: 1.85rem;
    font-weight: 700;
    margin: 0;
    letter-spacing: -0.02em;
    display: flex;
    align-items: center;
    gap: 10px;
    color: #0f172a;
}
.header-subtitle {
    font-size: 0.92rem;
    color: #334155;
    margin-top: 8px;
    margin-bottom: 0;
    font-weight: 400;
    line-height: 1.5;
}

/* 區塊標題橫幅 (取代原跨 markdown 空 div) */
.section-banner {
    background: linear-gradient(90deg, #f0f9ff 0%, #ffffff 100%);
    border: 1px solid #e2e8f0;
    border-left: 5px solid #0284c7;
    border-radius: 12px;
    padding: 0.85rem 1.25rem;
    margin: 1.25rem 0 1rem 0;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.03);
}
.section-banner h3 {
    margin: 0 !important;
    padding: 0 !important;
    font-size: 1.2rem !important;
    font-weight: 700 !important;
    color: #0f172a !important;
    letter-spacing: -0.01em;
}

/* 自定義精美卡片樣式 (單一 HTML 區塊使用) */
.custom-card {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 1.5rem 1.75rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 14px -2px rgba(15, 23, 42, 0.05), 0 2px 4px -1px rgba(15, 23, 42, 0.03);
    color: #0f172a;
}

/* 數據指標卡片 (Metric Cards) 淺色高對比 */
.metric-card {
    background: linear-gradient(180deg, #ffffff 0%, #fcfdff 100%);
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1.15rem 1.25rem;
    box-shadow: 0 2px 8px -2px rgba(15, 23, 42, 0.05), 0 1px 3px 0 rgba(15, 23, 42, 0.03);
    border-left: 5px solid #0284c7;
    transition: transform 0.2s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.2s ease;
    color: #0f172a;
    height: 100%;
}
.metric-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 12px 20px -5px rgba(15, 23, 42, 0.1), 0 4px 8px -2px rgba(15, 23, 42, 0.04);
}
.metric-title {
    font-size: 0.82rem;
    color: #475569;
    font-weight: 600;
    margin-bottom: 6px;
    letter-spacing: 0.01em;
}
.metric-value {
    font-family: 'JetBrains Mono', 'Inter', 'Noto Sans TC', sans-serif;
    font-size: 1.65rem;
    color: #0284c7;
    font-weight: 700;
    line-height: 1.2;
    letter-spacing: -0.02em;
    font-variant-numeric: tabular-nums;
}
.metric-desc {
    font-size: 0.78rem;
    color: #64748b;
    margin-top: 6px;
    line-height: 1.4;
}

/* Streamlit Tabs 頁籤樣式優化 (膠囊分段列風格) */
.stTabs [data-baseweb="tab-list"] {
    gap: 6px;
    background-color: #e2e8f080;
    padding: 6px;
    border-radius: 12px;
    border-bottom: none;
    flex-wrap: wrap;
}
.stTabs [data-baseweb="tab"] {
    height: 38px;
    background-color: transparent;
    border-radius: 8px;
    padding: 6px 14px;
    color: #475569;
    font-size: 0.88rem;
    font-weight: 500;
    border: none !important;
    transition: all 0.18s ease;
}
.stTabs [data-baseweb="tab"]:hover {
    color: #0284c7;
    background-color: rgba(255, 255, 255, 0.65);
}
.stTabs [aria-selected="true"] {
    background-color: #ffffff !important;
    color: #0369a1 !important;
    font-weight: 700 !important;
    box-shadow: 0 1px 4px rgba(15, 23, 42, 0.1) !important;
}
.stTabs [data-baseweb="tab-highlight"] {
    display: none !important;
}

/* 針對 Streamlit 內建按鈕的微調 */
button[kind="primary"], [data-testid="stDownloadButton"] button[kind="primary"] {
    background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;
    border: 1px solid #0369a1 !important;
    color: #ffffff !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    padding: 0.6rem 1.4rem !important;
    box-shadow: 0 2px 6px rgba(2, 132, 199, 0.25) !important;
    transition: all 0.2s ease !important;
}
button[kind="primary"]:hover:not(:disabled), [data-testid="stDownloadButton"] button[kind="primary"]:hover:not(:disabled) {
    transform: translateY(-1px) !important;
    box-shadow: 0 6px 16px rgba(2, 132, 199, 0.35) !important;
    filter: brightness(1.04);
}
button[kind="secondary"] {
    background-color: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    color: #334155 !important;
    font-weight: 600 !important;
    border-radius: 10px !important;
    transition: all 0.18s ease !important;
}
button[kind="secondary"]:hover:not(:disabled) {
    border-color: #0284c7 !important;
    color: #0369a1 !important;
    background-color: #f0f9ff !important;
}

/* 檔案上傳框外觀優化 */
[data-testid="stFileUploader"] {
    border: 2px dashed #38bdf8;
    border-radius: 14px;
    padding: 1.3rem 1.5rem;
    background: linear-gradient(180deg, #ffffff 0%, #f0f9ff 100%);
    color: #0f172a;
    box-shadow: 0 2px 10px rgba(2, 132, 199, 0.05);
    transition: all 0.2s ease;
}
[data-testid="stFileUploader"]:hover {
    border-color: #0284c7;
    box-shadow: 0 4px 16px rgba(2, 132, 199, 0.1);
}

/* 摺疊區塊 (Expander) 與資料表 (DataFrame) 優化 */
[data-testid="stExpander"] {
    background-color: #ffffff;
    border: 1px solid #e2e8f0 !important;
    border-radius: 12px !important;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
    overflow: hidden;
    margin-bottom: 0.75rem;
}
[data-testid="stExpander"] summary {
    font-weight: 600;
    color: #1e293b;
}
[data-testid="stExpander"] summary:hover {
    color: #0284c7;
}
[data-testid="stDataFrame"] {
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    overflow: hidden;
    background-color: #ffffff;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.03);
}

/* 自定義漸層分隔線 */
.custom-divider {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(2, 132, 199, 0.2) 20%, rgba(2, 132, 199, 0.2) 80%, transparent);
    margin: 1.4rem 0;
}

/* ---- 介面優化：步驟條（首頁流程可視化，純顯示，不影響狀態） ---- */
.steps-bar {
    display: flex;
    gap: 0;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    overflow: hidden;
    margin-bottom: 1rem;
    box-shadow: 0 2px 6px rgba(15, 23, 42, 0.03);
}
.steps-bar .step {
    flex: 1;
    padding: 0.75rem 1rem;
    font-size: 0.84rem;
    border-right: 1px solid #f1f5f9;
    color: #1e293b;
    transition: background-color 0.18s ease;
}
.steps-bar .step:hover {
    background-color: #f8fafc;
}
.steps-bar .step:last-child { border-right: none; }
.steps-bar .step .n {
    display: inline-block;
    width: 24px;
    height: 24px;
    border-radius: 50%;
    background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%);
    color: #fff;
    text-align: center;
    font-size: 0.75rem;
    line-height: 24px;
    margin-right: 8px;
    font-weight: 700;
    box-shadow: 0 2px 4px rgba(2, 132, 199, 0.25);
}
.steps-bar .step small {
    display: block;
    color: #64748b;
    margin-top: 3px;
    margin-left: 32px;
    font-size: 0.74rem;
}

/* ---- 模式分類導覽列與當前模式說明卡 ---- */
.mode-category-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 8px;
    margin-bottom: 0.85rem;
}
@media (max-width: 992px) {
    .mode-category-grid { grid-template-columns: repeat(2, 1fr); }
}
.mode-cat-chip {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 0.55rem 0.75rem;
    font-size: 0.76rem;
    color: #475569;
    line-height: 1.45;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.02);
}
.mode-cat-chip b {
    display: block;
    color: #0f172a;
    font-size: 0.8rem;
    margin-bottom: 2px;
}
.mode-desc-banner {
    background: linear-gradient(90deg, #f0f9ff 0%, #ffffff 100%);
    border: 1px solid #bae6fd;
    border-left: 4px solid #0284c7;
    border-radius: 10px;
    padding: 0.65rem 1rem;
    margin-top: 0.35rem;
    margin-bottom: 1rem;
    font-size: 0.86rem;
    color: #1e293b;
    line-height: 1.5;
}
.mode-desc-banner .mode-tag {
    display: inline-block;
    background: #0284c7;
    color: #ffffff;
    font-weight: 600;
    font-size: 0.76rem;
    padding: 2px 9px;
    border-radius: 999px;
    margin-right: 8px;
    vertical-align: middle;
}

/* ---- 介面優化：橫向單選按鈕 (Radio Group) 膠囊卡片化 ---- */
div[data-testid="stRadio"] > div[role="radiogroup"] {
    flex-wrap: wrap !important;
    gap: 8px !important;
    padding: 4px 0;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label {
    background: #ffffff !important;
    border: 1px solid #cbd5e1 !important;
    border-radius: 999px !important;
    padding: 5px 14px 5px 10px !important;
    font-size: 0.85rem !important;
    font-weight: 500 !important;
    color: #334155 !important;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.03) !important;
    transition: all 0.18s ease !important;
    margin-right: 0 !important;
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label:hover {
    border-color: #38bdf8 !important;
    background-color: #f0f9ff !important;
    transform: translateY(-1px);
}
div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {
    background: linear-gradient(135deg, #e0f2fe 0%, #bae6fd 100%) !important;
    border-color: #0284c7 !important;
    color: #0369a1 !important;
    font-weight: 700 !important;
    box-shadow: 0 2px 6px rgba(2, 132, 199, 0.18) !important;
}

/* ---- 使用說明頁：步驟網格與欄位標籤 ---- */
.usage-steps-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 10px;
    margin-top: 0.85rem;
}
@media (max-width: 992px) {
    .usage-steps-grid { grid-template-columns: repeat(2, 1fr); }
}
.usage-step-card {
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-top: 3px solid #0284c7;
    border-radius: 10px;
    padding: 0.85rem 0.9rem;
    font-size: 0.82rem;
    color: #334155;
    line-height: 1.5;
}
.usage-step-card b {
    display: block;
    color: #0f172a;
    font-size: 0.9rem;
    margin-bottom: 4px;
}
.col-badge {
    display: inline-block;
    font-family: 'JetBrains Mono', monospace;
    background: #f0f9ff;
    color: #0369a1;
    border: 1px solid #bae6fd;
    border-radius: 6px;
    padding: 1px 7px;
    font-size: 0.8rem;
    font-weight: 600;
}

/* ---- 介面優化：頂部下載列（結果頁置頂，免捲動） ---- */
.result-topbar {
    position: sticky;
    top: 0;
    z-index: 5;
    background: rgba(255, 255, 255, 0.96);
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 0.6rem 0.9rem;
    margin-bottom: 1rem;
    box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
}
</style>
""", unsafe_allow_html=True)

# 3. 多頁導航（切換頁面不會清空已產生的分析結果，widget 狀態以 session_state 保留）
pages = [
    st.Page("views/analyzer_page.py", title="分析工具", icon="📊", default=True),
    st.Page("views/usage_page.py", title="使用說明", icon="📖"),
]
pg = st.navigation(pages)
pg.run()
