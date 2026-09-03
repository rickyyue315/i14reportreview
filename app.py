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
/* 載入 Google 字型 */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+TC:wght@300;400;500;700&display=swap');

/* 全域字型與背景微調 (淺色主題) */
html, body, [class*="css"] {
    font-family: 'Inter', 'Noto Sans TC', -apple-system, BlinkMacSystemFont, sans-serif;
    color: #1a1a2e;
}

/* 主容器間距調整 */
.main .block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1200px;
}

/* 專業漸層 Header Banner (淺色) */
.header-container {
    background: linear-gradient(135deg, #e0f7ff 0%, #ffffff 100%);
    padding: 2.2rem;
    border-radius: 14px;
    color: #1a1a2e;
    margin-bottom: 2rem;
    box-shadow: 0 10px 25px -5px rgba(0, 80, 120, 0.12);
    border: 1px solid #00a3c433;
}
.header-title {
    font-size: 2rem;
    font-weight: 700;
    margin: 0;
    letter-spacing: -0.02em;
    display: flex;
    align-items: center;
    gap: 12px;
    color: #1a1a2e;
}
.header-subtitle {
    font-size: 0.95rem;
    color: #0077a3;
    margin-top: 8px;
    margin-bottom: 0;
    font-weight: 400;
}

/* 自定義精美卡片樣式 (淺色) */
.custom-card {
    background-color: #ffffff;
    border: 1px solid #e3e6ec;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.06), 0 2px 4px -1px rgba(0, 0, 0, 0.04);
    color: #1a1a2e;
}

/* 數據指標卡片 (Metric Cards) 淺色高對比 */
.metric-card {
    background-color: #ffffff;
    border: 1px solid #e3e6ec;
    border-radius: 10px;
    padding: 1.2rem;
    box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
    border-left: 5px solid #00a3c4;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
    color: #1a1a2e;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 16px -3px rgba(0, 0, 0, 0.1);
    border-color: #00a3c4;
}
.metric-title {
    font-size: 0.82rem;
    color: #4a4a5e;
    font-weight: 500;
    margin-bottom: 4px;
}
.metric-value {
    font-size: 1.6rem;
    color: #0077a3;
    font-weight: 700;
    line-height: 1.2;
}
.metric-desc {
    font-size: 0.82rem;
    color: #6a6a7a;
    margin-top: 4px;
}

/* Streamlit Tabs 頁籤樣式優化 (淺色) */
.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
    border-bottom: 1px solid #e3e6ec;
}
.stTabs [data-baseweb="tab"] {
    height: 44px;
    background-color: transparent;
    border-radius: 8px 8px 0px 0px;
    padding: 10px 16px;
    color: #5a5a6e;
    font-weight: 500;
    transition: all 0.2s;
}
.stTabs [data-baseweb="tab"]:hover {
    color: #0077a3;
    background-color: #f0f8fb;
}
.stTabs [aria-selected="true"] {
    color: #0077a3 !important;
    border-bottom: 2px solid #0077a3 !important;
    font-weight: 700 !important;
}

/* 針對 Streamlit 內建按鈕的微調 (淺色) */
button[kind="primary"] {
    background: linear-gradient(135deg, #00a3c4 0%, #0077a3 100%) !important;
    border: none !important;
    color: #ffffff !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    padding: 0.55rem 1.4rem !important;
    transition: all 0.2s ease !important;
}
button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(0, 119, 163, 0.35) !important;
}

/* 檔案上傳框外觀優化 (淺色) */
[data-testid="stFileUploader"] {
    border: 2px dashed #00a3c466;
    border-radius: 12px;
    padding: 1.2rem;
    background-color: #f8fbfc;
    color: #1a1a2e;
}

/* 自定義漸層分隔線 */
.custom-divider {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(0,119,163,0.15) 20%, rgba(0,119,163,0.15) 80%, transparent);
    margin: 1.5rem 0;
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
