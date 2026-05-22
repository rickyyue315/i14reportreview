import streamlit as st
import pandas as pd
from datetime import datetime
from zoneinfo import ZoneInfo
from transformer import transform, generate_excel, validate_input, _resolve_column_name
import time
from openpyxl import load_workbook
from io import BytesIO
from version import __version__

# 1. 設置網頁基本資訊
st.set_page_config(
    page_title="I14 報表轉換與分析系統",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. 注入自定義進階 CSS 樣式以實現極簡專業配色（含明暗主題變數）
st.markdown("""
<style>
/* 載入 Google 字型 */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+TC:wght@300;400;500;700&display=swap');

/* ========== 日間模式 (Light) CSS 變數 ========== */
:root {
    --bg-primary: #F8FAFC;
    --bg-secondary: #F1F5F9;
    --bg-card: #FFFFFF;
    --bg-file-uploader: #F8FAFC;
    --bg-download: #F8FAFC;
    --bg-tab-hover: #F8FAFC;
    --bg-error: #FDF2F2;
    --text-primary: #0F172A;
    --text-secondary: #64748B;
    --text-muted: #94A3B8;
    --text-download-desc: #475569;
    --border-primary: #E2E8F0;
    --border-accent: #CBD5E1;
    --border-download: #1E40AF;
    --border-error: #F8B4B4;
    --accent-primary: #1E40AF;
    --accent-primary-light: #2563EB;
    --accent-secondary: #93C5FD;
    --shadow-card: 0 4px 6px -1px rgba(0,0,0,0.02), 0 2px 4px -1px rgba(0,0,0,0.01);
    --shadow-card-hover: 0 8px 16px -3px rgba(0,0,0,0.04);
    --shadow-header: 0 10px 25px -5px rgba(30,64,175,0.12);
    --shadow-button: 0 4px 12px rgba(37,99,235,0.25);
    --header-gradient-start: #1E40AF;
    --header-gradient-end: #0F172A;
    --header-text: #FFFFFF;
    --header-subtitle: #93C5FD;
    --button-gradient-start: #1E40AF;
    --button-gradient-end: #2563EB;
    --button-text: white;
    --error-heading: #9B1C1C;
    --error-desc: #771D1D;
    --metric-green: #059669;
    --metric-red: #DC2626;
    --tab-text: #64748B;
    --tab-active: #1E40AF;
    --sidebar-title: #1E40AF;
    --sidebar-subtitle: #64748B;
    --sidebar-small: #94A3B8;
    --footer-text: #94A3B8;
    --toggle-bg: #FFFFFF;
    --toggle-border: #E2E8F0;
}

/* ========== 夜間模式 (Dark) CSS 變數 — Kilo.ai 配色風格 ========== */
html.dark {
    --bg-primary: #0A0A0A;
    --bg-secondary: #121212;
    --bg-card: #1E1E1E;
    --bg-file-uploader: #141414;
    --bg-download: #141414;
    --bg-tab-hover: #1E1E1E;
    --bg-error: #2D1B1B;
    --text-primary: #E5E5E5;
    --text-secondary: #9CA3AF;
    --text-muted: #6B7280;
    --text-download-desc: #9CA3AF;
    --border-primary: #333333;
    --border-accent: #444444;
    --border-download: #5B8FEF;
    --border-error: #662222;
    --accent-primary: #5B8FEF;
    --accent-primary-light: #7BA5F7;
    --accent-secondary: #7BA5F7;
    --shadow-card: 0 4px 6px -1px rgba(0,0,0,0.3), 0 2px 4px -1px rgba(0,0,0,0.2);
    --shadow-card-hover: 0 8px 16px -3px rgba(0,0,0,0.4);
    --shadow-header: 0 10px 25px -5px rgba(0,0,0,0.4);
    --shadow-button: 0 4px 12px rgba(91,143,239,0.3);
    --header-gradient-start: #162240;
    --header-gradient-end: #0A0A0A;
    --header-text: #E5E5E5;
    --header-subtitle: #7BA5F7;
    --button-gradient-start: #5B8FEF;
    --button-gradient-end: #7BA5F7;
    --button-text: white;
    --error-heading: #F87171;
    --error-desc: #FCA5A5;
    --metric-green: #34D399;
    --metric-red: #F87171;
    --tab-text: #9CA3AF;
    --tab-active: #5B8FEF;
    --sidebar-title: #5B8FEF;
    --sidebar-subtitle: #9CA3AF;
    --sidebar-small: #6B7280;
    --footer-text: #6B7280;
    --toggle-bg: #1E1E1E;
    --toggle-border: #333333;
}

/* ========== Streamlit 原生元素夜間模式適配 ========== */
html.dark .stApp {
    background-color: var(--bg-primary);
}
html.dark section[data-testid="stSidebar"] {
    background-color: #0E0E0E;
}
html.dark section[data-testid="stSidebar"] * {
    color: var(--text-primary);
}
html.dark .stMarkdown, html.dark .stMarkdown p, html.dark .stMarkdown span,
html.dark h1, html.dark h2, html.dark h3, html.dark h4, html.dark h5, html.dark h6 {
    color: var(--text-primary);
}
html.dark .stExpander {
    background-color: var(--bg-card);
    border-color: var(--border-primary);
}
html.dark .stExpander p, html.dark .stExpander span {
    color: var(--text-primary);
}
html.dark hr, html.dark .stDivider {
    border-color: var(--border-primary) !important;
}
html.dark .stProgress > div > div {
    background-color: var(--border-primary);
}
html.dark [data-testid="stDataFrame"] {
    color: var(--text-primary);
}
html.dark .st-emotion-cache-1qg05tj {
    color: var(--text-primary);
}
html.dark .st-cb, html.dark .st-cm, html.dark .st-cd, html.dark .st-cg, html.dark .st-ch {
    color: var(--text-primary);
}
html.dark .stDataFrame table, html.dark .stTable table {
    color: var(--text-primary);
}

/* ========== 全域字型與背景微調 ========== */
html, body, [class*="css"] {
    font-family: 'Inter', 'Noto Sans TC', -apple-system, BlinkMacSystemFont, sans-serif;
}

.main .block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1200px;
}

/* ========== 專業漸層 Header Banner ========== */
.header-container {
    background: linear-gradient(135deg, var(--header-gradient-start) 0%, var(--header-gradient-end) 100%);
    padding: 2.2rem;
    border-radius: 14px;
    color: var(--header-text);
    margin-bottom: 2rem;
    box-shadow: var(--shadow-header);
}
.header-title {
    font-size: 2rem;
    font-weight: 700;
    margin: 0;
    letter-spacing: -0.02em;
    display: flex;
    align-items: center;
    gap: 12px;
}
.header-subtitle {
    font-size: 0.95rem;
    color: var(--header-subtitle);
    margin-top: 8px;
    margin-bottom: 0;
    font-weight: 300;
}

/* ========== 自定義精美卡片樣式 ========== */
.custom-card {
    background-color: var(--bg-card);
    border: 1px solid var(--border-primary);
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: var(--shadow-card);
}

/* ========== 數據指標卡片 (Metric Cards) ========== */
.metric-card {
    background-color: var(--bg-card);
    border: 1px solid var(--border-primary);
    border-radius: 10px;
    padding: 1.2rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    border-left: 5px solid var(--accent-primary);
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: var(--shadow-card-hover);
}
.metric-title {
    font-size: 0.82rem;
    color: var(--text-secondary);
    font-weight: 500;
    margin-bottom: 4px;
}
.metric-value {
    font-size: 1.6rem;
    color: var(--text-primary);
    font-weight: 700;
    line-height: 1.2;
}
.metric-desc {
    font-size: 0.72rem;
    color: var(--text-muted);
    margin-top: 4px;
}

/* ========== Streamlit Tabs 頁籤樣式優化 ========== */
.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
    border-bottom: 1px solid var(--border-primary);
}
.stTabs [data-baseweb="tab"] {
    height: 44px;
    background-color: transparent;
    border-radius: 8px 8px 0px 0px;
    padding: 10px 16px;
    color: var(--tab-text);
    font-weight: 500;
    transition: all 0.2s;
}
.stTabs [data-baseweb="tab"]:hover {
    color: var(--accent-primary);
    background-color: var(--bg-tab-hover);
}
.stTabs [aria-selected="true"] {
    color: var(--tab-active) !important;
    border-bottom: 2px solid var(--tab-active) !important;
    font-weight: 700 !important;
}

/* ========== Streamlit 內建按鈕微調 ========== */
button[kind="primary"] {
    background: linear-gradient(135deg, var(--button-gradient-start) 0%, var(--button-gradient-end) 100%) !important;
    border: none !important;
    color: var(--button-text) !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    padding: 0.55rem 1.4rem !important;
    transition: all 0.2s ease !important;
}
button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: var(--shadow-button) !important;
}

/* ========== 檔案上傳框外觀優化 ========== */
[data-testid="stFileUploader"] {
    border: 2px dashed var(--border-accent);
    border-radius: 12px;
    padding: 1.2rem;
    background-color: var(--bg-file-uploader);
}

/* ========== 主題切換按鈕樣式 ========== */
#theme-toggle-btn {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: var(--toggle-bg);
    color: var(--text-primary);
    border: 1px solid var(--toggle-border);
    border-radius: 20px;
    padding: 8px 18px;
    cursor: pointer;
    font-size: 0.85rem;
    font-family: 'Inter', 'Noto Sans TC', sans-serif;
    transition: all 0.2s ease;
    white-space: nowrap;
}
#theme-toggle-btn:hover {
    border-color: var(--accent-primary);
    color: var(--accent-primary);
}
</style>

<script>
function toggleTheme() {
    const html = document.documentElement;
    const isDark = html.classList.toggle('dark');
    localStorage.setItem('kilo-theme', isDark ? 'dark' : 'light');
    updateToggleButton(isDark);
}
function updateToggleButton(isDark) {
    const btn = document.getElementById('theme-toggle-btn');
    if (btn) {
        btn.innerHTML = isDark ? '☀️ 日間模式' : '🌙 夜間模式';
    }
}
(function() {
    const saved = localStorage.getItem('kilo-theme');
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    const isDark = saved === 'dark' || (!saved && prefersDark);
    if (isDark) document.documentElement.classList.add('dark');
    function init() { updateToggleButton(isDark); }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
</script>
""", unsafe_allow_html=True)

# 3. 側邊欄設計 (Sidebar)
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 15px 0 5px 0;">
        <h2 style="color: var(--sidebar-title); margin-bottom: 2px; font-weight: 700; font-size: 1.6rem;">I14 Analyzer</h2>
        <p style="color: var(--sidebar-subtitle); font-size: 0.85rem; margin-top: 0;">自動化報表與數據轉換引擎</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    st.markdown("""
    <div style="text-align: center; margin: 6px 0 12px 0;">
        <button id="theme-toggle-btn" onclick="toggleTheme()">🌙 夜間模式</button>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 📋 系統操作指南")
    st.markdown("""
    1. 📥 **準備來源檔案**：確認您的 Excel 檔案（By SKUxShop）包含下方檢核表所列之必要欄位。
    2. 📂 **上傳檔案**：將 Excel 拖曳至右側上傳區域。
    3. 🔍 **預覽確認**：系統解析完成後，可於主介面預覽前 20 筆資料，核對檔案結構。
    4. ⚙️ **執行轉換**：點擊「執行報表轉換與分析」，系統將依特定規則過濾、計算並自動產生彙總表。
    5. 💾 **下載成果**：於視覺化看板中即時查閱統計資訊，並一鍵導出美化後的 Excel 報表。
    """)
    
    st.divider()
    
    st.markdown("### 🔍 必要欄位檢核標準")
    st.markdown("""
    為確保計算正確，來源檔案必須包含以下欄位：
    - ✅ **Buyer** <span style="color: var(--sidebar-subtitle);font-size:0.8em">(採購人員)</span>
    - ✅ **Brand** <span style="color: var(--sidebar-subtitle);font-size:0.8em">(品牌名稱)</span>
    - ✅ **SKU** <span style="color: var(--sidebar-subtitle);font-size:0.8em">(品號編碼)</span>
    - ✅ **Shop Stock Status** <span style="color: var(--sidebar-subtitle);font-size:0.8em">(庫存狀態明細)</span>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    st.markdown("### ⚙️ 核心計算規則")
    st.markdown("""
    系統會針對 `Shop Stock Status` 欄位過濾以下庫存過賸（Overstock）條件並歸類：
    1. **Overstock (P90D 0 Sales)**
    2. **Overstock (>2*Monthly Qty & >2pcs)**
    3. 自動生成採購與品牌的交叉匯總樞紐分析表。
    """)
    
    st.divider()
    st.markdown(
        '<div style="color: var(--sidebar-small); font-size: 0.75rem; text-align: center; line-height: 1.5;">'
        f'I14 報表計算工具 v{__version__}<br>'
        'Developed by Ricky Yue<br>'
        'Powered by Streamlit & Zeabur'
        '</div>',
        unsafe_allow_html=True
    )

# 4. 主介面標頭 (Header)
st.markdown("""
<div class="header-container">
    <h1 class="header-title">📊 I14 報表轉換與分析系統</h1>
    <p class="header-subtitle">上傳 By SKUxShop 來源檔案，一鍵完成庫存異常篩選、多維度統計，並導出 I14 by SKUxShop_To Buyer 格式報表。</p>
</div>
""", unsafe_allow_html=True)

# 5. 初始化 Session State (確保下載與重整時狀態持久化)
if "processed" not in st.session_state:
    st.session_state.processed = False
    st.session_state.summary_df = None
    st.session_state.raw_df = None
    st.session_state.categories = None
    st.session_state.excel_data = None
    st.session_state.last_uploaded_name = None

# 6. 檔案上傳區域
uploaded_file = st.file_uploader(
    "📂 上傳 By SKUxShop 原始 Excel 檔案 (.xlsx)",
    type=["xlsx"],
    help="支援標準 Excel (.xlsx) 格式，檔案上限 200MB，最佳化支援 5 萬行以上大數據處理"
)

# 7. 處理上傳的檔案
if uploaded_file is not None:
    # 當上傳了新檔案，自動重設先前的處理狀態
    if st.session_state.last_uploaded_name != uploaded_file.name:
        st.session_state.last_uploaded_name = uploaded_file.name
        st.session_state.processed = False
        st.session_state.summary_df = None
        st.session_state.raw_df = None
        st.session_state.categories = None
        st.session_state.excel_data = None

    # 若尚未載入資料，則執行讀取
    if "df_input" not in st.session_state or st.session_state.last_uploaded_name != uploaded_file.name:
        _prog = st.progress(0, text="讀取檔案中... 0% | 初始化中...")
        _t0 = time.time()
        try:
            _raw = uploaded_file.read()
            _wb = load_workbook(BytesIO(_raw), data_only=True)
            _ws = _wb.active
            _n = _ws.max_row
            if _n <= 1:
                df_input = pd.DataFrame()
            else:
                _hdr = list(next(_ws.iter_rows(min_row=1, max_row=1, values_only=True)))
                _rows = []
                for _i, _r in enumerate(_ws.iter_rows(min_row=2, max_row=_n, values_only=True), start=1):
                    _rows.append(list(_r))
                    if _i % 500 == 0 or _i == _n - 1:
                        _pct = _i / (_n - 1)
                        _elapsed = time.time() - _t0
                        _eta = (_elapsed / _pct) - _elapsed if _pct > 0 else 0
                        _prog.progress(min(_pct, 1.0), text=f"讀取檔案中... {_pct*100:.0f}% | 已用: {_elapsed:.0f}s | 預計剩餘: {_eta:.0f}s")
                df_input = pd.DataFrame(_rows, columns=_hdr)
            st.session_state.df_input = df_input
        except Exception as e:
            _prog.empty()
            st.error(f"❌ 讀取檔案失敗: {e}")
            st.stop()
        _prog.empty()
    else:
        df_input = st.session_state.df_input

    # 必要欄位驗證
    missing = validate_input(df_input)
    if missing:
        st.markdown(f"""
        <div style="background-color: var(--bg-error); border: 1px solid var(--border-error); padding: 15px; border-radius: 8px; margin-bottom: 1.5rem;">
            <h4 style="color: var(--error-heading); margin-top: 0; margin-bottom: 8px;">❌ 來源檔案欄位檢驗未通過</h4>
            <p style="color: var(--error-desc); margin-bottom: 0; font-size: 0.9rem;">
                檔案內缺少必要欄位：<b>{', '.join(missing)}</b>。<br>
                請確認您上傳的 Excel 檔案包含以下欄位：Buyer, Brand, SKU, Shop Stock Status。
            </p>
        </div>
        """, unsafe_allow_html=True)
        st.stop()

    # 顯示上傳檔案基本資訊與預覽卡片
    st.markdown('<div class="custom-card">', unsafe_allow_html=True)
    st.markdown("### 📊 來源檔案成功載入")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(f"""
        <div class="metric-card" style="border-left-color: #10B981;">
            <div class="metric-title">檔案解析狀態</div>
            <div class="metric-value" style="color: var(--metric-green);">驗證通過</div>
            <div class="metric-desc">檔案格式符合解析標準</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card" style="border-left-color: #3B82F6;">
            <div class="metric-title">總資料筆數</div>
            <div class="metric-value">{len(df_input):,}</div>
            <div class="metric-desc">原始資料列數 (Rows)</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-card" style="border-left-color: #F59E0B;">
            <div class="metric-title">欄位總數</div>
            <div class="metric-value">{len(df_input.columns)}</div>
            <div class="metric-desc">原始資料欄數 (Columns)</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
    
    # 預覽原始資料摺疊區塊
    with st.expander("🔍 預覽前 20 筆原始資料", expanded=False):
        st.dataframe(df_input.head(20), use_container_width=True)
        st.caption(f"目前顯示第 1 至第 20 筆明細（總共 {len(df_input):,} 筆原始資料）")
    st.markdown('</div>', unsafe_allow_html=True)

    # 執行報表轉換按鈕
    st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)
    if st.button("⚙️ 執行報表轉換與分析", type="primary", use_container_width=True):
        _prog2 = st.progress(0, text="處理中... 0% | 開始處理...")
        _t0_2 = time.time()
        
        def _cb(pct, msg):
            _e = time.time() - _t0_2
            _eta2 = (_e / pct) - _e if pct > 0 else 0
            _prog2.progress(min(pct, 1.0), text=f"處理中... {pct*100:.0f}% | {msg} | 已用: {_e:.0f}s | 預計剩餘: {_eta2:.0f}s")
            
        try:
            summary_df, raw_df, categories = transform(df_input, progress_callback=_cb)
            buf = generate_excel(summary_df, raw_df, categories, progress_callback=_cb)
            
            # 將處理結果儲存至 session state
            st.session_state.processed = True
            st.session_state.summary_df = summary_df
            st.session_state.raw_df = raw_df
            st.session_state.categories = categories
            st.session_state.excel_data = buf
        except Exception as e:
            _prog2.empty()
            st.error(f"❌ 報表處理失敗: {e}")
            st.stop()
            
        _prog2.empty()
        st.success("🎉 報表產生成功！數據彙總已完成。")

    # 8. 顯示轉換成果區塊 (若已處理完成，則保持顯示)
    if st.session_state.processed:
        summary_df = st.session_state.summary_df
        raw_df = st.session_state.raw_df
        categories = st.session_state.categories
        excel_data = st.session_state.excel_data
        
        # 數據統計分析面板 (Dashboard)
        st.markdown('<div class="custom-card" style="border-top: 4px solid var(--accent-primary);">', unsafe_allow_html=True)
        st.markdown("### 📈 異常庫存數據分析摘要 (Analysis Dashboard)")
        
        # 計算看板指標
        total_overstock = summary_df[summary_df["Buyer"] == "總計"]["總計"].values[0]
        total_buyers = len(summary_df[summary_df["Buyer"] != "總計"]["Buyer"].dropna().unique())
        total_brands = len(summary_df[summary_df["Buyer"] != "總計"]["Brand"].dropna().unique())
        total_skus = raw_df["SKU"].dropna().nunique()
        
        # 取得各分類統計值
        cat_counts = {}
        for cat in categories:
            cat_counts[cat] = summary_df[summary_df["Buyer"] == "總計"][cat].values[0]
            
        r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
        with r1_col1:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #EF4444;">
                <div class="metric-title">總 Overstock 筆數</div>
                <div class="metric-value" style="color: var(--metric-red);">{int(total_overstock):,}</div>
                <div class="metric-desc">所有庫存異常之總計筆數</div>
            </div>
            """, unsafe_allow_html=True)
        with r1_col2:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #3B82F6;">
                <div class="metric-title">影響採購人員數 (Buyers)</div>
                <div class="metric-value">{total_buyers:,}</div>
                <div class="metric-desc">位採購人員名下有異常品項</div>
            </div>
            """, unsafe_allow_html=True)
        with r1_col3:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #8B5CF6;">
                <div class="metric-title">影響品牌總數 (Brands)</div>
                <div class="metric-value">{total_brands:,}</div>
                <div class="metric-desc">個品牌包含異常庫存商品</div>
            </div>
            """, unsafe_allow_html=True)
        with r1_col4:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #10B981;">
                <div class="metric-title">影響SKU總數</div>
                <div class="metric-value">{total_skus:,}</div>
                <div class="metric-desc">個包含異常庫存商品(SKUs)</div>
            </div>
            """, unsafe_allow_html=True)
            
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        
        transfer_col = _resolve_column_name("After new Safety Qty / ND can transfer out of QTY", set(raw_df.columns))
        transfer_qty_total = 0
        if transfer_col:
            transfer_qty_total = int(pd.to_numeric(raw_df[transfer_col], errors="coerce").fillna(0).sum())

        r2_col1, r2_col2, r2_col3 = st.columns(3)
        with r2_col1:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #F59E0B;">
                <div class="metric-title">Overstock (P90D 0 Sales)</div>
                <div class="metric-value">{int(cat_counts.get(categories[0], 0)):,}</div>
                <div class="metric-desc">「過去 90 天內零銷售」之異常筆數</div>
            </div>
            """, unsafe_allow_html=True)
        with r2_col2:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #EC4899;">
                <div class="metric-title">Overstock (>2*Monthly Sold Qty & >2pcs)</div>
                <div class="metric-value">{int(cat_counts.get(categories[1], 0)):,}</div>
                <div class="metric-desc">「大於月均銷售量兩倍且庫存大於 2 件」之異常筆數</div>
            </div>
            """, unsafe_allow_html=True)
        with r2_col3:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #06B6D4;">
                <div class="metric-title">New Safety Qty / ND 調整後的預計總數</div>
                <div class="metric-value">{transfer_qty_total:,}</div>
                <div class="metric-desc">After new Safety Qty / ND can transfer out of TOTAL QTY</div>
            </div>
            """, unsafe_allow_html=True)
            
        st.markdown('</div>', unsafe_allow_html=True)

        # 詳細資料索引標籤
        tab1, tab2 = st.tabs(["📋 統計彙總表 (Shop Stock Status)", "📄 原始明細資料預覽 (RAW)"])
        
        with tab1:
            st.markdown("#### 採購人員與品牌之庫存異常交叉統計表")
            st.markdown("本表即為 Excel 導出檔中「Shop Stock Status」工作表之即時數據：")
            st.dataframe(summary_df, use_container_width=True, hide_index=True)
            
        with tab2:
            st.markdown("#### RAW 原始資料庫存明細 (前 50 筆)")
            st.markdown("此表為過濾與重整後，完整導出之明細資料數據預覽：")
            st.dataframe(raw_df.head(50), use_container_width=True)
            st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

        # 導出與下載區塊
        st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)
        st.markdown('<div class="custom-card" style="background-color: var(--bg-download); border: 1px solid var(--border-download); text-align: center; padding: 2rem 1.5rem;">', unsafe_allow_html=True)
        st.markdown("<h3 style='color: var(--accent-primary); margin-top: 0; margin-bottom: 8px;'>🎉 報表處理完畢，已就緒！</h3>", unsafe_allow_html=True)
        st.markdown("<p style='color: var(--text-download-desc); font-size: 0.95rem; margin-bottom: 20px;'>本工具已為您自動設定好專業的 Excel 表格樣式、欄寬自動調整與自定義篩選器。</p>", unsafe_allow_html=True)
        
        hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
        download_name = f"I14 by SKUxShop_To Buyer_{hk_time}.xlsx"
        
        st.download_button(
            label="📥 下載完整分析報表 (.xlsx)",
            data=excel_data,
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)

# 9. 底部頁尾資訊
st.divider()
st.markdown(
    '<div style="text-align: center; color: var(--footer-text); font-size: 0.8rem; padding: 10px 0;">'
    'I14 報表轉換與分析系統 &nbsp;&nbsp;•&nbsp;&nbsp; Streamlit + Zeabur &nbsp;&nbsp;•&nbsp;&nbsp; '
    f'v{__version__} &nbsp;&nbsp;•&nbsp;&nbsp; Developed by Ricky Yue'
    '</div>',
    unsafe_allow_html=True,
)
