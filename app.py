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

# 2. 注入自定義進階 CSS 樣式以實現暗色專業配色 (參考 kilo.ai 風格)
st.markdown("""
<style>
/* 載入 Google 字型 */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Noto+Sans+TC:wght@300;400;500;700&display=swap');

/* 全域字型與背景微調 (暗色主題) */
html, body, [class*="css"] {
    font-family: 'Inter', 'Noto Sans TC', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* 主容器間距調整 */
.main .block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1200px;
}

/* 專業漸層 Header Banner (暗色) */
.header-container {
    background: linear-gradient(135deg, #1e3a8a 0%, #0f172a 100%);
    padding: 2.2rem;
    border-radius: 14px;
    color: #FFFFFF;
    margin-bottom: 2rem;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
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
    color: #93C5FD;
    margin-top: 8px;
    margin-bottom: 0;
    font-weight: 300;
}

/* 自定義精美卡片樣式 (暗色) */
.custom-card {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2), 0 2px 4px -1px rgba(0, 0, 0, 0.15);
    color: #f1f5f9;
}

/* 數據指標卡片 (Metric Cards) 暗色高對比 */
.metric-card {
    background-color: #1e293b;
    border: 1px solid #334155;
    border-radius: 10px;
    padding: 1.2rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.2);
    border-left: 5px solid #6366f1;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
    color: #f1f5f9;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 16px -3px rgba(0, 0, 0, 0.3);
}
.metric-title {
    font-size: 0.82rem;
    color: #94a3b8;
    font-weight: 500;
    margin-bottom: 4px;
}
.metric-value {
    font-size: 1.6rem;
    color: #f1f5f9;
    font-weight: 700;
    line-height: 1.2;
}
.metric-desc {
    font-size: 0.72rem;
    color: #64748b;
    margin-top: 4px;
}

/* Streamlit Tabs 頁籤樣式優化 (暗色) */
.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
    border-bottom: 1px solid #334155;
}
.stTabs [data-baseweb="tab"] {
    height: 44px;
    background-color: transparent;
    border-radius: 8px 8px 0px 0px;
    padding: 10px 16px;
    color: #94a3b8;
    font-weight: 500;
    transition: all 0.2s;
}
.stTabs [data-baseweb="tab"]:hover {
    color: #a5b4fc;
    background-color: #1e293b;
}
.stTabs [aria-selected="true"] {
    color: #6366f1 !important;
    border-bottom: 2px solid #6366f1 !important;
    font-weight: 700 !important;
}

/* 針對 Streamlit 內建按鈕的微調 (暗色) */
button[kind="primary"] {
    background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%) !important;
    border: none !important;
    color: white !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    padding: 0.55rem 1.4rem !important;
    transition: all 0.2s ease !important;
}
button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(99, 102, 241, 0.4) !important;
}

/* 檔案上傳框外觀優化 (暗色) */
[data-testid="stFileUploader"] {
    border: 2px dashed #475569;
    border-radius: 12px;
    padding: 1.2rem;
    background-color: #1e293b;
    color: #e2e8f0;
}
</style>
""", unsafe_allow_html=True)

# 3. 側邊欄設計 (Sidebar)
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 15px 0 5px 0;">
        <h2 style="color: #6366f1; margin-bottom: 2px; font-weight: 700; font-size: 1.6rem;">I14 Analyzer</h2>
        <p style="color: #94a3b8; font-size: 0.85rem; margin-top: 0;">自動化報表與數據轉換引擎</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
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
    - ✅ **Buyer** <span style="color:#94a3b8;font-size:0.8em">(採購人員)</span>
    - ✅ **Brand** <span style="color:#94a3b8;font-size:0.8em">(品牌名稱)</span>
    - ✅ **SKU** <span style="color:#94a3b8;font-size:0.8em">(品號編碼)</span>
    - ✅ **Shop Stock Status** <span style="color:#94a3b8;font-size:0.8em">(庫存狀態明細)</span>
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
        '<div style="color: #64748b; font-size: 0.75rem; text-align: center; line-height: 1.5;">'
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
    if "df_input" not in st.session_state:
        _prog = st.progress(0, text="讀取檔案中... 0% | 初始化中...")
        _t0 = time.time()
        try:
            _raw = uploaded_file.read()
            _wb = load_workbook(BytesIO(_raw), data_only=True)
            _ws = _wb.active
            _n = _ws.max_row
            if _n <= 1:
                _prog.empty()
                st.error("❌ 上傳的 Excel 檔案沒有資料列（僅含標題或完全空白），請確認檔案內容後再試。")
                st.stop()
            else:
                _hdr = list(next(_ws.iter_rows(min_row=1, max_row=1, values_only=True)))
                _rows = []
                for _i, _r in enumerate(_ws.iter_rows(min_row=2, max_row=_n, values_only=True), start=1):
                    _rows.append(list(_r))
                    if _i == _n - 1 or (_i % 500 == 0 and _i != _n - 1):
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
        <div style="background-color: #450a0a; border: 1px solid #b91c1c; padding: 15px; border-radius: 8px; margin-bottom: 1.5rem;">
            <h4 style="color: #f87171; margin-top: 0; margin-bottom: 8px;">❌ 來源檔案欄位檢驗未通過</h4>
            <p style="color: #fecaca; margin-bottom: 0; font-size: 0.9rem;">
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
            <div class="metric-value" style="color: #4ade80;">驗證通過</div>
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
        st.markdown('<div class="custom-card" style="border-top: 4px solid #6366f1;">', unsafe_allow_html=True)
        st.markdown("### 📈 異常庫存數據分析摘要 (Analysis Dashboard)")
        
        # 計算看板指標
        _total_row = summary_df[summary_df["Buyer"] == "總計"]
        if _total_row.empty:
            st.error("❌ 摘要資料中未找到總計列，無法顯示分析看板。")
            st.stop()
        total_overstock = int(_total_row["總計"].values[0])
        total_buyers = len(summary_df[summary_df["Buyer"] != "總計"]["Buyer"].dropna().unique())
        total_brands = len(summary_df[summary_df["Buyer"] != "總計"]["Brand"].dropna().unique())
        total_skus = raw_df["SKU"].dropna().nunique()

        # 取得各分類統計值
        cat_counts = {}
        for cat in categories:
            cat_counts[cat] = int(_total_row[cat].values[0])
            
        r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
        with r1_col1:
            st.markdown(f"""
            <div class="metric-card" style="border-left-color: #EF4444;">
                <div class="metric-title">總 Overstock 筆數</div>
                <div class="metric-value" style="color: #f87171;">{int(total_overstock):,}</div>
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
        
        _transfer_col_name = "After new Safety Qty / ND can transfer out of QTY"
        transfer_col_resolved = _resolve_column_name(_transfer_col_name, set(summary_df.columns))
        transfer_qty_total = int(_total_row[_transfer_col_name].values[0]) if transfer_col_resolved is not None and _transfer_col_name in _total_row.columns else int(pd.to_numeric(raw_df[_transfer_col_name], errors="coerce").fillna(0).sum()) if _transfer_col_name in raw_df.columns else 0

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
        st.markdown('<div class="custom-card" style="background-color: #1e293b; border: 1px solid #6366f1; text-align: center; padding: 2rem 1.5rem;">', unsafe_allow_html=True)
        st.markdown("<h3 style='color: #6366f1; margin-top: 0; margin-bottom: 8px;'>🎉 報表處理完畢，已就緒！</h3>", unsafe_allow_html=True)
        st.markdown("<p style='color: #94a3b8; font-size: 0.95rem; margin-bottom: 20px;'>本工具已為您自動設定好專業的 Excel 表格樣式、欄寬自動調整與自定義篩選器。</p>", unsafe_allow_html=True)
        
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
    '<div style="text-align: center; color: #64748b; font-size: 0.8rem; padding: 10px 0;">'
    'I14 報表轉換與分析系統 &nbsp;&nbsp;•&nbsp;&nbsp; Streamlit + Zeabur &nbsp;&nbsp;•&nbsp;&nbsp; '
    f'v{__version__} &nbsp;&nbsp;•&nbsp;&nbsp; Developed by Ricky Yue'
    '</div>',
    unsafe_allow_html=True,
)
