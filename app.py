import streamlit as st
import pandas as pd
import calendar
from datetime import datetime
from zoneinfo import ZoneInfo
from transformer import transform, generate_excel, validate_input, _resolve_column_name, transform_oos, generate_excel_oos
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
    background: linear-gradient(135deg, #0f2027 0%, #0a0a0f 100%);
    padding: 2.2rem;
    border-radius: 14px;
    color: #FFFFFF;
    margin-bottom: 2rem;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.5);
    border: 1px solid #00d4ff33;
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
    color: #00d4ff;
    margin-top: 8px;
    margin-bottom: 0;
    font-weight: 300;
}

/* 自定義精美卡片樣式 (暗色) */
.custom-card {
    background-color: #12121a;
    border: 1px solid #1e1e2e;
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.3), 0 2px 4px -1px rgba(0, 0, 0, 0.2);
    color: #ffffff;
}

/* 數據指標卡片 (Metric Cards) 暗色高對比 */
.metric-card {
    background-color: #12121a;
    border: 1px solid #1e1e2e;
    border-radius: 10px;
    padding: 1.2rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.3);
    border-left: 5px solid #00d4ff;
    transition: transform 0.2s ease, box-shadow 0.2s ease;
    color: #ffffff;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 16px -3px rgba(0, 0, 0, 0.4);
    border-color: #00d4ff;
}
.metric-title {
    font-size: 0.82rem;
    color: #b0b0c0;
    font-weight: 500;
    margin-bottom: 4px;
}
.metric-value {
    font-size: 1.6rem;
    color: #00d4ff;
    font-weight: 700;
    line-height: 1.2;
}
.metric-desc {
    font-size: 0.72rem;
    color: #9999aa;
    margin-top: 4px;
}

/* Streamlit Tabs 頁籤樣式優化 (暗色) */
.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
    border-bottom: 1px solid #1e1e2e;
}
.stTabs [data-baseweb="tab"] {
    height: 44px;
    background-color: transparent;
    border-radius: 8px 8px 0px 0px;
    padding: 10px 16px;
    color: #888899;
    font-weight: 500;
    transition: all 0.2s;
}
.stTabs [data-baseweb="tab"]:hover {
    color: #00d4ff;
    background-color: #12121a;
}
.stTabs [aria-selected="true"] {
    color: #00d4ff !important;
    border-bottom: 2px solid #00d4ff !important;
    font-weight: 700 !important;
}

/* 針對 Streamlit 內建按鈕的微調 (暗色) */
button[kind="primary"] {
    background: linear-gradient(135deg, #00d4ff 0%, #0099cc 100%) !important;
    border: none !important;
    color: #0a0a0f !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    padding: 0.55rem 1.4rem !important;
    transition: all 0.2s ease !important;
}
button[kind="primary"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 12px rgba(0, 212, 255, 0.4) !important;
}

/* 檔案上傳框外觀優化 (暗色) */
[data-testid="stFileUploader"] {
    border: 2px dashed #00d4ff44;
    border-radius: 12px;
    padding: 1.2rem;
    background-color: #12121a;
    color: #ffffff;
}
</style>
""", unsafe_allow_html=True)

# 3. 側邊欄設計 (Sidebar)
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 15px 0 5px 0;">
        <h2 style="color: #00d4ff; margin-bottom: 2px; font-weight: 700; font-size: 1.6rem;">I14 Analyzer</h2>
        <p style="color: #888899; font-size: 0.85rem; margin-top: 0;">自動化報表與數據轉換引擎</p>
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
    - ✅ **Buyer** <span style="color:#00d4ff;font-size:0.8em">(採購人員)</span>
    - ✅ **Brand** <span style="color:#00d4ff;font-size:0.8em">(品牌名稱)</span>
    - ✅ **SKU** <span style="color:#00d4ff;font-size:0.8em">(品號編碼)</span>
    - ✅ **Shop Stock Status** <span style="color:#00d4ff;font-size:0.8em">(庫存狀態明細)</span>
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
        '<div style="color: #444455; font-size: 0.75rem; text-align: center; line-height: 1.5;">'
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

# 5. 模式選擇器
mode_options = ["Overstock Review", "OOS Review", "Combined Review"]
current_mode = st.radio("選擇分析模式", mode_options, horizontal=True, index=0)

# 6. 初始化 Session State (確保下載與重整時狀態持久化)
if "processed" not in st.session_state:
    st.session_state.processed = False
    st.session_state.summary_df = None
    st.session_state.raw_df = None
    st.session_state.categories = None
    st.session_state.excel_data = None
    st.session_state.last_uploaded_key = None
    st.session_state.warnings = []

if "oos_processed" not in st.session_state:
    st.session_state.oos_processed = False
    st.session_state.oos_summary_df = None
    st.session_state.oos_raw_df = None
    st.session_state.oos_categories = None
    st.session_state.oos_excel_data = None
    st.session_state.oos_warnings = []

# 7. 模式切換狀態清理
prev_mode = st.session_state.get("mode")
if prev_mode is not None and prev_mode != current_mode:
    st.session_state.processed = False
    st.session_state.summary_df = None
    st.session_state.raw_df = None
    st.session_state.categories = None
    st.session_state.excel_data = None
    st.session_state.warnings = []
    st.session_state.oos_processed = False
    st.session_state.oos_summary_df = None
    st.session_state.oos_raw_df = None
    st.session_state.oos_categories = None
    st.session_state.oos_excel_data = None
    st.session_state.oos_warnings = []
    st.session_state.pop("last_uploaded_key", None)
st.session_state.mode = current_mode
mode = current_mode

# 8. 檔案上傳區域
uploaded_file = st.file_uploader(
    "📂 上傳 By SKUxShop 原始 Excel 檔案 (.xlsx)",
    type=["xlsx"],
    help="支援標準 Excel (.xlsx) 格式，檔案上限 200MB，最佳化支援 5 萬行以上大數據處理"
)

# 9. 處理上傳的檔案
if uploaded_file is not None:
    # 當上傳了新檔案，自動重設先前的處理狀態
    _file_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("last_uploaded_key") != _file_key:
        st.session_state.last_uploaded_key = _file_key
        st.session_state.processed = False
        st.session_state.summary_df = None
        st.session_state.raw_df = None
        st.session_state.categories = None
        st.session_state.excel_data = None
        st.session_state.warnings = []
        st.session_state.pop("df_input", None)

    # 若尚未載入資料，則執行讀取
    if "df_input" not in st.session_state:
        _prog = st.progress(0, text="讀取檔案中... 0% | 初始化中...")
        _t0 = time.time()
        try:
            _raw = uploaded_file.read()
            _wb = load_workbook(BytesIO(_raw), data_only=True, read_only=True)
            _ws = _wb.active
            _hdr = list(next(_ws.iter_rows(min_row=1, max_row=1, values_only=True)))
            _rows = []
            for _i, _r in enumerate(_ws.iter_rows(min_row=2, values_only=True), start=1):
                _rows.append(list(_r))
                if _i % 500 == 0:
                    _elapsed = time.time() - _t0
                    _prog.progress(min(_i / (_i + 500), 0.99), text=f"讀取檔案中... {_i} 行 | 已用: {_elapsed:.0f}s")
            if _i == 0:
                _prog.empty()
                _wb.close()
                del _raw, _wb, _rows, _hdr
                st.error("❌ 上傳的 Excel 檔案沒有資料列（僅含標題或完全空白），請確認檔案內容後再試。")
                st.stop()
            df_input = pd.DataFrame(_rows, columns=_hdr)
            _wb.close()
            del _raw, _wb, _rows, _hdr
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
        <div style="background-color: #1a0a0a; border: 1px solid #ff4444; padding: 15px; border-radius: 8px; margin-bottom: 1.5rem;">
            <h4 style="color: #ff6666; margin-top: 0; margin-bottom: 8px;">❌ 來源檔案欄位檢驗未通過</h4>
            <p style="color: #ffaaaa; margin-bottom: 0; font-size: 0.9rem;">
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
        <div class="metric-card" style="border-left-color: #00ff88;">
            <div class="metric-title">檔案解析狀態</div>
            <div class="metric-value" style="color: #00ff88;">驗證通過</div>
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

    if mode == "Combined Review":
        st.info("📢 Combined Review 模式即將推出，敬請期待！")
    else:
        if mode == "OOS Review":
            mtd_date = st.date_input(
                "📅 MTD 日期 (Sold Qty CM-MTD 的截止日期)",
                value=datetime.now().date(),
                help="選擇 MTD (Month-To-Date) 截止日期，系統將據此推算月度預估銷售量。"
            )
            mtd_days = mtd_date.day
            month_total_days = calendar.monthrange(mtd_date.year, mtd_date.month)[1]
            st.caption(f"MTD: {mtd_days}日，當月共 {month_total_days} 日")

        st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

        SHOP_NATURE_OPTIONS = ["Tourist", "Mix", "Non Tourist"]
        selected_shop_natures = st.multiselect(
            "🏷️ 選擇要處理的 Shop Nature 類型",
            options=SHOP_NATURE_OPTIONS,
            default=SHOP_NATURE_OPTIONS,
            help="選擇要處理的店鋪性質。預設全選所有類型。"
        )

        with st.expander("⚙️ 自定義 Shop Nature Rate (進階設定)", expanded=False):
            st.markdown("調整各 Shop Nature 類型的安全庫存建議倍率（% of Monthly Sold Qty）。")
            _rate_cols = st.columns(3)
            with _rate_cols[0]:
                rate_tourist = st.slider("Tourist Rate (%)", 0, 100, 30, 1,
                    help="Tourist 類型：安全庫存建議 = floor(Monthly Sold × Rate)，最低 3")
            with _rate_cols[1]:
                rate_mix = st.slider("Mix Rate (%)", 0, 100, 40, 1,
                    help="Mix 類型：安全庫存建議 = floor(Monthly Sold × Rate)，最低 3")
            with _rate_cols[2]:
                rate_non_tourist = st.slider("Non Tourist Rate (%)", 0, 100, 50, 1,
                    help="Non Tourist 及其他類型：安全庫存建議 = floor(Monthly Sold × Rate)，最低 3")
            st.caption(f"目前設定 → Tourist: **{rate_tourist}%** | Mix: **{rate_mix}%** | Non Tourist: **{rate_non_tourist}%**")

        rates = {
            "tourist": rate_tourist / 100.0,
            "mix": rate_mix / 100.0,
            "default": rate_non_tourist / 100.0,
        }

        if st.button("⚙️ 執行報表轉換與分析", type="primary", use_container_width=True):
            _prog2 = st.progress(0, text="處理中... 0% | 開始處理...")
            _t0_2 = time.time()

            def _cb(pct, msg):
                _e = time.time() - _t0_2
                _eta2 = (_e / pct) - _e if pct > 0 else 0
                _prog2.progress(min(pct, 1.0), text=f"處理中... {pct*100:.0f}% | {msg} | 已用: {_e:.0f}s | 預計剩餘: {_eta2:.0f}s")

            try:
                def _cb_transform(pct, msg):
                    _cb(0.6 * pct, msg)

                def _cb_excel(pct, msg):
                    _cb(0.6 + 0.4 * pct, msg)

                if mode == "Overstock Review":
                    summary_df, raw_df, categories, _warnings = transform(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                        rates=rates,
                    )
                    buf = generate_excel(summary_df, raw_df, categories, progress_callback=_cb_excel)
                    st.session_state.processed = True
                    st.session_state.summary_df = summary_df
                    st.session_state.raw_df = raw_df
                    st.session_state.categories = categories
                    st.session_state.excel_data = buf
                    st.session_state.warnings = _warnings
                else:  # OOS Review
                    summary_df, raw_df, categories, _warnings = transform_oos(
                        df_input,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                        rates=rates,
                        mtd_days=mtd_days,
                        month_total_days=month_total_days,
                    )
                    buf = generate_excel_oos(summary_df, raw_df, categories, progress_callback=_cb_excel)
                    st.session_state.oos_processed = True
                    st.session_state.oos_summary_df = summary_df
                    st.session_state.oos_raw_df = raw_df
                    st.session_state.oos_categories = categories
                    st.session_state.oos_excel_data = buf
                    st.session_state.oos_warnings = _warnings
            except Exception as e:
                _prog2.empty()
                st.error(f"❌ 報表處理失敗: {e}")
                st.stop()

            _prog2.empty()
            st.success("🎉 報表產生成功！數據彙總已完成。")
            if _warnings:
                for _w in _warnings:
                    st.warning(f"⚠️ {_w}")

        # 顯示轉換成果區塊
        if mode == "Overstock Review" and st.session_state.processed:
            summary_df = st.session_state.summary_df
            raw_df = st.session_state.raw_df
            categories = st.session_state.categories
            excel_data = st.session_state.excel_data

            st.markdown('<div class="custom-card" style="border-top: 4px solid #6366f1;">', unsafe_allow_html=True)
            st.markdown("### 📈 異常庫存數據分析摘要 (Analysis Dashboard)")

            _total_row = summary_df[summary_df["Buyer"] == "總計"]
            if _total_row.empty:
                st.error("❌ 摘要資料中未找到總計列，無法顯示分析看板。")
                st.stop()
            total_overstock = int(_total_row["總計"].values[0])
            total_buyers = len(summary_df[summary_df["Buyer"] != "總計"]["Buyer"].dropna().unique())
            total_brands = len(summary_df[summary_df["Buyer"] != "總計"]["Brand"].dropna().unique())
            total_skus = raw_df["SKU"].dropna().nunique()

            cat_counts = {}
            for cat in categories:
                cat_counts[cat] = int(_total_row[cat].values[0])

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff5555;">
                    <div class="metric-title">總 Overstock 筆數</div>
                    <div class="metric-value" style="color: #ff5555;">{int(total_overstock):,}</div>
                    <div class="metric-desc">所有庫存異常之總計筆數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #5599ff;">
                    <div class="metric-title">影響採購人員數 (Buyers)</div>
                    <div class="metric-value" style="color: #5599ff;">{total_buyers:,}</div>
                    <div class="metric-desc">位採購人員名下有異常品項</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #9944ff;">
                    <div class="metric-title">影響品牌總數 (Brands)</div>
                    <div class="metric-value" style="color: #9944ff;">{total_brands:,}</div>
                    <div class="metric-desc">個品牌包含異常庫存商品</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #00ff88;">
                    <div class="metric-title">影響SKU總數</div>
                    <div class="metric-value" style="color: #00ff88;">{total_skus:,}</div>
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
                <div class="metric-card" style="border-left-color: #ff9944;">
                    <div class="metric-title">Overstock (P90D 0 Sales)</div>
                    <div class="metric-value" style="color: #ff9944;">{int(cat_counts.get(categories[0], 0)):,}</div>
                    <div class="metric-desc">「過去 90 天內零銷售」之異常筆數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff44ff;">
                    <div class="metric-title">Overstock (>2*Monthly Sold Qty & >2pcs)</div>
                    <div class="metric-value" style="color: #ff44ff;">{int(cat_counts.get(categories[1], 0)):,}</div>
                    <div class="metric-desc">「大於月均銷售量兩倍且庫存大於 2 件」之異常筆數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #00d4ff;">
                    <div class="metric-title">New Safety Qty / ND 調整後的預計總數</div>
                    <div class="metric-value" style="color: #00d4ff;">{transfer_qty_total:,}</div>
                    <div class="metric-desc">After new Safety Qty / ND can transfer out of TOTAL QTY</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

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

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)
            st.markdown('<div class="custom-card" style="background-color: #12121a; border: 1px solid #00d4ff; text-align: center; padding: 2rem 1.5rem;">', unsafe_allow_html=True)
            st.markdown("<h3 style='color: #00d4ff; margin-top: 0; margin-bottom: 8px;'>🎉 報表處理完畢，已就緒！</h3>", unsafe_allow_html=True)
            st.markdown("<p style='color: #888899; font-size: 0.95rem; margin-bottom: 20px;'>本工具已為您自動設定好專業的 Excel 表格樣式、欄寬自動調整與自定義篩選器。</p>", unsafe_allow_html=True)

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

        if mode == "OOS Review" and st.session_state.oos_processed:
            summary_df = st.session_state.oos_summary_df
            raw_df = st.session_state.oos_raw_df
            categories = st.session_state.oos_categories
            excel_data = st.session_state.oos_excel_data

            st.markdown('<div class="custom-card" style="border-top: 4px solid #ef4444;">', unsafe_allow_html=True)
            st.markdown("### 📈 OOS 庫存數據分析摘要 (OOS Analysis Dashboard)")

            _total_row = summary_df[summary_df["Buyer"] == "總計"]
            if _total_row.empty:
                st.error("❌ 摘要資料中未找到總計列，無法顯示分析看板。")
                st.stop()
            total_oos = int(_total_row["總計"].values[0])
            total_buyers = len(summary_df[summary_df["Buyer"] != "總計"]["Buyer"].dropna().unique())
            total_brands = len(summary_df[summary_df["Buyer"] != "總計"]["Brand"].dropna().unique())
            total_skus = raw_df["SKU"].dropna().nunique()

            cat_counts = {}
            for cat in categories:
                cat_counts[cat] = int(_total_row[cat].values[0])

            safety_increase_total = int(pd.to_numeric(raw_df.get("Safety Qty Increase", pd.Series(0, index=raw_df.index)), errors="coerce").fillna(0).sum())

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff5555;">
                    <div class="metric-title">總 OOS 筆數</div>
                    <div class="metric-value" style="color: #ff5555;">{total_oos:,}</div>
                    <div class="metric-desc">所有 OOS 異常之總計筆數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #5599ff;">
                    <div class="metric-title">影響採購人員數 (Buyers)</div>
                    <div class="metric-value" style="color: #5599ff;">{total_buyers:,}</div>
                    <div class="metric-desc">位採購人員名下有 OOS 品項</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #9944ff;">
                    <div class="metric-title">影響品牌總數 (Brands)</div>
                    <div class="metric-value" style="color: #9944ff;">{total_brands:,}</div>
                    <div class="metric-desc">個品牌包含 OOS 商品</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #00ff88;">
                    <div class="metric-title">影響SKU總數</div>
                    <div class="metric-value" style="color: #00ff88;">{total_skus:,}</div>
                    <div class="metric-desc">個包含 OOS 商品(SKUs)</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            r2_col1, r2_col2, r2_col3, r2_col4 = st.columns(4)
            with r2_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff9944;">
                    <div class="metric-title">OOS (0 Stock)</div>
                    <div class="metric-value" style="color: #ff9944;">{int(cat_counts.get(categories[0], 0)):,}</div>
                    <div class="metric-desc">現有庫存為 0 之 OOS 筆數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff44ff;">
                    <div class="metric-title">Potential OOS (<7 Turnday)</div>
                    <div class="metric-value" style="color: #ff44ff;">{int(cat_counts.get(categories[1], 0)):,}</div>
                    <div class="metric-desc">Turnday 不足 7 之潛在 OOS 筆數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #06b6d4;">
                    <div class="metric-title">Normal (Safety Qty Insufficient)</div>
                    <div class="metric-value" style="color: #06b6d4;">{int(cat_counts.get(categories[2], 0)):,}</div>
                    <div class="metric-desc">正常庫存但安全庫存不足筆數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #8b5cf6;">
                    <div class="metric-title">建議 Safety Qty 增加總量</div>
                    <div class="metric-value" style="color: #8b5cf6;">{safety_increase_total:,}</div>
                    <div class="metric-desc">建議新增 Safety Qty 之總量</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            tab1, tab2 = st.tabs(["📋 統計彙總表 (Shop Stock Status)", "📄 原始明細資料預覽 (RAW)"])

            with tab1:
                st.markdown("#### 採購人員與品牌之 OOS 交叉統計表")
                st.markdown("本表即為 Excel 導出檔中「Shop Stock Status」工作表之即時數據：")
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

            with tab2:
                st.markdown("#### RAW 原始資料庫存明細 (前 50 筆)")
                st.markdown("此表為過濾與重整後，完整導出之明細資料數據預覽：")
                st.dataframe(raw_df.head(50), use_container_width=True)
                st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)
            st.markdown('<div class="custom-card" style="background-color: #12121a; border: 1px solid #00d4ff; text-align: center; padding: 2rem 1.5rem;">', unsafe_allow_html=True)
            st.markdown("<h3 style='color: #00d4ff; margin-top: 0; margin-bottom: 8px;'>🎉 OOS 報表處理完畢，已就緒！</h3>", unsafe_allow_html=True)
            st.markdown("<p style='color: #888899; font-size: 0.95rem; margin-bottom: 20px;'>本工具已為您自動設定好專業的 Excel 表格樣式、欄寬自動調整與自定義篩選器。</p>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 OOS Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 OOS 分析報表 (.xlsx)",
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
    '<div style="text-align: center; color: #444455; font-size: 0.8rem; padding: 10px 0;">'
    'I14 報表轉換與分析系統 &nbsp;&nbsp;•&nbsp;&nbsp; Streamlit + Zeabur &nbsp;&nbsp;•&nbsp;&nbsp; '
    f'v{__version__} &nbsp;&nbsp;•&nbsp;&nbsp; Developed by Ricky Yue'
    '</div>',
    unsafe_allow_html=True,
)
