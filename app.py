import streamlit as st
import pandas as pd
import calendar
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from transformer import (
    transform, generate_excel, validate_input, _resolve_column_name,
    transform_oos, generate_excel_oos,
    transform_carry, generate_excel_carry,
    transform_combined, generate_excel_combined,
    transform_nd, generate_excel_nd,
    transform_stock_impact, generate_excel_stock_impact,
)
import time
import traceback
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

# 2. 注入自定義進階 CSS 樣式以實現淺色專業配色
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
</style>
""", unsafe_allow_html=True)

# 3. 側邊欄設計 (Sidebar)
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 15px 0 5px 0;">
        <h2 style="color: #0077a3; margin-bottom: 2px; font-weight: 700; font-size: 1.6rem;">I14 Analyzer</h2>
        <p style="color: #5a5a6e; font-size: 0.85rem; margin-top: 0;">自動化報表與數據轉換引擎</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    st.markdown("### 🎯 分析模式說明")
    st.markdown("""
    於主頁面頂端選擇分析模式後，系統將自動切換對應的計算流程與輸出：

    **📦 Overstock Review**
    - 篩選庫存過賸品項（P90D 零銷售 / 超過月均兩倍）
    - 計算新 Safety Qty 與可轉出貨量
    - 輸出：RAW、Shop Stock Status、By Shop、分析摘要

    **🚨 OOS Review**
    - 篩選缺貨（0 Stock）及潛在缺貨（Turnday < 7）
    - 偵測 Safety Qty 不足之正常庫存品項
    - 支援 MTD 日期選擇與月度預估銷售計算
    - 輸出：RAW、Shop Stock Status、By Shop、分析摘要

    **📋 Carry SKU Review**
    - 檢視各店舖的 Carry SKU 數量（扣除 ND 且庫存為 0 之誤差列）
    - 按 RP Type（ND / RF）與 P90D 0 Sales 交叉統計
    - 輸出：RAW、Summary（雙區塊樞紐：Adj. MCH1 × Adj. SST）、Carry SKU 分析摘要

    **✂️ SKU 精簡 Review**
    - 針對 RP Type 為 ND 的 SKU 建議退倉，區分 Supply Source
    - Warehouse 貨：建議全數退倉；Vendor 貨：標記「另有安排」不可退倉
    - 依 Supply Source × Shop Stock Status 雙層動態分類
    - 輸出：RAW、Shop Stock Status、By Shop、ND SKU 數據分析摘要

    **🔒 Safety Stock Review**
    - 全面審查各 SKU×Shop 的安全庫存設定是否合理
    - 結合 Overstock Cap（庫存上限）與 OOS Floor（缺貨底線），雙向平衡調整 Safety Stock
    - 支援逐 Site 檢視、Overstock / OOS Rate 自訂、Safety 只增或加減模式

    **📦 Stock Impact Analysis（缺貨銷售影響分析）**
    - 以 Shop Stock Qty = 0（缺貨）為核心，分析存貨對 Sold Qty 的因果影響（無貨就無銷售）
    - 比對 M-2 / M-1 / CM 三期的 Shop Stock 與 Sold Qty 對應關係
    - 缺貨銷售損失雙基準並列：M-2 Sold 為需求基準 vs P90D 90 天月均為基準
    - Shop Nature 優先級：Tourist > Mix > Non Tourist
    - SKU 級跨 Nature 調貨建議：Tourist 缺貨但 Non-Tourist/Mix 仍有存貨 → 揪出存貨配置不當
    - 預設只分析 Warehouse SKU，提供切換可納入 Vendor
    """)
    
    st.divider()
    
    st.markdown("### 📋 系統操作流程")
    st.markdown("""
    1. 📥 **準備來源檔案**：確認您的 Excel 檔案（By SKUxShop）包含必要欄位。
    2. 📂 **上傳檔案**：將 Excel 拖曳至右側上傳區域。
    3. 🎯 **選擇模式**：於主頁面頂端選擇 Overstock Review 或 OOS Review。
    4. 🔧 **設定參數**：根據模式需要，調整 MTD 日期、Shop Nature Rate 等進階選項。
    5. ⚙️ **執行轉換**：點擊「執行報表轉換與分析」，系統將自動過濾、計算並產生彙總表。
    6. 💾 **下載成果**：於視覺化看板中即時查閱統計資訊，並一鍵導出美化後的 Excel 報表。
    """)
    
    st.divider()
    
    st.markdown("### 🔍 必要欄位檢核標準")
    st.markdown("""
    為確保計算正確，來源檔案必須包含以下欄位：
    - ✅ **Buyer** <span style="color:#0077a3;font-size:0.8em">(採購人員)</span>
    - ✅ **Brand** <span style="color:#0077a3;font-size:0.8em">(品牌名稱)</span>
    - ✅ **SKU** <span style="color:#0077a3;font-size:0.8em">(品號編碼)</span>
    - ✅ **Shop Stock Status** <span style="color:#0077a3;font-size:0.8em">(庫存狀態明細)</span>

    OOS Review 額外建議欄位（非必要，缺少時將顯示提示）：
    - 📌 **Sold Qty (CM - MTD)** <span style="color:#f59e0b;font-size:0.8em">(月度銷售)</span>
    - 📌 **Adj. Safety Stock** <span style="color:#f59e0b;font-size:0.8em">(現有安全庫存)</span>
    - 📌 **Shop Nature** <span style="color:#f59e0b;font-size:0.8em">(店鋪性質)</span>
    """, unsafe_allow_html=True)
    
    st.divider()
    
    st.markdown("### ⚙️ 核心計算規則")
    st.markdown("""
    **Overstock Review** — 針對 `Shop Stock Status` 過濾以下庫存過賸條件：
    1. **P90D 0 Sales** — 過去 90 天內零銷售
    2. **Overstock (>2*Monthly Qty & >2pcs)** — 庫存超過月均銷售兩倍且大於 2 件

    **OOS Review** — 針對 `Shop Stock Status` 過濾以下缺貨條件：
    1. **OOS (0 Stock)** — 現有庫存為 0
    2. **Potential OOS (<7 Turnday)** — Turnday 不足 7 天
    3. **Normal (Safety Qty Insufficient)** — 庫存正常但 Safety Qty 不足以覆蓋預估銷售

    **Carry SKU Review** — 檢視各店舖實際 Carry SKU 數量：
    1. 先清理無效列（刪除 RP Type 為 ND 且 Shop Stock Qty 為 0 之誤差列）
    2. 依 RP Type (ND / RF) 與 P90D 0 Sales 狀態做雙維度交叉統計
    3. 雙區塊樞紐：左 (Adj. MCH1) + 右 (Adj. SST)

    **SKU 精簡 Review** — 針對 ND SKU 建議退倉（區分 Supply Source）：
    1. 篩選 RP Type = ND 的所有 SKU（不分 Shop Stock Status）
    2. **Warehouse 貨**：建議全數退倉 `transfer = Shop Stock Qty (CM)`
    3. **Vendor 貨**：標記「另有安排」，不可退倉 `transfer = 0`
    4. ND SKU 依 Supply Source × Shop Stock Status 雙層動態分類（如 Warehouse - Overstock、vendor - Normal）

     **🔒 Safety Stock Review** — 全面審查 Safety Stock 設定，計算平衡建議值：
     1. **Overstock Cap (X)** = floor(月均銷售 × Overstock Rate)，上限不超過現有 Adj. Safety Stock
     2. **OOS Floor (Y)** = ceil(MTD 預估月銷售 × OOS Rate)，最低 3
     3. **建議 Safety Stock** = max(Y, Z if Z ≤ X else X)，其中 Z = 現有 Adj. Safety Stock
     4. **Safety Qty Change** = 建議 Safety Stock − Z（正數=增加、負數=減少、零=不變）
     5. 品項分為 Overstock / OOS / Balanced (Safety Adjusted) 三類，標示調整原因

     **📦 Stock Impact Analysis** — 缺貨銷售影響分析判別：
     1. **缺貨判定** = Shop Stock Qty = 0（嚴格標準，三期 M-2 / M-1 / CM 各自判定）
     2. **缺貨模式**：長期缺貨（≥2 期）/ 間歇缺貨（1 期）/ 恢復供貨（前期缺貨 → CM 有貨）/ 無缺貨
     3. **銷售損失（M-2 基準）** = max(0, Sold Qty M-2 − 缺貨當期 Sold)，以 M-2 銷售為需求基準
     4. **銷售損失（P90D 基準）** = max(0, Shop Monthly Sold Qty P90D − 缺貨當期 Sold)，以 90 天月均為基準
     5. **影響判定** = 任一期缺貨且有需求證據（M-2 Sold 或 P90D > 0），且只計算 RP Type = RF 的店舖
     6. **Shop Nature 優先級**：Tourist > Mix > Non Tourist（彙總按 Tourist 缺貨嚴重度排序）
     7. **SKU 調貨建議**：Tourist 缺貨但同 SKU 在 Non-Tourist/Mix 仍有存貨 → 揪出存貨配置不當
     8. 預設只分析 Warehouse SKU；支援多維度彙總：By Site / Brand / Buyer / Category (Adj. MCH1)

     所有模式均自動生成採購與品牌的交叉匯總樞紐分析表。
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
mode_options = ["Overstock Review", "OOS Review", "Carry SKU Review", "SKU 精簡 Review", "Safety Stock Review", "Stock Impact Analysis"]
current_mode = st.radio("選擇分析模式", mode_options, horizontal=True, index=0)

# 6. 初始化 Session State (確保下載與重整時狀態持久化)
MODE_KEYS = {
    "Overstock Review": {"prefix": "", "processed": False, "summary_df": None, "raw_df": None, "categories": None, "excel_data": None, "warnings": []},
    "OOS Review": {"prefix": "oos_", "processed": False, "summary_df": None, "raw_df": None, "categories": None, "excel_data": None, "warnings": []},
    "Carry SKU Review": {"prefix": "carry_", "processed": False, "summary_left": None, "summary_right": None, "dashboard_df": None, "raw_df": None, "excel_data": None, "warnings": []},
    "SKU 精簡 Review": {"prefix": "nd_", "processed": False, "summary_df": None, "raw_df": None, "categories": None, "excel_data": None, "warnings": []},
    "Safety Stock Review": {"prefix": "combined_", "processed": False, "summary_df": None, "raw_df": None, "categories": None, "excel_data": None, "warnings": []},
    "Stock Impact Analysis": {"prefix": "stock_impact_", "processed": False, "summary_site": None, "summary_brand": None, "summary_buyer": None, "summary_category": None, "sku_allocation": None, "sku_diagnosis": None, "raw_df": None, "excel_data": None, "warnings": []},
}

for cfg in MODE_KEYS.values():
    for attr, default in cfg.items():
        if attr == "prefix":
            continue
        key = f"{cfg['prefix']}{attr}"
        if key not in st.session_state:
            st.session_state[key] = [] if isinstance(default, list) else default

if "last_uploaded_key" not in st.session_state:
    st.session_state.last_uploaded_key = None

# 7. 模式切換狀態清理
prev_mode = st.session_state.get("mode")
if prev_mode is not None and prev_mode != current_mode:
    for cfg in MODE_KEYS.values():
        for attr, default in cfg.items():
            if attr == "prefix":
                continue
            key = f"{cfg['prefix']}{attr}"
            st.session_state[key] = [] if isinstance(default, list) else default
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
    # 當上傳了新檔案，自動重設所有先前的處理狀態
    _file_key = f"{uploaded_file.name}_{uploaded_file.size}"
    if st.session_state.get("last_uploaded_key") != _file_key:
        st.session_state.last_uploaded_key = _file_key
        for cfg in MODE_KEYS.values():
            for attr, default in cfg.items():
                if attr == "prefix":
                    continue
                key = f"{cfg['prefix']}{attr}"
                st.session_state[key] = [] if isinstance(default, list) else default
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
            _i = 0
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
        <div style="background-color: #fff5f5; border: 1px solid #ff4444; padding: 15px; border-radius: 8px; margin-bottom: 1.5rem;">
            <h4 style="color: #c0392b; margin-top: 0; margin-bottom: 8px;">❌ 來源檔案欄位檢驗未通過</h4>
            <p style="color: #922b21; margin-bottom: 0; font-size: 0.9rem;">
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
            <div class="metric-title">總資料行數</div>
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
        st.dataframe(df_input.head(20), width='stretch')
        st.caption(f"目前顯示第 1 至第 20 筆明細（總共 {len(df_input):,} 筆原始資料）")
    st.markdown('</div>', unsafe_allow_html=True)

    if mode == "Safety Stock Review":
        mtd_date = st.date_input(
            "📅 MTD 日期 (Sold Qty CM-MTD 的截止日期)",
            value=datetime.now().date() - timedelta(days=1),
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
            st.markdown("**Overstock Rate** — 各 Shop Nature 庫存上限計算倍率。")
            col_os1, col_os2, col_os3 = st.columns(3)
            with col_os1:
                os_rate_tourist = st.slider("Overstock Tourist Rate (%)", 0, 100, 30, 1, key="os_tourist_cb",
                    help="Tourist 類型 Overstock 上限 = floor(Monthly Sold × Rate)")
            with col_os2:
                os_rate_mix = st.slider("Overstock Mix Rate (%)", 0, 100, 40, 1, key="os_mix_cb",
                    help="Mix 類型 Overstock 上限 = floor(Monthly Sold × Rate)")
            with col_os3:
                os_rate_non_tourist = st.slider("Overstock Non Tourist Rate (%)", 0, 100, 50, 1, key="os_non_tourist_cb",
                    help="Non Tourist 類型 Overstock 上限 = floor(Monthly Sold × Rate)")

            st.markdown("<hr style='margin: 1rem 0;'>", unsafe_allow_html=True)
            st.markdown("**OOS Rate** — 各 Shop Nature OOS 底線計算倍率。")
            col_oos1, col_oos2, col_oos3 = st.columns(3)
            with col_oos1:
                oos_rate_tourist = st.slider("OOS Tourist Rate (%)", 0, 100, 50, 1, key="oos_tourist_cb",
                    help="Tourist 類型 OOS 底線 = ceil(Projected Monthly × Rate)")
            with col_oos2:
                oos_rate_mix = st.slider("OOS Mix Rate (%)", 0, 100, 50, 1, key="oos_mix_cb",
                    help="Mix 類型 OOS 底線 = ceil(Projected Monthly × Rate)")
            with col_oos3:
                oos_rate_non_tourist = st.slider("OOS Non Tourist Rate (%)", 0, 100, 50, 1, key="oos_non_tourist_cb",
                    help="Non Tourist 類型 OOS 底線 = ceil(Projected Monthly × Rate)")

            st.caption(f"Overstock → Tourist: **{os_rate_tourist}%** | Mix: **{os_rate_mix}%** | Non Tourist: **{os_rate_non_tourist}%**")
            st.caption(f"OOS → Tourist: **{oos_rate_tourist}%** | Mix: **{oos_rate_mix}%** | Non Tourist: **{oos_rate_non_tourist}%**")

        rates_overstock = {
            "tourist": os_rate_tourist / 100.0,
            "mix": os_rate_mix / 100.0,
            "default": os_rate_non_tourist / 100.0,
        }
        rates_oos = {
            "tourist": oos_rate_tourist / 100.0,
            "mix": oos_rate_mix / 100.0,
            "default": oos_rate_non_tourist / 100.0,
        }

        SAFETY_MODE_OPTIONS = ["只可增加", "可加可減"]
        safety_mode = st.radio(
            "🔒 Safety Qty 計算模式",
            options=SAFETY_MODE_OPTIONS,
            index=0,
            horizontal=True,
            help="只可增加：Safety Qty 只增不減；可加可減：允許減少。"
        )
        safety_mode_value = "increase_only" if safety_mode == "只可增加" else "allow_both"

        _all_site_codes = sorted(df_input["Site Code"].dropna().unique()) if "Site Code" in df_input.columns else []

        if st.button("⚙️ 執行報表轉換與分析", type="primary", width='stretch'):
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

                summary_df, raw_df, categories, _warnings = transform_combined(
                    df_input,
                    progress_callback=_cb_transform,
                    shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                    rates_overstock=rates_overstock,
                    rates_oos=rates_oos,
                    mtd_days=mtd_days,
                    month_total_days=month_total_days,
                    safety_mode=safety_mode_value,
                )
                buf = generate_excel_combined(summary_df, raw_df, categories, progress_callback=_cb_excel)
                st.session_state.combined_processed = True
                st.session_state.combined_summary_df = summary_df
                st.session_state.combined_raw_df = raw_df
                st.session_state.combined_categories = categories
                st.session_state.combined_excel_data = buf
                st.session_state.combined_warnings = _warnings
            except Exception as e:
                _prog2.empty()
                _tb = traceback.format_exc()
                print(f"❌ Safety Stock Review 處理失敗: {e}\n{_tb}")
                st.error(f"❌ Safety Stock Review 處理失敗: {e}\n\n詳細錯誤: {_tb}")
                st.stop()

            _prog2.empty()
            st.success("🎉 Safety Stock Review 報表產生成功！數據彙總已完成。")
            if _warnings:
                for _w in _warnings:
                    st.warning(f"⚠️ {_w}")

        if st.session_state.combined_processed:
            summary_df = st.session_state.combined_summary_df
            raw_df = st.session_state.combined_raw_df
            categories = st.session_state.combined_categories
            excel_data = st.session_state.combined_excel_data

            st.markdown('<div class="custom-card" style="border-top: 4px solid #6366f1;">', unsafe_allow_html=True)
            st.markdown("### 📈 Safety Stock 數據分析摘要 (Safety Stock Analysis Dashboard)")

            _total_row = summary_df[summary_df["Buyer"] == "總計"]
            if _total_row.empty:
                st.error("❌ 摘要資料中未找到總計列，無法顯示分析看板。")
                st.stop()
            total_combined = int(_total_row["總計"].values[0])
            total_buyers = len(summary_df[summary_df["Buyer"] != "總計"]["Buyer"].dropna().unique())
            total_brands = len(summary_df[summary_df["Buyer"] != "總計"]["Brand"].dropna().unique())
            total_skus = raw_df["SKU"].dropna().nunique()

            cat_counts = {}
            for cat in categories:
                cat_counts[cat] = int(_total_row[cat].values[0]) if cat in _total_row.columns else 0

            safety_change_total = int(pd.to_numeric(raw_df.get("Safety Qty Change", pd.Series(0, index=raw_df.index)), errors="coerce").fillna(0).sum())

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff5555;">
                    <div class="metric-title">總 Safety Stock 審查行數</div>
                    <div class="metric-value" style="color: #ff5555;">{total_combined:,}</div>
                    <div class="metric-desc">所有 Safety Stock 審查之總計行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #5599ff;">
                    <div class="metric-title">影響採購人員數 (Buyers)</div>
                    <div class="metric-value" style="color: #5599ff;">{total_buyers:,}</div>
                    <div class="metric-desc">位採購人員名下有需審查品項</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #9944ff;">
                    <div class="metric-title">影響品牌總數 (Brands)</div>
                    <div class="metric-value" style="color: #9944ff;">{total_brands:,}</div>
                    <div class="metric-desc">個品牌包含需審查商品</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #00ff88;">
                    <div class="metric-title">影響SKU總數</div>
                    <div class="metric-value" style="color: #00ff88;">{total_skus:,}</div>
                    <div class="metric-desc">個包含需審查商品(SKUs)</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            r2_col1, r2_col2, r2_col3, r2_col4 = st.columns(4)
            with r2_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff9944;">
                    <div class="metric-title">Overstock 分類行數</div>
                    <div class="metric-value" style="color: #ff9944;">{int(cat_counts.get(categories[0], 0)):,}</div>
                    <div class="metric-desc">庫存過剩分類行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff44ff;">
                    <div class="metric-title">OOS 分類行數</div>
                    <div class="metric-value" style="color: #ff44ff;">{int(cat_counts.get(categories[1], 0)):,}</div>
                    <div class="metric-desc">缺貨分類行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #06b6d4;">
                    <div class="metric-title">Safety 調整行數</div>
                    <div class="metric-value" style="color: #06b6d4;">{int(cat_counts.get(categories[2], 0)):,}</div>
                    <div class="metric-desc">Balanced (Safety Adjusted) 行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #8b5cf6;">
                    <div class="metric-title">Safety 調整總量</div>
                    <div class="metric-value" style="color: #8b5cf6;">{safety_change_total:,}</div>
                    <div class="metric-desc">Safety Qty Change 總和</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

            selected_site_code = st.selectbox(
                "🏪 選擇 Site Code 檢視明細",
                options=["全部"] + _all_site_codes,
                index=0,
                help="選擇特定 Site Code 過濾下方表格。僅影響 RAW 明細預覽。"
            )

            tab1, tab2, tab3 = st.tabs(["📋 統計彙總表 (Shop Stock Status)", "🏪 By Shop 明細", "📄 RAW 原始明細"])

            with tab1:
                st.markdown("#### 採購人員與品牌之 Safety Stock 交叉統計表")
                st.markdown("本表即為 Excel 導出檔中「Shop Stock Status」工作表之即時數據：")
                st.dataframe(summary_df, width='stretch', hide_index=True)

            with tab2:
                by_shop_df = summary_df  # placeholder, real by_shop is built in excel gen
                st.markdown("#### 各 Site 層級彙總分析")
                _raw_for_shop = raw_df.copy()
                _by_shop_display = None
                try:
                    from transformer import _build_combined_by_shop_summary
                    _by_shop_display = _build_combined_by_shop_summary(_raw_for_shop)
                except Exception:
                    pass
                if _by_shop_display is not None:
                    st.dataframe(_by_shop_display, width='stretch', hide_index=True)
                else:
                    st.info("By Shop 明細無法產生，請檢查資料。")

            with tab3:
                st.markdown("#### RAW 原始資料庫存明細")
                _filtered_raw = raw_df.copy()
                if selected_site_code != "全部" and "Site Code" in _filtered_raw.columns:
                    _filtered_raw = _filtered_raw[_filtered_raw["Site Code"].astype(str).str.strip() == selected_site_code]
                st.dataframe(_filtered_raw.head(50), width='stretch')
                st.caption(f"僅顯示前 50 筆（{'全部' if selected_site_code == '全部' else selected_site_code}），完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 Safety Stock Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 Safety Stock Review 分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )
    elif mode == "Stock Impact Analysis":
        st.caption("📦 以 **Shop Stock Qty = 0（缺貨）** 為核心，分析存貨對 Sold Qty 的因果影響（無貨就無銷售）。重點關注 Tourist 店舖。")

        warehouse_only = st.checkbox(
            "🏗️ 只分析 Warehouse SKU（建議）",
            value=True,
            help="Warehouse SKU 受 D001 倉庫補貨影響，缺貨分析最具意義。取消勾選可納入 Vendor SKU。"
        )
        st.caption("💡 Vendor SKU 無 D001 倉庫概念，缺貨多為供應商因素而非存貨配置問題。")

        mtd_date = st.date_input(
            "📅 參考日期 (Sold Qty CM-MTD 的截止日期)",
            value=datetime.now().date() - timedelta(days=1),
            help="僅作 Sold Qty (CM-MTD) 語意參考。本模式直接使用原始 CM-MTD 數值，不做全月預估投射。"
        )
        st.caption(f"參考日期: {mtd_date.strftime('%Y-%m-%d')}（CM = 當月至今 {mtd_date.day} 日）")

        st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

        if st.button("⚙️ 執行缺貨銷售影響分析", type="primary", width='stretch'):
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

                summary_site, summary_brand, summary_buyer, summary_category, sku_allocation, sku_diagnosis, raw_df, _warnings = transform_stock_impact(
                    df_input,
                    progress_callback=_cb_transform,
                    warehouse_only=warehouse_only,
                )
                buf = generate_excel_stock_impact(summary_site, summary_brand, summary_buyer, summary_category, sku_allocation, sku_diagnosis, raw_df, progress_callback=_cb_excel)

                st.session_state.stock_impact_processed = True
                st.session_state.stock_impact_summary_site = summary_site
                st.session_state.stock_impact_summary_brand = summary_brand
                st.session_state.stock_impact_summary_buyer = summary_buyer
                st.session_state.stock_impact_summary_category = summary_category
                st.session_state.stock_impact_sku_allocation = sku_allocation
                st.session_state.stock_impact_sku_diagnosis = sku_diagnosis
                st.session_state.stock_impact_raw_df = raw_df
                st.session_state.stock_impact_excel_data = buf
                st.session_state.stock_impact_warnings = _warnings
            except Exception as e:
                _prog2.empty()
                _tb = traceback.format_exc()
                print(f"❌ Stock Impact Analysis 處理失敗: {e}\n{_tb}")
                st.error(f"❌ Stock Impact Analysis 處理失敗: {e}\n\n詳細錯誤: {_tb}")
                st.stop()

            _prog2.empty()
            st.success("🎉 缺貨銷售影響分析完成！")
            if _warnings:
                for _w in _warnings:
                    st.warning(f"⚠️ {_w}")

        if st.session_state.stock_impact_processed:
            summary_site = st.session_state.stock_impact_summary_site
            summary_brand = st.session_state.stock_impact_summary_brand
            summary_buyer = st.session_state.stock_impact_summary_buyer
            summary_category = st.session_state.stock_impact_summary_category
            sku_allocation = st.session_state.stock_impact_sku_allocation
            sku_diagnosis = st.session_state.stock_impact_sku_diagnosis
            raw_df = st.session_state.stock_impact_raw_df
            excel_data = st.session_state.stock_impact_excel_data

            impacted = raw_df[raw_df["Is Stock Impacted"]].copy()
            total_impacted = len(impacted)

            from transformer import _safe_numeric, _safe_default_series
            nature_counts = {"Tourist": 0, "Mix": 0, "Non Tourist": 0}
            if total_impacted and "Shop Nature" in impacted.columns:
                for nat, cnt in impacted["Shop Nature"].value_counts(dropna=False).items():
                    if nat in nature_counts:
                        nature_counts[nat] = int(cnt)

            loss_m2 = int(_safe_numeric(impacted.get("Total Sales Loss (M-2)", _safe_default_series(0, impacted.index))).sum()) if total_impacted else 0
            loss_p90d = int(_safe_numeric(impacted.get("Total Sales Loss (P90D)", _safe_default_series(0, impacted.index))).sum()) if total_impacted else 0
            tourist_pct = (nature_counts["Tourist"] / total_impacted * 100) if total_impacted else 0

            # 受影響最嚴重 Buyer
            top_buyer_name = "—"
            top_buyer_loss = 0
            if summary_buyer is not None and not summary_buyer.empty and "Buyer" in summary_buyer.columns:
                _buyer_data = summary_buyer[~summary_buyer["Buyer"].isin(["總計", ""])]
                if not _buyer_data.empty and "最大銷售損失(P90D)" in _buyer_data.columns:
                    _top = _buyer_data.sort_values("最大銷售損失(P90D)", ascending=False).iloc[0]
                    top_buyer_name = str(_top["Buyer"])
                    top_buyer_loss = int(_top.get("最大銷售損失(P90D)", 0))

            st.markdown('<div class="custom-card" style="border-top: 4px solid #dc2626;">', unsafe_allow_html=True)
            st.markdown("### 📦 缺貨銷售影響分析摘要 (Stock Impact Dashboard)")

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #EF4444;">
                    <div class="metric-title">總缺貨影響行數</div>
                    <div class="metric-value" style="color: #EF4444;">{total_impacted:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">行</span></div>
                    <div class="metric-desc">Shop Stock=0 且有需求（只計 RF 店舖）｜單位：行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #DC2626;">
                    <div class="metric-title">⭐ Tourist 缺貨</div>
                    <div class="metric-value" style="color: #DC2626;">{nature_counts['Tourist']:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">行</span></div>
                    <div class="metric-desc">公司最重視的店舖類型｜單位：行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #F97316;">
                    <div class="metric-title">Mix 缺貨</div>
                    <div class="metric-value" style="color: #F97316;">{nature_counts['Mix']:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">行</span></div>
                    <div class="metric-desc">次重視店舖類型｜單位：行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #F59E0B;">
                    <div class="metric-title">Non Tourist 缺貨</div>
                    <div class="metric-value" style="color: #F59E0B;">{nature_counts['Non Tourist']:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">行</span></div>
                    <div class="metric-desc">一般店舖類型｜單位：行數</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            r2_col1, r2_col2, r2_col3, r2_col4 = st.columns(4)
            with r2_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3B82F6;">
                    <div class="metric-title">總銷售損失 (M-2基準)</div>
                    <div class="metric-value" style="color: #3B82F6;">{loss_m2:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">件</span></div>
                    <div class="metric-desc">以 M-2 Sold 為需求基準｜單位：件數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #8B5CF6;">
                    <div class="metric-title">總銷售損失 (P90D基準)</div>
                    <div class="metric-value" style="color: #8B5CF6;">{loss_p90d:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">件</span></div>
                    <div class="metric-desc">以 90 天月均為需求基準｜單位：件數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #06B6D4;">
                    <div class="metric-title">受影響最嚴重 Buyer</div>
                    <div class="metric-value" style="color: #06B6D4; font-size: 1.1rem;">{top_buyer_name}</div>
                    <div class="metric-desc">最大損失(P90D): {top_buyer_loss:,} 件</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #1E40AF;">
                    <div class="metric-title">Tourist 缺貨占比</div>
                    <div class="metric-value" style="color: #1E40AF;">{tourist_pct:.1f}%</div>
                    <div class="metric-desc">Tourist 缺貨行數 / 總缺貨行數</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

            impact_tabs = st.tabs(["🏪 By Site", "🏷️ By Brand", "👤 By Buyer", "📦 By Category", "🔄 SKU 調貨建議", "🔍 SKU 缺貨診斷"])

            with impact_tabs[0]:
                st.markdown("#### Site 層級缺貨影響分析")
                if summary_site is not None and not summary_site.empty:
                    st.dataframe(summary_site, width='stretch', hide_index=True)
                else:
                    st.info("無 Site 層級缺貨影響數據")

            with impact_tabs[1]:
                st.markdown("#### Brand 層級缺貨影響分析")
                if summary_brand is not None and not summary_brand.empty:
                    st.dataframe(summary_brand, width='stretch', hide_index=True)
                else:
                    st.info("無 Brand 層級缺貨影響數據")

            with impact_tabs[2]:
                st.markdown("#### Buyer 層級缺貨影響分析")
                if summary_buyer is not None and not summary_buyer.empty:
                    st.dataframe(summary_buyer, width='stretch', hide_index=True)
                else:
                    st.info("無 Buyer 層級缺貨影響數據")

            with impact_tabs[3]:
                st.markdown("#### Category (Adj. MCH1) 層級缺貨影響分析")
                if summary_category is not None and not summary_category.empty:
                    st.dataframe(summary_category, width='stretch', hide_index=True)
                elif "Adj. MCH1" not in raw_df.columns:
                    st.info("來源檔案缺少 Adj. MCH1 欄位，無法產出 By Category 分析")
                else:
                    st.info("無 Category 層級缺貨影響數據")

            with impact_tabs[4]:
                st.markdown("#### 🔄 SKU 跨 Nature 調貨建議")
                st.caption("Tourist 店舖缺貨但同 SKU 在 Non-Tourist / Mix 店舖仍有存貨 → 揪出存貨配置不當，建議調貨支援。")
                if sku_allocation is not None and not sku_allocation.empty:
                    st.dataframe(sku_allocation, width='stretch', hide_index=True)
                else:
                    st.info("無可調貨 SKU（無 Tourist 缺貨且同時其他 Nature 有存貨的情況）")

            with impact_tabs[5]:
                st.markdown("#### 🔍 by-SKU 缺貨診斷（Tourist + Mix）")
                st.caption("每個缺貨 SKU 的 Tourist/Mix 缺貨店數、D001 倉庫三期貨量與店舖總存貨。🟧 **倉庫有貨但店舖缺貨** = 補貨流程問題；⬜ **源頭斷貨** = 倉庫也斷貨需上游補貨。")
                if sku_diagnosis is not None and not sku_diagnosis.empty:
                    # 統計摘要
                    d001_restock = int((sku_diagnosis["D001 狀態"] == "倉庫有貨但店舖缺貨").sum())
                    d001_source = int((sku_diagnosis["D001 狀態"] == "源頭斷貨").sum())
                    st.markdown(f"**共 {len(sku_diagnosis):,} 個缺貨 SKU** — 🟧 補貨問題: {d001_restock:,} 個｜⬜ 源頭斷貨: {d001_source:,} 個")
                    st.dataframe(sku_diagnosis, width='stretch', hide_index=True)
                else:
                    st.info("無 Tourist/Mix 缺貨 SKU")

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 Stock Impact Analysis_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載缺貨銷售影響分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )
    else:
        if mode == "OOS Review":
            oos_calc_mode = st.radio(
                "📊 Safety Qty 計算基準",
                options=["MTD (Sold Qty CM-MTD)", "P90D (Shop Monthly Sold Qty)"],
                index=0,
                horizontal=True,
                help="MTD：使用本月至今銷售（CM-MTD）推算全月；P90D：使用過去90天月均銷售（P90D）"
            )
            oos_calc_mode_value = "mtd" if oos_calc_mode.startswith("MTD") else "p90d"

            if oos_calc_mode_value == "mtd":
                mtd_date = st.date_input(
                    "📅 MTD 日期 (Sold Qty CM-MTD 的截止日期)",
                    value=datetime.now().date() - timedelta(days=1),
                    help="選擇 MTD (Month-To-Date) 截止日期，系統將據此推算月度預估銷售量。"
                )
            else:
                mtd_date = st.date_input(
                    "📅 MTD 日期 (僅供 Projected Monthly 顯示)",
                    value=datetime.now().date() - timedelta(days=1),
                    help="Projected Monthly Sold Qty 欄位顯示用，不影響 Safety Qty 建議計算。"
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

        _is_overstock_mode = (mode == "Overstock Review")
        _is_carry_mode = (mode == "Carry SKU Review")
        # Modes that don't use the Shop Nature Rate sliders (no Safety Qty calculation):
        # Carry SKU Review (pivot only) and SKU 精簡 Review (full 退倉, no rate needed).
        _skip_rates_mode = _is_carry_mode or (mode == "SKU 精簡 Review")

        if not _skip_rates_mode:
            _tourist_default = 30 if _is_overstock_mode else 50
            _mix_default = 40 if _is_overstock_mode else 50
            _non_tourist_default = 50

            with st.expander("⚙️ 自定義 Shop Nature Rate (進階設定)", expanded=False):
                st.markdown("調整各 Shop Nature 類型的安全庫存建議倍率（% of Monthly Sold Qty）。")
                _rate_cols = st.columns(3)
                with _rate_cols[0]:
                    rate_tourist = st.slider("Tourist Rate (%)", 0, 100, _tourist_default, 1,
                        help="Tourist 類型：安全庫存建議 = floor(Monthly Sold × Rate)，最低 3")
                with _rate_cols[1]:
                    rate_mix = st.slider("Mix Rate (%)", 0, 100, _mix_default, 1,
                        help="Mix 類型：安全庫存建議 = floor(Monthly Sold × Rate)，最低 3")
                with _rate_cols[2]:
                    rate_non_tourist = st.slider("Non Tourist Rate (%)", 0, 100, _non_tourist_default, 1,
                        help="Non Tourist 及其他類型：安全庫存建議 = floor(Monthly Sold × Rate)，最低 3")
                st.caption(f"目前設定 → Tourist: **{rate_tourist}%** | Mix: **{rate_mix}%** | Non Tourist: **{rate_non_tourist}%**")

            rates = {
                "tourist": rate_tourist / 100.0,
                "mix": rate_mix / 100.0,
                "default": rate_non_tourist / 100.0,
            }
        else:
            rates = {"tourist": 0.5, "mix": 0.5, "default": 0.5}

        safety_mode_value = "increase_only"
        if mode == "OOS Review":
            SAFETY_MODE_OPTIONS = ["只可增加", "可加可減"]
            safety_mode = st.radio(
                "🔒 Safety Qty 計算模式",
                options=SAFETY_MODE_OPTIONS,
                index=0,
                horizontal=True,
                help="只可增加：Safety Qty 只增不減（不處理 Overstock 的 Safety）；可加可減：同步處理 Overstock 的 Safety，可升可降。"
            )
            safety_mode_value = "increase_only" if safety_mode == "只可增加" else "allow_both"

        if st.button("⚙️ 執行報表轉換與分析", type="primary", width='stretch'):
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
                elif mode == "Carry SKU Review":
                    summary_left, summary_right, dashboard_df, raw_df, _warnings = transform_carry(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                    )
                    buf = generate_excel_carry(summary_left, summary_right, dashboard_df, raw_df, progress_callback=_cb_excel)
                    st.session_state.carry_processed = True
                    st.session_state.carry_summary_left = summary_left
                    st.session_state.carry_summary_right = summary_right
                    st.session_state.carry_dashboard_df = dashboard_df
                    st.session_state.carry_raw_df = raw_df
                    st.session_state.carry_excel_data = buf
                    st.session_state.carry_warnings = _warnings
                elif mode == "SKU 精簡 Review":
                    summary_df, raw_df, categories, _warnings = transform_nd(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                    )
                    buf = generate_excel_nd(summary_df, raw_df, categories, progress_callback=_cb_excel)
                    st.session_state.nd_processed = True
                    st.session_state.nd_summary_df = summary_df
                    st.session_state.nd_raw_df = raw_df
                    st.session_state.nd_categories = categories
                    st.session_state.nd_excel_data = buf
                    st.session_state.nd_warnings = _warnings
                else:  # OOS Review
                    summary_df, raw_df, categories, _warnings = transform_oos(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                        rates=rates,
                        mtd_days=mtd_days,
                        month_total_days=month_total_days,
                        safety_mode=safety_mode_value,
                        oos_calc_mode=oos_calc_mode_value,
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
                _tb = traceback.format_exc()
                print(f"❌ 報表處理失敗: {e}\n{_tb}")
                st.error(f"❌ 報表處理失敗: {e}\n\n詳細錯誤: {_tb}")
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
                    <div class="metric-title">總 Overstock 行數</div>
                    <div class="metric-value" style="color: #ff5555;">{int(total_overstock):,}</div>
                    <div class="metric-desc">所有庫存異常之總計行數</div>
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
            if transfer_col_resolved is not None and transfer_col_resolved in _total_row.columns:
                transfer_qty_total = int(_total_row[transfer_col_resolved].values[0])
            elif transfer_col_resolved is not None and transfer_col_resolved in raw_df.columns:
                transfer_qty_total = int(pd.to_numeric(raw_df[transfer_col_resolved], errors="coerce").fillna(0).sum())
            else:
                transfer_qty_total = 0

            r2_col1, r2_col2, r2_col3 = st.columns(3)
            with r2_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff9944;">
                    <div class="metric-title">P90D 0 Sales</div>
                    <div class="metric-value" style="color: #ff9944;">{int(cat_counts.get(categories[0], 0)):,}</div>
                    <div class="metric-desc">「過去 90 天內零銷售」之異常行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff44ff;">
                    <div class="metric-title">Overstock (>2*Monthly Sold Qty & >2pcs)</div>
                    <div class="metric-value" style="color: #ff44ff;">{int(cat_counts.get(categories[1], 0)):,}</div>
                    <div class="metric-desc">「大於月均銷售量兩倍且庫存大於 2 件」之異常行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #00d4ff;">
                    <div class="metric-title">調整後的預計可退倉總數</div>
                    <div class="metric-value" style="color: #00d4ff;">{transfer_qty_total:,}</div>
                    <div class="metric-desc">After new Safety Qty / ND can transfer out of TOTAL QTY</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            tab1, tab2 = st.tabs(["📋 統計彙總表 (Shop Stock Status)", "📄 原始明細資料預覽 (RAW)"])

            with tab1:
                st.markdown("#### 採購人員與品牌之庫存異常交叉統計表")
                st.markdown("本表即為 Excel 導出檔中「Shop Stock Status」工作表之即時數據：")
                st.dataframe(summary_df, width='stretch', hide_index=True)

            with tab2:
                st.markdown("#### RAW 原始資料庫存明細 (前 50 筆)")
                st.markdown("此表為過濾與重整後，完整導出之明細資料數據預覽：")
                st.dataframe(raw_df.head(50), width='stretch')
                st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 by SKUxShop_To Buyer_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載完整分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "OOS Review" and st.session_state.oos_processed:
            summary_df = st.session_state.oos_summary_df
            raw_df = st.session_state.oos_raw_df
            categories = st.session_state.oos_categories
            excel_data = st.session_state.oos_excel_data

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
                    <div class="metric-title">總 OOS 行數</div>
                    <div class="metric-value" style="color: #ff5555;">{total_oos:,}</div>
                    <div class="metric-desc">所有 OOS 異常之總計行數</div>
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
                    <div class="metric-desc">現有庫存為 0 之 OOS 行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff44ff;">
                    <div class="metric-title">Potential OOS (<7 Turnday)</div>
                    <div class="metric-value" style="color: #ff44ff;">{int(cat_counts.get(categories[1], 0)):,}</div>
                    <div class="metric-desc">Turnday 不足 7 之潛在 OOS 行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #06b6d4;">
                    <div class="metric-title">Normal (Safety Qty Insufficient)</div>
                    <div class="metric-value" style="color: #06b6d4;">{int(cat_counts.get(categories[2], 0)):,}</div>
                    <div class="metric-desc">正常庫存但安全庫存不足行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #8b5cf6;">
                    <div class="metric-title">建議 Safety Qty 修改</div>
                    <div class="metric-value" style="color: #8b5cf6;">{safety_increase_total:,}</div>
                    <div class="metric-desc">建議 Safety Qty 修改 之總量</div>
                </div>
                """, unsafe_allow_html=True)

            tab1, tab2 = st.tabs(["📋 統計彙總表 (Shop Stock Status)", "📄 原始明細資料預覽 (RAW)"])

            with tab1:
                st.markdown("#### 採購人員與品牌之 OOS 交叉統計表")
                st.markdown("本表即為 Excel 導出檔中「Shop Stock Status」工作表之即時數據：")
                st.dataframe(summary_df, width='stretch', hide_index=True)

            with tab2:
                st.markdown("#### RAW 原始資料庫存明細 (前 50 筆)")
                st.markdown("此表為過濾與重整後，完整導出之明細資料數據預覽：")
                st.dataframe(raw_df.head(50), width='stretch')
                st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 OOS Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 OOS 分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "SKU 精簡 Review" and st.session_state.nd_processed:
            summary_df = st.session_state.nd_summary_df
            raw_df = st.session_state.nd_raw_df
            categories = st.session_state.nd_categories
            excel_data = st.session_state.nd_excel_data

            st.markdown('<div class="custom-card" style="border-top: 4px solid #6366f1;">', unsafe_allow_html=True)
            st.markdown("### 📈 ND SKU 數據分析摘要 (ND SKU Analysis Dashboard)")

            _total_row = summary_df[summary_df["Buyer"] == "總計"]
            if _total_row.empty:
                st.error("❌ 摘要資料中未找到總計列，無法顯示分析看板。")
                st.stop()
            total_nd = int(_total_row["總計"].values[0])

            # Warehouse (可退倉) vs Vendor (另有安排) breakdown.
            _nd_rp = raw_df.get("RP Type", pd.Series("", index=raw_df.index)).fillna("").astype(str).str.strip().str.upper()
            _nd_ss = raw_df.get("Supply Source", pd.Series("", index=raw_df.index)).fillna("").astype(str).str.strip().str.lower()
            _nd_mask = (_nd_rp == "ND")
            _nd_wh_mask = _nd_mask & (_nd_ss == "warehouse")
            _nd_vendor_mask = _nd_mask & (_nd_ss == "vendor")
            nd_warehouse_rows = int(_nd_wh_mask.sum())
            nd_vendor_rows = int(_nd_vendor_mask.sum())

            # Buyers / Brands / SKU totals must count ONLY affected ND rows (raw_df keeps
            # all cleaned rows incl. non-ND, so counting over raw_df would over-count).
            _nd_df = raw_df[_nd_mask]
            total_buyers = _nd_df["Buyer"].dropna().nunique() if len(_nd_df) else 0
            total_brands = _nd_df["Brand"].dropna().nunique() if len(_nd_df) else 0
            total_skus = _nd_df["SKU"].dropna().nunique() if len(_nd_df) else 0

            _transfer_col_name = "After new Safety Qty / ND can transfer out of QTY"
            transfer_col_resolved = _resolve_column_name(_transfer_col_name, set(summary_df.columns))
            if transfer_col_resolved is not None and transfer_col_resolved in _total_row.columns:
                transfer_qty_total = int(_total_row[transfer_col_resolved].values[0])
            elif transfer_col_resolved is not None and transfer_col_resolved in raw_df.columns:
                transfer_qty_total = int(pd.to_numeric(raw_df[transfer_col_resolved], errors="coerce").fillna(0).sum())
            else:
                transfer_qty_total = 0

            # 另有安排 total qty = sum of Shop Stock Qty over ND+Vendor rows (cannot be returned).
            _shop_stock_col = _resolve_column_name("Shop Stock Qty (CM)", set(raw_df.columns))
            other_arranged_qty_total = 0
            if _shop_stock_col:
                other_arranged_qty_total = int(pd.to_numeric(raw_df.loc[_nd_vendor_mask, _shop_stock_col], errors="coerce").fillna(0).sum())

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff5555;">
                    <div class="metric-title">總 ND SKU 行數</div>
                    <div class="metric-value" style="color: #ff5555;">{total_nd:,}</div>
                    <div class="metric-desc">所有 ND SKU 總計行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #00ff88;">
                    <div class="metric-title">Warehouse 可退倉行數</div>
                    <div class="metric-value" style="color: #00ff88;">{nd_warehouse_rows:,}</div>
                    <div class="metric-desc">Supply Source=Warehouse 可退倉之 ND 行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff9944;">
                    <div class="metric-title">Vendor 另有安排行數</div>
                    <div class="metric-value" style="color: #ff9944;">{nd_vendor_rows:,}</div>
                    <div class="metric-desc">Supply Source=Vendor 不可退倉、另有安排之 ND 行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #5599ff;">
                    <div class="metric-title">影響採購人員數 (Buyers)</div>
                    <div class="metric-value" style="color: #5599ff;">{total_buyers:,}</div>
                    <div class="metric-desc">位採購人員名下有 ND SKU</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            r2_col1, r2_col2, r2_col3, r2_col4 = st.columns(4)
            with r2_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #9944ff;">
                    <div class="metric-title">影響品牌總數 (Brands)</div>
                    <div class="metric-value" style="color: #9944ff;">{total_brands:,}</div>
                    <div class="metric-desc">個品牌包含 ND SKU</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">影響SKU總數</div>
                    <div class="metric-value" style="color: #10b981;">{total_skus:,}</div>
                    <div class="metric-desc">個包含 ND SKU (SKUs)</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #00d4ff;">
                    <div class="metric-title">預計可退倉總數量</div>
                    <div class="metric-value" style="color: #00d4ff;">{transfer_qty_total:,}</div>
                    <div class="metric-desc">Warehouse 可退倉數量總和</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f97316;">
                    <div class="metric-title">預計另有安排總數量</div>
                    <div class="metric-value" style="color: #f97316;">{other_arranged_qty_total:,}</div>
                    <div class="metric-desc">Vendor 另有安排之庫存數量總和</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            _all_site_codes = sorted(df_input["Site Code"].dropna().unique()) if "Site Code" in df_input.columns else []
            selected_site_code = st.selectbox(
                "🏪 選擇 Site Code 檢視明細",
                options=["全部"] + _all_site_codes,
                index=0,
                help="選擇特定 Site Code 過濾下方 RAW 明細。僅影響 RAW 明細預覽。"
            )

            tab1, tab2 = st.tabs(["📋 統計彙總表 (Shop Stock Status)", "📄 原始明細資料預覽 (RAW)"])

            with tab1:
                st.markdown("#### 採購人員與品牌之 ND SKU 交叉統計表")
                st.markdown("ND SKU 依 Supply Source × Shop Stock Status 雙層動態分類（Warehouse 可退倉 / Vendor 另有安排）。本表即為 Excel 導出檔中「Shop Stock Status」工作表之即時數據：")
                st.dataframe(summary_df, width='stretch', hide_index=True)

            with tab2:
                st.markdown("#### RAW 原始資料庫存明細 (前 50 筆)")
                st.markdown("此表為過濾與重整後，完整導出之明細資料數據預覽（RP Type=ND：Warehouse 建議退倉、Vendor 標記「另有安排」）：")
                _filtered_raw = raw_df.copy()
                if selected_site_code != "全部" and "Site Code" in _filtered_raw.columns:
                    _filtered_raw = _filtered_raw[_filtered_raw["Site Code"].astype(str).str.strip() == selected_site_code]
                st.dataframe(_filtered_raw.head(50), width='stretch')
                st.caption(f"僅顯示前 50 筆（{'全部' if selected_site_code == '全部' else selected_site_code}），完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 SKU 精簡 Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 SKU 精簡 Review 分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "Carry SKU Review" and st.session_state.carry_processed:
            summary_left = st.session_state.carry_summary_left
            summary_right = st.session_state.carry_summary_right
            dashboard_df = st.session_state.carry_dashboard_df
            raw_df = st.session_state.carry_raw_df
            excel_data = st.session_state.carry_excel_data

            st.markdown('<div class="custom-card" style="border-top: 4px solid #6366f1;">', unsafe_allow_html=True)
            st.markdown("### 📈 Carry SKU 分析摘要 (Carry SKU Analysis Dashboard)")

            _total_row = dashboard_df[dashboard_df["Site Code"] == "總計"]
            if not _total_row.empty:
                tr = _total_row.iloc[0]
            else:
                tr = None

            shop_count = len(dashboard_df[dashboard_df["Site Code"] != "總計"])
            total_sku = int(tr["SKU 總數"]) if tr is not None else 0
            total_rf = int(tr["RF SKU 總數"]) if tr is not None else 0
            total_nd = int(tr["ND SKU 總數"]) if tr is not None else 0
            total_p90d0 = int(tr["P90D 0 Sales SKU 總數"]) if tr is not None else 0
            total_stock_val = tr["SKU Stock Value (CM)"] if tr is not None else 0
            total_p90d0_stock_val = tr["P90D 0 Sales Stock Value (CM)"] if tr is not None else 0
            total_suggested_nd = int(tr.get("建議ND SKU 總數", 0)) if tr is not None else 0
            total_suggested_nd_stock_val = tr.get("建議ND Stock Value (CM)", 0) if tr is not None else 0
            total_suggested_nd_stock_qty = tr.get("建議ND Shop Stock Qty (CM)", 0) if tr is not None else 0

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff5555;">
                    <div class="metric-title">影響店舖數</div>
                    <div class="metric-value" style="color: #ff5555;">{shop_count:,}</div>
                    <div class="metric-desc">個店舖含有 Carry SKU</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #5599ff;">
                    <div class="metric-title">Carry SKU 總數</div>
                    <div class="metric-value" style="color: #5599ff;">{total_sku:,}</div>
                    <div class="metric-desc">所有店舖 Carry SKU 總列數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #9944ff;">
                    <div class="metric-title">RF SKU 總數</div>
                    <div class="metric-value" style="color: #9944ff;">{total_rf:,}</div>
                    <div class="metric-desc">RP Type 為 RF 的 Carry SKU</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #00ff88;">
                    <div class="metric-title">ND SKU 總數</div>
                    <div class="metric-value" style="color: #00ff88;">{total_nd:,}</div>
                    <div class="metric-desc">RP Type 為 ND 的 Carry SKU</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            r2_col1, r2_col2, r2_col3 = st.columns(3)
            with r2_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff9944;">
                    <div class="metric-title">P90D 0 Sales SKU 總數</div>
                    <div class="metric-value" style="color: #ff9944;">{total_p90d0:,}</div>
                    <div class="metric-desc">「過去 90 天內零銷售」之 SKU 行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ff44ff;">
                    <div class="metric-title">SKU Stock Value (CM)</div>
                    <div class="metric-value" style="color: #ff44ff;">{total_stock_val:,.0f}</div>
                    <div class="metric-desc">所有 Carry SKU 庫存金額總和</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #00d4ff;">
                    <div class="metric-title">P90D 0 Sales Stock Value (CM)</div>
                    <div class="metric-value" style="color: #00d4ff;">{total_p90d0_stock_val:,.0f}</div>
                    <div class="metric-desc">P90D 0 Sales SKU 的庫存金額總和</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            r3_col1, r3_col2, r3_col3 = st.columns(3)
            with r3_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">建議ND SKU 總數</div>
                    <div class="metric-value" style="color: #10b981;">{total_suggested_nd:,}</div>
                    <div class="metric-desc">建議轉 ND 的 SKU 行數 (RF + P90D 0 Sales)</div>
                </div>
                """, unsafe_allow_html=True)
            with r3_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f97316;">
                    <div class="metric-title">建議ND Stock Value (CM)</div>
                    <div class="metric-value" style="color: #f97316;">{total_suggested_nd_stock_val:,.0f}</div>
                    <div class="metric-desc">建議轉 ND 的 SKU 庫存金額總和</div>
                </div>
                """, unsafe_allow_html=True)
            with r3_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #06b6d4;">
                    <div class="metric-title">建議ND Shop Stock Qty (CM)</div>
                    <div class="metric-value" style="color: #06b6d4;">{total_suggested_nd_stock_qty:,.0f}</div>
                    <div class="metric-desc">建議轉 ND 的 SKU 庫存數量總和</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

            st.markdown("#### 各店舖 Carry SKU 指標總覽")
            st.dataframe(dashboard_df, width='stretch', hide_index=True)

            tab1, tab2 = st.tabs(["📋 Summary 樞紐表預覽", "📄 RAW 明細（前 50 筆）"])
            with tab1:
                st.markdown("#### 左區塊 (Site Code × Adj. MCH1)")
                st.dataframe(summary_left, width='stretch', hide_index=True)
                st.markdown("#### 右區塊 (Site Code × Adj. SST)")
                st.dataframe(summary_right, width='stretch', hide_index=True)
            with tab2:
                st.markdown("#### RAW 原始資料明細（前 50 筆）")
                st.dataframe(raw_df.head(50), width='stretch')
                st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<div style='margin: 1.5rem 0;'></div>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 Carry SKU Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 Carry SKU 分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

# 9. 底部頁尾資訊
st.divider()
st.markdown(
    '<div style="text-align: center; color: #444455; font-size: 0.8rem; padding: 10px 0;">'
    'I14 報表轉換與分析系統 &nbsp;&nbsp;•&nbsp;&nbsp; Streamlit + Zeabur &nbsp;&nbsp;•&nbsp;&nbsp; '
    f'v{__version__} &nbsp;&nbsp;•&nbsp;&nbsp; Developed by Ricky Yue'
    '</div>',
    unsafe_allow_html=True,
)
