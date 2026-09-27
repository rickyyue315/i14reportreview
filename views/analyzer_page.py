import streamlit as st
import pandas as pd
import numpy as np
import calendar
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from transformer import (
    transform, generate_excel, validate_input, _resolve_column_name,
    _display_safety_header,
    transform_oos, generate_excel_oos,
    transform_carry, generate_excel_carry,
    transform_combined, generate_excel_combined,
    transform_nd, generate_excel_nd,
    transform_stock_impact, generate_excel_stock_impact,
    _build_sku_monthly_analysis_simplified, generate_excel_stock_impact_simplified,
    _build_low_stock_item_class_summary, _build_sku_soh_tier_summary,
    _build_sell_through_band_summary,
    transform_ndrf_comparison, generate_excel_ndrf_comparison,
    transform_p90d_zero_sales, generate_excel_p90d_zero_sales,
    transform_stock_checking, generate_excel_stock_checking,
    _stock_checking_raw_display_columns,
    transform_forecast, generate_excel_forecast,
    _forecast_raw_display_columns,
    _FC_TARGET_DAYS_DEFAULT, _FC_TARGET_DAYS_MIN, _FC_TARGET_DAYS_MAX,
    _FC_CAP_MONTHS_DEFAULT, _FC_CAP_MONTHS_MIN, _FC_CAP_MONTHS_MAX,
    _FC_SURGE_MULT_DEFAULT, _FC_SURGE_MULT_MIN, _FC_SURGE_MULT_MAX,
    _FC_SURGE_MIN_BASE_DEFAULT, _FC_SURGE_MIN_BASE_MIN, _FC_SURGE_MIN_BASE_MAX,
    _FC_ACCEL_MULT_DEFAULT, _FC_ACCEL_MULT_MIN, _FC_ACCEL_MULT_MAX,
    _FC_DAMP_DEFAULT, _FC_DAMP_MIN, _FC_DAMP_MAX,
    _FC_WEIGHT_P90D_DEFAULT, _FC_WEIGHT_M1_DEFAULT, _FC_WEIGHT_MTD_DEFAULT,
    _FC_WEIGHT_MIN, _FC_WEIGHT_MAX,
    _FC_TREND_CAP_DEFAULT, _FC_TREND_CAP_MIN, _FC_TREND_CAP_MAX,
    _FC_SURGE_COVER_DEFAULT, _FC_SURGE_COVER_MIN, _FC_SURGE_COVER_MAX,
    _FC_CLASS_NEW, _FC_CLASS_ACCEL, _FC_CLASS_SURGE,
    transform_hybrid, generate_excel_hybrid, _hybrid_dashboard_metrics,
    _hybrid_raw_display_columns, _build_hybrid_transfer_detail,
    _build_hybrid_safety_change_summary, _build_hybrid_safety_sku_detail,
    _build_hybrid_full_safety_change_summary, _build_hybrid_full_safety_sku_detail,
    _build_hybrid_full_safety_detail,
    _build_hybrid_sales_summary,
    _hybrid_buyer_need_stats, _hybrid_buyer_safety_stats,
    _hybrid_buyer_combined_stats, _hybrid_buyer_combined_detail,
    _build_stock_checking_shop_issues, _build_combined_shop_issues,
    _build_hybrid_shop_issues,
    _build_hybrid_return_d001,
    _HYBRID_RETURN_D001_COLUMN, _HYBRID_RETURN_D001_SHEET, _HYBRID_RETURN_D001_COLUMNS,
    _build_combined_dimension_cross_tab,
    _STOCK_CHECKING_SHOP_ISSUE_COLUMNS, _COMBINED_SHOP_ISSUE_COLUMNS,
    _HYBRID_SHOP_ISSUE_COLUMNS,
    _HYBRID_SAFETY_DAYS_DEFAULT, _HYBRID_SAFETY_DAYS_MIN, _HYBRID_SAFETY_DAYS_MAX,
    _hybrid_safety_days_label,
    _HYBRID_DIMENSION_LABELS,
    _hybrid_skew_hint_text,
    _hybrid_shortage_hint_text,
    _HYBRID_SALES_OOS_DAYS_DEFAULT, _HYBRID_SALES_OOS_DAYS_MIN, _HYBRID_SALES_OOS_DAYS_MAX,
    _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT, _HYBRID_SALES_OVERSTOCK_DAYS_MIN, _HYBRID_SALES_OVERSTOCK_DAYS_MAX,
    _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT, _HYBRID_SALES_OVERSTOCK_MIN_UNITS_MIN, _HYBRID_SALES_OVERSTOCK_MIN_UNITS_MAX,
    _optimize_input_dtypes,
    transform_sales_loss, generate_excel_sales_loss,
    transform_excess_capital, generate_excel_excess_capital,
    transform_rp_review, generate_excel_rp_review,
    transform_d001_health, generate_excel_d001_health,
    transform_new_sku, generate_excel_new_sku,
    _HYBRID_FULL_SAFETY_SHEET,
)
import time
import traceback
import gc
from openpyxl import load_workbook
from io import BytesIO
from version import __version__
from views.analyzer_helpers import (
    build_hybrid_params,
    canonical_site_codes,
    filter_by_site_code,
    hybrid_params_changed,
)

def _process_rss_mb():
    """讀取目前行程的實體記憶體用量 (MB)。僅 Linux（如 Zeabur 容器）有效，其他平台回傳 None。"""
    try:
        with open("/proc/self/status", "r") as f:
            for line in f:
                if line.startswith("VmRSS:"):
                    return int(line.split()[1]) / 1024
    except Exception:
        pass
    return None

def _release_os_memory():
    """gc 後再呼叫 malloc_trim 將已釋放記憶體真正歸還作業系統。
    單靠 gc.collect() 只會把物件還給 Python 堆，glibc 不會主動還頁給 OS，RSS 不會下降。"""
    gc.collect()
    try:
        import ctypes
        ctypes.CDLL("libc.so.6").malloc_trim(0)
    except Exception:
        pass


def _top_download_button(excel_data, prefix, key):
    """結果頁置頂下載按鈕（內容與底部下載按鈕相同，免捲動；計算邏輯不變）。"""
    if excel_data is None:
        return
    try:
        _hk = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
    except Exception:
        _hk = datetime.now().strftime("%Y%m%d_%H%M")
    st.download_button(
        label=f"📥 下載{prefix}分析報表 (.xlsx)（置頂）",
        data=excel_data,
        file_name=f"{prefix}_{_hk}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        key=key,
    )


def _style_inventory_skew(df: pd.DataFrame, zero_sales: bool = False, dimension: str = "nature", shortage: bool = True):
    """Summary 顯示：存貨偏重提示非空時，以醒目橘底加強顯示提示欄與偏重分組的庫存佔比欄。
    分組名稱按提示文字解析（Shop Nature 三組／RP Type 兩組／Nature＋RP Type 六組通用）。
    zero_sales=True 時（僅銷量兩維度傳入）另以醒目紅底標示零銷組庫存佔比大於 0% 的儲存格
    （零銷售而仍有存貨，理論上全屬 Overstock；同格與橘底重疊時以紅底優先）。
    shortage=True 時（預設開啟，兩模式全維度生效）另以淺藍底標示存貨偏低分組的庫存佔比欄——
    銷量佔比達標且銷量佔比 − 庫存佔比達差距門檻（門檻較偏重側寬鬆，只標示儲存格，不新增文字欄）；
    與橘底／紅底互斥。
    P0：列數超過 3000 時略過逐列樣式，直接回傳原表，避免 apply(axis=1) 卡頓。"""
    from transformer import _hybrid_shortage_thresholds as _short_thr
    hint_col = "存貨偏重提示"
    suffix = " 加總 - Shop Stock Qty (CM) %"
    if df is None or df.empty:
        return df
    if len(df) > 3000:
        return df
    if hint_col not in df.columns and not zero_sales and not shortage:
        return df
    try:
        _short_share, _short_gap = _short_thr(dimension)
    except Exception:
        _short_share, _short_gap = (0.30, 0.10)

    def _row_styles(row):
        hint = str(row.get(hint_col, "") or "")
        styles = [""] * len(row)
        targets = {}
        if hint.strip():
            targets[hint_col] = "background-color: #FFC000; font-weight: 700"
        # 由欄位反推群組名（不以正則解析提示文字，避免群組名含全形括號時截斷）
        # 先收集各組庫存／月銷加總，供偏低判定計算佔比
        _groups = []
        for col in row.index:
            if isinstance(col, str) and col.endswith(suffix):
                _groups.append(col[:-len(suffix)])
        _stock_sums = {}
        _sold_sums = {}
        for gname in _groups:
            try:
                _s = row.get(f"{gname} 加總 - Shop Stock Qty (CM)")
                _m = row.get(f"{gname} 加總 - Shop Monthly Sold Qty (P90D)")
                _stock_sums[gname] = 0.0 if _s is None or pd.isna(_s) else float(_s)
                _sold_sums[gname] = 0.0 if _m is None or pd.isna(_m) else float(_m)
            except (TypeError, ValueError):
                _stock_sums[gname] = 0.0
                _sold_sums[gname] = 0.0
        _total_stock = sum(_stock_sums.values())
        _total_sold = sum(_sold_sums.values())
        for col in row.index:
            if not (isinstance(col, str) and col.endswith(suffix)):
                continue
            gname = col[:-len(suffix)]
            if hint.strip() and f"存貨偏重 {gname}" in hint:
                targets[col] = "background-color: #FFC000; font-weight: 700"
            if zero_sales and (gname == "零銷" or gname.endswith("-零銷")):
                stock = row.get(f"{gname} 加總 - Shop Stock Qty (CM)")
                if stock is not None and pd.notna(stock):
                    try:
                        if float(stock) > 0:
                            targets[col] = "background-color: #FFC7CE; font-weight: 700; color: #9C0006"
                    except (TypeError, ValueError):
                        pass
            if shortage and col not in targets and _total_sold > 0:
                s_share = (_stock_sums.get(gname, 0.0) / _total_stock) if _total_stock > 0 else 0.0
                v_share = _sold_sums.get(gname, 0.0) / _total_sold
                if v_share >= _short_share and (v_share - s_share) >= _short_gap:
                    targets[col] = "background-color: #BDD7EE; font-weight: 700; color: #1F4E78"
        for i, col in enumerate(row.index):
            if col in targets:
                styles[i] = targets[col]
        return pd.Series(styles, index=row.index)

    return df.style.apply(_row_styles, axis=1)

# 1. 側邊欄設計 (Sidebar)
with st.sidebar:
    st.markdown("""
    <div style="text-align: center; padding: 15px 0 5px 0;">
        <h2 style="color: #0077a3; margin-bottom: 2px; font-weight: 700; font-size: 1.6rem;">I14 Analyzer</h2>
        <p style="color: #5a5a6e; font-size: 0.85rem; margin-top: 0;">自動化報表與數據轉換引擎</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
    
    st.info("📖 各模式的詳細說明與核心計算規則，已移至左側選單的「📖 使用說明」頁。")

    with st.expander("📋 系統操作流程", expanded=False):
        st.markdown("""
        1. 🎯 **選擇模式**：於主頁面頂端選擇分析模式。
        2. 📂 **上傳檔案**：將 Excel 拖曳至主頁面的上傳區域。
        3. 🔧 **設定參數**：按需要調整 MTD 日期、調貨門檻或 Safety 基準天數。
        4. ⚙️ **執行轉換**：點擊「執行報表轉換與分析」。
        5. 💾 **下載成果**：於看板查閱統計資訊，並導出 Excel 報表。
        """)

    with st.expander("🔍 必要欄位檢核標準", expanded=False):
        st.markdown("""
        共用及各模式必要欄位會按目前選擇的分析模式檢核。

        OOS／Safety Stock Review 會使用 `Shop Stock Status` 等模式專用欄位；
        **Safety Stock x NetStock Review** 不要求 `Shop Stock Status`，專用必要欄位詳見使用說明頁；另建議包含 `RP Type`（RF＝可補貨／ND＝不可補貨，ND 列不作調貨需求方；缺失時全部列按 RF 處理並退回 Shop Nature 維度）。
        """, unsafe_allow_html=True)
    
    st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

    if st.button("🧹 釋放記憶體", type="secondary", use_container_width=True,
                 help="清除所有快取資料、st.cache_data 緩存並移除已上傳的 Excel 檔案，將記憶體歸還系統。清除後需重新上傳檔案。"):
        for _key in list(st.session_state.keys()):
            del st.session_state[_key]
        # 清除 st.cache_data 模組級緩存（如 transformer 的欄位計算緩存）
        st.cache_data.clear()
        # 更換 uploader 的 widget key 徹底釋放已上傳檔案 —
        # 否則檔案仍掛在 uploader 上，重跑時會即時重新讀取整個 Excel，記憶體不減反增
        st.session_state.uploader_reset = st.session_state.get("uploader_reset", 0) + 1
        _release_os_memory()
        st.rerun()

    _rss = _process_rss_mb()
    _mem_line = f'Memory Usage: {_rss:.0f} MB<br>' if _rss is not None else ''
    st.markdown(
        '<div style="color: #444455; font-size: 0.75rem; text-align: center; line-height: 1.5;">'
        f'I14 報表計算工具 v{__version__}<br>'
        f'{_mem_line}'
        'Developed by Ricky Yue<br>'
        'Powered by Streamlit & Zeabur'
        '</div>',
        unsafe_allow_html=True
    )

# 2. 主介面標頭 (Header)
st.markdown("""
<div class="header-container">
    <h1 class="header-title">📊 I14 報表轉換與分析系統</h1>
    <p class="header-subtitle">上傳 By SKUxShop 來源檔案，一鍵完成庫存異常篩選、多維度統計，並導出 I14 by SKUxShop_To Buyer 格式報表。</p>
</div>
""", unsafe_allow_html=True)

# 3. 模式選擇器（維持單列，不分組；下方註解說明各模式用途；另加步驟條與分類提示）
st.markdown("""
<div class="steps-bar">
    <div class="step"><span class="n">1</span><b>選擇模式</b><small>下方向常用模式開始</small></div>
    <div class="step"><span class="n">2</span><b>上傳檔案</b><small>.xlsx，上限 200MB</small></div>
    <div class="step"><span class="n">3</span><b>設定參數</b><small>只顯示與模式相關選項</small></div>
    <div class="step"><span class="n">4</span><b>執行並下載</b><small>看板預覽＋一鍵導出</small></div>
</div>
""", unsafe_allow_html=True)
mode_options = ["Overstock Review", "OOS Review", "Carry SKU Review", "SKU 精簡 Review", "NDRF 月份對比", "Safety Stock Review", "Stock Impact Analysis", "Stock Impact Analysis (簡化)", "Past 90 Days 0 Sales", "Stock Checking", "Safety Stock x NetStock Review", "Forecast Review", "銷售損失量化 Review", "積壓資金 Review", "RP Type 覆核 Review", "D001 健康檢查 Review", "新品追蹤 Review"]
st.caption("🔄 調貨與補貨：Stock Checking · Safety Stock x NetStock Review　｜　🔒 安全庫存：Safety Stock Review · Forecast Review　｜　📦 庫存健康：Overstock · OOS · Stock Impact　｜　✂️ 商品與對比：Carry SKU · SKU 精簡 · NDRF 月份對比 · 90 日零銷售")
MODE_DESCRIPTIONS = {
    "Overstock Review": "篩選庫存過剩品項，計算建議安全庫存與可轉出貨量。",
    "OOS Review": "篩選缺貨與潛在缺貨品項，偵測安全庫存不足。",
    "Carry SKU Review": "統計各店舖實際經營品項數量與結構。",
    "SKU 精簡 Review": "針對不可補貨品項，建議退倉安排。",
    "NDRF 月份對比": "對比兩個月份各店舖經營與不可補貨品項變化。",
    "Safety Stock Review": "全面審查安全庫存設定，計算平衡建議值。",
    "Stock Impact Analysis": "分析缺貨對銷售的影響與調貨機會。",
    "Stock Impact Analysis (簡化)": "快速檢視各品項三期存貨與銷售趨勢。",
    "Past 90 Days 0 Sales": "篩選過去 90 天零銷售但仍有存貨的記錄。",
    "Stock Checking": "檢查遊客區缺貨，並由其他區域調貨補足。",
    "Safety Stock x NetStock Review": "以調貨事件審查安全庫存設定是否合理。",
    "Forecast Review": "按品項計算建議訂貨量，並檢視店舖覆蓋日數。",
    "銷售損失量化 Review": "彙總缺貨造成的銷售損失金額，排列補貨優先次序。",
    "積壓資金 Review": "彙總佔用資金最多的滯銷存貨，排列去庫存優先次序。",
    "RP Type 覆核 Review": "檢查可補貨與不可補貨設定與實際銷量是否相符。",
    "D001 健康檢查 Review": "檢查倉庫備貨與店舖需求是否脫節。",
    "新品追蹤 Review": "追蹤上市 90 天內新品的鋪貨與起量情況。",
}
current_mode = st.radio("選擇分析模式", mode_options, horizontal=True, index=0, key="mode_selector")
st.caption(f"用途：{MODE_DESCRIPTIONS.get(current_mode, '')}")

# 3a. 分類維度選項（Safety Stock x NetStock Review／Stock Checking／Safety Stock Review 共用）
_DIMENSION_OPTIONS = {
    "🏬 以 Shop Nature 分類（預設）": "nature",
    "🏷️ 以 RP Type 分類": "rp",
    "🔀 以 Shop Nature＋RP Type 分類": "both",
    "🗺️ 以 OM 分類": "om",
    "🧩 以 Shop Nature＋OM 分類": "nature_om",
    "📈 以銷量分類": "sales",
    "🧮 以銷量＋Shop Nature分類": "sales_nature",
}
_DIMENSION_HELP_COMMON = (
    "決定 Summary／交叉表的分組欄位；非法值自動退回 Shop Nature，不影響任何建議與調貨計算。\n"
    "以 Shop Nature 分類：按 Tourist／Mix／Non Tourist 三組呈現。\n"
    "以 RP Type 分類：按 RF（可補貨）／ND（不可補貨）兩組呈現。\n"
    "以 Shop Nature＋RP Type 分類：按 6 組（如 Tourist-RF、Mix-ND）呈現。\n"
    "以 OM 分類：按來源 OM 欄的不重複值分組呈現（空白自成「空白」組）。\n"
    "以 Shop Nature＋OM 分類：按 Nature×OM 組合（如 Tourist-XXX）呈現。\n"
    "以銷量分類：按銷量帶分組——同 SKU 內按 Sold Qty (P90D) 排名，排名超過 75%＝高銷、"
    "排名在最低 25% 內＝低銷、其餘＝中銷、P90D≤0＝零銷（有銷店舖少於 4 間時不區分高低，全歸中銷）。\n"
    "以銷量＋Shop Nature分類：按 Nature×銷量帶 12 組（如 Tourist-高銷、Mix-低銷）呈現；"
    "銷量帶準則與「以銷量分類」相同，分帶結果隨本次上傳資料的篩選範圍而變。\n"
    "來源缺「OM」欄或全欄空白時，OM 兩個維度退回 Shop Nature 維度。"
)


def _sales_threshold_inputs(prefix: str):
    """銷量存貨分類三條門檻輸入（多模式共用；回傳 缺貨線／積壓線／保底件數）。"""
    st.markdown("**📈 銷量存貨分類門檻**（店舖周轉日數 = Shop Stock Qty (CM) ÷ 月銷 × 30.5）")
    _cols = st.columns(3)
    with _cols[0]:
        _oos = st.number_input(
            "📉 高銷缺貨線 (日)",
            min_value=_HYBRID_SALES_OOS_DAYS_MIN, max_value=_HYBRID_SALES_OOS_DAYS_MAX,
            value=_HYBRID_SALES_OOS_DAYS_DEFAULT, step=1, key=f"{prefix}_sales_oos",
            help="店舖周轉日數低於此值且銷量帶＝高銷 → 「高銷缺貨」。預設 7 日。",
        )
    with _cols[1]:
        _over = st.number_input(
            "📦 低銷積壓線 (日)",
            min_value=_HYBRID_SALES_OVERSTOCK_DAYS_MIN, max_value=_HYBRID_SALES_OVERSTOCK_DAYS_MAX,
            value=_HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT, step=1, key=f"{prefix}_sales_over",
            help="銷量帶＝低銷且店舖周轉日數不少於此值且庫存達保底件數 → 「低銷積壓」（零銷不計周轉，只看保底件數）。預設 30 日。",
        )
    with _cols[2]:
        _minu = st.number_input(
            "🧱 積壓保底件數 (件)",
            min_value=_HYBRID_SALES_OVERSTOCK_MIN_UNITS_MIN, max_value=_HYBRID_SALES_OVERSTOCK_MIN_UNITS_MAX,
            value=_HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT, step=1, key=f"{prefix}_sales_minu",
            help="低銷／零銷積壓分類要求庫存不少於此件數，避免尾數庫存誤報積壓。預設 2 件。",
        )
    return int(_oos), int(_over), int(_minu)

# 4. 初始化 Session State (確保下載與重整時狀態持久化)
MODE_KEYS = {
    "Overstock Review": {"prefix": "", "processed": False, "summary_df": None, "raw_df": None, "categories": None, "excel_data": None, "warnings": []},
    "OOS Review": {"prefix": "oos_", "processed": False, "summary_df": None, "raw_df": None, "categories": None, "excel_data": None, "warnings": []},
    "Carry SKU Review": {"prefix": "carry_", "processed": False, "summary_left": None, "summary_right": None, "dashboard_df": None, "raw_df": None, "excel_data": None, "warnings": []},
    "SKU 精簡 Review": {"prefix": "nd_", "processed": False, "summary_df": None, "raw_df": None, "categories": None, "excel_data": None, "warnings": []},
    "NDRF 月份對比": {"prefix": "ndrf_comp_", "processed": False, "comparison_by_shop": None, "adj_sst_pivot": None, "mch2_pivot": None, "raw_df": None, "excel_data": None, "warnings": []},
    "Safety Stock Review": {"prefix": "combined_", "processed": False, "summary_df": None, "raw_df": None, "categories": None, "excel_data": None, "warnings": [], "dimension": "nature", "sales_oos_days": _HYBRID_SALES_OOS_DAYS_DEFAULT, "sales_overstock_days": _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT, "sales_overstock_min_units": _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT},
	    "Stock Impact Analysis": {"prefix": "stock_impact_", "processed": False, "summary_brand": None, "summary_buyer": None, "summary_category": None, "sku_diagnosis": None, "inv_distribution": None, "raw_slim": None, "excel_data": None, "warnings": []},
    "Stock Impact Analysis (簡化)": {"prefix": "stock_impact_simp_", "processed": False, "raw_df": None, "sku_monthly_analysis": None, "sku_soh_summary": None, "excel_data": None, "warnings": []},
    "Past 90 Days 0 Sales": {"prefix": "p90ds_", "processed": False, "summary_df": None, "raw_df": None, "excel_data": None, "warnings": []},
    "Stock Checking": {"prefix": "stock_chk_", "processed": False, "summary_df": None, "raw_df": None, "stats": None, "excel_data": None, "warnings": []},
    "Safety Stock x NetStock Review": {"prefix": "hybrid_", "processed": False, "raw_df": None, "sc_summary": None, "diagnosis": None, "diagnosis_om": None, "diag_summary": None, "stats": None, "excel_data": None, "params": None, "warnings": []},
    "Forecast Review": {"prefix": "fc_", "processed": False, "main_df": None, "coverage_df": None, "buyer_df": None, "raw_df": None, "stats": None, "excel_data": None, "warnings": []},
    "銷售損失量化 Review": {"prefix": "loss_", "processed": False, "raw_df": None, "by_sku": None, "by_shop": None, "excel_data": None, "warnings": []},
    "積壓資金 Review": {"prefix": "exc_", "processed": False, "raw_df": None, "by_sku": None, "by_shop": None, "excel_data": None, "warnings": []},
    "RP Type 覆核 Review": {"prefix": "rp_", "processed": False, "raw_df": None, "by_sku": None, "by_shop": None, "excel_data": None, "warnings": []},
    "D001 健康檢查 Review": {"prefix": "d001_", "processed": False, "check_df": None, "by_buyer": None, "raw_df": None, "excel_data": None, "warnings": []},
    "新品追蹤 Review": {"prefix": "ns_", "processed": False, "raw_df": None, "by_sku": None, "by_shop": None, "excel_data": None, "warnings": [], "ref_date": None},
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

# 5. 模式切換狀態清理
prev_mode = st.session_state.get("mode")
if prev_mode is not None and prev_mode != current_mode:
    for cfg in MODE_KEYS.values():
        for attr, default in cfg.items():
            if attr == "prefix":
                continue
            key = f"{cfg['prefix']}{attr}"
            st.session_state[key] = [] if isinstance(default, list) else default
    st.session_state.pop("last_uploaded_key", None)
    gc.collect()
st.session_state.mode = current_mode
mode = current_mode

# 6. 檔案上傳區域（key 隨 uploader_reset 變化，「釋放記憶體」按鈕藉此徹底清除已上傳檔案）
uploaded_file = st.file_uploader(
    "📂 上傳 By SKUxShop 原始 Excel 檔案 (.xlsx)",
    type=["xlsx"],
    key=f"src_uploader_{st.session_state.get('uploader_reset', 0)}",
    help="支援標準 Excel (.xlsx) 格式，檔案上限 200MB，最佳化支援 5 萬行以上大數據處理"
)

# 7. 處理上傳的檔案
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
            _t_read0 = time.time()
            df_input = None
            # P0：優先向量化讀取（pandas＋openpyxl），較逐列 iter_rows 快數倍
            try:
                _prog.progress(0.3, text="讀取檔案中... 30% | 向量化解析中...")
                df_input = pd.read_excel(BytesIO(_raw), engine="openpyxl")
            except Exception:
                df_input = None
            if df_input is None:
                # 後備：逐列讀取（相容特殊格式）
                _wb = load_workbook(BytesIO(_raw), data_only=True, read_only=True)
                _ws = _wb.active
                _hdr = list(next(_ws.iter_rows(min_row=1, max_row=1, values_only=True)))
                _rows = []
                _i = 0
                for _i, _r in enumerate(_ws.iter_rows(min_row=2, values_only=True), start=1):
                    _rows.append(list(_r))
                    if _i % 2000 == 0:
                        _elapsed = time.time() - _t0
                        _prog.progress(min(_i / (_i + 2000), 0.99), text=f"讀取檔案中... {_i} 行 | 已用: {_elapsed:.0f}s")
                if _i == 0:
                    _prog.empty()
                    _wb.close()
                    del _raw, _wb, _rows, _hdr
                    st.error("❌ 上傳的 Excel 檔案沒有資料列（僅含標題或完全空白），請確認檔案內容後再試。")
                    st.stop()
                df_input = pd.DataFrame(_rows, columns=_hdr)
                _wb.close()
                del _wb, _rows, _hdr
            if df_input is None or df_input.empty:
                _prog.empty()
                del _raw
                st.error("❌ 上傳的 Excel 檔案沒有資料列（僅含標題或完全空白），請確認檔案內容後再試。")
                st.stop()
            _prog.progress(0.8, text=f"讀取檔案中... 80% | 瘦身優化中（{len(df_input):,} 行）...")
            # P0：上傳後一次瘦身（去全空 Unnamed 欄＋數值降精度），不改欄名與數值
            try:
                df_input = _optimize_input_dtypes(df_input)
            except Exception:
                pass
            del _raw
            gc.collect()
            _read_s = time.time() - _t_read0
            st.session_state.df_read_seconds = round(_read_s, 2)
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

    # P0：全域匯出選項（各模式共用，略過 RAW 明細可省三至七成匯出時間）
    _skip_raw = st.checkbox(
        "⚡ 略過 RAW 明細工作表（匯出較快、檔案較小）",
        value=False,
        key="perf_skip_raw",
        help="勾選後 Excel 仍保留 RAW 工作表名稱，但只寫入一行說明，不寫入逐列明細。摘要與主表不受影響。",
    )
    _include_raw = not _skip_raw

    if mode == "Safety Stock Review":
        mtd_date = st.date_input(
            "📅 MTD 日期 (Sold Qty CM-MTD 的截止日期)",
            value=datetime.now().date() - timedelta(days=1),
            key="ss_mtd_date",
            help="選擇 MTD (Month-To-Date) 截止日期，系統將據此推算月度預估銷售量。"
        )
        mtd_days = mtd_date.day
        month_total_days = calendar.monthrange(mtd_date.year, mtd_date.month)[1]
        _w_pct = round(mtd_days / month_total_days * 100)
        st.caption(
            f"MTD: {mtd_days}日，當月共 {month_total_days} 日。"
            f"月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均（P90D 權重 {100 - _w_pct}%／MTD 權重 {_w_pct}%，"
            "月初以穩定的 P90D 為主，隨月份推進收斂至 MTD）"
        )

        st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

        SHOP_NATURE_OPTIONS = ["Tourist", "Mix", "Non Tourist"]
        selected_shop_natures = st.multiselect(
            "🏷️ 選擇要處理的 Shop Nature 類型",
            options=SHOP_NATURE_OPTIONS,
            default=SHOP_NATURE_OPTIONS,
            key="ss_natures",
            help="選擇要處理的店舖性質，預設全選。只處理部分類型時（如只選 Tourist），其他類型的列會從報表剔除。注意：Non Tourist 選項同時涵蓋 Shop Nature 為「Other」的店（與倍率邏輯一致，非 Tourist/Mix 一律按 Non Tourist 的 Rate 計算）；全部取消選取會導致無資料可處理。"
        )

        with st.expander("⚙️ 自定義 Shop Nature Rate (進階設定)", expanded=False):
            st.info(
                "**倍率如何影響建議值**（建議 Safety Stock = max(OOS 底線 Y, 現有 Safety 若 ≤ Overstock 上限 X，否則 X)，計算結果最低 3 件——最低 3 件為保障自動補貨的保底，受保底影響的建議會在「店舖問題 SKU」建議方法加註；建議值周轉日數見 RAW 的「Recommend Safety Stock Turnday」，店舖問題建議方法亦會顯示）\n\n"
                "- **Overstock Rate（庫存上限 X 的倍率）**：X = floor(P90D 月均銷售 × Rate)。Rate 30% ≈ 保留 9 日銷售量（30% × 30 日）；數值越低，建議上限越低，越積極壓縮庫存\n"
                "- **OOS Rate（缺貨底線 Y 的倍率）**：Y = ceil(月銷預估 × Rate)。Rate 50% ≈ 覆蓋 15 日銷售；數值越高，缺貨保護越強\n\n"
                "**如何選擇合適數值**\n\n"
                "- 想整體減少庫存 → 調低 **Overstock Rate**（壓低上限，需配合「可加可減」模式才有減量效果）\n"
                "- 想降低缺貨風險 → 調高 **OOS Rate**（抬高底線，「只可增加」模式下同樣生效）\n"
                "- 補貨週期長的店（如 Non Tourist）需要較高緩衝 → Overstock Rate 較高；補貨頻密的店（如 Tourist）可較低\n"
                "- ⚠️ 「只可增加」模式下建議值不會低於現有 Safety，Overstock Rate 實際上不產生壓縮效果；要壓縮庫存請選「可加可減」"
            )
            st.markdown("**Overstock Rate** — 各 Shop Nature 庫存上限計算倍率（月銷 × Rate = 建議值上限）。")
            col_os1, col_os2, col_os3 = st.columns(3)
            with col_os1:
                os_rate_tourist = st.slider("Overstock Tourist Rate (%)", 0, 100, 30, 1, key="os_tourist_cb",
                    help="Tourist 類型 Overstock 上限 X = floor(P90D 月均銷售 × Rate)。30% ≈ 保留 9 日銷售量。補貨頻密、週轉快的店可設較低（更積極壓縮庫存）")
            with col_os2:
                os_rate_mix = st.slider("Overstock Mix Rate (%)", 0, 100, 40, 1, key="os_mix_cb",
                    help="Mix 類型 Overstock 上限 X = floor(P90D 月均銷售 × Rate)。40% ≈ 保留 12 日銷售量。")
            with col_os3:
                os_rate_non_tourist = st.slider("Overstock Non Tourist Rate (%)", 0, 100, 50, 1, key="os_non_tourist_cb",
                    help="Non Tourist（及其他類型）Overstock 上限 X = floor(P90D 月均銷售 × Rate)。50% ≈ 保留 15 日銷售量。補貨週期長的店需要較高緩衝")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
            st.markdown("**OOS Rate** — 各 Shop Nature OOS 底線計算倍率（月銷預估 × Rate = 建議值底線）。")
            col_oos1, col_oos2, col_oos3 = st.columns(3)
            with col_oos1:
                oos_rate_tourist = st.slider("OOS Tourist Rate (%)", 0, 100, 50, 1, key="oos_tourist_cb",
                    help="Tourist 類型 OOS 底線 Y = ceil(月銷預估 × Rate)。50% ≈ 覆蓋半個月銷售。缺貨影響大的店可調高（保護更強）")
            with col_oos2:
                oos_rate_mix = st.slider("OOS Mix Rate (%)", 0, 100, 50, 1, key="oos_mix_cb",
                    help="Mix 類型 OOS 底線 Y = ceil(月銷預估 × Rate)。50% ≈ 覆蓋 15 日銷售。")
            with col_oos3:
                oos_rate_non_tourist = st.slider("OOS Non Tourist Rate (%)", 0, 100, 50, 1, key="oos_non_tourist_cb",
                    help="Non Tourist（及其他類型）OOS 底線 Y = ceil(月銷預估 × Rate)。月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均（月初以 P90D 為主）")

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
            key="ss_safety_mode",
            help="只可增加：Safety Qty 只增不減（ND / P90D 0 Sales 亦不歸零，全表無扣減）；可加可減：允許減少，並套用退場建議（ND → 0、P90D 0 Sales → 0+ND，Launch Date 90 天內新品除外）"
        )
        safety_mode_value = "increase_only" if safety_mode == "只可增加" else "allow_both"

        _ss_dim_label = st.radio(
            "🗂️ 分類維度",
            options=list(_DIMENSION_OPTIONS),
            index=0,
            horizontal=True,
            key="ss_dim",
            help="決定 SUMMARY「③ Safety Qty Change 分類 × 分類維度交叉表」的分組欄位。" + _DIMENSION_HELP_COMMON +
                 "\n本模式所有維度均只改分組顯示，Safety 建議計算（倍率、MTD 月銷預估、退場建議）不隨維度改變。"
                 "RAW／店舖問題 SKU 另恆附「銷量帶」與「銷量存貨分類」顯示欄（見下方門檻）。"
        )
        _ss_dimension = _DIMENSION_OPTIONS[_ss_dim_label]
        _ss_sales_oos, _ss_sales_over, _ss_sales_minu = _sales_threshold_inputs("ss")

        _all_site_codes = canonical_site_codes(df_input["Site Code"]) if "Site Code" in df_input.columns else []

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
                    ref_date=mtd_date,
                    sales_oos_days=_ss_sales_oos,
                    sales_overstock_days=_ss_sales_over,
                    sales_overstock_min_units=_ss_sales_minu,
                )
                buf = generate_excel_combined(
                    summary_df, raw_df, categories, progress_callback=_cb_excel,
                    dimension=_ss_dimension,
                    stats={
                        "safety_mode": safety_mode_value,
                        "rates_overstock": rates_overstock,
                        "rates_oos": rates_oos,
                        "mtd_date": str(mtd_date),
                        "mtd_days": mtd_days,
                        "month_total_days": month_total_days,
                        "ref_date": str(mtd_date),
                        "sales_oos_days": _ss_sales_oos,
                        "sales_overstock_days": _ss_sales_over,
                        "sales_overstock_min_units": _ss_sales_minu,
                    },
                    include_raw=_include_raw,
                )
                st.session_state.combined_processed = True
                st.session_state.combined_summary_df = summary_df
                st.session_state.combined_raw_df = raw_df
                st.session_state.combined_categories = categories
                st.session_state.combined_excel_data = buf
                st.session_state.combined_warnings = _warnings
                st.session_state.combined_dimension = _ss_dimension
                st.session_state.combined_sales_oos_days = _ss_sales_oos
                st.session_state.combined_sales_overstock_days = _ss_sales_over
                st.session_state.combined_sales_overstock_min_units = _ss_sales_minu
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

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📈 Safety Stock 數據分析摘要 (Safety Stock Analysis Dashboard)")
            _top_download_button(excel_data, "I14 Safety Stock Review", "dl_top_combined")

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

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            selected_site_code = st.selectbox(
                "🏪 選擇 Site Code 檢視明細",
                options=["全部"] + _all_site_codes,
                index=0,
                key="ss_site",
                help="選擇特定 Site Code 過濾下方表格。僅影響 RAW 明細預覽。"
            )

            tab1, tab_cross, tab2, tab3, tab4 = st.tabs(["📋 統計彙總表 (Shop Stock Status)", "🧮 分類交叉表", "🏪 By Shop 明細", "🏪 店舖問題 SKU", "📄 RAW 原始明細"])

            with tab1:
                st.markdown("#### 採購人員與品牌之 Safety Stock 交叉統計表")
                st.markdown("本表即為 Excel 導出檔中「Shop Stock Status」工作表之即時數據：")
                st.dataframe(summary_df, width='stretch', hide_index=True)

            with tab_cross:
                _ss_applied_dim = st.session_state.get("combined_dimension") or "nature"
                _ss_dim_txt = _HYBRID_DIMENSION_LABELS.get(_ss_applied_dim, "Shop Nature")
                _ss_oos_txt = int(st.session_state.get("combined_sales_oos_days") or _HYBRID_SALES_OOS_DAYS_DEFAULT)
                _ss_over_txt = int(st.session_state.get("combined_sales_overstock_days") or _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT)
                _ss_minu_txt = int(st.session_state.get("combined_sales_overstock_min_units") or _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT)
                st.markdown(f"#### ① Safety Qty Change 分類 × {_ss_dim_txt} 交叉表（列數 / SKU 數 / Safety Qty Change 合計）")
                st.markdown("即為 Excel 導出檔 SUMMARY 工作表區塊 ③ 之即時數據（只改分組顯示，Safety 建議計算不變）：")
                try:
                    _ss_cross = _build_combined_dimension_cross_tab(raw_df, dimension=_ss_applied_dim)
                except Exception:
                    _ss_cross = None
                if _ss_cross is not None and not _ss_cross.empty:
                    st.dataframe(_ss_cross, width='stretch', hide_index=True)
                    st.caption("增加／減少／不變按「Safety Qty Change」正負分行；非標準 Shop Nature 值（如 Other／空白）不佔分組欄、只計入合計；"
                               "總計列 SKU 數為全表不重複 SKU（同一 SKU 可跨分類）。")
                else:
                    st.info("無符合資料")

                st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
                st.markdown("#### ② 銷量 × 存貨分類分佈（母體為全部店舖×SKU 列）")
                try:
                    _ss_sales_block = _build_hybrid_sales_summary(raw_df)
                except Exception:
                    _ss_sales_block = None
                if _ss_sales_block is not None and not _ss_sales_block.empty:
                    st.dataframe(_ss_sales_block, width='stretch', hide_index=True)
                    st.caption(f"分類準則：高銷缺貨＝同 SKU 內屬高銷帶且店舖周轉日數 < {_ss_oos_txt} 日；"
                               f"低銷積壓＝低銷帶且周轉日數 ≥ {_ss_over_txt} 日且庫存 ≥ {_ss_minu_txt} 件；"
                               f"零銷積壓＝P90D 零銷售且庫存 ≥ {_ss_minu_txt} 件；其餘為正常。"
                               "銷量帶按本次上傳資料內同 SKU 排名計算，篩選範圍不同時結果會不同。")
                else:
                    st.info("無符合資料（來源檔缺少「Sold Qty (P90D)」或「Shop Stock Qty (CM)」時不產生銷量存貨分類）")

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
                st.markdown("#### 每間店舖需行動 SKU（問題＋建議方法）")
                st.markdown("只收錄需行動列（Safety 需增加／減少、ND 退場），按店舖分組、店內按嚴重度排序；建議方法附建議值周轉日數，受 OOS 底線最低 3 件保底影響時另加註保底原因：")
                try:
                    _ss_issues = _build_combined_shop_issues(raw_df)
                except Exception:
                    _ss_issues = None
                if _ss_issues is not None and not _ss_issues.empty:
                    _ss_cols = [c for c in _COMBINED_SHOP_ISSUE_COLUMNS if c in _ss_issues.columns]
                    st.dataframe(_ss_issues[_ss_cols].rename(columns=_display_safety_header), width='stretch', hide_index=True)
                    st.caption(f"共 {len(_ss_issues):,} 列需行動（只收需行動列，正常列不佔版面）。")
                else:
                    st.info("無符合資料")

            with tab4:
                st.markdown("#### RAW 原始資料庫存明細")
                _filtered_raw = raw_df.copy()
                if selected_site_code != "全部" and "Site Code" in _filtered_raw.columns:
                    _filtered_raw = filter_by_site_code(_filtered_raw, selected_site_code)
                # 隱藏內部欄（如保底偵測旗標），只顯示使用者可見欄位
                _filtered_raw = _filtered_raw[[c for c in _filtered_raw.columns if not str(c).startswith("_")]]
                st.dataframe(_filtered_raw.head(50).rename(columns=_display_safety_header), width='stretch')
                st.caption(f"僅顯示前 50 筆（{'全部' if selected_site_code == '全部' else selected_site_code}），完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

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
    elif mode == "Safety Stock x NetStock Review":
        with st.expander("🔀 調貨事件 Review Safety 設定說明（點選展開）", expanded=False):
            st.caption("🔀 **調貨事件 Review Safety 設定**：沿用 Stock Checking 調貨規則（14 日目標、NT 優先→Mix 補足）後，以 **Safety 基準周轉天數**（按 Shop Nature 分別設定，預設 Tourist 7 / Mix 9 / Non Tourist 10 日，可調 7–10）Review Safety Stock 設定 — 需調入列若現有設定低於所需量 → 「Safety 設定不足」；有貨可轉出列若現有設定高於生效基準 → 「Safety 設定過高」，並給出建議調整值與調整量（**下調建議補底 3 件**——受限 SAP MOQ，Safety 1–2 件難以補貨：現有 ≥ 3 件而下調至 1–2 件者建議最低 3 件；現有本身 1–2 件不強制上調；零銷售歸零建議不受保底；上調建議照常）。另附 **全列預防性 Safety 設定建議**（母體為全部店舖×SKU 列，不依賴調貨事件；ND 列不作上調）：RAW「全列」五欄與「🛡️ Safety 設定建議 (全列)」分頁供一次過設定適合的 Safety Stock，減少日後調貨次數。主表另附「建議 Safety 設定值 Turnday」（建議值 ÷ (P90D 月均 ÷ 30)，月均 0 留空），受保底影響的建議在店舖問題 SKU 建議方法加註保底原因。Safety 基準與調貨的 14 日線**互相獨立**（調整 Safety 天數不會改變任何調貨數字）。調貨結果分 **不限OM** 與 **同OM**（來源店須與需求店同 OM）兩套對照；OM 限制只影響調貨分配，不影響 Safety Review 與建議調整。調貨段除 Stock Checking 規則外，另含本模式的資料品質補正（Shop Nature 正規化、月銷空白／為 0 但有 P90D 銷售時逐列回補、D001 逐列 Net→Stock 回補），同檔與 Stock Checking 模式的結果可能有細微差異。可勾選 **允許雙向調貨**：存貨偏重遊客區時，Mix / Non Tourist 缺貨店舖亦可調入、遊客區店舖可轉出。可選擇 **分類維度**（預設 Shop Nature）：改以 RP Type（RF＝可補貨／ND＝不可補貨）、Shop Nature＋RP Type、OM、Shop Nature＋OM、銷量（高銷／中銷／低銷／零銷 四帶）或 銷量＋Shop Nature（Nature×銷量帶 12 組）分類；其中 OM／銷量／銷量＋Shop Nature 維度計算與預設相同，僅 Summary／交叉表改按 OM、Nature×OM、銷量帶或 Nature×銷量帶分組（TurnDay 與 Safety 基準天數恆按 Shop Nature）；**RP Type 為 ND 的店舖列一律不作調貨需求方（仍可作供應方轉出）**。來源缺「OM」欄或全欄空白時，OM 兩個維度退回 Shop Nature。可勾選 **ND 優先轉出**（預設關閉）：同一 Shop Nature 池內 **RP Type＝ND 的店舖先於 RF 店舖轉出**（池間順序仍為 Non Tourist → Mix → Tourist，轉出資格一律不放寬），只改變逐列轉出量與調貨來源，**不影響 Safety Review 與建議調整量**。RAW／主表恆附 **銷量帶**（同 SKU 內按 Sold Qty (P90D) 排名分帶）與 **銷量存貨分類**（高銷缺貨／低銷積壓／零銷積壓／正常；店舖周轉日數門檻預設缺貨 7 日、積壓 30 日、保底 2 件，與調貨 14 日線及 Safety 基準天數互相獨立）；「Review 摘要」另附 **⑤ 銷量 × 存貨分類交叉表**（母體為全部店舖×SKU 列）與 **⑥ 全列 Safety 調整方向分類／明細**。可勾選 **啟用銷量再平衡**（預設關閉）：高銷缺貨的 RF 店舖列即使所屬 Nature 池不缺貨亦可加入調貨需求方（仍須 D001 低於門檻、補到 14 日需求 > 0 且同 SKU 有貨可調），積壓列經既有轉出判定提供來源；跟隨單向／雙向框架，自我調貨／同 Nature／同OM 守衛與池級歸零不放寬。另附 SKU 級 **Summary**（按所選維度分組的庫存/銷量佔比與「存貨偏重提示」，偏重數值以醒目橘色底色加強顯示，存貨偏低分組另以淺藍底標示庫存佔比欄；以銷量／銷量＋Shop Nature 分類時，零銷組「加總 - Shop Stock Qty (CM) %」大於 0% 另以醒目紅底標示——零銷售而仍有存貨，理論上全屬 Overstock）。")

        sc_top = st.columns(2)
        with sc_top[0]:
            turnday_threshold = st.number_input(
                "🔁 Tourist 加總 TurnDay 門檻 (日)",
                min_value=1, max_value=365, value=14, step=1, key="hy_ttd",
                help="該 SKU 的 Tourist 加總 TurnDay（庫存加總 ÷ 月銷加總 × 30.5）低於此值時才列入調貨候選；"
                     "勾選雙向調貨後，此門檻同樣套用於各 Nature 的加總 TurnDay 判定。Safety Review 基準另行設定（見下方）。"
            )
        with sc_top[1]:
            d001_threshold = st.number_input(
                "🏭 D001 庫存門檻 (件)",
                min_value=0, max_value=1000000, value=100, step=10, key="hy_d001",
                help="當 D001 Net Stock Qty (CM) 低於此值且需求 Nature 加總 TurnDay 低於門檻時，才標示需要調貨。"
            )
        _hybrid_dim_options = _DIMENSION_OPTIONS
        _hybrid_dim_label = st.radio(
            "🗂️ 分類維度",
            options=list(_hybrid_dim_options),
            index=0,
            horizontal=True,
            key="hy_dim",
            help="決定調貨需求方過濾與 Summary／Review 摘要交叉表的分組欄位；"
                 "TurnDay 池與 Safety 基準天數恆按 Shop Nature（Tourist／Mix／Non Tourist）設定，不隨維度改變。\n"
                 "以 Shop Nature 分類：需求方＝Tourist 且 RP Type＝RF 的店舖；供應方＝Mix／Non Tourist（有貨即可轉出，不論 RF／ND）。\n"
                 "以 RP Type 分類：需求方＝任何 Nature 且 RF 的店舖（Mix／Non Tourist 缺貨亦可調入，含同 Nature 店間再平衡）；"
                 "供應方維持 Mix／Non Tourist。\n"
                 "以 Shop Nature＋RP Type 分類：計算與預設相同，Summary／交叉表按 6 組（如 Tourist-RF、Mix-ND）呈現。\n"
                 "以 OM 分類：計算與預設相同，Summary／交叉表按來源 OM 欄的不重複值分組呈現。\n"
                 "以 Shop Nature＋OM 分類：計算與預設相同，Summary／交叉表按 Nature×OM 組合（如 Tourist-XXX）呈現。\n"
                 "以銷量分類：計算與預設相同，Summary／交叉表按銷量帶分組——同 SKU 內按 Sold Qty (P90D) 排名，"
                 "排名超過 75%＝高銷、排名在最低 25% 內＝低銷、其餘＝中銷、P90D≤0＝零銷（有銷店舖少於 4 間時不區分高低，全歸中銷）。\n"
                 "以銷量＋Shop Nature分類：計算與預設相同，Summary／交叉表按 Nature×銷量帶 12 組（如 Tourist-高銷、Mix-低銷）呈現；"
                 "銷量帶準則與「以銷量分類」相同，分帶結果隨本次上傳資料的篩選範圍而變。\n"
                 "RP Type 為 ND（不可補貨）的店舖列一律不作需求方（仍可作供應方）；其他／空白值歸入 ND 並列 warning；"
                 "來源缺「RP Type」欄時全部列按 RF 處理並退回 Shop Nature 維度（OM／銷量維度不退回）。"
                 "來源缺「OM」欄或全欄空白時，OM 兩個維度退回 Shop Nature 維度。"
        )
        _dimension = _hybrid_dim_options[_hybrid_dim_label]
        st.markdown("**🛡️ Safety 基準周轉天數（按 Shop Nature 分別設定；與調貨段的 14 日線互不影響）**")
        _sd_cols = st.columns(3)
        _sd_help = ("Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 此天數)。此天數只決定 Safety Stock 的 Review 基準線，"
                    "不影響調貨需求與可轉出貨量（兩者固定補到／扣除 14 日周轉）。")
        with _sd_cols[0]:
            sd_tourist = st.number_input(
                "🧳 Tourist (日)", min_value=_HYBRID_SAFETY_DAYS_MIN, max_value=_HYBRID_SAFETY_DAYS_MAX,
                value=_HYBRID_SAFETY_DAYS_DEFAULT["Tourist"], step=1, key="hy_sd_tourist",
                help=_sd_help + " 預設 7 日。",
            )
        with _sd_cols[1]:
            sd_mix = st.number_input(
                "🔀 Mix (日)", min_value=_HYBRID_SAFETY_DAYS_MIN, max_value=_HYBRID_SAFETY_DAYS_MAX,
                value=_HYBRID_SAFETY_DAYS_DEFAULT["Mix"], step=1, key="hy_sd_mix",
                help=_sd_help + " 預設 9 日。",
            )
        with _sd_cols[2]:
            sd_nt = st.number_input(
                "🏪 Non Tourist (日)", min_value=_HYBRID_SAFETY_DAYS_MIN, max_value=_HYBRID_SAFETY_DAYS_MAX,
                value=_HYBRID_SAFETY_DAYS_DEFAULT["Non Tourist"], step=1, key="hy_sd_nt",
                help=_sd_help + " 預設 10 日。",
            )
        _safety_days = {
            "Tourist": int(sd_tourist), "Mix": int(sd_mix), "Non Tourist": int(sd_nt),
        }
        bidirectional = st.checkbox(
            "🔁 允許雙向調貨（任何 Nature 缺貨皆可調入，遊客區店舖亦可轉出）",
            value=False, key="hy_bidir",
            help="預設關閉：維持現行規則（只向遊客區調入，來源 Mix / Non Tourist）。"
                 "開啟後任何 Nature 加總 TurnDay 低於門檻且其他任一 Nature 有貨即標示需要調貨；"
                 "新增「可轉出貨量 (Tourist)」與 Summary「由Tourist 店舖預計調出數量」欄位，"
                 "來源優先序固定 Non Tourist → Mix → Tourist。"
        )
        nd_priority = st.checkbox(
            "🥇 ND 優先轉出（同一 Shop Nature 池內，RP Type＝ND 的店舖先於 RF 轉出）",
            value=False, key="hy_nd_priority",
            help="預設關閉：池內按來源檔列序取用（現行規則）。"
                 "開啟後「同一 Shop Nature 池內」RP Type＝ND 的來源列先於 RF 列被取用，"
                 "池間順序仍為 Non Tourist → Mix → Tourist（雙向時才含 Tourist），"
                 "展開為 Non Tourist-ND → Non Tourist-RF → Mix-ND → Mix-RF → Tourist-ND → Tourist-RF；"
                 "ND／RF 兩段各自維持來源檔列序（不按可轉出貨量重排）。\n"
                 "轉出資格一律不放寬：仍須該池加總 TurnDay ≥ 門檻、該列有 14 日超額（零銷售有貨列為全數庫存）、"
                 "非自我調貨、非同 Nature（「以 RP Type 分類」維度允許同 Nature 互調）、同OM 變體另須同 OM；"
                 "ND 列仍不作調貨需求方。\n"
                 "本開關只改變調貨分配結果（轉出量／調貨來源／事件類型／調貨後庫存與 TurnDay），"
                 "不影響 Safety Review、建議 Safety 設定值與調整量。"
                  "「ND」的定義為正規化後非 RF（含無法辨識而歸入 ND 的列）；來源檔缺「RP Type」欄時自動失效並列 warning。"
         )
        st.markdown("**📈 銷量分類與銷量再平衡**")
        _sales_hint = (
            "「銷量帶」與「銷量存貨分類」兩個顯示欄位恆輸出（RAW／主表，全列覆蓋），不受本開關影響；"
            "開關只控制是否讓「高銷缺貨」店舖列加入調貨需求方與積壓列經既有來源判定轉出。"
        )
        sales_rebalance = st.checkbox(
            "⚖️ 啟用銷量再平衡（高銷缺貨店舖可調入＋低銷／零銷積壓列可轉出）",
            value=False, key="hy_sales_reb",
            help="預設關閉：調貨需求判定維持現行規則（按 Nature 池加總 TurnDay）。"
                 "開啟後，同 SKU 內銷量屬高銷帶且店舖周轉日數低於缺貨線的店舖×SKU 列，"
                 "即使所屬 Nature 池整體不缺货，也可加入調貨需求方（仍須 RP Type＝RF、"
                 "補到 14 日需求 > 0、D001 < 門檻且同 SKU 有貨可調）；"
                 "積壓列（低銷／零銷且達積壓線與保底件數）經既有轉出判定提供來源。"
                 "單向模式需求方仍限 Tourist 列、轉出池仍限 Mix／Non Tourist（跟隨「允許雙向調貨」框架）；"
                 "自我調貨／同 Nature／同OM 守衛與池級歸零全部保留；"
                 "Safety 分類欄位屬觀察層，與調貨的 14 日線及 Safety 基準天數互相獨立。"
        )
        _sales_cols = st.columns(3)
        with _sales_cols[0]:
            sales_oos_days = st.number_input(
                "📉 高銷缺貨線 (日)",
                min_value=_HYBRID_SALES_OOS_DAYS_MIN, max_value=_HYBRID_SALES_OOS_DAYS_MAX,
                value=_HYBRID_SALES_OOS_DAYS_DEFAULT, step=1, key="hy_sales_oos",
                help="店舖周轉日數（Shop Stock Qty (CM) ÷ 生效月銷 × 30.5）低於此值且銷量帶＝高銷 → 「高銷缺貨」。"
                     + _sales_hint
            )
        with _sales_cols[1]:
            sales_overstock_days = st.number_input(
                "📦 低銷積壓線 (日)",
                min_value=_HYBRID_SALES_OVERSTOCK_DAYS_MIN, max_value=_HYBRID_SALES_OVERSTOCK_DAYS_MAX,
                value=_HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT, step=1, key="hy_sales_over",
                help="銷量帶＝低銷且店舖周轉日數不少於此值且庫存達保底件數 → 「低銷積壓」（零銷不計周轉，只看保底件數）。"
                     + _sales_hint
            )
        with _sales_cols[2]:
            sales_overstock_min_units = st.number_input(
                "🧱 積壓保底件數 (件)",
                min_value=_HYBRID_SALES_OVERSTOCK_MIN_UNITS_MIN, max_value=_HYBRID_SALES_OVERSTOCK_MIN_UNITS_MAX,
                value=_HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT, step=1, key="hy_sales_minu",
                help="低銷／零銷積壓分類要求庫存不少於此件數，避免尾數庫存誤報積壓。"
                     + _sales_hint
            )
        return_to_d001 = st.checkbox(
            "📤 建議退D001（中銷／低銷／零銷退回超出 Safety 的庫存；高銷須庫存高於 Safety 200%）",
            value=False, key="hy_return_d001",
            help="預設關閉：只輸出店舖間調貨建議。"
                 "開啟後新增 RAW 欄「建議退D001 量」與「退D001 建議」工作表："
                 "接收目標為 D001（Warehouse，可無限接收貨量），故不跑分配。"
                 "退貨基準為 Safety Stock (現時)（非調貨段的 14 日線，與池級門檻脫鉤，逐列判定）："
                 "中銷／低銷／零銷列退回超出 Safety 的庫存（庫存 − Safety）；"
                 "高銷列原則上不退，僅當 Safety > 0 且庫存 ≥ 3×Safety（高於 Safety 200%）時退至 2×Safety。"
                 "調貨優先：先扣已轉出量，需要調貨列一律不退。"
                 "本欄只作退倉建議，不影響調貨分配、Safety Review 與建議調整量。"
        )
        hy_ref_date = st.date_input(
            "📅 參考日期（New SKU 判定基準；Launch Date 起 90 天內視為新品）",
            value=datetime.now().date() - timedelta(days=1),
            key="hy_ref_date",
            help="Past 90 Days 0 Sales 轉 ND 候選的新品保護以此日期起計 90 天（Launch Date 在此日期前 90 天內視為新品，不作 ND 提示）。預設為昨日，與報表資料基準一致。",
        )
        _current_hybrid_params = build_hybrid_params(
            turnday_threshold, d001_threshold, _safety_days, bidirectional, _dimension,
            hy_ref_date, nd_priority,
            sales_rebalance, sales_oos_days, sales_overstock_days, sales_overstock_min_units,
            return_to_d001,
        )
        if bidirectional:
            st.caption(
                f"🔁 雙向調貨觸發：該 SKU **任一 Nature 加總 TurnDay < {int(turnday_threshold)} 日** 且 **D001 < {int(d001_threshold)} 件**，"
                "且**其他任一 Nature 加總 TurnDay ≥ 門檻**（有貨可調）→ 該 Nature 缺貨店舖（RP Type＝RF）標示需要調貨；"
                "調入補到 14 日周轉，來源固定 Non Tourist → Mix → Tourist 依序轉出（需求 Nature 自身一律排除——池級歸零與分配層守衛雙重保證；"
                "「以 RP Type 分類」維度例外——允許同 Nature 店間再平衡）。"
            )
        else:
            st.caption(
                f"🔁 調貨觸發（同 Stock Checking）：該 SKU 的 **Tourist 加總 TurnDay < {int(turnday_threshold)} 日** 且 **D001 < {int(d001_threshold)} 件**，"
                "且 Mix / Non Tourist 至少一個 Nature 加總 TurnDay ≥ 門檻（有貨可調）。調入補到 14 日周轉，NT 優先再由 Mix 補足。"
            )
        if nd_priority:
            st.caption(
                "🥇 ND 優先轉出已開啟：來源取用順序為 **Non Tourist-ND → Non Tourist-RF → Mix-ND → Mix-RF**"
                + ("** → Tourist-ND → Tourist-RF**（雙向已開啟，故含 Tourist 池）" if bidirectional else "（未開雙向時不取 Tourist 貨）")
                + "；池間順序與轉出資格（池級加總 TurnDay 門檻、逐列 14 日超額、自我調貨／同 Nature／同OM 守衛）完全不變，"
                "ND 列仍不作調貨需求方。正常一店×一SKU 一列的資料下，各 Nature 轉出合計與調入總量不變"
                "（若同一 Site Code 同時出現在需求與來源端——「Site Code × SKU」重複列異常，系統會列 warning——"
                "自我調貨守衛會令池內取用順序影響合計與缺口），"
                "**變動的是逐列轉出量、調貨來源、事件類型與調貨後庫存／TurnDay**；"
                "Safety Review、建議 Safety 設定值與調整量不受影響。"
            )
        if sales_rebalance:
            st.caption(
                f"⚖️ 銷量再平衡已開啟：同 SKU 內屬**高銷帶**且店舖周轉日數 < 缺貨線 {int(sales_oos_days)} 日 的 RF 列"
                f"（補到 14 日需求 > 0、D001 < {int(d001_threshold)} 件且同 SKU 有貨可調）即使所屬 Nature 池不缺貨也可調入；"
                f"積壓列（低銷且周轉 ≥ {int(sales_overstock_days)} 日，或零銷，且庫存 ≥ {int(sales_overstock_min_units)} 件）"
                "經既有轉出判定提供來源。"
                + ("目前已勾選雙向：兩端放寬至任何 Nature。" if bidirectional
                   else "未勾選雙向：需求方仍限 Tourist 列、轉出池仍限 Mix／Non Tourist。")
                + "自我調貨／同 Nature／同OM 守衛與池級歸零全部保留；ND 列仍不作需求方。"
            )
        if return_to_d001:
            st.caption(
                f"📤 建議退D001 已開啟：接收目標為 **D001（Warehouse，可無限接收貨量）**，故不跑分配。"
                "退貨基準為 **Safety Stock (現時)**（非調貨段的 14 日線，與池級門檻脫鉤，逐列判定）："
                "**中銷／低銷／零銷**列退回超出 Safety 的庫存（庫存 − Safety）；"
                "**高銷**列原則上不退，僅當 Safety > 0 且**庫存 ≥ 3×Safety（高於 Safety 200%）**時退至 2×Safety。"
                "調貨優先：建議退D001 量＝可退D001 貨量 − 轉出量（不限OM），需要調貨列一律不退。"
                "本欄只作退倉建議，不影響調貨分配、Safety Review 與建議調整量；"
                "OM 限制與分類維度均不影響本欄。"
            )
        if _dimension == "rp":
            st.caption(
                f"🏷️ 目前維度＝以 RP Type 分類：需求方＝**任何 Nature 且 RP Type＝RF** 的缺貨店舖"
                f"（該 Nature 加總 TurnDay < {int(turnday_threshold)} 日 且 D001 < {int(d001_threshold)} 件 且補到 14 日需求 > 0 且同 SKU 有貨可調，"
                "允許同 Nature 店間再平衡）；供應方維持 Mix／Non Tourist（ND 店可轉出）；"
                "ND 列一律不作需求方。Summary 與交叉表按 RF／ND 分組。"
            )
        elif _dimension == "both":
            st.caption(
                "🔀 目前維度＝Shop Nature＋RP Type：調貨計算與預設相同（需求 Tourist-RF、供應 Mix／Non Tourist），"
                "Summary 與交叉表按 6 組（Tourist-RF、Tourist-ND、Mix-RF、Mix-ND、Non Tourist-RF、Non Tourist-ND）呈現。"
            )
        elif _dimension == "om":
            st.caption(
                "🗺️ 目前維度＝以 OM 分類：調貨計算與預設相同（需求 Tourist-RF、供應 Mix／Non Tourist），"
                "Summary 與交叉表按來源 OM 欄的不重複值分組呈現；來源缺「OM」欄或全欄空白時退回 Shop Nature 維度。"
            )
        elif _dimension == "nature_om":
            st.caption(
                "🧩 目前維度＝Shop Nature＋OM：調貨計算與預設相同（需求 Tourist-RF、供應 Mix／Non Tourist），"
                "Summary 與交叉表按 Nature×OM 組合（如 Tourist-XXX）呈現；來源缺「OM」欄或全欄空白時退回 Shop Nature 維度。"
            )
        elif _dimension == "sales":
            st.caption(
                "📈 目前維度＝以銷量分類：調貨計算與預設相同（需求 Tourist-RF、供應 Mix／Non Tourist），"
                "Summary 與交叉表按銷量帶（高銷／中銷／低銷／零銷）分組呈現；"
                "銷量帶按本次上傳資料內同 SKU 各店舖 Sold Qty (P90D) 排名計算，篩選範圍不同時分帶結果會不同。"
            )
        elif _dimension == "sales_nature":
            st.caption(
                "🧮 目前維度＝以銷量＋Shop Nature分類：調貨計算與預設相同（需求 Tourist-RF、供應 Mix／Non Tourist），"
                "Summary 與交叉表按 Nature×銷量帶 12 組（Tourist-高銷／中銷／低銷／零銷、Mix-…、Non Tourist-…）呈現；"
                "銷量帶準則與「以銷量分類」相同——按本次上傳資料內同 SKU 各店舖 Sold Qty (P90D) 排名計算，"
                "篩選範圍不同時分帶結果會不同。"
            )
        else:
            st.caption(
                "🏬 目前維度＝以 Shop Nature 分類（預設）：需求方僅 Tourist 且 RP Type＝RF 的店舖"
                "（ND 列不作需求方，仍可作供應方）；Summary 與交叉表按 Tourist／Mix／Non Tourist 分組。"
            )
        st.caption(
            f"📦 存貨偏重提示門檻（隨分類維度切換，目前維度：{_hybrid_skew_hint_text(_dimension)}；"
            "觸發的分組以醒目橘底加強顯示）——完整對照：Shop Nature／RP Type／Nature＋RP Type 為庫存佔比 ≥ 40% 且差距 ≥ 20 個百分點；"
            "OM 為 ≥ 25% 且差距 ≥ 12 個百分點；Nature＋OM／銷量＋Shop Nature 為 ≥ 12% 且差距 ≥ 8 個百分點；銷量 為 ≥ 30% 且差距 ≥ 15 個百分點。"
        )
        st.caption(
            f"🔵 存貨偏低標示門檻（隨分類維度切換，目前維度：{_hybrid_shortage_hint_text(_dimension)}；"
            "觸發的分組以淺藍底標示庫存佔比欄，只標示儲存格）——完整對照：Shop Nature／RP Type／Nature＋RP Type 為銷量佔比 ≥ 30% 且差距 ≥ 10 個百分點；"
            "OM 為 ≥ 20% 且差距 ≥ 8 個百分點；銷量為 ≥ 20% 且差距 ≥ 10 個百分點；Nature＋OM／銷量＋Shop Nature 為 ≥ 10% 且差距 ≥ 5 個百分點。"
        )
        st.caption(
            f"🧭 Review 基準：Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 基準天數)，"
            f"目前基準 = {_hybrid_safety_days_label(_safety_days)}；"
            "調入列現有設定 < 所需量 → 「Safety 設定不足」；"
            "轉出列現有設定 > 生效基準 → 「Safety 設定過高」；建議 Safety 設定值 = 不足列取所需量、"
            "過高列取生效基準（僅不足／過高列給建議）。下調生效基準補底 3 件（SAP MOQ：Safety 1–2 件難補貨）："
            "現有 ≥ 3 件者最低建議 3 件、現有 1–2 件不強制上調、零銷售歸零不受保底；上調建議照常。"
            "主表與 RAW 另附「建議 Safety 設定值 Turnday」欄（建議值 ÷ (P90D 月均 ÷ 30)，月均 0 留空）；"
            "受保底影響的建議會在店舖問題 SKU 的建議方法加註保底原因。"
            "此基準不影響調貨需求與可轉出貨量（固定 14 日）。"
        )

        if st.button("⚙️ 執行報表轉換與分析", type="primary", width='stretch'):
            _prog2 = st.progress(0, text="處理中... 0% | 開始處理...")
            _t0_2 = time.time()

            def _cb(pct, msg):
                _e = time.time() - _t0_2
                _eta2 = (_e / pct) - _e if pct > 0 else 0
                _prog2.progress(min(pct, 1.0), text=f"處理中... {pct*100:.0f}% | {msg} | 已用: {_e:.0f}s | 預計剩餘: {_eta2:.0f}s")

            try:
                def _cb_transform(pct, msg):
                    _cb(0.55 * pct, msg)

                def _cb_excel(pct, msg):
                    _cb(0.55 + 0.45 * pct, msg)

                raw_df, sc_summary, diagnosis, diagnosis_om, diag_summary, _stats, _warnings = transform_hybrid(
                    df_input,
                    turnday_threshold=int(turnday_threshold),
                    d001_threshold=int(d001_threshold),
                    bidirectional=bool(bidirectional),
                    safety_days=_safety_days,
                    dimension=_dimension,
                    nd_priority=bool(nd_priority),
                    progress_callback=_cb_transform,
                    ref_date=hy_ref_date,
                    sales_rebalance=bool(sales_rebalance),
                    sales_oos_days=int(sales_oos_days),
                    sales_overstock_days=int(sales_overstock_days),
                    sales_overstock_min_units=int(sales_overstock_min_units),
                    return_to_d001=bool(return_to_d001),
                )
                buf = generate_excel_hybrid(
                    raw_df, sc_summary, diagnosis, diagnosis_om, diag_summary, _stats,
                    progress_callback=_cb_excel,
                    include_raw=_include_raw,
                )
                st.session_state.hybrid_processed = True
                st.session_state.hybrid_raw_df = raw_df
                st.session_state.hybrid_sc_summary = sc_summary
                st.session_state.hybrid_diagnosis = diagnosis
                st.session_state.hybrid_diagnosis_om = diagnosis_om
                st.session_state.hybrid_diag_summary = diag_summary
                st.session_state.hybrid_stats = _stats
                st.session_state.hybrid_excel_data = buf
                st.session_state.hybrid_params = _current_hybrid_params
                st.session_state.hybrid_warnings = _warnings
            except Exception as e:
                _prog2.empty()
                _tb = traceback.format_exc()
                print(f"❌ Safety Stock x NetStock Review 處理失敗: {e}\n{_tb}")
                st.error(f"❌ Safety Stock x NetStock Review 處理失敗: {e}\n\n詳細錯誤: {_tb}")
                st.stop()

            _prog2.empty()
            st.success("🎉 Safety Stock x NetStock Review 報表產生成功！數據彙總已完成。")
            if _warnings:
                for _w in _warnings:
                    st.warning(f"⚠️ {_w}")

        if st.session_state.hybrid_processed:
            raw_df = st.session_state.hybrid_raw_df
            sc_summary = st.session_state.hybrid_sc_summary
            diagnosis = st.session_state.hybrid_diagnosis
            diagnosis_om = st.session_state.hybrid_diagnosis_om
            diag_summary = st.session_state.hybrid_diag_summary
            stats = st.session_state.hybrid_stats
            excel_data = st.session_state.hybrid_excel_data
            applied_hybrid_days = (stats or {}).get("safety_days") or _safety_days
            applied_hybrid_dimension = (stats or {}).get("dimension") or _dimension
            try:
                applied_hybrid_ref = pd.to_datetime((stats or {}).get("ref_date")).date() if (stats or {}).get("ref_date") else hy_ref_date
            except (TypeError, ValueError):
                applied_hybrid_ref = hy_ref_date
            applied_hybrid_params = st.session_state.get("hybrid_params") or build_hybrid_params(
                (stats or {}).get("turnday_threshold", turnday_threshold),
                (stats or {}).get("d001_threshold", d001_threshold),
                applied_hybrid_days,
                (stats or {}).get("bidirectional", bidirectional),
                applied_hybrid_dimension,
                applied_hybrid_ref,
                (stats or {}).get("nd_priority", nd_priority),
                (stats or {}).get("sales_rebalance", sales_rebalance),
                (stats or {}).get("sales_oos_days", sales_oos_days),
                (stats or {}).get("sales_overstock_days", sales_overstock_days),
                (stats or {}).get("sales_overstock_min_units", sales_overstock_min_units),
                return_to_d001=(stats or {}).get("return_to_d001", return_to_d001),
            )
            hybrid_is_stale = hybrid_params_changed(applied_hybrid_params, _current_hybrid_params)
            if hybrid_is_stale:
                st.warning("目前設定已與最後一次執行結果不同，請重新執行後再下載報表；以下結果仍為最後一次執行內容。")

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 🔀 Safety Stock x NetStock Review 分析摘要 (Hybrid Dashboard)")
            _top_download_button(excel_data, "I14 Safety Stock x NetStock Review", "dl_top_hybrid")

            _hyb_rows = _hybrid_dashboard_metrics(raw_df, stats)

            for _row_m in _hyb_rows:
                _cols = st.columns(4)
                for _c, (_name, _val, _desc, _color) in zip(_cols, _row_m):
                    with _c:
                        st.markdown(f"""
                        <div class="metric-card" style="border-left-color: #{_color};">
                            <div class="metric-title">{_name}</div>
                            <div class="metric-value" style="color: #{_color};">{_val:,}</div>
                            <div class="metric-desc">{_desc}</div>
                        </div>
                        """, unsafe_allow_html=True)
                st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            try:
                _hyb_buyer_combined = _hybrid_buyer_combined_detail(raw_df)
            except Exception:
                _hyb_buyer_combined = []
            if _hyb_buyer_combined:
                st.markdown("#### 👤 Buyer 統計（需調貨＋需調整 Safety）")
                _b_total_need = sum(n for _, n, _, _, _ in _hyb_buyer_combined)
                _b_total_fix = sum(f for _, _, _, f, _ in _hyb_buyer_combined)
                _b_df = pd.DataFrame(
                    _hyb_buyer_combined,
                    columns=["Buyer", "需調貨 SKU 數", "調入需求量合計", "需調整 Safety SKU 數", "Safety 建議調整量合計"],
                )
                _b_df["需調貨 SKU 佔比"] = (_b_df["需調貨 SKU 數"] / _b_total_need * 100) if _b_total_need else 0.0
                _b_df["需調整 Safety SKU 佔比"] = (_b_df["需調整 Safety SKU 數"] / _b_total_fix * 100) if _b_total_fix else 0.0
                _b_df = _b_df[["Buyer", "需調貨 SKU 數", "需調貨 SKU 佔比", "調入需求量合計",
                               "需調整 Safety SKU 數", "需調整 Safety SKU 佔比", "Safety 建議調整量合計"]]
                st.dataframe(
                    _b_df, width='stretch', hide_index=True,
                    column_config={
                        "需調貨 SKU 佔比": st.column_config.NumberColumn("需調貨 SKU 佔比", format="%.1f%%"),
                        "需調整 Safety SKU 佔比": st.column_config.NumberColumn("需調整 Safety SKU 佔比", format="%.1f%%"),
                    },
                )
                st.caption(f"共 {len(_hyb_buyer_combined):,} 個 Buyer（已刪除重複；Buyer 取同 SKU 首個非空值，空白歸入「(未填 Buyer)」；需調整母體為 Safety 設定不足＋過高，貨量與調整量為淨額）。")

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            tab_diag, tab_diag_om, tab_sc, tab_issue, tab_return, tab_cross, tab_full, tab_detail, tab_raw = st.tabs([
                "🧭 調貨 Review Safety Stock（不限OM）",
                "🔗 調貨 Review Safety Stock（同OM）",
                "📋 Summary（SKU 存貨偏重）",
                "🏪 店舖問題 SKU",
                "📤 退D001 建議",
                "📊 Review 摘要（Safety 調整統計）",
                "🛡️ Safety 設定建議 (全列)",
                "🧾 簡易調貨明細",
                "📄 RAW 完整明細",
            ])
            with tab_diag:
                st.markdown("#### 店舖×SKU 逐列 Review — 不限OM調貨（需要調入／實際轉出／可轉出未使用；"
                            f"Safety 基準 {_hybrid_safety_days_label(applied_hybrid_days)}，與調貨的 14 日線獨立）")
                if diagnosis is not None and not diagnosis.empty:
                    st.dataframe(diagnosis.rename(columns=_display_safety_header), width='stretch', hide_index=True)
                    st.caption(f"共 {len(diagnosis):,} 列（按 Brand / SKU / 事件類型 / Site Code 排序）。")
                else:
                    st.info("無符合資料")
            with tab_diag_om:
                st.markdown("#### 店舖×SKU 逐列 Review — 同OM調貨（來源店須與需求店同 OM；Safety Review 與建議調整不受 OM 限制影響）")
                if diagnosis_om is not None and not diagnosis_om.empty:
                    st.dataframe(diagnosis_om.rename(columns=_display_safety_header), width='stretch', hide_index=True)
                    st.caption(f"共 {len(diagnosis_om):,} 列（按 Brand / SKU / 事件類型 / Site Code 排序）。")
                else:
                    st.info("無符合資料")
            _dim_label_applied = _HYBRID_DIMENSION_LABELS.get(applied_hybrid_dimension, "Shop Nature")
            with tab_sc:
                st.markdown(f"#### SKU 級彙總（按 {_dim_label_applied} 分組）：各組庫存/月銷加總、TurnDay、涉及店舖數、佔比 (%)、強制調貨提示、存貨偏重提示")
                _hy_zero_on = applied_hybrid_dimension in ("sales", "sales_nature")
                if sc_summary is not None and not sc_summary.empty:
                    _skewed = sc_summary[sc_summary["存貨偏重提示"].fillna("").astype(str) != ""] if "存貨偏重提示" in sc_summary.columns else sc_summary.iloc[0:0]
                    if not _skewed.empty:
                        st.markdown(f"**存貨偏重 SKU（{len(_skewed):,} 個）**")
                        st.dataframe(_style_inventory_skew(_skewed, zero_sales=_hy_zero_on, dimension=applied_hybrid_dimension), width='stretch', hide_index=True)
                    st.markdown("**全部 SKU Summary**")
                    st.dataframe(_style_inventory_skew(sc_summary, zero_sales=_hy_zero_on, dimension=applied_hybrid_dimension), width='stretch', hide_index=True)
                    _hy_zero_txt = "；零銷組庫存佔比大於 0% 以紅底標示（零銷售而仍有存貨，理論上全屬 Overstock）" if _hy_zero_on else ""
                    st.caption(f"共 {len(sc_summary):,} 個 SKU（按 Brand / SKU 排序）；存貨偏重提示 = {_hybrid_skew_hint_text(applied_hybrid_dimension)}；偏重數值以橘色底色加強顯示{_hy_zero_txt}；存貨偏低（{_hybrid_shortage_hint_text(applied_hybrid_dimension)}）以淺藍底標示庫存佔比欄（Excel 同步）。")
                else:
                    st.info("無符合資料")
            with tab_issue:
                st.markdown("#### 每間店舖需行動 SKU（問題＋建議方法）")
                try:
                    _hy_ref_txt = pd.to_datetime((stats or {}).get("ref_date")).date().isoformat()
                except (TypeError, ValueError):
                    _hy_ref_txt = ""
                _hy_ref_note = f"（參考日期 {_hy_ref_txt} 起 90 天內上市視為新品）" if _hy_ref_txt else ""
                st.markdown(f"只收錄 Safety 設定不足／過高列與 Past 90 Days 0 Sales 轉 ND 候選（四期皆零銷售且有庫存，新品除外{_hy_ref_note}；設定正常列不佔版面），按店舖分組、店內按嚴重度排序；同OM數值作對照欄；建議方法附建議值周轉日數，受下調補底 3 件影響時加註保底原因：")
                try:
                    _hy_issues = _build_hybrid_shop_issues(raw_df, ref_date=(stats or {}).get("ref_date"))
                except Exception:
                    _hy_issues = None
                if _hy_issues is not None and not _hy_issues.empty:
                    _hy_cols = [c for c in _HYBRID_SHOP_ISSUE_COLUMNS if c in _hy_issues.columns]
                    st.dataframe(_hy_issues[_hy_cols].rename(columns=_display_safety_header), width='stretch', hide_index=True)
                    st.caption(f"共 {len(_hy_issues):,} 列需行動（Safety 設定不足＋過高＋ND 退場；調入缺口另於建議方法加註）。")
                else:
                    st.info("無符合資料")
            with tab_return:
                st.markdown("#### 退D001 建議 — 退回超出 Safety 的庫存至 D001（Warehouse）")
                if _HYBRID_RETURN_D001_COLUMN in raw_df.columns:
                    st.markdown(
                        "接收目標為 **D001（Warehouse，可無限接收貨量）**，故不跑分配。"
                        "退貨基準為 **Safety Stock (現時)**：**中銷／低銷／零銷**列可退＝庫存 − Safety；"
                        "**高銷**列原則上不退，僅當 Safety > 0 且**庫存 ≥ 3×Safety（高於 Safety 200%）**時"
                        "可退＝庫存 − 2×Safety。"
                        "**調貨優先**：「建議退D001 量」＝可退D001 貨量 − 轉出量（不限OM），需要調貨列一律不退。"
                        "本表只作退倉建議，實際調貨仍按「調貨 Review Safety Stock」的分配結果。"
                    )
                    try:
                        _ret_df = _build_hybrid_return_d001(raw_df)
                    except Exception:
                        _ret_df = None
                    if _ret_df is not None and not _ret_df.empty:
                        _ret_cols = [c for c in _HYBRID_RETURN_D001_COLUMNS if c in _ret_df.columns]
                        st.dataframe(_ret_df[_ret_cols].rename(columns=_display_safety_header), width='stretch', hide_index=True)
                        _ret_sku_n = _ret_df["SKU"].nunique() if "SKU" in _ret_df.columns else 0
                        _ret_total = int(_ret_df[_HYBRID_RETURN_D001_COLUMN].sum())
                        st.caption(
                            f"共 {len(_ret_df):,} 列 / {_ret_sku_n:,} 個 SKU / 建議退回合計 {_ret_total:,} 件"
                            "（按 Site Code / Brand / SKU 排序；與 Excel「退D001 建議」工作表同一份清單）。"
                        )
                    else:
                        st.info("無符合資料（僅中銷／低銷／零銷列退回超出 Safety 的庫存；高銷列須庫存高於 Safety 200%；需要調貨列一律不退）")
                else:
                    st.info("未開啟「建議退D001」：本次結果不含退倉建議；請勾選開關後重新執行。")
            with tab_cross:
                st.markdown(f"#### ① Safety Review × {_dim_label_applied} 交叉表（列數 / SKU 數 / 建議調整量合計）")
                if diag_summary is not None and not diag_summary.empty:
                    st.dataframe(diag_summary, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")

                st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
                st.markdown("#### ② Safety 調整方向分類（增加 / 減少 / 不變）")
                _saf_change = _build_hybrid_safety_change_summary(raw_df)
                if not _saf_change.empty:
                    st.dataframe(_saf_change, width='stretch', hide_index=True)
                    st.caption("統計範圍為有調貨事件的店舖×SKU 列（與上方交叉表相同）；分類以「Safety 建議調整量」的正負為準"
                               "（正=增加、負=減少）；總計列的 SKU 數為不重複 SKU（同一 SKU 可跨分類），合計為淨調整量。")
                else:
                    st.info("無符合資料")

                st.markdown("#### ③ Safety 調整明細（SKU 級彙總）")
                _saf_sku = _build_hybrid_safety_sku_detail(raw_df)
                if not _saf_sku.empty:
                    st.dataframe(_saf_sku, width='stretch', hide_index=True)
                    st.caption(f"共 {len(_saf_sku):,} 個 SKU（含調貨事件列，按 Brand / SKU 排序；店舖數按 Site Code 不重複計算）；"
                               "「現有／建議 Safety 設定合計」只累加需要調整的列（Safety 設定不足＋過高）；"
                               "淨調整量 = 該 SKU 各列「Safety 建議調整量」加總。")
                else:
                    st.info("無符合資料")

                st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
                st.markdown("#### ④ Buyer 統計（需調貨＋需調整 Safety 合併表）")
                try:
                    _hyb_buyer_combined_tab = _hybrid_buyer_combined_detail(raw_df)
                except Exception:
                    _hyb_buyer_combined_tab = []
                if _hyb_buyer_combined_tab:
                    _bt_total_need = sum(n for _, n, _, _, _ in _hyb_buyer_combined_tab)
                    _bt_total_fix = sum(f for _, _, _, f, _ in _hyb_buyer_combined_tab)
                    _bt_df = pd.DataFrame(
                        _hyb_buyer_combined_tab,
                        columns=["Buyer", "需調貨 SKU 數", "調入需求量合計", "需調整 Safety SKU 數", "Safety 建議調整量合計"],
                    )
                    _bt_df["需調貨 SKU 佔比"] = (_bt_df["需調貨 SKU 數"] / _bt_total_need * 100) if _bt_total_need else 0.0
                    _bt_df["需調整 Safety SKU 佔比"] = (_bt_df["需調整 Safety SKU 數"] / _bt_total_fix * 100) if _bt_total_fix else 0.0
                    _bt_df = _bt_df[["Buyer", "需調貨 SKU 數", "需調貨 SKU 佔比", "調入需求量合計",
                                     "需調整 Safety SKU 數", "需調整 Safety SKU 佔比", "Safety 建議調整量合計"]]
                    st.dataframe(
                        _bt_df, width='stretch', hide_index=True,
                        column_config={
                            "需調貨 SKU 佔比": st.column_config.NumberColumn("需調貨 SKU 佔比", format="%.1f%%"),
                            "需調整 Safety SKU 佔比": st.column_config.NumberColumn("需調整 Safety SKU 佔比", format="%.1f%%"),
                        },
                    )
                    st.caption("同一 Buyer 的需調貨與需調整 Safety 並列顯示（已刪除重複；需調貨按 Stock Checking 同一準則；需調整母體為 Safety 設定不足＋過高，貨量與調整量為淨額）。")
                else:
                    st.info("無符合條件的 Buyer 統計")

                st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
                _soos = int((stats or {}).get("sales_oos_days") or _HYBRID_SALES_OOS_DAYS_DEFAULT)
                _sover = int((stats or {}).get("sales_overstock_days") or _HYBRID_SALES_OVERSTOCK_DAYS_DEFAULT)
                _sminu = int((stats or {}).get("sales_overstock_min_units") or _HYBRID_SALES_OVERSTOCK_MIN_UNITS_DEFAULT)
                st.markdown("#### ⑤ 銷量 × 存貨分類交叉表（母體為全部店舖×SKU 列）")
                _sales_block = _build_hybrid_sales_summary(raw_df)
                if not _sales_block.empty:
                    st.dataframe(_sales_block, width='stretch', hide_index=True)
                    st.caption(f"分類準則：高銷缺貨＝同 SKU 內屬高銷帶且店舖周轉日數 < {_soos} 日；"
                               f"低銷積壓＝低銷帶且周轉日數 ≥ {_sover} 日且庫存 ≥ {_sminu} 件；"
                               f"零銷積壓＝P90D 零銷售且庫存 ≥ {_sminu} 件；其餘為正常。"
                               "母體為全部店舖×SKU 列（有別於 ①–④ 的調貨事件列母體），無調貨事件時仍可查閱；"
                               "銷量帶按本次上傳資料內同 SKU 排名計算，篩選範圍不同時結果會不同。")
                else:
                    st.info("無符合資料")

                st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
                st.markdown("#### ⑥ 全列 Safety 調整方向分類（增加 / 減少 / 不變；母體為全部列）")
                _full_change = _build_hybrid_full_safety_change_summary(raw_df)
                if not _full_change.empty:
                    st.dataframe(_full_change, width='stretch', hide_index=True)
                    st.caption("統計範圍為全部店舖×SKU 列（預防性建議，不依賴調貨事件；ND 列不作上調）；"
                               "分類以「全列 Safety 建議調整量」的正負為準（正=增加、負=減少）；"
                               "總計列的 SKU 數為不重複 SKU，合計為淨調整量。")
                else:
                    st.info("無符合資料")

                st.markdown("#### ⑥ 全列 Safety 調整明細（SKU 級彙總）")
                _full_sku = _build_hybrid_full_safety_sku_detail(raw_df)
                if not _full_sku.empty:
                    st.dataframe(_full_sku, width='stretch', hide_index=True)
                    st.caption(f"共 {len(_full_sku):,} 個 SKU（全部 SKU，按 Brand / SKU 排序；店舖數按 Site Code 不重複計算）；"
                               "「現有／建議 Safety 設定合計」只累加需要調整的列；"
                               "淨調整量 = 該 SKU 各列「全列 Safety 建議調整量」加總。")
                else:
                    st.info("無符合資料")
            with tab_full:
                st.markdown("#### 全列預防性 Safety 設定建議 — 一次過設定適合的 Safety Stock，減少日後調貨次數")
                st.caption("母體為全部店舖×SKU 列，不依賴調貨事件；與事件版同口徑（Safety 周轉所需量＋下調補底 3 件）；"
                           "ND 列不作上調，只可下調；以下只列出需要調整的列（設定正常列不佔版面）。")
                _full_detail = _build_hybrid_full_safety_detail(raw_df)
                if _full_detail is not None and not _full_detail.empty:
                    st.dataframe(_full_detail.rename(columns=_display_safety_header), width='stretch', hide_index=True)
                    _full_inc = int((stats or {}).get("full_under_rows", 0))
                    _full_dec = int((stats or {}).get("full_over_rows", 0))
                    st.caption(f"共 {len(_full_detail):,} 列需調整（需增加 {_full_inc:,} 列／需減少 {_full_dec:,} 列；"
                               "按 Brand / SKU / Site Code 排序；與 Excel「Safety 設定建議 (全列)」工作表同一份清單）。")
                else:
                    st.info("無符合資料（全部列設定正常，無須調整）")
            with tab_detail:
                st.markdown("#### 簡易調貨明細（一列一筆來源店→接收店調貨；來源如 HB30:2; HB49:3 會拆成兩行）")
                _dv = st.radio(
                    "調貨方式", ["不限OM", "同OM"], horizontal=True, key="hy_detail_variant",
                    help="不限OM：來源固定 Non Tourist → Mix → Tourist 依序分配；"
                         "同OM：來源店須與接收店同 OM。Safety Review 與建議調整不受 OM 限制影響。"
                )
                _detail_df = _build_hybrid_transfer_detail(raw_df, variant="om" if _dv == "同OM" else "free")
                if not _detail_df.empty:
                    st.dataframe(_detail_df, width='stretch', hide_index=True)
                    st.caption(f"共 {len(_detail_df):,} 列（只收錄實際有調入的列，調貨數量按來源店逐行拆分）；"
                               "兩套結果另見下載 Excel 的「簡易調貨明細 (不限OM)／(同OM)」工作表。")
                else:
                    st.info("無實際調貨資料")
            with tab_raw:
                st.markdown("#### RAW 原始資料（Review 欄位已附於每一列；不顯示 Item_Shop 與 Value 金額欄）")
                _display_cols = [c for c in _hybrid_raw_display_columns(raw_df) if c in raw_df.columns]
                _display = raw_df[_display_cols]
                st.dataframe(_display.head(50).rename(columns=_display_safety_header), width='stretch')
                st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料（已按可讀性重排欄序）將包含在下載的 Excel 報表中。")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 Safety Stock x NetStock Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 Safety Stock x NetStock Review 分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                disabled=hybrid_is_stale,
                width='stretch',
            )
    elif mode == "Stock Impact Analysis":
        st.caption("📦 以 **Shop Stock Qty = 0（缺貨）** 為核心，分析存貨對 Sold Qty 的因果影響（無貨就無銷售）。重點留意 Tourist 店舖。")

        warehouse_only = st.checkbox(
            "🏗️ 只分析 Warehouse SKU（建議）",
            value=True,
            key="si_warehouse_only",
            help="Warehouse SKU 受 D001 倉庫補貨影響，缺貨分析最具意義。取消勾選可納入 Vendor SKU。"
        )
        st.caption("💡 Vendor SKU 無 D001 倉庫概念，缺貨多為供應商因素而非存貨配置問題。")

        mtd_date = st.date_input(
            "📅 參考日期 (Sold Qty CM-MTD 的截止日期)",
            value=datetime.now().date() - timedelta(days=1),
            key="si_mtd_date",
            help="僅作 Sold Qty (CM-MTD) 語意參考。本模式直接使用原始 CM-MTD 數值，不做全月預估投射。"
        )
        st.caption(f"參考日期: {mtd_date.strftime('%Y-%m-%d')}（CM = 當月至今 {mtd_date.day} 日）")

        stock_threshold = st.number_input(
            "📦 低存貨門檻（只影響低存貨配置決策，不限制整體分析範圍）",
            value=1000,
            min_value=0,
            step=10,
            key="si_stock_threshold",
            help="設定為 0 時不標記低存貨 SKU。\n"
            "當設定為 N（N > 0），系統會在「SKU 月銷量和存儲量分析」\n"
            "和「低存貨 SKU 配置決策」中標記 D001 + Shop Stock < N 的低存貨 SKU。\n"
            "整體分析（缺貨、Brand/Buyer 彙總等）不受此門檻限制，涵蓋所有 SKU。"
        )
        st.caption(f"當前門檻: {stock_threshold} 件 — 只影響低存貨標記與配置決策，不限制報表範圍。")

        st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

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

                _, summary_brand, summary_buyer, summary_category, summary_item_class, _, sku_diagnosis, inv_distribution, sku_monthly_analysis, raw_df, _warnings = transform_stock_impact(
                    df_input,
                    progress_callback=_cb_transform,
                    warehouse_only=warehouse_only,
                    stock_threshold=stock_threshold,
                    ref_date=mtd_date,
                )
                buf = generate_excel_stock_impact(summary_brand, summary_buyer, summary_category, summary_item_class, sku_diagnosis, inv_distribution, raw_df, progress_callback=_cb_excel, sku_monthly_analysis=sku_monthly_analysis, stock_threshold=stock_threshold, include_raw=_include_raw)

                st.session_state.stock_impact_processed = True
                st.session_state.stock_impact_summary_brand = summary_brand
                st.session_state.stock_impact_summary_buyer = summary_buyer
                st.session_state.stock_impact_summary_category = summary_category
                st.session_state.stock_impact_sku_diagnosis = sku_diagnosis
                st.session_state.stock_impact_inv_distribution = inv_distribution
                st.session_state.stock_impact_sku_monthly_analysis = sku_monthly_analysis
                # 只存 Dashboard 需要的欄位（節省 ~90% 記憶體）
                _slim_cols = ["Is Stock Impacted", "Shop Nature", "Shop Stock Qty (CM)",
                              "Turnday (Days)", "Mix Excess Stock", "Non Tourist Excess Stock",
                              "Adj. Safety Stock"]
                _available = [c for c in _slim_cols if c in raw_df.columns]
                st.session_state.stock_impact_raw_slim = raw_df[_available].copy() if _available else pd.DataFrame()
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
            summary_brand = st.session_state.stock_impact_summary_brand
            summary_buyer = st.session_state.stock_impact_summary_buyer
            summary_category = st.session_state.stock_impact_summary_category
            sku_diagnosis = st.session_state.stock_impact_sku_diagnosis
            inv_distribution = st.session_state.stock_impact_inv_distribution
            sku_monthly_analysis = st.session_state.stock_impact_sku_monthly_analysis
            raw_slim = st.session_state.stock_impact_raw_slim
            excel_data = st.session_state.stock_impact_excel_data

            from transformer import _safe_numeric, _safe_default_series
            import pandas as _pd

            # 存貨分佈指標計算（使用 raw_slim 節省記憶體）
            nature_s = raw_slim["Shop Nature"].fillna("").astype(str) if "Shop Nature" in raw_slim.columns else _pd.Series([""]*len(raw_slim))
            cm_stock_s = _safe_numeric(raw_slim.get("Shop Stock Qty (CM)", _safe_default_series(0, raw_slim.index)))
            turnday_s = _pd.to_numeric(raw_slim.get("Turnday (Days)", _safe_default_series(0, raw_slim.index)), errors="coerce")
            impacted_mask = raw_slim["Is Stock Impacted"] if "Is Stock Impacted" in raw_slim.columns else _pd.Series([False]*len(raw_slim))

            # Tourist 缺貨行數（高銷售店無貨）
            tourist_oos_rows = int(((nature_s == "Tourist") & (cm_stock_s <= 0) & impacted_mask).sum())
            # Tourist 缺貨數量（缺失件數 = max(0, Adj. Safety Stock - Shop Stock Qty (CM))）
            _adj_ss_s = _safe_numeric(raw_slim.get("Adj. Safety Stock", _safe_default_series(0, raw_slim.index)))
            _oos_mask = (nature_s == "Tourist") & (cm_stock_s <= 0) & impacted_mask
            tourist_oos_qty = int(np.maximum(0, (_adj_ss_s - cm_stock_s))[_oos_mask].sum())
            # 曾缺貨 Tourist 行數（含已恢復，不限 CM 當期缺貨）
            tourist_oos_broader = int(((nature_s == 'Tourist') & impacted_mask).sum())
            # Non Tourist 可釋存貨（>14日標準，與 Mix 一致）
            nt_excess_total = int(_safe_numeric(raw_slim.get("Non Tourist Excess Stock", _safe_default_series(0, raw_slim.index))).sum())
            # Mix 可釋放存貨件數（>14日標準）
            mix_excess_total = int(_safe_numeric(raw_slim.get("Mix Excess Stock", _safe_default_series(0, raw_slim.index))).sum())
            # 存貨失衡 SKU 數
            imbalance_sku = inv_distribution.get("imbalance_sku", _pd.DataFrame()) if inv_distribution else _pd.DataFrame()
            imbalance_count = len(imbalance_sku)
            # 可調貨優先分析（Tourist缺貨 + 同SKU在Mix/NT有可釋放件數）
            tp_data = inv_distribution.get("transfer_priority", {}) if inv_distribution else {}
            tp_summary = tp_data.get("summary", _pd.DataFrame()) if isinstance(tp_data, dict) else _pd.DataFrame()
            tp_detail = tp_data.get("detail", _pd.DataFrame()) if isinstance(tp_data, dict) else _pd.DataFrame()
            transferrable_tourist_rows = 0
            transferrable_sku_count = 0
            if tp_summary is not None and not tp_summary.empty:
                _s = tp_summary.iloc[0]
                transferrable_tourist_rows = int(_s.get("可調貨Tourist行數", 0))
                transferrable_sku_count = int(_s.get("可調貨SKU數", 0))
            # 各 Nature 平均 Turnday
            tourist_td = turnday_s[nature_s == "Tourist"]
            tourist_avg_td = round(float(tourist_td.mean()), 1) if tourist_td.notna().any() and len(tourist_td) else 0
            mix_td = turnday_s[nature_s == "Mix"]
            mix_avg_td = round(float(mix_td.mean()), 1) if mix_td.notna().any() and len(mix_td) else 0
            # Non Tourist 平均 Turnday
            nt_td = turnday_s[nature_s == "Non Tourist"]
            nt_avg_td = round(float(nt_td.mean()), 1) if nt_td.notna().any() and len(nt_td) else 0
            # 受影響最嚴重 Buyer
            top_buyer_name = "—"
            top_buyer_loss = 0
            if summary_buyer is not None and not summary_buyer.empty and "Buyer" in summary_buyer.columns:
                _buyer_data = summary_buyer[~summary_buyer["Buyer"].isin(["總計", ""])]
                if not _buyer_data.empty and "Tourist 缺貨行數" in _buyer_data.columns:
                    _top = _buyer_data.sort_values("Tourist 缺貨行數", ascending=False).iloc[0]
                    top_buyer_name = str(_top["Buyer"])
                    top_buyer_loss = int(_top.get("Tourist 缺貨行數", 0))

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📦 存貨分佈分析摘要 (Stock Impact Dashboard)")
            _top_download_button(excel_data, "I14 Stock Impact Analysis", "dl_top_impact")
            st.caption("🎯 重點：高銷售 Tourist 店無貨 vs Non Tourist/Mix 大量囤積｜Non Tourist/Mix 只需保留 7-14 日 Turnday")

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #DC2626;">
                    <div class="metric-title">⭐ Tourist 缺貨行數</div>
                    <div class="metric-value" style="color: #DC2626;">{tourist_oos_rows:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">行</span></div>
                    <div class="metric-desc">高銷售店無貨（只計 RF）｜單位：行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #F59E0B;">
                    <div class="metric-title">Non Tourist 可釋存貨</div>
                    <div class="metric-value" style="color: #F59E0B;">{nt_excess_total:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">件</span></div>
                    <div class="metric-desc">超過14日標準可釋放｜單位：件數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #F97316;">
                    <div class="metric-title">Mix 可釋放存貨</div>
                    <div class="metric-value" style="color: #F97316;">{mix_excess_total:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">件</span></div>
                    <div class="metric-desc">超過14日標準可釋放｜單位：件數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #EF4444;">
                    <div class="metric-title">存貨失衡 SKU 數</div>
                    <div class="metric-value" style="color: #EF4444;">{imbalance_count:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">個</span></div>
                    <div class="metric-desc">Tourist缺貨+其他過剩｜單位：SKU數</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            r2_col1, r2_col2, r2_col3, r2_col4 = st.columns(4)
            with r2_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #8B5CF6;">
                    <div class="metric-title">Tourist 平均Turnday</div>
                    <div class="metric-value" style="color: #8B5CF6;">{tourist_avg_td:.1f}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">日</span></div>
                    <div class="metric-desc">Tourist 平均存貨天數｜單位：日</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3B82F6;">
                    <div class="metric-title">Mix 平均Turnday</div>
                    <div class="metric-value" style="color: #3B82F6;">{mix_avg_td:.1f}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">日</span></div>
                    <div class="metric-desc">Mix 平均存貨天數｜單位：日</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #06B6D4;">
                    <div class="metric-title">Non Tourist 平均Turnday</div>
                    <div class="metric-value" style="color: #06B6D4;">{nt_avg_td:.1f}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">日</span></div>
                    <div class="metric-desc">Non Tourist 平均存貨天數｜單位：日</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #1E40AF;">
                    <div class="metric-title">受影響最嚴重 Buyer</div>
                    <div class="metric-value" style="color: #1E40AF; font-size: 1.1rem;">{top_buyer_name}</div>
                    <div class="metric-desc">Tourist 缺貨行數: {top_buyer_loss:,} 行</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            r3_col1, r3_col2, r3_col3, r3_col4 = st.columns(4)
            with r3_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #059669;">
                    <div class="metric-title">🔄 可調貨Tourist行數</div>
                    <div class="metric-value" style="color: #059669;">{transferrable_tourist_rows:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">行</span></div>
                    <div class="metric-desc">Tourist缺貨+同SKU在Mix/NT有可釋放件數｜單位：行數</div>
                </div>
                """, unsafe_allow_html=True)
            with r3_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #0D9488;">
                    <div class="metric-title">📦 可調貨SKU數</div>
                    <div class="metric-value" style="color: #0D9488;">{transferrable_sku_count:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">個</span></div>
                    <div class="metric-desc">涉及的不重複SKU｜單位：SKU數</div>
                </div>
                """, unsafe_allow_html=True)
            with r3_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #DC2626;">
                    <div class="metric-title">📉 Tourist 缺貨數量</div>
                    <div class="metric-value" style="color: #DC2626;">{tourist_oos_qty:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">件</span></div>
                    <div class="metric-desc">缺失件數 = Safety − Shop Net Stock｜單位：件數</div>
                </div>
                """, unsafe_allow_html=True)

            with r3_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #6366F1;">
                    <div class="metric-title">🕐 曾缺貨Tourist行數(含已恢復)</div>
                    <div class="metric-value" style="color: #6366F1;">{tourist_oos_broader:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">行</span></div>
                    <div class="metric-desc">任何一期曾缺貨的Tourist RF行數｜含已恢復供貨</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            low_stock_total = 0
            low_stock_item_class = pd.DataFrame()
            if sku_monthly_analysis is not None and not sku_monthly_analysis.empty and "低存貨" in sku_monthly_analysis.columns:
                low_stock_total = int(sku_monthly_analysis["低存貨"].sum())
                low_stock_item_class = _build_low_stock_item_class_summary(sku_monthly_analysis)

            r4_col1, r4_col2, r4_col3, r4_col4 = st.columns(4)
            with r4_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #DC2626;">
                    <div class="metric-title">🔴 低存貨 SKU 總數</div>
                    <div class="metric-value" style="color: #DC2626;">{low_stock_total:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">個</span></div>
                    <div class="metric-desc">D001+店舖存貨&lt;{stock_threshold}件｜單位：SKU數</div>
                </div>
                """, unsafe_allow_html=True)

            if low_stock_item_class is not None and not low_stock_item_class.empty:
                st.markdown("#### 🔴 低存貨 SKU 按 Item Class 分類")
                st.dataframe(low_stock_item_class, width='stretch', hide_index=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            impact_tabs = st.tabs(["📊 存貨分佈", "📦 SKU月銷量和存儲", "🏷️ By Brand", "👤 By Buyer", "📦 By Category", "🔍 SKU 缺貨 Review"])

            with impact_tabs[0]:
                st.markdown("#### 📊 存貨分佈分析")
                st.caption("🎯 焦點：Tourist 高銷售店無貨 vs Non Tourist/Mix 大量囤積；Non Tourist/Mix 只需保留 7-14 日 Turnday。")
                if inv_distribution:
                    mix_td = inv_distribution.get("mix_turnday")
                    nt_td = inv_distribution.get("nt_turnday")
                    nat_cmp = inv_distribution.get("nature_compare")
                    imb = inv_distribution.get("imbalance_sku")
                    if mix_td is not None and not mix_td.empty:
                        st.markdown("**Mix 店舖 Turnday 分級**")
                        st.dataframe(mix_td, width='stretch', hide_index=True)
                    if nt_td is not None and not nt_td.empty:
                        st.markdown("**Non Tourist 店舖 Turnday 分級**")
                        st.dataframe(nt_td, width='stretch', hide_index=True)
                    if nat_cmp is not None and not nat_cmp.empty:
                        st.markdown("**各 Shop Nature 存貨分佈對比**")
                        st.dataframe(nat_cmp, width='stretch', hide_index=True)
                    if imb is not None and not imb.empty:
                        st.markdown(f"**存貨失衡 SKU 清單**（{len(imb):,} 個 SKU，Tourist 缺貨 + Non Tourist/Mix 過剩）")
                        st.dataframe(imb, width='stretch', hide_index=True)
                        # 註解（計算方式說明）
                        if hasattr(imb, "attrs") and "annotation" in imb.attrs:
                            with st.expander("📝 計算方式說明"):
                                st.caption(imb.attrs["annotation"])
                    else:
                        st.info("無存貨失衡 SKU")
                    # B4: 低存貨 SKU 配置決策
                    low_alloc = inv_distribution.get("low_stock_allocation")
                    if low_alloc is not None and not low_alloc.empty:
                        st.markdown(f"**⚖️ 低存貨 SKU 配置決策**（D001 + Shop Stock < {int(low_alloc.iloc[0]['低存貨門檻(件)'])} 件 → 是否減少 Non Tourist RF 店舖）")
                        st.caption("NT_RF 無銷售有存貨 = 浪費應減店；ND 囤積 = 計畫淘汰佔用存貨")
                        # 轉置為 2 欄（指標/數值）方便閱讀
                        la_display = low_alloc.T.reset_index()
                        la_display.columns = ["指標", "數值"]
                        st.dataframe(la_display, width='stretch', hide_index=True)
                    else:
                        st.info("無低存貨 SKU 配置決策數據")
                    # B5: 可調貨優先 SKU 清單
                    if tp_detail is not None and not tp_detail.empty:
                        summary_text = f"可調貨Tourist行數: {transferrable_tourist_rows:,} 行｜可調貨SKU數: {transferrable_sku_count:,} 個"
                        st.markdown(f"**🔄 可調貨優先 SKU 清單**（Tourist缺貨 + 同SKU在Mix/NT有可釋放件數 → 優先調貨）")
                        st.caption(summary_text)
                        st.dataframe(tp_detail, width='stretch', hide_index=True)
                    else:
                        st.info("無可調貨優先 SKU（所有Tourist缺貨SKU在Mix/NT均無可釋放存貨）")
                else:
                    st.info("無存貨分佈數據")

            with impact_tabs[1]:
                st.markdown("#### 📦 SKU 月銷量和存儲量分析（按 Shop Nature）")
                st.caption("各 SKU 在 Tourist/Mix/Non Tourist 店舖的月均銷售（P90D/3）與當期庫存（CM Shop Stock）。低存貨 SKU 優先顯示。")
                if sku_monthly_analysis is not None and not sku_monthly_analysis.empty:
                    low_count = int(sku_monthly_analysis["低存貨"].sum()) if "低存貨" in sku_monthly_analysis.columns else 0
                    st.markdown(f"**共 {len(sku_monthly_analysis):,} 個 SKU** — 🔴 低存貨 SKU: {low_count:,} 個（D001 + Shop Stock < {stock_threshold} 件）")
                    st.dataframe(sku_monthly_analysis, width='stretch', hide_index=True)
                else:
                    st.info("無 SKU 月銷量/存儲量數據")

            with impact_tabs[2]:
                st.markdown("#### Brand 層級缺貨影響分析")
                st.caption("💡 「Tourist 缺貨行數」含已恢復供貨的店舖（不限當期 CM 缺貨）。Dashboard 的 ⭐「Tourist 缺貨行數」僅計當期仍在缺貨的店舖。")
                if summary_brand is not None and not summary_brand.empty:
                    st.dataframe(summary_brand, width='stretch', hide_index=True)
                else:
                    st.info("無 Brand 層級缺貨影響數據")

            with impact_tabs[3]:
                st.markdown("#### Buyer 層級缺貨影響分析")
                st.caption("💡 「Tourist 缺貨行數」含已恢復供貨的店舖（不限當期 CM 缺貨）。Dashboard 的 ⭐「Tourist 缺貨行數」僅計當期仍在缺貨的店舖。")
                if summary_buyer is not None and not summary_buyer.empty:
                    st.dataframe(summary_buyer, width='stretch', hide_index=True)
                else:
                    st.info("無 Buyer 層級缺貨影響數據")

            with impact_tabs[4]:
                st.markdown("#### Category (Adj. MCH1) 層級缺貨影響分析")
                st.caption("💡 「Tourist 缺貨行數」含已恢復供貨的店舖（不限當期 CM 缺貨）。Dashboard 的 ⭐「Tourist 缺貨行數」僅計當期仍在缺貨的店舖。")
                if summary_category is not None and not summary_category.empty:
                    st.dataframe(summary_category, width='stretch', hide_index=True)
                else:
                    st.info("無 Category 層級缺貨影響數據（若缺少 Adj. MCH1 欄位則無法產出）")

            with impact_tabs[5]:
                st.markdown("#### 🔍 by-SKU 缺貨 Review（Tourist + Mix + Non Tourist）")
                st.caption("每個缺貨 SKU 的 Tourist/Mix/Non Tourist 缺貨店數、D001 倉庫三期貨量與店舖總存貨。🟧 **倉庫有貨** = 補貨流程問題；⬜ **源頭斷貨** = 倉庫也斷貨需上游補貨。")
                if sku_diagnosis is not None and not sku_diagnosis.empty:
                    # 統計摘要
                    d001_restock = int((sku_diagnosis["D001 狀態"] == "倉庫有貨").sum())
                    d001_source = int((sku_diagnosis["D001 狀態"] == "源頭斷貨").sum())
                    st.markdown(f"**共 {len(sku_diagnosis):,} 個缺貨 SKU** — 🟧 補貨問題: {d001_restock:,} 個｜⬜ 源頭斷貨: {d001_source:,} 個")
                    st.dataframe(sku_diagnosis, width='stretch', hide_index=True)
                else:
                    st.info("無 Tourist / Mix / Non Tourist 缺貨 SKU")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

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
    elif mode == "Stock Impact Analysis (簡化)":
        st.caption("📦 簡化版缺貨銷售影響分析 — 輸出 **RAW**（含 Total Stock/SOH Tier 等擴充欄位）、**SKU Monthly Sales&Store**、**New**（店舖 SOH Tier 分佈）、**Stock Impact Dashboard** 四個工作表。")

        warehouse_only = st.checkbox(
            "🏗️ 只分析 Warehouse SKU（建議）",
            value=True,
            key="sir_warehouse_only",
            help="Warehouse SKU 受 D001 倉庫補貨影響，缺貨分析最具意義。取消勾選可納入 Vendor SKU。"
        )

        mtd_date = st.date_input(
            "📅 參考日期 (Sold Qty CM-MTD 的截止日期)",
            value=datetime.now().date() - timedelta(days=1),
            key="sir_mtd_date",
            help="僅作 Sold Qty (CM-MTD) 語意參考。"
        )

        stock_threshold = st.number_input(
            "📦 低存貨門檻",
            value=1000,
            min_value=0,
            step=10,
            key="sir_stock_threshold",
            help="設定為 0 時不標記低存貨 SKU。"
        )

        st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

        if st.button("⚙️ 執行簡化版缺貨銷售影響分析", type="primary", width='stretch'):
            _prog3 = st.progress(0, text="處理中... 0% | 開始處理...")
            _t0_3 = time.time()

            def _cb(pct, msg):
                _e = time.time() - _t0_3
                _eta3 = (_e / pct) - _e if pct > 0 else 0
                _prog3.progress(min(pct, 1.0), text=f"處理中... {pct*100:.0f}% | {msg} | 已用: {_e:.0f}s | 預計剩餘: {_eta3:.0f}s")

            try:
                def _cb_transform(pct, msg):
                    _cb(0.6 * pct, msg)

                def _cb_excel(pct, msg):
                    _cb(0.6 + 0.4 * pct, msg)

                _, _, _, _, _, _, _, _, _, raw_df, _warnings = transform_stock_impact(
                    df_input,
                    progress_callback=_cb_transform,
                    warehouse_only=warehouse_only,
                    stock_threshold=stock_threshold,
                    ref_date=mtd_date,
                )
                sku_monthly_simp = _build_sku_monthly_analysis_simplified(raw_df, stock_threshold=stock_threshold)
                sku_soh_summary = _build_sku_soh_tier_summary(raw_df)
                buf = generate_excel_stock_impact_simplified(raw_df, progress_callback=_cb_excel, sku_monthly_analysis=sku_monthly_simp, stock_threshold=stock_threshold, sku_soh_summary=sku_soh_summary, include_raw=_include_raw)

                # 儲存 Dashboard 需要的精簡欄位（節省 ~90% 記憶體）
                _slim_cols = ["Is Stock Impacted", "Shop Nature", "Shop Stock Qty (CM)",
                              "Turnday (Days)", "Mix Excess Stock", "Non Tourist Excess Stock",
                              "Adj. Safety Stock"]
                _available = [c for c in _slim_cols if c in raw_df.columns]
                st.session_state.stock_impact_simp_raw_slim = raw_df[_available].copy() if _available else pd.DataFrame()

                del raw_df
                gc.collect()

                st.session_state.stock_impact_simp_processed = True
                st.session_state.stock_impact_simp_sku_monthly_analysis = sku_monthly_simp
                st.session_state.stock_impact_simp_sku_soh_summary = sku_soh_summary
                st.session_state.stock_impact_simp_excel_data = buf
                st.session_state.stock_impact_simp_warnings = _warnings
            except Exception as e:
                _prog3.empty()
                _tb = traceback.format_exc()
                print(f"❌ Stock Impact Analysis (簡化) 處理失敗: {e}\n{_tb}")
                st.error(f"❌ Stock Impact Analysis (簡化) 處理失敗: {e}\n\n詳細錯誤: {_tb}")
                st.stop()

            _prog3.empty()
            st.success("🎉 簡化版缺貨銷售影響分析完成！")
            if _warnings:
                for _w in _warnings:
                    st.warning(f"⚠️ {_w}")

        if st.session_state.stock_impact_simp_processed:
            sku_monthly_simp = st.session_state.stock_impact_simp_sku_monthly_analysis
            excel_data = st.session_state.stock_impact_simp_excel_data
            raw_slim = st.session_state.stock_impact_simp_raw_slim

            # ── 簡化版 Dashboard：8 個指標（全從 raw_df 計算） ──
            if raw_slim is not None and not raw_slim.empty:
                from transformer import _safe_numeric, _safe_default_series
                import pandas as _pd

                nature_s = raw_slim["Shop Nature"].fillna("").astype(str) if "Shop Nature" in raw_slim.columns else _pd.Series([""]*len(raw_slim))
                cm_stock_s = _safe_numeric(raw_slim.get("Shop Stock Qty (CM)", _safe_default_series(0, raw_slim.index)))
                turnday_s = _pd.to_numeric(raw_slim.get("Turnday (Days)", _safe_default_series(0, raw_slim.index)), errors="coerce")
                impacted_mask = raw_slim["Is Stock Impacted"] if "Is Stock Impacted" in raw_slim.columns else _pd.Series([False]*len(raw_slim))

                # 指標計算
                tourist_oos_rows = int(((nature_s == "Tourist") & (cm_stock_s <= 0) & impacted_mask).sum())
                _adj_ss_s = _safe_numeric(raw_slim.get("Adj. Safety Stock", _safe_default_series(0, raw_slim.index)))
                _oos_mask = (nature_s == "Tourist") & (cm_stock_s <= 0) & impacted_mask
                tourist_oos_qty = int(np.maximum(0, (_adj_ss_s - cm_stock_s))[_oos_mask].sum())
                tourist_oos_broader = int(((nature_s == 'Tourist') & impacted_mask).sum())
                nt_excess_total = int(_safe_numeric(raw_slim.get("Non Tourist Excess Stock", _safe_default_series(0, raw_slim.index))).sum())
                mix_excess_total = int(_safe_numeric(raw_slim.get("Mix Excess Stock", _safe_default_series(0, raw_slim.index))).sum())
                tourist_td = turnday_s[nature_s == "Tourist"]
                tourist_avg_td = round(float(tourist_td.mean()), 1) if tourist_td.notna().any() and len(tourist_td) else 0
                mix_td = turnday_s[nature_s == "Mix"]
                mix_avg_td = round(float(mix_td.mean()), 1) if mix_td.notna().any() and len(mix_td) else 0
                nt_td = turnday_s[nature_s == "Non Tourist"]
                nt_avg_td = round(float(nt_td.mean()), 1) if nt_td.notna().any() and len(nt_td) else 0

                st.markdown('<div class="custom-card">', unsafe_allow_html=True)
                st.markdown("### 📦 存貨分佈分析摘要 (Stock Impact Dashboard - 簡化)")
                _top_download_button(excel_data, "I14 Stock Impact Analysis (Simplified)", "dl_top_impact_simp")
                st.caption("🎯 重點：高銷售 Tourist 店無貨 vs Non Tourist/Mix 大量囤積")

                r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
                with r1_col1:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #DC2626;">
                        <div class="metric-title">⭐ Tourist 缺貨行數</div>
                        <div class="metric-value" style="color: #DC2626;">{tourist_oos_rows:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">行</span></div>
                        <div class="metric-desc">高銷售店無貨（只計 RF）｜單位：行數</div>
                    </div>
                    """, unsafe_allow_html=True)
                with r1_col2:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #F59E0B;">
                        <div class="metric-title">Non Tourist 可釋存貨</div>
                        <div class="metric-value" style="color: #F59E0B;">{nt_excess_total:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">件</span></div>
                        <div class="metric-desc">超過14日標準可釋放｜單位：件數</div>
                    </div>
                    """, unsafe_allow_html=True)
                with r1_col3:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #F97316;">
                        <div class="metric-title">Mix 可釋放存貨</div>
                        <div class="metric-value" style="color: #F97316;">{mix_excess_total:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">件</span></div>
                        <div class="metric-desc">超過14日標準可釋放｜單位：件數</div>
                    </div>
                    """, unsafe_allow_html=True)
                with r1_col4:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #DC2626;">
                        <div class="metric-title">📉 Tourist 缺貨數量</div>
                        <div class="metric-value" style="color: #DC2626;">{tourist_oos_qty:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">件</span></div>
                        <div class="metric-desc">缺失件數 = Safety − Shop Net Stock｜單位：件數</div>
                    </div>
                    """, unsafe_allow_html=True)

                st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

                r2_col1, r2_col2, r2_col3, r2_col4 = st.columns(4)
                with r2_col1:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #8B5CF6;">
                        <div class="metric-title">Tourist 平均Turnday</div>
                        <div class="metric-value" style="color: #8B5CF6;">{tourist_avg_td:.1f}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">日</span></div>
                        <div class="metric-desc">Tourist 平均存貨天數｜單位：日</div>
                    </div>
                    """, unsafe_allow_html=True)
                with r2_col2:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #3B82F6;">
                        <div class="metric-title">Mix 平均Turnday</div>
                        <div class="metric-value" style="color: #3B82F6;">{mix_avg_td:.1f}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">日</span></div>
                        <div class="metric-desc">Mix 平均存貨天數｜單位：日</div>
                    </div>
                    """, unsafe_allow_html=True)
                with r2_col3:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #06B6D4;">
                        <div class="metric-title">Non Tourist 平均Turnday</div>
                        <div class="metric-value" style="color: #06B6D4;">{nt_avg_td:.1f}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">日</span></div>
                        <div class="metric-desc">Non Tourist 平均存貨天數｜單位：日</div>
                    </div>
                    """, unsafe_allow_html=True)
                with r2_col4:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #6366F1;">
                        <div class="metric-title">🕐 曾缺貨Tourist行數(含已恢復)</div>
                        <div class="metric-value" style="color: #6366F1;">{tourist_oos_broader:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">行</span></div>
                        <div class="metric-desc">任何一期曾缺貨的Tourist RF行數｜含已恢復供貨</div>
                    </div>
                    """, unsafe_allow_html=True)
                st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

                low_stock_total_s = 0
                low_stock_item_class_s = pd.DataFrame()
                if sku_monthly_simp is not None and not sku_monthly_simp.empty and "低存貨" in sku_monthly_simp.columns:
                    low_stock_total_s = int(sku_monthly_simp["低存貨"].sum())
                    low_stock_item_class_s = _build_low_stock_item_class_summary(sku_monthly_simp)

                # 有賣VS有貨新指標（從 sku_soh_summary 計算）
                cm_lt50_s = 0
                cm_lt30_s = 0
                cm_no_stock_s = 0
                band_df_s = pd.DataFrame()
                sku_soh = st.session_state.get("stock_impact_simp_sku_soh_summary")
                if sku_soh is not None and not sku_soh.empty:
                    cm_oh = pd.to_numeric(sku_soh.get("有貨店舖 (CM)", pd.Series(0)), errors="coerce").fillna(0)
                    cm_ratio = pd.to_numeric(sku_soh.get("有賣VS有貨 (CM)", pd.Series(np.nan)), errors="coerce")
                    cm_lt50_s = int(((cm_oh > 0) & (cm_ratio < 0.5)).sum())
                    cm_lt30_s = int(((cm_oh > 0) & (cm_ratio < 0.3)).sum())
                    cm_no_stock_s = int((cm_oh == 0).sum())
                    band_df_s = _build_sell_through_band_summary(sku_soh)

                r3_col1, r3_col2, r3_col3, r3_col4 = st.columns(4)
                with r3_col1:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #DC2626;">
                        <div class="metric-title">🔴 低存貨 SKU 總數</div>
                        <div class="metric-value" style="color: #DC2626;">{low_stock_total_s:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">個</span></div>
                        <div class="metric-desc">D001+店舖存貨&lt;{stock_threshold}件｜單位：SKU數</div>
                    </div>
                    """, unsafe_allow_html=True)
                with r3_col2:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #F97316;">
                        <div class="metric-title">📉 CM 有賣/有貨&lt;50% SKU</div>
                        <div class="metric-value" style="color: #F97316;">{cm_lt50_s:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">個</span></div>
                        <div class="metric-desc">有貨店舖>0 且 有賣÷有貨&lt;50%｜單位：SKU數</div>
                    </div>
                    """, unsafe_allow_html=True)
                with r3_col3:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #DC2626;">
                        <div class="metric-title">🔻 CM 有賣/有貨&lt;30% SKU</div>
                        <div class="metric-value" style="color: #DC2626;">{cm_lt30_s:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">個</span></div>
                        <div class="metric-desc">有貨店舖>0 且 有賣÷有貨&lt;30%｜單位：SKU數</div>
                    </div>
                    """, unsafe_allow_html=True)
                with r3_col4:
                    st.markdown(f"""
                    <div class="metric-card" style="border-left-color: #6B7280;">
                        <div class="metric-title">⬜ CM 完全無貨 SKU</div>
                        <div class="metric-value" style="color: #6B7280;">{cm_no_stock_s:,}<span style="font-size:0.45em; font-weight:normal; margin-left:4px;">個</span></div>
                        <div class="metric-desc">有貨店舖 (CM)=0｜單位：SKU數</div>
                    </div>
                    """, unsafe_allow_html=True)

                # 有賣VS有貨比率分佈表
                if band_df_s is not None and not band_df_s.empty:
                    st.markdown("#### 📊 有賣VS有貨比率分佈（按 SKU 數）")
                    st.dataframe(band_df_s, width='stretch', hide_index=True)

                if low_stock_item_class_s is not None and not low_stock_item_class_s.empty:
                    st.markdown("#### 🔴 低存貨 SKU 按 Item Class 分類")
                    st.dataframe(low_stock_item_class_s, width='stretch', hide_index=True)

                st.markdown('</div>', unsafe_allow_html=True)
                st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            if sku_monthly_simp is not None and not sku_monthly_simp.empty:
                low_count = int(sku_monthly_simp["低存貨"].sum()) if "低存貨" in sku_monthly_simp.columns else 0
                st.markdown(f"**共 {len(sku_monthly_simp):,} 個 SKU** — 🔴 低存貨 SKU: {low_count:,} 個")
                st.dataframe(sku_monthly_simp, width='stretch', hide_index=True)
            else:
                st.info("無 SKU 月銷量/存儲量數據")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 Stock Impact Analysis (Simplified)_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載簡化版缺貨銷售影響分析報表 (.xlsx)",
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
                key="oos_calc_mode",
                help="MTD：使用本月至今銷售（CM-MTD）推算全月；P90D：使用過去90天月均銷售（P90D）"
            )
            oos_calc_mode_value = "mtd" if oos_calc_mode.startswith("MTD") else "p90d"

            if oos_calc_mode_value == "mtd":
                mtd_date = st.date_input(
                    "📅 MTD 日期 (Sold Qty CM-MTD 的截止日期)",
                    value=datetime.now().date() - timedelta(days=1),
                    key="oo_mtd_date",
                    help="選擇 MTD (Month-To-Date) 截止日期，系統將據此推算月度預估銷售量。"
                )
            else:
                mtd_date = st.date_input(
                    "📅 MTD 日期 (僅供 Projected Monthly 顯示)",
                    value=datetime.now().date() - timedelta(days=1),
                    key="oo_mtd_date",
                    help="Projected Monthly Sold Qty 欄位顯示用，不影響 Safety Qty 建議計算。"
                )
            mtd_days = mtd_date.day
            month_total_days = calendar.monthrange(mtd_date.year, mtd_date.month)[1]
            st.caption(f"MTD: {mtd_days}日，當月共 {month_total_days} 日")

        if mode == "NDRF 月份對比":
            _all_months = sorted(df_input["Month"].dropna().unique())
            if len(_all_months) < 2:
                st.error(f"❌ 來源檔案的「Month」欄位至少需有 2 個不同月份值，目前僅有: {_all_months}")
                st.stop()
            col_month_a, col_month_b = st.columns(2)
            with col_month_a:
                month_a_selected = st.selectbox(
                    "📅 選擇對比月份 A",
                    options=_all_months,
                    index=0,
                    key="ndrf_month_a",
                    help="選擇月份 A（對比的基準月份）"
                )
            with col_month_b:
                month_b_selected = st.selectbox(
                    "📅 選擇對比月份 B",
                    options=_all_months,
                    index=min(1, len(_all_months) - 1),
                    key="ndrf_month_b",
                    help="選擇月份 B（對比的比較月份）"
                )
            if month_a_selected == month_b_selected:
                st.warning("⚠️ 月份 A 與月份 B 相同，對比將無差異。請選擇不同月份。")

        if mode == "Stock Checking":
            sc_top = st.columns(2)
            with sc_top[0]:
                turnday_threshold = st.number_input(
                    "🔁 Tourist 加總 TurnDay 門檻 (日)",
                    min_value=1, max_value=365, value=14, step=1,
                    key="sc_ttd",
                    help="該 SKU 所有 Tourist 店的庫存加總 ÷ 月銷加總 × 30.5（與 Summary 同一準則）低於此值時，才列入調貨候選。"
                )
            with sc_top[1]:
                d001_threshold = st.number_input(
                    "🏭 D001 庫存門檻 (件)",
                    min_value=0, max_value=1000000, value=100, step=10,
                    key="sc_d001",
                    help="當 D001 Net Stock Qty (CM) 低於此值且 Tourist 加總 TurnDay 低於門檻時，才標示需要調貨。"
                )
            st.caption(
                f"🔄 觸發條件（同時符合）：該 SKU 的 **Tourist 加總 TurnDay < {int(turnday_threshold)} 日** "
                f"且 **D001 Net Stock < {int(d001_threshold)} 件**，且 **Mix / Non Tourist 至少一個 Nature 的加總 TurnDay ≥ {int(turnday_threshold)} 日**（有貨可調）。"
                "加總 TurnDay ≥ 門檻的 SKU 一律無須調貨；Mix 與 Non Tourist 加總 TurnDay 皆低於門檻時視為無貨可調，不標示。"
                "只有已標示 SKU 的 RAW 才計算調貨數量；調入需求補到 14 日周轉，來源優先 Non Tourist 店舖，再由 Mix 店舖補足（RF 與 ND 店皆可轉出）。"
            )
            _sc_dim_label = st.radio(
                "🗂️ 分類維度",
                options=list(_DIMENSION_OPTIONS),
                index=0,
                horizontal=True,
                key="sc_dim",
                help="決定 Summary 的分組欄位。" + _DIMENSION_HELP_COMMON +
                     "\n本模式的調貨計算一律按 Shop Nature 準則（上述觸發條件），不隨維度改變。"
                     "RAW／店舖問題 SKU 另恆附「銷量帶」與「銷量存貨分類」顯示欄（見下方門檻）。"
            )
            _sc_dimension = _DIMENSION_OPTIONS[_sc_dim_label]
            _sc_sales_oos, _sc_sales_over, _sc_sales_minu = _sales_threshold_inputs("sc")

        if mode == "Forecast Review":
            fc_mtd_date = st.date_input(
                "📅 MTD 日期 (Sold Qty CM-MTD 的截止日期)",
                value=datetime.now().date() - timedelta(days=1),
                key="fc_mtd_date",
                help="MTD 月化＝Sold Qty (CM - MTD) ÷ 有效日數 × 當月日數（開賣未滿月按開賣後有效日數）。預設為昨日，與報表資料基準一致。"
            )
            _fc_mtd_days = fc_mtd_date.day
            _fc_month_total_days = calendar.monthrange(fc_mtd_date.year, fc_mtd_date.month)[1]
            st.caption(
                f"MTD: {_fc_mtd_days}日，當月共 {_fc_month_total_days} 日。"
                "v2 需求模型：MTD 月化按已過日數與月均基準混合（月初 MTD 佔比較低），樂觀情境再乘趨勢衰減係數。"
            )
            fc_top = st.columns(3)
            with fc_top[0]:
                fc_target_days = st.number_input(
                    "🎯 目標覆蓋天數 (日)",
                    min_value=_FC_TARGET_DAYS_MIN, max_value=_FC_TARGET_DAYS_MAX,
                    value=_FC_TARGET_DAYS_DEFAULT, step=1,
                    key="fc_target_days",
                    help="平穩類目標覆蓋量＝ceil(月預估 ÷ 30.5 × 天數)。預設 30 日。"
                )
            with fc_top[1]:
                fc_cap_months = st.number_input(
                    "🧱 上限月數 (個月)",
                    min_value=_FC_CAP_MONTHS_MIN, max_value=_FC_CAP_MONTHS_MAX,
                    value=_FC_CAP_MONTHS_DEFAULT, step=0.5,
                    key="fc_cap_months",
                    help="上限空間＝ceil(月預估 × 月數) − 位置；建議訂量＝min(缺口, 上限空間)。防止嚴重 Overstock 的封頂。預設 2.0 個月。"
                )
            with fc_top[2]:
                fc_surge_cover_months = st.number_input(
                    "🔥 爆升目標覆蓋月數 (個月)",
                    min_value=_FC_SURGE_COVER_MIN, max_value=_FC_SURGE_COVER_MAX,
                    value=_FC_SURGE_COVER_DEFAULT, step=0.25,
                    key="fc_surge_cover_months",
                    help="新起量／加速中／爆升的目標覆蓋量＝ceil(月預估 × 月數)；平穩類沿用目標覆蓋天數。"
                         "預設 2.0 個月（熱銷 SKU 預留較長覆蓋，仍受上限月數封頂）。"
                )
            fc_mid = st.columns(3)
            with fc_mid[0]:
                fc_surge_mult = st.number_input(
                    "🚀 爆升倍數 (倍)",
                    min_value=_FC_SURGE_MULT_MIN, max_value=_FC_SURGE_MULT_MAX,
                    value=_FC_SURGE_MULT_DEFAULT, step=0.1,
                    key="fc_surge_mult",
                    help="max(M-1, MTD 月化) ÷ max(M-2, 新起量基數) 達此倍數判爆升，用樂觀情境。預設 1.5 倍。"
                )
            with fc_mid[1]:
                fc_surge_min_base = st.number_input(
                    "🌱 新起量基數 (件)",
                    min_value=_FC_SURGE_MIN_BASE_MIN, max_value=_FC_SURGE_MIN_BASE_MAX,
                    value=_FC_SURGE_MIN_BASE_DEFAULT, step=1,
                    key="fc_surge_min_base",
                    help="M-2＋M-1 合計不超過此數且近期有量 → 新起量（用樂觀情境）。預設 10 件。"
                )
            with fc_mid[2]:
                fc_accel_mult = st.number_input(
                    "📈 加速倍數 (倍)",
                    min_value=_FC_ACCEL_MULT_MIN, max_value=_FC_ACCEL_MULT_MAX,
                    value=_FC_ACCEL_MULT_DEFAULT, step=0.05,
                    key="fc_accel_mult",
                    help="M-1 ≥ 此倍數 × M-2 且 MTD 月化 ≥ 此倍數 × M-1 → 加速中（用樂觀情境）。預設 1.15 倍。"
                )
            fc_bot = st.columns(2)
            with fc_bot[0]:
                fc_damp = st.number_input(
                    "🎚️ 衰減係數",
                    min_value=_FC_DAMP_MIN, max_value=_FC_DAMP_MAX,
                    value=_FC_DAMP_DEFAULT, step=0.05,
                    key="fc_damp",
                    help="樂觀情境＝max(M-1 折算後, MTD 調整月化) × (1 ＋ 係數 × 趨勢增速)，避免直接追逐尖峰。預設 0.8。"
                )
            with fc_bot[1]:
                fc_show_surge_only = st.checkbox(
                    "🔥 只顯示爆升 SKU（新起量／加速中／爆升）",
                    value=True,
                    key="fc_surge_only",
                    help="只影響主表與店舖覆蓋表的顯示；摘要與按 Buyer 彙總一律按全母體計算。"
                )
            with st.expander("⚙️ 進階模型設定 (v2 加權混合)", expanded=False):
                st.caption(
                    "v2 需求模型（SKU 級聚合）：標準情境＝三項加權混合；樂觀情境＝趨勢外推（增速上限＋衰減）。"
                    "三項權重會自動正規化；全為 0 時退回預設。"
                )
                _fc_w_cols = st.columns(3)
                with _fc_w_cols[0]:
                    fc_w_p90d = st.number_input(
                        "⚖️ 月均基準權重",
                        min_value=_FC_WEIGHT_MIN, max_value=_FC_WEIGHT_MAX,
                        value=_FC_WEIGHT_P90D_DEFAULT, step=0.05,
                        key="fc_w_p90d",
                        help="Shop Monthly Sold Qty (P90D) 月均在標準情境的權重。預設 0.4。"
                    )
                with _fc_w_cols[1]:
                    fc_w_m1 = st.number_input(
                        "⚖️ M-1 權重",
                        min_value=_FC_WEIGHT_MIN, max_value=_FC_WEIGHT_MAX,
                        value=_FC_WEIGHT_M1_DEFAULT, step=0.05,
                        key="fc_w_m1",
                        help="上月銷售（經缺貨還原與開賣折算）在標準情境的權重。預設 0.3。"
                    )
                with _fc_w_cols[2]:
                    fc_w_mtd = st.number_input(
                        "⚖️ MTD 調整月化權重",
                        min_value=_FC_WEIGHT_MIN, max_value=_FC_WEIGHT_MAX,
                        value=_FC_WEIGHT_MTD_DEFAULT, step=0.05,
                        key="fc_w_mtd",
                        help="MTD 月化按已過日數與月均基準混合後，在標準情境的權重。預設 0.3。"
                    )
                fc_trend_cap = st.number_input(
                    "📈 趨勢增速上限 (倍)",
                    min_value=_FC_TREND_CAP_MIN, max_value=_FC_TREND_CAP_MAX,
                    value=_FC_TREND_CAP_DEFAULT, step=0.1,
                    key="fc_trend_cap",
                    help="樂觀情境的趨勢增速 g＝M-1 ÷ M-2 − 1（只取正增速）上限；預設 1.0 倍（最多翻倍）。"
                )
            st.caption(
                f"📐 本次參數：目標覆蓋 **{int(fc_target_days)} 日**｜爆升目標覆蓋 **{float(fc_surge_cover_months):g} 個月**｜上限 **{float(fc_cap_months):g} 個月**｜"
                f"爆升 **{float(fc_surge_mult):g} 倍**（基數 {int(fc_surge_min_base)} 件）｜加速 **{float(fc_accel_mult):g} 倍**｜"
                f"衰減 **{float(fc_damp):g}**｜權重 **{float(fc_w_p90d):g}／{float(fc_w_m1):g}／{float(fc_w_mtd):g}**｜趨勢上限 **{float(fc_trend_cap):g} 倍**。"
                "Warehouse(D001) 貨位置＝店存＋D001 Net＋D001 PO＋在途，落單：D001 向 Vendor 下單補貨（D001 PO）；"
                "Vendor 貨位置＝店存＋Shop PO＋在途，落單：店舖向供應商下單（Shop PO）。"
            )

        if mode == "新品追蹤 Review":
            ns_ref_date = st.date_input(
                "📅 參考日期（Launch Date 起 90 天內視為新品）",
                value=datetime.now().date() - timedelta(days=1),
                key="ns_ref_date_input",
                help="上市日數＝參考日期 − Launch Date；0 至 90 天內視為新品。預設為昨日，與報表資料基準一致。"
            )
            st.caption(f"參考日期：{ns_ref_date.isoformat()}（上市 90 天內視為新品；無法解析或晚於參考日期的列視為非新品）。")

        st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

        SHOP_NATURE_OPTIONS = ["Tourist", "Mix", "Non Tourist"]
        if mode in ("Stock Checking", "Forecast Review"):
            # 調貨分配／Forecast SKU 彙總需要完整涵蓋所有 Shop Nature，此兩個模式不做 Nature 篩選
            selected_shop_natures = SHOP_NATURE_OPTIONS
        else:
            selected_shop_natures = st.multiselect(
                "🏷️ 選擇要處理的 Shop Nature 類型",
                options=SHOP_NATURE_OPTIONS,
                default=SHOP_NATURE_OPTIONS,
                key="oo_natures",
                help="選擇要處理的店舖性質。預設全選所有類型。"
            )

        _is_overstock_mode = (mode == "Overstock Review")
        _is_carry_mode = (mode == "Carry SKU Review")
        # Modes that don't use the Shop Nature Rate sliders (no Safety Qty calculation):
        # Carry SKU Review, SKU 精簡 Review, NDRF 月份對比 (pivot/comparison only).
        _skip_rates_mode = _is_carry_mode or (mode == "SKU 精簡 Review") or (mode == "NDRF 月份對比") or (mode == "Stock Checking") or (mode == "Forecast Review") or (mode == "新品追蹤 Review")

        if not _skip_rates_mode:
            _tourist_default = 30 if _is_overstock_mode else 50
            _mix_default = 40 if _is_overstock_mode else 50
            _non_tourist_default = 50

            with st.expander("⚙️ 自定義 Shop Nature Rate (進階設定)", expanded=False):
                st.markdown("調整各 Shop Nature 類型的安全庫存建議倍率（% of Monthly Sold Qty）。")
                _rate_cols = st.columns(3)
                with _rate_cols[0]:
                    rate_tourist = st.slider("Tourist Rate (%)", 0, 100, _tourist_default, 1,
                        key="oo_rate_tourist",
                        help="Tourist 類型：安全庫存建議 = floor(Monthly Sold × Rate)，最低 3")
                with _rate_cols[1]:
                    rate_mix = st.slider("Mix Rate (%)", 0, 100, _mix_default, 1,
                        key="oo_rate_mix",
                        help="Mix 類型：安全庫存建議 = floor(Monthly Sold × Rate)，最低 3")
                with _rate_cols[2]:
                    rate_non_tourist = st.slider("Non Tourist Rate (%)", 0, 100, _non_tourist_default, 1,
                        key="oo_rate_non_tourist",
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
                key="oo_safety_mode",
                help="只可增加：Safety Qty 只增不減（不處理 Overstock 的 Safety）；可加可減：同步處理 Overstock 的 Safety，可升可降。"
            )
            safety_mode_value = "increase_only" if safety_mode == "只可增加" else "allow_both"

        # P0 全域選項已於檔案預覽後定義（_include_raw），此處直接沿用
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
                    buf = generate_excel(summary_df, raw_df, categories, progress_callback=_cb_excel, include_raw=_include_raw)
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
                    buf = generate_excel_carry(summary_left, summary_right, dashboard_df, raw_df, progress_callback=_cb_excel, include_raw=_include_raw)
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
                    buf = generate_excel_nd(summary_df, raw_df, categories, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.nd_processed = True
                    st.session_state.nd_summary_df = summary_df
                    st.session_state.nd_raw_df = raw_df
                    st.session_state.nd_categories = categories
                    st.session_state.nd_excel_data = buf
                    st.session_state.nd_warnings = _warnings
                elif mode == "NDRF 月份對比":
                    comparison_by_shop, adj_sst_pivot, mch2_pivot, raw_df, _warnings = transform_ndrf_comparison(
                        df_input,
                        month_a=month_a_selected,
                        month_b=month_b_selected,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                    )
                    buf = generate_excel_ndrf_comparison(
                        comparison_by_shop, adj_sst_pivot, mch2_pivot, raw_df,
                        month_a_selected, month_b_selected, progress_callback=_cb_excel,
                        include_raw=_include_raw,
                    )
                    st.session_state.ndrf_comp_processed = True
                    st.session_state.ndrf_comp_comparison_by_shop = comparison_by_shop
                    st.session_state.ndrf_comp_adj_sst_pivot = adj_sst_pivot
                    st.session_state.ndrf_comp_mch2_pivot = mch2_pivot
                    st.session_state.ndrf_comp_raw_df = raw_df
                    st.session_state.ndrf_comp_excel_data = buf
                    st.session_state.ndrf_comp_warnings = _warnings
                elif mode == "OOS Review":  # OOS Review
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
                    buf = generate_excel_oos(summary_df, raw_df, categories, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.oos_processed = True
                    st.session_state.oos_summary_df = summary_df
                    st.session_state.oos_raw_df = raw_df
                    st.session_state.oos_categories = categories
                    st.session_state.oos_excel_data = buf
                    st.session_state.oos_warnings = _warnings
                elif mode == "Past 90 Days 0 Sales":
                    raw_df, summary_df, _warnings = transform_p90d_zero_sales(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                    )
                    buf = generate_excel_p90d_zero_sales(raw_df, summary_df, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.p90ds_processed = True
                    st.session_state.p90ds_summary_df = summary_df
                    st.session_state.p90ds_raw_df = raw_df
                    st.session_state.p90ds_excel_data = buf
                    st.session_state.p90ds_warnings = _warnings
                elif mode == "銷售損失量化 Review":
                    raw_df, by_sku_df, by_shop_df, _warnings = transform_sales_loss(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                    )
                    buf = generate_excel_sales_loss(raw_df, by_sku_df, by_shop_df, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.loss_processed = True
                    st.session_state.loss_raw_df = raw_df
                    st.session_state.loss_by_sku = by_sku_df
                    st.session_state.loss_by_shop = by_shop_df
                    st.session_state.loss_excel_data = buf
                    st.session_state.loss_warnings = _warnings
                elif mode == "積壓資金 Review":
                    raw_df, by_sku_df, by_shop_df, _warnings = transform_excess_capital(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                    )
                    buf = generate_excel_excess_capital(raw_df, by_sku_df, by_shop_df, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.exc_processed = True
                    st.session_state.exc_raw_df = raw_df
                    st.session_state.exc_by_sku = by_sku_df
                    st.session_state.exc_by_shop = by_shop_df
                    st.session_state.exc_excel_data = buf
                    st.session_state.exc_warnings = _warnings
                elif mode == "RP Type 覆核 Review":
                    raw_df, by_sku_df, by_shop_df, _warnings = transform_rp_review(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                    )
                    buf = generate_excel_rp_review(raw_df, by_sku_df, by_shop_df, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.rp_processed = True
                    st.session_state.rp_raw_df = raw_df
                    st.session_state.rp_by_sku = by_sku_df
                    st.session_state.rp_by_shop = by_shop_df
                    st.session_state.rp_excel_data = buf
                    st.session_state.rp_warnings = _warnings
                elif mode == "D001 健康檢查 Review":
                    check_df, by_buyer_df, raw_df, _warnings = transform_d001_health(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                    )
                    buf = generate_excel_d001_health(check_df, by_buyer_df, raw_df, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.d001_processed = True
                    st.session_state.d001_check_df = check_df
                    st.session_state.d001_by_buyer = by_buyer_df
                    st.session_state.d001_raw_df = raw_df
                    st.session_state.d001_excel_data = buf
                    st.session_state.d001_warnings = _warnings
                elif mode == "新品追蹤 Review":
                    raw_df, by_sku_df, by_shop_df, _warnings = transform_new_sku(
                        df_input,
                        progress_callback=_cb_transform,
                        shop_natures=selected_shop_natures if set(selected_shop_natures) != set(SHOP_NATURE_OPTIONS) else None,
                        ref_date=ns_ref_date,
                    )
                    buf = generate_excel_new_sku(raw_df, by_sku_df, by_shop_df, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.ns_processed = True
                    st.session_state.ns_raw_df = raw_df
                    st.session_state.ns_by_sku = by_sku_df
                    st.session_state.ns_by_shop = by_shop_df
                    st.session_state.ns_excel_data = buf
                    st.session_state.ns_warnings = _warnings
                    st.session_state.ns_ref_date = ns_ref_date.isoformat()
                elif mode == "Stock Checking":
                    raw_df, summary_df, _stats, _warnings = transform_stock_checking(
                        df_input,
                        turnday_threshold=int(turnday_threshold),
                        d001_threshold=int(d001_threshold),
                        progress_callback=_cb_transform,
                        dimension=_sc_dimension,
                        sales_oos_days=_sc_sales_oos,
                        sales_overstock_days=_sc_sales_over,
                        sales_overstock_min_units=_sc_sales_minu,
                    )
                    buf = generate_excel_stock_checking(raw_df, summary_df, _stats, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.stock_chk_processed = True
                    st.session_state.stock_chk_summary_df = summary_df
                    st.session_state.stock_chk_raw_df = raw_df
                    st.session_state.stock_chk_stats = _stats
                    st.session_state.stock_chk_excel_data = buf
                    st.session_state.stock_chk_warnings = _warnings
                elif mode == "Forecast Review":
                    _main_df, _cov_df, _buyer_df, _raw_df, _fc_stats, _warnings = transform_forecast(
                        df_input,
                        mtd_days=int(_fc_mtd_days),
                        month_total_days=int(_fc_month_total_days),
                        target_days=float(fc_target_days),
                        cap_months=float(fc_cap_months),
                        surge_mult=float(fc_surge_mult),
                        surge_min_base=float(fc_surge_min_base),
                        accel_mult=float(fc_accel_mult),
                        damp=float(fc_damp),
                        w_p90d=float(fc_w_p90d),
                        w_m1=float(fc_w_m1),
                        w_mtd=float(fc_w_mtd),
                        trend_cap=float(fc_trend_cap),
                        surge_cover_months=float(fc_surge_cover_months),
                        mtd_date=fc_mtd_date,
                        progress_callback=_cb_transform,
                    )
                    _fc_stats = dict(_fc_stats or {})
                    _fc_stats["mtd_date"] = fc_mtd_date.isoformat()
                    _fc_stats["surge_only"] = bool(fc_show_surge_only)
                    buf = generate_excel_forecast(_main_df, _cov_df, _buyer_df, _raw_df, _fc_stats, progress_callback=_cb_excel, include_raw=_include_raw)
                    st.session_state.fc_processed = True
                    st.session_state.fc_main_df = _main_df
                    st.session_state.fc_coverage_df = _cov_df
                    st.session_state.fc_buyer_df = _buyer_df
                    st.session_state.fc_raw_df = _raw_df
                    st.session_state.fc_stats = _fc_stats
                    st.session_state.fc_excel_data = buf
                    st.session_state.fc_warnings = _warnings
            except Exception as e:
                _prog2.empty()
                _tb = traceback.format_exc()
                print(f"❌ 報表處理失敗: {e}\n{_tb}")
                st.error(f"❌ 報表處理失敗: {e}\n\n詳細錯誤: {_tb}")
                st.stop()

            _prog2.empty()
            _total_s = time.time() - _t0_2
            _rss_after = _process_rss_mb()
            _rss_txt = f"｜記憶體 {_rss_after:.0f} MB" if _rss_after else ""
            _raw_txt = "（已略過 RAW）" if not _include_raw else ""
            st.success(f"🎉 報表產生成功！數據彙總已完成。耗時 {_total_s:.1f} 秒{_raw_txt}{_rss_txt}")
            st.session_state.last_run_seconds = round(_total_s, 2)
            gc.collect()
            if _warnings:
                for _w in _warnings:
                    st.warning(f"⚠️ {_w}")

        # 顯示轉換成果區塊
        if mode == "Overstock Review" and st.session_state.processed:
            summary_df = st.session_state.summary_df
            raw_df = st.session_state.raw_df
            categories = st.session_state.categories
            excel_data = st.session_state.excel_data

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📈 異常庫存數據分析摘要 (Analysis Dashboard)")
            _top_download_button(excel_data, "I14 by SKUxShop_To Buyer", "dl_top_overstock")

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
                st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

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
            _top_download_button(excel_data, "I14 OOS Review", "dl_top_oos")

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
                st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

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

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📈 ND SKU 數據分析摘要 (ND SKU Analysis Dashboard)")
            _top_download_button(excel_data, "I14 SKU 精簡 Review", "dl_top_nd")

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

            _all_site_codes = canonical_site_codes(df_input["Site Code"]) if "Site Code" in df_input.columns else []
            selected_site_code = st.selectbox(
                "🏪 選擇 Site Code 檢視明細",
                options=["全部"] + _all_site_codes,
                index=0,
                key="oo_site",
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
                    _filtered_raw = filter_by_site_code(_filtered_raw, selected_site_code)
                st.dataframe(_filtered_raw.head(50).rename(columns=_display_safety_header), width='stretch')
                st.caption(f"僅顯示前 50 筆（{'全部' if selected_site_code == '全部' else selected_site_code}），完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

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

        if mode == "NDRF 月份對比" and st.session_state.ndrf_comp_processed:
            comparison_by_shop = st.session_state.ndrf_comp_comparison_by_shop
            adj_sst_pivot = st.session_state.ndrf_comp_adj_sst_pivot
            mch2_pivot = st.session_state.ndrf_comp_mch2_pivot
            raw_df = st.session_state.ndrf_comp_raw_df
            excel_data = st.session_state.ndrf_comp_excel_data

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📊 NDRF 月份對比分析摘要")
            _top_download_button(excel_data, "I14 NDRF 月份對比", "dl_top_ndrf")

            _total_row = comparison_by_shop[comparison_by_shop["Site Code"] == "總計"]
            tr = _total_row.iloc[0] if not _total_row.empty else None

            _cols = list(comparison_by_shop.columns)
            _carry_cols = [c for c in _cols if c.startswith("Carry (")]
            _rf_cols = [c for c in _cols if c.startswith("RF (")]
            _nd_cols = [c for c in _cols if c.startswith("ND With Stock (")]

            shop_count = len(comparison_by_shop[comparison_by_shop["Site Code"] != "總計"])
            total_carry_a = int(tr[_carry_cols[0]]) if tr is not None and _carry_cols else 0
            total_carry_b = int(tr[_carry_cols[1]]) if tr is not None and len(_carry_cols) > 1 else 0
            total_rf_a = int(tr[_rf_cols[0]]) if tr is not None and _rf_cols else 0
            total_rf_b = int(tr[_rf_cols[1]]) if tr is not None and len(_rf_cols) > 1 else 0
            total_nd_a = int(tr[_nd_cols[0]]) if tr is not None and _nd_cols else 0
            total_nd_b = int(tr[_nd_cols[1]]) if tr is not None and len(_nd_cols) > 1 else 0

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">影響店舖數</div>
                    <div class="metric-value" style="color: #10b981;">{shop_count:,}</div>
                    <div class="metric-desc">個店舖含有 Carry SKU</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                _dc = total_carry_b - total_carry_a
                _sc = "+" if _dc >= 0 else ""
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3B82F6;">
                    <div class="metric-title">Carry SKU 變化</div>
                    <div class="metric-value" style="color: #3B82F6;">{total_carry_a:,} → {total_carry_b:,}</div>
                    <div class="metric-desc" style="color: {'#ef4444' if _dc > 0 else '#10b981'};">變化: {_sc}{_dc:,}</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                _dc = total_rf_b - total_rf_a
                _sc = "+" if _dc >= 0 else ""
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #8B5CF6;">
                    <div class="metric-title">RF SKU 變化</div>
                    <div class="metric-value" style="color: #8B5CF6;">{total_rf_a:,} → {total_rf_b:,}</div>
                    <div class="metric-desc" style="color: {'#ef4444' if _dc > 0 else '#10b981'};">變化: {_sc}{_dc:,}</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                _dc = total_nd_b - total_nd_a
                _sc = "+" if _dc >= 0 else ""
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #F59E0B;">
                    <div class="metric-title">ND With Stock 變化</div>
                    <div class="metric-value" style="color: #F59E0B;">{total_nd_a:,} → {total_nd_b:,}</div>
                    <div class="metric-desc" style="color: {'#ef4444' if _dc > 0 else '#10b981'};">變化: {_sc}{_dc:,}</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            st.markdown("#### 店舖級月份對比總覽")
            st.dataframe(comparison_by_shop, width='stretch', hide_index=True)

            tab1, tab2, tab3, tab4 = st.tabs(["月份對比 (By Shop)", "Adj. SST 樞紐對比", "Adj. MCH1 樞紐對比", "RAW (前50筆)"])
            with tab1:
                st.dataframe(comparison_by_shop, width='stretch', hide_index=True)
            with tab2:
                if adj_sst_pivot is not None and not adj_sst_pivot.empty:
                    st.dataframe(adj_sst_pivot, width='stretch', hide_index=True)
                else:
                    st.info("無 Adj. SST 樞紐資料")
            with tab3:
                if mch2_pivot is not None and not mch2_pivot.empty:
                    st.dataframe(mch2_pivot, width='stretch', hide_index=True)
                else:
                    st.info("無 Adj. MCH1 樞紐資料")
            with tab4:
                st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                st.caption(f"僅顯示前 50 筆 RAW 明細，完整 {len(raw_df):,} 筆將包含在下載的 Excel 報表中。")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 NDRF 月份對比_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 NDRF 月份對比分析報表 (.xlsx)",
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

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📈 Carry SKU 分析摘要 (Carry SKU Analysis Dashboard)")
            _top_download_button(excel_data, "I14 Carry SKU Review", "dl_top_carry")

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

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

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
                st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

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

        if mode == "Past 90 Days 0 Sales" and st.session_state.p90ds_processed:
            summary_df = st.session_state.p90ds_summary_df
            raw_df = st.session_state.p90ds_raw_df
            excel_data = st.session_state.p90ds_excel_data

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📉 Past 90 Days 0 Sales 分析摘要")
            _top_download_button(excel_data, "I14 Past 90 Days 0 Sales", "dl_top_p90d")

            _total_row = summary_df[summary_df["Site Code"] == "總計"] if summary_df is not None and not summary_df.empty else None
            if _total_row is not None and not _total_row.empty:
                _tr = _total_row.iloc[0]
                total_skus = int(_tr.get("Past 90 Days 0 Sales SKU 數量", 0))
                total_stock_lm = float(_tr.get("Shop Stock Qty (LM) 總數", 0))
                total_stock_cm = float(_tr.get("Shop Stock Qty (CM) 總數", 0))
            else:
                total_skus = int(raw_df["SKU"].dropna().nunique()) if raw_df is not None and not raw_df.empty else 0
                total_stock_lm = 0.0
                total_stock_cm = 0.0
            total_rows = len(raw_df) if raw_df is not None else 0

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ef4444;">
                    <div class="metric-title">符合行數</div>
                    <div class="metric-value" style="color: #ef4444;">{total_rows:,}</div>
                    <div class="metric-desc">Sold Qty 三期(M-2/M-1/MTD)皆≤0 且庫存LM/CM>0</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">SKU 數（已刪除重複）</div>
                    <div class="metric-value" style="color: #f59e0b;">{total_skus:,}</div>
                    <div class="metric-desc">符合條件的 SKU 總數（已刪除重複）</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">店舖庫存 (LM)</div>
                    <div class="metric-value" style="color: #3b82f6;">{total_stock_lm:,.0f}</div>
                    <div class="metric-desc">Shop Stock Qty (LM) 總數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">店舖庫存 (CM)</div>
                    <div class="metric-value" style="color: #10b981;">{total_stock_cm:,.0f}</div>
                    <div class="metric-desc">Shop Stock Qty (CM) 總數</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            tab_raw, tab_shop = st.tabs(["RAW", "Summary by Shop"])
            with tab_raw:
                st.markdown("#### RAW 原始明細（前 50 筆）")
                if raw_df is not None and not raw_df.empty:
                    st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                    st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")
                else:
                    st.info("無符合資料")
            with tab_shop:
                st.markdown("#### 各店舖 Past 90 Days 0 Sales 彙總")
                if summary_df is not None and not summary_df.empty:
                    st.dataframe(summary_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 Past 90 Days 0 Sales_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 Past 90 Days 0 Sales 分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "Stock Checking" and st.session_state.stock_chk_processed:
            raw_df = st.session_state.stock_chk_raw_df
            summary_df = st.session_state.stock_chk_summary_df
            stats = st.session_state.stock_chk_stats
            excel_data = st.session_state.stock_chk_excel_data

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 🔄 Stock Checking 調貨分析摘要 (Transfer Analysis Dashboard)")
            _top_download_button(excel_data, "I14 Stock Checking", "dl_top_stockchk")
            st.caption("SKU 級判定：Tourist 加總 TurnDay < 門檻 且 D001 < 門檻 且 Mix / NT 至少一個 Nature 加總 TurnDay ≥ 門檻（有貨可調）→ 標示需要調貨；調入需求補到 14 日周轉，來源優先 Non Tourist 店舖，再由 Mix 店舖補足。")

            total_rows = len(raw_df) if raw_df is not None else 0
            need_rows = int(stats.get("need_rows", 0)) if stats else 0
            need_skus = int(stats.get("need_skus", 0)) if stats else 0
            total_demand = int(stats.get("total_demand", 0)) if stats else 0
            total_actual = int(stats.get("total_actual", 0)) if stats else 0
            total_gap = int(stats.get("total_gap", 0)) if stats else 0
            nt_transfer = int(stats.get("nt_transfer", 0)) if stats else 0
            mix_transfer = int(stats.get("mix_transfer", 0)) if stats else 0

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ef4444;">
                    <div class="metric-title">需要調貨列數</div>
                    <div class="metric-value" style="color: #ef4444;">{need_rows:,}</div>
                    <div class="metric-desc">Tourist 加總 TurnDay &amp; D001 皆低於門檻的 SKU</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">需調貨 SKU 數</div>
                    <div class="metric-value" style="color: #f59e0b;">{need_skus:,}</div>
                    <div class="metric-desc">符合條件的 SKU 總數（已刪除重複）</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">調入需求量</div>
                    <div class="metric-value" style="color: #3b82f6;">{total_demand:,}</div>
                    <div class="metric-desc">補到 14 日周轉的總需求（件）</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">實際調入量</div>
                    <div class="metric-value" style="color: #10b981;">{total_actual:,}</div>
                    <div class="metric-desc">NT 優先、Mix 補足後的總調入（件）</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)

            r2_col1, r2_col2, r2_col3, r2_col4 = st.columns(4)
            with r2_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #8b5cf6;">
                    <div class="metric-title">Non Tourist 轉出量</div>
                    <div class="metric-value" style="color: #8b5cf6;">{nt_transfer:,}</div>
                    <div class="metric-desc">優先由 Non Tourist 店舖轉出（件）</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #06b6d4;">
                    <div class="metric-title">Mix 轉出量</div>
                    <div class="metric-value" style="color: #06b6d4;">{mix_transfer:,}</div>
                    <div class="metric-desc">NT 不足時由 Mix 店舖補足（件）</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f97316;">
                    <div class="metric-title">無法足額缺口</div>
                    <div class="metric-value" style="color: #f97316;">{total_gap:,}</div>
                    <div class="metric-desc">來源可轉出量不足的缺口（件）</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #64748b;">
                    <div class="metric-title">總明細行數</div>
                    <div class="metric-value" style="color: #64748b;">{total_rows:,}</div>
                    <div class="metric-desc">RAW 完整明細總行數</div>
                </div>
                """, unsafe_allow_html=True)

            buyer_stats = (stats or {}).get("buyer_need_skus") or []
            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
            st.markdown("#### 👤 Buyer 需調貨 SKU 統計")
            if buyer_stats:
                _b_total = sum(c for _, c in buyer_stats)
                _b_df = pd.DataFrame(buyer_stats, columns=["Buyer", "需調貨 SKU 數"])
                _b_df["佔比"] = (_b_df["需調貨 SKU 數"] / _b_total * 100) if _b_total else 0.0
                _b_df = _b_df.sort_values(["需調貨 SKU 數", "Buyer"], ascending=[False, True], kind="stable")
                st.dataframe(
                    _b_df, width='stretch', hide_index=True,
                    column_config={"佔比": st.column_config.NumberColumn("佔比", format="%.1f%%")},
                )
                st.caption(f"共 {len(buyer_stats)} 個 Buyer 涉及 {_b_total:,} 個需調貨 SKU（已刪除重複；Buyer 取同 SKU 首個非空值，空白歸入「(未填 Buyer)」）。")
            else:
                st.info("無符合調貨條件的 Buyer 統計")

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            tab_sum, tab_issue, tab_raw = st.tabs(["📊 Summary (SKU 彙總)", "🏪 店舖問題 SKU", "RAW 完整明細"])
            with tab_sum:
                _sc_applied_dim = (stats or {}).get("dimension") or "nature"
                _sc_dim_txt = _HYBRID_DIMENSION_LABELS.get(_sc_applied_dim, "Shop Nature")
                st.markdown(f"#### SKU 級 Summary（按 {_sc_dim_txt} 分組）：各組庫存/月銷加總與 TurnDay、涉及店舖數、佔比 (%)、調貨提示與調出量")
                st.markdown("即為 Excel 導出檔中「Summary」工作表之即時數據（加總/TurnDay 取整，TurnDay = 四捨五入到整數日；佔比 (%) = 該分組加總 ÷ 全部分組合計，頁面顯示小數、Excel 顯示為 %；調貨計算不隨維度改變）：")
                if summary_df is not None and not summary_df.empty:
                    st.dataframe(_style_inventory_skew(summary_df, dimension=_sc_applied_dim), width='stretch', hide_index=True)
                    st.caption(f"共 {len(summary_df):,} 個 SKU（按 Brand / SKU 排序）；存貨偏重提示 = {_hybrid_skew_hint_text(_sc_applied_dim)}；偏重數值以橘色底色加強顯示；存貨偏低（{_hybrid_shortage_hint_text(_sc_applied_dim)}）以淺藍底標示庫存佔比欄（Excel 同步）。")
                else:
                    st.info("無符合資料")
            with tab_issue:
                st.markdown("#### 每間店舖需行動 SKU（問題＋建議方法）")
                st.markdown("只收錄需行動列（需要調入／可轉出供應），按店舖分組、店內按嚴重度排序：")
                try:
                    _sc_issues = _build_stock_checking_shop_issues(
                        raw_df,
                        turnday_threshold=int((stats or {}).get("turnday_threshold", 14)),
                        d001_threshold=int((stats or {}).get("d001_threshold", 100)),
                    )
                except Exception:
                    _sc_issues = None
                if _sc_issues is not None and not _sc_issues.empty:
                    _sc_cols = [c for c in _STOCK_CHECKING_SHOP_ISSUE_COLUMNS if c in _sc_issues.columns]
                    st.dataframe(_sc_issues[_sc_cols], width='stretch', hide_index=True)
                    st.caption(f"共 {len(_sc_issues):,} 列需行動（只收需行動列，正常列不佔版面）。")
                else:
                    st.info("無符合資料")
            with tab_raw:
                st.markdown("#### RAW 原始資料（調貨計算欄位已附於每一列；不顯示 Item_Shop 與 Value 金額欄）")
                _display_cols = [c for c in _stock_checking_raw_display_columns(raw_df) if c in raw_df.columns]
                _display = raw_df[_display_cols]
                st.dataframe(_display.head(50).rename(columns=_display_safety_header), width='stretch')
                st.caption(f"僅顯示前 50 筆，完整 {total_rows:,} 筆資料（已按可讀性重排欄序）將包含在下載的 Excel 報表中。")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"By SKUxShop_Stock Checking_Output_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 Stock Checking 分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "Forecast Review" and st.session_state.fc_processed:
            main_df = st.session_state.fc_main_df
            coverage_df = st.session_state.fc_coverage_df
            buyer_df = st.session_state.fc_buyer_df
            raw_df = st.session_state.fc_raw_df
            stats = st.session_state.fc_stats or {}
            excel_data = st.session_state.fc_excel_data
            total_rows = int(stats.get("raw_rows", len(raw_df) if raw_df is not None else 0))

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📈 Forecast 數據分析摘要 (Forecast Dashboard)")
            _top_download_button(excel_data, "I14 Forecast Review", "dl_top_fc")

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">SKU 數</div>
                    <div class="metric-value" style="color: #3b82f6;">{int(stats.get("total_skus", 0)):,}</div>
                    <div class="metric-desc">本次計算的 SKU 總數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">爆升 SKU 數</div>
                    <div class="metric-value" style="color: #f59e0b;">{int(stats.get("surge_skus", 0)):,}</div>
                    <div class="metric-desc">新起量／加速中／爆升合計</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ef4444;">
                    <div class="metric-title">建議總訂量</div>
                    <div class="metric-value" style="color: #ef4444;">{int(stats.get("total_order", 0)):,}</div>
                    <div class="metric-desc">全部 SKU 建議訂量合計（件）</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #8b5cf6;">
                    <div class="metric-title">其中 D001 PO</div>
                    <div class="metric-value" style="color: #8b5cf6;">{int(stats.get("total_order_d001", 0)):,}</div>
                    <div class="metric-desc">Warehouse(D001) 貨向 Vendor 下單（件）</div>
                </div>
                """, unsafe_allow_html=True)
            r2_col1, r2_col2, r2_col3, r2_col4 = st.columns(4)
            with r2_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #06b6d4;">
                    <div class="metric-title">其中 Shop PO</div>
                    <div class="metric-value" style="color: #06b6d4;">{int(stats.get("total_order_shop", 0)):,}</div>
                    <div class="metric-desc">Vendor 貨向供應商下單（件）</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f97316;">
                    <div class="metric-title">已封頂 SKU 數</div>
                    <div class="metric-value" style="color: #f97316;">{int(stats.get("capped_skus", 0)):,}</div>
                    <div class="metric-desc">缺口超過上限空間的 SKU</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #64748b;">
                    <div class="metric-title">ND 不訂 SKU 數</div>
                    <div class="metric-value" style="color: #64748b;">{int(stats.get("nd_skus", 0)):,}</div>
                    <div class="metric-desc">全 ND 列的 SKU（建議訂量 0）</div>
                </div>
                """, unsafe_allow_html=True)
            with r2_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">店舖覆蓋列數</div>
                    <div class="metric-value" style="color: #10b981;">{int(stats.get("coverage_rows", 0)):,}</div>
                    <div class="metric-desc">有訂量 SKU 的店舖列數</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
            st.markdown("#### 👤 按 Buyer 建議訂量統計")
            if buyer_df is not None and not buyer_df.empty:
                _b_total_d001 = int(buyer_df["FC 建議訂量合計 (D001 PO)"].sum())
                _b_total_shop = int(buyer_df["FC 建議訂量合計 (Shop PO)"].sum())
                st.dataframe(buyer_df, width='stretch', hide_index=True)
                st.caption(f"共 {len(buyer_df)} 個 Buyer；D001 PO 合計 {_b_total_d001:,} 件，Shop PO 合計 {_b_total_shop:,} 件（全母體計算，不受「只顯示爆升」影響）。")
            else:
                st.info("無符合資料")

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            _surge_classes = {_FC_CLASS_NEW, _FC_CLASS_ACCEL, _FC_CLASS_SURGE}
            _surge_only = bool(stats.get("surge_only", True))
            _main_disp = main_df
            _cov_disp = coverage_df
            if _surge_only and main_df is not None and not main_df.empty:
                _main_disp = main_df[main_df["FC 爆升分類"].isin(_surge_classes)]
                if _cov_disp is not None and not _cov_disp.empty and not _main_disp.empty:
                    _cov_disp = _cov_disp[_cov_disp["SKU"].isin(set(_main_disp["SKU"]))]
                elif _main_disp.empty:
                    _cov_disp = _cov_disp.iloc[0:0] if _cov_disp is not None else _cov_disp
            tab_main, tab_cov, tab_raw = st.tabs(["📊 Forecast 主表 (SKU 建議訂貨)", "🏪 店舖覆蓋", "RAW 完整明細"])
            with tab_main:
                st.markdown("#### SKU 級 Forecast 主表（一 SKU 一列）")
                st.markdown("即為 Excel 導出檔中「Forecast 主表」工作表之即時數據（建議訂量＝min(缺口, 上限空間)；Warehouse(D001) 貨由 D001 向 Vendor 下單補貨（D001 PO），Vendor 貨由店舖向供應商下單（Shop PO）。「FC 本月 (M)」＝MTD 調整月化；M+1／M+2／M+3 預估僅供參考，建議訂量仍按「FC 月預估 (採用)」計算）：")
                if _main_disp is not None and not _main_disp.empty:
                    st.dataframe(_main_disp, width='stretch', hide_index=True)
                    _scope = "爆升 SKU（新起量／加速中／爆升）" if _surge_only else "全部 SKU"
                    st.caption(f"共 {len(_main_disp):,} 個 SKU（{_scope}；按 Buyer／Brand／SKU 排序）。")
                else:
                    st.info("無符合資料")
            with tab_cov:
                st.markdown("#### 店舖覆蓋（有訂量 SKU 的每店可用日數與風險旗）")
                st.markdown("先跟進「已缺貨」與「≤7日」列；訂貨到店前的空窗期以此表安排：")
                if _cov_disp is not None and not _cov_disp.empty:
                    st.dataframe(_cov_disp.head(200), width='stretch', hide_index=True)
                    st.caption(f"僅顯示前 200 筆，完整 {len(_cov_disp):,} 筆資料將包含在下載的 Excel 報表中。")
                else:
                    st.info("無符合資料（無建議訂量＞0 的 SKU）")
            with tab_raw:
                st.markdown("#### RAW 原始資料（逐列訊號欄已附於每一列；不顯示 Item_Shop 與 Value 金額欄）")
                _fc_display_cols = [c for c in _forecast_raw_display_columns(raw_df) if c in raw_df.columns]
                _display = raw_df[_fc_display_cols]
                st.dataframe(_display.head(50).rename(columns=_display_safety_header), width='stretch')
                st.caption(f"僅顯示前 50 筆，完整 {total_rows:,} 筆資料（已按可讀性重排欄序）將包含在下載的 Excel 報表中。")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"By SKUxShop_Forecast_Output_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 Forecast 分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "銷售損失量化 Review" and st.session_state.loss_processed:
            raw_df = st.session_state.loss_raw_df
            by_sku_df = st.session_state.loss_by_sku
            by_shop_df = st.session_state.loss_by_shop
            excel_data = st.session_state.loss_excel_data

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 💸 銷售損失量化分析摘要")

            total_rows = len(raw_df) if raw_df is not None else 0
            if by_sku_df is not None and not by_sku_df.empty:
                _data = by_sku_df[by_sku_df["Buyer"] != "總計"]
                total_skus = int(_data["SKU"].dropna().nunique()) if not _data.empty else 0
                total_m2 = int(pd.to_numeric(_data["Total Sales Loss (M-2) 合計"], errors="coerce").fillna(0).sum()) if not _data.empty else 0
                total_p90d = int(pd.to_numeric(_data["Total Sales Loss (P90D) 合計"], errors="coerce").fillna(0).sum()) if not _data.empty else 0
            else:
                total_skus = int(raw_df["SKU"].dropna().nunique()) if raw_df is not None and not raw_df.empty else 0
                total_m2 = 0
                total_p90d = 0

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ef4444;">
                    <div class="metric-title">受影響行數</div>
                    <div class="metric-value" style="color: #ef4444;">{total_rows:,}</div>
                    <div class="metric-desc">缺貨影響或損失大於 0 的列數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">受影響 SKU 數</div>
                    <div class="metric-value" style="color: #f59e0b;">{total_skus:,}</div>
                    <div class="metric-desc">已刪除重複</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">總損失 (M-2 基準)</div>
                    <div class="metric-value" style="color: #3b82f6;">{total_m2:,}</div>
                    <div class="metric-desc">Total Sales Loss (M-2) 合計（件）</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">總損失 (P90D 基準)</div>
                    <div class="metric-value" style="color: #10b981;">{total_p90d:,}</div>
                    <div class="metric-desc">Total Sales Loss (P90D) 合計（件）</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            tab_sku, tab_shop, tab_raw = st.tabs(["📊 By SKU（補貨優先）", "🏪 By Shop", "RAW"])
            with tab_sku:
                st.markdown("#### 按 SKU 彙總（按 P90D 損失由大至小排序）")
                if by_sku_df is not None and not by_sku_df.empty:
                    st.dataframe(by_sku_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_shop:
                st.markdown("#### 按店舖彙總（按 P90D 損失由大至小排序）")
                if by_shop_df is not None and not by_shop_df.empty:
                    st.dataframe(by_shop_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_raw:
                st.markdown("#### RAW 原始明細（前 50 筆）")
                if raw_df is not None and not raw_df.empty:
                    st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                    st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")
                else:
                    st.info("無符合資料")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 銷售損失量化 Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載銷售損失量化分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "積壓資金 Review" and st.session_state.exc_processed:
            raw_df = st.session_state.exc_raw_df
            by_sku_df = st.session_state.exc_by_sku
            by_shop_df = st.session_state.exc_by_shop
            excel_data = st.session_state.exc_excel_data

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 📦 積壓資金分析摘要")

            total_rows = len(raw_df) if raw_df is not None else 0
            if by_sku_df is not None and not by_sku_df.empty:
                _data = by_sku_df[by_sku_df["Buyer"] != "總計"]
                total_skus = int(_data["SKU"].dropna().nunique()) if not _data.empty else 0
                total_qty = int(pd.to_numeric(_data["Excess Stock Qty 合計"], errors="coerce").fillna(0).sum()) if not _data.empty else 0
                total_val = int(pd.to_numeric(_data["Excess Stock Value 合計"], errors="coerce").fillna(0).sum()) if not _data.empty else 0
            else:
                total_skus = int(raw_df["SKU"].dropna().nunique()) if raw_df is not None and not raw_df.empty else 0
                total_qty = 0
                total_val = 0

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ef4444;">
                    <div class="metric-title">積壓行數</div>
                    <div class="metric-value" style="color: #ef4444;">{total_rows:,}</div>
                    <div class="metric-desc">超額存貨大於 0 的列數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">積壓 SKU 數</div>
                    <div class="metric-value" style="color: #f59e0b;">{total_skus:,}</div>
                    <div class="metric-desc">已刪除重複</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">超額存貨總量</div>
                    <div class="metric-value" style="color: #3b82f6;">{total_qty:,}</div>
                    <div class="metric-desc">Excess Stock Qty 合計（件）</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">積壓資金總值</div>
                    <div class="metric-value" style="color: #10b981;">{total_val:,}</div>
                    <div class="metric-desc">Excess Stock Value 合計</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            tab_sku, tab_shop, tab_raw = st.tabs(["📊 By SKU（去庫存優先）", "🏪 By Shop", "RAW"])
            with tab_sku:
                st.markdown("#### 按 SKU 彙總（按積壓資金由大至小排序）")
                if by_sku_df is not None and not by_sku_df.empty:
                    st.dataframe(by_sku_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_shop:
                st.markdown("#### 按店舖彙總（按積壓資金由大至小排序）")
                if by_shop_df is not None and not by_shop_df.empty:
                    st.dataframe(by_shop_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_raw:
                st.markdown("#### RAW 原始明細（前 50 筆）")
                if raw_df is not None and not raw_df.empty:
                    st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                    st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")
                else:
                    st.info("無符合資料")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 積壓資金 Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載積壓資金分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "RP Type 覆核 Review" and st.session_state.rp_processed:
            raw_df = st.session_state.rp_raw_df
            by_sku_df = st.session_state.rp_by_sku
            by_shop_df = st.session_state.rp_by_shop
            excel_data = st.session_state.rp_excel_data

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 🏷️ RP Type 覆核分析摘要")

            total_rows = len(raw_df) if raw_df is not None else 0
            if raw_df is not None and not raw_df.empty:
                _rec = raw_df["覆核建議"].fillna("").astype(str)
                to_rf = int((_rec == "疑似應轉 RF").sum())
                to_nd = int((_rec == "疑似應轉 ND").sum())
                nd_oos = int((_rec == "ND 缺貨跟進").sum())
            else:
                to_rf = 0
                to_nd = 0
                nd_oos = 0
            if by_sku_df is not None and not by_sku_df.empty:
                _data = by_sku_df[by_sku_df["Buyer"] != "總計"]
                total_skus = int(_data["SKU"].dropna().nunique()) if not _data.empty else 0
            else:
                total_skus = 0

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ef4444;">
                    <div class="metric-title">覆核行數</div>
                    <div class="metric-value" style="color: #ef4444;">{total_rows:,}</div>
                    <div class="metric-desc">疑似誤設列數（已刪除重複計算前）</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">涉及 SKU 數</div>
                    <div class="metric-value" style="color: #f59e0b;">{total_skus:,}</div>
                    <div class="metric-desc">已刪除重複</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">疑似應轉 RF</div>
                    <div class="metric-value" style="color: #3b82f6;">{to_rf:,}</div>
                    <div class="metric-desc">ND 但近期有銷量（列數）</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">疑似應轉 ND</div>
                    <div class="metric-value" style="color: #10b981;">{to_nd:,}</div>
                    <div class="metric-desc">RF 但長期零銷售且有庫存（列數）</div>
                </div>
                """, unsafe_allow_html=True)

            if nd_oos:
                st.caption(f"另有 {nd_oos:,} 列屬「ND 缺貨跟進」（不可補貨且已缺貨），見 By SKU／By Shop 明細。")

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            tab_sku, tab_shop, tab_raw = st.tabs(["📊 By SKU", "🏪 By Shop", "RAW"])
            with tab_sku:
                st.markdown("#### 按 SKU 彙總（按涉及店舖數由大至小排序）")
                if by_sku_df is not None and not by_sku_df.empty:
                    st.dataframe(by_sku_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_shop:
                st.markdown("#### 按店舖彙總（按涉及 SKU 數由大至小排序）")
                if by_shop_df is not None and not by_shop_df.empty:
                    st.dataframe(by_shop_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_raw:
                st.markdown("#### RAW 原始明細（前 50 筆）")
                if raw_df is not None and not raw_df.empty:
                    st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                    st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")
                else:
                    st.info("無符合資料")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 RP Type 覆核 Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 RP Type 覆核分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "D001 健康檢查 Review" and st.session_state.d001_processed:
            check_df = st.session_state.d001_check_df
            by_buyer_df = st.session_state.d001_by_buyer
            raw_df = st.session_state.d001_raw_df
            excel_data = st.session_state.d001_excel_data

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 🏭 D001 健康檢查分析摘要")

            if check_df is not None and not check_df.empty:
                _rec = check_df["檢查結論"].fillna("").astype(str)
                total_skus = int(check_df["SKU"].dropna().nunique())
                shortage = int((_rec == "D001 缺貨需補").sum())
                overstock = int((_rec == "倉庫積壓").sum())
                low_cover = int((_rec == "覆蓋偏低").sum())
            else:
                total_skus = 0
                shortage = 0
                overstock = 0
                low_cover = 0

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ef4444;">
                    <div class="metric-title">需跟進 SKU 數</div>
                    <div class="metric-value" style="color: #ef4444;">{total_skus:,}</div>
                    <div class="metric-desc">已刪除重複</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">D001 缺貨需補</div>
                    <div class="metric-value" style="color: #f59e0b;">{shortage:,}</div>
                    <div class="metric-desc">倉庫無貨但店舖有需求</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">倉庫積壓</div>
                    <div class="metric-value" style="color: #3b82f6;">{overstock:,}</div>
                    <div class="metric-desc">倉庫有貨但全渠道零銷售</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">覆蓋偏低</div>
                    <div class="metric-value" style="color: #10b981;">{low_cover:,}</div>
                    <div class="metric-desc">倉庫加在途不足一個月銷量</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            tab_check, tab_buyer, tab_raw = st.tabs(["📊 D001 檢查（SKU 級）", "👤 By Buyer", "RAW"])
            with tab_check:
                st.markdown("#### SKU 級檢查表（按嚴重度與覆蓋日數排序）")
                if check_df is not None and not check_df.empty:
                    st.dataframe(check_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_buyer:
                st.markdown("#### 按採購人員彙總")
                if by_buyer_df is not None and not by_buyer_df.empty:
                    st.dataframe(by_buyer_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_raw:
                st.markdown("#### RAW 原始明細（前 50 筆，需跟進 SKU 的店舖列）")
                if raw_df is not None and not raw_df.empty:
                    st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                    st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")
                else:
                    st.info("無符合資料")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 D001 健康檢查 Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載 D001 健康檢查報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

        if mode == "新品追蹤 Review" and st.session_state.ns_processed:
            raw_df = st.session_state.ns_raw_df
            by_sku_df = st.session_state.ns_by_sku
            by_shop_df = st.session_state.ns_by_shop
            excel_data = st.session_state.ns_excel_data
            _ns_ref_txt = st.session_state.get("ns_ref_date", "")

            st.markdown('<div class="custom-card">', unsafe_allow_html=True)
            st.markdown("### 🌱 新品追蹤分析摘要")
            if _ns_ref_txt:
                st.caption(f"參考日期 {_ns_ref_txt} 起 90 天內上市視為新品。")

            total_rows = len(raw_df) if raw_df is not None else 0
            if by_sku_df is not None and not by_sku_df.empty:
                _data = by_sku_df[by_sku_df["Buyer"] != "總計"]
                total_skus = int(_data["SKU"].dropna().nunique()) if not _data.empty else 0
            else:
                total_skus = 0
            if raw_df is not None and not raw_df.empty:
                _st = raw_df["起量狀態"].fillna("").astype(str)
                started_rows = int((_st == "已起量").sum())
                not_started_rows = int((_st == "未起量").sum())
            else:
                started_rows = 0
                not_started_rows = 0

            r1_col1, r1_col2, r1_col3, r1_col4 = st.columns(4)
            with r1_col1:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #10b981;">
                    <div class="metric-title">新品行數</div>
                    <div class="metric-value" style="color: #10b981;">{total_rows:,}</div>
                    <div class="metric-desc">上市 90 天內的店舖列數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col2:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #3b82f6;">
                    <div class="metric-title">新品 SKU 數</div>
                    <div class="metric-value" style="color: #3b82f6;">{total_skus:,}</div>
                    <div class="metric-desc">已刪除重複</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col3:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #f59e0b;">
                    <div class="metric-title">已起量行數</div>
                    <div class="metric-value" style="color: #f59e0b;">{started_rows:,}</div>
                    <div class="metric-desc">任一期有銷售的列數</div>
                </div>
                """, unsafe_allow_html=True)
            with r1_col4:
                st.markdown(f"""
                <div class="metric-card" style="border-left-color: #ef4444;">
                    <div class="metric-title">未起量行數</div>
                    <div class="metric-value" style="color: #ef4444;">{not_started_rows:,}</div>
                    <div class="metric-desc">四期皆零銷售的列數</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown('</div>', unsafe_allow_html=True)

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            tab_sku, tab_shop, tab_raw = st.tabs(["📊 By SKU", "🏪 By Shop", "RAW"])
            with tab_sku:
                st.markdown("#### 按 SKU 彙總（按上市日期由新至舊排序）")
                if by_sku_df is not None and not by_sku_df.empty:
                    st.dataframe(by_sku_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_shop:
                st.markdown("#### 按店舖彙總（按新品 SKU 數由大至小排序）")
                if by_shop_df is not None and not by_shop_df.empty:
                    st.dataframe(by_shop_df, width='stretch', hide_index=True)
                else:
                    st.info("無符合資料")
            with tab_raw:
                st.markdown("#### RAW 原始明細（前 50 筆）")
                if raw_df is not None and not raw_df.empty:
                    st.dataframe(raw_df.head(50).rename(columns=_display_safety_header), width='stretch')
                    st.caption(f"僅顯示前 50 筆，完整 {len(raw_df):,} 筆資料將包含在下載的 Excel 報表中。")
                else:
                    st.info("無符合資料")

            st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)

            hk_time = datetime.now(ZoneInfo("Asia/Hong_Kong")).strftime("%Y%m%d_%H%M")
            download_name = f"I14 新品追蹤 Review_{hk_time}.xlsx"

            st.download_button(
                label="📥 下載新品追蹤分析報表 (.xlsx)",
                data=excel_data,
                file_name=download_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width='stretch',
            )

# 8. 底部頁尾資訊
st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
st.markdown(
    '<div style="text-align: center; color: #444455; font-size: 0.8rem; padding: 10px 0;">'
    'I14 報表轉換與分析系統 &nbsp;&nbsp;•&nbsp;&nbsp; Streamlit + Zeabur &nbsp;&nbsp;•&nbsp;&nbsp; '
    f'v{__version__} &nbsp;&nbsp;•&nbsp;&nbsp; Developed by Ricky Yue'
    '</div>',
    unsafe_allow_html=True,
)




