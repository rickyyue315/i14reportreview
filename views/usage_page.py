# -*- coding: utf-8 -*-
import streamlit as st
from version import __version__

# 1. 頁首橫幅
st.markdown("""
<div class="header-container">
    <h1 class="header-title">📖 使用說明</h1>
    <p class="header-subtitle">於「分析工具」頁頂端選擇分析模式後，系統將自動切換對應的計算流程與輸出。本頁彙整各模式的用途說明與核心計算規則。</p>
</div>
""", unsafe_allow_html=True)

# 2. 系統操作流程（單一 HTML 區塊，避免跨 st.markdown 的 div 被隔離而產生空卡片）
st.markdown("""
<div class="custom-card">
    <h3 style="margin-top: 0;">📋 系統操作流程</h3>
    <ol>
        <li>📥 <b>準備來源檔案</b>：確認您的 Excel 檔案（By SKUxShop）包含必要欄位。</li>
        <li>📂 <b>上傳檔案</b>：將 Excel 拖曳至「分析工具」頁的上傳區域。</li>
        <li>🎯 <b>選擇模式</b>：於主頁面頂端選擇要執行的分析模式。</li>
        <li>🔧 <b>設定參數</b>：根據模式需要，調整 MTD 日期、Shop Nature Rate 等進階選項。</li>
        <li>⚙️ <b>執行轉換</b>：點擊「執行報表轉換與分析」，系統將自動過濾、計算並產生彙總表。</li>
        <li>💾 <b>下載成果</b>：於視覺化看板中即時查閱統計資訊，並一鍵導出美化後的 Excel 報表。</li>
    </ol>
</div>
""", unsafe_allow_html=True)

# 3. 各模式說明與計算規則（每個模式一個收合區塊）
st.markdown("### 🎯 各模式說明與核心計算規則")
st.caption("點開任一模式，查看該模式的用途說明（🎯）與核心計算規則（⚙️）。")

MODES = [
    ("📦 Overstock Review",
     """
- 篩選庫存過剩品項（P90D 零銷售 / 超過月均兩倍）
- 計算新 Safety Qty 與可轉出貨量
- 輸出：RAW、Shop Stock Status、By Shop、分析摘要
     """,
     """
針對 `Shop Stock Status` 過濾以下庫存過剩條件：
1. **P90D 0 Sales** — 過去 90 天內零銷售
2. **Overstock (>2*Monthly Qty & >2pcs)** — 庫存超過月均銷售兩倍且大於 2 件
     """),

    ("🚨 OOS Review",
     """
- 篩選缺貨（0 Stock）及潛在缺貨（Turnday < 7）
- 偵測 Safety Qty 不足之正常庫存品項
- 支援 MTD 日期選擇與月度預估銷售計算
- 輸出：RAW、Shop Stock Status、By Shop、分析摘要
     """,
     """
針對 `Shop Stock Status` 過濾以下缺貨條件：
1. **OOS (0 Stock)** — 現有庫存為 0
2. **Potential OOS (<7 Turnday)** — Turnday 不足 7 天
3. **Normal (Safety Qty Insufficient)** — 庫存正常但 Safety Qty 不足以覆蓋預估銷售
     """),

    ("📋 Carry SKU Review",
     """
- 檢視各店舖的 Carry SKU 數量（扣除 ND 且庫存為 0 之誤差列）
- 按 RP Type（ND / RF）與 P90D 0 Sales 交叉統計
- 輸出：RAW、Summary（雙區塊樞紐：Adj. MCH1 × Adj. SST）、Carry SKU 分析摘要
     """,
     """
檢視各店舖實際 Carry SKU 數量：
1. 先清理無效列（刪除 RP Type 為 ND 且 Shop Stock Qty 為 0 之誤差列）
2. 依 RP Type (ND / RF) 與 P90D 0 Sales 狀態做雙維度交叉統計
3. 雙區塊樞紐：左 (Adj. MCH1) + 右 (Adj. SST)
     """),

    ("✂️ SKU 精簡 Review",
     """
- 針對 RP Type 為 ND 的 SKU 建議退倉，區分 Supply Source
- Warehouse 貨：建議全數退倉；Vendor 貨：標記「另有安排」不可退倉
- 依 Supply Source × Shop Stock Status 雙層動態分類
- 輸出：RAW、Shop Stock Status、By Shop、ND SKU 數據分析摘要
     """,
     """
針對 ND SKU 建議退倉（區分 Supply Source）：
1. 篩選 RP Type = ND 的所有 SKU（不分 Shop Stock Status）
2. **Warehouse 貨**：建議全數退倉 `transfer = Shop Stock Qty (CM)`
3. **Vendor 貨**：標記「另有安排」，不可退倉 `transfer = 0`
4. ND SKU 依 Supply Source × Shop Stock Status 雙層動態分類（如 Warehouse - Overstock、vendor - Normal）
     """),

    ("📊 NDRF 月份對比",
     """
- 對比兩個月份（如 Mar2026 vs Jul2026）各店舖的 Carry SKU、RF SKU、ND With Stock SKU 數量
- 支援按 Adj. SST 交叉分析與 SKU 級明細檢視
- 輸出：月份對比 (By Shop)、Adj. SST 樞紐表 (x2)、SKU 明細、RAW
     """,
     None),

    ("🔒 Safety Stock Review",
     """
- 全面審查各 SKU×Shop 的安全庫存設定是否合理
- 結合 Overstock Cap（庫存上限）與 OOS Floor（缺貨底線），雙向平衡調整 Safety Stock
- 支援逐 Site 檢視、Overstock / OOS Rate 自訂、Safety 只增或加減模式
- 輸出：RAW、Shop Stock Status、By Shop、SUMMARY（Safety Qty Change 分類 / ND 建議分類統計）、分析摘要
     """,
     """
全面審查 Safety Stock 設定，計算平衡建議值：
1. **Overstock Cap (X)** = floor(月均銷售 × Overstock Rate)，上限不超過現有 Adj. Safety Stock
2. **OOS Floor (Y)** = ceil(月銷預估 × OOS Rate)，最低 3；月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均（加權混合，月初以 P90D 為主避免頭幾日外推失真，月尾收斂至 MTD；P90D 欄位缺失時退回純 MTD 線性外推）
3. **建議 Safety Stock** = max(Y, Z if Z ≤ X else X)，其中 Z = 現有 Adj. Safety Stock
4. **Recommend Safety Stock Turnday** = 建議值 ÷ (P90D 月均 ÷ 30)，與來源「Turnday (P90D Qty)」同口徑（P90D 日均），可直接對比現時與建議的周轉日數；P90D 月均 = 0 時留空（無周轉意義），建議值 = 0 顯示 0（零覆蓋）
5. **Safety Qty Change** = 建議 Safety Stock − Z（正數=增加、負數=減少、零=不變），並輸出「Safety Qty Change 分類」欄（增加 / 減少 / 不變；「只可增加」模式下不會出現「減少」）
6. **New SKU 保護** — Launch Date 距今 90 天內（≈3 個月）的列排除「P90D 0 Sales → Safety=0/ND」override（新品未出單屬上架時間問題），改走正常 X/Y 計算，並在「Recommendation of new Safety Qty / ND」標記 New SKU (3 Month)；Launch Date 缺失或無法解析時視為非新品
7. **計算模式** — 「只可增加」：全表只有增加或保持原有 Safety，ND / P90D 0 Sales 的歸零+ND 退場建議不套用（ND 欄僅保留 New SKU 標記）；「可加可減」：允許減少並套用退場建議（ND 列 → 0；P90D 0 Sales 舊品 → ND 建議，並依 Adj. MCH1 / Adj. SST 加註：02 - MAKE UP → ND(02 - MAKE UP)、03 - Private Label 系 → ND(Private Label)、兩者皆是 → ND(Private Label - MAKE UP)）
8. 品項分為 Overstock / OOS / Balanced (Safety Adjusted) 三類，標示調整原因
9. **SUMMARY 工作表** — 統計 Safety Qty Change 分類（增加 / 減少 / 不變）與 Recommendation of new Safety Qty / ND 分類（空白維持 / New SKU / 0 已是 ND / ND 系列加註）各行數與 SKU 數；總計列 SKU 數為全表不重複 SKU（同一 SKU 可跨分類）
     """),

    ("📦 Stock Impact Analysis",
     """
- 以 Shop Stock Qty = 0（缺貨）為核心，分析存貨對 Sold Qty 的因果影響（無貨就無銷售）
- 缺貨模式判別：長期缺貨（≥2 期）/ 間歇缺貨（1 期）/ 恢復供貨 / 無缺貨
- 銷售損失雙基準並列：M-2 Sold 為需求基準 vs P90D 90 天月均為基準（只計 RF 店）
- 📦 **SKU 月銷量和存儲量分析**：按 SKU 計算各 Nature（Tourist/Mix/NT）月均銷售與庫存，低存貨優先
- 🎯 **存貨分佈分析**：Mix / Non Tourist Turnday 分級（7-14 日標準）、Nature 存貨對比
- ⚠️ **存貨失衡 SKU 判定**：Tourist 缺貨 + Non Tourist/Mix 過剩 → 揪出配置不當
- 🔄 **跨 Nature 調貨建議**：Tourist 缺貨但 Mix/NT 有可釋放件數 → 可調貨優先清單
- ⚖️ **低存貨 SKU 配置決策**：D001 + Shop Stock < N 件時分析 NT RF 浪費與 ND 囤積
- 🔍 **by-SKU 缺貨診斷**：含 D001 倉庫三期貨量 → 區分「倉庫有貨（補貨問題）」vs「源頭斷貨」
- 📊 **多維度彙總**：By Brand / Buyer / Category (Adj. MCH1) 缺貨影響
- 門檻只影響低存貨配置決策，不限制整體分析的 SKU 範圍
- 預設只分析 Warehouse SKU，可切換納入 Vendor
     """,
     """
缺貨銷售影響分析判別：
1. **缺貨判定** = Shop Stock Qty = 0（嚴格標準，三期 M-2 / M-1 / CM 各自判定）
2. **缺貨模式**：長期缺貨（≥2 期）/ 間歇缺貨（1 期）/ 恢復供貨（前期缺貨 → CM 有貨）/ 無缺貨
3. **銷售損失（M-2 基準）** = max(0, Sold Qty M-2 − 缺貨當期 Sold)，以 M-2 銷售為需求基準
4. **銷售損失（P90D 基準）** = max(0, Shop Monthly Sold Qty P90D − 缺貨當期 Sold)，以 90 天月均為基準
5. **影響判定** = 任一期缺貨且有需求證據（M-2 Sold 或 P90D > 0），且只計算 RP Type = RF 的店舖
6. **Shop Nature 優先級**：Tourist > Mix > Non Tourist（彙總按 Tourist 缺貨嚴重度排序）
7. **SKU 調貨建議**：Tourist 缺貨但同 SKU 在 Non-Tourist/Mix 仍有存貨 → 揪出存貨配置不當
8. 預設只分析 Warehouse SKU；支援多維度彙總：By Site / Brand / Buyer / Category (Adj. MCH1)
     """),

    ("📦 Stock Impact Analysis (簡化)",
     """
- 簡化版缺貨銷售影響分析，只輸出 RAW 和 SKU Monthly Sales&Store 兩個工作表
- SKU Monthly Sales&Store **保留原版所有欄位**（含 Vendor MC/Adj. MCH1/Item Status/Launch Date/Item Class/Is New SKU），並 **新增** M-2, M-1 的 D001 存貨、店舖總存貨及 M-2, M-1, MTD 的店舖總銷售量（不分 Shop Nature，D001_MTD/ShopStock_MTD 已存在於原版不重複）
- 適合快速查看各 SKU 三期存貨與銷售趨勢
     """,
     None),

    ("📉 Past 90 Days 0 Sales",
     """
- 篩選「過去 90 日零銷售（Sold Qty M-2 / M-1 / CM-MTD 及 Sold Qty (P90D) 皆 ≤ 0）、但店舖仍有存貨（Shop Stock Qty LM 與 CM 皆 > 0）」的 SKU×Shop 記錄
- RAW 只輸出 23 個指定欄位（Site Code / Shop Name / OM / Shop Nature / Brand / SKU / SKU Desc / Vendor MC / Item Status / Launch Date / Item Class / RP Type / Shop Stock Status / Safety Stock (現時) / Sold Qty M-2 / M-1 / CM-MTD / P90D / Shop Stock Qty CM / LM / Excess Stock Qty / Shop On Order PO Qty (CM) / Buyer）
- 輸出：RAW、Summary by Shop
     """,
     None),

    ("🔄 Stock Checking（調貨檢查）",
     """
- 觸發條件：該 SKU 的 **Tourist 加總 TurnDay**（各 Tourist 店 Shop Stock Qty (CM) 加總 ÷ Shop Monthly Sold Qty (P90D) 加總 × 30.5，與 Summary 同口徑）< 14 日（UI 可調）**且** D001 Net Stock Qty (CM) < 100 件（UI 可調）**且** Mix / Non Tourist 至少一個 Nature 的加總 TurnDay ≥ 14 日（有貨可調）→ 標示「需要調貨」；加總 TurnDay ≥ 門檻一律無須調貨
- 供應門檻：SKU 級 Mix / Non Tourist 加總 TurnDay < 14 日時，該 Nature 視為無貨可調（可轉出量歸零）；兩個 Nature 皆低於門檻時整個 SKU 無須標示
- 調入需求量 = 補到 14 日周轉：ceil(P90D 日均 × 14) − 現有店舖庫存
- Mix / Non Tourist 鋪可轉出貨量 = 超過 14 日周轉的部分（無銷售但有存貨全數可釋放，RF 與 ND 店皆可）
- 分配優先性：同一 SKU 的調貨需求先由 Non Tourist 鋪轉出，不足再由 Mix 鋪補足
- 輸出：Summary（SKU 彙總：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001、各 Nature 庫存/月銷加總（取整）、TurnDay（取整）、各 Nature 庫存/月銷佔比 (%)（0.00%）、強制調貨提示、NT/Mix 預計調出量）+ RAW（原始欄位 + 調貨標示、調入需求量、實際調入量、調入缺口、可轉出貨量、轉出量、調貨來源）+ 調貨分析摘要（Dashboard：8 項指標卡 + Buyer 需調貨 SKU 統計表）
     """,
     """
Tourist 缺貨、Mix/NT 調撥計算：
1. **觸發條件（SKU 級，同時符合才標示）** = 該 SKU 的「Tourist 加總 TurnDay」（各 Tourist 店 Shop Stock Qty (CM) 加總 ÷ Shop Monthly Sold Qty (P90D) 加總 × 30.5，與 Summary 同口徑）< 門檻（預設 14 日，UI 可調），且 D001 Net Stock Qty (CM) < 門檻（預設 100 件，UI 可調），且 Mix / Non Tourist 至少一個 Nature 的加總 TurnDay ≥ 門檻（有貨可調）；Tourist 加總 TurnDay ≥ 門檻一律無須調貨，Mix 與 Non Tourist 加總 TurnDay 皆低於門檻時視為無貨可調、不標示
2. **調入需求量** = ceil(P90D / 90 × 14) − Shop Stock Qty (CM)（P90D = Sold Qty (P90D)，90 日總量換算日均 × 14 日）；只為已標示「需要調貨」的 SKU 計算，個別 Tourist 店需求 ≤ 0 不標示
3. **可轉出貨量（Mix / Non Tourist 來源店）** = max(0, Shop Stock Qty (CM) − P90D / 90 × 14)；「無銷售但有存貨」全數可釋放；RP Type 為 RF 與 ND 的店皆可作為來源；該 Nature 的 SKU 級加總 TurnDay < 門檻時視為無貨可調（可轉出量歸零）
4. **分配優先** = 按 SKU 彙總調入需求，Non Tourist 店可轉出量優先，不足部分由 Mix 店補足；需求列按需求量大→小依序分配
5. **輸出** = Summary（SKU 級彙總：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001；各 Nature 庫存/月銷加總（取整）、TurnDay = ROUND(加總庫存 ÷ 加總月銷 × 30.5, 0)、各 Nature 庫存/月銷佔比 (%) = 該 Nature 加總 ÷ 三 Nature 合計（0.00%）、強制調貨提示、由 NT/Mix 店舖預計調出數量）+ RAW 完整明細（原始欄位全部保留），RAW 新增：需要調貨 / 調入需求量 (補到14日) / 實際調入量 / 調入缺口 / 可轉出貨量 (Mix) / 可轉出貨量 (Non Tourist) / 轉出量 / 調貨來源 (Site Code) + 調貨分析摘要（Dashboard：需要調貨列數 / 需調貨 SKU 數 / 調入需求量 / 實際調入量 / NT 轉出量 / Mix 轉出量 / 無法足額缺口 / 總明細行數 + Buyer 需調貨 SKU 統計表：各 Buyer 涉及需調貨的去重 SKU 數與佔比）
     """),

    ("🔀 Safety Stock x NetStock Review",
     """
- 同一次執行「Safety Stock Review × Stock Checking」雙引擎，並以調貨結果反饋 Safety 診斷（**不改寫** Recommend Safety Stock）
- Safety 段沿用現行公式（Overstock Cap × OOS Floor、參數全可調）；調貨段沿用現行規則（14 日目標、NT 優先、D001/TurnDay 門檻）
- 反饋欄位：調貨後庫存 = 原始 + 實際調入 − 轉出；調貨前/後 Safety 缺口 = max(0, Recommend Safety Stock − 庫存)；調貨後 TurnDay = 調貨後庫存 ÷ 月銷 × 30.5
- Shop Nature 固定全三種；Other/空白/大小寫差異一律正規化為 Non Tourist；混合必要欄位嚴格檢核
- 輸出：RAW + Safety（Shop Stock Status / By Shop / SUMMARY / 數據分析摘要）+ 調貨（Summary / 調貨分析摘要）+ Hybrid Summary 共 8 個工作表
     """,
     """
合併兩套既有引擎，調貨結果只反饋為 Safety 診斷：
1. **Safety 段** = 與 Safety Stock Review 完全同公式（X/Y、Projected、Safety 只增或加減、New SKU 保護、ND 加註），參數（MTD 日期 / Overstock×OOS Rate / Safety Qty 模式）全數開放
2. **調貨段** = 與 Stock Checking 完全同規則（SKU 級加總 TurnDay 門檻、D001 門檻、14 日需求/供應、NT 優先→Mix 補足、需求量大→小）
3. **反饋（不改寫 Safety）** = 調貨後庫存 = Shop Stock Qty (CM) + 實際調入量 − 轉出量；調貨前 Safety 缺口 = max(0, Recommend Safety Stock − 原始庫存)；調貨後 Safety 缺口 = max(0, Recommend Safety Stock − 調貨後庫存)；調貨後 TurnDay = 調貨後庫存 ÷ Shop Monthly Sold Qty (P90D) × 30.5（月銷 0 留空）
4. **Shop Nature** = 固定三種；Other/空白/大小寫差異一律正規化為 Non Tourist，映射列數列 warning；不做 Nature 篩選（避免破壞調貨供需母體）
5. **輸入檢核** = 基礎四欄 + Shop Nature / Site Code / Shop Stock Qty (CM) / Sold Qty (P90D) / Adj. Safety Stock / Shop Monthly Sold Qty (P90D) / Sold Qty (CM - MTD) 皆為混合必要欄位，缺任一即停止；D001 沿用 Net→Stock fallback，兩者皆缺視為 0+warning
6. **調貨目標固定 14 日**（沿用現行規則），不另開第三套可調目標
7. **輸出 8 工作表** = RAW / Shop Stock Status / By Shop / SUMMARY / Safety Stock 數據分析摘要 / Summary（調貨）/ 調貨分析摘要 / Hybrid Summary（調貨反饋 Safety：SKU 級調貨前後庫存、Safety 缺口前/後、改善件數、仍有缺口）
     """),
]

for _label, _desc, _rules in MODES:
    with st.expander(_label, expanded=False):
        st.markdown("#### 🎯 模式說明")
        st.markdown(_desc)
        if _rules:
            st.markdown("#### ⚙️ 核心計算規則")
            st.markdown(_rules)

# 4. 必要欄位檢核標準（單一 HTML 區塊）
st.markdown("""
<div class="custom-card">
    <h3 style="margin-top: 0;">🔍 必要欄位檢核標準</h3>
    <p>為確保計算正確，來源檔案必須包含以下欄位：</p>
    <ul>
        <li>✅ <b>Buyer</b> <span style="color:#0077a3;font-size:0.8em">(採購人員)</span></li>
        <li>✅ <b>Brand</b> <span style="color:#0077a3;font-size:0.8em">(品牌名稱)</span></li>
        <li>✅ <b>SKU</b> <span style="color:#0077a3;font-size:0.8em">(品號編碼)</span></li>
        <li>✅ <b>Shop Stock Status</b> <span style="color:#0077a3;font-size:0.8em">(庫存狀態明細)</span></li>
    </ul>
    <p>OOS Review 額外建議欄位（非必要，缺少時將顯示提示）：</p>
    <ul>
        <li>📌 <b>Sold Qty (CM - MTD)</b> <span style="color:#f59e0b;font-size:0.8em">(月度銷售)</span></li>
        <li>📌 <b>Adj. Safety Stock</b> <span style="color:#f59e0b;font-size:0.8em">(現有安全庫存)</span></li>
        <li>📌 <b>Shop Nature</b> <span style="color:#f59e0b;font-size:0.8em">(店鋪性質)</span></li>
    </ul>
</div>
""", unsafe_allow_html=True)

# 5. 頁尾
st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
st.caption(f"💡 所有模式均自動生成採購與品牌的交叉匯總樞紐分析表。｜I14 報表計算工具 v{__version__}")
