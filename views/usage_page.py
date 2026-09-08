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
        <li>🎯 <b>選擇模式</b>：於「分析工具」頁主頁面頂端選擇要執行的分析模式。</li>
        <li>📂 <b>上傳檔案</b>：將 Excel 拖曳至主頁面的上傳區域。</li>
        <li>🔧 <b>設定參數</b>：根據模式需要，調整 MTD 日期、Shop Nature Rate 等進階選項。</li>
        <li>⚙️ <b>執行轉換</b>：點擊「執行報表轉換與分析」，系統將自動過濾、計算並產生彙總表。</li>
        <li>💾 <b>下載成果</b>：於視覺化看板中即時查閱統計資訊，並一鍵導出美化後的 Excel 報表。</li>
    </ol>
</div>
""", unsafe_allow_html=True)

# 2b. 圖解操作流程（收合於展開區塊，避免進入頁面時即載入圖片拖慢顯示）
with st.expander("🖼️ 圖解操作流程（實際畫面擷圖）", expanded=False):
    st.caption("以下為實際操作畫面：先選模式，再上傳檔案。")
    st.image("assets/guide/01_mode_select_annotated.png", caption="步驟 1：於主頁面頂端選擇分析模式", width="stretch")
    st.image("assets/guide/02_upload_annotated.png", caption="步驟 2：將 By SKUxShop Excel 拖曳至上傳區域", width="stretch")

# 3. 各模式說明與計算規則（每個模式一個收合區塊）
st.markdown("### 🎯 各模式說明與核心計算規則")
st.caption("點開任一模式，查看該模式的用途說明（🎯）與核心計算規則（⚙️）。")

MODES = [
    ("📦 Overstock Review",
     """
- 篩選庫存過剩品項（P90D 零銷售 / 超過月均兩倍）
- 計算新 Safety Qty 與可轉出貨量
- 輸出：RAW、Shop Stock Status、By Shop、異常庫存數據分析摘要
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
- 輸出：RAW、Shop Stock Status、By Shop、OOS庫存數據分析摘要
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
- 輸出：RAW、Shop Stock Status、By Shop、ND SKU數據分析摘要
     """,
     """
針對 ND SKU 建議退倉（區分 Supply Source）：
1. 篩選 RP Type = ND 的所有 SKU（不分 Shop Stock Status）
2. **Warehouse 貨**：建議全數退倉 `transfer = Shop Stock Qty (CM)`
3. **Vendor 貨**：標記「另有安排」，不可退倉 `transfer = 0`
4. ND SKU 依 Supply Source × Shop Stock Status 雙層動態分類（如 Warehouse - Overstock、Vendor - Normal）
     """),

    ("📊 NDRF 月份對比",
     """
- 對比兩個月份（如 Mar2026 vs Jul2026）各店舖的 Carry SKU、RF SKU、ND With Stock SKU 數量
- 支援按 Adj. SST 與 Adj. MCH1 交叉分析（兩月份並列）
- 輸出：月份對比 (By Shop)、Adj. SST 樞紐對比、Adj. MCH1 樞紐對比、RAW
     """,
     None),

    ("🔒 Safety Stock Review",
     """
- 全面審查各 SKU×Shop 的安全庫存設定是否合理
- 結合 Overstock Cap（庫存上限）與 OOS Floor（缺貨底線），雙向平衡調整 Safety Stock
- 支援逐 Site 檢視、Overstock / OOS Rate 自訂、Safety 只增或加減模式
- 輸出：RAW、Shop Stock Status、By Shop、店舖問題 SKU、SUMMARY（Safety Qty Change 分類 / ND 建議分類統計）、Safety Stock 數據分析摘要
- 新增「🏪 店舖問題 SKU」：每間店舖需行動列（Safety 需增加／減少、ND 退場；New SKU 標記不收錄），按店舖分組、店內按嚴重度排序，逐列顯示問題說明與建議方法（數量＋行動指引）
     """,
     """
全面審查 Safety Stock 設定，計算平衡建議值：
1. **Overstock Cap (X)** = floor(月均銷售 × Overstock Rate)，上限不超過現有 Safety 設定（Safety Stock (現時)，即來源檔案的 Adj. Safety Stock）
2. **OOS Floor (Y)** = ceil(月銷預估 × OOS Rate)，最低 3；月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均（加權混合，月初以 P90D 為主避免頭幾日外推失真，月尾收斂至 MTD；P90D 欄位缺失時退回純 MTD 線性外推）
3. **建議 Safety Stock** = max(Y, Z if Z ≤ X else X)，其中 Z = 現有 Safety 設定（Safety Stock (現時)，即來源檔案的 Adj. Safety Stock）
4. **Recommend Safety Stock Turnday** = 建議值 ÷ (P90D 月均 ÷ 30)，與來源「Turnday (P90D Qty)」同口徑（P90D 日均），可直接對比現時與建議的周轉日數；P90D 月均 = 0 時留空（無周轉意義），建議值 = 0 顯示 0（零覆蓋）
5. **Safety Qty Change** = 建議 Safety Stock − Z（正數=增加、負數=減少、零=不變），並輸出「Safety Qty Change 分類」欄（增加 / 減少 / 不變；「只可增加」模式下不會出現「減少」）
6. **New SKU 保護** — Launch Date 距 MTD 截止日期 90 天內（≈3 個月）的列排除「P90D 0 Sales → Safety=0/ND」override（新品未出單屬上架時間問題），改走正常 X/Y 計算，並在「Recommendation of new Safety Qty / ND」標記 New SKU (3 Month)；Launch Date 缺失或無法解析時視為非新品
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
- 🔍 **by-SKU 缺貨 Review**：含 D001 倉庫三期貨量 → 區分「倉庫有貨（補貨問題）」vs「源頭斷貨」
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
8. 預設只分析 Warehouse SKU；支援多維度彙總：By Brand / Buyer / Category (Adj. MCH1) / Item Class
     """),

    ("📦 Stock Impact Analysis (簡化)",
     """
- 簡化版缺貨銷售影響分析，輸出 RAW、SKU Monthly Sales&Store、New、Stock Impact Dashboard 四個工作表
- SKU Monthly Sales&Store **保留原版所有欄位**（含 Vendor MC/Adj. MCH1/Item Status/Launch Date/Item Class/Is New SKU），並 **新增** M-2, M-1 的 D001 存貨、店舖總存貨及 M-2, M-1, MTD 的店舖總銷售量（不分 Shop Nature，D001_MTD/ShopStock_MTD 已存在於原版不重複）
- New 工作表為 SKU 級 SOH Tier 分佈及有賣與有貨比率；Dashboard 另加低存貨 SKU 及有賣與有貨比率指標
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
- Mix / Non Tourist 店舖可轉出貨量 = 超過 14 日周轉的部分（無銷售但有存貨全數可釋放，RF 與 ND 店皆可）
- 分配優先性：同一 SKU 的調貨需求先由 Non Tourist 店舖轉出，不足再由 Mix 店舖補足
- 存貨偏重提示：某 Nature 庫存佔比 ≥ 40% 且庫存佔比 − 銷量佔比 ≥ 20 個百分點 →「存貨偏重 {Nature}（庫存 x% / 銷量 y%）」，多個 Nature 以；分隔（差距大者在前）；偏重提示欄與偏重 Nature 的庫存佔比欄以醒目橘底加強顯示（Streamlit 與 Excel 同步）
- 新增「🏪 店舖問題 SKU」：每間店舖需行動列（需要調入／可轉出供應，缺口在同一列加註），按店舖分組、店內按嚴重度排序，逐列顯示問題說明與建議方法（數量＋行動指引；建議方法只計涉及來源店舖數量，不重複列出店號，店號請查閱「調貨來源 (Site Code)」欄）；另顯示 Supply Source 與 D001 Net Stock Qty (CM)；缺口跟進按供貨來源分支 — Warehouse＋D001 有貨由 D001 補貨、Warehouse＋D001 無貨寫 D001缺貨、Vendor 貨向供應商追貨（Vendor 由供應商直送店舖，不經 D001）
- 輸出：Summary（SKU 彙總：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001、各 Nature 庫存/月銷加總（取整）、TurnDay（取整）、涉及店舖數（同 SKU 同組 Site Code 不重複計數，空白不計，置於該組 TurnDay 之後）、各 Nature 庫存/月銷佔比 (%)（0.00%）、強制調貨提示、NT/Mix 預計調出量、存貨偏重提示）+ RAW（欄位按可讀性重排——識別→品項屬性→庫存與銷售→SKU 級判定→調貨執行→其他；不顯示「Item_Shop」與所有含「Value」的金額欄；調貨計算欄位附於每一列）+ 店舖問題 SKU + 調貨分析摘要（Dashboard：8 項指標卡 + Buyer 需調貨 SKU 統計表）
     """,
     """
Tourist 缺貨、Mix/NT 調撥計算：
1. **觸發條件（SKU 級，同時符合才標示）** = 該 SKU 的「Tourist 加總 TurnDay」（各 Tourist 店 Shop Stock Qty (CM) 加總 ÷ Shop Monthly Sold Qty (P90D) 加總 × 30.5，與 Summary 同口徑）< 門檻（預設 14 日，UI 可調），且 D001 Net Stock Qty (CM) < 門檻（預設 100 件，UI 可調），且 Mix / Non Tourist 至少一個 Nature 的加總 TurnDay ≥ 門檻（有貨可調）；Tourist 加總 TurnDay ≥ 門檻一律無須調貨，Mix 與 Non Tourist 加總 TurnDay 皆低於門檻時視為無貨可調、不標示
2. **調入需求量** = ceil(P90D / 90 × 14) − Shop Stock Qty (CM)（P90D = Sold Qty (P90D)，90 日總量換算日均 × 14 日）；只為已標示「需要調貨」的 SKU 計算，個別 Tourist 店需求 ≤ 0 不標示
3. **可轉出貨量（Mix / Non Tourist 來源店）** = max(0, Shop Stock Qty (CM) − P90D / 90 × 14)；「無銷售但有存貨」全數可釋放；RP Type 為 RF 與 ND 的店皆可作為來源；該 Nature 的 SKU 級加總 TurnDay < 門檻時視為無貨可調（可轉出量歸零）
4. **分配優先** = 按 SKU 彙總調入需求，Non Tourist 店可轉出量優先，不足部分由 Mix 店補足；需求列按需求量大→小依序分配；來源店與需求店同為一店（Site Code 相同）時一律排除，防止自我調貨
5. **存貨偏重提示（Summary）** = 個別 Nature 的庫存佔比 ≥ 40% 且（庫存佔比 − 銷量佔比）≥ 20 個百分點 →「存貨偏重 {Nature}（庫存 x% / 銷量 y%）」；全 SKU 月銷為 0 時銷量佔比以 0 計；多個 Nature 符合時以；分隔、差距大者在前；總庫存為 0 的 SKU 不提示；偏重提示欄（AC）與偏重 Nature 的庫存佔比欄（T/V/X）以醒目橘底（FFC000）加粗體加強顯示（Excel 與 Streamlit 同步）
6. **店舖問題 SKU** = 每間店舖需行動列（需要調入／可轉出供應，缺口在同一列加註），按店舖分組、店內按嚴重度排序，逐列顯示問題說明與建議方法（數量＋行動指引；建議方法只計涉及來源店舖數量，不重複列出店號，店號請查閱「調貨來源 (Site Code)」欄）
7. **輸出** = Summary（SKU 級彙總：Brand / SKU / SKU Desc / Buyer / Adj. SST / Supply Source / D001；各 Nature 庫存/月銷加總（取整）、TurnDay = ROUND(加總庫存 ÷ 加總月銷 × 30.5, 0)、涉及店舖數（同 SKU 同組 Site Code 不重複計數，空白不計，置於該組 TurnDay 之後）、各 Nature 庫存/月銷佔比 (%) = 該 Nature 加總 ÷ 三 Nature 合計（0.00%）、強制調貨提示、由 NT/Mix 店舖預計調出數量、存貨偏重提示——偏重提示欄與偏重 Nature 的庫存佔比欄以醒目橘底加強顯示）+ RAW 完整明細（欄序按可讀性重排：識別（Buyer/Brand/SKU/SKU Desc/Site Code/Shop Name/OM/Shop Nature）→品項屬性→庫存與銷售（含 D001 三欄）→SKU 級判定（各 Nature 加總 TurnDay／需要調貨／調入需求量）→調貨執行（可轉出貨量 NT/Mix／實際調入量／調入缺口／轉出量／調貨來源）→其他原始欄按原序；不顯示「Item_Shop」與所有含「Value」的金額欄，RAW dataframe 仍保留原欄供稽核），RAW 新增：需要調貨 / 調入需求量 (補到14日) / 實際調入量 / 調入缺口 / 可轉出貨量 (Mix) / 可轉出貨量 (Non Tourist) / 轉出量 / 調貨來源 (Site Code) + 調貨分析摘要（Dashboard：需要調貨列數 / 需調貨 SKU 數 / 調入需求量 / 實際調入量 / NT 轉出量 / Mix 轉出量 / 無法足額缺口 / 總明細行數 + Buyer 需調貨 SKU 統計表：各 Buyer 涉及需調貨的 SKU 數（已刪除重複）與佔比）
     """),

    ("🔀 Safety Stock x NetStock Review",
      """
- 以「調貨事件」為入口 Review Safety Stock 設定是否為成因（**Safety 基準周轉天數按 Shop Nature 分別設定**，預設 Tourist 7 / Mix 9 / Non Tourist 10 日，可調 7–10；與調貨段的 14 日線互相獨立）
- 調貨段沿用 Stock Checking 規則（14 日目標、NT 優先→Mix 補足、D001/TurnDay 門檻），並疊加本模式的資料品質補正（Shop Nature 正規化、月銷逐列回補至 Effective 稽核欄、D001 逐列 Net→Stock 回補），同檔與 Stock Checking 模式的結果可能有細微差異；分配跑兩遍：**不限OM** 與 **同OM**（來源店須與需求店同 OM，例 Ivy HBA6 → Ivy HB94）；不再執行 Safety Stock Review 引擎（無建議值 / Rate / MTD 參數）
- 可選擇 **分類維度**（預設「以 Shop Nature 分類」）：① 以 Shop Nature 分類——需求方＝Tourist 且 RP Type＝RF；② 以 RP Type 分類——需求方＝任何 Nature 的 RF 列（Mix／Non Tourist 缺貨亦可調入，允許同 Nature 店間再平衡）；③ 以 Shop Nature＋RP Type 分類——計算同 ①，Summary／交叉表按 6 組（如 Tourist-RF、Mix-ND）呈現；④ 以 OM 分類——計算同 ①，Summary／交叉表按來源 OM 欄的不重複值分組呈現；⑤ 以 Shop Nature＋OM 分類——計算同 ①，Summary／交叉表按 Nature×OM 組合（如 Tourist-XXX）呈現；⑥ 以銷量分類——計算同 ①，Summary／交叉表按銷量帶（高銷／中銷／低銷／零銷）分組呈現；⑦ 以銷量＋Shop Nature分類——計算同 ①，Summary／交叉表按 Nature×銷量帶 12 組（如 Tourist-高銷、Mix-低銷）呈現，銷量帶口徑與 ⑥ 相同。RP Type 語意＝該店舖可否補貨（RF＝可補貨、ND＝不可補貨）；**ND 列一律不作調貨需求方（仍可作供應方）**；TurnDay 池與 Safety 基準天數恆按 Shop Nature，不隨維度改變；來源缺「OM」欄或全欄空白時，④⑤ 退回 Shop Nature 維度
- RAW／主表恆附 **銷量帶** 與 **銷量存貨分類** 兩個顯示欄（全列覆蓋，與「銷量再平衡」開關無關）：銷量帶＝同 SKU 內按 Sold Qty (P90D) 排名分帶（前 25%＝高銷、最低 25%＝低銷、其餘＝中銷、P90D≤0＝零銷；有銷店舖少於 4 間時不區分高低，全歸中銷；並列以平均排名同帶處理）；銷量存貨分類＝高銷且店舖周轉日數（Shop Stock Qty (CM) ÷ 生效月銷 × 30.5）低於缺貨線（預設 7 日）→「高銷缺貨」；低銷且周轉日數達積壓線（預設 30 日）且庫存達保底件數（預設 5 件）→「低銷積壓」；零銷且庫存達保底件數→「零銷積壓」；其餘→「正常」。三條門檻於參數區可調，與調貨 14 日線及 Safety 基準天數互相獨立；銷量帶按本次上傳資料內同 SKU 排名計算，篩選範圍不同時分帶結果會不同
- 可勾選 **「啟用銷量再平衡」**（預設關閉）：高銷缺貨的 RF 店舖列即使所屬 Nature 池不缺貨，亦可加入調貨需求方（仍須補到 14 日需求 > 0、D001 < 門檻且同 SKU 有貨可調）；積壓列（低銷／零銷積壓）經既有轉出判定提供來源（轉出量公式與池級歸零不變——需求方擴大後更多 SKU 進入調貨，健康池內積壓列自動入事件表）。跟隨單向／雙向框架：單向時銷量需求方仍限 Tourist 列、轉出池仍限 Mix／Non Tourist；雙向時兩端放寬至任何 Nature。自我調貨／同 Nature／同OM 守衛與 ND 不作需求方全部不放寬；新事件列自動流入 Safety Review、店舖問題 SKU、Buyer 統計與簡易調貨明細；未勾選時輸出欄位與數值與現行完全一致（銷量分類兩欄除外，屬新增顯示欄）；本選項只適用於本模式
- 可勾選 **「允許雙向調貨」**（預設關閉）：存貨偏重遊客區而其他區域賣得好時，任何 Nature 缺貨店舖均可調入、遊客區店舖亦可轉出（來源固定 Non Tourist → Mix → Tourist，需求 Nature 自身一律排除；「以 RP Type 分類」維度例外——允許同 Nature 互調）；本選項只適用於本模式，Stock Checking 仍為單向
- 可勾選 **「ND 優先轉出」**（預設關閉）：**同一 Shop Nature 池內**，RP Type＝ND（不可補貨）的來源店先於 RF 店轉出，池間順序仍為 Non Tourist → Mix → Tourist（雙向時才含 Tourist），展開為 Non Tourist-ND → Non Tourist-RF → Mix-ND → Mix-RF → Tourist-ND → Tourist-RF 六段，ND／RF 兩段各自維持來源檔列序；**轉出資格一律不放寬**（池級 TurnDay 門檻、14 日超額、自我調貨／同 Nature／同OM 守衛全部照舊，ND 列仍不作調貨需求方），正常一店×一SKU 一列的資料下各 Nature 轉出合計與調入總量不變（若同一 Site Code 同時出現在需求與來源端——「Site Code × SKU」重複列異常，系統會列 warning——自我調貨守衛會令池內取用順序影響合計與缺口），只改變逐列轉出量、調貨來源、事件類型與調貨後庫存／TurnDay，**不影響 Safety Review 與建議調整量**；來源檔缺「RP Type」欄時自動失效並列 warning；未勾選時輸出欄位與數值與現行完全一致；本選項只適用於本模式
- 需調入列：現有 Safety Stock (現時)（即來源檔案的 Adj. Safety Stock）< Safety 周轉所需量 → 「Safety 設定不足」；有貨可轉出列（實際轉出或可轉出未使用）：現有設定 > 生效基準（所需量按下調補底 3 件後的值）→ 「Safety 設定過高」，並給出建議調整值與調整量（下調建議補底 3 件——受限 SAP MOQ，Safety 1–2 件難以補貨：現有 ≥ 3 件而下調至 1–2 件者建議最低 3 件；現有本身 1–2 件不強制上調；零銷售歸零建議不受保底；上調建議照常）；OM 限制只影響調貨分配，**Safety Review 與建議調整兩變體相同**
- 新增「🏪 店舖問題 SKU」：每間店舖需行動列（只收 Safety 設定不足／過高，另加 Past 90 Days 0 Sales 轉 ND 候選——M-2／M-1／CM-MTD／P90D 皆零銷售且有庫存，距參考日期 90 天內新品、需要調貨列與 ND 列除外；設定正常列排除；調入缺口在建議方法加註，建議方法只計涉及來源店舖數量，不重複列出店號），按店舖分組、店內按嚴重度排序；主表以不限OM為準，另顯示 Supply Source、D001 Net Stock Qty (CM) 與三個同OM對照欄；缺口跟進按供貨來源分支 — Warehouse＋D001 有貨由 D001 補貨、Warehouse＋D001 無貨寫 D001缺貨、Vendor 貨向供應商追貨
- 參考日期：新增「📅 參考日期」輸入（預設昨日），New SKU 判定（Launch Date 起 90 天內視為新品）以此日期起計，與報表數據口徑一致
- Shop Nature 固定全三種；Other/空白/大小寫差異一律正規化為 Non Tourist；RP Type 其他/空白值歸入 ND 並列 warning（來源缺「RP Type」欄時全部列按 RF 處理；rp／both 維度退回 Shop Nature，om／nature_om 維度不退回）；混合必要欄位嚴格檢核（不再要求 Sold Qty (CM - MTD)）；OM 為選填欄位（缺失時同OM＝不限OM＋warning；om／nature_om 維度另退回 Shop Nature）
- 輸出：RAW（欄位按可讀性重排——識別→品項屬性→庫存與銷售（含銷量帶／銷量存貨分類與 D001 三欄）→SKU 級判定→Review 結果→調貨 (不限OM)→調貨 (同OM)→其他；不顯示「Item_Shop」與所有含「Value」的金額欄）/ Summary（SKU 級彙總，按所選分類維度分組：nature 三組／rp 兩組 RF／ND／both 六組／om 動態 OM 名單／nature_om 三 Nature × OM／sales 四銷量帶／sales_nature 三 Nature × 四帶 12 組，各組庫存/月銷加總、TurnDay、涉及店舖數（同 SKU 同組 Site Code 不重複計數，空白不計）、佔比，含存貨偏重提示（門檻按維度切換：nature／rp／both 為庫存佔比 ≥40% 且差距 ≥20 個百分點，om 為 ≥25% 且差距 ≥12 個百分點，nature_om／sales_nature 為 ≥12% 且差距 ≥8 個百分點，sales 為 ≥30% 且差距 ≥15 個百分點）——偏重提示欄與偏重分組的庫存佔比欄以醒目橘底加強顯示；nature 雙向時另加「由Tourist 店舖預計調出數量」欄）/ 調貨 Review Safety Stock (不限OM) / 調貨 Review Safety Stock (同OM)（店舖×SKU 主表含 RP Type 與銷量帶／銷量存貨分類欄）/ 店舖問題 SKU（每間店舖需行動列：問題＋建議方法）/ Review 摘要（指標卡 16 項（雙向或「ND 優先轉出」有實際轉出時 18–20 項；「銷量再平衡」開啟時末尾另加銷量分類概況一排 4 項）+ ① Review × 分類維度交叉表 + ② Safety 調整方向分類（增加／減少／不變）+ ③ Safety 調整明細（SKU 級彙總）+ ④ Buyer 統計（需調貨＋需調整 Safety 合併表）+ ⑤ 銷量 × 存貨分類交叉表（母體為全部店舖×SKU 列，無調貨事件時仍輸出））/ 簡易調貨明細 (不限OM) / 簡易調貨明細 (同OM)（一列一筆來源店→接收店調貨，來源如「HB30:2; HB49:3」拆成兩行）共 9 個工作表（含首頁「說明」導讀）；結果頁另有「🏪 店舖問題 SKU」tab 與「🧾 簡易調貨明細」tab（可切換不限OM／同OM）與「📊 Review 摘要（Safety 調整統計）」tab（①–⑤ 五張表）
      """,
      """
調貨事件 Review Safety 設定：
1. **調貨段** = 與 Stock Checking 同規則（SKU 級加總 TurnDay 門檻、D001 門檻、14 日需求/供應、NT 優先→Mix 補足、需求量大→小），並疊加本模式特有的資料品質補正：① Shop Nature 正規化（大小寫不敏感、無法辨識值歸 Non Tourist；Stock Checking 為原值精確比對）、② 月銷空白／為 0 但 P90D 有銷售時逐列以 P90D ÷ 90 × 30.5 回補（實際採用值另存「Effective Shop Monthly Sold Qty (P90D)」稽核欄，Stock Checking 以 0 計）、③ D001 逐列 Net 空白時回退 Stock 欄（Stock Checking 為整欄級 fallback）——同一份輸入在兩個模式的調貨結果可能出現細微差異，本模式的補正屬較嚴謹的正確化處理；分配跑兩遍——**不限OM**（契約欄位：實際調入量 / 調入缺口 / 轉出量 / 調貨來源 (Site Code) / 調貨後庫存 / 調貨後 TurnDay）與**同OM**（同欄位加「(同OM)」後綴；來源店 OM 與需求店 OM 相同才可轉出，公司規定某些情況只允許同 OM 店舖調貨）；來源店與需求店同為一店（Site Code 相同）時一律排除，防止自我調貨
2. **分類維度（預設 Shop Nature）** = 決定調貨需求方過濾與 Summary／交叉表分組；**TurnDay 池與 Safety 基準天數恆按 Shop Nature，不隨維度改變**。RP Type 正規化：strip＋大寫後 RF→RF、ND→ND，其他／空白→ND（保守視為不可補貨）並列 warning；**需求方一律排除 ND 列（ND 仍可作供應方）**——① nature：需求方＝Tourist 且 RF（其餘同現行規則，含池級歸零與「需求 Nature 自身除外」守衛）；② rp：需求方＝任何 Nature 且 RF（該 Nature 加總 TurnDay < 門檻），供應方維持 Mix／Non Tourist（雙向時另加 Tourist），並取消池級歸零與同 Nature 守衛（可轉出貨量按逐列 14 日超額計，允許同 Nature 店間再平衡；有貨可調＝同 SKU 任一列有超額）；③ both：計算與 ① 完全相同，Summary／交叉表按 6 組（Tourist-RF、Tourist-ND、Mix-RF、Mix-ND、Non Tourist-RF、Non Tourist-ND）分組；④ om：計算與 ① 完全相同，Summary／交叉表按來源 OM 欄的不重複值分組（去除前後空格，空白自成「空白」組，大小寫不敏感，按字串排序）；⑤ nature_om：計算與 ① 完全相同，Summary／交叉表按 Nature×OM 組合（如 Tourist-XXX）分組；⑥ sales：計算與 ① 完全相同，Summary／交叉表按銷量帶（高銷／中銷／低銷／零銷）分組（銷量帶為常開產出欄，缺欄時顯式退回 Shop Nature 鍵，正常輸入不會觸發）；⑦ sales_nature：計算與 ① 完全相同，Summary／交叉表按 Nature×銷量帶 12 組（如 Tourist-高銷、Mix-低銷）分組，銷量帶口徑與 ⑥ 相同。來源缺「RP Type」欄時全部列按 RF 處理（ND 排除規則無法套用），僅 ②③ 退回 nature 維度＋warning（④⑤⑥⑦ 不退回）；來源缺「OM」欄或全欄空白時 ④⑤ 退回 nature 維度＋warning；非法維度值退回 nature 不拋錯
3. **雙向調貨（勾選後，預設不勾）** = 任何 Nature 缺貨皆可調入：該 SKU 的該 Nature 加總 TurnDay < 門檻 且 D001 < 門檻 且個別店補到 14 日需求 > 0，且**其他任一 Nature** 加總 TurnDay ≥ 門檻（有貨可調）即標示需要調貨（需求列另須 RP Type＝RF）；新增 RAW 欄「可轉出貨量 (Tourist)」，遊客區店舖亦可轉出，來源優先序固定 **Non Tourist → Mix → Tourist**，需求 Nature 自身一律排除（分配層顯式守衛＋池級歸零雙重保證：缺貨 Nature 的超額列可轉出歸零、不進入來源池，亦不作同 Nature 互調；「以 RP Type 分類」維度則取消此兩項限制，允許同 Nature 互調）；未勾選時維持單向（只向遊客區調入），輸出欄位與數值完全不變
   - **ND 優先轉出（勾選後，預設不勾）** = 只改**池內**取用順序：同一 Shop Nature 池內 RP Type＝ND 的來源列先於 RF 列，池間順序仍為 Non Tourist → Mix → Tourist（雙向時才含 Tourist），展開為 Non Tourist-ND → Non Tourist-RF → Mix-ND → Mix-RF → Tourist-ND → Tourist-RF；ND 段與 RF 段各自維持來源檔列序（不按可轉出貨量重排）。「ND」口徑＝正規化後非 RF（含無法辨識而歸入 ND 的列）；轉出資格一律不放寬（池級歸零、逐列 14 日超額、自我調貨／同 Nature／同OM 守衛照舊），ND 列仍不作需求方。因轉出列母體與分配順序無關，本開關只改變 `轉出量`／`調貨來源`／`事件類型`／`調貨後庫存`／`調貨後 TurnDay`（含同OM變體），**Safety Review、建議 Safety 設定值與調整量、各 Nature 轉出合計與調入總量全部不變**；缺「RP Type」欄時自動失效並列 warning；未勾選時輸出與現行完全一致；第 5 排另加「ND 來源轉出量」／「ND 來源轉出量 (同OM)」兩卡（開啟且 ND 確實轉出時）
   - **銷量再平衡（勾選後，預設不勾）** = 需求方在原判定之外**增列**「高銷缺貨」列：同 SKU 內屬高銷帶且店舖周轉日數低於缺貨線的 RF 列，豁免所屬 Nature 池加總 TurnDay 門檻（本通道存在目的），但仍須補到 14 日需求 > 0、D001 < 門檻、同 SKU 有可轉出量；單向模式需求方仍限 Tourist 列、轉出池仍限 Mix／Non Tourist（跟隨方向框架；雙向／rp 維度任何 Nature）。供應方轉出量公式不變（逐列 14 日超額＋池級歸零）——需求方擴大後更多 SKU 進入調貨，健康池內積壓列經既有 donor 判定自動入事件表；自我調貨／同 Nature／同OM 守衛與 ND 不作需求方全部保留。stats 另記 `sales_need_rows`（僅由本通道新增的需求列數）；開啟時指標卡末尾另加銷量分類概況一排（高銷缺貨列數／SKU 數、積壓列數／積壓庫存件數）
   - **銷量帶與銷量存貨分類（常開顯示欄）** = 「銷量帶」按同 SKU 內 Sold Qty (P90D) 的店舖**平均排名**分帶：排名超過 75%＝高銷、排名在最低 25% 內＝低銷、其餘＝中銷、P90D≤0＝零銷（有銷店舖少於 4 間時不產生高銷／低銷，全歸中銷，避免小樣本假區分；並列同值同帶）；「銷量存貨分類」以店舖周轉日數（Shop Stock Qty (CM) ÷ 生效月銷 × 30.5）與三條可調門檻判定：高銷缺貨＝高銷且周轉 < 缺貨線（預設 7 日，可調 1–14）；低銷積壓＝低銷且周轉 ≥ 積壓線（預設 30 日，可調 15–90）且庫存 ≥ 保底件數（預設 5 件，可調 1–100）；零銷積壓＝零銷且庫存 ≥ 保底件數；其餘＝正常。邊界口徑：缺貨用「低於」、積壓用「不少於」。兩欄恆輸出（不受銷量再平衡開關影響），與調貨 14 日線及 Safety 基準天數互相獨立；零銷有存但未達保底件數的列歸正常，仍由「店舖問題 SKU」的 Past 90 Days 0 Sales 轉 ND 提示涵蓋
4. **Review 基準** = Safety 周轉所需量 = ceil(Sold Qty (P90D) ÷ 90 × 該列 Nature 的基準天數)（無銷售列為 0——該列為轉出列時，任何大於 0 的現有設定都會判「Safety 設定過高」並建議降至 0，與調貨段 P90D 口徑一致，屬設計取捨）；基準天數按 Nature 於左側分別設定（預設 Tourist 7 / Mix 9 / Non Tourist 10，可調 7–10），與調貨觸發門檻及調貨的 14 日線完全無關 — 改變基準天數不會影響任何調貨欄位；RAW 與主表逐列回顯「Safety 基準天數」供稽核
5. **事件列與 Review** = ①「需要調入」：需要調貨=是 的列（單向只出現在 Tourist-RF；雙向任何 Nature 的 RF；rp 維度任何 Nature 的 RF）——現有 Safety Stock (現時) < 所需量 → 「Safety 設定不足」，否則「設定正常—缺貨非設定因素」；②「實際轉出」／「可轉出未使用」：已標示調貨 SKU 且有轉出量或可轉出貨量的列（單向為 Mix / Non Tourist；雙向含 Tourist；ND 列亦可為來源）——現有設定 > 生效基準（所需量按下調補底後的值，見第 6 點）→ 「Safety 設定過高」，否則「設定正常—過剩非設定因素」；等於邊界一律視為正常；未標示調貨的 SKU 不入 Review；個別店缺貨但其 Nature 加總 TurnDay ≥ 門檻時不列為需求（同 Nature 店間再平衡不在範圍內，與 Stock Checking 同口徑；「以 RP Type 分類」維度與「銷量再平衡」開啟時的高銷缺貨列除外）；入表列集合兩變體相同，僅事件標籤與調貨數值按變體計算
6. **建議調整（不受 OM 限制影響）** = 建議 Safety 設定值：不足列 = Safety 周轉所需量（上調屬 TurnDay 驅動，照常給足，含上調至 1–2 件）；過高列 = 生效基準（**下調補底 3 件**——受限 SAP MOQ，Safety 1–2 件難以補貨而 I14 原檔無 MOQ 欄：所需量 ≥ 1 且現有設定 ≥ 3 件時生效基準 = max(所需量, 3)，補底不高於現值，現有 3 件而所需 1–2 件者補底後相等 → 改判設定正常；現有設定本身只有 1–2 件者維持原所需量，不強制上調到 3；所需量 = 0 的歸零建議屬退場，不受保底影響）；僅「不足／過高」列給建議；Safety 建議調整量 = 建議值 − 現有設定（正=增加、負=減少；正常列留空）；「Safety 周轉所需量」欄維持原所需量作稽核口徑，保底只體現在建議值與調整量；Safety Stock (現時) 空白或非數值時 Review 視為 0 並列 warning
7. **上下文欄位** = 調貨後庫存 = Shop Stock Qty (CM) + 實際調入量 − 轉出量；調貨後 TurnDay = 調貨後庫存 ÷ Shop Monthly Sold Qty (P90D) × 30.5（月銷 0 留空；月銷為回補後值，實際採用值另存「Effective Shop Monthly Sold Qty (P90D)」稽核欄）；同OM變體另存「(同OM)」後綴欄
8. **OM 欄位** = 選填；缺「OM」欄時同OM調貨退化為與不限OM相同（全部店舖同一群組）並列 warning；OM 空白列自成「空白」群組（大小寫不敏感）；om／nature_om 分類維度的分組名單即此欄的不重複值（排序固定），缺欄或全欄空白時該兩維度退回 Shop Nature
9. **Shop Nature** = 固定三種；Other/空白/大小寫差異一律正規化為 Non Tourist，映射列數列 warning；不做 Nature 篩選（避免破壞調貨供需母體）
10. **輸入檢核** = `Buyer / Brand / SKU / Shop Nature / Site Code / Shop Stock Qty (CM) / Sold Qty (P90D) / Safety Stock (現時)／Adj. Safety Stock（兩名皆接受，語意相同：Safety Stock (現時) ＝ 來源檔案的 Adj. Safety Stock） / Shop Monthly Sold Qty (P90D)` 皆為混合必要欄位，不要求未參與計算的 `Shop Stock Status`，缺任一即停止；`RP Type` 為選填欄位（缺失時 ND 排除規則無法套用；rp／both 維度退回 nature 並列 warning）；`OM` 為選填欄位（缺失或全欄空白時 om／nature_om 維度退回 nature 並列 warning）；`Shop Stock Qty (CM)` 必須為非負整數件數，小數、負數及非有限值即拒絕；店舖銷售、月銷及 Safety 設定的非有限值即拒絕，負數銷售按 0 處理並列 warning。月銷空白／非數值／負數，以及月銷為 0 但 P90D 有正數銷售時，逐列以 P90D 換算並列 warning；D001 沿用 Net→Stock→0 fallback，非有限 D001 視為空白，負 Net 值保留，Summary 顯示實際採用值；`Site Code × SKU` 重複列、同店 Shop Nature 不一致及 OM 全欄空白均列 warning（正常輸入應為一店×一SKU一列）
11. **簡易調貨明細（tab 與 Excel）** = 一列一筆「來源店舖 → 接收店舖」調貨，欄位固定 Brand / SKU / SKU Desc / Supply Site Code(調貨來源) / Supply Shop Name / Supply RP Type / Supply OM / 調貨數量 / Target Site Code(接收店舖) / Target Shop Name / Target RP Type / Target OM（RP Type 緊接對應 Shop Name 之後，調貨數量放在接收 Site Code 前作 Supply／Target 分隔）；調貨來源字串（如「HB30:2; HB49:3」）按分號逐段拆行，同一接收列同一來源店重複出現時合併加總；只收錄「需要調貨=是」且實際有調入（實際調入量 > 0）的列，無貨可調的缺口列不進簡易版；來源店的 Shop Name／OM／RP Type 以 RAW 全量 Site Code 查找表補回（查不到留空）；按 Brand / SKU / 來源店 / 接收店排序；結果頁 tab 可切換不限OM／同OM，Excel 則固定輸出兩個工作表各一套
12. **輸出 9 工作表** = 說明（首頁導讀：回顯本次執行參數（分類維度／調貨門檻／D001 門檻／Safety 基準天數／各開關／銷量分類門檻／參考日期）＋核心邏輯（調貨段 14 日線、Review 段 Safety 基準、銷量帶與銷量存貨分類、0 Sales 轉 ND）＋工作表導讀（每頁內容與建議關注的部分）＋速查提示；打開 Excel 預設停在此頁）/ RAW（欄序按可讀性重排：識別（Buyer/Brand/SKU/SKU Desc/Site Code/Shop Name/OM/Shop Nature）→品項屬性（含 RP Type）→庫存與銷售（含銷量帶／銷量存貨分類與 D001 三欄）→SKU 級判定（各 Nature 加總 TurnDay／需要調貨／調入需求量）→Review 結果（事件類型 / Safety 基準天數 / Safety 周轉所需量 / Safety Review / 建議 Safety 設定值 / Safety 建議調整量）→調貨 (不限OM)→調貨 (同OM)→其他原始欄，非事件列 Review 欄留空；不顯示「Item_Shop」與所有含「Value」的金額欄；雙向時另含「可轉出貨量 (Tourist)」）/ Summary（SKU 級彙總，按所選分類維度分組：nature 三組＝與 Stock Checking 同格式同值（A–AC 契約欄序：各組加總庫存／加總月銷／TurnDay／涉及店舖數，店舖數為同 SKU 同組 Site Code 不重複計數，空白不計）、rp 兩組（RF／ND）、both 六組、om（動態 OM 名單）、nature_om（三 Nature × OM）、sales（高銷／中銷／低銷／零銷 四帶）、sales_nature（三 Nature × 四帶 12 組）；各組庫存/月銷加總、TurnDay、涉及店舖數、佔比、強制調貨提示、存貨偏重提示（門檻按維度切換：nature／rp／both 為庫存佔比 ≥40% 且差距 ≥20 個百分點，om 為 ≥25% 且差距 ≥12 個百分點，nature_om／sales_nature 為 ≥12% 且差距 ≥8 個百分點，sales 為 ≥30% 且差距 ≥15 個百分點）——偏重提示欄與偏重分組的庫存佔比欄以醒目橘底（Streamlit 與 Excel 同步）加強顯示；nature 雙向時末尾另加「由Tourist 店舖預計調出數量」欄）/ 調貨 Review Safety Stock (不限OM) / 調貨 Review Safety Stock (同OM)（店舖×SKU 主表，含 Site Code／OM／Shop Nature／RP Type 與銷量帶／銷量存貨分類，按 Brand / SKU / 事件類型 / Site Code 排序）/ 店舖問題 SKU（每間店舖需行動列：Safety 設定不足／過高＋Past 90 Days 0 Sales 轉 ND 候選（四期皆零銷售且有庫存，新品除外；重疊列保留原問題類型並加註）；按店舖分組、店內按嚴重度排序；主表以不限OM為準，另附三個同OM對照欄；調入缺口在建議方法加註）/ Review 摘要（指標卡——第 4 排為同OM調貨對照、第 5 排為轉出來源對照（雙向時含 Tourist 轉出量兩卡，「ND 優先轉出」開啟且 ND 店舖有實際轉出時另加「ND 來源轉出量」兩卡，該排 2–4 卡）、末尾為銷量分類概況一排（「銷量再平衡」開啟時：高銷缺貨列數／高銷缺貨 SKU 數／積壓列數／積壓庫存件數；無第 5 排時該排為第 5 排） + Safety Review × 分類維度交叉表（欄位按維度切換）：列數 / SKU 數 / 建議調整量合計，含總計列；其下另加「Safety 調整方向分類」（增加 / 減少 / 不變 × 列數 / SKU 數 / Safety 建議調整量合計，含總計列）與「Safety 調整明細（SKU 級彙總）」，最後為「銷量 × 存貨分類交叉表」（母體為全部店舖×SKU 列：各分類 × Shop Nature 列數、總計列數／SKU 數／店舖數、庫存與 P90D 銷量合計，含總計列；無調貨事件時仍輸出））/ 簡易調貨明細 (不限OM) / 簡易調貨明細 (同OM)（見第 11 點；凍結表頭 + 自動篩選；無實際調貨時顯示「無符合資料」）
13. **Safety 調整統計（「Review 摘要」與同名 tab）** = 以有調貨事件的店舖×SKU 列為母體（與 Safety Review × 分類維度交叉表相同），統計「調整 Safety」的影響：① **Safety 調整方向分類** — 按「Safety 建議調整量」的正負分為 增加／減少／不變，各列 列數、SKU 數、Safety 建議調整量合計，末列「總計」（SKU 數為不重複計算，同一 SKU 可跨分類；合計為淨調整量）；四捨五入後調整量為 0 的列歸「不變」。② **Safety 調整明細（SKU 級彙總）** — 每個有事件的 SKU 一列：Review 列數、需增加店舖數、需減少店舖數、設定正常店舖數（三者的店舖數按 Site Code 不重複計算）、現有 Safety 設定合計（調整列）、建議 Safety 設定合計（調整列）（兩個合計只累加「Safety 設定不足＋過高」的列）與 淨調整量（= 該 SKU 各列「Safety 建議調整量」加總），按 Brand / SKU 排序。兩個統計不受 OM 限制、「允許雙向調貨」、「ND 優先轉出」與分類維度影響（Safety 建議與調貨分配順序無關）
14. **Buyer 統計（需調貨＋需調整 Safety 合併表）** = 與 Stock Checking 同口徑：同 SKU 取全表首個非空 Buyer（空白歸入「(未填 Buyer)」），已刪除重複；需調貨與需調整 Safety 按 Buyer 外連接合併為同一列（缺席以 0 補齊），按需調貨數降序、再按需調整數降序。欄位為 Buyer／需調貨 SKU 數／需調貨 SKU 佔比／調入需求量合計／需調整 Safety SKU 數／需調整 Safety SKU 佔比／Safety 建議調整量合計（需調整母體為 Safety 設定不足＋過高；調入需求量與調整量為淨額，分別與指標卡總需求及調整合計對帳）。合併表跟隨現行分類維度與調貨結果，不受 OM 限制影響，空集合時不輸出
      """, ("assets/sales_band_guide.png",)),
]

for _label, _desc, _rules, *_extra in MODES:
    with st.expander(_label, expanded=False):
        st.markdown("#### 🎯 模式說明")
        st.markdown(_desc)
        if _rules:
            st.markdown("#### ⚙️ 核心計算規則")
            st.markdown(_rules)
        if _extra and _extra[0]:
            with st.expander("🖼️ 銷量帶分類說明圖（高銷／中銷／低銷／零銷）", expanded=False):
                st.image(_extra[0], width="stretch")
                st.caption("同 SKU 內按店舖 Sold Qty (P90D) 排名分帶；與輸出 Excel「說明」工作表的口徑一致。")

# 4. 必要欄位檢核標準（單一 HTML 區塊）
st.markdown("""
<div class="custom-card">
    <h3 style="margin-top: 0;">🔍 必要欄位檢核標準</h3>
    <p>共用必要欄位（依模式檢核，混合模式不要求未參與計算的 Shop Stock Status）：</p>
    <ul>
        <li>✅ <b>Buyer</b> <span style="color:#0077a3;font-size:0.8em">(採購人員，全部模式)</span></li>
        <li>✅ <b>Brand</b> <span style="color:#0077a3;font-size:0.8em">(品牌名稱，全部模式)</span></li>
        <li>✅ <b>SKU</b> <span style="color:#0077a3;font-size:0.8em">(品號編碼，全部模式)</span></li>
        <li>✅ <b>Shop Stock Status</b> <span style="color:#0077a3;font-size:0.8em">(庫存狀態明細；除 Safety Stock x NetStock Review 外全部模式)</span></li>
    </ul>
    <p>各模式額外必要欄位（缺少時停止並提示）：</p>
    <ul>
        <li>✅ <b>Stock Checking</b> <span style="color:#0077a3;font-size:0.8em">(Shop Nature／Sold Qty (P90D)／Shop Stock Qty (CM)／Site Code)</span></li>
        <li>✅ <b>Safety Stock x NetStock Review</b> <span style="color:#0077a3;font-size:0.8em">(Shop Nature／Site Code／Shop Stock Qty (CM)／Sold Qty (P90D)／Safety Stock (現時) 或 Adj. Safety Stock／Shop Monthly Sold Qty (P90D))</span></li>
        <li>✅ <b>NDRF 月份對比</b> <span style="color:#0077a3;font-size:0.8em">(Month／RP Type／Shop Stock Qty (CM)／Adj. SST／Site Code)</span></li>
        <li>✅ <b>Past 90 Days 0 Sales</b> <span style="color:#0077a3;font-size:0.8em">(Sold Qty (M-2)／Sold Qty (M-1)／Sold Qty (CM - MTD)／Shop Stock Qty (LM)／Shop Stock Qty (CM)／Site Code)</span></li>
    </ul>
    <p>OOS Review 額外建議欄位（非必要，缺少時將顯示提示）：</p>
    <ul>
        <li>📌 <b>Sold Qty (CM - MTD)</b> <span style="color:#f59e0b;font-size:0.8em">(月度銷售)</span></li>
        <li>📌 <b>Adj. Safety Stock</b> <span style="color:#f59e0b;font-size:0.8em">(現有安全庫存；報表內顯示為 Safety Stock (現時)，並非系統建議值——系統建議值另見「建議 Safety 設定值」欄)</span></li>
        <li>📌 <b>Shop Nature</b> <span style="color:#f59e0b;font-size:0.8em">(店舖性質)</span></li>
    </ul>
    <p>Safety Stock x NetStock Review 另建議包含 RP Type（RF＝可補貨／ND＝不可補貨）——ND 列不作調貨需求方；缺失時全部列按 RF 處理並退回 Shop Nature 維度（列 warning，不停止）。</p>
</div>
""", unsafe_allow_html=True)

# 5. 頁尾
st.markdown("<hr class='custom-divider'>", unsafe_allow_html=True)
st.caption(f"💡 所有模式均自動生成採購與品牌的交叉匯總樞紐分析表。｜I14 報表計算工具 v{__version__}")
