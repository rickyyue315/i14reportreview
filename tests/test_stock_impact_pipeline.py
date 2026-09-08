"""Validation tests for Stock Impact Analysis functions (v1.6.1).

v1.6.0: Shop Stock = 0（缺貨）對 Sold Qty 的因果影響，損失計算雙基準（M-2 & P90D）並列。
v1.6.1: 升級為存貨分佈分析 — Mix Turnday 分級（7-14日標準）、存貨失衡 SKU（Tourist缺貨+其他過剩）。
"""
import warnings
warnings.filterwarnings("ignore")
import numpy as np
import pandas as pd
from io import BytesIO
from transformer import (
    _add_stock_impact_columns, _build_stock_impact_summary,
    _build_sku_allocation_summary, _build_sku_stock_diagnosis,
    _build_inventory_distribution_summary, _classify_d001_status,
    _build_low_stock_item_class_summary,
    _build_sku_soh_tier_summary, _build_sell_through_band_summary,
    generate_excel_stock_impact_simplified,
    transform_stock_impact, generate_excel_stock_impact,
)

# 基準列：Warehouse SKU、Tourist 店舖、RF 店舖、有 M-2 銷售紀錄
BASE = {
    "Buyer": "B1", "Brand": "A", "SKU": "1", "SKU Desc": "Test Product", "Site Code": "S01",
    "Item_Shop": "Prod-S01", "Shop Stock Status": "Normal",
    "Supply Source": "Warehouse", "Shop Nature": "Tourist", "RP Type": "RF",
    "Shop Monthly Sold Qty (P90D)": 10.0,
    "Turnday (P90D Qty)": 15.0,
    "Shop Stock Qty (M-2)": 5, "Shop Stock Qty (M-1)": 5, "Shop Stock Qty (CM)": 5,
    "Sold Qty (M-2)": 10.0, "Sold Qty (M-1)": 8.0, "Sold Qty (CM - MTD)": 4.0,
    "OM": "OM1", "Adj. MCH1": "CatX", "Item Class": "Cream",
}


def _row(**kw):
    r = dict(BASE)
    r.update(kw)
    return r


def _df(rows):
    return pd.DataFrame([_row(**r) for r in rows])


# ──────────────────────────────────────────────────────────────────────────────
# Test 1: 缺貨判定（Shop Stock = 0）
# ──────────────────────────────────────────────────────────────────────────────
def test_oos_detection():
    df = _df([
        # r0: CM Shop Stock=0 → Is OOS (CM)=True
        {"SKU": "1", "Shop Stock Qty (CM)": 0, "Sold Qty (M-2)": 10, "Sold Qty (CM - MTD)": 0},
        # r1: 全期有貨 → 無缺貨
        {"SKU": "2", "Shop Stock Qty (M-2)": 5, "Shop Stock Qty (M-1)": 5, "Shop Stock Qty (CM)": 5},
        # r2: M-1 缺貨，M-2/CM 有貨 → 恢復供貨（M-1 OOS 但 CM 已補回）
        {"SKU": "3", "Shop Stock Qty (M-1)": 0, "Sold Qty (M-2)": 10},
        # r3: 全期缺貨 → 長期缺貨
        {"SKU": "4", "Shop Stock Qty (M-2)": 0, "Shop Stock Qty (M-1)": 0, "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 10},
    ])
    raw = _add_stock_impact_columns(df)

    assert raw.iloc[0]["Is OOS (CM)"] == True
    assert raw.iloc[0]["Is OOS (M-2)"] == False
    print("OK r0: CM 缺貨判定")

    assert raw.iloc[1]["Is OOS (CM)"] == False
    assert int(raw.iloc[1]["OOS Periods Count"]) == 0
    assert raw.iloc[1]["OOS Pattern"] == "無缺貨"
    print("OK r1: 全期有貨 → 無缺貨")

    assert raw.iloc[2]["Is OOS (M-1)"] == True
    assert int(raw.iloc[2]["OOS Periods Count"]) == 1
    assert raw.iloc[2]["OOS Pattern"] == "恢復供貨"
    print("OK r2: M-1 缺貨 → 恢復供貨")

    assert int(raw.iloc[3]["OOS Periods Count"]) == 3
    assert raw.iloc[3]["OOS Pattern"] == "長期缺貨"
    print("OK r3: 全期缺貨 → 長期缺貨")


# ──────────────────────────────────────────────────────────────────────────────
# Test 2: 損失計算（雙基準 M-2 & P90D）
# ──────────────────────────────────────────────────────────────────────────────
def test_sales_loss_calculation():
    df = _df([
        # r0: CM 缺貨，M-2 Sold=10, P90D=30 (已是月均), CM Sold=2
        #   Loss vs M-2 = max(0, 10-2) = 8；Loss vs P90D = max(0, 30-2) = 28
        {"SKU": "1", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 10.0, "Sold Qty (CM - MTD)": 2.0,
         "Shop Monthly Sold Qty (P90D)": 30.0},
        # r1: CM 缺貨，M-2 Sold=20, P90D=15 (已是月均), CM Sold=1
        #   Loss vs M-2 = max(0, 20-1) = 19；Loss vs P90D = max(0, 15-1) = 14
        {"SKU": "2", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 20.0, "Sold Qty (CM - MTD)": 1.0,
         "Shop Monthly Sold Qty (P90D)": 15.0},
        # r2: 有貨 → 不計損失（即使 Sold 下降）
        {"SKU": "3", "Shop Stock Qty (CM)": 5,
         "Sold Qty (M-2)": 30.0, "Sold Qty (CM - MTD)": 0.0},
    ])
    raw = _add_stock_impact_columns(df)

    assert int(raw.iloc[0]["Max Sales Loss (M-2)"]) == 8
    assert int(raw.iloc[0]["Max Sales Loss (P90D)"]) == 28
    assert raw.iloc[0]["Is Stock Impacted"] == True
    print("OK r0: 缺貨損失 M-2=8, P90D=28")

    assert int(raw.iloc[1]["Max Sales Loss (M-2)"]) == 19
    assert int(raw.iloc[1]["Max Sales Loss (P90D)"]) == 14
    print("OK r1: 缺貨損失 M-2=19, P90D=14（雙基準差異）")

    assert int(raw.iloc[2]["Max Sales Loss (M-2)"]) == 0
    assert raw.iloc[2]["Is Stock Impacted"] == False
    print("OK r2: 有貨不計損失")


# ──────────────────────────────────────────────────────────────────────────────
# Test 3: OOS Pattern 分類（含恢復供貨）
# ──────────────────────────────────────────────────────────────────────────────
def test_oos_pattern_classification():
    df = _df([
        # r0: M-2 缺貨 → CM 有貨 = 恢復供貨
        {"SKU": "1", "Shop Stock Qty (M-2)": 0, "Shop Stock Qty (CM)": 5, "Sold Qty (M-2)": 10},
        # r1: 僅 CM 缺貨 = 間歇缺貨
        {"SKU": "2", "Shop Stock Qty (CM)": 0, "Sold Qty (M-2)": 10},
        # r2: M-2 + CM 缺貨（2 期）= 長期缺貨
        {"SKU": "3", "Shop Stock Qty (M-2)": 0, "Shop Stock Qty (CM)": 0, "Sold Qty (M-2)": 10},
        # r3: 無需求（M-2 Sold=0, P90D=0）即使缺貨也不算 Impacted
        {"SKU": "4", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 0, "Shop Monthly Sold Qty (P90D)": 0},
        # r4: M-1 缺貨，M-2/CM 有貨 → 恢復供貨（M-1 OOS 但 CM 已補回）
        {"SKU": "5", "Shop Stock Qty (M-1)": 0, "Shop Stock Qty (M-2)": 5,
         "Shop Stock Qty (CM)": 5, "Sold Qty (M-2)": 10},
    ])
    raw = _add_stock_impact_columns(df)

    assert raw.iloc[0]["OOS Pattern"] == "恢復供貨"
    print("OK r0: 恢復供貨")

    assert raw.iloc[1]["OOS Pattern"] == "間歇缺貨"
    print("OK r1: 間歇缺貨")

    assert raw.iloc[2]["OOS Pattern"] == "長期缺貨"
    print("OK r2: 長期缺貨")

    assert raw.iloc[3]["Is Stock Impacted"] == False, "無需求證據不應算影響"
    print("OK r3: 無需求 → 不算 Impacted")

    # r4: M-1 缺貨，M-2/CM 有貨 → 恢復供貨（M-1 OOS 但 CM 已補回）
    assert raw.iloc[4]["OOS Pattern"] == "恢復供貨", "M-1 OOS 但 CM 有貨應為恢復供貨"
    print("OK r4: M-1 缺貨 → 恢復供貨（M-1 OOS 但 M-2/CM 有貨，已補回）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 4: 多維度彙總（含 Nature 分佈與排序）
# ──────────────────────────────────────────────────────────────────────────────
def test_build_stock_impact_summary():
    df = _df([
        {"SKU": "1", "Brand": "A", "Shop Nature": "Tourist", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 10.0, "Sold Qty (CM - MTD)": 2.0, "Shop Monthly Sold Qty (P90D)": 10.0},
        {"SKU": "2", "Brand": "A", "Shop Nature": "Tourist", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 20.0, "Sold Qty (CM - MTD)": 5.0, "Shop Monthly Sold Qty (P90D)": 15.0},
        {"SKU": "3", "Brand": "B", "Shop Nature": "Non Tourist", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 30.0, "Sold Qty (CM - MTD)": 10.0, "Shop Monthly Sold Qty (P90D)": 20.0},
    ])
    raw = _add_stock_impact_columns(df)
    brand_summary = _build_stock_impact_summary(raw, ["Brand"])

    # Brand A 應排在 Brand B 前面（Tourist 缺貨 2 vs 0）
    assert str(brand_summary.iloc[0]["Brand"]) == "A"
    assert int(brand_summary.iloc[0]["Tourist 缺貨行數"]) == 2
    assert str(brand_summary.iloc[1]["Brand"]) == "B"
    assert int(brand_summary.iloc[1]["Non Tourist 缺貨行數"]) == 1

    # 總計列
    total = brand_summary[brand_summary["Brand"] == "總計"]
    assert len(total) == 1
    assert int(total["缺貨行數"].values[0]) == 3
    assert int(total["Tourist 缺貨行數"].values[0]) == 2
    print("OK _build_stock_impact_summary（Tourist 排序 + Nature 分佈 + 總計）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 5: SKU 級跨 Nature 調貨建議 (已於 v1.7.1 取消輸出，保留函數備用)
# ──────────────────────────────────────────────────────────────────────────────


# ──────────────────────────────────────────────────────────────────────────────
# Test 6: 完整 transform_stock_impact 主流程
# ──────────────────────────────────────────────────────────────────────────────
def test_transform_stock_impact():
    df = _df([
        {"SKU": "1", "Adj. MCH1": "CatA", "Shop Nature": "Tourist", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 10.0, "Sold Qty (CM - MTD)": 2.0, "Shop Monthly Sold Qty (P90D)": 10.0},
        {"SKU": "2", "Adj. MCH1": "CatB", "Shop Nature": "Mix", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 20.0, "Sold Qty (CM - MTD)": 5.0, "Shop Monthly Sold Qty (P90D)": 15.0},
        {"SKU": "3", "Adj. MCH1": "CatA", "Shop Nature": "Non Tourist", "Shop Stock Qty (CM)": 5,
         "Sold Qty (M-2)": 30.0, "Sold Qty (CM - MTD)": 10.0},
     ])
    _, brand, buyer, cat, item_class, _, diag, inv, sku_monthly, raw, warns = transform_stock_impact(df)

    assert len(brand) > 0
    assert len(buyer) > 0
    assert len(cat) > 0
    assert len(item_class) > 0
    assert sku_monthly is not None and not sku_monthly.empty, "應有 SKU 月銷量/存儲量分析"
    impacted_count = int(raw["Is Stock Impacted"].sum())
    assert impacted_count == 2, f"應有 2 行缺貨影響，得 {impacted_count}"
    print(f"OK transform_stock_impact（3 維度彙總 + SKU月銷量 + 缺貨影響 {impacted_count} 行）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 7: Warehouse SKU 篩選
# ──────────────────────────────────────────────────────────────────────────────
def test_warehouse_filter():
    df = _df([
        {"SKU": "1", "Supply Source": "Warehouse", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 10.0, "Shop Monthly Sold Qty (P90D)": 10.0},
        {"SKU": "2", "Supply Source": "Vendor", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 10.0, "Shop Monthly Sold Qty (P90D)": 10.0},
    ])
    # warehouse_only=True（預設）→ 只剩 1 行
    _, _, _, _, _, _, _, _, _, raw_wh, _ = transform_stock_impact(df, warehouse_only=True)
    assert len(raw_wh) == 1
    assert str(raw_wh.iloc[0]["Supply Source"]) == "Warehouse"
    print("OK warehouse_only=True（過濾 Vendor）")

    # warehouse_only=False → 保留 2 行
    _, _, _, _, _, _, _, _, _, raw_all, _ = transform_stock_impact(df, warehouse_only=False)
    assert len(raw_all) == 2
    print("OK warehouse_only=False（保留 Vendor）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 6a: Mix Turnday 分級與可釋放存貨（v1.6.1 存貨分佈分析核心）
# ──────────────────────────────────────────────────────────────────────────────
def test_mix_turnday_banding():
    df = _df([
        # r0: Mix, Turnday 60日 → >30日(過剩), 存貨200, P90D=300 (90日總量)
        #   理想=300/90*14=46.7, 可釋放=200-46.7=153
        {"SKU": "1", "Shop Nature": "Mix", "Shop Stock Qty (CM)": 200,
         "Shop Monthly Sold Qty (P90D)": 300.0, "Turnday (P90D Qty)": 60},
        # r1: Mix, Turnday 10日 → 7-14日(理想), P90D=360 (90日總量)
        #   理想=360/90*14=56, 可釋放=0
        {"SKU": "2", "Shop Nature": "Mix", "Shop Stock Qty (CM)": 40,
         "Shop Monthly Sold Qty (P90D)": 360.0, "Turnday (P90D Qty)": 10},
        # r2: Mix, P90D=0 有存貨 → 無銷售有存貨(純囤積), 全數50可釋放
        {"SKU": "3", "Shop Nature": "Mix", "Shop Stock Qty (CM)": 50,
         "Shop Monthly Sold Qty (P90D)": 0.0, "Turnday (P90D Qty)": float("inf")},
        # r3: Tourist, Turnday 60日 → 不標 Mix Band (空字串), Excess=0
        {"SKU": "4", "Shop Nature": "Tourist", "Shop Stock Qty (CM)": 500,
         "Shop Monthly Sold Qty (P90D)": 100.0, "Turnday (P90D Qty)": 150},
    ])
    raw = _add_stock_impact_columns(df)

    assert raw.iloc[0]["Mix Turnday Band"] == ">30日(過剩)"
    assert int(raw.iloc[0]["Mix Excess Stock"]) == 153, f"得 {raw.iloc[0]['Mix Excess Stock']}"
    print("OK r0: Mix >30日過剩, 可釋放 153 件")

    assert raw.iloc[1]["Mix Turnday Band"] == "7-14日(理想)"
    assert int(raw.iloc[1]["Mix Excess Stock"]) == 0
    print("OK r1: Mix 7-14日理想, 可釋放 0")

    assert raw.iloc[2]["Mix Turnday Band"] == "無銷售有存貨(RF)"
    assert int(raw.iloc[2]["Mix Excess Stock"]) == 50, "無銷售應全數可釋放"
    print("OK r2: Mix 無銷售有存貨(RF), 全數 50 件可釋放")

    assert raw.iloc[3]["Mix Turnday Band"] == "", "Tourist 不應有 Mix Band"
    assert int(raw.iloc[3]["Mix Excess Stock"]) == 0, "Tourist 不計 Mix Excess"
    print("OK r3: Tourist 不套 Mix 標準")


# ──────────────────────────────────────────────────────────────────────────────
# Test 6b: 存貨分佈彙總（Mix Turnday 分級表 + Nature 對比 + 失衡 SKU）
# ──────────────────────────────────────────────────────────────────────────────
def test_inventory_distribution_summary():
    df = _df([
        # SKU 1: Tourist 缺貨（CM=0, 有需求）
        {"SKU": "1", "Brand": "A", "Shop Nature": "Tourist", "Site Code": "T01",
         "Shop Stock Qty (CM)": 0, "Sold Qty (M-2)": 10.0, "Shop Monthly Sold Qty (P90D)": 10.0},
        # SKU 1: 同 SKU 在 Non Tourist 過剩（Turnday>30）
        {"SKU": "1", "Brand": "A", "Shop Nature": "Non Tourist", "Site Code": "N01",
         "Shop Stock Qty (CM)": 80, "Shop Monthly Sold Qty (P90D)": 20.0, "Turnday (P90D Qty)": 120},
        # SKU 2: Mix 過剩（Turnday>14）
        {"SKU": "2", "Brand": "B", "Shop Nature": "Mix", "Site Code": "M01",
         "Shop Stock Qty (CM)": 100, "Shop Monthly Sold Qty (P90D)": 30.0, "Turnday (P90D Qty)": 100},
        # SKU 2: 同 SKU Tourist 缺貨
        {"SKU": "2", "Brand": "B", "Shop Nature": "Tourist", "Site Code": "T02",
         "Shop Stock Qty (CM)": 0, "Sold Qty (M-2)": 8.0, "Shop Monthly Sold Qty (P90D)": 8.0},
    ])
    raw = _add_stock_impact_columns(df)
    inv = _build_inventory_distribution_summary(raw)

    # B1: Mix Turnday 分級表應有資料
    mix_td = inv["mix_turnday"]
    assert not mix_td.empty, "Mix Turnday 分級表應有資料"
    assert "可釋放件數" in mix_td.columns
    assert "總計" in mix_td["Turnday 區間"].values
    # 分級正確性驗證：至少包含一種分級 + 總計
    assert len(mix_td) >= 2, f"至少應含分級資料 + 總計，得 {len(mix_td)} 列"
    print(f"OK B1a: Mix Turnday 分級表（{len(mix_td)} 列）")

    # B1b: NT Turnday
    nt_td = inv["nt_turnday"]
    assert not nt_td.empty, "NT Turnday 分級表應有資料"
    assert "總計" in nt_td["Turnday 區間"].values
    print(f"OK B1b: Non Tourist Turnday 分級表（{len(nt_td)} 列）")

    # B2: Nature 對比表應有 3 個 Nature
    nat_cmp = inv["nature_compare"]
    assert not nat_cmp.empty
    assert set(nat_cmp["Shop Nature"]) == {"Tourist", "Mix", "Non Tourist"}
    print(f"OK B2: Nature 對比表（3 個 Nature）")

    # B3: 存貨失衡 SKU（SKU 1 & 2 應出現：Tourist缺貨+其他過剩）
    imb = inv["imbalance_sku"]
    assert not imb.empty, "應有存貨失衡 SKU"
    skus = set(imb["SKU"].astype(str))
    assert "1" in skus and "2" in skus, "SKU 1 & 2 應為存貨失衡"
    assert "建議調貨量" in imb.columns
    # 驗證註解屬性存在
    assert hasattr(imb, "attrs") and "annotation" in imb.attrs, "應有計算方式註解"
    print(f"OK B3: 存貨失衡 SKU（{len(imb)} 個，含建議調貨量 + 計算註解）")


    # B4: 低存貨 SKU 配置決策應存在
    la = inv.get("low_stock_allocation")
    assert la is not None and not la.empty, "應有低存貨 SKU 配置決策表"
    assert "NT_RF無銷售有存貨行數" in la.columns
    assert "ND無銷售浪費存貨(件)" in la.columns
    print(f"OK B4: 低存貨配置決策表（欄位 {list(la.columns)}）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 8: 空資料 / 無缺貨處理
# ──────────────────────────────────────────────────────────────────────────────
def test_no_impacted_items():
    df = _df([
        {"SKU": "1", "Shop Stock Qty (CM)": 5, "Sold Qty (M-2)": 10.0},
        {"SKU": "2", "Shop Stock Qty (CM)": 3, "Sold Qty (M-2)": 20.0},
    ])
    _, brand, buyer, cat, item_class, _, diag, inv, _, raw, warns = transform_stock_impact(df)
    assert int(raw["Is Stock Impacted"].sum()) == 0
    print("OK 無缺貨影響：空彙總")


# ──────────────────────────────────────────────────────────────────────────────
# Test 6b: 只計算 RF 店舖（ND 店舖不納入影響計算）
# ──────────────────────────────────────────────────────────────────────────────
def test_rf_only_filter():
    df = _df([
        # r0: RF 店舖缺貨 → 應納入影響
        {"SKU": "1", "RP Type": "RF", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 10.0, "Shop Monthly Sold Qty (P90D)": 10.0},
        # r1: ND 店舖缺貨（同樣有需求）→ 不應納入影響
        {"SKU": "2", "RP Type": "ND", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 20.0, "Shop Monthly Sold Qty (P90D)": 15.0},
        # r2: 髒資料 RP Type " rf "（含空白）→ strip+upper 後應視為 RF
        {"SKU": "3", "RP Type": " rf ", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 8.0, "Shop Monthly Sold Qty (P90D)": 8.0},
        # r3: 空白 RP Type → 非 RF，不應納入
        {"SKU": "4", "RP Type": "", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 30.0, "Shop Monthly Sold Qty (P90D)": 20.0},
    ])
    raw = _add_stock_impact_columns(df)

    assert raw.iloc[0]["Is Stock Impacted"] == True, "RF 缺貨應納入"
    assert raw.iloc[1]["Is Stock Impacted"] == False, "ND 缺貨不應納入"
    assert raw.iloc[2]["Is Stock Impacted"] == True, "髒資料 ' rf ' 應正規化為 RF"
    assert raw.iloc[3]["Is Stock Impacted"] == False, "空白 RP Type 不應納入"
    impacted_count = int(raw["Is Stock Impacted"].sum())
    assert impacted_count == 2, f"應只有 2 行影響（RF），得 {impacted_count}"
    print(f"OK RF 篩選：{impacted_count} 行（排除 ND/空白，含髒資料正規化）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 7: D001 狀態分類
# ──────────────────────────────────────────────────────────────────────────────
def test_classify_d001_status():
    assert _classify_d001_status(0) == "源頭斷貨"
    assert _classify_d001_status(300) == "倉庫有貨"
    print("OK _classify_d001_status（0=源頭斷貨, >0=補貨問題）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 7b: by-SKU 缺貨 Review 表（Tourist+Mix+Non Tourist, 含 D001 + 店舖貨量, v1.7.1 新增 Non Tourist / 取消缺貨行數）
# ──────────────────────────────────────────────────────────────────────────────
def test_sku_stock_diagnosis():
    df = _df([
        # SKU 1: Tourist 缺貨, D001=0 (源頭斷貨)
        {"SKU": "1", "Brand": "A", "Buyer": "B1", "SKU Desc": "Test Product A", "Shop Nature": "Tourist", "Site Code": "T01",
         "Shop Stock Qty (CM)": 0, "Shop Stock Qty (M-2)": 5,
         "Sold Qty (M-2)": 10.0, "Sold Qty (CM - MTD)": 2.0, "Shop Monthly Sold Qty (P90D)": 10.0,
         "D001 Stock Qty (M-2)": 100, "D001 Stock Qty (M-1)": 50, "D001 Stock Qty (CM)": 0},
        # SKU 1 同 SKU 在另一間 Tourist 店缺貨
        {"SKU": "1", "Brand": "A", "Buyer": "B1", "SKU Desc": "Test Product A", "Shop Nature": "Tourist", "Site Code": "T02",
         "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 8.0, "Sold Qty (CM - MTD)": 1.0, "Shop Monthly Sold Qty (P90D)": 8.0,
         "D001 Stock Qty (M-2)": 100, "D001 Stock Qty (M-1)": 50, "D001 Stock Qty (CM)": 0},
        # SKU 2: Mix 缺貨, D001>0 (補貨問題)
        {"SKU": "2", "Brand": "B", "Buyer": "B2", "SKU Desc": "Test Product B", "Shop Nature": "Mix", "Site Code": "M01",
         "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 20.0, "Sold Qty (CM - MTD)": 5.0, "Shop Monthly Sold Qty (P90D)": 15.0,
         "D001 Stock Qty (M-2)": 200, "D001 Stock Qty (M-1)": 200, "D001 Stock Qty (CM)": 150},
        # SKU 3: Non Tourist 缺貨 → v1.7.1 起納入 Review 表
        {"SKU": "3", "Brand": "C", "Buyer": "B3", "SKU Desc": "Test Product C", "Shop Nature": "Non Tourist", "Site Code": "N01",
         "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 30.0, "Sold Qty (CM - MTD)": 10.0, "Shop Monthly Sold Qty (P90D)": 20.0},
    ])
    raw = _add_stock_impact_columns(df)
    diag = _build_sku_stock_diagnosis(raw)

    assert not diag.empty, "應有缺貨 SKU"
    skus = set(diag["SKU"].astype(str))
    assert "1" in skus and "2" in skus and "3" in skus, "SKU 1, 2, 3 皆應在 Review 表（含 Non Tourist）"

    # SKU Desc 欄位應存在
    assert "SKU Desc" in diag.columns, "應有 SKU Desc 欄位"
    # NDRF 欄位應存在
    assert "NDRF_ND行數" in diag.columns, "應有 NDRF_ND行數 欄位"
    assert "NDRF_RF行數" in diag.columns, "應有 NDRF_RF行數 欄位"
    # Non Tourist 欄位應存在
    assert "Non Tourist缺貨店數" in diag.columns, "應有 Non Tourist缺貨店數 欄位"
    # 缺貨行數欄位應已取消
    assert "Tourist缺貨行數" not in diag.columns, "缺貨行數應已取消"
    assert "Mix缺貨行數" not in diag.columns, "Mix缺貨行數應已取消"

    # SKU 1: Tourist 缺貨 2 間, D001 CM=0 → 源頭斷貨
    row1 = diag[diag["SKU"].astype(str) == "1"].iloc[0]
    assert int(row1["Tourist缺貨店數"]) == 2, f"SKU1 應有 2 間 Tourist 缺貨, 得 {row1['Tourist缺貨店數']}"
    assert int(row1["D001_Stock_CM"]) == 0
    assert row1["D001 狀態"] == "源頭斷貨"
    assert int(row1["D001_Stock_M2"]) == 100
    assert int(row1["D001_Stock_M1"]) == 50
    assert int(row1["NDRF_RF行數"]) == 2
    print("OK SKU 1: Tourist 缺貨 2 間, 源頭斷貨, NDRF_RF=2")

    # SKU 2: Mix 缺貨, D001 CM=150 → 補貨問題
    row2 = diag[diag["SKU"].astype(str) == "2"].iloc[0]
    assert int(row2["Mix缺貨店數"]) == 1
    assert int(row2["D001_Stock_CM"]) == 150
    assert row2["D001 狀態"] == "倉庫有貨"
    print("OK SKU 2: Mix 缺貨, 倉庫有貨（D001=150 但店舖缺貨）")

    # SKU 3: Non Tourist 缺貨（v1.7.1 新增）
    row3 = diag[diag["SKU"].astype(str) == "3"].iloc[0]
    assert int(row3["Non Tourist缺貨店數"]) == 1
    print("OK SKU 3: Non Tourist 缺貨 1 間（v1.7.1 起納入 Review）")

    # 排序: Tourist 缺貨店數多的在前 → SKU 1 (2間) 應在最前
    assert str(diag.iloc[0]["SKU"]) == "1"
    print("OK by-SKU 缺貨 Review 表（Tourist+Mix+Non Tourist, 含 SKU Desc + NDRF 行數, 取消缺貨行數）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 9: Excel 生成
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_excel_stock_impact():
    df = _df([
        {"SKU": "1", "Adj. MCH1": "CatA", "Shop Nature": "Tourist", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 10.0, "Sold Qty (CM - MTD)": 2.0, "Shop Monthly Sold Qty (P90D)": 10.0},
        {"SKU": "2", "Adj. MCH1": "CatB", "Shop Nature": "Mix", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 20.0, "Sold Qty (CM - MTD)": 5.0, "Shop Monthly Sold Qty (P90D)": 15.0},
    ])
    _, brand, buyer, cat, item_class, _, diag, inv, sku_monthly, raw, _ = transform_stock_impact(df)
    buf = generate_excel_stock_impact(brand, buyer, cat, item_class, diag, inv, raw, sku_monthly_analysis=sku_monthly)
    assert isinstance(buf, BytesIO)
    data = buf.getvalue()
    assert len(data) > 100
    assert data[:2] == b"PK", "應為有效 xlsx（PK zip signature）"

    # 驗證 10 個 sheet（v1.8.0 新增 SKU Monthly Sales&Store）
    import openpyxl
    wb = openpyxl.load_workbook(BytesIO(data))
    expected = ["RAW", "OOS Impact Detail", "By Brand", "By Buyer",
                "By Category", "By Item Class", "SKU Monthly Sales&Store",
                "SKU Stock Diagnosis", "Inventory Distribution", "Stock Impact Dashboard"]
    assert wb.sheetnames == expected, f"sheets mismatch: {wb.sheetnames}"
    print("OK generate_excel_stock_impact（9 sheets valid xlsx）")


# ──────────────────────────────────────────────────────────────────────────────
# Test: low_stock_item_class_summary
# ──────────────────────────────────────────────────────────────────────────────
def test_build_low_stock_item_class_summary():
    # Normal case: mix of low-stock / non-low-stock, multiple Item Classes
    df = pd.DataFrame({
        "SKU": ["S1", "S2", "S3", "S4", "S5", "S6"],
        "低存貨": [True, True, False, True, True, False],
        "Item Class": ["Cream", "Lotion", "Cream", "Cream", None, "Lotion"],
    })
    out = _build_low_stock_item_class_summary(df)
    assert not out.empty, "should return data for valid input"
    assert list(out.columns) == ["Item Class", "低存貨SKU隻數"], f"unexpected columns: {out.columns}"
    # Cream: S1, S4 → 2; Lotion: S2 → 1; New SKU: S5 → 1
    cream_row = out[out["Item Class"] == "Cream"]
    assert int(cream_row["低存貨SKU隻數"].iloc[0]) == 2, f"Cream expected 2, got {cream_row['低存貨SKU隻數'].iloc[0]}"
    lotion_row = out[out["Item Class"] == "Lotion"]
    assert int(lotion_row["低存貨SKU隻數"].iloc[0]) == 1
    newsku_row = out[out["Item Class"] == "New SKU"]
    assert int(newsku_row["低存貨SKU隻數"].iloc[0]) == 1
    total_row = out[out["Item Class"] == "總計"]
    assert int(total_row["低存貨SKU隻數"].iloc[0]) == 4, f"total expected 4, got {total_row['低存貨SKU隻數'].iloc[0]}"
    # Sorted desc: Cream(2) first, then Lotion(1) or New SKU(1), then 總計 last
    assert out.iloc[0]["Item Class"] == "Cream", f"first should be Cream, got {out.iloc[0]['Item Class']}"
    assert out.iloc[-1]["Item Class"] == "總計", f"last should be 總計, got {out.iloc[-1]['Item Class']}"
    print("OK _build_low_stock_item_class_summary - normal case")

    # NaN Item Class → New SKU
    df_nan = pd.DataFrame({
        "SKU": ["X1", "X2"],
        "低存貨": [True, True],
        "Item Class": [np.nan, ""],
    })
    out_nan = _build_low_stock_item_class_summary(df_nan)
    assert len(out_nan) == 2
    assert out_nan.iloc[0]["Item Class"] == "New SKU"
    assert int(out_nan[out_nan["Item Class"] == "總計"]["低存貨SKU隻數"].iloc[0]) == 2
    print("OK _build_low_stock_item_class_summary - NaN/blank → New SKU")

    # All non-low-stock
    df_empty = pd.DataFrame({
        "SKU": ["Y1", "Y2"],
        "低存貨": [False, False],
        "Item Class": ["Cream", "Lotion"],
    })
    assert _build_low_stock_item_class_summary(df_empty).empty
    print("OK _build_low_stock_item_class_summary - no low stock → empty")

    # Missing columns
    assert _build_low_stock_item_class_summary(pd.DataFrame({"SKU": ["Z"]})).empty
    print("OK _build_low_stock_item_class_summary - missing columns → empty")

    # None/empty input
    assert _build_low_stock_item_class_summary(None).empty
    assert _build_low_stock_item_class_summary(pd.DataFrame()).empty
    print("OK _build_low_stock_item_class_summary - None/empty → empty")

    print("ALL _build_low_stock_item_class_summary TESTS PASSED")


# ──────────────────────────────────────────────────────────────────────────────
# Test 10: _build_sku_soh_tier_summary marker 修正（有貨>0 & <50% → Y 其餘 ""）
# ──────────────────────────────────────────────────────────────────────────────
def test_sku_soh_tier_summary_marker():
    raw = pd.DataFrame({
        "SKU": ["S1", "S1", "S2", "S2", "S3", "S3", "S4", "S4"],
        "SKU Desc": ["Prod A", "Prod A", "Prod B", "Prod B", "Prod C", "Prod C", "Prod D", "Prod D"],
        "Brand": ["A", "A", "A", "A", "B", "B", "B", "B"],
        "Item Class": ["Cream", "Cream", "Lotion", "Lotion", "Cream", "Cream", "Lotion", "Lotion"],
        "Shop Stock Qty (CM)": [5, 5, 3, 3, 0, 0, 10, 10],
        "Sold Qty (CM - MTD)": [1, 1, 0, 0, 0, 0, 8, 8],
        "Shop Monthly Sold Qty (P90D)": [10, 10, 10, 10, 5, 5, 10, 10],
        "Shop Stock Qty (M-2)": [5, 5, 3, 3, 0, 0, 10, 10],
        "Shop Stock Qty (M-1)": [5, 5, 3, 3, 0, 0, 10, 10],
        "Sold Qty (M-2)": [1, 1, 0, 0, 0, 0, 8, 8],
        "Sold Qty (M-1)": [1, 1, 0, 0, 0, 0, 8, 8],
        "Supply Source": ["Warehouse", "Warehouse", "Warehouse", "Warehouse", "Warehouse", "Warehouse", "Warehouse", "Warehouse"],
    }, dtype=object)
    # Enrich 需要特定欄位，補上必要欄
    raw["Shop Nature"] = "Tourist"
    raw["RP Type"] = "RF"
    raw["OOS Periods Count"] = 0
    raw["OOS Pattern"] = "無缺貨"
    raw["Is Stock Impacted"] = False
    raw["Is OOS (M-2)"] = False
    raw["Is OOS (M-1)"] = False
    raw["Is OOS (CM)"] = False
    raw["D001 Stock Qty (M-2)"] = 0
    raw["D001 Stock Qty (M-1)"] = 0
    raw["D001 Stock Qty (CM)"] = 0

    new_df = _build_sku_soh_tier_summary(raw)
    assert not new_df.empty, "should produce output"
    marker_col = "Low 有賣VS有貨 (CM<50%)"
    assert marker_col in new_df.columns, f"missing {marker_col}"

    # S1: 有貨 2 間 >0，Sold=1/2間 → 有賣VS有貨=0.5(≥50%) → marker should be ""
    s1 = new_df[new_df["SKU"].astype(str) == "S1"].iloc[0]
    assert str(s1[marker_col]) == "", f"S1 (售罄率>=50%) should be '', got '{s1[marker_col]}'"
    print(f"OK S1: 有貨 2 間、有賣 2 間 → 比率≥50% → marker=''")

    # S2: 有貨 2 間 >0，Sold=0 (0%) → <50% → marker should be "Y"
    s2 = new_df[new_df["SKU"].astype(str) == "S2"].iloc[0]
    assert str(s2[marker_col]) == "Y", f"S2 (售罄率0%) should be 'Y', got '{s2[marker_col]}'"
    print(f"OK S2: 有貨 2 間、有賣 0 間 → 比率=0% → marker='Y'")

    # S3: 有貨=0 → ratio NaN → marker should be ""
    s3 = new_df[new_df["SKU"].astype(str) == "S3"].iloc[0]
    assert str(s3[marker_col]) == "", f"S3 (無貨) should be '', got '{s3[marker_col]}'"
    print(f"OK S3: 有貨=0 間 → 比率空白 → marker=''")

    # S4: 有貨 2 間，有賣 2 間 → 比率=1.0 (100%) ≥50% → marker should be ""
    s4 = new_df[new_df["SKU"].astype(str) == "S4"].iloc[0]
    assert str(s4[marker_col]) == "", f"S4 (售罄率100%) should be '', got '{s4[marker_col]}'"
    print(f"OK S4: 有貨 2 間、有賣 2 間 → 比率=100% → marker=''")

    print("OK _build_sku_soh_tier_summary marker（新準則，有貨>0 & <50% → Y）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 11: _build_sell_through_band_summary
# ──────────────────────────────────────────────────────────────────────────────
def test_build_sell_through_band_summary():
    raw = pd.DataFrame({
        "SKU": ["S1", "S1", "S2", "S2", "S3", "S3"],
        "SKU Desc": ["P1", "P1", "P2", "P2", "P3", "P3"],
        "Brand": ["A", "A", "A", "A", "B", "B"],
        "Item Class": ["Cream", "Cream", "Lotion", "Lotion", "Cream", "Cream"],
        "Shop Stock Qty (CM)": [5, 5, 3, 3, 0, 0],
        "Sold Qty (CM - MTD)": [1, 1, 0, 0, 0, 0],
        "Shop Monthly Sold Qty (P90D)": [10, 10, 10, 10, 5, 5],
        "Shop Stock Qty (M-2)": [5, 5, 3, 3, 0, 0],
        "Shop Stock Qty (M-1)": [5, 5, 3, 3, 0, 0],
        "Sold Qty (M-2)": [1, 1, 0, 0, 0, 0],
        "Sold Qty (M-1)": [1, 1, 0, 0, 0, 0],
        "Supply Source": ["Warehouse"] * 6,
        "Shop Nature": ["Tourist"] * 6,
        "RP Type": ["RF"] * 6,
        "OOS Periods Count": [0] * 6,
        "OOS Pattern": ["無缺貨"] * 6,
        "Is Stock Impacted": [False] * 6,
        "Is OOS (M-2)": [False] * 6,
        "Is OOS (M-1)": [False] * 6,
        "Is OOS (CM)": [False] * 6,
        "D001 Stock Qty (M-2)": [0] * 6,
        "D001 Stock Qty (M-1)": [0] * 6,
        "D001 Stock Qty (CM)": [0] * 6,
    }, dtype=object)
    new_df = _build_sku_soh_tier_summary(raw)
    band = _build_sell_through_band_summary(new_df)

    assert not band.empty, "should produce band table"
    expected_cols = ["有賣VS有貨比率區間", "SKU數 (M-2)", "SKU數 (M-1)", "SKU數 (CM)"]
    assert list(band.columns) == expected_cols, f"cols mismatch: {list(band.columns)}"

    # 總計 row 每期 = 3 (S1 + S2 + S3)
    total = band[band["有賣VS有貨比率區間"] == "總計"]
    assert len(total) == 1
    for p in ["M-2", "M-1", "CM"]:
        assert int(total[f"SKU數 ({p})"].iloc[0]) == 3, f"total {p} should be 3"

    # 驗證各期無貨 row = 1 (S3)
    no_stock_row = band[band["有賣VS有貨比率區間"] == "無貨（有貨=0）"]
    assert len(no_stock_row) == 1
    for p in ["M-2", "M-1", "CM"]:
        assert int(no_stock_row[f"SKU數 ({p})"].iloc[0]) == 1, f"無貨 {p} should be 1"

    # 空 input
    assert _build_sell_through_band_summary(pd.DataFrame()).empty
    assert _build_sell_through_band_summary(pd.DataFrame({"SKU": ["Z"]})).empty
    print("OK _build_sell_through_band_summary（行/列/總計/空資料）")


# ──────────────────────────────────────────────────────────────────────────────
# Test 12: 簡化版 Excel 生成 smoke test
# ──────────────────────────────────────────────────────────────────────────────
def test_generate_excel_stock_impact_simplified():
    df = _df([
        {"SKU": "1", "Shop Nature": "Tourist", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 10.0, "Sold Qty (CM - MTD)": 2.0, "Shop Monthly Sold Qty (P90D)": 10.0},
        {"SKU": "2", "Shop Nature": "Mix", "Shop Stock Qty (CM)": 0,
         "Sold Qty (M-2)": 20.0, "Sold Qty (CM - MTD)": 5.0, "Shop Monthly Sold Qty (P90D)": 15.0},
    ])
    _, _, _, _, _, _, _, _, sku_monthly, raw, _ = transform_stock_impact(df)
    buf = generate_excel_stock_impact_simplified(raw, sku_monthly_analysis=sku_monthly)
    assert isinstance(buf, BytesIO)
    data = buf.getvalue()
    assert len(data) > 100
    assert data[:2] == b"PK", "should be valid xlsx"

    import openpyxl
    wb = openpyxl.load_workbook(BytesIO(data))
    expected = ["RAW", "SKU Monthly Sales&Store", "New", "Stock Impact Dashboard"]
    assert wb.sheetnames == expected, f"sheets mismatch: {wb.sheetnames}"

    # Verify New sheet has marker column
    ws_new = wb["New"]
    header_row = [c.value for c in ws_new[4]]
    assert "Low 有賣VS有貨 (CM<50%)" in header_row, f"missing marker col: {header_row}"

    # Verify Dashboard sheet has title
    ws_dash = wb["Stock Impact Dashboard"]
    title = ws_dash.cell(row=1, column=1).value
    assert title and "Stock Impact Dashboard" in str(title), f"unexpected title: {title}"
    print("OK generate_excel_stock_impact_simplified（4 sheets, marker col present）")


if __name__ == "__main__":
    test_oos_detection()
    test_sales_loss_calculation()
    test_oos_pattern_classification()
    test_build_stock_impact_summary()
    test_classify_d001_status()
    test_sku_stock_diagnosis()
    test_build_low_stock_item_class_summary()
    test_transform_stock_impact()
    test_warehouse_filter()
    test_no_impacted_items()
    test_rf_only_filter()
    test_mix_turnday_banding()
    test_inventory_distribution_summary()
    test_generate_excel_stock_impact()
    test_sku_soh_tier_summary_marker()
    test_build_sell_through_band_summary()
    test_generate_excel_stock_impact_simplified()
    print("\nALL STOCK IMPACT VALIDATION TESTS PASSED")
