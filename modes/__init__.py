# -*- coding: utf-8 -*-
"""分析模式註冊表：模式顯示名稱 → 模組與函數對照。

新模式只需在 modes/ 下新增一個模組並在此註冊，無需改動 transformer.py
墊片與舊有分派邏輯。函數以延遲字串記載，避免循環引用。
"""

MODE_REGISTRY = {
    "Past 90 Days 0 Sales": {
        "module": "modes.p90d",
        "transform": "transform_p90d_zero_sales",
        "generate_excel": "generate_excel_p90d_zero_sales",
    },
}

MODE_REGISTRY["OOS Review"] = {
    "module": "modes.oos",
    "transform": "transform_oos",
    "generate_excel": "generate_excel_oos",
}

MODE_REGISTRY["SKU 精簡 Review"] = {
    "module": "modes.nd",
    "transform": "transform_nd",
    "generate_excel": "generate_excel_nd",
}

MODE_REGISTRY["Overstock Review"] = {
    "module": "modes.overstock",
    "transform": "transform",
    "generate_excel": "generate_excel",
}

MODE_REGISTRY["Carry SKU Review"] = {
    "module": "modes.carry",
    "transform": "transform_carry",
    "generate_excel": "generate_excel_carry",
}

MODE_REGISTRY["NDRF 月份對比"] = {
    "module": "modes.ndrf",
    "transform": "transform_ndrf_comparison",
    "generate_excel": "generate_excel_ndrf_comparison",
}

MODE_REGISTRY["Stock Checking"] = {
    "module": "modes.stock_checking",
    "transform": "transform_stock_checking",
    "generate_excel": "generate_excel_stock_checking",
}

MODE_REGISTRY["Safety Stock Review"] = {
    "module": "modes.safety_review",
    "transform": "transform_combined",
    "generate_excel": "generate_excel_combined",
}

MODE_REGISTRY["Stock Impact Analysis"] = {
    "module": "modes.stock_impact",
    "transform": "transform_stock_impact",
    "generate_excel": "generate_excel_stock_impact",
}

MODE_REGISTRY["Safety Stock x NetStock Review"] = {
    "module": "modes.hybrid",
    "transform": "transform_hybrid",
    "generate_excel": "generate_excel_hybrid",
}

MODE_REGISTRY["Forecast Review"] = {
    "module": "modes.forecast",
    "transform": "transform_forecast",
    "generate_excel": "generate_excel_forecast",
}

MODE_REGISTRY["Stock Impact Analysis (簡化)"] = {
    "module": "modes.stock_impact",
    "transform": "transform_stock_impact",
    "generate_excel": "generate_excel_stock_impact_simplified",
}

MODE_REGISTRY["P90D 0 Sales Review(Shop回覆)"] = {
    "module": "modes.overstock_p90d",
    "transform": "transform_overstock_p90d",
    "generate_excel": "generate_excel_overstock_p90d",
}

MODE_REGISTRY["P90D 0 Sales Review(個別店舖)"] = {
    "module": "modes.overstock_p90d",
    "transform": "transform_overstock_p90d_shop",
    "generate_excel": "generate_excel_overstock_p90d_shop",
}

MODE_REGISTRY["Safety Stock 月度變化分析"] = {
    "module": "modes.safety_monthly",
    "transform": "transform_safety_monthly",
    "generate_excel": "generate_excel_safety_monthly",
}

MODE_REGISTRY["P90D 0 Sales Review(轉貨模式)"] = {
    "module": "modes.transfer_p90d",
    "transform": "transform_p90d_transfer",
    "generate_excel": "generate_excel_p90d_transfer",
}
