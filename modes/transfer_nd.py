# -*- coding: utf-8 -*-
"""ND SKU with Stock Review(轉貨模式)：指定轉出店舖（RP Type＝ND 且有庫存）轉貨至有銷售記錄的 RF 店舖。

流程與 P90D 0 Sales Review(轉貨模式) 相同，只把轉出條件由「0 Sales」換成「RP Type＝ND」，
並把接收店收窄為 RP Type＝RF：
1. ``build_nd_transfer_candidates``：列出候選轉出店舖（RP Type＝ND 且 Shop Stock Qty (CM) > 0），
   供介面勾選指定轉出店舖。
2. ``transform_nd_transfer``：依勾選的轉出店舖產生轉貨單與四張彙總表。
3. ``generate_excel_nd_transfer``：產生六個工作表（說明、轉貨單、轉出店舖彙總、接收店舖彙總、
   SKU 轉貨彙總、MCH1 分類彙總；無 RAW）。

規則（不分 OM 群組，轉貨數量不設上限）：
- 轉出：已選轉出店舖中 RP Type＝ND 且有庫存的列，轉出數量為 Shop Stock Qty (CM) 全數
  （不看銷售記錄，即使該列仍有銷售亦照轉）。
- 接收：同 SKU、非轉出店舖、RP Type＝RF 且 Sold Qty (P90D) > 0 的列（有銷售記錄），不論有無庫存。
- 分配：同 SKU 的轉出總量按接收店 Sold Qty (P90D) 比例分配，以最大餘數法取整數件、全數分完；
  只有一間接收店時全數給該店；無接收店時該 SKU 不轉出，列入未分配。
- RP Type 比對為去除前後空格、大小寫不敏感；``RP Type`` 為必要欄位，缺欄即停止計算，
  空白或其他非 RF／ND 的值一律不納入轉出與接收（列 warning）。
- MCH 篩選（選填）：執行前可按 Adj. MCH1 排除部分分類（傳 None 或包含全部值＝不排除），
  對轉出與接收兩側同時生效；報表另附 MCH1 分類彙總（按 Adj. MCH1）。

分配、配對與彙總沿用 modes/transfer_p90d.py 的同一引擎（``nd_mode=True``），本模組只提供對外介面。
"""
from modes.transfer_p90d import (
    build_mch_options,
    build_transfer_candidates,
    generate_excel_p90d_transfer,
    transform_p90d_transfer,
)

__all__ = [
    "build_nd_transfer_candidates",
    "build_mch_options",
    "transform_nd_transfer",
    "generate_excel_nd_transfer",
]


def build_nd_transfer_candidates(df_input, shop_natures: list[str] | None = None, mch1_filter=None):
    """候選轉出店舖一覽：RP Type＝ND 且 Shop Stock Qty (CM) > 0 的列，依店舖彙總。

    ``mch1_filter`` 為執行前的 MCH1 篩選（None 或包含全部值＝不排除）。
    """
    return build_transfer_candidates(
        df_input, shop_natures=shop_natures, mch1_filter=mch1_filter, nd_mode=True,
    )


def transform_nd_transfer(
    df_input,
    transfer_out_sites: list[str],
    progress_callback=None,
    shop_natures: list[str] | None = None,
    mch1_filter=None,
):
    """依指定轉出店舖產生轉貨計畫（RP Type＝ND 且有庫存 → RP Type＝RF 且有銷售記錄的店舖）。

    ``mch1_filter`` 為執行前的 MCH1 篩選（None 或包含全部值＝不排除），
    對轉出與接收兩側同時生效。

    回傳 (plan_df, out_shop_df, in_shop_df, sku_df, mch_df, stats, warnings)。
    """
    return transform_p90d_transfer(
        df_input,
        transfer_out_sites,
        progress_callback=progress_callback,
        shop_natures=shop_natures,
        mch1_filter=mch1_filter,
        nd_mode=True,
    )


def generate_excel_nd_transfer(
    plan_df,
    out_shop_df,
    in_shop_df,
    sku_df,
    mch_df=None,
    stats: dict | None = None,
    nature_label: str = "全部",
    mch_filter_label: str = "全部（未排除）",
    progress_callback=None,
):
    """產生六個工作表：說明、轉貨單、轉出店舖彙總、接收店舖彙總、SKU 轉貨彙總、MCH1 分類彙總（無 RAW）。"""
    return generate_excel_p90d_transfer(
        plan_df,
        out_shop_df,
        in_shop_df,
        sku_df,
        mch_df=mch_df,
        stats=stats,
        nature_label=nature_label,
        mch_filter_label=mch_filter_label,
        progress_callback=progress_callback,
        nd_mode=True,
    )
