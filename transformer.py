import pandas as pd
import numpy as np
import unicodedata
import re
import gc
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import datetime
from concurrent.futures import ThreadPoolExecutor





















































# ─── Safety Stock Review Functions ─────────────────────────────────────────────










# ── SUMMARY sheet：兩大分類統計的固定顯示順序與說明 ──────────────────────────















# ─── Forecast Review（SKU 級建議訂貨量；第二維度防缺貨）Functions ──────────────
# 以下函數集中置於檔尾 Stock Checking 區之後，加 `forecast_`／`_fc_` 字首，
# 避免與現有函數名碰撞；不改動既有 11 個模式任何計算、欄位、欄序。













# ─── 店舖問題 SKU（每間店舖需行動列：問題＋建議方法）────────────────────────






































# ─── ND SKU Review (SKU 精簡 Review) Functions ─────────────────────────────────












# ─── OOS Review Functions ──────────────────────────────────────────────────────













# ─── Carry SKU Review Functions ─────────────────────────────────────────────────













# ─── NDRF Month Comparison Functions ─────────────────────────────────────────────

















# ─── Stock Impact Analysis Functions ────────────────────────────────────────────
#
# 取代 Sales Decline Detection (v1.5.1)。改以 Shop Stock = 0（缺貨）對 Sold Qty 的
# 因果影響為核心，回答「無貨就無銷售」。Shop Nature 優先級 Tourist > Mix > Non Tourist。


# Shop Nature 排序權重（越小越優先）





























































# ─── Safety Stock x NetStock Review (調貨事件 Review Safety 設定) Functions ──────























































_HYBRID_P90D_ZERO_SALES_ANNOTATION = "另符合 Past 90 Days 0 Sales，建議檢視是否轉 ND 停補貨"












































































# ─── Stock Impact Analysis (Simplified) Functions ─────────────────────────────
# 簡化版本：只輸出 RAW 和 SKU Monthly Sales&Store 兩個工作表
# SKU Monthly Sales&Store 不分 Shop Nature，只顯示 M-2, M-1, MTD 的 D001 存貨、店舖存貨和總銷售量













# ─── Past 90 Days 0 Sales Functions ─────────────────────────────────────────────
# P1 拆分試點：實作已遷至 modes/p90d.py，此處僅保留相容 re-export，
# 舊有 from transformer import ... 路徑維持不變。
from modes.p90d import (
    _P90D_ZERO_SALES_REQUIRED,
    _P90D_RAW_OUTPUT_COLUMNS,
    _p90d_raw_columns,
    _build_p90d_zero_sales_by_shop_summary,
    transform_p90d_zero_sales,
    generate_excel_p90d_zero_sales,
)



























































# ─── OOS Review：實作已遷至 modes/oos.py（墊片 re-export）───
from modes.oos import *  # noqa: F401,F403


# ─── SKU 精簡 Review：實作已遷至 modes/nd.py（墊片 re-export）───
from modes.nd import *  # noqa: F401,F403


# ─── Overstock Review：實作已遷至 modes/overstock.py（墊片 re-export）───
from modes.overstock import *  # noqa: F401,F403


# ─── Carry SKU Review：實作已遷至 modes/carry.py（墊片 re-export）───
from modes.carry import *  # noqa: F401,F403


# ─── NDRF 月份對比：實作已遷至 modes/ndrf.py（墊片 re-export）───
from modes.ndrf import *  # noqa: F401,F403


# ─── Stock Checking：實作已遷至 modes/stock_checking.py（墊片 re-export）───
from modes.stock_checking import *  # noqa: F401,F403


# ─── Safety Stock Review：實作已遷至 modes/safety_review.py（墊片 re-export）───
from modes.safety_review import *  # noqa: F401,F403



# ─── 共用核心墊片（實作已遷至 modes/common／issues／framework）───
from modes.common import *  # noqa: F401,F403
from modes.issues import *  # noqa: F401,F403
from modes.framework import *  # noqa: F401,F403


# ─── Stock Impact Analysis：實作已遷至 modes/stock_impact.py（墊片 re-export）───
from modes.stock_impact import *  # noqa: F401,F403


# ─── Safety Stock x NetStock Review：實作已遷至 modes/hybrid.py（墊片 re-export）───
from modes.hybrid import *  # noqa: F401,F403


# ─── Forecast Review：實作已遷至 modes/forecast.py（墊片 re-export）───
from modes.forecast import *  # noqa: F401,F403


# ─── 銷售損失量化 Review：實作已遷至 modes/sales_loss.py（墊片 re-export）───
from modes.sales_loss import *  # noqa: F401,F403


# ─── 積壓資金 Review：實作已遷至 modes/excess_capital.py（墊片 re-export）───
from modes.excess_capital import *  # noqa: F401,F403


# ─── RP Type 覆核 Review：實作已遷至 modes/rp_review.py（墊片 re-export）───
from modes.rp_review import *  # noqa: F401,F403


# ─── D001 健康檢查 Review：實作已遷至 modes/d001_health.py（墊片 re-export）───
from modes.d001_health import *  # noqa: F401,F403
