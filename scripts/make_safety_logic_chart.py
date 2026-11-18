# -*- coding: utf-8 -*-
"""Safety Stock 計算邏輯圖（供使用說明頁與對外講解用；非應用程式元件）。

準則與 modes/safety_review.py 的 `_add_combined_recommendation_columns` 完全一致：
- 月銷預估 = MTD 實績 + (1 − w) × P90D 月均，w = MTD 日數 ÷ 當月日數
- X0 = floor(P90D 月均 × Overstock Rate)
- Y  = ceil(月銷預估 × OOS Rate)
- X 欄 = min(X0, Z)（即 RAW 的「Overstock Cap (X)」欄）
- 建議值 = max(Y, X 欄) —— 兩步：先壓上限、再同底線取較大
- 周轉天數基準 = ceil(Sold Qty (P90D) ÷ 有效日數 × 基準天數)，
  有效日數 = min(90, max(上市至參考日（含首尾）, 3))
- 保底 3 件只適用於「有銷量或 90 日內上市新品」；舊品零銷量維持 0
圖片輸出：assets/safety_logic_guide.png
"""
import os

import matplotlib

matplotlib.use("Agg")
from matplotlib import font_manager
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

for _f in ("msjh.ttc", "msjhbd.ttc"):
    _p = os.path.join(r"C:\Windows\Fonts", _f)
    if os.path.exists(_p):
        font_manager.fontManager.addfont(_p)
plt.rcParams["font.family"] = ["Microsoft JhengHei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False

INK = "#1A1A2E"
SUB = "#5A6472"
LINE = "#D8DEE6"
TEAL = "#0077A3"
TEAL_BG = "#E3F1F7"
GREEN = "#22885A"
GREEN_BG = "#E4F3EB"
ORANGE = "#C05A11"
ORANGE_BG = "#FBF0E0"
GREY = "#5B6470"
HEAD_BG = "#E4EEF5"
ZEBRA = "#F6F9FB"

FIG_W, FIG_H = 12.5, 15.0
fig = plt.figure(figsize=(FIG_W, FIG_H), dpi=140)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
fig.patch.set_facecolor("white")


def T(x, y, s, size=10, color=INK, weight="normal", ha="left", va="center", ls=1.5):
    ax.text(x, y, s, fontsize=size, color=color, fontweight=weight, ha=ha, va=va,
            transform=ax.transAxes, linespacing=ls)


def box(x, y, w, h, fc, ec="none", lw=0.9, r=0.006):
    ax.add_patch(FancyBboxPatch((x, y), w, h,
                                boxstyle=f"round,pad=0,rounding_size={r}",
                                facecolor=fc, edgecolor=ec, linewidth=lw,
                                transform=ax.transAxes))


def arrow(x1, y1, x2, y2, color="#8A94A2", lw=1.3):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>",
                                 mutation_scale=12, color=color, lw=lw,
                                 transform=ax.transAxes, shrinkA=0, shrinkB=0))


def chip(x, y, num, color=TEAL, size=0.019):
    box(x, y - size / 2, size, size, color, r=0.004)
    T(x + size / 2, y, str(num), size=10.2, color="white", weight="bold", ha="center")


def section(x, y, num, title, color=TEAL):
    chip(x, y, num, color)
    T(x + 0.028, y, title, size=12.5, color=INK, weight="bold")


def step_band(y, h, num, title, formula, lines, accent, accent_bg):
    """一步一條橫幅：左邊序號＋標題＋公式，右邊兩三行白話說明。"""
    box(0.05, y, 0.90, h, accent_bg)
    box(0.05, y, 0.006, h, accent, r=0.003)
    chip(0.066, y + h - 0.028, num, accent)
    T(0.096, y + h - 0.028, title, size=10.6, color=INK, weight="bold")
    T(0.096, y + 0.026, formula, size=12.4, color=INK, weight="bold")
    _ly = y + h - 0.032
    for _ln in lines:
        T(0.520, _ly, _ln, size=8.6, color=SUB)
        _ly -= 0.016


# ── 標題 ──
T(0.05, 0.974, "Safety Stock 計算邏輯", size=25, weight="bold")
T(0.05, 0.947,
  "Safety Stock Review｜計算單位＝一間店舖 × 一個 SKU（報表一行）。系統用銷量算出上限 X0 與底線 Y，再以兩者為界調整你現時的設定 Z。",
  size=10.6, color=SUB)

# ── 一句話 ──
box(0.05, 0.884, 0.90, 0.042, HEAD_BG)
box(0.05, 0.884, 0.006, 0.042, TEAL, r=0.003)
T(0.068, 0.905,
  "一句話：先將現時設定壓到不超過上限（得出 X 欄），再同缺貨底線比較——哪個大，就用哪個。",
  size=10.8, color=INK, weight="bold")

# ── ① 三個輸入 ──
section(0.05, 0.862, "1", "讀入三個數字（每個店舖 × SKU 一行）")

_b1 = 0.756
box(0.05, _b1, 0.286, 0.094, TEAL_BG)
T(0.064, _b1 + 0.077, "銷售數據（來源檔）", size=10.8, color=TEAL, weight="bold")
T(0.064, _b1 + 0.048,
  "Sold Qty (P90D)：過去 90 日累計\n"
  "Shop Monthly Sold Qty (P90D)：月均\n"
  "Sold Qty (CM - MTD)：本月至今", size=9.2, color=INK)
T(0.064, _b1 + 0.010, "單位：件；P90D＝Past 90 Days（過去 90 日）", size=7.8, color=SUB)

box(0.357, _b1, 0.286, 0.094, "#EAEDF1")
T(0.371, _b1 + 0.077, "現時設定 Z", size=10.8, color=GREY, weight="bold")
T(0.371, _b1 + 0.051,
  "Safety Stock (現時)\n"
  "＝來源檔的 Adj. Safety Stock\n"
  "（你現在為這間店設幾件）", size=9.2, color=INK)
T(0.371, _b1 + 0.010, "空白或非數值 → 視為 0 並提示", size=7.8, color=SUB)

box(0.664, _b1, 0.286, 0.094, ORANGE_BG)
T(0.678, _b1 + 0.077, "店舖與品項屬性", size=10.8, color=ORANGE, weight="bold")
T(0.678, _b1 + 0.048,
  "Shop Nature：決定用哪組倍率\n"
  "Launch Date：90 日內上市＝新品\n"
  "RP Type：RF（可補貨）／ND（不可）", size=9.2, color=INK)
T(0.678, _b1 + 0.010, "其他 Nature 值一律按 Non Tourist 倍率", size=7.8, color=SUB)

arrow(0.5, 0.754, 0.5, 0.736)

# ── ② 月銷預估 ──
_y2, _h2 = 0.660, 0.072
box(0.05, _y2, 0.90, _h2, TEAL_BG)
box(0.05, _y2, 0.006, _h2, TEAL, r=0.003)
chip(0.066, _y2 + 0.048, "2", TEAL)
T(0.096, _y2 + 0.048, "先算月銷預估（Projected Monthly Sold Qty）", size=10.8, color=TEAL, weight="bold")
T(0.096, _y2 + 0.025,
  "月銷預估 = MTD 實績 + 剩餘日數 × P90D 月均　（權重 w = MTD 日數 ÷ 當月日數）", size=10.2, color=INK)
T(0.096, _y2 + 0.008,
  "月初以穩定的 P90D 為主，隨月份推進收斂至 MTD；輸出取整、下限 0。", size=8.6, color=SUB)

arrow(0.5, 0.658, 0.5, 0.641)

# ── ③ 兩條界線 ──
section(0.05, 0.622, "3", "用銷量算出兩條界線（上限與底線）")

_y3, _h3 = 0.504, 0.096
box(0.05, _y3, 0.440, _h3, ORANGE_BG)
T(0.066, _y3 + 0.078, "上限 X0（Overstock Cap）", size=10.8, color=ORANGE, weight="bold")
T(0.066, _y3 + 0.052, "X0 = floor（P90D 月均 × Overstock Rate）", size=10.4, color=INK, weight="bold")
T(0.066, _y3 + 0.024,
  "預設 Tourist 30%／Mix 40%／Non Tourist 50%，即 9–15 日銷售量。\n"
  "用途：限制設定可以高到幾多。", size=8.6, color=SUB)

box(0.510, _y3, 0.440, _h3, GREEN_BG)
T(0.526, _y3 + 0.078, "底線 Y（OOS Floor）", size=10.8, color=GREEN, weight="bold")
T(0.526, _y3 + 0.052, "Y = ceil（月銷預估 × OOS Rate）", size=10.4, color=INK, weight="bold")
T(0.526, _y3 + 0.024,
  "預設一律 50%，即約半個月銷售。\n"
  "用途：保證設定不會低到缺貨。", size=8.6, color=SUB)

arrow(0.5, 0.502, 0.5, 0.494)

# ── ④ 第一步：壓上限 ──
_y4, _h4 = 0.412, 0.080
step_band(
    _y4, _h4, "4", "第一步：把現時設定壓到不超過上限",
    "X 欄 = min（X0, Z）",
    ["Z 未超上限 → X 欄 = Z：保留你現有設定",
     "Z 已超上限 → X 欄 = X0：只保留到上限",
     "RAW「Overstock Cap (X)」欄＝這個 X 欄"],
    ORANGE, "#FDF6EE")

arrow(0.5, 0.410, 0.5, 0.398)

# ── ⑤ 第二步：取較大 ──
_y5, _h5 = 0.316, 0.080
step_band(
    _y5, _h5, "5", "第二步：再同底線比較，哪個大用哪個",
    "建議值 = max（Y, X 欄）",
    ["Y 較大 → 用 Y：補到缺貨底線",
     "X 欄較大 → 用 X 欄：受上限壓住",
     "Tourist 上限 30% < 底線 50% → 必用 Y"],
    GREEN, "#F4FAF6")

arrow(0.5, 0.314, 0.5, 0.302)

# ── ⑥ 調整量 ──
_y6, _h6 = 0.246, 0.042
box(0.05, _y6, 0.90, _h6, HEAD_BG)
box(0.05, _y6, 0.006, _h6, TEAL, r=0.003)
chip(0.068, _y6 + _h6 / 2, "6", TEAL)
T(0.098, _y6 + _h6 / 2, "調整量 = 建議值 − 現時設定 Z", size=10.6, color=INK, weight="bold")

_chips = [("增加", "> 0", TEAL, TEAL_BG), ("不變", "= 0", GREY, "#EDF0F3"), ("減少", "< 0", ORANGE, ORANGE_BG)]
_cx = 0.522
for _name, _rule, _fg, _bg in _chips:
    box(_cx, _y6 + 0.006, 0.132, 0.030, _bg, ec=_fg, lw=0.8, r=0.005)
    T(_cx + 0.012, _y6 + 0.021, _name, size=9.6, color=_fg, weight="bold")
    T(_cx + 0.120, _y6 + 0.021, _rule, size=9.0, color=_fg, ha="right")
    _cx += 0.142

# ── 例子 ──
_ye = 0.230
T(0.05, _ye, "例子：同一間 Mix 店、同一個 SKU 的五種現時設定", size=11.4, weight="bold")
T(0.05, _ye - 0.016,
  "假設 P90D 月均 80 件、月銷預估 80 件 → X0 = floor(80 × 40%) = 32、Y = ceil(80 × 50%) = 40、X 欄 = min(32, Z)",
  size=8.8, color=SUB)

_rows = [
    ("10 件", "10", "40 件", "+30", "增加", TEAL, TEAL_BG),
    ("32 件", "32", "40 件", "+8", "增加", TEAL, TEAL_BG),
    ("33 件", "32", "40 件", "+7", "增加", TEAL, TEAL_BG),
    ("40 件", "32", "40 件", "0", "不變", GREY, "#EDF0F3"),
    ("60 件", "32", "40 件", "-20", "減少", ORANGE, ORANGE_BG),
]
_tx, _tw = 0.05, 0.520
_cols = [0.050, 0.155, 0.305, 0.418, 0.498]
_hdr = 0.199
box(_tx, _hdr - 0.015, _tw, 0.021, HEAD_BG, r=0.0)
for _i, _h in enumerate(["現時設定 Z", "第一步 X 欄", "建議值", "調整量", "分類"]):
    T(_cols[_i] + 0.009, _hdr - 0.0045, _h, size=9.2, color=INK, weight="bold")
_rh = 0.022
for _r, (_z, _over, _rec, _delta, _cat, _fg, _bg) in enumerate(_rows):
    _ry = _hdr - 0.015 - _rh * (_r + 1)
    if _r % 2 == 1:
        box(_tx, _ry, _tw, _rh, ZEBRA, r=0.0)
    for _i, _v in enumerate([_z, _over, _rec, _delta]):
        T(_cols[_i] + 0.009, _ry + _rh / 2, _v, size=9.2, color=INK)
    box(_cols[4] + 0.006, _ry + 0.003, 0.052, 0.016, _bg, r=0.003)
    T(_cols[4] + 0.032, _ry + _rh / 2, _cat, size=8.8, color=_fg, weight="bold", ha="center")

# ── 右側說明卡 ──
_rx, _rw = 0.596, 0.354
box(_rx, 0.132, _rw, 0.092, "#F7F9FB", ec=LINE, lw=0.9)
T(_rx + 0.012, 0.208, "保底與特殊情況（優先於上方公式）", size=9.3, color=TEAL, weight="bold")
T(_rx + 0.012, 0.194,
  "• 最少 3 件：有銷量或 90 日內上市新品；零銷量舊品不保底\n"
  "• P90D 0 Sales 舊品：可加可減 → 0，並建議轉 ND\n"
  "• RP Type＝ND：可加可減 → 0；只可增加 → 維持\n"
  "• 新品（90 日內）：照公式，不套零銷量歸零",
  size=8.2, color=INK, ls=1.62, va="top")

box(_rx, 0.026, _rw, 0.092, "#F7F9FB", ec=LINE, lw=0.9)
T(_rx + 0.012, 0.102, "另一個基準：周轉天數（按 Nature）", size=9.3, color=TEAL, weight="bold")
T(_rx + 0.012, 0.088,
  "• 建議值 = ceil（P90D ÷ 有效日數 × 基準天數）\n"
  "• 有效日數 = min(90, 上市至參考日（含首尾）)，舊品 90\n"
  "• 預設 Tourist 10／Mix 9／Non Tourist 7（可調 7–10）\n"
  "• 兩個基準恆並列計算，切換只決定建議值用哪條",
  size=8.2, color=INK, ls=1.62, va="top")

# ── 頁尾 ──
T(0.05, 0.016,
  "Z＝Safety Stock (現時)（來源檔 Adj. Safety Stock）｜X0＝理論上限｜X 欄＝min(X0, Z)，即 RAW「Overstock Cap (X)」欄｜Y＝OOS Floor",
  size=7.9, color=SUB)
T(0.05, 0.006,
  "「只可增加」模式調整量下限 0，不會出現「減少」；周轉日數對照欄 Recommend Safety Stock Turnday = 建議值 ÷（P90D 月均 ÷ 30）。",
  size=7.9, color=SUB)

out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "assets", "safety_logic_guide.png")
fig.savefig(out, dpi=140, facecolor="white")
print(out)
