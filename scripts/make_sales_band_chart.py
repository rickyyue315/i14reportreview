# 銷量帶分類說明圖（供電郵講解用；非應用程式元件）
# 分帶準則與 transformer._hybrid_sales_bands 一致：P90D≤0＝零銷；
# 有銷列平均排名百分比 >75%＝高銷、≤25%＝低銷、其餘＝中銷；
# 有銷店舖 <4 間不產生高銷／低銷；並列以平均排名處理（同值同帶）。
import os
import textwrap

import matplotlib

matplotlib.use("Agg")
from matplotlib import font_manager
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle

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
GREY_BG = "#EAEDF1"
HEAD_BG = "#E4EEF5"
ZEBRA = "#F6F9FB"

BAND_FG = {"高銷": GREEN, "中銷": TEAL, "低銷": ORANGE, "零銷": GREY}
BAND_BG = {"高銷": GREEN_BG, "中銷": TEAL_BG, "低銷": ORANGE_BG, "零銷": GREY_BG}

fig = plt.figure(figsize=(12.5, 11.0), dpi=130)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, 1)
ax.set_ylim(0, 1)
ax.axis("off")
fig.patch.set_facecolor("white")


def T(x, y, s, size=11, color=INK, weight="normal", ha="left", va="center"):
    ax.text(x, y, s, fontsize=size, color=color, fontweight=weight,
            ha=ha, va=va, transform=ax.transAxes)


def box(x, y, w, h, fc, ec="none", lw=0.8, round_=True):
    if round_:
        ax.add_patch(FancyBboxPatch((x, y), w, h,
                                    boxstyle="round,pad=0,rounding_size=0.008",
                                    facecolor=fc, edgecolor=ec, linewidth=lw,
                                    transform=ax.transAxes))
    else:
        ax.add_patch(Rectangle((x, y), w, h, facecolor=fc, edgecolor=ec,
                               linewidth=lw, transform=ax.transAxes))


# ── 標題 ──
T(0.06, 0.958, "銷量帶分類說明", size=24, weight="bold")
T(0.06, 0.922, "Safety Stock x NetStock Review｜以同 SKU 各店舖的 Sold Qty (P90D) 排名分帶（Sold Qty (P90D)＝過去 90 日店舖銷量）",
  size=11.5, color=SUB)

# ── 銷量帶示意條 ──
BAR_Y, BAR_H = 0.845, 0.055
box(0.06, BAR_Y, 0.20, BAR_H, GREY_BG, round_=False)
T(0.16, BAR_Y + BAR_H * 0.60, "零銷", size=13, color=GREY, weight="bold", ha="center")
T(0.16, BAR_Y + BAR_H * 0.22, "Sold Qty (P90D) ≤ 0（不參與排名）", size=8.5, color=GREY, ha="center")

segs = [("低銷", "排名百分比 ≤ 25%", 0.25), ("中銷", "> 25% 且 ≤ 75%", 0.50), ("高銷", "> 75%", 0.25)]
bx = 0.30
for name, rng, frac in segs:
    w = 0.64 * frac
    box(bx, BAR_Y, w, BAR_H, BAND_BG[name], round_=False)
    T(bx + w / 2, BAR_Y + BAR_H * 0.60, name, size=13, color=BAND_FG[name], weight="bold", ha="center")
    T(bx + w / 2, BAR_Y + BAR_H * 0.22, rng, size=8.5, color=BAND_FG[name], ha="center")
    bx += w
for frac in (0.0, 0.25, 0.75, 1.0):
    xt = 0.30 + 0.64 * frac
    ax.plot([xt, xt], [BAR_Y, BAR_Y - 0.012], color=SUB, lw=0.8, transform=ax.transAxes)
    T(xt, BAR_Y - 0.028, f"{frac:.0%}", size=9, color=SUB, ha="center")

# ── 表格繪製工具 ──
def table(x, y, w, col_ws, headers, rows, row_h=0.040, header_h=0.038,
          band_col=None, aligns=None, header_size=10.5, cell_size=10.5):
    """表格：col_ws 為各欄寬（總和＝1），rows 內容為字串；band_col 指定該欄以銷量帶配色。"""
    aligns = aligns or ["left"] * len(col_ws)
    xs = [x]
    for cw in col_ws[:-1]:
        xs.append(xs[-1] + cw * w)

    def cell_xy(i):
        return xs[i] + 0.008, xs[i] + col_ws[i] * w - 0.008

    box(x, y - header_h, w, header_h, HEAD_BG, round_=False)
    for i, htxt in enumerate(headers):
        lx, rx = cell_xy(i)
        ha = aligns[i]
        tx = {"left": lx, "center": (lx + rx) / 2, "right": rx}[ha]
        T(tx, y - header_h / 2, htxt, size=header_size, color=INK, weight="bold", ha=ha)
    for r, row in enumerate(rows):
        ry = y - header_h - (r + 1) * row_h
        if r % 2 == 1:
            box(x, ry, w, row_h, ZEBRA, round_=False)
        ax.plot([x, x + w], [ry, ry], color=LINE, lw=0.6, transform=ax.transAxes)
        for i, val in enumerate(row):
            lx, rx = cell_xy(i)
            ha = aligns[i]
            tx = {"left": lx, "center": (lx + rx) / 2, "right": rx}[ha]
            if band_col is not None and i == band_col and val in BAND_FG:
                pad_x, pad_y = 0.006, row_h * 0.16
                box(lx, ry + pad_y, rx - lx - pad_x, row_h - 2 * pad_y,
                    BAND_BG[val], round_=False)
                T((lx + rx) / 2, ry + row_h / 2, val, size=cell_size,
                  color=BAND_FG[val], weight="bold", ha="center")
            else:
                color = BAND_FG.get(val, INK) if band_col is not None and i == band_col else INK
                T(tx, ry + row_h / 2, str(val), size=cell_size, color=color,
                  weight="bold" if (band_col is not None and i == band_col) else "normal", ha=ha)
    ax.plot([x, x + w], [y - header_h - len(rows) * row_h, y - header_h - len(rows) * row_h],
            color=LINE, lw=0.8, transform=ax.transAxes)
    return y - header_h - len(rows) * row_h


# ── 一、判定條件 ──
T(0.06, 0.775, "一、判定條件（同一 SKU 內比較）", size=14, weight="bold")
rows1 = [
    ("高銷", "有銷售店舖中，排名百分比 > 75%", "同 SKU 內銷量最高的店舖"),
    ("中銷", "其餘有銷售店舖（排名百分比 > 25% 且 ≤ 75%）", "中間地帶，不作高低區分"),
    ("低銷", "有銷售店舖中，排名百分比 ≤ 25%", "同 SKU 內銷量最少的店舖"),
    ("零銷", "Sold Qty (P90D) ≤ 0（不參與排名）", "由「Past 90 Days 0 Sales 轉 ND」提示另行跟進"),
]
table(0.06, 0.758, 0.88, [0.10, 0.44, 0.46],
      ["銷量帶", "判定條件", "說明"], rows1,
      row_h=0.042, band_col=0, aligns=["center", "left", "left"])

# ── 二、例子 ──
y2 = 0.524
T(0.06, y2, "二、例子：同一 SKU 在 10 間店舖（8 間有銷售、2 間零銷售）", size=14, weight="bold")
rank_pct = ["100%", "87.5%", "75%", "62.5%", "50%", "37.5%", "25%", "12.5%", "—", "—"]
qty = [90, 80, 70, 60, 50, 40, 30, 20, 0, 0]
bands = ["高銷", "高銷", "中銷", "中銷", "中銷", "中銷", "低銷", "低銷", "零銷", "零銷"]
rows2 = []
for i in range(10):
    rows2.append((f"S{i + 1:02d}", f"{qty[i]} 件", rank_pct[i], bands[i]))
table(0.06, y2 - 0.028, 0.50, [0.16, 0.28, 0.30, 0.26],
      ["店舖", "Sold Qty (P90D)", "排名百分比（÷8）", "銷量帶"], rows2,
      row_h=0.033, header_h=0.048, band_col=3,
      aligns=["center", "right", "right", "center"])

# ── 備註卡（與例子表並排）──
NX, NW = 0.62, 0.32
notes = [
    "排名百分比＝平均排名 ÷ 有銷售店舖數；並列（同銷量）以平均排名處理，同值同帶。",
    "邊界準則：> 75% 屬高銷（恰好 75% 屬中銷）；≤ 25% 屬低銷。",
    "同 SKU 有銷售店舖少於 4 間時，不設高銷／低銷，全部歸中銷。",
    "按本次上傳資料內同 SKU 排名計算；篩選範圍不同，分帶結果可能不同。",
]
NY, NH = 0.480, 0.362
box(NX, NY - NH, NW, NH, "#F6F9FB", ec=LINE, lw=0.9)
T(NX + 0.018, NY - 0.030, "備註", size=13, weight="bold", color=TEAL)
ny = NY - 0.064
for note in notes:
    T(NX + 0.018, ny, "•", size=10.5, color=TEAL)
    wrapped = textwrap.fill(note, width=23, break_long_words=True,
                            break_on_hyphens=False)
    n_lines = wrapped.count("\n") + 1
    ax.text(NX + 0.034, ny, wrapped, fontsize=9.8, color=INK, ha="left", va="top",
            transform=ax.transAxes, linespacing=1.45)
    ny -= 0.024 * n_lines + 0.022

T(0.06, 0.028, "銷量帶為 RAW／調貨 Review 主表的常開顯示欄，與調貨 14 日線及 Safety 基準周轉天數互相獨立。",
  size=9.5, color=SUB)

out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "assets", "sales_band_guide.png")
fig.savefig(out, dpi=130, facecolor="white")
print(out)
