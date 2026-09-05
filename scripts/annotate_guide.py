"""圖像教學標註腳本：為使用說明擷圖加上紅框與繁體中文標籤。"""
from PIL import Image, ImageDraw, ImageFont

FONT_PATH = "C:/Windows/Fonts/msjh.ttc"
FONT_SIZE = 26


def annotate(src: str, box: tuple, dst: str, label: str) -> None:
    im = Image.open(src).convert("RGB")
    draw = ImageDraw.Draw(im)
    font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
    draw.rectangle(box, outline=(227, 30, 30), width=6)
    x0, y0, _x1, _y1 = box
    bbox = draw.textbbox((0, 0), label, font=font)
    pad_x, pad_y = 14, 8
    label_box = (x0, max(y0 - (bbox[3] - bbox[1]) - pad_y * 2, 0), x0 + (bbox[2] - bbox[0]) + pad_x * 2, y0)
    draw.rectangle(label_box, fill=(227, 30, 30))
    draw.text((label_box[0] + pad_x, label_box[1] + pad_y), label, fill=(255, 255, 255), font=font)
    im.save(dst)
    print("wrote", dst, im.size)


jobs = [
    ("assets/guide/01_mode_select.png", (300, 310, 1160, 475), "assets/guide/01_mode_select_annotated.png", "1. 先選分析模式"),
    ("assets/guide/02_upload.png", (300, 265, 1160, 445), "assets/guide/02_upload_annotated.png", "2. 拖曳 Excel 至此上傳"),
]
for job in jobs:
    annotate(*job)
