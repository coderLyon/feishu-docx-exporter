"""
生成三套扩展图标方案（含各尺寸预览），供确认后再落盘为正式图标。

用法：python tools/make-icons.py
产物：samples/out/icon-sheet.png（三方案 × 128/48/32/16）
      samples/out/icons/方案{A,B,C}_{size}.png
"""
import os
from PIL import Image, ImageDraw, ImageFont

OUT = os.path.join("samples", "out", "icons")
os.makedirs(OUT, exist_ok=True)

BLUE = (51, 112, 255, 255)        # 飞书蓝 #3370FF
BLUE_DARK = (36, 88, 214, 255)
WHITE = (255, 255, 255, 255)
INK = (31, 35, 41, 255)

S = 512  # 超采样基准


def font(size, cjk=False):
    """英文用 Arial Bold，中文用微软雅黑（预览图标题需要）"""
    candidates = (
        [r"C://Windows//Fonts//msyhbd.ttc", r"C://Windows//Fonts//simhei.ttf"]
        if cjk
        else [r"C://Windows//Fonts//arialbd.ttf", r"C://Windows//Fonts//segoeuib.ttf"]
    )
    for path in candidates:
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default()


def rounded(draw, box, radius, fill):
    draw.rounded_rectangle(box, radius=radius, fill=fill)


def doc_shape(draw, x, y, w, h, color, fold=0.32):
    """文档：矩形 + 右上折角"""
    r = int(w * 0.10)
    draw.rounded_rectangle([x, y, x + w, y + h], radius=r, fill=color)
    fx = x + w - int(w * fold)
    draw.polygon([(fx, y), (x + w, y), (x + w, y + int(w * fold))], fill=(0, 0, 0, 0))


def arrow_down(draw, cx, cy, size, color, width):
    """向下箭头 + 底座横线（导出/保存意象）"""
    head = size // 2
    draw.line([(cx, cy - head), (cx, cy + head // 2)], fill=color, width=width)
    draw.polygon(
        [(cx - head, cy + head // 4), (cx + head, cy + head // 4), (cx, cy + head)],
        fill=color,
    )
    draw.line(
        [(cx - head, cy + head + int(width * 1.4)), (cx + head, cy + head + int(width * 1.4))],
        fill=color,
        width=width,
    )


def icon_a():
    """A：蓝色圆角底 + 白色文档 + 导出箭头"""
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    rounded(d, [0, 0, S - 1, S - 1], int(S * 0.22), BLUE)
    # 文档（略偏左上）
    dx, dy, dw, dh = int(S * 0.16), int(S * 0.12), int(S * 0.34), int(S * 0.42)
    d.rectangle([dx, dy, dx + dw, dy + dh], fill=(255, 255, 255, 235))
    # 文本行
    for i in range(3):
        yy = dy + int(dh * (0.24 + i * 0.20))
        d.rectangle([dx + int(dw * 0.18), yy, dx + int(dw * 0.82), yy + int(S * 0.022)],
                    fill=(120, 150, 220, 255))
    # 箭头（右下）
    arrow_down(d, int(S * 0.70), int(S * 0.64), int(S * 0.34), WHITE, int(S * 0.075))
    return img


def icon_b():
    """B：白底蓝边 + 大 W（Word 意象）"""
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    rounded(d, [int(S * 0.03), int(S * 0.03), S - int(S * 0.03), S - int(S * 0.03)],
            int(S * 0.20), WHITE)
    d.rounded_rectangle(
        [int(S * 0.03), int(S * 0.03), S - int(S * 0.03), S - int(S * 0.03)],
        radius=int(S * 0.20), outline=BLUE, width=int(S * 0.055),
    )
    f = font(int(S * 0.62))
    text = "W"
    bbox = d.textbbox((0, 0), text, font=f)
    d.text(((S - (bbox[2] - bbox[0])) / 2 - bbox[0], (S - (bbox[3] - bbox[1])) / 2 - bbox[1] - int(S * 0.02)),
           text, font=f, fill=BLUE)
    return img


def icon_c():
    """C：蓝色圆角底 + 文档折角 + 右下角圆形箭头徽标"""
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    rounded(d, [0, 0, S - 1, S - 1], int(S * 0.22), BLUE_DARK)
    # 白页
    px, py, pw, ph = int(S * 0.20), int(S * 0.14), int(S * 0.44), int(S * 0.52)
    d.rounded_rectangle([px, py, px + pw, py + ph], radius=int(S * 0.05), fill=WHITE)
    # 折角
    fold = int(pw * 0.34)
    d.polygon([(px + pw - fold, py), (px + pw, py), (px + pw, py + fold)], fill=(198, 214, 250, 255))
    # 文本行
    for i in range(3):
        yy = py + int(ph * (0.30 + i * 0.17))
        d.rectangle([px + int(pw * 0.16), yy, px + int(pw * 0.84), yy + int(S * 0.020)],
                    fill=(150, 172, 220, 255))
    # 右下圆形徽标 + 箭头
    r = int(S * 0.26)
    cx, cy = int(S * 0.70), int(S * 0.70)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=WHITE)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=BLUE_DARK, width=int(S * 0.022))
    arrow_down(d, cx, cy - int(S * 0.01), int(S * 0.22), BLUE_DARK, int(S * 0.045))
    return img


SIZES = [128, 48, 32, 16]
labels = {"A": "方案 A：蓝色底 + 文档 + 导出箭头", "B": "方案 B：白底蓝框 + W 字标", "C": "方案 C：深蓝底 + 文档 + 圆形导出徽标"}
makers = {"A": icon_a, "B": icon_b, "C": icon_c}

for key, fn in makers.items():
    base = fn()
    for s in SIZES:
        base.resize((s, s), Image.LANCZOS).save(os.path.join(OUT, f"{key}_{s}.png"))

# 预览：每行 128 原始尺寸 + 48/32/16 三倍放大 + 16px 原始尺寸
sheet = Image.new("RGBA", (1000, 700), (246, 247, 249, 255))
d = ImageDraw.Draw(sheet)
f_title = font(20, cjk=True)
f_cap = font(14, cjk=True)

y = 28
for key in ("A", "B", "C"):
    d.text((32, y), labels[key], font=f_title, fill=INK)
    base = makers[key]()
    sheet.alpha_composite(base.resize((128, 128), Image.LANCZOS), (32, y + 34))
    d.text((32, y + 168), "128px", font=f_cap, fill=(120, 126, 134, 255))

    x = 210
    for sz, zoom in ((48, 3), (32, 3), (16, 4)):
        small = base.resize((sz, sz), Image.LANCZOS)
        big = small.resize((sz * zoom, sz * zoom), Image.NEAREST)
        sheet.alpha_composite(big, (x, y + 34))
        d.text((x, y + 34 + sz * zoom + 6), f"{sz}px 放大{zoom}倍", font=f_cap, fill=(120, 126, 134, 255))
        x += sz * zoom + 40
    # 实际大小
    sheet.alpha_composite(base.resize((16, 16), Image.LANCZOS), (x + 20, y + 34))
    d.text((x + 10, y + 34 + 26), "16px 实际", font=f_cap, fill=(120, 126, 134, 255))
    y += 215

sheet.convert("RGB").save(os.path.join("samples", "out", "icon-sheet.png"))
print("已生成 samples/out/icon-sheet.png 与各尺寸 PNG")
