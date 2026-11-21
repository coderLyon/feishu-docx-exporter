"""从同一份画稿生成 Chrome 工具栏和扩展页的 PNG 图标。"""
from pathlib import Path

from PIL import Image, ImageDraw

OUT = Path(__file__).resolve().parents[1] / "src" / "icons"
SCALE = 4
SIZE = 128


def xy(points):
    return tuple(round(value * SCALE) for value in points)


def make_icon():
    size = SIZE * SCALE
    mask = Image.new("L", (size, size))
    ImageDraw.Draw(mask).rounded_rectangle(xy((3, 3, 125, 125)), radius=26 * SCALE, fill=255)

    image = Image.new("RGBA", (size, size))
    draw = ImageDraw.Draw(image)
    for y in range(size):
        fraction = y / (size - 1)
        color = tuple(round(a + (b - a) * fraction) for a, b in zip((30, 83, 174), (19, 40, 90)))
        draw.line((0, y, size - 1, y), fill=(*color, 255))
    image.putalpha(mask)

    page = [(34, 21), (73, 21), (97, 45), (97, 100), (92, 107), (34, 107), (29, 102), (29, 26)]
    draw.polygon([xy(point) for point in page], fill="#f8fbff")
    draw.polygon([xy(point) for point in [(73, 21), (97, 45), (73, 45)]], fill="#a9c9ff")

    ink = "#214f9b"
    arrow = [xy(point) for point in [(50, 68), (63, 81), (76, 68)]]
    draw.line(arrow, fill=ink, width=8 * SCALE, joint="curve")
    draw.rounded_rectangle(xy((59, 49, 67, 79)), radius=4 * SCALE, fill=ink)
    draw.rounded_rectangle(xy((46, 90, 80, 96)), radius=3 * SCALE, fill="#4ea4af")
    return image


if __name__ == "__main__":
    master = make_icon()
    OUT.mkdir(parents=True, exist_ok=True)
    for size in (16, 32, 48, 128):
        master.resize((size, size), Image.Resampling.LANCZOS).save(OUT / f"icon{size}.png", optimize=True)
