"""Draw the app icon (green sheet with a grid) -> assets/simple_sheets.ico. Build-time only (needs Pillow)."""
import os

from PIL import Image, ImageDraw

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "assets", "simple_sheets.ico")


def draw(size):
    s = size / 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([16 * s, 8 * s, 240 * s, 248 * s], radius=28 * s, fill=(24, 128, 56, 255))   # green page
    x0, y0, x1, y1 = 48 * s, 56 * s, 208 * s, 216 * s
    d.rectangle([x0, y0, x1, y1], fill=(255, 255, 255, 255))                                          # white sheet
    d.rectangle([x0, y0, x1, y0 + 40 * s], fill=(200, 230, 201, 255))                                  # header row
    line = max(int(6 * s), 1)
    for i in range(1, 4):                                                                              # grid
        y = y0 + i * 40 * s
        d.line([x0, y, x1, y], fill=(24, 128, 56, 255), width=line)
    for x in (x0 + 64 * s, x0 + 112 * s):
        d.line([x, y0, x, y1], fill=(24, 128, 56, 255), width=line)
    return img


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    sizes = [16, 24, 32, 48, 64, 128, 256]
    big = draw(256)
    big.save(OUT, sizes=[(n, n) for n in sizes], append_images=[draw(n) for n in sizes[:-1]])
    print("wrote", OUT)


if __name__ == "__main__":
    main()
