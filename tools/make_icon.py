"""Draw the app icon -> assets/simple_sheets.ico (16…256 px) and build/icon-1024.png (the macOS build). It's the Simple
apps' family (simplecore.brand_icon): every app the same tile - a soft gradient of its accent, a shine, a white
symbol of one weight - only the colour and the symbol ("sheets") are its own. Build-time only (needs Pillow)."""
import os

from simplecore import brand_icon

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "assets", "simple_sheets.ico")
BIG = os.path.join(ROOT, "build", "icon-1024.png")
PNG = None


def main():
    brand_icon.save("sheets", OUT, BIG, PNG)
    print("wrote", OUT, "and", BIG)


if __name__ == "__main__":
    main()
