"""openpyxl needs Pillow to keep the pictures in a workbook; without it (as in the light exe) it silently
drops every picture on save. This tiny stand-in gives openpyxl what it asks Pillow for — format, size and
the raw bytes of PNG / GIF / JPEG / BMP / EMF / TIFF pictures — so pictures are kept, untouched."""
import struct
from io import BytesIO


def sniff(data):
    """(format, (width, height)) of picture bytes; raises OSError for anything unknown."""
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", data[16:24])
        return "PNG", (w, h)
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "GIF", struct.unpack("<HH", data[6:10])
    if data[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                i += 2
                continue
            size = struct.unpack(">H", data[i + 2:i + 4])[0]
            if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return "JPEG", (w, h)
            i += 2 + size
        raise OSError("JPEG without a size")
    if data[:2] == b"BM":
        w, h = struct.unpack("<ii", data[18:26])
        return "BMP", (w, abs(h))
    if data[:4] == b"\x01\x00\x00\x00" and data[40:44] == b" EMF":
        left, top, right, bottom = struct.unpack("<iiii", data[24:40])   # frame in 0.01 mm
        return "EMF", (max((right - left) * 96 // 2540, 1), max((bottom - top) * 96 // 2540, 1))
    if data[:4] in (b"II*\x00", b"MM\x00*"):
        return "TIFF", (100, 100)
    raise OSError("not a picture format Simple Sheets knows")


class Image:
    """What openpyxl uses of a PIL image."""

    def __init__(self, data):
        self.data = data
        self.fp = BytesIO(data)
        self.format, self.size = sniff(data)

    def save(self, fp, format=None):
        fp.write(self.data)               # never converted: the bytes go back as they came

    def close(self):
        pass


class PIL:
    """Stands in for the `PIL.Image` module."""
    Image = Image

    @staticmethod
    def open(src):
        if hasattr(src, "read"):
            src.seek(0)
            data = src.read()
        else:
            with open(src, "rb") as f:
                data = f.read()
        return Image(data)


def install(force=False):
    """Use the stand-in when Pillow isn't there (force=True: always, for tests). Returns True if installed."""
    import openpyxl.drawing.image as oimg
    import openpyxl.reader.drawings as odraw
    if oimg.PILImage and not force:
        return False
    oimg.PILImage = PIL
    odraw.PILImage = PIL
    return True


def uninstall():
    """Back to real Pillow (tests)."""
    import openpyxl.drawing.image as oimg
    import openpyxl.reader.drawings as odraw
    try:
        from PIL import Image as real
    except ImportError:
        real = False
    oimg.PILImage = real
    odraw.PILImage = real
