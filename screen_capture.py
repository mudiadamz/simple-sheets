"""Pictures for the AI prompt without extra libraries: screen grabs (Windows GDI through ctypes), images
copied to the clipboard (Win+Shift+S, Print Screen, a picture copied in a browser…), and a small PNG
writer (zlib). Snip: dim the screen, drag a box, get that part as PNG."""
import struct
import sys
import zlib

# ---------- PNG ----------

def png(w, h, rgb):
    """Raw RGB bytes (top row first) -> PNG file bytes."""
    stride = w * 3
    raw = b"".join(b"\0" + rgb[y * stride:(y + 1) * stride] for y in range(h))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 6)) + chunk(b"IEND", b""))


def bgra_to_rgb(buf):
    """32-bit BGRA pixels -> RGB bytes."""
    n = len(buf) // 4
    out = bytearray(n * 3)
    out[0::3], out[1::3], out[2::3] = buf[2::4], buf[1::4], buf[0::4]
    return bytes(out)


# ---------- the screen ----------

def virtual_screen():
    """(x, y, w, h) of all monitors together."""
    if sys.platform != "win32":
        raise OSError("screen capture works on Windows")
    import ctypes
    m = ctypes.windll.user32.GetSystemMetrics
    return m(76), m(77), m(78), m(79)          # SM_XVIRTUALSCREEN, Y, CX, CY


def grab(x, y, w, h):
    """Screen pixels of a rectangle (screen coordinates) -> PNG bytes."""
    if sys.platform != "win32":
        raise OSError("screen capture works on Windows")
    import ctypes
    from ctypes import wintypes
    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    w, h = max(int(w), 1), max(int(h), 1)
    screen = user32.GetDC(0)
    mem = gdi32.CreateCompatibleDC(screen)
    bmp = gdi32.CreateCompatibleBitmap(screen, w, h)
    old = gdi32.SelectObject(mem, bmp)
    try:
        gdi32.BitBlt(mem, 0, 0, w, h, screen, int(x), int(y), 0x00CC0020 | 0x40000000)   # SRCCOPY | CAPTUREBLT

        class BITMAPINFOHEADER(ctypes.Structure):
            _fields_ = [("biSize", wintypes.DWORD), ("biWidth", wintypes.LONG), ("biHeight", wintypes.LONG),
                        ("biPlanes", wintypes.WORD), ("biBitCount", wintypes.WORD),
                        ("biCompression", wintypes.DWORD), ("biSizeImage", wintypes.DWORD),
                        ("biXPelsPerMeter", wintypes.LONG), ("biYPelsPerMeter", wintypes.LONG),
                        ("biClrUsed", wintypes.DWORD), ("biClrImportant", wintypes.DWORD)]
        info = BITMAPINFOHEADER(ctypes.sizeof(BITMAPINFOHEADER), w, -h, 1, 32, 0, 0, 0, 0, 0, 0)   # top-down
        buf = ctypes.create_string_buffer(w * h * 4)
        gdi32.GetDIBits(mem, bmp, 0, h, buf, ctypes.byref(info), 0)
        return png(w, h, bgra_to_rgb(buf.raw))
    finally:
        gdi32.SelectObject(mem, old)
        gdi32.DeleteObject(bmp)
        gdi32.DeleteDC(mem)
        user32.ReleaseDC(0, screen)


def grab_window(win):
    """A Tk window as it is on screen -> PNG bytes."""
    win.update_idletasks()
    return grab(win.winfo_rootx(), win.winfo_rooty(), win.winfo_width(), win.winfo_height())


# ---------- the clipboard ----------

def parse_dib(data):
    """A CF_DIB clipboard bitmap (BITMAPINFOHEADER + pixels, 24 or 32 bit) -> PNG bytes, or None."""
    if len(data) < 40:
        return None
    size, w, h, _planes, bits, comp = struct.unpack_from("<IiiHHI", data, 0)
    if bits not in (24, 32) or comp not in (0, 3) or w <= 0 or h == 0:
        return None
    offset = size + (12 if comp == 3 and size == 40 else 0)   # BI_BITFIELDS: three colour masks follow
    top_down, h = h < 0, abs(h)
    stride = ((w * bits + 31) // 32) * 4
    step = bits // 8
    rows = []
    for y in range(h):
        src = y if top_down else h - 1 - y
        row = data[offset + src * stride: offset + src * stride + w * step]
        if len(row) < w * step:
            return None
        if step == 4:
            rows.append(bgra_to_rgb(row))
        else:
            out = bytearray(w * 3)
            out[0::3], out[1::3], out[2::3] = row[2::3], row[1::3], row[0::3]
            rows.append(bytes(out))
    return png(w, h, b"".join(rows))


def clipboard_image():
    """The picture on the clipboard as PNG bytes (None when there's no picture)."""
    if sys.platform != "win32":
        return None
    import ctypes
    user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
    user32.GetClipboardData.restype = ctypes.c_void_p
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalSize.argtypes = [ctypes.c_void_p]
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    CF_DIB = 8
    if not user32.IsClipboardFormatAvailable(CF_DIB) or not user32.OpenClipboard(0):
        return None
    try:
        handle = user32.GetClipboardData(CF_DIB)
        if not handle:
            return None
        ptr = kernel32.GlobalLock(handle)
        try:
            data = ctypes.string_at(ptr, kernel32.GlobalSize(handle))
        finally:
            kernel32.GlobalUnlock(handle)
        return parse_dib(data)
    finally:
        user32.CloseClipboard()


# ---------- snip: drag a box on the screen ----------

class Snip:
    """A dimmed see-through layer over all screens: drag a box (Esc cancels). done(png_bytes or None)."""

    def __init__(self, master, done, grab_fn=None):
        import tkinter as tk

        from theme import C
        self.done, self.grab_fn = done, grab_fn or grab
        self.start = None
        x, y, w, h = virtual_screen()
        self.origin = (x, y)
        t = self.win = tk.Toplevel(master)
        t.overrideredirect(True)
        t.geometry(f"{w}x{h}+{x}+{y}")
        t.attributes("-topmost", True)
        t.attributes("-alpha", 0.35)
        cv = self.cv = tk.Canvas(t, background="#000000", highlightthickness=0, cursor="crosshair")
        cv.pack(fill="both", expand=True)
        cv.create_text(w // 2, 40, text="Drag a box around what to send  ·  Esc cancels", fill="#ffffff",
                       font=("Segoe UI", 16, "bold"))
        self.box = cv.create_rectangle(0, 0, 0, 0, outline=C["accent"], width=3)
        cv.bind("<ButtonPress-1>", self.press)
        cv.bind("<B1-Motion>", self.motion)
        cv.bind("<ButtonRelease-1>", self.release)
        t.bind("<Escape>", lambda e: self.finish(None))
        t.focus_force()

    def press(self, e):
        self.start = (e.x, e.y)

    def motion(self, e):
        if self.start:
            self.cv.coords(self.box, *self.start, e.x, e.y)

    def release(self, e):
        if not self.start:
            return
        (x1, y1), (x2, y2) = self.start, (e.x, e.y)
        rect = (min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1))
        self.win.withdraw()                         # the layer must not be in the picture
        if rect[2] < 4 or rect[3] < 4:
            self.finish(None)
            return
        ox, oy = self.origin
        self.win._root().after(120, lambda: self.finish(self.grab_fn(rect[0] + ox, rect[1] + oy, rect[2], rect[3])))

    def finish(self, data):
        if self.win.winfo_exists():
            self.win.destroy()
        self.done(data)
