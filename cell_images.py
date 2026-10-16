"""=IMAGE(source, [alt_text], [sizing], [height], [width]): a picture from the web drawn inside its cell.

The formula's value is an ImageValue (text = the alt text, so other formulas and CSV see that). The picture
downloads on a thread (http / https / data: URLs, up to 10 MB) and the grid redraws when it arrives. Tk shows
PNG / GIF itself; JPEG / BMP / TIFF / ICO… are turned into PNG with Windows' GDI+ (or Pillow when installed).
sizing: 0 fit in the cell keeping its shape (default) · 1 fill the cell · 2 original size · 3 height x width px."""
import base64
import os
import queue
import sys
import threading
import urllib.error
import urllib.parse
import urllib.request

MAX_BYTES = 10 * 1024 * 1024


class ImageValue(str):
    """The value of =IMAGE(...): behaves as its alt text; carries the picture's address and sizing."""

    def __new__(cls, url, alt="", sizing=0, height=None, width=None):
        v = super().__new__(cls, alt)
        v.url, v.sizing, v.height, v.width = url, sizing, height, width
        return v


# ---------- downloading ----------

def data_url(url):
    """The bytes of a data: URL (data:image/png;base64,...)."""
    head, _, body = url.partition(",")
    return base64.b64decode(body) if ";base64" in head else urllib.parse.unquote_to_bytes(body)


def fetch(url, timeout=15):
    """The picture's bytes (tests swap this)."""
    if url.startswith("data:"):
        return data_url(url)
    if not url.lower().startswith(("http://", "https://")):
        raise ValueError("Only web (http / https) pictures can be shown.")
    req = urllib.request.Request(url, headers={            # some sites refuse requests that don't look like a browser
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) SimpleSheets",
        "Accept": "image/png,image/jpeg,image/gif,image/webp,image/*;q=0.8"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = r.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise ValueError("The picture is bigger than 10 MB.")
    return data


def to_png_or_gif(data):
    """Bytes Tk can show (PNG / GIF), converting other formats; None when it can't (why_not says why)."""
    if data[:8] == b"\x89PNG\r\n\x1a\n" or data[:6] in (b"GIF87a", b"GIF89a"):
        return data
    if looks_like_text(data):                       # a web page / SVG: no decoder will read it
        return None
    if sys.platform == "win32":
        for convert in (wic_png, gdiplus_png):       # WIC reads WebP / HEIC / JPEG…; GDI+ for older Windows
            try:
                return convert(data)
            except Exception:
                pass
    try:
        import io

        from PIL import Image as PILImage  # when Pillow happens to be installed
        out = io.BytesIO()
        PILImage.open(io.BytesIO(data)).save(out, "PNG")
        return out.getvalue()
    except Exception:
        return None


def looks_like_text(data):
    head = data[:512].lstrip().lower()
    return head.startswith((b"<!doctype", b"<html", b"<?xml", b"<svg", b"{")) or b"<svg" in head


def why_not(data):
    """Why a download isn't a picture that can be shown, in plain words."""
    head = data[:512].lstrip().lower()
    if b"<svg" in head:
        return "SVG drawings can't be shown - use a PNG or JPEG link."
    if head.startswith((b"<!doctype", b"<html")) or b"<html" in head:
        return ("the link opens a web page, not a picture. In the browser, right-click the picture → "
                "Copy image address, and use that link.")
    return "this kind of picture can't be read (PNG, JPEG, GIF, WebP, BMP work)."


def png_bytes(w, h, rgba):
    """A PNG of w x h RGBA pixels (stdlib zlib)."""
    import struct
    import zlib
    stride = w * 4
    raw = b"".join(b"\x00" + bytes(rgba[y * stride:(y + 1) * stride]) for y in range(h))

    def chunk(kind, body):
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 3)) + chunk(b"IEND", b""))


def wic_png(data):
    """Decode any picture Windows Imaging Component reads (JPEG, WebP, HEIC, TIFF, BMP, ICO…) -> PNG bytes."""
    import ctypes
    from ctypes import POINTER, byref, c_uint, c_void_p

    class GUID(ctypes.Structure):
        _fields_ = [("a", ctypes.c_uint32), ("b", ctypes.c_uint16), ("c", ctypes.c_uint16), ("d", ctypes.c_ubyte * 8)]

    def guid(s):
        h = s.replace("-", "")
        return GUID(int(h[:8], 16), int(h[8:12], 16), int(h[12:16], 16),
                    (ctypes.c_ubyte * 8)(*bytes.fromhex(h[16:])))

    def method(obj, index, *argtypes):
        """COM method `index` of obj's vtable, as a callable taking obj first."""
        vtable = ctypes.cast(ctypes.cast(obj, POINTER(c_void_p))[0], POINTER(c_void_p))
        return ctypes.WINFUNCTYPE(ctypes.HRESULT, c_void_p, *argtypes)(vtable[index])

    def release(obj):
        if obj:
            ctypes.WINFUNCTYPE(c_uint, c_void_p)(
                ctypes.cast(ctypes.cast(obj, POINTER(c_void_p))[0], POINTER(c_void_p))[2])(obj)

    ole32 = ctypes.windll.ole32
    started = ole32.CoInitializeEx(None, 0) in (0, 1)       # S_OK / S_FALSE (RPC_E_CHANGED_MODE: already set)
    factory, stream, decoder, frame, conv = (c_void_p() for _ in range(5))
    buf = ctypes.create_string_buffer(bytes(data), len(data))
    try:
        ole32.CoCreateInstance.argtypes = [c_void_p, c_void_p, c_uint, c_void_p, c_void_p]
        hr = ole32.CoCreateInstance(byref(guid("cacaf262-9370-4615-a13b-9f5539da4c0a")), None, 1,
                                    byref(guid("ec5ec8a9-c395-4314-9c77-54d7a935ff70")), byref(factory))
        if hr != 0:
            raise OSError("WIC isn't available")
        method(factory, 14, POINTER(c_void_p))(factory, byref(stream))                     # CreateStream
        method(stream, 16, c_void_p, c_uint)(stream, buf, len(data))                       # InitializeFromMemory
        method(factory, 4, c_void_p, c_void_p, c_uint, POINTER(c_void_p))(                  # CreateDecoderFromStream
            factory, stream, None, 0, byref(decoder))
        method(decoder, 13, c_uint, POINTER(c_void_p))(decoder, 0, byref(frame))           # GetFrame(0)
        method(factory, 10, POINTER(c_void_p))(factory, byref(conv))                       # CreateFormatConverter
        bgra = guid("6fddc324-4e03-4bfe-b185-3d77768dc90f")                                # 32bppBGRA
        method(conv, 8, c_void_p, c_void_p, c_uint, c_void_p, ctypes.c_double, c_uint)(   # Initialize
            conv, frame, byref(bgra), 0, None, 0.0, 0)
        w, h = c_uint(), c_uint()
        method(conv, 3, POINTER(c_uint), POINTER(c_uint))(conv, byref(w), byref(h))        # GetSize
        if not (0 < w.value <= 8000 and 0 < h.value <= 8000):
            raise OSError("picture too big")
        stride = w.value * 4
        pixels = ctypes.create_string_buffer(stride * h.value)
        method(conv, 7, c_void_p, c_uint, c_uint, c_void_p)(conv, None, stride, len(pixels), pixels)  # CopyPixels
        rgba = bytearray(pixels.raw)
        rgba[0::4], rgba[2::4] = rgba[2::4], rgba[0::4]                                    # BGRA -> RGBA
        return png_bytes(w.value, h.value, rgba)
    finally:
        for obj in (conv, frame, decoder, stream, factory):
            release(obj)
        if started:
            ole32.CoUninitialize()


def gdiplus_png(data):
    """Any picture Windows can read (JPEG, BMP, TIFF, ICO…) -> PNG bytes, with GDI+ (no Pillow needed)."""
    import ctypes
    import tempfile
    from ctypes import wintypes as wt
    gdip, shlwapi = ctypes.windll.gdiplus, ctypes.windll.shlwapi
    shlwapi.SHCreateMemStream.argtypes = [ctypes.c_char_p, ctypes.c_uint]
    gdip.GdipCreateBitmapFromStream.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)]
    gdip.GdipSaveImageToFile.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_void_p, ctypes.c_void_p]
    gdip.GdipDisposeImage.argtypes = [ctypes.c_void_p]

    class StartupInput(ctypes.Structure):
        _fields_ = [("version", ctypes.c_uint32), ("callback", ctypes.c_void_p),
                    ("no_thread", wt.BOOL), ("no_codecs", wt.BOOL)]

    class GUID(ctypes.Structure):
        _fields_ = [("a", ctypes.c_uint32), ("b", ctypes.c_uint16), ("c", ctypes.c_uint16), ("d", ctypes.c_ubyte * 8)]

    token = ctypes.c_size_t()
    si = StartupInput(1, None, False, False)
    if gdip.GdiplusStartup(ctypes.byref(token), ctypes.byref(si), None) != 0:
        raise OSError("GDI+ didn't start")
    bitmap = ctypes.c_void_p()
    shlwapi.SHCreateMemStream.restype = ctypes.c_void_p
    stream = shlwapi.SHCreateMemStream(data, len(data))
    fd, path = tempfile.mkstemp(suffix=".png", prefix="ss-img-")
    os.close(fd)
    try:
        if not stream or gdip.GdipCreateBitmapFromStream(stream, ctypes.byref(bitmap)) != 0:
            raise OSError("not a picture Windows can read")
        png = GUID(0x557CF406, 0x1A04, 0x11D3, (ctypes.c_ubyte * 8)(0x9A, 0x73, 0x00, 0x00, 0xF8, 0x1E, 0xF3, 0x2E))
        if gdip.GdipSaveImageToFile(bitmap, path, ctypes.byref(png), None) != 0:
            raise OSError("couldn't convert the picture")
        gdip.GdipDisposeImage(bitmap)
        bitmap = ctypes.c_void_p()
        with open(path, "rb") as f:
            return f.read()
    finally:
        if bitmap:
            gdip.GdipDisposeImage(bitmap)
        if stream:
            # IStream::Release (vtable slot 2)
            release = ctypes.cast(ctypes.cast(stream, ctypes.POINTER(ctypes.c_void_p))[0],
                                  ctypes.POINTER(ctypes.c_void_p))[2]
            ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(release)(stream)
        gdip.GdiplusShutdown(token)
        try:
            os.remove(path)
        except OSError:
            pass


# ---------- cache: url -> ("loading" | "ok" | "error", png / gif bytes or message) ----------

_cache = {}
_done = queue.Queue()


def state(url):
    return _cache.get(url, (None, None))


def request(grid, url):
    """Start downloading a picture once; the grid redraws when it's there."""
    if url in _cache:
        return _cache[url]
    _cache[url] = ("loading", None)

    def work():
        try:
            raw = fetch(url)
            data = to_png_or_gif(raw)
            _done.put((url, ("ok", data) if data else ("error", why_not(raw))))
        except urllib.error.HTTPError as e:
            _done.put((url, ("error", {403: "the website refused (403) - it may block pictures used elsewhere.",
                                       404: "nothing at that link (404)."}.get(e.code, f"the website said {e.code}."))))
        except Exception as e:
            reason = getattr(e, "reason", None) or str(e) or type(e).__name__
            _done.put((url, ("error", f"couldn't download it ({reason}).")))
    threading.Thread(target=work, daemon=True).start()
    _watch(grid)
    return _cache[url]


def _watch(grid):
    if getattr(grid, "_image_watch", False):
        return
    grid._image_watch = True

    def check():
        got = False
        while True:
            try:
                url, result = _done.get_nowait()
            except queue.Empty:
                break
            _cache[url] = result
            got = True
        if got and grid.winfo_exists():
            grid.redraw()
        if any(s == "loading" for s, _ in _cache.values()) and grid.winfo_exists():
            grid._root().after(150, check)
        else:
            grid._image_watch = False
    grid._root().after(150, check)


def clear():
    _cache.clear()


# ---------- drawing ----------

def box(v, w, h, pw, ph, px):
    """(width, height) the picture is drawn at in a w x h cell, for the formula's sizing."""
    if v.sizing == 1:                                  # fill the cell
        return w, h
    if v.sizing == 2:                                  # original size
        return pw, ph
    if v.sizing == 3 and v.height and v.width:         # custom (pixels)
        return px(v.width), px(v.height)
    k = min(w / max(pw, 1), h / max(ph, 1))            # fit, keeping its shape
    return max(int(pw * k), 1), max(int(ph * k), 1)


def draw(grid, v, x, y, w, h):
    """Draw the picture of an ImageValue in its cell -> True, or False to show the alt text instead."""
    import tkinter as tk

    import images
    st, data = state(v.url)
    if st is None:
        st, data = request(grid, v.url)
    if st != "ok":
        return False
    cache = grid.__dict__.setdefault("cell_image_cache", {})
    key = (v.url, w, h, v.sizing, v.height, v.width, round(grid.zoom, 2))
    photo = cache.get(key)
    if photo is None:
        try:
            src = tk.PhotoImage(master=grid, data=base64.b64encode(data))
        except tk.TclError:
            return False
        pad = 2
        bw, bh = box(v, max(w - 2 * pad, 1), max(h - 2 * pad, 1), src.width(), src.height(), grid.px)
        photo = images.scaled(src, bw, bh) if (bw, bh) != (src.width(), src.height()) else src
        if len(cache) > 200:
            cache.clear()
        cache[key] = photo
    grid.cv.create_image(x + w / 2, y + h / 2, image=photo, anchor="center", tags="cellimage")
    return True
