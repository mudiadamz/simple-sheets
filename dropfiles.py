"""Drag & drop files from Explorer onto the window (Windows: WM_DROPFILES through a subclassed window
procedure; Tk has no file drop of its own). Elsewhere it does nothing."""
import sys

WM_DROPFILES = 0x0233
GWLP_WNDPROC = -4


def enable(window, on_drop):
    """Call on_drop([paths]) when files are dropped on `window` (a Tk toplevel). Returns True if enabled."""
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        from ctypes import wintypes
        user32, shell32 = ctypes.windll.user32, ctypes.windll.shell32
        LRESULT = ctypes.c_ssize_t
        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
        user32.CallWindowProcW.restype = LRESULT
        user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT,
                                           wintypes.WPARAM, wintypes.LPARAM]
        setter = getattr(user32, "SetWindowLongPtrW", None) or user32.SetWindowLongW
        setter.restype = ctypes.c_void_p
        setter.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT, ctypes.c_wchar_p, wintypes.UINT]
        shell32.DragQueryFileW.restype = wintypes.UINT
        shell32.DragFinish.argtypes = [ctypes.c_void_p]
        shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]

        window.update_idletasks()
        hwnd = int(window.wm_frame(), 16)     # the real top-level window around Tk's frame

        def proc(h, msg, wp, lp):
            if msg == WM_DROPFILES:
                n = shell32.DragQueryFileW(wp, 0xFFFFFFFF, None, 0)
                paths = []
                for i in range(n):
                    size = shell32.DragQueryFileW(wp, i, None, 0) + 1
                    buf = ctypes.create_unicode_buffer(size)
                    shell32.DragQueryFileW(wp, i, buf, size)
                    paths.append(buf.value)
                shell32.DragFinish(wp)
                window.after_idle(on_drop, paths)
                return 0
            return user32.CallWindowProcW(old, h, msg, wp, lp)

        callback = WNDPROC(proc)
        old = setter(hwnd, GWLP_WNDPROC, ctypes.cast(callback, ctypes.c_void_p).value)
        window._drop_proc = callback          # keep the callback alive as long as the window
        shell32.DragAcceptFiles(hwnd, True)
        return True
    except Exception:
        return False
