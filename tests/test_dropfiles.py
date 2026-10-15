"""Drag & drop files onto the window."""
import ctypes
import struct
import sys

import pytest

import workbooks


def test_open_dropped_logic(app, tmp_path, monkeypatch):
    spawned = []
    monkeypatch.setattr(app, "new_window", lambda path=None: spawned.append(path))
    a = workbooks.plain(tmp_path / "a.xlsx")
    b = workbooks.sales(tmp_path / "b.xlsx")
    junk = tmp_path / "notes.txt"
    junk.write_text("x")
    app.open_dropped([str(junk)])
    assert app.path is None and app.dialogs[-1][0] == "showwarning"
    app.open_dropped([str(a), str(junk), str(b)])
    assert app.path == str(a) and spawned == [str(b)]      # empty window: first one opens here
    app.open_dropped([str(a)])
    assert spawned == [str(b), str(a)]                     # a file is open: new window


@pytest.mark.skipif(sys.platform != "win32", reason="Windows drop messages")
def test_real_drop_message(app, tmp_path):
    assert app.drop_enabled
    path = str(workbooks.plain(tmp_path / "dropped.xlsx"))
    data = struct.pack("<IiiII", 20, 0, 0, 0, 1) + (path + "\0\0").encode("utf-16-le")
    k32 = ctypes.windll.kernel32
    k32.GlobalAlloc.restype = ctypes.c_void_p
    k32.GlobalLock.restype = ctypes.c_void_p
    k32.GlobalLock.argtypes = k32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    h = k32.GlobalAlloc(0x0042, len(data))                 # GHND (moveable, zeroed)
    ctypes.memmove(k32.GlobalLock(h), data, len(data))
    k32.GlobalUnlock(h)
    u32 = ctypes.windll.user32
    u32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_void_p]
    u32.SendMessageW(int(app.wm_frame(), 16), 0x0233, h, None)
    app.update()
    assert app.path == path and app.table.text(0, 0)
