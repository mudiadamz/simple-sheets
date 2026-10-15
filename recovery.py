"""Crash recovery (like Excel's AutoRecover): while a workbook has unsaved changes, a copy is written
every minute to <config>/recovery. The copy is removed when the workbook is saved or closed normally;
copies left behind by a window that no longer runs are offered back on the next start."""
import itertools
import json
import os
import sys
import time

from recent import config_dir

INTERVAL_MS = 60_000
_ids = itertools.count(1)


def folder():
    return os.path.join(config_dir(), "recovery")


def new_id():
    return f"{os.getpid()}-{next(_ids)}"


def paths(rid):
    base = os.path.join(folder(), rid)
    return base + ".xlsx", base + ".json"


def save(rid, sheets, names, source):
    """Write the snapshot (to a temp name first, so a crash mid-write never leaves half a file)."""
    import sheet_io as io
    xlsx, meta = paths(rid)
    os.makedirs(folder(), exist_ok=True)
    tmp = xlsx + ".tmp.xlsx"
    io.write_xlsx(tmp, sheets, names)
    os.replace(tmp, xlsx)
    with open(meta, "w", encoding="utf-8") as f:
        json.dump({"pid": os.getpid(), "source": source, "time": time.time()}, f)


def discard(rid):
    for p in paths(rid):
        try:
            os.remove(p)
        except OSError:
            pass


def alive(pid):
    if pid == os.getpid():
        return True
    if sys.platform == "win32":
        import ctypes
        k32 = ctypes.windll.kernel32
        k32.OpenProcess.restype = ctypes.c_void_p
        h = k32.OpenProcess(0x1000, False, pid)        # PROCESS_QUERY_LIMITED_INFORMATION
        if not h:
            return False
        code = ctypes.c_ulong()
        k32.GetExitCodeProcess(ctypes.c_void_p(h), ctypes.byref(code))
        k32.CloseHandle(ctypes.c_void_p(h))
        return code.value == 259                        # STILL_ACTIVE
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def orphans():
    """[(xlsx path, meta)] left by windows that are gone, newest first."""
    out = []
    try:
        names = os.listdir(folder())
    except OSError:
        return out
    for n in names:
        if not n.endswith(".json"):
            continue
        meta_path = os.path.join(folder(), n)
        xlsx = meta_path[:-5] + ".xlsx"
        try:
            with open(meta_path, encoding="utf-8") as f:
                meta = json.load(f)
        except (OSError, ValueError):
            continue
        if not os.path.exists(xlsx) or alive(int(meta.get("pid", 0))):
            continue
        out.append((xlsx, meta))
    out.sort(key=lambda x: -x[1].get("time", 0))
    return out


def forget(xlsx):
    for p in (xlsx, xlsx[:-5] + ".json"):
        try:
            os.remove(p)
        except OSError:
            pass
