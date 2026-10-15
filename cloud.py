"""Files in OneDrive / Google Drive / Dropbox / iCloud folders: say which service syncs the file, warn about
conflict copies the sync program made next to it."""
import json
import os
import re


def roots(env=None):
    """{folder: service} for the sync folders on this computer."""
    env = os.environ if env is None else env
    home = env.get("USERPROFILE") or os.path.expanduser("~")
    out = {}
    for var in ("OneDriveCommercial", "OneDriveConsumer", "OneDrive"):
        if env.get(var):
            out.setdefault(os.path.normcase(os.path.abspath(env[var])), "OneDrive")
    for sub, name in (("Dropbox", "Dropbox"), ("Google Drive", "Google Drive"), ("iCloudDrive", "iCloud Drive"),
                      ("My Drive", "Google Drive")):
        p = os.path.join(home, sub)
        if os.path.isdir(p):
            out.setdefault(os.path.normcase(os.path.abspath(p)), name)
    info = os.path.join(env.get("LOCALAPPDATA", ""), "Dropbox", "info.json")
    try:
        with open(info, encoding="utf-8") as f:
            for acct in json.load(f).values():
                if isinstance(acct, dict) and acct.get("path"):
                    out.setdefault(os.path.normcase(os.path.abspath(acct["path"])), "Dropbox")
    except (OSError, ValueError):
        pass
    for letter in "GHIJKLMNOPQRSTUVWXYZ":                 # Google Drive for desktop: a drive with "My Drive"
        p = f"{letter}:\\My Drive"
        if env is os.environ and os.path.isdir(p):
            out.setdefault(os.path.normcase(p), "Google Drive")
    return out


def service(path, env=None):
    """'OneDrive' / 'Google Drive' / 'Dropbox' / 'iCloud Drive' when path is in a sync folder, else None."""
    p = os.path.normcase(os.path.abspath(path))
    for root, name in sorted(roots(env).items(), key=lambda kv: -len(kv[0])):
        if p == root or p.startswith(root.rstrip("\\/") + os.sep):
            return name
    return None


def conflicts(path):
    """Conflict copies next to path: 'Book-LAPTOP7.xlsx' (OneDrive), 'Book (Ann's conflicted copy).xlsx'
    (Dropbox), 'Book (1).xlsx' / 'Book (conflict).xlsx' (Google Drive and others)."""
    folder, base = os.path.split(os.path.abspath(path))
    stem, ext = os.path.splitext(base)
    device = r"(?-i:-(?=[A-Za-z0-9_-]*[A-Z])[A-Z0-9][A-Za-z0-9_-]{2,})"      # -LAPTOP7: a computer's name
    pat = re.compile(rf"^{re.escape(stem)}(?:{device}|\s\([^)]*conflict[^)]*\)|\s\(\d+\))"
                     rf"{re.escape(ext)}$", re.I)
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    return sorted(os.path.join(folder, n) for n in names if n != base and pat.match(n))


def check(app, path, env=None):
    """After opening: note the sync service in the status bar; warn about conflict copies. Returns the copies."""
    name = service(path, env)
    if not name:
        return []
    found = conflicts(path)
    note = f"Synced by {name}"
    try:
        app.status["text"] = (app.status["text"] + "  ·  " if app.status["text"] else "") + note
    except Exception:
        pass
    if found:
        from tkinter import messagebox
        shown = "\n".join("• " + os.path.basename(p) for p in found[:8])
        more = f"\n…and {len(found) - 8} more" if len(found) > 8 else ""
        if messagebox.askyesno(f"{name}: conflict copies", f"{name} kept other versions of this file next to it:\n\n"
                               f"{shown}{more}\n\nSomeone may have edited it on another computer at the same time. "
                               "Open the newest copy too, to compare?", icon="warning", parent=app):
            newest = max(found, key=lambda p: os.path.getmtime(p))
            import book_tabs
            book_tabs.open_in_tab(app, newest)
    return found
