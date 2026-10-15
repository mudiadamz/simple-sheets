"""Help → Check for updates: reads a small JSON file (the release feed set in Tools → Settings, e.g. the
GitHub API 'releases/latest' address) and says whether a newer version exists.

Automatic updates: when the feed names the setup file and its SHA-256, the setup is downloaded to the
settings folder and checked against that hash; it is run silently (Setup.exe /S /relaunch) the next time
Simple Sheets starts, before the window opens. A download whose hash doesn't match is deleted, never run.
Feed: {"version", "url", "download", "sha256"} or a GitHub release (assets + a "SHA-256: …" line in its
notes, or a "<setup>.sha256" asset)."""
import hashlib
import json
import os
import re
import subprocess
import sys
import urllib.request
import webbrowser
from tkinter import messagebox

import settings
from version import APP_VERSION

SETUP_NAME = re.compile(r"setup.*\.exe$", re.I)


def parse_version(v):
    return tuple(int(x) for x in re.findall(r"\d+", str(v))[:4]) or (0,)


def _get(url, timeout=8):
    req = urllib.request.Request(url, headers={"User-Agent": f"SimpleSheets/{APP_VERSION}"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def info(url, timeout=8):
    """The feed -> {"version", "page", "download", "sha256"} (download / sha256 may be "")."""
    data = json.loads(_get(url, timeout).decode("utf-8"))
    version = str(data.get("version") or data.get("tag_name") or data.get("name") or "").lstrip("vV")
    out = {"version": version, "page": data.get("url") or data.get("html_url") or "",
           "download": data.get("download") or "", "sha256": (data.get("sha256") or "").lower()}
    assets = data.get("assets") or []
    if not out["download"]:
        setup = next((a for a in assets if SETUP_NAME.search(a.get("name", ""))), None)
        if setup:
            out["download"] = setup.get("browser_download_url") or ""
            if not out["sha256"]:
                side = next((a for a in assets if a.get("name", "").lower() == setup["name"].lower() + ".sha256"), None)
                if side:
                    try:
                        found = re.search(r"\b[0-9a-fA-F]{64}\b", _get(side["browser_download_url"], timeout).decode())
                        out["sha256"] = found.group().lower() if found else ""
                    except Exception:
                        pass
    if not out["sha256"]:
        m = re.search(r"SHA-?256\W+([0-9a-fA-F]{64})", data.get("body") or "", re.I)
        out["sha256"] = m.group(1).lower() if m else ""
    return out


def latest(url, timeout=8):
    """(version, page url) - kept for callers that only need those."""
    i = info(url, timeout)
    return i["version"], i["page"]


def folder():
    import recent
    return os.path.join(recent.config_dir(), "updates")


def pending_file():
    return os.path.join(folder(), "pending.json")


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def download(i, timeout=60):
    """Fetch the setup named by info `i`, check its SHA-256, remember it for the next start. -> path.
    Raises ValueError when the hash is missing or doesn't match (the file is deleted)."""
    if not i.get("download") or not re.fullmatch(r"[0-9a-f]{64}", i.get("sha256") or ""):
        raise ValueError("The update has no SHA-256 checksum, so it isn't installed automatically.")
    os.makedirs(folder(), exist_ok=True)
    dst = os.path.join(folder(), f"SimpleSheets-Setup-{i['version']}.exe")
    tmp = dst + ".part"
    req = urllib.request.Request(i["download"], headers={"User-Agent": f"SimpleSheets/{APP_VERSION}"})
    with urllib.request.urlopen(req, timeout=timeout) as r, open(tmp, "wb") as f:
        while block := r.read(1 << 16):
            f.write(block)
    if sha256_of(tmp) != i["sha256"]:
        os.remove(tmp)
        raise ValueError("The downloaded update didn't match its SHA-256 checksum and was deleted.")
    os.replace(tmp, dst)
    with open(pending_file(), "w", encoding="utf-8") as f:
        json.dump({"version": i["version"], "path": dst, "sha256": i["sha256"]}, f)
    return dst


def pending():
    try:
        with open(pending_file(), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def clear_pending():
    p = pending()
    for path in ((p or {}).get("path"), pending_file()):
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                pass


def apply_pending(run=None):
    """At start: a downloaded, verified newer setup is run silently (it closes nothing: this process exits
    right after and the setup waits for it, then starts the new version). Returns True when the caller
    should exit. Old or tampered downloads are removed."""
    p = pending()
    if not p:
        return False
    path = p.get("path") or ""
    if parse_version(p.get("version")) <= parse_version(APP_VERSION) or not os.path.exists(path) \
            or sha256_of(path) != p.get("sha256"):
        clear_pending()
        return False
    os.remove(pending_file())               # once only, even if the setup fails
    run = run or (lambda args: subprocess.Popen(args, close_fds=True))
    run([path, "/S", "/relaunch"])
    return True


def check(app, url=None, quiet=False, known=None):
    """known: the feed already read (by the start-up check)."""
    url = url or settings.get("update_url")
    if not url and known is None:
        if not quiet:
            messagebox.showinfo("Check for updates", "No update address is set.\n\n"
                                "Add one in Tools → Settings (for example a GitHub releases feed).", parent=app)
        return None
    try:
        i = known or info(url)
    except Exception as e:
        if not quiet:
            messagebox.showwarning("Check for updates", f"Couldn't check for updates:\n{e}", parent=app)
        return None
    newer = parse_version(i["version"]) > parse_version(APP_VERSION)
    if newer:
        p = pending()
        if p and p.get("version") == i["version"]:
            if not quiet:
                messagebox.showinfo("Update ready", f"Version {i['version']} is downloaded; it is installed the "
                                    "next time Simple Sheets starts.", parent=app)
            return newer
        # only the built app installs itself (from source there is nothing to replace)
        # and not a portable copy (the setup would install a second copy in the user profile)
        import recent
        can_auto = bool(i["download"] and i["sha256"]) and (getattr(sys, "frozen", False) and not recent.portable_dir()
                                                           or getattr(app, "allow_dev_update", False))
        if can_auto:
            if messagebox.askokcancel("Update available", f"Version {i['version']} is available (you have "
                                      f"{APP_VERSION}).\n\nDownload it now? It is checked and installed the next "
                                      "time you start Simple Sheets.", parent=app):
                try:
                    download(i)
                    app.status["text"] = f"Update {i['version']} downloaded: installed at the next start"
                except Exception as e:
                    messagebox.showwarning("Update", f"The update couldn't be downloaded:\n{e}", parent=app)
        elif messagebox.askokcancel("Update available", f"Version {i['version']} is available (you have "
                                    f"{APP_VERSION}).\n\nOpen the download page?", parent=app) and i["page"]:
            webbrowser.open(i["page"])
    elif not quiet:
        messagebox.showinfo("Check for updates", f"You have the latest version ({APP_VERSION}).", parent=app)
    return newer


def check_in_background(app, url=None):
    """The start-up check: the network part on a thread, the questions back on the window's thread."""
    import queue
    import threading
    url = url or settings.get("update_url")
    q = queue.Queue()

    def work():
        try:
            q.put(info(url))
        except Exception:
            q.put(None)
    threading.Thread(target=work, daemon=True).start()

    def poll():
        try:
            i = q.get_nowait()
        except queue.Empty:
            app.after(200, poll)
            return
        if i and parse_version(i["version"]) > parse_version(APP_VERSION):
            check(app, url, quiet=True, known=i)
    app.after(200, poll)
