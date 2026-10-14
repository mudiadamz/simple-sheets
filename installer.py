"""Simple Sheets installer / uninstaller (per user, no admin rights).

Installs to %LOCALAPPDATA%\\Programs\\SimpleSheets, adds Start menu (and optional desktop) shortcuts,
puts the app in Explorer's "Open with" list for spreadsheet files, and registers an uninstaller in
Settings > Apps. The built SimpleSheets.exe is bundled inside the setup exe ("payload").
Run with --uninstall to remove everything again.
"""
import ctypes
import os
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk

from version import APP_ID, APP_NAME, APP_VERSION, EXE_NAME

EXTS = (".xlsx", ".xlsm", ".xls", ".ods", ".csv", ".tsv")
PROGID = f"{APP_ID}.Spreadsheet"
UNINSTALL_EXE = "Uninstall.exe"
NO_WINDOW = 0x08000000   # CREATE_NO_WINDOW


def default_dir():
    return os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "Programs", APP_ID)


def payload():
    """The app exe bundled inside the setup exe (or dist/ when run from source)."""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return os.path.join(base, "payload", EXE_NAME)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", EXE_NAME)


# ---------- registry (HKEY_CURRENT_USER; `base` lets tests use a scratch key) ----------

def _set(key_path, name, value, base="Software"):
    import winreg
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, f"{base}\\{key_path}", 0, winreg.KEY_WRITE) as k:
        kind = winreg.REG_DWORD if isinstance(value, int) else winreg.REG_SZ
        winreg.SetValueEx(k, name, 0, kind, value)


def _delete_tree(key_path, base="Software"):
    import winreg
    full = f"{base}\\{key_path}"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, full, 0, winreg.KEY_ALL_ACCESS) as k:
            while True:
                try:
                    child = winreg.EnumKey(k, 0)
                except OSError:
                    break
                _delete_tree(f"{key_path}\\{child}", base)
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, full)
    except FileNotFoundError:
        pass


def _delete_value(key_path, name, base="Software"):
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, f"{base}\\{key_path}", 0, winreg.KEY_ALL_ACCESS) as k:
            winreg.DeleteValue(k, name)
    except FileNotFoundError:
        pass


def register(install_dir, open_with=True, base="Software"):
    """Explorer integration + Settings > Apps entry."""
    exe = os.path.join(install_dir, EXE_NAME)
    command = f'"{exe}" "%1"'
    if open_with:
        _set(f"Classes\\{PROGID}", "", f"Spreadsheet ({APP_NAME})", base)
        _set(f"Classes\\{PROGID}\\DefaultIcon", "", f'"{exe}",0', base)
        _set(f"Classes\\{PROGID}\\shell\\open\\command", "", command, base)
        app_key = f"Classes\\Applications\\{EXE_NAME}"
        _set(app_key, "FriendlyAppName", APP_NAME, base)
        _set(f"{app_key}\\DefaultIcon", "", f'"{exe}",0', base)
        _set(f"{app_key}\\shell\\open\\command", "", command, base)
        for ext in EXTS:
            _set(f"{app_key}\\SupportedTypes", ext, "", base)
            _set(f"Classes\\{ext}\\OpenWithProgids", PROGID, "", base)   # shows up in "Open with"
    _set(f"Microsoft\\Windows\\CurrentVersion\\App Paths\\{EXE_NAME}", "", exe, base)
    un = f"Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_ID}"
    _set(un, "DisplayName", APP_NAME, base)
    _set(un, "DisplayVersion", APP_VERSION, base)
    _set(un, "Publisher", APP_NAME, base)
    _set(un, "DisplayIcon", f'"{exe}",0', base)
    _set(un, "InstallLocation", install_dir, base)
    _set(un, "UninstallString", f'"{os.path.join(install_dir, UNINSTALL_EXE)}" --uninstall', base)
    _set(un, "NoModify", 1, base)
    _set(un, "NoRepair", 1, base)
    size_kb = sum(os.path.getsize(os.path.join(install_dir, f)) for f in os.listdir(install_dir)) // 1024
    _set(un, "EstimatedSize", size_kb, base)
    _notify_explorer()


def unregister(base="Software"):
    _delete_tree(f"Classes\\{PROGID}", base)
    _delete_tree(f"Classes\\Applications\\{EXE_NAME}", base)
    for ext in EXTS:
        _delete_value(f"Classes\\{ext}\\OpenWithProgids", PROGID, base)
    _delete_tree(f"Microsoft\\Windows\\CurrentVersion\\App Paths\\{EXE_NAME}", base)
    _delete_tree(f"Microsoft\\Windows\\CurrentVersion\\Uninstall\\{APP_ID}", base)
    _notify_explorer()


def _notify_explorer():
    try:
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)   # SHCNE_ASSOCCHANGED
    except Exception:
        pass


# ---------- shortcuts ----------

def shortcut_paths():
    start = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs")
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    return os.path.join(start, f"{APP_NAME}.lnk"), os.path.join(desktop, f"{APP_NAME}.lnk")


def make_shortcut(lnk, target):
    os.makedirs(os.path.dirname(lnk), exist_ok=True)
    ps = (f"$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{lnk}');"
          f"$s.TargetPath='{target}';$s.WorkingDirectory='{os.path.dirname(target)}';"
          f"$s.IconLocation='{target},0';$s.Description='{APP_NAME}';$s.Save()")
    subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", ps],
                   creationflags=NO_WINDOW, check=True)


# ---------- install / uninstall ----------

def install(install_dir, open_with=True, start_menu=True, desktop=False, base="Software", shortcuts=True):
    """Copy the files, register, make shortcuts. Raises PermissionError if the app is running."""
    os.makedirs(install_dir, exist_ok=True)
    shutil.copy2(payload(), os.path.join(install_dir, EXE_NAME))
    me = sys.executable if getattr(sys, "frozen", False) else None
    if me:   # the setup exe doubles as the uninstaller
        shutil.copy2(me, os.path.join(install_dir, UNINSTALL_EXE))
    register(install_dir, open_with, base)
    if shortcuts:
        exe = os.path.join(install_dir, EXE_NAME)
        start, desk = shortcut_paths()
        if start_menu:
            make_shortcut(start, exe)
        if desktop:
            make_shortcut(desk, exe)


def uninstall(install_dir, base="Software", shortcuts=True, remove_files=True):
    unregister(base)
    if shortcuts:
        for lnk in shortcut_paths():
            if os.path.exists(lnk):
                os.remove(lnk)
    if remove_files and os.path.isdir(install_dir):
        running_from_inside = os.path.abspath(sys.executable).lower().startswith(os.path.abspath(install_dir).lower())
        if running_from_inside:   # can't delete our own exe while running: let cmd do it a moment later
            subprocess.Popen(f'cmd /c ping -n 3 127.0.0.1 >nul & rmdir /s /q "{install_dir}"',
                             creationflags=NO_WINDOW, close_fds=True)
        else:
            shutil.rmtree(install_dir, ignore_errors=True)


# ---------- window ----------

class SetupWindow(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Install {APP_NAME} {APP_VERSION}")
        try:
            self.iconbitmap(default=os.path.join(getattr(sys, "_MEIPASS", ""), "assets", "simple_sheets.ico"))
        except tk.TclError:
            pass
        self.resizable(False, False)
        f = ttk.Frame(self, padding=18)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=f"{APP_NAME} {APP_VERSION}", font=("Segoe UI", 14, "bold")).pack(anchor="w")
        ttk.Label(f, text="A light viewer / editor for spreadsheet files.").pack(anchor="w", pady=(0, 12))
        ttk.Label(f, text="Install folder:").pack(anchor="w")
        self.dir = tk.StringVar(self, value=default_dir())
        ttk.Entry(f, textvariable=self.dir, width=60).pack(anchor="w", fill="x", pady=(0, 10))
        self.open_with = tk.BooleanVar(self, value=True)
        self.start_menu = tk.BooleanVar(self, value=True)
        self.desktop = tk.BooleanVar(self, value=False)
        ttk.Checkbutton(f, text="Show in \"Open with\" for " + " ".join(EXTS), variable=self.open_with).pack(anchor="w")
        ttk.Checkbutton(f, text="Start menu shortcut", variable=self.start_menu).pack(anchor="w")
        ttk.Checkbutton(f, text="Desktop shortcut", variable=self.desktop).pack(anchor="w")
        buttons = ttk.Frame(f)
        buttons.pack(anchor="e", pady=(16, 0))
        self.go = ttk.Button(buttons, text="Install", command=self.do_install)
        self.go.pack(side="left", padx=4)
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="left")
        self.status = ttk.Label(f, text="")
        self.status.pack(anchor="w", pady=(8, 0))

    def do_install(self):
        self.go["state"] = "disabled"
        self.status["text"] = "Installing…"
        self.update()
        try:
            install(self.dir.get(), self.open_with.get(), self.start_menu.get(), self.desktop.get())
        except PermissionError:
            messagebox.showerror("Install", f"Please close {APP_NAME} first, then try again.", parent=self)
            self.go["state"] = "normal"
            self.status["text"] = ""
            return
        except Exception as e:
            messagebox.showerror("Install", f"Install failed:\n{e}", parent=self)
            self.go["state"] = "normal"
            return
        if messagebox.askyesno("Installed", f"{APP_NAME} is installed.\n\nStart it now?", parent=self):
            subprocess.Popen([os.path.join(self.dir.get(), EXE_NAME)], close_fds=True)
        self.destroy()


def run_uninstall():
    root = tk.Tk()
    root.withdraw()
    here = os.path.dirname(os.path.abspath(sys.executable))
    if messagebox.askyesno("Uninstall", f"Remove {APP_NAME} from this computer?", parent=root):
        uninstall(here)
        messagebox.showinfo("Uninstall", f"{APP_NAME} was removed.", parent=root)
    root.destroy()


if __name__ == "__main__":
    if "--uninstall" in sys.argv:
        run_uninstall()
    else:
        SetupWindow().mainloop()
