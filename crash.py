"""Unexpected errors: instead of failing quietly, show what happened (with "Copy details" for a bug
report) and keep a log in <config>/errors.log. The app keeps running; unsaved work is still there and
crash recovery keeps a copy."""
import datetime
import os
import platform
import sys
import tkinter as tk
import traceback
from tkinter import ttk

from recent import config_dir
from version import APP_VERSION

MAX_LOG = 200_000   # bytes kept in the log


def log_path():
    return os.path.join(config_dir(), "errors.log")


def details(exc_type, exc, tb):
    when = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    head = f"Simple Sheets {APP_VERSION} · Python {platform.python_version()} · {platform.platform()} · {when}"
    return head + "\n" + "".join(traceback.format_exception(exc_type, exc, tb))


def write_log(text):
    try:
        os.makedirs(config_dir(), exist_ok=True)
        path = log_path()
        old = ""
        if os.path.exists(path):
            with open(path, encoding="utf-8", errors="replace") as f:
                old = f.read()[-MAX_LOG:]
        with open(path, "w", encoding="utf-8") as f:
            f.write(old + text + "\n" + "-" * 60 + "\n")
    except OSError:
        pass


class ErrorDialog(tk.Toplevel):
    def __init__(self, app, text):
        super().__init__(app)
        from theme import C
        self.title("Something went wrong")
        self.transient(app)
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Simple Sheets ran into an unexpected problem. Your work is still open - save it if "
                          "you can.\nThe details below help to fix it (they're also in errors.log).",
                  wraplength=520).pack(anchor="w")
        box = tk.Text(f, height=12, width=80, wrap="none", font=("Consolas", 9),
                      background=C["field_bg"], foreground=C["fg"])
        box.insert("1.0", text)
        box.configure(state="disabled")
        box.pack(fill="both", expand=True, pady=8)
        b = ttk.Frame(f)
        b.pack(anchor="e")
        ttk.Button(b, text="Copy details", command=lambda: self.copy(text)).pack(side="left", padx=4)
        ttk.Button(b, text="Continue", command=self.destroy).pack(side="left")
        self.text = text
        self.bind("<Escape>", lambda e: self.destroy())

    def copy(self, text):
        self.clipboard_clear()
        self.clipboard_append(text)


def report(app, exc_type, exc, tb):
    """Tk calls this for an exception in a callback (App.report_callback_exception)."""
    text = details(exc_type, exc, tb)
    write_log(text)
    app.last_error = text
    print(text, file=sys.stderr)
    if getattr(app, "_error_dialog_open", False):   # one dialog at a time (errors can repeat quickly)
        return None
    app._error_dialog_open = True
    try:
        d = ErrorDialog(app, text)
        d.bind("<Destroy>", lambda e: setattr(app, "_error_dialog_open", False) if e.widget is d else None)
        return d
    except tk.TclError:
        app._error_dialog_open = False
        return None
