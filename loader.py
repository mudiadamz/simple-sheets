"""Open big files without freezing the window: the file is read on a worker thread while a small
progress window (with Cancel) keeps the app responsive. Small files load directly."""
import os
import threading
import time
import tkinter as tk
from tkinter import ttk

import sheet_io as io

BACKGROUND_BYTES = 1024 * 1024   # files this big or bigger load in the background


class Cancelled(Exception):
    pass


def load(app, path):
    """{sheet: Sheet} for `path`; None if the user cancelled. Errors are raised as io.load raises them."""
    if os.path.getsize(path) < BACKGROUND_BYTES:
        return io.load(path)
    state = {"fraction": 0.0, "result": None, "error": None, "cancel": False, "done": False}

    def progress(f):
        if state["cancel"]:
            raise Cancelled()
        state["fraction"] = f

    def work():
        try:
            state["result"] = io.load(path, progress)
        except BaseException as e:      # handed to the UI thread
            state["error"] = e
        finally:
            state["done"] = True

    win = Progress(app, os.path.basename(path), lambda: state.update(cancel=True))
    threading.Thread(target=work, daemon=True).start()
    try:
        while not state["done"]:
            win.set(state["fraction"])
            app.update()
            time.sleep(0.03)
    finally:
        win.close()
    if isinstance(state["error"], Cancelled):
        return None
    if state["error"] is not None:
        raise state["error"]
    return state["result"]


class Progress(tk.Toplevel):
    """'Opening big.xlsx…' with a bar and Cancel; holds the input grab so nothing else runs meanwhile."""

    def __init__(self, app, name, on_cancel):
        super().__init__(app)
        self.title("Opening")
        self.transient(app)
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", on_cancel)
        box = ttk.Frame(self, padding=14)
        box.pack(fill="both", expand=True)
        ttk.Label(box, text=f"Opening {name}…").pack(anchor="w")
        self.var = tk.DoubleVar(master=self, value=0.0)
        self.bar = ttk.Progressbar(box, length=320, maximum=1.0, variable=self.var)
        self.bar.pack(fill="x", pady=8)
        self.cancel_btn = ttk.Button(box, text="Cancel", command=on_cancel)
        self.cancel_btn.pack(anchor="e")
        self.update_idletasks()
        x = app.winfo_rootx() + (app.winfo_width() - self.winfo_width()) // 2
        y = app.winfo_rooty() + (app.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        try:
            self.grab_set()
        except tk.TclError:
            pass
        app.loading = self

    def set(self, f):
        self.var.set(f)

    def close(self):
        self.master.loading = None
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()
