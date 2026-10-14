"""The ☰ app menu: new, open, open recent, save, save as, undo/redo, exit."""
import os
import tkinter as tk
from tkinter import ttk

import recent


def build(app, parent):
    """☰ button (with its menu) plus small undo/redo buttons, packed left into `parent`."""
    mb = ttk.Menubutton(parent, text="☰", style="Icon.Toolbutton", width=2)
    mb.pack(side="left")
    menu = tk.Menu(mb, tearoff=False)
    recent_menu = tk.Menu(menu, tearoff=False, postcommand=lambda: fill_recent(app, recent_menu))
    menu.add_command(label="New", accelerator="Ctrl+N", command=app.new_window)
    menu.add_command(label="Open…", accelerator="Ctrl+O", command=app.open_dialog)
    menu.add_cascade(label="Open recent", menu=recent_menu)
    menu.add_separator()
    menu.add_command(label="Save", accelerator="Ctrl+S", command=app.save)
    menu.add_command(label="Save as…", accelerator="Ctrl+Shift+S", command=app.save_as)
    menu.add_separator()
    menu.add_command(label="Undo", accelerator="Ctrl+Z", command=app.undo)
    menu.add_command(label="Redo", accelerator="Ctrl+Y", command=app.redo)
    menu.add_separator()
    menu.add_checkbutton(label="Dark mode", variable=app.dark_var, command=app.toggle_dark)
    menu.add_separator()
    menu.add_command(label="Exit", command=app.on_close)
    mb["menu"] = menu
    app.app_menu, app.recent_menu = menu, recent_menu

    from format_bar import Tooltip
    Tooltip(mb, "Menu")
    # save: blue disk when there are unsaved changes, grey (disabled) when everything is saved
    app.save_icons = {"on": disk_icon(app, app.scale, "#1a73e8"), "off": disk_icon(app, app.scale, "#b5b5b5")}
    app.save_btn = ttk.Button(parent, image=(app.save_icons["on"], "disabled", app.save_icons["off"]),
                              style="Icon.Toolbutton", command=app.save, state="disabled")
    app.save_btn.pack(side="left")
    Tooltip(app.save_btn, "Save (Ctrl+S)")
    app.undo_btn = ttk.Button(parent, text="↶", style="Icon.Toolbutton", width=2, command=app.undo, state="disabled")
    app.undo_btn.pack(side="left")
    Tooltip(app.undo_btn, "Undo (Ctrl+Z)")
    app.redo_btn = ttk.Button(parent, text="↷", style="Icon.Toolbutton", width=2, command=app.redo, state="disabled")
    app.redo_btn.pack(side="left")
    Tooltip(app.redo_btn, "Redo (Ctrl+Y)")
    ttk.Separator(parent, orient="vertical").pack(side="left", fill="y", padx=6, pady=(0, 4))


def disk_icon(master, scale, color):
    """Floppy disk: coloured body, white shutter at the top, white label at the bottom."""
    s = max(int(16 * scale), 16)
    k = s / 16
    img = tk.PhotoImage(master=master, width=s, height=s)

    def box(x1, y1, x2, y2, c):
        img.put(c, to=(int(x1 * k), int(y1 * k), max(int(x2 * k), int(x1 * k) + 1), max(int(y2 * k), int(y1 * k) + 1)))
    box(1, 1, 15, 15, color)          # body
    box(13, 1, 15, 3, "#ffffff")       # clipped corner
    box(4, 1, 11, 6, "#ffffff")        # metal shutter
    box(8, 2, 10, 5, color)            # shutter slot
    box(3, 9, 13, 15, "#ffffff")       # label
    box(4, 11, 12, 12, color)          # a line on the label
    return img


def fill_recent(app, menu):
    """Rebuilt each time the submenu opens."""
    menu.delete(0, "end")
    items = recent.load()
    if not items:
        menu.add_command(label="(none yet)", state="disabled")
        return
    for i, path in enumerate(items):
        label = f"{i + 1}  {os.path.basename(path)}   —   {os.path.dirname(path)}"
        menu.add_command(label=label, underline=0, command=lambda p=path: app.open_recent(p))
    menu.add_separator()
    menu.add_command(label="Clear list", command=recent.clear)
