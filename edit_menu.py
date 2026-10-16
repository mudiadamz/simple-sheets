"""Right-click menu for text boxes (Text / Entry): Undo, Cut, Copy, Paste, Delete, Select all.
Read-only boxes (like the AI chat) get only Copy and Select all. build() makes the menu without showing it
(tests use it: a popped-up menu blocks on Windows); attach() shows it on a right click."""
import tkinter as tk


def _is_text(w):
    return isinstance(w, tk.Text)


def has_selection(w):
    try:
        return bool(w.tag_ranges("sel")) if _is_text(w) else w.selection_present()
    except tk.TclError:
        return False


def editable(w):
    return str(w.cget("state")) not in ("disabled", "readonly")


def select_all(w):
    if _is_text(w):
        w.tag_add("sel", "1.0", "end-1c")
        w.mark_set("insert", "end-1c")
    else:
        w.selection_range(0, "end")
        w.icursor("end")
    w.focus_set()


def delete_selection(w):
    if has_selection(w):
        w.delete("sel.first", "sel.last")


def clipboard_text(w):
    try:
        return w.clipboard_get()
    except tk.TclError:
        return ""


def build(w, paste=None, extra=None):
    """The menu for widget w. paste: own paste command (e.g. the AI box also takes pictures);
    extra: [(label, command)] added at the top (e.g. Open link)."""
    m = tk.Menu(w, tearoff=False)
    sel, can_edit = has_selection(w), editable(w)
    for label, cmd in extra or ():
        m.add_command(label=label, command=cmd)
    if extra:
        m.add_separator()
    if can_edit and _is_text(w) and str(w.cget("undo")) in ("1", "true"):
        m.add_command(label="Undo", accelerator="Ctrl+Z", command=lambda: w.event_generate("<<Undo>>"))
        m.add_command(label="Redo", accelerator="Ctrl+Y", command=lambda: w.event_generate("<<Redo>>"))
        m.add_separator()
    if can_edit:
        m.add_command(label="Cut", accelerator="Ctrl+X", command=lambda: w.event_generate("<<Cut>>"),
                      state="normal" if sel else "disabled")
    m.add_command(label="Copy", accelerator="Ctrl+C", command=lambda: w.event_generate("<<Copy>>"),
                  state="normal" if sel else "disabled")
    if can_edit:
        m.add_command(label="Paste", accelerator="Ctrl+V",
                      command=paste or (lambda: w.event_generate("<<Paste>>")))
        m.add_command(label="Delete", command=lambda: delete_selection(w), state="normal" if sel else "disabled")
    m.add_separator()
    m.add_command(label="Select all", accelerator="Ctrl+A", command=lambda: select_all(w))
    return m


def attach(w, paste=None, extra=None):
    """Right click on w shows its edit menu. extra: a function (event) -> [(label, command)] for that spot."""
    def popup(e):
        w.focus_set()
        m = build(w, paste, extra(e) if extra else None)
        try:
            m.tk_popup(e.x_root, e.y_root)
        finally:
            m.grab_release()
        return "break"
    w.bind("<Button-3>", popup, add="+")
    return popup
