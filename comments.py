"""Cell comments (notes): red corner marker, hover popup, add / edit / delete dialog."""
import getpass
import tkinter as tk
from tkinter import ttk

from simplecore.theme import C


def author():
    try:
        return getpass.getuser()
    except Exception:
        return ""


class CommentDialog(tk.Toplevel):
    """One cell's comment. A note: edit it (empty deletes). A conversation (threaded comment): read it and
    reply; a new comment can start a conversation."""

    def __init__(self, app, r, c):
        super().__init__(app)
        self.app, self.r, self.c = app, r, c
        from grid import col_name
        self.title(f"Comment – {col_name(c)}{r + 1}")
        self.transient(app)
        self.configure(background=C["bg"])
        note = app.current.comments.get((r, c))
        self.thread = app.current.threads.get((r, c))
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        if self.thread:
            convo = tk.Text(f, width=44, height=8, wrap="word", background=C["bg"], foreground=C["fg"],
                            relief="flat")
            convo.tag_configure("who", font=("Segoe UI", 9, "bold"))
            for author, text, when, _ in self.thread:
                convo.insert("end", f"{author or 'Someone'}  {when[:16].replace('T', ' ')}\n", "who")
                convo.insert("end", text + "\n\n")
            convo.configure(state="disabled")
            convo.pack(fill="both", expand=True)
            self.conversation = convo
            ttk.Label(f, text="Reply").pack(anchor="w", pady=(6, 0))
        self.text = tk.Text(f, width=40, height=4 if self.thread else 6, wrap="word", background=C["field_bg"],
                            foreground=C["fg"], insertbackground=C["fg"], relief="solid", bd=1)
        self.text.pack(fill="both", expand=True)
        if note and not self.thread:
            self.text.insert("1.0", note[0])
        self.start = tk.BooleanVar(master=self, value=False)
        if not self.thread and not note:
            ttk.Checkbutton(f, text="Start a conversation (others can reply)", variable=self.start).pack(anchor="w")
        b = ttk.Frame(f)
        b.pack(fill="x", pady=(8, 0))
        ttk.Button(b, text="Reply" if self.thread else "OK", command=self.ok).pack(side="right")
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="right", padx=4)
        if self.thread:
            ttk.Button(b, text="Delete conversation", command=self.delete_thread).pack(side="left")
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Control-Return>", lambda e: self.ok())
        self.text.focus_set()

    def ok(self):
        text = self.text.get("1.0", "end-1c").strip()
        if self.thread:
            if text:
                self.app.add_reply(self.r, self.c, text)
        elif self.start.get() and text:
            self.app.add_reply(self.r, self.c, text, start=True)
        else:
            self.app.set_comment(self.r, self.c, text)
        self.destroy()

    def delete_thread(self):
        self.app.delete_thread(self.r, self.c)
        self.destroy()


class LinkDialog(tk.Toplevel):
    """Ctrl+K: text to show + link address; Remove link."""

    def __init__(self, app, r, c):
        super().__init__(app)
        self.app, self.r, self.c = app, r, c
        self.title("Link")
        self.transient(app)
        self.resizable(False, False)
        self.configure(background=C["bg"])
        g = app.table
        f = ttk.Frame(self, padding=10)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Text").grid(row=0, column=0, sticky="w")
        self.text = ttk.Entry(f, width=44)
        self.text.grid(row=0, column=1, pady=2)
        self.text.insert(0, g.text(r, c))
        ttk.Label(f, text="Link").grid(row=1, column=0, sticky="w", padx=(0, 8))
        self.url = ttk.Entry(f, width=44)
        self.url.grid(row=1, column=1, pady=2)
        current = app.current.links.get((r, c), "")
        self.url.insert(0, current)
        b = ttk.Frame(f)
        b.grid(row=2, column=0, columnspan=2, sticky="e", pady=(8, 0))
        if current:
            ttk.Button(b, text="Remove link", command=self.remove).pack(side="left", padx=(0, 12))
        ttk.Button(b, text="OK", command=self.ok).pack(side="left")
        ttk.Button(b, text="Cancel", command=self.destroy).pack(side="left", padx=4)
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Return>", lambda e: self.ok())
        (self.url if self.text.get() else self.text).focus_set()

    def ok(self):
        url = self.url.get().strip()
        if not url:
            return self.remove()
        if "://" not in url and not url.lower().startswith(("mailto:", "#")) and "@" in url:
            url = "mailto:" + url
        elif "://" not in url and not url.lower().startswith(("mailto:", "#")):
            url = "https://" + url
        self.app.set_link(self.r, self.c, self.text.get() or url, url)
        self.destroy()

    def remove(self):
        self.app.set_link(self.r, self.c, self.app.table.text(self.r, self.c), None)
        self.destroy()


class HoverNote:
    """Shows a cell's comment next to it while the mouse rests on the cell."""

    def __init__(self, grid):
        self.grid, self.cell, self.tip, self.job = grid, None, None, None

    def motion(self, x, y):
        g = self.grid
        r, c = g.hit(x, y) if g.sheet else (-1, -1)
        cell = (r, c) if r >= 0 and c >= 0 and (r, c) in g.sheet.comments else None
        if cell == self.cell:
            return
        self.hide()
        self.cell = cell
        if cell:
            self.job = g.after(350, self.show)

    def show(self):
        g, self.job = self.grid, None
        if not self.cell or self.cell not in g.sheet.comments or g.editor is not None:
            return
        r, c = self.cell
        text, who = g.sheet.comments[self.cell]
        thread = g.sheet.threads.get(self.cell)
        if thread:                          # a conversation: everyone's messages, not Excel's fallback note
            text, who = "\n".join(f"{a or 'Someone'}: {m}" for a, m, _, _ in thread), ""
        x = g.cv.winfo_rootx() + g.col_x(c + 1) + 4
        y = g.cv.winfo_rooty() + g.row_y(r)
        self.tip = tk.Toplevel(g)
        self.tip.wm_overrideredirect(True)
        self.tip.attributes("-topmost", True)
        body = f"{who}:\n{text}" if who else text
        tk.Label(self.tip, text=body, justify="left", wraplength=280, background=C["tooltip_bg"],
                 foreground=C["tooltip_fg"], relief="solid", bd=1, padx=6, pady=4).pack()
        self.tip.wm_geometry(f"+{x}+{y}")

    def hide(self, _=None):
        if self.job:
            self.grid.after_cancel(self.job)
            self.job = None
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None
        self.cell = None
