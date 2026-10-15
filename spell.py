"""Spelling (Tools → Spelling, F7) through the Windows spell checker (ISpellChecker, Windows 8+), called
with ctypes — no extra libraries. Checks the text cells of the current sheet from the active cell on."""
import ctypes
import sys
import tkinter as tk
from tkinter import messagebox, ttk

CLSID_FACTORY = "{7AB36653-1796-484B-BDFA-E74F1DB7C1DC}"
IID_FACTORY = "{8E018A9D-2415-4677-BF08-794EA61F94BB}"
S_OK = 0


class GUID(ctypes.Structure):
    _fields_ = [("a", ctypes.c_ulong), ("b", ctypes.c_ushort), ("c", ctypes.c_ushort), ("d", ctypes.c_ubyte * 8)]


def _guid(s):
    g = GUID()
    ctypes.oledll.ole32.CLSIDFromString(ctypes.c_wchar_p(s), ctypes.byref(g))
    return g


def _call(ptr, index, *args, argtypes=()):
    """Call method `index` of a COM interface pointer; returns the HRESULT."""
    vtbl = ctypes.cast(ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    proto = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *argtypes)
    return proto(vtbl[index])(ptr, *args)


def _release(ptr):
    if ptr:
        _call(ptr, 2)


class Checker:
    """One language's spell checker. Raises OSError when Windows has none for that language."""

    def __init__(self, lang="en-US"):
        if sys.platform != "win32":
            raise OSError("spell checking uses the Windows spell checker")
        ole = ctypes.windll.ole32
        ole.CoInitializeEx(None, 2)                      # apartment threaded; fine if already done
        factory = ctypes.c_void_p()
        hr = ole.CoCreateInstance(ctypes.byref(_guid(CLSID_FACTORY)), None, 1, ctypes.byref(_guid(IID_FACTORY)),
                                  ctypes.byref(factory))
        if hr != S_OK:
            raise OSError("The Windows spell checker isn't available")
        try:
            ok = ctypes.c_int()
            _call(factory, 4, ctypes.c_wchar_p(lang), ctypes.byref(ok), argtypes=(ctypes.c_wchar_p, ctypes.c_void_p))
            if not ok.value:
                raise OSError(f"No spell checker for {lang}")
            self.ptr = ctypes.c_void_p()
            hr = _call(factory, 5, ctypes.c_wchar_p(lang), ctypes.byref(self.ptr),
                       argtypes=(ctypes.c_wchar_p, ctypes.c_void_p))
            if hr != S_OK:
                raise OSError("Couldn't start the spell checker")
        finally:
            _release(factory)
        self.ignored = set()

    def check(self, text):
        """[(start, length, word)] of misspelled words in `text`."""
        enum = ctypes.c_void_p()
        if _call(self.ptr, 4, ctypes.c_wchar_p(text), ctypes.byref(enum), argtypes=(ctypes.c_wchar_p, ctypes.c_void_p)) != S_OK:
            return []
        out = []
        try:
            while True:
                err = ctypes.c_void_p()
                if _call(enum, 3, ctypes.byref(err), argtypes=(ctypes.c_void_p,)) != S_OK or not err:
                    break
                start, length = ctypes.c_ulong(), ctypes.c_ulong()
                _call(err, 3, ctypes.byref(start), argtypes=(ctypes.c_void_p,))
                _call(err, 4, ctypes.byref(length), argtypes=(ctypes.c_void_p,))
                _release(err)
                word = text[start.value:start.value + length.value]
                if word.casefold() not in self.ignored:
                    out.append((start.value, length.value, word))
        finally:
            _release(enum)
        return out

    def suggest(self, word, limit=8):
        enum = ctypes.c_void_p()
        if _call(self.ptr, 5, ctypes.c_wchar_p(word), ctypes.byref(enum), argtypes=(ctypes.c_wchar_p, ctypes.c_void_p)) != S_OK:
            return []
        out = []
        try:
            while len(out) < limit:
                s, got = ctypes.c_wchar_p(), ctypes.c_ulong()
                if _call(enum, 3, 1, ctypes.byref(s), ctypes.byref(got),
                         argtypes=(ctypes.c_ulong, ctypes.c_void_p, ctypes.c_void_p)) != S_OK or not got.value:
                    break
                out.append(s.value)
                ctypes.windll.ole32.CoTaskMemFree(ctypes.cast(s, ctypes.c_void_p))
        finally:
            _release(enum)
        return out

    def add(self, word):
        """Add to the user's dictionary (shared with other Windows apps)."""
        _call(self.ptr, 6, ctypes.c_wchar_p(word), argtypes=(ctypes.c_wchar_p,))

    def ignore(self, word):
        self.ignored.add(word.casefold())

    def __del__(self):
        try:
            _release(self.ptr)
        except Exception:
            pass


def language():
    """Windows display language as a BCP-47 tag, e.g. 'en-US'."""
    buf = ctypes.create_unicode_buffer(85)
    try:
        if ctypes.windll.kernel32.GetUserDefaultLocaleName(buf, 85):
            return buf.value
    except (AttributeError, OSError):
        pass
    return "en-US"


def cells_from(g, start):
    """Text cells of the sheet in reading order, starting at `start` and wrapping around."""
    rows = g.sheet.rows
    cells = [(r, c) for r in range(len(rows)) for c in range(len(rows[r]))
             if rows[r][c] and not rows[r][c].startswith("=") and any(ch.isalpha() for ch in rows[r][c])]
    k = next((i for i, rc in enumerate(cells) if rc >= start), 0)
    return cells[k:] + cells[:k]


class SpellDialog(tk.Toplevel):
    @classmethod
    def open(cls, app, checker=None):
        try:
            checker = checker or Checker(language())
        except OSError:
            try:
                checker = Checker("en-US")
            except OSError as e:
                messagebox.showinfo("Spelling", str(e), parent=app)
                return None
        d = cls(app, checker)
        if not d.next_error():
            d.destroy()
            messagebox.showinfo("Spelling", "Spelling check complete: no mistakes found.", parent=app)
            return None
        return d

    def __init__(self, app, checker):
        super().__init__(app)
        self.app, self.checker = app, checker
        self.title("Spelling")
        self.transient(app)
        self.resizable(False, False)
        g = app.table
        self.queue = cells_from(g, g.cur)
        self.current = None        # (r, c, start, length, word)
        self.changed_all = {}
        f = ttk.Frame(self, padding=12)
        f.pack(fill="both", expand=True)
        self.context = ttk.Label(f, text="", wraplength=360)
        self.context.grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(f, text="Change to").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.change_to = tk.StringVar(master=self)
        ttk.Entry(f, textvariable=self.change_to, width=30).grid(row=2, column=0, sticky="w")
        self.list = tk.Listbox(f, height=6, width=30, exportselection=False)
        self.list.grid(row=3, column=0, sticky="nsew", pady=4)
        self.list.bind("<<ListboxSelect>>", lambda e: self.change_to.set(self.list.get(self.list.curselection()[0]))
                       if self.list.curselection() else None)
        b = ttk.Frame(f)
        b.grid(row=2, column=1, rowspan=2, sticky="n", padx=(8, 0))
        for text, cmd in (("Change", self.change), ("Change all", self.change_all), ("Ignore", self.ignore_once),
                          ("Ignore all", self.ignore_all), ("Add to dictionary", self.add), ("Close", self.destroy)):
            ttk.Button(b, text=text, command=cmd).pack(fill="x", pady=1)
        self.bind("<Escape>", lambda e: self.destroy())

    def next_error(self):
        """Move to the next misspelled word; False when there is none left."""
        g = self.app.table
        while True:
            if self.current is not None:
                r, c, start, length, _ = self.current
                errs = [e for e in self.checker.check(g.text(r, c)) if e[0] >= start + length]
            else:
                if not self.queue:
                    return False
                (r, c), start, length = self.queue.pop(0), 0, 0
                errs = self.checker.check(g.text(r, c))
            if not errs:
                self.current = None
                continue
            self.current = (r, c) + tuple(errs[0])
            word = errs[0][2]
            if word in self.changed_all:          # "Change all" earlier: fix it and go on
                self.replace(self.changed_all[word])
                continue
            self.show(r, c, *errs[0])
            return True

    def show(self, r, c, start, length, word):
        self.current = (r, c, start, length, word)
        self.app.table.goto(r, c)
        self.context.configure(text=f"Not in the dictionary: “{word}”  ({self.app.table.text(r, c)[:80]})")
        self.list.delete(0, "end")
        sugg = self.checker.suggest(word)
        for s in sugg:
            self.list.insert("end", s)
        self.change_to.set(sugg[0] if sugg else word)

    def replace(self, new):
        r, c, start, length, word = self.current
        text = self.app.table.text(r, c)
        self.app.set_cells([(r, c, text[:start] + new + text[start + length:])])
        self.current = (r, c, start, len(new), new)

    def advance(self):
        if not self.next_error():
            self.destroy()
            messagebox.showinfo("Spelling", "Spelling check complete.", parent=self.app)

    def change(self):
        if self.current:
            self.replace(self.change_to.get())
        self.advance()

    def change_all(self):
        if self.current:
            self.changed_all[self.current[4]] = self.change_to.get()
            self.replace(self.change_to.get())
        self.advance()

    def ignore_once(self):
        self.advance()

    def ignore_all(self):
        if self.current:
            self.checker.ignore(self.current[4])
        self.advance()

    def add(self):
        if self.current:
            self.checker.add(self.current[4])
            self.checker.ignore(self.current[4])
        self.advance()
