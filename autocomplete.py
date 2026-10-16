"""Formula autocomplete: while typing `=SU…` a small list of matching functions (and named ranges)
drops down under the editor; inside `FUNC(` a hint shows the arguments, the current one in bold.
Works on the in-cell editor and on the formula bar. Keys: ↑/↓ choose, Tab/Enter insert, Esc close."""
import re
import tkinter as tk
from tkinter import font as tkfont

import formula
from theme import C

SIGNATURES = {
    "SUM": "number1, [number2], …", "AVERAGE": "number1, [number2], …", "MIN": "number1, [number2], …",
    "MAX": "number1, [number2], …", "COUNT": "value1, [value2], …", "COUNTA": "value1, [value2], …",
    "COUNTBLANK": "range", "PRODUCT": "number1, [number2], …", "ROUND": "number, digits",
    "ROUNDUP": "number, digits", "ROUNDDOWN": "number, digits", "INT": "number", "MOD": "number, divisor",
    "ABS": "number", "SQRT": "number", "POWER": "number, power", "IF": "condition, value_if_true, [value_if_false]",
    "IFS": "condition1, value1, [condition2, value2], …", "IFERROR": "value, value_if_error",
    "AND": "logical1, [logical2], …", "OR": "logical1, [logical2], …", "NOT": "logical",
    "CHOOSE": "index, value1, [value2], …", "CONCATENATE": "text1, [text2], …", "CONCAT": "text1, [text2], …",
    "TEXTJOIN": "delimiter, ignore_empty, text1, …", "LEN": "text", "UPPER": "text", "LOWER": "text",
    "PROPER": "text", "TRIM": "text", "LEFT": "text, [count]", "RIGHT": "text, [count]",
    "MID": "text, start, count", "SUBSTITUTE": "text, old, new, [instance]", "REPLACE": "text, start, count, new",
    "FIND": "find, within, [start]", "SEARCH": "find, within, [start]", "VALUE": "text", "TEXT": "value, format",
    "REPT": "text, times", "SUMIF": "range, criteria, [sum_range]", "SUMIFS": "sum_range, range1, criteria1, …",
    "COUNTIF": "range, criteria", "COUNTIFS": "range1, criteria1, …", "AVERAGEIF": "range, criteria, [avg_range]",
    "AVERAGEIFS": "avg_range, range1, criteria1, …", "VLOOKUP": "value, table, column, [approximate]",
    "HLOOKUP": "value, table, row, [approximate]", "XLOOKUP": "value, lookup_range, return_range, [if_not_found]",
    "INDEX": "range, row, [column]", "MATCH": "value, range, [match_type]", "TODAY": "", "NOW": "",
    "DATE": "year, month, day", "YEAR": "date", "MONTH": "date", "DAY": "date", "WEEKDAY": "date, [type]",
    "DAYS": "end_date, start_date", "MEDIAN": "number1, [number2], …", "LARGE": "range, k", "SMALL": "range, k",
    "PI": "", "RAND": "", "RANDBETWEEN": "low, high", "ISBLANK": "value", "ISNUMBER": "value", "ISTEXT": "value",
    "ISERROR": "value", "PMT": "rate, nper, pv, [fv], [type]", "FV": "rate, nper, pmt, [pv], [type]",
    "PV": "rate, nper, pmt, [fv], [type]", "NPER": "rate, pmt, pv, [fv], [type]",
    "RATE": "nper, pmt, pv, [fv], [type], [guess]", "NPV": "rate, value1, [value2], …", "IRR": "values, [guess]",
    "STDEV": "number1, [number2], …", "STDEV.P": "number1, [number2], …", "VAR": "number1, [number2], …",
    "VAR.P": "number1, [number2], …", "MODE": "number1, [number2], …", "PERCENTILE": "range, k",
    "QUARTILE": "range, quart", "RANK": "number, range, [ascending]", "COUNTUNIQUE": "value1, [value2], …",
    "SUMPRODUCT": "range1, [range2], …", "SUMSQ": "number1, [number2], …", "CEILING": "number, [significance]",
    "FLOOR": "number, [significance]", "MROUND": "number, multiple", "FACT": "number", "GCD": "number1, …",
    "LCM": "number1, …", "EDATE": "start_date, months", "EOMONTH": "start_date, months",
    "NETWORKDAYS": "start_date, end_date, [holidays]", "WORKDAY": "start_date, days, [holidays]",
    "DATEDIF": "start_date, end_date, unit", "TIME": "hour, minute, second", "HOUR": "time", "MINUTE": "time",
    "SECOND": "time", "DATEVALUE": "text", "TIMEVALUE": "text",
    "IMAGE": "source, [alt_text], [sizing], [height], [width]", "CHAR": "number", "CODE": "text", "CLEAN": "text",
    "FIXED": "number, [decimals], [no_commas]", "ROW": "[reference]", "COLUMN": "[reference]",
    "OFFSET": "reference, rows, cols, [height], [width]", "INDIRECT": "ref_text",
    "ADDRESS": "row, column, [abs], [a1], [sheet]",
    "SEQUENCE": "rows, [columns], [start], [step]", "FILTER": "array, include, [if_empty]",
    "UNIQUE": "array, [by_col], [exactly_once]", "SORT": "array, [sort_index], [sort_order], [by_col]",
    "SORTBY": "array, by_array1, [order1], …", "TRANSPOSE": "array", "VSTACK": "array1, [array2], …",
    "HSTACK": "array1, [array2], …", "TAKE": "array, rows, [columns]", "DROP": "array, rows, [columns]",
    "TEXTSPLIT": "text, col_delimiter, [row_delimiter], [ignore_empty]", "XMATCH": "value, array, [match_mode], [search_mode]",
    "LET": "name1, value1, [name2, value2], …, calculation", "LAMBDA": "[parameter1, …], calculation",
    "MAP": "array, lambda", "BYROW": "array, lambda", "BYCOL": "array, lambda",
    "REDUCE": "initial_value, array, lambda", "SCAN": "initial_value, array, lambda",
    "MAKEARRAY": "rows, columns, lambda", "SWITCH": "value, match1, result1, …, [default]",
    "TEXTBEFORE": "text, delimiter, [instance]", "TEXTAFTER": "text, delimiter, [instance]",
    "WEEKNUM": "date, [return_type]", "ISOWEEKNUM": "date", "MAXIFS": "max_range, range1, criteria1, …",
    "MINIFS": "min_range, range1, criteria1, …", "IFNA": "value, value_if_na", "XOR": "logical1, [logical2], …",
}

WORD = re.compile(r"[A-Za-z_][A-Za-z0-9_.]*$")
MAX_ITEMS = 8


def context(text, cur):
    """(word being typed, its start) and (innermost open function, argument index) at `cur`."""
    head = text[:cur]
    if not head.startswith("="):
        return None, None
    calls, word_start, in_str = [], None, False
    i = 1
    while i < len(head):
        ch = head[i]
        if ch == '"':
            in_str = not in_str
        elif in_str:
            pass
        elif ch == "'":                       # quoted sheet name: skip to its end
            j = head.find("'", i + 1)
            if j < 0:
                return None, None
            i = j
        elif ch == "(":
            m = WORD.search(head[:i])
            calls.append([m.group().upper() if m else "", 0])
        elif ch == ")":
            if calls:
                calls.pop()
        elif ch == "," and calls:
            calls[-1][1] += 1
        i += 1
    if in_str:
        return None, None
    m = WORD.search(head)
    word = None
    if m:
        word_start = m.start()
        before = head[word_start - 1] if word_start else ""
        if before not in "!$:" and not (word_start and head[word_start - 1].isdigit()):
            word = (m.group(), word_start)
    call = next(((n, k) for n, k in reversed(calls) if n), None)
    return word, call


class Autocomplete:
    def __init__(self, widget, app):
        self.w, self.app = widget, app
        self.top = self.listbox = self.hint = None
        self.items, self.word = [], None
        tag = f"autocomplete{id(self)}"
        widget.bindtags((tag,) + widget.bindtags())   # runs before the widget's own Return/Tab/Escape
        widget.bind_class(tag, "<KeyRelease>", self.on_key)
        for seq, f in (("<Down>", lambda: self.move(1)), ("<Up>", lambda: self.move(-1)),
                       ("<Tab>", self.accept), ("<Return>", self.accept), ("<KP_Enter>", self.accept),
                       ("<Escape>", self.close)):
            widget.bind_class(tag, seq, lambda e, f=f: "break" if self.items and f() is not False else None)
        widget.bind_class(tag, "<FocusOut>", lambda e: widget._root().after(150, self.close_if_unfocused))
        widget.bind_class(tag, "<Destroy>", lambda e: self.destroy())

    # ---- what to suggest ----
    def candidates(self, prefix):
        p = prefix.upper()
        names = sorted(n for n in formula.FUNCS if n.startswith(p))
        names.sort(key=lambda n: (n != p, "." in n, len(n)))      # exact, then plain, then short first
        named = sorted(n for n in getattr(self.app, "names", {}) or {} if n.upper().startswith(p))
        return (names + named)[:MAX_ITEMS]

    def on_key(self, e):
        if e.keysym in ("Up", "Down", "Return", "KP_Enter", "Tab", "Escape", "Shift_L", "Shift_R",
                        "Control_L", "Control_R"):
            return
        self.refresh()

    def refresh(self):
        try:
            text, cur = self.w.get(), self.w.index("insert")
        except tk.TclError:
            return
        word, call = context(text, cur)
        follows = text[cur:cur + 1]
        self.items, self.word = [], None
        if word and follows != "(":
            items = self.candidates(word[0])
            if items:
                self.items, self.word = items, word
        hint = None
        if call and call[0] in formula.FUNCS and not self.items:
            hint = call
        if not self.items and not hint:
            self.hide()
            return
        self.show(hint)

    # ---- popup ----
    def show(self, hint):
        if self.top is None:
            self.top = tk.Toplevel(self.w)
            self.top.overrideredirect(True)
            self.top.attributes("-topmost", True)
            self.top.configure(background=C["header_line"])
            base = tkfont.nametofont("TkDefaultFont", root=self.w)
            self.bold = base.copy()
            self.bold.configure(weight="bold")
            self.hint = tk.Text(self.top, height=1, width=10, relief="flat", bd=0, padx=6, pady=3,
                                background=C["field_bg"], foreground=C["fg"], font=base, cursor="arrow")
            self.hint.tag_configure("arg", font=self.bold)
            self.hint.tag_configure("muted", foreground=C["muted"])
            self.listbox = tk.Listbox(self.top, height=MAX_ITEMS, activestyle="none", relief="flat", bd=0,
                                      exportselection=False, background=C["field_bg"], foreground=C["fg"],
                                      selectbackground=C["menu_active"], selectforeground=C["fg"],
                                      highlightthickness=0, font=base)
            self.listbox.bind("<Button-1>", self.click)   # handled here: the editor keeps the focus
        self.hint.pack_forget()
        self.listbox.pack_forget()
        if self.items:
            self.listbox.delete(0, "end")
            for n in self.items:
                self.listbox.insert("end", n)
            self.listbox.configure(height=len(self.items), width=max(12, max(len(n) for n in self.items) + 2))
            self.listbox.selection_set(0)
            self.listbox.pack(fill="both", padx=1, pady=1)
            first = self.items[0]
            if first in SIGNATURES:          # the hint for the highlighted function, too
                self.set_hint(first, -1)
                self.hint.pack(fill="x", padx=1, pady=(0, 1))
        elif hint:
            self.set_hint(*hint)
            self.hint.pack(fill="x", padx=1, pady=1)
        self.place()
        self.top.deiconify()

    def set_hint(self, name, index):
        args = SIGNATURES.get(name, "…")
        parts = [a.strip() for a in args.split(",")] if args else []
        self.hint.configure(state="normal")
        self.hint.delete("1.0", "end")
        self.hint.insert("end", name + "(")
        repeat = parts and parts[-1] == "…"
        for k, p in enumerate(parts):
            if k:
                self.hint.insert("end", ", ")
            current = k == index or (repeat and index >= len(parts) - 1 and k == len(parts) - 2)
            self.hint.insert("end", p, ("arg",) if current else ())
        self.hint.insert("end", ")")
        self.hint.configure(width=len(self.hint.get("1.0", "end-1c")) + 2, state="disabled")

    def place(self):
        try:
            box = self.w.bbox("insert") if isinstance(self.w, tk.Text) else self.w.bbox(self.w.index("insert"))
        except tk.TclError:
            box = None
        x = self.w.winfo_rootx() + (box[0] if box else 0)
        y = self.w.winfo_rooty() + self.w.winfo_height()
        self.top.geometry(f"+{x}+{y}")

    def hide(self):
        self.items = []
        if self.top is not None:
            self.top.withdraw()

    def close(self):
        self.hide()

    def close_if_unfocused(self):
        try:
            if self.w.focus_get() is not self.w:
                self.hide()
        except (tk.TclError, KeyError):
            self.hide()

    def destroy(self):
        if self.top is not None:
            self.top.destroy()
            self.top = None
        self.items = []

    # ---- keys ----
    def click(self, e):
        i = self.listbox.nearest(e.y)
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(i)
        self.accept()
        return "break"

    def move(self, d):
        i = (self.listbox.curselection() or (0,))[0]
        i = max(0, min(len(self.items) - 1, i + d))
        self.listbox.selection_clear(0, "end")
        self.listbox.selection_set(i)
        self.listbox.see(i)
        if self.items[i] in SIGNATURES:
            self.set_hint(self.items[i], -1)
            if not self.hint.winfo_ismapped():
                self.hint.pack(fill="x", padx=1, pady=(0, 1))
        else:
            self.hint.pack_forget()

    def accept(self):
        if not self.items or not self.word:
            return False
        i = (self.listbox.curselection() or (0,))[0]
        name = self.items[i]
        _, start = self.word
        cur = self.w.index("insert")
        self.w.delete(start, cur)
        is_func = name in formula.FUNCS
        self.w.insert(start, name + ("(" if is_func else ""))
        self.w.icursor(start + len(name) + (1 if is_func else 0))
        if hasattr(self.w, "fit"):
            self.w.fit()
        self.hide()
        if is_func:
            self.refresh()               # now show the argument hint
        show_refs = getattr(getattr(self.app, "table", None), "show_refs", None)
        if show_refs:
            show_refs(self.w.get())
        return True


def attach(widget, app):
    widget.autocomplete = Autocomplete(widget, app)
    return widget.autocomplete
