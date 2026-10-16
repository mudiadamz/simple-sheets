"""Tools → AI assistant: ask an AI model (through OpenRouter, with your own API key) about the sheet or to
change it. The model answers in words and, when cells should change, with a script for the script engine
(scripting.py). The script is shown first; it only runs when you press Run (one undo step, nothing changes
if it fails).

Settings (Tools → Settings… → AI, or the assistant's ⚙): your OpenRouter API key (kept encrypted for your
Windows account) and the model. What is sent: the question, the sheet's names and the selected cells or
the sheet's data (up to MAX_CELLS cells) - to OpenRouter and the model's provider.
"""
import base64
import json
import queue
import re
import sys
import threading
import urllib.error
import urllib.request

import settings
from formula import col_letters
from version import APP_VERSION

API = "https://openrouter.ai/api/v1"
MAX_CELLS = 3000          # cells sent with a question (the rest is summarised as "… more rows")
MAX_TEXT = 200            # characters per cell sent
HISTORY = 12              # messages of the conversation kept

SYSTEM = """You are the assistant inside Simple Sheets, a spreadsheet app. You see the user's sheet as a table
with cell addresses. Answer questions about the data briefly and clearly.

When the user wants cells changed (values, formulas, formats, new columns, cleaning, sorting by writing values,
summaries...), reply with ONE short explanation sentence and ONE Python script in a ```python block. The app
shows the script to the user, who runs it. The script runs inside the app with these names:

sheet                    the active sheet; selection = the selected block; book = the workbook
sheet["A1"]              read a value (number, text, True/False, None when empty; formulas come back calculated)
sheet["A1:C3"]           list of rows;  sheet["B:B"] a whole column
sheet["A1"] = 5          write; text starting with "=" is a formula; None empties
sheet["A3"] = [[1, 2], [3, 4]]   rows spread from A3;  sheet["E1:E9"] = 0 fills a block
sheet.cell(row, col)     1-based read;  sheet.cell(row, col, value) write
sheet.formula("B2")      the typed text
sheet.format("A1:C1", bold=True, italic=True, fg="#rrggbb", bg="#rrggbb", size=12, numfmt="0.00",
             align="left|center|right", wrap=True,
             border="all|outline|top|bottom|left|right|none", border_style="thin|medium|thick|double|dashed|dotted",
             border_color="#rrggbb")   (False removes a format)
sheet.freeze(1) (top row) / sheet.freeze(cols=1) / sheet.freeze("B2") / sheet.unfreeze()
sheet.merge("A1:D1") / sheet.unmerge("A1:D1")
sheet.width("B", 20) / sheet.width("B:D", 12) / sheet.autofit() / sheet.autofit("A:C") / sheet.height(1, 30)
sheet.hide_rows("3:5") / show_rows / hide_cols("C") / show_cols
sheet.insert_rows(before=3, count=2) / delete_rows(3, 1) / insert_cols("C", 1) / delete_cols("C", 1)
sheet.sort("A1:E40", by="C", descending=True, header=True)      (by may be a list of columns)
sheet.filter("A1:E40") (filter buttons) / sheet.filter("A1:E40", column="E", values=["Delivered"]) / clear_filter()
sheet.dropdown("D2:D50", ["Yes", "No"])
sheet.highlight("C2:C40", ">", 100, bg="#ffc7ce", fg="#9c0006")   rules: > < >= <= = != between contains begins
             ends top bottom above_average below_average formula(value="=$C2>$D2") blank
sheet.color_scale("C2:C40") / sheet.data_bar("D2:D40", color="#638ec6") / sheet.clear_rules()
sheet.comment("A1", "text") / sheet.link("A1", "https://...") / sheet.tab_color("#ff0000")
book.add_sheet("Summary") -> new sheet;  book.rename_sheet("Old", "New")
sheet.last_row, sheet.last_col, sheet.rows(), sheet.name
selection.row, selection.col (1-based first cell), selection.rows(), selection.address
book["Name"], book.sheets, book.active;  ref(row, col) -> "C2";  print(...) shows text to the user
sheet.clear() (empty the used cells; clear("A2:F99"); formats=True also removes formatting)
attachments["file name"]   files the user attached in this chat: a CSV / spreadsheet is a list of rows (first row
                 = header, values are text), a text file is its text. To put attached data in the sheet, write it from
                 here (sheet["A1"] = attachments["x.csv"]) - never retype the data in the script.

Rules: prefer real formulas (=SUM, =AVERAGE, =IF, =XLOOKUP...) when the result should stay live. Only use the
names above plus plain Python (math, statistics, re, datetime). Never use files, the network, os, sys or
subprocess. Keep scripts short. If no cells should change, don't write a script. Everything listed above can be
done by the script - never tell the user to do one of those by hand (e.g. use sheet.freeze, not "View > Freeze")."""


# ---------- the API key, kept encrypted on Windows ----------

def _blob(data):
    import ctypes
    from ctypes import wintypes

    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]
    buf = ctypes.create_string_buffer(data, len(data))
    return DATA_BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char))), buf, DATA_BLOB


def _dpapi(data, protect):
    import ctypes
    blob_in, _keep, DATA_BLOB = _blob(data)
    blob_out = DATA_BLOB()
    crypt = ctypes.windll.crypt32
    fn = crypt.CryptProtectData if protect else crypt.CryptUnprotectData
    ok = fn(ctypes.byref(blob_in), None, None, None, None, 0x1, ctypes.byref(blob_out))   # UI forbidden
    if not ok:
        raise OSError("Windows couldn't encrypt / decrypt the key")
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(blob_out.pbData)


def store_key(key):
    key = (key or "").strip()
    if not key:
        settings.set("ai_key", "")
        return
    if sys.platform == "win32":
        try:
            settings.set("ai_key", "dpapi:" + base64.b64encode(_dpapi(key.encode(), True)).decode())
            return
        except OSError:
            pass
    settings.set("ai_key", "plain:" + key)


def load_key():
    v = settings.get("ai_key", "") or ""
    if v.startswith("dpapi:"):
        try:
            return _dpapi(base64.b64decode(v[6:]), False).decode()
        except (OSError, ValueError):
            return ""
    return v[6:] if v.startswith("plain:") else v


# ---------- talking to OpenRouter ----------

class AIError(Exception):
    pass


def http(method, url, key=None, body=None, timeout=90):
    """Default transport: -> parsed JSON. Tests pass their own."""
    headers = {"Content-Type": "application/json", "HTTP-Referer": "https://github.com/mudiadamz/simple-sheets",
               "X-Title": "Simple Sheets", "User-Agent": f"SimpleSheets/{APP_VERSION}"}
    if key:
        headers["Authorization"] = f"Bearer {key}"
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            detail = json.loads(e.read().decode("utf-8")).get("error", {}).get("message", "")
        except Exception:
            detail = ""
        hint = {401: "The API key was refused - check it in the assistant's ⚙ settings.",
                402: "Your OpenRouter account has no credits left for this model.",
                404: "That model wasn't found - pick another one.",
                429: "Too many requests or the model is busy - wait a moment and try again."}.get(e.code, "")
        raise AIError(f"OpenRouter said {e.code}. {hint} {detail}".strip()) from None
    except (urllib.error.URLError, TimeoutError, OSError) as e:
        raise AIError(f"Couldn't reach OpenRouter: {getattr(e, 'reason', e)}") from None


def models(transport=http):
    """[(id, name, price per million tokens text)] of OpenRouter's models, cheapest first within a name."""
    data = transport("GET", f"{API}/models")
    out = []
    for m in data.get("data", []):
        p = m.get("pricing") or {}
        try:
            cost = float(p.get("prompt", 0)) * 1e6, float(p.get("completion", 0)) * 1e6
            price = ("free" if cost == (0, 0) else "price of the model it picks" if min(cost) < 0
                     else f"${cost[0]:.2g} / ${cost[1]:.2g} per M tokens")
        except (TypeError, ValueError):
            price = ""
        reads = ", ".join((m.get("architecture") or {}).get("input_modalities") or ["text"])
        out.append((m.get("id", ""), m.get("name", m.get("id", "")), price, reads))
    return sorted((x for x in out if x[0]), key=lambda x: x[0])


def ask(messages, model, key, transport=http):
    """-> the model's reply text."""
    if not key:
        raise AIError("Add your OpenRouter API key first (⚙ settings).")
    if not model:
        raise AIError("Choose a model first (⚙ settings).")
    data = transport("POST", f"{API}/chat/completions", key, {"model": model, "messages": messages})
    if data.get("error"):
        raise AIError(f"OpenRouter: {data['error'].get('message', data['error'])}")
    try:
        return data["choices"][0]["message"]["content"] or ""
    except (KeyError, IndexError, TypeError):
        raise AIError("The model sent an empty or unexpected reply.") from None


# ---------- what the model sees ----------

def _cell_text(g, r, c):
    t = g.text(r, c)
    if t.startswith("=") and len(t) > 1:
        v = g.value(r, c)
        return f"{t} → {v}"[:MAX_TEXT]
    return t[:MAX_TEXT]


def context(app, scope="sheet"):
    """The sheet as text for the model: names, selection and a table of cells with their addresses."""
    g = app.table
    s = g.sheet
    r1, c1, r2, c2 = g.sel_cells()
    sel = f"{col_letters(c1)}{r1 + 1}" + ("" if (r1, c1) == (r2, c2) else f":{col_letters(c2)}{r2 + 1}")
    if scope == "selection":
        a1, b1, a2, b2 = r1, c1, r2, c2
    else:
        a1, b1 = 0, 0
        a2 = max((r for r, row in enumerate(s.rows) if any(t != "" for t in row)), default=-1)
        b2 = max((max((c for c, t in enumerate(row) if t != ""), default=-1) for row in s.rows), default=-1)
    lines = [f"Workbook sheets: {', '.join(app.sheets)}", f"Active sheet: {app.sheet}", f"Selection: {sel}"]
    if a2 < a1 or b2 < b1:
        lines.append("The sheet is empty." if scope == "sheet" else "The selection is empty.")
        return "\n".join(lines)
    width = b2 - b1 + 1
    max_rows = max(MAX_CELLS // max(width, 1), 1)
    shown_to = min(a2, a1 + max_rows - 1)
    lines.append(f"Cells {col_letters(b1)}{a1 + 1}:{col_letters(b2)}{a2 + 1} (row number, then column: value; "
                 "formulas shown as =formula → result; empty cells left out):")
    for r in range(a1, shown_to + 1):
        cells = [f"{col_letters(c)}{r + 1}: {v}" for c in range(b1, b2 + 1) if (v := _cell_text(g, r, c)) != ""]
        if cells:
            lines.append(" | ".join(cells))
    if shown_to < a2:
        lines.append(f"… {a2 - shown_to} more rows not shown (up to row {a2 + 1}).")
    return "\n".join(lines)


IMAGE_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif",
               ".webp": "image/webp"}
SHEET_TYPES = (".xlsx", ".xlsm", ".xlsb", ".xls", ".ods", ".csv", ".tsv", ".xltx", ".xltm")
MAX_FILE = 8 * 1024 * 1024          # bytes of a picture / PDF sent
MAX_ATTACHED_TEXT = 60_000          # characters of a text file sent
PREVIEW_ROWS = 30                   # rows of an attached CSV / spreadsheet shown to the model (scripts get all)


class Attachment:
    """Something sent with the next question: a picture, a PDF, or text (a text file, a workbook's cells)."""

    def __init__(self, name, kind, data, mime="", table=None):
        self.name, self.kind, self.data, self.mime, self.table = name, kind, data, mime, table

    def part(self):
        """-> one OpenRouter message content part."""
        if self.kind == "image":
            return {"type": "image_url",
                    "image_url": {"url": f"data:{self.mime};base64,{base64.b64encode(self.data).decode()}"}}
        if self.kind == "pdf":
            return {"type": "file", "file": {"filename": self.name, "file_data":
                                             f"data:application/pdf;base64,{base64.b64encode(self.data).decode()}"}}
        if self.table is not None:                   # a CSV / spreadsheet: shape + first rows; scripts get it all
            n, m = len(self.table), max((len(r) for r in self.table), default=0)
            head = "\n".join(f"row {i + 1}: " + " | ".join(str(v)[:100] for v in r)
                              for i, r in enumerate(self.table[:PREVIEW_ROWS]))
            more = f"\n… {n - PREVIEW_ROWS} more rows" if n > PREVIEW_ROWS else ""
            return {"type": "text", "text": f"Attached file {self.name}: {n} rows x {m} columns. A script gets ALL of it "
                                            f"as attachments[{self.name!r}] (list of rows, first row = header, values "
                                            f"are text). First rows:\n{head}{more}"}
        return {"type": "text", "text": f"Attached file {self.name} (a script gets it as attachments[{self.name!r}]):"
                                        f"\n{self.data}"}


def workbook_rows(path):
    """A CSV / spreadsheet file -> its first sheet's rows (text values, trailing empty rows / cells dropped)."""
    import sheet_io as io
    book = io.load(path)
    sheet = next(iter(book.values()))
    rows = [list(r) for r in sheet.rows]
    while rows and not any(v != "" for v in rows[-1]):
        rows.pop()
    for r in rows:
        while r and r[-1] == "":
            r.pop()
    return rows


def attach_file(path):
    """A file -> Attachment. Raises AIError for files that can't be sent."""
    import os
    name = os.path.basename(path)
    ext = os.path.splitext(name)[1].lower()
    if ext in IMAGE_TYPES or ext == ".pdf":
        size = os.path.getsize(path)
        if size > MAX_FILE:
            raise AIError(f"{name} is {size / 1e6:.1f} MB: pictures and PDFs up to {MAX_FILE // 1024 // 1024} MB.")
        with open(path, "rb") as f:
            data = f.read()
        return Attachment(name, "pdf", data) if ext == ".pdf" else Attachment(name, "image", data, IMAGE_TYPES[ext])
    if ext in SHEET_TYPES:
        try:
            return Attachment(name, "text", "", table=workbook_rows(path))
        except Exception as e:
            raise AIError(f"Couldn't read {name}: {e}") from None
    with open(path, "rb") as f:
        raw = f.read(MAX_ATTACHED_TEXT * 4)
    if b"\0" in raw[:4096]:
        raise AIError(f"{name} isn't text, a picture, a PDF or a spreadsheet, so it can't be sent.")
    text = raw.decode("utf-8", errors="replace")
    if len(text) > MAX_ATTACHED_TEXT:
        text = text[:MAX_ATTACHED_TEXT] + "\n… (cut: the file is bigger)"
    return Attachment(name, "text", text)


def extract_code(reply):
    """The ```python block of a reply (or None), and the reply's text without it."""
    m = re.search(r"```(?:python|py)?\s*\n(.*?)```", reply, re.S)
    if not m:
        return None, reply.strip()
    return m.group(1).strip() + "\n", (reply[:m.start()] + reply[m.end():]).strip()


def chats_dir():
    import os

    from recent import config_dir
    d = os.path.join(config_dir(), "ai_chats")
    os.makedirs(d, exist_ok=True)
    return d


def save_chat_file(chat_id, data):
    import os
    path = os.path.join(chats_dir(), f"{chat_id}.json")
    try:
        with open(path + ".tmp", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        os.replace(path + ".tmp", path)
        return path
    except OSError:
        return None


def load_chat_file(chat_id):
    import os
    try:
        with open(os.path.join(chats_dir(), f"{chat_id}.json"), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def list_chats():
    """Saved chats, newest first: [{"id", "title", "updated"}]."""
    import glob
    import os
    out = []
    for p in glob.glob(os.path.join(chats_dir(), "*.json")):
        try:
            with open(p, encoding="utf-8") as f:
                d = json.load(f)
            out.append({"id": d.get("id") or os.path.basename(p)[:-5], "title": d.get("title") or "(untitled)",
                        "updated": d.get("updated", 0)})
        except (OSError, ValueError):
            continue
    return sorted(out, key=lambda d: d["updated"], reverse=True)


def clear_chats():
    import glob
    import os
    for p in glob.glob(os.path.join(chats_dir(), "*.json")):
        try:
            os.remove(p)
        except OSError:
            pass


class Files(dict):
    """attachments[...] in a script: a clear message when a name is wrong."""

    def __missing__(self, name):
        import scripting
        have = ", ".join(repr(k) for k in self) or "none in this chat"
        raise scripting.ScriptError(f"No attached file {name!r} (attached: {have})")


class Conversation:
    def __init__(self):
        self.messages = []

    def build(self, app, question, scope, attachments=()):
        msgs = [{"role": "system", "content": SYSTEM}]
        msgs += self.messages[-HISTORY:]
        text = f"{context(app, scope)}\n\nRequest: {question}"
        if attachments:                      # text first, then the pictures / files / attached text
            msgs.append({"role": "user", "content": [{"type": "text", "text": text}] + [a.part() for a in attachments]})
        else:
            msgs.append({"role": "user", "content": text})
        return msgs

    def add(self, question, reply, attachments=()):
        if attachments:                      # remembered by name only (not sent again with every question)
            question += f"  [attached: {', '.join(a.name for a in attachments)}]"
        self.messages += [{"role": "user", "content": question}, {"role": "assistant", "content": reply}]

    def note(self, text):
        """What happened when the script ran (so a follow-up question can build on it)."""
        self.messages.append({"role": "user", "content": f"(The script was run. Result: {text})"})
        self.messages.append({"role": "assistant", "content": "Noted."})


# ---------- the pane (docked on the right of the window) ----------

INLINE = re.compile(r"(\*\*[^*\n]+\*\*|`[^`\n]+`)")


def markdown_lines(text):
    """Reply text -> [(line, kind)] with kind 'h' (heading), 'bullet', 'code' or 'p'; fences dropped."""
    out, in_code = [], False
    for line in text.strip().split("\n"):
        if line.strip().startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            out.append((line, "code"))
        elif re.match(r"^\s*#{1,6}\s", line):
            out.append((re.sub(r"^\s*#{1,6}\s+", "", line), "h"))
        elif re.match(r"^\s*[-*•]\s+", line):
            indent = len(line) - len(line.lstrip())
            out.append(("  " * (indent // 2) + "• " + re.sub(r"^\s*[-*•]\s+", "", line), "bullet"))
        else:
            out.append((line, "p"))
    return out


class Assistant:
    """The AI assistant as a pane on the right of the main window (Ctrl+Shift+A shows / hides it)."""

    def __init__(self, app, transport=http):
        import tkinter as tk
        from tkinter import ttk

        from theme import C
        self.app, self.transport, self.conv = app, transport, Conversation()
        self.q = queue.Queue()
        self.busy = False
        old = getattr(app, "ai_pane", None)                  # one pane per window
        if old is not None and old.win.winfo_exists():
            old.win.destroy()
        app.ai_pane = self
        self.width = int(settings.get("ai_width", 0) or 380 * app.scale)
        outer = self.win = tk.Frame(app, width=self.width, highlightthickness=0)
        outer.pack_propagate(False)
        self.grip = tk.Frame(outer, width=5, cursor="sb_h_double_arrow")       # drag to resize the pane
        self.grip.pack(side="left", fill="y")
        self.grip.bind("<B1-Motion>", self.drag_width)
        self.grip.bind("<ButtonRelease-1>", lambda e: settings.set("ai_width", self.width))
        body = self.body = ttk.Frame(outer, padding=(8, 6, 8, 8))
        body.pack(side="left", fill="both", expand=True)

        # one slim row: ✦ model ▾ (click: change it)  …  ✎ new chat · ↺ history · ⚙ · ✕
        head = ttk.Frame(body)
        head.pack(fill="x", pady=(0, 6))
        from app_menu import sparkle_icon
        from format_bar import Tooltip
        self.mark = sparkle_icon(outer, app.scale, C["accent"])
        ttk.Label(head, image=self.mark).pack(side="left", padx=(0, 4))
        self.model_label = ttk.Label(head, text="", cursor="hand2")
        self.model_label.pack(side="left")
        self.model_label.bind("<Button-1>", lambda e: self.settings())   # click the model to change it
        Tooltip(self.model_label, "AI model - click to change it, or your OpenRouter key")
        for text, cmd, tip in (("✕", self.hide, "Hide (Ctrl+Shift+A)"), ("⚙", self.settings, "AI settings"),
                               ("↺", self.history_menu, "Earlier chats"), ("✎", self.new_chat, "New chat")):
            b = ttk.Button(head, text=text, width=2, style="Toolbutton", command=cmd)
            b.pack(side="right")
            Tooltip(b, tip)
            if text == "↺":
                self.history_btn = b
        self.chat_id, self.transcript = None, []
        self.files = {}                       # attached CSV / spreadsheets / text, for scripts (this chat)
        self.request = 0                      # which question is being answered (Stop ignores older ones)

        self.chat = tk.Text(body, wrap="word", state="disabled", relief="flat", padx=10, pady=8, cursor="arrow",
                            font=("Segoe UI", 10), spacing1=1, spacing3=1, highlightthickness=0, borderwidth=0)
        self.chat._own_colors = True

        # a proposed script: a small card (what it does, ▶ Run, the code only when asked for)
        self.script_box = tk.Frame(body, highlightthickness=1, padx=10, pady=8)
        self.script_box._own_colors = True
        top = tk.Frame(self.script_box)
        top._own_colors = True
        top.pack(fill="x")
        self.card_title = tk.Label(top, text="Script ready", font=("Segoe UI", 9, "bold"), anchor="w")
        self.card_title._own_colors = True
        self.card_title.pack(side="left")
        self.run_btn = ttk.Button(top, text="▶ Run", command=self.run, state="disabled")
        self.run_btn.pack(side="right")
        self.card_desc = tk.Label(self.script_box, text="", anchor="w", justify="left", font=("Segoe UI", 9))
        self.card_desc._own_colors = True
        self.card_desc.pack(fill="x", pady=(2, 4))
        links = tk.Frame(self.script_box)
        links._own_colors = True
        links.pack(fill="x")
        self.card_links = []
        for text, cmd in (("Code ▾", self.toggle_code), ("Edit", self.to_script_window),
                          ("Discard", lambda: self.set_code(None))):
            lab = tk.Label(links, text=text, cursor="hand2", font=("Segoe UI", 9, "underline"))
            lab._own_colors = True
            lab.pack(side="left", padx=(0, 12))
            lab.bind("<Button-1>", lambda e, c=cmd: c())
            self.card_links.append(lab)
        self.code_toggle = self.card_links[0]
        self.code = tk.Text(self.script_box, wrap="none", height=8, font=("Consolas", 9), undo=True, relief="flat",
                            padx=6, pady=4, highlightthickness=0)
        self.code._own_colors = True
        self.script_box.bind("<Configure>", lambda e: self.card_desc.configure(wraplength=max(e.width - 30, 100)))

        # the composer: one rounded box with + (attach / capture / what to send), the question and a send button
        self.ask_box = ttk.Frame(body)
        self.attachments = []
        self.chips = ttk.Frame(self.ask_box)              # attached files, each with ✕ (shown when any)
        self.composer = tk.Canvas(self.ask_box, height=44, highlightthickness=0, borderwidth=0)
        self.composer._own_colors = True
        self.composer.pack(fill="x")
        self.plus = tk.Label(self.composer, text="+", font=("Segoe UI", 16), cursor="hand2", padx=2, pady=0)
        self.plus._own_colors = True
        self.prompt = tk.Text(self.composer, height=1, wrap="word", font=("Segoe UI", 10), relief="flat",
                              borderwidth=0, highlightthickness=0, padx=2, pady=max(int(4 * app.scale), 3))
        self.prompt._own_colors = True
        self.send_btn = tk.Canvas(self.composer, width=30, height=30, highlightthickness=0, borderwidth=0,
                                  cursor="hand2")
        self.send_btn._own_colors = True
        self.mode_chip = tk.Label(self.composer, text="Ask", cursor="hand2", padx=4)
        self.mode_chip._own_colors = True
        self.mode_chip.bind("<Button-1>", lambda e: self.toggle_mode())
        self.sending = False
        self.scope = tk.StringVar(master=outer, value=settings.get("ai_scope", "sheet"))
        self.plus_menu = tk.Menu(self.plus, tearoff=False)
        self.plus_menu.add_command(label="Attach a file…", command=self.attach)
        self.plus_menu.add_command(label="Capture part of the screen", command=self.capture)
        self.plus_menu.add_command(label="Paste a picture (Ctrl+V)", command=lambda: self.paste() or None)
        self.plus_menu.add_separator()
        self.plus_menu.add_radiobutton(label="Send the sheet", value="sheet", variable=self.scope)
        self.plus_menu.add_radiobutton(label="Send only the selection", value="selection", variable=self.scope)
        self.plus.bind("<Button-1>", lambda e: self.plus_menu.tk_popup(e.x_root, e.y_root))
        self.send_btn.bind("<Button-1>", lambda e: self.stop() if self.busy else self.send())
        from format_bar import Tooltip
        Tooltip(self.plus, "Attach a file, capture the screen, or choose what to send (sheet / selection)")
        Tooltip(self.mode_chip, "Ask: you press Run on the AI's scripts.  Auto: they run straight away "
                                "(still one Ctrl+Z).  Click to switch.")
        Tooltip(self.send_btn, "Ask (Enter). Your question and the sent cells go to OpenRouter and the model's "
                               "provider; scripts only run when you press Run.")
        self.composer.bind("<Configure>", lambda e: self.layout_composer())
        self.prompt.bind("<KeyRelease>", lambda e: self.layout_composer(), add="+")
        self.ask_box.pack(side="bottom", fill="x", pady=(8, 0))   # bottom first, then the chat takes the rest
        self.chat.pack(side="top", fill="both", expand=True)

        self.prompt.bind("<Return>", lambda e: (self.send(), "break")[-1])               # Enter sends
        self.prompt.bind("<Control-Return>", lambda e: (self.send(), "break")[-1])
        self.prompt.bind("<Shift-Return>", lambda e: None)                                # new line
        self.prompt.bind("<Control-v>", self.paste)            # a copied picture / screenshot becomes an attachment
        self.prompt.bind("<Control-V>", self.paste)
        self.prompt.bind("<FocusIn>", lambda e: self.hint(False))
        self.prompt.bind("<FocusOut>", lambda e: self.hint(True))
        self.recolor()
        self.paint_mode()
        self.hint(True)
        self.refresh_label()
        self.show()
        if not load_key() or not settings.get("ai_model"):
            app.after(200, lambda: outer.winfo_exists() and self.settings())

    # --- look ---

    def recolor(self):
        """Colours of the chat, the script and the question box (also after a theme switch)."""
        from theme import C
        self.chat.configure(background=C["chat_bg"], foreground=C["fg"], selectbackground=C["select_bg"])
        self.code.configure(background=C["field_bg"], foreground=C["fg"], insertbackground=C["fg"],
                            highlightbackground=C["header_line"], highlightcolor=C["accent"])
        self.prompt.configure(background=C["field_bg"], foreground=C["fg"], insertbackground=C["fg"])
        self.plus.configure(background=C["field_bg"], foreground=C["fg"])
        card = C["chat_bg"]
        self.script_box.configure(background=card, highlightbackground=C["header_line"])
        for w in [self.card_title.master, self.card_links[0].master]:
            w.configure(background=card)
        self.card_title.configure(background=card, foreground=C["fg"])
        self.card_desc.configure(background=card, foreground=C["muted"])
        for lab in self.card_links:
            lab.configure(background=card, foreground=C["accent"])
        self.code.configure(background=C["field_bg"], foreground=C["fg"], insertbackground=C["fg"])
        if hasattr(self, "mode_chip"):
            self.mode_chip.configure(background=C["field_bg"])
        self.composer.configure(background=C["bg"])
        self.send_btn.configure(background=C["field_bg"])
        self.layout_composer()
        self.grip.configure(background=C["header_line"])
        self.win.configure(background=C["bg"])
        t = self.chat
        t.tag_configure("you_name", foreground=C["accent"], font=("Segoe UI", 9, "bold"), lmargin1=48, spacing1=8)
        t.tag_configure("you", background=C["chat_user"], lmargin1=48, lmargin2=48, rmargin=6, spacing1=4,
                        spacing3=4)
        t.tag_configure("ai_name", foreground=C["muted"], font=("Segoe UI", 9, "bold"), spacing1=8)
        t.tag_configure("h", font=("Segoe UI", 11, "bold"), spacing1=6)
        t.tag_configure("bold", font=("Segoe UI", 10, "bold"))
        t.tag_configure("code", font=("Consolas", 9), background=C["chat_code"])
        t.tag_configure("bullet", lmargin1=6, lmargin2=20)
        t.tag_configure("note", foreground=C["muted"], font=("Segoe UI", 9, "italic"), spacing1=4)
        t.tag_configure("working", foreground=C["accent"], font=("Segoe UI", 9), spacing1=6)
        self.prompt.tag_configure("hint", foreground=C["muted"])

    def layout_composer(self):
        """The rounded box (grows to 4 lines with the question) with its +, text and send button."""
        from tkinter import font as tkfont

        from theme import C
        cv = self.composer
        k = self.app.scale                                           # everything follows the display scaling
        w = max(cv.winfo_width(), int(200 * k))
        btn = self.send_size()
        plus_x = int(10 * k) + self.plus.winfo_reqwidth() // 2       # inside the rounded end, never on its edge
        text_x = plus_x + self.plus.winfo_reqwidth() // 2 + int(6 * k)
        send_x = w - int(8 * k) - btn // 2
        chip_w = self.mode_chip.winfo_reqwidth()
        chip_x = send_x - btn // 2 - int(6 * k) - chip_w // 2
        text_w = max(chip_x - chip_w // 2 - int(6 * k) - text_x, 60)
        font = tkfont.Font(root=cv, font=self.prompt.cget("font"))
        line_h = font.metrics("linespace")                           # the real font, at this display scaling
        # wrapped lines of the question at the width it gets (measured, so it's right before it is drawn)
        avail = max(text_w - 2 * int(self.prompt.cget("padx")) - 4, 20)
        lines = 0
        for para in self.prompt.get("1.0", "end-1c").split("\n"):
            used, n = 0, 1
            for word in para.split(" "):
                wd = font.measure(word + " ")
                if used and used + wd > avail:
                    n, used = n + 1, 0
                used += wd
                while used > avail:                                  # a word longer than the line
                    n, used = n + 1, used - avail
            lines += n
        lines = max(1, min(lines, 4))
        text_h = lines * line_h + 2 * int(self.prompt.cget("pady")) + 2
        h = max(text_h + int(14 * k), btn + int(14 * k), self.plus.winfo_reqheight() + int(8 * k))
        if int(cv.cget("height")) != h:
            cv.configure(height=h)
        cv.delete("box")
        r = min(h // 2, int(20 * k)) if lines == 1 else int(14 * k)
        x1, y1, x2, y2 = 1, 1, w - 2, h - 2
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2, x1, y2,
               x1, y2 - r, x1, y1 + r, x1, y1]
        cv.create_polygon(pts, smooth=True, fill=C["field_bg"], outline=C["header_line"], width=1, tags="box")
        cv.tag_lower("box")
        mid = h // 2
        if not cv.find_withtag("plus"):
            cv.create_window(plus_x, mid, window=self.plus, tags="plus")
            cv.create_window(text_x, mid, window=self.prompt, anchor="w", tags="prompt")
            cv.create_window(send_x, mid, window=self.send_btn, tags="send")
            cv.create_window(chip_x, mid, window=self.mode_chip, tags="chip")
        cv.coords("plus", plus_x, mid)
        cv.coords("prompt", text_x, mid)
        cv.itemconfigure("prompt", width=text_w, height=text_h)
        cv.coords("send", send_x, mid)
        cv.coords("chip", chip_x, mid)
        self.draw_send()

    def send_size(self):
        return max(int(30 * self.app.scale), 24)

    def draw_send(self):
        """Round accent button with an up arrow (grey while waiting for the answer)."""
        from theme import C, shade
        b = self.send_btn
        s = self.send_size()
        if int(b.cget("width")) != s:
            b.configure(width=s, height=s)
        b.delete("all")
        fill = C["accent"]
        b.create_oval(1, 1, s - 1, s - 1, fill=fill, outline=shade(fill, 0.15) if not C.get("contrast") else fill)
        ink, m, k = C["on_accent"], s / 2, s / 30
        if self.sending:                         # working: a stop square
            b.create_rectangle(m - 5 * k, m - 5 * k, m + 5 * k, m + 5 * k, fill=ink, outline=ink)
            return
        b.create_line(m, m + 6 * k, m, m - 6 * k, fill=ink, width=max(2, round(2 * k)), capstyle="round")
        b.create_line(m - 5 * k, m - 1 * k, m, m - 6 * k, m + 5 * k, m - 1 * k, fill=ink, width=max(2, round(2 * k)),
                      capstyle="round", joinstyle="round")

    def set_sending(self, on):
        self.sending = bool(on)
        self.draw_send()

    def hint(self, show):
        """Grey 'Ask about your data…' in the empty question box."""
        text = self.prompt.get("1.0", "end-1c")
        if show and not text.strip():
            self.prompt.delete("1.0", "end")
            self.prompt.insert("1.0", "Ask about your data, or say what to change…", "hint")
            self.prompt._hinting = True
        elif not show and getattr(self.prompt, "_hinting", False):
            self.prompt.delete("1.0", "end")
            self.prompt._hinting = False

    def question_text(self):
        return "" if getattr(self.prompt, "_hinting", False) else self.prompt.get("1.0", "end").strip()

    def drag_width(self, e):
        self.width = max(int(260 * self.app.scale), min(self.width - e.x, int(self.app.winfo_width() * 0.7)))
        self.win.configure(width=self.width)

    def show(self):
        if not self.win.winfo_ismapped():
            self.win.pack(side="right", fill="y", before=self.app.table)
        self.prompt.focus_set()
        self.hint(False)
        __import__("app_menu").refresh_ai_button(self.app)

    def hide(self):
        self.win.pack_forget()
        self.app.table.cv.focus_set()
        __import__("app_menu").refresh_ai_button(self.app)

    def visible(self):
        return bool(self.win.winfo_manager())

    def refresh_label(self):
        """The model's short name (the part after the provider) and ▾; the full id is in settings."""
        model = settings.get("ai_model") or ""
        short = model.split("/")[-1] if model else "choose a model"
        if len(short) > 28:
            short = short[:27] + "…"
        self.model_label.configure(text=f"{short} ▾")

    # --- chat ---

    def add_text(self, text, base=()):
        """Text with **bold** and `code` inline, under extra tags `base`."""
        for part in INLINE.split(text):
            if not part:
                continue
            if part.startswith("**") and part.endswith("**"):
                self.chat.insert("end", part[2:-2], base + ("bold",))
            elif part.startswith("`") and part.endswith("`"):
                self.chat.insert("end", part[1:-1], base + ("code",))
            else:
                self.chat.insert("end", part, base)

    def say(self, who, text, record=True):
        """who: 'You', a model name, or None for a grey note."""
        if record:
            self.transcript.append([who, text])
        c = self.chat
        c.configure(state="normal")
        c.delete("thinking.first", "thinking.last") if c.tag_ranges("thinking") else None
        if who == "You":
            c.insert("end", "You\n", "you_name")
            c.insert("end", text.strip() + "\n", "you")
        elif who:
            c.insert("end", who + "\n", "ai_name")
            for line, kind in markdown_lines(text):
                if kind == "code":
                    c.insert("end", line + "\n", ("code",))
                else:
                    self.add_text(line, {"h": ("h",), "bullet": ("bullet",)}.get(kind, ()))
                    c.insert("end", "\n", {"h": ("h",), "bullet": ("bullet",)}.get(kind, ()))
        else:
            c.insert("end", text.strip() + "\n", "note")
        c.see("end")
        c.configure(state="disabled")

    def thinking(self, what=""):
        """A live 'working' line: animated dots, what it is doing, seconds so far (until the answer comes)."""
        import time
        self.think_started, self.think_what, self.think_step = time.monotonic(), what, 0
        self.paint_thinking()
        self.app.after(400, self.tick_thinking, self.request)

    def paint_thinking(self):
        import time
        c = self.chat
        secs = int(time.monotonic() - getattr(self, "think_started", time.monotonic()))
        dots = "●" * (self.think_step % 3 + 1) + "○" * (2 - self.think_step % 3)
        text = f"{dots}  {self.think_what or 'Thinking'}…" + (f"  {secs} s" if secs >= 2 else "") + \
            ("   (■ to stop)" if secs >= 5 else "")
        c.configure(state="normal")
        if c.tag_ranges("thinking"):
            c.delete("thinking.first", "thinking.last")
        c.insert("end", text + "\n", ("working", "thinking"))
        c.see("end")
        c.configure(state="disabled")

    def tick_thinking(self, request):
        if not self.busy or request != self.request or not self.win.winfo_exists():
            return
        self.think_step += 1
        self.paint_thinking()
        self.app.after(400, self.tick_thinking, request)

    def stop(self):
        """■: stop waiting; the answer to this question is ignored when it comes."""
        if not self.busy:
            return False
        self.request += 1
        self.busy = False
        self.set_sending(False)
        self.say(None, "Stopped.")
        return True

    def new_chat(self):
        """Start over. The chat so far stays in the history (↺)."""
        self.save_chat()
        self.conv = Conversation()
        self.chat_id, self.transcript, self.files = None, [], {}
        self.chat.configure(state="normal")
        self.chat.delete("1.0", "end")
        self.chat.configure(state="disabled")
        self.set_code(None)

    # --- history: chats kept on this computer (settings folder / ai_chats), never sent anywhere ---

    def save_chat(self):
        if not any(who == "You" for who, _ in self.transcript):
            return None
        import time
        if self.chat_id is None:
            self.chat_id = time.strftime("%Y%m%d-%H%M%S") + f"-{int(time.time() * 1000) % 1000:03d}"
        title = next(text for who, text in self.transcript if who == "You").split("\n")[0][:70]
        return save_chat_file(self.chat_id, {"id": self.chat_id, "title": title, "updated": time.time(),
                                             "model": settings.get("ai_model") or "", "messages": self.conv.messages,
                                             "shown": self.transcript})

    def open_chat(self, chat_id):
        data = load_chat_file(chat_id)
        if not data:
            return False
        self.save_chat()
        self.new_chat_view()
        self.chat_id = chat_id
        self.conv.messages = list(data.get("messages") or [])
        self.transcript = [list(x) for x in data.get("shown") or []]
        for who, text in self.transcript:
            self.say(who, text, record=False)
        return True

    def new_chat_view(self):
        self.conv = Conversation()
        self.chat.configure(state="normal")
        self.chat.delete("1.0", "end")
        self.chat.configure(state="disabled")
        self.set_code(None)

    def history_menu(self, x=None, y=None):
        """↺: earlier chats (click one to continue it) and Clear history."""
        m = self.build_history_menu()
        if x is None:
            b = getattr(self, "history_btn", self.win)
            x, y = b.winfo_rootx(), b.winfo_rooty() + b.winfo_height()
        try:
            m.tk_popup(x, y)
        finally:
            m.grab_release()
        return m

    def build_history_menu(self):
        import tkinter as tk
        m = tk.Menu(self.win, tearoff=False)
        chats = list_chats()[:20]
        if not chats:
            m.add_command(label="(no earlier chats yet)", state="disabled")
        for c in chats:
            when = __import__("time").strftime("%d %b %H:%M", __import__("time").localtime(c.get("updated", 0)))
            mark = "• " if c["id"] == self.chat_id else ""
            m.add_command(label=f"{mark}{c['title']}   ({when})", command=lambda i=c["id"]: self.open_chat(i))
        m.add_separator()
        m.add_command(label="Clear history…", command=self.clear_history, state="normal" if chats else "disabled")
        return m

    def clear_history(self):
        from tkinter import messagebox
        if not messagebox.askyesno("AI chats", "Delete every saved AI chat on this computer?",
                                   parent=self.win.winfo_toplevel()):
            return False
        clear_chats()
        self.chat_id = None
        return True

    def set_code(self, code):
        """Show the card for a proposed script (code hidden until 'Code ▾'), or hide it (None)."""
        self.code.delete("1.0", "end")
        if code:
            self.code.insert("1.0", code)
            lines = [x for x in code.strip().split("\n") if x.strip()]
            what = next((x.strip().lstrip("#").strip() for x in lines if x.strip().startswith("#")), "")
            self.card_title.configure(text=f"Script ready · {len(lines)} line{'s' if len(lines) != 1 else ''}")
            self.card_desc.configure(text=what or "Check it with Code ▾, then press ▶ Run.")
            if not self.script_box.winfo_manager():
                self.script_box.pack(side="bottom", fill="x", pady=(8, 0), before=self.chat)   # above the question
        else:
            self.script_box.pack_forget()
        self.show_code(False)
        self.run_btn.configure(state="normal" if code else "disabled")

    def show_code(self, on):
        if on and not self.code.winfo_manager():
            self.code.pack(fill="x", pady=(6, 0))
        elif not on:
            self.code.pack_forget()
        self.code_toggle.configure(text="Code ▴" if on else "Code ▾")

    def toggle_code(self):
        self.show_code(not self.code.winfo_manager())

    # --- Ask / Auto: run the AI's scripts after checking them, or straight away ---

    def auto(self):
        return bool(settings.get("ai_auto_run", False))

    def toggle_mode(self, confirm=True):
        from tkinter import messagebox
        if not self.auto() and confirm and not messagebox.askyesno(
                "Run AI scripts automatically?",
                "In Auto mode the AI's scripts run as soon as they arrive, without you reading them first.\n\n"
                "Each run is still one undo step (Ctrl+Z) and a failing script changes nothing - but a script can "
                "do anything Python can on this computer.\n\nSwitch to Auto?", parent=self.win.winfo_toplevel()):
            return self.auto()
        settings.set("ai_auto_run", not self.auto())
        self.paint_mode()
        return self.auto()

    def paint_mode(self):
        from theme import C
        on = self.auto()
        self.mode_chip.configure(text="Auto" if on else "Ask", foreground=C["accent"] if on else C["muted"],
                                 background=C["field_bg"], font=("Segoe UI", 9, "bold" if on else "normal"))
        self.layout_composer()

    def send(self, question=None):
        if self.busy:
            return None
        question = (question or self.question_text()).strip()
        if not question:
            return None
        settings.set("ai_scope", self.scope.get())
        attached = list(self.attachments)
        for a in attached:                    # attached data stays usable by scripts for the rest of the chat
            if a.table is not None:
                self.files[a.name] = a.table
            elif a.kind == "text":
                self.files[a.name] = a.data
        msgs = self.conv.build(self.app, question, self.scope.get(), attached)
        key, model = load_key(), settings.get("ai_model")
        self.prompt.delete("1.0", "end")
        self.prompt._hinting = False
        self.say("You", question + (f"\n📎 {', '.join(a.name for a in attached)}" if attached else ""))
        self.clear_attachments()
        self.request += 1
        request = self.request
        self.busy = True
        model_name = (model or "the model").split("/")[-1]
        self.thinking(f"Reading {attached[0].name}" if attached else f"Asking {model_name}")
        self.set_sending(True)

        def work():
            try:
                self.q.put(("ok", (question, attached, request), ask(msgs, model, key, self.transport)))
            except AIError as e:
                hint = (" This model may not read pictures or PDFs: pick one whose list entry says 'image' / 'file' "
                        "(⚙).") if any(a.kind != "text" for a in attached) else ""
                self.q.put(("error", (question, attached, request), str(e) + hint))
            except Exception as e:                       # never leave the pane waiting
                self.q.put(("error", (question, attached, request), f"{type(e).__name__}: {e}"))
        threading.Thread(target=work, daemon=True).start()
        self.app.after(100, self.poll)
        return msgs

    def poll(self):
        try:
            kind, question, text = self.q.get_nowait()
        except queue.Empty:
            if self.win.winfo_exists():
                self.app.after(100, self.poll)
            return
        if question[2] != self.request:      # stopped (or an older question): ignore this answer
            if self.busy and self.win.winfo_exists():
                self.app.after(100, self.poll)
            return
        self.busy = False
        if not self.win.winfo_exists():
            return
        self.set_sending(False)
        if kind == "error":
            self.say(None, "⚠ " + text)
            return
        question, attached, _request = question
        self.conv.add(question, text, attached)
        code, words = extract_code(text)
        self.say(settings.get("ai_model") or "AI", words or ("Here is a script for that." if code else "(no answer)"))
        self.set_code(code)
        if code and self.auto():
            self.run(auto=True)                       # Auto mode: straight away (one Ctrl+Z, rolled back on errors)
        elif code:
            self.say(None, "A script is ready below: check it, then press ▶ Run.")
        self.save_chat()

    # --- attachments ---

    def add_attachment(self, att):
        import tkinter as tk
        from tkinter import ttk
        self.attachments.append(att)
        chip = ttk.Frame(self.chips, padding=(4, 1))
        chip.pack(side="left", padx=(0, 4))
        icon = {"image": "🖼", "pdf": "📄"}.get(att.kind, "📝")
        ttk.Label(chip, text=f"{icon} {att.name}").pack(side="left")
        from theme import C
        x = tk.Label(chip, text="✕", cursor="hand2", background=C["bg"], foreground=C["muted"])
        x.pack(side="left", padx=(4, 0))
        x.bind("<Button-1>", lambda e, a=att, c=chip: self.remove_attachment(a, c))
        att.chip = chip
        if not self.chips.winfo_manager():
            self.chips.pack(fill="x", pady=(0, 4), before=self.composer)
        return att

    def remove_attachment(self, att, chip=None):
        if att in self.attachments:
            self.attachments.remove(att)
        (chip or att.chip).destroy()
        if not self.attachments:
            self.chips.pack_forget()

    def clear_attachments(self):
        for a in list(self.attachments):
            self.remove_attachment(a)

    def attach(self, path=None):
        from tkinter import filedialog
        path = path or filedialog.askopenfilename(
            parent=self.win.winfo_toplevel(), title="Attach to the question",
            filetypes=[("Pictures, PDF, text, spreadsheets", "*.png *.jpg *.jpeg *.gif *.webp *.pdf *.txt *.csv *.tsv "
                        "*.json *.md *.xml *.xlsx *.xlsm *.xlsb *.xls *.ods"), ("All files", "*.*")])
        if not path:
            return None
        try:
            return self.add_attachment(attach_file(path))
        except (AIError, OSError) as e:
            self.say(None, f"⚠ {e}")
            return None

    def capture(self, grab_fn=None):
        """Dim the screen and drag a box: that part is attached as a picture."""
        import screen_capture
        stamp = __import__("datetime").datetime.now().strftime("%H%M%S")

        def done(data):
            if data:
                self.add_attachment(Attachment(f"screen-{stamp}.png", "image", data, "image/png"))
            self.prompt.focus_set()
        try:
            return screen_capture.Snip(self.win.winfo_toplevel(), done, grab_fn)
        except OSError as e:
            self.say(None, f"⚠ {e}")
            return None

    def paste(self, e=None):
        """Ctrl+V: a picture on the clipboard is attached; text pastes as usual."""
        import screen_capture
        data = screen_capture.clipboard_image()
        if not data:
            return None
        stamp = __import__("datetime").datetime.now().strftime("%H%M%S")
        self.hint(False)
        self.add_attachment(Attachment(f"pasted-{stamp}.png", "image", data, "image/png"))
        return "break"

    def run(self, auto=False):
        import scripting
        code = self.code.get("1.0", "end-1c")
        if not code.strip():
            return None
        out = []
        ok, msg = scripting.run(self.app, code, out=out.append, filename="<AI script>",
                                extra={"attachments": Files(self.files)})
        undo = "  (Ctrl+Z undoes it)" if ok and "no cells changed" not in msg else ""
        self.say(None, ("".join(out) + ("✔ " if ok else "✖ ") + ("Ran automatically: " if auto else "") + msg
                        + undo).strip())
        if not ok and auto:                           # failed in Auto mode: keep the card to look at / fix
            self.show_code(True)
        self.conv.note(msg + (" Output: " + "".join(out)[:500] if out else ""))
        self.save_chat()
        if ok:
            self.set_code(None)
        return ok

    def to_script_window(self):
        import scripting
        w = scripting.open_window(self.app)
        w.load_text(self.code.get("1.0", "end-1c"))
        return w

    def settings(self):
        """The AI settings window - only ever one: asking again brings the open one to the front."""
        d = getattr(self, "settings_dialog", None)
        if d is not None and d.win.winfo_exists():
            d.win.deiconify()
            d.win.lift()
            d.win.focus_force()
            return d
        self.settings_dialog = SettingsDialog(self.win.winfo_toplevel(), self)
        return self.settings_dialog


class SettingsDialog:
    def __init__(self, parent, assistant=None):
        import tkinter as tk
        from tkinter import ttk
        self.assistant = assistant
        transport = assistant.transport if assistant else http
        self.transport = transport
        w = self.win = tk.Toplevel(parent)
        w.title("AI settings")
        w.transient(parent)
        w._root().after_idle(lambda: w.winfo_exists() and w.grab_set())   # modal: one settings window at a time
        f = ttk.Frame(w, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="OpenRouter API key").grid(row=0, column=0, sticky="w")
        self.key = tk.StringVar(master=w, value=load_key())
        ttk.Entry(f, textvariable=self.key, show="•", width=46).grid(row=0, column=1, sticky="w", pady=2)
        link = ttk.Label(f, text="Get a key at openrouter.ai/keys", foreground="#1a73e8", cursor="hand2")
        link.grid(row=1, column=1, sticky="w")
        link.bind("<Button-1>", lambda e: __import__("webbrowser").open("https://openrouter.ai/keys"))
        ttk.Label(f, text="Model").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.model = ttk.Combobox(f, width=44)
        self.model.set(settings.get("ai_model") or "")
        self.model.grid(row=2, column=1, sticky="w", pady=(10, 0))
        ttk.Label(f, text="Type to filter, or pick from the list (Load models).").grid(row=3, column=1, sticky="w")
        self.info = ttk.Label(f, text="", wraplength=360)
        self.info.grid(row=4, column=1, sticky="w")
        ttk.Label(f, text="Your questions, attachments and the sent cells go to OpenRouter and the model's provider. "
                          "AI scripts only run when you press Run.", wraplength=420, foreground="#777777").grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))
        b = ttk.Frame(f)
        b.grid(row=6, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(b, text="Auto model", command=self.auto_model).pack(side="left", padx=4)
        ttk.Button(b, text="Load models", command=self.load_models).pack(side="left", padx=4)
        ttk.Button(b, text="Save", command=self.save).pack(side="left", padx=4)
        ttk.Button(b, text="Cancel", command=w.destroy).pack(side="left")
        self.all = []
        self.model.bind("<KeyRelease>", lambda e: self.filter())
        self.model.bind("<<ComboboxSelected>>", lambda e: self.show_price())
        self.info.configure(text="Loading the model list…")
        w._root().after(150, lambda: w.winfo_exists() and not self.all and self.load_models())

    def load_models(self):
        try:
            self.all = models(self.transport)
        except AIError as e:
            self.info.configure(text=str(e))
            return []
        self.filter()
        self.info.configure(text=f"{len(self.all)} models. Pick one; prices are per million tokens (in / out).")
        return self.all

    def filter(self):
        text = self.model.get().lower()
        self.model["values"] = [m[0] for m in self.all if text in m[0].lower() or text in m[1].lower()][:300]

    def show_price(self):
        m = next((m for m in self.all if m[0] == self.model.get()), None)
        if m:
            self.info.configure(text=f"{m[1]} - {m[2]} - reads: {m[3]}")

    def auto_model(self):
        """OpenRouter's Auto Router: it picks a suitable model for each question."""
        self.model.set("openrouter/auto")
        self.info.configure(text="Auto: OpenRouter picks a suitable model for each question (you pay that model's "
                                 "price).")

    def save(self):
        store_key(self.key.get())
        settings.set("ai_model", self.model.get().strip())
        if self.assistant:
            self.assistant.refresh_label()
        self.win.destroy()
        return True


def open_assistant(app, transport=None):
    """Show the assistant pane (made the first time), or hide it when it is showing (Ctrl+Shift+A)."""
    app.table.commit()
    pane = getattr(app, "ai_pane", None)
    if pane is not None and pane.win.winfo_exists():
        if pane.visible():
            pane.hide()
        else:
            pane.show()
        return pane
    app.ai_pane = Assistant(app, transport or http)
    return app.ai_pane
