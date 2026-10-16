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
sheet.on_click("G2", code)   a button in the sheet: clicking G2 runs `code` (a Python string, same names as here,
                 plus `clicked` = "G2"); on_click("G2", None) removes it. Use it when the user wants a cell
                 that does something when clicked (e.g. "go to today": code sets the year / month cells).
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
WEB_NOTE = """
Web search is on: results from the web come with this question. Use them for facts the sheet doesn't have
(prices, figures, lists…), put the found values into the sheet with a script when asked, and mention where
they came from. Never invent figures you didn't find."""


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
        remember_context(m)
    return sorted((x for x in out if x[0]), key=lambda x: x[0])


# ---------- usage: tokens and cost counted on this computer; credits / spending from OpenRouter ----------

def record_usage(entries, today=None):
    """Add calls' {"in", "out", "cost"} to the per-day totals in settings (ai_usage: {"days": {date: [requests,
    in, out, cost]}, "older": [...] for days dropped after a year})."""
    import datetime as dt
    if not entries:
        return None
    day = today or dt.date.today().isoformat()
    u = settings.get("ai_usage") or {}
    days = dict(u.get("days") or {})
    row = list(days.get(day) or [0, 0, 0, 0.0])
    for e in entries:
        row = [row[0] + 1, row[1] + e["in"], row[2] + e["out"], round(row[3] + e["cost"], 6)]
    days[day] = row
    older = list(u.get("older") or [0, 0, 0, 0.0])
    for d in sorted(days)[:-400]:                   # keep about a year day by day
        older = [a + b for a, b in zip(older, days.pop(d))]
    settings.set("ai_usage", {"days": days, "older": older})
    return row


def usage_totals(today=None):
    """{"Today", "This month", "All time"} -> [requests, input tokens, output tokens, cost]."""
    import datetime as dt
    day = today or dt.date.today().isoformat()
    u = settings.get("ai_usage") or {}
    days = u.get("days") or {}

    def total(rows, start=None):
        return [sum(r[i] for r in rows) for i in range(4)] if rows else [0, 0, 0, 0.0]
    all_rows = list(days.values()) + [u.get("older") or [0, 0, 0, 0.0]]
    return {"Today": total([days[day]] if day in days else []),
            "This month": total([r for d, r in days.items() if d[:7] == day[:7]]),
            "All time": total(all_rows)}


def reset_usage():
    settings.set("ai_usage", {})


def account(key, transport=http):
    """What OpenRouter says about this key: {"credits", "used", "limit", "limit_remaining", "daily", "weekly",
    "monthly", "free_tier", "free_requests": (used, limit) or None}. Raises AIError."""
    if not key:
        raise AIError("Add your OpenRouter API key to see your credits and spending.")
    k = (transport("GET", f"{API}/key", key) or {}).get("data") or {}
    out = {"used": k.get("usage"), "limit": k.get("limit"), "limit_remaining": k.get("limit_remaining"),
           "daily": k.get("usage_daily"), "weekly": k.get("usage_weekly"), "monthly": k.get("usage_monthly"),
           "free_tier": bool(k.get("is_free_tier")), "credits": None, "account_used": None, "free_requests": None}
    fr = k.get("free_model_daily_requests")
    if isinstance(fr, dict) and fr.get("limit"):
        out["free_requests"] = (fr.get("used", 0), fr["limit"])
    try:                                           # the account's credits (may be refused for some keys)
        c = (transport("GET", f"{API}/credits", key) or {}).get("data") or {}
        if c.get("total_credits") is not None:
            out["credits"], out["account_used"] = float(c["total_credits"]), float(c.get("total_usage") or 0)
    except AIError:
        pass
    return out


def money(x):
    if x is None:
        return "—"
    return f"${x:,.2f}" if abs(x) >= 0.01 or x == 0 else f"${x:.4f}"


def count_text(n):
    return f"{n / 1e6:.1f}M" if n >= 1e6 else f"{n / 1e3:.1f}k" if n >= 1e4 else f"{n:,}"


# ---------- fitting a question into the model's context window ----------

DEFAULT_CONTEXT = 32_000      # tokens, when the model's size isn't known (yet)
AUTO_CONTEXT = 128_000        # Auto model: what most of the models it picks from hold
HISTORY_MAX = 40              # messages of the chat sent at most (older ones: the summary)
_context_sizes = {}           # model id -> context length in tokens (from OpenRouter's model list)


def remember_context(m):
    try:
        n = int(m.get("context_length") or (m.get("top_provider") or {}).get("context_length") or 0)
    except (TypeError, ValueError):
        n = 0
    if m.get("id") and n > 0:
        _context_sizes[m["id"]] = n


def context_size(model):
    """Tokens the model takes (question + answer): known from the model list / settings, else a safe guess."""
    if model == "openrouter/auto":                  # says 2M, but the model it picks may hold far less
        return AUTO_CONTEXT
    if model in _context_sizes:
        return _context_sizes[model]
    saved = settings.get("ai_model_ctx") or {}
    if saved.get("model") == model and saved.get("tokens"):
        return int(saved["tokens"])
    return DEFAULT_CONTEXT


def fetch_context_sizes(transport=http):
    """Fill the model sizes from OpenRouter's model list (once, on a thread; errors are ignored)."""
    try:
        for m in transport("GET", f"{API}/models").get("data", []):
            remember_context(m)
    except Exception:
        pass


def tokens(x):
    """A rough, safe-side token count of a message / text (about 3 characters per token; pictures, PDFs fixed)."""
    if isinstance(x, str):
        return len(x) // 3 + 4
    if isinstance(x, dict):
        if x.get("type") == "image_url":
            return 1500
        if x.get("type") == "file":
            return 4000
        return tokens(x.get("content") if "content" in x else x.get("text", ""))
    if isinstance(x, list):
        return sum(tokens(p) for p in x)
    return 0


def too_big(error):
    """Does an OpenRouter error say the question didn't fit the model?"""
    return bool(re.search(r"context.{0,20}(length|window|limit)|too many tokens|maximum.{0,30}tokens|too long|"
                          r"reduce the length", error, re.I))


def input_budget(ctx):
    """Tokens a question may use: the context minus room for the answer, with some slack for the estimate."""
    return int((ctx - min(4000, ctx // 4)) * 0.9)


def cut_sheet(text, room):
    """The sheet text cut to about `room` tokens: whole rows from the top, then '… N more rows not shown'.
    -> (text, rows kept or None when nothing was cut)."""
    if tokens(text) <= room:
        return text, None
    lines = text.split("\n")
    head_n = next((i + 1 for i, x in enumerate(lines) if x.startswith("Cells ")), len(lines))
    head, rows = lines[:head_n], lines[head_n:]
    tail = rows.pop() if rows and rows[-1].startswith("… ") else ""
    kept, used = [], tokens("\n".join(head)) + 40
    for x in rows:
        used += tokens(x)
        if used > room:
            break
        kept.append(x)
    more = len(rows) - len(kept)
    note = f"… {more} more rows not shown (too many for this model; ask about a selection to see them)."
    return "\n".join(head + kept + [note] + ([tail] if tail else [])), len(kept)


def ask(messages, model, key, transport=http, web=False, usage=None):
    """-> the model's reply text. web=True: OpenRouter's web search plugin looks things up first, and the
    pages it used are added at the end as "Sources". usage: a list that gets {"in", "out", "cost"} of this call
    (filled on the worker thread; record_usage stores it on the main thread)."""
    if not key:
        raise AIError("Add your OpenRouter API key first (⚙ settings).")
    if not model:
        raise AIError("Choose a model first (⚙ settings).")
    body = {"model": model, "messages": messages, "usage": {"include": True}}
    if web:
        body["plugins"] = [{"id": "web", "max_results": WEB_RESULTS}]
        if messages and messages[0].get("role") == "system":
            body["messages"] = [{**messages[0], "content": messages[0]["content"] + WEB_NOTE}] + messages[1:]
    data = transport("POST", f"{API}/chat/completions", key, body)
    if data.get("error"):
        raise AIError(f"OpenRouter: {data['error'].get('message', data['error'])}")
    u = data.get("usage")
    if usage is not None and isinstance(u, dict):
        try:
            usage.append({"in": int(u.get("prompt_tokens") or 0), "out": int(u.get("completion_tokens") or 0),
                          "cost": float(u.get("cost") or 0)})
        except (TypeError, ValueError):
            pass
    try:
        msg = data["choices"][0]["message"]
        text = msg["content"] or ""
    except (KeyError, IndexError, TypeError):
        raise AIError("The model sent an empty or unexpected reply.") from None
    sources = web_sources(msg)
    if sources and not any(url in text for _, url in sources):
        text = text.rstrip() + "\n\n**Sources**\n" + "\n".join(f"- [{t}]({u})" for t, u in sources)
    return text


WEB_RESULTS = 5          # pages the web search reads per question (OpenRouter charges per result)
AUTO_FIXES = 2           # Auto mode: how many times a failing script goes back to the AI by itself


def web_sources(msg):
    """[(title, url)] of the pages a web-search answer used (OpenRouter's url_citation annotations)."""
    out, seen = [], set()
    for a in (msg or {}).get("annotations") or []:
        cite = a.get("url_citation") if isinstance(a, dict) else None
        url = (cite or {}).get("url")
        if url and url.startswith(("http://", "https://")) and url not in seen:
            seen.add(url)
            out.append(((cite.get("title") or url).replace("]", ")").replace("[", "(")[:80], url))
    return out


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


def delete_chat_file(chat_id):
    import os
    try:
        os.remove(os.path.join(chats_dir(), f"{os.path.basename(chat_id)}.json"))
        return True
    except OSError:
        return False


def clear_chats():
    import glob
    import os
    for p in glob.glob(os.path.join(chats_dir(), "*.json")):
        try:
            os.remove(p)
        except OSError:
            pass


ICON_FONTS = ("Segoe Fluent Icons", "Segoe MDL2 Assets")   # Windows' own icon fonts (11 / 10)
ICONS = {"new": "\ue70f", "history": "\ue81c", "settings": "\ue713", "hide": "\ue89f", "add": "\ue710",
         "auto": "\ue945", "web": "\ue774", "delete": "\ue74d"}
FALLBACK = {"new": "✎", "history": "↺", "settings": "⚙", "hide": "»", "add": "+", "auto": "⚡", "web": "◎",
            "delete": "✕"}


def icon_font(root):
    """The Windows icon font when there is one (crisp, matching outline icons), else None."""
    from tkinter import font as tkfont
    for fam in ICON_FONTS:
        try:
            if tkfont.Font(root=root, family=fam, size=10).actual("family") == fam:
                return fam
        except Exception:
            pass
    return None


def icon(root, name):
    """(text, font) for an icon button: the icon font's glyph, or a plain symbol."""
    fam = icon_font(root)
    return (ICONS[name], (fam, 11)) if fam else (FALLBACK[name], ("Segoe UI Symbol", 11))


class Files(dict):
    """attachments[...] in a script: a clear message when a name is wrong."""

    def __missing__(self, name):
        import scripting
        have = ", ".join(repr(k) for k in self) or "none in this chat"
        raise scripting.ScriptError(f"No attached file {name!r} (attached: {have})")


class Conversation:
    """The chat as the model sees it. Each question is fitted to the model's context window: the question,
    attachments and the sheet first (the sheet cut to fewer rows if needed), then as much of the chat as fits,
    newest first. Older messages are folded into a short summary (made by the model, see summary_request)."""

    def __init__(self):
        self.messages = []
        self.summary, self.summary_upto = "", 0    # summary of messages[:summary_upto]
        self.dropped = 0                            # messages after summary_upto left out last time (to summarise)
        self.sheet_rows = None                      # rows of the sheet sent last time when it had to be cut

    def build(self, app, question, scope, attachments=(), ctx=None):
        budget = input_budget(ctx or DEFAULT_CONTEXT)
        system = SYSTEM + (f"\n\nEarlier in this chat (summary):\n{self.summary}" if self.summary else "")
        parts = [a.part() for a in attachments]
        fixed = tokens(system) + tokens(question) + tokens(parts) + 30
        sheet, self.sheet_rows = cut_sheet(context(app, scope), max(int(budget * 0.7) - fixed, 300))
        text = f"{sheet}\n\nRequest: {question}"
        room = budget - fixed - tokens(sheet)
        recent = self.messages[self.summary_upto:]
        keep = 0
        while keep + 2 <= min(len(recent), HISTORY_MAX):       # whole question / answer pairs, newest first
            pair = recent[len(recent) - keep - 2:len(recent) - keep]
            if tokens(pair) > room:
                break
            room -= tokens(pair)
            keep += 2
        self.dropped = len(recent) - keep
        msgs = [{"role": "system", "content": system}] + (recent[len(recent) - keep:] if keep else [])
        if parts:                            # text first, then the pictures / files / attached text
            msgs.append({"role": "user", "content": [{"type": "text", "text": text}] + parts})
        else:
            msgs.append({"role": "user", "content": text})
        return msgs

    def summary_request(self):
        """(messages for the model that write the new summary, index it covers up to), or None when nothing
        was left out. The messages that didn't fit last time are folded into the summary so far."""
        if self.dropped <= 0:
            return None
        upto = self.summary_upto + self.dropped
        old = self.messages[self.summary_upto:upto]
        convo = "\n\n".join(f"{m['role'].upper()}: {m['content'] if isinstance(m['content'], str) else '(attachment)'}"
                            [:3000] for m in old)
        prompt = ("Summarise this part of a conversation between a user and a spreadsheet assistant, to be "
                  "remembered later, in under 200 words: what the user wants, decisions made, what was changed "
                  "in the sheet, and any facts or numbers worth keeping. Plain text, no code.\n\n"
                  + (f"Summary of what came before:\n{self.summary}\n\n" if self.summary else "")
                  + f"Conversation:\n{convo}")
        return [{"role": "user", "content": prompt}], upto

    def set_summary(self, text, upto):
        if text.strip() and upto > self.summary_upto:
            end = self.summary_upto + self.dropped            # left out at the last question (maybe more by now)
            self.summary, self.summary_upto = text.strip(), upto
            self.dropped = max(0, end - upto)

    def add(self, question, reply, attachments=()):
        if attachments:                      # remembered by name only (not sent again with every question)
            question += f"  [attached: {', '.join(a.name for a in attachments)}]"
        self.messages += [{"role": "user", "content": question}, {"role": "assistant", "content": reply}]

    def note(self, text):
        """What happened when the script ran (so a follow-up question can build on it)."""
        self.messages.append({"role": "user", "content": f"(The script was run. Result: {text})"})
        self.messages.append({"role": "assistant", "content": "Noted."})


# ---------- the pane (docked on the right of the window) ----------

INLINE = re.compile(r"(\*\*[^*\n]+\*\*|`[^`\n]+`|\[[^\]\n]+\]\(https?://[^)\s]+\)|https?://[^\s)\]>]+)")
LINK = re.compile(r"\[([^\]\n]+)\]\((https?://[^)\s]+)\)")


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


class HistoryPanel:
    """↺ drop-down: one row per saved chat (title, when, a bin that deletes it), Clear all at the bottom.
    A borderless window under the button; a click outside it or Esc closes it."""

    def __init__(self, assistant):
        import tkinter as tk

        from theme import C
        self.a = assistant
        root = assistant.win.winfo_toplevel()
        self.top = tk.Toplevel(root)
        self.top.withdraw()
        self.top.overrideredirect(True)
        self.top.configure(background=C["header_line"])
        self.box = tk.Frame(self.top, background=C["menu_bg"], padx=4, pady=4)
        self.box.pack(fill="both", expand=True, padx=1, pady=1)      # 1 px line round it
        self.title_font = ("Segoe UI", 9)
        self.icon_text, self.icon_font = icon(root, "delete")
        self.rows = {}
        self.fill()
        self.top.bind("<Escape>", lambda e: self.destroy())
        self.top.bind("<Button-1>", self.click_outside)

    def fill(self):
        import time
        import tkinter as tk

        from theme import C
        for w in self.box.winfo_children():
            w.destroy()
        self.rows = {}
        chats = list_chats()[:30]
        tk.Label(self.box, text="Earlier chats", background=C["menu_bg"], foreground=C["muted"],
                 font=("Segoe UI", 8, "bold"), anchor="w", padx=6).pack(fill="x", pady=(0, 2))
        if not chats:
            tk.Label(self.box, text="No earlier chats yet", background=C["menu_bg"], foreground=C["muted"],
                     font=("Segoe UI", 9, "italic"), anchor="w", padx=6, pady=4).pack(fill="x")
        for c in chats:
            row = tk.Frame(self.box, background=C["menu_bg"], cursor="hand2")
            row.pack(fill="x")
            current = c["id"] == self.a.chat_id
            title = c["title"] if len(c["title"]) <= 38 else c["title"][:37] + "…"
            lab = tk.Label(row, text=title, anchor="w", padx=6, pady=3, background=C["menu_bg"],
                           foreground=C["accent"] if current else C["fg"],
                           font=self.title_font + (("bold",) if current else ()))
            lab.pack(side="left", fill="x", expand=True)
            bin_ = tk.Label(row, text=self.icon_text, font=self.icon_font, padx=6, background=C["menu_bg"],
                            foreground=C["muted"], cursor="hand2")
            bin_.pack(side="right")
            when = tk.Label(row, text=time.strftime("%d %b %H:%M", time.localtime(c.get("updated", 0))),
                            padx=4, background=C["menu_bg"], foreground=C["muted"], font=("Segoe UI", 8))
            when.pack(side="right")
            parts = (row, lab, when, bin_)
            for w in parts:
                w.bind("<Enter>", lambda e, p=parts: self.hover(p, True))
                w.bind("<Leave>", lambda e, p=parts: self.hover(p, False))
            for w in (row, lab, when):
                w.bind("<Button-1>", lambda e, i=c["id"]: (self.open(i), "break")[-1])
            bin_.bind("<Button-1>", lambda e, i=c["id"]: (self.delete(i), "break")[-1])
            bin_.bind("<Enter>", lambda e, b=bin_: b.configure(foreground=C["danger"]), add="+")
            bin_.bind("<Leave>", lambda e, b=bin_: b.configure(foreground=C["muted"]), add="+")
            self.rows[c["id"]] = parts
        if chats:
            tk.Frame(self.box, height=1, background=C["header_line"]).pack(fill="x", pady=4)
            clear = tk.Label(self.box, text="Clear all…", anchor="w", padx=6, pady=2, cursor="hand2",
                             background=C["menu_bg"], foreground=C["danger"], font=self.title_font)
            clear.pack(fill="x")
            clear.bind("<Button-1>", lambda e: (self.clear_all(), "break")[-1])

    def hover(self, parts, on):
        from theme import C
        bg = C["menu_active"] if on else C["menu_bg"]
        for w in parts:
            w.configure(background=bg)

    def show_under(self, button):
        """Right-aligned under the ↺ button, then take the mouse so a click outside closes it."""
        self.top.update_idletasks()
        w = max(self.top.winfo_reqwidth(), int(280 * self.a.app.scale))
        x = button.winfo_rootx() + button.winfo_width() - w
        self.top.geometry(f"{w}x{self.top.winfo_reqheight()}+{max(x, 0)}+{button.winfo_rooty() + button.winfo_height() + 2}")
        self.top.deiconify()
        self.top.lift()
        self.top.focus_force()
        try:
            self.top.grab_set()
        except Exception:
            pass

    def click_outside(self, e):
        x, y = e.x_root - self.top.winfo_rootx(), e.y_root - self.top.winfo_rooty()
        if not (0 <= x < self.top.winfo_width() and 0 <= y < self.top.winfo_height()):
            self.destroy()

    def open(self, chat_id):
        self.destroy()
        return self.a.open_chat(chat_id)

    def delete(self, chat_id):
        self.a.delete_chat(chat_id)
        self.fill()                                    # the list without it; the panel stays open
        if self.top.winfo_viewable():
            self.top.geometry(f"{self.top.winfo_width()}x{self.top.winfo_reqheight()}")
        return True

    def clear_all(self):
        try:
            self.top.grab_release()
        except Exception:
            pass
        self.destroy()
        return self.a.clear_history()

    def destroy(self):
        if self.top.winfo_exists():
            self.top.destroy()

    def winfo_exists(self):
        return self.top.winfo_exists()


class Assistant:
    """The AI assistant as a pane on the right of the main window (Ctrl+Shift+A shows / hides it)."""

    def __init__(self, app, transport=http):
        import tkinter as tk
        from tkinter import ttk

        from theme import C
        self.app, self.transport, self.conv = app, transport, Conversation()
        self.fetch_sizes = transport is http          # look up the models' context sizes (not with test transports)
        self.learn_sizes(settings.get("ai_model"))
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

        # one slim row: ✦ AI Assistant  …  ⚡ Auto · 🌐 Web | ✎ new chat · ↺ history · ⚙ · hide
        head = ttk.Frame(body)
        head.pack(fill="x", pady=(0, 6))
        from app_menu import sparkle_icon
        from format_bar import Tooltip
        self.mark = sparkle_icon(outer, app.scale, C["accent"])
        ttk.Label(head, image=self.mark).pack(side="left", padx=(0, 4))
        self.title_label = ttk.Label(head, text="AI Assistant", font=("Segoe UI", 10, "bold"))
        self.title_label.pack(side="left")
        self.title_tip = Tooltip(self.title_label, "")              # which model answers (chosen in ⚙)
        self.icon_buttons = {}
        for name, cmd, tip in (("hide", self.hide, "Hide the panel (it keeps working) - Ctrl+Shift+A"),
                               ("settings", self.settings, "AI settings: model and OpenRouter key"),
                               ("history", self.history_menu, "Earlier chats"),
                               ("new", self.new_chat, "New chat"),
                               # toggles: lit (accent) while on
                               ("web", lambda: self.toggle_web(),
                                "Search the web: the AI looks things up online before answering and lists its "
                                "sources (OpenRouter charges a little extra per search).  Click to switch."),
                               ("auto", lambda: self.toggle_mode(),
                                "Auto: the AI's scripts run straight away (still one Ctrl+Z).  Off (Ask): you "
                                "press Run.  Click to switch.")):
            text, font = icon(outer, name)
            b = tk.Label(head, text=text, font=font, cursor="hand2", padx=5, pady=1)
            b._own_colors = True
            b.pack(side="right", padx=(0, 8) if name == "web" else 0)   # a gap between toggles and actions
            b.bind("<Button-1>", lambda e, c=cmd: c())
            Tooltip(b, tip)
            self.icon_buttons[name] = b
        self.mode_chip, self.web_chip = self.icon_buttons["auto"], self.icon_buttons["web"]
        self.history_btn = self.icon_buttons["history"]
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
        # after a failed run: send the error back so the AI corrects its script
        self.fix_btn = ttk.Button(top, text="↻ Fix it", command=self.fix)
        self.failed = None                    # (code, error) of the last failed run, until the card changes
        self.auto_fixes = 0                   # automatic fix tries in Auto mode for this question
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
        plus_text, plus_font = icon(outer, "add")
        self.plus = tk.Label(self.composer, text=plus_text, font=(plus_font[0], 12), cursor="hand2", padx=4, pady=0)
        self.plus._own_colors = True
        self.prompt = tk.Text(self.composer, height=1, wrap="word", font=("Segoe UI", 10), relief="flat",
                              borderwidth=0, highlightthickness=0, padx=2, pady=max(int(4 * app.scale), 3))
        self.prompt._own_colors = True
        self.send_btn = tk.Canvas(self.composer, width=30, height=30, highlightthickness=0, borderwidth=0,
                                  cursor="hand2")
        self.send_btn._own_colors = True
        self.sending = False
        self.scope = tk.StringVar(master=outer, value=settings.get("ai_scope", "sheet"))
        self.plus_menu = tk.Menu(self.plus, tearoff=False)
        self.plus_menu.add_command(label="Attach a file…", command=self.attach)
        self.plus_menu.add_command(label="Capture part of the screen", command=self.capture)
        self.plus_menu.add_command(label="Paste a picture (Ctrl+V)", command=lambda: self.paste() or None)
        self.plus_menu.add_separator()
        self.plus_menu.add_radiobutton(label="Send the sheet", value="sheet", variable=self.scope)
        self.plus_menu.add_radiobutton(label="Send only the selection", value="selection", variable=self.scope)
        self.web_var = tk.BooleanVar(master=outer, value=self.web())
        self.plus_menu.add_checkbutton(label="Search the web", variable=self.web_var,
                                       command=lambda: self.toggle_web(self.web_var.get()))
        self.plus_menu.add_separator()
        self.mode_var = tk.StringVar(master=outer, value="auto" if settings.get("ai_auto_run") else "ask")
        self.plus_menu.add_radiobutton(label="Scripts: I check them first (Ask)", value="ask", variable=self.mode_var,
                                       command=lambda: self.set_mode(False))
        self.plus_menu.add_radiobutton(label="Scripts: run automatically (Auto)", value="auto",
                                       variable=self.mode_var, command=lambda: self.set_mode(True))
        self.plus.bind("<Button-1>", lambda e: self.plus_menu.tk_popup(e.x_root, e.y_root))
        self.send_btn.bind("<Button-1>", lambda e: self.stop() if self.busy else self.send())
        from format_bar import Tooltip
        Tooltip(self.plus, "Attach a file, capture the screen, or choose what to send (sheet / selection)")
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
        import edit_menu  # right click: copy / paste / select all…
        edit_menu.attach(self.chat, extra=self.link_items)
        edit_menu.attach(self.prompt, paste=self.paste_text)
        edit_menu.attach(self.code)
        self.prompt.bind("<FocusIn>", lambda e: self.hint(False))
        self.prompt.bind("<FocusOut>", lambda e: self.hint(True))
        self.recolor()
        self.paint_mode()
        self.hint(True)
        self.refresh_label()
        self.show()
        if not load_key() or not settings.get("ai_model"):
            self.later(200, lambda: outer.winfo_exists() and self.settings())

    def later(self, ms, fn, *args):
        """A timer on the main window that is cancelled if the pane closes first (no stray callbacks)."""
        if not hasattr(self, "_timers"):
            self._timers = set()
            self.win.bind("<Destroy>", lambda e: e.widget is self.win and self.cancel_timers(), add="+")
        tid = self.app.after(ms, lambda: (self._timers.discard(tid), fn(*args)))
        self._timers.add(tid)
        return tid

    def cancel_timers(self):
        for tid in list(getattr(self, "_timers", ())):
            try:
                self.app.after_cancel(tid)
            except Exception:
                pass
        self._timers = set()

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
        self.card_title.configure(background=card,
                                  foreground=C["danger"] if getattr(self, "failed", None) else C["fg"])
        self.card_desc.configure(background=card, foreground=C["muted"])
        for lab in self.card_links:
            lab.configure(background=card, foreground=C["accent"])
        self.code.configure(background=C["field_bg"], foreground=C["fg"], insertbackground=C["fg"])
        for b in getattr(self, "icon_buttons", {}).values():
            b.configure(background=C["bg"], foreground=C["icon"])
        if hasattr(self, "mode_chip"):
            self.paint_mode()
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
        t.tag_configure("working", foreground=C["muted"], font=("Segoe UI", 9), spacing1=8, spacing3=4)
        t.tag_configure("link", foreground=C["accent"], underline=True)
        t.tag_bind("link", "<Enter>", lambda e: t.configure(cursor="hand2"))
        t.tag_bind("link", "<Leave>", lambda e: t.configure(cursor="arrow"))
        if getattr(self, "dots", None) is not None:
            self.dots.configure(background=C["chat_bg"])
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
        text_w = max(send_x - btn // 2 - int(6 * k) - text_x, 60)
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
        cv.coords("plus", plus_x, mid)
        cv.coords("prompt", text_x, mid)
        cv.itemconfigure("prompt", width=text_w, height=text_h)
        cv.coords("send", send_x, mid)
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
        fill = C["danger"] if self.sending else C["accent"]      # working: a red Stop button
        b.create_oval(1, 1, s - 1, s - 1, fill=fill, outline=shade(fill, 0.15) if not C.get("contrast") else fill)
        ink, m, k = C["on_accent"], s / 2, s / 30
        if self.sending:
            square = "#ffffff" if not C["dark"] else C["on_accent"]
            b.create_rectangle(m - 5 * k, m - 5 * k, m + 5 * k, m + 5 * k, fill=square, outline=square)
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
        """The title's tooltip names the model (chosen in ⚙ settings)."""
        model = settings.get("ai_model") or ""
        self.title_tip.text = f"Model: {model}  (change it in ⚙)" if model else "No model yet: choose one in ⚙"

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
            elif LINK.fullmatch(part) or part.startswith(("http://", "https://")):
                m = LINK.fullmatch(part)
                label, url = (m.group(1), m.group(2)) if m else (part, part)
                self.links = getattr(self, "links", {})
                tag = f"link{len(self.links)}"           # each link its own tag: a click opens its page
                self.links[tag] = url
                self.chat.insert("end", label, base + ("link", tag))
                self.chat.tag_bind(tag, "<Button-1>", lambda e, u=url: self.open_link(u))
            else:
                self.chat.insert("end", part, base)

    def open_link(self, url):
        import webbrowser
        if url.startswith(("http://", "https://")):     # only web pages, never a file or program
            webbrowser.open(url)
        return url

    def say(self, who, text, record=True):
        """who: 'You', a model name, or None for a grey note."""
        if record:
            self.transcript.append([who, text])
        self.clear_thinking()
        c = self.chat
        c.configure(state="normal")
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
        """A live 'working' line: three pulsing dots, what it is doing, seconds so far (until the answer comes)."""
        import time
        import tkinter as tk

        from theme import C
        self.think_started, self.think_what, self.think_step = time.monotonic(), what, 0
        self.clear_thinking()
        c, k = self.chat, self.app.scale
        r = max(round(3.5 * k), 3)
        gap = 2 * r + max(round(4 * k), 3)
        self.dots = tk.Canvas(c, width=gap * 2 + r * 2 + 2, height=r * 2 + 2, highlightthickness=0, borderwidth=0,
                              background=C["chat_bg"])
        self.dot_items = [self.dots.create_oval(1 + i * gap, 1, 1 + i * gap + 2 * r, 1 + 2 * r, outline="")
                          for i in range(3)]
        c.configure(state="normal")
        c.mark_set("think", "end-1c")
        c.mark_gravity("think", "left")
        c.insert("end", " ", ("working",))
        c.window_create("end", window=self.dots, align="center")
        c.insert("end", "  ", ("working",))
        c.insert("end", "x", ("working", "think_text"))
        c.insert("end", "\n", ("working",))
        c.tag_add("thinking", "think", "end-1c")
        c.configure(state="disabled")
        self.paint_thinking()
        self.later(140, self.tick_thinking, self.request)

    def paint_thinking(self):
        """Next animation frame: one dot bright, the others faded; the text says what and for how long."""
        import time

        from grid import blend
        from theme import C
        c = self.chat
        if not c.tag_ranges("think_text"):
            return
        secs = int(time.monotonic() - getattr(self, "think_started", time.monotonic()))
        lit = self.think_step % 3
        for i, item in enumerate(self.dot_items):
            self.dots.itemconfigure(item, fill=C["accent"] if i == lit else blend(C["accent"], C["chat_bg"], 0.65))
        text = f"{self.think_what or 'Thinking'}…" + (f"  ·  {secs}s" if secs >= 2 else "")
        if c.get("think_text.first", "think_text.last") != text:
            c.configure(state="normal")
            c.delete("think_text.first", "think_text.last")
            c.insert("thinking.last -1c", text, ("working", "think_text", "thinking"))
            c.configure(state="disabled")
        c.see("end")

    def clear_thinking(self):
        c = self.chat
        if c.tag_ranges("thinking"):
            c.configure(state="normal")
            c.delete("thinking.first", "thinking.last")
            c.configure(state="disabled")
        if getattr(self, "dots", None) is not None:
            self.dots.destroy()
            self.dots = None

    def tick_thinking(self, request):
        if not self.busy or request != self.request or not self.win.winfo_exists():
            return
        self.think_step += 1
        self.paint_thinking()
        self.later(140 if self.think_step % 3 else 420, self.tick_thinking, request)   # a short pause each round

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
        self.fit_noted = set()
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
                                             "shown": self.transcript, "summary": self.conv.summary,
                                             "summary_upto": self.conv.summary_upto})

    def open_chat(self, chat_id):
        data = load_chat_file(chat_id)
        if not data:
            return False
        self.save_chat()
        self.new_chat_view()
        self.chat_id = chat_id
        self.conv.messages = list(data.get("messages") or [])
        self.conv.summary = data.get("summary") or ""
        self.conv.summary_upto = min(int(data.get("summary_upto") or 0), len(self.conv.messages))
        self.transcript = [list(x) for x in data.get("shown") or []]
        for who, text in self.transcript:
            self.say(who, text, record=False)
        return True

    def new_chat_view(self):
        self.conv = Conversation()
        self.fit_noted = set()
        self.chat.configure(state="normal")
        self.chat.delete("1.0", "end")
        self.chat.configure(state="disabled")
        self.set_code(None)

    def history_menu(self, show=True):
        """↺: a drop-down list of earlier chats under the button: click one to continue it, its bin deletes it;
        Clear all at the bottom. show=False builds it without taking the mouse (tests)."""
        old = getattr(self, "history_panel", None)
        if old is not None and old.winfo_exists():
            old.destroy()
            return None
        p = self.history_panel = HistoryPanel(self)
        if show:
            p.show_under(self.history_btn)
        return p

    def delete_chat(self, chat_id):
        """Delete one saved chat; deleting the one on screen starts an empty chat."""
        delete_chat_file(chat_id)
        if chat_id == self.chat_id:
            self.new_chat_view()
            self.chat_id, self.transcript, self.files = None, [], {}
        return True

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
        self.set_failed(None)
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

    def set_mode(self, auto):
        if auto != self.auto():
            self.toggle_mode()
        self.mode_var.set("auto" if self.auto() else "ask")

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
        if hasattr(self, "mode_var"):
            self.mode_var.set("auto" if self.auto() else "ask")
        return self.auto()

    def paint_mode(self):
        from grid import blend
        from theme import C
        lit = blend(C["accent"], C["bg"], 0.82)                  # a soft accent tint behind a switched-on icon
        for chip, on in ((self.mode_chip, self.auto()), (self.web_chip, self.web())):
            chip.configure(foreground=C["accent"] if on else C["icon"], background=lit if on else C["bg"])

    def web(self):
        return bool(settings.get("ai_web", False))

    def toggle_web(self, on=None):
        """Web search on / off (remembered)."""
        settings.set("ai_web", (not self.web()) if on is None else bool(on))
        if hasattr(self, "web_var"):
            self.web_var.set(self.web())
        self.paint_mode()
        return self.web()

    def send(self, question=None, shown=None, keep_fixes=False):
        """Ask the model. shown: a shorter text for the chat than what is sent (↻ Fix it sends the whole script)."""
        if self.busy:
            return None
        question = (question or self.question_text()).strip()
        if not question:
            return None
        if not keep_fixes:
            self.auto_fixes = 0                    # a new question: Auto mode may try fixing its script again
        settings.set("ai_scope", self.scope.get())
        attached = list(self.attachments)
        for a in attached:                    # attached data stays usable by scripts for the rest of the chat
            if a.table is not None:
                self.files[a.name] = a.table
            elif a.kind == "text":
                self.files[a.name] = a.data
        key, model = load_key(), settings.get("ai_model")
        msgs = self.conv.build(self.app, question, self.scope.get(), attached, ctx=context_size(model))
        self.prompt.delete("1.0", "end")
        self.prompt._hinting = False
        self.say("You", (shown or question) + (f"\n📎 {', '.join(a.name for a in attached)}" if attached else ""))
        self.fit_notes(model)
        self.clear_attachments()
        self.request += 1
        request = self.request
        self.busy = True
        web = self.web()
        self.thinking(f"Reading {attached[0].name}" if attached else "Searching the web" if web
                      else "Thinking")                                  # the model is in the header
        self.set_sending(True)

        used = []                                      # tokens / cost of this call (counted in ⚙ → Usage)

        def work():
            try:
                self.q.put(("ok", (question, attached, request, used),
                            ask(msgs, model, key, self.transport, web=web, usage=used)))
            except AIError as e:
                if too_big(str(e)):
                    hint = (" The question is too big for this model. Try: + → Send only the selection, fewer or "
                            "smaller attachments, ✎ a new chat, or a model with a larger context (⚙).")
                elif any(a.kind != "text" for a in attached):
                    hint = " This model may not read pictures or PDFs: pick one whose list entry says 'image' / 'file' (⚙)."
                else:
                    hint = ""
                self.q.put(("error", (question, attached, request), str(e) + hint))
            except Exception as e:                       # never leave the pane waiting
                self.q.put(("error", (question, attached, request), f"{type(e).__name__}: {e}"))
        threading.Thread(target=work, daemon=True).start()
        self.later(100, self.poll)
        return msgs

    def poll(self):
        try:
            kind, question, text = self.q.get_nowait()
        except queue.Empty:
            if self.win.winfo_exists():
                self.later(100, self.poll)
            return
        if len(question) > 3:
            record_usage(question[3])         # paid for even when stopped: counted in ⚙ → Usage
        if question[2] != self.request:      # stopped (or an older question): ignore this answer
            if self.busy and self.win.winfo_exists():
                self.later(100, self.poll)
            return
        self.busy = False
        if not self.win.winfo_exists():
            return
        self.set_sending(False)
        if kind == "error":
            self.say(None, "⚠ " + text)
            return
        question, attached = question[0], question[1]
        self.conv.add(question, text, attached)
        code, words = extract_code(text)
        self.say(settings.get("ai_model") or "AI", words or ("Here is a script for that." if code else "(no answer)"))
        self.set_code(code)
        if code and self.auto():
            self.run(auto=True)                       # Auto mode: straight away (one Ctrl+Z, rolled back on errors)
        elif code:
            self.say(None, "A script is ready below: check it, then press ▶ Run.")
        self.save_chat()
        self.summarize()

    # --- long chats: fitted to the model, older messages summarised ---

    def fit_notes(self, model):
        """Grey notes (once per chat) when the question had to be fitted: chat history left out, sheet cut."""
        self.learn_sizes(model)
        noted = self.__dict__.setdefault("fit_noted", set())
        if (self.conv.dropped or self.conv.summary) and "chat" not in noted:
            noted.add("chat")
            self.say(None, "This chat is getting long for the model: older messages are kept as a short summary. "
                           "✎ starts a fresh chat.")
        if self.conv.sheet_rows is not None and "sheet" not in noted:
            noted.add("sheet")
            self.say(None, f"The sheet is too big for this model: only its first {self.conv.sheet_rows} rows were "
                           "sent. Select the cells that matter and use + → Send only the selection, or pick a model "
                           "with a larger context (⚙).")

    def learn_sizes(self, model):
        """Look up the models' context sizes once (in the background) when this model's isn't known."""
        if model and model not in _context_sizes and self.fetch_sizes and not getattr(self, "sizes_asked", False):
            self.sizes_asked = True
            threading.Thread(target=fetch_context_sizes, args=(self.transport,), daemon=True).start()

    def summarize(self):
        """Fold the messages that no longer fit into the chat's summary (one model call, in the background)."""
        req = self.conv.summary_request()
        if not req or getattr(self, "summarizing", False):
            return None
        msgs, upto = req
        conv, key, model = self.conv, load_key(), settings.get("ai_model")
        self.summarizing = True
        out = queue.Queue()

        used = []

        def work():
            try:
                out.put(ask(msgs, model, key, self.transport, usage=used))
            except Exception:
                out.put(None)                          # no summary this time: tried again after the next answer

        def done():
            try:
                text = out.get_nowait()
            except queue.Empty:
                return self.later(200, done) if self.win.winfo_exists() else None
            self.summarizing = False
            record_usage(used)
            if text and conv is self.conv:
                conv.set_summary(text, upto)
                self.save_chat()
        threading.Thread(target=work, daemon=True).start()
        self.later(200, done)
        return upto

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

    def paste_text(self):
        """Paste from the right-click menu: a picture becomes an attachment, text goes into the question."""
        if self.paste() == "break":
            return "picture"
        self.hint(False)
        self.prompt.event_generate("<<Paste>>")
        self.layout_composer()
        return "text"

    def link_items(self, e):
        """Right click on a link in the chat: Open link / Copy link above the usual items."""
        tag = next((t for t in self.chat.tag_names(f"@{e.x},{e.y}") if t in getattr(self, "links", {})), None)
        if not tag:
            return []
        url = self.links[tag]

        def copy():
            self.chat.clipboard_clear()
            self.chat.clipboard_append(url)
        return [("Open link", lambda: self.open_link(url)), ("Copy link", copy)]

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
        self.conv.note(msg + (" Output: " + "".join(out)[:500] if out else ""))
        self.save_chat()
        if ok:
            self.set_code(None)
            self.auto_fixes = 0
        elif not msg.startswith("Stopped"):          # an error: offer ↻ Fix it (Auto mode tries it by itself)
            self.set_failed((code, msg, "".join(out)))
            if auto and self.auto_fixes < AUTO_FIXES:
                self.auto_fixes += 1
                self.say(None, f"Asking the AI to fix it (try {self.auto_fixes} of {AUTO_FIXES})…")
                self.fix(auto=True)
            elif auto:
                self.show_code(True)                 # still failing: the card stays, to look at / edit / retry
        return ok

    def set_failed(self, failed):
        """The card after a failed run: 'Script failed', the error, and ↻ Fix it next to ▶ Run."""
        from theme import C
        self.failed = failed
        if failed:
            first = failed[1].split("\n", 1)[0]
            self.card_title.configure(text="Script failed", foreground=C["danger"])
            self.card_desc.configure(text=first)
            if not self.fix_btn.winfo_manager():
                self.fix_btn.pack(side="right", padx=(0, 6), before=self.run_btn)
        else:
            self.card_title.configure(foreground=C["fg"])
            self.fix_btn.pack_forget()

    def fix(self, auto=False):
        """↻ Fix it: send the failed script and its error back; the AI answers with a corrected script."""
        if not self.failed or self.busy:
            return None
        code, error, output = self.failed
        if not auto:
            self.auto_fixes = AUTO_FIXES          # a click is one try (Auto mode doesn't start another round)
        question = ("The script failed. Error:\n```\n" + error + "\n```\n"
                    + (f"It printed before failing:\n```\n{output[:1500]}\n```\n" if output.strip() else "")
                    + "The script that ran:\n```python\n" + code[:12000] + "\n```\n"
                    "Find the cause and send the whole corrected script (nothing was changed by the failed run).")
        return self.send(question, shown="↻ Fix the script: " + error.split("\n", 1)[0], keep_fixes=True)

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


def center(win, parent):
    """Put a dialog in the middle of its parent window (kept on the screen)."""
    win.update_idletasks()
    top = parent.winfo_toplevel()
    w, h = win.winfo_reqwidth(), win.winfo_reqheight()
    x = top.winfo_rootx() + (top.winfo_width() - w) // 2
    y = top.winfo_rooty() + (top.winfo_height() - h) // 3
    x = max(0, min(x, win.winfo_screenwidth() - w))
    y = max(0, min(y, win.winfo_screenheight() - h - 40))
    win.geometry(f"+{x}+{y}")
    return x, y


class SettingsDialog:
    """⚙: OpenRouter key (with the account's credits and spending), model, and the usage counted here."""

    def __init__(self, parent, assistant=None):
        import tkinter as tk
        from tkinter import ttk

        from theme import C
        self.assistant = assistant
        self.transport = assistant.transport if assistant else http
        scale = getattr(parent.winfo_toplevel(), "scale", 1.0) or 1.0
        w = self.win = tk.Toplevel(parent)
        w.withdraw()
        w.title("AI settings")
        w.transient(parent.winfo_toplevel())
        w.resizable(False, False)
        w.configure(background=C["bg"])
        self.results = queue.Queue()                # answers of the background loads, for the main thread
        f = ttk.Frame(w, padding=(18, 14, 18, 14))
        f.pack(fill="both", expand=True)
        f.columnconfigure(0, weight=1)
        wrap = int(470 * scale)

        def heading(text, row, top=14):
            ttk.Label(f, text=text, font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", pady=(top, 4))

        def muted(parent_, text, **kw):
            return tk.Label(parent_, text=text, background=C["bg"], foreground=C["muted"], font=("Segoe UI", 8),
                            anchor="w", justify="left", **kw)

        # --- account: key, credits, spending ---
        heading("OpenRouter account", 0, top=0)
        keyrow = ttk.Frame(f)
        keyrow.grid(row=1, column=0, sticky="ew")
        keyrow.columnconfigure(0, weight=1)
        self.key = tk.StringVar(master=w, value=load_key())
        self.key_entry = ttk.Entry(keyrow, textvariable=self.key, show="•")
        self.key_entry.grid(row=0, column=0, sticky="ew")
        self.show_key = tk.BooleanVar(master=w, value=False)
        ttk.Checkbutton(keyrow, text="Show", variable=self.show_key,
                        command=lambda: self.key_entry.configure(show="" if self.show_key.get() else "•")).grid(
            row=0, column=1, padx=(8, 0))
        links = tk.Frame(f, background=C["bg"])
        links.grid(row=2, column=0, sticky="w", pady=(3, 0))
        for text, cmd in (("Get a key", lambda: self.open_page("https://openrouter.ai/keys")),
                          ("Add credits", lambda: self.open_page("https://openrouter.ai/settings/credits")),
                          ("↻ Refresh", self.start_account)):
            lab = tk.Label(links, text=text, background=C["bg"], foreground=C["accent"], cursor="hand2",
                           font=("Segoe UI", 9, "underline"))
            lab.pack(side="left", padx=(0, 14))
            lab.bind("<Button-1>", lambda e, c=cmd: c())
        card = self.card = tk.Frame(f, background=C["chat_bg"], highlightthickness=1,
                                    highlightbackground=C["header_line"], padx=12, pady=10)
        card.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        card.columnconfigure(1, weight=1)
        self.credits_lbl = tk.Label(card, text="—", background=C["chat_bg"], foreground=C["fg"],
                                    font=("Segoe UI", 16, "bold"), anchor="w")
        self.credits_lbl.grid(row=0, column=0, sticky="w")
        self.credits_sub = tk.Label(card, text="credits left", background=C["chat_bg"], foreground=C["muted"],
                                    font=("Segoe UI", 9), anchor="w")
        self.credits_sub.grid(row=0, column=1, sticky="sw", padx=(8, 0), pady=(0, 3))
        self.bar = ttk.Progressbar(card, mode="determinate", maximum=100)
        self.bar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 8))
        self.spend = {}
        spend = tk.Frame(card, background=C["chat_bg"])
        spend.grid(row=2, column=0, columnspan=2, sticky="ew")
        for i, name in enumerate(("Today", "This week", "This month", "All time")):
            spend.columnconfigure(i, weight=1)
            tk.Label(spend, text=name, background=C["chat_bg"], foreground=C["muted"], font=("Segoe UI", 8),
                     anchor="w").grid(row=0, column=i, sticky="w")
            self.spend[name] = tk.Label(spend, text="—", background=C["chat_bg"], foreground=C["fg"],
                                        font=("Segoe UI", 10, "bold"), anchor="w")
            self.spend[name].grid(row=1, column=i, sticky="w")
        self.account_note = tk.Label(card, text="", background=C["chat_bg"], foreground=C["muted"],
                                     font=("Segoe UI", 8), anchor="w", justify="left", wraplength=wrap - 30)
        self.account_note.grid(row=3, column=0, columnspan=2, sticky="w", pady=(8, 0))

        # --- model ---
        heading("Model", 4)
        mrow = ttk.Frame(f)
        mrow.grid(row=5, column=0, sticky="ew")
        mrow.columnconfigure(0, weight=1)
        self.model = ttk.Combobox(mrow)
        self.model.set(settings.get("ai_model") or "")
        self.model.grid(row=0, column=0, sticky="ew")
        ttk.Button(mrow, text="Auto model", command=self.auto_model).grid(row=0, column=1, padx=(8, 0))
        self.info = muted(f, "Loading the model list…", wraplength=wrap)
        self.info.grid(row=6, column=0, sticky="w", pady=(4, 0))

        # --- usage counted on this computer ---
        heading("Usage on this computer", 7)
        table = tk.Frame(f, background=C["bg"])
        table.grid(row=8, column=0, sticky="ew")
        cols = ("", "Requests", "Input tokens", "Output tokens", "Cost")
        for i, name in enumerate(cols):
            table.columnconfigure(i, weight=1 if i else 0, minsize=int(90 * scale) if not i else 0)
            tk.Label(table, text=name, background=C["bg"], foreground=C["muted"], font=("Segoe UI", 8),
                     anchor="e" if i else "w").grid(row=0, column=i, sticky="ew")
        self.usage_cells = {}
        for r, period in enumerate(("Today", "This month", "All time"), start=1):
            tk.Label(table, text=period, background=C["bg"], foreground=C["fg"], font=("Segoe UI", 9),
                     anchor="w").grid(row=r, column=0, sticky="w")
            for i in range(1, 5):
                lab = tk.Label(table, text="", background=C["bg"], foreground=C["fg"], font=("Segoe UI", 9),
                               anchor="e")
                lab.grid(row=r, column=i, sticky="ew")
                self.usage_cells[(period, i)] = lab
        foot = tk.Frame(f, background=C["bg"])
        foot.grid(row=9, column=0, sticky="ew", pady=(2, 0))
        muted(foot, "Counted from OpenRouter's answers to this app.").pack(side="left")
        reset = tk.Label(foot, text="Reset", background=C["bg"], foreground=C["accent"], cursor="hand2",
                         font=("Segoe UI", 8, "underline"))
        reset.pack(side="right")
        reset.bind("<Button-1>", lambda e: self.reset_usage())
        self.show_usage()

        # --- privacy + buttons ---
        muted(f, "Your questions, attachments and the sent cells go to OpenRouter and the model's provider. "
                 "AI scripts only run when you press Run (or in Auto mode).", wraplength=wrap).grid(
            row=10, column=0, sticky="w", pady=(14, 0))
        ttk.Separator(f).grid(row=11, column=0, sticky="ew", pady=(12, 10))
        b = ttk.Frame(f)
        b.grid(row=12, column=0, sticky="e")
        ttk.Button(b, text="Cancel", command=w.destroy).pack(side="right")
        ttk.Button(b, text="Save", command=self.save, default="active").pack(side="right", padx=(0, 8))
        w.bind("<Escape>", lambda e: w.destroy())
        w.bind("<Return>", lambda e: self.save() if e.widget is not self.model else None)
        f.update_idletasks()
        f.grid_propagate(True)
        w.minsize(int(520 * scale), 0)
        self.all = []
        self.model.bind("<KeyRelease>", lambda e: self.filter())
        self.model.bind("<<ComboboxSelected>>", lambda e: self.show_price())
        center(w, parent)
        w.deiconify()
        w._root().after_idle(lambda: w.winfo_exists() and w.grab_set())   # modal: one settings window at a time
        self.key_entry.focus_set()
        self.start_models()
        self.start_account()

    # --- background loads (the window never freezes while OpenRouter answers) ---

    def background(self, kind, fn):
        def work():
            try:
                self.results.put((kind, fn(), None))
            except Exception as e:
                self.results.put((kind, None, str(e)))
        threading.Thread(target=work, daemon=True).start()
        self.check_later()

    def check_later(self):
        """Look for answers again in a moment (cancelled when the window closes)."""
        if getattr(self, "timer", None) is None:
            self.win.bind("<Destroy>", lambda e: e.widget is self.win and self.cancel_timer(), add="+")
        self.cancel_timer()
        self.timer = self.win._root().after(100, self.take_results)

    def cancel_timer(self):
        if getattr(self, "timer", None):
            try:
                self.win._root().after_cancel(self.timer)
            except Exception:
                pass
        self.timer = ""

    def take_results(self):
        self.timer = ""
        if not self.win.winfo_exists():
            return
        try:
            kind, value, error = self.results.get_nowait()
        except queue.Empty:
            self.check_later()
            return
        if kind == "models":
            self.models_loaded(value, error)
        else:
            self.show_account(value, error)
        if not self.results.empty():
            self.take_results()

    def start_models(self):
        transport = self.transport
        self.background("models", lambda: models(transport))

    def start_account(self):
        key, transport = self.key.get().strip(), self.transport
        self.account_note.configure(text="Asking OpenRouter…")
        self.background("account", lambda: account(key, transport))

    def open_page(self, url):
        import webbrowser
        webbrowser.open(url)
        return url

    # --- account ---

    def load_account(self):
        """The account's credits and spending, now (tests; the window uses start_account)."""
        try:
            info = account(self.key.get().strip(), self.transport)
        except AIError as e:
            return self.show_account(None, str(e))
        return self.show_account(info, None)

    def show_account(self, info, error):
        import time
        if error or not info:
            self.credits_lbl.configure(text="—")
            self.credits_sub.configure(text="credits left")
            self.bar.configure(value=0)
            for lab in self.spend.values():
                lab.configure(text="—")
            self.account_note.configure(text=error or "No answer from OpenRouter.")
            return None
        left = None
        if info["credits"] is not None:                       # the account's balance
            left = info["credits"] - (info["account_used"] or 0)
            self.credits_lbl.configure(text=money(left))
            self.credits_sub.configure(text=f"credits left of {money(info['credits'])}")
            self.bar.configure(value=0 if not info["credits"] else max(0, min(100, 100 * left / info["credits"])))
        elif info["limit"] is not None:                       # a key with its own spending limit
            left = info["limit_remaining"]
            self.credits_lbl.configure(text=money(left))
            self.credits_sub.configure(text=f"left of this key's {money(info['limit'])} limit")
            self.bar.configure(value=0 if not info["limit"] else max(0, min(100, 100 * (left or 0) / info["limit"])))
        else:
            self.credits_lbl.configure(text=money(info["used"]))
            self.credits_sub.configure(text="spent with this key (no limit set)")
            self.bar.configure(value=0)
        for name, k in (("Today", "daily"), ("This week", "weekly"), ("This month", "monthly"), ("All time", "used")):
            self.spend[name].configure(text=money(info[k]))
        notes = []
        if info["limit"] is not None and info["credits"] is not None:
            notes.append(f"This key's limit: {money(info['limit'])} ({money(info['limit_remaining'])} left).")
        if info["free_requests"]:
            used, limit = info["free_requests"]
            notes.append(f"Free models today: {used} of {limit} requests used.")
        if info["free_tier"]:
            notes.append("Free tier: add credits to use paid models.")
        notes.append(f"Spending of this key, from OpenRouter · updated {time.strftime('%H:%M')}.")
        self.account_note.configure(text="  ".join(notes))
        return info

    # --- usage counted here ---

    def show_usage(self):
        totals = usage_totals()
        for period, row in totals.items():
            for i, text in enumerate((f"{row[0]:,}", count_text(row[1]), count_text(row[2]), money(row[3])), start=1):
                self.usage_cells[(period, i)].configure(text=text)
        return totals

    def reset_usage(self):
        from tkinter import messagebox
        if messagebox.askyesno("AI usage", "Set the usage counted on this computer back to zero?\n\n"
                                           "(Your OpenRouter account isn't changed.)", parent=self.win):
            reset_usage()
            self.show_usage()
            return True
        return False

    # --- model ---

    def load_models(self):
        """The model list, now (tests; the window loads it in the background)."""
        try:
            return self.models_loaded(models(self.transport), None)
        except AIError as e:
            return self.models_loaded(None, str(e))

    def models_loaded(self, found, error):
        if error:
            self.info.configure(text=error)
            return []
        self.all = found or []
        self.filter()
        self.info.configure(text=f"{len(self.all)} models. Type to filter; prices are per million tokens (in / out).")
        if self.model.get() and any(m[0] == self.model.get() for m in self.all):
            self.show_price()
        return self.all

    def filter(self):
        text = self.model.get().lower()
        self.model["values"] = [m[0] for m in self.all if text in m[0].lower() or text in m[1].lower()][:300]

    def show_price(self):
        m = next((m for m in self.all if m[0] == self.model.get()), None)
        if m:
            size = _context_sizes.get(m[0])
            room = f" · {size // 1000:,}k context" if size else ""
            self.info.configure(text=f"{m[1]} · {m[2]} · reads {m[3]}{room}")

    def auto_model(self):
        """OpenRouter's Auto Router: it picks a suitable model for each question."""
        self.model.set("openrouter/auto")
        self.info.configure(text="Auto: OpenRouter picks a suitable model for each question (you pay that model's "
                                 "price).")

    def save(self):
        store_key(self.key.get())
        model = self.model.get().strip()
        settings.set("ai_model", model)
        if model in _context_sizes:                  # its context size, for fitting questions next time
            settings.set("ai_model_ctx", {"model": model, "tokens": _context_sizes[model]})
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
