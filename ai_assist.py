"""Simple Sheets' side of the AI assistant (the pane itself is simplecore.ai_assist): what the model sees (the
sheet as a table with cell addresses, or the selection), how its scripts run (scripting.run, one undo step), and
the spreadsheet files it reads as attachments.

Settings (Tools → Settings… → AI, or the assistant's ⚙): your OpenRouter API key (kept encrypted for your
Windows account) and the model. What is sent: the question, the sheet's names and the selected cells or
the sheet's data (up to MAX_CELLS cells) - to OpenRouter and the model's provider.
"""
from simplecore import ai_assist as core
from simplecore.ai_assist import *  # noqa: F401,F403  (the pane, keys, models, chats… as before)
from simplecore.ai_assist import _context_sizes  # noqa: F401  (shared dict: tests set sizes here)
from simplecore.widgets import center  # noqa: F401  (older callers find it here)
from simplecore.widgets import dpapi as _dpapi  # noqa: F401

from formula import col_letters

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


def cut_sheet(text, room):
    """The sheet text cut to about `room` tokens: whole rows from the top, then '… N more rows not shown'.
    -> (text, rows kept or None when nothing was cut)."""
    tokens = core.tokens
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


workbook_rows = core.workbook_rows

SHEET_TYPES = (".xlsx", ".xlsm", ".xlsb", ".xls", ".ods", ".csv", ".tsv", ".xltx", ".xltm")


def _sheet_table(path):
    import sheet_io as io
    book = io.load(path)
    return [list(r) for r in next(iter(book.values())).rows]


class SheetsProfile(core.Profile):
    name = "sheet"
    scope_all = "sheet"
    system = SYSTEM
    web_note = WEB_NOTE
    summary_role = "a spreadsheet assistant"
    sent = "the sent cells"
    placeholder = "Ask about your data, or say what to change…"
    plus_tip = "Attach a file, capture the screen, or choose what to send (sheet / selection)"

    def context(self, app, scope):
        return context(app, scope)

    def cut(self, text, room):
        return cut_sheet(text, room)

    def too_big(self, kept):
        return (f"The sheet is too big for this model: only its first {kept} rows were sent. Select the cells that "
                "matter and use + → Send only the selection, or pick a model with a larger context (⚙).")

    def run_script(self, app, code, extra):
        import scripting
        out = []
        ok, msg = scripting.run(app, code, out=out.append, filename="<AI script>", extra=extra)
        return ok, msg, "".join(out), ok and "no cells changed" not in msg

    def open_script(self, app, code):
        import scripting
        w = scripting.open_window(app)
        w.load_text(code)
        return w

    def script_error(self, message):
        import scripting
        return scripting.ScriptError(message)

    def pane_parent(self, app):
        return app

    def main_widget(self, app):
        return app.table

    def before_open(self, app):
        app.table.commit()


Assistant = core.Assistant
core.TABLE_READERS.update({ext: _sheet_table for ext in SHEET_TYPES})
core.use(SheetsProfile())
