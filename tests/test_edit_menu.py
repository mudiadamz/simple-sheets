"""Right-click edit menu on text boxes, and on the AI pane (chat, question, script code)."""
import tkinter as tk

import ai_assist as ai
import edit_menu
import screen_capture
from test_ai_assist import FakeRouter


def labels(m):
    return [m.entrycget(i, "label") for i in range(m.index("end") + 1) if m.type(i) != "separator"]


def state(m, label):
    i = next(i for i in range(m.index("end") + 1) if m.type(i) != "separator" and m.entrycget(i, "label") == label)
    return str(m.entrycget(i, "state"))


def invoke(m, label):
    i = next(i for i in range(m.index("end") + 1) if m.type(i) != "separator" and m.entrycget(i, "label") == label)
    m.invoke(i)


def test_editable_text_menu(app):
    t = tk.Text(app, undo=True)
    t.insert("1.0", "hello world")
    m = edit_menu.build(t)
    assert labels(m) == ["Undo", "Redo", "Cut", "Copy", "Paste", "Delete", "Select all"]
    assert state(m, "Copy") == "disabled" and state(m, "Cut") == "disabled"     # nothing selected
    invoke(m, "Select all")
    assert t.get("sel.first", "sel.last") == "hello world"
    m = edit_menu.build(t)
    assert state(m, "Copy") == "normal"
    invoke(m, "Copy")
    assert app.clipboard_get() == "hello world"
    invoke(m, "Delete")
    assert t.get("1.0", "end-1c") == ""
    t.destroy()


def test_read_only_text_and_entry(app):
    t = tk.Text(app)
    t.insert("1.0", "answer")
    t.configure(state="disabled")
    assert labels(edit_menu.build(t)) == ["Copy", "Select all"]
    e = tk.Entry(app)
    e.insert(0, "abc")
    m = edit_menu.build(e)
    assert labels(m) == ["Cut", "Copy", "Paste", "Delete", "Select all"]
    invoke(m, "Select all")
    assert e.selection_present()
    invoke(edit_menu.build(e), "Cut")
    assert e.get() == "" and app.clipboard_get() == "abc"
    invoke(edit_menu.build(e), "Paste")
    assert e.get() == "abc"
    t.destroy()
    e.destroy()


def test_ai_pane_has_menus(app, monkeypatch):
    a = ai.Assistant(app, FakeRouter("ok"))
    a.say("model", "See [the docs](https://example.com/docs) for more.")
    for w in (a.chat, a.prompt, a.code):
        assert w.bind("<Button-3>")                                      # right click is bound
    chat_menu = edit_menu.build(a.chat)
    assert labels(chat_menu) == ["Copy", "Select all"]                   # the chat is read-only
    invoke(chat_menu, "Select all")
    invoke(edit_menu.build(a.chat), "Copy")
    assert "See the docs for more." in app.clipboard_get()
    # a right click on a link adds Open link / Copy link
    a.chat.update_idletasks()
    a.win.update()
    tag = next(iter(a.links))
    x, y, w_, h = a.chat.bbox(f"{tag}.first")
    items = a.link_items(type("E", (), {"x": x + 2, "y": y + h // 2})())
    assert [lab for lab, _ in items] == ["Open link", "Copy link"]
    items[1][1]()
    assert app.clipboard_get() == "https://example.com/docs"
    assert a.link_items(type("E", (), {"x": 1, "y": 1})()) == []           # not on a link
    # Paste in the question box: text goes in, a picture becomes an attachment
    monkeypatch.setattr(screen_capture, "clipboard_image", lambda: None)
    app.clipboard_clear()
    app.clipboard_append("sum column C")
    assert a.paste_text() == "text" and a.question_text() == "sum column C"
    monkeypatch.setattr(screen_capture, "clipboard_image", lambda: b"\x89PNG fake")
    assert a.paste_text() == "picture" and a.attachments and a.attachments[-1].kind == "image"
    assert labels(edit_menu.build(a.code))[:2] == ["Undo", "Redo"]       # the script box is editable
    a.win.destroy()
