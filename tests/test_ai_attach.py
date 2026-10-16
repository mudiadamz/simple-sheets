"""AI assistant: toolbar button, attachments (files, screen capture, pasted pictures), PNG / clipboard decoding."""
import base64
import struct
import sys
import time
import tkinter as tk

import pytest

import ai_assist as ai
import screen_capture as sc
import settings
from test_ai_assist import FakeRouter, wait


@pytest.fixture(autouse=True)
def configured(request):
    if "app" in request.fixturenames:                 # after the app fixture has reset the settings folder
        request.getfixturevalue("app")
    ai.store_key("sk-or-test")
    settings.set("ai_model", "vision/model")
    yield
    settings.set("ai_key", "")
    settings.set("ai_model", "")


def png_size(data):
    assert data[:8] == b"\x89PNG\r\n\x1a\n"
    return struct.unpack(">II", data[16:24])


def test_png_writer_reads_back(app):
    rgb = bytes([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 255])     # 2 x 2: red green / blue white
    data = sc.png(2, 2, rgb)
    assert png_size(data) == (2, 2)
    img = tk.PhotoImage(master=app, data=base64.b64encode(data))
    assert img.get(0, 0) == (255, 0, 0) and img.get(1, 0) == (0, 255, 0) and img.get(0, 1) == (0, 0, 255)


def test_clipboard_bitmap_decoding(app):
    # a 2 x 2 bottom-up 24-bit DIB (rows padded to 4 bytes): bottom row first
    header = struct.pack("<IiiHHIIiiII", 40, 2, 2, 1, 24, 0, 0, 0, 0, 0, 0)
    bottom = bytes([255, 0, 0, 255, 255, 255]) + b"\0\0"                # blue, white (BGR)
    top = bytes([0, 0, 255, 0, 255, 0]) + b"\0\0"                       # red, green
    data = sc.parse_dib(header + bottom + top)
    img = tk.PhotoImage(master=app, data=base64.b64encode(data))
    assert img.get(0, 0) == (255, 0, 0) and img.get(1, 0) == (0, 255, 0) and img.get(0, 1) == (0, 0, 255)
    assert sc.parse_dib(b"short") is None
    assert sc.parse_dib(struct.pack("<IiiHHIIiiII", 40, 2, 2, 1, 8, 0, 0, 0, 0, 0, 0)) is None   # 8-bit: no


@pytest.mark.skipif(sys.platform != "win32", reason="Windows screen")
def test_grab_screen_region():
    assert png_size(sc.grab(0, 0, 30, 20)) == (30, 20)
    x, y, w, h = sc.virtual_screen()
    assert w > 0 and h > 0


def test_attach_files(tmp_path):
    pic = tmp_path / "chart.png"
    pic.write_bytes(sc.png(1, 1, b"\x00\x00\x00"))
    a = ai.attach_file(str(pic))
    assert a.kind == "image" and a.part()["image_url"]["url"].startswith("data:image/png;base64,")
    pdf = tmp_path / "report.pdf"
    pdf.write_bytes(b"%PDF-1.4 tiny")
    p = ai.attach_file(str(pdf)).part()
    assert p["type"] == "file" and p["file"]["filename"] == "report.pdf"
    txt = tmp_path / "notes.txt"
    txt.write_text("prices go up 5%", encoding="utf-8")
    assert "prices go up 5%" in ai.attach_file(str(txt)).part()["text"]
    import workbooks
    book = workbooks.plain(tmp_path / "other.xlsx")
    t = ai.attach_file(str(book))
    assert t.table and t.table[0] and "attachments['other.xlsx']" in t.part()["text"]   # rows, for scripts
    blob = tmp_path / "x.bin"
    blob.write_bytes(b"\0\1\2" * 10)
    with pytest.raises(ai.AIError, match="can't be sent"):
        ai.attach_file(str(blob))
    big = tmp_path / "big.png"
    big.write_bytes(b"\0" * (ai.MAX_FILE + 1))
    with pytest.raises(ai.AIError, match="MB"):
        ai.attach_file(str(big))


def test_attachments_sent_once_with_the_question(app, tmp_path):
    router = FakeRouter("I see a chart.")
    a = ai.Assistant(app, router)
    pic = tmp_path / "chart.png"
    pic.write_bytes(sc.png(1, 1, b"\xff\x00\x00"))
    a.attach(str(pic))
    assert len(a.attachments) == 1 and a.chips.winfo_manager()
    a.send("what's in this picture?")
    assert not a.attachments and not a.chips.winfo_manager()          # sent: the chips go away
    wait(app, a)
    content = router.calls[-1][3]["messages"][-1]["content"]
    assert content[0]["type"] == "text" and "what's in this picture?" in content[0]["text"]
    assert content[1]["type"] == "image_url"
    a.send("and now?")
    wait(app, a)
    history = router.calls[-1][3]["messages"]
    assert isinstance(history[-1]["content"], str)                   # not sent again
    assert any("[attached: chart.png]" in m["content"] for m in history if isinstance(m["content"], str))
    assert "📎 chart.png" in a.chat.get("1.0", "end")
    a.win.destroy()


def test_remove_attachment_chip(app, tmp_path):
    a = ai.Assistant(app, FakeRouter("x"))
    f = tmp_path / "a.txt"
    f.write_text("hello")
    att = a.attach(str(f))
    a.remove_attachment(att)
    assert not a.attachments and not a.chips.winfo_manager()
    a.win.destroy()


def test_picture_error_explains_model(app, tmp_path):
    a = ai.Assistant(app, FakeRouter(error="OpenRouter said 404. No endpoints found that support image input"))
    pic = tmp_path / "p.png"
    pic.write_bytes(sc.png(1, 1, b"\0\0\0"))
    a.attach(str(pic))
    a.send("look")
    wait(app, a)
    assert "may not read pictures" in a.chat.get("1.0", "end")
    a.win.destroy()


def test_capture_snip_and_paste(app, monkeypatch):
    if sys.platform != "win32":
        pytest.skip("Windows screen")
    a = ai.Assistant(app, FakeRouter("x"))
    got = []
    snip = a.capture(grab_fn=lambda x, y, w, h: got.append((w, h)) or sc.png(w, h, b"\0\0\0" * w * h))
    app.update()
    snip.press(type("E", (), {"x": 10, "y": 10})())
    snip.motion(type("E", (), {"x": 70, "y": 50})())
    snip.release(type("E", (), {"x": 70, "y": 50})())
    for _ in range(30):
        app.update()
        if a.attachments:
            break
        app.after(20)
    assert got == [(60, 40)] and a.attachments[0].name.startswith("screen-") and a.attachments[0].kind == "image"
    assert not snip.win.winfo_exists()
    monkeypatch.setattr(sc, "clipboard_image", lambda: sc.png(1, 1, b"\0\0\0"))
    assert a.paste() == "break" and a.attachments[-1].name.startswith("pasted-")
    monkeypatch.setattr(sc, "clipboard_image", lambda: None)
    assert a.paste() is None                                          # plain text: the normal paste runs
    s2 = a.capture(grab_fn=lambda *r: b"")
    s2.finish(None)                                                   # Esc: nothing attached
    assert len(a.attachments) == 2
    a.win.destroy()


def test_toolbar_button_toggles_pane(app):
    import app_menu
    import theme
    b = app.ai_btn
    assert str(b.cget("background")) == theme.C["accent"] and not b.open
    b.event_generate("<Button-1>")
    app.update()
    assert app.ai_pane.visible() and b.open and str(b.cget("relief")) == "sunken"
    b.event_generate("<Button-1>")
    app.update()
    assert not app.ai_pane.visible() and not b.open
    app.set_theme(True)
    try:
        app_menu.refresh_ai_button(app)
        assert str(b.cget("background")) == theme.DARK["accent"] and str(b.cget("foreground")) == theme.DARK["on_accent"]
    finally:
        app.set_theme(False)


def test_models_say_what_they_read():
    def router(method, url, key=None, body=None, timeout=90):
        return {"data": [{"id": "a/vision", "name": "V", "pricing": {},
                          "architecture": {"input_modalities": ["text", "image", "file"]}},
                         {"id": "b/text", "name": "T", "pricing": {}}]}
    found = {m[0]: m[3] for m in ai.models(router)}
    assert found == {"a/vision": "text, image, file", "b/text": "text"}


def test_composer_is_one_compact_box(app):
    a = ai.Assistant(app, FakeRouter("ok"))
    app.update()
    assert not hasattr(a, "note")                                    # no notice under the box any more
    m = a.plus_menu
    labels = [m.entrycget(i, "label") for i in range(m.index("end") + 1) if m.type(i) != "separator"]
    assert labels == ["Attach a file…", "Capture part of the screen", "Paste a picture (Ctrl+V)",
                      "Send the sheet", "Send only the selection"]
    m.invoke(labels.index("Send only the selection") + 1)            # (+1: the separator before the radios)
    assert a.scope.get() == "selection"
    one = int(a.composer.cget("height"))
    a.hint(False)
    a.prompt.insert("1.0", "a long question " * 40)
    a.layout_composer()
    assert int(a.composer.cget("height")) > one                      # grows with the text (up to 4 lines)
    a.prompt.insert("end", "more " * 400)
    a.layout_composer()
    from tkinter import font as tkfont
    line = tkfont.Font(root=app, font=a.prompt.cget("font")).metrics("linespace")
    four = int(a.composer.cget("height"))
    assert four <= 4 * line + 60                                     # never more than 4 lines tall
    a.prompt.insert("end", "and more " * 400)
    a.layout_composer()
    assert int(a.composer.cget("height")) == four
    a.prompt.delete("1.0", "end")
    a.prompt.insert("1.0", "hello")
    a.send_btn.event_generate("<Button-1>")                          # the round button sends
    assert a.sending
    wait(app, a)
    assert not a.sending and "ok" in a.chat.get("1.0", "end")
    a.win.destroy()


def test_privacy_notice_now_in_settings(app):
    a = ai.Assistant(app, FakeRouter())
    d = a.settings()
    texts = [w.cget("text") for w in d.win.winfo_children()[0].winfo_children() if w.winfo_class() == "TLabel"]
    assert any("go to OpenRouter" in str(t) for t in texts)
    d.win.destroy()
    a.win.destroy()


def test_composer_fits_its_text_and_plus(app):
    """The question's text is never cut off and + sits inside the rounded box (any display scaling)."""
    from tkinter import font as tkfont
    a = ai.Assistant(app, FakeRouter("x"))
    app.update()
    a.layout_composer()
    app.update()
    line = tkfont.Font(root=app, font=a.prompt.cget("font")).metrics("linespace")
    assert a.prompt.winfo_height() >= line                          # a whole line shows
    assert a.prompt.winfo_y() >= 0 and a.prompt.winfo_y() + a.prompt.winfo_height() <= a.composer.winfo_height()
    assert a.plus.winfo_x() >= 6                                      # not on the box's rounded edge
    send = a.send_btn
    assert send.winfo_x() + send.winfo_width() <= a.composer.winfo_width() - 4
    a.win.destroy()


def test_compact_header(app):
    a = ai.Assistant(app, FakeRouter("x"))
    settings.set("ai_model", "meta/muse-spark-open")
    a.refresh_label()
    assert a.model_label["text"] == "muse-spark-open ▾"
    head = a.model_label.master
    buttons = [w["text"] for w in head.winfo_children() if w.winfo_class() == "TButton"]
    assert sorted(buttons) == sorted(["✕", "⚙", "↺", "✎"])           # icons only, one row
    assert not any(w.winfo_class() == "TLabel" and "AI assistant" in str(w["text"]) for w in head.winfo_children())
    a.win.destroy()


def test_chats_saved_and_reopened(app):
    a = ai.Assistant(app, FakeRouter("The total is **7**."))
    a.send("what's the total?")
    wait(app, a)
    chats = ai.list_chats()
    assert len(chats) == 1 and chats[0]["title"] == "what's the total?"
    first = chats[0]["id"]
    a.new_chat()                                                      # the old chat stays in the history
    assert a.chat.get("1.0", "end").strip() == "" and a.conv.messages == []
    a.send("second question")
    wait(app, a)
    assert len(ai.list_chats()) == 2
    m = a.build_history_menu()                                        # (shown by ↺; built here without showing)
    labels = [m.entrycget(i, "label") for i in range(m.index("end") + 1) if m.type(i) == "command"]
    assert any("what's the total?" in x for x in labels) and labels[-1] == "Clear history…"
    assert a.open_chat(first)                                         # reopen: shown again, model remembers it
    text = a.chat.get("1.0", "end")
    assert "what's the total?" in text and "The total is 7." in text and "second question" not in text
    assert a.conv.messages[0]["content"] == "what's the total?"
    a.send("and the average?")
    wait(app, a)
    sent = a.transport.calls[-1][3]["messages"]
    assert any(x.get("content") == "what's the total?" for x in sent)   # the reopened chat goes on from there
    assert len(ai.list_chats()) == 2                                  # continued, not a new one
    import tkinter.messagebox as mb
    orig = mb.askyesno
    mb.askyesno = lambda *x, **k: True
    try:
        assert a.clear_history() and ai.list_chats() == []
    finally:
        mb.askyesno = orig
    a.win.destroy()


SCRIPT_REPLY = ("Here you go.\n```python\n# Professional formatting for the sprint sheet\nsheet.freeze(1)\n"
                "sheet.format('A1:C1', bold=True)\nsheet['A1'] = 'Task'\n```")


def test_script_card(app):
    a = ai.Assistant(app, FakeRouter(SCRIPT_REPLY))
    a.send("format it")
    wait(app, a)
    app.update()
    assert a.script_box.winfo_manager()
    assert a.card_title["text"] == "Script ready · 4 lines"
    assert a.card_desc["text"] == "Professional formatting for the sprint sheet"     # its first comment
    assert not a.code.winfo_manager()                                 # the code stays folded away…
    a.toggle_code()
    assert a.code.winfo_manager() and a.code_toggle["text"] == "Code ▴"            # …until asked for
    assert app.table.sheet.freeze == (0, 0)                           # Ask mode: nothing ran yet
    assert a.run() and app.table.sheet.freeze == (1, 0) and not a.script_box.winfo_manager()
    assert "Ctrl+Z undoes it" in a.chat.get("1.0", "end")
    a.win.destroy()


def test_auto_mode_runs_scripts_straight_away(app, monkeypatch):
    import tkinter.messagebox as mb
    a = ai.Assistant(app, FakeRouter(SCRIPT_REPLY))
    assert a.mode_chip["text"] == "Ask"
    monkeypatch.setattr(mb, "askyesno", lambda *x, **k: False)
    assert a.toggle_mode() is False and a.mode_chip["text"] == "Ask"   # said no: stays Ask
    monkeypatch.setattr(mb, "askyesno", lambda *x, **k: True)
    assert a.toggle_mode() is True and a.mode_chip["text"] == "Auto" and settings.get("ai_auto_run") is True
    a.send("format it")
    wait(app, a)
    assert app.table.sheet.freeze == (1, 0) and app.table.text(0, 0) == "Task"     # ran without a click
    assert "Ran automatically" in a.chat.get("1.0", "end") and not a.script_box.winfo_manager()
    app.undo()                                                        # still one undo step
    assert app.table.sheet.freeze == (0, 0) and app.table.text(0, 0) == ""
    a.transport.reply = "```python\nsheet['A1'] = 'x'\n1/0\n```"
    a.send("break it")
    wait(app, a)
    assert app.table.text(0, 0) == "" and a.script_box.winfo_manager() and a.code.winfo_manager()  # kept to look at
    assert a.toggle_mode() is False and settings.get("ai_auto_run") is False      # back to Ask (no question)
    a.win.destroy()


def test_auto_model_choice(app):
    a = ai.Assistant(app, FakeRouter())
    d = a.settings()
    d.auto_model()
    assert d.model.get() == "openrouter/auto" and "OpenRouter picks" in d.info["text"]
    d.win.destroy()
    a.win.destroy()
    found = {m[0]: m[2] for m in ai.models(lambda *x, **k: {"data": [
        {"id": "openrouter/auto", "name": "Auto Router", "pricing": {"prompt": "-1", "completion": "-1"}}]})}
    assert found["openrouter/auto"] == "price of the model it picks"


class SlowRouter(FakeRouter):
    def __call__(self, *a, **k):
        time.sleep(self.delay)
        return super().__call__(*a, **k)


def csv_file(tmp_path, n=100):
    p = tmp_path / "INF Sprint 24.csv"
    p.write_text("Key,Summary,Status\n" + "".join(f"ISP-{i},Task {i},To Do\n" for i in range(1, n + 1)),
                 encoding="utf-8")
    return p


def test_attached_data_goes_to_scripts_not_retyped(app, tmp_path):
    reply = ("Replacing the sheet with the file.\n```python\nsheet.clear()\n"
             "sheet['A1'] = attachments['INF Sprint 24.csv']\n```")
    router = FakeRouter(reply)
    a = ai.Assistant(app, router)
    app.set_cells([(0, 0, "old"), (5, 3, "old too")])
    a.attach(str(csv_file(tmp_path)))
    a.send("replace with this data")
    wait(app, a)
    part = router.calls[-1][3]["messages"][-1]["content"][1]["text"]
    assert "101 rows x 3 columns" in part and "row 30:" in part and "row 31:" not in part   # a preview only
    assert "attachments['INF Sprint 24.csv']" in part
    assert a.run()
    g = app.table
    assert g.text(0, 0) == "Key" and g.text(100, 1) == "Task 100" and g.text(5, 3) == ""      # all 101 rows in
    router.reply = "```python\nsheet['D1'] = len(attachments['INF Sprint 24.csv'])\n```"
    a.send("how many rows was that file?")                            # later question: still usable
    wait(app, a)
    assert a.run() and g.text(0, 3) == "101"
    router.reply = "```python\nsheet['E1'] = len(attachments['nope.csv'])\n```"
    a.send("x")
    wait(app, a)
    assert not a.run() and "No attached file 'nope.csv'" in a.chat.get("1.0", "end")
    a.new_chat()
    assert a.files == {}                                               # a new chat starts without them
    a.win.destroy()


def test_working_line_and_stop(app, tmp_path):
    router = SlowRouter("late answer")
    router.delay = 1.2
    a = ai.Assistant(app, router)
    a.attach(str(csv_file(tmp_path, 3)))
    a.send("replace with this data")
    for _ in range(30):
        app.update()
        time.sleep(0.03)
    chat = a.chat.get("1.0", "end")
    assert "Reading INF Sprint 24.csv…" in chat and "●" in chat           # says what it is doing, animated
    assert any(a.send_btn.type(i) == "rectangle" for i in a.send_btn.find_all())   # ■ stop button
    assert a.stop() and not a.busy
    time.sleep(1.4)
    for _ in range(20):
        app.update()
        time.sleep(0.02)
    chat = a.chat.get("1.0", "end")
    assert "Stopped." in chat and "late answer" not in chat and "Reading" not in chat   # the late answer is ignored
    assert not any(a.send_btn.type(i) == "rectangle" for i in a.send_btn.find_all())
    router.delay = 0
    router.reply = "fresh answer"
    a.send("again")
    wait(app, a)
    assert "fresh answer" in a.chat.get("1.0", "end")
    a.win.destroy()
