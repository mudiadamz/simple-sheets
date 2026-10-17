"""AI assistant (OpenRouter): key storage, context sent, replies, scripts run only on Run. No network."""
import sys
import time

import pytest
from simplecore import settings

import ai_assist as ai


class FakeRouter:
    """Stands in for OpenRouter: records requests, answers with `reply` (or raises)."""

    def __init__(self, reply="", error=None):
        self.reply, self.error, self.calls = reply, error, []

    def __call__(self, method, url, key=None, body=None, timeout=90):
        self.calls.append((method, url, key, body))
        if self.error:
            raise ai.AIError(self.error)
        if url.endswith("/models"):
            return {"data": [{"id": "anthropic/test-model", "name": "Test", "pricing": {"prompt": "0.000003",
                                                                                         "completion": "0.000015"}},
                             {"id": "free/model", "name": "Free one", "pricing": {"prompt": "0", "completion": "0"}}]}
        return {"choices": [{"message": {"content": self.reply}}]}


def wait(app, a, seconds=5):
    end = time.monotonic() + seconds
    while a.busy and time.monotonic() < end:
        app.update()
        time.sleep(0.02)
    app.update()


@pytest.fixture
def configured():
    ai.store_key("sk-or-test-123")
    settings.set("ai_model", "anthropic/test-model")
    yield
    settings.set("ai_key", "")
    settings.set("ai_model", "")


def test_key_is_not_stored_as_plain_text(configured):
    stored = settings.get("ai_key")
    assert "sk-or-test-123" not in stored and ai.load_key() == "sk-or-test-123"
    if sys.platform == "win32":
        assert stored.startswith("dpapi:")
    ai.store_key("")
    assert ai.load_key() == ""


def test_context_sent(app):
    g = app.table
    app.set_cells([(0, 0, "Item"), (0, 1, "Qty"), (1, 0, "Pens"), (1, 1, "3"), (2, 1, "=B2*2")])
    g.select(1, 0, 1, 1)
    text = ai.context(app, "sheet")
    assert "Selection: A2:B2" in text and "A1: Item | B1: Qty" in text and "B3: =B2*2 → 6" in text
    sel = ai.context(app, "selection")
    assert "A2: Pens | B2: 3" in sel and "Item" not in sel
    old = ai.MAX_CELLS
    ai.MAX_CELLS = 4
    try:
        app.set_cells([(r, 0, str(r)) for r in range(3, 20)])
        assert "more rows not shown" in ai.context(app, "sheet")
    finally:
        ai.MAX_CELLS = old


def test_extract_code():
    code, words = ai.extract_code("Sure, this adds a total.\n```python\nsheet['A9'] = '=SUM(A1:A8)'\n```\nDone.")
    assert code == "sheet['A9'] = '=SUM(A1:A8)'\n" and words == "Sure, this adds a total.\n\nDone."
    assert ai.extract_code("The average is 4.") == (None, "The average is 4.")


def test_models_list():
    found = ai.models(FakeRouter())
    assert found[0][0] == "anthropic/test-model" and "per M tokens" in found[0][2] and found[1][2] == "free"


def test_ask_sends_key_model_and_messages(configured):
    r = FakeRouter("hi")
    assert ai.ask([{"role": "user", "content": "x"}], "m/1", "k", r) == "hi"
    method, url, key, body = r.calls[0]
    assert url.endswith("/chat/completions") and key == "k" and body["model"] == "m/1"
    with pytest.raises(ai.AIError):
        ai.ask([], "", "k", r)
    with pytest.raises(ai.AIError):
        ai.ask([], "m", "", r)


def test_question_answer_without_changes(app, configured):
    router = FakeRouter("The total quantity is 7.")
    a = ai.Assistant(app, router)
    app.set_cells([(0, 0, "3"), (1, 0, "4")])
    msgs = a.send("What's the total?")
    wait(app, a)
    assert msgs[0]["role"] == "system" and "A1: 3" in msgs[-1]["content"]
    chat = a.chat.get("1.0", "end")
    assert "The total quantity is 7." in chat and str(a.run_btn["state"]) == "disabled"
    a.win.destroy()


def test_script_shown_then_run_on_click(app, configured):
    router = FakeRouter("This writes a total.\n```python\nsheet['A3'] = '=SUM(A1:A2)'\n"
                        "sheet.format('A3', bold=True)\n```")
    a = ai.Assistant(app, router)
    app.set_cells([(0, 0, "3"), (1, 0, "4")])
    a.send("Add a total under the numbers")
    wait(app, a)
    assert "SUM(A1:A2)" in a.code.get("1.0", "end") and app.table.text(2, 0) == ""   # not run by itself
    assert a.run() is True
    assert app.table.value(2, 0) == "7" and app.table.sheet.styles[(2, 0)]["bold"]
    assert "✔" in a.chat.get("1.0", "end")
    app.undo()                                                       # one undo step
    assert app.table.text(2, 0) == ""
    a.send("and bold it")                                            # the follow-up knows what happened
    wait(app, a)
    sent = router.calls[-1][3]["messages"]
    assert any("The script was run" in m["content"] for m in sent) and any("Add a total" in m["content"] for m in sent)
    a.win.destroy()


def test_failing_script_changes_nothing(app, configured):
    a = ai.Assistant(app, FakeRouter("```python\nsheet['A1'] = 'x'\nboom()\n```"))
    a.send("do it")
    wait(app, a)
    assert a.run() is False and app.table.text(0, 0) == "" and "Nothing was changed" in a.chat.get("1.0", "end")
    a.win.destroy()


def test_errors_shown(app, configured):
    a = ai.Assistant(app, FakeRouter(error="OpenRouter said 401. The API key was refused"))
    a.send("hello")
    wait(app, a)
    assert "401" in a.chat.get("1.0", "end") and not a.busy
    a.win.destroy()


def test_settings_dialog(app):
    router = FakeRouter()
    a = ai.Assistant(app, router)
    d = a.settings()
    d.key.set("sk-new")
    assert len(d.load_models()) == 2
    d.model.set("free")
    d.filter()
    assert list(d.model["values"]) == ["free/model"]
    d.model.set("free/model")
    d.show_price()
    assert "free" in d.info["text"]
    assert d.save()
    assert ai.load_key() == "sk-new" and settings.get("ai_model") == "free/model"
    assert a.title_tip.text.startswith("Model: free/model")        # the title's tooltip names the model
    a.win.destroy()
    settings.set("ai_key", "")
    settings.set("ai_model", "")


def test_http_errors_explained(monkeypatch):
    import io
    import urllib.error

    def boom(req, timeout=0):
        raise urllib.error.HTTPError(req.full_url, 402, "Payment", {}, io.BytesIO(b'{"error": {"message": "no"}}'))
    monkeypatch.setattr(ai.urllib.request, "urlopen", boom)
    with pytest.raises(ai.AIError, match="no credits"):
        ai.real_http("POST", "https://openrouter.ai/api/v1/chat/completions", "k", {})   # (conftest blocks ai.http)


def test_pane_docked_right_and_toggles(app, configured):
    a = ai.open_assistant(app, FakeRouter("ok"))
    app.update()
    assert a.win.master is app and a.visible()                      # in the main window, not a pop-up
    assert a.win.winfo_rootx() > app.table.winfo_rootx()            # to the right of the grid
    assert ai.open_assistant(app) is a and not a.visible()          # Ctrl+Shift+A again: hidden
    assert ai.open_assistant(app) is a and a.visible()              # and back, same conversation
    a.hide()
    assert not a.visible()


def test_chat_looks_like_a_chat(app, configured):
    from simplecore import theme
    reply = "Your sheet has **32 providers**:\n\n* **Columns:** No, Code\n* `true` in 30 rows\n# Summary\nAll good."
    a = ai.Assistant(app, FakeRouter(reply))
    a.send("read data")
    wait(app, a)
    text = a.chat.get("1.0", "end")
    assert "**" not in text and "* " not in text and "• Columns: No, Code" in text and "Thinking" not in text
    assert a.chat.tag_ranges("bold") and a.chat.tag_ranges("code") and a.chat.tag_ranges("h")
    assert str(a.chat.cget("background")) == theme.C["chat_bg"] != str(a.prompt.cget("background"))
    assert a.chat.tag_ranges("you")                                  # your question in its own bubble
    assert not a.script_box.winfo_ismapped()                         # no script: no script box
    a.win.destroy()


def test_script_box_only_when_needed_and_enter_sends(app, configured):
    a = ai.Assistant(app, FakeRouter("```python\nsheet['A1'] = 1\n```"))
    app.update()
    a.prompt.focus_force()
    app.update()
    a.prompt.insert("end", "set A1")
    a.prompt.event_generate("<Return>")
    wait(app, a)
    app.update()
    assert a.script_box.winfo_ismapped() and "sheet['A1'] = 1" in a.code.get("1.0", "end")
    assert a.run() and not a.script_box.winfo_ismapped() and app.table.text(0, 0) == "1"
    a.win.destroy()


def test_hint_in_empty_box(app, configured):
    a = ai.Assistant(app, FakeRouter("x"))
    a.hint(True)
    assert a.question_text() == "" and "Ask about your data" in a.prompt.get("1.0", "end")
    assert a.send() is None                                         # the hint is never sent
    a.hint(False)
    assert a.prompt.get("1.0", "end").strip() == ""
    a.win.destroy()


def test_pane_follows_theme(app, configured):
    from simplecore import theme
    a = ai.Assistant(app, FakeRouter("x"))
    app.set_theme(True)
    try:
        assert str(a.chat.cget("background")) == theme.DARK["chat_bg"]
        assert str(a.prompt.cget("background")) == theme.DARK["field_bg"]
    finally:
        app.set_theme(False)
    assert str(a.chat.cget("background")) == theme.LIGHT["chat_bg"]
    a.win.destroy()


def test_markdown_lines():
    assert ai.markdown_lines("# Title\n- one\n  - two\n```\ncode\n```\nplain") == [
        ("Title", "h"), ("• one", "bullet"), ("  • two", "bullet"), ("code", "code"), ("plain", "p")]


def test_only_one_settings_window(app):
    """No key yet: the pane opens the settings by itself; ⚙, clicking the model and Ctrl+Shift+A again
    must not stack more windows."""
    settings.set("ai_key", "")
    settings.set("ai_model", "")
    a = ai.open_assistant(app, FakeRouter())
    for _ in range(20):                                         # the automatic one appears
        app.update()
        if getattr(a, "settings_dialog", None):
            break
        time.sleep(0.02)
    first = a.settings()
    assert a.settings() is first                                  # ⚙ again
    a.settings()                                                  # ⚙ once more
    ai.open_assistant(app)
    ai.open_assistant(app)                                        # hide / show the pane
    app.update()
    dialogs = [w for w in app.winfo_children() if w.winfo_class() == "Toplevel" and w.winfo_exists()
               and w.title() == "AI settings"]
    assert dialogs == [first.win]
    first.win.destroy()
    assert a.settings() is not first and a.settings_dialog.win.winfo_exists()   # closed: a new one opens
    a.settings_dialog.win.destroy()
    a.win.destroy()


def test_key_survives_an_update_and_uninstall(tmp_path, monkeypatch):
    """Updating = the setup removes the old program folder and installs the new one. The key lives in the
    settings folder (outside it) and is encrypted for the Windows user, not for one version."""
    if sys.platform != "win32":
        pytest.skip("Windows installer")
    from simplecore import installer
    monkeypatch.setenv("SIMPLE_SHEETS_HOME", str(tmp_path / "AppData-SimpleSheets"))
    ai.store_key("sk-or-keep-me")
    settings.set("ai_model", "anthropic/test-model")
    before = (tmp_path / "AppData-SimpleSheets" / "settings.json").read_bytes()
    payload = tmp_path / "payload.exe"
    payload.write_bytes(b"MZ v1")
    monkeypatch.setattr(installer, "payload", lambda: str(payload))
    monkeypatch.setattr(installer, "_notify_explorer", lambda: None)
    base = r"Software\SimpleSheetsTest"
    folder = tmp_path / "Programs" / "SimpleSheets"
    try:
        installer.install(str(folder), base=base, shortcuts=False)          # version 1
        payload.write_bytes(b"MZ v2")
        installer.install(str(folder), base=base, shortcuts=False)          # the update: old one removed first
        assert (folder / installer.EXE_NAME).read_bytes() == b"MZ v2"
        assert (tmp_path / "AppData-SimpleSheets" / "settings.json").read_bytes() == before
        assert ai.load_key() == "sk-or-keep-me" and settings.get("ai_model") == "anthropic/test-model"
        installer.uninstall(str(folder), base=base, shortcuts=False)        # even uninstalling keeps it
        assert ai.load_key() == "sk-or-keep-me"
    finally:
        installer._delete_tree("", base)
        settings.set("ai_key", "")
        settings.set("ai_model", "")
