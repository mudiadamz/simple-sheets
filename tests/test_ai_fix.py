"""AI assistant: a failing script can go back to the AI with its error (↻ Fix it; Auto mode by itself)."""
from simplecore import settings, theme

import ai_assist as ai
from test_ai_assist import FakeRouter, configured, wait  # noqa: F401  (configured is a fixture)

BAD = "```python\n# Put a total in B1\nsheet['B1'] = totl\n```"
GOOD = "```python\n# Put a total in B1\nsheet['B1'] = 42\n```"


class Replies(FakeRouter):
    """Answers with each reply in turn (the last one repeats)."""

    def __init__(self, *replies):
        super().__init__()
        self.replies = list(replies)

    def __call__(self, method, url, key=None, body=None, timeout=90):
        self.calls.append((method, url, key, body))
        reply = self.replies.pop(0) if len(self.replies) > 1 else self.replies[0]
        return {"choices": [{"message": {"content": reply}}]}


def test_fix_it_sends_error_and_script(app, configured):  # noqa: F811
    router = Replies(BAD, GOOD)
    a = ai.Assistant(app, router)
    a.send("total in B1")
    wait(app, a)
    assert not a.run()                                              # NameError: nothing changed
    assert a.failed and "NameError (line 2)" in a.failed[1]
    assert a.fix_btn.winfo_manager() and a.card_title.cget("text") == "Script failed"
    assert a.card_title.cget("foreground") == theme.C["danger"] and "totl" in a.card_desc.cget("text")
    a.fix()
    wait(app, a)
    sent = router.calls[-1][3]["messages"][-1]["content"]
    sent = sent if isinstance(sent, str) else " ".join(p.get("text", "") for p in sent)
    assert "NameError (line 2)" in sent and "sheet['B1'] = totl" in sent and "corrected script" in sent
    chat = a.chat.get("1.0", "end")
    assert "↻ Fix the script: NameError (line 2)" in chat and "sheet['B1'] = totl" not in chat   # short in the chat
    assert a.failed is None and not a.fix_btn.winfo_manager()       # the new script's card
    assert a.card_title.cget("text").startswith("Script ready")
    assert a.run() and app.table.text(0, 1) == "42"
    a.win.destroy()


def test_auto_mode_fixes_by_itself(app, configured):  # noqa: F811
    settings.set("ai_auto_run", True)
    router = Replies(BAD, GOOD)
    a = ai.Assistant(app, router)
    a.send("total in B1")
    for _ in range(3):
        wait(app, a)
    assert app.table.text(0, 1) == "42"                              # failed once, fixed, ran
    assert len(router.calls) == 2 and "try 1 of 2" in a.chat.get("1.0", "end")
    a.win.destroy()


def test_auto_mode_gives_up_after_two_tries(app, configured):  # noqa: F811
    settings.set("ai_auto_run", True)
    router = Replies(BAD)
    a = ai.Assistant(app, router)
    a.send("total in B1")
    for _ in range(6):
        wait(app, a)
    assert len(router.calls) == 1 + ai.AUTO_FIXES                    # the question + two fix tries, then stop
    assert a.failed and a.fix_btn.winfo_manager() and a.code.winfo_manager()   # left to the user
    a.fix()                                                         # a click still asks once more
    wait(app, a)
    assert len(router.calls) == 2 + ai.AUTO_FIXES
    a.win.destroy()


def test_stopped_script_offers_no_fix(app, configured):  # noqa: F811
    a = ai.Assistant(app, Replies(GOOD))
    a.set_code("x = 1")
    import scripting
    real = scripting.run
    scripting.run = lambda *x, **k: (False, "Stopped: nothing was changed.")
    try:
        assert not a.run()
    finally:
        scripting.run = real
    assert a.failed is None and not a.fix_btn.winfo_manager()
    a.win.destroy()
