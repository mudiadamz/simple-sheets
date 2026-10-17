"""AI settings: account credits / spending from OpenRouter, usage counted here, a centred window."""
import time

from simplecore import settings

import ai_assist as ai
from test_ai_assist import FakeRouter, configured, wait  # noqa: F401  (configured is a fixture)

KEY_INFO = {"data": {"label": "sk-or-v1-abc...", "limit": None, "limit_remaining": None, "usage": 5.79,
                     "usage_daily": 0.12, "usage_weekly": 0.8, "usage_monthly": 2.1, "is_free_tier": False,
                     "free_model_daily_requests": {"used": 3, "limit": 50, "remaining": 47}}}
CREDITS = {"data": {"total_credits": 10, "total_usage": 5.79}}


class Account(FakeRouter):
    def __init__(self, reply="ok", key_info=KEY_INFO, credits=CREDITS):
        super().__init__(reply)
        self.key_info, self.credits = key_info, credits

    def __call__(self, method, url, key=None, body=None, timeout=90):
        if url.endswith("/key"):
            self.calls.append((method, url, key, body))
            return self.key_info
        if url.endswith("/credits"):
            self.calls.append((method, url, key, body))
            if self.credits is None:
                raise ai.AIError("OpenRouter said 403.")
            return self.credits
        out = super().__call__(method, url, key, body, timeout)
        if url.endswith("/chat/completions"):
            out = {**out, "usage": {"prompt_tokens": 1200, "completion_tokens": 300, "cost": 0.0042}}
        return out


def test_usage_counted_per_day():
    ai.reset_usage()
    ai.record_usage([{"in": 100, "out": 20, "cost": 0.01}], today="2026-10-15")
    ai.record_usage([{"in": 50, "out": 5, "cost": 0.002}, {"in": 10, "out": 1, "cost": 0}], today="2026-10-16")
    t = ai.usage_totals(today="2026-10-16")
    assert t["Today"] == [2, 60, 6, 0.002]
    assert t["This month"][:3] == [3, 160, 26] and abs(t["This month"][3] - 0.012) < 1e-9
    assert t["All time"][0] == 3 and ai.usage_totals(today="2026-11-01")["This month"] == [0, 0, 0, 0.0]
    for i in range(420):                                   # more than a year: old days fold into "older"
        ai.record_usage([{"in": 1, "out": 1, "cost": 0}], today=f"2024-{1 + i // 28:02d}-{1 + i % 28:02d}")
    u = settings.get("ai_usage")
    assert len(u["days"]) == 400 and ai.usage_totals()["All time"][0] == 423
    assert ai.record_usage([]) is None
    ai.reset_usage()
    assert ai.usage_totals()["All time"] == [0, 0, 0, 0.0]


def test_questions_record_tokens_and_cost(app, configured):  # noqa: F811
    ai.reset_usage()
    router = Account("The answer.")
    a = ai.Assistant(app, router)
    a.send("hi")
    wait(app, a)
    assert router.calls[-1][3]["usage"] == {"include": True}           # OpenRouter adds the cost to its answer
    assert ai.usage_totals()["Today"] == [1, 1200, 300, 0.0042]
    a.win.destroy()


def test_account_numbers():
    info = ai.account("sk-or-x", Account())
    assert info["credits"] == 10 and info["account_used"] == 5.79 and info["daily"] == 0.12
    assert info["free_requests"] == (3, 50)
    no_credits = ai.account("sk-or-x", Account(credits=None))          # /credits refused: the key's numbers only
    assert no_credits["credits"] is None and no_credits["used"] == 5.79
    import pytest
    with pytest.raises(ai.AIError):
        ai.account("", Account())
    assert ai.money(4.21) == "$4.21" and ai.money(0.0042) == "$0.0042" and ai.money(None) == "—"
    assert ai.count_text(1500) == "1,500" and ai.count_text(25000) == "25.0k" and ai.count_text(2_300_000) == "2.3M"


def test_settings_window_shows_account_and_usage(app):
    ai.reset_usage()
    ai.record_usage([{"in": 1200, "out": 300, "cost": 0.0042}])
    ai.store_key("sk-or-test")
    a = ai.Assistant(app, Account())
    d = a.settings()
    end = time.monotonic() + 3
    while d.credits_lbl["text"] == "—" and time.monotonic() < end:      # loads in the background
        app.update()
        time.sleep(0.02)
    assert d.credits_lbl["text"] == "$4.21" and "of $10.00" in d.credits_sub["text"]
    assert abs(float(d.bar["value"]) - 42.1) < 0.1
    assert d.spend["Today"]["text"] == "$0.12" and d.spend["All time"]["text"] == "$5.79"
    assert "Free models today: 3 of 50" in d.account_note["text"]
    assert d.usage_cells[("Today", 1)]["text"] == "1" and d.usage_cells[("Today", 2)]["text"] == "1,200"
    assert d.usage_cells[("Today", 4)]["text"] == "$0.0042"
    while not d.all and time.monotonic() < end:
        app.update()
        time.sleep(0.02)
    assert len(d.all) == 2                                               # the model list came too
    # a key with its own limit, no account credits
    d.transport = Account(key_info={"data": {**KEY_INFO["data"], "limit": 20, "limit_remaining": 15}}, credits=None)
    d.load_account()
    assert d.credits_lbl["text"] == "$15.00" and "limit" in d.credits_sub["text"]
    d.transport = FakeRouter(error="OpenRouter said 401. The API key was refused")
    d.load_account()
    assert d.credits_lbl["text"] == "—" and "refused" in d.account_note["text"]
    d.win.destroy()
    a.win.destroy()


def test_dialog_centred_on_its_window(app):
    import tkinter as tk
    parent = tk.Toplevel(app)
    parent.geometry("900x700+100+80")
    parent.update()
    win = tk.Toplevel(parent)
    tk.Label(win, text="x" * 40).pack(padx=40, pady=40)
    x, y = ai.center(win, parent)
    w = win.winfo_reqwidth()
    assert abs((x + w / 2) - (parent.winfo_rootx() + 450)) <= 2              # in the middle across
    assert parent.winfo_rooty() < y < parent.winfo_rooty() + 350             # in the upper middle
    parent.geometry("300x200+-5000+50")                                      # a parent off the screen
    parent.update()
    x, _ = ai.center(win, parent)
    assert x == 0                                                            # kept on the screen
    parent.destroy()
