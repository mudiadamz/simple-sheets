"""AI assistant: questions fitted to the model's context window (sheet cut, older chat summarised)."""
import time

import pytest

import ai_assist as ai
import settings
from test_ai_assist import FakeRouter, configured, wait  # noqa: F401  (configured is a fixture)

MODEL = "anthropic/test-model"


def ctx_with_room(extra):
    """A context size whose question budget is the system prompt plus about `extra` tokens."""
    ctx = 1000
    while ai.input_budget(ctx) < ai.tokens(ai.SYSTEM) + extra:
        ctx += 100
    return ctx


@pytest.fixture
def small_model():
    ai._context_sizes[MODEL] = ctx_with_room(1500)
    yield ai._context_sizes[MODEL]
    ai._context_sizes.pop(MODEL, None)


class Router(FakeRouter):
    """Answers questions with a long reply; summary requests with a summary."""

    def __call__(self, method, url, key=None, body=None, timeout=90):
        self.calls.append((method, url, key, body))
        last = body["messages"][-1]["content"]
        if isinstance(last, str) and last.startswith("Summarise this part"):
            return {"choices": [{"message": {"content": "SUMMARY: the user is checking sprint totals."}}]}
        return {"choices": [{"message": {"content": "A fairly long answer. " * 40}}]}


def test_context_sizes_from_the_model_list():
    ai.remember_context({"id": "x/big", "context_length": 200000})
    ai.remember_context({"id": "x/top", "top_provider": {"context_length": 8000}})
    ai.remember_context({"id": "x/none"})
    assert ai.context_size("x/big") == 200000 and ai.context_size("x/top") == 8000
    assert ai.context_size("x/none") == ai.DEFAULT_CONTEXT
    ai.remember_context({"id": "openrouter/auto", "context_length": 2000000})
    assert ai.context_size("openrouter/auto") == ai.AUTO_CONTEXT          # not the 2M it claims
    ai._context_sizes.pop("openrouter/auto")
    settings.set("ai_model_ctx", {"model": "y/saved", "tokens": 64000})
    assert ai.context_size("y/saved") == 64000                         # remembered from ⚙ settings
    assert ai.input_budget(100000) < 100000 - 3999                       # room left for the answer
    for k in ("x/big", "x/top"):
        ai._context_sizes.pop(k)


def test_big_sheet_is_cut_to_fit(app):
    app.set_cells([(r, c, f"value {r}-{c} with some text") for r in range(400) for c in range(5)])
    conv = ai.Conversation()
    msgs = conv.build(app, "sum it", "sheet", ctx=ctx_with_room(2000))
    text = msgs[-1]["content"]
    assert conv.sheet_rows and 10 < conv.sheet_rows < 400
    assert "more rows not shown (too many for this model" in text and text.endswith("Request: sum it")
    assert sum(ai.tokens(m) for m in msgs) <= ai.input_budget(ctx_with_room(2000))
    big = ai.Conversation()
    big.build(app, "sum it", "sheet", ctx=1_000_000)
    assert big.sheet_rows is None                                       # a big model gets it all


def test_history_fits_newest_first(app):
    conv = ai.Conversation()
    for i in range(30):
        conv.add(f"question {i} " + "x" * 300, f"answer {i} " + "y" * 300)
    msgs = conv.build(app, "next", "sheet", ctx=ctx_with_room(1500))
    sent = [m["content"] for m in msgs[1:-1]]
    assert sent and sent[-1].startswith("answer 29") and sent[-2].startswith("question 29")   # newest kept
    assert len(sent) < 60 and len(sent) % 2 == 0 and conv.dropped == 60 - len(sent)
    assert [m["role"] for m in msgs[1:-1]] == ["user", "assistant"] * (len(sent) // 2)
    req, upto = conv.summary_request()
    assert upto == conv.dropped and "question 0" in req[0]["content"]
    conv.set_summary("S1", upto)
    assert conv.summary_upto == upto and conv.dropped == 0
    msgs = conv.build(app, "next", "sheet", ctx=ctx_with_room(1500))
    assert "Earlier in this chat (summary):\nS1" in msgs[0]["content"]
    assert ai.Conversation().summary_request() is None


def test_long_chat_gets_summarised(app, configured, small_model):  # noqa: F811
    router = Router()
    a = ai.Assistant(app, router)
    for i in range(8):
        a.send(f"question {i} about the sprint")
        wait(app, a)
        end = time.monotonic() + 3
        while getattr(a, "summarizing", False) and time.monotonic() < end:   # the summary comes in the background
            app.update()
            time.sleep(0.02)
    summaries = [c for c in router.calls if c[3]["messages"][-1]["content"].startswith("Summarise this part")]
    assert summaries and a.conv.summary.startswith("SUMMARY:") and a.conv.summary_upto > 0
    last = [c for c in router.calls if c not in summaries][-1][3]["messages"]
    assert "Earlier in this chat (summary):\nSUMMARY" in last[0]["content"]
    chat = a.chat.get("1.0", "end")
    assert chat.count("older messages are kept as a short summary") == 1   # said once
    chat_id = a.chat_id
    a.new_chat()
    assert a.conv.summary == "" and a.open_chat(chat_id)
    assert a.conv.summary.startswith("SUMMARY:") and a.conv.summary_upto > 0   # kept with the saved chat
    a.win.destroy()


def test_too_big_error_explains_what_to_do(app, configured):  # noqa: F811
    router = FakeRouter(error="OpenRouter said 400. This endpoint's maximum context length is 8192 tokens.")
    a = ai.Assistant(app, router)
    a.send("hello")
    wait(app, a)
    chat = a.chat.get("1.0", "end")
    assert "too big for this model" in chat and "Send only the selection" in chat
    assert ai.too_big("prompt is too long: 250000 tokens > 200000 maximum")
    assert not ai.too_big("OpenRouter said 401. The API key was refused")
    a.win.destroy()


def test_settings_remember_the_models_size(app):
    ai.models(lambda *x, **k: {"data": [{"id": "z/model", "name": "Z", "context_length": 131072,
                                         "pricing": {"prompt": "0", "completion": "0"}}]})
    a = ai.Assistant(app, FakeRouter("x"))
    d = a.settings()
    d.model.set("z/model")
    d.key.set("sk-or-test")
    d.save()
    assert settings.get("ai_model_ctx") == {"model": "z/model", "tokens": 131072}
    ai._context_sizes.pop("z/model")
    assert ai.context_size("z/model") == 131072
    a.win.destroy()
