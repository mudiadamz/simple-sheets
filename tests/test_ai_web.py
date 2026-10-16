"""AI assistant: web search (OpenRouter's web plugin), sources as clickable links."""
import ai_assist as ai
import settings
import theme
from test_ai_assist import FakeRouter, configured, wait  # noqa: F401  (configured is a fixture)

CITED = {"choices": [{"message": {"content": "Bitcoin is about 65,000 USD today.", "annotations": [
    {"type": "url_citation", "url_citation": {"url": "https://example.com/btc", "title": "BTC price"}},
    {"type": "url_citation", "url_citation": {"url": "https://example.com/btc", "title": "dup"}},
    {"type": "url_citation", "url_citation": {"url": "file:///C:/secret.txt", "title": "not a web page"}}]}}]}


class WebRouter(FakeRouter):
    def __call__(self, method, url, key=None, body=None, timeout=90):
        self.calls.append((method, url, key, body))
        return CITED


def test_ask_with_web_adds_plugin_and_sources():
    r = WebRouter()
    text = ai.ask([{"role": "system", "content": "sys"}, {"role": "user", "content": "btc?"}], "m/x", "k", r, web=True)
    body = r.calls[0][3]
    assert body["plugins"] == [{"id": "web", "max_results": ai.WEB_RESULTS}]
    assert ai.WEB_NOTE.strip() in body["messages"][0]["content"]
    assert "[BTC price](https://example.com/btc)" in text and text.count("example.com/btc") == 1
    assert "secret" not in text                                     # only http(s) pages count as sources


def test_ask_without_web_sends_no_plugin():
    r = WebRouter()
    ai.ask([{"role": "system", "content": "sys"}, {"role": "user", "content": "x"}], "m/x", "k", r)
    body = r.calls[0][3]
    assert "plugins" not in body and body["messages"][0]["content"] == "sys"


def test_web_chip_switches_and_is_remembered(app, configured):  # noqa: F811
    router = WebRouter()
    a = ai.Assistant(app, router)
    assert not a.web() and a.web_chip.cget("foreground") == theme.C["icon"]
    assert a.toggle_web() is True                                       # what a click on the chip does
    assert a.web() and settings.get("ai_web") is True and a.web_var.get()
    assert a.web_chip.cget("foreground") == theme.C["accent"]           # lit while on
    a.send("bitcoin price into A1")
    wait(app, a)
    assert router.calls[-1][3]["plugins"][0]["id"] == "web"
    chat = a.chat.get("1.0", "end")
    assert "Sources" in chat and "BTC price" in chat and "https://example.com/btc" not in chat   # shown as a link
    tag = next(t for t, u in a.links.items() if u == "https://example.com/btc")
    assert "link" in a.chat.tag_names(f"{tag}.first")
    opened = []
    import webbrowser
    real = webbrowser.open
    webbrowser.open = opened.append
    try:
        a.open_link("https://example.com/btc")
        a.open_link("file:///C:/Windows/notepad.exe")                    # never anything but a web page
    finally:
        webbrowser.open = real
    assert opened == ["https://example.com/btc"]
    a.toggle_web(False)
    assert not a.web() and not a.web_var.get()
    a.win.destroy()


def test_working_line_says_searching(app, configured):  # noqa: F811
    settings.set("ai_web", True)
    a = ai.Assistant(app, WebRouter())
    seen = []
    real = a.thinking
    a.thinking = lambda what="": (seen.append(what), real(what))[1]
    a.send("x")
    wait(app, a)
    assert seen == ["Searching the web"]
    a.win.destroy()
