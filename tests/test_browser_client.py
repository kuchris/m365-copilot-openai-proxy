from __future__ import annotations

import asyncio
import json
import os
from unittest.mock import AsyncMock

import httpx
import websockets

import pytest
from fastapi.testclient import TestClient

from m365_copilot_openai_proxy.app import create_app
from m365_copilot_openai_proxy import browser_client
from m365_copilot_openai_proxy.browser_client import (
    BrowserCopilotClient,
    _CDPPage,
    _check_page_blocked,
    _completed_reply,
    _conversation_url,
    _SNAPSHOT_JS,
    _load_chat,
    _enter_prompt,
    _FOCUS_JS,
    _SUBMIT_JS,
)
from m365_copilot_openai_proxy.config import Settings
from m365_copilot_openai_proxy.session_store import PersistentSession
from m365_copilot_openai_proxy.substrate_client import SubstrateCopilotError


@pytest.mark.parametrize("busy,reply", [
    (False, {"id": "old", "text": "Previous answer", "complete": True}),
    (True, {"id": "new", "text": "Still generating", "complete": True}),
    (False, {"id": "new", "text": "Partial answer", "complete": False}),
    (False, {"id": "new", "text": "  ", "complete": True}),
])
def test_only_a_new_finished_reply_can_complete_a_request(busy, reply):
    assert _completed_reply({"busy": busy, "replies": [reply]}, {"old"}) is None


def test_completed_reply_does_not_include_previous_messages():
    snapshot = {"busy": False, "replies": [
        {"id": "old", "text": "Previous answer", "complete": True},
        {"id": "new", "text": "New answer\nwith a second line", "complete": True},
    ]}
    assert _completed_reply(snapshot, {"old"}) == "New answer\nwith a second line"


def test_older_hydrated_message_is_not_returned_while_new_reply_is_pending():
    snapshot = {"replies": [
        {"id": "older-history", "text": "Old answer", "complete": True},
        {"id": "new", "text": "Partial", "complete": False},
    ]}
    assert _completed_reply(snapshot, set()) is None


def test_security_checks_require_human_action():
    with pytest.raises(SubstrateCopilotError, match="Complete it yourself in Edge"):
        _check_page_blocked({"verification": True, "dialog_open": True})


def test_transient_security_check_can_finish_without_resubmitting(monkeypatch):
    page = type("Page", (), {"evaluate": AsyncMock(side_effect=[
        {"verification": True, "dialog_open": True},
        {"verification": False, "dialog_open": False, "ready": True},
    ])})()
    monkeypatch.setattr(browser_client.asyncio, "sleep", AsyncMock())
    assert asyncio.run(browser_client._read_snapshot(page))["ready"]
    assert page.evaluate.await_count == 2


def test_persistent_security_check_still_requires_human(monkeypatch):
    page = type("Page", (), {"evaluate": AsyncMock(return_value={"verification": True})})()
    monkeypatch.setattr(browser_client.asyncio, "sleep", AsyncMock())
    with pytest.raises(SubstrateCopilotError, match="Complete it yourself"):
        asyncio.run(browser_client._read_snapshot(page))


def test_opt_in_verification_attempt_is_bounded_and_does_not_resubmit(monkeypatch):
    class Page:
        auto_verify = True
        verification_attempted = False
        evaluate = AsyncMock(return_value={"verification": True})
    page = Page()
    click = AsyncMock(return_value=True)
    monkeypatch.setattr(browser_client, "_try_verification_checkbox", click)
    monkeypatch.setattr(browser_client.asyncio, "sleep", AsyncMock())
    with pytest.raises(SubstrateCopilotError, match="Complete it yourself"):
        asyncio.run(browser_client._read_snapshot(page))
    assert click.await_count == 1
    assert page.verification_attempted


def test_auto_verification_only_clicks_a_recognized_checkbox_in_its_own_frame():
    class Page:
        verification_attempted = False
        inputs = []
        async def call(self, method, params=None, *, session_id=None):
            if method == "Target.getTargets":
                return {"targetInfos": [
                    {"type": "iframe", "url": "https://example.com/", "targetId": "unrelated"},
                    {"type": "iframe", "url": "https://challenges.cloudflare.com/", "targetId": "verification"},
                ]}
            if method == "DOM.getFrameOwner":
                assert params["frameId"] == "verification"
                return {"backendNodeId": 1}
            if method == "Target.attachToTarget":
                return {"sessionId": "frame-session"}
            if method == "DOM.resolveNode":
                return {"object": {"objectId": "visible-frame"}}
            if method == "Runtime.callFunctionOn":
                return {"result": {"value": True}}
            if method == "Runtime.releaseObject":
                return {}
            if method == "Accessibility.getFullAXTree":
                assert session_id == "frame-session"
                return {"nodes": [{"role": {"value": "checkbox"}, "name": {"value": "Verify you are human"}, "backendDOMNodeId": 2}]}
            if method == "DOM.getBoxModel":
                return {"model": {"content": [5, 5, 25, 5, 25, 25, 5, 25] if session_id else [100, 200, 400, 200, 400, 270, 100, 270], "width": 20, "height": 20}}
            if method == "Input.dispatchMouseEvent":
                self.inputs.append(params)
                return {}
            assert method == "Target.detachFromTarget"
            return {}
    page = Page()
    assert asyncio.run(browser_client._try_verification_checkbox(page))
    assert [(event["type"], event["x"], event["y"]) for event in page.inputs] == [
        ("mousePressed", 115, 215), ("mouseReleased", 115, 215),
    ]


def test_auto_verification_refuses_frames_from_other_tabs():
    class Page:
        verification_attempted = False
        async def call(self, method, params=None, **kwargs):
            if method == "Target.getTargets":
                return {"targetInfos": [{"type": "iframe", "url": "https://challenges.cloudflare.com/", "targetId": "other-tab"}]}
            assert method == "DOM.getFrameOwner"
            raise SubstrateCopilotError("Frame is not in this page")
    assert asyncio.run(browser_client._try_verification_checkbox(Page())) is False


def test_unexpected_dialogs_are_not_submitted_by_pressing_enter():
    with pytest.raises(SubstrateCopilotError, match="dialog is blocking"):
        _check_page_blocked({"dialog_open": True})


@pytest.mark.parametrize("url", [
    "https://login.microsoftonline.com/chat/conversation/test",
    "https://m365.cloud.microsoft.example/chat/conversation/test",
    "http://m365.cloud.microsoft/chat/conversation/test",
    "https://m365.cloud.microsoft/chat",
])
def test_persistent_session_does_not_save_an_unrelated_page(url):
    assert not _conversation_url(url)
    assert _conversation_url("https://m365.cloud.microsoft/chat/conversation/test")


def test_browser_requests_are_serialized(monkeypatch):
    async def run():
        client = BrowserCopilotClient()
        active = 0
        maximum_active = 0
        order = []

        async def turn(text, session):
            nonlocal active, maximum_active
            active += 1
            maximum_active = max(maximum_active, active)
            order.append(text)
            await asyncio.sleep(0.01)
            active -= 1
            return text.upper()

        monkeypatch.setattr(client, "_chat_turn", turn)
        replies = await asyncio.gather(client.chat("one", []), client.chat("two", []))
        assert replies == ["ONE", "TWO"]
        assert order == ["one", "two"]
        assert maximum_active == 1

    asyncio.run(run())


def test_timeout_does_not_resend_and_releases_the_queue(monkeypatch):
    async def run():
        client = BrowserCopilotClient(timeout_seconds=0.02)
        calls = []

        async def turn(text, session):
            calls.append(text)
            if text == "slow":
                await asyncio.sleep(10)
            return "ok"

        monkeypatch.setattr(client, "_chat_turn", turn)
        with pytest.raises(SubstrateCopilotError, match="not automatically resubmitted"):
            await client.chat("slow", [])
        assert await client.chat("next", []) == "ok"
        assert calls == ["slow", "next"]

    asyncio.run(run())


def test_cdp_evaluate_reads_values_and_ignores_browser_events():
    class Socket:
        async def send(self, raw):
            command = json.loads(raw)
            assert command["params"]["returnByValue"] is True
            self.responses = iter([
                json.dumps({"method": "Page.frameNavigated", "params": {}}),
                json.dumps({"id": command["id"], "result": {"result": {"value": {"ready": True}}}}),
            ])

        async def recv(self):
            return next(self.responses)

    assert asyncio.run(_CDPPage(Socket()).evaluate("document.readyState")) == {"ready": True}


def test_navigation_waits_for_new_document_before_using_the_editor():
    class Page:
        frame_checks = 0

        async def call(self, method, params=None):
            if method == "Page.navigate":
                return {"loaderId": "new-document"}
            assert method == "Page.getFrameTree"
            self.frame_checks += 1
            return {"frameTree": {"frame": {"loaderId": "old-document" if self.frame_checks == 1 else "new-document"}}}

        async def evaluate(self, expression):
            assert self.frame_checks >= 2, "Do not read the stale editor while navigation is pending"
            return {"ready": True, "url": "https://m365.cloud.microsoft/chat", "replies": []}

    snapshot = asyncio.run(_load_chat(Page(), "https://m365.cloud.microsoft/chat", None))
    assert snapshot["ready"]


def test_persistent_navigation_waits_for_saved_reply_to_load():
    class Page:
        snapshots = 0

        async def call(self, method, params=None):
            if method == "Page.navigate":
                return {"loaderId": "new"}
            return {"frameTree": {"frame": {"loaderId": "new"}}}

        async def evaluate(self, expression):
            self.snapshots += 1
            replies = [] if self.snapshots == 1 else [{"id": "saved-reply", "text": "Previous answer", "complete": True}]
            return {"ready": True, "url": "https://m365.cloud.microsoft/chat/conversation/test", "replies": replies}

    page = Page()
    snapshot = asyncio.run(_load_chat(page, "https://m365.cloud.microsoft/chat/conversation/test", "saved-reply"))
    assert page.snapshots == 2
    assert snapshot["replies"][0]["id"] == "saved-reply"


def test_prompt_insertion_never_overwrites_a_changed_draft():
    class Page:
        inserted = []
        snapshots = iter([{}, {"draft": "User changed the draft"}])

        async def evaluate(self, expression):
            return True if expression == _FOCUS_JS else next(self.snapshots)

        async def call(self, method, params):
            assert method == "Input.insertText"
            self.inserted.append(params["text"])

    page = Page()
    with pytest.raises(SubstrateCopilotError, match="changed unexpectedly"):
        asyncio.run(_enter_prompt(page, "API prompt"))
    assert page.inserted == ["API prompt"]


def test_persistent_turn_waits_for_delayed_conversation_url(monkeypatch):
    url = "https://m365.cloud.microsoft/chat/conversation/saved"
    completed = {"url": "https://m365.cloud.microsoft/chat", "replies": [
        {"id": "new-reply", "text": "answer", "complete": True},
    ]}

    class Page:
        snapshots = iter([{}, {"replies": []}, completed, completed, {**completed, "url": url}])

        async def evaluate(self, expression):
            return True if expression == _SUBMIT_JS else next(self.snapshots)

    class Connection:
        async def __aenter__(self):
            return object()

        async def __aexit__(self, *args):
            pass

    async def run():
        client = BrowserCopilotClient()
        session = PersistentSession()
        monkeypatch.setattr(client, "_open_page", AsyncMock(return_value={"webSocketDebuggerUrl": "ws://localhost/test"}))
        monkeypatch.setattr(browser_client.websockets, "connect", lambda *args, **kwargs: Connection())
        monkeypatch.setattr(browser_client, "_CDPPage", lambda ws: Page())
        monkeypatch.setattr(browser_client, "_load_chat", AsyncMock(return_value={"replies": []}))
        monkeypatch.setattr(browser_client, "_enter_prompt", AsyncMock())
        assert await client.chat("prompt", [], session) == "answer"
        saved = client._conversations[session.conversation_id]
        assert saved.url == url
        assert saved.last_reply_id == "new-reply"

    asyncio.run(run())


def test_browser_backend_supports_all_api_formats_without_token(monkeypatch):
    calls = []

    async def chat(self, prompt, context, session=None):
        calls.append((prompt, context))
        return "browser reply"

    async def status(self):
        return {"connected": True, "copilot_tab_open": True, "bridge_tab_open": True}

    monkeypatch.setattr(BrowserCopilotClient, "chat", chat)
    monkeypatch.setattr(BrowserCopilotClient, "status", status)
    app = create_app(Settings(M365_BACKEND="browser", M365_ACCESS_TOKEN="", _env_file=None))
    with TestClient(app) as client:
        assert client.get("/healthz").json()["backend"] == "browser"
        assert client.get("/v1/token/status").json() == {"required": False, "backend": "browser"}
        requests = [
            ("/v1/chat/completions", {"messages": [{"role": "user", "content": "chat"}]}),
            ("/v1/responses", {"input": "responses"}),
            ("/v1/messages", {"messages": [{"role": "user", "content": "messages"}]}),
        ]
        for path, body in requests:
            response = client.post(path, json={"model": "m365-copilot", **body})
            assert response.status_code == 200
            assert "browser reply" in response.text
        response = client.post("/v1/chat/completions", json={
            "model": "m365-copilot", "stream": True,
            "messages": [{"role": "user", "content": "stream"}],
        })
        assert "browser reply" in response.text
        assert "data: [DONE]" in response.text
    assert [prompt for prompt, _ in calls] == ["chat", "responses", "messages", "stream"]


def test_browser_security_check_is_reported_as_an_error(monkeypatch):
    async def chat(self, *args, **kwargs):
        _check_page_blocked({"verification": True})

    monkeypatch.setattr(BrowserCopilotClient, "chat", chat)
    app = create_app(Settings(M365_BACKEND="browser", _env_file=None))
    with TestClient(app) as client:
        response = client.post("/v1/chat/completions", json={
            "model": "m365-copilot", "messages": [{"role": "user", "content": "test"}],
        })
        assert response.status_code == 502
        assert "security check" in response.json()["detail"]


@pytest.mark.skipif(os.environ.get("M365_TEST_CDP") != "1", reason="Opt-in DOM test requires debug Edge on port 9222")
def test_dom_snapshot_distinguishes_hidden_dialogs_and_finished_replies():
    async def run():
        async with httpx.AsyncClient(trust_env=False) as client:
            tab = (await client.put("http://127.0.0.1:9222/json/new?about:blank")).json()
            try:
                async with websockets.connect(tab["webSocketDebuggerUrl"]) as ws:
                    page = _CDPPage(ws)
                    html = '''<span id="m365-chat-editor-target-element" contenteditable="true"> </span>
                        <div id="security" role="dialog" aria-label="Security check required"
                             style="visibility:hidden;width:340px;height:125px">Verify you are human</div>
                        <button aria-label="Stop generating" style="visibility:hidden">Stop</button>
                        <div data-testid="copilot-message-reply-div"><div id="copilot-message-test"><div data-testid="lastChatMessage">
                            <div data-testid="markdown-reply" data-message-id="new" data-message-type="Chat">Answer</div>
                        </div></div><button data-testid="CopyButtonTestId">Copy</button></div>'''
                    await page.evaluate(f"document.body.innerHTML = {json.dumps(html)}")
                    snapshot = await page.evaluate(_SNAPSHOT_JS)
                    assert snapshot["verification"] is False
                    assert snapshot["dialog_open"] is False
                    assert snapshot["busy"] is False
                    assert _completed_reply(snapshot, set()) == "Answer"
                    await page.evaluate("document.getElementById('security').style.visibility = 'visible'")
                    snapshot = await page.evaluate(_SNAPSHOT_JS)
                    assert snapshot["verification"] is True
                    with pytest.raises(SubstrateCopilotError, match="security check"):
                        _check_page_blocked(snapshot)
            finally:
                await client.get(f"http://127.0.0.1:9222/json/close/{tab['id']}")

    asyncio.run(run())
