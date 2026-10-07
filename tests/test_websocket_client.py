from __future__ import annotations

import asyncio
import json
from unittest.mock import AsyncMock
from urllib.parse import parse_qs, urlsplit

import pytest

from m365_copilot_openai_proxy.websocket_client import CapturedCredentials, WebSocketCopilotClient
from m365_copilot_openai_proxy.session_store import PersistentSession
from m365_copilot_openai_proxy.substrate_client import SubstrateCopilotError


def credentials():
    return CapturedCredentials(
        url="wss://substrate.office.com/m365Copilot/Chathub/user@tenant?access_token=opaque.encrypted.token.parts.value&ConversationId=bootstrap&X-SessionId=authenticated-session&XRoutingParameterSessionKey=routing&clientrequestid=old&chatsessionid=old&source=%22officeweb%22",
        invocation={"type": 4, "target": "chat", "invocationId": "0", "arguments": [{
            "sessionId": "authenticated-session", "requestId": "old", "traceId": "old", "clientCorrelationId": "old",
            "message": {"text": "bootstrap", "requestId": "old"}, "isStartOfSession": False,
        }]},
        prelude=[{"type": 1, "target": "send", "arguments": [{"type": "ChallengeProof", "method": "CloudflareTurnstile", "token": "secret-proof"}]}],
    )


def test_request_ids_match_url_and_payload_without_changing_authenticated_session():
    captured = credentials()
    url, payload = captured.request("actual prompt", "new-conversation", True)
    query = parse_qs(urlsplit(url).query)
    arg = payload["arguments"][0]
    request = arg["requestId"]
    assert len(request) == 32 and "-" not in request
    assert query["chatsessionid"] == query["clientrequestid"] == [request]
    assert arg["traceId"] == arg["clientCorrelationId"] == arg["message"]["requestId"] == request
    assert query["ConversationId"] == ["new-conversation"]
    assert query["X-SessionId"] == [arg["sessionId"]] == ["authenticated-session"]
    assert query["XRoutingParameterSessionKey"] == ["routing"]
    assert query["source"] == ['"officeweb"']
    assert arg["message"]["text"] == "actual prompt"
    assert arg["isStartOfSession"] is True
    assert captured.invocation["arguments"][0]["message"]["text"] == "bootstrap"
    assert "secret-proof" not in repr(captured) and "opaque" not in repr(captured)


def test_capture_does_not_accept_an_unrelated_websocket_host():
    captured = credentials()
    captured.url = captured.url.replace("substrate.office.com", "substrate.office.com.example")
    with pytest.raises(SubstrateCopilotError):
        captured.request("prompt", "conversation", True)


class Connection:
    def __init__(self, frames, sent):
        self.frames = iter(frames)
        self.sent = sent

    async def __aenter__(self):
        return self

    async def __aexit__(self, *args):
        pass

    async def send(self, data):
        self.sent.append(json.loads(data.rstrip("\x1e")))

    async def recv(self):
        return "{}\x1e"

    def __aiter__(self):
        return self

    async def __anext__(self):
        try:
            return json.dumps(next(self.frames)) + "\x1e"
        except StopIteration:
            raise StopAsyncIteration


def success(text):
    return [
        {"type": 1, "target": "update", "arguments": [{"writeAtCursor": text}]},
        {"type": 2, "item": {"result": {"value": "Success"}}},
        {"type": 3},
    ]


def test_persistent_turns_reconnect_but_keep_conversation(monkeypatch):
    from m365_copilot_openai_proxy import websocket_client
    provider = type("Provider", (), {"get": AsyncMock(return_value=credentials())})()
    calls = []
    def connect(url, **kwargs):
        sent = []
        calls.append((url, sent))
        return Connection(success("answer"), sent)
    monkeypatch.setattr(websocket_client.websockets, "connect", connect)
    async def run():
        client = WebSocketCopilotClient(provider)
        session = PersistentSession()
        assert await client.chat("first", [], session) == "answer"
        assert await client.chat("second", [], session) == "answer"
        assert session.turn_count == 2
    asyncio.run(run())
    assert len(calls) == 2
    queries = [parse_qs(urlsplit(url).query) for url, _ in calls]
    assert queries[0]["ConversationId"] == queries[1]["ConversationId"]
    assert queries[0]["chatsessionid"] != queries[1]["chatsessionid"]
    for _, sent in calls:
        assert [m.get("target") for m in sent] == [None, None, "send", "chat"]
    assert calls[0][1][-1]["arguments"][0]["isStartOfSession"] is True
    assert calls[1][1][-1]["arguments"][0]["isStartOfSession"] is False


def test_auth_rejection_refreshes_once_without_returning_an_empty_answer(monkeypatch):
    from m365_copilot_openai_proxy import websocket_client
    provider = type("Provider", (), {"get": AsyncMock(return_value=credentials()), "invalidate": lambda self, value: None})()
    connections = []
    def connect(*args, **kwargs):
        frames = success("fresh answer") if connections else [{"type": 2, "item": {"result": {"value": "UnauthorizedRequest", "errorCode": "ChallengeProofRejected"}}}, {"type": 3}]
        connections.append(frames)
        return Connection(frames, [])
    monkeypatch.setattr(websocket_client.websockets, "connect", connect)
    assert asyncio.run(WebSocketCopilotClient(provider).chat("prompt", [])) == "fresh answer"
    assert len(connections) == 2


def test_generic_failure_is_not_automatically_resubmitted(monkeypatch):
    from m365_copilot_openai_proxy import websocket_client
    provider = type("Provider", (), {"get": AsyncMock(return_value=credentials())})()
    connect = lambda *args, **kwargs: Connection([{"type": 2, "item": {"result": {"value": "InvalidRequest", "errorCode": "MultipleRequestsNotSupported", "message": "secret-server-text"}}}, {"type": 3}], [])
    monkeypatch.setattr(websocket_client.websockets, "connect", connect)
    with pytest.raises(SubstrateCopilotError, match="MultipleRequestsNotSupported") as error:
        asyncio.run(WebSocketCopilotClient(provider).chat("prompt", []))
    assert "secret-server-text" not in str(error.value)
    assert provider.get.await_count == 1


def test_quota_errors_do_not_trigger_credential_refresh(monkeypatch):
    from m365_copilot_openai_proxy import websocket_client
    provider = type("Provider", (), {"get": AsyncMock(return_value=credentials())})()
    frames = [{"type": 2, "item": {"result": {"value": "OutOfCredits"}}}, {"type": 3}]
    monkeypatch.setattr(websocket_client.websockets, "connect", lambda *args, **kwargs: Connection(frames, []))
    with pytest.raises(SubstrateCopilotError, match="OutOfCredits"):
        asyncio.run(WebSocketCopilotClient(provider).chat("prompt", []))
    assert provider.get.await_count == 1


def test_expired_access_token_at_http_handshake_refreshes_once(monkeypatch):
    from m365_copilot_openai_proxy import websocket_client
    class Unauthorized(Exception):
        status_code = 401
    provider = type("Provider", (), {"get": AsyncMock(return_value=credentials()), "invalidate": lambda self, value: None})()
    calls = []
    def connect(*args, **kwargs):
        calls.append(True)
        if len(calls) == 1:
            raise Unauthorized("must not expose token-bearing URLs")
        return Connection(success("renewed"), [])
    monkeypatch.setattr(websocket_client.websockets, "connect", connect)
    assert asyncio.run(WebSocketCopilotClient(provider).chat("prompt", [])) == "renewed"
    assert len(calls) == 2


def test_snapshot_before_first_delta_is_not_lost_or_duplicated():
    from m365_copilot_openai_proxy.websocket_client import _read_reply
    frames = [
        {"type": 1, "arguments": [{"messages": [{"author": "bot", "messageType": "Chat", "text": "WS"}]}]},
        {"type": 1, "arguments": [{"writeAtCursor": "_API_READY"}]},
        {"type": 2, "item": {"result": {"value": "Success"}, "messages": [{"author": "bot", "text": "WS_API_READY"}]}},
        {"type": 3},
    ]
    async def run():
        return [part async for part in _read_reply(Connection(frames, []))]
    assert "".join(asyncio.run(run())) == "WS_API_READY"


def test_expired_credentials_refresh_once_for_concurrent_callers(monkeypatch):
    from m365_copilot_openai_proxy.websocket_client import BrowserCredentialProvider
    async def run():
        provider = BrowserCredentialProvider(refresh_seconds=2700)
        old, fresh = credentials(), credentials()
        old.created_at -= 2701
        provider._credentials = old
        capture = AsyncMock(return_value=fresh)
        monkeypatch.setattr(provider, "_capture", capture)
        results = await asyncio.gather(provider.get(), provider.get())
        assert results == [fresh, fresh]
        assert capture.await_count == 1
        assert provider.refresh_count == 1
        provider.invalidate(old)
        assert provider._credentials is fresh
    asyncio.run(run())


def test_partial_response_is_never_retried_on_auth_error(monkeypatch):
    from m365_copilot_openai_proxy import websocket_client
    provider = type("Provider", (), {"get": AsyncMock(return_value=credentials())})()
    frames = [
        {"type": 1, "arguments": [{"writeAtCursor": "partial"}]},
        {"type": 2, "item": {"result": {"value": "UnauthorizedRequest", "errorCode": "ChallengeProofRejected"}}},
    ]
    monkeypatch.setattr(websocket_client.websockets, "connect", lambda *args, **kwargs: Connection(frames, []))
    async def run():
        seen = []
        with pytest.raises(SubstrateCopilotError, match="ChallengeProofRejected"):
            async for part in WebSocketCopilotClient(provider).chat_stream("prompt", []):
                seen.append(part)
        assert seen == ["partial"]
    asyncio.run(run())
    assert provider.get.await_count == 1


def test_stateless_requests_use_distinct_conversations(monkeypatch):
    from m365_copilot_openai_proxy import websocket_client
    provider = type("Provider", (), {"get": AsyncMock(return_value=credentials())})()
    conversations = []
    def connect(url, **kwargs):
        conversations.append(parse_qs(urlsplit(url).query)["ConversationId"])
        return Connection(success("ok"), [])
    monkeypatch.setattr(websocket_client.websockets, "connect", connect)
    async def run():
        client = WebSocketCopilotClient(provider)
        assert await client.chat("first", []) == "ok"
        assert await client.chat("second", []) == "ok"
    asyncio.run(run())
    assert conversations[0] != conversations[1]


def test_api_websocket_backend_uses_the_shared_client(monkeypatch):
    from fastapi.testclient import TestClient
    from m365_copilot_openai_proxy.app import create_app
    from m365_copilot_openai_proxy.config import Settings
    monkeypatch.setattr(WebSocketCopilotClient, "chat", AsyncMock(return_value="direct reply"))
    app = create_app(Settings(M365_BACKEND="websocket", _env_file=None))
    with TestClient(app) as client:
        for endpoint, body in [
            ("/v1/chat/completions", {"messages": [{"role": "user", "content": "hello"}]}),
            ("/v1/responses", {"input": "hello"}),
            ("/v1/messages", {"messages": [{"role": "user", "content": "hello"}], "max_tokens": 20}),
        ]:
            response = client.post(endpoint, json={"model": "m365-copilot", **body})
            assert response.status_code == 200
            assert "direct reply" in response.text
