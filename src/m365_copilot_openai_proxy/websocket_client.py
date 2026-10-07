from __future__ import annotations

import asyncio
import copy
import json
import re
import time
import uuid
from collections.abc import AsyncIterator
from dataclasses import dataclass, field
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

import websockets

from .browser_client import BrowserCopilotClient
from .session_store import PersistentSession
from .substrate_client import SIGNALR_SEP, SubstrateCopilotError, _combine_text


def _frames(raw: str | bytes):
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8")
    for part in raw.split(SIGNALR_SEP):
        if part.strip():
            try:
                value = json.loads(part)
            except ValueError:
                continue
            if isinstance(value, dict):
                yield value


@dataclass(repr=False)
class CapturedCredentials:
    """Opaque credentials and protocol template; never written to disk or logged."""

    url: str
    invocation: dict
    prelude: list[dict]
    created_at: float = field(default_factory=time.monotonic)

    def request(self, text: str, conversation_id: str, first_turn: bool) -> tuple[str, dict]:
        parsed = urlsplit(self.url)
        if (parsed.scheme != "wss" or parsed.netloc != "substrate.office.com"
                or not parsed.path.startswith("/m365Copilot/Chathub/")):
            raise SubstrateCopilotError("Captured connection is not a Copilot WebSocket.")
        query = dict(parse_qsl(parsed.query))
        if not query.get("access_token"):
            raise SubstrateCopilotError("Captured connection has no access token.")
        request_id = uuid.uuid4().hex
        for key in list(query):
            if key.lower() in ("clientrequestid", "chatsessionid"):
                query[key] = request_id
            elif key.lower() == "conversationid":
                query[key] = conversation_id
        payload = copy.deepcopy(self.invocation)
        payload["invocationId"] = "0"
        argument = payload["arguments"][0]
        argument.update(clientCorrelationId=request_id, requestId=request_id,
                        traceId=request_id, isStartOfSession=first_turn)
        argument["message"].update(text=text, requestId=request_id)
        # X-SessionId, routing key and client session identify the authenticated
        # browser session. Changing them invalidates its verification context.
        return urlunsplit((parsed.scheme, parsed.netloc, parsed.path, urlencode(query), "")), payload


class BrowserCredentialProvider:
    """Refresh with one visible bootstrap chat; API prompts never enter that tab."""

    def __init__(self, cdp_port: int = 9222, timeout_seconds: float = 120,
                 refresh_seconds: float = 2700, auto_verify: bool = False):
        self._browser = BrowserCopilotClient(cdp_port, timeout_seconds, auto_verify=auto_verify)
        self._refresh_seconds = refresh_seconds
        self._credentials: CapturedCredentials | None = None
        self._lock = asyncio.Lock()
        self.refresh_count = 0

    async def get(self) -> CapturedCredentials:
        async with self._lock:
            if self._credentials is None or time.monotonic() - self._credentials.created_at >= self._refresh_seconds:
                self._credentials = await self._capture()
                self.refresh_count += 1
            return self._credentials

    def invalidate(self, credentials: CapturedCredentials) -> None:
        if self._credentials is credentials:
            self._credentials = None

    async def status(self) -> dict:
        browser = await self._browser.status()
        remaining = max(0, self._refresh_seconds - (time.monotonic() - self._credentials.created_at)) if self._credentials else 0
        return {"ready": remaining > 0, "refresh_in_seconds": int(remaining),
                "refresh_count": self.refresh_count, "browser_connected": browser["connected"]}

    async def _capture(self) -> CapturedCredentials:
        tab = await self._browser._open_page()
        captures: dict[str, dict] = {}
        try:
            async with websockets.connect(tab["webSocketDebuggerUrl"], max_size=8 * 1024 * 1024) as monitor:
                await monitor.send(json.dumps({"id": 1, "method": "Network.enable"}))
                async with asyncio.timeout(10):
                    while json.loads(await monitor.recv()).get("id") != 1:
                        pass

                async def listen():
                    async for raw in monitor:
                        event = json.loads(raw)
                        method, params = event.get("method"), event.get("params", {})
                        request_id = params.get("requestId")
                        if method == "Network.webSocketCreated":
                            url = params.get("url", "")
                            if urlsplit(url).netloc == "substrate.office.com":
                                captures[request_id] = {"url": url, "prelude": []}
                        if method != "Network.webSocketFrameSent" or request_id not in captures:
                            continue
                        capture = captures[request_id]
                        for frame in _frames(params.get("response", {}).get("payloadData", "")):
                            if frame.get("type") == 1 and frame.get("target") == "send":
                                capture["prelude"].append(frame)
                            elif frame.get("type") == 4 and frame.get("target") == "chat":
                                capture["invocation"] = frame

                task = asyncio.create_task(listen())
                marker = "PROXY_AUTH_" + uuid.uuid4().hex[:12]
                try:
                    reply = await self._browser.chat("Connection initialization. Reply with exactly " + marker, [])
                    await asyncio.sleep(0.1)
                    if marker not in reply:
                        raise SubstrateCopilotError("Copilot credential initialization did not complete. Check its dedicated Edge tab.")
                    if task.done():
                        task.result()
                    capture = next((c for c in reversed(list(captures.values())) if "invocation" in c), None)
                    if capture is None:
                        raise SubstrateCopilotError("Copilot did not expose a chat WebSocket during initialization.")
                    return CapturedCredentials(**capture)
                finally:
                    task.cancel()
                    await asyncio.gather(task, return_exceptions=True)
        except SubstrateCopilotError:
            raise
        except Exception as exc:
            raise SubstrateCopilotError("Cannot initialize WebSocket credentials from Edge. Keep the dedicated profile signed in.") from exc


class _UpstreamError(SubstrateCopilotError):
    def __init__(self, value: str, code: str = ""):
        self.code = code or value
        super().__init__(f"Copilot rejected the WebSocket request: {value}" + (f" ({code})." if code else "."))


def _safe_code(value) -> str:
    return value if isinstance(value, str) and re.fullmatch(r"[A-Za-z0-9_]{1,80}", value) else "UnknownError"


async def _read_reply(ws) -> AsyncIterator[str]:
    fallback = ""
    emitted = False
    async for raw in ws:
        for message in _frames(raw):
            kind = message.get("type")
            if kind == 2:
                item = message.get("item") or {}
                result = item.get("result") or {}
                if result.get("value") not in (None, "Success"):
                    raise _UpstreamError(_safe_code(result.get("value")), _safe_code(result["errorCode"]) if result.get("errorCode") else "")
                containers = [item]
            elif kind == 1:
                containers = message.get("arguments") or []
            elif kind == 3:
                if message.get("error"):
                    raise _UpstreamError("InvocationFailed")
                if not emitted and fallback:
                    yield fallback
                elif not emitted:
                    raise _UpstreamError("EmptyResponse")
                return
            else:
                continue
            for container in containers:
                delta = container.get("writeAtCursor")
                if isinstance(delta, str) and delta:
                    if not emitted and fallback:
                        yield fallback
                    emitted = True
                    yield delta
                entries = container.get("messages") or []
                if isinstance(entries, dict):
                    entries = [entries]
                for entry in reversed(entries):
                    if entry.get("author") == "bot" and entry.get("messageType", "Chat") == "Chat" and entry.get("text"):
                        fallback = entry["text"]
                        break
    raise _UpstreamError("ConnectionEndedBeforeCompletion")


class WebSocketCopilotClient:
    def __init__(self, provider: BrowserCredentialProvider, timeout_seconds: float = 120):
        self._provider = provider
        self._timeout = timeout_seconds
        self._lock = asyncio.Lock()

    async def chat(self, prompt: str, additional_context: list[str], session: PersistentSession | None = None) -> str:
        return "".join([part async for part in self.chat_stream(prompt, additional_context, session)])

    async def chat_stream(self, prompt: str, additional_context: list[str], session: PersistentSession | None = None) -> AsyncIterator[str]:
        text = _combine_text(prompt, additional_context)
        if not text.strip():
            raise SubstrateCopilotError("A non-empty prompt is required.")
        async with self._lock:
            conversation = session.conversation_id if session else str(uuid.uuid4())
            first = session is None or session.turn_count == 0
            for attempt in range(2):
                credentials = await self._provider.get()
                emitted = False
                try:
                    async with asyncio.timeout(self._timeout):
                        async for part in self._turn(credentials, text, conversation, first):
                            emitted = True
                            yield part
                    if session:
                        session.turn_count += 1
                    return
                except _UpstreamError as exc:
                    if not emitted and attempt == 0 and exc.code in {
                        "ChallengeProofRequired", "ChallengeProofRejected", "UnauthorizedRequest",
                        "InvalidAuthenticationToken", "TokenExpired", "Unauthorized",
                    }:
                        self._provider.invalidate(credentials)
                        continue
                    raise
                except TimeoutError as exc:
                    raise SubstrateCopilotError("Copilot WebSocket timed out. The request was not automatically resubmitted.") from exc
                except SubstrateCopilotError:
                    raise
                except Exception as exc:
                    response = getattr(exc, "response", None)
                    status = getattr(response, "status_code", None) or getattr(exc, "status_code", None)
                    if not emitted and attempt == 0 and status in (401, 403):
                        self._provider.invalidate(credentials)
                        continue
                    raise SubstrateCopilotError("Copilot WebSocket connection failed. The request was not automatically resubmitted.") from exc

    async def _turn(self, credentials: CapturedCredentials, text: str, conversation: str, first: bool) -> AsyncIterator[str]:
        url, payload = credentials.request(text, conversation, first)
        # Copilot returns MultipleRequestsNotSupported if a socket is reused.
        async with websockets.connect(url, additional_headers={"Origin": "https://m365.cloud.microsoft"}, max_size=8 * 1024 * 1024) as ws:
            await ws.send(json.dumps({"protocol": "json", "version": 1}) + SIGNALR_SEP)
            handshake = list(_frames(await ws.recv()))
            if {} not in handshake:
                raise _UpstreamError("HandshakeRejected")
            await ws.send(json.dumps({"type": 6}) + SIGNALR_SEP)
            for frame in credentials.prelude:
                await ws.send(json.dumps(frame, ensure_ascii=False) + SIGNALR_SEP)
            await ws.send(json.dumps(payload, ensure_ascii=False) + SIGNALR_SEP)
            async for part in _read_reply(ws):
                yield part
