from __future__ import annotations

import asyncio
import json
import re
from collections.abc import AsyncIterator
from dataclasses import dataclass
from urllib.parse import urlsplit

import httpx
import websockets

from .session_store import PersistentSession
from .substrate_client import SubstrateCopilotError, _combine_text

_CHAT_URL = "https://m365.cloud.microsoft/chat"
_EDITOR_JS = "document.querySelector('#m365-chat-editor-target-element, [contenteditable=true][role=textbox], textarea')"
_SNAPSHOT_JS = """
(() => {
    const editor = EDITOR;
    const visible = e => Boolean(e && e.getBoundingClientRect().width
        && e.checkVisibility({checkOpacity: true, checkVisibilityCSS: true}));
    const buttons = [...document.querySelectorAll('button')].filter(visible);
    const dialogs = [...document.querySelectorAll('[role="dialog"], [role="alertdialog"]')].filter(visible);
    const verification = dialogs.some(e =>
        /security check|verify.*human|verification required|captcha|安全性檢查|安全性检查|安全驗證|安全验证|セキュリティ.*(確認|チェック)/i
            .test((e.getAttribute('aria-label') || '') + ' ' + e.innerText));
    const busy = buttons.some(b =>
        /stop responding|stop generating|停止回應|停止回覆|停止生成|停止回答|応答を停止|生成を停止/i.test(b.getAttribute('aria-label') || '')
        || /stop.*(respond|generat)/i.test(b.getAttribute('data-testid') || ''));
    const replies = [...document.querySelectorAll('[data-testid="markdown-reply"]')]
        .filter(e => !e.dataset.messageType || e.dataset.messageType === 'Chat')
        .map(e => {
            const container = e.closest('[data-testid="copilot-message-reply-div"]');
            const copy = container?.querySelector('[data-testid="CopyButtonTestId"]');
            return {id: e.dataset.messageId, text: e.innerText,
                complete: Boolean(copy && !copy.disabled && visible(copy))};
        });
    return {url: location.href, ready: document.readyState === 'complete' && visible(editor),
        draft: editor ? (editor.value ?? editor.innerText).replace(/[\u200b-\u200d\ufeff]/g, '').trim() : '',
        busy, replies, verification, dialog_open: dialogs.length > 0};
})()
""".replace("EDITOR", _EDITOR_JS)
_FOCUS_JS = """
(() => {
    const editor = EDITOR;
    if (!editor || editor.disabled || editor.getAttribute('aria-disabled') === 'true') return false;
    if ((editor.value ?? editor.innerText).replace(/[\u200b-\u200d\ufeff]/g, '').trim()) return false;
    editor.focus();
    return document.activeElement === editor;
})()
""".replace("EDITOR", _EDITOR_JS)
_SUBMIT_JS = """
(() => {
    const button = document.querySelector('button.fai-SendButton[type="submit"]');
    if (!button || button.disabled || !button.checkVisibility({checkOpacity: true, checkVisibilityCSS: true})) return false;
    button.click();
    return true;
})()
"""


class _CDPPage:
    """One caller owns the connection; unrelated browser events are discarded."""

    def __init__(self, ws, auto_verify: bool = False):
        self._ws = ws
        self._sequence = 0
        self.auto_verify = auto_verify
        self.verification_attempted = False

    async def call(self, method: str, params: dict | None = None, *, session_id: str | None = None) -> dict:
        self._sequence += 1
        request_id = self._sequence
        command = {"id": request_id, "method": method, "params": params or {}}
        if session_id:
            command["sessionId"] = session_id
        await self._ws.send(json.dumps(command))
        async with asyncio.timeout(10):
            while True:
                response = json.loads(await self._ws.recv())
                if response.get("id") != request_id:
                    continue
                if "error" in response:
                    raise SubstrateCopilotError(f"Edge rejected browser command {method}.")
                return response.get("result", {})

    async def evaluate(self, expression: str):
        result = await self.call("Runtime.evaluate", {
            "expression": expression, "returnByValue": True, "awaitPromise": True,
        })
        if "exceptionDetails" in result:
            raise SubstrateCopilotError("Cannot read the Copilot page. Reload it in Edge and retry.")
        return result.get("result", {}).get("value")


def _conversation_url(url: str) -> bool:
    parsed = urlsplit(url)
    return parsed.scheme == "https" and parsed.netloc == "m365.cloud.microsoft" and parsed.path.startswith("/chat/conversation/")


def _completed_reply(snapshot: dict, previous_ids: set[str]) -> str | None:
    replies = snapshot.get("replies", [])
    if snapshot.get("busy") or not replies:
        return None
    reply = replies[-1]
    if reply.get("id") and reply["id"] not in previous_ids and reply.get("complete") and reply.get("text", "").strip():
        return reply["text"]
    return None


def _check_page_blocked(snapshot: dict) -> None:
    if snapshot.get("verification"):
        raise SubstrateCopilotError(
            "Microsoft requires a security check in the bridge tab. "
            "Complete it yourself in Edge, then retry. The bridge will not solve or bypass it."
        )
    if snapshot.get("dialog_open"):
        raise SubstrateCopilotError("A dialog is blocking the bridge tab. Review it in Edge, then retry.")


async def _load_chat(page: _CDPPage, destination: str, last_reply_id: str | None) -> dict:
    navigation = await page.call("Page.navigate", {"url": destination})
    if navigation.get("errorText"):
        raise SubstrateCopilotError("Edge could not load the Copilot chat page.")
    while True:
        # The old document can still report readyState=complete immediately
        # after Page.navigate. Do not type until the NEW document has committed.
        tree = await page.call("Page.getFrameTree")
        if navigation.get("loaderId") and tree["frameTree"]["frame"].get("loaderId") != navigation["loaderId"]:
            await asyncio.sleep(0.1)
            continue
        snapshot = await _read_snapshot(page)
        if snapshot and snapshot.get("ready") and snapshot.get("url", "").rstrip("/") == destination.rstrip("/"):
            replies = snapshot.get("replies", [])
            history_loaded = (bool(replies) and replies[-1].get("id") == last_reply_id) if last_reply_id else not replies
            if not snapshot.get("busy") and history_loaded:
                return snapshot
        await asyncio.sleep(0.3)


async def _read_snapshot(page: _CDPPage) -> dict:
    snapshot = await page.evaluate(_SNAPSHOT_JS) or {}
    # Give the website time to complete its normal checks. Opt-in interaction is
    # limited to one recognized checkbox in a Cloudflare frame owned by this tab.
    attempted = getattr(page, "verification_attempted", False)
    for poll in range(60 if getattr(page, "auto_verify", False) else 20):
        if not snapshot.get("verification"):
            break
        if poll >= 3 and getattr(page, "auto_verify", False) and not attempted:
            attempted = await _try_verification_checkbox(page)
            if attempted:
                page.verification_attempted = True
        await asyncio.sleep(0.3)
        snapshot = await page.evaluate(_SNAPSHOT_JS) or {}
    _check_page_blocked(snapshot)
    return snapshot


async def _try_verification_checkbox(page: _CDPPage) -> bool:
    """Click only a visible, named verification checkbox; no puzzle solving."""
    targets = (await page.call("Target.getTargets")).get("targetInfos", [])
    for target in targets:
        parsed = urlsplit(target.get("url", ""))
        if target.get("type") != "iframe" or parsed.scheme != "https" or parsed.netloc != "challenges.cloudflare.com":
            continue
        session_id = None
        try:
            # A frame from another tab cannot be resolved by this page's DOM.
            owner = await page.call("DOM.getFrameOwner", {"frameId": target["targetId"]})
            remote = await page.call("DOM.resolveNode", {"backendNodeId": owner["backendNodeId"]})
            object_id = remote["object"]["objectId"]
            try:
                visibility = await page.call("Runtime.callFunctionOn", {
                    "objectId": object_id, "returnByValue": True,
                    "functionDeclaration": """function() {
                        const dialog = this.closest('[role="dialog"], [role="alertdialog"]');
                        return Boolean(dialog && this.checkVisibility({checkOpacity: true, checkVisibilityCSS: true})
                            && /security check|verification|verify|驗證|验证|安全性檢查|安全性检查|セキュリティ/i.test(
                                (dialog.getAttribute('aria-label') || '') + ' ' + dialog.innerText));
                    }""",
                })
            finally:
                await page.call("Runtime.releaseObject", {"objectId": object_id})
            if visibility.get("result", {}).get("value") is not True:
                continue
            outer = (await page.call("DOM.getBoxModel", {"backendNodeId": owner["backendNodeId"]}))["model"]["content"]
            attached = await page.call("Target.attachToTarget", {"targetId": target["targetId"], "flatten": True})
            session_id = attached["sessionId"]
            nodes = (await page.call("Accessibility.getFullAXTree", session_id=session_id)).get("nodes", [])
            checkboxes = [node for node in nodes if not node.get("ignored")
                and node.get("role", {}).get("value") == "checkbox"
                and re.search(r"verify.*human|驗證.*人|验证.*人|人間.*確認", node.get("name", {}).get("value", ""), re.I)
                and not any(p.get("name") in ("disabled", "checked") and p.get("value", {}).get("value") in (True, "true") for p in node.get("properties", []))]
            if len(checkboxes) != 1 or not checkboxes[0].get("backendDOMNodeId"):
                continue
            model = (await page.call("DOM.getBoxModel", {"backendNodeId": checkboxes[0]["backendDOMNodeId"]}, session_id=session_id))["model"]
            quad = model["content"]
            # Frame-local checkbox coordinates translated to the parent viewport.
            x = min(outer[::2]) + sum(quad[::2]) / 4
            y = min(outer[1::2]) + sum(quad[1::2]) / 4
            if model.get("width", 0) <= 0 or model.get("height", 0) <= 0:
                continue
            if not (min(outer[::2]) < x < max(outer[::2]) and min(outer[1::2]) < y < max(outer[1::2])):
                continue
            # Record before dispatch: an ambiguous input result must not repeat.
            page.verification_attempted = True
            await page.call("Input.dispatchMouseEvent", {"type": "mousePressed", "button": "left", "clickCount": 1, "x": x, "y": y})
            await page.call("Input.dispatchMouseEvent", {"type": "mouseReleased", "button": "left", "clickCount": 1, "x": x, "y": y})
            return True
        except (SubstrateCopilotError, KeyError, TypeError, ValueError):
            if getattr(page, "verification_attempted", False):
                return True
        finally:
            if session_id:
                try:
                    await page.call("Target.detachFromTarget", {"sessionId": session_id})
                except SubstrateCopilotError:
                    pass
    return False


async def _enter_prompt(page: _CDPPage, text: str) -> None:
    # Copilot may replace its editor while the app hydrates after navigation.
    # Retry insertion only when that replacement leaves an empty editor; never
    # overwrite a different draft or resubmit a message.
    await asyncio.sleep(0.7)
    expected = text.replace("\r\n", "\n").strip()
    for _ in range(3):
        await _read_snapshot(page)
        if not await page.evaluate(_FOCUS_JS):
            raise SubstrateCopilotError("Copilot's message box is unavailable or contains an unsent draft. Check Edge before retrying.")
        await page.call("Input.insertText", {"text": text})
        await asyncio.sleep(0.3)
        snapshot = await _read_snapshot(page)
        draft = (snapshot or {}).get("draft", "")
        if draft == expected:
            return
        if draft:
            raise SubstrateCopilotError("The message box content changed unexpectedly. Review the unsent draft in Edge.")
        await asyncio.sleep(0.5)
    raise SubstrateCopilotError("Copilot's editor did not retain the prompt. Wait for the page to finish loading, then retry.")


@dataclass(frozen=True)
class _BrowserConversation:
    url: str
    last_reply_id: str


class BrowserCopilotClient:
    """Send through the signed-in website, without exporting or replaying tokens.

    A single bridge-owned tab is reused. The lock prevents simultaneous requests
    from mixing replies. Persistent sessions save their conversation URL; ordinary
    requests always navigate to a new chat.
    """

    def __init__(self, cdp_port: int = 9222, timeout_seconds: float = 120, auto_verify: bool = False):
        self._base_url = f"http://127.0.0.1:{cdp_port}"
        self._timeout = timeout_seconds
        self._lock = asyncio.Lock()
        self._tab_id: str | None = None
        self._conversations: dict[str, _BrowserConversation] = {}
        self._phase = "connecting to Edge"
        self._auto_verify = auto_verify

    async def status(self) -> dict:
        try:
            async with httpx.AsyncClient(timeout=2, trust_env=False) as client:
                response = await client.get(f"{self._base_url}/json")
                response.raise_for_status()
                tabs = response.json()
            return {"connected": True, "copilot_tab_open": any(
                tab.get("type") == "page" and tab.get("url", "").startswith(_CHAT_URL)
                for tab in tabs
            ), "bridge_tab_open": any(tab.get("id") == self._tab_id for tab in tabs)}
        except (httpx.HTTPError, ValueError):
            return {"connected": False, "copilot_tab_open": False, "bridge_tab_open": False}

    async def _open_page(self) -> dict:
        async with httpx.AsyncClient(timeout=5, trust_env=False) as client:
            response = await client.get(f"{self._base_url}/json")
            response.raise_for_status()
            tabs = response.json()
            owned = next((tab for tab in tabs if tab.get("id") == self._tab_id), None)
            if owned:
                return owned
            # The dedicated profile shares the user's sign-in, not their open tab.
            response = await client.put(f"{self._base_url}/json/new?{_CHAT_URL}")
            response.raise_for_status()
            tab = response.json()
            self._tab_id = tab["id"]
            return tab

    async def chat(
        self,
        prompt: str,
        additional_context: list[str],
        session: PersistentSession | None = None,
    ) -> str:
        text = _combine_text(prompt, additional_context)
        if not text.strip():
            raise SubstrateCopilotError("A non-empty prompt is required.")
        async with self._lock:
            try:
                async with asyncio.timeout(self._timeout):
                    return await self._chat_turn(text, session)
            except TimeoutError as exc:
                raise SubstrateCopilotError(
                    f"Browser bridge timed out while {self._phase} after {self._timeout:g} seconds. "
                    "Check the bridge tab in Edge for sign-in, verification, or an error. "
                    "The request was not automatically resubmitted."
                ) from exc
            except SubstrateCopilotError:
                raise
            except Exception as exc:
                # Browser errors may contain URLs or page content. Keep them out of API errors.
                raise SubstrateCopilotError(
                    "Cannot communicate with the Copilot browser tab. "
                    "Keep the dedicated Edge window open and signed in, then retry."
                ) from exc

    async def _chat_turn(self, text: str, session: PersistentSession | None) -> str:
        self._phase = "connecting to Edge"
        tab = await self._open_page()
        session_key = session.conversation_id if session else None
        conversation = self._conversations.get(session_key)
        destination = conversation.url if conversation else _CHAT_URL
        async with websockets.connect(tab["webSocketDebuggerUrl"], max_size=4 * 1024 * 1024) as ws:
            page = _CDPPage(ws, auto_verify=True) if self._auto_verify else _CDPPage(ws)
            # Refuse to discard text if the user has typed into our bridge tab.
            current = await _read_snapshot(page)
            if current and current.get("draft"):
                raise SubstrateCopilotError("The bridge tab contains an unsent draft. Clear or send it in Edge before retrying.")
            self._phase = "loading the chat page"
            snapshot = await _load_chat(page, destination, conversation.last_reply_id if conversation else None)
            previous_ids = {reply["id"] for reply in snapshot.get("replies", []) if reply.get("id")}
            self._phase = "entering the prompt"
            await _enter_prompt(page, text)
            snapshot = await _read_snapshot(page)
            previous_ids.update(reply["id"] for reply in snapshot.get("replies", []) if reply.get("id"))
            self._phase = "submitting the prompt"
            while not await page.evaluate(_SUBMIT_JS):
                await _read_snapshot(page)
                await asyncio.sleep(0.1)
            self._phase = "waiting for the completed reply"
            last_complete: str | None = None
            while True:
                snapshot = await _read_snapshot(page)
                reply = _completed_reply(snapshot or {}, previous_ids)
                # Confirm completion twice to allow the DOM's last render to settle.
                if reply is not None and reply == last_complete:
                    if session_key:
                        if not _conversation_url(snapshot["url"]):
                            # The site can update the route after rendering the
                            # final answer. Wait before recording session state.
                            self._phase = "waiting for Copilot to save the conversation"
                            await asyncio.sleep(0.3)
                            continue
                        self._conversations[session_key] = _BrowserConversation(snapshot["url"], snapshot["replies"][-1]["id"])
                    return reply
                last_complete = reply
                await asyncio.sleep(0.3)

    async def chat_stream(
        self,
        prompt: str,
        additional_context: list[str],
        session: PersistentSession | None = None,
    ) -> AsyncIterator[str]:
        # DOM text can be rewritten during generation. Emit the completed reply
        # once, through the existing SSE adapters, instead of incorrect deltas.
        yield await self.chat(prompt, additional_context, session)
