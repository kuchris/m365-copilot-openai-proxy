from __future__ import annotations

import argparse
import asyncio
import json
import logging
import msvcrt
import re
import subprocess
import threading
import time
from pathlib import Path

import httpx
import uvicorn
import websockets

from .app import create_app
from .config import Settings
from .token_store import decode_jwt_payload, is_substrate_token_claims


class _SuppressCtrlC(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        return "CTRL+C" not in record.getMessage()


logging.getLogger("uvicorn.error").addFilter(_SuppressCtrlC())

_CDP_JS = """
(() => {
    const candidates = [];
    for (const store of [sessionStorage, localStorage]) {
        for (const key of ['LokiAuthToken', ...Object.keys(store).filter(k => k.startsWith('LokiAuthToken'))]) {
            const token = store.getItem(key);
            if (token && token.startsWith('eyJ')) candidates.push(token);
        }
    }
    for (const entry of performance.getEntriesByType('resource')) {
        if (!entry.name.includes('substrate.office.com') ||
            !entry.name.includes('access_token=')) continue;
        const match = entry.name.match(/[?&]access_token=([^&]+)/);
        if (match) candidates.push(decodeURIComponent(match[1]));
    }
    const stores = [sessionStorage, localStorage];
    for (const store of stores) {
        for (const k of Object.keys(store)) {
            if (!k.includes('accesstoken')) continue;
            try {
                const v = JSON.parse(store.getItem(k));
                if (v && v.secret && v.secret.startsWith('eyJ') &&
                    ((v.target && v.target.includes('substrate')) || k.includes('substrate'))) {
                    candidates.push(v.secret);
                }
            } catch {}
        }
    }
    return candidates;
})()
"""

_CDP_NUDGE_JS = """
(() => {
    const input = document.querySelector('[aria-label="Message Copilot"], textarea, [contenteditable="true"], [role="textbox"]');
    if (!input) return false;
    input.focus();
    return true;
})()
"""


async def _cdp_extract_token(port: int, *, allow_nudge: bool = True) -> str | None:
    try:
        async with httpx.AsyncClient(timeout=1) as client:
            tabs = (await client.get(f"http://localhost:{port}/json")).json()
    except Exception:
        return None

    tab = _find_m365_page(tabs)
    if not tab:
        return None

    try:
        async with websockets.connect(tab["webSocketDebuggerUrl"]) as ws:
            await ws.send(json.dumps({"id": 1, "method": "Runtime.evaluate", "params": {"expression": _CDP_JS}}))
            result = json.loads(await ws.recv())
            candidates = result.get("result", {}).get("result", {}).get("value") or []
            for token in candidates:
                if _is_substrate_token(token):
                    return token
            if not allow_nudge:
                return None
            return await _cdp_nudge_and_wait_for_token(ws)
    except Exception:
        return None


async def _cdp_capture_websocket_token(port: int, timeout_seconds: int) -> str | None:
    deadline = asyncio.get_running_loop().time() + timeout_seconds
    while asyncio.get_running_loop().time() < deadline:
        try:
            async with httpx.AsyncClient(timeout=3) as client:
                tabs = (await client.get(f"http://localhost:{port}/json")).json()
        except Exception:
            await asyncio.sleep(1)
            continue

        tab = _find_m365_page(tabs)
        if not tab:
            await asyncio.sleep(1)
            continue

        try:
            async with websockets.connect(tab["webSocketDebuggerUrl"]) as ws:
                await ws.send(json.dumps({"id": 1, "method": "Network.enable"}))
                token = await _wait_for_substrate_websocket_token(ws, deadline)
                if token:
                    return token
        except Exception:
            await asyncio.sleep(1)
            continue
    return None


async def _wait_for_substrate_websocket_token(ws, deadline: float) -> str | None:
    while asyncio.get_running_loop().time() < deadline:
        try:
            raw = await asyncio.wait_for(ws.recv(), timeout=1)
        except asyncio.TimeoutError:
            continue
        msg = json.loads(raw)
        if msg.get("method") != "Network.webSocketCreated":
            continue
        url = msg.get("params", {}).get("url", "")
        if "substrate.office.com" not in url:
            continue
        match = re.search(r"[?&]access_token=([^&]+)", url)
        if not match:
            continue
        token = match.group(1)
        if _is_substrate_token(token):
            return token
    return None


def _find_m365_page(tabs: list[dict]) -> dict | None:
    return next(
        (
            tab for tab in tabs
            if tab.get("type") == "page"
            and tab.get("url", "").startswith("https://m365.cloud.microsoft/")
        ),
        None,
    )


def _wait_for_m365_page(cdp_port: int, timeout_seconds: int) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        try:
            with httpx.Client(timeout=1) as client:
                tabs = client.get(f"http://localhost:{cdp_port}/json").json()
        except Exception:
            time.sleep(0.5)
            continue
        if _find_m365_page(tabs):
            return True
        time.sleep(0.5)
    return False


def _capture_token_to_env(cdp_port: int, timeout_seconds: int) -> bool:
    token = asyncio.run(_cdp_capture_websocket_token(cdp_port, timeout_seconds))
    if not token:
        return False
    _write_token(token)
    return True


def _needs_substrate_token(token: str | None) -> bool:
    if not token or not _is_substrate_token(token):
        return True
    try:
        return _seconds_remaining(token) <= 0
    except Exception:
        return True


def _startup_capture_loop(cdp_port: int, timeout_seconds: int) -> None:
    print("Waiting for the debug Edge M365 tab...")
    _wait_for_m365_page(cdp_port, min(timeout_seconds, 30))
    print("Trying to refresh Substrate token from the debug Edge tab...")
    if _try_auto_refresh(cdp_port):
        return
    print("Waiting for a Substrate token from the debug Edge M365 Copilot tab...")
    print("If needed: press F5 in Copilot, click the message box, and type one character.")
    if _capture_token_to_env(cdp_port, timeout_seconds):
        print(".env updated with Substrate token.")
    else:
        print("Startup token capture timed out. Manual set-token is still available.")

async def _cdp_nudge_and_wait_for_token(ws) -> str | None:
    await ws.send(json.dumps({"id": 2, "method": "Network.enable"}))
    await ws.send(json.dumps({"id": 3, "method": "Runtime.evaluate", "params": {"expression": _CDP_NUDGE_JS}}))
    await ws.send(json.dumps({"id": 4, "method": "Input.insertText", "params": {"text": " "}}))
    await ws.send(json.dumps({
        "id": 5,
        "method": "Input.dispatchKeyEvent",
        "params": {
            "type": "keyDown",
            "windowsVirtualKeyCode": 8,
            "nativeVirtualKeyCode": 8,
            "key": "Backspace",
            "code": "Backspace",
        },
    }))
    await ws.send(json.dumps({
        "id": 6,
        "method": "Input.dispatchKeyEvent",
        "params": {
            "type": "keyUp",
            "windowsVirtualKeyCode": 8,
            "nativeVirtualKeyCode": 8,
            "key": "Backspace",
            "code": "Backspace",
        },
    }))
    deadline = asyncio.get_running_loop().time() + 10
    while asyncio.get_running_loop().time() < deadline:
        try:
            raw = await asyncio.wait_for(ws.recv(), timeout=1)
        except asyncio.TimeoutError:
            continue
        msg = json.loads(raw)
        if msg.get("method") != "Network.webSocketCreated":
            continue
        url = msg.get("params", {}).get("url", "")
        if "substrate.office.com" not in url:
            continue
        match = re.search(r"[?&]access_token=([^&]+)", url)
        if not match:
            continue
        token = match.group(1)
        if _is_substrate_token(token):
            return token
    return None


def _is_substrate_token(token: str) -> bool:
    try:
        claims = decode_jwt_payload(token)
    except Exception:
        return False
    return is_substrate_token_claims(claims)


def _try_auto_refresh(cdp_port: int, *, allow_nudge: bool = True) -> bool:
    token = asyncio.run(_cdp_extract_token(cdp_port, allow_nudge=allow_nudge))
    if not token:
        return False
    _write_token(token)
    print("Token refreshed automatically.")
    return True


def _read_token() -> str | None:
    env_path = Path(".env")
    if not env_path.exists():
        return None
    text = env_path.read_text(encoding="utf-8")
    match = re.search(r"(?m)^M365_ACCESS_TOKEN=(.*)$", text)
    return match.group(1).strip().strip("\"'") if match else None


def _seconds_remaining(token: str) -> int:
    claims = decode_jwt_payload(token)
    return int(claims["exp"]) - int(time.time())


def _auto_refresh_loop(
    cdp_port: int,
    refresh_before_seconds: int,
    retry_seconds: int,
    stop_event: threading.Event,
) -> None:
    while not stop_event.is_set():
        token = _read_token()
        if not token:
            stop_event.wait(retry_seconds)
            continue

        try:
            remaining = _seconds_remaining(token)
        except Exception as exc:
            print(f"Auto-refresh skipped: cannot decode current token: {exc}")
            stop_event.wait(retry_seconds)
            continue

        if remaining > refresh_before_seconds:
            wait_seconds = min(remaining - refresh_before_seconds, 300)
            stop_event.wait(wait_seconds)
            continue

        print(f"Token expires in {max(remaining, 0)} seconds; refreshing from Edge...")
        if not _try_auto_refresh(cdp_port):
            print("Auto-refresh failed; will retry later.")
        stop_event.wait(retry_seconds)


def _write_token(token: str) -> None:
    env_path = Path(".env")
    token_line_pattern = r"(?m)^M365_ACCESS_TOKEN=.*$"
    if env_path.exists():
        text = env_path.read_text(encoding="utf-8")
        if re.search(token_line_pattern, text):
            text = re.sub(token_line_pattern, f"M365_ACCESS_TOKEN={token}", text)
        else:
            text += f"\nM365_ACCESS_TOKEN={token}\n"
    else:
        text = f"M365_ACCESS_TOKEN={token}\n"
    env_path.write_text(text, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(prog="copilot-openai-proxy")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("set-token").set_defaults(func=set_token_command)
    capture_parser = subparsers.add_parser("capture-token")
    capture_parser.add_argument("--cdp-port", type=int, default=9222)
    capture_parser.add_argument("--timeout-seconds", type=int, default=60)
    capture_parser.set_defaults(func=capture_token_command)

    launch_parser = subparsers.add_parser("launch-edge")
    launch_parser.add_argument("--cdp-port", type=int, default=9222)
    launch_parser.set_defaults(func=launch_edge_command)

    serve_parser = subparsers.add_parser("serve")
    serve_parser.add_argument("--host", default="127.0.0.1")
    serve_parser.add_argument("--port", type=int, default=8000)
    serve_parser.add_argument("--cdp-port", type=int, default=9222)
    serve_parser.add_argument("--backend", choices=["substrate", "websocket", "browser"], default=None)
    serve_parser.add_argument("--browser-timeout-seconds", type=float, default=None)
    serve_parser.add_argument("--auto-verify", action="store_true", default=None,
                              help="Try one recognized Cloudflare verification checkbox automatically.")
    serve_parser.add_argument("--no-auto-refresh", action="store_true")
    serve_parser.add_argument("--no-launch-edge", action="store_true")
    serve_parser.add_argument("--no-capture-on-start", action="store_true")
    serve_parser.add_argument("--capture-timeout-seconds", type=int, default=180)
    serve_parser.add_argument("--refresh-before-seconds", type=int, default=300)
    serve_parser.add_argument("--refresh-retry-seconds", type=int, default=60)
    serve_parser.set_defaults(func=serve_command)

    args = parser.parse_args()
    args.func(args)


def launch_edge_command(args: argparse.Namespace) -> None:
    _launch_debug_edge(args.cdp_port)


def _launch_debug_edge(cdp_port: int) -> None:
    profile_dir = Path.home() / ".m365-copilot-openai-proxy" / "edge-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    subprocess.Popen([
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        f"--remote-debugging-port={cdp_port}",
        f"--user-data-dir={profile_dir}",
        "--no-first-run",
        "https://m365.cloud.microsoft/chat",
    ])
    print(f"Edge launched with remote debugging on port {cdp_port}.")
    print(f"Dedicated Edge profile: {profile_dir}")
    print("Sign in to M365 Copilot in that window once, then retry refresh.")


def set_token_command(_args) -> None:
    print("Paste the full WebSocket URL (or just the access_token value), then press Enter:")
    raw = input().strip()
    match = re.search(r"access_token=([^&\s]+)", raw)
    token = match.group(1) if match else raw
    if not token.startswith("eyJ"):
        print("Error: could not find a valid token. Make sure you copied the full WebSocket URL.")
        return
    if not _is_substrate_token(token):
        print("Error: token is not a substrate.office.com WebSocket token.")
        print("Copy the full wss://substrate.office.com/... URL from the Network WebSocket request.")
        return
    _write_token(token)
    print(".env updated.")


def capture_token_command(args: argparse.Namespace) -> None:
    print("Listening for a Substrate WebSocket token...")
    print("In the debug Edge M365 Copilot tab, click the message box and type one character. Do not need to send.")
    token = asyncio.run(_cdp_capture_websocket_token(args.cdp_port, args.timeout_seconds))
    if not token:
        print("Error: no Substrate WebSocket token captured before timeout.")
        return
    _write_token(token)
    print(".env updated with Substrate token.")


def serve_command(args: argparse.Namespace) -> None:
    cdp_port: int = args.cdp_port
    settings = Settings()
    settings_updates = {"cdp_port": cdp_port}
    if args.backend is not None:
        settings_updates["backend"] = args.backend
    if args.auto_verify is not None:
        settings_updates["browser_auto_verify"] = args.auto_verify
    if args.browser_timeout_seconds is not None:
        if args.browser_timeout_seconds <= 0:
            raise SystemExit("--browser-timeout-seconds must be greater than zero.")
        settings_updates["browser_timeout_seconds"] = args.browser_timeout_seconds
    settings = settings.model_copy(update=settings_updates)
    browser_mode = settings.backend in ("browser", "websocket")
    while True:
        app = create_app(settings=settings)
        config = uvicorn.Config(app, host=args.host, port=args.port)
        server = uvicorn.Server(config)
        stop_auto_refresh = threading.Event()
        auto_refresh_thread = None
        capture_thread = None

        if not args.no_launch_edge:
            _launch_debug_edge(cdp_port)

        thread = threading.Thread(target=server.run, daemon=True)
        thread.start()
        if not browser_mode and not args.no_capture_on_start and _needs_substrate_token(_read_token()):
            capture_thread = threading.Thread(
                target=_startup_capture_loop,
                args=(cdp_port, args.capture_timeout_seconds),
                daemon=True,
            )
            capture_thread.start()
        if not browser_mode and not args.no_auto_refresh:
            auto_refresh_thread = threading.Thread(
                target=_auto_refresh_loop,
                args=(
                    cdp_port,
                    args.refresh_before_seconds,
                    args.refresh_retry_seconds,
                    stop_auto_refresh,
                ),
                daemon=True,
            )
            auto_refresh_thread.start()

        while not server.started and thread.is_alive():
            time.sleep(0.05)
        auto_refresh_label = "on demand" if settings.backend == "websocket" else "not required" if browser_mode else ("off" if args.no_auto_refresh else "on")
        capture_label = "not required" if browser_mode else ("off" if args.no_capture_on_start else "on")
        refresh_action = "refresh credentials" if settings.backend == "websocket" else "reconnect browser" if browser_mode else "refresh token"
        print(
            f"\n  backend: {settings.backend}    [q] quit    [r] {refresh_action}"
            f"    auto-refresh: {auto_refresh_label}"
            f"    startup-capture: {capture_label}\n"
        )

        action = None
        while thread.is_alive():
            if msvcrt.kbhit():
                key = msvcrt.getwch().lower()
                if key == "q":
                    action = "quit"
                    server.should_exit = True
                    break
                elif key == "r":
                    action = "refresh"
                    server.should_exit = True
                    break
            time.sleep(0.05)

        stop_auto_refresh.set()
        thread.join()
        if auto_refresh_thread:
            auto_refresh_thread.join(timeout=1)
        if capture_thread:
            capture_thread.join(timeout=1)

        if action == "refresh":
            if browser_mode:
                print("Refreshing browser credentials..." if settings.backend == "websocket" else "Reconnecting browser bridge...")
                continue
            print("Refreshing token...")
            if not _try_auto_refresh(cdp_port):
                print("Auto-refresh failed (Edge not running with --remote-debugging-port).")
                print("Falling back to manual mode.")
                set_token_command(None)
            # Reload credentials written by manual refresh before recreating the app.
            settings = Settings().model_copy(update=settings_updates)
            print("Restarting server...")
        else:
            break


if __name__ == "__main__":
    main()
