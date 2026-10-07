"""Optional real-Edge test against a local fixture, without contacting Copilot."""
import asyncio
import os
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import httpx
import pytest
import websockets

from m365_copilot_openai_proxy.browser_client import _CDPPage, _try_verification_checkbox


@pytest.mark.skipif(os.environ.get("M365_TEST_CDP") != "1", reason="Requires debug Edge")
def test_verification_click_uses_the_correct_cross_origin_frame_coordinates():
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args):
            pass

        def do_GET(self):
            if self.path == "/frame":
                html = '<label><input type="checkbox">Verify you are human</label>'
            else:
                html = f'<div role="dialog" aria-label="Security check"><iframe src="http://localhost:{self.server.server_port}/frame" style="position:absolute;left:100px;top:200px;width:320px;height:100px"></iframe></div>'
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            self.wfile.write(html.encode())

    async def run(server):
        base = "http://127.0.0.1:9222"
        async with httpx.AsyncClient(trust_env=False) as http:
            tab = (await http.put(base + f"/json/new?http://127.0.0.1:{server.server_port}")).json()
            try:
                async with websockets.connect(tab["webSocketDebuggerUrl"]) as ws:
                    page = _CDPPage(ws, auto_verify=True)
                    frame = None
                    for _ in range(30):
                        targets = (await page.call("Target.getTargets"))["targetInfos"]
                        frame = next((t for t in targets if t.get("type") == "iframe" and t.get("url") == f"http://localhost:{server.server_port}/frame"), None)
                        if frame:
                            break
                        await asyncio.sleep(0.2)
                    assert frame, "Edge did not create the local cross-origin iframe"

                    class FixturePage:
                        verification_attempted = False

                        async def call(self, method, params=None, **kwargs):
                            if method == "Target.getTargets":
                                # Only this test's local iframe is eligible. No
                                # production origin rule is relaxed or changed.
                                return {"targetInfos": [{**frame, "url": "https://challenges.cloudflare.com/fixture"}]}
                            return await page.call(method, params, **kwargs)

                    assert await _try_verification_checkbox(FixturePage())
                    attached = await page.call("Target.attachToTarget", {"targetId": frame["targetId"], "flatten": True})
                    session = attached["sessionId"]
                    result = await page.call("Runtime.evaluate", {"expression": "document.querySelector('input').checked", "returnByValue": True}, session_id=session)
                    assert result["result"]["value"] is True
                    await page.call("Target.detachFromTarget", {"sessionId": session})
            finally:
                await http.get(base + "/json/close/" + tab["id"])

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        asyncio.run(run(server))
    finally:
        server.shutdown()
        server.server_close()
