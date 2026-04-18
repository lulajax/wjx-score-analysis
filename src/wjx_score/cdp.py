"""CDP (Chrome DevTools Protocol) websocket 通信"""

import json
import asyncio
import time
import urllib.request

import websockets

_msg_counter = 1000


async def cdp_send(ws, method, params=None, timeout=15):
    global _msg_counter
    _msg_counter += 1
    mid = _msg_counter
    await ws.send(json.dumps({"id": mid, "method": method, **({"params": params} if params else {})}))
    dl = time.time() + timeout
    while time.time() < dl:
        resp = json.loads(await asyncio.wait_for(ws.recv(), timeout=max(dl - time.time(), 1)))
        if resp.get("id") == mid:
            return resp
    raise TimeoutError(f"CDP timeout: {method}")


async def cdp_eval(ws, expr, timeout=15):
    r = (await cdp_send(ws, "Runtime.evaluate",
         {"expression": expr, "returnByValue": True, "awaitPromise": True}, timeout)
        ).get("result", {}).get("result", {})
    return r.get("value", "") if r.get("type") == "string" else r.get("value")


async def wait_ready(ws, sel, max_wait=20):
    for _ in range(max_wait):
        try:
            if await cdp_eval(ws, "document.readyState", 5) == "complete":
                if (await cdp_eval(ws, f"document.querySelectorAll('{sel}').length", 5) or 0) > 0:
                    return True
        except Exception:
            pass
        await asyncio.sleep(1)
    return False


async def connect(host="localhost", port=9222):
    """连接到 Chrome CDP，返回 websocket 连接"""
    tabs = json.loads(urllib.request.urlopen(f"http://{host}:{port}/json").read())
    pts = [t for t in tabs if t.get("type") == "page"]
    if not pts:
        raise RuntimeError("需要至少 1 个浏览器标签页")
    ws = await websockets.connect(pts[0]["webSocketDebuggerUrl"], max_size=10 * 1024 * 1024, ping_interval=None)
    await cdp_send(ws, "Page.enable")
    return ws
