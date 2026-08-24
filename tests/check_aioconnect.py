"""Adapter unit validation — runs WITHOUT mcp/pywin32/Office (python3.11).

Validates the AiConnect adapter layer (license gate + envelope wrap) that the
real server calls from run_server.py. All OfficeMCP tools are SYNC defs, so
the sync wrapper path is exercised explicitly (FastMCP freezes is_async at
registration — an async wrapper on a sync tool leaks a coroutine).
"""
import asyncio
import base64
import hashlib
import hmac
import json
import os
import sys
import time
from pathlib import Path

FORK = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(FORK))

# AiConnect SDK is external (not vendored — IP boundary). Dev default points
# at the aiconnector monorepo; override with AICONNECT_SDK_PATH.
os.environ.setdefault("AICONNECT_SDK_PATH", os.environ.get("AICONNECT_SDK_PATH") or str(FORK.parent / "connector-sdk" / "python"))


def asyncio_run(coro):
    return asyncio.get_event_loop().run_until_complete(coro)

from officemcp import aioconnect  # noqa: E402
from mcp_license_sdk import LicenseError  # noqa: E402

SECRET = "0123456789abcdef0123456789abcdef"
os.environ["JWT_SECRET"] = SECRET


def mint(entitlements, ttl=600, subject=None):
    def b64(b):
        return base64.urlsafe_b64encode(b).rstrip(b"=")
    header = b64(json.dumps({"alg": "HS256", "typ": "JWT"}).encode())
    payload = b64(json.dumps({
        "sub": subject or "connector:office-mcp", "iat": int(time.time()), "exp": int(time.time()) + ttl,
        "entitlements": entitlements,
    }).encode())
    sig = b64(hmac.new(SECRET.encode(), header + b"." + payload, hashlib.sha256).digest())
    return f"{header.decode()}.{payload.decode()}.{sig.decode()}"


results = []
def check(name, cond):
    results.append((name, bool(cond)))
    print(("PASS" if cond else "FAIL"), name)


# 1. disabled (AICONNECT_ENABLE unset) → everything is a no-op
os.environ.pop("AICONNECT_ENABLE", None)
os.environ.pop("MCP_LICENSE_TOKEN", None)
aioconnect.ensure_licensed()  # must NOT raise
check("disabled: ensure_licensed no-op without env", True)
check("disabled: wrap_tools wraps 0", aioconnect.wrap_tools(object()) == 0)

# 2. enabled + missing token → refuse at startup
os.environ["AICONNECT_ENABLE"] = "1"
try:
    aioconnect.ensure_licensed()
    check("enabled: missing token refuses", False)
except LicenseError:
    check("enabled: missing token refuses", True)

# 3. enabled + valid token → passes; wrong subject → refuses; envelope works
os.environ["MCP_LICENSE_TOKEN"] = mint(["office-mcp"])
aioconnect.ensure_licensed()
check("enabled: valid token passes", True)

os.environ["MCP_LICENSE_TOKEN"] = mint(["office-mcp"], subject="connector:other-mcp")
try:
    aioconnect.ensure_licensed()
    check("enabled: wrong subject refuses", False)
except LicenseError:
    check("enabled: wrong subject refuses", True)
os.environ["MCP_LICENSE_TOKEN"] = mint(["office-mcp"])

env = json.loads(aioconnect._wrap_result('{"a":1}'))
check("envelope: JSON → ok(data)", env.get("success") is True and env.get("data", {}).get("a") == 1)
env = json.loads(aioconnect._wrap_result("raw boom"))
check("envelope: non-JSON → fail TOOL_ERROR", env.get("success") is False and env.get("error", {}).get("code") == "TOOL_ERROR")
env = json.loads(aioconnect._wrap_result(""))
check("envelope: empty → ok", env.get("success") is True)

# 4. per-call recheck: expired token rejected inside the wrapped call (async)
os.environ["MCP_LICENSE_TOKEN"] = mint(["office-mcp"], ttl=-400000)
async def fake_tool():
    return '{"ok": true}'
wrapped = aioconnect._wrap(fake_tool)
out = asyncio_run(wrapped())
env = json.loads(out)
check("per-call recheck: expired token → fail envelope", env.get("success") is False)

# 5. SYNC tool wrapping: envelope on success, fail envelope on exception
os.environ["MCP_LICENSE_TOKEN"] = mint(["office-mcp"])
def sync_ok():
    return '{"ok": true}'
wrapped_sync = aioconnect._wrap(sync_ok)
env = json.loads(wrapped_sync())
check("sync tool: wrapped call returns envelope", env.get("success") is True)

def sync_boom():
    raise RuntimeError("COM not reachable")
wrapped_sync = aioconnect._wrap(sync_boom)
env = json.loads(wrapped_sync())
check("sync tool exception → fail envelope", env.get("success") is False and "COM not reachable" in env.get("error", {}).get("message", ""))

# 6. wrap_tools on a fake registry wraps all tools exactly once
class FakeTool:
    def __init__(self, fn):
        self.fn = fn
class FakeManager:
    def __init__(self):
        self._tools = {"t1": FakeTool(lambda: '{"x":1}'), "t2": FakeTool(lambda: "y")}
mgr = FakeManager()
n = aioconnect.wrap_tools(mgr)
check("wrap_tools wraps all", n == 2)
check("wrap_tools idempotent", aioconnect.wrap_tools(mgr) == 0)
env = json.loads(mgr._tools["t1"].fn())
check("wrapped sync tool returns envelope", env.get("success") is True)

failed = [n for n, ok in results if not ok]
print(f"\n{len(results) - len(failed)}/{len(results)} adapter checks passed")
sys.exit(1 if failed else 0)
