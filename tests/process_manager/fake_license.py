"""Shared helpers for the Process Manager compatibility suite.

Mints real HS256 JWTs with the same claim shape as apps/gateway/src/auth.rs
(audit §8: missing / invalid / valid token, no external service). Also
provides a minimal MCP-over-stdio client for lifecycle tests.
"""
import base64
import hashlib
import hmac
import json
import os
import subprocess
import sys
import time
from pathlib import Path

CONNECTOR_ROOT = Path(__file__).resolve().parent.parent.parent
RUN_SERVER = CONNECTOR_ROOT / "run_server.py"
SECRET = "0123456789abcdef0123456789abcdef"

CONNECTOR_ID = "office-mcp"

# The shared AiConnect SDK is NOT vendored in this repo (IP boundary). Point
# AICONNECT_SDK_PATH at the installed SDK; dev default = aiconnector monorepo.
DEFAULT_SDK = "/project/aiconnector/connectors/sdk/python"


def mint(entitlements=None, subject=None, ttl=600, secret=SECRET):
    """Mint a token exactly like gateway auth.rs::mint (HS256)."""
    def b64(b):
        return base64.urlsafe_b64encode(b).rstrip(b"=")

    now = int(time.time())
    header = b64(json.dumps({"alg": "HS256", "typ": "JWT"}).encode())
    payload = b64(json.dumps({
        "sub": subject or f"connector:{CONNECTOR_ID}",
        "iat": now,
        "exp": now + ttl,
        "entitlements": entitlements or [CONNECTOR_ID],
    }).encode())
    sig = b64(hmac.new(secret.encode(), header + b"." + payload, hashlib.sha256).digest())
    return f"{header.decode()}.{payload.decode()}.{sig.decode()}"


def spawn_server(env_extra=None, interpreter=None):
    """Spawn run_server.py the way the Process Manager does.

    The gateway spawn inherits the parent environment (tokio Command
    default), so this harness inherits os.environ and overrides what the
    test needs. The connector must not depend on IDE/developer extras —
    the clean-environment case is covered explicitly in test_env.py.
    """
    env = dict(os.environ)
    env["PYTHONPATH"] = str(CONNECTOR_ROOT)
    env.setdefault("AICONNECT_SDK_PATH", DEFAULT_SDK)
    env["AICONNECT_FAKE_BRIDGE"] = "1"
    if env_extra:
        env.update(env_extra)
    return subprocess.Popen(
        [interpreter or sys.executable, str(RUN_SERVER)],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
    )


def recv_json(proc, timeout=15.0):
    """Read one newline-delimited JSON-RPC message from the child stdout."""
    assert proc.stdout is not None
    line = proc.stdout.readline()
    if not line:
        return None
    return json.loads(line.decode())


def mcp_initialize(proc):
    """Perform MCP initialize + notifications/initialized, return result."""
    assert proc.stdin is not None
    req = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "pm-compat-test", "version": "0"},
        },
    }
    proc.stdin.write((json.dumps(req) + "\n").encode())
    proc.stdin.flush()
    resp = recv_json(proc)
    assert resp and resp.get("id") == 1, f"no initialize response: {resp}"
    proc.stdin.write((json.dumps({
        "jsonrpc": "2.0",
        "method": "notifications/initialized",
    }) + "\n").encode())
    proc.stdin.flush()
    return resp


def mcp_tools_list(proc):
    """List tools after initialize; returns the result payload."""
    assert proc.stdin is not None
    req = {"jsonrpc": "2.0", "id": 2, "method": "tools/list", "params": {}}
    proc.stdin.write((json.dumps(req) + "\n").encode())
    proc.stdin.flush()
    resp = recv_json(proc)
    assert resp and resp.get("id") == 2, f"no tools/list response: {resp}"
    return resp.get("result", {})


def stop(proc):
    """Close stdin (EOF → FastMCP stdio exits cleanly), wait, return code."""
    try:
        if proc.stdin:
            proc.stdin.close()
    except BrokenPipeError:
        pass
    try:
        return proc.wait(timeout=10)
    except subprocess.TimeoutExpired:
        proc.kill()
        return proc.wait()
