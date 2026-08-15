"""Startup + MCP initialization + tool dispatch tests.

Proves the PM lifecycle contract WITHOUT Windows/Office: the connector
spawns, initializes MCP over stdio, exposes the expected tools, and
dispatches tool calls through the fake officer (AICONNECT_FAKE_BRIDGE=1).
No COM/Office/license is touched.
"""
import json

from fake_license import mcp_initialize, mcp_tools_list, spawn_server, stop

# 13 curated tools — upstream RunPython tool REMOVED (see README security
# posture). ReadME + Instructions are resources, not tools.
EXPECTED_TOOLS = {
    "AvailableApps",
    "RunningApps",
    "IsAppAvailable",
    "DownloadImage",
    "RootFolder",
    "Visible",
    "Launch",
    "ScreenShot",
    "IsFileExists",
    "Quit",
    "Speak",
    "Beep",
    "Demonstrate",
}


def test_starts_without_office_and_initializes():
    proc = spawn_server()
    try:
        init = mcp_initialize(proc)
        assert init.get("result", {}).get("serverInfo", {}).get("name") == "OfficeMCP"
    finally:
        stop(proc)


def test_exposes_expected_tools():
    proc = spawn_server()
    try:
        mcp_initialize(proc)
        result = mcp_tools_list(proc)
        names = {t["name"] for t in result.get("tools", [])}
        assert names == EXPECTED_TOOLS, f"missing: {EXPECTED_TOOLS - names}"
    finally:
        stop(proc)


def test_tool_dispatch_through_fake_officer():
    """AvailableApps + Launch work through the fake officer — proves MCP
    server + tool registration + dispatch without Office (COM-free)."""
    proc = spawn_server()
    try:
        mcp_initialize(proc)
        assert proc.stdin is not None
        req = {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {"name": "AvailableApps", "arguments": {}},
        }
        proc.stdin.write((json.dumps(req) + "\n").encode())
        proc.stdin.flush()
        resp = json.loads(proc.stdout.readline().decode())
        content = resp["result"]["content"][0]["text"]
        assert "Excel" in content and "Word" in content

        req = {
            "jsonrpc": "2.0",
            "id": 4,
            "method": "tools/call",
            "params": {"name": "Launch", "arguments": {"app_name": "Excel"}},
        }
        proc.stdin.write((json.dumps(req) + "\n").encode())
        proc.stdin.flush()
        resp = json.loads(proc.stdout.readline().decode())
        content = resp["result"]["content"][0]["text"]
        assert "true" in content
    finally:
        stop(proc)


def test_runpython_not_in_surface():
    """Security regression: RunPython must NOT be registered."""
    proc = spawn_server()
    try:
        mcp_initialize(proc)
        result = mcp_tools_list(proc)
        names = {t["name"] for t in result.get("tools", [])}
        assert "RunPython" not in names
    finally:
        stop(proc)


def test_exits_cleanly_on_stdin_close():
    proc = spawn_server()
    mcp_initialize(proc)
    code = stop(proc)
    assert code == 0, f"expected clean exit, got {code}"
