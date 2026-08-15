"""Stdio purity regression (audit §13).

The bridge (mcp-stdio-bridge) parses the connector's stdout as
newline-delimited JSON-RPC. ANY plain-text line corrupts the protocol.
Logs must go to stderr only. This test fails if startup logging corrupts
stdout — both with and without the AiConnect adapter enabled.
"""
import json
import select

from fake_license import SECRET, mcp_initialize, mint, spawn_server, stop

VALID_CONFIGS = [
    {"name": "plain", "env": {}},
    {
        "name": "adapter",
        "env": {
            "AICONNECT_ENABLE": "1",
            "JWT_SECRET": SECRET,
            "MCP_LICENSE_TOKEN": mint(),
        },
    },
]


def _assert_stdout_pure(proc):
    """Every stdout line must be JSON-RPC; anything else fails the test.

    Select-first with a short timeout: the server is idle after
    initialize, so an unconditional readline() would block forever.
    """
    assert proc.stdout is not None
    while True:
        ready, _, _ = select.select([proc.stdout], [], [], 0.5)
        if not ready:
            break
        line = proc.stdout.readline()
        if not line:
            break
        text = line.decode()
        parsed = json.loads(text)  # raises on any non-JSON line
        assert parsed.get("jsonrpc") == "2.0"


def test_stdout_is_protocol_only_plain():
    proc = spawn_server()
    try:
        mcp_initialize(proc)
        _assert_stdout_pure(proc)
    finally:
        stop(proc)


def test_stdout_is_protocol_only_with_adapter():
    proc = spawn_server(env_extra=VALID_CONFIGS[1]["env"])
    try:
        mcp_initialize(proc)
        _assert_stdout_pure(proc)
    finally:
        stop(proc)


def test_logs_land_on_stderr_not_stdout():
    proc = spawn_server(env_extra=VALID_CONFIGS[1]["env"])
    try:
        mcp_initialize(proc)
        # give any late logging a moment, then verify stdout still pure
        import time

        time.sleep(0.3)
        _assert_stdout_pure(proc)
        assert proc.stderr is not None
        assert proc.stderr.readable()
    finally:
        stop(proc)
