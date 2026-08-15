"""Environment + entitlement injection tests (audit §8, §9, §22).

Covers missing / invalid / valid / wrong-subject tokens with a fake
validator (minted locally, no production license service). Also the
test-only crash hook (audit §16) and clean-environment startup.
"""
import os
import subprocess

from fake_license import (
    RUN_SERVER,
    SECRET,
    mcp_initialize,
    mint,
    spawn_server,
    stop,
)

VALID_TOKEN = mint()
INVALID_TOKEN = mint(secret="other-secret")
WRONG_SUBJECT = mint(subject="connector:other-mcp")


def _exit_code(proc):
    code = stop(proc)
    return code


def test_missing_token_fails_closed():
    proc = spawn_server(env_extra={"AICONNECT_ENABLE": "1"})
    code = _exit_code(proc)
    assert code != 0, "missing token must fail closed (adapter enabled)"


def test_invalid_token_fails_closed():
    proc = spawn_server(
        env_extra={"AICONNECT_ENABLE": "1", "MCP_LICENSE_TOKEN": INVALID_TOKEN},
    )
    code = _exit_code(proc)
    assert code != 0, "invalid signature must fail closed"


def test_wrong_subject_fails_closed():
    proc = spawn_server(
        env_extra={"AICONNECT_ENABLE": "1", "MCP_LICENSE_TOKEN": WRONG_SUBJECT},
    )
    code = _exit_code(proc)
    assert code != 0, "token minted for another connector must fail closed"


def test_valid_token_starts_and_serves():
    proc = spawn_server(
        env_extra={
            "AICONNECT_ENABLE": "1",
            "JWT_SECRET": SECRET,
            "MCP_LICENSE_TOKEN": VALID_TOKEN,
        },
    )
    try:
        init = mcp_initialize(proc)
        assert init.get("result", {}).get("serverInfo", {}).get("name") == "OfficeMCP"
    finally:
        assert _exit_code(proc) == 0


def test_clean_environment_startup():
    """No developer shell PATH/IDE/SAP2000 required for PM lifecycle (§22).

    Minimal env: PATH + HOME (python needs HOME for its user site-packages
    layout — that is the runtime install, not a developer extra). No
    PYTHONPATH, no VIRTUAL_ENV, no SAP2000, no IDE vars.
    """
    env = {
        "PATH": "/usr/local/bin:/usr/bin:/bin",
        "HOME": os.environ.get("HOME", "/tmp"),
        "AICONNECT_FAKE_BRIDGE": "1",
    }
    proc = subprocess.Popen(
        ["/usr/local/bin/python3", str(RUN_SERVER)],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
    )
    try:
        mcp_initialize(proc)
    finally:
        stop(proc)


def test_crash_hook_exits_nonzero():
    proc = spawn_server(env_extra={"FAKE_CONNECTOR_MODE": "crash"})
    code = _exit_code(proc)
    assert code != 0, "FAKE_CONNECTOR_MODE=crash must exit non-zero for backoff tests"
