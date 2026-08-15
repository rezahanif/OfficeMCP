"""AiConnect adapter for the OfficeMCP fork (integration layer — the 11
upstream tools are NOT per-tool edited; RunPython was REMOVED from the MCP
surface — see README "Security posture").

Reuses the shared Python SDK (connectors/sdk/python):
  1. License gate — startup + per-call check of the token the Process Manager
     injects via MCP_LICENSE_TOKEN (manifest token_env_var).
  2. Response envelope — every registered tool's return is wrapped centrally
     at registration time (ok/fail), so no tool needs per-tool edits.

Env-gated integration points:
  AICONNECT_ENABLE=1        — install the license gate + envelope wrap
  MCP_LICENSE_TOKEN         — the token to validate
  JWT_SECRET                — token signing secret (default matches gateway dev)

Security posture (D1 decision, 2026-08-15): the upstream RunPython tool
(unrestricted exec with full COM Officer in scope) is REMOVED. Only curated
lifecycle/query tools remain — no arbitrary Python execution.
"""
import asyncio
import functools
import json
import os
import sys
from pathlib import Path

# SDK resolution: AICONNECT_SDK_PATH env wins (installed AiConnect SDK,
# keeps the public fork IP-free); else monorepo-relative fallback
# (parents[3]/sdk/python = connectors/<category>/<id>/officemcp/aioconnect.py).
_env_sdk = os.environ.get("AICONNECT_SDK_PATH", "")
_SDK = Path(_env_sdk).resolve() if _env_sdk else Path(__file__).resolve().parents[3] / "sdk" / "python"
if str(_SDK) not in sys.path:
    sys.path.insert(0, str(_SDK))

from mcp_license_sdk import LicenseError, LicenseValidator, fail, ok  # noqa: E402

CONNECTOR_ID = "office-mcp"


def _enabled() -> bool:
    return os.environ.get("AICONNECT_ENABLE", "") == "1"


def _validate() -> dict:
    """Validate the PM-injected token AND its connector binding.

    The Process Manager mints tokens with subject `connector:<id>` and
    entitlements `[<id>]` (auth.rs::mint). Signature + expiry are checked
    by the SDK; binding is asserted here so a token minted for another
    connector can never authorize this one.
    """
    claims = LicenseValidator(os.environ.get("JWT_SECRET", "dev-secret-change-me")).ensure_licensed()
    if claims.get("sub") != f"connector:{CONNECTOR_ID}":
        raise LicenseError(f"token not bound to {CONNECTOR_ID}")
    scopes = claims.get("entitlements") or claims.get("scopes") or []
    if CONNECTOR_ID not in scopes:
        raise LicenseError(f"token lacks scope {CONNECTOR_ID}")
    return claims


def ensure_licensed() -> None:
    if not _enabled():
        return
    _validate()


def _wrap_result(r):
    if isinstance(r, str):
        text = r.strip()
        if text:
            try:
                return json.dumps(ok(json.loads(text)))
            except json.JSONDecodeError:
                return json.dumps(fail("TOOL_ERROR", "non-JSON tool output"))
        return json.dumps(ok({"result": ""}))
    return json.dumps(ok(r))


async def _call(fn, args, kwargs):
    if asyncio.iscoroutinefunction(fn):
        return await fn(*args, **kwargs)
    return fn(*args, **kwargs)


def _make_sync_wrapper(fn):
    """For SYNC-registered tools: FastMCP freezes is_async at registration, so
    replacing tool.fn with an async wrapper makes FastMCP call it synchronously
    and leak a coroutine. Sync tools get a sync wrapper instead.
    functools.wraps preserves the ORIGINAL signature via __wrapped__ — FastMCP
    3.x derives input model/ctx injection from inspect.signature."""
    @functools.wraps(fn)
    def _w(*args, **kwargs):
        if not _enabled():
            return fn(*args, **kwargs)
        try:
            _validate()  # per-call recheck
            return _wrap_result(fn(*args, **kwargs))
        except LicenseError as e:
            return json.dumps(fail("LICENSE", str(e)))
        except Exception as e:
            return json.dumps(fail("TOOL_ERROR", str(e)))
    return _w


def _wrap(fn):
    if not asyncio.iscoroutinefunction(fn):
        return _make_sync_wrapper(fn)

    @functools.wraps(fn)
    async def _w(*args, **kwargs):
        if not _enabled():
            return await _call(fn, args, kwargs)
        try:
            _validate()  # per-call recheck
            result = await _call(fn, args, kwargs)
            return _wrap_result(result)
        except LicenseError as e:
            return json.dumps(fail("LICENSE", str(e)))
        except Exception as e:
            return json.dumps(fail("TOOL_ERROR", str(e)))
    return _w


def wrap_tools(mcp) -> int:
    """Legacy per-tool fn swap (FastMCP <3.x / fastmcp 1.x fallback).

    Do NOT use on FastMCP 3.4.7: replacing tool.fn keeps the frozen
    fn_metadata output model, so a wrapped JSON-string return fails
    structured-output validation (DictModel input_type=str). Prefer
    install_call_interceptor on 3.x.
    """
    if not _enabled():
        return 0
    registry = None
    for candidate in (getattr(mcp, "_tool_manager", None), getattr(mcp, "_tools", None)):
        if candidate is None:
            continue
        reg = getattr(candidate, "_tools", None) or getattr(candidate, "tools", None)
        if isinstance(reg, dict):
            registry = reg
            break
        if isinstance(candidate, dict) and candidate:
            registry = candidate
            break
    if registry is None:
        print("aioconnect: tool manager not found — envelope wrap skipped", file=sys.stderr)
        return 0
    wrapped = 0
    for name, tool in list(registry.items()):
        fn = getattr(tool, "fn", None) or tool
        if fn is None or getattr(fn, "_aioconnect_wrapped", False):
            continue
        wrapped_fn = _wrap(fn)
        wrapped_fn._aioconnect_wrapped = True
        if hasattr(tool, "fn"):
            tool.fn = wrapped_fn
        else:
            registry[name] = wrapped_fn
        wrapped += 1
    return wrapped


def install_envelope_middleware(mcp) -> bool:
    """Install a tools/call envelope middleware via FastMCP's SUPPORTED API.

    This FastMCP class (fastmcp.server.server.FastMCP 3.x) exposes a first-
    class middleware system (mcp.add_middleware). Our middleware:
      1. per-call license recheck + binding (fail → LICENSE envelope),
      2. delegates to the next handler (full validation/conversion — tools
         stay untouched: fn/signature/schema intact),
      3. envelopes the ToolResult content (ok/fail) — replacing tool.fn
         here breaks on FastMCP 3.4.7 (JSON-string return vs frozen
         output-model validation).

    Returns True when installed; False when unavailable/disabled — caller
    falls back to wrap_tools (fastmcp <3.x / mcp-SDK FastMCP without this
    middleware API use the low-level call_tool interceptor instead).
    """
    if not _enabled():
        return False
    if not hasattr(mcp, "add_middleware"):
        print("aioconnect: add_middleware not available — middleware skipped", file=sys.stderr)
        return False

    from fastmcp.server.middleware import Middleware
    from fastmcp.tools.base import ToolResult
    from mcp.types import CallToolRequestParams, TextContent

    class _EnvelopeMiddleware(Middleware):
        async def on_call_tool(self, context: CallToolRequestParams, call_next) -> ToolResult:
            try:
                _validate()  # per-call recheck + binding
            except LicenseError as e:
                return ToolResult(content=[TextContent(type="text", text=json.dumps(fail("LICENSE", str(e))))])
            try:
                result = await call_next(context)
            except Exception as e:
                return ToolResult(content=[TextContent(type="text", text=json.dumps(fail("TOOL_ERROR", str(e))))])
            text = ""
            for block in (result.content or []):
                if getattr(block, "type", None) == "text":
                    text = getattr(block, "text", "") or ""
                    break
            if not text:
                payload = result.structured_content
                text = json.dumps(payload if payload is not None else result, default=str)
            return ToolResult(
                content=[TextContent(type="text", text=_wrap_result(text))],
                structured_content=result.structured_content,
            )

    mcp.add_middleware(_EnvelopeMiddleware())
    return True
