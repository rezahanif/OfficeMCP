"""AiConnect gateway entrypoint for the OfficeMCP connector.

The Process Manager spawns `python3 run_server.py` (manifest entry) behind the
mcp-stdio-bridge. Adapter wiring happens in main() — license gate first
(fail-closed: invalid/missing token → LicenseError → nonzero exit → PM
crash/backoff), then central envelope wrap, then the upstream server.
"""
import os
import sys

try:
    from officemcp import aioconnect
except ImportError:  # standalone upstream run without the shared SDK
    aioconnect = None

from officemcp.OfficeMCP import mcp, RunOfficeMCP  # noqa: E402


def main() -> None:
    if aioconnect is not None:
        aioconnect.ensure_licensed()
        # FastMCP 3.x: envelope via the server's SUPPORTED middleware API
        # (tools stay untouched — envelope applied post-validation).
        # Fallback: legacy per-tool wrap for fastmcp <3.x.
        if not aioconnect.install_envelope_middleware(mcp):
            aioconnect.wrap_tools(mcp)
    RunOfficeMCP()


if __name__ == "__main__":
    if os.environ.get("FAKE_CONNECTOR_MODE", "") == "crash":
        sys.exit(1)  # test-only: PM crash/backoff exercises
    main()
