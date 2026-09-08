"""AiConnect gateway entrypoint for the OfficeMCP connector.

The Process Manager spawns `python3 run_server.py` (manifest entry) behind the
mcp-stdio-bridge. Adapter wiring happens in main() — license gate first
(fail-closed: invalid/missing token → LicenseError → nonzero exit → PM
crash/backoff), then central envelope wrap, then the upstream server.
"""
import os
import sys
from pathlib import Path

# Vendored dependencies (`stage-python-vendor.py`), shipped inside the
# package. The connector used to ship source-only, so `from officemcp.OfficeMCP
# import mcp` below failed with `ModuleNotFoundError: No module named 'fastmcp'`
# on any machine without fastmcp/pydantic/pywin32 already installed
# system-wide. AI CONNECT bundles the INTERPRETER; the connector brings its
# own LIBRARIES, and this is where they are. Inserted (not appended) ahead of
# the pywin32 subpaths specifically — pywin32's vendored layout splits across
# win32/, win32/lib/ and pythonwin/, none of which sit on sys.path by default.
_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(_ROOT))
_VENDOR = _ROOT / "_vendor"
if _VENDOR.is_dir():
    sys.path.insert(0, str(_VENDOR))
    sys.path.insert(0, str(_VENDOR / "win32"))
    sys.path.insert(0, str(_VENDOR / "win32" / "lib"))
    sys.path.insert(0, str(_VENDOR / "pythonwin"))
    if sys.platform == "win32":
        pywin32_sys32 = _VENDOR / "pywin32_system32"
        if pywin32_sys32.is_dir() and hasattr(os, "add_dll_directory"):
            try:
                os.add_dll_directory(str(pywin32_sys32))
            except Exception:
                pass

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
