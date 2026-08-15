"""Authoritative-installation boundary tests (audit §19, §20, §21).

The connector must NEVER modify active.json, install metadata, or attempt
to reinstall itself — the Process Manager owns installation state. These
tests run the real entrypoint against a fake CP22-style install tree and
assert the tree is byte-identical afterwards.
"""
import json
import os
import time
from pathlib import Path

from fake_license import mcp_initialize, spawn_server, stop


def _make_installed_tree(root: Path) -> Path:
    conn = root / "connectors" / "office-mcp"
    ver = conn / "1.0.5"
    ver.mkdir(parents=True)
    (conn / "active.json").write_text(
        json.dumps({"connector_id": "office-mcp", "version": "1.0.5"}),
    )
    (ver / ".aiconnect-install.json").write_text(
        json.dumps({
            "connector_id": "office-mcp",
            "version": "1.0.5",
            "platform_os": "windows",
            "platform_arch": "x64",
        }),
    )
    (ver / "manifest.json").write_text(
        json.dumps({
            "id": "office-mcp",
            "name": "Office Connector",
            "version": "1.0.5",
            "manifest_schema_version": 1,
            "package_format_version": 1,
            "platform": {"os": "windows", "arch": "x64"},
            "runtime": "python",
            "entry": "run_server.py",
            "stdio": True,
        }),
    )
    return conn


def _snapshot(root: Path):
    out = {}
    for p in sorted(root.rglob("*")):
        if p.is_file():
            out[str(p.relative_to(root))] = (p.read_bytes(), p.stat().st_mtime_ns)
    return out


def test_connector_never_touches_installed_state(tmp_path):
    conn = _make_installed_tree(tmp_path)
    before = _snapshot(tmp_path)

    proc = spawn_server(env_extra={"AICONNECT_CONNECTOR_ID": "office-mcp"})
    try:
        mcp_initialize(proc)
        time.sleep(0.3)  # let any (mis)behavior manifest
    finally:
        stop(proc)

    after = _snapshot(tmp_path)
    assert after == before, "connector modified the installation tree (active.json / metadata)"
    assert (conn / "active.json").read_text().startswith('{"connector_id": "office-mcp"')
