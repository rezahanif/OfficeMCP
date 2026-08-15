"""Manifest contract tests (audit §7).

Validates connectors/office/office-mcp/manifest.json against BOTH parsers
that matter:

1. Gateway runtime parser (apps/gateway/src/manifest.rs `Manifest`).
2. CP18 package manifest parser (packages/connector-package) which
   REQUIRES manifest_schema_version, package_format_version, platform —
   without them an installed connector is rejected (InvalidManifest) and
   can never run.

CP18 is a Rust parser; this test statically verifies the JSON fields it
requires. The gateway Rust test suite runs the real parser against this
same file (see gateway manifest tests).
"""
import json
from pathlib import Path

import pytest

CONNECTOR_ROOT = Path(__file__).resolve().parent.parent.parent

REQUIRED_GATEWAY_FIELDS = ["id", "name", "version", "runtime", "entry"]
REQUIRED_CP18_FIELDS = [
    "manifest_schema_version",
    "package_format_version",
    "platform",
]


@pytest.fixture(scope="module")
def manifest():
    with open(CONNECTOR_ROOT / "manifest.json") as f:
        return json.load(f)


def test_gateway_runtime_fields_present(manifest):
    for field in REQUIRED_GATEWAY_FIELDS:
        assert manifest.get(field), f"gateway Manifest requires non-empty {field!r}"


def test_cp18_fields_present(manifest):
    for field in REQUIRED_CP18_FIELDS:
        assert field in manifest, f"connector-package CP18 requires {field!r}"


def test_cp18_schema_versions(manifest):
    assert manifest["manifest_schema_version"] == 1
    assert manifest["package_format_version"] == 1


def test_cp18_platform(manifest):
    assert manifest["platform"]["os"] == "windows"
    assert manifest["platform"]["arch"] == "x64"


def test_stdio_transport(manifest):
    assert manifest["stdio"] is True


def test_env_var_contract(manifest):
    assert manifest["port_env_var"] == "MCP_PORT"
    assert manifest["token_env_var"] == "MCP_LICENSE_TOKEN"


def test_runtime_supported_by_pm(manifest):
    # gateway build_command maps python|python3 -> python3
    assert manifest["runtime"] in ("python", "python3")


def test_entry_is_single_entrypoint(manifest):
    assert manifest["entry"] == "run_server.py"


def test_extra_fields_are_tolerated(manifest):
    # platform_requirements is informational; the gateway serde struct
    # ignores unknown fields — it must not break discovery.
    assert "platform_requirements" in manifest
