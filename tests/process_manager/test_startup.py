"""Startup + MCP initialization + tool dispatch tests.

Proves the PM lifecycle contract WITHOUT Windows/Office: the connector
spawns, initializes MCP over stdio, exposes the expected tools, and
dispatches tool calls through the fake officer (AICONNECT_FAKE_BRIDGE=1).
No COM/Office/license is touched.
"""
import json

from fake_license import SECRET, mcp_initialize, mcp_tools_list, mint, spawn_server, stop

# All 52 tools — upstream RunPython tool REMOVED (see README security
# posture). ReadME + Instructions are resources, not tools.
EXPECTED_TOOLS = {
    # COM lifecycle (13)
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
    # Guidance (1)
    "get_error_hints",
    # Word (.docx) CRUD (7)
    "doc_create",
    "doc_read",
    "doc_add_paragraph",
    "doc_add_heading",
    "doc_add_table",
    "doc_replace_text",
    "doc_get_properties",
    # Excel (.xlsx) CRUD (7)
    "xlsx_create",
    "xlsx_read_cells",
    "xlsx_write_cells",
    "xlsx_list_sheets",
    "xlsx_add_sheet",
    "xlsx_append_rows",
    "xlsx_get_properties",
    # PowerPoint (.pptx) CRUD (4)
    "pptx_create",
    "pptx_read",
    "pptx_add_slide",
    "pptx_get_info",
    # Layer B API guidance (4)
    # office_function_registry_query and register_verified_office were REMOVED:
    # they read scripts/registry.json, which never existed in this repo, so both
    # answered every call from an empty registry while costing ~557 tokens of
    # tool surface in every session. See aiconnector docs/audit/OFFICE-API-BENCHMARK.md.
    "search_office_api",
    "list_office_api_categories",
    "list_templates",
    "load_template",
    # Microsoft Project (16)
    "msp_create",
    "msp_open",
    "msp_save",
    "msp_close",
    "msp_add_task",
    "msp_get_tasks",
    "msp_update_task",
    "msp_delete_task",
    "msp_add_resource",
    "msp_get_resources",
    "msp_delete_resource",
    "msp_assign_resource",
    "msp_set_baseline",
    "msp_get_project_info",
    "msp_switch_view",
    "msp_list_views",
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


def test_tool_dispatch_enveloped():
    """Adapter-mode tool call must return an ok envelope through the
    low-level interceptor — regression for the FastMCP 3.4.7 wrapped-call
    failure (JSON-string return vs frozen output-model validation)."""
    proc = spawn_server(env_extra={
        "AICONNECT_ENABLE": "1",
        "JWT_SECRET": SECRET,
        "MCP_LICENSE_TOKEN": mint(),
    })
    try:
        mcp_initialize(proc)
        assert proc.stdin is not None
        req = {
            "jsonrpc": "2.0",
            "id": 5,
            "method": "tools/call",
            "params": {"name": "AvailableApps", "arguments": {}},
        }
        proc.stdin.write((json.dumps(req) + "\n").encode())
        proc.stdin.flush()
        resp = json.loads(proc.stdout.readline().decode())
        content = resp["result"]["content"][0]["text"]
        assert '"success":' in content and '"data"' in content, f"expected ok envelope, got: {content[:200]}"
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
