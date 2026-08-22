# Office Connector (office-mcp) — AiConnect adaptation of OfficeMCP

Adapted from upstream [OfficeMCP/OfficeMCP](https://github.com/OfficeMCP/OfficeMCP)
(commit `188140dc`, v1.0.5, fork `rezahanif/OfficeMCP`).

> **LICENSE — DISTRIBUTION BLOCKED.** Upstream has **no LICENSE file**. Local
> rework and internal fixture/testing use only. Public/customer redistribution
> is BLOCKED until an upstream license/redistribution grant is obtained
> (author: youngfe@live.com). Decision 2026-08-15 (D2).

## Classification (AiConnect compatibility v2)

- **Family**: COM / LOCAL RUNTIME (same shape as sap2000-mcp)
- **Platform**: Windows only (COM via pywin32) · manifest `platform: windows/x64`
- **Transport**: stdio (`stdio: true`) via mcp-stdio-bridge
- **host_plugin**: none — connector talks COM directly to Office apps
- **Gateway/PM changes**: NONE — connector fits the existing contract

## Security posture (D1, 2026-08-15)

Upstream `RunPython` (unrestricted `exec()` with builtins + full COM `Officer`
in scope) is **REMOVED from the MCP surface**. The connector exposes only the
curated lifecycle/query tools plus cross-platform document CRUD:

**COM lifecycle (Windows-only, 13 tools):**
`AvailableApps, RunningApps, IsAppAvailable, DownloadImage, RootFolder,
Visible, Launch, ScreenShot, IsFileExists, Quit, Speak, Beep, Demonstrate`

**Document CRUD (cross-platform OOXML, 18 tools):**
`doc_create, doc_read, doc_add_paragraph, doc_add_heading, doc_add_table,
doc_replace_text, doc_get_properties` (python-docx) ·
`xlsx_create, xlsx_read_cells, xlsx_write_cells, xlsx_list_sheets,
xlsx_add_sheet, xlsx_append_rows, xlsx_get_properties` (openpyxl) ·
`pptx_create, pptx_read, pptx_add_slide, pptx_get_info` (python-pptx)

**API guidance (Layer B, 7 tools):**
`search_office_api, list_office_api_categories, office_function_registry_query,
register_verified_office, list_templates, load_template, get_error_hints`

No arbitrary execution anywhere. Document CRUD is pure library file
operations on OOXML files (ZIP+XML) — no COM, no Office install required.

## AiConnect integration

- `manifest.json` — CP18 package fields + gateway runtime fields; stdio:true.
- `run_server.py` — gateway entrypoint; license gate first (fail-closed:
  invalid/missing token → nonzero exit → PM crash/backoff), then central
  envelope wrap, then upstream `mcp.run(stdio)`.
- `officemcp/aioconnect.py` — adapter: `AICONNECT_ENABLE=1` activates license
  gate (token binding: `sub == connector:office-mcp` + entitlement) + ok/fail
  response envelope on every tool; standalone run stays plain upstream.
- Lazy imports — `pywintypes`/`win32com.client`/`winreg` load inside methods;
  server boots headless on non-Windows (tools fail with structured TOOL_ERROR).
- `AICONNECT_FAKE_BRIDGE=1` — COM-free `FakeOfficer` for headless fixture tests.

## Upstream fixes included

- stdout `print()` pollution → stderr (stdio protocol purity)
- `Beep` passed duration as frequency
- `DownloadImage` returned None
- `Demonstrate` NameError on failure path
- `Launch` param typo `visilbe`
- hardcoded `D:\@OfficeMCP` default folder → `OFFICE_MCP_ROOT` env / `~/@OfficeMCP`

## Validation status

| Suite | Status |
|---|---|
| Adapter unit (`tests/check_aioconnect.py`) | VERIFIED on Linux sandbox |
| PM harness (`tests/process_manager/`) | VERIFIED on Linux sandbox (fake officer) |
| Real Office COM round-trip | **BLOCKED** — requires Windows + Office |
| Gateway fixture round-trip | pending — gateway `manifest.rs` has in-flight Agent 2 WIP |

## Layout

```
manifest.json          gateway/CP18 manifest contract
run_server.py          gateway entrypoint (adapter wiring)
officemcp/             adapted upstream package
  OfficeMCP.py         MCP server (RunPython removed, stderr logging, fixes)
  Officer.py           COM layer (lazy Windows imports)
  aioconnect.py        AiConnect adapter (license + envelope)
  fake_officer.py      COM-free test double
tests/
  check_aioconnect.py  adapter unit checks (no pytest)
  process_manager/     PM compatibility harness (fake_license + 5 suites)
```
