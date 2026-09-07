# Changelog — OfficeMCP

## v1.0.6 (2026-09-07)

### Added
- Complete rewrite of TUTORIAL.md covering all 3 operational modes (Pure OOXML, live COM automation, MS Project), all 52 tool signatures and parameters, multi-client setups, and end-to-end workflows.
- Explicit package includes for `templates/` and `API/` in `manifest.json`.

### Fixed
- Removed stray `D:\@OfficeMCP` directory to enforce path hygiene.

## v1.0.5 (2026-08-24)

### Added
- MIT LICENSE file covering AiConnect original work (documents, project tools, adapter, Layer B).
- `Annotated[Field(description=...)]` on all 97 tool parameters for full schema completeness.
- pydantic added as explicit dependency in pyproject.toml.
- This CHANGELOG.

### Changed
- fastmcp pinned to `<3.0.0` upper bound in pyproject.toml.
- `PermissionError` renamed to `OfficePermissionError` in errors.py (no longer shadows builtin).
- DEFINITION_QUALITY.md re-scored: 82/100 (up from 73.8) after schema and test fixes.
- test_startup.py EXPECTED_TOOLS updated from 13 to 54 (all registered tools).

### Upstream
- COM lifecycle wrappers (13 tools) remain under upstream permission (LICENSE_PERMISSION.md).
