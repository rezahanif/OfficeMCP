# coding=utf-8
import json
import os
import sys
from pathlib import Path
from typing import Annotated, Literal

from fastmcp import FastMCP
from fastmcp.resources import TextResource
from pydantic import Field
from officemcp.Officer import TheOfficer
from officemcp.errors import (
    COMConnectionError, AppNotInstalledError, AppNotRunningError,
    FileOperationError, ScreenshotError, TTSError, OfficeError, ERROR_HINTS,
)


def _log(*args):
    # AiConnect: stdout is reserved for the JSON-RPC stdio protocol (bridge
    # parses it line-by-line). ALL diagnostics go to stderr.
    print(*args, file=sys.stderr)


def _create_officer():
    # AiConnect: test hook — COM-free officer for headless/Linux fixture
    # validation (same pattern as SAP2000 AICONNECT_FAKE_BRIDGE).
    if os.environ.get("AICONNECT_FAKE_BRIDGE", "") == "1":
        from officemcp.fake_officer import FakeOfficer
        return FakeOfficer()
    return TheOfficer()


mcp = FastMCP("OfficeMCP")
mcp.isrunnning = False
mcp.Officer = _create_officer()
Officer = mcp.Officer

# region Office Apps ------------------------------

@mcp.tool()
def AvailableApps() -> list:
    """Office applications INSTALLED on this machine, whether running or not."""
    return Officer.AvailableApps()

@mcp.tool()
def RunningApps() -> list:
    """Office applications currently OPEN, a subset of AvailableApps."""
    return Officer.RunningApps()

@mcp.tool()
def IsAppAvailable(
    app_name: Annotated[
        Literal["Word", "Excel", "PowerPoint", "Outlook", "MSProject", "Access"],
        Field(description="Office application name to check. Requires Windows + the app installed (COM)."),
    ] = "Word",
) -> bool:
    """Check if the specified Office application is installed."""
    return Officer.IsAppAvailable(app_name)

@mcp.tool()
def DownloadImage(
    url: Annotated[
        str,
        Field(description="URL of the image to download."),
    ] = "https://www.bing.com/favicon.ico",
    save_path: Annotated[
        str,
        Field(description="Local file path to save the downloaded image."),
    ] = "favicon.ico",
) -> str:
    """ Download an image from the given URL and save it to the specified path."""
    Officer.Print(f'Tool.DownloadImage....{url}  to  {save_path}')
    path = Officer.DownloadImage(url, save_path)
    return path

@mcp.tool()
def RootFolder() -> str:
    """ return the default folder for this OfficeMCP server."""
    return Officer.RootFolder


@mcp.tool()
def Visible(
    app_name: Annotated[
        Literal["Word", "Excel", "PowerPoint", "Outlook", "MSProject", "Access"],
        Field(description="Office application to show or hide. Requires Windows + the app running (COM)."),
    ] = "Word",
    visible: Annotated[
        bool,
        Field(description="True shows the window, False hides it."),
    ] = True,
) -> bool:
    """Show or hide an Office application window."""
    return Officer.Visible(app_name,visible)

@mcp.tool()
def Launch(
    app_name: Annotated[
        Literal["Word", "Excel", "PowerPoint", "Outlook", "MSProject", "Access"],
        Field(description="Office application to launch. Requires Windows + the app installed (COM)."),
    ] = "Word",
    visible: Annotated[
        bool,
        Field(description="Whether to show the application window."),
    ] = True,
) -> dict:
    """Launch an Office application or attach to the running instance.

    Returns:
        dict with success status and details

    Raises:
        AppNotInstalledError: if the application is not installed
        COMConnectionError: if COM automation is unavailable
    """
    Officer.Print('Tool.Launch....')
    try:
        if not Officer.IsAppAvailable(app_name):
            raise AppNotInstalledError(
                f"{app_name} is not installed",
                details={"available_apps": Officer.AvailableApps()},
            )
        app = Officer.Application(app_name)
        app.Visible = visible
        Officer.Print('    Launched:')
        return {"success": True, "app": app_name, "visible": visible}
    except OfficeError:
        raise
    except Exception as e:
        raise COMConnectionError(
            f"Failed to launch {app_name}: {e}",
            details={"app_name": app_name},
        ) from e

@mcp.tool()
def ScreenShot(
    save_path: Annotated[
        str | None,
        Field(description="Optional file path to save the screenshot. None for auto-generated path."),
    ] = None,
) -> dict:
    """Capture a screenshot of the entire screen.

    Returns:
        dict with path to saved screenshot

    Raises:
        ScreenshotError: if screen capture fails
    """
    Officer.Print('Tool.ScreenShot....')
    try:
        path = Officer.ScreenShot(save_path)
        if not path:
            raise ScreenshotError("Screen capture returned empty path")
        Officer.Print(f'   saved to {path}: ')
        return {"success": True, "path": path}
    except OfficeError:
        raise
    except Exception as e:
        raise ScreenshotError(
            f"Screen capture failed: {e}",
            details={"save_path": save_path},
        ) from e

@mcp.resource("resource://README.md")
def ReadME() -> TextResource:
    """ Read the readme file."""
    return TextResource(Officer.FilePath("README.md"))

@mcp.tool()
def IsFileExists(
    sub_file_path: Annotated[
        str,
        Field(description="Relative file path to check for existence under RootFolder."),
    ],
) -> bool:
    """Check whether a file exists under the server's root folder."""
    return Officer.IsFileExists(sub_file_path)

@mcp.tool()
def Quit(
    app_name: Annotated[
        Literal["Word", "Excel", "PowerPoint", "Outlook", "MSProject", "Access"],
        Field(description="Office application to quit. Requires Windows + the app running (COM)."),
    ] = "Word",
    force: Annotated[
        bool,
        Field(description="True to close without saving prompts."),
    ] = False,
) -> bool:
    """Quit an Office application."""
    _log('Tool.Quit:')
    return Officer.Quit(app_name,force)

@mcp.tool()
def Speak(
    text: Annotated[
        str,
        Field(description="Text to speak via Windows SAPI."),
    ] = "I'm office mcp server , how are you",
    volume: Annotated[
        int,
        Field(description="Volume level (0-100)."),
    ] = 80,
    rate: Annotated[
        int,
        Field(description="Speech rate (-10 to 10)."),
    ] = 0,
) -> dict:
    """ Speak the text using Windows SAPI.

    Returns:
        dict with success status

    Raises:
        TTSError: if speech synthesis fails
    """
    _log('Tool.Speak:')
    try:
        result = Officer.Speak(text, volume, rate)
        if not result:
            raise TTSError("Speech synthesis returned false")
        return {"success": True, "text": text[:50]}
    except OfficeError:
        raise
    except Exception as e:
        raise TTSError(
            f"Speech synthesis failed: {e}",
            details={"text": text[:50], "volume": volume, "rate": rate},
        ) from e


@mcp.tool()
def get_error_hints(
    error_code: Annotated[
        str | None,
        Field(description='Optional error code to get hints for (e.g. "com_connection"). None returns all hints.'),
    ] = None,
) -> dict:
    """Get recovery hints for a specific error code or all error types.

    Returns:
        dict with error hints and recovery guidance
    """
    if error_code:
        hint = ERROR_HINTS.get(error_code)
        if hint:
            return {"error_code": error_code, **hint}
        return {"error_code": error_code, "hint": "No hints available for this error code"}
    return {"error_types": list(ERROR_HINTS.keys()), "hints": ERROR_HINTS}

@mcp.tool()
def Beep(
    frequency: Annotated[
        int,
        Field(description="Tone frequency in Hz (37-32767)."),
    ] = 500,
    duration: Annotated[
        int,
        Field(description="Duration in milliseconds (0-65535)."),
    ] = 500,
) -> bool:
    """Beep the PC speaker."""
    _log('Tool.Beep:')
    return Officer.Beep(frequency, duration)
    
@mcp.tool()
def Demonstrate() -> dict:
    """ Demonstrate for you to see some functions in this OfficeMCP server."""
    _log('Tool.Demonstrate:')
    output = ""
    try:
        output = Officer.Demonstrate()
        return {"success": True, "output": output}
    except Exception as e:
        _log(e)
        return {"success": False, "error": str(e), "output": output}

# endregion

# region Document CRUD (OOXML — cross-platform, no COM) ---------------
#
# AiConnect Phase 1: dedicated document manipulation tools built on
# python-docx / openpyxl / python-pptx. These read/write .docx/.xlsx/.pptx
# files directly (OOXML = ZIP+XML) and work on ANY OS without Office
# installed. Complements the 13 Windows-only COM lifecycle tools above.
# Security posture unchanged: pure library file operations, no exec.

from typing import Any as _Any

from officemcp import documents as _docs


# --- Word (.docx) ---

@mcp.tool()
def doc_create(
    path: Annotated[
        str,
        Field(description="File path relative to the connector root folder (or absolute) for the new .docx."),
    ],
) -> dict:
    """Create a new empty Word document (.docx).

    Works on any OS — no Office installation required.
    """
    return _docs.doc_create(path)

@mcp.tool()
def doc_read(
    path: Annotated[
        str,
        Field(description="Path to the .docx file to read."),
    ],
) -> dict:
    """Read all paragraph texts and tables from a Word document (.docx).

    Returns {paragraphs: [str], tables: [[[str]]], counts}. Cross-platform.
    """
    return _docs.doc_read(path)

@mcp.tool()
def doc_add_paragraph(
    path: Annotated[
        str,
        Field(description="Path to the .docx file to modify."),
    ],
    text: Annotated[
        str,
        Field(description="Paragraph text to append."),
    ],
    style: Annotated[
        str | None,
        Field(description='Optional style name like "Normal", "Quote", "List Bullet".'),
    ] = None,
) -> dict:
    """Append a paragraph to a Word document (.docx)."""
    return _docs.doc_add_paragraph(path, text, style)

@mcp.tool()
def doc_add_heading(
    path: Annotated[
        str,
        Field(description="Path to the .docx file to modify."),
    ],
    text: Annotated[
        str,
        Field(description="Heading text."),
    ],
    level: Annotated[
        int,
        Field(description="Heading level 0-4 (0 = Title, 1 = Heading 1, ...)."),
    ] = 1,
) -> dict:
    """Add a heading to a Word document (.docx)."""
    return _docs.doc_add_heading(path, text, level)

@mcp.tool()
def doc_add_table(
    path: Annotated[
        str,
        Field(description="Path to the .docx file to modify."),
    ],
    rows: Annotated[
        list[list[str]],
        Field(description="2D array of cell strings for table rows."),
    ],
    headers: Annotated[
        list[str] | None,
        Field(description="Optional first-row header labels (prepended to rows)."),
    ] = None,
) -> dict:
    """Append a table to a Word document (.docx)."""
    return _docs.doc_add_table(path, rows, headers)

@mcp.tool()
def doc_replace_text(
    path: Annotated[
        str,
        Field(description="Path to the .docx file to modify."),
    ],
    find: Annotated[
        str,
        Field(description="Text string to search for across paragraphs and table cells."),
    ],
    replace: Annotated[
        str,
        Field(description="Replacement text."),
    ],
) -> dict:
    """Replace all occurrences of find with replace across paragraphs and table cells.

    Returns the number of replacements made. Handles both simple runs and
    split-run paragraphs (Word splits text arbitrarily across XML runs).
    """
    return _docs.doc_replace_text(path, find, replace)

@mcp.tool()
def doc_get_properties(
    path: Annotated[
        str,
        Field(description="Path to the .docx file to inspect."),
    ],
) -> dict:
    """Read core properties of a Word document (.docx): title, author, dates, counts."""
    return _docs.doc_get_properties(path)


# --- Excel (.xlsx) ---

@mcp.tool()
def xlsx_create(
    path: Annotated[
        str,
        Field(description="File path for the new .xlsx workbook."),
    ],
) -> dict:
    """Create a new Excel workbook (.xlsx) with one default sheet. Cross-platform."""
    return _docs.xlsx_create(path)

@mcp.tool()
def xlsx_read_cells(
    path: Annotated[
        str,
        Field(description="Path to the .xlsx file to read."),
    ],
    sheet: Annotated[
        str | None,
        Field(description="Sheet name (default: first sheet)."),
    ] = None,
    cell_range: Annotated[
        str | None,
        Field(description='Optional A1-style range ("A1:C10") or single cell ("B2"). Omit to read all non-empty rows.'),
    ] = None,
) -> dict:
    """Read cell values from an Excel workbook (.xlsx)."""
    return _docs.xlsx_read_cells(path, sheet, cell_range)

@mcp.tool()
def xlsx_write_cells(
    path: Annotated[
        str,
        Field(description="Path to the .xlsx file to modify."),
    ],
    sheet: Annotated[
        str,
        Field(description="Sheet name to write to. Creates the sheet if missing."),
    ],
    start_cell: Annotated[
        str,
        Field(description='A1-style anchor (e.g. "A1") indicating where to start writing.'),
    ],
    data: Annotated[
        list[list[_Any]],
        Field(description="2D array of values; numbers stay numbers, strings stay strings."),
    ],
) -> dict:
    """Write a 2D array of values starting at start_cell on the named sheet."""
    return _docs.xlsx_write_cells(path, sheet, start_cell, data)

@mcp.tool()
def xlsx_list_sheets(
    path: Annotated[
        str,
        Field(description="Path to the .xlsx file to inspect."),
    ],
) -> dict:
    """List all sheet names in an Excel workbook (.xlsx)."""
    return _docs.xlsx_list_sheets(path)

@mcp.tool()
def xlsx_add_sheet(
    path: Annotated[
        str,
        Field(description="Path to the .xlsx file to modify."),
    ],
    name: Annotated[
        str,
        Field(description="Name for the new sheet. Fails if the name already exists."),
    ],
) -> dict:
    """Add a new sheet to an Excel workbook (.xlsx). Fails if the name exists."""
    return _docs.xlsx_add_sheet(path, name)

@mcp.tool()
def xlsx_append_rows(
    path: Annotated[
        str,
        Field(description="Path to the .xlsx file to modify."),
    ],
    sheet: Annotated[
        str,
        Field(description="Sheet name to append to. Creates the sheet if missing."),
    ],
    rows: Annotated[
        list[list[_Any]],
        Field(description="2D array of rows to append at the end of the sheet."),
    ],
) -> dict:
    """Append rows at the end of a sheet (creates the sheet if missing)."""
    return _docs.xlsx_append_rows(path, sheet, rows)

@mcp.tool()
def xlsx_get_properties(
    path: Annotated[
        str,
        Field(description="Path to the .xlsx file to inspect."),
    ],
) -> dict:
    """Read workbook metadata: title, creator, created date, sheet list."""
    return _docs.xlsx_get_properties(path)


# --- PowerPoint (.pptx) ---

@mcp.tool()
def pptx_create(
    path: Annotated[
        str,
        Field(description="File path for the new .pptx presentation."),
    ],
) -> dict:
    """Create a new presentation (.pptx) with one blank title slide. Cross-platform."""
    return _docs.pptx_create(path)

@mcp.tool()
def pptx_read(
    path: Annotated[
        str,
        Field(description="Path to the .pptx file to read."),
    ],
) -> dict:
    """Extract all text from every slide of a presentation (.pptx).

    Returns [{index, texts: [str]}] per slide, including table cell text.
    """
    return _docs.pptx_read(path)

@mcp.tool()
def pptx_add_slide(
    path: Annotated[
        str,
        Field(description="Path to the .pptx file to modify."),
    ],
    title: Annotated[
        str,
        Field(description="Slide title text."),
    ],
    content: Annotated[
        str | None,
        Field(description="Optional body paragraph text. Placed in layout body placeholder or a new textbox."),
    ] = None,
) -> dict:
    """Add a slide with a title and optional body content to a presentation (.pptx)."""
    return _docs.pptx_add_slide(path, title, content)

@mcp.tool()
def pptx_get_info(
    path: Annotated[
        str,
        Field(description="Path to the .pptx file to inspect."),
    ],
) -> dict:
    """Read presentation metadata: title, author, slide count, dimensions (EMU)."""
    return _docs.pptx_get_info(path)

# endregion

# region Layer B — API guidance tools (AiConnect Phase 1) ---------------
#
# Fallback path: when no dedicated tool covers the intent, the agent searches the
# Office API docs and adapts a template. There is deliberately no exec hatch —
# RunPython was removed (D1, 2026-08-15) — so discovery ends at documentation,
# not at arbitrary execution.

from officemcp.doc_search import doc_index as _doc_index  # noqa: E402

_TEMPLATES_PATH = Path(__file__).resolve().parent.parent / "templates" / "templates.json"


def _load_templates() -> dict:
    try:
        return json.loads(Path(_TEMPLATES_PATH).read_text(encoding="utf-8")).get("templates", {})
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


@mcp.tool()
def search_office_api(
    query: Annotated[
        str,
        Field(description='What you need (e.g. "open workbook", "add slide", "replace text").'),
    ],
    category: Annotated[
        str | None,
        Field(description="Optional category filter from list_office_api_categories results."),
    ] = None,
) -> list[dict]:
    """Search the Office API documentation for objects/methods matching a query.

    Use BEFORE writing custom automation: returns object names, syntax,
    descriptions, and examples for both the COM model and cross-platform
    OOXML libraries.
    """
    return _doc_index.search(query=query, category=category)


@mcp.tool()
def list_office_api_categories() -> list[dict]:
    """List available Office API documentation categories with section counts.

    Returns [{category, sections}] — e.g. Excel Worksheet Object (5).
    """
    return _doc_index.list_categories()


@mcp.tool()
def list_templates() -> list[dict]:
    """List ready-to-run Office document templates (report skeleton, data table,
    outline deck, find/replace...). Pair with load_template for full code."""
    return [
        {"id": tid, "name": t["name"], "description": t["description"], "category": t["category"]}
        for tid, t in sorted(_load_templates().items())
    ]


@mcp.tool()
def load_template(
    template_id: Annotated[
        str,
        Field(description="Template identifier from list_templates (e.g. 'report-skeleton')."),
    ],
) -> dict:
    """Load one template's full tool-call sequence by id (from list_templates).

    The sequences are verified — adapt file names to yours.
    """
    t = _load_templates().get(template_id)
    if t is None:
        return {"error": f"Unknown template: {template_id}", "available": sorted(_load_templates())}
    return {"id": template_id, **t}

# endregion

# region Microsoft Project tools ------------------------------------------

import officemcp.project as _proj


@mcp.tool()
def msp_create(
    path: Annotated[
        str,
        Field(description="File path for the new .mpp project file."),
    ],
) -> dict:
    """Create a new blank Microsoft Project file (.mpp)."""
    return _proj.msp_create(path)


@mcp.tool()
def msp_open(
    path: Annotated[
        str,
        Field(description="Path to an existing .mpp project file to open."),
    ],
) -> dict:
    """Open an existing .mpp project file."""
    return _proj.msp_open(path)


@mcp.tool()
def msp_save(
    path: Annotated[
        str | None,
        Field(description="Optional path for SaveAs. None saves to current location."),
    ] = None,
) -> dict:
    """Save the active project. Pass path to SaveAs."""
    return _proj.msp_save(path)


@mcp.tool()
def msp_close(
    save: Annotated[
        bool,
        Field(description="True to save before closing."),
    ] = True,
) -> dict:
    """Close the active project."""
    return _proj.msp_close(save)


@mcp.tool()
def msp_add_task(
    name: Annotated[
        str,
        Field(description="Task name."),
    ],
    duration: Annotated[
        str | None,
        Field(description='Optional duration string (e.g. "3d", "1w", "4h").'),
    ] = None,
    start: Annotated[
        str | None,
        Field(description="Optional start date string."),
    ] = None,
    predecessors: Annotated[
        str | None,
        Field(description='Optional predecessor references (e.g. "1,3FS+2d").'),
    ] = None,
    notes: Annotated[
        str | None,
        Field(description="Optional task notes."),
    ] = None,
) -> dict:
    """Add an activity to the project schedule, with duration and predecessors.

    The unit of work in a plan: an activity, task or job with a duration."""
    return _proj.msp_add_task(name, duration, start, predecessors, notes)


@mcp.tool()
def msp_get_tasks(
    filter_name: Annotated[
        str | None,
        Field(description='Optional filter (e.g. "Incomplete Tasks").'),
    ] = None,
    field: Annotated[
        str | None,
        Field(description="Optional field name to sort by."),
    ] = None,
    max_results: Annotated[
        int,
        Field(description="Maximum number of tasks to return."),
    ] = 50,
) -> dict:
    """List the activities in the project schedule — the work in the plan.

    Optional filter and sort field. Returns tasks with durations and dates."""
    return _proj.msp_get_tasks(filter_name, field, max_results)


@mcp.tool()
def msp_update_task(
    task_id: Annotated[
        int,
        Field(description="Unique ID of the task to update."),
    ],
    name: Annotated[
        str | None,
        Field(description="New task name."),
    ] = None,
    duration: Annotated[
        str | None,
        Field(description='New duration string (e.g. "5d").'),
    ] = None,
    start: Annotated[
        str | None,
        Field(description="New start date string."),
    ] = None,
    finish: Annotated[
        str | None,
        Field(description="New finish date string."),
    ] = None,
    percent_complete: Annotated[
        int | None,
        Field(description="Completion percentage (0-100)."),
    ] = None,
    predecessors: Annotated[
        str | None,
        Field(description='New predecessor references (e.g. "1FS+2d").'),
    ] = None,
    notes: Annotated[
        str | None,
        Field(description="New task notes."),
    ] = None,
) -> dict:
    """Update a task by UniqueID. Only provided fields are changed."""
    return _proj.msp_update_task(
        task_id, name, duration, start, finish, percent_complete, predecessors, notes
    )


@mcp.tool()
def msp_delete_task(
    task_id: Annotated[
        int,
        Field(description="Unique ID of the task to delete."),
    ],
) -> dict:
    """Delete a task by UniqueID."""
    return _proj.msp_delete_task(task_id)


@mcp.tool()
def msp_add_resource(
    name: Annotated[
        str,
        Field(description="Resource name."),
    ],
    resource_type: Annotated[
        str,
        Field(description="Resource type: 'work', 'material', or 'cost'."),
    ] = "work",
    email: Annotated[
        str | None,
        Field(description="Optional email for work resources."),
    ] = None,
) -> dict:
    """Add a resource."""
    return _proj.msp_add_resource(name, resource_type, email)


@mcp.tool()
def msp_get_resources(
    max_results: Annotated[
        int,
        Field(description="Maximum number of resources to return."),
    ] = 50,
) -> dict:
    """List resources in the active project."""
    return _proj.msp_get_resources(max_results)


@mcp.tool()
def msp_delete_resource(
    resource_id: Annotated[
        int,
        Field(description="Unique ID of the resource to delete."),
    ],
) -> dict:
    """Delete a resource by UniqueID."""
    return _proj.msp_delete_resource(resource_id)


@mcp.tool()
def msp_assign_resource(
    task_id: Annotated[
        int,
        Field(description="Unique ID of the task to assign to."),
    ],
    resource_id: Annotated[
        int,
        Field(description="Unique ID of the resource to assign."),
    ],
    units: Annotated[
        int,
        Field(description="Allocation percentage: 100=full time, 50=half time."),
    ] = 100,
) -> dict:
    """Assign a resource to a task."""
    return _proj.msp_assign_resource(task_id, resource_id, units)


@mcp.tool()
def msp_set_baseline(
    baseline_number: Annotated[
        int,
        Field(description="Baseline slot (0-10). 0=Baseline, 1=Baseline 1, etc."),
    ] = 0,
) -> dict:
    """Save a baseline (0-10)."""
    return _proj.msp_set_baseline(baseline_number)


@mcp.tool()
def msp_get_project_info() -> dict:
    """Get info about the active project (name, tasks, resources, dates)."""
    return _proj.msp_get_project_info()


@mcp.tool()
def msp_switch_view(
    view_name: Annotated[
        str,
        Field(description='View name (e.g. "Gantt Chart", "Resource Sheet").'),
    ],
) -> dict:
    """Switch to a named view."""
    return _proj.msp_switch_view(view_name)


@mcp.tool()
def msp_list_views() -> dict:
    """List available views in the active project."""
    return _proj.msp_list_views()

# endregion

def RunOfficeMCP() -> None:
    r"""OfficeMCP server entry point with command line arguments support.
    Usage examples:
    1. OfficeMCP (stdio mode by default)
    2. OfficeMCP sse --port 8080 --host 127.0.0.1 --folder d:\@OfficeMCP (alternative syntax)
    """
    import sys as _sys
    args = _sys.argv[1:]
    transport = "stdio"
    thePort = 8888
    theHost = "127.0.0.1"
    theFolder = os.environ.get("OFFICE_MCP_ROOT") or os.path.join(os.path.expanduser("~"), "@OfficeMCP")

    # 更健壮的参数解析
    i = 1
    if args:
        transport = args[0].lower()
        while i < len(args):
            arg = args[i]
            if arg == '--port' and i+1 < len(args):
                try:
                    thePort = int(args[i+1])
                except Exception:
                    pass
                i += 2
            elif arg == '--host' and i+1 < len(args):
                theHost = args[i+1]
                i += 2
            elif arg == '--folder' and i+1 < len(args):
                theFolder = args[i+1]
                i += 2
            elif arg.isdigit():
                thePort = int(arg)
                i += 1
            elif not arg.startswith('--') and theHost == "127.0.0.1":
                # 只有没有明确指定 --host 时才用
                theHost = arg
                i += 1
            else:
                i += 1

    # 校验文件夹路径
    if not theFolder or not isinstance(theFolder, str) or not os.path.isabs(theFolder):
        theFolder = os.path.join(os.path.expanduser("~"), "@OfficeMCP")
    if not os.path.exists(theFolder):
        try:
            os.makedirs(theFolder)
        except Exception:
            _log(f"Warning: Could not create folder {theFolder}, fallback to {os.path.join(os.path.expanduser('~'), '@OfficeMCP')}")
            theFolder = os.path.join(os.path.expanduser("~"), "@OfficeMCP")
            if not os.path.exists(theFolder):
                os.makedirs(theFolder)
    mcp.Officer._default_folder = theFolder

    try:
        if transport == "stdio":
            _log(f"OfficeMCP running in stdio mode, root folder: {theFolder}")
            mcp.run("stdio")
        else:
            _log(f"OfficeMCP running in SSE mode, host: {theHost}, port: {thePort}, root folder: {theFolder}")
            mcp.run(transport="sse", host=theHost, port=thePort)
    except ValueError as e:
        _log(f"OfficeMCP Error parsing arguments: {e}")
    except Exception as e:
        _log(f"OfficeMCP Server startup failed: {e}")
