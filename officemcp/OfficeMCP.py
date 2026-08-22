# coding=utf-8
import os
import sys
from fastmcp import FastMCP
from fastmcp.resources import TextResource
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
    """Get Microsoft Office applications availability. """
    return Officer.AvailableApps()

@mcp.tool()
def RunningApps() -> list:
    """Get Microsoft Office applications availability. """
    return Officer.RunningApps()

@mcp.tool()
def IsAppAvailable(app_name: str = "Word") -> bool:
    """ Check if the specified application is installed."""
    return Officer.IsAppAvailable(app_name)


@mcp.tool()
def DownloadImage(url: str='https://www.bing.com/favicon.ico', save_path: str='favicon.ico') -> str:
    """ Download an image from the given URL and save it to the specified path."""
    Officer.Print(f'Tool.DownloadImage....{url}  to  {save_path}')
    path = Officer.DownloadImage(url, save_path)
    return path

@mcp.tool()
def RootFolder() -> str:
    """ return the default folder for this OfficeMCP server."""
    return Officer.RootFolder


@mcp.tool()
def Visible(app_name: str="Word", visible: bool = True) -> bool:
    """ Check if the microsoft excel application is visible."""
    return Officer.Visible(app_name,visible)

@mcp.tool()
def Launch(app_name: str ="Word", visible: bool = True)->dict:
    """ Launch a new Microsoft Office application or use the existing one.
    
    Args:
        app_name: Office application name (Word, Excel, PowerPoint, etc.)
        visible: Whether to show the application window
    
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
def ScreenShot(save_path: str = None) -> dict:
    """Capture a screenshot of the entire screen.
    
    Args:
        save_path: Optional file path to save the screenshot
    
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
def IsFileExists(sub_file_path: str) -> bool:
    return Officer.IsFileExists(sub_file_path)

@mcp.tool()
def Quit(app_name: str="Word",force:bool=False)->bool:
    """ Quit the microsoft excel application."""
    _log('Tool.Quit:')
    return Officer.Quit(app_name,force)

@mcp.tool()
def Speak(text: str = "I'm office mcp server , how are you", volume: int = 80, rate: int = 0)->dict:
    """ Speak the text using Windows SAPI.
    
    Args:
        text: Text to speak
        volume: Volume level (0-100)
        rate: Speech rate (-10 to 10)
    
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
def get_error_hints(error_code: str = None) -> dict:
    """Get recovery hints for a specific error code or all error types.
    
    Args:
        error_code: Optional error code to get hints for (e.g. "com_connection")
    
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
def Beep(frequency:int=500,duration:int=500)->bool:
    """ Beep the computer. frequency range is 37 to 32767, duration range is 0 to 65535."""
    _log('Tool.Beep:')
    return Officer.Beep(frequency, duration)
    
@mcp.tool()
def Demonstrate()->dict:
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
def doc_create(path: str) -> dict:
    """Create a new empty Word document (.docx).

    path: File path relative to the connector root folder (or absolute).
    Works on any OS — no Office installation required.
    """
    return _docs.doc_create(path)

@mcp.tool()
def doc_read(path: str) -> dict:
    """Read all paragraph texts and tables from a Word document (.docx).

    Returns {paragraphs: [str], tables: [[[str]]], counts}. Cross-platform.
    """
    return _docs.doc_read(path)

@mcp.tool()
def doc_add_paragraph(path: str, text: str, style: str | None = None) -> dict:
    """Append a paragraph to a Word document (.docx).

    style: Optional style name like "Normal", "Quote", "List Bullet".
    """
    return _docs.doc_add_paragraph(path, text, style)

@mcp.tool()
def doc_add_heading(path: str, text: str, level: int = 1) -> dict:
    """Add a heading to a Word document (.docx).

    level: Heading level 0-4 (0 = Title, 1 = Heading 1, ...).
    """
    return _docs.doc_add_heading(path, text, level)

@mcp.tool()
def doc_add_table(path: str, rows: list[list[str]], headers: list[str] | None = None) -> dict:
    """Append a table to a Word document (.docx).

    rows: 2D array of cell strings.
    headers: Optional first-row header labels (prepended to rows).
    """
    return _docs.doc_add_table(path, rows, headers)

@mcp.tool()
def doc_replace_text(path: str, find: str, replace: str) -> dict:
    """Replace all occurrences of find with replace across paragraphs and table cells.

    Returns the number of replacements made. Handles both simple runs and
    split-run paragraphs (Word splits text arbitrarily across XML runs).
    """
    return _docs.doc_replace_text(path, find, replace)

@mcp.tool()
def doc_get_properties(path: str) -> dict:
    """Read core properties of a Word document (.docx): title, author, dates, counts."""
    return _docs.doc_get_properties(path)


# --- Excel (.xlsx) ---

@mcp.tool()
def xlsx_create(path: str) -> dict:
    """Create a new Excel workbook (.xlsx) with one default sheet. Cross-platform."""
    return _docs.xlsx_create(path)

@mcp.tool()
def xlsx_read_cells(path: str, sheet: str | None = None, cell_range: str | None = None) -> dict:
    """Read cell values from an Excel workbook (.xlsx).

    sheet: Sheet name (default: first sheet).
    cell_range: Optional A1-style range ("A1:C10") or single cell ("B2").
                Omit to read all non-empty rows.
    """
    return _docs.xlsx_read_cells(path, sheet, cell_range)

@mcp.tool()
def xlsx_write_cells(path: str, sheet: str, start_cell: str, data: list[list[_Any]]) -> dict:
    """Write a 2D array of values starting at start_cell on the named sheet.

    start_cell: A1-style anchor ("A1"). Creates the sheet if missing.
    data: 2D array; numbers stay numbers, strings stay strings.
    """
    return _docs.xlsx_write_cells(path, sheet, start_cell, data)

@mcp.tool()
def xlsx_list_sheets(path: str) -> dict:
    """List all sheet names in an Excel workbook (.xlsx)."""
    return _docs.xlsx_list_sheets(path)

@mcp.tool()
def xlsx_add_sheet(path: str, name: str) -> dict:
    """Add a new sheet to an Excel workbook (.xlsx). Fails if the name exists."""
    return _docs.xlsx_add_sheet(path, name)

@mcp.tool()
def xlsx_append_rows(path: str, sheet: str, rows: list[list[_Any]]) -> dict:
    """Append rows at the end of a sheet (creates the sheet if missing)."""
    return _docs.xlsx_append_rows(path, sheet, rows)

@mcp.tool()
def xlsx_get_properties(path: str) -> dict:
    """Read workbook metadata: title, creator, created date, sheet list."""
    return _docs.xlsx_get_properties(path)


# --- PowerPoint (.pptx) ---

@mcp.tool()
def pptx_create(path: str) -> dict:
    """Create a new presentation (.pptx) with one blank title slide. Cross-platform."""
    return _docs.pptx_create(path)

@mcp.tool()
def pptx_read(path: str) -> dict:
    """Extract all text from every slide of a presentation (.pptx).

    Returns [{index, texts: [str]}] per slide, including table cell text.
    """
    return _docs.pptx_read(path)

@mcp.tool()
def pptx_add_slide(path: str, title: str, content: str | None = None) -> dict:
    """Add a slide with a title and optional body content to a presentation (.pptx).

    content: Body paragraph text. Placed in the layout's body placeholder,
    or a new textbox if the layout has none.
    """
    return _docs.pptx_add_slide(path, title, content)

@mcp.tool()
def pptx_get_info(path: str) -> dict:
    """Read presentation metadata: title, author, slide count, dimensions (EMU)."""
    return _docs.pptx_get_info(path)

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
