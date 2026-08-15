# coding=utf-8
import os
import sys
from fastmcp import FastMCP
from fastmcp.resources import TextResource
from officemcp.Officer import TheOfficer


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
def Launch(app_name: str ="Word", visible: bool = True)->bool:
    """ Launch an new microsoft excel application or use the existed one."""
    Officer.Print('Tool.Launch....')
    try:
        app = Officer.Application(app_name)
        app.Visible = visible
        Officer.Print('    Launched:')
        return True
    except Exception as e:
        Officer.Print(f'    failed{e}:')
        return False

@mcp.tool()
def ScreenShot(save_path: str = None) -> str:
    """ Launch an new microsoft excel application or use the existed one."""
    Officer.Print('Tool.ScreenShot....')
    try:
        path = Officer.ScreenShot(save_path)
        Officer.Print(f'   saved to {path}: ')
        return path
    except Exception as e:
        Officer.Print(f'   failed {e}: ')
        return ""

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
def Speak(text: str = "I'm office mcp server , how are you", volume: int = 80, rate: int = 0)->bool:
    """ Speak the text. volume range is 0-100, rate range is -10 to 10."""
    _log('Tool.Speak:')
    return Officer.Speak(text, volume, rate)

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

@mcp.resource("resource://Instructions")
def Instructions() -> str:
    return """
    There're some base tools for you to control Microsoft applications.
    Use tool AvailableApps / RunningApps / IsAppAvailable to check applications.
    Use tool Launch / Visible / Quit to control application lifecycle.
    Use tool ScreenShot to capture the screen.
    Use tool Speak / Beep for simple feedback.
    """

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
