"""AiConnect: COM-free TheOfficer stand-in for headless/Linux fixture tests.

Activated by AICONNECT_FAKE_BRIDGE=1 (same pattern as the SAP2000 fake
bridge). Proves MCP server boot, tools/list, tool dispatch, and the
license/envelope adapter WITHOUT Windows or Office installed.
"""
import os

from officemcp.Officer import TheOfficer


class FakeApp:
    def __init__(self, name):
        self.Name = name
        self.Visible = True


class FakeOfficer(TheOfficer):
    """Overrides every COM-touching method with deterministic no-COM results."""

    _INSTALLED = ["Word", "Excel", "PowerPoint", "Outlook", "MSProject"]

    def AvailableApps(self) -> list:
        return list(self._INSTALLED)

    def RunningApps(self) -> list:
        return []

    def IsAppAvailable(self, app_name: str) -> bool:
        if app_name.endswith(".Application"):
            app_name = app_name[: -len(".Application")]
        return app_name in self._INSTALLED

    def Application(self, app_name: str, asNewInstance: bool = False):
        if not self.IsAppAvailable(app_name):
            return None
        app = FakeApp(app_name)
        self.__dict__["_" + app_name.lower()] = app
        return app

    def GetComObject(self, com_name: str, active: bool = True):
        return FakeApp(com_name)

    def Visible(self, app_name: str, visible=None) -> bool:
        return True

    def Quit(self, app_name: str, force: bool = False) -> bool:
        return True

    def Speak(self, text: str = "", volume: int = 80, rate: int = 2) -> bool:
        return True

    def Beep(self, frequency: int = 500, duration: int = 500) -> bool:
        return True

    def Demonstrate(self) -> str:
        return "fake demo succeeded"

    def ScreenShot(self, save_path: str = None) -> str:
        return os.path.join(self.RootFolder, save_path or "fake_screenshot.png")

    def DownloadImage(self, url: str, save_path: str = None) -> str:
        target = self.FilePath(save_path or "fake_download.png")
        os.makedirs(os.path.dirname(target) or self.RootFolder, exist_ok=True)
        with open(target, "wb") as f:
            f.write(b"fake image bytes")
        return target
