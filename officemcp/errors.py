"""Typed error hierarchy for OfficeMCP — architecture copied from
ltspice-mcp (not code). Provides structured error codes, recovery
hints, and actionable next steps for agent self-healing.

Pattern: each error type carries error_code + hint + recovery list.
Agent sees {"error": "...", "hint": "...", "recovery": [...]} instead
of bare False/""/None returns.
"""
from __future__ import annotations
from typing import Any


class OfficeError(Exception):
    """Base class for typed Office errors."""
    error_code: str = "office_error"
    hint: str = ""
    recovery: list[str] = []

    def __init__(self, message: str, *, details: dict[str, Any] | None = None) -> None:
        super().__init__(message)
        self.message = message
        self.details = details or {}

    def to_dict(self) -> dict[str, Any]:
        return {
            "error": self.error_code,
            "message": self.message,
            "hint": self.hint,
            "recovery": self.recovery,
            **self.details,
        }


class COMConnectionError(OfficeError):
    """COM automation layer unavailable (host app not running or not installed)."""
    error_code = "com_connection"
    hint = "Ensure the target Office application is installed and running on Windows."
    recovery = [
        "Call AvailableApps() to check installed applications",
        "Call Launch(app_name) to start the application",
        "Verify running on Windows with COM automation enabled",
    ]


class AppNotInstalledError(OfficeError):
    """Requested Office application is not installed."""
    error_code = "app_not_installed"
    hint = "The application is not registered in the Windows registry."
    recovery = [
        "Call AvailableApps() to list installed applications",
        "Install the required Office application",
        "Check if a WPS/Kingsoft alternative is available (Ket, Kwps, Kwpp)",
    ]


class AppNotRunningError(OfficeError):
    """Requested Office application is not currently running."""
    error_code = "app_not_running"
    hint = "Launch the application before performing operations on it."
    recovery = [
        "Call Launch(app_name) to start the application",
        "Call RunningApps() to check which apps are running",
    ]


class FileOperationError(OfficeError):
    """File read/write/save operation failed."""
    error_code = "file_operation"
    hint = "Check file path, permissions, and format."
    recovery = [
        "Call RootFolder() to verify the working directory",
        "Call IsFileExists(path) to check if the file exists",
        "Ensure the file is not open in another application",
    ]


class OfficePermissionError(OfficeError):
    """Access denied — file or COM object requires elevated permissions."""
    error_code = "permission_denied"
    hint = "Run as Administrator or check file permissions."
    recovery = [
        "Restart the MCP server with elevated privileges",
        "Check file/folder permissions in the RootFolder",
        "Ensure the Office application has write access to the target path",
    ]


class ScreenshotError(OfficeError):
    """Screen capture failed."""
    error_code = "screenshot_failed"
    hint = "Screen capture requires a visible desktop session."
    recovery = [
        "Ensure the application window is visible (Visible(app_name, True))",
        "Check if running in a headless environment (screenshots require a display)",
        "Verify PIL/Pillow is installed for image processing",
    ]


class TTSError(OfficeError):
    """Text-to-speech operation failed."""
    error_code = "tts_failed"
    hint = "SAPI speech synthesis requires Windows SAPI service."
    recovery = [
        "Ensure running on Windows with SAPI service available",
        "Check system audio output is configured",
        "Verify volume is not muted (Speak param volume > 0)",
    ]


# Hint registry — maps error_code to recovery guidance
# Agent can query this for self-healing without parsing exception messages
ERROR_HINTS: dict[str, dict[str, Any]] = {
    "com_connection": {
        "description": "COM automation unavailable",
        "primary_action": "Check Windows COM registration",
        "diagnostic_tool": "AvailableApps()",
    },
    "app_not_installed": {
        "description": "Office application not found in registry",
        "primary_action": "Verify installation or use alternative",
        "diagnostic_tool": "AvailableApps()",
    },
    "app_not_running": {
        "description": "Application process not found",
        "primary_action": "Launch the application first",
        "diagnostic_tool": "RunningApps()",
    },
    "file_operation": {
        "description": "File I/O failed",
        "primary_action": "Check path and permissions",
        "diagnostic_tool": "IsFileExists(path)",
    },
    "permission_denied": {
        "description": "Access denied",
        "primary_action": "Elevate privileges or fix permissions",
        "diagnostic_tool": "RootFolder()",
    },
    "screenshot_failed": {
        "description": "Screen capture failed",
        "primary_action": "Ensure visible desktop session",
        "diagnostic_tool": "Visible(app_name, True)",
    },
    "tts_failed": {
        "description": "Text-to-speech unavailable",
        "primary_action": "Check Windows SAPI service",
        "diagnostic_tool": "Speak('test')",
    },
}
