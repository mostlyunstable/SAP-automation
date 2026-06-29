"""Windows Save dialog automation for SAP export workflows.

SAP GUI Scripting cannot control the Windows Save dialog.
This module uses Windows COM automation (WScript.Shell) to:

1. Detect the Save dialog
2. Enter the filename and path
3. Handle overwrite confirmation
4. Verify the file was saved

The approach is based on the proven pattern used in enterprise
SAP automation: AppActivate to find the dialog, SendKeys to
fill the filename and trigger Save.

Windows-only -- requires pywin32.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import Any

from ..core.exceptions import ExportError
from ..core.logger import get_logger

log = get_logger("windows_dialog")

# Lazy import -- win32com is only available on Windows
_win32com: Any = None


def _get_win32com() -> Any:
    """Lazy-load win32com.client."""
    global _win32com
    if _win32com is None:
        try:
            import win32com.client  # type: ignore[import-untyped]
            _win32com = win32com.client
        except ImportError:
            raise ExportError(
                "pywin32 is required for Save dialog automation. "
                "Install it: pip install pywin32"
            ) from None
    return _win32com


class SaveDialogHandler:
    """Handles the Windows Save dialog that appears during SAP export.

    Usage:
        handler = SaveDialogHandler()
        handler.save_file(
            filename="VA23_1000012345.xlsx",
            folder="C:/output",
        )
    """

    def __init__(
        self,
        dialog_timeout: float = 15.0,
        key_delay: float = 0.1,
    ):
        """
        Args:
            dialog_timeout: Max seconds to wait for Save dialog to appear.
            key_delay: Delay between keystrokes in seconds.
        """
        self.dialog_timeout = dialog_timeout
        self.key_delay = key_delay

    def save_file(
        self,
        filename: str,
        folder: str,
        delete_existing: bool = True,
        overwrite_confirm: bool = True,
    ) -> Path:
        """Fill in the Windows Save dialog and save the file.

        Args:
            filename: Desired filename (e.g., "VA23_1000012345.xlsx").
            folder: Target folder path.
            delete_existing: If True, delete existing file before saving.
            overwrite_confirm: If True, handle overwrite confirmation dialog.

        Returns:
            Path to the saved file.

        Raises:
            ExportError: If the Save dialog cannot be found or file not saved.
        """
        folder_path = Path(folder)
        folder_path.mkdir(parents=True, exist_ok=True)
        full_path = folder_path / filename

        if delete_existing and full_path.exists():
            try:
                full_path.unlink()
                log.info("Deleted existing file: %s", full_path)
            except OSError as exc:
                log.warning("Could not delete existing file: %s", exc)

        file_path_str = str(full_path).replace("/", "\\")
        log.info("Save dialog: target=%s", file_path_str)

        shell = _get_win32com().Dispatch("WScript.Shell")

        self._wait_for_save_dialog(shell)
        self._fill_filename(shell, file_path_str)
        self._trigger_save(shell)

        if overwrite_confirm:
            self._handle_overwrite_confirm(shell)

        self._wait_for_file(full_path)

        log.info("File saved: %s", full_path)
        return full_path

    def close_excel_window(self, timeout: float = 5.0) -> None:
        """Close the Excel window that SAP opens after xlsx export.

        SAP opens the exported file in Excel automatically.
        This method closes that window to prevent file locks.
        """
        shell = _get_win32com().Dispatch("WScript.Shell")

        # Try common Excel window title patterns
        excel_titles = ["Microsoft Excel", "Excel"]
        for title in excel_titles:
            if shell.AppActivate(title):
                time.sleep(0.3)
                shell.SendKeys("%{F4}")  # Alt+F4 to close
                time.sleep(0.5)
                log.info("Closed Excel window: %s", title)
                return

        log.debug("No Excel window found to close")

    def _wait_for_save_dialog(self, shell: Any) -> None:
        """Wait for the Windows Save dialog to appear."""
        start = time.monotonic()
        dialog_titles = ["Save as", "Save As", "Speichern unter"]

        while time.monotonic() - start < self.dialog_timeout:
            for title in dialog_titles:
                if shell.AppActivate(title):
                    time.sleep(0.3)  # Let dialog fully render
                    log.info("Save dialog detected: '%s'", title)
                    return
            time.sleep(0.2)

        raise ExportError(
            f"Save dialog not found within {self.dialog_timeout}s",
            details={
                "timeout": str(self.dialog_timeout),
                "suggestion": (
                    "The Windows Save dialog may have a different title. "
                    "Check your Windows language settings. "
                    "Try running the export manually to see the dialog title."
                ),
            },
        )

    def _fill_filename(self, shell: Any, file_path: str) -> None:
        """Enter the filename in the Save dialog.

        Uses Alt+N (or Alt+D) to focus the filename field,
        then types the full path.
        """
        # Focus the filename field
        # Alt+N is the standard shortcut for "File name" in English Windows
        # Alt+D is an alternative that works in some versions
        for shortcut in ["%n", "%d"]:
            shell.SendKeys(shortcut)
            time.sleep(self.key_delay)

        # Clear any existing text
        shell.SendKeys("^a")  # Ctrl+A to select all
        time.sleep(self.key_delay)

        # Type the full file path
        # SendKeys has a limit -- use clipboard for long paths
        self._set_clipboard(shell, file_path)
        shell.SendKeys("^v")  # Ctrl+V to paste
        time.sleep(self.key_delay)

        log.debug("Filename entered: %s", file_path)

    def _trigger_save(self, shell: Any) -> None:
        """Trigger the Save button in the dialog."""
        shell.SendKeys("%s")  # Alt+S to Save
        time.sleep(0.5)
        log.debug("Save triggered")

    def _handle_overwrite_confirm(self, shell: Any) -> None:
        """Handle the 'file already exists' overwrite confirmation.

        The confirmation dialog has an "Yes" button.
        Uses Alt+Y to confirm overwrite.
        """
        time.sleep(0.3)
        # Check for overwrite confirmation dialog
        confirm_titles = ["Confirm Save As", "Confirm", "Speichern unter"]
        for title in confirm_titles:
            if shell.AppActivate(title):
                time.sleep(0.2)
                shell.SendKeys("%y")  # Alt+Y for Yes
                time.sleep(0.3)
                log.info("Overwrite confirmed")
                return

    def _wait_for_file(
        self,
        file_path: Path,
        timeout: float = 30.0,
    ) -> None:
        """Wait for the file to appear on disk and have content."""
        start = time.monotonic()
        while time.monotonic() - start < timeout:
            if file_path.exists():
                # Verify file has content
                try:
                    size = file_path.stat().st_size
                    if size > 0:
                        log.debug(
                            "File verified: %s (%d bytes)",
                            file_path,
                            size,
                        )
                        return
                except OSError:
                    pass
            time.sleep(0.3)

        raise ExportError(
            f"Exported file not found or empty after {timeout}s: {file_path}",
            details={
                "path": str(file_path),
                "timeout": str(timeout),
                "suggestion": (
                    "The file may not have been saved. "
                    "Check the target folder permissions and disk space. "
                    "Verify the SAP export completed successfully."
                ),
            },
        )

    def _set_clipboard(self, shell: Any, text: str) -> None:
        """Set clipboard content using WScript.Shell for SendKeys compatibility."""
        # Use PowerShell to set clipboard — avoids shell=True security issue
        try:
            import subprocess
            subprocess.run(
                [  # noqa: S607 — powershell is a system command
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    f"Set-Clipboard -Value '{text}'",
                ],
                capture_output=True,
                timeout=5,
                check=False,
            )
        except Exception:
            # Fallback: use SendKeys with the text directly
            # This has a length limit but works for typical file paths
            shell.SendKeys(text, 1)  # 1 = Wait for completion
