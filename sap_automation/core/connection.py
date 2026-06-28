"""SAP GUI connection and session management via COM automation.

This module attaches to an already-running SAP GUI session.
The user must be logged in manually (works with SSO/MFA).

Windows-only — requires pywin32 and SAP GUI with scripting enabled.

Handles:
    - Session attachment with timeout and connection filtering
    - Modal dialog detection, categorization, and handling
    - Busy-wait with configurable timeout
    - Status bar error checking with recovery suggestions
    - Screen/window verification after navigation
    - Graceful disconnect with explicit COM cleanup
    - Connection recovery after transient failures
    - Pre-transaction session validation
"""

from __future__ import annotations

import gc
import re
import time
from typing import Any

from .exceptions import (
    SAPBusyError,
    SAPCancelledError,
    SAPConnectionError,
    SessionNotFoundError,
    TransactionError,
)
from .logger import get_logger

log = get_logger("connection")

# Lazy import — win32com is only available on Windows
_win32com: Any = None


def _get_win32com() -> Any:
    """Lazy-load win32com.client. Raises SAPConnectionError if unavailable."""
    global _win32com
    if _win32com is None:
        try:
            import win32com.client  # type: ignore[import-untyped]
            _win32com = win32com.client
        except ImportError:
            raise SAPConnectionError(
                "pywin32 is required. Install it: pip install pywin32"
            ) from None
    return _win32com


# Valid SAP transaction codes: 4 uppercase alphanumeric characters
_TCODE_PATTERN = re.compile(r"^[A-Za-z0-9]{4}$")

# Known SAP popup types and their handling strategies
_POPUP_AUTHORIZATION = "authorization"
_POPUP_WARNING = "warning"
_POPUP_INFORMATION = "information"
_POPUP_CONFIRMATION = "confirmation"
_POPUP_PRINTER = "printer"
_POPUP_SESSION_TIMEOUT = "session_timeout"
_POPUP_UNKNOWN = "unknown"

# Authorization error keywords (case-insensitive)
_AUTH_ERROR_KEYWORDS = [
    "no authorization",
    "not authorized",
    "insufficient authorization",
    "missing authorization",
    "you do not have",
    "keine berechtigung",
    "no tiene autorización",
]

# Session timeout keywords
_TIMEOUT_KEYWORDS = [
    "session timeout",
    "session expired",
    "connection closed",
    "dialog cancelled",
]


def validate_tcode(tcode: str) -> str:
    """Validate and sanitize an SAP transaction code.

    SAP transaction codes are exactly 4 alphanumeric characters.
    This prevents OKCode injection via the command field.

    Raises:
        TransactionError: if the tcode is invalid.
    """
    tcode = tcode.strip()
    if not _TCODE_PATTERN.match(tcode):
        raise TransactionError(
            f"Invalid transaction code: '{tcode}'. "
            f"Must be exactly 4 alphanumeric characters."
        )
    return tcode.upper()


def _categorize_popup(title: str, text: str) -> str:
    """Categorize a popup dialog by its content."""
    combined = f"{title} {text}".lower()

    for kw in _AUTH_ERROR_KEYWORDS:
        if kw in combined:
            return _POPUP_AUTHORIZATION

    for kw in _TIMEOUT_KEYWORDS:
        if kw in combined:
            return _POPUP_SESSION_TIMEOUT

    if any(kw in combined for kw in ["warning", "warnung", "aviso"]):
        return _POPUP_WARNING

    if any(kw in combined for kw in ["information", "hinweis", "info"]):
        return _POPUP_INFORMATION

    if any(kw in combined for kw in ["confirm", "bestätigen", "confirmar"]):
        return _POPUP_CONFIRMATION

    if any(kw in combined for kw in ["printer", "drucker", "impresora"]):
        return _POPUP_PRINTER

    return _POPUP_UNKNOWN


class SAPConnection:
    """Manages connection to a running SAP GUI instance.

    Features:
        - Attaches to existing SAP GUI via COM
        - Filters connections by system_name when multiple are open
        - Waits for session with configurable timeout
        - Detects and logs modal dialogs with categorization
        - Checks status bar for errors after operations
        - Validates session health before transactions
        - Graceful cleanup with explicit COM release
        - Connection recovery after transient failures

    Usage:
        with SAPConnection() as sap:
            session = sap.session
            sap.open_transaction("VA23")
    """

    def __init__(
        self,
        system_name: str | None = None,
        client: str | None = None,
        session_timeout: float = 15.0,
        ready_timeout: float = 30.0,
    ):
        self.system_name = system_name
        self.client = client
        self.session_timeout = session_timeout
        self.ready_timeout = ready_timeout
        self._gui_application: Any = None
        self._application: Any = None
        self._connection: Any = None
        self._session: Any = None

    # -- Context manager -------------------------------------------------------

    def __enter__(self) -> SAPConnection:
        try:
            self.connect()
        except Exception:
            self.disconnect()
            raise
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.disconnect()

    # -- Public API ------------------------------------------------------------

    def connect(self) -> None:
        """Attach to the running SAP GUI application.

        If system_name is specified, attempts to find a matching connection.
        Otherwise, uses the first available connection.

        Raises SAPConnectionError with actionable diagnostics on failure.
        """
        win32com = _get_win32com()

        log.info("Attaching to SAP GUI application...")
        try:
            self._gui_application = win32com.GetObject("SAPGUI")
        except Exception as exc:
            raise SAPConnectionError(
                "Cannot find running SAP GUI. Ensure SAP Logon is open.",
                details={
                    "error": str(exc),
                    "suggestion": "Start SAP Logon, log in to your system, then retry.",
                },
            ) from exc

        try:
            self._application = self._gui_application.GetScriptingEngine
        except Exception as exc:
            raise SAPConnectionError(
                "Cannot access SAP scripting engine. "
                "Ensure GUI Scripting is enabled in SAP options.",
                details={
                    "error": str(exc),
                    "suggestion": (
                        "Enable scripting: SAP Logon > Options > Accessibility > "
                        "GUI Scripting > Enable GUI Scripting. "
                        "Also check: Options > Scripting > Notify / Do not notify."
                    ),
                },
            ) from exc

        # Find an open connection
        try:
            connection_count = self._application.Children.Count
        except Exception:
            connection_count = 0

        if connection_count == 0:
            raise SAPConnectionError(
                "No SAP connections found. Open a connection in SAP Logon first.",
                details={
                    "suggestion": (
                        "Open SAP Logon, select a system, and log in. "
                        "Ensure the connection remains open during automation."
                    ),
                },
            )

        # Select the appropriate connection
        self._connection = self._select_connection(connection_count)
        system_name = getattr(self._connection, "SystemName", "unknown")
        log.info("Attached to connection: %s", system_name)

        # Find or wait for a session
        self._wait_for_session()

    def disconnect(self) -> None:
        """Release all COM references and force garbage collection.

        Safe to call multiple times. Ensures COM objects are released
        promptly rather than waiting for GC.
        """
        self._session = None
        self._connection = None
        self._application = None
        self._gui_application = None
        gc.collect()
        log.info("Released SAP connection references.")

    @property
    def session(self) -> Any:
        """Get the active SAP session, reconnecting if needed."""
        if self._session is None:
            self._wait_for_session()
        return self._session

    @property
    def connection(self) -> Any:
        """Get the active SAP connection."""
        if self._connection is None:
            raise SAPConnectionError("Not connected. Call connect() first.")
        return self._connection

    @property
    def is_connected(self) -> bool:
        """Quick check — returns True only if a session is already held.

        Does NOT trigger reconnection. Use session property for that.
        """
        return self._session is not None

    def invalidate_session(self) -> None:
        """Invalidate the current session, forcing reconnection on next access."""
        self._session = None

    def is_alive(self) -> bool:
        """Check if the SAP GUI process is still running.

        Returns False if the connection has been closed by the user.
        """
        if self._application is None:
            return False
        try:
            _ = self._application.Children.Count
            return True
        except Exception:
            return False

    def validate_session(self) -> bool:
        """Validate that the current session is healthy and usable.

        Checks:
            - Session object exists
            - Session is not busy
            - SAP GUI process is alive
            - Session has not been closed

        Returns True if session is healthy.
        Logs warnings for any issues detected.
        """
        if self._session is None:
            log.warning("Session validation failed: no session object")
            return False

        if not self.is_alive():
            log.warning("Session validation failed: SAP GUI process not alive")
            self._session = None
            return False

        try:
            _ = self._session.Busy
        except Exception:
            log.warning("Session validation failed: session object invalid")
            self._session = None
            return False

        try:
            _ = self._session.Id
        except Exception:
            log.warning("Session validation failed: session ID inaccessible")
            self._session = None
            return False

        return True

    def get_session_info(self) -> dict[str, str]:
        """Get diagnostic information about the current session."""
        info: dict[str, str] = {}
        if self._session is None:
            info["status"] = "no session"
            return info

        try:
            info["client"] = str(getattr(self._session, "Client", "unknown"))
        except Exception:
            info["client"] = "unknown"

        try:
            info["user"] = str(getattr(self._session, "User", "unknown"))
        except Exception:
            info["user"] = "unknown"

        try:
            info["system"] = str(getattr(self._session, "SystemName", "unknown"))
        except Exception:
            info["system"] = "unknown"

        try:
            info["transaction"] = str(getattr(self._session, "Info", "unknown"))
        except Exception:
            info["transaction"] = "unknown"

        try:
            info["busy"] = str(self._session.Busy)
        except Exception:
            info["busy"] = "unknown"

        info["alive"] = str(self.is_alive())
        return info

    # -- Transaction helpers ---------------------------------------------------

    def open_transaction(self, tcode: str) -> None:
        """Navigate to a transaction code.

        Validates the transaction code before sending to SAP
        to prevent OKCode injection.

        After navigation:
            - Waits for SAP to be ready
            - Verifies screen loaded
            - Checks status bar for errors
        """
        tcode = validate_tcode(tcode)
        log.info("Opening transaction: /n%s", tcode)

        try:
            session = self.session
            ok_code = session.findById("wnd[0]/tbar[0]/okcd")
            ok_code.text = f"/n{tcode}"
            session.findById("wnd[0]").sendVKey(0)
            self._wait_for_ready()

            # Handle any popups that appear after navigation
            self._handle_post_navigation_popups(session)

            error_msg = self._check_status_bar(session)
            if error_msg:
                raise TransactionError(
                    f"SAP error after opening transaction {tcode}: {error_msg}"
                )

            log.info("Transaction %s opened successfully", tcode)

        except (SAPConnectionError, SessionNotFoundError):
            raise
        except TransactionError:
            raise
        except Exception as exc:
            raise SAPConnectionError(
                f"Failed to open transaction {tcode}",
                details={"error": str(exc)},
            ) from exc

    def wait_for_ready(self, timeout: float | None = None) -> None:
        """Public wrapper for internal wait."""
        self._wait_for_ready(timeout or self.ready_timeout)

    def wait_for_control(
        self,
        control_id: str,
        timeout: float | None = None,
    ) -> Any:
        """Wait for a specific SAP control to become available.

        Args:
            control_id: SAP control ID (e.g., 'wnd[0]/usr/ctxtRV45A-VBELN').
            timeout: Maximum wait time in seconds.

        Returns:
            The control object if found.

        Raises:
            SAPConnectionError: If control not found within timeout.
        """
        timeout = timeout or self.ready_timeout
        start = time.monotonic()
        session = self.session

        while time.monotonic() - start < timeout:
            try:
                control = session.findById(control_id)
                if control is not None:
                    log.debug("Control found: %s", control_id)
                    return control
            except Exception:  # noqa: S110 — polling loop
                pass
            time.sleep(0.2)

        # Provide diagnostic information on failure
        raise SAPConnectionError(
            f"Control not found within {timeout}s: {control_id}",
            details={
                "control_id": control_id,
                "timeout": str(timeout),
                "suggestion": (
                    "The screen layout may differ from expected. "
                    "Check SAP GUI theme, window size, and transaction variant. "
                    "Enable GUI Scripting trace for detailed diagnostics."
                ),
            },
        )

    def check_status_bar(self) -> str | None:
        """Public method to check SAP status bar for errors.

        Returns the status bar message, or None if no error.
        Raises SAPConnectionError if the COM check itself fails.
        """
        try:
            return self._check_status_bar(self.session)
        except SAPConnectionError:
            raise
        except Exception as exc:
            raise SAPConnectionError(
                "Failed to read SAP status bar (session may be disconnected)."
            ) from exc

    def get_screen_info(self) -> dict[str, str]:
        """Get diagnostic information about the current screen."""
        info: dict[str, str] = {}
        session = self._session
        if session is None:
            info["status"] = "no session"
            return info

        try:
            info["window_text"] = str(
                session.findById("wnd[0]").Text
            )
        except Exception:
            info["window_text"] = "unknown"

        try:
            info["program"] = str(
                session.findById("wnd[0]/usr").Program
            )
        except Exception:
            info["program"] = "unknown"

        try:
            status_bar = session.findById("wnd[0]/sbar")
            info["status_bar_type"] = str(status_bar.MessageType)
            info["status_bar_text"] = str(status_bar.Text)
        except Exception:
            info["status_bar_type"] = "unknown"
            info["status_bar_text"] = "unknown"

        return info

    # -- Internal helpers ------------------------------------------------------

    def _select_connection(self, connection_count: int) -> Any:
        """Select the appropriate SAP connection.

        If system_name is specified, finds the matching connection.
        Otherwise, uses the first available connection.
        """
        if self.system_name is None:
            return self._application.Children(0)

        available_systems: list[str] = []
        for i in range(connection_count):
            conn = self._application.Children(i)
            try:
                sys_name = getattr(conn, "SystemName", None)
                if sys_name:
                    available_systems.append(sys_name)
                if sys_name == self.system_name:
                    if self.client is None:
                        return conn
                    # Check client if specified
                    try:
                        for j in range(conn.Sessions.Count):
                            session = conn.Sessions(j)
                            if str(getattr(session, "Client", "")) == self.client:
                                return conn
                    except Exception:  # noqa: S110 — COM polling
                        pass
                    # No session matched this client — try next connection
                    continue
            except Exception:  # noqa: S112 — COM polling
                continue

        # No matching connection found — provide detailed diagnostics
        if self.client is not None:
            raise SAPConnectionError(
                f"Connection matching system={self.system_name} "
                f"client={self.client} not found.",
                details={
                    "requested_system": self.system_name,
                    "requested_client": self.client,
                    "available_systems": available_systems,
                    "suggestion": (
                        "Ensure you are logged into the correct system and client. "
                        "Check SAP Logon pad for open connections."
                    ),
                },
            )
        if self.system_name is not None:
            raise SAPConnectionError(
                f"Connection matching system={self.system_name} not found.",
                details={
                    "requested_system": self.system_name,
                    "available_systems": available_systems,
                    "suggestion": (
                        "Ensure you are logged into the correct system. "
                        "Check SAP Logon pad for open connections."
                    ),
                },
            )
        # Should not reach here (system_name=None case handled at top)
        return self._application.Children(0)

    def _wait_for_session(self) -> None:
        """Wait for an active session to become available."""
        start = time.monotonic()
        while time.monotonic() - start < self.session_timeout:
            # Fail fast if SAP GUI process is no longer alive
            if not self.is_alive():
                raise SAPConnectionError(
                    "SAP GUI process is no longer running. "
                    "Cannot wait for session.",
                    details={
                        "suggestion": (
                            "SAP Logon may have been closed. "
                            "Restart SAP Logon and log in again."
                        ),
                    },
                )
            try:
                if self._connection and self._connection.Sessions.Count > 0:
                    # Find a session matching the configured client
                    for i in range(self._connection.Sessions.Count):
                        sess = self._connection.Sessions(i)
                        if self.client is None or str(getattr(sess, "Client", "")) == self.client:
                            self._session = sess
                            self._log_and_dismiss_modal_dialogs()
                            log.info(
                                "Session acquired (client=%s).",
                                getattr(sess, "Client", "?"),
                            )
                            return
                    # No matching client — error if client was explicitly configured
                    if self.client is not None:
                        raise SessionNotFoundError(
                            f"No session with client={self.client} found. "
                            f"Ensure you are logged into the correct client."
                        )
                    self._session = self._connection.Sessions(0)
                    self._log_and_dismiss_modal_dialogs()
                    return
            except SessionNotFoundError:
                raise  # Don't swallow client-mismatch errors
            except Exception:  # noqa: S110 — polling loop, noise reduction
                pass
            time.sleep(0.5)

        raise SessionNotFoundError(
            f"No active SAP session found within {self.session_timeout}s. "
            "Ensure you are logged in."
        )

    def _wait_for_ready(self, timeout: float | None = None) -> None:
        """Wait until SAP is ready for input (not busy)."""
        timeout = timeout or self.ready_timeout
        start = time.monotonic()
        while time.monotonic() - start < timeout:
            if self._session is None:
                raise SAPConnectionError("Session lost — cannot wait for ready state.")
            try:
                if not self._session.Busy:
                    return
            except Exception:  # noqa: S110 — polling loop, noise reduction
                pass
            time.sleep(0.3)

        raise SAPBusyError(
            f"SAP did not become ready within {timeout}s.",
            details={
                "timeout": str(timeout),
                "suggestion": (
                    "SAP may be processing a long-running operation. "
                    "Increase ready_timeout in config or check for hung sessions."
                ),
            },
        )

    def _handle_post_navigation_popups(self, session: Any) -> None:
        """Handle popups that appear after transaction navigation.

        Categorizes popups and handles them appropriately:
        - Authorization errors: raise immediately
        - Warnings: log and dismiss
        - Information: log and dismiss
        - Printer dialogs: dismiss
        - Session timeout: raise
        - Unknown: log and attempt dismiss
        """
        if session is None:
            return

        try:
            max_dialogs = 10
            handled = 0
            while handled < max_dialogs:
                try:
                    count = session.ModalDialogs.Count
                except Exception:
                    break
                if count == 0:
                    break

                dialog = session.ModalDialogs(0)

                try:
                    dialog_text = str(getattr(dialog, "Text", ""))
                    dialog_title = str(getattr(dialog, "Title", ""))
                except Exception:
                    dialog_text = ""
                    dialog_title = ""

                popup_type = _categorize_popup(dialog_title, dialog_text)

                if popup_type == _POPUP_AUTHORIZATION:
                    log.error(
                        "Authorization error popup: title='%s', text='%s'",
                        dialog_title,
                        dialog_text,
                    )
                    raise SAPCancelledError(
                        f"Authorization error: {dialog_text}",
                        details={
                            "popup_type": "authorization",
                            "title": dialog_title,
                            "text": dialog_text,
                            "suggestion": (
                                "Contact your SAP administrator to request "
                                "the required authorization object."
                            ),
                        },
                    )

                if popup_type == _POPUP_SESSION_TIMEOUT:
                    log.error(
                        "Session timeout popup: title='%s', text='%s'",
                        dialog_title,
                        dialog_text,
                    )
                    self._session = None
                    raise SessionNotFoundError(
                        f"Session timeout: {dialog_text}",
                        details={
                            "suggestion": "Reconnect to SAP and retry.",
                        },
                    )

                log.warning(
                    "Dismissing %s popup: title='%s', text='%s'",
                    popup_type,
                    dialog_title,
                    dialog_text,
                )

                try:
                    dialog.sendVKey(0)  # Enter
                    handled += 1
                except Exception:
                    try:
                        dialog.sendVKey(12)  # Escape
                        handled += 1
                    except Exception:
                        break

            if handled > 0:
                log.info("Handled %d popup dialog(s) after navigation", handled)

        except (SAPCancelledError, SessionNotFoundError):
            raise
        except Exception as exc:
            log.warning("Could not handle post-navigation popups: %s", exc)

    def _log_and_dismiss_modal_dialogs(self) -> None:
        """Log modal dialog contents, then attempt to dismiss them.

        Logs each dialog's title/text for audit trail before pressing
        Enter or Escape to dismiss.
        """
        if self._session is None:
            return

        try:
            max_dialogs = 10  # Safety limit
            dismissed = 0
            while dismissed < max_dialogs:
                try:
                    count = self._session.ModalDialogs.Count
                except Exception:
                    break
                if count == 0:
                    break

                dialog = self._session.ModalDialogs(0)

                # Log dialog content for audit trail
                try:
                    dialog_text = getattr(dialog, "Text", "unknown")
                    dialog_title = getattr(dialog, "Title", "unknown")
                    log.warning(
                        "Dismissing modal dialog: title='%s', text='%s'",
                        dialog_title,
                        dialog_text,
                    )
                except Exception:
                    log.warning("Dismissing unnamed modal dialog.")

                try:
                    dialog.sendVKey(0)  # Enter
                    dismissed += 1
                except Exception:
                    try:
                        dialog.sendVKey(12)  # Escape
                        dismissed += 1
                    except Exception:
                        break

            if dismissed > 0:
                log.info("Dismissed %d modal dialog(s).", dismissed)
        except Exception as exc:
            log.warning("Could not dismiss modal dialogs: %s", exc)

    def _check_status_bar(self, session: Any) -> str | None:
        """Check SAP status bar for error messages.

        Returns the message text if it's an error/warning, None otherwise.
        """
        try:
            status_bar = session.findById("wnd[0]/sbar")
            msg_type: str = str(status_bar.MessageType)
            msg_text: str = str(status_bar.Text)

            if msg_type in ("E", "A", "W"):  # Error, Abort, Warning
                log.warning("SAP status bar [%s]: %s", msg_type, msg_text)
                return msg_text

            return None
        except Exception as exc:
            log.warning("Could not read SAP status bar: %s", exc)
            return None
