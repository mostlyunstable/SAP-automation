"""Abstract base class for SAP transaction handlers and result dataclass."""

from __future__ import annotations

import time
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .config import Config
from .exceptions import (
    SAPConnectionError,
    TransactionError,
    ValidationError,
)
from .logger import get_logger

log = get_logger("base_transaction")


@dataclass
class TransactionResult:
    """Result of a single transaction execution.

    Attributes:
        transaction_code: SAP transaction code (e.g., 'VA23').
        document_number: Document number processed.
        success: Whether the transaction completed successfully.
        message: Human-readable result message.
        error_type: Exception class name if failed.
        exported_files: List of file paths exported.
        data: Additional key-value data extracted from SAP.
        duration_seconds: Wall-clock time for this execution.
        timestamp: ISO timestamp when this result was created.
        retry_count: Number of retries attempted.
        screen_info: Diagnostic screen info at time of failure.
    """

    transaction_code: str
    document_number: str
    success: bool = False
    message: str = ""
    error_type: str = ""
    exported_files: list[Path] = field(default_factory=list)
    data: dict[str, Any] = field(default_factory=dict)
    duration_seconds: float = 0.0
    timestamp: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    retry_count: int = 0
    screen_info: dict[str, str] = field(default_factory=dict)

    @property
    def summary(self) -> str:
        """One-line summary for logging and console output."""
        status = "[OK]" if self.success else "[FAIL]"
        parts = [
            status,
            self.transaction_code,
            self.document_number,
            f"{self.duration_seconds:.1f}s",
        ]
        if self.message:
            parts.append(self.message)
        if self.exported_files:
            parts.append(f"Files: {len(self.exported_files)}")
        if self.retry_count > 0:
            parts.append(f"Retries: {self.retry_count}")
        return " | ".join(parts)

    def to_dict(self) -> dict[str, Any]:
        """Serialize result to dictionary for JSON output."""
        return {
            "transaction_code": self.transaction_code,
            "document_number": self.document_number,
            "success": self.success,
            "message": self.message,
            "error_type": self.error_type,
            "exported_files": [str(f) for f in self.exported_files],
            "data": self.data,
            "duration_seconds": self.duration_seconds,
            "timestamp": self.timestamp,
            "retry_count": self.retry_count,
            "screen_info": self.screen_info,
        }


class BaseTransaction(ABC):
    """Abstract base for all SAP transaction handlers.

    Subclasses must implement:
        - tcode: SAP transaction code property
        - description: Human-readable description property
        - validate_record(record): Validate a single input record
        - execute(session, record): Execute the transaction in SAP

    The template method run() handles the full lifecycle:
        1. Validate input
        2. Connect to SAP
        3. Pre-transaction validation (session health, screen state)
        4. Open transaction
        5. Execute (calls subclass)
        6. Post-transaction validation
        7. Export results
        8. Return TransactionResult
    """

    def __init__(self, config: Config) -> None:
        self.config = config

    @property
    @abstractmethod
    def tcode(self) -> str:
        """SAP transaction code (e.g., 'VA23')."""

    @property
    @abstractmethod
    def description(self) -> str:
        """Human-readable transaction description."""

    @abstractmethod
    def validate_record(self, record: dict[str, Any]) -> None:
        """Validate a single input record.

        Args:
            record: Input data dictionary.

        Raises:
            ValidationError: If the record is invalid.
        """

    @abstractmethod
    def execute(self, session: Any, record: dict[str, Any]) -> dict[str, Any]:
        """Execute the transaction for a single record.

        Args:
            session: Active SAP GUI session object.
            record: Input data dictionary.

        Returns:
            Dictionary of extracted data from SAP.

        Raises:
            TransactionError: If the transaction fails.
        """

    def pre_validate(
        self,
        session: Any,
        record: dict[str, Any],
    ) -> None:
        """Validate SAP state before transaction execution.

        Override in subclasses for transaction-specific pre-checks.

        Default checks:
            - Session is not None
            - SAP is not busy
            - OKCode field is accessible

        Args:
            session: Active SAP GUI session.
            record: Input data dictionary.

        Raises:
            TransactionError: If pre-validation fails.
        """
        if session is None:
            raise TransactionError("No active SAP session")

        try:
            if session.Busy:
                raise TransactionError(
                    "SAP is busy when trying to start transaction"
                )
        except TransactionError:
            raise
        except Exception:  # noqa: S110 — cannot check busy state
            pass  # Cannot check busy state — continue

        # Verify OKCode field is accessible (confirms screen is ready)
        try:
            session.findById("wnd[0]/tbar[0]/okcd")
        except Exception as exc:
            raise TransactionError(
                "Cannot access OKCode field — SAP screen may not be ready",
                details={
                    "suggestion": (
                        "Ensure no other transaction is running. "
                        "Close any open popups and wait for SAP to be idle."
                    ),
                },
            ) from exc

    def post_validate(self, session: Any) -> None:
        """Validate SAP state after transaction execution.

        Override in subclasses for transaction-specific post-checks.

        Default checks:
            - Status bar has no errors
            - Session is still alive

        Args:
            session: Active SAP GUI session.

        Raises:
            TransactionError: If post-validation fails.
        """
        if session is None:
            return

        try:
            status_bar = session.findById("wnd[0]/sbar")
            msg_type = str(status_bar.MessageType)
            msg_text = str(status_bar.Text)
            if msg_type in ("E", "A"):
                raise TransactionError(
                    f"SAP error after transaction: {msg_text}"
                )
        except TransactionError:
            raise
        except Exception:  # noqa: S110 — status bar not accessible
            pass  # Status bar not accessible

    def get_screen_diagnostics(self, session: Any) -> dict[str, str]:
        """Get diagnostic information about the current screen.

        Useful for debugging field ID mismatches and screen layout issues.
        """
        info: dict[str, str] = {}
        if session is None:
            return info

        try:
            info["window_text"] = str(session.findById("wnd[0]").Text)
        except Exception:
            info["window_text"] = "unknown"

        try:
            info["program"] = str(session.findById("wnd[0]/usr").Program)
        except Exception:
            info["program"] = "unknown"

        try:
            status_bar = session.findById("wnd[0]/sbar")
            info["status_type"] = str(status_bar.MessageType)
            info["status_text"] = str(status_bar.Text)
        except Exception:
            info["status_type"] = "unknown"
            info["status_text"] = "unknown"

        return info

    def run(self, record: dict[str, Any]) -> TransactionResult:
        """Execute the full transaction lifecycle for a single record.

        Creates a new SAP connection, processes the record, and disconnects.
        For batch processing, use run_with_session() with a shared connection.
        """
        from .connection import SAPConnection

        system_name = self.config.get("sap.system_name")
        client = self.config.get("sap.client")

        with SAPConnection(
            system_name=system_name,
            client=client,
        ) as sap:
            return self.run_with_session(sap, record)

    def run_with_session(
        self,
        sap: Any,
        record: dict[str, Any],
    ) -> TransactionResult:
        """Execute the transaction using an existing SAP connection.

        Args:
            sap: Active SAPConnection instance.
            record: Input data dictionary.

        Returns:
            TransactionResult with execution details.
        """
        start_time = time.monotonic()
        result = TransactionResult(
            transaction_code=self.tcode,
            document_number=str(record.get("document_number", "")),
        )

        try:
            self.validate_record(record)

            if not sap.validate_session():
                raise SAPConnectionError(
                    "Session validation failed before transaction"
                )

            sap.open_transaction(self.tcode)

            self.pre_validate(sap.session, record)

            extracted_data = self.execute(sap.session, record)
            result.data = extracted_data

            self.post_validate(sap.session)

            if extracted_data:
                files = self._export(record, extracted_data)
                result.exported_files = files
                self._verify_exports(files)

            result.success = True
            result.message = "Completed successfully"

        except ValidationError as exc:
            result.success = False
            result.message = f"Validation error: {exc}"
            result.error_type = "ValidationError"
            log.error(
                "Validation failed for %s: %s",
                result.document_number,
                exc,
            )

        except SAPConnectionError as exc:
            result.success = False
            result.message = f"Connection error: {exc}"
            result.error_type = "SAPConnectionError"
            log.error(
                "Connection error for %s: %s",
                result.document_number,
                exc,
            )

        except TransactionError as exc:
            result.success = False
            result.message = f"Transaction error: {exc}"
            result.error_type = "TransactionError"
            log.error(
                "Transaction failed for %s: %s",
                result.document_number,
                exc,
            )

        except Exception as exc:
            result.success = False
            result.message = str(exc)
            result.error_type = type(exc).__name__
            log.error(
                "Unexpected error for %s: %s",
                result.document_number,
                exc,
                exc_info=True,
            )

        result.duration_seconds = time.monotonic() - start_time
        return result

    def _export(
        self,
        record: dict[str, Any],
        data: dict[str, Any],
    ) -> list[Path]:
        """Export extracted data to file(s)."""
        from ..utils.exporter import ReportExporter

        output_dir = self.config.get("paths.output_dir", "./output")
        export_format = self.config.get("export.format", "xlsx")

        exporter = ReportExporter(output_dir=output_dir)

        filename = f"{self.tcode}_{record.get('document_number', 'unknown')}"

        rows = [data] if isinstance(data, dict) and not isinstance(data, list) else data
        if isinstance(rows, dict):
            rows = [rows]

        return exporter.export(rows, filename=filename, formats=[export_format])

    def _verify_exports(self, files: list[Path]) -> None:
        """Verify exported files exist and are readable."""
        for file_path in files:
            if not file_path.exists():
                raise TransactionError(
                    f"Export file not created: {file_path}"
                )
            if not file_path.is_file():
                raise TransactionError(
                    f"Export path is not a file: {file_path}"
                )
            if file_path.stat().st_size == 0:
                raise TransactionError(
                    f"Export file is empty: {file_path}"
                )
