"""Batch orchestration for SAP transaction processing.

Features:
    - Connection lifecycle: one connection per batch, not per record
    - Configurable retry with exponential backoff for transient SAP errors
    - Dry run mode for validation without SAP modification
    - Environment info collection for diagnostics
    - Detailed failure reporting with suggestions
"""

from __future__ import annotations

import sys
import time
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from importlib import import_module
from typing import Any

from .core.base_transaction import BaseTransaction, TransactionResult
from .core.config import Config
from .core.exceptions import (
    ConfigError,
    SAPBusyError,
    SessionNotFoundError,
)
from .core.logger import get_logger, set_run_id

log = get_logger("orchestrator")

_RETRYABLE_EXCEPTIONS: tuple[type[BaseException], ...] = (
    SAPBusyError,
    SessionNotFoundError,
)


def _collect_environment_info() -> dict[str, str]:
    """Collect environment information for diagnostics."""
    import platform

    info: dict[str, str] = {}
    info["python_version"] = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    info["platform"] = sys.platform
    info["os"] = platform.system()
    info["os_release"] = platform.release()
    info["machine"] = platform.machine()

    if sys.platform == "win32":
        try:
            import win32com.client  # type: ignore[import-untyped]

            gui = win32com.client.GetObject("SAPGUI")
            info["sap_gui_version"] = str(getattr(gui, "Version", "unknown"))
        except Exception:
            info["sap_gui_version"] = "not available"

    return info


@dataclass
class RunSummary:
    """Summary of a batch automation run.

    Attributes:
        run_id: Unique identifier for this run.
        transaction_name: Name of the transaction executed.
        total_records: Total number of records processed.
        successful: Number of successful records.
        failed: Number of failed records.
        retried: Number of records that succeeded after retry.
        skipped: Number of skipped records (e.g., KeyboardInterrupt).
        started_at: ISO timestamp when the run started.
        ended_at: ISO timestamp when the run ended.
        results: List of individual TransactionResult objects.
        environment: Environment information for diagnostics.
        dry_run: Whether this was a dry run.
    """

    run_id: str = field(default_factory=lambda: uuid.uuid4().hex[:12])
    transaction_name: str = ""
    total_records: int = 0
    successful: int = 0
    failed: int = 0
    retried: int = 0
    skipped: int = 0
    started_at: str = field(default_factory=lambda: datetime.now(timezone.utc).isoformat())
    ended_at: str = ""
    results: list[TransactionResult] = field(default_factory=list)
    environment: dict[str, str] = field(default_factory=dict)
    dry_run: bool = False

    @property
    def success_rate(self) -> float:
        if self.total_records == 0:
            return 0.0
        return (self.successful / self.total_records) * 100.0

    @property
    def failed_documents(self) -> list[str]:
        return [
            r.document_number
            for r in self.results
            if not r.success
        ]

    def report(self) -> str:
        """Generate a human-readable run report."""
        lines = [
            "=" * 60,
            "AUTOMATION RUN SUMMARY",
            "=" * 60,
            f"Run ID:        {self.run_id}",
            f"Transaction:   {self.transaction_name}",
            f"Dry Run:       {'Yes' if self.dry_run else 'No'}",
            f"Started:       {self.started_at}",
            f"Ended:         {self.ended_at}",
            "-" * 60,
            f"Total Records: {self.total_records}",
            f"Successful:    {self.successful}",
            f"Failed:        {self.failed}",
            f"Retried:       {self.retried}",
            f"Skipped:       {self.skipped}",
            f"Success Rate:  {self.success_rate:.1f}%",
            "-" * 60,
        ]

        if self.failed_documents:
            lines.append("Failed Documents:")
            for doc in self.failed_documents:
                lines.append(f"  - {doc}")
            lines.append("-" * 60)

        if self.environment:
            lines.append("Environment:")
            for key, val in self.environment.items():
                lines.append(f"  {key}: {val}")
            lines.append("-" * 60)

        lines.append("=" * 60)
        return "\n".join(lines)

    def to_dict(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "transaction_name": self.transaction_name,
            "total_records": self.total_records,
            "successful": self.successful,
            "failed": self.failed,
            "retried": self.retried,
            "skipped": self.skipped,
            "success_rate": self.success_rate,
            "started_at": self.started_at,
            "ended_at": self.ended_at,
            "failed_documents": self.failed_documents,
            "environment": self.environment,
            "dry_run": self.dry_run,
            "results": [r.to_dict() for r in self.results],
        }


class Orchestrator:
    """Batch orchestrator for processing Excel records through SAP transactions.

    Manages connection lifecycle: one connection per batch, shared across all records.
    Applies configurable retry with exponential backoff for transient SAP errors.
    """

    def __init__(self, config: Config) -> None:
        self.config = config

    @staticmethod
    def _load_handler(class_path: str, config: Config) -> BaseTransaction:
        """Dynamically load a transaction handler class."""
        if not class_path.startswith("sap_automation.transactions."):
            raise ConfigError(
                f"Transaction class path must start with 'sap_automation.transactions.', "
                f"got: '{class_path}'"
            )

        try:
            module_path, class_name = class_path.rsplit(".", 1)
            module = import_module(module_path)
            cls = getattr(module, class_name)
        except ImportError as exc:
            raise ConfigError(
                f"Cannot import module '{module_path}': {exc}"
            ) from exc
        except AttributeError as exc:
            raise ConfigError(
                f"Class '{class_name}' not found in module '{module_path}': {exc}"
            ) from exc

        if not (isinstance(cls, type) and issubclass(cls, BaseTransaction)):
            raise ConfigError(
                f"'{class_path}' is not a BaseTransaction subclass"
            )

        return cls(config)

    def _get_retry_config(self) -> dict[str, Any]:
        """Extract retry configuration with sensible defaults."""
        return {
            "max_attempts": self.config.get("retry.max_attempts", 3),
            "delay_seconds": self.config.get("retry.delay_seconds", 2.0),
            "backoff_multiplier": self.config.get("retry.backoff_multiplier", 2.0),
            "max_delay_seconds": self.config.get("retry.max_delay_seconds", 60.0),
        }

    def validate_input(
        self,
        transaction_name: str,
        records: list[dict[str, Any]],
    ) -> list[str]:
        """Validate input without executing any SAP transactions."""
        errors: list[str] = []

        txn_config = self.config.get(f"transactions.{transaction_name}")
        if not txn_config:
            errors.append(f"Transaction '{transaction_name}' not found in config")
            return errors

        class_path = txn_config.get("class")
        if not class_path:
            errors.append(f"Transaction '{transaction_name}' missing 'class' in config")
            return errors

        try:
            handler = self._load_handler(class_path, self.config)
        except ConfigError as exc:
            errors.append(f"Cannot load handler: {exc}")
            return errors

        for i, record in enumerate(records):
            doc_num = record.get("document_number", f"record_{i}")
            try:
                handler.validate_record(record)
            except Exception as exc:
                errors.append(f"Record {i + 1} ({doc_num}): {exc}")

        return errors

    def dry_run(
        self,
        transaction_name: str,
        records: list[dict[str, Any]],
    ) -> RunSummary:
        """Validate configuration, input, and connectivity without modifying SAP."""
        summary = RunSummary(transaction_name=transaction_name, dry_run=True)
        summary.total_records = len(records)
        summary.environment = _collect_environment_info()
        set_run_id(summary.run_id)

        log.info("Starting dry run: %s (%d records)", transaction_name, len(records))

        errors = self.validate_input(transaction_name, records)
        if errors:
            for error in errors:
                log.error("Dry run validation: %s", error)
            summary.failed = len(errors)
            summary.results = [
                TransactionResult(
                    transaction_code=transaction_name,
                    document_number=str(records[i].get("document_number", f"record_{i}")),
                    success=False,
                    message=error,
                    error_type="ValidationError",
                )
                for i, error in enumerate(errors)
            ]
        else:
            summary.successful = len(records)
            summary.results = [
                TransactionResult(
                    transaction_code=transaction_name,
                    document_number=str(r.get("document_number", "")),
                    success=True,
                    message="Input validation passed",
                )
                for r in records
            ]

        if sys.platform == "win32":
            try:
                from .core.connection import SAPConnection

                system_name = self.config.get("sap.system_name")
                client = self.config.get("sap.client")
                with SAPConnection(
                    system_name=system_name,
                    client=client,
                    session_timeout=5.0,
                ) as sap:
                    info = sap.get_session_info()
                    summary.environment["sap_connection"] = "connected"
                    summary.environment["sap_client"] = info.get("client", "unknown")
                    summary.environment["sap_user"] = info.get("user", "unknown")
                    log.info("Dry run: SAP connection successful")
            except Exception as exc:
                summary.environment["sap_connection"] = f"failed: {exc}"
                log.warning("Dry run: SAP connection test failed: %s", exc)

        summary.ended_at = datetime.now(timezone.utc).isoformat()

        log.info(
            "Dry run complete: %d/%d records valid (%.1f%%)",
            summary.successful,
            summary.total_records,
            summary.success_rate,
        )

        return summary

    def run(
        self,
        transaction_name: str,
        records: list[dict[str, Any]],
    ) -> RunSummary:
        """Execute a transaction for all records.

        Connects once to SAP, processes the entire batch, then disconnects.
        Retries transient failures (SAPBusyError, SessionNotFoundError) using
        exponential backoff from config. Never retries validation errors.
        """
        summary = RunSummary(transaction_name=transaction_name)
        summary.total_records = len(records)
        summary.environment = _collect_environment_info()

        txn_config = self.config.get(f"transactions.{transaction_name}")
        if not txn_config:
            raise ConfigError(
                f"Transaction '{transaction_name}' not found in config"
            )

        class_path = txn_config.get("class")
        if not class_path:
            raise ConfigError(
                f"Transaction '{transaction_name}' missing 'class' in config"
            )

        handler = self._load_handler(class_path, self.config)
        retry_cfg = self._get_retry_config()
        set_run_id(summary.run_id)

        log.info(
            "Starting batch run: %s (%d records)",
            transaction_name,
            len(records),
        )

        from .core.connection import SAPConnection

        system_name = self.config.get("sap.system_name")
        client = self.config.get("sap.client")

        i = 0
        try:
            with SAPConnection(
                system_name=system_name,
                client=client,
            ) as sap:
                for i, record in enumerate(records, 1):
                    result = self._execute_with_retry(
                        handler, sap, record, i, len(records), retry_cfg
                    )
                    summary.results.append(result)

                    if result.success:
                        summary.successful += 1
                        if result.retry_count > 0:
                            summary.retried += 1
                    else:
                        summary.failed += 1

        except KeyboardInterrupt:
            log.warning("Interrupted during batch run")
            remaining = records[i:] if i > 0 else records
            summary.skipped = len(remaining)
            for r in remaining:
                summary.results.append(TransactionResult(
                    transaction_code=handler.tcode,
                    document_number=str(r.get("document_number", "")),
                    success=False,
                    message="Skipped (interrupted)",
                ))

        summary.ended_at = datetime.now(timezone.utc).isoformat()

        log.info(
            "Batch run complete: %d/%d successful (%.1f%%)",
            summary.successful,
            summary.total_records,
            summary.success_rate,
        )

        return summary

    def _execute_with_retry(
        self,
        handler: BaseTransaction,
        sap: Any,
        record: dict[str, Any],
        record_num: int,
        total: int,
        retry_cfg: dict[str, Any],
    ) -> TransactionResult:
        """Execute a single record with retry for transient SAP errors.

        Validation errors are never retried — they are deterministic.
        Only SAPBusyError and SessionNotFoundError trigger retry with
        exponential backoff.
        """
        max_attempts = retry_cfg["max_attempts"]
        delay = retry_cfg["delay_seconds"]
        backoff = retry_cfg["backoff_multiplier"]
        max_delay = retry_cfg["max_delay_seconds"]

        last_result: TransactionResult | None = None

        for attempt in range(1, max_attempts + 1):
            log.info(
                "Record %d/%d: %s (attempt %d/%d)",
                record_num,
                total,
                record.get("document_number", "unknown"),
                attempt,
                max_attempts,
            )

            result = handler.run_with_session(sap, record)

            if result.success:
                if attempt > 1:
                    log.info(
                        "Record %d/%d succeeded after %d retries",
                        record_num,
                        total,
                        attempt - 1,
                    )
                result.retry_count = attempt - 1
                return result

            last_result = result

            if result.error_type == "ValidationError":
                return result

            if attempt < max_attempts:
                log.warning(
                    "Record %d/%d failed (%s), retrying in %.1fs",
                    record_num,
                    total,
                    result.message,
                    delay,
                )
                time.sleep(delay)
                delay = min(delay * backoff, max_delay)

        log.error(
            "Record %d/%d failed after %d attempts: %s",
            record_num,
            total,
            max_attempts,
            last_result.message if last_result else "unknown",
        )
        if last_result is None:
            last_result = TransactionResult(
                transaction_code=handler.tcode,
                document_number="unknown",
                success=False,
                message="Unknown error after retries",
            )
        last_result.retry_count = max_attempts - 1
        return last_result
