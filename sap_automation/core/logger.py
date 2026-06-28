"""Thread-safe logging setup with secret filtering and execution context."""

import logging
import logging.handlers
import re
import threading
from pathlib import Path
from typing import ClassVar, Optional

# Global state
_initialized = False
_init_lock = threading.Lock()
_execution_filter: Optional["_ExecutionContextFilter"] = None


class _SecretFilter(logging.Filter):
    """Filter that redacts sensitive values from log messages.

    Redacts patterns like:
    - password=xxx, password="xxx", password='xxx'
    - token=xxx, token="xxx", token='xxx'
    - secret=xxx, secret="xxx", secret='xxx'
    - api_key=xxx, api_key="xxx", api_key='xxx'
    - auth=xxx, auth="xxx", auth='xxx'
    - credential=xxx, credential="xxx", credential='xxx'
    """

    _SECRET_PATTERNS: ClassVar[list[str]] = [
        r"(password|token|secret|api_key|auth|credential)\s*=\s*([^\s\"']+)",
        r'(password|token|secret|api_key|auth|credential)\s*=\s*"([^"]*)"',
        r"(password|token|secret|api_key|auth|credential)\s*=\s*'([^']*)'",
    ]

    def filter(self, record: logging.LogRecord) -> bool:
        """Redact secrets from the log message."""
        try:
            message = record.getMessage()
            for pattern in self._SECRET_PATTERNS:
                message = re.sub(
                    pattern,
                    lambda m: f"{m.group(1)}=REDACTED",
                    message,
                    flags=re.IGNORECASE,
                )
            record.msg = message
            record.args = ()
        except Exception:  # noqa: S110 — never let filtering break logging
            pass  # Never let filtering break logging
        return True


class _ExecutionContextFilter(logging.Filter):
    """Filter that injects execution context (run_id) into log records."""

    def __init__(self, run_id: str = "") -> None:
        super().__init__()
        self.run_id = run_id

    def filter(self, record: logging.LogRecord) -> bool:
        record.run_id = self.run_id  # type: ignore[attr-defined]
        return True


def _get_formatter(fmt: str) -> logging.Formatter:
    """Create formatter with run_id support."""
    return logging.Formatter(fmt)


def setup_logging(
    log_dir: str | None = None,
    level: str = "INFO",
    fmt: str | None = None,
    run_id: str = "",
) -> logging.Logger:
    """Initialize logging with rotation, secret filtering, and execution context.

    Args:
        log_dir: Directory for log files. If None, logs to console only.
        level: Log level (DEBUG, INFO, WARNING, ERROR).
        fmt: Custom format string.
        run_id: Execution ID for correlation across modules.

    Returns:
        Root logger for the sap_automation namespace.
    """
    global _initialized, _execution_filter

    with _init_lock:
        if _initialized:
            # Already initialized, just update run_id
            if _execution_filter:
                _execution_filter.run_id = run_id
            return logging.getLogger("sap_automation")

        root = logging.getLogger("sap_automation")
        root.setLevel(getattr(logging, level.upper()))

        for h in root.handlers[:]:
            root.removeHandler(h)

        if fmt is None:
            fmt = "%(asctime)s | %(levelname)-8s | %(name)s | %(run_id)s | %(message)s"

        formatter = _get_formatter(fmt)

        # Add secret filter
        secret_filter = _SecretFilter()

        # Add execution context filter
        _execution_filter = _ExecutionContextFilter(run_id)

        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        console_handler.addFilter(secret_filter)
        console_handler.addFilter(_execution_filter)
        root.addHandler(console_handler)

        # File handler if log_dir specified
        if log_dir:
            log_path = Path(log_dir)
            log_path.mkdir(parents=True, exist_ok=True)
            file_handler = logging.handlers.TimedRotatingFileHandler(
                log_path / "sap_automation.log",
                when="midnight",
                interval=1,
                backupCount=30,
                encoding="utf-8",
            )
            file_handler.setFormatter(formatter)
            file_handler.addFilter(secret_filter)
            file_handler.addFilter(_execution_filter)
            root.addHandler(file_handler)

        _initialized = True
        return root


def get_logger(name: str) -> logging.Logger:
    """Get a child logger under the sap_automation namespace.

    Args:
        name: Logger name (e.g., "txn.VA23", "connection").

    Returns:
        Configured logger instance.
    """
    if not _initialized:
        setup_logging()
    return logging.getLogger(f"sap_automation.{name}")


def set_run_id(run_id: str) -> None:
    """Update the execution context run_id for all loggers.

    Args:
        run_id: Execution identifier for correlation.
    """
    global _execution_filter
    if _execution_filter is None:
        setup_logging(run_id=run_id)
    else:
        _execution_filter.run_id = run_id
