"""SAP Automation Framework - Core package."""

from .base_transaction import BaseTransaction, TransactionResult
from .config import Config
from .connection import SAPConnection, validate_tcode
from .exceptions import (
    ConfigError,
    ExcelReadError,
    ExportError,
    FieldError,
    NavigationError,
    SAPAutomationError,
    SAPBusyError,
    SAPCancelledError,
    SAPConnectionError,
    SecurityError,
    SessionNotFoundError,
    TransactionError,
    ValidationError,
)
from .logger import get_logger, set_run_id, setup_logging
from .retry import retry_on_exception

__all__ = [
    "BaseTransaction",
    "Config",
    "ConfigError",
    "ExcelReadError",
    "ExportError",
    "FieldError",
    "NavigationError",
    "SAPAutomationError",
    "SAPBusyError",
    "SAPCancelledError",
    "SAPConnection",
    "SAPConnectionError",
    "SecurityError",
    "SessionNotFoundError",
    "TransactionError",
    "TransactionResult",
    "ValidationError",
    "get_logger",
    "retry_on_exception",
    "set_run_id",
    "setup_logging",
    "validate_tcode",
]
