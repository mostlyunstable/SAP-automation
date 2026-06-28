"""Type-safe exception hierarchy for SAP Automation Framework."""

from typing import Any


class SAPAutomationError(Exception):
    """Base exception for all SAP automation errors.

    Attributes:
        message: The error message.
        details: Optional dictionary with additional error context.
    """

    def __init__(self, message: str, details: dict[str, Any] | None = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}

    def __str__(self) -> str:
        if self.details:
            return f"{self.message} | Details: {self.details}"
        return self.message


class SAPBusyError(SAPAutomationError):
    """Raised when SAP is busy and retry is needed."""


class SAPConnectionError(SAPAutomationError):
    """Raised for SAP GUI connection issues."""


class SessionNotFoundError(SAPConnectionError):
    """Raised when SAP session cannot be found."""


class SAPCancelledError(SAPAutomationError):
    """Raised when user cancels SAP operation."""


class TransactionError(SAPAutomationError):
    """Raised for general transaction failures."""


class NavigationError(TransactionError):
    """Raised when navigation to a transaction fails."""


class FieldError(TransactionError):
    """Raised when SAP field interaction fails."""


class ExportError(TransactionError):
    """Raised when report export fails."""


class ConfigError(SAPAutomationError):
    """Raised for configuration errors."""


class ExcelReadError(SAPAutomationError):
    """Raised for Excel file reading errors."""


class ValidationError(TransactionError):
    """Raised when input validation fails."""


class SecurityError(SAPAutomationError):
    """Raised for security policy violations."""
