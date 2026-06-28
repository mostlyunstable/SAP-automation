"""Tests for sap_automation.core.exceptions — Exception hierarchy."""

from sap_automation.core.exceptions import (
    ConfigError,
    ExcelReadError,
    ExportError,
    FieldError,
    SAPAutomationError,
    SAPBusyError,
    SAPCancelledError,
    SAPConnectionError,
    SecurityError,
    SessionNotFoundError,
    TransactionError,
    ValidationError,
)


class TestExceptionHierarchy:
    """Verify exception inheritance and messages."""

    def test_base_exception(self) -> None:
        exc = SAPAutomationError("test msg", details={"key": "val"})
        assert str(exc) == "test msg | Details: {'key': 'val'}"
        assert exc.message == "test msg"
        assert exc.details == {"key": "val"}

    def test_base_no_details(self) -> None:
        exc = SAPAutomationError("simple msg")
        assert str(exc) == "simple msg"

    def test_connection_inherits_base(self) -> None:
        assert issubclass(SAPConnectionError, SAPAutomationError)

    def test_session_not_found_inherits_connection(self) -> None:
        assert issubclass(SessionNotFoundError, SAPConnectionError)

    def test_busy_inherits_base(self) -> None:
        assert issubclass(SAPBusyError, SAPAutomationError)

    def test_cancelled_inherits_base(self) -> None:
        assert issubclass(SAPCancelledError, SAPAutomationError)

    def test_transaction_inherits_base(self) -> None:
        assert issubclass(TransactionError, SAPAutomationError)

    def test_field_inherits_base(self) -> None:
        assert issubclass(FieldError, SAPAutomationError)

    def test_export_inherits_base(self) -> None:
        assert issubclass(ExportError, SAPAutomationError)

    def test_config_inherits_base(self) -> None:
        assert issubclass(ConfigError, SAPAutomationError)

    def test_excel_inherits_base(self) -> None:
        assert issubclass(ExcelReadError, SAPAutomationError)

    def test_validation_inherits_base(self) -> None:
        assert issubclass(ValidationError, SAPAutomationError)

    def test_security_inherits_base(self) -> None:
        assert issubclass(SecurityError, SAPAutomationError)

    def test_catch_all_with_base(self) -> None:
        """All framework exceptions should be catchable with base class."""
        exceptions = [
            SAPConnectionError("e"),
            SessionNotFoundError("e"),
            SAPBusyError("e"),
            TransactionError("e"),
            FieldError("e"),
            ExportError("e"),
            ConfigError("e"),
            ExcelReadError("e"),
            ValidationError("e"),
            SecurityError("e"),
        ]
        for exc in exceptions:
            assert isinstance(exc, SAPAutomationError)
