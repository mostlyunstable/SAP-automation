"""Tests for sap_automation.core.connection — SAP connection and tcode validation."""

from __future__ import annotations

import pytest

from sap_automation.core.connection import SAPConnection, validate_tcode
from sap_automation.core.exceptions import SAPConnectionError, TransactionError


class TestValidateTcode:
    """Test transaction code validation (OKCode injection prevention)."""

    def test_valid_tcode(self) -> None:
        assert validate_tcode("VA23") == "VA23"

    def test_lowercase_converted(self) -> None:
        assert validate_tcode("va23") == "VA23"

    def test_with_whitespace(self) -> None:
        assert validate_tcode("  VA23  ") == "VA23"

    def test_too_short(self) -> None:
        with pytest.raises(TransactionError, match="Invalid transaction code"):
            validate_tcode("VA2")

    def test_too_long(self) -> None:
        with pytest.raises(TransactionError, match="Invalid transaction code"):
            validate_tcode("VA234")

    def test_special_characters(self) -> None:
        with pytest.raises(TransactionError, match="Invalid transaction code"):
            validate_tcode("VA;2")

    def test_injection_attempt(self) -> None:
        with pytest.raises(TransactionError, match="Invalid transaction code"):
            validate_tcode("/nSE38")

    def test_sql_injection(self) -> None:
        with pytest.raises(TransactionError, match="Invalid transaction code"):
            validate_tcode("';--")

    def test_alphanumeric(self) -> None:
        assert validate_tcode("AB12") == "AB12"


class TestSAPConnection:
    """Test connection lifecycle (without actual SAP)."""

    def test_init_defaults(self) -> None:
        conn = SAPConnection()
        assert conn.system_name is None
        assert conn.client is None
        assert conn.is_connected is False

    def test_init_custom(self) -> None:
        conn = SAPConnection(system_name="QAS", client="200", session_timeout=5.0)
        assert conn.system_name == "QAS"
        assert conn.client == "200"
        assert conn.session_timeout == 5.0

    def test_disconnect_safe(self) -> None:
        conn = SAPConnection()
        # Should not raise
        conn.disconnect()

    def test_connection_property_when_disconnected(self) -> None:
        conn = SAPConnection()
        with pytest.raises(SAPConnectionError):
            _ = conn.connection

    def test_context_manager(self) -> None:
        conn = SAPConnection()
        # Without SAP, connect() will fail, but disconnect should be safe
        try:
            with conn:
                pass
        except Exception:  # noqa: S110
            pass
        # After context exit, should be disconnected
        assert conn._session is None
