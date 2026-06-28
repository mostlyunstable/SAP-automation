"""Tests for sap_automation.core.base_transaction — Base transaction and result."""

from __future__ import annotations

from dataclasses import fields

from sap_automation.core.base_transaction import TransactionResult


class TestTransactionResult:
    """Test the TransactionResult dataclass."""

    def test_defaults(self) -> None:
        result = TransactionResult(transaction_code="VA23", document_number="123")
        assert result.transaction_code == "VA23"
        assert result.document_number == "123"
        assert result.success is False
        assert result.message == ""
        assert result.exported_files == []
        assert result.data == {}
        assert result.duration_seconds == 0.0
        assert result.timestamp  # Should have a timestamp

    def test_summary_success(self) -> None:
        result = TransactionResult(
            transaction_code="VA23",
            document_number="123",
            success=True,
            duration_seconds=1.5,
        )
        summary = result.summary
        assert "[OK]" in summary
        assert "VA23" in summary
        assert "123" in summary
        assert "1.5s" in summary

    def test_summary_failure(self) -> None:
        result = TransactionResult(
            transaction_code="VA23",
            document_number="123",
            success=False,
            duration_seconds=0.5,
        )
        summary = result.summary
        assert "[FAIL]" in summary

    def test_summary_with_files(self) -> None:
        from pathlib import Path

        result = TransactionResult(
            transaction_code="VA23",
            document_number="123",
            success=True,
            exported_files=[Path("file1.xlsx"), Path("file2.pdf")],
        )
        assert "Files: 2" in result.summary

    def test_all_fields_present(self) -> None:
        field_names = {f.name for f in fields(TransactionResult)}
        expected = {
            "transaction_code",
            "document_number",
            "success",
            "message",
            "error_type",
            "exported_files",
            "data",
            "duration_seconds",
            "timestamp",
            "retry_count",
            "screen_info",
        }
        assert expected == field_names
