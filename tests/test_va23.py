"""Tests for sap_automation.transactions.va23_display_quotation."""

from __future__ import annotations

import pytest

from sap_automation.core.config import Config
from sap_automation.core.exceptions import ValidationError
from sap_automation.transactions.va23_display_quotation import VA23DisplayQuotation


@pytest.fixture
def va23_handler() -> VA23DisplayQuotation:
    cfg = Config.from_dict({"sap": {"system_name": "PRD", "client": "100"}})
    return VA23DisplayQuotation(cfg)


class TestValidateRecord:
    def test_valid_record(self, va23_handler: VA23DisplayQuotation) -> None:
        va23_handler.validate_record({"document_number": "20001234"})

    def test_missing_document_number(self, va23_handler: VA23DisplayQuotation) -> None:
        with pytest.raises(ValidationError, match="Missing document_number"):
            va23_handler.validate_record({})

    def test_empty_document_number(self, va23_handler: VA23DisplayQuotation) -> None:
        with pytest.raises(ValidationError, match="Missing document_number"):
            va23_handler.validate_record({"document_number": ""})

    def test_none_document_number(self, va23_handler: VA23DisplayQuotation) -> None:
        with pytest.raises(ValidationError, match="Missing document_number"):
            va23_handler.validate_record({"document_number": None})

    def test_non_numeric_document_number(self, va23_handler: VA23DisplayQuotation) -> None:
        with pytest.raises(ValidationError, match="Must be numeric"):
            va23_handler.validate_record({"document_number": "ABC123"})

    def test_too_long_document_number(self, va23_handler: VA23DisplayQuotation) -> None:
        with pytest.raises(ValidationError, match="too long"):
            va23_handler.validate_record({"document_number": "12345678901"})

    def test_whitespace_stripped(self, va23_handler: VA23DisplayQuotation) -> None:
        va23_handler.validate_record({"document_number": "  20001234  "})


class TestVA23Properties:
    def test_tcode(self, va23_handler: VA23DisplayQuotation) -> None:
        assert va23_handler.tcode == "VA23"

    def test_description(self, va23_handler: VA23DisplayQuotation) -> None:
        assert va23_handler.description == "Display Quotation"
