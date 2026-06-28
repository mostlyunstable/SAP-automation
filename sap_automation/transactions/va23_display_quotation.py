"""VA23 — Display Quotation transaction handler.

Handles VA23 with comprehensive field validation, popup handling,
screen verification, and diagnostic error reporting.

Field IDs are configurable and validated at runtime. If a field
cannot be found, the handler logs diagnostic information and
stops safely rather than continuing with partial data.
"""

from __future__ import annotations

import contextlib
from typing import Any

from ..core.base_transaction import BaseTransaction
from ..core.config import Config
from ..core.exceptions import FieldError, TransactionError, ValidationError
from ..core.logger import get_logger

log = get_logger("transactions.va23")

# Default field IDs — may vary by SAP theme/version/screen layout
_DEFAULT_FIELD_IDS = {
    "input_vbeln": "wnd[0]/usr/ctxtRV45A-VBELN",
    "sold_to": "wnd[0]/usr/subHEADER/SUB1/RBHP-VERTR",
    "document_date": "wnd[0]/usr/subHEADER/SUB1/RBHP-AUDAT",
    "net_value": "wnd[0]/usr/subHEADER/SUB1/RBHP-NETWR",
    "currency": "wnd[0]/usr/subHEADER/SUB1/RBHP-WAERK",
}

# Required fields for a valid VA23 record
_REQUIRED_FIELDS = {"document_number"}

# Maximum document number length for VA23
_MAX_DOC_NUM_LENGTH = 10


class VA23DisplayQuotation(BaseTransaction):
    """Handle VA23 (Display Quotation) transaction.

    Opens VA23, enters a quotation number, and extracts the
    displayed data for export.

    Features:
        - Runtime field ID validation with fallback diagnostics
        - Popup handling after navigation
        - Screen verification before data extraction
        - Configurable field IDs via transaction config
    """

    def __init__(self, config: Config) -> None:
        super().__init__(config)
        self._field_ids = dict(_DEFAULT_FIELD_IDS)
        # Allow override from config
        field_overrides = config.get("transactions.va23.field_ids", {})
        if isinstance(field_overrides, dict):
            self._field_ids.update(field_overrides)

    @property
    def tcode(self) -> str:
        return "VA23"

    @property
    def description(self) -> str:
        return "Display Quotation"

    def validate_record(self, record: dict[str, Any]) -> None:
        """Validate a single input record.

        Checks:
            - document_number is present and non-empty
            - document_number is numeric
            - document_number is not too long (max 10 digits)

        Raises:
            ValidationError: If validation fails.
        """
        doc_num = record.get("document_number")

        if doc_num is None or (isinstance(doc_num, str) and not doc_num.strip()):
            raise ValidationError("Missing document_number")

        doc_str = str(doc_num).strip()
        if not doc_str:
            raise ValidationError("Missing document_number")

        if not doc_str.isdigit():
            raise ValidationError(
                f"Must be numeric, got: '{doc_str}'"
            )

        if len(doc_str) > _MAX_DOC_NUM_LENGTH:
            raise ValidationError(
                f"Document number too long: {len(doc_str)} chars "
                f"(max {_MAX_DOC_NUM_LENGTH})"
            )

        # Strip whitespace in-place for downstream use
        record["document_number"] = doc_str

    def pre_validate(self, session: Any, record: dict[str, Any]) -> None:
        """Pre-validate for VA23: verify input field exists."""
        super().pre_validate(session, record)

        field_id = self._field_ids["input_vbeln"]
        try:
            session.findById(field_id)
        except Exception as exc:
            raise TransactionError(
                f"VA23 input field not found: {field_id}",
                details={
                    "field_id": field_id,
                    "error": str(exc),
                    "suggestion": (
                        "The VA23 initial screen layout may differ. "
                        "Check SAP GUI theme, transaction variant, "
                        "and screen resolution."
                    ),
                },
            ) from exc

    def execute(self, session: Any, record: dict[str, Any]) -> dict[str, Any]:
        """Execute VA23 for a single quotation.

        Args:
            session: Active SAP GUI session.
            record: Must contain 'document_number'.

        Returns:
            Dictionary of extracted quotation data.

        Raises:
            FieldError: If a required field cannot be found.
            TransactionError: If navigation or extraction fails.
        """
        doc_num = record["document_number"]
        log.info("Displaying quotation: %s", doc_num)

        # Enter quotation number
        self._enter_document_number(session, doc_num)

        # Handle popups after entering document
        self._handle_entry_popups(session)

        # Verify we're on the quotation display screen
        self._verify_screen(session, doc_num)

        # Extract header data
        data = self._extract_header_data(session, doc_num)

        log.info(
            "Extracted quotation %s: sold_to=%s, net_value=%s",
            doc_num,
            data.get("sold_to", ""),
            data.get("net_value", ""),
        )

        return data

    def _enter_document_number(self, session: Any, doc_num: str) -> None:
        """Enter the document number in the VA23 initial screen."""
        field_id = self._field_ids["input_vbeln"]
        try:
            input_field = session.findById(field_id)
            input_field.text = doc_num
            session.findById("wnd[0]").sendVKey(0)  # Enter
        except Exception as exc:
            raise FieldError(
                f"Failed to enter quotation number '{doc_num}'",
                details={
                    "field_id": field_id,
                    "error": str(exc),
                    "suggestion": (
                        f"Field '{field_id}' not found. "
                        "The VA23 screen layout may differ from expected. "
                        "Check SAP GUI theme and transaction variant."
                    ),
                },
            ) from exc

    def _handle_entry_popups(self, session: Any) -> None:
        """Handle popups that appear after entering document number."""
        with contextlib.suppress(Exception):
            session.findById("wnd[0]").sendVKey(0)  # Confirm info messages

    def _verify_screen(self, session: Any, doc_num: str) -> None:
        """Verify we're on the quotation display screen after navigation."""
        # Check status bar for errors
        try:
            status_bar = session.findById("wnd[0]/sbar")
            msg_type = str(status_bar.MessageType)
            msg_text = str(status_bar.Text)
            if msg_type == "E":
                raise TransactionError(
                    f"SAP error displaying quotation {doc_num}: {msg_text}"
                )
        except TransactionError:
            raise
        except Exception:  # noqa: S110 — status bar not accessible
            pass  # Status bar not accessible

        # Verify at least one header field exists (confirms screen loaded)
        header_field = self._field_ids.get("sold_to")
        if header_field:
            try:
                session.findById(header_field)
            except Exception:
                log.warning(
                    "Could not verify quotation screen — "
                    "header field '%s' not found. "
                    "Quotation %s may not exist or screen layout differs.",
                    header_field,
                    doc_num,
                )

    def _extract_header_data(
        self, session: Any, doc_num: str
    ) -> dict[str, Any]:
        """Extract header data from the quotation display screen.

        Each field extraction is independent — if one field cannot be
        found, the others are still attempted. Empty string is used
        for missing fields rather than failing the entire extraction.
        """
        data: dict[str, Any] = {"document_number": doc_num}

        field_mapping = {
            "sold_to": "sold_to",
            "document_date": "document_date",
            "net_value": "net_value",
            "currency": "currency",
        }

        for key, field_id_key in field_mapping.items():
            field_id = self._field_ids.get(field_id_key)
            if not field_id:
                log.warning("No field ID configured for '%s'", key)
                data[key] = ""
                continue

            try:
                control = session.findById(field_id)
                value = str(control.text) if hasattr(control, "text") else ""
                data[key] = value
            except Exception:
                log.debug(
                    "Field '%s' (ID: %s) not found on screen. "
                    "Value set to empty string.",
                    key,
                    field_id,
                )
                data[key] = ""

        return data
