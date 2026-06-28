"""Tests for sap_automation.core.logger — Logging setup."""

from __future__ import annotations

import logging
from pathlib import Path

from sap_automation.core.logger import (
    _ExecutionContextFilter,
    _SecretFilter,
    get_logger,
    set_run_id,
    setup_logging,
)


class TestSetupLogging:
    """Test logging initialization."""

    def test_setup_creates_logger(self, tmp_path: Path) -> None:
        import sap_automation.core.logger as logger_mod
        logger_mod._initialized = False  # Reset for test

        root = setup_logging(log_dir=tmp_path, level="DEBUG")
        assert root is not None
        assert root.name == "sap_automation"
        assert root.level == logging.DEBUG

    def test_idempotent(self, tmp_path: Path) -> None:
        import sap_automation.core.logger as logger_mod
        logger_mod._initialized = False

        logger1 = setup_logging(log_dir=tmp_path, level="INFO")
        logger2 = setup_logging(log_dir=tmp_path, level="DEBUG")

        # Should return same logger, not re-add handlers
        assert logger1 is logger2

    def test_creates_log_dir(self, tmp_path: Path) -> None:
        log_dir = tmp_path / "subdir" / "logs"
        import sap_automation.core.logger as logger_mod
        logger_mod._initialized = False

        setup_logging(log_dir=log_dir)
        assert log_dir.exists()

    def test_get_child_logger(self) -> None:
        logger = get_logger("test_child")
        assert logger.name == "sap_automation.test_child"


class TestExecutionContextFilter:
    """Test run_id injection."""

    def test_run_id_injected(self) -> None:
        filt = _ExecutionContextFilter(run_id="abc123")
        record = logging.LogRecord(
            name="test", level=logging.INFO, pathname="", lineno=0,
            msg="test", args=(), exc_info=None,
        )
        assert filt.filter(record)
        assert record.run_id == "abc123"  # type: ignore[attr-defined]


class TestSecretFilter:
    """Test secret redaction."""

    def test_redacts_password(self) -> None:
        filt = _SecretFilter()
        record = logging.LogRecord(
            name="test", level=logging.INFO, pathname="", lineno=0,
            msg="Connecting with password=secret123",
            args=(), exc_info=None,
        )
        assert filt.filter(record)
        assert "REDACTED" in record.msg

    def test_redacts_token(self) -> None:
        filt = _SecretFilter()
        record = logging.LogRecord(
            name="test", level=logging.INFO, pathname="", lineno=0,
            msg="Using api_token=abc123",
            args=(), exc_info=None,
        )
        assert filt.filter(record)
        assert "REDACTED" in record.msg

    def test_passes_normal_messages(self) -> None:
        filt = _SecretFilter()
        record = logging.LogRecord(
            name="test", level=logging.INFO, pathname="", lineno=0,
            msg="Normal log message",
            args=(), exc_info=None,
        )
        assert filt.filter(record)
        assert record.msg == "Normal log message"


class TestSetRunId:
    """Test run_id updates."""

    def test_set_run_id(self) -> None:
        from sap_automation.core.logger import _execution_filter
        set_run_id("test_run_123")
        assert _execution_filter.run_id == "test_run_123"
