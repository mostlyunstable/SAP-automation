"""Tests for sap_automation.orchestrator — Batch orchestration."""

from __future__ import annotations

import pytest

from sap_automation.core.base_transaction import TransactionResult
from sap_automation.orchestrator import Orchestrator, RunSummary


class TestRunSummary:
    """Test the RunSummary dataclass."""

    def test_defaults(self) -> None:
        summary = RunSummary()
        assert summary.total_records == 0
        assert summary.successful == 0
        assert summary.failed == 0
        assert summary.skipped == 0
        assert summary.run_id  # Should have auto-generated ID
        assert summary.started_at

    def test_success_rate(self) -> None:
        summary = RunSummary(total_records=10, successful=7)
        assert summary.success_rate == 70.0

    def test_success_rate_zero_records(self) -> None:
        summary = RunSummary(total_records=0)
        assert summary.success_rate == 0.0

    def test_report(self) -> None:
        summary = RunSummary(
            transaction_name="va23",
            total_records=5,
            successful=4,
            failed=1,
        )
        report = summary.report()
        assert "AUTOMATION RUN SUMMARY" in report
        assert "5" in report
        assert "4" in report
        assert "80.0%" in report

    def test_failed_documents(self) -> None:
        results = [
            TransactionResult("VA23", "001", success=True),
            TransactionResult("VA23", "002", success=False, message="Error"),
            TransactionResult("VA23", "003", success=True),
            TransactionResult("VA23", "004", success=False, message="Error"),
        ]
        summary = RunSummary(results=results)
        assert summary.failed_documents == ["002", "004"]

    def test_to_dict(self) -> None:
        summary = RunSummary(transaction_name="va23", total_records=5)
        d = summary.to_dict()
        assert d["transaction_name"] == "va23"
        assert d["total_records"] == 5
        assert "run_id" in d
        assert "success_rate" in d


class TestOrchestrator:
    """Test orchestrator initialization and handler loading."""

    def test_load_handler_invalid_path(self) -> None:
        from sap_automation.core.config import Config
        from sap_automation.core.exceptions import ConfigError

        cfg = Config.from_dict({"sap": {"system_name": "PRD", "client": "100"}})

        with pytest.raises(ConfigError):
            Orchestrator._load_handler("invalid.path", cfg)

    def test_load_handler_wrong_namespace(self) -> None:
        from sap_automation.core.config import Config
        from sap_automation.core.exceptions import ConfigError

        cfg = Config.from_dict({"sap": {"system_name": "PRD", "client": "100"}})

        with pytest.raises(ConfigError, match="must start with"):
            Orchestrator._load_handler(
                "sap_automation.core.config.NonExistentClass", cfg
            )

    def test_load_handler_nonexistent_class(self) -> None:
        from sap_automation.core.config import Config
        from sap_automation.core.exceptions import ConfigError

        cfg = Config.from_dict({"sap": {"system_name": "PRD", "client": "100"}})

        with pytest.raises(ConfigError):
            Orchestrator._load_handler(
                "sap_automation.transactions.nonexistent.Handler", cfg
            )

    def test_load_handler_not_subclass(self) -> None:
        """Test that importing a non-BaseTransaction class raises ConfigError.

        We use Config (which is in sap_automation.core, not transactions)
        to verify the namespace check rejects it before we even get to
        the subclass check.
        """
        from sap_automation.core.config import Config
        from sap_automation.core.exceptions import ConfigError

        cfg = Config.from_dict({"sap": {"system_name": "PRD", "client": "100"}})

        with pytest.raises(ConfigError, match="must start with"):
            Orchestrator._load_handler(
                "sap_automation.core.config.Config",
                cfg,
            )

    def test_load_handler_not_subclass_actual(self) -> None:
        """Test that a non-BaseTransaction class in the transactions namespace is rejected."""
        import sys
        import types

        from sap_automation.core.config import Config
        from sap_automation.core.exceptions import ConfigError

        mod = types.ModuleType("sap_automation.transactions._test_helper")

        class NotATransaction:
            pass

        mod.Handler = NotATransaction
        sys.modules["sap_automation.transactions._test_helper"] = mod
        try:
            cfg = Config.from_dict({"sap": {"system_name": "PRD", "client": "100"}})
            with pytest.raises(ConfigError, match="not a BaseTransaction"):
                Orchestrator._load_handler(
                    "sap_automation.transactions._test_helper.Handler",
                    cfg,
                )
        finally:
            del sys.modules["sap_automation.transactions._test_helper"]
