"""Tests for sap_automation.cli — CLI commands."""

from __future__ import annotations

import pytest
from click.testing import CliRunner

from sap_automation.cli import main


@pytest.fixture
def runner() -> CliRunner:
    return CliRunner()


class TestCLI:
    """Test CLI commands."""

    def test_version(self, runner: CliRunner) -> None:
        result = runner.invoke(main, ["--version"])
        assert result.exit_code == 0
        assert "1.0.0" in result.output

    def test_help(self, runner: CliRunner) -> None:
        result = runner.invoke(main, ["--help"])
        assert result.exit_code == 0
        assert "SAP Automation" in result.output

    def test_run_missing_transaction(self, runner: CliRunner) -> None:
        result = runner.invoke(main, ["run"])
        assert result.exit_code != 0

    def test_run_missing_input(self, runner: CliRunner) -> None:
        result = runner.invoke(main, ["run", "-t", "va23"])
        assert result.exit_code != 0

    def test_validate_missing_config(self, runner: CliRunner) -> None:
        result = runner.invoke(main, ["validate", "-c", "/nonexistent.yaml"])
        assert result.exit_code != 0

    def test_list_transactions_missing_config(self, runner: CliRunner) -> None:
        result = runner.invoke(main, ["list-transactions", "-c", "/nonexistent.yaml"])
        assert result.exit_code != 0

    def test_init_creates_dirs(self, runner: CliRunner, tmp_path) -> None:
        import os
        original_cwd = os.getcwd()
        try:
            os.chdir(tmp_path)
            result = runner.invoke(main, ["init", "test_output"])
            assert result.exit_code == 0
            assert (tmp_path / "test_output").exists()
        finally:
            os.chdir(original_cwd)
