"""Tests for sap_automation.core.config — Config loading, access, validation."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
import yaml

from sap_automation.core.config import Config, _DictAccessor
from sap_automation.core.exceptions import ConfigError


@pytest.fixture
def sample_data() -> dict[str, Any]:
    return {
        "sap": {"system_name": "PRD", "client": "100", "language": "EN"},
        "paths": {"output_dir": "./output", "log_dir": "./logs"},
        "retry": {"max_attempts": 3, "delay_seconds": 2.0},
    }


@pytest.fixture
def sample_config(sample_data: dict[str, Any]) -> Config:
    return Config.from_dict(sample_data)


class TestConfigAccess:
    """Test dot-notation and attribute access."""

    def test_get_simple(self, sample_config: Config) -> None:
        assert sample_config.get("sap.system_name") == "PRD"

    def test_get_nested(self, sample_config: Config) -> None:
        assert sample_config.get("sap.client") == "100"

    def test_get_default(self, sample_config: Config) -> None:
        assert sample_config.get("nonexistent.key", "fallback") == "fallback"

    def test_get_missing_returns_none(self, sample_config: Config) -> None:
        assert sample_config.get("missing.path") is None

    def test_attribute_access(self, sample_config: Config) -> None:
        assert sample_config.sap.system_name == "PRD"

    def test_attribute_access_missing(self, sample_config: Config) -> None:
        with pytest.raises(ConfigError):
            _ = sample_config.nonexistent

    def test_contains(self, sample_config: Config) -> None:
        assert "sap.system_name" in sample_config
        assert "missing.key" not in sample_config

    def test_repr(self, sample_config: Config) -> None:
        assert "Config" in repr(sample_config)


class TestConfigLoading:
    """Test YAML loading and merging."""

    def test_from_file(self, tmp_path: Path) -> None:
        config_file = tmp_path / "test.yaml"
        config_file.write_text(yaml.dump({"sap": {"system_name": "DEV", "client": "200"}}))

        cfg = Config.from_file(config_file)
        assert cfg.get("sap.system_name") == "DEV"
        assert cfg.get("sap.client") == "200"

    def test_file_not_found(self) -> None:
        with pytest.raises(ConfigError, match="Config file not found"):
            Config.from_file("/nonexistent/config.yaml")

    def test_invalid_yaml(self, tmp_path: Path) -> None:
        config_file = tmp_path / "bad.yaml"
        config_file.write_text("{{{{invalid yaml:::")

        with pytest.raises(ConfigError, match="Invalid YAML"):
            Config.from_file(config_file)

    def test_deep_merge(self, tmp_path: Path) -> None:
        default = tmp_path / "default.yaml"
        default.write_text(yaml.dump({
            "sap": {"system_name": "PRD", "client": "100"},
            "retry": {"max_attempts": 3},
        }))

        user = tmp_path / "user.yaml"
        user.write_text(yaml.dump({
            "sap": {"client": "200"},
            "retry": {"max_attempts": 5},
        }))

        # Override default path
        import sap_automation.core.config as cfg_mod
        original = cfg_mod._DEFAULT_CONFIG_PATH
        cfg_mod._DEFAULT_CONFIG_PATH = default
        try:
            cfg = Config.from_file(user)
            assert cfg.get("sap.system_name") == "PRD"  # from default
            assert cfg.get("sap.client") == "200"  # overridden
            assert cfg.get("retry.max_attempts") == 5  # overridden
        finally:
            cfg_mod._DEFAULT_CONFIG_PATH = original


class TestConfigValidation:
    """Test startup validation."""

    def test_valid_config(self, sample_data: dict[str, Any]) -> None:
        # Should not raise
        Config.from_dict(sample_data)

    def test_missing_system_name(self) -> None:
        data = {"sap": {"client": "100"}}
        with pytest.raises(ConfigError, match=r"sap\.system_name"):
            Config.from_dict(data)

    def test_missing_client(self) -> None:
        data = {"sap": {"system_name": "PRD"}}
        with pytest.raises(ConfigError, match=r"sap\.client"):
            Config.from_dict(data)

    def test_invalid_retry_attempts(self) -> None:
        data = {
            "sap": {"system_name": "PRD", "client": "100"},
            "retry": {"max_attempts": -1},
        }
        with pytest.raises(ConfigError, match=r"retry\.max_attempts"):
            Config.from_dict(data)

    def test_invalid_retry_delay(self) -> None:
        data = {
            "sap": {"system_name": "PRD", "client": "100"},
            "retry": {"delay_seconds": "not_a_number"},
        }
        with pytest.raises(ConfigError, match=r"retry\.delay_seconds"):
            Config.from_dict(data)


class TestEnvironmentVariables:
    """Test ${ENV_VAR} resolution."""

    def test_env_var_resolution(self, monkeypatch: Any) -> None:
        monkeypatch.setenv("TEST_SAP_SYSTEM", "QAS")
        data = {
            "sap": {"system_name": "${TEST_SAP_SYSTEM}", "client": "100"},
        }
        cfg = Config.from_dict(data)
        assert cfg.get("sap.system_name") == "QAS"

    def test_unresolved_var_raises_error(self) -> None:
        from sap_automation.core.exceptions import ConfigError

        data = {
            "sap": {"system_name": "${NONEXISTENT_VAR_12345}", "client": "100"},
        }
        with pytest.raises(ConfigError, match="is not set"):
            Config.from_dict(data)


class TestDictAccessor:
    """Test the _DictAccessor helper."""

    def test_dict_access(self) -> None:
        accessor = _DictAccessor({"a": {"b": "value"}})
        assert accessor.a.b == "value"

    def test_key_error(self) -> None:
        accessor = _DictAccessor({"a": 1})
        with pytest.raises(ConfigError, match="Unknown config key"):
            _ = accessor.nonexistent

    def test_non_dict_error(self) -> None:
        accessor = _DictAccessor("not_a_dict")
        with pytest.raises(ConfigError, match="Cannot access"):
            _ = accessor.key

    def test_repr(self) -> None:
        accessor = _DictAccessor({"a": 1})
        assert repr(accessor) == "{'a': 1}"

    def test_eq(self) -> None:
        a = _DictAccessor({"x": 1})
        b = _DictAccessor({"x": 1})
        assert a == b
        assert a != _DictAccessor({"y": 2})
        assert a != "something_else"

    def test_iter(self) -> None:
        accessor = _DictAccessor({"a": 1, "b": 2})
        assert set(accessor) == {"a", "b"}

    def test_nonexistent_private_attr(self) -> None:
        accessor = _DictAccessor({})
        with pytest.raises(ConfigError, match="Unknown config key"):
            _ = accessor._nonexistent
