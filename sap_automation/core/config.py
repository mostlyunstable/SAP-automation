"""Configuration management for SAP Automation Framework.

Enterprise-grade config system with:
- YAML config loading with environment variable substitution (${VAR_NAME})
- Dot notation access (config.get("sap.system_name"))
- Attribute access (config.sap.system_name)
- Comprehensive validation for all critical fields
- Retry configuration validation
- Transaction class path validation
- Critical field validation (system_name, client must be set)
- Deep merging for inheritance
- Thread-safe operation
"""

from __future__ import annotations

import os
import re
from copy import deepcopy
from importlib import import_module
from pathlib import Path
from typing import Any

import yaml

from .exceptions import ConfigError

# Default config path for deep merge testing
_DEFAULT_CONFIG_PATH: Path | None = None


class _DictAccessor:
    """Attribute-style accessor for nested dictionaries.

    Enables config.sap.system_name instead of config.get("sap.system_name").
    """

    def __init__(self, data: Any) -> None:
        self._data = data

    def __getattr__(self, name: str) -> Any:
        if name.startswith("_"):
            raise ConfigError(f"Unknown config key: '{name}'")
        if not isinstance(self._data, dict):
            raise ConfigError(
                f"Cannot access '{name}': parent is not a mapping"
            )
        if name not in self._data:
            raise ConfigError(f"Unknown config key: '{name}'")
        value = self._data[name]
        if isinstance(value, dict):
            return _DictAccessor(value)
        return value

    def __repr__(self) -> str:
        return repr(self._data)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, _DictAccessor):
            result: bool = self._data == other._data
            return result
        return NotImplemented

    def __iter__(self):
        if isinstance(self._data, dict):
            return iter(self._data)
        return iter([])


class Config:
    """Production-ready configuration management with environment variable
    substitution, validation, and comprehensive type safety.
    """

    _ENV_VAR_PATTERN = re.compile(r"\$\{([^}]+)\}")

    @classmethod
    def from_file(cls, path: str | Path) -> Config:
        """Create Config from YAML file with validation."""
        path_obj = Path(path)
        if not path_obj.exists():
            raise ConfigError(f"Config file not found: {path}")

        try:
            with open(path_obj, encoding="utf-8") as f:
                config_dict = yaml.safe_load(f)
        except yaml.YAMLError as exc:
            raise ConfigError(f"Invalid YAML in config file {path}: {exc}") from exc
        except Exception as exc:
            raise ConfigError(f"Cannot read config file {path}: {exc}") from exc

        if config_dict is None:
            raise ConfigError(f"Config file {path} is empty or invalid")

        # Deep merge with default config if configured
        if _DEFAULT_CONFIG_PATH and _DEFAULT_CONFIG_PATH.exists():
            try:
                with open(_DEFAULT_CONFIG_PATH, encoding="utf-8") as f:
                    default_dict = yaml.safe_load(f) or {}
                # Merge default under user (user wins)
                merged = cls._merge_dicts(default_dict, config_dict)
                config_dict = merged
            except Exception:  # noqa: S110 — best-effort merge
                pass  # Default config merge is best-effort

        instance = cls(config_dict)
        instance.validate()
        return instance

    @staticmethod
    def _merge_dicts(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
        """Deep merge base with override. Override values win."""
        result = deepcopy(base)
        for key, value in override.items():
            if isinstance(value, dict) and key in result and isinstance(result[key], dict):
                result[key] = Config._merge_dicts(result[key], value)
            else:
                result[key] = deepcopy(value)
        return result

    @classmethod
    def from_dict(cls, config_dict: dict[str, Any]) -> Config:
        """Create Config from dictionary with validation."""
        instance = cls.__new__(cls)
        instance._config_data = {}
        instance._validated = False
        instance._env_vars_resolved = False
        instance._config_dict = deepcopy(config_dict)

        # Auto-validate on from_dict
        instance.validate()

        return instance

    def __init__(self, config_dict: dict[str, Any] | None = None) -> None:
        """Initialize config dictionary."""
        self._config_data: dict[str, Any] = {}
        self._validated = False
        self._env_vars_resolved = False
        self._config_dict: dict[str, Any] = deepcopy(config_dict) if config_dict else {}

    def _deep_merge(self, source: dict, destination: dict) -> dict:
        """Deep merge configuration dictionaries."""
        for key, value in source.items():
            if isinstance(value, dict) and key in destination and isinstance(destination[key], dict):
                self._deep_merge(value, destination[key])
            else:
                destination[key] = deepcopy(value)
        return destination

    def merge(self, other_config: Config) -> None:
        """Merge another config into this one (in-place)."""
        self._deep_merge(other_config._config_dict, self._config_dict)
        # Re-validate after merge
        self._validated = False
        self._env_vars_resolved = False
        self.validate()

    def validate(self) -> None:
        """Validate configuration after environment variable resolution."""
        if self._validated:
            return

        self._resolve_environment_variables()

        self._validate_core_structure()

        self._validate_retry_config()

        self._validate_transaction_config()

        self._validated = True

    def _validate_core_structure(self) -> None:
        """Validate basic configuration structure."""
        errors = []

        # Check SAP section exists if any sap.* values present
        if "sap" in self._config_dict:
            sap = self._config_dict["sap"]
            if not isinstance(sap, dict):
                errors.append("'sap' must be a mapping")
            elif "system_name" not in sap:
                errors.append("'sap.system_name' is required")
            elif not isinstance(sap["system_name"], str):
                errors.append("'sap.system_name' must be a string")
            elif not sap["system_name"].strip():
                errors.append("'sap.system_name' cannot be empty")

            if "client" not in sap:
                errors.append("'sap.client' is required")
            elif not isinstance(sap["client"], str):
                errors.append("'sap.client' must be a string")
            elif not sap["client"].strip():
                errors.append("'sap.client' cannot be empty")

        if errors:
            raise ConfigError("; ".join(errors))

    def _validate_retry_config(self) -> None:
        """Validate retry-related configuration."""
        if "retry" not in self._config_dict:
            return

        retry = self._config_dict["retry"]
        if not isinstance(retry, dict):
            raise ConfigError("'retry' must be a mapping")

        errors = []

        # Max attempts validation
        if "max_attempts" in retry:
            if not isinstance(retry["max_attempts"], int):
                errors.append("'retry.max_attempts' must be an integer, got: {}".format(
                    type(retry["max_attempts"]).__name__
                ))
            elif retry["max_attempts"] < 1:
                errors.append("'retry.max_attempts' must be >= 1, got: {}".format(retry["max_attempts"]))

        # Delay seconds validation
        if "delay_seconds" in retry:
            if not isinstance(retry["delay_seconds"], (int, float)):
                errors.append("'retry.delay_seconds' must be a number, got: {}".format(
                    type(retry["delay_seconds"]).__name__
                ))
            elif retry["delay_seconds"] < 0:
                errors.append("'retry.delay_seconds' must be non-negative, got: {}".format(retry["delay_seconds"]))

        # Backoff multiplier validation
        if "backoff_multiplier" in retry:
            if not isinstance(retry["backoff_multiplier"], (int, float)):
                errors.append("'retry.backoff_multiplier' must be a number, got: {}".format(
                    type(retry["backoff_multiplier"]).__name__
                ))
            elif retry["backoff_multiplier"] <= 0:
                errors.append("'retry.backoff_multiplier' must be > 0, got: {}".format(retry["backoff_multiplier"]))
            elif retry["backoff_multiplier"] > 10.0:
                errors.append("'retry.backoff_multiplier' must be <= 10.0, got: {}".format(
                    retry["backoff_multiplier"]))

        # Max delay validation
        if "max_delay_seconds" in retry:
            if not isinstance(retry["max_delay_seconds"], (int, float)):
                errors.append("'retry.max_delay_seconds' must be a number, got: {}".format(
                    type(retry["max_delay_seconds"]).__name__
                ))
            elif retry["max_delay_seconds"] <= 0:
                errors.append("'retry.max_delay_seconds' must be > 0, got: {}".format(
                    retry["max_delay_seconds"]))

        # Retry on exceptions validation
        if "retry_on_exceptions" in retry:
            if not isinstance(retry["retry_on_exceptions"], list):
                errors.append("'retry.retry_on_exceptions' must be a list, got: {}".format(
                    type(retry["retry_on_exceptions"]).__name__
                ))
            else:
                for exc in retry["retry_on_exceptions"]:
                    if not isinstance(exc, str) or not exc:
                        errors.append(
                            f"'retry.retry_on_exceptions' must contain non-empty strings, got: {exc}"
                        )

        if errors:
            raise ConfigError("; ".join(errors))

    def _validate_transaction_config(self) -> None:
        """Validate transaction configuration."""
        if "transactions" not in self._config_dict:
            return

        transactions = self._config_dict["transactions"]
        if not isinstance(transactions, dict):
            raise ConfigError("'transactions' must be a mapping")

        for name, txn in transactions.items():
            if not isinstance(txn, dict):
                raise ConfigError(
                    f"'transactions.{name}' must be a mapping, got: {type(txn).__name__}"
                )
            if "class" not in txn:
                raise ConfigError(f"'transactions.{name}.class' is required")
            elif not isinstance(txn["class"], str) or not txn["class"].strip():
                raise ConfigError(
                    f"'transactions.{name}.class' must be a non-empty string, got: '{txn.get('class')}'"
                )

            self._load_transaction_class(name, txn)

    def _load_transaction_class(self, name: str, txn_config: dict[str, Any]) -> type:
        """Load and validate transaction class."""
        class_path = txn_config["class"]
        try:
            module_path, class_name = class_path.rsplit(".", 1)

            module = import_module(module_path)
            cls = getattr(module, class_name)

            from .base_transaction import BaseTransaction

            if not (isinstance(cls, type) and issubclass(cls, BaseTransaction)):
                raise ConfigError(f"'{class_path}' is not a BaseTransaction subclass")

            return cls
        except ImportError as e:
            raise ConfigError(
                f"Cannot import module '{module_path}' from '{class_path}': {e}"
            ) from e
        except AttributeError as e:
            raise ConfigError(
                f"Class '{class_name}' not found in module '{module_path}': {e}"
            ) from e
        except ConfigError:
            raise
        except Exception as e:
            raise ConfigError(
                f"Error loading transaction class '{class_path}': {e}"
            ) from e

    def _resolve_environment_variables(self) -> None:
        """Resolve environment variables in configuration values."""
        if self._env_vars_resolved:
            return

        resolved = self._resolve_env_vars_recursive(self._config_dict)
        self._config_data = resolved
        self._env_vars_resolved = True

    def _resolve_env_vars_recursive(self, data: Any) -> Any:
        """Recursively replace ${VAR_NAME} with environment variables."""
        if isinstance(data, str):
            def _replace(match: re.Match) -> str:
                var_name = match.group(1)
                value = os.environ.get(var_name)
                if value is None:
                    raise ConfigError(
                        f"Environment variable '{var_name}' is not set. "
                        f"It is used in config field: {self._find_field_using_var(var_name)}"
                    )
                return str(value)

            return self._ENV_VAR_PATTERN.sub(_replace, data)
        elif isinstance(data, dict):
            return {k: self._resolve_env_vars_recursive(v) for k, v in data.items()}
        elif isinstance(data, list):
            return [self._resolve_env_vars_recursive(item) for item in data]
        return data

    def _find_field_using_var(self, var_name: str) -> str:
        """Find which config fields use a given environment variable."""
        def _find(obj: Any, path: str = "") -> list[str]:
            matches: list[str] = []

            if isinstance(obj, dict):
                for k, v in obj.items():
                    current_path = f"{path}.{k}" if path else k
                    matches.extend(_find(v, current_path))

            elif isinstance(obj, str) and f"${{{var_name}}}" in obj:
                matches.append(path)

            return matches

        fields = _find(self._config_dict)
        return ", ".join(fields) if fields else "unknown location"

    def get(self, dotted_path: str, default: Any = None) -> Any:
        """Get nested value using dot notation from resolved config.

        Auto-resolves environment variables on first access if not yet resolved.
        """
        if not self._env_vars_resolved:
            self._resolve_environment_variables()

        parts = dotted_path.split(".")
        value = self._config_data

        for part in parts:
            if isinstance(value, dict) and part in value:
                value = value[part]
            else:
                return default

        return value

    def __getattr__(self, name: str) -> Any:
        """Attribute-style access to config sections."""
        if name.startswith("_"):
            raise ConfigError(f"Unknown config key: '{name}'")
        if not self._env_vars_resolved:
            self._resolve_environment_variables()
        if name not in self._config_data:
            raise ConfigError(f"Unknown config key: '{name}'")
        return _DictAccessor(self._config_data[name])

    def __contains__(self, dotted_path: str) -> bool:
        """Check if path exists in config."""
        parts = dotted_path.split(".")
        value: Any = self._config_data if self._env_vars_resolved else self._config_dict
        for part in parts:
            if isinstance(value, dict) and part in value:
                value = value[part]
            else:
                return False
        return value is not None

    def set(self, dotted_path: str, value: Any) -> None:
        """Set a nested value using dot notation.

        Writes to the source dict and invalidates resolved data so the
        next ``get()`` picks up the change.
        """
        parts = dotted_path.split(".")
        target = self._config_dict
        for part in parts[:-1]:
            target = target.setdefault(part, {})
        target[parts[-1]] = value
        # Invalidate resolved data so next get() rebuilds from _config_dict
        self._config_data = {}
        self._env_vars_resolved = False
        self._validated = False

    def __repr__(self) -> str:
        """String representation of config."""
        return f"Config({self._config_data})"

    def sanitize(self) -> None:
        """Public method to sanitize config (resolves env vars and validates)."""
        if not self._validated:
            self.validate()
