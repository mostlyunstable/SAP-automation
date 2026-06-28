"""Tests for sap_automation.core.retry — Retry decorator."""

from __future__ import annotations

import pytest

from sap_automation.core.exceptions import SAPBusyError, TransactionError
from sap_automation.core.retry import retry_on_exception


class TestRetryDecorator:
    """Test retry behavior."""

    def test_success_no_retry(self) -> None:
        call_count = 0

        @retry_on_exception(max_attempts=3, delay=0.01)
        def success_func() -> str:
            nonlocal call_count
            call_count += 1
            return "ok"

        assert success_func() == "ok"
        assert call_count == 1

    def test_retry_then_success(self) -> None:
        call_count = 0

        @retry_on_exception(max_attempts=3, delay=0.01)
        def flaky_func() -> str:
            nonlocal call_count
            call_count += 1
            if call_count < 3:
                raise SAPBusyError("busy")
            return "ok"

        assert flaky_func() == "ok"
        assert call_count == 3

    def test_all_attempts_fail(self) -> None:
        @retry_on_exception(max_attempts=2, delay=0.01)
        def always_fails() -> None:
            raise ValueError("always fails")

        with pytest.raises(ValueError, match="always fails"):
            always_fails()

    def test_only_catches_specified_exceptions(self) -> None:
        call_count = 0

        @retry_on_exception(
            max_attempts=3, delay=0.01, exceptions=(SAPBusyError,)
        )
        def wrong_exception() -> None:
            nonlocal call_count
            call_count += 1
            raise TransactionError("not caught")

        with pytest.raises(TransactionError):
            wrong_exception()
        assert call_count == 1  # No retry for wrong exception type

    def test_invalid_max_attempts(self) -> None:
        with pytest.raises(ValueError, match="max_attempts must be >= 1"):
            retry_on_exception(max_attempts=0)

    def test_invalid_delay(self) -> None:
        with pytest.raises(ValueError, match="delay must be >= 0"):
            retry_on_exception(delay=-1)

    def test_preserves_function_name(self) -> None:
        @retry_on_exception(max_attempts=1, delay=0.01)
        def my_function() -> None:
            pass

        assert my_function.__name__ == "my_function"

    def test_single_attempt(self) -> None:
        call_count = 0

        @retry_on_exception(max_attempts=1, delay=0.01)
        def fails_once() -> None:
            nonlocal call_count
            call_count += 1
            raise ValueError("fail")

        with pytest.raises(ValueError):
            fails_once()
        assert call_count == 1
