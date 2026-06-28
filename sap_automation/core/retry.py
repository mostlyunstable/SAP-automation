"""Retry decorator with exponential backoff and selective exception handling."""

from __future__ import annotations

import functools
import time
from collections.abc import Callable
from typing import Any

from .logger import get_logger

log = get_logger("retry")


def retry_on_exception(
    max_attempts: int = 3,
    delay: float = 1.0,
    backoff: float = 2.0,
    max_delay: float = 60.0,
    exceptions: tuple[type[BaseException], ...] = (Exception,),
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Decorator that retries a function on specified exceptions.

    Args:
        max_attempts: Maximum number of attempts (>= 1).
        delay: Initial delay between retries in seconds (>= 0).
        backoff: Multiplier for delay after each retry (> 0).
        max_delay: Maximum delay between retries in seconds (> 0).
        exceptions: Tuple of exception types to catch and retry.

    Returns:
        Decorated function with retry logic.

    Raises:
        ValueError: If parameters are invalid.
    """
    if max_attempts < 1:
        raise ValueError("max_attempts must be >= 1")
    if delay < 0:
        raise ValueError("delay must be >= 0")
    if backoff <= 0:
        raise ValueError("backoff must be > 0")
    if max_delay <= 0:
        raise ValueError("max_delay must be > 0")

    def decorator(func: Callable[..., Any]) -> Callable[..., Any]:
        @functools.wraps(func)
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            current_delay = delay
            last_exception: BaseException | None = None

            for attempt in range(1, max_attempts + 1):
                try:
                    return func(*args, **kwargs)
                except exceptions as exc:
                    last_exception = exc
                    if attempt == max_attempts:
                        log.error(
                            "All %d attempts failed for %s: %s",
                            max_attempts,
                            func.__qualname__,
                            exc,
                        )
                        raise
                    log.warning(
                        "Attempt %d/%d failed for %s: %s — retrying in %.1fs",
                        attempt,
                        max_attempts,
                        func.__qualname__,
                        exc,
                        current_delay,
                    )
                    time.sleep(current_delay)
                    current_delay = min(current_delay * backoff, max_delay)

            # Should not reach here, but safety fallback
            if last_exception:
                raise last_exception
            raise RuntimeError("Retry logic failed unexpectedly")

        return wrapper

    return decorator
