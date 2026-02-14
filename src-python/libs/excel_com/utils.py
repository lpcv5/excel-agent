"""Utility functions for Excel COM operations."""

import logging
import os
import time
from typing import Any, Callable, TypeVar

import pythoncom

from .constants import HRESULT_MAP, TRANSIENT_HRESULTS

logger = logging.getLogger("app.excel")

F = TypeVar("F", bound=Callable[..., Any])


def normalize_path(path: str) -> str:
    """Normalize a file path for consistent registry lookups."""
    return os.path.normcase(os.path.realpath(path))


def pump_messages() -> None:
    """Pump COM message queue to prevent channel blocking."""
    pythoncom.PumpWaitingMessages()


def hresult_to_message(hresult: int) -> str:
    """Map an HRESULT code to a human-readable message."""
    return HRESULT_MAP.get(hresult, f"COM error 0x{hresult & 0xFFFFFFFF:08X}")


def is_transient(exc: Exception) -> bool:
    """Check if a COM exception is transient and safe to retry."""
    hr = getattr(exc, "hresult", None)
    if isinstance(hr, int):
        return hr in TRANSIENT_HRESULTS
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0] in TRANSIENT_HRESULTS
    return False


def com_retry(max_retries: int = 2, delay: float = 0.5):
    """Decorator that retries on transient COM errors."""

    def decorator(func: F) -> F:
        def wrapper(*args: Any, **kwargs: Any) -> Any:
            last_exc: Exception | None = None
            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_exc = e
                    if is_transient(e) and attempt < max_retries:
                        logger.warning(
                            "Transient COM error in %s (attempt %d/%d): %s",
                            func.__name__,
                            attempt + 1,
                            max_retries,
                            e,
                        )
                        time.sleep(delay * (attempt + 1))
                        pump_messages()
                        continue
                    raise
            raise last_exc  # type: ignore[misc]

        return wrapper  # type: ignore[return-value]

    return decorator
