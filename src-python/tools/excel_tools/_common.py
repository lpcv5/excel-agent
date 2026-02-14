"""Common utilities for Excel tools layer."""

import json
import logging
import time
import traceback
from functools import wraps
from typing import Any, Callable

from libs.excel_com.com_thread import run_on_com_thread
from libs.excel_com.errors import ExcelComError

logger = logging.getLogger("app.excel")


def format_result(success: bool, results: list[dict], file_saved: bool = False) -> str:
    """Format tool results as JSON string."""
    return json.dumps(
        {"success": success, "results": results, "file_saved": file_saved},
        ensure_ascii=False,
        default=str,
    )


def safe_excel_call(func: Callable) -> Callable:
    """Decorator that dispatches the call onto the dedicated COM thread
    and catches Excel errors, returning structured error strings."""

    @wraps(func)
    def wrapper(*args: Any, **kwargs: Any) -> str:
        t0 = time.perf_counter()
        try:
            result = run_on_com_thread(func, *args, **kwargs)
            elapsed = (time.perf_counter() - t0) * 1000
            logger.info("Tool %s completed in %.0fms", func.__name__, elapsed)
            return result
        except ExcelComError as e:
            elapsed = (time.perf_counter() - t0) * 1000
            logger.error(
                "Tool %s ExcelComError in %.0fms: %s",
                func.__name__,
                elapsed,
                e.user_message,
            )
            return format_result(
                False, [{"status": "error", "message": e.user_message}]
            )
        except Exception as e:
            elapsed = (time.perf_counter() - t0) * 1000
            logger.error(
                "Tool %s unexpected error in %.0fms: %s\n%s",
                func.__name__,
                elapsed,
                e,
                traceback.format_exc(),
            )
            return format_result(
                False, [{"status": "error", "message": f"内部错误: {e}"}]
            )

    return wrapper
