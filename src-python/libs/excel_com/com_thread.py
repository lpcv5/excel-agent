"""Dedicated COM thread for Excel operations.

All Excel COM calls must run on the same STA thread. Since LangGraph dispatches
tool calls via asyncio.to_thread() to arbitrary thread-pool threads, COM objects
created on one thread become invalid on another.

This module provides a single long-lived thread with CoInitialize() called once,
and a `run_on_com_thread(fn, *args, **kwargs)` helper that marshals any callable
onto that thread and returns the result.
"""

import logging
import queue
import threading
from typing import Any, Callable

import pythoncom

logger = logging.getLogger("app.excel")

_REQUEST = tuple  # (callable, args, kwargs, result_queue)


class _ComThread(threading.Thread):
    """Daemon thread that owns all COM objects."""

    def __init__(self) -> None:
        super().__init__(name="excel-com-thread", daemon=True)
        self._queue: queue.Queue[_REQUEST | None] = queue.Queue()
        self._ready = threading.Event()

    def run(self) -> None:
        pythoncom.CoInitialize()
        logger.info("COM thread started (thread=%s)", threading.current_thread().ident)
        self._ready.set()
        try:
            while True:
                item = self._queue.get()
                if item is None:
                    break
                fn, args, kwargs, result_q = item
                try:
                    result = fn(*args, **kwargs)
                    result_q.put((True, result))
                except Exception as e:
                    result_q.put((False, e))
        finally:
            pythoncom.CoUninitialize()
            logger.info("COM thread exiting")

    def submit(self, fn: Callable, *args: Any, timeout: float = 30.0, **kwargs: Any) -> Any:
        """Run *fn* on the COM thread and return its result (blocking)."""
        result_q: queue.Queue[tuple[bool, Any]] = queue.Queue()
        self._queue.put((fn, args, kwargs, result_q))
        try:
            ok, value = result_q.get(timeout=timeout)
        except queue.Empty:
            raise TimeoutError(f"Excel COM call timed out after {timeout}s")
        if ok:
            return value
        raise value

    def shutdown(self) -> None:
        self._queue.put(None)


# Module-level singleton
_thread: _ComThread | None = None
_lock = threading.Lock()


def _ensure_thread() -> _ComThread:
    global _thread
    if _thread is not None and _thread.is_alive():
        return _thread
    with _lock:
        if _thread is not None and _thread.is_alive():
            return _thread
        _thread = _ComThread()
        _thread.start()
        _thread._ready.wait(timeout=5.0)
        return _thread


def run_on_com_thread(fn: Callable, *args: Any, timeout: float = 30.0, **kwargs: Any) -> Any:
    """Execute *fn(*args, **kwargs)* on the dedicated COM thread.

    Blocks the calling thread until the result is available.
    Exceptions raised inside *fn* are re-raised on the caller.
    Raises TimeoutError if the call takes longer than *timeout* seconds.
    """
    t = _ensure_thread()
    return t.submit(fn, *args, timeout=timeout, **kwargs)


def shutdown_com_thread() -> None:
    """Shut down the COM thread (call on app exit)."""
    global _thread
    if _thread is not None and _thread.is_alive():
        _thread.shutdown()
        _thread.join(timeout=5.0)
        _thread = None
