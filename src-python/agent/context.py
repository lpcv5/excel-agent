"""Application context - centralized resource management."""

import atexit
import logging
import threading
from typing import Optional, TYPE_CHECKING

if TYPE_CHECKING:
    from agent.config import AgentConfig


class AppContext:
    _instance: Optional["AppContext"] = None
    _lock: threading.Lock = threading.Lock()
    _initialized: bool
    _config: Optional["AgentConfig"]
    _logger: Optional[logging.Logger]
    _excel_lock: threading.RLock

    def __new__(cls) -> "AppContext":
        with cls._lock:
            if cls._instance is None:
                cls._instance = super().__new__(cls)
                cls._instance._initialized = False
                cls._instance._config = None
                cls._instance._logger = None
                cls._instance._excel_lock = threading.RLock()
                atexit.register(cls._atexit_cleanup)
            return cls._instance

    @property
    def initialized(self) -> bool:
        return self._initialized

    @property
    def config(self) -> Optional["AgentConfig"]:
        return self._config

    @property
    def logger(self) -> Optional[logging.Logger]:
        return self._logger

    @property
    def excel_lock(self) -> threading.RLock:
        return self._excel_lock

    def initialize(self, config: "AgentConfig") -> None:
        if self._initialized:
            return
        self._config = config
        if config.logging.enabled:
            from agent.logging_config import setup_logging

            self._logger = setup_logging(config.logging)
        self._initialized = True

    async def ainitialize(self, config: "AgentConfig") -> None:
        self.initialize(config)

    def cleanup(self) -> None:
        import gc

        for _ in range(2):
            gc.collect()
        if self._logger is not None:
            self._logger.handlers.clear()
            self._logger = None
        self._initialized = False
        self._config = None

    async def acleanup(self) -> None:
        self.cleanup()

    async def __aenter__(self) -> "AppContext":
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb) -> None:
        self.cleanup()

    @classmethod
    def _atexit_cleanup(cls) -> None:
        if cls._instance is not None:
            try:
                cls._instance.cleanup()
            except Exception:
                pass

    def reset(self) -> None:
        self.cleanup()


def get_app_context() -> AppContext:
    return AppContext()
