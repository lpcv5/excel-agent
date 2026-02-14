"""Logging configuration module for LLM call logging."""

import logging
import sys
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

from agent.config import LoggingConfig

_APP_LOGGERS = [
    "agent.llm",
    "excel_tools",
    "app.fastapi",
    "app.agent",
    "app.excel",
    "app.stream",
    "app.analysis",
    "app.project",
]
_app_logging_initialized = False


def setup_app_logging(log_dir: Path | None = None) -> None:
    """Initialize all app loggers with a daily-rotating file handler.

    Safe to call multiple times — subsequent calls are no-ops.
    """
    global _app_logging_initialized
    if _app_logging_initialized:
        return
    _app_logging_initialized = True

    if log_dir is None:
        # src-python/agent/logging_config.py -> project root is 3 levels up
        log_dir = Path(__file__).parent.parent.parent / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"app_{timestamp}.log"

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(name)-20s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = TimedRotatingFileHandler(
        log_file,
        when="midnight",
        backupCount=7,
        encoding="utf-8",
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    for name in _APP_LOGGERS:
        lg = logging.getLogger(name)
        lg.setLevel(logging.DEBUG)
        lg.propagate = False
        if not lg.handlers:
            lg.addHandler(file_handler)


def setup_logging(config: LoggingConfig) -> logging.Logger:
    """Configure agent.llm and excel_tools loggers.

    If setup_app_logging() has already been called, the file handler is already
    attached — this function only adds a console handler when requested.
    """
    level = getattr(logging, config.level.upper(), logging.DEBUG)

    llm_logger = logging.getLogger("agent.llm")
    llm_logger.setLevel(level)
    llm_logger.propagate = False

    tools_logger = logging.getLogger("excel_tools")
    tools_logger.setLevel(level)
    tools_logger.propagate = False

    if config.output in ("console", "both"):
        formatter = logging.Formatter(
            fmt="%(asctime)s | %(name)-20s | %(levelname)-8s | %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        console_handler = logging.StreamHandler(sys.stderr)
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        # Avoid duplicate console handlers on repeated calls
        if not any(isinstance(h, logging.StreamHandler) and h.stream is sys.stderr for h in llm_logger.handlers):
            llm_logger.addHandler(console_handler)
        if not any(isinstance(h, logging.StreamHandler) and h.stream is sys.stderr for h in tools_logger.handlers):
            tools_logger.addHandler(console_handler)

    return llm_logger
