"""Logging configuration module for LLM call logging.

This module provides utilities to set up and configure logging
for LLM API calls during development.
"""

import logging
import sys
from pathlib import Path

from excel_agent.config import LoggingConfig


def setup_logging(config: LoggingConfig) -> logging.Logger:
    """Set up and return the LLM logger.

    Args:
        config: Logging configuration

    Returns:
        Configured logger instance
    """
    logger = logging.getLogger("excel_agent.llm")

    # Clear existing handlers
    logger.handlers.clear()

    # Set log level
    level = getattr(logging, config.level.upper(), logging.DEBUG)
    logger.setLevel(level)

    # Create formatter
    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Console output
    if config.output in ("console", "both"):
        console_handler = logging.StreamHandler(sys.stderr)
        console_handler.setLevel(level)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    # File output
    if config.output in ("file", "both") and config.file_path:
        file_path = Path(config.file_path)
        file_path.parent.mkdir(parents=True, exist_ok=True)

        file_handler = logging.FileHandler(
            file_path,
            encoding="utf-8",
            mode="a",  # Append mode
        )
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    # Prevent propagation to root logger
    logger.propagate = False

    return logger
