"""PowerQuery / M Code operations for Excel COM."""

import logging
from typing import Any

from .errors import QueryError
from .utils import com_retry, pump_messages

logger = logging.getLogger("app.excel")


def list_queries(wb: Any) -> list[dict[str, str]]:
    """List all Power Query queries in the workbook."""
    result = []
    try:
        for i in range(1, wb.Queries.Count + 1):
            q = wb.Queries(i)
            result.append(
                {
                    "name": q.Name,
                    "description": q.Description or "",
                }
            )
    except Exception as e:
        logger.debug("Error listing queries: %s", e)
    return result
