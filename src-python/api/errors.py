"""Structured error codes and FastAPI exception handlers."""

from enum import Enum

from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse


class ErrorCode(str, Enum):
    AGENT_BUSY = "AGENT_BUSY"
    PROJECT_SWITCH_DURING_STREAM = "PROJECT_SWITCH_DURING_STREAM"
    SETTINGS_INVALID = "SETTINGS_INVALID"
    COM_TIMEOUT = "COM_TIMEOUT"
    COM_EXCEL_UNAVAILABLE = "COM_EXCEL_UNAVAILABLE"
    FILE_NOT_FOUND = "FILE_NOT_FOUND"
    PROJECT_NOT_FOUND = "PROJECT_NOT_FOUND"
    UNAUTHORIZED = "UNAUTHORIZED"
    INTERNAL_ERROR = "INTERNAL_ERROR"


class AppError(Exception):
    def __init__(self, code: ErrorCode, message: str, status: int = 400) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.status = status


def register_exception_handlers(app: FastAPI) -> None:
    @app.exception_handler(AppError)
    async def app_error_handler(request: Request, exc: AppError) -> JSONResponse:
        return JSONResponse(
            status_code=exc.status,
            content={"error": exc.code, "message": exc.message},
        )
