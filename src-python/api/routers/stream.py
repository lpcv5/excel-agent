"""Stream endpoints: POST /api/stream, POST /api/stream/cancel"""

import asyncio
import json
import logging
import uuid
from typing import Optional

from fastapi import APIRouter, Depends
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from api.deps import get_core, get_project_service, verify_token
from services.project_service import ProjectService

logger = logging.getLogger("app.fastapi")
stream_logger = logging.getLogger("app.stream")

router = APIRouter(dependencies=[Depends(verify_token)])


class StreamRequest(BaseModel):
    message: str
    thread_id: Optional[str] = None


@router.post("/api/stream/cancel")
async def cancel_stream(project_svc: ProjectService = Depends(get_project_service)):
    project_svc.cancel_stream()
    stream_logger.info(">>> CANCEL requested")
    return {"ok": True}


@router.post("/api/stream")
async def stream_endpoint(
    req: StreamRequest,
    core=Depends(get_core),
    project_svc: ProjectService = Depends(get_project_service),
):
    async def generate():
        from agent.events import (
            ErrorEvent,
            QueryEndEvent,
            TextEvent,
            ThinkingEvent,
            TodoUpdateEvent,
            ToolCallStartEvent,
            ToolResultEvent,
        )

        def _sse(payload: dict) -> str:
            return f"data: {json.dumps(payload)}\n\n"

        stream_logger.info(
            ">>> REQUEST thread_id=%s message=%r", req.thread_id, req.message
        )

        _text_buf: list[str] = []
        _thinking_buf: list[str] = []

        def _flush_text():
            if _text_buf:
                stream_logger.info("<<< stream:text %r", "".join(_text_buf))
                _text_buf.clear()

        def _flush_thinking():
            if _thinking_buf:
                stream_logger.info("<<< stream:thinking %r", "".join(_thinking_buf))
                _thinking_buf.clear()

        def _close_workbooks_sync():
            try:
                from libs.excel_com import ExcelInstanceManager
                from libs.excel_com.com_thread import run_on_com_thread

                run_on_com_thread(ExcelInstanceManager().close_agent_workbooks)
            except Exception as e:
                logger.warning("close_agent_workbooks failed: %s", e)

        project_svc.mark_stream_start()
        try:
            async for event in core.astream_query(req.message, thread_id=req.thread_id):
                if isinstance(event, TextEvent):
                    if _thinking_buf:
                        _flush_thinking()
                    _text_buf.append(event.content or "")
                    yield _sse({"type": "stream:text", "token": event.content})

                elif isinstance(event, ThinkingEvent):
                    if _text_buf:
                        _flush_text()
                    _thinking_buf.append(event.content or "")
                    yield _sse({"type": "stream:thinking", "token": event.content})

                elif isinstance(event, ToolCallStartEvent):
                    _flush_text()
                    _flush_thinking()
                    call_id = event.tool_call_id or str(uuid.uuid4())
                    args: dict = {}
                    if event.tool_args:
                        try:
                            args = json.loads(event.tool_args)
                        except Exception:
                            args = {}
                    stream_logger.info("<<< tool:start id=%s name=%s", call_id, event.tool_name)
                    yield _sse({
                        "type": "tool:start",
                        "id": call_id,
                        "name": event.tool_name,
                        "args": args,
                    })

                elif isinstance(event, ToolResultEvent):
                    _flush_text()
                    _flush_thinking()
                    call_id = event.tool_call_id or ""
                    status = "error" if event.data.get("status") == "error" else "success"
                    duration_ms = event.data.get("duration_ms")
                    # Recover accumulated args from the core buffer (args stream in chunks)
                    args_str = core._tool_args_buffer.get(call_id, "") if call_id else ""
                    args_patch: dict = {}
                    if args_str:
                        try:
                            args_patch = json.loads(args_str)
                        except Exception:
                            args_patch = {}
                    stream_logger.info(
                        "<<< tool:result id=%s name=%s status=%s duration_ms=%s result=%r",
                        call_id, event.tool_name, status, duration_ms, event.content,
                    )
                    yield _sse({
                        "type": "tool:result",
                        "id": call_id,
                        "name": event.tool_name,
                        "status": status,
                        "result": event.content,
                        "duration_ms": duration_ms,
                        "args": args_patch,
                    })

                elif isinstance(event, TodoUpdateEvent):
                    _flush_text()
                    _flush_thinking()
                    tasks = [
                        {
                            "id": t.get("id") or str(
                                uuid.uuid5(uuid.NAMESPACE_OID, t.get("content", t.get("label", str(i))))
                            ),
                            "label": t.get("content", t.get("label", "")),
                            "status": t.get("status", "pending"),
                        }
                        for i, t in enumerate(event.todos or [])
                    ]
                    stream_logger.info("<<< tasks:update count=%d", len(tasks))
                    yield _sse({"type": "tasks:update", "tasks": tasks})

                elif isinstance(event, QueryEndEvent):
                    _flush_text()
                    _flush_thinking()
                    stream_logger.info("<<< stream:done")
                    yield _sse({"type": "stream:done"})
                    await asyncio.to_thread(_close_workbooks_sync)

                elif isinstance(event, ErrorEvent):
                    _flush_text()
                    _flush_thinking()
                    stream_logger.error("<<< stream:done error=%r", event.error_message)
                    yield _sse({"type": "stream:done", "error": event.error_message})
                    await asyncio.to_thread(_close_workbooks_sync)

        except Exception as exc:
            _flush_text()
            _flush_thinking()
            logger.exception("stream_endpoint unhandled error")
            yield _sse({"type": "stream:done", "error": str(exc)})
            await asyncio.to_thread(_close_workbooks_sync)
        finally:
            project_svc.mark_stream_end()

    return StreamingResponse(
        generate(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )
