"""WebSocket file watcher: /ws/files/watch"""

import asyncio
import logging
import os
import secrets
from pathlib import Path

from fastapi import APIRouter, WebSocket, WebSocketDisconnect

logger = logging.getLogger("app.fastapi")

router = APIRouter()

_HEARTBEAT_INTERVAL = 20.0  # seconds between server pings
_RECEIVE_TIMEOUT = 30.0     # seconds to wait for client message before closing


@router.websocket("/ws/files/watch")
async def ws_files_watch(websocket: WebSocket, path: str = "", token: str = ""):
    # WebSocket can't send custom headers — auth via query param
    if os.environ.get("EXCEL_AGENT_DEV") != "1":
        app_token: str = websocket.app.state.app_token
        if not secrets.compare_digest(token, app_token):
            await websocket.close(code=1008)
            return

    await websocket.accept()

    async def _watch_loop(watch_path: Path, stop: asyncio.Event):
        from watchfiles import Change, awatch

        change_map = {
            Change.added: "added",
            Change.modified: "modified",
            Change.deleted: "deleted",
        }
        try:
            async for changes in awatch(str(watch_path), stop_event=stop, recursive=False):
                affected_dirs: set[str] = set()
                events = []
                for change_type, changed_path in changes:
                    parent = str(Path(changed_path).parent)
                    affected_dirs.add(parent)
                    events.append({
                        "type": change_map.get(change_type, "modified"),
                        "path": changed_path,
                        "dir": parent,
                    })
                await websocket.send_json({
                    "event": "changes",
                    "changes": events,
                    "affected_dirs": list(affected_dirs),
                })
        except Exception as exc:
            logger.debug("ws_files_watch loop ended: %s", exc)

    async def _heartbeat_loop(stop: asyncio.Event):
        while not stop.is_set():
            await asyncio.sleep(_HEARTBEAT_INTERVAL)
            if stop.is_set():
                break
            try:
                await websocket.send_json({"type": "ping"})
            except Exception:
                break

    watch_path = Path(path).expanduser() if path else Path.home()
    stop_event = asyncio.Event()
    watch_task = asyncio.create_task(_watch_loop(watch_path, stop_event))
    heartbeat_task = asyncio.create_task(_heartbeat_loop(stop_event))

    try:
        while True:
            try:
                msg = await asyncio.wait_for(
                    websocket.receive_json(), timeout=_RECEIVE_TIMEOUT
                )
            except asyncio.TimeoutError:
                logger.debug("ws_files_watch: client timed out, closing")
                break

            if msg.get("type") == "pong":
                continue

            if msg.get("action") == "watch" and msg.get("path"):
                stop_event.set()
                watch_task.cancel()
                try:
                    await watch_task
                except (asyncio.CancelledError, Exception):
                    pass
                stop_event = asyncio.Event()
                watch_path = Path(msg["path"]).expanduser()
                watch_task = asyncio.create_task(_watch_loop(watch_path, stop_event))
                await websocket.send_json({"event": "watching", "path": str(watch_path)})
    except WebSocketDisconnect:
        pass
    except Exception as exc:
        logger.debug("ws_files_watch client error: %s", exc)
    finally:
        stop_event.set()
        heartbeat_task.cancel()
        watch_task.cancel()
        for task in (heartbeat_task, watch_task):
            try:
                await task
            except (asyncio.CancelledError, Exception):
                pass
