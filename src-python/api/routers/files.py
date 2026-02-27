"""File tree and dialog endpoints."""

import asyncio
import logging
import os
import shutil
from pathlib import Path

from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel

from api.deps import verify_token

logger = logging.getLogger("app.fastapi")

router = APIRouter(dependencies=[Depends(verify_token)])


class FileRenameBody(BaseModel):
    path: str
    new_name: str


class FileDeleteBody(BaseModel):
    path: str


class FileMoveBody(BaseModel):
    src: str
    dst_dir: str


class DialogOpenBody(BaseModel):
    multiple: bool = False
    filters: list[str] = []


class DialogSaveBody(BaseModel):
    default_path: str = ""
    filters: list[str] = []


def _sync_list_files(target: str) -> dict:
    with os.scandir(target) as it:
        raw = list(it)
    raw.sort(key=lambda e: (not e.is_dir(follow_symlinks=False), e.name.lower()))
    entries = []
    for entry in raw:
        is_dir = entry.is_dir(follow_symlinks=False)
        try:
            size = entry.stat(follow_symlinks=False).st_size if not is_dir else None
        except OSError:
            size = None
        entries.append({
            "name": entry.name,
            "path": entry.path,
            "type": "dir" if is_dir else "file",
            "size": size,
        })
    return {"path": target, "entries": entries}


@router.get("/api/files")
async def list_files(path: str = ""):
    target = Path(path).expanduser() if path else Path.home()
    if not target.is_dir():
        raise HTTPException(status_code=400, detail="Not a directory")
    try:
        return await asyncio.to_thread(_sync_list_files, str(target))
    except PermissionError:
        raise HTTPException(status_code=403, detail="Permission denied")


@router.post("/api/files/rename")
async def rename_file(body: FileRenameBody):
    src = Path(body.path)
    dst = src.parent / body.new_name
    if not src.exists():
        raise HTTPException(status_code=404, detail="Not found")
    if dst.exists():
        raise HTTPException(status_code=409, detail="Name already exists")
    src.rename(dst)
    return {"path": str(dst)}


@router.delete("/api/files")
async def delete_file(body: FileDeleteBody):
    target = Path(body.path)
    if not target.exists():
        raise HTTPException(status_code=404, detail="Not found")
    if target.is_dir():
        shutil.rmtree(target)
    else:
        target.unlink()
    return {"ok": True}


@router.post("/api/files/move")
async def move_file(body: FileMoveBody):
    src = Path(body.src)
    dst_dir = Path(body.dst_dir)
    if not src.exists():
        raise HTTPException(status_code=404, detail="Source not found")
    if not dst_dir.is_dir():
        raise HTTPException(status_code=400, detail="Destination is not a directory")
    dst = dst_dir / src.name
    if dst.exists():
        raise HTTPException(status_code=409, detail="Name already exists at destination")
    shutil.move(str(src), str(dst))
    return {"path": str(dst)}


def _get_webview_window(request: Request):
    window = getattr(request.app.state, "webview_window", None)
    if window is None:
        raise HTTPException(status_code=503, detail="File dialog unavailable (no desktop window)")
    return window


@router.post("/api/dialog/open")
async def dialog_open(body: DialogOpenBody, request: Request):
    window = _get_webview_window(request)
    try:
        import webview  # type: ignore[import]

        result = window.create_file_dialog(
            webview.FileDialog.OPEN,
            allow_multiple=body.multiple,
            file_types=tuple(body.filters),
        )
        return {"paths": list(result) if result else None}
    except HTTPException:
        raise
    except Exception:
        logger.exception("dialog_open failed")
        raise HTTPException(status_code=503, detail="File dialog unavailable (no desktop window)")


@router.post("/api/dialog/save")
async def dialog_save(body: DialogSaveBody, request: Request):
    window = _get_webview_window(request)
    try:
        import webview  # type: ignore[import]

        result = window.create_file_dialog(
            webview.FileDialog.SAVE,
            directory=body.default_path,
            file_types=tuple(body.filters),
        )
        return {"path": result[0] if result else None}
    except HTTPException:
        raise
    except Exception:
        logger.exception("dialog_save failed")
        raise HTTPException(status_code=503, detail="File dialog unavailable (no desktop window)")


@router.post("/api/dialog/folder")
async def dialog_folder(request: Request):
    window = _get_webview_window(request)
    try:
        import webview  # type: ignore[import]

        result = window.create_file_dialog(webview.FileDialog.FOLDER)
        return {"path": result[0] if result else None}
    except HTTPException:
        raise
    except Exception:
        logger.exception("dialog_folder failed")
        raise HTTPException(status_code=503, detail="Folder dialog unavailable")
