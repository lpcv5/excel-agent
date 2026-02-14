"""ExcelInstanceManager — singleton managing Excel.Application and workbook registry."""

import atexit
import logging
import os
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, ClassVar, Optional

import pythoncom
import win32com.client
from pywintypes import com_error as COMError

from .constants import xlCalculationManual
from .errors import (
    ExcelInstanceError,
    SheetNotFoundError,
    WorkbookNotFoundError,
    WorkbookReadOnlyError,
)
from .utils import com_retry, normalize_path, pump_messages

logger = logging.getLogger("app.excel")


@dataclass
class WorkbookEntry:
    file_path: str
    workbook: Any  # CDispatch
    app: Any  # Excel.Application that owns this workbook
    was_already_open: bool
    is_editing: bool = False
    opened_at: datetime = field(default_factory=datetime.now)
    last_operation_at: datetime = field(default_factory=datetime.now)


class ExcelInstanceManager:
    _instance: ClassVar[Optional["ExcelInstanceManager"]] = None
    _VALIDATION_INTERVAL: ClassVar[float] = 5.0  # seconds

    # Type annotations only — values are set in __new__
    _app: Optional[Any]
    _app_is_attached: bool
    _registry: dict[str, WorkbookEntry]
    _last_validation_time: float

    def __new__(cls) -> "ExcelInstanceManager":
        if cls._instance is None:
            inst = super().__new__(cls)
            inst._registry = {}
            inst._app = None
            inst._app_is_attached = False
            inst._last_validation_time = 0.0
            cls._instance = inst
            atexit.register(inst.cleanup)
        return cls._instance

    # ── Application lifecycle ──────────────────────────────────────────

    def _ensure_app(self) -> Any:
        """Get or create Excel.Application (steps 2+3 of the three-step strategy).

        Step 2: attach to an already-running Excel instance.
        Step 3: launch a new Excel instance if none is running.
        """
        pythoncom.CoInitialize()  # safe to call repeatedly (returns S_FALSE if already done)

        if self._app is not None:
            try:
                _ = self._app.Version
                return self._app
            except (COMError, AttributeError):
                logger.warning("Excel.Application reference is stale, reconnecting")
                self._app = None
                self._app_is_attached = False

        # Step 2: attach to existing instance
        try:
            self._app = win32com.client.GetActiveObject("Excel.Application")
            self._app_is_attached = True
            logger.info(
                "Attached to existing Excel instance (version %s)", self._app.Version
            )
            return self._app
        except COMError:
            pass

        # Step 3: launch new instance
        try:
            self._app = win32com.client.DispatchEx("Excel.Application")
            self._app.Visible = False
            self._app.DisplayAlerts = False
            self._app_is_attached = False
            logger.info(
                "Created new Excel instance (version %s)", self._app.Version
            )
            return self._app
        except COMError as e:
            raise ExcelInstanceError(f"无法启动Excel: {e}", e)

    @property
    def app(self) -> Any:
        """Public access to the Excel.Application COM object."""
        return self._ensure_app()

    # ── ROT helpers ────────────────────────────────────────────────────

    @staticmethod
    def _is_file_in_rot(norm_path: str) -> bool:
        """Check if a file path is registered in the Running Object Table.

        This prevents GetObject from silently *opening* a file that is not
        already open, which would cause was_already_open to be set incorrectly.
        """
        pythoncom.CoInitialize()
        try:
            ctx = pythoncom.CreateBindCtx(0)
            rot = pythoncom.GetRunningObjectTable()
            for moniker in rot.EnumRunning():
                try:
                    display = moniker.GetDisplayName(ctx, None)
                    if normalize_path(display) == norm_path:
                        return True
                except pythoncom.com_error:
                    continue
        except Exception:
            pass
        return False

    # ── Registry health ────────────────────────────────────────────────

    def _validate_registry(self, *, force: bool = False) -> None:
        """Remove stale entries whose COM references are dead.

        Throttled to at most once per _VALIDATION_INTERVAL seconds unless
        *force* is True — avoids redundant cross-process COM calls on
        hot paths like get_workbook().
        """
        now = time.monotonic()
        if not force and (now - self._last_validation_time) < self._VALIDATION_INTERVAL:
            return
        self._last_validation_time = now

        stale_keys: list[str] = []
        for key, entry in self._registry.items():
            try:
                _ = entry.workbook.Name
            except (COMError, AttributeError):
                stale_keys.append(key)
        for key in stale_keys:
            logger.warning("Removing stale workbook entry: %s", key)
            del self._registry[key]

    # ── Workbook operations ────────────────────────────────────────────

    @com_retry(max_retries=2, delay=0.5)
    def open_workbook(self, file_path: str) -> WorkbookEntry:
        """Open a workbook using a three-step strategy:

        1. File already open in Excel → attach via GetObject (COM moniker)
        2. Excel running but file not open → open in existing instance
        3. Excel not running → launch new instance and open file
        """
        self._validate_registry()
        abs_path = os.path.abspath(file_path)
        norm = normalize_path(abs_path)

        # Already in registry?
        if norm in self._registry:
            entry = self._registry[norm]
            try:
                _ = entry.workbook.Name
                logger.debug("Workbook already in registry: %s", norm)
                return entry
            except (COMError, AttributeError):
                del self._registry[norm]

        if not os.path.exists(abs_path):
            raise WorkbookNotFoundError(f"文件不存在: {abs_path}")

        # Step 1: file already open in Excel → GetObject returns it directly.
        # We MUST check the ROT first: GetObject on some Office versions will
        # silently open the file if it is not already open, which would make
        # was_already_open=True incorrect and cause the file to never be closed.
        if self._is_file_in_rot(norm):
            try:
                wb = win32com.client.GetObject(abs_path)
                wb_app = wb.Application
                logger.info(
                    "Attached to user-opened workbook via GetObject: %s", norm
                )
                entry = WorkbookEntry(
                    file_path=norm,
                    workbook=wb,
                    app=wb_app,
                    was_already_open=True,
                )
                self._registry[norm] = entry
                return entry
            except COMError:
                pass

        # Step 2 & 3: use existing or new Excel instance to open the file
        app = self._ensure_app()
        try:
            wb = app.Workbooks.Open(abs_path, UpdateLinks=False, ReadOnly=False)
            pump_messages()
        except COMError as e:
            raise WorkbookNotFoundError(f"无法打开文件 '{abs_path}': {e}", e)

        if wb.ReadOnly:
            wb.Close(SaveChanges=False)
            raise WorkbookReadOnlyError(
                f"文件已被其他程序以独占方式打开，无法写入: {abs_path}"
            )

        entry = WorkbookEntry(
            file_path=norm, workbook=wb, app=app, was_already_open=False
        )
        self._registry[norm] = entry
        logger.info("Opened workbook: %s", norm)
        return entry

    @com_retry(max_retries=2, delay=0.5)
    def create_workbook(self, file_path: str) -> WorkbookEntry:
        """Create a new workbook and save it to *file_path*."""
        abs_path = os.path.abspath(file_path)
        norm = normalize_path(abs_path)

        # Guard: already tracked
        self._validate_registry()
        if norm in self._registry:
            logger.debug("Workbook already in registry: %s", norm)
            return self._registry[norm]

        # Guard: file already exists on disk
        if os.path.exists(abs_path):
            raise ExcelInstanceError(
                f"文件已存在: {abs_path}，请使用 open_workbook 打开"
            )

        app = self._ensure_app()
        os.makedirs(os.path.dirname(abs_path) or ".", exist_ok=True)

        wb = app.Workbooks.Add()
        pump_messages()
        wb.SaveAs(abs_path)
        pump_messages()

        entry = WorkbookEntry(
            file_path=norm, workbook=wb, app=app, was_already_open=False
        )
        self._registry[norm] = entry
        logger.info("Created workbook: %s", norm)
        return entry

    def get_workbook(self, file_path: str) -> WorkbookEntry:
        """Get an already-open workbook entry, or open it."""
        norm = normalize_path(os.path.abspath(file_path))
        self._validate_registry()  # throttled — cheap on hot path
        if norm in self._registry:
            return self._registry[norm]
        return self.open_workbook(file_path)

    @com_retry(max_retries=2, delay=0.5)
    def save_workbook(self, file_path: str) -> None:
        """Save a workbook."""
        entry = self.get_workbook(file_path)
        entry.workbook.Save()
        entry.last_operation_at = datetime.now()
        pump_messages()
        logger.info("Saved workbook: %s", entry.file_path)

    def close_workbook(self, file_path: str, save: bool = True) -> None:
        """Close a workbook. User-opened files are saved but not closed."""
        norm = normalize_path(os.path.abspath(file_path))
        self._validate_registry()
        if norm not in self._registry:
            logger.debug("Workbook not in registry, nothing to close: %s", norm)
            return

        entry = self._registry[norm]
        if entry.was_already_open:
            if save:
                try:
                    entry.workbook.Save()
                    pump_messages()
                except COMError as e:
                    logger.warning(
                        "Error saving user-opened workbook %s: %s", norm, e
                    )
            logger.info("Saved (not closing) user-opened workbook: %s", norm)
            return

        try:
            entry.workbook.Close(SaveChanges=save)
            pump_messages()
        except Exception as e:  # best-effort close
            logger.warning("Error closing workbook %s: %s", norm, e)
        del self._registry[norm]
        logger.info("Closed workbook: %s", norm)

    def get_sheet(self, file_path: str, sheet: str | int | None = None) -> Any:
        """Get a worksheet COM object. Defaults to active sheet."""
        entry = self.get_workbook(file_path)
        wb = entry.workbook
        if sheet is None:
            return wb.ActiveSheet
        try:
            return wb.Sheets(sheet)
        except COMError as e:
            available = [
                wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)
            ]
            raise SheetNotFoundError(
                f"工作表 '{sheet}' 不存在，可用工作表: {available}", e
            )

    # ── Workbook status helpers ────────────────────────────────────────

    def is_workbook_saved(self, file_path: str) -> bool:
        """Check whether a workbook has unsaved changes (queries COM directly)."""
        entry = self.get_workbook(file_path)
        try:
            return bool(entry.workbook.Saved)
        except (COMError, AttributeError):
            return True  # can't determine → assume saved

    def mark_dirty(self, file_path: str) -> None:
        """Update last_operation_at — call after tool writes to cells."""
        norm = normalize_path(os.path.abspath(file_path))
        entry = self._registry.get(norm)
        if entry:
            entry.last_operation_at = datetime.now()

    # ── Batch operations ───────────────────────────────────────────────

    class _BatchContext:
        """Context manager that optimises Excel settings for bulk writes."""

        def __init__(self, mgr: "ExcelInstanceManager", file_path: str):
            self._mgr = mgr
            self._file_path = file_path
            self._app_ref: Any = None
            self._prev_calc = None
            self._prev_events = None
            self._prev_screen = None
            self._prev_alerts = None

        def __enter__(self) -> "ExcelInstanceManager._BatchContext":
            entry = self._mgr.get_workbook(self._file_path)
            # Use the app that *owns* this workbook — not _ensure_app(),
            # which may return a different Application instance.
            self._app_ref = entry.app
            app = self._app_ref

            self._prev_calc = app.Calculation
            self._prev_events = app.EnableEvents
            self._prev_screen = app.ScreenUpdating
            self._prev_alerts = app.DisplayAlerts

            app.Calculation = xlCalculationManual
            app.EnableEvents = False
            app.DisplayAlerts = False
            # Keep ScreenUpdating on for user-opened files so the user can
            # see changes happening in real time.
            if not entry.was_already_open:
                app.ScreenUpdating = False

            entry.is_editing = True
            return self

        def __exit__(self, *_: object) -> None:
            app = self._app_ref
            if app is None:
                return

            # Check liveness — do NOT call _ensure_app() which would spin
            # up a new Excel instance if the original one crashed.
            try:
                _ = app.Version
            except (COMError, AttributeError):
                logger.warning(
                    "Excel exited during batch operation, skipping settings restore"
                )
                return

            try:
                if self._prev_calc is not None:
                    app.Calculation = self._prev_calc
                if self._prev_events is not None:
                    app.EnableEvents = self._prev_events
                if self._prev_screen is not None:
                    app.ScreenUpdating = self._prev_screen
                if self._prev_alerts is not None:
                    app.DisplayAlerts = self._prev_alerts
            except (COMError, AttributeError) as e:
                logger.warning("Error restoring Excel settings: %s", e)

            norm = normalize_path(os.path.abspath(self._file_path))
            entry = self._mgr._registry.get(norm)
            if entry:
                entry.is_editing = False
                entry.last_operation_at = datetime.now()

    def batch_operation(self, file_path: str) -> _BatchContext:
        """Return a context manager that optimises Excel for batch ops."""
        return self._BatchContext(self, file_path)

    # ── Status & cleanup ───────────────────────────────────────────────

    def list_workbooks(self) -> list[dict[str, Any]]:
        """Return status of all tracked workbooks."""
        self._validate_registry(force=True)
        result = []
        for _key, entry in self._registry.items():
            try:
                saved = bool(entry.workbook.Saved)
            except (COMError, AttributeError):
                saved = True
            result.append(
                {
                    "file_path": entry.file_path,
                    "was_already_open": entry.was_already_open,
                    "is_editing": entry.is_editing,
                    "is_saved": saved,
                }
            )
        return result

    def _close_registry_entries(self, *, quit_app_if_empty: bool = False) -> None:
        """Shared logic for close_agent_workbooks and cleanup."""
        self._validate_registry(force=True)
        for key in list(self._registry.keys()):
            entry = self._registry[key]
            try:
                if entry.was_already_open:
                    if not entry.workbook.Saved:
                        entry.workbook.Save()
                        pump_messages()
                    logger.info("Saved user-opened workbook: %s", key)
                else:
                    entry.workbook.Close(SaveChanges=True)
                    pump_messages()
                    del self._registry[key]
                    logger.info("Closed agent-opened workbook: %s", key)
            except Exception as e:  # best-effort — COM may already be dead
                logger.warning("Error closing %s: %s", key, e)
                # User-opened: keep entry, will be cleaned up by next validation.
                # Agent-opened: COM reference is likely dead, remove now.
                if not entry.was_already_open:
                    self._registry.pop(key, None)

        if quit_app_if_empty and self._app and not self._app_is_attached:
            try:
                if self._app.Workbooks.Count == 0:
                    self._app.Quit()
                    logger.info("Quit Excel application")
            except Exception as e:
                logger.warning("Error quitting Excel: %s", e)
            self._app = None

    def close_agent_workbooks(self) -> None:
        """Close agent-opened workbooks and save user-opened ones."""
        self._close_registry_entries(quit_app_if_empty=False)

    def cleanup(self) -> None:
        """Close all workbooks and quit agent-owned Excel instance."""
        self._close_registry_entries(quit_app_if_empty=True)