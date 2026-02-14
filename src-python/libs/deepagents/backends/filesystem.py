"""`FilesystemBackend`: Read and write files directly from the filesystem.

Optimized for Windows platform — uses os.scandir() for cached stat data,
plain open() instead of O_NOFOLLOW ceremony, and CREATE_NO_WINDOW for subprocesses.
"""

import json
import os
import re
import subprocess
from datetime import UTC, datetime
from pathlib import Path

import wcmatch.glob as wcglob

from libs.deepagents.backends.protocol import (
    BackendProtocol,
    EditResult,
    FileDownloadResponse,
    FileInfo,
    FileUploadResponse,
    GrepMatch,
    WriteResult,
)
from libs.deepagents.backends.utils import (
    check_empty_content,
    format_content_with_line_numbers,
    perform_string_replacement,
)

# Windows: prevent console window flash when spawning ripgrep from GUI
_STARTUPINFO = subprocess.STARTUPINFO()
_STARTUPINFO.dwFlags |= subprocess.STARTF_USESHOWWINDOW
_STARTUPINFO.wShowWindow = subprocess.SW_HIDE
_CREATION_FLAGS = subprocess.CREATE_NO_WINDOW


class FilesystemBackend(BackendProtocol):
    """Backend that reads and writes files directly from the filesystem.

    Files are accessed using their actual filesystem paths. Absolute paths are
    used as-is; relative paths are resolved relative to the current working directory.
    Content is read/written as plain text, and metadata (timestamps) are derived
    from filesystem stats.
    """

    def __init__(
        self,
        root_dir: str | Path | None = None,
        max_file_size_mb: int = 10,
    ) -> None:
        """Initialize filesystem backend.

        Args:
            root_dir: Optional root directory for file operations.
                If not provided, defaults to the current working directory.
                Only affects relative path resolution — absolute paths are used as-is.

            max_file_size_mb: Maximum file size in megabytes for operations like
                grep's Python fallback search.
                Files exceeding this limit are skipped during search. Defaults to 10 MB.
        """
        self.cwd = Path(root_dir).resolve() if root_dir else Path.cwd()
        self.max_file_size_bytes = max_file_size_mb * 1024 * 1024

    def _resolve_path(self, key: str) -> Path:
        """Resolve a file path.

        Absolute paths are used as-is; relative paths resolve under cwd.

        Args:
            key: File path (absolute or relative).

        Returns:
            Resolved absolute `Path` object.
        """
        if os.path.isabs(key):
            return Path(key)
        return Path(os.path.normpath(os.path.join(str(self.cwd), key)))

    def ls_info(self, path: str) -> list[FileInfo]:
        """List files and directories in the specified directory (non-recursive).

        Args:
            path: Absolute directory path to list files from.

        Returns:
            List of `FileInfo`-like dicts for files and directories directly in the
                directory. Directories have a trailing `/` in their path and
                `is_dir=True`.
        """
        dir_path = self._resolve_path(path)
        if not dir_path.exists() or not dir_path.is_dir():
            return []

        results: list[FileInfo] = []

        # os.scandir() uses FindFirstFile/FindNextFile on Windows,
        # which returns size + timestamps in the directory entry itself.
        # entry.stat(follow_symlinks=False) is cached — no extra syscall.
        try:
            with os.scandir(dir_path) as entries:
                for entry in entries:
                    try:
                        is_file = entry.is_file(follow_symlinks=False)
                        is_dir = entry.is_dir(follow_symlinks=False)
                    except OSError:
                        continue

                    abs_path = entry.path  # already a string

                    if is_file:
                        try:
                            st = entry.stat(follow_symlinks=False)
                            results.append(
                                {
                                    "path": abs_path,
                                    "is_dir": False,
                                    "size": st.st_size,
                                    "modified_at": datetime.fromtimestamp(
                                        st.st_mtime, tz=UTC
                                    ).isoformat(),
                                }
                            )
                        except OSError:
                            results.append({"path": abs_path, "is_dir": False})
                    elif is_dir:
                        try:
                            st = entry.stat(follow_symlinks=False)
                            results.append(
                                {
                                    "path": abs_path + "\\",
                                    "is_dir": True,
                                    "size": 0,
                                    "modified_at": datetime.fromtimestamp(
                                        st.st_mtime, tz=UTC
                                    ).isoformat(),
                                }
                            )
                        except OSError:
                            results.append({"path": abs_path + "\\", "is_dir": True})
        except (OSError, PermissionError):
            pass

        # Keep deterministic order by path
        results.sort(key=lambda x: x.get("path", ""))
        return results

    def read(
        self,
        file_path: str,
        offset: int = 0,
        limit: int = 2000,
    ) -> str:
        """Read file content with line numbers.

        Args:
            file_path: Absolute or relative file path.
            offset: Line offset to start reading from (0-indexed).
            limit: Maximum number of lines to read.

        Returns:
            Formatted file content with line numbers, or error message.
        """
        resolved_path = self._resolve_path(file_path)

        if not resolved_path.exists() or not resolved_path.is_file():
            return f"Error: File '{file_path}' not found"

        try:
            # Paginated path: skip lines efficiently without loading entire file
            if offset > 0 or limit < 2000:
                with open(resolved_path, encoding="utf-8") as f:
                    # Skip to offset
                    for _ in range(offset):
                        if not f.readline():
                            return f"Error: Line offset {offset} exceeds file length"
                    selected = []
                    for _ in range(limit):
                        line = f.readline()
                        if not line:
                            break
                        selected.append(line.rstrip("\r\n"))
                if not selected and offset == 0:
                    return check_empty_content("") or ""
                return format_content_with_line_numbers(selected, start_line=offset + 1)

            # Default path: read entire file
            content = resolved_path.read_text(encoding="utf-8")

            empty_msg = check_empty_content(content)
            if empty_msg:
                return empty_msg

            lines = content.splitlines()
            end_idx = min(limit, len(lines))
            return format_content_with_line_numbers(lines[:end_idx], start_line=1)
        except (OSError, UnicodeDecodeError) as e:
            return f"Error reading file '{file_path}': {e}"

    def write(
        self,
        file_path: str,
        content: str,
    ) -> WriteResult:
        """Create a new file with content.

        Args:
            file_path: Path where the new file will be created.
            content: Text content to write to the file.

        Returns:
            `WriteResult` with path on success, or error message if the file
                already exists or write fails. External storage sets `files_update=None`.
        """
        resolved_path = self._resolve_path(file_path)

        if resolved_path.exists():
            return WriteResult(
                error=f"Cannot write to {file_path} because it already exists. Read and then make an edit, or write to a new path."
            )

        try:
            resolved_path.parent.mkdir(parents=True, exist_ok=True)
            resolved_path.write_text(content, encoding="utf-8")
            return WriteResult(path=file_path, files_update=None)
        except (OSError, UnicodeEncodeError) as e:
            return WriteResult(error=f"Error writing file '{file_path}': {e}")

    def edit(
        self,
        file_path: str,
        old_string: str,
        new_string: str,
        replace_all: bool = False,
    ) -> EditResult:
        """Edit a file by replacing string occurrences.

        Args:
            file_path: Path to the file to edit.
            old_string: The text to search for and replace.
            new_string: The replacement text.
            replace_all: If `True`, replace all occurrences. If `False` (default),
                replace only if exactly one occurrence exists.

        Returns:
            `EditResult` with path and occurrence count on success, or error
                message if file not found or replacement fails. External storage sets
                `files_update=None`.
        """
        resolved_path = self._resolve_path(file_path)

        if not resolved_path.exists() or not resolved_path.is_file():
            return EditResult(error=f"Error: File '{file_path}' not found")

        try:
            content = resolved_path.read_text(encoding="utf-8")

            result = perform_string_replacement(
                content, old_string, new_string, replace_all
            )

            if isinstance(result, str):
                return EditResult(error=result)

            new_content, occurrences = result
            resolved_path.write_text(new_content, encoding="utf-8")

            return EditResult(
                path=file_path, files_update=None, occurrences=int(occurrences)
            )
        except (OSError, UnicodeDecodeError, UnicodeEncodeError) as e:
            return EditResult(error=f"Error editing file '{file_path}': {e}")

    def grep_raw(
        self,
        pattern: str,
        path: str | None = None,
        glob: str | None = None,
    ) -> list[GrepMatch] | str:
        """Search for a literal text pattern in files.

        Uses ripgrep if available, falling back to Python search.

        Args:
            pattern: Literal string to search for (NOT regex).
            path: Directory or file path to search in. Defaults to current directory.
            glob: Optional glob pattern to filter which files to search.

        Returns:
            List of GrepMatch dicts containing path, line number, and matched text.
        """
        # Resolve base path
        try:
            base_full = self._resolve_path(path or ".")
        except ValueError:
            return []

        if not base_full.exists():
            return []

        # Try ripgrep first (with -F flag for literal search)
        results = self._ripgrep_search(pattern, base_full, glob)
        if results is None:
            # Python fallback needs escaped pattern for literal search
            results = self._python_search(re.escape(pattern), base_full, glob)

        matches: list[GrepMatch] = []
        for fpath, items in results.items():
            for line_num, line_text in items:
                matches.append(
                    {"path": fpath, "line": int(line_num), "text": line_text}
                )
        return matches

    def _ripgrep_search(
        self, pattern: str, base_full: Path, include_glob: str | None
    ) -> dict[str, list[tuple[int, str]]] | None:
        """Search using ripgrep with fixed-string (literal) mode.

        Args:
            pattern: Literal string to search for (unescaped).
            base_full: Resolved base path to search in.
            include_glob: Optional glob pattern to filter files.

        Returns:
            Dict mapping file paths to list of `(line_number, line_text)` tuples.
                Returns `None` if ripgrep is unavailable or times out.
        """
        cmd = ["rg", "--json", "-F"]  # -F enables fixed-string (literal) mode
        if include_glob:
            cmd.extend(["--glob", include_glob])
        cmd.extend(["--", pattern, str(base_full)])

        try:
            proc = subprocess.run(  # noqa: S603
                cmd,
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
                startupinfo=_STARTUPINFO,
                creationflags=_CREATION_FLAGS,
            )
        except (subprocess.TimeoutExpired, FileNotFoundError):
            return None

        results: dict[str, list[tuple[int, str]]] = {}
        for line in proc.stdout.splitlines():
            try:
                data = json.loads(line)
            except json.JSONDecodeError:
                continue
            if data.get("type") != "match":
                continue
            pdata = data.get("data", {})
            ftext = pdata.get("path", {}).get("text")
            if not ftext:
                continue
            p = Path(ftext)
            ln = pdata.get("line_number")
            lt = pdata.get("lines", {}).get("text", "").rstrip("\n")
            if ln is None:
                continue
            results.setdefault(str(p), []).append((int(ln), lt))

        return results

    def _python_search(
        self, pattern: str, base_full: Path, include_glob: str | None
    ) -> dict[str, list[tuple[int, str]]]:
        """Fallback search using Python when ripgrep is unavailable.

        Uses os.walk() for efficient directory traversal on Windows and reads
        files line-by-line to avoid loading entire files into memory.

        Args:
            pattern: Escaped regex pattern (from re.escape) for literal search.
            base_full: Resolved base path to search in.
            include_glob: Optional glob pattern to filter files by name.

        Returns:
            Dict mapping file paths to list of `(line_number, line_text)` tuples.
        """
        regex = re.compile(pattern)
        results: dict[str, list[tuple[int, str]]] = {}
        root = str(base_full if base_full.is_dir() else base_full.parent)

        for dirpath, _dirnames, filenames in os.walk(root):
            for filename in filenames:
                if include_glob and not wcglob.globmatch(
                    filename, include_glob, flags=wcglob.BRACE
                ):
                    continue
                filepath = os.path.join(dirpath, filename)
                try:
                    if os.path.getsize(filepath) > self.max_file_size_bytes:
                        continue
                except OSError:
                    continue
                try:
                    with open(filepath, encoding="utf-8") as f:
                        for line_num, line in enumerate(f, 1):
                            line = line.rstrip("\n\r")
                            if regex.search(line):
                                results.setdefault(filepath, []).append(
                                    (line_num, line)
                                )
                except (UnicodeDecodeError, PermissionError, OSError):
                    continue

        return results

    def glob_info(self, pattern: str, path: str = ".") -> list[FileInfo]:
        """Find files matching a glob pattern.

        Args:
            pattern: Glob pattern to match files against (e.g., `'*.py'`, `'**/*.txt'`).
            path: Base directory to search from. Defaults to current working directory.

        Returns:
            List of `FileInfo` dicts for matching files, sorted by path. Each dict
                contains `path`, `is_dir`, `size`, and `modified_at` fields.
        """
        search_path = self._resolve_path(path)
        if not search_path.exists() or not search_path.is_dir():
            return []

        results: list[FileInfo] = []
        try:
            for matched_path in search_path.rglob(pattern):
                try:
                    is_file = matched_path.is_file()
                except (PermissionError, OSError):
                    continue
                if not is_file:
                    continue
                abs_path = str(matched_path)
                try:
                    st = matched_path.stat()
                    results.append(
                        {
                            "path": abs_path,
                            "is_dir": False,
                            "size": int(st.st_size),
                            "modified_at": datetime.fromtimestamp(
                                st.st_mtime, tz=UTC
                            ).isoformat(),
                        }
                    )
                except OSError:
                    results.append({"path": abs_path, "is_dir": False})
        except (OSError, ValueError):
            pass

        results.sort(key=lambda x: x.get("path", ""))
        return results

    def upload_files(self, files: list[tuple[str, bytes]]) -> list[FileUploadResponse]:
        """Upload multiple files to the filesystem.

        Args:
            files: List of (path, content) tuples where content is bytes.

        Returns:
            List of FileUploadResponse objects, one per input file.
            Response order matches input order.
        """
        responses: list[FileUploadResponse] = []
        for path, content in files:
            try:
                resolved_path = self._resolve_path(path)
                resolved_path.parent.mkdir(parents=True, exist_ok=True)
                resolved_path.write_bytes(content)

                responses.append(FileUploadResponse(path=path, error=None))
            except FileNotFoundError:
                responses.append(FileUploadResponse(path=path, error="file_not_found"))
            except PermissionError:
                responses.append(
                    FileUploadResponse(path=path, error="permission_denied")
                )
            except (ValueError, OSError):
                responses.append(FileUploadResponse(path=path, error="invalid_path"))

        return responses

    def download_files(self, paths: list[str]) -> list[FileDownloadResponse]:
        """Download multiple files from the filesystem.

        Args:
            paths: List of file paths to download.

        Returns:
            List of FileDownloadResponse objects, one per input path.
        """
        responses: list[FileDownloadResponse] = []
        for path in paths:
            try:
                resolved_path = self._resolve_path(path)
                content = resolved_path.read_bytes()
                responses.append(
                    FileDownloadResponse(path=path, content=content, error=None)
                )
            except FileNotFoundError:
                responses.append(
                    FileDownloadResponse(
                        path=path, content=None, error="file_not_found"
                    )
                )
            except PermissionError:
                responses.append(
                    FileDownloadResponse(
                        path=path, content=None, error="permission_denied"
                    )
                )
            except IsADirectoryError:
                responses.append(
                    FileDownloadResponse(path=path, content=None, error="is_directory")
                )
            except ValueError:
                responses.append(
                    FileDownloadResponse(path=path, content=None, error="invalid_path")
                )
            # Let other errors propagate
        return responses
