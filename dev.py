"""Dev runner: interactive terminal UI for Vite + pywebview."""

from __future__ import annotations

import os
import queue
import re
import subprocess
import sys
import threading
import time
from enum import Enum

from rich.console import Console
from rich.panel import Panel
from rich.text import Text

console = Console(highlight=False)

_ANSI_RE = re.compile(r"\033\[[0-9;]*[mGKHF]")

# ── Process state ─────────────────────────────────────────────────────────────

class ProcessState(Enum):
    STOPPED  = "stopped"
    STARTING = "starting"
    READY    = "ready"
    ERROR    = "error"

# ── ManagedProcess ────────────────────────────────────────────────────────────

class ManagedProcess:
    def __init__(
        self,
        name: str,
        prefix: str,
        style: str,
        cmd: list[str],
        output_q: queue.Queue,
        cwd: str | None = None,
        env: dict | None = None,
        ready_marker: str | None = None,
        use_shell: bool = False,
    ) -> None:
        self.name = name
        self.prefix = prefix
        self.style = style
        self.cmd = cmd
        self.output_q = output_q
        self.cwd = cwd
        self.env = env
        self.ready_marker = ready_marker
        self.use_shell = use_shell
        self.state = ProcessState.STOPPED
        self._proc: subprocess.Popen | None = None

    def start(self) -> None:
        self.state = ProcessState.STARTING
        kwargs: dict = {}
        if sys.platform == "win32":
            kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP
        self._proc = subprocess.Popen(
            self.cmd,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            cwd=self.cwd,
            env=self.env,
            shell=self.use_shell,
            **kwargs,
        )
        t = threading.Thread(target=self._reader_thread, args=(self._proc,), daemon=True)
        t.start()

    def _reader_thread(self, proc: subprocess.Popen) -> None:
        try:
            assert proc.stdout is not None
            for raw in proc.stdout:
                line = raw.decode("utf-8", errors="replace").rstrip()
                if not line:
                    continue
                plain = _ANSI_RE.sub("", line)
                if self.ready_marker and self.ready_marker in plain:
                    self.state = ProcessState.READY
                self.output_q.put((self.prefix, self.style, plain))
        except Exception:
            pass
        if self.state == ProcessState.STARTING:
            self.state = ProcessState.ERROR

    def kill(self) -> None:
        if self._proc and self._proc.poll() is None:
            if sys.platform == "win32":
                subprocess.call(
                    ["taskkill", "/F", "/T", "/PID", str(self._proc.pid)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            else:
                self._proc.terminate()
                try:
                    self._proc.wait(timeout=5)
                except subprocess.TimeoutExpired:
                    self._proc.kill()
        self.state = ProcessState.STOPPED

    def restart(self) -> None:
        self.kill()
        self.start()

    def is_alive(self) -> bool:
        return self._proc is not None and self._proc.poll() is None


# ── DevUI ─────────────────────────────────────────────────────────────────────

class DevUI:
    def print_banner(self, dev_url: str, port: int) -> None:
        content = Text()
        content.append("Frontend:  ", style="dim")
        content.append(dev_url, style="cyan underline")
        content.append("\nBackend:   ", style="dim")
        content.append(f"http://localhost:{port}", style="cyan underline")
        content.append("\n\nPress ", style="dim")
        content.append("h", style="bold green")
        content.append(" for shortcuts", style="dim")
        console.print(Panel(content, title="[bold magenta]ExcelAgent Dev[/]", border_style="magenta"))

    def print_help(self) -> None:
        t = Text("\n")
        for key, desc in [("r", "restart Python backend"), ("v", "restart Vite"),
                           ("h", "show this help"), ("q", "quit")]:
            t.append(f"  {key}", style="bold green")
            t.append(f"  {desc}\n", style="dim")
        console.print(t)

    def log(self, prefix: str, style: str, line: str) -> None:
        t = Text()
        t.append(prefix, style=f"bold {style}")
        t.append(line)
        console.print(t)

    def system(self, msg: str) -> None:
        t = Text()
        t.append("[dev]   ", style="bold yellow")
        t.append(msg, style="dim")
        console.print(t)

    def error(self, msg: str) -> None:
        t = Text()
        t.append("[dev]   ", style="bold red")
        t.append(msg, style="red")
        console.print(t)


# ── KeyboardPoller ────────────────────────────────────────────────────────────

if sys.platform == "win32":
    import msvcrt as _msvcrt

    class KeyboardPoller:
        def kbhit(self) -> bool:
            return _msvcrt.kbhit()  # type: ignore[attr-defined]

        def getch(self) -> str:
            ch = _msvcrt.getch()  # type: ignore[attr-defined]
            if ch in (b"\x00", b"\xe0"):
                _msvcrt.getch()  # type: ignore[attr-defined]
                return ""
            try:
                return ch.decode("utf-8")
            except UnicodeDecodeError:
                return ""

        def restore(self) -> None:
            pass

else:
    import tty as _tty
    import termios as _termios
    import select as _select

    class KeyboardPoller:  # type: ignore[no-redef]
        def __init__(self) -> None:
            self._fd = sys.stdin.fileno()
            self._old = _termios.tcgetattr(self._fd)
            _tty.setraw(self._fd)

        def kbhit(self) -> bool:
            return bool(_select.select([sys.stdin], [], [], 0)[0])

        def getch(self) -> str:
            return sys.stdin.read(1)

        def restore(self) -> None:
            _termios.tcsetattr(self._fd, _termios.TCSADRAIN, self._old)


# ── Main loop ─────────────────────────────────────────────────────────────────

def main_loop(
    vite: ManagedProcess,
    app: ManagedProcess,
    dev_url: str,
    port: int,
    no_window: bool,
) -> None:
    q = vite.output_q
    kb = KeyboardPoller()
    ui = DevUI()

    ui.print_banner(dev_url, port)
    ui.system("Starting Vite...")
    vite.start()

    app_started = False

    try:
        while True:
            for _ in range(50):
                try:
                    prefix, style, line = q.get_nowait()
                    ui.log(prefix, style, line)
                except queue.Empty:
                    break

            if not app_started and vite.state == ProcessState.READY:
                app_started = True
                if no_window:
                    ui.system(f"Vite ready — open in browser: {dev_url}")
                else:
                    ui.system("Vite ready — starting Python backend...")
                    app.start()

            if not vite.is_alive() and vite.state == ProcessState.ERROR:
                ui.error("Vite failed to start.")
                break

            if app_started and not no_window and not app.is_alive():
                ui.system("App exited.")
                break

            if kb.kbhit():
                ch = kb.getch().lower()
                if ch in ("q", "\x03"):
                    ui.system("Quitting...")
                    break
                elif ch == "r":
                    if no_window:
                        ui.system("No app to restart in --no-window mode.")
                    else:
                        ui.system("Restarting Python backend...")
                        app.restart()
                        app_started = True
                elif ch == "v":
                    ui.system("Restarting Vite...")
                    if not no_window:
                        app.kill()
                    vite.restart()
                    app_started = False
                elif ch == "h":
                    ui.print_help()

            time.sleep(0.02)

    except KeyboardInterrupt:
        pass
    finally:
        kb.restore()
        app.kill()
        vite.kill()


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="ExcelAgent dev runner")
    parser.add_argument(
        "--pm",
        default=(
            os.environ.get("npm_execpath", "npm")
            .split("\\")[-1]
            .split("/")[-1]
            .replace(".js", "")
            .replace(".cmd", "")
            or "npm"
        ),
    )
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--dev-url", default="http://localhost:5173")
    parser.add_argument("--no-window", action="store_true")
    args = parser.parse_args()

    dev_url: str = args.dev_url
    pm: str = args.pm
    port: int = args.port

    root = os.path.dirname(os.path.abspath(__file__))
    src_python = os.path.join(root, "src-python")
    main_py = os.path.join(src_python, "main.py")

    q: queue.Queue = queue.Queue()

    vite = ManagedProcess(
        name="vite",
        prefix="[vite]  ",
        style="magenta",
        cmd=[pm, "run", "dev"],
        output_q=q,
        ready_marker="Local:",
        use_shell=(sys.platform == "win32"),
    )

    app = ManagedProcess(
        name="app",
        prefix="[app]   ",
        style="cyan",
        cmd=[sys.executable, main_py, "--port", str(port), "--dev-frontend", dev_url],
        output_q=q,
        cwd=src_python,
        env={**os.environ, "EXCEL_AGENT_DEV": "1"},
        ready_marker="API server ready.",
    )

    main_loop(vite, app, dev_url, port, args.no_window)


if __name__ == "__main__":
    main()



