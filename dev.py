"""Dev runner: starts Vite + pywebview (via main.py) in one command."""

import os
import subprocess
import sys
import time
import urllib.request


def wait_for_url(url: str, timeout: int = 30) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(url, timeout=2)
            return True
        except Exception:
            time.sleep(0.3)
    return False


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--pm",
        default=os.environ.get("npm_execpath", "npm").split("\\")[-1].split("/")[-1].replace(".js", "").replace(".cmd", "") or "npm",
    )
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--dev-url", default="http://localhost:5173")
    parser.add_argument("--no-window", action="store_true", help="Skip pywebview, just print the URL")
    args = parser.parse_args()

    dev_url = args.dev_url
    pm = args.pm
    port = args.port

    # Dev mode: no token needed
    env = {**os.environ, "EXCEL_AGENT_DEV": "1"}

    root = os.path.dirname(os.path.abspath(__file__))
    src_python = os.path.join(root, "src-python")

    kwargs = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP

    # 1. Start Vite
    vite = subprocess.Popen(
        [pm, "run", "dev"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        shell=(sys.platform == "win32"),
    )

    def _kill(proc: subprocess.Popen) -> None:
        if proc.poll() is not None:
            return
        if sys.platform == "win32":
            subprocess.call(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        else:
            proc.terminate()
            try:
                proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                proc.kill()

    app_proc: subprocess.Popen | None = None

    def _cleanup():
        _kill(vite)
        if app_proc is not None:
            _kill(app_proc)

    try:
        # 2. Wait for Vite
        print(f"Waiting for Vite on {dev_url}...")
        if not wait_for_url(dev_url):
            print("Vite failed to start.", file=sys.stderr)
            _cleanup()
            sys.exit(1)
        print("Vite is ready.")

        if args.no_window:
            print(f"\nOpen in browser: {dev_url}\n")
            vite.wait()
        else:
            # 3. Launch main.py --dev-frontend; it starts uvicorn + pywebview on its main thread
            main_py = os.path.join(src_python, "main.py")
            app_proc = subprocess.Popen(
                [sys.executable, main_py, "--port", str(port), "--dev-frontend", dev_url],
                cwd=src_python,
                env=env,
                **kwargs,
            )
            app_proc.wait()
            _kill(vite)

    except KeyboardInterrupt:
        pass
    finally:
        _cleanup()


if __name__ == "__main__":
    main()
