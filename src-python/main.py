import argparse
import os
import sys
import threading
import time
import urllib.request


def _wait_for_health(url: str, timeout: int = 30) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(url, timeout=2)
            return True
        except Exception:
            time.sleep(0.3)
    return False


def main() -> None:
    if "__compiled__" in globals():
        # Nuitka bundle — serve from dist/ next to the executable
        import uvicorn
        from server import app, APP_TOKEN

        base = __nuitka_binary_dir  # type: ignore[name-defined]  # noqa: F821
        dist_dir = os.path.join(base, "dist")
        port = 8765

        from fastapi.staticfiles import StaticFiles

        app.mount("/", StaticFiles(directory=dist_dir, html=True), name="static")

        def _serve():
            uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")

        t = threading.Thread(target=_serve, daemon=True)
        t.start()

        health_url = f"http://127.0.0.1:{port}/health"
        if not _wait_for_health(health_url):
            print("Server failed to start.", file=sys.stderr)
            sys.exit(1)

        import webview
        from server import set_webview_window

        url = f"http://127.0.0.1:{port}?token={APP_TOKEN}"
        window = webview.create_window("ExcelAgent", url=url, width=1600, height=900)
        set_webview_window(window)
        webview.start(debug=False)
        return

    parser = argparse.ArgumentParser(description="ExcelAgent")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument(
        "--no-window",
        action="store_true",
        help="B/S mode: print URL and run server only",
    )
    parser.add_argument(
        "--dev-frontend",
        default=None,
        help="Dev frontend URL, e.g. http://localhost:5173",
    )
    args = parser.parse_args()

    import uvicorn
    from server import app, APP_TOKEN

    port = args.port

    if args.no_window:
        print(f"http://127.0.0.1:{port}?token={APP_TOKEN}", flush=True)
        uvicorn.run(app, host="127.0.0.1", port=port)
        return

    # Desktop mode: start uvicorn in background thread, then open pywebview
    server = uvicorn.Server(
        uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning")
    )

    def _serve():
        server.run()

    t = threading.Thread(target=_serve, daemon=True)
    t.start()

    health_url = f"http://127.0.0.1:{port}/health"
    print(f"Waiting for API server on {health_url}...")
    if not _wait_for_health(health_url):
        print("API server failed to start.", file=sys.stderr)
        sys.exit(1)
    print("API server ready.")

    import webview
    from server import set_webview_window

    frontend_url = args.dev_frontend or f"http://127.0.0.1:{port}"
    url = f"{frontend_url}?token={APP_TOKEN}"
    window = webview.create_window("ExcelAgent", url=url, width=1600, height=900)
    set_webview_window(window)
    webview.start(debug=bool(args.dev_frontend))

    # Window closed — force exit so uvicorn and all threads are terminated
    os._exit(0)


if __name__ == "__main__":
    main()
