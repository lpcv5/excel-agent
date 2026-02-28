"""Cross-platform Nuitka build script for pywebview + React app."""

import os
import platform
import subprocess
import sys


def run(cmd: list[str]) -> None:
    print(f">>> {' '.join(cmd)}")
    subprocess.check_call(cmd, shell=(sys.platform == "win32"))


# ── Packages to FULLY EXCLUDE (never needed at runtime) ─────────────────────
# --nofollow-import-to means "do NOT include this at all".
# Only put packages here that are NEVER used at runtime.
EXCLUDE_PACKAGES = [
    # Dev tools
    "pyright",
    "ruff",
    "nodeenv",
    "nuitka",
    # Testing
    "*.tests",
    "*.test",
    "*.tests._testing",
    "unittest",
    "doctest",
    "test",
    "pytest",
    # Too large to compile to C (assembler overflow)
    "google.genai",
    # Unused large stdlib modules
    "tkinter",
    "sqlite3",
    "xml.etree",
    "xmlrpc",
    "pydoc",
    "lib2to3",
    "ensurepip",
    "venv",
    "distutils",
    "setuptools",
    "pip",
]

# ── Packages whose data files must be preserved ─────────────────────────────
INCLUDE_PACKAGE_DATA = [
    "pydantic",
    "pydantic_core",
    "tiktoken",
    "certifi",
    "uvicorn",
    "starlette",
    "fastapi",
    "langchain",
    "langchain_core",
    "langgraph",
]

# ── Dynamically imported modules that Nuitka cannot discover automatically ───
INCLUDE_MODULES = [
    # uvicorn selects HTTP/WS implementation and event loop at runtime
    "uvicorn.lifespan.on",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.http.h11_impl",
    "uvicorn.protocols.http.httptools_impl",
    "uvicorn.protocols.websockets.auto",
    "uvicorn.protocols.websockets.wsproto_impl",
    "uvicorn.loops.auto",
    "uvicorn.loops.asyncio",
    # stdlib modules used by FastAPI/Starlette
    "email.mime.multipart",
    "email.mime.text",
]

# ── Packages to include but NOT compile to C (bytecode only) ────────────────
# This speeds up build significantly for large pure-Python packages.
# Nuitka flag: --include-package=X ensures inclusion,
#              --force-dll-dependency-cache-update is not needed here.
# We use --module-parameter to keep them as bytecode where supported,
# otherwise we just let Nuitka compile everything (slower but correct).
BYTECODE_ONLY_PACKAGES = [
    "langchain",
    "langchain_core",
    "langchain_anthropic",
    "langchain_openai",
    "langchain_google_genai",
    "langgraph",
    "langgraph_sdk",
    "langsmith",
    "anthropic",
    "openai",
    "google",
    "httpx",
    "httpcore",
    "requests",
    "urllib3",
    "anyio",
    "sniffio",
    "h11",
    "httptools",
    "websockets",
    "watchfiles",
    "cryptography",
    "cffi",
    "pycparser",
    "orjson",
    "ormsgpack",
    "jiter",
    "tiktoken",
    "regex",
    "certifi",
    "charset_normalizer",
    "idna",
    "tenacity",
    "tqdm",
    "yaml",
    "dotenv",
    "jsonpatch",
    "jsonpointer",
    "zstandard",
    "xxhash",
    "packaging",
    "typing_extensions",
    "typing_inspection",
    "annotated_types",
    "annotated_doc",
    "docstring_parser",
    "filetype",
    "wcmatch",
    "bracex",
    "bottle",
    "proxy_tools",
    "colorama",
    "distro",
    "pydantic",
    "pydantic_core",
    "fastapi",
    "starlette",
    "uvicorn",
    "click",
]


def build_nuitka_cmd(*, release: bool = False) -> list[str]:
    cmd = [
        "uv",
        "run",
        "python",
        "-m",
        "nuitka",
        "--onefile" if release else "--standalone",
        f"--jobs={os.cpu_count() or 4}",
        "--include-data-dir=dist=dist",
        "--output-dir=build",
        "--assume-yes-for-downloads",
        # Suppress the deployment error about excluded modules
        "--no-deployment-flag=excluded-module-usage",
    ]

    if not release:
        cmd.append("--lto=no")

    # ── Fully exclude unused packages ──
    for pkg in EXCLUDE_PACKAGES:
        cmd.append(f"--nofollow-import-to={pkg}")

    # ── Include packages (bundle as bytecode, don't compile to C) ──
    for pkg in BYTECODE_ONLY_PACKAGES:
        cmd.append(f"--include-package={pkg}")

    # ── Preserve data files ──
    for pkg in INCLUDE_PACKAGE_DATA:
        cmd.append(f"--include-package-data={pkg}")

    # ── Explicitly include dynamically imported modules ──
    for mod in INCLUDE_MODULES:
        cmd.append(f"--include-module={mod}")

    # ── Platform-specific options ──
    system = platform.system()
    if system == "Windows":
        cmd.extend([
            "--windows-console-mode=disable",
            "--mingw64",
        ])
    elif system == "Darwin":
        cmd.append("--macos-create-app-bundle")

    cmd.append("src-python/main.py")
    return cmd


def main() -> None:
    release = "--release" in sys.argv

    mode = "release (onefile)" if release else "dev (standalone)"
    print(f"\n{'='*50}")
    print(f"  Building in {mode} mode")
    print(f"  CPU cores: {os.cpu_count()}")
    print(f"  Platform:  {platform.system()}")
    print(f"{'='*50}\n")

    run(["bun", "run", "build"])

    cmd = build_nuitka_cmd(release=release)
    run(cmd)

    print(f"\n✅ Build complete ({mode})! Check the build/ directory.")


if __name__ == "__main__":
    main()