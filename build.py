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
    "doctest",
    "test",
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
    "pip",
    "lzma",
    "bz2",
    # Unused win32 modules (only need win32com, win32api, win32event,
    # win32ui, win32gui, win32file, win32security, win32process)
    "dde",
    "mmapfile",
    "odbc",
    "perfmon",
    "servicemanager",
    "timer",
    "win32clipboard",
    "win32console",
    "win32cred",
    "win32crypt",
    "win32evtlog",
    "win32help",
    "win32inet",
    "win32job",
    "win32lz",
    "win32net",
    "win32pdh",
    "win32pipe",
    "win32print",
    "win32profile",
    "win32ras",
    "win32service",
    "win32trace",
    "win32transaction",
    "win32ts",
    "win32uiole",
    "win32wnet",
    "_winxptheme",
    "_wmi",
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

# ── Packages with lazy imports that need --include-package for discovery ─────
# These use __getattr__ / importlib.import_module in sub-package __init__.py,
# so Nuitka can't discover submodules via static analysis alone.
FORCE_INCLUDE_PACKAGES = [
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
    "pydantic",
    "anyio",
]

# ── Packages to include as bytecode only (not compiled to C) ────────────────
# Uses Nuitka anti-bloat plugin's --noinclude-custom-mode=PKG:bytecode.
# These packages are included but kept as .pyc, drastically reducing C modules.
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
    "adodbapi",
    "win32com",
    "email",
    "http",
    "asyncio",
    "multiprocessing",
    "concurrent",
    "importlib",
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
    deepagents_base_prompt = os.path.join(
        "src-python", "libs", "deepagents", "base_prompt.md"
    )
    cmd.append(
        f"--include-data-file={deepagents_base_prompt}=libs/deepagents/base_prompt.md"
    )

    if not release:
        cmd.extend(["--lto=no", "--python-flag=no_docstrings"])
    else:
        cmd.append("--python-flag=no_asserts")

    # ── Fully exclude unused packages ──
    for pkg in EXCLUDE_PACKAGES:
        cmd.append(f"--nofollow-import-to={pkg}")

    # ── Force-include packages with lazy imports (for submodule discovery) ──
    for pkg in FORCE_INCLUDE_PACKAGES:
        cmd.append(f"--include-package={pkg}")

    # ── Include packages as bytecode only (not compiled to C) ──
    for pkg in BYTECODE_ONLY_PACKAGES:
        cmd.append(f"--noinclude-custom-mode={pkg}:bytecode")

    # ── Preserve data files ──
    for pkg in INCLUDE_PACKAGE_DATA:
        cmd.append(f"--include-package-data={pkg}")

    # ── Explicitly include dynamically imported modules ──
    for mod in INCLUDE_MODULES:
        cmd.append(f"--include-module={mod}")

    # ── Built-in anti-bloat modes ──
    cmd.extend([
        "--noinclude-setuptools-mode=nofollow",
        "--noinclude-pytest-mode=nofollow",
        "--noinclude-unittest-mode=nofollow",
    ])

    # ── Platform-specific options ──
    system = platform.system()
    if system == "Windows":
        cmd.extend([
            "--windows-console-mode=disable" if release else "--windows-console-mode=attach",
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
