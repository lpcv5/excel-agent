"""Cross-platform Nuitka build script for pywebview + React app."""

import platform
import subprocess
import sys


def run(cmd: list[str]) -> None:
    print(f">>> {' '.join(cmd)}")
    subprocess.check_call(cmd, shell=(sys.platform == "win32"))


def main() -> None:
    # 1. Build the React frontend
    run(["npm", "run", "build"])

    # 2. Assemble Nuitka command
    cmd = [
        "uv",
        "run",
        "python",
        "-m",
        "nuitka",
        "--standalone",
        "--onefile",
        "--enable-plugin=pywebview",
        "--include-data-dir=dist=dist",
        "--output-dir=build",
        "--assume-yes-for-downloads",
    ]

    system = platform.system()
    if system == "Windows":
        cmd.append("--windows-console-mode=disable")
    elif system == "Darwin":
        cmd.append("--macos-create-app-bundle")

    cmd.append("src-python/main.py")

    # 3. Compile
    run(cmd)
    print("\nBuild complete! Check the build/ directory for output.")


if __name__ == "__main__":
    main()
