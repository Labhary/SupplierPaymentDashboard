from __future__ import annotations

import socket
import subprocess
import sys
import threading
import time
import webbrowser
import os
from pathlib import Path
from urllib.request import urlopen


def _base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent


def _find_free_port(start: int = 8501, attempts: int = 20) -> int:
    for port in range(start, start + attempts):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.2)
            if sock.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return start


def _open_browser_when_ready(url: str, timeout_seconds: int = 30) -> None:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        try:
            with urlopen(url, timeout=1):
                webbrowser.open(url)
                return
        except Exception:
            time.sleep(0.5)
    webbrowser.open(url)


def main() -> int:
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

    base_dir = _base_dir()
    app_path = base_dir / "app.py"
    if not app_path.exists():
        print(f"Unable to find app.py at: {app_path}")
        return 1

    port = _find_free_port()
    url = f"http://localhost:{port}"
    browser_thread = threading.Thread(target=_open_browser_when_ready, args=(url,), daemon=True)
    browser_thread.start()

    args = [
        "streamlit",
        "run",
        str(app_path),
        "--server.address=localhost",
        f"--server.port={port}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--global.developmentMode=false",
    ]

    if getattr(sys, "frozen", False):
        sys.path.insert(0, str(base_dir))
        from streamlit.web import cli as streamlit_cli

        sys.argv = args
        streamlit_cli.main()
        return 0

    command = [sys.executable, "-m", *args]
    return subprocess.call(command, cwd=str(base_dir), env=os.environ.copy())


if __name__ == "__main__":
    raise SystemExit(main())
