"""Start PPTcheck.

    python -m pptcheck                 start the dashboard and open the browser
    python -m pptcheck --port 9000     use a different port
    python -m pptcheck --no-browser    don't open the browser
    python -m pptcheck --scan PATH [--out report.json]
                                       command-line mode, no browser

When packaged as a desktop app (PyInstaller) there is no terminal window:
  * opening the app a second time just reopens the browser tab;
  * the app quits via the dashboard's "Quit" button, or by itself after
    IDLE_MINUTES with no dashboard tab open;
  * log output goes to a pptcheck.log file in the system temp folder.
"""
from __future__ import annotations

import argparse
import json
import os
import socket
import sys
import tempfile
import threading
import time
import urllib.request
import webbrowser

DEFAULT_PORT = 8765
IDLE_MINUTES = 15


def _redirect_output_if_windowed() -> str | None:
    """Windowed apps have no console: send stdout/stderr to a log file."""
    if sys.stdout is not None and sys.stderr is not None:
        return None
    log_path = os.path.join(tempfile.gettempdir(), "pptcheck.log")
    log = open(log_path, "a", encoding="utf-8", buffering=1)
    sys.stdout = sys.stdout or log
    sys.stderr = sys.stderr or log
    return log_path


def _pptcheck_running_on(port: int) -> bool:
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{port}/api/health", timeout=1.5) as r:
            return json.load(r).get("app") == "pptcheck"
    except Exception:
        return False


def _port_free(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(("127.0.0.1", port))
            return True
        except OSError:
            return False


def _pick_port(preferred: int) -> int:
    for port in range(preferred, preferred + 20):
        if _port_free(port):
            return port
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def _scan(args) -> None:
    from .analyzer import analyze_pptx
    from .server import is_presentation
    target = os.path.abspath(args.scan)
    if os.path.isdir(target):
        paths = [os.path.join(r, n) for r, _, ns in os.walk(target) for n in sorted(ns) if is_presentation(n)]
    else:
        paths = [target]
    reports = []
    for i, p in enumerate(paths, 1):
        print(f"[{i}/{len(paths)}] {p}", file=sys.stderr)
        reports.append(analyze_pptx(p, display_name=os.path.relpath(p, target) if os.path.isdir(target) else None,
                                    source_dir=os.path.dirname(p)))
    text = json.dumps(reports, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"Wrote {args.out}", file=sys.stderr)
    else:
        print(text)


def main() -> None:
    log_path = _redirect_output_if_windowed()
    parser = argparse.ArgumentParser(prog="pptcheck", description="PowerPoint technical auditor")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--no-browser", action="store_true", help="don't open the browser automatically")
    parser.add_argument("--scan", metavar="PATH", help="command-line mode: analyse a .pptx file or folder")
    parser.add_argument("--out", metavar="FILE", help="with --scan: write JSON here (default: print)")
    # macOS passes "-psn_..." when launching some apps; ignore unknown args.
    args, _unknown = parser.parse_known_args()

    if args.scan:
        _scan(args)
        return

    # Already running? Just show it again.
    if _pptcheck_running_on(args.port):
        if not args.no_browser:
            webbrowser.open(f"http://127.0.0.1:{args.port}")
        print(f"PPTcheck is already running at http://127.0.0.1:{args.port}")
        return

    import uvicorn

    from . import server
    from .media import engine_name
    from .paths import is_frozen

    port = _pick_port(args.port)
    url = f"http://127.0.0.1:{port}"
    print(f"\n  PPTcheck is running at {url}\n  Use the Quit button in the dashboard (or Ctrl+C here) to stop it.\n")
    if log_path:
        print(f"  (log file: {log_path})")
    if engine_name() is None:
        print("  WARNING: no video engine found (MediaInfo or FFmpeg); video details will be unavailable.\n")

    config = uvicorn.Config(server.app, host="127.0.0.1", port=port, log_level="warning",
                            log_config=None if log_path else uvicorn.config.LOGGING_CONFIG)
    srv = uvicorn.Server(config)
    server.server_handle = srv

    if not args.no_browser:
        def open_when_ready():
            for _ in range(100):
                if _pptcheck_running_on(port):
                    webbrowser.open(url)
                    return
                time.sleep(0.2)
        threading.Thread(target=open_when_ready, daemon=True).start()

    if is_frozen():
        # No terminal window to close, so stop by ourselves once no dashboard
        # tab has checked in for a while.
        def idle_watchdog():
            while not srv.should_exit:
                time.sleep(30)
                if time.monotonic() - server.last_heartbeat > IDLE_MINUTES * 60:
                    print("No dashboard open; shutting down.")
                    srv.should_exit = True
        threading.Thread(target=idle_watchdog, daemon=True).start()

    srv.run()


if __name__ == "__main__":
    main()
