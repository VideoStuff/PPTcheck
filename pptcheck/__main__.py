"""Start PPTcheck:  python -m pptcheck  [--port 8765] [--no-browser]

Also usable without the web UI:
    python -m pptcheck --scan "C:\\path\\to\\folder" --out report.json
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import threading
import webbrowser


def main() -> None:
    parser = argparse.ArgumentParser(prog="pptcheck", description="PowerPoint technical auditor")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true", help="don't open the browser automatically")
    parser.add_argument("--scan", metavar="PATH", help="command-line mode: analyse a .pptx file or folder")
    parser.add_argument("--out", metavar="FILE", help="with --scan: write JSON here (default: print)")
    args = parser.parse_args()

    if args.scan:
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
        return

    import uvicorn
    from .media import ffprobe_available

    url = f"http://127.0.0.1:{args.port}"
    print(f"\n  PPTcheck is running at {url}\n  Press Ctrl+C in this window to stop it.\n")
    if not ffprobe_available():
        print("  WARNING: ffprobe (FFmpeg) was not found. Video duration/FPS/resolution\n"
              "  will be unavailable until FFmpeg is installed.\n")
    if not args.no_browser:
        threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    uvicorn.run("pptcheck.server:app", host="127.0.0.1", port=args.port, log_level="warning")


if __name__ == "__main__":
    main()
