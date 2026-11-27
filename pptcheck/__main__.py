"""Start PPTcheck.

    python -m pptcheck                        open the PPTcheck window
    python -m pptcheck --scan PATH [--out report.json] [--pdf report.pdf]
                                              command-line mode, no window
    python -m pptcheck --gui-selftest OUT.json
                                              build check: open the window, analyse the
                                              test deck through the interface, write results

When packaged as an app there is no console window; any messages go to a
pptcheck.log file in the system temp folder.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile


def _redirect_output_if_windowed() -> None:
    """Windowed apps have no console: send stdout/stderr to a log file."""
    if sys.stdout is not None and sys.stderr is not None:
        return
    log = open(os.path.join(tempfile.gettempdir(), "pptcheck.log"), "a", encoding="utf-8", buffering=1)
    sys.stdout = sys.stdout or log
    sys.stderr = sys.stderr or log


def _scan(args) -> None:
    from .analyzer import analyze_pptx
    from .files import list_folder
    target = os.path.abspath(args.scan)
    paths = list_folder(target) if os.path.isdir(target) else [target]
    reports = []
    for i, p in enumerate(paths, 1):
        print(f"[{i}/{len(paths)}] {p}", file=sys.stderr)
        reports.append(analyze_pptx(p, display_name=os.path.relpath(p, target) if os.path.isdir(target) else None,
                                    source_dir=os.path.dirname(p)))
    if args.pdf:
        from .report_pdf import build_pdf
        build_pdf(reports, args.pdf)
        print(f"Wrote {args.pdf}", file=sys.stderr)
        if not args.out:
            return
    text = json.dumps(reports, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"Wrote {args.out}", file=sys.stderr)
    else:
        print(text)


def main() -> None:
    _redirect_output_if_windowed()
    parser = argparse.ArgumentParser(prog="pptcheck", description="PowerPoint technical auditor")
    parser.add_argument("--scan", metavar="PATH", help="command-line mode: analyse a .pptx file or folder")
    parser.add_argument("--out", metavar="FILE", help="with --scan: write JSON here (default: print)")
    parser.add_argument("--pdf", metavar="FILE", help="with --scan: write a PDF report here")
    parser.add_argument("--gui-selftest", metavar="OUT", help=argparse.SUPPRESS)
    parser.add_argument("--selftest-file", metavar="PPTX", help=argparse.SUPPRESS)
    # macOS may pass "-psn_..." when launching apps; ignore unknown arguments.
    args, _unknown = parser.parse_known_args()

    if args.scan:
        _scan(args)
        return

    from .desktop import run
    sys.exit(run(selftest_out=args.gui_selftest, selftest_file=args.selftest_file))


if __name__ == "__main__":
    main()
