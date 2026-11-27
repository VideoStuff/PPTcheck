"""Check that a packaged PPTcheck build actually works.

Usage: python packaging/smoke_test.py PATH_TO_PACKAGED_EXECUTABLE [--gui]

1. Runs the app in --scan mode on tests/fixtures/sample_deck.pptx and checks
   the results, including video metrics (proves MediaInfo was bundled).
2. With --gui: opens the real app window, analyses the same deck through the
   interface (choose -> analyse -> detail view), and checks what the window
   showed, and creates the PDF report. A screenshot is saved to smoke/gui.png
   and the PDF to smoke/gui.pdf.
"""
import json
import os
import subprocess
import sys

exe = os.path.abspath(sys.argv[1])
gui = "--gui" in sys.argv
root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
deck = os.path.join(root, "tests", "fixtures", "sample_deck.pptx")
out_dir = os.path.join(root, "smoke")
os.makedirs(out_dir, exist_ok=True)
failed = False


def check(name, ok):
    global failed
    print(f"{'PASS' if ok else 'FAIL'}  {name}")
    failed |= not ok


env = dict(os.environ)

report_path = os.path.join(out_dir, "report.json")
subprocess.run([exe, "--scan", deck, "--out", report_path], check=True, timeout=300, env=env)
report = json.load(open(report_path, encoding="utf-8"))[0]
video = report["videos"][0]
check("slide count", report["slide_count"] == 4)
check("video duration", video["duration"] == "00:00:05")
check("video fps", video["fps"] == 30.0)
check("video resolution", video["resolution"] == "1280 × 720 px")
check("slide size in inches", (report["slide_size"] or {}).get("text") == "10 × 7.5 in (960 × 720 px) · 4:3")
check("video aspect ratio", video["aspect_ratio"] == "16:9")
# "MPEG-4" is how MediaInfo names the container (ffprobe would say "QuickTime / MOV"),
# which proves the bundled MediaInfo library is the one in use.
check("bundled MediaInfo engine used", video["container"] == "MPEG-4")
check("transitions", {t["type"] for t in report["transitions"]} == {"Morph", "Fade"})
check("slido add-in", "Slido" in report["summary"]["addins"])
check("fonts", "Gotham Rounded" in report["summary"]["fonts_to_install"])

if gui:
    gui_path = os.path.join(out_dir, "gui.json")
    proc = subprocess.run([exe, "--gui-selftest", gui_path, "--selftest-file", deck], timeout=300, env=env)
    check("window exited cleanly", proc.returncode == 0)
    result = json.load(open(gui_path, encoding="utf-8")) if os.path.exists(gui_path) else {"ok": False, "error": "no result file"}
    check("window self-test completed", result.get("ok") is True)
    page = result.get("page") or {}
    check("window: one row in table", page.get("rows") == 1)
    check("window: video metrics shown", "00:00:05" in page.get("detail", ""))
    check("window: fonts section shown", "Gotham Rounded" in page.get("detail", ""))
    check("window: video engine ready", "video metrics" in page.get("engine", ""))
    pdf_path = os.path.splitext(gui_path)[0] + ".pdf"
    pdf_ok = os.path.exists(pdf_path) and open(pdf_path, "rb").read(5) == b"%PDF-" and os.path.getsize(pdf_path) > 5000
    check("window: PDF report created", page.get("pdf_written") is True and pdf_ok)
    if not result.get("ok"):
        print(json.dumps(result, indent=2)[:3000])

if failed:
    sys.exit(1)
print("Packaged app works.")
