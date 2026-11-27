"""Check that a packaged PPTcheck build actually works.

Usage: python packaging/smoke_test.py PATH_TO_PACKAGED_EXECUTABLE
Runs the app in --scan mode on tests/fixtures/sample_deck.pptx and checks the
results, including video metrics (proves the MediaInfo library was bundled).
"""
import json
import os
import subprocess
import sys
import tempfile

exe = os.path.abspath(sys.argv[1])
root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
deck = os.path.join(root, "tests", "fixtures", "sample_deck.pptx")
out = os.path.join(tempfile.mkdtemp(), "report.json")

# Hide any system FFmpeg so only the bundled engine can be used.
env = dict(os.environ, PATH=os.path.dirname(exe))
subprocess.run([exe, "--scan", deck, "--out", out], check=True, timeout=300, env=env)

report = json.load(open(out, encoding="utf-8"))[0]
video = report["videos"][0]
checks = {
    "slide count": report["slide_count"] == 4,
    "video duration": video["duration"] == "00:00:05",
    "video fps": video["fps"] == 30.0,
    "video resolution": video["resolution"] == "1280x720",
    "transitions": {t["type"] for t in report["transitions"]} == {"Morph", "Fade"},
    "slido add-in": "Slido" in report["summary"]["addins"],
    "fonts": "Gotham Rounded" in report["summary"]["fonts_to_install"],
}
for name, ok in checks.items():
    print(f"{'PASS' if ok else 'FAIL'}  {name}")
if not all(checks.values()):
    print(json.dumps(video, indent=2))
    sys.exit(1)
print("Packaged app works.")
