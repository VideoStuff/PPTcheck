"""Run with:  python -m pytest tests  (needs FFmpeg + python-pptx)."""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from make_sample import build  # noqa: E402
from pptcheck.analyzer import analyze_pptx  # noqa: E402
from pptcheck.media import ffprobe_available  # noqa: E402


@pytest.fixture(scope="module")
def report(tmp_path_factory):
    path = build(str(tmp_path_factory.mktemp("deck")))
    return analyze_pptx(path)


def test_overview(report):
    assert report["error"] is None
    assert report["slide_count"] == 4
    assert report["hidden_slide_count"] == 1
    assert report["summary"]["has_notes"] and report["summary"]["slides_with_notes"] == 1


def test_video(report):
    assert report["summary"]["video_count"] == 1
    v = report["videos"][0]
    assert v["slides"] == [2] and v["source"] == "Embedded" and v["size_bytes"] > 0
    if ffprobe_available():
        assert v["duration"] == "00:00:05"
        assert v["fps"] == 30.0
        assert v["resolution"] == "1280x720"
        assert v["video_codec"].startswith("h264")


def test_transitions_and_animations(report):
    t = {x["slide"]: x for x in report["transitions"]}
    assert t[1]["type"] == "Morph"
    assert t[2]["type"] == "Fade" and t[2]["auto_advance_seconds"] == 5.0 and not t[2]["advance_on_click"]
    assert report["summary"]["has_animations"]
    anim = {a["slide"]: a for a in report["animations"]}
    assert anim[1]["entrance"] == 1


def test_addins(report):
    a = {x["name"]: x for x in report["addins"]}
    assert a["Slido"]["slides"] == [3] and a["Slido"]["confidence"] == "high"
    assert "Office Web Add-in" in a["Slido"]["sources"]
    assert 2 in a["Mentimeter"]["slides"]


def test_bad_file(tmp_path):
    p = tmp_path / "broken.pptx"
    p.write_bytes(b"not a zip")
    r = analyze_pptx(str(p))
    assert r["error"]
