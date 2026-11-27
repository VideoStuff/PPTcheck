"""Run with:  python -m pytest tests  (needs FFmpeg + python-pptx)."""
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from make_sample import build  # noqa: E402
from pptcheck.analyzer import analyze_pptx  # noqa: E402
from pptcheck.media import media_available as ffprobe_available  # noqa: E402


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
        assert v["video_codec"].startswith("H.264")


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


def test_fonts(report):
    f = {x["name"]: x for x in report["fonts"]}
    assert f["Montserrat SemiBold"]["status"] == "google" and f["Montserrat SemiBold"]["needs_install"]
    assert f["Montserrat SemiBold"]["links"][0]["url"] == "https://fonts.google.com/specimen/Montserrat"
    assert f["Gotham Rounded"]["status"] == "nonstandard" and f["Gotham Rounded"]["slides"] == [3]
    assert f["Segoe UI"]["status"] == "windows" and f["Segoe UI"]["platform_warning"]
    assert f["Calibri"]["status"] == "office" and "Theme body font" in f["Calibri"]["uses"]
    assert f["Brand Sans"]["embedded"] and not f["Brand Sans"]["needs_install"]
    assert sorted(report["summary"]["fonts_to_install"]) == ["Gotham Rounded", "Montserrat SemiBold"]


def test_font_classification():
    from pptcheck.fonts import classify
    assert classify("Source Sans Pro")["family"] == "Source Sans 3"
    assert classify("Open Sans Light")["status"] == "google"
    assert classify("Aptos Display")["status"] == "cloud"
    assert classify("Arial")["status"] == "standard"
    assert classify("Proxima Nova")["status"] == "nonstandard"
