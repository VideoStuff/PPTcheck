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
        assert v["resolution"] == "1280 × 720 px"
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


def test_expand_paths(tmp_path):
    from pptcheck.files import expand_paths
    (tmp_path / "Decks" / "sub").mkdir(parents=True)
    for rel in ("Decks/a.pptx", "Decks/~$a.pptx", "Decks/notes.txt", "Decks/sub/b.pptx", "other/a.pptx"):
        f = tmp_path / rel
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_bytes(b"x")
    res = expand_paths([str(tmp_path / "Decks"), str(tmp_path / "other" / "a.pptx"), str(tmp_path / "missing.pptx")])
    names = [j["name"] for j in res["jobs"]]
    assert names == ["Decks/a.pptx", "Decks/sub/b.pptx", "a.pptx"]
    assert res["errors"] and "missing.pptx" in res["errors"][0]
    top = expand_paths([str(tmp_path / "Decks")], recursive=False)
    assert [j["name"] for j in top["jobs"]] == ["Decks/a.pptx"]


def test_page_is_self_contained():
    import re
    from pptcheck.desktop import build_page_html
    html = build_page_html()
    assert "Content-Security-Policy" in html and "connect-src 'none'" in html
    assert '<script src=' not in html and 'rel="stylesheet"' not in html
    # The page must not reference any outside resource (only links the user clicks).
    assert not re.search(r'(src|href)="https?://', html)
    assert "fetch(" not in html


def test_pdf_report(report, tmp_path):
    from pptcheck.report_pdf import build_pdf, items_to_check
    items = " ".join(items_to_check(report))
    assert "Gotham Rounded" in items and "Slido" in items and "automatically" in items
    assert "Brand Sans" not in items  # embedded fonts don't need installing
    single = build_pdf([report], str(tmp_path / "one.pdf"))
    data = open(single, "rb").read()
    assert data.startswith(b"%PDF-") and len(data) > 5000
    broken = {"file_name": "broken.pptx", "error": "Not a valid .pptx", "slide_count": 0}
    multi = build_pdf([report, broken], str(tmp_path / "many.pdf"))
    assert open(multi, "rb").read(5) == b"%PDF-"


def test_slide_size_and_video_shape(report):
    size = report["slide_size"]
    assert (size["width_in"], size["height_in"]) == (10.0, 7.5)
    assert (size["width_px"], size["height_px"]) == (960, 720) and size["aspect_ratio"] == "4:3"
    v = report["videos"][0]
    assert v["resolution"] == "1280 × 720 px" and v["aspect_ratio"] == "16:9"
    p = v["placements"][0]
    # The sample places the 16:9 video in a 4 x 3 inch frame, so it is squashed.
    assert (p["slide"], p["width_in"], p["height_in"], p["aspect_ratio"]) == (2, 4.0, 3.0, "4:3")
    assert p["stretched"] and v["stretched_on_slides"] == [2]
    from pptcheck.report_pdf import items_to_check
    assert any("stretched" in i for i in items_to_check(report))


def test_geometry_helpers():
    from pptcheck.geometry import aspect_label, is_stretched, placement, slide_size
    assert aspect_label(1920, 1080) == "16:9" and aspect_label(1024, 768) == "4:3"
    assert aspect_label(1080, 1920) == "9:16" and aspect_label(2.39, 1) == "2.39:1"
    wide = slide_size(12192000, 6858000)          # PowerPoint's default widescreen slide
    assert wide["text"] == "13.33 × 7.5 in (1280 × 720 px) · 16:9" and wide["name"] == "Widescreen"
    frame = placement(6096000, 3429000)            # 6.67 x 3.75 in, 16:9
    assert not is_stretched(frame, 16 / 9) and is_stretched(frame, 4 / 3)
    # Cropping the sides of a 16:9 video to a 4:3 frame is not distortion.
    cropped = placement(4572000, 3429000, {"l": 12500, "r": 12500})
    assert cropped["aspect_ratio"] == "4:3" and not is_stretched(cropped, 16 / 9)
