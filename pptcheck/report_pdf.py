"""PDF summary report of PPTcheck findings (generated locally with ReportLab).

build_pdf(reports, path) writes one PDF covering one or more analysed
presentations: an overview table (for several files), then for each
presentation an "Items to check" list followed by compact tables of videos,
transitions, add-ins and fonts.
"""
from __future__ import annotations

import datetime as _dt
import os
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (PageBreak, Paragraph, SimpleDocTemplate, Spacer,
                                Table, TableStyle)

from . import __version__
from .paths import resource_dir

ACCENT = colors.HexColor("#c4442b")
TEXT = colors.HexColor("#1b2230")
MUTED = colors.HexColor("#667085")
BORDER = colors.HexColor("#dde2ea")
HEAD_BG = colors.HexColor("#f0f3f8")
WARN_BG = colors.HexColor("#fdf1d8")
GOOD = colors.HexColor("#1f7a4d")
BAD = colors.HexColor("#b42318")

LARGE_VIDEO_MB = 100
# Formats Microsoft recommends for PowerPoint: .mp4 with H.264 video + AAC audio.
WIDELY_SUPPORTED_VIDEO = ("H.264", "H.265")

_FONTS_READY = False


def _register_fonts() -> tuple[str, str]:
    """Use bundled DejaVu Sans (covers accented/most non-English text)."""
    global _FONTS_READY
    if not _FONTS_READY:
        folder = os.path.join(resource_dir(), "assets", "fonts")
        try:
            pdfmetrics.registerFont(TTFont("PPT-Regular", os.path.join(folder, "DejaVuSans.ttf")))
            pdfmetrics.registerFont(TTFont("PPT-Bold", os.path.join(folder, "DejaVuSans-Bold.ttf")))
            _FONTS_READY = True
        except Exception:
            return "Helvetica", "Helvetica-Bold"
    return "PPT-Regular", "PPT-Bold"


def _styles():
    regular, bold = _register_fonts()

    def mk(name, bold_font=False, **kw):
        opts = dict(fontName=bold if bold_font else regular, textColor=TEXT, alignment=TA_LEFT)
        opts.update(kw)
        return ParagraphStyle(name, **opts)

    return {
        "title": mk("title", True, fontSize=20, leading=24, spaceAfter=4),
        "meta": mk("meta", fontSize=8.5, leading=11, textColor=MUTED),
        "h1": mk("h1", True, fontSize=14, leading=18, spaceBefore=4, spaceAfter=4),
        "h2": mk("h2", True, fontSize=10.5, leading=14, spaceBefore=10, spaceAfter=4, textColor=ACCENT),
        "body": mk("body", fontSize=9, leading=12),
        "cell": mk("cell", fontSize=7.8, leading=9.6),
        "cellb": mk("cellb", True, fontSize=7.8, leading=9.6),
        "bullet": mk("bullet", fontSize=9, leading=12, leftIndent=12, bulletIndent=2),
        "ok": mk("ok", fontSize=9, leading=12, textColor=GOOD),
        "err": mk("err", fontSize=9, leading=12, textColor=BAD),
    }


def _p(text, style) -> Paragraph:
    return Paragraph(escape(str(text if text is not None else "—")), style)


def _plural(n: int, word: str) -> str:
    return f"{n} {word}{'' if n == 1 else 's'}"


def _slides(nums) -> str:
    nums = list(nums or [])
    if not nums:
        return "—"
    return ("slide " if len(nums) == 1 else "slides ") + ", ".join(str(n) for n in nums)


def _table(headers, rows, widths, st, header_style="cellb"):
    data = [[_p(h, st[header_style]) for h in headers]]
    for row in rows:
        data.append([c if isinstance(c, Paragraph) else _p(c, st["cell"]) for c in row])
    t = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HEAD_BG),
        ("LINEBELOW", (0, 0), (-1, -1), 0.4, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t


# --------------------------------------------------------------------------- #
# "Items to check"
# --------------------------------------------------------------------------- #

def items_to_check(r: dict) -> list[str]:
    """Plain-language list of things worth checking before presenting."""
    if r.get("error"):
        return [f"PPTcheck could not read this file: {r['error']}"]
    s = r.get("summary") or {}
    items: list[str] = []
    fonts = r.get("fonts") or []

    install = [f for f in fonts if f.get("needs_install")]
    if install:
        items.append("Fonts that are not standard (install them on the presenting computer, or embed fonts "
                     "via File > Options > Save > Embed fonts): "
                     + "; ".join(f"{f['name']} ({f['status_label']})" for f in install) + ".")
    platform = [f for f in fonts if f.get("platform_warning")]
    if platform:
        items.append("Fonts that may be substituted on some computers: "
                     + "; ".join(f"{f['name']} ({f['status_label']})" for f in platform) + ".")

    for a in r.get("addins") or []:
        if a.get("confidence") == "low":
            continue
        where = ", ".join(a.get("locations") or []) or "—"
        items.append(f"Interactive add-in: {a['name']} ({where}). Needs an internet connection and the "
                     f"add-in/account available on the presenting computer.")

    for v in (r.get("videos") or []) + (r.get("audio") or []):
        label = f"\"{v.get('name') or v.get('file')}\" ({_slides(v.get('slides'))})"
        source = v.get("source") or ""
        if source == "Linked (file not found)":
            items.append(f"Linked media file not found: {label}. Copy it next to the presentation "
                         f"or embed it (Insert > Video > This Device).")
        elif source == "Linked (online)":
            items.append(f"Online media: {label}. Needs an internet connection while presenting.")
        if v in (r.get("videos") or []):
            codec = v.get("video_codec") or ""
            if codec and not codec.startswith(WIDELY_SUPPORTED_VIDEO):
                items.append(f"Video {label} uses {codec}, which may not play on every computer. "
                             f"Microsoft recommends .mp4 with H.264 video and AAC audio.")
            if (v.get("size_mb") or 0) >= LARGE_VIDEO_MB:
                items.append(f"Large video {label}: {v['size_mb']} MB. Consider "
                             f"File > Compress Media to reduce the file size.")
        if v.get("probe_error") and source == "Embedded":
            items.append(f"Could not measure {label}: {v['probe_error']}.")

    auto = [t for t in r.get("transitions") or [] if t.get("auto_advance")]
    if auto:
        items.append("Slides advance automatically after a set time: "
                     + ", ".join(f"slide {t['slide']} ({t['auto_advance_seconds']}s)" for t in auto) + ".")
    if r.get("hidden_slide_count"):
        hidden = [str(sl["number"]) for sl in r.get("slides") or [] if sl.get("hidden")]
        items.append(f"Hidden slides (won't show in the slide show): {', '.join(hidden)}.")
    for w in r.get("warnings") or []:
        items.append(f"Note: {w}")
    return items


# --------------------------------------------------------------------------- #
# Document
# --------------------------------------------------------------------------- #

def _overview(reports, st):
    rows = []
    for r in reports:
        s = r.get("summary") or {}
        if r.get("error"):
            rows.append([_p(r["file_name"], st["cellb"]), "Error", "—", "—", "—", "—", "—", "—"])
            continue
        vids = f"{s.get('video_count', 0)}" + (f" ({s.get('total_video_size_mb')} MB)" if s.get("video_count") else "")
        rows.append([
            _p(r["file_name"], st["cellb"]),
            str(r.get("slide_count", 0)),
            vids,
            "Yes" if s.get("has_transitions") else "No",
            "Yes" if s.get("has_animations") else "No",
            ", ".join(s.get("addins") or []) or "—",
            ", ".join(s.get("fonts_to_install") or []) or "—",
            str(len(items_to_check(r))),
        ])
    widths = [1.7 * inch, 0.5 * inch, 0.85 * inch, 0.75 * inch, 0.75 * inch, 1.0 * inch, 1.3 * inch, 0.45 * inch]
    return _table(["Presentation", "Slides", "Videos", "Transitions", "Animations", "Add-ins",
                   "Fonts to install", "Items"], rows, widths, st)


def _presentation(r, st) -> list:
    flow: list = [Paragraph(escape(r["file_name"]), st["h1"])]
    s = r.get("summary") or {}
    if r.get("error"):
        flow.append(_p(f"Could not be analysed: {r['error']}", st["err"]))
        return flow

    hidden = f" ({r['hidden_slide_count']} hidden)" if r.get("hidden_slide_count") else ""
    facts = [
        f"{_plural(r.get('slide_count', 0), 'slide')}{hidden}",
        f"{r.get('file_size_mb')} MB",
        f"{_plural(s.get('video_count', 0), 'video')}"
        + (f" ({s.get('total_video_size_mb')} MB, {s.get('total_video_duration')})" if s.get("video_count") else ""),
        _plural(s.get("audio_count", 0), "audio clip"),
        _plural(s.get("image_count", 0), "image"),
        f"{s.get('text_density', '—').lower()} text ({s.get('avg_words_per_slide')} words/slide)",
        f"speaker notes on {s.get('slides_with_notes', 0)} of {r.get('slide_count', 0)} slides",
    ]
    flow.append(_p(" · ".join(facts), st["meta"]))

    # Items to check
    flow.append(Paragraph("Items to check", st["h2"]))
    items = items_to_check(r)
    if items:
        box = [Paragraph(escape(i), st["bullet"], bulletText="•") for i in items]
        t = Table([[box]], colWidths=[7.1 * inch])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), WARN_BG),
                               ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                               ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
        flow.append(t)
    else:
        flow.append(_p("✓ Nothing to check — no non-standard fonts, add-ins or media problems found.", st["ok"]))

    # Videos
    videos = r.get("videos") or []
    flow.append(Paragraph(f"Videos ({len(videos)})", st["h2"]))
    if videos:
        rows = [[_p(v.get("name") or v.get("file"), st["cellb"]), _slides(v.get("slides")).replace("slides ", "").replace("slide ", ""),
                 v.get("size_mb") if v.get("size_mb") is not None else "—", v.get("duration") or "—",
                 v.get("fps") or "—", v.get("resolution") or "—", v.get("codec") or "—", v.get("source")]
                for v in videos]
        flow.append(_table(["Video", "Slides", "MB", "Length", "FPS", "Resolution", "Codec", "Source"], rows,
                           [1.5 * inch, 0.5 * inch, 0.45 * inch, 0.65 * inch, 0.45 * inch, 0.8 * inch,
                            1.4 * inch, 1.55 * inch], st))
    else:
        flow.append(_p("No videos.", st["body"]))

    # Transitions & animations
    trans = [t for t in r.get("transitions") or []]
    flow.append(Paragraph("Transitions & animations", st["h2"]))
    if trans:
        rows = [[str(t["slide"]), t["type"], f"after {t['auto_advance_seconds']}s" if t.get("auto_advance") else "On click"]
                for t in trans]
        flow.append(_table(["Slide", "Transition", "Advances"], rows, [0.6 * inch, 2.4 * inch, 1.4 * inch], st))
    else:
        flow.append(_p("No slide transitions.", st["body"]))
    anims = [a for a in r.get("animations") or [] if a.get("effects")]
    if anims:
        flow.append(Spacer(1, 4))
        flow.append(_p(f"Animations: {_plural(s.get('animation_effect_count', 0), 'effect')} on "
                       f"{_slides([a['slide'] for a in anims])}.", st["body"]))
    else:
        flow.append(_p("No animations.", st["body"]))

    # Add-ins
    addins = [a for a in r.get("addins") or [] if a.get("confidence") != "low"]
    mentions = [a for a in r.get("addins") or [] if a.get("confidence") == "low"]
    flow.append(Paragraph(f"Add-ins & interactive tools ({len(addins)})", st["h2"]))
    if addins:
        rows = [[_p(a["name"], st["cellb"]), ", ".join(a.get("locations") or []) or "—",
                 ", ".join(a.get("sources") or []), a.get("confidence")] for a in addins]
        flow.append(_table(["Add-in", "Where", "Detected via", "Confidence"], rows,
                           [1.6 * inch, 1.6 * inch, 2.9 * inch, 1.0 * inch], st))
    else:
        flow.append(_p("No add-ins detected.", st["body"]))
    if mentions:
        flow.append(_p("Only mentioned in slide text (not an actual add-in): "
                       + ", ".join(f"{a['name']} ({', '.join(a.get('locations') or [])})" for a in mentions) + ".",
                       st["meta"]))

    # Fonts
    fonts = r.get("fonts") or []
    flow.append(Paragraph(f"Fonts ({len(fonts)})", st["h2"]))
    if fonts:
        rows = []
        for f in fonts:
            links = " ".join(
                f'<link href="{escape(l["url"])}" color="#2a4fa8"><u>{escape(l["label"])}</u></link>'
                for l in f.get("links") or []
                if str(l.get("url", "")).startswith("https://") and not f.get("embedded"))
            status_color = "#b42318" if f.get("needs_install") else "#9a6400" if f.get("platform_warning") else "#1f7a4d"
            label = "Embedded in this file" if f.get("embedded") else f["status_label"]
            rows.append([
                _p(f["name"], st["cellb"]),
                Paragraph(f'<font color="{status_color}">{escape(label)}</font>', st["cell"]),
                _slides(f.get("slides")) if f.get("slides") else ", ".join(f.get("uses") or []),
                "Yes" if f.get("embedded") else "No",
                {True: "Yes", False: "No"}.get(f.get("installed_here"), "Unknown"),
                Paragraph(links or "—", st["cell"]),
            ])
        flow.append(_table(["Font", "Status", "Used on", "Embedded", "Installed*", "Download / find"], rows,
                           [1.45 * inch, 1.45 * inch, 1.3 * inch, 0.8 * inch, 0.75 * inch, 1.55 * inch], st))
        flow.append(_p("* Installed on the computer that created this report.", st["meta"]))
    else:
        flow.append(_p("No fonts found.", st["body"]))
    return flow


def build_pdf(reports: list[dict], path: str) -> str:
    """Write the PDF report for the given analysis reports. Returns the path."""
    st = _styles()
    regular, _bold = _register_fonts()
    generated = _dt.datetime.now().strftime("%B %d, %Y at %H:%M").replace(" 0", " ")

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(regular, 7.5)
        canvas.setFillColor(MUTED)
        canvas.drawString(0.6 * inch, 0.45 * inch, f"PPTcheck report · {generated}")
        canvas.drawRightString(letter[0] - 0.6 * inch, 0.45 * inch, f"Page {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(path, pagesize=letter, leftMargin=0.6 * inch, rightMargin=0.6 * inch,
                            topMargin=0.6 * inch, bottomMargin=0.75 * inch,
                            title="PPTcheck report", author="PPTcheck", creator=f"PPTcheck {__version__}")
    flow: list = [
        Paragraph('PPT<font color="#c4442b">check</font> report', st["title"]),
        _p(f"Generated {generated} · {_plural(len(reports), 'presentation')} · PPTcheck v{__version__}. "
           f"Created on this computer; no files were uploaded.", st["meta"]),
        Spacer(1, 10),
    ]
    if len(reports) > 1:
        flow += [Paragraph("Overview", st["h2"]), _overview(reports, st)]
        for r in reports:
            flow += [PageBreak()] + _presentation(r, st)
    else:
        flow += _presentation(reports[0], st)
    doc.build(flow, onFirstPage=footer, onLaterPages=footer)
    return path
