"""Build a sample .pptx with known contents for testing PPTcheck.

Usage:  python tests/make_sample.py OUTPUT_DIR
Requires FFmpeg (to make a short test video) and python-pptx.
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import zipfile

from pptx import Presentation
from pptx.util import Inches

WEBEXT_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{11111111-2222-3333-4444-555555555555}">
  <we:reference id="wa000000001" version="1.0.0.0" store="en-US" storeType="OMEX"/>
  <we:alternateReferences/>
  <we:properties>
    <we:property name="pollUrl" value="&quot;https://app.sli.do/event/abc123&quot;"/>
  </we:properties>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>"""

GRAPHIC_FRAME = """<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:nvGraphicFramePr><p:cNvPr id="99" name="Slido poll"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="1000000"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/webextensions/webextension/2010/11">
<we:webextensionref xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" r:id="rIdWE1"/>
</a:graphicData></a:graphic></p:graphicFrame>"""

MORPH = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
         '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
         '<p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="2000">'
         '<p159:morph option="byObject"/></p:transition></mc:Choice>'
         '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
FADE_AUTO = '<p:transition spd="med" advClick="0" advTm="5000"><p:fade/></p:transition>'
ANIM = ('<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        '<p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:childTnLst><p:par><p:cTn id="4" fill="hold">'
        '<p:childTnLst><p:par><p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">'
        '<p:childTnLst/></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>'
        '</p:childTnLst></p:cTn><p:prevCondLst/><p:nextCondLst/></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>')


def build(out_dir: str) -> str:
    os.makedirs(out_dir, exist_ok=True)
    tmp = tempfile.mkdtemp()
    video = os.path.join(tmp, "clip.mp4")
    subprocess.run(["ffmpeg", "-y", "-v", "error", "-f", "lavfi", "-i", "testsrc=size=1280x720:rate=30:duration=5",
                    "-f", "lavfi", "-i", "sine=frequency=440:duration=5", "-shortest",
                    "-c:v", "libx264", "-pix_fmt", "yuv420p", "-c:a", "aac", video], check=True)
    poster = os.path.join(tmp, "poster.png")
    subprocess.run(["ffmpeg", "-y", "-v", "error", "-i", video, "-frames:v", "1", poster], check=True)

    prs = Presentation()
    layout = prs.slide_layouts[1]
    s1 = prs.slides.add_slide(layout)
    s1.shapes.title.text = "Welcome"
    s1.placeholders[1].text = "Join the poll at https://www.menti.com/abc"
    s1.notes_slide.notes_text_frame.text = "Remember to greet everyone."
    s1.shapes.add_picture(poster, Inches(5), Inches(4), Inches(2))

    s2 = prs.slides.add_slide(layout)
    s2.shapes.title.text = "A video"
    s2.shapes.add_movie(video, Inches(1), Inches(2), Inches(4), Inches(3), poster_frame_image=poster, mime_type="video/mp4")
    link_run = s2.placeholders[1].text_frame.paragraphs[0].add_run()
    link_run.text = "Mentimeter"
    link_run.hyperlink.address = "https://www.mentimeter.com/app/presentation/xyz"

    s3 = prs.slides.add_slide(layout)
    s3.shapes.title.text = "Interactive poll"
    for run_text, font in (("Google font", "Montserrat SemiBold"), ("Commercial font", "Gotham Rounded"),
                           ("Windows font", "Segoe UI")):
        r = s3.placeholders[1].text_frame.add_paragraph().add_run()
        r.text = run_text
        r.font.name = font

    s4 = prs.slides.add_slide(layout)
    s4.shapes.title.text = "Backup (hidden)"

    out = os.path.join(out_dir, "sample_deck.pptx")
    raw = os.path.join(tmp, "raw.pptx")
    prs.save(raw)

    with zipfile.ZipFile(raw) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            name = item.filename
            if name == "[Content_Types].xml":
                data = data.replace(b"</Types>", b'<Override PartName="/ppt/webextensions/webextension1.xml" '
                                    b'ContentType="application/vnd.ms-office.webextension+xml"/></Types>')
            elif name == "ppt/slides/slide1.xml":
                data = data.decode().replace("</p:clrMapOvr>", "</p:clrMapOvr>" + MORPH).encode()
                data = data.replace(b"</p:sld>", ANIM.encode() + b"</p:sld>")
            elif name == "ppt/slides/slide2.xml":
                data = re.sub(r"(</p:clrMapOvr>)", r"\1" + FADE_AUTO, data.decode()).encode()
            elif name == "ppt/slides/slide3.xml":
                data = data.decode().replace("</p:spTree>", GRAPHIC_FRAME + "</p:spTree>").encode()
            elif name == "ppt/slides/_rels/slide3.xml.rels":
                data = data.replace(b"</Relationships>", b'<Relationship Id="rIdWE1" Type="http://schemas.microsoft.com/'
                                    b'office/2011/relationships/webextension" Target="../webextensions/webextension1.xml"/>'
                                    b"</Relationships>")
            elif name == "ppt/presentation.xml":
                # Declare an embedded font (as PowerPoint's "Embed fonts in the file" does).
                data = data.replace(b"<p:defaultTextStyle>", b'<p:embeddedFontLst><p:embeddedFont>'
                                    b'<p:font typeface="Brand Sans"/><p:regular r:id="rIdF1"/></p:embeddedFont>'
                                    b"</p:embeddedFontLst><p:defaultTextStyle>", 1)
            elif name == "ppt/slides/slide4.xml":
                data = data.replace(b"<p:sld ", b'<p:sld show="0" ', 1)
            zout.writestr(item, data)
        zout.writestr("ppt/webextensions/webextension1.xml", WEBEXT_XML)
    return out


if __name__ == "__main__":
    print(build(sys.argv[1] if len(sys.argv) > 1 else "."))
