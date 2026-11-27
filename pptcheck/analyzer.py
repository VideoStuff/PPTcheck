"""Core .pptx inspector.

A .pptx file is a ZIP archive of XML "parts". This module opens the archive
directly (zipfile + lxml) instead of relying only on python-pptx, because the
high-level libraries ignore transitions, web add-ins, slide tags and the
media relationships that we need for an accurate audit.

Key parts that are inspected:
  ppt/presentation.xml (+ _rels)        slide order, presentation-level tags
  ppt/slides/slideN.xml                 hidden flag, text, pictures, media,
                                        <p:transition>, <p:timing> (animations)
  ppt/slides/_rels/slideN.xml.rels      links to media, notes, web add-ins,
                                        tags, OLE/ActiveX, hyperlinks
  ppt/media/*                           embedded video/audio -> ffprobe
  ppt/webextensions/*                   Office Web Add-ins (e.g. Slido)
  ppt/tags/*, customXml/*, docProps/*   add-in fingerprints
"""
from __future__ import annotations

import datetime as _dt
import os
import posixpath
import re
import zipfile
from typing import Any
from urllib.parse import unquote, urlparse

from lxml import etree

from .addins import match_addins
from .fonts import classify as classify_font, installed_font_names, is_installed
from .geometry import aspect_label, is_stretched, placement as frame_placement, slide_size as describe_slide_size
from .media import format_duration, probe_file, probe_zip_member

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
    "we": "http://schemas.microsoft.com/office/webextensions/webextension/2010/11",
    "wetp": "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11",
}
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
R_EMBED = f"{{{NS['r']}}}embed"
R_LINK = f"{{{NS['r']}}}link"
R_ID = f"{{{NS['r']}}}id"

VIDEO_EXT = {".mp4", ".m4v", ".mov", ".wmv", ".avi", ".mpg", ".mpeg", ".mkv",
             ".webm", ".asf", ".3gp", ".flv", ".ogv", ".ts", ".mts"}
AUDIO_EXT = {".mp3", ".wav", ".m4a", ".wma", ".aac", ".mid", ".midi", ".ogg",
             ".flac", ".aif", ".aiff", ".au"}
IMAGE_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf",
             ".wmf", ".svg", ".webp", ".jfif", ".ico"}

TRANSITION_NAMES = {
    "fade": "Fade", "push": "Push", "wipe": "Wipe", "split": "Split",
    "cut": "Cut", "dissolve": "Dissolve", "random": "Random",
    "randomBar": "Random Bars", "blinds": "Blinds", "checker": "Checkerboard",
    "circle": "Circle", "comb": "Comb", "cover": "Cover", "diamond": "Diamond",
    "newsflash": "Newsflash", "plus": "Plus", "pull": "Uncover",
    "strips": "Strips", "wedge": "Wedge", "wheel": "Wheel", "zoom": "Zoom",
    "morph": "Morph", "vortex": "Vortex", "switch": "Switch", "flip": "Flip",
    "ripple": "Ripple", "honeycomb": "Honeycomb", "prism": "Prism",
    "doors": "Doors", "window": "Window", "ferris": "Ferris Wheel",
    "gallery": "Gallery", "conveyor": "Conveyor", "pan": "Pan",
    "glitter": "Glitter", "warp": "Warp", "flythrough": "Fly Through",
    "flash": "Flash", "shred": "Shred", "reveal": "Reveal",
    "wheelReverse": "Wheel Reverse",
}
IGNORED_TRANSITION_CHILDREN = {"extLst", "sndAc"}
NOTES_SKIP_PH = {"sldNum", "sldImg", "dt", "ftr", "hdr"}
CONFIDENCE_RANK = {"low": 0, "medium": 1, "high": 2}

_PARSER = etree.XMLParser(resolve_entities=False, no_network=True,
                          huge_tree=True, recover=True)


def _local(el) -> str:
    try:
        return etree.QName(el).localname
    except ValueError:
        return ""


def _camel_to_title(name: str) -> str:
    return re.sub(r"(?<!^)(?=[A-Z])", " ", name).title() if name else name


def _mb(n: int | None) -> float | None:
    return round(n / (1024 * 1024), 2) if n is not None else None


class Package:
    """Thin helper around the ZIP archive: XML parsing and relationship lookup."""

    def __init__(self, zf: zipfile.ZipFile):
        self.zf = zf
        self.names = set(zf.namelist())
        self._xml_cache: dict[str, Any] = {}
        self._rels_cache: dict[str, dict] = {}

    def has(self, part: str) -> bool:
        return part in self.names

    def read_bytes(self, part: str, limit: int | None = None) -> bytes:
        if part not in self.names:
            return b""
        with self.zf.open(part) as fh:
            return fh.read(limit) if limit else fh.read()

    def text(self, part: str, limit: int = 2_000_000) -> str:
        raw = self.read_bytes(part, limit).decode("utf-8", errors="ignore")
        return raw + "\n" + unquote(raw)

    def xml(self, part: str):
        if part not in self._xml_cache:
            data = self.read_bytes(part)
            root = None
            if data:
                try:
                    root = etree.fromstring(data, _PARSER)
                except etree.XMLSyntaxError:
                    root = None
            self._xml_cache[part] = root
        return self._xml_cache[part]

    def rels(self, part: str) -> dict[str, dict]:
        """Return {rId: {type, kind, target, external}} for a part."""
        if part in self._rels_cache:
            return self._rels_cache[part]
        directory, base = posixpath.split(part)
        rels_part = posixpath.join(directory, "_rels", base + ".rels")
        result: dict[str, dict] = {}
        root = self.xml(rels_part)
        if root is not None:
            for rel in root.iter(f"{{{PKG_REL}}}Relationship"):
                rtype = rel.get("Type", "")
                target = rel.get("Target", "")
                external = rel.get("TargetMode") == "External"
                if not external:
                    if target.startswith("/"):
                        target = target.lstrip("/")
                    else:
                        target = posixpath.normpath(posixpath.join(directory, target))
                result[rel.get("Id", "")] = {
                    "type": rtype,
                    "kind": rtype.rsplit("/", 1)[-1],
                    "target": target,
                    "external": external,
                }
        self._rels_cache[part] = result
        return result


# --------------------------------------------------------------------------- #
# Slide-level extraction helpers
# --------------------------------------------------------------------------- #

def _slide_parts(pkg: Package) -> list[str]:
    pres_part = "ppt/presentation.xml"
    pres = pkg.xml(pres_part)
    rels = pkg.rels(pres_part)
    parts: list[str] = []
    if pres is not None:
        for sld in pres.iter(f"{{{NS['p']}}}sldId"):
            rel = rels.get(sld.get(R_ID, ""))
            if rel and pkg.has(rel["target"]):
                parts.append(rel["target"])
    if not parts:  # fallback: numeric order of slide files
        found = [n for n in pkg.names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
        parts = sorted(found, key=lambda n: int(re.search(r"(\d+)\.xml$", n).group(1)))
    return parts


def _text_of(el) -> str:
    return "".join(t.text or "" for t in el.iter(f"{{{NS['a']}}}t"))


def _slide_title(root) -> str | None:
    for sp in root.iter(f"{{{NS['p']}}}sp"):
        ph = sp.find(f".//{{{NS['p']}}}nvPr/{{{NS['p']}}}ph")
        if ph is not None and ph.get("type") in ("title", "ctrTitle"):
            text = " ".join(_text_of(p) for p in sp.iter(f"{{{NS['a']}}}p")).strip()
            if text:
                return re.sub(r"\s+", " ", text)[:200]
    return None


def _word_count(root) -> int:
    words = 0
    for para in root.iter(f"{{{NS['a']}}}p"):
        words += len(_text_of(para).split())
    return words


def _media_from_slide(pkg: Package, root, rels: dict) -> tuple[list[dict], int]:
    """Return (media_refs, image_count) for one slide."""
    media_refs: list[dict] = []
    images = 0
    used_rids: set[str] = set()
    for pic in root.iter(f"{{{NS['p']}}}pic"):
        nvpr = pic.find(f"{{{NS['p']}}}nvPicPr/{{{NS['p']}}}nvPr")
        cnv = pic.find(f"{{{NS['p']}}}nvPicPr/{{{NS['p']}}}cNvPr")
        shape_name = cnv.get("name") if cnv is not None else None
        kind = None
        rids: list[tuple[str, str]] = []  # (rid, element localname)
        if nvpr is not None:
            for el in nvpr.iter():
                ln = _local(el)
                if ln in ("videoFile", "quickTimeFile"):
                    kind = "video"
                elif ln in ("audioFile", "wavAudioFile", "audioCd"):
                    kind = kind or "audio"
                elif ln != "media":
                    continue
                rid = el.get(R_EMBED) or el.get(R_LINK)
                if rid:
                    rids.append((rid, ln))
        if kind is None and rids:  # only p14:media found; decide by extension
            for rid, _ in rids:
                rel = rels.get(rid)
                if rel:
                    ext = posixpath.splitext(urlparse(rel["target"]).path)[1].lower()
                    kind = "audio" if ext in AUDIO_EXT else "video"
                    break
        if kind:
            # Prefer the embedded p14:media part (actual bytes) over the link.
            chosen = None
            for rid, ln in sorted(rids, key=lambda x: 0 if x[1] == "media" else 1):
                rel = rels.get(rid)
                if not rel:
                    continue
                if not rel["external"] and pkg.has(rel["target"]):
                    chosen = rel
                    break
                chosen = chosen or rel
            used_rids.update(r for r, _ in rids)
            if chosen:
                media_refs.append({"kind": kind, "rel": chosen, "shape_name": shape_name,
                                   "frame": _frame_of(pic)})
        else:
            if pic.find(f".//{{{NS['a']}}}blip") is not None:
                images += 1

    # Media relationships that no <p:pic> referenced (rare, but be thorough).
    for rid, rel in rels.items():
        if rid in used_rids or rel["kind"] not in ("video", "audio", "media"):
            continue
        ext = posixpath.splitext(urlparse(rel["target"]).path)[1].lower()
        if rel["kind"] == "audio" or ext in AUDIO_EXT:
            kind = "audio"
        else:
            kind = "video"
        if any(m["rel"]["target"] == rel["target"] for m in media_refs):
            continue
        media_refs.append({"kind": kind, "rel": rel, "shape_name": None, "frame": None})
    return media_refs, images


def _frame_of(pic) -> dict | None:
    """Size of a picture/video frame on the slide (inches), including any cropping."""
    ext = pic.find(f"{{{NS['p']}}}spPr/{{{NS['a']}}}xfrm/{{{NS['a']}}}ext")
    if ext is None:
        return None
    try:
        cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
    except ValueError:
        return None
    crop = {}
    src = pic.find(f"{{{NS['p']}}}blipFill/{{{NS['a']}}}srcRect")
    if src is not None:
        for side in ("l", "t", "r", "b"):
            try:
                crop[side] = int(src.get(side, 0))
            except ValueError:
                pass
    return frame_placement(cx, cy, crop)


def _transition(root) -> dict | None:
    trans = next(root.iter(f"{{{NS['p']}}}transition"), None)
    if trans is None:
        return None
    effect_el = None
    for child in trans:
        if not isinstance(child.tag, str):
            continue
        if _local(child) not in IGNORED_TRANSITION_CHILDREN:
            effect_el = child
            break
    ttype = None
    variant = None
    if effect_el is not None:
        ln = _local(effect_el)
        if ln == "prstTrans":
            prst = effect_el.get("prst") or "preset"
            ttype = _camel_to_title(prst)
        else:
            ttype = TRANSITION_NAMES.get(ln, _camel_to_title(ln))
        variant = (effect_el.get("dir") or effect_el.get("option")
                   or effect_el.get("orient") or effect_el.get("pattern"))
    dur_ms = trans.get(f"{{{NS['p14']}}}dur")
    adv_tm = trans.get("advTm")
    adv_click = trans.get("advClick")
    sound = None
    snd = trans.find(f".//{{{NS['p']}}}snd")
    if snd is not None:
        sound = snd.get("name") or "sound"
    auto_secs = None
    if adv_tm is not None:
        try:
            auto_secs = round(int(adv_tm) / 1000, 2)
        except ValueError:
            auto_secs = None
    if ttype is None and auto_secs is None and sound is None:
        return None
    return {
        "type": ttype or "None (timing only)",
        "has_effect": ttype is not None,
        "variant": variant,
        "speed": trans.get("spd"),
        "duration_ms": int(dur_ms) if dur_ms and dur_ms.isdigit() else None,
        "auto_advance": auto_secs is not None,
        "auto_advance_seconds": auto_secs,
        "advance_on_click": adv_click != "0",
        "sound": sound,
    }


def _animations(root) -> dict:
    counts = {"entrance": 0, "emphasis": 0, "exit": 0, "motion_path": 0,
              "media_triggers": 0, "other": 0}
    timing = root.find(f"{{{NS['p']}}}timing")
    if timing is not None:
        mapping = {"entr": "entrance", "emph": "emphasis", "exit": "exit",
                   "path": "motion_path", "mediacall": "media_triggers"}
        for ctn in timing.iter(f"{{{NS['p']}}}cTn"):
            cls = ctn.get("presetClass")
            if cls:
                counts[mapping.get(cls, "other")] += 1
    counts["effects"] = sum(v for k, v in counts.items() if k != "media_triggers")
    return counts


def _notes_text(pkg: Package, rels: dict) -> str:
    for rel in rels.values():
        if rel["kind"] == "notesSlide" and not rel["external"]:
            root = pkg.xml(rel["target"])
            if root is None:
                return ""
            chunks = []
            for sp in root.iter(f"{{{NS['p']}}}sp"):
                ph = sp.find(f".//{{{NS['p']}}}nvPr/{{{NS['p']}}}ph")
                if ph is not None and ph.get("type") in NOTES_SKIP_PH:
                    continue
                chunks.append(" ".join(_text_of(p) for p in sp.iter(f"{{{NS['a']}}}p")))
            return re.sub(r"\s+", " ", " ".join(chunks)).strip()
    return ""


def _parse_webextension(pkg: Package, part: str) -> dict:
    root = pkg.xml(part)
    info = {"part": part, "store_id": None, "store": None, "store_type": None,
            "version": None, "properties": {}}
    if root is not None:
        ref = root.find(f"{{{NS['we']}}}reference")
        if ref is not None:
            info.update(store_id=ref.get("id"), store=ref.get("store"),
                        store_type=ref.get("storeType"), version=ref.get("version"))
        for prop in root.iter(f"{{{NS['we']}}}property"):
            name = prop.get("name")
            if name:
                info["properties"][name] = (prop.get("value") or "")[:300]
    return info


def _webext_detail(info: dict) -> str:
    bits = []
    if info.get("store_id"):
        bits.append(f"Store ID {info['store_id']}")
    if info.get("store_type"):
        bits.append(f"store type {info['store_type']}")
    if info.get("version"):
        bits.append(f"v{info['version']}")
    return ", ".join(bits) or info["part"]


# --------------------------------------------------------------------------- #
# Fonts
# --------------------------------------------------------------------------- #

FONT_ELEMENTS = {"latin", "ea", "cs", "sym", "buFont"}
FONT_PART_KINDS = {"chart": "Chart", "diagramData": "SmartArt", "diagramDrawing": "SmartArt"}


def _rel_target(pkg: Package, part: str, kind: str) -> str | None:
    for rel in pkg.rels(part).values():
        if rel["kind"] == kind and not rel["external"] and pkg.has(rel["target"]):
            return rel["target"]
    return None


def _theme_fonts(pkg: Package, theme_part: str | None) -> dict[str, str]:
    """Map theme placeholders (mj-lt, mn-lt, mj-ea...) to real font names."""
    result: dict[str, str] = {}
    root = pkg.xml(theme_part) if theme_part else None
    if root is None:
        return result
    scheme = root.find(f".//{{{NS['a']}}}fontScheme")
    if scheme is None:
        return result
    for kind, prefix in (("majorFont", "mj"), ("minorFont", "mn")):
        el = scheme.find(f"{{{NS['a']}}}{kind}")
        if el is None:
            continue
        for tag, suffix in (("latin", "lt"), ("ea", "ea"), ("cs", "cs")):
            t = el.find(f"{{{NS['a']}}}{tag}")
            if t is not None and (t.get("typeface") or "").strip():
                result[f"{prefix}-{suffix}"] = t.get("typeface").strip()
    return result


class FontCollector:
    def __init__(self):
        self.fonts: dict[str, dict] = {}

    def add(self, name: str | None, slide: int | None, use: str):
        name = (name or "").strip().lstrip("@")
        if not name or name.startswith("+"):
            return
        key = name.lower()
        item = self.fonts.setdefault(key, {"name": name, "slides": set(), "uses": set(),
                                           "embedded": False, "embedded_styles": []})
        if slide is not None:
            item["slides"].add(slide)
        item["uses"].add(use)

    def scan(self, root, theme: dict[str, str], slide: int | None, use: str) -> set[str]:
        found = set()
        if root is None:
            return found
        for el in root.iter():
            if not isinstance(el.tag, str) or _local(el) not in FONT_ELEMENTS:
                continue
            face = (el.get("typeface") or "").strip()
            if not face:
                continue
            if face.startswith("+"):
                face = theme.get(face[1:], "")
            if face:
                self.add(face, slide, use)
                found.add(face.lstrip("@"))
        return found

    def result(self, installed: set[str] | None) -> list[dict]:
        out = []
        for item in self.fonts.values():
            info = classify_font(item["name"])
            embedded = item["embedded"]
            needs_install = info["status"] in ("google", "nonstandard") and not embedded
            warning = info["status"] in ("windows", "mac", "cloud") and not embedded
            slides = sorted(item["slides"])
            out.append({
                "name": item["name"],
                "status": info["status"],
                "status_label": info["label"],
                "note": ("Embedded in this file, so it travels with the presentation "
                         "(if the font's license allows embedding). " if embedded else "") + info["note"],
                "family": info["family"],
                "links": info["links"],
                "embedded": embedded,
                "embedded_styles": item["embedded_styles"],
                "installed_here": is_installed(item["name"], installed),
                "needs_install": needs_install,
                "platform_warning": warning,
                "slides": slides,
                "uses": sorted(item["uses"]),
            })
        out.sort(key=lambda f: (not f["needs_install"], not f["platform_warning"], f["name"].lower()))
        return out


# --------------------------------------------------------------------------- #
# Add-in collector
# --------------------------------------------------------------------------- #

class AddinCollector:
    def __init__(self):
        self.items: dict[str, dict] = {}

    def add(self, name: str, category: str, slide: int | None, source: str,
            confidence: str, detail: str | None = None):
        item = self.items.setdefault(name, {
            "name": name, "category": category, "slides": set(),
            "presentation_level": False, "sources": set(),
            "confidence": confidence, "details": [],
        })
        if slide is None:
            item["presentation_level"] = True
        else:
            item["slides"].add(slide)
        item["sources"].add(source)
        if CONFIDENCE_RANK[confidence] > CONFIDENCE_RANK[item["confidence"]]:
            item["confidence"] = confidence
        if detail and detail not in item["details"] and len(item["details"]) < 25:
            item["details"].append(detail)

    def scan(self, text: str, slide: int | None, source: str, confidence: str,
             detail: str | None = None) -> list[str]:
        names = []
        for hit in match_addins(text):
            self.add(hit["name"], hit["category"], slide, source, confidence,
                     detail or f"matched '{hit['keyword']}'")
            names.append(hit["name"])
        return names

    def result(self) -> list[dict]:
        out = []
        for item in self.items.values():
            slides = sorted(item["slides"])
            locations = [f"Slide {s}" for s in slides]
            if item["presentation_level"]:
                locations.append("Presentation-level")
            out.append({**item, "slides": slides, "sources": sorted(item["sources"]),
                        "locations": locations})
        out.sort(key=lambda x: (-CONFIDENCE_RANK[x["confidence"]], x["name"]))
        return out


# --------------------------------------------------------------------------- #
# Main entry point
# --------------------------------------------------------------------------- #

def _resolve_linked_path(target: str, source_dir: str | None) -> str | None:
    """Try to find a linked (external) media file on disk."""
    parsed = urlparse(target)
    if parsed.scheme in ("http", "https"):
        return None
    if parsed.scheme == "file":
        candidate = unquote(parsed.path)
        if re.match(r"^/[A-Za-z]:", candidate):  # file:///C:/...
            candidate = candidate[1:]
    else:
        candidate = unquote(target)
    candidates = [candidate]
    if source_dir and not os.path.isabs(candidate):
        candidates.insert(0, os.path.join(source_dir, candidate))
    if source_dir:  # same file name next to the presentation
        candidates.append(os.path.join(source_dir, os.path.basename(candidate.replace("\\", "/"))))
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


def _empty_report(file_name: str, size: int, error: str) -> dict:
    return {"file_name": file_name, "file_size_bytes": size, "file_size_mb": _mb(size),
            "error": error, "slide_count": 0, "hidden_slide_count": 0,
            "slide_size": None, "summary": {}, "videos": [], "audio": [], "transitions": [],
            "animations": [], "addins": [], "fonts": [], "slides": [], "warnings": [error]}


def analyze_pptx(path: str, display_name: str | None = None,
                 source_dir: str | None = None) -> dict:
    """Analyse one .pptx file on disk and return a JSON-serialisable report.

    display_name: name to show (e.g. relative path inside an uploaded folder)
    source_dir:   directory the original file lives in, used to resolve
                  linked (not embedded) videos. None for browser uploads.
    """
    file_name = display_name or os.path.basename(path)
    size = os.path.getsize(path)
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        return _empty_report(file_name, size,
                             "Not a valid .pptx (ZIP) file — it may be corrupt, "
                             "password-protected, or an old .ppt file renamed.")
    with zf:
        pkg = Package(zf)
        if not pkg.has("ppt/presentation.xml"):
            return _empty_report(file_name, size, "ZIP archive has no ppt/presentation.xml — not a PowerPoint file.")
        return _analyze_package(pkg, file_name, size, source_dir)


def _analyze_package(pkg: Package, file_name: str, size: int, source_dir: str | None) -> dict:
    warnings: list[str] = []
    addins = AddinCollector()
    slide_parts = _slide_parts(pkg)
    slides: list[dict] = []
    transitions: list[dict] = []
    animations: list[dict] = []
    media_groups: dict[tuple, dict] = {}
    referenced_webext: set[str] = set()
    fonts = FontCollector()
    theme_cache: dict[str | None, dict[str, str]] = {}
    scanned_layouts: set[str] = set()

    for number, part in enumerate(slide_parts, start=1):
        root = pkg.xml(part)
        if root is None:
            warnings.append(f"Slide {number}: could not parse {part}")
            continue
        rels = pkg.rels(part)
        hidden = root.get("show") in ("0", "false")
        words = _word_count(root)
        layout_part = _rel_target(pkg, part, "slideLayout")
        master_part = _rel_target(pkg, layout_part, "slideMaster") if layout_part else None
        if master_part not in theme_cache:
            theme_part = _rel_target(pkg, master_part, "theme") if master_part else None
            theme_cache[master_part] = _theme_fonts(pkg, theme_part or "ppt/theme/theme1.xml")
        theme = theme_cache[master_part]
        for p_ in (layout_part, master_part):
            if p_ and p_ not in scanned_layouts:
                scanned_layouts.add(p_)
                fonts.scan(pkg.xml(p_), theme, None, "Slide master / layout")
        slide_fonts = fonts.scan(root, theme, number, "Slide text")
        for rel in rels.values():
            use = FONT_PART_KINDS.get(rel["kind"])
            if use and not rel["external"]:
                slide_fonts |= fonts.scan(pkg.xml(rel["target"]), theme, number, use)
        if words:
            for key, use in (("mj-lt", "Theme heading font"), ("mn-lt", "Theme body font")):
                if theme.get(key):
                    fonts.add(theme[key], number, use)
                    slide_fonts.add(theme[key])
        media_refs, images = _media_from_slide(pkg, root, rels)
        trans = _transition(root)
        anim = _animations(root)
        notes = _notes_text(pkg, rels)
        slide_addins: set[str] = set()
        hyperlinks: list[str] = []
        ole_objects: list[str] = []

        # --- media grouping (same file may appear on several slides) ------
        for ref in media_refs:
            rel = ref["rel"]
            key = (ref["kind"], rel["target"], rel["external"])
            grp = media_groups.setdefault(key, {
                "kind": ref["kind"], "target": rel["target"],
                "external": rel["external"], "slides": [], "shape_names": [], "frames": [],
            })
            if ref.get("frame"):
                grp["frames"].append({"slide": number, **ref["frame"]})
            if number not in grp["slides"]:
                grp["slides"].append(number)
            if ref["shape_name"] and ref["shape_name"] not in grp["shape_names"]:
                grp["shape_names"].append(ref["shape_name"])
            if rel["external"]:
                addins.scan(rel["target"], number, "Linked media URL", "medium",
                            f"Linked media: {rel['target'][:200]}")

        # --- relationships: add-ins, hyperlinks, OLE/ActiveX, tags ----------
        for rel in rels.values():
            kind = rel["kind"]
            if kind == "hyperlink" and rel["external"]:
                hyperlinks.append(rel["target"])
                slide_addins.update(addins.scan(
                    rel["target"], number, "Hyperlink", "medium",
                    f"Hyperlink: {rel['target'][:200]}"))
            elif kind == "webextension" and not rel["external"]:
                referenced_webext.add(rel["target"])
                info = _parse_webextension(pkg, rel["target"])
                names = addins.scan(pkg.text(rel["target"]), number,
                                    "Office Web Add-in", "high", _webext_detail(info))
                if not names:
                    name = f"Office Web Add-in ({info.get('store_id') or posixpath.basename(rel['target'])})"
                    addins.add(name, "Office Web Add-in (unrecognised)", number,
                               "Office Web Add-in", "high", _webext_detail(info))
                    names = [name]
                slide_addins.update(names)
            elif kind == "tags" and not rel["external"]:
                slide_addins.update(addins.scan(pkg.text(rel["target"]), number,
                                                "Slide tags (custom data)", "high",
                                                f"Tag part {rel['target']}"))
            elif kind in ("oleObject", "package", "control", "activeXControlBinary"):
                target = rel["target"]
                label = f"{kind}: {posixpath.basename(target) if not rel['external'] else target}"
                text = target
                if kind == "control" and not rel["external"]:
                    text += pkg.text(target, limit=200_000)
                slide_addins.update(addins.scan(text, number, "Embedded object / ActiveX",
                                                "high", label))
            elif kind in ("customXml", "customData") and not rel["external"]:
                slide_addins.update(addins.scan(pkg.text(target := rel["target"]), number,
                                                "Slide custom data", "high", f"Part {target}"))

        for ole in root.iter(f"{{{NS['p']}}}oleObj"):
            prog = ole.get("progId") or "OLE object"
            label = " / ".join(x for x in (prog, ole.get("name")) if x)
            ole_objects.append(label)
            slide_addins.update(addins.scan(label, number, "Embedded object / ActiveX",
                                            "high", f"OLE {label}"))
        for ctrl in root.iter(f"{{{NS['p']}}}control"):
            label = ctrl.get("name") or "ActiveX control"
            ole_objects.append(f"ActiveX: {label}")
            slide_addins.update(addins.scan(label, number, "Embedded object / ActiveX",
                                            "high", f"ActiveX {label}"))

        # Plain slide text / alt text mentions (lowest confidence).
        slide_raw = etree.tostring(root, encoding="unicode")
        slide_addins.update(addins.scan(slide_raw, number, "Slide text / alt text mention", "low"))

        if trans:
            transitions.append({"slide": number, **trans})
        if anim["effects"] or anim["media_triggers"]:
            animations.append({"slide": number, **anim})

        slides.append({
            "number": number,
            "part": part,
            "title": _slide_title(root),
            "hidden": hidden,
            "word_count": words,
            "image_count": images,
            "video_count": sum(1 for m in media_refs if m["kind"] == "video"),
            "audio_count": sum(1 for m in media_refs if m["kind"] == "audio"),
            "transition": trans["type"] if trans else None,
            "auto_advance_seconds": trans["auto_advance_seconds"] if trans else None,
            "animation_count": anim["effects"],
            "has_notes": bool(notes),
            "notes_preview": notes[:200] if notes else None,
            "hyperlinks": hyperlinks,
            "ole_objects": ole_objects,
            "addins": sorted(slide_addins),
            "fonts": sorted(slide_fonts, key=str.lower),
        })

    # --- embedded fonts (File > Options > Save > Embed fonts) ---------------
    pres_root = pkg.xml("ppt/presentation.xml")
    if pres_root is not None:
        for emb in pres_root.iter(f"{{{NS['p']}}}embeddedFont"):
            font_el = emb.find(f"{{{NS['p']}}}font")
            face = font_el.get("typeface") if font_el is not None else None
            if not face:
                continue
            fonts.add(face, None, "Embedded in file")
            item = fonts.fonts.get(face.strip().lstrip("@").lower())
            if item:
                item["embedded"] = True
                item["embedded_styles"] = [_local(c) for c in emb if isinstance(c.tag, str) and _local(c) != "font"]

    # --- presentation-level add-in scan ------------------------------------
    taskpanes = "ppt/webextensions/taskpanes.xml"
    if pkg.has(taskpanes):
        for rel in pkg.rels(taskpanes).values():
            if rel["kind"] == "webextension" and not rel["external"]:
                referenced_webext.add(rel["target"])
                info = _parse_webextension(pkg, rel["target"])
                if not addins.scan(pkg.text(rel["target"]), None, "Task pane add-in",
                                   "high", "Task pane: " + _webext_detail(info)):
                    addins.add(f"Office Web Add-in ({info.get('store_id') or 'task pane'})",
                               "Office Web Add-in (unrecognised)", None, "Task pane add-in",
                               "high", "Task pane: " + _webext_detail(info))
    for name in sorted(pkg.names):
        if re.fullmatch(r"ppt/webextensions/webextension\d*\.xml", name) and name not in referenced_webext:
            info = _parse_webextension(pkg, name)
            if not addins.scan(pkg.text(name), None, "Office Web Add-in (unplaced)", "high",
                               _webext_detail(info)):
                addins.add(f"Office Web Add-in ({info.get('store_id') or posixpath.basename(name)})",
                           "Office Web Add-in (unrecognised)", None,
                           "Office Web Add-in (unplaced)", "high", _webext_detail(info))
        elif name.startswith("customXml/") and name.endswith(".xml") and "/_rels/" not in name:
            addins.scan(pkg.text(name), None, "Custom XML part", "medium", f"Part {name}")
        elif name in ("docProps/custom.xml", "docProps/app.xml"):
            addins.scan(pkg.text(name), None, "Document properties", "medium", f"Part {name}")
        elif name.startswith("ppt/tags/") and name.endswith(".xml"):
            # Tag parts referenced by slides were already scanned with a slide
            # number; scanning again only adds presentation-level hits for tags
            # attached to the presentation itself.
            pass
    for rel in pkg.rels("ppt/presentation.xml").values():
        if rel["kind"] in ("tags", "customXml") and not rel["external"]:
            addins.scan(pkg.text(rel["target"]), None, "Presentation tags", "high",
                        f"Part {rel['target']}")

    # --- media probing -----------------------------------------------------
    videos, audio = [], []
    for grp in media_groups.values():
        entry = _probe_group(pkg, grp, source_dir, warnings)
        (videos if grp["kind"] == "video" else audio).append(entry)
    videos.sort(key=lambda v: (v["slides"][0] if v["slides"] else 0, v["file"]))
    audio.sort(key=lambda v: (v["slides"][0] if v["slides"] else 0, v["file"]))

    # --- summary -----------------------------------------------------------
    font_list = fonts.result(installed_font_names())
    slide_count = len(slides)
    total_words = sum(s["word_count"] for s in slides)
    avg_words = round(total_words / slide_count, 1) if slide_count else 0
    density = ("Light" if avg_words < 40 else "Moderate" if avg_words <= 100 else "Heavy")
    images_total = sum(s["image_count"] for s in slides)
    image_files = sorted(n for n in pkg.names if n.startswith("ppt/media/")
                         and posixpath.splitext(n)[1].lower() in IMAGE_EXT)
    total_video_bytes = sum(v["size_bytes"] or 0 for v in videos)
    slides_with_notes = sum(1 for s in slides if s["has_notes"])
    addin_list = addins.result()
    trans_types = sorted({t["type"] for t in transitions if t["has_effect"]})
    animated = [a for a in animations if a["effects"]]

    summary = {
        "text_density": density,
        "total_words": total_words,
        "avg_words_per_slide": avg_words,
        "image_count": images_total,
        "unique_image_files": len(image_files),
        "video_count": len(videos),
        "embedded_video_count": sum(1 for v in videos if v["source"] == "Embedded"),
        "linked_video_count": sum(1 for v in videos if v["source"] != "Embedded"),
        "total_video_size_bytes": total_video_bytes,
        "total_video_size_mb": _mb(total_video_bytes),
        "total_video_duration": format_duration(sum(v["duration_seconds"] or 0 for v in videos)) if videos else None,
        "audio_count": len(audio),
        "has_notes": slides_with_notes > 0,
        "slides_with_notes": slides_with_notes,
        "has_transitions": any(t["has_effect"] for t in transitions),
        "transition_slide_count": sum(1 for t in transitions if t["has_effect"]),
        "transition_types": trans_types,
        "has_auto_advance": any(t["auto_advance"] for t in transitions),
        "auto_advance_slide_count": sum(1 for t in transitions if t["auto_advance"]),
        "has_animations": bool(animated),
        "animated_slide_count": len(animated),
        "animation_effect_count": sum(a["effects"] for a in animated),
        "addins": [a["name"] for a in addin_list if a["confidence"] != "low"],
        "addin_mentions": [a["name"] for a in addin_list if a["confidence"] == "low"],
        "external_hyperlink_count": sum(len(s["hyperlinks"]) for s in slides),
        "font_count": len(font_list),
        "fonts": [f["name"] for f in font_list],
        "fonts_to_install": [f["name"] for f in font_list if f["needs_install"]],
        "font_platform_warnings": [f["name"] for f in font_list if f["platform_warning"]],
        "embedded_fonts": [f["name"] for f in font_list if f["embedded"]],
    }
    hidden_count = sum(1 for s in slides if s["hidden"])
    summary["description"] = _describe(slide_count, hidden_count, summary)

    sld_sz = pres_root.find(f"{{{NS['p']}}}sldSz") if pres_root is not None else None
    slide_size = None
    if sld_sz is not None:
        try:
            slide_size = describe_slide_size(int(sld_sz.get("cx", 0)), int(sld_sz.get("cy", 0)), sld_sz.get("type"))
        except ValueError:
            slide_size = None
    summary["slide_size"] = slide_size["text"] if slide_size else None
    summary["stretched_videos"] = [v["file"] for v in videos if v.get("stretched_on_slides")]

    return {
        "file_name": file_name,
        "file_size_bytes": size,
        "file_size_mb": _mb(size),
        "slide_size": slide_size,
        "slide_count": slide_count,
        "hidden_slide_count": hidden_count,
        "summary": summary,
        "videos": videos,
        "audio": audio,
        "transitions": transitions,
        "animations": animations,
        "addins": addin_list,
        "fonts": font_list,
        "slides": slides,
        "warnings": warnings,
        "error": None,
        "analyzed_at": _dt.datetime.now().isoformat(timespec="seconds"),
    }


def _probe_group(pkg: Package, grp: dict, source_dir: str | None, warnings: list[str]) -> dict:
    target = grp["target"]
    parsed_path = urlparse(target).path if grp["external"] else target
    file_label = posixpath.basename(unquote(parsed_path).replace("\\", "/")) or target
    ext = posixpath.splitext(file_label)[1].lower().lstrip(".") or None
    entry = {
        "file": file_label,
        "name": grp["shape_names"][0] if grp["shape_names"] else file_label,
        "source": "Embedded",
        "target": target,
        "slides": sorted(grp["slides"]),
        "format": ext.upper() if ext else None,
        "size_bytes": None,
        "size_mb": None,
    }
    meta: dict[str, Any]
    if not grp["external"]:
        if pkg.has(target):
            entry["size_bytes"] = pkg.zf.getinfo(target).file_size
            meta = probe_zip_member(pkg.zf, target)
        else:
            meta = {"probe_error": f"Media part {target} missing from file"}
            warnings.append(f"Media part {target} referenced but missing")
    else:
        local = _resolve_linked_path(target, source_dir)
        if urlparse(target).scheme in ("http", "https"):
            entry["source"] = "Linked (online)"
            meta = {"probe_error": "Online video — not downloaded"}
        elif local:
            entry["source"] = "Linked (file found)"
            entry["size_bytes"] = os.path.getsize(local)
            meta = probe_file(local)
        else:
            entry["source"] = "Linked (file not found)"
            meta = {"probe_error": "Linked file is not available on this computer"}
    entry["size_mb"] = _mb(entry["size_bytes"])
    for key in ("duration", "duration_seconds", "fps", "resolution", "width", "height",
                "video_codec", "audio_codec", "container", "bitrate_kbps", "probe_error"):
        entry[key] = meta.get(key)
    codec_bits = [c for c in (entry["video_codec"], entry["audio_codec"]) if c]
    entry["codec"] = " / ".join(codec_bits) if codec_bits else None

    # Aspect ratio of the video itself, and how it is placed on each slide.
    ratio = meta.get("aspect_value")
    entry["aspect_value"] = ratio
    entry["aspect_ratio"] = aspect_label(ratio, 1) if ratio else None
    if entry["width"] and entry["height"]:
        rotated = f" (rotated {meta['rotation']}°)" if meta.get("rotation") else ""
        entry["resolution"] = f"{entry['width']} × {entry['height']} px{rotated}"
    placements = []
    for f in grp.get("frames", []):
        stretched = is_stretched(f, ratio) if grp["kind"] == "video" else False
        placements.append({k: v for k, v in f.items() if not k.startswith("_")} | {"stretched": stretched})
    entry["placements"] = placements
    entry["stretched_on_slides"] = sorted({p["slide"] for p in placements if p["stretched"]})
    sizes = []
    for p in placements:
        label = f"{p['text']} ({p['aspect_ratio']})"
        if label not in sizes:
            sizes.append(label)
    entry["size_on_slide"] = "; ".join(sizes) if sizes else None
    return entry


def _plural(n: int, word: str) -> str:
    return f"{n} {word}{'' if n == 1 else 's'}"


def _describe(slide_count: int, hidden: int, s: dict) -> str:
    parts = [f"{slide_count} slide{'s' if slide_count != 1 else ''}"
             + (f" ({hidden} hidden)" if hidden else "")]
    parts.append(f"{s['text_density'].lower()} text density (avg {s['avg_words_per_slide']} words/slide)")
    parts.append(f"{s['image_count']} image{'s' if s['image_count'] != 1 else ''}")
    if s["video_count"]:
        parts.append(f"{s['video_count']} video{'s' if s['video_count'] != 1 else ''} "
                     f"({s['total_video_size_mb']} MB embedded/linked)")
    else:
        parts.append("no videos")
    parts.append(f"{s['audio_count']} audio clip{'s' if s['audio_count'] != 1 else ''}")
    parts.append(f"speaker notes on {s['slides_with_notes']} of {slide_count} slides"
                 if s["has_notes"] else "no speaker notes")
    if s["has_transitions"]:
        parts.append(f"transitions on {_plural(s['transition_slide_count'], 'slide')} ({', '.join(s['transition_types'])})")
    else:
        parts.append("no transitions")
    if s["has_auto_advance"]:
        parts.append(f"auto-advance timings on {_plural(s['auto_advance_slide_count'], 'slide')}")
    parts.append(f"animations on {_plural(s['animated_slide_count'], 'slide')}" if s["has_animations"] else "no animations")
    parts.append("add-ins: " + ", ".join(s["addins"]) if s["addins"] else "no add-ins detected")
    font_text = _plural(s["font_count"], "font")
    if s["fonts_to_install"]:
        font_text += f" ({len(s['fonts_to_install'])} not standard: {', '.join(s['fonts_to_install'])})"
    else:
        font_text += " (all standard or embedded)"
    parts.append(font_text)
    return "; ".join(parts) + "."
