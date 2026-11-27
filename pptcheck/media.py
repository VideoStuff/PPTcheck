"""Media inspection via ffprobe (part of FFmpeg).

Embedded media is streamed out of the .pptx ZIP into a temporary file and
probed with ffprobe, which reads the real container/stream headers. This is
what gives accurate duration, FPS, resolution and codec values.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import zipfile
from fractions import Fraction
from typing import Any

FFPROBE = shutil.which("ffprobe")


def ffprobe_available() -> bool:
    return FFPROBE is not None


def format_duration(seconds: float | None) -> str | None:
    if seconds is None:
        return None
    total = int(round(seconds))
    h, rem = divmod(total, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _parse_rate(rate: str | None) -> float | None:
    if not rate or rate in ("0/0", "0"):
        return None
    try:
        value = float(Fraction(rate))
    except (ValueError, ZeroDivisionError):
        return None
    return round(value, 3) if value > 0 else None


def probe_file(path: str) -> dict[str, Any]:
    """Run ffprobe on a file on disk and return normalised metadata."""
    result: dict[str, Any] = {
        "duration_seconds": None,
        "duration": None,
        "fps": None,
        "resolution": None,
        "width": None,
        "height": None,
        "video_codec": None,
        "audio_codec": None,
        "container": None,
        "bitrate_kbps": None,
        "probe_error": None,
    }
    if not FFPROBE:
        result["probe_error"] = "ffprobe not installed (install FFmpeg)"
        return result
    try:
        proc = subprocess.run(
            [FFPROBE, "-v", "error", "-print_format", "json",
             "-show_format", "-show_streams", path],
            capture_output=True, text=True, timeout=120,
        )
    except subprocess.TimeoutExpired:
        result["probe_error"] = "ffprobe timed out"
        return result
    if proc.returncode != 0:
        result["probe_error"] = (proc.stderr.strip() or "ffprobe failed")[:300]
        return result
    try:
        data = json.loads(proc.stdout or "{}")
    except json.JSONDecodeError:
        result["probe_error"] = "Could not parse ffprobe output"
        return result

    fmt = data.get("format", {}) or {}
    streams = data.get("streams", []) or []
    # Ignore cover-art "video" streams (attached pictures) when picking video.
    video = next((s for s in streams if s.get("codec_type") == "video"
                  and not (s.get("disposition") or {}).get("attached_pic")), None)
    audio = next((s for s in streams if s.get("codec_type") == "audio"), None)

    duration = None
    for candidate in (fmt.get("duration"), (video or {}).get("duration"),
                      (audio or {}).get("duration")):
        try:
            if candidate is not None:
                duration = float(candidate)
                break
        except ValueError:
            continue
    result["duration_seconds"] = round(duration, 3) if duration is not None else None
    result["duration"] = format_duration(duration)

    names = (fmt.get("format_long_name") or fmt.get("format_name") or "")
    result["container"] = names or None
    try:
        if fmt.get("bit_rate"):
            result["bitrate_kbps"] = round(int(fmt["bit_rate"]) / 1000)
    except ValueError:
        pass

    if video:
        result["fps"] = _parse_rate(video.get("avg_frame_rate")) or _parse_rate(video.get("r_frame_rate"))
        w, h = video.get("width"), video.get("height")
        if w and h:
            result["width"], result["height"] = w, h
            result["resolution"] = f"{w}x{h}"
        codec = video.get("codec_name")
        if codec and video.get("profile"):
            codec = f"{codec} ({video['profile']})"
        result["video_codec"] = codec
    if audio:
        result["audio_codec"] = audio.get("codec_name")
    return result


def probe_zip_member(zf: zipfile.ZipFile, member: str) -> dict[str, Any]:
    """Extract one ZIP member to a temp file (streamed, not loaded in RAM) and probe it."""
    if not FFPROBE:
        return probe_file("")  # returns the "not installed" result
    suffix = os.path.splitext(member)[1] or ".bin"
    fd, tmp_path = tempfile.mkstemp(prefix="pptcheck_", suffix=suffix)
    try:
        with os.fdopen(fd, "wb") as out, zf.open(member) as src:
            shutil.copyfileobj(src, out, length=1024 * 1024)
        return probe_file(tmp_path)
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass
