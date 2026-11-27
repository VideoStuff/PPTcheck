"""Media inspection: duration, FPS, resolution and codec of embedded video/audio.

Two engines are supported:
  1. MediaInfo (via the pymediainfo package). The pip package ships the
     MediaInfo library inside it, so nothing extra needs installing. This is
     the default, and it is what the packaged desktop app uses.
  2. ffprobe (part of FFmpeg), used only if MediaInfo is unavailable. A
     bundled ffprobe next to the app (bin/ffprobe[.exe]) is preferred over
     one found on PATH.

Embedded media is streamed out of the .pptx ZIP into a temporary file (not
loaded into memory) and inspected there.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from fractions import Fraction
from typing import Any

from .paths import resource_dir

try:
    from pymediainfo import MediaInfo
    _MEDIAINFO_OK = MediaInfo.can_parse()
except Exception:  # library missing or the shared library failed to load
    MediaInfo = None
    _MEDIAINFO_OK = False


def _find_ffprobe() -> str | None:
    exe = "ffprobe.exe" if sys.platform.startswith("win") else "ffprobe"
    bundled = os.path.join(resource_dir(), "bin", exe)
    if os.path.isfile(bundled):
        return bundled
    return shutil.which("ffprobe")


FFPROBE = _find_ffprobe()

VIDEO_CODEC_NAMES = {
    "avc": "H.264", "h264": "H.264", "hevc": "H.265 / HEVC", "h265": "H.265 / HEVC",
    "mpeg-4 visual": "MPEG-4 Part 2", "mpeg4": "MPEG-4 Part 2", "vc-1": "VC-1 (WMV)",
    "wmv3": "VC-1 (WMV)", "vp8": "VP8", "vp9": "VP9", "av1": "AV1", "prores": "Apple ProRes",
    "mpeg video": "MPEG-1/2", "mpeg2video": "MPEG-2", "mjpeg": "Motion JPEG",
}


def engine_name() -> str | None:
    """Human-readable name of the media engine in use, or None."""
    if _MEDIAINFO_OK:
        return "MediaInfo"
    if FFPROBE:
        return "ffprobe"
    return None


def media_available() -> bool:
    return engine_name() is not None


# Kept for backwards compatibility with earlier versions.
ffprobe_available = media_available


def format_duration(seconds: float | None) -> str | None:
    if seconds is None:
        return None
    total = int(round(seconds))
    h, rem = divmod(total, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _friendly_codec(name: str | None, profile: str | None = None) -> str | None:
    if not name:
        return None
    friendly = VIDEO_CODEC_NAMES.get(name.lower(), name)
    if profile:
        friendly += f" ({profile.split('@')[0]})"
    return friendly


def _empty() -> dict[str, Any]:
    return {"duration_seconds": None, "duration": None, "fps": None, "resolution": None,
            "width": None, "height": None, "video_codec": None, "audio_codec": None,
            "container": None, "bitrate_kbps": None, "probe_error": None}


def _probe_mediainfo(path: str) -> dict[str, Any]:
    result = _empty()
    try:
        info = MediaInfo.parse(path)
    except Exception as exc:  # corrupt or unreadable file
        result["probe_error"] = f"MediaInfo could not read file: {exc}"[:300]
        return result
    general = next((t for t in info.tracks if t.track_type == "General"), None)
    video = next((t for t in info.tracks if t.track_type == "Video"), None)
    audio = next((t for t in info.tracks if t.track_type == "Audio"), None)
    if video is None and audio is None:
        result["probe_error"] = "No audio or video stream found"

    def num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    duration_ms = None
    for track in (general, video, audio):
        if track is not None and num(track.duration) is not None:
            duration_ms = num(track.duration)
            break
    if duration_ms is not None:
        result["duration_seconds"] = round(duration_ms / 1000, 3)
        result["duration"] = format_duration(duration_ms / 1000)
    if general is not None:
        result["container"] = general.format
        br = num(general.overall_bit_rate)
        result["bitrate_kbps"] = round(br / 1000) if br else None
    if video is not None:
        fps = num(video.frame_rate)
        result["fps"] = round(fps, 3) if fps else None
        w, h = num(video.width), num(video.height)
        if w and h:
            result["width"], result["height"] = int(w), int(h)
            result["resolution"] = f"{int(w)}x{int(h)}"
        result["video_codec"] = _friendly_codec(video.format, video.format_profile)
    if audio is not None:
        result["audio_codec"] = audio.format
    return result


def _parse_rate(rate: str | None) -> float | None:
    if not rate or rate in ("0/0", "0"):
        return None
    try:
        value = float(Fraction(rate))
    except (ValueError, ZeroDivisionError):
        return None
    return round(value, 3) if value > 0 else None


def _probe_ffprobe(path: str) -> dict[str, Any]:
    result = _empty()
    try:
        proc = subprocess.run(
            [FFPROBE, "-v", "error", "-print_format", "json", "-show_format", "-show_streams", path],
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
    video = next((s for s in streams if s.get("codec_type") == "video"
                  and not (s.get("disposition") or {}).get("attached_pic")), None)
    audio = next((s for s in streams if s.get("codec_type") == "audio"), None)
    duration = None
    for candidate in (fmt.get("duration"), (video or {}).get("duration"), (audio or {}).get("duration")):
        try:
            if candidate is not None:
                duration = float(candidate)
                break
        except ValueError:
            continue
    result["duration_seconds"] = round(duration, 3) if duration is not None else None
    result["duration"] = format_duration(duration)
    result["container"] = fmt.get("format_long_name") or fmt.get("format_name") or None
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
        result["video_codec"] = _friendly_codec(video.get("codec_name"), video.get("profile"))
    if audio:
        result["audio_codec"] = audio.get("codec_name")
    return result


def probe_file(path: str) -> dict[str, Any]:
    """Return normalised metadata for a media file on disk."""
    if _MEDIAINFO_OK:
        return _probe_mediainfo(path)
    if FFPROBE:
        return _probe_ffprobe(path)
    result = _empty()
    result["probe_error"] = "No media engine available (MediaInfo or FFmpeg)"
    return result


def probe_zip_member(zf: zipfile.ZipFile, member: str) -> dict[str, Any]:
    """Extract one ZIP member to a temp file (streamed, not loaded in RAM) and probe it."""
    if not media_available():
        return probe_file("")
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
