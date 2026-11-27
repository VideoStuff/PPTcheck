"""FastAPI web server: serves the dashboard and the analysis API (localhost only)."""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .analyzer import analyze_pptx
from .fonts import preload_installed_fonts
from .media import FFPROBE, ffprobe_available

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")
PPT_EXTENSIONS = (".pptx", ".pptm", ".ppsx", ".ppsm")
ALLOWED_HOSTS = {"127.0.0.1", "localhost", "[::1]"}

app = FastAPI(title="PPTcheck", version="1.1.0")


@app.on_event("startup")
def _startup() -> None:
    # Reading the installed-font list takes a few seconds; start it early.
    preload_installed_fonts()


@app.middleware("http")
async def only_local_hosts(request: Request, call_next):
    # Protects against "DNS rebinding": a web page on the internet pretending
    # to be localhost to read files through this API.
    host = (request.headers.get("host") or "").rsplit(":", 1)[0].lower()
    if host not in ALLOWED_HOSTS:
        return JSONResponse({"detail": "PPTcheck only accepts requests from this computer."},
                            status_code=403)
    return await call_next(request)


def is_presentation(name: str) -> bool:
    base = os.path.basename(name)
    return (base.lower().endswith(PPT_EXTENSIONS)
            and not base.startswith("~$") and not base.startswith("._"))


@app.get("/api/health")
def health():
    version = None
    if FFPROBE:
        try:
            out = subprocess.run([FFPROBE, "-version"], capture_output=True, text=True, timeout=10)
            version = (out.stdout.splitlines() or [""])[0]
        except (OSError, subprocess.TimeoutExpired):
            version = None
    return {"ok": True, "ffprobe": ffprobe_available(), "ffprobe_version": version}


@app.post("/api/analyze")
def analyze_upload(file: UploadFile = File(...), relative_path: str = Form("")):
    """Analyse a presentation uploaded from the browser (drag-and-drop / picker)."""
    name = relative_path or file.filename or "presentation.pptx"
    if not is_presentation(name):
        raise HTTPException(400, f"{name} is not a .pptx presentation")
    fd, tmp = tempfile.mkstemp(prefix="pptcheck_upload_", suffix=".pptx")
    try:
        with os.fdopen(fd, "wb") as out:
            shutil.copyfileobj(file.file, out, length=1024 * 1024)
        return analyze_pptx(tmp, display_name=name, source_dir=None)
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


@app.get("/api/list-folder")
def list_folder(path: str, recursive: bool = True):
    """List presentations inside a folder on this computer (no upload needed)."""
    folder = os.path.abspath(os.path.expanduser(path.strip().strip('"')))
    if not os.path.isdir(folder):
        raise HTTPException(404, f"Folder not found: {folder}")
    files = []
    if recursive:
        for root, dirs, names in os.walk(folder):
            dirs[:] = sorted(d for d in dirs if not d.startswith("."))
            for n in sorted(names):
                if is_presentation(n):
                    full = os.path.join(root, n)
                    files.append(full)
    else:
        for n in sorted(os.listdir(folder)):
            full = os.path.join(folder, n)
            if os.path.isfile(full) and is_presentation(n):
                files.append(full)
    return {
        "folder": folder,
        "files": [{"path": f, "relative_path": os.path.relpath(f, folder),
                   "size_bytes": os.path.getsize(f)} for f in files],
    }


class PathRequest(BaseModel):
    path: str
    display_name: str | None = None


@app.post("/api/analyze-path")
def analyze_path(req: PathRequest):
    """Analyse a presentation directly from disk. Linked videos can be resolved."""
    path = os.path.abspath(os.path.expanduser(req.path))
    if not os.path.isfile(path):
        raise HTTPException(404, f"File not found: {path}")
    if not is_presentation(path):
        raise HTTPException(400, "Only .pptx/.pptm/.ppsx files can be analysed")
    return analyze_pptx(path, display_name=req.display_name or os.path.basename(path),
                        source_dir=os.path.dirname(path))


app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")
