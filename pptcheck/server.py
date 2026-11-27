"""FastAPI web server: serves the dashboard and the analysis API (localhost only)."""
from __future__ import annotations

import os
import shutil
import tempfile
import time
from contextlib import asynccontextmanager

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from .analyzer import analyze_pptx
from . import __version__
from .fonts import preload_installed_fonts
from .media import engine_name, media_available
from .paths import resource_dir

STATIC_DIR = os.path.join(resource_dir(), "static")
PPT_EXTENSIONS = (".pptx", ".pptm", ".ppsx", ".ppsm")
ALLOWED_HOSTS = {"127.0.0.1", "localhost", "[::1]"}
APP_ID = "pptcheck"

# Updated whenever an open dashboard tab checks in (used for idle shutdown).
last_heartbeat = time.monotonic()
# Set by __main__ to the running uvicorn.Server so /api/quit can stop it.
server_handle = None


@asynccontextmanager
async def lifespan(_app):
    # Reading the installed-font list takes a few seconds; start it early.
    preload_installed_fonts()
    yield


app = FastAPI(title="PPTcheck", version=__version__, lifespan=lifespan)


def _hostname(value: str) -> str:
    value = value.split("://", 1)[-1]
    if value.startswith("["):
        return value.split("]", 1)[0] + "]"
    return value.rsplit(":", 1)[0].lower()


@app.middleware("http")
async def only_local_requests(request: Request, call_next):
    # 1. Host check: blocks "DNS rebinding" (a web page on the internet
    #    pretending to be localhost to read files through this API).
    # 2. Origin check: blocks other websites open in the browser from sending
    #    requests to PPTcheck.
    host = _hostname(request.headers.get("host") or "")
    origin = request.headers.get("origin")
    if host not in ALLOWED_HOSTS or (origin and _hostname(origin) not in ALLOWED_HOSTS):
        return JSONResponse({"detail": "PPTcheck only accepts requests from this computer."},
                            status_code=403)
    return await call_next(request)


def is_presentation(name: str) -> bool:
    base = os.path.basename(name)
    return (base.lower().endswith(PPT_EXTENSIONS)
            and not base.startswith("~$") and not base.startswith("._"))


@app.get("/api/health")
def health():
    return {"ok": True, "app": APP_ID, "version": __version__,
            "media_engine": engine_name(), "media_available": media_available(),
            "can_quit": server_handle is not None}


@app.get("/api/heartbeat")
def heartbeat():
    global last_heartbeat
    last_heartbeat = time.monotonic()
    return {"ok": True}


@app.post("/api/quit")
def quit_app(request: Request):
    # The custom header can't be sent by other websites without permission.
    if request.headers.get("x-pptcheck") != "1":
        raise HTTPException(403, "Missing header")
    if server_handle is None:
        raise HTTPException(400, "Not running as the desktop app")
    server_handle.should_exit = True
    return {"ok": True}


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
