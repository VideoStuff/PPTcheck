"""PPTcheck desktop window.

The dashboard (static/index.html + app.js + styles.css) is shown inside a
native app window using pywebview, which uses the web engine built into the
operating system (Edge WebView2 on Windows, WebKit on macOS).

There is no web server: the page is loaded directly into the window, and it
calls the Api methods below through pywebview's in-process bridge. Files are
read straight from disk by their location; nothing is uploaded or sent over
any network. The page itself is also locked down by a Content-Security-Policy
that forbids network requests.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import threading
import webbrowser

from . import __version__
from .analyzer import analyze_pptx
from .files import expand_paths, is_presentation
from .media import engine_name, media_available
from .paths import resource_dir

WINDOW_TITLE = "PPTcheck"
WEBVIEW2_URL = "https://developer.microsoft.com/microsoft-edge/webview2/"
# pywebview's webview.FileDialog values (kept here so Api doesn't need to import pywebview).
OPEN_DIALOG, FOLDER_DIALOG, SAVE_DIALOG = 10, 20, 30


# --------------------------------------------------------------------------- #
# Page assembly
# --------------------------------------------------------------------------- #

def build_page_html() -> str:
    """index.html with its stylesheet and script inlined (one self-contained page)."""
    static = os.path.join(resource_dir(), "static")

    def read(name: str) -> str:
        with open(os.path.join(static, name), encoding="utf-8") as fh:
            return fh.read()

    html = read("index.html")
    html = html.replace('<link rel="stylesheet" href="styles.css">', "<style>\n" + read("styles.css") + "\n</style>")
    html = html.replace('<script src="app.js"></script>', "<script>\n" + read("app.js") + "\n</script>")
    return html


# --------------------------------------------------------------------------- #
# Functions the page can call (window.pywebview.api.*)
# --------------------------------------------------------------------------- #

class Api:
    """Public methods are callable from JavaScript. Names starting with "_"
    are hidden from the page by pywebview."""

    def __init__(self, selftest_out: str | None = None):
        self._window = None
        self._selftest_out = selftest_out
        self._selftest_result: dict | None = None
        self._selftest_event = threading.Event()

    # -- information ---------------------------------------------------------
    def info(self) -> dict:
        return {"version": __version__, "media_engine": engine_name(),
                "media_available": media_available()}

    # -- choosing files --------------------------------------------------------
    def choose_files(self) -> list[str]:
        result = self._window.create_file_dialog(
            OPEN_DIALOG, allow_multiple=True,
            file_types=("PowerPoint presentations (*.pptx;*.pptm;*.ppsx;*.ppsm)", "All files (*.*)"))
        return [p for p in (result or []) if is_presentation(p)]

    def choose_folder(self) -> list[str]:
        result = self._window.create_file_dialog(FOLDER_DIALOG)
        if isinstance(result, str):
            result = [result]
        return list(result or [])

    def expand_paths(self, paths: list[str], recursive: bool = True) -> dict:
        return expand_paths(paths, bool(recursive))

    # -- analysis ------------------------------------------------------------
    def analyze(self, path: str, display_name: str | None = None) -> dict:
        try:
            if not os.path.isfile(path) or not is_presentation(path):
                return {"bridge_error": f"Not a PowerPoint file: {path}"}
            return analyze_pptx(path, display_name=display_name or os.path.basename(path),
                                source_dir=os.path.dirname(path))
        except Exception as exc:  # report, don't crash the window
            return {"bridge_error": f"{type(exc).__name__}: {exc}"}

    # -- exporting -------------------------------------------------------------
    def save_file(self, suggested_name: str, content: str) -> dict:
        ext = os.path.splitext(suggested_name)[1].lower()
        types = (("CSV file (*.csv)",) if ext == ".csv" else ("JSON file (*.json)",)) + ("All files (*.*)",)
        target = self._window.create_file_dialog(SAVE_DIALOG, save_filename=suggested_name,
                                                 file_types=types)
        if isinstance(target, (list, tuple)):
            target = target[0] if target else None
        if not target:
            return {"saved": None}
        if ext and not target.lower().endswith(ext):
            target += ext
        try:
            # utf-8 (with the BOM the CSVs already start with) opens correctly in Excel.
            with open(target, "w", encoding="utf-8", newline="") as fh:
                fh.write(content)
        except OSError as exc:
            return {"error": exc.strerror or str(exc)}
        return {"saved": target}

    # -- links -------------------------------------------------------------------
    def open_url(self, url: str) -> bool:
        """Open a font download/search link in the normal web browser."""
        if isinstance(url, str) and url.startswith("https://"):
            webbrowser.open(url)
            return True
        return False

    # -- automated self-test (used by the build) ---------------------------------
    def selftest_done(self, result: dict) -> None:
        self._selftest_result = result
        self._selftest_event.set()


# --------------------------------------------------------------------------- #
# Windows: make sure the modern web engine (WebView2) is present
# --------------------------------------------------------------------------- #

def _windows_has_webview2() -> bool:
    import winreg
    key = r"{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    places = [
        (winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{key}"),
        (winreg.HKEY_LOCAL_MACHINE, rf"SOFTWARE\Microsoft\EdgeUpdate\Clients\{key}"),
        (winreg.HKEY_CURRENT_USER, rf"SOFTWARE\Microsoft\EdgeUpdate\Clients\{key}"),
    ]
    for hive, path in places:
        try:
            with winreg.OpenKey(hive, path) as k:
                version, _ = winreg.QueryValueEx(k, "pv")
                if version and version != "0.0.0.0":
                    return True
        except OSError:
            continue
    return False


def _windows_message(title: str, text: str) -> None:
    import ctypes
    ctypes.windll.user32.MessageBoxW(None, text, title, 0x40)  # MB_ICONINFORMATION


# --------------------------------------------------------------------------- #
# Self-test helpers
# --------------------------------------------------------------------------- #

def _screenshot(path: str) -> None:
    """Best-effort screenshot of the whole screen (build machines only)."""
    try:
        if sys.platform == "darwin":
            subprocess.run(["screencapture", "-x", path], timeout=30, check=False)
        elif sys.platform.startswith("win"):
            ps = (
                "Add-Type -AssemblyName System.Windows.Forms,System.Drawing;"
                "$b=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds;"
                "$i=New-Object System.Drawing.Bitmap $b.Width,$b.Height;"
                "$g=[System.Drawing.Graphics]::FromImage($i);"
                "$g.CopyFromScreen($b.Location,[System.Drawing.Point]::Empty,$b.Size);"
                f"$i.Save('{path}',[System.Drawing.Imaging.ImageFormat]::Png)"
            )
            subprocess.run(["powershell", "-NoProfile", "-Command", ps], timeout=30, check=False,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:
        pass


def _write_json(path: str, data: dict) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def run(selftest_out: str | None = None, selftest_file: str | None = None) -> int:
    """Open the PPTcheck window. Returns a process exit code."""
    if sys.platform.startswith("win") and not _windows_has_webview2():
        msg = ("PPTcheck needs the Microsoft Edge WebView2 Runtime, a free Microsoft component "
               "that is already included in Windows 11 and most Windows 10 computers.\n\n"
               "Click OK to open Microsoft's download page. Install the \"Evergreen Bootstrapper\", "
               "then open PPTcheck again.")
        if selftest_out:
            _write_json(selftest_out, {"ok": False, "error": "WebView2 runtime not installed"})
            return 3
        _windows_message("PPTcheck", msg)
        webbrowser.open(WEBVIEW2_URL)
        return 3

    import webview
    from webview.dom import DOMEventHandler

    api = Api(selftest_out)
    window = webview.create_window(
        WINDOW_TITLE, html=build_page_html(), js_api=api,
        width=1280, height=860, min_size=(820, 560), text_select=True,
    )
    api._window = window

    def on_drop(event):
        files = (event.get("dataTransfer") or {}).get("files") or []
        paths = [f.get("pywebviewFullPath") for f in files if f.get("pywebviewFullPath")]
        if paths:
            window.evaluate_js(f"window.PPTcheck.dropped({json.dumps(paths)})")
        elif files:
            window.evaluate_js("window.PPTcheck && document.querySelector('#toast') && "
                               "(document.querySelector('#toast').textContent='Could not read the dropped "
                               "items. Use Choose files… or Choose folder… instead.', "
                               "document.querySelector('#toast').className='toast bad')")

    def on_started():
        # Called on a background thread once the window exists.
        window.events.loaded.wait(60)
        window.dom.document.events.drop += DOMEventHandler(on_drop, prevent_default=True)

        if selftest_out:
            result: dict = {"ok": False}
            try:
                window.evaluate_js(f"window.PPTcheck.selfTest({json.dumps(selftest_file)})")
                if api._selftest_event.wait(120):
                    page = api._selftest_result or {}
                    result = {"ok": True, "page": page}
                    shot = os.path.splitext(selftest_out)[0] + ".png"
                    threading.Event().wait(2)  # let the detail panel finish drawing
                    _screenshot(shot)
                else:
                    result = {"ok": False, "error": "Page did not finish the self-test in time"}
            except Exception as exc:
                result = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
            _write_json(selftest_out, result)
            window.destroy()

    gui = "edgechromium" if sys.platform.startswith("win") else None
    storage = os.path.join(tempfile.gettempdir(), "pptcheck-webview")
    webview.start(on_started, gui=gui, private_mode=True, storage_path=storage)
    return 0
