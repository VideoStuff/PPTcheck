# PyInstaller recipe for the PPTcheck desktop app.
# Build with:  pyinstaller packaging/pptcheck.spec --noconfirm
#   Windows -> dist/PPTcheck.exe          (single file)
#   macOS   -> dist/PPTcheck.app          (app bundle)
#   Linux   -> dist/PPTcheck              (single file)
import glob
import os
import sys

import pymediainfo
from PyInstaller.utils.hooks import collect_submodules

ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))
IS_MAC = sys.platform == "darwin"

hidden = collect_submodules("fontTools.ttLib.tables")   # font tables are imported on demand
# The window: pywebview picks its platform module at run time.
if sys.platform.startswith("win"):
    hidden += ["webview.platforms.winforms", "webview.platforms.edgechromium", "clr_loader", "pythonnet"]
elif IS_MAC:
    hidden += ["webview.platforms.cocoa"]

# The MediaInfo library ships inside the pymediainfo package (MediaInfo.dll /
# libmediainfo.0.dylib / libmediainfo.so.0). Add it explicitly so it is always
# bundled next to pymediainfo, where pymediainfo looks for it.
_mi_dir = os.path.dirname(pymediainfo.__file__)
mediainfo_libs = [(f, "pymediainfo") for f in glob.glob(os.path.join(_mi_dir, "*"))
                  if os.path.basename(f).lower().startswith(("mediainfo", "libmediainfo"))]
if not mediainfo_libs:
    raise SystemExit("MediaInfo library not found inside pymediainfo - run: pip install pymediainfo")

a = Analysis(
    [os.path.join(ROOT, "packaging", "launcher.py")],
    pathex=[ROOT],
    binaries=mediainfo_libs,
    datas=[(os.path.join(ROOT, "pptcheck", "static"), os.path.join("pptcheck", "static")),
           (os.path.join(ROOT, "pptcheck", "assets"), os.path.join("pptcheck", "assets"))],
    hiddenimports=hidden,
    excludes=["tkinter", "pptx", "numpy", "pytest", "IPython", "PyQt5", "PyQt6", "PySide2", "PySide6",
              "gi", "cefpython3", "fastapi", "uvicorn"],
    noarchive=False,
)
pyz = PYZ(a.pure)

icon = None
for candidate in ("pptcheck.ico", "pptcheck.icns"):
    path = os.path.join(ROOT, "packaging", candidate)
    if os.path.exists(path) and (candidate.endswith(".icns") == IS_MAC):
        icon = path

if IS_MAC:
    # A .app bundle: a folder-style build wrapped as a normal Mac application.
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name="PPTcheck",
              console=False, icon=icon, upx=False)
    coll = COLLECT(exe, a.binaries, a.datas, name="PPTcheck", upx=False)
    app = BUNDLE(
        coll,
        name="PPTcheck.app",
        icon=icon,
        bundle_identifier="io.github.pptcheck",
        info_plist={
            "CFBundleName": "PPTcheck",
            "CFBundleDisplayName": "PPTcheck",
            "CFBundleShortVersionString": os.environ.get("PPTCHECK_VERSION", "1.5.0").lstrip("v"),
            "NSHighResolutionCapable": True,
            "NSRequiresAquaSystemAppearance": False,   # follow light/dark mode
        },
    )
else:
    # One self-contained executable file, no console window.
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name="PPTcheck",
              console=False, icon=icon, upx=False, runtime_tmpdir=None)
