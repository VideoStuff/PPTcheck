"""Locate bundled files both when running from source and as a packaged app.

PyInstaller unpacks data files to a folder exposed as sys._MEIPASS. When running
from source, files live next to this package.
"""
from __future__ import annotations

import os
import sys


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def resource_dir() -> str:
    """Folder that contains the 'pptcheck' resources (static/, bin/)."""
    if is_frozen():
        return os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(sys.executable)), "pptcheck")
    return os.path.dirname(os.path.abspath(__file__))
