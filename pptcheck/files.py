"""Finding presentation files on disk."""
from __future__ import annotations

import os

PPT_EXTENSIONS = (".pptx", ".pptm", ".ppsx", ".ppsm")


def is_presentation(name: str) -> bool:
    """True for .pptx-family files, ignoring PowerPoint's "~$" lock files and macOS "._" files."""
    base = os.path.basename(name)
    return (base.lower().endswith(PPT_EXTENSIONS)
            and not base.startswith("~$") and not base.startswith("._"))


def list_folder(folder: str, recursive: bool = True) -> list[str]:
    """All presentations in a folder (and its subfolders if recursive), sorted."""
    found: list[str] = []
    if recursive:
        for root, dirs, names in os.walk(folder):
            dirs[:] = sorted(d for d in dirs if not d.startswith("."))
            found += [os.path.join(root, n) for n in sorted(names) if is_presentation(n)]
    else:
        for n in sorted(os.listdir(folder)):
            full = os.path.join(folder, n)
            if os.path.isfile(full) and is_presentation(n):
                found.append(full)
    return found


def expand_paths(paths: list[str], recursive: bool = True) -> dict:
    """Turn chosen/dropped files and folders into a list of analysis jobs.

    Returns {"jobs": [{"path", "name"}], "errors": [...]}. Files inside a folder
    are named "Folder/sub/file.pptx" so the dashboard shows where they came from.
    """
    jobs: list[dict] = []
    errors: list[str] = []
    seen: set[str] = set()
    for raw in paths or []:
        path = os.path.abspath(raw)
        if os.path.isdir(path):
            parent = os.path.dirname(path.rstrip(os.sep)) or path
            try:
                for f in list_folder(path, recursive):
                    if f not in seen:
                        seen.add(f)
                        jobs.append({"path": f, "name": os.path.relpath(f, parent).replace(os.sep, "/")})
            except OSError as exc:
                errors.append(f"Could not read folder {os.path.basename(path)}: {exc.strerror or exc}")
        elif os.path.isfile(path):
            if is_presentation(path) and path not in seen:
                seen.add(path)
                jobs.append({"path": path, "name": os.path.basename(path)})
        else:
            errors.append(f"Not found: {raw}")
    # Two different files with the same name (from different folders) need
    # different labels, or the second would replace the first in the table.
    counts: dict[str, int] = {}
    for job in jobs:
        n = counts[job["name"]] = counts.get(job["name"], 0) + 1
        if n > 1:
            job["name"] = f"{job['name']} ({n})"
    return {"jobs": jobs, "errors": errors}
