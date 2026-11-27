#!/usr/bin/env bash
# Start PPTcheck on macOS / Linux. First run creates a private Python environment.
set -e
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "First run: setting up (this takes a minute)..."
  python3 -m venv .venv
  ./.venv/bin/python -m pip install --upgrade pip >/dev/null
  ./.venv/bin/python -m pip install -r requirements.txt
fi
command -v ffprobe >/dev/null || echo "WARNING: ffprobe not found. Install FFmpeg for video duration/FPS (see README)."
exec ./.venv/bin/python -m pptcheck "$@"
