@echo off
REM Start PPTcheck on Windows. First run creates a private Python environment.
cd /d "%~dp0"
if not exist .venv (
  echo First run: setting up ^(this takes a minute^)...
  py -3 -m venv .venv || python -m venv .venv
  .venv\Scripts\python -m pip install --upgrade pip
  .venv\Scripts\python -m pip install -r requirements.txt
)
where ffprobe >nul 2>nul || echo WARNING: ffprobe not found. Install FFmpeg for video duration/FPS (see README).
.venv\Scripts\python -m pptcheck %*
pause
