# PPTcheck — PowerPoint technical auditor

PPTcheck runs on your own computer. It opens `.pptx` files (one at a time or whole
folders) and reports:

- **Overview:** file size, slide count, hidden slides, text density, images, audio, speaker notes
- **Videos:** each video's file name, which slides it's on, size (MB), duration (HH:MM:SS),
  frame rate (FPS), resolution, codec and format. Covers both embedded and linked videos.
- **Transitions & animations:** which slides have transitions, the type (Morph, Fade, Push…),
  auto-advance timings, and animation counts per slide
- **Add-ins:** Slido, Mentimeter, Poll Everywhere, Kahoot!, Microsoft Forms and other
  Office Web Add-ins, with the slide numbers where they appear

It then shows a sortable, filterable dashboard that you can export to **CSV** (opens in Excel)
or **JSON**.

Nothing gets uploaded to the internet. The "server" runs only on your computer (`127.0.0.1`)
and accepts no connections from other machines.

---

## How it works

A `.pptx` file is actually a ZIP archive full of XML files. PPTcheck has a small **Python**
program (built with FastAPI) that opens that archive directly. It reads the slide XML for
`<p:transition>` and animation timing, follows each slide's relationship file
(`ppt/slides/_rels/slideN.xml.rels`) to find videos, web add-ins, tags and hyperlinks, and
scans `ppt/webextensions/`. Each embedded video in `ppt/media/` is copied out to a temporary
file and measured with **ffprobe** (part of FFmpeg), which is where the exact duration, FPS and
codec come from. The dashboard is a plain web page that the Python program serves to your browser.

```
Browser dashboard  ──(sends file)──▶  Python server (FastAPI)  ──▶  unzip + read XML (lxml)
   (static/*.html/js)  ◀──(JSON report)──                        └─▶  ffprobe on each video
```

---

## Setup (one-time)

You need two free programs: **Python** (version 3.10 or newer), which runs PPTcheck, and
**FFmpeg**, which measures videos.

### Windows

1. **Install Python:** go to <https://www.python.org/downloads/> and click the big
   "Download Python" button. When the installer opens, **tick "Add python.exe to PATH"** at the
   bottom before you click *Install Now*.
2. **Install FFmpeg:** open the Start menu, type `PowerShell`, open it, and paste:
   ```powershell
   winget install --id Gyan.FFmpeg -e
   ```
   Then **close PowerShell and open a new window** so it picks up the new program.
   (If `winget` isn't available: download the "release essentials" zip from
   <https://www.gyan.dev/ffmpeg/builds/>, unzip it to `C:\ffmpeg`, and add `C:\ffmpeg\bin` to
   your PATH.)
3. Check that both work by typing these in the new PowerShell window. Each one should print a version number:
   ```powershell
   python --version
   ffprobe -version
   ```

### macOS

1. Install Homebrew (a tool installer) if you don't have it. Open the **Terminal** app and paste
   the command shown at <https://brew.sh>.
2. Then run:
   ```bash
   brew install python ffmpeg
   ```

### Linux (Ubuntu/Debian)

```bash
sudo apt update && sudo apt install -y python3 python3-venv ffmpeg
```

### Get the PPTcheck files

Either download the repository as a ZIP from GitHub (green **Code** button → **Download ZIP**,
then unzip it), or if you use git:

```bash
git clone https://github.com/VideoStuff/PPTcheck.git
cd PPTcheck
```

---

## Running it

### Easiest way: double-click

- **Windows:** double-click **`run.bat`** in the PPTcheck folder.
- **macOS/Linux:** in Terminal, `cd` into the PPTcheck folder and run `./run.sh`.

The first run takes a minute because it creates a private Python environment in a `.venv`
folder and installs the libraries listed in `requirements.txt`. That way nothing else on your
computer gets changed. Later runs start in a couple of seconds.

Your browser then opens **http://127.0.0.1:8765**. Leave the black terminal window open while
you use the app, and close it (or press `Ctrl+C`) when you're done.

### Manual way (if you prefer typing the commands)

```bash
python -m venv .venv
# Windows:      .venv\Scripts\activate
# macOS/Linux:  source .venv/bin/activate
pip install -r requirements.txt
python -m pptcheck
```

Options: `python -m pptcheck --port 9000` uses a different port, and `--no-browser` stops it
from opening the browser automatically.

---

## Using the dashboard

1. **Add presentations**, in one of three ways:
   - Drag `.pptx` files **or whole folders** onto the drop area.
   - Click **Choose files…** or **Choose folder…**. Untick *Include subfolders* to scan only the
     top level of a folder.
   - Paste a folder path, such as `C:\Users\Jon\Documents\Decks`, into the box and click
     **Scan folder**. This reads the files straight from your disk without copying them, so it
     is faster for large folders. It is also the **only way linked (not embedded) videos can be
     measured**, because PPTcheck needs to know where the presentation lives to find them.
2. A progress bar shows the scan. Files are processed two at a time in the background, so the page
   stays responsive. **Cancel** stops the scan after the files currently being processed.
3. **Batch summary table:** click any column header to sort, and use the search box and drop-downs to
   filter (for example *Add-ins: Detected*).
4. **Click a row** to open the detailed view, which shows overview stats, the per-video table,
   transitions, animations, add-ins with their slide numbers, and a slide-by-slide breakdown.
5. **Export:** the buttons above the table export the rows currently shown (Summary, Videos,
   Transitions, Add-ins and Slides as CSV, or everything as JSON). The detail view has its own
   **Export JSON / Export CSV** buttons for a single presentation.

Files starting with `~$` are skipped automatically. PowerPoint creates these temporary lock
files while a deck is open. `.pptm`, `.ppsx` and `.ppsm` files are accepted too, since they
use the same format. Old-style `.ppt` files (PowerPoint 97–2003) are **not** supported; open
them in PowerPoint and *Save As* `.pptx` first.

### Command-line mode (no browser)

```bash
python -m pptcheck --scan "C:\path\to\folder" --out report.json
```

---

## How detection works (and its limits)

| Item | Where PPTcheck looks |
|---|---|
| Slide order / hidden slides | `ppt/presentation.xml` and `show="0"` on each `<p:sld>` |
| Videos & audio | `<a:videoFile>`, `<a:audioFile>`, `<p14:media>` in slide XML → slide `.rels` → `ppt/media/*` |
| Video metrics | ffprobe on each extracted media file (duration, avg frame rate, width×height, codecs) |
| Transitions | `<p:transition>`, including Morph (`p159:morph`) and newer effects in `mc:AlternateContent`; `advTm` = auto-advance |
| Animations | `<p:timing>` → `<p:cTn presetClass=…>` (entrance/emphasis/exit/motion path). Video "play" triggers are counted separately. |
| Add-ins | `ppt/webextensions/*` (Office Web Add-ins placed on slides or as task panes), slide tag parts, OLE/ActiveX objects, hyperlinks, custom XML parts, document properties |

**Confidence levels for add-ins:**
- **high:** found inside an add-in, tag or embedded-object part.
- **medium:** found in a hyperlink or document metadata.
- **low:** only mentioned in slide text.

Low-confidence matches appear in the detail view but are left out of the summary table's
"Detected add-ins" column.

**Unrecognised Office Web Add-ins** are still reported, labelled with their Office Store ID
(e.g. `Office Web Add-in (wa104381909)`). To teach PPTcheck a new add-in name, add an entry to
`KNOWN_ADDINS` in `pptcheck/addins.py`.

**Online videos** (YouTube, Stream, etc.) are listed as "Linked (online)". Their duration and
FPS can't be measured because PPTcheck doesn't download anything from the internet.

---

## Project layout

```
PPTcheck/
├── run.bat / run.sh          one-click launchers (create .venv on first run)
├── requirements.txt          Python libraries
├── pptcheck/
│   ├── __main__.py           starts the server (or --scan command-line mode)
│   ├── server.py             FastAPI web API + serves the dashboard
│   ├── analyzer.py           opens the .pptx ZIP and reads the XML
│   ├── media.py              ffprobe video/audio measurement
│   ├── addins.py             list of known add-ins and their fingerprints
│   └── static/               the dashboard (index.html, app.js, styles.css)
└── tests/
    ├── make_sample.py        builds a test deck with a video, Morph, Fade, Slido, …
    └── test_analyzer.py      automated checks
```

## Running the tests

```bash
pip install pytest
python -m pytest tests
```

## Troubleshooting

- **"FFmpeg not found" badge in the top-right:** FFmpeg isn't installed, or the terminal was
  opened before you installed it. Close the terminal, open a new one, and run `ffprobe -version`
  to check. Without FFmpeg everything else still works; only the duration, FPS, resolution and
  codec fields stay empty.
- **`python` is not recognised (Windows):** reinstall Python and tick *Add python.exe to PATH*,
  or use `py` instead of `python`.
- **Port already in use:** run `python -m pptcheck --port 9000` and open http://127.0.0.1:9000.
- **A file shows "Error":** hover over the red tag to see why. Common causes are password-protected
  files, corrupt files, or `.ppt` files that were renamed to `.pptx`.
