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
- **Fonts:** every font used, where it's used, and whether it's standard. Fonts that need
  installing come with a download link (free Google Fonts) or search links (other fonts).
  It also shows whether each font is embedded in the file or installed on your computer.

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
file and measured with **MediaInfo** (built in), which is where the exact duration, FPS and codec
come from. The dashboard is a plain web page that the Python program serves to your browser.

```
Browser dashboard  ──(sends file)──▶  Python server (FastAPI)  ──▶  unzip + read XML (lxml)
   (static/*.html/js)  ◀──(JSON report)──                        └─▶  MediaInfo on each video
```

---

## Download and run (no installation needed)

PPTcheck is packaged as a standalone app. Python, the video-analysis engine and everything
else it needs are built in, so people don't need to install anything.

1. Go to the repository's **Releases** page on GitHub and download the zip for your computer:
   - **Windows:** `PPTcheck-Windows.zip`
   - **Mac with Apple chip (M1 and newer):** `PPTcheck-macOS-AppleSilicon.zip`
   - **Older Intel Mac:** `PPTcheck-macOS-Intel.zip`
2. Unzip it. On a Mac, drag `PPTcheck.app` into your Applications folder.
3. Double-click **PPTcheck**. Your browser opens the dashboard.
4. When you're finished, click **Quit PPTcheck** at the top-right of the dashboard. It also
   stops by itself about 15 minutes after you close the browser tab.

**First launch security prompts:** the app isn't signed with a paid Apple/Microsoft developer
certificate, so each computer asks for confirmation once:
- **Windows** ("Windows protected your PC"): click **More info**, then **Run anyway**.
- **Mac** ("cannot be opened" / "Apple could not verify"): click **Done**, then open
  **System Settings → Privacy & Security**, scroll down and click **Open Anyway**.
  If it says the app "is damaged", run `xattr -cr /Applications/PPTcheck.app` in Terminal.

The same instructions are in `READ ME FIRST.txt` inside each zip.

---

## For developers: building the downloadable apps

The apps are built automatically by **GitHub Actions**: GitHub's own Windows and Mac machines
run the steps in `.github/workflows/build.yml`.

- **Every push** that changes the app runs the tests and builds the apps on Windows and macOS
  (Apple Silicon and Intel). It also opens each built app and checks it against a test
  presentation (`packaging/smoke_test.py`). The zips are attached to that run: open the
  repository's **Actions** tab, click the run, then go to **Artifacts**. Downloading these
  requires a GitHub login.
- **Publishing a release:** create a version tag, e.g. `v1.2.0`. On GitHub, go to
  **Releases → Draft a new release**, type the tag `v1.2.0`, and publish. Or, with git, run
  `git tag v1.2.0 && git push origin v1.2.0`. The workflow then builds the apps and attaches
  the zips to the release. People can download them from the Releases page without a GitHub
  account, as long as the repository is public.

To build on your own computer instead (this produces an app for that computer's operating
system only):

```bash
pip install -r requirements.txt pyinstaller
pyinstaller packaging/pptcheck.spec --noconfirm
# result: dist/PPTcheck.exe (Windows), dist/PPTcheck.app (Mac), dist/PPTcheck (Linux)
```

### Running from source (for development)

Requires Python 3.10+. FFmpeg is optional: video details come from the MediaInfo library,
which installs automatically with `pymediainfo`.

- **Windows:** double-click `run.bat`. **macOS/Linux:** run `./run.sh`.
  The first run creates a private Python environment in a `.venv` folder.
  If you ran an older version before, delete `.venv` so the new libraries get installed.
- Or manually:

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
   Transitions, Add-ins, Fonts and Slides as CSV, or everything as JSON). The detail view has its own
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
| Video metrics | MediaInfo (or ffprobe, if MediaInfo is unavailable) on each extracted media file: duration, frame rate, width×height, codecs |
| Transitions | `<p:transition>`, including Morph (`p159:morph`) and newer effects in `mc:AlternateContent`; `advTm` = auto-advance |
| Animations | `<p:timing>` → `<p:cTn presetClass=…>` (entrance/emphasis/exit/motion path). Video "play" triggers are counted separately. |
| Fonts | `<a:latin>`, `<a:ea>`, `<a:cs>`, `<a:sym>`, `<a:buFont>` in slides, charts, SmartArt, masters and layouts; theme `<a:fontScheme>`; `<p:embeddedFontLst>` |
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

### Fonts

PPTcheck reads fonts from slide text, bullets, charts, SmartArt, slide masters/layouts, the
theme (the "Headings" and "Body" fonts), and the list of fonts embedded in the file. Each font
gets one of these statuses:

| Status | Meaning | Action |
|---|---|---|
| Standard (Windows & Mac) | Arial, Times New Roman, Verdana… | None |
| Included with Microsoft Office | Calibri, Cambria, Century Gothic… | None if Office is installed |
| Microsoft 365 cloud font | Aptos, Grandview, Arial Nova… | Office downloads these automatically when online; older or offline Office versions substitute them |
| Windows only / Mac only | Segoe UI / Helvetica Neue, Avenir… | Check the deck on the other platform |
| Free — Google Fonts | Montserrat, Open Sans, Roboto… | Click **Download from Google Fonts** and install the font |
| Not standard — must be installed | Gotham, Proxima Nova, company fonts… | Use the search links to find it. It may be a paid font, so check the license. |

A font that is **embedded in the file** doesn't need installing, as long as its license allows
embedding. **"Installed here"** shows whether the font is installed on the computer running
PPTcheck, which helps you check a presenting laptop. Style words are ignored when matching, so
"Montserrat SemiBold" is matched as "Montserrat". To change how a font is classified, add its
name to one of the lists at the top of `pptcheck/fonts.py`.

**Online videos** (YouTube, Stream, etc.) are listed as "Linked (online)". Their duration and
FPS can't be measured because PPTcheck doesn't download anything from the internet.

---

## Project layout

```
PPTcheck/
├── run.bat / run.sh          run from source (developers)
├── packaging/                desktop-app build recipe, icon, smoke test, READ ME FIRST.txt
├── .github/workflows/        automatic Windows/Mac builds on GitHub
├── requirements.txt          Python libraries
├── pptcheck/
│   ├── __main__.py           starts the server (or --scan command-line mode)
│   ├── server.py             FastAPI web API + serves the dashboard
│   ├── analyzer.py           opens the .pptx ZIP and reads the XML
│   ├── media.py              video/audio measurement (MediaInfo, ffprobe fallback)
│   ├── paths.py              finds bundled files when running as a packaged app
│   ├── addins.py             list of known add-ins and their fingerprints
│   ├── fonts.py              font lists (standard/Office/Google…), download links, installed-font check
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

- **"Video engine missing" badge in the top-right:** this only happens when running from source
  if `pymediainfo` didn't install. Run `pip install -r requirements.txt` again. Everything else
  still works without it; only the duration, FPS, resolution and codec fields stay empty.
- **"PPTcheck has stopped running" banner:** the app was quit, or it closed itself after being
  idle. Open PPTcheck again.
- **`python` is not recognised (Windows):** reinstall Python and tick *Add python.exe to PATH*,
  or use `py` instead of `python`.
- **Port already in use:** run `python -m pptcheck --port 9000` and open http://127.0.0.1:9000.
- **A file shows "Error":** hover over the red tag to see why. Common causes are password-protected
  files, corrupt files, or `.ppt` files that were renamed to `.pptx`.
