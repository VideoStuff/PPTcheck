# PPTcheck — PowerPoint technical auditor

PPTcheck runs on your own computer. It opens `.pptx` files (one at a time or whole
folders) and reports:

- **Overview:** file size, slide count, hidden slides, **slide size** (in inches, PowerPoint's
  units, plus pixels at 96 px/inch and the aspect ratio, e.g. `13.33 × 7.5 in (1280 × 720 px) · 16:9`),
  text density, images, audio, speaker notes
- **Videos:** each video's file name, which slides it's on, size (MB), duration (HH:MM:SS),
  frame rate (FPS), resolution in pixels, aspect ratio, codec and format. Covers both embedded
  and linked videos.
- **Transitions & animations:** which slides have transitions, the type (Morph, Fade, Push…),
  auto-advance timings, and animation counts per slide
- **Add-ins:** Slido, Mentimeter, Poll Everywhere, Kahoot!, Microsoft Forms and other
  Office Web Add-ins, with the slide numbers where they appear
- **Fonts:** every font used, where it's used, and whether it's standard. Fonts that need
  installing come with a download link (free Google Fonts) or search links (other fonts).
  It also shows whether each font is embedded in the file or installed on your computer.

It then shows a sortable, filterable dashboard, and can save a **PDF summary report** of the
findings.

**Privacy:** PPTcheck is a desktop app with its own window. It reads presentations directly
from your disk and never uploads or sends them anywhere. It has no web server and makes no
network connections. The page shown in the window is also locked down so it *can't* make
network requests. The only time the internet is used is when you click a font download/search
link, which opens your normal web browser.

---

## How it works

A `.pptx` file is actually a ZIP archive full of XML files. PPTcheck is a **Python** program
that opens that archive directly. It reads the slide XML for
`<p:transition>` and animation timing, follows each slide's relationship file
(`ppt/slides/_rels/slideN.xml.rels`) to find videos, web add-ins, tags and hyperlinks, and
scans `ppt/webextensions/`. Each embedded video in `ppt/media/` is copied out to a temporary
file and measured with **MediaInfo** (built in), which is where the exact duration, FPS and codec
come from. The interface is shown in a native window using **pywebview**, which uses the web
engine built into the operating system (Edge WebView2 on Windows, WebKit on macOS). The window
passes the file's *location* (not its contents) to the Python code through pywebview's
in-process bridge.

```
App window (static/*.html/js)  ──(file location)──▶  Python (pptcheck/desktop.py)  ──▶  read .pptx from disk (lxml)
                               ◀──(report)─────────                                 └─▶  MediaInfo on each video
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
3. Double-click **PPTcheck**. It opens in its own window.
4. When you're finished, close the window.

On Windows, the window uses Microsoft's **WebView2** component, which is part of Windows 11 and
almost all Windows 10 computers. If it's missing, PPTcheck says so and opens Microsoft's free
download page.

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
  On Windows and macOS the check also opens the **real app window**, analyses the test deck
  through the interface, and saves a screenshot (artifact `selftest-…`).
- **Publishing a release:** go to **Actions → Build desktop apps → Run workflow** and type the
  version (e.g. `v1.3.1`) in *Publish a release*. Or push a version tag with git:
  `git tag v1.3.1 && git push origin v1.3.1`. The workflow then builds the apps and attaches
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

Requires Python 3.10+ on Windows or macOS. (On Linux, pywebview also needs GTK or Qt; see the
pywebview documentation.) FFmpeg is optional: video details come from the MediaInfo library,
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

---

## Using the dashboard

1. **Add presentations:** drag `.pptx` files **or whole folders** onto the window, or click
   **Choose files…** / **Choose folder…**. Untick *Include subfolders* to scan only the top level
   of a folder. Files are read straight from disk, so **linked** (not embedded) videos stored
   next to the presentation are found and measured too.
2. A progress bar shows the scan. Files are processed two at a time in the background, so the
   window stays responsive. **Cancel** stops the scan after the files currently being processed.
3. **Batch summary table:** click any column header to sort, and use the search box and drop-downs to
   filter (for example *Add-ins: Detected*).
4. **Click a row** to open the detailed view, which shows overview stats, the per-video table,
   transitions, animations, add-ins with their slide numbers, and a slide-by-slide breakdown.
5. **PDF report:** **Export PDF report** (above the table) creates a PDF covering the
   presentations currently shown, so filters apply. The detail view's **Export PDF** button does
   the same for just that presentation. A normal *Save as* window asks where to save, and the
   PDF opens when it's ready. Each presentation gets an **Items to check** list (non-standard
   fonts, add-ins, missing or online videos, video formats that may not play everywhere,
   auto-advance timings, hidden slides), followed by tables of its videos, transitions, add-ins
   and fonts, with clickable font download links.

Files starting with `~$` are skipped automatically. PowerPoint creates these temporary lock
files while a deck is open. `.pptm`, `.ppsx` and `.ppsm` files are accepted too, since they
use the same format. Old-style `.ppt` files (PowerPoint 97–2003) are **not** supported; open
them in PowerPoint and *Save As* `.pptx` first.

### Command-line mode (no window)

```bash
python -m pptcheck --scan "C:\path\to\folder" --pdf report.pdf    # PDF report
python -m pptcheck --scan "C:\path\to\folder" --out report.json  # raw data as JSON
```

---

## How detection works (and its limits)

| Item | Where PPTcheck looks |
|---|---|
| Slide order / hidden slides | `ppt/presentation.xml` and `show="0"` on each `<p:sld>` |
| Videos & audio | `<a:videoFile>`, `<a:audioFile>`, `<p14:media>` in slide XML → slide `.rels` → `ppt/media/*` |
| Slide size | `<p:sldSz cx cy>` in `ppt/presentation.xml` (EMU: 914,400 per inch) |
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
│   ├── __main__.py           starts the app (or --scan command-line mode)
│   ├── desktop.py            the app window and the functions the interface can call
│   ├── files.py              finding presentations in chosen/dropped folders
│   ├── analyzer.py           opens the .pptx ZIP and reads the XML
│   ├── media.py              video/audio measurement (MediaInfo, ffprobe fallback)
│   ├── paths.py              finds bundled files when running as a packaged app
│   ├── addins.py             list of known add-ins and their fingerprints
│   ├── report_pdf.py         the PDF summary report ("Items to check" + tables)
│   ├── assets/fonts/         DejaVu Sans font used in the PDF (free license included)
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
- **"PPTcheck needs the Microsoft Edge WebView2 Runtime" (Windows):** install the free
  "Evergreen Bootstrapper" from the page that opens, then start PPTcheck again.
- **Dropping files does nothing:** use **Choose files…** / **Choose folder…** instead, and let us
  know which Windows/macOS version you're on.
- **`python` is not recognised (Windows, running from source):** reinstall Python and tick
  *Add python.exe to PATH*, or use `py` instead of `python`.
- **A file shows "Error":** hover over the red tag to see why. Common causes are password-protected
  files, corrupt files, or `.ppt` files that were renamed to `.pptx`.
