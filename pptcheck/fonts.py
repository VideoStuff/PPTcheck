"""Font classification: is a font standard, and where can it be downloaded?

Each font found in a presentation is classified as one of:

  standard     Installed on both Windows and macOS (Arial, Times New Roman...)
  office       Installed with Microsoft Office (Calibri, Cambria, Century Gothic...)
  cloud        Microsoft 365 cloud font: Office downloads it automatically when
               it's used (Aptos, Grandview, Arial Nova...). Older/offline Office
               versions will substitute it.
  windows      Ships with Windows only: may be substituted on a Mac
  mac          Ships with macOS only: may be substituted on Windows
  google       Free on Google Fonts: direct download link provided
  nonstandard  Not a standard font: must be installed (link to search for it)

The lists below cover common fonts. Add names to them to change how a font
is classified. Matching ignores case and trailing style words such as
"Light", "SemiBold" or "Italic" (so "Montserrat SemiBold" -> "Montserrat").

The module can also check which fonts are installed on the computer running
PPTcheck (installed_font_names), using fontTools to read font files.
"""
from __future__ import annotations

import os
import re
import sys
import threading
from urllib.parse import quote, quote_plus

STANDARD = {
    "Arial", "Arial Black", "Arial Narrow", "Comic Sans MS", "Courier New", "Georgia",
    "Impact", "Times New Roman", "Trebuchet MS", "Verdana", "Tahoma", "Webdings",
    "Wingdings", "Wingdings 2", "Wingdings 3", "Symbol", "Brush Script MT",
    "Microsoft Sans Serif", "Arial Unicode MS", "Andale Mono", "Courier",
}

OFFICE = {
    "Calibri", "Calibri Light", "Cambria", "Cambria Math", "Candara", "Consolas",
    "Constantia", "Corbel", "Century Gothic", "Gill Sans MT", "Gill Sans MT Condensed",
    "Gill Sans Ultra Bold", "Franklin Gothic Book", "Franklin Gothic Medium",
    "Franklin Gothic Demi", "Franklin Gothic Heavy", "Garamond", "Book Antiqua",
    "Bookman Old Style", "Century", "Century Schoolbook", "Lucida Sans", "Lucida Bright",
    "Lucida Calligraphy", "Lucida Fax", "Lucida Handwriting", "Tw Cen MT",
    "Tw Cen MT Condensed", "Rockwell", "Rockwell Condensed", "Rockwell Extra Bold",
    "Bodoni MT", "Bernard MT Condensed", "Goudy Old Style", "Baskerville Old Face",
    "Perpetua", "Perpetua Titling MT", "Haettenschweiler", "Harrington", "Agency FB",
    "Algerian", "Arial Rounded MT Bold", "Bauhaus 93", "Berlin Sans FB", "Britannic Bold",
    "Broadway", "Castellar", "Chiller", "Colonna MT", "Cooper Black",
    "Copperplate Gothic Bold", "Copperplate Gothic Light", "Curlz MT",
    "Edwardian Script ITC", "Elephant", "Engravers MT", "Eras Bold ITC", "Eras Demi ITC",
    "Eras Light ITC", "Eras Medium ITC", "Felix Titling", "Footlight MT Light", "Forte",
    "Freestyle Script", "French Script MT", "Gigi", "Gloucester MT Extra Condensed",
    "Goudy Stout", "High Tower Text", "Imprint MT Shadow", "Informal Roman", "Jokerman",
    "Juice ITC", "Kristen ITC", "Kunstler Script", "Magneto", "Maiandra GD",
    "Matura MT Script Capitals", "Mistral", "Modern No. 20", "Monotype Corsiva",
    "Niagara Engraved", "Niagara Solid", "Old English Text MT", "Onyx", "Palace Script MT",
    "Papyrus", "Parchment", "Playbill", "Poor Richard", "Pristina", "Rage Italic", "Ravie",
    "Script MT Bold", "Showcard Gothic", "Snap ITC", "Stencil", "Tempus Sans ITC",
    "Viner Hand ITC", "Vivaldi", "Vladimir Script", "Wide Latin", "Calisto MT",
    "Californian FB", "Centaur", "Bell MT", "Blackadder ITC", "Bradley Hand ITC",
    "Lucida Sans Typewriter", "MS Reference Sans Serif", "MS Reference Specialty",
    "MT Extra", "Bookshelf Symbol 7", "Garamond Premier Pro",
}

CLOUD = {
    "Aptos", "Aptos Display", "Aptos Narrow", "Aptos Black", "Aptos ExtraBold",
    "Aptos Light", "Aptos SemiBold", "Aptos Serif", "Aptos Mono", "Grandview",
    "Grandview Display", "Seaford", "Seaford Display", "Skeena", "Skeena Display",
    "Tenorite", "Tenorite Display", "Bierstadt", "Bierstadt Display", "Arial Nova",
    "Arial Nova Cond", "Georgia Pro", "Georgia Pro Cond", "Verdana Pro", "Verdana Pro Cond",
    "Gill Sans Nova", "Rockwell Nova", "Neue Haas Grotesk Text Pro", "Avenir Next LT Pro",
    "Univers", "Univers Condensed", "Trade Gothic Next", "Sagona", "Walbaum Display",
    "Walbaum Text", "Posterama", "Amasis MT Pro", "Avenir Next LT Pro Demi",
}

WINDOWS = {
    "Segoe UI", "Segoe UI Light", "Segoe UI Semilight", "Segoe UI Semibold",
    "Segoe UI Black", "Segoe UI Emoji", "Segoe UI Symbol", "Segoe UI Historic",
    "Segoe UI Variable", "Segoe UI Variable Display", "Segoe UI Variable Text",
    "Segoe Print", "Segoe Script", "Segoe MDL2 Assets", "Segoe Fluent Icons",
    "Bahnschrift", "Cascadia Code", "Cascadia Mono", "Ebrima", "Gabriola", "Gadugi",
    "Ink Free", "Javanese Text", "Leelawadee", "Leelawadee UI", "Lucida Console",
    "Lucida Sans Unicode", "Malgun Gothic", "Marlett", "Microsoft Himalaya",
    "Microsoft JhengHei", "Microsoft New Tai Lue", "Microsoft PhagsPa",
    "Microsoft Tai Le", "Microsoft YaHei", "Microsoft Yi Baiti", "MingLiU-ExtB",
    "Mongolian Baiti", "MS Gothic", "MS PGothic", "MS UI Gothic", "MS Mincho",
    "MS PMincho", "MV Boli", "Myanmar Text", "Nirmala UI", "Palatino Linotype",
    "SimSun", "SimSun-ExtB", "NSimSun", "SimHei", "Sitka", "Sitka Text",
    "Sitka Heading", "Sitka Display", "Sylfaen", "Yu Gothic", "Yu Gothic UI",
    "Yu Mincho", "HoloLens MDL2 Assets", "Meiryo", "Meiryo UI", "Batang", "Gulim",
    "Dotum", "Gungsuh", "Mangal", "Nyala", "Plantagenet Cherokee", "Raavi", "Shruti",
    "Tunga", "Vrinda", "Kartika", "Latha", "Gautami", "Kalinga", "Euphemia",
    "Estrangelo Edessa", "Iskoola Pota", "DaunPenh", "DokChampa", "Khmer UI", "Lao UI",
    "Aharoni", "David", "FrankRuehl", "Miriam", "Narkisim", "Rod", "Levenim MT",
    "Traditional Arabic", "Simplified Arabic", "Arabic Typesetting", "Andalus",
    "Sakkal Majalla", "Urdu Typesetting", "Franklin Gothic", "Lucida Sans Unicode",
}

MAC = {
    "Helvetica", "Helvetica Neue", "Avenir", "Avenir Next", "Avenir Next Condensed",
    "Futura", "Gill Sans", "Optima", "Baskerville", "Didot", "Hoefler Text", "Menlo",
    "Monaco", "SF Pro", "SF Pro Display", "SF Pro Text", "SF Mono", "SF Compact",
    "New York", ".SF NS", ".SF NS Text", ".SF NS Display", "San Francisco",
    "Lucida Grande", "Geneva", "American Typewriter", "Chalkboard", "Chalkboard SE",
    "Chalkduster", "Marker Felt", "Noteworthy", "Copperplate", "Skia", "Big Caslon",
    "Charter", "Iowan Old Style", "Palatino", "Seravek", "Superclarendon", "Athelas",
    "Kefa", "Avenir Book", "Avenir Light", "Avenir Medium", "Avenir Heavy",
    "Avenir Black", "Bodoni 72", "Bodoni 72 Oldstyle", "Bodoni 72 Smallcaps",
    "Bradley Hand", "Cochin", "Snell Roundhand", "Zapfino", "Trattatello", "Phosphate",
    "Herculanum", "Party LET", "Savoye LET", "SignPainter", "Apple Chancery",
    "Apple SD Gothic Neo", "Hiragino Sans", "Hiragino Kaku Gothic Pro",
    "Hiragino Mincho Pro", "PingFang SC", "PingFang TC", "Heiti SC", "Songti SC",
    "Apple Color Emoji", "Arial Hebrew", "Geeza Pro", "Kailasa", "Thonburi", "Rockwell",
    "Courier", "Times", "Zapf Dingbats", "Al Nile", "Damascus", "Farah", "Mishafi",
}

# Fonts available free on Google Fonts (family name -> exact Google family).
GOOGLE = {name: name for name in [
    "Roboto", "Roboto Condensed", "Roboto Slab", "Roboto Mono", "Roboto Flex",
    "Open Sans", "Lato", "Montserrat", "Montserrat Alternates", "Poppins", "Inter",
    "Raleway", "Nunito", "Nunito Sans", "Merriweather", "Merriweather Sans",
    "Playfair Display", "Oswald", "Noto Sans", "Noto Serif", "Noto Sans JP",
    "Noto Sans KR", "Noto Sans SC", "Noto Sans TC", "Noto Sans Arabic",
    "Noto Sans Hebrew", "PT Sans", "PT Serif", "PT Mono", "PT Sans Narrow", "Work Sans",
    "Rubik", "Fira Sans", "Fira Sans Condensed", "Fira Code", "Fira Mono", "Barlow",
    "Barlow Condensed", "Barlow Semi Condensed", "Quicksand", "DM Sans", "DM Serif Display",
    "DM Serif Text", "DM Mono", "Mulish", "Karla", "Libre Baskerville", "Libre Franklin",
    "Lora", "Bebas Neue", "Manrope", "IBM Plex Sans", "IBM Plex Serif", "IBM Plex Mono",
    "IBM Plex Sans Condensed", "Josefin Sans", "Josefin Slab", "Archivo", "Archivo Black",
    "Archivo Narrow", "Cabin", "Titillium Web", "Ubuntu", "Ubuntu Mono", "Ubuntu Condensed",
    "Arimo", "Tinos", "Cousine", "Carlito", "Caladea", "Heebo", "Hind", "Hind Siliguri",
    "Kanit", "Exo", "Exo 2", "Oxygen", "Anton", "Abril Fatface", "Pacifico",
    "Dancing Script", "Lobster", "Caveat", "Comfortaa", "Righteous", "Space Grotesk",
    "Space Mono", "Source Sans 3", "Source Serif 4", "Source Code Pro", "Crimson Text",
    "Crimson Pro", "EB Garamond", "Cormorant", "Cormorant Garamond", "Spectral",
    "Zilla Slab", "Arvo", "Bitter", "Varela Round", "Varela", "Asap", "Asap Condensed",
    "Maven Pro", "Signika", "Teko", "Prompt", "Sarabun", "Outfit", "Lexend",
    "Plus Jakarta Sans", "Urbanist", "Sora", "Figtree", "Red Hat Display", "Red Hat Text",
    "Public Sans", "Alegreya", "Alegreya Sans", "Questrial", "Overpass", "Catamaran",
    "Assistant", "Nanum Gothic", "Nanum Myeongjo", "Saira", "Chivo", "Encode Sans", "Jost",
    "Fjalla One", "Shadows Into Light", "Indie Flower", "Amatic SC", "Permanent Marker",
    "Satisfy", "Great Vibes", "Prata", "Old Standard TT", "Domine", "Cardo", "Vollkorn",
    "Gelasio", "Literata", "Atkinson Hyperlegible", "Courier Prime", "Inconsolata",
    "JetBrains Mono", "Mukta", "Yanone Kaffeesatz", "Kalam", "Patrick Hand", "Baloo 2",
    "Fredoka", "Lilita One", "Passion One", "Bree Serif", "Russo One", "Black Ops One",
    "Orbitron", "Audiowide", "Press Start 2P", "Bungee", "Staatliches", "Alfa Slab One",
    "Ultra", "Cinzel", "Philosopher", "Marcellus", "Tenor Sans", "Didact Gothic",
    "Poiret One", "Yeseva One", "Rajdhani", "Khand", "Abel", "Acme", "Advent Pro",
    "Alata", "Albert Sans", "Amiri", "Antic Slab", "Arapey", "Architects Daughter",
    "Bai Jamjuree", "Be Vietnam Pro", "Bodoni Moda", "Cairo", "Changa", "Comic Neue",
    "Courgette", "Cuprum", "Dosis", "Epilogue", "Fraunces", "Gloria Hallelujah",
    "Gothic A1", "Gruppo", "Hanken Grotesk", "Ibarra Real Nova", "Instrument Sans",
    "Instrument Serif", "Kumbh Sans", "League Spartan", "Libre Caslon Text",
    "Lexend Deca", "Mada", "Martel", "Montserrat Subrayada", "Newsreader", "Noticia Text",
    "Onest", "Pathway Gothic One", "Philosopher", "Play", "Playfair", "Quattrocento",
    "Quattrocento Sans", "Readex Pro", "Reem Kufi", "Sawarabi Gothic", "Schibsted Grotesk",
    "Sen", "Shippori Mincho", "Signika Negative", "Slabo 27px", "Sofia Sans", "Syne",
    "Tajawal", "Tilt Warp", "Titan One", "Unbounded", "Vazirmatn", "Wix Madefor Display",
    "Ysabeau", "Zen Kaku Gothic New", "M PLUS 1p", "M PLUS Rounded 1c", "Kosugi Maru",
    "Almarai", "Noto Kufi Arabic", "El Messiri", "Lobster Two", "Sacramento",
    "Homemade Apple", "Rock Salt", "Kaushan Script", "Allura", "Parisienne", "Cookie",
    "Tangerine", "Merienda", "Handlee", "Neucha", "Annie Use Your Telescope",
    "Covered By Your Grace", "Reenie Beanie", "Nothing You Could Do", "Bangers",
    "Luckiest Guy", "Fredoka One", "Carter One", "Chewy", "Special Elite", "Creepster",
    "Monoton", "Lato Black",
]}
# Old Google family names that were renamed.
GOOGLE_ALIASES = {
    "Source Sans Pro": "Source Sans 3",
    "Source Serif Pro": "Source Serif 4",
    "Lato Black": "Lato",
    "Fredoka One": "Fredoka",
}

STYLE_WORDS = {
    "thin", "hairline", "extralight", "ultralight", "light", "semilight", "book",
    "regular", "normal", "medium", "semibold", "demibold", "demi", "bold", "extrabold",
    "ultrabold", "heavy", "black", "extrablack", "ultra", "italic", "oblique", "it",
    "cond", "condensed", "semicondensed", "extracondensed", "narrow", "wide", "expanded",
    "extended", "display", "text", "caption", "lt", "std", "pro", "mt", "variable", "vf",
}
# Also matches the same style words written with a space ("Semi Bold", "Extra Light").
_TWO_WORD_STYLES = re.compile(r"\s+(semi|demi|extra|ultra)\s+(bold|light|condensed|expanded)$", re.I)

STATUS_LABELS = {
    "embedded": "Embedded in file",
    "standard": "Standard (Windows & Mac)",
    "office": "Included with Microsoft Office",
    "cloud": "Microsoft 365 cloud font",
    "windows": "Windows only",
    "mac": "Mac only",
    "google": "Free — Google Fonts",
    "nonstandard": "Not standard — must be installed",
}
STATUS_NOTES = {
    "standard": "Available on virtually every Windows and Mac computer.",
    "office": "Installed with Microsoft Office on Windows and Mac.",
    "cloud": "Office (Microsoft 365) downloads this automatically when online. "
             "Older or offline Office versions will substitute another font.",
    "windows": "Comes with Windows. A Mac will substitute a different font.",
    "mac": "Comes with macOS. A Windows PC will substitute a different font.",
    "google": "Free font. Download and install it on every computer that will present this deck.",
    "nonstandard": "Not a standard font. Install it on every presenting computer (it may be a "
                   "paid/commercial font — check the license), or embed fonts in PowerPoint "
                   "(File > Options > Save > Embed fonts in the file).",
}
NEEDS_ACTION = {"google", "nonstandard", "windows", "mac", "cloud"}

_LOOKUP: dict[str, tuple[str, str]] = {}
for _status, _names in (("standard", STANDARD), ("office", OFFICE), ("cloud", CLOUD),
                        ("windows", WINDOWS), ("mac", MAC)):
    for _n in _names:
        _LOOKUP.setdefault(_n.lower(), (_status, _n))
for _n, _g in {**GOOGLE, **GOOGLE_ALIASES}.items():
    _LOOKUP.setdefault(_n.lower(), ("google", GOOGLE.get(_g, _g)))


def _candidates(name: str):
    """Yield the name, then progressively shorter versions without style words."""
    current = re.sub(r"\s+", " ", name).strip()
    yield current
    for _ in range(4):
        stripped = _TWO_WORD_STYLES.sub("", current)
        if stripped == current:
            words = current.split(" ")
            if len(words) > 1 and words[-1].lower().replace("-", "") in STYLE_WORDS:
                stripped = " ".join(words[:-1])
        if stripped == current or not stripped:
            return
        current = stripped
        yield current


def classify(name: str) -> dict:
    """Return {status, label, note, family, links} for a font name."""
    status, family = "nonstandard", name
    for cand in _candidates(name):
        hit = _LOOKUP.get(cand.lower())
        if hit:
            status, family = hit
            # A Google family stays "google" only if matched; keep original for others.
            if status != "google":
                family = cand
            break
    links: list[dict] = []
    if status == "google":
        links.append({"label": "Download from Google Fonts",
                      "url": f"https://fonts.google.com/specimen/{quote_plus(family)}"})
    elif status == "nonstandard":
        links.append({"label": "Search Google Fonts",
                      "url": f"https://fonts.google.com/?query={quote(family)}"})
        links.append({"label": "Search the web",
                      "url": f"https://www.google.com/search?q={quote_plus(chr(34) + name + chr(34) + ' font download')}"})
    return {"status": status, "label": STATUS_LABELS[status], "note": STATUS_NOTES[status],
            "family": family, "links": links}


# --------------------------------------------------------------------------- #
# Fonts installed on this computer
# --------------------------------------------------------------------------- #

_installed: set[str] | None = None
_installed_lock = threading.Lock()


def _font_dirs() -> list[str]:
    home = os.path.expanduser("~")
    if sys.platform.startswith("win"):
        windir = os.environ.get("WINDIR", r"C:\Windows")
        local = os.environ.get("LOCALAPPDATA", os.path.join(home, "AppData", "Local"))
        return [os.path.join(windir, "Fonts"), os.path.join(local, "Microsoft", "Windows", "Fonts")]
    if sys.platform == "darwin":
        return ["/System/Library/Fonts", "/Library/Fonts", os.path.join(home, "Library", "Fonts"),
                "/System/Library/AssetsV2", "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts"]
    return ["/usr/share/fonts", "/usr/local/share/fonts", os.path.join(home, ".fonts"),
            os.path.join(home, ".local", "share", "fonts")]


def _names_from_file(path: str) -> set[str]:
    from fontTools.ttLib import TTCollection, TTFont
    names: set[str] = set()
    try:
        fonts = (TTCollection(path, lazy=True).fonts if path.lower().endswith((".ttc", ".otc"))
                 else [TTFont(path, lazy=True, fontNumber=0)])
    except Exception:
        return names
    for font in fonts:
        try:
            for rec in font["name"].names:
                if rec.nameID in (1, 4, 16):
                    try:
                        names.add(rec.toUnicode().strip().lower())
                    except Exception:
                        continue
        except Exception:
            continue
        finally:
            try:
                font.close()
            except Exception:
                pass
    return names


def installed_font_names() -> set[str] | None:
    """Lower-case family/full names of installed fonts, or None if unavailable."""
    global _installed
    with _installed_lock:
        if _installed is not None:
            return _installed or None
        try:
            import fontTools  # noqa: F401
        except ImportError:
            _installed = set()
            return None
        found: set[str] = set()
        for d in _font_dirs():
            if not os.path.isdir(d):
                continue
            for root, _, files in os.walk(d):
                for f in files:
                    if f.lower().endswith((".ttf", ".otf", ".ttc", ".otc")):
                        found |= _names_from_file(os.path.join(root, f))
        _installed = found
        return found or None


def is_installed(name: str, installed: set[str] | None) -> bool | None:
    if installed is None:
        return None
    return any(c.lower() in installed for c in _candidates(name))


def preload_installed_fonts() -> None:
    """Start scanning installed fonts in the background (called at server start)."""
    threading.Thread(target=installed_font_names, daemon=True).start()
