"""Signatures used to recognise third-party add-ins / interactive integrations.

Detection is keyword based: each signature lists lowercase substrings that are
searched for in web-extension XML, slide tag parts, hyperlink targets,
OLE/ActiveX metadata and custom XML parts. Add your own entries to
KNOWN_ADDINS to detect more tools.
"""
from __future__ import annotations

KNOWN_ADDINS: list[dict] = [
    {"name": "Slido", "category": "Polling / Q&A",
     "keywords": ["slido", "sli.do"]},
    {"name": "Mentimeter", "category": "Polling / Q&A",
     "keywords": ["mentimeter", "menti.com"]},
    {"name": "Poll Everywhere", "category": "Polling / Q&A",
     "keywords": ["polleverywhere", "poll everywhere", "pollev.com", "pollev"]},
    {"name": "Kahoot!", "category": "Quiz",
     "keywords": ["kahoot"]},
    {"name": "Microsoft Forms", "category": "Forms / Survey",
     "keywords": ["forms.office.com", "forms.microsoft.com", "microsoft forms", "forms.office365"]},
    {"name": "Google Forms", "category": "Forms / Survey",
     "keywords": ["docs.google.com/forms", "forms.gle"]},
    {"name": "AhaSlides", "category": "Polling / Q&A",
     "keywords": ["ahaslides"]},
    {"name": "Wooclap", "category": "Polling / Q&A",
     "keywords": ["wooclap"]},
    {"name": "Vevox", "category": "Polling / Q&A",
     "keywords": ["vevox"]},
    {"name": "ClassPoint", "category": "Classroom interaction",
     "keywords": ["classpoint"]},
    {"name": "Nearpod", "category": "Classroom interaction",
     "keywords": ["nearpod"]},
    {"name": "Pear Deck", "category": "Classroom interaction",
     "keywords": ["peardeck", "pear deck"]},
    {"name": "Quizizz / Wayground", "category": "Quiz",
     "keywords": ["quizizz", "wayground"]},
    {"name": "Socrative", "category": "Quiz",
     "keywords": ["socrative"]},
    {"name": "Sendsteps", "category": "Polling / Q&A",
     "keywords": ["sendsteps"]},
    {"name": "Crowdpurr", "category": "Polling / Q&A",
     "keywords": ["crowdpurr"]},
    {"name": "LiveWeb (web page embed)", "category": "Web embed",
     "keywords": ["liveweb"]},
    {"name": "Microsoft Web Viewer", "category": "Web embed",
     "keywords": ["web viewer", "webviewer"]},
    {"name": "Microsoft Stream", "category": "Web video",
     "keywords": ["microsoftstream.com", "stream.microsoft.com"]},
    {"name": "Power BI", "category": "Live data",
     "keywords": ["powerbi", "power bi"]},
]


def match_addins(text: str) -> list[dict]:
    """Return every known add-in signature whose keyword occurs in text."""
    lowered = text.lower()
    hits = []
    for sig in KNOWN_ADDINS:
        for kw in sig["keywords"]:
            if kw in lowered:
                hits.append({"name": sig["name"], "category": sig["category"], "keyword": kw})
                break
    return hits
