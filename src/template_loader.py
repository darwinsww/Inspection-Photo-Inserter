"""Parse the Markdown output-template definition (``doc/Template.md``).

The template is a plain-text file that defines the generated report layout: the
cover page, the "Areas" index, and the fixed list of section titles that photos
are inserted under. Keeping it in Markdown lets the titles change without code
changes (spec 3.6).
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

logger = logging.getLogger(__name__)

# Template.md lives in the project's doc/ folder, next to the source tree.
DEFAULT_TEMPLATE_PATH = Path(__file__).resolve().parent.parent / "doc" / "Template.md"

_HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
_BOLD_LINE = re.compile(r"^\*\*(.+)\*\*$")
_BULLET = re.compile(r"^-(?:\s+(.*))?$")
_ORDERED = re.compile(r"^\d+\.\s+(.*)$")


@dataclass
class Template:
    """Parsed output-template definition."""

    cover_title: str = ""
    cover_lines: List[str] = field(default_factory=list)
    areas_heading: str = "Areas"
    areas: List[str] = field(default_factory=list)
    titles: List[str] = field(default_factory=list)


def normalize_title(name: str) -> str:
    """Normalize a title/folder name for case- and whitespace-insensitive match."""
    return name.strip().lower()


def load_template(path: Optional[Path] = None) -> Template:
    """Load and parse the template definition from *path* (or the default)."""
    template_path = path or DEFAULT_TEMPLATE_PATH
    text = template_path.read_text(encoding="utf-8")
    template = _parse(text)
    logger.info(
        "Loaded template: %d title(s), %d cover line(s)",
        len(template.titles),
        len(template.cover_lines),
    )
    return template


def _parse(text: str) -> Template:
    template = Template()
    section: Optional[str] = None

    for raw in text.splitlines():
        stripped = raw.strip()

        heading = _HEADING.match(stripped)
        if heading:
            level = len(heading.group(1))
            title = heading.group(2).strip()
            low = title.lower()
            if level == 2:
                # Top-level section headings select the current parse mode.
                if "cover" in low:
                    section = "cover"
                elif "areas index" in low or "index" in low:
                    section = "areas"
                elif "section titles" in low or low.startswith("pages"):
                    section = "titles"
                else:
                    section = None
            elif level == 3:
                if section == "areas":
                    template.areas_heading = title
                elif section == "titles":
                    template.titles.append(title)
            # Level-1 (document title) and other headings are ignored.
            continue

        if not stripped or stripped.startswith(">"):
            continue

        if section == "cover":
            bold = _BOLD_LINE.match(stripped)
            if bold:
                template.cover_title = bold.group(1).strip()
                continue
            bullet = _BULLET.match(stripped)
            if bullet:
                # An empty bullet (``-``) becomes a blank line on the cover page.
                template.cover_lines.append((bullet.group(1) or "").strip())
        elif section == "areas":
            ordered = _ORDERED.match(stripped)
            if ordered:
                template.areas.append(ordered.group(1).strip())
                continue
            bullet = _BULLET.match(stripped)
            if bullet:
                template.areas.append(bullet.group(1).strip())
        # In the "titles" section only the ### headings matter; prose is ignored.

    return template
