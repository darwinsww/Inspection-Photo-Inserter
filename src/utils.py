"""General utilities: logging setup, natural sorting, and image discovery."""
from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import List

# Image extensions that are always supported.
BASE_EXTENSIONS = {".jpg", ".jpeg", ".png"}
# HEIC is only usable when pillow-heif is installed.
HEIC_EXTENSIONS = {".heic"}


def setup_logging(level: int = logging.INFO) -> None:
    """Configure application-wide logging."""
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )


def supported_extensions(include_heic: bool) -> set[str]:
    """Return the set of supported file extensions."""
    extensions = set(BASE_EXTENSIONS)
    if include_heic:
        extensions |= HEIC_EXTENSIONS
    return extensions


def _natural_chunks(text: str) -> list:
    """Split a string into text/number chunks for natural ordering."""
    return [
        int(chunk) if chunk.isdigit() else chunk.lower()
        for chunk in re.split(r"(\d+)", text)
    ]


def natural_sort_key(path: Path) -> list:
    """Sort key producing natural (human) filename ordering.

    Ensures ``2.jpg`` sorts before ``10.jpg`` rather than lexicographically.
    """
    return _natural_chunks(path.name)


def list_images(folder: Path, include_heic: bool) -> List[Path]:
    """Return supported image files in *folder*, naturally sorted.

    Unsupported files are ignored. Sub-folders are not traversed.
    """
    extensions = supported_extensions(include_heic)
    files = [
        entry
        for entry in folder.iterdir()
        if entry.is_file() and entry.suffix.lower() in extensions
    ]
    files.sort(key=natural_sort_key)
    return files


@dataclass
class FolderImages:
    """Supported images found directly inside one physical folder.

    ``section_name`` is the immediate parent folder name and is used to group
    same-named folders (from different paths) into a single document section.
    """

    folder: Path
    section_name: str
    images: List[Path]


def discover_folders(root: Path, include_heic: bool) -> List[FolderImages]:
    """Recursively find every folder under *root* that contains images.

    The whole tree is traversed, regardless of nesting depth (spec 3.4). Each
    physical folder that directly contains at least one supported image becomes
    one :class:`FolderImages` entry, tagged with its own folder name as the
    section name (spec 3.5). Traversal and file order are deterministic and
    natural.
    """
    extensions = supported_extensions(include_heic)
    results: List[FolderImages] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # Sort sub-directories so traversal order is deterministic and natural.
        dirnames.sort(key=_natural_chunks)
        folder = Path(dirpath)
        images = [
            folder / name
            for name in filenames
            if Path(name).suffix.lower() in extensions
        ]
        if images:
            images.sort(key=natural_sort_key)
            results.append(
                FolderImages(
                    folder=folder,
                    section_name=folder.name,
                    images=images,
                )
            )
    return results
