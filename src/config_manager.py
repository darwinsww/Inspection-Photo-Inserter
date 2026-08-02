"""Persistent user preferences stored in ``settings.json``."""
from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass, fields
from pathlib import Path

logger = logging.getLogger(__name__)


@dataclass
class Settings:
    """User preferences remembered between sessions."""

    last_word_template: str = ""
    last_photo_folder: str = ""
    last_output_folder: str = ""


class ConfigManager:
    """Loads and saves :class:`Settings` to a JSON file on disk."""

    def __init__(self, path: Path | None = None) -> None:
        # Keep settings.json at project root even though code lives in src/.
        self.path = path or Path(__file__).resolve().parent.parent / "settings.json"
        self.settings = self._load()

    def _load(self) -> Settings:
        if self.path.exists():
            try:
                data = json.loads(self.path.read_text(encoding="utf-8"))
                valid = {f.name: data.get(f.name, "") for f in fields(Settings)}
                return Settings(**valid)
            except (json.JSONDecodeError, OSError, TypeError) as exc:
                logger.warning("Could not read settings (%s); using defaults", exc)
        return Settings()

    def save(self) -> None:
        """Persist the current settings to disk."""
        try:
            self.path.write_text(
                json.dumps(asdict(self.settings), indent=2),
                encoding="utf-8",
            )
        except OSError as exc:
            logger.warning("Could not save settings: %s", exc)
