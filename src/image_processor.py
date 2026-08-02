"""Image loading, EXIF handling, resizing, and orientation classification.

Original images are never modified. Each source image is opened, corrected for
EXIF orientation, resized, and written to a temporary JPEG copy that is later
embedded into the Word document and then deleted.
"""
from __future__ import annotations

import logging
import shutil
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from tempfile import mkdtemp
from typing import Optional

from PIL import Image, ImageOps

logger = logging.getLogger(__name__)

# Enable HEIC support if pillow-heif is available.
try:
    import pillow_heif  # type: ignore

    pillow_heif.register_heif_opener()
    HEIC_SUPPORTED = True
    logger.debug("pillow-heif registered: HEIC support enabled")
except Exception:  # pragma: no cover - optional dependency
    HEIC_SUPPORTED = False
    logger.debug("pillow-heif not available: HEIC support disabled")


class Orientation(Enum):
    """Image orientation category used for layout grouping."""

    LANDSCAPE = "landscape"
    PORTRAIT = "portrait"


@dataclass(frozen=True)
class ImageConfig:
    """Resize and encoding configuration."""

    max_long_edge: int = 1800
    jpeg_quality: int = 90


@dataclass
class ProcessedImage:
    """A resized temporary image ready for insertion."""

    original_path: Path
    temp_path: Path
    orientation: Orientation
    width: int
    height: int


class ImageProcessor:
    """Creates resized temporary copies of source images.

    Temporary files live in a dedicated directory that is removed by
    :meth:`cleanup`.
    """

    def __init__(self, config: ImageConfig | None = None) -> None:
        self.config = config or ImageConfig()
        self._temp_dir = Path(mkdtemp(prefix="report_photos_"))
        self._counter = 0
        logger.debug("Temporary directory: %s", self._temp_dir)

    @property
    def temp_dir(self) -> Path:
        return self._temp_dir

    def process(self, path: Path) -> Optional[ProcessedImage]:
        """Resize a single image and return metadata, or ``None`` if unreadable."""
        try:
            with Image.open(path) as raw:
                # Apply EXIF orientation BEFORE deciding portrait/landscape.
                oriented = ImageOps.exif_transpose(raw)
                rgb = oriented.convert("RGB")
                resized = self._resize(rgb)
        except (OSError, ValueError, Image.DecompressionBombError) as exc:
            logger.warning("Skipping unreadable image %s (%s)", path.name, exc)
            return None

        orientation = self._classify(resized.width, resized.height)
        self._counter += 1
        temp_path = self._temp_dir / f"{self._counter:04d}_{path.stem}.jpg"
        try:
            resized.save(
                temp_path,
                format="JPEG",
                quality=self.config.jpeg_quality,
                optimize=True,
            )
        except OSError as exc:
            logger.warning("Could not write temp copy for %s (%s)", path.name, exc)
            return None

        return ProcessedImage(
            original_path=path,
            temp_path=temp_path,
            orientation=orientation,
            width=resized.width,
            height=resized.height,
        )

    def _resize(self, image: Image.Image) -> Image.Image:
        """Downscale so the long edge does not exceed the configured maximum."""
        long_edge = max(image.width, image.height)
        if long_edge <= self.config.max_long_edge:
            return image.copy()
        scale = self.config.max_long_edge / long_edge
        new_size = (round(image.width * scale), round(image.height * scale))
        return image.resize(new_size, Image.LANCZOS)

    @staticmethod
    def _classify(width: int, height: int) -> Orientation:
        """Landscape when width >= height, otherwise portrait."""
        return Orientation.LANDSCAPE if width >= height else Orientation.PORTRAIT

    def cleanup(self) -> None:
        """Delete all temporary images and the temporary directory."""
        shutil.rmtree(self._temp_dir, ignore_errors=True)
        logger.debug("Removed temporary directory: %s", self._temp_dir)
