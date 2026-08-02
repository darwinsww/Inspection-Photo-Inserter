"""Application entry point for the Report Photo Inserter."""
from __future__ import annotations

from gui import ReportPhotoInserterApp
from utils import setup_logging


def main() -> None:
    setup_logging()
    app = ReportPhotoInserterApp()
    app.run()


if __name__ == "__main__":
    main()
