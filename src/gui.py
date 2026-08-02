"""Tkinter GUI for the Report Photo Inserter application.

The GUI collects the photo (root) folder and the output file, then runs the
photo-insertion pipeline on a background thread so the interface stays
responsive. The output is generated from the Markdown template definition
(spec 3.6).
"""
from __future__ import annotations

import logging
import threading
from pathlib import Path
from tkinter import StringVar, Tk, filedialog, messagebox, ttk
from typing import Callable, List, Optional

from config_manager import ConfigManager
from image_processor import HEIC_SUPPORTED, ImageConfig, ImageProcessor
from template_loader import load_template, normalize_title
from utils import discover_folders
from word_processor import FolderPhotos, LayoutConfig, WordProcessor

logger = logging.getLogger(__name__)

DEFAULT_OUTPUT_NAME = "OriginalReport_Photos.docx"
TODO_OUTPUT_NAME = "To do. txt"
TODO_LINES = [
    "1. 删除分隔的红线",
    "2. 删除 UV 灯拍的照片",
    "3. 写评语",
    "4. 检查 docx 中照片数量是否和文件夹中的一致",
    "5. 所有内容备份到 G 盘",
]


def _directory_of(path_text: str) -> str:
    """Return an existing directory to seed a file dialog, or empty string."""
    if not path_text:
        return ""
    path = Path(path_text)
    if path.is_dir():
        return str(path)
    if path.parent.is_dir():
        return str(path.parent)
    return ""


class ReportPhotoInserterApp:
    """Main application window."""

    def __init__(self) -> None:
        self.config = ConfigManager()
        settings = self.config.settings

        self.root = Tk()
        self.root.title("Report Photo Inserter")
        self.root.resizable(False, False)

        self.folder_var = StringVar(value=settings.last_photo_folder)
        self.output_var = StringVar()
        self.status_var = StringVar(value="Ready")

        self._build_ui()

    # UI construction -----------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 6}
        frame = ttk.Frame(self.root, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(frame, text="Photo folder (scanned recursively):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.folder_var, width=50).grid(row=0, column=1, **pad)
        ttk.Button(frame, text="Browse", command=self._browse_folder).grid(row=0, column=2, **pad)

        ttk.Label(frame, text="Output file:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frame, textvariable=self.output_var, width=50).grid(row=1, column=1, **pad)
        ttk.Button(frame, text="Browse", command=self._browse_output).grid(row=1, column=2, **pad)

        self.progress = ttk.Progressbar(frame, length=420, mode="determinate", maximum=100)
        self.progress.grid(row=2, column=0, columnspan=3, sticky="we", **pad)

        buttons = ttk.Frame(frame)
        buttons.grid(row=3, column=0, columnspan=3, **pad)
        self.generate_btn = ttk.Button(buttons, text="Generate", command=self._on_generate)
        self.generate_btn.grid(row=0, column=0, padx=6)
        ttk.Button(buttons, text="Exit", command=self._on_exit).grid(row=0, column=1, padx=6)

        status = ttk.Label(frame, textvariable=self.status_var, relief="sunken", anchor="w")
        status.grid(row=4, column=0, columnspan=3, sticky="we", **pad)

    # Dialog handlers -----------------------------------------------------

    def _browse_folder(self) -> None:
        path = filedialog.askdirectory(
            title="Select photo folder (sub-folders are scanned recursively)",
            initialdir=_directory_of(self.folder_var.get())
            or _directory_of(self.config.settings.last_photo_folder),
        )
        if path:
            self.folder_var.set(path)

    def _browse_output(self) -> None:
        path = self._ask_output()
        if path:
            self.output_var.set(path)

    def _ask_output(self) -> Optional[str]:
        path = filedialog.asksaveasfilename(
            title="Save report as",
            initialdir=_directory_of(self.config.settings.last_output_folder),
            initialfile=DEFAULT_OUTPUT_NAME,
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")],
        )
        return path or None

    # Actions -------------------------------------------------------------

    def _on_generate(self) -> None:
        folder_text = self.folder_var.get().strip()
        if not folder_text or not Path(folder_text).is_dir():
            messagebox.showerror("Missing folder", "Please select a valid photo folder.")
            return

        output_text = self.output_var.get().strip()
        if not output_text:
            output_text = self._ask_output()
            if not output_text:
                return
            self.output_var.set(output_text)

        folder_path = Path(folder_text)
        output_path = Path(output_text)
        # Force a .docx extension so Windows shows the Word icon and the file
        # opens in Word (a missing/incorrect extension yields a blank icon).
        if output_path.suffix.lower() != ".docx":
            output_path = output_path.with_suffix(".docx")
            self.output_var.set(str(output_path))

        self._persist(folder_text, output_path)

        self.generate_btn.config(state="disabled")
        self.progress["value"] = 0
        self._set_status("Processing...")

        thread = threading.Thread(
            target=self._worker,
            args=(folder_path, output_path),
            daemon=True,
        )
        thread.start()

    def _worker(
        self,
        folder_path: Path,
        output_path: Path,
    ) -> None:
        """Run the full pipeline on a background thread."""
        processor = ImageProcessor(ImageConfig())
        word_processor = WordProcessor(LayoutConfig())
        folders_photos: list[FolderPhotos] = []
        processed_count = 0
        skipped = 0
        try:
            template = load_template()
            title_keys = {normalize_title(title) for title in template.titles}

            discovered = discover_folders(folder_path, HEIC_SUPPORTED)
            if not discovered:
                self._ui(lambda: self._on_empty())
                return

            # Only folders whose name matches a template title are processed;
            # the rest are reported as skipped (spec 3.6).
            matched_dirs = [
                folder
                for folder in discovered
                if normalize_title(folder.section_name) in title_keys
            ]
            unmatched_names = list(
                dict.fromkeys(
                    folder.section_name
                    for folder in discovered
                    if normalize_title(folder.section_name) not in title_keys
                )
            )

            total = max(sum(len(folder.images) for folder in matched_dirs), 1)
            done = 0
            for folder in matched_dirs:
                processed_images = []
                for path in folder.images:
                    result = processor.process(path)
                    done += 1
                    if result is None:
                        skipped += 1
                    else:
                        processed_images.append(result)
                    fraction = 0.5 * done / total
                    self._ui(lambda f=fraction: self._set_progress(f, "Processing..."))
                if processed_images:
                    folders_photos.append(
                        FolderPhotos(
                            section_name=folder.section_name,
                            folder=folder.folder,
                            images=processed_images,
                        )
                    )
                    processed_count += len(processed_images)

            processed_total = max(processed_count, 1)

            def on_insert(inserted: int, _total: int) -> None:
                fraction = 0.5 + 0.5 * inserted / processed_total
                self._ui(lambda f=fraction: self._set_progress(f, "Inserting..."))

            document, _ = word_processor.build_report(
                template, folders_photos, progress=on_insert
            )
            word_processor.save(document, output_path)
            todo_path = output_path.parent / TODO_OUTPUT_NAME
            self._write_todo_file(todo_path)
        except Exception as exc:  # noqa: BLE001 - report any failure to the user
            logger.exception("Generation failed")
            self._ui(lambda message=str(exc): self._on_error(message))
            return
        finally:
            processor.cleanup()

        self._ui(
            lambda: self._on_success(
                processed_count, skipped, unmatched_names, output_path
            )
        )

    def _write_todo_file(self, todo_path: Path) -> None:
        """Write a post-generation checklist next to the output document."""
        content = "\n".join(TODO_LINES) + "\n"
        todo_path.write_text(content, encoding="utf-8")
        logger.info("Saved todo file: %s", todo_path)

    # UI update helpers (always called on the main thread) ----------------

    def _ui(self, func: Callable[[], None]) -> None:
        self.root.after(0, func)

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _set_progress(self, fraction: float, status: str) -> None:
        self.progress["value"] = max(0, min(100, fraction * 100))
        self._set_status(status)

    def _on_success(
        self,
        processed: int,
        skipped: int,
        unmatched: List[str],
        output_path: Path,
    ) -> None:
        self.progress["value"] = 100
        self._set_status("Completed")
        self.generate_btn.config(state="normal")
        lines = [
            f"Processed: {processed} images",
            f"Skipped (unreadable): {skipped} images",
        ]
        if unmatched:
            names = ", ".join(unmatched)
            lines.append(
                f"Skipped (no matching title): {len(unmatched)} folder(s): {names}"
            )
        lines.append(f"Output: {output_path.name}")
        messagebox.showinfo("Report complete", "\n".join(lines))

    def _on_error(self, message: str) -> None:
        self._set_status("Error")
        self.generate_btn.config(state="normal")
        messagebox.showerror("Generation failed", message)

    def _on_empty(self) -> None:
        self._set_status("Ready")
        self.generate_btn.config(state="normal")
        messagebox.showwarning(
            "No images found",
            "The selected folder contains no supported images.",
        )

    # Persistence ---------------------------------------------------------

    def _persist(self, folder_text: str, output_path: Path) -> None:
        settings = self.config.settings
        settings.last_photo_folder = folder_text
        settings.last_output_folder = str(output_path.parent)
        self.config.save()

    def _on_exit(self) -> None:
        self.config.save()
        self.root.destroy()

    def run(self) -> None:
        self.root.protocol("WM_DELETE_WINDOW", self._on_exit)
        self.root.mainloop()
