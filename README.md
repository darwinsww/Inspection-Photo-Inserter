# Report Photo Inserter

A standalone Python desktop application that automatically inserts photos into a
Microsoft Word document. It is designed for generating inspection, engineering,
insurance and similar reports that contain large numbers of photos.

## Features

- Generates the report from the Markdown template definition
  ([doc/Template.md](doc/Template.md)) — a fixed cover page, an "Areas" index,
  and 12 fixed section titles.
- Scans a chosen photo folder **recursively** and collects every supported image.
- Inserts photos under the template title matching each photo's **immediate
  parent folder name** (case-insensitive); the 12 titles keep their order and
  are never changed. Folders matching no title are skipped and reported.
- Same-named folders from different paths share a title but are separated by a
  **thick red divider line**.
- **Original images are never modified** — resized temporary copies are used and
  deleted automatically after processing.
- Automatic resizing (long edge capped, LANCZOS resampling, JPEG quality 90).
- Correct **EXIF orientation** handling before portrait/landscape classification.
- Portrait and landscape images are **never mixed on the same row**; single-image
  rows are minimized by pairing same-orientation photos within a folder.
- Two images per row: full rows are centered, a single-image row is left-aligned;
  image width is derived automatically from the page size and margins (works with
  A4, Letter, landscape and custom pages).
- Styled output: all text in Calibri (Body); cover fields 12pt with bold labels,
  the "Areas" heading (Heading 1) in 22pt, and the 12 section titles (Heading 3)
  in 16pt, numbered `1)`, `2)`, ….
- After report generation, also creates a checklist text file named `To do. txt`
  in the same output folder.
- Remembers the last used photo folder and output folder.
- Handles 200–500 photos with low memory usage (images are processed
  incrementally, never all loaded into RAM at once).
- Optional HEIC support when `pillow-heif` is installed.

## Project structure

```
ReportPhotoInserter/
├── src/
│   ├── main.py              # Entry point
│   ├── gui.py               # Tkinter GUI
│   ├── word_processor.py    # Report building, layout, template rendering
│   ├── template_loader.py   # Parses doc/Template.md
│   ├── image_processor.py   # Resizing, EXIF, orientation
│   ├── config_manager.py    # settings.json persistence
│   └── utils.py             # Logging, natural sort, recursive image discovery
├── doc/
│   ├── Template.md          # Output template definition
│   └── Software Design Specification.md
├── scripts/
│   └── install-deps.ps1  # Project-only dependency install
├── settings.json        # Remembered folder locations
├── .pip/
│   └── pip.ini          # Project-local pip config (Artifactory mirror)
├── requirements.txt
└── README.md
```

## Files

- `src/main.py` — entry point (logging + launches GUI)
- `src/gui.py` — Tkinter GUI with browse fields, progress bar, status bar, background worker thread
- `src/word_processor.py` — template rendering, folder-to-title matching, row compaction, borderless tables (single-image rows left-aligned)
- `src/template_loader.py` — parses `doc/Template.md` (cover page, areas index, 12 titles)
- `src/image_processor.py` — EXIF transpose, LANCZOS resize, orientation classification, temp copies
- `src/config_manager.py` — `settings.json` persistence
- `src/utils.py` — logging, natural sort, recursive image discovery
- `doc/Template.md` — output template definition (rendered into the report)
- `scripts/install-deps.ps1` — installs requirements with project-local pip config
- `.pip/pip.ini` — project-only package index (Artifactory mirror)
- `settings.json`, `requirements.txt`, `README.md`

## Spec compliance highlights

- **Layout algorithm**: portrait and landscape never share a row; two per row, pairing same-orientation photos within a folder to minimize single-image rows (which are left-aligned); pairing never crosses a folder boundary.
- **Width derived from page**: `(page_width - left - right - gap) / 2`, so A4/Letter/landscape/custom all work with no code changes; aspect ratio preserved (width-only sizing).
- **Originals untouched**: resized JPEG temp copies (long edge 1800, quality 90), deleted right after embedding; output always via Save As.
- **EXIF** handled via `ImageOps.exif_transpose()` before classification.
- **Memory-safe**: images processed incrementally to disk temps, never all in RAM; temps removed after embedding.
- Natural filename sort verified (`1,2,3,10,11`); remembers last photo/output locations; skip-and-summarize error handling; type hints, dataclasses, logging throughout.

## Requirements

- Python 3.11+
- Dependencies listed in `requirements.txt`
  (`tkinter` ships with the standard CPython installer).

## Installation

### Windows PowerShell (virtual environment)

1. Open PowerShell in the project folder.
2. Create a virtual environment:

```powershell
python -m venv .venv
```

3. Activate the virtual environment:

```powershell
.\.venv\Scripts\Activate.ps1
```

4. If activation is blocked by execution policy, run this once and activate again:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

5. (Optional - work laptops only):   
   5.1 configure the project-local package source. Edit
      `.pip/pip.ini` and replace `<USERNAME>` with your Iress username and
      `<TOKEN>` with a fresh Artifactory Identity Token (generated at
      `https://iress.jfrog.io` under your profile).

      > Public PyPI (`files.pythonhosted.org`) is blocked on the Iress network, so
      > this project installs from the Iress Artifactory mirror instead. The token
      > lives only in this project's `.pip/pip.ini` — do not commit real credentials.

   5.2 install dependencies with the project-only
      wrapper script:

      ```powershell
      .\scripts\install-deps.ps1
      ```

      This script temporarily sets `PIP_CONFIG_FILE` to `.pip/pip.ini` so only this
project uses the configured index, then restores your previous environment.

6. : install dependencies with pip:

```powershell
pip install -r requirements.txt
```

Use this step instead of step 6 when you are not on a work laptop or do not need
the project-local mirror.

7. Run the app:

```powershell
python src/main.py
```

8. When finished, you can leave the virtual environment with:

```powershell
deactivate
```

> HEIC support is optional. If `pillow-heif` fails to install on your platform,
> the application still runs and simply ignores `.heic` files.

## Usage

```powershell
python src/main.py
```

1. **Photo folder** — select the root folder to scan. Sub-folders are scanned
   **recursively**; photos are grouped by their immediate parent folder name and
   inserted under the matching section title from the template. Supported
   formats: `.jpg`, `.jpeg`, `.png`, and `.heic` (if available).
2. **Output file** — choose where to save the generated report via *Save As*.
3. Click **Generate**. A progress bar and status bar report progress, and a
   summary dialog shows how many images were processed, how many were skipped as
  unreadable, and any folders that matched no template title.
4. After completion, the app writes `To do. txt` next to the generated `.docx`
  with this checklist:

```text
1. 删除分隔的红线
2. 删除 UV 灯拍的照片
3. 写评语
4. 检查 docx 中照片数量是否和文件夹中的一致
5. 所有内容备份到 G 盘
```

## Layout logic

Photos are grouped by their source folder and placed two per row. Within a
folder, same-orientation photos are paired together — even when not adjacent — to
minimize rows that contain a single photo; a leftover single photo is
left-aligned. Portrait and landscape images never share a row, and pairing never
crosses a folder boundary.

```
Folder photos (in order):  L P L L P
Output:                    [L L]
                           [P P]
                           [L]
```

## Configuration

Defaults (edit the dataclasses in the source to change):

- Maximum long edge: `1800` px (`ImageConfig.max_long_edge`)
- JPEG quality: `90` (`ImageConfig.jpeg_quality`)
- Images per row: `2` (`LayoutConfig.images_per_row`)
- Column gap: `0.5` cm (`LayoutConfig.column_gap_cm`)
- Spacing after each row: `3` pt (`LayoutConfig.row_space_after_pt`)

## Future extension points

The architecture is intentionally modular to allow later additions such as
placeholder-based insertion (`{{PHOTOS}}`), captions, figure numbering,
multiple folders/sections, PDF export, and additional sorting modes
(EXIF date, modified date).
