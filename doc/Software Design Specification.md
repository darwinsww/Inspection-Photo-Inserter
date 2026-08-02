# Report Photo Inserter — Software Design Specification

## 1. Project Overview

Develop a standalone Python desktop application that automatically inserts photos into a Microsoft Word document.

The application is intended for generating inspection reports, engineering reports, insurance reports, and similar documents containing large numbers of photos.

The software should be designed for long-term use, with maintainable code and modular architecture.

## 2. Technical Requirements

- Python 3.11+

Libraries:

- `tkinter`
- `python-docx`
- `Pillow`
- `pillow-heif` (optional, for HEIC support)
- `json`
- `pathlib`
- `tempfile`

Code should be modular rather than a single script.

Recommended structure:

```text
ReportPhotoInserter/
│
├── main.py
├── gui.py
├── word_processor.py
├── image_processor.py
├── config_manager.py
├── utils.py
├── settings.json
├── requirements.txt
└── README.md
```

## 3. Functional Requirements

### 3.1 Word Document

When the program starts:

- Prompt the user to choose an existing Word (`.docx`) document.
- If the user cancels the dialog, create a new empty Word document instead.

The original Word file must never be overwritten. The final output should always be saved as a new file chosen by the user via a **Save As** dialog.

### 3.2 Photo Folder

- Prompt the user to select ONE photo folder.
- Only one folder is processed each time.

Supported formats:

- jpg
- jpeg
- png
- heic (if `pillow-heif` is available)

Ignore unsupported files.

### 3.3 Remember Previous Locations

Store user preferences in `settings.json`.

Remember:

- last Word template
- last photo folder
- last output folder

Next time the application starts, automatically use these as the initial directory.

### 3.4 Root Folder Batch Mode (Recursive Scanning)

In addition to selecting a single photo folder (3.2), the user may assign a **root folder**.

- The root folder contains multiple subfolders. Each subfolder may itself contain further subfolders that hold the actual photos (potentially hundreds each).
- The application must **recursively** scan the entire tree under the root folder and read every supported image, regardless of nesting depth.
- All photos found anywhere under the root folder are inserted into **one single output `.docx` file**.
- Within that single file, photos are placed under the matching **template title** (see 3.6), based on their source folder name.

Example tree the application must handle:

```text
Root/
├── BuildingA/
│   ├── Kitchen/     (Kitchen1.jpg, Kitchen2.jpg, ... hundreds)
│   └── Garage/      (Garage1.jpg, ...)
├── BuildingB/
│   └── Kitchen/     (KitchenA.jpg, KitchenB.jpg, ...)
└── BuildingC/
    └── Bathroom/    (Bath1.jpg, ...)
```

### 3.5 Sectioning by Immediate Parent Folder Name

Photos are grouped into sections based on the **name of their immediate parent folder** (the folder directly containing the photo), not the full path.

- All photos whose immediate parent folder shares the **same name** belong to the **same section**, even if those folders live in different paths.
- Each section corresponds to one of the fixed template titles (see 3.6); the shared parent-folder name is matched to the template title of the same name.
- Within a section, photos that come from **different physical folders** (same name, different path) must be visually separated by a **thick red horizontal divider line**. Photos from the same physical folder are **not** separated.

Using the tree in 3.4, `BuildingA/Kitchen` and `BuildingB/Kitchen` both have the immediate parent folder name `Kitchen`, so they share one section. Sections always appear in the fixed template order defined in 3.6 (Kitchen before Bathroom before Garage), **not** in folder-discovery order:

```text
# Kitchen                     <- template title (matched to folder name)
Kitchen1   Kitchen2           (from BuildingA/Kitchen)
Kitchen3
═══════════════════════════   <- thick RED separator: same name, different path
KitchenA   KitchenB           (from BuildingB/Kitchen)

# Bathroom
Bath1      Bath2

# Garage
Garage1    Garage2
```

### 3.6 Output Template (`Template.md`)

The output layout is defined by `doc/Template.md`, a plain-text, version-controllable template. The application builds the output document **programmatically** from this definition — there is **no binary `.docx` template file** to open.

- The application generates the base document from the template definition: the cover page, the "Areas" index, and the 12 section titles, then saves the result as a new output file.
- **The cover page and the "Areas" index must never receive any photos.** They are generated from the template and left as content only (page 1 and page 2 of the output).
- The "Areas" index heading uses Word **Heading 1**; each of the 12 section titles uses Word **Heading 3**.
- **Fonts and numbering** (all text uses the **Calibri (Body)** font):
  - Cover fields use **Calibri (Body), 12pt**. Each label is **bold** up to and including its colon, followed by a non-bold space so text the user types afterwards is not bold. An empty line follows the report title, and two empty lines separate "Inspection Date" and "General comments".
  - The "Areas" heading uses **Calibri (Body), 22pt, black** with **line spacing: multiple 2.41**; its 12 area lines use **Calibri (Body), 12pt**, are **not numbered**, and use **line spacing: 1.5 lines**.
  - Each of the 12 section titles uses **Calibri (Body), 16pt, black**, is **numbered** `1)`, `2)`, … in template order, and uses **line spacing: multiple 1.73**.
- Photo insertion happens only under the **12 fixed titles** defined by the template (in this order):

  1. Smoke Alarm
  2. Living Room
  3. Kitchen
  4. Dining Room
  5. Ensuite
  6. Bedroom 1
  7. Bedroom 2
  8. Bathroom
  9. Laundry
  10. Garage
  11. Exterior
  12. Keys for Tenant

- Photos are inserted by **matching the immediate parent folder name to the title name** (see 3.5). All photos from a folder named `Kitchen` are placed under the `Kitchen` title; photos from `Bathroom` go under `Bathroom`, and so on.
- Matching is case-insensitive and ignores surrounding whitespace. Folder names that do **not** match any of the 12 titles are skipped and reported in the summary (section 11).
- The 12 titles come from the template and are **never renamed, reordered, added to, or removed** — photos are only inserted beneath the existing titles.

**Title order in the output.** The generated `.docx` always contains all 12 titles in exactly the sequence listed above (3.6), independent of the order in which folders are discovered or matched. A title that has **no** matching photos still appears in the output with no photos beneath it. Titles are never sorted by folder name, merged, split, or omitted.

Example — given this photo tree:

```text
Root/
├── House1/
│   ├── Kitchen/     (K1.jpg, K2.jpg, K3.jpg)
│   └── Garage/      (G1.jpg)
├── House2/
│   └── Kitchen/     (KA.jpg, KB.jpg)
└── Misc/
    └── Backyard/    (B1.jpg)      <- no matching title
```

The output document keeps the cover page and Areas index as content only, and inserts photos under the matching titles:

```text
[Page 1] Cover page                     (generated, no photos)
[Page 2] Areas index                    (generated, no photos)
...
Kitchen                                 <- template title
  K1     K2                             (from House1/Kitchen)
  K3
  ══════════════════════════           <- thick RED separator (different path)
  KA     KB                             (from House2/Kitchen)
...
Garage                                  <- template title
  G1                                    (from House1/Garage)
...
```

`Misc/Backyard` does not match any of the 12 titles, so its photo `B1.jpg` is skipped and counted in the summary.

## 4. Image Processing

Original images must NEVER be modified. Instead:

- Create resized temporary copies.
- Temporary images should be deleted automatically after processing.

### 4.1 Resize

Phone images are typically `4032×3024` or `3024×4032`.

Do NOT insert original images directly into Word. Resize before insertion.

Configuration:

- Maximum long edge: 1800 pixels
- JPEG quality: 90
- Use Pillow LANCZOS resampling.

### 4.2 EXIF Orientation

Many phone images rely on EXIF orientation.

Use `ImageOps.exif_transpose()` before determining orientation. Failure to do this will produce incorrect portrait/landscape classification.

## 5. Image Classification

Images are classified as:

- **Landscape:** width >= height
- **Portrait:** height > width

## 6. Sorting

Default sorting: natural filename order.

Example:

```text
Correct:            Incorrect:
1.jpg               1.jpg
2.jpg               10.jpg
3.jpg               11.jpg
10.jpg              2.jpg
11.jpg
```

Future support:

- EXIF date
- Modified date

Design code so new sorting modes can be added easily.

## 7. Layout Algorithm

This is the most important requirement.

- DO NOT separate all landscape images from all portrait images.
- The original photo order must be preserved.
- However, landscape and portrait images must NEVER appear on the same row.

Example input:

```text
01 Landscape
02 Landscape
03 Landscape
04 Portrait
05 Portrait
06 Landscape
07 Landscape
08 Landscape
09 Portrait
```

Correct output:

```text
01    02
03
04    05
06    07
08
09
```

Explanation — whenever image orientation changes:

- Finish the current row.
- Start a new orientation group.

Within each orientation group:

- Two images per row.
- Never mix portrait and landscape in one row.

### 7.1 Row Compaction Within a Folder

Avoid producing too many rows that contain only a single photo. Within the photos of a **single physical folder**, pair images of the **same orientation** together — even if a lone image of that orientation appears earlier and its partner appears later.

Suggested algorithm (per folder): scan photos in order while keeping at most one "pending" landscape and one "pending" portrait. When a second image of the same orientation arrives, emit that orientation's pair as one row. After all photos are scanned, emit any leftover single images. This guarantees at most one single-image row per orientation, per folder.

Example — a single folder containing, in order: `L1, P1, L2, L3, P2` (L = landscape, P = portrait):

```text
Step-by-step:
L1  -> pending landscape = L1
P1  -> pending portrait  = P1
L2  -> pairs with L1     -> emit row [L1  L2]
L3  -> pending landscape = L3
P2  -> pairs with P1     -> emit row [P1  P2]
end -> leftover landscape -> emit row [L3]
```

Resulting layout:

```text
L1  L2
P1  P2
L3
```

Only one single-image row remains, instead of five single-image rows.

### 7.2 Folder Boundaries in Layout

Row compaction must **never cross folder boundaries**. Images from two different physical folders must never share a row — even when the folders have the **same immediate parent-folder name** and appear in the **same section**.

Example — one `Kitchen` section built from two different folders:

```text
Folder A (BuildingA/Kitchen), in order: L1, P1, L2
Folder B (BuildingB/Kitchen), in order: P3, L3
```

Each folder is compacted independently, then separated by the divider line from 3.5:

```text
# Kitchen
L1  L2          (Folder A: the two landscapes pair up)
P1              (Folder A: lone portrait)
════════════    (thick RED separator: different path, same name)
P3              (Folder B: not paired with Folder A's P1)
L3              (Folder B: not paired with Folder A's landscapes)
```

Note how `P1` (Folder A) and `P3` (Folder B) are **not** paired, and `L3` (Folder B) is **not** paired with `L1`/`L2` (Folder A), because pairing cannot cross a folder boundary.

## 8. Word Layout

- Two images per row.
- Full rows (two images) are centered; a row with a single image is **left-aligned** (the lone image sits on the left, not in the middle of the line).
- Constant spacing between columns.
- Suggested spacing: 0.5 cm

DO NOT hardcode image width. Instead, read:

- page width
- left margin
- right margin

Calculate available width automatically.

Formula:

```text
available_width = page_width − left_margin − right_margin
image_width      = (available_width − column_gap) / 2
```

This allows the software to work on A4, Letter, Landscape pages, and custom page sizes without code changes.

- Maintain aspect ratio.
- Never stretch images.

## 9. Paragraph Spacing

- Each image row should have small spacing after it. Suggested: 3 pt
- No unnecessary blank pages.

## 10. Output

- Prompt user with a **Save As** dialog.
- Suggested default filename: `OriginalReport_Photos.docx`
- Never overwrite the original document.

## 11. Error Handling

Skip unreadable images. Display a summary.

Example:

```text
Processed: 152 images
Skipped:   3 images
Output:    Inspection Report_Photos.docx
```

## 12. GUI

Simple Tkinter GUI.

Fields:

- Word document — `[ Browse ]`
- Photo Folder — `[ Browse ]`
- Output File — `[ Browse ]`

Buttons:

- Generate
- Exit

Status bar:

- Ready
- Processing...
- Completed

Progress bar preferred.

## 13. Performance

The application should comfortably handle 200–500 photos without excessive memory usage.

- Images should be processed incrementally.
- Do NOT load every image into RAM simultaneously.

## 14. Future Extension Points

The architecture should make these features easy to add later.

**Placeholder insertion** — Insert photos after `{{PHOTOS}}` instead of always appending to the end.

**Caption** — Optional filename below image. Example:

```text
Kitchen Sink Leak
[photo]
```

**Figure numbering** — Automatically generated:

```text
Figure 1
Figure 2
```

**Multiple folders** — Process `Kitchen/`, `Bathroom/`, `Roof/`. Each folder becomes a separate section.

**PDF Export** — Export directly to PDF.

**Compression settings** — Allow users to configure maximum pixel size and JPEG quality.

## 15. Coding Style

Requirements:

- Type hints
- Dataclasses where appropriate
- PEP8 compliant
- Small reusable functions
- Object-oriented architecture preferred
- Clear comments
- Logging instead of excessive `print()`
- Proper exception handling

## 16. Acceptance Criteria

The software is considered complete when it satisfies all of the following:

- User can generate a report without editing source code.
- Original images remain untouched.
- Original Word document remains untouched.
- Images are automatically resized.
- EXIF orientation is handled correctly.
- Original photo order is preserved.
- Portrait and landscape images are never mixed on the same row.
- Exactly two images per row whenever possible.
- A root folder is scanned recursively, and photos at any nesting depth are collected.
- The output document is generated from the `Template.md` definition, with the cover page and Areas index kept as content only (no photos).
- Photos are inserted only under the 12 fixed template titles, matched by parent-folder name; the titles themselves are never changed.
- The output always contains all 12 titles in the exact sequence defined in 3.6, even when a title has no matching photos.
- Folder names that match no template title are skipped and reported in the summary.
- Photos are grouped into sections by their immediate parent-folder name, with same-name folders from different paths sharing one section.
- Same-name folders from different paths are separated by a thick red divider line within their shared section.
- Single-image rows are minimized by pairing same-orientation photos within a folder.
- Row pairing never crosses folder boundaries.
- Image widths adapt automatically to page size.
- The program remembers the last used folders.
- Processing 300+ phone photos completes successfully without excessive memory usage.
- The codebase is modular and maintainable.
