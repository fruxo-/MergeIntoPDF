# MergeIntoPDF

Combine a folder of PDFs, images, videos, spreadsheets and audio files into a **single PDF**, with every item labeled with its original filename.

It is handy for assembling evidence bundles, attachments or case files from mixed material (phone photos, screen recordings, Excel sheets, existing PDFs) into one document that can be printed or shared.

## What it does

| Input type | Extensions | How it appears in the output |
|---|---|---|
| Images | `.png`, `.jpg`, `.jpeg` | One page per image, labeled with the filename. EXIF orientation is respected, so phone photos are right-side up. |
| Videos | `.mp4`, `.mov`, `.m4v`, `.avi`, `.mkv` | First frame on a page, with a play-button badge and a caption explaining it is a video clip. |
| PDFs | `.pdf` | A title page with the filename, followed by the original pages. |
| Spreadsheets | `.xlsx`, `.xls`, `.xlsm` | Converted to PDF with LibreOffice, then inserted like a regular PDF (title page + pages). |
| Audio | `.qta` | A placeholder page with the filename and a speaker icon, noting that playback requires the digital version. |

Image layout rules:

- **Landscape** images are scaled to fill the full page width.
- **Portrait / square** images are placed at their natural size, centered, and only shrunk if they would not fit (never enlarged).

Extras:

- **Table of contents** at the front of the PDF, listing every processed file with its starting page number (it spans multiple pages if needed).
- **PDF bookmarks** (sidebar outline) for each file.
- **Page numbering** ("Pagina N di TOTALE") on every page.
- **Natural sorting** of filenames, so `file2` comes before `file10`.
- **Fault tolerance**: a file that fails to process is skipped, reported as a red warning, and listed in a summary at the end. It does not stop the run.

> **Note:** The text the tool writes into the PDF (table of contents title, captions, page numbers) is in **Italian**. To change the language, edit the string constants in `mtpdf.py` (e.g. `TOC_TITLE`, the video and audio captions, and the `Pagina N di TOTALE` format in `add_page_numbers`).

## Requirements

- **Python 3.9 or newer**
- Python packages: [`pymupdf`](https://pymupdf.readthedocs.io/), [`pillow`](https://python-pillow.org/), [`opencv-python-headless`](https://pypi.org/project/opencv-python-headless/)
- **LibreOffice**, only needed if your folder contains spreadsheets (`.xlsx`, `.xls`, `.xlsm`)

## Setting up the local environment

There are two ways to run the tool. The first (uv) is the easiest because the script declares its own dependencies.

### Option A: uv (recommended)

The script uses [PEP 723](https://peps.python.org/pep-0723/) inline metadata, so [`uv`](https://docs.astral.sh/uv/) can create an isolated environment and install the dependencies automatically on first run. There is nothing to install manually.

1. Install uv:

   ```bash
   # macOS (Homebrew)
   brew install uv

   # macOS / Linux (official installer)
   curl -LsSf https://astral.sh/uv/install.sh | sh

   # Windows (PowerShell)
   powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
   ```

2. Clone the repository and run the script:

   ```bash
   git clone <your-repo-url>
   cd MergeIntoPDF
   uv run mtpdf.py /path/to/input_folder /path/to/output.pdf
   ```

   On macOS/Linux you can also make it executable and run it directly:

   ```bash
   chmod +x mtpdf.py
   ./mtpdf.py /path/to/input_folder /path/to/output.pdf
   ```

### Option B: virtual environment with pip

```bash
git clone <your-repo-url>
cd MergeIntoPDF

python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install pymupdf pillow opencv-python-headless

python3 mtpdf.py /path/to/input_folder /path/to/output.pdf
```

### Installing LibreOffice (for spreadsheets)

LibreOffice is used in headless mode to convert spreadsheets to PDF. If you never merge spreadsheets, you can skip this step.

```bash
# macOS
brew install --cask libreoffice

# Debian / Ubuntu
sudo apt install libreoffice-calc

# Fedora
sudo dnf install libreoffice-calc
```

The script looks for the `soffice` command on your `PATH`, and on macOS also at `/Applications/LibreOffice.app/Contents/MacOS/soffice`. On **Windows**, add LibreOffice's `program` folder (for example `C:\Program Files\LibreOffice\program`) to your `PATH` so that `soffice` can be found.

## Usage

```bash
uv run mtpdf.py <input_folder> <output.pdf>
```

Example:

```bash
uv run mtpdf.py ~/Documents/case-files ~/Documents/case-files-merged.pdf
```

Typical output:

```
Found 12 files. Merging...
[1/12] 01_contract.pdf
[2/12] 02_photo.jpg
...
Done. Saved combined PDF to: /Users/me/Documents/case-files-merged.pdf
```

Things to know:

- Only files **directly inside** the input folder are processed (subfolders are not scanned), and hidden files (names starting with `.`) are ignored.
- Files are merged in natural filename order. To control the order, prefix names with numbers (`01_`, `02_`, ...).
- The output page size is **A4**. Original PDF pages are appended as they are, without resizing.
- Each spreadsheet conversion has a 180-second timeout.

## Troubleshooting

| Problem | Likely cause / fix |
|---|---|
| `LibreOffice not found` | LibreOffice is not installed or `soffice` is not on your `PATH`. See [Installing LibreOffice](#installing-libreoffice-for-spreadsheets). |
| Spreadsheet conversion fails or hangs | Close any open LibreOffice windows and try again; a running instance can interfere with headless conversion. |
| `ModuleNotFoundError: No module named 'fitz'` | Dependencies are not installed in the active environment. Use `uv run`, or activate your venv and run `pip install pymupdf pillow opencv-python-headless`. |
| `No PDF/image/video files found in that folder.` | The folder has no supported file types at its top level. Check the extensions listed above. |
| A video shows an error | OpenCV could not read a frame from it. The file is skipped and reported as a warning; other files are unaffected. |

## License

This project is licensed under the **GNU General Public License v3.0**. See the [LICENSE](LICENSE) file for details.
