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

## Output language

All text the tool writes **into the PDF** (table of contents title, video and audio captions, page numbers) is translated. Pick the language with `--lang` (default: `it`, as before):

```bash
uv run mtpdf.py ./input out.pdf --lang en
uv run mtpdf.py ./input out.pdf --lang ja
uv run mtpdf.py --list-languages
```

| Code | Language | Code | Language |
|---|---|---|---|
| `it` | Italiano (default) | `ru` | Русский |
| `en` | English | `zh` | 中文 (简体) |
| `fr` | Français | `ja` | 日本語 |
| `es` | Español | `hi` | हिन्दी (Hindi) |
| `de` | Deutsch | | |

Regional forms such as `zh-CN` or `pt_BR` are accepted and reduced to their base language code.

### Fonts for non-Latin languages

Italian, English, French, Spanish and German use the built-in Helvetica font, so nothing extra is needed. Russian, Chinese, Japanese and Hindi need a font that contains their characters. The tool looks for one, in this order:

1. any `.ttf` / `.otf` / `.ttc` file in a **`fonts/`** folder next to `mtpdf.py`;
2. well-known fonts installed on your system (Arial Unicode, Noto, DejaVu, PingFang, Microsoft YaHei, Nirmala, ...);
3. whatever `fontconfig` suggests (Linux/macOS with `fc-match`).

A font is only accepted if it really contains every character the chosen language needs. If none is found you get a clear error. In that case either download a font (for example from [Google Fonts / Noto](https://fonts.google.com/noto)) into `fonts/`, or point to one directly:

```bash
uv run mtpdf.py ./input out.pdf --lang hi --font ~/Downloads/NotoSansDevanagari-Regular.ttf
```

For Chinese and Japanese, if no font is found the tool falls back to PyMuPDF's built-in CJK font, which is not embedded in the PDF (the viewer must supply it).

Notes:

- File names are drawn with the same font as the rest of the text. If you use a Latin-only language (for example `--lang it`) but your file names contain Cyrillic, CJK or other non-Latin characters, pass a Unicode font with `--font`.
- Hindi is drawn through PyMuPDF's HTML engine, which performs the text shaping Devanagari requires.
- With Noto Sans CJK, Chinese and Japanese share one font file, so some characters use the same glyph variant in both.
- Right-to-left languages (Arabic, Hebrew, Urdu) are not supported yet.

### Adding a language

Languages live in the `LANGUAGES` table near the top of `mtpdf.py`. To add one:

1. Copy the `"en"` block and change its key to the new language code.
2. Translate the strings, keeping the `{placeholders}`. Any string you leave out falls back to English.
3. If the language uses characters outside Western Latin-1 (Cyrillic, Greek, CJK, Indic, ...), optionally add a `"fonts"` list of font paths to try and a `"font_hint"` for the error message. Add `"shaped": True` for scripts that need complex text shaping.

Example (Portuguese, Latin-1 only, so no font settings are needed):

```python
"pt": {
    "name": "Português",
    "strings": {
        "toc_title": "Índice de anexos",
        "toc_title_continued": "Índice de anexos (continuação)",
        "video_first_frame": "{label} (primeiro fotograma)",
        "pdf_contents_of": "Conteúdo de: {label}",
        "video_caption": "Videoclipe — é mostrado o primeiro fotograma. Consulte a versão digital para ver o vídeo completo.",
        "audio_note": "Ficheiro de áudio — consulte a versão digital deste documento para ouvir este ficheiro.",
        "page_x_of_y": "Página {page} de {total}",
    },
},
```

The new code then appears automatically in `--list-languages` and `--help`.

## Requirements

- **Python 3.9 or newer**
- Python packages: [`pymupdf`](https://pymupdf.readthedocs.io/) (1.24 or newer), [`pillow`](https://python-pillow.org/), [`opencv-python-headless`](https://pypi.org/project/opencv-python-headless/), [`fonttools`](https://pypi.org/project/fonttools/) (used to shrink embedded fonts)
- A font covering the language, only for Russian, Chinese, Japanese and Hindi (see [Fonts for non-Latin languages](#fonts-for-non-latin-languages))
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

pip install "pymupdf>=1.24" pillow opencv-python-headless fonttools

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
uv run mtpdf.py <input_folder> <output.pdf> [--lang CODE] [--font FILE]
uv run mtpdf.py --list-languages
```

| Option | Description |
|---|---|
| `-l`, `--lang CODE` | Language of the text written into the PDF (default: `it`). See [Output language](#output-language). |
| `--font FILE` | Use this `.ttf` / `.otf` font for all text instead of auto-detecting one. |
| `--list-languages` | Print the available languages and exit. |

Example:

```bash
uv run mtpdf.py ~/Documents/case-files ~/Documents/case-files-merged.pdf --lang en
```

Typical output:

```
Found 12 files. Merging... (PDF language: English [en])
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
| `ModuleNotFoundError: No module named 'fitz'` | Dependencies are not installed in the active environment. Use `uv run`, or activate your venv and run `pip install "pymupdf>=1.24" pillow opencv-python-headless fonttools`. |
| `no installed font covers the 'xx' text` | Russian, Chinese, Japanese and Hindi need a font. Put a `.ttf`/`.otf` in a `fonts/` folder next to the script, or use `--font`. See [Fonts for non-Latin languages](#fonts-for-non-latin-languages). |
| Empty boxes instead of file names | The font lacks the characters in the file names. Use `--font` with a broader Unicode font such as Noto Sans. |
| `No PDF/image/video files found in that folder.` | The folder has no supported file types at its top level. Check the extensions listed above. |
| A video shows an error | OpenCV could not read a frame from it. The file is skipped and reported as a warning; other files are unaffected. |

## License

This project is licensed under the **GNU General Public License v3.0**. See the [LICENSE](LICENSE) file for details.
