#!/usr/bin/env -S uv run
# /// script
# requires-python = ">=3.9"
# dependencies = [
#     "pymupdf",
#     "pillow",
#     "opencv-python-headless",
# ]
# ///
"""
merge_to_pdf.py

Combine a folder of PDFs, images, videos, and XLSX spreadsheets into a
single PDF, labeling each item with its original filename.

Layout rules:
  - EXIF orientation is respected first (so phone photos display right-side up).
  - Landscape images: scaled to fill the full page width.
  - Portrait images: placed at their natural size, centered, only shrunk
    if they're too big to fit on the page (never enlarged).
  - Videos: the first frame is extracted and shown, with a play-button
    badge and a caption clarifying it's a video clip (not a still photo).
  - PDFs: a title page is inserted, then the PDF's original pages are appended.
  - XLSX/XLS files: converted to PDF (via LibreOffice) and inserted the
    same way as native PDFs (title page + converted pages).
  - QTA (audio) files: a placeholder page is inserted with the filename
    and an audio icon, noting that playback requires the digital version.

A table of contents is generated at the very front of the output,
listing every successfully processed file with its starting page
number, followed by real PDF bookmarks (sidebar outline) for navigation.

Any file that fails to process is skipped (not fatal) and reported as a
red warning on stdout, with a final summary at the end.

Usage:
    python3 merge_to_pdf.py /path/to/input_folder /path/to/output.pdf

Requires:
    pip3 install pymupdf pillow opencv-python-headless
    LibreOffice installed (for XLSX conversion): brew install --cask libreoffice
"""

import sys
import re
import os
import tempfile
import shutil
import subprocess

import fitz  # PyMuPDF
from PIL import Image, ImageOps
import cv2

PAGE_WIDTH, PAGE_HEIGHT = 595, 842  # A4 in points
MARGIN = 36
LABEL_HEIGHT = 28
FONT_SIZE = 11

IMAGE_EXTS = {".png", ".jpg", ".jpeg"}
VIDEO_EXTS = {".mp4", ".mov", ".m4v", ".avi", ".mkv"}
SPREADSHEET_EXTS = {".xlsx", ".xls", ".xlsm"}
AUDIO_EXTS = {".qta"}
PDF_EXT = ".pdf"

# ANSI colors for stdout warnings.
RED = "\033[91m"
BOLD = "\033[1m"
RESET = "\033[0m"

TEMP_DIR = tempfile.mkdtemp(prefix="merge_to_pdf_")

# Common locations for LibreOffice's command-line binary on macOS.
SOFFICE_CANDIDATES = [
    "soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def natural_key(s):
    """Sort key so 'file2' comes before 'file10'."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def wrap_text(text, fontname, fontsize, max_width):
    """Break text into lines that each fit within max_width."""
    words = text.split(" ")
    lines = []
    current = ""

    def fits(s):
        return fitz.get_text_length(s, fontname=fontname, fontsize=fontsize) <= max_width

    for word in words:
        candidate = f"{current} {word}".strip()
        if fits(candidate):
            current = candidate
        else:
            if current:
                lines.append(current)
            # word itself might be longer than max_width (long filename, no spaces)
            if fits(word):
                current = word
            else:
                # hard-break the long word/token by characters
                chunk = ""
                for ch in word:
                    if fits(chunk + ch):
                        chunk += ch
                    else:
                        lines.append(chunk)
                        chunk = ch
                current = chunk
    if current:
        lines.append(current)
    return lines


def add_label(page, text, y=MARGIN, fontsize=FONT_SIZE, max_width=None):
    """
    Draw a filename label near the top of a page, wrapping to multiple
    lines if it's too wide. Returns the total height the label occupied
    (so callers can lay out content below it).
    """
    if max_width is None:
        max_width = PAGE_WIDTH - 2 * MARGIN

    lines = wrap_text(text, "helv", fontsize, max_width)
    line_height = fontsize * 1.3

    cursor_y = y
    for line in lines:
        page.insert_text(
            (MARGIN, cursor_y),
            line,
            fontsize=fontsize,
            fontname="helv",
            color=(0, 0, 0),
        )
        cursor_y += line_height

    underline_y = cursor_y - line_height + fontsize * 0.4
    page.draw_line(
        (MARGIN, underline_y),
        (PAGE_WIDTH - MARGIN, underline_y),
        color=(0.6, 0.6, 0.6),
        width=0.5,
    )

    total_height = cursor_y - y
    return total_height


def extract_first_frame(video_path):
    """Grab the first frame of a video and return a PIL Image (RGB)."""
    cap = cv2.VideoCapture(video_path)
    ok, frame = cap.read()
    cap.release()
    if not ok or frame is None:
        raise RuntimeError("could not read a frame from video")
    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
    return Image.fromarray(frame_rgb)


_SOFFICE_PATH = None


def find_soffice():
    """Locate the LibreOffice command-line binary, or raise a clear error."""
    global _SOFFICE_PATH
    if _SOFFICE_PATH:
        return _SOFFICE_PATH

    for candidate in SOFFICE_CANDIDATES:
        path = shutil.which(candidate) if os.path.sep not in candidate else candidate
        if path and os.path.exists(path):
            _SOFFICE_PATH = path
            return path

    raise RuntimeError(
        "LibreOffice not found. Install it with: brew install --cask libreoffice"
    )


def convert_xlsx_to_pdf(xlsx_path):
    """Convert an XLSX/XLS/XLSM file to PDF using LibreOffice headless mode.
    Returns the path to the converted PDF (in TEMP_DIR)."""
    soffice = find_soffice()
    convert_dir = tempfile.mkdtemp(dir=TEMP_DIR)

    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--norestore",
            "--convert-to", "pdf",
            "--outdir", convert_dir,
            xlsx_path,
        ],
        capture_output=True,
        text=True,
        timeout=180,
    )

    base = os.path.splitext(os.path.basename(xlsx_path))[0]
    converted_path = os.path.join(convert_dir, f"{base}.pdf")

    if result.returncode != 0 or not os.path.exists(converted_path):
        raise RuntimeError(
            f"LibreOffice conversion failed: {result.stderr.strip() or result.stdout.strip()}"
        )

    return converted_path


def prepare_image(src_path, is_video):
    """
    Load an image (or first video frame), apply correct EXIF orientation
    (no forced rotation beyond that), and save as a temp JPEG.
    Returns the temp file path.
    """
    if is_video:
        img = extract_first_frame(src_path)
    else:
        img = Image.open(src_path)
        img = ImageOps.exif_transpose(img)  # respect camera orientation tag

    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")

    fd, temp_path = tempfile.mkstemp(suffix=".jpg", dir=TEMP_DIR)
    os.close(fd)
    img.save(temp_path, "JPEG", quality=92)
    return temp_path


def compute_placement(img_w, img_h, avail_w, avail_h):
    """
    Landscape (w > h): fill the available width, scale height to match.
    Portrait/square (w <= h): natural size, only shrink (never enlarge)
    if it doesn't fit the available area.
    """
    if img_w > img_h:
        scale = avail_w / img_w
        draw_w, draw_h = avail_w, img_h * scale
        if draw_h > avail_h:
            scale = avail_h / img_h
            draw_w, draw_h = img_w * scale, avail_h
    else:
        scale = min(1.0, avail_w / img_w, avail_h / img_h)
        draw_w, draw_h = img_w * scale, img_h * scale
    return draw_w, draw_h


def add_image_page(out_doc, image_path, label, is_video=False):
    prepared_path = prepare_image(image_path, is_video)

    page = out_doc.new_page(width=PAGE_WIDTH, height=PAGE_HEIGHT)
    label_text = f"{label} (primo fotogramma)" if is_video else label
    label_height = add_label(page, label_text)

    img_top = MARGIN + label_height + 6
    avail_w = PAGE_WIDTH - 2 * MARGIN

    # Reserve space at the bottom for the video icon + caption, if needed.
    caption_reserve = 70 if is_video else 0
    avail_h = PAGE_HEIGHT - img_top - MARGIN - caption_reserve

    with Image.open(prepared_path) as im:
        img_w, img_h = im.size

    draw_w, draw_h = compute_placement(img_w, img_h, avail_w, avail_h)
    x0 = MARGIN + (avail_w - draw_w) / 2
    y0 = img_top + (avail_h - draw_h) / 2
    rect = fitz.Rect(x0, y0, x0 + draw_w, y0 + draw_h)

    page.insert_image(rect, filename=prepared_path)

    if is_video:
        caption_y = img_top + avail_h + 12
        add_video_caption(page, caption_y)


def add_pdf_with_title(out_doc, pdf_path, label):
    title_page = out_doc.new_page(width=PAGE_WIDTH, height=PAGE_HEIGHT)
    label_height = add_label(title_page, label)

    big_fontsize = 16
    max_width = PAGE_WIDTH - 2 * MARGIN
    big_lines = wrap_text(f"Contenuto di: {label}", "helv", big_fontsize, max_width)
    big_line_height = big_fontsize * 1.3

    # Vertically center the (possibly multiline) big text in the space
    # below the header label.
    block_height = len(big_lines) * big_line_height
    top_bound = MARGIN + label_height + 20
    bottom_bound = PAGE_HEIGHT - MARGIN
    start_y = top_bound + max(0, (bottom_bound - top_bound - block_height) / 2)

    cursor_y = start_y
    for line in big_lines:
        title_page.insert_text(
            (MARGIN, cursor_y),
            line,
            fontsize=big_fontsize,
            fontname="helv",
            color=(0, 0, 0),
        )
        cursor_y += big_line_height

    with fitz.open(pdf_path) as src:
        out_doc.insert_pdf(src)


def draw_audio_icon(page, cx, cy, scale=1.0):
    """Draw a simple speaker + sound-wave icon centered at (cx, cy)."""
    color = (0.25, 0.25, 0.25)

    # Speaker body: rectangle base + trapezoid cone, as one closed polygon.
    speaker_points = [
        fitz.Point(cx - 50 * scale, cy - 15 * scale),
        fitz.Point(cx - 25 * scale, cy - 15 * scale),
        fitz.Point(cx, cy - 40 * scale),
        fitz.Point(cx, cy + 40 * scale),
        fitz.Point(cx - 25 * scale, cy + 15 * scale),
        fitz.Point(cx - 50 * scale, cy + 15 * scale),
        fitz.Point(cx - 50 * scale, cy - 15 * scale),
    ]
    shape = page.new_shape()
    shape.draw_polyline(speaker_points)
    shape.finish(color=color, fill=color, closePath=True)

    # Sound waves: three curved arcs radiating to the right of the cone tip.
    for radius in (15 * scale, 30 * scale, 45 * scale):
        top = fitz.Point(cx + radius * 0.3, cy - radius)
        bottom = fitz.Point(cx + radius * 0.3, cy + radius)
        ctrl = fitz.Point(cx + radius * 1.3, cy)
        shape.draw_curve(top, ctrl, bottom)
    shape.finish(color=color, width=3 * scale)

    shape.commit()


def draw_video_badge(page, cx, cy, radius=18):
    """Draw a circular play-button badge (video indicator) at (cx, cy)."""
    color = (0.25, 0.25, 0.25)
    shape = page.new_shape()
    shape.draw_circle(fitz.Point(cx, cy), radius)
    shape.finish(color=color, fill=(1, 1, 1), width=2)

    tri_size = radius * 0.8
    p1 = fitz.Point(cx - tri_size * 0.35, cy - tri_size * 0.55)
    p2 = fitz.Point(cx - tri_size * 0.35, cy + tri_size * 0.55)
    p3 = fitz.Point(cx + tri_size * 0.55, cy)
    shape.draw_polyline([p1, p2, p3, p1])
    shape.finish(color=color, fill=color, closePath=True)

    shape.commit()


def add_video_caption(page, y):
    """Draw a play-button icon plus a centered 'video clip' caption at
    vertical position y. Returns the height the caption block occupied."""
    icon_radius = 12
    gap_after_icon = 8
    caption_fontsize = 12
    line_height = caption_fontsize * 1.4

    max_width = PAGE_WIDTH - 2 * MARGIN
    lines = wrap_text(
        "Videoclip \u2014 mostrato il primo fotogramma. Fare riferimento "
        "alla versione digitale per guardare il video completo.",
        "helv", caption_fontsize, max_width,
    )

    icon_cy = y + icon_radius
    draw_video_badge(page, PAGE_WIDTH / 2, icon_cy, radius=icon_radius)

    text_top = icon_cy + icon_radius + gap_after_icon
    for i, line in enumerate(lines):
        line_width = fitz.get_text_length(line, fontname="helv", fontsize=caption_fontsize)
        x = (PAGE_WIDTH - line_width) / 2
        page.insert_text(
            (x, text_top + i * line_height),
            line,
            fontsize=caption_fontsize,
            fontname="helv",
            color=(0.25, 0.25, 0.25),
        )

    return (text_top + len(lines) * line_height) - y


def add_audio_placeholder_page(out_doc, label):
    """Insert a placeholder page for an audio file: filename label,
    a speaker icon, and a note to consult the digital version."""
    page = out_doc.new_page(width=PAGE_WIDTH, height=PAGE_HEIGHT)
    add_label(page, label)

    cx, cy = PAGE_WIDTH / 2, PAGE_HEIGHT / 2 - 40
    draw_audio_icon(page, cx, cy, scale=1.4)

    note_fontsize = 13
    max_width = PAGE_WIDTH - 2 * MARGIN
    note_lines = wrap_text(
        "File audio \u2014 fare riferimento alla versione digitale di "
        "questo documento per ascoltare questo file.",
        "helv", note_fontsize, max_width,
    )
    line_height = note_fontsize * 1.4
    text_top = cy + 90
    # Center the note block horizontally line by line.
    for i, line in enumerate(note_lines):
        line_width = fitz.get_text_length(line, fontname="helv", fontsize=note_fontsize)
        x = (PAGE_WIDTH - line_width) / 2
        page.insert_text(
            (x, text_top + i * line_height),
            line,
            fontsize=note_fontsize,
            fontname="helv",
            color=(0.25, 0.25, 0.25),
        )


TOC_TITLE = "Indice degli allegati"
TOC_TITLE_CONTINUED = "Indice degli allegati (continua)"
TOC_FONTSIZE = 11
TOC_TITLE_FONTSIZE = 18
TOC_LINE_HEIGHT = TOC_FONTSIZE * 1.6
TOC_PAGE_NUM_COL_WIDTH = 50


def _toc_title_block_height():
    return (TOC_TITLE_FONTSIZE * 1.8) + 10


def _toc_entries_per_page():
    avail_h = PAGE_HEIGHT - MARGIN - MARGIN - _toc_title_block_height()
    return max(1, int(avail_h // TOC_LINE_HEIGHT))


def compute_toc_page_count(num_entries):
    """How many TOC pages will be needed for this many entries."""
    if num_entries == 0:
        return 0
    per_page = _toc_entries_per_page()
    return -(-num_entries // per_page)  # ceil division


def truncate_to_width(text, fontname, fontsize, max_width):
    """Truncate text with an ellipsis so it fits within max_width."""
    if fitz.get_text_length(text, fontname=fontname, fontsize=fontsize) <= max_width:
        return text
    for i in range(len(text), 0, -1):
        candidate = text[:i].rstrip() + "\u2026"
        if fitz.get_text_length(candidate, fontname=fontname, fontsize=fontsize) <= max_width:
            return candidate
    return "\u2026"


def build_toc_pages(toc_doc, entries, page_offset):
    """
    Render the table of contents into toc_doc.
    entries: list of (label, start_page_index_0based_in_content).
    page_offset: number of TOC pages preceding the content (so final,
    1-indexed page numbers can be computed).
    """
    per_page = _toc_entries_per_page()
    label_max_width = (PAGE_WIDTH - 2 * MARGIN) - TOC_PAGE_NUM_COL_WIDTH - 10

    for start in range(0, len(entries), per_page):
        chunk = entries[start:start + per_page]
        is_first = (start == 0)

        page = toc_doc.new_page(width=PAGE_WIDTH, height=PAGE_HEIGHT)
        title_text = TOC_TITLE if is_first else TOC_TITLE_CONTINUED
        page.insert_text(
            (MARGIN, MARGIN + TOC_TITLE_FONTSIZE),
            title_text,
            fontsize=TOC_TITLE_FONTSIZE,
            fontname="helv",
            color=(0, 0, 0),
        )

        cursor_y = MARGIN + _toc_title_block_height()
        for label, start_idx in chunk:
            page_num = start_idx + page_offset + 1
            page_num_text = str(page_num)
            trunc_label = truncate_to_width(label, "helv", TOC_FONTSIZE, label_max_width)

            page.insert_text(
                (MARGIN, cursor_y + TOC_FONTSIZE),
                trunc_label,
                fontsize=TOC_FONTSIZE,
                fontname="helv",
                color=(0, 0, 0),
            )
            num_width = fitz.get_text_length(page_num_text, fontname="helv", fontsize=TOC_FONTSIZE)
            page.insert_text(
                (PAGE_WIDTH - MARGIN - num_width, cursor_y + TOC_FONTSIZE),
                page_num_text,
                fontsize=TOC_FONTSIZE,
                fontname="helv",
                color=(0, 0, 0),
            )
            cursor_y += TOC_LINE_HEIGHT


def add_page_numbers(doc):
    """Stamp 'Pagina N di TOTALE' centered in the bottom margin of every page."""
    fontsize = 9
    total = doc.page_count
    y = PAGE_HEIGHT - 20

    for i in range(total):
        page = doc[i]
        text = f"Pagina {i + 1} di {total}"
        width = fitz.get_text_length(text, fontname="helv", fontsize=fontsize)
        x = (PAGE_WIDTH - width) / 2
        page.insert_text(
            (x, y),
            text,
            fontsize=fontsize,
            fontname="helv",
            color=(0.4, 0.4, 0.4),
        )


def main():
    if len(sys.argv) != 3:
        print("Usage: python3 merge_to_pdf.py <input_folder> <output.pdf>")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.isdir(input_folder):
        print(f"Error: {input_folder} is not a folder")
        sys.exit(1)

    all_exts = IMAGE_EXTS | VIDEO_EXTS | SPREADSHEET_EXTS | AUDIO_EXTS | {PDF_EXT}
    files = [
        f for f in os.listdir(input_folder)
        if os.path.splitext(f)[1].lower() in all_exts
        and not f.startswith(".")
    ]
    files.sort(key=natural_key)

    if not files:
        print("No PDF/image/video files found in that folder.")
        sys.exit(1)

    print(f"Found {len(files)} files. Merging...")

    content_doc = fitz.open()
    toc_doc = None
    final_doc = None
    failures = []
    entries = []  # (label, start_page_index_0based_in_content)

    try:
        for i, fname in enumerate(files, 1):
            path = os.path.join(input_folder, fname)
            ext = os.path.splitext(fname)[1].lower()
            print(f"[{i}/{len(files)}] {fname}")
            start_idx = content_doc.page_count
            try:
                if ext == PDF_EXT:
                    add_pdf_with_title(content_doc, path, fname)
                elif ext in IMAGE_EXTS:
                    add_image_page(content_doc, path, fname, is_video=False)
                elif ext in VIDEO_EXTS:
                    add_image_page(content_doc, path, fname, is_video=True)
                elif ext in SPREADSHEET_EXTS:
                    converted_pdf = convert_xlsx_to_pdf(path)
                    add_pdf_with_title(content_doc, converted_pdf, fname)
                elif ext in AUDIO_EXTS:
                    add_audio_placeholder_page(content_doc, fname)
                entries.append((fname, start_idx))
            except Exception as e:
                failures.append((fname, str(e)))
                print(f"{RED}{BOLD}WARNING: failed to process '{fname}': {e}{RESET}")

        toc_page_count = compute_toc_page_count(len(entries))
        if toc_page_count > 0:
            toc_doc = fitz.open()
            build_toc_pages(toc_doc, entries, toc_page_count)

        final_doc = fitz.open()
        if toc_doc is not None:
            final_doc.insert_pdf(toc_doc)
        final_doc.insert_pdf(content_doc)

        if entries:
            bookmarks = [
                [1, label, start_idx + toc_page_count + 1]
                for label, start_idx in entries
            ]
            try:
                final_doc.set_toc(bookmarks)
            except Exception:
                pass  # bookmarks are a bonus; don't fail the whole run over them

        add_page_numbers(final_doc)
        final_doc.save(output_path)
        print(f"\nDone. Saved combined PDF to: {output_path}")

        if failures:
            print(
                f"\n{RED}{BOLD}WARNING: {len(failures)} of {len(files)} "
                f"file(s) failed to process and were skipped:{RESET}"
            )
            for fname, reason in failures:
                print(f"{RED}  - {fname}: {reason}{RESET}")
    finally:
        content_doc.close()
        if toc_doc is not None:
            toc_doc.close()
        if final_doc is not None:
            final_doc.close()
        shutil.rmtree(TEMP_DIR, ignore_errors=True)


if __name__ == "__main__":
    main()
