#!/usr/bin/env -S uv run
# /// script
# requires-python = ">=3.9"
# dependencies = [
#     "pymupdf>=1.24,<2",
#     "pillow",
#     "opencv-python-headless",
#     "fonttools",
# ]
# ///
"""
mtpdf.py

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

Language:
  All text written INTO the PDF (table of contents title, captions, page
  numbers) is translated. Choose it with --lang (default: it). Run
  `mtpdf.py --list-languages` to see the available languages.
  To add a language, see "HOW TO ADD A LANGUAGE" below the LANGUAGES table.

Usage:
    python3 mtpdf.py INPUT_FOLDER OUTPUT.pdf [--lang CODE] [--font FILE]
    python3 mtpdf.py --list-languages

Requires:
    pip3 install "pymupdf>=1.24,<2" pillow opencv-python-headless fonttools
    LibreOffice installed (for XLSX conversion): brew install --cask libreoffice
"""

import argparse
import glob
import html
import os
import re
import shutil
import string
import subprocess
import sys
import tempfile

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
YELLOW = "\033[93m"
BOLD = "\033[1m"
RESET = "\033[0m"

TEMP_DIR = tempfile.mkdtemp(prefix="merge_to_pdf_")

# Common locations for LibreOffice's command-line binary on macOS.
SOFFICE_CANDIDATES = [
    "soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


# ---------------------------------------------------------------------------
# Languages
# ---------------------------------------------------------------------------
#
# Every language is one entry in LANGUAGES, keyed by its ISO 639-1 code.
#
# HOW TO ADD A LANGUAGE
#   1. Copy any block below (the English one is a good template).
#   2. Change the key to the new language code, e.g. "pt".
#   3. Translate the values of "strings". Keep the {placeholders} as they are
#      (they are filled in by the program). Any string you leave out falls
#      back to English, so a partial translation still works.
#   4. Fonts:
#        - If every character you used exists in Western Latin-1
#          (Portuguese, Dutch, Swedish, ...), you are done: the built-in
#          Helvetica font is used and nothing else is needed.
#        - Otherwise (Cyrillic, Greek, Arabic, CJK, Indic, ...) the program
#          looks for a suitable font installed on the computer. You can
#          help it with the optional keys:
#            "fonts":        list of font file paths / glob patterns to try
#            "font_hint":    font name shown in the error message if none is found
#            "builtin_font": last-resort PyMuPDF CJK font (china-s, china-t,
#                            japan, korea); not embedded in the PDF
#            "shaped":       True for scripts that need complex text shaping
#                            (Devanagari and other Indic scripts, ...)
#      A font file placed in a "fonts" folder next to this script is always
#      tried first, and --font FILE overrides everything.
#
# Right-to-left scripts (Arabic, Hebrew, Urdu) are not supported yet.

LANGUAGES = {
    "it": {
        "name": "Italiano",
        "strings": {
            "toc_title": "Indice degli allegati",
            "toc_title_continued": "Indice degli allegati (continua)",
            "video_first_frame": "{label} (primo fotogramma)",
            "pdf_contents_of": "Contenuto di: {label}",
            "video_caption": (
                "Videoclip — mostrato il primo fotogramma. Fare riferimento "
                "alla versione digitale per guardare il video completo."
            ),
            "audio_note": (
                "File audio — fare riferimento alla versione digitale di "
                "questo documento per ascoltare questo file."
            ),
            "page_x_of_y": "Pagina {page} di {total}",
        },
    },
    "en": {
        "name": "English",
        "strings": {
            "toc_title": "Index of attachments",
            "toc_title_continued": "Index of attachments (continued)",
            "video_first_frame": "{label} (first frame)",
            "pdf_contents_of": "Contents of: {label}",
            "video_caption": (
                "Video clip — first frame shown. Refer to the digital "
                "version to watch the full video."
            ),
            "audio_note": (
                "Audio file — refer to the digital version of this "
                "document to listen to this file."
            ),
            "page_x_of_y": "Page {page} of {total}",
        },
    },
    "fr": {
        "name": "Français",
        "strings": {
            "toc_title": "Index des pièces jointes",
            "toc_title_continued": "Index des pièces jointes (suite)",
            "video_first_frame": "{label} (première image)",
            "pdf_contents_of": "Contenu de : {label}",
            "video_caption": (
                "Clip vidéo — première image affichée. Consultez la "
                "version numérique pour regarder la vidéo complète."
            ),
            "audio_note": (
                "Fichier audio — consultez la version numérique de ce "
                "document pour écouter ce fichier."
            ),
            "page_x_of_y": "Page {page} sur {total}",
        },
    },
    "es": {
        "name": "Español",
        "strings": {
            "toc_title": "Índice de anexos",
            "toc_title_continued": "Índice de anexos (continuación)",
            "video_first_frame": "{label} (primer fotograma)",
            "pdf_contents_of": "Contenido de: {label}",
            "video_caption": (
                "Videoclip — se muestra el primer fotograma. Consulte la "
                "versión digital para ver el vídeo completo."
            ),
            "audio_note": (
                "Archivo de audio — consulte la versión digital de este "
                "documento para escuchar este archivo."
            ),
            "page_x_of_y": "Página {page} de {total}",
        },
    },
    "de": {
        "name": "Deutsch",
        "strings": {
            "toc_title": "Anlagenverzeichnis",
            "toc_title_continued": "Anlagenverzeichnis (Fortsetzung)",
            "video_first_frame": "{label} (erstes Einzelbild)",
            "pdf_contents_of": "Inhalt von: {label}",
            "video_caption": (
                "Videoclip — es wird das erste Einzelbild angezeigt. Das "
                "vollständige Video finden Sie in der digitalen Version."
            ),
            "audio_note": (
                "Audiodatei — diese Datei können Sie in der digitalen "
                "Version dieses Dokuments anhören."
            ),
            "page_x_of_y": "Seite {page} von {total}",
        },
    },
    "ru": {
        "name": "Русский",
        "font_hint": "Noto Sans, Arial or DejaVu Sans",
        "strings": {
            "toc_title": "Список приложений",
            "toc_title_continued": "Список приложений (продолжение)",
            "video_first_frame": "{label} (первый кадр)",
            "pdf_contents_of": "Содержимое: {label}",
            "video_caption": (
                "Видеоклип — показан первый кадр. "
                "Полное видео смотрите в цифровой версии."
            ),
            "audio_note": (
                "Аудиофайл — чтобы прослушать этот файл, "
                "обратитесь к цифровой версии этого документа."
            ),
            "page_x_of_y": "Страница {page} из {total}",
        },
    },
    "zh": {
        "name": "中文（简体）",
        "font_hint": "Noto Sans CJK SC",
        "builtin_font": "china-s",
        "fonts": [
            # macOS
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Medium.ttc",
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
            # Linux
            "/usr/share/fonts/**/NotoSansCJK*-Regular.ttc",
            "/usr/share/fonts/**/NotoSansSC*.[ot]tf",
            "/usr/share/fonts/**/wqy-*.ttc",
            "/usr/share/fonts/**/DroidSansFallback*.ttf",
            # Windows
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simsun.ttc",
            "C:/Windows/Fonts/simhei.ttf",
        ],
        "strings": {
            "toc_title": "附件目录",
            "toc_title_continued": "附件目录（续）",
            "video_first_frame": "{label}（第一帧）",
            "pdf_contents_of": "文件内容：{label}",
            "video_caption": "视频片段：此处显示的是第一帧。如需观看完整视频，请参阅电子版。",
            "audio_note": "音频文件：如需收听此文件，请参阅本文档的电子版。",
            "page_x_of_y": "第 {page} 页，共 {total} 页",
        },
    },
    "ja": {
        "name": "日本語",
        "font_hint": "Noto Sans CJK JP",
        "builtin_font": "japan",
        "fonts": [
            # macOS
            "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
            "/System/Library/Fonts/Hiragino Sans GB.ttc",
            # Linux
            "/usr/share/fonts/**/NotoSansCJK*-Regular.ttc",
            "/usr/share/fonts/**/NotoSansJP*.[ot]tf",
            "/usr/share/fonts/**/ipag*.ttf",
            "/usr/share/fonts/**/takao*.ttf",
            # Windows
            "C:/Windows/Fonts/YuGothR.ttc",
            "C:/Windows/Fonts/meiryo.ttc",
            "C:/Windows/Fonts/msgothic.ttc",
        ],
        "strings": {
            "toc_title": "添付資料目次",
            "toc_title_continued": "添付資料目次（続き）",
            "video_first_frame": "{label}（最初のフレーム）",
            "pdf_contents_of": "内容：{label}",
            "video_caption": "ビデオクリップ：最初のフレームを表示しています。完全な動画はデジタル版でご覧ください。",
            "audio_note": "音声ファイル：このファイルを再生するには、本書のデジタル版をご覧ください。",
            "page_x_of_y": "{total} ページ中 {page} ページ",
        },
    },
    "hi": {
        "name": "हिन्दी",
        "font_hint": "Noto Sans Devanagari",
        "shaped": True,  # Devanagari needs complex text shaping
        "fonts": [
            # macOS
            "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
            "/Library/Fonts/Arial Unicode.ttf",
            # Linux
            "/usr/share/fonts/**/NotoSansDevanagari-Regular.ttf",
            "/usr/share/fonts/**/NotoSansDevanagariUI-Regular.ttf",
            "/usr/share/fonts/**/Lohit-Devanagari.ttf",
            "/usr/share/fonts/**/Mukta-Regular.ttf",
            # Windows
            "C:/Windows/Fonts/Nirmala.ttf",
            "C:/Windows/Fonts/mangal.ttf",
        ],
        "strings": {
            "toc_title": "संलग्नकों की अनुक्रमणिका",
            "toc_title_continued": "संलग्नकों की अनुक्रमणिका (जारी)",
            "video_first_frame": "{label} (पहला फ़्रेम)",
            "pdf_contents_of": "सामग्री: {label}",
            "video_caption": (
                "वीडियो क्लिप: पहला फ़्रेम दिखाया गया है। "
                "पूरा वीडियो देखने के लिए डिजिटल संस्करण देखें।"
            ),
            "audio_note": (
                "ऑडियो फ़ाइल: इस फ़ाइल को सुनने के लिए "
                "इस दस्तावेज़ का डिजिटल संस्करण देखें।"
            ),
            "page_x_of_y": "पृष्ठ {page} / {total}",
        },
    },
}

DEFAULT_LANGUAGE = "it"   # kept as the default so existing usage is unchanged
FALLBACK_LANGUAGE = "en"  # used for any string a language does not define

# Set by configure_language() at startup.
STRINGS = {}
FACE = None


def tr(key, **kwargs):
    """Return the translated string for `key` in the selected language."""
    return STRINGS[key].format(**kwargs)


def language_code(value):
    """argparse type: accept 'it', 'IT', 'zh-CN', 'zh_cn', ... -> 'it' / 'zh'."""
    key = value.strip().lower().replace("_", "-")
    for candidate in (key, key.split("-")[0]):
        if candidate in LANGUAGES:
            return candidate
    raise argparse.ArgumentTypeError(
        f"unknown language '{value}'. Available: {', '.join(sorted(LANGUAGES))}"
    )


# ---------------------------------------------------------------------------
# Fonts
# ---------------------------------------------------------------------------
#
# Latin-1 languages use the built-in Helvetica. Other scripts need a real font:
# we look for one on the system (or in ./fonts next to this script), and only
# accept a font that actually contains every character the language needs.

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOCAL_FONT_DIR = os.path.join(SCRIPT_DIR, "fonts")

# Broad-coverage fonts tried for any language after its own "fonts" list.
GENERIC_FONT_PATTERNS = [
    # macOS
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "~/Library/Fonts/NotoSans-Regular.ttf",
    # Linux
    "/usr/share/fonts/**/NotoSans-Regular.ttf",
    "/usr/share/fonts/**/DejaVuSans.ttf",
    "/usr/share/fonts/**/LiberationSans-Regular.ttf",
    "~/.local/share/fonts/**/NotoSans-Regular.ttf",
    "~/.fonts/**/NotoSans-Regular.ttf",
    # Windows
    "C:/Windows/Fonts/arial.ttf",
    "C:/Windows/Fonts/segoeui.ttf",
]


def warn(message):
    print(f"{YELLOW}{BOLD}WARNING:{RESET}{YELLOW} {message}{RESET}")


def _chars_needed(strings):
    """Characters that must exist in the font (all translated text + digits)."""
    text = "".join(re.sub(r"\{[^}]*\}", "", s) for s in strings.values())
    text += string.digits
    return sorted({c for c in text if not c.isspace()})


def _missing_glyphs(font, chars):
    return [c for c in chars if not font.has_glyph(ord(c))]


def _fits_builtin_font(chars):
    """True if all chars can be drawn with the built-in Latin-1 Helvetica."""
    try:
        "".join(chars).encode("cp1252")
        return True
    except UnicodeEncodeError:
        return False


def _expand(patterns):
    seen = set()
    for pattern in patterns:
        for path in sorted(glob.glob(os.path.expanduser(pattern), recursive=True)):
            if os.path.isfile(path) and path not in seen:
                seen.add(path)
                yield path


def _fontconfig_candidates(code):
    """Ask fontconfig (Linux/macOS with fontconfig) for a font for this language."""
    exe = shutil.which("fc-match")
    if not exe:
        return
    for query in (f":lang={code}", f"sans:lang={code}"):
        try:
            out = subprocess.run(
                [exe, "-f", "%{file}", query],
                capture_output=True, text=True, timeout=10,
            ).stdout.strip()
        except Exception:
            continue
        if out and os.path.isfile(out):
            yield out


def _font_candidates(code, pack):
    local = [os.path.join(glob.escape(LOCAL_FONT_DIR), ext) for ext in ("*.ttf", "*.otf", "*.ttc")]
    yield from _expand(local)
    yield from _expand(pack.get("fonts", []))
    yield from _expand(GENERIC_FONT_PATTERNS)
    yield from _fontconfig_candidates(code)


def _font_usable(path, required):
    """True if the font has every required glyph AND PyMuPDF can embed it."""
    try:
        font = fitz.Font(fontfile=path)
        if _missing_glyphs(font, required):
            return False
        probe = fitz.open()
        probe.new_page().insert_text((10, 20), "A", fontname="probe", fontfile=path)
        probe.close()
        return True
    except Exception:
        return False


class Typeface:
    """
    Measures and draws text with one font, hiding the differences between the
    built-in fonts, font files, and scripts that need text shaping.
    """

    def __init__(self, kind, ref, shaped=False):
        self.kind = kind      # "builtin" or "file"
        self.ref = ref        # PyMuPDF font name, or path of the font file
        self.shaped = shaped  # draw through the HTML engine (complex scripts)
        self._font = None
        self._alias = "MTFont"
        self._archive = None
        self._story_font = None

        if kind == "file":
            self._font = fitz.Font(fontfile=ref)
            if shaped:
                # The HTML engine loads fonts through an archive (a folder).
                ext = os.path.splitext(ref)[1].lower() or ".ttf"
                self._story_font = f"shapedfont{ext}"
                shutil.copyfile(ref, os.path.join(TEMP_DIR, self._story_font))
                self._archive = fitz.Archive(TEMP_DIR)

    def width(self, text, size):
        if self.kind == "builtin":
            return fitz.get_text_length(text, fontname=self.ref, fontsize=size)
        return self._font.text_length(text, fontsize=size)

    def draw(self, page, x, y, text, size, color=(0, 0, 0)):
        """Draw `text` with its baseline at y and its left edge at x."""
        if self.kind == "builtin":
            page.insert_text((x, y), text, fontsize=size, fontname=self.ref, color=color)
        elif self.shaped:
            self._draw_shaped(page, x, y, text, size, color)
        else:
            page.insert_text(
                (x, y), text, fontsize=size,
                fontname=self._alias, fontfile=self.ref, color=color,
            )

    def _draw_shaped(self, page, x, y, text, size, color):
        # insert_text() does not shape complex scripts (matras and conjuncts
        # would come out wrong), so these go through the HTML engine, which
        # does. Wrapping is already done by the caller, hence nowrap.
        width = self.width(text, size) * 1.25 + size
        rect = fitz.Rect(x, y - size * 1.0, x + width, y + size * 0.8)
        r, g, b = (round(c * 255) for c in color)
        css = (
            "@font-face { font-family: tf; src: url(%s); } "
            "body, p { margin: 0; padding: 0; font-family: tf; font-size: %.2fpt; "
            "line-height: 1.2; white-space: nowrap; color: rgb(%d, %d, %d); }"
        ) % (self._story_font, size, r, g, b)
        page.insert_htmlbox(
            rect, "<p>%s</p>" % html.escape(text), css=css, archive=self._archive
        )


def build_typeface(code, pack, strings, font_override=None):
    """Pick the font used for the whole PDF (see the notes above)."""
    shaped = bool(pack.get("shaped"))
    required = _chars_needed(strings)

    def check_latin(path):
        # Not fatal: only matters for the Latin letters in filenames.
        try:
            missing = _missing_glyphs(fitz.Font(fontfile=path), string.ascii_letters)
        except Exception:
            return
        if missing:
            warn(
                f"font '{os.path.basename(path)}' has no Latin letters, so filenames "
                "may show up as empty boxes. Use --font with a font that covers both."
            )

    if font_override:
        if not os.path.isfile(font_override):
            sys.exit(f"Error: font file not found: {font_override}")
        try:
            font = fitz.Font(fontfile=font_override)
        except Exception as e:
            sys.exit(f"Error: cannot load font '{font_override}': {e}")
        missing = _missing_glyphs(font, required)
        if missing:
            warn(
                f"font '{os.path.basename(font_override)}' lacks these characters: "
                + " ".join(missing)
            )
        check_latin(font_override)
        return Typeface("file", font_override, shaped)

    if _fits_builtin_font(required):
        return Typeface("builtin", "helv")

    for path in _font_candidates(code, pack):
        if _font_usable(path, required):
            check_latin(path)
            return Typeface("file", path, shaped)

    if pack.get("builtin_font"):
        warn(
            f"no suitable font found for '{code}'; falling back to the built-in "
            f"'{pack['builtin_font']}' font, which is not embedded in the PDF."
        )
        return Typeface("builtin", pack["builtin_font"])

    hint = pack.get("font_hint", "Noto Sans")
    sys.exit(
        f"Error: no installed font covers the '{code}' text.\n"
        f"Install one (for example {hint}) and either put the .ttf/.otf file in\n"
        f"  {LOCAL_FONT_DIR}\n"
        "or pass it explicitly with --font /path/to/font.ttf"
    )


def configure_language(code, font_override=None):
    """Select the PDF language: loads its strings and picks its font."""
    global STRINGS, FACE
    pack = LANGUAGES[code]
    STRINGS = {**LANGUAGES[FALLBACK_LANGUAGE]["strings"], **pack["strings"]}
    FACE = build_typeface(code, pack, STRINGS, font_override)


# ---------------------------------------------------------------------------
# Layout helpers
# ---------------------------------------------------------------------------

def natural_key(s):
    """Sort key so 'file2' comes before 'file10'."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def wrap_text(text, fontsize, max_width):
    """Break text into lines that each fit within max_width."""
    words = text.split(" ")
    lines = []
    current = ""

    def fits(s):
        return FACE.width(s, fontsize) <= max_width

    for word in words:
        candidate = f"{current} {word}".strip()
        if fits(candidate):
            current = candidate
        else:
            if current:
                lines.append(current)
            # word itself might be longer than max_width (long filename, or
            # a language written without spaces such as Chinese or Japanese)
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

    lines = wrap_text(text, fontsize, max_width)
    line_height = fontsize * 1.3

    cursor_y = y
    for line in lines:
        FACE.draw(page, MARGIN, cursor_y, line, fontsize)
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
    label_text = tr("video_first_frame", label=label) if is_video else label
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
    big_lines = wrap_text(tr("pdf_contents_of", label=label), big_fontsize, max_width)
    big_line_height = big_fontsize * 1.3

    # Vertically center the (possibly multiline) big text in the space
    # below the header label.
    block_height = len(big_lines) * big_line_height
    top_bound = MARGIN + label_height + 20
    bottom_bound = PAGE_HEIGHT - MARGIN
    start_y = top_bound + max(0, (bottom_bound - top_bound - block_height) / 2)

    cursor_y = start_y
    for line in big_lines:
        FACE.draw(title_page, MARGIN, cursor_y, line, big_fontsize)
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


def draw_centered_lines(page, lines, top_y, fontsize, line_height, color):
    """Draw each line horizontally centered on the page, starting at top_y."""
    for i, line in enumerate(lines):
        x = (PAGE_WIDTH - FACE.width(line, fontsize)) / 2
        FACE.draw(page, x, top_y + i * line_height, line, fontsize, color)


def add_video_caption(page, y):
    """Draw a play-button icon plus a centered 'video clip' caption at
    vertical position y. Returns the height the caption block occupied."""
    icon_radius = 12
    gap_after_icon = 8
    caption_fontsize = 12
    line_height = caption_fontsize * 1.4

    max_width = PAGE_WIDTH - 2 * MARGIN
    lines = wrap_text(tr("video_caption"), caption_fontsize, max_width)

    icon_cy = y + icon_radius
    draw_video_badge(page, PAGE_WIDTH / 2, icon_cy, radius=icon_radius)

    text_top = icon_cy + icon_radius + gap_after_icon
    draw_centered_lines(page, lines, text_top, caption_fontsize, line_height, (0.25, 0.25, 0.25))

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
    note_lines = wrap_text(tr("audio_note"), note_fontsize, max_width)
    line_height = note_fontsize * 1.4
    text_top = cy + 90
    draw_centered_lines(page, note_lines, text_top, note_fontsize, line_height, (0.25, 0.25, 0.25))


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


def truncate_to_width(text, fontsize, max_width):
    """Truncate text with an ellipsis so it fits within max_width."""
    if FACE.width(text, fontsize) <= max_width:
        return text
    for i in range(len(text), 0, -1):
        candidate = text[:i].rstrip() + "\u2026"
        if FACE.width(candidate, fontsize) <= max_width:
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
        title_text = tr("toc_title") if is_first else tr("toc_title_continued")
        FACE.draw(page, MARGIN, MARGIN + TOC_TITLE_FONTSIZE, title_text, TOC_TITLE_FONTSIZE)

        cursor_y = MARGIN + _toc_title_block_height()
        for label, start_idx in chunk:
            page_num = start_idx + page_offset + 1
            page_num_text = str(page_num)
            trunc_label = truncate_to_width(label, TOC_FONTSIZE, label_max_width)

            FACE.draw(page, MARGIN, cursor_y + TOC_FONTSIZE, trunc_label, TOC_FONTSIZE)
            num_width = FACE.width(page_num_text, TOC_FONTSIZE)
            FACE.draw(
                page, PAGE_WIDTH - MARGIN - num_width, cursor_y + TOC_FONTSIZE,
                page_num_text, TOC_FONTSIZE,
            )
            cursor_y += TOC_LINE_HEIGHT


def add_page_numbers(doc):
    """Stamp the translated 'Page N of TOTAL' centered in the bottom margin of every page."""
    fontsize = 9
    total = doc.page_count
    y = PAGE_HEIGHT - 20

    for i in range(total):
        page = doc[i]
        text = tr("page_x_of_y", page=i + 1, total=total)
        x = (PAGE_WIDTH - FACE.width(text, fontsize)) / 2
        FACE.draw(page, x, y, text, fontsize, (0.4, 0.4, 0.4))


# ---------------------------------------------------------------------------
# Command line
# ---------------------------------------------------------------------------

def print_languages():
    print("Available languages (use with --lang):")
    for code in sorted(LANGUAGES):
        default = "  (default)" if code == DEFAULT_LANGUAGE else ""
        print(f"  {code:<4}{LANGUAGES[code]['name']}{default}")


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        prog="mtpdf.py",
        description=(
            "Combine a folder of PDFs, images, videos, spreadsheets and audio "
            "files into a single PDF, labeling each item with its filename."
        ),
    )
    parser.add_argument("input_folder", nargs="?", help="folder containing the files to merge")
    parser.add_argument("output", nargs="?", help="path of the PDF to create")
    parser.add_argument(
        "-l", "--lang",
        type=language_code,
        default=DEFAULT_LANGUAGE,
        metavar="CODE",
        help=(
            "language of the text written into the PDF "
            f"(default: {DEFAULT_LANGUAGE}; available: {', '.join(sorted(LANGUAGES))}). "
            "Regional forms such as zh-CN are accepted."
        ),
    )
    parser.add_argument(
        "--font",
        metavar="FILE",
        help="use this .ttf/.otf font file for all text (needed if no suitable font is auto-detected)",
    )
    parser.add_argument(
        "--list-languages",
        action="store_true",
        help="list the available languages and exit",
    )
    args = parser.parse_args(argv)

    if args.list_languages:
        print_languages()
        sys.exit(0)
    if not args.input_folder or not args.output:
        parser.error("the following arguments are required: input_folder, output")
    return args


def main():
    args = parse_args()
    input_folder = args.input_folder
    output_path = args.output

    if not os.path.isdir(input_folder):
        print(f"Error: {input_folder} is not a folder")
        sys.exit(1)

    configure_language(args.lang, args.font)

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

    print(
        f"Found {len(files)} files. Merging... "
        f"(PDF language: {LANGUAGES[args.lang]['name']} [{args.lang}])"
    )

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

        if FACE.kind == "file":
            # Embedded font files can be large (CJK fonts especially) and are
            # added once per page. Keep only the glyphs actually used, and
            # merge duplicate objects. Both are optional size optimizations.
            try:
                final_doc.subset_fonts()
            except Exception:
                pass
            final_doc.save(output_path, garbage=4, deflate=True)
        else:
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
