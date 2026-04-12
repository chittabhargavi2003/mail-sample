"""
MSG to PDF Converter
====================
Converts Microsoft Outlook .msg email files to PDF.

For each .msg file the script produces a PDF that contains:
  1. Email header block  (From / Sent / To / CC / Subject / Attachments)
  2. Email body          (HTML decoded to plain-text, or raw plain-text)
  3. Inline attachments  (PDF attachments merged page-by-page; images embedded
                          on a dedicated page; other file types listed on a
                          placeholder page)

Attachment resolution order
---------------------------
  a) Attachments embedded directly inside the .msg file (most common).
  b) Companion files found next to the .msg file whose names match attachment
     references in the email body (e.g. the .msg was saved alongside its PDF).
  c) Files found in the directory supplied with --attachments-dir / -a.

Usage
-----
    # Convert a single file
    python msg_to_pdf.py email.msg

    # Convert several files
    python msg_to_pdf.py file1.msg file2.msg

    # Convert every .msg file in a directory
    python msg_to_pdf.py /path/to/folder

    # Specify a custom output directory
    python msg_to_pdf.py /path/to/folder --output /path/to/output

    # Specify where to look for companion attachment files
    python msg_to_pdf.py /path/to/folder --attachments-dir /path/to/attachments

Dependencies (install via pip)
-------------------------------
    pip install extract-msg reportlab pillow pypdf html2text beautifulsoup4
"""

from __future__ import annotations

import argparse
import io
import logging
import os
import re
import sys
from pathlib import Path
from typing import List, Optional

import extract_msg
import html2text
from bs4 import BeautifulSoup
from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    HRFlowable,
    Image as RLImage,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Page geometry
# ---------------------------------------------------------------------------
PAGE_WIDTH, PAGE_HEIGHT = A4
MARGIN = 2 * cm
CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN

# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
_BASE_STYLES = getSampleStyleSheet()

STYLE_HEADER_TITLE = ParagraphStyle(
    "HeaderTitle",
    parent=_BASE_STYLES["Normal"],
    fontSize=13,
    fontName="Helvetica-Bold",
    spaceAfter=4,
)

STYLE_HEADER_LABEL = ParagraphStyle(
    "HeaderLabel",
    parent=_BASE_STYLES["Normal"],
    fontSize=10,
    fontName="Helvetica-Bold",
    leading=14,
)

STYLE_HEADER_VALUE = ParagraphStyle(
    "HeaderValue",
    parent=_BASE_STYLES["Normal"],
    fontSize=10,
    fontName="Helvetica",
    leading=14,
)

STYLE_BODY = ParagraphStyle(
    "Body",
    parent=_BASE_STYLES["Normal"],
    fontSize=10,
    fontName="Helvetica",
    leading=14,
    spaceAfter=4,
    alignment=TA_LEFT,
)

STYLE_ATTACHMENT_TITLE = ParagraphStyle(
    "AttachmentTitle",
    parent=_BASE_STYLES["Normal"],
    fontSize=11,
    fontName="Helvetica-Bold",
    spaceAfter=6,
    spaceBefore=10,
)

STYLE_PLACEHOLDER = ParagraphStyle(
    "Placeholder",
    parent=_BASE_STYLES["Normal"],
    fontSize=10,
    fontName="Helvetica-Oblique",
    textColor=colors.grey,
)

# ---------------------------------------------------------------------------
# Module-level constants shared across helpers
# ---------------------------------------------------------------------------

# Regex pattern for filenames with common attachment extensions referenced in
# email body text.  Uses negative look-around so it does not match filenames
# that are part of a longer word or a domain segment.
_FILENAME_PATTERN = re.compile(
    r"(?<!\w)([\w\-. ]+\.(?:pdf|docx?|xlsx?|png|jpe?g|gif|txt|csv))(?!\w)",
    re.IGNORECASE,
)

# Image file extensions that can be embedded as a page in the output PDF.
_IMAGE_EXTENSIONS = frozenset(
    {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp"}
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _html_to_text(html: bytes | str) -> str:
    """
    Convert HTML bytes/string to clean plain text.

    Removes the injected Kutools-for-Outlook 'Attachments:' div (we already
    list attachments in the header block) before conversion.
    """
    if isinstance(html, bytes):
        html = html.decode("utf-8", errors="replace")

    # Strip the Kutools injection div and any other purely decorative divs
    # that only contain attachment metadata.
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup.find_all("div", class_="KutoolsforOutlook"):
        tag.decompose()
    for tag in soup.find_all("p", class_="KutoolsforOutlook"):
        tag.decompose()
    # Also strip the injected header div (already shown in header block)
    for tag in soup.find_all(id="injectedHeader"):
        tag.decompose()
    clean_html = str(soup)

    converter = html2text.HTML2Text()
    converter.ignore_links = True
    converter.ignore_images = True
    converter.ignore_emphasis = True
    converter.body_width = 0  # no forced line-wrapping
    return converter.handle(clean_html).strip()


def _clean_body(msg: extract_msg.Message) -> str:
    """Return the best available body text for the message."""
    if msg.htmlBody:
        return _html_to_text(msg.htmlBody)
    if msg.body:
        return (msg.body or "").strip()
    return ""


def _escape_xml(text: str) -> str:
    """Escape special characters for ReportLab Paragraph markup."""
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    return text


def _format_date(dt) -> str:
    """Format a datetime object as 'Weekday, Month DD, YYYY HH:MM AM/PM'."""
    if dt is None:
        return ""
    try:
        return dt.strftime("%A, %B %d, %Y %I:%M %p").replace(" 0", " ")
    except Exception:
        return str(dt)


def _recipient_display_names(msg: extract_msg.Message) -> List[str]:
    """Return a list of display names for all TO recipients."""
    names: List[str] = []
    for r in msg.recipients:
        if r.type == 1:  # MAPI_TO
            names.append(r.name or r.email or "")
    return names


def _attachment_names(msg: extract_msg.Message) -> List[str]:
    """Return display names for all file attachments (embedded + body-referenced)."""
    names: List[str] = []
    for att in msg.attachments:
        name = att.longFilename or att.shortFilename or "attachment"
        names.append(name)
    # Also surface attachment references in the plain-text body that are not
    # already covered by an embedded attachment entry.
    if msg.body:
        for f in _FILENAME_PATTERN.findall(msg.body):
            fname = os.path.basename(f.replace("\\", "/"))
            if fname and fname not in names:
                names.append(fname)
    return names


def _extract_referenced_filenames(msg: extract_msg.Message) -> List[str]:
    """
    Return the list of filenames referenced in the message body but NOT
    embedded as proper attachments.  These are candidates for companion-file
    lookup on disk.
    """
    embedded = {
        att.longFilename or att.shortFilename or ""
        for att in msg.attachments
    }
    referenced: List[str] = []
    if msg.body:
        for f in _FILENAME_PATTERN.findall(msg.body):
            fname = os.path.basename(f.replace("\\", "/"))
            if fname and fname not in embedded and fname not in referenced:
                referenced.append(fname)
    return referenced


def _find_companion_file(filename: str, search_dirs: List[Path]) -> Optional[Path]:
    """
    Search *search_dirs* (in order) for a file matching *filename*.
    Returns the first match found, or None.
    """
    for d in search_dirs:
        candidate = d / filename
        if candidate.is_file():
            return candidate
    return None


# ---------------------------------------------------------------------------
# Build the "header + body" portion as an in-memory PDF
# ---------------------------------------------------------------------------

def _build_header_body_pdf(msg: extract_msg.Message) -> bytes:
    """
    Render the email header and body into a PDF and return the raw bytes.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=MARGIN,
    )

    story = []

    # --- Primary recipient line (shown at top, like Outlook's print view) ---
    # Matches the Outlook convention of showing the first TO recipient's display
    # name above the header block when printing or exporting an email.
    to_names = _recipient_display_names(msg)
    if to_names:
        primary = to_names[0]
        story.append(Paragraph(_escape_xml(primary), STYLE_HEADER_TITLE))

    # --- Separator ---
    story.append(HRFlowable(width="100%", thickness=1, color=colors.black, spaceAfter=6))

    # --- Header table (label | value) ---
    sender_name = ""
    if msg.sender:
        # Extract display name before the angle-bracket email
        m = re.match(r"^(.*?)\s*<", msg.sender)
        sender_name = m.group(1).strip() if m else msg.sender.strip()

    att_names = _attachment_names(msg)
    att_display = "; ".join(att_names) if att_names else "(none)"

    # Build the header rows; always include Attachments.
    # Optional rows (CC) are omitted when empty.
    optional_rows = [
        ("CC:", msg.cc or ""),
    ]
    header_rows = [
        ("From:", sender_name or msg.sender or ""),
        ("Sent:", _format_date(msg.date)),
        ("To:", msg.to or ""),
    ]
    header_rows += [(lbl, val) for lbl, val in optional_rows if val]
    header_rows += [
        ("Subject:", msg.subject or ""),
        ("Attachments:", att_display),
    ]

    tbl_data = [
        [
            Paragraph(_escape_xml(lbl), STYLE_HEADER_LABEL),
            Paragraph(_escape_xml(val), STYLE_HEADER_VALUE),
        ]
        for lbl, val in header_rows
    ]

    tbl = Table(tbl_data, colWidths=[2.5 * cm, CONTENT_WIDTH - 2.5 * cm])
    tbl.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    story.append(tbl)
    story.append(Spacer(1, 0.4 * cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey, spaceAfter=8))

    # --- Body ---
    body_text = _clean_body(msg)
    if body_text:
        for line in body_text.splitlines():
            stripped = line.strip()
            if stripped:
                story.append(Paragraph(_escape_xml(stripped), STYLE_BODY))
            else:
                story.append(Spacer(1, 6))

    doc.build(story)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Attachment rendering helpers
# ---------------------------------------------------------------------------

def _attachment_to_pdf_bytes(att) -> Optional[bytes]:
    """
    Try to convert a single attachment into PDF bytes suitable for merging.

    Returns:
        bytes  – PDF bytes if the attachment is a PDF file
        None   – for non-PDF attachments (caller handles image / other)
    """
    name = (att.longFilename or att.shortFilename or "").lower()
    data = att.data
    if data and name.endswith(".pdf"):
        return data
    return None


def _build_image_page_from_bytes(image_data: bytes, display_name: str) -> Optional[bytes]:
    """
    Render image bytes as a single-page PDF (A4) and return the bytes.
    Returns None on failure.
    """
    try:
        img = Image.open(io.BytesIO(image_data))
        img_w, img_h = img.size

        max_w = PAGE_WIDTH - 2 * MARGIN
        max_h = PAGE_HEIGHT - 2 * MARGIN
        scale = min(max_w / img_w, max_h / img_h, 1.0)
        draw_w = img_w * scale
        draw_h = img_h * scale

        buf = io.BytesIO()
        doc = SimpleDocTemplate(
            buf,
            pagesize=A4,
            leftMargin=MARGIN,
            rightMargin=MARGIN,
            topMargin=MARGIN,
            bottomMargin=MARGIN,
        )
        story = [
            Paragraph(_escape_xml(display_name), STYLE_ATTACHMENT_TITLE),
            Spacer(1, 0.3 * cm),
            RLImage(io.BytesIO(image_data), width=draw_w, height=draw_h),
        ]
        doc.build(story)
        return buf.getvalue()
    except Exception as exc:
        log.warning("Could not render image '%s': %s", display_name, exc)
        return None


def _build_image_page(att) -> Optional[bytes]:
    """
    Render an image attachment as a single-page PDF (A4) and return the bytes.
    Returns None if the attachment is not a supported image.
    """
    name = (att.longFilename or att.shortFilename or "").lower()
    if Path(name).suffix not in _IMAGE_EXTENSIONS:
        return None

    data = att.data
    if not data:
        return None

    display_name = att.longFilename or att.shortFilename or "image"
    return _build_image_page_from_bytes(data, display_name)


def _build_placeholder_page(att) -> bytes:
    """
    Build a simple placeholder page for attachments that cannot be rendered
    (non-PDF, non-image) showing the filename and size.
    """
    name = att.longFilename or att.shortFilename or "attachment"
    size = len(att.data) if att.data else 0

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=MARGIN,
        rightMargin=MARGIN,
        topMargin=MARGIN,
        bottomMargin=MARGIN,
    )
    story = [
        Paragraph("Attachment", STYLE_ATTACHMENT_TITLE),
        HRFlowable(width="100%", thickness=0.5, color=colors.lightgrey, spaceAfter=8),
        Paragraph(_escape_xml(f"File name : {name}"), STYLE_BODY),
        Paragraph(_escape_xml(f"File size : {size:,} bytes"), STYLE_BODY),
        Spacer(1, 0.5 * cm),
        Paragraph(
            "This attachment type cannot be rendered inline.",
            STYLE_PLACEHOLDER,
        ),
    ]
    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Core conversion function
# ---------------------------------------------------------------------------

def msg_to_pdf(
    msg_path: str | Path,
    output_path: Optional[str | Path] = None,
    attachments_dir: Optional[str | Path] = None,
) -> Path:
    """
    Convert a single .msg file to a PDF file.

    Parameters
    ----------
    msg_path        : path to the source .msg file
    output_path     : destination PDF path (optional).  When omitted, the PDF is
                      placed in the same directory as the .msg file with the same
                      stem (e.g. email.msg → email.pdf).
    attachments_dir : optional extra directory to search for companion attachment
                      files that are referenced in the email body but not embedded
                      in the .msg file itself.

    Returns
    -------
    Path of the generated PDF file.
    """
    msg_path = Path(msg_path).resolve()
    if output_path is None:
        output_path = msg_path.with_suffix(".pdf")
    else:
        output_path = Path(output_path).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)

    log.info("Converting '%s' → '%s'", msg_path.name, output_path.name)

    msg = extract_msg.Message(str(msg_path))

    # Directories to search when resolving companion attachment files.
    # Priority: msg directory first, then user-supplied attachments_dir.
    search_dirs: List[Path] = [msg_path.parent]
    if attachments_dir is not None:
        extra = Path(attachments_dir).resolve()
        if extra not in search_dirs:
            search_dirs.append(extra)

    # 1. Build the header + body PDF
    header_body_bytes = _build_header_body_pdf(msg)

    # 2. Merge attachments
    writer = PdfWriter()

    # Add all pages from the header/body PDF
    reader = PdfReader(io.BytesIO(header_body_bytes))
    for page in reader.pages:
        writer.add_page(page)

    # 2a. Process each *embedded* attachment
    for att in msg.attachments:
        att_name = att.longFilename or att.shortFilename or "attachment"

        # Try PDF
        pdf_bytes = _attachment_to_pdf_bytes(att)
        if pdf_bytes:
            log.info("  Embedding PDF attachment: %s", att_name)
            try:
                att_reader = PdfReader(io.BytesIO(pdf_bytes))
                for page in att_reader.pages:
                    writer.add_page(page)
                continue
            except Exception as exc:
                log.warning("  Could not parse PDF attachment '%s': %s", att_name, exc)

        # Try image
        img_bytes = _build_image_page(att)
        if img_bytes:
            log.info("  Embedding image attachment: %s", att_name)
            img_reader = PdfReader(io.BytesIO(img_bytes))
            for page in img_reader.pages:
                writer.add_page(page)
            continue

        # Fallback: placeholder page
        log.info("  Adding placeholder for attachment: %s", att_name)
        placeholder_bytes = _build_placeholder_page(att)
        ph_reader = PdfReader(io.BytesIO(placeholder_bytes))
        for page in ph_reader.pages:
            writer.add_page(page)

    # 2b. Resolve *companion* files referenced in the body but not embedded.
    #     Search alongside the .msg file and in any user-supplied directory.
    referenced = _extract_referenced_filenames(msg)
    for ref_name in referenced:
        companion = _find_companion_file(ref_name, search_dirs)
        if companion is None:
            log.debug("  Companion file not found on disk: %s", ref_name)
            continue

        ext = companion.suffix.lower()
        log.info("  Merging companion file: %s", companion.name)

        if ext == ".pdf":
            try:
                att_reader = PdfReader(str(companion))
                for page in att_reader.pages:
                    writer.add_page(page)
            except Exception as exc:
                log.warning("  Could not merge companion PDF '%s': %s", companion.name, exc)

        elif ext in _IMAGE_EXTENSIONS:
            img_bytes = _build_image_page_from_bytes(companion.read_bytes(), companion.name)
            if img_bytes:
                img_reader = PdfReader(io.BytesIO(img_bytes))
                for page in img_reader.pages:
                    writer.add_page(page)
        else:
            log.info("  Companion file '%s' cannot be rendered inline (unsupported type)", companion.name)

    # 3. Write the final PDF
    with open(output_path, "wb") as f:
        writer.write(f)

    log.info("  ✓ Saved: %s", output_path)
    return output_path


# ---------------------------------------------------------------------------
# Batch conversion
# ---------------------------------------------------------------------------

def convert_directory(
    input_dir: str | Path,
    output_dir: Optional[str | Path] = None,
    attachments_dir: Optional[str | Path] = None,
) -> List[Path]:
    """
    Convert all .msg files found in *input_dir* (non-recursive).

    Parameters
    ----------
    input_dir       : directory containing .msg files
    output_dir      : directory for output PDFs (defaults to *input_dir*)
    attachments_dir : optional extra directory to search for companion
                      attachment files (see :func:`msg_to_pdf`)

    Returns
    -------
    List of generated PDF paths.
    """
    input_dir = Path(input_dir).resolve()
    msg_files = sorted(input_dir.glob("*.msg"))
    if not msg_files:
        log.warning("No .msg files found in '%s'", input_dir)
        return []

    results: List[Path] = []
    for msg_file in msg_files:
        if output_dir is not None:
            pdf_path = Path(output_dir) / (msg_file.stem + ".pdf")
        else:
            pdf_path = None
        results.append(msg_to_pdf(msg_file, pdf_path, attachments_dir=attachments_dir))
    return results


# ---------------------------------------------------------------------------
# CLI entry-point
# ---------------------------------------------------------------------------

def _parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Convert Outlook .msg files to PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        metavar="INPUT",
        help="One or more .msg files or directories containing .msg files",
    )
    parser.add_argument(
        "--output",
        "-o",
        metavar="DIR",
        help="Output directory for generated PDF files (default: same as each input file)",
    )
    parser.add_argument(
        "--attachments-dir",
        "-a",
        metavar="DIR",
        help=(
            "Directory to search for companion attachment files that are referenced "
            "in the email body but not embedded in the .msg file itself. "
            "The .msg file's own directory is always searched first."
        ),
    )
    return parser.parse_args(argv)


def main(argv=None):
    args = _parse_args(argv)
    output_dir = Path(args.output) if args.output else None
    attachments_dir = Path(args.attachments_dir) if args.attachments_dir else None

    generated: List[Path] = []
    for inp in args.inputs:
        p = Path(inp)
        if p.is_dir():
            generated.extend(convert_directory(p, output_dir, attachments_dir=attachments_dir))
        elif p.is_file() and p.suffix.lower() == ".msg":
            pdf_dest = (output_dir / (p.stem + ".pdf")) if output_dir else None
            generated.append(msg_to_pdf(p, pdf_dest, attachments_dir=attachments_dir))
        else:
            log.error("Skipping '%s': not a .msg file or directory", inp)

    if generated:
        print(f"\nGenerated {len(generated)} PDF file(s):")
        for pdf in generated:
            print(f"  {pdf}")
    else:
        print("No PDF files were generated.")
        sys.exit(1)


if __name__ == "__main__":
    main()
