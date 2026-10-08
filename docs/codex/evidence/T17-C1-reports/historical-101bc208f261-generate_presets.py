"""Development-only original synthetic preset corpus; never an application engine.

All text, diagrams and image pixels are authored here (CC0-1.0). No document,
network content or user font is read. Exact pinned packages make invariant
ReportLab output reproducible. Outputs must be a new explicitly owned directory.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import random
import sys

import PIL
from PIL import Image, ImageDraw, ImageFont
import reportlab
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas


VERSIONS = {"python": "3.12.14", "reportlab": "4.4.9", "pillow": "12.3.0"}
PAGE_SIZE = (612, 792)
SEED = 170041


def require_versions() -> None:
    actual = {"python": sys.version.split()[0], "reportlab": reportlab.Version,
              "pillow": PIL.__version__}
    if actual != VERSIONS:
        raise RuntimeError(f"Development generator pins differ: {actual}")


def heading(c: Canvas, title: str, identifier: str) -> None:
    c.setFillColorRGB(0.08, 0.18, 0.28)
    c.setFont("Helvetica-Bold", 19)
    c.drawString(36, 750, title)
    c.setFont("Helvetica", 10)
    c.drawString(36, 728, "Original synthetic test content - compare master, screen and ebook at 100%.")
    c.setStrokeColorRGB(0.25, 0.40, 0.52)
    c.line(36, 718, 576, 718)
    c.setFillColorRGB(0.08, 0.18, 0.28)
    c.setFont("Helvetica", 11)
    c.drawString(36, 22, identifier)
    c.setFont("Helvetica", 8)
    c.drawRightString(576, 22, "CC0 original fixture; no forms, signatures or private content")


def vector_page(c: Canvas, identifier: str, title: str) -> None:
    heading(c, title, identifier)
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(36, 690, "Small-print ladder and line detail")
    y = 660
    for points in (12, 10, 8, 7, 6, 5):
        c.setFont("Helvetica", points)
        c.drawString(42, y, f"{points} pt: Invoice 170041 | 0123456789 | AaBb CcDd MmNn | $123.45 | 0O 1Il")
        y -= 34
    c.setFont("Helvetica-Bold", 11)
    c.drawString(36, 430, "Fine rules, grayscale and color detail")
    for index, width in enumerate((0.25, 0.5, 0.75, 1.0, 1.5)):
        y = 405 - index * 19
        c.setLineWidth(width)
        c.setStrokeColorRGB(0, 0, 0)
        c.line(42, y, 356, y)
        c.setFont("Helvetica", 8)
        c.drawString(370, y - 3, f"{width:g} pt rule")
    for index, gray in enumerate((0, 0.15, 0.35, 0.55, 0.75, 0.9)):
        c.setFillColorRGB(gray, gray, gray)
        c.rect(42 + index * 83, 263, 74, 36, fill=1, stroke=0)
    for index, rgb in enumerate(((0.8, 0.1, 0.1), (0.1, 0.55, 0.2), (0.15, 0.3, 0.8),
                                 (0.85, 0.55, 0.1), (0.55, 0.2, 0.65), (0.1, 0.6, 0.65))):
        c.setFillColorRGB(*rgb)
        c.rect(42 + index * 83, 210, 74, 36, fill=1, stroke=0)
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(36, 175, "Original table")
    for row, text in enumerate(("Item        Units        Rate          Total",
                                "Paper       12           3.25          39.00",
                                "Ink         4            17.50         70.00",
                                "Review      1            0.00          0.00")):
        c.setFont("Courier", 8)
        c.drawString(42, 152 - row * 19, text)
    c.showPage()


def scanned_image() -> Image.Image:
    width, height = 2250, 2800
    # Deterministic low-amplitude scanner-like noise; no external source image.
    pixels = random.Random(SEED).randbytes(width * height)
    pixels = bytes(249 + value % 7 for value in pixels)
    image = Image.frombytes("L", (width, height), pixels)
    draw = ImageDraw.Draw(image)
    scale = width / 540
    def text(x: int, y: int, value: str, points: int) -> None:
        draw.text((int(x * scale), int(y * scale)), value, fill=20,
                  font=ImageFont.load_default(size=int(points * scale)))
    text(20, 18, "SCANNED ORIGINAL - raster text and diagram", 15)
    text(20, 48, "Original pixels, deterministic noise, no OCR layer", 10)
    for row, points in enumerate((12, 10, 8, 7, 6)):
        text(20, 90 + row * 37,
             f"{points} pt: Account 170041 / 0123456789 / AaBb MmNn / $123.45", points)
    text(20, 293, "Raster diagram: lines, circles and grayscale", 11)
    for index, line_width in enumerate((1, 2, 3, 4, 6)):
        y = int((325 + index * 17) * scale)
        draw.line((int(20 * scale), y, int(320 * scale), y), fill=10, width=line_width)
    for index in range(5):
        x, y = int((25 + index * 98) * scale), int(435 * scale)
        radius = int(34 * scale)
        draw.ellipse((x, y, x + radius, y + radius), outline=15, width=2)
        draw.rectangle((x, y + 2 * radius, x + radius, y + 3 * radius),
                       fill=int(index * 52))
    text(20, 580, "Check legibility at 100%; smaller bytes alone do not prove fidelity.", 9)
    return image


def scan_page(c: Canvas, image: Image.Image, identifier: str, title: str) -> None:
    heading(c, title, identifier)
    c.drawImage(ImageReader(image), 36, 43, width=540, height=672)
    c.showPage()


def generate(output: Path) -> dict:
    require_versions()
    output.mkdir(parents=True, exist_ok=True)
    image = scanned_image()
    specifications = (
        ("small-print.pdf", ["T03-17-P01"], "vector_small_print"),
        ("scan.pdf", ["T03-17-P02"], "original_scanned_text_diagram"),
        ("mixed.pdf", ["T03-17-P03", "T03-17-P04"], "mixed_vector_raster"),
    )
    fixtures = []
    for name, identifiers, category in specifications:
        path = output / name
        if path.exists():
            raise FileExistsError(f"Refusing to overwrite fixture: {path}")
        c = Canvas(str(path), pagesize=PAGE_SIZE, pageCompression=1, invariant=1)
        c.setTitle(f"Original T17 synthetic {category}")
        c.setAuthor("WinPDFMerger synthetic fixture generator")
        c.setSubject("CC0 original synthetic preset review; development only")
        if category == "vector_small_print":
            vector_page(c, identifiers[0], "Vector small-print specimen")
        elif category == "original_scanned_text_diagram":
            scan_page(c, image, identifiers[0], "Scanned text and diagram specimen")
        else:
            vector_page(c, identifiers[0], "Mixed specimen - vector page")
            scan_page(c, image, identifiers[1], "Mixed specimen - scanned page")
        c.save()
        raw = path.read_bytes()
        fixtures.append({"file": name, "category": category, "bytes": len(raw),
                         "sha256": hashlib.sha256(raw).hexdigest(),
                         "page_count": len(identifiers), "page_identifiers": identifiers,
                         "page_size_points": list(PAGE_SIZE)})
    return {"schema_version": 1, "license": "CC0-1.0", "versions": VERSIONS,
            "provenance": "Original synthetic vector text/diagrams and seeded raster pixels; no external documents, user data, OCR or network.",
            "scan_pixel_dimensions": [2250, 2800], "scan_seed": SEED,
            "generator_sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
            "fixtures": fixtures,
            "scope": "Development fixture recipe and structural expectations; visual preset acceptance requires actual rendered master/derivative review."}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    print(json.dumps(generate(args.output), indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
