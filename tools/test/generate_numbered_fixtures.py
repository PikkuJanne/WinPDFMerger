"""Create or check the original, deterministic T03 PDF corpus (development only)."""

from __future__ import annotations

import argparse
import hashlib
import io
import json
from pathlib import Path

import reportlab
from reportlab.pdfgen import canvas


REPORTLAB_VERSION = "4.4.9"
FIXTURE_DIRECTORY = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "numbered"
PAGE_SIZE = (432, 288)
FIXTURES = (
    ("1.pdf", ("T03-01-P01",)),
    ("2.pdf", ("T03-02-P01", "T03-02-P02")),
    ("10.pdf", ("T03-10-P01",)),
)


def make_pdf(filename: str, identifiers: tuple[str, ...]) -> bytes:
    """Use only original text/vector artwork; no external font or image asset."""
    if reportlab.Version != REPORTLAB_VERSION:
        raise RuntimeError(f"Fixture generation requires ReportLab {REPORTLAB_VERSION}.")
    output = io.BytesIO()
    document = canvas.Canvas(
        output,
        pagesize=PAGE_SIZE,
        invariant=1,
        pageCompression=0,
        pdfVersion=(1, 4),
    )
    document.setTitle(f"WinPDFMerger synthetic fixture {filename}")
    document.setAuthor("WinPDFMerger test corpus")
    document.setSubject("Original development-only numbered PDF")
    for number, identifier in enumerate(identifiers, start=1):
        document.setStrokeColorRGB(0.12, 0.24, 0.42)
        document.setLineWidth(2)
        document.rect(20, 20, 392, 248)
        document.setFillColorRGB(0.12, 0.24, 0.42)
        document.setFont("Helvetica-Bold", 28)
        document.drawCentredString(216, 164, identifier)
        document.setFont("Helvetica", 12)
        document.drawCentredString(216, 131, f"File {filename} - page {number} of {len(identifiers)}")
        document.setFont("Helvetica", 9)
        document.drawCentredString(216, 75, "Original synthetic test content. No private document data.")
        document.showPage()
    document.save()
    return output.getvalue()


def expected_files() -> dict[str, bytes]:
    """Return PDF bytes and oracle metadata without filesystem writes."""
    files = {}
    entries = []
    for filename, identifiers in FIXTURES:
        content = make_pdf(filename, identifiers)
        files[filename] = content
        entries.append({
            "file": filename,
            "sha256": hashlib.sha256(content).hexdigest(),
            "bytes": len(content),
            "page_count": len(identifiers),
            "page_identifiers": list(identifiers),
            "page_size_points": list(PAGE_SIZE),
        })
    manifest = {
        "schema_version": 1,
        "provenance": "Original synthetic text and vector artwork created by this repository's generator.",
        "license": "Repository LICENSE; no external document, image or font file included.",
        "generator": "tools/test/generate_numbered_fixtures.py",
        "generator_version": 1,
        "reportlab_version": REPORTLAB_VERSION,
        "fixtures": entries,
        "merge_order": [filename for filename, _ in FIXTURES],
        "expected_merged_page_count": sum(len(identifiers) for _, identifiers in FIXTURES),
        "expected_merged_page_identifiers": [identifier for _, identifiers in FIXTURES for identifier in identifiers],
    }
    files["manifest.json"] = (json.dumps(manifest, indent=2) + "\n").encode("utf-8")
    return files


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=FIXTURE_DIRECTORY)
    parser.add_argument("--write", action="store_true", help="Write the intended corpus; default is read-only byte comparison.")
    arguments = parser.parse_args()
    files = expected_files()
    if arguments.write:
        arguments.output_dir.mkdir(parents=True, exist_ok=True)
        for name, content in files.items():
            (arguments.output_dir / name).write_bytes(content)
        print(f"Wrote {len(FIXTURES)} original PDFs and manifest.")
        return 0
    mismatches = [name for name, content in files.items() if not (arguments.output_dir / name).is_file() or (arguments.output_dir / name).read_bytes() != content]
    if mismatches:
        parser.exit(1, "Fixture bytes differ or are missing: " + ", ".join(mismatches) + "\n")
    print(f"PASS: {len(FIXTURES)} PDFs and manifest reproduce byte-for-byte.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
