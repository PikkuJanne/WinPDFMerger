"""Independent native PDFium checks for numbered fixtures; never an app engine."""

from __future__ import annotations

import argparse
from contextlib import closing
import hashlib
import json
import math
import re
from pathlib import Path

import pypdfium2 as pdfium


PYPDFIUM_VERSION = "5.13.0"
PDFIUM_VERSION = "153.0.7999.0"
DEFAULT_MANIFEST = Path(__file__).resolve().parents[2] / "tests" / "fixtures" / "numbered" / "manifest.json"
IDENTIFIER = re.compile(r"T03-[0-9]{2}-P[0-9]{2}")


def require_pinned_oracle() -> None:
    if str(pdfium.PYPDFIUM_INFO) != PYPDFIUM_VERSION or str(pdfium.PDFIUM_INFO) != PDFIUM_VERSION:
        raise RuntimeError(f"Oracle requires pypdfium2 {PYPDFIUM_VERSION} / PDFium {PDFIUM_VERSION}.")


def validate_identifiers(identifiers: object) -> None:
    if not isinstance(identifiers, list) or not identifiers or any(not isinstance(value, str) or not IDENTIFIER.fullmatch(value) for value in identifiers):
        raise ValueError("Expected a nonempty list of T03 page identifiers.")
    if len(identifiers) != len(set(identifiers)):
        raise ValueError("Page identifiers must be unique.")


def load_manifest(path: Path) -> dict:
    """Reject empty or inconsistent expectations before inspecting any PDF."""
    manifest = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict) or type(manifest.get("schema_version")) is not int or manifest["schema_version"] != 1:
        raise ValueError("Unsupported fixture manifest schema.")
    fixtures = manifest.get("fixtures")
    if not isinstance(fixtures, list) or not fixtures:
        raise ValueError("Fixture manifest must contain a nonempty corpus.")
    names = []
    identifiers = []
    for fixture in fixtures:
        if not isinstance(fixture, dict):
            raise ValueError("Each fixture must be an object.")
        name = fixture.get("file")
        if not isinstance(name, str) or not name.lower().endswith(".pdf") or Path(name).name != name or any(character in name for character in '\\/:*?"<>|\0'):
            raise ValueError("Fixture name must be a local PDF basename.")
        if re.fullmatch(r"(?i)(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?", name):
            raise ValueError("Fixture name must not identify a Windows device.")
        if name.lower() in [previous.lower() for previous in names]:
            raise ValueError("Fixture names must be unique, ignoring case.")
        names.append(name)
        if type(fixture.get("bytes")) is not int or fixture["bytes"] <= 0:
            raise ValueError("Fixture byte length must be a positive integer.")
        if not isinstance(fixture.get("sha256"), str) or not re.fullmatch(r"[0-9a-f]{64}", fixture["sha256"]):
            raise ValueError("Fixture SHA-256 must contain 64 lowercase hexadecimal digits.")
        page_identifiers = fixture.get("page_identifiers")
        validate_identifiers(page_identifiers)
        if type(fixture.get("page_count")) is not int or fixture["page_count"] <= 0 or fixture["page_count"] != len(page_identifiers):
            raise ValueError("Fixture page count must be positive and match its identifiers.")
        size = fixture.get("page_size_points")
        if not isinstance(size, list) or len(size) != 2 or any(type(value) not in (int, float) or not math.isfinite(value) or value <= 0 for value in size):
            raise ValueError("Fixture dimensions must be two finite positive numbers.")
        identifiers.extend(page_identifiers)
    validate_identifiers(identifiers)
    if manifest.get("merge_order") != names:
        raise ValueError("Manifest input order differs.")
    if type(manifest.get("expected_merged_page_count")) is not int or manifest["expected_merged_page_count"] != len(identifiers):
        raise ValueError("Manifest merged page count differs.")
    if manifest.get("expected_merged_page_identifiers") != identifiers:
        raise ValueError("Manifest merged identifiers differ.")
    return manifest


def inspect_pdf(path: Path, expected_identifiers: list[str], render_directory: Path | None = None) -> dict:
    """Count/read pages with native PDFium, independent of ReportLab generation."""
    validate_identifiers(expected_identifiers)
    require_pinned_oracle()
    identifiers = []
    with pdfium.PdfDocument(path) as document:
        if len(document) != len(expected_identifiers):
            raise ValueError(f"Page count differs: expected {len(expected_identifiers)}, got {len(document)}.")
        sizes = []
        for index in range(len(document)):
            with closing(document[index]) as page:
                sizes.append(list(page.get_size()))
                with closing(page.get_textpage()) as text_page:
                    found = IDENTIFIER.findall(text_page.get_text_range())
                expected = expected_identifiers[index]
                if found != [expected]:
                    raise ValueError(f"Page {index + 1} identifier differs: expected {expected}, got {found}.")
                identifiers.append(found[0])
                if render_directory is not None:
                    render_directory.mkdir(parents=True, exist_ok=True)
                    with closing(page.render(scale=1.5)) as bitmap:
                        bitmap.to_pil().save(render_directory / f"{path.stem}-page-{index + 1}.png")
    return {"page_count": len(identifiers), "page_identifiers": identifiers, "page_sizes_points": sizes}


def verify_corpus(manifest_path: Path, render_directory: Path | None = None) -> dict:
    """Check committed hashes, real page totals and identifiers; no mocked parser."""
    manifest = load_manifest(manifest_path)
    results = []
    all_identifiers = []
    total = 0
    for fixture in manifest["fixtures"]:
        name = fixture["file"]
        path = manifest_path.parent / name
        content = path.read_bytes()
        if len(content) != fixture["bytes"] or hashlib.sha256(content).hexdigest() != fixture["sha256"]:
            raise ValueError(f"Fixture hash/length differs: {name}.")
        result = inspect_pdf(path, fixture["page_identifiers"], render_directory)
        if result["page_count"] != fixture["page_count"]:
            raise ValueError(f"Manifest page total differs: {name}.")
        if any(size != fixture["page_size_points"] for size in result["page_sizes_points"]):
            raise ValueError(f"Page size differs: {name}.")
        results.append({"file": name, "sha256": fixture["sha256"], **result})
        total += result["page_count"]
        all_identifiers.extend(result["page_identifiers"])
    if manifest["merge_order"] != [result["file"] for result in results]:
        raise ValueError("Manifest input order differs.")
    if total != manifest["expected_merged_page_count"] or all_identifiers != manifest["expected_merged_page_identifiers"]:
        raise ValueError("Manifest merged expectation differs.")
    return {
        "evidence_class": "fixture_native_inspection",
        "product_merge_executed": False,
        "pdftk_executed": False,
        "ghostscript_executed": False,
        "pypdfium2_version": str(pdfium.PYPDFIUM_INFO),
        "pdfium_version": str(pdfium.PDFIUM_INFO),
        "fixtures": results,
        "total_pages": total,
        "page_identifiers": all_identifiers,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST)
    parser.add_argument("--render-dir", type=Path, help="Optional development PNG directory; no output writes otherwise.")
    parser.add_argument("--pdf", type=Path, help="Inspect a real merge result against the manifest's merged identifiers.")
    arguments = parser.parse_args()
    if arguments.pdf is None:
        result = verify_corpus(arguments.manifest, arguments.render_dir)
    else:
        manifest = load_manifest(arguments.manifest)
        result = inspect_pdf(arguments.pdf, manifest["expected_merged_page_identifiers"], arguments.render_dir)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
