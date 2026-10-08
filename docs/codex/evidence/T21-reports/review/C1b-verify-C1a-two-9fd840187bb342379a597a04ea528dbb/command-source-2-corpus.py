"""Reconstruct/inspect the tracked original synthetic corpus; development only.

No downloads, application import, runtime dependency, or user document input.
Materialization writes only a new owned directory under tests/.work. Verification
is read-only. Native application acceptance is a separate Pester suite.
"""
from __future__ import annotations

import argparse
from contextlib import closing
import ctypes
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import re
import subprocess
import sys

REPO = Path(__file__).resolve().parents[2]
CATALOG = REPO / "tests/fixtures/corpus.json"
PINS = {"python": "3.12.14", "reportlab": "4.4.9", "pypdf": "6.10.0",
        "pillow": "12.3.0", "pypdfium2": "5.13.0", "pdfium": "153.0.7999.0"}
IDENTIFIER = re.compile(r"T03-[0-9]{2}-P[0-9]{2}|T19-[AB]-P[12]")
GROUPS = {
    "numbered": ("tools/test/generate_numbered_fixtures.py", "tests/fixtures/numbered/manifest.json"),
    "presets": ("tests/fixtures/presets/generate_presets.py", "tests/fixtures/presets/manifest.json"),
    "features": ("tests/fixtures/features/generate_features.py", "tests/fixtures/features/manifest.json"),
    "envelopes": ("tools/test/generate_pdf_envelope_fixtures.py", None),
}


def sha256(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def same_json(actual, expected) -> bool:
    """JSON numbers and booleans must not compare as interchangeable types."""
    return json.dumps(actual, sort_keys=True) == json.dumps(expected, sort_keys=True)


def load_recipe(group: str):
    specification = importlib.util.spec_from_file_location("corpus_" + group, REPO / GROUPS[group][0])
    module = importlib.util.module_from_spec(specification)
    specification.loader.exec_module(module)
    return module


def versions() -> dict:
    import reportlab
    import pypdf
    import pypdfium2
    import PIL
    actual = {"python": sys.version.split()[0], "reportlab": reportlab.Version,
              "pypdf": pypdf.__version__, "pillow": PIL.__version__,
              "pypdfium2": str(pypdfium2.PYPDFIUM_INFO), "pdfium": str(pypdfium2.PDFIUM_INFO)}
    if actual != PINS:
        raise RuntimeError("Use the recorded development pins; no dependency is installed by this tool.")
    return actual


def safe_relative(value: str) -> Path:
    if not isinstance(value, str) or not value or "\\" in value:
        raise ValueError("Expected a nonempty forward-slash relative corpus path.")
    path = Path(value)
    if path.is_absolute() or any(part in ("", ".", "..") or
        any(character in part for character in ':*?"<>|\0') or part.endswith((".", " ")) or
        re.fullmatch(r"(?i)(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\..*)?", part) for part in value.split("/")):
        raise ValueError("Corpus path must stay within its owned directory.")
    return path


def encrypted(raw: bytes, restricted: bool = False) -> bytes:
    from pypdf import PdfReader, PdfWriter
    from pypdf.constants import UserAccessPermissions
    from pypdf.generic import ArrayObject, ByteStringObject
    writer = PdfWriter()
    writer.clone_document_from_reader(PdfReader(io.BytesIO(raw), strict=True))
    identity = hashlib.sha256(b"T21 original restricted" if restricted else b"T21 original encrypted").digest()[:16]
    writer._ID = ArrayObject([ByteStringObject(identity), ByteStringObject(identity)])
    writer.encrypt("" if restricted else "T21-user", "T21-owner", algorithm="RC4-128",
                   permissions_flag=UserAccessPermissions.PRINT)
    stream = io.BytesIO()
    writer.write(stream)
    return stream.getvalue()


def safety_recipes() -> tuple[dict, dict[str, bytes]]:
    """Explicit expected order, independent of the application's comparator."""
    numbered = load_recipe("numbered")
    files = {}
    scenarios = {}

    def scenario(key, names, bad=None):
        directory = "safety/" + ("invalid/" + key[8:] if key.startswith("invalid-") else key)
        rows = []
        identifiers = []
        for index, name in enumerate(names, 1):
            ids = [f"T03-21-P{index:02d}"]
            if key == "matrix" and name == "10.pdf":
                ids.append("T03-21-P90")
            raw = numbered.make_pdf(name, tuple(ids))
            row = {"path": name, "sha256": sha256(raw), "bytes": len(raw), "hidden": False,
                   "page_count": len(ids), "page_identifiers": ids, "disposition": "included"}
            rows.append(row)
            identifiers.extend(ids)
            files[directory + "/" + name] = raw
        if bad is not None:
            name, raw = bad
            rows.append({"path": name, "sha256": sha256(raw), "bytes": len(raw), "hidden": False,
                         "page_count": None, "page_identifiers": [], "disposition": "rejected"})
            files[directory + "/" + name] = raw
        scenarios[key] = {"source_directory": directory, "source_files": rows,
            "ordered_names": names if bad is None else [names[0], bad[0], names[1]],
            "expected_page_count": len(identifiers) if bad is None else None,
            "expected_page_identifiers": identifiers if bad is None else [],
            "expected_exit_code": 0 if bad is None else 1}

    names = ["0.pdf", "00.pdf", "1.PDF", "01.pdf", "2 [x] ! & (y's).pdf", "10.pdf",
             "12-café.pdf", "13-Å.pdf", "14-ä.pdf", "20-part2.pdf", "20-part02.pdf", "20-part10.pdf",
             "2147483648.pdf", "99999999999999999999999999999999999999.pdf",
             "100000000000000000000000000000000000000.pdf", "item2-part2.pdf", "item2-part10.pdf",
             "WinPDFMerge_legitimate.pdf"]
    scenario("matrix", names)
    matrix = scenarios["matrix"]
    for name, hidden, disposition in (("hidden.PDF", True, "hidden"), ("nested/3.PDF", False, "nested")):
        raw = numbered.make_pdf(name, ("T03-21-P91",))
        files[matrix["source_directory"] + "/" + name] = raw
        matrix["source_files"].append({"path": name, "sha256": sha256(raw), "bytes": len(raw),
            "hidden": hidden, "page_count": 1, "page_identifiers": ["T03-21-P91"], "disposition": disposition})
    raw = b"T21 original inert non-PDF sentinel.\n"
    files[matrix["source_directory"] + "/notes.txt"] = raw
    matrix["source_files"].append({"path": "notes.txt", "sha256": sha256(raw), "bytes": len(raw),
        "hidden": False, "page_count": None, "page_identifiers": [], "disposition": "non_pdf"})
    matrix["source_files"].append({"path": "directory.pdf", "hidden": False, "disposition": "directory"})
    scenario("single-uppercase", ["single.PDF"])
    scenario("names-unicode", ["1.pdf", "2-漢字.pdf"])
    scenarios["names-unicode"]["source_files"][1]["disposition"] = "rejected"
    scenarios["names-unicode"]["expected_exit_code"] = 1
    scenarios["names-unicode"]["expected_page_count"] = None
    scenarios["names-unicode"]["expected_page_identifiers"] = []
    base = numbered.make_pdf("original.pdf", ("T03-21-P92",))
    for kind, raw in (("empty", b""), ("truncated", base[:len(base) // 2]),
                      ("malformed", b"%PDF-1.4\nOriginal corrupt synthetic content\nstartxref\n0\n%%EOF\n"),
                      ("encrypted", encrypted(base)), ("owner-restricted", encrypted(base, True))):
        scenario("invalid-" + kind, ["1.pdf", "10.pdf"], ("5_" + kind + ".pdf", raw))
    return scenarios, files


def catalog_recipe() -> dict:
    scenarios, _ = safety_recipes()
    groups = {}
    for name, (recipe, manifest) in GROUPS.items():
        row = {"generator": recipe, "generator_sha256": sha256((REPO / recipe).read_bytes())}
        if manifest:
            data = json.loads((REPO / manifest).read_text(encoding="utf-8"))
            row.update(manifest=manifest, manifest_sha256=sha256((REPO / manifest).read_bytes()),
                provenance=data["provenance"], license=data["license"], fixtures=data["fixtures"])
            row["merge_order"] = data.get("merge_order", ["mixed.pdf", "scan.pdf", "small-print.pdf"])
            row["expected_page_count"] = sum(item["page_count"] for item in data["fixtures"])
            by_name = {item["file"]: item for item in data["fixtures"]}
            row["expected_page_identifiers"] = [value for name in row["merge_order"] for value in by_name[name]["page_identifiers"]]
        else:
            envelope = load_recipe(name)
            raw = {"conventional-lf.pdf": envelope.conventional()[0], "conventional-cr.pdf": envelope.conventional(b"\r")[0],
                   "incremental.pdf": envelope.incremental(), "xref-stream.pdf": envelope.cross_reference_stream()}
            row.update(provenance="Original manual PDF objects; approved Ghostscript creates only the linearized derivative.",
                license="Repository LICENSE", fixtures=[{"file": key, "bytes": len(value), "sha256": sha256(value),
                "page_count": 1, "page_identifiers": [envelope.IDENTIFIER]} for key, value in raw.items()],
                variable_fixture={"file": "linearized.pdf", "page_count": 1, "page_identifiers": [envelope.IDENTIFIER],
                    "expected_linearized": True, "byte_reproducibility": "Receipt binds actual bytes; native dates/IDs may vary."},
                merge_order=["conventional-cr.pdf", "conventional-lf.pdf", "incremental.pdf", "linearized.pdf", "xref-stream.pdf"], expected_page_count=5,
                expected_page_identifiers=[envelope.IDENTIFIER] * 5)
        groups[name] = row
        fixture_rows = list(row["fixtures"])
        if "variable_fixture" in row:
            fixture_rows.append(row["variable_fixture"])
        scenarios[name] = {"source_directory": name, "source_files": [
            {"path": fixture["file"], "hidden": False, "disposition": "included",
             "page_count": fixture["page_count"], "page_identifiers": fixture["page_identifiers"],
             **({"sha256": fixture["sha256"], "bytes": fixture["bytes"]} if "sha256" in fixture else {})}
            for fixture in fixture_rows], "ordered_names": row["merge_order"],
            "expected_page_count": row["expected_page_count"], "expected_page_identifiers": row["expected_page_identifiers"],
            "expected_exit_code": 0}
    return {"schema_version": 1, "development_only": True, "versions": PINS,
        "provenance": "Original synthetic corpus only; no private PDFs, external images or embedded font files.",
        "groups": groups, "safety": {"generator": "tools/test/corpus.py", "license": "Repository LICENSE",
            "encryption": {"algorithm": "RC4-128", "user_password": "T21-user", "owner_password": "T21-owner",
                "restricted_user_password": "", "purpose": "Synthetic unsupported-input rejection; no production encryption advice."},
            "scenarios": scenarios},
        "preservation_expectations": {"numbered_and_envelopes": "Page IDs/count/order; source hashes unchanged. No general preservation claim.",
            "presets": "Small print/vector/seeded scan/mixed pages. Screen and ebook visual fidelity assessed separately by T17 native evidence.",
            "features_master": "T19: canonical values/widget relations retained; second shared_text name renamed; named-destination/document attachment indexes and tag/ParentTree lost; page attachments retained.",
            "features_email": "T19: forms/widgets lost; navigation and page attachments retained in measured corpus; screen orientation changed; ebook retains sideways appearance. Potentially lossy rewrite.",
            "scope": "Historical measured observations, not universal guarantees. Signature/XFA/PDF-A/accessibility/malware certification unvalidated; retain originals."}}


def load_catalog() -> dict:
    expected = catalog_recipe()
    actual = json.loads(CATALOG.read_text(encoding="utf-8"))
    if not same_json(actual, expected):
        raise ValueError("Tracked corpus catalog differs from its pinned recipes/expectations.")
    return actual


def hidden(path: Path) -> bool:
    return bool(getattr(path.stat(), "st_file_attributes", 0) & 2)


def inventory(root: Path) -> list[dict]:
    rows = []
    for path in sorted(root.rglob("*"), key=lambda value: value.relative_to(root).as_posix()):
        if path.is_symlink() or (hasattr(path, "is_junction") and path.is_junction()):
            raise ValueError("Corpus tree cannot contain links or junctions.")
        relative = path.relative_to(root).as_posix()
        safe_relative(relative)
        if relative == "corpus.json":
            continue
        row = {"path": relative, "kind": "directory" if path.is_dir() else "file", "hidden": hidden(path)}
        if path.is_file():
            raw = path.read_bytes()
            row.update(bytes=len(raw), sha256=sha256(raw))
        rows.append(row)
    return rows


def intended_paths(catalog: dict) -> set[str]:
    """Fixed artifact set; editing a generated receipt cannot admit extra files."""
    paths = set()
    for group, metadata in catalog["groups"].items():
        paths.add(group + "/manifest.json")
        if group == "features":
            paths.add(group + "/generation.json")
        for fixture in metadata["fixtures"]:
            paths.add(group + "/" + fixture["file"])
        if "variable_fixture" in metadata:
            paths.add(group + "/" + metadata["variable_fixture"]["file"])
    for scenario in catalog["safety"]["scenarios"].values():
        paths.update(scenario["source_directory"] + "/" + row["path"] for row in scenario["source_files"])
    for value in list(paths):
        parent = safe_relative(value).parent
        while parent != Path("."):
            paths.add(parent.as_posix())
            parent = parent.parent
    return paths


def inspect_pdf(path: Path, expected: list[str]) -> dict:
    if not isinstance(expected, list) or not expected or any(not isinstance(value, str) or not IDENTIFIER.fullmatch(value) for value in expected):
        raise ValueError("Expected a nonempty identifier list; repeated original page IDs are allowed.")
    import pypdfium2 as pdfium
    if str(pdfium.PYPDFIUM_INFO) != PINS["pypdfium2"] or str(pdfium.PDFIUM_INFO) != PINS["pdfium"]:
        raise RuntimeError("Independent inspection requires the recorded native PDFium pins.")
    before = path.read_bytes()
    pages = []
    with pdfium.PdfDocument(io.BytesIO(before)) as document:
        if len(document) != len(expected):
            raise ValueError("Page count differs from corpus expectation.")
        for index in range(len(document)):
            with closing(document[index]) as page, closing(page.get_textpage()) as text:
                found = IDENTIFIER.findall(text.get_text_range())
                if found != [expected[index]]:
                    raise ValueError(f"Page {index + 1} visible identifier/order differs.")
                pages.append({"identifier": found[0], "size_points": list(page.get_size()), "rotation": page.get_rotation()})
    if path.read_bytes() != before:
        raise ValueError("Source changed during read-only inspection.")
    return {"sha256": sha256(before), "page_count": len(pages), "page_identifiers": expected, "pages": pages}


def materialize(output: Path, ghostscript: Path) -> dict:
    versions()
    catalog = load_catalog()
    output = output.resolve()
    if output.exists() or not output.is_relative_to(REPO / "tests/.work"):
        raise ValueError("Use a new explicitly owned directory under tests/.work.")
    if os.name != "nt":
        raise RuntimeError("Corpus includes real Windows hidden attributes and approved Windows Ghostscript; no native pass on another OS.")
    load_recipe("envelopes").verify_ghostscript(ghostscript, REPO)
    output.mkdir(parents=True, exist_ok=False)
    numbered = output / "numbered"
    numbered.mkdir()
    for name, raw in load_recipe("numbered").expected_files().items():
        (numbered / name).write_bytes(raw)
    for group in ("presets", "features"):
        result = load_recipe(group).generate(output / group)
        target = output / group / "manifest.json"
        with target.open("x", encoding="utf-8") as stream:
            stream.write(json.dumps(result, indent=2, ensure_ascii=False) + "\n")
    command = [sys.executable, "-B", str(REPO / GROUPS["envelopes"][0]), "--output", str(output / "envelopes"), "--ghostscript", str(ghostscript)]
    result = subprocess.run(command, stdin=subprocess.DEVNULL, capture_output=True, timeout=30,
                            creationflags=subprocess.CREATE_NO_WINDOW, shell=False)
    if result.returncode:
        raise RuntimeError("Envelope reconstruction failed: " + result.stderr.decode("utf-8", errors="replace"))
    scenarios = catalog["safety"]["scenarios"]
    _, files = safety_recipes()
    for relative, raw in files.items():
        path = output / safe_relative(relative)
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("xb") as stream:
            stream.write(raw)
    for scenario in scenarios.values():
        for row in scenario["source_files"]:
            path = output / safe_relative(scenario["source_directory"]) / safe_relative(row["path"])
            if row["disposition"] == "directory":
                path.mkdir(exist_ok=False)
            if row["hidden"] and not ctypes.windll.kernel32.SetFileAttributesW(str(path), 2):
                raise OSError("Could not set the real Windows hidden attribute.")
    receipt = {"schema_version": 1, "catalog_sha256": sha256(CATALOG.read_bytes()), "versions": PINS,
               "groups": catalog["groups"], "safety": catalog["safety"], "inventory": inventory(output)}
    with (output / "corpus.json").open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(receipt, indent=2, ensure_ascii=False) + "\n")
    return verify(output)


def verify(root: Path) -> dict:
    versions()
    catalog = load_catalog()
    root = root.resolve()
    receipt = json.loads((root / "corpus.json").read_text(encoding="utf-8"))
    if type(receipt.get("schema_version")) is not int or receipt["schema_version"] != 1 or receipt.get("catalog_sha256") != sha256(CATALOG.read_bytes()) or \
        not same_json(receipt.get("versions"), PINS) or not same_json(receipt.get("groups"), catalog["groups"]) or not same_json(receipt.get("safety"), catalog["safety"]):
        raise ValueError("Generated receipt differs from tracked corpus expectations.")
    actual_inventory = inventory(root)
    if not same_json(actual_inventory, receipt.get("inventory")) or {row["path"] for row in actual_inventory} != intended_paths(catalog):
        raise ValueError("Corpus inventory/hash/attribute differs; verification never repairs or rewrites files.")
    observed = []
    for group, metadata in catalog["groups"].items():
        fixtures = list(metadata["fixtures"])
        if "variable_fixture" in metadata:
            fixtures.append(metadata["variable_fixture"])
        for row in fixtures:
            path = root / group / safe_relative(row["file"])
            raw = path.read_bytes()
            if "sha256" in row and (sha256(raw) != row["sha256"] or len(raw) != row["bytes"]):
                raise ValueError("Fixture differs from tracked original bytes: " + group + "/" + row["file"])
            if row.get("expected_linearized") and not load_recipe("envelopes").envelope(raw)["linearized_dictionary_in_first1024"]:
                raise ValueError("Native envelope derivative is not linearized.")
            observation = inspect_pdf(path, row["page_identifiers"])
            if "page_size_points" in row and any(page["size_points"] != row["page_size_points"] for page in observation["pages"]):
                raise ValueError("Page dimensions differ from tracked expectations.")
            if "structural_rotations" in row and [page["rotation"] for page in observation["pages"]] != row["structural_rotations"]:
                raise ValueError("Page rotations differ from tracked expectations.")
            observed.append({"path": group + "/" + row["file"], **observation})
    for scenario in catalog["safety"]["scenarios"].values():
        for row in scenario["source_files"]:
            path = root / safe_relative(scenario["source_directory"]) / safe_relative(row["path"])
            if row["disposition"] == "directory":
                if not path.is_dir():
                    raise ValueError("Missing PDF-named directory sentinel.")
                continue
            raw = path.read_bytes()
            if ("sha256" in row and (sha256(raw) != row["sha256"] or len(raw) != row["bytes"])) or hidden(path) != row["hidden"]:
                raise ValueError("Safety source bytes/attribute differs: " + row["path"])
            if row["page_identifiers"]:
                relative = path.relative_to(root).as_posix()
                if not any(item["path"] == relative for item in observed):
                    observed.append({"path": relative, **inspect_pdf(path, row["page_identifiers"])})
    return {"result": "pass", "catalog_sha256": sha256(CATALOG.read_bytes()), "versions": PINS,
            "inventory_entries": len(actual_inventory), "inspected_files": len(observed), "observations": observed,
            "product_merge_executed": False, "ghostscript_executed_by_verification": False,
            "scope": "Fixture reconstruction and independent PDFium inspection; no application/native/manual acceptance claim."}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    build = commands.add_parser("materialize")
    build.add_argument("--output", type=Path, required=True)
    build.add_argument("--ghostscript", type=Path, required=True)
    check = commands.add_parser("verify")
    check.add_argument("--root", type=Path, required=True)
    inspect = commands.add_parser("inspect")
    inspect.add_argument("--pdf", type=Path, required=True)
    inspect.add_argument("--expected-identifiers", type=Path, required=True)
    arguments = parser.parse_args()
    if arguments.command == "materialize":
        result = materialize(arguments.output, arguments.ghostscript)
    elif arguments.command == "verify":
        result = verify(arguments.root)
    else:
        result = inspect_pdf(arguments.pdf, json.loads(arguments.expected_identifiers.read_text(encoding="utf-8-sig")))
    # ASCII JSON also survives legacy Windows console code pages for CJK paths.
    print(json.dumps(result))


if __name__ == "__main__":
    main()
