"""Generate original structural-envelope fixtures for development tests only.

No application dependency or external PDF/font input. Manual PDF syntax follows
ISO 32000-1 sections 7.5.2-7.5.8; linearized output is produced by the explicitly
verified approved Ghostscript cache, using its documented FastWebView option.
All output files are new and failure leaves the owned directory for diagnosis.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import time

IDENTIFIER = "T03-01-P01"
GS_TIMEOUT_SECONDS = 10


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def page_objects(eol: bytes) -> list[bytes]:
    content = b"BT /F1 24 Tf 40 100 Td (" + IDENTIFIER.encode("ascii") + b") Tj ET" + eol
    return [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
        # ISO 32000-1 7.3.8.1 requires LF or CRLF after stream, including
        # when the surrounding structural and footer lines use CR.
        b"<< /Length " + str(len(content)).encode() + b" >>" + eol + b"stream\n" + content + b"endstream",
    ]


def body(version: str, eol: bytes) -> tuple[bytearray, dict[int, int]]:
    result = bytearray(b"%PDF-" + version.encode() + eol + b"%\xe2\xe3\xcf\xd3" + eol)
    offsets = {}
    for number, value in enumerate(page_objects(eol), 1):
        offsets[number] = len(result)
        result.extend(str(number).encode() + b" 0 obj" + eol + value + eol + b"endobj" + eol)
    return result, offsets


def conventional(eol: bytes = b"\n") -> tuple[bytes, int]:
    result, offsets = body("1.4", eol)
    xref = len(result)
    result.extend(b"xref" + eol + b"0 6" + eol + b"0000000000 65535 f " + eol)
    for number in range(1, 6):
        result.extend(f"{offsets[number]:010d} 00000 n ".encode() + eol)
    result.extend(b"trailer" + eol + b"<< /Size 6 /Root 1 0 R >>" + eol + b"startxref" + eol + str(xref).encode() + eol + b"%%EOF" + eol)
    return bytes(result), xref


def incremental() -> bytes:
    original, previous = conventional()
    result = bytearray(original)
    info_offset = len(result)
    result.extend(b"6 0 obj\n<< /Title (Original incremental update fixture) >>\nendobj\n")
    xref = len(result)
    result.extend(b"xref\n6 1\n" + f"{info_offset:010d} 00000 n \n".encode())
    result.extend(b"trailer\n<< /Size 7 /Root 1 0 R /Info 6 0 R /Prev " + str(previous).encode() + b" >>\nstartxref\n" + str(xref).encode() + b"\n%%EOF\n")
    return bytes(result)


def cross_reference_stream() -> bytes:
    result, offsets = body("1.5", b"\n")
    offsets[6] = len(result)
    entries = b"\x00" + (0).to_bytes(4, "big") + (65535).to_bytes(2, "big")
    entries += b"".join(b"\x01" + offsets[number].to_bytes(4, "big") + (0).to_bytes(2, "big") for number in range(1, 7))
    result.extend(b"6 0 obj\n<< /Type /XRef /Size 7 /Root 1 0 R /W [1 4 2] /Index [0 7] /Length " + str(len(entries)).encode() + b" >>\nstream\n")
    result.extend(entries + b"\nendstream\nendobj\nstartxref\n" + str(offsets[6]).encode() + b"\n%%EOF\n")
    return bytes(result)


def envelope(data: bytes) -> dict:
    values = []
    for match in re.finditer(rb"startxref[\x00\x09\x0A\x0C\x0D\x20]+([0-9]+)", data):
        target = int(match.group(1))
        values.append({"keyword_offset": match.start(), "target_offset": target,
                       "target_preview": data[target:target + 100].decode("latin1")})
    return {"startxref_records": values, "final_startxref": values[-1]["target_offset"],
            "eof_offsets": [match.start() for match in re.finditer(rb"%%EOF", data)],
            "linearized_dictionary_in_first1024": bool(re.search(rb"/Linearized\s+1(?:\.0)?\b", data[:1024]))}


def verify_ghostscript(path: Path, repo: Path) -> dict:
    receipt_path = repo / "docs/codex/evidence/T09-gs-acquisition.json"
    receipt = json.loads(receipt_path.read_bytes())
    root = Path(os.path.expandvars(receipt["cache_root"])).resolve()
    expected = (root / receipt["version_probe"]["executable"]).resolve()
    if path.resolve() != expected:
        raise ValueError("Use the previously approved canonical Ghostscript executable.")
    for item in receipt["ghostscript_extraction"]["selected_files"]:
        if digest((root / item["relative_path"]).read_bytes()) != item["sha256"]:
            raise ValueError("Approved Ghostscript console or DLL bytes changed.")
    return {"version": "10.08.0", "acquisition_receipt": "docs/codex/evidence/T09-gs-acquisition.json",
            "acquisition_receipt_sha256": digest(receipt_path.read_bytes())}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--ghostscript", type=Path, required=True)
    args = parser.parse_args()
    output = args.output.resolve()
    if output.exists() and (not output.is_dir() or any(output.iterdir())):
        raise ValueError("Fixture output must be a new or empty explicitly owned directory.")
    repo = Path(__file__).resolve().parents[2]
    gs = args.ghostscript.resolve()
    approval = verify_ghostscript(gs, repo)
    output.mkdir(parents=True, exist_ok=True)
    originals = {
        "conventional-lf.pdf": conventional()[0],
        "conventional-cr.pdf": conventional(b"\r")[0],
        "incremental.pdf": incremental(),
        "xref-stream.pdf": cross_reference_stream(),
    }
    for name, data in originals.items():
        with (output / name).open("xb") as target:
            target.write(data)
    linearized = output / "linearized.pdf"
    command = [str(gs), "-dBATCH", "-dNOPAUSE", "-dSAFER", "-dPDFSTOPONERROR", "-sDEVICE=pdfwrite",
               "-dCompatibilityLevel=1.7", "-dFastWebView=true", "-o", str(linearized), "-f", str(output / "conventional-lf.pdf")]
    environment = {key: value for key, value in os.environ.items() if key.upper() != "GS_OPTIONS"}
    started = time.monotonic()
    # Direct fixed vector, closed stdin, dual captured streams and an owned timeout.
    result = subprocess.run(command, shell=False, stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            timeout=GS_TIMEOUT_SECONDS, env=environment, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    elapsed = round((time.monotonic() - started) * 1000)
    if result.returncode != 0 or not linearized.is_file() or not linearized.stat().st_size:
        raise RuntimeError("Approved Ghostscript failed to create the owned linearized fixture: " + result.stderr.decode("utf-8", errors="replace"))
    originals["linearized.pdf"] = linearized.read_bytes()
    fixtures = [{"file": name, "bytes": len(data), "sha256": digest(data), "expected_page_count": 1,
                 "page_identifiers": [IDENTIFIER], "envelope": envelope(data)} for name, data in originals.items()]
    if not fixtures[-1]["envelope"]["linearized_dictionary_in_first1024"]:
        raise RuntimeError("FastWebView output did not include the expected real linearization dictionary.")
    manifest = {"schema_version": 1, "generator": "tools/test/generate_pdf_envelope_fixtures.py",
                "generator_sha256": digest(Path(__file__).read_bytes()), "provenance": "Original one-page manual text/byte fixtures; no external PDF or font input.",
                "license": "Repository LICENSE", "fixtures": fixtures,
                "cr_fixture_line_endings": "CR structural/header/xref/footer lines; required LF follows the stream keyword (not an all-CR stream marker).",
                "ghostscript": {**approval, "command": command, "timeout_ms": GS_TIMEOUT_SECONDS * 1000,
                                "exit_code": result.returncode, "elapsed_ms": elapsed, "stdout": result.stdout.decode("utf-8", errors="replace"),
                                "stderr": result.stderr.decode("utf-8", errors="replace"), "child_gs_options_removed": True},
                "sources": ["https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/PDF32000_2008.pdf",
                            "https://ghostscript.readthedocs.io/en/gs10.08.0/VectorDevices.html"],
                "scope": "Development envelope regression fixtures only, not PDF validation/fidelity or application FastWebView behavior."}
    with (output / "manifest.json").open("xb") as target:
        target.write((json.dumps(manifest, indent=2) + "\n").encode("utf-8"))
    print(json.dumps({"manifest": str(output / "manifest.json"), "fixtures": fixtures}, indent=2))


if __name__ == "__main__":
    main()
