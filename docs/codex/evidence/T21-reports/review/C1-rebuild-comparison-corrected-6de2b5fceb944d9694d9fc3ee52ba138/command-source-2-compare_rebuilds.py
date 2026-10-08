"""Independent comparison of separately reconstructed T21 corpus inventories."""
from __future__ import annotations
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--first", type=Path, required=True)
    parser.add_argument("--second", type=Path, required=True)
    args = parser.parse_args()
    roots = [args.first.resolve(), args.second.resolve()]
    receipts = [json.loads((root / "corpus.json").read_text(encoding="utf-8")) for root in roots]
    assert roots[0] != roots[1]
    assert receipts[0]["catalog_sha256"] == receipts[1]["catalog_sha256"]
    assert receipts[0]["versions"] == receipts[1]["versions"]
    before, after = [{row["path"]: row for row in receipt["inventory"]} for receipt in receipts]
    assert set(before) == set(after)
    variable = {"envelopes/linearized.pdf", "envelopes/manifest.json"}
    fixed = []
    for relative in before:
        first, second = before[relative], after[relative]
        assert (first["kind"], first["hidden"]) == (second["kind"], second["hidden"]), relative
        if first["kind"] == "file":
            for root, row in zip(roots, (first, second)):
                path = root / relative
                assert digest(path) == row["sha256"] and path.stat().st_size == row["bytes"], relative
        if relative in variable:
            continue
        assert first == second, relative
        fixed.append(relative)
    manifests = [json.loads((root / "envelopes/manifest.json").read_text(encoding="utf-8")) for root in roots]
    fixed_envelope_rows = [{row["file"]: row for row in manifest["fixtures"] if row["file"] != "linearized.pdf"} for manifest in manifests]
    assert fixed_envelope_rows[0] == fixed_envelope_rows[1]
    commands = [manifest["ghostscript"]["command"] for manifest in manifests]
    normalized = [[operand.replace(str(root), "<OWNED_CORPUS>") for operand in command] for command, root in zip(commands, roots)]
    assert normalized[0] == normalized[1]
    assert all(manifest["ghostscript"]["exit_code"] == 0 and manifest["ghostscript"]["version"] == "10.08.0" for manifest in manifests)
    print(json.dumps({"task": "T21", "reviewer": "/root/review", "authorship": "Independent reviewer, no tracked implementation authorship.", "observed_at_utc": datetime.now(timezone.utc).isoformat(), "producer_sha256": digest(Path(__file__)), "result": "pass", "roots": [str(root) for root in roots], "inventory_entries": len(before), "fixed_entries_compared": len(fixed), "fixed_entries": fixed, "variable_entries": [{"path": relative, "first": before[relative], "second": after[relative]} for relative in sorted(variable)], "variable_scope": "Approved native Ghostscript linearized PDF may vary in native dates/IDs. Its generation manifest binds actual hash, unique output argv and measured elapsed time; fixed envelope originals and normalized native argv remain identical. Both roots were separately fully verified with PDFium.", "catalog_sha256": receipts[0]["catalog_sha256"]}, indent=2))

if __name__ == "__main__":
    main()
