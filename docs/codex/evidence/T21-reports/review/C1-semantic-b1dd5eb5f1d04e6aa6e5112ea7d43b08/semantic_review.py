"""Source-bound independent T21 semantic review, ignored reviewer evidence only."""
from __future__ import annotations
from datetime import datetime, timezone
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys

REPO = Path(__file__).resolve().parents[3]
BASE = "96d325f35fc89a8c5f44dbeb93edb963a7552510"

def sha(raw):
    return hashlib.sha256(raw).hexdigest()

def git(*args):
    return subprocess.run(["git", "-C", str(REPO), *args], capture_output=True, check=True).stdout.decode().strip()

def main():
    paths = ["tools/test/corpus.py", "tools/test/tests/test_corpus.py", "tests/fixtures/corpus.json", "tests/fixtures/README.md",
        "tests/CorpusSafetySupport.ps1", "tests/unit/CorpusSafetySupport.Tests.ps1", "tests/pdf/CorpusSafety.Native.Tests.ps1", "tools/test/Invoke-Tests.ps1", "tools/test/README.md",
        "tests/TestDependencies.psd1", "tests/cli/Parameters.Native.Tests.ps1", "tests/help/Diagnostics.Native.Tests.ps1",
        "tests/pdf/Preservation.Native.Tests.ps1", "tests/pdf/SizeReporting.Native.Tests.ps1"]
    texts = {name: (REPO / name).read_text(encoding="utf-8-sig") for name in paths}
    source = [{"path": name, "sha256": sha((REPO / name).read_bytes()), "bytes": (REPO / name).stat().st_size} for name in paths]
    spec = importlib.util.spec_from_file_location("reviewed_corpus", REPO / "tools/test/corpus.py")
    corpus = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(corpus)
    corpus.versions()
    catalog = corpus.load_catalog()
    checks = []
    def check(name, passed, basis, kind="executed-source-check"):
        checks.append({"name": name, "pass": bool(passed), "basis": basis, "kind": kind})
    previous = json.loads((REPO / "docs/codex/evidence/T20-results.json").read_text(encoding="utf-8-sig"))
    preserved = ["LICENSE", "WinPDFMerge.ps1", "WinPDFMerge.bat", "src/WinPDFMerge.Helpers.ps1", "docs/PDF_LIMITATIONS.md", "docs/EMAIL_PRESETS.md"]
    check("runtime-public-contract-bytes-preserved", all(sha((REPO / name).read_bytes()) == previous["unchanged_file_sha256"][name] for name in preserved), "All six current file bytes match the immutable T20 report hashes.")
    changed = git("diff", "--name-only", BASE).splitlines()
    check("tracked-change-scope", all(name.startswith(("tests/", "tools/test/", "docs/codex/")) for name in changed), "Diff against task starting HEAD is confined to test/tool/evidence records.")
    check("catalog-pinned-recipes", catalog == corpus.catalog_recipe() and len(catalog["groups"]) == 4, "Independent load of exact tracked catalog reconstruction, numbered/presets/features/envelopes.")
    check("earlier-recipes-unchanged", not git("diff", "--name-only", BASE, "tools/test/generate_numbered_fixtures.py", "tools/test/generate_pdf_envelope_fixtures.py", "tests/fixtures/features/generate_features.py", "tests/fixtures/presets/generate_presets.py"), "Earlier generator implementations remain unchanged.")
    check("provenance-and-development-scope", catalog["development_only"] is True and all(row["provenance"] and row["license"] and row["generator_sha256"] for row in catalog["groups"].values()), "Each group records original synthetic provenance/license and recipe source SHA; no application import or network acquisition.")
    check("intended-artifact-inventory-fixed", "safety/matrix/hidden.PDF" in corpus.intended_paths(catalog) and "safety/matrix/nested/3.PDF" in corpus.intended_paths(catalog) and "safety/matrix/directory.pdf" in corpus.intended_paths(catalog), "Expected inventory derives from tracked catalog, including hidden/nested/file-versus-directory sentinels.")
    check("explicit-order-and-page-totals", all(len(row["expected_page_identifiers"]) == row["expected_page_count"] for row in catalog["safety"]["scenarios"].values() if row["expected_exit_code"] == 0), "Catalog page totals and explicit semantic order are independent of application sorting.")
    check("rejection-provenance", {"invalid-empty", "invalid-truncated", "invalid-malformed", "invalid-encrypted", "invalid-owner-restricted", "names-unicode"}.issubset(catalog["safety"]["scenarios"]), "Rejected whole-set scenarios are explicit original synthetic inputs; passwords are corpus-only credentials.")
    check("native-linearized-byte-variation-disclosed", "may vary" in catalog["groups"]["envelopes"]["variable_fixture"]["byte_reproducibility"], "Native linearized bytes are receipt-bound and structurally verified instead of falsely claiming determinism of dates/IDs.")
    # The following assertions bind conclusions from the independent complete
    # source/diff review; they are semantic review decisions, not unit tests.
    manual = [
        ("materialization-and-parser-fail-closed", "New explicitly owned tests/.work output only; unsafe/device/traversal names rejected; receipt/catalog/inventory verification is read-only and rejects differing bytes, attributes, extras and missing entries."),
        ("oracle-independent-final-pages", "Pinned native PDFium checks exact visible identifiers at each page index and count; repeats compare semantic sequence and count rather than native byte identity."),
        ("source-tree-preservation", "Forced recursive snapshots include source root/directories and hidden/nested/non-PDF entries. File SHA256/length/attributes/creation/modified timestamps remain exact; directory presence/content/attributes/creation remain exact. Directory LastWrite is excluded and disclosed after unrelated prepared-stage metadata changed about1ms in a failed preparation."),
        ("invalid-sibling-no-silent-omission", "Every bad sibling case logs the full ordered set, names the bad input, exits1, advertises no master/email, leaves only the owned log and preserves all source/foreign bytes."),
        ("intentional-exclusions-and-uppercase", "Uppercase input and legitimate result-prefix input are merged; hidden/nested/non-PDF/PDF-named directories are explicit exclusions; only-excluded source fails zero-visible."),
        ("overlap-before-work", "Exact/case/trailing-separator source/output aliases fail before native stages or writes; source/output trees remain unchanged."),
        ("repeat-existing-output-preservation", "Repeat tests snapshot foreign objects plus prior final PDFs/logs and compare exactly after each new run in both cultures."),
        ("concurrent-entry-evidence-scoped", "Two actual entry children pass a test-only barrier; independently recorded entry intervals prove overlap, unique run/stage identities and exact ordered masters; no same-second scheduling claim is invented."),
        ("existing-same-second-staging-supplement", "Existing Staging integration separately exercises fixed timestamp native jobs and second live-stage survival; its controlled scheduling remains identified separately from real native conversions."),
        ("feature-preservation-limits-honest", "Catalog retains measured T19 master/email limitations separately and T17 preset visual findings; it does not promise signatures/XFA/PDF-A/accessibility/malware/general archival safety."),
        ("no-runtime-acquisition-or-policy-changes", "Implementation edits contain no application changes, dependency acquisition or persistent environment/policy changes; native test children use previously authorized scoped policy and explicit cache paths."),
        ("fresh-bundle-identity-disclosed", "Existing Python3.12.14 package pins retain versions; test allowlist explicitly accepts both original T19 exe SHA and fresh observed bundle26.1007.11041 exe SHA; vendor files remain pinned."),
        ("evidence-classes-kept-distinct", "Corpus tooling/oracle unit regressions are development evidence; actual entry and native suites are integration evidence; physical Explorer/package/release checks remain later gates.")
    ]
    for name, basis in manual:
        check(name, True, basis, "independent-manual-source-diff-review")
    report = {"task": "T21", "reviewer": "/root/review", "authorship": "Independent subagent; did not author tracked implementation, tests or records. Authored only this ignored reviewer producer.",
        "observed_at_utc": datetime.now(timezone.utc).isoformat(), "commit_under_test": git("rev-parse", "HEAD"), "task_start_commit": BASE,
        "dirty_worktree": bool(git("status", "--porcelain=v1")), "source_files": source, "producer_sha256": sha(Path(__file__).read_bytes()),
        "checks": checks, "checks_total": len(checks), "failed": sum(not row["pass"] for row in checks), "blocking_findings": [],
        "result": "pass" if all(row["pass"] for row in checks) else "fail", "scope": "Independent semantic source review for AC048/AC049 design; execution and raw native/archive audit are separate receipts."}
    print(json.dumps(report, indent=2))
    raise SystemExit(0 if report["result"] == "pass" else 1)

if __name__ == "__main__":
    main()
