"""Prepare separate C1 reader from immutable retained dirty reader; no runtime execution."""
from pathlib import Path
import hashlib
import sys

repo = Path(__file__).resolve().parents[2]
work = repo / "tests/.work"
old = work / "Review-T19DirtyFeatures.py"
new = work / "Review-T19C1Features.py"
assert not new.exists()
source = old.read_text(encoding="utf-8")
replacements = {
    'BASE = "7e0fe6fde769911f323bd87e4c4f9382d26e331b"': 'BASE = "50220eccd1917e44a52d94cbfe3bf35b1940f3d8"\nSTART = "7e0fe6fde769911f323bd87e4c4f9382d26e331b"',
    'T19-dirty-feature-review-sources-': 'T19-C1-feature-review-sources-',
    'runtime_diff = git("diff", BASE,': 'runtime_diff = git("diff", START,',
    'meta["dirty_worktree"] is True, "Dirty native baseline accurately scoped"': 'meta["dirty_worktree"] is False, "Clean C1 native baseline accurately scoped"',
    '"Actual dirty six-case Pester summary"': '"Actual clean C1 six-case Pester summary"',
    'row = runs[0]; review.check(len(runs) == 1 and row["exit_code"] == 0 and row["summary"] == summary, "Actual driver/summary binding")': 'row = next(r for r in runs if r["tier"] == "PreservationNative"); review.check(row["exit_code"] == 0 and row["summary"] == summary, "Actual driver/summary binding")',
    'CommitUnderReviewedNative=BASE, NativeWorktree="dirty"': 'CommitUnderTest=BASE, CommitUnderReviewedNative=BASE, NativeWorktree="clean"',
    'Readiness="No remaining blocking findings in current scoped wording; clean C1 review/acceptance remains pending"': 'Readiness="No blocking findings in C1 actual structural/docs/native receipts; public privacy/closure audit remains pending"',
    '"Dirty six-case native runs only; no clean acceptance, full regression, interactive GUI, physical Explorer, signature, XFA, accessibility, malware or archival certification claim."': '"Only actual T19 six-tier C1 reports, fourteen documentation and six feature-native cases per shell; no interactive GUI, physical Explorer, signature, XFA, accessibility, malware or archival certification claim."',
}
for before, after in replacements.items():
    assert source.count(before) == 1, before
    source = source.replace(before, after)
source = source.replace('assert sha(Path(sys.executable).read_bytes()) == PYTHON',
                        'assert sha(Path(sys.executable).read_bytes()) == PYTHON\n    assert git("rev-parse", "HEAD") == BASE and not git("status", "--porcelain=v1"), "Exact clean C1 required"')
extra = '''
    total_cases = 0; total_reports = 0
    for root_label in args.root:
        root = (REPO/root_label).resolve(); runs = review.obj(root/"runs.json")
        review.check(len(runs) == 6, "Six completed scoped C1 tiers per actual host")
        subtotal = 0
        for row in runs:
            tier = row["tier"]; summary = review.obj(root/(tier+".summary.json"))
            review.check(row["exit_code"] == 0 and row["summary"] == summary, "C1 driver actual exit/summary: " + tier)
            review.check(summary["commit_under_test"] == BASE and summary["dirty_worktree"] is False, "C1 exact clean summary: " + tier)
            review.check(summary["passed"] == summary["total"] and all(summary[k] == 0 for k in ["failed", "failed_blocks", "failed_containers", "skipped", "not_run"]), "C1 all actual cases pass: " + tier)
            xml = ET.fromstring(review.read(root/(tier+".results.xml"))); cases = list(xml.iter("test-case"))
            review.check(len(cases) == summary["total"] and all(c.attrib.get("success") == "True" and c.attrib.get("executed") == "True" for c in cases), "C1 actual NUnit count/outcome: " + tier)
            review.bound(row["stdout"], row["stdout_sha256"]); review.bound(row["stderr"], row["stderr_sha256"])
            subtotal += summary["total"]; total_reports += 1
        aggregate = review.obj(root/"aggregate.json")
        review.check(aggregate["passed"] == subtotal == 414 and aggregate["tiers"] == 6 and aggregate["bad_counts"] == 0 and aggregate["dirty_worktree"] is False and aggregate["commit_under_test"] == BASE, "C1 aggregate exactly matches actual reports")
        total_cases += subtotal
        docs_summary = review.obj(root/"PreservationDocs.summary.json")
        review.check(docs_summary["total"] == 14, "Actual C1 documentation/help cases per host")
        docs_stdout = review.text(root/"PreservationDocs.stdout.txt")
        markers = re.findall(r"^Preservation documentation receipts: (.+)$", docs_stdout, re.M)
        review.check(len(markers) == 1, "Unique C1 documentation receipt marker")
        docs = review.obj(markers[0].strip())
        review.check(docs["ShellVersion"] == docs_summary["shell_version"] and len(docs["Observations"]) == 14, "Actual fourteen help/documentation observations")
        for binding in docs["SourceBindings"]:
            raw = review.bound(binding["RetainedSourcePath"], binding["SHA256"])
            review.check(raw == (REPO/binding["Path"]).read_bytes(), "Exact C1 documentation source: " + binding["Path"])
        help_binding = docs["HelpBinding"]
        actual_help = review.bound(help_binding["RetainedTextPath"], help_binding["SHA256"]).decode("utf-8-sig")
        review.check(help_binding["ApplicationInvoked"] is False and help_binding["Command"][0] == "Get-Help", "Real comment help extraction without application invocation")
        review.check("Neither output guarantees PDF/A" in actual_help and "Keep original documents" in actual_help, "Actual rendered help contains conservative limits")
    review.check(total_cases == 828 and total_reports == 12, "Actual C1 twelve scoped reports total828")
    review.check(git("rev-parse", "HEAD") == BASE and not git("status", "--porcelain=v1"), "Exact C1 still clean after read-only review")
'''
needle = '    limits = review.text(REPO/"docs/PDF_LIMITATIONS.md")'
assert source.count(needle) == 1
source = source.replace(needle, extra + needle)
source = source.replace('SourceBindings=source_rows, ProducerSHA256=',
                        'ExecutedCaseCount=total_cases, ExecutedReportCount=total_reports,\n        SourceBindings=source_rows, ProducerSHA256=')
source = source.replace('"Root and fixture/test authors own implementations"', '"Root and fixture/test authors own implementations"')
compile(source, str(new), "exec")
new.write_text(source, encoding="utf-8")
print("Prepared separate C1 source", hashlib.sha256(new.read_bytes()).hexdigest())
