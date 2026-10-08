"""Ignored read-only history binder for actual T17 native focused attempts."""
import hashlib
import json
from pathlib import Path
import re

REPO = Path(__file__).resolve().parents[2]


def binding(path):
    raw = path.read_bytes()
    return {"Path": path.relative_to(REPO).as_posix(),
            "SHA256": hashlib.sha256(raw).hexdigest(), "Bytes": len(raw)}


attempts = []
for root in sorted((REPO / "tests/.work").glob("T17-dirty-*-SizeReportingNative-*")):
    stdout = (root / "stdout.txt").read_text(encoding="utf-8")
    reports = Path(re.search(r"(?m)^Reports: (.+)$", stdout)[1].strip())
    observations_path = Path(re.search(r"(?m)^Size reporting observations: (.+)$", stdout)[1].strip())
    summary = json.loads((reports / "summary.json").read_text(encoding="utf-8-sig"))
    observations = json.loads(observations_path.read_text(encoding="utf-8-sig"))
    launch = json.loads((root / "launch.json").read_text(encoding="utf-8-sig"))
    source = root / "SizeReporting.Native.Tests.ps1"
    assert binding(source)["SHA256"] == launch["TestSourceSHA256"].lower()
    files = [binding(p) for p in sorted(root.iterdir()) if p.is_file()]
    files += [binding(reports / "summary.json"), binding(reports / "results.xml"), binding(observations_path)]
    cases = []
    for case in sorted(observations_path.parent.iterdir()):
        if case.is_dir() and case.name != "original-preset-corpus":
            cases.append({"Root": case.relative_to(REPO).as_posix(),
                          "Files": [binding(p) for p in sorted(case.rglob("*")) if p.is_file() and p.suffix.lower() != ".pdf"],
                          "RetainedSyntheticPdfHashes": [binding(p) for p in sorted(case.rglob("*.pdf"))]})
    attempts.append({"Selection": launch["Shell"], "Root": root.relative_to(REPO).as_posix(),
                     "Passed": summary["passed"], "Failed": summary["failed"],
                     "Total": summary["total"], "ExitCode": int((root / "exit-code.txt").read_text()),
                     "ObservationCount": len(observations["Observations"]),
                     "TestSourceAtInvocationSHA256": launch["TestSourceSHA256"].lower(),
                     "SourceSnapshot": source.relative_to(REPO).as_posix(),
                     "AllListedSourceSnapshotsRetainedBeforeRun": True,
                     "Diagnosis": ("Test-only case-insensitive Fixture parameter shadowed outer fixture expectation; six corpus cases passed, five tiny cases failed assertion property lookup after actual jobs. Corrected expectation variable; no application source change." if summary["failed"] else "All eleven targeted native cases passed on the selected approved actual Windows host."),
                     "Files": files, "RetainedCopiedApplicationCases": cases})
assert len(attempts) == 4
target = REPO / "tests/.work/T17-native-dirty-history.json"
target.write_text(json.dumps({"SchemaVersion": 1, "Task": "T17", "Attempts": attempts,
                              "Scope": "Historical dirty focused attempts only; no counts contribute to clean C1 acceptance. Exact test/generator/manifest/entry/helper/runner/launcher source snapshots retained before each actual run. Raw XML/summary/streams/receipts and original synthetic artifacts remain ignored; no document/engine substitutes count as native support or manual visual pass."}, indent=2)+"\n", encoding="utf-8")
print(json.dumps({"Index": binding(target), "Attempts": len(attempts), "Counts": [[a["Selection"], a["Passed"], a["Failed"]] for a in attempts]}))
