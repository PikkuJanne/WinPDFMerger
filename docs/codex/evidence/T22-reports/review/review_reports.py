"""Independent machine-readable inspection of source-bound T22 clean reports."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import uuid
import xml.etree.ElementTree as ET

REPO = Path(__file__).resolve().parents[3]
BASELINE = "75df26187202104995e58f020c2442d973c7b075"


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def git(*arguments):
    return subprocess.check_output(["git", "-C", str(REPO), *arguments]).decode("utf-8").strip()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--run", action="append", type=Path, required=True)
    parser.add_argument("--static", action="append", type=Path, default=[])
    args = parser.parse_args()
    output = REPO / "tests/.work/T22-review" / ("reports-" + uuid.uuid4().hex)
    output.mkdir(parents=True)
    shutil.copyfile(__file__, output / "review_reports.py")
    commit = git("rev-parse", "HEAD")
    review = {
        "task": "T22", "observed_at_utc": datetime.now(timezone.utc).isoformat(),
        "commit_under_review": commit, "producer_sha256": sha(__file__), "runs": [], "static": [],
        "command": [sys.executable, *sys.argv], "python_version": sys.version, "python_sha256": sha(sys.executable),
        "manual_diff_review": "One bounded completed-task observation correction; existing fixture-local automatic-variable renames, UTF8 BOM and Verbose retry diagnostic preserve assertions; focused fault/checker/harness regressions strengthen truthful refusal. No actionable findings.",
        "limitations": "Independent report/source/diff inspection; this does not supply missing Explorer/manual acceptance or broader T23 native coverage.",
    }
    try:
        assert not git("status", "--porcelain=v1"), "Clean source checkpoint required"
        labels = git("-c", "core.quotepath=false", "ls-files", "--cached", "--others", "--exclude-standard", "--", "WinPDFMerge.ps1", "WinPDFMerge.bat", "src", "tests", "tools/test", "PSScriptAnalyzerSettings.psd1").splitlines()
        current = {relative: sha(REPO / relative) for relative in labels}
        before = subprocess.check_output(["git", "-C", str(REPO), "show", BASELINE + ":src/WinPDFMerge.Helpers.ps1"]).decode("utf-8-sig").replace("\r\n", "\n")
        actual = (REPO / "src/WinPDFMerge.Helpers.ps1").read_text(encoding="utf-8-sig").replace("\r\n", "\n")
        old = "        # Result is read only after IsCompleted; this never waits for a pipe.\n        $length = $State.PendingRead.Result"
        new = "        # Observe completed task faults through a method: PowerShell property\n        # access can hide Task.Result getter exceptions. IsCompleted ensures\n        # GetResult never waits for a pipe while IO failures reach this catch.\n        $length = $State.PendingRead.GetAwaiter().GetResult()"
        assert before.count(old) == 1 and before.replace(old, new) == actual, "Runtime change exceeds reviewed one-method defect correction"
        assert not git("diff", BASELINE, commit, "--", "WinPDFMerge.ps1", "WinPDFMerge.bat"), "Established entry and launcher changed"
        review["runtime_scope_verified"] = True
        review["source_bindings"] = current
        expected_shells = {"ps51": ("5.1.26100.9444", "Desktop"), "ps7": ("7.6.6", "Core")}
        total_checks = 0
        for root in args.run:
            root = root.resolve()
            meta, aggregate, rows = load(root / "metadata.json"), load(root / "aggregate.json"), load(root / "runs.json")
            assert meta["commit_under_test"] == aggregate["commit_under_test"] == commit
            assert meta["dirty_worktree"] is aggregate["dirty_worktree"] is False
            assert aggregate["result"] == "pass"
            assert {row["path"]: row["sha256"] for row in meta["sources"]} == current
            guard = load(root / "source-guard.json")
            assert guard["result"] == "pass" and len(guard["bindings"]) == len(current)
            assert all(row["before_sha256"] == row["after_sha256"] == current[row["path"]] for row in guard["bindings"])
            run_result = {"root": str(root), "shell": meta["shell"], "metadata_sha256": sha(root / "metadata.json"), "tiers": []}
            for row in rows:
                tier = row["tier"]
                summary = load(root / (tier + ".summary.json"))
                assert row["exit_code"] == 0 and summary == row["summary"]
                assert summary["commit_under_test"] == commit and summary["dirty_worktree"] is False
                assert summary["source_unchanged"] is True and summary["runner_error"] is None and summary["result"] == "pass"
                assert summary["source_start"] == summary["source_end"]
                assert summary["source_start"]["status"] == []
                assert {item["path"]: item["sha256"] for item in summary["source_start"]["sources"]} == current
                assert (summary["shell_version"], summary["shell_edition"]) == expected_shells[meta["shell"]]
                assert summary["pester_version"] == "6.2.0" and summary["execution_policy"] == "RemoteSigned" and summary["process_64_bit"] is True
                assert type(summary["passed"]) is int and summary["passed"] == summary["total"] > 0
                assert all(type(summary[key]) is int and summary[key] == 0 for key in ("failed", "failed_blocks", "failed_containers", "skipped", "inconclusive", "not_run"))
                stdout = root / (tier + ".stdout.txt")
                stderr = root / (tier + ".stderr.txt")
                assert sha(stdout) == row["stdout_sha256"] and sha(stderr) == row["stderr_sha256"]
                console = re.findall(r"Tests Passed: (\d+), Failed: (\d+), Skipped: (\d+), Inconclusive: (\d+), NotRun: (\d+)", stdout.read_text(encoding="utf-8-sig"))
                assert len(console) == 1 and tuple(map(int, console[0])) == (summary["passed"], 0, 0, 0, 0)
                xml_path = root / (tier + ".results.xml")
                xml = ET.parse(xml_path).getroot()
                assert xml.tag == "test-results" and int(xml.attrib["total"]) == summary["total"]
                assert all(int(xml.attrib[key]) == 0 for key in ("errors", "failures", "not-run", "inconclusive", "ignored", "skipped", "invalid"))
                leaves = xml.findall(".//test-case")
                assert len(leaves) == summary["total"] and all(leaf.attrib.get("result") == "Success" and leaf.attrib.get("executed") == "True" for leaf in leaves)
                run_result["tiers"].append({"tier": tier, "passed": summary["passed"], "summary_sha256": sha(root / (tier + ".summary.json")), "xml_sha256": sha(xml_path), "verified": True})
                total_checks += summary["passed"]
            assert sum(row["passed"] for row in run_result["tiers"]) == aggregate["passed"]
            assert len(rows) == aggregate["tiers"]
            review["runs"].append(run_result)
        expected_static = {str((REPO / relative).resolve()) for relative in labels if Path(relative).suffix.lower() in (".ps1", ".psm1", ".psd1")}
        for path in args.static:
            report = load(path)
            execution = load(path.with_name('execution.json'))
            assert execution['commit_under_test'] == commit and execution['dirty_worktree'] is False and execution['exit_code'] == 0
            assert sha(path.with_name('stdout.txt')) == execution['stdout_sha256'] and sha(path.with_name('stderr.txt')) == execution['stderr_sha256']
            assert report["scope"] == "all-maintained-powershell" and report["result"] == "pass"
            assert report["commit_under_test"] == report["commit_after"] == commit and report["dirty_worktree"] is False
            assert report["analyzer_version"] == "1.25.0" and report["syntax_target_versions"] == ["5.1", "7.6"]
            assert (report['shell_version'], report['shell_edition']) == expected_shells[execution['shell']]
            assert report['process_64_bit'] is True and report['execution_policy'] == 'RemoteSigned'
            assert all(row['sha256'] == sha(row['path']) for row in report['analyzer_module_files'])
            assert set(row["path"] for row in report["files"]) == expected_static
            assert report["files_checked"] == report["parser_passed"] == report["analyzer_passed"] == len(expected_static)
            assert all(report[key] == 0 for key in ("parser_failed", "parser_errors", "analyzer_failed", "analyzer_not_run", "skipped", "selected_errors", "selected_warnings", "selected_information", "selected_suppressions", "source_guard_failed", "checkpoint_guard_failed"))
            assert all(row["unchanged"] and row["before_sha256"] == row["after_sha256"] == sha(row["path"]) for row in report["source_bindings"])
            assert all(row["parser_result"] == row["analyzer_result"] == "pass" and not row["selected_findings"] and not row["suppressed_findings"] and row["sha256"] == sha(row["path"]) for row in report["files"])
            assert sum(len(row["advisory_findings"]) for row in report["files"]) == sum(report[key] for key in ("advisory_errors", "advisory_warnings", "advisory_information"))
            review["static"].append({"path": str(path.resolve()), "sha256": sha(path), "files_checked": report["files_checked"], "selected_rules": len(report["selected_rules"]), "advisory_errors": report["advisory_errors"], "advisory_warnings": report["advisory_warnings"], "advisory_information": report["advisory_information"], "verified": True})
        review["pester_checks_verified"] = total_checks
        review["result"] = "pass"
    except Exception as error:
        review["result"] = "fail"
        review["error"] = repr(error)
        raise
    finally:
        (output / "review.json").write_text(json.dumps(review, indent=2) + "\n", encoding="utf-8")
        if review['result'] == 'pass':
            shutil.copyfile(output / 'review.json', Path(__file__).resolve().parent / 'C1-reports-review.json')
        print("Review report: " + str(output), flush=True)


if __name__ == "__main__":
    main()
