"""Independent T22 audit of the development runner using synthetic local repos.

These controlled Pester outcomes exercise reporting, never PDF engines or manual
Windows acceptance. No dependency is acquired and no persistent policy changes.
"""
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import uuid
import xml.etree.ElementTree as ET

REPO = Path(__file__).resolve().parents[3]
PESTER = Path(r"<USERPROFILE>\.cache\WinPDFMerger-T03\pester-6.2.0-fe69858b0b7b4b92ae85caf8dd8fcc67\module\Pester.psd1")
SHELLS = {
    "ps51": Path(r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"),
    "ps7": Path(r"<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe"),
}
COPIED = ("tools/test/Invoke-Tests.ps1", "tools/test/TestRunSupport.ps1", "tests/TestDependencies.psd1")
SCENARIOS = {
    "passed": "Describe 'T22 controlled pass' { It 'passes' { 1 | Should -Be 1 } }\n",
    "failed": "Describe 'T22 controlled failure' { It 'fails' { 1 | Should -Be 2 } }\n",
    "skipped": "Describe 'T22 controlled skip' { It 'is visibly skipped' -Skip { 1 | Should -Be 1 } }\n",
    "not_run": "Describe 'T22 controlled setup fault' { BeforeAll { throw 'T22 synthetic setup fault' }; It 'cannot run' { 1 | Should -Be 1 } }\n",
    "discovery": "Describe 'T22 controlled discovery parser fault' { It 'has no closing body' {\n",
    "empty": "# T22 synthetic empty Pester suite\n",
    "source_mutation": "Describe 'T22 controlled source mutation' { It 'mutates only synthetic source' { [IO.File]::WriteAllText((Join-Path $PSScriptRoot '../../WinPDFMerge.ps1'), '# T22 controlled modified bytes'); 1 | Should -Be 1 } }\n",
    "xml_write_fault": "Describe 'T22 controlled report IO fault' { BeforeAll { $runs=Join-Path $PSScriptRoot '../.work/pester'; $run=@(Get-ChildItem -LiteralPath $runs -Directory); if($run.Count -ne 1){throw 'Expected one controlled report directory'}; [void][IO.Directory]::CreateDirectory((Join-Path $run[0].FullName 'results.xml')) }; It 'passes before report writing throws' { 1 | Should -Be 1 } }\n",
}


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def utc():
    return datetime.now(timezone.utc).isoformat()


def write_json(path, value):
    Path(path).write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def run(argv, destination, cwd):
    environment = {k: v for k, v in os.environ.items() if k.upper() != "PSMODULEPATH"}
    started = utc()
    result = subprocess.run([str(x) for x in argv], cwd=cwd, env=environment, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=90)
    (destination / "stdout.txt").write_bytes(result.stdout)
    (destination / "stderr.txt").write_bytes(result.stderr)
    invocation = {
        "started_at_utc": started, "completed_at_utc": utc(), "argv": [str(x) for x in argv],
        "cwd": str(cwd), "exit_code": result.returncode, "stdout_sha256": digest(destination / "stdout.txt"),
        "stderr_sha256": digest(destination / "stderr.txt"), "child_environment_remove": ["PSMODULEPATH"],
        "process_only_policy": "RemoteSigned", "no_dependency_acquisition": True,
    }
    write_json(destination / "invocation.json", invocation)
    return result


def git(path, *args):
    result = subprocess.run(["git", "-C", str(path), *args], stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=20)
    if result.returncode:
        raise AssertionError("Synthetic Git operation failed: " + result.stderr.decode("utf-8", "replace"))
    return result.stdout.decode("utf-8").strip()


def audit_scenario(shell_label, scenario, output):
    destination = output / shell_label / scenario
    synthetic = destination / "synthetic repo with spaces"
    synthetic.mkdir(parents=True)
    for relative in COPIED:
        target = synthetic / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(REPO / relative, target)
    (synthetic / "WinPDFMerge.ps1").write_text("# T22 inert synthetic application source\n", encoding="utf-8")
    (synthetic / ".gitignore").write_text("tests/.work/\n", encoding="utf-8")
    suite = synthetic / "tests/unit/Synthetic.Tests.ps1"
    suite.parent.mkdir(parents=True)
    suite.write_text(SCENARIOS[scenario], encoding="utf-8")
    git(synthetic, "init", "--quiet")
    git(synthetic, "add", "--", *COPIED, "WinPDFMerge.ps1", ".gitignore", "tests/unit/Synthetic.Tests.ps1")
    git(synthetic, "-c", "user.name=T22Synthetic", "-c", "user.email=t22@example.invalid", "commit", "--quiet", "--no-gpg-sign", "-m", "T22 controlled runner audit")
    tested_commit = git(synthetic, "rev-parse", "HEAD")
    copied_bindings = {relative: digest(synthetic / relative) for relative in COPIED}
    result = run([SHELLS[shell_label], "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File", synthetic / "tools/test/Invoke-Tests.ps1", "-PesterModulePath", PESTER, "-Tier", "Unit"], destination, synthetic)
    reports = list((synthetic / "tests/.work/pester").glob("*/summary.json"))
    assert len(reports) == 1, f"{shell_label}/{scenario}: expected one retained truthful summary"
    summary = json.loads(reports[0].read_text(encoding="utf-8-sig"))
    shutil.copyfile(reports[0], destination / "summary.json")
    xml_path = reports[0].with_name("results.xml")
    xml_root = None
    if xml_path.is_file():
        shutil.copyfile(xml_path, destination / "results.xml")
        xml_root = ET.parse(xml_path).getroot()
    assert summary["commit_under_test"] == tested_commit
    assert summary["dirty_worktree"] is False
    assert summary["source_start"]["status"] == []
    assert summary["shell_version"] == ("5.1.26100.9444" if shell_label == "ps51" else "7.6.6")
    assert summary["execution_policy"] == "RemoteSigned"
    assert summary["pester_version"] == "6.2.0"
    assert summary["process_64_bit"] is True
    start = {item["path"]: item["sha256"] for item in summary["source_start"]["sources"]}
    for relative, sha in copied_bindings.items():
        assert start[relative] == sha
    if scenario == "passed":
        assert result.returncode == 0 and summary["result"] == "pass"
        assert summary["passed"] == summary["total"] == 1
    else:
        assert result.returncode == 1 and summary["result"] == "fail"
    if scenario == "failed":
        assert summary["failed"] == 1 and summary["total"] == 1
    elif scenario == "skipped":
        assert summary["skipped"] == 1 and summary["passed"] == 0
    elif scenario == "not_run":
        assert summary["not_run"] == 1 and summary["failed_blocks"] == 1 and summary["passed"] == 0
    elif scenario == "discovery":
        assert summary["failed_containers"] == 1 and summary["passed"] == 0
    elif scenario == "empty":
        assert summary["total"] in (None, 0) and summary["passed"] in (None, 0)
    elif scenario == "source_mutation":
        assert summary["source_unchanged"] is False and summary["passed"] == summary["total"] == 1
        assert summary["source_end"]["status"] != []
    elif scenario == "xml_write_fault":
        assert summary["runner_error"] and summary["total"] is None and summary["passed"] is None
        assert xml_root is None
    if scenario != "source_mutation":
        assert summary["source_unchanged"] is True
    if xml_root is not None:
        assert xml_root.tag == "test-results"
        assert int(xml_root.attrib["total"]) == summary["total"]
        leaves = xml_root.findall(".//test-case")
        assert len(leaves) == summary["total"]
        if scenario == "passed":
            assert all(x.attrib.get("result") == "Success" for x in leaves)
        if scenario == "failed":
            assert any(x.attrib.get("result") == "Failure" for x in leaves)
        if scenario == "skipped":
            assert any(x.attrib.get("executed") == "False" for x in leaves)
    observation = {
        "shell": shell_label, "scenario": scenario, "verified": True,
        "scope": "controlled synthetic Pester reporting; no application/native/manual acceptance",
        "synthetic_commit_under_test": tested_commit, "exit_code": result.returncode,
        "counts": {key: summary[key] for key in ("passed", "failed", "failed_blocks", "failed_containers", "skipped", "not_run", "total")},
        "source_unchanged": summary["source_unchanged"], "runner_error": summary["runner_error"],
        "copied_sources": copied_bindings, "xml_present": xml_root is not None,
    }
    write_json(destination / "observation.json", observation)
    print(json.dumps({"shell": shell_label, "scenario": scenario, "counts": observation["counts"], "verified": True}), flush=True)
    return observation


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--phase", default="preparation")
    parser.add_argument("--shell", choices=tuple(SHELLS))
    arguments = parser.parse_args()
    output = REPO / "tests/.work/T22-review" / (arguments.phase + "-" + uuid.uuid4().hex)
    output.mkdir(parents=True)
    shutil.copyfile(__file__, output / "review_harness.py")
    review = {
        "task": "T22", "phase": arguments.phase, "started_at_utc": utc(),
        "producer_sha256": digest(__file__), "python": sys.version,
        "python_sha256": digest(sys.executable), "commit_under_review": git(REPO, "rev-parse", "HEAD"),
        "dirty_status_at_start": git(REPO, "status", "--porcelain=v1", "--untracked-files=all").splitlines(),
        "source_bindings": {relative: digest(REPO / relative) for relative in COPIED},
        "pester_manifest_sha256": digest(PESTER), "observations": [],
        "limitations": "Synthetic local repos and controlled Pester outcomes; no native engines, actual application entry, Explorer, or manual acceptance.",
    }
    write_json(output / "review.json", review)
    try:
        for shell_label in ([arguments.shell] if arguments.shell else SHELLS):
            for scenario in SCENARIOS:
                review["observations"].append(audit_scenario(shell_label, scenario, output))
        review["source_bindings_end"] = {relative: digest(REPO / relative) for relative in COPIED}
        review["sources_unchanged"] = review["source_bindings"] == review["source_bindings_end"]
        assert review["sources_unchanged"], "Reviewed runner source changed during independent audit"
        review["result"] = "pass"
    except Exception as error:
        review["result"] = "fail"
        review["error"] = repr(error)
        raise
    finally:
        review["completed_at_utc"] = utc()
        write_json(output / "review.json", review)
        print("Review report: " + str(output), flush=True)


if __name__ == "__main__":
    main()
