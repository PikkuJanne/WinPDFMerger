"""Independent source-bound before/after probe for the completed Task read defect."""
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import uuid

REPO = Path(__file__).resolve().parents[3]
BASELINE = "75df26187202104995e58f020c2442d973c7b075"
SHELLS = {
    "ps51": Path(r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"),
    "ps7": Path(r"<USERPROFILE>\AppData\Local\WinPDFMergerDevCache\T09-ps7-6719cb8d846e47c9bf69d6c7f5948a6a\portable\pwsh.exe"),
}
DRIVER = r"""param([string]$HelperPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. $HelperPath
$stream = [IO.MemoryStream]::new([byte[]]@())
$reader = [IO.StreamReader]::new($stream)
try {
    $state = New-NativeStreamCapture -Reader $reader
    [void]$state.Text.Append('T22 retained prefix')
    $completion = New-Object 'System.Threading.Tasks.TaskCompletionSource[int]'
    $completion.SetException([IO.IOException]::new('T22 independently controlled completed read fault'))
    $state.PendingRead = $completion.Task
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $received = Receive-NativeStreamCapture -State $state -MaximumCaptureCharacters 8192
    $watch.Stop()
    [ordered]@{
        observed_at_utc=[DateTime]::UtcNow.ToString('o')
        shell_version=$PSVersionTable.PSVersion.ToString()
        shell_edition=$PSVersionTable.PSEdition
        process_64_bit=[Environment]::Is64BitProcess
        execution_policy=(Get-ExecutionPolicy).ToString()
        helper_sha256=(Get-FileHash -LiteralPath $HelperPath -Algorithm SHA256).Hash.ToLowerInvariant()
        received=$received
        error=$state.Error
        closed=$state.Closed
        pending_read_present=($null -ne $state.PendingRead)
        retained_text=$state.Text.ToString()
        truncated=$state.Truncated
        elapsed_ms=$watch.ElapsedMilliseconds
    } | ConvertTo-Json
    # Observe the deliberately faulted test task even in the historical path.
    [void]$completion.Task.Exception
} finally { $reader.Dispose(); $stream.Dispose() }
"""


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def utc():
    return datetime.now(timezone.utc).isoformat()


def write(path, value):
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def main():
    output = REPO / "tests/.work/T22-review" / ("capture-" + uuid.uuid4().hex)
    output.mkdir(parents=True)
    shutil.copyfile(__file__, output / "review_capture.py")
    historical = subprocess.run(["git", "-C", str(REPO), "show", BASELINE + ":src/WinPDFMerge.Helpers.ps1"], check=True, stdout=subprocess.PIPE).stdout
    (output / "historical-helpers.ps1").write_bytes(historical)
    shutil.copyfile(REPO / "src/WinPDFMerge.Helpers.ps1", output / "current-helpers.ps1")
    driver = output / "capture-probe.ps1"
    driver.write_text(DRIVER, encoding="utf-8")
    reviewed_commit = subprocess.run(["git", "-C", str(REPO), "rev-parse", "HEAD"], check=True, stdout=subprocess.PIPE).stdout.decode().strip()
    dirty_status = subprocess.run(["git", "-C", str(REPO), "status", "--porcelain=v1", "--untracked-files=all"], check=True, stdout=subprocess.PIPE).stdout.decode().splitlines()
    review = {
        "task": "T22", "started_at_utc": utc(), "historical_commit": BASELINE,
        "commit_under_review": reviewed_commit, "producer_sha256": sha(__file__),
        "dirty_worktree": bool(dirty_status), "dirty_status_at_start": dirty_status,
        "driver_sha256": sha(driver), "python": sys.version, "python_sha256": sha(sys.executable),
        "historical_source_sha256": sha(output / "historical-helpers.ps1"),
        "current_source_sha256": sha(output / "current-helpers.ps1"), "observations": [],
        "limitations": "Controlled completed Task IO fault against imported helper definitions only; no native engines or application orchestration.",
    }
    try:
        for shell_label, executable in SHELLS.items():
            for version in ("historical", "current"):
                destination = output / (shell_label + "-" + version)
                destination.mkdir()
                source = output / (version + "-helpers.ps1")
                argv = [str(executable), "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File", str(driver), "-HelperPath", str(source)]
                environment = {k: v for k, v in os.environ.items() if k.upper() != "PSMODULEPATH"}
                started = utc()
                result = subprocess.run(argv, cwd=REPO, env=environment, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
                (destination / "stdout.txt").write_bytes(result.stdout)
                (destination / "stderr.txt").write_bytes(result.stderr)
                invocation = {"argv": argv, "started_at_utc": started, "completed_at_utc": utc(), "exit_code": result.returncode, "stdout_sha256": sha(destination / "stdout.txt"), "stderr_sha256": sha(destination / "stderr.txt"), "source_sha256": sha(source), "child_environment_remove": ["PSMODULEPATH"], "process_only_policy": "RemoteSigned"}
                write(destination / "invocation.json", invocation)
                assert result.returncode == 0, result.stderr.decode("utf-8", "replace")
                observation = json.loads(result.stdout.decode("utf-8-sig"))
                assert observation["helper_sha256"] == sha(source)
                assert observation["received"] is True
                assert observation["retained_text"] == "T22 retained prefix"
                assert observation["truncated"] is False
                assert observation["elapsed_ms"] < 1000
                if version == "historical":
                    assert observation["closed"] is False and observation["error"] is None
                    assert observation["pending_read_present"] is True
                else:
                    assert observation["closed"] is True
                    assert "independently controlled completed read fault" in observation["error"]
                    assert observation["pending_read_present"] is False
                observation.update({"shell": shell_label, "version": version, "verified": True})
                write(destination / "observation.json", observation)
                review["observations"].append(observation)
                print(json.dumps({"shell": shell_label, "version": version, "closed": observation["closed"], "error": observation["error"], "verified": True}), flush=True)
        review["current_source_sha256_end"] = sha(REPO / "src/WinPDFMerge.Helpers.ps1")
        assert review["current_source_sha256_end"] == review["current_source_sha256"]
        review["result"] = "pass"
    except Exception as error:
        review["result"] = "fail"
        review["error"] = repr(error)
        raise
    finally:
        review["completed_at_utc"] = utc()
        write(output / "review.json", review)
        print("Review report: " + str(output), flush=True)


if __name__ == "__main__":
    main()
