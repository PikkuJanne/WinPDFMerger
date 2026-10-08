"""T20 independent source/document policy review; writes only its ignored receipts."""
import argparse
import datetime
import hashlib
import json
import platform
import re
import subprocess
from pathlib import Path

START = "440ea7f6f4d92892fd6302d349848955b440f93d"
ORIGINAL = {
    "LICENSE": "714ffa7a21614e637d7dbb17a2e86e4575d7ecd2b6d36b5b67ab3fdcc4193477",
    "WinPDFMerge.ps1": "e702b0b1583ab841589494cc524da1d9bd4148cf8849577793b08fd8f54b641d",
    "WinPDFMerge.bat": "288287beefea974c64d39b34266332d83112b9106c952efd8c304edfedebff83",
    "src/WinPDFMerge.Helpers.ps1": "3df10eafeac04493b06300d276cff05a9c5ef196a01b7b157a6afd3d2768eaa7",
}
PUBLIC = ["README.md", "SECURITY.md", "docs/DEPENDENCIES.md", "docs/USAGE.md", "docs/TROUBLESHOOTING.md", "docs/PDF_LIMITATIONS.md", "docs/EMAIL_PRESETS.md"]

def sha(data):
    return hashlib.sha256(data).hexdigest()

def git(root, *args):
    return subprocess.run(["git", "-C", str(root), *args], capture_output=True, check=True, text=True).stdout.strip()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--label", required=True)
    parser.add_argument("--expected-head", required=True)
    parser.add_argument("--require-clean", action="store_true")
    args = parser.parse_args()
    if not re.fullmatch(r"[a-zA-Z0-9-]+", args.label):
        raise ValueError("label must be a simple owned receipt directory name")
    root = Path(__file__).resolve().parents[3]
    output = Path(__file__).resolve().parent / args.label
    output.mkdir(exist_ok=False)
    paths = PUBLIC + list(ORIGINAL)
    original_bytes = {name: (root / name).read_bytes() for name in paths}
    text = {name: re.sub(r"\s+", " ", data.decode("utf-8-sig")) for name, data in original_bytes.items()}
    bindings = []
    for name, data in original_bytes.items():
        snap = output / (name.replace("/", "-") + ".source.txt")
        snap.write_bytes(data)
        bindings.append({"path": name, "sha256": sha(data), "bytes": len(data), "retained_source": snap.relative_to(root).as_posix()})
    checks = []
    def check(name, condition, observation, references):
        checks.append({"id": name, "result": "pass" if condition else "fail", "observation": observation, "evidence": references})
    head = git(root, "rev-parse", "HEAD")
    dirty = bool(git(root, "status", "--porcelain=v1"))
    check("head", head == args.expected_head, "The review is bound to the explicitly requested current Git head; dirty source hashes are separate.", ["git rev-parse HEAD"])
    check("clean-if-required", not args.require_clean or not dirty, "Clean status is required only for the clean checkpoint review.", ["git status --porcelain=v1"])
    for name, digest in ORIGINAL.items():
        check("unchanged-" + name.replace("/", "-"), sha(original_bytes[name]) == digest, "Exact LICENSE/application/BAT/helper bytes match the T19 starting source.", [name, "source-audit.json"])
    dependencies = text["docs/DEPENDENCIES.md"]
    security = text["SECURITY.md"]
    readme = text["README.md"]
    usage = text["docs/USAGE.md"]
    trouble = text["docs/TROUBLESHOOTING.md"]
    all_public = " ".join(text[p] for p in PUBLIC)
    runtime = " ".join(text[p] for p in ORIGINAL if p != "LICENSE")
    helper = text["src/WinPDFMerge.Helpers.ps1"]
    check("mit-separate", "copyright 2025 Janne Vuorela" in dependencies and "does not relicense" in dependencies, "Project MIT attribution and unchanged copyright are separate from dependency licenses.", ["LICENSE", "docs/DEPENDENCIES.md"])
    check("pdftk-vendor-terms", "GPL version 2" in dependencies and "https://www.pdflabs.com/docs/pdftk-license/" in dependencies, "PDFtk GPLv2 and vendor redistribution information match the official vendor source.", ["docs/DEPENDENCIES.md", "https://www.pdflabs.com/docs/pdftk-license/"])
    check("gs-vendor-terms", "AGPL version 3 or commercial" in dependencies and "https://artifex.com/licensing" in dependencies, "Ghostscript AGPLv3/commercial alternatives are separate from project MIT.", ["docs/DEPENDENCIES.md", "https://artifex.com/licensing"])
    check("no-license-exemption", "not a legal determination" in dependencies and "redistribution is exempt" in dependencies, "The prose makes no legal integration/redistribution exemption claim.", ["docs/DEPENDENCIES.md"])
    check("independent-dependencies", "application package policy excludes third-party executables" in dependencies and "not application runtime requirements" in dependencies and "Python is not needed" in readme, "Declared package policy excludes vendor binaries without asserting an unbuilt package is verified; development libraries are not runtime requirements.", ["README.md", "docs/DEPENDENCIES.md", "docs/codex/SECURITY_AND_DEPENDENCIES.md"])
    check("required-pdftk-optional-gs", "PDFtk is required" in readme and "it is optional" in readme and "Ghostscript is optional" in security, "Required PDFtk and optional Ghostscript are clearly separated.", ["README.md", "SECURITY.md", "WinPDFMerge.ps1"])
    check("runtime-local-source", not re.search(r"Invoke-WebRequest|Invoke-RestMethod|WebClient|HttpClient|Invoke-Expression", runtime, re.I), "Actual entry/helper/BAT contain no application network acquisition, telemetry endpoint, or shell evaluation call.", ["WinPDFMerge.ps1", "WinPDFMerge.bat", "src/WinPDFMerge.Helpers.ps1"])
    check("local-scope-prose", "no network calls" in dependencies and "uploads, telemetry" in security and "separate actions" in security, "Runtime processing is local; dependency/repository/reporting actions are accurately distinguished.", ["docs/DEPENDENCIES.md", "SECURITY.md"])
    check("native-safety-not-sandbox", "do not make this tool a sandbox or PDF malware sanitizer" in security and "user account's permissions" in security and "'-dSAFER'" in helper and "'-dNOSAFER'" not in helper, "Fixed GS restrictions stay enabled; native parser user permissions and absence of hostile-file isolation are disclosed.", ["SECURITY.md", "src/WinPDFMerge.Helpers.ps1", "https://ghostscript.readthedocs.io/en/latest/Use.html#dsafer"])
    check("child-only-environment", "$startInfo.EnvironmentVariables.Remove($name)" in helper and "caller\u2019s environment".replace("\u2019", "'") in usage, "GS_OPTIONS removal is limited to selected child environment; caller environment remains intact.", ["docs/USAGE.md", "src/WinPDFMerge.Helpers.ps1"])
    check("confidential-log-details", all(word in security for word in ("PDF metadata", "command arguments", "native stdout/stderr", "process ID", "known temporary paths")), "Log streams and ownership markers can expose names, paths, metadata and run/process facts.", ["SECURITY.md", "Write-NativeProcessLog", "New-PdfStaging"])
    check("ordinary-permissions", "not automatically redacted, encrypted" in security and "ordinary directory/filesystem access controls" in security and "not an access-isolated or encrypted directory" in security, "No special encryption/ACL/redaction or private-staging isolation is promised.", ["SECURITY.md", "src/WinPDFMerge.Helpers.ps1"])
    check("retention-and-external-storage", "no log-retention or automatic deletion schedule" in security and "sync/backup service or network folder" in security, "Owner controls retention and storage; external services can independently copy local data.", ["SECURITY.md"])
    check("sanitized-sharing", "copy the log" in security and "including both native streams" in security and "Do not publish private PDFs" in security, "Public reports require a reviewed sanitized copy and synthetic reproduction; private documents/secrets are excluded.", ["SECURITY.md", "docs/TROUBLESHOOTING.md"])
    check("safe-orphan-cleanup", "all merge runs have stopped" in trouble and "abandoned run you own" in trouble and "Do not sweep by filename prefix" in trouble, "Manual cleanup preserves active/unknown/foreign staging and validated finals.", ["docs/TROUBLESHOOTING.md", "src/WinPDFMerge.Helpers.ps1"])
    check("unsigned-accurate", "scripts are unsigned" in security and "process-only" in security and "cannot override Group Policy" in security, "Actual NotSigned PS source status and existing BAT process Bypass limits are accurately disclosed.", ["SECURITY.md", "Get-AuthenticodeSignature read-only observation", "WinPDFMerge.bat"])
    check("no-persistent-policy-workaround", not re.search(r"(?im)^\s*Set-ExecutionPolicy\b", "\n".join(original_bytes[p].decode("utf-8-sig") for p in PUBLIC)) and "Do not disable security software" in trouble and "change user/machine policy" in trouble, "No instructions weaken persistent execution policy, enterprise policy or security software.", ["docs/TROUBLESHOOTING.md", "SECURITY.md", "https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies?view=powershell-5.1"])
    check("checksum-not-authentication", "not a digital signature, proof of publisher identity or a malware verdict" in security and "same site as a download" in security, "SHA-256 integrity is distinguished from signatures, publisher identity and safety.", ["SECURITY.md"])
    check("no-admin-normal-use", "normal use requires no administrator" in dependencies and "do not disable protections, request elevation for normal use" in security, "Normal standard-user use is documented without elevation or control disablement.", ["docs/DEPENDENCIES.md", "SECURITY.md"])
    check("tested-scope-limits", "do not certify every build or complete release acceptance" in dependencies and "Windows 10, ARM, 32-bit hosts and live UNC shares are unvalidated" in dependencies and "Version numbers here identify tests" in dependencies, "Known Windows/PS/native observations do not become broad support, future-security, Explorer or release certification.", ["docs/DEPENDENCIES.md", "README.md", "docs/codex/COMPATIBILITY_MATRIX.md"])
    check("current-vendor-maintenance", "https://www.ghostscript.com/releases/cve/index.html" in dependencies and "/about/cve.html" not in dependencies and "check current vendor" in dependencies, "Correct current vendor CVE route replaces the actual 404 and updates remain an owner/vendor action.", ["docs/DEPENDENCIES.md", "https://www.ghostscript.com/releases/cve/index.html"])
    check("reporting-truthful", "No dedicated private vulnerability reporting route is currently configured" in security and "any later change" in security and "wait for that channel" in security and "No support SLA" in security, "Fresh API disabled-state observation is disclosed without inventing a private channel, exploit disclosure request or response deadline.", ["SECURITY.md", "gh api repos/PikkuJanne/WinPDFMerger/private-vulnerability-reporting"])
    check("feature-safety-limits", "Neither output guarantees PDF/A, signature validity, accessibility, universal feature retention, archival certification or malware removal" in readme and "signed/feature-rich originals" in security, "Preservation/sanitization limitations remain explicit and original retention is advised.", ["README.md", "SECURITY.md", "docs/PDF_LIMITATIONS.md"])
    check("no-website-scope", "website work, cloud conversion, OCR, a GUI editor and an installer are outside its scope" in readme and "automatic updates" in security, "No website/cloud/OCR/editor/installer/update expansion is introduced.", ["README.md", "SECURITY.md", "AGENTS.md"])
    check("release-not-claimed", "release is not yet published" in readme and "not publication evidence" in security, "Public instructions do not claim an available tested/public v1.0.0 ZIP or completed release.", ["README.md", "SECURITY.md"])
    check("public-privacy", not re.search(r"(?i)(?:C:[/\\]Users[/\\]|Bearer\s+[A-Za-z0-9]|gh[pousr]_[A-Za-z0-9]{20}|AKIA[A-Z0-9]{16}|-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY)", all_public), "Candidate public prose contains synthetic examples and no observed profile paths or standard secret-key/token markers.", PUBLIC)
    after = {name: sha((root / name).read_bytes()) for name in paths}
    check("stable-read", all(after[name] == sha(original_bytes[name]) for name in paths), "All reviewed source/document bytes remained unchanged through the snapshot/check pass.", ["before/after SHA-256 bindings"])
    report = {
        "schema_version": 1, "task": "T20", "acceptance_case": "AC047", "reviewer": "independent-policy-source-review",
        "observed_at_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(), "label": args.label,
        "starting_commit": START, "commit_at_review": head, "dirty_worktree": dirty,
        "scope": "Source/document/current-official-reference review only. Regex consistency checks support manually inspected claims; no application, native parser, Windows manual/Explorer, redistribution clearance, full secret audit or release acceptance.",
        "environment": {"os_family": platform.system(), "review_python": platform.python_version()},
        "source_bindings": bindings, "checks": checks,
        "check_count": len(checks), "passed": sum(c["result"] == "pass" for c in checks),
        "blockers": [c for c in checks if c["result"] != "pass"],
        "result": "pass" if all(c["result"] == "pass" for c in checks) else "fail",
        "official_reference_observations": "source-audit.json plus final session verification of PowerShell project MIT and Microsoft installation links; observed 2026-10-08, not perpetual latest/safe claims.",
        "actual_read_only_observations": {"authenticode": "WinPDFMerge.ps1 and helper NotSigned/no signer", "private_vulnerability_reporting": {"enabled": False}, "old_cve_route_http_status": 404, "correct_cve_route": "https://www.ghostscript.com/releases/cve/index.html"},
        "resolved_finding": "Original final-candidate dependencies used an obsolete 404 CVE route; root corrected it before this exact candidate review. No runtime or license change was required.",
        "limitations": ["Public secret scan is limited to reviewed candidate prose and common markers, not a T25 full tracked-history audit.", "No privileged/security setting changes, install, runtime/native execution, paid signing, publishing or Git mutation performed.", "Clean C1 rebinding is required before using dirty-candidate results as clean implementation evidence.", "Vendor licenses are separate primary references; no legal compatibility/exemption determination is made."]
    }
    report_path = output / "review.json"
    report_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"path": report_path.relative_to(root).as_posix(), "sha256": sha(report_path.read_bytes()), "checks": report["check_count"], "passed": report["passed"], "result": report["result"], "blocker_ids": [c["id"] for c in report["blockers"]]}))
    return 0 if report["result"] == "pass" else 1

if __name__ == "__main__":
    raise SystemExit(main())
