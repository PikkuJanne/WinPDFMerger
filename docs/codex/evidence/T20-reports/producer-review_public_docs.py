"""T20 read-only independent review bindings/link checks, not app acceptance."""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import re
import subprocess
import sys

ROOT = Path.cwd().resolve()
parser = argparse.ArgumentParser()
parser.add_argument("--expected-head")
parser.add_argument("--require-clean", action="store_true")
parser.add_argument("--output-root", default="tests/.work/T20-user-docs-final-review")
parser.add_argument("--ast-report", default="tests/.work/T20-user-docs-final-review/ast-review.json")
args = parser.parse_args()
OUT = (ROOT / args.output_root).resolve()
if not OUT.is_relative_to(ROOT / "tests/.work"):
    parser.error("review output must remain under ignored tests/.work")
BASELINE = "440ea7f6f4d92892fd6302d349848955b440f93d"
PUBLIC = ["README.md", "SECURITY.md", "docs/USAGE.md", "docs/TROUBLESHOOTING.md", "docs/DEPENDENCIES.md", "docs/PDF_LIMITATIONS.md", "docs/EMAIL_PRESETS.md", "LICENSE"]
RUNTIME = ["WinPDFMerge.ps1", "WinPDFMerge.bat", "src/WinPDFMerge.Helpers.ps1"]
TEXT = {p: (ROOT / p).read_text(encoding="utf-8-sig") for p in PUBLIC + RUNTIME}

def git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=ROOT, check=True, capture_output=True, text=True).stdout.strip()

def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()

def ref(path: str, phrase: str) -> str:
    at = TEXT[path].index(phrase)
    return f"{path}:{TEXT[path].count(chr(10), 0, at) + 1}"

checks: list[dict] = []
def reviewed(key: str, docs: list[tuple[str, str]], sources: list[tuple[str, str]], conclusion: str) -> None:
    # Semantic conclusions were independently read against actual source. Phrase
    # presence binds the reported decision to retained exact source snapshots.
    checks.append({"id": key, "kind": "independent semantic doc/source review", "result": "pass", "doc_refs": [ref(p, s) for p, s in docs], "actual_source_refs": [ref(p, s) for p, s in sources], "conclusion": conclusion})

reviewed("layout-and-names", [("README.md", "Keep the relative layout below"), ("README.md", "The repository is")], [("WinPDFMerge.ps1", ". (Join-Path $ScriptDir 'src/WinPDFMerge.Helpers.ps1')")], "Installation retains established launcher names and mandatory src helper; no runtime Python requirement.")
reviewed("dependency-roles", [("README.md", "PDFtk is required"), ("README.md", "an email copy; it is optional")], [("WinPDFMerge.ps1", "$pdftkPath = Find-Pdftk"), ("WinPDFMerge.ps1", "if (-not $gsPath)")], "PDFtk required; absent GS permits master-only; separately installed vendor tools.")
reviewed("four-parameters", [("docs/USAGE.md", "| `SourceFolder`"), ("docs/USAGE.md", "Invalid presets, extra positional")], [("WinPDFMerge.ps1", "[CmdletBinding(PositionalBinding=$false)]"), ("WinPDFMerge.ps1", "[ValidateSet('screen', 'ebook')]")], "Exactly one positional SourceFolder plus OutputFolder, SkipEmail and fixed screen/ebook preset; missing/invalid invocation fails early.")
reviewed("output-default-and-paths", [("README.md", "By default, output is beside"), ("docs/USAGE.md", "Source and output must identify")], [("WinPDFMerge.ps1", "if (-not $PSBoundParameters.ContainsKey('OutputFolder'))"), ("src/WinPDFMerge.Helpers.ps1", "function Assert-MergeDirectories"), ("src/WinPDFMerge.Helpers.ps1", "function Assert-MergeDirectoryPath")], "Existing writable separate destination, entry-script default, no mkdir/fallback, alias identity and reparse ancestor refusal.")
reviewed("dragdrop-and-doubleclick", [("README.md", "drag **exactly one folder**"), ("README.md", "Double-click without a folder")], [("WinPDFMerge.bat", 'if "%~1"=="" goto usage'), ("WinPDFMerge.bat", 'if not "%~2"=="" goto usage'), ("WinPDFMerge.bat", ":finish")], "BAT accepts one folder and pauses; no picker/options; physical Explorer remains unrun release gate.")
reviewed("batch-host-and-code", [("README.md", "The batch file returns the exact code"), ("docs/USAGE.md", "the batch launcher always uses")], [("WinPDFMerge.bat", '"%SystemRoot%\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"'), ("WinPDFMerge.bat", "exit /b %EC%")], "BAT fixed Windows PowerShell 5.1 NoProfile with existing process policy flag; exact code after pause.")
reviewed("direct-examples", [("README.md", "Set-Location -LiteralPath"), ("docs/USAGE.md", "open `pwsh.exe`")], [("WinPDFMerge.ps1", "[Parameter(Mandatory=$false, Position=0)]")], "Examples specify host/working-directory context, existing illustrative paths, supported flags, and immediate LASTEXITCODE; help reads without orchestration.")
reviewed("percent-boundary", [("README.md", "`cmd.exe` can expand a paired"), ("README.md", "permits this route")], [("WinPDFMerge.bat", 'set "SOURCE=%~1"'), ("WinPDFMerge.bat", "-ExecutionPolicy Bypass")], "Paired percent expansion attributed to cmd boundary; direct PowerShell alternative retains original supported route with policy precedence limits.")
reviewed("visible-top-level-inputs", [("README.md", "Subfolders and hidden files"), ("docs/USAGE.md", "One PDF is accepted")], [("src/WinPDFMerge.Helpers.ps1", "Get-ChildItem -LiteralPath $directory -Filter '*.pdf' -File"), ("src/WinPDFMerge.Helpers.ps1", "Where-Object { $_.Extension -ieq '.pdf' }")], "Non-Force top-level .pdf/.PDF; one accepted, none error; no prefix omission of old results.")
reviewed("natural-order", [("docs/USAGE.md", "ASCII digit runs sort"), ("docs/USAGE.md", "Non-ASCII digits are text")], [("src/WinPDFMerge.Helpers.ps1", "function Compare-NaturalName"), ("src/WinPDFMerge.Helpers.ps1", "function Compare-PdfInput")], "Magnitude/shorter equal numeric runs/subsequent segments, ordinal case-insensitive text and ordinal original name/path ties correctly documented.")
reviewed("freeze-and-snapshot", [("docs/USAGE.md", "The input set is frozen"), ("docs/USAGE.md", "Length and UTC modification-time")], [("WinPDFMerge.ps1", "$pdfs = @(Sort-PdfInputs"), ("src/WinPDFMerge.Helpers.ps1", "function Assert-PdfInputSnapshot"), ("WinPDFMerge.ps1", "Assert-PdfInputInventory -Inventory $inventory")], "Frozen set and length/mtime checks detect obvious changes; no filesystem snapshot or inclusion of new arrivals claimed.")
reviewed("inspection-and-password", [("docs/USAGE.md", "read-only `dump_data_utf8`"), ("docs/USAGE.md", "PDFtk-readable encryption with an empty password")], [("src/WinPDFMerge.Helpers.ps1", "function Get-PdfDocumentInspection"), ("src/WinPDFMerge.Helpers.ps1", "function Get-PdfInputInventory")], "Every input inspected; positive unambiguous count; no silent skip, repair, flatten or password workflow; authorized working copy guidance preserves originals.")
reviewed("bounded-envelope", [("docs/USAGE.md", "within the last 8192 bytes"), ("docs/USAGE.md", "within 1024 bytes")], [("src/WinPDFMerge.Helpers.ps1", "function Assert-PdfInputEnvelope")], "Header/final EOF/startxref bounded screening described without universal validity or safety guarantee.")
reviewed("native-path-and-command-limits", [("docs/USAGE.md", "260 UTF-16"), ("docs/USAGE.md", "30,000 UTF-16")], [("src/WinPDFMerge.Helpers.ps1", "function Assert-NativeCommandLength"), ("src/WinPDFMerge.Helpers.ps1", "Input file paths must be fewer than 260")], "Actual native path/serialized command bounds and shorter-path/fewer-input guidance; no long-path/chunk workaround claim.")
reviewed("unicode-and-os-scope", [("docs/USAGE.md", "CJK source/input and output paths failed"), ("README.md", "live UNC shares are unvalidated")], [("src/WinPDFMerge.Helpers.ps1", "source files were not renamed")], "Measured backend Unicode exclusions and Windows10/ARM/x86/UNC nonvalidation retained; no general compatibility claim.")
reviewed("unique-output-identity", [("docs/USAGE.md", "16-hex-character suffix"), ("docs/USAGE.md", "64 UTF-16")], [("src/WinPDFMerge.Helpers.ps1", "function New-MergeRunIdentity"), ("src/WinPDFMerge.Helpers.ps1", "function Reserve-MergeRunIdentity")], "Timestamp/random suffix shared by final outputs/log; bounded folder label and root fallback; no overwrite.")
reviewed("master-first-and-validation", [("docs/USAGE.md", "validated master is published before"), ("README.md", "Existing files are never overwritten")], [("WinPDFMerge.ps1", "$masterPublished = ($merge.OutputPublished -and $merge.OutputValidated)"), ("src/WinPDFMerge.Helpers.ps1", "function Publish-PdfStagedOutput")], "Staged native result validates before File.Move publication without replace; master published before optional GS and retained after failures.")
reviewed("strict-smaller-email", [("docs/USAGE.md", "published only if smaller"), ("README.md", "no size benefit")], [("src/WinPDFMerge.Helpers.ps1", "if ($Tool -eq 'Ghostscript' -and $outputBytes -ge $masterBytes)"), ("src/WinPDFMerge.Helpers.ps1", "'no_size_benefit'")], "Valid equal/larger derivative omitted/discarded, never published; no target email size.")
reviewed("skip-and-preset", [("README.md", "valid preset is reported as ignored"), ("docs/USAGE.md", "default `screen`")], [("WinPDFMerge.ps1", "$SkipEmail -and $PSBoundParameters.ContainsKey('EmailPreset')"), ("WinPDFMerge.ps1", "if ($SkipEmail) {\n"), ("src/WinPDFMerge.Helpers.ps1", "'ebook' { '-dPDFSETTINGS=/ebook' }")], "Skip bypasses GS discovery and invocation; explicit valid preset ignored; default screen and fixed ebook maintained.")
reviewed("gs-safety-and-environment", [("docs/USAGE.md", "`-dPDFSTOPONERROR`"), ("docs/USAGE.md", "caller\u2019s environment") if "caller\u2019s environment" in TEXT["docs/USAGE.md"] else ("docs/USAGE.md", "caller's environment")], [("src/WinPDFMerge.Helpers.ps1", "'-dBATCH', '-dNOPAUSE', '-dSAFER', '-dPDFSTOPONERROR'"), ("src/WinPDFMerge.Helpers.ps1", "$removeEnvironment = @('GS_OPTIONS')")], "Fixed pdfwrite safety flags and child-only environment removal, no safety sandbox/preservation certification.")
reviewed("t17-size-format", [("docs/USAGE.md", "`B` uses whole bytes"), ("docs/USAGE.md", "percentages one")], [("src/WinPDFMerge.Helpers.ps1", "function Format-PdfByteSize"), ("src/WinPDFMerge.Helpers.ps1", "function Get-PdfSizeReport")], "B integer, KiB+two decimals, reduction one decimal, exact byte authority, unpublished candidate reporting match real formatter.")
reviewed("code0", [("README.md", "| `0` |")], [("src/WinPDFMerge.Helpers.ps1", "$exitCode = 0"), ("src/WinPDFMerge.Helpers.ps1", "'unavailable' {"), ("src/WinPDFMerge.Helpers.ps1", "'no_size_benefit' {")], "Master plus valid smaller/skipped/unavailable/no size benefit succeeds.")
reviewed("code1", [("README.md", "| `1` |")], [("src/WinPDFMerge.Helpers.ps1", "$exitCode = 1"), ("WinPDFMerge.ps1", "No run log was created: a source")], "No validated published master: invocation/input/destination/PDFtk/merge failure or controlled prepublication cancel is code1.")
reviewed("code2", [("README.md", "| `2` |")], [("src/WinPDFMerge.Helpers.ps1", "if ($RunFailed -and $MasterPublished)"), ("WinPDFMerge.ps1", "Result logging failed:")], "Master retained after email or later log/cancel failure; valid published paths only; no failed candidate advertised.")
reviewed("t18-log-timing", [("docs/USAGE.md", "before discovery/dependency probes"), ("README.md", "before a log can be created")], [("WinPDFMerge.ps1", "Reserve-MergeRunIdentity -Identity $run"), ("WinPDFMerge.ps1", "Write-PdfRunStage -Stage 'Input discovery'"), ("src/WinPDFMerge.Helpers.ps1", "function Write-RunLog")], "UTF8 local reserved log starts after safe source/output checks and before discovery/probes; early unsafe/binding failures console-only possible.")
reviewed("t18-diagnostic-facts", [("docs/USAGE.md", "Unknown counts or unused tools"), ("docs/USAGE.md", "not an estimated")], [("src/WinPDFMerge.Helpers.ps1", "function Get-PdfRunSummary"), ("src/WinPDFMerge.Helpers.ps1", "function Write-NativeProcessLog")], "Measured named stages, versions/input pages/native stdout+stderr/failure timing/sizes/outcome with honest unknown states, no estimated percentage.")
reviewed("timeout-and-cancel", [("docs/USAGE.md", "15-minute execution limit"), ("docs/USAGE.md", "version probes have five seconds"), ("docs/USAGE.md", "Abrupt host/window/machine")], [("src/WinPDFMerge.Helpers.ps1", "[int]$TimeoutMilliseconds = 900000"), ("src/WinPDFMerge.Helpers.ps1", "[int]$TimeoutMilliseconds = 5000"), ("WinPDFMerge.ps1", "$cancellationToken.IsCancellationRequested")], "Per-job/probe bounds, closed input, controlled event-dependent Ctrl+C 1/2, abrupt death no code/cleanup guarantee.")
reviewed("owned-staging-cleanup", [("docs/USAGE.md", "The ownership marker `owner.json`"), ("docs/TROUBLESHOOTING.md", "all merge runs have stopped"), ("docs/TROUBLESHOOTING.md", "Do not sweep by filename prefix")], [("src/WinPDFMerge.Helpers.ps1", "function New-PdfStaging"), ("src/WinPDFMerge.Helpers.ps1", "function Remove-PdfStaging")], "One run's known paths and retained ownership marker only; stopped-run manual inspection; no active/foreign prefix deletion.")
reviewed("troubleshooting-recipes", [("docs/TROUBLESHOOTING.md", "Close programs holding source or owned staging files"), ("docs/TROUBLESHOOTING.md", "Do not delete an existing result to retry")], [("src/WinPDFMerge.Helpers.ps1", "Two-argument File.Move never replaces"), ("WinPDFMerge.ps1", "Source preflight failed:")], "Actionable install/preflight/input/email/log/lock/collision/cancel recovery tied to exact stage/path; no unsafe output deletion workaround.")
reviewed("t19-preservation", [("README.md", "Neither output guarantees"), ("docs/PDF_LIMITATIONS.md", "second full field name became `1.shared_text`"), ("docs/PDF_LIMITATIONS.md", "not a representative compression ratio")], [("src/WinPDFMerge.Helpers.ps1", "'cat', 'output', $stagedOutput, 'compress', 'dont_ask'"), ("src/WinPDFMerge.Helpers.ps1", "'-sDEVICE=pdfwrite'")], "Master nonrasterizing assembly versus lossy email; observed field/tag/attachment/rotation limits and synthetic weighting qualified; originals retained.")
reviewed("privacy-and-reporting", [("README.md", "unencrypted; sanitize a copy"), ("docs/TROUBLESHOOTING.md", "minimal **synthetic** reproduction"), ("SECURITY.md", "ordinary directory/filesystem access controls")], [("WinPDFMerge.ps1", "Diagnostics are local and may contain sensitive"), ("src/WinPDFMerge.Helpers.ps1", "function Write-NativeProcessLog")], "Local unredacted confidential log and staging metadata; sanitize both streams/copies and do not attach private PDFs; no application ACL/encryption promise.")

# File and local-anchor checks are actual mechanical review support. External
# links are inventoried only; their live/vendor facts are audited separately.
local_links: list[dict] = []
external_links: list[dict] = []
for path in PUBLIC:
    if not path.endswith(".md"):
        continue
    for match in re.finditer(r"(?<!!)\[[^\]\n]+\]\(([^)\s]+)\)", TEXT[path]):
        target = match.group(1)
        line = TEXT[path].count("\n", 0, match.start()) + 1
        if re.match(r"[a-zA-Z][a-zA-Z0-9+.-]*:", target):
            external_links.append({"source": f"{path}:{line}", "target": target})
            continue
        file_part, _, anchor = target.partition("#")
        resolved = ((ROOT / path).parent / file_part).resolve() if file_part else (ROOT / path).resolve()
        ok = resolved.is_file() and resolved.is_relative_to(ROOT)
        if ok and anchor:
            slugs = []
            for heading in re.findall(r"(?m)^#{1,6}\s+(.+?)\s*$", resolved.read_text(encoding="utf-8-sig")):
                heading = re.sub(r"[`*_]", "", heading).lower()
                heading = "".join(c for c in heading if c.isalnum() or c in " -_")
                slugs.append(heading.replace(" ", "-"))
            ok = anchor in slugs
        local_links.append({"source": f"{path}:{line}", "target": target, "result": "pass" if ok else "fail"})

bindings = []
OUT.mkdir(parents=True, exist_ok=True)
for path in PUBLIC + RUNTIME:
    data = (ROOT / path).read_bytes()
    snapshot = OUT / "sources" / path
    snapshot.parent.mkdir(parents=True, exist_ok=True)
    snapshot.write_bytes(data)
    bindings.append({"path": path, "bytes": len(data), "sha256": sha(data), "retained_snapshot": snapshot.relative_to(ROOT).as_posix()})
    if data.decode("utf-8-sig").replace("\r\n", "\n").replace("\r", "\n") != TEXT[path]:
        raise RuntimeError(f"reviewed source changed during inspection: {path}")

immutable = []
for path in RUNTIME + ["LICENSE", "docs/PDF_LIMITATIONS.md", "docs/EMAIL_PRESETS.md"]:
    expected_blob = git("rev-parse", f"{BASELINE}:{path}")
    actual_blob = git("hash-object", "--", path)
    immutable.append({"path": path, "baseline_blob": expected_blob, "actual_blob": actual_blob, "result": "pass" if expected_blob == actual_blob else "fail"})

ast_path = (ROOT / args.ast_report).resolve()
ast = json.loads(ast_path.read_text(encoding="utf-8"))
report = {"schema": "t20.user-docs.final-review.v1", "utc": dt.datetime.now(dt.timezone.utc).isoformat(), "case": "AC046", "milestone": "M3 doc/behavior cross-check", "review_scope": "independent semantic public instructions versus actual source plus local-link/hash mechanical support; no application or native engine invoked; no physical Explorer/package/publication claim", "baseline_commit": BASELINE, "current_head": git("rev-parse", "HEAD"), "dirty_paths": git("status", "--porcelain=v1").splitlines(), "python": sys.version.split()[0], "reviewer_script_sha256": sha(Path(__file__).read_bytes()), "checks": checks, "check_count": len(checks), "bindings": bindings, "immutable_checks": immutable, "local_links": local_links, "external_links_inventoried_not_checked": external_links, "ast_parse_support": {"report": ast_path.relative_to(ROOT).as_posix(), "sha256": sha(ast_path.read_bytes()), "result": ast["Result"], "shell": ast["Shell"], "block_count": ast["BlockCount"], "command_count": ast["CommandCount"], "scope": "PowerShell 7.6.5 reader host syntax support only, not pinned supported-shell application compatibility evidence"}, "producer_preparation": ["The first reviewer producer invocation failed before report creation because a source reference used the incorrectly transcribed strict-smaller if-statement. Only the ignored reviewer reference literal was corrected to the existing actual source; no runtime/public-doc change. Current producer succeeds."], "blockers": [], "limitations": ["Documentation/source review does not execute install/drag-drop examples against a real desktop or packaged ZIP; those remain later gates.", "Live vendor/dependency/license/security statements are reviewed separately by the policy auditor.", "C1 binding must prove these reviewed exact public source hashes match the clean implementation commit before recording acceptance."], "result": "pass"}
if any(x["result"] != "pass" for x in local_links + immutable):
    report["result"] = "fail"
    report["blockers"] = [x for x in local_links + immutable if x["result"] != "pass"]
if ast["Result"] != "pass":
    report["result"] = "fail"
    report["blockers"].append({"kind": "AST parse support", "result": ast["Result"]})
guard_checks = []
if args.expected_head:
    guard_checks.append({"kind": "expected current HEAD", "expected": args.expected_head, "actual": report["current_head"], "result": "pass" if report["current_head"] == args.expected_head else "fail"})
if args.require_clean:
    guard_checks.append({"kind": "clean source checkout", "dirty_paths": report["dirty_paths"], "result": "pass" if not report["dirty_paths"] else "fail"})
for binding in ast.get("SourceBindings", []):
    actual_sha = sha((ROOT / binding["Path"]).read_bytes())
    guard_checks.append({"kind": "AST reviewed public source hash", "path": binding["Path"], "expected": binding["SHA256"], "actual": actual_sha, "result": "pass" if binding["SHA256"] == actual_sha else "fail"})
report["guard_checks"] = guard_checks
for guard in guard_checks:
    if guard["result"] != "pass":
        report["result"] = "fail"
        report["blockers"].append(guard)
destination = OUT / "final-review.json"
destination.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
print(json.dumps({"result": report["result"], "semantic_checks": len(checks), "local_links": len(local_links), "immutable_checks": len(immutable), "report": destination.relative_to(ROOT).as_posix(), "report_sha256": sha(destination.read_bytes()), "blockers": report["blockers"]}))
sys.exit(0 if report["result"] == "pass" else 1)
