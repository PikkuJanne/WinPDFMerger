"""Owned local preparation only: copy exact source/approved original fixtures."""
from pathlib import Path
import datetime, hashlib, json, os, shutil, subprocess, sys, uuid

repo = Path.cwd().resolve()
candidate = "e2451141217efdd00a1d49d72a04df054872dffc"
kit = repo / "tests/.work/T26-desktop" / ("walkthrough-" + uuid.uuid4().hex[:12])
kit.mkdir(parents=True, exist_ok=False)
hashbytes = lambda raw: hashlib.sha256(raw).hexdigest()
hashfile = lambda p: hashbytes(p.read_bytes())
gitblob = lambda label: subprocess.check_output(["git", "show", candidate + ":" + label], cwd=repo)
inventory = json.loads((repo / "tests/.work/T23-environment.json").read_text())
assert inventory["result"] == "pass"
assert hashfile(Path(sys.executable)) == inventory["python_sha256"]

app = kit / "Candidate with spaces"
labels = ["WinPDFMerge.ps1", "WinPDFMerge.bat", "src/WinPDFMerge.Helpers.ps1", "README.md", "LICENSE",
          "docs/USAGE.md", "docs/DEPENDENCIES.md", "docs/EMAIL_PRESETS.md", "docs/PDF_LIMITATIONS.md", "docs/TROUBLESHOOTING.md"]
source_rows = []
for label in labels:
    raw = gitblob(label)
    target = app / label
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(raw)
    assert hashfile(target) == hashbytes(raw)
    source_rows.append({"repository_path": label, "kit_path": target.relative_to(kit).as_posix(),
                        "sha256": hashbytes(raw), "bytes": len(raw), "source_commit": candidate})

catalog = json.loads((repo / "tests/fixtures/corpus.json").read_text())
existing = repo / "tests/.work/corpus-safety/f875eb04f7c343f2bb3cf5a518b17373/corpus"
chosen = [("numbered", "1.pdf", "1.pdf"), ("numbered", "2.pdf", "2.pdf"), ("numbered", "10.pdf", "10.pdf"),
          ("presets", "small-print.pdf", "11-small-print.pdf"), ("presets", "scan.pdf", "12-scan.pdf"),
          ("presets", "mixed.pdf", "13-mixed.pdf"), ("features", "1-feature-A.pdf", "14-rotation-feature-A.pdf")]
mixed = kit / "Sources" / "Mixed numbered scans rotation"
special = kit / "Sources" / "Special [ ] & !"
rejected = kit / "Sources" / "Rejected empty input"
for folder in (mixed, special, rejected): folder.mkdir(parents=True, exist_ok=False)
fixture_rows, ids = [], []
for group, original, final in chosen:
    row = next(r for r in catalog["groups"][group]["fixtures"] if r["file"] == original)
    origin = repo / "tests/fixtures/numbered" / original if group == "numbered" else existing / group / original
    assert hashfile(origin) == row["sha256"]
    for folder in (mixed, special):
        target = folder / final
        shutil.copyfile(origin, target)
        assert hashfile(target) == row["sha256"]
        fixture_rows.append({"kit_path": target.relative_to(kit).as_posix(), "catalog_group": group,
                            "catalog_file": original, "sha256": row["sha256"], "bytes": target.stat().st_size,
                            "page_identifiers": row["page_identifiers"], "page_count": row["page_count"]})
    ids.extend(row["page_identifiers"])
shutil.copyfile(repo / "tests/fixtures/numbered/1.pdf", rejected / "1.pdf")
(rejected / "2-empty.pdf").write_bytes(b"")
for p in sorted(rejected.iterdir()):
    fixture_rows.append({"kit_path": p.relative_to(kit).as_posix(), "sha256": hashfile(p), "bytes": p.stat().st_size,
                         "provenance": "Tracked numbered original" if p.name == "1.pdf" else "T21 safety empty recipe b''"})

for label in ("PS51 explicit output", "PS7 explicit output", "PS51 ebook", "PS7 ebook", "PS51 skip", "PS7 skip"):
    (kit / "CLI outputs" / label).mkdir(parents=True, exist_ok=False)

native = json.loads((repo / "docs/codex/evidence/T25-reports/native-cache-review.json").read_text())["native_selected_files"]
dependency_rows = []
for row in native:
    actual = Path(row["path"].replace("<USERPROFILE>", os.environ["USERPROFILE"]))
    assert hashfile(actual) == row["sha256"] and actual.stat().st_size == row["bytes"]
    dependency_rows.append({"path_template": row["path"], "sha256": row["sha256"], "bytes": row["bytes"]})
pwshrow = next(row for row in inventory["approved_selected_files"] if Path(row["path"]).name == "pwsh.exe")
assert hashfile(Path(pwshrow["path"])) == pwshrow["sha256"]
dependency_rows.append({"path_template": pwshrow["path"].replace(os.environ["USERPROFILE"], "<USERPROFILE>"),
                        "sha256": pwshrow["sha256"], "bytes": Path(pwshrow["path"]).stat().st_size})

start = r'''# Private, process-only setup. No application invocation, install, or policy change.
[CmdletBinding()]
param([switch]$OpenExplorer)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$global:T26Kit = $PSScriptRoot
$global:T26App = Join-Path $T26Kit 'Candidate with spaces'
$global:T26Mixed = Join-Path $T26Kit 'Sources\Mixed numbered scans rotation'
$global:T26Special = Join-Path $T26Kit 'Sources\Special [ ] & !'
$global:T26Rejected = Join-Path $T26Kit 'Sources\Rejected empty input'
$global:T26PS51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$global:T26OriginalPath = $env:Path
$manifest = Get-Content -LiteralPath (Join-Path $T26Kit 'kit-manifest.json') -Raw | ConvertFrom-Json
foreach ($row in @($manifest.source_files) + @($manifest.fixtures)) {
    $path = Join-Path $T26Kit $row.kit_path
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $row.sha256) {
        throw ('Kit bytes changed: ' + $row.kit_path)
    }
}
$dependencies = foreach ($row in $manifest.dependencies) {
    $path = $row.path_template.Replace('<USERPROFILE>', $env:USERPROFILE)
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $row.sha256) {
        throw 'Approved dependency bytes changed; stop and report privately.'
    }
    [pscustomobject]@{ Name=[IO.Path]::GetFileName($path); Path=$path }
}
$pdf = ($dependencies | Where-Object Name -eq 'pdftk.exe').Path
$gs = ($dependencies | Where-Object Name -eq 'gswin64c.exe').Path
$global:T26PS7 = ($dependencies | Where-Object Name -eq 'pwsh.exe').Path
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if ($principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Elevated token detected. Close this window and use a normal standard-user session.'
}
$env:Path = (Split-Path -Parent $pdf) + ';' + (Split-Path -Parent $gs) + ';' + $T26OriginalPath
Set-Location -LiteralPath $T26App
Write-Host 'Exact candidate/fixture and five selected dependency-file hashes checked.'
Write-Host 'Only this process PATH and its children were adjusted. Confirm a standard-user account separately.'
Write-Host 'Candidate: e2451141217efdd00a1d49d72a04df054872dffc; 10 expected mixed pages.'
Write-Host 'Private session variables: $T26Kit $T26App $T26Mixed $T26Special $T26Rejected $T26PS51 $T26PS7'
Write-Host 'Follow WALKTHROUGH.md. Record actual dependency paths from each application log.'
if ($OpenExplorer) {
    Start-Process -FilePath (Join-Path $env:SystemRoot 'explorer.exe') -ArgumentList @('/separate', ('"{0}"' -f $T26Kit)) -WindowStyle Normal
    Write-Host 'Separate Explorer requested. Existing shell reuse/inheritance is unproven; actual drop/run logs decide.'
}
'''
(kit / "Start-Walkthrough.ps1").write_text(start, encoding="utf-8", newline="\r\n")

walkthrough = '''# T26 private walkthrough — all observations not_run

This is a local copied-source preparation kit, not a candidate distribution ZIP.
Source commit: `e2451141217efdd00a1d49d72a04df054872dffc`.
Application, README and selected public docs are exact Git blob copies. No third-party binary is copied.

1. Use a normal Windows 11 x64 standard-user account. Confirm Settings > System > About gives edition/version/full OS build; Settings > Windows Update > Windows Insider Program gives actual enrollment/channel status. Record the displayed facts privately. A build-number lookup alone does not establish enrollment. Do not change settings.
2. In an ordinary PowerShell window, start a disposable Windows PowerShell 5.1 child with the already authorized process-only RemoteSigned flag; Group Policy remains authoritative:

   `& "$env:SystemRoot\\System32\\WindowsPowerShell\\v1.0\\powershell.exe" -NoProfile -ExecutionPolicy RemoteSigned -NoExit -Command ". '<KIT>\\Start-Walkthrough.ps1' -OpenExplorer"`

   Replace `<KIT>` with this kit's absolute path. If enterprise policy blocks this, record the block; do not change policy. Setup checks copied bytes and five approved cached files, adjusts only child-process PATH, and requests a separate visible Explorer. It never invokes the application. Existing Explorer may reuse the shell process: success requires the actual application log to identify the approved PDFtk/GS paths. A missing dependency is a real observation, not a passed run. Close the disposable shell when finished; no machine/user PATH or policy is written.
3. Read `Candidate with spaces\\README.md`. In actual Explorer, drag `Sources\\Mixed numbered scans rotation` onto the exact `Candidate with spaces\\WinPDFMerge.bat`. Observe the numbered list, stages, final result/code and pause. Do not substitute a terminal BAT command. Default output must be beside the copied launchers. Preserve the window until result is recorded.
4. Open the advertised master in a local PDF viewer. Confirm 10 pages and IDs below. Inspect vector small text on page 5, scan marks/text/diagram on page 6, raster/vector pages 7–8, and rotated feature page 10. Open a published smaller email file, if present, and record the same count/order, text/scan readability, actual orientation and any visible defects. Compare with original PDFs. Screen may rotate text upright; forms/tags/attachments/navigation have documented limits. Do not follow links or open attachments. Record viewer/version and zoom used. If no smaller email is published, record the real reason; do not invent one.
5. Repeat actual Explorer drag/drop with `Sources\\Special [ ] & !`, then `Sources\\Rejected empty input`. Special case expects the same 10 pages. Rejected case expects exit 1, an explicit invalid/empty input diagnostic, no advertised/final master/email, preserved sources, and a retained pause.
6. In the prepared PS5.1 child, use the exact public CLI examples adapted to these synthetic paths. Record each actual console/log result/code. These are CLI checks, not Explorer observations:

```powershell
& $T26PS51 -NoProfile -ExecutionPolicy RemoteSigned -File (Join-Path $T26App 'WinPDFMerge.ps1') -SourceFolder $T26Mixed
& $T26PS51 -NoProfile -ExecutionPolicy RemoteSigned -File (Join-Path $T26App 'WinPDFMerge.ps1') -SourceFolder $T26Mixed -OutputFolder (Join-Path $T26Kit 'CLI outputs\\PS51 explicit output')
& $T26PS51 -NoProfile -ExecutionPolicy RemoteSigned -File (Join-Path $T26App 'WinPDFMerge.ps1') -SourceFolder $T26Mixed -EmailPreset ebook -OutputFolder (Join-Path $T26Kit 'CLI outputs\\PS51 ebook')
& $T26PS51 -NoProfile -ExecutionPolicy RemoteSigned -File (Join-Path $T26App 'WinPDFMerge.ps1') -SourceFolder $T26Mixed -SkipEmail -OutputFolder (Join-Path $T26Kit 'CLI outputs\\PS51 skip')
& $T26PS7 -NoProfile -ExecutionPolicy RemoteSigned -File (Join-Path $T26App 'WinPDFMerge.ps1') -SourceFolder $T26Mixed
& $T26PS7 -NoProfile -ExecutionPolicy RemoteSigned -File (Join-Path $T26App 'WinPDFMerge.ps1') -SourceFolder $T26Mixed -OutputFolder (Join-Path $T26Kit 'CLI outputs\\PS7 explicit output')
& $T26PS7 -NoProfile -ExecutionPolicy RemoteSigned -File (Join-Path $T26App 'WinPDFMerge.ps1') -SourceFolder $T26Mixed -EmailPreset ebook -OutputFolder (Join-Path $T26Kit 'CLI outputs\\PS7 ebook')
& $T26PS7 -NoProfile -ExecutionPolicy RemoteSigned -File (Join-Path $T26App 'WinPDFMerge.ps1') -SourceFolder $T26Mixed -SkipEmail -OutputFolder (Join-Path $T26Kit 'CLI outputs\\PS7 skip')
```

`-SkipEmail` must show a complete master-only result/code 0 without GS discovery/use. It satisfies the template's master-only branch; no dependency is removed and no missing-GS observation is implied. Use `$LASTEXITCODE` after each CLI child. Inspect at least the master-only master and ebook output as well. Record actual PS5.1/PS7/PDFtk/GS versions shown in logs; selected pinned PS7 is 7.6.6. Keep the real logs private until sanitized.

7. Rerun the setup hash guard without opening Explorer: `. (Join-Path $T26Kit 'Start-Walkthrough.ps1')`. This compares every source/copy with the original manifest. Record unchanged hashes or exact changed synthetic path. Observe whether only validated final outputs were advertised; any new run logs are expected outside Sources. Record all actual observations in `OBSERVATIONS.md`; it initially contains only blanks/not_run.

Expected mixed and special sequence:

| Page | Visible ID | Original |
|---|---|---|
| 1 | T03-01-P01 | 1.pdf |
| 2 | T03-02-P01 | 2.pdf |
| 3 | T03-02-P02 | 2.pdf |
| 4 | T03-10-P01 | 10.pdf |
| 5 | T03-17-P01 | 11-small-print.pdf |
| 6 | T03-17-P02 | 12-scan.pdf |
| 7 | T03-17-P03 | 13-mixed.pdf |
| 8 | T03-17-P04 | 13-mixed.pdf |
| 9 | T19-A-P1 | 14-rotation-feature-A.pdf |
| 10 | T19-A-P2 | 14-rotation-feature-A.pdf (source /Rotate 90) |

Only provide a sanitized summary/synthetic screenshots; no private identity paths, PDFs or raw unreviewed logs. The later T28/T30 exact-ZIP extraction/distribution gate remains unrun.
'''
(kit / "WALKTHROUGH.md").write_text(walkthrough, encoding="utf-8")
observations = '''# T26 actual observer record — not_run

Observer (public-safe name/role): [not supplied]
Date/time UTC: [not supplied]
Windows edition/version/full build/architecture: [not supplied]
Settings Windows Insider Program enrollment/channel as actually displayed: [not supplied]
Standard-user account status and non-elevated token: [not supplied]
PS5.1 / pinned PS7 / PDFtk / GS versions: [not supplied]
PDF viewer/version and inspection zoom: [not supplied]
Copied source candidate: e2451141217efdd00a1d49d72a04df054872dffc
Kit source/docs/fixture and selected dependency hash check: [actual result not_run]
Distribution ZIP/hash: not applicable; T28/T30 distribution check remains not_run
Extraction/copy path: <OWNED_WORK>\\T26-desktop\\<KIT>\\Candidate with spaces

| Case | Actual observation | Result | Evidence |
|---|---|---|---|
| Standard-user OS/enrollment evidence | [blank] | not_run | [blank] |
| Actual Explorer mixed drop/list/stages/default output/pause | [blank] | not_run | [blank] |
| Visible master 10 IDs/small text/scans/rotation | [blank] | not_run | [blank] |
| Actual screen email outcome and visible inspection | [blank] | not_run | [blank] |
| Actual Explorer special [ ] & ! drop | [blank] | not_run | [blank] |
| Actual Explorer rejected empty input | [blank] | not_run | [blank] |
| PS5.1 four public CLI examples/master-only inspection | [blank] | not_run | [blank] |
| PS7 four public CLI examples/master-only inspection | [blank] | not_run | [blank] |
| Source hashes unchanged; only validated finals advertised | [blank] | not_run | [blank] |

Limitations/exclusions: [actual statement not supplied]
Observer acceptance or specific failure: [not supplied]
'''
(kit / "OBSERVATIONS.md").write_text(observations, encoding="utf-8")

manifest = {"task": "T26", "purpose": "private copied-source human desktop preparation, not executed/manual/distribution evidence",
            "prepared_at_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(), "source_commit": candidate,
            "application_copy_hashes_match_source": True, "fixtures_match_tracked_catalog_hashes": True,
            "source_files": source_rows, "fixtures": fixture_rows, "dependencies": dependency_rows,
            "fixture_provenance": catalog["provenance"], "development_python_version": sys.version.split()[0],
            "development_python_sha256": hashfile(Path(sys.executable)), "mixed_ordered_names": [row[2] for row in chosen],
            "mixed_expected_page_count": len(ids), "mixed_expected_page_identifiers": ids,
            "special_expected_page_identifiers": ids, "rejected_expected_exit": 1,
            "all_manual_results": "not_run", "distribution_zip": None,
            "new_dependency_acquisition": False, "application_executed": False, "Explorer_executed": False,
            "new_install_or_persistent_changes": False}
(kit / "kit-manifest.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
for row in source_rows + fixture_rows:
    assert hashfile(kit / row["kit_path"]) == row["sha256"]
print(json.dumps({"kit_path": str(kit), "manifest_sha256": hashfile(kit / "kit-manifest.json"),
                  "application_sources": len(source_rows), "fixture_copies": len(fixture_rows),
                  "mixed_pages": len(ids), "manual_results": "not_run", "distribution_zip": None}))
