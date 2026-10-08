# T16 completion — small public parameter interface

Implementation under test: `26ac1b73e3733a23099de53d944e00e4ee412982` (C1),
branch `codex/v1.0.0-readiness`, unchanged origin
`https://github.com/PikkuJanne/WinPDFMerger.git`. Started clean and synchronized
at `fe8210d9ca604435de8f7261e7afdd9774291951`; owner merged PR15 and live main
was `c1fdb3bf99b5b5e2631f08634b3b5b918635e220`. No tags/releases observed.
[Draft PR16](__PR_URL__) contains implementation and this records closure.

`EmailPreset` accepts only screen/ebook, case-insensitively, defaults to screen,
and selects literal allowlisted Ghostscript flags. Output stays beside entry
unless an existing writable literal OutputFolder is selected. Legacy positional
and named SourceFolder invocation stays available. Missing source prints usage
before helper import, with closed standard input; unknown/extra arguments and
invalid presets bind before application effects. SkipEmail never discovers,
probes or launches Ghostscript and explains an explicitly supplied valid preset
in console/log. Master validation, owned staging/jobs, fixed safety arguments,
strictly smaller email publication and source/final preservation remain.

AC038 integration and AC039 unit pass. Actual Windows PowerShell 5.1.26100.9444
Desktop x64 and pinned PowerShell 7.6.6 Core x64 each pass 536 cases in 14 relevant
tiers: **1,072 total**, 28 complete reports, all failure/block/container/skip/
not-run counts zero. Pester6.2.0; PDFtk2.02; Ghostscript10.08.0;
development Python3.12.14/pypdfium2 5.13.0/PDFium153.0.7999.0.

| Tier | Each shell |
| --- | ---: |
| Unit | 335 |
| Parameters | 31 |
| ParametersNative | 9 |
| EmailOutcome | 11 |
| MasterValidation | 7 |
| Staging | 9 |
| InputPreflight | 22 |
| Destination | 15 |
| ToolInvocation | 12 |
| GhostscriptPaths | 13 |
| Launcher | 24 |
| LauncherNative | 2 |
| FaultIO | 32 |
| FaultRecovery | 14 |

The clean drivers execute the explicit approved shell with `-NoProfile
-ExecutionPolicy RemoteSigned -File tests/.work/Run-T16Checkpoint.ps1
-ShellLabel ps51|ps7 -ExpectedCommit 26ac1b73e3733a23099de53d944e00e4ee412982
-ExpectedCountsPath tests/.work/T16-expected-counts.json -Checkpoint C1`.
The exact selected executables, individual tier argument vectors, source/cache
hashes, timestamps, bounds and separate tier stdout/stderr are in collector/
runs/aggregate/manifest receipts. The two outer driver captures combine streams;
there is no fabricated separate outer stderr. Both drivers exit0. Only each
child's inherited PSModulePath is removed. All tiers verify clean C1 before/after.

Native parameter cases run real engine copies, record actual argument vectors,
complete job/inspection receipts and fresh PDFtk/PDFium published-file reads.
They cover positional/named/default output, screen/ebook/mixed case, named
destination, usage and both skip cases. Actual cmd/BAT delivery starts PS5.1
under either parent test shell. Skip cases disclose discovery/probe/launch
throwing controls while real Ghostscript remains available. The deterministic
single raster master is 2,882,118 bytes; screen and ebook derivatives are smaller
with matching visible ID, count and geometry. This is structural integration
evidence, not preset-quality acceptance. Unit cases use controlled receipts;
they do not count as native engine acceptance.

Independent runtime/source reviews find no blocking findings, with authorship
limits disclosed. Scoped actual PSA1.25.0 over seven changed PowerShell files
reports 0 errors, 71 warnings and 34 informational findings per shell, reviewed
nonblocking; no lint-clean/fullT22 claim. Independent native audit passes 2,273
checks, 18 case audits and 28 fresh retained PDFtk/PDFium reads. Independent
archive review passes __ARCHIVE_CHECKS__ checks of the 335-file collector plan.
Actual root write exactly matches that reviewed manifest/results/file plan.
Raw and sanitized public hashes remain distinct; support indexes bind exact
retained originals. File-specific literal diagnostic whitespace waivers preserve
captured bytes. No original document or private PDF is uploaded.

Dirty preparation is retained separately and excluded from 1,072 clean passes:
initial Parameters runs were24pass/7fail each because the test wrapper did not
propagate LASTEXITCODE; corrected wrappers pass31each. Initial native PS5.1 was
1pass/8fail from JSON-array receipt enumeration while real jobs succeeded;
PS7passed9, corrected checks pass9both. Initial native observations contain only
the completed missing-input case; raw application copies/receipts bind the
other runs. Initial unit source was not separately frozen; original invocation
hashes and executed copies exist. The initial native source is separately frozen.
Dirty root335-per-shell runs, focused finals and dirty scoped analyzer findings
remain separate. The first independent native auditor's16 CRLF log-comparison
failures were corrected in its ignored byte reader, with exact auditor/wrapper
snapshots and both28-read attempts retained. Only the final28 reads certify that
audit. First collector check failed an overly literal sentinel-label assertion;
the corrected semantic check passed, preserving failed source/captures without
changing application/tests or adding acceptance counts. Absent historical
producer outputs/source snapshots are explicitly disclosed, never reconstructed.

Ordinary read-only PS5.1 inventory at C1 reports standard user, x64, local NTFS,
Windows10.0.26300.0, Restricted effective policy/all scopes Undefined; support
channel unknown. Previously authorized test-child RemoteSigned is separate.
Approved caches were rehashed and reused:14selected files/5dependencies plus
separate Python/PDFium pins, no acquisition/admin/persistent PATH/policy/security
or parent environment changes. Cached original acquisition/full-audit receipts
remain linked; inventory is not an application/manual pass.

Root C1 normal push and fresh read-only `handoff.py sync --repo .` verify clean
local/live equality. Records-only C2 binds C1, retains historical matrix and
attribute prefixes, updates only T16/AC038/AC039 and selects T17. Root actual
public/raw/privacy/index checks and independent C2 semantic review precede the
final intended-path commit/push/live check. C2's own SHA/equality is reported in
the session after commit, avoiding a committed self-reference. `check-plan` is
structural validation, not additional Windows/manual/release evidence.

Physical Explorer drag/drop remains a manual/release gate distinct from actual
BAT argument delivery. T17 AC040 size reporting and AC041 small-print/scanned/
mixed fidelity observations remain pending. Unchanged NativeRunner,
DependencyEntry, SourceDiscovery, PdftkPaths and NativeFixture tiers retain
prior evidence, with no new run claim. Broader OS/UNC/CI/security/package/
publication gates remain open. No universal preservation/security, signatures,
PDF/A, malware removal, crash cleanup or archival guarantee is added. M3 stays
in progress. No tag/release/publication occurs in T16.
