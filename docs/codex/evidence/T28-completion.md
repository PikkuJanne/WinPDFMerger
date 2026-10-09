# T28 clean allowlisted packaging and checksum completion

2026-10-09. AC065/AC066 pass at clean tested source C1b
`8917938820f60e499e2c20caa9cb03171678be72`. Initial implementation C1 was
`fd0acc77907134ed424e11b74edc227b6b1480c8`; C1b only corrects capture labels
and adds their validator/regressions plus the interruption record. Runtime,
public docs, builder, allowlist and PowerShell tests are identical across them.
Neither commit is the later accepted merged release source R.

## Implemented contract

Build-Release.ps1 accepts a full lowercase 40-character SourceCommit equal to
clean HEAD and a new OutputDirectory outside the repository beneath an existing
ordinary parent. Optional RepositoryRoot supports isolated test checkouts; the
tracked builder must be invoked from that root. Direct checked Git subprocesses
read exact regular-file blob bytes; no shell evaluation, filters, wildcard
checkout copying, acquisition or network call is used.

The reviewed release-files.json contains 15 source files: both familiar entry
points, src/WinPDFMerge.Helpers.ps1, VERSION, README, MIT LICENSE, SECURITY,
CHANGELOG and seven public instruction/release-note documents. Unused icons/
posters, tests, developer handoff/tools, PDFs/logs, vendors and unrelated assets
are absent. Missing/untracked sources, changed/staged/untracked/other ignored
content, hidden-change index flags, links/reparse ancestors, unsafe names,
case-insensitive duplicates, nonnumeric schema versions and version/contract
disagreements fail. Only ignored tests/.work caches may coexist; they are never
read into the archive. Both excluded development-doc links now target immutable
GitHub source; packaged relative Markdown links all resolve.

One WinPDFMerger-v1.0.0 root contains those 15 files and generated BUILD_INFO.json.
VERSION feeds names and metadata. BUILD_INFO binds full source SHA, actual host/
Git/CLR/OS identity, committed builder/allowlist/contract digests and complete
per-file inventory; it excludes its own hash. ZIP entries have ordinal order,
fixed 2000 timestamp/zero attributes and NoCompression. ZIP readback verifies exact
inventory and bytes before an owned sibling staging directory is moved without
overwrite to the new destination. Existing output directories are refused;
cleanup deletes only this run's known files and empty stage, never recursively.
Source identity/cleanliness and ordinary paths are rechecked before publication.
This is local asset publication, not a GitHub release.

SHA256SUMS.txt is the exact one-line ZIP digest/name. The return value separately
records both complete asset hashes. Integrity checksums are not signatures or
publisher identity. Unsigned application status and all PDF/dependency limits
remain disclosed; no source/runtime/native behavior or default is changed.

## Actual clean execution

The tracked capture-T28.py ran with approved workspace Python 3.12.14 after
rehashing348 previously approved cache payloads. Actual x64 hosts were Windows
PowerShell 5.1.26100.9444 Desktop and pinned PowerShell 7.6.6 Core, Pester 6.2.0,
analyzer 1.25.0 and Git 2.56.0.windows.2. Windows registry observed Professional
26H2/full26300.9457; OSVersion10.0.26300.0, both nonadministrator tokens. Null
channel registry fields do not prove enrollment. Child Process RemoteSigned and
child-only module-path cleanup made no persistent policy/PATH/module/security
change. No dependency installation, elevation or acquisition occurred.

| Clean C1b tier | PS5.1 | Pinned PS7.6.6 | Scope |
| --- | ---: | ---: | --- |
| Package | 66 | 66 | Actual synthetic Git/build children, independent ZIP/blob/refusal checks |
| Unit | 545 | 545 | Unit/controlled helper/harness checks |
| Version | 21 | 21 | Version/static and actual nonmerging preflight children |
| PublicDocs | 22 | 22 | Public contracts and isolated binding/helper checks |
| Static | 9 | 9 | Actual selected-analyzer checker regressions on synthetic source |
| Total | 663 | 663 | 1326passes; every bad/discovery/skip count0 |

All guarded source snapshots and outer clean/driver guards pass. Selected-file
static checks cover7 changed PowerShell files/41 rules per host: parser errors,
selected findings/suppressions, analyzer not_run and source/checkpoint failures0.
Vendor default advisories remain visible:0 errors/18 warnings/0 information each.
Three AST-only capture-label stdlib regressions additionally pass. The27 handoff
helper tests separately give26 pass/1 skip for unavailable symlink creation; this
is not an application/native pass and no elevation was requested.

Actual commands (exact sanitized argument/exit/hash ledgers retained):

```text
<approved-python> -B docs/codex/evidence/T28-reports/scripts/capture-T28.py 8917938820f60e499e2c20caa9cb03171678be72
<host> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File <tracked environment probe>
<host> ... -File tools/test/Invoke-Tests.ps1 -Tier <Package|Unit|Version|PublicDocs|Static> -PesterModulePath <approved>
# Static tier additionally supplies the explicit approved analyzer path.
<host> ... -Command <guarded Export-CiTestReport with exact source/tier/shell>
<host> ... -Command <Invoke-StaticChecks.ps1 with seven explicit source paths>
<host> ... -Command <Build-Release.ps1 -SourceCommit <C1b> -OutputDirectory <new external path>; JSON result>
<approved-python> -B -m unittest discover -s tests/package -p test_*.py -v
<approved-python> -B -m unittest discover -s tools/codex/tests -v
<approved-python> -B tools/codex/handoff.py check-plan --repo .
```

There are31 actual successful capture invocations,10 guarded original/exported
JSON+NUnit pairs and four actual repository builds: first/repeat per host.
Each ZIP has16 files; both same-environment pairs match byte-for-byte. PS5.1 CLR
4.0.30319.42000 encodes NoCompression as ZIP method 8; PS7 CLR10.0.12 uses method 0.
Recorded host metadata also differs. Cross-shell ZIP identity is not promised;
both carry identical exact15 Git payload blobs and valid provenance. Neither
package is executed by this task. Exact hashes are in T28-results.json and
T28-reports/builds; PS5.1 ZIP 191010 bytes, PS7 ZIP 189617 bytes, each checksum 90 bytes.

## Independent review and retained failures

Independent ZIP/source audits pass260 checks/0 issues each at clean C1b, covering
safe metadata, inventory, all blob bytes, BUILD_INFO, both hashes, local-doc
closure and repeated bytes. Original/exported receipt/stream/cache/source/static/
build audit passes3647 checks/0 issues. These are review checks, not additional
application/native cases. T28-review.md and unchanged manifested auditors/results
retain exact scope and commands. The public projection excludes application ZIP
assets and private original diagnostics; its manifest preserves exact bytes.
Independent public projection/manifest audit passes 1335 checks/0 issues across
55 manifested payloads; T28-archive-review.json binds the exact manifest hash.

Dirty failed preparations and stable66/66 preparations remain separate in
T28-preparation.*. Independent preparation audit passes122 checks/0 issues.
C1's outer capture failed after1326 passing leaves because Static/static filenames
collided on Windows; it created no repository package. It remains incomplete/
failed, with an independent2958 check/0 issue audit and interruption note. C1b's
casefold preflight and regression cover that failure. Its full successful rerun
alone supplies the clean acceptance totals above; no failed receipt is rewritten.

## Synchronization and handoff

C1 normal push/clean live equality was observed at14:24:06UTC; C1b at14:29:46UTC.
Both C1 push37943890288/PR37943900641 and C1b push37944594925/PR37944606821
finished4/4 jobs successfully. These are platform metadata observations, not
new hosted application test-count or exact-package claims. Fresh14:34:33UTC
platform inspection shows PR26 draft/open/unmerged, main e245114 and no tags or
releases. Reviewer independently confirmed clean/live C1b after ZIP audits.

Final records C2 stores completed task/case/status/continuation and retained
evidence. Its intended-file diff, plan/manifest/staged-byte checks, normal push
and own fresh clean/live equality must be verified after creating it and are
reported in the session, avoiding a self-referential future SHA claim.
The final byte check caught CRLF normalization of the standalone staged-audit
JSON. C2 adds a narrow -text attribute for that report and stages its unchanged
original bytes; no audit receipt is rewritten to match normalization. This
record-preservation attribute is the only C2 path outside docs/codex.

Next task is exactly T29: extract and operate the exact candidate ZIP on actual
Windows in both required shells with real PDFtk/GS, independent PDF inspection
and unchanged sources. T30 must refresh preparation/Unreleased wording from
actual evidence before source acceptance. AC058 remains excluded/unperformed,
never pass; Windows10/liveUNC/ARM/32-bit exclusions and all prior limitations
remain. Final accepted-R operation, publication, independent published download
and synchronized closure are later required gates. The project is not complete.
