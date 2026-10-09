# T27 preparation checkpoint

Prepared 2026-10-09 from clean/live-synchronized readiness source
`5a5c6e08699175f5fcc3b945d9b2aa233f5375bb`. Both origin routes target
PikkuJanne/WinPDFMerger; main is `e2451141217efdd00a1d49d72a04df054872dffc`.
PR26 is draft/open/unmerged; live tag/release lists are empty. The first release
list attempt requested an unsupported CLI JSON field `url`; the supported-field
retry succeeded and returned no releases. No GitHub mutation occurred in those reads.

VERSION contains 1.0.0. The runtime reads it as data for the startup/usage banner
and established run log; help names the source. Package naming/build-info
contract and public note titles are validated against it. The builder remains
T28 work. Missing/invalid metadata fails before merge work. Copied test installs
carry VERSION; source receipts bind its bytes and the release contracts.

Dirty preparation command used hash-approved portable PS7.6.6 with Pester6.2.0:
`<pinned-host> -NoProfile -NonInteractive -ExecutionPolicy RemoteSigned -File
tools/test/Invoke-Tests.ps1 -Tier Version -PesterModulePath <approved-manifest>`.
The first run had20pass/1fail/0skip/0not_run/0inconclusive: the test's expected
asset array lacked parentheses around string concatenation, producing one
expected item. Fixed that assertion and reran21pass/0fail with all bad counts0.
These dirty executions are preparation only, not clean implementation acceptance.
Ignored captures are under tests/.work/T27-preparation; final evidence will
bind the clean implementation commit and independently review original receipts.

Independent preliminary review caught public layout/recovery instructions that
omitted mandatory VERSION. README/usage/troubleshooting and the existing
PublicDocs regression now require it. A preflight regression also now runs in
an isolated copied layout to avoid unrelated ignored-work I/O. Final static
pre-execution review reports no unresolved finding.

Notes retain actual source/native versus controlled/CI scope, exact reference
versions, dependency/PDF/path/privacy limits, unsigned status, AC058
excluded/unperformed and Windows10/liveUNC/ARM/32-bit-host exclusions.
No package/download/publication pass is claimed.

T27 remains in_progress pending clean dual-shell targeted tests, selected-file
parser/analyzer, independent AC063/AC064 review and synchronized completion
records. No dependency acquisition, elevation, persistent policy/PATH/security
change, package, tag or release occurred. Defaults remain within the product contract.
