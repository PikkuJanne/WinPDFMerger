# Compatibility and test environments

T02 measured the following environment on 2026-10-07 at
`4ad96bfe67ffa86753ace9a728dbad21192bb556`. Inventory and isolated primitive
observations are not application compatibility certification. Evidence:
`evidence/T02-baseline.md`, `T02-environment.json`, `T02-probe-ps7.json`.

| Environment | Release scope | Version/build | Evidence | Status |
|---|---|---|---|---|
| Windows 11 x64 standard-user desktop | Required | Pro 10.0.26300/build 26300, DisplayVersion 26H2; non-administrator observation token; NTFS fixed drive | T02 baseline/environment | INVENTORIED; application/Explorer acceptance NOT TESTED; OS support channel unestablished |
| Windows PowerShell 5.1 on reference desktop | Required | 5.1.26100.9444 Desktop x64; test Process RemoteSigned; ordinary Restricted/all scopes Undefined | T09 completion/C3 reports + prior task evidence | T09/M1 narrow217cases PASS with actualPDFtk/GS; full interruption/fidelity/Explorer/release NOT TESTED |
| Supported PowerShell 7 x64 on reference desktop | Required | Supported portable7.6.6 Core x64; Microsoft-signed host; Process RemoteSigned | T09 PS7 acquisition + completion/C3 reports | T09/M1 narrow217cases PASS with actualPDFtk/GS; full application/release compatibility NOT VALIDATED |
| PDFtk native Windows build | Required | Vendor Server2.02 unsigned x86; approved exact external cache | T03 acquisition + T09 completion/C3 | Tested punctuation/Latin paths andCJKinstall PASS; CJKoperands safeFAIL;258PASS/260preflight; bounded encrypted/lock/ACL/collision PASS; no fidelity/release claim |
| Ghostscript native Windows build | Required for email support | Verified10.08.0 x64 console/interpreter unsigned; signed installer; explicit external cache | T09 GS acquisition + completion/C3 | Tested punctuation/Latin/CJK paths,258input, boundedfailure/entry PASS; PDFSTOPONERROR nativeerror regression PASS; full email validation/fidelity/release NOT TESTED |
| GitHub Windows runner | Required CI evidence; not desktop certification | UNKNOWN | None | NOT TESTED |
| Windows 10 | Optional/excludable | UNKNOWN | None | NOT TESTED |
| Live UNC network share | Optional/excludable | UNKNOWN | None | NOT TESTED |
| ARM/32-bit host | Optional/excludable | UNKNOWN | None | NOT TESTED |

Record tool version, architecture, execution-policy context, input/output filesystem, and test commit. Redact private host/user details. README and release notes must not claim broader validation than this matrix supports. A scoped exclusion can satisfy honesty requirements, but it is not a passing compatibility test.

No policy/dependency installation change was made in T02. Pester 3.4.0 is
available but unpinned/unexecuted; PSScriptAnalyzer was not discovered. Windows
PowerShell's successful read-only inline inventory/parser command is not its
blocked script probe passing. Later release gates require a supported recorded
PowerShell 7 update and actual native/manual evidence. No optional case was
excluded during this inventory.

T03 authorized exact Pester6.2.0 and official PDFtk dev-cache acquisition plus
process-only RemoteSigned tests; all T03 gates passed at clean C2
`fca1e20c0240995d8b888fee25c2f24ddf0da418`. No user/machine policy, PATH,
system install, security tool or MOTW change. Pester6.2.0 package/signatures
verified; PSScriptAnalyzer remains absent/unrun. These small harness/fixture
passes are not application or release compatibility certification. T03 receipts
supersede earlier availability/policy constraints only for their stated scope.

T04 extends evidence to narrow real entry/master merges at clean C1
`2f15e8291459940473b4e69c49a27f73fb7eb690`, plus source-provider/array/hidden/nested/
bracket tests. Current non-elevated token and normal PS5.1 policy were rechecked;
no user/machine policy or parent environment changes. Native cases use short ASCII
script/output paths without spaces, child-only dependency environment, GS excluded.
Other punctuation/Unicode source tests are helper-level; no later native quoting,
email, fidelity, output publication or desktop/release pass is inferred.

T05 extends evidence at clean C1 `769817ba9aa8923a3b0001252637865c9a25f7a0`
to actual terminal cmd/BAT path delivery/exit/pause with controlled PS5.1 receivers
and narrow real BAT/application/PDFtk smoke. Each PS5.1/PS7 driver passes24+2;
all child applications are PS5.1. Controlled metacharacter/status cases do not
certify PDF engine path support or real email partial success. Test-child-only
PSModulePath normalization isolates driver inheritance; native dependency overrides
exclude GS. Ordinary PS5.1 policy/token were rechecked without changes. Explorer,
current supported PS7 application build, native paths/email/fidelity/publication
and release gates remain pending. See `evidence/T05-completion.md` for exact scope.

T06 clean C1 `c3c839a264ea35608d5ca6116c595bab4f3050d0` passes Unit57 and
SourceDiscovery4 in EACH actual PS5.1/PS7 shell. Comparator magnitude/leading-zero/
text/case/path ties and en-US/tr-TR/reversed-input consistency are unit evidence.
Real entry/PDFtk multi-input log proves1,2,10,legitimate-prefix order; master page
counts are2,2,5 per shell and sources remain unchanged. Independent final PDF page
order remains downstream. Same narrow native paths/GS exclusion/environment and
current-supported-PS7/desktop/release limitations apply; see T06 completion.

T07 clean C1 `08733007b5dc1ffcbefe225e9cd97fe91b5a9f84` passes Unit99,
DependencyEntry9 and SourceDiscovery4 in EACH actual shell. Strict executable/x86/
numeric GS/version parsing is42new unit cases. Dependency tier is5actual entry
faults,2real-PDFtk2page smokes and2direct controlled probe cases; controlled GS
version failure alone retains master/returns2. Exact timeout PID exit and caller
GS_OPTIONS states are observed. PS5.1 requested empty normalizes to absent; PS7
exposes empty. Test-only process environment setup restores actual state in finally;
application overrides remain child-only. Five real master merges per shell are
inspected; numeric/x86 GS lookup is not native GS/x86-host support. Fixed version
probe bounds do not certify general T08 streams/descendants/cancellation. Native
GS/current-supported-PS7/finalpageorder/fidelity/desktop/release remain unrun.
Exact commands/results/limits and report/build-receipt hashes: T07 completion/C1.

T08 clean C1 `2726c73ad2ff8660540c964789910a9da1f8b3a8` passes Unit99,
NativeRunner36, DependencyEntry9 and SourceDiscovery4 in each actual shell.
Controlled compiled Windows echo verifies exact empty/quote/backslash/control/
Unicode argument vectors; fair capped dual capture/nonzero/launch/inherited-pipe
faults, exact immediate-child timeout/token cancellation, unrelated same-image
survival, child-only GS_OPTIONS and UTF8-noBOM logs pass. These do not certify
native PDF engine paths/encoding, descendant termination or full application
interruption cleanup. Conversion routing/real paths/prompts/command-length guards
remain T09; descendant/run cleanup T15. Five real PDFtk master smokes per shell
retain 2,2 and 2,2,5 page totals and source snapshots. Actual PS7 remains7.6.5,
current-supported update/native GS/final page order/fidelity/desktop/release
unrun. Exact commands/results/limits and report/build hashes: T08 completion/C1.

T09 clean implementation C1 `139ecdcc9f31bc8c55bd0e26e63aaa43e2acdbfc` passes
203 tests per actual PS5.1/PS7 shell (406 total), all failure/skip/not_run zero.
Native PDFtk2.02 source/input/output/install/actual-entry punctuation+Latin paths,
readable CJK operand failures/CJK install success,258 input/260 preflight guard,
password-required, exclusive lock, owned-file ReadData ACL denial with restored
SDDL/source snapshots and final collision refusal pass. Controlled job wiring
and M1 regression tiers are separate evidence. Actual GS remains absent/unrun;
new exact external-cache acquisition authorization is pending. T09/AC019/20 remain
blocked/incomplete as whole required gates; AC021 passes. No supported-current
PS7 or release/Explorer/fidelity claim. See `evidence/T09-evidence.md` and C1 reports.

T09 followup supersedes historical dependency availability/PS7 gaps above. Clean C3
`894bbb4862105ae497bc9801decc46575fedd87a` passes217Pester tests per actual shell/434total,
Python10tests/exact regeneration/PDFium oracle and scoped analyzer0errors/20warnings/
4information. Required AC019/20/21 pass; M1 complete. Owner authorized all needed
test dependencies, retained outside repository/system. Acquisition receipts and
C3 manifest/review disclose actual signatures/hashes, static warnings and historical
GS exit0/password-error wrongpage failure fixed with PDFSTOPONERROR. Full release/
OS support channel/Explorer/fidelity/structural/email validation remain later gates.

T10 clean C1 `962f05ec1ea9f8edfbeb8e5fc755dacac24156ed` passes202 focused tests per actual
PS5.1.26100.9444/supportedPS7.6.6 shell,404total; AC022/23/24 pass. Actual local
NTFS standard-user copied application paths prove owned WriteData/AddFile denial,
default refusal/explicit writable recovery/original SDDL restoration, physical
same/case/available8.3 aliases, four leaf/ancestor junction refusals and two
concurrently live PDFtk2.02/GS10.08.0 applications with distinct shared identities.
Sources/foreign finals and owned probe/stage cleanup are checked. Unit naming/
metadata mocks remain separate. Static0errors/29warnings/5info retained/reviewed
nonblocking. Known caches reverified without acquisition/system changes. Earlier
full NativeRunner/Launcher/Python fixture suites were not repeated; T11–T15,
OS support-channel/UNC/Explorer/fidelity/CI/full-release gates remain pending.
See T10 completion/C1 reports/review/live sync and cache audit.
