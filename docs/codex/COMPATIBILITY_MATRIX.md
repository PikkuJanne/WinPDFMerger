# Compatibility and test environments

T02 measured the following environment on 2026-10-07 at
`4ad96bfe67ffa86753ace9a728dbad21192bb556`. Inventory and isolated primitive
observations are not application compatibility certification. Evidence:
`evidence/T02-baseline.md`, `T02-environment.json`, `T02-probe-ps7.json`.

| Environment | Release scope | Version/build | Evidence | Status |
|---|---|---|---|---|
| Windows 11 x64 standard-user desktop | Required | Pro 10.0.26300/build 26300, DisplayVersion 26H2; non-administrator observation token; NTFS fixed drive | T02 baseline/environment + T23 completion/reports | T23 actual native application tests PASS; physical Explorer/desktop acceptance remains; OS support channel unestablished |
| Windows PowerShell 5.1 on reference desktop | Required | 5.1.26100.9444 Desktop x64; test Process RemoteSigned; ordinary Restricted/all scopes Undefined | T23 completion/results/reports + prior task evidence | T23 full relevant suite: 854 checks / 29 tiers PASS; controlled classes separate; physical Explorer, CI and release gates remain |
| Supported PowerShell 7 x64 on reference desktop | Required | Supported portable7.6.6 Core x64; Microsoft-signed host; Process RemoteSigned | T09 PS7 acquisition + T23 completion/results/reports | T23 same relevant suite: 854 checks / 29 tiers PASS; OS channel unestablished; broader desktop/release compatibility remains |
| PDFtk native Windows build | Required | Vendor Server2.02 unsigned x86; approved exact external cache | T03 acquisition + T23 completion/reports | T23 paths, ordered pages, structural validation, job bounds and source/publication safety PASS; CJK operands/inspection remain safely refused; no broader release claim |
| Ghostscript native Windows build | Required for email support | Verified10.08.0 x64 console/interpreter unsigned; signed installer; explicit external cache | T09 GS acquisition + T23 completion/reports | T23 paths, both presets, validation, warning disposition and email outcomes PASS; PDFtk inspection limits retained; visual/desktop/package/release gates remain |
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
full NativeRunner/Launcher/Python fixture suites were not repeated; T11â€“T15,
OS support-channel/UNC/Explorer/fidelity/CI/full-release gates remain pending.
See T10 completion/C1 reports/review/live sync and cache audit.

T11 clean C1 `822f84eb4f7f0d75736887d9d79d0d363ae551a7` passes287 focused cases in each actual PS5.1.26100.9444 /
supportedPS7.6.6 host,574 total, all bad counts zero. AC025/26 pass: actual whole-job
named damaged/protected/zero-page/ambiguous-label refusal and exact natural ordered
page inventory, with eight independent PDFium visible-ID checks per shell. Actual
corrected CR structural/footer, incremental, xref-stream and GS10.08.0 linearized
fixtures preserve pages;75 source/foreign snapshot references per shell match.
Exclusive sharing denial stops before native launch; controlled source-change/
parser/overflow/logging cases remain unit evidence. Guard is bounded plausibility,
not universal corruption/encryption detection, full fidelity or transactional
snapshot. Static0errors/31warnings/6information retained/reviewed nonblocking;
selected approved caches reverified with no new acquisition/system change.
T12â€“T15, OS support-channel/UNC/Explorer/fidelity/CI/package/release gates remain.
Exact commands/hashes/results/limits and separate failed/tolerance-only history:
T11 completion/C1 reports/results/review/live receipts.

T12 clean C1 `1a4901b8914d02af6c539bc5b82181705ca3c2e9` passes324 focused cases in each actual
PS5.1.26100.9444 / supportedPS7.6.6 host,648 total,all bad counts zero.
AC027/28/29 pass: genuine PDFtk2.02/GS10.08.0 successful shared staging,
file/directory collisions after real native success,actual same-second overlapping
selected-shell processes with independent staging/final identities,and preserved
source/foreign/master snapshots. Controlled collision scheduling and barriers
remain disclosed; locks/changed ownership/late-child marker restoration are unit
filesystem evidence. Static0errors/51warnings/11info over five-file scope is
retained/reviewed nonblocking. Current nonempty-only publication is not structural
master/email acceptance or full fidelity;T13–T15 and later gates remain.
No acquisition,admin,persistent policy/PATH/security change or broader OS/UNC/
Explorer/release claim. Exact commands/results/reviews/history/live receipt:
T12 completion/C1 results/reports/native audit.

T13 clean C1 `e74ffa92588eff876c12ebd3321e2f3b0d6f110c` passes369focused cases each actual
PS5.1.26100.9444 / supportedPS7.6.6,738total,all bad counts0. AC030/31pass:
separate complete bounded masterinspection/exact frozen count before publication,
actual single2page and natural1/01/2/10 fivepage mixed0/90/180/270rotations,
independent PDFium IDs/rotation/dimensions and unchanged source/foreign hashes.
Native substitutions and scheduling are disclosed,not naturally failing vendor
output. Independent audit rereads four masters with realPDFtk/pinnedPDFium.
Static0errors112warnings49information over10files eachshell retained/reviewed
nonblocking. Narrow structural/localWindows evidence is not full fidelity,
universal validity/security,signatures,OS support-channel/UNC/Explorer/CI/package/
release acceptance. T14/T15 and later gates remain. No new acquisition/admin/
persistent policy/PATH/security/runtime network/engine change. Exact commands,
reports/history/reviews/native audit/live receipt: T13 completion/C1 evidence.

T14 clean C1 `588a96586518059ed38e2de01ad2312003defc7c` passes476relevant cases each actual
PS5.1.26100.9444 / supportedPS7.6.6,952total,all bad counts0. AC032/33integration
and AC034unit pass: explicit skip/missingGS/no-benefit master-only0, separately
validated strictly-smaller email0, realGS partial/count/post-master failure2 with
validated master retained, actualBAT0/1/2 and explicit final publication state.
PinnedPDFium reads visibleID/count/rotation/dimensions; source/master/foreign
snapshots match. Controlled corruptinput/wrongcount/discovery/logger faults and
owned staging4096whitespace preparation remain disclosed. GS CJKstage conversion0
then PDFtk2.02 inspectionfailure refuses publication on this host. No new broad
path-support/fidelity claim. ScopedPSA0errors108warnings42info each10files retained/
reviewed nonblocking. HistoricalPS5.1 host/locale test failures remain separate
from corrected dirty passes and clean952. Independent code/static/native/evidence
review and fresh C1live equality retained. No acquisition/admin/persistent policy/
PATH/security/runtime network change. T15/T16/T17 and OS/UNC/Explorer/fidelity/CI/
package/release gates remain. Exact commands/hashes/results/limits/history:
T14 completion/C1results/reports/review/native audit/live receipt.

T15 clean C1 `53d0923c95a86ae6a44bc89bab51cac6786c1e32`:17 tiers,559 passes each actual
PS5.1.26100.9444 / pinnedPS7.6.6,1118total,all bad counts0. AC035/036 integration
and AC037 unit pass on standard-user localNTFS Windows11x64/build26300. Actual
Win32 unset/empty/value GS_OPTIONS survives real engine success, invalid-image
start failure and controlled logger errors. Compiled inherited nested-process
cancellation before/after master stops owned descendants, preserves unrelated
same-image sentinel and validated master. Real locks and injected denied/full/
write/log/cleanup/ownership receipt faults remain truthful. Independent M2
source review no blocks; native2515checks/26freshPDFtk+PDFium retained-final
reads and evidence4611checks pass. PSA0errors152warnings45info each13files,
reviewed nonblocking; not lint-clean/fullT22. Dirty/preparation failures, missing
bootstrap XML/summary and corrected focused passes stay separate from1118.
Physical Ctrl+C/host/window/crash cleanup, broadACL/diskexhaustion, external
brokers,OS/UNC/Explorer/fidelity/signatures/PDF-A/CI/package/release not certified.
Approved selected cache bytes reused, no acquisition/admin/persistent policy/
PATH/security changes. M2 complete within stated scope; T16/T17 and later gates
remain. Exact commands/environment/hashes/limits: T15 completion/results/manifest,
M2review/nativeaudit/evidencereview/provenance/C1live receipt.

T16 clean C1 `26ac1b73e3733a23099de53d944e00e4ee412982`:14tiers,536each actual
PS5.1.26100.9444/pinnedPS7.6.6,1072total,all badcounts0. AC038integration and
AC039unit pass,standarduser/localNTFS/build26300. Real fixedscreen/ebook and
case-insensitive selection,named/default destination,actualcmd/BAT PS5.1,
closedstdinusage and SkipEmail bypass/explanation. Native2273checks/18cases/
28freshretainedPDFtk+PDFiumreads; archive6397checks/335files pass.
Source/runtime reviews no blocks; PSA0errors71warnings34info each7files
nonblocking,not lint-clean/fullT22. Dirty/preparation failures kept separate.
PhysicalExplorer,T17fidelity/size reporting,broadOS/UNC/CI/security/package/
release/signatures/PDF-A/universal preservation remain unclaimed. Approved
cache reuse,no acquisition/admin/persistentpolicy/PATH/security changes.
M3inprogress,T17next; exact commands/context/hashes/limits:T16completion,
results,manifest,reviews,provenance and C1live receipt.

T17 clean C1 `040176695fdb79e614ba2a821118fbc979a33115`:16tiers579each actualPS5.1.26100.9444/
pinnedPS7.6.6,1158total32reports,allbadcounts0. AC040integration/AC041manual
pass on standarduser/localNTFS/build26300. Real exactsize/decimalreduction/
candidate omission/success0 plus actual Codex40page10unique144DPI synthetic
visual review bothshells; screen smallest rastertext/outlines degrade, ebook
clearer, vector-only grows/omits. Not owner/Explorer/manualdesktop acceptance.
Independent source/runtime/native/archive reviews pass; archive11660
checks/632files. PSA5PSfiles0errors51warnings14info each nonblocking,
notfullT22. Dirty/preparation/absences separate. Approved cache reuse/noacquisition/
admin/persistentpolicy/environment/security changes. PhysicalExplorer/broadOS/
UNC/CI/security/package/release/universalfidelity/signatures/PDF-A remain open.
M3inprogress,T18nextpending; exact commands/hashes/context/limits:T17completion/
results/manifest/manualreview/source/runtime/native/archive/provenance/C1live.

T18 clean C1b `e506d73797379f355a1a0b731c857e71f4c1d251`:17tiers590each actual
PS5.1.26100.9444/pinnedPS7.6.6 x64,1180total34reports,allbadcounts0.
AC042integration/AC043independentreview pass,standarduser/localNTFS/build26300.
Real Get-Help/four examples/five documented routes (explicit percent route
uses PS5.1 under either outer context), measured stages/local owned logs,
truthful count/page/version/size/outcome and both native streams. Controlled
faults remain distinct. Source/runtime/native/diagnostic/archive reviews pass;
diagnostic5720checks/94observations/90native receipts. PSA12PSfiles0errors/
97warnings/82info each reviewed nonblocking,notfullT22/lint-clean. Actual
dirty/C1a/preparation failures kept separate. No fresh visual/manual desktop
claim,private binary upload,acquisition/admin/persistentpolicy/environment/
security change. PhysicalExplorer/broadOS/UNC/CI/security/featurepreservation/
package/release/signatures/PDF-A remain open. M3inprogress,T19nextpending;
exact commands/hashes/context/limits:T18completion/results/manifest/provenance.

T19 clean C1 `50220eccd1917e44a52d94cbfe3bf35b1940f3d8`:6tiers414 each actualPS5.1.26100.9444/
pinnedPS7.6.6 x64,828total12reports,allbadcounts0;12oraclegraphregressions.
AC044integration/AC045independentreview pass,standarduser/localNTFS/build26300.
OriginalCC0corpus54checks; actual master-only/screen/ebook separate structures.
Renamed second form; missing document attachment index/tag tree; intact page
payload hashes/navigation offsets; no editable email fields/widgets. Root48dirty
pages10unique144DPIgroups viewed;48cleanpagepixelmatches. Masterpixelsoriginal;
screenorientationchanges,ebooksidewaystext. Not GUI/Explorer/ownerdesktop,
signature/XFA/PDF-A/accessibility/malware/universalarchivalcertification.
Independentfeature/docs1101checks/cleansource504checks pass; PSA3PSfiles0errors/
16warnings5info each reviewednonblocking,notfullT22. Earlier dirty/preparation/
receiptfailures disclosed and excluded from828. No runtime/flags/defaults/
package change,acquisition/admin/persistentpolicy/environment/security changes.
PublicT19-onlytextarchiveidentity/pathsubstitutions/rawhashbindings; synthetic
PDF/PNG/nativebinaryhash-only. BroadOS/UNC/CI/security/package/release remain
open. M3inprogress,T20nextpending; commands/context/hashes/limits:T19completion/
results/manifest and independent reviews. C2sync reported after normalpush.

## T20 documentation and policy review checkpoint

Clean C1 `a9e93319ad260646d57278f207872a0e03dbb155` on standard-user Windows x64/localNTFS:
PublicDocs18/PreservationDocs14/Parameters31/Diagnostics36 per actualPS5.1.26100.9444
and pinnedPS7.6.6,198total/eightNUnit pairs,allbadcounts0. These are docs, isolated
actual ParamBlock and controlled entry/helper checks; no nativePDF/Explorer pass.
AC046/M3 review31semantic/35links/6immutable and AC04733policy checks pass.
Scoped PSA1.25.0 two test-tooling files0errors/4warnings each, reviewed nonblocking.
MIT/license/runtime/defaults/nativeflags/priorPDFobservations unchanged. No acquisition
or persistent policy/environment/security change. Supporting AST readerPS7.6.5 only
proves syntax; required test evidence usesPS7.6.6. Broad OS/UNC/fullnative/lint/security/
CI/package/release gates remain open. Evidence:evidence/T20-completion.md.

## T21 synthetic corpus and source-safety checkpoint

Clean C1b `abf8976e84f2c3f851efc42a844037a880519b26`, standard-user Windows
reference desktop/build 26300/local NTFS: 451 checks per actual PS5.1.26100.9444
and pinned PS7.6.6 x64, 902 total/22 NUnit pairs, all bad counts zero. AC048
review/AC049 native integration pass; 40 Python tests, 23 semantic checks, 73 corpus entries/
46 valid PDFs, 71 fixed entries reproducible. Native audit: 42 observations,
880 snapshot object checks, 38 fresh PDFium reads, 348 dependency files rehashed.
Actual concurrent intervals overlap 5.57s/5.07s with separate identity/stages and
exact new-output union. Directory modified timestamps excluded after observed
pre-invocation metadata change; all file invariants and tree inventory remain.
Scoped eight-file PSA 1.25.0: 0 errors/44 warnings/59 information per host, reviewed
nonblocking harness findings, not full T22. Runtime/defaults/native flags unchanged.
No acquisition/admin/persistent policy/environment/security changes. Earlier
preparation/C1a guard failures remain separate. Full fault/static/broader native,
Explorer/OS/UNC/CI/security/package/release gates remain open. Evidence:
evidence/T21-completion.md, results, manifest and archive review. T22 next.


T22 passes scoped fault/static checks at clean C1 `d159486cdfb66c39cf3ca6b35a23ebd08e1b2932`:
actual PS5.1.26100.9444 Desktop/pinned PS7.6.6 Core, 715 checks each,
53 maintained PS files/41 selected analyzer1.25.0 rules, all required bad counts
zero and no suppressions. Vendor-default advisory0errors/298warnings/161info
remain visible. Real engines are used only in the explicitly classified affected
native tiers; no mock/skip/inconclusive or missing environment is native evidence.
The unchanged supplemental handoff helper symlink case skipped because creation
was not permitted. This does not certify live UNC, physical Explorer, OS support,
CI, package/release or the full T23 integration gate. T22 environment rehashed348
selected dependency files unchanged, using scoped child policies and standard
user localNTFS with no persistent security/environment change. See T22-completion.

## T23 current native integration scope

Clean C1 `8fa2032c66f94199b121fc1914792d6d71bb6202` passes the same 29
relevant tiers in both actual required hosts: 854 checks each, 1708 total,
58 original NUnit/JSON pairs, every failure/block/container/skip/not_run/
inconclusive count zero. Declared controlled and documentation classes remain
separate from actual PDF-engine evidence. All source guards and 348 selected
approved dependency rehashes pass. Runtime and familiar defaults are unchanged.

New real-helper cases verify 138 distinct pages at a 29475-character command,
180-input / 38337-character prelaunch refusal, both presets on 24 original
vector/raster pages, and logged nonfatal GS warnings with strict validation and
no-size-benefit omission. The malformed warning fixture is explicitly refused
by application preflight. Original reports, independent raw/public audits and
eight extra PDFium audit reads support the scoped results; modest sample timings
are not a throughput or fidelity guarantee. See `evidence/T23-completion.md`
and `evidence/T23-results.json`.

OS support channel remains unestablished. Physical Explorer/desktop, CI,
security, broader OS/UNC, distribution and final publication remain subsequent
gates. No Windows 10, ARM, 32-bit host or live UNC pass is inferred.
