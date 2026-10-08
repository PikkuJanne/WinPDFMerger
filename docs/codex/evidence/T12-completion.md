# T12 — Completed owned staging and no-overwrite publication

Clean implementation C1: `1a4901b8914d02af6c539bc5b82181705ca3c2e9`.
Required AC027, AC028 and AC029 pass. Normal same-branch push and fresh clean
local/live equality are retained in `T12-C1-live-sync.json`. Draft
[PR12](https://github.com/PikkuJanne/WinPDFMerger/pull/12) follows owner-merged PR11.
Final C2 changes records only; its own post-push SHA/equality is reported in the
session without self-reference. Next task is T13; M2/release gates remain.

One atomically new `.WinPDFMerge_<32hex>.tmp` under the destination now owns both
master and email work. CreateDirectoryW decides the reservation race; an existing
candidate never acquires cleanup ownership. Create-new/flushed `owner.json`
records schema, run identity, stage suffix, creationUTC, PID and known files.
A retained read handle permits reading but denies write/delete sharing. Context
records original parent/stage physical identities and fixed master.pdf/email.pdf
paths. Import defines functions only; compilation/native work remains lazy.

Publication rechecks context/marker/identities/reparse paths, direct same-parent
final location and a nonempty regular known staged file, then two-argument
File.Move decides collision without overwrite/retry. Existing final files and
directories survive races after genuine native success. Shared entry staging is
cleaned in finally even on early exit. Published masters are never cleanup
targets and survive failed email publication. Explicit emailPublished replaces
file-existence inference in the summary. Defaults, input inventory/order and
native command vectors remain preserved.

Cleanup inspects only the owned stage, refuses unknown children/reparse content,
deletes only the two known temporary PDFs and marker, then the empty directory
nonrecursively. Locks/changed ownership/unknown content report the exact retained
folder for manual inspection after all runs stop; no orphan/prefix sweep exists.
If directory deletion fails after marker removal, restoration uses CreateNew only
after original stage AND parent identities/reparse checks. Restoration failure is
disclosed. Abrupt termination still cannot guarantee cleanup; missing markers
need manual investigation, never an automatic ownership assumption.

| Tier | PS5.1 | PS7.6.6 | Evidence class |
|---|---:|---:|---|
| Unit |225|225|Parser/helpers and controlled process/filesystem;28 new staging cases|
| Staging |9|9|Real PDFtk/GS/filesystem plus controlled publication/concurrency scheduling|
| InputPreflight |22|22|Actual entry/PDFtk and independent PDFium visible order|
| Destination |15|15|Actual standard-user entry/PDFtk/GS destination regressions|
| ToolInvocation |12|12|Controlled fixed vectors/private filesystem faults|
| PdftkPaths |13|13|Actual PDFtk path/noninteractive regressions|
| GhostscriptPaths |13|13|Actual GS path/noninteractive regressions|
| SourceDiscovery |4|4|Actual entry/PDFtk literal/top-level/order regressions|
| DependencyEntry |9|9|Actual entry faults/real PDFtk and controlled probes|
| LauncherNative |2|2|Actual BAT/application/PDFtk;child hostPS5.1|

324 each,648 total; failures/failed blocks/containers/skips/not_run all zero.
Each summary/XML identifies exact C1 and clean state. New native observations
include shared success, two existing-final refusals, four file/directory collision
races after real PDFtk/GS exit0, real PDFtk bad-data failure and overlapping actual
selected-shell children. Fixed same timestamp, separate run/stage/final identities
and controlled barriers prove the live second stage/sentinel survives first
cleanup; first final snapshots survive second completion. Source/foreign hashes,
lengths and UTCmtime remain identical. Controlled scheduling is disclosed and
does not count as an uncontrolled stress test or mocked-engine validation.
Independent native audit additionally reads16 retained outputs through exact
pinned PDFtk2.02:all exit0/pages3. These are narrow structural observations.

Unit28 staging cases cover atomic suffix file/directory collisions, marker
readability/immutability/disposed handle, known-only cleanup, foreign prefixes,
unknown children, actual locked files and junction targets, layout/identity/marker
faults, late-child marker restoration/refusal and publication path/shape/lock
boundaries. Controlled identity changes are distinguished from actual filesystem
effects. Clean orchestration bounds each suite at180000ms and owns dependency
fixture receipts via a bounded lock. Existing native bounds remain900000ms
execution,1000ms termination/capture and30000UTF16 complete command length;
all input/output operands remain below260.

Exact commands/pins/environment/results and raw/sanitized byte hashes are in
T12-C1-results/reports manifest. Independent C1 code/static/evidence review and
native audit are retained separately. Original dirty ToolInvocation6pass/6fail
each shell came from marker sharing and the path-budget diagnostic; both fixed
with regressions. The late-child unit fixture initially had27pass/1fail each shell
due to a Pester mock default; fixed28/28, transcript-only history disclosed where
no raw report existed. Corrected dirty runs remain separate from clean648.

Evidence Git attributes preserve exact retained byte hashes through Windows
line-ending conversion. Six stdout files retain captured empty FileVersion
diagnostics after `: `; only their trailing-space lint is waived. Staged Git blobs
are checked against retained file bytes before the records checkpoint.

Scoped PSA1.25.0 over entry/helpers/runner/new tests:0errors/51warnings/11information
per shell, retained/reviewed nonblocking,not lint-clean/fullT22. Groups concern
console output,private-helper ShouldProcess,naming,Pester scope/parameters and
existing comment-block BOM;information is OutputType/positional test helpers.

Actual local NTFS Windows11x64/build26300 standard-user evidence uses
PS5.1.26100.9444/supported portablePS7.6.6,Pester6.2.0,PDFtk2.02/GS10.08.0.
Approved caches:14 selected files across5 dependencies plus development
Python3.12.14/PDFium pins/hashes reverified. No acquisition,installer/admin,
persistent policy/PATH/security change,vendor redistribution or privatePDF upload.
Ordinary Restricted/all-scopesUndefined reobserved;RemoteSigned is authorized
test-child-only. No runtime network/Python/new engine/source renaming added.

Current publication remains nonempty-only; T13 adds expected-page structural
master validation before publication. T14 owns email validation/size/outcomes;
T15 owns interruption/descendants. No universal validity/fidelity/signature,
transactional snapshot,Explorer/UNC/OS-support-channel/CI/package/release claim.
Unchanged full NativeRunner/Launcher/generalPython-helper suites were not repeated
for this focused task; prior evidence remains historical. No tag/release change.
