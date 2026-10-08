# T12 implementation checkpoint

Started 2026-10-08 from clean `280468546a63cc89dc617491f29d823790db6bab`
on `codex/v1.0.0-readiness`; live same-branch ref matched. Origin fetch/push is
`https://github.com/PikkuJanne/WinPDFMerger.git`. Owner-merged PR11 and live main
`ce27e08841063d82fcb1ce86a7a427b9c17899c7` were inspected; no tags/releases exist.
No reset, stash, force, origin change or release action is part of T12.

Implementation reserves one atomic create-new Windows staging directory per run
beneath the destination, holds a readable immutable ownership marker, records
physical directory identities and restricts publication/cleanup to known files.
Final publication uses the same-parent two-argument no-overwrite move. Shared
entry cleanup is in finally; master publication survives email failure. Email
summary uses this run's explicit published state. Unknown content, locks and
changed ownership fail cleanup safely with exact-path manual diagnostics. A
failed empty-directory deletion restores marker evidence best effort only after
original physical identity checks. No orphan/prefix sweep or recursive removal.

Dirty development verification against the starting SHA (not immutable acceptance):
full Unit225 per shell, ToolInvocation12 per shell, focused Staging9 per shell.
Actual PS5.1.26100.9444 / pinned supported PS7.6.6, Pester6.2.0,
PDFtk2.02 / GS10.08.0. All these corrected runs have zero failures/skips/not_run.
Unit includes28 new staging faults; native tier uses genuine engines, fixed
publication collision scheduling and actual concurrent shell barriers.

Historical first ToolInvocation run had6pass/6fail per shell: a ReadWrite marker
handle prevented File.ReadAllText and the path-budget message lacked260.
The marker now retains a read-only/no-write-delete-sharing handle, and the
diagnostic retains the explicit limit. Regression assertions cover readability
and immutability. A separate late-child unit fixture mock initially had27pass/1fail
per shell because unmatched parent identity returned null; corrected explicit
real default behavior passes28/28. Original historical reports remain separate
from clean acceptance counts. No native gate was inferred from these faults.

Precommit PSA1.25.0 on entry/helpers/runner/new tests:0errors/51warnings/11information
per shell. Findings remain retained/reviewable, not lint-clean or fullT22.
Read-only approved-cache audit rechecked14 selected files across5 dependencies
and Python3.12.14/PDFium pins. No acquisition/admin/PATH/persistent policy/security
change. Ordinary standard-user PS5.1 Restricted/all-scopesUndefined was observed;
RemoteSigned is authorized test-child-only. Local Windows11x64/build26300 NTFS
synthetic evidence, not Explorer/UNC/OS-support-channel/full-fidelity acceptance.

AC027/28/29 remain pending the committed clean dual-shell rerun, final evidence,
normal push and fresh local/live equality. T12 retains the existing nonempty-only
publication gate; structural master validation is T13 and email validation/size
outcomes are T14. Interruption/descendants remain T15. C1 cannot name its own
future SHA. Final records-only C2 will bind the tested C1 and its live receipt;
C2's own clean/post-push equality is reported in the session.
