# T13 implementation checkpoint

Started 2026-10-08 from clean `d9e2038934abd3e0554f8b814c73ec3ee9072c2d`
on `codex/v1.0.0-readiness`; live same-branch ref matched. Origin fetch/push is
`https://github.com/PikkuJanne/WinPDFMerger.git`. Owner-merged PR12 and live main
`c7cb75ef6aa5f2d55025e8af3ed17e4a7d86eace` were inspected; no tags/releases exist.
No reset, stash, force, origin change or release action is part of T13.

Every PDFtk merge now requires a positive frozen ExpectedPageCount. Native merge
success and a nonempty staged regular file precede bounded PDF envelope/native
inspection with the same selected PDFtk. Complete successful inspection and an
exact positive expected total are required before final no-overwrite publication.
Ownership/reparse guards and staged length/UTCmtime are rechecked afterward.
Merge and validation native receipts remain separate; validated and published
states remain distinct on collisions. Entry rechecks frozen inputs before merge,
logs a planned path first, then advertises the master/pages after validated move.
Ghostscript follows successful master publication; email validation remains T14.

Dirty development verification against the starting SHA is not immutable
acceptance: 38 new unit cases plus12 controlled invocation cases pass50/50 in
each actual PS5.1.26100.9444 / supported PS7.6.6. Earlier49/49 focused runs remain
separate history before the additional missing-native-receipt regression.
Actual MasterValidation7/7 passes each shell; Staging9/PdftkPaths13/GSPaths13
pass in PS7. Native cases include single2page and natural1/01/2/10 fivepage
entry masters, independently checked visible IDs/rotations/dimensions via pinned
PDFium, unchanged source/foreign hashes, real wrong expected total refusal and
disclosed staged substitutions after genuine native success. No failed run has
been observed in T13 development; clean acceptance still requires committed rerun.

Scoped precommit PSA1.25.0 over ten changed PowerShell files:0errors/112warnings/
49information each shell. Findings remain retained for independent disposition;
this is not lint-clean/fullT22. Approved-cache read-only audit rechecked14 selected
files across5 dependencies and Python3.12.14/PDFium pins. No acquisition/admin/
PATH/persistent policy/security change. Ordinary standard-user PS5.1 Restricted/
all-scopesUndefined was observed; RemoteSigned is authorized test-child-only.

Required AC030/31 remain pending clean C1 dual-shell eleven-tier rerun, independent
review/native audit, normal same-branch push and fresh local/live equality. C1
cannot name its own future SHA; records-only C2 will bind the tested C1 and its
live receipt, with C2's own post-push equality reported in the session. Narrow
local synthetic Windows structural evidence does not prove universal PDF validity,
security, fidelity/signatures, transactional snapshots, Explorer/UNC/OS support,
CI/package or release acceptance. T14/T15 and remaining project gates stay open.
