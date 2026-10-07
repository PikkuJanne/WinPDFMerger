# Destination and shared run identity tests

The `Destination` tier executes the real entry point in isolated application
copies on Windows. It requires explicit real PDFtk and Ghostscript paths, gated
against their recorded acquisition hashes. No dependency is downloaded or installed.

```powershell
& $selectedShell -NoProfile -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 `
    -PesterModulePath $pesterManifest -Tier Destination `
    -PdftkPath $pdftkExecutable -GhostscriptPath $ghostscriptExecutable
```

Use the previously authorized test-process policy and verified external caches;
select the recorded PowerShell 7.6.6 host explicitly. Run separately in Windows
PowerShell 5.1. The tier fails closed if actual required tools/Windows are absent.

AC022 covers default/explicit writable destinations, missing/file/wildcard/provider
failures, actual standard-user write denial on one owned application directory,
and explicit-output recovery. The ACL fixture restores the original descriptor
in `finally`; source and existing foreign-file snapshots are compared. The
create-new/DeleteOnClose probe must leave no residue even after later dependency
failure. A directory's ReadOnly attribute alone is not a Windows write-denial test.

AC023 covers physical same-directory/case identity and real leaf/ancestor junction
paths. Junction/reparse paths are explicitly unsupported at preflight. A read-only
short-name alias observation, where the filesystem provides one, is separate from
mandatory case/junction evidence; no volume short-name policy is changed.

AC024 unit cases cover root/empty/dot-only labels, trailing dots/spaces/separators,
punctuation/Unicode, surrogate-safe truncation, invariant timestamp formatting,
complete final/private-output budgets and create-new collision refusal. The native
tier starts two actual PDFtk/GS application processes together, verifies distinct
shared master/email/log identities and unchanged source/foreign bytes, then inspects
the synthetic PDF page totals. Counts are not visible-fidelity certification.

Directory identity uses a small lazy Windows `FILE_ID_INFO` adapter (volume ID and
128-bit file ID). It does not compile/run when helpers are imported. Unsupported
or ambiguous metadata fails closed; only the actual filesystem tested in the
evidence is a compatibility claim. Reparse ancestors are checked before any probe.
This is ordinary path preflight, not a transactional filesystem snapshot against
an unrelated process replacing directories while a run executes.

The implementation has no runtime network calls, source renaming, extra native
PDF engine, administrative step or fallback output location. Input inspection,
general staging/state validation, email outcomes and full interruption cleanup
remain T11–T15. The public `OutputFolder` option begins in T10 because AC022 needs
it; T16 owns the remaining parameter interface.
