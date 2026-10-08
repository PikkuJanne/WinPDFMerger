# T12 staging and publication checks

Run `tools/test/Invoke-Tests.ps1 -Tier Staging -PdftkPath <explicit vendor exe>
-GhostscriptPath <explicit vendor exe> -PesterModulePath <pinned manifest>`
in each actual required shell. The runner downloads/installs nothing. Tests use
original synthetic PDFs and owned ignored directories. Both native tools must
match the approved acquisition receipts before integration starts.

One run owns a new `.WinPDFMerge_<32hex>.tmp` beneath its destination. Windows
CreateDirectoryW atomically refuses existing candidates; an existing candidate
does not acquire ownership. `owner.json` is create-new, flushed, and retained
with a read handle that denies write/delete sharing. The context records original
directory identities, marker bytes and exactly `master.pdf` / `email.pdf`.
Importing helpers creates no files, compiles nothing and starts no orchestration.

Publication checks the context, path/identity/reparse boundaries, known staged
file and nonempty regular-file shape, then uses two-argument File.Move in the
same output directory. Its result decides races; there is no overwrite/retry.
T12 retains the current nonempty-file publication gate. Structural expected-page
validation is T13; email validation/size-benefit decisions are T14. These tests
do not claim full PDF fidelity, signature validity, Explorer or release acceptance.

Integration uses real native conversion plus controlled scheduling at publication
to create collision files/directories immediately before the final operation.
The concurrent test runs two actual selected-shell processes with a fixed same
timestamp and independent identities/stages. A barrier keeps the second stage
live while the first publishes and cleans. Its known sentinel must survive.
Native exit details, hashes, page inspections and exact cleanup results are
retained in the suite's observations; controlled scheduling is not a mock-engine
or uncontrolled stress-test claim.

Cleanup enumerates only its owned directory, rejects unknown children/reparse
substitution and deletes only known paths, then the empty directory. Locked files,
marker/identity changes and unexpected content retain an exact-path diagnostic
for manual inspection after all runs stop. After a failed directory removal, the
marker is restored best effort only in the same original directory. Restoration
failure is disclosed. No recovery scan or recursive deletion exists. Abrupt
termination cannot guarantee cleanup. Neither a PID nor a filename prefix alone
proves that an orphan is safe to delete.

Primary API contracts: [CreateDirectoryW](https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-createdirectoryw),
[File.Move](https://learn.microsoft.com/en-us/dotnet/api/system.io.file.move?view=netframework-4.8.1).
