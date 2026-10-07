# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Current milestone: M0 — baseline and safe setup.
Completed tasks: T01 and T02.
Current task: T03 — test seams and fixture harness, in progress.
Publication: NOT STARTED.

Prepared: four unchanged baseline helpers with import boundaries, exact Pester
6.2.0 runner, controlled native process, and original one/multipage PDFs with
independent native PDFium oracle/render review. Entry behavior/defaults remain;
narrow ignores keep intended synthetic PDFs trackable.

Complete AC005/AC006 gates remain not_run: exact Pester unavailable, PS5.1
normal scripts rejected by Restricted, PDFtk/GS undiscovered. Prerequisite
authorization requested; unanswered requests do not authorize installation or
policy changes. PS7 import smoke, native-renderer fixture and fake-process checks
are partial evidence. PS7 7.6.5 does not support a current-supported-build claim.

Start clean live readiness: `8aa2b63dd701f499ec9936ae2252cceed1e41b25`.
PR #2 was already merged; readiness safely fast-forwarded to live main
`86f56c133b92507f88e01dfdac5960bbce24e847` before edits. No tags/releases or
open PRs at inspection. Final records commit must be verified after push; its
own hash/live proof belongs in session output, never fabricated here.

Evidence: `evidence/T03-harness.md`. T01/T02 evidence retains its recorded
review/helper/primitive scope. No application, Explorer, package, CI or release
check passed. T04 onward remains pending. Close T03/M0 only after required gates.
