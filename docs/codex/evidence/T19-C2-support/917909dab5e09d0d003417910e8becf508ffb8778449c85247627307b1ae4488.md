# T19 development checkpoint

Started clean/local/live at `7e0fe6fde769911f323bd87e4c4f9382d26e331b` on
`codex/v1.0.0-readiness`; fetch/push origin remains
`https://github.com/PikkuJanne/WinPDFMerger.git`. Owner merged PR18; fetched
main `802b13da6785cc12c43158d1f6b63e6fb40a12da` has the same tree. No tags or
releases observed. T18 remains complete; publication remains NOT STARTED.

T19 will characterize a deterministic original CC0 two-document feature corpus:
links/bookmarks, canonical forms and widget appearances including repeated
field names, annotations, rotations, embedded attachments and minimal tags.
Master and screen/ebook email observations remain separate; rendering/page
success is not a structural, accessibility or signature guarantee. Signed PDFs
and XFA have explicit preservation non-guarantees rather than fabricated tests.
Application/native flags and default workflow remain unchanged. No flattening
or repair is added to the product. New parser/generator tooling is development-only.

Approved caches and preinstalled bundled Python 3.12.14, ReportLab 4.4.9,
pypdf 6.10.0 and pypdfium2 5.13.0/PDFium 153.0.7999.0 are being rehashed.
No acquisition/admin/persistent policy/environment/security changes. Child
RemoteSigned and case-insensitive child-only module-path removal remain scoped.
Synthetic binaries and PNGs stay ignored; publish privacy-reviewed text receipts only.

AC044 integration and AC045 review remain not_run until clean tested C1,
independent review and records-only C2 closure. T20 is pending and unstarted.
Actual development/preparation results and limitations will be added below;
this plan is not executed acceptance evidence.

## Actual development results before C1

The approved cached/native binaries and 332 development package Python sources
were rehashed; no dependencies were acquired. The ordinary standard-user host
reported Windows PowerShell 5.1.26100.9444 Desktop x64, Windows 10.0.26300.0,
NTFS, Restricted policy and all five policy scopes Undefined. Actual test children
use RemoteSigned and remove inherited PSModulePath only in their environment.
Pinned PowerShell 7 is 7.6.6; Pester is 6.2.0, PDFtk is 2.02 and Ghostscript is
10.08.0. Python 3.12.14, ReportLab 4.4.9, pypdf 6.10.0, pypdfium2 5.13.0,
PDFium 153.0.7999.0 and Pillow 12.3.0 are development inspection tooling.

First original PDF authoring followed the PDF skill's marker exactly once.
The generator SHA256 was
`a6c4146fba847ccb4c961c11277456503bed61dc3db3672cc9a0fb2aebfddbd8`;
the tracked manifest binds the actual two generated PDF hashes and model.
Independent original characterization passed 54 scoped checks, with strict
parser warnings empty and original bytes unchanged. Twelve oracle regression
tests passed using in-memory object graphs only; they are not native app tests.

Actual focused development commands, from the repository root with the approved
Python executable, used the retained task-local execution wrapper and driver:

```text
python -B tests/.work/Run-T19Command.py --name T19-dirty-ps51-native --script tests/.work/Run-T19Tests.py --shell ps51 --phase dirty --tiers PreservationNative
python -B tests/.work/Run-T19Command.py --name T19-dirty-ps7-native --script tests/.work/Run-T19Tests.py --shell ps7 --phase dirty --tiers PreservationNative
```

Both six-case native runs passed, with all failure/skip/not-run counts zero,
against dirty sources based on start commit
`7e0fe6fde769911f323bd87e4c4f9382d26e331b`. These invoked the unchanged copied
application under each actual pinned shell, using explicit acquired-engine paths,
master-only, screen and ebook routes. Sources and foreign outputs were preserved.
The measured forms/navigation/annotations/attachments/tags/rotation observations
are documented in `docs/PDF_LIMITATIONS.md`; retention losses are observations,
not a reason to add flattening, repair or different application flags.

The root reviewer viewed all ten distinct full-page pixel groups at 144 DPI,
covering 48 rendered pages from the two actual native runs. Pixel-hash bindings
show all masters matching originals. Screen changed rotated-page text to upright;
ebook retained sideways text orientation with changed note appearance. This is
a limited visual review, not interactive editing, physical Explorer acceptance,
signature validation, accessibility, PDF/A or malware certification.

Development failures are retained rather than counted as passes: the first
inventory producer did not expand a Windows environment-variable path and exited
1, then passed after a task-local path fix. Both first PreservationDocs runs had
14 assertions pass but one failed container in AfterAll due to generic-list array
conversion; each complete run therefore failed. Independent review also found
the fixture README's incorrect generator option, corrected to `--output`.
The test receipt helper now uses explicit generic-list `ToArray()` conversion.
Actual `PreservationDocs` rechecks passed 14/14 on each pinned shell, with every
bad count zero. Their argv differ only in the documented tier and shell labels
from the focused native commands above. The first failed streams/reports remain
retained and are not included in successful totals.

Actual PSScriptAnalyzer 1.25.0 ran over the two new tests and modified runner
under both pinned hosts. The final sources yielded zero errors, 16 warnings and
five informational findings per host. Earlier runs and correction of an automatic
variable name and trailing whitespace are retained separately. Remaining findings
concern Pester cross-block variable use, test receipt output, helper naming and
positional style; this is static test tooling evidence, not native preservation.

Artificial seeded, nonpainting comments ensure synthetic email outputs satisfy
the normal smaller-file rule. The measured sizes are not representative image
quality or compression evidence. No signed/XFA corpus or validators are present;
their preservation is expressly unvalidated. Clean C1 verification and C2 records
closure remain pending, and T20 remains unstarted.

## Clean C1 closure

The development chronology above is historical. Clean implementation C1
`50220eccd1917e44a52d94cbfe3bf35b1940f3d8` passed 828 Pester cases across12 reports (all bad counts zero), plus12
oracle graph regressions. AC044 integration and AC045 independent review pass.
Full actual receipts, commands, clean review/visual bindings, retained failures
and limitations are linked in [completion](T19-completion.md),
[results](T19-results.json) and [archive manifest](T19-C1-reports/manifest.json).
T20 is pending and next. Records-only C2 synchronization is verified after its
normal push and reported in the session, without a self-referential future SHA.
