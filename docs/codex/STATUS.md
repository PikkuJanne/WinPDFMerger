# Project status

Target: published and independently verified v1.0.0 in PikkuJanne/WinPDFMerger.
Completed milestones: M1/M2/M3 within recorded Windows/test/review scope.
Current milestone: M4, in_progress. Completed tasks: T01 through T25.
Current task: T26 - desktop acceptance and compatibility scoping, in_progress.
Publication: NOT STARTED. No release or tag exists in the fresh T26 inspection.

T26 has documented the permitted Windows10/liveUNC/ARM/32-bit-host validation
exclusions (AC060-AC062 are excluded, never pass), with public rationale and a
documentation regression. Required AC058 Explorer/visible PDF acceptance and
AC059 complete required-environment review remain not_run. Synthetic walkthrough
preparation cannot replace supplied human observations. T27 is not dependency-ready.

Fresh targeted registry observation identifies current Windows Pro26H2 full
26300.9457 and a standard-user x64 token. Microsoft's current release table lists
that exact GA revision; prior OSVersion10.0.26300.0 omitted the revision. Insider
enrollment/channel still requires Settings evidence. Historical receipts retain
their own dates and limited observations. See evidence/T26-compatibility-review.md.

PR25 was freshly observed MERGED at main e245114. The clean readiness checkout
safely fast-forwarded from6180b73 to same-tree main; no runtime/source change or
history rewrite occurred. This current platform observation supersedes the old
continuation's draft/open status; immutable T25 evidence remains unchanged.

T23 clean C1 8fa2032 retains actual PS5.1.26100.9444/pinnedPS7.6.6 dual-shell
native evidence,854checks/29tiers each. T24 retains actual hosted Windows Server/
admin CI and deliberate failure gates; T25 retains scoped security/dependency/
privacy review. These are supporting evidence, not physical Explorer acceptance.

Clean T26 documentation/static test results and its normal push/live equality
will be recorded after execution. M4 and later package, accepted-source,
publication/download/closure gates remain incomplete; only final v1.0.0 is allowed.
