# T30 coverage and merge-gate review: implementation checkpoint

Started 2026-10-09 from clean, live-synchronized
`681e0e3f97a44b8c23fb67f45aa9605e4d7fcd32` on `codex/v1.0.0-readiness`.
Both origin routes are PikkuJanne/WinPDFMerger; live main is
`e2451141217efdd00a1d49d72a04df054872dffc`. PR26 is draft/open/unmerged and
all eight current push/PR jobs are successful. No tags or releases exist.

Independent reviews cover all 17 improvements and before-merge evidence,
runtime/source safety, package provenance, public claims and dated primary
vendor information. Initial review found stale candidate-acceptance wording
and account-class language in public documentation. The changes report actual
T28/T29 premerge ZIP operation while retaining accepted-R/final/download gates,
Unreleased status, AC058 exclusion and unsigned/PDF/dependency/platform limits.
Two public-doc regression cases distinguish these evidence classes.
Runtime, native arguments, package allowlist and builder are unchanged.

The development-only capture-full.py reuses the explicit approved 348-file
cache inventory and selected shells/modules/engines. It requires a clean exact
source commit, records each child's actual argv, streams, exits and report
hashes, and checks source/cache/driver guards before and after. Full32-tier
dual-shell regression, maintained-source static checks, helper tests and final
independent review remain pending at this implementation checkpoint.
No tests, future push or merge readiness are claimed here. AC069/070 remain
not_run; T30 is in_progress. Final records will identify tested C1, actual
outcomes and independently reviewed evidence without self-referential hashes.

Preparation correction: refreshing the dependency heading date broke SECURITY's
local section link. The existing public-link regression caught it (PS5.1:
23 pass/one fail). Preparation commit `7aa1d57bdb5cacd7baffe6cb3ba7cbb2fb03f002`
was pushed before that result was inspected, so it is not an accepted test SHA.
The anchor is corrected; the full clean-source run must use the subsequent
corrected commit. Original failed preparation streams/report remain retained.

T31 merge/source freeze is later. No tag, release or publication is authorized
by a successful plan check alone; exact final and downloaded assets remain gates.
