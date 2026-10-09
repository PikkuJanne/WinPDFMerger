# T31 final record writer preparation

`Write-T31Records.py` has not been executed. It is prepared for root to use after
the final full runs, original/native/CI/final-gate reviews, immutable public packet,
independent public audit and external fresh main/freeze proof all exist.
`bindings.template.json` deliberately lacks their filenames/hashes and cannot
complete the gate. Populate a new ignored bindings file with observed facts.

The writer checks the entire frozen packet's inventory/hashes/XML, all32tiers in
both shells,64NUnit/JSON pairs/2144passes,68files/41selected static rules and
retained349/175advisories, exact-R2 ten extras commands/85passes/1skip, and exact-R2
main push CI fourjobs/20pairs/1370passes. All six required independent review
bindings must have their actual pass result, positive completed checks, no issues
and explicit JSON pointer identifying R2. The public audit must additionally bind
the exact manifest hash and its independent source hash.

The external proof schema agreed with root requires task/result, exact source/
local HEAD/local main/live main R2, branch main, clean/synchronized/frozen true,
UTC time and successful actual command receipts. Its time must follow both full
captures and be at most one hour old. Its raw hash must also appear in the public
manifest at the explicitly bound public label. The writer itself freshly checks
origin/main, both origin routes, exact R2/tree/parents, the evidence branch and
the frozen package/runtime/public-doc/version/builder blobs.

The only write surface is TASKS/ACCEPTANCE_CASES/RELEASE_STATE/STATUS/NEXT_SESSION
and T31 completion/results under docs/codex. All staged bytes are prepared before
writing. Default mode writes these seven files as ignored drafts; `--apply` is the
root-owned final tracked write. It performs no Git mutation, commit, push, PR,
package/tag/draft/publication action. It cannot infer a future evidence commit or
push result. R2's earlier in-progress/fail handoff snapshot remains immutable,
while later docs-only E1 records close AC071/072 with actual execution evidence.

Initial rejected R, strict fixture34pass/1fail/5errors, stopped35pairs/1549partial
passes, scoped first-R CI/static results, red2errors, dirty-base85pass/1skip,
260/72check reviews and reviewer preparation failures remain historical. Required
source failure is closed only by actual final exact-R2 evidence. AC058 and the
other existing exclusions remain unchanged; T32-T34 and all six later cases stay
pending/not_run. Release state remains not_started with exact source R2 and null
asset hashes/URL/publication time.

After root supplies completed bindings and opens the evidence branch:

```text
<approved-python> -B tests/.work/T31-record-preparation/Write-T31Records.py --bindings <completed-ignored-bindings.json>
```

Inspect the ignored draft diff and outcome, then root may use `--apply`. Root
must run the record validator, review/stage intended docs-only changes, commit/
push normally and report actual clean/live equality. Source runtime remains R2.
