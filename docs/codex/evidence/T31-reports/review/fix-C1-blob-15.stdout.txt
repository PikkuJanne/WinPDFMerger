# T31 — unaccepted merged candidate and checkout-byte fix

PR26 was reviewed at277e8cbb7de98b4cb07850def58590473ec636b9 and merged through
the normal allowed merge path at2026-10-09T17:46:48Z. First candidate
de5f30155c68755dbd5af691625a0651e3fb7230 has the exact reviewed PR tree and parents
e2451141217efdd00a1d49d72a04df054872dffc and277e8cbb7de98b4cb07850def58590473ec636b9.
Clean local/live main equality was observed at17:47:03UTC. No tag or release exists.

This candidate is **unaccepted**. Its fresh Windows checkout under the existing
system core.autocrlf=true converted four LF-pinned development fixture files.
The actual Python fixture-oracle suite ran40tests:34pass,1failure,5errors; exit1.
The original stderr SHA-256 is
7c11d8c17c901c3ec2b8ddd0bc6c0f775f5b1dac722a3529594b3f0986fbcf88.
The numbered manifest failed exact reproducibility, and five corpus tests rejected
raw recipe/manifest pin mismatches. These are required developer-tool failures;
successful scoped CI and static checks do not make this candidate accepted.

The fix adds five explicit -text rules. Four canonical LF blobs already match
the existing pins. The presets manifest's original pinned CRLF bytes are now
stored verbatim in Git, removing its opposite autocrlf=false mismatch. All
catalogue expectations, hashes, fixture semantics, existing feature/PDF rules,
runtime/version/native arguments, builder and allowlist remain unchanged.
Future evidence byte rules can remain inside docs/codex/evidence/.gitattributes,
which is allowed on the separate M6 evidence-only branch after source freeze.

The new test creates a small owned local Git repository, commits the fixture
bytes and attributes, and performs two real fresh clones with command-scoped
core.autocrlf=true and false. Both call actual corpus.load_catalog and compare
numbered generator output with checked-out bytes. The initial regression draft
failed before the fix (2errors,exit1); original stderr SHA-256
5e52d98f2e71a6de473b7a14783e35a214c5b5267c595d79efd64209fd09090c.
Both passed afterward; the full fixture helper suite passed42tests without skips.
The final test adds command-scoped empty hooksPath to isolate global hooks;
the original draft is reconstructed byte-for-byte and matches its recorded SHA
72e653efc3664c1ab2af58b92b7d7430ca7cadda7b2da8a433c4dc255ccedbfe.
The final test source SHA-256 is
1ad01e47a75b06fed7eff56cdfa1b3eb1ef9394fa897f08c698862df08a0e220.
Green preparation executions are explicitly dirty-base executions at de5f301,
not accepted final-source evidence. Exact commands/times/source hashes and red/
green streams remain in owned tests/.work/T31-checkout-regression-{red,green}.
Python is the approved pinned3.12.14 runtime; no acquisition/elevation or
persistent policy, Git-config, PATH or security change was performed.

Final five-rule dirty-base preparation at18:02:02–13UTC ran
`<approvedPython> -B -m unittest discover -s <scope> -p test_*.py -v`
for tools/codex/tests, tools/test/tests and tests/package:26pass/1symlinkskip,
42pass/no skips and17pass/no skips respectively (85pass/1skip,all exit0).
Original stderr SHA-256 values in that order are
0871a620a3141fc65a653359ed2c5755b30bc59c0ab6c14fde9d6f2a4b53312d,
655925fd22427cd266b529d5454cbe53c1c8c9eda385ae9a57966396e6d68610 and
c43564d005a8c08a6b05163ac6c04aed0bb924bd730e52c5f5c32626ee937bab.
Their exact invocation ledger/streams remain in
tests/.work/T31-fix-preparation-1d9deb93c5a64a9f89938e051f5c4b98.
Record-only check-plan --require-ready passes30done/66pass/4excluded; it
does not execute the application or establish accepted final-source results.

First-candidate full native test captures were intentionally stopped after the
known required failure: PS5.1 completed15tiers/726passes and pinnedPS7 completed
20tiers/823passes. Incomplete tiers, remaining tiers and final guards are not
accepted. The first owned PS5.1 tree signal returned1 after terminating its
selected parent/three children; one child reported unsupported termination.
The separate PS7 signal returned0 and refreshed creation-identity query confirmed
the selected parent absent. No application cancellation/cleanup pass is inferred;
partial streams and owned synthetic leftovers are retained. Source stayed unchanged
before stopping; no broad process sweep or elevated retry was used.

AC071 remains not_run; AC072 is fail for the unaccepted candidate's actual
required fixture failure. T31 remains in_progress until the reviewed fix is
normally merged, clean/live main is verified, and all required exact new merged
source checks pass. T32–T34 remain pending. AC058 remains excluded, unperformed
and never pass. Final task records will preserve first-candidate failures and
bind the accepted source, original reports and independent reviews separately.
