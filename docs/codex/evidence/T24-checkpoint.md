# T24 checkpoint

Starting clean source: `cda0d8aed7a9892ec8679b82c3938a021d6320ac` on
`codex/v1.0.0-readiness`, equal to a fresh live origin read. Both origin
destinations resolve to `PikkuJanne/WinPDFMerger`. Live main is
`884275096087f8e38f1f87de9d978edfb4a1c1b7`; PR23 is merged, superseding the
earlier draft observation. No open PR, v1.0.0 tag or release was found.

T24 begins in progress. AC054/AC055 remain not_run during preparation.
Scope is restrained Windows hosted CI with separate controlled/unit and
actual native engine jobs, explicit PS5.1 and pinned portable PS7.6.6,
hash-checked temporary dependencies and sanitized machine-readable reports.
No runtime behavior, launcher, native flags, defaults or application download
behavior changes. Hosted administrator/Server evidence does not close desktop,
standard-user, Explorer or release gates.

Implementation, clean-commit executions, actual hosted failure verification,
independent review and fresh push/live synchronization will be recorded only
after observation. No later task or publication is started.

Dirty preparation: the full Unit tier passes 539 checks in each actual local
PS5.1.26100.9444 and pinned PS7.6.6 host, every bad count zero and unchanged
source guards. All 63 maintained PowerShell files parse and pass the 41
selected analyzer rules in both hosts, without suppressions. Vendor advisories
are 0 errors / 321 warnings / 172 information per host and are nonblocking.
These are preparation results, not clean-commit or hosted acceptance.

Focused regressions cover 63 report/export cases, 45 dependency/path/hash
cases and seven static/workflow cases per host. Real cached ZIP extraction
and archive listing rehearsal passes without new local downloads. Preparatory
schema/array/XML and PS5.1 fixture-generation issues were corrected and rerun;
no skipped case was relabeled a pass. Independent review found and resolved
native bootstrap's optional analyzer path, runner-label consistency and empty
PS5.1 native argv delivery. Selected-shell startup reconstructs the default
module path; both actual hosts were probed successfully. The seven affected
root regressions pass again after the final invocation change.

C1 will bind the implementation. Hosted pass, PR execution and deliberate
failure remain required before AC054/AC055 can be accepted.

First pushed C1 `7f3426deac5834c9e070855f65eaab02059aa0bf` was clean/live
synchronized. Draft PR24 was created and attached. Hosted push run
`37808521621` failed before creating any job or artifact; it is not test
evidence. Investigation found `runner.temp` in job-level `env`, where GitHub's
context-availability contract excludes `runner`. The setting is moved into
step-level `env`, where it is supported, with a workflow regression assertion.
The next clean implementation checkpoint and hosted rerun will supersede this
failed preparation, retaining its actual run ID and zero executed jobs.

Corrected C1a `4eb308236ebe25e62225ebb76f99642d961548e1` was clean/live
synchronized. Push run `37808743156` and PR run `37808743041` each executed
four jobs and uploaded four sanitized artifacts. Both workflows failed their
underlying tests. Eight downloaded artifact ZIP hashes independently match
GitHub API digests; no result is relabeled passed.

The local clean driver reproduced PS5.1 Launcher's 18 failures: an explicitly
empty child module path could not survive a test's unset/restore roundtrip.
The selected driver already starts with its runtime's reconstructed compatible
default path. Tier children now inherit that nonempty selected-shell path.
A workflow/driver regression guards the boundary. PublicDocs case13 fails on
hosted checkout's license EOL rewrite; a per-command autocrlf checkout probe
reproduces the changed SHA without altering the preserved local license.
An explicit LICENSE -text attribute retains its original tested bytes.

Each native job had NativeFixture1 and SourceDiscovery4 passes, but case12 of
each 13-case backend path suite failed (current-user ReadData ACL denial).
Those standard-user/path characterizations remain required in their local
native scope and retain T23 evidence. Hosted Server/admin CI now uses a
dedicated genuine-engine merge/preset/no-overwrite smoke suite, rather than
claiming those desktop/ACL checks. The original failed counts and scoped
replacement are disclosed. Normal green and deliberate failure runs against
the final implementation remain pending.

The focused genuine-engine CiNativeSmoke tier passes four checks per actual
local shell with unchanged source guards in dirty preparation. Its two changed
PowerShell files parse and pass the selected 41 analyzer rules per host with no
findings/suppressions. Seven root CI regressions pass again, and the autocrlf
checkout probe now retains the exact original license SHA. Final C1b will rerun
the complete selected CI driver against clean source in both hosts and on GitHub.
