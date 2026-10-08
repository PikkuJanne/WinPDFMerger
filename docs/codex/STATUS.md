# Project status

Target: published and independently verified v1.0.0 in PikkuJanne/WinPDFMerger.
Completed milestones: M1/M2/M3 within recorded Windows/test/review scope.
Current milestone: M4, in_progress. Completed tasks: T01 through T24.
Current task: T25 - focused security and public-repository review, pending.
Publication: NOT STARTED. No release or tag is created.

T24 AC054/AC055 pass at clean C1b
`e626e45a5ba375456f23b506f0ded7ca7d68f1e3`. Actual Windows hosted push/PR
runs each pass four jobs / 1306 checks / 18 sanitized report pairs. Actual PR
synthetic merge has the same Git tree as C1b. Same-source negative probe retains
one real failed assertion per unit host and fails the workflow; both native
jobs pass and all four artifacts upload. Full-SHA Actions, verified temporary
vendor dependencies and actual minimum token permissions pass independent
review; routine CI has no release publishing capability.

Clean local selected driver passes 1306 checks / 18 pairs in actual
PS5.1.26100.9444 and pinned PS7.6.6 under a nonadmin token. Source guards pass.
All 64 maintained PowerShell files pass 41 selected rules per host, zero
findings/suppressions; advisories 0 errors / 325 warnings / 175 information per
host remain nonblocking. Runtime/launchers/native flags/defaults unchanged.

Independent final hosted/local audits pass 28950/5511 integrity checks;
all 20 downloaded ZIP digests, including failed preparation, match GitHub API.
Distinct controlled/unit and real native-smoke classes and truthful counts
remain retained. Earlier parser/module-path/license/hosted ACL failures and
corrected scope are disclosed. See T24 completion/results/manifest/reviews.

C1b normal push/fresh clean live equality is retained; PR24 remains draft and
unmerged. Records-only C2 own clean/live equality and current PR head are checked
after push in session without self-reference. T25 next. Hosted Server/admin CI
is not standard-user ACL, physical Explorer or manual desktop acceptance.
T23 broader native evidence and T19 preservation limits remain within their
scope. OS support channel, security, broader OS/UNC, desktop/manual, package
and publication gates remain; T24 closes none of those later gates.
