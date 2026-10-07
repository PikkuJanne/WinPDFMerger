# Project status

Target: published and independently verified `v1.0.0` in `PikkuJanne/WinPDFMerger`.
Completed milestone: M1 — paths, ordering and native execution.
Completed tasks: T01 through T11.
Next task: T12 — no-overwrite staging and publication, in M2.
Publication: NOT STARTED.

T11 and required AC025/AC026 pass at clean implementation C1 `822f84eb4f7f0d75736887d9d79d0d363ae551a7`.
All nine focused tiers pass in each actual PS5.1.26100.9444 / supported PS7.6.6:
287 each,574 total, all failure/skip/not-run counts zero. A successful merge requires
every frozen input to pass the envelope check and bounded read-only PDFtk inspection,
with strict positive labeled Int64 count, guarded
expected total, named whole-job refusal and fresh length/UTCmtime checks around
inspection and before merging. Natural order and defaults remain preserved.

A small read-only byte-envelope guard rejects measured backend-recovered damage.
Actual corrected CR structural/footer, incremental, xref-stream and GS-linearized
controls pass; independent PDFium checks prove narrow visible page/order results.
Neither envelope nor native inspection is universal PDF validation, full fidelity,
encryption detection or a transactional snapshot. Readable empty-password
encryption can remain supported. Helpers remain PS5.1-compatible/import-only.

Original dirty native9/4 logger-return failures were fixed with a regression;
historical failures and invalid CR-stream tolerance controls remain separate.
Static analysis0errors/31warnings/6information per shell is retained and reviewed
nonblocking, not lint-clean/full T22. See `evidence/T11-completion.md`, exact clean
reports/results/review/live receipts and separate historical checkpoint evidence.

Normal C1 push and fresh clean/local/live equality verified; draft [PR11](https://github.com/PikkuJanne/WinPDFMerger/pull/11)
is open after owner-merged PR10/main f790c51. Final C2 contains records only;
its own post-push SHA/equality is reported in the session without self-reference.
Prior authorized external caches were reused; no acquisition/admin/PATH/persistent
policy/security change. Ordinary Restricted/all-Undefined policy remains observed;
RemoteSigned is test-child-only. Evidence is local NTFS Windows11x64/build26300
standard-user synthetic work. T12 staging, T13 master validation, T14 email/outcomes,
T15 interruption and OS support/UNC/Explorer/fidelity/CI/package/release gates remain.
