# T31 corrective PR27 CI review

PR27 CI run [37971199628](https://github.com/PikkuJanne/WinPDFMerger/actions/runs/37971199628) passes its premerge scope: four successful jobs, 20 original JSON/NUnit pairs and 1,370 passes. Unit scopes contribute 676 per shell; native smoke contributes 9 per shell. The 1,851 independent metadata/receipt/hash/count checks found no issue.

The original GitHub API trigger/head is reviewed C1 `30560516a0248636769e988b0420466214c25e3b`. All artifact execution commits equal synthetic checkout `0d1a1b39ec8c15378d63f3c1d4da237f1240878d`. Its API parents are observed main `de5f30155c68755dbd5af691625a0651e3fb7230` and C1. The synthetic checkout and C1 have exactly the same API tree `5014f5bdf4f374aee828ced4c39cb93bfeb6465a`. Trigger and checkout commit identities remain separate.

Both hosted maintained-source static receipts pass 68 files with zero selected findings/suppressions or other bad counts; advisory totals remain 349 warnings and 175 information each. Actual hosts are Windows Server, PS5.1.26100.33438 and pinned PS7.6.6 with administrator tokens. Controlled/unit/documentation/static and native smoke classes remain separate; this is no human standard-user or Explorer evidence.

The original download has 50 already-sanitized artifact files. Captured read/download/API argv, timestamps, process exits, raw stdout/stderr, commit objects and a byte/hash index are preserved alongside `review.json`. This reviewer ran no application/test and made no source/Git configuration change. Hosted dependency payloads were not independently rehashed by downloading exporter receipts.

This is a premerge PR27 CI gate result. A new normal merge and fresh exact merged-source full dual-shell, static, supplementary and CI receipts remain required for AC071/AC072. AC058 remains owner-excluded/unperformed; no release-source acceptance, tag, package/download or publication completion is claimed.
