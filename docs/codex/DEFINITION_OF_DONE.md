# Definition of done

**The project is complete only when `v1.0.0` is published on GitHub and its exact download has been verified and smoke-tested.** Source-ready, merge-ready, a local build, a pushed tag, a CI artifact, or a draft release is not complete.

## Product
The existing local folder workflow and entry points remain. All 17 accepted improvements have implementation/evidence coverage. Source hashes remain unchanged; no source/existing output overwrites; deterministic order; bounded correctly quoted native execution; parse/page validation; correct master/email/partial states; preserved defaults and small parameters; honest limits/help/privacy. No rewrite, replacement engine, website, server-side conversion, runtime Python/network/telemetry, OCR/editor/installer/auto-updater.

## Demonstrated quality
Required unit/fault/native/CI/package and real Windows desktop checks pass. Windows PowerShell 5.1 and one recorded supported PowerShell 7 x64 environment are exercised with real dependencies. The expected source/output/page order/fidelity checks are observed, not inferred from file existence. No known blocker is concealed; skipped or excluded cases are never counted as passed. Windows 10/live UNC/other architectures may be explicitly excluded only with accurate public documentation.

## Release provenance
The accepted merge commit R is known; final package was built/tested from clean R using a reviewed allowlist. Version and BUILD_INFO match R/1.0.0. Published assets contain no private documents, developer files, vendor executables or unrelated outputs. Exact ZIP and checksums-file SHA-256 values were recorded before publication. Script signing is optional and unsigned status is truthfully disclosed.

## Actual GitHub state
Exactly one public GitHub Release exists, tagged `v1.0.0`, not draft or prerelease, with real public URL/time and exactly the expected ZIP/checksum assets. The live annotated tag peels to R. No intermediate published release/tag was created for development progress. Public assets downloaded without the authenticated draft cache match the independently recorded accepted hashes and manifest. A fresh extraction of that downloaded ZIP passes the Windows standard-user smoke workflow.

## Continuity and closure
All 34 tasks and required acceptance cases have real evidence; scoped exclusions have rationale. RELEASE_STATE records verified publication. The evidence-only PR is merged normally, final local main is clean and equals the live remote main, and E descends from R with only `docs/codex/` changes after the release-source freeze. Tag remains at R. Final report includes release URL, R, hashes, verified downloaded operation, E/live-sync proof, and limitations. Do not invent future evidence or require an impossible self-referential commit hash.

A missing credential, protected-branch approval, required human observation, failed test, wrong release artifact, network outage or failed push is a **specific blocker**, not a completed project. Preserve work and record the exact next action rather than announcing "done" or asking again for permission already given.
