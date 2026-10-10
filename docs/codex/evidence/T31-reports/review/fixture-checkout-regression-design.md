# T31 fixture checkout regression design

Test the behavior that failed: raw byte-pinned fixture files must survive a real Git add and fresh checkout under command-scoped `core.autocrlf=true` and `false`.

Create an owned temporary seed repository containing `.gitattributes`, the strict corpus loader/catalog, every pinned recipe and manifest, and the numbered PDF fixtures. Add and commit using ordinary Windows newline conversion (`core.autocrlf=true`), then make two real local clones with `--no-local`, one under each autocrlf setting. In each clone, run the strict catalog loader so every pinned recipe/manifest is hash-checked; also compare every committed numbered output, including the manifest, to independently regenerated expected bytes. Each setting is a separate regression case.

All Git settings belong to individual commands and the owned temporary repositories. Disable external attributes, hooks and templates in those commands; set only synthetic local commit identity and signing configuration. Invoke Git directly with `shell=False`, capture stderr and bound subprocess time. No repository/global Git configuration change, download, PDF upload, dependency installation or application orchestration is needed.

`tools/test/tests/test_fixture_checkout.py` implements this design. The independent read confirms that real Git materialization and the strict catalog are exercised, rather than a mocked newline conversion. The root's recorded preparation results are red before the fix for both cases and green after the fixture correction, with 42 fixture-suite passes. Those dirty-branch preparation outcomes are separate from the required fresh clean merged-R2 executions. The test does not claim Windows native engine, account-class, Explorer or package/download acceptance.
