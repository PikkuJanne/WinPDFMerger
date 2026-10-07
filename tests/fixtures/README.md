# T03 numbered fixtures

`numbered/1.pdf`, `2.pdf` and `10.pdf` are original synthetic PDFs with four
pages total. Every page has one large, visible `T03-xx-Pxx` identifier.
`numbered/manifest.json` records input hashes, lengths, page totals, dimensions
and the expected `1, 2, 10` merge order. The files contain only original text
and vector artwork, under the repository LICENSE. They contain no user PDFs,
external images or embedded font files.

The generator uses ReportLab 4.4.9 in invariant mode, with PDF 1.4 and
uncompressed page streams. The independent oracle uses the native PDFium
153.0.7999.0 library through pypdfium2 5.13.0 to read page count, dimensions
and identifiers in page order. It can render every page to PNG for visual
review. Versions are pinned in `tools/test/requirements-fixtures.txt`; this is
development tooling, with no automatic dependency installation or application
runtime dependency. Use the configured bundled workspace Python when available.

From the repository root, using that Python executable:

```text
python -B tools/test/generate_numbered_fixtures.py
python -B tools/test/fixture_oracle.py
python -B -m unittest discover -s tools/test/tests -p test_fixture_oracle.py -v
python -B tools/test/fixture_oracle.py --render-dir tests/.work/fixtures
```

Generation defaults to read-only byte-for-byte comparison. Explicitly pass
`--write` to regenerate the intended corpus, or `--output-dir <directory>
--write` to create a comparison copy. An actual PDFtk or Ghostscript result can
later be checked with `fixture_oracle.py --pdf <result.pdf>` against the expected
merged identifiers; the oracle does not create or merge that result.

These checks inspect real PDFs with an independent native parser. They prove
the tiny corpus and its oracle, and must be recorded separately from controlled
fake-process tests. They do not prove PDFtk/Ghostscript execution, product
merging, feature preservation or desktop acceptance. Mixed sizes, rotation,
images/scans, corrupt/encrypted and feature-rich PDFs belong to later tasks.
