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
images/scans and corrupt/encrypted PDFs belong to later tasks.

## T19 synthetic feature corpus

`features/generate_features.py` creates two original CC0 PDFs in an explicitly
selected development directory. `features/manifest.json` binds their exact
bytes and expected four-page order. The corpus covers links and internal
destinations, bookmarks, AcroForm values and repeated field names, two widgets
per field, annotations, rotations, document and page attachments, and minimal
tag relationships. It contains no user documents or external artwork.

```text
python -B tests/fixtures/features/generate_features.py --output tests/.work/features
python -B tools/test/feature_oracle.py --pdf tests/.work/features/1-feature-A.pdf --output tests/.work/features-A.json --render-dir tests/.work/features-A-renders
python -B -m unittest discover -s tools/test/tests -p test_feature_oracle.py -v
```

Use the pinned Python 3.12.14 and development dependencies recorded in the
manifest and `tools/test/requirements-fixtures.txt`. These tools neither install
dependencies nor participate in the application runtime. The read-only oracle
uses pypdf for raw document structures and native PDFium for identifiers and
rendered pages. It does not merge, repair, fill, flatten or rewrite tested PDFs.

`PreservationNative` runs the unchanged application with explicitly selected
PDFtk and Ghostscript paths and records originals, masters, screen copies and
ebook copies separately. Artificial, nonpainting page-stream comments make the
synthetic inputs large enough for the normal smaller-email publication rule;
these sizes do not measure representative compression or image fidelity.

Tag relationships do not certify accessibility. Signatures and XFA have no
validated preservation guarantee in this corpus. See `docs/PDF_LIMITATIONS.md`
for measured native results and limits. Generated PDFs and rendered QA images
remain in ignored development directories; public evidence retains text and
hashes without private PDFs.
