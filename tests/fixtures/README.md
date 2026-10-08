# Original synthetic regression corpus

The tracked [corpus catalog](corpus.json) consolidates the earlier numbered,
envelope, preset and feature fixtures with T21 source-safety scenarios. It binds
the recipes, development pins, original hashes, page identifiers, explicit merge
order and preservation expectations. All text, vector marks, seeded scan pixels,
forms and inert attachment bytes are original synthetic content. Licenses are
recorded per group (repository LICENSE or CC0-1.0). There are no private PDFs,
personal document names, external images or embedded font files.

| Group | Reconstruction and expected result |
| --- | --- |
| `numbered` | Tracked `1.pdf`, `2.pdf`, `10.pdf`: four pages in natural order, visible `T03-xx-Pxx` identifiers, 432 by 288 points. |
| `envelopes` | Original manual objects with LF/CR, incremental update, xref stream and an actual approved Ghostscript linearized derivative: five one-page inputs. Explicit native order is conventional CR, conventional LF, incremental, linearized, xref stream. |
| `presets` | Seeded scan, small vector print and mixed vector/raster originals: four pages in `mixed.pdf`, `scan.pdf`, `small-print.pdf` order. The scan pixels are generated, not external artwork or OCR. |
| `features` | Two two-page originals: links/destinations, bookmarks, forms with repeated field names and distinct values, widgets, annotations, rotations, document/page attachments and minimal tags. Four identifiers in A, B order. |
| `safety/matrix` | Eighteen included files, nineteen pages: zero/leading-zero numbers, integers beyond Int32, multiple numeric segments, uppercase extension, spaces/brackets/`!`/`&`/parentheses/apostrophe, short Latin Unicode names and a legitimate `WinPDFMerge_` name. Real hidden PDF, nested PDF, non-PDF file and PDF-named directory are explicit exclusions. |
| Other `safety` scenarios | One uppercase single-page input; CJK path rejection on the pinned reference backend; valid siblings around empty, truncated, malformed, user-encrypted and owner-restricted PDFs. Invalid inputs require failure with no silently omitted operand or final master. |

The natural order and page identifiers are explicit data, independent of the
application comparator. Envelope copies can intentionally repeat the same
original identifier. Hashes cover hidden/nested/non-PDF sources as well as
included inputs. The native safety suite checks actual application output,
repeat runs under two cultures, full source-tree preservation, overlap refusal,
existing output preservation and concurrent run isolation.
Snapshots compare every file's bytes/hash, length, attributes, creation and
modified timestamps. They retain directory presence, contents, attributes and
creation timestamps. Directory modified timestamps are excluded: one prepared
foreign directory changed that metadata before application invocation on the
reference host. File invariants and unexpected directory/file detection remain
strict; this does not claim preservation of every filesystem metadata field.

Use the configured bundled development Python **3.12.14** with ReportLab
**4.4.9**, pypdf **6.10.0**, Pillow **12.3.0**, pypdfium2 **5.13.0** and native
PDFium **153.0.7999.0**. Library pins are in
[`tools/test/requirements-fixtures.txt`](../../tools/test/requirements-fixtures.txt).
Envelope reconstruction requires the explicitly selected approved Windows
Ghostscript **10.08.0** executable; the generator verifies the earlier acquisition
receipt and console/DLL hashes before launching it. No tool downloads or installs
dependencies. Python and generators are development-only and must stay outside
runtime packaging. WinPDFMerger has no runtime Python dependency.
The refreshed workspace bundle reports the same Python and PDFium versions
with different executable and DLL bytes. Test-only Python readers and the
feature oracle accept only the two exact observed hashes of each, record the
selected bytes, and retain version pins. Unknown DLL bytes remain a failure
even when version strings match; this is a development cache compatibility
allowlist, not a runtime installation or general trust of same-version binaries.

From the repository root, substitute the selected executable paths:

```text
python -B tools/test/corpus.py materialize --output tests/.work/corpus-1 --ghostscript <approved-gswin64c.exe>
python -B tools/test/corpus.py verify --root tests/.work/corpus-1
python -B -m unittest discover -s tools/test/tests -v
```

`materialize` creates only a new explicitly owned directory below `tests/.work`
and refuses an existing output. It reconstructs all four groups and safety
scenarios, sets the actual Windows hidden attribute, then verifies them. The
generated `corpus.json` freezes the complete file/directory inventory, byte
lengths, SHA-256 hashes and hidden attributes. It carries `safety.scenarios`
with source directories, source rows, ordered names, expected page identifiers,
page totals and exit codes. `verify` reads only, rejects missing/unexpected or
changed files/attributes, and independently reads valid page IDs, dimensions
and rotations through native PDFium. It never repairs a fixture or invokes the
application. Generated PDFs, receipts and renderings stay ignored.

Reconstruct again into a fresh `tests/.work/corpus-2` directory to compare
original bytes. Numbered, preset, feature and safety sources reproduce exactly
under the pins, including deterministic RC4-128 rejection samples with fixed
synthetic IDs. Synthetic credentials are `T21-user` and `T21-owner`; the
owner-restricted sample has an empty user password. These are test data, not
secrets or production encryption advice. Ghostscript's linearized derivative
may vary in native date/ID bytes: its receipt binds the actual hash while the
tracked recipe requires its linearization dictionary, page count and visible
identifier. It is not claimed to be byte-for-byte reproducible.

To inspect a real master independently, write its expected identifiers as a
JSON array in an owned development file:

```text
python -B tools/test/corpus.py inspect --pdf <master.pdf> --expected-identifiers <expected-identifiers.json>
```

This accepts repeated original identifiers while requiring exact count/order
and one expected visible identifier per page. It leaves the inspected bytes
unchanged. CLI JSON is ASCII-escaped so CJK synthetic paths remain usable under
legacy Windows console encodings.

The earlier focused commands remain available:

```text
python -B tools/test/generate_numbered_fixtures.py
python -B tools/test/fixture_oracle.py
python -B tools/test/fixture_oracle.py --render-dir tests/.work/numbered-renders
python -B tests/fixtures/features/generate_features.py --output tests/.work/features-1
python -B tools/test/feature_oracle.py --pdf tests/.work/features-1/1-feature-A.pdf --output tests/.work/features-A.json --render-dir tests/.work/features-A-renders
```

The numbered generator defaults to read-only exact comparison; `--write` is
explicit regeneration. The feature oracle reads raw structures with pypdf and
uses PDFium for page IDs/rendering. It never follows links, opens attachments,
fills/repairs/flattens or rewrites inspected documents.

Master and email preservation expectations are separate. Historical T19 native
observations include renamed master form names, loss of the named-destination
index/document attachment index/tag relationships, retained resolved navigation
and page attachments, lost editable email fields/widgets, and changed screen
orientation. Seeded nonpainting page-stream comments exercise the smaller-email
publication rule; they do not measure representative compression or fidelity.
See [`docs/PDF_LIMITATIONS.md`](../../docs/PDF_LIMITATIONS.md) and the native
`PreservationNative`/`SizeReportingNative` suites for measured structures and
rendered appearances. Minimal tags do not certify accessibility. Signature,
XFA, PDF/A, malware and universal archival preservation remain unvalidated.

Corpus reconstruction/oracle checks and their Python regressions prove fixture
provenance and inspection only. They are recorded separately from real Windows
PDFtk/Ghostscript application integration, physical Explorer acceptance and
distribution checks; none substitutes for those release gates.
