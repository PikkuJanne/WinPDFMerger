"""Record the six contact images actually inspected with original-detail view_image."""
import hashlib
import json
from pathlib import Path

root = Path(__file__).resolve().parent
visual = root / 'visual-published-outputs'
sha = lambda raw: hashlib.sha256(raw).hexdigest()
render_path = visual / 'render-contact-report.json'
render_raw = render_path.read_bytes()
if sha(render_raw) != 'd82d1ab887b13ee1ca791e28f325328da68c3da024fdc197cc2d3a0c5017834c':
    raise ValueError('Original actual render report changed')
sheets = sorted(visual.glob('contact-*.png'))
if len(sheets) != 6:
    raise ValueError('Exactly the six actually viewed contact images required')
report = {'task': 'T33', 'result': 'pass_for_recorded_rendered_contact_sheet_inspection',
          'source_commit': '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a',
          'harness_commit': 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232',
          'zip_sha256': '2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2',
          'checksums_sha256': 'd39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca',
          'original_render_report_sha256': sha(render_raw), 'pdf_count': 21, 'page_count': 106, 'contact_sheets': 6,
          'method': 'Independent agent actually viewed contact01 through06 with functions.view_image detail=original; this record preserves those observed tool views',
          'observations': ['Visible page identifiers/order, blue frames, text and margins present across all21PDF106page renders',
                           'No unexpected blank pages, page clipping, overlap, missing page frames or raster loss observed',
                           'Raster noise source and email renditions visible in expected sixth pages; vector tiny masters retain two expected pages'],
          'sheets': [{'path': path.relative_to(root).as_posix(), 'bytes': path.stat().st_size, 'sha256': sha(path.read_bytes())} for path in sheets],
          'human_acceptance': 'excluded/unperformed; never pass', 'application_native_or_download_reexecuted': False,
          'limitations': ['Agent render inspection is not a physical human PDF-viewer walkthrough',
                          'No universal fidelity, PDF/A, signature validity, malicious-document isolation or archival guarantee'], 'issues': []}
destination = root / 'visual-published-review.json'
if destination.exists():
    raise ValueError('Preserve the original visual record')
destination.write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'result': report['result'], 'report': str(destination), 'sha256': sha(destination.read_bytes())}))
