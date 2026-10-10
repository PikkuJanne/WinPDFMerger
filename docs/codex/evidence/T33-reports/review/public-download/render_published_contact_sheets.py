"""Read retained final outputs and render visual contacts; no app/native PDF engine run."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from PIL import Image, ImageDraw

sha = lambda raw: hashlib.sha256(raw).hexdigest()
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--capture', type=Path, required=True)
parser.add_argument('--output', type=Path, required=True)
args = parser.parse_args()
root = Path(__file__).resolve().parent
if not args.output.resolve().is_relative_to(root) or args.output.exists():
    raise ValueError('Use a NEW visual directory under ignored T33-public-download-review')
args.output.mkdir()
poppler = Path('<USERPROFILE>/.cache/codex-runtimes/codex-primary-runtime/dependencies/native/poppler/Library/bin/pdftoppm.exe')
version = subprocess.run([str(poppler), '-v'], capture_output=True, timeout=20)
rows, calls, sheets = [], [], []
for shell in ('PS51', 'PS7'):
    result = json.loads((args.capture / (shell + '-reports/result.json')).read_text(encoding='utf-8-sig'))
    if result['task'] != 'T33' or result['result'] != 'pass' or result['preparation'] or result['candidate_source_commit'] != '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a':
        raise AssertionError('Only actual complete final-R retained outputs')
    for case in result['cases']:
        outputs = [Path(path) for path in case['output_paths'] if Path(path).suffix.lower() == '.pdf']
        if len(outputs) != len(case['independent_pdfs']):
            raise AssertionError('Exact retained output receipt inventory')
        for output, receipt in zip(outputs, case['independent_pdfs']):
            raw = output.read_bytes()
            if sha(raw) != receipt['sha256'] or len(raw) != receipt['bytes']:
                raise AssertionError('Original retained PDF changed')
            label = shell + '-' + case['label'] + ('-email' if output.name.endswith('_email.pdf') else '-master')
            prefix = args.output / label
            argv = [str(poppler), '-r', '72', '-f', '1', '-l', str(receipt['strict_pypdf_pages']), '-png', str(output), str(prefix)]
            run = subprocess.run(argv, capture_output=True, timeout=40)
            (args.output / (label + '.stdout.txt')).write_bytes(run.stdout)
            (args.output / (label + '.stderr.txt')).write_bytes(run.stderr)
            calls.append({'label': label, 'argv': argv, 'exit_code': run.returncode, 'stdout_sha256': sha(run.stdout), 'stderr_sha256': sha(run.stderr)})
            if run.returncode != 0:
                raise AssertionError('Poppler render failed: ' + label)
            pages = sorted(args.output.glob(label + '-*.png'), key=lambda path: int(path.stem.rsplit('-', 1)[1]))
            if len(pages) != receipt['strict_pypdf_pages'] or output.read_bytes() != raw:
                raise AssertionError('Complete renders and unchanged original PDF')
            rows.append({'label': label, 'pdf_bytes': len(raw), 'pdf_sha256': sha(raw), 'pages': [{'path': page.name, 'sha256': sha(page.read_bytes()), 'identifier': receipt['pages'][index]['identifier']} for index, page in enumerate(pages)]})
for index in range(0, len(rows), 4):
    selected = rows[index:index + 4]
    sheet = Image.new('RGB', (1328, 184 * len(selected) + 10), '#e5e7eb')
    draw = ImageDraw.Draw(sheet)
    for row_index, row in enumerate(selected):
        y = row_index * 184
        draw.text((10, y + 4), row['label'] + ' | original PDF SHA256 ' + row['pdf_sha256'][:16], fill='black')
        for page_index, page in enumerate(row['pages']):
            with Image.open(args.output / page['path']) as image:
                thumb = image.convert('RGB')
                thumb.thumbnail((216, 144))
                sheet.paste(thumb, (10 + page_index * 218, y + 22))
            draw.text((10 + page_index * 218, y + 168), str(page_index + 1) + ': ' + page['identifier'], fill='black')
    path = args.output / ('contact-' + str(index // 4 + 1).zfill(2) + '.png')
    sheet.save(path)
    sheets.append({'path': path.name, 'sha256': sha(path.read_bytes()), 'pdf_labels': [row['label'] for row in selected]})
if len(rows) != 21 or sum(len(row['pages']) for row in rows) != 106:
    raise AssertionError('Exact final retained PDF/page matrix')
report = {'task': 'T33', 'scope': 'Read-only Poppler render/contact preparation for visual QA, no application/PDFtk/Ghostscript/human pass', 'result': 'rendered_pending_visual_review', 'source_commit': '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a', 'capture': str(args.capture), 'renderer_path': str(poppler), 'renderer_sha256': sha(poppler.read_bytes()), 'renderer_version_stdout': version.stdout.decode(errors='replace'), 'renderer_version_stderr': version.stderr.decode(errors='replace'), 'pdf_count': len(rows), 'page_count': 106, 'sheets': sheets, 'files': rows, 'invocations': calls, 'script_sha256': sha(Path(__file__).read_bytes()), 'command': [sys.executable, '-B', str(Path(__file__).resolve()), *sys.argv[1:]]}
(args.output / 'render-contact-report.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps({'PDFs': len(rows), 'pages': 106, 'sheets': len(sheets), 'report': str(args.output / 'render-contact-report.json')}))
