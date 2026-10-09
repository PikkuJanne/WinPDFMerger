"""Independent read-only audit of T29 preparation projections and retained inputs."""
from __future__ import annotations
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, required=True)
    parser.add_argument('--report', type=Path, required=True)
    args = parser.parse_args()
    repo = args.repo.absolute()
    checks = []
    issues = []

    def check(label, condition):
        checks.append({'check': label, 'pass': bool(condition)})
        if not condition:
            issues.append(label)

    replacements = [(str(repo), '<REPO>'), (os.environ['USERPROFILE'], '<USERPROFILE>')]

    def projected(value):
        if isinstance(value, dict):
            return {key: projected(item) for key, item in value.items()}
        if isinstance(value, list):
            return [projected(item) for item in value]
        if type(value) is str:
            for prefix, token in replacements:
                value = value.replace(prefix, token).replace(prefix.replace('\\', '/'), token)
        return value

    def compare(expected, actual, label):
        check(label + ' type preserved', type(expected) is type(actual))
        if type(expected) is not type(actual):
            return
        if isinstance(expected, dict):
            check(label + ' keys preserved', expected.keys() == actual.keys())
            for key in expected.keys() & actual.keys():
                compare(expected[key], actual[key], label + '/' + key)
        elif isinstance(expected, list):
            check(label + ' length preserved', len(expected) == len(actual))
            for index, (left, right) in enumerate(zip(expected, actual)):
                compare(left, right, label + '/' + str(index))
        else:
            check(label + ' value matches declared projection', expected == actual)

    prep = repo / 'docs/codex/evidence/T29-reports/preparation'
    names = ['initial-context.json', 'initial-platform.json', 'initial-probe.json', 'initial-raster-probe.json']
    check('preparation contains exactly four projections', sorted(p.name for p in prep.iterdir()) == sorted(names))
    originals = {}
    receipts = []
    for name in names:
        path = prep / name
        public_raw = path.read_bytes()
        public = json.loads(public_raw.decode('utf-8-sig'))
        original_path = repo / public['original_ledger']
        raw = original_path.read_bytes()
        original = json.loads(raw.decode('utf-8-sig'))
        originals[name] = original
        check(name + ' original byte count/hash', public['original_bytes'] == len(raw) and public['original_sha256'] == sha(raw))
        check(name + ' exact projection envelope', set(public) == {'task', 'scope', 'original_ledger', 'original_bytes', 'original_sha256', 'redactions', 'payload'})
        check(name + ' preparation only declared scope', public['task'] == 'T29' and public['scope'] == 'initial context or dirty exploratory preparation; not accepted T29 operation')
        check(name + ' only declared path substitutions', public['redactions'] == ['<REPO>', '<USERPROFILE>'])
        compare(projected(original), public['payload'], name + '/payload')
        text = public_raw.decode('utf-8-sig').replace('\\\\', '\\').casefold()
        check(name + ' raw repository/user prefix absent', all(prefix.casefold() not in text and prefix.replace('\\', '/').casefold() not in text for prefix, _ in replacements))
        receipts.append({'public_path': path.relative_to(repo).as_posix(), 'public_bytes': len(public_raw), 'public_sha256': sha(public_raw), 'original_path': public['original_ledger'], 'original_bytes': len(raw), 'original_sha256': sha(raw)})

    context = originals['initial-context.json']
    c0 = '5f962aaab1dacf0401ded92916f5a1ca73095354'
    check('initial local observation C0 binding', context['head'] == c0 and context['approved_cache_payloads_rehashed'] == 348)
    for asset in context['assets']:
        directory = Path(asset['directory'])
        check(asset['label'] + ' exactly ZIP/checksum assets', sorted(p.name for p in directory.iterdir()) == ['SHA256SUMS.txt', 'WinPDFMerger-v1.0.0.zip'])
        for row in asset['files']:
            raw = (directory / row['name']).read_bytes()
            check(asset['label'] + '/' + row['name'] + ' retained exact bytes/hash', len(raw) == row['bytes'] and sha(raw) == row['sha256'])
    cache = json.loads((repo / 'docs/codex/evidence/T23-reports/context/T23-environment.json').read_text(encoding='utf-8-sig'))['approved_selected_files']
    check('approved original cache inventory exactly 348', len(cache) == 348)
    for index, row in enumerate(cache):
        actual = Path(row['path'].replace('<USERPROFILE>', os.environ['USERPROFILE']))
        check('approved cache file ' + str(index) + ' original SHA256', sha(actual.read_bytes()) == row['sha256'])

    platform = originals['initial-platform.json']
    check('initial platform observation C0 binding', platform['source_head'] == c0)
    for row in platform['invocations']:
        check('initial platform ' + row['label'] + ' actual exit0', type(row['exit_code']) is int and row['exit_code'] == 0)
        for stream in ('stdout', 'stderr'):
            raw = (repo / 'tests/.work/T29-context' / ('initial-' + row['label'] + '.' + stream)).read_bytes()
            check('initial platform ' + row['label'] + '/' + stream + ' raw byte hash', sha(raw) == row[stream + '_sha256'])
    check('initial platform PR26 draft/open C0', platform['pr']['number'] == 26 and platform['pr']['isDraft'] is True and platform['pr']['state'] == 'OPEN' and platform['pr']['headRefOid'] == c0)
    check('initial eight checks successful platform metadata only', len(platform['pr']['statusCheckRollup']) == 8 and all(row['conclusion'] == 'SUCCESS' and row['status'] == 'COMPLETED' for row in platform['pr']['statusCheckRollup']))
    check('initial no releases/tags and refs bound', platform['releases'] == [] and platform['refs'].splitlines() == [c0 + '\trefs/heads/codex/v1.0.0-readiness', 'e2451141217efdd00a1d49d72a04df054872dffc\trefs/heads/main'])

    observations = {}
    for name in ('initial-probe.json', 'initial-raster-probe.json'):
        rows = originals[name]
        by_label = {row['label']: row for row in rows}
        check(name + ' unique observed labels', len(by_label) == len(rows))
        check(name + ' native version then owned init failure', by_label['gs-lib-version']['exit_code'] == 0 and by_label['gs-lib-version']['stdout'].strip() == '10.08.0' and by_label['gs-lib-render']['exit_code'] == 1 and 'Initialization file gs_init.ps does not begin with an integer' in by_label['gs-lib-render']['stderr'])
        for label, code in [('candidate-default', 0), ('candidate-skip', 0), ('candidate-email-failure', 2)]:
            row = by_label[label]
            check(name + '/' + label + ' actual exploratory exit', row['exit_code'] == code)
            output = Path(row['command'][row['command'].index('-OutputFolder') + 1])
            for artifact in row['outputs']:
                raw = (output / artifact['name']).read_bytes()
                check(name + '/' + label + '/' + artifact['name'] + ' retained exact bytes/hash', len(raw) == artifact['bytes'] and sha(raw) == artifact['sha256'])
        check(name + ' exact provenance excludes acceptance', by_label['candidate-experiment-provenance']['scope'] == 'experiment only; not accepted clean execution')
        check(name + ' recorded source set unchanged', by_label['candidate-experiment-provenance']['source_unchanged'] is True and by_label['candidate-experiment-provenance']['before'] == by_label['candidate-experiment-provenance']['after'])
        observations[name] = [{'label': row['label'], 'exit_code': row.get('exit_code')} for row in rows]
    for label in ('candidate-raster-default', 'candidate-raster-ebook'):
        row = next(row for row in originals['initial-raster-probe.json'] if row['label'] == label)
        check(label + ' exploratory smaller email observed', row['exit_code'] == 0 and 'Email result: published' in row['stdout'])
    prep_text = (repo / 'docs/codex/evidence/T29-preparation.md').read_text(encoding='utf-8-sig')
    check('preparation note distinguishes init/config fault and pending clean acceptance', 'controlled initialization failure path' in prep_text and 'not accepted clean T29' in prep_text and 'AC058 remains excluded/unperformed' in prep_text)
    attrs = (repo / '.gitattributes').read_text(encoding='utf-8-sig')
    check('T29 attribute scoped byte preservation only', '/docs/codex/evidence/T29-reports/** -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol' in attrs)
    result = {'task': 'T29', 'audit': 'independent_preparation_projection_and_retained_hashes', 'acceptance_evidence': False, 'auditor_sha256': sha(Path(__file__).read_bytes()), 'command': projected([sys.executable, '-B', str(Path(__file__).absolute()), *sys.argv[1:]]), 'checks': len(checks), 'issues': issues, 'receipts': receipts, 'observations': observations, 'details': checks}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, indent=2) + '\n', encoding='utf-8', newline='\n')
    print(json.dumps({'checks': len(checks), 'issues': len(issues), 'report': projected(str(args.report))}))
    return bool(issues)


if __name__ == '__main__':
    raise SystemExit(main())
