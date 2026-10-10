"""Narrow T33 operation auditor derivative; no operation or reader execution."""
import ast
import difflib
import hashlib
import json
from pathlib import Path

root = Path(__file__).resolve().parent
repo = root.parents[2]
base = (repo / 'tests/.work/T32-review/audit_final_operation.py').read_bytes()
sha = lambda raw: hashlib.sha256(raw).hexdigest()
if sha(base) != '6e926e44d99df7839a568589f9d22bda903afcee6c80b38a572908d26c2b476f':
    raise ValueError('Frozen independent T32 operation auditor required')
text = base.decode()
changes = [
    ('T32', 'T33'),
    ('02811fe05e6fbc5955634fc5ed62d4f3650834bd0fce60c4c62c806093b8557a', '01cf9fa377dde5719fea16153791bc2a85fa412872ceee9cfa6029a5f270c8b7'),
    ('080016af3ec0e579d648d75d584306d9d7e25a854079c1ce1a5df6bd90ce99e5', '5ea6bd48c400f1ffe9becf7ba5fb2b7fd355019dbbab3a8aa295a29c727626f8'),
    ('actual_final_R_package_operation', 'actual_independently_anonymously_downloaded_published_v1_0_0_package_operation'),
    ('ignored T33-review', 'ignored T33-public-download-review'),
    ('independent_original_final_R_receipts_exact_shared_assets_and_final_pdf_operation', 'independent_original_anonymous_published_download_receipts_exact_shared_assets_and_pdf_operation')]
for old, new in changes:
    if old not in text:
        raise ValueError('Expected narrow derivation source absent: ' + old)
    text = text.replace(old, new)
addition = '''
def published_download_matches_ledger(public, ledger, zip_hash, sums_hash):
    directory = Path(public['download_directory']).absolute()
    return (public['task'] == 'T33' and public['result'] == 'pass_for_unauthenticated_published_release_and_independent_download' and
            public['issues'] == [] and public['source_commit'] == SOURCE and public['harness_commit'] == ledger['harness_commit'] and
            public['authentication_used'] is False and public['cookies_used'] is False and public['gh_download_used'] is False and
            public['download_directory_previously_existed'] is False and public['application_executed'] is False and public['remote_mutations'] is False and
            public['draft'] is False and public['prerelease'] is False and public['published_at'] == '2026-10-10T07:12:01Z' and
            public['release_id'] == 408603768 and public['tag_object_sha'] == '7818645de07b902ad8f2b815e90ee1d74d2724d6' and
            public['source_sha256'] == 'f0f3bcba724e99c249787b01092d111c4085e296aee43d1403f3e7eaee9df6dc' and
            public['package_auditor_sha256'] == '85fb65be23bffd953f76fc4a52b6ac84bf3d89400a0cbdd75bcd2379a20bec34' and
            public['zip_sha256'] == ledger['shared_assets']['zip_sha256'] == zip_hash and
            public['checksums_sha256'] == ledger['shared_assets']['checksums_sha256'] == sums_hash and
            public['package_audit']['checks'] == 266 and public['package_audit']['issues'] == [] and
            Path(ledger['shared_assets']['zip_path']).absolute() == directory / 'WinPDFMerger-v1.0.0.zip' and
            Path(ledger['shared_assets']['checksums_path']).absolute() == directory / 'SHA256SUMS.txt')

'''
marker = 'def main():\r\n' if '\r\n' in text else 'def main():\n'
text = text.replace(marker, addition + marker, 1)
parser_marker = "    parser.add_argument('--capture', type=Path, required=True)"
text = text.replace(parser_marker, parser_marker + "\n    parser.add_argument('--public-download-report', type=Path, required=True)\n    parser.add_argument('--public-download-report-sha256', required=True)", 1)
binding = '''
        check('external accepted anonymous download report pin', args.public_download_report_sha256 == '0f5f94923e47d0ae5bcc247df9d5ba56e02b39b1513977b32b27b3c1727faa19')
        public = load(args.public_download_report)
        bind(args.public_download_report, args.public_download_report_sha256, 'actual independent public download report')
        check('actual independently downloaded public pair paths and scopes', published_download_matches_ledger(public, ledger, args.expected_zip_sha256, args.expected_checksums_sha256))
        for command in public['commands']:
            check('public original ' + command['label'] + ' actual exit', command['exit_code'] == 0)
            for name, stream in command['streams'].items():
                raw = bind(repo / stream['path'], stream['sha256'], command['label'] + '/' + name)
                check('public original command stream byte length', len(raw) == stream['bytes'])
        for request in public['public_http_requests']:
            raw = bind(repo / request['response_path'], request['sha256'], 'actual public API response')
            check('actual public API unauthenticated response facts', request['method'] == 'GET' and request['status'] == 200 and request['authentication'] is False and request['cookies'] is False and len(raw) == request['bytes'] and not any(key.lower() in ('authorization', 'cookie') for key in request['request_headers']))
        package_audit = load(repo / public['package_audit']['path'])
        bind(repo / public['package_audit']['path'], public['package_audit']['sha256'], 'actual published downloaded package audit')
        check('actual original public download byte audit complete', package_audit['task'] == 'T33' and package_audit['source_commit'] == SOURCE and package_audit['checks_total'] == 266 and package_audit['issues'] == [] and all(item['pass'] is True for item in package_audit['checks']))
'''
binding_marker = "        bind(capture / 'invocations.json')"
text = text.replace(binding_marker, binding_marker + '\n' + binding, 1)
out = root / 'audit_published_operation.py'
raw = text.encode()
ast.parse(text)
out.write_bytes(raw)
delta = ''.join(difflib.unified_diff(base.decode().splitlines(keepends=True), text.splitlines(keepends=True), fromfile='frozen-T32/audit_final_operation.py', tofile='prepared-T33/audit_published_operation.py'))
(root / 'operation-auditor-derivation.diff.txt').write_text(delta, encoding='utf-8', newline='\n')
metadata = {'task': 'T33', 'result': 'prepared_only', 'base_sha256': sha(base), 'derived_sha256': sha(raw), 'delta_sha256': sha(delta.encode()),
            'changes': changes, 'added_binding_scope': 'Original actual anonymous download hash, API/command raw bytes, strict no-auth/cookie flags, exact same original directory asset paths and published metadata; existing native/PDF/source/canary guards preserved',
            'application_or_native_or_reader_executed': False, 'frozen_T32_source_modified': False}
(root / 'operation-auditor-derivation.json').write_text(json.dumps(metadata, indent=2) + '\n', encoding='utf-8')
print(json.dumps(metadata))
