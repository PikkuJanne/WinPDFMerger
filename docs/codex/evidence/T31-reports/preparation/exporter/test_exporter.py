"""Synthetic development-tool checks only; no T31/native/CI acceptance claims."""
from pathlib import Path
import importlib.util, json, os, tempfile, unittest, xml.etree.ElementTree as ET

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[2]
spec = importlib.util.spec_from_file_location('t31_exporter', HERE/'Export-T31.py')
exporter = importlib.util.module_from_spec(spec); spec.loader.exec_module(exporter)
R2 = 'a'*40; INITIAL = 'de5f30155c68755dbd5af691625a0651e3fb7230'

def save(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False)+'\n', encoding='utf-8', newline='\n')

class ExporterChecks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Path(tempfile.mkdtemp(prefix='synthetic-tool-checks-', dir=HERE))
        cls.raw = cls.root/'raw'; cls.destination = cls.root/'public'; roots = []
        profile = os.environ['USERPROFILE']; cls.profile = profile
        cls.semantic = {'true': True, 'false': False, 'integer': 17, 'decimal': 1.25,
                        'null': None, 'array': [False, 0, None], 'path': profile+'\\synthetic-only'}
        cls.observation = cls.root/'receipt-owned-by-run'
        save(cls.observation/'semantics.json', cls.semantic)
        (cls.observation/'local-only.pdf').write_bytes(b'%PDF-synthetic-not-public')
        for shell in ('ps51', 'ps7'):
            path = cls.raw/shell; runs = []
            for index in range(32):
                tier = 'tier'+str(index)
                summary = {'result': 'pass', 'commit_under_test': R2, 'dirty_worktree': False, 'source_unchanged': True,
                           'total': 1, 'passed': 1, 'failed': 0}
                save(path/(tier+'.summary.json'), summary)
                xml = '\ufeff<test-results><environment user="'+Path(profile).name+'" machine-name="'+os.environ.get('COMPUTERNAME', 'SYNTHETIC')+'" user-domain="SYNTHETIC" cwd="'+str(REPO)+'"/><test-case name="'+str(REPO)+'\\synthetic" result="Success"/></test-results>'
                (path/(tier+'.results.xml')).write_bytes(xml.encode('utf-8'))
                runs.append({'tier': tier, 'summary': summary, 'exit_code': 0, 'process_error': None,
                             'observation_receipts': [('inline', json.dumps(cls.semantic)), ('owned-outside-named-root', str(cls.observation))] if index == 0 else []})
            save(path/'runs.json', runs)
            save(path/'aggregate.json', {'result': 'pass', 'commit_under_test': R2, 'dirty_worktree': False,
                                        'tiers': 32, 'bad_counts': 0, 'passed': 32, 'shell': shell})
            (path/'typed.json').write_bytes(b'\xef\xbb\xbf'+json.dumps(cls.semantic).encode('utf-8'))
            (path/'inline.stdout').write_bytes(('\ufeff'+json.dumps(cls.semantic)+'\r\n').encode('utf-8'))
            (path/'excluded.pdf').write_bytes(b'%PDF-synthetic-not-public')
            roots.append({'source': str(path), 'label': shell, 'mode': 'flat', 'role': 'accepted_full',
                          'shell': shell, 'observations': True, 'provenance': 'Synthetic tool-only full guard fixture'})
            static = cls.raw/('static-'+shell)
            value = {k: 0 for k in exporter.STATIC_BAD}
            value.update(result='pass', commit_under_test=R2, dirty_worktree=False, files_checked=2,
                         selected_rules=['SyntheticRule'], advisory_warnings=3, advisory_information=2,
                         shell_edition='Desktop' if shell == 'ps51' else 'Core', shell_version='5.1.0' if shell == 'ps51' else '7.0.0')
            save(static/'analysis.json', value)
            roots.append({'source': str(static), 'label': 'static/'+shell, 'mode': 'flat',
                          'role': 'accepted_static', 'shell': shell, 'provenance': 'Synthetic tool-only static guard fixture'})
        extras = cls.raw/'extras'
        save(extras/'aggregate.json', {'result': 'pass', 'commit_under_test': R2, 'commands': 1})
        save(extras/'invocations.json', [{'exit_code': 0, 'commit_under_test': R2, 'dirty_worktree': False}])
        roots.append({'source': str(extras), 'label': 'extras', 'mode': 'recursive', 'role': 'accepted_extras',
                      'provenance': 'Synthetic tool-only extras guard fixture'})
        ci = cls.raw/'ci'
        save(ci/'push.json', {'headSha': R2, 'conclusion': 'success', 'event': 'push',
                             'jobs': [{'conclusion': 'success'} for _ in range(4)]})
        for job in range(4):
            path = ci/'push'/('job'+str(job))
            save(path/'job.json', {'commit_under_test': R2, 'result': 'pass', 'source_unchanged': True})
            for tier in range(5):
                save(path/('tier'+str(tier))/'summary.json', {'commit_under_test': R2, 'result': 'pass',
                    'accepted': True, 'source_unchanged': True, 'passed': 1, 'total': 1})
                (path/('tier'+str(tier))/'results.xml').write_text('<test-results><test-case result="Success"/></test-results>', encoding='utf-8')
        roots.append({'source': str(ci), 'label': 'ci', 'mode': 'recursive', 'role': 'accepted_ci',
                      'metadata': 'push.json', 'artifacts': 'push', 'event': 'push', 'provenance': 'Synthetic tool-only CI guard fixture'})
        failed = cls.raw/'failed'
        save(failed/'aggregate.json', {'result': 'fail', 'commit_under_test': INITIAL, 'tiers_completed': 0})
        (failed/'audited-packet-diff.stdout').write_text('synthetic huge-diff placeholder', encoding='utf-8')
        (failed/'failure.stderr').write_text('synthetic fixture oracle failure\r\n', encoding='utf-8')
        roots.append({'source': str(failed), 'label': 'unaccepted-R/initial', 'mode': 'recursive', 'role': 'unaccepted',
                      'historical_commit': INITIAL, 'scope': 'Synthetic preserved failure only', 'provenance': 'Synthetic tool-only historical guard fixture'})
        cls.config = {'schema_version': 1, 'task': 'T31', 'source_commit': R2,
                      'initial_unaccepted_R': INITIAL, 'roots': roots, 'files': []}
        cls.config_path = cls.root/'selection.json'; save(cls.config_path, cls.config)
        cls.result = exporter.Projector(REPO, cls.config_path).run(cls.destination, write=True)
        cls.manifest = json.loads((cls.destination/'manifest.json').read_text())

    def test_01_types_and_paths(self):
        actual = json.loads((self.destination/'ps51/typed.json').read_text())
        expected = dict(self.semantic); expected['path'] = '<USERPROFILE>\\synthetic-only'
        self.assertEqual(actual, expected)
        self.assertFalse((self.destination/'ps51/typed.json').read_bytes().startswith(b'\xef\xbb\xbf'))
        for key in ('true', 'false', 'integer', 'decimal', 'null'):
            self.assertIs(type(actual[key]), type(self.semantic[key]))

    def test_02_xml_entities_identities_and_bom(self):
        raw = (self.destination/'ps7/tier0.results.xml').read_bytes()
        self.assertTrue(raw.startswith(b'\xef\xbb\xbf'))
        self.assertIn(b'&lt;REPO&gt;', raw)
        parsed = ET.fromstring(raw); env = parsed.find('environment')
        self.assertEqual(env.attrib['user'], '<USER>')
        self.assertEqual(env.attrib['machine-name'], '<COMPUTER>')
        self.assertEqual(env.attrib['user-domain'], '<COMPUTER>')
        self.assertEqual(env.attrib['cwd'], '<REPO>')

    def test_03_inline_json_and_text_stream_name(self):
        raw = (self.destination/'ps51/inline.stdout.txt').read_bytes()
        self.assertTrue(raw.startswith(b'\xef\xbb\xbf') and raw.endswith(b'\r\n'))
        self.assertEqual(json.loads(raw.decode('utf-8-sig'))['path'], '<USERPROFILE>\\synthetic-only')
        self.assertEqual(len(json.loads((self.destination/'ps51/runs.json').read_text())[0]['observation_receipts']), 2)
        observation = json.loads((self.destination/'ps51/observations/tier0/1/semantics.json').read_text())
        self.assertIs(observation['false'], False)
        self.assertTrue((self.destination/'unaccepted-R/initial/failure.stderr.txt').is_file())

    def test_04_raw_public_hashes_and_binary_exclusion(self):
        paths = {x['path'] for x in self.manifest['files']}
        self.assertNotIn('ps51/excluded.pdf', paths)
        self.assertNotIn('unaccepted-R/initial/audited-packet-diff.stdout.txt', paths)
        self.assertEqual(len(self.manifest['omitted_text_bindings']), 1)
        for row in self.manifest['files']:
            public = (self.destination/row['path']).read_bytes()
            self.assertEqual(exporter.sha(public), row['sha256'])
            self.assertEqual(len(public), row['bytes'])
            self.assertEqual(len(row['raw_sha256']), 64)
        self.assertEqual(self.manifest['post_manifest_review_files'], exporter.POST)

    def test_05_idempotent_immutable_packet(self):
        before = {p: p.read_bytes() for p in self.destination.rglob('*') if p.is_file()}
        again = exporter.Projector(REPO, self.config_path).run(self.destination, write=True)
        self.assertEqual(again['manifest_sha256'], self.result['manifest_sha256'])
        self.assertTrue(all(p.read_bytes() == b for p, b in before.items()))

    def test_06_unknown_or_initial_R_rejected(self):
        for value in (None, INITIAL):
            bad = dict(self.config); bad['source_commit'] = value
            path = self.root/('bad-'+str(value)+'.json'); save(path, bad)
            with self.assertRaises(ValueError): exporter.Projector(REPO, path)

    def test_07_wrong_full_source_fails_before_write(self):
        path = self.raw/'ps51/aggregate.json'; old = path.read_bytes()
        value = json.loads(old); value['commit_under_test'] = INITIAL; save(path, value)
        destination = self.root/'must-not-be-created'
        try:
            with self.assertRaises(ValueError): exporter.Projector(REPO, self.config_path).run(destination, write=True)
            self.assertFalse(destination.exists())
        finally: path.write_bytes(old)

    def test_08_escaped_prefixes(self):
        projector = exporter.Projector(REPO, self.config_path)
        for slash_count in (1, 2, 4, 8, 16):
            escaped = self.profile.replace('\\', '\\'*slash_count)
            self.assertEqual(projector.replace(escaped+'\\fixture'), '<USERPROFILE>\\fixture')
        self.assertEqual(projector.replace(self.profile.replace('\\', '/').replace('/', '\\/')), '<USERPROFILE>')

    def test_09_unaccepted_failure_classification(self):
        failed = json.loads((self.destination/'unaccepted-R/initial/aggregate.json').read_text())
        self.assertEqual(failed['result'], 'fail'); self.assertEqual(failed['commit_under_test'], INITIAL)
        self.assertEqual(self.manifest['initial_unaccepted_R'], INITIAL)
        self.assertTrue(all(x['source_commit'] == R2 for x in self.manifest['accepted_guard_results']))

    def test_10_username_plaintext_fails_before_write(self):
        path = self.raw/'extras/private.txt'; path.write_text(Path(self.profile).name, encoding='utf-8')
        destination = self.root/'privacy-must-not-be-created'
        try:
            with self.assertRaises(ValueError): exporter.Projector(REPO, self.config_path).run(destination, write=True)
            self.assertFalse(destination.exists())
        finally: path.unlink()

    def test_11_key_collision_is_rejected(self):
        projector = exporter.Projector(REPO, self.config_path)
        with self.assertRaises(ValueError): projector.typed({str(REPO): 1, '<REPO>': 2})

    def test_12_external_receipt_link_rejected(self):
        projector = exporter.Projector(REPO, self.config_path)
        with self.assertRaises(ValueError): projector.owned(REPO/'README.md')

    def test_13_wrong_CI_checkout_fails_before_write(self):
        path = self.raw/'ci/push/job0/job.json'; old = path.read_bytes()
        value = json.loads(old); value['commit_under_test'] = INITIAL; save(path, value)
        destination = self.root/'ci-must-not-be-created'
        try:
            with self.assertRaises(ValueError): exporter.Projector(REPO, self.config_path).run(destination, write=True)
            self.assertFalse(destination.exists())
        finally: path.write_bytes(old)

    def test_14_failed_extras_cannot_be_accepted(self):
        path = self.raw/'extras/invocations.json'; old = path.read_bytes()
        value = json.loads(old); value[0]['exit_code'] = 1; save(path, value)
        destination = self.root/'extras-must-not-be-created'
        try:
            with self.assertRaises(ValueError): exporter.Projector(REPO, self.config_path).run(destination, write=True)
            self.assertFalse(destination.exists())
        finally: path.write_bytes(old)

    def test_15_binary_disguised_as_text_is_rejected(self):
        path = self.raw/'extras/disguised.txt'; path.write_bytes(b'MZsynthetic-only')
        destination = self.root/'binary-must-not-be-created'
        try:
            with self.assertRaises(ValueError): exporter.Projector(REPO, self.config_path).run(destination, write=True)
            self.assertFalse(destination.exists())
        finally: path.unlink()

if __name__ == '__main__': unittest.main(verbosity=2)
