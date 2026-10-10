"""Constrained privacy-projection checks only; no application/native/CI execution."""
from pathlib import Path
import copy, hashlib, importlib.util, json, os, tempfile, unittest

HERE = Path(__file__).resolve().parent; REPO = HERE.parents[2]
spec = importlib.util.spec_from_file_location('new_projector', HERE/'Export-T31.py')
producer = importlib.util.module_from_spec(spec); spec.loader.exec_module(producer)
REGISTRY = json.loads((HERE/'identity-registry.json').read_bytes())
RAW = REPO/'tests/.work/T31-review/fix-CI-37971199628-925c9f3d01c141048866809aa332093c'
sha = lambda b: hashlib.sha256(b).hexdigest()

class IdentityChecks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = Path(tempfile.mkdtemp(prefix='tool-only-', dir=HERE))
        config = json.loads((REPO/'tests/.work/T31-export-selection-v3.json').read_bytes())
        config['github_metadata_identity_receipts'] = REGISTRY['github_metadata_identity_receipts']
        cls.config = cls.root/'config.raw'; cls.config.write_text(json.dumps(config), encoding='utf-8')
        cls.base = config

    def make(self, row, raw=None):
        raw = (RAW/Path(row['path']).name).read_bytes() if raw is None else raw
        config = copy.deepcopy(self.base); config['github_metadata_identity_receipts'] = [dict(row)]
        path = self.root/'case-config.raw'; path.write_text(json.dumps(config), encoding='utf-8')
        p = producer.Projector(REPO, path)
        return p, raw

    def assert_only_emails(self, original, public, kind):
        expected = copy.deepcopy(original)
        target = expected if kind == 'git_commit' else expected['head_commit']
        target['author']['email'] = target['committer']['email'] = '<EMAIL>'
        self.assertEqual(json.loads(public), expected)

    def test_01_actual_git_commit_preserves_every_other_fact(self):
        row = REGISTRY['github_metadata_identity_receipts'][0]; p, raw = self.make(row)
        public = p.github_metadata(raw, row['path'])
        self.assert_only_emails(json.loads(raw), public, row['kind'])
        self.assertIs(type(json.loads(public)['verification']['verified']), bool)
        self.assertIs(type(json.loads(public)['verification']['signature']), type(json.loads(raw)['verification']['signature']))

    def test_02_actual_actions_run_preserves_every_other_fact(self):
        row = REGISTRY['github_metadata_identity_receipts'][1]; p, raw = self.make(row)
        public = p.github_metadata(raw, row['path'])
        self.assert_only_emails(json.loads(raw), public, row['kind'])
        self.assertEqual(json.loads(public)['id'], 37971199628)

    def test_03_raw_hash_tampering_rejected(self):
        row = REGISTRY['github_metadata_identity_receipts'][0]; p, raw = self.make(row)
        with self.assertRaises(ValueError): p.github_metadata(raw+b' ', row['path'])

    def test_04_wrong_commit_pin_rejected(self):
        row = dict(REGISTRY['github_metadata_identity_receipts'][0]); row['git_commit_sha'] = 'a'*40
        p, raw = self.make(row)
        with self.assertRaises(ValueError): p.github_metadata(raw, row['path'])

    def test_05_wrong_run_pin_rejected(self):
        row = dict(REGISTRY['github_metadata_identity_receipts'][1]); row['run_id'] += 1
        p, raw = self.make(row)
        with self.assertRaises(ValueError): p.github_metadata(raw, row['path'])

    def test_06_wrong_schema_rejected_even_if_new_raw_hash_is_declared(self):
        row = dict(REGISTRY['github_metadata_identity_receipts'][0])
        value = json.loads((RAW/Path(row['path']).name).read_bytes()); del value['verification']
        raw = json.dumps(value).encode(); row['raw_sha256'] = sha(raw); p, raw = self.make(row, raw)
        with self.assertRaises(ValueError): p.github_metadata(raw, row['path'])

    def test_07_unknown_private_identity_outside_email_still_blocks(self):
        row = dict(REGISTRY['github_metadata_identity_receipts'][0])
        value = json.loads((RAW/Path(row['path']).name).read_bytes()); value['message'] = Path(os.environ['USERPROFILE']).name
        raw = json.dumps(value).encode(); row['raw_sha256'] = sha(raw); p, raw = self.make(row, raw)
        path = self.root/'unknown-identity.raw'; path.write_bytes(raw); p.selected = {row['path']: (path, 'Synthetic negative privacy test')}
        with self.assertRaises(ValueError): p.payloads()

    def test_08_BOM_is_normalized_as_typed_JSON(self):
        row = dict(REGISTRY['github_metadata_identity_receipts'][0]); raw = b'\xef\xbb\xbf'+(RAW/Path(row['path']).name).read_bytes()
        row['raw_sha256'] = sha(raw); p, raw = self.make(row, raw)
        public = p.github_metadata(raw, row['path'])
        self.assertFalse(public.startswith(b'\xef\xbb\xbf'))
        self.assert_only_emails(json.loads(raw.decode('utf-8-sig')), public, row['kind'])

    def test_09_exact_new_rule_is_recorded(self):
        row = REGISTRY['github_metadata_identity_receipts'][0]; p, raw = self.make(row)
        p.selected = {row['path']: (RAW/Path(row['path']).name, 'Actual immutable GitHub receipt privacy projection check')}
        staged, rows = p.payloads()
        self.assertEqual(rows[0]['projection'], REGISTRY['projection_rule'])
        self.assertEqual(rows[0]['raw_sha256'], row['raw_sha256'])
        self.assertEqual(sha(staged[0][1]), rows[0]['sha256'])

    def test_10_general_JSON_and_nonidentity_field_types_remain(self):
        row = dict(REGISTRY['github_metadata_identity_receipts'][0]); value = json.loads((RAW/Path(row['path']).name).read_bytes())
        value['extra_facts'] = {'bool': True, 'false': False, 'integer': 7, 'decimal': 0.5, 'null': None}
        raw = json.dumps(value).encode(); row['raw_sha256'] = sha(raw); p, raw = self.make(row, raw)
        public = p.github_metadata(raw, row['path']); actual = json.loads(public)
        self.assert_only_emails(value, public, row['kind'])
        for key, original in value['extra_facts'].items(): self.assertIs(type(actual['extra_facts'][key]), type(original))

    def test_11_undeclared_receipt_still_uses_old_privacy_refusal(self):
        row = REGISTRY['github_metadata_identity_receipts'][0]; p, raw = self.make(row)
        p.github_identity_receipts = {}; p.selected = {row['path']: (RAW/Path(row['path']).name, 'Synthetic undeclared receipt check')}
        with self.assertRaises(ValueError): p.payloads()

    def test_12_original_frozen_sources_are_still_exact(self):
        self.assertEqual(sha((REPO/'tests/.work/T31-export-preparation/Export-T31.py').read_bytes()), '03288f27e9def307e3fbedfee7b3a3c8eb637dfaac8369d5a81994823f1dce2c')
        self.assertEqual(sha((REPO/'tests/.work/T31-export-selection-v3.json').read_bytes()), '318006165313cbdf83a56c134b969b5c66d63b3f35944f5026c711fb80000487')
        for row in REGISTRY['github_metadata_identity_receipts']:
            self.assertEqual(sha((RAW/Path(row['path']).name).read_bytes()), row['raw_sha256'])

if __name__ == '__main__': unittest.main(verbosity=2)
