"""Independent auditor synthetic checks; no exporter or real packet is imported."""
from pathlib import Path
import copy, hashlib, importlib.util, json, os, unittest

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[2]
spec = importlib.util.spec_from_file_location('independent_privacy_auditor', HERE / 'public-review.py')
module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)


class GitIdentityAuditorTests(unittest.TestCase):
    def setUp(self):
        self.auditor = module.Auditor(REPO, HERE, HERE, HERE, '0' * 64)
        private_user = Path(os.environ['USERPROFILE']).name
        self.metadata = {'sha': 'a' * 40, 'node_id': 'synthetic',
                         'url': 'https://api.github.com/repos/synthetic/project/git/commits/' + 'a' * 40,
                         'html_url': 'https://github.com/synthetic/project/commit/' + 'a' * 40,
                         'author': {'name': 'Synthetic author', 'email': private_user + '@example.invalid', 'date': '2026-10-09T00:00:00Z'},
                         'committer': {'name': 'Synthetic committer', 'email': 'commit-' + private_user + '@example.invalid', 'date': '2026-10-09T00:00:01Z'},
                         'tree': {'sha': 'b' * 40, 'url': 'https://api.github.com/synthetic/tree'},
                         'parents': [{'sha': 'c' * 40, 'url': 'https://api.github.com/synthetic/parent'}],
                         'message': 'Synthetic message kept exactly',
                         'verification': {'verified': False, 'reason': 'unsigned', 'signature': None, 'payload': None},
                         'additional_facts': [True, False, 0, 1, 1.25, None, '<USERPROFILE>']}
        self.raw = (json.dumps(self.metadata, ensure_ascii=False) + '\n').encode()
        self.declaration = {'path': 'synthetic/commit.stdout.txt', 'raw_sha256': hashlib.sha256(self.raw).hexdigest(), 'kind': 'git_commit', 'git_commit_sha': 'a' * 40}

    def project(self, raw=None, declaration=None):
        return self.auditor.github_metadata_identity(raw or self.raw, 'synthetic/commit.stdout.txt', declaration or self.declaration)

    def actions_receipt(self):
        value = {'id': 123, 'head_sha': self.metadata['sha'], 'event': 'pull_request', 'status': 'completed',
                 'conclusion': 'success', 'url': 'https://api.github.com/synthetic/run/123',
                 'html_url': 'https://github.com/synthetic/project/actions/runs/123',
                 'head_commit': {'id': self.metadata['sha'], 'tree_id': self.metadata['tree']['sha'],
                                 'message': self.metadata['message'], 'timestamp': '2026-10-09T00:00:00Z',
                                 'author': {key: self.metadata['author'][key] for key in ('name', 'email')},
                                 'committer': {key: self.metadata['committer'][key] for key in ('name', 'email')}},
                 'other_facts': [True, False, 1, None]}
        raw = json.dumps(value).encode()
        declaration = {'path': 'synthetic/run.stdout.txt', 'raw_sha256': hashlib.sha256(raw).hexdigest(),
                       'kind': 'actions_run', 'git_commit_sha': 'a' * 40, 'run_id': 123}
        return value, raw, declaration

    def test_only_two_declared_email_fields_change(self):
        original, comparison = self.project()
        expected = copy.deepcopy(self.metadata)
        expected['author']['email'] = expected['committer']['email'] = '<EMAIL>'
        self.assertEqual(original, self.metadata)
        self.assertEqual(comparison, expected)
        self.assertFalse(self.auditor.issues)

    def test_boolean_number_null_facts_are_not_coerced(self):
        _, comparison = self.project()
        public = self.auditor.projected(comparison)
        self.auditor.walk(comparison, public, 'synthetic')
        self.assertFalse(self.auditor.issues)
        self.assertIs(public['verification']['verified'], False)
        self.assertIsNone(public['verification']['payload'])
        public['verification']['verified'] = 0
        self.auditor.walk(comparison, public, 'synthetic-tampered-type')
        self.assertTrue(any('JSON type retained' in issue for issue in self.auditor.issues))

    def test_changed_tree_or_parent_fact_is_rejected(self):
        _, comparison = self.project()
        public = self.auditor.projected(comparison)
        public['tree']['sha'] = 'd' * 40
        public['parents'][0]['sha'] = 'e' * 40
        self.auditor.walk(comparison, public, 'synthetic-tampered-provenance')
        self.assertEqual(sum('prefix-only string change' in issue for issue in self.auditor.issues), 2)

    def test_wrong_raw_hash_or_commit_pin_cannot_pass(self):
        bad = dict(self.declaration, raw_sha256='0' * 64, git_commit_sha='0' * 40)
        self.project(declaration=bad)
        self.assertTrue(any('raw SHA pin' in issue for issue in self.auditor.issues))
        self.assertTrue(any('commit SHA pin' in issue for issue in self.auditor.issues))

    def test_unrecognized_or_type_confused_API_schema_aborts(self):
        for mutate in (lambda value: value.pop('author'), lambda value: value['committer'].update(email=True),
                       lambda value: value.update(tree={'sha': 17}), lambda value: value.update(message={'text': 'unknown'})):
            with self.subTest(mutation=mutate):
                value = copy.deepcopy(self.metadata); mutate(value)
                raw = json.dumps(value).encode()
                declaration = dict(self.declaration, raw_sha256=hashlib.sha256(raw).hexdigest())
                with self.assertRaises(ValueError): self.project(raw, declaration)

    def test_unknown_identity_elsewhere_is_kept_and_fails_privacy(self):
        value = copy.deepcopy(self.metadata)
        value['message'] = Path(os.environ['USERPROFILE']).name + ' synthetic-only'
        raw = json.dumps(value).encode()
        declaration = dict(self.declaration, raw_sha256=hashlib.sha256(raw).hexdigest())
        _, comparison = self.project(raw, declaration)
        self.assertEqual(comparison['message'], value['message'])
        public = json.dumps(self.auditor.projected(comparison)).encode()
        self.auditor.privacy(public, 'synthetic-unknown-identity')
        self.assertTrue(any('original Windows identity' in issue for issue in self.auditor.issues))

    def test_JSON_BOM_is_parsed_and_canonical_output_drops_it(self):
        raw = b'\xef\xbb\xbf' + self.raw
        declaration = dict(self.declaration, raw_sha256=hashlib.sha256(raw).hexdigest())
        _, comparison = self.project(raw, declaration)
        public = (json.dumps(self.auditor.projected(comparison), indent=2, ensure_ascii=False) + '\n').encode()
        self.assertFalse(public.startswith(b'\xef\xbb\xbf'))
        self.assertTrue(public.endswith(b'\n'))

    def test_actions_run_changes_only_nested_email_fields(self):
        value, raw, declaration = self.actions_receipt()
        original, comparison = self.project(raw, declaration)
        expected = copy.deepcopy(value)
        expected['head_commit']['author']['email'] = expected['head_commit']['committer']['email'] = '<EMAIL>'
        self.assertEqual(original, value)
        self.assertEqual(comparison, expected)
        self.auditor.walk(comparison, self.auditor.projected(comparison), 'synthetic-actions')
        self.assertFalse(self.auditor.issues)

    def test_actions_run_and_commit_pins_are_both_required(self):
        _, raw, declaration = self.actions_receipt()
        bad = dict(declaration, run_id=124, git_commit_sha='e' * 40)
        self.project(raw, bad)
        self.assertTrue(any('run ID pin' in issue for issue in self.auditor.issues))
        self.assertTrue(any('head/commit SHA pins' in issue for issue in self.auditor.issues))

    def test_actions_boolean_ID_cannot_equal_integer_pin(self):
        value, _, declaration = self.actions_receipt()
        value['id'] = True
        raw = json.dumps(value).encode()
        bad = dict(declaration, run_id=1, raw_sha256=hashlib.sha256(raw).hexdigest())
        with self.assertRaises(ValueError): self.project(raw, bad)

    def test_public_email_redaction_is_not_broadened(self):
        value = copy.deepcopy(self.metadata)
        value['author']['email'] = 'public-author@example.invalid'
        raw = json.dumps(value).encode()
        declaration = dict(self.declaration, raw_sha256=hashlib.sha256(raw).hexdigest())
        self.project(raw, declaration)
        self.assertTrue(any('scope not broadened' in issue for issue in self.auditor.issues))


if __name__ == '__main__': unittest.main(verbosity=2)
