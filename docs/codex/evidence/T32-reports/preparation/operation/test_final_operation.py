"""Focused T32 capture guards only; never run application/native orchestration."""
import copy, importlib.util, json
from pathlib import Path
import tempfile, types, unittest
from unittest import mock

HERE = Path(__file__).resolve().parent


def load(name, filename):
    spec = importlib.util.spec_from_file_location(name, HERE / filename)
    module = importlib.util.module_from_spec(spec); spec.loader.exec_module(module)
    return module


smoke = load('t32_final_smoke', 'final_package_smoke.py')
capture = load('t32_final_capture', 'capture-T32.py')


class FinalOperationGuards(unittest.TestCase):
    def test_import_has_no_capture_or_application_side_effects(self):
        with mock.patch('subprocess.Popen') as process:
            load('t32_import_smoke', 'final_package_smoke.py')
            load('t32_import_capture', 'capture-T32.py')
            process.assert_not_called()

    def test_historical_source_is_refused_before_processes_or_owned_directories(self):
        arguments = types.SimpleNamespace(candidate_source_commit='8917938820f60e499e2c20caa9cb03171678be72')
        with mock.patch('subprocess.Popen') as process, mock.patch.object(Path, 'mkdir') as mkdir:
            with self.assertRaisesRegex(RuntimeError, 'Exact accepted release source R'):
                smoke.run(arguments)
            process.assert_not_called(); mkdir.assert_not_called()
        smoke.require_release_source(capture.SOURCE)

    def test_both_hosts_use_identical_one_asset_pair_and_frozen_source(self):
        assets = {'zip_path': 'synthetic final ZIP.zip', 'zip_sha256': 'a' * 64,
                  'checksums_path': 'synthetic checksums.txt', 'checksums_sha256': 'b' * 64}
        selected = {'pdftk.exe': 'synthetic pdftk.exe', 'gswin64c.exe': 'synthetic gs.exe'}
        calls = [capture.operation_command(Path('repo'), Path('harness'), 'c' * 40, assets, kind + '.exe', kind,
                 selected, Path('manifest'), Path('external spaces'), Path('capture')) for kind in ('PS51', 'PS7')]
        for flag in ('--zip', '--zip-sha256', '--checksums', '--checksums-sha256', '--candidate-source-commit'):
            self.assertEqual(calls[0][calls[0].index(flag) + 1], calls[1][calls[1].index(flag) + 1])
        self.assertEqual(calls[0][calls[0].index('--candidate-source-commit') + 1], capture.SOURCE)
        self.assertNotIn('--preparation', calls[0]); self.assertNotIn('--preparation', calls[1])
        self.assertNotEqual(calls[0][calls[0].index('--work-root') + 1], calls[1][calls[1].index('--work-root') + 1])

    def test_whole_checksum_asset_and_separate_zip_pin_are_required(self):
        with tempfile.TemporaryDirectory(prefix='T32 unit assets ') as temporary:
            root = Path(temporary); archive = root / 'WinPDFMerger-v1.0.0.zip'; checksums = root / 'SHA256SUMS.txt'
            archive.write_bytes(b'synthetic unit bytes, never an operated application ZIP')
            zip_hash = capture.digest(archive); checksums.write_bytes((zip_hash + '  WinPDFMerger-v1.0.0.zip\n').encode())
            sum_hash = capture.digest(checksums)
            capture.validate_assets(archive, zip_hash, checksums, sum_hash)
            with self.assertRaises(ValueError): capture.validate_assets(archive, 'a' * 64, checksums, sum_hash)
            checksums.write_bytes(checksums.read_bytes() + b'extra unreviewed row\n')
            with self.assertRaises(ValueError): capture.validate_assets(archive, zip_hash, checksums, sum_hash)
            with self.assertRaises(ValueError): capture.validate_assets(archive, zip_hash, checksums, capture.digest(checksums))

    def valid_report(self, kind):
        assets = {'zip_sha256': 'a' * 64, 'checksums_sha256': 'b' * 64}
        return {'task': 'T32', 'result': 'pass', 'preparation': False, 'harness_commit': 'c' * 40,
                'candidate_source_commit': capture.SOURCE, 'shell_kind': kind, 'candidate': dict(assets),
                'cases': [{} for _ in range(14 if kind == 'PS51' else 11)],
                'source_guard': {key: True for key in capture.GUARDS},
                'manual_acceptance': 'excluded/unperformed; never pass'}, assets

    def test_accepted_report_requires_every_source_guard_and_exact_metadata(self):
        for kind in ('PS51', 'PS7'):
            valid, assets = self.valid_report(kind)
            capture.validate_report(valid, kind, 'c' * 40, assets)
            for mutate in (lambda value: value.update(preparation=True), lambda value: value['source_guard'].pop('clean'),
                           lambda value: value['source_guard'].update(driver_unchanged=False),
                           lambda value: value.update(candidate_source_commit='8917938820f60e499e2c20caa9cb03171678be72'),
                           lambda value: value['candidate'].update(zip_sha256='d' * 64),
                           lambda value: value['cases'].pop(), lambda value: value.update(manual_acceptance='pass')):
                bad = copy.deepcopy(valid); mutate(bad)
                with self.assertRaises(ValueError): capture.validate_report(bad, kind, 'c' * 40, assets)


if __name__ == '__main__': unittest.main(verbosity=2)
