"""T33 developer regressions. Never execute a shell/application/native engine."""
import ast
import copy
import importlib.util
from pathlib import Path
import types
import unittest
from unittest import mock

HERE = Path(__file__).resolve().parent
OLD = HERE.parent/'T32-operation-preparation'
def module(name,path):
    spec = importlib.util.spec_from_file_location(name,path)
    result = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(result)
    return result
smoke = module('T33_published_smoke',HERE/'final_package_smoke.py')
capture = module('T33_published_capture',HERE/'capture-T33.py')

class PublishedDerivative(unittest.TestCase):
    def test_wrong_harness_M_refused_before_any_process_or_directory(self):
        args = types.SimpleNamespace(candidate_source_commit=capture.SOURCE,expected_harness_commit='a'*40)
        with mock.patch('subprocess.Popen') as process, mock.patch.object(Path,'mkdir') as mkdir:
            with self.assertRaisesRegex(RuntimeError,'Exact owner-merged T33 harness commit M'):
                smoke.run(args)
            process.assert_not_called()
            mkdir.assert_not_called()

    def test_outer_wrong_harness_M_refused_before_process_or_directory(self):
        argv = ['capture','--expected-harness-commit','a'*40,'--zip','unit.zip','--zip-sha256','a'*64,
                '--checksums','unit.txt','--checksums-sha256','b'*64]
        with mock.patch('sys.argv',argv),mock.patch('subprocess.Popen') as process,mock.patch.object(Path,'mkdir') as mkdir:
            with self.assertRaisesRegex(ValueError,'exact owner-merged T33 harness commit M'):
                capture.main()
            process.assert_not_called()
            mkdir.assert_not_called()

    def test_task_and_download_class_cannot_reuse_T32_report(self):
        pins={'zip_sha256':'a'*64,'checksums_sha256':'b'*64}
        for kind,count in [('PS51',14),('PS7',11)]:
            valid={'task':'T33','evidence_class':capture.EVIDENCE_CLASS,'result':'pass','preparation':False,
                   'harness_commit':capture.EXPECTED_HARNESS,'candidate_source_commit':capture.SOURCE,
                   'shell_kind':kind,'candidate':pins.copy(),'cases':[{}]*count,
                   'source_guard':{k:True for k in capture.GUARDS},'manual_acceptance':'excluded/unperformed; never pass'}
            capture.validate_report(valid,kind,capture.EXPECTED_HARNESS,pins)
            for change in [{'task':'T32'},{'evidence_class':'actual_final_R_package_operation'},
                           {'manual_acceptance':'pass'},{'preparation':True}]:
                invalid=copy.deepcopy(valid);invalid.update(change)
                with self.assertRaises(ValueError):
                    capture.validate_report(invalid,kind,capture.EXPECTED_HARNESS,pins)

    def test_scenario_and_native_safe_process_body_unchanged(self):
        original=ast.parse((OLD/'final_package_smoke.py').read_text(encoding='utf-8'))
        derived=ast.parse((HERE/'final_package_smoke.py').read_text(encoding='utf-8'))
        old_run=next(n for n in original.body if isinstance(n,ast.FunctionDef) and n.name=='run')
        new_run=next(n for n in derived.body if isinstance(n,ast.FunctionDef) and n.name=='run')
        old_cases=next(i for i,n in enumerate(old_run.body) if isinstance(n,ast.Try))
        new_cases=next(i for i,n in enumerate(new_run.body) if isinstance(n,ast.Try))
        self.assertEqual(ast.dump(old_run.body[old_cases],include_attributes=False),ast.dump(new_run.body[new_cases],include_attributes=False))

    def test_all_nonorchestrating_harness_functions_ast_unchanged(self):
        def functions(path):
            return {n.name:ast.dump(n,include_attributes=False) for n in ast.parse(path.read_text(encoding='utf-8')).body if isinstance(n,ast.FunctionDef)}
        a,b=functions(OLD/'final_package_smoke.py'),functions(HERE/'final_package_smoke.py')
        self.assertEqual(set(a),set(b))
        for name in a:
            if name!='run':
                self.assertEqual(a[name],b[name],name)

    def test_legacy_candidate_CLI_retains_exact_R_and_same_pair_both_hosts(self):
        assets={'zip_path':'anonymous ZIP.zip','zip_sha256':'a'*64,'checksums_path':'anonymous sum.txt','checksums_sha256':'b'*64}
        selected={'pdftk.exe':'unit pdftk.exe','gswin64c.exe':'unit gs.exe'}
        commands=[capture.operation_command(Path('repo'),Path('harness'),capture.EXPECTED_HARNESS,assets,k+'.exe',k,
            selected,Path('manifest'),Path('fresh spaces'),Path('capture'))for k in ['PS51','PS7']]
        for flag in ['--zip','--zip-sha256','--checksums','--checksums-sha256','--candidate-source-commit','--expected-harness-commit']:
            self.assertEqual(commands[0][commands[0].index(flag)+1],commands[1][commands[1].index(flag)+1])
        self.assertEqual(commands[0][commands[0].index('--candidate-source-commit')+1],capture.SOURCE)
        self.assertEqual(capture.SOURCE,smoke.ACCEPTED_SOURCE)
        self.assertEqual(capture.EXPECTED_HARNESS,smoke.EXPECTED_HARNESS)

if __name__=='__main__':
    unittest.main(verbosity=2)
