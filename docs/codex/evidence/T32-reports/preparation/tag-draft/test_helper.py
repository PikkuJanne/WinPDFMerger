"""Synthetic helper safety checks only: never invoke git/gh or remote writes."""
import copy
import importlib.util
import json
from pathlib import Path
import subprocess
import tempfile
import unittest
from types import SimpleNamespace

HERE=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('tagdraft',HERE/'TagDraft-T32.py')
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)

class Safety(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name);self.calls=[]
    def tearDown(self): self.tmp.cleanup()
    def fake(self,argv,**kw):
        self.calls.append(argv);kw['stdout'].write(b'captured');kw['stderr'].write(b'')
        return SimpleNamespace(returncode=0)
    def capture(self,execute=False,runner=None):
        p=self.root/'receipt';p.mkdir()
        return module.Capture(self.root,p,execute,'a'*64,runner=runner or self.fake)
    def write(self,name,value):
        p=self.root/name;p.write_text(json.dumps(value));return p
    def gates(self):
        z,s='1'*64,'2'*64
        children=[]
        for kind in ('PS51','PS7'):
            child={'task':'T32','result':'pass','preparation':False,'candidate_source_commit':module.R,
                'harness_commit':'3'*40,'shell_kind':kind,'candidate':{'zip_sha256':z,'checksums_sha256':s},
                'cases':[{'package_guard':True,'source_foreign_guard':True} for _ in range(14 if kind=='PS51' else 11)],
                'public_help':{'exit_code':0,'package_unchanged':True,'no_outputs':True},
                'source_guard':{k:True for k in module.GUARDS},'manual_acceptance':'excluded/unperformed; never pass'}
            p=self.write(kind+'.json',child);children.append({'shell':kind,'path':str(p),'sha256':module.sha(p)})
        ledger={'task':'T32','source_commit':module.R,'result':'pass','harness_commit':'3'*40,
            'source_clean_before_after':True,'driver_unchanged':True,'cache_and_assets_unchanged':True,
            'approved_cache_files_verified':348,'manual_acceptance':'excluded/unperformed; never pass',
            'shared_assets':{'zip_sha256':z,'checksums_sha256':s},'candidate_reports':children}
        lp=self.write('ledger.json',ledger)
        review={'task':'T32','source_commit':module.R,'result':'pass_for_exact_final_package_bytes',
            'zip_sha256':z,'checksums_sha256':s,'checks':12,'issues':[]}
        ap=self.write('asset.json',review);np=self.write('native.json',{**review,'result':'pass_for_actual_final_native_scope'})
        rows=[]
        for role,path,result in [('native_ledger',lp,'pass'),('asset_review',ap,review['result']),('independent_native_review',np,'pass_for_actual_final_native_scope')]:
            rows.append({'role':role,'path':str(path),'sha256':module.sha(path),'expected_result':result,
                'result_pointer':'/result','source_pointer':'/source_commit',
                'zip_pointer':'/shared_assets/zip_sha256' if role=='native_ledger' else '/zip_sha256',
                'checksums_pointer':'/shared_assets/checksums_sha256' if role=='native_ledger' else '/checksums_sha256',
                'issues_pointer':'/issues','checks_pointer':'/checks'})
        return {'task':'T32','repo':str(self.root),'source_commit':module.R,'gates':rows},z,s
    def repin(self,c,role,change):
        g=next(x for x in c['gates'] if x['role']==role);d=module.read(g['path']);change(d)
        Path(g['path']).write_text(json.dumps(d));g['sha256']=module.sha(g['path'])
    def test_default_blocks_mutation_before_runner(self):
        c=self.capture()
        with self.assertRaises(ValueError):c.call('blocked',['gh','release','create'],True)
        self.assertEqual(self.calls,[]);self.assertEqual(c.ledger['commands'],[])
    def test_read_receipt_contains_hashes_and_argv(self):
        c=self.capture();self.assertEqual(c.call('read',['gh','api','repos/example']),b'captured')
        r=c.ledger['commands'][0];self.assertEqual(r['exit_code'],0);self.assertEqual(r['stdout']['bytes'],8)
        self.assertEqual(r['argv'],self.calls[0]);self.assertFalse(c.ledger['remote_write_started'])
    def test_duplicate_command_label_never_runs_again(self):
        c=self.capture();c.call('read',['gh','api'])
        with self.assertRaises(ValueError):c.call('read',['gh','api'])
        self.assertEqual(len(self.calls),1)
    def test_unknown_mutation_outcome_preserves_and_stops(self):
        def timeout(argv,**kw): self.calls.append(argv);raise subprocess.TimeoutExpired(argv,1)
        c=self.capture(True,timeout)
        with self.assertRaises(subprocess.TimeoutExpired):c.call('write',['gh','release','create'],True)
        self.assertEqual(c.ledger['result'],'unknown_mutation_outcome_requires_read_only_inspection')
        self.assertEqual(c.ledger['commands'][0]['state'],'unknown_outcome')
        with self.assertRaises(ValueError):c.call('again',['gh','release','create'],True)
        self.assertEqual(len(self.calls),1)
    def test_nonzero_mutation_stops_without_retry(self):
        def failure(argv,**kw):self.calls.append(argv);return SimpleNamespace(returncode=1)
        c=self.capture(True,failure)
        with self.assertRaises(ValueError):c.call('write',['git','push'],True)
        self.assertEqual(c.ledger['result'],'mutation_error_requires_read_only_inspection')
        with self.assertRaises(ValueError):c.call('again',['git','push'],True)
        self.assertEqual(len(self.calls),1)
    def test_complete_synthetic_gate_schema(self):
        c,z,s=self.gates();self.assertEqual(set(module.validate_gates(c,z,s)),{'native_ledger','asset_review','independent_native_review'})
    def test_missing_independent_gate_refused(self):
        c,z,s=self.gates();c['gates'].pop()
        with self.assertRaises(ValueError):module.validate_gates(c,z,s)
    def test_tampered_report_hash_refused(self):
        c,z,s=self.gates();Path(c['gates'][1]['path']).write_text('{}')
        with self.assertRaises(ValueError):module.validate_gates(c,z,s)
    def test_other_commit_refused(self):
        c,z,s=self.gates();self.repin(c,'asset_review',lambda d:d.update(source_commit='4'*40))
        with self.assertRaises(ValueError):module.validate_gates(c,z,s)
    def test_other_asset_refused(self):
        c,z,s=self.gates();self.repin(c,'independent_native_review',lambda d:d.update(zip_sha256='4'*64))
        with self.assertRaises(ValueError):module.validate_gates(c,z,s)
    def test_preparation_native_report_refused(self):
        c,z,s=self.gates();g=c['gates'][0];d=module.read(g['path']);row=d['candidate_reports'][0];n=module.read(row['path'])
        n['preparation']=True;Path(row['path']).write_text(json.dumps(n));row['sha256']=module.sha(row['path'])
        Path(g['path']).write_text(json.dumps(d));g['sha256']=module.sha(g['path'])
        with self.assertRaises(ValueError):module.validate_gates(c,z,s)
    def test_native_source_guard_failure_refused(self):
        c,z,s=self.gates();self.repin(c,'native_ledger',lambda d:d.update(source_clean_before_after=False))
        with self.assertRaises(ValueError):module.validate_gates(c,z,s)
    def test_independent_issues_refused(self):
        c,z,s=self.gates();self.repin(c,'independent_native_review',lambda d:d.update(issues=['unresolved']))
        with self.assertRaises(ValueError):module.validate_gates(c,z,s)
    def test_zero_independent_checks_refused(self):
        c,z,s=self.gates();self.repin(c,'asset_review',lambda d:d.update(checks=0))
        with self.assertRaises(ValueError):module.validate_gates(c,z,s)
    def test_draft_metadata_exact_unpublished_pair(self):
        paths={name:self.root/name for name in module.ASSETS}
        for p in paths.values():p.write_bytes(b'bytes')
        d={'tag_name':module.TAG,'draft':True,'prerelease':False,'published_at':None,'assets':[{'name':n,'state':'uploaded','size':5} for n in module.ASSETS]}
        module.validate_draft(d,paths)
    def test_published_or_extra_asset_metadata_refused(self):
        paths={name:self.root/name for name in module.ASSETS}
        for p in paths.values():p.write_bytes(b'bytes')
        d={'tag_name':module.TAG,'draft':False,'prerelease':False,'published_at':'now','assets':[{'name':n,'state':'uploaded','size':5} for n in module.ASSETS]}
        with self.assertRaises(ValueError):module.validate_draft(d,paths)
        d.update(draft=True,published_at=None);d['assets'].append({'name':'extra','size':5,'state':'uploaded'})
        with self.assertRaises(ValueError):module.validate_draft(d,paths)
    def test_pointer_preserves_typed_values(self):
        self.assertIs(module.pointer({'a':[True,None,2]},'/a/0'),True)
        self.assertIsNone(module.pointer({'a':[True,None,2]},'/a/1'))
        self.assertEqual(module.pointer({'a':[True,None,2]},'/a/2'),2)

if __name__=='__main__':unittest.main(verbosity=2)
