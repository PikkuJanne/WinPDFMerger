"""Developer-only actual-schema/refusal probes; no actual exporter or app action."""
from copy import deepcopy
import json, tempfile, unittest
from pathlib import Path
from test_projector import module,root

class ClosureAdapter(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(dir=root,prefix='schema-');self.addCleanup(self.tmp.cleanup)
        self.base=Path(self.tmp.name);self.repo=root.parents[2]
        inputs=json.loads((self.repo/'tests/.work/T34-record-preparation/writer-inputs.json').read_text())['roles']
        self.rows=inputs
        self.g={k:json.loads((self.repo/v['path']).read_bytes().decode('utf-8-sig'))for k,v in inputs.items()}
        download=self.g['published_download'];historical=next(arg for command in download['commands']for arg in command['argv']if isinstance(arg,str)and 'WinPDFMerger-t32-source-'in arg)
        self.value={'task':'T34','schema_version':1,'source_commit':module.R,'owner_merged_main':module.M,'asset_sha256':module.PAIR,'roots':[],
            'acceptance_gates':[{'role':role,'source':inputs[role]['path'],'raw_sha256':inputs[role]['sha256']}for role in ('source_merge_review','readiness_review','published_download')],
            'native_reuse_binding':{'source':inputs['native_reuse']['path'],'raw_sha256':inputs['native_reuse']['sha256']},
            'local_path_aliases':[{'original':str(Path(download['download_directory'])),'alias':'<T34_PUBLIC_DOWNLOAD>','provenance':'Actual frozen anonymous GateF root'},
                                  {'original':str(Path(historical)),'alias':'<T34_SOURCE>','provenance':'Actual frozen anonymous/package cleanR argv'}]}
    def p(self):
        q=self.base/'config.json';q.write_text(json.dumps(self.value))
        return module.Projector(self.repo,q)
    def change(self,role,path,value):
        data=deepcopy(self.g[role]);o=data
        for key in path[:-1]:o=o[key]
        o[path[-1]]=value
        p=self.base/(role+'.json');p.write_text(json.dumps(data))
        if role=='native_reuse':self.value['native_reuse_binding']={'source':str(p),'raw_sha256':module.sha(p.read_bytes())}
        else:next(x for x in self.value['acceptance_gates']if x['role']==role).update(source=str(p),raw_sha256=module.sha(p.read_bytes()))
    def test_actual_three_roles_child_reuse_pass(self):
        p=self.p();p.acceptance();self.assertEqual(len(p.required_sources),5);self.assertEqual(len(p.guards),5)
    def test_wrong_original_hash(self):
        self.value['acceptance_gates'][0]['raw_sha256']='0'*64
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_missing_role(self):
        self.value['acceptance_gates'].pop()
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_wrong_owner_main(self):
        self.value['owner_merged_main']='0'*40
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_false_source_surface(self):
        self.change('source_merge_review',['facts','source_proof','all_changed_paths_docs_codex'],False)
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_future_readiness_completion_refused(self):
        self.change('readiness_review',['completed_tasks'],34)
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_new_native_readiness_claim_refused(self):
        self.change('readiness_review',['scope','application_native_CI_or_helper_tests_reexecuted'],True)
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_authenticated_download_refused(self):
        self.change('published_download',['authentication_used'],True)
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_download_native_execution_refused(self):
        self.change('published_download',['application_executed'],True)
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_wrong_download_alias_refused(self):
        self.value['local_path_aliases'][0]['original']='C:/projects/WinPDFMerger-t34-public-download-'+'a'*32
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_wrong_source_alias_refused(self):
        self.value['local_path_aliases'][1]['original']='C:/projects/WinPDFMerger-t32-source-'+'a'*32
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_reuse_new_native_claim_refused(self):
        self.change('native_reuse',['new_T34_native_pass_claimed'],True)
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_reuse_gate_hash_tamper_refused(self):
        self.change('native_reuse',['fresh_public_download','sha256'],'0'*64)
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_prior_raw_scope_hash_tamper_refused(self):
        self.change('native_reuse',['prior_native','raw_sha256'],'0'*64)
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_prior_native_case_inflation_refused(self):
        self.change('native_reuse',['downloaded_windows_cases'],26)
        with self.assertRaises(ValueError):self.p().acceptance()
    def test_package_false_check_refused(self):
        data=deepcopy(self.g['package_review']);data['checks'][0]['pass']=False
        q=self.base/'package.json';q.write_text(json.dumps(data));d=deepcopy(self.g['published_download'])
        d['package_audit'].update(path=str(q),sha256=module.sha(q.read_bytes()));p=self.base/'download.json';p.write_text(json.dumps(d))
        next(x for x in self.value['acceptance_gates']if x['role']=='published_download').update(source=str(p),raw_sha256=module.sha(p.read_bytes()))
        with self.assertRaisesRegex(ValueError,'Every exact fresh package'):self.p().acceptance()
    def test_package_native_or_application_scope_refused(self):
        for field in ('application_executed','native_engines_executed'):
            data=deepcopy(self.g['package_review']);data[field]=True
            q=self.base/'package-scope.json';q.write_text(json.dumps(data));d=deepcopy(self.g['published_download'])
            d['package_audit'].update(path=str(q),sha256=module.sha(q.read_bytes()));p=self.base/'download-scope.json';p.write_text(json.dumps(d))
            next(x for x in self.value['acceptance_gates']if x['role']=='published_download').update(source=str(p),raw_sha256=module.sha(p.read_bytes()))
            with self.assertRaisesRegex(ValueError,'Every exact fresh package'):self.p().acceptance()
    def test_changed_package_child_hash_refused(self):
        self.change('published_download',['package_audit','sha256'],'0'*64)
        with self.assertRaisesRegex(ValueError,'Original closure gate hash'):self.p().acceptance()
    def test_curated_nul_exact_bindings(self):
        raw=b'docs/codex/A.md\0docs/codex/B.md\0';q=self.base/'paths.txt';q.write_bytes(raw)
        argv=['git','diff','--name-only','-z',module.R,module.M]
        command={'argv':argv,'exit_code':0,'streams':{'stdout':{'path':q.name,'sha256':module.sha(raw),'bytes':len(raw)}}}
        ledger=self.base/'ledger.json';ledger.write_text(json.dumps([command]))
        row={'source':str(q),'raw_sha256':module.sha(raw),'raw_bytes':len(raw),'entry_count':2,'kind':'git_diff_names_z','all_entries_are_docs_codex':True,'argv':argv,'ledger':{'source':str(ledger),'raw_sha256':module.sha(ledger.read_bytes()),'command_index':0}}
        self.value['curated_git_z_omissions']=[row];p=self.p();p.validate_curated_omissions();p.omit_curated(q);self.assertEqual(len(p.omitted),1)
        for key,bad in [('raw_sha256','0'*64),('raw_bytes',True),('entry_count',True),('all_entries_are_docs_codex',False),('argv',['git','diff','--name-only','-z',module.R,'0'*40])]:
            self.value['curated_git_z_omissions']=[{**row,key:bad}]
            with self.assertRaises(ValueError):self.p().validate_curated_omissions()
    def test_bare_action_preserves_complete_typed_object(self):
        private=Path(__import__('os').environ['USERPROFILE']).name
        v={'id':123,'head_sha':'a'*40,'head_commit':{'id':'a'*40,'tree_id':'b'*40,'message':'Recorded','timestamp':'2026-10-10T00:00:00Z','author':{'name':'Author','email':private+'@private.invalid'},'committer':{'name':'Committer','email':private+'@private.invalid'}},'event':'pull_request','status':'completed','conclusion':'success','url':'https://api.github.com/example','html_url':'https://github.com/example','bool':True,'number':3,'null':None}
        raw=b'\xef\xbb\xbf'+json.dumps(v).encode()
        self.value['github_metadata_identity_receipts']=[{'path':'api.txt','kind':'actions_run','raw_sha256':module.sha(raw),'git_commit_sha':'a'*40,'run_id':123,'page_index':0,'run_index':0}]
        p=self.p();public=p.github_metadata(raw,'api.txt');self.assertTrue(public.startswith(b'\xef\xbb\xbf'))
        v['head_commit']['author']['email']='<EMAIL>';v['head_commit']['committer']['email']='<EMAIL>'
        self.assertEqual(json.loads(public.decode('utf-8-sig')),v)
    def test_bare_action_nonzero_index_refused(self):
        self.value['github_metadata_identity_receipts']=[{'path':'api.txt','kind':'actions_run','raw_sha256':'a'*64,'git_commit_sha':'a'*40,'run_id':123,'page_index':0,'run_index':1}]
        with self.assertRaises(ValueError):self.p()

if __name__=='__main__':unittest.main()
