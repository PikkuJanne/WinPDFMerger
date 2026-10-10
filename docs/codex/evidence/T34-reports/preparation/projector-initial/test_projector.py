"""Synthetic T34 projection checks, never application/native/platform acceptance."""
from copy import deepcopy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
import xml.etree.ElementTree as ET

root=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('projector',root/'Export-T34.py')
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)

class Tests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.config=Path(self.tmp.name)/'config.json'
        self.private=Path(os.environ['USERPROFILE']).name
        self.value={'schema_version':1,'task':'T34','source_commit':module.R,'roots':[],
                    'github_metadata_identity_receipts':[]}
        self.repo=root.parents[2]
    def projector(self):
        self.config.write_text(json.dumps(self.value),encoding='utf-8')
        return module.Projector(self.repo,self.config)
    def pages(self):
        person={'name':'Recorded author','email':self.private+'@private.invalid'}
        return [{'total_count':1,'workflow_runs':[{'id':123,'head_sha':'a'*40,'head_commit':{
            'id':'a'*40,'tree_id':'b'*40,'message':'Recorded message','timestamp':'2026-10-10T00:00:00Z',
            'author':deepcopy(person),'committer':deepcopy(person)},'event':'push','status':'completed',
            'conclusion':'success','url':'https://api.github.com/example','html_url':'https://github.com/example',
            'null':None,'bool':True,'number':7.25}]}]
    def tag(self):
        return {'sha':'c'*40,'node_id':'recorded','tag':'v1.0.0','message':'Recorded message',
                'tagger':{'name':'Recorded tagger','email':self.private+'@private.invalid','date':'2026-10-10T00:00:00Z'},
                'object':{'type':'commit','sha':module.R},'verification':{'verified':False,'payload':None}}
    def registered(self,kind,value,bom=False):
        raw=(b'\xef\xbb\xbf' if bom else b'')+json.dumps(value).encode()
        row={'path':'metadata.stdout.txt','raw_sha256':module.sha(raw),'kind':kind}
        if kind=='actions_run_pages':row.update(git_commit_sha='a'*40,run_id=123,page_index=0,run_index=0)
        else:row.update(tag_object_sha='c'*40,target_commit_sha=module.R,tag='v1.0.0')
        self.value['github_metadata_identity_receipts']=[row]
        return self.projector(),raw
    def test_types(self):
        p=self.projector();v={'bool':True,'false':False,'null':None,'integer':8,'number':2.5,'text':str(self.repo)}
        projected=p.typed(v);self.assertEqual({k:type(x) for k,x in projected.items()},{k:type(x) for k,x in v.items()})
        self.assertEqual(projected['text'],'<REPO>')
    def test_json_bom(self):
        self.assertTrue(self.projector().encode_json({'pass':True},b'\xef\xbb\xbf{}').startswith(b'\xef\xbb\xbf'))
    def test_no_added_bom(self):self.assertFalse(self.projector().encode_json({},b'{}').startswith(b'\xef\xbb\xbf'))
    def test_path_variants(self):
        p=self.projector();original=str(self.repo);forward=original.replace('\\','/')
        for depth in range(5):
            for v in [original.replace('\\','\\'*(2**depth)),forward.replace('/','\\'*(2**depth)+'/')]:
                self.assertEqual(p.replace(v),'<REPO>')
    def test_xml(self):
        p=self.projector();raw='<test cwd="'+str(self.repo)+'"><environment user="private" machine-name="private" user-domain="private" /></test>'
        parsed=ET.fromstring(p.xml(raw));self.assertEqual(parsed.attrib['cwd'],'<REPO>')
        self.assertEqual(parsed[0].attrib,{'user':'<USER>','machine-name':'<COMPUTER>','user-domain':'<COMPUTER>'})
    def test_xml_bom(self):self.assertTrue(self.projector().xml('\ufeff<test />').startswith('\ufeff'))
    def test_email_only(self):
        v=self.pages();p,raw=self.registered('actions_run_pages',v);a=json.loads(p.github_metadata(raw,'metadata.stdout.txt'))
        e=deepcopy(v);e[0]['workflow_runs'][0]['head_commit']['author']['email']='<EMAIL>';e[0]['workflow_runs'][0]['head_commit']['committer']['email']='<EMAIL>'
        self.assertEqual(a,e)
    def test_tag_only(self):
        v=self.tag();p,raw=self.registered('annotated_tag',v);a=json.loads(p.github_metadata(raw,'metadata.stdout.txt'));v['tagger']['email']='<EMAIL>';self.assertEqual(a,v)
    def test_metadata_bom(self):
        p,raw=self.registered('annotated_tag',self.tag(),True);self.assertTrue(p.github_metadata(raw,'metadata.stdout.txt').startswith(b'\xef\xbb\xbf'))
    def test_wrong_raw(self):
        p,raw=self.registered('annotated_tag',self.tag())
        with self.assertRaises(ValueError):p.github_metadata(raw+b' ','metadata.stdout.txt')
    def test_wrong_run(self):
        v=self.pages();v[0]['workflow_runs'][0]['id']=124;p,raw=self.registered('actions_run_pages',v)
        with self.assertRaises(ValueError):p.github_metadata(raw,'metadata.stdout.txt')
    def test_boolean_run(self):
        v=self.pages();v[0]['workflow_runs'][0]['id']=True;p,raw=self.registered('actions_run_pages',v)
        with self.assertRaises(ValueError):p.github_metadata(raw,'metadata.stdout.txt')
    def test_wrong_head(self):
        v=self.pages();v[0]['workflow_runs'][0]['head_commit']['id']='d'*40;p,raw=self.registered('actions_run_pages',v)
        with self.assertRaises(ValueError):p.github_metadata(raw,'metadata.stdout.txt')
    def test_unknown_pages(self):
        v=self.pages()[0];p,raw=self.registered('actions_run_pages',v)
        with self.assertRaises(ValueError):p.github_metadata(raw,'metadata.stdout.txt')
    def test_wrong_tag_object(self):
        v=self.tag();v['sha']='d'*40;p,raw=self.registered('annotated_tag',v)
        with self.assertRaises(ValueError):p.github_metadata(raw,'metadata.stdout.txt')
    def test_wrong_tag_target(self):
        v=self.tag();v['object']['sha']='d'*40;p,raw=self.registered('annotated_tag',v)
        with self.assertRaises(ValueError):p.github_metadata(raw,'metadata.stdout.txt')
    def test_public_email_unmodified_refusal(self):
        v=self.tag();v['tagger']['email']='public@example.org';p,raw=self.registered('annotated_tag',v)
        with self.assertRaises(ValueError):p.github_metadata(raw,'metadata.stdout.txt')
    def test_bad_source(self):
        self.value['source_commit']='d'*40
        with self.assertRaises(ValueError):self.projector()
    def test_bad_task(self):
        self.value['task']='T31'
        with self.assertRaises(ValueError):self.projector()
    def test_unsafe_labels(self):
        for label in ('../private','/private','C:/private','root\\file','root//file'):
            with self.assertRaises(ValueError):module.label_safe(label)
    def test_owned_outside(self):
        with self.assertRaises(ValueError):self.projector().owned(self.config)
    def test_extra_alias(self):
        original='C:/projects/WinPDFMerger-t34-public-download-'+'a'*32
        self.value['local_path_aliases']=[{'original':original,'alias':'<T34_SOURCE>','provenance':'synthetic exact source'}]
        self.assertEqual(self.projector().replace(original),'<T34_SOURCE>')
    def test_broad_alias_refused(self):
        self.value['local_path_aliases']=[{'original':'C:/projects','alias':'<T34_SOURCE>','provenance':'invalid broad'}]
        with self.assertRaises(ValueError):self.projector()
    def test_real_junction_parent(self):
        import subprocess
        base=Path(self.tmp.name); target=base/'target';target.mkdir();link=base/'junction'
        run=subprocess.run(['cmd.exe','/d','/c','mklink','/J',str(link),str(target)],capture_output=True)
        self.assertEqual(run.returncode,0,'Local synthetic junction fixture creation failed')
        try:
            with self.assertRaises(ValueError):module.ordinary_ancestors(link/'future'/'payload.txt')
            with self.assertRaises(ValueError):module.ordinary_ancestors(link)
        finally:
            os.rmdir(link)
        self.assertTrue(target.is_dir())
    def test_ordinary_new_ancestors(self):
        module.ordinary_ancestors(Path(self.tmp.name)/'new'/'payload.txt')
    def test_original_destination_checked_before_resolve(self):
        import inspect
        code=inspect.getsource(module.main)
        self.assertLess(code.index('ordinary_ancestors(destination)'),code.index('destination = destination.resolve()'))
    def test_write_checks_target_before_after_mkdir(self):
        import inspect
        code=inspect.getsource(module.Projector.run)
        self.assertGreaterEqual(code.count('ordinary_ancestors(target)'),3)
    def test_synthetic_code_prefix_is_not_private_path(self):
        example="original='C:/"+"projects/WinPDFMerger-t32-source-'+'a'*32"
        self.assertFalse(module.private_windows_task_path(example))
    def test_complete_undeclared_task_uuid_refused(self):
        actual='C:/'+'projects/WinPDFMerger-t32-source-'+'a'*32+'/README.md'
        self.assertTrue(module.private_windows_task_path(actual))
        self.assertTrue(module.private_windows_task_path('C:/'+'Users/private/file.txt'))
    def test_literal_source_probe_is_permitted_but_data_refused(self):
        fake='C:/'+'projects/WinPDFMerger-t32-source-'+'a'*32+'/README.md'
        base=Path(self.tmp.name);source=base/'probe.py';data=base/'probe.txt'
        for item in (source,data):item.write_text("path="+repr(fake),encoding='utf-8')
        p=self.projector();p.selected={'probe.py':(source,'Synthetic source predicate probe')};_,rows=p.payloads();self.assertEqual(len(rows),1)
        p.selected={'probe.txt':(data,'Synthetic unknown data path')}
        with self.assertRaises(ValueError):p.payloads()
    def test_source_current_private_identity_still_refused(self):
        source=Path(self.tmp.name)/'probe.py';source.write_text('identity='+repr(self.private),encoding='utf-8')
        p=self.projector();p.selected={'probe.py':(source,'Synthetic private source identity')}
        with self.assertRaises(ValueError):p.payloads()
    def test_missing_final_gate(self):
        self.value['acceptance_gates']=[]
        with self.assertRaises(ValueError):self.projector().acceptance()

if __name__=='__main__':unittest.main()
