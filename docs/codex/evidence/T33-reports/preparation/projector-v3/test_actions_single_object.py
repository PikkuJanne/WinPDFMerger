"""Single-object GitHub metadata developer tests; no API/remote operations."""
from copy import deepcopy
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
HERE=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('T33_flat_actions',HERE/'Export-T33.py')
p=importlib.util.module_from_spec(spec);spec.loader.exec_module(p)

class SingleObjectActions(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.config=Path(self.tmp.name)/'config.json'
        private=Path(os.environ['USERPROFILE']).name
        person={'name':'Recorded author','email':private+'@private.invalid'}
        self.data={'total_count':1,'workflow_runs':[{'id':123,'head_sha':'a'*40,'head_commit':{'id':'a'*40,'tree_id':'b'*40,
            'message':'Recorded original message','timestamp':'2026-10-10T00:00:00Z','author':deepcopy(person),'committer':deepcopy(person)},
            'event':'push','status':'completed','conclusion':'success','url':'https://api.github.com/example','html_url':'https://github.com/example',
            'bool':True,'false':False,'null':None,'integer':4,'number':7.25}]}
    def projector(self,value=None,bom=False,**pins):
        value=self.data if value is None else value
        raw=(b'\xef\xbb\xbf'if bom else b'')+json.dumps(value).encode()
        declaration={'path':'flat.stdout.txt','kind':'actions_runs','raw_sha256':p.sha(raw),'git_commit_sha':'a'*40,'run_id':123,'page_index':0,'run_index':0,**pins}
        self.config.write_text(json.dumps({'schema_version':1,'task':'T33','source_commit':p.R,'github_metadata_identity_receipts':[declaration]}))
        return p.Projector(HERE.parents[2],self.config),raw
    def test_only_two_pinned_email_fields_change_object_type_preserved(self):
        exporter,raw=self.projector();actual=json.loads(exporter.github_metadata(raw,'flat.stdout.txt'))
        expected=deepcopy(self.data)
        for role in ['author','committer']:expected['workflow_runs'][0]['head_commit'][role]['email']='<EMAIL>'
        self.assertEqual(actual,expected);self.assertIs(type(actual),dict)
    def test_all_bool_number_null_types_and_BOM_preserved(self):
        exporter,raw=self.projector(bom=True);projected=exporter.github_metadata(raw,'flat.stdout.txt')
        self.assertTrue(projected.startswith(b'\xef\xbb\xbf'))
        actual=json.loads(projected.decode('utf-8-sig'))['workflow_runs'][0]
        self.assertEqual([type(actual[k])for k in ['bool','false','null','integer','number']],[bool,bool,type(None),int,float])
    def test_list_cannot_masquerade_as_single_object(self):
        exporter,raw=self.projector([self.data])
        with self.assertRaises(ValueError):exporter.github_metadata(raw,'flat.stdout.txt')
    def test_page_index_must_be_typed_zero(self):
        for value in [1,True]:
            with self.subTest(value=value),self.assertRaises(ValueError):self.projector(page_index=value)
    def test_wrong_run_and_head_and_unknown_schema_rejected(self):
        for mutate in [lambda x:x['workflow_runs'][0].update(id=124),lambda x:x['workflow_runs'][0]['head_commit'].update(id='d'*40),
                       lambda x:x.update(total_count=True),lambda x:x['workflow_runs'][0]['head_commit']['author'].pop('name')]:
            data=deepcopy(self.data);mutate(data);exporter,raw=self.projector(data)
            with self.assertRaises(ValueError):exporter.github_metadata(raw,'flat.stdout.txt')
    def test_raw_hash_change_is_refused(self):
        exporter,raw=self.projector()
        with self.assertRaises(ValueError):exporter.github_metadata(raw+b' ','flat.stdout.txt')
    def test_unknown_private_identity_elsewhere_still_refused(self):
        data=deepcopy(self.data);data['extra']=Path(os.environ['USERPROFILE']).name
        exporter,raw=self.projector(data);projected=exporter.github_metadata(raw,'flat.stdout.txt')
        source=Path(self.tmp.name)/'flat.json';source.write_bytes(raw)
        exporter.selected={'flat.stdout.txt':(source,'Synthetic narrow privacy test')}
        with self.assertRaises(ValueError):exporter.payloads()

if __name__=='__main__':unittest.main(verbosity=2)
