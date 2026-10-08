"""Synthetic privacy-copy tests only; no application/native/export invocation."""
from pathlib import Path
import hashlib,importlib.util,json,shutil,unittest,uuid,xml.etree.ElementTree as ET
from unittest import mock
repo=Path.cwd().resolve();source=repo/'tests/.work/Export-T19Evidence.py';root=repo/'tests/.work'/('T19-exporter-privacy-tests-'+uuid.uuid4().hex);root.mkdir()
for path in [source,Path(__file__)]:shutil.copyfile(path,root/path.name)
spec=importlib.util.spec_from_file_location('exporter',source);module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
rows=[(['repo'],r'Z:\fictional\Repo','<REPO>','path'),(['userprofile'],r'Z:\Users\fixture-user','<USERPROFILE>','path'),
      (['localappdata'],r'Z:\Users\fixture-user\AppData\Local','%LOCALAPPDATA%','path'),(['appdata'],r'Z:\Users\fixture-user\AppData\Roaming','%APPDATA%','path'),
      (['username'],'fixture-user','<USERNAME>','identity'),(['computername'],'fixture-host','<COMPUTERNAME>','identity'),(['userdomain'],'fixture-domain','<USERDOMAIN>','identity')]
document={'schema_version':1,'task':'T19','replacements':[{'fields':fields,'value':value,'token':token,'kind':kind} for fields,value,token,kind in rows]}
with mock.patch.object(module,'load',return_value=document),mock.patch.object(Path,'read_bytes',return_value=json.dumps(document).encode()):privacy=module.PrivacyMap(root/'never-written-fake-map.json')

class PrivacyTests(unittest.TestCase):
    def test_longest_specific_path_precedes_profile_and_username(self):
        value=(r'Z:\Users\fixture-user\AppData\Local\cache'+ '\n').encode()
        public,counts=privacy.copy_observation(value)
        self.assertEqual(public,b'%LOCALAPPDATA%\\cache\n');self.assertEqual(counts,{'%LOCALAPPDATA%':1})
    def test_native_forward_json_and_nested_json_variants(self):
        base=rows[0][1]
        for value in [base,base.replace('\\','/'),base.replace('\\','\\\\'),base.replace('\\','\\\\\\\\'),base.replace('\\','\\/')]:
            with self.subTest(value=value):self.assertEqual(privacy.text(value),'<REPO>')
    def test_json_copy_remains_json_and_preserves_nonidentity_facts(self):
        raw=json.dumps({'source':rows[0][1]+r'\synthetic.pdf','exit':0,'pages':4,'bytes':12345,'ordered_ids':['T19-A-P1','T19-B-P2']}).encode()
        facts=json.loads(privacy.raw_copy(raw));self.assertEqual(facts['source'],r'<REPO>\synthetic.pdf')
        self.assertEqual({key:facts[key] for key in ['exit','pages','bytes','ordered_ids']},{'exit':0,'pages':4,'bytes':12345,'ordered_ids':['T19-A-P1','T19-B-P2']})
    def test_xml_identity_fields_and_stream_order_are_preserved(self):
        raw=b'<environment user="fixture-user" machine="FIXTURE-HOST" domain="fixture-domain" />\r\nstdout-first\r\nstderr-second\r\n'
        self.assertEqual(privacy.raw_copy(raw,is_xml=True),b'<environment user="&lt;USERNAME&gt;" machine="&lt;COMPUTERNAME&gt;" domain="&lt;USERDOMAIN&gt;" />\r\nstdout-first\r\nstderr-second\r\n')
    def test_xml_public_copy_remains_parseable_with_original_counts(self):
        raw=b'<?xml version="1.0"?><test-results total="6" failures="0"><environment user="fixture-user" machine="FIXTURE-HOST" /><case path="Z:\\fictional\\Repo\\case.pdf" /></test-results>\r\n'
        root=ET.fromstring(privacy.raw_copy(raw,is_xml=True))
        self.assertEqual(root.attrib,{'total':'6','failures':'0'});self.assertEqual(root.find('environment').attrib,{'user':'<USERNAME>','machine':'<COMPUTERNAME>'})
        self.assertEqual(root.find('case').attrib['path'],r'<REPO>\case.pdf')
    def test_no_identity_recipe_is_byte_exact(self):
        raw=b'print("T19-A-P1")\r\n# original fixture source\r\n';public,counts=privacy.copy_observation(raw)
        self.assertEqual(public,raw);self.assertEqual(counts,{})
    def test_encoding_bom_and_crcrlf_survive_sanitization(self):
        raw=b'\xff\xfe'+('fixture-user\r\r\nexit=0\r\n').encode('utf-16-le')
        self.assertEqual(privacy.raw_copy(raw),b'\xff\xfe'+('<USERNAME>\r\r\nexit=0\r\n').encode('utf-16-le'))
    def test_identity_boundary_avoids_changes_to_unrelated_words(self):
        self.assertEqual(privacy.text('fixture-user-extra fixture-hosting fixture-domainish'),'fixture-user-extra fixture-hosting fixture-domainish')
    def test_public_manifest_labels_and_binary_paths_are_sanitized(self):
        value={'copy_from':rows[0][1]+r'\case.json','original':rows[1][1]+r'\cache.dll','source_sha256':'a'*64,'bytes':123}
        result=privacy.labels(value);self.assertEqual(result['copy_from'],r'<REPO>\case.json');self.assertEqual(result['original'],r'<USERPROFILE>\cache.dll')
        self.assertEqual(result['source_sha256'],'a'*64);self.assertEqual(result['bytes'],123)

suite=unittest.defaultTestLoader.loadTestsFromTestCase(PrivacyTests);runner=unittest.TextTestRunner(verbosity=2);result=runner.run(suite)
receipt={'task':'T19','scope':'Synthetic ignored exporter privacy-copy tests only; no actual application/native/manual/export acceptance.',
         'result':'pass' if result.wasSuccessful() else 'fail','tests':result.testsRun,'failures':len(result.failures),'errors':len(result.errors),
         'exporter_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'producer_sha256':hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
         'clear_actual_environment_identity_used':False,'public_export_performed':False}
path=root/'receipt.json';path.write_text(json.dumps(receipt,indent=2)+'\n');print(json.dumps({'receipt':str(path),'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),**receipt}))
raise SystemExit(0 if result.wasSuccessful() else 1)
