"""Synthetic acceptance-interface regressions, not release/native execution."""
from copy import deepcopy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

HERE=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('T33_projector_gates',HERE/'Export-T33.py')
p=importlib.util.module_from_spec(spec);spec.loader.exec_module(p)
ZIP='2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2'
SUM='d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'
DOWNLOAD=Path('C:/'+'projects/WinPDFMerger-t33-public-download-'+'a'*32)

class T33ActualGateSchema(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.root=Path(self.tmp.name)
        self.p=p.Projector.__new__(p.Projector)
        self.p.aliases=[(str(DOWNLOAD),'<T33_PUBLIC_DOWNLOAD>')]
        self.p.guards=[]
        self.records={}
        self.children={}
        common={'task':'T33','source_commit':p.R,'harness_commit':p.M,'issues':[]}
        hashes={'zip_sha256':ZIP,'checksums_sha256':SUM}
        self.records['accepted_publication']={**common,'result':'pass_for_verified_final_publication','mode':'publish','publication_attempted':True,'evidence_commit':p.M,
            'draft':False,'prerelease':False,'live_peeled_commit':p.R,'tag_object_sha':'7818645de07b902ad8f2b815e90ee1d74d2724d6',
            'release_id':408603768,'published_at':'2026-10-10T07:12:01Z','release_url':'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0',
            'assets':{'WinPDFMerger-v1.0.0.zip':[193669,'sha256:'+ZIP],'SHA256SUMS.txt':[90,'sha256:'+SUM]}}
        self.records['accepted_public_download']={**common,**hashes,'result':'pass_for_unauthenticated_published_release_and_independent_download',
            **{k:False for k in ['draft','prerelease','authentication_used','cookies_used','gh_download_used','download_directory_previously_existed','remote_mutations','application_executed']},
            'download_directory':str(DOWNLOAD),'tag_object_sha':'7818645de07b902ad8f2b815e90ee1d74d2724d6','release_id':408603768,
            'published_at':'2026-10-10T07:12:01Z','release_url':'https://github.com/PikkuJanne/WinPDFMerger/releases/tag/v1.0.0',
            'manual_acceptance':'excluded/unperformed; never pass','package_audit':{'sha256':'','checks':266,'issues':[]}}
        self.records['accepted_package_review']={**common,'result':'pass_for_exact_published_download_package_bytes',
            'recorded_expected_zip_sha256':ZIP,'recorded_expected_checksums_sha256':SUM,'checks_total':266}
        self.records['accepted_operation_review']={**common,**hashes,'result':'pass','checks':5673,'application_cases':25,
            'independent_pdf_count':21,'independent_pdf_pages':106,'manual_acceptance':'excluded/unperformed',
            'download_directory':str(DOWNLOAD),'public_download_report':{'sha256':''}}
        self.records['accepted_decoded_image_review']={**common,'candidate_source_commit':p.R,'result':'pass',
            'retained_pdf_count':21,'retained_output_page_count':106,'checks':1360}
        self.records['accepted_native']={**common,'evidence_class':p.NATIVE_CLASS,'result':'pass','source_clean_before_after':True,
            'driver_unchanged':True,'cache_and_assets_unchanged':True,'approved_cache_files_verified':348,
            'manual_acceptance':'excluded/unperformed; never pass','shared_assets':{**hashes,'zip_path':str(DOWNLOAD/'WinPDFMerger-v1.0.0.zip'),'checksums_path':str(DOWNLOAD/'SHA256SUMS.txt')},'candidate_reports':[]}
        guards=['expected_head','status_unchanged','clean','driver_unchanged','approved_cache_unchanged','candidate_assets_unchanged','parent_environment_unchanged']
        for shell,count in [('PS51',14),('PS7',11)]:
            self.children[shell]={'task':'T33','evidence_class':p.NATIVE_CLASS,'result':'pass','preparation':False,'candidate_source_commit':p.R,
                'harness_commit':p.M,'shell_kind':shell,'candidate':deepcopy(self.records['accepted_native']['shared_assets']),
                'cases':[{'package_guard':True,'source_foreign_guard':True}for _ in range(count)],'source_guard':{k:True for k in guards},
                'public_help':{'exit_code':0,'package_unchanged':True,'no_outputs':True},'manual_acceptance':'excluded/unperformed; never pass'}
        self.p.owned=lambda value:self.root/Path(value).name

    def write(self):
        def put(name,value):
            target=self.root/(name+'.json');target.write_text(json.dumps(value),encoding='utf-8')
            return {'source':target.name,'raw_sha256':p.sha(target.read_bytes())}
        child=[]
        for shell,value in self.children.items():
            row=put(shell,value);child.append({'shell':shell,'path':row['source'],'sha256':row['raw_sha256']})
        self.records['accepted_native']['candidate_reports']=child
        package=put('accepted_package_review',self.records['accepted_package_review'])
        self.records['accepted_public_download']['package_audit']['sha256']=package['raw_sha256']
        download=put('accepted_public_download',self.records['accepted_public_download'])
        self.records['accepted_operation_review']['public_download_report']['sha256']=download['raw_sha256']
        rows=[{'role':role,**put(role,value)}for role,value in self.records.items()]
        self.p.config={'acceptance_gates':rows,'asset_sha256':{'zip':ZIP,'checksums':SUM}}
    def accept(self):
        self.write();self.p.acceptance()
    def test_all_six_actual_interface_roles_accept_typed_synthetic_facts(self):
        self.accept();self.assertEqual(len(self.p.guards),6)
    def test_historical_draft_or_build_roles_cannot_substitute(self):
        self.write();self.p.config['acceptance_gates'][0]['role']='accepted_tag_draft'
        with self.assertRaises(ValueError):self.p.acceptance()
    def test_authenticated_or_reused_download_is_rejected(self):
        for key in ['authentication_used','cookies_used','gh_download_used','download_directory_previously_existed']:
            with self.subTest(key=key):
                self.records['accepted_public_download'][key]=True
                with self.assertRaises(ValueError):self.accept()
                self.records['accepted_public_download'][key]=False
    def test_native_matching_hashes_from_another_path_are_rejected(self):
        self.records['accepted_native']['shared_assets']['zip_path']=str(self.root/'WinPDFMerger-v1.0.0.zip')
        with self.assertRaises(ValueError):self.accept()
    def test_native_child_wrong_path_or_guard_rejected(self):
        for mutate in [lambda x:x['candidate'].update(checksums_path='wrong-sum.txt'),lambda x:x['source_guard'].update(clean=False)]:
            original=deepcopy(self.children['PS7']);mutate(self.children['PS7'])
            with self.assertRaises(ValueError):self.accept()
            self.children['PS7']=original
    def test_publication_array_sizes_and_digests_exact(self):
        self.records['accepted_publication']['assets']['SHA256SUMS.txt'][0]=91
        with self.assertRaises(ValueError):self.accept()
    def test_draft_metadata_cannot_pass_publication(self):
        self.records['accepted_publication']['draft']=True
        with self.assertRaises(ValueError):self.accept()
    def test_operation_requires_original_download_report_binding(self):
        self.write();role=next(x for x in self.p.config['acceptance_gates']if x['role']=='accepted_operation_review')
        data=json.loads((self.root/role['source']).read_text());data['public_download_report']['sha256']='b'*64
        (self.root/role['source']).write_text(json.dumps(data));role['raw_sha256']=p.sha((self.root/role['source']).read_bytes())
        with self.assertRaises(ValueError):self.p.acceptance()
    def test_actual_original_gate_hash_tampering_rejected(self):
        self.write();row=self.p.config['acceptance_gates'][0]
        (self.root/row['source']).write_bytes((self.root/row['source']).read_bytes()+b' ')
        with self.assertRaises(ValueError):self.p.acceptance()
    def test_decoded_wrong_R_or_pages_rejected(self):
        self.records['accepted_decoded_image_review']['retained_output_page_count']=105
        with self.assertRaises(ValueError):self.accept()
    def test_download_alias_is_exact_and_required(self):
        self.p.aliases=[]
        with self.assertRaises(ValueError):self.accept()
    def test_additional_exact_public_download_alias_projects_all_escaped_variants(self):
        config=self.root/'config.json';config.write_text(json.dumps({'schema_version':1,'task':'T33','source_commit':p.R,
            'local_path_aliases':[{'original':str(DOWNLOAD),'alias':'<T33_PUBLIC_DOWNLOAD>','provenance':'synthetic exact download scope'}]}))
        projector=p.Projector(HERE.parents[2],config)
        for depth in range(5):
            value=str(DOWNLOAD).replace('\\','\\'*(2**depth))
            self.assertEqual(projector.replace(value),'<T33_PUBLIC_DOWNLOAD>')

if __name__=='__main__':unittest.main(verbosity=2)
