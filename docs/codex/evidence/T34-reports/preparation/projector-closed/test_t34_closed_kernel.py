"""Developer-only T34 projection guards; no release/native/closure acceptance."""
from pathlib import Path
import json, os, tempfile, unittest
from test_projector import module, root

class T34Kernel(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(dir=root,prefix='synthetic-');self.addCleanup(self.tmp.cleanup)
        self.base=Path(self.tmp.name);self.config=self.base/'config.json'
        self.value={'schema_version':1,'task':'T34','source_commit':module.R,'roots':[],
            'asset_sha256':{'zip':'2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2','checksums':'d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'},
            'acceptance_gates':[],'github_metadata_identity_receipts':[]}
    def p(self):
        self.config.write_text(json.dumps(self.value),encoding='utf-8')
        return module.Projector(root.parents[2],self.config)
    def test_prior_task_aliases_refused(self):
        for kind in ('t32-source','t32-artifacts','t33-public-download'):
            self.value['local_path_aliases']=[{'original':'C:/projects/WinPDFMerger-'+kind+'-'+'a'*32,'alias':'<T34_PUBLIC_DOWNLOAD>','provenance':'synthetic alias refusal'}]
            with self.assertRaises(ValueError):self.p()
    def test_t34_escaped_aliases(self):
        original='C:'+chr(92)+'projects'+chr(92)+'WinPDFMerger-t34-public-download-'+'a'*32
        self.value['local_path_aliases']=[{'original':original,'alias':'<T34_PUBLIC_DOWNLOAD>','provenance':'synthetic exact alias'}]
        p=self.p()
        for depth in range(5):
            for value in (original.replace('\\','\\'*(2**depth)),original.replace('\\','/').replace('/','\\'*(2**depth)+'/')):
                self.assertEqual(p.replace(value),'<T34_PUBLIC_DOWNLOAD>')
    def test_unknown_t34_data_path_detected(self):
        value='C:/projects/WinPDFMerger-t34-public-download-'+'b'*32+'/file.txt'
        self.assertTrue(module.private_windows_task_path(value))
    def test_previous_owned_receipt_refused(self):
        p=self.p()
        with self.assertRaises(ValueError):p.owned(root.parent/'T33-export-preparation-v3/Export-T33.py')
    def test_no_automatic_historical_diff_skip(self):
        q=self.base/'audited-packet-diff.stdout.txt';q.write_text('owned synthetic text')
        p=self.p();p.choose(q,'ledger/text.txt','Synthetic retained text')
        self.assertEqual(len(p.selected),1);self.assertEqual(p.omitted,[])
    def test_unknown_actual_interfaces_stay_closed(self):
        with self.assertRaisesRegex(ValueError,'not bound'):self.p().acceptance()
    def test_unbound_text_omission_refused(self):
        q=self.base/'fixture.txt';q.write_text('owned synthetic text')
        self.value['roots']=[{'source':str(self.base),'label':'fixture','role':'preparation','mode':'flat','scope':'synthetic','provenance':'synthetic','exclude':['fixture.txt']}]
        p=self.p();p.acceptance=lambda:None
        with self.assertRaisesRegex(ValueError,'exact raw-hash'):p.select()
    def test_nul_text_refused(self):
        q=self.base/'fixture.txt';q.write_bytes(b'unsafe\0text')
        p=self.p();p.selected={'fixture.txt':(q,'Synthetic NUL rejection')}
        with self.assertRaisesRegex(ValueError,'NUL'):p.payloads()
    def test_dry_no_creation_and_immutable_local_fixture(self):
        q=self.base/'fixture.json';q.write_text('{"bool":true,"number":3,"null":null}')
        p=self.p();p.select=lambda:None;p.selected={'fixture.json':(q,'Synthetic immutable byte fixture')}
        target=self.base/'prospective'
        result=p.run(target);self.assertEqual(result['mode'],'dry_run');self.assertFalse(target.exists())
        first=p.run(target,True);self.assertTrue(target.is_dir())
        self.assertEqual(p.run(target,True)['manifest_sha256'],first['manifest_sha256'])
        (target/'fixture.json').write_text('modified')
        with self.assertRaisesRegex(ValueError,'cannot be rewritten'):p.run(target,True)

if __name__=='__main__':unittest.main()
