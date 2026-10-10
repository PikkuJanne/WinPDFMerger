"""Capture one authorized exact-published-download T33 native invocation."""
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import uuid

REPO=Path(__file__).resolve().parents[3]
HERE=Path(__file__).resolve().parent
R='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
M='f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'
ZIP='2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2'
SUM='d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'
DOWNLOAD=Path('<T33_PUBLIC_DOWNLOAD>')
PREP=REPO/'tests/.work/T33-operation-preparation'
PUBLIC=REPO/'tests/.work/T33-public-download-review/actual-054fec6b5bc445beb4dd71d4fbf017aa'
BINDINGS=[(PUBLIC/'public-download-review.json','0f5f94923e47d0ae5bcc247df9d5ba56e02b39b1513977b32b27b3c1727faa19'),
          (PUBLIC/'downloaded-package-byte-audit.json','f9b277562ce359443a09a2ed2242398dc054659c46c74415775475a87611172d'),
          (REPO/'tests/.work/T33-public-download-review/actions/independent-unauthenticated-published-download-42c3a95a78dd424c962654da306d58e6/receipt.json','acfb9a5843e81c64f1291df1409dd0a6965b8de7066195578370ca48f6a8b4e5'),
          (PREP/'capture-T33.py','01cf9fa377dde5719fea16153791bc2a85fa412872ceee9cfa6029a5f270c8b7'),
          (PREP/'final_package_smoke.py','5ea6bd48c400f1ffe9becf7ba5fb2b7fd355019dbbab3a8aa295a29c727626f8')]
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def load(p):return json.loads(p.read_text(encoding='utf-8-sig'))
def git(*argv):return subprocess.check_output(['git',*argv],cwd=REPO,text=True).strip()
def guard():
    return {'clean_exact_M':git('rev-parse','HEAD')==M and not git('status','--porcelain=v1'),
            'sources_and_original_gates_unchanged':all(sha(p)==h for p,h in BINDINGS),
            'downloaded_pair_unchanged':sha(DOWNLOAD/'WinPDFMerger-v1.0.0.zip')==ZIP and sha(DOWNLOAD/'SHA256SUMS.txt')==SUM}

def main():
    gate=load(PUBLIC/'public-download-review.json')
    assert gate['task']=='T33' and gate['result']=='pass_for_unauthenticated_published_release_and_independent_download'
    assert gate['source_commit']==R and gate['harness_commit']==M and Path(gate['download_directory'])==DOWNLOAD
    assert gate['zip_sha256']==ZIP and gate['checksums_sha256']==SUM and gate['issues']==[]
    assert gate['release_id']==408603768 and gate['published_at']=='2026-10-10T07:12:01Z'
    assert all(gate[k] is False for k in ['draft','prerelease','authentication_used','cookies_used','gh_download_used'])
    assert load(BINDINGS[2][0])['exit_code']==0
    assert {p.name for p in DOWNLOAD.iterdir()}=={'WinPDFMerger-v1.0.0.zip','SHA256SUMS.txt'}
    before=guard()
    assert all(before.values()), 'Before-execution guard failed'
    root=HERE/('actual-'+uuid.uuid4().hex)
    root.mkdir(exist_ok=False)
    command=[sys.executable,'-B',str(PREP/'capture-T33.py'),'--repo',str(REPO),
             '--expected-harness-commit',M,'--zip',str(DOWNLOAD/'WinPDFMerger-v1.0.0.zip'),'--zip-sha256',ZIP,
             '--checksums',str(DOWNLOAD/'SHA256SUMS.txt'),'--checksums-sha256',SUM]
    receipt={'task':'T33','scope':'Actual automated dual-shell Windows/native operation of the exact independently anonymously downloaded published asset pair',
             'source_commit':R,'harness_commit':M,'started_at_utc':datetime.now(timezone.utc).isoformat(),
             'argv':command,'capture_source_sha256':sha(Path(__file__)),'state':'running','guard_before':before,
             'original_bindings':[{'path':str(p),'sha256':h}for p,h in BINDINGS],
             'download_directory':str(DOWNLOAD),'zip_sha256':ZIP,'checksums_sha256':SUM,
             'manual_acceptance':'AC058 excluded/nonrequired/unperformed; never pass'}
    target=root/'receipt.json'
    def save():target.write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
    save()
    failure=None;code=None
    try:
        with (root/'stdout.txt').open('xb') as out,(root/'stderr.txt').open('xb') as err:
            run=subprocess.run(command,cwd=REPO,stdin=subprocess.DEVNULL,stdout=out,stderr=err,timeout=3000,shell=False)
        code=run.returncode
    except Exception as error:
        failure=type(error).__name__+': '+str(error)
    receipt.update(exit_code=code,launch_or_wait_error=failure,finished_at_utc=datetime.now(timezone.utc).isoformat())
    receipt['streams']={n:{'path':str(root/(n+'.txt')),'bytes':(root/(n+'.txt')).stat().st_size,'sha256':sha(root/(n+'.txt'))}for n in ['stdout','stderr']}
    try:
        after=guard();receipt['guard_after']=after
        lines=(root/'stdout.txt').read_text(encoding='utf-8-sig').splitlines()
        rows=[json.loads(line)for line in lines if line.startswith('{')]
        last=next(row for row in reversed(rows)if 'ledger' in row)
        ledger_path=Path(last['ledger']);ledger=load(ledger_path)
        assert ledger['shared_assets']['zip_path']==str(DOWNLOAD/'WinPDFMerger-v1.0.0.zip')
        assert ledger['shared_assets']['checksums_path']==str(DOWNLOAD/'SHA256SUMS.txt')
        assert ledger['result']=='pass' and ledger['source_commit']==R and ledger['harness_commit']==M
        assert [(x['shell'],x['cases'])for x in ledger['candidate_reports']]==[('PS51',14),('PS7',11)]
        receipt['native_ledger']={'path':str(ledger_path),'sha256':sha(ledger_path),'application_cases':25,
                                  'hosts':[{'shell':x['shell'],'cases':x['cases'],'invocations':x['invocations'],'raw_report_sha256':x['sha256']}for x in ledger['candidate_reports']]}
        passed=code==0 and failure is None and all(after.values())
    except Exception as error:
        passed=False;receipt['verification_error']=type(error).__name__+': '+str(error)
    receipt.update(state='complete',result='pass_for_exact_independently_anonymously_downloaded_published_pair_native_capture' if passed else 'fail')
    save()
    print(json.dumps({'receipt':target.relative_to(REPO).as_posix(),'receipt_sha256':sha(target),
                      'result':receipt['result'],'native_ledger':receipt.get('native_ledger')}))
    return 0 if passed else 1
if __name__=='__main__':raise SystemExit(main())
