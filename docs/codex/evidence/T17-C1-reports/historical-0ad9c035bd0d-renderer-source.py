"""Read-only PDF rendering for explicit Codex visual review, never automated manual acceptance."""
from pathlib import Path
import argparse,datetime,hashlib,json,subprocess,uuid
parser=argparse.ArgumentParser();parser.add_argument('--phase',choices=['dirty','C1'],required=True)
parser.add_argument('--observations',nargs='+',required=True);parser.add_argument('--expected-commit')
args=parser.parse_args();repo=Path.cwd().resolve();work=repo/'tests/.work'
sha=lambda b:hashlib.sha256(b).hexdigest();load=lambda p:json.loads(p.read_bytes().decode('utf-8-sig'))
renderer=Path('<USERPROFILE>/.cache/codex-runtimes/codex-primary-runtime/dependencies/native/poppler/Library/bin/pdftoppm.exe')
assert renderer.is_file()
head=subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()
dirty=bool(subprocess.check_output(['git','status','--porcelain=v1']))
if args.phase=='C1':assert args.expected_commit==head and not dirty
root=work/('T17-'+args.phase+'-visual-renders-'+uuid.uuid4().hex);root.mkdir()
(root/'renderer-source.py').write_bytes(Path(__file__).read_bytes())
version=subprocess.run([str(renderer),'-v'],capture_output=True,timeout=30)
assert version.returncode==0
(root/'renderer-version.stdout.txt').write_bytes(version.stdout);(root/'renderer-version.stderr.txt').write_bytes(version.stderr)
labels={f'actual-{kind}-{preset}' for kind in ('small-print','scan','mixed') for preset in ('screen','ebook')}
records=[];seen=set();inputs=[]
for selection in args.observations:
    shell,path=selection.split('=',1);assert shell in ('ps51','ps7')
    observation_path=Path(path).resolve();assert observation_path.is_relative_to(work)
    raw=observation_path.read_bytes();obs=load(observation_path)
    assert obs['CommitUnderTest']==head and obs['DirtyWorktree']==dirty
    if args.phase=='C1':assert obs['DirtyWorktree'] is False
    rows=[r for r in obs['Observations'] if r['Label'] in labels];assert len(rows)==6
    inputs.append({'shell':shell,'path':str(observation_path),'sha256':sha(raw),'observed_test_source_sha256':obs['TestSourceSHA256'],
        'shell_version':obs['ShellVersion'],'pdftk_version':obs['PdfTkVersion'],'ghostscript_version':obs['GhostscriptVersion']})
    for row in rows:
        proof=row['Proof'];pages=row['FixtureExpectation']['page_count'];assert pages in (1,2)
        for role,pdf in [('original',row['OriginalFixturePath']),('master',proof['MasterPath']),('candidate',proof['EmailJob']['RetainedValidatedCandidate']['Path'])]:
            pdf=Path(pdf).resolve();assert pdf.is_relative_to(work) and pdf.is_file()
            if (shell,role,str(pdf)) in seen:continue
            seen.add((shell,role,str(pdf)));original=pdf.read_bytes()
            if role=='candidate':assert sha(original)==proof['EmailJob']['RetainedValidatedCandidate']['SHA256'].lower()
            for page in range(1,pages+1):
                label=shell+'-'+row['Label']+'-'+role+'-'+str(page);prefix=root/label
                command=[str(renderer),'-png','-r','144','-f',str(page),'-singlefile',str(pdf),str(prefix)]
                started=datetime.datetime.now(datetime.timezone.utc).isoformat();run=subprocess.run(command,capture_output=True,timeout=120)
                stdout=root/(label+'.stdout.txt');stderr=root/(label+'.stderr.txt')
                stdout.write_bytes(run.stdout);stderr.write_bytes(run.stderr)
                png=Path(str(prefix)+'.png');assert run.returncode==0 and png.is_file() and png.stat().st_size>0
                assert pdf.read_bytes()==original
                records.append({'shell':shell,'label':row['Label'],'role':role,'page':page,'page_count':pages,
                    'native_publication_state':proof['ExpectedState'],'source_pdf':str(pdf),'source_pdf_sha256':sha(original),'source_pdf_bytes':len(original),
                    'command':command,'started_at_utc':started,'finished_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'exit_code':run.returncode,
                    'stdout':str(stdout),'stdout_sha256':sha(run.stdout),'stderr':str(stderr),'stderr_sha256':sha(run.stderr),
                    'png':str(png),'png_sha256':sha(png.read_bytes()),'png_bytes':png.stat().st_size})
receipt={'task':'T17','phase':args.phase,'commit_under_test':head,'dirty_worktree':dirty,
 'result':'rendered_pending_visual_review','scope':'Actual read-only Poppler renders of original/master/validated candidates at equal144dpi; rendering success is not a manual pass.',
 'renderer':str(renderer),'renderer_sha256':sha(renderer.read_bytes()),'renderer_version_stdout_sha256':sha(version.stdout),
 'renderer_version_stderr_sha256':sha(version.stderr),'source_sha256':sha(Path(__file__).read_bytes()),'observations':inputs,'renders':records,
 'limits':['Retained omitted candidates are development evidence copies, not published email outputs.',
           'Actual stderr retained, including nonfatal font-cache warnings; pixel review remains required.',
           'No modification to PDF/source/final bytes; no installation/runtime dependency, OCR or network introduced.',
           'PDF/A/signatures/forms/accessibility/security and physical Explorer acceptance are not assessed.']}
(root/'receipt.json').write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
unique={}
for row in records:unique.setdefault(row['png_sha256'],row['png'])
print(json.dumps({'result':receipt['result'],'root':str(root),'rendered_pages':len(records),'unique_pngs':len(unique),'inspect':list(unique.values())},indent=2))
