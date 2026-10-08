"""Audit root-observed visual corpus bindings; no rendering or manual-pass inference."""
import argparse
from collections import Counter
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import PIL
from PIL import Image

parser=argparse.ArgumentParser()
parser.add_argument('--commit',required=True)
parser.add_argument('--review',required=True,type=Path)
parser.add_argument('--output',required=True,type=Path)
args=parser.parse_args()
repo=Path(__file__).resolve().parent
while not (repo/'WinPDFMerge.ps1').is_file():
    repo=repo.parent
work=repo/'tests/.work'
checks,findings,bindings,images=[],[],[],[]
seen=set()


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def label(path):
    return path.resolve().relative_to(repo).as_posix()


def bind(path):
    path=Path(path).resolve()
    assert path.is_relative_to(work) and path.is_file(),'Existing owned evidence file required'
    if path not in seen:
        bindings.append(dict(Path=label(path),SHA256=sha(path),Bytes=path.stat().st_size))
        seen.add(path)
    return path


def load(path):
    return json.loads(bind(path).read_bytes().decode('utf-8-sig'))


def check(condition,message):
    checks.append(dict(Check=message,Passed=bool(condition)))
    if not condition:
        findings.append(message)


def git(*arguments):
    return subprocess.check_output(['git',*arguments],cwd=repo,text=True).strip()


partial=True
started=datetime.now(timezone.utc).isoformat()
review={}
render={}
try:
    check(git('rev-parse','HEAD')==args.commit and not git('status','--porcelain=v1'),'Exact clean C1 before visual-binding audit')
    check(sys.version.split()[0]=='3.12.14' and PIL.__version__=='12.3.0','Actual approved Python/Pillow inspection versions')
    review=load(args.review)
    check(review['CommitUnderTest']==args.commit and review['Phase']=='C1' and review['DirtyWorktree'] is False,'Exact clean root visual-review context')
    check(review['Result']=='pass' and review['ManualVisualInspectionPerformed'] is True,'Root records actual performed visual observation')
    check(review['Observer']=='root Codex actual visual inspection' and review['AcceptanceCases']==['AC041'],'Exact limited declared observer/acceptance class')
    check('No automated fidelity inference' in review['InspectionMethod'],'Root explicitly distinguishes inspection from automated rendering')
    check(any('not owner/Explorer' in value and 'AC058/T26 remains open' in value for value in review['Limitations']),'Owner/Explorer manual-desktop gate remains distinct')
    render_path=bind(review['RendererReceiptPath'])
    check(sha(render_path)==review['RendererReceiptSHA256'],'Exact root renderer receipt binding')
    render=load(render_path)
    render_root=render_path.parent
    source=bind(review['RendererSourcePath'])
    check(sha(source)==review['RendererSourceSHA256']==render['source_sha256'],'Exact renderer source snapshot')
    check(render['phase']=='C1' and render['commit_under_test']==args.commit and render['dirty_worktree'] is False,'Exact clean rendering context')
    check(render['result']=='rendered_pending_visual_review','Render-only receipt never rewritten into automatic manual pass')
    renderer=Path(render['renderer'])
    check(renderer.is_file() and sha(renderer)==render['renderer_sha256'],'Actual retained renderer executable binding')
    check(renderer.name=='pdftoppm.exe','Bundled Poppler executable selected')
    for stream in ('stdout','stderr'):
        version=bind(render_root/('renderer-version.'+stream+'.txt'))
        check(sha(version)==render['renderer_version_'+stream+'_sha256'],'Exact renderer version '+stream+' stream')
    check(b'pdftoppm version' in (render_root/'renderer-version.stderr.txt').read_bytes(),'Actual Poppler version output retained')
    audit=load(work/'T17-C1-native-audit.json')
    check(audit['Result']=='pass' and audit['Partial'] is False and audit['CommitUnderTest']==args.commit,'Exact independent native size/count audit pass')
    fresh={row['Path']:row for row in audit['FreshReads']}
    expected={}
    labels={f'actual-{kind}-{preset}' for kind in ('small-print','scan','mixed') for preset in ('screen','ebook')}
    observation_bindings={}
    for selection in render['observations']:
        shell=selection['shell']
        observation_path=bind(selection['path'])
        check(sha(observation_path)==selection['sha256'],'Exact rendering observation input '+shell)
        observation=load(observation_path)
        check(observation['CommitUnderTest']==args.commit and observation['DirtyWorktree'] is False,'Exact clean native observation input '+shell)
        check(observation['ShellVersion']==selection['shell_version'] and observation['TestSourceSHA256']==selection['observed_test_source_sha256'],
              'Rendering native-source/shell binding '+shell)
        rows=[row for row in observation['Observations'] if row['Label'] in labels]
        check({row['Label'] for row in rows}==labels,'Complete actual preset corpus labels '+shell)
        for row in rows:
            proof=row['Proof']
            expectation=row['FixtureExpectation']
            paths=dict(original=row['OriginalFixturePath'],master=proof['MasterPath'],
                       candidate=proof['EmailJob']['RetainedValidatedCandidate']['Path'])
            for role,pdf in paths.items():
                if role=='original' and row['Label'].endswith('-ebook'):
                    continue
                for page in range(1,expectation['page_count']+1):
                    expected[(shell,row['Label'],role,page)]=dict(PDFPath=pdf,State=proof['ExpectedState'],Expectation=expectation)
        observation_bindings[shell]=dict(Path=label(observation_path),SHA256=sha(observation_path))
    check(set(observation_bindings)=={'ps51','ps7'},'Both actual approved shell observation sets')
    renders=render['renders']
    keys=[(row['shell'],row['label'],row['role'],row['page']) for row in renders]
    check(len(keys)==len(set(keys)) and set(keys)==set(expected),'Complete unique original/master/screen/ebook render coverage')
    check(len(renders)==review['RenderedPageCount']==len(expected),'Actual rendered page count matches declared observation coverage')
    pixels={}
    for row in renders:
        key=(row['shell'],row['label'],row['role'],row['page'])
        context='/'.join(map(str,key))+' '
        target=expected[key]
        pdf=bind(row['source_pdf'])
        png=bind(row['png'])
        check(str(pdf)==target['PDFPath'] and sha(pdf)==row['source_pdf_sha256']
              and pdf.stat().st_size==row['source_pdf_bytes'],context+' exact native original/master/candidate PDF bytes')
        check(label(pdf) in fresh and fresh[label(pdf)]['SHA256']==sha(pdf),context+' PDF independently freshly reread by approved engines')
        check(row['native_publication_state']==target['State'] and row['page_count']==target['Expectation']['page_count'],context+' exact native state/page count')
        check(sha(png)==row['png_sha256'] and png.stat().st_size==row['png_bytes'],context+' actual PNG bytes')
        prefix=png.with_suffix('')
        check(row['command']==[str(renderer),'-png','-r','144','-f',str(row['page']),'-singlefile',str(pdf),str(prefix)],
              context+' exact read-only equal-144DPI render argv')
        check(row['exit_code']==0 and row['started_at_utc']<=row['finished_at_utc'],context+' actual completed successful render receipt')
        for stream in ('stdout','stderr'):
            stream_path=bind(row[stream])
            check(sha(stream_path)==row[stream+'_sha256'],context+' exact retained render '+stream)
        with Image.open(png) as image:
            image.load()
            size=image.size
            rgb=image.convert('RGB')
            pixel_hash=hashlib.sha256(rgb.tobytes()).hexdigest()
        expected_size=tuple(round(value*2) for value in target['Expectation']['page_size_points'])
        check(size==expected_size,context+' exact rendered dimensions at144DPI')
        pixels[sha(png)]=pixel_hash
        images.append(dict(Shell=key[0],Label=key[1],Role=key[2],Page=key[3],Path=label(png),
                           PNGSHA256=sha(png),PixelSHA256=pixel_hash,SizePixels=list(size)))
    unique=review['ReviewedUniqueImages']
    check(len(unique)==review['ReviewedUniqueImageCount']==len({row['png_sha256'] for row in renders}),'Actual ten unique byte classes match declared viewed representatives')
    viewed={}
    for row in unique:
        path=bind(row['Path'])
        check(sha(path)==row['SHA256'] and path.stat().st_size==row['Bytes'],'Viewed representative exact retained bytes '+label(path))
        check('Actual root Codex full-page view_image(detail=original)' in row['Inspection'],'Declared actual full-page root observation '+label(path))
        viewed[str(path)]=row['SHA256']
    coverage=review['Coverage']
    coverage_keys=[(row['Shell'],row['Label'],row['Role'],row['Page']) for row in coverage]
    check(len(coverage_keys)==len(set(coverage_keys)) and set(coverage_keys)==set(expected),'Every bound render has exactly one declared view-coverage row')
    render_by_key={key:row for key,row in zip(keys,renders)}
    for row in coverage:
        key=(row['Shell'],row['Label'],row['Role'],row['Page'])
        actual=render_by_key[key]
        check(row['PDFPath']==actual['source_pdf'] and row['PDFSHA256']==actual['source_pdf_sha256'],'Covered exact source PDF '+str(key))
        check(row['PNGPath']==actual['png'] and row['PNGSHA256']==actual['png_sha256'],'Covered exact render PNG '+str(key))
        check(row['ReviewedViaIdenticalPNG'] in viewed and viewed[row['ReviewedViaIdenticalPNG']]==row['PNGSHA256'],
              'Exact byte-identical viewed representative covers '+str(key))
    by_case={(row['Shell'],row['Label'],row['Role'],row['Page']):row for row in images}
    for shell in ('ps51','ps7'):
        for kind in ('small-print','scan','mixed'):
            page_count=2 if kind=='mixed' else 1
            for page in range(1,page_count+1):
                original=by_case[(shell,'actual-'+kind+'-screen','original',page)]['PixelSHA256']
                for preset in ('screen','ebook'):
                    master=by_case[(shell,'actual-'+kind+'-'+preset,'master',page)]['PixelSHA256']
                    check(original==master,shell+' '+kind+' '+preset+' original/master pixel identity page '+str(page))
    check({row['Document'] for row in review['Observations']}=={'small-print','scan','mixed'},'Actual observation text covers all three required document classes')
    doc=(repo/'docs/EMAIL_PRESETS.md').read_text(encoding='utf-8-sig')
    check('visually compared by Codex' in doc and '144 DPI' in doc,'User documentation labels actual Codex equal-DPI comparison')
    check('not size targets or promises' in doc and 'does not establish fidelity for every PDF' in doc,'User documentation keeps corpus-size/fidelity limits')
    check('3.54 KiB; larger, omitted' in doc and 'fine circle outlines into dots' in doc and 'much clearer' in doc,'User tradeoff description agrees with recorded root observations')
    check(git('rev-parse','HEAD')==args.commit and not git('status','--porcelain=v1'),'Exact clean C1 after visual-binding audit')
    partial=False
except Exception as failure:
    check(False,'Visual-binding audit exception: '+type(failure).__name__+': '+str(failure))

report=dict(SchemaVersion=1,Task='T17',Result='pass' if not findings and not partial else 'fail',Partial=partial,
    CommitUnderTest=args.commit,StartedAtUtc=started,CompletedAtUtc=datetime.now(timezone.utc).isoformat(),
    AuditorSHA256=sha(Path(__file__)),CheckCount=len(checks),BlockingFindings=findings,
    RootVisualReviewPath=label(args.review),RootVisualReviewSHA256=sha(args.review),
    RenderedPageCount=len(images),ReviewedUniqueImageCount=review.get('ReviewedUniqueImageCount'),
    PixelClassCount=len({row['PixelSHA256'] for row in images}),
    ActualManualInspectionPerformedByThisReviewer=False,RenderingOrApplicationRerun=False,
    Checks=checks,RawBindings=bindings,ImageBindings=images,
    Limits=['This certificate verifies source/native/render/declared-view bindings; root Codex owns the actual visual observations and view_image events.',
            'This reviewer did not infer a manual pass from successful rendering or reperform owner/Explorer acceptance.',
            '40 complete retained render pages and10 explicit viewed representatives are verified by exact bytes and decoded pixel hashes; dedup coverage is explicit.',
            'No additional PDF generation, engine conversion, suite/application run or image modification was performed.',
            'No universal fidelity/signature/form/PDF-A/accessibility/security or package/release acceptance claim.'])
assert args.output.resolve().is_relative_to(work) and not args.output.exists()
with args.output.open('x',encoding='utf-8') as stream:
    json.dump(report,stream,indent=2)
    stream.write('\n')
print(json.dumps(dict(Result=report['Result'],Partial=partial,CheckCount=len(checks),RenderedPageCount=len(images),
                      ReviewedUniqueImageCount=report['ReviewedUniqueImageCount'],Report=label(args.output),
                      SHA256=sha(args.output),BlockingFindings=findings)))
sys.exit(0 if report['Result']=='pass' else 1)
