"""Bind root's actual image inspection; this script does not inspect pixels."""
from pathlib import Path
import argparse,datetime,hashlib,json,subprocess
p=argparse.ArgumentParser();p.add_argument('--receipt',required=True);p.add_argument('--phase',choices=['dirty','C1'],required=True);p.add_argument('--viewed',nargs='+',required=True);a=p.parse_args()
repo=Path.cwd().resolve();work=repo/'tests/.work';path=Path(a.receipt).resolve();assert path.is_relative_to(work)
sha=lambda b:hashlib.sha256(b).hexdigest();r=json.loads(path.read_bytes());head=subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip();dirty=bool(subprocess.check_output(['git','status','--porcelain=v1']))
assert r['phase']==a.phase and r['commit_under_test']==head and r['dirty_worktree']==dirty
if a.phase=='C1':assert not dirty and len(r['renders'])==40 and {x['shell'] for x in r['observations']}=={'ps51','ps7'}
reviewed=[];seen=set()
for name in a.viewed:
    rows=[x for x in r['renders'] if Path(x['png']).name==name];assert len(rows)==1,name
    x=rows[0];assert sha(Path(x['png']).read_bytes())==x['png_sha256'];assert x['png_sha256'] not in seen
    seen.add(x['png_sha256']);reviewed.append({'Path':x['png'],'SHA256':x['png_sha256'],'Bytes':x['png_bytes'],'Inspection':'Actual root Codex full-page view_image(detail=original) inspection in tool conversation.'})
assert seen=={x['png_sha256'] for x in r['renders']}
coverage=[]
for x in r['renders']:
    assert sha(Path(x['png']).read_bytes())==x['png_sha256'] and sha(Path(x['source_pdf']).read_bytes())==x['source_pdf_sha256']
    coverage.append({'Shell':x['shell'],'Label':x['label'],'Role':x['role'],'Page':x['page'],'PDFPath':x['source_pdf'],'PDFSHA256':x['source_pdf_sha256'],'PNGPath':x['png'],'PNGSHA256':x['png_sha256'],'ReviewedViaIdenticalPNG':next(v['Path'] for v in reviewed if v['SHA256']==x['png_sha256'])})
observations=[
 {'Document':'small-print','Observation':'Original/master and both candidate vector pages show readable12/10/8/7/6/5pt text,8pt table,0.25-to-1.5pt rules and grayscale/color blocks. Layout/identifier intact; candidates larger and correctly omitted.'},
 {'Document':'scan','Observation':'Original/master raster12/10/8/7/6pt ladder and circles readable/continuous. Screen smallest8/7/6pt lines difficult to read,6pt particularly degraded, thin circles dotted/broken and line-weight detail reduced. Ebook all ladder lines much clearer, continuous thin circles/rule weights visible. Grayscale/layout/identifier retained.'},
 {'Document':'mixed','Observation':'Vectorpage readable12-to-5pt ladder/table/fine rules with both presets; scanpage same screen small-text/circle degradation and clearer ebook detail. ExpectedP03/P04 order, page layout/identifiers unclipped.'}]
out={'Task':'T17','Phase':a.phase,'CommitUnderTest':head,'DirtyWorktree':dirty,'Result':'pass' if a.phase=='C1' else 'development_observations_only','ManualVisualInspectionPerformed':True,'Observer':'root Codex actual visual inspection','InspectionMethod':'All unique full-page Poppler144DPI PNGs opened with view_image detail=original; identical SHA256 pixels share explicit coverage. Comparisons at equal144DPI. No automated fidelity inference.','ReviewedAtUTC':datetime.datetime.now(datetime.timezone.utc).isoformat(),'AcceptanceCases':['AC041'] if a.phase=='C1' else [],'RendererReceiptPath':str(path),'RendererReceiptSHA256':sha(path.read_bytes()),'RendererSourcePath':str(path.parent/'renderer-source.py'),'RendererSourceSHA256':sha((path.parent/'renderer-source.py').read_bytes()),'RenderedPageCount':len(r['renders']),'ReviewedUniqueImageCount':len(reviewed),'ReviewedUniqueImages':reviewed,'Coverage':coverage,'Observations':observations,'Limitations':['Codex visual review, not owner/Explorer/manual-desktop acceptance; AC058/T26 remains open.','Original synthetic corpus only; no private PDFs, no universal fidelity/signature/form/PDF-A/accessibility/security claim.','Omitted validated candidates are evidence copies, never final email outputs.','Dirty preparation does not count as clean acceptance.']}
target=work/f'T17-{a.phase}-visual-review.json';assert not target.exists();target.write_text(json.dumps(out,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'path':str(target),'sha256':sha(target.read_bytes()),'result':out['Result'],'pages':len(coverage),'unique_views':len(reviewed)}))
