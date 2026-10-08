"""Bind already generated original artifacts; no PDF authoring or render rerun."""
import hashlib
import json
from pathlib import Path
from PIL import Image

repo=Path(__file__).resolve().parents[2]
runtime=Path.home()/'.cache/codex-runtimes/codex-primary-runtime'
root=repo/'tests/.work/T17-original-presets-5912939f506e4774bcf1b2b2f1bc1e76'
def bind(path):
    raw=path.read_bytes()
    return {'Path':str(path),'SHA256':hashlib.sha256(raw).hexdigest(),'Bytes':len(raw)}
generation=json.loads((root/'generation.json').read_text(encoding='utf-8-sig'))
renders=[]
for p in sorted((root/'renders').glob('*.png')):
    with Image.open(p) as image:
        renders.append({**bind(p),'PixelDimensions':list(image.size)})
receipt={'SchemaVersion':1,'Task':'T17','EvidenceClass':'original synthetic corpus authoring and preliminary visual QA; not clean C1 preset acceptance',
 'ArtifactMarker':{'Executable':str(runtime/'dependencies/node/bin/node.exe'),
   'Script':str(runtime/'plugins/openai-primary-runtime/plugins/pdf/skills/pdf/container_tools/mark_artifact_operation_started.mjs'),
   'Arguments':['--operation-kind','create','--expected-output-count','3','--output-format','pdf'],
   'ObservedExitCode':0,'SuccessfulInvocationCount':1,
   'Scheduling':'Marker tool command completed successfully immediately before first three-PDF generator authoring command.',
   'RawCaptureLimit':'Actual tool transcript reports exit0 with empty output; no separate local stdout/stderr files were captured for this marker. It was not run again.'},
 'Generator':{'Executable':str(runtime/'dependencies/python/python.exe'),
   'Arguments':['-B','tests/fixtures/presets/generate_presets.py','--output',str(root)],
   'Source':bind(repo/'tests/fixtures/presets/generate_presets.py'),
   'RawStdout':bind(root/'generation.json'),'Result':generation,
   'CaptureLimit':'Initial shell generation JSON redirection retained; separate generator stderr/native exit receipt not captured in this authoring command. Both subsequent actual shell suites reran generator via captured direct child and independently verified exact manifest hashes with native exit0.'},
 'OriginalPdfs':[bind(root/f['file']) for f in generation['fixtures']],
 'Renderer':{'Executable':str(runtime/'dependencies/native/poppler/Library/bin/pdftoppm.exe'),
   'ExecutableSHA256':bind(runtime/'dependencies/native/poppler/Library/bin/pdftoppm.exe')['SHA256'],
   'ArgumentsTemplate':['-r','100','-png','<original-pdf>','<owned-renders-prefix>'],
   'ObservedExitCodes':[0,0,0],'Images':renders,
   'RawCaptureLimit':'Expanded original rendering commands checked each native exit0. Font-cache warnings about unavailable unrelated display fonts appeared in actual tool output; separate raw stderr files were not retained for initial corpus rendering. Rendered fixture glyphs were visually intact.'},
 'ActualCodexOriginalVisualInspection':{'PagesInspected':4,'ImagesInspected':[r['Path'] for r in renders],
   'Observation':'All complete original pages inspected: headers/footer IDs, 5-12pt vector and 6-12pt scanned ladders, fine rules, grayscale/color blocks, circles and table unclipped and readable at the viewed scale.',
   'Limitation':'Original-authoring QA only. No derived-image manual acceptance, owner/Explorer input, printer test or universal small-print/fidelity claim.'}}
target=repo/'tests/.work/T17-original-corpus-preparation.json'
target.write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8')
print(json.dumps(bind(target)))
