"""Developer/source adapter probes only; no exporter, app or network execution."""
import ast
import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[2]
spec = importlib.util.spec_from_file_location('independent_t34_auditor',ROOT/'public-review.py')
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
checks = []
def record(value,label):
    checks.append({'check':label,'pass':bool(value)})
def refused(call,label):
    try:
        call()
    except (ValueError,KeyError,IndexError):
        record(True,label)
    else:
        record(False,label)
def auditor(config=None):
    value = mod.Auditor.__new__(mod.Auditor)
    value.repo = REPO
    value.work = REPO/'tests/.work'
    value.config = copy.deepcopy(config or {})
    value.checks = 0
    value.issues = []
    value.counts = {}
    value.expected = {}
    value.metadata_receipts = []
    value.registry = {}
    value.private = (Path(os.environ['USERPROFILE']).name,os.environ.get('COMPUTERNAME',''))
    return value

base = ast.parse((REPO/'tests/.work/T34-public-review-preparation/public-review.py').read_text())
current = ast.parse((ROOT/'public-review.py').read_text())
def methods(tree):
    klass = next(node for node in tree.body if isinstance(node,ast.ClassDef) and node.name=='Auditor')
    return {node.name:node for node in klass.body if isinstance(node,ast.FunctionDef)}
for name in ('walk','xml_bytes','xml_facts','privacy','payloads','owned','replace','projected'):
    record(ast.dump(methods(base)[name],include_attributes=False)==ast.dump(methods(current)[name],include_attributes=False),'Preserved independent '+name+' AST')
record(not any(isinstance(node,(ast.Import,ast.ImportFrom)) and ('export' in ast.unparse(node).lower() or 'producer' in ast.unparse(node).lower()) for node in ast.walk(current)),'No producer/exporter imports')
record(mod.PAIR=={'zip':'2b95e90cc3eb3d47b5619710acd1b6cf551769e90ac89813a1dbf0c899c63fc2','checksums':'d39084cb335c56bb99fa51424ec5aed2d95179f3c44974c81a68a8d3bf1e01ca'},'Fixed actual accepted asset pair')

for kind,name in (('actions_run','owner-pr30-CI-run.stdout.txt'),('annotated_tag','actual-annotated-tag.stdout.txt')):
    raw = (REPO/'tests/.work/T34-preflight-v2'/name).read_bytes()
    registry = {'kind':kind,'raw_sha256':mod.sha(raw)}
    if kind=='actions_run':
        registry.update(git_commit_sha=mod.E33,run_id=38035180410,page_index=0,run_index=0)
    else:
        registry.update(tag_object_sha=mod.TAG,target_commit_sha=mod.R,tag='v1.0.0')
    value = auditor()
    value.registry = {'probe.json':registry}
    original = json.loads(raw)
    output = value.metadata(raw,'probe.json')
    baseline = copy.deepcopy(original)
    if kind=='actions_run':
        baseline['head_commit']['author']['email']='<EMAIL>'
        baseline['head_commit']['committer']['email']='<EMAIL>'
        record('workflow_runs' not in output and output['id']==38035180410,'Bare Actions object remains bare, exact run id retained')
    else:
        baseline['tagger']['email']='<EMAIL>'
    record(output==baseline and not value.issues,kind+': only exact declared email fields differ')
    bad = auditor(); bad.registry={'probe.json':{**registry,'raw_sha256':'0'*64}}
    refused(lambda:bad.metadata(raw,'probe.json'),kind+': wrong raw SHA refused')
    bad = auditor(); bad.registry={'probe.json':{**registry,'kind':'actions_runs'}}
    refused(lambda:bad.metadata(raw,'probe.json'),kind+': broader metadata shape refused')
    if kind=='actions_run':
        bad = auditor(); bad.registry={'probe.json':{**registry,'run_id':True}}
        refused(lambda:bad.metadata(raw,'probe.json'),'Boolean run identifier cannot satisfy integer pin')
        bad = auditor(); bad.registry={'probe.json':{**registry,'git_commit_sha':mod.M}}
        refused(lambda:bad.metadata(raw,'probe.json'),'Merged main cannot replace PR30 head')

closure = REPO/'tests/.work/T34-closure-review'
paths = {'source_merge_review':REPO/'tests/.work/T34-preflight-v2/preflight-result.json','readiness_review':closure/'closure-readiness-review.json','published_download':next(closure.glob('actual-*/public-download-review.json'))}
reuse_path = closure/'identical-native-reuse-binding.json'
download = mod.read(paths['published_download'])
source_paths = {arg for command in download['commands'] for arg in command['argv'] if isinstance(arg,str) and 'WinPDFMerger-t32-source-' in arg}
record(len(source_paths)==1,'Actual original commands expose exactly one clean-R checkout root')
config = {'acceptance_gates':[{'role':role,'source':str(path),'raw_sha256':mod.sha(path.read_bytes())} for role,path in paths.items()],'native_reuse_binding':{'source':str(reuse_path),'raw_sha256':mod.sha(reuse_path.read_bytes())},'local_path_aliases':[{'original':download['download_directory'],'alias':'<T34_PUBLIC_DOWNLOAD>'},{'original':next(iter(source_paths)),'alias':'<T34_SOURCE>'}]}
all_paths = set(paths.values())|{reuse_path,REPO/download['package_audit']['path']}
def gate_auditor(candidate):
    value = auditor(candidate)
    value.expected = {str(index):(path.resolve(),'Developer actual-original schema probe') for index,path in enumerate(all_paths)}
    return value
value = gate_auditor(config)
value.gates()
record(not value.issues and value.counts['new_T34_native_cases']==0 and value.counts['pending_closure_ids']==['AC077','AC078'],'Actual originals satisfy scoped gates without native rerun or closure upgrade')
bad = copy.deepcopy(config);bad['acceptance_gates'].pop()
refused(lambda:gate_auditor(bad).gates(),'Missing actual gate role refused')
bad = copy.deepcopy(config);bad['acceptance_gates'][0]['raw_sha256']='0'*64
refused(lambda:gate_auditor(bad).gates(),'Changed original gate hash refused')
bad = copy.deepcopy(config);bad['native_reuse_binding']['raw_sha256']='0'*64
refused(lambda:gate_auditor(bad).gates(),'Changed original reuse hash refused')
bad = copy.deepcopy(config);bad['local_path_aliases'][1]['original']=download['download_directory']
refused(lambda:gate_auditor(bad).gates(),'Missing exact historical clean-R alias refused')
value = gate_auditor(config);value.expected={}
refused(value.gates,'Unselected required original source/child/reuse refused')

issues = [row['check'] for row in checks if not row['pass']]
report = {'task':'T34','result':'pass_for_developer_independent_adapter_probes' if not issues else 'fail','checks':len(checks),'details':checks,'issues':issues,'auditor_sha256':mod.sha((ROOT/'public-review.py').read_bytes()),'scope':'Developer/source/read-only original schema probes, no producer imports/execution, no actual packet acceptance, no new native/app/CI/network/Git/source writes'}
(ROOT/'adapter-probes.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8',newline='\n')
print(json.dumps({key:report[key] for key in ('result','checks','issues','auditor_sha256')}))
raise SystemExit(bool(issues))
