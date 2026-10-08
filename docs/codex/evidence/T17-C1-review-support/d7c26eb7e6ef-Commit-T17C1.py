from pathlib import Path
import hashlib,json,subprocess
repo=Path.cwd();work=repo/'tests/.work';base='27527e839e3b6b37bc554356618bba2ec169a83a'
git=lambda *a:subprocess.check_output(['git',*a]);sha=lambda b:hashlib.sha256(b).hexdigest()
assert git('rev-parse','HEAD').decode().strip()==base
assert git('branch','--show-current').decode().strip()=='codex/v1.0.0-readiness'
assert git('remote','get-url','origin').decode().strip()==git('remote','get-url','--push','origin').decode().strip()=='https://github.com/PikkuJanne/WinPDFMerger.git'
r=json.loads((work/'T17-runtime-review-dirty.json').read_bytes());assert r['Result']=='no_blocking_findings' and not r['Findings']
r=json.loads((work/'T17-precommit-review.json').read_bytes())
for key in ['CodeReview','StaticAnalysisReview']:assert r[key]['Result']=='pass' and not r[key]['BlockingFindings']
for name,expected in [('tests/pdf/SizeReporting.Tests.ps1','de6fcc7a106e90ec461fe06aac8b96e1eb13b7308e66a7cef66b24893bcb3273'),('tests/pdf/SizeReporting.Native.Tests.ps1','07cf1bc40f65de01497af16b94dafbe6b5dea8b34ebba93f9c6d1933710d53d6')]:assert sha((repo/name).read_bytes())==expected
for leaf,count in [('T17-unit-dirty-history.json',32),('T17-native-dirty-history.json',11)]:
    h=json.loads((work/leaf).read_bytes())
    for shell in ['ps51','ps7']:
        if leaf=='T17-unit-dirty-history.json':
            r=[x for x in h['Attempts'] if x['Invocation']['Selection']==shell][-1];assert r['Execution']['ExitCode']==0 and int(r['NUnitRootAttributes']['total'])==count and int(r['NUnitRootAttributes']['failures'])==0
        else:
            r=[x for x in h['Attempts'] if x['Selection']==shell][-1];assert r['Passed']==r['Total']==count and r['Failed']==0 and r['ExitCode']==0
tasks_path=repo/'docs/codex/TASKS.json';tasks=json.loads(tasks_path.read_bytes());task=next(x for x in tasks['tasks'] if x['id']=='T17');assert task['status']=='in_progress'
task['notes']='T17 reporting implementation frozen for clean C1 validation: exact master/email bytes, binary readable sizes and decimal reduction; equal/larger validated candidate labelled not published and existing master-only success0 preserved. Screen default/fixed flags unchanged. Dirty Unit335, SizeReporting32 and SizeReportingNative11 pass in each PS5.1.26100.9444/pinnedPS7.6.6. Initial focused unit26pass6fail each (logger emits earlier truthful lines) and native6pass5fail each (test fixture name collision) corrected in tests only; raw history retained, initial unit fullsource unavailable explicitly disclosed. Dirty20-page/10unique actual Codex visual comparisons show screen scan small-text/outline degradation and clearer ebook; not clean AC041. PSA1.25.0 scoped5PSfiles0errors51warnings14info each reviewed nonblocking. AC040/041 remain not_run pending clean16-tier579each/1158, actual clean visual review, independent review and live synchronization. T18 not started; broader Windows/Explorer/CI/package/security/publication gates remain.'
tasks_path.write_text(json.dumps(tasks,indent=2)+'\n',encoding='utf-8')
with (repo/'docs/codex/evidence/T17-checkpoint.md').open('ab') as f:f.write(b'''\nActual dirty preparation:335 existing unit,32 size-reporting and11 native\nsize-reporting cases pass in each required actual shell, all bad counts zero\nin the final focused runs. Initial focused unit26pass6fail each and native\n6pass5fail each were test assumptions/fixture shadowing, corrected in tests\nonly; retained raw histories distinguish these from the final focused runs.\nInitial unit full test source was not snapshotted, and this absence is\nexplicitly recorded rather than reconstructed.\n\nRoot Codex actually inspected20 equal144DPI page renders via10 unique PNGs.\nVector text/rules/table remained readable; screen scan smallest text degraded\nand thin circles broke into dots, while ebook preserved clearer text/outlines.\nOriginal/master renders matched in this corpus; no clipping/order change seen.\nThis is dirty development review, not clean AC041 or owner/Explorer acceptance.\nUser preset tradeoffs are documented in docs/EMAIL_PRESETS.md.\n\nScoped PSA1.25.0 over5 changed PS files:0errors51warnings14info each, all\nreviewed nonblocking; not fullT22/lint-clean. Two independent runtime/source\nreviews found no blocking finding. Clean16-tier expected579each/1158, actual\nclean visual acceptance, reviews and matching normal push/live sync pending.\n''')
paths=['README.md','WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','tests/pdf/SizeReporting.Tests.ps1','tests/pdf/SizeReporting.Native.Tests.ps1','tests/fixtures/presets/generate_presets.py','tests/fixtures/presets/manifest.json','docs/EMAIL_PRESETS.md','docs/codex/TASKS.json','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md','docs/codex/evidence/T17-checkpoint.md']
changed=set(filter(None,(git('diff','--name-only','-z')+git('ls-files','--others','--exclude-standard','-z')).decode().split('\0')));assert changed==set(paths),(changed-set(paths),set(paths)-changed)
subprocess.run(['git','diff','--check'],check=True);subprocess.run(['git','add','--',*paths],check=True)
assert set(filter(None,git('diff','--cached','--name-only','-z').decode().split('\0')))==set(paths)
subprocess.run(['git','diff','--cached','--check'],check=True);subprocess.run(['git','diff','--cached','--stat'],check=True)
subprocess.run(['git','commit','-m','Report actual PDF sizes and observed preset tradeoffs (T17)'],check=True)
assert not git('status','--porcelain=v1');c1=git('rev-parse','HEAD').decode().strip()
with (work/'T17-C1-commit.txt').open('x',encoding='utf-8') as f:f.write(c1+'\n')
inventory=(work/'Inventory-T16.py').read_bytes().replace(b'T16',b'T17').replace(b"c1='26ac1b73e3733a23099de53d944e00e4ee412982'",b"c1=(work/'T17-C1-commit.txt').read_text(encoding='utf-8').strip()")
with (work/'Inventory-T17.py').open('xb') as f:f.write(inventory)
print(json.dumps({'task':'T17','C1':c1,'clean':True,'acceptance':'not_run; clean579each/1158 validation next'}))
