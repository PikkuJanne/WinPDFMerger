"""Read-only semantic review of root-authored prepared T16 C2 records.

Writes only a unique ignored receipt selected by the retaining runner.
No application, test suite, native engine or remote query is executed.
"""
from pathlib import Path
import argparse,datetime,hashlib,json,re,runpy,subprocess,xml.etree.ElementTree as ET

repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work';e=repo/'docs/codex/evidence'
c1='26ac1b73e3733a23099de53d944e00e4ee412982'
sha=lambda raw:hashlib.sha256(raw).hexdigest()
load=lambda path:json.loads(path.read_bytes().decode('utf-8-sig'))
checks=[];findings=[];bindings={}
def check(name,ok,detail=None):
    checks.append({'Check':name,'Pass':bool(ok)})
    if not ok:findings.append({'Check':name,'Detail':detail})
def bind(path):
    path=path.resolve();raw=path.read_bytes();bindings[path.relative_to(repo).as_posix()]={'SHA256':sha(raw),'Bytes':len(raw)};return raw
def git(*args):return subprocess.check_output(['git','-C',str(repo),*args])
def normalized(raw):return raw.replace(b'\r\n',b'\n')

def main(destination):
    initial_status=git('status','--porcelain=v1','--untracked-files=all')
    check('Exact C1 HEAD retained until root closure commit',git('rev-parse','HEAD').decode().strip()==c1)
    changed={line[3:] for line in initial_status.decode().splitlines()}|set(git('diff',c1,'--name-only').decode().splitlines())
    check('All records changes confined to docs/codex and .gitattributes',all(p=='.gitattributes' or p.startswith('docs/codex/') for p in changed),sorted(changed))
    check('No tested source changes from C1 including staged changes',not git('diff',c1,'--','WinPDFMerge.ps1','WinPDFMerge.bat','src','tests','tools','README.md'))
    old_tasks=json.loads(git('show',c1+':docs/codex/TASKS.json'));tasks=load(repo/'docs/codex/TASKS.json')
    old={r['id']:r for r in old_tasks['tasks']};new={r['id']:r for r in tasks['tasks']}
    check('Only T16 task changed',old.keys()==new.keys() and {k for k in old if old[k]!=new[k]}=={'T16'})
    check('T16 done and T17 pending',new['T16']['status']=='done' and new['T17']['status']=='pending')
    old_cases=json.loads(git('show',c1+':docs/codex/ACCEPTANCE_CASES.json'));cases=load(repo/'docs/codex/ACCEPTANCE_CASES.json')
    previous={r['id']:r for r in old_cases['cases']};current={r['id']:r for r in cases['cases']}
    check('Only AC038 and AC039 changed',previous.keys()==current.keys() and {k for k in previous if previous[k]!=current[k]}=={'AC038','AC039'})
    for key,mode in [('AC038','integration'),('AC039','unit')]:
        row=current[key]
        check(key+' required correct-class pass',row['result']=='pass' and row['mode']==mode and row['required'] and row['exclusion_reason'] is None)
        check(key+' contract unchanged',{k:v for k,v in row.items() if k not in ('result','evidence')}=={k:v for k,v in previous[key].items() if k not in ('result','evidence')})
        for ref in row['evidence']:check(key+' referenced evidence '+Path(ref).name,(repo/ref).is_file())
    for ref in new['T16']['evidence']:check('Task evidence exists '+Path(ref).name,(repo/ref).is_file())
    for leaf in ('.gitattributes','docs/codex/COMPATIBILITY_MATRIX.md'):
        check('Historical prefix intact '+leaf,normalized(bind(repo/leaf)).startswith(normalized(git('show',c1+':'+leaf))))
    for leaf in ('docs/codex/TASKS.json','docs/codex/ACCEPTANCE_CASES.json','docs/codex/STATUS.md','docs/codex/NEXT_SESSION.md'):
        bind(repo/leaf)
    proof_root=work/'T16-collector-check-0ae1a1ba50214a7f9afd0371dc9cfbb9';write_root=work/'T16-collector-write'
    proof=load(proof_root/'stdout.txt');written=load(write_root/'stdout.txt')
    for root in (proof_root,write_root):
        execution=load(root/'execution.json');bind(root/'execution.json')
        check('Collector check/write exit0 '+root.name,execution['ExitCode']==0)
        check('Collector check/write exact source '+root.name,execution.get('CollectorSHA256',execution.get('CollectorSourceSHA256'))==sha(bind(work/'Collect-T16Evidence.py')))
        for stream in ('Stdout','Stderr'):check('Collector exact '+stream+' '+root.name,sha(bind(root/(stream.lower()+'.txt')))==execution[stream+'SHA256'])
    for field in ('files','manifest_sha256','results_sha256','literal_whitespace_waiver_suggestions'):
        check('335-plan check/write equivalence '+field,proof[field]==written[field])
    check('Check/write335files28reports1072cases',proof['public_files']==written['public_files']==335 and proof['clean_reports']==written['clean_reports']==28
        and proof['total_passed']==written['total_passed']==1072 and proof['check_only'] is True and written['check_only'] is False)
    planned={r['file']:r['sha256'] for r in proof['files']}
    check('335plan distinct filenames',len(planned)==335)
    for name,digest in planned.items():check('Collector byte plan '+Path(name).name,sha(bind(repo/name))==digest)
    results=load(e/'T16-C1-results.json');manifest=load(e/'T16-C1-reports/manifest.json')
    check('Results exact bytes and AC038/39 clean acceptance',sha(bind(e/'T16-C1-results.json'))==proof['results_sha256'] and
        results['task']=='T16' and results['commit_under_test']==c1 and results['dirty_worktree'] is False and results['ac038']==results['ac039']=='pass'
        and results['total_passed']==1072 and results['clean_reports']==28 and results['all_failures_skips_not_run']==0)
    check('Manifest exact bytes and frozen context',sha(bind(e/'T16-C1-reports/manifest.json'))==proof['manifest_sha256'] and manifest['task']=='T16'
        and manifest['commit_under_test']==c1 and manifest['dirty_worktree'] is False and manifest['clean_reports']==28
        and manifest['total_clean_passed']==1072 and manifest['results_sha256']==proof['results_sha256'])
    clean=[r for r in manifest['records'] if r['classification']=='clean implementation acceptance/regression execution']
    check('Exactly28 clean report records',len(clean)==28)
    expected={'Unit':335,'Parameters':31,'ParametersNative':9,'EmailOutcome':11,'MasterValidation':7,'Staging':9,'InputPreflight':22,
        'Destination':15,'ToolInvocation':12,'GhostscriptPaths':13,'Launcher':24,'LauncherNative':2,'FaultIO':32,'FaultRecovery':14}
    for record in clean:
        root=e/'T16-C1-reports';summary=load(root/record['summary_file']);xml=ET.fromstring(bind(root/record['xml_file']))
        check('Report context/bytes/count '+record['shell']+'/'+record['tier'],sha(bind(root/record['summary_file']))==record['summary_sha256'] and
            sha((root/record['xml_file']).read_bytes())==record['xml_sha256'] and summary['commit_under_test']==c1 and summary['dirty_worktree'] is False
            and summary['passed']==summary['total']==record['counts']['passed']==int(xml.attrib['total'])==expected[record['tier']]
            and all(summary[k]==0 for k in ('failed','failed_blocks','failed_containers','skipped','not_run')))
    for label in ('ps51','ps7'):
        rows=[r for r in clean if r['shell']==label]
        check(label+' exact14tiers536cases',{r['tier']:r['counts']['passed'] for r in rows}==expected and sum(r['counts']['passed'] for r in rows)==536)
    provenance=load(e/'T16-C1-review-provenance.json');bind(e/'T16-C1-review-provenance.json')
    check('Supplemental provenance clean335-plan bindings',provenance['task']=='T16' and provenance['implementation_commit']==c1 and provenance['result']=='pass'
        and provenance['collector_plan_files']==335 and provenance['collector_manifest_sha256']==proof['manifest_sha256'] and provenance['collector_results_sha256']==proof['results_sha256'])
    check('Supplemental source/public filenames unique',len({r['source'] for r in provenance['bindings']})==len(provenance['bindings'])==len({r['file'] for r in provenance['bindings']}))
    helpers=runpy.run_path(str(work/'Collect-T16Evidence.py'),run_name='readonly_closure_sanitizer');sanitizer=helpers['T16Collector'](repo,c1)
    for row in provenance['bindings']:
        source=repo/row['source'];target=repo/row['file'];raw=bind(source);public=bind(target)
        check('Supplemental exact raw/public '+target.name,sha(raw)==row['raw_sha256'] and sha(public)==row['public_sha256'] and len(raw)==row['raw_bytes']
            and len(public)==row['public_bytes'] and (raw!=public)==row['privacy_changed_bytes'])
        expected_public=helpers['json_bytes'](sanitizer.sanitize_value(load(source))) if source.suffix=='.json' else sanitizer.sanitize_string(raw.decode('utf-8-sig')).encode('utf-8')
        check('Supplemental only intended privacy/canonicalization '+target.name,expected_public==public)
    native=load(e/'T16-C1-native-audit.json');archive=load(e/'T16-C1-evidence-review.json');runtime=load(e/'T16-C1-runtime-review.json');sync=load(e/'T16-C1-live-sync.json')
    check('Existing native independent2273checks18cases28freshreads',native['CommitUnderTest']==c1 and native['Result']=='pass' and native['Partial'] is False
        and native['CheckCount']==len(native['Checks'])==2273 and native['CaseCount']==18 and native['FreshFinalReads']==28
        and not native['Findings'] and all(r['passed'] for r in native['Checks']))
    check('Evidence independent archive result coherent',archive['CommitUnderTest']==c1 and archive['Result']=='pass' and not archive['BlockingFindings']
        and archive['CheckCount']==len(archive['Checks']) and all(r.get('Pass',r.get('passed',r.get('pass'))) for r in archive['Checks']))
    check('Runtime independent no-block review coherent',runtime['ImplementationCommit']==c1 and runtime['Result']=='no_blocking_findings' and not runtime['Findings'])
    check('Actual C1 live sync receipt coherent',sync['local_head']==sync['live_remote_head']==c1 and sync['clean'] and sync['synchronized']
        and sync['repository']=='PikkuJanne/WinPDFMerger' and sync['branch']=='codex/v1.0.0-readiness')
    for label in ('ps51','ps7'):
        psa=load(e/'T16-C1-reports'/f'{label}-PSScriptAnalyzer-findings.json')
        check(label+' scoped seven-file static0/71/34',psa['CommitUnderTest']==c1 and psa['DirtyWorktree'] is False and
            (psa['Errors'],psa['Warnings'],psa['Information'])==(0,71,34))
    completion=bind(e/'T16-completion.md').decode('utf-8-sig');status=bind(repo/'docs/codex/STATUS.md').decode('utf-8-sig');next_text=bind(repo/'docs/codex/NEXT_SESSION.md').decode('utf-8-sig')
    for name,text in [('completion',completion),('status',status),('next',next_text)]:
        check(name+' C1/counts/continuation coherent',c1 in text and '536' in text and ('1072' in text or '1,072' in text) and 'T17' in text)
        check(name+' physical/fidelity/release limits explicit','Explorer' in text and 'fidelity' in text and 'release' in text and 'UNC' in text)
    check('Completion28reports/native2273/history/control limits','28 complete reports' in completion and '2,273' in completion
        and 'controlled' in completion.lower() and '1pass/8fail' in completion and '24pass/7fail' in completion and 'sentinel-label' in completion
        and 'Physical Explorer' in completion and 'T17' in completion)
    check('Continuation T17 selected/publication unstarted','Selected task: T17' in next_text and 'Publication: NOT STARTED' in status)
    check('Completion no unresolved template markers','__PR_URL__' not in completion and '__ARCHIVE_CHECKS__' not in completion)
    check('Completion PR16 reference','https://github.com/PikkuJanne/WinPDFMerger/pull/16' in completion)
    flat_completion=' '.join(completion.split())
    check('Completion current C2 commit/live gate still pending',"C2's own SHA/equality is reported in the session after commit" in flat_completion
        and 'final intended-path commit/push/live check' in flat_completion)
    attrs=bind(repo/'.gitattributes').decode('utf-8-sig');old_attrs=git('show',c1+':.gitattributes').decode('utf-8-sig')
    additions=normalized(attrs.encode())[len(normalized(old_attrs.encode())):].decode()
    for line in additions.splitlines():
        if 'whitespace=-' in line:check('Literal whitespace waiver is file-specific '+line.split()[0],not any(c in line.split()[0] for c in '*?['))
    waiver_path=work/'T16-C2-whitespace-waivers.json';waivers=load(waiver_path);bind(waiver_path)
    for row in waivers['waivers']:
        check('Exact waiver payload digest '+Path(row['file']).name,sha(bind(repo/row['file']))==row['sha256'])
        attr=git('check-attr','text','whitespace','--',row['file']).decode()
        check('Exact waiver attribute '+Path(row['file']).name,'text: unset' in attr and
            (not row['trailing'] or '-blank-at-eol' in attr) and (not row['blank_eof'] or '-blank-at-eof' in attr))
    check('Suggested collector waivers covered',{r['file'] for r in proof['literal_whitespace_waiver_suggestions']}<={r['file'] for r in waivers['waivers']})
    public_paths=list(dict.fromkeys([repo/name for name in planned]+[repo/r['file'] for r in provenance['bindings']]+[e/'T16-C1-review-provenance.json',e/'T16-completion.md']))
    identities=set()
    for label in ('ps51','ps7'):
        jobs=load(work/('T16-C1-'+label)/'runs.json');env=ET.fromstring((Path(jobs[0]['report'])/'results.xml').read_bytes()).find('environment')
        identities.update(env.attrib.get(k,'') for k in ('user','user-domain','machine-name'))
    for path in public_paths:
        text=bind(path).decode('utf-8-sig')
        check('Public privacy '+path.name,not re.search(r'(?i)[A-Z]:[\\/]+Users[\\/]+[^<>\\/\s]+',text)
            and not re.search(r'(?i)ghp_[A-Za-z0-9]+|github_pat_[A-Za-z0-9_]+|https://[^/\s]+@',text)
            and not any(len(identity)>=4 and identity!='REDACTED' and re.search(r'(?<![\w])'+re.escape(identity)+r'(?![\w])',text,re.I) for identity in identities))
    diff=subprocess.run(['git','-C',str(repo),'diff',c1,'--check'],capture_output=True)
    check('Tracked C1-to-prepared diff whitespace gate',diff.returncode==0,diff.stdout.decode('utf-8-sig',errors='replace'))
    check('Current origin unchanged',git('remote','get-url','origin').decode().strip()=='https://github.com/PikkuJanne/WinPDFMerger.git')
    check('Read-only review does not mutate tracked/public files',git('status','--porcelain=v1','--untracked-files=all')==initial_status)
    record={'SchemaVersion':1,'Task':'T16','Phase':'prepared records-only C2 semantic review','ImplementationCommit':c1,
        'ObservedAtUtc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'Result':'pass' if not findings else 'fail','Findings':findings,
        'CheckCount':len(checks),'Checks':checks,'SourceBindings':bindings,'ReviewerSourceSHA256':sha(Path(__file__).read_bytes()),
        'ReviewerAuthorship':'Reviewer authored Parameters.Native.Tests.ps1 and T16 collector. This independently reviews root-authored closure/publisher semantics and bytes; it is not independent review of own tests/collector. Safety/source/native and archive reviewers supply separate independence.',
        'Scope':{'CleanReports':28,'CleanPasses':1072,'PerShell':536,'CollectorPlanFiles':335,'SupplementalBindings':len(provenance['bindings']),
            'PublicFilesScanned':len(public_paths),'ExistingNativeAuditChecks':2273,'ExistingIndependentArchiveChecks':archive['CheckCount']},
        'Limitations':['Read-only JSON/XML/hash/source/static-record validation; no app/suite/native PDF engine or remote state query.',
            'Existing native liveness/command/read and C1 live equality events checked via retained receipts; not recreated here.',
            'Root still owns final staged Git-blob verification, C2 commit/push and fresh clean live equality; no future publication certification.',
            'Historical .gitattributes/compatibility prefixes compare only CRLF/LF-normalized text; every evidence hash remains exact-byte based.',
            'No physical Explorer, T17 visual/manual fidelity, broad OS/UNC/CI/security/package/release acceptance added.']}
    require=not destination.exists()
    if not require:raise RuntimeError('Unique ignored receipt destination already exists')
    destination.write_bytes((json.dumps(record,indent=2)+'\n').encode())
    print(json.dumps({'Result':record['Result'],'CheckCount':len(checks),'Findings':findings,'Receipt':destination.relative_to(repo).as_posix(),
        'SHA256':sha(destination.read_bytes()),'PublicFilesScanned':len(public_paths),'SupplementalBindings':len(provenance['bindings'])},indent=2))
    return 0 if not findings else 1

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--receipt',required=True,type=Path);args=parser.parse_args()
    destination=args.receipt.resolve()
    if not destination.is_relative_to(work):raise RuntimeError('Receipt must remain ignored tests/.work')
    raise SystemExit(main(destination))
