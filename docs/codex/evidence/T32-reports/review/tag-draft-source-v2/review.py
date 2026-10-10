"""Read-only source/schema/receipt review of prepared T32 tag/draft helper; never invoke it."""
import ast, datetime, hashlib, json, pathlib, subprocess, sys
ROOT = pathlib.Path(__file__).resolve().parent
REPO = ROOT.parents[3]
PREP = REPO / 'tests/.work/T32-tag-draft-preparation'
R = '95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
sha = lambda data: hashlib.sha256(data).hexdigest()
checks, issues, bindings = [], [], {}
def check(label, good):
    checks.append({'check':label,'pass':bool(good)})
    if not good: issues.append(label)
def bind(path):
    raw=path.read_bytes();bindings[str(path.relative_to(REPO))]={'sha256':sha(raw),'bytes':len(raw)};return raw
source=bind(PREP/'TagDraft-T32.py'); parsed=ast.parse(source.decode('utf-8'))
check('stable reviewed helper source SHA',sha(source)=='d894da59dc0bad2322a8d9271ae381f54d88b3d3d38f4b949ba146a764e5fd1f')
result=json.loads(bind(PREP/'preparation-result.json'))
check('truthful synthetic helper-only receipt',result['result']=='pass_for_helper_developer_checks' and result['exit_code']==0 and result['helper_remote_writes_executed'] is False and result['synthetic_checks_are_application_acceptance'] is False and result['helper_sha256']==sha(source))
test=bind(PREP/'test_helper.py');ast.parse(test.decode('utf-8'))
check('helper tests source bound',sha(test)==result['test_source_sha256'])
stdout,stderr=bind(PREP/'helper-tests.stdout.txt'),bind(PREP/'helper-tests.stderr.txt')
check('original raw helper streams hash binding',sha(stdout)==result['stdout_sha256'] and sha(stderr)==result['stderr_sha256'])
check('actual seventeen synthetic checks complete',b'Ran 17 tests' in stderr and stderr.rstrip().endswith(b'OK') and b'skipped=' not in stderr)
template=json.loads(bind(PREP/'gates.template.json'))
check('deliberately incomplete template cannot authorize action',template['source_commit']==R and len(template['gates'])==3 and all(g['path'] is None and g['sha256'] is None for g in template['gates']) and all(v['path'] is None and v['sha256'] is None for v in template['assets'].values()))
text=source.decode('utf-8')
reviewed=[
 ('default read-only and explicit execute control',"p.add_argument('--execute',action='store_true')" in text and "need(not mutate or self.execute" in text),
 ('complete hash-bound three-role exact R asset gates',"expected = {'native_ledger','asset_review','independent_native_review'}" in text and "pointer(d,g['source_pointer']) == R" in text and "pointer(d,g['zip_pointer']) == zip_hash and pointer(d,g['checksums_pointer']) == sums_hash" in text),
 ('actual complete both-host operation guards',"n['preparation'] is False" in text and "len(n['cases']) == (14 if kind == 'PS51' else 11)" in text and "all(x['package_guard'] is True and x['source_foreign_guard'] is True for x in n['cases'])" in text and "all(x is True for x in n['source_guard'].values())" in text),
 ('native cache/help/manual scope preserved',"d['approved_cache_files_verified'] == 348" in text and "n['public_help']['no_outputs'] is True" in text and "excluded/unperformed; never pass" in text),
 ('independent issue/check gates and safe placeholders',"pointer(d,g['issues_pointer']) == []" in text and "type(checks) is int and checks > 0" in text),
 ('fresh conflict/rules/workflow/origin and source-freeze checks',all(x in text for x in ['fresh-all-tags','fresh-all-releases','fresh-all-rulesets','fresh-all-workflow-metadata','fresh-main-tree','fresh-workflow','fresh-origin-fetch','fresh-origin-push','fresh-only-M6-source-diff','fresh-clean-primary'])),
 ('normal annotated R tag and exact nonforce push/peel',"['git','tag','-a',TAG,R,'-m','WinPDFMerger v1.0.0']" in text and "['git','push','origin','refs/tags/'+TAG+':refs/tags/'+TAG]" in text and "refs/tags/'+TAG+'^{}':R" in text),
 ('only unpublished verify-tag draft action',"['gh','release','create',TAG,'--repo',TARGET,'--verify-tag','--draft'" in text and "d['published_at'] is None" in text and "d['prerelease'] is False" in text),
 ('exact notes/assets and fresh authenticated download',"draft['body'].encode('utf-8')==(c.root/'frozen-R-notes.md').read_bytes()" in text and "download.mkdir(exist_ok=False)" in text and "sha(download/name)==config['assets'][name]['sha256']" in text),
 ('all actual mutations require stable reviewed source/gates/assets',"sha(__file__)==c.ledger['producer_sha256']" in text and "for g in config['gates']: need(sha(g['path'])==g['sha256']" in text and "Assets changed before draft upload" in text),
 ('raw commands saved before action and no unknown-outcome retry',"No command retry/duplicate label" in text and "unknown_mutation_outcome_requires_read_only_inspection" in text and "mutation_error_requires_read_only_inspection" in text and "Stopped transaction requires read-only inspection; no retry" in text),
 ('direct argv no shell/elevation/publish operation',"shell=False" in text and '--clobber' not in text and '--force' not in text and "['gh','release','edit'" not in text and 'Bypass' not in text)
]
for label,good in reviewed:check('reviewed source: '+label,good)
frozen={}
for path,expected in [('docs/RELEASE_NOTES_v1.0.0.md','38866d8ab69626f49a5ed50381f839d21dca9e702dbcc59c4338ee37f8894fbd'),('.github/workflows/windows-tests.yml','910a39e0b827a3698f6db875514fd50c8e5d8766cd38b5f5faa6d3c9a2db74fe')]:
    p=subprocess.run(['git','-C',str(REPO),'cat-file','blob',R+':'+path],capture_output=True)
    for kind,data in [('stdout',p.stdout),('stderr',p.stderr)]:
        name=ROOT/('frozen-'+str(len(frozen))+'.'+kind+'.bin')
        with name.open('xb') as out:out.write(data)
    frozen[path]={'argv':['git','-C',str(REPO),'cat-file','blob',R+':'+path],'exit_code':p.returncode,'stdout_sha256':sha(p.stdout),'stderr_sha256':sha(p.stderr)}
    check('fixed source pin actual R blob '+path,p.returncode==0 and sha(p.stdout)==expected)
report={'schema_version':1,'task':'T32','result':'pass_for_tag_draft_helper_source_preparation' if not issues else 'fail','source_commit':R,'issues':issues,'checks_total':len(checks),'checks':checks,'file_bindings':bindings,'frozen_git_blob_reads':frozen,'reviewer_source_sha256':sha(pathlib.Path(__file__).read_bytes()),'recorded_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'limitations':[
 'Read-only preparation review and original synthetic helper receipt inspection; helper full preflight/transaction was not invoked by reviewer.',
 'Root must bind actual complete native capture and both independent final asset/native audits, then capture fresh read-only preflight and clean/live session facts before explicit execute.',
 'Seventeen synthetic safety checks are developer evidence only; no tag/draft/platform/application/native pass is asserted here.',
 'Authenticated draft-download byte checks are limited to this later unpublished draft step; independent published-download application operation remains required at T33.',
 'AC058 remains excluded and never passed; no elevation or persistent policy/security changes are needed.'
]}
with (ROOT/'tag-draft-source-review.json').open('x',encoding='utf-8') as out:json.dump(report,out,indent=2);out.write('\n')
print(json.dumps({'result':report['result'],'checks_total':len(checks),'issues':issues,'report_sha256':sha((ROOT/'tag-draft-source-review.json').read_bytes())}))
sys.exit(0 if not issues else 1)
