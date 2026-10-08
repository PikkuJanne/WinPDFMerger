"""Read-only T19 source/runtime invariant and scoped static-analysis review."""
import ast, hashlib, json, subprocess, uuid
from datetime import datetime, timezone
from pathlib import Path

repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
baseline='7e0fe6fde769911f323bd87e4c4f9382d26e331b'
review_commit=(work/'T19-C1-commit.txt').read_text(encoding='utf-8-sig').strip()
target=work/'T19-runtime-review-dirty.json'
assert not target.exists(),'Preserve prior review output'
root=work/('T19-dirty-source-review-'+uuid.uuid4().hex);root.mkdir();(root/'sources').mkdir()
sha=lambda raw:hashlib.sha256(raw).hexdigest()
git=lambda *args:subprocess.check_output(['git',*args],cwd=repo)
checks=[];bindings=[]
def check(ok,message):
    if not ok:raise AssertionError(message)
    checks.append(message)
def bind(path):
    raw=path.read_bytes();row={'Path':path.relative_to(repo).as_posix(),'SHA256':sha(raw),'Bytes':len(raw)}
    bindings.append(row);return row
def capture(argv,label):
    result=subprocess.run(argv,cwd=repo,stdin=subprocess.DEVNULL,stdout=subprocess.PIPE,stderr=subprocess.PIPE)
    out=root/(label+'.stdout.txt');err=root/(label+'.stderr.txt');out.write_bytes(result.stdout);err.write_bytes(result.stderr)
    record=root/(label+'.execution.json');record.write_text(json.dumps({'Task':'T19','Command':argv,'ExitCode':result.returncode,'StdoutSHA256':sha(result.stdout),'StderrSHA256':sha(result.stderr)},indent=2)+'\n',encoding='utf-8')
    for path in [out,err,record]:bind(path)
    check(result.returncode==0,'Read-only captured command succeeded: '+label)
    return result.stdout
check(git('rev-parse','HEAD').decode().strip()==review_commit,'Source review binds actual frozen C1 HEAD')
check(git('branch','--show-current').decode().strip()=='codex/v1.0.0-readiness','Expected development branch')
check(not git('status','--porcelain=v1').strip(),'Actual review source context is now clean C1; dirty static receipts remain historical')
paths=['WinPDFMerge.ps1','WinPDFMerge.bat','src/WinPDFMerge.Helpers.ps1','docs/codex/PACKAGE_CONTRACT.json',
 'tools/codex/handoff.py','tests/TestSupport.ps1','tests/fixtures/features/generate_features.py',
 'tests/fixtures/features/manifest.json','tools/test/feature_oracle.py','tools/test/tests/test_feature_oracle.py',
 'tests/pdf/Preservation.Native.Tests.ps1','tests/help/PreservationDocs.Tests.ps1','tools/test/Invoke-Tests.ps1',
 'tools/test/requirements-fixtures.txt','tests/fixtures/README.md','README.md','docs/PDF_LIMITATIONS.md',
 'tests/.work/Run-T19Tests.py']
sources=[]
for index,path in enumerate(paths):
    raw=(repo/path).read_bytes();copy=root/'sources'/(str(index).zfill(2)+'-'+Path(path).name);copy.write_bytes(raw);bind(copy)
    sources.append({'Path':path,'SHA256':sha(raw),'Bytes':len(raw),'RetainedSourcePath':copy.relative_to(repo).as_posix()})
for path in paths[:6]:
    before=git('show',baseline+':'+path);working=(repo/path).read_bytes()
    check(before.replace(b'\r\n',b'\n')==working.replace(b'\r\n',b'\n'),'Baseline source blob unchanged: '+path)
    retained=root/'sources'/('baseline-'+Path(path).name);retained.write_bytes(before);bind(retained)
helper=(repo/'src/WinPDFMerge.Helpers.ps1').read_text(encoding='utf-8-sig')
for needle in ["@('cat', 'output', $stagedOutput, 'compress', 'dont_ask')","'-dSAFER'","'-dPDFSTOPONERROR'",
 "'-sDEVICE=pdfwrite'","'-dCompatibilityLevel=1.6'","'-dDetectDuplicateImages=true'","'-dPDFSETTINGS=/screen'","'-dPDFSETTINGS=/ebook'"]:
    check(needle in helper,'Historical fixed native vector retained: '+needle)
runtime=(repo/'WinPDFMerge.ps1').read_text(encoding='utf-8-sig')+'\n'+helper+'\n'+(repo/'WinPDFMerge.bat').read_text(encoding='utf-8-sig')
check('python' not in runtime.casefold(),'No runtime Python dependency introduced in actual entry/helper/BAT')
package=json.loads((repo/'docs/codex/PACKAGE_CONTRACT.json').read_text(encoding='utf-8-sig'))
check(not any(path.endswith('.py') for path in package['required_files']),'Public package required files introduce no Python')
check('".py"' in (repo/'tools/codex/handoff.py').read_text(encoding='utf-8-sig'),'Existing package verifier still refuses Python payloads')
generator=repo/'tests/fixtures/features/generate_features.py';oracle=repo/'tools/test/feature_oracle.py'
for path in [generator,oracle,repo/'tools/test/tests/test_feature_oracle.py']:
    tree=ast.parse(path.read_text(encoding='utf-8-sig'))
    imports={node.names[0].name.split('.')[0] for node in ast.walk(tree) if isinstance(node,ast.Import)}
    imports|={node.module.split('.')[0] for node in ast.walk(tree) if isinstance(node,ast.ImportFrom) and node.module}
    check(not imports.intersection({'requests','urllib','http','socket','subprocess'}),'Reviewed development source has no network/process launch imports: '+path.name)
manifest=json.loads((repo/'tests/fixtures/features/manifest.json').read_text(encoding='utf-8-sig'))
check(manifest['generator_sha256']==sha(generator.read_bytes()),'Manifest binds exact current generator bytes')
check(manifest['license']=='CC0-1.0' and len(manifest['fixtures'])==2 and manifest['expected_merged_page_count']==4,'Original two-document/four-page fixture provenance')
recipe=generator.read_text(encoding='utf-8-sig');reader=oracle.read_text(encoding='utf-8-sig')
check('output.exists()' in recipe and 'output.is_relative_to(root / "tests" / ".work")' in recipe and '.open("xb")' in recipe,'Generator restricts new exclusive output to owned ignored workspace')
check('strict=True' in reader and 'path.read_bytes() != before' in reader,'Oracle strict parsing and unchanged PDF-byte recheck')
check('xfa_disabled=True' in reader and 'uris_followed": False' in reader and 'javascript_platform_provided": False' in reader,'Oracle disables XFA form rendering and records no URI/JavaScript platform use')
check('validation_performed": False' in reader and 'no PDF/UA' in reader,'No signature or accessibility certification is inferred')
check('canonical_fields' in reader and 'parent_chain_reaches_canonical' in reader and 'effective_value_matches_canonical' in reader and 'text_literals' in reader,'Form tree, widget relationships and normal appearances observed separately')
check('walk_pairs' in reader and 'element_content_reference_matches' in reader,'Raw destination trees and actual tag associations observed')
check('Nonpainting' in recipe and 'not a representative' in recipe and 'artificial_size_weighting' in recipe,'Artificial size weighting explicitly disclosed')
native=(repo/'tests/pdf/Preservation.Native.Tests.ps1').read_text(encoding='utf-8-sig')
check(native.count("    It '")==6,'Six native characterization cases remain separate from doc contract cases')
check("-Options @('-SkipEmail')" in native and "-Options @('-EmailPreset','ebook')" in native and "-Label 'screen' -Options @()" in native,'Actual unchanged application routes include master-only, default screen and explicit ebook')
check('Invoke-TestChildProcess' in native and 'ParentEnvironmentPreserved=$true' in native,'Shared bounded child runner and parent environment preservation assertion present')
check('ApplicationFlagsUnchanged=$true' in native and 'AutomaticFlattenOrRepairAdded=$false' in native,'Observed feature losses do not trigger native flag changes or automatic flatten/repair')
runner=(repo/'tools/test/Invoke-Tests.ps1').read_text(encoding='utf-8-sig')
check("$Tier -eq 'PreservationDocs'" in runner and "$Tier -eq 'PreservationNative'" in runner,'Only explicit new documentation/native tier entry points added')
check('not native PDF preservation or manual acceptance' in runner and 'visual review separate' in runner,'Runner evidence classes distinguish text contract and native observations')
fixture_doc=(repo/'tests/fixtures/README.md').read_text(encoding='utf-8-sig')
check('generate_features.py --output tests/.work/features' in fixture_doc and '--output-dir' not in fixture_doc,'Earlier reviewed fixture-command mismatch corrected to actual argparse option')
capture(['git','diff','--no-ext-diff','--stat',baseline,review_commit],'diff-stat')
capture(['git','diff','--no-ext-diff',baseline,review_commit,'--','WinPDFMerge.ps1','WinPDFMerge.bat','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','README.md','tests/fixtures/README.md','tools/test/requirements-fixtures.txt','.gitattributes'],'source-diff')
analyzer_roots=[
 'T19-dirty-analyzer-ps51-2df0ae5393d444a5ab378e5ae7b20f57','T19-dirty-analyzer-ps7-2e35351bb6be40168cb2c562b0579c5a',
 'T19-dirty-analyzer-ps51-84ba86dbfd8c48cea678f174f7b7b165','T19-dirty-analyzer-ps7-45f8cb076807480d8ab41bf152c6a8f7',
 'T19-dirty-analyzer-ps51-cde0c6079ee2435280912118b7c6dd26','T19-dirty-analyzer-ps7-870ad6b33d544ba7817532c936044a8c']
static=[]
for name in analyzer_roots:
    directory=work/name;analysis=json.loads((directory/'analysis.json').read_text(encoding='utf-8-sig'));execution=json.loads((directory/'execution.json').read_text(encoding='utf-8-sig'))
    check(analysis['Task']=='T19' and analysis['CommitUnderTest']==baseline and analysis['DirtyWorktree'] and analysis['AnalyzerVersion']=='1.25.0','Actual analyzer receipt binds dirty T19 context: '+name)
    check(analysis['Errors']==0 and execution['ExitCode']==0 and execution['SourcesUnchanged'],'Actual analyzer completed without source changes or errors: '+name)
    check(analysis['ShellVersion'] in ['5.1.26100.9444','7.6.6'] and analysis['Process64Bit'] and not analysis['Administrator'],'Required actual standard-user x64 shell: '+name)
    for filename in ['analysis.json','execution.json','stdout.txt','stderr.txt','scope.json']:
        bind(directory/filename)
    for item in execution['SourceBindings']:
        copy=Path(item['RetainedSourcePath']);check(sha(copy.read_bytes())==item['SHA256'],'Pre-run exact analyzer source snapshot: '+name+'/'+item['Path']);bind(copy)
    check(sha((directory/'stdout.txt').read_bytes())==execution['StdoutSHA256'] and sha((directory/'stderr.txt').read_bytes())==execution['StderrSHA256'],'Exact captured static streams: '+name)
    if name in analyzer_roots[-2:]:
        for item in execution['SourceBindings'][:3]:check(sha((repo/item['Path']).read_bytes())==item['SHA256'],'Final historical analyzer source equals current C1 bytes: '+item['Path'])
    static.append({'Root':directory.relative_to(repo).as_posix(),'ShellVersion':analysis['ShellVersion'],'Errors':analysis['Errors'],'Warnings':analysis['Warnings'],'Information':analysis['Information'],'ReceiptSHA256':sha((directory/'analysis.json').read_bytes()),'CurrentFinalSource':name in analyzer_roots[-2:]})
check(all(item['Warnings']==16 and item['Information']==5 for item in static[-2:]),'Final actual scoped counts0errors/16warnings/5info each')
check(all(sha((repo/item['Path']).read_bytes())==item['SHA256'] for item in sources),'Reviewed source bytes remain unchanged during review')
document={'Task':'T19','Phase':'source-review-at-frozen-C1-with-historical-dirty-static','CommitUnderTest':review_commit,'HistoricalCleanStartCommit':baseline,'CleanAcceptanceClaimed':False,'Result':'pass','CheckCount':len(checks),'ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'Checks':checks,
 'Findings':[],'ResolvedSourceIssue':'Fixture README --output-dir was corrected by root to the actual generator --output option; no PDF engine/runtime change.',
 'Sources':sources,'StaticAnalysis':static,'StaticClassification':{
 'Errors':0,'FinalWarningsPerShell':16,'FinalInformationPerShell':5,'Blocking':False,
 'Reason':'Receipt Write-Host and private test-helper naming/positional style are local test conventions; unused-variable/parameter warnings cross Pester BeforeAll/It/AfterAll boundaries. No suppression or lint-clean claim. Initial automatic Matches assignment and one trailing space were corrected only in reviewer-owned test.'},
 'SupportBindings':bindings,'ProducerSource':{'Path':'tests/.work/Review-T19Dirty.py','SHA256':sha(Path(__file__).read_bytes())},
 'Limits':['Reviewer authored14 preservation documentation cases and performed their narrow test-only receipt/style corrections; own test design is not independently certified.',
 'Generator/oracle/native-suite/root runner/runtime/docs source were independently read; no application/native PDF job or PDF authoring was performed by this review.',
 'Clean C1 was frozen while the review was being prepared. The first captured reader guard failed before reading sources because HEAD advanced; that preparation-only failure is retained, with no application/test failure claim. Dirty analyzer results do not count as clean C1 acceptance. Prior17warning6information pair and intermediate16warning5information pair are historical, retained separately.',
 'Root reported first documentation runs14passed plus1AfterAll container failure; these remain failures despite passing assertions. Corrected focused runs and clean C1 raw reports will be independently bound later.',
 'Current preservation docs include measured observations, but this source review alone does not establish their native accuracy, visual fidelity, signatures, XFA, accessibility, PDF/A, malware removal, package distribution, Explorer or release acceptance.'],
 'TrackedWrites':False,'ApplicationRerun':False,'PublicEvidenceWrites':False}
with target.open('x',encoding='utf-8') as stream:json.dump(document,stream,indent=2);stream.write('\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Result':'pass','CheckCount':len(checks),'SupportFiles':len(bindings)}))
