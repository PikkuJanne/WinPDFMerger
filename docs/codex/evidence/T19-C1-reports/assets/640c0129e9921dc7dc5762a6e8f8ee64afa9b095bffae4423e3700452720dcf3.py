"""Independent read-only clean C1 source/static/raw-report/feature receipt review."""
import hashlib,json,re,subprocess,uuid,xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime,timezone
from pathlib import Path
repo=Path(__file__).resolve().parents[2];work=repo/'tests/.work'
c1=(work/'T19-C1-commit.txt').read_text(encoding='utf-8-sig').strip();baseline='7e0fe6fde769911f323bd87e4c4f9382d26e331b'
target=work/'T19-C1-runtime-review.json';assert not target.exists()
sha=lambda raw:hashlib.sha256(raw).hexdigest();load=lambda path:json.loads(Path(path).read_text(encoding='utf-8-sig'))
git=lambda *args:subprocess.check_output(['git',*args],cwd=repo)
checks=[];bindings={};binary={}
def check(ok,label):
    if not ok:raise AssertionError(label)
    checks.append(label)
def bind(path,kind='text'):
    path=Path(path).resolve();path.relative_to(work.resolve());raw=path.read_bytes();key=path.relative_to(repo).as_posix()
    record={'Path':key,'SHA256':sha(raw),'Bytes':len(raw)}
    (binary if kind=='binary' else bindings)[key]=record
    return record
check(git('rev-parse','HEAD').decode().strip()==c1 and not git('status','--porcelain=v1').strip(),'Exact clean C1 current source context')
prior=load(work/'T19-runtime-review-dirty.json');bind(work/'T19-runtime-review-dirty.json')
check(prior['Result']=='pass' and prior['CommitUnderTest']==c1 and not prior['CleanAcceptanceClaimed'],'Prior independent source review is current and does not invent early clean acceptance')
for item in prior['Sources']:
    check(sha((repo/item['Path']).read_bytes())==item['SHA256'],'Fresh current source equals earlier reviewed source: '+item['Path'])
    bind(repo/item['RetainedSourcePath'])
for path in ['WinPDFMerge.ps1','WinPDFMerge.bat','src/WinPDFMerge.Helpers.ps1','tests/TestSupport.ps1','docs/codex/PACKAGE_CONTRACT.json','tools/codex/handoff.py']:
    check((repo/path).read_bytes().replace(b'\r\n',b'\n')==git('show',baseline+':'+path).replace(b'\r\n',b'\n'),'Unchanged baseline runtime/support/package contract: '+path)
index_path=work/'T19-C1-drivers.json';index=load(index_path);bind(index_path)
check(index['task']=='T19' and index['commit']==c1 and {x['shell'] for x in index['drivers']}=={'ps51','ps7'},'Actual clean driver index identifies both required shells')
counts={'PreservationDocs':14,'PreservationNative':6,'Unit':335,'Diagnostics':36,'DiagnosticsNative':11,'ToolInvocation':12}
reports=[];shells=[];features=[];raw_oracle_reads=0;rendered_pages=0
manifest=load(repo/'tests/fixtures/features/manifest.json')
for driver in index['drivers']:
    shell=driver['shell'];root=Path(driver['root']);root.resolve().relative_to(work.resolve())
    aggregate=load(root/'aggregate.json');metadata=load(root/'metadata.json');runs=load(root/'runs.json')
    for path in [root/'aggregate.json',root/'metadata.json',root/'runs.json',root/'sources/driver.py']:bind(path)
    check(sha((root/'aggregate.json').read_bytes())==driver['aggregate_sha256'],'Actual bound driver aggregate hash: '+shell)
    check(aggregate['result']=='pass' and aggregate['phase']=='C1' and aggregate['commit_under_test']==c1 and not aggregate['dirty_worktree'] and aggregate['passed']==sum(counts.values()) and aggregate['bad_counts']==0 and aggregate['tiers']==6,'Actual clean aggregate414/6/allbad0: '+shell)
    check(metadata['commit_under_test']==c1 and metadata['phase']=='C1' and not metadata['dirty_worktree'] and metadata['child_modulepath_removed'] and metadata['no_acquisition_or_persistent_policy_changes'],'Actual clean driver pre-run context: '+shell)
    for item in metadata['sources']:
        copy=Path(item['retained_source']);bind(copy)
        check(sha(copy.read_bytes())==item['sha256']==sha((repo/item['path']).read_bytes()),'Exact pre-run source/current bytes: '+shell+'/'+item['path'])
    check(len(runs)==6 and {row['tier'] for row in runs}==set(counts),'Exactly six intended completed tiers: '+shell)
    for row in runs:
        tier=row['tier'];label=shell+'/'+tier;summary=row['summary'];report=Path(row['report'])
        for path in [Path(row['stdout']),Path(row['stderr']),report/'summary.json',report/'results.xml',root/(tier+'.summary.json'),root/(tier+'.results.xml')]:bind(path)
        check(sha(Path(row['stdout']).read_bytes())==row['stdout_sha256'] and sha(Path(row['stderr']).read_bytes())==row['stderr_sha256'],'Actual outer streams match recorded hashes: '+label)
        check(load(report/'summary.json')==summary and (report/'summary.json').read_bytes()==(root/(tier+'.summary.json')).read_bytes() and (report/'results.xml').read_bytes()==(root/(tier+'.results.xml')).read_bytes(),'Copied/original immutable summary and NUnit bytes agree: '+label)
        check(row['exit_code']==0 and summary['passed']==summary['total']==counts[tier] and all(summary[key]==0 for key in ['failed','failed_blocks','failed_containers','skipped','not_run']),'Actual cases and all bad counts: '+label)
        check(summary['commit_under_test']==c1 and not summary['dirty_worktree'] and summary['process_64_bit'] and summary['pester_version']=='6.2.0' and summary['execution_policy']=='RemoteSigned','Actual test context and selected pins/policy: '+label)
        expected_shell=('5.1.26100.9444','Desktop') if shell=='ps51' else ('7.6.6','Core')
        check((summary['shell_version'],summary['shell_edition'])==expected_shell,'Required actual runtime: '+label)
        argv=row['argv'];check(argv[argv.index('-Tier')+1]==tier and argv[argv.index('-ExecutionPolicy')+1]=='RemoteSigned' and '-NoProfile' in argv,'Actual selected command vector: '+label)
        check(Path(argv[0]).name.lower()==('powershell.exe' if shell=='ps51' else 'pwsh.exe'),'Actual executable matches required host: '+label)
        xml=ET.parse(report/'results.xml').getroot();cases=list(xml.iter('test-case'))
        check(int(xml.attrib['total'])==counts[tier] and len(cases)==counts[tier],'Independently parsed NUnit case total: '+label)
        check(all(int(xml.attrib[key])==0 for key in ['errors','failures','not-run','inconclusive','ignored','skipped','invalid']) and all(case.attrib['result']=='Success' and case.attrib['success']=='True' and case.attrib['executed']=='True' for case in cases),'NUnit independently confirms all cases executed successfully: '+label)
        reports.append({'Shell':shell,'Tier':tier,'Passed':counts[tier],'Report':report.relative_to(work).as_posix(),'NUnitSHA256':sha((report/'results.xml').read_bytes()),'SummarySHA256':sha((report/'summary.json').read_bytes()),'EvidenceClass':summary['evidence_class']})
    doc_path=Path(driver['doc_observations']);doc=load(doc_path);bind(doc_path)
    check(sha(doc_path.read_bytes())==driver['doc_observations_sha256'] and doc['Task']=='T19' and doc['ShellVersion']==expected_shell[0] and len(doc['Observations'])==14,'Actual14 documentation observations/full help record: '+shell)
    check(len({x['Label'] for x in doc['Observations']})==14 and doc['HelpBinding']['ApplicationInvoked'] is False,'Documentation observations distinct and application not invoked by help tier: '+shell)
    for item in doc['SourceBindings']:
        copy=Path(item['RetainedSourcePath']);bind(copy)
        check(item['Exists'] and sha(copy.read_bytes())==item['SHA256']==sha((repo/item['Path']).read_bytes()),'Actual doc/help source snapshot matches C1: '+shell+'/'+item['Path'])
    help_path=Path(doc['HelpBinding']['RetainedTextPath']);bind(help_path)
    check(sha(help_path.read_bytes())==doc['HelpBinding']['SHA256'] and doc['HelpBinding']['Command'][0]=='Get-Help','Actual full-help text/command hash bound: '+shell)
    native_path=Path(driver['native_observations']);native=load(native_path);bind(native_path)
    check(sha(native_path.read_bytes())==driver['native_observations_sha256'] and native['Task']=='T19' and native['ShellVersion']==expected_shell[0] and len(native['Observations'])==6,'Actual six native characterization observations: '+shell)
    observations={x['Label']:x for x in native['Observations']};check(set(observations)=={'original-corpus','master-only','screen','ebook','structural-scope','preservation'},'Exact native observation roles distinguish control/reporting from engine runs: '+shell)
    check(observations['preservation']['SourceBefore']==observations['preservation']['SourceAfter'] and observations['preservation']['ParentEnvironmentPreserved'] and observations['preservation']['CopiedApplicationBytesIdentical'],'Actual source snapshots/environment/copied application preserved: '+shell)
    check(observations['structural-scope']['ApplicationFlagsUnchanged'] and not observations['structural-scope']['AutomaticFlattenOrRepairAdded'],'Actual native scope reports unchanged flags and no extra flatten/repair: '+shell)
    for capture in native['Captures']:
        check(capture['ExitCode']==0,'Actual captured generator/entry/oracle exit0: '+shell+'/'+capture['Label'])
        for stream in ['Stdout','Stderr']:
            path=Path(capture[stream]);bind(path);check(sha(path.read_bytes())==capture[stream+'SHA256'],'Actual child stream bytes: '+shell+'/'+capture['Label']+'/'+stream)
        bind(Path(native['Work'])/(capture['Label']+'.execution.json'))
    pdfs=[]
    for original in observations['original-corpus']['Originals']:pdfs.append(('original',original))
    for route in ['master-only','screen','ebook']:
        run=observations[route];log_path=Path(run['Log']);bind(log_path);log=log_path.read_text(encoding='utf-8-sig')
        check(run['ExitCode']==0 and sha(log_path.read_bytes())==run['LogSHA256'],'Actual entry exit/log hash: '+shell+'/'+route)
        check('Expected page total: 4' in log and 'Email result: '+('skipped' if route=='master-only' else 'published') in log,'Actual logged expected pages and explicit email state: '+shell+'/'+route)
        check('Ownership released: True' in log and 'Exit code: 0' in log,'Actual logged successful released native receipts: '+shell+'/'+route)
        if route=='master-only':check(run['Email'] is None and 'Ghostscript version probe' not in log,'Explicit skip omits GS version probe and email output: '+shell)
        else:check('-dPDFSETTINGS=/'+route in log and '-dSAFER' in log and '-dPDFSTOPONERROR' in log and run['Email']['Snapshot']['file']['bytes']<run['Master']['Snapshot']['file']['bytes'],'Actual fixed requested preset/safety/smaller publication: '+shell+'/'+route)
        foreign=Path(run['Output'])/'foreign-existing.txt';bind(foreign);check(sha(foreign.read_bytes())==run['ForeignSHA256'],'Actual foreign file bytes preserved: '+shell+'/'+route)
        pdfs.append((route+'-master',run['Master']))
        if run['Email'] is not None:pdfs.append((route+'-email',run['Email']))
    for role,record in pdfs:
        pdf=Path(record['Pdf']);bind(pdf,'binary');snapshot_path=Path(record['SnapshotPath']);bind(snapshot_path);snapshot=load(snapshot_path)
        check(snapshot==record['Snapshot'] and sha(snapshot_path.read_bytes())==record['SnapshotSHA256'],'Actual full structural/PDFium snapshot equals recorded native observation: '+shell+'/'+role)
        check(sha(pdf.read_bytes())==snapshot['file']['sha256'] and pdf.stat().st_size==snapshot['file']['bytes'],'Current retained actual PDF bytes match snapshot: '+shell+'/'+role)
        check(snapshot['parser']['strict'] is True and snapshot['signature_observations']['validation_performed'] is False and not snapshot['active_content_observations']['uris_followed'] and not snapshot['active_content_observations']['javascript_platform_provided'],'Strict record has no signature/URI/JavaScript certification inference: '+shell+'/'+role)
        check(snapshot['versions']['pypdf']=='6.10.0' and snapshot['versions']['pdfium']=='153.0.7999.0' and snapshot['versions']['pypdfium2']=='5.13.0','Actual recorded independent parser versions: '+shell+'/'+role)
        check(snapshot['page_count']==len(snapshot['page_identifiers'])==len(snapshot['pages']) and len(set(snapshot['page_identifiers']))==snapshot['page_count'],'Actual distinct page identities and structural/native counts: '+shell+'/'+role)
        for render in snapshot['renders']:
            path=Path(render['path']);bind(path,'binary');check(sha(path.read_bytes())==render['sha256'] and path.stat().st_size==render['bytes'] and render['dpi']==144 and render['draw_forms'] and render['draw_annots'],'Actual retained render bytes/dpi/appearance flags: '+shell+'/'+role+'/'+render['identifier']);rendered_pages+=1
        raw_oracle_reads+=1
        if role!='original':
            check(snapshot['page_identifiers']==manifest['expected_merged_page_identifiers'],'Actual merged four-page order: '+shell+'/'+role)
            if role.endswith('-master'):
                check(len(snapshot['forms']['canonical_fields'])==2 and len(snapshot['forms']['widgets'])==4 and all(row['canonical_in_field_tree'] and row['widget_in_canonical_kids'] and row['parent_chain_reaches_canonical'] and row['page_association_matches'] and row['effective_value_matches_canonical'] for row in snapshot['forms']['relations']),'Measured master canonical field/widget relationships retained: '+shell+'/'+role)
                check({row['full_name'] for row in snapshot['forms']['canonical_fields']}=={'shared_text','1.shared_text'} and {row['value'] for row in snapshot['forms']['canonical_fields']}=={'value-A','value-B'},'Measured repeated name rename and canonical values match public limitation table: '+shell+'/'+role)
            else:check(not snapshot['forms']['canonical_fields'] and not snapshot['forms']['widgets'],'Measured rewritten email has no canonical fields/widgets: '+shell+'/'+role)
            check(len(snapshot['bookmarks'])==4 and not snapshot['named_destinations'],'Measured bookmark count and absent named-destination index: '+shell+'/'+role)
            check(not snapshot['embedded_files'] and not snapshot['tagging']['root_present'] and not snapshot['tagging']['parent_tree_entries'],'Measured missing document attachment entries/structure-tree/ParentTree: '+shell+'/'+role)
            check(all(not row['element_in_structure_tree'] and not row['element_content_reference_matches'] for row in snapshot['tagging']['mcid_associations']),'Residual MCIDs do not imply surviving tags/accessibility: '+shell+'/'+role)
            annots=[item for page in snapshot['pages'] for item in page['annotations']];types=Counter(item['subtype'] for item in annots)
            check(types['/Link']==8 and types['/Text']==4 and types['/Highlight']==4 and types['/FileAttachment']==2,'Measured annotation subtype counts: '+shell+'/'+role)
            payloads={item['file_attachment']['streams']['/F']['decoded_sha256'] for item in annots if item['subtype']=='/FileAttachment'}
            check(payloads=={item['attachment']['sha256'] for item in manifest['fixtures']},'Measured page attachment bytes retained separately from absent document entries: '+shell+'/'+role)
        features.append({'Shell':shell,'Role':role,'PDFSHA256':snapshot['file']['sha256'],'Bytes':snapshot['file']['bytes'],'Pages':snapshot['page_count'],'Rotations':[row['structural']['rotation_degrees'] for row in snapshot['pages']],'CanonicalFields':len(snapshot['forms']['canonical_fields']),'Widgets':len(snapshot['forms']['widgets']),'DocumentAttachments':len(snapshot['embedded_files']),'TagStructureRoot':snapshot['tagging']['root_present'],'ObservedReaderWarnings':snapshot['parser']['warnings']})
    shells.append({'Shell':shell,'Version':expected_shell[0],'Passed':aggregate['passed'],'Reports':len(runs),'NativeCases':6,'DocumentationCases':14})
check(len(reports)==12 and sum(row['Passed'] for row in reports)==828 and raw_oracle_reads==14 and rendered_pages==48,'Actual totals12reports/828cases/14recordedPDFreads/48recordedrenders')
static=[]
for name in ['T19-C1-analyzer-ps51-22f115f953604c9486bb6205d5e0fdfb','T19-C1-analyzer-ps7-8e4936ef3b1344e9840bf691278621c4']:
    root=work/name;analysis=load(root/'analysis.json');execution=load(root/'execution.json')
    check(analysis['CommitUnderTest']==c1 and not analysis['DirtyWorktree'] and analysis['Phase']=='C1' and analysis['Errors']==0 and analysis['Warnings']==16 and analysis['Information']==5,'Actual clean scoped analyzer0/16/5: '+name)
    check(analysis['ShellVersion'] in ['5.1.26100.9444','7.6.6'] and analysis['AnalyzerVersion']=='1.25.0' and not analysis['Administrator'] and analysis['Process64Bit'],'Actual approved standard-user analyzer runtime: '+name)
    for leaf in ['analysis.json','execution.json','stdout.txt','stderr.txt','scope.json']:bind(root/leaf)
    for item in execution['SourceBindings']:
        path=Path(item['RetainedSourcePath']);bind(path);check(sha(path.read_bytes())==item['SHA256']==sha((repo/item['Path']).read_bytes()),'Actual clean pre-run/current static source byte binding: '+name+'/'+item['Path'])
    check(execution['ExitCode']==0 and execution['SourcesUnchanged'] and not execution['PersistentChanges'],'Actual clean analyzer successful/no persistent changes: '+name)
    static.append({'Root':root.relative_to(repo).as_posix(),'ShellVersion':analysis['ShellVersion'],'Errors':0,'Warnings':16,'Information':5})
oracle_root=work/'T19-C1-oracle-46240b5511b64ca8aa6ff126ce849adb';oracle=load(oracle_root/'receipt.json')
check(oracle['commit_under_test']==c1 and not oracle['dirty_worktree'] and oracle['exit_code']==0,'Actual clean oracle graph test context/exit0')
for leaf in ['receipt.json','stdout.txt','stderr.txt']:bind(oracle_root/leaf)
check(sha((oracle_root/'stdout.txt').read_bytes())==oracle['stdout_sha256'] and sha((oracle_root/'stderr.txt').read_bytes())==oracle['stderr_sha256'],'Actual clean graph-test streams hash bound')
oracle_stderr=(oracle_root/'stderr.txt').read_text(encoding='utf-8-sig');check(re.search(r'Ran 12 tests',oracle_stderr) and re.search(r'(?m)^OK\s*$',oracle_stderr),'Actual twelve graph regressions pass; separate from828Pester/native counts')
for item in oracle['sources']:
    path=Path(item['retained_source']);bind(path);check(sha(path.read_bytes())==item['sha256']==sha((repo/item['path']).read_bytes()),'Actual clean graph-test full source snapshot: '+item['path'])
proof_path=Path(index['proof']);proof=load(proof_path);bind(proof_path)
check(proof['commit_under_test']==c1 and proof['passed_pester']==828 and proof['reports']==12 and proof['bad_counts']==0 and proof['live_sync']['local_head']==proof['live_sync']['live_remote_head']==c1 and proof['live_sync']['clean'],'Root proof agrees with independently decoded raw counts/current C1 sync')
check(proof['pr']['headRefOid']==c1 and proof['pr']['isDraft'] and proof['pr']['state']=='OPEN' and proof['pr']['number']==19,'Actual draft PR19 proof metadata')
check(git('rev-parse','HEAD').decode().strip()==c1 and not git('status','--porcelain=v1').strip(),'Exact C1 sources remain clean after read-only review')
document={'Task':'T19','CommitUnderTest':c1,'Result':'pass','Findings':[],'ReviewedAtUtc':datetime.now(timezone.utc).isoformat(),'CheckCount':len(checks),'Checks':checks,
 'PassedPester':828,'Reports':12,'SelectedTierCounts':counts,'Shells':shells,'RawReports':reports,'RecordedNativePDFReadsReviewed':raw_oracle_reads,'RecordedRenderedPagesHashChecked':rendered_pages,'FeatureObservations':features,'StaticAnalysis':static,'SeparateOracleGraphTests':12,
 'SourceBindings':prior['Sources'],'SupportBindings':list(bindings.values()),'BinaryOriginalsRetainedIgnored':list(binary.values()),'ProducerSource':{'Path':'tests/.work/Review-T19C1.py','SHA256':sha(Path(__file__).read_bytes())},
 'Limits':['Reviewer authored14 documentation contract cases; own test design is not independently certified. Assertions and previously failed AfterAll receipt generation remain distinct.',
 'This reviewer independently decoded original/copied NUnit, summaries, command/context, pre-run sources and native feature receipts; no application, native PDF reader, generator or manual image view was rerun.',
 'Fourteen existing pypdf/PDFium PDF reads and48recorded render hashes were reviewed. No new native/interactive/signature/XFA/accessibility/PDF-A/malware or broad fidelity verification is claimed.',
 'Actual displayed field values/orientation claims rely on the separately scoped root image observations; structural form/widget/tag/attachment observations were checked independently against exact retained records.',
 'Artificial inert comment size weighting is only publication characterization; not representative compression/quality evidence. Both original and output feature losses are documented.',
 'Only six selected Pester tiers are claimed here. Twelve in-memory oracle graph tests are separately counted. Full T22/Explorer/package/release acceptance remains open.',
 'Two reader-preparation failures, two dirty doc14pass/1container failures and initial/intermediate dirty analyzer receipts remain historical, excluded from clean totals.'],
 'TrackedWrites':False,'ApplicationOrNativeRerun':False,'ManualInspectionPerformed':False}
with target.open('x',encoding='utf-8') as stream:json.dump(document,stream,indent=2);stream.write('\n')
print(json.dumps({'Path':str(target),'SHA256':sha(target.read_bytes()),'Result':'pass','Checks':len(checks),'PassedPester':828,'Reports':12,'SupportFiles':len(bindings),'IgnoredBinaryBindings':len(binary)}))
