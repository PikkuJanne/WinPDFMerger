"""T16 collector: read-only check by default; root alone authorizes --write.

Runs no application, suites or PDF engines. Reuses reviewed legacy Pester,
privacy, build and native schemas. Dirty history never contributes to totals.
"""
from pathlib import Path
import argparse,json,os,re,runpy,subprocess

PRIOR_PATH=Path(__file__).with_name('Collect-T15Evidence.py')
P=runpy.run_path(str(PRIOR_PATH),run_name='t16_reviewed_prior_primitives')
Prior=P['T15Collector'];Base=P['Base'];Legacy=P['LegacyCollector'];LEGACY=P['LEGACY']
require,sha,load_json,json_bytes=(P[n] for n in ('require','sha','load_json','json_bytes'))
COUNTS={'Unit':335,'Parameters':31,'ParametersNative':9,'EmailOutcome':11,'MasterValidation':7,'Staging':9,
        'InputPreflight':22,'Destination':15,'ToolInvocation':12,'GhostscriptPaths':13,'Launcher':24,
        'LauncherNative':2,'FaultIO':32,'FaultRecovery':14}
TIERS=tuple(COUNTS);Legacy.clean_shell.__globals__['TIERS']=TIERS
ORACLE={'python':'3.12.14','pypdfium2':'5.13.0','pdfium':'153.0.7999.0'}
FROZEN_COMMIT='26ac1b73e3733a23099de53d944e00e4ee412982'

def assert_clean(repo,commit):
    head=subprocess.run(['git','-C',str(repo),'rev-parse','HEAD'],capture_output=True,text=True,check=True).stdout.strip()
    status=subprocess.run(['git','-C',str(repo),'status','--porcelain=v1'],capture_output=True,text=True,check=True).stdout
    require(FROZEN_COMMIT is not None and commit==FROZEN_COMMIT and head==commit and not status,'Exact reviewed clean frozen T16 C1 required')

class T16Collector(Prior):
    def __init__(self,repo,commit):
        super().__init__(repo,commit);self.support=[];self.preparation_history=[]

    def support_file(self,path,expected=None):
        path=self.owned_input(path);raw=path.read_bytes();digest=sha(raw)
        if expected is not None:require(digest==expected.lower(),'Retained supporting raw bytes changed: '+path.name)
        row={'source_relative_path':path.relative_to(self.work).as_posix(),'raw_sha256':digest,'bytes':len(raw),
             'publication':'Raw supporting original retained ignored; structured observation/decision values are archived separately.'}
        if row not in self.support:self.support.append(row)
        return row

    def clean_shell(self,label,root):
        aggregate=Legacy.clean_shell(self,label,root)
        require(aggregate['counts_per_tier']==COUNTS and aggregate['passed']==536,'Exact fourteen-tier536 counts required')
        root=self.owned_input(root);metadata_raw,metadata=load_json(root/'collector.json')
        require(metadata['task']=='T16' and metadata['checkpoint']=='C1' and metadata['commit_under_test']==self.commit
                and metadata['dirty_worktree'] is False and metadata['shell']==label and not metadata['acquisition_performed'],
                'Clean collector metadata differs from frozen context')
        require(metadata['expected_counts_sha256']==sha(self.owned_input(metadata['expected_counts_file']).read_bytes())==self.counts_digest,
                'Driver expected count input changed')
        require(metadata['orchestration_script_sha256']==sha((self.work/'Run-T16Checkpoint.ps1').read_bytes()),'Driver source changed')
        require(metadata['child_only_modulepath_removed'] is True and metadata['suite_timeout_ms']==180000
                and metadata['stream_capture_timeout_ms']==1000,'Driver child environment/bounds differ')
        self.selected_pdftk=metadata['pdftk']
        for leaf,raw in [('collector',metadata_raw),('runs',(root/'runs.json').read_bytes()),('aggregate',(root/'aggregate.json').read_bytes())]:
            name=f'{label}-{leaf}.json';self.add_payload(name,raw)
            aggregate.update({leaf+'_file':name,leaf+'_raw_sha256':sha(raw),leaf+'_sha256':sha(self.payloads[name])})
        _,jobs=load_json(root/'runs.json')
        for job in jobs:
            tier=job['tier'];row=self.record_for(label,tier)
            row['summary_raw_sha256']=row['summary_sha256'];row['summary_sha256']=sha(self.payloads[row['summary_file']])
            require(job['native_test_host_started'] and not job['timed_out'] and not job['capture_error'] and not job['termination_error'],
                    'Clean test-host lifecycle did not complete')
            require(job['executable']==metadata['driver'],'Test host differs from selected pinned driver')
            argv=['-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(self.repo/'tools/test/Invoke-Tests.ps1'),
                  '-PesterModulePath',metadata['pester_manifest'],'-Tier',tier]
            if tier in ('ParametersNative','EmailOutcome','MasterValidation','Staging','InputPreflight','Destination','GhostscriptPaths','LauncherNative','FaultRecovery'):argv+=['-PdftkPath',metadata['pdftk']]
            if tier in ('ParametersNative','EmailOutcome','Staging','InputPreflight','Destination','GhostscriptPaths','FaultRecovery'):argv+=['-GhostscriptPath',metadata['ghostscript']]
            if tier in ('ParametersNative','EmailOutcome','MasterValidation','InputPreflight','FaultRecovery'):argv+=['-PythonPath',metadata['python']]
            require(job['arguments']==argv,'Executed argument vector differs from frozen selection')
            row.update(command_executable=self.sanitize_string(job['executable']),command_arguments=self.sanitize_value(argv),
                       started_at_utc=job['started_at_utc'],completed_at_utc=job['completed_at_utc'],elapsed_ms=job['elapsed_ms'])
            for stream,key in [('stdout','log'),('stderr','stderr_log')]:
                raw=self.owned_input(job[key]).read_bytes();clean=self.sanitize_string(raw.decode('utf-8-sig')).encode('utf-8')
                name=f'{label}-{tier}-{stream}.txt';self.add_payload(name,clean)
                row.update({stream+'_file':name,stream+'_raw_sha256':sha(raw),stream+'_sha256':sha(clean)})
            if tier in ('EmailOutcome','MasterValidation','Staging','InputPreflight','Destination'):
                self.native[(label,tier)]=Base.observations(self,label,tier,job,aggregate['shell_version'],metadata)
            elif tier=='GhostscriptPaths':self.path_observations(job,metadata)
            elif tier=='FaultRecovery':
                _,summary=load_json(self.owned_input(job['report'])/'summary.json')
                row.update(self.build_receipt(summary['native_fixture_build_receipt'],summary['native_fixture_build_receipt_sha256'],
                                              label,tier,'FakeNative.exe','tests/native/FakeNative.cs'))
                self.fault_recovery(label,job,metadata,aggregate['shell_version'])
            elif tier=='FaultIO':self.fault_io(label,job)
            elif tier=='ParametersNative':self.parameters_native(label,job,metadata,aggregate['shell_version'])
            elif tier=='Parameters':self.parameters_unit(label,job,metadata)
        require(self.record_for(label,'GhostscriptPaths')['native_version']=='10.08.0','Actual GS version mismatch')
        return aggregate

    def snapshot_rows(self,rows):
        require(rows and len({r['Path'] for r in rows})==len(rows),'Unambiguous synthetic snapshot paths required')
        for row in rows:
            path=self.owned_input(row['Path']);raw=path.read_bytes()
            require(sha(raw)==row['SHA256'].lower() and len(raw)==row['Length'] and row['ModifiedUtcTicks']>0,'Source/final/foreign snapshot bytes changed')
            info=path.stat()
            require(info.st_mtime_ns//100+621355968000000000==row['ModifiedUtcTicks'] and info.st_file_attributes==row['Attributes'],
                    'Retained snapshot modification time/attributes changed')

    def parameters_native(self,label,job,metadata,version):
        stdout=self.owned_input(job['log']).read_bytes().decode('utf-8-sig')
        matches=re.findall(r'(?m)^Parameters observations: (.+?)\r?$',stdout);require(len(matches)==1,'One native parameter receipt required')
        path=self.owned_input(matches[0]);raw,obs=load_json(path)
        require(obs['CommitUnderTest']==self.commit and obs['DirtyWorktree'] is False and obs['ShellVersion']==version
                and obs['StandardUser'] and obs['Process64Bit'] and obs['PdfTkVersion']=='2.02' and obs['GhostscriptVersion']=='10.08.0'
                and obs['OracleVersions']==ORACLE,'Native parameter context/pins differ')
        require(obs['TestSourceSHA256'].lower()==sha((self.repo/'tests/cli/Parameters.Native.Tests.ps1').read_bytes()),'Native parameter test source differs')
        for leaf,field in [('independent-parameter-inspection.py','OracleSHA256'),('original-parameter-raster.py','GeneratorSHA256')]:
            source=self.owned_input(path.parent/leaf).read_bytes();require(sha(source)==obs[field].lower(),'Native parameter development helper changed')
            self.add_payload(f'{label}-ParametersNative-{leaf}',source)
        require(obs['PythonSHA256'].lower()==sha(Path(metadata['python']).read_bytes()),'Native parameter Python bytes changed')
        _,pdf=load_json(self.repo/'docs/codex/evidence/T03-pdftk-acquisition.json');_,gs=load_json(self.repo/'docs/codex/evidence/T09-gs-acquisition.json')
        expected={Path(x['relative_path']).name:x['sha256'] for x in pdf['extracted_files']}
        expected.update({Path(x['relative_path']).name:x['sha256'] for x in gs['ghostscript_extraction']['selected_files']})
        require({x['Name']:x['SHA256'] for x in obs['EngineSHA256']}==expected,'Native parameter executable/DLL pins differ')
        _,fixture_manifest=load_json(self.repo/'tests/fixtures/numbered/manifest.json')
        fixture=next(f for f in fixture_manifest['fixtures'] if f['file']=='1.pdf')
        require(obs['OriginalFixtureSHA256'].lower()==fixture['sha256']==sha((self.repo/'tests/fixtures/numbered/1.pdf').read_bytes()),'Original synthetic fixture pin changed')
        labels={'actual-positional-default-screen','actual-named-default-screen','actual-batch-default-screen',
                'actual-missing-input-usage-no-interactive-prompt','actual-named-output-ebook','actual-case-insensitive-screen',
                'actual-case-insensitive-ebook','actual-SkipEmail-bound-preset-False','actual-SkipEmail-bound-preset-True'}
        rows=obs['Observations'];require(len(rows)==9 and {r['Label'] for r in rows}==labels,'Exact nine native parameter cases required')
        for row in rows:
            name=row['Label'];proof=row['Proof'];require(row['Before']==row['After'],'Source or foreign snapshot changed');self.snapshot_rows(row['After'])
            entry=self.owned_input(Path(row['AppFolder'])/'WinPDFMerge.ps1')
            require(sha(entry.read_bytes())==row['EntrySHA256'].lower()==sha((self.repo/'WinPDFMerge.ps1').read_bytes()),'Copied actual entry differs')
            self.support_file(entry,row['EntrySHA256']);self.support_file(Path(row['AppFolder'])/'src/WinPDFMerge.Helpers.ps1',row['CopiedHelperSHA256'])
            cap=Path(row['AppFolder']).parent/'captured-calls'
            vector=row['RequestedVector'];source=row['SourceFolder'];preset='ebook' if name.endswith('ebook') else 'screen';skip='SkipEmail' in name
            if row['FixtureGeneration']:
                generation=row['FixtureGeneration'];source_row=next(r for r in row['After'] if str(Path(r['Path']).parent)==source)
                require(generation['seed']==160038 and generation['pixel_dimensions']==[1200,800] and generation['page_size_points']==[432,288]
                    and generation['visible_id']=='T03-16-P01' and generation['pages']==1 and generation['bytes']==source_row['Length']
                    and generation['sha256']==source_row['SHA256'].lower(),'Owned original raster provenance/bytes differ')
            expected_vector={'actual-positional-default-screen':[source],'actual-named-default-screen':['-SourceFolder',source],
                'actual-batch-default-screen':[source],'actual-missing-input-usage-no-interactive-prompt':[],
                'actual-named-output-ebook':['-SourceFolder',source,'-OutputFolder',row['NamedOutputFolder'],'-EmailPreset','ebook'],
                'actual-case-insensitive-screen':[source,'-EmailPreset','ScReEn'],'actual-case-insensitive-ebook':[source,'-EmailPreset','eBoOk'],
                'actual-SkipEmail-bound-preset-False':['-SourceFolder',source,'-SkipEmail'],
                'actual-SkipEmail-bound-preset-True':['-SourceFolder',source,'-SkipEmail','-EmailPreset','ebook']}[name]
            require(vector==expected_vector,'Requested native parameter vector differs')
            if 'missing-input' in name:
                require(proof['Result']['ExitCode']==1 and proof['ClosedStdin'] and not proof['HelperImported'] and proof['NoRunOutputs']
                        and 'Usage: WinPDFMerge.ps1 <FolderWithPDFs>' in proof['Result']['Stdout'] and not list(cap.iterdir()),'Missing input prompted/imported/produced outputs')
                continue
            require(proof['Result']['ExitCode']==0 and proof['Child']['ShellVersion']==('5.1.26100.9444' if 'batch' in name else version)
                    and proof['Child']['EntrySHA256'].lower()==row['EntrySHA256'].lower(),'Actual application shell/result differs')
            master=proof['MasterJob'];self.native_result(master['Job'],metadata['pdftk'])
            require('EmailPreset' not in master['BoundParameterKeys'] and master['ExpectedPageCount']==1,'Preset leaked into master job')
            output=row['NamedOutputFolder'] if name=='actual-named-output-ebook' else row['AppFolder']
            require(proof['OutputDirectory']==output and str(Path(proof['MasterPath']).parent)==output,'Default/named destination differs')
            reads=proof['FinalReads'];require(len(reads)==(1 if skip else 2),'Published final list differs')
            for read in reads:
                self.snapshot_rows([read['Snapshot']]);require(read['PdfTkRead']['ExitCode']==read['OracleRead']['ExitCode']==0
                    and len(re.findall(r'(?m)^NumberOfPages:\s*1\s*$',read['PdfTkRead']['Stdout']))==1
                    and read['Oracle']['page_count']==1 and read['Oracle']['pages']==[{'identifier':'T03-01-P01' if skip else 'T03-16-P01','rotation_degrees':0,'size_points':[432,288]}],
                    'Actual independent final count/ID/geometry differs')
            log_path=self.owned_input(proof['LogPath']);log=log_path.read_bytes().decode('utf-8-sig');require(log==proof['Log'],'Final parameter log changed')
            self.support_file(log_path)
            require('Master validation OK: 1 expected pages inspected' in log and not Path(master['StageDirectory']).exists()
                    and not list(Path(output).glob('.WinPDFMerge*')),'Master validation or staging cleanup differs')
            calls=proof['NativeCalls'];[self.complete_native(c['Result']) for c in calls]
            gs_calls=[c for c in calls if c['Executable']==metadata['ghostscript']]
            if skip:
                require(not gs_calls and not proof['EmailJob'] and not proof['EmailPaths'] and not list(cap.glob('unexpected-GS-*'))
                    and 'Email result: skipped' in log,'Skipped GS discovery/probe/launch not bypassed')
                explained="EmailPreset 'ebook' is ignored because -SkipEmail was supplied."
                require((explained in log)==name.endswith('True') and (explained in proof['Result']['Stdout'])==name.endswith('True'),'Ignored-preset explanation differs')
                require('throws/writes only if GS discovery, version probe or native launch is attempted' in row['ControlledHooks']
                        and 'available' in row['ControlledHooks'],'Skip controls not disclosed')
            else:
                email=proof['EmailJob'];self.native_result(email['Job'],metadata['ghostscript'])
                require(email['RequestedEmailPreset'].lower()==preset and 'EmailPreset' in email['BoundParameterKeys']
                    and email['InspectionExecutable']==metadata['pdftk'] and email['MasterBefore']==email['MasterAfter'],'Preset/inspector/master preservation differs')
                self.snapshot_rows([email['MasterAfter']])
                conversions=[c for c in gs_calls if '-sDEVICE=pdfwrite' in c['Arguments']];require(len(conversions)==1,'One real GS conversion required')
                flags=['-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dCompatibilityLevel=1.6',
                       '-dPDFSETTINGS=/'+preset,'-dDetectDuplicateImages=true','-o',str(Path(email['StageDirectory'])/'email.pdf'),'-f',proof['MasterPath']]
                require(conversions[0]['Arguments']==flags and conversions[0]['RemoveEnvironmentVariables']==['GS_OPTIONS'],'Actual fixed GS vector differs')
                require(Path(proof['EmailPaths'][0]).stat().st_size<Path(proof['MasterPath']).stat().st_size,'Published derivative not strictly smaller')
            for item in cap.glob('*.json'):self.support_file(item)
        name=f'{label}-ParametersNative-observations.json';self.add_payload(name,raw)
        self.record_for(label,'ParametersNative').update(observations_file=name,raw_observations_sha256=sha(raw),observations_sha256=sha(self.payloads[name]),
              observation_count=9,scope='Actual real engines/cmd BAT; copied recording/Skip sentinels disclosed; no physical Explorer or T17 quality pass')

    def parameters_unit(self,label,job,metadata):
        stdout=self.owned_input(job['log']).read_bytes().decode('utf-8-sig')
        matches=re.findall(r'(?m)^Parameters receipts: (.+?)\r?$',stdout);require(len(matches)==1,'One retained unit parameter receipt required')
        raw,rows=load_json(self.owned_input(matches[0]));require(len(rows)==31 and len({r['Label'] for r in rows})==31,'Exact31distinct parameter unit observations required')
        copied=[r for r in rows if 'ReceiptPath' in r];helpers=[r for r in rows if 'ReceiptPath' not in r]
        require(len(copied)==26 and len(helpers)==5,'Controlled entry/helper split differs')
        for index,row in enumerate(copied,1):
            require(row['Before']==row['After'] and row['PersistentEnvironmentChanges'] is False,'Unit preserved bytes/environment changed')
            result=row['Result'];require(result['Started'] and result['OwnershipReleased'] and not result['TimedOut'] and not result['Cancelled']
                and not result['LaunchError'] and not result['CaptureError'] and not result['TerminationError'],'Unit actual shell lifecycle incomplete')
            receipt_path=self.owned_input(row['ReceiptPath']);root=receipt_path.parent;receipt_raw,receipt=load_json(receipt_path)
            require(sha(receipt_raw)==row['ReceiptSHA256'],'Unit copied decision receipt changed')
            for rel,field in [('app/WinPDFMerge.ps1','EntrySHA256'),('app/src/WinPDFMerge.Helpers.ps1','HelperWithHooksSHA256'),
                              ('Invoke-Entry.ps1','WrapperSHA256'),('app/parameter-config.json','ConfigurationSHA256'),('stdout.txt','StdoutSHA256'),('stderr.txt','StderrSHA256')]:
                self.support_file(root/rel,row[field])
            require(row['Command'][0]==metadata['driver'],'Unit copied entry host differs')
            _,configuration=load_json(root/'app/parameter-config.json')
            paths=[Path(configuration['Receipt']).parent/'source [x] ! &/1.pdf',Path(configuration['Receipt']).parent/'output [x] ! &/foreign-existing.pdf']
            require(':'.join(sha(self.owned_input(p).read_bytes()) for p in paths)==row['After'],'Unit source/foreign hashes changed')
            if result['ExitCode']==1:
                require(receipt['NativeCalls']==receipt['PdftkDiscovery']==receipt['GSDiscovery']==receipt['GSVersion']==0 and not receipt['Jobs'],
                        'Rejected unit input reached native work')
            else:
                require(result['ExitCode']==0 and receipt['Outcome']['ExitCode']==0 and receipt['Jobs'][0]['Tool']=='Pdftk'
                    and not receipt['Jobs'][0]['PresetBound'],'Accepted unit master/result differs')
                skip=receipt['Outcome']['EmailState']=='skipped'
                require(receipt['GSDiscovery']==receipt['GSVersion']==(0 if skip else 1) and len(receipt['Jobs'])==(1 if skip else 2),'Unit optional engine decision differs')
            name=f'{label}-Parameters-entry-{index:02d}-receipt.json';self.add_payload(name,receipt_raw)
            self.support_file(receipt_path,row['ReceiptSHA256'])
        for row in helpers:
            require(row['Before']==row['After'] and row['Removals']==['GS_OPTIONS'] and row['Job']['OutputValidated']
                    and row['Job']['OutputPublished'] and row['Job']['Succeeded'],'Controlled helper preset result differs')
            require(row['Arguments'][0:8]==['-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dCompatibilityLevel=1.6',
                    '-dPDFSETTINGS=/'+row['Expected'],'-dDetectDuplicateImages=true'],'Controlled helper fixed flags differ')
        name=f'{label}-Parameters-observations.json';self.add_payload(name,raw)
        self.record_for(label,'Parameters').update(observations_file=name,raw_observations_sha256=sha(raw),observations_sha256=sha(self.payloads[name]),
                  observation_count=31,scope='26actual copied shell-entry controls plus5mocked helper vectors; no PDF-engine support claim')

    def historical(self):
        for index_name in ('T16-unit-dirty-history.json','T16-native-dirty-history.json'):
            path=self.work/index_name;_,index=load_json(path);require(index['Task']=='T16','History ownership differs');self.archive(path)
            for attempt in index['Attempts']:
                for item in attempt['Files']:
                    original=self.repo/item['Path'];require(sha(original.read_bytes())==item['SHA256'],'Indexed historical bytes changed');self.archive(original)
                copies=attempt.get('RetainedCopiedEntryCases',attempt.get('RetainedCopiedApplicationCases',[]))
                for case in copies:
                    for item in case['Files']:self.support_file(self.repo/item['Path'],item['SHA256'])
        for path in sorted(self.work.glob('T16-dirty-*.execution.json')):
            _,execution=load_json(path);self.archive(path)
            for key,digest in [('stdout','stdout_sha256'),('stderr','stderr_sha256')]:
                original=self.owned_input(execution[key]);require(sha(original.read_bytes())==execution[digest],'Dirty root capture changed');self.archive(original)
            text=Path(execution['stdout']).read_bytes().decode('utf-8-sig')
            matches=re.findall(r'(?m)^Reports: (.+?)\r?$',text)
            if len(matches)==1:
                report=self.owned_input(matches[0]);self.archive(report/'summary.json');self.archive(report/'results.xml')
            else:self.records.append({'classification':'historical missing report explicitly absent; never fabricated','source_relative_path':path.relative_to(self.work).as_posix()})
        for path in sorted(self.work.glob('T16-precommit-analyzer-*')):
            if path.is_file():self.archive(path,'historical dirty static findings; not acceptance passes')
            else:
                for child in sorted(path.iterdir()):
                    if child.is_file():self.archive(child,'historical dirty analyzer execution/capture; not acceptance passes')
        for leaf in ('Invoke-T16NativeSmoke.ps1','Build-T16NativeHistory.py'):
            self.archive(self.work/leaf,'historical dirty focused authoring source; not acceptance counts')
        for path in sorted(self.work.glob('T16*review*dirty*.json')):
            _,review=load_json(path);self.archive(path,'historical dirty read-only source review; separate from clean acceptance counts')
            support=review.get('ReviewSupport',{})
            items=([support['Producer']] if 'Producer' in support else [])+support.get('ReadOnlyDiffExecution',[])
            for item in items:
                original=self.repo/item['Path'];require(sha(original.read_bytes())==item['SHA256'],'Dirty review support changed')
                self.archive(original,'historical dirty read-only review support; not acceptance counts')
        analyzer_source=self.work/'Run-T16Analyzer.py'
        if analyzer_source.exists():self.archive(analyzer_source,'historical dirty scoped analyzer source; not acceptance counts')
        for root in sorted(self.work.glob('T16-collector-check-*')):
            if not (root/'execution.json').exists():continue
            _,attempt=load_json(root/'execution.json')
            if attempt['ExitCode']==0:continue
            require(attempt['ApplicationOrNativeTestsExecuted'] is False and not attempt['PublicWriteRequested']
                    and attempt['CleanReports'] is None and attempt['TotalPassed'] is None,'Failed collector preparation misclassified')
            require(sha((root/'collector-source.py').read_bytes())==attempt['CollectorSourceSHA256']
                    and sha((root/'stdout.txt').read_bytes())==attempt['StdoutSHA256']
                    and sha((root/'stderr.txt').read_bytes())==attempt['StderrSHA256'],'Collector failed preparation binding changed')
            for leaf in ('collector-source.py','stdout.txt','stderr.txt','execution.json'):
                self.archive(root/leaf,'historical collector preparation only; no application failure or acceptance counts')
            self.preparation_history.append({'source_relative_root':root.relative_to(self.work).as_posix(),'exit_code':attempt['ExitCode'],
                'execution_receipt_raw_sha256':sha((root/'execution.json').read_bytes()),'collector_source_sha256':attempt['CollectorSourceSHA256'],
                'stdout_raw_sha256':attempt['StdoutSHA256'],'stderr_raw_sha256':attempt['StderrSHA256'],'acceptance_counts_available':False,
                'diagnosis':'Collector demanded a literal sentinel label; actual copied-hook receipt accurately disclosed throws/writes on GS discovery/version/native attempts. Corrected semantic disclosure check; no application/source/test failure.'})

    def environment_and_pins(self,path):
        Legacy.environment_and_pins(self,path);raw,receipt=load_json(self.owned_input(path));self.standalone['T16-environment.json']=self.payloads['ordinary-ps51-environment.json']
        require(receipt['commit_under_inventory']==self.commit and receipt['dirty_worktree'] is False and receipt['task']=='T16'
            and receipt['application_or_native_test_claim'] is False,'Fresh ordinary inventory context differs')
        for leaf,digest in [('T16-InventoryCommand.ps1','inventory_source_sha256'),('T16-environment.stdout.txt','raw_stdout_sha256'),('T16-environment.stderr.txt','raw_stderr_sha256')]:
            source=self.work/leaf;require(sha(source.read_bytes())==receipt[digest],'Fresh inventory source/capture changed')
            self.archive(source,'separate read-only ordinary inventory; no application acceptance pass')

    def verified_cache_receipt(self,path):
        raw,receipt=load_json(self.owned_input(path));require(receipt['task']=='T16' and not receipt['acquisition_performed'],'Cache scope changed')
        require({r['dependency']:r['version'] for r in receipt['dependencies']}=={'Pester':'6.2.0','PDFtk':'pdftk 2.02','Ghostscript':'10.08.0',
                'PowerShell7':'7.6.6','PSScriptAnalyzer':'1.25.0'},'Cache version pins changed')
        for row in receipt['dependencies']:
            require(row['all_selected_file_hashes_match'] and row['selected_files_verified']==len(row['selected_files'])
                and sha((self.repo/row['source_receipt']).read_bytes())==row['source_receipt_sha256'],'Cache source receipt changed')
            root=Path(os.path.expandvars(row['cache_root'].replace('<USERPROFILE>',os.environ['USERPROFILE']))).resolve()
            for selected in row['selected_files']:
                candidate=(root/selected['relative_path']).resolve();require(candidate.is_relative_to(root) and sha(candidate.read_bytes())==selected['sha256'],'Approved selected cache bytes changed')
        require(sum(r['selected_files_verified'] for r in receipt['dependencies'])==14,'Exact14cache inputs required')
        require(sha((self.repo/receipt['previous_full_audit']).read_bytes())==receipt['previous_full_audit_sha256'],'Retained full cache audit changed')
        runtime=receipt['development_oracle_runtime'];require({k:runtime[k] for k in ORACLE}==ORACLE,'Development oracle pins changed')
        for key in ('python','pdfium_dll'):
            path=Path(os.path.expandvars(runtime[key+'_path'].replace('<USERPROFILE>',os.environ['USERPROFILE'])))
            require(sha(path.read_bytes())==runtime[key+'_sha256'],'Development oracle runtime bytes changed')
        self.add_payload('approved-cache-verification.json',raw);self.standalone['T16-cache-verification.json']=self.payloads['approved-cache-verification.json']

    def analyzer(self,label,path,version):
        raw,report=load_json(self.owned_input(path))
        require(report['Task']=='T16' and report['Phase']=='C1' and report['CommitUnderTest']==self.commit and report['DirtyWorktree'] is False
                and report['ShellVersion']==version and report['AnalyzerVersion']=='1.25.0','Clean scoped analyzer context differs')
        for severity,key in [(2,'Errors'),(1,'Warnings'),(0,'Information')]:
            require(report[key]==sum(f['Severity']==severity for f in report['Findings']),'Scoped analyzer counts differ from findings')
        require((report['Errors'],report['Warnings'],report['Information'])==(0,71,34),'Frozen scoped analyzer findings differ')
        name=f'{label}-PSScriptAnalyzer-findings.json';self.add_payload(name,raw)
        self.records.append({'classification':'clean C1 scoped seven-file static findings; warnings/info retained, not test passes',
            'shell':label,'file':name,'raw_sha256':sha(raw),'sha256':sha(self.payloads[name]),'analyzer_version':'1.25.0',
            'error_count':0,'warning_count':71,'information_count':34})
        executions=list(self.work.glob('T16-C1-analyzer-execution-'+label+'-*'));require(len(executions)==1,'One clean scoped analyzer execution required')
        _,execution=load_json(executions[0]/'execution.json')
        require(execution['started'] and execution['exit_code']==0 and not execution['timed_out'] and not execution['error']
            and execution['analyzer_report_sha256']==sha(raw) and execution['commit_after']==self.commit and not execution['git_status_after']
            and execution['source_bytes_unchanged'],'Scoped analyzer lifecycle/source binding differs')
        require(len(execution['source_bindings_after'])==7,'Exactly seven changed PowerShell source bindings required')
        for source,digest in execution['source_bindings_after'].items():require(sha((self.repo/source).read_bytes())==digest,'Scoped analyzer source bytes changed')
        for item in executions[0].iterdir():
            if item.is_file():self.archive(item,'clean C1 scoped analyzer execution/capture; separate from application case counts')

    def results_document(self,shells,destination):
        return {'schema_version':1,'task':'T16','checkpoint':'C1','commit_under_test':self.commit,'dirty_worktree':False,
                'implementation_acceptance':'pass','ac038':'pass','ac039':'pass','cases_per_shell':COUNTS,'selected_tiers_in_execution_order':list(TIERS),
                'passed_per_shell':{s['shell']:s['passed'] for s in shells},'total_passed':1072,'clean_reports':28,'all_failures_skips_not_run':0,
                'scope':'AC038 actual Windows PS5.1/PS7.6.6 legacy/named/default interface, real approved PDFtk/GS fixed screen and ebook selection, closed-stdin missing-input usage, named output and actual cmd/BAT one-folder delivery. AC039 actual advanced binding and copied-entry/mock controls validate allowlists, preflight, unsupported arguments and SkipEmail; controls do not certify native engines.',
                'reports_manifest':'docs/codex/evidence/T16-C1-reports/manifest.json','selected_cache_verification':'docs/codex/evidence/T16-cache-verification.json',
                'ordinary_environment_receipt':'docs/codex/evidence/T16-environment.json','pester_version':'6.2.0','pdftk_version':'2.02','ghostscript_version':'10.08.0','development_oracle':ORACLE,
                'collector_command_template':'<selected shell> -NoProfile -ExecutionPolicy RemoteSigned -File tests/.work/Run-T16Checkpoint.ps1 -ShellLabel ps51|ps7 -ExpectedCommit '+self.commit+' -ExpectedCountsPath tests/.work/T16-expected-counts.json -Checkpoint C1',
                'evidence_collector_command':'<approved Python> -B tests/.work/Collect-T16Evidence.py --repo . --commit '+self.commit+' [--write]',
                'historical_rule':'All failed/corrected dirty focused attempts and static observations remain separate, never added to clean1072. Original records and copied support hash inventories retained; absent outputs never fabricated.',
                'historical_findings':{'Parameters':'Initial31-case unit runs24passed/7failed each due to test wrapper not propagating entry LASTEXITCODE; actual refusals correct. Corrected31passes each.',
                    'ParametersNative':'Initial PS5.1 native run1passed/8failed due to test JSON array pipeline parsing; real entry/native jobs succeeded. Corrected reader9passes both; initialPS7already9passes.'},
                'collector_preparation_history':self.preparation_history,
                'environment':{'inventory':self.inventory,'standard_user_non_elevated':True,'test_process_policy':'Previously authorized process-only RemoteSigned',
                    'persistent_policy_security_parent_environment_changed':False,'acquisition_performed':False},
                'static_analysis':{'scope':'Seven changed PowerShell files; separate from application cases and full T22 lint gate',
                    'analyzer_version':'1.25.0','each_shell':{'errors':0,'warnings':71,'information':34}},
                'limitations':['Actual cmd/BAT selects Windows PowerShell5.1 and is distinct from the open physical Explorer/manual gate.',
                    'Deterministic synthetic raster page count/ID/geometry and strict byte reduction do not certify T17 preset quality, full fidelity, signatures, PDF/A or security.',
                    'Copied helper recording and SkipEmail throwing sentinels are controlled and disclosed. Earlier affected fault cases retain controlled corrupt-input, owned padding, logger/token/native-tree seams.',
                    'Only14affected tiers selected; unchanged NativeRunner, DependencyEntry, SourceDiscovery, PdftkPaths and NativeFixture tiers retain prior accepted evidence rather than a new run claim.',
                    'No Windows support-channel, UNC/Explorer/CI/package or published-release acceptance follows; downstream gates remain open.'],
                'checkpoint_note':'This certifies clean implementation gates; root separately maintains task/status and final records synchronization.'}

    def finish(self,shells,destination,write):
        require(sum(s['passed'] for s in shells)==1072,'Exact1072clean total required')
        evidence=self.repo/'docs/codex/evidence';destination=destination.resolve();require(destination==evidence/'T16-C1-reports','Exact root-authorized destination required')
        self.add_payload('retained-support-bindings.json',json_bytes(self.support))
        result_path=evidence/'T16-C1-results.json';result_payload=json_bytes(self.sanitize_value(self.results_document(shells,destination)))
        manifest={'schema_version':1,'task':'T16','checkpoint':'C1','commit_under_test':self.commit,'dirty_worktree':False,'clean_reports':28,
                  'total_clean_passed':1072,'shells':shells,'records':self.records,'payload_bindings':self.payload_bindings.copy(),
                  'evidence_collector_sha256':sha(Path(__file__).read_bytes()),'prior_collector_primitives_sha256':sha(PRIOR_PATH.read_bytes()),
                  'native_schema_primitives_sha256':sha(P['BASE_PATH'].read_bytes()),'legacy_primitives_sha256':sha(P['B']['LEGACY_PATH'].read_bytes()),
                  'xml_redactions':['environment.'+k for k in LEGACY['XML_IDENTITY']],
                  'privacy_scope':'XML machine/user/domain/cwd and repository/user/cache/temp path prefixes sanitized; synthetic IDs, whitespace, vectors, diagnostics and native results retained.',
                  'exact_byte_policy':'Raw and public sanitized/canonical SHA bindings are distinct. Supporting copied raw files retained ignored and verified in support inventory; only original synthetic document paths permitted.',
                  'results_file':'docs/codex/evidence/T16-C1-results.json','results_sha256':sha(result_payload),
                  'standalone_receipts':[{'file':'docs/codex/evidence/'+n,'sha256':sha(p)} for n,p in self.standalone.items()],
                  'frozen_implementation_source_bytes':[{'path':p,'sha256':sha((self.repo/p).read_bytes())} for p in
                     ['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','tests/cli/Parameters.Tests.ps1','tests/cli/Parameters.Native.Tests.ps1','tests/faults/FaultIO.Tests.ps1','tests/pdf/EmailOutcome.Native.Tests.ps1']]}
        self.add_payload('manifest.json',json_bytes(manifest))
        for name,payload in self.payloads.items():self.privacy_gate(payload,name)
        self.privacy_gate(result_payload,result_path.name)
        for name,payload in self.standalone.items():self.privacy_gate(payload,name)
        files=[{'file':'docs/codex/evidence/T16-C1-reports/'+n,'sha256':sha(p)} for n,p in self.payloads.items()]
        files+=[{'file':'docs/codex/evidence/'+n,'sha256':sha(p)} for n,p in self.standalone.items()]
        files.append({'file':'docs/codex/evidence/T16-C1-results.json','sha256':sha(result_payload)})
        outcome={'task':'T16','clean_commit':self.commit,'check_only':not write,'clean_reports':28,'total_passed':1072,'per_shell':shells,
                 'historical_records':sum(r['classification'].startswith('historical') for r in self.records),'public_files':len(files),
                 'manifest_sha256':sha(self.payloads['manifest.json']),'results_sha256':sha(result_payload),'files':files,
                 'literal_whitespace_waiver_suggestions':[{'file':'docs/codex/evidence/T16-C1-reports/'+n,
                    'trailing_whitespace_lines':sum(bool(re.search(r'[ \t]+$',line)) for line in p.decode('utf-8-sig').splitlines()),
                    'reason':'Retain literal original diagnostic/source whitespace and byte binding; generated evidence only.'}
                    for n,p in self.payloads.items() if any(re.search(r'[ \t]+$',line) for line in p.decode('utf-8-sig').splitlines())]}
        if write:
            assert_clean(self.repo,self.commit);require(not destination.exists() and not result_path.exists() and all(not (evidence/n).exists() for n in self.standalone),'Public evidence is never overwritten')
            destination.mkdir()
            for path,payload in [(destination/n,p) for n,p in self.payloads.items()]+[(result_path,result_payload)]+[(evidence/n,p) for n,p in self.standalone.items()]:
                with path.open('xb') as stream:stream.write(payload)
                require(sha(path.read_bytes())==sha(payload),'Published evidence bytes differ')
        print(json.dumps(outcome,indent=2))

def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--repo',type=Path,default=Path.cwd());parser.add_argument('--commit',required=True)
    group=parser.add_mutually_exclusive_group();group.add_argument('--check-only',action='store_true');group.add_argument('--write',action='store_true')
    args=parser.parse_args();repo=args.repo.resolve();assert_clean(repo,args.commit);collector=T16Collector(repo,args.commit)
    raw,counts=load_json(collector.work/'T16-expected-counts.json');require(counts==COUNTS and tuple(counts)==TIERS,'Frozen14tier counts changed')
    collector.counts_digest=sha(raw);collector.add_payload('frozen-expected-counts.json',raw)
    shells=[collector.clean_shell(label,collector.work/('T16-C1-'+label)) for label in ('ps51','ps7')]
    collector.historical()
    for leaf in ('Run-T16Checkpoint.ps1','T16-expected-counts.json','Collect-T16Evidence.py','Collect-T15Evidence.py','Collect-T14Evidence.py','Collect-T09C3Evidence.py',
                 'T16-C1-ps51.outer.txt','T16-C1-ps7.outer.txt'):
        collector.archive(collector.work/leaf,'clean C1 execution source/capture bytes; separate from Pester case counts')
    collector.environment_and_pins(collector.work/'T16-environment.json');collector.verified_cache_receipt(collector.work/'T16-cache-verification.json')
    for shell in shells:collector.analyzer(shell['shell'],collector.work/('T16-C1-analyzer-'+shell['shell']+'.json'),shell['shell_version'])
    assert_clean(repo,args.commit);collector.finish(shells,repo/'docs/codex/evidence/T16-C1-reports',args.write)

if __name__=='__main__':main()
