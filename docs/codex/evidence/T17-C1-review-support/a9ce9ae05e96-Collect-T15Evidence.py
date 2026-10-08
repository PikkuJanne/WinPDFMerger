"""T15 evidence collector: checks only by default; --write is root's final records action.

Runs no suites or PDF engines. Existing T09 and T14 collector primitives validate
exact Pester reports and prior native receipt schemas. Dirty history is excluded.
"""
from pathlib import Path
import argparse, json, os, re, runpy, subprocess

BASE_PATH=Path(__file__).with_name('Collect-T14Evidence.py')
B=runpy.run_path(str(BASE_PATH),run_name='t15_existing_evidence_primitives')
Base=B['T14Collector']; LegacyCollector=B['LegacyCollector']; LEGACY=B['LEGACY']
require,sha,load_json,json_bytes=(B[n] for n in ('require','sha','load_json','json_bytes'))
COMMIT='53d0923c95a86ae6a44bc89bab51cac6786c1e32'
COUNTS={'Unit':335,'EmailOutcome':11,'MasterValidation':7,'Staging':9,'InputPreflight':22,'Destination':15,
        'ToolInvocation':12,'PdftkPaths':13,'GhostscriptPaths':13,'SourceDiscovery':4,'DependencyEntry':9,
        'Launcher':24,'LauncherNative':2,'NativeFixture':1,'NativeRunner':36,'FaultIO':32,'FaultRecovery':14}
TIERS=tuple(COUNTS); LegacyCollector.clean_shell.__globals__['TIERS']=TIERS
ORACLE={'python':'3.12.14','pypdfium2':'5.13.0','pdfium':'153.0.7999.0'}

def assert_clean(repo,commit):
    head=subprocess.run(['git','-C',str(repo),'rev-parse','HEAD'],capture_output=True,text=True,check=True).stdout.strip()
    status=subprocess.run(['git','-C',str(repo),'status','--porcelain=v1'],capture_output=True,text=True,check=True).stdout
    require(head==commit and not status,'Exact clean frozen T15 C1 is required before collection or writes')

class T15Collector(Base):
    def __init__(self,repo,commit):
        super().__init__(repo,commit);self.payload_bindings={};self.history_reports=set();self.history_files=set()

    def add_payload(self,name,payload):
        raw=payload
        if name.endswith('.json'):
            payload=json_bytes(self.sanitize_value(json.loads(raw.decode('utf-8-sig'))))
        self.payload_bindings[name]={'input_sha256':sha(raw),'public_sha256':sha(payload),'privacy_changed_bytes':raw!=payload}
        super().add_payload(name,payload)

    def build_receipt(self,*args,**kwargs):
        result=super().build_receipt(*args,**kwargs)
        result['build_receipt_raw_sha256']=result['build_receipt_sha256']
        result['build_receipt_sha256']=sha(self.payloads[result['build_receipt']])
        return result

    @staticmethod
    def complete_native(native):
        Base.complete_native(native)
        require(native.get('OwnershipReleased') is True,'Native ownership release must be confirmed')

    def clean_shell(self,label,root):
        aggregate=LegacyCollector.clean_shell(self,label,root)
        require(aggregate['counts_per_tier']==COUNTS and aggregate['passed']==559,'Exact seventeen-tier559 counts required')
        root=self.owned_input(root);metadata_raw,metadata=load_json(root/'collector.json')
        require(metadata['task']=='T15' and metadata['checkpoint']=='C1' and metadata['commit_under_test']==self.commit
                and metadata['dirty_worktree'] is False and metadata['shell']==label and not metadata['acquisition_performed'],
                'Clean collector metadata differs from frozen context')
        require(metadata['expected_counts_sha256']==sha(self.owned_input(metadata['expected_counts_file']).read_bytes())==self.counts_digest,
                'Collector expected count input changed')
        require(metadata['orchestration_script_sha256']==sha((self.work/'Run-T15Checkpoint.ps1').read_bytes()),'Driver source changed')
        require(metadata['child_only_modulepath_removed'] is True and metadata['suite_timeout_ms']==180000
                and metadata['stream_capture_timeout_ms']==1000,'Driver child environment or bounds differ')
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
            if tier in ('EmailOutcome','MasterValidation','Staging','InputPreflight','Destination','PdftkPaths','GhostscriptPaths',
                        'SourceDiscovery','DependencyEntry','LauncherNative','NativeFixture','FaultRecovery'):argv+=['-PdftkPath',metadata['pdftk']]
            if tier in ('EmailOutcome','Staging','InputPreflight','Destination','GhostscriptPaths','FaultRecovery'):argv+=['-GhostscriptPath',metadata['ghostscript']]
            if tier in ('EmailOutcome','MasterValidation','InputPreflight','FaultRecovery'):argv+=['-PythonPath',metadata['python']]
            require(job['arguments']==argv,'Executed argument vector differs from frozen selection')
            row.update(command_executable=self.sanitize_string(job['executable']),command_arguments=self.sanitize_value(argv),
                       started_at_utc=job['started_at_utc'],completed_at_utc=job['completed_at_utc'],elapsed_ms=job['elapsed_ms'])
            for stream,key in [('stdout','log'),('stderr','stderr_log')]:
                raw=self.owned_input(job[key]).read_bytes();clean=self.sanitize_string(raw.decode('utf-8-sig')).encode('utf-8')
                name=f'{label}-{tier}-{stream}.txt';self.add_payload(name,clean)
                row.update({stream+'_file':name,stream+'_raw_sha256':sha(raw),stream+'_sha256':sha(clean)})
            if tier in ('EmailOutcome','MasterValidation','Staging','InputPreflight','Destination'):
                self.native[(label,tier)]=Base.observations(self,label,tier,job,aggregate['shell_version'],metadata)
            elif tier in ('PdftkPaths','GhostscriptPaths'):self.path_observations(job,metadata)
            elif tier=='FaultRecovery':
                _,summary=load_json(self.owned_input(job['report'])/'summary.json')
                row.update(self.build_receipt(summary['native_fixture_build_receipt'],summary['native_fixture_build_receipt_sha256'],
                                              label,tier,'FakeNative.exe','tests/native/FakeNative.cs'))
                self.fault_recovery(label,job,metadata,aggregate['shell_version'])
            elif tier=='FaultIO':self.fault_io(label,job)
        require(self.record_for(label,'PdftkPaths')['native_version']=='2.02' and
                self.record_for(label,'GhostscriptPaths')['native_version']=='10.08.0','Actual engine version mismatch')
        return aggregate

    def fault_recovery(self,label,job,metadata,version):
        stdout=self.owned_input(job['log']).read_text(encoding='utf-8-sig')
        matches=re.findall(r'(?m)^Fault recovery observations: (.+?)\r?$',stdout);require(len(matches)==1,'One fault receipt required')
        path=self.owned_input(matches[0]);raw,obs=load_json(path)
        require(obs['CommitUnderTest']==self.commit and obs['DirtyWorktree'] is False and obs['ShellVersion']==version
                and obs['StandardUser'] and obs['Process64Bit'] and obs['PdfTkVersion']=='2.02' and obs['GhostscriptVersion']=='10.08.0'
                and obs['OracleVersions']==ORACLE,'Fault recovery context/pins differ')
        require(obs['OuterCallerBefore']==obs['OuterCallerAfter'] and obs['OuterPathBeforeSHA256']==obs['OuterPathAfterSHA256'],
                'Outer caller environment changed')
        _,pdftk=load_json(self.repo/'docs/codex/evidence/T03-pdftk-acquisition.json')
        _,gs=load_json(self.repo/'docs/codex/evidence/T09-gs-acquisition.json')
        expected_hashes={Path(x['relative_path']).name:x['sha256'] for x in pdftk['extracted_files']}
        expected_hashes.update({Path(x['relative_path']).name:x['sha256'] for x in gs['ghostscript_extraction']['selected_files']})
        require({x['Name']:x['SHA256'] for x in obs['EngineSHA256']}==expected_hashes,'Fault vendor executable/DLL pins differ')
        for leaf,digest in [('independent-fault-inspection.py','OracleSHA256'),('original-fault-raster.py','GeneratorSHA256')]:
            source=self.owned_input(path.parent/leaf).read_bytes();require(sha(source)==obs[digest].lower(),'Fault development helper changed')
            self.add_payload(f'{label}-FaultRecovery-{leaf}',source)
        _,build=load_json(self.owned_input(obs['ControlledFixtureBuildReceipt']))
        require(build['source_sha256']==obs['ControlledFixtureSourceSHA256']==sha((self.repo/'tests/native/FakeNative.cs').read_bytes())
                and build['executable_sha256']==obs['ControlledFixtureSHA256'],'Controlled fixture source/version mismatch')
        expected={f'environment-{state}-{mode}' for state in ('unset','empty','value') for mode in ('success','start-failure','log-fault')}
        expected.update(('cancel-before-master','cancel-after-master','cancel-after-real-master-staged-inspection-before-move',
                         'controlled-nested-tree-timeout','controlled-nested-tree-parent-exits-first'))
        rows=obs['Observations'];require(len(rows)==14 and {r['Label'] for r in rows}==expected,'Exact fourteen fault observations required')
        for r in rows:
            name=r['Label'];proof=r.get('Proof')
            if 'SourceAndForeignBefore' in r:
                require(r['SourceAndForeignBefore']==r['SourceAndForeignAfter'],'Source or foreign snapshot changed')
                self.snapshot(r['SourceAndForeignAfter'],current=True)
            if proof:
                cap=proof['Capture'];require(cap['CallerBefore']==cap['CallerAfter'],'Actual child caller environment changed')
                for call in cap['NativeCalls']:
                    require(call['CallerBefore']==call['CallerAfter']==cap['CallerBefore'],'Native call changed caller environment')
                    native=call['Result'];require(native.get('OwnershipReleased') is True,'Native ownership remains unresolved')
                    if call['MasterBefore']:
                        require(call['MasterBefore']==call['MasterAfter'],'Published master changed during email work')
                        self.snapshot(json.dumps(call['MasterAfter']),current=True)
                reads=proof['FinalReads'];reads=reads if isinstance(reads,list) else ([] if reads is None else [reads])
                for read in reads:
                    final=self.owned_input(read['Path']);require(sha(final.read_bytes())==read['SHA256'].lower(),'Retained final changed')
                    require(read['PdfTk']['ExitCode']==read['OracleExit']==0 and read['Oracle']['page_count']==1
                            and read['Oracle']['pages']==[{'identifier':'T03-15-P01' if 'log-fault' in name else 'T03-01-P01',
                                                          'rotation_degrees':0,'size_points':[432,288]}],'Actual separate PDFtk/PDFium final reread differs')
                require(not list(self.owned_input(proof['Output']).glob('.WinPDFMerge*')),'Recovered entry left staging residue')
                log=self.owned_input(proof['LogPath']).read_bytes().decode('utf-8-sig');require(log==proof['Log'],'Retained log changed')
                if name.startswith('environment-'):
                    state=r['RequestedOSState'];mode=r['ControlledMode'];require(cap['CallerBefore']['State']==state and
                        cap['CallerBefore']['Present']==(state!='unset'),'Requested genuine OS state not observed')
                    if state=='empty':require(cap['CallerBefore']['Value']=='','Genuine empty state collapsed')
                    calls=[c for c in cap['NativeCalls'] if c['Phase']=='email'];require(len(calls)==1,'One email attempt required')
                    call=calls[0];self.complete_native(call['EnvironmentProbe']);require(call['EnvironmentProbe']['Stdout'].strip()=='<unset>'
                        and call['RemovedEnvironmentVariables']==['GS_OPTIONS'],'Child-only removal not demonstrated')
                    require(proof['Result']['ExitCode']==(0 if mode=='success' else 2) and len(reads)==(2 if mode=='log-fault' else 1),
                            'Entry failure/publication outcome mismatch')
                    if mode=='start-failure':require(not call['Result']['Started'] and call['Result']['LaunchError'] and call['ControlledSubstitution'],
                                                   'Actual controlled OS start failure missing')
                    else:self.complete_native(call['Result']);require(call['Result']['Executable']==metadata['ghostscript'],'Real selected GS missing')
                    if mode=='log-fault':require(cap['LogFaultReached'] and cap['Outcome']['ExitCode']==2 and cap['Outcome']['EmailState']=='published'
                        and cap['OutcomeParameters']['RunFailed'] and len(cap['Outcome']['PublishedPaths'])==2
                        and self.owned_input(reads[1]['Path']).stat().st_size<self.owned_input(reads[0]['Path']).stat().st_size,
                        'Logging failure lost truthful published email/master result')
                elif name.startswith('cancel-'):
                    require(cap['CancellationRequested'] and proof['Result']['ExitCode']==(2 if name=='cancel-after-master' else 1)
                            and len(reads)==(1 if name=='cancel-after-master' else 0),'Cancellation outcome/final inventory mismatch')
                    if name.startswith('cancel-after-real-'):
                        require(cap['FastCancelReached'],'Fast-success token seam not reached')
                        actual=[c['Result'] for c in cap['NativeCalls'] if c['Phase']=='master' or
                                ('dump_data_utf8' in c['RequestedArguments'] and c['RequestedArguments'][0].endswith('master.pdf'))]
                        require(len(actual)==2,'Real merge and staged inspection missing');[self.complete_native(n) for n in actual]
                    else:
                        calls=[c for c in cap['NativeCalls'] if c['ControlledSubstitution'] and 'owned parent-child-grandchild' in c['ControlledSubstitution']]
                        require(len(calls)==1,'Controlled tree conversion missing');call=calls[0];native=call['Result']
                        require(native['Cancelled'] and not native['TimedOut'] and not native['Succeeded'] and native['Started']
                                and call['StagedExistsBeforeEntryCleanup'] and call['StagedBytes']>0 and not Path(call['StagePath']).exists(),
                                'Cancelled owned partial was not released before known cleanup')
            if 'Tree' in r:
                tree=r['Tree'];require(len(tree)==3 and len({x['pid'] for x in tree})==3 and tree[1]['parent_pid']==tree[0]['pid']
                                      and tree[2]['parent_pid']==tree[1]['pid'] and all(x['start_utc_ticks']>0 for x in tree)
                                      and r['UnrelatedSurvived']['Alive'] is True,'Owned tree identity/unrelated preservation evidence missing')
                if name.startswith('controlled-nested-'):
                    native=r['NativeResult'];require(native['OwnershipReleased'] and not native['TerminationError'] and not native['CaptureError'],
                                                   'Nested owned launch did not release')
                    if name.endswith('timeout'):require(native['TimedOut'] and not native['Succeeded'],'Actual owned timeout missing')
                    else:self.complete_native(native)
        name=f'{label}-FaultRecovery-observations.json';clean=json_bytes(self.sanitize_value(obs));self.add_payload(name,clean)
        self.record_for(label,'FaultRecovery').update(observations_file=name,raw_observations_sha256=sha(raw),observations_sha256=sha(clean),observation_count=14)

    def fault_io(self,label,job):
        log=self.owned_input(job['log']).read_text(encoding='utf-8-sig')
        matches=re.findall(r'(?m)^Fault IO observations:\r?\n(\[[^\r\n]*\])\r?$',log)
        require(len(matches)==1,'FaultIO inline observations missing');obs=json.loads(matches[0]);require(len(obs)==27,'FaultIO32cases require27disclosed receipts')
        for row in obs:
            if 'Before' in row and 'After' in row:require(row['Before']==row['After'],'Controlled IO changed preserved snapshot')
        raw=matches[0].encode('utf-8');name=f'{label}-FaultIO-observations.json';self.add_payload(name,raw)
        self.record_for(label,'FaultIO').update(observations_file=name,raw_observations_sha256=sha(raw),observations_sha256=sha(self.payloads[name]),
                                               observation_count=27,scope='Controlled unit receipts;4outcome+1lockedlog cases lack receipts; no PDF engine support pass')

    def archive(self,path,classification='historical dirty evidence; excluded from clean totals',name=None):
        path=self.owned_input(path)
        if path in self.history_files:return
        self.history_files.add(path);raw=path.read_bytes()
        name=name or ('historical-'+sha(str(path.relative_to(self.work)).encode())[:12]+'-'+path.name)
        if path.suffix=='.xml':
            _,summary=load_json(path.parent/'summary.json')
            if 'Total' in summary:summary={'total':summary['Total'],'failed':summary['Failed']}
            clean=self.sanitized_xml(raw,False,summary)
        elif path.suffix=='.json':clean=json_bytes(self.sanitize_value(json.loads(raw.decode('utf-8-sig'))))
        else:clean=self.sanitize_string(raw.decode('utf-8-sig')).encode('utf-8')
        self.add_payload(name,clean)
        self.records.append({'classification':classification,'source_relative_path':str(path.relative_to(self.work)).replace('\\','/'),
                             'file':name,'raw_sha256':sha(raw),'sha256':sha(self.payloads[name])})
        if path.name=='summary.json':
            summary=json.loads(raw.decode('utf-8-sig'));record=self.records[-1]
            if 'total' in summary:
                record.update(commit_under_test=summary['commit_under_test'],dirty_worktree=summary['dirty_worktree'],
                              counts={k:summary[k] for k in LEGACY['COUNTS']})
            elif 'Total' in summary:record.update(commit_under_test=summary['CommitUnderTest'],dirty_worktree=summary['DirtyWorktree'],
                counts={k:summary[v] for k,v in [('passed','Passed'),('failed','Failed'),('failed_blocks','FailedBlocks'),
                                               ('failed_containers','FailedContainers'),('skipped','Skipped'),('not_run','NotRun'),('total','Total')]})

    def historical(self):
        for path in sorted(self.work.glob('T15-dirty-*')):
            if path.is_file():self.archive(path)
            elif path.is_dir():
                for item in sorted(path.iterdir()):
                    if item.is_file():self.archive(item)
        for path in sorted(self.history_files):
            if path.suffix=='.txt' and ('stdout' in path.name or path.name.endswith('-1.txt') or path.name.endswith('-2.txt')):
                log=path.read_text(encoding='utf-8-sig')
                for match in re.findall(r'(?m)^Reports: (.+?)\r?$',log):
                    report=self.owned_input(match);self.history_reports.add(report)
                    for leaf in ('summary.json','results.xml'):self.archive(report/leaf)
                    _,summary=load_json(report/'summary.json');require(summary['dirty_worktree'] is True,'Historical summary must remain dirty')
                    if 'native_fixture_build_receipt' in summary:
                        build_path=self.owned_input(summary['native_fixture_build_receipt']);build_raw,build=load_json(build_path)
                        require(sha(build_raw)==summary['native_fixture_build_receipt_sha256'] and
                                sha((build_path.parent/'FakeNative.exe').read_bytes())==build['executable_sha256'],
                                'Historical actual compiled fixture/build binding changed')
                        self.archive(build_path)
                for match in re.findall(r'(?m)^(?:Fault recovery|Email) observations: (.+?)\r?$',log):self.archive(Path(match))
        for root in sorted(self.work.glob('T15-focused-FaultIO-*')):
            _,execution=load_json(root/'execution.json');require(execution['exit_code']==0,'Historical FaultIO unexpectedlyfailed')
            for leaf,digest in execution['files'].items():require(sha((root/leaf).read_bytes())==digest,'Focused raw binding changed')
            _,summary=load_json(root/'summary.json');require(summary['DirtyWorktree'] and summary['Total']==summary['Passed'] in (28,32),
                'Exact historical FaultIO counts changed')
            for item in sorted(root.iterdir()):
                if item.is_file():self.archive(item)
        for name in ('T15-native-dirty-smoke-history.json','T15-M2-runtime-review-dirty.json','T15-precommit-analyzer-ps51.json',
                     'T15-precommit-analyzer-ps7.json','T15-owned-launch-smoke-ps51.json','T15-owned-launch-smoke-final-ps51.json',
                     'T15-owned-launch-smoke-final-ps7.json','T15-OwnedNativeLaunch-initial.txt','T15-OwnedNativeLaunch.txt',
                     'T15-OwnedNativeSmoke-initial.ps1','T15-OwnedNativeSmoke.ps1','Invoke-T15NativeSmoke.ps1','Run-T15Precommit.py',
                     'Run-T15FaultIOFocus.ps1','Run-T15FaultIOFocus.py'):
            self.archive(self.work/name)
        _,index=load_json(self.work/'T15-native-dirty-smoke-history.json')
        for attempt in index['Attempts']:
            for binding in attempt['Bindings']:
                source=self.repo/binding['Path'];require(sha(source.read_bytes())==binding['SHA256'],'Historical native raw binding changed');self.archive(source)
        for path in self.work.glob('T15-owned-launch-smoke*.json'):
            _,receipt=load_json(path);require(receipt['Result']=='pass' and len(receipt['Cases'])==8 and all(c['Pass'] for c in receipt['Cases']),
                                            'Historical adapter smoke counts changed')
            for field,digest in [('AdapterSource','AdapterSHA256'),('SmokeSource','SmokeSHA256'),('Fixture','FixtureSHA256')]:
                source=self.repo/receipt[field]
                if sha(source.read_bytes())!=receipt[digest] and path.name=='T15-owned-launch-smoke-ps51.json' and field!='Fixture':
                    source=source.with_name(source.stem+'-initial'+source.suffix)
                    self.records.append({'classification':'historical adapter source version binding; excluded from clean totals',
                                         'receipt':path.name,'original_source_path':receipt[field],
                                         'preserved_initial_snapshot':str(source.relative_to(self.repo)).replace('\\','/'),
                                         'source_sha256':receipt[digest],
                                         'note':'Original smoke recorded then-current path; exact initial bytes were retained separately before subsequent adapter refinement.'})
                require(sha(source.read_bytes())==receipt[digest],'Historical adapter source/fixture changed')
            build=(self.repo/receipt['Fixture']).parent/'build-info.json';require(sha(build.read_bytes())==receipt['FixtureBuildReceiptSHA256'],'Adapter buildreceipt changed');self.archive(build)

    def analyzer(self,label,path,version):
        raw,report=load_json(self.owned_input(path));require(report['Task']=='T15' and report['Phase']=='C1' and report['CommitUnderTest']==self.commit
             and report['DirtyWorktree'] is False and report['ShellVersion']==version and report['AnalyzerVersion']=='1.25.0','Clean static context differs')
        for severity,key in [(2,'Errors'),(1,'Warnings'),(0,'Information')]:require(report[key]==sum(f['Severity']==severity for f in report['Findings']),
                                                                                'Static counts differ from raw findings')
        require(report['Errors']==0,'Static errors remain');name=f'{label}-PSScriptAnalyzer-findings.json';self.add_payload(name,raw)
        self.records.append({'classification':'clean implementation static findings, warnings/info retained; not test passes','shell':label,
                             'file':name,'raw_sha256':sha(raw),'sha256':sha(self.payloads[name]),'analyzer_version':'1.25.0',
                             'error_count':report['Errors'],'warning_count':report['Warnings'],'information_count':report['Information']})

    def verified_cache_receipt(self,path):
        raw,receipt=load_json(self.owned_input(path));require(receipt['task']=='T15' and not receipt['acquisition_performed'],'Cache reuse scope changed')
        require({r['dependency']:r['version'] for r in receipt['dependencies']}=={'Pester':'6.2.0','PDFtk':'pdftk 2.02','Ghostscript':'10.08.0',
                  'PowerShell7':'7.6.6','PSScriptAnalyzer':'1.25.0'},'Cache version pins changed')
        for row in receipt['dependencies']:
            require(row['all_selected_file_hashes_match'] and row['selected_files_verified']==len(row['selected_files'])
                    and sha((self.repo/row['source_receipt']).read_bytes())==row['source_receipt_sha256'],'Cache source receipt changed')
            root=Path(os.path.expandvars(row['cache_root'].replace('<USERPROFILE>',os.environ['USERPROFILE']))).resolve()
            for selected in row['selected_files']:
                candidate=(root/selected['relative_path']).resolve();require(candidate.is_relative_to(root) and sha(candidate.read_bytes())==selected['sha256'],
                                                                            'Previously approved selected bytes changed')
        require(sum(r['selected_files_verified'] for r in receipt['dependencies'])==14,'Exact14selected cache inputs required')
        require(sha((self.repo/receipt['previous_full_audit']).read_bytes())==receipt['previous_full_audit_sha256'],
                'Previously retained full approved-cache audit changed')
        self.oracle_runtime=receipt['development_oracle_runtime'];require({k:self.oracle_runtime[k] for k in ORACLE}==ORACLE,'Development oracle pins changed')
        for key in ('python','pdfium_dll'):
            path=Path(os.path.expandvars(self.oracle_runtime[key+'_path'].replace('<USERPROFILE>',os.environ['USERPROFILE'])))
            require(sha(path.read_bytes())==self.oracle_runtime[key+'_sha256'],'Development runtime byte pin changed')
        self.add_payload('approved-cache-verification.json',raw);self.standalone['T15-cache-verification.json']=raw

    def environment_and_pins(self,path):
        LegacyCollector.environment_and_pins(self,path);self.standalone['T15-environment.json']=self.owned_input(path).read_bytes()
        _,receipt=load_json(self.owned_input(path));require(receipt['commit_under_inventory']==self.commit and receipt['dirty_worktree'] is False
            and receipt['task']=='T15' and receipt['application_or_native_test_claim'] is False,'Fresh ordinary inventory context mismatch')
        for leaf,digest in [('T15-InventoryCommand.ps1','inventory_source_sha256'),('T15-environment.stdout.txt','raw_stdout_sha256'),
                            ('T15-environment.stderr.txt','raw_stderr_sha256')]:
            source=self.work/leaf;require(sha(source.read_bytes())==receipt[digest],'Fresh inventory source/capture binding changed')
            self.archive(source,'separate read-only ordinary host inventory; no application acceptance pass')

    def results_document(self,shells,destination):
        result=Base.results_document(self,shells,destination)
        for key in ('ac032','ac033','ac034'):result.pop(key,None)
        result.update(task='T15',ac035='pass',ac036='pass',ac037='pass',cases_per_shell=COUNTS,total_passed=1118,
            selected_tiers_in_execution_order=list(TIERS),selected_cache_verification='docs/codex/evidence/T15-cache-verification.json',
            ordinary_environment_receipt='docs/codex/evidence/T15-environment.json',
            scope='AC035 integration: actual Win32 unset/empty/value survives real success, controlled invalid-image OS start failure and controlled logging faults in both required shells. AC036 integration: owned nested process cancellation before/after master and after real staged inspection; unrelated same-image process and validated published master survive. AC037 unit: controlled locks/denied/full/write/log/cleanup faults yield truthful publication outcomes. All earlier affected regressions retained.',
            historical_rule='Dirty failures, preparation-only missing XML/summary, corrected focused passes, adapter8case smokes and static warnings are retained separately, never added to clean1118.',
            collector_command_template='<explicit shell> -NoProfile -ExecutionPolicy RemoteSigned -File tests/.work/Run-T15Checkpoint.ps1 -ShellLabel ps51|ps7 -ExpectedCommit '+self.commit+' -ExpectedCountsPath tests/.work/T15-expected-counts.json -Checkpoint C1',
            evidence_collector_command='<approved Python> -B tests/.work/Collect-T15Evidence.py --repo . --commit '+self.commit+' [--write]',
            limitations=['Controlled copied-helper token/logger/invalid-image/compiled-tree substitutions are disclosed; no physical keyboard Ctrl+C, Explorer interruption or hard-crash cleanup guarantee.',
                         'No public cancellation option is introduced. Unresolved native ownership retains known staging rather than deleting active-owned data.',
                         'Retained T14 controlled corrupt GS input and4096owned whitespace fixture preparation remain explicitly scoped.',
                         'Structure/visible-page checks do not certify full fidelity, security, signatures, PDF/A, OS/UNC/Explorer/CI/package/release gates; T16/T17 remain open.',
                         'Static warnings/information retained; no lint-clean claim. Previously authorized caches reused without acquisition/admin/persistent policy/PATH/security changes.'])
        result['command_template']='Per-tier exact executable/argument vectors are retained in manifest and sanitized collector/runs metadata.'
        failed=self.work/'T15-collector-check-7b67256b192747f3a927c97177cb60fe'
        _,receipt=load_json(failed/'execution.json')
        result['collector_preparation_history']={
            'classification':'Evidence collector preparation only; no application/suite failure or acceptance counts',
            'retained_failed_attempt_source_sha256':receipt['CollectorSHA256'],'exit_code':receipt['ExitCode'],
            'stdout_sha256':receipt['StdoutSHA256'],'stderr_sha256':receipt['StderrSHA256'],
            'execution_receipt_sha256':sha((failed/'execution.json').read_bytes()),
            'diagnosis':'Python read_text newline translation caused exact retained CRLF log versus captured .NET string comparison failure. Corrected to raw bytes decoded UTF8; original application bytes/receipts were intact.',
            'unretained_exploratory_checks':'Earlier incomplete-driver count gate and initial-source snapshot mapping checks were read-only authoring probes; no fabricated report/XML/counts claimed.'}
        return result

    def finish(self,shells,destination,write):
        require(sum(s['passed'] for s in shells)==1118,'Exact1118clean total required')
        destination=destination.resolve();evidence=self.repo/'docs/codex/evidence'
        require(destination==evidence/'T15-C1-reports','Exact root-authorized public destination required')
        results_path=evidence/'T15-C1-results.json';results_payload=json_bytes(self.sanitize_value(self.results_document(shells,destination)))
        manifest={'schema_version':1,'task':'T15','checkpoint':'C1','commit_under_test':self.commit,'dirty_worktree':False,
                  'clean_reports':34,'total_clean_passed':1118,'shells':shells,'records':self.records,
                  'xml_redactions':['environment.'+k for k in LEGACY['XML_IDENTITY']],
                  'privacy_scope':'Only XML machine/user/domain/cwd and repository/user/cache/temp path prefixes sanitized; synthetic IDs, whitespace, arguments, diagnostics and native results retained.',
                  'exact_byte_policy':'Every source/raw and public/sanitized SHA is separately bound. JSON containing absolute fixture/profile paths is sanitized, including NativeRunner/FaultRecovery summaries; no private raw path bytes are published.',
                  'evidence_collector_sha256':sha(Path(__file__).read_bytes()),'legacy_primitives_sha256':sha(B['LEGACY_PATH'].read_bytes()),
                  'prior_native_schema_primitives_sha256':sha(BASE_PATH.read_bytes()),'payload_bindings':self.payload_bindings.copy(),
                  'results_file':'docs/codex/evidence/T15-C1-results.json','results_sha256':sha(results_payload),
                  'historical_rule':'Unavailable historical XML/summary/counts are explicitly absent, never fabricated; dirty history/adapter/static checks excluded from1118.',
                  'standalone_receipts':[{'file':'docs/codex/evidence/'+n,'sha256':sha(p)} for n,p in self.standalone.items()]}
        source_paths=['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','tests/native/FakeNative.cs',
                      'tests/native/NativeRunner.Tests.ps1','tests/faults/FaultRecovery.Native.Tests.ps1','tests/faults/FaultIO.Tests.ps1',
                      'tests/pdf/EmailOutcome.Native.Tests.ps1']
        manifest['frozen_implementation_source_bytes']=[{'path':p,'sha256':sha((self.repo/p).read_bytes())} for p in source_paths]
        self.add_payload('manifest.json',json_bytes(manifest))
        for name,payload in self.payloads.items():self.privacy_gate(payload,name)
        self.privacy_gate(results_payload,results_path.name)
        for name,payload in self.standalone.items():self.privacy_gate(payload,name)
        outcome={'task':'T15','clean_commit':self.commit,'check_only':not write,'clean_reports':34,'total_passed':1118,
                 'per_shell':shells,'historical_records':sum(r['classification'].startswith('historical') for r in self.records),
                 'public_files':len(self.payloads)+len(self.standalone)+1,'manifest_sha256':sha(self.payloads['manifest.json']),
                 'results_sha256':sha(results_payload),'files':[{'file':'docs/codex/evidence/T15-C1-reports/'+n,'sha256':sha(p)} for n,p in self.payloads.items()]}
        outcome['literal_whitespace_waiver_suggestions']=[
            {'file':'docs/codex/evidence/T15-C1-reports/'+name,
             'trailing_whitespace_lines':sum(bool(re.search(r'[ \t]+$',line)) for line in payload.decode('utf-8-sig').splitlines()),
             'reason':'Preserve literal original diagnostic/source whitespace with its byte-hash binding; generated evidence only.'}
            for name,payload in self.payloads.items()
            if any(re.search(r'[ \t]+$',line) for line in payload.decode('utf-8-sig').splitlines())]
        if write:
            assert_clean(self.repo,self.commit);require(not destination.exists() and not results_path.exists() and
                all(not (evidence/n).exists() for n in self.standalone),'Public evidence is never overwritten')
            destination.mkdir()
            for path,payload in [(destination/n,p) for n,p in self.payloads.items()]+[(results_path,results_payload)]+[(evidence/n,p) for n,p in self.standalone.items()]:
                with path.open('xb') as f:f.write(payload)
                require(sha(path.read_bytes())==sha(payload),'Published evidence bytes differ')
        print(json.dumps(outcome,indent=2))

def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--repo',type=Path,default=Path.cwd());parser.add_argument('--commit',default=COMMIT)
    parser.add_argument('--environment',type=Path);group=parser.add_mutually_exclusive_group();group.add_argument('--check-only',action='store_true');group.add_argument('--write',action='store_true')
    args=parser.parse_args();require(args.commit==COMMIT,'Only frozen reviewed T15 C1 accepted');repo=args.repo.resolve();assert_clean(repo,args.commit)
    collector=T15Collector(repo,args.commit);raw,counts=load_json(collector.work/'T15-expected-counts.json')
    require(counts==COUNTS and tuple(counts)==TIERS,'Frozen expected counts changed');collector.counts_digest=sha(raw);collector.add_payload('frozen-expected-counts.json',raw)
    shells=[collector.clean_shell(label,collector.work/('T15-C1-'+label)) for label in ('ps51','ps7')]
    collector.historical()
    failure=collector.work/'T15-collector-check-7b67256b192747f3a927c97177cb60fe'
    _,failed=load_json(failure/'execution.json');require(failed['ExitCode']==1,'Retained collector preparation exit changed')
    for leaf,key in [('stdout.txt','StdoutSHA256'),('stderr.txt','StderrSHA256')]:
        require(sha((failure/leaf).read_bytes())==failed[key],'Collector preparation raw capture changed')
    for leaf in ('execution.json','stdout.txt','stderr.txt'):
        collector.archive(failure/leaf,'historical collector preparation only; no application/suite acceptance count')
    for leaf in ('Run-T15Checkpoint.ps1','T15-expected-counts.json','Collect-T15Evidence.py','Collect-T14Evidence.py','Collect-T09C3Evidence.py',
                 'T15-C1-ps51.stdout.txt','T15-C1-ps51.stderr.txt','T15-C1-ps7.stdout.txt','T15-C1-ps7.stderr.txt'):
        collector.archive(collector.work/leaf,'clean C1 execution source/capture bytes; separate from Pester case counts')
    for shell in shells:collector.analyzer(shell['shell'],collector.work/('T15-C1-analyzer-'+shell['shell']+'.json'),shell['shell_version'])
    collector.environment_and_pins(args.environment or collector.work/'T15-environment.json');collector.verified_cache_receipt(collector.work/'T15-cache-verification.json')
    assert_clean(repo,args.commit);collector.finish(shells,repo/'docs/codex/evidence/T15-C1-reports',args.write)

if __name__=='__main__':main()
