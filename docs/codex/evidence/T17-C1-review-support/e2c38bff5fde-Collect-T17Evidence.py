"""T17 collector: read-only check by default; root alone authorizes --write.

Runs no application, suites or PDF engines. Reuses reviewed legacy Pester,
privacy, build and native schemas. Dirty history never contributes to totals.
"""
from pathlib import Path
import argparse,json,os,re,runpy,subprocess
from decimal import Decimal, localcontext

PRIOR_PATH=Path(__file__).with_name('Collect-T15Evidence.py')
P=runpy.run_path(str(PRIOR_PATH),run_name='t17_reviewed_prior_primitives')
Prior=P['T15Collector'];Base=P['Base'];Legacy=P['LegacyCollector'];LEGACY=P['LEGACY']
require,sha,load_json,json_bytes=(P[n] for n in ('require','sha','load_json','json_bytes'))
COUNTS={'Unit':335,'SizeReporting':32,'SizeReportingNative':11,'Parameters':31,'ParametersNative':9,'EmailOutcome':11,'MasterValidation':7,'Staging':9,
        'InputPreflight':22,'Destination':15,'ToolInvocation':12,'GhostscriptPaths':13,'Launcher':24,
        'LauncherNative':2,'FaultIO':32,'FaultRecovery':14}
TIERS=tuple(COUNTS);Legacy.clean_shell.__globals__['TIERS']=TIERS
ORACLE={'python':'3.12.14','pypdfium2':'5.13.0','pdfium':'153.0.7999.0'}
FROZEN_PATH=Path(__file__).with_name('T17-C1-commit.txt')
FROZEN_COMMIT=FROZEN_PATH.read_text(encoding='utf-8-sig').strip() if FROZEN_PATH.exists() else None

def assert_clean(repo,commit):
    head=subprocess.run(['git','-C',str(repo),'rev-parse','HEAD'],capture_output=True,text=True,check=True).stdout.strip()
    status=subprocess.run(['git','-C',str(repo),'status','--porcelain=v1'],capture_output=True,text=True,check=True).stdout
    require(FROZEN_COMMIT is not None and commit==FROZEN_COMMIT and head==commit and not status,'Exact reviewed clean frozen T17 C1 required')

class T17Collector(Prior):
    def __init__(self,repo,commit):
        super().__init__(repo,commit);self.support=[];self.preparation_history=[];self.size_observations={}

    def support_file(self,path,expected=None):
        path=self.owned_input(path);raw=path.read_bytes();digest=sha(raw)
        if expected is not None:require(digest==expected.lower(),'Retained supporting raw bytes changed: '+path.name)
        row={'source_relative_path':path.relative_to(self.work).as_posix(),'raw_sha256':digest,'bytes':len(raw),
             'publication':'Raw supporting original retained ignored; structured observation/decision values are archived separately.'}
        if row not in self.support:self.support.append(row)
        return row

    def clean_shell(self,label,root):
        aggregate=Legacy.clean_shell(self,label,root)
        require(aggregate['counts_per_tier']==COUNTS and aggregate['passed']==579,'Exact sixteen-tier579 counts required')
        root=self.owned_input(root);metadata_raw,metadata=load_json(root/'collector.json')
        require(metadata['task']=='T17' and metadata['checkpoint']=='C1' and metadata['commit_under_test']==self.commit
                and metadata['dirty_worktree'] is False and metadata['shell']==label and not metadata['acquisition_performed'],
                'Clean collector metadata differs from frozen context')
        require(metadata['expected_counts_sha256']==sha(self.owned_input(metadata['expected_counts_file']).read_bytes())==self.counts_digest,
                'Driver expected count input changed')
        require(metadata['orchestration_script_sha256']==sha((self.work/'Run-T17Checkpoint.ps1').read_bytes()),'Driver source changed')
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
            if tier in ('SizeReportingNative','ParametersNative','EmailOutcome','MasterValidation','Staging','InputPreflight','Destination','GhostscriptPaths','LauncherNative','FaultRecovery'):argv+=['-PdftkPath',metadata['pdftk']]
            if tier in ('SizeReportingNative','ParametersNative','EmailOutcome','Staging','InputPreflight','Destination','GhostscriptPaths','FaultRecovery'):argv+=['-GhostscriptPath',metadata['ghostscript']]
            if tier in ('SizeReportingNative','ParametersNative','EmailOutcome','MasterValidation','InputPreflight','FaultRecovery'):argv+=['-PythonPath',metadata['python']]
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
            elif tier=='SizeReportingNative':self.size_native(label,job,metadata,aggregate['shell_version'])
            elif tier=='SizeReporting':self.size_unit(label,job,metadata)
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

    @staticmethod
    def size_lines(master,candidate=None,published=False):
        require(type(master) is int and master>0,'Exact positive master bytes required')
        def human(value):
            units=('B','KiB','MiB','GiB','TiB','PiB','EiB');unit=0;value=Decimal(value)
            while value>=1024 and unit<len(units)-1:value/=1024;unit+=1
            return (str(value.quantize(Decimal('1'))) if unit==0 else format(value,'.2f'))+' '+units[unit]
        lines=[f'Master size: {master} bytes ({human(master)}).']
        if candidate is not None:
            require(type(candidate) is int and candidate>0,'Exact positive candidate bytes required')
            with localcontext() as context:
                context.prec=50;percent=Decimal(100)*(Decimal(master)-Decimal(candidate))/Decimal(master)
                display=format(percent,'.1f')
            if published:lines+=[f'Email size: {candidate} bytes ({human(candidate)}).',f'Email reduction: {display}%.']
            else:lines+=[f'Validated email candidate size: {candidate} bytes ({human(candidate)}); not published.',
                       f'Email candidate reduction: {display}% (no size benefit; candidate not published).']
        return lines

    def size_snapshots(self,rows):
        for row in rows:
            path=Path(row['Path']).resolve()
            if path.is_relative_to(self.work):self.snapshot_rows([row])
            else:
                require(path==(self.repo/'tests/fixtures/numbered/1.pdf').resolve(),'Only known tracked original may be outside ignored evidence')
                require(sha(path.read_bytes())==row['SHA256'].lower() and path.stat().st_size==row['Length']
                    and path.stat().st_mtime_ns//100+621355968000000000==row['ModifiedUtcTicks']
                    and path.stat().st_file_attributes==row['Attributes'],'Original tracked fixture snapshot changed')

    def size_native(self,label,job,metadata,version):
        text=self.owned_input(job['log']).read_bytes().decode('utf-8-sig')
        matches=re.findall(r'(?m)^Size reporting observations: (.+?)\r?$',text);require(len(matches)==1,'One native size receipt required')
        path=self.owned_input(matches[0]);raw,obs=load_json(path)
        require(obs['CommitUnderTest']==self.commit and obs['DirtyWorktree'] is False and obs['ShellVersion']==version
            and obs['StandardUser'] and obs['Process64Bit'] and obs['PdfTkVersion']=='2.02' and obs['GhostscriptVersion']=='10.08.0'
            and obs['OracleVersions']==ORACLE,'Native size context/engine/oracle pins differ')
        require(obs['TestSourceSHA256'].lower()==sha((self.repo/'tests/pdf/SizeReporting.Native.Tests.ps1').read_bytes()),'Native size source differs')
        manifest_raw,manifest=load_json(self.repo/'tests/fixtures/presets/manifest.json')
        require(obs['FixtureManifest']==manifest and obs['FixtureManifestSHA256'].lower()==sha(manifest_raw)
            and obs['GeneratorSHA256'].lower()==manifest['generator_sha256']==sha((self.repo/'tests/fixtures/presets/generate_presets.py').read_bytes()),'Original recipe/manifest binding differs')
        require(manifest['license']=='CC0-1.0' and manifest['versions']=={'python':'3.12.14','reportlab':'4.4.9','pillow':'12.3.0'}
            and manifest['scan_pixel_dimensions']==[2250,2800] and manifest['scan_seed']==170041,'Original generation provenance differs')
        require(obs['GeneratorResult']['ExitCode']==0 and json.loads(obs['GeneratorResult']['Stdout'])==manifest
            and obs['GeneratorCommand']==['-B',str(self.repo/'tests/fixtures/presets/generate_presets.py'),'--output',obs['OriginalCorpusDirectory']],
            'Actual development fixture generation command/result differs')
        require(obs['PythonSHA256'].lower()==sha(Path(metadata['python']).read_bytes()),'Development Python pin differs')
        _,pdf=load_json(self.repo/'docs/codex/evidence/T03-pdftk-acquisition.json');_,gs=load_json(self.repo/'docs/codex/evidence/T09-gs-acquisition.json')
        expected={Path(x['relative_path']).name:x['sha256'] for x in pdf['extracted_files']}
        expected.update({Path(x['relative_path']).name:x['sha256'] for x in gs['ghostscript_extraction']['selected_files']})
        require({x['Name']:x['SHA256'] for x in obs['EngineSHA256']}==expected,'Native size engine/DLL pins differ')
        oracle=path.parent/'independent-size-inspection.py';require(sha(oracle.read_bytes())==obs['OracleSHA256'].lower(),'Native size oracle source changed')
        self.add_payload(f'{label}-SizeReportingNative-oracle.py',oracle.read_bytes())
        for fixture in manifest['fixtures']:
            original=self.owned_input(Path(obs['OriginalCorpusDirectory'])/fixture['file'])
            require(sha(original.read_bytes())==fixture['sha256'] and original.stat().st_size==fixture['bytes'],'Recreated original fixture bytes differ')
            self.support_file(original,fixture['sha256'])
        rows=obs['Observations']
        labels={f'actual-{name}-{preset}' for name in ('small-print','scan','mixed') for preset in ('screen','ebook')}
        labels.update(('actual-tiny-larger-no-benefit-code0','controlled-equal-after-actual-GS0-code0','actual-master-only-skip-code0',
                       'actual-master-only-missing-code0','actual-master-only-failure-code2'))
        require(len(rows)==11 and {r['Label'] for r in rows}==labels,'Exact eleven native size observations required')
        for row in rows:
            proof=row['Proof'];name=row['Label'];control=row['Control'];expectation=row['FixtureExpectation'];state=proof['ExpectedState']
            require(row['Before']==row['After'],'Native source/original/foreign preservation differs');self.size_snapshots(row['After'])
            app=Path(row['SourceFolder']).parent/'app';cap=app.parent/'captured-calls'
            require(sha((app/'WinPDFMerge.ps1').read_bytes())==row['EntrySHA256'].lower()==sha((self.repo/'WinPDFMerge.ps1').read_bytes()),'Actual copied size entry differs')
            self.support_file(app/'WinPDFMerge.ps1',row['EntrySHA256']);self.support_file(app/'src/WinPDFMerge.Helpers.ps1',row['CopiedHelperSHA256'])
            command_path=self.owned_input(proof['CommandReceiptPath']);_,command=load_json(command_path);self.support_file(command_path)
            arguments=['-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(app/'WinPDFMerge.ps1'),row['SourceFolder'],'-OutputFolder',row['OutputFolder']]
            preset='ebook' if name.endswith('-ebook') else 'screen'
            if preset=='ebook':arguments+=['-EmailPreset','ebook']
            if control=='skip':arguments+=['-SkipEmail']
            require(command['Executable']==metadata['driver'] and command['Arguments']==arguments and proof['ExpectedPreset']==preset,'Actual size entry arguments/preset differ')
            expected_code=2 if control=='corrupt' else 0
            require(command['ExitCode']==proof['Result']['ExitCode']==expected_code,'Actual native size exit code differs')
            for stream in ('Stdout','Stderr'):
                source=command_path.parent/('entry.'+stream.lower()+'.txt');self.support_file(source,command[stream+'SHA256'])
                require(source.read_bytes().decode('utf-8-sig')==proof['Result'][stream],'Actual entry raw stream differs')
            master=proof['MasterJob'];self.native_result(master['Job'],metadata['pdftk']);self.complete_native(master['Job']['ValidationResult']['NativeResult'])
            master_path=self.owned_input(proof['MasterPath']);master_bytes=master_path.stat().st_size
            require(master_bytes==proof['MasterBytes'] and master['ExpectedPageCount']==expectation['page_count'],'Published master bytes/count differ')
            require(str(master_path.parent)==row['OutputFolder'] and not Path(master['StageDirectory']).exists()
                and not list(Path(row['OutputFolder']).glob('.WinPDFMerge*')),'Native output directory/owned cleanup differs')
            reads=proof['FinalReads'];require(len(reads)==(2 if state in ('skipped','unavailable','failed') else 4 if state=='published' else 3),'Actual inspected source/final/candidate list differs')
            for read in reads:
                self.snapshot_rows([read['Snapshot']]);expected_pages=[{'identifier':i,'rotation_degrees':0,'size_points':expectation['page_size_points']} for i in expectation['page_identifiers']]
                require(read['PdfTkRead']['ExitCode']==read['OracleRead']['ExitCode']==0 and read['Oracle']['pages']==expected_pages
                    and read['Oracle']['page_count']==expectation['page_count'],'Independent source/master/candidate count/IDs/geometry differ')
            for call in proof['NativeCalls']:self.complete_native(call['Result']) if call['Result']['Succeeded'] else require(control=='corrupt' and call['Result']['ExitCode']==1 and call['Result']['OwnershipReleased'],'Unexpected native failure or unreleased ownership')
            email=proof['EmailJob'];candidate=None;gs_calls=[c for c in proof['NativeCalls'] if c['Executable']==metadata['ghostscript']]
            if state in ('skipped','unavailable'):
                require(not email and not proof['EmailPaths'] and not gs_calls and not list(cap.glob('unexpected-GS-*')),'Master-only mode discovered/probed/launched GS')
                require((state=='skipped')==(control=='skip'),'Master-only control label differs')
                if control=='skip':require(str(Path(metadata['ghostscript']).parent) in command['ChildPath'],'Skip sentinel must leave real GS available')
                else:require(str(Path(metadata['ghostscript']).parent) not in command['ChildPath'],'Scoped missing optional GS PATH differs')
            else:
                require(email['MasterBefore']==email['MasterAfter'] and email['RequestedEmailPreset'].lower()==preset,'Original master/preset changed')
                self.snapshot_rows([email['MasterAfter']]);require(email['MasterAfter']['Length']==master_bytes,'Preserved original master size differs')
                conversions=[c for c in gs_calls if '-sDEVICE=pdfwrite' in c['Arguments']];require(len(conversions)==1,'One actual GS conversion required')
                actual_input=str(app.parent/'owned-corrupt-gs-input.pdf') if control=='corrupt' else str(master_path)
                vector=['-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dCompatibilityLevel=1.6',
                        '-dPDFSETTINGS=/'+preset,'-dDetectDuplicateImages=true','-o',email['StagedEmailPath'],'-f',actual_input]
                require(conversions[0]['Arguments']==vector and conversions[0]['RemoveEnvironmentVariables']==['GS_OPTIONS']
                    and email['ActualInputPaths']==[actual_input],'Actual GS fixed vector/safety/input differs')
                require(not Path(email['StageDirectory']).exists() and not Path(email['StagedEmailPath']).exists(),'Native email staging residue remains')
                result=email['Job']
                if state=='failed':
                    require(control=='corrupt' and result['NativeResult']['ExitCode']==1 and not result['OutputValidated'] and not result['OutputPublished'] and not proof['EmailPaths'],'Actual engine failure was advertised as validated/published')
                    partial=email['FailedPartialBeforeCleanup'];retained=cap/'retained-failed-partial.pdf'
                    require(partial['Length']>0 and sha(retained.read_bytes())==partial['SHA256'].lower() and retained.stat().st_size==partial['Length'],'Actual partial-before-cleanup binding differs')
                    require(result['MasterBytes']==Path(actual_input).stat().st_size and 'substituted corrupt input length' in row['ControlledHooks'],'Corrupt input versus original master metrics not disclosed')
                    self.support_file(retained,partial['SHA256'])
                else:
                    self.complete_native(result['NativeResult']);self.validated_email(result,metadata['pdftk']);self.complete_native(result['ValidationResult']['NativeResult'])
                    require(result['Succeeded'] and result['MasterBytes']==master_bytes,'Validated email master metrics/result differ')
                    retained=email['RetainedValidatedCandidate'];self.snapshot_rows([retained]);candidate=retained['Length']
                    require(candidate==result['OutputBytes']==proof['ValidatedCandidateBytes'],'Retained validated candidate metrics differ')
                    self.support_file(retained['Path'],retained['SHA256'])
                    require((state=='published')==(candidate<master_bytes) and (bool(proof['EmailPaths']))==(state=='published'),'Strict size-benefit/final-path decision differs')
                    if proof['EmailPaths']:require(len(proof['EmailPaths'])==1 and sha(Path(proof['EmailPaths'][0]).read_bytes())==retained['SHA256'].lower(),'Published email differs from retained validated bytes')
                    if control=='equal':
                        boundary=proof['EqualBoundary'];require(candidate==master_bytes and boundary['InjectedCandidate']['SHA256']==boundary['UnmodifiedMaster']['SHA256']
                            and 'not a genuine GS equal-size result' in boundary['Scope'],'Controlled equality disclosure/boundary differs')
                        actual=boundary['ActualGhostscriptCandidate'];self.support_file(boundary['RetainedActualGhostscriptCandidate'],actual['SHA256'])
                    if name=='actual-tiny-larger-no-benefit-code0':require(candidate>master_bytes,'Genuine larger case is not larger')
            expected_lines=self.size_lines(master_bytes,candidate,state=='published');require(proof['ExpectedSizeLines']==expected_lines,'Independent byte/unit/reduction display differs')
            log_path=self.owned_input(proof['LogPath']);log=log_path.read_bytes().decode('utf-8-sig');require(log==proof['Log'],'Actual size log changed');self.support_file(log_path)
            summary=re.search(r'(?m)^(?:SUCCESS|PARTIAL SUCCESS):',proof['Result']['Stdout']);require(summary is not None,'Actual final console summary missing')
            for output in (log,proof['Result']['Stdout'][summary.start():]):
                actual_lines=[line for line in output.splitlines() if re.match(r'^(Master size:|Email size:|Email reduction:|Validated email candidate size:|Email candidate reduction:)',line)]
                require(actual_lines==expected_lines,'Actual log/final console metrics differ from files')
            for line in proof['Result']['Stdout'].splitlines():
                if re.match(r'^(Master size:|Email size:|Email reduction:|Validated email candidate size:|Email candidate reduction:)',line):require(line in expected_lines,'Earlier echoed metric is untruthful')
            require('Email result: '+state in log and (' - Email-optimized:' in proof['Result']['Stdout'])==(state=='published'),'Actual result/final email list differs')
            for item in cap.glob('*.json'):self.support_file(item)
        name=f'{label}-SizeReportingNative-observations.json';self.add_payload(name,raw)
        self.record_for(label,'SizeReportingNative').update(observations_file=name,raw_observations_sha256=sha(raw),observations_sha256=sha(self.payloads[name]),observation_count=11,
            scope='Actual entry/selected engines and independent structural reads; genuine smaller/larger, controlled equal/corrupt/skip cases disclosed; manual rendered-image acceptance separate')
        self.size_observations[label]={'path':path,'sha256':sha(raw),'receipt':obs}

    def size_unit(self,label,job,metadata):
        stdout=self.owned_input(job['log']).read_bytes().decode('utf-8-sig')
        matches=re.findall(r'(?m)^Size reporting unit receipts: (.+?)\r?$',stdout);require(len(matches)==1,'One unit size receipt required')
        raw,rows=load_json(self.owned_input(matches[0]));require(len(rows)==32 and len({r['Label'] for r in rows})==32,'Exact32 distinct unit size cases required')
        copied=[r for r in rows if 'ReceiptPath' in r]
        require(len(copied)==7 and {r['Label'] for r in copied}=={'entry-published','entry-equal','entry-larger','entry-skipped','entry-unavailable','entry-failed','entry-log-failed'},
                'Exactly seven labelled controlled copied-entry cases required')
        for index,row in enumerate(copied,1):
            require(row['Before']==row['After'] and row['PersistentEnvironmentChanges'] is False and row['Command'][0]==metadata['driver']
                and row['Culture']=='de-DE','Controlled unit source/environment/selected shell differs')
            result=row['Result'];require(result['Started'] and result['OwnershipReleased'] and result['ProcessId']>0
                and not result['TimedOut'] and not result['Cancelled'] and not result['LaunchError'] and not result['CaptureError']
                and not result['TerminationError'] and not result['StdoutTruncated'] and not result['StderrTruncated'],
                'Controlled actual shell lifecycle/capture incomplete')
            receipt_path=self.owned_input(row['ReceiptPath']);root=receipt_path.parent;receipt_raw,receipt=load_json(receipt_path)
            require(sha(receipt_raw)==row['ReceiptSHA256'] and receipt==row['Receipt'],'Controlled size decision receipt changed')
            for relative,field in [('app/WinPDFMerge.ps1','EntrySHA256'),('app/src/WinPDFMerge.Helpers.ps1','HelperWithHooksSHA256'),
                                  ('Invoke-Entry.ps1','WrapperSHA256'),('app/size-config.json','ConfigurationSHA256'),('stdout.txt','StdoutSHA256'),('stderr.txt','StderrSHA256')]:self.support_file(root/relative,row[field])
            _,config=load_json(root/'app/size-config.json')
            require(':'.join(sha(self.owned_input(p).read_bytes()) for p in (Path(config['Source'])/'1.pdf',Path(config['Output'])/'foreign-existing.pdf'))==row['After'],'Controlled source/foreign raw hashes changed')
            expected_code=2 if row['Label'] in ('entry-failed','entry-log-failed') else 0
            require(result['ExitCode']==receipt['Outcome']['ExitCode']==expected_code,'Controlled outcome code differs')
            finals=row['Finals'];require(finals and len(finals)==(2 if receipt['Outcome']['EmailState']=='published' else 1),'Controlled published final list differs')
            for final in finals:self.support_file(final['Path'],final['SHA256']);require(Path(final['Path']).stat().st_size==final['Bytes'],'Controlled final length changed')
            for report in receipt['Reports']:
                r=report['Report'];candidate=r['EmailBytes'] if report['EmailBytesBound'] else None
                require(r['Lines']==self.size_lines(report['MasterBytes'],candidate,report['EmailPublished']),'Controlled unit report arithmetic differs')
            name=f'{label}-SizeReporting-entry-{index:02d}-receipt.json';self.add_payload(name,receipt_raw)
            self.support_file(receipt_path,row['ReceiptSHA256']);self.support_file(root/'invocation.json')
        name=f'{label}-SizeReporting-observations.json';self.add_payload(name,raw)
        self.record_for(label,'SizeReporting').update(observations_file=name,raw_observations_sha256=sha(raw),observations_sha256=sha(self.payloads[name]),observation_count=32,
            scope='25 numeric import-only cases and7controlled copied-entry/native receipts; actual bounded shell hosts, no PDF-engine/manual support claim')

    def historical(self):
        for index_name in ('T17-unit-dirty-history.json','T17-native-dirty-history.json'):
            path=self.work/index_name;_,index=load_json(path);require(index['Task']=='T17','History ownership differs');self.archive(path)
            for attempt in index['Attempts']:
                for item in attempt['Files']:
                    original=self.repo/item['Path'];require(sha(original.read_bytes())==item['SHA256'],'Indexed historical bytes changed');self.archive(original)
                copies=attempt.get('RetainedCopiedEntryCases',attempt.get('RetainedCopiedApplicationCases',[]))
                for case in copies:
                    for item in case['Files']:self.support_file(self.repo/item['Path'],item['SHA256'])
                    for item in case.get('RetainedSyntheticPdfHashes',[]):self.support_file(self.repo/item['Path'],item['SHA256'])
        def retain_bound_values(value):
            if isinstance(value,dict):
                if 'Path' in value and 'SHA256' in value:
                    original=(self.repo/value['Path']).resolve()
                    require(sha(original.read_bytes())==value['SHA256'].lower(),'Historical retained source/binary bytes differ: '+original.name)
                    if original.is_relative_to(self.work):
                        if original.suffix.lower() in ('.pdf','.png'):self.support_file(original,value['SHA256'])
                        else:self.archive(original,'historical source/capture/provenance binding; not clean acceptance counts')
                    else:
                        require(original==(self.repo/'tests/fixtures/presets/generate_presets.py').resolve(),'Only known tracked original fixture recipe may bind outside ignored history')
                        self.add_payload('original-corpus-generator.py',original.read_bytes())
                        self.records.append({'classification':'historical original authoring recipe; exact bytes also tracked at clean C1, not acceptance counts',
                            'source_relative_path':'tests/fixtures/presets/generate_presets.py','file':'original-corpus-generator.py',
                            'raw_sha256':value['SHA256'].lower(),'sha256':sha(self.payloads['original-corpus-generator.py'])})
                for item in value.values():retain_bound_values(item)
            elif isinstance(value,list):
                for item in value:retain_bound_values(item)
        for leaf in ('T17-unit-source-retention.json','T17-original-corpus-preparation.json','T17-dirty-render-inputs.json','T17-dirty-visual-review.json','T17-root-preparation-history.json'):
            original=self.work/leaf;_,value=load_json(original);self.archive(original,'historical preparation/retention disclosure; not clean visual/native acceptance')
            retain_bound_values(value)
        for leaf in ('Collect-T17NativeFocus.py','Bind-T17CorpusPreparation.py','Index-T17UnitHistory.py','Run-T17SizeReporting.py','Run-T17SizeReporting-before-snapshots.py'):
            self.archive(self.work/leaf,'historical focused history/retention producer; not clean acceptance counts')
        for render_root in sorted(self.work.glob('T17-dirty-visual-renders-*')):
            _,render=load_json(render_root/'receipt.json');require(render['dirty_worktree'] and render['phase']=='dirty','Dirty visual context differs')
            self.archive(render_root/'receipt.json','historical dirty render receipts; preliminary visual scope, not clean AC041')
            for row in render['renders']:
                for key in ('source_pdf','png'):self.support_file(row[key],row[key+'_sha256'])
                for key in ('stdout','stderr'):
                    original=self.owned_input(row[key]);require(sha(original.read_bytes())==row[key+'_sha256'],'Dirty render raw capture changed');self.archive(original,'historical dirty renderer capture; not clean visual acceptance')
            for child in render_root.iterdir():
                if child.is_file() and child.suffix.lower() not in ('.pdf','.png'):self.archive(child,'historical dirty renderer source/capture; not clean visual acceptance')
        for leaf in ('Bind-T17VisualReview.py','Render-T17Presets.py'):
            self.archive(self.work/leaf,'root render/review preparation source; actual image inspection is separately bound')
        for path in sorted(self.work.glob('T17-dirty-*.execution.json')):
            _,execution=load_json(path);self.archive(path)
            for key,digest in [('stdout','stdout_sha256'),('stderr','stderr_sha256')]:
                original=self.owned_input(execution[key]);require(sha(original.read_bytes())==execution[digest],'Dirty root capture changed');self.archive(original)
            text=Path(execution['stdout']).read_bytes().decode('utf-8-sig')
            matches=re.findall(r'(?m)^Reports: (.+?)\r?$',text)
            if len(matches)==1:
                report=self.owned_input(matches[0]);self.archive(report/'summary.json');self.archive(report/'results.xml')
            else:self.records.append({'classification':'historical missing report explicitly absent; never fabricated','source_relative_path':path.relative_to(self.work).as_posix()})
        for path in sorted(self.work.glob('T17-precommit-analyzer-*')):
            if path.is_file():self.archive(path,'historical dirty static findings; not acceptance passes')
            else:
                for child in sorted(path.iterdir()):
                    if child.is_file():self.archive(child,'historical dirty analyzer execution/capture; not acceptance passes')
        for leaf in ('Invoke-T17NativeSmoke.ps1','Collect-T17NativeFocus.py','Bind-T17CorpusPreparation.py'):
            self.archive(self.work/leaf,'historical dirty focused authoring source; not acceptance counts')
        for path in sorted(self.work.glob('T17*review*dirty*.json')):
            _,review=load_json(path);self.archive(path,'historical dirty read-only source review; separate from clean acceptance counts')
            support=review.get('ReviewSupport',{})
            items=([support['Producer']] if 'Producer' in support else [])+support.get('ReadOnlyDiffExecution',[])+support.get('SourceAndDiffCaptures',[])
            for item in items:
                original=self.repo/item['Path'];require(sha(original.read_bytes())==item['SHA256'],'Dirty review support changed')
                self.archive(original,'historical dirty read-only review support; not acceptance counts')
        for leaf in ('T17-precommit-review.json','T17-precommit-source-review-support-index.json'):
            _,review=load_json(self.work/leaf);self.archive(self.work/leaf,'historical scoped precommit source/static review; not clean counts')
            if leaf.endswith('support-index.json'):retain_bound_values(review)
        self.archive(self.work/'Commit-T17C1.py','corrected root checkpoint producer; initial failed guard source/captures absent, not reconstructed')
        self.archive(self.work/'Run-T17Closure.py','corrected root evidence command wrapper; actual argument-preparation failure excluded from acceptance')
        self.archive(self.work/'T17-prefix-capture-b94699af067347afa08719c152b76e9d/wrapper-source.py',
                     'original wrapper retained by earlier successful prefix capture; not relabelled as failed-command snapshot')
        analyzer_source=self.work/'Run-T17Analyzer.py'
        if analyzer_source.exists():self.archive(analyzer_source,'historical dirty scoped analyzer source; not acceptance counts')
        for root in sorted(self.work.glob('T17-collector-check-*')):
            if not (root/'execution.json').exists():continue
            _,attempt=load_json(root/'execution.json')
            if attempt['ExitCode']==0:continue
            require(attempt['ApplicationOrNativeTestsExecuted'] is False and not attempt['PublicWriteRequested']
                    and attempt['CleanReports'] is None and attempt['TotalPassed'] is None,'Failed collector preparation misclassified')
            require(sha((root/'collector-source.py').read_bytes())==attempt['CollectorSourceSHA256']
                    and sha((root/'stdout.txt').read_bytes())==attempt['StdoutSHA256']
                    and sha((root/'stderr.txt').read_bytes())==attempt['StderrSHA256'],'Collector failed preparation binding changed')
            for leaf,digest in attempt['SourceSHA256'].items():require(sha((root/leaf).read_bytes())==digest,'Failed collector exact source/primitives/wrapper snapshot changed')
            for original in sorted(root.iterdir()):
                if original.is_file():self.archive(original,'historical collector preparation only; no application failure or acceptance counts')
            self.preparation_history.append({'source_relative_root':root.relative_to(self.work).as_posix(),'exit_code':attempt['ExitCode'],
                'execution_receipt_raw_sha256':sha((root/'execution.json').read_bytes()),'collector_source_sha256':attempt['CollectorSourceSHA256'],
                'stdout_raw_sha256':attempt['StdoutSHA256'],'stderr_raw_sha256':attempt['StderrSHA256'],'acceptance_counts_available':False,
                'recorded_stderr_final_line':self.sanitize_string((root/'stderr.txt').read_text(encoding='utf-8-sig').rstrip().splitlines()[-1]),
                'diagnosis':'Collector-only preparation failure; exact original stderr/source retained. First failure expected entry-smaller rather than actual entry-published; second rejected original tracked generator binding as though every history source were ignored. Semantic readers corrected; no application/source/test acceptance change.'})

    def environment_and_pins(self,path):
        Legacy.environment_and_pins(self,path);raw,receipt=load_json(self.owned_input(path));self.standalone['T17-environment.json']=self.payloads['ordinary-ps51-environment.json']
        require(receipt['commit_under_inventory']==self.commit and receipt['dirty_worktree'] is False and receipt['task']=='T17'
            and receipt['application_or_native_test_claim'] is False,'Fresh ordinary inventory context differs')
        for leaf,digest in [('T17-InventoryCommand.ps1','inventory_source_sha256'),('T17-environment.stdout.txt','raw_stdout_sha256'),('T17-environment.stderr.txt','raw_stderr_sha256')]:
            source=self.work/leaf;require(sha(source.read_bytes())==receipt[digest],'Fresh inventory source/capture changed')
            self.archive(source,'separate read-only ordinary inventory; no application acceptance pass')

    def verified_cache_receipt(self,path):
        raw,receipt=load_json(self.owned_input(path));require(receipt['task']=='T17' and not receipt['acquisition_performed'],'Cache scope changed')
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
        self.add_payload('approved-cache-verification.json',raw);self.standalone['T17-cache-verification.json']=self.payloads['approved-cache-verification.json']

    def analyzer(self,label,path,version):
        raw,report=load_json(self.owned_input(path))
        require(report['Task']=='T17' and report['Phase']=='C1' and report['CommitUnderTest']==self.commit and report['DirtyWorktree'] is False
                and report['ShellVersion']==version and report['AnalyzerVersion']=='1.25.0','Clean scoped analyzer context differs')
        for severity,key in [(2,'Errors'),(1,'Warnings'),(0,'Information')]:
            require(report[key]==sum(f['Severity']==severity for f in report['Findings']),'Scoped analyzer counts differ from findings')
        require((report['Errors'],report['Warnings'],report['Information'])==(0,51,14),'Frozen scoped analyzer findings differ')
        name=f'{label}-PSScriptAnalyzer-findings.json';self.add_payload(name,raw)
        self.records.append({'classification':'clean C1 scoped five-file static findings; warnings/info retained, not test passes',
            'shell':label,'file':name,'raw_sha256':sha(raw),'sha256':sha(self.payloads[name]),'analyzer_version':'1.25.0',
            'error_count':0,'warning_count':51,'information_count':14})
        executions=list(self.work.glob('T17-C1-analyzer-execution-'+label+'-*'));require(len(executions)==1,'One clean scoped analyzer execution required')
        _,execution=load_json(executions[0]/'execution.json')
        require(execution['started'] and execution['exit_code']==0 and not execution['timed_out'] and not execution['error']
            and execution['analyzer_report_sha256']==sha(raw) and execution['commit_after']==self.commit and not execution['git_status_after']
            and execution['source_bytes_unchanged'],'Scoped analyzer lifecycle/source binding differs')
        require(len(execution['source_bindings_after'])==5,'Exactly five changed PowerShell source bindings required')
        for source,digest in execution['source_bindings_after'].items():require(sha((self.repo/source).read_bytes())==digest,'Scoped analyzer source bytes changed')
        for item in executions[0].iterdir():
            if item.is_file():self.archive(item,'clean C1 scoped analyzer execution/capture; separate from application case counts')

    def manual_visual_review(self,path):
        raw,review=load_json(self.owned_input(path))
        require(review['Task']=='T17' and review['Phase']=='C1' and review['CommitUnderTest']==self.commit and review['DirtyWorktree'] is False
            and review['Result']=='pass' and review['ManualVisualInspectionPerformed'] is True and review['AcceptanceCases']==['AC041']
            and review['Observer']=='root Codex actual visual inspection','Explicit clean Codex visual review is required; render success alone is not acceptance')
        render_path=self.owned_input(review['RendererReceiptPath']);render_raw,renderer=load_json(render_path)
        require(sha(render_raw)==review['RendererReceiptSHA256'] and renderer['task']=='T17' and renderer['phase']=='C1'
            and renderer['commit_under_test']==self.commit and renderer['dirty_worktree'] is False and renderer['result']=='rendered_pending_visual_review',
            'Bound renderer clean context differs')
        source=self.owned_input(review['RendererSourcePath']);require(sha(source.read_bytes())==review['RendererSourceSHA256']==renderer['source_sha256'],'Actual rendering source changed')
        require(sha(Path(renderer['renderer']).read_bytes())==renderer['renderer_sha256'],'Actual Poppler executable changed')
        require(len(renderer['observations'])==2 and {r['shell'] for r in renderer['observations']}=={'ps51','ps7'},'Both clean shell native receipt inputs required')
        expected={};labels={f'actual-{kind}-{preset}' for kind in ('small-print','scan','mixed') for preset in ('screen','ebook')}
        for selected in renderer['observations']:
            native=self.size_observations[selected['shell']]
            require(Path(selected['path'])==native['path'] and selected['sha256']==native['sha256']
                and selected['observed_test_source_sha256']==native['receipt']['TestSourceSHA256'],'Rendering input native receipt/source binding differs')
            seen=set()
            for row in native['receipt']['Observations']:
                if row['Label'] not in labels:continue
                for role,pdf in (('original',row['OriginalFixturePath']),('master',row['Proof']['MasterPath']),
                                 ('candidate',row['Proof']['EmailJob']['RetainedValidatedCandidate']['Path'])):
                    if (role,pdf) in seen:continue
                    seen.add((role,pdf))
                    for page in range(1,row['FixtureExpectation']['page_count']+1):
                        expected[(selected['shell'],row['Label'],role,page)]=pdf
        require(len(expected)==40 and len(renderer['renders'])==review['RenderedPageCount']==len(review['Coverage'])==40,'Complete forty original/master/derivative render pages required')
        rendered={};coverage={};unique=set()
        for row in renderer['renders']:
            key=(row['shell'],row['label'],row['role'],row['page']);require(key in expected and key not in rendered and expected[key]==row['source_pdf'],'Unexpected/missing/duplicate rendered corpus page')
            require(row['exit_code']==0,'Actual renderer failure recorded')
            png=self.owned_input(row['png']);pdf=self.owned_input(row['source_pdf'])
            require(sha(png.read_bytes())==row['png_sha256'] and png.stat().st_size==row['png_bytes']
                and sha(pdf.read_bytes())==row['source_pdf_sha256'] and pdf.stat().st_size==row['source_pdf_bytes'],'Reviewed PNG/PDF raw bytes changed')
            vector=[renderer['renderer'],'-png','-r','144','-f',str(row['page']),'-singlefile',row['source_pdf'],str(png.with_suffix(''))]
            require(row['command']==vector,'Actual equal-scale renderer command differs')
            for stream in ('stdout','stderr'):
                original=self.owned_input(row[stream]);require(sha(original.read_bytes())==row[stream+'_sha256'],'Renderer raw capture changed')
                self.archive(original,'clean C1 read-only renderer capture; explicit visual review is separately bound')
            self.support_file(png,row['png_sha256']);self.support_file(pdf,row['source_pdf_sha256'])
            rendered[key]=row;unique.add(row['png_sha256'])
        reviewed={}
        for image in review['ReviewedUniqueImages']:
            png=self.owned_input(image['Path']);require(sha(png.read_bytes())==image['SHA256'] and png.stat().st_size==image['Bytes']
                and image['SHA256'] not in reviewed and 'Actual root Codex full-page view_image(detail=original)' in image['Inspection'],
                'Every distinct PNG requires a bound explicit actual full-page visual inspection')
            reviewed[image['SHA256']]=image['Path']
        require(set(reviewed)==unique and len(reviewed)==review['ReviewedUniqueImageCount'],'Actual unique image views do not cover every rendered page')
        for row in review['Coverage']:
            key=(row['Shell'],row['Label'],row['Role'],row['Page']);require(key in rendered and key not in coverage,'Visual coverage key differs')
            render=rendered[key]
            require(row['PDFPath']==render['source_pdf'] and row['PDFSHA256']==render['source_pdf_sha256']
                and row['PNGPath']==render['png'] and row['PNGSHA256']==render['png_sha256']
                and row['ReviewedViaIdenticalPNG']==reviewed[row['PNGSHA256']],'Visual dedup coverage does not bind identical raw image pixels')
            coverage[key]=row
        require(set(coverage)==set(expected) and {r['Document'] for r in review['Observations']}=={'small-print','scan','mixed'}
            and all(r['Observation'].strip() for r in review['Observations']),'Explicit comparative qualitative corpus observations missing')
        for leaf,field in [('renderer-version.stdout.txt','renderer_version_stdout_sha256'),('renderer-version.stderr.txt','renderer_version_stderr_sha256')]:
            version_stream=render_path.parent/leaf;require(sha(version_stream.read_bytes())==renderer[field],'Actual renderer version capture changed');self.archive(version_stream,'clean C1 renderer version capture; no application counts')
        self.add_payload('manual-visual-review.json',raw);self.add_payload('manual-render-receipt.json',render_raw)
        self.archive(source,'clean C1 renderer source; no pixel review is inferred from rendering')
        self.archive(self.work/'Bind-T17VisualReview.py','clean explicit manual visual review binder; does not inspect pixels automatically')
        for pattern in ('T17-C1-render-capture-*','T17-C1-visual-bind-capture-*'):
            captures=list(self.work.glob(pattern));require(len(captures)==1,'One actual clean render/bind command capture required')
            capture=captures[0];_,execution=load_json(capture/'execution.json')
            require(execution['Task']=='T17' and execution['ExitCode']==0
                and execution['ProducerSHA256']==sha((capture/'producer-source.py').read_bytes())
                and execution['StdoutSHA256']==sha((capture/'stdout.txt').read_bytes())
                and execution['StderrSHA256']==sha((capture/'stderr.txt').read_bytes()),'Actual render/bind command lifecycle/source/captures differ')
            for child in capture.iterdir():
                if child.is_file():self.archive(child,'clean C1 actual read-only render/visual-binding command; no application test rerun')
        self.manual_review_binding={'file':'docs/codex/evidence/T17-C1-reports/manual-visual-review.json',
            'raw_sha256':sha(raw),'sha256':sha(self.payloads['manual-visual-review.json']),
            'renderer_receipt_file':'docs/codex/evidence/T17-C1-reports/manual-render-receipt.json',
            'renderer_receipt_raw_sha256':sha(render_raw),'renderer_receipt_sha256':sha(self.payloads['manual-render-receipt.json']),
            'observer':review['Observer'],'rendered_pages':40,'explicit_unique_images_viewed':len(reviewed),
            'scope':'Clean C1 actual Codex full-page comparisons with explicit identical-pixel coverage; no owner/physical Explorer desktop or universal fidelity claim'}

    def results_document(self,shells,destination):
        return {'schema_version':1,'task':'T17','checkpoint':'C1','commit_under_test':self.commit,'dirty_worktree':False,
                'implementation_acceptance':'pass','ac040':'pass','ac041':'pass','cases_per_shell':COUNTS,'selected_tiers_in_execution_order':list(TIERS),
                'passed_per_shell':{s['shell']:s['passed'] for s in shells},'total_passed':1158,'clean_reports':32,'all_failures_skips_not_run':0,
                'scope':'AC040 actual Windows PS5.1/PS7.6.6 entry size lines independently agree with retained real master, validated candidate and published derivative bytes. Real smaller scan/mixed derivatives and larger vector/tiny candidates, controlled equal boundary and master-only optional/failure states are separately labelled. AC041 is explicit clean-C1 Codex visual review of bound rendered original/master/screen/ebook pages, separate from native structure/count assertions and owner/Explorer interaction.',
                'reports_manifest':'docs/codex/evidence/T17-C1-reports/manifest.json','selected_cache_verification':'docs/codex/evidence/T17-cache-verification.json',
                'ordinary_environment_receipt':'docs/codex/evidence/T17-environment.json','pester_version':'6.2.0','pdftk_version':'2.02','ghostscript_version':'10.08.0','development_oracle':ORACLE,
                'collector_command_template':'<selected shell> -NoProfile -ExecutionPolicy RemoteSigned -File tests/.work/Run-T17Checkpoint.ps1 -ShellLabel ps51|ps7 -ExpectedCommit '+self.commit+' -ExpectedCountsPath tests/.work/T17-expected-counts.json -Checkpoint C1',
                'evidence_collector_command':'<approved Python> -B tests/.work/Collect-T17Evidence.py --repo . --commit '+self.commit+' [--write]',
                'historical_rule':'All failed/corrected dirty focused attempts and static observations remain separate, never added to clean1158. Original records and copied support hash inventories retained; absent outputs never fabricated.',
                'historical_findings':{'SizeReporting':'Initial32-case unit focuses26passed/6failed each assumed only one stdout metric emission; existing logger echoes before final summary. Corrected test assertions32passes each. Initial full failing test source absent/not reconstructed; original pre-run hashes and raw failures retained. Final full source captured after runs matches invocation hashes.',
                    'SizeReportingNative':'Initial11-case native focuses6passed/5failed each: tiny-case expectation object shadowed by case-insensitive Fixture parameter. Actual six corpus cases passed and tiny jobs executed; corrected test-only variable11passes each. Exact full sources snapshotted before every native attempt.'},
                'collector_preparation_history':self.preparation_history,
                'checkpoint_preparation_disclosure':{'classification':'root orchestration preparation only; no application/test/manual counts',
                    'parent_reported_failure_count':2,'receipt':'Historical T17-root-preparation-history.json is archived with explicit original stream/source-retention limits.',
                    'diagnoses':['First Commit-T17C1 invocation raised KeyError Result before tracked mutation; corrected reader then committed C1.',
                                 'Initial Run-T17Closure wrapper rejected --phase before launching renderer; corrected forwarding then executed actual rendering.'],
                    'capture_limit':'Failed raw outputs are actual session-transcript only, not reconstructed. Initial commit producer source absent; original closure wrapper retained from earlier successful prefix command, not relabelled as failed-command snapshot.'},
                'environment':{'inventory':self.inventory,'standard_user_non_elevated':True,'test_process_policy':'Previously authorized process-only RemoteSigned',
                    'persistent_policy_security_parent_environment_changed':False,'acquisition_performed':False},
                'static_analysis':{'scope':'Five changed PowerShell files; separate from application cases and full T22 lint gate',
                    'analyzer_version':'1.25.0','each_shell':{'errors':0,'warnings':51,'information':14}},
                'limitations':['Actual cmd/BAT selects Windows PowerShell5.1 and is distinct from the open physical Explorer/manual gate.',
                    'Manual observations apply only to this original synthetic small-print, scan and mixed corpus at the recorded render/view scales. Page count/ID/geometry and smaller bytes alone are not fidelity evidence; no universal readability, target attachment size, signature, PDF/A or security promise.',
                    'Copied helper recording and SkipEmail throwing sentinels are controlled and disclosed. Earlier affected fault cases retain controlled corrupt-input, owned padding, logger/token/native-tree seams.',
                    'Only16affected tiers selected; unchanged NativeRunner, DependencyEntry, SourceDiscovery, PdftkPaths and NativeFixture tiers retain prior accepted evidence rather than a new run claim.',
                    'No Windows support-channel, UNC/Explorer/CI/package or published-release acceptance follows; downstream gates remain open.'],
                'manual_visual_review':self.manual_review_binding,
                'checkpoint_note':'This certifies clean implementation/explicit visual gates; root separately maintains task/status and final records synchronization.'}

    def finish(self,shells,destination,write):
        require(sum(s['passed'] for s in shells)==1158,'Exact1158clean total required')
        evidence=self.repo/'docs/codex/evidence';destination=destination.resolve();require(destination==evidence/'T17-C1-reports','Exact root-authorized destination required')
        self.add_payload('retained-support-bindings.json',json_bytes(self.support))
        result_path=evidence/'T17-C1-results.json';result_payload=json_bytes(self.sanitize_value(self.results_document(shells,destination)))
        manifest={'schema_version':1,'task':'T17','checkpoint':'C1','commit_under_test':self.commit,'dirty_worktree':False,'clean_reports':32,
                  'total_clean_passed':1158,'shells':shells,'records':self.records,'payload_bindings':self.payload_bindings.copy(),
                  'evidence_collector_sha256':sha(Path(__file__).read_bytes()),'prior_collector_primitives_sha256':sha(PRIOR_PATH.read_bytes()),
                  'native_schema_primitives_sha256':sha(P['BASE_PATH'].read_bytes()),'legacy_primitives_sha256':sha(P['B']['LEGACY_PATH'].read_bytes()),
                  'xml_redactions':['environment.'+k for k in LEGACY['XML_IDENTITY']],
                  'privacy_scope':'XML machine/user/domain/cwd and repository/user/cache/temp path prefixes sanitized; synthetic IDs, whitespace, vectors, diagnostics and native results retained.',
                  'exact_byte_policy':'Raw and public sanitized/canonical SHA bindings are distinct. Supporting copied raw files retained ignored and verified in support inventory; only original synthetic document paths permitted.',
                  'results_file':'docs/codex/evidence/T17-C1-results.json','results_sha256':sha(result_payload),
                  'standalone_receipts':[{'file':'docs/codex/evidence/'+n,'sha256':sha(p)} for n,p in self.standalone.items()],
                  'frozen_implementation_source_bytes':[{'path':p,'sha256':sha((self.repo/p).read_bytes())} for p in
                     ['WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1','tests/pdf/SizeReporting.Tests.ps1','tests/pdf/SizeReporting.Native.Tests.ps1','tests/fixtures/presets/generate_presets.py','tests/fixtures/presets/manifest.json','README.md','docs/EMAIL_PRESETS.md']]}
        self.add_payload('manifest.json',json_bytes(manifest))
        for name,payload in self.payloads.items():self.privacy_gate(payload,name)
        self.privacy_gate(result_payload,result_path.name)
        for name,payload in self.standalone.items():self.privacy_gate(payload,name)
        files=[{'file':'docs/codex/evidence/T17-C1-reports/'+n,'sha256':sha(p)} for n,p in self.payloads.items()]
        files+=[{'file':'docs/codex/evidence/'+n,'sha256':sha(p)} for n,p in self.standalone.items()]
        files.append({'file':'docs/codex/evidence/T17-C1-results.json','sha256':sha(result_payload)})
        outcome={'task':'T17','clean_commit':self.commit,'check_only':not write,'clean_reports':32,'total_passed':1158,'per_shell':shells,
                 'historical_records':sum(r['classification'].startswith('historical') for r in self.records),'public_files':len(files),
                 'manifest_sha256':sha(self.payloads['manifest.json']),'results_sha256':sha(result_payload),'files':files,
                 'literal_whitespace_waiver_suggestions':[{'file':'docs/codex/evidence/T17-C1-reports/'+n,
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
    args=parser.parse_args();repo=args.repo.resolve();assert_clean(repo,args.commit);collector=T17Collector(repo,args.commit)
    raw,counts=load_json(collector.work/'T17-expected-counts.json');require(counts==COUNTS and tuple(counts)==TIERS,'Frozen16tier counts changed')
    collector.counts_digest=sha(raw);collector.add_payload('frozen-expected-counts.json',raw)
    shells=[collector.clean_shell(label,collector.work/('T17-C1-'+label)) for label in ('ps51','ps7')]
    collector.historical()
    for leaf in ('Run-T17Checkpoint.ps1','T17-expected-counts.json','Collect-T17Evidence.py','Collect-T15Evidence.py','Collect-T14Evidence.py','Collect-T09C3Evidence.py',
                 'Run-T17Clean.py','T17-C1-commit.txt'):
        collector.archive(collector.work/leaf,'clean C1 execution source/capture bytes; separate from Pester case counts')
    for label in ('ps51','ps7'):
        execution_path=collector.work/f'T17-C1-{label}.outer.execution.json';_,execution=load_json(execution_path)
        require(execution['task']=='T17' and execution['phase']=='C1' and execution['commit_under_test']==collector.commit
            and execution['exit_code']==0 and execution['child_only_modulepath_removed'] is True
            and execution['wrapper_sha256']==sha((collector.work/'Run-T17Clean.py').read_bytes())
            and execution['driver_sha256']==sha((collector.work/'Run-T17Checkpoint.ps1').read_bytes()),'Outer actual clean driver wrapper/context differs')
        collector.archive(execution_path,'clean actual outer driver invocation/execution; no additional test counts')
        for stream in ('stdout','stderr'):
            source=collector.owned_input(execution[stream]);require(source.name==f'T17-C1-{label}.outer.{stream}.txt'
                and sha(source.read_bytes())==execution[stream+'_sha256'],'Outer genuine separate raw capture differs')
            collector.archive(source,'clean actual outer driver separate raw '+stream+'; no additional test counts')
    collector.environment_and_pins(collector.work/'T17-environment.json');collector.verified_cache_receipt(collector.work/'T17-cache-verification.json')
    for shell in shells:collector.analyzer(shell['shell'],collector.work/('T17-C1-analyzer-'+shell['shell']+'.json'),shell['shell_version'])
    collector.manual_visual_review(collector.work/'T17-C1-visual-review.json')
    assert_clean(repo,args.commit);collector.finish(shells,repo/'docs/codex/evidence/T17-C1-reports',args.write)

if __name__=='__main__':main()
