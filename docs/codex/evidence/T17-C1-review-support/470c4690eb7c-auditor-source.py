"""Independent read-only audit of retained T17 native size-accounting evidence."""
import argparse
from contextlib import closing
from datetime import datetime, timezone
from decimal import Decimal, localcontext, ROUND_HALF_EVEN
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
import pypdfium2 as pdfium

REPO=Path(__file__).resolve().parent
while not (REPO/'WinPDFMerge.ps1').is_file():
    REPO=REPO.parent
WORK=REPO/'tests/.work'
SOURCE_FILES=[
    'WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tools/test/Invoke-Tests.ps1',
    'tests/pdf/SizeReporting.Tests.ps1','tests/pdf/SizeReporting.Native.Tests.ps1',
    'tests/fixtures/presets/generate_presets.py','tests/fixtures/presets/manifest.json',
    'tests/fixtures/numbered/1.pdf','tests/fixtures/numbered/manifest.json',
    'README.md','docs/EMAIL_PRESETS.md',
]
ALLOWED_SOURCE={REPO/name for name in SOURCE_FILES}
checks,findings,bindings,fresh,cases,run_bindings,history=[],[],[],[],[],[],[]
seen=set()


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def label(path):
    return Path(path).resolve().relative_to(REPO).as_posix()


def owned(path,file=True):
    path=Path(path).resolve()
    if not path.is_relative_to(WORK) or (file and not path.is_file()):
        raise ValueError('Existing owned synthetic evidence operand required')
    return path


def bind(path):
    path=Path(path).resolve()
    if path not in ALLOWED_SOURCE:
        path=owned(path)
    if path not in seen:
        bindings.append(dict(Path=label(path),SHA256=sha(path),Bytes=path.stat().st_size))
        seen.add(path)
    return path


def load(path):
    raw=bind(path).read_bytes()
    return json.loads(raw.decode('utf-16' if raw.startswith((b'\xff\xfe',b'\xfe\xff')) else 'utf-8-sig'))


def check(condition,message):
    checks.append(dict(Check=message,Passed=bool(condition)))
    if not condition:
        findings.append(message)


def git(*arguments):
    return subprocess.check_output(['git',*arguments],cwd=REPO,text=True).strip()


def canonical(value):
    return hashlib.sha256(json.dumps(value,sort_keys=True,separators=(',',':')).encode()).hexdigest()


def snapshot(row,context):
    path=bind(row['Path'])
    stat=path.stat()
    check(sha(path)==row['SHA256'].lower(),context+' current SHA256 '+label(path))
    check(stat.st_size==row['Length'],context+' current exact length '+label(path))
    check(stat.st_mtime_ns//100+621355968000000000==row['ModifiedUtcTicks'],context+' current UTC ticks '+label(path))
    check(stat.st_file_attributes==row['Attributes'],context+' current attributes '+label(path))
    check(not stat.st_file_attributes&1024,context+' ordinary non-reparse file '+label(path))
    return path


def native(receipt,context,executable,exit_code=0):
    fields=('Started','ExitCode','Succeeded','TimedOut','Cancelled','LaunchError','CaptureError',
            'TerminationError','StdoutTruncated','StderrTruncated','OwnershipReleased')
    check(all(name in receipt for name in fields),context+' required receipt fields')
    check(receipt.get('Started') is True and receipt.get('ExitCode')==exit_code
          and receipt.get('Succeeded') is (exit_code==0),context+' actual bounded process outcome')
    check(receipt.get('TimedOut') is False and receipt.get('Cancelled') is False,context+' no timeout/cancel')
    check(receipt.get('OwnershipReleased') is True and receipt.get('StdoutTruncated') is False
          and receipt.get('StderrTruncated') is False,context+' ownership released and complete captures')
    check(all(not receipt.get(name) for name in ('LaunchError','CaptureError','TerminationError')),context+' no launcher/capture/termination error')
    check(receipt.get('ProcessId',0)>0,context+' actual retained process identity')
    check(Path(receipt['Executable'])==executable,context+' selected executable')


def expected_pages(expectation):
    return [dict(identifier=identifier,rotation_degrees=0,
                 size_points=[float(value) for value in expectation['page_size_points']])
            for identifier in expectation['page_identifiers']]


def inspect_pdf(path,expectation,context,kind):
    path=owned(path)
    before=sha(path)
    stat=path.stat()
    command=[str(PDFTK),str(path),'dump_data_utf8','output','-','dont_ask']
    started=time.monotonic()
    process=subprocess.Popen(command,stdin=subprocess.DEVNULL,stdout=subprocess.PIPE,stderr=subprocess.PIPE,
                             cwd=REPO,creationflags=subprocess.CREATE_NO_WINDOW)
    try:
        stdout,stderr=process.communicate(timeout=15)
    except subprocess.TimeoutExpired:
        process.kill()
        process.communicate(timeout=2)
        raise RuntimeError('Read-only approved PDFtk exceeded finite inspection timeout')
    stdout_text=stdout.decode('utf-8')
    stderr_text=stderr.decode('utf-8')
    counts=re.findall(r'^NumberOfPages:\s*([0-9]+)\s*$',stdout_text,re.M)
    check(process.returncode==0 and counts==[str(expectation['page_count'])],context+' fresh PDFtk exact count')
    pages=[]
    with pdfium.PdfDocument(path) as document:
        for index in range(len(document)):
            with closing(document[index]) as page:
                with closing(page.get_textpage()) as text:
                    identifiers=re.findall(r'T03-[0-9]{2}-P[0-9]{2}',text.get_text_range())
                check(len(identifiers)==1,context+' fresh PDFium one visible identifier page '+str(index+1))
                pages.append(dict(identifier=identifiers[0] if len(identifiers)==1 else identifiers,
                                  rotation_degrees=page.get_rotation(),size_points=list(page.get_size())))
    check(pages==expected_pages(expectation),context+' fresh PDFium ordered IDs/rotations/dimensions')
    after=path.stat()
    check(sha(path)==before and after.st_size==stat.st_size and after.st_mtime_ns==stat.st_mtime_ns
          and after.st_file_attributes==stat.st_file_attributes,context+' fresh reads preserved bytes and metadata')
    fresh.append(dict(Path=label(path),SHA256=before,Bytes=stat.st_size,Kind=kind,
        PdfTk=dict(Command=['<approved PDFtk 2.02>',label(path),'dump_data_utf8','output','-','dont_ask'],
                   ProcessId=process.pid,ExitCode=process.returncode,ElapsedMilliseconds=int((time.monotonic()-started)*1000),
                   Stdout=stdout_text,Stderr=stderr_text,StdoutSHA256=hashlib.sha256(stdout).hexdigest(),
                   StderrSHA256=hashlib.sha256(stderr).hexdigest()),
        Pdfium=dict(PageCount=len(pages),Pages=pages)))


def human(bytes_count):
    units=('B','KiB','MiB','GiB','TiB','PiB','EiB')
    unit=min((int(bytes_count).bit_length()-1)//10,len(units)-1)
    with localcontext() as context:
        context.prec=60
        value=Decimal(bytes_count)/Decimal(1<<(10*unit))
        rounded=value.quantize(Decimal(1) if unit==0 else Decimal('0.01'),rounding=ROUND_HALF_EVEN)
        return format(rounded,'f')+' '+units[unit]


EXPECTED={
    'actual-small-print-screen':('screen','no_size_benefit','record','small-print.pdf'),
    'actual-small-print-ebook':('ebook','no_size_benefit','record','small-print.pdf'),
    'actual-scan-screen':('screen','published','record','scan.pdf'),
    'actual-scan-ebook':('ebook','published','record','scan.pdf'),
    'actual-mixed-screen':('screen','published','record','mixed.pdf'),
    'actual-mixed-ebook':('ebook','published','record','mixed.pdf'),
    'actual-tiny-larger-no-benefit-code0':('screen','no_size_benefit','record','1.pdf'),
    'controlled-equal-after-actual-GS0-code0':('screen','no_size_benefit','equal','1.pdf'),
    'actual-master-only-skip-code0':('screen','skipped','skip','1.pdf'),
    'actual-master-only-missing-code0':('screen','unavailable','record','1.pdf'),
    'actual-master-only-failure-code2':('screen','failed','corrupt','1.pdf'),
}


def audit_observations(path,commit,selection):
    data=load(path)
    prefix=selection+' '
    check(data['CommitUnderTest']==commit and data['DirtyWorktree'] is False,prefix+' clean exact C1 observations')
    check(data['ShellVersion']=={'ps51':'5.1.26100.9444','ps7':'7.6.6'}[selection],prefix+' actual approved shell')
    check(data['StandardUser'] is True and data['Process64Bit'] is True,prefix+' recorded standard-user Windows x64')
    check(data['PdfTkVersion']=='2.02' and data['GhostscriptVersion']=='10.08.0',prefix+' exact selected native versions')
    check(data['PythonSHA256']==sha(sys.executable),prefix+' development Python pin')
    check(data['OracleVersions']==dict(python='3.12.14',pypdfium2='5.13.0',pdfium='153.0.7999.0'),prefix+' independent oracle pins')
    check(data['TestSourceSHA256'].lower()==sha(REPO/'tests/pdf/SizeReporting.Native.Tests.ps1'),prefix+' frozen native test source')
    check(data['GeneratorSHA256'].lower()==sha(REPO/'tests/fixtures/presets/generate_presets.py'),prefix+' frozen generator source')
    check(data['FixtureManifestSHA256'].lower()==sha(REPO/'tests/fixtures/presets/manifest.json'),prefix+' fixture manifest bytes')
    check(data['FixtureManifest']==MANIFEST,prefix+' exact typed original fixture manifest')
    check(data['OriginalFixtureSHA256'].lower()==sha(REPO/'tests/fixtures/numbered/1.pdf'),prefix+' tiny fixture pin')
    check({row['Name']:row['SHA256'].lower() for row in data['EngineSHA256']}==ENGINE_HASHES,prefix+' all four selected engine hashes')
    root=owned(path).parent
    oracle=bind(root/'independent-size-inspection.py')
    check(sha(oracle)==data['OracleSHA256'].lower(),prefix+' actual suite oracle source binding')
    check(data['GeneratorCommand']==['-B',str(REPO/'tests/fixtures/presets/generate_presets.py'),'--output',data['OriginalCorpusDirectory']],
          prefix+' exact original generation vector')
    check(data['GeneratorResult']['ExitCode']==0 and not data['GeneratorResult']['Stderr'],prefix+' actual generator success receipt')
    check(json.loads(data['GeneratorResult']['Stdout'])==MANIFEST,prefix+' actual generation output matches manifest')
    originals=owned(data['OriginalCorpusDirectory'],file=False)
    for expectation in MANIFEST['fixtures']:
        original=bind(originals/expectation['file'])
        check(sha(original)==expectation['sha256'] and original.stat().st_size==expectation['bytes'],prefix+' regenerated original exact bytes '+expectation['file'])
        inspect_pdf(original,expectation,prefix+'original '+expectation['file'],'original corpus')
    observations=data['Observations']
    check(len(observations)==len(EXPECTED) and {row['Label'] for row in observations}==set(EXPECTED),
          prefix+' exact eleven distinct native observations')
    for row in observations:
        case=prefix+row['Label']+' '
        preset,state,control,fixture=EXPECTED[row['Label']]
        expectation=next(value for value in MANIFEST['fixtures'] if value['file']==fixture) if fixture!='1.pdf' else TINY
        check(row['FixtureExpectation']==expectation,case+' exact synthetic fixture expectation')
        check(row['Control']==control and bool(row['ControlledHooks']),case+' explicitly disclosed recording/control mode')
        if control=='equal':
            check('not actual GS equality' in row['ControlledHooks'],case+' equality supplement not relabelled genuine GS equality')
        if control=='corrupt':
            check('Job.MasterBytes is substituted corrupt input length' in row['ControlledHooks'],case+' substituted input metric limitation disclosed')
        output=owned(row['OutputFolder'],file=False)
        case_root=output.parent
        app=owned(case_root/'app',file=False)
        capture=owned(case_root/'captured-calls',file=False)
        entry=bind(app/'WinPDFMerge.ps1')
        helper=bind(app/'src/WinPDFMerge.Helpers.ps1')
        check(entry.read_bytes()==(REPO/'WinPDFMerge.ps1').read_bytes() and sha(entry)==row['EntrySHA256'].lower(),case+' exact frozen entry copy')
        check(helper.read_bytes().startswith((REPO/'src/WinPDFMerge.Helpers.ps1').read_bytes())
              and sha(helper)==row['CopiedHelperSHA256'].lower(),case+' original helper bytes and labelled recording suffix')
        check(row['Before']==row['After'] and len(row['After'])==3,case+' source/foreign/original before-after preservation')
        for preserved in row['After']:
            snapshot(preserved,case+'preservation')
        check(Path(row['After'][0]['Path'])==Path(row['SourceFolder'])/'1.pdf',case+' explicit source operand')
        check(Path(row['After'][2]['Path'])==Path(row['OriginalFixturePath']),case+' explicit original operand')
        check(row['After'][0]['SHA256'].lower()==expectation['sha256'] and row['After'][0]['Length']==expectation['bytes'],
              case+' copied source matches exact original fixture')
        proof=row['Proof']
        code=2 if state=='failed' else 0
        check(proof['ExpectedState']==state and proof['ExpectedPreset']==preset and proof['Result']['ExitCode']==code,case+' explicit observed outcome/exit/preset')
        check(Path(proof['OutputDirectory'])==output and Path(proof['MasterPath']).parent==output,case+' intended explicit destination')
        command=load(proof['CommandReceiptPath'])
        check(command['ExitCode']==code,case+' raw child exit binding')
        expected_shell=SHELLS[selection]
        check(Path(command['Executable'])==expected_shell,case+' actual approved child shell path')
        arguments=command['Arguments']
        check(arguments[:5]==['-NoProfile','-ExecutionPolicy','RemoteSigned','-File',str(entry)]
              and arguments[5]==row['SourceFolder'],case+' direct actual entry vector and source')
        check(arguments[arguments.index('-OutputFolder')+1]==str(output),case+' explicit output argument')
        if row['Label'].endswith('-ebook'):
            check(arguments[-2:]==['-EmailPreset','ebook'],case+' explicit ebook argument')
        else:
            check('-EmailPreset' not in arguments,case+' unchanged omitted screen default')
        check(('-SkipEmail' in arguments)==(state=='skipped'),case+' exact requested skip flag')
        check(command['ChildEnvironment']['GS_OPTIONS']=='-T17-invalid-inherited-child-option',case+' inherited child-only GS poison')
        check(str(GHOSTSCRIPT.parent) in command['ChildPath'] if state!='unavailable' else str(GHOSTSCRIPT.parent) not in command['ChildPath'],
              case+' selected engine availability through recorded child PATH')
        stdout_path=bind(case_root/'entry.stdout.txt')
        stderr_path=bind(case_root/'entry.stderr.txt')
        stdout=stdout_path.read_bytes().decode('utf-8-sig')
        stderr=stderr_path.read_bytes().decode('utf-8-sig')
        check(stdout==proof['Result']['Stdout'] and stderr==proof['Result']['Stderr'],case+' exact raw child stream binding')
        check(sha(stdout_path)==command['StdoutSHA256'].lower() and sha(stderr_path)==command['StderrSHA256'].lower(),case+' recorded raw child stream hashes')
        master=load(capture/'Pdftk-job.json')
        check(master==proof['MasterJob'] and 'EmailPreset' not in master['BoundParameterKeys'],case+' raw master job and unchanged PDFtk call contract')
        native(master['Job']['NativeResult'],case+'master merge',PDFTK)
        native(master['Job']['ValidationResult']['NativeResult'],case+'master inspection',PDFTK)
        check(master['Job']['Succeeded'] is True and master['Job']['OutputPublished'] is True
              and master['Job']['OutputValidated'] is True and master['Job']['ValidatedPageCount']==expectation['page_count'],
              case+' successful full native master validation/publication')
        check(master['Job']['NativeResult']['ProcessId']!=master['Job']['ValidationResult']['NativeResult']['ProcessId'],case+' separate actual merge/inspection processes')
        master_path=bind(proof['MasterPath'])
        master_bytes=master_path.stat().st_size
        check(proof['MasterBytes']==master_bytes==master['Job']['OutputBytes'],case+' exact published master length')
        log_path=bind(proof['LogPath'])
        log=log_path.read_bytes().decode('utf-8-sig')
        check(log==proof['Log'],case+' exact raw log text binding without newline translation')
        calls=load(capture/'native-calls.json')
        check(calls==proof['NativeCalls'],case+' exact original native call array')
        for index,call in enumerate(calls):
            executable=Path(call['Executable'])
            check(executable in (PDFTK,GHOSTSCRIPT),case+' selected direct executable call '+str(index))
            failure=state=='failed' and '-sDEVICE=pdfwrite' in call['Arguments']
            native(call['Result'],case+'native call '+str(index),executable,1 if failure else 0)
        merges=[call for call in calls if 'cat' in call['Arguments']]
        check(len(merges)==1 and merges[0]['Arguments']==[str(Path(row['SourceFolder'])/'1.pdf'),'cat','output',
              str(Path(master['StageDirectory'])/'master.pdf'),'dont_ask'],case+' exact original PDFtk master vector')
        check(merges[0]['Result']==master['Job']['NativeResult'],case+' actual master conversion receipt binding')
        gs_calls=[call for call in calls if Path(call['Executable'])==GHOSTSCRIPT]
        expected_lines=['Master size: '+str(master_bytes)+' bytes ('+human(master_bytes)+').']
        email=None
        candidate_bytes=None
        if state in ('skipped','unavailable'):
            check(not gs_calls and proof['EmailJob'] is None and proof['EmailPaths']==[],case+' no GS call/job/email final')
            check(not (capture/'Ghostscript-job.json').exists() and not list(capture.glob('unexpected-GS-*')),case+' no controlled discovery/probe/native sentinel reached')
            check(proof['ValidatedCandidateBytes'] is None and proof['ReductionPercentInvariant'] is None,case+' no invented candidate metric')
        else:
            email=load(capture/'Ghostscript-job.json')
            check(email==proof['EmailJob'] and email['RequestedEmailPreset'].lower()==preset and 'EmailPreset' in email['BoundParameterKeys'],case+' exact raw email job/preset')
            check(email['Control']==control,case+' same disclosed job control')
            check(email['MasterBefore']==email['MasterAfter'],case+' published master before/after optional work unchanged')
            snapshot(email['MasterAfter'],case+'master preservation')
            check(Path(email['MasterAfter']['Path'])==master_path and email['MasterAfter']['Length']==master_bytes,case+' actual published master snapshot operand')
            conversions=[call for call in gs_calls if '-sDEVICE=pdfwrite' in call['Arguments']]
            check(len(conversions)==1,case+' exactly one actual GS conversion')
            native_input=str(case_root/'owned-corrupt-gs-input.pdf') if state=='failed' else str(master_path)
            vector=['-dBATCH','-dNOPAUSE','-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite',
                    '-dCompatibilityLevel=1.6','-dPDFSETTINGS=/'+preset,'-dDetectDuplicateImages=true',
                    '-o',email['StagedEmailPath'],'-f',native_input]
            check(conversions[0]['Arguments']==vector,case+' exact fixed allowlisted GS vector/safety/output/input')
            check(conversions[0]['RemoveEnvironmentVariables']==['GS_OPTIONS'],case+' child-only GS_OPTIONS removal')
            check(conversions[0]['Result']==email['Job']['NativeResult'],case+' actual GS conversion receipt binding')
            check(email['OriginalInputPaths']==[str(master_path)] and email['ActualInputPaths']==[native_input],case+' original/actual input control disclosure')
            native(email['Job']['NativeResult'],case+'email conversion',GHOSTSCRIPT,1 if state=='failed' else 0)
            if state=='failed':
                check(email['Job']['OutputValidated'] is False and email['Job']['OutputPublished'] is False
                      and email['Job']['Succeeded'] is False and email['Job']['OutputState']=='failed',case+' actual failed candidate never validated/published')
                check(proof['ValidatedCandidateBytes'] is None and proof['ReductionPercentInvariant'] is None,case+' failed candidate supplies no success metric')
                partial=email['FailedPartialBeforeCleanup']
                retained=bind(capture/'retained-failed-partial.pdf')
                check(partial['Length']>0 and partial['SHA256'].lower()==sha(retained)
                      and partial['Length']==retained.stat().st_size,case+' honest retained failed partial exact bytes')
                check(not Path(partial['Path']).exists(),case+' owned partial removed from original staging')
                check('Email result: failed' in log and 'PARTIAL SUCCESS:' in stdout,case+' actual failure console/log outcome')
            else:
                native(email['Job']['ValidationResult']['NativeResult'],case+'email strict inspection',PDFTK)
                check(email['Job']['OutputValidated'] is True and email['Job']['Succeeded'] is True
                      and email['Job']['ValidatedPageCount']==expectation['page_count'],case+' successful full native derivative validation')
                check(email['Job']['NativeResult']['ProcessId']!=email['Job']['ValidationResult']['NativeResult']['ProcessId'],case+' separate actual email conversion/inspection processes')
                candidate= snapshot(email['RetainedValidatedCandidate'],case+'retained validated candidate')
                candidate_bytes=candidate.stat().st_size
                check(candidate_bytes==email['Job']['OutputBytes']==proof['ValidatedCandidateBytes'],case+' exact validated candidate length')
                check(email['Job']['MasterBytes']==master_bytes,case+' candidate comparison uses unchanged published master length')
                with localcontext() as decimal_context:
                    decimal_context.prec=60
                    reduction=Decimal(100)*(Decimal(master_bytes)-Decimal(candidate_bytes))/Decimal(master_bytes)
                    displayed=format(reduction.quantize(Decimal('0.1'),rounding=ROUND_HALF_EVEN),'f')
                recorded_reduction=Decimal(proof['ReductionPercentInvariant'])
                check(abs(recorded_reduction-reduction)<=Decimal('1e-25'),case+' independently recomputed decimal reduction')
                if state=='published':
                    check(candidate_bytes<master_bytes and email['Job']['OutputPublished'] is True,case+' actual strictly-smaller publication')
                    check(proof['EmailPaths']==[email['OutputPath']] and Path(email['OutputPath']).is_file(),case+' explicit published email path')
                    check(sha(email['OutputPath'])==sha(candidate) and Path(email['OutputPath']).stat().st_size==candidate_bytes,case+' final/candidate byte identity')
                    expected_lines+=['Email size: '+str(candidate_bytes)+' bytes ('+human(candidate_bytes)+').',
                                     'Email reduction: '+displayed+'%.']
                else:
                    check(candidate_bytes>=master_bytes and email['Job']['OutputPublished'] is False
                          and proof['EmailPaths']==[] and not Path(email['OutputPath']).exists(),case+' no-benefit candidate omitted from final outputs')
                    if control=='record':
                        check(candidate_bytes>master_bytes,case+' genuine real GS larger candidate')
                    expected_lines+=['Validated email candidate size: '+str(candidate_bytes)+' bytes ('+human(candidate_bytes)+'); not published.',
                                     'Email candidate reduction: '+displayed+'% (no size benefit; candidate not published).']
                    check(' - Email-optimized:' not in stdout,case+' no false successful email-final path')
                if control=='equal':
                    boundary=load(capture/'equal-boundary.json')
                    check(boundary==proof['EqualBoundary'],case+' exact equality supplement receipt')
                    actual=bind(boundary['RetainedActualGhostscriptCandidate'])
                    original_candidate=boundary['ActualGhostscriptCandidate']
                    check(actual.stat().st_size==original_candidate['Length'] and sha(actual)==original_candidate['SHA256'].lower(),case+' actual GS candidate before injection retained')
                    check(boundary['InjectedCandidate']['SHA256'].lower()==sha(candidate)==sha(master_path)
                          and candidate_bytes==master_bytes,case+' equality substitution exact unchanged master bytes')
                    check('not a genuine GS equal-size result' in boundary['Scope'],case+' truthful controlled equality scope')
                    inspect_pdf(actual,expectation,case+'actual GS pre-equality candidate','actual GS before equality control')
        check(proof['ExpectedSizeLines']==expected_lines,case+' independently computed authoritative metric lines')
        size_pattern=r'^(?:Master size:|Email size:|Email reduction:|Validated email candidate size:|Email candidate reduction:)'
        all_console=[line for line in stdout.splitlines() if re.match(size_pattern,line)]
        summary_text=re.split(r'(?m)^(?:SUCCESS|PARTIAL SUCCESS|FAILURE):',stdout)[-1]
        summary_lines=[line for line in summary_text.splitlines() if re.match(size_pattern,line)]
        log_lines=[line for line in log.splitlines() if re.match(size_pattern,line)]
        check(all(line in expected_lines for line in all_console),case+' all echoed console metrics truthful')
        check(summary_lines==expected_lines and log_lines==expected_lines,case+' exact ordered final-summary and log metric lines')
        check('Email result: '+state in log and 'Result: '+('PARTIAL SUCCESS' if code else 'SUCCESS')+'; exit code: '+str(code) in log,
              case+' exact logged stage/final exit states')
        check(not Path(master['StageDirectory']).exists(),case+' original master/email stage removed')
        if email is not None:
            check(not Path(email['StagedEmailPath']).exists(),case+' owned candidate absent after cleanup')
        for directory in (app,Path(row['SourceFolder']),output):
            check(not list(directory.glob('.WinPDFMerge*')),case+' no owned orphan stage in '+label(directory))
        final_paths=sorted(path.name for path in output.glob('WinPDFMerge_*.pdf'))
        explicit=[Path(proof['MasterPath']).name]+[Path(path).name for path in proof['EmailPaths']]
        check(final_paths==sorted(explicit),case+' only explicit validated finals, foreign preserved separately')
        reads=proof['FinalReads']
        expected_read_paths=[row['After'][0]['Path'],proof['MasterPath']]
        if candidate_bytes is not None:
            expected_read_paths.append(email['RetainedValidatedCandidate']['Path'])
        if state=='published':
            expected_read_paths+=proof['EmailPaths']
        check([read['Snapshot']['Path'] for read in reads]==expected_read_paths,case+' complete explicit original/master/candidate/final read operands')
        fresh_before=len(fresh)
        for read in reads:
            path=snapshot(read['Snapshot'],case+'original retained read')
            check(read['PdfTkRead']['ExitCode']==0 and re.findall(r'^NumberOfPages:\s*([0-9]+)\s*$',read['PdfTkRead']['Stdout'],re.M)==[str(expectation['page_count'])],
                  case+' recorded exact PDFtk read count')
            check(read['Oracle']['page_count']==expectation['page_count'] and read['Oracle']['pages']==expected_pages(expectation),case+' recorded independent PDFium result')
            check(json.loads(read['OracleRead']['Stdout'])==read['Oracle'] and read['OracleRead']['ExitCode']==0,case+' original oracle stream/result binding')
            inspect_pdf(path,expectation,case+'fresh retained read','source/master/candidate/final')
        cases.append(dict(Shell=selection,Label=row['Label'],ObservedState=state,ExitCode=code,Control=control,
                          MasterBytes=master_bytes,ValidatedCandidateBytes=candidate_bytes,
                          ExpectedSizeLines=expected_lines,FreshRetainedReads=len(fresh)-fresh_before))


parser=argparse.ArgumentParser()
parser.add_argument('--commit',required=True)
parser.add_argument('--observations',action='append',required=True)
parser.add_argument('--run-snapshot',action='append',required=True)
parser.add_argument('--prior-attempt',action='append',default=[])
parser.add_argument('--output',required=True,type=Path)
args=parser.parse_args()
output=owned(args.output,file=False)
assert not output.exists(),'Never overwrite an audit report'
started=datetime.now(timezone.utc).isoformat()
partial=True
source_before={}
try:
    check(sys.platform=='win32','Actual independent audit runs on Windows')
    check(git('rev-parse','HEAD')==args.commit and not git('status','--porcelain=v1'),'Exact clean C1 before audit')
    source_before={name:sha(REPO/name) for name in SOURCE_FILES}
    check(sys.version.split()[0]=='3.12.14' and str(pdfium.PYPDFIUM_INFO)=='5.13.0'
          and str(pdfium.PDFIUM_INFO)=='153.0.7999.0','Actual pinned Python/PDFium versions')
    cache=load(WORK/'T17-cache-verification.json')
    check(sha(sys.executable)==cache['development_oracle_runtime']['python_sha256'],'Actual development Python executable bytes')
    dll=Path(cache['development_oracle_runtime']['pdfium_dll_path'].replace('<USERPROFILE>',os.environ['USERPROFILE']))
    check(sha(dll)==cache['development_oracle_runtime']['pdfium_dll_sha256'],'Actual independent PDFium engine bytes')
    ENGINE_HASHES={}
    engines={}
    for dependency in cache['dependencies']:
        if dependency['dependency'] not in ('PDFtk','Ghostscript'):
            continue
        cache_root=Path(os.path.expandvars(dependency['cache_root']))
        for file in dependency['selected_files']:
            executable=cache_root/file['relative_path']
            check(sha(executable)==file['sha256'],'Actual selected engine bytes '+executable.name)
            ENGINE_HASHES[executable.name]=file['sha256']
            engines[executable.name]=executable
    PDFTK,GHOSTSCRIPT=engines['pdftk.exe'],engines['gswin64c.exe']
    ps7_receipt=json.loads((REPO/'docs/codex/evidence/T09-ps7-acquisition.json').read_bytes())
    SHELLS=dict(ps51=Path(os.environ['SystemRoot'])/'System32/WindowsPowerShell/v1.0/powershell.exe',
                ps7=Path(os.path.expandvars(ps7_receipt['cache']['directory_label']))/ps7_receipt['cache']['executable_relative_path'])
    check(sha(SHELLS['ps7'])==ps7_receipt['executable']['sha256'],'Actual approved PS7 executable bytes')
    MANIFEST=json.loads((REPO/'tests/fixtures/presets/manifest.json').read_bytes())
    TINY=next(row for row in json.loads((REPO/'tests/fixtures/numbered/manifest.json').read_bytes())['fixtures'] if row['file']=='1.pdf')
    for prior in args.prior_attempt:
        root=owned(prior,file=False)
        invocation=load(root/'invocation.json')
        execution=load(root/'execution.json')
        report=load(root/'report.json')
        check(report['CommitUnderTest']==args.commit and report['Result']!='pass','Prior audit-only attempt context')
        check(sha(root/'auditor-source.py')==invocation['AuditScriptSHA256'],'Prior auditor exact source snapshot')
        for path,digest in execution['RawBindings'].items():
            check(sha(REPO/path)==digest,'Prior exact retained attempt component '+path)
        history.append(dict(Attempt=label(root),Result=report['Result'],Partial=report['Partial'],ExitCode=execution['ExitCode'],
                            CheckCount=report['CheckCount'],FreshFinalReads=report['FreshFinalReads'],Findings=report['Findings'],
                            Scope='Auditor preparation/execution only; no additional application/suite acceptance cases.'))
    snapshots=dict(value.split('=',1) for value in args.run_snapshot)
    check(set(snapshots)=={'ps51','ps7'} and len(args.observations)==2,'Both actual shell run snapshots/observation operands supplied')
    for selection in ('ps51','ps7'):
        snapshot_path=owned(snapshots[selection])
        runs=load(snapshot_path)
        selected=[run for run in runs if run['tier']=='SizeReportingNative']
        check(len(selected)==1,selection+' exactly one clean native size tier')
        run=selected[0]
        current_runs=json.loads((WORK/('T17-C1-'+selection)/'runs.json').read_bytes().decode('utf-8-sig'))
        check([row for row in current_runs if row['tier']=='SizeReportingNative']==[run],selection+' native row unchanged in active run index')
        run_bindings.append(dict(Shell=selection,Source='tests/.work/T17-C1-'+selection+'/runs.json',
                                 SourceSHA256AtCapture=sha(snapshot_path),Snapshot=label(snapshot_path),
                                 NativeRowSHA256=canonical(run),
                                 Limit='Exact run-index bytes captured while other tiers may append; immutable native row checked separately.'))
        check(run['exit_code']==0 and run['native_test_host_started'] is True and run['timed_out'] is False
              and not run['capture_error'] and not run['termination_error'],selection+' actual bounded native-tier process outcome')
        summary=load(Path(run['report'])/'summary.json')
        check(summary==run['summary'],selection+' raw summary binding')
        check(summary['commit_under_test']==args.commit and summary['dirty_worktree'] is False
              and summary['shell_version']=={'ps51':'5.1.26100.9444','ps7':'7.6.6'}[selection]
              and summary['total']==summary['passed']==run['expected_count']==len(EXPECTED)
              and all(summary[name]==0 for name in ('failed','failed_blocks','failed_containers','skipped','not_run')),
              selection+' clean exact eleven-case total with no failed/absent cases')
        xml=ET.fromstring(bind(Path(run['report'])/'results.xml').read_bytes())
        xml_cases=list(xml.iter('test-case'))
        check(len(xml_cases)==len(EXPECTED) and all(case.attrib.get('executed')=='True'
              and case.attrib.get('success')=='True' and case.attrib.get('result')=='Success' for case in xml_cases),
              selection+' all individual raw NUnit cases passed')
        check(Path(run['executable'])==SHELLS[selection],selection+' actual approved tier host')
        arguments=run['arguments']
        for option,value in [('-File',str(REPO/'tools/test/Invoke-Tests.ps1')),('-Tier','SizeReportingNative'),
                             ('-PdftkPath',str(PDFTK)),('-GhostscriptPath',str(GHOSTSCRIPT)),('-PythonPath',sys.executable)]:
            check(option in arguments and arguments[arguments.index(option)+1]==value,selection+' exact tier argument '+option)
        bind(run['stderr_log'])
        stdout=bind(run['log']).read_bytes().decode('utf-8-sig')
        markers=re.findall(r'Size reporting observations: ([^\r\n]+)',stdout)
        check(len(markers)==1,selection+' unique native observation marker')
        path=owned(markers[0])
        check(path in [owned(value) for value in args.observations],selection+' explicit native observation operand bound to raw stdout')
        audit_observations(path,args.commit,selection)
    check(len(cases)==2*len(EXPECTED),'All distinct native observations audited, not guessed from aggregate count')
    check({name:sha(REPO/name) for name in SOURCE_FILES}==source_before,'Frozen tracked sources unchanged after audit')
    check(git('rev-parse','HEAD')==args.commit and not git('status','--porcelain=v1'),'Exact clean C1 after audit')
    partial=False
except Exception as failure:
    check(False,'Audit preparation/execution exception: '+type(failure).__name__+': '+str(failure))

report=dict(SchemaVersion=1,Task='T17',Result='pass' if not findings and not partial else 'fail',Partial=partial,
    CommitUnderTest=args.commit,StartedAtUtc=started,CompletedAtUtc=datetime.now(timezone.utc).isoformat(),
    Auditor='tests/.work/Audit-T17Native.py',AuditorSHA256=sha(__file__),
    ActualVersions=dict(Python=sys.version.split()[0],Pypdfium2=str(pdfium.PYPDFIUM_INFO),Pdfium=str(pdfium.PDFIUM_INFO)),
    CheckCount=len(checks),CaseCount=len(cases),FreshFinalReads=len(fresh),
    FreshReadKinds={kind:sum(read['Kind']==kind for read in fresh) for kind in sorted({read['Kind'] for read in fresh})},
    FreshPageInspections=sum(read['Pdfium']['PageCount'] for read in fresh),
    Findings=findings,Checks=checks,Cases=cases,RawBindings=bindings,FreshReads=fresh,RunSnapshotBindings=run_bindings,
    SourceBindings=[dict(Path=name,SHA256=digest) for name,digest in source_before.items()],AuditHistory=history,
    Limits=['No suite/application/GS conversion/fixture generation rerun; historical native calls bind original exact retained receipts and source copies.',
            'Fresh actual read-only PDFtk/PDFium checks cover explicit original/source/master/validated-candidate/final operands; failed partial is hashed but not called a valid PDF.',
            'Controlled equality injects unchanged published-master bytes after actual GS success; corrupt input and skip sentinels are disclosed controlled supplements.',
            'This structural/size audit is not visual fidelity, owner or Explorer desktop acceptance. Root records actual visual observations separately.',
            'Metadata/hash preservation verifies retained synthetic files at audit time; no transaction guarantee against later external modification.',
            'Full remaining tiers, public archive, feature preservation, package/release and universal PDF/security acceptance are separate.'])
with output.open('x',encoding='utf-8') as stream:
    json.dump(report,stream,indent=2)
    stream.write('\n')
print(json.dumps(dict(Result=report['Result'],Partial=partial,CheckCount=len(checks),CaseCount=len(cases),
                      FreshFinalReads=len(fresh),FreshPageInspections=report['FreshPageInspections'],
                      Report=label(output),SHA256=sha(output),Findings=findings)))
sys.exit(0 if report['Result']=='pass' else 1)
