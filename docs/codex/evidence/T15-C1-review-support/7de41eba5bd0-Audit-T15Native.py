"""Ignored independent T15 fault/native audit; never reruns application/suites.

Requires both exact clean C1 FaultRecovery observation paths. Reads their retained
synthetic final PDFs through approved PDFtk and independent pinned PDFium. Case
expectations are explicit here, not inferred from filesystem success or copied
Oracle values. This receipt excludes full XML/archive/publication closure review.
"""
from __future__ import annotations
import argparse
from contextlib import closing
import ctypes
from datetime import datetime, timezone
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

REPO = Path(__file__).resolve().parents[2]
WORK = REPO / 'tests/.work'
checks, findings, receipts, fresh, audited_cases = [], [], [], [], []
seen_receipts = set()

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def owned(path, must_exist=True):
    path = Path(path).resolve()
    if not path.is_relative_to(WORK) or (must_exist and not path.is_file()):
        raise ValueError('Audit requires an existing exact owned synthetic operand.')
    return path

def label(path):
    return Path(path).resolve().relative_to(REPO).as_posix()

def load(path):
    path = owned(path)
    bind(path)
    raw = path.read_bytes()
    encoding = 'utf-16' if raw.startswith((b'\xff\xfe',b'\xfe\xff')) else 'utf-8-sig'
    return json.loads(raw.decode(encoding))

def bind(path):
    path = owned(path)
    if path not in seen_receipts:
        receipts.append(dict(path=label(path), sha256=sha(path), bytes=path.stat().st_size))
        seen_receipts.add(path)
    return path

def check(condition, text):
    passed = bool(condition)
    checks.append(dict(check=text, **{'pass':passed}))
    if not passed:
        findings.append(text)

def read_pdftk(executable, path):
    begin = time.monotonic()
    process = subprocess.Popen([str(executable),str(path),'dump_data_utf8','output','-','dont_ask'],
        stdin=subprocess.DEVNULL,stdout=subprocess.PIPE,stderr=subprocess.PIPE,cwd=REPO,
        creationflags=subprocess.CREATE_NO_WINDOW)
    try:
        stdout,stderr = process.communicate(timeout=15)
    except subprocess.TimeoutExpired:
        process.kill()
        process.communicate(timeout=2)
        raise RuntimeError('Exact read-only PDFtk inspection exceeded the finite bound.')
    return dict(started=True,process_id=process.pid,exit_code=process.returncode,
        elapsed_ms=int((time.monotonic()-begin)*1000),stdout=stdout.decode('utf-8','replace'),stderr=stderr.decode('utf-8','replace'))

def audit_snapshot(text):
    entries = [json.loads(line) for line in text.splitlines() if line.strip()]
    check(bool(entries), 'Nonempty preservation snapshot')
    for row in entries:
        path = owned(row['Path'])
        stat = path.stat()
        check(sha(path) == row.get('SHA256',row.get('Hash','')).lower(), 'Current preserved hash: '+label(path))
        if 'Length' in row:
            check(stat.st_size == row['Length'], 'Current preserved length: '+label(path))
        ticks = row.get('ModifiedUtcTicks',row.get('Modified'))
        if ticks is not None:
            check(stat.st_mtime_ns//100+621355968000000000 == ticks, 'Current preserved UTC metadata: '+label(path))
        if 'Attributes' in row:
            check(stat.st_file_attributes == row['Attributes'], 'Current preserved attributes: '+label(path))

def native(result, name, executable, kind='success'):
    fields=('Started','ExitCode','Succeeded','TimedOut','Cancelled','LaunchError',
            'CaptureError','TerminationError','StdoutTruncated','StderrTruncated','OwnershipReleased')
    check(all(key in result for key in fields), name+' complete native receipt')
    check(result['Executable'] == str(executable), name+' exact selected executable')
    check(result['OwnershipReleased'] is True and not result['CaptureError'] and
          not result['TerminationError'] and result['StdoutTruncated'] is False and
          result['StderrTruncated'] is False, name+' released ownership and complete captures')
    if kind == 'start-failure':
        check(result['Started'] is False and result['Succeeded'] is False and bool(result['LaunchError'])
              and result['ProcessId'] is None and result['ExitCode'] is None,
              name+' real invalid-image launch failure')
    else:
        check(result['Started'] is True and result['ProcessId'] > 0 and not result['LaunchError'],name+' actual process start')
        if kind == 'success':
            check(result['Succeeded'] is True and result['ExitCode'] == 0 and
                  result['TimedOut'] is False and result['Cancelled'] is False,name+' successful native exit')
        else:
            check(result['Succeeded'] is False and result['ExitCode'] != 0 and
                  result['TimedOut'] is (kind == 'timeout') and result['Cancelled'] is (kind == 'cancel'),name+' explicit '+kind+' outcome')
    check(0 <= result['ElapsedMilliseconds'] < 45000,name+' finite observed invocation')

def state(value, name):
    expected=value['State']
    check(expected in ('unset','empty','value') and value['Present'] is (expected != 'unset'),name+' OS presence oracle')
    check((value['Value'] is None if expected == 'unset' else
           value['Value'] == '' if expected == 'empty' else value['Value'] == '-T15-invalid-inherited-option'),name+' actual OS value')

def tree(observation, result, fixture, name):
    rows=observation['Tree']
    check(len(rows) == 3 and len({row['pid'] for row in rows}) == 3,name+' three distinct owned process identities')
    check(rows[0]['pid'] == result['ProcessId'] and rows[1]['parent_pid'] == rows[0]['pid']
          and rows[2]['parent_pid'] == rows[1]['pid'],name+' actual parent-child-grandchild chain')
    for row in rows:
        check(row['start_utc_ticks'] > 0,name+' recorded start identity '+row['role'])
    vector=re.fullmatch(r'"owned-tree" "30000" "([^"]+)" "[^"]+" "(?:stay|exit-parent)" "0"',result['RenderedArguments'])
    check(vector is not None,name+' exact controlled tree vector')
    if vector is not None:
        prefix=owned(vector[1],must_exist=False)
        for row in rows:
            retained=load(str(prefix)+'-'+row['role']+'.json')
            check(retained == row,name+' exact retained role receipt '+row['role'])
        bind(str(prefix)+'-ready.txt')
    sentinel=observation['UnrelatedSurvived']
    check(sentinel['Alive'] is True and sentinel['ProcessId'] not in {row['pid'] for row in rows}
          and sentinel['StartedUtcTicks'] > 0 and sentinel['Executable'] == str(fixture),name+' asserted unrelated same-image survival')
    # These historical liveness assertions are independently bound to test source,
    # observations and passing XML later. Do not imply a new concurrent observation.

def tier_binding(path, shell, commit, executable, gs):
    shell_label='ps51' if shell == '5.1.26100.9444' else 'ps7'
    runs=load(WORK/('T15-C1-'+shell_label)/'runs.json')
    check(len(runs) == 17,shell+' complete clean tier command inventory before audit')
    matching=[row for row in runs if row['tier'] == 'FaultRecovery']
    check(len(matching) == 1,shell+' exactly one clean native tier command')
    run=matching[0];name=shell+' FaultRecovery tier '
    check(run['exit_code'] == 0 and run['native_test_host_started'] is True and run['timed_out'] is False
          and not run['capture_error'] and not run['termination_error'],name+'actual host completion')
    summary=load(Path(run['report'])/'summary.json')
    check(run['summary'] == summary,name+'raw summary matches outer command receipt')
    check(summary['commit_under_test'] == commit and summary['dirty_worktree'] is False and
          summary['shell_version'] == shell and summary['process_64_bit'] is True and
          summary['pester_version'] == '6.2.0' and summary['execution_policy'] == 'RemoteSigned',name+'clean C1 actual pinned environment')
    check(run['expected_count'] == summary['total'] == summary['passed'] == 14 and
          all(summary[key] == 0 for key in ('failed','failed_blocks','failed_containers','skipped','not_run')),name+'all fourteen cases passed without exclusions')
    arguments=run['arguments']
    check(arguments[arguments.index('-Tier')+1] == 'FaultRecovery' and
          Path(arguments[arguments.index('-File')+1]).resolve() == REPO/'tools/test/Invoke-Tests.ps1',name+'exact test entry and tier command')
    check(Path(arguments[arguments.index('-PdftkPath')+1]).resolve() == executable.resolve() and
          Path(arguments[arguments.index('-GhostscriptPath')+1]).resolve() == gs.resolve() and
          Path(arguments[arguments.index('-PythonPath')+1]).resolve() == Path(sys.executable).resolve(),name+'explicit approved engine/oracle operands')
    log=bind(run['log']);stderr=bind(run['stderr_log'])
    check(stderr.stat().st_size == 0,name+'no outer stderr')
    check('Fault recovery observations: '+str(Path(path).resolve()) in log.read_text(encoding='utf-8-sig'),name+'original stdout binds exact native observation path')
    xml=bind(Path(run['report'])/'results.xml');document=ET.parse(xml).getroot()
    check(document.attrib['total'] == '14' and document.attrib['errors'] == '0' and document.attrib['failures'] == '0'
          and document.attrib['not-run'] == '0',name+'raw NUnit summary counts')
    cases=document.findall('.//test-case')
    check(len(cases) == 14 and all(row.attrib.get('success') == 'True' and row.attrib.get('executed') == 'True'
                                 and row.attrib.get('result') == 'Success' for row in cases),name+'all individual NUnit outcomes')

def audit_observations(path, commit, executable, gs, engine_pins):
    document=load(path)
    shell=document['ShellVersion'];prefix=shell+' '
    tier_binding(path,shell,commit,executable,gs)
    check(document['CommitUnderTest'] == commit and document['DirtyWorktree'] is False,prefix+'clean native C1 header')
    check(document['StandardUser'] is True and document['Process64Bit'] is True,prefix+'actual standard-user x64 observation')
    check(document['PdfTkVersion'] == '2.02' and document['GhostscriptVersion'] == '10.08.0',prefix+'approved engine versions')
    check({item['Name']:item['SHA256'] for item in document['EngineSHA256']} == engine_pins,prefix+'four engine pin bindings')
    check(document['OracleVersions'] == {'python':'3.12.14','pypdfium2':'5.13.0','pdfium':'153.0.7999.0'},prefix+'pinned independent oracle versions')
    check(document['OuterCallerBefore'] == document['OuterCallerAfter'] and
          document['OuterPathBeforeSHA256'] == document['OuterPathAfterSHA256'],prefix+'outer caller GS_OPTIONS and PATH restored')
    build=load(document['ControlledFixtureBuildReceipt'])
    fixture=owned(Path(path).parent/'ControlledNative.exe')
    check(sha(fixture) == document['ControlledFixtureSHA256'] == build['executable_sha256'],prefix+'controlled binary exact build binding')
    check(sha(REPO/'tests/native/FakeNative.cs') == document['ControlledFixtureSourceSHA256'] == build['source_sha256'],prefix+'controlled fixture source binding')
    oracle=bind(document['OraclePath']);generator=bind(oracle.parent/'original-fault-raster.py')
    check(sha(oracle) == document['OracleSHA256'].lower() and sha(generator) == document['GeneratorSHA256'].lower(),prefix+'retained oracle/generator binding')
    labels=[item['Label'] for item in document['Observations']]
    expected={f'environment-{s}-{m}' for s in ('unset','empty','value') for m in ('success','start-failure','log-fault')}
    expected.update(('cancel-before-master','cancel-after-master','cancel-after-real-master-staged-inspection-before-move',
                     'controlled-nested-tree-timeout','controlled-nested-tree-parent-exits-first'))
    check(len(labels) == len(expected) == 14 and set(labels) == expected,prefix+'complete 14-case inventory without duplicates')
    operands=[]
    for observation in document['Observations']:
        case=observation['Label'];name=prefix+case
        if case.startswith('controlled-nested-tree-'):
            result=observation['NativeResult'];kind='timeout' if case.endswith('-timeout') else 'success'
            native(result,name,fixture,kind);tree(observation,result,fixture,name)
            audited_cases.append(dict(shell=shell,label=case,scope='Controlled actual owned process tree; historical sentinel liveness assertion'))
            continue
        before=observation['SourceAndForeignBefore'];after=observation['SourceAndForeignAfter']
        check(before == after,name+' source/foreign before-after equality');audit_snapshot(after)
        proof=observation['Proof'];capture=proof['Capture'];outcome=capture['Outcome'];parameters=capture['OutcomeParameters']
        output=owned(proof['Output'],must_exist=False)
        check(output.is_dir(),name+' retained owned output directory')
        check(not any(item.name.startswith('.WinPDFMerge') for item in output.iterdir()),name+' no remaining private stage after confirmed ownership release')
        app=output.parent/'app';entry=bind(app/'WinPDFMerge.ps1');helper=bind(app/'src/WinPDFMerge.Helpers.ps1')
        check(entry.read_bytes() == (REPO/'WinPDFMerge.ps1').read_bytes(),name+' exact original entry copy')
        check(helper.read_bytes().startswith((REPO/'src/WinPDFMerge.Helpers.ps1').read_bytes()),name+' original helper prefix plus disclosed controlled seams')
        captured=load(output.parent/'fault-capture.json');check(captured == capture,name+' capture file matches observation object')
        log=bind(proof['LogPath']);raw=log.read_bytes();encoding='utf-16' if raw.startswith((b'\xff\xfe',b'\xfe\xff')) else 'utf-8-sig'
        check(raw.decode(encoding) == proof['Log'],name+' exact final log bytes')
        state(capture['CallerBefore'],name+' caller before');state(capture['CallerAfter'],name+' caller after')
        check(capture['CallerBefore'] == capture['CallerAfter'],name+' caller GS_OPTIONS restored')
        code=0 if case.endswith('-success') else 1 if case in ('cancel-before-master','cancel-after-real-master-staged-inspection-before-move') else 2
        count=2 if case.endswith('-log-fault') else 0 if code == 1 else 1
        summary={0:'SUCCESS',1:'FAILURE',2:'PARTIAL SUCCESS'}[code]
        email='published' if count == 2 else 'no_size_benefit' if code == 0 else 'not_started' if code == 1 else 'failed'
        check(proof['Result']['ExitCode'] == outcome['ExitCode'] == code and outcome['Summary'] == summary,name+' original entry explicit exit/summary')
        check(outcome['EmailState'] == parameters['EmailState'] == email and parameters['MasterPublished'] is (code != 1),name+' explicit publication state')
        check(parameters['RunFailed']['IsPresent'] is (code != 0),
              name+' separate run failure flag')
        published=outcome['PublishedPaths'];reads=proof['FinalReads'] or []
        check(len(reads) == len(published) == count and [row['Path'] for row in published] == [row['Path'] for row in reads],name+' only explicit published final paths')
        actual=[item for item in output.iterdir() if item.is_file() and item.name.startswith('WinPDFMerge_') and item.suffix == '.pdf']
        check({str(item) for item in actual} == {row['Path'] for row in reads},name+' retained final inventory matches receipt')
        stdout=proof['Result']['Stdout']
        check(len(re.findall(r'^ - Merged master:',stdout,re.M)) == int(count > 0) and
              len(re.findall(r'^ - Email-optimized:',stdout,re.M)) == int(count == 2),name+' console published paths match state')
        check(re.search(r'^'+re.escape(summary)+r':',stdout,re.M) is not None,name+' truthful final console summary')
        if code != 0:check(re.search(r'^Result: SUCCESS;',proof['Log'],re.M) is None,name+' no erroneous logged success after failure')
        if proof['FinalSnapshots']:audit_snapshot(proof['FinalSnapshots'])
        identity='T03-15-P01' if case.endswith('-log-fault') else 'T03-01-P01'
        for read in reads:
            check(read['OracleExit'] == read['PdfTk']['ExitCode'] == 0 and re.findall(r'^NumberOfPages:[ \t]*([0-9]+)[ \t]*\r?$',read['PdfTk']['Stdout'],re.M) == ['1'],name+' recorded real PDFtk inspection')
            check(read['Oracle']['page_count'] == 1 and read['Oracle']['pages'] == [{'identifier':identity,'rotation_degrees':0,'size_points':[432.0,288.0]}],name+' independent recorded synthetic ID/rotation/dimensions')
            operands.append(dict(Path=read['Path'],SHA256=read['SHA256'],IDs=[identity],Rotations=[0],Sizes=[[432.0,288.0]],Purpose=name+' retained validated final'))
        if count == 2:
            check(Path(reads[1]['Path']).stat().st_size < Path(reads[0]['Path']).stat().st_size,name+' strictly smaller published email')
            check(capture['LogFaultReached'] is True and observation['FixtureGeneration']['visible_id'] == 'T03-15-P01',name+' once-only logger fault after original raster generation')
        calls=capture['NativeCalls'];master=[call for call in calls if call['Phase'] == 'master'];emails=[call for call in calls if call['Phase'] == 'email']
        check(len(master) == 1 and len(emails) == int(code != 1),name+' conversion sequencing')
        for index,call in enumerate(calls):
            call_name=name+' call '+str(index+1)
            check(call['CallerBefore'] == call['CallerAfter'] == capture['CallerBefore'],call_name+' caller environment preserved')
            kind='start-failure' if case.endswith('-start-failure') and call['Phase'] == 'email' else 'cancel' if case in ('cancel-before-master','cancel-after-master') and call['ControlledSubstitution'] else 'success'
            selected=owned(output.parent/'owned-invalid-image.exe') if kind == 'start-failure' else fixture if kind == 'cancel' else gs if call['Phase'] == 'email' or '--version' in call['RequestedArguments'] and call['RequestedExecutable'] == str(gs) else executable
            native(call['Result'],call_name,selected,kind)
            if call['MasterBefore'] is not None:
                check(call['MasterBefore'] == call['MasterAfter'],call_name+' validated master unchanged during email job')
                audit_snapshot(json.dumps(call['MasterAfter']))
            if call['Phase'] == 'email':
                native(call['EnvironmentProbe'],call_name+' environment probe',fixture)
                check(call['EnvironmentProbe']['Stdout'].strip() == '<unset>' and call['RemovedEnvironmentVariables'] == ['GS_OPTIONS'],call_name+' GS_OPTIONS absent from child block')
                check(call['RequestedExecutable'] == str(gs) and all(flag in call['RequestedArguments'] for flag in ('-dSAFER','-dPDFSTOPONERROR','-sDEVICE=pdfwrite','-dPDFSETTINGS=/screen')),call_name+' original Ghostscript vector retains safety/defaults')
            if call['StagePath'] is not None:
                stage=owned(call['StagePath'],must_exist=False)
                check(not stage.exists(),call_name+' exact stage removed after released ownership')
        if case.startswith('environment-'):
            check(capture['CallerBefore']['State'] == observation['RequestedOSState'],name+' true requested OS state')
            if not case.endswith('-start-failure'):check('Email validation exit: 0;' in proof['Log'],name+' strict actual derivative validation logged')
            else:check(bool(emails[0]['ControlledSubstitution']),name+' disclosed OS invalid-image substitution')
        elif case in ('cancel-before-master','cancel-after-master'):
            target=master[0] if case == 'cancel-before-master' else emails[0]
            check(target['StagedExistsBeforeEntryCleanup'] is True and target['StagedBytes'] > 0 and capture['CancellationRequested'] is True,name+' controlled writer partial and cancellation observed')
            tree(observation,target['Result'],fixture,name)
        else:
            check(capture['FastCancelReached'] is True and capture['CancellationRequested'] is True,name+' token cancelled after real staged inspection')
            inspections=[call for call in calls if 'dump_data_utf8' in call['RequestedArguments'] and Path(call['RequestedArguments'][0]).name == 'master.pdf']
            check(len(inspections) == 1,name+' genuine successful staged master inspection before final move refusal')
        audited_cases.append(dict(shell=shell,label=case,exit_code=code,published_finals=count,source_foreign_preserved=True,
                                 scope='Original entry/runtime with disclosed copied-helper controlled fault seams'))
    check(len(operands) == 13,prefix+'exact thirteen retained validated final operands')
    return shell,operands

def audit_pdf(executable, row):
    path = owned(row['Path'])
    before = sha(path)
    check(before == row['SHA256'].lower(), 'Explicit retained PDF hash: '+label(path))
    result = read_pdftk(executable,path)
    check(result['exit_code'] == 0 and not result['stderr'], 'Fresh actual PDFtk read success: '+label(path))
    count = re.findall(r'^NumberOfPages:[ \t]*([0-9]+)[ \t]*\r?$',result['stdout'],re.M)
    check(count == [str(len(row['IDs']))] and len(row['IDs']) > 0, 'Fresh strict expected PDFtk page count: '+label(path))
    pages = []
    with pdfium.PdfDocument(path) as document:
        for index in range(len(document)):
            with closing(document[index]) as page:
                with closing(page.get_textpage()) as text:
                    ids = re.findall(r'T03-[0-9]{2}-P[0-9]{2}',text.get_text_range())
                pages.append(dict(identifiers=ids,rotation_degrees=page.get_rotation(),
                    rotation_quarter_turns=int(pdfium.raw.FPDFPage_GetRotation(page)),size_points=list(page.get_size())))
    check(len(pages) == len(row['IDs']) == len(row['Rotations']) == len(row['Sizes']), 'Fresh PDFium exact page inventory: '+label(path))
    for index,page in enumerate(pages):
        if index >= len(row['IDs']):
            continue
        check(page['identifiers'] == [row['IDs'][index]] and page['rotation_degrees'] == row['Rotations'][index]
              and page['rotation_quarter_turns']*90 == row['Rotations'][index]
              and page['size_points'] == row['Sizes'][index], 'Fresh synthetic ID/rotation/dimensions: '+label(path)+'/'+str(index+1))
    check(sha(path) == before, 'Read-only audit preserved PDF bytes: '+label(path))
    fresh.append(dict(path=label(path),sha256=before,bytes=path.stat().st_size,purpose=row['Purpose'],
        pdftk_command=['<approved PDFtk2.02>',label(path),'dump_data_utf8','output','-','dont_ask'],pdftk=result,
        pdfium=dict(page_count=len(pages),pages=pages,python=sys.version.split()[0],pypdfium2=str(pdfium.PYPDFIUM_INFO),pdfium=str(pdfium.PDFIUM_INFO))))

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--commit',required=True)
    parser.add_argument('--observations',action='append',required=True)
    parser.add_argument('--prior-attempt')
    parser.add_argument('--prior-failed-report')
    parser.add_argument('--output',required=True)
    args=parser.parse_args()
    history=[]
    if bool(args.prior_attempt) != bool(args.prior_failed_report):
        raise ValueError('Audit history requires both exact prior attempt and failed report operands.')
    if args.prior_attempt:
        attempt=load(args.prior_attempt);prior=load(args.prior_failed_report)
        if attempt['ExitCode'] != 1 or attempt['ReportSHA256'] != sha(args.prior_failed_report) or prior['result'] != 'fail':
            raise ValueError('Prior audit preparation failure bytes do not match their original attempt receipt.')
        for stream in ('Stdout','Stderr'):
            stream_path=bind(REPO/attempt[stream])
            if sha(stream_path) != attempt[stream+'SHA256']:
                raise ValueError('Prior captured audit stream bytes changed.')
        history.append(dict(kind='Audit preparation failure; case-sensitive uppercase receipt hash comparison',
            attempt=label(args.prior_attempt),attempt_sha256=sha(args.prior_attempt),
            failed_report=label(args.prior_failed_report),failed_report_sha256=sha(args.prior_failed_report),
            original_output_label=attempt['Report'],original_exit_code=attempt['ExitCode'],
            checks=prior['check_count'],fresh_pdf_reads=len(prior['fresh_readonly_pdf_inspections']),findings=prior['findings'],
            disposition='Preserved exact failed receipt bytes at this relocated path; normalized only audit hash letter case, then reran read-only audit.'))
    output=owned(args.output,must_exist=False)
    if output.exists():
        raise ValueError('Never overwrite an earlier independent audit receipt.')
    current=subprocess.check_output(['git','-C',str(REPO),'rev-parse','HEAD'],text=True).strip()
    dirty=subprocess.check_output(['git','-C',str(REPO),'status','--porcelain=v1'],text=True)
    if not re.fullmatch(r'[0-9a-f]{40}',args.commit) or current != args.commit or dirty:
        raise ValueError('Fresh retained native audit requires exact clean implementation C1.')
    if os.name != 'nt' or ctypes.sizeof(ctypes.c_void_p) != 8 or ctypes.windll.shell32.IsUserAnAdmin():
        raise ValueError('Actual standard-user Windows x64 audit required.')
    check(sys.version.split()[0] == '3.12.14' and str(pdfium.PYPDFIUM_INFO) == '5.13.0'
          and str(pdfium.PDFIUM_INFO) == '153.0.7999.0', 'Pinned development Python/PDFium versions')
    cache=load(WORK/'T15-cache-verification.json')
    oracle_runtime=cache['development_oracle_runtime']
    check(sha(sys.executable) == oracle_runtime['python_sha256'],'Pinned actual development Python executable bytes')
    oracle_dll=Path(oracle_runtime['pdfium_dll_path'].replace('<USERPROFILE>',str(Path.home())))
    check(sha(oracle_dll) == oracle_runtime['pdfium_dll_sha256'],'Pinned actual independent PDFium DLL bytes')
    pdf_dependency=next(item for item in cache['dependencies'] if item['dependency'] == 'PDFtk')
    root=pdf_dependency['cache_root'].replace('<USERPROFILE>',str(Path.home()))
    root=os.path.expandvars(root)
    selected=next(item for item in pdf_dependency['selected_files'] if item['relative_path'].endswith('/pdftk.exe'))
    executable=Path(root)/selected['relative_path']
    check(sha(executable) == selected['sha256'], 'Approved actual PDFtk executable bytes')
    gs_dependency=next(item for item in cache['dependencies'] if item['dependency'] == 'Ghostscript')
    gs_root=Path(os.path.expandvars(gs_dependency['cache_root'].replace('<USERPROFILE>',str(Path.home()))))
    gs_selected=next(item for item in gs_dependency['selected_files'] if item['relative_path'].endswith('/gswin64c.exe'))
    gs=gs_root/gs_selected['relative_path']
    engine_pins={}
    for dependency,base in ((pdf_dependency,Path(root)),(gs_dependency,gs_root)):
        source=REPO/dependency['source_receipt']
        check(sha(source) == dependency['source_receipt_sha256'],'Approved acquisition receipt bytes: '+dependency['dependency'])
        acquisition=json.loads(source.read_text(encoding='utf-8-sig'))
        approved=acquisition['extracted_files'] if dependency['dependency'] == 'PDFtk' else acquisition['ghostscript_extraction']['selected_files']
        for item in dependency['selected_files']:
            if Path(item['relative_path']).name in ('pdftk.exe','libiconv2.dll','gswin64c.exe','gsdll64.dll'):
                pins=[row['sha256'] for row in approved if Path(row['relative_path']).name == Path(item['relative_path']).name]
                check(pins == [item['sha256']],'Selected file agrees with original approved acquisition: '+Path(item['relative_path']).name)
                check(sha(base/item['relative_path']) == item['sha256'],'Approved selected engine bytes: '+Path(item['relative_path']).name)
                engine_pins[Path(item['relative_path']).name]=item['sha256']
    shell_versions=[];operands=[]
    for path in args.observations:
        shell,rows=audit_observations(path,args.commit,executable,gs,engine_pins)
        shell_versions.append(shell);operands.extend(rows)
    check(len(shell_versions) == 2 and set(shell_versions) == {'5.1.26100.9444','7.6.6'}, 'Both required actual native observation shells exactly once')
    check(len(operands) == len({row['Path'] for row in operands}) == 26,'Explicit exact 26 retained final PDF inspection inventory')
    for row in operands:
        audit_pdf(executable,row)
    check(subprocess.check_output(['git','-C',str(REPO),'rev-parse','HEAD'],text=True).strip() == args.commit and
          not subprocess.check_output(['git','-C',str(REPO),'status','--porcelain=v1'],text=True),
          'Exact clean C1 remains unchanged after read-only audit')
    command=['<approved development Python 3.12.14>','-B',label(__file__),'--commit',args.commit]
    for observation in args.observations:command.extend(('--observations',label(observation)))
    if args.prior_attempt:command.extend(('--prior-attempt',label(args.prior_attempt),'--prior-failed-report',label(args.prior_failed_report)))
    command.extend(('--output',label(output)))
    report=dict(schema_version=1,task='T15',implementation_commit=args.commit,observed_at_utc=datetime.now(timezone.utc).isoformat(),
        audit_kind='Independent fault native observations, preservation and fresh retained PDF inspections',partial=False,
        result='pass' if not findings else 'fail',check_count=len(checks),checks=checks,findings=findings,
        command=command,audit_history=history,raw_receipts=receipts,case_audits=audited_cases,fresh_readonly_pdf_inspections=fresh,audit_script_sha256=sha(__file__),audit_python_sha256=sha(sys.executable),
        reviewed_source_sha256={file:sha(REPO/file) for file in ('WinPDFMerge.ps1','src/WinPDFMerge.Helpers.ps1','tests/faults/FaultRecovery.Native.Tests.ps1','tests/native/FakeNative.cs')},
        approved_engine_sha256=engine_pins,pdftk_sha256=sha(executable),limitations=['No suite/application rerun. Actual read-only PDFtk/PDFium inspections of explicitly supplied synthetic retained PDFs only.',
        'This native-focused receipt cannot by itself close AC035/36/37 or T15; full suite XML/control/archive semantic review remains separate.',
        'Historical tree termination and unrelated sentinel liveness assertions bind original observations/test source; audit does not repeat their concurrent scheduling.',
        'No actual disk-exhaustion, hard crash, physical Ctrl+C, Explorer, broad feature/signature/fidelity or release acceptance claim.'])
    payload=json.dumps(report,indent=2,ensure_ascii=True)+'\n'
    if re.search(r'[A-Za-z]:[\\/]',payload):
        raise ValueError('Audit payload contains an absolute private path.')
    output.parent.mkdir(parents=True,exist_ok=True)
    with output.open('x',encoding='utf-8') as file:
        file.write(payload)
    print(json.dumps(dict(result=report['result'],partial=report['partial'],checks=len(checks),reads=len(fresh),report=label(output),sha256=sha(output))))
    return 0 if not findings else 1

if __name__ == '__main__':
    raise SystemExit(main())
