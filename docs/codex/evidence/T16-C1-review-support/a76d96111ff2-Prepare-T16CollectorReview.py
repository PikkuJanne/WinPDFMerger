"""Adapt the retained independently authored reviewer, without editing originals."""
import ast
from pathlib import Path

root=Path(__file__).resolve().parent
text=(root/'Review-T15Collector.py').read_text(encoding='utf-8')
text=text.replace('53d0923c95a86ae6a44bc89bab51cac6786c1e32','26ac1b73e3733a23099de53d944e00e4ee412982')
text=text.replace('T15','T16').replace('t15','t16')
text=text.replace('317','335').replace('1118','1072').replace('559','536').replace('114','92')
text=text.replace('==34','==28').replace('Exactly 34','Exactly 28').replace('34 clean','28 clean').replace('==17','==14')
text=text.replace("'CleanReports':34","'CleanReports':28")
text=text.replace('len(waivers)==9','len(waivers)==3').replace('Exactly nine literal','Exactly three literal')
text=text.replace('T16-collector-check-f66e06ad6b224f088937891419321432','T16-collector-check-0ae1a1ba50214a7f9afd0371dc9cfbb9')
text=text.replace("manifest['prior_native_schema_primitives_sha256']==file_sha(WORK/'Collect-T14Evidence.py')",
                  "manifest['native_schema_primitives_sha256']==file_sha(WORK/'Collect-T14Evidence.py') and manifest['prior_collector_primitives_sha256']==file_sha(WORK/'Collect-T15Evidence.py')")
text=text.replace("for name in ('Collect-T14Evidence.py','Collect-T09C3Evidence.py')",
                  "for name in ('Collect-T15Evidence.py','Collect-T14Evidence.py','Collect-T09C3Evidence.py')")
old="""    check(payload==(WORK/name).read_bytes(),'Standalone original receipt byte preservation: '+name)
    collector.privacy_gate(payload,name)"""
new="""    raw=(WORK/name).read_bytes()
    expected=canonical(independent_value(json.loads(raw.decode('utf-8-sig')),collector.prefixes))
    check(payload==expected,'Standalone original receipt canonical/privacy-only transformation: '+name)
    collector.privacy_gate(payload,name)"""
assert old in text
text=text.replace(old,new)
old="""check(execution['CommitUnderCheck']==C1 and execution['DirtyWorktree'] is False and execution['ExitCode']==0 and
      execution['CollectorSHA256']==source_before,'Original executed collector check proof clean/source/exit binding')"""
new="""check(execution['CommitUnderTest']==C1 and execution['DirtyWorktreeBefore'] is False and execution['DirtyWorktreeAfter'] is False and execution['ExitCode']==0 and
      execution['CollectorSourceSHA256']==source_before and execution['PublicWriteRequested'] is False and execution['ApplicationOrNativeTestsExecuted'] is False,
      'Original executed collector check proof clean/source/exit binding')
check(file_sha(proof_root/'collector-source.py')==source_before,'Original collector check exact source snapshot')
check(file_sha(WORK/'Invoke-T16CollectorCheck.py')==execution['LauncherSourceSHA256'],'Original collector check launcher source binding')"""
assert old in text
text=text.replace(old,new)
start=text.index("native_path=WORK/'T16-C1-native-audit.json'")
end=text.index("check(not captured['destination'].exists()",start)
text=text[:start]+"""native_path=WORK/'T16-C1-native-audit.json';native=load(native_path)
check(native['CommitUnderTest']==C1 and native['Result']=='pass' and native['Partial'] is False and native['CheckCount']==2273
      and native['CaseCount']==len(native['Cases'])==18 and native['FreshFinalReads']==len(native['FreshReads'])==28,
      'Separate actual native audit receipt; reviewer authored and executed that audit, not a second independent auditor')
check(len(native['AuditHistory'])==1 and native['AuditHistory'][0]['ExitCode']==1 and native['AuditHistory'][0]['FreshFinalReads']==28
      and native['AuditHistory'][0]['FailedChecks']==16,'Native audit CRLF-comparison preparation failure honestly separate')
for binding in native['RawBindings']:check(file_sha(REPO/binding['path'])==binding['sha256'],'Original independently audited byte binding still intact: '+binding['path'])
for binding in native['RunSnapshotBindings']:
    jobs=load(REPO/binding['source']);row=next(job for job in jobs if job['tier']=='ParametersNative')
    digest=sha(json.dumps(row,sort_keys=True,separators=(',',':')).encode())
    check(digest==binding['native_row_sha256'],'Native command row unchanged in completed index: '+binding['shell'])
support=json.loads(payloads['retained-support-bindings.json'])
check(support==collector.support,'Exact supporting original inventory payload')
for binding in support:
    path=(WORK/binding['source_relative_path']).resolve()
    check(path.is_relative_to(WORK) and path.is_file() and file_sha(path)==binding['raw_sha256'] and path.stat().st_size==binding['bytes'],
          'Each retained original support file remains exact: '+binding['source_relative_path'])
preparation=results['collector_preparation_history']
check(len(preparation)==1 and preparation[0]['exit_code']==1 and preparation[0]['acceptance_counts_available'] is False,
      'Collector preparation failure is separate and never added to clean counts')
for history_root in WORK.glob('T16-collector-check-*'):
    original_execution=load(history_root/'execution.json')
    if original_execution['ExitCode']:
        check(file_sha(history_root/'collector-source.py')==original_execution['CollectorSourceSHA256']
              and file_sha(history_root/'stdout.txt')==original_execution['StdoutSHA256']
              and file_sha(history_root/'stderr.txt')==original_execution['StderrSHA256']
              and original_execution['CleanReports'] is None and original_execution['TotalPassed'] is None,
              'Exact failed collector source/captures and unavailable counts retained')
whitespace_details=[]
for waiver in waivers:
    name=Path(waiver['file']).name;payload=payloads[name];lines=payload.decode('utf-8-sig').splitlines()
    trailing=[index+1 for index,line in enumerate(lines) if re.search(r'[ \\t]+$',line)]
    check(len(trailing)==waiver['trailing_whitespace_lines'],'Exact literal whitespace line count: '+name)
    whitespace_details.append(dict(File=waiver['file'],TrailingWhitespaceLines=trailing,RawTailSHA256=sha(payload[-128:]),
                                   BlankLinesAtEOF=max(0,len(payload.decode('utf-8-sig').splitlines())-len(payload.decode('utf-8-sig').rstrip('\\r\\n').splitlines()))))
"""+text[end:]
text=text.replace("'Result':native['result'],'Checks':native['check_count'],'FreshPDFReads':26,'Cases':28",
                  "'Result':native['Result'],'Checks':native['CheckCount'],'FreshPDFReads':28,'Cases':18")
text=text.replace("'RawCleanSummaryXMLBindings':raw_clean_files,'ReviewScriptSHA256':file_sha(__file__),",
                  "'RawCleanSummaryXMLBindings':raw_clean_files,'SupportingOriginalFiles':len(support),'LiteralWhitespaceDetails':whitespace_details,'ReviewScriptSHA256':file_sha(__file__),")
text=text.replace('no failing-source byte snapshot or fabricated historical XML/summary is claimed.',
                  'both failing and passing collector source snapshots are retained; no absent historical XML/summary is fabricated.')
text=text.replace('Reviewer authored the launch/cancellation adapter and four legacy receipt additions; this is an independent collector/evidence review, not a fresh independent review of those changes.',
                  'Reviewer authored the unchanged T15 adapter and two T16 legacy wrapper additions, and executed the separate native audit; this independently reviews another agent collector and planned evidence, not a second independent audit of those authored pieces.')
text=text.replace("'ReviewerScope':'Independent review", "'ReviewerScope':'Independent review")
assert "native['result']" not in text and "native['implementation_commit']" not in text
ast.parse(text)
destination=root/'Review-T16Collector.py'
assert not destination.exists()
destination.write_text(text,encoding='utf-8')
print('Prepared ignored Review-T16Collector.py; no execution or public write.')
