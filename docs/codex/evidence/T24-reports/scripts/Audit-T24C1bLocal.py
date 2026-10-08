"""Read-only local sanitized receipt/invocation audit; no tests are invoked."""
import hashlib
import json
import pathlib
import re
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timezone

repo = next(parent for parent in pathlib.Path(__file__).resolve().parents if (parent/'.git').exists())
commit = 'e626e45a5ba375456f23b506f0ded7ca7d68f1e3'
base = repo / 'tests/.work/T24-clean-driver' / commit
counts = ('passed','failed','failed_blocks','failed_containers','skipped','not_run','inconclusive','total')
allowed_keys = {'schema_version','commit_under_test','tier','evidence_class','runner_label','shell','shell_version','shell_edition','process_64_bit','pester_version','source_unchanged','runner_error_present','manual_desktop_acceptance','result','accepted','nunit_discovery_errors',*counts}
classes = {'Unit':'unit-controlled','Static':'static','Launcher':'controlled-launcher','NativeRunner':'controlled-native-process','ToolInvocation':'unit-controlled','PublicDocs':'documentation','NativeFixture':'windows-native-integration','SourceDiscovery':'windows-native-integration','CiNativeSmoke':'windows-native-integration'}
checks = 0
def check(condition,reason):
    global checks
    checks += 1
    if not condition: raise AssertionError(reason)
def read_json(path):
    return json.loads(path.read_text(encoding='utf-8-sig'))
def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()
groups = []
for shell in ('PS51','PS7'):
    for group in ('unit','native'):
        label = shell+'-'+group
        path = base / label
        job = read_json(path/'job.json')
        invocation = read_json(base/(label+'.invocation.json'))
        check(invocation['commit'] == commit and invocation['shell'] == shell and invocation['group'] == group and invocation['exit_code'] == 0, 'Actual local invocation did not pass from the expected commit')
        check((base/(label+'.stderr.txt')).read_bytes() == b'', 'Local driver stderr is nonempty')
        check(job['schema_version'] == 1 and job['commit_under_test'] == commit and job['group'] == group and job['shell'] == shell and job['runner_label'] == 'windows-local', 'Local job context mismatch')
        check(job['result'] == 'pass' and job['source_unchanged'] is True and job['administrator_token'] is False and job['manual_desktop_acceptance'] is False and job['failure_probe_requested'] is False and job['child_inherits_selected_shell_module_path'] is True, 'Local job source/privilege/scope mismatch')
        check(job['process_64_bit'] is True and job['shell_edition'] == ('Desktop' if shell == 'PS51' else 'Core') and job['shell_version'] == ('5.1.26100.9444' if shell == 'PS51' else '7.6.6') and job['execution_policy'] == 'RemoteSigned', 'Local shell/runtime mismatch')
        tiers = {t['tier']:t for t in job['tiers']}
        expected = {'Unit','Static','Launcher','NativeRunner','ToolInvocation','PublicDocs'} if group == 'unit' else {'NativeFixture','SourceDiscovery','CiNativeSmoke'}
        check(set(tiers) == expected, 'Local required tier inventory mismatch')
        record = {'label':label,'actual_invocation_exit_code':invocation['exit_code'],'invocation_sha256':digest(base/(label+'.invocation.json')),'job_sha256':digest(path/'job.json'),'elapsed_seconds':job['elapsed_seconds'],'shell_version':job['shell_version'],'administrator_token':False,'pairs':[],'totals':dict.fromkeys(counts,0)}
        for tier in sorted(tiers):
            sp = path/tier/'summary.json'
            xp = path/tier/'results.xml'
            s = read_json(sp)
            check(set(s) == allowed_keys and s['schema_version'] == 1, 'Local summary schema differs from sanitized allowlist')
            check(s['commit_under_test'] == commit and s['tier'] == tier and s['shell'] == shell and s['shell_version'] == job['shell_version'] and s['runner_label'] == 'windows-local' and s['evidence_class'] == classes[tier], 'Local summary source/runtime/evidence mismatch')
            check(s['source_unchanged'] is True and s['manual_desktop_acceptance'] is False and s['runner_error_present'] is False and s['process_64_bit'] is True and s['pester_version'] == '6.2.0', 'Local summary guard/pin/scope mismatch')
            check(s['accepted'] is True and s['result'] == 'pass' and s['passed'] == s['total'] > 0 and all(type(s[k]) is int and s[k] == 0 for k in counts if k not in ('passed','total')), 'Local summary does not pass with truthful zero bad counts')
            check(all(tiers[tier][k] == s[k] for k in counts) and tiers[tier]['result'] == 'pass' and tiers[tier]['process_exit_code'] == 0, 'Local job/summary/child exit mismatch')
            check(not re.search(r'(?i)(C:\\Users\\|synthetic-private|stack-trace|source_start|source_end)',sp.read_text(encoding='utf-8-sig')+xp.read_text(encoding='utf-8-sig')), 'Raw identity/path/diagnostic data found')
            raw_xml=xp.read_bytes()
            check(b'<!DOCTYPE' not in raw_xml and b'<!ENTITY' not in raw_xml, 'Unexpected XML DTD/entity')
            root=ET.fromstring(raw_xml)
            check(root.tag == 'test-results' and root.attrib['name'] == tier, 'Local XML root differs from tier')
            leaves=root.findall('.//test-case')
            states=Counter(c.attrib['result'] for c in leaves)
            check(len(leaves) == int(root.attrib['total']) == s['total'] and states['Success'] == s['total'] and len(states) == 1, 'Local XML leaves/count/state mismatch')
            check(all(int(root.attrib[k]) == 0 for k in ('errors','failures','not-run','inconclusive','ignored','skipped','invalid')), 'Local XML has nonzero bad counts')
            for node in root.iter():
                check(node.tag in {'test-results','test-suite','test-case','results'} and set(node.attrib) <= {'total','errors','failures','not-run','inconclusive','ignored','skipped','invalid','executed','result','success','time','asserts','type','name'} and not (node.text or '').strip() and not (node.tail or '').strip(), 'Unexpected XML field/text disclosed')
                if node.tag in ('test-suite','test-case'):
                    check(node.attrib['result'] == 'Success' and node.attrib['success'] == node.attrib['executed'] == 'True', 'Local XML suite/case execution mismatch')
                    check(re.fullmatch(('suite' if node.tag=='test-suite' else 'case')+r'-[1-9][0-9]*',node.attrib['name']) is not None, 'Original local test/suite name disclosed')
            record['pairs'].append({'tier':tier,'summary_sha256':digest(sp),'xml_sha256':digest(xp),'counts':{k:s[k] for k in counts}})
            for k in counts: record['totals'][k]+=s[k]
        check(record['totals']['total'] == (644 if group == 'unit' else 9), 'Local group execution count mismatch')
        if group == 'unit':
            st=read_json(path/'static.json')
            check(st['result']=='pass' and st['accepted'] is True and st['commit_under_test']==commit and st['shell']==shell and st['analyzer_version']=='1.25.0' and st['files_checked']==st['parser_passed']==st['analyzer_passed']==64, 'Local maintained static scope/count mismatch')
            check(all(st[k]==0 for k in ('parser_failed','parser_errors','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','source_guard_failed','checkpoint_guard_failed')), 'Local static gate has a bad count')
            check(st['advisory_errors']==0 and st['advisory_warnings']==325 and st['advisory_information']==175, 'Local advisory counts not retained')
            record['static_sha256']=digest(path/'static.json')
        groups.append(record)
totals={k:sum(g['totals'][k] for g in groups) for k in counts}
check(totals['passed']==totals['total']==1306 and all(totals[k]==0 for k in counts if k not in ('passed','total')), 'Local overall counts differ from expected observed executions')
audit={'task':'T24','audit':'C1b-local-driver-independent-invocation-sanitized-source-count-privacy-checks','observed_at_utc':datetime.now(timezone.utc).isoformat(),'result':'pass','commit_under_test':commit,'checks':checks,'groups':groups,'report_pairs':18,'totals':totals,'limitations':['This audit inspects retained actual invocation receipts and sanitized artifacts; it invokes no tests and is not added to application execution counts.','Controlled/helper/native/documentation classes remain distinct.','Local standard-user CLI operation is not Explorer drag-and-drop or manual desktop acceptance.','Hosted preparation failures and deliberate probe failures are retained in separate audits.']}
(repo/'tests/.work/T24-C1b-local-audit.json').write_text(json.dumps(audit,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:audit[k] for k in ('result','checks','report_pairs','totals')},indent=2))
