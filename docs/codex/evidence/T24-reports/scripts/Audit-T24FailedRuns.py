"""Independent read-only inspection of preserved GitHub artifact ZIP bytes."""
import hashlib
import json
import pathlib
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timezone

repo = next(parent for parent in pathlib.Path(__file__).resolve().parents if (parent/'.git').exists())
base = repo / 'tests/.work/T24-hosted'
counts = ('passed', 'failed', 'failed_blocks', 'failed_containers', 'skipped', 'not_run', 'inconclusive', 'total')
summary_keys = {'schema_version', 'commit_under_test', 'tier', 'evidence_class', 'runner_label', 'shell', 'shell_version', 'shell_edition', 'process_64_bit', 'pester_version', 'source_unchanged', 'runner_error_present', 'manual_desktop_acceptance', 'result', 'accepted', 'nunit_discovery_errors', *counts}
classes = {'Unit':'unit-controlled', 'Static':'static', 'Launcher':'controlled-launcher', 'NativeRunner':'controlled-native-process', 'ToolInvocation':'unit-controlled', 'PublicDocs':'documentation', 'NativeFixture':'windows-native-integration', 'SourceDiscovery':'windows-native-integration', 'PdftkPaths':'windows-native-integration', 'GhostscriptPaths':'windows-native-integration'}
xml_attrs = {'total','errors','failures','not-run','inconclusive','ignored','skipped','invalid','executed','result','success','time','asserts','type','name'}
checks = 0
def check(condition, reason):
    global checks
    checks += 1
    if not condition:
        raise AssertionError(reason)

def sha(data):
    return hashlib.sha256(data).hexdigest()

runs = []
for rid in ('37808743041', '37808743156'):
    path = base / rid
    raw_metadata = (path / 'run.json').read_bytes()
    metadata = json.loads(raw_metadata)
    check(metadata['id'] == int(rid) and metadata['status'] == 'completed' and metadata['conclusion'] == 'failure', 'Expected actual completed failed workflow')
    check(len(metadata['artifacts']) == 4 and len(metadata['jobs']) == 4, 'Expected four shell/group jobs/artifacts')
    run = {'run_id':int(rid), 'run_url':metadata['html_url'], 'event':metadata['event'], 'head_sha':metadata['head_sha'], 'conclusion':metadata['conclusion'], 'metadata_sha256':sha(raw_metadata), 'artifacts':[], 'totals':dict.fromkeys(counts, 0)}
    commits = set()
    for artifact in metadata['artifacts']:
        name = artifact['name']
        group, shell = re.fullmatch(r'windows-(unit|native)-(PS51|PS7)-1', name).groups()
        raw_zip = (path / (name + '.zip')).read_bytes()
        digest = sha(raw_zip)
        check('sha256:' + digest == artifact['digest'], 'Downloaded ZIP digest differs from GitHub digest')
        check(len(raw_zip) == artifact['size_in_bytes'], 'Downloaded ZIP size differs from GitHub metadata')
        files_metadata = {f['path']: f for f in artifact['files']}
        record = {'name':name, 'artifact_id':artifact['id'], 'zip_sha256':digest, 'github_digest_match':True, 'file_count':len(files_metadata), 'pairs':[], 'totals':dict.fromkeys(counts,0)}
        with zipfile.ZipFile(path / (name + '.zip')) as z:
            check(set(z.namelist()) == set(files_metadata), 'ZIP inventory differs from preserved download metadata')
            contents = {n:z.read(n) for n in z.namelist()}
        for filename, data in contents.items():
            check(filename.endswith(('.json','.xml')) and not filename.startswith(('/', '\\')) and '..' not in pathlib.PurePosixPath(filename).parts, 'Unexpected artifact file type/path')
            check(len(data) == files_metadata[filename]['bytes'] and sha(data) == files_metadata[filename]['sha256'], 'Artifact file digest differs from recorded ZIP metadata')
            check(data == (path / name / filename).read_bytes(), 'Extracted artifact differs from ZIP bytes')
            text = data.decode('utf-8-sig')
            check(not re.search(r'(?i)(C:\\Users\\|D:\\a\\|synthetic-private|stack-trace|source_start|source_end|private_extra|runner_error"\s*:)', text), 'Private path/identity/raw diagnostic field found')
        job = json.loads(contents['job.json'])
        check(job['schema_version'] == 1 and job['group'] == group and job['shell'] == shell and job['runner_label'] == 'windows-2025', 'Job context mismatch')
        check(job['result'] == 'fail' and job['source_unchanged'] is True and job['manual_desktop_acceptance'] is False and job['administrator_token'] is True, 'Hosted failure/privilege/manual context mismatch')
        check(job['process_64_bit'] is True and job['shell_edition'] == ('Desktop' if shell == 'PS51' else 'Core') and (re.fullmatch(r'5\.1\.\d+\.\d+',job['shell_version']) if shell == 'PS51' else job['shell_version'] == '7.6.6'), 'Selected shell mismatch')
        commits.add(job['commit_under_test'])
        if metadata['event'] == 'push':
            check(job['commit_under_test'] == metadata['head_sha'], 'Push job did not test its event commit')
        dep = json.loads(contents['dependencies.json'])
        check(dep['result'] == 'pass' and dep['group'] == group and dep['requested_shell'] == shell and dep['runner_label'] == 'windows-2025', 'Dependency context mismatch')
        check(all(d['verified'] is True and re.fullmatch('[0-9a-f]{64}',d['sha256']) for d in dep['downloads']), 'Dependency verification receipt incomplete')
        check(all(v is False for v in dep['boundaries'].values()), 'Unexpected installation/security/runtime-network boundary')
        tiers = {tier['tier']: tier for tier in job['tiers']}
        pairs = sorted(n[:-len('/summary.json')] for n in contents if n.endswith('/summary.json'))
        check(set(pairs) == set(tiers), 'Job tiers differ from paired result artifacts')
        for tier in pairs:
            raw_json = contents[tier+'/summary.json']
            raw_xml = contents[tier+'/results.xml']
            s = json.loads(raw_json)
            check(set(s) == summary_keys and s['schema_version'] == 1, 'Unexpected summary schema or data disclosure')
            check(s['commit_under_test'] == job['commit_under_test'] and s['tier'] == tier and s['shell'] == shell and s['shell_version'] == job['shell_version'] and s['runner_label'] == 'windows-2025', 'Summary context mismatch')
            check(s['pester_version'] == '6.2.0' and s['source_unchanged'] is True and s['runner_error_present'] is False and s['manual_desktop_acceptance'] is False, 'Summary pin/source/manual mismatch')
            check(s['evidence_class'] == classes[tier], 'Controlled/native evidence classification mismatch')
            check(all(type(s[k]) is int and 0 <= s[k] <= 2147483647 for k in counts), 'Invalid summary count type')
            check(s['passed']+s['failed']+s['skipped']+s['not_run']+s['inconclusive'] == s['total'], 'Summary count total mismatch')
            check(all(tiers[tier][k] == s[k] for k in counts) and tiers[tier]['result'] == s['result'], 'Job/summary count mismatch')
            check(tiers[tier]['process_exit_code'] == (0 if s['accepted'] else 1), 'Observed child exit differs from acceptance')
            check(b'<!DOCTYPE' not in raw_xml and b'<!ENTITY' not in raw_xml, 'Unexpected XML DTD/entity')
            root = ET.fromstring(raw_xml)
            check(root.tag == 'test-results' and root.attrib['name'] == tier, 'NUnit root mismatch')
            leaves = root.findall('.//test-case')
            leaf_states = Counter(c.attrib['result'] for c in leaves)
            check(len(leaves) == s['total']-s['not_run'] == int(root.attrib['total']), 'NUnit total/leaf mismatch')
            check(leaf_states['Success'] == s['passed'] and leaf_states['Failure'] == s['failed'] and leaf_states['Ignored'] == s['skipped'] and leaf_states['Inconclusive'] == s['inconclusive'], 'NUnit leaf state/summary mismatch')
            for attribute, key in (('failures','failed'),('not-run','not_run'),('inconclusive','inconclusive'),('skipped','skipped'),('errors','nunit_discovery_errors')):
                check(int(root.attrib[attribute]) == s[key], 'NUnit root counter/summary mismatch')
            suites = root.findall('.//test-suite')
            blocking_suite = any(x.attrib['result'] != 'Success' or x.attrib['success'] != 'True' or x.attrib['executed'] != 'True' for x in suites)
            expected_pass = s['total'] > 0 and s['passed'] == s['total'] and s['failed_blocks'] == s['failed_containers'] == s['nunit_discovery_errors'] == 0 and not blocking_suite
            check(s['accepted'] is expected_pass and s['result'] == ('pass' if expected_pass else 'fail'), 'Pass/fail acceptance does not match actual counts/states')
            failed_cases = []
            for node in root.iter():
                check(node.tag in {'test-results','test-suite','test-case','results','failure','reason','message'} and set(node.attrib) <= xml_attrs, 'Unexpected XML data field')
                check(not (node.tail or '').strip(), 'Unexpected XML tail text')
                if node.tag == 'message':
                    check(node.text == 'Details omitted from public CI artifact.', 'Original diagnostic text disclosed')
                else:
                    check(not (node.text or '').strip(), 'Unexpected XML data text')
                if node.tag == 'test-suite':
                    check(re.fullmatch(r'suite-[1-9][0-9]*',node.attrib['name']) is not None, 'Original suite name disclosed')
                if node.tag == 'test-case':
                    check(re.fullmatch(r'case-[1-9][0-9]*',node.attrib['name']) is not None, 'Original case name disclosed')
                    check(node.attrib['success'] == ('True' if node.attrib['result'] == 'Success' else 'False') and node.attrib['executed'] == ('False' if node.attrib['result'] == 'Ignored' else 'True'), 'Leaf execution state mismatch')
                    if node.attrib['result'] == 'Failure': failed_cases.append(node.attrib['name'])
            check(len(failed_cases) == s['failed'], 'Failure details count mismatch')
            record['pairs'].append({'tier':tier,'result':s['result'],'counts':{k:s[k] for k in counts},'failed_cases':failed_cases,'summary_sha256':sha(raw_json),'xml_sha256':sha(raw_xml)})
            for k in counts:
                record['totals'][k] += s[k]
                run['totals'][k] += s[k]
        if group == 'unit':
            static = json.loads(contents['static.json'])
            check(static['result'] == 'pass' and static['accepted'] is True and static['commit_under_test'] == job['commit_under_test'] and static['shell'] == shell and static['manual_desktop_acceptance'] is False, 'Static receipt context/state mismatch')
            check(static['files_checked'] == static['parser_passed'] == static['analyzer_passed'] > 0 and static['analyzer_version'] == '1.25.0', 'Static pass count/pin mismatch')
        check(any(t['result'] == 'fail' and t['process_exit_code'] == 1 for t in tiers.values()), 'Failed hosted job lacks observed failed child')
        record['job_result']='fail'
        run['artifacts'].append(record)
    check(len(commits) == 1, 'Run shells/groups disagree on tested commit')
    run['reported_tested_commit'] = next(iter(commits))
    runs.append(run)

audit = {'task':'T24','audit':'preserved-failed-hosted-runs-independent-ZIP-and-sanitized-receipt-checks','observed_at_utc':datetime.now(timezone.utc).isoformat(),'result':'pass','checks':checks,'workflow_runs':2,'artifact_zips':8,'report_pairs':sum(len(a['pairs']) for r in runs for a in r['artifacts']),'runs':runs,'limitations':['Audit establishes downloaded artifact integrity/schema/count/state/privacy consistency, not the omitted underlying exception text.','PR event metadata identifies source head; its consistently reported tested commit is the GitHub synthetic merge checkout selected by github.sha.','Both audited workflows actually failed; these are preparation failures, not final AC054 passing evidence.','Hosted runner uses an administrator token and is not Windows desktop/Explorer acceptance.','Native case-12 and PublicDocs case-13 sanitized failures retain no original diagnostic; root task separately preserves the investigated cause.']}
(repo/'tests/.work/T24-failed-run-audit.json').write_text(json.dumps(audit,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:audit[k] for k in ('result','checks','workflow_runs','artifact_zips','report_pairs')},indent=2))
print(json.dumps([{'id':r['run_id'],'totals':r['totals'],'reported_tested_commit':r['reported_tested_commit']} for r in runs],indent=2))
