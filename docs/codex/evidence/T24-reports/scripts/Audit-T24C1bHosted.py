"""Independent read-only audit of final C1b normal/negative GitHub artifacts."""
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
classes = {'Unit':'unit-controlled', 'Static':'static', 'Launcher':'controlled-launcher', 'NativeRunner':'controlled-native-process', 'ToolInvocation':'unit-controlled', 'PublicDocs':'documentation', 'NativeFixture':'windows-native-integration', 'SourceDiscovery':'windows-native-integration', 'CiNativeSmoke':'windows-native-integration', 'CiFailureProbe':'ci-controlled-deliberate-failure'}
xml_attrs = {'total','errors','failures','not-run','inconclusive','ignored','skipped','invalid','executed','result','success','time','asserts','type','name'}
checks = 0
def check(condition, reason):
    global checks
    checks += 1
    if not condition:
        raise AssertionError(reason)

def sha(data):
    return hashlib.sha256(data).hexdigest()

expected_head = 'e626e45a5ba375456f23b506f0ded7ca7d68f1e3'
expected_merge = '1c42402a05f4b18f0df35cf2a4225ffb928405df'
context_path = repo / 'tests/.work/T24-C1b-PRMerge-source-context.json'
context = json.loads(context_path.read_text(encoding='utf-8-sig'))
pins = {d['id']:d for d in json.loads((repo/'tests/ci-dependencies.json').read_text(encoding='utf-8-sig'))['dependencies']}
check(context['head']['sha'] == expected_head and context['merge']['sha'] == expected_merge and context['tree_equal'] is True and context['expected_head_is_merge_parent'] is True, 'PR merge source context mismatch')
check(context['head']['tree'] == context['merge']['tree'] == 'a3c2b9045be56d7ff34650aac92275eebcde031d' and expected_head in context['merge']['parents'], 'PR merge does not match reviewed source tree')
runs = []
for rid in ('37809847683', '37809856744', '37809951854'):
    path = base / rid
    raw_metadata = (path / 'run.json').read_bytes()
    metadata = json.loads(raw_metadata)
    negative = rid == '37809951854'
    expected_conclusion = 'failure' if negative else 'success'
    expected_event = 'pull_request' if rid == '37809856744' else 'push'
    expected_source = expected_merge if expected_event == 'pull_request' else expected_head
    check(metadata['id'] == int(rid) and metadata['status'] == 'completed' and metadata['conclusion'] == expected_conclusion and metadata['event'] == expected_event and metadata['head_sha'] == expected_head, 'Expected completed workflow/source/event mismatch')
    check(metadata['head_branch'] == ('codex/t24-ci-failure-probe' if negative else 'codex/v1.0.0-readiness'), 'Failure probe source branch mismatch')
    check(len(metadata['artifacts']) == 4 and len(metadata['jobs']) == 4, 'Expected four shell/group jobs/artifacts')
    run = {'run_id':int(rid), 'run_url':metadata['html_url'], 'event':metadata['event'], 'head_sha':metadata['head_sha'], 'conclusion':metadata['conclusion'], 'negative_probe':negative, 'metadata_sha256':sha(raw_metadata), 'artifacts':[], 'totals':dict.fromkeys(counts, 0)}
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
        job_failed = negative and group == 'unit'
        check(job['result'] == ('fail' if job_failed else 'pass') and job['source_unchanged'] is True and job['manual_desktop_acceptance'] is False and job['administrator_token'] is True, 'Hosted failure/privilege/manual context mismatch')
        check(job['commit_under_test'] == expected_source and job['failure_probe_requested'] is job_failed and job['child_inherits_selected_shell_module_path'] is True, 'Tested commit/probe/module context mismatch')
        job_metadata = [j for j in metadata['jobs'] if j['name'] == group+' / '+shell]
        check(len(job_metadata) == 1 and job_metadata[0]['conclusion'] == ('failure' if job_failed else 'success') and job_metadata[0]['labels'] == ['windows-2025'], 'GitHub job conclusion/runner differs from actual receipts')
        check(job['process_64_bit'] is True and job['shell_edition'] == ('Desktop' if shell == 'PS51' else 'Core') and (re.fullmatch(r'5\.1\.\d+\.\d+',job['shell_version']) if shell == 'PS51' else job['shell_version'] == '7.6.6'), 'Selected shell mismatch')
        commits.add(job['commit_under_test'])
        if metadata['event'] == 'push':
            check(job['commit_under_test'] == metadata['head_sha'], 'Push job did not test its event commit')
        dep = json.loads(contents['dependencies.json'])
        check(dep['result'] == 'pass' and dep['group'] == group and dep['requested_shell'] == shell and dep['runner_label'] == 'windows-2025', 'Dependency context mismatch')
        check(all(d['verified'] is True and re.fullmatch('[0-9a-f]{64}',d['sha256']) for d in dep['downloads']), 'Dependency verification receipt incomplete')
        expected_downloads = {id for id,pin in pins.items() if pin['group'] in ('all',group) and pin['shell'] in ('all',shell)}
        check({d['id'] for d in dep['downloads']} == expected_downloads and len(dep['downloads']) == len(expected_downloads), 'Missing, duplicate or unexpected dependency download')
        for download in dep['downloads']:
            pin = pins[download['id']]
            check(download['version'] == pin['version'] and download['url'] == pin['url'] and download['sha256'] == pin['sha256'], 'Dependency version/source/archive digest differs from reviewed CI manifest')
            expected_selected = {f['path']:f['sha256'] for f in pin['files']}
            check({f['path']:f['sha256'] for f in download['selected_files']} == expected_selected, 'Selected dependency file digest/inventory differs from reviewed CI manifest')
        check(all(v is False for v in dep['boundaries'].values()), 'Unexpected installation/security/runtime-network boundary')
        tiers = {tier['tier']: tier for tier in job['tiers']}
        pairs = sorted(n[:-len('/summary.json')] for n in contents if n.endswith('/summary.json'))
        check(set(pairs) == set(tiers), 'Job tiers differ from paired result artifacts')
        expected_tiers = {'Unit','Static','Launcher','NativeRunner','ToolInvocation','PublicDocs'} if group == 'unit' else {'NativeFixture','SourceDiscovery','CiNativeSmoke'}
        if job_failed: expected_tiers.add('CiFailureProbe')
        check(set(pairs) == expected_tiers, 'Unexpected or missing required CI tier')
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
            if tier == 'CiFailureProbe':
                check(negative and group == 'unit' and s['failed'] == 1 and s['passed'] == 0 and s['total'] == 1 and all(s[k] == 0 for k in ('failed_blocks','failed_containers','skipped','not_run','inconclusive')) and s['accepted'] is False, 'Deliberate failure probe did not fail with exact actual counts')
            else:
                check(s['accepted'] is True and s['passed'] == s['total'] and all(s[k] == 0 for k in ('failed','failed_blocks','failed_containers','skipped','not_run','inconclusive')), 'Normal CI tier did not pass cleanly')
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
        check(any(t['result'] == 'fail' and t['process_exit_code'] == 1 for t in tiers.values()) is job_failed, 'Job success/failure differs from child observations')
        record['job_result']='fail' if job_failed else 'pass'
        record['test_source_commit']=expected_source
        run['artifacts'].append(record)
    check(len(commits) == 1, 'Run shells/groups disagree on tested commit')
    run['reported_tested_commit'] = next(iter(commits))
    runs.append(run)

check(sum(len(a['pairs']) for r in runs for a in r['artifacts']) == 56, 'Expected final 56 artifact pairs')
check(all(r['totals']['passed'] == 1306 and r['totals']['failed'] == (2 if r['negative_probe'] else 0) and r['totals']['total'] == (1308 if r['negative_probe'] else 1306) for r in runs), 'Final observed workflow totals mismatch')
audit = {'task':'T24','audit':'C1b-final-hosted-normal-and-deliberate-failure-independent-ZIP-and-sanitized-receipt-checks','observed_at_utc':datetime.now(timezone.utc).isoformat(),'result':'pass','checks':checks,'workflow_runs':3,'artifact_zips':12,'report_pairs':56,'source_context':context,'source_context_sha256':sha(context_path.read_bytes()),'runs':runs,'normal_positive_totals':{k:sum(r['totals'][k] for r in runs if not r['negative_probe']) for k in counts},'negative_probe_totals':next(r['totals'] for r in runs if r['negative_probe']),'limitations':['Audit establishes downloaded artifact integrity/schema/count/state/privacy consistency; original diagnostic text is deliberately omitted.','Normal push and PR workflows actually pass; the negative workflow fails only its two deliberate one-test unit probes.','PR event head and tested synthetic merge are distinct commits with independently queried identical Git tree and expected parents.','Hosted Windows Server runner uses an administrator token and is not Windows desktop/Explorer/standard-user acceptance.','Native smoke inspects structural page totals with real PDFtk; it does not claim independent rendering or feature-preservation acceptance.','Earlier failed preparation runs are audited separately and excluded from the positive totals.']}
(repo/'tests/.work/T24-C1b-hosted-audit.json').write_text(json.dumps(audit,indent=2)+'\n',encoding='utf-8')
print(json.dumps({k:audit[k] for k in ('result','checks','workflow_runs','artifact_zips','report_pairs')},indent=2))
print(json.dumps([{'id':r['run_id'],'totals':r['totals'],'reported_tested_commit':r['reported_tested_commit']} for r in runs],indent=2))
