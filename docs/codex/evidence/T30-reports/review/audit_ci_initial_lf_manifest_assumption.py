"""Independent downloaded sanitized T30 CI report audit; does not run tests."""
from pathlib import Path
import datetime, hashlib, json, subprocess, xml.etree.ElementTree as ET
repo = Path.cwd().resolve()
root = repo / 'tests/.work/T30-ci-4d6a02d3d1f147f0b920dec292a1b145'
source = '1e4f2b79fb9a025d71d72e7cec9f566a7c11c930'
merge = '672b6f1bd1b02ccc91004b7ca1c01d585aea2e51'
sha = lambda data: hashlib.sha256(data).hexdigest()
read = lambda path: json.loads(Path(path).read_text(encoding='utf-8-sig'))
checks, issues = 0, []
def check(value, label):
    global checks
    checks += 1
    if not value:
        issues.append(label)
tree = subprocess.check_output(['git','rev-parse',source+'^{tree}'],cwd=repo).decode().strip()
api_bytes = subprocess.check_output(['gh','api','repos/PikkuJanne/WinPDFMerger/git/commits/'+merge],cwd=repo)
api = json.loads(api_bytes)
check(api['sha'] == merge and api['tree']['sha'] == tree, 'PR synthetic merge has exact C1 source tree')
check([item['sha'] for item in api['parents']] == ['e2451141217efdd00a1d49d72a04df054872dffc',source], 'PR synthetic merge actual parents')
manifest_bytes = subprocess.check_output(['git','show',source+':tests/ci-dependencies.json'],cwd=repo)
manifest = json.loads(manifest_bytes)
pins = {item['id']:item for item in manifest['dependencies']}
downloads = read(root / 'downloads.json')
check(len(downloads) == 2, 'Both actual CI artifact downloads')
events = []
for event, run, commit in [('push',37956550873,source),('pull_request',37956557899,merge)]:
    metadata = read(root / (event + '.json'))
    check(metadata['databaseId'] == run and metadata['headSha'] == source and metadata['event'] == event, event + 'actual platform run/source metadata')
    check(metadata['status'] == 'completed' and metadata['conclusion'] == 'success' and len(metadata['jobs']) == 4, event + 'four successful completed jobs')
    for job in metadata['jobs']:
        check(job['conclusion'] == 'success' and job['status'] == 'completed', event + 'platform job conclusion')
    receipt = [row for row in downloads if row['event'] == event]
    check(len(receipt) == 1 and receipt[0]['run'] == run and receipt[0]['exit_code'] == 0, event + 'actual gh download exit')
    artifacts = sorted((root / event).iterdir())
    check(len(artifacts) == 4 and all(item.is_dir() for item in artifacts), event + 'exact four downloaded artifact groups')
    jobs, total = [], 0
    for artifact in artifacts:
        job = read(artifact / 'job.json')
        prefix = event + '/' + artifact.name + ': '
        shell, group = job['shell'], job['group']
        tiers = ['NativeFixture','SourceDiscovery','CiNativeSmoke'] if group == 'native' else ['Unit','Static','Launcher','NativeRunner','ToolInvocation','PublicDocs','Version']
        check(job['commit_under_test'] == commit and job['runner_label'] == 'windows-2025', prefix + 'actual job source/runner')
        check(job['result'] == 'pass' and job['source_unchanged'] is True and job['manual_desktop_acceptance'] is False and job['failure_probe_requested'] is False, prefix + 'source/scope/result')
        check(job['process_64_bit'] is True and job['execution_policy'] == 'RemoteSigned' and job['administrator_token'] is True, prefix + 'actual host/token/policy facts')
        expected_version, expected_edition = ('5.1.26100.33438','Desktop') if shell == 'PS51' else ('7.6.6','Core')
        check(job['shell_version'] == expected_version and job['shell_edition'] == expected_edition and job['os_version'] == '10.0.26100.0', prefix + 'actual hosted shell/OS')
        check([row['tier'] for row in job['tiers']] == tiers, prefix + 'complete selected tier list')
        job_total = 0
        for row in job['tiers']:
            tier = row['tier']
            summary = read(artifact / tier / 'summary.json')
            xml = ET.parse(artifact / tier / 'results.xml').getroot()
            leaves = xml.findall('.//test-case')
            check(summary['commit_under_test'] == commit and summary['tier'] == tier and summary['shell'] == shell and summary['runner_label'] == 'windows-2025', prefix + tier + ' source identity')
            check(summary['shell_version'] == expected_version and summary['shell_edition'] == expected_edition and summary['process_64_bit'] is True and summary['pester_version'] == '6.2.0', prefix + tier + ' recorded exact host')
            check(summary['accepted'] is True and summary['result'] == row['result'] == 'pass' and summary['source_unchanged'] is True and summary['runner_error_present'] is False and summary['manual_desktop_acceptance'] is False, prefix + tier + ' actual scope/guard')
            check(row['process_exit_code'] == 0 and summary['passed'] == summary['total'] == row['passed'] == row['total'] == len(leaves) == int(xml.get('total')), prefix + tier + ' JSON/job/NUnit count equality')
            for field in ['failed','failed_blocks','failed_containers','skipped','not_run','inconclusive']:
                check(summary[field] == row[field] == 0, prefix + tier + ' ' + field)
            check(summary['nunit_discovery_errors'] == 0, prefix + tier + ' discovery errors')
            for field in ['errors','failures','not-run','inconclusive','ignored','skipped','invalid']:
                check(int(xml.get(field)) == 0, prefix + tier + ' NUnit ' + field)
            for leaf in leaves:
                check(leaf.get('result') == 'Success' and leaf.get('success') == 'True' and leaf.get('executed') == 'True', prefix + tier + ' successful executed leaf')
            job_total += summary['passed']
        dependency = read(artifact / 'dependencies.json')
        check(dependency['result'] == 'pass' and dependency['runner_environment'] == 'github-hosted' and dependency['runner_os'] == 'Windows' and dependency['manifest_sha256'] == sha(manifest_bytes), prefix + 'exact committed dependency manifest receipt')
        for item in dependency['downloads']:
            pin = pins[item['id']]
            check(item['version'] == pin['version'] and item['url'] == pin['url'] and item['sha256'] == pin['sha256'] and item['bytes'] == pin['bytes'] and item['verified'] is True, prefix + item['id'] + ' acquisition integrity receipt')
            check(item['selected_files'] == pin['files'], prefix + item['id'] + ' selected executable/module digest receipts')
        check(all(value is False for value in dependency['boundaries'].values()), prefix + 'no installer/elevation/persistent/runtime-network scope')
        if group == 'native':
            check({'pdftk','ghostscript','pester','innoextract'}.issubset({item['id'] for item in dependency['downloads']}), prefix + 'both genuine selected engine acquisition receipts')
        else:
            static = read(artifact / 'static.json')
            check(static['accepted'] is True and static['result'] == 'pass' and static['commit_under_test'] == commit and static['files_checked'] == static['parser_passed'] == static['analyzer_passed'] == 68, prefix + 'all maintained source static gate')
            for field in ['parser_failed','parser_errors','analyzer_failed','analyzer_not_run','skipped','selected_errors','selected_warnings','selected_information','selected_suppressions','source_guard_failed','checkpoint_guard_failed']:
                check(static[field] == 0, prefix + 'static ' + field)
        total += job_total
        jobs.append({'artifact':artifact.name,'group':group,'shell':shell,'source_commit':commit,'passed':job_total,'pairs':len(tiers),'native_smoke_scope':'Real pinned engine/helper structural smoke only; no PDF output/render/archive inspection by this audit' if group == 'native' else None})
    events.append({'event':event,'run':run,'platform_head_sha':source,'actual_checkout_sha':commit,'passed':total,'pairs':sum(item['pairs'] for item in jobs),'jobs':jobs})
report = {'schema_version':1,'task':'T30','evidence_class':'independent_downloaded_sanitized_CI_receipt_review',
          'source_head':source,'PR_synthetic_merge':merge,'exact_equal_tree':tree,
          'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),'auditor_sha256':sha(Path(__file__).read_bytes()),
          'checks':checks,'issues':issues,'result':'fail' if issues else 'pass','events':events,
          'limitations':['Downloaded CI artifacts are sanitized JSON/NUnit exports, not raw native observations or produced PDFs.','Selected engine/dependency digest receipts and source assertions were reviewed; absent hosted vendor binaries were not independently rehashed.','Hosted Windows Server/admin-token smoke does not certify Windows desktop, Explorer, account class, visual fidelity or exact release asset operation.','No test/application rerun occurred; actual checkout SHA is distinct from platform headSha for pull_request.']}
Path('tests/.work/T30-review/ci-original-audit.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({key:report[key] for key in ('checks','issues','result','events')}))
raise SystemExit(1 if issues else 0)
