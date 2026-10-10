"""Prepare a closed T34 independent comparison kernel; no producer imports."""
from pathlib import Path
import ast,difflib,hashlib,json
root=Path(__file__).resolve().parent
repo=root.parents[2]
base=repo/'docs/codex/evidence/T33-reports/review/public-review.py';raw=base.read_bytes()
assert hashlib.sha256(raw).hexdigest()=='b2dc0064e1dcf4d6844ac731a0f288167bfcd8022f3d0cce0f972e8925ddac53'
original=raw.decode('utf-8');text=original.replace('T33','T34')
text=text.replace("M = 'f3d8f3c8e8a8c582171c36ff1ceed82d84b09232'", "M = 'b6897ea75037d2d1f1d8ed88e08d214a25d3b143'")
text=text.replace("(?:t32-(?:source|artifacts)|t33-public-download)", "(?:t32-source|t34-public-download)")
text=text.replace("self.must(Counter(row['kind'] for row in declarations) == {'actions_runs': 1, 'annotated_tag': 3}, 'Only one owner evidence Actions response and three separately captured final tag originals')", "self.must(all(row['kind'] in {'actions_runs','actions_run_pages','annotated_tag'} for row in declarations), 'Only explicit pinned metadata kinds; final exact registry adapter remains closed')")
text=text.replace("return resolved\n    def add", "self.must(resolved.relative_to(self.work).parts[0].startswith('T34'), 'Only owned explicit T34 originals selected; prior native receipts referenced without reexport')\n        return resolved\n    def add")
start=text.index("        if path.name.startswith('audited-packet-diff.stdout'):")
end=text.index("        if path.suffix.lower()",start)
text=text[:start]+text[end:]
start=text.index("            for omission in row.get(")
end=text.index('            members = ',start)
text=text[:start]+"            self.must(not row.get('curated_git_z_omissions') and not row.get('exclude'), 'Exact final T34 curated omissions adapter not supplied; no generic omission allowed')\n"+text[end:]
start=text.index('                if any(relative == excluded')
end=text.index('        for row in c.get(',start)
text=text[:start]+"                public_relative = relative[:-4] + '.txt' if stream else relative\n                self.add(row['label'] + '/' + public_relative, path, row['provenance'] + '; ' + row['scope'])\n"+text[end:]
text=text.replace("self.check(len(self.curated_omissions) == 5 and sorted(item['entry_count'] for item in self.curated_omissions) == [2406,2406,2406,2406,9093] and len(self.omitted) == 5, 'Only five authorized Git-z omissions with independently decoded exact counts/classifications')", "self.check(not self.curated_omissions and not self.omitted, 'No generic text/NUL omission in closed prepared kernel')")
start=text.index('    def metadata(');end=text.index('    def xml_bytes(',start)
text=text[:start]+"    def metadata(self, raw, label):\n        raise ValueError('Exact final T34 metadata receipt/hash/run/index/person adapter not supplied')\n"+text[end:]
text=text.replace("        operation_root = self.owned(next(row['source'] for row in self.config['acceptance_gates'] if row['role'] == 'accepted_native')).parent\n",'')
text=text.replace("                    if raw_path.is_relative_to(operation_root):\n                        stream_count += 1\n                    else:\n                        other_stream_count += 1", "                    other_stream_count += 1")
text=text.replace("self.check(stream_count == 210, 'All 210 original operation stdout/stderr byte streams independently included')", "self.check(stream_count == 0, 'No prior native payload rerun or reexport counted as new T34 operation')")
start=text.index('    def gates(');end=text.index('    def run(',start)
text=text[:start]+"    def gates(self, manifest=None):\n        raise ValueError('Final actual T34 gate schemas/role set/hash links not supplied; no acceptance')\n"+text[end:]
text=text.replace("'T34_synchronized_closure_inferred': False", "'final_main_closure_inferred': False")
ast.parse(text);out=root/'public-review.py';assert not out.exists();out.write_text(text,encoding='utf-8',newline='\n')
(root/'kernel-derivation.diff.txt').write_text(''.join(difflib.unified_diff(original.splitlines(True),text.splitlines(True),fromfile='accepted-T33-public-review.py',tofile='closed-T34-public-review.py')),encoding='utf-8')
report={'task':'T34','result':'pass_for_closed_independent_comparison_kernel_preparation','base_sha256':hashlib.sha256(raw).hexdigest(),'derived_sha256':hashlib.sha256(out.read_bytes()).hexdigest(),'closed_adapters':['All actual gate schemas/roles/paths/hashes','Explicit final NUL-curation rawSHA/size/kind/count/source/argv/ledger','Exact metadata receipt rawSHA/run/head/index and allowed email pointers'],'preserved':['TypedJSON key order/node type/scalar comparison','BOM and XML decoded facts/exact byte projection','All real local identity/exactprefix checks on source/data; only supplemental unknownUUID heuristic source exemption','Complete selected original/public inventory/hash/size/provenance and safe paths','Strict UTF8 captured streams/no generic NUL or binary omission'],'scope':{'producer_imported_or_executed':False,'application_native_or_network_execution':False,'actual_acceptance_or_final_main_closure_claimed':False}}
(root/'kernel-preparation.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8');print(json.dumps(report))
