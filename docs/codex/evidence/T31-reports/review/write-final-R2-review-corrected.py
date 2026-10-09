"""Write final review summary from immutable actual independent audit reports."""
from pathlib import Path
import datetime, hashlib, json

repo=Path.cwd().resolve()
folder=repo/'tests/.work/T31-review'
read=lambda path:json.loads(path.read_text(encoding='utf-8'))
sha=lambda path:hashlib.sha256(path.read_bytes()).hexdigest()
expected='95e0a19e6cc5fc01cd4bec4ac15f989f9830840a'
original=read(folder/'final-R2-original-audit.json')
native=read(folder/'final-R2-native-original-audit.json')
invocations=read(folder/'final-R2-auditor-invocations.json')
lineage=read(folder/'source-R2-lineage-audit-corrected.json')
assert original['result']==native['result']==invocations['result']=='pass'
assert lineage['result']=='pass_for_R2_lineage_source_and_capture_preparation' and lineage['checks']==255 and lineage['issues']==[]
assert original['source_commit']==native['source_commit']==invocations['source_commit']==expected
assert original['issues']==native['issues']==[] and original['incomplete']==[]
assert invocations['source_unchanged'] and invocations['source_end']['head']==invocations['source_end']['live_main']==expected
assert invocations['source_end']['branch']=='main' and invocations['source_end']['status']==''
outputs=[folder/'final-R2-source-review.json',folder/'final-R2-source-review.md']
assert not any(path.exists() for path in outputs)
paths=['source-R2-lineage-audit-corrected.py','source-R2-lineage-audit-corrected.json',
       'audit-R2-original.py','audit-R2-original-corrected.py','audit-R2-native.py',
       'R2-auditor-static-correction-provenance.json','final-R2-original-audit.json',
       'final-R2-native-original-audit.json','invoke-final-R2-audits.py',
       'final-R2-auditor-invocations.json']
bindings=[{'path':str((folder/name).relative_to(repo)).replace('\\','/'),
           'sha256':sha(folder/name)} for name in paths]
constraints=[
    'R1 de5f30155c68755dbd5af691625a0651e3fb7230 remains unaccepted. Its fixture failure, aborted full scopes and original diagnostics remain historical failures, separate from actual corrected R2 tests.',
    'The reviewed corrective C1 30560516a0248636769e988b0420466214c25e3b and normal PR27 merge produce R2 with tree 5014f5bdf4f374aee828ced4c39cb93bfeb6465a. Exact runtime, VERSION, builder, allowlist, native arguments, workflow and public-document Git identities remain unchanged from reviewed C2/R1.',
    'All seven recipe/manifest raw pins remain strict and match both actual R2 Git blobs and working bytes. Presets manifest change preserves its existing CRLF pin and JSON semantics; no expectation/hash weakening or manual acceptance normalization is used.',
    'Full count 2144 is actual R2 local execution across two required hosts, 32 tiers per host and 64 JSON/NUnit pairs. Evidence classes include controlled, unit, synthetic-package and document checks; the total does not mean 2144 PDF-engine or manual cases.',
    'Static is 68 maintained files and 41 selected rules per host, zero selected findings/suppressions. Visible all-rule advisories are 0 errors, 349 warnings, 175 information per host. Actual combined default-both capture argv is legitimate and bound to both actual pinned child hosts.',
    'Development helpers are 85 passed plus one explicitly disclosed symlink-creation skip: handoff 26 pass/1 skip of 27, fixture/oracles 42 pass, package helpers 17 pass. The skip is never relabeled as application, Windows/native, manual or source failure/pass.',
    'Actual native evidence is from original producer receipts and retained bytes; controlled fault hooks and native usage-only cases remain disclosed. Larger benign-warning candidates were inspected then cleaned and cannot be rehashed as retained PDFs by this reviewer.',
    'Legacy StandardUser=true is only the original nonadministrator-token predicate. AC058 remains owner-excluded and unperformed; no human account class, Explorer, physical viewer or enrollment acceptance is inferred.',
    'Four current PR27 check jobs are the actual configured premerge gate. Actual merged-main CI, coverage closure and public evidence projection audit remain separately reviewed scopes; this summary does not convert PR synthetic head or prior-source CI into exact R2 execution.',
    'T32 exact final assets/package operation, T33 publication and independent downloaded operation, and T34 synchronized closure remain later gates. No release/tag/publication completion is claimed.',
    'Earlier prepared static-selector assumption was corrected before final auditor execution, with the unexecuted old source retained. It is an auditor preparation correction, not a product/test/static failure.'
]
report={'schema_version':1,'task':'T31','phase':'R2','source_commit':expected,
        'evidence_class':'independent_final_exact_R2_source_and_original_evidence_review',
        'observed_at_utc':datetime.datetime.now(datetime.timezone.utc).isoformat(),
        'summary_source_sha256':sha(Path(__file__)),
        'result':'pass','findings':[],
        'lineage_review_checks':lineage['checks'],
        'original_receipt_audit_checks':original['checks'],
        'native_original_audit_checks':native['checks'],
        'full_hosts':[{'shell':row['shell'],'tiers':row['tiers'],'passed':row['passed'],'complete':row['complete']} for row in original['full_hosts']],
        'local_full_checks_total':sum(row['passed'] for row in original['full_hosts']),
        'local_json_nunit_pairs':sum(row['tiers'] for row in original['full_hosts']),
        'static_hosts':original['static_hosts'],
        'development_helper_counts':original['R2_development_helper_counts'],
        'extras_actual_commands':original['R2_extras_commands'],
        'distinct_source_bindings':original['distinct_source_blob_hashes_verified'],
        'native_original_receipts':len(native['receipts']),
        'native_retained_files_verified':native['retained_files_verified'],
        'native_observation_records':sum(row['observation_records'] for row in native['receipts']),
        'native_capture_records':sum(row['capture_records'] for row in native['receipts']),
        'source_after_audits':invocations['source_end'],
        'bindings':bindings,'completion_record_constraints':constraints,
        'reviewer_execution_scope':'Read-only source/receipt/hash audit; no application, test runner, PDF engine, rendering or static producer reexecution.'}
outputs[0].write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
paragraphs=[
    '# T31 independent final R2 source and evidence review',
    'PASS, no findings. Accepted review source is actual normally merged R2 `'+expected+'`; the read-only audit wrapper freshly verified local/current/live clean `main` at this SHA before and after auditing.',
    'The prior independent lineage/source audit passed '+str(lineage['checks'])+' checks. The completed original source/report audit passed '+str(original['checks'])+' checks with zero issues, verifying 312 distinct source bindings, all 64 original JSON/NUnit pairs, per-tier outcomes/leaf counts/actual selected hosts, raw streams, clean source equality, unchanged capture drivers and all 348 approved dependency payloads.',
    'Both required full runs completed 32 tiers and 1,072 checks each (2,144 total), with zero failure/skip/not-run/blocked counts. Both full static reports cover 68 maintained files and 41 selected rules, zero selected findings/suppressions, with all-rule advisories 0 errors/349 warnings/175 information each. Ten extras commands passed; developer helpers total 85 passed and one disclosed symlink skip (26+skip, 42, 17).',
    'The separate native original audit passed '+str(native['checks'])+' checks with zero issues: '+str(len(native['receipts']))+' actual original receipts, 288 retained files, 358 observation records and 22 capture records. Source/foreign/prior-output preservation, pinned native dependencies, original independent PDFium receipts and real retained output byte/hash bindings were checked without running the application or engines.',
    'Actual commands, raw stdout/stderr hashes, source hashes and timing are retained in `final-R2-auditor-invocations.json`. Exact authoritative result files are `final-R2-original-audit.json` and `final-R2-native-original-audit.json`; the summary JSON binds each report/source by SHA256.',
    '## Constraints for completion records',
    *['- '+item for item in constraints],
    'This reviewer scope is complete and files are stable for the parent core freeze. Any later completion-record/diff review must use a new ignored output outside the frozen review payload.'
]
outputs[1].write_text('\n\n'.join(paragraphs)+'\n',encoding='utf-8')
print(json.dumps({'result':'pass','outputs':[{'path':str(path.relative_to(repo)).replace('\\','/'),'sha256':sha(path)} for path in outputs]},indent=2))
