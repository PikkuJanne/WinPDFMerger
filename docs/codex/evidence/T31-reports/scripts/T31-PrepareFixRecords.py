from pathlib import Path
import json
base=Path('docs/codex')
p=base/'TASKS.json';v=json.loads(p.read_text(encoding='utf-8'))
task=next(t for t in v['tasks'] if t['id']=='T31');task.update(status='in_progress',evidence=['docs/codex/evidence/T31-checkout-fix.md'],notes='PR26 merged normally to first candidate de5f301; fresh checkout exposed required fixture-byte failures. Candidate unaccepted. Five-file byte contract and real true/false Git checkout regression are under reviewed fix; new merged exact-source acceptance remains required.')
p.write_text(json.dumps(v,indent=2,ensure_ascii=False)+'\n',encoding='utf-8',newline='\n')
p=base/'STATUS.md';s=p.read_text(encoding='utf-8')
s=s.replace('Next task T31 is pending: merge the accepted PR and freeze release-source R.','T31 is in_progress. PR26 merged normally at2026-10-09T17:46:48Z to first\ncandidate de5f30155c68755dbd5af691625a0651e3fb7230; it is unaccepted. Fresh\nWindows checkout exposed required fixture-byte provenance failures (34pass,\n1failure,5errors). Reviewed five-file byte-preservation fix and real fresh Git\ncheckout regression are needed before a new merged source can be accepted.\nSee evidence/T31-checkout-fix.md. AC071/072 remain not_run.')
start=s.index('C1b clean/live synchronization is recorded;')
s=s[:start]+'''T30 C1b/C2 clean/live synchronization was recorded and PR26's exact eight checks
were freshly verified before its normal merge. First candidate scoped static/CI
passes do not supersede the required fixture failure. No release source is frozen;
release state remains not_started. T31 accepted new merged source/tests,
T32final exact assets, T33publication/independent downloaded operation and
T34synchronized closure remain required. Unsigned/dependency/PDF/privacy limits
remain. First-candidate partial captures and preparation failures are retained.
'''
p.write_text(s,encoding='utf-8',newline='\n')
(base/'NEXT_SESSION.md').write_text('''# Next session

Continue exactly T31: merge accepted source and freeze release-source R.
T01–T30 remain done; T31 in_progress, AC071/072 not_run, later T32–T34 pending.
Read AGENTS/INDEX/STATUS, T31 TASKS/brief, PRODUCT_SPEC, ACCEPTANCE_CASES,
GITHUB_WORKFLOW, TEST_STRATEGY, SECURITY_AND_DEPENDENCIES, RELEASE_RUNBOOK,
DEFINITION_OF_DONE and evidence/T31-checkout-fix.md.

PR26 was normally merged at2026-10-09T17:46:48Z. First candidate
de5f30155c68755dbd5af691625a0651e3fb7230 is explicitly unaccepted: fresh
Windows checkout under existing system autocrlf=true converted four pinned LF
development fixture files, causing34pass/1failure/5errors. Presets manifest's
pinned CRLF bytes were also normalized in Git, breaking autocrlf=false. Preserve
these failures, partial full runs and successful scoped CI/static facts separately.
Original receipts/producers/reviews remain under ignored tests/.work/T31-*.

Reuse codex/t31-fixture-checkout for the reviewed five-file -text fix and actual
true/false fresh Git checkout regression. No pin/expectation/runtime/version/
builder/allowlist changes. Verify the preset manifest staged bytes retain SHA
383ec07bec823f131ab86cf07de2ffb96e10803ad4f2f4997ba834835ee01e1f.
Dirty-base green preparation42fixture-helper passes are not final acceptance.
Normally commit/push, verify matching live ref and clean state, review exact fix
head and CI, then normally merge without admin bypass. Fetch/fast-forward main;
run required full32-tier regression in actual PS5.1 and pinnedPS7.6.6 plus static,
helpers and exact new merged CI using a genuinely fresh checkout. Identify exact
accepted R only after required passes. No freeze/tag/draft/publication yet.

After accepted R, open separate codex/v1.0.0-release-evidence branch. Changes
then ONLY docs/codex. Future report byte-preservation rules belong in existing
docs/codex/evidence/.gitattributes. Preserve raw/public hashes, typed JSON and
entity-escaped XML identity/path projections; no private PDFs/native binaries.
No record may claim its own future commit/push. Verify clean/live synchronization.

Reuse approved348-file caches and workspacePython3.12.14; no silent install,
elevation or persistent policy/PATH/security/Git-config change. Ambient7.6.5 is
not pinned7.6.6 evidence. Record actual Professional26H2/full26300.9457 and
nonadministrator token facts without inferring account class/Insider enrollment.
AC058 human standard-user/Explorer/viewer acceptance remains excluded,
unperformed and never pass. Windows10/liveUNC/ARM/32-bit-host exclusions and
unsigned/dependency/PDF/privacy limits remain. Only T32 final assets, T33 final
publication/independent download and T34 closure complete the public v1.0.0;
no extra ceremonial owner approval is required.
''',encoding='utf-8',newline='\n')
print('T31 fix checkpoint records prepared; no acceptance claimed')
