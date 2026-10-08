from pathlib import Path
import hashlib,json,subprocess
repo=Path.cwd();work=repo/'tests/.work';p='docs/codex/evidence/T15-completion.md';sha=lambda b:hashlib.sha256(b).hexdigest()
raw=(repo/p).read_bytes();indexed=subprocess.check_output(['git','show',':'+p])
assert raw!=indexed and raw.replace(b'\r\n',b'\n')==indexed
receipt={'Task':'T15','Classification':'records-only staged byte check, no suite or application failure',
 'Command':'<approved Python> -B tests/.work/Validate-T15C2.py --stage','ObservedExitCode':1,
 'ObservedToolTranscriptError':'AssertionError: Exact staged Git evidence bytes: docs/codex/evidence/T15-completion.md',
 'ObservedTranscriptOnly':True,'SeparateOriginalStdoutStderrFilesRetained':False,
 'File':p,'WorkingSHA256':sha(raw),'ThenIndexedSHA256':sha(indexed),'NormalizationEquivalent':True,
 'ValidatorSourceSHA256':sha((work/'Validate-T15C2.py').read_bytes()),
 'Correction':'Add a file-specific -text attribute for generated completion note; restage and recheck literal staged bytes.',
 'NoApplicationChange':True,'NoApplicationRerun':True}
target=work/'T15-C2-staged-normalization-preparation.json';assert not target.exists()
target.write_text(json.dumps(receipt,indent=2)+'\n',encoding='utf-8');print(json.dumps(receipt))
