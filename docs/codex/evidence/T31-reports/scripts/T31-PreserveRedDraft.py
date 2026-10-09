from pathlib import Path
import hashlib,json
p=Path('tools/test/tests/test_fixture_checkout.py');raw=p.read_bytes()
initial=raw.replace(b'"core.attributesFile=", "-c", "core.hooksPath=", *arguments',b'"core.attributesFile=", *arguments')
expected='72e653efc3664c1ab2af58b92b7d7430ca7cadda7b2da8a433c4dc255ccedbfe'
for name,value in [('unchanged-final',raw),('remove-hooks',initial),('remove-hooks-final-blank',initial+b'\n'),('remove-hooks-CRLF',initial.replace(b'\n',b'\r\n'))]:
    print(name,hashlib.sha256(value).hexdigest())
    if hashlib.sha256(value).hexdigest()==expected:
        target=Path('tests/.work/T31-checkout-regression-red/test_fixture_checkout.initial-draft.py');target.write_bytes(value)
        (target.parent/'source-reconstruction.json').write_text(json.dumps({'original_source_sha256':expected,'reconstructed_sha256':hashlib.sha256(value).hexdigest(),'matched_original_recorded_bytes':True,'derivation':name,'final_source_sha256':hashlib.sha256(raw).hexdigest(),'limitations':'Reconstructed source byte payload matches immutable original red ledger digest; draft underwent only scoped safety refinement before green test.'},indent=2)+'\n',encoding='utf-8')
