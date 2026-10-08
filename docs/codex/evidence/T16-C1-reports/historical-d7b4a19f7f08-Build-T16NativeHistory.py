from pathlib import Path
import hashlib,json,re

repo=Path(__file__).resolve().parents[2]
work=repo/'tests/.work'
def digest(raw):return hashlib.sha256(raw).hexdigest()
def load(path):return json.loads(path.read_bytes().decode('utf-8-sig'))
def binding(path):
    raw=path.read_bytes()
    return {'Path':path.relative_to(repo).as_posix(),'SHA256':digest(raw),'Bytes':len(raw)}
source=(repo/'tests/cli/Parameters.Native.Tests.ps1').read_bytes()
initial=source.replace(b'        # PS5.1 emits a JSON array as one pipeline object. Language foreach\n        # enumerates the parsed array explicitly in both supported shells.\n        $calls = @(foreach ($call in (ConvertFrom-Json -InputObject ([IO.File]::ReadAllText((Join-Path $App.Capture \'native-calls.json\'),[Text.Encoding]::UTF8)))) { $call })',b"        $calls = @(Get-Content -LiteralPath (Join-Path $App.Capture 'native-calls.json') -Raw | ConvertFrom-Json)")
assert digest(initial)=='7bd389fe8fbd94e70fc6f11fa6b776a9c0058bb1daf01f9d1adc532b08f3a919'
initial_path=work/'T16-Parameters.Native-initial.Tests.ps1'
if initial_path.exists():assert initial_path.read_bytes()==initial
else:initial_path.write_bytes(initial)
final_path=work/'T16-Parameters.Native-focused.Tests.ps1'
if final_path.exists():assert final_path.read_bytes()==source
else:final_path.write_bytes(source)
attempts=[]
for root in sorted(work.glob('T16-dirty-ps*-ParametersNative-*')):
    launch=load(root/'launch.json');out=(root/'stdout.txt').read_bytes().decode('utf-8-sig')
    report=Path(re.findall(r'(?m)^Reports: (.+?)\r?$',out)[0]);summary=load(report/'summary.json')
    receipt=Path(re.findall(r'(?m)^Parameters observations: (.+?)\r?$',out)[0]);obs=load(receipt)
    expected=initial_path if launch['TestSourceSHA256'].lower()==digest(initial) else final_path
    assert digest(expected.read_bytes())==launch['TestSourceSHA256'].lower()
    files=[root/'launch.json',root/'stdout.txt',root/'stderr.txt',root/'exit-code.txt',report/'summary.json',report/'results.xml',receipt,expected]
    copies=[]
    for case in sorted(p for p in receipt.parent.iterdir() if p.is_dir()):
        evidence=list((case/'captured-calls').glob('*.json'))+list((case/'app with spaces').glob('*.log'))+list((case/'named output').glob('*.log'))
        evidence+=[case/'app with spaces/WinPDFMerge.ps1',case/'app with spaces/src/WinPDFMerge.Helpers.ps1',case/'app with spaces/WinPDFMerge.bat']
        copies.append({'Root':case.relative_to(repo).as_posix(),'Files':[binding(p) for p in sorted(evidence)]})
    attempts.append({'Selection':launch['Shell'],'Root':root.relative_to(repo).as_posix(),'Passed':summary['passed'],'Failed':summary['failed'],'Total':summary['total'],'ExitCode':int((root/'exit-code.txt').read_text()),'ObservationCount':len(obs['Observations']),'TestSourceAtInvocationSHA256':launch['TestSourceSHA256'].lower(),'SourceSnapshot':expected.relative_to(repo).as_posix(),'Diagnosis':('Test receipt reader used pipeline ConvertFrom-Json array output; PS5.1 retained the array as one object and shorthand Where-Object could not resolve Executable. Actual entry/native jobs succeeded; one missing-input assertion passed. Corrected test-only language foreach enumeration. No runtime modification.' if summary['failed'] else 'All nine focused native parameter cases passed; dirty run excluded from subsequent clean acceptance.'),'Files':[binding(p) for p in files],'RetainedCopiedApplicationCases':copies})
index={'SchemaVersion':1,'Task':'T16','Attempts':attempts,'Scope':'Historical dirty focused runs; all raw XML/summary/streams/receipts/source snapshots retained. Source input PDFs, validated outputs and complete copied-helper evidence remain ignored for later read-only independent inspection. No dirty counts contribute to clean acceptance.'}
path=work/'T16-native-dirty-history.json';path.write_bytes((json.dumps(index,indent=2)+'\n').encode())
print(json.dumps({'Path':path.relative_to(repo).as_posix(),'SHA256':digest(path.read_bytes()),'Attempts':[{'Shell':a['Selection'],'Passed':a['Passed'],'Failed':a['Failed'],'Total':a['Total'],'Observations':a['ObservationCount']} for a in attempts]},indent=2))
