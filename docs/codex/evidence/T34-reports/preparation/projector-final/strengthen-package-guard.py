from pathlib import Path
import hashlib,difflib,json
HERE=Path(__file__).resolve().parent
p=HERE/'Export-T34.py';raw=p.read_bytes()
assert hashlib.sha256(raw).hexdigest()=='53c2d5e974fa01d5f27f574140211f11b720b34db9f579da95f47ed3dac011c6'
old=raw.decode();(HERE/'unexecuted-first-adapter.py').write_bytes(raw)
needle="        binding_row=self.config['native_reuse_binding'];reuse=self.bound_gate(binding_row)"
new=old.replace(needle,"        require(len(package['checks'])==266 and all(x['pass']is True for x in package['checks']) and package['application_executed']is False and package['native_engines_executed']is False, 'Every exact fresh package check must pass without application/native execution')\n"+needle)
assert new!=old
p.write_bytes(new.encode())
diff=''.join(difflib.unified_diff(old.splitlines(True),new.splitlines(True),fromfile='unexecuted-first-adapter.py',tofile='Export-T34.py'))
(HERE/'package-child-guard.diff').write_bytes(diff.encode())
(HERE/'final-source-derivation.json').write_bytes((json.dumps({'task':'T34','scope':'One explicit package child scope/count guard added before probes or exporter invocation; initial source preserved','previous_source_sha256':hashlib.sha256(raw).hexdigest(),'source_sha256':hashlib.sha256(new.encode()).hexdigest(),'diff_sha256':hashlib.sha256(diff.encode()).hexdigest()},indent=2)+'\n').encode())
print(json.dumps({'source_sha256':hashlib.sha256(new.encode()).hexdigest()}))
