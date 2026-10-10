"""Preserve initial selection builder and align original local ledger path schema."""
from pathlib import Path
import hashlib
import json
HERE=Path(__file__).resolve().parent
source=(HERE/'build-selection.py').read_text(encoding='utf-8')
old="stream=command['streams']['stdout'];s=Path(stream['path']);s=s if s.is_absolute()else REPO/s"
new="stream=command['streams']['stdout'];s=Path(stream['path']);s=s if s.is_absolute()else ledgerpath.parent/s"
assert source.count(old)==1
source=source.replace(old,new)
marker=" root('T33-export-preparation','preparation/projector-initial'"
index=source.index(marker)
addition=" root('T33-record-preparation','preparation/record-writer','preparation','Prepared V1/V2 writer and exact original correction/probes; no record write yet'),\n root('T33-writer-review','review/writer-source','review','Independent53 source/semantic developer checks only'),\n root('T33-final-actions/prepare-reviewed-helper-guards-v2-cd92090a20f84ca4b3258f25577f1d5b','actions/prepare-helper-v2','preparation','Original guard-source derivation/parse checks; no Git checkpoint execution'),\n"
source=source[:index]+addition+source[index:]
target=HERE/'build-selection-v2.py';assert not target.exists();target.write_bytes(source.encode())
(HERE/'selection-preparation-history.json').write_text(json.dumps({'task':'T33','scope':'Selection helper preparation only; no exporter invocation/public/native failure',
    'preserved_original_builder_sha256':hashlib.sha256((HERE/'build-selection.py').read_bytes()).hexdigest(),
    'corrected_builder_sha256':hashlib.sha256(target.read_bytes()).hexdigest(),
    'failure':'Initial builder expected ledger stream paths relative to repo; original task-owned ledgers actually record filenames relative to their own capture root. Exact ledger rawSHA/argv/source/count remain required.',
    'original_failure_retention':'Tool-returned exit1/AssertionError plus original builder source retained; no standalone failed-builder stream receipt existed.',
    'added_frozen_roots':['T33-record-preparation','T33-writer-review','prepare-reviewed-helper-guards-v2-cd92090a20f84ca4b3258f25577f1d5b']},indent=2)+'\n',encoding='utf-8')
