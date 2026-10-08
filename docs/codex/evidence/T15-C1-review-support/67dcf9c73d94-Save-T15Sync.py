from pathlib import Path
import argparse
import json
import subprocess
import sys

parser=argparse.ArgumentParser()
parser.add_argument('--phase',choices=['C1','C2'],required=True)
parser.add_argument('--expected',required=True)
args=parser.parse_args()
destination=Path('tests/.work')/('T15-'+args.phase+'-live-sync.json')
assert not destination.exists(), 'Preserve previous live receipt.'
proc=subprocess.run([sys.executable,'-B','tools/codex/handoff.py','sync','--repo','.'],capture_output=True)
if proc.returncode:
    sys.stdout.buffer.write(proc.stdout)
    sys.stderr.buffer.write(proc.stderr)
    raise SystemExit(proc.returncode)
result=json.loads(proc.stdout)
assert result['local_head']==result['live_remote_head']==args.expected
assert result['clean'] and result['synchronized']
destination.write_bytes(proc.stdout)
sys.stdout.buffer.write(proc.stdout)
