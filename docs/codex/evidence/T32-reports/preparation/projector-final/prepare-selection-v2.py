"""Derive explicit V2 selection preserving V1. Final source review is a later explicit append."""
from pathlib import Path
import hashlib
root=Path(__file__).resolve().parent
initial=root.parent/'T32-export-preparation/build-selection.py'
text=initial.read_text()
text=text.replace("add('T32-export-preparation','preparation/projector','preparation',scope='Compact explicit text projection source and synthetic checks only; no application/native/platform acceptance')",
                  "add('T32-export-preparation','preparation/projector-initial','preparation',scope='Preserved initial projector and 24 synthetic checks; later parent-link correction does not promote them to application/native/platform acceptance')\nadd('T32-export-final-preparation','preparation/projector-final','preparation',scope='Corrected projector source and 28 synthetic checks, including actual NTFS junction fixture; no application/native/platform acceptance')")
text=text.replace("'T32-PolishWriter.py','T32-WritePreparation.py']",
                  "'T32-PolishWriter.py','T32-WritePreparation.py','T32-WriteCompletionV2.diff.txt','T32-WriteCompletionV3.py','T32-WriteCompletionV3.diff.txt','T32-PolishWriterV3.py','T32-EvidenceCheckpoint.py']")
text=text.replace("T32-export-selection-v1.json", "T32-export-selection-v2.json")
with (root/'build-selection-v2.py').open('x',encoding='utf-8',newline='\n') as f:f.write(text)
print(hashlib.sha256((root/'build-selection-v2.py').read_bytes()).hexdigest())
