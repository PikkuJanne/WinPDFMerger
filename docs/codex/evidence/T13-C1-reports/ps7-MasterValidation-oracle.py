import json
import re
import sys
from contextlib import closing
from pathlib import Path
import pypdfium2 as pdfium
if str(pdfium.PYPDFIUM_INFO) != "5.13.0" or str(pdfium.PDFIUM_INFO) != "153.0.7999.0":
    raise RuntimeError("Independent master oracle requires the recorded PDFium pins.")
if sys.argv[1:] == ["--versions"]:
    print(json.dumps({"python": sys.version.split()[0], "pypdfium2": str(pdfium.PYPDFIUM_INFO), "pdfium": str(pdfium.PDFIUM_INFO)}))
    raise SystemExit(0)
expected = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8-sig"))
actual = []
with pdfium.PdfDocument(Path(sys.argv[1])) as document:
    for index in range(len(document)):
        with closing(document[index]) as page:
            with closing(page.get_textpage()) as text:
                identifiers = re.findall(r"T03-[0-9]{2}-P[0-9]{2}", text.get_text_range())
            if len(identifiers) != 1:
                raise ValueError(f"Page {index + 1} has unexpected visible identifiers: {identifiers}")
            quarter_turns = int(pdfium.raw.FPDFPage_GetRotation(page))
            degrees = page.get_rotation()
            if quarter_turns < 0 or quarter_turns * 90 != degrees:
                raise ValueError("Raw PDFium rotation and pypdfium2 degrees differ.")
            actual.append({"identifier": identifiers[0], "rotation_degrees": degrees,
                           "rotation_quarter_turns": quarter_turns, "size_points": list(page.get_size())})
if actual != expected:
    raise ValueError(f"Merged pages differ: expected {expected}, actual {actual}")
print(json.dumps({"page_count": len(actual), "pages": actual, "pypdfium2": str(pdfium.PYPDFIUM_INFO), "pdfium": str(pdfium.PDFIUM_INFO)}))