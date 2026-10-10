"""Read-only focused image/PDF review of retained accepted T33 final-R outputs."""
import argparse
from contextlib import closing
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import random
import re
import subprocess
import sys

import PIL
import pypdf
import pypdfium2 as pdfium

ROOT = Path(__file__).resolve().parent
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--capture', type=Path, required=True)
parser.add_argument('--repo', type=Path, required=True)
parser.add_argument('--expected-harness-commit', required=True)
parser.add_argument('--report', type=Path, required=True)
args = parser.parse_args()
CAPTURE = args.capture.resolve()
if not args.report.resolve().is_relative_to(ROOT) or args.report.exists():
    raise ValueError('Use a NEW report under ignored T33-public-download-review')
EXPECTED_IDS = ["T03-01-P01", "T03-01-P02", "T03-02-P01", "T03-02-P02", "T03-10-P01", "T03-14-P01"]
EXPECTED_COMMIT = args.expected_harness_commit
SOURCE_COMMIT = "95e0a19e6cc5fc01cd4bec4ac15f989f9830840a"
ID_PATTERN = re.compile(r"T03-[0-9]{2}-P[0-9]{2}")
ORIGINAL_PIXELS = random.Random(140032).randbytes(1200 * 800 * 3)
SHA = lambda raw: hashlib.sha256(raw).hexdigest()
checks = 0

def check(condition, message):
    global checks
    checks += 1
    if not condition:
        raise AssertionError(message)

def regular(path):
    for candidate in [path, *path.parents]:
        info = candidate.stat(follow_symlinks=False)
        check(not candidate.is_symlink() and not getattr(info, "st_file_attributes", 0) & 0x400,
              "Audit input contains a link/reparse path")
    check(path.is_file(), "Audit input must be an existing regular file")

def inspect(path, expected_ids, alias):
    regular(path)
    raw = path.read_bytes()
    images = []
    with path.open("rb") as stream:
        reader = pypdf.PdfReader(stream, strict=True)
        check(len(reader.pages) == len(expected_ids), alias + ": strict page count")
        with pdfium.PdfDocument(path) as document:
            check(len(document) == len(expected_ids), alias + ": PDFium page count")
            pages = []
            for index, expected in enumerate(expected_ids):
                pypdf_page = reader.pages[index]
                check(ID_PATTERN.findall(pypdf_page.extract_text()) == [expected], alias + ": pypdf page IDs/order")
                check([float(v) for v in pypdf_page.mediabox] == [0.0, 0.0, 432.0, 288.0], alias + ": original media box")
                check(int(pypdf_page.get("/Rotate", 0)) == 0, alias + ": original rotation")
                with closing(document[index]) as page:
                    with closing(page.get_textpage()) as text:
                        check(ID_PATTERN.findall(text.get_text_range()) == [expected], alias + ": PDFium page IDs/order")
                    check(list(page.get_size()) == [432.0, 288.0] and page.get_rotation() == 0, alias + ": PDFium geometry")
                page_images = []
                for name, image in pypdf_page.images.items():
                    check(image.indirect_reference is not None, alias + ": expected ordinary image XObject")
                    obj = image.indirect_reference.get_object()
                    decoded_stream = obj.get_data()
                    pixels = image.image.convert("RGB")
                    row = {"page": index + 1, "name": str(name), "width": int(obj["/Width"]),
                           "height": int(obj["/Height"]), "bits_per_component": int(obj["/BitsPerComponent"]),
                           "color_space": str(obj["/ColorSpace"]), "filter": str(obj.get("/Filter", "none")),
                           "pypdf_get_data_bytes": len(decoded_stream), "pypdf_get_data_sha256": SHA(decoded_stream),
                           "rgb_pixel_bytes": len(pixels.tobytes()), "rgb_pixel_sha256": SHA(pixels.tobytes()),
                           "pil_width": pixels.width, "pil_height": pixels.height}
                    check([row["width"], row["height"]] == [pixels.width, pixels.height], alias + ": image dictionary/PIL geometry")
                    page_images.append(row)
                    images.append(row)
                pages.append({"page": index + 1, "identifier": expected, "size_points": [432, 288], "rotation": 0,
                              "image_count": len(page_images)})
    return {"alias": alias, "bytes": len(raw), "sha256": SHA(raw), "pages": pages, "images": images}

def original_image(row, label):
    check(len(row["images"]) == 1, label + ": one original raster image")
    image = row["images"][0]
    check([image["width"], image["height"], image["bits_per_component"], image["color_space"]] == [1200, 800, 8, "/DeviceRGB"],
          label + ": original raster resolution/color/bits retained")
    check(image["pypdf_get_data_bytes"] == len(ORIGINAL_PIXELS) and image["pypdf_get_data_sha256"] == SHA(ORIGINAL_PIXELS),
          label + ": raw decoded RGB stream matches original seeded bytes")
    check(image["rgb_pixel_bytes"] == len(ORIGINAL_PIXELS) and image["rgb_pixel_sha256"] == SHA(ORIGINAL_PIXELS),
          label + ": independent PIL decoded pixels match original seeded bytes")

result = {"task": "T33", "scope": "Focused read-only retained source/master/email PDF and decoded image inspection; no new application run, manual pass, universal fidelity, PDF/A or signature claim",
          "observed_at_utc": datetime.now(timezone.utc).isoformat(), "result": "fail", "issues": [],
          "harness_commit": EXPECTED_COMMIT, "candidate_source_commit": SOURCE_COMMIT,
          "versions": {"python": sys.version.split()[0], "pypdf": pypdf.__version__, "pillow": PIL.__version__,
                       "pypdfium2": str(pdfium.PYPDFIUM_INFO), "pdfium": str(pdfium.PDFIUM_INFO)},
          "expected_original_rgb_sha256": SHA(ORIGINAL_PIXELS), "files": [], "source_rasters": [], "email_observations": []}
try:
    check(result["versions"] == {"python": "3.12.14", "pypdf": "6.10.0", "pillow": "12.3.0", "pypdfium2": "5.13.0", "pdfium": "153.0.7999.0"}, "Approved reader pins")
    repo = args.repo.resolve()
    status_before = subprocess.check_output(["git", "-C", str(repo), "status", "--porcelain=v1", "--untracked-files=all"])
    check(not status_before, "Clean actual harness checkout required")
    head = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"], capture_output=True, timeout=10, check=True).stdout.decode().strip()
    check(head == EXPECTED_COMMIT, "Actual harness HEAD differs")
    result["input_receipts"] = []
    for shell in ("PS51", "PS7"):
        receipt_path = CAPTURE / (shell + "-reports/result.json")
        report = json.loads(receipt_path.read_text(encoding="utf-8-sig"))
        check(report["result"] == "pass" and not report["preparation"], shell + ": accepted receipt")
        check(report["task"] == "T33" and report["evidence_class"] == "actual_independently_anonymously_downloaded_published_v1_0_0_package_operation", shell + ": actual final-R package receipt")
        check(report["harness_commit"] == EXPECTED_COMMIT and report["candidate_source_commit"] == SOURCE_COMMIT, shell + ": observed source identity")
        result["input_receipts"].append({"path": receipt_path.relative_to(CAPTURE).as_posix(), "sha256": SHA(receipt_path.read_bytes())})
        work = Path(report["work_root"])
        for case in report["cases"]:
            case_root = Path(case["case_root"])
            check(case_root.is_relative_to(work), shell + ": case belongs to actual capture work root")
            outputs = [Path(p) for p in case["output_paths"] if Path(p).suffix.lower() == ".pdf"]
            for output in outputs:
                check(output.is_relative_to(case_root), shell + ": output belongs to actual case")
            raster_path = case_root / "synthetic inputs/20.pdf"
            is_normal = raster_path.is_file()
            if is_normal:
                source = inspect(raster_path, ["T03-14-P01"], shell + "/" + case["label"] + "/source-20.pdf")
                original_image(source, source["alias"])
                result["source_rasters"].append(source)
            retained_master = None
            retained_email = None
            for output in outputs:
                kind = "email" if output.name.endswith("_email.pdf") else "master"
                observed = inspect(output, EXPECTED_IDS if is_normal else ["T03-01-P01"], shell + "/" + case["label"] + "/" + kind)
                result["files"].append(observed)
                if kind == "master":
                    retained_master = observed
                    if is_normal:
                        original_image(observed, observed["alias"])
                        check(observed["images"][0]["page"] == 6, observed["alias"] + ": raster remains in natural final page")
                    else:
                        check(not observed["images"], observed["alias"] + ": tiny vector source remains vector/image-free")
                else:
                    retained_email = observed
            if retained_email:
                check(retained_master is not None and retained_email["bytes"] < retained_master["bytes"], shell + ": email exact size benefit")
                check(len(retained_email["images"]) == 1, shell + ": one rewritten email raster")
                image = retained_email["images"][0]
                check(0 < image["width"] <= 1200 and 0 < image["height"] <= 800, shell + ": actual email raster bounds")
                downsampled = [image["width"], image["height"]] != [1200, 800]
                check(downsampled or image["filter"] == "/DCTDecode", shell + ": actual downsampling or lossy JPEG rewrite")
                check(image["rgb_pixel_sha256"] != SHA(ORIGINAL_PIXELS), shell + ": actual email pixels changed")
                result["email_observations"].append({"case": shell + "/" + case["label"], "master_bytes": retained_master["bytes"],
                                                     "email_bytes": retained_email["bytes"], "width": image["width"], "height": image["height"],
                                                     "filter": image["filter"], "downsampled": downsampled,
                                                     "rgb_pixel_sha256": image["rgb_pixel_sha256"]})
    result["retained_pdf_count"] = len(result["files"])
    result["retained_output_page_count"] = sum(len(row["pages"]) for row in result["files"])
    result["normal_master_image_count"] = sum(1 for row in result["files"] if row["alias"].endswith("/master") and row["images"])
    result["rewritten_email_image_count"] = len(result["email_observations"])
    result["source_raster_count"] = len(result["source_rasters"])
    status_after = subprocess.check_output(["git", "-C", str(repo), "status", "--porcelain=v1", "--untracked-files=all"])
    head_after = subprocess.check_output(["git", "-C", str(repo), "rev-parse", "HEAD"]).decode().strip()
    check(status_after == status_before and head_after == EXPECTED_COMMIT, "Actual harness source unchanged throughout read-only review")
    result["result"] = "pass"
except Exception as error:
    result["issues"].append(type(error).__name__ + ": " + str(error))
result["checks"] = checks
result["review_script_sha256"] = SHA(Path(__file__).read_bytes())
args.report.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
print(json.dumps({k:v for k,v in result.items() if k not in ("files", "source_rasters")},indent=2))
raise SystemExit(0 if result["result"] == "pass" else 1)
