"""Independent existing-PNG pixel/provenance binding; no PDF rendering or visual-pass claim."""
import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import PIL
from PIL import Image

REPO = Path(__file__).resolve().parents[2]
WORK = REPO / "tests/.work"
C1 = "50220eccd1917e44a52d94cbfe3bf35b1940f3d8"
PYTHON = "dd5f8d19f6755d6491ee7c4bef2fe35ddd521334cc3ca3ed8fc93ebcadf135d0"
sha = lambda data: hashlib.sha256(data).hexdigest()


def main():
    parser = argparse.ArgumentParser(); parser.add_argument("--proof", required=True)
    parser.add_argument("--output", required=True); args = parser.parse_args()
    output = Path(args.output).resolve(); assert output.is_relative_to(WORK) and not output.exists()
    assert sha(Path(sys.executable).read_bytes()) == PYTHON and PIL.__version__ == "12.3.0"
    assert subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip() == C1
    assert not subprocess.check_output(["git", "status", "--porcelain=v1"], cwd=REPO, text=True).strip()
    bindings = {}; pngs = {}; pixels = {}; failures = []; checks = 0

    def check(value, label):
        nonlocal checks
        checks += 1
        if not value: failures.append(label)

    def read(path, binary=False):
        path = Path(path).resolve(); assert path.is_relative_to(WORK)
        data = path.read_bytes(); row = dict(Path=path.relative_to(REPO).as_posix(), SHA256=sha(data), Bytes=len(data))
        (pngs if binary else bindings)[row["Path"]] = row
        return data

    def obj(path): return json.loads(read(path).decode("utf-8-sig"))

    def pixel(path):
        key = str(Path(path).resolve())
        if key not in pixels:
            raw = read(path, True)
            with Image.open(path) as image:
                rgb = image.convert("RGB")
                pixels[key] = dict(PixelSHA256=sha(rgb.tobytes()), Width=rgb.width, Height=rgb.height, PNGSHA256=sha(raw))
                rgb.close()
        return pixels[key]

    proof = obj(REPO/args.proof); dirty = obj(WORK/"T19-dirty-visual-review.json")
    feature = obj(WORK/"T19-C1-feature-review.json")
    check(proof["commit_under_test"] == C1 and proof["dirty_worktree"] is False, "Exact root C1 proof scope")
    check(proof["visual_comparison"]["prior_visual_review_sha256"] == sha((WORK/"T19-dirty-visual-review.json").read_bytes()), "Prior actual root visual receipt bound")
    check(proof["visual_comparison"]["no_new_interactive_or_certification_claim"] is True, "Mechanical C1 comparison retains limited visual class")
    groups = dirty["reviewed_groups"]
    check(len(groups) == 10 and len(dirty["coverage"]) == 48, "Root actual inspected ten-group/48 dirty coverage scope")
    representatives = {}
    for group in groups:
        actual = pixel(group["representative"])
        check(actual["PixelSHA256"] == group["pixel_sha256"] and actual["Width"] == group["width"] and actual["Height"] == group["height"], "Decoded root representative pixels/dimensions")
        representatives[group["representative"]] = actual
    expected = {}
    native_paths = [row["Path"] for row in feature["RawBindings"] if row["Path"].endswith("/feature-observations.json")]
    check(len(native_paths) == 2, "Two exact C1 native observation sources")
    for native_path in native_paths:
        native = obj(REPO/native_path)
        for observation in native["Observations"]:
            if observation["Label"] == "original-corpus": records = observation["Originals"]
            elif observation["Label"] in ["master-only", "screen", "ebook"]: records = [observation["Master"]] + ([observation["Email"]] if observation["Email"] else [])
            else: continue
            for record in records:
                for render in record["Snapshot"]["renders"]:
                    expected[render["path"]] = dict(Shell=native["ShellVersion"], PDFSHA256=record["Snapshot"]["file"]["sha256"], Identifier=render["identifier"], PNGSHA256=render["sha256"], Width=render["pixel_width"], Height=render["pixel_height"])
    coverage = proof["visual_comparison"]["coverage"]
    check(len(coverage) == len(expected) == 48 and {row["png_path"] for row in coverage} == set(expected), "Exact complete clean native page coverage")
    for row in coverage:
        fact = expected[row["png_path"]]; actual = pixel(row["png_path"])
        check(actual["PNGSHA256"] == row["png_sha256"] == fact["PNGSHA256"], "Clean PNG exact byte/source binding")
        check(row["pdf_sha256"] == fact["PDFSHA256"] and row["shell"] == fact["Shell"] and row["identifier"] == fact["Identifier"], "Clean source PDF/shell/page identity")
        check(actual["PixelSHA256"] == row["pixel_sha256"] and actual["Width"] == fact["Width"] and actual["Height"] == fact["Height"], "Clean decoded pixel/dimension observation")
        previous = representatives.get(row["previously_viewed_representative"])
        check(previous is not None and all(actual[key] == previous[key] for key in ["PixelSHA256", "Width", "Height"]), "Clean full-page image equals previously inspected representative")
    check(len({pixel(row["png_path"])["PixelSHA256"] for row in coverage}) == 10, "Ten unique clean full-page pixel groups")
    for row in dirty["coverage"]:
        actual = pixel(row["png_path"])
        check(actual["PNGSHA256"] == row["png_sha256"] and actual["PixelSHA256"] == row["pixel_sha256"], "Dirty coverage retained pixel/PNG binding")
    report = dict(SchemaVersion=1, Task="T19", Result="pass" if not failures else "fail", CommitUnderTest=C1,
        ObservedAtUtc=dt.datetime.now(dt.timezone.utc).isoformat(), CheckCount=checks, BlockingFindings=failures,
        CleanRenderedPages=len(coverage), PreviouslyInspectedGroups=len(groups),
        SourceBindings=list(bindings.values()), IgnoredPNGByteBindings=list(pngs.values()),
        ProducerSHA256=sha(Path(__file__).read_bytes()), PythonSHA256=PYTHON, PillowVersion=PIL.__version__,
        Scope="Independent exact PNG/RGB/dimension/page identity binding of root's C1 mechanical comparison to root's previously completed dirty visual observations",
        Limits=["Existing PNG decoding only; no PDF parser/native application/renderer/suite execution or image authoring.",
                "This reviewer did not use view_image or make new visual/manual/interactive/Explorer claims; root authored the underlying scoped observations.",
                "Pixel equality establishes identity to those viewed page images, not editable forms, structure, signatures, accessibility or universal fidelity.",
                "Only current T19 proof, native feature receipts and direct dirty visual support were read; no prior-task evidence recursion."])
    output.write_text(json.dumps(report, indent=2)+"\n", encoding="utf-8")
    print(json.dumps(dict(Result=report["Result"], CheckCount=checks, Report=output.relative_to(REPO).as_posix(), ReportSHA256=sha(output.read_bytes()), BlockingFindings=failures)))
    return 0 if not failures else 1


if __name__ == "__main__": sys.exit(main())
