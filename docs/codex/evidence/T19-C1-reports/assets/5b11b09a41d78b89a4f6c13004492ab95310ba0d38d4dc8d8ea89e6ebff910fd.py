"""Independent retained T19 feature/capture/docs audit; no native/PDF renderer run."""
import argparse
from collections import Counter
import datetime as dt
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys
import uuid
import xml.etree.ElementTree as ET

REPO = Path(__file__).resolve().parents[2]
WORK = REPO / "tests/.work"
BASE = "50220eccd1917e44a52d94cbfe3bf35b1940f3d8"
START = "7e0fe6fde769911f323bd87e4c4f9382d26e331b"
PYTHON = "dd5f8d19f6755d6491ee7c4bef2fe35ddd521334cc3ca3ed8fc93ebcadf135d0"
VERSIONS = dict(python="3.12.14", pypdf="6.10.0", pypdfium2="5.13.0",
                pdfium="153.0.7999.0", pillow="12.3.0",
                pdfium_dll_sha256="958e5342ed7e2e20fb914adde238bbae0ac8ad4a3267aa49d0b9dd266c7667f2")


def sha(data): return hashlib.sha256(data).hexdigest()


def git(*args): return subprocess.check_output(["git", *args], cwd=REPO, text=True).strip()


class Review:
    def __init__(self):
        self.checks = 0; self.failures = []; self.bindings = {}; self.binaries = {}; self.outcomes = []

    def check(self, value, label):
        self.checks += 1
        if not value: self.failures.append(label)

    def read(self, path, binary=False):
        path = Path(path).resolve()
        assert path.is_relative_to(REPO), "Only current task repository files may be read"
        data = path.read_bytes()
        row = dict(Path=path.relative_to(REPO).as_posix(), SHA256=sha(data), Bytes=len(data))
        (self.binaries if binary else self.bindings)[row["Path"]] = row
        return data

    def obj(self, path): return json.loads(self.read(path).decode("utf-8-sig"))

    def text(self, path): return self.read(path).decode("utf-8-sig")

    def bound(self, path, digest, binary=False):
        data = self.read(path, binary)
        self.check(sha(data) == digest.lower(), "Exact retained bytes: " + Path(path).name)
        return data

    def capture(self, row):
        out = self.bound(row["Stdout"], row["StdoutSHA256"])
        self.bound(row["Stderr"], row["StderrSHA256"])
        self.check(row["ExitCode"] == 0, "Actual child exit zero: " + row["Label"])
        self.check(Path(row["Executable"]).is_file(), "Selected executable still exists: " + row["Label"])
        return out.decode("utf-8-sig")

    def snapshot(self, record, label, shell):
        raw = self.bound(record["SnapshotPath"], record["SnapshotSHA256"])
        snap = json.loads(raw.decode("utf-8-sig"))
        self.check(snap == record["Snapshot"], "Embedded snapshot equals exact standalone snapshot: " + label)
        pdf = self.bound(record["Pdf"], snap["file"]["sha256"], binary=True)
        self.check(len(pdf) == snap["file"]["bytes"], "Actual PDF length: " + label)
        self.check(snap["file"]["path"] == record["Pdf"], "Exact PDF operand: " + label)
        self.check(snap["versions"] == VERSIONS, "Pinned structural/native reader versions: " + label)
        self.check(snap["parser"] == {"strict": True, "warnings": []}, "Strict no-warning receipt: " + label)
        self.check(snap["page_count"] == len(snap["pages"]) == len(snap["page_identifiers"]), "Reader page/count agreement: " + label)
        for page in snap["pages"]:
            self.check(snap["page_identifiers"][page["index"]] == page["identifier"], "Page identifier mapping: " + label)
            self.check(page["pdfium"]["rotation_degrees"] == page["structural"]["rotation_degrees"], "Raw/native rotation agreement: " + label)
        self.check(len(snap["renders"]) == snap["page_count"], "One retained render per actual page: " + label)
        for render in snap["renders"]:
            png = self.bound(render["path"], render["sha256"], binary=True)
            self.check(len(png) == render["bytes"] and png.startswith(b"\x89PNG\r\n\x1a\n"), "Bound PNG bytes: " + label)
            self.check(render["dpi"] == 144 and render["draw_forms"] and render["draw_annots"], "Declared render settings: " + label)
        self.check(snap["signature_observations"]["validation_performed"] is False, "No signature validation claim: " + label)
        self.check(snap["signature_observations"]["xfa_present"] is False, "No invented XFA sample: " + label)
        active = snap["active_content_observations"]
        self.check(active["uris_followed"] is False and active["javascript_platform_provided"] is False and active["xfa_render_disabled"] is True,
                   "Inert reader boundary: " + label)
        forms = snap["forms"]
        refs = {field["object_ref"]: field for field in forms["canonical_fields"]}
        widgets = {widget["object_ref"]: widget for widget in forms["widgets"]}
        for relation in forms["relations"]:
            widget = widgets[relation["widget_ref"]]; field = refs.get(relation["canonical_field_ref"])
            self.check(relation["canonical_in_field_tree"] == (field is not None), "Canonical membership fact: " + label)
            self.check(relation["widget_in_canonical_kids"] == (field is not None and widget["object_ref"] in field["widget_refs"]), "Kids association fact: " + label)
            self.check(relation["parent_chain_reaches_canonical"] == (field is not None and (field["object_ref"] in widget["parent_chain"] or field["object_ref"] == widget["object_ref"])), "Parent association fact: " + label)
            self.check(relation["page_association_matches"] == (widget["page_index"] == widget["declared_page_index"]), "Declared P association fact: " + label)
            self.check(relation["effective_value_matches_canonical"] == (field is not None and widget["effective_value"] == field["value"]), "Effective field value fact: " + label)
        outcome = self.summarize(snap, shell, label)
        self.outcomes.append(outcome)
        return snap, outcome

    def summarize(self, snap, shell, label):
        annots = [a for page in snap["pages"] for a in page["annotations"]]
        links = [dict(SourcePage=page["index"], URI=a["uri"], TargetPage=(a["destination"] or {}).get("page_index"),
                      TargetIdentifier=(a["destination"] or {}).get("page_identifier"))
                 for page in snap["pages"] for a in page["annotations"] if a["subtype"] == "/Link"]
        attachments = [dict(Name=a["file_attachment"]["filename"],
                            Payloads={k: v.get("decoded_sha256") for k, v in a["file_attachment"]["streams"].items()})
                       for a in annots if "file_attachment" in a]
        return dict(Shell=shell, Label=label, Bytes=snap["file"]["bytes"], PageIdentifiers=snap["page_identifiers"],
            CanonicalFields=[dict(Name=f["full_name"], Value=f["value"], Type=f["field_type"], Widgets=len(f["widget_refs"])) for f in snap["forms"]["canonical_fields"]],
            Widgets=[dict(Page=w["page_index"], Name=w["full_name"], EffectiveValue=w["effective_value"],
                          AppearancePresent=w["appearance"]["present"], AppearanceText=(w["appearance"]["normal"] or {}).get("text_literals")) for w in snap["forms"]["widgets"]],
            FormRelations=[{k: v for k, v in r.items() if not k.endswith("_ref")} for r in snap["forms"]["relations"]],
            DuplicateCanonicalNames=snap["forms"]["duplicate_canonical_names"], FormIssues=snap["forms"]["issues"],
            Bookmarks=snap["bookmarks"], NamedDestinations=snap["named_destinations"], Links=links,
            AnnotationCounts=dict(sorted(Counter(a["subtype"] for a in annots).items())),
            DocumentEmbeddedFiles=[dict(Name=f["name"], Filename=f["filename"], Payloads={k:v.get("decoded_sha256") for k,v in f["streams"].items()}) for f in snap["embedded_files"]],
            PageAttachments=attachments, Tagging=dict(Marked=snap["tagging"]["marked"], RootPresent=snap["tagging"]["root_present"],
                NodeCount=len(snap["tagging"]["nodes"]), ParentEntries=len(snap["tagging"]["parent_tree_entries"]),
                MCIDs=len(snap["tagging"]["mcid_associations"]), MatchedMCIDs=sum(a["element_in_structure_tree"] and a["element_content_reference_matches"] for a in snap["tagging"]["mcid_associations"]),
                Issues=snap["tagging"]["issues"]),
            Geometry=[dict(Rotate=p["structural"]["rotation_degrees"], MediaBox=p["structural"]["media_box"], PdfiumSize=p["pdfium"]["size_points"]) for p in snap["pages"]])


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", action="append", required=True)
    parser.add_argument("--original-receipt", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    review = Review(); output = Path(args.output).resolve()
    assert output.is_relative_to(WORK) and not output.exists(), "Use a fresh ignored report"
    assert sha(Path(sys.executable).read_bytes()) == PYTHON
    assert git("rev-parse", "HEAD") == BASE and not git("status", "--porcelain=v1"), "Exact clean C1 required"
    source_dir = WORK / ("T19-C1-feature-review-sources-" + uuid.uuid4().hex); source_dir.mkdir()
    sources = ["README.md", "docs/PDF_LIMITATIONS.md", "tests/fixtures/README.md", "WinPDFMerge.ps1",
               "WinPDFMerge.bat", "src/WinPDFMerge.Helpers.ps1", "tools/test/Invoke-Tests.ps1",
               "tests/pdf/Preservation.Native.Tests.ps1", "tests/help/PreservationDocs.Tests.ps1",
               "tests/fixtures/features/generate_features.py", "tests/fixtures/features/manifest.json",
               "tools/test/feature_oracle.py", "tools/test/tests/test_feature_oracle.py"]
    source_rows = []
    for index, relative in enumerate(sources):
        raw = review.read(REPO/relative); snapshot = source_dir / f"{index:02d}-{Path(relative).name}"
        snapshot.write_bytes(raw); review.bound(snapshot, sha(raw))
        source_rows.append(dict(Path=relative, SHA256=sha(raw), RetainedSource=snapshot.relative_to(REPO).as_posix()))
    runtime_diff = git("diff", START, "--", "WinPDFMerge.ps1", "WinPDFMerge.bat", "src/WinPDFMerge.Helpers.ps1")
    review.check(not runtime_diff, "Application/helper/BAT byte semantics unchanged versus starting commit")
    manifest = review.obj(REPO/"tests/fixtures/features/manifest.json")
    review.check(manifest["generator_sha256"] == sha((REPO/manifest["generator"]).read_bytes()), "Generator source bound to manifest")
    expected_ids = manifest["expected_merged_page_identifiers"]
    shell_outcomes = {}
    for root_label in args.root:
        root = (REPO/root_label).resolve(); meta = review.obj(root/"metadata.json")
        runs = review.obj(root/"runs.json"); summary = review.obj(root/"PreservationNative.summary.json")
        aggregate = review.obj(root/"aggregate.json"); xml = ET.fromstring(review.read(root/"PreservationNative.results.xml"))
        review.check(meta["commit_under_test"] == BASE and meta["dirty_worktree"] is False, "Clean C1 native baseline accurately scoped")
        review.check(summary["passed"] == summary["total"] == 6 and all(summary[k] == 0 for k in ["failed", "failed_blocks", "failed_containers", "skipped", "not_run"]), "Actual clean C1 six-case Pester summary")
        cases = list(xml.iter("test-case"))
        review.check(len(cases) == 6 and all(c.attrib.get("success") == "True" and c.attrib.get("executed") == "True" for c in cases), "Six actual executed successful NUnit cases")
        for binding in meta["sources"]: review.bound(binding["retained_source"], binding["sha256"])
        row = next(r for r in runs if r["tier"] == "PreservationNative"); review.check(row["exit_code"] == 0 and row["summary"] == summary, "Actual driver/summary binding")
        stdout = review.bound(row["stdout"], row["stdout_sha256"]).decode("utf-8-sig")
        review.bound(row["stderr"], row["stderr_sha256"])
        matches = re.findall(r"^Preservation native observations: (.+)$", stdout, re.M)
        review.check(len(matches) == 1, "Unique actual native observation marker")
        observation = review.obj(matches[0].strip()); shell = observation["ShellVersion"]
        review.check(shell == summary["shell_version"] and shell in ("5.1.26100.9444", "7.6.6"), "Actual approved shell")
        capture_by_label = {c["Label"]: c for c in observation["Captures"]}
        review.check(len(capture_by_label) == 11, "Generation, seven reads and three unchanged entry captures")
        captured_stdout = {}
        for capture in observation["Captures"]: captured_stdout[capture["Label"]] = review.capture(capture)
        work = Path(observation["Work"])
        for relative in ["WinPDFMerge.ps1", "src/WinPDFMerge.Helpers.ps1"]:
            review.check(review.read(work/"application"/relative) == (REPO/relative).read_bytes(), "Unmodified copied application: " + relative)
        for relative in ["WinPDFMerge.ps1", "src/WinPDFMerge.Helpers.ps1", "tests/pdf/Preservation.Native.Tests.ps1", "tools/test/feature_oracle.py", "tests/fixtures/features/generate_features.py"]:
            review.check(review.read(work/Path(relative).name) == (REPO/relative).read_bytes(), "Retained executing source snapshot: " + relative)
        orig_row = next(o for o in observation["Observations"] if o["Label"] == "original-corpus")
        originals = []
        for record, fixture in zip(orig_row["Originals"], manifest["fixtures"]):
            snap, outcome = review.snapshot(record, fixture["file"], shell); originals.append(snap)
            review.check(json.loads(captured_stdout[fixture["file"].replace(".pdf", "")+"-oracle"]) == snap, "Actual original reader stdout equals snapshot")
            review.check(snap["file"]["sha256"] == fixture["sha256"] and snap["file"]["bytes"] == fixture["bytes"], "Exact reproducible original")
            review.check(outcome["CanonicalFields"] == [dict(Name="shared_text", Value=fixture["canonical_field"]["value"], Type="/Tx", Widgets=2)], "Distinct original canonical name/value")
            review.check(len(outcome["Widgets"]) == 2 and all(w["AppearanceText"] == [fixture["canonical_field"]["value"]] for w in outcome["Widgets"]), "Original widget appearance separate from canonical value")
            review.check(outcome["Tagging"]["MatchedMCIDs"] == 2, "Original minimal MCID relationships")
        runs_by_label = {o["Label"]: o for o in observation["Observations"] if o["Label"] in ["master-only", "screen", "ebook"]}
        local_outcomes = []
        for label, run in runs_by_label.items():
            review.check(run["ExitCode"] == 0, "Actual application route exit zero: " + label)
            log = review.bound(run["Log"], run["LogSHA256"]).decode("utf-8-sig")
            review.check("Expected page total: 4" in log and "Result: SUCCESS; exit code: 0" in log, "Runtime actual page count and outcome: " + label)
            review.check(" cat " not in log or '"cat"' in log, "Native vector serialization remains literal: " + label)
            review.check('"compress" "dont_ask"' in log and '"flatten"' not in log, "PDFtk vector contains compress, no flatten option: " + label)
            child = capture_by_label[label]
            review.check(child["Arguments"][1:4] == ["-ExecutionPolicy", "RemoteSigned", "-File"], "Entry child uses documented scoped policy: " + label)
            review.check(child["Arguments"][-len(run["Options"]):] == run["Options"] if run["Options"] else True, "Actual entry options: " + label)
            review.check(not list(Path(run["Output"]).glob(".WinPDFMerge_*.tmp")), "Owned staging cleaned after route: " + label)
            review.bound(Path(run["Output"])/"foreign-existing.txt", run["ForeignSHA256"])
            for kind in ["Master", "Email"]:
                record = run[kind]
                if record is None: continue
                snap, outcome = review.snapshot(record, label+"-"+kind.lower(), shell)
                review.check(json.loads(captured_stdout[label+"-"+kind.lower()+"-oracle"]) == snap, "Actual output reader stdout equals snapshot")
                review.check(snap["page_identifiers"] == expected_ids, "Actual output expected page identifiers")
                for link in outcome["Links"]:
                    if link["TargetPage"] is not None:
                        expected = 1 if link["SourcePage"] < 2 else 3
                        review.check(link["TargetPage"] == expected and link["TargetIdentifier"] == expected_ids[expected], "Resolved local merged offset")
                    else: review.check(link["URI"] == ("https://example.invalid/T19/A" if link["SourcePage"] < 2 else "https://example.invalid/T19/B"), "Exact inert URI observation")
                for attached in outcome["PageAttachments"]:
                    fixture = next(f for f in manifest["fixtures"] if f["attachment"]["name"] == attached["Name"])
                    review.check(all(digest == fixture["attachment"]["sha256"] for digest in attached["Payloads"].values()), "Decoded original attachment payload hash")
                local_outcomes.append(outcome)
            if label == "master-only": review.check(run["Email"] is None and "Email result: skipped" in log and "Ghostscript version probe" not in log, "True skip route")
            else:
                review.check(run["Email"]["Snapshot"]["file"]["bytes"] < run["Master"]["Snapshot"]["file"]["bytes"] and "Email result: published" in log, "Actual strictly smaller final derivative")
                for flag in ["-dSAFER", "-dPDFSTOPONERROR", "-dCompatibilityLevel=1.6", "-dDetectDuplicateImages=true", "-dPDFSETTINGS=/"+label]:
                    review.check(flag in log, "Exact retained Ghostscript flag: " + flag)
        preservation = next(o for o in observation["Observations"] if o["Label"] == "preservation")
        review.check(preservation["SourceBefore"] == preservation["SourceAfter"] and preservation["ParentEnvironmentPreserved"] and preservation["CopiedApplicationBytesIdentical"], "Retained original/timestamp/environment assertions")
        for item in preservation["SourceAfter"]: review.bound(work/"source"/item["Name"], item["SHA256"])
        shell_outcomes[shell] = [{k:v for k,v in outcome.items() if k != "Shell"} for outcome in local_outcomes]
    review.check(set(shell_outcomes) == {"5.1.26100.9444", "7.6.6"}, "Both required shell receipts")
    review.check(shell_outcomes["5.1.26100.9444"] == shell_outcomes["7.6.6"], "Measured feature/size/geometry outcomes agree across shells")
    original = review.obj(REPO/args.original_receipt)
    review.check(original["result"] == "pass" and not original["partial"] and original["check_count"] == 54 and all(c["pass"] for c in original["checks"]), "Independent original receipt scope and actual checks")
    for row in original["results"]:
        review.bound(row["stdout"], row["stdout_sha256"]); review.bound(row["stderr"], row["stderr_sha256"])
        raw = review.bound(row["snapshot"], row["snapshot_sha256"]); snap = json.loads(raw)
        review.bound(snap["file"]["path"], row["source_before_sha256"], binary=True)
        review.check(row["source_before_sha256"] == row["source_after_sha256"] and row["exit_code"] == 0, "Original read-only receipt preserves source")
        for render in row["renders"]: review.bound(render["path"], render["sha256"], binary=True)

    total_cases = 0; total_reports = 0
    for root_label in args.root:
        root = (REPO/root_label).resolve(); runs = review.obj(root/"runs.json")
        review.check(len(runs) == 6, "Six completed scoped C1 tiers per actual host")
        subtotal = 0
        for row in runs:
            tier = row["tier"]; summary = review.obj(root/(tier+".summary.json"))
            review.check(row["exit_code"] == 0 and row["summary"] == summary, "C1 driver actual exit/summary: " + tier)
            review.check(summary["commit_under_test"] == BASE and summary["dirty_worktree"] is False, "C1 exact clean summary: " + tier)
            review.check(summary["passed"] == summary["total"] and all(summary[k] == 0 for k in ["failed", "failed_blocks", "failed_containers", "skipped", "not_run"]), "C1 all actual cases pass: " + tier)
            xml = ET.fromstring(review.read(root/(tier+".results.xml"))); cases = list(xml.iter("test-case"))
            review.check(len(cases) == summary["total"] and all(c.attrib.get("success") == "True" and c.attrib.get("executed") == "True" for c in cases), "C1 actual NUnit count/outcome: " + tier)
            review.bound(row["stdout"], row["stdout_sha256"]); review.bound(row["stderr"], row["stderr_sha256"])
            subtotal += summary["total"]; total_reports += 1
        aggregate = review.obj(root/"aggregate.json")
        review.check(aggregate["passed"] == subtotal == 414 and aggregate["tiers"] == 6 and aggregate["bad_counts"] == 0 and aggregate["dirty_worktree"] is False and aggregate["commit_under_test"] == BASE, "C1 aggregate exactly matches actual reports")
        total_cases += subtotal
        docs_summary = review.obj(root/"PreservationDocs.summary.json")
        review.check(docs_summary["total"] == 14, "Actual C1 documentation/help cases per host")
        docs_stdout = review.text(root/"PreservationDocs.stdout.txt")
        markers = re.findall(r"^Preservation documentation receipts: (.+)$", docs_stdout, re.M)
        review.check(len(markers) == 1, "Unique C1 documentation receipt marker")
        docs = review.obj(markers[0].strip())
        review.check(docs["ShellVersion"] == docs_summary["shell_version"] and len(docs["Observations"]) == 14, "Actual fourteen help/documentation observations")
        for binding in docs["SourceBindings"]:
            raw = review.bound(binding["RetainedSourcePath"], binding["SHA256"])
            review.check(raw == (REPO/binding["Path"]).read_bytes(), "Exact C1 documentation source: " + binding["Path"])
        help_binding = docs["HelpBinding"]
        actual_help = review.bound(help_binding["RetainedTextPath"], help_binding["SHA256"]).decode("utf-8-sig")
        review.check(help_binding["ApplicationInvoked"] is False and help_binding["Command"][0] == "Get-Help", "Real comment help extraction without application invocation")
        review.check("Neither output guarantees PDF/A" in actual_help and "Keep original documents" in actual_help, "Actual rendered help contains conservative limits")
    review.check(total_cases == 828 and total_reports == 12, "Actual C1 twelve scoped reports total828")
    review.check(git("rev-parse", "HEAD") == BASE and not git("status", "--porcelain=v1"), "Exact C1 still clean after read-only review")
    limits = review.text(REPO/"docs/PDF_LIMITATIONS.md"); readme = review.text(REPO/"README.md")
    fixture_readme = review.text(REPO/"tests/fixtures/README.md"); entry = review.text(REPO/"WinPDFMerge.ps1")
    review.check(not re.search(r"\blossless\b|archive.safe", readme, re.I), "README removes all unqualified lossless/archive-safe claims")
    review.check("--output tests/.work/features" in fixture_readme and "generate_features.py --output-dir" not in fixture_readme, "Fixture README command matches argparse/native actual command")
    for term in ["1.shared_text", "Document entries absent", "Struct", "signature preservation is unvalidated", "XFA", "not a malware sanitizer", "not a representative compression", "144 DPI"]:
        review.check(term.lower() in limits.lower(), "Scoped measured/general documentation statement: " + term)
    review.check("Neither output guarantees PDF/A, signature validity, universal feature retention" in entry, "Real source comment help remains conservative")
    review.check("docs/PDF_LIMITATIONS.md" in readme, "README routes users to feature limits")
    for row in source_rows: review.check(sha((REPO/row["Path"]).read_bytes()) == row["SHA256"], "Current reviewed source remains stable: " + row["Path"])
    report = dict(SchemaVersion=1, Task="T19", Result="pass" if not review.failures else "fail",
        CommitUnderTest=BASE, CommitUnderReviewedNative=BASE, NativeWorktree="clean", CurrentHeadAtReview=git("rev-parse", "HEAD"),
        ObservedAtUtc=dt.datetime.now(dt.timezone.utc).isoformat(), CheckCount=review.checks, BlockingFindings=review.failures,
        EvidenceClass="independent retained structural/native-capture and current documentation review",
        ExecutedCaseCount=total_cases, ExecutedReportCount=total_reports,
        SourceBindings=source_rows, ProducerSHA256=sha(Path(__file__).read_bytes()), RawBindings=list(review.bindings.values()),
        IgnoredBinaryBindings=list(review.binaries.values()), Outcomes=review.outcomes,
        DocumentationReview=dict(Result="pass" if not review.failures else "fail",
            Surfaces=["README.md", "docs/PDF_LIMITATIONS.md", "tests/fixtures/README.md", "WinPDFMerge.ps1 real comment help"],
            Readiness="No blocking findings in C1 actual structural/docs/native receipts; public privacy/closure audit remains pending",
            VisualStatements="Attributed to root's separate actual view inspection; this reviewer did not view images or independently certify those visual statements"),
        Limits=["No application, suite, PDF authoring, native reader or renderer rerun by this reviewer; retained PDF/PNG bytes were hashed read-only.",
            "Native snapshots originate from the separately authored strict-pypdf/PDFium oracle; independently reviewed source and recomputed record associations/targets/payload comparisons, not a second parser implementation.",
            "Own source research/authorship only; application, corpus/oracle and Pester tests authored by other agents.",
            "Initial exploratory Get-Content used a nonexistent PreservationNative.txt basename twice; actual stdout.txt paths then enumerated/read. No application failure or fabricated report followed.",
            "Only actual T19 six-tier C1 reports, fourteen documentation and six feature-native cases per shell; no interactive GUI, physical Explorer, signature, XFA, accessibility, malware or archival certification claim.",
            "Only current T19 task-local raw records/captures and current source/docs inspected; no recursive prior-task evidence graph."])
    output.write_text(json.dumps(report, indent=2, ensure_ascii=False)+"\n", encoding="utf-8")
    print(json.dumps(dict(Result=report["Result"], CheckCount=review.checks, Report=output.relative_to(REPO).as_posix(),
                         ReportSHA256=sha(output.read_bytes()), BlockingFindings=review.failures,
                         RawBindings=len(review.bindings), BinaryBindings=len(review.binaries))))
    return 0 if not review.failures else 1


if __name__ == "__main__": sys.exit(main())
