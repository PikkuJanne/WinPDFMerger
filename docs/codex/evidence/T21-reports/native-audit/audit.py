"""Independent T21 raw-result reader; no application launches or tracked edits.

Author also wrote corpus.py; independent from the safety suite and root driver.
Uses retained raw NUnit/JSON/observations and fresh PDFium reads only.
"""
from __future__ import annotations
import argparse
import datetime
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import subprocess
import xml.etree.ElementTree as ET

REPO = Path(__file__).resolve().parents[3]
C1 = "abf8976e84f2c3f851efc42a844037a880519b26"
TIERS = ["Unit", "SourceDiscovery", "InputPreflight", "MasterValidation", "Staging", "Destination",
         "SizeReportingNative", "PreservationNative", "ParametersNative", "DiagnosticsNative", "CorpusSafety"]
BAD = ["failed", "failed_blocks", "failed_containers", "skipped", "not_run"]
UTC = datetime.timezone.utc
TICKS_EPOCH = 621355968000000000


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def exact(first, second):
    return json.dumps(first, sort_keys=True) == json.dumps(second, sort_keys=True)


def require(condition, message):
    if not condition:
        raise ValueError(message)


def snapshot_check(raw):
    rows = read_snapshot = json.loads(raw)
    require(isinstance(rows, list) and rows, "Empty snapshot cannot establish preservation.")
    recorded_paths = {Path(row["Path"]) for row in rows}
    roots = [path for path in recorded_paths if not any(parent in recorded_paths for parent in path.parents)]
    actual_paths = set()
    for root in roots:
        actual_paths.add(root)
        if root.is_dir():
            actual_paths.update(root.rglob("*"))
    require(actual_paths == recorded_paths, "Current complete tree has extra/missing objects.")
    for row in read_snapshot:
        path = Path(row["Path"])
        require(path.exists(), "Retained source/foreign/final object is missing.")
        stat = path.stat()
        require(path.is_dir() == (row["Kind"] == "directory"), "Object type changed.")
        require(stat.st_file_attributes == row["Attributes"], "File/directory attributes changed.")
        birth_ns = getattr(stat, "st_birthtime_ns", stat.st_ctime_ns)
        require(TICKS_EPOCH + birth_ns // 100 == row["CreatedUtcTicks"], "Creation timestamp changed.")
        if row["Kind"] == "file":
            require(stat.st_size == row["Length"] and digest(path) == row["SHA256"], "Retained file bytes changed.")
            require(TICKS_EPOCH + stat.st_mtime_ns // 100 == row["ModifiedUtcTicks"], "File modification timestamp changed.")
    return len(rows)


def native_audit(path):
    report = read(path)
    require(report["CommitUnderTest"] == C1 and report["DirtyWorktree"] is False, "Native observations are not clean C1.")
    require(report["Process64Bit"] is True and report["StandardUser"] is True, "Native Windows standard-user facts missing.")
    require(digest(report["CorpusReceipt"]) == report["CorpusReceiptSHA256"], "Corpus receipt hash mismatch.")
    require(digest(REPO / "tools/test/corpus.py") == report["CorpusToolSHA256"], "Corpus tool hash mismatch.")
    environment = read(REPO / "tests/.work/T21-environment.json")
    expected_engines = {Path(row["path"]).name: row["sha256"] for row in environment["approved_selected_files"]
                        if Path(row["path"]).name in ("pdftk.exe", "libiconv2.dll", "gswin64c.exe", "gsdll64.dll")}
    require(exact({row["Name"]: row["SHA256"] for row in report["EngineHashes"]}, expected_engines), "Native engine hash facts differ from selected-byte inventory.")
    require(report["Generation"]["ExitCode"] == 0, "Corpus generation failed.")
    spec = importlib.util.spec_from_file_location("fresh_corpus_audit", REPO / "tools/test/corpus.py")
    oracle = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(oracle)
    oracle.versions()
    observations = report["Observations"]
    require(len(observations) == 21, "Unexpected native observation coverage.")
    fresh = []
    snapshots = 0
    concurrent = []
    outcomes = []
    for observation in observations:
        label = observation["Label"]
        scenario = observation["Expected"]
        source_root = Path(json.loads(observation["SourceBefore"])[0]["Path"])
        for category in ("Source", "Foreign"):
            require(exact(json.loads(observation[category + "Before"]), json.loads(observation[category + "After"])), label + ": preservation snapshots differ.")
            snapshots += snapshot_check(observation[category + "After"])
        if observation["FinalSnapshots"]:
            snapshots += snapshot_check(observation["FinalSnapshots"])
        extra = observation.get("Extra") or {}
        for before, after in (("PriorOutputsBefore", "PriorOutputsAfter"), ("OutputBefore", "OutputAfter")):
            if before in extra:
                require(exact(json.loads(extra[before]), json.loads(extra[after])), label + ": protected outputs differ.")
                snapshots += snapshot_check(extra[after])
        results = observation["Result"] if isinstance(observation["Result"], list) else [observation["Result"]]
        expected_code = 1 if label.startswith(("whole-set-refusal-", "overlap-", "only-documented-exclusions-")) else 0
        require(all(row["ExitCode"] == expected_code for row in results), label + ": exit differs.")
        log = observation["Log"]
        if expected_code == 1 and observation["FinalSnapshots"]:
            final_rows = json.loads(observation["FinalSnapshots"])
            require(len(final_rows) == 1 and Path(final_rows[0]["Path"]).suffix == ".log", label + ": partial PDF was published.")
            require("Master validation OK" not in log and "Published Master:" not in log and "Ghostscript arguments:" not in log, label + ": partial success advertised.")
        if observation["LogPath"] and not label.startswith("concurrent-"):
            output = Path(observation["LogPath"]).parent
            permitted_rows = json.loads(observation["ForeignAfter"]) + json.loads(observation["FinalSnapshots"])
            if "PriorOutputsAfter" in extra:
                permitted_rows += json.loads(extra["PriorOutputsAfter"])
            # A later repeat adds its own immutable outputs to this same folder;
            # include every recorded final from that exact destination.
            for later in observations:
                if later["LogPath"] and Path(later["LogPath"]).parent == output and later["FinalSnapshots"]:
                    permitted_rows += json.loads(later["FinalSnapshots"])
            permitted = {row["Path"] for row in permitted_rows if Path(row["Path"]).parent == output}
            require({str(item) for item in output.iterdir()} == permitted, label + ": final output inventory has omissions/residue.")
        if observation["Oracle"]:
            require(expected_code == 0 and log, label + ": no successful log.")
            require(Path(observation["LogPath"]).read_bytes().decode("utf-8-sig") == log, label + ": log bytes differ from observation.")
            actual_names = [Path(value).name for value in re.findall(r"(?m)^Input [0-9]+: (.+)\r?$", log)]
            actual_names = [name.rstrip("\r") for name in actual_names]
            require(actual_names == scenario["ordered_names"], label + ": numbered order differs.")
            counts = [int(value) for value in re.findall(r"(?m)^Input [0-9]+ pages: ([0-9]+)\r?$", log)]
            by_name = {item["path"]: item for item in scenario["source_files"]}
            require(counts == [by_name[name]["page_count"] for name in scenario["ordered_names"]], label + ": individual page counts differ.")
            require(sum(counts) == scenario["expected_page_count"], label + ": total pages differ.")
            require(f"Expected page total: {scenario['expected_page_count']}" in log, label + ": expected total log missing.")
            for kind in ("Master", "Email"):
                recorded = observation["Oracle"].get(kind)
                if not recorded:
                    continue
                require(recorded["ExitCode"] == 0, label + ": recorded oracle failed.")
                arguments = recorded["Arguments"]
                pdf = Path(arguments[arguments.index("--pdf") + 1])
                current = oracle.inspect_pdf(pdf, scenario["expected_page_identifiers"])
                require(exact(current, recorded["Inspection"]), label + ": fresh PDFium differs from retained inspection.")
                require(current["page_count"] == scenario["expected_page_count"], label + ": fresh count differs.")
                fresh.append({"label": label, "kind": kind, "pdf": str(pdf), "sha256": digest(pdf),
                    "pages": current["page_count"], "identifiers": current["page_identifiers"]})
        elif label.startswith("whole-set-refusal-"):
            rejected = [item["path"] for item in scenario["source_files"] if item["disposition"] == "rejected"]
            require(len(rejected) == 1 and rejected[0] in results[0]["Stdout"] + results[0]["Stderr"], label + ": rejected operand not named.")
            names = [Path(value.rstrip("\r")).name for value in re.findall(r"(?m)^Input [0-9]+: (.+)\r?$", log)]
            require(names == scenario["ordered_names"] and "PDFtk failed during input preflight" in log, label + ": bad input silently omitted.")
        if label.startswith("concurrent-"):
            intervals = extra["EntryIntervals"]
            require(len(intervals) == 2 and intervals[0]["ProcessId"] != intervals[1]["ProcessId"], "Concurrency needs two exact children.")
            require(all(item["StartUtcTicks"] < item["EndUtcTicks"] for item in intervals), "Invalid entry interval.")
            require(max(item["StartUtcTicks"] for item in intervals) < min(item["EndUtcTicks"] for item in intervals), "Actual entry intervals did not overlap.")
            require(all(item["StartUtcTicks"] >= extra["ReleasedAtUtcTicks"] for item in intervals), "Entry began before barrier release.")
            require(not Path(extra["Stage"]).exists(), "Owned concurrent stage remains.")
            concurrent.append(observation)
        outcomes.append({"label": label, "exit_code": expected_code, "scenario": observation["Scenario"]})
    require(len(concurrent) == 2, "Missing either concurrent output observation.")
    logs = [Path(item["LogPath"]) for item in concurrent]
    require(len({log.stem for log in logs}) == 2 and len({item["Extra"]["Stage"] for item in concurrent}) == 2, "Concurrent identities/stages reused.")
    expected_union = set()
    for item in concurrent:
        expected_union.update(row["Path"] for row in json.loads(item["FinalSnapshots"]))
    foreign = {row["Path"] for row in json.loads(concurrent[0]["ForeignAfter"])}
    output = logs[0].parent
    actual_union = {str(path) for path in output.iterdir() if str(path) not in foreign}
    require(actual_union == expected_union, "Concurrent publication union differs or has residue.")
    return {"observation_receipt": str(path), "observation_sha256": digest(path), "observations": len(observations),
        "snapshot_objects_rechecked": snapshots, "fresh_pdfium_inspections": len(fresh), "fresh_pdfs": fresh,
        "concurrency": {"entry_intervals": concurrent[0]["Extra"]["EntryIntervals"], "overlap_ticks": min(item["EndUtcTicks"] for item in concurrent[0]["Extra"]["EntryIntervals"]) - max(item["StartUtcTicks"] for item in concurrent[0]["Extra"]["EntryIntervals"]),
            "unique_run_identities": 2, "unique_stages": 2, "publication_union_files": len(expected_union), "owned_residue": 0}, "outcomes": outcomes}


def host_audit(root):
    metadata = read(root / "metadata.json")
    aggregate = read(root / "aggregate.json")
    runs = read(root / "runs.json")
    guard = read(root / "source-guard.json")
    require(metadata["commit_under_test"] == C1 and metadata["dirty_worktree"] is False, "Driver did not test clean C1.")
    require(metadata["verified_inventory_sha256"] == digest(REPO / "tests/.work/T21-environment.json"), "Environment inventory changed after execution.")
    require(aggregate["result"] == "pass" and aggregate["tiers"] == 11 and aggregate["bad_counts"] == 0, "Missing complete eleven-tier aggregate.")
    require([row["tier"] for row in runs] == TIERS, "Tier coverage/order differs.")
    require(guard["result"] == "pass" and len(guard["bindings"]) == len(metadata["sources"]), "Incomplete source guard.")
    require(exact({item["path"]: item["sha256"] for item in metadata["sources"]}, {item["path"]: item["before_sha256"] for item in guard["bindings"]}), "Source guard does not bind initial metadata.")
    for row in guard["bindings"]:
        require(row["before_sha256"] == row["after_sha256"] == digest(REPO / row["path"]), "Guarded implementation changed.")
    audited = []
    native = None
    for run in runs:
        tier = run["tier"]
        summary_path = root / (tier + ".summary.json")
        xml_path = root / (tier + ".results.xml")
        summary = read(summary_path)
        require(exact(summary, run["summary"]) and exact(summary, read(Path(run["report"]) / "summary.json")), tier + ": summary copies differ.")
        require(xml_path.read_bytes() == (Path(run["report"]) / "results.xml").read_bytes(), tier + ": raw NUnit copy differs.")
        require(run["exit_code"] == 0 and summary["passed"] == summary["total"] > 0 and all(summary[key] == 0 for key in BAD), tier + ": failed/skipped/unrun/block counts.")
        require(summary["commit_under_test"] == C1 and summary["dirty_worktree"] is False and summary["pester_version"] == "6.2.0", tier + ": source/runtime binding differs.")
        require(summary["process_64_bit"] is True and summary["execution_policy"] == "RemoteSigned", tier + ": shell facts differ.")
        require(summary["shell_version"] == ("5.1.26100.9444" if metadata["shell"] == "ps51" else "7.6.6"), tier + ": shell pin differs.")
        for stream in ("stdout", "stderr"):
            require(digest(run[stream]) == run[stream + "_sha256"], tier + ": stream hash differs.")
        xml = ET.parse(xml_path).getroot()
        require(int(xml.attrib["total"]) == summary["total"], tier + ": NUnit total differs.")
        require(all(int(xml.attrib[key]) == 0 for key in ("errors", "failures", "not-run", "inconclusive", "ignored", "skipped", "invalid")), tier + ": NUnit bad count nonzero.")
        cases = list(xml.iter("test-case"))
        require(len(cases) == summary["passed"] and all(case.attrib.get("executed") == "True" and case.attrib.get("success") == "True" and case.attrib.get("result") == "Success" for case in cases), tier + ": raw test cases disagree.")
        require(all(node.attrib.get("success") == "True" for node in xml.iter("test-suite")), tier + ": raw suite failure.")
        audited.append({"tier": tier, "passed": len(cases), "summary_sha256": digest(summary_path), "nunit_sha256": digest(xml_path), "all_bad_counts": 0})
        if tier == "CorpusSafety":
            receipts = [item[1] for item in run["observation_receipts"] if item[0] == "Corpus safety receipts"]
            require(len(receipts) == 1, "Exactly one safety receipt required.")
            native = native_audit(Path(receipts[0]) / "native-observations.json")
    require(sum(item["passed"] for item in audited) == aggregate["passed"], "Host aggregate total differs.")
    return {"shell": metadata["shell"], "root": str(root), "guarded_sources": len(guard["bindings"]),
        "passed": aggregate["passed"], "tiers": audited, "native": native,
        "bindings": {name: digest(root / name) for name in ("metadata.json", "aggregate.json", "runs.json", "source-guard.json", "driver.py")}}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, action="append", required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    require(not args.output.exists() and args.output.resolve().is_relative_to(REPO / "tests/.work"), "Audit output must be a new owned ignored file.")
    require(subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO).decode().strip() == C1, "Current HEAD differs from guarded C1.")
    environment_path = REPO / "tests/.work/T21-environment.json"
    environment = read(environment_path)
    for row in environment["approved_selected_files"]:
        require(digest(row["path"]) == row["sha256"], "Selected external dependency bytes changed.")
    dll = [row for row in environment["approved_selected_files"] if Path(row["path"]).name == "pdfium.dll"]
    require(len(dll) == 1 and dll[0]["sha256"] == "524ecbe6a7d49103909b1ed39fe512d2d4e612e35dac1336c9274371d20c5d90", "Fresh PDFium DLL differs from the exact observed T21 bundle.")
    result = {"task": "T21", "result": "pass", "commit_under_test": C1, "observed_at_utc": datetime.datetime.now(UTC).isoformat(),
              "author_disclosure": "Auditor also authored corpus.py; this reader is independent from the native safety suite and root driver.",
              "environment_inventory_sha256": digest(environment_path), "external_dependency_files_rehashed": len(environment["approved_selected_files"]), "pdfium_dll_sha256": dll[0]["sha256"],
              "audit_script_sha256": digest(__file__), "oracle_sha256": digest(REPO / "tools/test/corpus.py"), "hosts": [host_audit(root.resolve()) for root in args.root],
              "scope": "Independent raw-result/source guard/snapshot/log/order/concurrency audit and fresh retained-PDF reads. No merges rerun; no desktop/visual/signature/accessibility certification."}
    with args.output.open("x", encoding="utf-8") as stream:
        stream.write(json.dumps(result, indent=2) + "\n")
    print(json.dumps({"result": "pass", "output": str(args.output), "hosts": [{"shell": host["shell"], "passed": host["passed"], "tiers": len(host["tiers"]), "observations": host["native"]["observations"], "fresh_inspections": host["native"]["fresh_pdfium_inspections"]} for host in result["hosts"]]}))


if __name__ == "__main__":
    main()
