"""Ignored T14 collector. Default is read-only; only --write creates public evidence.

All inputs are checked before writes. Clean acceptance, dirty focused reports,
standalone native smoke reports and controlled scheduling remain explicitly
distinct. This collector runs no PDF engines or test suites.
"""
from pathlib import Path
import argparse
import json
import os
import re
import runpy
import subprocess

LEGACY_PATH = Path(__file__).with_name("Collect-T09C3Evidence.py")
LEGACY = runpy.run_path(str(LEGACY_PATH), run_name="t14_report_helpers")
LegacyCollector = LEGACY["Collector"]
require, sha, load_json, json_bytes = (LEGACY[key] for key in ("require", "sha", "load_json", "json_bytes"))
COUNTS = {
    "Unit": 335, "EmailOutcome": 11, "MasterValidation": 7, "Staging": 9, "InputPreflight": 22, "Destination": 15,
    "ToolInvocation": 12, "PdftkPaths": 13, "GhostscriptPaths": 13,
    "SourceDiscovery": 4, "DependencyEntry": 9, "Launcher": 24, "LauncherNative": 2,
}
TIERS = tuple(COUNTS)
LegacyCollector.clean_shell.__globals__["TIERS"] = TIERS
COMMIT = "588a96586518059ed38e2de01ad2312003defc7c"


def assert_clean(repo, commit):
    head = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"], check=True,
                          capture_output=True, text=True).stdout.strip()
    dirty = subprocess.run(["git", "-C", str(repo), "status", "--porcelain=v1"], check=True,
                           capture_output=True, text=True).stdout
    require(head == commit and not dirty, "Collection requires the exact clean T14 C1 implementation SHA")


class T14Collector(LegacyCollector):
    def __init__(self, repo, commit):
        super().__init__(repo, commit)
        self.standalone = {}
        self.native = {}
        self.oracle_runtime = None
        self.counts_digest = None

    def record_for(self, label, tier):
        return next(row for row in self.records if row.get("shell") == label and row.get("tier") == tier)

    def clean_shell(self, label, root):
        aggregate = super().clean_shell(label, root)
        require(aggregate["counts_per_tier"] == COUNTS and aggregate["passed"] == 476,
                "Frozen T14 tier/count mismatch")
        for tier, version in (("PdftkPaths", "2.02"), ("GhostscriptPaths", "10.08.0")):
            require(self.record_for(label, tier)["native_version"] == version, "Native path-suite engine version pin mismatch")
        root = self.owned_input(root)
        metadata_raw, metadata = load_json(root / "collector.json")
        require(metadata["task"] == "T14" and metadata["checkpoint"] == "C1" and
                metadata["commit_under_test"] == self.commit and metadata["dirty_worktree"] is False and
                metadata["shell"] == label and metadata["acquisition_performed"] is False,
                "T14 collector metadata ownership/context mismatch")
        require(metadata["expected_counts_sha256"] == sha(self.owned_input(metadata["expected_counts_file"]).read_bytes())
                == self.counts_digest, "Executed collector used different expected counts")
        require(metadata["orchestration_script_sha256"] == sha((self.work / "Run-T14Checkpoint.ps1").read_bytes()),
                "Executed collector source changed")
        name = label + "-collector.json"
        cleaned = json_bytes(self.sanitize_value(metadata))
        self.add_payload(name, cleaned)
        aggregate.update(collector_file=name, collector_raw_sha256=sha(metadata_raw), collector_sha256=sha(cleaned))
        _, jobs = load_json(root / "runs.json")
        for job in jobs:
            tier = job["tier"]
            require(job["native_test_host_started"] is True and job["timed_out"] is False and
                    job["capture_error"] is None and job["termination_error"] is None,
                    "Owned test host did not complete cleanly")
            require(job["executable"] == metadata["driver"], "Test-host executable differs from collector metadata")
            arguments = job["arguments"]
            expected = ["-NoProfile", "-ExecutionPolicy", "RemoteSigned", "-File",
                        str(self.repo / "tools/test/Invoke-Tests.ps1"), "-PesterModulePath",
                        metadata["pester_manifest"], "-Tier", tier]
            if tier in ("EmailOutcome", "MasterValidation", "Staging", "InputPreflight", "Destination", "PdftkPaths", "GhostscriptPaths",
                        "SourceDiscovery", "DependencyEntry", "LauncherNative"):
                expected += ["-PdftkPath", metadata["pdftk"]]
            if tier in ("EmailOutcome", "Staging", "InputPreflight", "Destination", "GhostscriptPaths"):
                expected += ["-GhostscriptPath", metadata["ghostscript"]]
            if tier in ("EmailOutcome", "InputPreflight", "MasterValidation"):
                expected += ["-PythonPath", metadata["python"]]
            require(arguments == expected, "Unexpected clean test-host argument vector")
            row = self.record_for(label, tier)
            row.update(elapsed_ms=job["elapsed_ms"], started_at_utc=job["started_at_utc"],
                       completed_at_utc=job["completed_at_utc"], command_executable=self.sanitize_string(job["executable"]),
                       command_arguments=self.sanitize_value(arguments))
            for stream, key in (("stdout", "log"), ("stderr", "stderr_log")):
                raw = self.owned_input(job[key]).read_bytes()
                clean = self.sanitize_string(raw.decode("utf-8-sig")).encode("utf-8")
                log_name = f"{label}-{tier}-{stream}.txt"
                self.add_payload(log_name, clean)
                row.update({f"collector_{stream}_raw_sha256": sha(raw), f"collector_{stream}_file": log_name,
                            f"collector_{stream}_sha256": sha(clean)})
            if tier in ("EmailOutcome", "MasterValidation", "Staging", "Destination", "InputPreflight"):
                observations = self.observations(label, tier, job, aggregate["shell_version"], metadata)
                self.native[(label, tier)] = observations
            elif tier in ("PdftkPaths", "GhostscriptPaths"):
                self.path_observations(job, metadata)
        return aggregate

    def snapshot(self, value, current=False):
        require(isinstance(value, str) and value, "Missing nonempty synthetic file snapshot")
        rows = [json.loads(line) for line in value.splitlines() if line.strip()]
        require(rows and len({row["Path"] for row in rows}) == len(rows), "Ambiguous snapshot paths")
        for row in rows:
            require(re.fullmatch(r"[0-9a-fA-F]{64}", row["SHA256"]) and row["Length"] > 0 and
                    row["ModifiedUtcTicks"] > 0, "Invalid synthetic file snapshot metadata")
            if current:
                path = self.owned_input(row["Path"])
                raw = path.read_bytes()
                require(len(raw) == row["Length"] and sha(raw) == row["SHA256"].lower(),
                        "Retained source/final/sentinel bytes differ from recorded snapshot")
        return {row["Path"]: row for row in rows}

    def native_result(self, result, executable, success=True):
        native = result["NativeResult"]
        require(native and native["Started"] is True and native["Executable"] == executable and native["ProcessId"] > 0,
                "Expected actual selected engine execution was not recorded")
        require(native["Succeeded"] is success and ((native["ExitCode"] == 0) is success),
                "Actual engine result/exit mismatch")
        if success:
            self.complete_native(native)
            require(result["Succeeded"] is True and result["OutputPublished"] is True and not result["OutputError"],
                    "Real successful native result was not published truthfully")
            if Path(executable).name.lower() == "pdftk.exe":
                self.validated_master(result, executable)
            else:
                self.validated_email(result, self.selected_pdftk)
        return native

    @staticmethod
    def complete_native(native):
        require(native and native["Started"] is True and native["Succeeded"] is True and native["ExitCode"] == 0 and
                native["ProcessId"] > 0 and native["TimedOut"] is False and native["Cancelled"] is False and
                not native["LaunchError"] and not native["CaptureError"] and not native["TerminationError"] and
                native["StdoutTruncated"] is False and native["StderrTruncated"] is False,
                "Actual native receipt does not establish complete bounded successful capture")

    def inspected_output(self, job, executable, expected, staged_name):
        validation = job["ValidationResult"]
        require(validation and validation["Succeeded"] is True and not validation["InputError"] and
                validation["PageCount"] == expected > 0, "Master validation result/page count mismatch")
        native = validation["NativeResult"]
        self.complete_native(native)
        require(native["Executable"] == executable and native["ProcessId"] != job["NativeResult"]["ProcessId"],
                "Validation is not a distinct actual selected PDFtk process")
        labels = re.findall(r"(?m)^NumberOfPages:[^\r\n]*", native["Stdout"])
        require(len(labels) == 1 and re.fullmatch(r"NumberOfPages:[ \t]*[0-9]+[ \t]*", labels[0]) and
                int(labels[0].split(":", 1)[1].strip()) == expected <= 9223372036854775807,
                "Master validation native stdout does not contain one strict positive matching Int64 label")
        require("dump_data_utf8" in native["RenderedArguments"] and staged_name in native["RenderedArguments"],
                "Validation does not identify the read-only staged inspection")
        return validation

    def inspected_master(self, job, executable, expected):
        return self.inspected_output(job, executable, expected, "master.pdf")

    def validated_master(self, job, executable):
        count = job["ValidatedPageCount"]
        require(job["OutputValidated"] is True and isinstance(count, int) and count > 0,
                "Successful PDFtk result lacks explicit positive validated master state")
        self.inspected_master(job, executable, count)

    def validated_email(self, job, inspector):
        count = job["ValidatedPageCount"]
        require(job["OutputValidated"] is True and isinstance(count, int) and count > 0,
                "Successful Ghostscript result lacks explicit positive validated state")
        self.inspected_output(job, inspector, count, "email.pdf")
        require(isinstance(job["MasterBytes"], int) and job["MasterBytes"] > 0 and
                isinstance(job["OutputBytes"], int) and job["OutputBytes"] > 0,
                "Email size decision lacks exact nonzero byte counts")
        if job["OutputState"] == "published":
            require(job["OutputPublished"] is True and job["OutputBytes"] < job["MasterBytes"],
                    "Published email does not have a strict size benefit")
        else:
            require(job["OutputState"] == "no_size_benefit" and job["OutputPublished"] is False and
                    job["OutputBytes"] >= job["MasterBytes"] and not job["OutputError"],
                    "No-benefit email was misclassified or advertised")

    def padding_receipt(self, preparation, path):
        require(preparation and preparation["AddedWhitespaceBytes"] == 4096 and preparation["Scope"],
                "Synthetic owned-master preparation lacks exact disclosed padding provenance")
        raw = self.owned_input(path).read_bytes()
        require(raw.endswith(b" " * 4096), "Retained prepared master padding changed")
        if "BeforeSHA256" in preparation:
            require(sha(raw[:-4096]) == preparation["BeforeSHA256"].lower() and
                    sha(raw) == preparation["AfterSHA256"].lower(), "Owned-master preparation digest mismatch")
        else:
            require(sha(raw) == preparation["PreparedMasterSHA256"].lower(),
                    "Concurrent prepared master digest changed")

    @staticmethod
    def cleaned(result):
        require(result["Cleaned"] is True and not result["CleanupError"] and not result["OrphanPath"],
                "Native owned staging did not clean successfully")

    def staging_observations(self, observations, metadata):
        self.selected_pdftk = metadata["pdftk"]
        rows = observations["Observations"]
        expected = {"real-tools-one-owned-stage", "preexisting-final-Pdftk", "preexisting-final-Ghostscript",
                    "actual-pdftk-bad-input-owned-cleanup", "actual-same-second-live-concurrent-stage-isolation"}
        expected.update(f"actual-move-{kind}-collision-after-real-{engine}"
                        for kind in ("file", "directory") for engine in ("Pdftk", "Ghostscript"))
        require(len(rows) == 9 and {row["Label"] for row in rows} == expected,
                "Staging requires the nine exact native publication/concurrency observations")
        require(observations["Process64Bit"] is True, "Staging actual host is not x64")
        expected_engine_hashes = {}
        for receipt_name, entries_key in (("T03-pdftk-acquisition.json", "extracted_files"),):
            _, receipt = load_json(self.repo / "docs/codex/evidence" / receipt_name)
            expected_engine_hashes.update({Path(item["relative_path"]).name: item["sha256"] for item in receipt[entries_key]})
        _, gs = load_json(self.repo / "docs/codex/evidence/T09-gs-acquisition.json")
        expected_engine_hashes.update({Path(item["relative_path"]).name: item["sha256"]
                                      for item in gs["ghostscript_extraction"]["selected_files"]})
        require({row["Name"]: row["SHA256"] for row in observations["EngineSHA256"]} == expected_engine_hashes,
                "Staging actual vendor executable/interpreter hashes differ from approved receipts")
        for row in rows:
            if "Before" in row:
                require(row["Before"] == row["After"], "Staging changed source/existing final snapshot")
                self.snapshot(row["After"], current=True)
            if "SourceAndForeignBefore" in row:
                require(row["SourceAndForeignBefore"] == row["SourceAndForeignAfter"],
                        "Staging changed source/foreign snapshots")
                self.snapshot(row["SourceAndForeignAfter"], current=True)
            if "Cleanup" in row:
                self.cleaned(row["Cleanup"])
            label = row["Label"]
            if label == "real-tools-one-owned-stage":
                self.native_result(row["Master"], metadata["pdftk"])
                gs_native = self.native_result(row["Email"], metadata["ghostscript"])
                require("-dSAFER" in gs_native["RenderedArguments"] and "/screen" in gs_native["RenderedArguments"],
                        "Actual email did not retain safety/default flags")
                require(row["Master"]["StagingPath"] == row["Email"]["StagingPath"] == row["Stage"],
                        "Actual engines did not share the same owned stage")
                require(len(self.snapshot(row["FinalOutputs"], current=True)) == 2, "Expected master/email final snapshots")
                self.padding_receipt(row["MasterPreparation"], row["MasterPreparation"]["Path"])
            elif label.startswith("preexisting-final-"):
                require(row["RealNativeStarted"] is False and row["Result"]["NativeResult"] is None and
                        row["Result"]["Succeeded"] is False and row["Result"]["OutputPublished"] is False,
                        "Preexisting final was not refused before actual native work")
            elif label.startswith("actual-move-"):
                result = row["Result"]
                native = result["NativeResult"]
                executable = metadata["pdftk"] if label.endswith("Pdftk") else metadata["ghostscript"]
                require(row["RealNativeStarted"] is True and native["Started"] is True and
                        native["Succeeded"] is True and native["ExitCode"] == 0 and native["Executable"] == executable and
                        result["Succeeded"] is False and result["OutputPublished"] is False and result["OutputError"],
                        "Real native output/final collision result is inconsistent")
                require(re.fullmatch(r"[0-9A-Fa-f]{64}", row["StagedSHA256"]), "Missing real staged collision hash")
                self.complete_native(native)
                if label.endswith("Pdftk"):
                    self.validated_master(result, metadata["pdftk"])
                else:
                    require(result["OutputValidated"] is True and result["OutputBytes"] < result["MasterBytes"],
                            "GS publication collision did not follow a validated strict-smaller decision")
                    self.inspected_output(result, metadata["pdftk"], 3, "email.pdf")
                    self.padding_receipt(row["MasterPreparation"], row["MasterPreparation"]["Path"])
                sentinel = self.owned_input(row["ForeignSentinelPath"])
                require(sha(sentinel.read_bytes()) == row["ForeignFinalSHA256"].lower(), "Foreign final sentinel changed")
            elif label == "actual-pdftk-bad-input-owned-cleanup":
                self.native_result(row["Result"], metadata["pdftk"], success=False)
                require(row["Result"]["Succeeded"] is False and row["Result"]["OutputPublished"] is False,
                        "Invalid native merge advertised success")
            else:
                a, b = row["ReadyA"], row["ReadyB"]
                require(a["Timestamp"] == b["Timestamp"] and a["ProcessId"] != b["ProcessId"] and
                        a["Stage"] != b["Stage"] and a["BaseName"] != b["BaseName"],
                        "Actual same-second children shared run or staging identity")
                require(all(a[key] != b[key] for key in ("Master", "Email", "Log")),
                        "Actual concurrent children shared a final/log path")
                require(row["HeldSecondBefore"] == row["HeldSecondAfterFirstCleanup"],
                        "First cleanup changed the second live child's held staged sentinel")
                require(len(self.snapshot(row["HeldSecondBefore"])) == 1, "Missing held concurrent sentinel snapshot")
                finals = self.snapshot(row["FinalOutputs"], current=True)
                first = self.snapshot(row["FirstFinalsBeforeSecondFinishes"], current=True)
                require(len(finals) == 4 and len(first) == 2 and all(finals[path] == value for path, value in first.items()),
                        "Second run changed first published outputs")
                for result, ready in ((row["ResultA"], a), (row["ResultB"], b)):
                    require(result["ProcessId"] == ready["ProcessId"], "Concurrent result belongs to different child")
                    self.native_result(result["Master"], metadata["pdftk"])
                    self.native_result(result["Email"], metadata["ghostscript"])
                    self.padding_receipt(result["MasterPreparation"], ready["Master"])
                    require(result["Master"]["StagingPath"] == result["Email"]["StagingPath"] == ready["Stage"],
                            "Concurrent result used a foreign stage")
                    self.cleaned(result["Cleanup"])

    def path_observations(self, job, metadata):
        log = self.owned_input(job["log"]).read_text(encoding="utf-8-sig")
        paths = re.findall(r"(?m)^Native observations:\s*(.+?)\r?$", log)
        require(len(paths) == 1, "ToolPaths requires its exact owned receipt")
        _, observations = load_json(self.owned_input(paths[0].strip()))
        rows = observations["Observations"]
        require(len(rows) == 14 and len({row["Label"] for row in rows}) == 14,
                "ToolPaths thirteen cases require fourteen distinct observations")
        _, acquisition = load_json(self.repo / "docs/codex/evidence/T03-pdftk-acquisition.json")
        pdftk_hash = next(item["sha256"] for item in acquisition["extracted_files"] if item["relative_path"].endswith("/pdftk.exe"))
        for row in rows:
            if "Succeeded" not in row:
                require(row["Label"] in ("actual-entry-special-Latin-paths", "actual-encrypted-only-entry-failure-before-GS"),
                        "Unexpected aggregate native path observation shape")
                continue
            if row["Succeeded"]:
                require(row["NativeStarted"] is True and row["ExitCode"] == 0 and row["TimedOut"] is False and
                        row["OutputValidated"] is True and row["ValidatedPageCount"] == 2 and not row["OutputError"] and
                        not row["CleanupError"], "Successful path output lacks truthful native/validated state")
                validation = row["ValidationResult"]
                require(validation["Succeeded"] is True and validation["PageCount"] == 2 and not validation["InputError"],
                        "Path output lacks matching complete inspection")
                self.complete_native(validation["NativeResult"])
                executable = Path(validation["NativeResult"]["Executable"]).resolve()
                require(executable.name.lower() == "pdftk.exe" and sha(executable.read_bytes()) == pdftk_hash,
                        "Tool path inspector differs from selected approved PDFtk bytes")
                if row["OutputState"] == "published":
                    require(row["Published"] is True and self.owned_input(row["OutputPath"]).stat().st_size == row["OutputBytes"],
                            "Published native path output bytes are missing/changed")
                    if job["tier"] == "GhostscriptPaths":
                        require(row["OutputBytes"] < row["MasterBytes"], "Native path email was published without strict size benefit")
                else:
                    require(job["tier"] == "GhostscriptPaths" and row["OutputState"] == "no_size_benefit" and
                            row["Published"] is False and row["OutputBytes"] >= row["MasterBytes"] and
                            not Path(row["OutputPath"]).exists(), "Native path no-benefit outcome is inconsistent")
            elif job["tier"] == "GhostscriptPaths" and row["Label"] == "CJK-output-directory":
                require(row["NativeStarted"] is True and row["ExitCode"] == 0 and row["TimedOut"] is False and
                        row["OutputValidated"] is False and row["Published"] is False and row["OutputState"] == "failed" and
                        row["OutputError"] and not Path(row["OutputPath"]).exists(),
                        "CJK GS conversion/inspection failure was advertised as native success")
                inspection = row["ValidationResult"]
                require(inspection["Succeeded"] is False and inspection["InputError"] and
                        inspection["NativeResult"]["Executable"] == metadata["pdftk"] and
                        inspection["NativeResult"]["Started"] is True and inspection["NativeResult"]["ExitCode"] != 0,
                        "CJK output limitation lacks separate actual selected PDFtk failure")

    def native_log_success(self, log, operation, executable):
        matches = re.findall(r"(?m)^" + re.escape(operation) + r" exit: 0; elapsed: [0-9]+ ms; PID: ([0-9]+)\r?$", log)
        require(len(matches) == 1 and int(matches[0]) > 0, "Entry log lacks one actual successful native PID")
        require(operation + " executable: " + executable in log and
                operation + " started: True; timed out: False; cancelled: False; succeeded: True" in log and
                operation + " stdout truncated: False; stderr truncated: False" in log,
                "Entry log lacks selected executable or complete successful native capture")
        for fault in ("launch", "capture", "termination"):
            require(re.search(r"(?m)^" + re.escape(operation) + " " + fault + r" error:[ \t]*\r?$", log),
                    "Entry log has an omitted or nonempty native fault field")
        return int(matches[0])

    @staticmethod
    def log_page_count(log, operation, count):
        block = re.search(re.escape(operation) + r" stdout:\r?\n(.*?)" + re.escape(operation) + r" stderr:", log, re.DOTALL)
        labels = re.findall(r"(?m)^NumberOfPages:[^\r\n]*", block.group(1) if block else "")
        require(len(labels) == 1 and re.fullmatch(r"NumberOfPages:[ \t]*[0-9]+[ \t]*", labels[0]) and
                int(labels[0].split(":", 1)[1].strip()) == count > 0,
                "Logged inspection stdout does not have one strict matching positive page count")

    def email_observations(self, label, observations, metadata, receipt_path):
        states = {
            "actual-entry-missing-optional-GS": ("unavailable", 0),
            "actual-entry-SkipEmail-with-real-GS-available": ("skipped", 0),
            "actual-entry-real-smaller-screen-derivative": ("published", 0),
            "actual-entry-real-valid-no-size-benefit": ("no_size_benefit", 0),
            "actual-GS-exit1-partial-retained-master-code2": ("failed", 2),
            "actual-GS-exit0-real-inspection-injected-page-mismatch-code2": ("failed", 2),
            "real-master-controlled-discovery-exception-code2": ("failed", 2),
            "real-master-controlled-publication-log-exception-code2": ("failed", 2),
            "actual-cmd-BAT-native-code0": ("no_size_benefit", 0),
            "actual-cmd-BAT-native-code1": ("not_started", 1),
            "actual-cmd-BAT-native-code2": ("failed", 2),
        }
        rows = observations["Observations"]
        require(len(rows) == 11 and {row["Label"] for row in rows} == set(states),
                "EmailOutcome requires its eleven exact native entry/BAT observations")
        require(observations["Process64Bit"] is True and observations["OracleVersions"] ==
                {"python": "3.12.14", "pypdfium2": "5.13.0", "pdfium": "153.0.7999.0"}, "Email oracle/runtime pins mismatch")
        _, pdftk = load_json(self.repo / "docs/codex/evidence/T03-pdftk-acquisition.json")
        _, gs = load_json(self.repo / "docs/codex/evidence/T09-gs-acquisition.json")
        expected_hashes = {Path(item["relative_path"]).name: item["sha256"] for item in pdftk["extracted_files"]}
        expected_hashes.update({Path(item["relative_path"]).name: item["sha256"] for item in gs["ghostscript_extraction"]["selected_files"]})
        require({item["Name"]: item["SHA256"] for item in observations["EngineSHA256"]} == expected_hashes,
                "Email engines/interpreters differ from approved selected bytes")
        suite = self.repo / "tests/pdf/EmailOutcome.Native.Tests.ps1"
        suite_raw = suite.read_bytes()
        suite_text = suite_raw.decode("utf-8-sig")
        sources = (
            ("oracle", "independent-email-inspection.py", "OracleSHA256", r"\[IO.File\]::WriteAllText\(\$oracle,@'\r?\n(.*?)\r?\n'@"),
            ("generator", "original-raster.py", "GeneratorSHA256", r"\$generatorSource=@'\r?\n(.*?)\r?\n'@"),
        )
        record = self.record_for(label, "EmailOutcome")
        record.update(native_suite_sha256=sha(suite_raw), independent_final_inspections=[])
        for kind, leaf, digest, pattern in sources:
            raw = self.owned_input(receipt_path.parent / leaf).read_bytes()
            embedded = re.search(pattern, suite_text, re.DOTALL)
            require(sha(raw) == observations[digest].lower() and embedded and
                    embedded.group(1).replace("\r\n", "\n") == raw.decode("utf-8-sig").replace("\r\n", "\n"),
                    "Email retained oracle/generator differs from executed clean-suite source")
            name = f"{label}-EmailOutcome-{kind}.py"
            self.add_payload(name, raw)
            record.update({kind + "_file": name, kind + "_sha256": sha(raw)})
        for row in rows:
            name = row["Label"]
            state, code = states[name]
            require(row["Before"] == row["After"], "Email entry changed source/foreign snapshots")
            sources_snapshot = self.snapshot(row["After"], current=True)
            if name.startswith("actual-cmd-BAT-"):
                require(row["Result"]["ExitCode"] == code and "PowerShell5.1" in row["BatchChildShell"] and row["Scope"],
                        "Actual BAT result/shell/control disclosure mismatch")
            if code == 1:
                require(row["Proof"] is None and "Merge failed with exit code 1" in row["Result"]["Stdout"],
                        "Actual BAT pre-master failure advertised a master")
                output = self.owned_input(next(iter(sources_snapshot))).parent
                require(not list(output.glob("WinPDFMerge_*.pdf")), "Failed actual BAT created an advertised final PDF")
                continue
            proof = row["Proof"]
            require(proof["ExpectedState"] == state and proof["ActualResult"]["ExitCode"] == code,
                    "Actual email entry/result-state mismatch")
            stdout = proof["ActualResult"]["Stdout"]
            require(("PARTIAL SUCCESS:" if code == 2 else "SUCCESS:") in stdout,
                    "Actual entry summary disagrees with explicit complete/partial outcome")
            master = self.owned_input(proof["Master"])
            require(" - Merged master: " + str(master) in stdout, "Published master is missing from explicit final listing")
            require((" - Email-optimized: " in stdout) is (state == "published"),
                    "Actual final listing advertises an unpublished email or omits the published email")
            raw_log = self.owned_input(proof["LogPath"]).read_bytes()
            log = raw_log.decode("utf-8-sig")
            require(log == proof["Log"] and f"Email result: {state}" in log and
                    f"Result: {'PARTIAL SUCCESS' if code == 2 else 'SUCCESS'}; exit code: {code}" in log,
                    "Actual retained entry log/state/summary differ from observation")
            pids = [self.native_log_success(log, "PDFtk", metadata["pdftk"]),
                    self.native_log_success(log, "Master validation", metadata["pdftk"])]
            require(pids[0] != pids[1] and log.index("PDFtk arguments:") < log.index("Master validation arguments:"),
                    "Actual master merge and inspection are not separate ordered processes")
            self.log_page_count(log, "Master validation", 1)
            reads = proof["Oracle"]
            require(isinstance(reads, list) and len(reads) == (2 if state == "published" else 1) and reads[0]["Path"] == str(master),
                    "Independent oracle inventory does not follow explicit published paths")
            expected_id = "T03-14-P01" if "smaller-screen" in name else "T03-01-P01"
            for inspection in reads:
                path = self.owned_input(inspection["Path"])
                require(sha(path.read_bytes()) == inspection["SHA256"].lower() and inspection["ExitCode"] == 0,
                        "Independent published PDF bytes/oracle exit changed")
                actual = inspection["Result"]
                require(actual["page_count"] == 1 and actual["pypdfium2"] == "5.13.0" and actual["pdfium"] == "153.0.7999.0" and
                        actual["pages"] == [{"identifier": expected_id, "rotation_degrees": 0, "size_points": [432, 288]}],
                        "Independent published master/email visible identity, rotation or dimensions differ")
                record["independent_final_inspections"].append({"observation": name, "path": self.sanitize_string(str(path)),
                                                              "sha256": sha(path.read_bytes()), "page_count": 1})
            require(not list(master.parent.glob(".WinPDFMerge_*.tmp")), "Actual entry left an owned staging directory")
            if "FixtureGeneration" in row:
                generation = row["FixtureGeneration"]
                require(generation["seed"] == 140032 and generation["pixel_dimensions"] == [1200, 800] and
                        generation["page_size_points"] == [432, 288] and generation["visible_id"] == expected_id and
                        generation["pages"] == 1 and generation["provenance"], "Original raster fixture provenance changed")
                require(any(item["SHA256"].lower() == generation["sha256"] and item["Length"] == generation["bytes"]
                            for item in sources_snapshot.values()), "Generated source bytes are not bound to preserved source snapshot")
            capture = proof["EmailCapture"]
            if capture is None:
                require(state in ("unavailable", "skipped", "failed") and not re.search(r"(?m)^Ghostscript arguments:", log) and
                        not re.search(r"(?m)^Email validation arguments:", log), "No-conversion outcome nevertheless launched email work")
                if state == "skipped":
                    require(row["ControlledSentinel"] and not (master.parent.parent / "discovery-reached.txt").exists(),
                            "SkipEmail discovery sentinel was reached or its control was not disclosed")
                elif state == "failed":
                    require(row["ControlledFault"] and "T14 controlled" in log, "Post-master failure lacks disclosed controlled provenance")
                continue
            require(capture["OriginalMasterBefore"] == capture["OriginalMasterAfter"] and
                    capture["OriginalMasterAfter"]["Path"] == str(master), "Optional processing changed its published master snapshot")
            self.snapshot(json.dumps(capture["OriginalMasterAfter"]), current=True)
            job = capture["Job"]
            require(job["OutputState"] == state and job["NativeResult"]["Executable"] == metadata["ghostscript"] and
                    job["NativeResult"]["Started"] is True and job["NativeResult"]["ProcessId"] > 0 and
                    "-dSAFER" in job["NativeResult"]["RenderedArguments"] and "/screen" in job["NativeResult"]["RenderedArguments"] and
                    not job["CleanupError"], "Actual selected email process/state/safety/profile receipt mismatch")
            require(not Path(capture["StagedEmailPath"]).exists() and not Path(capture["StageDirectory"]).exists(),
                    "Actual entry did not clean its known staged derivative")
            if state in ("published", "no_size_benefit"):
                self.complete_native(job["NativeResult"])
                require(job["Succeeded"] is True and not job["OutputError"] and
                        job["MasterBytes"] == capture["OriginalMasterAfter"]["Length"], "Complete email decision lacks stable original master bytes")
                self.validated_email(job, metadata["pdftk"])
                gs_pid = self.native_log_success(log, "Ghostscript", metadata["ghostscript"])
                inspection_pid = self.native_log_success(log, "Email validation", metadata["pdftk"])
                require(gs_pid != inspection_pid and len(set(pids + [gs_pid, inspection_pid])) == 4,
                        "Master and email operations do not have distinct actual process identities")
                self.log_page_count(log, "Email validation", 1)
                require(log.index("Master validation OK:") < log.index("Ghostscript arguments:") < log.index("Email validation arguments:"),
                        "Actual optional work was not ordered after validated master publication")
                if state == "published":
                    require(job["OutputPath"] == reads[1]["Path"] and self.owned_input(job["OutputPath"]).stat().st_size == job["OutputBytes"],
                            "Published smaller derivative path/byte count differ from independent inspection")
                else:
                    require(not Path(job["OutputPath"]).exists() and "no size benefit" in log.lower(),
                            "No-benefit derivative was advertised or silently omitted")
            else:
                require(job["Succeeded"] is False and job["OutputValidated"] is False and job["OutputPublished"] is False and
                        job["OutputError"] and not Path(job["OutputPath"]).exists(), "Failed real email was advertised or published")
                if capture["Mode"] == "corrupt":
                    native = job["NativeResult"]
                    require(native["Succeeded"] is False and native["ExitCode"] == 1 and native["TimedOut"] is False and
                            native["Cancelled"] is False and not native["LaunchError"] and not native["CaptureError"] and not native["TerminationError"] and
                            native["StdoutTruncated"] is False and native["StderrTruncated"] is False and job["ValidationResult"] is None and
                            capture["StagedPartialExistsBeforeEntryCleanup"] is True and capture["StagedPartialBytes"] > 0 and
                            re.fullmatch(r"[0-9a-fA-F]{64}", capture["StagedPartialSHA256"]),
                            "Actual GS nonzero partial-output evidence is incomplete")
                    substituted = self.owned_input(capture["ActualNativeInput"][0])
                    require(str(substituted) != str(master) and substituted.stat().st_size == job["MasterBytes"] and
                            (row.get("ControlledInput") or row.get("Scope")), "Corrupt-input scheduling was not explicitly separated from preserved master")
                else:
                    require(capture["Mode"] == "wrong-count" and row["ControlledCount"], "Unknown actual email failure control")
                    self.complete_native(job["NativeResult"])
                    self.inspected_output(job, metadata["pdftk"], 1, "email.pdf")
                    require("expected 2, inspected 1" in job["OutputError"], "Injected expected-count mismatch changed")

    def master_observations(self, label, observations, metadata):
        rows = observations["Observations"]
        expected_labels = {"actual-entry-single-multipage-validated", "actual-entry-natural-five-pages-mixed-rotation",
                           "real-merge-and-inspection-wrong-expected-count"}
        expected_labels.update("real-native-success-controlled-staged-" + kind for kind in ("missing", "empty", "non-PDF", "wrong-page"))
        require(len(rows) == 7 and {row["Label"] for row in rows} == expected_labels,
                "Master validation requires its exact seven native observations")
        require(observations["Process64Bit"] is True and observations["OracleVersions"] ==
                {"python": "3.12.14", "pypdfium2": "5.13.0", "pdfium": "153.0.7999.0"},
                "Master oracle/runtime pins mismatch")
        _, acquisition = load_json(self.repo / "docs/codex/evidence/T03-pdftk-acquisition.json")
        require({item["Name"]: item["SHA256"] for item in observations["EngineSHA256"]} ==
                {Path(item["relative_path"]).name: item["sha256"] for item in acquisition["extracted_files"]},
                "Master actual engine/interpreter bytes differ from approved acquisition")
        oracle_path = self.owned_input(observations["OraclePath"])
        oracle_raw = oracle_path.read_bytes()
        require(sha(oracle_raw) == observations["OracleSHA256"].lower(), "Retained master oracle digest mismatch")
        suite_path = self.repo / "tests/pdf/MasterValidation.Native.Tests.ps1"
        suite_raw = suite_path.read_bytes()
        embedded = re.search(r"\$oracleSource = @'\r?\n(.*?)\r?\n'@", suite_raw.decode("utf-8-sig"), re.DOTALL)
        require(embedded and embedded.group(1).replace("\r\n", "\n") == oracle_raw.decode("utf-8-sig").replace("\r\n", "\n"),
                "Master oracle does not match the clean native-suite embedded source")
        oracle_name = label + "-MasterValidation-oracle.py"
        self.add_payload(oracle_name, oracle_raw)
        self.record_for(label, "MasterValidation").update(oracle_file=oracle_name, oracle_sha256=sha(oracle_raw),
                                                        oracle_native_suite_sha256=sha(suite_raw), oracle_expectations=[])
        rotations_seen = set()
        entry_number = 0
        for row in rows:
            require(row["SourceAndForeignBefore"] == row["SourceAndForeignAfter"], "Master operation changed source/foreign snapshots")
            self.snapshot(row["SourceAndForeignAfter"], current=True)
            if row["Label"].startswith("actual-entry-"):
                entry_number += 1
                expected = row["ExpectedPages"]
                require(isinstance(expected, list) and len(expected) == (2 if "single" in row["Label"] else 5),
                        "Unexpected master expected visible-page inventory")
                success = row["Success"]
                require(row["Entry"]["ExitCode"] == 0 and "SUCCESS:" in row["Entry"]["Stdout"],
                        "Actual master entry did not return/advertise complete requested success")
                master = self.owned_input(success["Master"])
                require(sha(master.read_bytes()) == success["MasterSHA256"].lower(), "Retained master bytes changed")
                actual_log = self.owned_input(success["LogPath"]).read_bytes().decode("utf-8-sig")
                require(actual_log == success["Log"], "Retained master log differs from captured observation")
                process_ids = []
                for operation in ("PDFtk", "Master validation"):
                    processes = re.findall(r"(?m)^" + operation + r" exit: 0; elapsed: [0-9]+ ms; PID: ([0-9]+)\r?$", actual_log)
                    require(len(processes) == 1 and int(processes[0]) > 0, "Master log missing one actual successful native PID")
                    process_ids.append(int(processes[0]))
                    require(operation + " started: True; timed out: False; cancelled: False; succeeded: True" in actual_log and
                            operation + " stdout truncated: False; stderr truncated: False" in actual_log,
                            "Master log does not establish complete native execution/capture")
                    for fault in ("launch", "capture", "termination"):
                        require(re.search(r"(?m)^" + operation + " " + fault + r" error:[ \t]*\r?$", actual_log),
                                "Master entry native log contains or omits a fault result")
                require(process_ids[0] != process_ids[1], "Entry master merge and validation did not use separate actual processes")
                data = re.search(r"Master validation stdout:\r?\n(.*?)Master validation stderr:", actual_log, re.DOTALL)
                labels = re.findall(r"(?m)^NumberOfPages:[^\r\n]*", data.group(1) if data else "")
                require(len(labels) == 1 and re.fullmatch(r"NumberOfPages:[ \t]*[0-9]+[ \t]*", labels[0]) and
                        int(labels[0].split(":", 1)[1].strip()) == len(expected), "Entry validation stdout count mismatch")
                merge_at = actual_log.index("PDFtk arguments:")
                inspection_at = actual_log.index("Master validation arguments:")
                published_at = actual_log.index("Master validation OK:")
                require(merge_at < inspection_at < published_at and "dump_data_utf8" in actual_log[inspection_at:published_at] and
                        f"Master validation OK: {len(expected)} expected pages inspected. Merged master published:" in actual_log,
                        "Entry merge/validation/publication log ordering is inconsistent")
                require("Ghostscript not found; skipping email-optimized copy." in actual_log and
                        not re.search(r"(?m)^Ghostscript arguments:", actual_log), "Owned master-only entry unexpectedly launched email work")
                oracle = success["Oracle"]
                expectation_path = self.owned_input(oracle["ExpectationPath"])
                expectation_raw, expectation = load_json(expectation_path)
                require(sha(expectation_raw) == oracle["ExpectationSHA256"].lower() and expectation == expected,
                        "Retained master oracle expectation digest/inventory mismatch")
                actual = oracle["Result"]
                require(oracle["ExitCode"] == 0 and not oracle["Stderr"] and actual["page_count"] == len(expected) and
                        actual["pages"] == expected and actual["pypdfium2"] == "5.13.0" and actual["pdfium"] == "153.0.7999.0",
                        "Independent actual master identifiers/rotation/size differ from expectations")
                for page in actual["pages"]:
                    degrees = page["rotation_degrees"]
                    require(degrees in (0, 90, 180, 270) and page["rotation_quarter_turns"] * 90 == degrees and
                            page["size_points"] == ([288, 432] if degrees in (90, 270) else [432, 288]) and
                            re.fullmatch(r"T03-[0-9]{2}-P[0-9]{2}", page["identifier"]),
                            "Independent visible-page units/dimensions/identifier contract mismatch")
                    rotations_seen.add(degrees)
                expected_name = f"{label}-MasterValidation-expectation-{entry_number}.json"
                self.add_payload(expected_name, expectation_raw)
                self.record_for(label, "MasterValidation")["oracle_expectations"].append(
                    {"file": expected_name, "sha256": sha(expectation_raw), "page_count": len(expected),
                     "native_merge_process_id": process_ids[0], "native_inspection_process_id": process_ids[1]})
                for creation in row.get("OwnedFixtureCreation", []):
                    self.complete_native(creation["NativeResult"])
                    require(creation["NativeResult"]["Executable"] == metadata["pdftk"] and
                            sha(self.owned_input(creation["Path"]).read_bytes()) == creation["SHA256"].lower(),
                            "Owned rotated fixture actual execution/hash mismatch")
            else:
                job = row["Job"]
                self.complete_native(job["NativeResult"])
                require(job["NativeResult"]["Executable"] == metadata["pdftk"] and job["Succeeded"] is False and
                        job["OutputValidated"] is False and job["OutputPublished"] is False and
                        job["ValidatedPageCount"] is None and job["OutputError"] and not job["CleanupError"],
                        "Rejected actual master was incorrectly advertised/validated/published")
                final = Path(job["OutputPath"]).resolve()
                stage = Path(job["StagingPath"]).resolve()
                require(final.is_relative_to(self.work) and stage.is_relative_to(self.work) and not final.exists() and not stage.exists(),
                        "Rejected actual master final or owned stage remains unexpectedly")
                if row["Label"] == "real-merge-and-inspection-wrong-expected-count":
                    require(row["ExpectedPageCount"] == 3 and row["ActualPageCount"] == 2, "Expected native count mismatch characterization changed")
                    self.inspected_master(job, metadata["pdftk"], 2)
                elif row["Label"].endswith("wrong-page"):
                    self.inspected_master(job, metadata["pdftk"], 1)
                    require(row["ControlledScheduling"] and re.fullmatch(r"[0-9A-Fa-f]{64}", row["ProducedStagedSHA256"]),
                            "Controlled staged substitution lacks actual original staged digest/provenance")
                else:
                    require(job["ValidationResult"]["Succeeded"] is False and job["ValidationResult"]["NativeResult"] is None and
                            job["ValidationResult"]["InputError"] and row["ControlledScheduling"] and
                            re.fullmatch(r"[0-9A-Fa-f]{64}", row["ProducedStagedSHA256"]),
                            "Missing/empty/non-PDF staged substitution was not refused before inspection launch")
        require(entry_number == 2 and rotations_seen == {0, 90, 180, 270}, "Master oracle did not demonstrate all four rotations")

    def observations(self, label, tier, job, shell_version, metadata):
        log = self.owned_input(job["log"]).read_text(encoding="utf-8-sig")
        prefixes = {"EmailOutcome": "Email", "MasterValidation": "Master", "Staging": "Staging", "Destination": "(?:Native|Destination)",
                    "InputPreflight": "(?:Native|Input preflight)"}
        matches = re.findall(r"(?m)^" + prefixes[tier] + r" observations:\s*(.+?)\r?$", log)
        require(len(matches) == 1, f"{tier} requires one standalone owned observations receipt")
        path = self.owned_input(matches[0].strip())
        raw, observations = load_json(path)
        require(observations["CommitUnderTest"] == self.commit and observations["DirtyWorktree"] is False and
                observations["ShellVersion"] == shell_version and observations["ShellEdition"] ==
                ("Desktop" if label == "ps51" else "Core") and observations["StandardUser"] is True and
                observations["PdfTkVersion"] == "2.02" and
                (tier == "MasterValidation" or observations["GhostscriptVersion"] == "10.08.0") and
                observations["Observations"], f"{tier} native commit/host/vendor pin mismatch")
        if tier == "EmailOutcome":
            self.email_observations(label, observations, metadata, path)
        elif tier == "MasterValidation":
            self.master_observations(label, observations, metadata)
        elif tier == "Staging":
            self.staging_observations(observations, metadata)
        elif tier == "InputPreflight":
            require(len(observations["Observations"]) == 22 and observations["OracleVersions"] ==
                    {"python": "3.12.14", "pypdfium2": "5.13.0", "pdfium": "153.0.7999.0"},
                    "InputPreflight observation/oracle pins mismatch")
            oracle_raw = self.owned_input(path.parent / "inspect-identifiers.py").read_bytes()
            require(sha(oracle_raw) == observations["OracleScriptSHA256"].lower(), "Owned oracle source hash mismatch")
            oracle_name = f"{label}-InputPreflight-oracle.py"
            self.add_payload(oracle_name, oracle_raw)
            envelope_path = self.owned_input(observations["EnvelopeFixtureReceipt"])
            envelope_raw, envelope = load_json(envelope_path)
            require(sha(envelope_raw) == observations["EnvelopeFixtureReceiptSHA256"].lower() and
                    envelope["generator_sha256"] == sha((self.repo / "tools/test/generate_pdf_envelope_fixtures.py").read_bytes()),
                    "Envelope fixture receipt differs from current generator")
            require(envelope["ghostscript"]["exit_code"] == 0 and envelope["ghostscript"]["version"] == "10.08.0" and
                    envelope["ghostscript"]["child_gs_options_removed"] is True and len(envelope["fixtures"]) == 5,
                    "Native envelope fixture generation pin/result mismatch")
            for fixture in envelope["fixtures"]:
                require(sha(self.owned_input(envelope_path.parent / fixture["file"]).read_bytes()) == fixture["sha256"],
                        "Retained envelope fixture changed")
            envelope_name = f"{label}-InputPreflight-envelope-build.json"
            clean_envelope = json_bytes(self.sanitize_value(envelope))
            self.add_payload(envelope_name, clean_envelope)
            self.record_for(label, tier).update(oracle_file=oracle_name, oracle_sha256=sha(oracle_raw),
                                               envelope_build_file=envelope_name, envelope_build_raw_sha256=sha(envelope_raw),
                                               envelope_build_sha256=sha(clean_envelope))
        else:
            labels = [row["Label"] for row in observations["Observations"]]
            fixed = {"default-entry-directory-writable", "explicit-literal-writable-output",
                     "invalid-output-nonexistent", "invalid-output-file", "invalid-output-wildcard",
                     "invalid-output-nonfilesystem provider", "default-actual-directory-ACL-denial",
                     "explicit-writable-recovery-from-restricted-install",
                     "owned-probe-cleanup-before-later-dependency-failure", "directory-overlap-same directory",
                     "directory-overlap-case alias", "actual-junction-output leaf to source",
                     "actual-junction-source leaf to output", "actual-junction-output ancestor",
                     "actual-junction-source ancestor", "actual-concurrent-processes"}
            concurrent = {value for value in labels if value.startswith("simultaneous-shared-identity-")}
            optional_alias = {value for value in labels if value == "available-actual-8.3-directory-alias-refused"}
            require(len(labels) == len(set(labels)) and len(concurrent) == 2 and
                    set(labels) == fixed | concurrent | optional_alias,
                    "Destination observation labels differ from actual source contract (15 cases produce18/19 receipts)")
        clean = json_bytes(self.sanitize_value(observations))
        name = f"{label}-{tier}-observations.json"
        self.add_payload(name, clean)
        self.record_for(label, tier).update(observations_file=name, raw_observations_sha256=sha(raw),
                                           observations_sha256=sha(clean), observation_count=len(observations["Observations"]))
        return observations

    def analyzer(self, label, path, shell_version):
        raw, report = load_json(self.owned_input(path))
        require(report["Task"] == "T14" and report["Phase"] == "C1" and report["CommitUnderTest"] == self.commit and
                report["DirtyWorktree"] is False and report["ShellVersion"] == shell_version and
                report["AnalyzerVersion"] == "1.25.0", "Static analysis commit/runtime pin mismatch")
        require(all(item["Severity"] in (0, 1, 2) for item in report["Findings"]), "Unknown analyzer severity")
        for severity, count in ((2, "Errors"), (1, "Warnings"), (0, "Information")):
            require(report[count] == sum(item["Severity"] == severity for item in report["Findings"]),
                    "Static analysis raw findings/count mismatch")
        require(report["Errors"] == 0, "Static analysis errors remain")
        cleaned = json_bytes(self.sanitize_value(report))
        name = f"{label}-PSScriptAnalyzer-findings.json"
        self.add_payload(name, cleaned)
        self.records.append({"shell": label, "shell_version": shell_version,
                             "classification": "clean implementation static findings; retained warnings/information are not test passes",
                             "commit_under_test": self.commit, "dirty_worktree": False, "analyzer_version": "1.25.0",
                             "file": name, "raw_findings_report_sha256": sha(raw), "sanitized_findings_report_sha256": sha(cleaned),
                             "error_count": report["Errors"], "warning_count": report["Warnings"],
                             "information_count": report["Information"]})

    def historical(self):
        focused = {
            "T14-focused-ps51-9a767746909344699c905ce840b971bc": (120, 103),
            "T14-focused-ps51-a2e6039ee6624e5d8522490581df3943": (120, 103),
            "T14-focused-ps51-7f30fb0a9acf4836a693f24fd7d52573": (120, 120),
            "T14-focused-ps51-ac20eddc2c3b46988e625dea1f221572": (None, None),
            "T14-focused-ps51-af5e6ce18d5941d5aa2963f6db6884ff": (120, 0),
            "T14-focused-ps7-fce8de7e9f22428eb79cfebc02be6ee3": (120, 0),
            "T14-focused-ps51-d3ac4b951fd24c63a50cc9e3042d08b8": (122, 0),
            "T14-focused-ps7-98a041459984417485355a0e52068083": (122, 0),
        }
        for leaf, (total, failed) in focused.items():
            root = self.owned_input(self.work / leaf)
            run_raw, run = load_json(root / "run.json")
            require(run["task"] == "T14" and run["dirty_worktree"] is True and not run["timed_out"] and
                    re.fullmatch(r"[0-9a-f]{40}", run["commit_under_test"]), "Historical focused host context changed")
            name = "historical-" + leaf
            clean_run = json_bytes(self.sanitize_value(run))
            self.add_payload(name + "-run.json", clean_run)
            stdout = (root / "stdout.txt").read_bytes()
            stderr = (root / "stderr.txt").read_bytes()
            require(run["streams_sha256"] == {"stdout.txt": sha(stdout), "stderr.txt": sha(stderr)},
                    "Historical focused raw stream hashes changed")
            if failed == 0:
                raw, summary = load_json(root / "summary.json")
                require(summary["tier"] == "FocusedEmailMasterTool" and summary["passed"] == summary["total"] == total and
                        summary["dirty_worktree"] is True and summary["pester_version"] == "6.2.0" and
                        summary["shell_version"] == ("7.6.6" if run["shell"] == "ps7" else "5.1.26100.9444") and
                        all(summary[key] == 0 for key in LEGACY["BAD_COUNTS"]) and run["exit_code"] == 0,
                        "Historical focused successful report counts/pins mismatch")
                self.archive_historical(name, raw, summary, root / "results.xml", root / "stdout.txt", root / "stderr.txt",
                                        "historical dirty focused controlled unit/job pass; excluded from clean totals")
                record = self.records[-1]
            else:
                require(not (root / "summary.json").exists() and run["exit_code"] == 1,
                        "Historical failed host summary availability/exit changed")
                record = {"classification": "historical dirty focused test-host failure; excluded from clean totals",
                          "commit_under_test": run["commit_under_test"], "dirty_worktree": True,
                          "summary_available": False, "summary_note": "Host fault prevented summary creation; no replacement summary was invented."}
                for stream, raw in (("stdout", stdout), ("stderr", stderr)):
                    clean = self.sanitize_string(raw.decode("utf-8-sig")).encode("utf-8")
                    self.add_payload(name + "-" + stream + ".txt", clean)
                    record.update({stream + "_file": name + "-" + stream + ".txt", stream + "_raw_sha256": sha(raw),
                                   stream + "_sha256": sha(clean)})
                if total is None:
                    require(not (root / "results.xml").exists() and not stdout and b"Get-ExecutionPolicy" in stderr,
                            "Historical bootstrap-only failure unexpectedly has executed test evidence")
                    record.update(xml_available=False, executed_test_counts=None, phase="host bootstrap before Pester")
                else:
                    xml_raw = self.owned_input(root / "results.xml").read_bytes()
                    tree = LEGACY["ET"].fromstring(xml_raw)
                    require(int(tree.attrib["total"]) == total and int(tree.attrib["failures"]) == failed,
                            "Historical failed XML actual counts changed")
                    printed = re.findall(r"Tests Passed: ([0-9]+), Failed: ([0-9]+), Skipped: ([0-9]+), Inconclusive: ([0-9]+), NotRun: ([0-9]+)", stdout.decode("utf-8-sig"))
                    require(len(printed) == 1 and tuple(map(int, printed[0])) == (total - failed, failed, 0, 0, 0),
                            "Historical failed XML/stdout printed counts differ")
                    xml_clean = self.sanitized_xml(xml_raw, False, {"total": total, "failed": failed})
                    self.add_payload(name + "-results.xml", xml_clean)
                    record.update(xml_available=True, xml_file=name + "-results.xml", raw_xml_sha256=sha(xml_raw), xml_sha256=sha(xml_clean),
                                  counts_from_xml_and_stdout={"passed": total - failed, "failed": failed, "total": total},
                                  container_failure_printed=b"Container failed: 1" in stdout)
                self.records.append(record)
            record.update(command_receipt_file=name + "-run.json", command_receipt_raw_sha256=sha(run_raw),
                          command_receipt_sha256=sha(clean_run), command_arguments=self.sanitize_value(run["arguments"]),
                          child_environment_note="Earlier case-sensitive removal metadata described intent; diagnosis receipt clarifies actual uppercase-key inheritance. Corrected hosts remove only child PSMODULEPATH case insensitively.")
        diagnosis_raw, diagnosis = load_json(self.owned_input(self.work / "T14-focused-host-diagnosis.json"))
        require(diagnosis["task"] == "T14" and diagnosis["fabricated_summary_or_xml"] is False and
                len(diagnosis["historical_attempts"]) == 4, "Focused host-fault diagnosis receipt changed")
        for attempt in diagnosis["historical_attempts"]:
            root = self.owned_input(self.repo / attempt["root"])
            require(not attempt["summary_available"] and not (root / "summary.json").exists(), "Diagnosis invents a missing historical summary")
            for leaf, digest in attempt["raw_sha256"].items():
                require(Path(leaf).name == leaf and sha(self.owned_input(root / leaf).read_bytes()) == digest,
                        "Historical diagnosis/raw artifact binding changed")
        diagnosis_clean = json_bytes(self.sanitize_value(diagnosis))
        self.add_payload("historical-focused-host-diagnosis.json", diagnosis_clean)
        self.records.append({"classification": "Historical host-fault diagnosis; no executed tests or clean pass contribution",
                             "file": "historical-focused-host-diagnosis.json", "raw_sha256": sha(diagnosis_raw), "sha256": sha(diagnosis_clean)})
        paths = sorted(path for path in self.work.glob("T14-*.txt") if re.fullmatch(
            r"T14-(EmailOutcome|Staging|Destination|PdftkPaths|GhostscriptPaths)-(ps51|ps7)-[0-9a-f]{32}\.txt", path.name))
        require(len(paths) >= 9, "Required retained T14 exploratory native logs are missing")
        for log_path in paths:
            match = re.fullmatch(r"T14-(EmailOutcome|Staging|Destination|PdftkPaths|GhostscriptPaths)-(ps51|ps7)-[0-9a-f]{32}\.txt", log_path.name)
            require(match is not None, "Unexpected exploratory native log ownership name")
            tier, label = match.groups()
            raw_log = self.owned_input(log_path).read_bytes()
            receipts = re.findall(r"(?m)^Reports:\s*(.+?)\r?$", raw_log.decode("utf-8-sig"))
            require(len(receipts) == 1, "Exploratory native log requires its one retained exact Reports path")
            root = self.owned_input(receipts[0].strip())
            raw, summary = load_json(root / "summary.json")
            require(summary["tier"] == tier and summary["pester_version"] == "6.2.0" and
                    summary["total"] in ((9, 11) if tier == "EmailOutcome" else (COUNTS[tier],)) and
                    summary["passed"] + summary["failed"] + summary["skipped"] + summary["not_run"] == summary["total"] and
                    summary["shell_version"] == ("7.6.6" if label == "ps7" else "5.1.26100.9444"),
                    "Historical native smoke counts/runtime mismatch")
            name = "historical-standalone-" + log_path.stem
            self.archive_historical(name, raw, summary, root / "results.xml", log_path, None,
                                    "historical dirty/standalone native smoke or failed execution; actual counts retained and excluded from clean totals")
            obs_paths = re.findall(r"(?m)^(?:Native|Staging|Email|Destination) observations:\s*(.+?)\r?$", raw_log.decode("utf-8-sig"))
            require(len(obs_paths) == 1, "Exploratory native receipt missing/ambiguous")
            obs_raw, observations = load_json(self.owned_input(obs_paths[0].strip()))
            require(observations["CommitUnderTest"] == summary["commit_under_test"] and
                    observations["DirtyWorktree"] == summary["dirty_worktree"] and
                    observations["ShellVersion"] == summary["shell_version"], "Historical native observations context mismatch")
            clean = json_bytes(self.sanitize_value(observations))
            self.add_payload(name + "-observations.json", clean)
            self.records[-1].update(observations_file=name + "-observations.json", raw_observations_sha256=sha(obs_raw),
                                    observations_sha256=sha(clean), observation_count=len(observations["Observations"]))

    def archive_historical(self, name, raw_summary, summary, xml_path, stdout_path, stderr_path, classification):
        self.add_payload(name + "-summary.json", raw_summary)
        raw_xml = self.owned_input(xml_path).read_bytes()
        clean_xml = self.sanitized_xml(raw_xml, False, summary)
        self.add_payload(name + "-results.xml", clean_xml)
        record = {"classification": classification, "shell_version": summary["shell_version"],
                  "commit_under_test": summary["commit_under_test"], "dirty_worktree": summary["dirty_worktree"],
                  "summary_file": name + "-summary.json", "summary_sha256": sha(raw_summary),
                  "xml_file": name + "-results.xml", "raw_xml_sha256": sha(raw_xml), "xml_sha256": sha(clean_xml),
                  "counts": {key: summary[key] for key in LEGACY["COUNTS"]}}
        for stream, path in (("stdout", stdout_path), ("stderr", stderr_path)):
            if path is not None:
                raw = self.owned_input(path).read_bytes()
                clean = self.sanitize_string(raw.decode("utf-8-sig")).encode("utf-8")
                self.add_payload(name + "-" + stream + ".txt", clean)
                record.update({stream + "_file": name + "-" + stream + ".txt", stream + "_raw_sha256": sha(raw),
                               stream + "_sha256": sha(clean)})
        self.records.append(record)

    def verified_cache_receipt(self, path):
        raw, verification = load_json(self.owned_input(path))
        require(verification["task"] == "T14" and verification["acquisition_performed"] is False,
                "Cache receipt must identify T14 read-only approved-cache reuse")
        expected = {"Pester": "6.2.0", "PDFtk": "pdftk 2.02", "Ghostscript": "10.08.0",
                    "PowerShell7": "7.6.6", "PSScriptAnalyzer": "1.25.0"}
        require({row["dependency"]: row["version"] for row in verification["dependencies"]} == expected,
                "Approved cache version pins mismatch")
        require(sha((self.repo / verification["previous_full_audit"]).read_bytes()) == verification["previous_full_audit_sha256"],
                "Prior full approved-cache audit changed")
        for row in verification["dependencies"]:
            require(row["all_selected_file_hashes_match"] is True and row["selected_files_verified"] == len(row["selected_files"]),
                    "Cache verification selected-file count/result mismatch")
            require(sha((self.repo / row["source_receipt"]).read_bytes()) == row["source_receipt_sha256"],
                    "Approved acquisition receipt changed")
            root = Path(os.path.expandvars(row["cache_root"].replace("<USERPROFILE>", os.environ["USERPROFILE"]))).resolve()
            for item in row["selected_files"]:
                selected = (root / item["relative_path"]).resolve()
                require(selected.is_relative_to(root) and sha(selected.read_bytes()) == item["sha256"],
                        "Approved selected cache bytes changed")
        self.oracle_runtime = verification["development_oracle_runtime"]
        require({key: self.oracle_runtime[key] for key in ("python", "pypdfium2", "pdfium")} ==
                {"python": "3.12.14", "pypdfium2": "5.13.0", "pdfium": "153.0.7999.0"}, "Oracle version pins changed")
        for field in ("python", "pdfium_dll"):
            path = Path(os.path.expandvars(self.oracle_runtime[field + "_path"].replace("<USERPROFILE>", os.environ["USERPROFILE"]))).resolve()
            require(sha(path.read_bytes()) == self.oracle_runtime[field + "_sha256"], "Approved oracle runtime bytes changed")
        self.add_payload("approved-cache-verification.json", raw)
        self.standalone["T14-cache-verification.json"] = raw
        self.records.append({"classification": "Read-only approved external-cache verification; no acquisition or native pass claim",
                             "file": "approved-cache-verification.json", "sha256": sha(raw), "acquisition_performed": False,
                             "selected_files_verified": sum(row["selected_files_verified"] for row in verification["dependencies"])})

    def environment_and_pins(self, path):
        super().environment_and_pins(path)
        self.standalone["T14-environment.json"] = self.owned_input(path).read_bytes()

    def results_document(self, shells, destination):
        manifest_path = str((destination / "manifest.json").relative_to(self.repo)).replace("\\", "/")
        pins = self.acquisitions
        paths = {
            "ps51": "%SystemRoot%/System32/WindowsPowerShell/v1.0/powershell.exe",
            "ps7": pins["ps7"]["cache"]["directory_label"] + "/" + pins["ps7"]["cache"]["executable_relative_path"],
            "pester": pins["pester"]["cache"]["directory_label"].replace("<USERPROFILE>", "%USERPROFILE%") + "/" + pins["pester"]["cache"]["module_manifest_relative_path"],
            "pdftk": pins["pdftk"]["cache_root"] + "/" + pins["pdftk"]["extracted_files"][0]["relative_path"],
            "ghostscript": pins["gs"]["cache_root"] + "/" + pins["gs"]["version_probe"]["executable"],
            "analyzer": pins["analyzer"]["cache"]["directory_label"] + "/" + pins["analyzer"]["cache"]["manifest_relative_path"],
            "development_python": self.oracle_runtime["python_path"].replace("<USERPROFILE>", "%USERPROFILE%"),
        }
        return {"schema_version": 1, "task": "T14", "checkpoint": "C1", "commit_under_test": self.commit,
                "dirty_worktree": False, "implementation_acceptance": "pass", "ac032": "pass", "ac033": "pass", "ac034": "pass",
                "checkpoint_note": "This certifies clean implementation tests only. Root records review, task closure and final push/live equality separately.",
                "scope": "Explicit optional-email and master result states with end-to-end 0/1/2 outcomes, complete selected PDFtk inspection and exact page agreement before strict-smaller email publication; actual Windows entry/BAT/GS/PDFtk and independent synthetic visible-page checks, preserved masters/sources/foreign hashes, plus affected regressions.",
                "pester_version": "6.2.0", "analyzer_version": "1.25.0", "pdftk_version": "2.02", "ghostscript_version": "10.08.0",
                "cases_per_shell": COUNTS, "passed_per_shell": {row["shell"]: row["passed"] for row in shells},
                "total_passed": 952, "all_failures_skips_not_run": 0, "reports_manifest": manifest_path,
                "selected_cache_verification": "docs/codex/evidence/T14-cache-verification.json",
                "ordinary_environment_receipt": "docs/codex/evidence/T14-environment.json",
                "command_paths_from_acquisition_receipts": paths,
                "selected_tiers_in_execution_order": list(TIERS),
                "command_template": "<explicit shell> -NoProfile -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -PesterModulePath <manifest> -Tier <tier, including Launcher> [-PdftkPath <real exe> for EmailOutcome/MasterValidation/Staging/InputPreflight/Destination/PdftkPaths/GhostscriptPaths/SourceDiscovery/DependencyEntry/LauncherNative] [-GhostscriptPath <real exe> for EmailOutcome/Staging/InputPreflight/Destination/GhostscriptPaths] [-PythonPath <approved Python> for EmailOutcome/MasterValidation/InputPreflight]",
                "collector_command_template": "<explicit shell> -NoProfile -ExecutionPolicy RemoteSigned -File tests/.work/Run-T14Checkpoint.ps1 -ShellLabel ps51|ps7 -ExpectedCommit " + self.commit + " -ExpectedCountsPath tests/.work/T14-expected-counts.json -Checkpoint C1",
                "evidence_collector_command": "<approved Python> -B tests/.work/Collect-T14Evidence.py --repo . --commit " + self.commit + " [--write]",
                "command_path_expansion": "Expand %EnvironmentVariable% labels and resolve literally. Per-tier exact sanitized argument vectors and raw log/metadata hashes appear in the manifest.",
                "environment": {"os": "Windows 11 x64 build26300", "ordinary_ps51_inventory": self.inventory,
                                "standard_user_non_elevated": True, "ps51": shells[0]["shell_version"], "ps7": shells[1]["shell_version"],
                                "test_process_policy": "RemoteSigned for authorized child test hosts only", "acquisition_performed": False,
                                "persistent_policy_path_security_changes": False, "os_support_channel_established": False},
                "static_analysis": [row for row in self.records if "analyzer_version" in row],
                "executions": [{"shell": label, "tier": summary["tier"], "exit_code": 0, "summary": summary,
                                "reports_manifest": manifest_path} for label, summary in self.summaries],
                "historical_rule": "Dirty focused 120/120 and 122/122 reports remain separate from clean totals. Three failed host attempts retain actual XML/stdout counts (17/103 twice, 0/120 plus one printed container failure); the bootstrap-only attempt has no XML. These four attempts have no fabricated summaries. Standalone native smoke/failed reports retain actual counts, including historical EmailOutcome PS5.1 1 passed/10 failed/11 total before locale-neutral dimension assertions.",
                "limitations": ["T15 interruption/descendants, T16 further options and T17 fidelity/features remain separate tasks.",
                                "Copied helper seams control discovery/log exceptions, separate corrupt GS input and injected expected counts after actual master publication. Existing controlled staged substitutions, foreign finals and concurrency barriers remain disclosed separately from actual engine and publication behavior.",
                                "Staging publication-race masters receive4096 bytes of owned synthetic whitespace before preservation snapshots so genuine GS reaches a strict-smaller move; this preparation never alters source fixtures or another run's outputs.",
                                "Actual GS accepts the retained CJK output path but selected PDFtk2.02 inspection fails there; the result fails closed without publishing an unvalidated email or renaming sources.",
                                "Cleanup remains best effort. Marker retention is attempted only for the original identifiable stage after late directory-removal failure; identity loss is disclosed and fails closed.",
                                "Focused local synthetic Windows evidence does not certify full PDF fidelity, signatures, Explorer drag/drop, live UNC, OS support channels, CI, package or release gates.",
                                "Static warnings/information are retained at actual counts; zero errors does not mean lint-clean.",
                                "Approved caches were reused; no acquisition, installer, elevation, security changes, vendor redistribution or private PDF upload."]}

    def finish(self, shells, destination, write):
        destination = destination.resolve()
        evidence = (self.repo / "docs/codex/evidence").resolve()
        require(destination == evidence / "T14-C1-reports", "T14 C1 reports must use their exact intended evidence directory")
        require(sum(row["passed"] for row in shells) == 952, "Expected exact clean total952")
        results_path = evidence / "T14-C1-results.json"
        results_payload = json_bytes(self.results_document(shells, destination))
        manifest = {"schema_version": 1, "task": "T14", "checkpoint": "C1", "commit_under_test": self.commit,
                    "dirty_worktree": False, "clean_reports": 26, "total_clean_passed": 952, "shells": shells,
                    "xml_redactions": ["environment." + key for key in LEGACY["XML_IDENTITY"]],
                    "native_json_redactions": "Recursive repository/user/cache/temp string-prefix redactions; XML machine/user/domain/cwd identity fields removed.",
                    "exact_byte_copies": "Original summary/build/environment/cache bytes; raw and sanitized XML/observations/log/analyzer hashes recorded separately.",
                    "evidence_collector_sha256": sha(Path(__file__).read_bytes()), "base_primitives_sha256": sha(LEGACY_PATH.read_bytes()),
                    "results_file": str(results_path.relative_to(self.repo)).replace("\\", "/"), "results_sha256": sha(results_payload),
                    "standalone_exact_receipts": [{"file": "docs/codex/evidence/" + name, "sha256": sha(payload)}
                                                  for name, payload in self.standalone.items()],
                    "historical_rule": "All dirty focused passes, failed host XML/bootstrap attempts, and standalone native smoke/failed records are excluded from clean acceptance totals; unavailable summaries/XML are disclosed rather than invented.", "records": self.records}
        self.add_payload("manifest.json", json_bytes(manifest))
        for name, payload in self.payloads.items():
            self.privacy_gate(payload, name)
        self.privacy_gate(results_payload, results_path.name)
        for name, payload in self.standalone.items():
            require(Path(name).name == name, "Standalone path must remain confined")
            self.privacy_gate(payload, name)
        outcome = {"task": "T14", "clean_commit": self.commit, "check_only": not write, "clean_reports": 26,
                   "total_passed": 952, "per_shell": shells, "historical_records": sum(row["classification"].startswith("historical") for row in self.records),
                   "public_files": len(self.payloads) + len(self.standalone) + 1,
                   "manifest_sha256": sha(self.payloads["manifest.json"]), "results_sha256": sha(results_payload)}
        if write:
            assert_clean(self.repo, self.commit)
            require(not destination.exists() and not results_path.exists() and
                    all(not (evidence / name).exists() for name in self.standalone), "Never overwrite existing public evidence")
            destination.mkdir()
            outputs = [(destination / name, payload) for name, payload in self.payloads.items()]
            outputs += [(results_path, results_payload)] + [(evidence / name, payload) for name, payload in self.standalone.items()]
            for path, payload in outputs:
                with path.open("xb") as handle:
                    handle.write(payload)
                require(sha(path.read_bytes()) == sha(payload), "Written evidence byte hash mismatch")
            outcome["destination"] = str(destination.relative_to(self.repo)).replace("\\", "/")
        print(json.dumps(outcome, indent=2))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=Path.cwd())
    parser.add_argument("--commit", default=COMMIT)
    parser.add_argument("--counts", type=Path)
    parser.add_argument("--ps51-root", type=Path)
    parser.add_argument("--ps7-root", type=Path)
    parser.add_argument("--analyzer-ps51", type=Path)
    parser.add_argument("--analyzer-ps7", type=Path)
    parser.add_argument("--environment", type=Path)
    parser.add_argument("--cache-verification", type=Path)
    choice = parser.add_mutually_exclusive_group()
    choice.add_argument("--check-only", action="store_true")
    choice.add_argument("--write", action="store_true")
    args = parser.parse_args()
    require(args.commit == COMMIT, "T14 C1 collector accepts only the explicit reviewed implementation SHA")
    repo = args.repo.resolve()
    assert_clean(repo, args.commit)
    collector = T14Collector(repo, args.commit)
    count_path = collector.owned_input(args.counts or collector.work / "T14-expected-counts.json")
    count_raw, counts = load_json(count_path)
    require(counts == COUNTS and tuple(counts) == TIERS, "Exact frozen T14 expected counts required")
    collector.counts_digest = sha(count_raw)
    collector.add_payload("frozen-expected-counts.json", count_raw)
    collector.records.append({"classification": "Frozen focused-tier expected counts; not executed tests",
                              "file": "frozen-expected-counts.json", "sha256": sha(count_raw)})
    shells = [collector.clean_shell("ps51", args.ps51_root or collector.work / "T14-C1-ps51"),
              collector.clean_shell("ps7", args.ps7_root or collector.work / "T14-C1-ps7")]
    collector.historical()
    collector.analyzer("ps51", args.analyzer_ps51 or collector.work / "T14-C1-analyzer-ps51.json", shells[0]["shell_version"])
    collector.analyzer("ps7", args.analyzer_ps7 or collector.work / "T14-C1-analyzer-ps7.json", shells[1]["shell_version"])
    collector.environment_and_pins(args.environment or collector.work / "T14-environment.json")
    collector.verified_cache_receipt(args.cache_verification or collector.work / "T14-cache-verification.json")
    collector.finish(shells, repo / "docs/codex/evidence/T14-C1-reports", args.write)


if __name__ == "__main__":
    main()
