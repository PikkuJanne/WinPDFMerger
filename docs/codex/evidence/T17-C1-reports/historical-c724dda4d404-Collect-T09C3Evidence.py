"""Ignored T09 C3 evidence collector. Invoke only on root-authorized clean SHA/roots.

Validates every input before creating any public output; never changes handoff status.
Exact summaries/build receipts retain their original bytes. XML environment identity
fields and recursively parsed native JSON strings receive explicit privacy redactions.
Historical failed/dirty records remain separate and never contribute to pass counts.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import xml.etree.ElementTree as ET


TIERS = (
    "Unit", "ToolInvocation", "PdftkPaths", "GhostscriptPaths", "NativeRunner",
    "DependencyEntry", "SourceDiscovery", "Launcher", "LauncherNative",
)
BAD_COUNTS = ("failed", "failed_blocks", "failed_containers", "skipped", "not_run")
COUNTS = ("passed",) + BAD_COUNTS + ("total",)
XML_IDENTITY = ("user", "user-domain", "machine-name", "cwd")
HISTORICAL_REPORT = "pester/f0ee29c67029494584d6293ef003dae6"
HISTORICAL_OBSERVATIONS = "tool-paths/994a375210e146bb9f264c6fe13ec6ca"


def require(condition, message):
    if not condition:
        raise ValueError(message)


def sha(data):
    return hashlib.sha256(data).hexdigest()


def load_json(path):
    raw = path.read_bytes()
    return raw, json.loads(raw.decode("utf-8-sig"))


def json_bytes(value):
    return (json.dumps(value, indent=2, ensure_ascii=False) + "\n").encode("utf-8")


class Collector:
    def __init__(self, repo, commit):
        self.repo = repo.resolve()
        self.work = (self.repo / "tests/.work").resolve()
        self.commit = commit
        self.payloads = {}
        self.records = []
        self.summaries = []
        self.inventory = None
        self.acquisitions = {}
        self.helper_results = None
        self.identities = set()
        self.prefixes = []
        # More specific prefixes are replaced first; slash-mixed and JSON-escaped
        # paths inside a native console/log string are covered by the same regex.
        for name in ("TEMP", "TMP", "LOCALAPPDATA", "APPDATA", "USERPROFILE", "HOME"):
            value = os.environ.get(name)
            if value:
                self.prefixes.append((value.rstrip("\\/"), f"<{name}>"))
        home = str(Path.home())
        if home:
            self.prefixes.append((home.rstrip("\\/"), "<USERPROFILE>"))
        self.prefixes.append((str(self.repo), "<repo>"))
        self.prefixes.sort(key=lambda item: len(item[0]), reverse=True)

    def owned_input(self, path):
        path = Path(path).resolve()
        require(path.is_relative_to(self.work), f"Input escapes owned test work: {path.name}")
        require(path.exists(), f"Missing owned input: {path.name}")
        return path

    def sanitize_string(self, value, xml=False):
        for prefix, replacement in self.prefixes:
            # Bracket markers are valid in XML attributes/text and CDATA alike.
            if xml:
                replacement = replacement.replace("<", "[").replace(">", "]")
            components = re.split(r"[\\/]+", prefix)
            pattern = r"[\\/]+".join(re.escape(part) for part in components)
            value = re.sub(pattern, lambda _: replacement, value, flags=re.IGNORECASE)
        return value

    def sanitize_value(self, value):
        if isinstance(value, str):
            return self.sanitize_string(value)
        if isinstance(value, list):
            return [self.sanitize_value(item) for item in value]
        if isinstance(value, dict):
            return {key: self.sanitize_value(item) for key, item in value.items()}
        return value

    def privacy_gate(self, payload, name):
        text = payload.decode("utf-8-sig")
        require(not re.search(r"(?i)[A-Z]:[\\/]+Users[\\/]+[^<>\\/\s]+", text),
                f"Unredacted user-profile path in {name}")
        require(not re.search(r"(?i)ghp_[A-Za-z0-9]+|github_pat_[A-Za-z0-9_]+|https://[^/\s]+@", text),
                f"Credential marker in {name}")
        for identity in self.identities:
            if len(identity) >= 4 and identity != "REDACTED":
                require(not re.search(r"(?<![\w])" + re.escape(identity) + r"(?![\w])", text, re.IGNORECASE),
                        f"Unredacted XML identity in {name}")

    def add_payload(self, name, payload):
        require(name not in self.payloads, f"Duplicate public filename: {name}")
        require(Path(name).name == name, "Public payload name must be a simple filename")
        self.payloads[name] = payload

    def sanitized_xml(self, raw, clean, summary):
        content = raw.decode("utf-8-sig")
        tree = ET.fromstring(content)
        require(tree.tag == "test-results", "Expected exact Pester NUnit2 report")
        require(int(tree.attrib["total"]) == summary["total"], "XML/summary total mismatch")
        require(int(tree.attrib["failures"]) == summary["failed"], "XML/summary failures mismatch")
        if clean:
            for key in ("errors", "failures", "not-run", "inconclusive", "ignored", "skipped", "invalid"):
                require(int(tree.attrib.get(key, "0")) == 0, f"Clean XML contains {key}")
        environments = re.findall(r"<environment\b[^>]*>", content)
        require(len(environments) == 1, "Missing or ambiguous NUnit environment")
        original = environments[0]
        redacted = original
        for key in XML_IDENTITY:
            match = re.search(r"\b" + key + r'="([^"]*)"', original)
            if match:
                if key != "cwd":
                    self.identities.add(match.group(1))
                redacted = re.sub(r"\b" + key + r'="[^"]*"', key + '="REDACTED"', redacted)
        content = content.replace(original, redacted, 1)
        output = self.sanitize_string(content, xml=True).encode("utf-8")
        ET.fromstring(output)
        return output

    def build_receipt(self, receipt_path, expected_digest, shell, tier, executable_name, source):
        path = self.owned_input(receipt_path)
        raw, receipt = load_json(path)
        require(sha(raw) == expected_digest.lower(), f"{tier} build receipt digest mismatch")
        require(sha((self.repo / source).read_bytes()) == receipt["source_sha256"].lower(),
                f"{tier} receipt differs from clean source")
        executable = self.owned_input(path.parent / executable_name)
        require(sha(executable.read_bytes()) == receipt["executable_sha256"].lower(),
                f"{tier} owned executable differs from build receipt")
        name = f"{shell}-{tier}-build.json"
        self.add_payload(name, raw)
        return {"build_receipt": name, "build_receipt_sha256": sha(raw),
                "source_sha256": receipt["source_sha256"],
                "controlled_executable_sha256": receipt["executable_sha256"],
                "scope": "Controlled process fixture only; no native PDF-engine support claim"}

    def clean_shell(self, label, root):
        root = self.owned_input(root)
        raw_runs, jobs = load_json(root / "runs.json")
        require(isinstance(jobs, list) and len(jobs) == len(TIERS), f"{label}: expected nine executed tiers")
        require(tuple(job["tier"] for job in jobs) == TIERS, f"{label}: unexpected tier selection/order")
        raw_aggregate, aggregate = load_json(root / "aggregate.json")
        require(aggregate["commit_under_test"] == self.commit and aggregate["dirty_worktree"] is False,
                f"{label}: aggregate not at requested clean SHA")
        require(aggregate["shell"] == label and aggregate["tiers"] == len(TIERS), "Aggregate shell/tier mismatch")
        total = 0
        versions = set()
        for job in jobs:
            tier = job["tier"]
            require(job["shell"] == label and job["exit_code"] == 0, f"{label}/{tier}: failed collector exit")
            report = self.owned_input(job["report"])
            raw, summary = load_json(report / "summary.json")
            require(summary["commit_under_test"] == self.commit and summary["dirty_worktree"] is False,
                    f"{label}/{tier}: requested clean implementation SHA not observed")
            require(summary["tier"] == tier and summary["pester_version"] == "6.2.0", "Tier/Pester pin mismatch")
            require(summary["process_64_bit"] is True, "Required test process is not x64")
            require(summary["execution_policy"] == "RemoteSigned", "Unexpected test-process execution policy")
            require(summary["total"] == summary["passed"] == job["expected_count"] > 0,
                    f"{label}/{tier}: unexpected test counts")
            require(all(summary[key] == 0 for key in BAD_COUNTS), f"{label}/{tier}: failure/skip/not-run count")
            # ConvertFrom/To-Json can remove insignificant trailing timestamp
            # zeros. Retain raw summary bytes and compare that duplicated value
            # at its original precision after removing only trailing zeros.
            job_summary = dict(job["summary"])
            raw_summary = dict(summary)
            def normalized_stamp(value):
                return re.sub(r"\.(\d+)(Z)$", lambda match:
                              ("." + match.group(1).rstrip("0") if match.group(1).rstrip("0") else "") + match.group(2), value)
            job_summary["observed_at_utc"] = normalized_stamp(job_summary["observed_at_utc"])
            raw_summary["observed_at_utc"] = normalized_stamp(raw_summary["observed_at_utc"])
            require(job_summary == raw_summary, f"{label}/{tier}: collector/raw summary mismatch")
            if label == "ps7":
                require(summary["shell_edition"] == "Core" and summary["shell_version"] == "7.6.6", "PS7 pin mismatch")
            else:
                require(summary["shell_edition"] == "Desktop" and summary["shell_version"].startswith("5.1."), "PS5.1 mismatch")
            versions.add(summary["shell_version"])
            total += summary["passed"]
            name = f"{label}-{tier}"
            summary_name = name + "-summary.json"
            self.add_payload(summary_name, raw)
            xml_raw = (report / "results.xml").read_bytes()
            xml_clean = self.sanitized_xml(xml_raw, True, summary)
            xml_name = name + "-results.xml"
            self.add_payload(xml_name, xml_clean)
            log_path = self.owned_input(job["log"])
            log_raw = log_path.read_bytes()
            log = log_raw.decode("utf-8-sig")
            entry = {"shell": label, "shell_version": summary["shell_version"], "tier": tier,
                     "classification": "clean implementation acceptance/regression execution",
                     "commit_under_test": self.commit, "dirty_worktree": False,
                     "summary_file": summary_name, "summary_sha256": sha(raw),
                     "raw_xml_sha256": sha(xml_raw), "xml_file": xml_name, "xml_sha256": sha(xml_clean),
                     "collector_log_raw_sha256": sha(log_raw),
                     "counts": {key: summary[key] for key in COUNTS}}
            if tier == "NativeRunner":
                entry.update(self.build_receipt(summary["native_fixture_build_receipt"],
                             summary["native_fixture_build_receipt_sha256"], label, tier,
                             "FakeNative.exe", "tests/native/FakeNative.cs"))
            elif tier == "DependencyEntry":
                require("dependency_fixture_build_receipt" in job and "dependency_fixture_build_receipt_sha256" in job,
                        "DependencyEntry collector needs its exact newly owned build receipt path/digest")
                entry.update(self.build_receipt(job["dependency_fixture_build_receipt"],
                             job["dependency_fixture_build_receipt_sha256"], label, tier,
                             "VersionProbeFixture.exe", "tests/dependencies/VersionProbeFixture.cs"))
            elif tier in ("PdftkPaths", "GhostscriptPaths"):
                matches = re.findall(r"(?m)^Native observations:\s*(.+?)\r?$", log)
                require(len(matches) == 1, "Missing/ambiguous native observations receipt")
                path = self.owned_input(matches[0].strip())
                obs_raw, observations = load_json(path)
                backend = "Pdftk" if tier == "PdftkPaths" else "Ghostscript"
                require(observations["CommitUnderTest"] == self.commit and observations["DirtyWorktree"] is False,
                        f"{backend} observations are not from requested clean SHA")
                require(observations["ToolBackend"] == backend and observations["ShellVersion"] == summary["shell_version"],
                        "Native observations backend/shell mismatch")
                require(observations["Observations"], "Native observations are empty")
                obs_clean = json_bytes(self.sanitize_value(observations))
                obs_name = name + "-observations.json"
                self.add_payload(obs_name, obs_clean)
                entry.update(observations_file=obs_name, raw_observations_sha256=sha(obs_raw),
                             observations_sha256=sha(obs_clean), native_backend=backend,
                             native_version=observations["ToolVersion"], observation_count=len(observations["Observations"]))
            self.records.append(entry)
            self.summaries.append((label, summary))
        require(len(versions) == 1, "A shell collector mixed runtime versions")
        require(total == aggregate["total_passed"] == aggregate["expected_total"], "Aggregate pass total mismatch")
        require(aggregate["all_failures_skips_not_run"] == 0, "Aggregate bad-count marker")
        return {"shell": label, "shell_version": next(iter(versions)), "tiers": len(jobs), "passed": total,
                "all_failure_skip_not_run_counts": 0, "runs_raw_sha256": sha(raw_runs),
                "aggregate_raw_sha256": sha(raw_aggregate), "counts_per_tier": {job["tier"]: job["expected_count"] for job in jobs}}

    def historical(self):
        report = self.owned_input(self.work / HISTORICAL_REPORT)
        raw, summary = load_json(report / "summary.json")
        require(summary["dirty_worktree"] is True and summary["failed"] > 0,
                "Historical failed GS record must remain explicitly dirty/failed")
        require(summary["tier"] == "GhostscriptPaths", "Unexpected historical GS tier")
        name = "historical-dirty-GhostscriptPaths-summary.json"
        self.add_payload(name, raw)
        xml_raw = (report / "results.xml").read_bytes()
        xml_clean = self.sanitized_xml(xml_raw, False, summary)
        xml_name = "historical-dirty-GhostscriptPaths-results.xml"
        self.add_payload(xml_name, xml_clean)
        self.records.append({"classification": "historical dirty failed GS suite; never counted as acceptance pass",
                             "commit_under_test": summary["commit_under_test"], "dirty_worktree": True,
                             "summary_file": name, "summary_sha256": sha(raw),
                             "raw_xml_sha256": sha(xml_raw), "xml_file": xml_name, "xml_sha256": sha(xml_clean),
                             "counts": {key: summary[key] for key in COUNTS}})
        obs_root = self.owned_input(self.work / HISTORICAL_OBSERVATIONS)
        obs_raw, observations = load_json(obs_root / "native-observations.json")
        require(observations["DirtyWorktree"] is True and observations["CommitUnderTest"] == summary["commit_under_test"],
                "Historical GS observation/suite context mismatch")
        obs_clean = json_bytes(self.sanitize_value(observations))
        obs_name = "historical-dirty-GhostscriptPaths-observations.json"
        self.add_payload(obs_name, obs_clean)
        self.records.append({"classification": "historical dirty GS observations; never counted as acceptance pass",
                             "commit_under_test": summary["commit_under_test"], "dirty_worktree": True,
                             "file": obs_name, "raw_sha256": sha(obs_raw), "sha256": sha(obs_clean)})
        direct_paths = sorted(obs_root.rglob("*PDFSTOPONERROR.json"))
        require(direct_paths, "Historical direct PDFSTOPONERROR probe not found beneath observation root")
        actual_count = 0
        for index, path in enumerate(direct_paths, 1):
            path = self.owned_input(path)
            raw_direct, direct = load_json(path)
            require(direct["Option"] == "-dPDFSTOPONERROR", "Unexpected direct historical probe flag")
            started = direct["Result"]["Started"]
            if started:
                actual_count += 1
                require(direct["Result"]["ExitCode"] == 1 and direct["Result"]["Succeeded"] is False,
                        "Historical actual GS stop-on-error probe expected exit1/failure")
            clean_direct = json_bytes(self.sanitize_value(direct))
            direct_name = f"historical-dirty-direct-PDFSTOPONERROR-{index}.json"
            self.add_payload(direct_name, clean_direct)
            self.records.append({"classification": ("historical direct GS stop-on-error characterization" if started else
                                                      "historical direct probe prelaunch harness fault") + "; never counted as acceptance pass",
                                 "context_commit_under_test": summary["commit_under_test"], "context_dirty_worktree": True,
                                 "commit_metadata_in_raw_receipt": False,
                                 "context_binding": "Owned path beneath the retained dirty historical GS observation root",
                                 "source_relative_path": str(path.relative_to(self.work)).replace("\\", "/"),
                                 "file": direct_name, "raw_sha256": sha(raw_direct), "sha256": sha(clean_direct),
                                 "native_started": started, "native_exit_code": direct["Result"]["ExitCode"],
                                 "output_exists": direct["OutputExists"]})
        require(actual_count == 1, "Historical actual direct GS characterization missing/ambiguous")

    def analyzer(self, label, path, shell_version):
        raw, report = load_json(self.owned_input(path))
        require(report["commit_under_test"] == self.commit and report["dirty_worktree"] is False,
                f"{label}: analyzer findings are not from requested clean implementation SHA")
        require(report["shell_version"] == shell_version and report["analyzer_version"] == "1.25.0",
                "Analyzer runtime/version pin mismatch")
        for severity, key in (("Error", "error_count"), ("Warning", "warning_count"), ("Information", "information_count")):
            require(report[key] == len([finding for finding in report["findings"] if finding["severity"] == severity]),
                    "Analyzer findings/count mismatch")
        require(report["error_count"] == 0, "Analyzer errors remain unresolved")
        clean = json_bytes(self.sanitize_value(report))
        name = f"{label}-PSScriptAnalyzer-findings.json"
        self.add_payload(name, clean)
        self.records.append({"shell": label, "shell_version": shell_version,
                             "classification": "clean implementation static findings; warnings/information retained; no native PDF acceptance claim",
                             "commit_under_test": self.commit, "dirty_worktree": False,
                             "analyzer_version": report["analyzer_version"], "file": name,
                             "raw_findings_report_sha256": sha(raw), "sanitized_findings_report_sha256": sha(clean),
                             "error_count": report["error_count"], "warning_count": report["warning_count"],
                             "information_count": report["information_count"]})

    def environment_and_pins(self, inventory_path):
        raw, inventory = load_json(self.owned_input(inventory_path))
        require(inventory["process_64_bit"] is True and inventory["administrator"] is False,
                "Ordinary PS5.1 inventory must be x64/non-elevated")
        require(inventory["policy"] == "Restricted" and inventory["shell_version"].startswith("5.1."),
                "Ordinary PS5.1 inventory changed")
        require(inventory["scopes"] == [name + "=Undefined" for name in
                ("MachinePolicy", "UserPolicy", "Process", "CurrentUser", "LocalMachine")], "Unexpected ordinary policy scopes")
        require(inventory["os_version"] == "10.0.26300.0", "Unexpected recorded reference OS build")
        self.inventory = inventory
        self.add_payload("ordinary-ps51-environment.json", raw)
        self.records.append({"classification": "Separate ordinary Windows PS5.1 read-only host inventory; no native test/pass claim",
                             "file": "ordinary-ps51-environment.json", "raw_sha256": sha(raw), "sha256": sha(raw),
                             "observed_at_utc": inventory["observed_at_utc"], "commit_metadata_in_raw_receipt": False})
        for label, relative in (
            ("pester", "T03-pester-acquisition.json"), ("pdftk", "T03-pdftk-acquisition.json"),
            ("gs", "T09-gs-acquisition.json"), ("ps7", "T09-ps7-acquisition.json"),
            ("analyzer", "T09-analyzer-acquisition.json"),
        ):
            path = self.repo / "docs/codex/evidence" / relative
            receipt_raw, receipt = load_json(path)
            self.acquisitions[label] = receipt
            self.records.append({"classification": "Existing authorized external-cache dependency acquisition; separate from application acceptance",
                                 "evidence_file": "docs/codex/evidence/" + relative, "sha256": sha(receipt_raw)})

    def python_helpers(self, root):
        root = self.owned_input(root)
        environment_raw, environment = load_json(root / "environment.json")
        self.add_payload("python-helper-environment.json", environment_raw)
        self.records.append({"classification": "Development-only Python fixture/oracle environment; excluded from application tier totals",
                             "file": "python-helper-environment.json", "sha256": sha(environment_raw)})
        helpers = []
        for label in ("reproduce", "oracle", "oracle-tests"):
            command_raw, command = load_json(root / (label + ".json"))
            require(command["CommitUnderTest"] == self.commit and command["DirtyWorktree"] is False and command["ExitCode"] == 0,
                    "Python helper is not a successful clean implementation execution")
            command_name = "python-helper-" + label + "-command.json"
            self.add_payload(command_name, command_raw)
            output_raw = (root / (label + ".txt")).read_bytes()
            output_text = output_raw.decode("utf-8-sig")
            if label == "oracle":
                oracle = json.loads(output_text)
                require(oracle["total_pages"] == 4 and len(oracle["fixtures"]) == 3,
                        "Fixture oracle expected three fixtures/four pages")
                require(all(oracle[key] is False for key in ("product_merge_executed", "pdftk_executed", "ghostscript_executed")),
                        "Development helper was incorrectly presented as application engine evidence")
                output_name = "python-helper-oracle-output.json"
                output_clean = json_bytes(self.sanitize_value(oracle))
            else:
                output_name = "python-helper-" + label + "-output.txt"
                output_clean = self.sanitize_string(output_text).encode("utf-8")
                if label == "reproduce":
                    require("PASS: 3 PDFs and manifest reproduce byte-for-byte." in output_text, "Fixture reproduction result missing")
                else:
                    require(re.search(r"Ran 10 tests in [^\r\n]+", output_text) and output_text.strip().endswith("OK"),
                            "Expected actual ten passing helper unittest cases")
            self.add_payload(output_name, output_clean)
            record = {"classification": "Clean development-only Python fixture/oracle helper; excluded from application tier totals",
                      "commit_under_test": self.commit, "dirty_worktree": False, "command": command["Command"],
                      "exit_code": 0, "command_file": command_name, "command_raw_sha256": sha(command_raw),
                      "output_file": output_name, "output_raw_sha256": sha(output_raw), "output_sha256": sha(output_clean)}
            self.records.append(record)
            helpers.append(record)
        self.helper_results = {"scope": "Development fixture/oracle helper checks, separate from434application/control/native tier passes",
                               "environment": environment, "commands": helpers,
                               "unit_tests_passed": 10, "fixtures_reproduced": 3, "fixture_oracle_pages": 4,
                               "application_merge_executed": False, "pdftk_executed": False, "ghostscript_executed": False}

    def results_document(self, shells, destination):
        pins = self.acquisitions
        paths = {
            "ps51": "%SystemRoot%/System32/WindowsPowerShell/v1.0/powershell.exe",
            "ps7": pins["ps7"]["cache"]["directory_label"] + "/" + pins["ps7"]["cache"]["executable_relative_path"],
            "pester": pins["pester"]["cache"]["directory_label"].replace("<USERPROFILE>", "%USERPROFILE%") + "/" + pins["pester"]["cache"]["module_manifest_relative_path"],
            "pdftk": pins["pdftk"]["cache_root"] + "/" + pins["pdftk"]["extracted_files"][0]["relative_path"],
            "ghostscript": pins["gs"]["cache_root"] + "/" + pins["gs"]["version_probe"]["executable"],
            "analyzer": pins["analyzer"]["cache"]["directory_label"] + "/" + pins["analyzer"]["cache"]["manifest_relative_path"],
        }
        require(pins["ps7"]["dependency"]["version"] == "7.6.6" and
                pins["gs"]["ghostscript_installer"]["version"] == "10.08.0" and
                pins["analyzer"]["dependency"]["version"] == "1.25.0", "Acquisition pin mismatch")
        require(all(shell["passed"] == 217 for shell in shells), "Expected complete focused tier counts")
        manifest_relative = str((destination / "manifest.json").relative_to(self.repo)).replace("\\", "/")
        static = [record for record in self.records if "analyzer_version" in record]
        return {"schema_version": 1, "task": "T09", "checkpoint": "C3", "commit_under_test": self.commit,
                "dirty_worktree": False, "implementation_acceptance": "pass", "ac019": "pass", "ac020": "pass", "ac021": "pass",
                "checkpoint_note": "Task/status and final records commit/push synchronization are maintained by root separately; this report certifies the executed clean implementation gates only.",
                "scope": "Focused M1 cross-cutting regressions and actual Windows PDFtk2.02/Ghostscript10.08.0 path/noninteractive integration; controlled fixtures are separate from real engines; no final fidelity/Explorer/package/release pass claim.",
                "pester_version": "6.2.0", "analyzer_version": "1.25.0", "pdftk_version": "2.02", "ghostscript_version": "10.08.0",
                "cases_per_shell": shells[0]["counts_per_tier"], "passed_per_shell": {s["shell"]: s["passed"] for s in shells},
                "total_passed": sum(s["passed"] for s in shells), "all_failures_skips_not_run": 0,
                "reports_manifest": manifest_relative,
                "command_template": "<explicit shell> -NoProfile -ExecutionPolicy RemoteSigned -File tools/test/Invoke-Tests.ps1 -PesterModulePath <pester manifest> -Tier <tier> [-PdftkPath <exe> for PdftkPaths/GhostscriptPaths/DependencyEntry/SourceDiscovery/LauncherNative] [-GhostscriptPath <exe> for GhostscriptPaths]",
                "collector_command_template": "<existing shell> -NoProfile -ExecutionPolicy RemoteSigned -File tests/.work/Run-T09Checkpoint-C3.ps1 -ShellLabel ps51|ps7 -ExpectedCommit " + self.commit,
                "command_paths_from_acquisition_receipts": paths,
                "command_path_expansion": "Expand each %EnvironmentVariable% path with [Environment]::ExpandEnvironmentVariables and resolve literally; the collectors select explicit executables/manifests rather than PATH aliases.",
                "environment": {"os": "Windows11 x64 build26300", "os_inventory": self.inventory,
                                "standard_user_non_elevated": True, "ps51": shells[0]["shell_version"], "ps7": shells[1]["shell_version"],
                                "test_process_policy": "RemoteSigned explicitly authorized for test process only",
                                "current_supported_ps7_build": "7.6.6, official current-LTS observation recorded in separate PS7 acquisition receipt",
                                "os_support_channel_established": False,
                                "native_ghostscript_executed": True, "persistent_policy_security_parent_environment_changed": False,
                                "owned_synthetic_file_acl": "Run-owned read-denial fixtures only; original descriptor/SDDL/source snapshot restored in finally in each actual shell."},
                "static_analysis": static,
                "supporting_development_checks": self.helper_results,
                "historical_findings": {"default_GS_behavior": "Dirty historical GS suite had11passes/1failure: an inaccessible password-protected input could yield exit0 and a placeholder output under default GS behavior. This historical failure is preserved separately and excluded from clean pass totals.",
                                        "direct_flag_probe": "Historical actual -dPDFSTOPONERROR probe produced exit1 while an owned output existed; output existence cannot establish success. A separate prelaunch slash-path harness fault is retained without a native execution claim.",
                                        "clean_fix_evidence": "The current GS wrapper includes fixed -dPDFSTOPONERROR; clean13-case GS suites in both actual shells verify the bounded encrypted-input result without publishing/claiming an invalid output."},
                "executions": [{"shell": label, "tier": summary["tier"], "exit_code": 0,
                                "summary": summary, "reports_manifest": manifest_relative} for label, summary in self.summaries],
                "limitations": ["Native page totals and unchanged synthetic source snapshots do not certify visible PDF fidelity, feature preservation, or signature validity.",
                                "Unsupported PDFtk Unicode/code-page and conservative file operands at260characters are safe documented limitations; no source renaming/chunking/shell workaround was added.",
                                "Actual ACL denial, exclusive-lock, encrypted-input and existing-output protection remain distinct integration observations.",
                                "PSScriptAnalyzer default static findings retain20warnings/4information items in each shell with0errors; warnings were not silently converted into native test passes.",
                                "General PDF validation, destination overlap/final staging, email outcomes, descendant interruption cleanup, final fidelity, Explorer/manual acceptance, CI/package and release gates remain downstream.",
                                "Windows OS support channel remains unestablished; no blanket OS-platform support or project/release completion claim follows."]}

    def finish(self, shells, destination, check_only):
        destination = destination.resolve()
        require(destination.is_relative_to(self.repo / "docs/codex/evidence"), "Public output must stay beneath evidence root")
        results_path = self.repo / "docs/codex/evidence/T09-C3-results.json"
        results_payload = json_bytes(self.results_document(shells, destination))
        manifest = {"schema_version": 1, "task": "T09", "checkpoint": "C3",
                    "commit_under_test": self.commit, "dirty_worktree": False,
                    "source_binding": "Every clean suite summary/native observation and collector aggregate explicitly names this SHA and clean state",
                    "xml_redactions": ["environment." + key for key in XML_IDENTITY],
                    "native_json_redactions": "Recursively parsed string values: repository, HOME/USERPROFILE, LOCALAPPDATA/APPDATA and TEMP/TMP prefixes, including mixed/escaped separators in console/log text",
                    "exact_byte_copies": "Original summary.json and controlled build-info.json bytes; raw XML/native JSON digests retained beside sanitized public digests",
                    "historical_rule": "Historical dirty/failed records and direct probe harness faults are separate records and excluded from clean totals",
                    "scope": "Native PDFtk/GS path and bounded noninteractive integration plus focused M1 regressions; controlled fixtures cannot certify engine support, final fidelity, Explorer or release gates",
                    "shells": shells, "total_clean_passed": sum(shell["passed"] for shell in shells),
                    "results_file": str(results_path.relative_to(self.repo)).replace("\\", "/"),
                    "results_sha256": sha(results_payload),
                    "clean_reports": len(shells) * len(TIERS), "records": self.records}
        self.add_payload("manifest.json", json_bytes(manifest))
        for name, payload in self.payloads.items():
            self.privacy_gate(payload, name)
        self.privacy_gate(results_payload, results_path.name)
        result = {"clean_commit": self.commit, "clean_reports": manifest["clean_reports"],
                  "total_passed": manifest["total_clean_passed"], "per_shell": shells,
                  "historical_records": len([r for r in self.records if r["classification"].startswith("historical")]),
                  "public_files": len(self.payloads) + 1, "manifest_sha256": sha(self.payloads["manifest.json"]),
                  "results_sha256": sha(results_payload),
                  "check_only": check_only}
        if not check_only:
            require(not destination.exists(), "Destination exists; retain prior evidence and select a new destination")
            require(not results_path.exists(), "Results file exists; never overwrite earlier evidence")
            destination.mkdir()
            for name, payload in self.payloads.items():
                with (destination / name).open("xb") as target:
                    target.write(payload)
                require(sha((destination / name).read_bytes()) == sha(payload), "Written byte hash mismatch")
            with results_path.open("xb") as target:
                target.write(results_payload)
            require(sha(results_path.read_bytes()) == sha(results_payload), "Written results byte hash mismatch")
            result["destination"] = str(destination.relative_to(self.repo)).replace("\\", "/")
        print(json.dumps(result, indent=2))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=Path.cwd())
    parser.add_argument("--commit", required=True)
    parser.add_argument("--ps51-root", type=Path, required=True)
    parser.add_argument("--ps7-root", type=Path, required=True)
    parser.add_argument("--destination", type=Path)
    parser.add_argument("--analyzer-ps51", type=Path)
    parser.add_argument("--analyzer-ps7", type=Path)
    parser.add_argument("--environment", type=Path)
    parser.add_argument("--python-root", type=Path)
    parser.add_argument("--check-only", action="store_true", help="Validate all input/privacy/hash gates without public writes")
    args = parser.parse_args()
    require(re.fullmatch(r"[0-9a-f]{40}", args.commit) is not None, "Explicit full lowercase implementation SHA required")
    repo = args.repo.resolve()
    head = subprocess.run(["git", "-C", str(repo), "rev-parse", "HEAD"], check=True, capture_output=True, text=True).stdout.strip()
    dirty = subprocess.run(["git", "-C", str(repo), "status", "--porcelain=v1"], check=True, capture_output=True, text=True).stdout
    require(head == args.commit and not dirty, "Evidence collection must start at requested implementation SHA with clean tracked worktree")
    collector = Collector(repo, args.commit)
    shells = [collector.clean_shell("ps51", args.ps51_root), collector.clean_shell("ps7", args.ps7_root)]
    collector.historical()
    collector.analyzer("ps51", args.analyzer_ps51 or collector.work / "T09-analyzer-C3-ps51.json", shells[0]["shell_version"])
    collector.analyzer("ps7", args.analyzer_ps7 or collector.work / "T09-analyzer-C3-ps7.json", shells[1]["shell_version"])
    collector.environment_and_pins(args.environment or collector.work / "T09-C3-environment.json")
    collector.python_helpers(args.python_root or collector.work / "T09-C3-python")
    destination = args.destination or repo / "docs/codex/evidence/T09-C3-reports"
    collector.finish(shells, destination, args.check_only)


if __name__ == "__main__":
    main()
