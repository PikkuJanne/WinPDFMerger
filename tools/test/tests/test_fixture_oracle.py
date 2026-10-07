"""T03 corpus checks use real PDFium; synthetic faults do not emulate PDFtk."""

from __future__ import annotations

import hashlib
import copy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock


TOOL_DIRECTORY = Path(__file__).resolve().parents[1]
FIXTURE_DIRECTORY = TOOL_DIRECTORY.parents[1] / "tests" / "fixtures" / "numbered"


def load_tool(name: str):
    specification = importlib.util.spec_from_file_location(name, TOOL_DIRECTORY / f"{name}.py")
    module = importlib.util.module_from_spec(specification)
    specification.loader.exec_module(module)
    return module


generator = load_tool("generate_numbered_fixtures")
oracle = load_tool("fixture_oracle")


class FixtureOracleTests(unittest.TestCase):
    def test_empty_manifest_cannot_report_native_inspection_success(self):
        manifest = {"schema_version": 1, "fixtures": [], "merge_order": [], "expected_merged_page_count": 0, "expected_merged_page_identifiers": []}
        with mock.patch.object(Path, "read_text", return_value=json.dumps(manifest)), mock.patch.object(oracle, "inspect_pdf") as inspect:
            with self.assertRaisesRegex(ValueError, "nonempty corpus"):
                oracle.verify_corpus(FIXTURE_DIRECTORY / "manifest.json")
            inspect.assert_not_called()

    def test_malformed_metadata_fails_before_native_inspection(self):
        original = json.loads((FIXTURE_DIRECTORY / "manifest.json").read_text(encoding="utf-8"))
        mutations = (
            ("boolean schema", lambda value: value.update(schema_version=True)),
            ("unknown schema", lambda value: value.update(schema_version=2)),
            ("traversal", lambda value: value["fixtures"][0].update(file="../1.pdf")),
            ("alternate stream", lambda value: value["fixtures"][0].update(file="1:stream.pdf")),
            ("duplicate filename", lambda value: value["fixtures"][1].update(file="1.PDF")),
            ("zero byte length", lambda value: value["fixtures"][0].update(bytes=0)),
            ("boolean page count", lambda value: value["fixtures"][0].update(page_count=True)),
            ("missing identifiers", lambda value: value["fixtures"][0].update(page_identifiers=[])),
            ("duplicate identifiers", lambda value: value["fixtures"][1].update(page_identifiers=["T03-02-P01", "T03-02-P01"])),
            ("bad hash", lambda value: value["fixtures"][0].update(sha256="bad")),
            ("bad dimensions", lambda value: value["fixtures"][0].update(page_size_points=[432, 0])),
            ("wrong total", lambda value: value.update(expected_merged_page_count=5)),
            ("wrong merged identifiers", lambda value: value.update(expected_merged_page_identifiers=list(reversed(value["expected_merged_page_identifiers"])))),
            ("wrong input order", lambda value: value.update(merge_order=list(reversed(value["merge_order"])))),
        )
        for label, mutate in mutations:
            with self.subTest(label=label):
                manifest = copy.deepcopy(original)
                mutate(manifest)
                with mock.patch.object(Path, "read_text", return_value=json.dumps(manifest)), mock.patch.object(oracle, "inspect_pdf") as inspect:
                    with self.assertRaises(ValueError):
                        oracle.verify_corpus(FIXTURE_DIRECTORY / "manifest.json")
                    inspect.assert_not_called()

    def test_pdf_cli_rejects_bad_schema_and_inconsistent_expectations(self):
        original = json.loads((FIXTURE_DIRECTORY / "manifest.json").read_text(encoding="utf-8"))
        for changes in ({"schema_version": True}, {"expected_merged_page_count": 5}):
            with self.subTest(changes=changes):
                manifest = copy.deepcopy(original)
                manifest.update(changes)
                with mock.patch("sys.argv", ["fixture_oracle.py", "--pdf", "unopened-result.pdf"]), mock.patch.object(Path, "read_text", return_value=json.dumps(manifest)), mock.patch.object(oracle, "inspect_pdf") as inspect:
                    with self.assertRaises(ValueError):
                        oracle.main()
                    inspect.assert_not_called()

    def test_committed_corpus_reproduces_exactly(self):
        for name, expected in generator.expected_files().items():
            self.assertEqual(expected, (FIXTURE_DIRECTORY / name).read_bytes(), name)

    def test_real_native_oracle_reads_total_order_and_sizes(self):
        result = oracle.verify_corpus(FIXTURE_DIRECTORY / "manifest.json")
        self.assertEqual(result["total_pages"], 4)
        self.assertEqual(result["page_identifiers"], ["T03-01-P01", "T03-02-P01", "T03-02-P02", "T03-10-P01"])
        self.assertFalse(result["pdftk_executed"])
        self.assertFalse(result["ghostscript_executed"])
        self.assertFalse(result["product_merge_executed"])

    def test_native_oracle_rejects_wrong_page_count(self):
        with self.assertRaisesRegex(ValueError, "Page count differs"):
            oracle.inspect_pdf(FIXTURE_DIRECTORY / "2.pdf", ["T03-02-P01"])

    def test_native_oracle_rejects_wrong_page_order(self):
        with self.assertRaisesRegex(ValueError, "identifier differs"):
            oracle.inspect_pdf(FIXTURE_DIRECTORY / "2.pdf", ["T03-02-P02", "T03-02-P01"])

    def test_hash_gate_detects_changed_fixture(self):
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            for name, content in generator.expected_files().items():
                (directory / name).write_bytes(content)
            path = directory / "1.pdf"
            path.write_bytes(path.read_bytes() + b"changed")
            with self.assertRaisesRegex(ValueError, "hash/length differs"):
                oracle.verify_corpus(directory / "manifest.json")

    def test_corpus_inspection_preserves_source_hashes(self):
        before = {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in FIXTURE_DIRECTORY.glob("*.pdf")}
        oracle.verify_corpus(FIXTURE_DIRECTORY / "manifest.json")
        after = {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in FIXTURE_DIRECTORY.glob("*.pdf")}
        self.assertEqual(before, after)

    def test_optional_render_creates_one_png_per_real_page(self):
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            oracle.inspect_pdf(FIXTURE_DIRECTORY / "2.pdf", ["T03-02-P01", "T03-02-P02"], directory)
            images = sorted(directory.glob("*.png"))
            self.assertEqual([path.name for path in images], ["2-page-1.png", "2-page-2.png"])
            self.assertTrue(all(path.read_bytes().startswith(b"\x89PNG\r\n\x1a\n") for path in images))


if __name__ == "__main__":
    unittest.main()
