"""Corpus tooling regressions; real PDFium inspection, never application acceptance."""
from __future__ import annotations

from contextlib import redirect_stdout
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

SOURCE = Path(__file__).resolve().parents[1] / "corpus.py"
SPEC = importlib.util.spec_from_file_location("corpus", SOURCE)
corpus = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(corpus)


class CorpusTests(unittest.TestCase):
    def test_catalog_reconstructs_exactly_from_pinned_original_recipes(self):
        expected = corpus.catalog_recipe()
        self.assertEqual(corpus.load_catalog(), expected)
        self.assertEqual(set(expected["groups"]), {"numbered", "envelopes", "features", "presets"})
        self.assertEqual(len(expected["safety"]["scenarios"]), 12)

    def test_tampered_catalog_is_rejected(self):
        expected = corpus.catalog_recipe()
        with mock.patch.object(Path, "read_text", return_value=json.dumps({**expected, "development_only": False})):
            with mock.patch.object(corpus, "catalog_recipe", return_value=expected):
                with self.assertRaisesRegex(ValueError, "catalog differs"):
                    corpus.load_catalog()

    def test_boolean_page_count_cannot_equal_integer_catalog_count(self):
        expected = corpus.catalog_recipe()
        altered = json.loads(json.dumps(expected))
        altered["groups"]["numbered"]["fixtures"][0]["page_count"] = True
        with mock.patch.object(Path, "read_text", return_value=json.dumps(altered)), \
                mock.patch.object(corpus, "catalog_recipe", return_value=expected):
            with self.assertRaisesRegex(ValueError, "catalog differs"):
                corpus.load_catalog()

    def test_intended_inventory_is_fixed_and_contains_nested_and_empty_sentinels(self):
        paths = corpus.intended_paths(corpus.load_catalog())
        self.assertIn("safety/matrix/nested", paths)
        self.assertIn("safety/matrix/nested/3.PDF", paths)
        self.assertIn("safety/matrix/directory.pdf", paths)
        self.assertIn("features/generation.json", paths)
        self.assertNotIn("safety/matrix/unexpected.pdf", paths)

    def test_type_confused_generated_receipt_is_rejected_before_inspection(self):
        catalog = corpus.load_catalog()
        for section in ("groups", "safety"):
            with self.subTest(section=section), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                receipt = {"schema_version": 1, "catalog_sha256": corpus.sha256(corpus.CATALOG.read_bytes()),
                           "versions": corpus.PINS, "groups": json.loads(json.dumps(catalog["groups"])),
                           "safety": json.loads(json.dumps(catalog["safety"])), "inventory": []}
                if section == "groups":
                    receipt["groups"]["numbered"]["fixtures"][0]["page_count"] = True
                else:
                    receipt["safety"]["scenarios"]["single-uppercase"]["expected_page_count"] = True
                (root / "corpus.json").write_text(json.dumps(receipt), encoding="utf-8")
                with mock.patch.object(corpus, "inspect_pdf") as inspection:
                    with self.assertRaisesRegex(ValueError, "receipt differs"):
                        corpus.verify(root)
                    inspection.assert_not_called()

    def test_safety_pdf_bytes_and_encrypted_bytes_repeat_exactly(self):
        first_scenarios, first_files = corpus.safety_recipes()
        second_scenarios, second_files = corpus.safety_recipes()
        self.assertEqual(first_scenarios, second_scenarios)
        self.assertEqual(first_files, second_files)
        self.assertEqual(first_scenarios["matrix"]["expected_page_count"], 19)
        self.assertEqual(first_scenarios["single-uppercase"]["ordered_names"], ["single.PDF"])
        self.assertNotIn("hidden.PDF", first_scenarios["matrix"]["ordered_names"])

    def test_encrypted_sources_have_recorded_synthetic_passwords(self):
        from pypdf import PdfReader
        _, files = corpus.safety_recipes()
        for kind, password in (("encrypted", "T21-user"), ("owner-restricted", "")):
            reader = PdfReader(io.BytesIO(files[f"safety/invalid/{kind}/5_{kind}.pdf"]), strict=True)
            self.assertTrue(reader.is_encrypted)
            self.assertTrue(reader.decrypt(password))
            self.assertEqual(len(reader.pages), 1)
            self.assertIn("T03-21-P92", reader.pages[0].extract_text())

    def test_expected_orders_are_explicit_and_identifier_totals_consistent(self):
        catalog = corpus.load_catalog()
        self.assertEqual(catalog["groups"]["presets"]["merge_order"], ["mixed.pdf", "scan.pdf", "small-print.pdf"])
        self.assertEqual(catalog["groups"]["envelopes"]["merge_order"], ["conventional-cr.pdf", "conventional-lf.pdf", "incremental.pdf", "linearized.pdf", "xref-stream.pdf"])
        for scenario in catalog["safety"]["scenarios"].values():
            if scenario["expected_exit_code"]:
                continue
            by_name = {row["path"]: row for row in scenario["source_files"]}
            ids = [value for name in scenario["ordered_names"] for value in by_name[name]["page_identifiers"]]
            self.assertEqual(ids, scenario["expected_page_identifiers"])
            self.assertEqual(len(ids), scenario["expected_page_count"])

    def test_unsafe_inventory_paths_are_rejected(self):
        for value in ("", "../private.pdf", "/private.pdf", "C:/private.pdf", "folder\\private.pdf",
                      "folder/../private.pdf", "folder//private.pdf", "folder/./private.pdf", "1.pdf:stream",
                      "NUL.pdf", "folder/COM1.txt", "trailing./1.pdf", "trailing /1.pdf"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                corpus.safe_relative(value)
        self.assertEqual(corpus.safe_relative("nested/3.PDF"), Path("nested/3.PDF"))
        self.assertEqual(corpus.safe_relative("2 [x] ! & (y's).pdf"), Path("2 [x] ! & (y's).pdf"))

    def test_inspection_accepts_repeated_ids_without_weakening_order(self):
        from pypdf import PdfReader, PdfWriter
        original = corpus.REPO / "tests/fixtures/numbered/1.pdf"
        writer = PdfWriter()
        page = PdfReader(original).pages[0]
        writer.add_page(page)
        writer.add_page(page)
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "repeat.pdf"
            writer.write(path)
            before = path.read_bytes()
            result = corpus.inspect_pdf(path, ["T03-01-P01", "T03-01-P01"])
            self.assertEqual(result["page_count"], 2)
            self.assertEqual(result["pages"][0]["size_points"], [432, 288])
            self.assertEqual(path.read_bytes(), before)
            with self.assertRaisesRegex(ValueError, "identifier/order"):
                corpus.inspect_pdf(path, ["T03-01-P01", "T03-10-P01"])
            with self.assertRaisesRegex(ValueError, "Page count"):
                corpus.inspect_pdf(path, ["T03-01-P01"])

    def test_bad_identifier_expectations_fail_before_pdf_open(self):
        for value in ([], [True], ["invented"], "T03-01-P01", ["T03-01-P01suffix"]):
            with self.subTest(value=value), self.assertRaises(ValueError):
                corpus.inspect_pdf(Path("unopened.pdf"), value)

    def test_generation_refuses_existing_or_outside_directories_before_native_work(self):
        with mock.patch.object(corpus, "versions"), mock.patch.object(corpus, "load_catalog"):
            for path in (corpus.REPO, corpus.REPO / "tests/fixtures", corpus.REPO / "tests/.work"):
                with self.subTest(path=path), self.assertRaisesRegex(ValueError, "new explicitly owned"):
                    corpus.materialize(path, Path("unopened-gs.exe"))

    def test_inventory_detects_additions_changes_and_empty_directories(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / "directory.pdf").mkdir()
            (root / "1.pdf").write_bytes(b"original")
            before = corpus.inventory(root)
            self.assertEqual(len(before), 2)
            (root / "1.pdf").write_bytes(b"changed")
            self.assertNotEqual(corpus.inventory(root), before)
            (root / "nested").mkdir()
            (root / "nested/unexpected.txt").write_bytes(b"unexpected")
            self.assertEqual(len(corpus.inventory(root)), 4)

    def test_verify_rejects_changed_inventory_without_inspecting_or_rewriting(self):
        catalog = corpus.load_catalog()
        for mutation in ("hash", "hidden", "unexpected", "directory"):
            with self.subTest(mutation=mutation), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                (root / "1.pdf").write_bytes(b"original")
                (root / "directory.pdf").mkdir()
                recorded = corpus.inventory(root)
                receipt = {"schema_version": 1, "catalog_sha256": corpus.sha256(corpus.CATALOG.read_bytes()),
                           "versions": corpus.PINS, "groups": catalog["groups"], "safety": catalog["safety"], "inventory": recorded}
                (root / "corpus.json").write_text(json.dumps(receipt), encoding="utf-8")
                if mutation == "hash":
                    (root / "1.pdf").write_bytes(b"changed")
                elif mutation == "hidden":
                    receipt["inventory"][0]["hidden"] = True
                    (root / "corpus.json").write_text(json.dumps(receipt), encoding="utf-8")
                elif mutation == "unexpected":
                    (root / "unexpected.txt").write_bytes(b"unexpected")
                else:
                    (root / "directory.pdf").rmdir()
                before = corpus.inventory(root)
                with mock.patch.object(corpus, "inspect_pdf") as inspection:
                    with self.assertRaisesRegex(ValueError, "inventory/hash/attribute"):
                        corpus.verify(root)
                    inspection.assert_not_called()
                self.assertEqual(corpus.inventory(root), before)

    def test_cli_json_handles_legacy_windows_console_encoding(self):
        result = {"path": "2-漢字.pdf", "page_count": 1}
        raw = io.BytesIO()
        stream = io.TextIOWrapper(raw, encoding="cp1252", write_through=True)
        with mock.patch("sys.argv", ["corpus.py", "verify", "--root", "unopened"]), \
                mock.patch.object(corpus, "verify", return_value=result), redirect_stdout(stream):
            corpus.main()
        self.assertEqual(json.loads(raw.getvalue().decode("cp1252")), result)


if __name__ == "__main__":
    unittest.main()
