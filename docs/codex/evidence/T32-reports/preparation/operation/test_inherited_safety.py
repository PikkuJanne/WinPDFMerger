"""Focused safety regressions for the development-only candidate operation driver."""
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

MODULE_PATH = Path(__file__).with_name("final_package_smoke.py")
SPEC = importlib.util.spec_from_file_location("candidate_smoke", MODULE_PATH)
smoke = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(smoke)


class CandidateSafety(unittest.TestCase):
    def make_package(self, extra=None, inventory=None, source="a" * 40):
        payload = {"WinPDFMerge.ps1": b"synthetic package test bytes\r\n", "VERSION": b"1.0.0\n"}
        rows = [{"path": path, "sha256": smoke.digest(raw)} for path, raw in payload.items()]
        info = {"version": "1.0.0", "source_commit": source, "build_environment": {"test": "synthetic unit fixture"},
                "files": rows if inventory is None else inventory}
        content = {**payload, "BUILD_INFO.json": json.dumps(info).encode(), **(extra or {})}
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, "w") as archive:
            for path, raw in content.items():
                entry = zipfile.ZipInfo(smoke.PACKAGE_ROOT + "/" + path, date_time=(2000, 1, 1, 0, 0, 0))
                entry.external_attr = 0x20  # zipfile substitutes mode if zero; remove below.
                archive.writestr(entry, raw)
        raw = bytearray(stream.getvalue())
        offset = 0
        while True:
            offset = raw.find(b"PK\x01\x02", offset)
            if offset < 0:
                break
            raw[offset + 38:offset + 42] = b"\0\0\0\0"
            offset += 46
        raw = bytes(raw)
        checksum = (smoke.digest(raw) + "  " + smoke.PACKAGE_ROOT + ".zip\n").encode()
        return raw, checksum, payload

    def validate(self, raw, checksums, payload, **kwargs):
        return smoke.validate_package(raw, kwargs.get("expected_hash", smoke.digest(raw)), checksums,
                                      kwargs.get("checksums_hash", smoke.digest(checksums)),
                                      kwargs.get("source", "a" * 40), payload)

    def test_import_has_no_application_or_capture_side_effects(self):
        self.assertTrue(callable(smoke.main))
        self.assertTrue(callable(smoke.run))

    def test_git_forward_slash_root_names_same_actual_windows_path(self):
        root = Path.cwd()
        self.assertEqual(Path(str(root).replace("\\", "/")), root)

    def test_safe_entries_refuse_traversal_absolute_alias_and_devices(self):
        for value in ("../x", "/x", "C:/x", "x\\y", "x//y", "x/./y", "x./y", "NUL.pdf", "x/COM1.txt"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                smoke.safe_entry(value)

    def test_exact_assets_and_inventory_validate(self):
        raw, checksums, payload = self.make_package()
        info, files = self.validate(raw, checksums, payload)
        self.assertEqual(info["source_commit"], "a" * 40)
        self.assertEqual(set(files), set(payload) | {"BUILD_INFO.json"})

    def test_independently_accepted_zip_hash_cannot_be_redefined_by_manifest(self):
        raw, checksums, payload = self.make_package()
        with self.assertRaises(RuntimeError):
            self.validate(raw, checksums, payload, expected_hash="b" * 64)

    def test_whole_manifest_hash_is_an_independent_gate(self):
        raw, checksums, payload = self.make_package()
        with self.assertRaises(RuntimeError):
            self.validate(raw, checksums, payload, checksums_hash="b" * 64)

    def test_unreviewed_file_is_refused_even_with_updated_exact_asset_hash(self):
        raw, checksums, payload = self.make_package({"tests/private.pdf": b"synthetic"})
        with self.assertRaises(RuntimeError):
            self.validate(raw, checksums, payload)

    def test_stale_source_and_duplicate_inventory_are_refused(self):
        raw, checksums, payload = self.make_package(source="b" * 40)
        with self.assertRaises(RuntimeError):
            self.validate(raw, checksums, payload)
        raw, checksums, payload = self.make_package(inventory=[{"path": "VERSION", "sha256": "c" * 64}] * 2)
        with self.assertRaises(RuntimeError):
            self.validate(raw, checksums, payload)

    def test_git_blob_byte_changes_are_refused(self):
        raw, checksums, payload = self.make_package()
        payload["WinPDFMerge.ps1"] = payload["WinPDFMerge.ps1"].replace(b"\r\n", b"\n")
        with self.assertRaises(RuntimeError):
            self.validate(raw, checksums, payload)

    def test_extraction_requires_new_directory_and_preserves_exact_bytes(self):
        with tempfile.TemporaryDirectory(prefix="T29 unit ") as directory:
            target = Path(directory) / "fresh extraction"
            contents = {"WinPDFMerge.ps1": b"unit bytes\r\n", "docs/USAGE.md": b"unit instructions\n"}
            app = smoke.extract(contents, target)
            self.assertEqual((app / "WinPDFMerge.ps1").read_bytes(), contents["WinPDFMerge.ps1"])
            with self.assertRaises(FileExistsError):
                smoke.extract(contents, target)

    def test_snapshots_detect_source_byte_and_file_set_changes(self):
        with tempfile.TemporaryDirectory(prefix="T29 guard ") as directory:
            root = Path(directory)
            first = root / "source.pdf"
            first.write_bytes(b"original synthetic")
            before = smoke.snapshot([first], root)
            first.write_bytes(b"changed synthetic")
            self.assertNotEqual(smoke.snapshot([first], root), before)
            first.write_bytes(b"original synthetic")
            second = root / "extra.pdf"
            second.write_bytes(b"foreign synthetic")
            self.assertNotEqual(smoke.snapshot([first, second], root), before)

    def test_recursive_inventory_refuses_new_file_and_empty_directory(self):
        with tempfile.TemporaryDirectory(prefix="T29 inventory ") as directory:
            root = Path(directory)
            (root / "src").mkdir()
            (root / "src/helper.ps1").write_bytes(b"synthetic")
            before = smoke.tree_inventory(root)
            (root / "src/extra.ps1").write_bytes(b"foreign synthetic")
            with self.assertRaises(RuntimeError):
                smoke.require_inventory(smoke.tree_inventory(root), before)
            (root / "src/extra.ps1").unlink()
            (root / "unexpected empty directory").mkdir()
            with self.assertRaises(RuntimeError):
                smoke.require_inventory(smoke.tree_inventory(root), before)

    def test_email_result_uses_actual_package_log_label_and_unique_state(self):
        smoke.require_email_result("Email result: no_size_benefit\r\n", "no_size_benefit")
        smoke.require_email_result("Email result: failed\r\n", "failed")
        for log in ("Email state: failed\r\n", "Email result: published\r\n", "Email result: failed\nEmail result: failed\n"):
            with self.subTest(log=log), self.assertRaises(RuntimeError):
                smoke.require_email_result(log, "failed")

    def test_raster_recipe_is_deterministic_original_pdf(self):
        first = smoke.raster_pdf()
        self.assertEqual(first, smoke.raster_pdf())
        self.assertTrue(first.startswith(b"%PDF-1.4"))
        self.assertIn(b"T03-14-P01", first)


if __name__ == "__main__":
    unittest.main()
