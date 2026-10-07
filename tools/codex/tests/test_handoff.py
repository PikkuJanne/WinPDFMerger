"""Offline helper tests; do not certify Windows, PDF merging or GitHub publication."""
from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import stat
import subprocess
import tempfile
import unittest
from unittest import mock
import zipfile

HELPER = Path(__file__).resolve().parents[1] / "handoff.py"
spec = importlib.util.spec_from_file_location("handoff_under_test", HELPER)
assert spec is not None and spec.loader is not None
h = importlib.util.module_from_spec(spec)
spec.loader.exec_module(h)


def write_json(path: Path, data: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data), encoding="utf-8")


def make_bundle(root: Path, payload: dict[str, str] | None = None) -> Path:
    root.mkdir(parents=True)
    payload = payload if payload is not None else {"AGENTS.md": "rules\n", "docs/codex/INDEX.md": "index\n"}
    for name, text in payload.items():
        p = root / "payload" / name
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text, encoding="utf-8")
    files = [{"path": p.relative_to(root).as_posix(), "sha256": h.digest(p)}
             for p in sorted(root.rglob("*")) if p.is_file()]
    write_json(root / "BUNDLE_MANIFEST.json", {"schema_version": 1, "repository": h.REPOSITORY, "files": files})
    return root


def run_git(repo: Path, *args: str) -> str:
    result = subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True, check=True)
    return result.stdout.strip()


def make_repo(root: Path) -> Path:
    root.mkdir()
    run_git(root, "init", "-b", "main")
    run_git(root, "config", "user.name", "Synthetic Test")
    run_git(root, "config", "user.email", "synthetic@example.invalid")
    run_git(root, "remote", "add", "origin", "https://github.com/PikkuJanne/WinPDFMerger.git")
    run_git(root, "config", "branch.main.remote", "origin")
    run_git(root, "config", "branch.main.merge", "refs/heads/main")
    (root / "README.md").write_text("Synthetic local test repository.\n", encoding="utf-8")
    run_git(root, "add", "README.md"); run_git(root, "commit", "-m", "synthetic baseline")
    return root


def package(path: Path, commit: str = "a" * 40, extra: dict[str, bytes] | None = None) -> None:
    content = {"WinPDFMerge.ps1": b"# synthetic\n", "WinPDFMerge.bat": b"@echo off\n",
               "README.md": b"synthetic instructions\n", "LICENSE": b"synthetic license\n"}
    content.update(extra or {})
    info = {"version": "1.0.0", "source_commit": commit,
            "files": [{"path": n, "sha256": hashlib.sha256(data).hexdigest()} for n, data in content.items()]}
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in content.items():
            archive.writestr(f"{h.PACKAGE_ROOT}/{name}", data)
        archive.writestr(f"{h.PACKAGE_ROOT}/BUILD_INFO.json", json.dumps(info))


class Helpers(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_safe_paths_accept_normal(self) -> None:
        self.assertEqual(str(h.safe_relative("docs/codex/T01.md")), "docs/codex/T01.md")

    def test_safe_paths_reject_traversal_windows_ambiguity(self) -> None:
        for name in (".", "", "../x", "/x", "a/../b", "a//b", "a\\b", "C:x", ".git/config",
                     "a. ", "NUL.txt", "docs/con", "a?b", "a\x01b"):
            with self.subTest(name=name), self.assertRaises(h.HandoffError):
                h.safe_relative(name)

    def test_canonical_remotes(self) -> None:
        for url in ("git@github.com:PikkuJanne/WinPDFMerger.git", "ssh://git@github.com/PikkuJanne/WinPDFMerger.git",
                    "https://github.com/pikkujanne/winpdfmerger", "https://github.com/PikkuJanne/WinPDFMerger.git"):
            self.assertEqual(h.canonical_remote(url), h.REPOSITORY)

    def test_remotes_reject_credentials_wrong_target_host(self) -> None:
        for url in ("https://token@github.com/PikkuJanne/WinPDFMerger.git", "https://github.com/other/repo",
                    "https://github.com.evil.invalid/PikkuJanne/WinPDFMerger", "file:///tmp/repo",
                    "ssh://root@github.com/PikkuJanne/WinPDFMerger.git"):
            with self.subTest(url=url), self.assertRaises(h.HandoffError):
                h.canonical_remote(url)

    def test_manifest_success_and_tamper(self) -> None:
        bundle = make_bundle(self.root / "bundle")
        self.assertTrue(h.verify_bundle(bundle)["verified"])
        (bundle / "payload/AGENTS.md").write_text("different", encoding="utf-8")
        with self.assertRaises(h.HandoffError):
            h.verify_bundle(bundle)

    def test_manifest_extra_file_refused(self) -> None:
        bundle = make_bundle(self.root / "bundle")
        (bundle / "unexpected.txt").write_text("extra", encoding="utf-8")
        with self.assertRaises(h.HandoffError):
            h.verify_bundle(bundle)

    def test_manifest_traversal_refused(self) -> None:
        bundle = make_bundle(self.root / "bundle")
        manifest = h.read_json(bundle / "BUNDLE_MANIFEST.json")
        manifest["files"][0]["path"] = "../outside"
        write_json(bundle / "BUNDLE_MANIFEST.json", manifest)
        with self.assertRaises(h.HandoffError):
            h.verify_bundle(bundle)

    def test_manifest_case_collision_refused(self) -> None:
        bundle = make_bundle(self.root / "bundle")
        manifest = h.read_json(bundle / "BUNDLE_MANIFEST.json")
        manifest["files"].append({"path": "PAYLOAD/agents.md", "sha256": "0" * 64})
        write_json(bundle / "BUNDLE_MANIFEST.json", manifest)
        with self.assertRaises(h.HandoffError):
            h.verify_bundle(bundle)

    @unittest.skipUnless(shutil.which("git"), "Git is required for local repo tests")
    def test_preview_never_writes(self) -> None:
        bundle, repo = make_bundle(self.root / "bundle"), make_repo(self.root / "repo")
        before = h.inventory(repo)
        result = h.import_bundle(bundle, repo)
        self.assertEqual(result["mode"], "preview")
        self.assertEqual(h.inventory(repo), before)
        self.assertEqual(run_git(repo, "status", "--porcelain"), "")

    @unittest.skipUnless(shutil.which("git"), "Git required")
    def test_apply_create_only_and_idempotent(self) -> None:
        bundle, repo = make_bundle(self.root / "bundle"), make_repo(self.root / "repo")
        original = h.digest(repo / "README.md")
        self.assertEqual(h.import_bundle(bundle, repo, True)["mode"], "applied")
        self.assertEqual(h.digest(repo / "README.md"), original)
        run_git(repo, "add", "AGENTS.md", "docs"); run_git(repo, "commit", "-m", "import synthetic handoff")
        result = h.import_bundle(bundle, repo, True)
        self.assertTrue(all(x["action"] == "skip-identical" for x in result["files"]))
        self.assertEqual(run_git(repo, "status", "--porcelain"), "")

    @unittest.skipUnless(shutil.which("git"), "Git required")
    def test_conflict_preflights_before_any_write(self) -> None:
        bundle, repo = make_bundle(self.root / "bundle"), make_repo(self.root / "repo")
        (repo / "AGENTS.md").write_text("existing instructions", encoding="utf-8")
        run_git(repo, "add", "AGENTS.md"); run_git(repo, "commit", "-m", "existing rules")
        with self.assertRaises(h.HandoffError):
            h.import_bundle(bundle, repo, True)
        self.assertFalse((repo / "docs").exists())
        self.assertEqual((repo / "AGENTS.md").read_text(), "existing instructions")

    @unittest.skipUnless(shutil.which("git"), "Git required")
    def test_dirty_apply_refused_preview_allowed(self) -> None:
        bundle, repo = make_bundle(self.root / "bundle"), make_repo(self.root / "repo")
        (repo / "unrelated.txt").write_text("preserve", encoding="utf-8")
        self.assertEqual(h.import_bundle(bundle, repo)["mode"], "preview")
        with self.assertRaises(h.HandoffError):
            h.import_bundle(bundle, repo, True)
        self.assertFalse((repo / "AGENTS.md").exists())

    @unittest.skipUnless(shutil.which("git"), "Git required")
    def test_wrong_push_target_refused(self) -> None:
        bundle, repo = make_bundle(self.root / "bundle"), make_repo(self.root / "repo")
        run_git(repo, "remote", "set-url", "--push", "origin", "https://github.com/other/repo.git")
        with self.assertRaises(h.HandoffError):
            h.import_bundle(bundle, repo, True)
        self.assertFalse((repo / "AGENTS.md").exists())

    @unittest.skipUnless(shutil.which("git"), "Git required")
    def test_runtime_payload_refused(self) -> None:
        bundle = make_bundle(self.root / "bundle", {"WinPDFMerge.ps1": "replacement"})
        repo = make_repo(self.root / "repo")
        with self.assertRaises(h.HandoffError):
            h.import_bundle(bundle, repo, True)
        self.assertFalse((repo / "WinPDFMerge.ps1").exists())

    def test_case_only_existing_path_refused(self) -> None:
        (self.root / "agents.md").write_text("existing", encoding="utf-8")
        with self.assertRaises(h.HandoffError):
            h.reject_case_alias(self.root, "AGENTS.md")

    def test_symlink_refused(self) -> None:
        target = self.root / "target"; target.mkdir()
        try:
            (self.root / "alias").symlink_to(target, target_is_directory=True)
        except OSError:
            self.skipTest("Symlink creation not permitted")
        with self.assertRaises(h.HandoffError):
            h.reject_reparse_chain(self.root / "alias" / "new-file")

    def test_windows_junction_attribute_refused(self) -> None:
        fake = mock.Mock(st_mode=stat.S_IFDIR, st_file_attributes=0x400)
        with mock.patch.object(Path, "lstat", return_value=fake), self.assertRaises(h.HandoffError):
            h.reject_reparse_chain(self.root)

    @unittest.skipUnless(shutil.which("git"), "Git required")
    def test_live_sync_clean_match_and_mismatch(self) -> None:
        repo = make_repo(self.root / "repo")
        head = run_git(repo, "rev-parse", "HEAD")
        original = h.git
        def reading(path: Path, *args: str) -> str:
            if args[0] == "ls-remote":
                return f"{head}\trefs/heads/main"
            return original(path, *args)
        with mock.patch.object(h, "git", side_effect=reading):
            self.assertTrue(h.sync(repo)["synchronized"])
            (repo / "dirty.txt").write_text("dirty", encoding="utf-8")
            self.assertFalse(h.sync(repo)["synchronized"])
        (repo / "dirty.txt").unlink()
        def mismatch(path: Path, *args: str) -> str:
            return f"{'b'*40}\trefs/heads/main" if args[0] == "ls-remote" else original(path, *args)
        with mock.patch.object(h, "git", side_effect=mismatch):
            self.assertFalse(h.sync(repo)["synchronized"])

    @unittest.skipUnless(shutil.which("git"), "Git required")
    def test_sync_network_failure_is_not_success(self) -> None:
        repo = make_repo(self.root / "repo")
        original = h.git
        def unavailable(path: Path, *args: str) -> str:
            if args[0] == "ls-remote":
                raise h.HandoffError("network unavailable")
            return original(path, *args)
        with mock.patch.object(h, "git", side_effect=unavailable), self.assertRaises(h.HandoffError):
            h.sync(repo)

    def test_live_ref_parser_rejects_bad_or_duplicate(self) -> None:
        for data in ("bad refs/heads/main", f"{'a'*40} refs/heads/main\n{'b'*40} refs/heads/main"):
            with self.assertRaises(h.HandoffError):
                h.parse_refs(data)

    def test_package_valid_with_build_inventory(self) -> None:
        path = self.root / "good.zip"; package(path)
        result = h.verify_package(path, "a" * 40)
        self.assertTrue(result["safe_structure"])
        self.assertFalse(result["application_executed"])

    def test_package_traversal_or_dev_content_refused(self) -> None:
        for extra in ({"../escape.txt": b"no"}, {"docs/codex/STATUS.md": b"no"}, {"tool.exe": b"no"}):
            path = self.root / "bad.zip"; package(path, extra=extra)
            with self.assertRaises(h.HandoffError):
                h.verify_package(path, "a" * 40)

    def test_package_wrong_commit_or_file_hash_refused(self) -> None:
        path = self.root / "bad.zip"; package(path)
        with self.assertRaises(h.HandoffError):
            h.verify_package(path, "b" * 40)
        # Replace the declared hashes without introducing duplicate ZIP entries.
        with zipfile.ZipFile(path) as source:
            contents = {x.filename: source.read(x) for x in source.infolist()}
        key = f"{h.PACKAGE_ROOT}/BUILD_INFO.json"
        info = json.loads(contents[key]); info["files"][0]["sha256"] = "0" * 64
        contents[key] = json.dumps(info).encode()
        with zipfile.ZipFile(path, "w") as dest:
            for name, data in contents.items(): dest.writestr(name, data)
        with self.assertRaises(h.HandoffError):
            h.verify_package(path, "a" * 40)

    def test_only_one_actual_final_public_release(self) -> None:
        release = {"id": 123, "tag_name": h.TAG, "draft": False, "prerelease": False,
                   "published_at": "2026-10-07T12:00:00Z", "html_url": f"https://github.com/{h.REPOSITORY}/releases/tag/{h.TAG}"}
        self.assertEqual(h.select_public_release([release])["id"], 123)
        for listing in ([], [release, dict(release, id=124, tag_name="v0.9.0")],
                        [dict(release, draft=True)], [dict(release, prerelease=True)], [dict(release, published_at=None)]):
            with self.assertRaises(h.HandoffError): h.select_public_release(listing)

    def test_public_pagination_includes_all_pages(self) -> None:
        with mock.patch.object(h, "public_json", side_effect=[[{"id": i} for i in range(100)], [{"id": 100}]]):
            self.assertEqual(len(h.public_pages("releases")), 101)

    def test_release_verification_offline_full_flow(self) -> None:
        # Synthetic archive/metadata and mocked public reads; no actual GitHub publish.
        original = self.root / "accepted.zip"; package(original)
        zhash = h.digest(original)
        sums = f"{zhash}  {h.ZIP_NAME}\n".encode()
        shash = hashlib.sha256(sums).hexdigest()
        release = {"id": 123, "tag_name": h.TAG, "draft": False, "prerelease": False,
                   "published_at": "2026-10-07T12:00:00Z", "html_url": f"https://github.com/{h.REPOSITORY}/releases/tag/{h.TAG}"}
        assets = [{"name": h.ZIP_NAME}, {"name": h.SUM_NAME}]
        def pages(endpoint: str): return [release] if endpoint == "releases" else assets
        def download(asset: dict, directory: Path) -> Path:
            p = directory / asset["name"]
            p.write_bytes(original.read_bytes() if asset["name"] == h.ZIP_NAME else sums)
            return p
        refs = f"{'c'*40}\trefs/tags/{h.TAG}\n{'a'*40}\trefs/tags/{h.TAG}^{{}}"
        with mock.patch.object(h, "inspect_repo", return_value={"dirty": False, "root": str(self.root / "repo")}), \
             mock.patch.object(h, "git", return_value=refs), mock.patch.object(h, "public_pages", side_effect=pages), \
             mock.patch.object(h, "download_asset", side_effect=download):
            result = h.verify_release(self.root / "repo", "a"*40, zhash, shash, self.root / "download")
            self.assertTrue(result["public_release_verified"])
            self.assertEqual(result["windows_downloaded_package_smoke_test"], "STILL REQUIRED")
            with self.assertRaises(h.HandoffError):
                h.verify_release(self.root / "repo", "a"*40, "0"*64, shash, self.root / "wrong-hash")
            with self.assertRaises(h.HandoffError):
                h.verify_release(self.root / "repo", "a"*40, zhash, shash, self.root / "download")

    def test_plan_initial_consistency_and_gate_fail_closed(self) -> None:
        # Locate the installed payload or repository from either bundled test copy.
        candidates = [HELPER.parents[1] / "payload", HELPER.parents[2]]
        repo = next((p for p in candidates if (p / "docs/codex/TASKS.json").is_file()), None)
        if repo is None:
            self.skipTest("Plan files not present beside helper")
        result = h.check_plan(repo)
        self.assertEqual(result["task_count"], 34)
        # Initial supplied records must not pass a release gate.
        if result["done_tasks"] == 0:
            with self.assertRaises(h.HandoffError): h.check_plan(repo, "ready")


if __name__ == "__main__":
    unittest.main()
