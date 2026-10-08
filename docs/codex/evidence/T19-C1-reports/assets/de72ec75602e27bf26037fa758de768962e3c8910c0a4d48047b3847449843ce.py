#!/usr/bin/env python3
"""Development-only WinPDFMerger handoff helpers (Python 3.10+, standard library).

No command publishes releases, pushes Git, rewrites history, or runs downloaded
application code. Import is preview-only unless --apply is explicitly supplied.
The helpers validate evidence structure; they cannot certify unobserved PDF tests.
"""
from __future__ import annotations

import argparse
import hashlib
import hmac
import json
import os
from pathlib import Path, PurePosixPath
import re
import stat
import subprocess
import sys
from datetime import datetime, timezone
from typing import Any
import urllib.parse
import urllib.request
import zipfile

REPOSITORY = "PikkuJanne/WinPDFMerger"
TAG = "v1.0.0"
ZIP_NAME = "WinPDFMerger-v1.0.0.zip"
SUM_NAME = "SHA256SUMS.txt"
PACKAGE_ROOT = "WinPDFMerger-v1.0.0"
MAX_DOWNLOAD = 256 * 1024 * 1024
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")


class HandoffError(RuntimeError):
    """A failed safety, consistency, or external verification check."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise HandoffError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def safe_relative(value: str) -> PurePosixPath:
    require(isinstance(value, str) and bool(value), "Missing relative path.")
    require(not any(c in value for c in ("\\", ":", "\x00", "\r", "\n")), "Unsafe path syntax.")
    p = PurePosixPath(value)
    require(not p.is_absolute() and value == p.as_posix(), "Path must be normalized and relative.")
    require(bool(p.parts) and all(part not in ("", ".", "..") for part in p.parts), "Unsafe relative path component.")
    require(all(ord(c) >= 32 and c not in '<>"|?*' for c in value), "Invalid Windows path character.")
    devices = {"con", "prn", "aux", "nul", *(f"com{i}" for i in range(1, 10)), *(f"lpt{i}" for i in range(1, 10))}
    require(all(part.split(".")[0].casefold() not in devices for part in p.parts), "Reserved Windows device name.")
    require(all(part.rstrip(" .") == part for part in p.parts), "Ambiguous Windows path component.")
    require(all(part.casefold() != ".git" for part in p.parts), "Git internals are not allowed.")
    return p


def reject_reparse_chain(path: Path) -> None:
    """Refuse symlinks and Windows reparse points, including existing parents."""
    absolute = path.absolute()
    for item in [*reversed(absolute.parents), absolute]:
        try:
            info = item.lstat()
        except FileNotFoundError:
            continue
        require(not stat.S_ISLNK(info.st_mode), "Symlink path refused.")
        attributes = getattr(info, "st_file_attributes", 0)
        require(not (attributes & getattr(stat, "FILE_ATTRIBUTE_REPARSE_POINT", 0x400)),
                "Windows reparse-point/junction path refused.")


def read_json(path: Path) -> Any:
    reject_reparse_chain(path)
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise HandoffError(f"Cannot read valid JSON: {path.name} ({type(exc).__name__}).") from exc


def inventory(root: Path) -> set[str]:
    found: set[str] = set()
    for base, dirs, files in os.walk(root, followlinks=False):
        for name in [*dirs, *files]:
            reject_reparse_chain(Path(base) / name)
        for name in files:
            found.add((Path(base) / name).relative_to(root).as_posix())
    return found


def verify_bundle(bundle: Path) -> dict[str, Any]:
    reject_reparse_chain(bundle)
    bundle = bundle.absolute()
    require(bundle.is_dir(), "Bundle directory does not exist.")
    manifest = read_json(bundle / "BUNDLE_MANIFEST.json")
    require(manifest.get("schema_version") == 1, "Unsupported bundle manifest schema.")
    require(manifest.get("repository") == REPOSITORY, "Bundle targets another repository.")
    entries = manifest.get("files")
    require(isinstance(entries, list) and bool(entries), "Empty bundle manifest.")
    paths: set[str] = set()
    folded: set[str] = set()
    for item in entries:
        rel = safe_relative(item.get("path", "")).as_posix()
        require(rel != "BUNDLE_MANIFEST.json", "Manifest cannot hash itself.")
        require(rel.casefold() not in folded, "Duplicate/case-colliding manifest path.")
        folded.add(rel.casefold()); paths.add(rel)
        expected = item.get("sha256", "")
        require(bool(HEX64.fullmatch(expected)), f"Invalid SHA-256 for {rel}.")
        p = bundle / rel
        reject_reparse_chain(p)
        require(p.is_file(), f"Missing bundle file: {rel}")
        require(hmac.compare_digest(digest(p), expected.lower()), f"Bundle checksum mismatch: {rel}")
    actual = inventory(bundle) - {"BUNDLE_MANIFEST.json"}
    require(actual == paths, "Bundle inventory differs from manifest (missing or extra files).")
    return {"verified": True, "repository": REPOSITORY, "file_count": len(paths),
            "paths": sorted(paths), "note": "Integrity only; not a digital signature."}


def canonical_remote(url: str) -> str:
    """Accept only credential-free canonical GitHub HTTPS or git-user SSH URLs."""
    value = url.strip()
    scp = re.fullmatch(r"git@github\.com:([^\s]+)", value, flags=re.IGNORECASE)
    if scp:
        repo = scp.group(1)
    else:
        parts = urllib.parse.urlsplit(value)
        require(parts.hostname and parts.hostname.casefold() == "github.com", "Unexpected origin host.")
        require(not parts.query and not parts.fragment, "Origin URL query/fragment refused.")
        try:
            port = parts.port
        except ValueError as exc:
            raise HandoffError("Invalid origin port.") from exc
        if parts.scheme == "https":
            require(parts.username is None and parts.password is None and port in (None, 443),
                    "Embedded credentials or unexpected HTTPS port refused.")
        elif parts.scheme == "ssh":
            require(parts.username == "git" and parts.password is None and port in (None, 22),
                    "Unexpected SSH account/port.")
        else:
            raise HandoffError("Origin must use canonical GitHub HTTPS or SSH.")
        repo = parts.path.lstrip("/")
    if repo.endswith(".git"):
        repo = repo[:-4]
    require(repo.casefold() == REPOSITORY.casefold(), "Origin does not target PikkuJanne/WinPDFMerger.")
    return REPOSITORY


def git(repo: Path, *args: str) -> str:
    env = os.environ.copy()
    env.update({"GIT_TERMINAL_PROMPT": "0", "GCM_INTERACTIVE": "Never", "GIT_OPTIONAL_LOCKS": "0"})
    try:
        result = subprocess.run(["git", "-C", str(repo), *args], check=False,
                                capture_output=True, text=True, encoding="utf-8", errors="replace",
                                timeout=60, env=env)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise HandoffError(f"Git {args[0] if args else 'command'} could not complete.") from exc
    # Do not echo raw stderr or URLs: remote configuration may contain credentials.
    require(result.returncode == 0, f"Git {args[0] if args else 'command'} failed (exit {result.returncode}); inspect locally.")
    return result.stdout.strip()


def inspect_repo(repo: Path) -> dict[str, Any]:
    reject_reparse_chain(repo)
    repo = repo.absolute()
    require(repo.is_dir(), "Repository directory does not exist.")
    root = Path(git(repo, "rev-parse", "--show-toplevel")).resolve()
    require(root == repo.resolve(), "Use the repository root, not a subdirectory.")
    for args in (("remote", "get-url", "--all", "origin"),
                 ("remote", "get-url", "--push", "--all", "origin")):
        urls = git(repo, *args).splitlines()
        require(len(urls) == 1, "Exactly one origin fetch/push target is required; inspect multiple URLs.")
        canonical_remote(urls[0])
    head = git(repo, "rev-parse", "HEAD")
    require(bool(HEX40.fullmatch(head)), "Expected a committed SHA-1 Git checkout.")
    return {"root": str(root), "head": head,
            "branch": git(repo, "branch", "--show-current"),
            "dirty": bool(git(repo, "status", "--porcelain=v1", "--untracked-files=all"))}


def allowed_payload(rel: str) -> bool:
    return (rel == "AGENTS.md" or rel.startswith("docs/codex/")
            or rel.startswith("tools/codex/"))


def reject_case_alias(root: Path, relative: str) -> None:
    """Prevent case-only path collisions, including on a case-sensitive dev host."""
    parent = root
    for part in PurePosixPath(relative).parts:
        if parent.is_dir():
            matches = [p.name for p in parent.iterdir() if p.name.casefold() == part.casefold()]
            require(not matches or matches == [part], "Case-only destination conflict; inspect before import.")
        parent = parent / part


def import_bundle(bundle: Path, repo: Path, apply: bool = False) -> dict[str, Any]:
    verified = verify_bundle(bundle)
    state = inspect_repo(repo)
    require(not state["dirty"] or not apply, "Apply requires a clean reconciled checkout; no automatic stash/reset.")
    root = Path(state["root"])
    plan: list[dict[str, str]] = []
    folded: set[str] = set()
    for name in verified["paths"]:
        if not name.startswith("payload/"):
            continue
        rel = safe_relative(name[len("payload/"):]).as_posix()
        require(allowed_payload(rel), f"Payload path is outside the handoff allowlist: {rel}")
        require(rel.casefold() not in folded, "Case-colliding payload destination.")
        folded.add(rel.casefold())
        src, dst = bundle / name, root / rel
        reject_case_alias(root, rel)
        reject_reparse_chain(dst)
        if dst.exists():
            require(dst.is_file(), f"Destination is not a regular file: {rel}")
            action = "skip-identical" if digest(src) == digest(dst) else "conflict"
        else:
            action = "create"
        plan.append({"path": rel, "action": action})
    require(bool(plan), "No payload files in bundle.")
    require(not any(x["action"] == "conflict" for x in plan),
            "Different destination files exist: inspect and manually merge; nothing was written.")
    if apply:
        # Recheck sources and dirty state immediately before writes. Every destination
        # was preflighted, and exclusive-create protects existing files from overwrite.
        verify_bundle(bundle)
        require(not inspect_repo(root)["dirty"], "Checkout became dirty before apply.")
        for item in plan:
            if item["action"] != "create":
                continue
            dst = root / item["path"]
            reject_reparse_chain(dst)
            dst.parent.mkdir(parents=True, exist_ok=True)
            reject_reparse_chain(dst)
            with (bundle / "payload" / item["path"]).open("rb") as source:
                with dst.open("xb") as target:
                    for block in iter(lambda: source.read(1024 * 1024), b""):
                        target.write(block)
        # On an unexpected IO error, some create-only files may remain. We never
        # delete them as rollback; inspect/commit safely and rerun identically.
    return {"mode": "applied" if apply else "preview", "repository": REPOSITORY,
            "checkout_dirty_before": state["dirty"], "files": plan,
            "git_metadata_changed_by_helper": False,
            "next": "Review/stage explicit files, commit/push normally, then verify live synchronization."}


def parse_refs(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for line in text.splitlines():
        fields = line.split()
        require(len(fields) == 2 and bool(HEX40.fullmatch(fields[0])), "Malformed live Git ref response.")
        require(fields[1] not in result, "Duplicate live Git ref.")
        result[fields[1]] = fields[0].lower()
    return result


def sync(repo: Path) -> dict[str, Any]:
    state = inspect_repo(repo)
    branch = state["branch"]
    require(bool(branch), "Detached HEAD is not a synchronized working branch.")
    require(git(repo, "config", "--get", f"branch.{branch}.remote") == "origin", "Upstream remote is not origin.")
    ref = f"refs/heads/{branch}"
    require(git(repo, "config", "--get", f"branch.{branch}.merge") == ref, "Upstream branch name differs.")
    refs = parse_refs(git(repo, "ls-remote", "--exit-code", "--heads", "origin", ref))
    require(ref in refs and len(refs) == 1, "Expected live origin branch not found.")
    matched = refs[ref] == state["head"].lower()
    return {"repository": REPOSITORY, "branch": branch, "local_head": state["head"],
            "live_remote_head": refs[ref], "clean": not state["dirty"],
            "synchronized": matched and not state["dirty"], "checked_at": utc_now(),
            "method": "live git ls-remote; no fetch, push or checkout mutation"}


def evidence_exists(repo: Path, values: Any) -> None:
    require(isinstance(values, list) and bool(values), "Completed result needs evidence paths.")
    for item in values:
        rel = safe_relative(item).as_posix()
        require(rel.startswith("docs/codex/evidence/") and rel != "docs/codex/evidence/README.md",
                "Evidence must be an actual task record under docs/codex/evidence/, not a template.")
        p = repo / rel
        reject_reparse_chain(p)
        require(p.is_file() and p.stat().st_size > 0, f"Missing/empty evidence: {rel}")


def check_plan(repo: Path, gate: str | None = None) -> dict[str, Any]:
    docs = repo / "docs/codex"
    task_doc, case_doc = read_json(docs / "TASKS.json"), read_json(docs / "ACCEPTANCE_CASES.json")
    coverage = read_json(docs / "IMPROVEMENT_COVERAGE.json")
    release = read_json(docs / "RELEASE_STATE.json")
    require(task_doc.get("repository") == REPOSITORY and task_doc.get("target_release") == TAG,
            "Plan targets wrong repository or release.")
    tasks, cases = task_doc.get("tasks", []), case_doc.get("cases", [])
    require(isinstance(tasks, list) and isinstance(cases, list), "Tasks/cases must be arrays.")
    tmap, cmap = {x["id"]: x for x in tasks}, {x["id"]: x for x in cases}
    require(len(tmap) == len(tasks) and set(tmap) == {f"T{n:02d}" for n in range(1, 35)}, "Task IDs missing/duplicated.")
    require(len(cmap) == len(cases) and len(cases) >= 78, "Acceptance IDs missing/duplicated.")
    statuses = {"pending", "in_progress", "blocked", "done"}
    results = {"not_run", "pass", "fail", "excluded"}
    require(sum(t["status"] == "in_progress" for t in tasks) <= 1, "More than one active conceptual task.")
    for t in tasks:
        require(t.get("status") in statuses, f"Invalid task status: {t['id']}")
        brief = safe_relative(t["brief"]).as_posix()
        require((repo / brief).is_file(), f"Task brief missing: {t['id']}")
        require(all(d in tmap and d != t["id"] for d in t.get("depends_on", [])), "Unknown/self task dependency.")
        require(bool(t.get("acceptance_ids")) and len(set(t["acceptance_ids"])) == len(t["acceptance_ids"]), "Missing/duplicate task acceptance links.")
        for cid in t["acceptance_ids"]:
            require(cid in cmap and cmap[cid]["task_id"] == t["id"], "Broken task/case linkage.")
        if t["status"] == "done":
            evidence_exists(repo, t.get("evidence"))
            require(all(tmap[d]["status"] == "done" for d in t["depends_on"]), "Done task has unfinished dependency.")
            require(all(cmap[c]["result"] in ("pass", "excluded") for c in t["acceptance_ids"]), "Done task has uncompleted acceptance cases.")
    visited, visiting = set(), set()
    def visit(tid: str) -> None:
        require(tid not in visiting, "Cyclic task dependencies.")
        if tid in visited:
            return
        visiting.add(tid)
        for dep in tmap[tid]["depends_on"]:
            visit(dep)
        visiting.remove(tid); visited.add(tid)
    for tid in tmap:
        visit(tid)
    linked = [c for t in tasks for c in t["acceptance_ids"]]
    require(len(linked) == len(cases) and set(linked) == set(cmap), "Unlinked or multiply linked acceptance case.")
    for c in cases:
        require(c.get("result") in results and isinstance(c.get("required"), bool), "Invalid case result/required flag.")
        require(c.get("stage") in {"pre_release", "accepted", "prepared", "published"}, "Invalid acceptance stage.")
        if c["result"] in ("pass", "excluded"):
            evidence_exists(repo, c.get("evidence"))
        if c["result"] == "excluded":
            require(not c["required"] and bool(c.get("exclusion_reason")), "Required case cannot be excluded; scoped exclusion needs rationale.")
    improvements = coverage.get("improvements", [])
    require({x["id"] for x in improvements} == {f"I{n:02d}" for n in range(1, 18)} and len(improvements) == 17,
            "All 17 improvement mappings are required.")
    for item in improvements:
        require(bool(item.get("task_ids")) and all(x in tmap for x in item["task_ids"]), "Invalid improvement coverage link.")
    if gate:
        limit = {"ready": 30, "prepared": 32, "complete": 34}[gate]
        selected = [t for t in tasks if int(t["id"][1:]) <= limit]
        require(all(t["status"] == "done" for t in selected), f"Gate {gate} has unfinished tasks.")
        for t in selected:
            for cid in t["acceptance_ids"]:
                c = cmap[cid]
                require(c["result"] == "pass" or (not c["required"] and c["result"] == "excluded"),
                        f"Gate {gate} has unverified case {cid}.")
        if gate in {"prepared", "complete"}:
            require(release.get("state") in {"prepared", "published", "verified", "complete"}, "Release state is not prepared.")
            require(bool(HEX40.fullmatch(release.get("release_commit") or "")), "Missing accepted release commit.")
            for key in ("zip_sha256", "checksums_sha256"):
                require(bool(HEX64.fullmatch(release.get(key) or "")), f"Missing accepted {key}.")
            evidence_exists(repo, release.get("readiness_evidence"))
        if gate == "complete":
            require(release.get("state") in {"verified", "complete"}, "Publication is not verified.")
            require(release.get("release_url") == f"https://github.com/{REPOSITORY}/releases/tag/{TAG}", "Wrong/missing published release URL.")
            require(bool(release.get("published_at")), "Missing observed published_at.")
            evidence_exists(repo, release.get("publication_evidence"))
            evidence_exists(repo, release.get("post_publication_smoke_evidence"))
    return {"valid": True, "gate": gate or "structure-only", "task_count": len(tasks),
            "case_count": len(cases), "done_tasks": sum(t["status"] == "done" for t in tasks),
            "passed_cases": sum(c["result"] == "pass" for c in cases),
            "excluded_cases": sum(c["result"] == "excluded" for c in cases),
            "note": "Record validation only; does not execute application tests or verify current GitHub state."}


def public_json(url: str) -> Any:
    require(url.startswith(f"https://api.github.com/repos/{REPOSITORY}/"), "Unexpected public API URL.")
    request = urllib.request.Request(url, headers={"User-Agent": "WinPDFMerger-release-verifier", "Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(request, timeout=60) as response:
        data = response.read(8 * 1024 * 1024 + 1)
    require(len(data) <= 8 * 1024 * 1024, "Oversized API response.")
    return json.loads(data.decode("utf-8"))


def public_pages(endpoint: str) -> list[dict[str, Any]]:
    all_items: list[dict[str, Any]] = []
    for page in range(1, 101):
        batch = public_json(f"https://api.github.com/repos/{REPOSITORY}/{endpoint}?per_page=100&page={page}")
        require(isinstance(batch, list), "Unexpected paginated API data.")
        all_items.extend(batch)
        if len(batch) < 100:
            return all_items
    raise HandoffError("Public API pagination exceeded safety limit; inspect manually.")


def select_public_release(releases: list[dict[str, Any]]) -> dict[str, Any]:
    public = [r for r in releases if not r.get("draft", False)]
    require(len(public) == 1, "Expected exactly one public GitHub Release, including any prereleases.")
    r = public[0]
    require(r.get("tag_name") == TAG and r.get("draft") is False and r.get("prerelease") is False,
            "v1.0.0 is missing, a draft, or a prerelease.")
    require(bool(r.get("published_at")), "Release is not actually published.")
    require(r.get("html_url") == f"https://github.com/{REPOSITORY}/releases/tag/{TAG}", "Unexpected release URL.")
    require(isinstance(r.get("id"), int), "Missing release ID.")
    return r


def download_asset(asset: dict[str, Any], directory: Path) -> Path:
    name = asset.get("name")
    require(name in (ZIP_NAME, SUM_NAME), "Unexpected asset name.")
    expected_url = f"https://github.com/{REPOSITORY}/releases/download/{TAG}/{name}"
    require(asset.get("browser_download_url") == expected_url, "Unexpected public asset URL.")
    size = asset.get("size")
    require(isinstance(size, int) and 0 < size <= MAX_DOWNLOAD, "Invalid/oversized release asset.")
    require(asset.get("state") == "uploaded", "Asset upload is not complete.")
    target = directory / name
    request = urllib.request.Request(expected_url, headers={"User-Agent": "WinPDFMerger-release-verifier"})
    with urllib.request.urlopen(request, timeout=60) as response:
        final = urllib.parse.urlsplit(response.geturl())
        host = (final.hostname or "").lower()
        require(final.scheme == "https" and (host == "github.com" or host.endswith(".githubusercontent.com")),
                "Unexpected asset redirect; inspect before downloading.")
        count = 0
        with target.open("xb") as stream:
            while True:
                block = response.read(1024 * 1024)
                if not block:
                    break
                count += len(block)
                require(count <= size and count <= MAX_DOWNLOAD, "Asset exceeds declared size.")
                stream.write(block)
    require(count == size, "Downloaded size differs from GitHub asset metadata.")
    return target


def verify_package(path: Path, expected_commit: str) -> dict[str, Any]:
    """Inspect ZIP structure and generated provenance without extracting/executing."""
    with zipfile.ZipFile(path) as archive:
        files: dict[str, zipfile.ZipInfo] = {}
        seen: set[str] = set()
        total = 0
        for info in archive.infolist():
            name = info.filename[:-1] if info.is_dir() else info.filename
            rel = safe_relative(name)
            require(rel.parts[0] == PACKAGE_ROOT, "ZIP has the wrong top-level folder.")
            require(name.casefold() not in seen, "Duplicate/case-colliding ZIP entry.")
            seen.add(name.casefold())
            mode = info.external_attr >> 16
            require(not stat.S_ISLNK(mode), "ZIP symlink refused.")
            require(not (info.flag_bits & 1), "Encrypted application package refused.")
            if info.is_dir():
                continue
            require(len(rel.parts) >= 2, "File must be inside the package root.")
            inner = PurePosixPath(*rel.parts[1:]).as_posix()
            require(not inner.startswith(("docs/codex/", "tools/", "tests/", ".github/")), "Developer content in application package.")
            require(not inner.lower().endswith((".pdf", ".log", ".exe", ".py")), "Unexpected PDF/log/vendor executable/Python runtime in package.")
            total += info.file_size
            require(total <= MAX_DOWNLOAD and info.file_size <= 32 * 1024 * 1024, "Unpacked package exceeds safety limits.")
            files[inner] = info
        require({"WinPDFMerge.ps1", "WinPDFMerge.bat", "README.md", "LICENSE", "BUILD_INFO.json"} <= set(files),
                "Required end-user package files are missing.")
        info = json.loads(archive.read(files["BUILD_INFO.json"]).decode("utf-8-sig"))
        require(info.get("version") == "1.0.0" and info.get("source_commit") == expected_commit,
                "BUILD_INFO does not match accepted version/commit.")
        entries = info.get("files")
        require(isinstance(entries, list), "BUILD_INFO lacks per-file inventory.")
        expected_files = set(files) - {"BUILD_INFO.json"}
        inventory_paths: set[str] = set()
        for item in entries:
            rel = safe_relative(item.get("path", "")).as_posix()
            require(rel in expected_files and rel not in inventory_paths, "BUILD_INFO has missing/duplicate/unexpected file.")
            inventory_paths.add(rel)
            expected_hash = item.get("sha256", "")
            require(bool(HEX64.fullmatch(expected_hash)), "Invalid BUILD_INFO file hash.")
            actual = hashlib.sha256(archive.read(files[rel])).hexdigest()
            require(hmac.compare_digest(actual, expected_hash.lower()), "Packaged file hash mismatch.")
        require(inventory_paths == expected_files, "BUILD_INFO inventory does not cover all package files.")
    return {"safe_structure": True, "source_commit": expected_commit, "file_count": len(files),
            "application_executed": False}


def verify_release(repo: Path, expected_commit: str, zip_hash: str, sums_hash: str, directory: Path) -> dict[str, Any]:
    require(bool(HEX40.fullmatch(expected_commit)), "Expected release commit must be a full 40-character SHA.")
    require(bool(HEX64.fullmatch(zip_hash)) and bool(HEX64.fullmatch(sums_hash)), "Both prepublication SHA-256 values are required.")
    state = inspect_repo(repo)
    require(not state["dirty"], "Public-release verification requires a clean local checkout.")
    tagref = f"refs/tags/{TAG}"
    refs = parse_refs(git(repo, "ls-remote", "--exit-code", "origin", tagref, tagref + "^{}"))
    require(tagref in refs and refs.get(tagref + "^{}") == expected_commit.lower(), "Live annotated release tag does not peel to accepted R.")
    release = select_public_release(public_pages("releases"))
    assets = public_pages(f"releases/{release['id']}/assets")
    require(len(assets) == 2 and {a.get("name") for a in assets} == {ZIP_NAME, SUM_NAME}, "Release assets are not exactly the accepted ZIP/checksum pair.")
    reject_reparse_chain(directory)
    directory = directory.absolute()
    require(not directory.resolve().is_relative_to(Path(state["root"]).resolve()), "Use a fresh download directory outside the repository.")
    require(not directory.exists() or (directory.is_dir() and not any(directory.iterdir())), "Download directory must be new or empty.")
    directory.mkdir(parents=True, exist_ok=True)
    paths = {a["name"]: download_asset(a, directory) for a in assets}
    actual_zip, actual_sums = digest(paths[ZIP_NAME]), digest(paths[SUM_NAME])
    require(hmac.compare_digest(actual_zip, zip_hash.lower()), "Published ZIP differs from accepted prepublication bytes.")
    require(hmac.compare_digest(actual_sums, sums_hash.lower()), "Published checksum file differs from accepted bytes.")
    content = paths[SUM_NAME].read_text(encoding="utf-8-sig").strip()
    match = re.fullmatch(r"([0-9a-fA-F]{64})  WinPDFMerger-v1\.0\.0\.zip", content)
    require(bool(match) and match.group(1).lower() == actual_zip, "SHA256SUMS contents do not match the public ZIP.")
    package = verify_package(paths[ZIP_NAME], expected_commit.lower())
    return {"public_release_verified": True, "repository": REPOSITORY, "tag": TAG,
            "release_id": release["id"], "release_url": release["html_url"],
            "published_at": release["published_at"], "checked_at": utc_now(),
            "release_commit": expected_commit.lower(), "zip_sha256": actual_zip,
            "checksums_sha256": actual_sums, "download_directory": str(directory),
            "package": package, "windows_downloaded_package_smoke_test": "STILL REQUIRED",
            "note": "No application was executed; record a real Windows smoke test before completion."}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subs = parser.add_subparsers(dest="command", required=True)
    verify = subs.add_parser("verify-bundle", help="Verify immutable bundle file inventory and SHA-256 hashes.")
    verify.add_argument("--bundle", type=Path, required=True)
    imp = subs.add_parser("import", help="Preview create-only handoff import; --apply is required to write.")
    imp.add_argument("--bundle", type=Path, required=True); imp.add_argument("--repo", type=Path, required=True)
    imp.add_argument("--apply", action="store_true")
    syn = subs.add_parser("sync", help="Read-only comparison of clean local HEAD with live matching origin branch.")
    syn.add_argument("--repo", type=Path, required=True)
    check = subs.add_parser("check-plan", help="Validate task/evidence records, not PDF behavior.")
    check.add_argument("--repo", type=Path, required=True)
    gates = check.add_mutually_exclusive_group()
    for name in ("ready", "prepared", "complete"):
        gates.add_argument(f"--require-{name}", dest="gate", action="store_const", const=name)
    rel = subs.add_parser("verify-release", help="Anonymously verify/download final public release; does not publish or execute it.")
    rel.add_argument("--repo", type=Path, required=True)
    rel.add_argument("--expected-release-commit", required=True)
    rel.add_argument("--expected-zip-sha256", required=True)
    rel.add_argument("--expected-checksums-sha256", required=True)
    rel.add_argument("--download-dir", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        if args.command == "verify-bundle":
            result = verify_bundle(args.bundle)
            result.pop("paths", None)
        elif args.command == "import":
            result = import_bundle(args.bundle, args.repo, args.apply)
        elif args.command == "sync":
            result = sync(args.repo)
        elif args.command == "check-plan":
            result = check_plan(args.repo, args.gate)
        else:
            result = verify_release(args.repo, args.expected_release_commit, args.expected_zip_sha256,
                                    args.expected_checksums_sha256, args.download_dir)
        print(json.dumps(result, indent=2, ensure_ascii=False))
        return 2 if args.command == "sync" and not result["synchronized"] else 0
    except (HandoffError, OSError, ValueError, KeyError, TypeError, zipfile.BadZipFile) as exc:
        # Some HTTP/OS error messages can contain local paths, so do not publish
        # unsanitized helper diagnostics as task evidence without review.
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
