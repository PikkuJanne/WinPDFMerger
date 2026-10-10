"""Independent T33 exact published-download accepted-R ZIP audit; no builder imports and no application execution."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import pathlib
import platform
import posixpath
import re
import stat
import subprocess
import sys
import zipfile


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--repo", required=True)
    p.add_argument("--commit", required=True)
    p.add_argument("--artifacts", required=True)
    p.add_argument("--repeat-artifacts")
    p.add_argument("--report", required=True)
    p.add_argument("--require-clean", action="store_true")
    p.add_argument("--expected-zip-sha256", required=True)
    p.add_argument("--expected-checksums-sha256", required=True)
    a = p.parse_args()
    repo = pathlib.Path(a.repo).resolve()
    artifacts = pathlib.Path(a.artifacts).resolve()
    report_path = pathlib.Path(a.report).resolve()
    review_root = pathlib.Path(__file__).resolve().parent
    if not report_path.is_relative_to(review_root) or report_path.exists():
        raise ValueError("Use a new report under ignored T33-public-download-review; never overwrite an attempt")
    checks: list[dict] = []

    def check(label: str, truth: bool) -> None:
        checks.append({"check": label, "pass": bool(truth)})

    def git(*args: str) -> bytes:
        completed = subprocess.run(["git", "-C", str(repo), *args],
                                   capture_output=True, check=True,
                                   env={**os.environ, "GIT_NO_REPLACE_OBJECTS": "1", "GIT_OPTIONAL_LOCKS": "0", "GIT_TERMINAL_PROMPT": "0"})
        return completed.stdout

    def blob(name: str) -> bytes:
        return git("show", f"{a.commit}:{name}")

    check("requested source is full lowercase SHA", bool(re.fullmatch(r"[0-9a-f]{40}", a.commit)))
    check("source resolves to requested commit", git("rev-parse", f"{a.commit}^{{commit}}").strip().decode() == a.commit)
    check("requested source is exact accepted release R", a.commit == "95e0a19e6cc5fc01cd4bec4ac15f989f9830840a")
    source_tree = git("rev-parse", f"{a.commit}^{{tree}}").strip().decode()
    check("accepted R tree matches frozen lineage", source_tree == "5014f5bdf4f374aee828ced4c39cb93bfeb6465a")
    status = git("status", "--porcelain=v1", "--untracked-files=all").decode("utf-8")
    check("HEAD equals requested source", git("rev-parse", "HEAD").strip().decode() == a.commit)
    if a.require_clean:
        check("actual audit checkout is clean", status == "")
    version = blob("VERSION").decode("utf-8").rstrip("\r\n")
    check("version is one strict semantic core", bool(re.fullmatch(r"[0-9]+\.[0-9]+\.[0-9]+", version)))
    contract = json.loads(blob("docs/codex/PACKAGE_CONTRACT.json"))
    allowlist = json.loads(blob("release-files.json"))
    check("source contract schema exactly numeric one", type(contract.get("schema_version")) is int and contract["schema_version"] == 1)
    check("source allowlist schema exactly numeric one", isinstance(allowlist, dict) and type(allowlist.get("schema_version")) is int and allowlist["schema_version"] == 1)
    # This supports both obvious explicit allowlist layouts without calling the producer.
    names = allowlist if isinstance(allowlist, list) else allowlist["files"]
    check("allowlist contains explicit strings", isinstance(names, list) and all(isinstance(x, str) for x in names))
    check("allowlist has no duplicate or case collision", len(names) == len(set(x.casefold() for x in names)))
    required_sources = {"WinPDFMerge.ps1", "WinPDFMerge.bat", "VERSION", "README.md", "LICENSE", "SECURITY.md", "src/WinPDFMerge.Helpers.ps1"}
    check("allowlist carries documented runtime/license/security layout", required_sources <= set(names))
    reviewed_sources = required_sources | {"CHANGELOG.md"} | {"docs/" + n + ".md" for n in ("USAGE", "TROUBLESHOOTING", "DEPENDENCIES", "COMPATIBILITY", "PDF_LIMITATIONS", "EMAIL_PRESETS", "RELEASE_NOTES_v1.0.0")}
    check("allowlist equals independently reviewed 15-file inventory", set(names) == reviewed_sources)
    check("generated BUILD_INFO is excluded from source allowlist", "BUILD_INFO.json" not in names)
    excluded_prefixes = (".git/", ".github/", "tests/", "tools/", "docs/codex/", "build/", "dist/")
    for name in names:
        check(f"allowlist source path safe: {name}", name == posixpath.normpath(name) and not name.startswith("/") and "\\" not in name and ":" not in name and all(part not in ("", ".", "..") for part in name.split("/")))
        check(f"allowlist source in runtime/user-doc scope: {name}", not name.casefold().startswith(excluded_prefixes) and pathlib.PurePosixPath(name).suffix.casefold() not in (".pdf", ".log", ".exe", ".dll", ".png", ".jpg", ".jpeg", ".ico") and name != "docs/DEVELOPMENT.md")
        row = git("ls-tree", a.commit, "--", name).decode("utf-8").strip()
        check(f"allowlist source tracked regular Git blob: {name}", bool(re.fullmatch(r"100(?:644|755) blob [0-9a-f]{40}\t" + re.escape(name), row)))
    assets = contract["assets"]
    root = contract["zip_root"]
    zip_name = f"WinPDFMerger-v{version}.zip"
    check("contract asset names derive from VERSION", assets == [zip_name, "SHA256SUMS.txt"])
    check("contract ZIP root derives from VERSION", root == f"WinPDFMerger-v{version}")
    check("contract build-info version agrees", contract["build_info_contract"]["version"] == version)
    check("artifact directory has exactly two assets", sorted(x.name for x in artifacts.iterdir()) == sorted(assets))
    zip_bytes = (artifacts / zip_name).read_bytes()
    sums_bytes = (artifacts / "SHA256SUMS.txt").read_bytes()
    zip_hash = sha256(zip_bytes)
    sums_hash = sha256(sums_bytes)
    check("expected ZIP hash is full lowercase SHA256", bool(re.fullmatch(r"[0-9a-f]{64}", a.expected_zip_sha256)))
    check("expected complete checksum-file hash is full lowercase SHA256", bool(re.fullmatch(r"[0-9a-f]{64}", a.expected_checksums_sha256)))
    check("actual final ZIP equals recorded build asset bytes", zip_hash == a.expected_zip_sha256)
    check("actual complete SHA256SUMS equals recorded build asset bytes", sums_hash == a.expected_checksums_sha256)
    check("checksums exact one-line bytes/hash agree", sums_bytes in (f"{zip_hash}  {zip_name}\n".encode(), f"{zip_hash}  {zip_name}\r\n".encode()))
    source_blobs = {name: blob(name) for name in names}
    expected = {root + "/" + name for name in names} | {root + "/BUILD_INFO.json"}
    zip_rows = []
    packaged = {}
    with zipfile.ZipFile(artifacts / zip_name) as z:
        infos = z.infolist()
        actual = [i.filename for i in infos]
        check("ZIP inventory equals exact allowlist plus BUILD_INFO", set(actual) == expected and len(actual) == len(expected))
        check("ZIP has no duplicated case-insensitive names", len(actual) == len(set(x.casefold() for x in actual)))
        check("ZIP has stable ordinal entry order", actual == sorted(actual))
        check("ZIP container comment empty", z.comment == b"")
        check("ZIP CRC validation succeeds", z.testzip() is None)
        for i in infos:
            name = i.filename
            mode = i.external_attr >> 16
            rel = name[len(root) + 1:] if name.startswith(root + "/") else name
            data = z.read(i)
            packaged[rel] = data
            check(f"ZIP normalized safe path: {rel}", name == posixpath.normpath(name) and name.startswith(root + "/") and "\\" not in name and ":" not in name and all(part not in ("", ".", "..") for part in name.split("/")))
            check(f"ZIP regular nonencrypted file: {rel}", not i.is_dir() and not stat.S_ISLNK(mode) and stat.S_IFMT(mode) in (0, stat.S_IFREG) and not i.flag_bits & 1)
            check(f"ZIP metadata excludes comments/extra data: {rel}", i.comment == b"" and i.extra == b"")
            check(f"ZIP fixed source-independent timestamp: {rel}", i.date_time == (2000, 1, 1, 0, 0, 0))
            check(f"ZIP fixed zero external attributes: {rel}", i.external_attr == 0)
            check(f"ZIP supported NoCompression container representation: {rel}", i.compress_type in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED))
            if rel in source_blobs:
                check(f"ZIP exact Git blob bytes: {rel}", data == source_blobs[rel])
            zip_rows.append({"path": rel, "bytes": len(data), "sha256": sha256(data), "date_time": i.date_time, "create_system": i.create_system, "mode_octal": oct(mode), "compression": i.compress_type})
    bi = json.loads(packaged["BUILD_INFO.json"])
    check("BUILD_INFO UTF-8 JSON object", isinstance(bi, dict))
    check("BUILD_INFO schema exactly numeric one", type(bi.get("schema_version")) is int and bi["schema_version"] == 1)
    check("BUILD_INFO has exactly reviewed provenance fields", set(bi) == {"schema_version", "version", "source_commit", "build_environment", "files"})
    check("BUILD_INFO version agrees", bi.get("version") == version)
    check("BUILD_INFO source exact full source commit", bi.get("source_commit") == a.commit)
    check("BUILD_INFO environment recorded", isinstance(bi.get("build_environment"), dict) and bool(bi["build_environment"]))
    build_environment = bi.get("build_environment", {})
    check("BUILD_INFO recorded builder identity matches tracked producer", build_environment.get("builder") == "tools/release/Build-Release.ps1")
    check("BUILD_INFO producer hash equals source Git blob", build_environment.get("builder_sha256") == sha256(blob("tools/release/Build-Release.ps1")))
    check("BUILD_INFO allowlist hash equals source Git blob", build_environment.get("allowlist_sha256") == sha256(blob("release-files.json")))
    check("BUILD_INFO contract hash equals source Git blob", build_environment.get("package_contract_sha256") == sha256(blob("docs/codex/PACKAGE_CONTRACT.json")))
    for field in ("powershell_version", "powershell_edition", "dotnet_version", "os_version", "process_architecture", "git_version", "zip_format"):
        check(f"BUILD_INFO recorded tool/environment identity: {field}", isinstance(build_environment.get(field), str) and bool(build_environment[field]))
    environment_text = json.dumps(bi.get("build_environment"), ensure_ascii=False)
    check("BUILD_INFO environment no obvious private filesystem paths", not re.search(r"(?i)([a-z]:[\\/]|[/\\]users[/\\]|[/\\]home[/\\])", environment_text))
    inventory = bi.get("files")
    check("BUILD_INFO full inventory is array", isinstance(inventory, list))
    expected_inventory = [{"path": n, "sha256": sha256(source_blobs[n])} for n in sorted(names)]
    check("BUILD_INFO exact per-file inventory/hash agreement", inventory == expected_inventory)
    check("BUILD_INFO does not inventory itself", all(x.get("path") != "BUILD_INFO.json" for x in inventory))
    for name, data in packaged.items():
        if not name.endswith(".md"):
            continue
        for link in re.finditer(r"\[[^\]]*\]\(([^)]+)\)", data.decode("utf-8")):
            target = link.group(1).strip()
            if re.match(r"^[A-Za-z][A-Za-z0-9+.-]*:", target) or target.startswith("#"):
                continue
            path = target.split("#", 1)[0].split("?", 1)[0]
            dest = posixpath.normpath(posixpath.join(posixpath.dirname(name), path))
            check(f"packaged documentation local link exists: {name} -> {path}", dest in packaged)
    repeat = None
    if a.repeat_artifacts:
        repeat_dir = pathlib.Path(a.repeat_artifacts).resolve()
        check("repeat artifact directory exactly two assets", sorted(x.name for x in repeat_dir.iterdir()) == sorted(assets))
        repeat = {n: sha256((repeat_dir / n).read_bytes()) for n in assets}
        check("repeat exact ZIP bytes deterministic", (repeat_dir / zip_name).read_bytes() == zip_bytes)
        check("repeat exact checksum bytes deterministic", (repeat_dir / "SHA256SUMS.txt").read_bytes() == sums_bytes)
    check("actual ZIP bytes unchanged throughout independent audit", sha256((artifacts / zip_name).read_bytes()) == zip_hash)
    check("actual checksum bytes unchanged throughout independent audit", sha256((artifacts / "SHA256SUMS.txt").read_bytes()) == sums_hash)
    check("source HEAD/status unchanged throughout independent audit", git("rev-parse", "HEAD").strip().decode() == a.commit and git("status", "--porcelain=v1", "--untracked-files=all").decode("utf-8") == status)
    issues = [x["check"] for x in checks if not x["pass"]]
    report = {
        "schema_version": 1, "evidence_class": "independent-package-byte-and-source-audit",
        "application_executed": False, "native_engines_executed": False,
        "task": "T33", "result": "pass_for_exact_published_download_package_bytes" if not issues else "fail_for_exact_published_download_package_bytes",
        "source_commit": a.commit, "source_tree": source_tree, "clean_audit_checkout": status == "",
        "recorded_expected_zip_sha256": a.expected_zip_sha256, "recorded_expected_checksums_sha256": a.expected_checksums_sha256,
        "environment": {"system": platform.system(), "release": platform.release(), "machine": platform.machine(), "python": platform.python_version(), "implementation": sys.implementation.name},
        "auditor_sha256": sha256(pathlib.Path(__file__).read_bytes()),
        "producer_source_sha256": sha256(blob("tools/release/Build-Release.ps1")),
        "allowlist_sha256": sha256(blob("release-files.json")),
        "contract_sha256": sha256(blob("docs/codex/PACKAGE_CONTRACT.json")),
        "assets": [{"name": zip_name, "bytes": len(zip_bytes), "sha256": zip_hash}, {"name": "SHA256SUMS.txt", "bytes": len(sums_bytes), "sha256": sums_hash}],
        "repeat_hashes": repeat, "build_info": bi, "zip_files": zip_rows,
        "checks": checks, "checks_total": len(checks), "issues": issues,
        "limitations": ["Review checks are not application/native/Windows-shell acceptance cases.", "No package application or native PDF engine executed by this auditor.", "No publication action or network download performed by this package auditor.", "No source commit is accepted final release R merely because it has a package."]}
    report_path.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"checks": len(checks), "issues": issues, "assets": report["assets"]}))
    return 1 if issues else 0


if __name__ == "__main__":
    raise SystemExit(main())
