"""Operate an exact candidate ZIP on Windows; development-only, no installation.

All application code comes from a hash-verified ZIP. Python, fixture generation
and independent PDF readers belong only to this test driver. Each scenario uses
a fresh extraction and an unrelated working directory. The GS failure scenario
uses a disclosed child-only resource configuration fault in the genuine engine.
"""
from __future__ import annotations

import argparse
from contextlib import closing
import ctypes
from ctypes import wintypes
import datetime as dt
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import random
import re
import stat
import subprocess
import sys
import time
import zipfile

PINS = {"python": "3.12.14", "reportlab": "4.4.9", "pypdf": "6.10.0",
        "pillow": "12.3.0", "pypdfium2": "5.13.0", "pdfium": "153.0.7999.0"}
PACKAGE_ROOT = "WinPDFMerger-v1.0.0"
PAGE_ID = re.compile(r"T03-[0-9]{2}-P[0-9]{2}")
BAD_GS_INIT = b"/T29FaultToken load\n"


def digest(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def file_hash(path: Path) -> str:
    return digest(path.read_bytes())


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def safe_entry(name: str) -> str:
    if not isinstance(name, str) or not re.fullmatch(r"[A-Za-z0-9_.-]+(?:/[A-Za-z0-9_.-]+)*", name):
        raise ValueError("Unsafe ZIP entry name")
    for part in name.split("/"):
        if part in (".", "..") or part.endswith(".") or re.match(r"(?i)^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)", part):
            raise ValueError("Unsafe or reserved ZIP path")
    return name


def ordinary_path(path: Path) -> None:
    """Reject links/junctions along all existing ancestors, without resolving them."""
    for candidate in (path, *path.parents):
        try:
            info = candidate.stat(follow_symlinks=False)
        except FileNotFoundError:
            continue
        require(not stat.S_ISLNK(info.st_mode) and not getattr(info, "st_file_attributes", 0) & 0x400,
                "Reparse/link test paths are refused")


def snapshot(paths: list[Path], root: Path) -> list[dict]:
    rows = []
    for path in sorted(paths, key=lambda p: str(p).casefold()):
        ordinary_path(path)
        info = path.stat()
        require(path.is_file(), "Guarded file is missing or nonregular")
        rows.append({"path": path.relative_to(root).as_posix(), "sha256": file_hash(path),
                     "bytes": info.st_size, "modified_ns": info.st_mtime_ns,
                     "attributes": getattr(info, "st_file_attributes", 0)})
    return rows


def environment_hash() -> str:
    return digest(json.dumps(dict(os.environ), sort_keys=True, ensure_ascii=True).encode())


def tree_inventory(root: Path) -> dict:
    files, directories = [], []
    for path in root.rglob("*"):
        ordinary_path(path)
        (directories if path.is_dir() else files).append(path.relative_to(root).as_posix())
    return {"files": sorted(files), "directories": sorted(directories)}


def require_inventory(actual: dict, expected: dict) -> None:
    require(actual == expected, "Complete recursive file/directory inventory changed")


def require_email_result(log: str, expected: str) -> None:
    require(re.findall(r"(?m)^Email result: ([a-z_]+)\r?$", log) == [expected], "Packaged email result differs")


class OwnedJob:
    """A Windows job owns each test child and descendants until capture completes."""
    def __init__(self, process):
        class Basic(ctypes.Structure):
            _fields_ = [("process_time", ctypes.c_longlong), ("job_time", ctypes.c_longlong),
                        ("flags", wintypes.DWORD), ("min_ws", ctypes.c_size_t), ("max_ws", ctypes.c_size_t),
                        ("active", wintypes.DWORD), ("affinity", ctypes.c_size_t),
                        ("priority", wintypes.DWORD), ("scheduling", wintypes.DWORD)]
        class Limits(ctypes.Structure):
            _fields_ = [("basic", Basic), ("io", ctypes.c_ulonglong * 6),
                        ("process_memory", ctypes.c_size_t), ("job_memory", ctypes.c_size_t),
                        ("peak_process", ctypes.c_size_t), ("peak_job", ctypes.c_size_t)]
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.CreateJobObjectW.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR]
        kernel.CreateJobObjectW.restype = wintypes.HANDLE
        kernel.SetInformationJobObject.argtypes = [wintypes.HANDLE, ctypes.c_int, ctypes.c_void_p, wintypes.DWORD]
        kernel.AssignProcessToJobObject.argtypes = [wintypes.HANDLE, wintypes.HANDLE]
        kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        self.kernel = kernel
        self.handle = kernel.CreateJobObjectW(None, None)
        if not self.handle:
            process.kill()
            process.wait(timeout=5)
            raise RuntimeError("Could not create owned child job")
        limits = Limits()
        limits.basic.flags = 0x2000  # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE.
        if not kernel.SetInformationJobObject(self.handle, 9, ctypes.byref(limits), ctypes.sizeof(limits)) or not kernel.AssignProcessToJobObject(self.handle, wintypes.HANDLE(int(process._handle))):
            self.close()
            process.kill()
            process.wait(timeout=5)
            raise RuntimeError("Could not assign test child to its owned job")

    def close(self):
        if self.handle:
            self.kernel.CloseHandle(self.handle)
            self.handle = None


class Capture:
    def __init__(self, root: Path):
        self.root = root
        self.calls = []

    def invoke(self, label: str, arguments, cwd: Path, overrides=None, stdin=None, timeout=60):
        overrides = overrides or {}
        env = dict(os.environ)
        for name, value in overrides.items():
            # Environment names on Windows are case-insensitive.
            for existing in list(env):
                if existing.casefold() == name.casefold():
                    del env[existing]
            if value is not None:
                env[name] = str(value)
        stem = f"{len(self.calls):03d}-{label}"
        require(re.fullmatch(r"[A-Za-z0-9_.-]+", stem) is not None, "Unsafe capture label")
        started = dt.datetime.now(dt.timezone.utc).isoformat()
        watch = time.monotonic()
        timed_out = False
        process = subprocess.Popen(arguments, cwd=cwd, env=env, shell=False,
                                   stdin=subprocess.PIPE if stdin is not None else subprocess.DEVNULL,
                                   stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                   creationflags=subprocess.CREATE_NO_WINDOW)
        job = OwnedJob(process)
        try:
            try:
                stdout, stderr = process.communicate(stdin, timeout=timeout)
            except subprocess.TimeoutExpired:
                timed_out = True
                job.close()
                stdout, stderr = process.communicate(timeout=5)
        finally:
            job.close()
        out = self.root / (stem + ".stdout.bin")
        err = self.root / (stem + ".stderr.bin")
        out.write_bytes(stdout)
        err.write_bytes(stderr)
        row = {"label": label, "arguments": arguments, "working_directory": str(cwd),
               "environment_overrides": overrides, "stdin_sha256": digest(stdin) if stdin is not None else None,
               "started_at_utc": started, "elapsed_seconds": round(time.monotonic() - watch, 3),
               "timeout_seconds": timeout, "timed_out": timed_out, "exit_code": process.returncode,
               "owned_job": True, "stdout": out.name, "stdout_sha256": digest(stdout),
               "stderr": err.name, "stderr_sha256": digest(stderr)}
        self.calls.append(row)
        (self.root / (stem + ".invocation.json")).write_text(json.dumps(row, indent=2) + "\n", encoding="utf-8")
        (self.root / "invocations.json").write_text(json.dumps(self.calls, indent=2) + "\n", encoding="utf-8")
        require(not timed_out, "Owned child timed out: " + label)
        return row, stdout.decode("utf-8-sig", errors="replace"), stderr.decode("utf-8-sig", errors="replace")


def load_versions() -> dict:
    import reportlab
    import pypdf
    import pypdfium2
    import PIL
    observed = {"python": sys.version.split()[0], "reportlab": reportlab.Version,
                "pypdf": pypdf.__version__, "pillow": PIL.__version__,
                "pypdfium2": str(pypdfium2.PYPDFIUM_INFO), "pdfium": str(pypdfium2.PDFIUM_INFO)}
    require(observed == PINS, "Approved exact independent-reader/development pins required")
    return observed


def inspect_pdf(path: Path, identifiers: list[str], render: Path) -> dict:
    import pypdfium2 as pdfium
    from pypdf import PdfReader
    strict = PdfReader(path, strict=True)
    require(len(strict.pages) == len(identifiers), "Independent strict PDF page count differs")
    pages = []
    render.mkdir(exist_ok=False)
    with pdfium.PdfDocument(path) as document:
        require(len(document) == len(identifiers), "Independent PDFium page count differs")
        for index, expected in enumerate(identifiers):
            with closing(document[index]) as page:
                with closing(page.get_textpage()) as text:
                    actual = PAGE_ID.findall(text.get_text_range())
                require(actual == [expected], "Independent page identifier/order differs")
                require(list(page.get_size()) == [432.0, 288.0] and page.get_rotation() == 0, "Independent page dimensions/rotation differ")
                png = render / f"page-{index + 1:02d}.png"
                with closing(page.render(scale=1.0)) as bitmap:
                    image = bitmap.to_pil()
                    require(image.convert("L").getextrema()[0] < 230, "Independent PDF rendering is blank")
                    image.save(png)
                pages.append({"identifier": expected, "size_points": list(page.get_size()), "rotation": page.get_rotation(),
                              "render_path": str(png), "render_sha256": file_hash(png)})
    return {"path": str(path), "sha256": file_hash(path), "bytes": path.stat().st_size,
            "strict_pypdf_pages": len(strict.pages), "pdfium_pages": len(pages), "pages": pages}


def raster_pdf() -> bytes:
    """Original deterministic synthetic RGB/text fixture; no private inputs."""
    pixels = random.Random(140032).randbytes(1200 * 800 * 3)
    content = b"q 432 0 0 260 0 28 cm /Im0 Do Q\nBT /F1 12 Tf 24 8 Td (T03-14-P01) Tj ET\n"
    objects = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
               b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 432 288] /Resources << /Font << /F1 4 0 R >> /XObject << /Im0 5 0 R >> >> /Contents 6 0 R >>",
               b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
               b"<< /Type /XObject /Subtype /Image /Width 1200 /Height 800 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length " + str(len(pixels)).encode() + b" >>\nstream\n" + pixels + b"\nendstream",
               b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"endstream"]
    output = io.BytesIO()
    output.write(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets = []
    for number, obj in enumerate(objects, 1):
        offsets.append(output.tell())
        output.write(str(number).encode() + b" 0 obj\n" + obj + b"\nendobj\n")
    xref = output.tell()
    output.write(b"xref\n0 7\n0000000000 65535 f \n")
    for offset in offsets:
        output.write(f"{offset:010} 00000 n \n".encode())
    output.write(b"trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n" + str(xref).encode() + b"\n%%EOF\n")
    return output.getvalue()


def validate_package(raw: bytes, expected_hash: str, checksums: bytes, checksums_hash: str,
                     source: str, payload: dict[str, bytes]) -> tuple[dict, dict[str, bytes]]:
    require(digest(raw) == expected_hash and digest(checksums) == checksums_hash, "Exact independently accepted asset hashes differ")
    require(checksums in [(expected_hash + "  " + PACKAGE_ROOT + ".zip" + ending).encode() for ending in ("\n", "\r\n")], "Exact checksum manifest content differs")
    contents = {}
    with zipfile.ZipFile(io.BytesIO(raw)) as archive:
        for entry in archive.infolist():
            name = safe_entry(entry.filename)
            require(name.startswith(PACKAGE_ROOT + "/"), "Unexpected package root")
            relative = name[len(PACKAGE_ROOT) + 1:]
            require(relative.casefold() not in {key.casefold() for key in contents}, "Duplicate package entry")
            mode = entry.external_attr >> 16
            require(not entry.is_dir() and not stat.S_ISLNK(mode) and entry.external_attr == 0, "Nonregular package entry")
            contents[relative] = archive.read(entry)
    require(set(contents) == set(payload) | {"BUILD_INFO.json"}, "Package inventory differs from reviewed tracked allowlist")
    require(all(contents[path] == raw_file for path, raw_file in payload.items()), "Packaged bytes differ from exact source Git blobs")
    info = json.loads(contents["BUILD_INFO.json"])
    require(info["version"] == "1.0.0" and info["source_commit"] == source, "Candidate provenance differs")
    require(isinstance(info["build_environment"], dict), "Missing actual builder environment")
    inventory = info["files"]
    require(isinstance(inventory, list) and len(inventory) == len(payload), "BUILD_INFO inventory count differs")
    require(len({row["path"].casefold() for row in inventory}) == len(payload), "Duplicate BUILD_INFO inventory")
    require({row["path"]: row["sha256"] for row in inventory} == {path: digest(raw_file) for path, raw_file in payload.items()}, "BUILD_INFO payload hashes differ")
    return info, contents


def extract(contents: dict[str, bytes], destination: Path) -> Path:
    ordinary_path(destination)
    destination.mkdir(exist_ok=False)
    app = destination / PACKAGE_ROOT
    app.mkdir()
    for relative, raw in contents.items():
        safe_entry(relative)
        path = app / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("xb") as stream:
            stream.write(raw)
    require({p.relative_to(app).as_posix() for p in app.rglob("*") if p.is_file()} == set(contents), "Fresh extraction inventory differs")
    return app


def run(arguments) -> dict:
    require(os.name == "nt", "Actual Windows required; unsupported environments are failures")
    repo = Path(os.path.abspath(arguments.repo))
    ordinary_path(repo)
    ordinary_path(arguments.work_root.absolute())
    ordinary_path(arguments.capture_root.absolute())
    work = Path(os.path.abspath(arguments.work_root))
    capture_root = Path(os.path.abspath(arguments.capture_root))
    require(" " in str(work) and repo not in work.parents and work != repo and work not in repo.parents, "Fresh external working path with spaces required")
    require((repo / "tests/.work") in capture_root.parents, "CaptureRoot must be an owned ignored tests/.work directory")
    require(not work.exists() and not capture_root.exists(), "Fresh work/capture directories required; nothing is overwritten")
    work.mkdir(exist_ok=False)
    capture_root.mkdir(parents=True, exist_ok=False)
    capture = Capture(capture_root)
    result = {"task": "T29", "evidence_class": "actual_candidate_package_operation", "result": "fail",
              "preparation": arguments.preparation, "harness_commit": arguments.expected_harness_commit,
              "candidate_source_commit": arguments.candidate_source_commit, "shell_kind": arguments.shell_kind,
              "work_root": str(work), "capture_root": str(capture_root), "cases": [], "manual_acceptance": "excluded/unperformed; never pass"}
    parent_env = environment_hash()
    driver_before = file_hash(Path(__file__))
    initial_status = None
    cache = []
    candidate_before = {}
    base_env = {"PSModulePath": None}

    def git(*args, binary=False):
        row, stdout, _ = capture.invoke("git-" + args[0], ["git", "-C", str(repo), *args], repo,
                                       {"GIT_NO_REPLACE_OBJECTS": "1", "GIT_OPTIONAL_LOCKS": "0"}, timeout=30)
        require(row["exit_code"] == 0, "Checked Git operation failed")
        return (capture_root / row["stdout"]).read_bytes() if binary else stdout.strip()

    try:
        require(re.fullmatch(r"[a-f0-9]{40}", arguments.expected_harness_commit) is not None and
                re.fullmatch(r"[a-f0-9]{40}", arguments.candidate_source_commit) is not None, "Full lowercase commits required")
        require(git("rev-parse", "HEAD") == arguments.expected_harness_commit, "Expected harness HEAD differs")
        initial_status = git("status", "--porcelain=v1", "--untracked-files=all")
        require(arguments.preparation or not initial_status, "Accepted operation requires clean harness HEAD")
        require(Path(git("rev-parse", "--show-toplevel")) == repo, "Actual repository root required")
        result["independent_reader_versions"] = load_versions()
        pins = (repo / "tests/TestDependencies.psd1").read_text(encoding="utf-8-sig")
        require(file_hash(Path(sys.executable)) in re.findall(r"[0-9a-f]{64}", pins), "Approved pinned developer Python executable required")
        context = json.loads(arguments.approved_cache_manifest.read_text(encoding="utf-8-sig"))
        for expected in context["approved_selected_files"]:
            path = Path(expected["path"].replace("<USERPROFILE>", os.environ["USERPROFILE"]))
            ordinary_path(path)
            require(file_hash(path) == expected["sha256"], "Approved dependency payload hash differs")
            cache.append({"path": str(path), "sha256": expected["sha256"]})
        require(len(cache) == 348, "Expected complete 348-file approved dependency inventory")
        approved = {row["path"].casefold(): row["sha256"] for row in cache}
        for tool in (arguments.pdftk, arguments.ghostscript):
            require(str(tool.absolute()).casefold() in approved, "Selected native tool is outside approved inventory")
        if arguments.shell_kind == "PS7":
            require(str(arguments.shell.absolute()).casefold() in approved, "Selected PS7 is outside approved inventory")
        else:
            require(arguments.shell.absolute() == Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe", "Batch/default Windows PS5.1 path required")
        result["approved_cache_files"] = cache
        result["python"] = {"path": sys.executable, "sha256": file_hash(Path(sys.executable))}
        result["driver_sha256_before"] = driver_before
        version_row, version_text, _ = capture.invoke("shell-inventory", [str(arguments.shell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-Command",
            "$ErrorActionPreference='Stop'; $w=Get-ItemProperty -LiteralPath 'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion'; $p=New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent()); [ordered]@{shell_version=$PSVersionTable.PSVersion.ToString();shell_edition=$PSVersionTable.PSEdition;process_64_bit=[Environment]::Is64BitProcess;os_version=[Environment]::OSVersion.Version.ToString();edition=$w.EditionID;display_version=$w.DisplayVersion;full_build=($w.CurrentBuildNumber+'.'+$w.UBR);is_administrator=$p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);policy=@(Get-ExecutionPolicy -List|ForEach-Object{[ordered]@{scope=$_.Scope.ToString();policy=$_.ExecutionPolicy.ToString()}})}|ConvertTo-Json -Depth 5"], work, base_env)
        require(version_row["exit_code"] == 0, "Actual shell inventory failed")
        result["environment"] = json.loads(version_text)
        require(result["environment"]["process_64_bit"], "Actual x64 host required")
        shell_version = result["environment"]["shell_version"]
        require(shell_version == "7.6.6" if arguments.shell_kind == "PS7" else shell_version.startswith("5.1."), "Required actual shell version differs")
        for label, tool, version in (("pdftk", arguments.pdftk, "2.02"), ("ghostscript", arguments.ghostscript, "10.08.0")):
            row, stdout, _ = capture.invoke(label + "-version", [str(tool), "--version"], work,
                                           {"GS_OPTIONS": None, "GS_LIB": None, "GS_DLL": None})
            require(row["exit_code"] == 0 and re.search(r"\b" + re.escape(version) + r"\b", stdout), "Approved actual native version differs")
        source = arguments.candidate_source_commit
        require(arguments.zip.name == PACKAGE_ROOT + ".zip" and arguments.checksums.name == "SHA256SUMS.txt", "Expected exact asset file names required")
        allowlist = json.loads(git("cat-file", "blob", source + ":release-files.json", binary=True))["files"]
        require(len(allowlist) == 15, "Reviewed candidate payload inventory differs")
        payload = {path: git("cat-file", "blob", source + ":" + safe_entry(path), binary=True) for path in allowlist}
        for path, raw in payload.items():
            require(git("cat-file", "blob", arguments.expected_harness_commit + ":" + path, binary=True) == raw,
                    "Current source runtime/public payload differs from candidate")
        candidate_before = {"zip": file_hash(arguments.zip), "checksums": file_hash(arguments.checksums)}
        info, contents = validate_package(arguments.zip.read_bytes(), arguments.zip_sha256,
                                         arguments.checksums.read_bytes(), arguments.checksums_sha256, source, payload)
        result["candidate"] = {"zip_path": str(arguments.zip), "zip_sha256": arguments.zip_sha256,
                               "checksums_path": str(arguments.checksums), "checksums_sha256": arguments.checksums_sha256,
                               "build_info": info, "payload_count": len(payload), "package_count": len(contents)}
        generator = repo / "tools/test/generate_numbered_fixtures.py"
        spec = importlib.util.spec_from_file_location("t29_original_numbered", generator)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        original = module.expected_files()
        require(all((repo / "tests/fixtures/numbered" / name).read_bytes() == raw for name, raw in original.items()), "Original fixtures do not reproduce")
        normal = {"1.pdf": original["1.pdf"], "01.pdf": module.make_pdf("01.pdf", ("T03-01-P02",)),
                  "2.PDF": original["2.pdf"], "10.pdf": original["10.pdf"], "20.pdf": raster_pdf()}
        identifiers = ["T03-01-P01", "T03-01-P02", "T03-02-P01", "T03-02-P02", "T03-10-P01", "T03-14-P01"]
        result["fixtures"] = {"generator": str(generator), "generator_sha256": file_hash(generator),
                              "provenance": "Original pinned numbered recipe, additional original 01 page, deterministic stdlib RGB raster; no private PDFs",
                              "files": [{"path": key, "sha256": digest(raw), "bytes": len(raw)} for key, raw in normal.items()],
                              "expected_order": list(normal), "expected_identifiers": identifiers}
        cases = [("default-screen", "normal", 0, "published", [], False, True),
                 ("ebook-output", "normal", 0, "published", ["-EmailPreset", "ebook"], False, False),
                 ("skip-email", "normal", 0, "skipped", ["-SkipEmail"], False, False),
                 ("skip-ignored-preset", "normal", 0, "skipped", ["-SkipEmail", "-EmailPreset", "ebook"], False, False),
                 ("tiny-no-benefit", "tiny", 0, "no_size_benefit", [], False, False),
                 ("optional-gs-absent", "tiny", 0, "unavailable", [], False, False),
                 ("missing-input", "empty", 1, None, [], False, True),
                 ("invalid-preset", "normal", 1, None, ["-EmailPreset", "invalid"], False, False),
                 ("empty-input", "empty", 1, None, [], False, False),
                 ("corrupt-input", "corrupt", 1, None, [], False, False),
                 ("gs-resource-failure", "normal", 2, "failed", [], False, False)]
        if arguments.shell_kind == "PS51":
            cases += [("batch-default", "normal", 0, "published", [], True, True),
                      ("batch-empty", "empty", 1, None, [], True, True),
                      ("batch-gs-failure", "normal", 2, "failed", [], True, True)]
        for label, fixture, code, email_state, options, batch, default_output in cases:
            case = work / ("case " + label)
            case.mkdir()
            app = extract(contents, case / "fresh install")
            source_dir, output, cwd, no_common, temp = [case / name for name in ("synthetic inputs", "separate output", "unrelated cwd", "no common engines", "child temp")]
            for directory in (source_dir, output, cwd, no_common, temp):
                directory.mkdir()
            selected = normal if fixture == "normal" else ({"1.pdf": original["1.pdf"]} if fixture == "tiny" else ({"1.pdf": b"%PDF-1.4\nT29 original corrupt fixture\n%%EOF\n"} if fixture == "corrupt" else {}))
            for name, raw in selected.items():
                (source_dir / name).write_bytes(raw)
            (source_dir / "not a PDF.txt").write_bytes(b"T29 source foreign canary\n")
            (source_dir / "ignored subfolder").mkdir()
            (source_dir / "ignored subfolder/99.pdf").write_bytes(original["1.pdf"])
            hidden = source_dir / "hidden.pdf"
            hidden.write_bytes(original["1.pdf"])
            require(ctypes.windll.kernel32.SetFileAttributesW(str(hidden), 2), "Could not set owned hidden fixture attribute")
            destination = app if default_output else output
            foreign = destination / "WinPDFMerge_existing_20000101_00000000.pdf"
            foreign.write_bytes(original["1.pdf"])
            cwd_canary = cwd / "foreign cwd.txt"
            cwd_canary.write_bytes(b"T29 foreign CWD canary\n")
            guarded = [app / path for path in contents] + [foreign, cwd_canary]
            guarded += [p for p in source_dir.rglob("*") if p.is_file()]
            fault = None
            if email_state == "failed":
                resources = case / "owned bad GS resources"
                resources.mkdir()
                fault_path = resources / "gs_init.ps"
                fault_path.write_bytes(BAD_GS_INIT)
                guarded.append(fault_path)
                fault = {"path": str(fault_path), "sha256": file_hash(fault_path), "bytes": len(BAD_GS_INIT),
                         "ascii_content": BAD_GS_INIT.decode("ascii"), "scope": "Owned child-only GS_LIB initialization configuration fault; no package/engine/PDF modification"}
            before = snapshot(guarded, case)
            source_before = snapshot([p for p in source_dir.rglob("*") if p.is_file()], case)
            source_inventory_before = tree_inventory(source_dir)
            cwd_before = tree_inventory(cwd)
            app_inventory_before = tree_inventory(app)
            before_outputs = {p.name for p in destination.iterdir()}
            child_path = [str(arguments.pdftk.parent), str(Path(os.environ["SystemRoot"]) / "System32"),
                          str(Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0")]
            if email_state != "unavailable":
                child_path.insert(0, str(arguments.ghostscript.parent))
            overrides = {"PATH": ";".join(child_path), "ProgramFiles": str(no_common), "ProgramFiles(x86)": str(no_common),
                         "PSModulePath": None, "GS_OPTIONS": "-T29-invalid-inherited-option", "GS_LIB": None, "GS_DLL": None,
                         "TEMP": str(temp), "TMP": str(temp)}
            if email_state == "failed":
                overrides["GS_LIB"] = str(resources)
            if batch:
                cmd = Path(os.environ["SystemRoot"]) / "System32/cmd.exe"
                require(not any(c in str(app) + str(source_dir) for c in '%"\r\n'), "Synthetic batch paths cannot use cmd expansion characters")
                command = f'"{cmd}" /d /v:off /s /c ""{app / "WinPDFMerge.bat"}" "{source_dir}""'
                row, stdout, stderr = capture.invoke(label, command, cwd, overrides, b"T29 automated pause input\r\n")
            else:
                command = [str(arguments.shell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-File", str(app / "WinPDFMerge.ps1")]
                if label != "missing-input":
                    command += [str(source_dir)]
                if not default_output:
                    command += ["-OutputFolder", str(destination)]
                command += options
                row, stdout, stderr = capture.invoke(label, command, cwd, overrides)
            require(row["exit_code"] == code, "Packaged exit code differs: " + label)
            require(snapshot(guarded, case) == before, "Source/package/foreign file snapshot changed: " + label)
            require(snapshot([p for p in source_dir.rglob("*") if p.is_file()], case) == source_before, "Source file set changed: " + label)
            require_inventory(tree_inventory(source_dir), source_inventory_before)
            require_inventory(tree_inventory(cwd), cwd_before)
            new = [p for p in destination.iterdir() if p.name not in before_outputs]
            require(not any(p.name.startswith(".WinPDFMerge") for p in new), "Owned application staging survives normal completion")
            masters = [p for p in new if p.suffix.lower() == ".pdf" and not p.name.endswith("_email.pdf")]
            emails = [p for p in new if p.name.endswith("_email.pdf")]
            logs = [p for p in new if p.suffix.lower() == ".log"]
            require(len(new) == len(masters) + len(emails) + len(logs), "Unexpected published file")
            require(all(p.name.startswith("WinPDFMerge_") for p in new), "Unexpected output naming")
            expected_app_files = set(contents) | ({foreign.name} | {p.name for p in new} if default_output else set())
            expected_app_inventory = {"files": sorted(expected_app_files), "directories": app_inventory_before["directories"]}
            require_inventory(tree_inventory(app), expected_app_inventory)
            require(len(masters) == (1 if code in (0, 2) else 0), "Packaged master publication differs")
            require(len(emails) == (1 if email_state == "published" else 0), "Packaged email publication differs")
            require(len(logs) == (0 if label in ("missing-input", "invalid-preset") else 1), "Packaged log publication differs")
            log = logs[0].read_text(encoding="utf-8-sig") if logs else ""
            inspected = []
            if masters:
                expected_ids = identifiers if fixture == "normal" else ["T03-01-P01"]
                inspected.append(inspect_pdf(masters[0], expected_ids, case / "master renders"))
                require(f"PowerShell: {shell_version} ({result['environment']['shell_edition']})" in log, "Actual packaged shell identity differs")
                require("Application version: 1.0.0" in log and "PDFtk" in log and "version 2.02" in log, "Packaged version/native identity missing")
                require_email_result(log, email_state)
                if fixture == "normal":
                    ordered = re.findall(r"(?m)^Input [0-9]+: (.+)\r?$", log)
                    require([Path(p.strip()).name for p in ordered] == list(normal), "Logged natural input order/top-level visibility differs")
                if emails:
                    inspected.append(inspect_pdf(emails[0], expected_ids, case / "email renders"))
                    require(emails[0].stat().st_size < masters[0].stat().st_size, "Email has no size benefit")
                if email_state in ("published", "no_size_benefit", "failed"):
                    require("Ghostscript executable: " + str(arguments.ghostscript) in log and "version 10.08.0" in log, "Selected actual GS identity absent")
                    require("-dSAFER" in log and ("-dPDFSETTINGS=/ebook" if "ebook" in options else "-dPDFSETTINGS=/screen") in log, "Actual approved GS safety/preset arguments absent")
                    require("Ghostscript exit: " + ("1;" if email_state == "failed" else "0;") in log, "Actual GS exit differs")
                else:
                    require("Ghostscript arguments:" not in log, "Ghostscript unexpectedly executed")
                if email_state == "failed":
                    require("Initialization file gs_init.ps does not begin with an integer" in log and "PARTIAL SUCCESS:" in stdout, "Actual GS resource failure/retained-master receipt absent")
                if label == "skip-ignored-preset":
                    require("ignored" in stdout.lower(), "Ignored explicit preset diagnostic absent")
            else:
                require("SUCCESS:" not in stdout and " - Merged master:" not in stdout, "Failure advertises a published master")
            if batch:
                require(("Merge completed successfully." if code == 0 else "Partial success (exit code 2)." if code == 2 else "Merge failed with exit code 1.") in stdout, "Batch exit presentation differs")
                require("Press any key to continue" in stdout, "Batch pause prompt absent")
            case_report = {"label": label, "exit_code": code, "email_state": email_state, "batch": batch,
                           "invocation_label": row["label"], "case_root": str(case), "package_guard": True,
                           "source_foreign_guard": True, "before": before, "after": snapshot(guarded, case),
                           "source_file_set_before": source_before, "source_file_set_after": snapshot([p for p in source_dir.rglob("*") if p.is_file()], case),
                           "source_inventory_before": source_inventory_before, "source_inventory_after": tree_inventory(source_dir),
                           "cwd_inventory_before": cwd_before, "cwd_inventory_after": tree_inventory(cwd),
                           "package_inventory_before": app_inventory_before, "package_inventory_after": tree_inventory(app),
                           "expected_package_inventory_after": expected_app_inventory, "gs_resource_fault": fault,
                           "output_paths": [str(p) for p in new], "independent_pdfs": inspected,
                           "logs": [{"path": str(p), "sha256": file_hash(p)} for p in logs],
                           "scope": "Actual intact candidate/native operation; GS_LIB owned configuration fault" if email_state == "failed" else "Actual intact candidate/native operation; redirected automated batch input" if batch else "Actual intact candidate operation"}
            for log_path in logs:
                (capture_root / (label + ".application.log")).write_bytes(log_path.read_bytes())
            result["cases"].append(case_report)
            (capture_root / (label + ".case.json")).write_text(json.dumps(case_report, indent=2) + "\n", encoding="utf-8")
            print(json.dumps({"case": label, "exit_code": code, "result": "pass", "batch": batch}), flush=True)
        help_case = work / "public help"
        help_case.mkdir()
        help_app = extract(contents, help_case / "fresh install")
        before_help = snapshot([help_app / path for path in contents], help_case)
        help_inventory_before = tree_inventory(help_app)
        help_path = str(help_app / "WinPDFMerge.ps1").replace("'", "''")
        row, stdout, stderr = capture.invoke("public-help", [str(arguments.shell), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "RemoteSigned", "-Command", f"Get-Help '{help_path}' -Examples | Out-String -Width 200"], work, base_env)
        require(row["exit_code"] == 0 and "SkipEmail" in stdout and "EmailPreset" in stdout, "Packaged public help unavailable")
        require(snapshot([p for p in help_app.rglob("*") if p.is_file()], help_case) == before_help, "Help starts orchestration or changes package")
        require_inventory(tree_inventory(help_app), help_inventory_before)
        result["public_help"] = {"exit_code": 0, "package_unchanged": True, "no_outputs": True,
                                 "inventory_before": help_inventory_before, "inventory_after": tree_inventory(help_app)}
        result["result"] = "preparation_pass" if arguments.preparation else "pass"
    except Exception as error:
        result["failure"] = type(error).__name__ + ": " + str(error)
    finally:
        try:
            after_head = git("rev-parse", "HEAD")
            after_status = git("status", "--porcelain=v1", "--untracked-files=all")
            cache_ok = bool(cache) and all(file_hash(Path(row["path"])) == row["sha256"] for row in cache)
            asset_ok = bool(candidate_before) and candidate_before == {"zip": file_hash(arguments.zip), "checksums": file_hash(arguments.checksums)}
            guard = {"expected_head": after_head == arguments.expected_harness_commit,
                     "status_unchanged": initial_status is not None and initial_status == after_status,
                     "clean": not after_status, "driver_unchanged": driver_before == file_hash(Path(__file__)),
                     "approved_cache_unchanged": cache_ok, "candidate_assets_unchanged": asset_ok,
                     "parent_environment_unchanged": parent_env == environment_hash()}
            result["source_guard"] = guard
            result["driver_sha256_after"] = file_hash(Path(__file__))
            if not all(value for key, value in guard.items() if key != "clean") or (not arguments.preparation and not guard["clean"]):
                result["result"] = "fail"
                result.setdefault("failure", "Final source/cache/asset/environment guard failed")
        except Exception as error:
            result["result"] = "fail"
            result["guard_failure"] = str(error)
        result["invocation_count"] = len(capture.calls)
        (capture_root / "result.json").write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
        print("Candidate report: " + str(capture_root / "result.json"), flush=True)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, required=True)
    parser.add_argument("--expected-harness-commit", required=True)
    parser.add_argument("--candidate-source-commit", required=True)
    parser.add_argument("--zip", type=Path, required=True)
    parser.add_argument("--zip-sha256", required=True)
    parser.add_argument("--checksums", type=Path, required=True)
    parser.add_argument("--checksums-sha256", required=True)
    parser.add_argument("--shell", type=Path, required=True)
    parser.add_argument("--shell-kind", choices=("PS51", "PS7"), required=True)
    parser.add_argument("--pdftk", type=Path, required=True)
    parser.add_argument("--ghostscript", type=Path, required=True)
    parser.add_argument("--approved-cache-manifest", type=Path, required=True)
    parser.add_argument("--work-root", type=Path, required=True)
    parser.add_argument("--capture-root", type=Path, required=True)
    parser.add_argument("--preparation", action="store_true", help="Disclosed dirty preparation; never acceptance evidence")
    args = parser.parse_args()
    result = run(args)
    return 0 if result["result"] in ("pass", "preparation_pass") else 1


if __name__ == "__main__":
    raise SystemExit(main())
