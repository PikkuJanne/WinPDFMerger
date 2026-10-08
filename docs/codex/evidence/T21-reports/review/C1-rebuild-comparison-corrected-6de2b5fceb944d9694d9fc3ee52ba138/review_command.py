"""Independent reviewer command receipts; ignored development evidence only."""
from __future__ import annotations
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import uuid

REPO = Path(__file__).resolve().parents[3]
ROOT = REPO / "tests" / ".work" / "T21-review"

def now():
    return datetime.now(timezone.utc).isoformat()

def git(*args):
    result = subprocess.run(["git", "-C", str(REPO), *args], capture_output=True, check=True)
    return result.stdout.decode("utf-8", errors="strict").strip()

def sha(raw):
    return hashlib.sha256(raw).hexdigest()

def sources():
    paths = git("ls-files", "-z", "--cached", "--others", "--exclude-standard").split("\0")
    records = []
    for value in sorted(set(paths)):
        if not value:
            continue
        path = REPO / value
        if path.is_file():
            raw = path.read_bytes()
            records.append({"path": value, "sha256": sha(raw), "bytes": len(raw)})
    return records

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--label", required=True)
    parser.add_argument("--expected-commit")
    parser.add_argument("--timeout", type=int, default=300)
    parser.add_argument("command", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    command = args.command[1:] if args.command[:1] == ["--"] else args.command
    if not command:
        raise ValueError("An explicit argv command is required.")
    if any(char not in "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_" for char in args.label):
        raise ValueError("Use a simple receipt label.")
    head = git("rev-parse", "HEAD")
    if args.expected_commit and head != args.expected_commit:
        raise ValueError("Current source does not match the expected reviewed commit.")
    directory = ROOT / (args.label + "-" + uuid.uuid4().hex)
    directory.mkdir(parents=True, exist_ok=False)
    producer = Path(__file__).read_bytes()
    (directory / "review_command.py").write_bytes(producer)
    command_sources = []
    for index, operand in enumerate(command):
        candidate = Path(operand)
        if candidate.suffix == ".py":
            candidate = (candidate if candidate.is_absolute() else REPO / candidate).resolve()
            if candidate.is_file() and candidate.is_relative_to(REPO):
                raw_source = candidate.read_bytes()
                copied_name = "command-source-" + str(index) + "-" + candidate.name
                (directory / copied_name).write_bytes(raw_source)
                command_sources.append({"path": str(candidate), "copied": copied_name, "sha256": sha(raw_source), "bytes": len(raw_source)})
    before = sources()
    metadata = {
        "task": "T21", "reviewer": "/root/review", "authorship": "Independent subagent reviewer; did not author tracked implementation or records.",
        "label": args.label, "observed_start_utc": now(), "command_argv": command,
        "working_directory": str(REPO), "commit_under_test": head,
        "status_before": git("status", "--porcelain=v1"), "source_before": before,
        "source_before_sha256": sha(json.dumps(before, sort_keys=True).encode()),
        "producer_sha256": sha(producer), "command_sources": command_sources, "python": sys.version, "platform": sys.platform,
        "scope": "Source-bound independently invoked review/development command; this receipt alone is not Windows native/manual acceptance."
    }
    (directory / "metadata.json").write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")
    try:
        execution = subprocess.run(command, cwd=REPO, capture_output=True, timeout=args.timeout)
        stdout, stderr, code = execution.stdout, execution.stderr, execution.returncode
    except subprocess.TimeoutExpired as error:
        stdout, stderr, code = error.stdout or b"", error.stderr or b"", 124
    (directory / "stdout.txt").write_bytes(stdout)
    (directory / "stderr.txt").write_bytes(stderr)
    after = sources()
    result = {**metadata,
        "observed_end_utc": now(), "exit_code": code,
        "stdout_sha256": sha(stdout), "stdout_bytes": len(stdout),
        "stderr_sha256": sha(stderr), "stderr_bytes": len(stderr),
        "source_after": after, "source_unchanged_during_command": before == after,
        "head_after": git("rev-parse", "HEAD"), "status_after": git("status", "--porcelain=v1"),
        "result": "pass" if code == 0 and before == after and head == git("rev-parse", "HEAD") else "fail"
    }
    (directory / "receipt.json").write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"result": result["result"], "receipt": str(directory / "receipt.json"), "exit_code": code, "source_unchanged": before == after}))
    if code:
        print(stderr.decode("utf-8", errors="replace"), file=sys.stderr)
    raise SystemExit(0 if result["result"] == "pass" else 1)

if __name__ == "__main__":
    main()
