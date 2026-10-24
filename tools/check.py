#!/usr/bin/env python3
"""Run the repository integrity checks locally and in CI; no Excel required."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

SCOPE = (
    "Repository tooling, label and issue-form configuration, test-case and changelog format, "
    "static VBA source and workflow checks only; no VBA has been compiled or run in Excel and no SFT "
    "calculation has been checked."
)
LABEL_POLICY = ".github/label-policy.json"


@dataclass(frozen=True)
class Gate:
    """One named check: a command, or a probe that returns its own outcome."""

    name: str
    tool: str
    command: tuple[str, ...] = ()
    probe: Callable[[Path, Path], tuple[int, str]] | None = None


@dataclass(frozen=True)
class Result:
    name: str
    status: str  # "pass", "fail" or "not-run"
    detail: str


def label(status: str) -> str:
    return status.upper().replace("-", " ")


def run(command: Sequence[str], root: Path, stdin: str | None = None) -> tuple[int, str]:
    try:
        completed = subprocess.run(
            list(command),
            cwd=root,
            input=stdin,
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as error:
        return 2, str(error)
    return completed.returncode, completed.stdout + completed.stderr


def ruff_complexity_probe(root: Path, scratch: Path) -> tuple[int, str]:
    """Prove the complexity ceiling is active: complexity 15 passes and 16 fails."""

    def source(branches: int) -> str:
        body = "".join(f"    if x == {i}:\n        return {i}\n" for i in range(branches - 1))
        return f"def probe(x: int) -> int:\n{body}    return -1\n"

    command = ["ruff", "check", "--no-cache", "--select", "C901",
               "--stdin-filename", "tools/complexity_probe.py", "-"]
    ok_code, ok_log = run(command, root, source(15))
    bad_code, bad_log = run(command, root, source(16))
    log = f"complexity 15 (expected pass):\n{ok_log}\ncomplexity 16 (expected fail):\n{bad_log}"
    rejected = bad_code != 0 and "is too complex (16 > 15)" in bad_log
    return (0 if ok_code == 0 and rejected else 1), log


def mypy_strict_probe(root: Path, scratch: Path) -> tuple[int, str]:
    """Prove strict mode is active: an unannotated function is rejected."""
    probe = scratch / "strict_probe.py"
    probe.write_text("def probe(value: int) -> int:\n    return value\n", encoding="utf-8")
    ok_code, ok_log = run(["mypy", "--no-pretty", "--no-incremental", str(probe)], root)
    probe.write_text("def probe(value):\n    return value\n", encoding="utf-8")
    bad_code, bad_log = run(["mypy", "--no-pretty", "--no-incremental", str(probe)], root)
    log = f"annotated (expected pass):\n{ok_log}\nunannotated (expected fail):\n{bad_log}"
    rejected = bad_code != 0 and "[no-untyped-def]" in bad_log
    return (0 if ok_code == 0 and rejected else 1), log


def default_gates() -> list[Gate]:
    python = sys.executable
    return [
        Gate("tool-tests", "python",
             (python, "-m", "unittest", "discover", "-s", "tools", "-p", "test_*.py")),
        Gate("labels-sync-self-test", "node",
             ("node", ".github/scripts/labels-sync.mjs", "--policy", LABEL_POLICY, "--self-test")),
        Gate("labels-manifest", "node",
             ("node", ".github/scripts/labels-sync.mjs", "--policy", LABEL_POLICY,
              "--mode", "validate")),
        Gate("labels-drift-self-test", "node",
             ("node", ".github/scripts/labels-drift.mjs", "--policy", LABEL_POLICY, "--self-test")),
        Gate("issue-config", "python",
             (python, "tools/check_issues.py", "--check-config")),
        Gate("test-cases", "python", (python, "tools/check_test_cases.py")),
        Gate("changelog", "python", (python, "tools/check_changelog.py")),
        Gate("vba-source", "python", (python, "tools/check_vba_source.py")),
        Gate("vba-conditionals", "python", (python, "tools/check_vba_conditionals.py")),
        Gate("vba-jumps", "python", (python, "tools/check_vba_jumps.py")),
        Gate("vba-public-api", "python", (python, "tools/check_vba_public_api.py")),
        Gate("ruff", "ruff", ("ruff", "check", "--no-cache", "tools")),
        Gate("ruff-complexity-probe", "ruff", probe=ruff_complexity_probe),
        Gate("mypy", "mypy", ("mypy", "--no-incremental")),
        Gate("mypy-strict-probe", "mypy", probe=mypy_strict_probe),
        Gate("actionlint", "actionlint", ("actionlint", "-no-color")),
    ]


def available(tool: str) -> bool:
    return tool == "python" or shutil.which(tool) is not None


def tool_versions(gates: Sequence[Gate], root: Path) -> dict[str, str]:
    flags = {"python": "--version", "node": "--version", "ruff": "--version",
             "mypy": "--version", "actionlint": "-version"}
    versions: dict[str, str] = {}
    for tool in sorted({gate.tool for gate in gates}):
        if not available(tool):
            versions[tool] = "not installed"
            continue
        executable = sys.executable if tool == "python" else tool
        code, log = run([executable, flags.get(tool, "--version")], root)
        lines = log.strip().splitlines()
        versions[tool] = lines[0] if code == 0 and lines else f"unknown (exit {code})"
    return versions


def run_gates(gates: Sequence[Gate], root: Path, output: Path) -> list[Result]:
    results = []
    with tempfile.TemporaryDirectory(prefix="probe-", dir=output) as scratch:
        for gate in gates:
            if not available(gate.tool):
                code, log = -1, f"{gate.tool} is not installed or not on PATH"
            elif gate.probe is not None:
                code, log = gate.probe(root, Path(scratch))
            else:
                code, log = run(gate.command, root)
            (output / f"{gate.name}.log").write_text(log, encoding="utf-8", newline="\n")
            status = "not-run" if code == -1 else ("pass" if code == 0 else "fail")
            detail = f"{gate.tool} not available" if code == -1 else f"exit {code}"
            results.append(Result(gate.name, status, detail))
            print(f"{gate.name}: {label(status)}", flush=True)
            if status == "fail":
                print(log, flush=True)
    return results


def overall_status(results: Sequence[Result], *, ci: bool, dirty: bool,
                   allow_missing: bool) -> str:
    """PASS needs every gate to pass; a gate that did not run is never a PASS."""
    if any(result.status == "fail" for result in results) or (ci and dirty):
        return "fail"
    if any(result.status == "not-run" for result in results):
        return "incomplete" if allow_missing and not ci else "fail"
    return "pass"


def git(root: Path, *args: str) -> str:
    return subprocess.run(["git", "-C", str(root), *args], capture_output=True,
                          text=True, check=True).stdout


def render_markdown(report: dict[str, Any]) -> str:
    lines = [
        "# Repository integrity",
        "",
        f"Status: **{label(report['status'])}**",
        "",
        f"Candidate: `{report['candidate_sha']}`; mode: `{report['mode']}`; "
        f"dirty: `{report['dirty']}`",
        "",
        report["scope"],
        "",
        "| Gate | Result | Detail |",
        "| --- | --- | --- |",
    ]
    lines += [f"| `{c['name']}` | {label(c['status'])} | {c['detail']} |"
              for c in report["checks"]]
    lines += ["", "| Tool | Version |", "| --- | --- |"]
    lines += [f"| {tool} | {version} |" for tool, version in report["tools"].items()]
    return "\n".join(lines) + "\n"


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    parser.add_argument("--output-dir", type=Path, default=Path("test-results"))
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--ci", action="store_true",
                      help="require a clean checkout and every tool")
    mode.add_argument("--allow-missing", action="store_true",
                      help="report gates whose tool is missing as NOT RUN and exit 0 "
                           "with status INCOMPLETE (never PASS)")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None, gates: Sequence[Gate] | None = None) -> int:
    args = parse_args(argv)
    root = args.root.resolve()
    output = (root / args.output_dir).resolve()
    output.mkdir(parents=True, exist_ok=True)
    selected = list(default_gates() if gates is None else gates)
    sha = git(root, "rev-parse", "HEAD").strip()
    dirty = bool(git(root, "status", "--porcelain", "--untracked-files=normal"))
    results = run_gates(selected, root, output)
    status = overall_status(results, ci=args.ci, dirty=dirty, allow_missing=args.allow_missing)
    if args.ci and dirty:
        print("working tree is dirty: --ci requires a clean checkout", flush=True)
    report: dict[str, Any] = {
        "schema_version": 1,
        "candidate_sha": sha,
        "dirty": dirty,
        "mode": "ci" if args.ci else "local",
        "status": status,
        "scope": SCOPE,
        "checks": [{"name": r.name, "status": r.status, "detail": r.detail} for r in results],
        "tools": tool_versions(selected, root),
    }
    (output / "repository-integrity.json").write_text(
        json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8", newline="\n")
    (output / "repository-integrity.md").write_text(
        render_markdown(report), encoding="utf-8", newline="\n")
    print(f"repository-integrity: {label(status)} ({SCOPE})", flush=True)
    return 0 if status in ("pass", "incomplete") else 1


if __name__ == "__main__":
    raise SystemExit(main())
