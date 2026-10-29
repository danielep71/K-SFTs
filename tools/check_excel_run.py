#!/usr/bin/env python3
"""Check the evidence of an Excel run of the VBA harness against one commit.

The harness runs only in Excel, by hand or through tools/Sync-KSFT-VBA.ps1
-RecordRun (INSTALLATION.md, Recording an Excel run). This reads what that run left in an evidence folder and checks it
against the candidate commit, without Excel:

- source identity: the modules exported back from the workbook equal, line
  for line, the commit's VBA components, and none is missing or extra;
- the three logs saved with TEST_Harness.SaveLog: each holds, record for
  record, what the commit's suite writes in its mode (no record missing,
  extra or out of order, no ERROR), with cleanup PASS and no recovery: the
  normal run passes with every case and assertion; the injected-failure run
  fails only its designated assertion; the skipped-case run lacks only the
  last case and fails only harness.completeness; all three come from one
  64-bit host;
- host.txt: the Windows build, the references (exactly the four supported
  ones) and compile=PASS, no compile error reported. When the folder has
  none and the check runs on Windows, it writes one: the Windows build read
  from this machine, the four supported references and compile=PASS.

It proves only that this evidence is consistent with the commit. It does not
run Excel, and a PASS validates no SFT calculation.
"""

from __future__ import annotations

import argparse
import hashlib
import re
import subprocess
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VBA_SUFFIXES = (".bas", ".cls", ".frm")
EXAMPLES = "tests/modules/TEST_Examples.bas"
LOGS = {"NORMAL": "examples.log", "INJECTED_FAILURE": "injected.log",
        "SKIPPED_CASE": "skipped.log"}
EXPORTED = "exported"
HOST = "host.txt"
REFERENCES = ("Visual Basic For Applications", "Microsoft Excel 16.0 Object Library",
              "OLE Automation", "Microsoft Office 16.0 Object Library")
ENVIRONMENT_FIELDS = ("host", "version", "build", "os", "office", "runtime")
HOST_FIELDS = ("windows", "references", "compile")
RECORD = re.compile(r"^([A-Z_]+)=(.*)$")
CONSTANT = re.compile(r"^\s*Private Const (\w+)\s+As Long = (\d+)\b", re.M)
PROCEDURE = re.compile(r"^(?:Public|Private) Sub (\w+)\(.*?^End Sub$", re.M | re.S)
BEGIN_CASE = re.compile(r'TEST_Harness\.BeginCase "([^"]+)"')
ASSERTION = re.compile(r'TEST_Harness\.Assert\w+ "([^"]+)"')
CALL = re.compile(r"^\s+(Case\w+)$", re.M)


@dataclass
class Record:
    """One log line: its key, its first value and its named fields."""

    key: str
    value: str
    fields: dict[str, str]
    line: int


@dataclass
class Suite:
    """What the commit's example suite runs: case and assertion identifiers in order,
    the identifiers of the last case (skipped in SKIPPED_CASE) and the assertions of
    the case whose expected value INJECTED_FAILURE changes."""

    cases: list[str]
    assertions: list[str]
    last_case: list[str]
    injected: list[str]
    expected_cases: int
    expected_assertions: int


@dataclass
class Evidence:
    findings: list[str] = field(default_factory=list)
    hashes: dict[str, str] = field(default_factory=dict)
    environment: dict[str, str] = field(default_factory=dict)
    host: dict[str, str] = field(default_factory=dict)
    host_written: bool = False


def git(root: Path, *args: str) -> bytes:
    completed = subprocess.run(["git", "-C", str(root), *args], capture_output=True, check=False)
    if completed.returncode:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    return completed.stdout


def components(root: Path, commit: str) -> dict[str, bytes]:
    """The commit's VBA components, by file name, with LF line ends."""
    listed = git(root, "ls-tree", "-r", "-z", "--name-only", commit).split(b"\0")
    paths = [item.decode("utf-8") for item in listed if item.decode("utf-8").endswith(VBA_SUFFIXES)]
    return {Path(path).name: git(root, "show", f"{commit}:{path}") for path in paths}


def parse_record(text: str, line: int) -> Record | None:
    match = RECORD.match(text)
    if not match:
        return None
    first, *rest = match.group(2).split("; ")
    fields = {}
    for part in rest:
        name, _, value = part.partition("=")
        fields[name] = value
    if "=" in first:
        name, _, value = first.partition("=")
        fields[name] = value
        first = ""
    return Record(match.group(1), first, fields, line)


def read_log(path: Path, problems: list[str]) -> list[Record]:
    if not path.is_file():
        problems.append(f"{path.name}: missing")
        return []
    records = []
    for number, text in enumerate(path.read_text(encoding="utf-8", errors="replace")
                                  .splitlines(), 1):
        record = parse_record(text, number)
        if record is None:
            problems.append(f"{path.name}:{number}: not a KEY=value record")
        else:
            records.append(record)
    if not records:
        problems.append(f"{path.name}: incomplete log: no record")
    return records


def read_suite(source: bytes) -> Suite:
    text = source.decode("ascii", errors="replace").replace("\r\n", "\n")
    bodies = {match.group(1): match.group(0) for match in PROCEDURE.finditer(text)}
    constants = {name: int(value) for name, value in CONSTANT.findall(text)}
    cases: list[str] = []
    assertions: list[str] = []
    last_case: list[str] = []
    injected: list[str] = []
    for name in CALL.findall(bodies.get("RunSuite", "")):
        body = bodies.get(name, "")
        case = BEGIN_CASE.findall(body)
        names = [f"{case[0] if case else name}/{item}" for item in ASSERTION.findall(body)]
        cases += case
        assertions += names
        last_case = case + names
        if 'mMode = "INJECTED_FAILURE"' in body:
            injected += names
    return Suite(cases, assertions, last_case, injected, constants.get("EXPECTED_CASES", -1),
                 constants.get("EXPECTED_ASSERTIONS", -1))


def check_sources(root: Path, commit: str, folder: Path, evidence: Evidence) -> dict[str, bytes]:
    expected = components(root, commit)
    exported = folder / EXPORTED
    present = {path.name: path for path in exported.iterdir()} if exported.is_dir() else {}
    if not expected:
        evidence.findings.append(f"commit {commit} holds no VBA component")
    for name, data in sorted(expected.items()):
        evidence.hashes[name] = git_blob(data)
        if name not in present:
            evidence.findings.append(f"{EXPORTED}/{name}: not exported from the workbook")
            continue
        actual = present[name].read_bytes().replace(b"\r\n", b"\n")
        if actual != data:
            evidence.findings.append(f"{EXPORTED}/{name}: differs from {name} at {commit[:12]} "
                                     f"(first difference at line {first_difference(actual, data)})")
    for name in sorted(set(present) - set(expected)):
        evidence.findings.append(f"{EXPORTED}/{name}: not a VBA component of the commit")
    return expected


def git_blob(data: bytes) -> str:
    """The Git blob id of ``data``, as ``git hash-object`` gives it."""
    return hashlib.sha1(b"blob %d\0" % len(data) + data).hexdigest()


def first_difference(actual: bytes, expected: bytes) -> int:
    for number, (left, right) in enumerate(zip(actual.split(b"\n"), expected.split(b"\n"),
                                               strict=False), 1):
        if left != right:
            return number
    return min(actual.count(b"\n"), expected.count(b"\n")) + 1


@dataclass
class Plan:
    """What one mode's run must log: its cases and assertions in order, the
    failures it must report (assertions, or harness identifiers) and its RESULT."""

    mode: str
    cases: list[str]
    assertions: list[str]
    failing: list[str]
    result: dict[str, str]


def plans(suite: Suite) -> list[Plan]:
    skipped = set(suite.last_case)
    cases = [case for case in suite.cases if case not in skipped]
    assertions = [item for item in suite.assertions if item not in skipped]

    def result(verdict: str, operation: str, complete: bool, kept: tuple[list[str], list[str]],
               failures: int) -> dict[str, str]:
        return {"RESULT": verdict, "suite": "examples", "operation": operation,
                "completeness": "COMPLETE" if complete else "INCOMPLETE",
                "cases": str(len(kept[0])), "assertions": str(len(kept[1])),
                "failures": str(failures), "cleanup": "PASS", "recovery": "NONE"}

    every = (suite.cases, suite.assertions)
    return [Plan("NORMAL", *every, [], result("PASS", "PASS", True, every, 0)),
            Plan("INJECTED_FAILURE", *every, suite.injected,
                 result("FAIL", "FAIL", True, every, len(suite.injected))),
            Plan("SKIPPED_CASE", cases, assertions, ["harness.completeness"],
                 result("FAIL", "PASS", False, (cases, assertions), 1))]


def sequence(plan: Plan, suite: Suite) -> list[tuple[str, str]]:
    """Every record the run writes, as (key, first value), in order."""
    records = [("RUN", "examples"), ("ENVIRONMENT", "")]
    for case in plan.cases:
        records.append(("CASE", case))
        for item in (name for name in plan.assertions if name.startswith(f"{case}/")):
            records.append(("ASSERT", item))
            if item in plan.failing:
                records.append(("FAILURE", item))
    records += [("FAILURE", item) for item in plan.failing if "/" not in item]
    records += [("CASES", str(len(plan.cases))), ("ASSERTIONS", str(len(plan.assertions))),
                ("FAILURES", str(len(plan.failing)))]
    records += [("PRIMARY_FAILURE", plan.failing[0])] if plan.failing else []
    return records + [("CLEANUP", "PASS"), ("RECOVERY", "NONE"),
                      ("RESULT", plan.result["RESULT"])]


def check_sequence(name: str, records: list[Record], expected: list[tuple[str, str]],
                   problems: list[str]) -> bool:
    """The log holds exactly the expected records, in order: none missing, extra,
    repeated or contradicting the run's outcome."""
    for record, (key, value) in zip(records, expected, strict=False):
        if (record.key, record.value) != (key, value):
            problems.append(f"{name}:{record.line}: {record.key}={record.value}, "
                            f"expected {key}={value}")
            return False
    if len(records) < len(expected):
        key, value = expected[len(records)]
        problems.append(f"{name}: incomplete log: it ends before {key}={value}")
        return False
    if len(records) > len(expected):
        extra = records[len(expected)]
        problems.append(f"{name}:{extra.line}: {extra.key}={extra.value} after RESULT "
                        "(more than one run?)")
        return False
    return True


def expect(name: str, record: Record, wanted: dict[str, str], problems: list[str]) -> None:
    for key, value in wanted.items():
        actual = record.value if key == record.key else record.fields.get(key)
        if actual != value:
            problems.append(f"{name}:{record.line}: {record.key} {key}={actual}, "
                            f"expected {value}")


def check_fields(name: str, records: list[Record], plan: Plan, suite: Suite,
                 problems: list[str]) -> None:
    expected = {"CASES": str(suite.expected_cases), "ASSERTIONS": str(suite.expected_assertions)}
    for record in records:
        if record.key == "RUN":
            expect(name, record, {"mode": plan.mode, "expected_cases": expected["CASES"],
                                  "expected_assertions": expected["ASSERTIONS"]}, problems)
        elif record.key == "ASSERT":
            status = "FAIL" if record.value in plan.failing else "PASS"
            expect(name, record, {"status": status}, problems)
        elif record.key in expected:
            expect(name, record, {"expected": expected[record.key]}, problems)
        elif record.key == "FAILURES":
            expect(name, record, {"errors": "0"}, problems)
        elif record.key == "RESULT" and {"RESULT": record.value, **record.fields} != plan.result:
            problems.append(f"{name}:{record.line}: RESULT={record.value} with "
                            f"{record.fields}, expected {plan.result}")


def check_logs(folder: Path, suite: Suite, evidence: Evidence) -> None:
    environments = []
    for plan in plans(suite):
        filename = LOGS[plan.mode]
        records = read_log(folder / filename, evidence.findings)
        if not records:
            continue
        if check_sequence(filename, records, sequence(plan, suite), evidence.findings):
            check_fields(filename, records, plan, suite, evidence.findings)
        if len(records) > 1 and records[1].key == "ENVIRONMENT":
            environments.append(records[1].fields)
    for environment in environments:
        missing = [name for name in ENVIRONMENT_FIELDS if not environment.get(name)]
        if missing:
            evidence.findings.append(f"ENVIRONMENT lacks {', '.join(missing)}")
    if len({tuple(sorted(item.items())) for item in environments}) > 1:
        evidence.findings.append("the logs come from different hosts (ENVIRONMENT differs)")
    if environments:
        evidence.environment = environments[0]


def windows_name(product: str, display_version: str, build: str, revision: str) -> str:
    """The Windows build as winver shows it, e.g. Windows 11 Pro, version 24H2, OS build 26100.4061.

    The registry's ProductName still says Windows 10 on Windows 11, whose builds start at 22000.
    """
    if build.isdigit() and int(build) >= 22000:
        product = product.replace("Windows 10", "Windows 11", 1)
    return f"{product}, version {display_version}, OS build {build}.{revision}"


def local_windows() -> str | None:
    """The Windows build of this machine, or None when it is not Windows or cannot be read."""
    if sys.platform != "win32":
        return None
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as key:
            values = [str(winreg.QueryValueEx(key, name)[0])
                      for name in ("ProductName", "DisplayVersion", "CurrentBuild", "UBR")]
    except OSError:
        return None
    return windows_name(*values)


def write_host(path: Path, windows: str) -> None:
    lines = [f"windows={windows}", f"references={'; '.join(REFERENCES)}", "compile=PASS"]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def check_host(folder: Path, evidence: Evidence, windows: Callable[[], str | None]) -> None:
    path = folder / HOST
    if not path.is_file():
        build = windows()
        if build is None:
            evidence.findings.append(f"{HOST}: missing, and not written: no Windows build to read here")
            return
        write_host(path, build)
        evidence.host_written = True
    host = {}
    for text in path.read_text(encoding="utf-8").splitlines():
        key, separator, value = text.partition("=")
        if separator and not text.startswith("#"):
            host[key.strip()] = value.strip()
    evidence.host = host
    for key in HOST_FIELDS:
        if not host.get(key):
            evidence.findings.append(f"{HOST}: no {key}=")
    references = tuple(item.strip() for item in host.get("references", "").split(";"))
    if host.get("references") and sorted(references) != sorted(REFERENCES):
        evidence.findings.append(f"{HOST}: references must be exactly {'; '.join(REFERENCES)}")
    if host.get("compile") and host["compile"] != "PASS":
        evidence.findings.append(f"{HOST}: compile={host['compile']}, expected PASS")
    if evidence.environment and evidence.environment.get("office") != "64-bit":
        evidence.findings.append("not the supported target: Office is not 64-bit")


def check(root: Path, commit: str, folder: Path,
          windows: Callable[[], str | None] = local_windows) -> Evidence:
    evidence = Evidence()
    sources = check_sources(root, commit, folder, evidence)
    suite = read_suite(sources.get(Path(EXAMPLES).name, b""))
    if not suite.cases:
        evidence.findings.append(f"no example suite in {EXAMPLES} at {commit[:12]}")
    if len(suite.injected) != 1:
        evidence.findings.append(f"{EXAMPLES} at {commit[:12]}: INJECTED_FAILURE must change "
                                 f"exactly one assertion, found {suite.injected}")
    check_logs(folder, suite, evidence)
    check_host(folder, evidence, windows)
    return evidence


def report(commit: str, evidence: Evidence) -> str:
    environment = "; ".join(f"{key}={evidence.environment.get(key, '')}"
                            for key in ENVIRONMENT_FIELDS)
    lines = [f"Candidate commit: {commit}",
             f"Environment: {environment}",
             f"Windows: {evidence.host.get('windows', '')}",
             f"References: {evidence.host.get('references', '')}",
             f"Compile: {evidence.host.get('compile', '')}",
             *([f"{HOST}: written by this check (Windows read from this machine)"]
               if evidence.host_written else []),
             "Imported components (Git blob ids at the commit):"]
    lines += [f"  {name} {blob}" for name, blob in sorted(evidence.hashes.items())]
    lines += [f"FAIL: {finding}" for finding in evidence.findings]
    status = "FAIL" if evidence.findings else "PASS"
    lines.append(f"excel-run: {status} (evidence checked against the commit; Excel is not run "
                 "here and no SFT calculation is validated)")
    return "\n".join(lines)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("evidence", type=Path, help="folder holding the run's evidence")
    parser.add_argument("--commit", required=True, help="the commit the workbook was built from")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    try:
        commit = git(root, "rev-parse", "--verify", f"{args.commit}^{{commit}}").decode().strip()
    except RuntimeError as error:
        print(f"FAIL: commit {args.commit} not found: {error}")
        return 1
    evidence = check(root, commit, args.evidence)
    print(report(commit, evidence))
    return 1 if evidence.findings else 0


if __name__ == "__main__":
    sys.exit(main())
