#!/usr/bin/env python3
"""Check the evidence of an Excel run of the VBA harness against one commit.

The harness runs only in Excel, by hand (INSTALLATION.md, Recording an Excel
run). This reads what that run left in an evidence folder and checks it
against the candidate commit, without Excel:

- source identity: the modules exported back from the workbook equal, line
  for line, the commit's VBA components, and none is missing or extra;
- the three logs saved with TEST_Harness.SaveLog: the normal run passes,
  complete, with every case and assertion of the commit's suite in order,
  cleanup PASS and no recovery; the injected-failure and skipped-case runs
  end as documented; all three come from one host;
- host.txt: the Windows build, the references (exactly the four supported
  ones) and that Debug > Compile VBAProject completed.

It proves only that this evidence is consistent with the commit. It does not
run Excel, and a PASS validates no SFT calculation.
"""

from __future__ import annotations

import argparse
import hashlib
import re
import subprocess
import sys
from collections.abc import Sequence
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
    """What the commit's example suite runs: case and assertion identifiers in order."""

    cases: list[str]
    assertions: list[str]
    last_case: list[str]
    expected_cases: int
    expected_assertions: int


@dataclass
class Evidence:
    findings: list[str] = field(default_factory=list)
    hashes: dict[str, str] = field(default_factory=dict)
    environment: dict[str, str] = field(default_factory=dict)
    host: dict[str, str] = field(default_factory=dict)


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
    return records


def read_suite(source: bytes) -> Suite:
    text = source.decode("ascii", errors="replace").replace("\r\n", "\n")
    bodies = {match.group(1): match.group(0) for match in PROCEDURE.finditer(text)}
    constants = {name: int(value) for name, value in CONSTANT.findall(text)}
    cases: list[str] = []
    assertions: list[str] = []
    last_case: list[str] = []
    for name in CALL.findall(bodies.get("RunSuite", "")):
        body = bodies.get(name, "")
        case = BEGIN_CASE.findall(body)
        names = [f"{case[0] if case else name}/{item}" for item in ASSERTION.findall(body)]
        cases += case
        assertions += names
        last_case = case + names
    return Suite(cases, assertions, last_case, constants.get("EXPECTED_CASES", -1),
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


def summary(records: list[Record], key: str) -> Record | None:
    return next((record for record in records if record.key == key), None)


def check_shape(name: str, records: list[Record], mode: str, suite: Suite,
                problems: list[str]) -> None:
    """RUN, then ENVIRONMENT, first; RESULT last and only once."""
    if len(records) < 2 or records[0].key != "RUN" or records[1].key != "ENVIRONMENT":
        problems.append(f"{name}: does not start with RUN and ENVIRONMENT records")
        return
    run = records[0]
    if run.value != "examples" or run.fields.get("mode") != mode:
        problems.append(f"{name}:1: RUN is not suite examples in mode {mode}")
    if (run.fields.get("expected_cases"), run.fields.get("expected_assertions")) != \
            (str(suite.expected_cases), str(suite.expected_assertions)):
        problems.append(f"{name}:1: expected counts differ from the commit's suite")
    results = [record for record in records if record.key == "RESULT"]
    if not results or records[-1].key != "RESULT":
        problems.append(f"{name}: incomplete log: the last record is not RESULT")
    if len(results) > 1 or [record.key for record in records].count("RUN") > 1:
        problems.append(f"{name}: more than one run in the log")


def check_order(name: str, records: list[Record], cases: list[str], assertions: list[str],
                problems: list[str]) -> None:
    logged_cases = [record.value for record in records if record.key == "CASE"]
    logged = [record.value for record in records if record.key == "ASSERT"]
    if logged_cases != cases:
        problems.append(f"{name}: cases {logged_cases} differ from the commit's {cases}")
    if logged != assertions:
        problems.append(f"{name}: assertions differ from the commit's suite, in name or order")


def expect(name: str, record: Record | None, wanted: dict[str, str],
           problems: list[str]) -> None:
    if record is None:
        problems.append(f"{name}: no {next(iter(wanted))} record")
        return
    for key, value in wanted.items():
        actual = record.value if key == record.key else record.fields.get(key)
        if actual != value:
            problems.append(f"{name}:{record.line}: {record.key} {key}={actual}, "
                            f"expected {value}")


def check_normal(records: list[Record], suite: Suite, problems: list[str]) -> None:
    name = LOGS["NORMAL"]
    check_order(name, records, suite.cases, suite.assertions, problems)
    failed = [record for record in records if record.key in {"FAILURE", "ERROR"} or
              (record.key == "ASSERT" and record.fields.get("status") != "PASS")]
    for record in failed:
        problems.append(f"{name}:{record.line}: {record.key}={record.value}")
    expect(name, summary(records, "CLEANUP"), {"CLEANUP": "PASS"}, problems)
    expect(name, summary(records, "RECOVERY"), {"RECOVERY": "NONE"}, problems)
    expect(name, summary(records, "RESULT"), {
        "RESULT": "PASS", "suite": "examples", "operation": "PASS", "completeness": "COMPLETE",
        "cases": str(suite.expected_cases), "assertions": str(suite.expected_assertions),
        "failures": "0", "cleanup": "PASS", "recovery": "NONE"}, problems)


def check_injected(records: list[Record], suite: Suite, problems: list[str]) -> None:
    name = LOGS["INJECTED_FAILURE"]
    check_order(name, records, suite.cases, suite.assertions, problems)
    expect(name, summary(records, "CLEANUP"), {"CLEANUP": "PASS"}, problems)
    expect(name, summary(records, "RESULT"), {
        "RESULT": "FAIL", "operation": "FAIL", "completeness": "COMPLETE",
        "cases": str(suite.expected_cases), "assertions": str(suite.expected_assertions),
        "failures": "1", "cleanup": "PASS", "recovery": "NONE"}, problems)


def check_skipped(records: list[Record], suite: Suite, problems: list[str]) -> None:
    name = LOGS["SKIPPED_CASE"]
    skipped = set(suite.last_case)
    cases = [case for case in suite.cases if case not in skipped]
    assertions = [item for item in suite.assertions if item not in skipped]
    check_order(name, records, cases, assertions, problems)
    expect(name, summary(records, "FAILURE"), {"FAILURE": "harness.completeness"}, problems)
    expect(name, summary(records, "CLEANUP"), {"CLEANUP": "PASS"}, problems)
    expect(name, summary(records, "RESULT"), {
        "RESULT": "FAIL", "operation": "PASS", "completeness": "INCOMPLETE",
        "cases": str(len(cases)), "assertions": str(len(assertions)), "failures": "1",
        "cleanup": "PASS", "recovery": "NONE"}, problems)


def check_logs(folder: Path, suite: Suite, evidence: Evidence) -> None:
    checks = {"NORMAL": check_normal, "INJECTED_FAILURE": check_injected,
              "SKIPPED_CASE": check_skipped}
    environments = []
    for mode, filename in LOGS.items():
        records = read_log(folder / filename, evidence.findings)
        if not records:
            continue
        check_shape(filename, records, mode, suite, evidence.findings)
        checks[mode](records, suite, evidence.findings)
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


def check_host(folder: Path, evidence: Evidence) -> None:
    path = folder / HOST
    if not path.is_file():
        evidence.findings.append(f"{HOST}: missing")
        return
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


def check(root: Path, commit: str, folder: Path) -> Evidence:
    evidence = Evidence()
    sources = check_sources(root, commit, folder, evidence)
    suite = read_suite(sources.get(Path(EXAMPLES).name, b""))
    if not suite.cases:
        evidence.findings.append(f"no example suite in {EXAMPLES} at {commit[:12]}")
    check_logs(folder, suite, evidence)
    check_host(folder, evidence)
    return evidence


def report(commit: str, evidence: Evidence) -> str:
    environment = "; ".join(f"{key}={evidence.environment.get(key, '')}"
                            for key in ENVIRONMENT_FIELDS)
    lines = [f"Candidate commit: {commit}",
             f"Environment: {environment}",
             f"Windows: {evidence.host.get('windows', '')}",
             f"References: {evidence.host.get('references', '')}",
             f"Compile: {evidence.host.get('compile', '')}",
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
