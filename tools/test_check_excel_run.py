"""Positive and negative evidence fixtures for check_excel_run.py.

Each fixture is an evidence folder built in a temporary directory: the
commit's VBA components as the VBE exports them (CRLF), the three logs the
harness would write for the commit's example suite, and host.txt. The logs
are synthetic: they follow tests/README.md (Log format) and the commit's
TEST_Examples.bas, not a recorded Excel run.
"""

from __future__ import annotations

import contextlib
import io
import re
import subprocess
import sys
import tempfile
import types
import unittest
from collections.abc import Callable, Iterator
from pathlib import Path
from unittest import mock

import check_excel_run
from check_excel_run import EXAMPLES, Suite, components, read_suite

ROOT = Path(__file__).resolve().parents[1]
COMMIT = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "HEAD"], capture_output=True,
                        check=True, text=True).stdout.strip()
ENVIRONMENT = ("ENVIRONMENT=host=Microsoft Excel; version=16.0; build=17928; "
               "os=Windows (64-bit) NT 10.00; office=64-bit; runtime=VBA7")
HOST = ("windows=Windows 11 Pro, version 24H2, OS build 26100.4061\n"
        "references=Visual Basic For Applications; Microsoft Excel 16.0 Object Library; "
        "OLE Automation; Microsoft Office 16.0 Object Library\n"
        "compile=PASS\n")
UNCHANGED = "CLEANUP=PASS; detail=Excel state unchanged (Calculation, DisplayAlerts, " \
            "EnableEvents, ScreenUpdating)"


def suite() -> Suite:
    return read_suite(components(ROOT, COMMIT)[Path(EXAMPLES).name])


def run_log(mode: str) -> list[str]:
    """The records the harness writes for one run of the example suite."""
    examples = suite()
    skipped = set(examples.last_case) if mode == "SKIPPED_CASE" else set()
    lines = [f"RUN=examples; mode={mode}; expected_cases={examples.expected_cases}; "
             f"expected_assertions={examples.expected_assertions}; started=2026-10-24 17:00:00",
             ENVIRONMENT]
    failures: list[str] = []
    cases = assertions = 0
    for case in examples.cases:
        if case in skipped:
            continue
        lines.append(f"CASE={case}")
        cases += 1
        for name in (item for item in examples.assertions if item.startswith(f"{case}/")):
            assertions += 1
            if mode == "INJECTED_FAILURE" and name in examples.injected:
                detail = "expected=1; actual=0.75"
                lines.append(f"ASSERT={name}; status=FAIL; {detail}")
                lines.append(f"FAILURE={name}; {detail}")
                failures.append(f"{name}; {detail}")
            else:
                lines.append(f"ASSERT={name}; status=PASS")
    if mode == "SKIPPED_CASE":
        detail = (f"expected cases={examples.expected_cases}, "
                  f"assertions={examples.expected_assertions}; observed cases={cases}, "
                  f"assertions={assertions}")
        lines.append(f"FAILURE=harness.completeness; {detail}")
        failures.append(f"harness.completeness; {detail}")
    complete = "INCOMPLETE" if skipped else "COMPLETE"
    operation = "FAIL" if mode == "INJECTED_FAILURE" else "PASS"
    verdict = "PASS" if mode == "NORMAL" else "FAIL"
    lines += [f"CASES={cases}; expected={examples.expected_cases}",
              f"ASSERTIONS={assertions}; expected={examples.expected_assertions}",
              f"FAILURES={len(failures)}; errors=0"]
    lines += [f"PRIMARY_FAILURE={failures[0]}"] if failures else []
    lines += [UNCHANGED, "RECOVERY=NONE",
              f"RESULT={verdict}; suite=examples; operation={operation}; completeness={complete}; "
              f"cases={cases}; assertions={assertions}; failures={len(failures)}; "
              "cleanup=PASS; recovery=NONE"]
    return lines


@contextlib.contextmanager
def evidence(change: Callable[[Path], None] | None = None) -> Iterator[Path]:
    """A complete, valid evidence folder, then ``change`` applied to it."""
    with tempfile.TemporaryDirectory(prefix="excel-run-") as temporary:
        folder = Path(temporary)
        (folder / "exported").mkdir()
        for name, data in components(ROOT, COMMIT).items():
            (folder / "exported" / name).write_bytes(data.replace(b"\n", b"\r\n"))
        for mode, filename in check_excel_run.LOGS.items():
            (folder / filename).write_bytes(("\r\n".join(run_log(mode)) + "\r\n").encode())
        (folder / "host.txt").write_text(HOST, encoding="utf-8")
        if change:
            change(folder)
        yield folder


def edit(path: str, old: str, new: str) -> Callable[[Path], None]:
    def apply(folder: Path) -> None:
        text = (folder / path).read_bytes().decode("ascii")
        assert old in text, (path, old)
        (folder / path).write_bytes(text.replace(old, new, 1).encode("ascii"))
    return apply


def findings(change: Callable[[Path], None] | None = None) -> list[str]:
    with evidence(change) as folder:
        return check_excel_run.check(ROOT, COMMIT, folder, lambda: None).findings


class EvidenceTest(unittest.TestCase):
    def assertFinding(self, change: Callable[[Path], None], text: str) -> None:
        found = findings(change)
        self.assertTrue(any(text in item for item in found), found)

    def test_complete_evidence_passes(self) -> None:
        self.assertEqual(findings(), [])

    def test_report_names_the_commit_and_every_component(self) -> None:
        with evidence() as folder, contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(check_excel_run.main([str(folder), "--commit", COMMIT,
                                                   "--root", str(ROOT)]), 0)
        text = output.getvalue()
        self.assertIn(f"Candidate commit: {COMMIT}", text)
        for name in components(ROOT, COMMIT):
            self.assertIn(f"  {name} ", text)
        self.assertIn("excel-run: PASS", text)

    def test_unknown_commit_fails(self) -> None:
        with evidence() as folder, contextlib.redirect_stdout(io.StringIO()):
            self.assertEqual(check_excel_run.main([str(folder), "--commit", "0" * 40,
                                                   "--root", str(ROOT)]), 1)

    def test_wrong_source_identity(self) -> None:
        self.assertFinding(edit("exported/TEST_Harness.bas", "Option Explicit",
                                "Option Explicit\r\n'local edit"), "differs from TEST_Harness.bas")
        self.assertFinding(lambda folder: (folder / "exported/TEST_Examples.bas").unlink(),
                           "TEST_Examples.bas: not exported")
        def extra(folder: Path) -> None:
            (folder / "exported/Module1.bas").write_bytes(b"x")
        self.assertFinding(extra, "Module1.bas: not a VBA component of the commit")

    def test_incomplete_logs(self) -> None:
        def truncate(folder: Path) -> None:
            lines = (folder / "examples.log").read_bytes().splitlines()
            (folder / "examples.log").write_bytes(b"\r\n".join(lines[:-1]))
        self.assertFinding(truncate, "examples.log: incomplete log")
        self.assertFinding(lambda folder: (folder / "skipped.log").unlink(), "skipped.log: missing")
        def empty(folder: Path) -> None:
            for filename in check_excel_run.LOGS.values():
                (folder / filename).write_bytes(b"")
        found = findings(empty)
        for filename in check_excel_run.LOGS.values():
            self.assertIn(f"{filename}: incomplete log: no record", found)
        self.assertFinding(edit("examples.log", "CASE=text-exact\r\n", ""),
                           "expected CASE=text-exact")
        self.assertFinding(edit("examples.log", "RESULT=PASS", "stray line\r\nRESULT=PASS"),
                           "not a KEY=value record")

    def test_a_run_that_did_not_pass(self) -> None:
        self.assertFinding(edit("examples.log", "status=PASS", "status=FAIL"),
                           "examples.log:4: ASSERT status=FAIL, expected PASS")
        self.assertFinding(edit("examples.log", "RESULT=PASS", "RESULT=FAIL"), "RESULT=FAIL")
        self.assertFinding(edit("examples.log", "mode=NORMAL", "mode=SKIPPED_CASE"),
                           "RUN mode=SKIPPED_CASE, expected NORMAL")
        self.assertFinding(edit("examples.log", "ASSERTIONS=18", "ERROR=run; number=5\r\n"
                                "ASSERTIONS=18"), "ERROR=run, expected ASSERTIONS=18")
        self.assertFinding(edit("examples.log", "RECOVERY=NONE\r\n",
                                "RECOVERY=NONE\r\nRUN=examples; mode=NORMAL\r\n"),
                           "RUN=examples, expected RESULT=PASS")
        def appended(folder: Path) -> None:
            log = (folder / "examples.log").read_bytes()
            (folder / "examples.log").write_bytes(log + log)
        self.assertFinding(appended, "after RESULT (more than one run?)")

    def test_failed_cleanup(self) -> None:
        def dirty(folder: Path) -> None:
            edit("examples.log", UNCHANGED, "CLEANUP=FAIL; detail=Excel state changed during "
                 "the run: Calculation")(folder)
            edit("examples.log", "RECOVERY=NONE", "RECOVERY=REQUIRED; reason=changed")(folder)
            edit("examples.log", "cleanup=PASS; recovery=NONE",
                 "cleanup=FAIL; recovery=REQUIRED")(folder)
        self.assertFinding(dirty, "CLEANUP=FAIL, expected CLEANUP=PASS")
        for filename in check_excel_run.LOGS.values():
            with self.subTest(log=filename):
                self.assertFinding(edit(filename, "RECOVERY=NONE\r\n",
                                        "RECOVERY=REQUIRED; reason=unsafe host state\r\n"),
                                   f"{filename}:")
                self.assertFinding(edit(filename, "RECOVERY=NONE\r\n", ""),
                                   "expected RECOVERY=NONE")
                self.assertFinding(edit(filename, "cleanup=PASS; recovery=NONE",
                                        "cleanup=PASS; recovery=REQUIRED"), "RESULT=")

    def test_failure_paths_must_end_as_documented(self) -> None:
        self.assertFinding(edit("injected.log", "failures=1; cleanup", "failures=2; cleanup"),
                           "injected.log")
        self.assertFinding(edit("skipped.log", "completeness=INCOMPLETE", "completeness=COMPLETE"),
                           "skipped.log")
        self.assertFinding(edit("skipped.log", "FAILURE=harness.completeness", "FAILURE=other"),
                           "FAILURE=other, expected FAILURE=harness.completeness")
        self.assertFinding(edit("skipped.log", "status=PASS", "status=FAIL; x"),
                           "skipped.log:4: ASSERT status=FAIL, expected PASS")

    def test_the_injected_failure_must_be_the_designated_assertion(self) -> None:
        designated = suite().injected[0]
        other = next(item for item in suite().assertions if item != designated)
        def moved(folder: Path) -> None:
            edit("injected.log", f"ASSERT={designated}; status=FAIL; expected=1; actual=0.75",
                 f"ASSERT={designated}; status=PASS")(folder)
            edit("injected.log", f"ASSERT={other}; status=PASS",
                 f"ASSERT={other}; status=FAIL; expected=1; actual=0.75")(folder)
            edit("injected.log", f"FAILURE={designated}", f"FAILURE={other}")(folder)
        self.assertFinding(moved, f"FAILURE={other}, expected FAILURE={designated}")

    def test_environment(self) -> None:
        self.assertFinding(edit("injected.log", "build=17928", "build=18000"),
                           "different hosts")
        self.assertFinding(edit("examples.log", "build=17928; ", ""), "lacks build")
        def bitness(folder: Path) -> None:
            for name in check_excel_run.LOGS.values():
                edit(name, "office=64-bit", "office=32-bit")(folder)
        self.assertFinding(bitness, "not the supported target")

    def test_host_record(self) -> None:
        self.assertFinding(lambda folder: (folder / "host.txt").unlink(), "host.txt: missing")
        self.assertFinding(edit("host.txt", "compile=PASS", "compile=FAIL"), "compile=FAIL")
        self.assertFinding(edit("host.txt", "compile=PASS\n", ""), "no compile=")
        self.assertFinding(edit("host.txt", "OLE Automation; ", "OLE Automation; Scripting; "),
                           "references must be exactly")
        self.assertFinding(edit("host.txt", "windows=Windows 11 Pro, version 24H2, OS build 26100.4061\n",
                                ""), "no windows=")

    def test_a_missing_host_record_is_written_on_windows(self) -> None:
        build = "Windows 11 Enterprise, version 25H2, OS build 26200.9711"
        with evidence(lambda folder: (folder / "host.txt").unlink()) as folder:
            result = check_excel_run.check(ROOT, COMMIT, folder, lambda: build)
            self.assertEqual(result.findings, [])
            self.assertEqual((folder / "host.txt").read_text(encoding="utf-8").splitlines(),
                             [f"windows={build}",
                              "references=Visual Basic For Applications; Microsoft Excel 16.0 "
                              "Object Library; OLE Automation; Microsoft Office 16.0 Object Library",
                              "compile=PASS"])
            self.assertIn("host.txt: written by this check", check_excel_run.report(COMMIT, result))

    def test_an_existing_host_record_is_kept(self) -> None:
        with evidence() as folder:
            result = check_excel_run.check(ROOT, COMMIT, folder, lambda: "Windows 11 Home")
            self.assertEqual(result.findings, [])
            self.assertFalse(result.host_written)
            self.assertEqual((folder / "host.txt").read_text(encoding="utf-8"), HOST)

    def test_local_windows_reads_the_registry(self) -> None:
        values: dict[str, object] = {"ProductName": "Windows 10 Enterprise", "DisplayVersion": "25H2",
                                     "CurrentBuild": "26200", "UBR": 9711}

        def query(key: str, name: str) -> tuple[object, int]:
            if name not in values:
                raise FileNotFoundError(name)
            return values[name], 1

        registry = types.SimpleNamespace(HKEY_LOCAL_MACHINE=0, QueryValueEx=query,
                                         OpenKey=lambda *args: contextlib.nullcontext("key"))
        with mock.patch.object(sys, "platform", "win32"), mock.patch.dict(sys.modules, {"winreg": registry}):
            self.assertEqual(check_excel_run.local_windows(),
                             "Windows 11 Enterprise, version 25H2, OS build 26200.9711")
            del values["DisplayVersion"]
            self.assertIsNone(check_excel_run.local_windows())
        with mock.patch.object(sys, "platform", "linux"):
            self.assertIsNone(check_excel_run.local_windows())

    def test_windows_name_matches_winver(self) -> None:
        name = check_excel_run.windows_name
        self.assertEqual(name("Windows 10 Enterprise", "25H2", "26200", "9711"),
                         "Windows 11 Enterprise, version 25H2, OS build 26200.9711")
        self.assertEqual(name("Windows 10 Pro", "22H2", "19045", "5011"),
                         "Windows 10 Pro, version 22H2, OS build 19045.5011")


class SuiteTest(unittest.TestCase):
    def test_suite_matches_the_documented_counts(self) -> None:
        examples = suite()
        self.assertEqual(len(examples.cases), examples.expected_cases)
        self.assertEqual(len(examples.assertions), examples.expected_assertions)
        self.assertEqual(examples.last_case[0], examples.cases[-1])
        self.assertEqual(len(examples.injected), 1)

    def test_the_entry_point_saves_the_logs_the_checker_reads(self) -> None:
        source = (ROOT / EXAMPLES).read_text(encoding="ascii")
        saved = dict(re.findall(r'SaveRun "(\w+)", Folder & "\\([\w.]+)"', source))
        self.assertEqual(saved, check_excel_run.LOGS)


if __name__ == "__main__":
    unittest.main()
