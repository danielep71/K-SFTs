"""Static consistency of the VBA harness and its example suite (tests/modules/).

The harness runs only in Excel. These tests check, without Excel, what its
source can show: that the example suite's expected counts equal the cases and
assertions it contains, that every case runs once with a well-formed unique
identifier, and that the documented passing RESULT line matches the counts.
"""

from __future__ import annotations

import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = ROOT / "tests/modules/TEST_Examples.bas"
HARNESS = ROOT / "tests/modules/TEST_Harness.bas"
CASE_ID = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")


def source(path: Path) -> str:
    return path.read_text(encoding="ascii").replace("\r\n", "\n")


def constant(text: str, name: str) -> int:
    match = re.search(rf"^\s*Private Const {name}\s+As Long = (\d+)\b", text, re.M)
    assert match, name
    return int(match.group(1))


def procedures(text: str) -> dict[str, str]:
    """Body of every Sub and Function, by name."""
    pattern = re.compile(r"^(?:Public|Private) (?:Sub|Function) (\w+)\(.*?^End (?:Sub|Function)$",
                         re.M | re.S)
    return {match.group(1): match.group(0) for match in pattern.finditer(text)}


class ExampleSuiteTest(unittest.TestCase):
    text = source(EXAMPLES)
    bodies = procedures(text)
    cases = {name: body for name, body in bodies.items() if name.startswith("Case")}

    def test_runner_calls_every_case_once_in_order(self) -> None:
        runner = self.bodies["RunSuite"]
        called = re.findall(r"^\s+(Case\w+)$", runner, re.M)
        self.assertEqual(sorted(called), sorted(self.cases))
        self.assertEqual(len(called), len(set(called)))

    def test_each_case_begins_once_with_a_unique_identifier(self) -> None:
        identifiers = []
        for name, body in self.cases.items():
            begun = re.findall(r'TEST_Harness\.BeginCase "([^"]+)"', body)
            self.assertEqual(len(begun), 1, name)
            self.assertRegex(begun[0], CASE_ID)
            self.assertIn("TEST_Harness.RecordCaseError", body, name)
            identifiers += begun
        self.assertEqual(len(identifiers), len(set(identifiers)))

    def test_expected_counts_match_the_source(self) -> None:
        assertions = sum(len(re.findall(r"TEST_Harness\.Assert\w+ ", body))
                         for body in self.cases.values())
        self.assertEqual(constant(self.text, "EXPECTED_CASES"), len(self.cases))
        self.assertEqual(constant(self.text, "EXPECTED_ASSERTIONS"), assertions)

    def test_documented_assertion_counts_per_case(self) -> None:
        for name, body in self.cases.items():
            stated = re.search(r"(\d+) assertions?\.", body)
            self.assertIsNotNone(stated, name)
            assert stated is not None
            self.assertEqual(int(stated.group(1)),
                             len(re.findall(r"TEST_Harness\.Assert\w+ ", body)), name)

    def test_documented_passing_line(self) -> None:
        cases = constant(self.text, "EXPECTED_CASES")
        assertions = constant(self.text, "EXPECTED_ASSERTIONS")
        line = ("RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE; "
                f"cases={cases}; assertions={assertions}; failures=0; cleanup=PASS; "
                "recovery=NONE")
        self.assertIn(line, (ROOT / "INSTALLATION.md").read_text(encoding="utf-8"))
        skipped = self.cases["CaseCleanupOnlyFailure"]
        remaining = assertions - len(re.findall(r"TEST_Harness\.Assert\w+ ", skipped))
        self.assertIn(f"`cases={cases - 1}`, `assertions={remaining}`",
                      (ROOT / "INSTALLATION.md").read_text(encoding="utf-8"))


class HarnessTest(unittest.TestCase):
    text = source(HARNESS)

    def test_every_documented_record_is_written(self) -> None:
        readme = (ROOT / "tests/README.md").read_text(encoding="utf-8")
        table = readme.split('<a id="log-format"></a>', 1)[1]
        cells = [line.split("|")[1] for line in table.splitlines() if line.startswith("| `")]
        documented = {key for cell in cells for key in re.findall(r"`([A-Z_]+)=", cell)}
        written = set(re.findall(r'"([A-Z_]+)=', self.text))
        self.assertEqual(documented, written - {"RESET"})
        self.assertIn("RESET=DONE", readme)

    def test_documented_failure_ids_are_written(self) -> None:
        for failure in ("harness.case-id", "harness.no-case", "harness.completeness",
                        "runner.unexpected-error", ".unexpected-error"):
            with self.subTest(failure=failure):
                self.assertIn(f'"{failure}"', self.text)


if __name__ == "__main__":
    unittest.main()
