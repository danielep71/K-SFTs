"""Tests for the test-case format check (tools/check_test_cases.py)."""

from __future__ import annotations

import contextlib
import copy
import io
import json
import tempfile
import unittest
from pathlib import Path
from typing import Any

import check_test_cases

ROOT = Path(__file__).resolve().parents[1]
REGISTER = """# Methodology

## 📚 Source register

| ID | Source |
| --- | --- |
| `CRR` | CRR |
| `BCBS` | CRE22 |
| `IMPL-EXAMPLE` | An external implementation |

## 📅 Version applicability

| Regime | Source ID | From | To |
| --- | --- | --- | --- |
| `CRR` | `CRR` | 2025-01-01 | |
| `BCBS` | `BCBS` | 2023-01-01 | 2028-10-31 |

## 🧭 Assumptions

| `NOT-A-SOURCE` | outside the register section |
"""


def fixture(case_id: str = "illustrative-case") -> dict[str, Any]:
    return {
        "schema_version": 1,
        "kind": "sft-fixture",
        "id": case_id,
        "synthetic": True,
        "description": "A test fixture",
        "category": "nominal",
        "calculation_currency_code": "EUR",
        "valuation_date": "2026-01-02",
        "netting_sets": [{
            "id": "NS1",
            "counterparty_id": "CP1",
            "inputs": {"master_netting_agreement_flag": True},
            "transactions": [{
                "id": "T1",
                "product": "repo",
                "inputs": {"currency_code": "EUR", "cash_amount": 100, "start_date": "2026-01-02",
                           "holding_period_business_days": 5, "haircut_rate": 0.02},
            }],
            "collateral": [{"id": "C1", "transaction_id": "T1",
                            "inputs": {"currency_code": "USD", "market_value_amount": 102}}],
        }],
    }


def reference(kind: str = "illustrative") -> dict[str, Any]:
    return {"class": kind, "source": None, "locator": None, "derivation": "Why it exists",
            "script": None, "derived_by": "Tester", "reviewed_by": None, "date": "2026-10-24"}


def number_output(**changes: Any) -> dict[str, Any]:
    output = {"quantity": "illustrative_total", "netting_set": "NS1", "transaction": None,
              "rule": None, "value": 100, "unit": "EUR",
              "tolerance": {"absolute": 1e-6, "relative": 1e-9}, "reference": reference()}
    output.update(changes)
    return output


def expected(case_id: str = "illustrative-case", regime: str = "CRR",
             *outputs: dict[str, Any]) -> dict[str, Any]:
    return {"schema_version": 1, "kind": "sft-expected", "fixture": case_id, "regime": regime,
            "outputs": list(outputs) or [number_output()]}


class Workspace:
    """A temporary repository root with a register and test-case files."""

    def __init__(self, root: Path) -> None:
        self.root = root
        (root / "docs/methodology").mkdir(parents=True)
        (root / "docs/methodology/README.md").write_text(REGISTER, encoding="utf-8")
        (root / "tests/fixtures").mkdir(parents=True)
        (root / "tests/expected").mkdir(parents=True)

    def write(self, relative: str, data: object) -> None:
        text = data if isinstance(data, str) else json.dumps(data, indent=2)
        (self.root / relative).write_text(text, encoding="utf-8")

    def case(self, fixture_data: dict[str, Any], *expected_files: tuple[str, dict[str, Any]]) -> None:
        case_id = fixture_data["id"]
        self.write(f"tests/fixtures/{case_id}.json", fixture_data)
        for regime, data in expected_files or (("crr", expected(case_id)),):
            self.write(f"tests/expected/{case_id}.{regime}.json", data)

    def findings(self) -> list[str]:
        return check_test_cases.check(self.root)[0]


class CheckTestCasesTests(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.ws = Workspace(Path(directory.name))

    def assertFinding(self, text: str) -> None:
        findings = self.ws.findings()
        self.assertTrue(any(text in f for f in findings), f"{text!r} not in {findings}")

    def test_valid_case_passes(self) -> None:
        self.ws.case(fixture())
        self.assertEqual(self.ws.findings(), [])

    def test_empty_directories_pass(self) -> None:
        self.assertEqual(self.ws.findings(), [])

    def test_repository_cases_pass(self) -> None:
        findings, fixtures, expected_files = check_test_cases.check(ROOT)
        self.assertEqual(findings, [])
        self.assertGreaterEqual(fixtures, 1)
        self.assertGreaterEqual(expected_files, 1)

    def test_register_ids_come_from_the_register_section_only(self) -> None:
        self.assertEqual(check_test_cases.register_ids(self.ws.root),
                         {"CRR", "BCBS", "IMPL-EXAMPLE"})

    def test_invalid_json_and_duplicate_keys(self) -> None:
        self.ws.write("tests/fixtures/broken.json", "{not json")
        self.ws.write("tests/fixtures/dup.json", '{"id": "a", "id": "b"}')
        findings = self.ws.findings()
        self.assertTrue(any("broken.json: not valid UTF-8 JSON" in f for f in findings))
        self.assertTrue(any("duplicate key 'id'" in f for f in findings))

    def test_null_documents_are_reported_not_skipped(self) -> None:
        self.ws.case(fixture())
        self.ws.write("tests/fixtures/extra-case.json", "null")
        self.ws.write("tests/expected/illustrative-case.bcbs.json", "null")
        findings = self.ws.findings()
        self.assertIn("tests/fixtures/extra-case.json: is not an object", findings)
        self.assertIn("tests/expected/illustrative-case.bcbs.json: is not an object", findings)

    def test_non_finite_numbers_are_rejected_in_every_category(self) -> None:
        for category in ("nominal", "boundary", "invalid"):
            with self.subTest(category=category):
                data = fixture()
                data["category"] = category
                text = json.dumps(data).replace('"cash_amount": 100', '"cash_amount": NaN')
                self.assertIn("NaN", text)
                self.ws.write("tests/fixtures/illustrative-case.json", text)
                self.assertFinding("not valid UTF-8 JSON (non-standard number NaN)")
        self.ws.write("tests/fixtures/illustrative-case.json", fixture())
        self.ws.write("tests/expected/illustrative-case.crr.json",
                      json.dumps(expected()).replace('"value": 100', '"value": -Infinity'))
        self.assertFinding("not valid UTF-8 JSON (non-standard number -Infinity)")

    def test_overflowing_numbers_are_rejected_in_every_category(self) -> None:
        for category in ("nominal", "boundary", "invalid"):
            with self.subTest(category=category):
                data = fixture()
                data["category"] = category
                text = json.dumps(data).replace('"cash_amount": 100', '"cash_amount": 1e400')
                self.ws.write("tests/fixtures/illustrative-case.json", text)
                self.assertFinding("input 'cash_amount' is not a finite number")

    def test_envelope_rules(self) -> None:
        data = fixture()
        data.update(synthetic=False, category="other", calculation_currency_code="eur",
                    valuation_date="2026-02-30")
        self.ws.case(data)
        for text in ("synthetic is not true", "unknown category", "not a currency code",
                     "valuation_date is not an ISO date"):
            self.assertFinding(text)

    def test_unknown_and_missing_keys(self) -> None:
        data = fixture()
        data["extra"] = 1
        del data["description"]
        self.ws.case(data)
        self.assertFinding("missing description; unknown extra")

    def test_id_must_equal_file_name(self) -> None:
        data = fixture()
        self.ws.write("tests/fixtures/illustrative-other.json", data)
        self.ws.write("tests/expected/illustrative-other.crr.json",
                      expected("illustrative-other"))
        self.assertFinding("id is not the kebab-case file name")

    def test_input_suffixes_and_values(self) -> None:
        data = fixture()
        inputs = data["netting_sets"][0]["transactions"][0]["inputs"]
        inputs.update(cash=1, start_date="02/01/2026", holding_period_business_days=2.5,
                      haircut_rate=True, Bad_Key=1)
        self.ws.case(data)
        for text in ("input 'cash' has no unit suffix", "'start_date' has an invalid value",
                     "'holding_period_business_days' has an invalid value",
                     "'haircut_rate' has an invalid value", "'Bad_Key' is not lower snake case"):
            self.assertFinding(text)

    def test_amount_needs_a_currency(self) -> None:
        data = fixture()
        del data["netting_sets"][0]["collateral"][0]["inputs"]["currency_code"]
        self.ws.case(data)
        self.assertFinding("inputs hold an amount but no currency_code")

    def test_identifiers_products_and_collateral_links(self) -> None:
        data = fixture()
        netting_set = data["netting_sets"][0]
        netting_set["transactions"][0]["product"] = "swap"
        netting_set["collateral"][0]["transaction_id"] = "T9"
        netting_set["collateral"].append(copy.deepcopy(netting_set["collateral"][0]))
        self.ws.case(data)
        for text in ("unknown product 'swap'", "transaction_id 'T9' is not in this netting set",
                     "id 'C1' is not unique in the fixture"):
            self.assertFinding(text)

    def test_expected_file_name_and_envelope(self) -> None:
        self.ws.case(fixture(), ("crr", expected(regime="BCBS")))
        self.ws.write("tests/expected/illustrative-case.eu.json", expected())
        self.ws.write("tests/expected/missing-case.crr.json", expected("missing-case"))
        self.assertFinding("envelope does not match")
        self.assertFinding("illustrative-case.eu.json: is not named")
        self.assertFinding("missing-case.crr.json: is not named")

    def test_fixture_without_expected_file(self) -> None:
        self.ws.write("tests/fixtures/illustrative-case.json", fixture())
        self.assertFinding("has no expected file")

    def test_output_scope(self) -> None:
        self.ws.case(fixture(), ("crr", expected("illustrative-case", "CRR",
                                                 number_output(netting_set="NS9"),
                                                 number_output(transaction="T9"))))
        self.assertFinding("netting_set 'NS9' is not in the fixture")
        self.assertFinding("transaction 'T9' is not in that netting set")

    def test_number_and_text_forms(self) -> None:
        text_with_unit = number_output(text="x")
        del text_with_unit["value"]
        self.ws.case(fixture(), ("crr", expected(
            "illustrative-case", "CRR",
            number_output(value="100"), number_output(unit="EUR/USD"),
            number_output(tolerance={"absolute": -1, "relative": 0}),
            number_output(tolerance={"absolute": 0}), text_with_unit)))
        for text in ("value is not a finite number", "unit 'EUR/USD'",
                     "bounds must be finite numbers of at least 0",
                     "must hold exactly absolute and relative", "unknown tolerance, unit"):
            self.assertFinding(text)

    def test_rules_follow_the_regime(self) -> None:
        good = number_output(rule="CRE22.65")
        crr_rule = number_output(rule="CRR.220.3")
        self.ws.case(fixture(), ("bcbs", expected("illustrative-case", "BCBS", good)),
                     ("crr", expected("illustrative-case", "CRR", good, crr_rule)))
        findings = self.ws.findings()
        self.assertEqual([f for f in findings if "rule" in f],
                         ["tests/expected/illustrative-case.crr.json: outputs[0]: rule 'CRE22.65' "
                          "is not a CRR rule ID"])

    def test_only_illustrative_values_may_lack_a_rule(self) -> None:
        published = reference("published")
        published.update(source="BCBS", locator="CRE22.65")
        self.ws.case(fixture("case"), ("crr", expected(
            "case", "CRR", number_output(quantity="total", reference=published))))
        self.assertFinding("only an illustrative value may have no rule")

    def test_reference_classes(self) -> None:
        published = reference("published")
        published.update(source="IMPL-EXAMPLE", locator="x")
        independent = reference("independent")
        independent.update(source="CRR", locator="x")
        unregistered = reference()
        unregistered["source"] = "NOT-A-SOURCE"
        no_derivation = reference()
        no_derivation["derivation"] = " "
        outputs = [number_output(quantity="total", rule="CRR.220", reference=r)
                   for r in (published, independent, unregistered, no_derivation)]
        self.ws.case(fixture("case"), ("crr", expected("case", "CRR", *outputs)))
        for text in ("a published value needs a register source (not IMPL-)",
                     "an independent value needs an IMPL- source",
                     "source 'NOT-A-SOURCE' is not in the register",
                     "an illustrative value needs a derivation"):
            self.assertFinding(text)

    def test_script_must_exist_under_derivations(self) -> None:
        (self.ws.root / "tests/derivations").mkdir()
        (self.ws.root / "tests/derivations/total.py").write_text("x = 1\n", encoding="utf-8")
        present, missing, outside = reference(), reference(), reference()
        present["script"] = "tests/derivations/total.py"
        missing["script"] = "tests/derivations/absent.py"
        outside["script"] = "tools/check.py"
        self.ws.case(fixture(), ("crr", expected("illustrative-case", "CRR", *[
            number_output(reference=r) for r in (present, missing, outside)])))
        findings = [f for f in self.ws.findings() if "script" in f]
        self.assertEqual(len(findings), 2)
        self.assertTrue(all("absent.py" in f or "tools/check.py" in f for f in findings))

    def test_illustrative_prefix_rule(self) -> None:
        published = reference("published")
        published.update(source="BCBS", locator="CRE22.65")
        self.ws.case(fixture(), ("bcbs", expected(
            "illustrative-case", "BCBS",
            number_output(quantity="total", rule="CRE22.65", reference=published))))
        self.ws.case(fixture("plain-case"), ("crr", expected(
            "plain-case", "CRR", number_output(quantity="total", rule="CRR.220"))))
        findings = [f for f in self.ws.findings() if "prefix must be used" in f]
        self.assertEqual(len(findings), 2)

    def test_illustrative_quantity_outside_illustrative_case(self) -> None:
        published = reference("published")
        published.update(source="CRR", locator="Article 220(3)")
        self.ws.case(fixture("plain-case"), ("crr", expected(
            "plain-case", "CRR", number_output(rule="CRR.220.3", reference=published))))
        self.assertFinding("an illustrative_ quantity outside an illustrative- case")

    def test_invalid_case_needs_a_status(self) -> None:
        data = fixture()
        data["category"] = "invalid"
        status = number_output(quantity="transaction_status", transaction="T1", text="REJECTED")
        for key in ("value", "unit", "tolerance"):
            del status[key]
        self.ws.case(data)
        self.assertFinding("an invalid case needs a _status text output")
        self.ws.case(data, ("crr", expected("illustrative-case", "CRR", status)))
        self.assertEqual(self.ws.findings(), [])

    def test_invalid_case_may_hold_out_of_domain_values(self) -> None:
        data = fixture()
        data["category"] = "invalid"
        netting_set = data["netting_sets"][0]
        netting_set["transactions"][0]["product"] = "swap"
        netting_set["transactions"][0]["inputs"].update(currency_code="eur", start_date="soon",
                                                        haircut_rate="high")
        del netting_set["collateral"][0]["inputs"]["currency_code"]
        status = number_output(quantity="transaction_status", transaction="T1", text="REJECTED")
        for key in ("value", "unit", "tolerance"):
            del status[key]
        self.ws.case(data, ("crr", expected("illustrative-case", "CRR", status)))
        self.assertEqual(self.ws.findings(), [])
        data["category"] = "nominal"
        self.ws.case(data, ("crr", expected("illustrative-case", "CRR", status)))
        for text in ("unknown product 'swap'", "'currency_code' has an invalid value",
                     "'start_date' has an invalid value", "'haircut_rate' has an invalid value",
                     "inputs hold an amount but no currency_code"):
            self.assertFinding(text)

    def test_invalid_case_still_checks_structure(self) -> None:
        data = fixture()
        data["category"] = "invalid"
        netting_set = data["netting_sets"][0]
        netting_set["transactions"][0]["product"] = 7
        netting_set["transactions"][0]["inputs"]["cash"] = 1
        netting_set["collateral"][0]["transaction_id"] = "T9"
        self.ws.case(data)
        for text in ("unknown product 7", "input 'cash' has no unit suffix",
                     "transaction_id 'T9' is not in this netting set",
                     "an invalid case needs a _status text output"):
            self.assertFinding(text)

    def test_methodology_tables_are_read(self) -> None:
        report = check_test_cases.Report()
        method = check_test_cases.read_methodology(self.ws.root, report)
        self.assertEqual(report.findings, [])
        self.assertEqual(method.windows, {"CRR": [("2025-01-01", None)],
                                          "BCBS": [("2023-01-01", "2028-10-31")]})

    def test_valuation_date_must_fall_in_a_registered_version(self) -> None:
        data = fixture()
        data["valuation_date"] = "2024-12-31"
        self.ws.case(data, ("crr", expected()), ("bcbs", expected(regime="BCBS")))
        findings = [f for f in self.ws.findings() if "outside every registered" in f]
        self.assertEqual(len(findings), 1)
        self.assertIn("illustrative-case.crr.json: valuation_date 2024-12-31", findings[0])
        data["valuation_date"] = "2028-11-01"
        self.ws.case(data, ("bcbs", expected(regime="BCBS")))
        self.assertFinding("valuation_date 2028-11-01 is outside every registered BCBS version")

    def test_regime_must_be_in_the_applicability_table(self) -> None:
        register = self.ws.root / "docs/methodology/README.md"
        register.write_text(register.read_text(encoding="utf-8").replace(
            "| `BCBS` | `BCBS` | 2023-01-01 | 2028-10-31 |\n", ""), encoding="utf-8")
        self.ws.case(fixture(), ("bcbs", expected(regime="BCBS")))
        self.ws.write("tests/expected/illustrative-case.CRR.json", expected())
        findings = self.ws.findings()
        self.assertTrue(any("illustrative-case.bcbs.json: is not named <case-id>.<crr>.json"
                            in f for f in findings), findings)
        self.assertTrue(any("illustrative-case.CRR.json: is not named" in f for f in findings))

    def test_malformed_applicability_rows_are_reported(self) -> None:
        register = self.ws.root / "docs/methodology/README.md"
        register.write_text(register.read_text(encoding="utf-8").replace(
            "| `BCBS` | `BCBS` | 2023-01-01 | 2028-10-31 |",
            "| `BCBS` | `NOT-REGISTERED` | 2023-01-01 | 2022-01-01 |\n| `XYZ` | `CRR` | 2025-01-01 | |"
            "\n| `CRR` | `CRR` |"), encoding="utf-8")
        findings = self.ws.findings()
        self.assertEqual(len([f for f in findings if "applicability row" in f]), 3, findings)

    def test_main_reports_and_exit_code(self) -> None:
        self.ws.case(fixture())
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            self.assertEqual(check_test_cases.main(["--root", str(self.ws.root)]), 0)
        self.assertIn("test-cases: PASS (1 fixture(s), 1 expected file(s)", out.getvalue())
        self.ws.write("tests/fixtures/illustrative-case.json", "[]")
        out = io.StringIO()
        with contextlib.redirect_stdout(out):
            self.assertEqual(check_test_cases.main(["--root", str(self.ws.root)]), 1)
        self.assertIn("FAIL: tests/fixtures/illustrative-case.json: is not an object",
                      out.getvalue())


if __name__ == "__main__":
    unittest.main()
