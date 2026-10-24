#!/usr/bin/env python3
"""Check the numerical test-case files against docs/methodology/TEST_CASES.md.

Static: it checks the format of tests/fixtures/*.json and tests/expected/*.json
and the references between them. It never computes or compares a value.
"""

from __future__ import annotations

import argparse
import datetime
import json
import math
import re
import sys
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
REGISTER = Path("docs/methodology/README.md")
FIXTURES = Path("tests/fixtures")
EXPECTED = Path("tests/expected")
DERIVATIONS = "tests/derivations/"
# Until this directory holds the input and output contract, only illustrative- cases may exist.
CONTRACTS = Path("docs/contracts")

CASE_ID = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
SNAKE = re.compile(r"^[a-z][a-z0-9]*(?:_[a-z0-9]+)*$")
CURRENCY = re.compile(r"^[A-Z]{3}$")
DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
REGISTER_ROW = re.compile(r"^\|\s*`([^`]+)`\s*\|")
# Rule-ID pattern of each regime that can be declared. A regime is in use only
# when the methodology's Version applicability table lists it (today: CRR).
RULES = {"CRR": re.compile(r"^CRR\.\d+[a-z]*(?:\.\d+)?$"), "BCBS": re.compile(r"^CRE22\.\d+$")}
PRODUCTS = frozenset({"repo", "reverse_repo", "securities_lending", "securities_borrowing",
                      "margin_lending"})
CATEGORIES = frozenset({"nominal", "boundary", "invalid"})
CLASSES = frozenset({"published", "independent", "illustrative"})
UNITS = frozenset({"1", "years", "business_days"})
SUFFIXES = ("_business_days", "_amount", "_factor", "_years", "_count", "_rate", "_date",
            "_flag", "_code", "_id")

FIXTURE_KEYS = frozenset({"schema_version", "kind", "id", "synthetic", "description",
                          "category", "calculation_currency_code", "valuation_date",
                          "netting_sets"})
NETTING_SET_KEYS = frozenset({"id", "counterparty_id", "inputs", "transactions", "collateral"})
TRANSACTION_KEYS = frozenset({"id", "product", "inputs"})
COLLATERAL_KEYS = frozenset({"id", "transaction_id", "inputs"})
EXPECTED_KEYS = frozenset({"schema_version", "kind", "fixture", "regime", "outputs"})
OUTPUT_KEYS = frozenset({"quantity", "netting_set", "transaction", "rule", "reference"})
NUMBER_KEYS = OUTPUT_KEYS | {"value", "unit", "tolerance"}
TEXT_KEYS = OUTPUT_KEYS | {"text"}
REFERENCE_KEYS = frozenset({"class", "source", "locator", "derivation", "script",
                            "derived_by", "reviewed_by", "date"})


@dataclass
class Case:
    """What an expected file needs to know about its fixture."""

    category: str
    valuation_date: str
    netting_sets: dict[str, set[str]]
    classes: set[str] = field(default_factory=set)
    expected_files: int = 0


@dataclass
class Methodology:
    """What the methodology README declares: register IDs and version windows."""

    sources: set[str]
    windows: dict[str, list[tuple[str, str | None]]]


@dataclass
class Report:
    findings: list[str] = field(default_factory=list)

    def add(self, where: str, message: str) -> None:
        self.findings.append(f"{where}: {message}")


def is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def is_text(value: object) -> bool:
    return isinstance(value, str) and value.strip() != ""


def is_date(value: object) -> bool:
    if not isinstance(value, str) or not DATE.match(value):
        return False
    try:
        datetime.date.fromisoformat(value)
    except ValueError:
        return False
    return True


def keys_problem(data: Mapping[str, Any], allowed: frozenset[str]) -> str | None:
    missing = sorted(allowed - data.keys())
    unknown = sorted(data.keys() - allowed)
    if missing or unknown:
        parts = ([f"missing {', '.join(missing)}"] if missing else []) + (
            [f"unknown {', '.join(unknown)}"] if unknown else [])
        return "; ".join(parts)
    return None


# Returned by load_json when a file cannot be read; distinct from a JSON null.
LOAD_FAILED = object()


def load_json(path: Path, root: Path, report: Report) -> Any:
    """Parse strict JSON: duplicate keys, the non-standard NaN and Infinity, and any
    number that overflows to infinity (such as 1e400), anywhere in the file, fail."""

    def no_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"duplicate key {key!r}")
            result[key] = value
        return result

    def no_constants(name: str) -> None:
        raise ValueError(f"non-standard number {name}")

    def finite_float(text: str) -> float:
        value = float(text)
        if not math.isfinite(value):
            raise ValueError(f"number {text} overflows to infinity")
        return value

    def finite_int(text: str) -> int:
        value = int(text)
        try:
            float(value)
        except OverflowError:
            raise ValueError(f"number {text} overflows to infinity") from None
        return value

    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=no_duplicates,
                          parse_constant=no_constants, parse_float=finite_float,
                          parse_int=finite_int)
    except (OSError, UnicodeDecodeError, ValueError) as error:
        report.add(rel(path, root), f"not valid UTF-8 JSON ({error})")
        return LOAD_FAILED


def rel(path: Path, root: Path) -> str:
    return path.relative_to(root).as_posix()


def register_ids(root: Path) -> set[str]:
    """Source IDs from the register tables of the methodology README."""
    ids: set[str] = set()
    in_register = False
    for line in (root / REGISTER).read_text(encoding="utf-8").splitlines():
        if line.startswith("## "):
            in_register = "Source register" in line
        elif in_register and (match := REGISTER_ROW.match(line)):
            ids.add(match.group(1))
    return ids


def section_rows(root: Path, title: str) -> list[list[str]]:
    """Cells of the table rows that start with a code span, in one README section."""
    rows = []
    inside = False
    for line in (root / REGISTER).read_text(encoding="utf-8").splitlines():
        if line.startswith("## "):
            inside = title in line
        elif inside and REGISTER_ROW.match(line):
            rows.append([cell.strip().strip("`") for cell in line.strip().strip("|").split("|")])
    return rows


def read_methodology(root: Path, report: Report) -> Methodology:
    """Read the register IDs and the version applicability table."""
    sources = register_ids(root)
    windows: dict[str, list[tuple[str, str | None]]] = {}
    for cells in section_rows(root, "Version applicability"):
        if len(cells) != 4:
            report.add(REGISTER.as_posix(), f"applicability row {cells} does not have 4 cells")
            continue
        regime, source, start, end = cells
        if (source not in sources or not RULES.get(regime) or not is_date(start)
                or (end and not is_date(end)) or (end and end < start)):
            report.add(REGISTER.as_posix(),
                       f"applicability row for {regime!r} needs a known regime, a register "
                       "source and valid dates")
            continue
        windows.setdefault(regime, []).append((start, end or None))
    return Methodology(sources, windows)


def input_value_problem(key: str, value: object, invalid: bool) -> str | None:
    """Check a key's unit suffix and, outside an invalid case, its value's domain."""
    suffix = next((s for s in SUFFIXES if key.endswith(s)), None)
    if suffix is None:
        return f"input {key!r} has no unit suffix"
    if value is None or invalid:
        return None
    if suffix in ("_amount", "_rate", "_factor", "_years"):
        valid = is_number(value)
    elif suffix in ("_business_days", "_count"):
        valid = isinstance(value, int) and not isinstance(value, bool) and value >= 0
    elif suffix == "_date":
        valid = is_date(value)
    elif suffix == "_flag":
        valid = isinstance(value, bool)
    elif key.endswith("currency_code"):
        valid = isinstance(value, str) and bool(CURRENCY.match(value))
    else:
        valid = is_text(value)
    return None if valid else f"input {key!r} has an invalid value for its suffix {suffix}"


def check_inputs(inputs: object, where: str, invalid: bool, report: Report) -> None:
    if not isinstance(inputs, dict):
        report.add(where, "inputs is not an object")
        return
    for key, value in inputs.items():
        if not SNAKE.match(key):
            report.add(where, f"input {key!r} is not lower snake case")
            continue
        if problem := input_value_problem(key, value, invalid):
            report.add(where, problem)
    if (not invalid and any(k.endswith("_amount") for k in inputs)
            and inputs.get("currency_code") is None):
        report.add(where, "inputs hold an amount but no currency_code")


def check_items(items: object, allowed: frozenset[str], where: str, seen: set[str],
                invalid: bool, report: Report) -> list[dict[str, Any]]:
    """Check a list of transactions or collateral items; return the valid objects."""
    if not isinstance(items, list):
        report.add(where, "is not a list")
        return []
    valid = []
    for index, item in enumerate(items):
        at = f"{where}[{index}]"
        if not isinstance(item, dict):
            report.add(at, "is not an object")
            continue
        if problem := keys_problem(item, allowed):
            report.add(at, problem)
            continue
        if not is_text(item["id"]):
            report.add(at, "id is not a text")
        elif item["id"] in seen:
            report.add(at, f"id {item['id']!r} is not unique in the fixture")
        else:
            seen.add(item["id"])
        check_inputs(item["inputs"], at, invalid, report)
        valid.append(item)
    return valid


def check_netting_set(item: dict[str, Any], at: str, seen: set[str], invalid: bool,
                      report: Report) -> set[str]:
    """Check one netting set; return its transaction IDs.

    In an invalid case the structure is still checked, but not the domain of the
    values: an unknown product or an out-of-domain input is what such a case tests.
    """
    if not is_text(item["counterparty_id"]):
        report.add(at, "counterparty_id is not a text")
    check_inputs(item["inputs"], at, invalid, report)
    transactions = check_items(item["transactions"], TRANSACTION_KEYS, f"{at}.transactions",
                               seen, invalid, report)
    ids = {str(t["id"]) for t in transactions}
    for index, transaction in enumerate(transactions):
        product = transaction["product"]
        if not is_text(product) or (not invalid and product not in PRODUCTS):
            report.add(f"{at}.transactions[{index}]",
                       f"unknown product {transaction['product']!r}")
    collateral = check_items(item["collateral"], COLLATERAL_KEYS, f"{at}.collateral", seen,
                             invalid, report)
    for index, entry in enumerate(collateral):
        if entry["transaction_id"] is not None and entry["transaction_id"] not in ids:
            report.add(f"{at}.collateral[{index}]",
                       f"transaction_id {entry['transaction_id']!r} is not in this netting set")
    return ids


def envelope_problems(data: dict[str, Any], stem: str) -> list[str]:
    problems = []
    if data["schema_version"] != 1:
        problems.append("schema_version is not 1")
    if data["kind"] != "sft-fixture":
        problems.append("kind is not 'sft-fixture'")
    if data["id"] != stem or not CASE_ID.match(stem):
        problems.append("id is not the kebab-case file name")
    if data["synthetic"] is not True:
        problems.append("synthetic is not true")
    if not is_text(data["description"]):
        problems.append("description is not a text")
    if data["category"] not in CATEGORIES:
        problems.append(f"unknown category {data['category']!r}")
    code = data["calculation_currency_code"]
    if not isinstance(code, str) or not CURRENCY.match(code):
        problems.append("calculation_currency_code is not a currency code")
    if not is_date(data["valuation_date"]):
        problems.append("valuation_date is not an ISO date")
    return problems


def check_fixture(path: Path, root: Path, report: Report) -> Case | None:
    where = rel(path, root)
    data = load_json(path, root, report)
    if data is LOAD_FAILED:
        return None
    if not isinstance(data, dict):
        report.add(where, "is not an object")
        return None
    if problem := keys_problem(data, FIXTURE_KEYS):
        report.add(where, problem)
        return None
    for problem in envelope_problems(data, path.stem):
        report.add(where, problem)
    netting_sets: dict[str, set[str]] = {}
    seen: set[str] = set()
    invalid = data["category"] == "invalid"
    sets = data["netting_sets"]
    if not isinstance(sets, list) or not sets:
        report.add(where, "netting_sets is not a non-empty list")
        sets = []
    for index, item in enumerate(sets):
        at = f"{where}: netting_sets[{index}]"
        if not isinstance(item, dict) or (problem := keys_problem(item, NETTING_SET_KEYS)):
            report.add(at, problem if isinstance(item, dict) and problem else "is not an object")
            continue
        if not is_text(item["id"]) or item["id"] in netting_sets:
            report.add(at, "id is not a unique text")
            continue
        netting_sets[item["id"]] = check_netting_set(item, at, seen, invalid, report)
    return Case(str(data["category"]), str(data["valuation_date"]), netting_sets)


def tolerance_problem(tolerance: object) -> str | None:
    if not isinstance(tolerance, dict) or set(tolerance) != {"absolute", "relative"}:
        return "tolerance must hold exactly absolute and relative"
    if not all(is_number(v) and v >= 0 for v in tolerance.values()):
        return "tolerance bounds must be finite numbers of at least 0"
    return None


def exact_tolerance_problem(output: dict[str, Any]) -> str | None:
    """An exact comparison (both bounds 0) is for whole numbers only."""
    tolerance, value = output["tolerance"], output["value"]
    if tolerance["absolute"] == 0 and tolerance["relative"] == 0 and value != int(value):
        return "an exact tolerance (absolute and relative 0) needs a whole-number value"
    return None


def form_problems(output: dict[str, Any]) -> list[str]:
    if "text" in output:
        if problem := keys_problem(output, TEXT_KEYS):
            return [problem]
        return [] if isinstance(output["text"], str) else ["text is not a text"]
    if problem := keys_problem(output, NUMBER_KEYS):
        return [problem]
    problems = []
    if not is_number(output["value"]):
        problems.append("value is not a finite number")
    unit = output["unit"]
    if not isinstance(unit, str) or not (unit in UNITS or CURRENCY.match(unit)):
        problems.append(f"unit {unit!r} is not a currency code, 1, years or business_days")
    if problem := tolerance_problem(output["tolerance"]):
        problems.append(problem)
    elif not problems and (problem := exact_tolerance_problem(output)):
        problems.append(problem)
    return problems


def reference_problems(reference: object, sources: set[str], root: Path) -> list[str]:
    if not isinstance(reference, dict):
        return ["reference is not an object"]
    if problem := keys_problem(reference, REFERENCE_KEYS):
        return [f"reference: {problem}"]
    kind, source = reference["class"], reference["source"]
    problems = []
    if kind not in CLASSES:
        return [f"unknown reference class {kind!r}"]
    if source is not None and source not in sources:
        problems.append(f"source {source!r} is not in the register")
    if kind == "published" and (source is None or str(source).startswith("IMPL-")
                                or not is_text(reference["locator"])):
        problems.append("a published value needs a register source (not IMPL-) and a locator")
    if kind == "independent" and not (isinstance(source, str) and source.startswith("IMPL-")
                                      and all(is_text(reference[k]) for k in
                                              ("locator", "derivation", "derived_by"))):
        problems.append("an independent value needs an IMPL- source, locator, derivation "
                        "and derived_by")
    if kind == "illustrative" and not is_text(reference["derivation"]):
        problems.append("an illustrative value needs a derivation")
    problems.extend(script_problems(reference["script"], root))
    if not is_date(reference["date"]):
        problems.append("reference date is not an ISO date")
    for key in ("derived_by", "reviewed_by"):
        if reference[key] is not None and not is_text(reference[key]):
            problems.append(f"{key} is not a text or null")
    return problems


def script_problems(script: object, root: Path) -> list[str]:
    if script is None:
        return []
    if (not isinstance(script, str) or not script.startswith(DERIVATIONS) or ".." in script
            or not (root / script).is_file()):
        return [f"script {script!r} is not an existing file under {DERIVATIONS}"]
    return []


def check_output(output: object, case_id: str, case: Case, regime: str,
                 sources: set[str], root: Path) -> list[str]:
    if not isinstance(output, dict):
        return ["output is not an object"]
    problems = form_problems(output)
    if problems:
        return problems
    quantity = output["quantity"]
    if not isinstance(quantity, str) or not SNAKE.match(quantity):
        problems.append(f"quantity {quantity!r} is not lower snake case")
    elif quantity.startswith("illustrative_") and not case_id.startswith("illustrative-"):
        problems.append("an illustrative_ quantity outside an illustrative- case")
    if output["netting_set"] not in case.netting_sets:
        problems.append(f"netting_set {output['netting_set']!r} is not in the fixture")
    elif (output["transaction"] is not None
          and output["transaction"] not in case.netting_sets[output["netting_set"]]):
        problems.append(f"transaction {output['transaction']!r} is not in that netting set")
    reference = output["reference"]
    kind = reference.get("class") if isinstance(reference, dict) else None
    rule = output["rule"]
    if rule is None and kind != "illustrative":
        problems.append("only an illustrative value may have no rule")
    elif rule is not None and not (isinstance(rule, str) and RULES[regime].match(rule)):
        problems.append(f"rule {rule!r} is not a {regime} rule ID")
    problems.extend(reference_problems(reference, sources, root))
    if isinstance(kind, str):
        case.classes.add(kind)
    return problems


def check_expected(path: Path, root: Path, cases: Mapping[str, Case], method: Methodology,
                   report: Report) -> None:
    where = rel(path, root)
    case_id, _, regime_name = path.name.removesuffix(".json").rpartition(".")
    regime = regime_name.upper()
    if not regime_name.islower() or regime not in method.windows or case_id not in cases:
        names = "|".join(sorted(r.lower() for r in method.windows))
        report.add(where, f"is not named <case-id>.<{names}>.json for an existing fixture")
        return
    data = load_json(path, root, report)
    if data is LOAD_FAILED:
        return
    if not isinstance(data, dict):
        report.add(where, "is not an object")
        return
    if problem := keys_problem(data, EXPECTED_KEYS):
        report.add(where, problem)
        return
    if (data["schema_version"] != 1 or data["kind"] != "sft-expected"
            or data["fixture"] != case_id or data["regime"] != regime):
        report.add(where, "envelope does not match: schema_version 1, kind 'sft-expected', "
                          "fixture and regime as in the file name")
    case = cases[case_id]
    case.expected_files += 1
    if not any(start <= case.valuation_date and (end is None or case.valuation_date <= end)
               for start, end in method.windows[regime]):
        report.add(where, f"valuation_date {case.valuation_date} is outside every registered "
                          f"{regime} version (methodology README, Version applicability)")
    outputs = data["outputs"] if isinstance(data["outputs"], list) else []
    if not outputs:
        report.add(where, "outputs is not a non-empty list")
    for index, output in enumerate(outputs):
        for problem in check_output(output, case_id, case, regime, method.sources, root):
            report.add(f"{where}: outputs[{index}]", problem)
    statuses = [o for o in outputs if isinstance(o, dict) and "text" in o
                and str(o.get("quantity", "")).endswith("_status")]
    if case.category == "invalid" and not statuses:
        report.add(where, "an invalid case needs a _status text output")


def check_cases(cases: Mapping[str, Case], report: Report, contract: bool) -> None:
    for case_id, case in sorted(cases.items()):
        where = f"{FIXTURES.as_posix()}/{case_id}.json"
        if not contract and not case_id.startswith("illustrative-"):
            report.add(where, f"only illustrative- cases may exist until the input contract "
                              f"is written in {CONTRACTS.as_posix()}/ (TEST_CASES.md, Fixture "
                              f"format)")
        if case.expected_files == 0:
            report.add(where, "has no expected file")
            continue
        all_illustrative = case.classes == {"illustrative"}
        if all_illustrative != case_id.startswith("illustrative-"):
            report.add(where, "the illustrative- prefix must be used exactly when every "
                              "expected value is illustrative")


def check(root: Path) -> tuple[list[str], int, int]:
    """Return the findings and the numbers of fixtures and expected files checked."""
    report = Report()
    method = read_methodology(root, report)
    fixtures = sorted((root / FIXTURES).glob("*.json"))
    expected = sorted((root / EXPECTED).glob("*.json"))
    cases: dict[str, Case] = {}
    for path in fixtures:
        if (case := check_fixture(path, root, report)) is not None:
            cases[path.stem] = case
    for path in expected:
        check_expected(path, root, cases, method, report)
    contract = (root / CONTRACTS).is_dir() and any((root / CONTRACTS).iterdir())
    check_cases(cases, report, contract)
    return report.findings, len(fixtures), len(expected)


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=(__doc__ or "").splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT,
                        help="repository root (default: this checkout)")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    findings, fixtures, expected = check(args.root.resolve())
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"test-cases: {status} ({fixtures} fixture(s), {expected} expected file(s); "
          "format only, no value computed)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
