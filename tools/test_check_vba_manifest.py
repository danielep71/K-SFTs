"""Valid and invalid fixtures for check_vba_manifest.py.

Each fixture is a throw-away Git repository (test_check_vba.repository) with
synthetic components and a tools/vba-projects.json; no workbook is involved.
"""

from __future__ import annotations

import contextlib
import io
import json
import unittest
from collections.abc import Mapping
from pathlib import Path
from typing import Any

import check_vba_manifest
from test_check_vba import CORE, FACADE, HARNESS, VALID, WORKBOOK, repository

ROOT = Path(__file__).resolve().parents[1]
FILES: dict[str, str | bytes] = {
    "src/core/CORE_Validation.bas": CORE,
    "src/modules/SFT_Functions.bas": FACADE,
    "src/workbook/ThisWorkbook.cls": WORKBOOK,
    "tests/modules/TEST_Harness.bas": HARNESS,
    "examples/modules/ExampleSingleNettingSet.bas": VALID["examples/modules/ExampleSingleNettingSet.bas"],
}
PRODUCTION = ["src/core/CORE_Validation.bas", "src/modules/SFT_Functions.bas"]


def manifest(**changes: Any) -> dict[str, Any]:
    data: dict[str, Any] = {
        "schema_version": 1,
        "document_modules": ["src/workbook/ThisWorkbook.cls"],
        "configurations": {
            "production": {"modules": list(PRODUCTION)},
            "development": {"modules": [*PRODUCTION, "tests/modules/TEST_Harness.bas"]},
        },
    }
    data.update(changes)
    return data


def findings(data: Any, files: Mapping[str, str | bytes] = FILES) -> list[str]:
    text = data if isinstance(data, str) else json.dumps(data)
    with repository({**files, check_vba_manifest.MANIFEST: text}) as root:
        return [str(finding) for finding in check_vba_manifest.check(root)[0]]


class ManifestTest(unittest.TestCase):
    def assertFinding(self, data: Any, text: str,
                      files: Mapping[str, str | bytes] = FILES) -> None:
        found = findings(data, files)
        self.assertTrue(any(text in item for item in found), found)

    def test_valid_manifest_passes(self) -> None:
        self.assertEqual(findings(manifest()), [])

    def test_repository_manifest_passes(self) -> None:
        self.assertEqual(check_vba_manifest.check(ROOT)[0], [])

    def test_unreadable_or_wrong_schema(self) -> None:
        self.assertFinding("{", "not valid UTF-8 JSON")
        self.assertFinding(manifest(schema_version=2), "schema_version 1")
        self.assertFinding(manifest(configurations={"production": []}),
                           "production: needs a list of paths")

    def test_a_component_left_out_or_misplaced(self) -> None:
        data = manifest()
        data["configurations"]["development"]["modules"].remove("tests/modules/TEST_Harness.bas")
        self.assertFinding(data, "development: tests/modules/TEST_Harness.bas is missing")
        data = manifest(document_modules=[])
        self.assertFinding(data, "document_modules: src/workbook/ThisWorkbook.cls is missing")
        data = manifest()
        data["configurations"]["production"]["modules"].append("tests/modules/TEST_Harness.bas")
        self.assertFinding(data, "production: tests/modules/TEST_Harness.bas does not belong")

    def test_examples_and_untracked_paths_are_refused(self) -> None:
        data = manifest()
        data["configurations"]["development"]["modules"].append(
            "examples/modules/ExampleSingleNettingSet.bas")
        self.assertFinding(data, "development: examples/modules/ExampleSingleNettingSet.bas "
                                 "does not belong")
        data = manifest()
        data["configurations"]["production"]["modules"].append("src/core/CORE_Gone.bas")
        self.assertFinding(data, "production: src/core/CORE_Gone.bas does not belong")

    def test_order_and_duplicates(self) -> None:
        data = manifest()
        data["configurations"]["production"]["modules"].reverse()
        self.assertFinding(data, "production: paths must be in sorted order")
        data = manifest()
        data["configurations"]["production"]["modules"].insert(0, PRODUCTION[0])
        self.assertFinding(data, "production: a path is listed twice")

    def test_missing_manifest(self) -> None:
        with repository(FILES) as root:
            self.assertEqual([str(item) for item in check_vba_manifest.check(root)[0]],
                             ["tools/vba-projects.json: missing"])

    def test_main_reports_the_verdict(self) -> None:
        with contextlib.redirect_stdout(io.StringIO()) as output:
            self.assertEqual(check_vba_manifest.main(["--root", str(ROOT)]), 0)
        self.assertIn("vba-manifest: PASS", output.getvalue())


if __name__ == "__main__":
    unittest.main()
