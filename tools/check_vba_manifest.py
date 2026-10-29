#!/usr/bin/env python3
"""Check tools/vba-projects.json against the tracked VBA components.

The manifest lists the components that tools/Sync-KSFT-VBA.ps1 imports into
each workbook built from the template. This keeps it from going stale:

- every entry is a tracked component, listed once, in sorted order;
- document modules are the ThisWorkbook and sh<Noun> classes of
  src/workbook/, and every one of them is listed;
- production lists every other component under src/ and nothing else;
- development lists the production components and every component of
  tests/modules/, and nothing else.

Examples are never imported. Static: no workbook is opened.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Any

from vba_lexer import Finding, report, tracked_components

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = "tools/vba-projects.json"


def is_document(path: str) -> bool:
    name = Path(path).stem
    return (path.startswith("src/workbook/") and path.endswith(".cls")
            and (name == "ThisWorkbook" or name.startswith("sh")))


def entries(data: dict[str, Any], key: str, findings: list[Finding]) -> list[str]:
    """One list of the manifest: strings, each once, in sorted order."""
    if key == "document_modules":
        value = data.get(key)
    else:
        configurations = data.get("configurations")
        configuration = configurations.get(key) if isinstance(configurations, dict) else None
        value = configuration.get("modules") if isinstance(configuration, dict) else None
    if not isinstance(value, list) or not all(isinstance(item, str) for item in value):
        findings.append(Finding(MANIFEST, None, f"{key}: needs a list of paths"))
        return []
    if len(set(value)) != len(value):
        findings.append(Finding(MANIFEST, None, f"{key}: a path is listed twice"))
    if value != sorted(value):
        findings.append(Finding(MANIFEST, None, f"{key}: paths must be in sorted order"))
    return value


def compare(key: str, listed: list[str], expected: set[str], findings: list[Finding]) -> None:
    for path in sorted(expected - set(listed)):
        findings.append(Finding(MANIFEST, None, f"{key}: {path} is missing"))
    for path in sorted(set(listed) - expected):
        findings.append(Finding(MANIFEST, None, f"{key}: {path} does not belong here "
                                                "(not a tracked component of this workbook)"))


def check(root: Path) -> tuple[list[Finding], int]:
    path = root / MANIFEST
    if not path.is_file():
        return [Finding(MANIFEST, None, "missing")], 0
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, ValueError) as error:
        return [Finding(MANIFEST, None, f"not valid UTF-8 JSON ({error})")], 0
    if not isinstance(data, dict) or data.get("schema_version") != 1:
        return [Finding(MANIFEST, None, "needs an object with schema_version 1")], 0
    findings: list[Finding] = []
    components = set(tracked_components(root))
    source = {item for item in components if item.startswith("src/")}
    documents = {item for item in source if is_document(item)}
    production = source - documents
    tests = {item for item in components if item.startswith("tests/modules/")}
    compare("document_modules", entries(data, "document_modules", findings), documents, findings)
    compare("production", entries(data, "production", findings), production, findings)
    compare("development", entries(data, "development", findings), production | tests, findings)
    return findings, len(components)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components = check(args.root.resolve())
    return report("vba-manifest", findings, f"{components} VBA component(s) against {MANIFEST}")


if __name__ == "__main__":
    sys.exit(main())
