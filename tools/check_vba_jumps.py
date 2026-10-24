#!/usr/bin/env python3
"""Check that every VBA jump target is a label of the procedure that jumps.

``GoTo``, ``GoSub``, ``Resume <label>`` and ``On Error GoTo <label>`` must
name a label defined in the same procedure, labels must be unique within a
procedure, and every procedure must be closed. Each compilation environment
of check_vba_conditionals.py is checked separately, so that mutually
exclusive ``#If`` branches are never combined. Adapted from K-SACCR; see
NOTICE.md.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path

from check_vba_conditionals import reachable_sources
from vba_lexer import Finding, logical_units, read_component, report, tracked_components

ROOT = Path(__file__).resolve().parents[1]
PROC_OPEN = re.compile(
    r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\b", re.IGNORECASE)
PROC_CLOSE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
NAMED_LABEL = re.compile(r"^\s*([A-Za-z_]\w*|\d+):")
NUMBERED_LABEL = re.compile(r"^\s*(\d+)(?=\s+\S)")
JUMP = re.compile(r"\b(GoTo|GoSub|Resume)\b(?:\s+([A-Za-z_]\w*|\d+|-1))?", re.IGNORECASE)
ON_ERROR = re.compile(r"\bon\s+error\s*$", re.IGNORECASE)
# Statements that can stand alone before a colon, as in "Else: x = 1"; a
# keyword is never a label.
KEYWORDS = frozenset({"beep", "do", "else", "end", "loop", "next", "resume", "return", "stop",
                      "wend"})


@dataclass
class Procedure:
    name: str
    line: int
    labels: dict[str, int] = field(default_factory=dict)
    jumps: list[tuple[int, str, str]] = field(default_factory=list)


def is_procedure(code: str) -> re.Match[str] | None:
    match = PROC_OPEN.match(code)
    return match if match and " declare " not in f" {code.casefold()} " else None


def label_at_start(code: str) -> str | None:
    match = NAMED_LABEL.match(code) or NUMBERED_LABEL.match(code)
    return match.group(1) if match and match.group(1).casefold() not in KEYWORDS else None


def statement_jumps(code: str) -> list[tuple[str, str]]:
    """``(operation, target)`` for each jump to a label in one statement.

    ``Resume``, ``Resume Next`` and ``On Error GoTo 0`` or ``-1`` name no label.
    """
    jumps: list[tuple[str, str]] = []
    for match in JUMP.finditer(code):
        operation, target = match.group(1), match.group(2)
        if target is None or (operation.casefold() == "resume" and target.casefold() == "next"):
            continue
        if operation.casefold() == "goto" and target in {"0", "-1"} \
                and ON_ERROR.search(code[:match.start()]):
            continue
        jumps.append((operation, target))
    return jumps


def scan(path: str, text: str) -> tuple[list[Procedure], list[Finding]]:
    """Collect each procedure's labels and jumps; report structural defects."""
    procedures: list[Procedure] = []
    findings: list[Finding] = []
    current: Procedure | None = None
    for line, _, _, code in logical_units(text.split("\n"), blank_strings=True):
        opener = is_procedure(code)
        if opener:
            if current is not None:
                findings.append(Finding(path, line, f"procedure {current.name} opened at line "
                                                    f"{current.line} is not closed before "
                                                    f"{opener.group(2)}"))
            current = Procedure(opener.group(2), line)
            procedures.append(current)
        elif PROC_CLOSE.match(code):
            if current is None:
                findings.append(Finding(path, line, "End statement outside any procedure"))
            current = None
        elif current is not None:
            label = label_at_start(code)
            if label is not None and label.casefold() in current.labels:
                findings.append(Finding(path, line, f"duplicate label {label!r} in procedure "
                                                    f"{current.name}; first defined at line "
                                                    f"{current.labels[label.casefold()]}"))
            elif label is not None:
                current.labels[label.casefold()] = line
            current.jumps.extend((line, operation, target)
                                 for operation, target in statement_jumps(code))
    if current is not None:
        findings.append(Finding(path, current.line, f"procedure {current.name} is not closed"))
    return procedures, findings


def analyze_source(path: str, text: str) -> list[Finding]:
    procedures, findings = scan(path, text)
    for procedure in procedures:
        findings.extend(Finding(path, line, f"{operation} target {target!r} is not a label of "
                                            f"procedure {procedure.name}")
                        for line, operation, target in procedure.jumps
                        if target.casefold() not in procedure.labels)
    return findings


def analyze(path: str, text: str) -> list[Finding]:
    """Findings of every environment, each listed once with the environments it occurs in."""
    sources, findings = reachable_sources(path, text)
    environments: dict[Finding, list[str]] = {}
    for environment, source in sources.items():
        for finding in analyze_source(path, source):
            environments.setdefault(finding, []).append(environment)
    findings.extend(Finding(item.path, item.line, f"{item.message} [{', '.join(names)}]")
                    for item, names in environments.items())
    return findings


def check(root: Path) -> tuple[list[Finding], int, int]:
    components = tracked_components(root)
    findings: list[Finding] = []
    procedures = 0
    for path in components:
        text = read_component(root, path)
        findings.extend(analyze(path, text))
        procedures += sum(bool(is_procedure(code)) for _, _, _, code
                          in logical_units(text.split("\n"), blank_strings=True))
    return findings, len(components), procedures


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components, procedures = check(args.root.resolve())
    return report("vba-jumps", findings,
                  f"{components} VBA component(s), {procedures} procedure(s)")


if __name__ == "__main__":
    sys.exit(main())
