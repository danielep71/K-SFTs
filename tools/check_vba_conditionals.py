#!/usr/bin/env python3
"""Check VBA conditional compilation and Declare statements in every modeled host.

Evaluates every ``#If`` block of the tracked VBA components in each modeled
compilation environment and rejects what it cannot evaluate: an unknown
symbol, a project ``#Const``, an unbalanced block. A ``Declare`` reachable
under VBA7 must be ``PtrSafe`` (docs/VBA_HOUSE_STYLE.md, Module rules). The
reachable-source model is shared with check_vba_jumps.py and
check_vba_public_api.py.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path

from vba_lexer import (
    Finding,
    logical_units,
    missing,
    read_component,
    report,
    tracked_components,
)

ROOT = Path(__file__).resolve().parents[1]
DECLARE = re.compile(
    r"^\s*(?:(?:Public|Private|Global)\s+)?Declare\s+(?P<ptrsafe>PtrSafe\s+)?(?:Function|Sub)\b",
    re.IGNORECASE)
DIRECTIVE = re.compile(r"^\s*#\s*(If|ElseIf|Else|End\s+If)\b(.*)$", re.IGNORECASE)
CONST = re.compile(r"^\s*#\s*Const\b", re.IGNORECASE)
TOKEN = re.compile(r"\s*(\(|\)|<>|=|-?\d+|[A-Za-z_]\w*)")

# Microsoft's compiler constants: VBA6 means version-6 compatible, so it is
# True under VBA7 as well; Win32 is True on 32-bit and 64-bit development
# platforms, and Win64 on 64-bit Office as well. 64-bit Excel (vba7-win64)
# is the supported target (INSTALLATION.md); vba7-win32 and the pre-VBA7
# vba6-win32 are modeled so that a branch written for them is still
# evaluated and its declarations checked.
ENVIRONMENTS: dict[str, dict[str, bool]] = {
    "vba6-win32": {"VBA6": True, "VBA7": False, "WIN32": True, "WIN64": False},
    "vba7-win32": {"VBA6": True, "VBA7": True, "WIN32": True, "WIN64": False},
    "vba7-win64": {"VBA6": True, "VBA7": True, "WIN32": True, "WIN64": True},
}
SYMBOLS = frozenset(ENVIRONMENTS["vba7-win64"])
VBA7_ENVIRONMENTS = tuple(name for name, symbols in ENVIRONMENTS.items() if symbols["VBA7"])


class ExpressionError(ValueError):
    """A conditional expression outside the supported model."""


class ExpressionParser:
    """Boolean expressions over VBA6, VBA7, Win32, Win64, True, False, -1 and 0."""

    def __init__(self, expression: str, symbols: dict[str, bool]) -> None:
        self.symbols = symbols
        self.tokens = self.tokenize(expression)
        self.index = 0

    @staticmethod
    def tokenize(expression: str) -> list[str]:
        tokens: list[str] = []
        position = 0
        while position < len(expression):
            match = TOKEN.match(expression, position)
            if match is None:
                if not expression[position:].strip():
                    break
                raise ExpressionError(f"unsupported token near {expression[position:]!r}")
            tokens.append(match.group(1))
            position = match.end()
        if not tokens:
            raise ExpressionError("empty conditional expression")
        return tokens

    def peek(self) -> str:
        return self.tokens[self.index] if self.index < len(self.tokens) else ""

    def take(self) -> str:
        token = self.peek()
        if not token:
            raise ExpressionError("unexpected end of expression")
        self.index += 1
        return token

    def parse(self) -> bool:
        value = self.parse_or()
        if self.peek():
            raise ExpressionError(f"unexpected token {self.peek()!r}")
        return value

    def parse_or(self) -> bool:
        value = self.parse_and()
        while self.peek().casefold() == "or":
            self.take()
            right = self.parse_and()
            value = value or right
        return value

    def parse_and(self) -> bool:
        value = self.parse_not()
        while self.peek().casefold() == "and":
            self.take()
            right = self.parse_not()
            value = value and right
        return value

    def parse_not(self) -> bool:
        if self.peek().casefold() == "not":
            self.take()
            return not self.parse_not()
        left = self.parse_atom()
        operator = self.peek()
        if operator not in {"=", "<>"}:
            return left
        self.take()
        right = self.parse_atom()
        return left == right if operator == "=" else left != right

    def parse_atom(self) -> bool:
        token = self.take()
        if token == "(":
            value = self.parse_or()
            if self.take() != ")":
                raise ExpressionError("missing closing parenthesis")
            return value
        if token.casefold() in {"true", "false"}:
            return token.casefold() == "true"
        if re.fullmatch(r"-?\d+", token):
            if int(token) not in {-1, 0}:
                raise ExpressionError("only the Boolean integer literals -1 and 0 are supported")
            return int(token) == -1
        if token.upper() not in SYMBOLS:
            raise ExpressionError(f"unsupported conditional-compilation symbol {token!r}; "
                                  "supported symbols are VBA6, VBA7, Win32 and Win64")
        return self.symbols[token.upper()]


def evaluate(expression: str, symbols: dict[str, bool]) -> bool:
    return ExpressionParser(expression, symbols).parse()


@dataclass
class Frame:
    """One open #If block in one environment."""

    parent_active: bool
    branch_taken: bool
    current_active: bool
    else_seen: bool = False


Stacks = dict[str, list[Frame]]


def active(stack: list[Frame]) -> bool:
    return all(frame.current_active for frame in stack)


def condition_values(path: str, line: int, label: str, remainder: str,
                     findings: list[Finding]) -> dict[str, bool]:
    """The condition's value per environment; False everywhere when indeterminate."""
    match = re.match(r"^(.*?)\s+Then\s*$", remainder.strip(), re.IGNORECASE)
    try:
        if match is None or not match.group(1).strip():
            raise ExpressionError(f"#{label} must end with Then and contain a condition")
        return {name: evaluate(match.group(1), symbols)
                for name, symbols in ENVIRONMENTS.items()}
    except ExpressionError as error:
        findings.append(Finding(path, line, f"indeterminate #{label} directive: {error}"))
        return {name: False for name in ENVIRONMENTS}


def handle_if(path: str, line: int, remainder: str, stacks: Stacks,
              findings: list[Finding]) -> None:
    values = condition_values(path, line, "If", remainder, findings)
    for name, stack in stacks.items():
        selected = active(stack) and values[name]
        stack.append(Frame(active(stack), selected, selected))


def handle_elseif(path: str, line: int, remainder: str, stacks: Stacks,
                  findings: list[Finding]) -> None:
    if any(not stack for stack in stacks.values()):
        findings.append(Finding(path, line, "#ElseIf without #If"))
        return
    if any(stack[-1].else_seen for stack in stacks.values()):
        findings.append(Finding(path, line, "#ElseIf after #Else"))
        return
    values = condition_values(path, line, "ElseIf", remainder, findings)
    for name, stack in stacks.items():
        frame = stack[-1]
        selected = frame.parent_active and not frame.branch_taken and values[name]
        frame.current_active = selected
        frame.branch_taken = frame.branch_taken or selected


def handle_else(path: str, line: int, remainder: str, stacks: Stacks,
                findings: list[Finding]) -> None:
    if remainder.strip():
        findings.append(Finding(path, line, "#Else must not contain trailing text"))
    elif any(not stack for stack in stacks.values()):
        findings.append(Finding(path, line, "#Else without #If"))
    elif any(stack[-1].else_seen for stack in stacks.values()):
        findings.append(Finding(path, line, "duplicate #Else"))
    else:
        for stack in stacks.values():
            frame = stack[-1]
            frame.current_active = frame.parent_active and not frame.branch_taken
            frame.branch_taken = True
            frame.else_seen = True


def handle_end_if(path: str, line: int, remainder: str, stacks: Stacks,
                  findings: list[Finding]) -> None:
    if remainder.strip():
        findings.append(Finding(path, line, "#End If must not contain trailing text"))
    elif any(not stack for stack in stacks.values()):
        findings.append(Finding(path, line, "#End If without #If"))
    else:
        for stack in stacks.values():
            stack.pop()


HANDLERS: dict[str, Callable[[str, int, str, Stacks, list[Finding]], None]] = {
    "if": handle_if, "elseif": handle_elseif, "else": handle_else, "end if": handle_end_if,
}


def handle_directive(path: str, line: int, code: str, stacks: Stacks,
                     findings: list[Finding]) -> None:
    if CONST.match(code):
        findings.append(Finding(path, line, "a project #Const symbol is outside the supported "
                                            "model; conditional compilation fails closed"))
        return
    directive = DIRECTIVE.match(code)
    if directive is None:
        findings.append(Finding(path, line, f"unsupported conditional-compilation directive: "
                                            f"{code}"))
        return
    kind = " ".join(directive.group(1).split()).casefold()
    HANDLERS[kind](path, line, directive.group(2), stacks, findings)


def unclosed(path: str, stacks: Stacks) -> list[Finding]:
    depth = max(len(stack) for stack in stacks.values())
    return [Finding(path, None, f"{depth} conditional-compilation block(s) not closed")] \
        if depth else []


def reachable_sources(path: str, text: str) -> tuple[dict[str, str], list[Finding]]:
    """Each environment's source with inactive lines and directives blanked,
    so that line numbers are preserved; plus the directive findings."""
    lines = text.split("\n")
    sources = {name: [""] * len(lines) for name in ENVIRONMENTS}
    stacks: Stacks = {name: [] for name in ENVIRONMENTS}
    findings: list[Finding] = []
    for start, end, kind, code in logical_units(lines):
        if kind == "directive":
            handle_directive(path, start, code, stacks, findings)
            continue
        for name, stack in stacks.items():
            if active(stack):
                sources[name][start - 1:end] = lines[start - 1:end]
    return {name: "\n".join(source) for name, source in sources.items()}, \
        findings + unclosed(path, stacks)


def analyze(path: str, text: str) -> list[Finding]:
    findings: list[Finding] = []
    stacks: Stacks = {name: [] for name in ENVIRONMENTS}
    for start, _, kind, code in logical_units(text.split("\n")):
        if kind == "directive":
            handle_directive(path, start, code, stacks, findings)
            continue
        declaration = DECLARE.match(code)
        reachable = [name for name in VBA7_ENVIRONMENTS if active(stacks[name])]
        if declaration and reachable and not declaration.group("ptrsafe"):
            findings.append(Finding(path, start, "Declare reachable under VBA7 without PtrSafe "
                                                 f"in {', '.join(reachable)}"))
    return findings + unclosed(path, stacks)


def check(root: Path) -> tuple[list[Finding], int]:
    components = tracked_components(root)
    findings: list[Finding] = []
    for path in components:
        if absent := missing(root, path):
            findings.append(absent)
            continue
        findings.extend(analyze(path, read_component(root, path)))
    return findings, len(components)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components = check(args.root.resolve())
    return report("vba-conditionals", findings, f"{components} VBA component(s), "
                  f"environments {', '.join(ENVIRONMENTS)}")


if __name__ == "__main__":
    sys.exit(main())
