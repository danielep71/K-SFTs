#!/usr/bin/env python3
"""Check the placement, names, encoding, storage and mechanical style of VBA source.

Applies the mechanical rules of docs/VBA_HOUSE_STYLE.md and the placement
rules of docs/REPOSITORY_STRUCTURE.md to every tracked file: VBA components
only in their homes, each with its role's name prefix and a VB_Name equal to
its file name; ASCII only; Option Explicit, and Option Private Module where
required; no Option Base 1, tab or trailing white space; string literals
closed; forms beside their .frx; LF storage with CRLF checkout; no workbook
committed but the approved template, which must hold no macros.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
import zipfile
from collections.abc import Sequence
from pathlib import Path

from vba_lexer import Finding, missing, read_component, report, tracked_files

ROOT = Path(__file__).resolve().parents[1]
COMPONENT_SUFFIXES = frozenset({".bas", ".cls", ".frm"})
# Component name pattern by (home, suffix): docs/VBA_HOUSE_STYLE.md, Naming.
# Any other suffix in a home, or a component outside these homes, is misplaced.
NAMING: dict[tuple[str, str], tuple[str, re.Pattern[str]]] = {
    ("src/core/", ".bas"): ("CORE_<Subject>", re.compile(r"CORE_[A-Z][A-Za-z0-9_]*")),
    ("src/modules/", ".bas"): ("SFT_<Subject>", re.compile(r"SFT_[A-Z][A-Za-z0-9_]*")),
    ("src/classes/", ".cls"): ("C<Noun>", re.compile(r"C[A-Z][A-Za-z0-9]*")),
    ("src/workbook/", ".cls"): ("ThisWorkbook or sh<Noun>",
                                re.compile(r"ThisWorkbook|sh[A-Z][A-Za-z0-9]*")),
    ("src/workbook/", ".bas"): ("M_<Subject>", re.compile(r"M_[A-Z][A-Za-z0-9_]*")),
    ("src/forms/", ".frm"): ("frm<Noun>", re.compile(r"frm[A-Z][A-Za-z0-9]*")),
    ("tests/modules/", ".bas"): ("TEST_<Subject>", re.compile(r"TEST_[A-Z][A-Za-z0-9_]*")),
    ("examples/modules/", ".bas"): ("Example<Subject>",
                                    re.compile(r"Example[A-Z][A-Za-z0-9]*")),
}
HOMES = tuple(sorted({home for home, _ in NAMING}))
# Standard modules whose Public procedures must stay inside the project.
PRIVATE_MODULE_HOMES = ("src/core/", "tests/modules/")
MAX_NAME = 31  # the VBE rejects a longer component name
# Workbooks are build outputs. One may be committed only at an exact path
# approved in an issue and re-included in .gitignore: today only the
# macro-free template of #102 (docs/REPOSITORY_STRUCTURE.md, Placement rules).
WORKBOOK_SUFFIXES = frozenset({".xls", ".xlsx", ".xlsm", ".xlsb", ".xlt", ".xltx", ".xltm",
                               ".xlam", ".xla", ".xlw"})
APPROVED_WORKBOOKS = frozenset({"src/workbook/SFT_Template.xlsx"})

VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]*)"[ \t]*$', re.M)
OPTION_EXPLICIT = re.compile(r"^[ \t]*Option[ \t]+Explicit[ \t]*(?:'.*)?$", re.M | re.I)
OPTION_PRIVATE = re.compile(r"^[ \t]*Option[ \t]+Private[ \t]+Module[ \t]*(?:'.*)?$",
                            re.M | re.I)
OPTION_BASE_ONE = re.compile(r"^[ \t]*Option[ \t]+Base[ \t]+1\b", re.M | re.I)
FRX_REFERENCE = re.compile(r'"([^"\n]+\.frx)":([0-9A-Fa-f]+)')
REM_STATEMENT = re.compile(r"Rem(?=[ \t]|$)", re.I)
LINE_NUMBER = re.compile(r"[ \t]*\d+(?=[ \t])")


def placement(path: str) -> list[Finding]:
    """A component, or a form's .frx, must sit in its home with its role's name."""
    suffix = Path(path).suffix.lower()
    if suffix == ".frx":
        return [] if path.startswith("src/forms/") and path.count("/") == 2 else \
            [Finding(path, None, "a form resource belongs beside its form in src/forms/")]
    home = next((home for home in HOMES if path.startswith(home)), None)
    if home is None or "/" in path[len(home):]:
        return [Finding(path, None, f"VBA component outside its homes ({', '.join(HOMES)}, "
                                    "no subdirectories)")]
    rule = NAMING.get((home, suffix))
    if rule is None:
        return [Finding(path, None, f"{home} holds no {suffix} components")]
    name = Path(path).stem
    findings = []
    if not rule[1].fullmatch(name):
        findings.append(Finding(path, None, f"a {suffix} component in {home} is named "
                                            f"{rule[0]}"))
    if len(name) > MAX_NAME:
        findings.append(Finding(path, None, f"component name longer than {MAX_NAME} characters"))
    return findings


def non_ascii(path: str, data: bytes) -> list[Finding]:
    if data.startswith(b"\xef\xbb\xbf"):
        return [Finding(path, 1, "byte-order mark; VBA source is ASCII only")]
    return [Finding(path, number, f"non-ASCII byte 0x{byte:02X}; VBA source is ASCII only, "
                                  "build the character with ChrW")
            for number, line in enumerate(data.split(b"\n"), 1)
            for byte in sorted(set(line)) if byte > 0x7F][:20]


def unclosed_literals(text: str) -> list[int]:
    """Lines whose code leaves a string literal open.

    A ' outside a literal and outside [...] (which quotes a sheet name, as in
    ['Sheet 1'!A1]), or Rem at the start of a statement, also after a line
    number, begins a comment; a comment ending in " _" continues on the next
    line.
    """
    open_lines: list[int] = []
    in_comment = False
    for number, line in enumerate(text.split("\n"), 1):
        if in_comment:
            in_comment = line.rstrip().endswith(" _")
            continue
        quoted, statement_start, brackets = False, True, 0
        label = LINE_NUMBER.match(line)
        for index in range(label.end() if label else 0, len(line)):
            char = line[index]
            if quoted:
                quoted = char != '"'
            elif char == '"':
                quoted = True
            elif (char == "'" and not brackets) or \
                    (statement_start and REM_STATEMENT.match(line, index)):
                in_comment = line.rstrip().endswith(" _")
                break
            elif char == "[":
                brackets += 1
            elif char == "]" and brackets:
                brackets -= 1
            statement_start = (char == ":" and not quoted) or \
                (statement_start and char.isspace())
        if quoted:
            open_lines.append(number)
    return open_lines


def layout(path: str, text: str) -> list[Finding]:
    """Tabs and trailing white space; docs/VBA_HOUSE_STYLE.md, Indentation."""
    findings = []
    for number, line in enumerate(text.split("\n"), 1):
        if "\t" in line:
            findings.append(Finding(path, number, "tab character; indent with spaces"))
        if line != line.rstrip(" \t"):
            findings.append(Finding(path, number, "trailing white space"))
    return findings


def form_resources(root: Path, path: str, text: str, tracked: set[str]) -> list[Finding]:
    references = FRX_REFERENCE.findall(text)
    if not references:
        return [Finding(path, None, "no form resource (.frx) reference")]
    findings = []
    for filename, offset in references:
        companion = f"{Path(path).parent.as_posix()}/{filename}"
        if Path(filename).name != filename or companion not in tracked:
            findings.append(Finding(path, None, f"form resource {filename} is not tracked "
                                                "beside the form"))
        elif absent := missing(root, companion):
            findings.append(absent)
        elif (root / companion).stat().st_size <= int(offset, 16):
            findings.append(Finding(path, None, f"resource offset is outside {filename}"))
    return findings


def content(root: Path, path: str, tracked: set[str], names: dict[str, str]) -> list[Finding]:
    data = (root / path).read_bytes()
    text = read_component(root, path)
    findings = non_ascii(path, data)
    vb_names = VB_NAME.findall(text)
    stem = Path(path).stem
    if vb_names != [stem]:
        findings.append(Finding(path, None, f'needs exactly one Attribute VB_Name = "{stem}"'))
    elif stem.casefold() in names:
        findings.append(Finding(path, None, f"component name also used by {names[stem.casefold()]}"
                                            " (VBA names are case-insensitive)"))
    else:
        names[stem.casefold()] = path
    suffix = Path(path).suffix.lower()
    if suffix == ".bas" and not text.startswith('Attribute VB_Name = "'):
        findings.append(Finding(path, 1, "a .bas export starts with its Attribute VB_Name line"))
    if not OPTION_EXPLICIT.search(text):
        findings.append(Finding(path, None, "missing Option Explicit"))
    if suffix == ".bas" and path.startswith(PRIVATE_MODULE_HOMES) and \
            not OPTION_PRIVATE.search(text):
        findings.append(Finding(path, None, "missing Option Private Module"))
    findings.extend(Finding(path, text[:match.start()].count("\n") + 1,
                            "Option Base 1; declare explicit array bounds")
                    for match in OPTION_BASE_ONE.finditer(text))
    findings.extend(Finding(path, number, "string literal not closed")
                    for number in unclosed_literals(text))
    findings.extend(layout(path, text))
    if suffix == ".frm":
        findings.extend(form_resources(root, path, text, tracked))
    return findings


def storage(root: Path, paths: list[str]) -> list[Finding]:
    """Exported VBA is stored with LF and checked out with CRLF (.gitattributes)."""
    if not paths:
        return []
    completed = subprocess.run(["git", "-C", str(root), "ls-files", "--eol", "-z", "--", *paths],
                               capture_output=True, check=False)
    if completed.returncode:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    findings = []
    for record in completed.stdout.split(b"\0"):
        if not record:
            continue
        info, _, path = record.decode("utf-8", errors="surrogateescape").partition("\t")
        fields = info.split()
        index_eol, attributes = fields[0].removeprefix("i/"), " ".join(fields[2:])
        if index_eol not in {"lf", "none", ""}:
            findings.append(Finding(path, None, f"stored with {index_eol} in Git; store with LF"))
        if "eol=crlf" not in attributes:
            findings.append(Finding(path, None, ".gitattributes must check it out with CRLF"))
    return findings


def template_problems(root: Path, path: str) -> list[Finding]:
    """The approved template is a macro-free Office Open XML workbook."""
    if absent := missing(root, path):
        return [absent]
    try:
        with zipfile.ZipFile(root / path) as package:
            names = package.namelist()
    except zipfile.BadZipFile:
        return [Finding(path, None, "the template is not an Office Open XML package")]
    if "xl/workbook.xml" not in names:
        return [Finding(path, None, "the template holds no xl/workbook.xml")]
    if any(Path(name).name.lower() == "vbaproject.bin" for name in names):
        return [Finding(path, None, "the template holds a VBA project; it must be macro-free")]
    return []


def check(root: Path) -> tuple[list[Finding], int]:
    tracked = tracked_files(root)
    tracked_set = set(tracked)
    components = [p for p in tracked if Path(p).suffix.lower() in COMPONENT_SUFFIXES]
    findings: list[Finding] = []
    for path in tracked:
        suffix = Path(path).suffix.lower()
        if suffix in COMPONENT_SUFFIXES or suffix == ".frx":
            findings.extend(placement(path))
        if suffix in WORKBOOK_SUFFIXES and path not in APPROVED_WORKBOOKS:
            findings.append(Finding(path, None, "workbook committed; workbooks are built from "
                                                "source and only the template path is approved"))
        elif path in APPROVED_WORKBOOKS:
            findings.extend(template_problems(root, path))
    names: dict[str, str] = {}
    for path in components:
        if absent := missing(root, path):
            findings.append(absent)
            continue
        findings.extend(content(root, path, tracked_set, names))
    findings.extend(storage(root, components))
    return findings, len(components)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, components = check(args.root.resolve())
    return report("vba-source", findings, f"{components} VBA component(s)")


if __name__ == "__main__":
    sys.exit(main())
