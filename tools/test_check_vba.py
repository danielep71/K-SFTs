"""Valid and invalid fixtures for the VBA static checks (tools/check_vba_*.py).

Each fixture is a throw-away Git repository holding synthetic VBA components,
written with CRLF as the VBE exports them and stored under this repository's
.gitattributes (one storage test deliberately omits it). The components are
fixtures for the checks only: no calculation, product or regulatory value is
implied by them.
"""

from __future__ import annotations

import contextlib
import io
import shutil
import subprocess
import tempfile
import unittest
from collections.abc import Callable, Iterator, Mapping
from pathlib import Path

import check_vba_conditionals
import check_vba_jumps
import check_vba_public_api
import check_vba_source
from vba_lexer import Finding, logical_units, strip_code

ROOT = Path(__file__).resolve().parents[1]

CORE = '''Attribute VB_Name = "CORE_Validation"
Option Explicit
Option Private Module

Public Function IsPositive( _
    ByVal Value As Double) _
    As Boolean

        On Error GoTo Failed
        IsPositive = Value > 0
        Exit Function

Failed:
        Err.Raise Err.Number, "CORE_Validation.IsPositive", Err.Description
End Function
'''
FACADE = '''Attribute VB_Name = "SFT_Functions"
Option Explicit

Public Const SFT_FORMAT_TEXT As String = "0.00"


Public Function SFT_IsPositive( _
    ByVal Value As Double) _
    As Variant

        SFT_IsPositive = CORE_Validation.IsPositive(Value)
End Function
'''
INVENTORY = (
    "# Supported VBA declarations\n"
    "SFT_Functions\tConst\tSFT_FORMAT_TEXT\n"
    '# SIG\tSFT_Functions\tConst\tSFT_FORMAT_TEXT\tPublic Const SFT_FORMAT_TEXT As String = "0.00"\n'
    "SFT_Functions\tFunction\tSFT_IsPositive\n"
    "# SIG\tSFT_Functions\tFunction\tSFT_IsPositive\t"
    "Public Function SFT_IsPositive(ByVal Value As Double) As Variant\n"
)
CLASS = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCollateralItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMarketValue As Double  'Synthetic value of the fixture


Public Property Get MarketValue() As Double
        MarketValue = mMarketValue
End Property


Public Property Let MarketValue(ByVal Value As Double)
        mMarketValue = Value
End Property
'''
WORKBOOK = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
'''
FORM = '''VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRunOptions
   Caption         =   "Run options"
   OleObjectBlob   =   "frmRunOptions.frx":0000
End
Attribute VB_Name = "frmRunOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''
HARNESS = '''Attribute VB_Name = "TEST_Harness"
Option Explicit
Option Private Module

Public Sub RunAll()
        Debug.Print "It's ""quoted"": GoTo Nowhere", ['Sheet 1'!A1]
        Rem A comment with an open " quote
End Sub
'''
EXAMPLE = '''Attribute VB_Name = "ExampleSingleNettingSet"
Option Explicit

Public Sub Run()
        Debug.Print SFT_Functions.SFT_IsPositive(1)
End Sub
'''
VALID: dict[str, str | bytes] = {
    "src/core/CORE_Validation.bas": CORE,
    "src/modules/SFT_Functions.bas": FACADE,
    "src/classes/CCollateralItem.cls": CLASS,
    "src/workbook/ThisWorkbook.cls": WORKBOOK,
    "src/forms/frmRunOptions.frm": FORM,
    "src/forms/frmRunOptions.frx": b"\x00\x01\x02\x03",
    "tests/modules/TEST_Harness.bas": HARNESS,
    "examples/modules/ExampleSingleNettingSet.bas": EXAMPLE,
    "docs/PUBLIC_API.txt": INVENTORY,
}


@contextlib.contextmanager
def repository(files: Mapping[str, str | bytes], *, attributes: bool = True) -> Iterator[Path]:
    """A Git repository holding ``files``: VBA components written with CRLF, other
    text as given, then staged."""
    with tempfile.TemporaryDirectory(prefix="vba-check-") as temporary:
        root = Path(temporary)
        subprocess.run(["git", "init", "-q", str(root)], check=True)
        if attributes:
            shutil.copyfile(ROOT / ".gitattributes", root / ".gitattributes")
        for relative, data in files.items():
            path = root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            if isinstance(data, str):
                data = data.replace("\n", "\r\n").encode("ascii", errors="surrogateescape") \
                    if relative.endswith((".bas", ".cls", ".frm")) else data.encode("utf-8")
            path.write_bytes(data)
        subprocess.run(["git", "-C", str(root), "-c", "core.autocrlf=false", "-c",
                        "core.safecrlf=false", "add", "--all"], check=True)
        yield root


def run_check(check: Callable[[Path], tuple[list[Finding], *tuple[int, ...]]],
              files: Mapping[str, str | bytes], *, attributes: bool = True) -> list[str]:
    with repository(files, attributes=attributes) as root:
        return [str(finding) for finding in check(root)[0]]


def changed(path: str, old: str, new: str, files: Mapping[str, str | bytes] = VALID
            ) -> dict[str, str | bytes]:
    text = files[path]
    assert isinstance(text, str) and old in text, (path, old)
    return {**files, path: text.replace(old, new, 1)}


def standalone(source: str) -> dict[str, str | bytes]:
    """One test module holding ``source``, plus an empty inventory."""
    return {"tests/modules/TEST_Fixture.bas": source, "docs/PUBLIC_API.txt": "# None\n"}


class SourceTest(unittest.TestCase):
    def findings(self, files: Mapping[str, str | bytes], **options: bool) -> list[str]:
        return run_check(check_vba_source.check, files, **options)

    def assertFinding(self, files: Mapping[str, str | bytes], text: str) -> None:
        findings = self.findings(files)
        self.assertTrue(any(text in item for item in findings), findings)

    def test_valid_project_passes(self) -> None:
        self.assertEqual(self.findings(VALID), [])

    def test_empty_repository_passes(self) -> None:
        self.assertEqual(self.findings({"README.md": "# Empty\n"}), [])

    def test_component_outside_its_homes(self) -> None:
        self.assertFinding({"src/CORE_Stray.bas": CORE.replace("CORE_Validation", "CORE_Stray")},
                           "outside its homes")
        self.assertFinding({"src/core/sub/CORE_Validation.bas": CORE}, "outside its homes")
        self.assertFinding({"tools/CORE_Validation.bas": CORE}, "outside its homes")

    def test_home_without_that_component_kind(self) -> None:
        self.assertFinding({"src/core/CORE_Validation.cls": CORE}, "holds no .cls components")
        self.assertFinding({"src/modules/frmRunOptions.frm": FORM}, "holds no .frm components")

    def test_role_prefixes(self) -> None:
        cases = {
            "src/core/Validation.bas": "CORE_<Subject>",
            "src/modules/Functions.bas": "SFT_<Subject>",
            "src/classes/CollateralItem.cls": "C<Noun>",
            "src/workbook/Inputs.cls": "ThisWorkbook or sh<Noun>",
            "src/workbook/Main.bas": "M_<Subject>",
            "tests/modules/Harness.bas": "TEST_<Subject>",
            "examples/modules/SingleNettingSet.bas": "Example<Subject>",
        }
        for path, rule in cases.items():
            with self.subTest(path=path):
                source = CORE.replace("CORE_Validation", Path(path).stem)
                self.assertFinding({path: source}, f"is named {rule}")

    def test_form_resource_rules(self) -> None:
        without = {k: v for k, v in VALID.items() if not k.endswith(".frx")}
        self.assertFinding(without, "form resource frmRunOptions.frx is not tracked")
        self.assertFinding({**VALID, "src/forms/frmRunOptions.frx": b""},
                           "resource offset is outside")
        self.assertFinding(changed("src/forms/frmRunOptions.frm",
                                   '"frmRunOptions.frx":0000', '"Other"'),
                           "no form resource")
        self.assertFinding({"src/workbook/frmLost.frx": b"\x00"}, "beside its form")

    def test_vb_name_must_equal_file_name(self) -> None:
        self.assertFinding(changed("src/core/CORE_Validation.bas", 'VB_Name = "CORE_Validation"',
                                   'VB_Name = "CORE_Checks"'),
                           'exactly one Attribute VB_Name = "CORE_Validation"')

    def test_bas_starts_with_vb_name(self) -> None:
        self.assertFinding(changed("src/core/CORE_Validation.bas", "Attribute VB_Name",
                                   "'Banner\nAttribute VB_Name"),
                           "starts with its Attribute VB_Name line")

    def test_names_are_case_insensitive(self) -> None:
        files = {"src/core/CORE_Ab.bas": CORE.replace("CORE_Validation", "CORE_Ab"),
                 "src/core/CORE_AB.bas": CORE.replace("CORE_Validation", "CORE_AB")}
        self.assertFinding(files, "component name also used by")

    def test_name_length_limit(self) -> None:
        name = "CORE_" + "A" * 27
        self.assertFinding({f"src/core/{name}.bas": CORE.replace("CORE_Validation", name)},
                           "longer than 31 characters")

    def test_module_options(self) -> None:
        self.assertFinding(changed("src/modules/SFT_Functions.bas", "Option Explicit\n", ""),
                           "missing Option Explicit")
        for path in ("src/core/CORE_Validation.bas", "tests/modules/TEST_Harness.bas"):
            with self.subTest(path=path):
                self.assertFinding(changed(path, "Option Private Module\n", ""),
                                   "missing Option Private Module")
        self.assertFinding(changed("src/core/CORE_Validation.bas", "Option Explicit\n",
                                   "Option Explicit\nOption Base 1\n"), "Option Base 1")

    def test_ascii_only(self) -> None:
        self.assertFinding(changed("src/core/CORE_Validation.bas", "Err.Description",
                                   "Err.Description & \"\udce9\""), "non-ASCII byte 0xE9")
        bom = {"src/core/CORE_Validation.bas": b"\xef\xbb\xbf" + CORE.encode("ascii")}
        self.assertFinding(bom, "byte-order mark")

    def test_tabs_and_trailing_white_space(self) -> None:
        self.assertFinding(changed("src/core/CORE_Validation.bas", "        IsPositive =",
                                   "\tIsPositive ="), "tab character")
        self.assertFinding(changed("src/core/CORE_Validation.bas", "Exit Function",
                                   "Exit Function  "), "trailing white space")

    def test_unclosed_string_literal(self) -> None:
        self.assertFinding(changed("tests/modules/TEST_Harness.bas", 'Debug.Print "It',
                                   'Debug.Print "Open\n        Debug.Print "It'),
                           "string literal not closed")

    def test_comment_continuation_is_comment(self) -> None:
        source = changed("tests/modules/TEST_Harness.bas", "End Sub",
                         "        ' A comment that continues _\n        on \" this line\nEnd Sub")
        self.assertEqual(self.findings(source), [])

    def test_storage_needs_lf_and_crlf_checkout(self) -> None:
        findings = self.findings({"src/core/CORE_Validation.bas": CORE}, attributes=False)
        self.assertTrue(any("stored with crlf" in item for item in findings), findings)
        self.assertTrue(any("check it out with CRLF" in item for item in findings), findings)

    def test_workbooks_are_not_committed(self) -> None:
        self.assertFinding({"src/workbook/CCR-K-SFTs.xlsm": b"PK\x03\x04"}, "workbook committed")

    def test_main_exit_codes(self) -> None:
        for files, expected in ((VALID, 0), ({"src/Stray.bas": CORE}, 1)):
            with self.subTest(expected=expected), repository(files) as root, \
                    contextlib.redirect_stdout(io.StringIO()) as output:
                self.assertEqual(check_vba_source.main(["--root", str(root)]), expected)
                self.assertIn("static check, nothing compiled or run", output.getvalue())


CONDITIONAL_VALID = {
    "nested": '''#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function F64 Lib "kernel32" () As LongPtr
    #ElseIf Win32 Then
        Private Declare PtrSafe Function F32 Lib "kernel32" () As Long
    #End If
#Else
    Private Declare Function FLegacy Lib "kernel32" () As Long
#End If
''',
    "inactive-legacy": '''#If Not VBA7 Then
    #If Win64 Then
        Private Declare Function Impossible Lib "kernel32" () As Long
    #End If
#End If
''',
    "boolean-expression": '''#If Not VBA7 Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#ElseIf VBA7 And (Win64 Or Win32) Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
    "continued-declare": '''#If VBA7 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" _
        () As LongPtr
#Else
    Private Declare Function Legacy Lib "kernel32" _
        () As Long
#End If
''',
    "literals": '''#If True And Not 0 Then
    Private Declare PtrSafe Function Modern Lib "kernel32" () As LongPtr
#End If
''',
}
CONDITIONAL_INVALID = {
    "declare-without-ptrsafe": ('''#If VBA7 And Win64 Then
    Private Declare Function Bad Lib "kernel32" () As Long
#End If
''', "without PtrSafe in vba7-win64"),
    "unconditional-declare": ('Private Declare Function Bad Lib "kernel32" () As Long\n',
                              "without PtrSafe in vba7-win32, vba7-win64"),
    "win32-is-true-on-win64": ('''#If Win32 Then
    Private Declare Function WrongOrder Lib "kernel32" () As Long
#ElseIf Win64 Then
    Private Declare PtrSafe Function NeverReached Lib "kernel32" () As LongPtr
#End If
''', "vba7-win64"),
    "unknown-symbol": ("#If Mac Then\n#End If\n", "unsupported conditional-compilation symbol"),
    "vba6-is-true-under-vba7": ('''#If VBA6 Then
    Private Declare Function Legacy Lib "kernel32" () As Long
#End If
''', "without PtrSafe in vba7-win32, vba7-win64"),
    "integer-literal": ("#If 2 Then\n#End If\n", "-1 and 0"),
    "missing-then": ("#If VBA7\n#End If\n", "must end with Then"),
    "project-const": ("#Const FEATURE = True\n", "project #Const"),
    "unclosed": ("#If VBA7 Then\n", "not closed"),
    "stray-elseif": ("#ElseIf VBA7 Then\n", "#ElseIf without #If"),
    "elseif-after-else": ("#If VBA7 Then\n#Else\n#ElseIf Win64 Then\n#End If\n",
                          "#ElseIf after #Else"),
    "duplicate-else": ("#If VBA7 Then\n#Else\n#Else\n#End If\n", "duplicate #Else"),
    "stray-end-if": ("#End If\n", "#End If without #If"),
    "unsupported-directive": ("#Region Main\n", "unsupported conditional-compilation directive"),
}
CONDITIONAL_VALID["continued-comment-hides-directive"] = "' retired guard _\n#If Mac Then\n"


def module(body: str) -> str:
    return f'Attribute VB_Name = "TEST_Fixture"\nOption Explicit\nOption Private Module\n{body}'


class ConditionalTest(unittest.TestCase):
    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(run_check(check_vba_conditionals.check, VALID), [])
        for name, body in CONDITIONAL_VALID.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_conditionals.check,
                                           standalone(module(body))), [])

    def test_invalid_fixtures_fail(self) -> None:
        for name, (body, text) in CONDITIONAL_INVALID.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_conditionals.check, standalone(module(body)))
                self.assertTrue(any(text in item for item in findings), findings)

    def test_reachable_sources_blank_inactive_lines(self) -> None:
        sources, findings = check_vba_conditionals.reachable_sources(
            "x.bas", "A\n#If VBA7 Then\nB\n#Else\nC\n#End If")
        self.assertEqual(findings, [])
        self.assertEqual(sources["vba7-win64"].split("\n"), ["A", "", "B", "", "", ""])
        self.assertEqual(sources["vba6-win32"].split("\n"), ["A", "", "", "", "C", ""])


JUMPS_VALID = {
    "local-handler": '''Public Sub First(ByVal Value As Long)
        On Error GoTo Handler
        GoSub Worker
        Resume 100
Worker:
        Return
100:    Exit Sub
Handler:
        Resume Next
End Sub
''',
    "same-label-in-two-procedures": '''Private Sub First()
Shared:
        Exit Sub
End Sub

Private Sub Second()
Shared:
        Exit Sub
End Sub
''',
    "numbered-label": '''Private Sub First()
        GoTo 100
100     Exit Sub
End Sub
''',
    "special-error-controls": '''Private Sub First()
        On Error GoTo 0
        On Error GoTo -1
        On Error Resume Next
        Resume Next
End Sub
''',
    "quoted-jump-text": '''Private Sub First()
        Debug.Print "GoSub Missing", "GoTo Lost", "Resume Nowhere"
        Debug.Print "He said ""GoSub Missing""; it's text"
        Debug.Print "GoSub Missing" & _
            "Resume Lost"
        Debug.Print "GoSub Missing": GoSub Worker
        Exit Sub
Worker:
        Return
End Sub
''',
    "keyword-before-colon": '''Private Sub First(ByVal Flag As Boolean)
        If Flag Then
            Debug.Print 1
        Else: Debug.Print 2
        End If
        If Not Flag Then
            Debug.Print 3
        Else: Debug.Print 4
        End If
End Sub
''',
    "rem-after-separator": '''Private Sub First()
        Debug.Print 1: Rem GoTo Missing
        Debug.Print 2:Rem GoSub Lost
End Sub
''',
    "one-line-procedure": '''Private Sub First(): Debug.Print 1: End Sub

Private Sub Second(): Debug.Print 2: End Sub
''',
    "rem-after-line-number": '''Private Sub First()
100 Rem GoTo Missing
        Exit Sub
End Sub
''',
    "numbered-comment-line-as-target": '''Private Sub First()
        GoTo 100
100 Rem target
        GoTo 200
200 ' target
End Sub
''',
    "bracketed-sheet-name": '''Private Sub First(): Debug.Print ['Sheet 1'!A1]: End Sub

Private Sub Second()
        Debug.Print [A1:B2]: GoTo Done
Done:
End Sub
''',
    "jump-keyword-in-a-sheet-name": '''Private Sub First()
        Debug.Print ['GoTo Missing'!A1], [GoSub!B2]: GoTo Done
Done:
End Sub
''',
    "computed-jump": '''Private Sub First(ByVal Choice As Long)
        On Choice GoTo One, Two
        On Choice GoSub One, Two
One:
Two:
End Sub
''',
    "colons-in-a-date-literal": '''Private Sub First()
        Debug.Print #12:30:00 PM#: GoTo Done
Done:
End Sub
''',
    "exclusive-branches": '''Private Sub First()
#If VBA7 Then
        GoTo Modern
Modern:
#Else
        GoTo Legacy
Legacy:
#End If
End Sub
''',
    "continued-comment-hides-jump": '''Private Sub First()
        ' a comment continued _
        GoTo Missing
End Sub
''',
    "continued-rem-hides-jump": '''Private Sub First()
        Rem a remark continued _
        GoTo Missing
End Sub
''',
    "comment-continued-twice": '''Private Sub First()
        ' first line _
          second line _
        GoTo Missing
End Sub
''',
    "continued-comment-after-label-and-separator": '''Private Sub First()
100     Debug.Print 1: ' note _
        GoTo Missing
        Debug.Print 2: Rem note _
        GoSub Lost
End Sub
''',
}
JUMPS_INVALID = {
    "cross-procedure": ('''Private Sub First()
        GoTo Shared
End Sub

Private Sub Second()
Shared:
        Exit Sub
End Sub
''', "GoTo target 'Shared' is not a label of procedure First"),
    "duplicate-label": ('''Private Sub First()
Again:
Again:
        Exit Sub
End Sub
''', "duplicate label 'Again' in procedure First; first defined at line"),
    "jump-after-string": ('''Private Sub First()
        Debug.Print "it's ""quoted""": GoSub Missing
End Sub
''', "GoSub target 'Missing'"),
    "resume-to-missing-label": ('''Private Sub First()
        On Error GoTo Handler
        Exit Sub
Handler:
        Resume Retry
End Sub
''', "Resume target 'Retry'"),
    "label-in-other-branch": ('''Private Sub First()
#If VBA7 Then
        GoTo Legacy
#Else
Legacy:
#End If
End Sub
''', "[vba7-win32, vba7-win64]"),
    "procedure-not-closed": ("Private Sub First()\n        Exit Sub\n", "is not closed"),
    "procedure-opened-twice": ("Private Sub First()\nPrivate Sub Second()\nEnd Sub\n",
                               "is not closed before Second"),
    "stray-end": ("End Sub\n", "End statement outside any procedure"),
    "computed-jump-missing-target": ("Private Sub First(ByVal Choice As Long)\n"
                                     "        On Choice GoTo Present, Missing\nPresent:\nEnd Sub\n",
                                     "GoTo target 'Missing'"),
    "jump-after-bracketed-sheet-name": ("Private Sub First()\n"
                                        "        Debug.Print ['Sheet 1'!A1]: GoTo Lost\nEnd Sub\n",
                                        "GoTo target 'Lost'"),
    "one-line-procedure-missing-label": ("Private Sub First(): GoTo Lost: End Sub\n",
                                         "GoTo target 'Lost'"),
    "continued-comment-hides-label": ("Private Sub First()\n        GoTo Missing\n"
                                      "        ' a comment continued _\nMissing:\nEnd Sub\n",
                                      "GoTo target 'Missing'"),
    "continued-comment-ends": ("Private Sub First()\n        ' a comment continued _\n"
                               "          on this line\n        GoTo Missing\nEnd Sub\n",
                               "GoTo target 'Missing'"),
    "continuation-mark-in-a-string": ("Private Sub First()\n        Debug.Print \"ends _\"\n"
                                      "        GoTo Missing\nEnd Sub\n", "GoTo target 'Missing'"),
    "continuation-mark-in-brackets": ("Private Sub First()\n        Debug.Print ['a _'!A1]\n"
                                      "        GoTo Missing\nEnd Sub\n", "GoTo target 'Missing'"),
}


class JumpTest(unittest.TestCase):
    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(run_check(check_vba_jumps.check, VALID), [])
        for name, body in JUMPS_VALID.items():
            with self.subTest(name=name):
                self.assertEqual(run_check(check_vba_jumps.check, standalone(module(body))), [])

    def test_invalid_fixtures_fail(self) -> None:
        for name, (body, text) in JUMPS_INVALID.items():
            with self.subTest(name=name):
                findings = run_check(check_vba_jumps.check, standalone(module(body)))
                self.assertTrue(any(text in item for item in findings), findings)

    def test_counts_procedures(self) -> None:
        with repository(VALID) as root:
            _, components, procedures = check_vba_jumps.check(root)
        self.assertEqual((components, procedures), (7, 6))


API_FACADE = '''Attribute VB_Name = "SFT_Api"
Option Explicit
Public Const SFT_LIMIT As Long = 7
Public Declare PtrSafe Function Tick Lib "kernel32" () As Long
Public gLastStatus As Long
Public Enum SftMode
    smFirst = 1
End Enum
Public Type SftPair
    Left As Long
    Right As Long
End Type
Private mHidden As Long

Public Function SFT_Echo( _
    ByVal Value As Long) _
    As Long
        SFT_Echo = Value
End Function

Public Property Get SFT_Current() As Long
        SFT_Current = gLastStatus
End Property

Public Property Let SFT_Current(ByVal Value As Long)
        gLastStatus = Value
End Property

Private Sub Helper()
End Sub
'''
API_ROWS = [
    ("Const", "SFT_LIMIT", "Public Const SFT_LIMIT As Long = 7"),
    ("Declare Function", "Tick", 'Public Declare PtrSafe Function Tick Lib "kernel32" () As Long'),
    ("Variable", "gLastStatus", "Public gLastStatus As Long"),
    ("Enum", "SftMode", "Public Enum SftMode | smFirst = 1 | End Enum"),
    ("Type", "SftPair", "Public Type SftPair | Left As Long | Right As Long | End Type"),
    ("Function", "SFT_Echo", "Public Function SFT_Echo(ByVal Value As Long) As Long"),
    ("Property Get", "SFT_Current", "Public Property Get SFT_Current() As Long"),
    ("Property Let", "SFT_Current", "Public Property Let SFT_Current(ByVal Value As Long)"),
]


def api_inventory(rows: list[tuple[str, str, str]], extra: str = "") -> str:
    lines = ["# Supported VBA declarations"]
    for kind, name, signature in rows:
        lines += [f"SFT_Api\t{kind}\t{name}", f"# SIG\tSFT_Api\t{kind}\t{name}\t{signature}"]
    return "\n".join(lines) + "\n" + extra


def api_files(facade: str = API_FACADE, inventory: str | None = None,
              **others: str) -> dict[str, str | bytes]:
    files: dict[str, str | bytes] = {
        "src/modules/SFT_Api.bas": facade,
        "src/core/CORE_Internal.bas": 'Attribute VB_Name = "CORE_Internal"\nOption Explicit\n'
                                      "Option Private Module\nPublic Function Inside() As Long\n"
                                      "End Function\n",
        "tests/modules/TEST_Api.bas": 'Attribute VB_Name = "TEST_Api"\nOption Explicit\n'
                                      "Option Private Module\nPublic Sub RunAll()\nEnd Sub\n",
        "docs/PUBLIC_API.txt": api_inventory(API_ROWS) if inventory is None else inventory,
    }
    files.update({f"src/modules/{name}.bas": text for name, text in others.items()})
    return files


DECLARE = 'Public Declare PtrSafe Function Tick Lib "kernel32" () As Long'
LEGACY = DECLARE.replace("PtrSafe ", "")
LEGACY_ROW = '# SIG\tSFT_Api\tDeclare Function\tTick\tPublic Declare Function Tick Lib "kernel32" () As Long\n'


class PublicApiTest(unittest.TestCase):
    def findings(self, files: Mapping[str, str | bytes]) -> list[str]:
        return run_check(check_vba_public_api.check, files)

    def assertFinding(self, files: Mapping[str, str | bytes], text: str) -> None:
        findings = self.findings(files)
        self.assertTrue(any(text in item for item in findings), findings)

    def test_valid_fixtures_pass(self) -> None:
        self.assertEqual(self.findings(VALID), [])
        self.assertEqual(self.findings(api_files()), [])

    def test_repository_inventory_matches_empty_source(self) -> None:
        inventory = (ROOT / "docs/PUBLIC_API.txt").read_text(encoding="utf-8")
        self.assertEqual(self.findings({"docs/PUBLIC_API.txt": inventory}), [])

    def test_implicit_visibility_is_rejected_everywhere(self) -> None:
        self.assertFinding(api_files(API_FACADE.replace("Public Function SFT_Echo",
                                                        "Function SFT_Echo")),
                           "implicit public procedure: SFT_Echo")
        hidden = 'Attribute VB_Name = "CORE_Implicit"\nOption Explicit\nOption Private Module\n' \
                 "Sub Hidden()\nEnd Sub\nEnum Mode\n    mA = 1\nEnd Enum\n"
        findings = self.findings({"src/core/CORE_Implicit.bas": hidden,
                                  "docs/PUBLIC_API.txt": "# None\n"})
        self.assertTrue(any("implicit public procedure: Hidden" in f for f in findings), findings)
        self.assertTrue(any("implicit public module-level declaration: Enum Mode" in f
                            for f in findings), findings)

    def test_one_identifier_per_public_statement(self) -> None:
        self.assertFinding(api_files(API_FACADE.replace("Public gLastStatus As Long",
                                                        "Public gLastStatus As Long, gOther")),
                           "Public variable statement must declare one identifier")
        self.assertFinding(api_files(API_FACADE.replace("= 7", "= 7, SFT_MORE As Long = 8")),
                           "Public Const statement must declare one identifier")

    def test_inventory_must_list_every_declaration(self) -> None:
        self.assertFinding(api_files(inventory=api_inventory(API_ROWS[1:])),
                           "public declaration not listed: SFT_Api\tConst\tSFT_LIMIT")
        rows = api_inventory(API_ROWS).replace("SFT_Api\tFunction\tSFT_Echo\n", "")
        self.assertFinding(api_files(inventory=rows), "not listed: SFT_Api\tFunction\tSFT_Echo")

    def test_signature_records(self) -> None:
        rows = api_inventory(API_ROWS).replace("(ByVal Value As Long) As Long", "() As Long")
        self.assertFinding(api_files(inventory=rows), "signature mismatch for SFT_Api\tFunction")
        missing = "\n".join(line for line in api_inventory(API_ROWS).splitlines()
                            if "SIG\tSFT_Api\tFunction" not in line) + "\n"
        self.assertFinding(api_files(inventory=missing), "signature record missing")
        duplicate = api_inventory(API_ROWS, api_inventory(API_ROWS[:1]).splitlines()[2] + "\n")
        self.assertFinding(api_files(inventory=duplicate), "duplicate signature record")
        self.assertFinding(api_files(inventory=api_inventory(API_ROWS, "# SIG\tSFT_Api\n")),
                           "malformed # SIG record")

    def test_stale_and_malformed_rows(self) -> None:
        stale = api_inventory(API_ROWS, "SFT_Api\tFunction\tGone\n"
                                        "# SIG\tSFT_Api\tFunction\tGone\tPublic Function Gone()\n")
        findings = self.findings(api_files(inventory=stale))
        self.assertTrue(any("listed declaration not in src/modules/" in f for f in findings))
        self.assertTrue(any("stale signature record" in f for f in findings))
        self.assertFinding(api_files(inventory=api_inventory(API_ROWS, "SFT_Api\tFunction\n")),
                           "needs component, kind and name")
        self.assertFinding(api_files(inventory=api_inventory(API_ROWS, "SFT_Api\tConst\tsft_limit\n")),
                           "duplicate inventory row")

    def test_only_the_facade_is_supported_api(self) -> None:
        rows = api_inventory(API_ROWS, "CORE_Internal\tFunction\tInside\n")
        self.assertFinding(api_files(inventory=rows),
                           "listed declaration not in src/modules/: CORE_Internal")

    def test_missing_inventory(self) -> None:
        files = {k: v for k, v in api_files().items() if k != "docs/PUBLIC_API.txt"}
        self.assertFinding(files, "public API inventory is missing")

    def test_collisions(self) -> None:
        twice = API_FACADE + "\nPublic Function SFT_Echo(ByVal Value As String) As String\n" \
                             "End Function\n"
        self.assertFinding(api_files(twice), "public declaration appears more than once")
        other = 'Attribute VB_Name = "SFT_Other"\nOption Explicit\n' \
                "Public Function SFT_Echo(ByVal Value As Long) As Long\nEnd Function\n"
        self.assertFinding(api_files(SFT_Other=other), "public name 'SFT_Echo' collides with "
                                                       "SFT_Api.SFT_Echo (Function)")

    def test_conditional_variants(self) -> None:
        conditional = API_FACADE.replace(DECLARE, f"#If VBA7 Then\n{DECLARE}\n#Else\n{LEGACY}\n"
                                                  "#End If")
        variants = api_inventory(API_ROWS, LEGACY_ROW)
        self.assertEqual(self.findings(api_files(conditional, variants)), [])
        self.assertFinding(api_files(conditional), "signature mismatch")
        self.assertFinding(api_files(inventory=variants), "signature mismatch")
        self.assertFinding(api_files(conditional + "\n" + DECLARE + "\n", variants),
                           "appears more than once")
        dead = API_FACADE + f"\n#If False Then\n{DECLARE.replace('Tick', 'Dead')}\n#End If\n"
        self.assertEqual(self.findings(api_files(dead)), [])
        self.assertFinding(api_files(conditional.replace("#If VBA7", "#If Unknown"), variants),
                           "indeterminate #If")

    def test_exclusive_names_across_modules(self) -> None:
        modern_only = API_FACADE.replace(DECLARE, f"#If VBA7 Then\n{DECLARE}\n#End If")
        legacy = f'Attribute VB_Name = "SFT_Legacy"\nOption Explicit\n#If Not VBA7 Then\n{LEGACY}\n' \
                 "#End If\n"
        rows = api_inventory(API_ROWS, "SFT_Legacy\tDeclare Function\tTick\n# SIG\tSFT_Legacy\t"
                                       'Declare Function\tTick\tPublic Declare Function Tick Lib '
                                       '"kernel32" () As Long\n')
        self.assertEqual(self.findings(api_files(modern_only, rows, SFT_Legacy=legacy)), [])
        self.assertFinding(api_files(API_FACADE, rows, SFT_Legacy=legacy), "collides")

    def test_signatures_keep_string_literals(self) -> None:
        facade = API_FACADE.replace("Private Sub Helper()", "Public Function SFT_Pad( _\n"
                                    '    Optional ByVal Text As String = "A  ( B )") _\n'
                                    "    As String\nEnd Function\n\nPrivate Sub Helper()")
        exact = api_inventory(API_ROWS + [("Function", "SFT_Pad", 'Public Function SFT_Pad('
                                           'Optional ByVal Text As String = "A  ( B )") '
                                           "As String")])
        self.assertEqual(self.findings(api_files(facade, exact)), [])
        collapsed = exact.replace('"A  ( B )"', '"A (B)"')
        self.assertFinding(api_files(facade, collapsed), "signature mismatch for SFT_Api\tFunction"
                                                         "\tSFT_Pad")

    def test_declare_inside_a_literal_is_text(self) -> None:
        facade = API_FACADE.replace("Private Sub Helper()", "Public Function SFT_Say("
                                    'Optional ByVal Text As String = " declare ") As String\n'
                                    "End Function\n\nPrivate Sub Helper()")
        rows = api_inventory(API_ROWS + [("Function", "SFT_Say", "Public Function SFT_Say("
                                          'Optional ByVal Text As String = " declare ") As String')])
        self.assertEqual(self.findings(api_files(facade, rows)), [])

    def test_colon_separated_and_commented_statements(self) -> None:
        facade = API_FACADE.replace("Private mHidden As Long",
                                    "Private mHidden As Long: Public gExtra As Long ' note: x")
        self.assertFinding(api_files(facade), "not listed: SFT_Api\tVariable\tgExtra")
        dated = API_FACADE.replace("Private mHidden As Long",
                                   "Private Const START_DATE As Date = #1/2/2026#: Rem : Public x")
        self.assertEqual(self.findings(api_files(dated)), [])

    def test_continued_comment_is_not_a_declaration(self) -> None:
        retired = API_FACADE.replace("Private mHidden As Long",
                                     "Private mHidden As Long ' retired: _\nPublic gRetired As Long")
        self.assertEqual(self.findings(api_files(retired)), [])


class LexerTest(unittest.TestCase):
    def test_comments_end_the_code(self) -> None:
        cases = {
            "x = 1 ' note": "x = 1",
            "Rem whole line": "",
            "  Rem": "",
            "Debug.Print 1: Rem GoTo Missing": "Debug.Print 1:",
            'Debug.Print "a: Rem b"': 'Debug.Print "a: Rem b"',
            "Call Run(Remainder:=1)": "Call Run(Remainder:=1)",
            "Report Name:= Rem": "Report Name:= Rem",
            "Remove 1": "Remove 1",
            "100 Rem GoTo Missing": "100",
            "  10 x = 1 ' note": "  10 x = 1",
            "10: Rem note": "10:",
            "x = ['Sheet 1'!A1] ' note": "x = ['Sheet 1'!A1]",
            "x = [A1] ' note": "x = [A1]",
            "x = 1: Remark = 2": "x = 1: Remark = 2",
            '"it\'s": GoTo Label': '"it\'s": GoTo Label',
        }
        for raw, expected in cases.items():
            with self.subTest(raw=raw):
                self.assertEqual(strip_code(raw), expected)
        self.assertEqual(strip_code('Debug.Print "GoTo x": GoTo y', blank_strings=True),
                         "Debug.Print         : GoTo y")
        self.assertEqual(strip_code("x = ['GoTo z'!A1]: GoTo y", blank_strings=True),
                         "x = [" + " " * len("'GoTo z'!A1") + "]: GoTo y")

    def test_a_comment_ending_in_a_continuation_continues(self) -> None:
        lines = ["x = 1 ' a _", "GoTo y _", "  z", "Rem b _", "w", "s = \"t _\"", "v = 2"]
        self.assertEqual([code for _, _, _, code in logical_units(lines)],
                         ["x = 1", "", "", "", "", 's = "t _"', "v = 2"])


class MainTest(unittest.TestCase):
    def test_every_gate_exits_0_when_valid_and_1_when_invalid(self) -> None:
        # Each gate has its own defect here: no Option Explicit, an unknown
        # #If symbol, a jump to a missing label and an implicit public function.
        invalid = standalone(module("Private Sub First()\n        GoTo Lost\nEnd Sub\n"
                                    "Function Implicit()\nEnd Function\n#If Mac Then\n#End If\n"))
        invalid["tests/modules/TEST_Fixture.bas"] = \
            str(invalid["tests/modules/TEST_Fixture.bas"]).replace("Option Explicit\n", "")
        mains = (check_vba_source.main, check_vba_conditionals.main, check_vba_jumps.main,
                 check_vba_public_api.main)
        for files, expected in ((VALID, 0), (invalid, 1)):
            with repository(files) as root:
                for main in mains:
                    with self.subTest(main=main.__module__, expected=expected), \
                            contextlib.redirect_stdout(io.StringIO()) as output:
                        self.assertEqual(main(["--root", str(root)]), expected)
                        verdict = output.getvalue().splitlines()[-1]
                        self.assertIn("PASS" if expected == 0 else "FAIL", verdict)
                        self.assertIn("static check, nothing compiled or run", verdict)

    def test_a_tracked_file_missing_from_the_working_tree_is_a_finding(self) -> None:
        def run_check_without(check: Callable[[Path], tuple[list[Finding], *tuple[int, ...]]],
                              deleted: str) -> list[str]:
            # Deleted after staging, so the file stays tracked.
            with repository(VALID) as root:
                (root / deleted).unlink()
                return [str(finding) for finding in check(root)[0]]

        checks = (check_vba_source.check, check_vba_conditionals.check, check_vba_jumps.check,
                  check_vba_public_api.check)
        missing = "src/modules/SFT_Functions.bas: tracked in Git but missing from the working tree"
        for check in checks:
            with self.subTest(check=check.__module__):
                self.assertIn(missing, run_check_without(check, "src/modules/SFT_Functions.bas"))
        self.assertIn("src/forms/frmRunOptions.frx: tracked in Git but missing from the "
                      "working tree",
                      run_check_without(check_vba_source.check, "src/forms/frmRunOptions.frx"))


if __name__ == "__main__":
    unittest.main()
