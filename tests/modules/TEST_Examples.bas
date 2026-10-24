Attribute VB_Name = "TEST_Examples"
'==============================================================================
' MODULE: TEST_Examples
'------------------------------------------------------------------------------
' PURPOSE
'   Neutral example suite of TEST_Harness: one case per kind of assertion,
'   and the reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup) with its cases. Nothing
'   here is an SFT calculation or a regulatory value.
'
' PUBLIC SURFACE
'   RunExampleTests is the documented test entry point.
'   RunExampleTestsWithInjectedFailure and RunExampleTestsWithSkippedCase
'   show what a failed and an incomplete run report. Option Private Module
'   keeps all three out of the workbook's external surface.
'
' DEPENDENCIES
'   TEST_Harness; the VBA and Excel object models.
'
' STATE OWNERSHIP
'   Owns the mode of the running suite. RunExampleOperation changes
'   Application.Calculation and Application.ScreenUpdating and restores them
'   before it returns, on the success and the error path.
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START.
'
' USAGE
'   In a development workbook holding the modules of tests/modules/, type
'   TEST_Examples.RunExampleTests in the Immediate window.
'
' UPDATED
'   2026-10-24
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the tests inside the project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once.
        Private Const SUITE_NAME              As String = "examples"
        Private Const EXPECTED_CASES          As Long = 9
        Private Const EXPECTED_ASSERTIONS     As Long = 18
    'Errors of the example procedures, test-only.
        Private Const ERR_EXAMPLE_INVALID     As Long = vbObjectError + 4080   'Input outside the domain
        Private Const ERR_EXAMPLE_OPERATION   As Long = vbObjectError + 4081   'Primary error of the operation
        Private Const ERR_EXAMPLE_RESTORE     As Long = vbObjectError + 4082   'Simulated restore failure
        Private Const ERR_CLEANUP_FAILED      As Long = vbObjectError + 4083   'Only the cleanup failed

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
        Private mMode   As String    'NORMAL, INJECTED_FAILURE or SKIPPED_CASE


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunExampleTests()
'
'==============================================================================
'                               RunExampleTests
'------------------------------------------------------------------------------
' PURPOSE
'   Run the example suite. A passing run ends with
'   RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE;
'   cases=9; assertions=18; failures=0; cleanup=PASS; recovery=NONE.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        RunSuite "NORMAL"

End Sub


Public Sub RunExampleTestsWithInjectedFailure()
'
'==============================================================================
'                      RunExampleTestsWithInjectedFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Show the failure path: exact-double expects a wrong value. Expected
'   end: RESULT=FAIL; ...; operation=FAIL; completeness=COMPLETE; cases=9;
'   assertions=18; failures=1, then the error ERR_TEST_FAILED.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        RunSuite "INJECTED_FAILURE"

End Sub


Public Sub RunExampleTestsWithSkippedCase()
'
'==============================================================================
'                        RunExampleTestsWithSkippedCase
'------------------------------------------------------------------------------
' PURPOSE
'   Show that an incomplete run is never a pass: the last case is skipped.
'   Expected end: RESULT=FAIL; ...; operation=PASS;
'   completeness=INCOMPLETE; cases=8; assertions=16; failures=1, the
'   failure being harness.completeness, then ERR_TEST_FAILED.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        RunSuite "SKIPPED_CASE"

End Sub


'
'------------------------------------------------------------------------------
'
'                                    RUNNER
'
'------------------------------------------------------------------------------
'

Private Sub RunSuite( _
    ByVal Mode As String)
'
'==============================================================================
'                                   RunSuite
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun.
'
' INPUTS
'   Mode: NORMAL, INJECTED_FAILURE or SKIPPED_CASE.
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it
'   and re-raises it unchanged. BeginRun's refusal propagates directly.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrNumber        As Long      'Runner error number; 0 for none
    Dim ErrSource        As String    'Runner error source
    Dim ErrDescription   As String    'Runner error description

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        mMode = Mode
        TEST_Harness.BeginRun SUITE_NAME, Mode, EXPECTED_CASES, EXPECTED_ASSERTIONS
        On Error GoTo RunFailed
        CaseExactDouble
        CaseAbsoluteTolerance
        CaseRelativeTolerance
        CaseExactText
        CaseErrorValue
        CaseRaisedError
        CaseCleanupRestoresState
        CaseCleanupKeepsPrimaryError
        If mMode <> "SKIPPED_CASE" Then
            CaseCleanupOnlyFailure
        End If

Finish:
        On Error GoTo 0
        mMode = vbNullString
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription
        Exit Sub

RunFailed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume Finish

End Sub


'
'------------------------------------------------------------------------------
'
'                                    CASES
'
'------------------------------------------------------------------------------
'

Private Sub CaseExactDouble()
'
'==============================================================================
'                               CaseExactDouble
'------------------------------------------------------------------------------
' PURPOSE
'   Exact equality on a result that is exact in binary. 1 assertion.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Expected   As Double    'Expected sum; wrong on purpose when injected

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "exact-double"
        Expected = 0.75
        If mMode = "INJECTED_FAILURE" Then
            Expected = 0.76
        End If
        TEST_Harness.AssertEqualDouble "half-plus-quarter", Expected, 0.5 + 0.25
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseAbsoluteTolerance()
'
'==============================================================================
'                            CaseAbsoluteTolerance
'------------------------------------------------------------------------------
' PURPOSE
'   The absolute bound: 0.1 + 0.2 is not exactly 0.3 in binary, and near
'   zero only the absolute bound counts. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "tolerance-absolute"
        TEST_Harness.AssertNear "tenth-plus-fifth", 0.3, 0.1 + 0.2, 0.000000000000001, 0#
        TEST_Harness.AssertNear "near-zero", 0#, 1E-20, 0.000000000001, 0.000000001
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseRelativeTolerance()
'
'==============================================================================
'                            CaseRelativeTolerance
'------------------------------------------------------------------------------
' PURPOSE
'   The relative bound governs large values, the bound is inclusive, and a
'   value outside both bounds is rejected. 3 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "tolerance-relative"
        TEST_Harness.AssertNear "large-value", 1000000000#, 1000000000.1, 0.000001, 0.000000001
        TEST_Harness.AssertTrue "bound-is-inclusive", _
            TEST_Harness.WithinTolerance(2#, 2.5, 0.5, 0#), "|2.5 - 2| = 0.5 must be within 0.5"
        TEST_Harness.AssertTrue "outside-both-bounds", _
            Not TEST_Harness.WithinTolerance(1#, 1.1, 0.01, 0.01), _
            "|1.1 - 1| = 0.1 must exceed max(0.01, 0.01 * 1)"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseExactText()
'
'==============================================================================
'                                CaseExactText
'------------------------------------------------------------------------------
' PURPOSE
'   Text is compared exactly, case included. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "text-exact"
        TEST_Harness.AssertEqualText "concatenation", "CASE-1", "CASE" & "-" & CStr(1)
        TEST_Harness.AssertTrue "case-matters", _
            StrComp("a", "A", vbBinaryCompare) <> 0, "a and A must differ"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseErrorValue()
'
'==============================================================================
'                                CaseErrorValue
'------------------------------------------------------------------------------
' PURPOSE
'   An expected worksheet error value, returned rather than raised.
'   1 assertion.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "error-value"
        TEST_Harness.AssertErrorValue "zero-denominator", xlErrDiv0, ExampleRatio(1#, 0#)
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseRaisedError()
'
'==============================================================================
'                               CaseRaisedError
'------------------------------------------------------------------------------
' PURPOSE
'   An expected raised error, captured on the line after the call, and the
'   same call with a valid input raising nothing. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Result        As Double    'Value of the checked square root
    Dim RaisedError   As Long      'Err.Number right after the call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "raised-error"
        On Error Resume Next
        Result = ExampleSquareRoot(-1#)
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "negative-input", ERR_EXAMPLE_INVALID, RaisedError
        On Error Resume Next
        Result = ExampleSquareRoot(4#)
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "valid-input", 0, RaisedError
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupRestoresState()
'
'==============================================================================
'                           CaseCleanupRestoresState
'------------------------------------------------------------------------------
' PURPOSE
'   A successful operation restores what it changed and reports no cleanup
'   failure. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details       As String     'Cleanup failures reported by the operation
    Dim SavedScreen   As Boolean    'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-restores-state"
        SavedScreen = Application.ScreenUpdating
        RunExampleOperation False, False, Details
        TEST_Harness.AssertEqualText "no-cleanup-failure", vbNullString, Details
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupKeepsPrimaryError()
'
'==============================================================================
'                         CaseCleanupKeepsPrimaryError
'------------------------------------------------------------------------------
' PURPOSE
'   When the operation and its cleanup both fail, the caller receives the
'   primary error, and the cleanup failure is still reported. 3 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details       As String    'Cleanup failures reported by the operation
    Dim RaisedError   As Long      'Err.Number right after the call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-keeps-primary-error"
        On Error Resume Next
        RunExampleOperation True, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "primary-error", ERR_EXAMPLE_OPERATION, RaisedError
        TEST_Harness.AssertTrue "calculation-failure-reported", _
            InStr(Details, "Calculation not restored") > 0, "details='" & Details & "'"
        TEST_Harness.AssertTrue "screen-failure-reported", _
            InStr(Details, "ScreenUpdating not restored") > 0, "details='" & Details & "'"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


Private Sub CaseCleanupOnlyFailure()
'
'==============================================================================
'                            CaseCleanupOnlyFailure
'------------------------------------------------------------------------------
' PURPOSE
'   When only the cleanup fails, the operation raises ERR_CLEANUP_FAILED,
'   never success. 2 assertions.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Details       As String    'Cleanup failures reported by the operation
    Dim RaisedError   As Long      'Err.Number right after the call

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
        On Error GoTo CaseFailed
        TEST_Harness.BeginCase "cleanup-only-failure"
        On Error Resume Next
        RunExampleOperation False, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
        TEST_Harness.AssertRaised "cleanup-error", ERR_CLEANUP_FAILED, RaisedError
        TEST_Harness.AssertTrue "failure-reported", Len(Details) > 0, "no cleanup detail returned"
        Exit Sub

CaseFailed:
        TEST_Harness.RecordCaseError

End Sub


'
'------------------------------------------------------------------------------
'
'                         EXAMPLE PROCEDURES UNDER TEST
'
'------------------------------------------------------------------------------
'

Private Function ExampleRatio( _
    ByVal Numerator As Double, _
    ByVal Denominator As Double) _
    As Variant
'
'==============================================================================
'                                 ExampleRatio
'------------------------------------------------------------------------------
' PURPOSE
'   Numerator / Denominator, or the worksheet error #DIV/0! for a zero
'   denominator, as a worksheet function returns errors (house style).
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If Denominator = 0# Then
            ExampleRatio = CVErr(xlErrDiv0)
        Else
            ExampleRatio = Numerator / Denominator
        End If

End Function


Private Function ExampleSquareRoot( _
    ByVal Value As Double) _
    As Double
'
'==============================================================================
'                              ExampleSquareRoot
'------------------------------------------------------------------------------
' PURPOSE
'   Square root of Value >= 0; a negative Value raises ERR_EXAMPLE_INVALID,
'   as a core procedure raises (house style).
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If Value < 0# Then
            Err.Raise ERR_EXAMPLE_INVALID, "TEST_Examples.ExampleSquareRoot", _
                "Value must be >= 0; got " & CStr(Value) & "."
        End If
        ExampleSquareRoot = Sqr(Value)

End Function


Private Sub RunExampleOperation( _
    ByVal FailOperation As Boolean, _
    ByVal FailRestore As Boolean, _
    ByRef CleanupDetails As String)
'
'==============================================================================
'                             RunExampleOperation
'------------------------------------------------------------------------------
' PURPOSE
'   Reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup): capture, change, restore
'   in one block on both paths, preserve the primary error, report every
'   cleanup failure.
'
' INPUTS
'   FailOperation: raise ERR_EXAMPLE_OPERATION inside the work.
'   FailRestore: test seam; each restore helper restores its setting and
'   then reports a simulated failure, so that the failure paths can be
'   tested without leaving Excel changed.
'
' RETURNS
'   CleanupDetails: the cleanup failures joined by "; ", empty on success;
'   set also when an error is raised.
'
' STATE OWNERSHIP
'   Changes Application.Calculation to manual and ScreenUpdating to False,
'   and restores both to the values it captured.
'
' ERROR POLICY
'   Re-raises the primary error after cleanup; when only the cleanup failed,
'   raises ERR_CLEANUP_FAILED with the details.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedCalculation   As XlCalculation    'Calculation mode before the operation
    Dim SavedScreen        As Boolean          'ScreenUpdating before the operation
    Dim StateCaptured      As Boolean          'True once both were read
    Dim ErrNumber          As Long             'Primary error number; 0 for none
    Dim ErrSource          As String           'Primary error source
    Dim ErrDescription     As String           'Primary error description

'------------------------------------------------------------------------------
' CAPTURE AND CHANGE
'------------------------------------------------------------------------------
        CleanupDetails = vbNullString
        On Error GoTo Failed
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

'------------------------------------------------------------------------------
' WORK
'------------------------------------------------------------------------------
        If FailOperation Then
            Err.Raise ERR_EXAMPLE_OPERATION, "TEST_Examples.RunExampleOperation", _
                "Simulated failure of the operation."
        End If

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
CleanUp:
        On Error GoTo 0
        If StateCaptured Then
            CleanupDetails = AppendDetail(RestoreCalculation(SavedCalculation, FailRestore), _
                                          RestoreScreenUpdating(SavedScreen, FailRestore))
        End If
        If ErrNumber <> 0 Then
            Err.Raise ErrNumber, ErrSource, ErrDescription
        End If
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "TEST_Examples.RunExampleOperation", CleanupDetails
        End If
        Exit Sub

'------------------------------------------------------------------------------
' HANDLE THE PRIMARY ERROR
'------------------------------------------------------------------------------
Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume CleanUp

End Sub


Private Function RestoreCalculation( _
    ByVal SavedCalculation As XlCalculation, _
    ByVal SimulateFailure As Boolean) _
    As String
'
'==============================================================================
'                              RestoreCalculation
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.Calculation; contain any error.
'
' RETURNS
'   Empty on success, otherwise what failed.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo Failed
        Application.Calculation = SavedCalculation
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreCalculation", "simulated"
        End If
        Exit Function

Failed:
        RestoreCalculation = "Calculation not restored: " & Err.Description

End Function


Private Function RestoreScreenUpdating( _
    ByVal SavedScreen As Boolean, _
    ByVal SimulateFailure As Boolean) _
    As String
'
'==============================================================================
'                            RestoreScreenUpdating
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.ScreenUpdating; contain any error.
'
' RETURNS
'   Empty on success, otherwise what failed.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        On Error GoTo Failed
        Application.ScreenUpdating = SavedScreen
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreScreenUpdating", "simulated"
        End If
        Exit Function

Failed:
        RestoreScreenUpdating = "ScreenUpdating not restored: " & Err.Description

End Function


Private Function AppendDetail( _
    ByVal First As String, _
    ByVal Second As String) _
    As String
'
'==============================================================================
'                                 AppendDetail
'------------------------------------------------------------------------------
' PURPOSE
'   Join two failure texts with "; ", skipping empty ones.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If Len(First) = 0 Then
            AppendDetail = Second
        ElseIf Len(Second) = 0 Then
            AppendDetail = First
        Else
            AppendDetail = First & "; " & Second
        End If

End Function
