Attribute VB_Name = "TEST_Examples"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Examples
'------------------------------------------------------------------------------
' PURPOSE
'   Neutral example suite of TEST_Harness: one case per kind of assertion,
'   and the reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup) with its cases. Nothing
'   here is an SFT calculation or a regulatory value.
'
' WHY THIS EXISTS
'   It shows how each kind of assertion of TEST_Harness is used, what a
'   passing, a failed and an incomplete run report, and it holds the
'   procedure that docs/VBA_HOUSE_STYLE.md names as the reference
'   implementation of the cleanup pattern, next to the cases that test it.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Runs nine cases in a fixed order between TEST_Harness.BeginRun and
'   TEST_Harness.EndRun. The mode of the run selects:
'     - NORMAL: every case, every expected value correct
'     - INJECTED_FAILURE: exact-double expects a wrong value
'     - SKIPPED_CASE: the last case is not run
'
' ERROR POLICY
'   As TEST_Harness.EndRun: a run that does not pass raises ERR_TEST_FAILED,
'   or the unexpected error that escaped the runner. A refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START. SaveExampleEvidence prints those
'   errors instead of raising them, after saving each run's log; an error
'   preparing its folder propagates.
'
' DEPENDENCIES
'   TEST_Harness; the VBA and Excel object models.
'
' NOTES
'   None
'
' PUBLIC SURFACE
'   RunExampleTests is the documented test entry point.
'   RunExampleTestsWithInjectedFailure and RunExampleTestsWithSkippedCase
'   show what a failed and an incomplete run report. SaveExampleEvidence
'   runs all three and saves their logs as the evidence of an Excel run.
'   Option Private Module keeps them out of the workbook's external surface.
'
' STATE OWNERSHIP
'   Owns the mode of the running suite. RunExampleOperation changes
'   Application.Calculation and Application.ScreenUpdating and restores them
'   before it returns, on the success and the error path.
'
' USAGE
'   In a development workbook holding the modules of tests/modules/, type
'   TEST_Examples.RunExampleTests in the Immediate window. To record an
'   Excel run, type TEST_Examples.SaveExampleEvidence "<folder>" instead
'   (INSTALLATION.md, Recording an Excel run).
'
' AUTHOR
'   Daniele Penza
'
' UPDATED
'   2026-10-26
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'A complete run starts every case and evaluates every assertion once
    Private Const SUITE_NAME                As String = "examples"          'Suite name written to the log
    Private Const EXPECTED_CASES            As Long = 9                     'Cases a complete run starts
    Private Const EXPECTED_ASSERTIONS       As Long = 18                    'Assertions a complete run evaluates

    'Errors of the example procedures, test-only
    Private Const ERR_EXAMPLE_INVALID       As Long = vbObjectError + 4080  'Input outside the domain
    Private Const ERR_EXAMPLE_OPERATION     As Long = vbObjectError + 4081  'Primary error of the operation
    Private Const ERR_EXAMPLE_RESTORE       As Long = vbObjectError + 4082  'Simulated restore failure
    Private Const ERR_CLEANUP_FAILED        As Long = vbObjectError + 4083  'Only the cleanup failed

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mMode   As String   'NORMAL, INJECTED_FAILURE or SKIPPED_CASE


'
'------------------------------------------------------------------------------
'
'                                 ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub RunExampleTests()

'
'------------------------------------------------------------------------------
'                            RUN THE EXAMPLE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Run the example suite. A passing run ends with
'   RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE;
'   cases=9; assertions=18; failures=0; cleanup=PASS; recovery=NONE.
'
' WHY THIS EXISTS
'   It is the documented test entry point of the suite.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the suite in NORMAL mode.
'
' ERROR POLICY
'   As RunSuite: a run that does not pass raises ERR_TEST_FAILED or the
'   unexpected error that escaped the runner; a refused start raises
'   TEST_Harness.ERR_TEST_DIRTY_START.
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   Typical call, in the Immediate window:
'
'       TEST_Examples.RunExampleTests
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run every case in NORMAL mode
        RunSuite "NORMAL"

End Sub

Public Sub RunExampleTestsWithInjectedFailure()

'
'------------------------------------------------------------------------------
'                              SHOW A FAILED RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Show the failure path: exact-double expects a wrong value. Expected
'   end: RESULT=FAIL; ...; operation=FAIL; completeness=COMPLETE; cases=9;
'   assertions=18; failures=1, then the error ERR_TEST_FAILED.
'
' WHY THIS EXISTS
'   It shows what a failed run reports.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the suite in INJECTED_FAILURE mode, in which CaseExactDouble
'   expects 0.76 instead of 0.75.
'
' ERROR POLICY
'   Raises ERR_TEST_FAILED at the end of the run, as expected; a refused
'   start raises TEST_Harness.ERR_TEST_DIRTY_START.
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run every case with one wrong expected value
        RunSuite "INJECTED_FAILURE"

End Sub

Public Sub RunExampleTestsWithSkippedCase()

'
'------------------------------------------------------------------------------
'                            SHOW AN INCOMPLETE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Show that an incomplete run is never a pass: the last case is skipped.
'   Expected end: RESULT=FAIL; ...; operation=PASS;
'   completeness=INCOMPLETE; cases=8; assertions=16; failures=1, the
'   failure being harness.completeness, then ERR_TEST_FAILED.
'
' WHY THIS EXISTS
'   It shows what an incomplete run reports.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the suite in SKIPPED_CASE mode, in which RunSuite does not run
'   CaseCleanupOnlyFailure.
'
' ERROR POLICY
'   Raises ERR_TEST_FAILED at the end of the run, as expected; a refused
'   start raises TEST_Harness.ERR_TEST_DIRTY_START.
'
' DEPENDENCIES
'   RunSuite
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run every case but the last
        RunSuite "SKIPPED_CASE"

End Sub

Public Sub SaveExampleEvidence( _
    ByVal Folder As String)
'
'------------------------------------------------------------------------------
'                      SAVE THE EVIDENCE OF AN EXCEL RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Run the suite in its three modes and save each complete log in Folder:
'   examples.log, injected.log and skipped.log, which
'   tools/check_excel_run.py checks against the commit. Folder is created
'   when missing, and the logs and host.txt of an earlier run are removed
'   first; its exported\ folder is left as it is.
'
' WHY THIS EXISTS
'   The three logs are the evidence of an Excel run that
'   tools/check_excel_run.py checks.
'
' INPUTS
'   Folder
'     A folder outside the checkout or in the ignored test-results/
'     directory, whose parent exists, without a trailing backslash.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   1. Creates Folder when it does not exist.
'   2. Deletes examples.log, injected.log, skipped.log and host.txt when an
'      earlier run left them in Folder.
'   3. Runs and saves the NORMAL, INJECTED_FAILURE and SKIPPED_CASE runs,
'      in that order, through SaveRun.
'
' ERROR POLICY
'   The error that ends each run is its reported outcome: it is printed
'   after the log is saved, not raised, so that all three logs are written.
'   Whether the runs ended as documented is judged by check_excel_run.py,
'   never here. An error of MkDir, Kill or SaveLog propagates.
'
' DEPENDENCIES
'   SaveRun; Dir$, MkDir and Kill of the VBA library.
'
' NOTES
'   Typical call, in the Immediate window (INSTALLATION.md, Recording an
'   Excel run):
'
'       TEST_Examples.SaveExampleEvidence "<folder>"
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.SaveExampleEvidence"

    Dim Leftover        As Variant  'File name an earlier run left in Folder

'------------------------------------------------------------------------------
' START A NEW EVIDENCE FOLDER
'------------------------------------------------------------------------------
    'Create the folder when it does not exist
        If VBA.Len(Dir$(Folder, vbDirectory)) = 0 Then
            MkDir Folder
        End If
    'Remove the logs and host.txt of an earlier run
        For Each Leftover In Array("examples.log", "injected.log", "skipped.log", "host.txt")
            If VBA.Len(Dir$(Folder & "\" & Leftover)) > 0 Then
                Kill Folder & "\" & Leftover
            End If
        Next Leftover

'------------------------------------------------------------------------------
' RUN AND SAVE
'------------------------------------------------------------------------------
    'Run each mode and save its log
        SaveRun "NORMAL", Folder & "\examples.log"
        SaveRun "INJECTED_FAILURE", Folder & "\injected.log"
        SaveRun "SKIPPED_CASE", Folder & "\skipped.log"

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
'------------------------------------------------------------------------------
'                         RUN EVERY CASE OF THE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Run every case in a fixed order between BeginRun and EndRun.
'
' WHY THIS EXISTS
'   The three entry points and SaveRun share one runner; only Mode differs.
'
' INPUTS
'   Mode
'     NORMAL, INJECTED_FAILURE or SKIPPED_CASE.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   1. Stores Mode for the cases and starts the harness run.
'   2. Runs the cases in order; in SKIPPED_CASE mode the last one is not
'      run.
'   3. Clears the mode and ends the harness run with the error that
'      escaped a case, if any.
'
' ERROR POLICY
'   An error escaping a case handler is passed to EndRun, which reports it
'   and re-raises it unchanged. BeginRun's refusal propagates directly.
'
' DEPENDENCIES
'   TEST_Harness.BeginRun and TEST_Harness.EndRun; the cases of this module.
'
' NOTES
'   Each case is called on a line of its own: tools/test_vba_harness.py and
'   tools/check_excel_run.py read the calls from this body.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.RunSuite"

    Dim ErrNumber       As Long     'Runner error number; 0 for none
    Dim ErrSource       As String   'Runner error source
    Dim ErrDescription  As String   'Runner error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Store the mode the cases read
        mMode = Mode
    'Start the run; a refused start propagates before the handler is set
        TEST_Harness.BeginRun SUITE_NAME, Mode, EXPECTED_CASES, EXPECTED_ASSERTIONS
    'Pass an error escaping a case to EndRun
        On Error GoTo RunFailed

'------------------------------------------------------------------------------
' RUN CASES
'------------------------------------------------------------------------------
    'Run every case in a fixed order
        CaseExactDouble
        CaseAbsoluteTolerance
        CaseRelativeTolerance
        CaseExactText
        CaseErrorValue
        CaseRaisedError
        CaseCleanupRestoresState
        CaseCleanupKeepsPrimaryError
    'Skip the last case in SKIPPED_CASE mode
        If mMode <> "SKIPPED_CASE" Then
            CaseCleanupOnlyFailure
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
Finish:
    'Restore normal error handling and clear the mode
        On Error GoTo 0
        mMode = vbNullString
    'End the run with the captured error, if any
        TEST_Harness.EndRun ErrNumber, ErrSource, ErrDescription
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
RunFailed:
    'Capture the error that escaped a case
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'End the run through the exit path
        Resume Finish

End Sub

Private Sub SaveRun( _
    ByVal Mode As String, _
    ByVal Path As String)
'
'------------------------------------------------------------------------------
'                        RUN ONE MODE AND SAVE ITS LOG
'------------------------------------------------------------------------------
' PURPOSE
'   Run the suite in Mode, save its log to Path and print how the run ended.
'
' INPUTS
'   Mode
'     NORMAL, INJECTED_FAILURE or SKIPPED_CASE.
'
'   Path
'     Full path of the log file to write.
'
' ERROR POLICY
'   The run's own error is captured and printed after the log is saved; an
'   error of SaveLog propagates.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TEST_Examples.SaveRun"

    Dim RunErrorNumber          As Long     'Error that ended the run; 0 for none
    Dim RunErrorDescription     As String   'Its description

'------------------------------------------------------------------------------
' RUN AND SAVE
'------------------------------------------------------------------------------
    'Run the suite and capture the error that ended it
        On Error Resume Next
        RunSuite Mode
        RunErrorNumber = Err.Number
        RunErrorDescription = Err.Description
        On Error GoTo 0
    'Save the log of the run
        TEST_Harness.SaveLog Path
    'Print how the run ended
        Debug.Print "SAVED=" & Path & "; mode=" & Mode & "; error=" & VBA.CStr(RunErrorNumber) & _
            IIf(RunErrorNumber = 0, vbNullString, " " & RunErrorDescription)

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
'------------------------------------------------------------------------------
'                         CHECK EXACT DOUBLE EQUALITY
'------------------------------------------------------------------------------
' PURPOSE
'   Exact equality on a result that is exact in binary. 1 assertion.
'
' NOTES
'   In INJECTED_FAILURE mode the expected value is wrong on purpose.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.CaseExactDouble"

    Dim Expected        As Double   'Expected sum; wrong on purpose when injected

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error as an error of the case
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Start the case
        TEST_Harness.BeginCase "exact-double"
    'Expect a wrong value when a failure is injected
        Expected = 0.75
        If mMode = "INJECTED_FAILURE" Then
            Expected = 0.76
        End If
    'Compare the sum exactly
        TEST_Harness.AssertEqualDouble "half-plus-quarter", Expected, 0.5 + 0.25

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseAbsoluteTolerance()

'
'------------------------------------------------------------------------------
'                         CHECK THE ABSOLUTE TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   The absolute bound: 0.1 + 0.2 is not exactly 0.3 in binary, and near
'   zero only the absolute bound counts. 2 assertions.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.CaseAbsoluteTolerance"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error as an error of the case
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Start the case
        TEST_Harness.BeginCase "tolerance-absolute"
    'Accept a sum that is not exact in binary
        TEST_Harness.AssertNear "tenth-plus-fifth", 0.3, 0.1 + 0.2, 0.000000000000001, 0#
    'Near zero, only the absolute bound counts
        TEST_Harness.AssertNear "near-zero", 0#, 1E-20, 0.000000000001, 0.000000001

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseRelativeTolerance()

'
'------------------------------------------------------------------------------
'                         CHECK THE RELATIVE TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   The relative bound governs large values, the bound is inclusive, and a
'   value outside both bounds is rejected. 3 assertions.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.CaseRelativeTolerance"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error as an error of the case
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Start the case
        TEST_Harness.BeginCase "tolerance-relative"
    'The relative bound governs a large value
        TEST_Harness.AssertNear "large-value", 1000000000#, 1000000000.1, 0.000001, 0.000000001
    'The bound is inclusive
        TEST_Harness.AssertTrue "bound-is-inclusive", _
            TEST_Harness.WithinTolerance(2#, 2.5, 0.5, 0#), "|2.5 - 2| = 0.5 must be within 0.5"
    'A value outside both bounds is rejected
        TEST_Harness.AssertTrue "outside-both-bounds", _
            Not TEST_Harness.WithinTolerance(1#, 1.1, 0.01, 0.01), _
            "|1.1 - 1| = 0.1 must exceed max(0.01, 0.01 * 1)"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseExactText()

'
'------------------------------------------------------------------------------
'                          CHECK EXACT TEXT EQUALITY
'------------------------------------------------------------------------------
' PURPOSE
'   Text is compared exactly, case included. 2 assertions.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.CaseExactText"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error as an error of the case
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Start the case
        TEST_Harness.BeginCase "text-exact"
    'Compare a built text exactly
        TEST_Harness.AssertEqualText "concatenation", "CASE-1", "CASE" & "-" & VBA.CStr(1)
    'Letter case matters in a binary comparison
        TEST_Harness.AssertTrue "case-matters", _
            VBA.StrComp("a", "A", vbBinaryCompare) <> 0, "a and A must differ"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseErrorValue()

'
'------------------------------------------------------------------------------
'                        CHECK AN EXPECTED ERROR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   An expected worksheet error value, returned rather than raised.
'   1 assertion.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.CaseErrorValue"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error as an error of the case
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Start the case
        TEST_Harness.BeginCase "error-value"
    'A zero denominator returns #DIV/0!
        TEST_Harness.AssertErrorValue "zero-denominator", xlErrDiv0, ExampleRatio(1#, 0#)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseRaisedError()

'
'------------------------------------------------------------------------------
'                        CHECK AN EXPECTED RAISED ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   An expected raised error, captured on the line after the call, and the
'   same call with a valid input raising nothing. 2 assertions.
'
' WHY THIS EXISTS
'   It shows how a case checks a procedure that raises: the error number
'   is read on the line right after the call.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   1. Calls ExampleSquareRoot with -1 and checks that it raised
'      ERR_EXAMPLE_INVALID.
'   2. Calls ExampleSquareRoot with 4 and checks that it raised nothing.
'
' ERROR POLICY
'   Each checked call runs under On Error Resume Next; the case handler is
'   restored right after. Any other error is passed to
'   TEST_Harness.RecordCaseError.
'
' DEPENDENCIES
'   TEST_Harness; ExampleSquareRoot.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.CaseRaisedError"

    Dim Result          As Double   'Value of the checked square root
    Dim RaisedError     As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error as an error of the case
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Start the case
        TEST_Harness.BeginCase "raised-error"
    'Call with a negative input and capture the error it raises
        On Error Resume Next
        Result = ExampleSquareRoot(-1#)
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'The negative input raised ERR_EXAMPLE_INVALID
        TEST_Harness.AssertRaised "negative-input", ERR_EXAMPLE_INVALID, RaisedError
    'Call with a valid input and capture any error it raises
        On Error Resume Next
        Result = ExampleSquareRoot(4#)
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'The valid input raised nothing
        TEST_Harness.AssertRaised "valid-input", 0, RaisedError

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupRestoresState()

'
'------------------------------------------------------------------------------
'                      CHECK THAT CLEANUP RESTORES STATE
'------------------------------------------------------------------------------
' PURPOSE
'   A successful operation restores what it changed and reports no cleanup
'   failure. 2 assertions.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.CaseCleanupRestoresState"

    Dim Details         As String   'Cleanup failures reported by the operation
    Dim SavedScreen     As Boolean  'ScreenUpdating before the operation

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error as an error of the case
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Start the case
        TEST_Harness.BeginCase "cleanup-restores-state"
    'Capture ScreenUpdating, then run an operation that succeeds
        SavedScreen = Application.ScreenUpdating
        RunExampleOperation False, False, Details
    'No cleanup failure is reported
        TEST_Harness.AssertEqualText "no-cleanup-failure", vbNullString, Details
    'ScreenUpdating is back to its value before the operation
        TEST_Harness.AssertTrue "screen-updating-restored", _
            Application.ScreenUpdating = SavedScreen, "ScreenUpdating differs after the operation"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupKeepsPrimaryError()

'
'------------------------------------------------------------------------------
'                  CHECK THAT CLEANUP KEEPS THE PRIMARY ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   When the operation and its cleanup both fail, the caller receives the
'   primary error, and the cleanup failure is still reported. 3 assertions.
'
' WHY THIS EXISTS
'   It is the case of docs/VBA_HOUSE_STYLE.md (Operation and cleanup) for
'   a primary error with failed cleanup.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   1. Runs RunExampleOperation with a failing operation and failing
'      restores, and captures the error it raises.
'   2. Checks that the error is ERR_EXAMPLE_OPERATION.
'   3. Checks that the details name both restore failures.
'
' ERROR POLICY
'   The operation runs under On Error Resume Next; the case handler is
'   restored right after. Any other error is passed to
'   TEST_Harness.RecordCaseError.
'
' DEPENDENCIES
'   TEST_Harness; RunExampleOperation.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.CaseCleanupKeepsPrimaryError"

    Dim Details         As String   'Cleanup failures reported by the operation
    Dim RaisedError     As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error as an error of the case
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Start the case
        TEST_Harness.BeginCase "cleanup-keeps-primary-error"
    'Run an operation whose work and cleanup both fail; capture its error
        On Error Resume Next
        RunExampleOperation True, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'The caller receives the primary error
        TEST_Harness.AssertRaised "primary-error", ERR_EXAMPLE_OPERATION, RaisedError
    'Both restore failures are reported
        TEST_Harness.AssertTrue "calculation-failure-reported", _
            VBA.InStr(1, Details, "Calculation not restored", vbBinaryCompare) > 0, _
            "details='" & Details & "'"
        TEST_Harness.AssertTrue "screen-failure-reported", _
            VBA.InStr(1, Details, "ScreenUpdating not restored", vbBinaryCompare) > 0, _
            "details='" & Details & "'"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error
        TEST_Harness.RecordCaseError

End Sub

Private Sub CaseCleanupOnlyFailure()

'
'------------------------------------------------------------------------------
'                     CHECK A FAILURE OF THE CLEANUP ONLY
'------------------------------------------------------------------------------
' PURPOSE
'   When only the cleanup fails, the operation raises ERR_CLEANUP_FAILED,
'   never success. 2 assertions.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.CaseCleanupOnlyFailure"

    Dim Details         As String   'Cleanup failures reported by the operation
    Dim RaisedError     As Long     'Err.Number right after the call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Record an unexpected error as an error of the case
        On Error GoTo CaseFailed

'------------------------------------------------------------------------------
' RUN CASE
'------------------------------------------------------------------------------
    'Start the case
        TEST_Harness.BeginCase "cleanup-only-failure"
    'Run an operation whose cleanup alone fails; capture its error
        On Error Resume Next
        RunExampleOperation False, True, Details
        RaisedError = Err.Number
        On Error GoTo CaseFailed
    'The operation raises ERR_CLEANUP_FAILED
        TEST_Harness.AssertRaised "cleanup-error", ERR_CLEANUP_FAILED, RaisedError
    'The cleanup failure is reported
        TEST_Harness.AssertTrue "failure-reported", _
            VBA.Len(Details) > 0, "no cleanup detail returned"

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
CaseFailed:
    'Record the unexpected error
        TEST_Harness.RecordCaseError

End Sub

'
'------------------------------------------------------------------------------
'
'                        EXAMPLE PROCEDURES UNDER TEST
'
'------------------------------------------------------------------------------
'

Private Function ExampleRatio( _
    ByVal Numerator As Double, _
    ByVal Denominator As Double) As Variant
'
'------------------------------------------------------------------------------
'                      DIVIDE OR RETURN A WORKSHEET ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Numerator / Denominator, or the worksheet error #DIV/0! for a zero
'   denominator, as a worksheet function returns errors (house style).
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return #DIV/0! for a zero denominator, the ratio otherwise
        If Denominator = 0# Then
            ExampleRatio = CVErr(xlErrDiv0)
        Else
            ExampleRatio = Numerator / Denominator
        End If

End Function

Private Function ExampleSquareRoot( _
    ByVal Value As Double) As Double
'
'------------------------------------------------------------------------------
'                          TAKE A CHECKED SQUARE ROOT
'------------------------------------------------------------------------------
' PURPOSE
'   Square root of Value >= 0; a negative Value raises ERR_EXAMPLE_INVALID,
'   as a core procedure raises (house style).
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a negative value
        If Value < 0# Then
            Err.Raise ERR_EXAMPLE_INVALID, "TEST_Examples.ExampleSquareRoot", _
                "Value must be >= 0; got " & VBA.CStr(Value) & "."
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the square root
        ExampleSquareRoot = Sqr(Value)

End Function

Private Sub RunExampleOperation( _
    ByVal FailOperation As Boolean, _
    ByVal FailRestore As Boolean, _
    ByRef CleanupDetails As String)
'
'------------------------------------------------------------------------------
'                RUN AN OPERATION THAT RESTORES WHAT IT CHANGES
'------------------------------------------------------------------------------
' PURPOSE
'   Reference implementation of the cleanup pattern of
'   docs/VBA_HOUSE_STYLE.md (Operation and cleanup): capture, change, restore
'   in one block on both paths, preserve the primary error, report every
'   cleanup failure.
'
' WHY THIS EXISTS
'   docs/VBA_HOUSE_STYLE.md names it as the reference implementation of the
'   pattern; the cleanup cases of this suite test it.
'
' INPUTS
'   FailOperation
'     Raise ERR_EXAMPLE_OPERATION inside the work.
'
'   FailRestore
'     Test seam; each restore helper restores its setting and then reports
'     a simulated failure, so that the failure paths can be tested without
'     leaving Excel changed.
'
'   CleanupDetails
'     Receives the cleanup failures joined by "; ", empty on success; set
'     also when an error is raised.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   1. Clears CleanupDetails.
'   2. Captures Application.Calculation and Application.ScreenUpdating,
'      then turns ScreenUpdating off and Calculation to manual.
'   3. Does the work: raises ERR_EXAMPLE_OPERATION when FailOperation.
'   4. On the success and the error path, restores both settings when they
'      were captured, each through its own helper, and joins the failures.
'   5. Re-raises the primary error; otherwise raises ERR_CLEANUP_FAILED
'      when a restore failed.
'
' ERROR POLICY
'   Re-raises the primary error after cleanup; when only the cleanup failed,
'   raises ERR_CLEANUP_FAILED with the details.
'
' DEPENDENCIES
'   Excel Application; RestoreCalculation, RestoreScreenUpdating and
'   AppendDetail.
'
' NOTES
'   The sketch in docs/VBA_HOUSE_STYLE.md (Operation and cleanup) shows the
'   same pattern without the test seams.
'
' STATE OWNERSHIP
'   Changes Application.Calculation to manual and ScreenUpdating to False,
'   and restores both to the values it captured.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "TEST_Examples.RunExampleOperation"

    Dim SavedCalculation    As XlCalculation    'Calculation mode before the operation
    Dim SavedScreen         As Boolean          'ScreenUpdating before the operation
    Dim StateCaptured       As Boolean          'True once both were read
    Dim ErrNumber           As Long             'Primary error number; 0 for none
    Dim ErrSource           As String           'Primary error source
    Dim ErrDescription      As String           'Primary error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start with no cleanup failure
        CleanupDetails = vbNullString
    'Send an error of the operation to the cleanup block
        On Error GoTo Failed

'------------------------------------------------------------------------------
' CAPTURE AND CHANGE
'------------------------------------------------------------------------------
    'Capture the settings this procedure changes
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
    'Change only what the operation needs
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

'------------------------------------------------------------------------------
' WORK
'------------------------------------------------------------------------------
    'Fail the operation when asked
        If FailOperation Then
            Err.Raise ERR_EXAMPLE_OPERATION, "TEST_Examples.RunExampleOperation", _
                "Simulated failure of the operation."
        End If

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
CleanUp:
    'Restore normal error handling; each restore helper contains its own errors
        On Error GoTo 0
    'Restore each captured setting and collect the failures
        If StateCaptured Then
            CleanupDetails = AppendDetail(RestoreCalculation(SavedCalculation, FailRestore), _
                RestoreScreenUpdating(SavedScreen, FailRestore))
        End If
    'Re-raise the primary error
        If ErrNumber <> 0 Then
            Err.Raise ErrNumber, ErrSource, ErrDescription
        End If
    'Never report success when only the cleanup failed
        If VBA.Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "TEST_Examples.RunExampleOperation", CleanupDetails
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Capture the primary error
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Restore through the cleanup block
        Resume CleanUp

End Sub

Private Function RestoreCalculation( _
    ByVal SavedCalculation As XlCalculation, _
    ByVal SimulateFailure As Boolean) As String
'
'------------------------------------------------------------------------------
'                         RESTORE THE CALCULATION MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.Calculation; contain any error.
'
' INPUTS
'   SavedCalculation
'     Calculation mode to restore.
'
'   SimulateFailure
'     Test seam: report a simulated failure after restoring.
'
' RETURNS
'   Empty on success, otherwise what failed.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.RestoreCalculation"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain any error of the restore
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restore the captured calculation mode
        Application.Calculation = SavedCalculation
    'Simulate a failure when asked
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreCalculation", "simulated"
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return what failed
        RestoreCalculation = "Calculation not restored: " & Err.Description

End Function

Private Function RestoreScreenUpdating( _
    ByVal SavedScreen As Boolean, _
    ByVal SimulateFailure As Boolean) As String
'
'------------------------------------------------------------------------------
'                           RESTORE SCREEN UPDATING
'------------------------------------------------------------------------------
' PURPOSE
'   Restore Application.ScreenUpdating; contain any error.
'
' INPUTS
'   SavedScreen
'     ScreenUpdating value to restore.
'
'   SimulateFailure
'     Test seam: report a simulated failure after restoring.
'
' RETURNS
'   Empty on success, otherwise what failed.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Examples.RestoreScreenUpdating"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Contain any error of the restore
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restore the captured ScreenUpdating value
        Application.ScreenUpdating = SavedScreen
    'Simulate a failure when asked
        If SimulateFailure Then
            Err.Raise ERR_EXAMPLE_RESTORE, "TEST_Examples.RestoreScreenUpdating", "simulated"
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return what failed
        RestoreScreenUpdating = "ScreenUpdating not restored: " & Err.Description

End Function

Private Function AppendDetail( _
    ByVal First As String, _
    ByVal Second As String) As String
'
'------------------------------------------------------------------------------
'                            JOIN TWO FAILURE TEXTS
'------------------------------------------------------------------------------
' PURPOSE
'   Join two failure texts with "; ", skipping empty ones.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Skip an empty text; join two texts with "; "
        If VBA.Len(First) = 0 Then
            AppendDetail = Second
        ElseIf VBA.Len(Second) = 0 Then
            AppendDetail = First
        Else
            AppendDetail = First & "; " & Second
        End If

End Function
