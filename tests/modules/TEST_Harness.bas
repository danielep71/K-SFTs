Attribute VB_Name = "TEST_Harness"

Option Explicit
Option Private Module

'
'==============================================================================
' MODULE: TEST_Harness
'------------------------------------------------------------------------------
' PURPOSE
'   Deterministic regression harness for K-SFTs's VBA test suites: one
'   run per suite, stable case identifiers, exact and tolerance assertions,
'   expected error values and raised errors, expected case and assertion
'   counts, and a complete key=value log. It holds no SFT calculation.
'
' WHY THIS EXISTS
'   The suites share one lifecycle, one set of assertions and one log
'   format, so that every run is reported the same way and its saved logs
'   can be checked without Excel (tools/check_excel_run.py).
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   One run at a time:
'     - BeginRun refuses an unsafe start, then claims the run
'     - BeginCase starts each case; the assertions count and log outcomes
'     - RecordCaseError records an error that ended a case
'     - EndRun checks completeness and the host state, then raises the
'       outcome
'   Counters and the log stay readable after EndRun, until the next
'   BeginRun or ResetTests.
'
' ERROR POLICY
'   Assertion failures are recorded and the run continues. EndRun re-raises
'   an unexpected runner error unchanged after cleanup and reporting, or
'   raises ERR_TEST_FAILED for any other failed outcome. A run that starts
'   while another is active, or while recovery is required, is refused with
'   ERR_TEST_DIRTY_START. Misuse of the lifecycle raises ERR_TEST_USAGE.
'
' DEPENDENCIES
'   The VBA and Excel object models only: no reference, workbook fixture,
'   worksheet or external framework.
'
' NOTES
'   Called by the suites in tests/modules/, for example
'   TEST_Examples.RunExampleTests; SaveLog and ResetTests are also typed in
'   the Immediate window. It has no suite of its own.
'
'   The test-only error numbers are kept apart from the production range in
'   use.
'
' PUBLIC SURFACE
'   Suite lifecycle: BeginRun, BeginCase, RecordCaseError, EndRun.
'   Assertions: AssertEqualDouble, AssertNear, AssertEqualText, AssertTrue,
'   AssertErrorValue, AssertRaised; WithinTolerance is the comparison rule.
'   Log: Say, SaveLog, LastResult. Recovery: ResetTests. Option Private
'   Module keeps all of them out of the workbook's external surface.
'
' STATE OWNERSHIP
'   Owns the run counters, the case identifiers, the failure records, the
'   retained log and the active-run and recovery-required flags. Reads four
'   Application properties to prove a run left them unchanged; never changes
'   them, and never touches a workbook or worksheet.
'
' LOG FORMAT
'   One record per line, KEY=value; fields separated by "; ". The keys and
'   their order are documented in tests/README.md (Log format).
'
' COMPATIBILITY
'   Excel VBA on Windows. The log records the host, its version and build,
'   the Office bitness and the VBA generation.
'
' AUTHOR
'   Daniele Penza
'
' UPDATED
'   2026-10-29
'==============================================================================

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    'Test-only error numbers, kept apart from the production range in use
    Public Const ERR_TEST_DIRTY_START   As Long = vbObjectError + 4064  'Run refused: active or recovery required
    Public Const ERR_TEST_FAILED        As Long = vbObjectError + 4065  'Run finished without PASS
    Public Const ERR_TEST_USAGE         As Long = vbObjectError + 4066  'Lifecycle used out of order

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'One run at a time. Counters and the log stay readable after EndRun,
    'until the next BeginRun or ResetTests
    Private mRunActive          As Boolean          'True from BeginRun to EndRun
    Private mRecoveryRequired   As Boolean          'True until ResetTests after an unsafe end
    Private mRecoveryReason     As String           'Why recovery is required

    Private mSuite              As String           'Name of the running or last suite
    Private mExpectedCases      As Long             'Cases a complete run starts
    Private mExpectedAssertions As Long             'Assertions a complete run evaluates
    Private mCaseCount          As Long             'Cases started in this run
    Private mAssertionCount     As Long             'Assertions evaluated in this run
    Private mFailureCount       As Long             'Failures of any kind in this run
    Private mErrorCount         As Long             'Unexpected errors in this run

    Private mCurrentCase        As String           'Identifier of the running case
    Private mCaseIds            As Collection       'Case identifiers seen in this run
    Private mPrimaryFailure     As String           'First failure recorded in this run
    Private mLastResult         As String           'RESULT record of the last finished run
    Private mLog                As Collection       'Log since the last BeginRun or ResetTests

    Private mSnapshotValid      As Boolean          'True once all four properties were read
    Private mInitialCalculation As XlCalculation    'Calculation at BeginRun
    Private mInitialAlerts      As Boolean          'DisplayAlerts at BeginRun
    Private mInitialEvents      As Boolean          'EnableEvents at BeginRun
    Private mInitialScreen      As Boolean          'ScreenUpdating at BeginRun


'
'------------------------------------------------------------------------------
'
'                               SUITE LIFECYCLE
'
'------------------------------------------------------------------------------
'

Public Sub BeginRun( _
    ByVal SuiteName As String, _
    ByVal Mode As String, _
    ByVal ExpectedCases As Long, _
    ByVal ExpectedAssertions As Long)
'
'------------------------------------------------------------------------------
'                              START A SUITE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Starts one suite run: refuses an unsafe start, resets the run state,
'   records the host state to compare at the end and writes the log header.
'
' WHY THIS EXISTS
'   An interrupted run or an unverified cleanup must be acknowledged with
'   ResetTests before a new run may overwrite its state, and EndRun needs
'   the host state as it was before the suite ran.
'
' INPUTS
'   SuiteName
'     Stable suite identifier, written in the RUN and RESULT records
'
'   Mode
'     NORMAL, or the name of a deliberate failure mode of the suite
'
'   ExpectedCases
'     Cases a complete run starts, at least 1
'
'   ExpectedAssertions
'     Assertions a complete run evaluates, at least 1
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - treats a run that is still active as interrupted
'     - when recovery is required, writes a NOT_RUN RESULT record and a
'       RECOVERY record, then refuses the start
'     - rejects expected counts below 1
'     - resets the run state, starts a new log and claims the run
'     - reads Calculation, DisplayAlerts, EnableEvents and ScreenUpdating
'     - writes the RUN and ENVIRONMENT records
'
' ERROR POLICY
'   Raises ERR_TEST_DIRTY_START, after logging why, when a run is active or
'   recovery is required; raises ERR_TEST_USAGE for counts below 1. A host
'   property that cannot be read leaves the snapshot invalid, which fails
'   cleanup at EndRun; it does not stop the run.
'
' DEPENDENCIES
'   Excel Application object; ResetRunState, Say, EnvironmentSummary
'
' NOTES
'   A project reset (End or Reset in the VBE) clears the module state, so a
'   run stopped that way is not seen as interrupted.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "TEST_Harness.BeginRun"

'------------------------------------------------------------------------------
' REFUSE AN UNSAFE START
'------------------------------------------------------------------------------
    'Treat a run that is still active as interrupted
        If mRunActive Then
            mRecoveryRequired = True
            mRecoveryReason = "run of " & mSuite & " still active (interrupted?)"
        End If
    'Log the refused start, then refuse it until ResetTests acknowledges it
        If mRecoveryRequired Then
            Say "RESULT=FAIL; suite=" & SuiteName & "; operation=NOT_RUN" & _
                "; completeness=INCOMPLETE; cases=0; assertions=0; failures=0" & _
                "; cleanup=NOT_RUN; recovery=REQUIRED"
            Say "RECOVERY=REQUIRED; reason=" & mRecoveryReason & _
                "; action=check Excel state, then run TEST_Harness.ResetTests"
            Err.Raise ERR_TEST_DIRTY_START, "TEST_Harness.BeginRun", _
                "Recovery required: " & mRecoveryReason & ". Run TEST_Harness.ResetTests."
        End If

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject expected counts below 1
        If ExpectedCases < 1 Or ExpectedAssertions < 1 Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.BeginRun", _
                "Expected counts must be at least 1; got cases=" & VBA.CStr(ExpectedCases) & _
                ", assertions=" & VBA.CStr(ExpectedAssertions) & "."
        End If

'------------------------------------------------------------------------------
' RESET AND CLAIM THE RUN
'------------------------------------------------------------------------------
    'Return the run state to its idle values and start a new log
        ResetRunState
        Set mLog = New Collection
    'Store the suite and its expected counts
        mSuite = SuiteName
        mExpectedCases = ExpectedCases
        mExpectedAssertions = ExpectedAssertions
    'Claim the run
        mRunActive = True

'------------------------------------------------------------------------------
' SNAPSHOT HOST STATE
'------------------------------------------------------------------------------
    'Send a failed read to the snapshot handler
        On Error GoTo SnapshotFailed
    'Read the four properties that EndRun compares
        mInitialCalculation = Application.Calculation
        mInitialAlerts = Application.DisplayAlerts
        mInitialEvents = Application.EnableEvents
        mInitialScreen = Application.ScreenUpdating
    'Mark the snapshot valid only after all four reads, so that cleanup never
    'compares a value that was not read
        mSnapshotValid = True

'------------------------------------------------------------------------------
' WRITE THE LOG HEADER
'------------------------------------------------------------------------------
WriteHeader:
    'Restore normal error handling
        On Error GoTo 0
    'Write the RUN record
        Say "RUN=" & SuiteName & "; mode=" & Mode & _
            "; expected_cases=" & VBA.CStr(ExpectedCases) & _
            "; expected_assertions=" & VBA.CStr(ExpectedAssertions) & _
            "; started=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    'Write the ENVIRONMENT record
        Say "ENVIRONMENT=" & EnvironmentSummary()

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the snapshot handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
SnapshotFailed:
    'Leave the snapshot invalid, which fails cleanup at EndRun
        mSnapshotValid = False
    'Write the log header all the same
        Resume WriteHeader

End Sub

Public Sub BeginCase(ByVal CaseId As String)

'
'------------------------------------------------------------------------------
'                                 START A CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Counts and logs the start of one case.
'
' WHY THIS EXISTS
'   Logs are compared case by case, so every case needs an identifier that
'   is well formed and used once per run.
'
' INPUTS
'   CaseId
'     Lower-case kebab-case identifier, unique within the run and stable
'     across runs, so that logs can be compared case by case
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts the case, makes it the current case and writes its CASE record.
'   A malformed or repeated identifier is then recorded as a failure; a new
'   one is remembered for the rest of the run.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A malformed or repeated identifier
'   is recorded as a failure; the case still counts and runs.
'
' DEPENDENCIES
'   RequireActiveRun, Say, IsCaseId, IsKnownCase, RecordFailure
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REGISTER THE CASE
'------------------------------------------------------------------------------
    'Refuse a case outside a run
        RequireActiveRun "BeginCase"
    'Count the case and make it the current one
        mCaseCount = mCaseCount + 1
        mCurrentCase = CaseId
    'Write the CASE record
        Say "CASE=" & CaseId

'------------------------------------------------------------------------------
' CHECK THE IDENTIFIER
'------------------------------------------------------------------------------
    'Record a malformed or repeated identifier; remember a new one
        If Not IsCaseId(CaseId) Then
            RecordFailure "harness.case-id", "malformed case identifier '" & CaseId & "'"
        ElseIf IsKnownCase(CaseId) Then
            RecordFailure "harness.case-id", "repeated case identifier '" & CaseId & "'"
        Else
            mCaseIds.Add CaseId, CaseId
        End If

End Sub

Public Sub RecordCaseError()

'
'------------------------------------------------------------------------------
'                             RECORD A CASE ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Records the error that ended the running case, so that the remaining
'   cases still run.
'
' WHY THIS EXISTS
'   An error in one case must not stop the suite; it is counted and logged
'   as a failure of that case instead.
'
' INPUTS
'   None. Reads the Err object before anything else can reset it.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Captures the number, source and description of the error, counts an
'   unexpected error, writes the ERROR record and records the failure
'   <case>.unexpected-error. Evaluates no assertion.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run.
'
' DEPENDENCIES
'   Err object; RequireActiveRun, Say, RecordFailure
'
' NOTES
'   Call it first thing in the case's error handler.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "TEST_Harness.RecordCaseError"

    Dim ErrorNumber         As Long     'Number of the case error
    Dim ErrorSource         As String   'Source of the case error
    Dim ErrorDescription    As String   'Description of the case error

'------------------------------------------------------------------------------
' CAPTURE THE ERROR
'------------------------------------------------------------------------------
    'Capture the case error before any other call can reset it
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description

'------------------------------------------------------------------------------
' RECORD THE ERROR
'------------------------------------------------------------------------------
    'Refuse a call outside a run
        RequireActiveRun "RecordCaseError"
    'Count the unexpected error
        mErrorCount = mErrorCount + 1
    'Write the ERROR record
        Say "ERROR=" & mCurrentCase & "; number=" & VBA.CStr(ErrorNumber) & _
            "; source=" & ErrorSource & "; description=" & ErrorDescription
    'Record the failure of the case
        RecordFailure mCurrentCase & ".unexpected-error", _
            "number=" & VBA.CStr(ErrorNumber) & "; source=" & ErrorSource

End Sub

Public Sub EndRun( _
    ByVal RunErrorNumber As Long, _
    ByVal RunErrorSource As String, _
    ByVal RunErrorDescription As String)
'
'------------------------------------------------------------------------------
'                              FINISH A SUITE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Finishes the run: checks completeness, releases the run, verifies the
'   host state, writes the summary and raises the outcome.
'
' WHY THIS EXISTS
'   What ran, completeness, cleanup and recovery are judged and reported
'   separately, and a runner error must not be replaced by a later cleanup
'   failure.
'
' INPUTS
'   RunErrorNumber
'     Number of the error that escaped the suite's runner, or 0 when none
'     did
'
'   RunErrorSource
'     Source of that error, or an empty string when none escaped
'
'   RunErrorDescription
'     Description of that error, or an empty string when none escaped
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   In order:
'     - counts, logs and records a runner error
'     - decides the operation outcome
'     - records harness.completeness when an expected count was not met
'     - compares the host state, releases the run and requires recovery
'       when cleanup failed
'     - writes the CASES, ASSERTIONS, FAILURES, PRIMARY_FAILURE (when any),
'       CLEANUP, RECOVERY and RESULT records
'     - raises the outcome
'
' ERROR POLICY
'   Re-raises the runner error unchanged, after reporting, so that cleanup
'   never replaces the primary failure; otherwise raises ERR_TEST_FAILED
'   unless RESULT=PASS. Raises ERR_TEST_USAGE outside a run.
'
' DEPENDENCIES
'   RequireActiveRun, Say, RecordFailure, VerifyHostState
'
' NOTES
'   None
'
' OUTCOMES
'   operation: what ran: PASS, FAIL (an assertion or a case identifier
'   failed) or ERROR (an unexpected error occurred). completeness: COMPLETE
'   only when both expected counts were met exactly. cleanup: PASS when the
'   host state is unchanged. recovery: REQUIRED when the host state changed
'   or could not be checked, or when a start was refused while this run was
'   active. RESULT=PASS needs operation PASS, COMPLETE, cleanup PASS and no
'   recovery.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Harness.EndRun"

    Dim CleanupDetail   As String   'Reason for the cleanup outcome
    Dim CleanupPassed   As Boolean  'True when the host state is unchanged
    Dim Completed       As Boolean  'True when both expected counts were met
    Dim Operation       As String   'PASS, FAIL or ERROR
    Dim Verdict         As String   'PASS or FAIL

'------------------------------------------------------------------------------
' COLLECT THE OUTCOMES
'------------------------------------------------------------------------------
    'Refuse a call outside a run
        RequireActiveRun "EndRun"
    'Attribute the failures recorded from here on to the run itself
        mCurrentCase = "run"
    'Count, log and record an error that escaped the runner
        If RunErrorNumber <> 0 Then
            mErrorCount = mErrorCount + 1
            Say "ERROR=run; number=" & VBA.CStr(RunErrorNumber) & _
                "; source=" & RunErrorSource & "; description=" & RunErrorDescription
            RecordFailure "runner.unexpected-error", _
                "number=" & VBA.CStr(RunErrorNumber) & "; source=" & RunErrorSource
        End If
    'The operation outcome covers what ran; completeness and cleanup are
    'judged separately below, so that each is reported on its own
        If mErrorCount > 0 Then
            Operation = "ERROR"
        ElseIf mFailureCount > 0 Then
            Operation = "FAIL"
        Else
            Operation = "PASS"
        End If
    'Exact counts: a skipped case or assertion, or an extra one, means the
    'run did not execute the suite as written
        Completed = (mCaseCount = mExpectedCases) And _
                    (mAssertionCount = mExpectedAssertions)
        If Not Completed Then
            RecordFailure "harness.completeness", _
                "expected cases=" & VBA.CStr(mExpectedCases) & _
                ", assertions=" & VBA.CStr(mExpectedAssertions) & _
                "; observed cases=" & VBA.CStr(mCaseCount) & _
                ", assertions=" & VBA.CStr(mAssertionCount)
        End If

'------------------------------------------------------------------------------
' VERIFY CLEANUP
'------------------------------------------------------------------------------
    'Compare the host state with the snapshot, then release the run
        CleanupPassed = VerifyHostState(CleanupDetail)
        mRunActive = False
    'Require recovery when the host state changed or could not be checked
        If Not CleanupPassed Then
            mRecoveryRequired = True
            mRecoveryReason = CleanupDetail
        End If

'------------------------------------------------------------------------------
' DECIDE THE VERDICT
'------------------------------------------------------------------------------
    'Pass only a complete, clean operation PASS that needs no recovery
        If Operation = "PASS" And Completed And CleanupPassed And Not mRecoveryRequired Then
            Verdict = "PASS"
        Else
            Verdict = "FAIL"
        End If

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Write the counts
        Say "CASES=" & VBA.CStr(mCaseCount) & "; expected=" & VBA.CStr(mExpectedCases)
        Say "ASSERTIONS=" & VBA.CStr(mAssertionCount) & _
            "; expected=" & VBA.CStr(mExpectedAssertions)
        Say "FAILURES=" & VBA.CStr(mFailureCount) & "; errors=" & VBA.CStr(mErrorCount)
    'Write the first failure, when any
        If VBA.Len(mPrimaryFailure) > 0 Then
            Say "PRIMARY_FAILURE=" & mPrimaryFailure
        End If
    'Write the cleanup outcome
        Say "CLEANUP=" & IIf(CleanupPassed, "PASS", "FAIL") & "; detail=" & CleanupDetail
    'Write the recovery state
        If mRecoveryRequired Then
            Say "RECOVERY=REQUIRED; reason=" & mRecoveryReason & _
                "; action=check Excel state, then run TEST_Harness.ResetTests"
        Else
            Say "RECOVERY=NONE"
        End If
    'Build, keep and write the RESULT record
        mLastResult = "RESULT=" & Verdict & "; suite=" & mSuite & _
            "; operation=" & Operation & _
            "; completeness=" & IIf(Completed, "COMPLETE", "INCOMPLETE") & _
            "; cases=" & VBA.CStr(mCaseCount) & _
            "; assertions=" & VBA.CStr(mAssertionCount) & _
            "; failures=" & VBA.CStr(mFailureCount) & _
            "; cleanup=" & IIf(CleanupPassed, "PASS", "FAIL") & _
            "; recovery=" & IIf(mRecoveryRequired, "REQUIRED", "NONE")
        Say mLastResult

'------------------------------------------------------------------------------
' RAISE THE OUTCOME
'------------------------------------------------------------------------------
    'Re-raise the runner error unchanged
        If RunErrorNumber <> 0 Then
            Err.Raise RunErrorNumber, RunErrorSource, RunErrorDescription
        End If
    'Raise a run that did not pass
        If Verdict <> "PASS" Then
            Err.Raise ERR_TEST_FAILED, "TEST_Harness.EndRun", _
                "Suite " & mSuite & " did not pass; see the log."
        End If

End Sub

Public Sub ResetTests()

'
'------------------------------------------------------------------------------
'                            ACKNOWLEDGE A RECOVERY
'------------------------------------------------------------------------------
' PURPOSE
'   Acknowledges an interrupted run or an unverified cleanup, so that the
'   next run may start.
'
' WHY THIS EXISTS
'   BeginRun refuses to start while a run is active or recovery is
'   required; only this deliberate reset clears that state.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the run state and both flags, starts a new log and writes the
'   RESET record.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   ResetRunState, Say
'
' NOTES
'   Run only after the interrupted run has stopped and the Excel state has
'   been checked by hand; it restores nothing itself. Typed in the
'   Immediate window.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESET THE HARNESS
'------------------------------------------------------------------------------
    'Return the run state to its idle values
        ResetRunState
    'Clear the recovery flag and its reason
        mRecoveryRequired = False
        mRecoveryReason = vbNullString
    'Start a new log and write the RESET record
        Set mLog = New Collection
        Say "RESET=DONE; run_active=no; recovery=NONE"

End Sub

'
'------------------------------------------------------------------------------
'
'                                  ASSERTIONS
'
'------------------------------------------------------------------------------
'

Public Sub AssertEqualDouble( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double)
'
'------------------------------------------------------------------------------
'                            ASSERT EXACT EQUALITY
'------------------------------------------------------------------------------
' PURPOSE
'   Requires exact Double equality, for results that are exact in binary.
'
' WHY THIS EXISTS
'   Results that are exact in binary are compared without a tolerance;
'   AssertNear covers the others.
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case
'
'   Expected
'     Expected value
'
'   Actual
'     Value to compare with Expected
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts and logs the assertion; a failure names both values.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A failed comparison is recorded,
'   not raised.
'
' DEPENDENCIES
'   LogAssertion, NumberText
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Log the exact comparison with both values
        LogAssertion AssertionName, (Actual = Expected), _
            "expected=" & NumberText(Expected) & "; actual=" & NumberText(Actual)

End Sub

Public Sub AssertNear( _
    ByVal AssertionName As String, _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal AbsoluteTolerance As Double, _
    ByVal RelativeTolerance As Double)
'
'------------------------------------------------------------------------------
'                      ASSERT EQUALITY WITHIN A TOLERANCE
'------------------------------------------------------------------------------
' PURPOSE
'   Requires |Actual - Expected| <= max(AbsoluteTolerance, RelativeTolerance
'   * |Expected|), the rule of docs/methodology/TEST_CASES.md (Tolerance).
'
' WHY THIS EXISTS
'   Results that are not exact in binary are compared with the documented
'   tolerance rule.
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case
'
'   Expected
'     Expected value
'
'   Actual
'     Value to compare with Expected
'
'   AbsoluteTolerance
'     Absolute bound, >= 0 and always stated. Near zero the absolute bound
'     governs
'
'   RelativeTolerance
'     Bound relative to |Expected|, >= 0 and always stated. Both 0 demands
'     exact equality
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Fails the assertion when a tolerance is negative; otherwise counts and
'   logs the result of WithinTolerance, naming both values and both
'   tolerances.
'
' ERROR POLICY
'   A negative tolerance fails the assertion; nothing is raised. Raises
'   ERR_TEST_USAGE outside a run.
'
' DEPENDENCIES
'   LogAssertion, WithinTolerance, NumberText
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Fail a negative tolerance; otherwise apply the tolerance rule
        If AbsoluteTolerance < 0# Or RelativeTolerance < 0# Then
            LogAssertion AssertionName, False, "tolerances must not be negative"
        Else
            LogAssertion AssertionName, _
                WithinTolerance(Expected, Actual, AbsoluteTolerance, RelativeTolerance), _
                "expected=" & NumberText(Expected) & "; actual=" & NumberText(Actual) & _
                "; absolute=" & NumberText(AbsoluteTolerance) & _
                "; relative=" & NumberText(RelativeTolerance)
        End If

End Sub

Public Sub AssertEqualText( _
    ByVal AssertionName As String, _
    ByVal Expected As String, _
    ByVal Actual As String)
'
'------------------------------------------------------------------------------
'                              ASSERT EQUAL TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Requires the same text, compared binary: case and spaces count.
'
' WHY THIS EXISTS
'   A text result must match exactly, so that a change of case or spacing
'   is reported.
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case
'
'   Expected
'     Expected text
'
'   Actual
'     Text to compare with Expected
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts and logs the binary comparison; a failure quotes both texts.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A failed comparison is recorded,
'   not raised.
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Log the binary comparison with both texts quoted
        LogAssertion AssertionName, (VBA.StrComp(Actual, Expected, vbBinaryCompare) = 0), _
            "expected='" & Expected & "'; actual='" & Actual & "'"

End Sub

Public Sub AssertTrue( _
    ByVal AssertionName As String, _
    ByVal Condition As Boolean, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                              ASSERT A CONDITION
'------------------------------------------------------------------------------
' PURPOSE
'   Requires a condition the suite evaluated; Detail is logged on failure.
'
' WHY THIS EXISTS
'   A check that the other assertions do not cover is evaluated by the
'   suite and still counted and logged as an assertion.
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case
'
'   Condition
'     Result of the check, True when it passed
'
'   Detail
'     Text logged when Condition is False
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts and logs the assertion.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A false condition is recorded, not
'   raised.
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Log the condition with its detail
        LogAssertion AssertionName, Condition, Detail

End Sub

Public Sub AssertErrorValue( _
    ByVal AssertionName As String, _
    ByVal ExpectedCode As Long, _
    ByVal Actual As Variant)
'
'------------------------------------------------------------------------------
'                        ASSERT A WORKSHEET ERROR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Requires a specific worksheet error value, such as CVErr(xlErrValue).
'
' WHY THIS EXISTS
'   Worksheet functions report invalid inputs as error values, which cannot
'   be compared with "=" like other values.
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case
'
'   ExpectedCode
'     An xlErr constant
'
'   Actual
'     The value returned
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Fails when Actual is not an error value, naming its type; otherwise
'   compares the two error values as text ("Error 2007").
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A wrong type or value is recorded,
'   not raised.
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   The error value is compared through its text, VBA.CStr(Actual) =
'   VBA.CStr(CVErr(ExpectedCode)); tools/test_vba_harness.py checks that form.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Check the type first, then compare the error values as text
    '("Error 2007"): comparing a value with an error value by "=" raises
    'error 13, Type mismatch
        If Not IsError(Actual) Then
            LogAssertion AssertionName, False, _
                "expected error " & VBA.CStr(ExpectedCode) & "; actual type=" & TypeName(Actual)
        Else
            LogAssertion AssertionName, (VBA.CStr(Actual) = VBA.CStr(CVErr(ExpectedCode))), _
                "expected error " & VBA.CStr(ExpectedCode) & "; actual=" & VBA.CStr(Actual)
        End If

End Sub

Public Sub AssertRaised( _
    ByVal AssertionName As String, _
    ByVal ExpectedNumber As Long, _
    ByVal ActualNumber As Long)
'
'------------------------------------------------------------------------------
'                            ASSERT A RAISED ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Requires that a call raised a specific error number.
'
' WHY THIS EXISTS
'   Error paths are part of the tested behavior; the suite captures the
'   number and this assertion counts and logs it like any other check.
'
' INPUTS
'   AssertionName
'     Stable name of the assertion within the case
'
'   ExpectedNumber
'     Error number the call must raise
'
'   ActualNumber
'     Err.Number, captured by the suite on the line after the call; 0
'     means nothing was raised
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts and logs the comparison of the two numbers; a failure names
'   both.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. A different number is recorded,
'   not raised.
'
' DEPENDENCIES
'   LogAssertion
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Log the comparison of the error numbers
        LogAssertion AssertionName, (ActualNumber = ExpectedNumber), _
            "expected error=" & VBA.CStr(ExpectedNumber) & _
            "; actual error=" & VBA.CStr(ActualNumber)

End Sub

Public Function WithinTolerance( _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal AbsoluteTolerance As Double, _
    ByVal RelativeTolerance As Double) As Boolean
'
'------------------------------------------------------------------------------
'                           APPLY THE TOLERANCE RULE
'------------------------------------------------------------------------------
' PURPOSE
'   The comparison rule of AssertNear.
'
' WHY THIS EXISTS
'   Exposed so that suites can test the rule itself.
'
' INPUTS
'   Expected
'     Expected value
'
'   Actual
'     Value to compare with Expected
'
'   AbsoluteTolerance
'     Absolute bound
'
'   RelativeTolerance
'     Bound relative to |Expected|
'
' RETURNS
'   True when |Actual - Expected| <= max(AbsoluteTolerance,
'   RelativeTolerance * |Expected|); the bound is inclusive.
'
' BEHAVIOR
'   Takes the larger of the two bounds and compares the absolute
'   difference with it.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   Negative tolerances are not checked here; AssertNear rejects them.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "TEST_Harness.WithinTolerance"

    Dim Bound       As Double   'Larger of the two bounds

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Start from the relative bound
        Bound = RelativeTolerance * Abs(Expected)
    'Use the absolute bound when it is larger
        If AbsoluteTolerance > Bound Then
            Bound = AbsoluteTolerance
        End If
    'Apply the inclusive bound
        WithinTolerance = (Abs(Actual - Expected) <= Bound)

End Function

'
'------------------------------------------------------------------------------
'
'                                     LOG
'
'------------------------------------------------------------------------------
'

Public Sub Say(ByVal Text As String)

'
'------------------------------------------------------------------------------
'                              WRITE A LOG RECORD
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one log record to the Immediate window and keeps it for SaveLog.
'
' WHY THIS EXISTS
'   The Immediate window keeps only its last 200 lines; the kept records
'   are the complete log.
'
' INPUTS
'   Text
'     One record, KEY=value, fields separated by "; "
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Starts a log when none is kept, appends the record and prints it.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA Debug object
'
' NOTES
'   Works outside a run as well; BeginRun uses it to log a refused start.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Start a log when none is kept
        If mLog Is Nothing Then
            Set mLog = New Collection
        End If
    'Keep the record for SaveLog
        mLog.Add Text
    'Print the record to the Immediate window
        Debug.Print Text

End Sub

Public Sub SaveLog(ByVal Path As String)

'
'------------------------------------------------------------------------------
'                            SAVE THE LOG TO A FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes every record kept since the last BeginRun or ResetTests to a text
'   file.
'
' WHY THIS EXISTS
'   The Immediate window keeps only its last 200 lines; the file is the
'   complete log.
'
' INPUTS
'   Path
'     Full path of the file, created or overwritten; outside the checkout
'     or in the ignored test-results/ directory
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens the file for output, writes one record per line and closes it.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE when no record is kept; a file error is re-raised
'   after the file is closed.
'
' DEPENDENCIES
'   VBA file statements (FreeFile, Open, Print #, Close)
'
' NOTES
'   Also typed in the Immediate window.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "TEST_Harness.SaveLog"

    Dim FileNumber          As Integer  'Handle of the open file
    Dim LogEntry            As Variant  'One kept record
    Dim ErrorNumber         As Long     'File error number
    Dim ErrorSource         As String   'File error source
    Dim ErrorDescription    As String   'File error description

'------------------------------------------------------------------------------
' VALIDATE THE LOG
'------------------------------------------------------------------------------
    'Refuse to save when no record is kept
        If mLog Is Nothing Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness.SaveLog", "No log record is kept."
        End If

'------------------------------------------------------------------------------
' WRITE THE FILE
'------------------------------------------------------------------------------
    'Create or overwrite the file
        FileNumber = FreeFile
        Open Path For Output As #FileNumber
    'Close the file before a write error is re-raised
        On Error GoTo WriteFailed
    'Write one record per line
        For Each LogEntry In mLog
            Print #FileNumber, VBA.CStr(LogEntry)
        Next LogEntry
    'Close the file
        Close #FileNumber

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the write error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
WriteFailed:
    'Capture the write error
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Restore normal error handling
        On Error GoTo 0
    'Close the file
        Close #FileNumber
    'Re-raise the write error unchanged
        Err.Raise ErrorNumber, ErrorSource, ErrorDescription

End Sub

Public Function LastResult() As String

'
'------------------------------------------------------------------------------
'                            RETURN THE LAST RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the RESULT record of the last finished run.
'
' WHY THIS EXISTS
'   The RESULT record stays readable after EndRun, until the next run
'   replaces it.
'
' INPUTS
'   None
'
' RETURNS
'   The RESULT record of the last finished run; empty before the first.
'
' BEHAVIOR
'   Returns the kept record.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   None
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the kept RESULT record
        LastResult = mLastResult

End Function

'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub LogAssertion( _
    ByVal AssertionName As String, _
    ByVal Passed As Boolean, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                              LOG ONE ASSERTION
'------------------------------------------------------------------------------
' PURPOSE
'   Counts one assertion, logs its outcome and records it when it failed.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE outside a run. An assertion outside a case is
'   counted and recorded as a failure.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD THE ASSERTION
'------------------------------------------------------------------------------
    'Refuse an assertion outside a run
        RequireActiveRun "an assertion"
    'Count the assertion
        mAssertionCount = mAssertionCount + 1
    'Record an assertion made outside a case
        If VBA.Len(mCurrentCase) = 0 Then
            RecordFailure "harness.no-case", "assertion '" & AssertionName & "' outside a case"
        End If
    'Write the ASSERT record; record a failed assertion
        If Passed Then
            Say "ASSERT=" & mCurrentCase & "/" & AssertionName & "; status=PASS"
        Else
            Say "ASSERT=" & mCurrentCase & "/" & AssertionName & "; status=FAIL; " & Detail
            RecordFailure mCurrentCase & "/" & AssertionName, Detail
        End If

End Sub

Private Sub RecordFailure( _
    ByVal FailureId As String, _
    ByVal Detail As String)
'
'------------------------------------------------------------------------------
'                               RECORD A FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Counts and logs one failure; the first one of a run is its primary
'   failure.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
    'Count the failure
        mFailureCount = mFailureCount + 1
    'Keep the first failure of the run as its primary failure
        If VBA.Len(mPrimaryFailure) = 0 Then
            mPrimaryFailure = FailureId & "; " & Detail
        End If
    'Write the FAILURE record
        Say "FAILURE=" & FailureId & "; " & Detail

End Sub

Private Sub RequireActiveRun(ByVal Caller As String)

'
'------------------------------------------------------------------------------
'                            REQUIRE AN ACTIVE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Refuses a lifecycle call or an assertion made outside BeginRun..EndRun.
'
' ERROR POLICY
'   Raises ERR_TEST_USAGE, naming Caller, when no run is active.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Refuse a call made outside a run
        If Not mRunActive Then
            Err.Raise ERR_TEST_USAGE, "TEST_Harness", _
                "No test run is active for " & Caller & "; call BeginRun first."
        End If

End Sub

Private Function VerifyHostState(ByRef Detail As String) As Boolean

'
'------------------------------------------------------------------------------
'                            VERIFY THE HOST STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Compares the four Application properties with their BeginRun values.
'
' WHY THIS EXISTS
'   EndRun reports cleanup PASS only when the run left the host state
'   unchanged.
'
' INPUTS
'   Detail
'     Receives the names of the changed properties, or why they could not
'     be compared
'
' RETURNS
'   True when all four are unchanged; Detail names the changed ones, or why
'   they could not be compared.
'
' BEHAVIOR
'   Fails without reading when the BeginRun snapshot is invalid; otherwise
'   reads Calculation, DisplayAlerts, EnableEvents and ScreenUpdating and
'   names each one that changed.
'
' ERROR POLICY
'   A failed read is not raised: the result is False and Detail names the
'   error.
'
' DEPENDENCIES
'   Excel Application object
'
' NOTES
'   The harness never restores the properties: it never changed them, and
'   the procedure that did owns their restoration.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "TEST_Harness.VerifyHostState"

    Dim Changed     As String   'Names of the changed properties

'------------------------------------------------------------------------------
' CHECK THE SNAPSHOT
'------------------------------------------------------------------------------
    'Fail without comparing when the snapshot could not be read
        If Not mSnapshotValid Then
            Detail = "Excel state not compared: it could not be read at BeginRun"
            VerifyHostState = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Send a failed read to the read handler
        On Error GoTo ReadFailed
    'Name each property that changed
        If Application.Calculation <> mInitialCalculation Then Changed = Changed & " Calculation"
        If Application.DisplayAlerts <> mInitialAlerts Then Changed = Changed & " DisplayAlerts"
        If Application.EnableEvents <> mInitialEvents Then Changed = Changed & " EnableEvents"
        If Application.ScreenUpdating <> mInitialScreen Then Changed = Changed & " ScreenUpdating"
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' DESCRIBE THE OUTCOME
'------------------------------------------------------------------------------
    'Pass only when no property changed
        If VBA.Len(Changed) = 0 Then
            Detail = "Excel state unchanged (Calculation, DisplayAlerts, EnableEvents, ScreenUpdating)"
            VerifyHostState = True
        Else
            Detail = "Excel state changed during the run:" & Changed
            VerifyHostState = False
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the read handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ReadFailed:
    'Fail and name the read error
        Detail = "Excel state not compared: read failed with error " & VBA.CStr(Err.Number) & _
            " (" & Err.Description & ")"
        VerifyHostState = False

End Function

Private Function IsCaseId(ByVal CaseId As String) As Boolean

'
'------------------------------------------------------------------------------
'                           CHECK A CASE IDENTIFIER
'------------------------------------------------------------------------------
' PURPOSE
'   Whether CaseId is lower-case kebab-case: a-z and 0-9 in groups joined by
'   single hyphens, as the test-case files use (TEST_CASES.md, Files).
'
' WHY THIS EXISTS
'   Logs are compared case by case, so identifiers follow one form.
'
' INPUTS
'   CaseId
'     Identifier to check
'
' RETURNS
'   True for lower-case kebab-case; False for an empty identifier, a
'   leading, trailing or double hyphen, or any other character.
'
' BEHAVIOR
'   Checks the hyphens first, then every character.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   The module has no Option Compare statement, so Like compares binary and
'   upper-case letters are rejected.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TEST_Harness.IsCaseId"

    Dim Character       As String   'Character under test
    Dim CharacterIndex  As Long     'Position in CaseId

'------------------------------------------------------------------------------
' CHECK THE HYPHENS
'------------------------------------------------------------------------------
    'Reject an empty identifier, or one that starts or ends with a hyphen
        If VBA.Len(CaseId) = 0 Or VBA.Left$(CaseId, 1) = "-" Or VBA.Right$(CaseId, 1) = "-" Then
            IsCaseId = False
            Exit Function
        End If
    'Reject a double hyphen
        If VBA.InStr(1, CaseId, "--", vbBinaryCompare) > 0 Then
            IsCaseId = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' CHECK THE CHARACTERS
'------------------------------------------------------------------------------
    'Reject any character other than a-z, 0-9 and the hyphen
        For CharacterIndex = 1 To VBA.Len(CaseId)
            Character = VBA.Mid$(CaseId, CharacterIndex, 1)
            If Not (Character Like "[a-z0-9-]") Then
                IsCaseId = False
                Exit Function
            End If
        Next CharacterIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Accept the identifier
        IsCaseId = True

End Function

Private Function IsKnownCase(ByVal CaseId As String) As Boolean

'
'------------------------------------------------------------------------------
'                          CHECK FOR A REPEATED CASE
'------------------------------------------------------------------------------
' PURPOSE
'   Whether CaseId already started in this run.
'
' ERROR POLICY
'   Suppresses the error that an unknown key raises; it is the answer.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "TEST_Harness.IsKnownCase"

    Dim Found       As Variant  'Item stored under CaseId, when present

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'A Collection raises error 5 for an unknown key; that error is the answer
        On Error Resume Next
        Found = mCaseIds.Item(CaseId)
        IsKnownCase = (Err.Number = 0)
    'Restore normal error handling
        On Error GoTo 0

End Function

Private Function NumberText(ByVal Value As Double) As String

'
'------------------------------------------------------------------------------
'                           FORMAT A NUMBER AS TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   A Double as text that does not depend on the regional settings: Str$
'   always uses a period as the decimal separator. Str$ drops the zero
'   before the point (" .75"), which is put back ("0.75", "-0.75").
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "TEST_Harness.NumberText"

    Dim Text        As String   'Str$ result without its sign space

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Format with Str$ and drop its sign space
        Text = VBA.Trim$(Str$(Value))
    'Put back the zero before the point
        If VBA.Left$(Text, 1) = "." Then
            Text = "0" & Text
        ElseIf VBA.Left$(Text, 2) = "-." Then
            Text = "-0" & VBA.Mid$(Text, 2)
        End If
    'Return the text
        NumberText = Text

End Function

Private Function EnvironmentSummary() As String

'
'------------------------------------------------------------------------------
'                           DESCRIBE THE ENVIRONMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Describes the Excel host and the compiled VBA environment.
'
' WHY THIS EXISTS
'   The ENVIRONMENT record ties each log to its host; the logs of one
'   recorded Excel run must all come from one 64-bit host
'   (tools/check_excel_run.py).
'
' INPUTS
'   None
'
' RETURNS
'   host, version, build, os, office (bitness) and runtime (VBA
'   generation), as KEY=value fields separated by "; ".
'
' BEHAVIOR
'   Reads the host fields from the Application object and the office and
'   runtime fields from conditional compilation.
'
' ERROR POLICY
'   None. Errors are not handled here and reach the caller.
'
' DEPENDENCIES
'   Excel Application object; the Win64 and VBA7 compiler constants
'
' NOTES
'   office and runtime come from conditional compilation, not from Windows.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "TEST_Harness.EnvironmentSummary"

    Dim Bitness     As String   'Office bitness
    Dim Runtime     As String   'VBA generation

'------------------------------------------------------------------------------
' DESCRIBE
'------------------------------------------------------------------------------
    'Name the Office bitness
        #If Win64 Then
            Bitness = "64-bit"
        #Else
            Bitness = "32-bit"
        #End If
    'Name the VBA generation
        #If VBA7 Then
            Runtime = "VBA7"
        #Else
            Runtime = "VBA6"
        #End If
    'Join the host and environment fields
        EnvironmentSummary = "host=" & Application.Name & "; version=" & Application.Version & _
            "; build=" & VBA.CStr(Application.Build) & "; os=" & Application.OperatingSystem & _
            "; office=" & Bitness & "; runtime=" & Runtime

End Function

Private Sub ResetRunState()

'
'------------------------------------------------------------------------------
'                             RESET THE RUN STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the counters, case identifiers, failures and flags of a run to
'   their idle values; the recovery flag and the log are left to callers.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Release the run and clear its suite and counts
        mRunActive = False
        mSuite = vbNullString
        mExpectedCases = 0
        mExpectedAssertions = 0
        mCaseCount = 0
        mAssertionCount = 0
        mFailureCount = 0
        mErrorCount = 0
    'Clear the current case and start a new set of case identifiers
        mCurrentCase = vbNullString
        Set mCaseIds = New Collection
    'Clear the primary failure and invalidate the snapshot
        mPrimaryFailure = vbNullString
        mSnapshotValid = False

End Sub
