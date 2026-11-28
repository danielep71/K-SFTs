Attribute VB_Name = "M_Main"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the macros behind the buttons on the sheets, run each operation
'   with Excel prepared for speed, put Excel back exactly as it was, and
'   report the outcome to the user.
'
' WHY THIS EXISTS
'   The engine calculates and writes the outputs; the sheet buttons need
'   macros around it that prepare Excel for the run, leave Excel as they
'   found it on success and on failure, and tell the user how the operation
'   ended.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Manages:
'     - the three button macros and RunSFT_Silent for automation
'     - sheet protection for the user interface only, when the workbook
'       opens and before each operation
'     - capture, preparation and restoration of calculation mode, events
'       and screen updating around each operation
'     - the outcome and failure messages, or the result line in silent mode
'
' ERROR POLICY
'   The primary error is copied before cleanup.
'   Each setting is restored by its own helper, so one failed restoration
'   does not stop the others, and a cleanup failure is reported separately
'   from the primary error. With gSilent False the outcome, or the error,
'   is shown in a message box; with gSilent True an error is raised to the
'   caller, and the outcome is printed to the Immediate window and returned
'   as the result line. Errors in the inputs are not VBA errors: the engine
'   lists them on Checks.
'
' DEPENDENCIES
'   CORE_Engine for the calculation and its run counters; CORE_Util and
'   CORE_Config for the output sheets and the error numbers.
'
' NOTES
'   The macros must stay visible to the sheet buttons, so this module is not
'   private (no Option Private Module).
'
' PUBLIC SURFACE
'   RunSFT: validate the inputs, calculate and write every output sheet.
'   ValidateInputs: validate only; findings go to the Checks sheet.
'   ClearOutputs: empty every output sheet.
'   RunSFT_Silent: RunSFT for automation. Returns a machine-readable result
'   line and raises errors instead of showing them.
'   gSilent: Public flag for silent mode, set by RunSFT_Silent.
'   Auto_Open: run by Excel when the workbook opens; protects every sheet
'   for the user interface only and lets the row groups of the README sheet
'   expand and collapse.
'   The sheet buttons call the first three by name. These macros are not
'   listed in docs/PUBLIC_API.txt.
'
' STATE OWNERSHIP
'   Protects every sheet for the user interface only when the workbook
'   opens and before each operation. Each operation captures calculation
'   mode, events and screen updating, changes them for the run, and
'   restores the captured values on success and on failure. Owns gSilent
'   (messages to the Immediate window instead of message boxes, and errors
'   raised to the caller) and the re-entry flag mRunning.
'
' KNOWN DEVIATION
'   This module calls CORE_Engine and CORE_Util in src/core directly; the
'   repository structure routes src/workbook through the src/modules
'   facade, which has no engine entry point yet.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' AUTHOR
'   Daniele Penza
'
' UPDATED
'   2026-10-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Operations run by ExecuteOperation
    Private Const OP_RUN                As Long = 1             'Calculate and write every output sheet
    Private Const OP_VALIDATE           As Long = 2             'Validate only
    Private Const OP_CLEAR              As Long = 3             'Clear every output sheet

    'Title of every message box
    Private Const MSG_TITLE             As String = "K-SFTs"    'The project name

'------------------------------------------------------------------------------
' PUBLIC STATE
'------------------------------------------------------------------------------
    'Switches messages to the Immediate window and errors to the caller
    Public gSilent                      As Boolean              'Silent mode for automation and tests

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'Refuses a second operation while one is in progress
    Private mRunning                    As Boolean              'True while an operation is in progress


'
'------------------------------------------------------------------------------
'
'                                BUTTON MACROS
'
'------------------------------------------------------------------------------
'

Public Sub RunSFT()

'
'------------------------------------------------------------------------------
'                           RUN THE SFT CALCULATION
'------------------------------------------------------------------------------
' PURPOSE
'   Button macro: validate the inputs, calculate E*, RWA and own funds of
'   every netting set, and write Results, Positions, LegCalc and Checks.
'
' WHY THIS EXISTS
'   A sheet button calls this macro by name.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs OP_RUN through ExecuteOperation, which protects the sheets,
'   prepares Excel, calculates, restores Excel and reports the outcome.
'
' ERROR POLICY
'   Handled by ExecuteOperation: an error is shown in a message box, or
'   raised to the caller while gSilent is True.
'
' DEPENDENCIES
'   ExecuteOperation
'
' NOTES
'   The result line of ExecuteOperation is discarded here; RunSFT_Silent
'   returns it.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Calculate and write every output sheet
        ExecuteOperation OP_RUN

End Sub

Public Function RunSFT_Silent() As String

'
'------------------------------------------------------------------------------
'                       RUN THE SFT CALCULATION SILENTLY
'------------------------------------------------------------------------------
' PURPOSE
'   RunSFT for automation: no message box; the outcome is returned.
'
' WHY THIS EXISTS
'   Automation and tests need a machine-readable result line and errors
'   raised to them instead of message boxes.
'
' INPUTS
'   None
'
' RETURNS
'   A result line: "OK" or "STOPPED", then the run class, the error and
'   warning counts and the totals per regime, separated by " | ".
'
' BEHAVIOR
'   Saves gSilent and sets it to True, runs OP_RUN through
'   ExecuteOperation, restores gSilent and returns the result line. On an
'   error it restores gSilent and raises the error again.
'
' ERROR POLICY
'   Raises the primary error, or ERR_CLEANUP_FAILED, to the caller.
'
' DEPENDENCIES
'   ExecuteOperation
'
' NOTES
'   The error is raised again with its own number, source and description.
'
' STATE OWNERSHIP
'   Sets gSilent for the operation and restores its previous value.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Main.RunSFT_Silent"

    Dim SavedSilent             As Boolean      'gSilent before the call
    Dim ErrNumber               As Long         'Primary error number
    Dim ErrSource               As String       'Primary error source
    Dim ErrText                 As String       'Primary error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Save the caller's silent flag and switch to silent mode
        SavedSilent = gSilent
        gSilent = True
    'Enable controlled error handling
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
    'Run the calculation and return its result line
        RunSFT_Silent = ExecuteOperation(OP_RUN)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Restore the caller's silent flag
        gSilent = SavedSilent
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Capture the primary error
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrText = Err.Description
    'Restore the caller's silent flag
        gSilent = SavedSilent
    'Raise the primary error again, unchanged
        Err.Raise ErrNumber, ErrSource, ErrText

End Function

Public Sub ValidateInputs()

'
'------------------------------------------------------------------------------
'                             VALIDATE THE INPUTS
'------------------------------------------------------------------------------
' PURPOSE
'   Button macro: check the inputs and list the findings on Checks, without
'   calculating; the other output sheets are cleared.
'
' WHY THIS EXISTS
'   A sheet button calls this macro by name.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs OP_VALIDATE through ExecuteOperation, which protects the sheets,
'   prepares Excel, validates, restores Excel and reports the outcome.
'
' ERROR POLICY
'   Handled by ExecuteOperation: an error is shown in a message box, or
'   raised to the caller while gSilent is True.
'
' DEPENDENCIES
'   ExecuteOperation
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
    'Validate the inputs only
        ExecuteOperation OP_VALIDATE

End Sub

Public Sub ClearOutputs()

'
'------------------------------------------------------------------------------
'                           CLEAR THE OUTPUT SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Button macro: empty every output sheet.
'
' WHY THIS EXISTS
'   A sheet button calls this macro by name.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs OP_CLEAR through ExecuteOperation, which protects the sheets,
'   prepares Excel, clears the outputs, restores Excel and reports the
'   outcome.
'
' ERROR POLICY
'   Handled by ExecuteOperation: an error is shown in a message box, or
'   raised to the caller while gSilent is True.
'
' DEPENDENCIES
'   ExecuteOperation
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
    'Empty every output sheet
        ExecuteOperation OP_CLEAR

End Sub

'
'------------------------------------------------------------------------------
'
'                                  OPEN MACRO
'
'------------------------------------------------------------------------------
'

Public Sub Auto_Open()

'
'------------------------------------------------------------------------------
'                        PROTECT THE SHEETS ON OPENING
'------------------------------------------------------------------------------
' PURPOSE
'   Run by Excel when the user opens the workbook: protects every sheet
'   again for the user interface only (ProtectSheets).
'
' WHY THIS EXISTS
'   Excel saves neither UserInterfaceOnly nor EnableOutlining with the
'   file, so the protection has to be applied again on opening.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls ProtectSheets.
'
' ERROR POLICY
'   Never raises.
'
' DEPENDENCIES
'   ProtectSheets
'
' NOTES
'   Excel runs the macro by this name; it must not be renamed.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PROTECT
'------------------------------------------------------------------------------
    'Protect every sheet for the user interface only
        ProtectSheets

End Sub

Private Sub ProtectSheets()

'
'------------------------------------------------------------------------------
'                             PROTECT EVERY SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Protect every sheet again, without a password and for the user
'   interface only, so that the macros can write the output sheets while
'   the user can change only the unlocked input cells. The input sheets
'   keep filtering, sorting and inserting rows, as in the template; the
'   README sheet keeps its row groups. Excel saves neither
'   UserInterfaceOnly nor EnableOutlining with the file, so this runs when
'   the workbook opens and before each operation.
'
' ERROR POLICY
'   Never raises: a sheet protected with a password stays as it is.
'
' DEPENDENCIES
'   ProtectSheet
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet                   As Worksheet    'Each worksheet

'------------------------------------------------------------------------------
' PROTECT
'------------------------------------------------------------------------------
    'Protect each worksheet of this workbook in turn
        For Each Sheet In ThisWorkbook.Worksheets
            ProtectSheet Sheet
        Next Sheet

End Sub

Private Sub ProtectSheet(ByVal Sheet As Worksheet)

'
'------------------------------------------------------------------------------
'                              PROTECT ONE SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Protect one sheet for the user interface only, with the options the
'   template gives it (tools/build_template.py).
'
' ERROR POLICY
'   Never raises: a sheet protected with a password stays as it is.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Leave the sheet as it is when it cannot be protected
        On Error GoTo Failed

'------------------------------------------------------------------------------
' PROTECT
'------------------------------------------------------------------------------
    'Keep sorting, filtering and inserting rows on the input sheets
        Select Case Sheet.Name
            Case SH_PARAMS, SH_NS, SH_TX
                Sheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                    UserInterfaceOnly:=True, AllowSorting:=True, AllowFiltering:=True, _
                    AllowInsertingRows:=True
            Case Else
                Sheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                    UserInterfaceOnly:=True
        End Select
    'Let the row groups of the README sheet expand and collapse
        If Sheet.Name = SH_README Then
            Sheet.EnableOutlining = True
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
    'Ignore the failure: a sheet protected with a password stays as it is

End Sub

'
'------------------------------------------------------------------------------
'
'                               OPERATION RUNNER
'
'------------------------------------------------------------------------------
'

Private Function ExecuteOperation(ByVal Operation As Long) As String

'
'------------------------------------------------------------------------------
'                              RUN ONE OPERATION
'------------------------------------------------------------------------------
' PURPOSE
'   Protect every sheet for the user interface only, run one operation
'   with Excel prepared for speed, restore Excel, then report the outcome.
'
' WHY THIS EXISTS
'   The three button macros and RunSFT_Silent share one path that prepares
'   and restores Excel and reports the outcome.
'
' INPUTS
'   Operation
'     OP_RUN, OP_VALIDATE or OP_CLEAR.
'
' RETURNS
'   The result line built by ReportOutcome; empty when the operation raised
'   an error or was refused because another was running (in silent mode
'   these are raised instead).
'
' BEHAVIOR
'   In order:
'     - refuses to start while another operation is running
'     - protects every sheet for the user interface only
'     - captures calculation mode, events and screen updating, then turns
'       off screen updating and events and sets manual calculation
'     - unless clearing, recalculates and stops with ERR_RECALC_PENDING
'       when Excel has not finished
'     - runs the operation
'     - restores the captured settings and reports the outcome or the
'       error
'
' ERROR POLICY
'   The primary error is copied into locals before cleanup and reported, or
'   raised in silent mode, after cleanup. A cleanup failure is reported
'   with it, or on its own when the operation succeeded.
'
' DEPENDENCIES
'   ProtectSheets, CORE_Engine.Calculate, ClearAllOutputs, RestoreSettings,
'   ReportOutcome, ReportFailure
'
' NOTES
'   The error handler resumes at CleanUp, so the restoration and the report
'   run on every path.
'
' STATE OWNERSHIP
'   Captures calculation mode, events and screen updating before changing
'   them, and restores the captured values on every path.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Main.ExecuteOperation"

    Dim SavedCalculation        As XlCalculation    'Calculation mode before the operation
    Dim SavedEvents             As Boolean          'EnableEvents before the operation
    Dim SavedScreen             As Boolean          'ScreenUpdating before the operation
    Dim StateCaptured           As Boolean          'True once all three were read
    Dim Completed               As Boolean          'True when the operation completed
    Dim ErrorNumber             As Long             'Primary error number; 0 for none
    Dim ErrorSource             As String           'Primary error source
    Dim ErrorDescription        As String           'Primary error description
    Dim CleanupDetails          As String           'Restoration failures; empty when all restored

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Refuse a second operation while one is in progress
        If mRunning Then
            ReportFailure ERR_RUN_ACTIVE, "M_Main.ExecuteOperation", _
                "An SFT operation is already running.", vbNullString
            Exit Function
        End If
    'Mark the operation as running
        mRunning = True
    'Enable controlled error handling
        On Error GoTo Failed

'------------------------------------------------------------------------------
' PROTECT FOR THE USER INTERFACE ONLY
'------------------------------------------------------------------------------
    'The protection Excel restores on opening blocks the macros as well;
    'protected again for the user interface only, the sheets can be written
        ProtectSheets

'------------------------------------------------------------------------------
' CAPTURE AND PREPARE EXCEL
'------------------------------------------------------------------------------
    'Capture the settings the operation changes
        SavedCalculation = Application.Calculation
        SavedEvents = Application.EnableEvents
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
    'Prepare Excel for speed
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

'------------------------------------------------------------------------------
' BRING FORMULA INPUTS UP TO DATE
'------------------------------------------------------------------------------
    'In manual calculation a formula input could still hold an old result;
    'Calculate brings every dirty cell up to date before the engine reads
        If Operation <> OP_CLEAR Then
            Application.Calculate
            If Application.CalculationState <> xlDone Then
                Err.Raise ERR_RECALC_PENDING, "M_Main.ExecuteOperation", _
                    "Excel did not finish recalculating the inputs, so the operation was not started."
            End If
        End If

'------------------------------------------------------------------------------
' RUN THE OPERATION
'------------------------------------------------------------------------------
    'Calculate, validate or clear
        Select Case Operation
            Case OP_RUN
                Completed = CORE_Engine.Calculate(True)
            Case OP_VALIDATE
                Completed = CORE_Engine.Calculate(False)
            Case OP_CLEAR
                ClearAllOutputs
                Completed = True
            Case Else
                'Any other value runs nothing and leaves Completed False
        End Select

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore normal error handling
        On Error GoTo 0
    'Restore Excel when its settings were captured
        If StateCaptured Then
            CleanupDetails = RestoreSettings(SavedCalculation, SavedEvents, SavedScreen)
        End If
    'Allow the next operation
        mRunning = False
    'Report the error, or the outcome and its result line
        If ErrorNumber <> 0 Then
            ReportFailure ErrorNumber, ErrorSource, ErrorDescription, CleanupDetails
        Else
            ExecuteOperation = ReportOutcome(Operation, Completed, CleanupDetails)
        End If
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Capture the primary error
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Restore Excel and report through the clean exit
        Resume CleanUp

End Function

Private Sub ClearAllOutputs()

'
'------------------------------------------------------------------------------
'                           CLEAR ALL OUTPUT TABLES
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the Results, Positions, LegCalc and Checks tables and write the
'   time of clearing as the run summary.
'
' ERROR POLICY
'   Raises ERR_OUTPUT_LAYOUT through OutputFirstRow when a table cannot be
'   located. Without the run-summary name the summary is not written.
'
' DEPENDENCIES
'   CORE_Util: ClearOutputBlock, GetSheet, OutputFirstRow, NamedCell
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Info                    As Range        'Run summary cell, named NM_RUNINFO

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Empty the four output tables
        ClearOutputBlock GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_NCOLS
        ClearOutputBlock GetSheet(SH_POSITIONS), OutputFirstRow(SH_POSITIONS), PS_NCOLS
        ClearOutputBlock GetSheet(SH_LEGCALC), OutputFirstRow(SH_LEGCALC), LC_NCOLS
        ClearOutputBlock GetSheet(SH_CHECKS), OutputFirstRow(SH_CHECKS), CK_NCOLS

'------------------------------------------------------------------------------
' WRITE RUN SUMMARY
'------------------------------------------------------------------------------
    'Record the time of clearing when the run-summary cell exists
        Set Info = NamedCell(NM_RUNINFO)
        If Not Info Is Nothing Then
            Info.Value = "Outputs cleared " & Format$(Now, "yyyy-mm-dd hh:mm:ss")
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                               RESTORE SETTINGS
'
'------------------------------------------------------------------------------
'

Private Function RestoreSettings( _
    ByVal SavedCalculation As XlCalculation, _
    ByVal SavedEvents As Boolean, _
    ByVal SavedScreen As Boolean) As String
'
'------------------------------------------------------------------------------
'                        RESTORE THE CAPTURED SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the three captured settings, each attempted on its own.
'
' RETURNS
'   The failures as text; empty when all three were restored.
'
' DEPENDENCIES
'   RestoreCalculation, RestoreEvents, RestoreScreenUpdating, AppendDetail
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Restore each setting and join the failures
        RestoreSettings = AppendDetail(AppendDetail(RestoreCalculation(SavedCalculation), _
            RestoreEvents(SavedEvents)), RestoreScreenUpdating(SavedScreen))

End Function

Private Function RestoreCalculation(ByVal SavedCalculation As XlCalculation) As String

'
'------------------------------------------------------------------------------
'                         RESTORE THE CALCULATION MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Put the calculation mode back.
'
' RETURNS
'   "" on success; the failure as text otherwise.
'
' ERROR POLICY
'   Contains its own error.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn a failure into the returned text
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Put the captured calculation mode back
        Application.Calculation = SavedCalculation

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return the failure as text
        RestoreCalculation = "calculation mode not restored (error " & Err.Number & ": " & _
            Err.Description & ")"

End Function

Private Function RestoreEvents(ByVal SavedEvents As Boolean) As String

'
'------------------------------------------------------------------------------
'                                RESTORE EVENTS
'------------------------------------------------------------------------------
' PURPOSE
'   Put EnableEvents back.
'
' RETURNS
'   "" on success; the failure as text otherwise.
'
' ERROR POLICY
'   Contains its own error.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn a failure into the returned text
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Put the captured EnableEvents value back
        Application.EnableEvents = SavedEvents

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return the failure as text
        RestoreEvents = "events not restored (error " & Err.Number & ": " & Err.Description & ")"

End Function

Private Function RestoreScreenUpdating(ByVal SavedScreen As Boolean) As String

'
'------------------------------------------------------------------------------
'                           RESTORE SCREEN UPDATING
'------------------------------------------------------------------------------
' PURPOSE
'   Put ScreenUpdating back.
'
' RETURNS
'   "" on success; the failure as text otherwise.
'
' ERROR POLICY
'   Contains its own error.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn a failure into the returned text
        On Error GoTo Failed

'------------------------------------------------------------------------------
' RESTORE
'------------------------------------------------------------------------------
    'Put the captured ScreenUpdating value back
        Application.ScreenUpdating = SavedScreen

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Return the failure as text
        RestoreScreenUpdating = "screen updating not restored (error " & Err.Number & ": " & _
            Err.Description & ")"

End Function

Private Function AppendDetail(ByVal Existing As String, ByVal Addition As String) As String

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
' JOIN
'------------------------------------------------------------------------------
    'Return the other text when one is empty, otherwise both joined
        If VBA.Len(Existing) = 0 Then
            AppendDetail = Addition
        ElseIf VBA.Len(Addition) = 0 Then
            AppendDetail = Existing
        Else
            AppendDetail = Existing & "; " & Addition
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                                  REPORTING
'
'------------------------------------------------------------------------------
'

Private Function ReportOutcome( _
    ByVal Operation As Long, _
    ByVal Completed As Boolean, _
    ByVal CleanupDetails As String) As String
'
'------------------------------------------------------------------------------
'                              REPORT THE OUTCOME
'------------------------------------------------------------------------------
' PURPOSE
'   Tell the user how the operation ended, and build the result line of
'   RunSFT_Silent. A completed run's message says that the risk weight,
'   RWA and own funds are illustrative (D7; the support matrix, #248).
'   A run stopped before its parameters were checked has no run class,
'   and the message leaves it out (#220).
'
' WHY THIS EXISTS
'   The message for the user and the result line for automation are built
'   from the same counters in one place.
'
' INPUTS
'   Operation
'     OP_RUN, OP_VALIDATE or OP_CLEAR.
'
'   Completed
'     True when the operation completed; False when the run stopped.
'
'   CleanupDetails
'     Restoration failures; empty when all settings were restored.
'
' RETURNS
'   "OK" or "STOPPED", the run class, the error and warning counts and the
'   totals per regime, separated by " | "; "OK | cleared" for OP_CLEAR.
'
' BEHAVIOR
'   Composes the message and the result line from the CORE_Engine totals,
'   run class and error and warning counts, then raises or appends a
'   cleanup failure and shows the message through Notify.
'
' ERROR POLICY
'   A cleanup failure is raised in silent mode as ERR_CLEANUP_FAILED, or
'   shown with the outcome otherwise.
'
' DEPENDENCIES
'   CORE_Engine: TotalsText, RunClass, ErrorCount, WarningCount;
'   AppendDetail, Notify
'
' NOTES
'   None
'
' UPDATED
'   2026-10-31
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Main.ReportOutcome"

    Dim Message                 As String       'Text for the user
    Dim Totals                  As String       'Totals per regime

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
    'Build the message and the result line for the operation and its outcome
        Select Case Operation
            Case OP_CLEAR
                Message = "Outputs cleared."
                ReportOutcome = "OK | cleared"
            Case Else
                Totals = AppendDetail(CORE_Engine.TotalsText(RG_CRR), _
                    CORE_Engine.TotalsText(RG_BCBS))
                If Completed Then
                    ReportOutcome = "OK"
                    If Operation = OP_RUN Then
                        Message = "SFT run completed." & vbLf & vbLf & _
                            Replace(Totals, "; ", vbLf) & vbLf & _
                            "The risk weight, RWA and own funds are illustrative: only E* is supported."
                    Else
                        Message = "Validation completed: outputs cleared."
                    End If
                Else
                    ReportOutcome = "STOPPED"
                    Message = "The run stopped: see the Checks sheet."
                End If
                ReportOutcome = ReportOutcome & " | " & CORE_Engine.RunClass & " | errors " & _
                    CORE_Engine.ErrorCount & " | warnings " & CORE_Engine.WarningCount
                If VBA.Len(Totals) > 0 Then
                    ReportOutcome = ReportOutcome & " | " & Totals
                End If
                Message = Message & vbLf & vbLf & "Errors " & CORE_Engine.ErrorCount & _
                    ", warnings " & CORE_Engine.WarningCount & " (see Checks)."
                If VBA.Len(CORE_Engine.RunClass) > 0 Then
                    Message = Message & " Run class: " & CORE_Engine.RunClass & "."
                End If
        End Select

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Raise a cleanup failure in silent mode, or add it to the message
        If VBA.Len(CleanupDetails) > 0 Then
            If gSilent Then
                Err.Raise ERR_CLEANUP_FAILED, "M_Main.ReportOutcome", CleanupDetails
            End If
            Message = Message & vbLf & vbLf & "Excel settings not fully restored: " & CleanupDetails
        End If
    'Show the message, or print it in silent mode
        Notify Message, IIf(Completed, vbInformation, vbExclamation)

End Function

Private Sub ReportFailure( _
    ByVal ErrorNumber As Long, _
    ByVal ErrorSource As String, _
    ByVal ErrorDescription As String, _
    ByVal CleanupDetails As String)
'
'------------------------------------------------------------------------------
'                          REPORT A FAILED OPERATION
'------------------------------------------------------------------------------
' PURPOSE
'   Show an error that stopped an operation, or raise it in silent mode,
'   with any cleanup failure appended.
'
' ERROR POLICY
'   Raises the given error number and source in silent mode.
'
' DEPENDENCIES
'   Notify
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Description             As String       'Description with the cleanup failures

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Append any cleanup failure to the description
        Description = ErrorDescription
        If VBA.Len(CleanupDetails) > 0 Then
            Description = Description & " Excel settings not fully restored: " & CleanupDetails
        End If
    'Raise the error in silent mode
        If gSilent Then
            Err.Raise ErrorNumber, ErrorSource, Description
        End If
    'Otherwise show it
        Notify "The operation failed." & vbLf & vbLf & "Error " & ErrorNumber & " in " & _
            ErrorSource & ":" & vbLf & Description, vbCritical

End Sub

Private Sub Notify(ByVal Message As String, ByVal Style As VbMsgBoxStyle)

'
'------------------------------------------------------------------------------
'                                SHOW A MESSAGE
'------------------------------------------------------------------------------
' PURPOSE
'   Show a message box, or print to the Immediate window in silent mode.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' NOTIFY
'------------------------------------------------------------------------------
    'Print on one line in silent mode, otherwise show a message box
        If gSilent Then
            Debug.Print MSG_TITLE & ": " & Replace(Message, vbLf, " ")
        Else
            MsgBox Message, Style, MSG_TITLE
        End If

End Sub
