Attribute VB_Name = "M_Main"
'==============================================================================
' MODULE: M_Main
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the macros behind the buttons on the sheets, run each operation
'   with Excel prepared for speed, put Excel back exactly as it was, and
'   report the outcome to the user.
'
' PUBLIC SURFACE
'   RunSFT: validate the inputs, calculate and write every output sheet.
'   ValidateInputs: validate only; findings go to the Checks sheet.
'   ClearOutputs: empty every output sheet.
'   RunSFT_Silent: RunSFT for automation. Returns a machine-readable result
'   line and raises errors instead of showing them.
'   The sheet buttons call the first three by name. These macros are not
'   listed in docs/PUBLIC_API.txt.
'
' DEPENDENCIES
'   CORE_Engine for the calculation and its run counters; CORE_Util and
'   CORE_Config for the output sheets and the error numbers.
'
' STATE OWNERSHIP
'   Each operation captures calculation mode, events and screen updating,
'   changes them for the run, and restores the captured values on success
'   and on failure. Owns gSilent (messages to the Immediate window instead
'   of message boxes, and errors raised to the caller) and the re-entry
'   flag mRunning.
'
' ERROR POLICY
'   As in K-SACCR's M_Main: the primary error is copied before cleanup.
'   Each setting is restored by its own helper, so one failed restoration
'   does not stop the others, and a cleanup failure is reported separately
'   from the primary error. With gSilent False the outcome is shown in a
'   message box; with gSilent True it is raised to the caller. Errors in
'   the inputs are not VBA errors: the engine lists them on Checks.
'
' KNOWN DEVIATION
'   As in K-SACCR, this standard module lives in src/workbook because the
'   sheet buttons call it.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
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
    'Require explicit declarations. The macros must stay visible to the
    'sheet buttons, so this module is not private.
    Option Explicit

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Operations run by ExecuteOperation.
        Private Const OP_RUN        As Long = 1    'Calculate and write every output sheet
        Private Const OP_VALIDATE   As Long = 2    'Validate only
        Private Const OP_CLEAR      As Long = 3    'Clear every output sheet

    'Title of every message box: the project name.
        Private Const MSG_TITLE     As String = "K-SFTs"

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'gSilent switches messages to the Immediate window and errors to the
    'caller. mRunning refuses a second operation while one is in progress.
        Public gSilent      As Boolean    'Silent mode for automation and tests
        Private mRunning    As Boolean    'True while an operation is in progress


'
'------------------------------------------------------------------------------
'
'                                BUTTON MACROS
'
'------------------------------------------------------------------------------
'

Public Sub RunSFT()
'
'==============================================================================
'                                    RunSFT
'------------------------------------------------------------------------------
' PURPOSE
'   Button macro: validate the inputs, calculate E*, RWA and own funds of
'   every netting set, and write Results, Positions, LegCalc and Checks.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        ExecuteOperation OP_RUN

End Sub


Public Function RunSFT_Silent() As String
'
'==============================================================================
'                                RunSFT_Silent
'------------------------------------------------------------------------------
' PURPOSE
'   RunSFT for automation: no message box; the outcome is returned.
'
' RETURNS
'   A result line: "OK" or "STOPPED", then the run class, the error and
'   warning counts and the totals per regime, separated by " | ".
'
' ERROR POLICY
'   Raises the primary error, or ERR_CLEANUP_FAILED, to the caller.
'
' STATE OWNERSHIP
'   Sets gSilent for the operation and restores its previous value.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedSilent   As Boolean    'gSilent before the call
    Dim ErrNumber     As Long       'Primary error number
    Dim ErrSource     As String     'Primary error source
    Dim ErrText       As String     'Primary error description

'------------------------------------------------------------------------------
' RUN
'------------------------------------------------------------------------------
        SavedSilent = gSilent
        gSilent = True
        On Error GoTo Failed
        RunSFT_Silent = ExecuteOperation(OP_RUN)
        gSilent = SavedSilent
        Exit Function

'------------------------------------------------------------------------------
' RESTORE AND RAISE
'------------------------------------------------------------------------------
Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrText = Err.Description
        gSilent = SavedSilent
        Err.Raise ErrNumber, ErrSource, ErrText

End Function


Public Sub ValidateInputs()
'
'==============================================================================
'                                ValidateInputs
'------------------------------------------------------------------------------
' PURPOSE
'   Button macro: check the inputs and list the findings on Checks, without
'   calculating; the other output sheets are cleared.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        ExecuteOperation OP_VALIDATE

End Sub


Public Sub ClearOutputs()
'
'==============================================================================
'                                 ClearOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Button macro: empty every output sheet.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        ExecuteOperation OP_CLEAR

End Sub


'
'------------------------------------------------------------------------------
'
'                              OPERATION RUNNER
'
'------------------------------------------------------------------------------
'

Private Function ExecuteOperation( _
    ByVal Operation As Long) _
    As String
'
'==============================================================================
'                               ExecuteOperation
'------------------------------------------------------------------------------
' PURPOSE
'   Run one operation with Excel prepared for speed, restore Excel, then
'   report the outcome.
'
' INPUTS
'   Operation: OP_RUN, OP_VALIDATE or OP_CLEAR.
'
' RETURNS
'   The result line built by ReportOutcome; empty when the operation raised
'   an error or was refused because another was running.
'
' STATE OWNERSHIP
'   Captures calculation mode, events and screen updating before changing
'   them, and restores the captured values on every path.
'
' ERROR POLICY
'   The primary error is copied into locals before cleanup and reported, or
'   raised in silent mode, after cleanup. A cleanup failure is reported
'   with it, or on its own when the operation succeeded.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedCalculation   As XlCalculation    'Calculation mode before the operation
    Dim SavedEvents        As Boolean          'EnableEvents before the operation
    Dim SavedScreen        As Boolean          'ScreenUpdating before the operation
    Dim StateCaptured      As Boolean          'True once all three were read
    Dim Completed          As Boolean          'True when the operation completed
    Dim ErrorNumber        As Long             'Primary error number; 0 for none
    Dim ErrorSource        As String           'Primary error source
    Dim ErrorDescription   As String           'Primary error description
    Dim CleanupDetails     As String           'Restoration failures; empty when all restored

'------------------------------------------------------------------------------
' REFUSE RE-ENTRY
'------------------------------------------------------------------------------
        If mRunning Then
            ReportFailure ERR_RUN_ACTIVE, "M_Main.ExecuteOperation", "An SFT operation is already running.", ""
            Exit Function
        End If
        mRunning = True
        On Error GoTo Failed

'------------------------------------------------------------------------------
' CAPTURE AND PREPARE EXCEL
'------------------------------------------------------------------------------
        SavedCalculation = Application.Calculation
        SavedEvents = Application.EnableEvents
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual

'------------------------------------------------------------------------------
' BRING FORMULA INPUTS UP TO DATE
'------------------------------------------------------------------------------
    'In manual calculation a formula input could still hold an old result;
    'Calculate brings every dirty cell up to date before the engine reads.
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
        Select Case Operation
            Case OP_RUN
                Completed = CORE_Engine.Calculate(True)
            Case OP_VALIDATE
                Completed = CORE_Engine.Calculate(False)
            Case OP_CLEAR
                ClearAllOutputs
                Completed = True
        End Select

'------------------------------------------------------------------------------
' RESTORE EXCEL AND REPORT
'------------------------------------------------------------------------------
CleanUp:
        On Error GoTo 0
        If StateCaptured Then
            CleanupDetails = RestoreSettings(SavedCalculation, SavedEvents, SavedScreen)
        End If
        mRunning = False
        If ErrorNumber <> 0 Then
            ReportFailure ErrorNumber, ErrorSource, ErrorDescription, CleanupDetails
        Else
            ExecuteOperation = ReportOutcome(Operation, Completed, CleanupDetails)
        End If
        Exit Function

'------------------------------------------------------------------------------
' HANDLE ERROR
'------------------------------------------------------------------------------
Failed:
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
        Resume CleanUp

End Function


Private Sub ClearAllOutputs()
'
'==============================================================================
'                               ClearAllOutputs
'------------------------------------------------------------------------------
' PURPOSE
'   Empty the Results, Positions, LegCalc and Checks tables and write the
'   time of clearing as the run summary.
'
' ERROR POLICY
'   Raises ERR_OUTPUT_LAYOUT through OutputFirstRow when a table cannot be
'   located.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        ClearOutputBlock GetSheet(SH_RESULTS), OutputFirstRow(SH_RESULTS), RS_NCOLS
        ClearOutputBlock GetSheet(SH_POSITIONS), OutputFirstRow(SH_POSITIONS), PS_NCOLS
        ClearOutputBlock GetSheet(SH_LEGCALC), OutputFirstRow(SH_LEGCALC), LC_NCOLS
        ClearOutputBlock GetSheet(SH_CHECKS), OutputFirstRow(SH_CHECKS), CK_NCOLS
        GetSheet(SH_RESULTS).Range(RUNINFO_CELL).Value = "Outputs cleared " & Format$(Now, "yyyy-mm-dd hh:mm:ss")

End Sub


'
'------------------------------------------------------------------------------
'
'                              RESTORE SETTINGS
'
'------------------------------------------------------------------------------
'

Private Function RestoreSettings( _
    ByVal SavedCalculation As XlCalculation, _
    ByVal SavedEvents As Boolean, _
    ByVal SavedScreen As Boolean) _
    As String
'
'==============================================================================
'                               RestoreSettings
'------------------------------------------------------------------------------
' PURPOSE
'   Restore the three captured settings, each attempted on its own.
'
' RETURNS
'   The failures as text; empty when all three were restored.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        RestoreSettings = AppendDetail(AppendDetail(RestoreCalculation(SavedCalculation), _
                                                    RestoreEvents(SavedEvents)), _
                                       RestoreScreenUpdating(SavedScreen))

End Function


Private Function RestoreCalculation( _
    ByVal SavedCalculation As XlCalculation) _
    As String
'
'==============================================================================
'                              RestoreCalculation
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
'==============================================================================
'
        On Error GoTo Failed
        Application.Calculation = SavedCalculation
        Exit Function
Failed:
        RestoreCalculation = "calculation mode not restored (error " & Err.Number & ": " & Err.Description & ")"

End Function


Private Function RestoreEvents( _
    ByVal SavedEvents As Boolean) _
    As String
'
'==============================================================================
'                                RestoreEvents
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
'==============================================================================
'
        On Error GoTo Failed
        Application.EnableEvents = SavedEvents
        Exit Function
Failed:
        RestoreEvents = "events not restored (error " & Err.Number & ": " & Err.Description & ")"

End Function


Private Function RestoreScreenUpdating( _
    ByVal SavedScreen As Boolean) _
    As String
'
'==============================================================================
'                            RestoreScreenUpdating
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
'==============================================================================
'
        On Error GoTo Failed
        Application.ScreenUpdating = SavedScreen
        Exit Function
Failed:
        RestoreScreenUpdating = "screen updating not restored (error " & Err.Number & ": " & Err.Description & ")"

End Function


Private Function AppendDetail( _
    ByVal Existing As String, _
    ByVal Addition As String) _
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
        If Len(Existing) = 0 Then
            AppendDetail = Addition
        ElseIf Len(Addition) = 0 Then
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
    ByVal CleanupDetails As String) _
    As String
'
'==============================================================================
'                                ReportOutcome
'------------------------------------------------------------------------------
' PURPOSE
'   Tell the user how the operation ended, and build the result line of
'   RunSFT_Silent.
'
' RETURNS
'   "OK" or "STOPPED", the run class, the error and warning counts and the
'   totals per regime, separated by " | ".
'
' ERROR POLICY
'   A cleanup failure is raised in silent mode as ERR_CLEANUP_FAILED, or
'   shown with the outcome otherwise.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Message   As String    'Text for the user
    Dim Totals    As String    'Totals per regime

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
        Select Case Operation
            Case OP_CLEAR
                Message = "Outputs cleared."
                ReportOutcome = "OK | cleared"
            Case Else
                Totals = AppendDetail(CORE_Engine.TotalsText(RG_CRR), CORE_Engine.TotalsText(RG_BCBS))
                If Completed Then
                    ReportOutcome = "OK"
                    If Operation = OP_RUN Then
                        Message = "SFT run completed." & vbLf & vbLf & Replace(Totals, "; ", vbLf)
                    Else
                        Message = "Validation completed: outputs cleared."
                    End If
                Else
                    ReportOutcome = "STOPPED"
                    Message = "The run stopped: see the Checks sheet."
                End If
                ReportOutcome = ReportOutcome & " | " & CORE_Engine.RunClass & " | errors " & CORE_Engine.ErrorCount & _
                                " | warnings " & CORE_Engine.WarningCount
                If Len(Totals) > 0 Then
                    ReportOutcome = ReportOutcome & " | " & Totals
                End If
                Message = Message & vbLf & vbLf & "Errors " & CORE_Engine.ErrorCount & ", warnings " & _
                          CORE_Engine.WarningCount & " (see Checks). Run class: " & CORE_Engine.RunClass & "."
        End Select

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        If Len(CleanupDetails) > 0 Then
            If gSilent Then
                Err.Raise ERR_CLEANUP_FAILED, "M_Main.ReportOutcome", CleanupDetails
            End If
            Message = Message & vbLf & vbLf & "Excel settings not fully restored: " & CleanupDetails
        End If
        Notify Message, IIf(Completed, vbInformation, vbExclamation)

End Function


Private Sub ReportFailure( _
    ByVal ErrorNumber As Long, _
    ByVal ErrorSource As String, _
    ByVal ErrorDescription As String, _
    ByVal CleanupDetails As String)
'
'==============================================================================
'                                ReportFailure
'------------------------------------------------------------------------------
' PURPOSE
'   Show an error that stopped an operation, or raise it in silent mode,
'   with any cleanup failure appended.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Description   As String    'Description with the cleanup failures

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        Description = ErrorDescription
        If Len(CleanupDetails) > 0 Then
            Description = Description & " Excel settings not fully restored: " & CleanupDetails
        End If
        If gSilent Then
            Err.Raise ErrorNumber, ErrorSource, Description
        End If
        Notify "The operation failed." & vbLf & vbLf & "Error " & ErrorNumber & " in " & ErrorSource & ":" & _
               vbLf & Description, vbCritical

End Sub


Private Sub Notify( _
    ByVal Message As String, _
    ByVal Style As VbMsgBoxStyle)
'
'==============================================================================
'                                    Notify
'------------------------------------------------------------------------------
' PURPOSE
'   Show a message box, or print to the Immediate window in silent mode.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If gSilent Then
            Debug.Print MSG_TITLE & ": " & Replace(Message, vbLf, " ")
        Else
            MsgBox Message, Style, MSG_TITLE
        End If

End Sub
