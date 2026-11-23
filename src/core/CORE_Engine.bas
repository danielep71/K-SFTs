Attribute VB_Name = "CORE_Engine"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Engine
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate the exposure value E* of every netting set of securities
'   financing transactions (SFTs) on the input sheets under the financial
'   collateral comprehensive method with supervisory haircuts, its risk
'   weight, risk-weighted exposure amount and own funds requirement, and
'   write the intermediate and final figures to the output sheets.
'
'     with an eligible master netting agreement             [CRE22.65]
'       E* = max(0, sum E - sum C + 0.4 net + 0.6 gross / sqrt(N)
'                   + sum_k |E_k^fx| Hfx)               [CRR Art. 220(3)]
'     without one, per transaction                          [CRE22.46]
'       E* = max(0, E (1 + He) - C (1 - Hc - Hfx))  [CRR Art. 223(2)-(5)]
'     haircuts  H = H10 sqrt((N_R + T_M - 1) / 10)  [CRE22.59; Art. 226]
'     RWA = E* x RW;  own funds requirement = 8% x RWA
'
' WHY THIS EXISTS
'   M_Main needs one routine that turns the input sheets into E*, RWA and
'   own funds requirements and writes every intermediate figure and every
'   rejected input to the output sheets, so that each result can be traced
'   back to its netting set, transactions and legs.
'
' INPUTS
'   None at module level. The engine reads the Params, NettingSets and
'   Transactions sheets (see STATE OWNERSHIP).
'
' RETURNS
'   Nothing at module level. Calculate reports whether the run completed;
'   the results are written to Results, Positions, LegCalc and Checks.
'
' BEHAVIOR
'   Calculate runs, in order:
'     - ClearOutputSheets, ValidateSchema: clear old results, check the
'       layout;
'     - LoadParams, LoadFXTable, LoadHaircutTable, LoadRiskWeightTable,
'       LoadNettingSets: read the inputs;
'     - ClassifyRun: regulatory or what-if run;
'     - LoadLegs: one row per leg, validated, valued and keyed to a haircut;
'     - BuildTransactions: legs grouped into transactions, which are checked
'       as a whole and rejected as a whole;
'     - CheckNettingSets: transaction types, short-term status, zero-haircut
'       conditions, issue consistency, holding period override, risk-weight
'       key and status of each netting set;
'     - ComputeNettingSets: holding period, haircuts, positions and E*;
'     - WriteResults, WritePositions, WriteLegCalc, WriteChecks: outputs;
'     - WriteRunInfo: the run summary.
'
' ERROR POLICY
'   Invalid inputs are not VBA errors. They are logged as ERROR, WARNING or
'   INFO lines for the Checks sheet: an invalid leg rejects its whole
'   transaction and its netting set is INCOMPLETE; an invalid netting set
'   is INVALID; in both cases E* is withheld. An invalid layout, parameter
'   or Params table, or no usable netting set, stops the run. A Checks sheet
'   that cannot be written withdraws the results and raises
'   ERR_CHECKS_WRITE. Any other unexpected error withdraws the results, is
'   logged on Checks and in the run summary, and is raised unchanged to
'   M_Main.
'
' DEPENDENCIES
'   CORE_Config for the layout, CORE_Util for conversions and sheet access,
'   SFT_Formulas for the holding period, haircut scaling, both exposure
'   formulas, the net and gross exposures, the issue count and the own
'   funds requirement.
'
' NOTES
'   Each netting set, transaction, leg, issue and currency position is
'   stored once in an array, with a Collection (CORE_Util.KeyIndex) mapping
'   its key to its array position.
'
' REGIME LAYER
'   CRR is the only supported regime (D1; the support matrix, #248). The
'   regime is set on Params and can be set per netting set, and a table row
'   applies to BOTH or CRR; a BCBS value anywhere is refused. The
'   two-regime totals stay in the code until the core is separated from
'   the worksheets (#249); they only ever hold CRR. The differences between
'   CRE22 and the CRR are documented on the Basel_vs_CRR sheet.
'
' PUBLIC SURFACE
'   None outside this VBA project. Calculate, TotalsText, RunClass and the
'   run counters are Public for M_Main; Option Private Module keeps them
'   off the supported external surface.
'
' STATE OWNERSHIP
'   Owns the module state below. ResetState clears it at the start of each
'   run; the run counters stay readable until the next run. Reads the input
'   sheets and Params; writes Results, Positions, LegCalc and Checks.
'   Changes no Excel application setting.
'
' KNOWN DEVIATION
'   This module lives in src/core but reads and writes worksheets, which
'   the repository structure reserves for src/workbook, calls SFT_Formulas
'   in src/modules, and keeps module state (PRIVATE STATE below), which the
'   house style excludes from src/core.
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
' PRIVATE TYPES
'------------------------------------------------------------------------------
    'One netting set, as read from the NettingSets sheet, with its results.
    'Amounts are in the reporting currency
    Private Type tNettingSet
        ID                  As String   'Netting-set ID as entered
        Row                 As Long     'Sheet row
        Counterparty        As String   'Counterparty name
        MNA                 As Boolean  'Covered by an eligible master netting agreement
        SettleCcy           As String   'Settlement currency of the agreement
        CptyClass           As String   'SOV, INST or CORP; QCCP is refused
        CQS                 As Long     'Credit quality step 1-6; 0 unrated
        SCRA                As String   'SCRA grade of an unrated institution
        RemarginBD          As Double   'Remargining frequency N_R, business days
        LargeOrIlliquid     As Boolean  'Over 5,000 trades or illiquid collateral
        Disputes            As Boolean  'Margin-call disputes: holding period doubled
        HoldingOverride     As Double   'Holding period override, business days; 0 none
        ZeroHaircut         As Boolean  'Conditions of the 0% haircut declared met
        HasRWOverride       As Boolean  'A risk weight override is given
        RWOverride          As Double   'Risk weight override, decimal
        Regime              As String   'CRR; BCBS is refused
        InputErrors         As Long     'Errors that make the netting set INVALID
        TxCount             As Long     'Transactions read for this netting set
        TxRejected          As Long     'Transactions rejected
        LegsUsed            As Long     'Legs recognised in the calculation
        HasRepoStyle        As Boolean  'Holds a repo-style transaction
        HasCapitalMarket    As Boolean  'Holds an other capital-market-driven transaction
        AllShortTerm        As Boolean  'Every transaction of 3 months or less
        HoldingBD           As Double   'Holding period T_M; the longest without netting
        HaircutScale        As Double   'sqrt((N_R + T_M - 1) / base); the longest without netting
        SumE                As Double   'Value lent, sold or posted
        SumC                As Double   'Value borrowed, bought or received
        NetExposure         As Double   '|sum E_s H_s|
        GrossExposure       As Double   'sum |E_s| H_s
        IssueCount          As Long     'N
        SystematicTerm      As Double   'NetWeight x net exposure
        IdiosyncraticTerm   As Double   'GrossWeight x gross exposure / sqrt(N)
        FXTerm              As Double   'Currency-mismatch term
        EStar               As Double   'Exposure value E*
        RWKey               As String   'Risk-weight key
        RiskWeight          As Double   'Risk weight applied
        RWA                 As Double   'Risk-weighted exposure amount
        OwnFunds            As Double   'Own funds requirement
        Status              As String   'VALID, INVALID, INCOMPLETE or NO TRANSACTIONS
    End Type

    'One leg: cash or one security issue lent or received in a transaction
    Private Type tLeg
        Row             As Long     'Sheet row
        LegID           As String   'Leg ID as entered
        TxKey           As String   'Transaction ID, upper case
        TxID            As String   'Transaction ID as entered
        TxIndex         As Long     'Index in mTx
        TxRejected      As Boolean  'Its transaction was rejected
        NsIndex         As Long     'Index in mNS; 0 when unknown
        TxType          As String   'Transaction type
        StartSerial     As Double   'Start date serial; -1 invalid
        MatSerial       As Double   'Maturity date serial; 0 open
        Side            As String   'LENT or RECEIVED
        Sign            As Long     '+1 lent, -1 received
        Asset           As String   'Asset type
        IssueKey        As String   'Issue ID, upper case; CASH and the currency for cash
        Issuer          As String   'Issuer type of a debt security
        CQS             As Long     'Credit quality step of the issue; 0 unrated
        SecMatSerial    As Double   'Security maturity date serial
        Listing         As String   'Equity listing
        FundH10         As Double   'Fund haircut H10
        MarketValue     As Double   'Market value in the leg currency
        Ccy             As String   'Leg currency
        Value           As Double   'Market value in the reporting currency
        ResidualYears   As Double   'Residual maturity of a debt security; -1 n/a
        HaircutKey      As String   'Key of the haircut-table row; FUND for a fund
        Eligible        As Boolean  'Eligible as collateral
        Recognised      As Boolean  'Used in the calculation
        H10             As Double   'Haircut before scaling
        Signature       As String   'Description of the issue, compared across its legs
        Valid           As Boolean  'No input error on this leg
        HoldingBD       As Double   'Holding period applied
        Scale           As Double   'Haircut scale applied
        H               As Double   'Scaled haircut
        Hfx             As Double   'Currency haircut, single-transaction formula only
        Adjusted        As Double   'V(1 + H) lent, V(1 - H - Hfx) received
        Messages        As String   'Warnings, or why the leg was excluded
    End Type

    'One transaction: the legs sharing a Transaction ID
    Private Type tTransaction
        ID              As String   'Transaction ID as entered
        NsIndex         As Long     'Netting set of its first leg; 0 unknown
        TxType          As String   'Transaction type of its first leg
        StartSerial     As Double   'Start date of its first leg
        MatSerial       As Double   'Maturity of its first leg; 0 open
        FirstRow        As Long     'Sheet row of its first leg
        LegCount        As Long     'Legs read
        LentCount       As Long     'Legs lent
        ReceivedCount   As Long     'Legs received
        LentNonCash     As Long     'Non-cash legs lent
        ReceivedNonCash As Long     'Non-cash legs received
        LentCcy         As String   'Currency of the legs lent; "*" when several
        Rejected        As Boolean  'Rejected as a whole
        ShortTerm       As Boolean  'Original maturity of 3 months or less
        HoldingBD       As Double   'Holding period, without netting
        E               As Double   'Value lent
        C               As Double   'Value received and recognised
        EAdjusted       As Double   'sum V (1 + H) lent
        CAdjusted       As Double   'sum V (1 - H - Hfx) received
        FXTerm          As Double   'sum V Hfx received
        EStar           As Double   'Exposure value without netting
    End Type

    'One security issue or currency of a netting set with netting
    Private Type tPosition
        NsIndex     As Long     'Netting set
        Kind        As String   'ISSUE or CURRENCY
        Key         As String   'Issue ID or currency, upper case
        Asset       As String   'Asset type; blank for a currency
        Ccy         As String   'Currency
        Legs        As Long     'Legs aggregated
        Lent        As Double   'Value lent
        Received    As Double   'Value received
        H           As Double   'Scaled haircut, or the currency haircut
        Counted     As Boolean  'Counted in N
    End Type

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    'Parameters of the run
    Private pAsOf       As Double   'Reporting date serial
    Private pRepCcy     As String   'Reporting currency
    Private pRegime     As String   'Default regime
    Private pDaysYear   As Double   'Days per year
    Private pHcBase     As Double   'Holding period of the H10 table
    Private pHoldRepo   As Double   'Minimum holding period, repo-style
    Private pHoldCMT    As Double   'Minimum holding period, other capital market
    Private pHoldLarge  As Double   'Holding period, large or illiquid
    Private pHfx        As Double   'Currency haircut H10
    Private pNonElig    As Double   'Haircut H10 of non-eligible securities lent
    Private pNetW       As Double   'Net weight
    Private pGrossW     As Double   'Gross weight
    Private pThreshold  As Double   'Issue-count threshold
    Private pCapRatio   As Double   'Own funds per unit of RWA

    'Lookup tables of Params
    Private mHcKeys()       As String       'Haircut keys
    Private mHcValues()     As Double       'H10
    Private mHcRegimes()    As String       'BOTH or CRR
    Private mHcCount        As Long         'Rows
    Private mHcIndex        As Collection   'Key to row
    Private mRwKeys()       As String       'Risk-weight keys
    Private mRwValues()     As Double       'Risk weights
    Private mRwRegimes()    As String       'BOTH or CRR
    Private mRwCount        As Long         'Rows
    Private mRwIndex        As Collection   'Key to row
    Private mFxCcy()        As String       'Currencies
    Private mFxRate()       As Double       'Units of reporting currency per unit
    Private mFxCount        As Long         'Rows
    Private mFxIndex        As Collection   'Currency to row

    'Inputs and intermediate results
    Private mNS()       As tNettingSet  'Netting sets
    Private mNsCount    As Long         'Netting sets read
    Private mNsIndex    As Collection   'Netting-set ID to index in mNS
    Private mLegs()     As tLeg         'Legs
    Private mLegCount   As Long         'Legs read
    Private mLegIndex   As Collection   'Leg ID to index in mLegs
    Private mTx()       As tTransaction 'Transactions
    Private mTxCount    As Long         'Transactions built
    Private mTxIndex    As Collection   'Transaction ID to index in mTx
    Private mPos()      As tPosition    'Issue and currency positions
    Private mPosCount   As Long         'Positions built
    Private mPosIndex   As Collection   'Position key to index in mPos

    'Checks messages, stored (field, message) so that ReDim Preserve can grow
    'the last dimension, and the run counters
    Private mLog()                  As Variant  'Checks messages, (field, message)
    Private mLogCount               As Long     'Messages logged
    Private mErrCount               As Long     'ERROR messages logged
    Private mWarnCount              As Long     'WARNING messages logged
    Private mRunClass               As String   'RUN_REGULATORY, RUN_WHATIF or ""
    Private mDeviations             As Long     'Differences from the regulatory set
    Private mTxUsed                 As Long     'Transactions in VALID netting sets
    Private mLegsUsed               As Long     'Legs recognised in VALID netting sets
    Private mWithheld               As Long     'Netting sets whose E* is withheld
    Private mRegEStar(1 To 2)       As Double   'Total E* per regime: 1 CRR, 2 BCBS
    Private mRegRWA(1 To 2)         As Double   'Total RWA per regime
    Private mRegOwnFunds(1 To 2)    As Double   'Total own funds requirement per regime
    Private mRegSets(1 To 2)        As Long     'Netting sets per regime
    Private mRegValid(1 To 2)       As Long     'VALID netting sets per regime


'
'------------------------------------------------------------------------------
'
'                                     RUN
'
'------------------------------------------------------------------------------
'

Public Function Calculate( _
    ByVal WriteOutputs As Boolean) As Boolean
'
'------------------------------------------------------------------------------
'                                RUN THE ENGINE
'------------------------------------------------------------------------------
' PURPOSE
'   Run the engine once: load the inputs, calculate every leg, transaction
'   and netting set, and write the outputs.
'
' WHY THIS EXISTS
'   M_Main starts every run through this one entry point, which fixes the
'   order of the steps and withdraws the results after a failure.
'
' INPUTS
'   WriteOutputs
'     True writes every output sheet; False validates only: it clears the
'     output tables and writes the Checks sheet and the run summary.
'
' RETURNS
'   True when the run completed. Individual transactions may still have
'   been rejected; the Checks sheet lists them. False when an input sheet,
'   input table or Params layout is not as expected, or a parameter, a
'   Params table or the netting sets could not be loaded, or an output
'   sheet or output header name is missing.
'
' BEHAVIOR
'   - Resets the module state and clears the old outputs first, for
'     validation too, so that a failed or validation-only run can never
'     leave earlier results looking current
'   - Checks the layout, loads the parameters, the FX, haircut and
'     risk-weight tables and the netting sets, and stops at the first step
'     that fails
'   - Classifies the run, loads the legs, builds the transactions, checks
'     and computes the netting sets
'   - Writes Results, Positions and LegCalc when WriteOutputs is True
'   - Writes the Checks sheet and the run summary in every case
'
' ERROR POLICY
'   Raises ERR_CHECKS_WRITE when the Checks sheet cannot be written, after
'   clearing the other output sheets. Any other error clears the output
'   sheets, is logged on Checks and in the run summary, and is raised with
'   its original number, source and description.
'
' DEPENDENCIES
'   The load, check, compute and write procedures of this module;
'   CORE_Config for the severities and ERR_CHECKS_WRITE.
'
' NOTES
'   The source of ERR_CHECKS_WRITE is the literal "CORE_Engine.Calculate".
'
' STATE OWNERSHIP
'   Resets all module state, then fills it for this run.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Engine.Calculate"

    Dim StartTime       As Double   'Timer value at the start
    Dim ChecksFailure   As String   'Why the Checks sheet could not be written
    Dim ErrNumber       As Long     'Unexpected error: number
    Dim ErrSource       As String   'Unexpected error: source
    Dim ErrDescription  As String   'Unexpected error: description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Note the start time and clear the state of the previous run
        StartTime = Timer
        ResetState
    'Enable controlled error handling
        On Error GoTo Failed

'------------------------------------------------------------------------------
' LOAD INPUTS
'------------------------------------------------------------------------------
    'Clear old outputs first, for validation too, so that a failed or
    'validation-only run can never leave earlier results looking current
        ClearOutputSheets
    'Check the layout and load the inputs; stop at the first step that fails
        If Not ValidateSchema() Then GoTo Finish
        If Not LoadParams() Then GoTo Finish
        If Not LoadFXTable() Then GoTo Finish
        If Not LoadHaircutTable() Then GoTo Finish
        If Not LoadRiskWeightTable() Then GoTo Finish
        If Not LoadNettingSets() Then GoTo Finish
    'Classify the run as regulatory or what-if
        ClassifyRun

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Read the legs, build the transactions, check and compute the netting sets
        LoadLegs
        BuildTransactions
        CheckNettingSets
        ComputeNettingSets
    'Write the detailed outputs unless the run only validates
        If WriteOutputs Then
            WriteResults
            WritePositions
            WriteLegCalc
        End If
    'Report the run as completed
        Calculate = True

'------------------------------------------------------------------------------
' WRITE CHECKS AND RUN SUMMARY
'------------------------------------------------------------------------------
Finish:
    'Withdraw the results and raise when the Checks sheet cannot be written
        If Not WriteChecks(ChecksFailure) Then
            Calculate = False
            ClearOutputSheets
            WriteRunInfo Timer - StartTime, False, WriteOutputs, "the Checks sheet could not be written (" & _
                ChecksFailure & "); results withdrawn"
            On Error GoTo 0
            Err.Raise ERR_CHECKS_WRITE, "CORE_Engine.Calculate", _
                "The Checks sheet could not be written (" & ChecksFailure & _
                "); results withdrawn. Unprotect or repair the Checks sheet and run again."
        End If
    'Write the run summary
        WriteRunInfo Timer - StartTime, Calculate, WriteOutputs, vbNullString

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Capture the original error
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Withdraw everything written so far and log the original error, both
    'best effort
        Calculate = False
        On Error Resume Next
        ClearOutputSheets
        LogMsg SEV_ERROR, vbNullString, vbNullString, "Run stopped by error " & ErrNumber & ": " & _
            ErrDescription & " - results withdrawn."
        WriteChecks ChecksFailure
        WriteRunInfo Timer - StartTime, False, WriteOutputs, "error " & ErrNumber & ": " & ErrDescription
    'Restore normal error handling
        On Error GoTo 0
    'Raise the original error unchanged
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function

Public Property Get ErrorCount() As Long

'
'------------------------------------------------------------------------------
'                         COUNT ERRORS OF THE LAST RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Number of ERROR messages logged by the last run.
'
' WHY THIS EXISTS
'   M_Main reports the outcome of a run without reading the Checks sheet.
'
' INPUTS
'   None
'
' RETURNS
'   The ERROR count; 0 before the first run.
'
' BEHAVIOR
'   Returns the counter kept by LogMsg.
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
    'Return the counter of the last run
        ErrorCount = mErrCount

End Property

Public Property Get WarningCount() As Long

'
'------------------------------------------------------------------------------
'                        COUNT WARNINGS OF THE LAST RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Number of WARNING messages logged by the last run.
'
' WHY THIS EXISTS
'   M_Main reports the outcome of a run without reading the Checks sheet.
'
' INPUTS
'   None
'
' RETURNS
'   The WARNING count; 0 before the first run.
'
' BEHAVIOR
'   Returns the counter kept by LogMsg.
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
    'Return the counter of the last run
        WarningCount = mWarnCount

End Property

Public Property Get RunClass() As String

'
'------------------------------------------------------------------------------
'                            CLASS OF THE LAST RUN
'------------------------------------------------------------------------------
' PURPOSE
'   REGULATORY or WHAT-IF for the last run; "" when it was not classified.
'
' WHY THIS EXISTS
'   M_Main reports the outcome of a run without reading the run summary.
'
' INPUTS
'   None
'
' RETURNS
'   The class set by ClassifyRun; "" when the run stopped before it.
'
' BEHAVIOR
'   Returns the class kept in the module state.
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
    'Return the class of the last run
        RunClass = mRunClass

End Property

Public Function TotalsText( _
    ByVal RegimeCode As String) As String
'
'------------------------------------------------------------------------------
'                      DESCRIBE THE TOTALS OF ONE REGIME
'------------------------------------------------------------------------------
' PURPOSE
'   One line with the totals of one regime of the last run, for the message
'   M_Main shows: E*, RWA and own funds requirement over the VALID netting
'   sets, and how many netting sets were withheld. Totals of different
'   regimes are never added together.
'
' WHY THIS EXISTS
'   M_Main shows the totals per regime without reading the output sheets.
'
' INPUTS
'   RegimeCode
'     RG_CRR or RG_BCBS.
'
' RETURNS
'   The line; "" when the run had no netting set of that regime.
'
' BEHAVIOR
'   Finds the totals slot of the regime and formats its E*, RWA, own funds
'   and VALID and total netting-set counts.
'
' ERROR POLICY
'   None raised by the procedure itself.
'
' DEPENDENCIES
'   RegimeIndex
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
    Const PROC_NAME As String = "CORE_Engine.TotalsText"

    Dim RegimeSlot  As Long 'Regime index

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
    'Find the totals slot of the regime
        RegimeSlot = RegimeIndex(RegimeCode)
    'Return an empty line when the run had no netting set of that regime
        If mRegSets(RegimeSlot) = 0 Then
            Exit Function
        End If
    'Compose the line from the totals of the regime
        TotalsText = RegimeCode & ": E* " & Format$(mRegEStar(RegimeSlot), "#,##0") & _
            ", RWA " & Format$(mRegRWA(RegimeSlot), "#,##0") & _
            ", own funds " & Format$(mRegOwnFunds(RegimeSlot), "#,##0") & _
            " (" & mRegValid(RegimeSlot) & " of " & mRegSets(RegimeSlot) & " netting set(s) VALID)"

End Function

Private Function RegimeIndex( _
    ByVal RegimeCode As String) As Long
'
'------------------------------------------------------------------------------
'                       FIND THE TOTALS SLOT OF A REGIME
'------------------------------------------------------------------------------
' PURPOSE
'   Index of a regime in the per-regime totals: 1 for CRR, 2 for BCBS.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Use slot 2 for BCBS and slot 1 for every other code
        If RegimeCode = RG_BCBS Then
            RegimeIndex = 2
        Else
            RegimeIndex = 1
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                                STATE AND LOG
'
'------------------------------------------------------------------------------
'

Private Sub ResetState()

'
'------------------------------------------------------------------------------
'                            CLEAR THE MODULE STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the module state before a run.
'
' WHY THIS EXISTS
'   The module state outlives a run; nothing of the previous run may leak
'   into the next one.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Allocates an empty message log of 64 messages
'   - Sets every counter to 0 and the run class to ""
'   - Replaces every key index with a new Collection
'   - Sets the per-regime totals to 0
'
' ERROR POLICY
'   None raised by the procedure itself.
'
' DEPENDENCIES
'   CORE_Config for CK_NCOLS.
'
' NOTES
'   The arrays of netting sets, legs, transactions and positions are not
'   erased; only their counters are reset.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.ResetState"

    Dim RegimeSlot  As Long 'Regime index

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
    'Start an empty message log
        ReDim mLog(1 To CK_NCOLS, 1 To 64)
    'Reset the counters and the run class
        mLogCount = 0
        mErrCount = 0
        mWarnCount = 0
        mRunClass = vbNullString
        mDeviations = 0
        mHcCount = 0
        mRwCount = 0
        mFxCount = 0
        mNsCount = 0
        mLegCount = 0
        mTxCount = 0
        mPosCount = 0
        mTxUsed = 0
        mLegsUsed = 0
        mWithheld = 0
    'Start empty key indexes
        Set mHcIndex = New Collection
        Set mRwIndex = New Collection
        Set mFxIndex = New Collection
        Set mNsIndex = New Collection
        Set mLegIndex = New Collection
        Set mTxIndex = New Collection
        Set mPosIndex = New Collection
    'Clear the totals of each regime
        For RegimeSlot = 1 To 2
            mRegEStar(RegimeSlot) = 0#
            mRegRWA(RegimeSlot) = 0#
            mRegOwnFunds(RegimeSlot) = 0#
            mRegSets(RegimeSlot) = 0
            mRegValid(RegimeSlot) = 0
        Next RegimeSlot

End Sub

Private Sub LogMsg( _
    ByVal Severity As String, _
    ByVal SheetName As String, _
    ByVal Reference As String, _
    ByVal Message As String, _
    Optional ByVal RowNumber As Long = 0)
'
'------------------------------------------------------------------------------
'                      LOG A MESSAGE FOR THE CHECKS SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Add one message for the Checks sheet and count errors and warnings.
'
' WHY THIS EXISTS
'   Invalid inputs are reported, not raised; every check of the engine
'   reports through this one log.
'
' INPUTS
'   Severity
'     SEV_ERROR, SEV_WARN or SEV_INFO.
'
'   SheetName
'     Sheet the message refers to.
'
'   Reference
'     The ID or parameter code concerned; may be empty.
'
'   Message
'     The message text.
'
'   RowNumber
'     Sheet row concerned; 0 leaves the Row column blank.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Doubles the log when it is full
'   - Stores severity, sheet, row, reference and message
'   - Counts ERROR and WARNING messages
'
' ERROR POLICY
'   None raised by the procedure itself.
'
' DEPENDENCIES
'   CORE_Config for CK_NCOLS and the severities.
'
' NOTES
'   The log is stored (field, message) so that ReDim Preserve can grow the
'   last dimension.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' STORE MESSAGE
'------------------------------------------------------------------------------
    'Take the next slot, doubling the log when it is full
        mLogCount = mLogCount + 1
        If mLogCount > UBound(mLog, 2) Then
            ReDim Preserve mLog(1 To CK_NCOLS, 1 To mLogCount * 2)
        End If
    'Store the severity and the sheet
        mLog(1, mLogCount) = Severity
        mLog(2, mLogCount) = SheetName
    'Store the row, or leave it blank when there is none
        If RowNumber > 0 Then
            mLog(3, mLogCount) = RowNumber
        Else
            mLog(3, mLogCount) = Empty
        End If
    'Store the reference and the message text
        mLog(4, mLogCount) = Reference
        mLog(5, mLogCount) = Message

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Count errors and warnings
        If Severity = SEV_ERROR Then
            mErrCount = mErrCount + 1
        ElseIf Severity = SEV_WARN Then
            mWarnCount = mWarnCount + 1
        End If

End Sub

Private Sub LegError( _
    ByVal LegIndex As Long, _
    ByVal Message As String)
'
'------------------------------------------------------------------------------
'                                 REJECT A LEG
'------------------------------------------------------------------------------
' PURPOSE
'   Reject a leg: log an ERROR on Transactions and keep the reason for its
'   LegCalc row. Its transaction is then rejected as a whole.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REJECT
'------------------------------------------------------------------------------
    'Mark the leg invalid, keep the reason and log it
        With mLegs(LegIndex)
            .Valid = False
            .Messages = AppendText(.Messages, Message)
            LogMsg SEV_ERROR, SH_TX, .LegID, Message & " - transaction " & .TxID & " rejected.", .Row
        End With

End Sub

Private Sub LegNote( _
    ByVal LegIndex As Long, _
    ByVal Severity As String, _
    ByVal Message As String)
'
'------------------------------------------------------------------------------
'                            NOTE A REMARK ON A LEG
'------------------------------------------------------------------------------
' PURPOSE
'   Log a WARNING or INFO about a leg that stays in the calculation, and
'   keep it for its LegCalc row.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' NOTE
'------------------------------------------------------------------------------
    'Keep the remark for the LegCalc row and log it
        With mLegs(LegIndex)
            .Messages = AppendText(.Messages, Message)
            LogMsg Severity, SH_TX, .LegID, Message, .Row
        End With

End Sub

Private Sub NsError( _
    ByVal NsIndex As Long, _
    ByVal Message As String)
'
'------------------------------------------------------------------------------
'                           INVALIDATE A NETTING SET
'------------------------------------------------------------------------------
' PURPOSE
'   Log an ERROR that makes a netting set INVALID, so that its E* is
'   withheld.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INVALIDATE
'------------------------------------------------------------------------------
    'Count the input error and log it
        With mNS(NsIndex)
            .InputErrors = .InputErrors + 1
            LogMsg SEV_ERROR, SH_NS, .ID, Message & " - exposure withheld.", .Row
        End With

End Sub

Private Function AppendText( _
    ByVal Existing As String, _
    ByVal Addition As String) As String
'
'------------------------------------------------------------------------------
'                            JOIN TWO MESSAGE TEXTS
'------------------------------------------------------------------------------
' PURPOSE
'   Join two message texts with "; ".
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' JOIN
'------------------------------------------------------------------------------
    'Return the addition alone when there is nothing yet
        If VBA.Len(Existing) = 0 Then
            AppendText = Addition
        Else
            AppendText = Existing & "; " & Addition
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                                 SCHEMA CHECK
'
'------------------------------------------------------------------------------
'

Private Function ValidateSchema() As Boolean

'
'------------------------------------------------------------------------------
'                          CHECK THE WORKBOOK LAYOUT
'------------------------------------------------------------------------------
' PURPOSE
'   Check the workbook layout before anything is read: every sheet, every
'   workbook name that locates a table or a parameter, and the header text
'   of every column the engine reads.
'
' WHY THIS EXISTS
'   The engine reads by column number, so a header that differs means a
'   column was inserted, deleted or moved.
'
' INPUTS
'   None
'
' RETURNS
'   True when the layout is as expected; otherwise False with the problems
'   found logged (missing sheets first; then the first header mismatch of
'   each table), and the run stops.
'
' BEHAVIOR
'   - Checks that every sheet exists; stops when one is missing
'   - Compares the headers of the input tables and the Params tables
'   - Checks the header names of the output sheets
'   - Checks the workbook name of every parameter
'
' ERROR POLICY
'   Layout problems are logged as ERROR messages, not raised.
'
' DEPENDENCIES
'   SheetPresent, HeaderProblems, NamedHeaderProblems, ParameterCodes;
'   CORE_Config for the sheet, header and name constants; CORE_Util for
'   TableHeaderRow, TableHeaderName and NamedCell.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.ValidateSchema"

    Dim SheetNames  As Variant  'Sheets the engine needs
    Dim Codes       As Variant  'Parameter codes
    Dim Problems    As Long     'Problems found
    Dim ItemIndex   As Long     'Position in a list

'------------------------------------------------------------------------------
' SHEETS
'------------------------------------------------------------------------------
    'Log every missing sheet
        SheetNames = Array(SH_PARAMS, SH_NS, SH_TX, SH_RESULTS, SH_POSITIONS, SH_LEGCALC, SH_CHECKS)
        For ItemIndex = LBound(SheetNames) To UBound(SheetNames)
            If Not SheetPresent(VBA.CStr(SheetNames(ItemIndex))) Then
                LogMsg SEV_ERROR, VBA.CStr(SheetNames(ItemIndex)), vbNullString, _
                    "Sheet missing - rebuild the workbook from the template."
                Problems = Problems + 1
            End If
        Next ItemIndex
    'Stop when a sheet is missing
        If Problems > 0 Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' TABLE HEADERS
'------------------------------------------------------------------------------
    'Compare the headers of the input tables and the Params tables
        Problems = Problems + HeaderProblems(SH_NS, TableHeaderRow(SH_NS), NS_HEADERS, NM_NS_HEADER)
        Problems = Problems + HeaderProblems(SH_TX, TableHeaderRow(SH_TX), TX_HEADERS, NM_TX_HEADER)
        Problems = Problems + NamedHeaderProblems(NM_HC_TABLE, HC_HEADERS)
        Problems = Problems + NamedHeaderProblems(NM_RW_TABLE, RW_HEADERS)
        Problems = Problems + NamedHeaderProblems(NM_FX_TABLE, FX_HEADERS)
        Problems = Problems + NamedHeaderProblems(NM_PRM_CODE, PRM_CODE_HEADER)
        Problems = Problems + NamedHeaderProblems(NM_PRM_VALUE, PRM_VALUE_HEADER)
    'Check the header name of every output sheet
        SheetNames = Array(SH_RESULTS, SH_POSITIONS, SH_LEGCALC, SH_CHECKS)
        For ItemIndex = LBound(SheetNames) To UBound(SheetNames)
            If TableHeaderRow(VBA.CStr(SheetNames(ItemIndex))) = 0 Then
                LogMsg SEV_ERROR, VBA.CStr(SheetNames(ItemIndex)), vbNullString, "Workbook name '" & _
                    TableHeaderName(VBA.CStr(SheetNames(ItemIndex))) & _
                    "' is missing or no longer refers to the first header cell in column C."
                Problems = Problems + 1
            End If
        Next ItemIndex

'------------------------------------------------------------------------------
' PARAMETER NAMES
'------------------------------------------------------------------------------
    'Check the workbook name of every parameter
        Codes = ParameterCodes()
        For ItemIndex = LBound(Codes) To UBound(Codes)
            If NamedCell(VBA.CStr(Codes(ItemIndex))) Is Nothing Then
                LogMsg SEV_ERROR, SH_PARAMS, VBA.CStr(Codes(ItemIndex)), "Workbook name of parameter missing or broken."
                Problems = Problems + 1
            End If
        Next ItemIndex
    'Pass only when no problem was found
        ValidateSchema = (Problems = 0)

End Function

Private Function ParameterCodes() As Variant

'
'------------------------------------------------------------------------------
'                           LIST THE PARAMETER CODES
'------------------------------------------------------------------------------
' PURPOSE
'   The codes of every parameter the engine reads, in Params order.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the codes in Params order
        ParameterCodes = Array(PRM_ASOF, PRM_REPCCY, PRM_REGIME, PRM_DAYSYEAR, PRM_HCBASE, PRM_HOLD_REPO, _
            PRM_HOLD_CMT, PRM_HOLD_LARGE, PRM_HFX, PRM_NONELIG, PRM_NETW, PRM_GROSSW, _
            PRM_NTHRESH, PRM_CAPRATIO)

End Function

Private Function SheetPresent( _
    ByVal SheetName As String) As Boolean
'
'------------------------------------------------------------------------------
'                         TELL WHETHER A SHEET EXISTS
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether this workbook has a worksheet with the given tab name.
'
' NOTES
'   The tab name is compared ignoring case.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.SheetPresent"

    Dim Sheet       As Worksheet    'Each worksheet

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Return True at the first worksheet with that tab name
        For Each Sheet In ThisWorkbook.Worksheets
            If VBA.StrComp(Sheet.Name, SheetName, vbTextCompare) = 0 Then
                SheetPresent = True
                Exit Function
            End If
        Next Sheet

End Function

Private Function HeaderProblems( _
    ByVal SheetName As String, _
    ByVal HeaderRow As Long, _
    ByVal Expected As String, _
    ByVal HeaderName As String) As Long
'
'------------------------------------------------------------------------------
'                    COMPARE THE HEADERS OF AN INPUT TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the header row of an input table with the expected headers,
'   ignoring case, spaces and line breaks; an empty expected entry is a
'   column that is not read.
'
' WHY THIS EXISTS
'   The engine reads input columns by number; a moved column must stop the
'   run before anything is read.
'
' INPUTS
'   SheetName
'     The input sheet.
'
'   HeaderRow
'     Sheet row of the headers; 0 when the workbook name is missing.
'
'   Expected
'     The expected headers, separated by "|", from column 1 of the table.
'
'   HeaderName
'     The workbook name that locates the header row, for the message.
'
' RETURNS
'   0 when they match; 1 with the first difference logged otherwise.
'
' BEHAVIOR
'   Logs a missing header name; otherwise compares the headers column by
'   column and logs the first difference.
'
' ERROR POLICY
'   Differences are logged as ERROR messages, not raised.
'
' DEPENDENCIES
'   HeaderKey, LogMsg; CORE_Util for GetSheet, TableCell and SafeStr.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.HeaderProblems"

    Dim Parts()     As String       'Expected headers, 0-based
    Dim Sheet       As Worksheet    'The input sheet
    Dim Found       As String       'Header in the sheet
    Dim Col         As Long         'Position in Parts; column Col + 1

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Report a missing or broken header name
        If HeaderRow = 0 Then
            LogMsg SEV_ERROR, SheetName, vbNullString, "Workbook name '" & HeaderName & _
                "' is missing or no longer refers to the first header cell in column C."
            HeaderProblems = 1
            Exit Function
        End If
    'Split the expected headers
        Set Sheet = GetSheet(SheetName)
        Parts = Split(Expected, "|")
    'Compare every read column and report the first difference
        For Col = 0 To UBound(Parts)
            If VBA.Len(Parts(Col)) > 0 Then
                Found = SafeStr(TableCell(Sheet, HeaderRow, Col + 1).Value)
                If HeaderKey(Found) <> HeaderKey(Parts(Col)) Then
                    LogMsg SEV_ERROR, SheetName, vbNullString, "Column " & (Col + 1) & " header is '" & Found & _
                        "', expected '" & Parts(Col) & "': a column was inserted, deleted or moved.", HeaderRow
                    HeaderProblems = 1
                    Exit Function
                End If
            End If
        Next Col

End Function

Private Function NamedHeaderProblems( _
    ByVal NameText As String, _
    ByVal Expected As String) As Long
'
'------------------------------------------------------------------------------
'                    COMPARE THE HEADERS OF A PARAMS TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the cells from a workbook name's top-left cell rightwards with
'   the expected headers of a Params table or parameter column.
'
' WHY THIS EXISTS
'   The Params tables are located by workbook names and read by offset; a
'   moved column must stop the run before anything is read.
'
' INPUTS
'   NameText
'     The workbook name of the table or column.
'
'   Expected
'     The expected headers, separated by "|"; an empty entry is not read.
'
' RETURNS
'   0 when they match; 1 with the problem logged otherwise.
'
' BEHAVIOR
'   Logs a missing workbook name; otherwise compares the headers cell by
'   cell and logs the first difference.
'
' ERROR POLICY
'   Differences are logged as ERROR messages, not raised.
'
' DEPENDENCIES
'   HeaderKey, LogMsg; CORE_Util for NamedCell and SafeStr.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.NamedHeaderProblems"

    Dim Cell        As Range    'The name's top-left cell
    Dim Parts()     As String   'Expected headers, 0-based
    Dim Found       As String   'Header in the sheet
    Dim Col         As Long     'Position in Parts

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Report a missing or broken workbook name
        Set Cell = NamedCell(NameText)
        If Cell Is Nothing Then
            LogMsg SEV_ERROR, SH_PARAMS, NameText, "Workbook name missing or broken - rebuild from the template."
            NamedHeaderProblems = 1
            Exit Function
        End If
    'Split the expected headers
        Parts = Split(Expected, "|")
    'Compare every read cell and report the first difference
        For Col = 0 To UBound(Parts)
            If VBA.Len(Parts(Col)) > 0 Then
                Found = SafeStr(Cell.Offset(0, Col).Value)
                If HeaderKey(Found) <> HeaderKey(Parts(Col)) Then
                    LogMsg SEV_ERROR, SH_PARAMS, NameText, "Header '" & Found & "' found where '" & Parts(Col) & _
                        "' was expected.", Cell.Row
                    NamedHeaderProblems = 1
                    Exit Function
                End If
            End If
        Next Col

End Function

Private Function HeaderKey( _
    ByVal Header As String) As String
'
'------------------------------------------------------------------------------
'                              NORMALISE A HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   Normalise a header for comparison: upper case, without spaces, tabs or
'   line breaks.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Remove spaces, tabs and line breaks, then convert to upper case
        HeaderKey = VBA.UCase$(Replace(Replace(Replace(Replace(Header, " ", vbNullString), vbTab, vbNullString), _
            vbCr, vbNullString), vbLf, vbNullString))

End Function

'
'------------------------------------------------------------------------------
'
'                                  PARAMETERS
'
'------------------------------------------------------------------------------
'

Private Function LoadParams() As Boolean

'
'------------------------------------------------------------------------------
'                             READ THE PARAMETERS
'------------------------------------------------------------------------------
' PURPOSE
'   Read and check every parameter on Params.
'
' WHY THIS EXISTS
'   Every later step relies on the parameters; an unusable one must stop
'   the run before any leg is read.
'
' INPUTS
'   None
'
' RETURNS
'   True when every parameter is usable; otherwise False with each problem
'   logged, and the run stops.
'
' BEHAVIOR
'   - Reads the reporting date and refuses one that is missing or before
'     FIRST_ASOF
'   - Reads the reporting currency and refuses a code that is not 3 letters
'   - Reads the regime and refuses anything but CRR
'   - Reads every numeric parameter within its range through NumParam
'
' ERROR POLICY
'   Problems are logged as ERROR messages, not raised.
'
' DEPENDENCIES
'   NumParam, LogMsg; CORE_Util for GetParam, ToSerial and UTxt.
'
' NOTES
'   The holding periods are checked as whole numbers of business days.
'
' UPDATED
'   2026-10-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.LoadParams"

    Dim Problems    As Long 'Parameters rejected

'------------------------------------------------------------------------------
' REPORTING DATE, CURRENCY AND REGIME
'------------------------------------------------------------------------------
    'Read the reporting date; refuse a missing date or one before FIRST_ASOF
        pAsOf = ToSerial(GetParam(PRM_ASOF))
        If pAsOf <= 0# Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_ASOF, "Reporting date missing or not a date."
            Problems = Problems + 1
        ElseIf pAsOf < FIRST_ASOF Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_ASOF, "Reporting date before 2025-01-01 is not supported: " & _
                "the CRR as amended by Regulation (EU) 2024/1623 applies from that date."
            Problems = Problems + 1
        End If
    'Read the reporting currency; refuse a code that is not 3 letters
        pRepCcy = UTxt(GetParam(PRM_REPCCY))
        If VBA.Len(pRepCcy) <> 3 Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REPCCY, "Reporting currency must be a 3-letter code."
            Problems = Problems + 1
        End If
    'Read the regime; refuse anything but CRR
        pRegime = UTxt(GetParam(PRM_REGIME))
        If pRegime = RG_BCBS Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REGIME, "Regime BCBS is not supported: CRR is the only regime."
            Problems = Problems + 1
        ElseIf pRegime <> RG_CRR Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REGIME, "Regime must be CRR."
            Problems = Problems + 1
        End If

'------------------------------------------------------------------------------
' NUMBERS
'------------------------------------------------------------------------------
    'Read every numeric parameter within its range
        Problems = Problems + NumParam(PRM_DAYSYEAR, 360#, 366#, pDaysYear)
        Problems = Problems + NumParam(PRM_HCBASE, 1#, 250#, pHcBase, True)
        Problems = Problems + NumParam(PRM_HOLD_REPO, 1#, 250#, pHoldRepo, True)
        Problems = Problems + NumParam(PRM_HOLD_CMT, 1#, 250#, pHoldCMT, True)
        Problems = Problems + NumParam(PRM_HOLD_LARGE, 1#, 250#, pHoldLarge, True)
        Problems = Problems + NumParam(PRM_HFX, 0#, 1#, pHfx)
        Problems = Problems + NumParam(PRM_NONELIG, 0#, 1#, pNonElig)
        Problems = Problems + NumParam(PRM_NETW, 0#, 1#, pNetW)
        Problems = Problems + NumParam(PRM_GROSSW, 0#, 1#, pGrossW)
        Problems = Problems + NumParam(PRM_NTHRESH, 0#, 1#, pThreshold)
        Problems = Problems + NumParam(PRM_CAPRATIO, 0#, 1#, pCapRatio)
    'Pass only when no parameter was rejected
        LoadParams = (Problems = 0)

End Function

Private Function NumParam( _
    ByVal Code As String, _
    ByVal Lowest As Double, _
    ByVal Highest As Double, _
    ByRef Result As Double, _
    Optional ByVal WholeNumber As Boolean = False) As Long
'
'------------------------------------------------------------------------------
'                          READ ONE NUMERIC PARAMETER
'------------------------------------------------------------------------------
' PURPOSE
'   Read one numeric parameter and check it lies in [Lowest, Highest] and,
'   when WholeNumber is True, that it is a whole number (business days).
'
' WHY THIS EXISTS
'   Every numeric parameter is read and checked the same way.
'
' INPUTS
'   Code
'     The parameter code.
'
'   Lowest, Highest
'     The accepted range, bounds included.
'
'   Result
'     Receives the value read.
'
'   WholeNumber
'     True when the value must be a whole number of business days.
'
' RETURNS
'   0 when usable, with Result set; 1 with the problem logged otherwise.
'
' BEHAVIOR
'   Refuses a missing or non-numeric value; otherwise sets Result and
'   refuses a value outside the range or a fraction where whole days are
'   required.
'
' ERROR POLICY
'   Problems are logged as ERROR messages, not raised.
'
' DEPENDENCIES
'   LogMsg; CORE_Util for GetParam and IsNum.
'
' NOTES
'   Result is set from the cell even when it is then rejected.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.NumParam"

    Dim Content     As Variant  'Raw value

'------------------------------------------------------------------------------
' READ AND CHECK
'------------------------------------------------------------------------------
    'Refuse a missing or non-numeric value
        Content = GetParam(Code)
        If Not IsNum(Content) Then
            LogMsg SEV_ERROR, SH_PARAMS, Code, "Parameter missing or not a number."
            NumParam = 1
            Exit Function
        End If
    'Refuse a value outside the range, or a fraction where whole days are required
        Result = VBA.CDbl(Content)
        If Result < Lowest Or Result > Highest Then
            LogMsg SEV_ERROR, SH_PARAMS, Code, "Parameter " & Result & " outside [" & Lowest & ", " & Highest & "]."
            NumParam = 1
        ElseIf WholeNumber And Result <> Int(Result) Then
            LogMsg SEV_ERROR, SH_PARAMS, Code, "Parameter " & Result & " is not a whole number of business days."
            NumParam = 1
        End If

End Function

Private Function LoadFXTable() As Boolean

'
'------------------------------------------------------------------------------
'                              READ THE FX TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Read the FX table below its header down to the first blank currency:
'   units of reporting currency per unit of each currency. A currency cell
'   holding a worksheet error value is reported, not read as blank.
'
' WHY THIS EXISTS
'   Every leg is converted to the reporting currency with this table.
'
' INPUTS
'   None
'
' RETURNS
'   True when every row is usable and the reporting currency is listed at
'   1; otherwise False with each problem logged, and the run stops.
'
' BEHAVIOR
'   - Reads each row below the header until the currency cell is blank
'   - Refuses an error value, a code that is not 3 letters, a currency
'     listed twice, and a rate that is missing, not a number or not above 0
'   - Stores each usable row and indexes it by currency
'   - Checks that the reporting currency is listed with the rate 1
'
' ERROR POLICY
'   Problems are logged as ERROR messages, not raised.
'
' DEPENDENCIES
'   LogMsg; CORE_Util for NamedCell, IsBlankCell, UTxt, IsNum, KeyIndex
'   and KeyAdd.
'
' NOTES
'   The rate is read from the third column of the table.
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.LoadFXTable"

    Dim Header      As Range    'Header cell FX_Ccy
    Dim Ccy         As String   'Currency of the row
    Dim Rate        As Variant  'Rate of the row
    Dim Problems    As Long     'Rows rejected
    Dim RowOffset   As Long     'Row offset below the header
    Dim Position    As Long     'Index of the reporting currency

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Locate the table and allocate the lists
        Set Header = NamedCell(NM_FX_TABLE)
        ReDim mFxCcy(1 To 16)
        ReDim mFxRate(1 To 16)
    'Read each row until the currency cell is blank
        RowOffset = 1
        Do While Not IsBlankCell(Header.Offset(RowOffset, 0).Value) Or IsError(Header.Offset(RowOffset, 0).Value)
            Ccy = UTxt(Header.Offset(RowOffset, 0).Value)
            Rate = Header.Offset(RowOffset, 2).Value2
            If IsError(Header.Offset(RowOffset, 0).Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, vbNullString, "FX currency holds a worksheet error value.", _
                    Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf VBA.Len(Ccy) <> 3 Then
                LogMsg SEV_ERROR, SH_PARAMS, Ccy, "FX currency must be a 3-letter code.", Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf KeyIndex(mFxIndex, Ccy) > 0 Then
                LogMsg SEV_ERROR, SH_PARAMS, Ccy, "Currency listed twice in the FX table.", Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf Not IsNum(Rate) Then
                LogMsg SEV_ERROR, SH_PARAMS, Ccy, "FX rate missing or not a number.", Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf VBA.CDbl(Rate) <= 0# Then
                LogMsg SEV_ERROR, SH_PARAMS, Ccy, "FX rate must be above 0.", Header.Row + RowOffset
                Problems = Problems + 1
            Else
                mFxCount = mFxCount + 1
                If mFxCount > UBound(mFxCcy) Then
                    ReDim Preserve mFxCcy(1 To mFxCount * 2)
                    ReDim Preserve mFxRate(1 To mFxCount * 2)
                End If
                mFxCcy(mFxCount) = Ccy
                mFxRate(mFxCount) = VBA.CDbl(Rate)
                KeyAdd mFxIndex, Ccy, mFxCount
            End If
            RowOffset = RowOffset + 1
        Loop

'------------------------------------------------------------------------------
' CHECK THE REPORTING CURRENCY
'------------------------------------------------------------------------------
    'Require the reporting currency in the table with the rate 1
        Position = KeyIndex(mFxIndex, pRepCcy)
        If Position = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, pRepCcy, "Reporting currency not in the FX table."
            Problems = Problems + 1
        ElseIf mFxRate(Position) <> 1# Then
            LogMsg SEV_ERROR, SH_PARAMS, pRepCcy, "Reporting currency must have the FX rate 1."
            Problems = Problems + 1
        End If
    'Pass only when no row was rejected
        LoadFXTable = (Problems = 0)

End Function

Private Function LoadHaircutTable() As Boolean

'
'------------------------------------------------------------------------------
'                            READ THE HAIRCUT TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Read the supervisory haircut table below its header down to the first
'   blank key: key, H10 and the regimes each row applies to. A key that is
'   not in the table makes the asset not eligible. A key cell holding a
'   worksheet error value is reported, not read as blank.
'
' WHY THIS EXISTS
'   The haircut of every leg is looked up by key in this table.
'
' INPUTS
'   None
'
' RETURNS
'   True when every row is usable; otherwise False and the run stops.
'
' BEHAVIOR
'   - Reads each row below the header until the key cell is blank
'   - Refuses an error value, a key listed twice, an H10 that is missing,
'     not a number or outside [0, 1], and regimes other than BOTH or CRR
'   - Stores each usable row and indexes it by key
'   - Refuses an empty table
'
' ERROR POLICY
'   Problems are logged as ERROR messages, not raised.
'
' DEPENDENCIES
'   LogMsg; CORE_Util for NamedCell, IsBlankCell, UTxt, IsNum, KeyIndex
'   and KeyAdd; CORE_Config for the column constants.
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.49; CRR Art. 224 Tables 1 and 3.
'
' UPDATED
'   2026-10-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.LoadHaircutTable"

    Dim Header      As Range    'Header cell HC_Key
    Dim Key         As String   'Key of the row
    Dim Value       As Variant  'H10 of the row
    Dim Regimes     As String   'Regimes of the row
    Dim Problems    As Long     'Rows rejected
    Dim RowOffset   As Long     'Row offset below the header

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Locate the table and allocate the lists
        Set Header = NamedCell(NM_HC_TABLE)
        ReDim mHcKeys(1 To 64)
        ReDim mHcValues(1 To 64)
        ReDim mHcRegimes(1 To 64)
    'Read each row until the key cell is blank
        RowOffset = 1
        Do While Not IsBlankCell(Header.Offset(RowOffset, HC_KEY - 1).Value) Or _
            IsError(Header.Offset(RowOffset, HC_KEY - 1).Value)
            Key = UTxt(Header.Offset(RowOffset, HC_KEY - 1).Value)
            Value = Header.Offset(RowOffset, HC_H10 - 1).Value2
            Regimes = UTxt(Header.Offset(RowOffset, HC_REGIMES - 1).Value)
            If IsError(Header.Offset(RowOffset, HC_KEY - 1).Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, vbNullString, "Haircut key holds a worksheet error value.", _
                    Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf KeyIndex(mHcIndex, Key) > 0 Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Haircut key listed twice.", Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf Not IsNum(Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Haircut missing or not a number.", Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf VBA.CDbl(Value) < 0# Or VBA.CDbl(Value) > 1# Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Haircut must be between 0 and 1.", Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf Regimes <> RG_BOTH And Regimes <> RG_CRR Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Regimes must be BOTH or CRR: CRR is the only regime.", _
                    Header.Row + RowOffset
                Problems = Problems + 1
            Else
                mHcCount = mHcCount + 1
                If mHcCount > UBound(mHcKeys) Then
                    ReDim Preserve mHcKeys(1 To mHcCount * 2)
                    ReDim Preserve mHcValues(1 To mHcCount * 2)
                    ReDim Preserve mHcRegimes(1 To mHcCount * 2)
                End If
                mHcKeys(mHcCount) = Key
                mHcValues(mHcCount) = VBA.CDbl(Value)
                mHcRegimes(mHcCount) = Regimes
                KeyAdd mHcIndex, Key, mHcCount
            End If
            RowOffset = RowOffset + 1
        Loop
    'Refuse an empty table
        If mHcCount = 0 And Problems = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, NM_HC_TABLE, "Haircut table is empty."
            Problems = 1
        End If
    'Pass only when no row was rejected
        LoadHaircutTable = (Problems = 0)

End Function

Private Function LoadRiskWeightTable() As Boolean

'
'------------------------------------------------------------------------------
'                          READ THE RISK-WEIGHT TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Read the risk-weight table below its header down to the first blank
'   key: key, risk weight and the regimes each row applies to. A key cell
'   holding a worksheet error value is reported, not read as blank.
'
' WHY THIS EXISTS
'   The risk weight of every netting set is looked up by key in this table.
'
' INPUTS
'   None
'
' RETURNS
'   True when every row is usable; otherwise False and the run stops.
'
' BEHAVIOR
'   - Reads each row below the header until the key cell is blank
'   - Refuses an error value, a key listed twice, a risk weight that is
'     missing, not a number or outside [0, 12.5], and regimes other than
'     BOTH or CRR
'   - Stores each usable row and indexes it by key
'   - Refuses an empty table
'
' ERROR POLICY
'   Problems are logged as ERROR messages, not raised.
'
' DEPENDENCIES
'   LogMsg; CORE_Util for NamedCell, IsBlankCell, UTxt, IsNum, KeyIndex
'   and KeyAdd; CORE_Config for the column constants.
'
' NOTES
'   None
'
' REFERENCE
'   CRE20.7, 20.18-20.21, CRE54.7; CRR Art. 114, 120-122, 306(1)(a).
'
' UPDATED
'   2026-10-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.LoadRiskWeightTable"

    Dim Header      As Range    'Header cell RW_Key
    Dim Key         As String   'Key of the row
    Dim Value       As Variant  'Risk weight of the row
    Dim Regimes     As String   'Regimes of the row
    Dim Problems    As Long     'Rows rejected
    Dim RowOffset   As Long     'Row offset below the header

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Locate the table and allocate the lists
        Set Header = NamedCell(NM_RW_TABLE)
        ReDim mRwKeys(1 To 64)
        ReDim mRwValues(1 To 64)
        ReDim mRwRegimes(1 To 64)
    'Read each row until the key cell is blank
        RowOffset = 1
        Do While Not IsBlankCell(Header.Offset(RowOffset, RW_KEY - 1).Value) Or _
            IsError(Header.Offset(RowOffset, RW_KEY - 1).Value)
            Key = UTxt(Header.Offset(RowOffset, RW_KEY - 1).Value)
            Value = Header.Offset(RowOffset, RW_WEIGHT - 1).Value2
            Regimes = UTxt(Header.Offset(RowOffset, RW_REGIMES - 1).Value)
            If IsError(Header.Offset(RowOffset, RW_KEY - 1).Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, vbNullString, "Risk-weight key holds a worksheet error value.", _
                    Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf KeyIndex(mRwIndex, Key) > 0 Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Risk-weight key listed twice.", Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf Not IsNum(Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Risk weight missing or not a number.", Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf VBA.CDbl(Value) < 0# Or VBA.CDbl(Value) > 12.5 Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Risk weight must be between 0 and 12.5.", Header.Row + RowOffset
                Problems = Problems + 1
            ElseIf Regimes <> RG_BOTH And Regimes <> RG_CRR Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Regimes must be BOTH or CRR: CRR is the only regime.", _
                    Header.Row + RowOffset
                Problems = Problems + 1
            Else
                mRwCount = mRwCount + 1
                If mRwCount > UBound(mRwKeys) Then
                    ReDim Preserve mRwKeys(1 To mRwCount * 2)
                    ReDim Preserve mRwValues(1 To mRwCount * 2)
                    ReDim Preserve mRwRegimes(1 To mRwCount * 2)
                End If
                mRwKeys(mRwCount) = Key
                mRwValues(mRwCount) = VBA.CDbl(Value)
                mRwRegimes(mRwCount) = Regimes
                KeyAdd mRwIndex, Key, mRwCount
            End If
            RowOffset = RowOffset + 1
        Loop
    'Refuse an empty table
        If mRwCount = 0 And Problems = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, NM_RW_TABLE, "Risk-weight table is empty."
            Problems = 1
        End If
    'Pass only when no row was rejected
        LoadRiskWeightTable = (Problems = 0)

End Function

Private Function FXRate( _
    ByVal Ccy As String) As Double
'
'------------------------------------------------------------------------------
'                              LOOK UP AN FX RATE
'------------------------------------------------------------------------------
' PURPOSE
'   Units of reporting currency per unit of a currency; 0 when the currency
'   is not in the FX table.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.FXRate"

    Dim Position    As Long 'Row in the FX table

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'Return the rate of a listed currency; 0 otherwise
        Position = KeyIndex(mFxIndex, Ccy)
        If Position > 0 Then
            FXRate = mFxRate(Position)
        End If

End Function

Private Function TableRow( _
    ByVal Lookup As Collection, _
    ByRef Regimes() As String, _
    ByVal Key As String, _
    ByVal RegimeCode As String) As Long
'
'------------------------------------------------------------------------------
'                        FIND A TABLE ROW FOR A REGIME
'------------------------------------------------------------------------------
' PURPOSE
'   Row of a haircut or risk-weight table key that applies to a regime.
'
' RETURNS
'   The row; 0 when the key is missing or its row is limited to the other
'   regime.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.TableRow"

    Dim Position    As Long 'Row of the key

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'Return 0 for a blank key
        If VBA.Len(Key) = 0 Then
            Exit Function
        End If
    'Return the row when it applies to both regimes or to this one
        Position = KeyIndex(Lookup, Key)
        If Position > 0 Then
            If Regimes(Position) = RG_BOTH Or Regimes(Position) = RegimeCode Then
                TableRow = Position
            End If
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                                RUN CLASSIFICATION
'
'------------------------------------------------------------------------------
'

Private Sub ClassifyRun()

'
'------------------------------------------------------------------------------
'                               CLASSIFY THE RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Classify the run as REGULATORY, when every parameter compared (all but
'   the reporting date, reporting currency, default regime and DaysPerYear),
'   the haircut table and the risk-weight table equal the regulatory
'   parameter set REF_SET_ID and no risk weight is overridden (the overrides
'   are read with the netting sets, which are loaded first), or WHAT-IF,
'   listing each difference on Checks as INFO.
'
' WHY THIS EXISTS
'   A run with other parameters than the regulatory set must not be taken
'   for a regulatory result.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Compares each parameter with its REF_ value
'   - Compares the haircut and risk-weight tables with the reference rows
'   - Counts each risk weight override as a difference
'   - Sets the run class from the number of differences
'
' ERROR POLICY
'   None raised by the procedure itself.
'
' DEPENDENCIES
'   CompareParam, CompareTable, Deviation; CORE_Config for the REF_
'   constants and the run classes.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Engine.ClassifyRun"

    Dim NettingSetIndex As Long 'Netting set

'------------------------------------------------------------------------------
' COMPARE PARAMETERS
'------------------------------------------------------------------------------
    'Compare each parameter with the regulatory set
        CompareParam PRM_HCBASE, pHcBase, REF_HCBASE
        CompareParam PRM_HOLD_REPO, pHoldRepo, REF_HOLD_REPO
        CompareParam PRM_HOLD_CMT, pHoldCMT, REF_HOLD_CMT
        CompareParam PRM_HOLD_LARGE, pHoldLarge, REF_HOLD_LARGE
        CompareParam PRM_HFX, pHfx, REF_HFX
        CompareParam PRM_NONELIG, pNonElig, REF_NONELIG
        CompareParam PRM_NETW, pNetW, REF_NETW
        CompareParam PRM_GROSSW, pGrossW, REF_GROSSW
        CompareParam PRM_NTHRESH, pThreshold, REF_NTHRESH
        CompareParam PRM_CAPRATIO, pCapRatio, REF_CAPRATIO

'------------------------------------------------------------------------------
' COMPARE TABLES AND OVERRIDES
'------------------------------------------------------------------------------
    'Compare the haircut and risk-weight tables with the regulatory set
        CompareTable NM_HC_TABLE, REF_HAIRCUTS, mHcIndex, mHcKeys, mHcValues, mHcRegimes, mHcCount
        CompareTable NM_RW_TABLE, REF_RISKWEIGHTS, mRwIndex, mRwKeys, mRwValues, mRwRegimes, mRwCount
    'Count every risk weight override as a difference
        For NettingSetIndex = 1 To mNsCount
            If mNS(NettingSetIndex).HasRWOverride Then
                Deviation SH_NS, mNS(NettingSetIndex).ID, "Risk weight override " & _
                    mNS(NettingSetIndex).RWOverride & "."
            End If
        Next NettingSetIndex
    'Classify the run from the differences found
        If mDeviations = 0 Then
            mRunClass = RUN_REGULATORY
        Else
            mRunClass = RUN_WHATIF
        End If

End Sub

Private Sub CompareParam( _
    ByVal Code As String, _
    ByVal Actual As Double, _
    ByVal Reference As Double)
'
'------------------------------------------------------------------------------
'                            COMPARE ONE PARAMETER
'------------------------------------------------------------------------------
' PURPOSE
'   Record a deviation when a parameter differs from the regulatory set.
'
' NOTES
'   Differences up to 1E-9 are treated as equal.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Record a difference beyond the tolerance
        If Abs(Actual - Reference) > 0.000000001 Then
            Deviation SH_PARAMS, Code, "Parameter " & Actual & " differs from " & Reference & " in " & _
                REF_SET_ID & "."
        End If

End Sub

Private Sub CompareTable( _
    ByVal TableName As String, _
    ByVal Reference As String, _
    ByVal Lookup As Collection, _
    ByRef Keys() As String, _
    ByRef Values() As Double, _
    ByRef Regimes() As String, _
    ByVal RowCount As Long)
'
'------------------------------------------------------------------------------
'                            COMPARE A PARAMS TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Record a deviation for every row of a Params table that differs from
'   the regulatory set: a value that differs, a row limited to one regime,
'   a row added, or a reference row missing.
'
' WHY THIS EXISTS
'   The haircut and risk-weight tables are compared the same way.
'
' INPUTS
'   TableName
'     The workbook name of the table, for the message.
'
'   Reference
'     The set's rows as KEY=value pairs separated by ";".
'
'   Lookup, Keys, Values, Regimes, RowCount
'     The table as loaded.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Checks each reference row against the table: missing, different
'     value, or limited to one regime
'   - Lists each table row that is not in the reference set
'
' ERROR POLICY
'   None raised by the procedure itself.
'
' DEPENDENCIES
'   Deviation; CORE_Util for KeyIndex and KeyAdd.
'
' NOTES
'   Values are compared with a tolerance of 1E-9.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.CompareTable"

    Dim Pairs()     As String       'Reference rows
    Dim Parts()     As String       'Key and value of one reference row
    Dim Seen        As Collection   'Reference keys
    Dim Position    As Long         'Row of a key in the table
    Dim RowIndex    As Long         'Reference row or table row

'------------------------------------------------------------------------------
' REFERENCE ROWS AGAINST THE TABLE
'------------------------------------------------------------------------------
    'Split the reference rows
        Set Seen = New Collection
        Pairs = Split(Reference, ";")
    'Record each reference row that is missing, differs or is limited to one regime
        For RowIndex = 0 To UBound(Pairs)
            Parts = Split(Pairs(RowIndex), "=")
            KeyAdd Seen, Parts(0), RowIndex + 1
            Position = KeyIndex(Lookup, Parts(0))
            If Position = 0 Then
                Deviation SH_PARAMS, Parts(0), TableName & ": reference row missing."
            ElseIf Abs(Values(Position) - Val(Parts(1))) > 0.000000001 Then
                Deviation SH_PARAMS, Parts(0), TableName & ": " & Values(Position) & " differs from " & _
                    Parts(1) & " in " & REF_SET_ID & "."
            ElseIf Regimes(Position) <> RG_BOTH Then
                Deviation SH_PARAMS, Parts(0), TableName & ": row limited to " & Regimes(Position) & "."
            End If
        Next RowIndex

'------------------------------------------------------------------------------
' ROWS ADDED TO THE TABLE
'------------------------------------------------------------------------------
    'Record each table row that is not in the reference set
        For RowIndex = 1 To RowCount
            If KeyIndex(Seen, Keys(RowIndex)) = 0 Then
                Deviation SH_PARAMS, Keys(RowIndex), TableName & ": row not in " & REF_SET_ID & "."
            End If
        Next RowIndex

End Sub

Private Sub Deviation( _
    ByVal SheetName As String, _
    ByVal Reference As String, _
    ByVal Message As String)
'
'------------------------------------------------------------------------------
'                       RECORD A DIFFERENCE FROM THE SET
'------------------------------------------------------------------------------
' PURPOSE
'   Count one difference from the regulatory set and list it as INFO.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RECORD
'------------------------------------------------------------------------------
    'Count the difference and log it
        mDeviations = mDeviations + 1
        LogMsg SEV_INFO, SheetName, Reference, "WHAT-IF: " & Message

End Sub

'
'------------------------------------------------------------------------------
'
'                                 NETTING SETS
'
'------------------------------------------------------------------------------
'

Private Function LoadNettingSets() As Boolean

'
'------------------------------------------------------------------------------
'                            READ THE NETTING SETS
'------------------------------------------------------------------------------
' PURPOSE
'   Read every netting set row. A row whose inputs are not usable is kept
'   with its errors counted, so that its transactions are reported against
'   it and its exposure is withheld; a duplicate ID makes the first row
'   INVALID and the duplicate is not read.
'
' WHY THIS EXISTS
'   The legs are attached to their netting sets by ID, so the netting sets
'   are read and indexed before any leg.
'
' INPUTS
'   None
'
' RETURNS
'   True when at least one netting set was read; otherwise False and the
'   run stops.
'
' BEHAVIOR
'   - Reads the NettingSets table in one block; stops when it is empty
'   - Skips empty rows and logs a row without an ID
'   - Invalidates the first row of a duplicated ID
'   - Indexes and reads every other row through ReadNettingSet
'   - Stops when no netting set was read
'
' ERROR POLICY
'   Problems are logged as ERROR messages, not raised.
'
' DEPENDENCIES
'   ReadNettingSet, NsError, LogMsg; CORE_Util for the sheet, block and
'   key helpers.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.LoadNettingSets"

    Dim Sheet       As Worksheet    'NettingSets sheet
    Dim FirstRow    As Long         'First data row
    Dim LastRow     As Long         'Last row with data
    Dim Block       As Variant      'Input block
    Dim ID          As String       'Netting-set ID of the row
    Dim Existing    As Long         'Index of an earlier row with the same ID
    Dim BlockRow    As Long         'Row in Block

'------------------------------------------------------------------------------
' READ THE BLOCK
'------------------------------------------------------------------------------
    'Find the rows with data; stop when there are none
        Set Sheet = GetSheet(SH_NS)
        FirstRow = TableFirstRow(SH_NS)
        LastRow = LastDataRowAny(Sheet, FirstRow, NS_NCOLS, 0)
        If LastRow < FirstRow Then
            LogMsg SEV_ERROR, SH_NS, vbNullString, "No netting sets entered."
            Exit Function
        End If
    'Read the table in one block and allocate the netting sets
        Block = TableRange(Sheet, FirstRow, 1, LastRow, NS_NCOLS).Value2
        ReDim mNS(1 To LastRow - FirstRow + 1)

'------------------------------------------------------------------------------
' READ EACH ROW
'------------------------------------------------------------------------------
    'Read each row with data, refusing a missing or duplicated ID
        For BlockRow = 1 To UBound(Block, 1)
            If RowHasData(Block, BlockRow, NS_NCOLS, 0) Then
                ID = SafeStr(Block(BlockRow, NS_ID))
                Existing = KeyIndex(mNsIndex, ID)
                If VBA.Len(ID) = 0 Then
                    LogMsg SEV_ERROR, SH_NS, vbNullString, "Netting-set ID missing: row not read.", _
                        FirstRow + BlockRow - 1
                ElseIf Existing > 0 Then
                    NsError Existing, "Netting-set ID also entered on row " & (FirstRow + BlockRow - 1)
                Else
                    mNsCount = mNsCount + 1
                    KeyAdd mNsIndex, ID, mNsCount
                    ReadNettingSet mNsCount, Block, BlockRow, FirstRow + BlockRow - 1
                End If
            End If
        Next BlockRow
    'Stop when no netting set could be read
        If mNsCount = 0 Then
            LogMsg SEV_ERROR, SH_NS, vbNullString, "No usable netting set."
            Exit Function
        End If
    'Report success
        LoadNettingSets = True

End Function

Private Sub ReadNettingSet( _
    ByVal NsIndex As Long, _
    ByRef Block As Variant, _
    ByVal r As Long, _
    ByVal SheetRow As Long)
'
'------------------------------------------------------------------------------
'                             READ ONE NETTING SET
'------------------------------------------------------------------------------
' PURPOSE
'   Read and check the fields of one netting set row.
'
' WHY THIS EXISTS
'   Every field of a netting set is checked once, here; an unusable field
'   makes the netting set INVALID instead of stopping the run.
'
' INPUTS
'   NsIndex
'     Index in mNS.
'
'   Block, r
'     The input block and the row in it.
'
'   SheetRow
'     The sheet row, for the messages.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Reads the ID, counterparty, netting agreement, settlement currency
'     (default: the reporting currency) and regime (default: Params)
'   - Checks the counterparty class, CQS and SCRA grade
'   - Reads the remargining frequency, large or illiquid flag, disputes,
'     holding period override and zero-haircut flag
'   - Reads the risk weight override
'   - Reports cells holding a worksheet error value
'
' ERROR POLICY
'   Problems are logged through NsError; a SCRA grade given where it does
'   not apply is a WARNING.
'
' DEPENDENCIES
'   NsError, LogMsg, ErrorColumns; CORE_Util for SafeStr, UTxt, TryBool,
'   TryDbl, IsNum, IsBlankCell and KeyIndex; CORE_Config for the columns
'   and codes.
'
' NOTES
'   r is the row in Block, not the sheet row.
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.ReadNettingSet"

    Dim Flag        As Boolean  'A Y/N field read
    Dim Number      As Double   'A numeric field read
    Dim Found       As String   'Headers of the cells holding an error value

'------------------------------------------------------------------------------
' IDENTITY AND AGREEMENT
'------------------------------------------------------------------------------
    'Work on the netting set being read
        With mNS(NsIndex)
        'Store the identity and the sheet row
            .ID = SafeStr(Block(r, NS_ID))
            .Row = SheetRow
            .Counterparty = SafeStr(Block(r, NS_CPTY))
        'Read the netting agreement flag, which must be Y or N
            If TryBool(Block(r, NS_MNA), False, Flag) And Not IsBlankCell(Block(r, NS_MNA)) Then
                .MNA = Flag
            Else
                NsError NsIndex, "Eligible master netting agreement must be Y or N"
            End If
        'Read the settlement currency; the reporting currency by default
            .SettleCcy = UTxt(Block(r, NS_SETTLE))
            If VBA.Len(.SettleCcy) = 0 Then
                .SettleCcy = pRepCcy
            ElseIf KeyIndex(mFxIndex, .SettleCcy) = 0 Then
                NsError NsIndex, "Settlement currency " & .SettleCcy & " not in the FX table"
            End If
        'Read the regime; the Params regime by default and when it is refused
            .Regime = UTxt(Block(r, NS_REGIME))
            If VBA.Len(.Regime) = 0 Then
                .Regime = pRegime
            ElseIf .Regime = RG_BCBS Then
                NsError NsIndex, "Regime BCBS is not supported: CRR is the only regime"
                .Regime = pRegime
            ElseIf .Regime <> RG_CRR Then
                NsError NsIndex, "Regime override must be blank or CRR"
                .Regime = pRegime
            End If

'------------------------------------------------------------------------------
' COUNTERPARTY
'------------------------------------------------------------------------------
        'Accept SOV, INST and CORP; refuse QCCP and anything else
            .CptyClass = UTxt(Block(r, NS_CLASS))
            Select Case .CptyClass
                Case CC_SOV, CC_INST, CC_CORP
                Case CC_QCCP
                    NsError NsIndex, "Counterparty class QCCP is not supported: exposures to CCPs are out of scope"
                Case Else
                    NsError NsIndex, "Counterparty class must be SOV, INST or CORP"
            End Select
        'Read the CQS, which must be blank or a whole number from 1 to 6
            If Not IsBlankCell(Block(r, NS_CQS)) Then
                If IsNum(Block(r, NS_CQS)) Then
                    Number = VBA.CDbl(Block(r, NS_CQS))
                Else
                    Number = 0#
                End If
                If Number < 1# Or Number > 6# Or Number <> Int(Number) Then
                    NsError NsIndex, "Counterparty CQS must be blank or a whole number from 1 to 6"
                Else
                    .CQS = VBA.CLng(Number)
                End If
            End If
        'Require a SCRA grade for an unrated institution; warn when it does not apply
            .SCRA = UTxt(Block(r, NS_SCRA))
            If .CptyClass = CC_INST And .CQS = 0 Then
                If VBA.InStr(1, " " & SCRA_GRADES & " ", " " & .SCRA & " ", vbBinaryCompare) = 0 Or _
                    VBA.Len(.SCRA) = 0 Then
                    NsError NsIndex, "Unrated institution: SCRA grade must be A, A_ENH, B or C"
                End If
            ElseIf VBA.Len(.SCRA) > 0 Then
                LogMsg SEV_WARN, SH_NS, .ID, "SCRA grade ignored: it applies to unrated institutions only.", .Row
            End If

'------------------------------------------------------------------------------
' HOLDING PERIOD DRIVERS
'------------------------------------------------------------------------------
        'Read the remargining frequency: whole business days, 1 or more
            .RemarginBD = 1#
            If TryDbl(Block(r, NS_FREQ), 1#, Number) Then
                If Number < 1# Then
                    NsError NsIndex, "Remargin frequency must be 1 business day or more"
                ElseIf Number <> Int(Number) Then
                    NsError NsIndex, "Remargin frequency must be a whole number of business days"
                Else
                    .RemarginBD = Number
                End If
            Else
                NsError NsIndex, "Remargin frequency must be a number"
            End If
        'Read the large-or-illiquid flag
            If TryBool(Block(r, NS_LARGE), False, Flag) Then
                .LargeOrIlliquid = Flag
            Else
                NsError NsIndex, ">5,000 trades or illiquid collateral must be Y or N"
            End If
        'Read the margin disputes flag
            If TryBool(Block(r, NS_DISPUTE), False, Flag) Then
                .Disputes = Flag
            Else
                NsError NsIndex, "Margin disputes must be Y or N"
            End If
        'Read the holding period override: whole business days, 0 or more
            If TryDbl(Block(r, NS_HOLDOVR), 0#, Number) Then
                If Number < 0# Then
                    NsError NsIndex, "Holding period override must be 0 or more"
                ElseIf Number <> Int(Number) Then
                    NsError NsIndex, "Holding period override must be a whole number of business days"
                Else
                    .HoldingOverride = Number
                End If
            Else
                NsError NsIndex, "Holding period override must be a number"
            End If
        'Read the zero-haircut flag
            If TryBool(Block(r, NS_ZEROHC), False, Flag) Then
                .ZeroHaircut = Flag
            Else
                NsError NsIndex, "0% haircut conditions met must be Y or N"
            End If

'------------------------------------------------------------------------------
' RISK WEIGHT OVERRIDE
'------------------------------------------------------------------------------
        'Read a risk weight override between 0 and 12.5 when one is given
            If Not IsBlankCell(Block(r, NS_RWOVR)) Then
                If TryDbl(Block(r, NS_RWOVR), 0#, Number) Then
                    If Number < 0# Or Number > 12.5 Then
                        NsError NsIndex, "Risk weight override must be between 0 and 12.5"
                    Else
                        .HasRWOverride = True
                        .RWOverride = Number
                    End If
                Else
                    NsError NsIndex, "Risk weight override must be a number"
                End If
            End If
        End With

'------------------------------------------------------------------------------
' WORKSHEET ERROR VALUES
'------------------------------------------------------------------------------
    'The text fields, the CQS and the risk weight override would read an
    'error value as blank and take their default; the Y/N and numeric
    'fields already reject it, so for them this is a second message
        Found = ErrorColumns(Block, r, NS_HEADERS, 0)
        If VBA.Len(Found) > 0 Then
            NsError NsIndex, "Worksheet error value in: " & Found
        End If

End Sub

'
'------------------------------------------------------------------------------
'
'                                     LEGS
'
'------------------------------------------------------------------------------
'

Private Sub LoadLegs()

'
'------------------------------------------------------------------------------
'                                READ THE LEGS
'------------------------------------------------------------------------------
' PURPOSE
'   Read every row of the Transactions sheet as a leg, check its fields,
'   value it in the reporting currency and find its supervisory haircut.
'
' WHY THIS EXISTS
'   Each row of the Transactions sheet is one leg; the transactions are
'   built from the legs afterwards.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Reads the Transactions table in one block; warns and returns when it
'     is empty
'   - Reads every row with data through ReadLeg
'
' ERROR POLICY
'   Problems are logged, not raised; an empty table is a WARNING.
'
' DEPENDENCIES
'   ReadLeg, LogMsg; CORE_Util for the sheet and block helpers.
'
' NOTES
'   The comment column does not count as data.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.LoadLegs"

    Dim Sheet       As Worksheet    'Transactions sheet
    Dim FirstRow    As Long         'First data row
    Dim LastRow     As Long         'Last row with data
    Dim Block       As Variant      'Input block
    Dim BlockRow    As Long         'Row in Block

'------------------------------------------------------------------------------
' READ THE BLOCK
'------------------------------------------------------------------------------
    'Find the rows with data; warn and return when there are none
        Set Sheet = GetSheet(SH_TX)
        FirstRow = TableFirstRow(SH_TX)
        LastRow = LastDataRowAny(Sheet, FirstRow, TX_NCOLS, TX_COMMENT)
        If LastRow < FirstRow Then
            LogMsg SEV_WARN, SH_TX, vbNullString, "No transactions entered."
            Exit Sub
        End If
    'Read the table in one block and allocate the legs
        Block = TableRange(Sheet, FirstRow, 1, LastRow, TX_NCOLS).Value2
        ReDim mLegs(1 To LastRow - FirstRow + 1)

'------------------------------------------------------------------------------
' READ EACH ROW
'------------------------------------------------------------------------------
    'Read each row with data as a leg
        For BlockRow = 1 To UBound(Block, 1)
            If RowHasData(Block, BlockRow, TX_NCOLS, TX_COMMENT) Then
                mLegCount = mLegCount + 1
                ReadLeg mLegCount, Block, BlockRow, FirstRow + BlockRow - 1
            End If
        Next BlockRow

End Sub

Private Sub ReadLeg( _
    ByVal LegIndex As Long, _
    ByRef Block As Variant, _
    ByVal r As Long, _
    ByVal SheetRow As Long)
'
'------------------------------------------------------------------------------
'                                 READ ONE LEG
'------------------------------------------------------------------------------
' PURPOSE
'   Read and check the fields of one leg; every input error is an ERROR
'   that rejects the leg and, through BuildTransactions, its transaction.
'   A valid leg is then keyed to its haircut (FindHaircut).
'
' WHY THIS EXISTS
'   Every field of a leg is checked once, here, before the legs are
'   grouped into transactions.
'
' INPUTS
'   LegIndex
'     Index in mLegs.
'
'   Block, r
'     The input block and the row in it.
'
'   SheetRow
'     The sheet row, for the messages.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Reads the leg, transaction and netting-set IDs; refuses a missing or
'     repeated leg ID, a missing transaction ID and an unknown netting set
'   - Checks the transaction type, start and maturity dates and side
'   - Reads the asset fields through ReadLegAsset
'   - Reads the market value and currency and values the leg in the
'     reporting currency
'   - Builds the issue key of cash and the signature of the issue
'   - Reports cells holding a worksheet error value
'   - Finds the haircut of a valid leg
'
' ERROR POLICY
'   Problems are logged through LegError, not raised.
'
' DEPENDENCIES
'   LegError, ReadLegAsset, ErrorColumns, FXRate, FindHaircut; CORE_Util
'   for SafeStr, UTxt, ToSerial, IsNum, IsBlankCell, KeyIndex and KeyAdd.
'
' NOTES
'   A missing leg or transaction ID is replaced by "(row n)" for the
'   messages.
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.ReadLeg"

    Dim Found       As String   'Headers of the cells holding an error value

'------------------------------------------------------------------------------
' IDENTITY
'------------------------------------------------------------------------------
    'Work on the leg being read
        With mLegs(LegIndex)
        'Start the leg as valid, with its IDs and no residual maturity
            .Row = SheetRow
            .Valid = True
            .LegID = SafeStr(Block(r, TX_LEG))
            .TxID = SafeStr(Block(r, TX_ID))
            .TxKey = VBA.UCase$(.TxID)
            .ResidualYears = -1#
        'Refuse a missing or repeated leg ID; index the leg otherwise
            If VBA.Len(.LegID) = 0 Then
                .LegID = "(row " & SheetRow & ")"
                LegError LegIndex, "Leg ID missing"
            ElseIf KeyIndex(mLegIndex, .LegID) > 0 Then
                LegError LegIndex, "Leg ID entered twice"
            Else
                KeyAdd mLegIndex, .LegID, LegIndex
            End If
        'Refuse a missing transaction ID
            If VBA.Len(.TxID) = 0 Then
                .TxID = "(row " & SheetRow & ")"
                .TxKey = VBA.UCase$(.TxID)
                LegError LegIndex, "Transaction ID missing"
            End If
        'Attach the leg to its netting set
            .NsIndex = KeyIndex(mNsIndex, SafeStr(Block(r, TX_NS)))
            If .NsIndex = 0 Then
                LegError LegIndex, "Netting set '" & SafeStr(Block(r, TX_NS)) & "' not on NettingSets"
            End If

'------------------------------------------------------------------------------
' TRANSACTION FIELDS
'------------------------------------------------------------------------------
        'Refuse an unknown transaction type
            .TxType = UTxt(Block(r, TX_TYPE))
            Select Case .TxType
                Case TT_REPO, TT_REVREPO, TT_SECLEND, TT_SECBORROW, TT_MARGINLEND
                Case Else
                    LegError LegIndex, "Transaction type '" & .TxType & "' is not REPO, REVREPO, SECLEND, " & _
                        "SECBORROW or MARGINLEND"
            End Select
        'Require a start date
            .StartSerial = ToSerial(Block(r, TX_START))
            If .StartSerial <= 0# Then
                LegError LegIndex, "Start date missing or not a date"
            End If
        'Read the maturity date; blank means open
            If IsBlankCell(Block(r, TX_MAT)) Then
                .MatSerial = 0#
            Else
                .MatSerial = ToSerial(Block(r, TX_MAT))
                If .MatSerial <= 0# Then
                    LegError LegIndex, "Maturity date not a date"
                ElseIf .StartSerial > 0# And .MatSerial < .StartSerial Then
                    LegError LegIndex, "Maturity date before the start date"
                End If
            End If
        'Set the sign from the side
            .Side = UTxt(Block(r, TX_SIDE))
            If .Side = SD_LENT Then
                .Sign = 1
            ElseIf .Side = SD_RECEIVED Then
                .Sign = -1
            Else
                LegError LegIndex, "Side must be LENT or RECEIVED"
            End If
        End With

'------------------------------------------------------------------------------
' ASSET AND AMOUNT
'------------------------------------------------------------------------------
    'Read the asset fields
        ReadLegAsset LegIndex, Block, r
    'Work on the leg being read
        With mLegs(LegIndex)
        'Require a market value above 0
            If IsNum(Block(r, TX_MV)) Then
                .MarketValue = VBA.CDbl(Block(r, TX_MV))
            End If
            If .MarketValue <= 0# Then
                LegError LegIndex, "Market value must be a number above 0"
            End If
        'Value the leg in the reporting currency
            .Ccy = UTxt(Block(r, TX_CCY))
            If FXRate(.Ccy) = 0# Then
                LegError LegIndex, "Currency '" & .Ccy & "' not in the FX table"
            Else
                .Value = .MarketValue * FXRate(.Ccy)
            End If
        'Key cash by its currency
            If .Asset = AT_CASH Then
                .IssueKey = "CASH " & .Ccy
            End If
        'Describe the issue, for the comparison across its legs
            .Signature = .Asset & "|" & .Issuer & "|" & .CQS & "|" & .SecMatSerial & "|" & .Listing & "|" & _
                .FundH10 & "|" & .Ccy
        End With

'------------------------------------------------------------------------------
' WORKSHEET ERROR VALUES
'------------------------------------------------------------------------------
    'An error value would read as a blank cell and take the field's default
        Found = ErrorColumns(Block, r, TX_HEADERS, TX_COMMENT)
        If VBA.Len(Found) > 0 Then
            LegError LegIndex, "Worksheet error value in: " & Found
        End If
    'Find the haircut of a valid leg
        If mLegs(LegIndex).Valid Then
            FindHaircut LegIndex
        End If

End Sub

Private Function ErrorColumns( _
    ByRef Block As Variant, _
    ByVal r As Long, _
    ByVal Headers As String, _
    ByVal SkipColumn As Long) As String
'
'------------------------------------------------------------------------------
'                      LIST CELLS HOLDING AN ERROR VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   List the headers of the cells of one input row that hold a worksheet
'   error value (#N/A, #VALUE! and others).
'
' WHY THIS EXISTS
'   An error value would read as a blank cell and take the field's
'   default; it is reported instead.
'
' INPUTS
'   Block, r
'     The input block and the row in it.
'
'   Headers
'     The table's headers, separated by |.
'
'   SkipColumn
'     A column that is not read, or 0.
'
' RETURNS
'   The headers, separated by "; ", or an empty string.
'
' BEHAVIOR
'   Scans every column of the row except SkipColumn and lists the header
'   of each cell holding an error value.
'
' ERROR POLICY
'   None raised by the procedure itself.
'
' DEPENDENCIES
'   AppendText.
'
' NOTES
'   A column beyond the headers is named "column n".
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.ErrorColumns"

    Dim Names()     As String   'Headers, 0-based
    Dim Col         As Long     'Column in Block
    Dim Label       As String   'Header of the column

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    'Split the headers
        Names = Split(Headers, "|")
    'List the header of every read cell that holds an error value
        For Col = 1 To UBound(Block, 2)
            If Col <> SkipColumn Then
                If IsError(Block(r, Col)) Then
                    If Col - 1 <= UBound(Names) Then
                        Label = Names(Col - 1)
                    Else
                        Label = "column " & Col
                    End If
                    ErrorColumns = AppendText(ErrorColumns, Label)
                End If
            End If
        Next Col

End Function

Private Sub ReadLegAsset( _
    ByVal LegIndex As Long, _
    ByRef Block As Variant, _
    ByVal r As Long)
'
'------------------------------------------------------------------------------
'                           READ THE ASSET OF A LEG
'------------------------------------------------------------------------------
' PURPOSE
'   Read and check the asset fields of one leg: asset type, issue ID and
'   the fields each asset type needs.
'
' WHY THIS EXISTS
'   The fields a leg needs depend on its asset type; this keeps those
'   checks apart from the transaction fields of ReadLeg.
'
' INPUTS
'   LegIndex
'     Index in mLegs.
'
'   Block, r
'     The input block and the row in it.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Refuses an unknown asset type and returns
'   - Requires an issue ID for every non-cash leg
'   - Debt: checks the issuer type, CQS and security maturity, and sets the
'     residual maturity in years
'   - Equity: checks the listing
'   - Fund: checks the fund haircut H10
'
' ERROR POLICY
'   Problems are logged through LegError, not raised.
'
' DEPENDENCIES
'   LegError; CORE_Util for UTxt, ToDbl, ToSerial, TryDbl and IsBlankCell;
'   CORE_Config for the columns and codes.
'
' NOTES
'   Cash, gold and other assets need no further field.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.ReadLegAsset"

    Dim Number      As Double   'A numeric field read

'------------------------------------------------------------------------------
' ASSET TYPE AND ISSUE
'------------------------------------------------------------------------------
    'Work on the leg being read
        With mLegs(LegIndex)
        'Refuse an unknown asset type
            .Asset = UTxt(Block(r, TX_ASSET))
            Select Case .Asset
                Case AT_CASH, AT_DEBT, AT_EQUITY, AT_GOLD, AT_FUND, AT_OTHER
                Case Else
                    LegError LegIndex, "Asset type '" & .Asset & "' is not CASH, DEBT, EQUITY, GOLD, FUND or OTHER"
                    Exit Sub
            End Select
        'Require an issue ID for every non-cash leg
            If .Asset <> AT_CASH Then
                .IssueKey = UTxt(Block(r, TX_ISSUE))
                If VBA.Len(.IssueKey) = 0 Then
                    LegError LegIndex, "Issue ID missing: every non-cash leg names its issue"
                End If
            End If

'------------------------------------------------------------------------------
' DEBT SECURITY
'------------------------------------------------------------------------------
        'Check the fields of each asset type
            Select Case .Asset
                Case AT_DEBT
                    .Issuer = UTxt(Block(r, TX_ISSUER))
                    Select Case .Issuer
                        Case IS_SOV, IS_INST, IS_CORP, IS_SEC
                        Case Else
                            LegError LegIndex, "Issuer type must be SOV, INST, CORP or SEC"
                    End Select
                    If Not IsBlankCell(Block(r, TX_CQS)) Then
                        Number = ToDbl(Block(r, TX_CQS), -1#)
                        If Number < 1# Or Number > 6# Or Number <> Int(Number) Then
                            LegError LegIndex, "Issue CQS must be blank or a whole number from 1 to 6"
                        Else
                            .CQS = VBA.CLng(Number)
                        End If
                    End If
                    .SecMatSerial = ToSerial(Block(r, TX_SECMAT))
                    If .SecMatSerial <= 0# Then
                        LegError LegIndex, "Security maturity date missing or not a date"
                    ElseIf .SecMatSerial <= pAsOf Then
                        LegError LegIndex, "Security matured on or before the reporting date"
                    Else
                        .ResidualYears = (.SecMatSerial - pAsOf) / pDaysYear
                    End If

'------------------------------------------------------------------------------
' EQUITY AND FUND
'------------------------------------------------------------------------------
                Case AT_EQUITY
                    .Listing = UTxt(Block(r, TX_LISTING))
                    Select Case .Listing
                        Case EL_MAIN, EL_LISTED, EL_OTHER
                        Case Else
                            LegError LegIndex, "Equity listing must be MAIN, LISTED or OTHER"
                    End Select
                Case AT_FUND
                    If TryDbl(Block(r, TX_FUNDH), -1#, Number) Then
                        If Number < 0# Or Number > 1# Then
                            LegError LegIndex, "Fund haircut H10 must be between 0 and 1 (look-through or " & _
                                "the highest haircut the fund may invest in)"
                        Else
                            .FundH10 = Number
                        End If
                    Else
                        LegError LegIndex, "Fund haircut H10 must be a number"
                    End If
                Case Else
                    'Cash, gold and other assets need no further field
            End Select
        End With

End Sub

Private Sub FindHaircut( _
    ByVal LegIndex As Long)
'
'------------------------------------------------------------------------------
'                          FIND THE HAIRCUT OF A LEG
'------------------------------------------------------------------------------
' PURPOSE
'   Build the haircut-table key of a valid leg, and set its H10 and
'   eligibility. A leg other than a fund whose key is not in the table for
'   its regime is not eligible: lent, it takes NonEligibleHaircut; received,
'   it is not recognised as collateral. A fund is eligible with the H10
'   entered on its leg.
'
' WHY THIS EXISTS
'   The haircut table is keyed by asset, issuer category, credit-quality
'   band and maturity band; the key is derived from the leg's fields.
'
' INPUTS
'   LegIndex
'     Index in mLegs of a valid leg.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Builds the key: CASH, GOLD, EQUITY_MAIN, EQUITY_LISTED, or DEBT_
'     with the issuer category, the credit-quality band and, except for
'     CQS 4, the maturity band
'   - Treats unrated institution debt as CQS 2-3, with an INFO
'   - Looks the key up for the regime of the leg's netting set; a fund
'     takes its own H10
'   - Recognises an eligible leg; a non-eligible leg lent takes
'     NonEligibleHaircut with an INFO; a non-eligible leg received is not
'     recognised, with a WARNING
'
' ERROR POLICY
'   None raised by the procedure itself.
'
' DEPENDENCIES
'   MaturityBand, TableRow, LegNote; CORE_Config for the codes.
'
' NOTES
'   An equity of listing OTHER and an asset of type OTHER get no key and
'   are therefore not eligible.
'
' REFERENCE
'   CRE22.49, 22.54; CRR Art. 197, 224(1), (4)-(6).
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.FindHaircut"

    Dim Regime      As String   'Regime of the leg's netting set
    Dim Category    As String   'Haircut-table issuer category: SOV, OTH or SEC
    Dim Band        As String   'Credit-quality band
    Dim Position    As Long     'Row in the haircut table

'------------------------------------------------------------------------------
' BUILD THE KEY
'------------------------------------------------------------------------------
    'Work on the leg
        With mLegs(LegIndex)
        'Take the regime of the leg's netting set
            Regime = mNS(.NsIndex).Regime
        'Build the haircut-table key of the asset
            Select Case .Asset
                Case AT_CASH
                    .HaircutKey = "CASH"
                Case AT_GOLD
                    .HaircutKey = "GOLD"
                Case AT_EQUITY
                    If .Listing = EL_MAIN Then
                        .HaircutKey = "EQUITY_MAIN"
                    ElseIf .Listing = EL_LISTED Then
                        .HaircutKey = "EQUITY_LISTED"
                    End If
                Case AT_DEBT
                    Select Case .Issuer
                        Case IS_SOV
                            Category = "SOV"
                        Case IS_SEC
                            Category = "SEC"
                        Case Else
                            Category = "OTH"
                    End Select
                    Select Case .CQS
                        Case 1
                            Band = "CQS1"
                        Case 2, 3
                            Band = "CQS23"
                        Case 4
                            Band = "CQS4"
                        Case 5, 6
                            Band = "CQS56"
                        Case Else
                            If .Issuer = IS_INST Then
                                Band = "CQS23"
                                LegNote LegIndex, SEV_INFO, "Unrated institution debt treated as CQS 2-3: the " & _
                                    "conditions of CRE22.34(3)(b) / CRR Art. 197(4) are assumed met"
                            End If
                    End Select
                    If Band = "CQS4" Then
                        .HaircutKey = "DEBT_" & Category & "_CQS4_ALL"
                    ElseIf VBA.Len(Band) > 0 Then
                        .HaircutKey = "DEBT_" & Category & "_" & Band & "_" & MaturityBand(.ResidualYears)
                    End If
                Case Else
                    'Funds are keyed below; other assets have no key
            End Select

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        'Take the fund's own H10, or look the key up for the regime
            If .Asset = AT_FUND Then
                .HaircutKey = "FUND"
                .Eligible = True
                .H10 = .FundH10
            Else
                Position = TableRow(mHcIndex, mHcRegimes, .HaircutKey, Regime)
                If Position > 0 Then
                    .Eligible = True
                    .H10 = mHcValues(Position)
                End If
            End If
        'Recognise an eligible leg; treat a non-eligible one by its side
            If .Eligible Then
                .Recognised = True
            ElseIf .Sign > 0 Then
                .Recognised = True
                .H10 = pNonElig
                LegNote LegIndex, SEV_INFO, "Not eligible as collateral: lent with the haircut " & _
                    "NonEligibleHaircut " & Format$(pNonElig, "0%")
            Else
                .Recognised = False
                LegNote LegIndex, SEV_WARN, "Not eligible as collateral: received but not recognised"
            End If
        End With

End Sub

Private Function MaturityBand( _
    ByVal Years As Double) As String
'
'------------------------------------------------------------------------------
'                            FIND THE MATURITY BAND
'------------------------------------------------------------------------------
' PURPOSE
'   Residual-maturity band of a debt security in the haircut keys: LE1,
'   1TO3, 3TO5, 5TO10 or GT10, upper bounds included.
'
' REFERENCE
'   CRE22.49; CRR Art. 224 Table 1.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Return the band whose upper bound is reached first
        If Years <= 1# Then
            MaturityBand = "LE1"
        ElseIf Years <= 3# Then
            MaturityBand = "1TO3"
        ElseIf Years <= 5# Then
            MaturityBand = "3TO5"
        ElseIf Years <= 10# Then
            MaturityBand = "5TO10"
        Else
            MaturityBand = "GT10"
        End If

End Function

'
'------------------------------------------------------------------------------
'
'                                 TRANSACTIONS
'
'------------------------------------------------------------------------------
'

Private Sub BuildTransactions()

'
'------------------------------------------------------------------------------
'                            BUILD THE TRANSACTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Group the legs into transactions by Transaction ID, check each
'   transaction as a whole, and reject it with all its legs when any leg is
'   invalid or the legs disagree. Count the transactions of each netting
'   set. A rejected transaction whose legs name several netting sets is
'   counted, as rejected, in each of them, so that none is reported VALID
'   without its legs.
'
' WHY THIS EXISTS
'   A transaction is accepted or rejected as a whole; its legs are read one
'   row at a time.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Creates a transaction from the first leg of each Transaction ID and
'     checks every further leg against it
'   - Counts each leg in its transaction and notes transactions with an
'     invalid leg
'   - Checks each transaction and counts it in the netting set of its
'     first leg
'   - Marks the legs of rejected transactions as not recognised and counts
'     the transaction, rejected, in every other netting set its legs name
'
' ERROR POLICY
'   Problems are logged, not raised.
'
' DEPENDENCIES
'   CheckSameTransaction, AddLegToTransaction, CheckTransaction,
'   AppendText; CORE_Util for KeyIndex and KeyAdd.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "CORE_Engine.BuildTransactions"

    Dim LegInvalid()        As Boolean      'Per transaction: some leg is invalid
    Dim RejectedKeys        As Collection   'Transaction|netting set pairs counted as rejected
    Dim TransactionIndex    As Long         'Transaction
    Dim LegIndex            As Long         'Leg
    Dim Ns                  As Long         'Netting set

'------------------------------------------------------------------------------
' GROUP
'------------------------------------------------------------------------------
    'Nothing to group without legs
        If mLegCount = 0 Then
            Exit Sub
        End If
    'Allocate at most one transaction per leg
        ReDim mTx(1 To mLegCount)
        ReDim LegInvalid(1 To mLegCount)
    'Attach each leg to its transaction, creating it from its first leg
        For LegIndex = 1 To mLegCount
            TransactionIndex = KeyIndex(mTxIndex, mLegs(LegIndex).TxKey)
            If TransactionIndex = 0 Then
                mTxCount = mTxCount + 1
                TransactionIndex = mTxCount
                KeyAdd mTxIndex, mLegs(LegIndex).TxKey, TransactionIndex
                With mTx(TransactionIndex)
                    .ID = mLegs(LegIndex).TxID
                    .NsIndex = mLegs(LegIndex).NsIndex
                    .TxType = mLegs(LegIndex).TxType
                    .StartSerial = mLegs(LegIndex).StartSerial
                    .MatSerial = mLegs(LegIndex).MatSerial
                    .FirstRow = mLegs(LegIndex).Row
                End With
            Else
                CheckSameTransaction TransactionIndex, LegIndex
            End If
            mLegs(LegIndex).TxIndex = TransactionIndex
            AddLegToTransaction TransactionIndex, LegIndex
            If Not mLegs(LegIndex).Valid Then
                LegInvalid(TransactionIndex) = True
            End If
        Next LegIndex

'------------------------------------------------------------------------------
' CHECK EACH TRANSACTION
'------------------------------------------------------------------------------
    'Check each transaction and count it in the netting set of its first leg
        For TransactionIndex = 1 To mTxCount
            CheckTransaction TransactionIndex, LegInvalid(TransactionIndex)
            If mTx(TransactionIndex).NsIndex > 0 Then
                With mNS(mTx(TransactionIndex).NsIndex)
                    .TxCount = .TxCount + 1
                    If mTx(TransactionIndex).Rejected Then
                        .TxRejected = .TxRejected + 1
                    End If
                End With
            End If
        Next TransactionIndex

'------------------------------------------------------------------------------
' MARK THE LEGS OF REJECTED TRANSACTIONS
'------------------------------------------------------------------------------
    'Mark the legs of rejected transactions and count the transaction once
    'in each other netting set its legs name
        Set RejectedKeys = New Collection
        For LegIndex = 1 To mLegCount
            TransactionIndex = mLegs(LegIndex).TxIndex
            If mTx(TransactionIndex).Rejected Then
                mLegs(LegIndex).TxRejected = True
                mLegs(LegIndex).Recognised = False
                If mLegs(LegIndex).Valid Then
                    mLegs(LegIndex).Messages = AppendText(mLegs(LegIndex).Messages, "Transaction rejected")
                End If
                Ns = mLegs(LegIndex).NsIndex
                If Ns > 0 And Ns <> mTx(TransactionIndex).NsIndex Then
                    If KeyIndex(RejectedKeys, TransactionIndex & "|" & Ns) = 0 Then
                        KeyAdd RejectedKeys, TransactionIndex & "|" & Ns, 1
                        With mNS(Ns)
                            .TxCount = .TxCount + 1
                            .TxRejected = .TxRejected + 1
                        End With
                    End If
                End If
            End If
        Next LegIndex

End Sub

Private Sub CheckSameTransaction( _
    ByVal t As Long, _
    ByVal i As Long)
'
'------------------------------------------------------------------------------
'                     CHECK A FURTHER LEG OF A TRANSACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a further leg of a transaction agrees with its first leg on
'   the netting set, type, start date and maturity date.
'
' INPUTS
'   t
'     Index in mTx of the transaction.
'
'   i
'     Index in mLegs of the further leg.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Reject the leg at the first field that differs from the first leg
        With mLegs(i)
            If .NsIndex <> mTx(t).NsIndex Then
                LegError i, "Netting set differs from the transaction's first leg (row " & mTx(t).FirstRow & ")"
            ElseIf .TxType <> mTx(t).TxType Then
                LegError i, "Transaction type differs from the transaction's first leg (row " & _
                    mTx(t).FirstRow & ")"
            ElseIf .StartSerial <> mTx(t).StartSerial Or .MatSerial <> mTx(t).MatSerial Then
                LegError i, "Start or maturity date differs from the transaction's first leg (row " & _
                    mTx(t).FirstRow & ")"
            End If
        End With

End Sub

Private Sub AddLegToTransaction( _
    ByVal t As Long, _
    ByVal i As Long)
'
'------------------------------------------------------------------------------
'                        COUNT A LEG IN ITS TRANSACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Count a leg in its transaction: sides, non-cash legs and the currency
'   of the legs lent.
'
' WHY THIS EXISTS
'   CheckTransaction judges a transaction from these counts.
'
' INPUTS
'   t
'     Index in mTx of the transaction.
'
'   i
'     Index in mLegs of the leg.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Counts the leg, its side and, when it is not cash, its side's non-cash
'   legs; keeps the currency of the legs lent.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   None
'
' NOTES
'   LentCcy becomes "*" when the legs lent are in several currencies.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Count the leg by side, non-cash legs, and the currency of the legs lent
        With mTx(t)
            .LegCount = .LegCount + 1
            If mLegs(i).Sign > 0 Then
                .LentCount = .LentCount + 1
                If mLegs(i).Asset <> AT_CASH Then
                    .LentNonCash = .LentNonCash + 1
                End If
                If VBA.Len(.LentCcy) = 0 Then
                    .LentCcy = mLegs(i).Ccy
                ElseIf .LentCcy <> mLegs(i).Ccy Then
                    .LentCcy = "*"
                End If
            ElseIf mLegs(i).Sign < 0 Then
                .ReceivedCount = .ReceivedCount + 1
                If mLegs(i).Asset <> AT_CASH Then
                    .ReceivedNonCash = .ReceivedNonCash + 1
                End If
            End If
        End With

End Sub

Private Sub CheckTransaction( _
    ByVal t As Long, _
    ByVal HasInvalidLeg As Boolean)
'
'------------------------------------------------------------------------------
'                        CHECK A TRANSACTION AS A WHOLE
'------------------------------------------------------------------------------
' PURPOSE
'   Check a transaction as a whole: rejected when a leg is invalid, when it
'   has no leg lent or no leg received, or, without netting, when its legs
'   lent are in several currencies (the single-transaction formula needs
'   one exposure currency). Warn when its legs do not look like its type,
'   and decide whether it is short-term.
'
' WHY THIS EXISTS
'   These conditions concern all the legs of a transaction together and
'   cannot be checked on one leg.
'
' INPUTS
'   t
'     Index in mTx of the transaction.
'
'   HasInvalidLeg
'     True when at least one of its legs is invalid.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Rejects the transaction, and returns, when a leg is invalid, a side
'     is missing, or the legs lent of a transaction without netting are in
'     several currencies
'   - Warns when the securities lent or received do not fit the type, and
'     when the maturity date is before the reporting date
'   - Sets ShortTerm when the original maturity is three months or less
'
' ERROR POLICY
'   Problems are logged as ERROR or WARNING messages, not raised.
'
' DEPENDENCIES
'   LogMsg; CORE_Config for the transaction types and SHORT_TERM_MONTHS.
'
' NOTES
'   An invalid leg has already been logged, so its rejection logs nothing
'   further here.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.CheckTransaction"

    Dim Reference   As String   'Transaction ID, for the messages

'------------------------------------------------------------------------------
' REJECT
'------------------------------------------------------------------------------
    'Work on the transaction
        With mTx(t)
        'Keep the ID for the messages
            Reference = .ID
        'Reject a transaction with an invalid leg
            If HasInvalidLeg Then
                .Rejected = True
                Exit Sub
            End If
        'Reject a transaction without a leg lent or a leg received
            If .LentCount = 0 Or .ReceivedCount = 0 Then
                .Rejected = True
                LogMsg SEV_ERROR, SH_TX, Reference, "Transaction needs at least one leg LENT and one RECEIVED - " & _
                    "transaction rejected.", .FirstRow
                Exit Sub
            End If
        'Without netting, reject legs lent in several currencies
            If Not mNS(.NsIndex).MNA And .LentCcy = "*" Then
                .Rejected = True
                LogMsg SEV_ERROR, SH_TX, Reference, "Without an eligible master netting agreement the legs lent " & _
                    "must share one currency - transaction rejected.", .FirstRow
                Exit Sub
            End If

'------------------------------------------------------------------------------
' WARN
'------------------------------------------------------------------------------
        'Warn when the securities lent or received do not fit the type
            Select Case .TxType
                Case TT_REPO, TT_SECLEND
                    If .LentNonCash = 0 Then
                        LogMsg SEV_WARN, SH_TX, Reference, .TxType & " without a security lent: check the type " & _
                            "or the sides.", .FirstRow
                    End If
                Case TT_REVREPO, TT_MARGINLEND
                    If .LentNonCash > 0 Then
                        LogMsg SEV_WARN, SH_TX, Reference, .TxType & " with a security lent: check the type " & _
                            "or the sides.", .FirstRow
                    End If
                Case TT_SECBORROW
                    If .ReceivedNonCash = 0 Then
                        LogMsg SEV_WARN, SH_TX, Reference, "SECBORROW without a security received: check the " & _
                            "type or the sides.", .FirstRow
                    End If
                Case Else
                    'No other type reaches this point: an unknown type invalidates its leg
            End Select
        'Warn when the transaction matured before the reporting date
            If .MatSerial > 0# And .MatSerial < pAsOf Then
                LogMsg SEV_WARN, SH_TX, Reference, "Maturity date before the reporting date: the transaction " & _
                    "is still calculated.", .FirstRow
            End If

'------------------------------------------------------------------------------
' SHORT-TERM
'------------------------------------------------------------------------------
        'Original maturity of three months or less; an open transaction is not
        'short-term
            If .MatSerial > 0# Then
                .ShortTerm = (VBA.CDbl(DateAdd("m", SHORT_TERM_MONTHS, VBA.CDate(.StartSerial))) >= .MatSerial)
            End If
        End With

End Sub

'
'------------------------------------------------------------------------------
'
'                            NETTING-SET CHECKS
'
'------------------------------------------------------------------------------
'

Private Sub CheckNettingSets()

'
'------------------------------------------------------------------------------
'                            CHECK THE NETTING SETS
'------------------------------------------------------------------------------
' PURPOSE
'   For every netting set: its transaction types, whether every transaction
'   is short-term, the zero-haircut conditions, the consistency of each
'   security issue across its legs, the holding period override, the
'   risk-weight key, and then its status.
'
' WHY THIS EXISTS
'   These checks need every transaction of a netting set, so they run once
'   all transactions are built.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Notes, from the transactions not rejected, whether each netting set
'     holds repo-style or capital-market transactions and whether all are
'     short-term
'   - Checks the zero-haircut conditions and the issues
'   - Checks the holding period override, finds the risk weight and sets
'     the status of each netting set
'
' ERROR POLICY
'   Problems are logged, not raised.
'
' DEPENDENCIES
'   CheckZeroHaircut, CheckIssues, CheckHoldingOverride, FindRiskWeight,
'   SetStatus.
'
' NOTES
'   Margin lending is the only capital-market-driven transaction type;
'   every other type is repo-style.
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "CORE_Engine.CheckNettingSets"

    Dim Ns                  As Long 'Netting set
    Dim TransactionIndex    As Long 'Transaction

'------------------------------------------------------------------------------
' TRANSACTION TYPES AND SHORT-TERM STATUS
'------------------------------------------------------------------------------
    'Start every netting set as all short-term
        For Ns = 1 To mNsCount
            mNS(Ns).AllShortTerm = True
        Next Ns
    'Note the types and the short-term status of the transactions not rejected
        For TransactionIndex = 1 To mTxCount
            If mTx(TransactionIndex).NsIndex > 0 And Not mTx(TransactionIndex).Rejected Then
                With mNS(mTx(TransactionIndex).NsIndex)
                    If mTx(TransactionIndex).TxType = TT_MARGINLEND Then
                        .HasCapitalMarket = True
                    Else
                        .HasRepoStyle = True
                    End If
                    If Not mTx(TransactionIndex).ShortTerm Then
                        .AllShortTerm = False
                    End If
                End With
            End If
        Next TransactionIndex

'------------------------------------------------------------------------------
' CONDITIONS, ISSUES, HOLDING PERIOD AND RISK WEIGHT
'------------------------------------------------------------------------------
    'Check the zero-haircut conditions and the issues
        CheckZeroHaircut
        CheckIssues
    'Check the override, find the risk weight and set the status of each netting set
        For Ns = 1 To mNsCount
            CheckHoldingOverride Ns
            FindRiskWeight Ns
            SetStatus Ns
        Next Ns

End Sub

Private Sub CheckZeroHaircut()

'
'------------------------------------------------------------------------------
'                      CHECK THE ZERO-HAIRCUT CONDITIONS
'------------------------------------------------------------------------------
' PURPOSE
'   For every netting set declared to meet the conditions of the 0%
'   haircut, check those the inputs show: repo-style transactions only,
'   every recognised leg cash or an eligible debt security of a central
'   government or central bank, one currency, daily remargining. The other
'   conditions (the 0% risk weight of the securities, core market
'   participant, documentation, settlement) are declared by the Y on
'   NettingSets and listed as INFO.
'
' WHY THIS EXISTS
'   The 0% haircut is declared by the user; the conditions the inputs can
'   show are checked rather than taken on trust.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Collects, per netting set, the conditions not met: margin lending,
'     remargining that is not daily
'   - Collects, per recognised leg, a second currency, a leg that is
'     neither cash nor a central government security, and a leg that is not
'     eligible
'   - Invalidates each declared netting set with a condition not met, and
'     lists the declared conditions as INFO for the others
'
' ERROR POLICY
'   Problems are logged through NsError, not raised.
'
' DEPENDENCIES
'   AppendText, NsError, LogMsg.
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.60, 22.36-22.37; CRR Art. 227.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.CheckZeroHaircut"

    Dim Failures()  As String   'Per netting set: conditions not met
    Dim FirstCcy()  As String   'Per netting set: currency the next leg is compared with
    Dim Ns          As Long     'Netting set
    Dim LegIndex    As Long     'Leg

'------------------------------------------------------------------------------
' NETTING-SET CONDITIONS
'------------------------------------------------------------------------------
    'Allocate the failures and currencies per netting set
        ReDim Failures(1 To mNsCount)
        ReDim FirstCcy(1 To mNsCount)
    'Collect margin lending and remargining that is not daily
        For Ns = 1 To mNsCount
            With mNS(Ns)
                If .ZeroHaircut And .HasCapitalMarket Then
                    Failures(Ns) = AppendText(Failures(Ns), "margin lending is not a repo-style transaction")
                End If
                If .ZeroHaircut And .RemarginBD > 1# Then
                    Failures(Ns) = AppendText(Failures(Ns), "remargining is not daily")
                End If
            End With
        Next Ns

'------------------------------------------------------------------------------
' LEG CONDITIONS
'------------------------------------------------------------------------------
    'Collect a second currency and legs that are not eligible central government
    'securities or cash
        For LegIndex = 1 To mLegCount
            Ns = mLegs(LegIndex).NsIndex
            If Ns > 0 Then
                If mNS(Ns).ZeroHaircut And mLegs(LegIndex).Recognised Then
                    If VBA.Len(FirstCcy(Ns)) = 0 Then
                        FirstCcy(Ns) = mLegs(LegIndex).Ccy
                    ElseIf mLegs(LegIndex).Ccy <> FirstCcy(Ns) Then
                        Failures(Ns) = AppendText(Failures(Ns), "legs in more than one currency (" & _
                            mLegs(LegIndex).LegID & ")")
                        FirstCcy(Ns) = mLegs(LegIndex).Ccy
                    End If
                    If Not (mLegs(LegIndex).Asset = AT_CASH Or _
                        (mLegs(LegIndex).Asset = AT_DEBT And mLegs(LegIndex).Issuer = IS_SOV)) Then
                        Failures(Ns) = AppendText(Failures(Ns), "leg " & mLegs(LegIndex).LegID & " is neither cash " & _
                            "nor a central government security")
                    ElseIf Not mLegs(LegIndex).Eligible Then
                        Failures(Ns) = AppendText(Failures(Ns), "leg " & mLegs(LegIndex).LegID & " is not eligible")
                    End If
                End If
            End If
        Next LegIndex

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Invalidate a declared netting set with a condition not met; list the
    'declared conditions as INFO otherwise
        For Ns = 1 To mNsCount
            If mNS(Ns).ZeroHaircut Then
                If VBA.Len(Failures(Ns)) > 0 Then
                    NsError Ns, "0% haircut declared but its conditions are not met: " & Failures(Ns)
                Else
                    LogMsg SEV_INFO, SH_NS, mNS(Ns).ID, "0% haircut applied: the conditions on the " & _
                        "0% risk weight of the securities, the counterparty, documentation and " & _
                        "settlement are taken as declared.", mNS(Ns).Row
                End If
            End If
        Next Ns

End Sub

Private Sub CheckIssues()

'
'------------------------------------------------------------------------------
'                      CHECK THE ISSUES ACROSS THEIR LEGS
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the recognised legs of one security issue in a netting set
'   describe it in the same way (asset type, issuer, CQS, maturity,
'   listing, fund haircut and currency), since the netting formula gives
'   each issue one haircut.
'
' WHY THIS EXISTS
'   The netting formula gives each issue one haircut, so its legs must
'   agree on everything that sets it.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Compares the signature of each recognised non-cash leg with the first
'   leg of the same issue in the same netting set.
'
' ERROR POLICY
'   A difference invalidates the netting set through NsError.
'
' DEPENDENCIES
'   NsError; CORE_Util for KeyIndex and KeyAdd.
'
' NOTES
'   The signature is built by ReadLeg.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.CheckIssues"

    Dim Seen        As Collection   'Netting set and issue key to the first leg
    Dim Key         As String       'Netting set and issue key
    Dim First       As Long         'First leg of an issue
    Dim LegIndex    As Long         'Leg

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Compare each recognised non-cash leg with the first leg of its issue
        Set Seen = New Collection
        For LegIndex = 1 To mLegCount
            If mLegs(LegIndex).NsIndex > 0 And mLegs(LegIndex).Recognised And mLegs(LegIndex).Asset <> AT_CASH Then
                Key = mLegs(LegIndex).NsIndex & "|" & mLegs(LegIndex).IssueKey
                First = KeyIndex(Seen, Key)
                If First = 0 Then
                    KeyAdd Seen, Key, LegIndex
                ElseIf mLegs(First).Signature <> mLegs(LegIndex).Signature Then
                    NsError mLegs(LegIndex).NsIndex, "Issue " & mLegs(LegIndex).IssueKey & _
                        " is described differently on legs " & mLegs(First).LegID & " and " & mLegs(LegIndex).LegID
                End If
            End If
        Next LegIndex

End Sub

Private Sub CheckHoldingOverride( _
    ByVal Ns As Long)
'
'------------------------------------------------------------------------------
'                      CHECK THE HOLDING PERIOD OVERRIDE
'------------------------------------------------------------------------------
' PURPOSE
'   Warn when a holding period override is not used: it can only lengthen
'   the holding period, and it is shorter than every one it could lengthen.
'   With netting that is the minimum of the netting set's transaction types
'   taken together. Without netting the minimum is set per transaction, so
'   an override between the repo-style and the capital-market minimum still
'   lengthens the repo-style transactions.
'
' WHY THIS EXISTS
'   An override that changes nothing would otherwise pass unnoticed.
'
' INPUTS
'   Ns
'     Index in mNS.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Returns without an override; otherwise finds the shortest minimum the
'   override could lengthen, applies the large-or-illiquid period, and
'   warns when the override is below it.
'
' ERROR POLICY
'   An unused override is a WARNING.
'
' DEPENDENCIES
'   MinimumHolding, Checked, SFT_HoldingPeriod; CORE_Util for Min2.
'
' NOTES
'   Disputes are left out of the minimum compared.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.CheckHoldingOverride"

    Dim Minimum     As Double   'Shortest holding period without override or disputes

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Work on the netting set
        With mNS(Ns)
        'Nothing to check without an override
            If .HoldingOverride <= 0# Then
                Exit Sub
            End If
        'Find the shortest minimum the override could lengthen
            If .MNA Or Not (.HasRepoStyle And .HasCapitalMarket) Then
                Minimum = MinimumHolding(.HasRepoStyle, .HasCapitalMarket)
            Else
                Minimum = Min2(pHoldRepo, pHoldCMT)
            End If
            Minimum = Checked(SFT_HoldingPeriod(Minimum, .LargeOrIlliquid, False, 0#, pHoldLarge), _
                "SFT_HoldingPeriod")
        'Warn when the override is below it
            If .HoldingOverride < Minimum Then
                LogMsg SEV_WARN, SH_NS, .ID, "Holding period override " & .HoldingOverride & _
                    " BD is below the minimum " & Minimum & " BD and is not used.", .Row
            End If
        End With

End Sub

Private Sub FindRiskWeight( _
    ByVal Ns As Long)
'
'------------------------------------------------------------------------------
'                    FIND THE RISK WEIGHT OF A NETTING SET
'------------------------------------------------------------------------------
' PURPOSE
'   Build the risk-weight key of a netting set from its counterparty
'   class, CQS, SCRA grade and short-term status, and find its risk weight;
'   an override replaces the table value. A netting set already INVALID is
'   not looked up, so that a missing input is not reported twice.
'
' WHY THIS EXISTS
'   The risk-weight table is keyed by counterparty class and credit
'   quality; the key is derived from the netting set's fields.
'
' INPUTS
'   Ns
'     Index in mNS.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Uses the short-term keys of an institution when every transaction
'     read and not rejected is short-term
'   - Builds SOV_ or CORP_ with the CQS or UNRATED, and INST_ with the CQS
'     or the SCRA grade; returns for any other class
'   - Uses an override with a WARNING, or looks the key up for the regime
'     of the netting set
'
' ERROR POLICY
'   A key missing from the table invalidates the netting set through
'   NsError.
'
' DEPENDENCIES
'   TableRow, NsError, LogMsg.
'
' NOTES
'   The key of an overridden netting set is marked "(override)".
'
' REFERENCE
'   CRE20.7, 20.18-20.21, CRE54.7; CRR Art. 114, 120-122, 306(1)(a).
'
' UPDATED
'   2026-10-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.FindRiskWeight"

    Dim ShortPart   As String   '"ST_" for a short-term institution exposure
    Dim Position    As Long     'Row in the risk-weight table

'------------------------------------------------------------------------------
' BUILD THE KEY
'------------------------------------------------------------------------------
    'Work on the netting set
        With mNS(Ns)
        'Use the short-term keys when every transaction is short-term
            If .AllShortTerm And .TxCount > .TxRejected Then
                ShortPart = "ST_"
            End If
        'Build the key from the counterparty class and credit quality
            Select Case .CptyClass
                Case CC_SOV, CC_CORP
                    If .CQS > 0 Then
                        .RWKey = .CptyClass & "_CQS" & .CQS
                    Else
                        .RWKey = .CptyClass & "_UNRATED"
                    End If
                Case CC_INST
                    If .CQS > 0 Then
                        .RWKey = "INST_" & ShortPart & "CQS" & .CQS
                    Else
                        .RWKey = "INST_" & ShortPart & "SCRA_" & .SCRA
                    End If
                Case Else
                    Exit Sub
            End Select

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        'Use an override instead of the table
            If .HasRWOverride Then
                .RiskWeight = .RWOverride
                LogMsg SEV_WARN, SH_NS, .ID, "Risk weight override " & Format$(.RWOverride, "0.0%") & _
                    " used instead of " & .RWKey & ".", .Row
                .RWKey = .RWKey & " (override)"
                Exit Sub
            End If
        'Do not look up a netting set already INVALID
            If .InputErrors > 0 Then
                Exit Sub
            End If
        'Look the key up for the regime of the netting set
            Position = TableRow(mRwIndex, mRwRegimes, .RWKey, .Regime)
            If Position = 0 Then
                NsError Ns, "Risk-weight key " & .RWKey & " not in the risk-weight table for " & .Regime
            Else
                .RiskWeight = mRwValues(Position)
            End If
        End With

End Sub

Private Sub SetStatus( _
    ByVal Ns As Long)
'
'------------------------------------------------------------------------------
'                       SET THE STATUS OF A NETTING SET
'------------------------------------------------------------------------------
' PURPOSE
'   Status of a netting set: INVALID with input errors, INCOMPLETE with a
'   rejected transaction, NO TRANSACTIONS, or VALID. Only a VALID netting
'   set receives an exposure value.
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' SET STATUS
'------------------------------------------------------------------------------
    'Take the first status that applies
        With mNS(Ns)
            If .InputErrors > 0 Then
                .Status = "INVALID: " & .InputErrors & " input error(s)"
            ElseIf .TxRejected > 0 Then
                .Status = "INCOMPLETE: " & .TxRejected & " of " & .TxCount & " transaction(s) rejected"
            ElseIf .TxCount = 0 Then
                .Status = "NO TRANSACTIONS"
            Else
                .Status = "VALID"
            End If
        End With

End Sub

'
'------------------------------------------------------------------------------
'
'                                 CALCULATION
'
'------------------------------------------------------------------------------
'

Private Sub ComputeNettingSets()

'
'------------------------------------------------------------------------------
'                          CALCULATE THE NETTING SETS
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate every netting set: the holding period of each netting set
'   with netting and of each transaction without, the haircut of every
'   leg in one pass, then E* of every VALID netting set, its RWA and own
'   funds requirement, and the totals per regime.
'
' WHY THIS EXISTS
'   The two exposure formulas need different holding periods and sums; this
'   procedure routes every leg to the right one.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Sets the holding period and haircut scale of each netting set with
'     netting, and the holding period of each transaction without
'   - Applies the haircut to every leg of a known netting set; adds the
'     recognised legs of VALID netting sets to the positions (with netting)
'     or to their transaction (without)
'   - Calculates E* with netting, then without
'   - Calculates the RWA and own funds of every VALID netting set and adds
'     them to the totals of its regime; counts the others as withheld
'
' ERROR POLICY
'   Raises ERR_FORMULA_DOMAIN through Checked when an SFT_ function
'   returns an error value.
'
' DEPENDENCIES
'   HoldingFor, ApplyHaircut, AddToPositions, AddToTransaction,
'   FinishNetted, FinishSingles, RegimeIndex, Checked;
'   SFT_CapitalRequirement.
'
' NOTES
'   The RWA is SFT_CapitalRequirement with a capital ratio of 1.
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "CORE_Engine.ComputeNettingSets"

    Dim Ns                  As Long 'Netting set
    Dim TransactionIndex    As Long 'Transaction
    Dim LegIndex            As Long 'Leg
    Dim RegimeSlot          As Long 'Regime index

'------------------------------------------------------------------------------
' HOLDING PERIODS
'------------------------------------------------------------------------------
    'Set the holding period and haircut scale of each netting set with netting
        For Ns = 1 To mNsCount
            With mNS(Ns)
                If .MNA Then
                    .HoldingBD = HoldingFor(Ns, .HasRepoStyle, .HasCapitalMarket)
                    .HaircutScale = Sqr((.RemarginBD + .HoldingBD - 1#) / pHcBase)
                End If
            End With
        Next Ns
    'Set the holding period of each transaction without netting from its type
        For TransactionIndex = 1 To mTxCount
            Ns = mTx(TransactionIndex).NsIndex
            If Ns > 0 Then
                If Not mNS(Ns).MNA Then
                    mTx(TransactionIndex).HoldingBD = HoldingFor(Ns, mTx(TransactionIndex).TxType <> TT_MARGINLEND, _
                        mTx(TransactionIndex).TxType = TT_MARGINLEND)
                End If
            End If
        Next TransactionIndex

'------------------------------------------------------------------------------
' LEGS
'------------------------------------------------------------------------------
    'Every leg of a known netting set gets its haircut. LegCalc shows it for
    'the recognised legs of transactions not rejected; only recognised legs
    'of VALID netting sets enter the sums
        For LegIndex = 1 To mLegCount
            Ns = mLegs(LegIndex).NsIndex
            If Ns > 0 Then
                If mNS(Ns).MNA Then
                    ApplyHaircut LegIndex, mNS(Ns).HoldingBD, vbNullString
                    If mLegs(LegIndex).Recognised And mNS(Ns).Status = "VALID" Then
                        AddToPositions Ns, LegIndex
                    End If
                Else
                    TransactionIndex = mLegs(LegIndex).TxIndex
                    If Ns = mTx(TransactionIndex).NsIndex Then
                        ApplyHaircut LegIndex, mTx(TransactionIndex).HoldingBD, mTx(TransactionIndex).LentCcy
                    Else
                    'A leg filed under another netting set than its transaction's
                    'first leg: the transaction is rejected, and any holding period
                    'it has is the first leg's netting set's, so the leg takes its
                    'own. The haircut is neither shown on LegCalc nor summed
                        ApplyHaircut LegIndex, HoldingFor(Ns, mTx(TransactionIndex).TxType <> TT_MARGINLEND, _
                            mTx(TransactionIndex).TxType = TT_MARGINLEND), mTx(TransactionIndex).LentCcy
                    End If
                    If mLegs(LegIndex).Recognised And mNS(Ns).Status = "VALID" Then
                        AddToTransaction TransactionIndex, LegIndex
                    End If
                End If
            End If
        Next LegIndex

'------------------------------------------------------------------------------
' EXPOSURE VALUE
'------------------------------------------------------------------------------
    'Calculate E* with netting, then without
        FinishNetted
        FinishSingles

'------------------------------------------------------------------------------
' CAPITAL AND TOTALS
'------------------------------------------------------------------------------
    'Calculate the capital of each VALID netting set and add it to the totals
    'of its regime; count the others as withheld
        For Ns = 1 To mNsCount
            With mNS(Ns)
                RegimeSlot = RegimeIndex(.Regime)
                mRegSets(RegimeSlot) = mRegSets(RegimeSlot) + 1
                If .Status = "VALID" Then
                    .RWA = Checked(SFT_CapitalRequirement(.EStar, .RiskWeight, 1#), "SFT_CapitalRequirement")
                    .OwnFunds = Checked(SFT_CapitalRequirement(.EStar, .RiskWeight, pCapRatio), _
                        "SFT_CapitalRequirement")
                    mRegValid(RegimeSlot) = mRegValid(RegimeSlot) + 1
                    mRegEStar(RegimeSlot) = mRegEStar(RegimeSlot) + .EStar
                    mRegRWA(RegimeSlot) = mRegRWA(RegimeSlot) + .RWA
                    mRegOwnFunds(RegimeSlot) = mRegOwnFunds(RegimeSlot) + .OwnFunds
                    mTxUsed = mTxUsed + .TxCount
                    mLegsUsed = mLegsUsed + .LegsUsed
                Else
                    mWithheld = mWithheld + 1
                End If
            End With
        Next Ns

End Sub

Private Function MinimumHolding( _
    ByVal RepoStyle As Boolean, _
    ByVal CapitalMarket As Boolean) As Double
'
'------------------------------------------------------------------------------
'                       FIND THE MINIMUM HOLDING PERIOD
'------------------------------------------------------------------------------
' PURPOSE
'   Minimum holding period of the transaction types covered: 5 business
'   days for repo-style transactions, 10 for other capital-market-driven
'   transactions, and the longer of the two when both are present.
'
' NOTES
'   The periods are the Params values pHoldRepo and pHoldCMT.
'
' REFERENCE
'   CRE22.57-22.58; CRR Art. 224(2). The rule for both present is in
'   CRE22.58 only; the CRR does not state it (PROTOTYPE.md, assumption 1).
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Take the longer minimum when both types are present
        If CapitalMarket And RepoStyle Then
            MinimumHolding = Max2(pHoldRepo, pHoldCMT)
        ElseIf CapitalMarket Then
            MinimumHolding = pHoldCMT
        Else
            MinimumHolding = pHoldRepo
        End If

End Function

Private Function HoldingFor( _
    ByVal Ns As Long, _
    ByVal RepoStyle As Boolean, _
    ByVal CapitalMarket As Boolean) As Double
'
'------------------------------------------------------------------------------
'                           FIND THE HOLDING PERIOD
'------------------------------------------------------------------------------
' PURPOSE
'   Holding period T_M of a netting set or of one transaction in it, from
'   the transaction types it covers and the netting set's drivers: the
'   20-day period for a large or illiquid netting set, the override, and
'   the doubling for disputes.
'
' REFERENCE
'   CRE22.57-22.58; CRR Art. 224(2), 285(3)-(4).
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Apply the netting set's drivers to the minimum of the types covered
        With mNS(Ns)
            HoldingFor = Checked(SFT_HoldingPeriod(MinimumHolding(RepoStyle, CapitalMarket), .LargeOrIlliquid, _
                .Disputes, .HoldingOverride, pHoldLarge), "SFT_HoldingPeriod")
        End With

End Function

Private Sub ApplyHaircut( _
    ByVal i As Long, _
    ByVal HoldingBD As Double, _
    ByVal ExposureCcy As String)
'
'------------------------------------------------------------------------------
'                          APPLY THE HAIRCUT TO A LEG
'------------------------------------------------------------------------------
' PURPOSE
'   Scale the haircut of one leg for its holding period and its netting
'   set's remargining frequency, and set its adjusted value. With
'   ExposureCcy given (single-transaction formula), a leg received in
'   another currency also takes the currency haircut. A 0% haircut netting
'   set takes no haircut at all.
'
' WHY THIS EXISTS
'   Both exposure formulas use the scaled haircut of every leg; the
'   currency haircut of a leg applies only without netting.
'
' INPUTS
'   i
'     The leg.
'
'   HoldingBD
'     T_M.
'
'   ExposureCcy
'     The currency of the legs lent of its transaction, or "" under
'     netting.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Stores the holding period and the haircut scale
'   - Scales H10 unless the 0% haircut applies
'   - Adjusts the value up when lent; down when received, with the
'     currency haircut when its currency differs from ExposureCcy
'
' ERROR POLICY
'   Raises ERR_FORMULA_DOMAIN through Checked when SFT_ScaledHaircut
'   returns an error value.
'
' DEPENDENCIES
'   Checked, SFT_ScaledHaircut.
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.46, 22.52, 22.59-22.60; CRR Art. 223, 224, 226, 227.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLY
'------------------------------------------------------------------------------
    'Work on the leg
        With mLegs(i)
        'Store the holding period and the scale; start without haircuts
            .HoldingBD = HoldingBD
            .Scale = Sqr((mNS(.NsIndex).RemarginBD + HoldingBD - 1#) / pHcBase)
            .H = 0#
            .Hfx = 0#
        'Scale the haircut unless the 0% haircut applies
            If Not mNS(.NsIndex).ZeroHaircut Then
                .H = Checked(SFT_ScaledHaircut(.H10, HoldingBD, mNS(.NsIndex).RemarginBD, pHcBase), _
                    "SFT_ScaledHaircut")
            End If
        'Adjust the value up when lent; down, with any currency haircut, when received
            If .Sign > 0 Then
                .Adjusted = .Value * (1# + .H)
            Else
                If VBA.Len(ExposureCcy) > 0 And .Ccy <> ExposureCcy And Not mNS(.NsIndex).ZeroHaircut Then
                    .Hfx = Checked(SFT_ScaledHaircut(pHfx, HoldingBD, mNS(.NsIndex).RemarginBD, pHcBase), _
                        "SFT_ScaledHaircut")
                End If
                .Adjusted = .Value * (1# - .H - .Hfx)
            End If
        End With

End Sub

Private Sub AddToPositions( _
    ByVal Ns As Long, _
    ByVal i As Long)
'
'------------------------------------------------------------------------------
'                          ADD A LEG TO THE POSITIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Add a recognised leg of a netting set with netting to the netting
'   set's totals, to the position of its currency and, for a security, to
'   the position of its issue.
'
' WHY THIS EXISTS
'   The netting formula works on positions per issue and per currency.
'
' INPUTS
'   Ns
'     The netting set.
'
'   i
'     The leg.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Counts the leg and adds its value to the lent or received total
'   - Adds it to the position of its currency
'   - Adds a security to the position of its issue, which takes its haircut
'
' ERROR POLICY
'   None raised by the procedure itself.
'
' DEPENDENCIES
'   PositionFor, AddLegValue.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Engine.AddToPositions"

    Dim PositionIndex   As Long 'Position

'------------------------------------------------------------------------------
' NETTING-SET TOTALS
'------------------------------------------------------------------------------
    'Work on the leg
        With mLegs(i)
        'Count the leg and add its value to the side it is on
            mNS(Ns).LegsUsed = mNS(Ns).LegsUsed + 1
            If .Sign > 0 Then
                mNS(Ns).SumE = mNS(Ns).SumE + .Value
            Else
                mNS(Ns).SumC = mNS(Ns).SumC + .Value
            End If

'------------------------------------------------------------------------------
' CURRENCY AND ISSUE
'------------------------------------------------------------------------------
        'Add the leg to the position of its currency
            PositionIndex = PositionFor(Ns, PK_CURRENCY, .Ccy, vbNullString, .Ccy)
            AddLegValue PositionIndex, i
        'Add a security to the position of its issue, with its haircut
            If .Asset <> AT_CASH Then
                PositionIndex = PositionFor(Ns, PK_ISSUE, .IssueKey, .Asset, .Ccy)
                mPos(PositionIndex).H = .H
                AddLegValue PositionIndex, i
            End If
        End With

End Sub

Private Function PositionFor( _
    ByVal Ns As Long, _
    ByVal Kind As String, _
    ByVal Key As String, _
    ByVal Asset As String, _
    ByVal Ccy As String) As Long
'
'------------------------------------------------------------------------------
'                          FIND OR CREATE A POSITION
'------------------------------------------------------------------------------
' PURPOSE
'   Index of a position of a netting set, created when new.
'
' WHY THIS EXISTS
'   Each issue and currency of a netting set has exactly one position.
'
' INPUTS
'   Ns
'     The netting set.
'
'   Kind
'     PK_ISSUE or PK_CURRENCY.
'
'   Key
'     The issue key or the currency.
'
'   Asset, Ccy
'     Stored on a new position.
'
' RETURNS
'   The index in mPos.
'
' BEHAVIOR
'   Looks the position up; creates it when new, growing the list as
'   needed.
'
' ERROR POLICY
'   None raised by the procedure itself.
'
' DEPENDENCIES
'   CORE_Util for KeyIndex and KeyAdd.
'
' NOTES
'   Positions are keyed by netting set, kind and key.
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Engine.PositionFor"

    Dim LookupKey       As String   'Netting set, kind and key
    Dim PositionIndex   As Long     'Position

'------------------------------------------------------------------------------
' FIND OR CREATE
'------------------------------------------------------------------------------
    'Look the position up
        LookupKey = Ns & "|" & Kind & "|" & Key
        PositionIndex = KeyIndex(mPosIndex, LookupKey)
    'Create it when new, growing the list as needed
        If PositionIndex = 0 Then
            mPosCount = mPosCount + 1
            If mPosCount = 1 Then
                ReDim mPos(1 To 64)
            ElseIf mPosCount > UBound(mPos) Then
                ReDim Preserve mPos(1 To mPosCount * 2)
            End If
            PositionIndex = mPosCount
            KeyAdd mPosIndex, LookupKey, PositionIndex
            mPos(PositionIndex).NsIndex = Ns
            mPos(PositionIndex).Kind = Kind
            mPos(PositionIndex).Key = Key
            mPos(PositionIndex).Asset = Asset
            mPos(PositionIndex).Ccy = Ccy
        End If
    'Return its index
        PositionFor = PositionIndex

End Function

Private Sub AddLegValue( _
    ByVal p As Long, _
    ByVal i As Long)
'
'------------------------------------------------------------------------------
'                        ADD A LEG VALUE TO A POSITION
'------------------------------------------------------------------------------
' PURPOSE
'   Add a leg's value to the lent or received side of a position.
'
' INPUTS
'   p
'     The position.
'
'   i
'     The leg.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
    'Count the leg and add its value to its side
        With mPos(p)
            .Legs = .Legs + 1
            If mLegs(i).Sign > 0 Then
                .Lent = .Lent + mLegs(i).Value
            Else
                .Received = .Received + mLegs(i).Value
            End If
        End With

End Sub

Private Sub AddToTransaction( _
    ByVal t As Long, _
    ByVal i As Long)
'
'------------------------------------------------------------------------------
'                         ADD A LEG TO ITS TRANSACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Add a recognised leg of a netting set without netting to its
'   transaction: the value lent or received and its adjusted value.
'
' INPUTS
'   t
'     The transaction.
'
'   i
'     The leg.
'
' NOTES
'   The currency term of a leg received is its value times its Hfx.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
    'Count the leg in its netting set and add its values to its side
        With mTx(t)
            mNS(.NsIndex).LegsUsed = mNS(.NsIndex).LegsUsed + 1
            If mLegs(i).Sign > 0 Then
                .E = .E + mLegs(i).Value
                .EAdjusted = .EAdjusted + mLegs(i).Adjusted
            Else
                .C = .C + mLegs(i).Value
                .CAdjusted = .CAdjusted + mLegs(i).Adjusted
                .FXTerm = .FXTerm + mLegs(i).Value * mLegs(i).Hfx
            End If
        End With

End Sub

Private Sub FinishNetted()

'
'------------------------------------------------------------------------------
'                          CALCULATE E* WITH NETTING
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate E* of every VALID netting set with an eligible master netting
'   agreement from its positions: the net position and haircut of each
'   security issue, the net and gross exposures, N, the FX term over the
'   currencies other than the settlement currency, and E*.
'
' WHY THIS EXISTS
'   The positions of all netting sets sit in one list; each netting set is
'   calculated from its own slice of it.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Returns when there is no position
'   - Groups the positions by netting set
'   - Calculates each VALID netting set with netting that has positions
'     through FinishOneNetted
'
' ERROR POLICY
'   Errors of FinishOneNetted are not handled here.
'
' DEPENDENCIES
'   FinishOneNetted.
'
' NOTES
'   None
'
' METHOD
'   The positions of all netting sets are grouped by netting set with a
'   counting sort, so each netting set is calculated from its own slice.
'
' REFERENCE
'   CRE22.62-22.65; CRR Art. 220.
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Engine.FinishNetted"

    Dim Starts()        As Long 'First slot of each netting set in Order
    Dim Filled()        As Long 'Slots filled so far per netting set
    Dim Order()         As Long 'Position indices grouped by netting set
    Dim Ns              As Long 'Netting set
    Dim PositionIndex   As Long 'Position or slot

'------------------------------------------------------------------------------
' GROUP THE POSITIONS BY NETTING SET
'------------------------------------------------------------------------------
    'Nothing to calculate without positions
        If mPosCount = 0 Then
            Exit Sub
        End If
    'Allocate the slots
        ReDim Starts(1 To mNsCount + 1)
        ReDim Filled(1 To mNsCount)
        ReDim Order(1 To mPosCount)
    'Count the positions of each netting set
        For PositionIndex = 1 To mPosCount
            Filled(mPos(PositionIndex).NsIndex) = Filled(mPos(PositionIndex).NsIndex) + 1
        Next PositionIndex
    'Set the first slot of each netting set
        Starts(1) = 1
        For Ns = 1 To mNsCount
            Starts(Ns + 1) = Starts(Ns) + Filled(Ns)
            Filled(Ns) = 0
        Next Ns
    'Place each position in the slice of its netting set
        For PositionIndex = 1 To mPosCount
            Ns = mPos(PositionIndex).NsIndex
            Order(Starts(Ns) + Filled(Ns)) = PositionIndex
            Filled(Ns) = Filled(Ns) + 1
        Next PositionIndex

'------------------------------------------------------------------------------
' CALCULATE EACH NETTING SET
'------------------------------------------------------------------------------
    'Calculate each VALID netting set with netting that has positions
        For Ns = 1 To mNsCount
            If mNS(Ns).MNA And mNS(Ns).Status = "VALID" And Filled(Ns) > 0 Then
                FinishOneNetted Ns, Order, Starts(Ns), Starts(Ns) + Filled(Ns) - 1
            End If
        Next Ns

End Sub

Private Sub FinishOneNetted( _
    ByVal Ns As Long, _
    ByRef Order() As Long, _
    ByVal FirstSlot As Long, _
    ByVal LastSlot As Long)
'
'------------------------------------------------------------------------------
'                        CALCULATE E* OF ONE NETTED SET
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate E* of one netting set with netting from its positions,
'   Order(FirstSlot) to Order(LastSlot).
'
' WHY THIS EXISTS
'   Keeps the netting formula of one netting set apart from the grouping
'   done by FinishNetted.
'
' INPUTS
'   Ns
'     The netting set.
'
'   Order, FirstSlot, LastSlot
'     The position indices grouped by netting set, and the slice of this
'     one.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Collects the net position and haircut of each issue and marks the
'     issues counted in N (non-zero and at least the threshold times the
'     largest)
'   - Sets the currency haircut of each currency position and adds the FX
'     term of the currencies other than the settlement currency
'   - Calculates the net and gross exposures, N, the systematic and
'     idiosyncratic terms and E*
'
' ERROR POLICY
'   Raises ERR_FORMULA_DOMAIN through Checked when an SFT_ function
'   returns an error value.
'
' DEPENDENCIES
'   Checked, SFT_ScaledHaircut, SFT_NetExposure, SFT_GrossExposure,
'   SFT_IssueCount, SFT_ExposureNetted.
'
' NOTES
'   A 0% haircut netting set takes no currency haircut.
'
' REFERENCE
'   CRE22.64-22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Engine.FinishOneNetted"

    Dim Es()            As Double   'Net position of each issue
    Dim Hs()            As Double   'Haircut of each issue
    Dim IssuePos()      As Long     'Position index of each issue
    Dim Issues          As Long     'Issues
    Dim Hfx             As Double   'Scaled currency haircut
    Dim Largest         As Double   'Largest |E_s|
    Dim SlotIndex       As Long     'Slot or issue
    Dim PositionIndex   As Long     'Position

'------------------------------------------------------------------------------
' ISSUES
'------------------------------------------------------------------------------
    'Allocate one entry per slot
        ReDim Es(1 To LastSlot - FirstSlot + 1)
        ReDim Hs(1 To LastSlot - FirstSlot + 1)
        ReDim IssuePos(1 To LastSlot - FirstSlot + 1)
    'Collect the net position and haircut of each issue and the largest |E_s|
        For SlotIndex = FirstSlot To LastSlot
            PositionIndex = Order(SlotIndex)
            If mPos(PositionIndex).Kind = PK_ISSUE Then
                Issues = Issues + 1
                Es(Issues) = mPos(PositionIndex).Lent - mPos(PositionIndex).Received
                Hs(Issues) = mPos(PositionIndex).H
                IssuePos(Issues) = PositionIndex
                If Abs(Es(Issues)) > Largest Then
                    Largest = Abs(Es(Issues))
                End If
            End If
        Next SlotIndex
    'Mark the issues counted in N
        For SlotIndex = 1 To Issues
            mPos(IssuePos(SlotIndex)).Counted = (Es(SlotIndex) <> 0# And _
                Abs(Es(SlotIndex)) >= pThreshold * Largest)
        Next SlotIndex

'------------------------------------------------------------------------------
' CURRENCIES
'------------------------------------------------------------------------------
    'Work on the netting set
        With mNS(Ns)
        'Scale the currency haircut unless the 0% haircut applies
            If Not .ZeroHaircut Then
                Hfx = Checked(SFT_ScaledHaircut(pHfx, .HoldingBD, .RemarginBD, pHcBase), "SFT_ScaledHaircut")
            End If
        'Set the haircut of each currency position; add the FX term of the
        'currencies other than the settlement currency
            For SlotIndex = FirstSlot To LastSlot
                PositionIndex = Order(SlotIndex)
                If mPos(PositionIndex).Kind = PK_CURRENCY Then
                    mPos(PositionIndex).H = Hfx
                    If mPos(PositionIndex).Ccy <> .SettleCcy Then
                        .FXTerm = .FXTerm + Abs(mPos(PositionIndex).Lent - mPos(PositionIndex).Received) * Hfx
                    End If
                End If
            Next SlotIndex

'------------------------------------------------------------------------------
' EXPOSURE VALUE
'------------------------------------------------------------------------------
        'Calculate the net and gross exposures and N over the issues
            If Issues > 0 Then
                ReDim Preserve Es(1 To Issues)
                ReDim Preserve Hs(1 To Issues)
                .NetExposure = Checked(SFT_NetExposure(Es, Hs), "SFT_NetExposure")
                .GrossExposure = Checked(SFT_GrossExposure(Es, Hs), "SFT_GrossExposure")
                .IssueCount = VBA.CLng(Checked(SFT_IssueCount(Es, pThreshold), "SFT_IssueCount"))
            End If
        'Calculate the systematic and idiosyncratic terms
            .SystematicTerm = pNetW * .NetExposure
            If .IssueCount > 0 Then
                .IdiosyncraticTerm = pGrossW * .GrossExposure / Sqr(.IssueCount)
            End If
        'Calculate E*
            .EStar = Checked(SFT_ExposureNetted(.SumE, .SumC, .NetExposure, .GrossExposure, .IssueCount, _
                .FXTerm, pNetW, pGrossW), "SFT_ExposureNetted")
        End With

End Sub

Private Sub FinishSingles()

'
'------------------------------------------------------------------------------
'                         CALCULATE E* WITHOUT NETTING
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate E* of every transaction of a VALID netting set without an
'   eligible master netting agreement with the single-transaction formula,
'   and add it to its netting set: E* of the netting set is the sum over
'   its transactions, and its holding period the longest of theirs.
'
' WHY THIS EXISTS
'   Without netting, E* is calculated per transaction and summed per
'   netting set.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Derives the basket haircuts of each transaction from its sums
'   - Calculates its E* with SFT_ExposureSingle
'   - Adds its sums and E* to its netting set and keeps the longest holding
'     period and its haircut scale
'
' ERROR POLICY
'   Raises ERR_FORMULA_DOMAIN through Checked when SFT_ExposureSingle
'   returns an error value.
'
' DEPENDENCIES
'   Checked, SFT_ExposureSingle; CORE_Util for Max2.
'
' NOTES
'   Negative basket haircuts are passed as 0.
'
' METHOD
'   The basket haircuts passed to SFT_ExposureSingle are the value-weighted
'   averages of the legs' haircuts [CRE22.55], so E(1 + He) and
'   C(1 - Hc - Hfx) equal the sums over the legs already kept in EAdjusted
'   and CAdjusted.
'
' REFERENCE
'   CRE22.46, 22.55; CRR Art. 223.
'
' UPDATED
'   2026-10-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "CORE_Engine.FinishSingles"

    Dim HeAvg               As Double   'Value-weighted haircut of the legs lent
    Dim HcAvg               As Double   'Value-weighted haircut of the legs received
    Dim HfxAvg              As Double   'Value-weighted currency haircut of the legs received
    Dim Ns                  As Long     'Netting set
    Dim TransactionIndex    As Long     'Transaction

'------------------------------------------------------------------------------
' EACH TRANSACTION
'------------------------------------------------------------------------------
    'Calculate each transaction of a VALID netting set without netting
        For TransactionIndex = 1 To mTxCount
            Ns = mTx(TransactionIndex).NsIndex
            If Ns > 0 Then
                If Not mNS(Ns).MNA And mNS(Ns).Status = "VALID" Then
                    With mTx(TransactionIndex)
                        HeAvg = 0#
                        HcAvg = 0#
                        HfxAvg = 0#
                        If .E > 0# Then
                            HeAvg = (.EAdjusted - .E) / .E
                        End If
                        If .C > 0# Then
                            HfxAvg = .FXTerm / .C
                            HcAvg = (.C - .CAdjusted - .FXTerm) / .C
                        End If
                        .EStar = Checked(SFT_ExposureSingle(.E, Max2(HeAvg, 0#), .C, Max2(HcAvg, 0#), _
                            Max2(HfxAvg, 0#)), "SFT_ExposureSingle")
                        mNS(Ns).SumE = mNS(Ns).SumE + .E
                        mNS(Ns).SumC = mNS(Ns).SumC + .C
                        mNS(Ns).FXTerm = mNS(Ns).FXTerm + .FXTerm
                        mNS(Ns).EStar = mNS(Ns).EStar + .EStar
                        If .HoldingBD > mNS(Ns).HoldingBD Then
                            mNS(Ns).HoldingBD = .HoldingBD
                            mNS(Ns).HaircutScale = Sqr((mNS(Ns).RemarginBD + .HoldingBD - 1#) / pHcBase)
                        End If
                    End With
                End If
            End If
        Next TransactionIndex

End Sub

Private Function Checked( _
    ByVal Result As Variant, _
    ByVal FunctionName As String) As Double
'
'------------------------------------------------------------------------------
'                            CHECK A FORMULA RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Return the number an SFT_* function gave the engine. The engine checks
'   every input before calling them, so an error value here means a defect:
'   the run stops with ERR_FORMULA_DOMAIN rather than writing a wrong
'   figure.
'
' ERROR POLICY
'   Raises ERR_FORMULA_DOMAIN for an error value.
'
' NOTES
'   The error source is the literal "CORE_Engine.Checked".
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Stop the run on an error value
        If IsError(Result) Then
            Err.Raise ERR_FORMULA_DOMAIN, "CORE_Engine.Checked", _
                FunctionName & " rejected the inputs the " & "engine passed: " & VBA.CStr(Result)
        End If
    'Return the number
        Checked = VBA.CDbl(Result)

End Function

'
'------------------------------------------------------------------------------
'
'                                   OUTPUTS
'
'------------------------------------------------------------------------------
'

Private Sub ClearOutputSheets()

'
'------------------------------------------------------------------------------
'                           CLEAR THE OUTPUT TABLES
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the Results, Positions and LegCalc tables before a run. Checks is
'   cleared and rewritten by WriteChecks.
'
' WHY THIS EXISTS
'   A failed or validation-only run must never leave earlier results
'   looking current.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the data rows of each output table whose sheet and header name
'   can be located.
'
' ERROR POLICY
'   A table whose sheet or header name is missing is skipped, so that
'   ValidateSchema reports it on Checks and the run stops there.
'
' DEPENDENCIES
'   SheetPresent; CORE_Util for TableFirstRow, GetSheet and
'   ClearOutputBlock.
'
' NOTES
'   None
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Engine.ClearOutputSheets"

    Dim SheetNames      As Variant  'Output sheets cleared
    Dim ColumnCounts    As Variant  'Their table widths
    Dim FirstRow        As Long     'First data row; 0 when the table cannot be located
    Dim ItemIndex       As Long     'Position in the lists

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'List the output tables and their widths
        SheetNames = Array(SH_RESULTS, SH_POSITIONS, SH_LEGCALC)
        ColumnCounts = Array(RS_NCOLS, PS_NCOLS, LC_NCOLS)
    'Clear each table that can be located
        For ItemIndex = LBound(SheetNames) To UBound(SheetNames)
            If SheetPresent(VBA.CStr(SheetNames(ItemIndex))) Then
                FirstRow = TableFirstRow(VBA.CStr(SheetNames(ItemIndex)))
                If FirstRow > 0 Then
                    ClearOutputBlock GetSheet(VBA.CStr(SheetNames(ItemIndex))), FirstRow, _
                        VBA.CLng(ColumnCounts(ItemIndex))
                End If
            End If
        Next ItemIndex

End Sub

Private Sub WriteResults()

'
'------------------------------------------------------------------------------
'                           WRITE THE RESULTS SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Write one Results row per netting set, then the TOTAL CRR row over
'   the VALID netting sets. A TOTAL BCBS row is written only when a
'   netting set is in that regime, which the refusal of BCBS (#248) rules
'   out. A netting set that is not VALID shows its identity and status
'   only.
'
' WHY THIS EXISTS
'   Results is the main output: one line per netting set with every figure
'   from E* to the own funds requirement.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Fills one row per netting set; the figures only when it is VALID, and
'     the netting terms only with netting
'   - Adds the TOTAL rows
'   - Writes the block, formats its columns and sets the TOTAL rows in bold
'
' ERROR POLICY
'   Errors are not handled here; they reach the handler of Calculate.
'
' DEPENDENCIES
'   CORE_Util for GetSheet, OutputFirstRow, WriteBlock, FormatColumns and
'   TableRange; CORE_Config for RS_NCOLS.
'
' NOTES
'   The column numbers are those of the Results table.
'
' UPDATED
'   2026-10-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.WriteResults"

    Dim Sheet       As Worksheet    'Results sheet
    Dim TopRow      As Long         'First data row
    Dim Block()     As Variant      'Output rows
    Dim RowCount    As Long         'Rows written
    Dim Ns          As Long         'Netting set
    Dim RegimeSlot  As Long         'Regime index
    Dim TotalRows   As Long         'TOTAL rows written

'------------------------------------------------------------------------------
' NETTING SETS
'------------------------------------------------------------------------------
    'Locate the table and allocate a row per netting set and per total
        Set Sheet = GetSheet(SH_RESULTS)
        TopRow = OutputFirstRow(SH_RESULTS)
        ReDim Block(1 To mNsCount + 2, 1 To RS_NCOLS)
    'Fill one row per netting set; the figures only when it is VALID
        For Ns = 1 To mNsCount
            RowCount = RowCount + 1
            With mNS(Ns)
                Block(RowCount, 1) = .ID
                Block(RowCount, 2) = .Counterparty
                Block(RowCount, 3) = .Regime
                Block(RowCount, 4) = IIf(.MNA, "Y", "N")
                Block(RowCount, 19) = .CptyClass
                Block(RowCount, 24) = .Status
                If .Status = "VALID" Then
                    Block(RowCount, 5) = .TxCount
                    Block(RowCount, 6) = .LegsUsed
                    Block(RowCount, 7) = .HoldingBD
                    Block(RowCount, 8) = .RemarginBD
                    Block(RowCount, 9) = .HaircutScale
                    Block(RowCount, 10) = .SumE
                    Block(RowCount, 11) = .SumC
                    If .MNA Then
                        Block(RowCount, 12) = .NetExposure
                        Block(RowCount, 13) = .GrossExposure
                        Block(RowCount, 14) = .IssueCount
                        Block(RowCount, 15) = .SystematicTerm
                        Block(RowCount, 16) = .IdiosyncraticTerm
                    End If
                    Block(RowCount, 17) = .FXTerm
                    Block(RowCount, 18) = .EStar
                    Block(RowCount, 20) = .RWKey
                    Block(RowCount, 21) = .RiskWeight
                    Block(RowCount, 22) = .RWA
                    Block(RowCount, 23) = .OwnFunds
                End If
            End With
        Next Ns

'------------------------------------------------------------------------------
' TOTALS
'------------------------------------------------------------------------------
    'Add the CRR total, and a BCBS total only when a netting set is in it
        For RegimeSlot = 1 To 2
            If RegimeSlot = 1 Or mRegSets(RegimeSlot) > 0 Then
                RowCount = RowCount + 1
                TotalRows = TotalRows + 1
                Block(RowCount, 1) = "TOTAL " & Choose(RegimeSlot, RG_CRR, RG_BCBS)
                Block(RowCount, 3) = Choose(RegimeSlot, RG_CRR, RG_BCBS)
                Block(RowCount, 18) = mRegEStar(RegimeSlot)
                Block(RowCount, 22) = mRegRWA(RegimeSlot)
                Block(RowCount, 23) = mRegOwnFunds(RegimeSlot)
                Block(RowCount, 24) = mRegValid(RegimeSlot) & " of " & mRegSets(RegimeSlot) & _
                    " netting set(s) VALID; " & (mRegSets(RegimeSlot) - mRegValid(RegimeSlot)) & " withheld"
            End If
        Next RegimeSlot

'------------------------------------------------------------------------------
' WRITE AND FORMAT
'------------------------------------------------------------------------------
    'Write the rows and format the columns
        WriteBlock Sheet, TopRow, Block, RowCount, RS_NCOLS
        FormatColumns Sheet, TopRow, RowCount, "@|@|@|@|0|0|0|0|0.0000|#,##0|#,##0|#,##0|#,##0|0|#,##0|#,##0|" & _
            "#,##0|#,##0|@|@|0.00%|#,##0|#,##0|@"
    'Set the TOTAL rows in bold
        TableRange(Sheet, TopRow + RowCount - TotalRows, 1, TopRow + RowCount - 1, RS_NCOLS).Font.Bold = True

End Sub

Private Sub WritePositions()

'
'------------------------------------------------------------------------------
'                          WRITE THE POSITIONS SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Write the Positions sheet for every VALID netting set: with netting, one
'   row per security issue and per currency; without netting, one row per
'   transaction, filed under its first leg's netting set. A netting set
'   that is not VALID has no rows, as on Results it has no figures.
'
' WHY THIS EXISTS
'   Positions shows the intermediate figures behind E* of each netting set.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - With netting: one row per position, with the net position, haircut,
'     net and gross terms and whether the issue counts in N, or the role of
'     the currency (SETTLEMENT or MISMATCH) and its FX term
'   - Without netting: one row per transaction not rejected, with its sums
'     and E*
'   - Writes and formats the rows; returns when there are none
'
' ERROR POLICY
'   Errors are not handled here; they reach the handler of Calculate.
'
' DEPENDENCIES
'   CORE_Util for GetSheet, OutputFirstRow, WriteBlock and FormatColumns;
'   CORE_Config for PS_NCOLS and the position kinds.
'
' NOTES
'   Positions exist only for VALID netting sets, so the netting branch
'   needs no status test.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "CORE_Engine.WritePositions"

    Dim Sheet               As Worksheet    'Positions sheet
    Dim TopRow              As Long         'First data row
    Dim Block()             As Variant      'Output rows
    Dim RowCount            As Long         'Rows written
    Dim Ns                  As Long         'Netting set
    Dim PositionIndex       As Long         'Position
    Dim TransactionIndex    As Long         'Transaction
    Dim NetPos              As Double       'Lent less received

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
    'Locate the table and allocate a row per position and per transaction
        Set Sheet = GetSheet(SH_POSITIONS)
        TopRow = OutputFirstRow(SH_POSITIONS)
        ReDim Block(1 To mPosCount + mTxCount + 1, 1 To PS_NCOLS)
    'Fill the rows of each netting set: positions with netting, transactions without
        For Ns = 1 To mNsCount
            If mNS(Ns).MNA Then
                For PositionIndex = 1 To mPosCount
                    If mPos(PositionIndex).NsIndex = Ns Then
                        RowCount = RowCount + 1
                        With mPos(PositionIndex)
                            NetPos = .Lent - .Received
                            Block(RowCount, 1) = mNS(Ns).ID
                            Block(RowCount, 2) = .Kind
                            Block(RowCount, 3) = .Key
                            Block(RowCount, 4) = .Asset
                            Block(RowCount, 5) = .Ccy
                            Block(RowCount, 6) = .Legs
                            Block(RowCount, 7) = mNS(Ns).HoldingBD
                            Block(RowCount, 8) = .Lent
                            Block(RowCount, 9) = .Received
                            Block(RowCount, 10) = NetPos
                            Block(RowCount, 11) = .H
                            If .Kind = PK_ISSUE Then
                                Block(RowCount, 12) = NetPos * .H
                                Block(RowCount, 13) = Abs(NetPos) * .H
                                Block(RowCount, 14) = IIf(.Counted, "Y", "N")
                            ElseIf .Ccy = mNS(Ns).SettleCcy Then
                                Block(RowCount, 14) = "SETTLEMENT"
                            Else
                                Block(RowCount, 14) = "MISMATCH"
                                Block(RowCount, 15) = Abs(NetPos) * .H
                            End If
                        End With
                    End If
                Next PositionIndex
            ElseIf mNS(Ns).Status = "VALID" Then
                For TransactionIndex = 1 To mTxCount
                    If mTx(TransactionIndex).NsIndex = Ns And Not mTx(TransactionIndex).Rejected Then
                        RowCount = RowCount + 1
                        With mTx(TransactionIndex)
                            Block(RowCount, 1) = mNS(Ns).ID
                            Block(RowCount, 2) = PK_TRANSACTION
                            Block(RowCount, 3) = .ID
                            Block(RowCount, 4) = .TxType
                            Block(RowCount, 5) = .LentCcy
                            Block(RowCount, 6) = .LegCount
                            Block(RowCount, 7) = .HoldingBD
                            Block(RowCount, 8) = .E
                            Block(RowCount, 9) = .C
                            Block(RowCount, 10) = .E - .C
                            Block(RowCount, 15) = .FXTerm
                            Block(RowCount, 16) = .EAdjusted
                            Block(RowCount, 17) = .CAdjusted
                            Block(RowCount, 18) = .EStar
                        End With
                    End If
                Next TransactionIndex
            End If
        Next Ns

'------------------------------------------------------------------------------
' WRITE AND FORMAT
'------------------------------------------------------------------------------
    'Nothing to write without rows
        If RowCount = 0 Then
            Exit Sub
        End If
    'Write the rows and format the columns
        WriteBlock Sheet, TopRow, Block, RowCount, PS_NCOLS
        FormatColumns Sheet, TopRow, RowCount, "@|@|@|@|@|0|0|#,##0|#,##0|#,##0|0.0000%|#,##0|#,##0|@|#,##0|" & _
            "#,##0|#,##0|#,##0"

End Sub

Private Sub WriteLegCalc()

'
'------------------------------------------------------------------------------
'                           WRITE THE LEGCALC SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Write one LegCalc row per leg, in input order, with every intermediate
'   figure of its haircut; an excluded leg shows why.
'
' WHY THIS EXISTS
'   LegCalc lets every haircut be traced back to its inputs.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Returns when there are no legs
'   - Fills one row per leg: its identity and messages; EXCLUDED for an
'     invalid leg or one of a rejected transaction; otherwise its status,
'     haircut key, eligibility, values and, for a recognised leg with a
'     holding period, every figure of its haircut
'   - Writes and formats the rows
'
' ERROR POLICY
'   Errors are not handled here; they reach the handler of Calculate.
'
' DEPENDENCIES
'   CORE_Util for GetSheet, OutputFirstRow, WriteBlock and FormatColumns;
'   CORE_Config for LC_NCOLS and the leg statuses.
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
    Const PROC_NAME As String = "CORE_Engine.WriteLegCalc"

    Dim Sheet       As Worksheet    'LegCalc sheet
    Dim TopRow      As Long         'First data row
    Dim Block()     As Variant      'Output rows
    Dim LegIndex    As Long         'Leg

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
    'Nothing to write without legs
        If mLegCount = 0 Then
            Exit Sub
        End If
    'Locate the table and allocate a row per leg
        Set Sheet = GetSheet(SH_LEGCALC)
        TopRow = OutputFirstRow(SH_LEGCALC)
        ReDim Block(1 To mLegCount, 1 To LC_NCOLS)
    'Fill one row per leg
        For LegIndex = 1 To mLegCount
            With mLegs(LegIndex)
                Block(LegIndex, 1) = .LegID
                Block(LegIndex, 2) = .TxID
                If .NsIndex > 0 Then
                    Block(LegIndex, 3) = mNS(.NsIndex).ID
                    Block(LegIndex, 19) = mNS(.NsIndex).Regime
                End If
                Block(LegIndex, 5) = .TxType
                Block(LegIndex, 6) = .Side
                Block(LegIndex, 7) = .Asset
                Block(LegIndex, 20) = .Messages
                If Not .Valid Or .TxRejected Then
                    Block(LegIndex, 4) = LS_EXCLUDED
                Else
                    If .Recognised Then
                        Block(LegIndex, 4) = LS_OK
                    Else
                        Block(LegIndex, 4) = LS_NOTRECOG
                    End If
                    Block(LegIndex, 8) = .HaircutKey
                    Block(LegIndex, 9) = IIf(.Eligible, "Y", "N")
                    If .ResidualYears >= 0# Then
                        Block(LegIndex, 10) = .ResidualYears
                    End If
                    Block(LegIndex, 11) = .Value
                    Block(LegIndex, 12) = .Sign * .Value
                    If .Recognised And .HoldingBD > 0# Then
                        Block(LegIndex, 13) = .H10
                        Block(LegIndex, 14) = .HoldingBD
                        Block(LegIndex, 15) = .Scale
                        Block(LegIndex, 16) = .H
                        Block(LegIndex, 17) = .Hfx
                        Block(LegIndex, 18) = .Adjusted
                    End If
                End If
            End With
        Next LegIndex

'------------------------------------------------------------------------------
' WRITE AND FORMAT
'------------------------------------------------------------------------------
    'Write the rows and format the columns
        WriteBlock Sheet, TopRow, Block, mLegCount, LC_NCOLS
        FormatColumns Sheet, TopRow, mLegCount, "@|@|@|@|@|@|@|@|@|0.00|#,##0|#,##0|0.00%|0|0.0000|0.0000%|" & _
            "0.0000%|#,##0|@|@"

End Sub

Private Function WriteChecks( _
    ByRef Failure As String) As Boolean
'
'------------------------------------------------------------------------------
'                            WRITE THE CHECKS SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Write the logged messages to the Checks sheet, errors in bold, or a
'   single "No issues found." line.
'
' WHY THIS EXISTS
'   Invalid inputs are reported on Checks rather than raised; Calculate
'   must know whether that report could be written.
'
' INPUTS
'   Failure
'     Set to the reason when writing failed; "" otherwise.
'
' RETURNS
'   True when the sheet was written. False when writing failed, with the
'   reason in Failure.
'
' BEHAVIOR
'   - Clears the Checks table
'   - Writes "No issues found." when nothing was logged
'   - Otherwise writes every message and sets the ERROR lines in bold
'
' ERROR POLICY
'   Contains any error and reports it through the result, never silently.
'
' DEPENDENCIES
'   CORE_Util for GetSheet, OutputFirstRow, ClearOutputBlock, WriteBlock
'   and TableCell.
'
' NOTES
'   The log is stored (field, message) and transposed for the sheet.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "CORE_Engine.WriteChecks"

    Dim ChecksSheet     As Worksheet    'Checks sheet
    Dim TopRow          As Long         'First Checks data row
    Dim OutputBlock()   As Variant      'Output rows, (message, field)
    Dim MessageIndex    As Long         'Message
    Dim FieldIndex      As Long         'Field

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling and start without a failure
        On Error GoTo Failed
        Failure = vbNullString

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Locate and clear the Checks table
        Set ChecksSheet = GetSheet(SH_CHECKS)
        TopRow = OutputFirstRow(SH_CHECKS)
        ClearOutputBlock ChecksSheet, TopRow, CK_NCOLS
    'Write a single line when nothing was logged
        If mLogCount = 0 Then
            ReDim OutputBlock(1 To 1, 1 To CK_NCOLS)
            OutputBlock(1, 1) = SEV_INFO
            OutputBlock(1, 5) = "No issues found."
            WriteBlock ChecksSheet, TopRow, OutputBlock, 1, CK_NCOLS
            WriteChecks = True
            Exit Function
        End If
    'Transpose the log into rows and write them
        ReDim OutputBlock(1 To mLogCount, 1 To CK_NCOLS)
        For MessageIndex = 1 To mLogCount
            For FieldIndex = 1 To CK_NCOLS
                OutputBlock(MessageIndex, FieldIndex) = mLog(FieldIndex, MessageIndex)
            Next FieldIndex
        Next MessageIndex
        WriteBlock ChecksSheet, TopRow, OutputBlock, mLogCount, CK_NCOLS
    'Set the ERROR lines in bold
        For MessageIndex = 1 To mLogCount
            If OutputBlock(MessageIndex, 1) = SEV_ERROR Then
                TableCell(ChecksSheet, TopRow + MessageIndex - 1, 1).Font.Bold = True
            End If
        Next MessageIndex
    'Report success
        WriteChecks = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Failed:
    'Report the failure through the result
        Failure = "error " & Err.Number & ": " & Err.Description

End Function

Private Sub WriteRunInfo( _
    ByVal RunSeconds As Double, _
    ByVal Completed As Boolean, _
    ByVal WroteOutputs As Boolean, _
    ByVal Failure As String)
'
'------------------------------------------------------------------------------
'                            WRITE THE RUN SUMMARY
'------------------------------------------------------------------------------
' PURPOSE
'   Write a one-line run summary to the Results cell named NM_RUNINFO: a
'   failure, a validation-only run, a completed run with its counts, run
'   class and duration, or a run stopped by input errors.
'
' WHY THIS EXISTS
'   The summary tells at a glance whether the figures on the sheets are
'   current and complete.
'
' INPUTS
'   RunSeconds
'     Duration of the run, in seconds.
'
'   Completed
'     True when the run completed.
'
'   WroteOutputs
'     False for a validation-only run.
'
'   Failure
'     Why the run failed; "" when it did not.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Composes the first summary that applies (failure, validation, completed
'   run, run stopped by input errors) and writes it to NM_RUNINFO.
'
' ERROR POLICY
'   Never raises: the summary is informative. A protected or missing
'   Results sheet, or a missing name, keeps its earlier summary.
'
' DEPENDENCIES
'   RunClassText; CORE_Util for NamedCell.
'
' NOTES
'   The write is the only statement under On Error Resume Next.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "CORE_Engine.WriteRunInfo"

    Dim Summary     As String   'Summary line
    Dim Stamp       As String   'Date and time of the run

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
    'Stamp the run
        Stamp = Format$(Now, "yyyy-mm-dd hh:mm:ss")
    'Compose the first summary that applies
        If VBA.Len(Failure) > 0 Then
            Summary = "Last run " & Stamp & " FAILED - " & Failure
        ElseIf Not WroteOutputs Then
            Summary = "Last validation " & Stamp & " | errors " & mErrCount & ", warnings " & mWarnCount & _
                RunClassText() & " | outputs cleared: press Run SFT to calculate"
        ElseIf Completed Then
            Summary = "Last run " & Stamp & " | reporting date " & Format$(VBA.CDate(pAsOf), "yyyy-mm-dd") & _
                " | ccy " & pRepCcy & " | transactions used " & mTxUsed & " of " & mTxCount & _
                ", legs " & mLegsUsed & " of " & mLegCount & " | errors " & mErrCount & _
                ", warnings " & mWarnCount & RunClassText() & " | " & Format$(RunSeconds, "0.00") & " s"
            If mWithheld > 0 Then
                Summary = Summary & " | exposure withheld for " & mWithheld & " netting set(s)"
            End If
        Else
            Summary = "Last run " & Stamp & " FAILED - see Checks sheet (" & mErrCount & " errors)"
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Write the summary, best effort
        On Error Resume Next
        NamedCell(NM_RUNINFO).Value = Summary
        On Error GoTo 0

End Sub

Private Function RunClassText() As String

'
'------------------------------------------------------------------------------
'                            DESCRIBE THE RUN CLASS
'------------------------------------------------------------------------------
' PURPOSE
'   The run class for the run summary: the regulatory set's identity, or
'   WHAT-IF with the number of differences; "" when not classified.
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Describe the class of the run
        If mRunClass = RUN_REGULATORY Then
            RunClassText = " | regulatory parameters " & REF_SET_ID
        ElseIf mRunClass = RUN_WHATIF Then
            RunClassText = " | WHAT-IF: " & mDeviations & " difference(s) from " & REF_SET_ID
        End If

End Function
