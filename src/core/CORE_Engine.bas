Attribute VB_Name = "CORE_Engine"
'==============================================================================
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
' REGIME LAYER
'   The regime is set on Params (CRR by default) and can be overridden per
'   netting set. With the CRR consolidated text of 26.06.2026 the haircuts,
'   holding periods, netting formula and risk weights used here are the
'   same in both regimes; the regime selects the rows of the haircut and
'   risk-weight tables whose Regimes column is BOTH or that regime, and
'   keeps the totals apart. The differences that remain are documented on
'   the Basel_vs_CRR sheet.
'
' FLOW
'   Calculate runs, in order:
'     ClearOutputSheets, ValidateSchema: clear old results, check the layout;
'     LoadParams, LoadFXTable, LoadHaircutTable, LoadRiskWeightTable,
'     LoadNettingSets: read the inputs;
'     ClassifyRun: regulatory or what-if run;
'     LoadLegs: one row per leg, validated, valued and keyed to a haircut;
'     BuildTransactions: legs grouped into transactions, which are checked
'     as a whole and rejected as a whole;
'     CheckNettingSets: transaction types, short-term status, zero-haircut
'     conditions, issue consistency, holding period override, risk-weight
'     key and status of each netting set;
'     ComputeNettingSets: holding period, haircuts, positions and E*;
'     WriteResults, WritePositions, WriteLegCalc, WriteChecks: outputs;
'     WriteRunInfo: the run summary.
'   Each netting set, transaction, leg, issue and currency position is
'   stored once in an array, with a Collection (CORE_Util.KeyIndex) mapping
'   its key to its array position.
'
' PUBLIC SURFACE
'   None outside this VBA project. Calculate, TotalsText, RunClass and the
'   run counters are Public for M_Main; Option Private Module keeps them off the supported
'   external surface.
'
' DEPENDENCIES
'   CORE_Config for the layout, CORE_Util for conversions and sheet access,
'   SFT_Formulas for the holding period, haircut scaling, both exposure
'   formulas, the net and gross exposures, the issue count and the own
'   funds requirement.
'
' STATE OWNERSHIP
'   Owns the module state below. ResetState clears it at the start of each
'   run; the run counters stay readable until the next run. Reads the input
'   sheets and Params; writes Results, Positions, LegCalc and Checks.
'   Changes no Excel application setting.
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
' KNOWN DEVIATION
'   This module lives in src/core but reads and writes worksheets, which
'   the repository structure reserves for src/workbook, calls SFT_Formulas
'   in src/modules, and keeps module state (MODULE STATE below), which the
'   house style excludes from src/core.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' UPDATED
'   2026-10-26
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the engine inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' DATA STRUCTURES
'------------------------------------------------------------------------------
    'One netting set, as read from the NettingSets sheet, with its results.
    'Amounts are in the reporting currency.
    Private Type tNettingSet
        ID                As String     'Netting-set ID as entered
        Row               As Long       'Sheet row
        Counterparty      As String     'Counterparty name
        MNA               As Boolean    'Covered by an eligible master netting agreement
        SettleCcy         As String     'Settlement currency of the agreement
        CptyClass         As String     'SOV, INST, CORP or QCCP
        CQS               As Long       'Credit quality step 1-6; 0 unrated
        SCRA              As String     'SCRA grade of an unrated institution
        RemarginBD        As Double     'Remargining frequency N_R, business days
        LargeOrIlliquid   As Boolean    'Over 5,000 trades or illiquid collateral
        Disputes          As Boolean    'Margin-call disputes: holding period doubled
        HoldingOverride   As Double     'Holding period override, business days; 0 none
        ZeroHaircut       As Boolean    'Conditions of the 0% haircut declared met
        HasRWOverride     As Boolean    'A risk weight override is given
        RWOverride        As Double     'Risk weight override, decimal
        Regime            As String     'BCBS or CRR
        InputErrors       As Long       'Errors that make the netting set INVALID
        TxCount           As Long       'Transactions read for this netting set
        TxRejected        As Long       'Transactions rejected
        LegsUsed          As Long       'Legs recognised in the calculation
        HasRepoStyle      As Boolean    'Holds a repo-style transaction
        HasCapitalMarket  As Boolean    'Holds an other capital-market-driven transaction
        AllShortTerm      As Boolean    'Every transaction of 3 months or less
        HoldingBD         As Double     'Holding period T_M; the longest without netting
        HaircutScale      As Double     'sqrt((N_R + T_M - 1) / base); the longest without netting
        SumE              As Double     'Value lent, sold or posted
        SumC              As Double     'Value borrowed, bought or received
        NetExposure       As Double     '|sum E_s H_s|
        GrossExposure     As Double     'sum |E_s| H_s
        IssueCount        As Long       'N
        SystematicTerm    As Double     'NetWeight x net exposure
        IdiosyncraticTerm As Double     'GrossWeight x gross exposure / sqrt(N)
        FXTerm            As Double     'Currency-mismatch term
        EStar             As Double     'Exposure value E*
        RWKey             As String     'Risk-weight key
        RiskWeight        As Double     'Risk weight applied
        RWA               As Double     'Risk-weighted exposure amount
        OwnFunds          As Double     'Own funds requirement
        Status            As String     'VALID, INVALID, INCOMPLETE or NO TRANSACTIONS
    End Type

    'One leg: cash or one security issue lent or received in a transaction.
    Private Type tLeg
        Row               As Long       'Sheet row
        LegID             As String     'Leg ID as entered
        TxKey             As String     'Transaction ID, upper case
        TxID              As String     'Transaction ID as entered
        TxIndex           As Long       'Index in mTx
        TxRejected        As Boolean    'Its transaction was rejected
        NsIndex           As Long       'Index in mNS; 0 when unknown
        TxType            As String     'Transaction type
        StartSerial       As Double     'Start date serial; -1 invalid
        MatSerial         As Double     'Maturity date serial; 0 open
        Side              As String     'LENT or RECEIVED
        Sign              As Long       '+1 lent, -1 received
        Asset             As String     'Asset type
        IssueKey          As String     'Issue ID, upper case; CASH and the currency for cash
        Issuer            As String     'Issuer type of a debt security
        CQS               As Long       'Credit quality step of the issue; 0 unrated
        SecMatSerial      As Double     'Security maturity date serial
        Listing           As String     'Equity listing
        FundH10           As Double     'Fund haircut H10
        MarketValue       As Double     'Market value in the leg currency
        Ccy               As String     'Leg currency
        Value             As Double     'Market value in the reporting currency
        ResidualYears     As Double     'Residual maturity of a debt security; -1 n/a
        HaircutKey        As String     'Key of the haircut-table row; FUND for a fund
        Eligible          As Boolean    'Eligible as collateral
        Recognised        As Boolean    'Used in the calculation
        H10               As Double     'Haircut before scaling
        Signature         As String     'Description of the issue, compared across its legs
        Valid             As Boolean    'No input error on this leg
        HoldingBD         As Double     'Holding period applied
        Scale             As Double     'Haircut scale applied
        H                 As Double     'Scaled haircut
        Hfx               As Double     'Currency haircut, single-transaction formula only
        Adjusted          As Double     'V(1 + H) lent, V(1 - H - Hfx) received
        Messages          As String     'Warnings, or why the leg was excluded
    End Type

    'One transaction: the legs sharing a Transaction ID.
    Private Type tTransaction
        ID                As String     'Transaction ID as entered
        NsIndex           As Long       'Netting set of its first leg; 0 unknown
        TxType            As String     'Transaction type of its first leg
        StartSerial       As Double     'Start date of its first leg
        MatSerial         As Double     'Maturity of its first leg; 0 open
        FirstRow          As Long       'Sheet row of its first leg
        LegCount          As Long       'Legs read
        LentCount         As Long       'Legs lent
        ReceivedCount     As Long       'Legs received
        LentNonCash       As Long       'Non-cash legs lent
        ReceivedNonCash   As Long       'Non-cash legs received
        LentCcy           As String     'Currency of the legs lent; "*" when several
        Rejected          As Boolean    'Rejected as a whole
        ShortTerm         As Boolean    'Original maturity of 3 months or less
        HoldingBD         As Double     'Holding period, without netting
        E                 As Double     'Value lent
        C                 As Double     'Value received and recognised
        EAdjusted         As Double     'sum V (1 + H) lent
        CAdjusted         As Double     'sum V (1 - H - Hfx) received
        FXTerm            As Double     'sum V Hfx received
        EStar             As Double     'Exposure value without netting
    End Type

    'One security issue or currency of a netting set with netting.
    Private Type tPosition
        NsIndex           As Long       'Netting set
        Kind              As String     'ISSUE or CURRENCY
        Key               As String     'Issue ID or currency, upper case
        Asset             As String     'Asset type; blank for a currency
        Ccy               As String     'Currency
        Legs              As Long       'Legs aggregated
        Lent              As Double     'Value lent
        Received          As Double     'Value received
        H                 As Double     'Scaled haircut, or the currency haircut
        Signature         As String     'Not used
        Counted           As Boolean    'Counted in N
    End Type

'------------------------------------------------------------------------------
' MODULE STATE
'------------------------------------------------------------------------------
    'Parameters of the run.
        Private pAsOf          As Double     'Reporting date serial
        Private pRepCcy        As String     'Reporting currency
        Private pRegime        As String     'Default regime
        Private pDaysYear      As Double     'Days per year
        Private pHcBase        As Double     'Holding period of the H10 table
        Private pHoldRepo      As Double     'Minimum holding period, repo-style
        Private pHoldCMT       As Double     'Minimum holding period, other capital market
        Private pHoldLarge     As Double     'Holding period, large or illiquid
        Private pHfx           As Double     'Currency haircut H10
        Private pNonElig       As Double     'Haircut H10 of non-eligible securities lent
        Private pNetW          As Double     'Net weight
        Private pGrossW        As Double     'Gross weight
        Private pThreshold     As Double     'Issue-count threshold
        Private pCapRatio      As Double     'Own funds per unit of RWA
    'Lookup tables of Params.
        Private mHcKeys()      As String     'Haircut keys
        Private mHcValues()    As Double     'H10
        Private mHcRegimes()   As String     'BOTH, CRR or BCBS
        Private mHcCount       As Long       'Rows
        Private mHcIndex       As Collection 'Key to row
        Private mRwKeys()      As String     'Risk-weight keys
        Private mRwValues()    As Double     'Risk weights
        Private mRwRegimes()   As String     'BOTH, CRR or BCBS
        Private mRwCount       As Long       'Rows
        Private mRwIndex       As Collection 'Key to row
        Private mFxCcy()       As String     'Currencies
        Private mFxRate()      As Double     'Units of reporting currency per unit
        Private mFxCount       As Long       'Rows
        Private mFxIndex       As Collection 'Currency to row
    'Inputs and intermediate results.
        Private mNS()          As tNettingSet
        Private mNsCount       As Long
        Private mNsIndex       As Collection
        Private mLegs()        As tLeg
        Private mLegCount      As Long
        Private mLegIndex      As Collection
        Private mTx()          As tTransaction
        Private mTxCount       As Long
        Private mTxIndex       As Collection
        Private mPos()         As tPosition
        Private mPosCount      As Long
        Private mPosIndex      As Collection
    'Checks messages, stored (field, message) so that ReDim Preserve can grow
    'the last dimension, and the run counters.
        Private mLog()         As Variant
        Private mLogCount      As Long
        Private mErrCount      As Long
        Private mWarnCount     As Long
        Private mRunClass      As String     'RUN_REGULATORY, RUN_WHATIF or ""
        Private mDeviations    As Long       'Differences from the regulatory set
        Private mTxUsed        As Long       'Transactions in VALID netting sets
        Private mLegsUsed      As Long       'Legs recognised in VALID netting sets
        Private mWithheld      As Long       'Netting sets whose E* is withheld
        Private mRegEStar(1 To 2)    As Double    'Total E* per regime: 1 CRR, 2 BCBS
        Private mRegRWA(1 To 2)      As Double    'Total RWA per regime
        Private mRegOwnFunds(1 To 2) As Double    'Total own funds requirement per regime
        Private mRegSets(1 To 2)     As Long      'Netting sets per regime
        Private mRegValid(1 To 2)    As Long      'VALID netting sets per regime


'
'------------------------------------------------------------------------------
'
'                                     RUN
'
'------------------------------------------------------------------------------
'

Public Function Calculate( _
    ByVal WriteOutputs As Boolean) _
    As Boolean
'
'==============================================================================
'                                  Calculate
'------------------------------------------------------------------------------
' PURPOSE
'   Run the engine once: load the inputs, calculate every leg, transaction
'   and netting set, and write the outputs.
'
' INPUTS
'   WriteOutputs: True writes every output sheet; False validates only: it
'   clears the output tables and writes the Checks sheet and the run
'   summary.
'
' RETURNS
'   True when the run completed. Individual transactions may still have
'   been rejected; the Checks sheet lists them. False when an input sheet,
'   input table or Params layout is not as expected, or a parameter, a
'   Params table or the netting sets could not be loaded. A missing output
'   sheet or output header name raises from ClearOutputSheets before the
'   layout is checked.
'
' ERROR POLICY
'   Raises ERR_CHECKS_WRITE when the Checks sheet cannot be written, after
'   clearing the other output sheets. Any other error clears the output
'   sheets, is logged on Checks and in the run summary, and is raised with
'   its original number, source and description.
'
' STATE OWNERSHIP
'   Resets all module state, then fills it for this run.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim StartTime        As Double    'Timer value at the start
    Dim ChecksFailure    As String    'Why the Checks sheet could not be written
    Dim ErrNumber        As Long      'Unexpected error: number
    Dim ErrSource        As String    'Unexpected error: source
    Dim ErrDescription   As String    'Unexpected error: description

'------------------------------------------------------------------------------
' LOAD INPUTS
'------------------------------------------------------------------------------
    'Old outputs are cleared first, for validation too, so that a failed or
    'validation-only run can never leave earlier results looking current.
        StartTime = Timer
        ResetState
        On Error GoTo Failed
        ClearOutputSheets
        If Not ValidateSchema() Then GoTo Finish
        If Not LoadParams() Then GoTo Finish
        If Not LoadFXTable() Then GoTo Finish
        If Not LoadHaircutTable() Then GoTo Finish
        If Not LoadRiskWeightTable() Then GoTo Finish
        If Not LoadNettingSets() Then GoTo Finish
        ClassifyRun

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        LoadLegs
        BuildTransactions
        CheckNettingSets
        ComputeNettingSets
        If WriteOutputs Then
            WriteResults
            WritePositions
            WriteLegCalc
        End If
        Calculate = True

'------------------------------------------------------------------------------
' WRITE CHECKS AND RUN SUMMARY
'------------------------------------------------------------------------------
Finish:
        If Not WriteChecks(ChecksFailure) Then
            Calculate = False
            ClearOutputSheets
            WriteRunInfo Timer - StartTime, False, WriteOutputs, "the Checks sheet could not be written (" & _
                         ChecksFailure & "); results withdrawn"
            On Error GoTo 0
            Err.Raise ERR_CHECKS_WRITE, "CORE_Engine.Calculate", "The Checks sheet could not be written (" & _
                      ChecksFailure & "); results withdrawn. Unprotect or repair the Checks sheet and run again."
        End If
        WriteRunInfo Timer - StartTime, Calculate, WriteOutputs, ""
        Exit Function

'------------------------------------------------------------------------------
' HANDLE UNEXPECTED ERROR
'------------------------------------------------------------------------------
    'Withdraw everything written so far and log the original error, both
    'best effort, then raise the original error unchanged.
Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Calculate = False
        On Error Resume Next
        ClearOutputSheets
        LogMsg SEV_ERROR, "", "", "Run stopped by error " & ErrNumber & ": " & ErrDescription & _
               " - results withdrawn."
        WriteChecks ChecksFailure
        WriteRunInfo Timer - StartTime, False, WriteOutputs, "error " & ErrNumber & ": " & ErrDescription
        On Error GoTo 0
        Err.Raise ErrNumber, ErrSource, ErrDescription

End Function


Public Property Get ErrorCount() As Long
'
'==============================================================================
'                                  ErrorCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number of ERROR messages logged by the last run.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        ErrorCount = mErrCount

End Property


Public Property Get WarningCount() As Long
'
'==============================================================================
'                                 WarningCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number of WARNING messages logged by the last run.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        WarningCount = mWarnCount

End Property


Public Property Get RunClass() As String
'
'==============================================================================
'                                   RunClass
'------------------------------------------------------------------------------
' PURPOSE
'   REGULATORY or WHAT-IF for the last run; "" when it was not classified.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        RunClass = mRunClass

End Property


Public Function TotalsText( _
    ByVal RegimeCode As String) _
    As String
'
'==============================================================================
'                                  TotalsText
'------------------------------------------------------------------------------
' PURPOSE
'   One line with the totals of one regime of the last run, for the message
'   M_Main shows: E*, RWA and own funds requirement over the VALID netting
'   sets, and how many netting sets were withheld. Totals of different
'   regimes are never added together.
'
' INPUTS
'   RegimeCode: RG_CRR or RG_BCBS.
'
' RETURNS
'   The line; "" when the run had no netting set of that regime.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Regime index

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
        r = RegimeIndex(RegimeCode)
        If mRegSets(r) = 0 Then
            Exit Function
        End If
        TotalsText = RegimeCode & ": E* " & Format$(mRegEStar(r), "#,##0") & _
                     ", RWA " & Format$(mRegRWA(r), "#,##0") & _
                     ", own funds " & Format$(mRegOwnFunds(r), "#,##0") & _
                     " (" & mRegValid(r) & " of " & mRegSets(r) & " netting set(s) VALID)"

End Function


Private Function RegimeIndex( _
    ByVal RegimeCode As String) _
    As Long
'
'==============================================================================
'                                 RegimeIndex
'------------------------------------------------------------------------------
' PURPOSE
'   Index of a regime in the per-regime totals: 1 for CRR, 2 for BCBS.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
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
'==============================================================================
'                                  ResetState
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the module state before a run.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim r   As Long    'Regime index

'------------------------------------------------------------------------------
' RESET
'------------------------------------------------------------------------------
        ReDim mLog(1 To CK_NCOLS, 1 To 64)
        mLogCount = 0
        mErrCount = 0
        mWarnCount = 0
        mRunClass = ""
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
        Set mHcIndex = New Collection
        Set mRwIndex = New Collection
        Set mFxIndex = New Collection
        Set mNsIndex = New Collection
        Set mLegIndex = New Collection
        Set mTxIndex = New Collection
        Set mPosIndex = New Collection
        For r = 1 To 2
            mRegEStar(r) = 0#
            mRegRWA(r) = 0#
            mRegOwnFunds(r) = 0#
            mRegSets(r) = 0
            mRegValid(r) = 0
        Next r

End Sub


Private Sub LogMsg( _
    ByVal Severity As String, _
    ByVal SheetName As String, _
    ByVal Reference As String, _
    ByVal Message As String, _
    Optional ByVal RowNumber As Long = 0)
'
'==============================================================================
'                                    LogMsg
'------------------------------------------------------------------------------
' PURPOSE
'   Add one message for the Checks sheet and count errors and warnings.
'
' INPUTS
'   Severity: SEV_ERROR, SEV_WARN or SEV_INFO.
'   SheetName: sheet the message refers to.
'   Reference: the ID or parameter code concerned; may be empty.
'   Message: the message text.
'   RowNumber: sheet row concerned; 0 leaves the Row column blank.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' STORE MESSAGE
'------------------------------------------------------------------------------
        mLogCount = mLogCount + 1
        If mLogCount > UBound(mLog, 2) Then
            ReDim Preserve mLog(1 To CK_NCOLS, 1 To mLogCount * 2)
        End If
        mLog(1, mLogCount) = Severity
        mLog(2, mLogCount) = SheetName
        If RowNumber > 0 Then
            mLog(3, mLogCount) = RowNumber
        Else
            mLog(3, mLogCount) = Empty
        End If
        mLog(4, mLogCount) = Reference
        mLog(5, mLogCount) = Message

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
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
'==============================================================================
'                                   LegError
'------------------------------------------------------------------------------
' PURPOSE
'   Reject a leg: log an ERROR on Transactions and keep the reason for its
'   LegCalc row. Its transaction is then rejected as a whole.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
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
'==============================================================================
'                                   LegNote
'------------------------------------------------------------------------------
' PURPOSE
'   Log a WARNING or INFO about a leg that stays in the calculation, and
'   keep it for its LegCalc row.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        With mLegs(LegIndex)
            .Messages = AppendText(.Messages, Message)
            LogMsg Severity, SH_TX, .LegID, Message, .Row
        End With

End Sub


Private Sub NsError( _
    ByVal NsIndex As Long, _
    ByVal Message As String)
'
'==============================================================================
'                                   NsError
'------------------------------------------------------------------------------
' PURPOSE
'   Log an ERROR that makes a netting set INVALID, so that its E* is
'   withheld.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        With mNS(NsIndex)
            .InputErrors = .InputErrors + 1
            LogMsg SEV_ERROR, SH_NS, .ID, Message & " - exposure withheld.", .Row
        End With

End Sub


Private Function AppendText( _
    ByVal Existing As String, _
    ByVal Addition As String) _
    As String
'
'==============================================================================
'                                  AppendText
'------------------------------------------------------------------------------
' PURPOSE
'   Join two message texts with "; ".
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If Len(Existing) = 0 Then
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
'==============================================================================
'                                ValidateSchema
'------------------------------------------------------------------------------
' PURPOSE
'   Check the workbook layout before anything is read: every sheet, every
'   workbook name that locates a table or a parameter, and the header text
'   of every column the engine reads. The engine reads by column number, so
'   a header that differs means a column was inserted, deleted or moved.
'
' RETURNS
'   True when the layout is as expected; otherwise False with the problems
'   found logged (missing sheets first; then the first header mismatch of
'   each table), and the run stops.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SheetNames   As Variant    'Sheets the engine needs
    Dim Codes        As Variant    'Parameter codes
    Dim Problems     As Long       'Problems found
    Dim i            As Long       'Position in a list

'------------------------------------------------------------------------------
' SHEETS
'------------------------------------------------------------------------------
        SheetNames = Array(SH_PARAMS, SH_NS, SH_TX, SH_RESULTS, SH_POSITIONS, SH_LEGCALC, SH_CHECKS)
        For i = LBound(SheetNames) To UBound(SheetNames)
            If Not SheetPresent(CStr(SheetNames(i))) Then
                LogMsg SEV_ERROR, CStr(SheetNames(i)), "", "Sheet missing - rebuild the workbook from the template."
                Problems = Problems + 1
            End If
        Next i
        If Problems > 0 Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' TABLE HEADERS
'------------------------------------------------------------------------------
        Problems = Problems + HeaderProblems(SH_NS, TableHeaderRow(SH_NS), NS_HEADERS, NM_NS_HEADER)
        Problems = Problems + HeaderProblems(SH_TX, TableHeaderRow(SH_TX), TX_HEADERS, NM_TX_HEADER)
        Problems = Problems + NamedHeaderProblems(NM_HC_TABLE, HC_HEADERS)
        Problems = Problems + NamedHeaderProblems(NM_RW_TABLE, RW_HEADERS)
        Problems = Problems + NamedHeaderProblems(NM_FX_TABLE, FX_HEADERS)
        Problems = Problems + NamedHeaderProblems(NM_PRM_CODE, PRM_CODE_HEADER)
        Problems = Problems + NamedHeaderProblems(NM_PRM_VALUE, PRM_VALUE_HEADER)
        SheetNames = Array(SH_RESULTS, SH_POSITIONS, SH_LEGCALC, SH_CHECKS)
        For i = LBound(SheetNames) To UBound(SheetNames)
            If TableHeaderRow(CStr(SheetNames(i))) = 0 Then
                LogMsg SEV_ERROR, CStr(SheetNames(i)), "", "Workbook name '" & TableHeaderName(CStr(SheetNames(i))) & _
                       "' is missing or no longer refers to the first header cell in column C."
                Problems = Problems + 1
            End If
        Next i

'------------------------------------------------------------------------------
' PARAMETER NAMES
'------------------------------------------------------------------------------
        Codes = ParameterCodes()
        For i = LBound(Codes) To UBound(Codes)
            If NamedCell(CStr(Codes(i))) Is Nothing Then
                LogMsg SEV_ERROR, SH_PARAMS, CStr(Codes(i)), "Workbook name of parameter missing or broken."
                Problems = Problems + 1
            End If
        Next i
        ValidateSchema = (Problems = 0)

End Function


Private Function ParameterCodes() As Variant
'
'==============================================================================
'                                ParameterCodes
'------------------------------------------------------------------------------
' PURPOSE
'   The codes of every parameter the engine reads, in Params order.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        ParameterCodes = Array(PRM_ASOF, PRM_REPCCY, PRM_REGIME, PRM_DAYSYEAR, PRM_HCBASE, PRM_HOLD_REPO, _
                               PRM_HOLD_CMT, PRM_HOLD_LARGE, PRM_HFX, PRM_NONELIG, PRM_NETW, PRM_GROSSW, _
                               PRM_NTHRESH, PRM_CAPRATIO)

End Function


Private Function SheetPresent( _
    ByVal SheetName As String) _
    As Boolean
'
'==============================================================================
'                                 SheetPresent
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether this workbook has a worksheet with the given tab name.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet   As Worksheet    'Each worksheet

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
        For Each Sheet In ThisWorkbook.Worksheets
            If StrComp(Sheet.Name, SheetName, vbTextCompare) = 0 Then
                SheetPresent = True
                Exit Function
            End If
        Next Sheet

End Function


Private Function HeaderProblems( _
    ByVal SheetName As String, _
    ByVal HeaderRow As Long, _
    ByVal Expected As String, _
    ByVal HeaderName As String) _
    As Long
'
'==============================================================================
'                                HeaderProblems
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the header row of an input table with the expected headers,
'   ignoring case, spaces and line breaks; an empty expected entry is a
'   column that is not read.
'
' RETURNS
'   0 when they match; 1 with the first difference logged otherwise.
'
' UPDATED
'   2026-10-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parts()   As String       'Expected headers, 0-based
    Dim Sheet     As Worksheet    'The input sheet
    Dim Found     As String       'Header in the sheet
    Dim Col       As Long         'Position in Parts; column Col + 1

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If HeaderRow = 0 Then
            LogMsg SEV_ERROR, SheetName, "", "Workbook name '" & HeaderName & _
                   "' is missing or no longer refers to the first header cell in column C."
            HeaderProblems = 1
            Exit Function
        End If
        Set Sheet = GetSheet(SheetName)
        Parts = Split(Expected, "|")
        For Col = 0 To UBound(Parts)
            If Len(Parts(Col)) > 0 Then
                Found = SafeStr(TableCell(Sheet, HeaderRow, Col + 1).Value)
                If HeaderKey(Found) <> HeaderKey(Parts(Col)) Then
                    LogMsg SEV_ERROR, SheetName, "", "Column " & (Col + 1) & " header is '" & Found & _
                           "', expected '" & Parts(Col) & "': a column was inserted, deleted or moved.", HeaderRow
                    HeaderProblems = 1
                    Exit Function
                End If
            End If
        Next Col

End Function


Private Function NamedHeaderProblems( _
    ByVal NameText As String, _
    ByVal Expected As String) _
    As Long
'
'==============================================================================
'                             NamedHeaderProblems
'------------------------------------------------------------------------------
' PURPOSE
'   Compare the cells from a workbook name's top-left cell rightwards with
'   the expected headers of a Params table or parameter column.
'
' RETURNS
'   0 when they match; 1 with the problem logged otherwise.
'
' UPDATED
'   2026-10-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cell      As Range        'The name's top-left cell
    Dim Parts()   As String       'Expected headers, 0-based
    Dim Found     As String       'Header in the sheet
    Dim Col       As Long         'Position in Parts

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        Set Cell = NamedCell(NameText)
        If Cell Is Nothing Then
            LogMsg SEV_ERROR, SH_PARAMS, NameText, "Workbook name missing or broken - rebuild from the template."
            NamedHeaderProblems = 1
            Exit Function
        End If
        Parts = Split(Expected, "|")
        For Col = 0 To UBound(Parts)
            If Len(Parts(Col)) > 0 Then
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
    ByVal Header As String) _
    As String
'
'==============================================================================
'                                  HeaderKey
'------------------------------------------------------------------------------
' PURPOSE
'   Normalise a header for comparison: upper case, without spaces, tabs or
'   line breaks.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        HeaderKey = UCase$(Replace(Replace(Replace(Replace(Header, " ", ""), vbTab, ""), vbCr, ""), vbLf, ""))

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
'==============================================================================
'                                  LoadParams
'------------------------------------------------------------------------------
' PURPOSE
'   Read and check every parameter on Params.
'
' RETURNS
'   True when every parameter is usable; otherwise False with each problem
'   logged, and the run stops.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Problems   As Long       'Parameters rejected

'------------------------------------------------------------------------------
' REPORTING DATE, CURRENCY AND REGIME
'------------------------------------------------------------------------------
        pAsOf = ToSerial(GetParam(PRM_ASOF))
        If pAsOf <= 0# Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_ASOF, "Reporting date missing or not a date."
            Problems = Problems + 1
        End If
        pRepCcy = UTxt(GetParam(PRM_REPCCY))
        If Len(pRepCcy) <> 3 Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REPCCY, "Reporting currency must be a 3-letter code."
            Problems = Problems + 1
        End If
        pRegime = UTxt(GetParam(PRM_REGIME))
        If pRegime <> RG_CRR And pRegime <> RG_BCBS Then
            LogMsg SEV_ERROR, SH_PARAMS, PRM_REGIME, "Regime must be CRR or BCBS."
            Problems = Problems + 1
        End If

'------------------------------------------------------------------------------
' NUMBERS
'------------------------------------------------------------------------------
        Problems = Problems + NumParam(PRM_DAYSYEAR, 360#, 366#, pDaysYear)
        Problems = Problems + NumParam(PRM_HCBASE, 1#, 250#, pHcBase)
        Problems = Problems + NumParam(PRM_HOLD_REPO, 1#, 250#, pHoldRepo)
        Problems = Problems + NumParam(PRM_HOLD_CMT, 1#, 250#, pHoldCMT)
        Problems = Problems + NumParam(PRM_HOLD_LARGE, 1#, 250#, pHoldLarge)
        Problems = Problems + NumParam(PRM_HFX, 0#, 1#, pHfx)
        Problems = Problems + NumParam(PRM_NONELIG, 0#, 1#, pNonElig)
        Problems = Problems + NumParam(PRM_NETW, 0#, 1#, pNetW)
        Problems = Problems + NumParam(PRM_GROSSW, 0#, 1#, pGrossW)
        Problems = Problems + NumParam(PRM_NTHRESH, 0#, 1#, pThreshold)
        Problems = Problems + NumParam(PRM_CAPRATIO, 0#, 1#, pCapRatio)
        LoadParams = (Problems = 0)

End Function


Private Function NumParam( _
    ByVal Code As String, _
    ByVal Lowest As Double, _
    ByVal Highest As Double, _
    ByRef Result As Double) _
    As Long
'
'==============================================================================
'                                   NumParam
'------------------------------------------------------------------------------
' PURPOSE
'   Read one numeric parameter and check it lies in [Lowest, Highest].
'
' RETURNS
'   0 when usable, with Result set; 1 with the problem logged otherwise.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Content   As Variant    'Raw value

'------------------------------------------------------------------------------
' READ AND CHECK
'------------------------------------------------------------------------------
        Content = GetParam(Code)
        If Not IsNum(Content) Then
            LogMsg SEV_ERROR, SH_PARAMS, Code, "Parameter missing or not a number."
            NumParam = 1
            Exit Function
        End If
        Result = CDbl(Content)
        If Result < Lowest Or Result > Highest Then
            LogMsg SEV_ERROR, SH_PARAMS, Code, "Parameter " & Result & " outside [" & Lowest & ", " & Highest & "]."
            NumParam = 1
        End If

End Function


Private Function LoadFXTable() As Boolean
'
'==============================================================================
'                                 LoadFXTable
'------------------------------------------------------------------------------
' PURPOSE
'   Read the FX table below its header down to the first blank currency:
'   units of reporting currency per unit of each currency. A currency cell
'   holding a worksheet error value is reported, not read as blank.
'
' RETURNS
'   True when every row is usable and the reporting currency is listed at
'   1; otherwise False with each problem logged, and the run stops.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header     As Range     'Header cell FX_Ccy
    Dim Ccy        As String    'Currency of the row
    Dim Rate       As Variant   'Rate of the row
    Dim Problems   As Long      'Rows rejected
    Dim r          As Long      'Row offset below the header
    Dim Position   As Long      'Index of the reporting currency

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set Header = NamedCell(NM_FX_TABLE)
        ReDim mFxCcy(1 To 16)
        ReDim mFxRate(1 To 16)
        r = 1
        Do While Not IsBlankCell(Header.Offset(r, 0).Value) Or IsError(Header.Offset(r, 0).Value)
            Ccy = UTxt(Header.Offset(r, 0).Value)
            Rate = Header.Offset(r, 2).Value2
            If IsError(Header.Offset(r, 0).Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, "", "FX currency holds a worksheet error value.", Header.Row + r
                Problems = Problems + 1
            ElseIf Len(Ccy) <> 3 Then
                LogMsg SEV_ERROR, SH_PARAMS, Ccy, "FX currency must be a 3-letter code.", Header.Row + r
                Problems = Problems + 1
            ElseIf KeyIndex(mFxIndex, Ccy) > 0 Then
                LogMsg SEV_ERROR, SH_PARAMS, Ccy, "Currency listed twice in the FX table.", Header.Row + r
                Problems = Problems + 1
            ElseIf Not IsNum(Rate) Then
                LogMsg SEV_ERROR, SH_PARAMS, Ccy, "FX rate missing or not a number.", Header.Row + r
                Problems = Problems + 1
            ElseIf CDbl(Rate) <= 0# Then
                LogMsg SEV_ERROR, SH_PARAMS, Ccy, "FX rate must be above 0.", Header.Row + r
                Problems = Problems + 1
            Else
                mFxCount = mFxCount + 1
                If mFxCount > UBound(mFxCcy) Then
                    ReDim Preserve mFxCcy(1 To mFxCount * 2)
                    ReDim Preserve mFxRate(1 To mFxCount * 2)
                End If
                mFxCcy(mFxCount) = Ccy
                mFxRate(mFxCount) = CDbl(Rate)
                KeyAdd mFxIndex, Ccy, mFxCount
            End If
            r = r + 1
        Loop

'------------------------------------------------------------------------------
' CHECK THE REPORTING CURRENCY
'------------------------------------------------------------------------------
        Position = KeyIndex(mFxIndex, pRepCcy)
        If Position = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, pRepCcy, "Reporting currency not in the FX table."
            Problems = Problems + 1
        ElseIf mFxRate(Position) <> 1# Then
            LogMsg SEV_ERROR, SH_PARAMS, pRepCcy, "Reporting currency must have the FX rate 1."
            Problems = Problems + 1
        End If
        LoadFXTable = (Problems = 0)

End Function


Private Function LoadHaircutTable() As Boolean
'
'==============================================================================
'                               LoadHaircutTable
'------------------------------------------------------------------------------
' PURPOSE
'   Read the supervisory haircut table below its header down to the first
'   blank key: key, H10 and the regimes each row applies to. A key that is
'   not in the table makes the asset not eligible. A key cell holding a
'   worksheet error value is reported, not read as blank.
'
' RETURNS
'   True when every row is usable; otherwise False and the run stops.
'
' REFERENCE
'   CRE22.49; CRR Art. 224 Tables 1 and 3.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header     As Range      'Header cell HC_Key
    Dim Key        As String     'Key of the row
    Dim Value      As Variant    'H10 of the row
    Dim Regimes    As String     'Regimes of the row
    Dim Problems   As Long       'Rows rejected
    Dim r          As Long       'Row offset below the header

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set Header = NamedCell(NM_HC_TABLE)
        ReDim mHcKeys(1 To 64)
        ReDim mHcValues(1 To 64)
        ReDim mHcRegimes(1 To 64)
        r = 1
        Do While Not IsBlankCell(Header.Offset(r, HC_KEY - 1).Value) Or IsError(Header.Offset(r, HC_KEY - 1).Value)
            Key = UTxt(Header.Offset(r, HC_KEY - 1).Value)
            Value = Header.Offset(r, HC_H10 - 1).Value2
            Regimes = UTxt(Header.Offset(r, HC_REGIMES - 1).Value)
            If IsError(Header.Offset(r, HC_KEY - 1).Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, "", "Haircut key holds a worksheet error value.", Header.Row + r
                Problems = Problems + 1
            ElseIf KeyIndex(mHcIndex, Key) > 0 Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Haircut key listed twice.", Header.Row + r
                Problems = Problems + 1
            ElseIf Not IsNum(Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Haircut missing or not a number.", Header.Row + r
                Problems = Problems + 1
            ElseIf CDbl(Value) < 0# Or CDbl(Value) > 1# Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Haircut must be between 0 and 1.", Header.Row + r
                Problems = Problems + 1
            ElseIf Regimes <> RG_BOTH And Regimes <> RG_CRR And Regimes <> RG_BCBS Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Regimes must be BOTH, CRR or BCBS.", Header.Row + r
                Problems = Problems + 1
            Else
                mHcCount = mHcCount + 1
                If mHcCount > UBound(mHcKeys) Then
                    ReDim Preserve mHcKeys(1 To mHcCount * 2)
                    ReDim Preserve mHcValues(1 To mHcCount * 2)
                    ReDim Preserve mHcRegimes(1 To mHcCount * 2)
                End If
                mHcKeys(mHcCount) = Key
                mHcValues(mHcCount) = CDbl(Value)
                mHcRegimes(mHcCount) = Regimes
                KeyAdd mHcIndex, Key, mHcCount
            End If
            r = r + 1
        Loop
        If mHcCount = 0 And Problems = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, NM_HC_TABLE, "Haircut table is empty."
            Problems = 1
        End If
        LoadHaircutTable = (Problems = 0)

End Function


Private Function LoadRiskWeightTable() As Boolean
'
'==============================================================================
'                             LoadRiskWeightTable
'------------------------------------------------------------------------------
' PURPOSE
'   Read the risk-weight table below its header down to the first blank
'   key: key, risk weight and the regimes each row applies to. A key cell
'   holding a worksheet error value is reported, not read as blank.
'
' RETURNS
'   True when every row is usable; otherwise False and the run stops.
'
' REFERENCE
'   CRE20.7, 20.18-20.21, CRE54.7; CRR Art. 114, 120-122, 306(1)(a).
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Header     As Range      'Header cell RW_Key
    Dim Key        As String     'Key of the row
    Dim Value      As Variant    'Risk weight of the row
    Dim Regimes    As String     'Regimes of the row
    Dim Problems   As Long       'Rows rejected
    Dim r          As Long       'Row offset below the header

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set Header = NamedCell(NM_RW_TABLE)
        ReDim mRwKeys(1 To 64)
        ReDim mRwValues(1 To 64)
        ReDim mRwRegimes(1 To 64)
        r = 1
        Do While Not IsBlankCell(Header.Offset(r, RW_KEY - 1).Value) Or IsError(Header.Offset(r, RW_KEY - 1).Value)
            Key = UTxt(Header.Offset(r, RW_KEY - 1).Value)
            Value = Header.Offset(r, RW_WEIGHT - 1).Value2
            Regimes = UTxt(Header.Offset(r, RW_REGIMES - 1).Value)
            If IsError(Header.Offset(r, RW_KEY - 1).Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, "", "Risk-weight key holds a worksheet error value.", Header.Row + r
                Problems = Problems + 1
            ElseIf KeyIndex(mRwIndex, Key) > 0 Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Risk-weight key listed twice.", Header.Row + r
                Problems = Problems + 1
            ElseIf Not IsNum(Value) Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Risk weight missing or not a number.", Header.Row + r
                Problems = Problems + 1
            ElseIf CDbl(Value) < 0# Or CDbl(Value) > 12.5 Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Risk weight must be between 0 and 12.5.", Header.Row + r
                Problems = Problems + 1
            ElseIf Regimes <> RG_BOTH And Regimes <> RG_CRR And Regimes <> RG_BCBS Then
                LogMsg SEV_ERROR, SH_PARAMS, Key, "Regimes must be BOTH, CRR or BCBS.", Header.Row + r
                Problems = Problems + 1
            Else
                mRwCount = mRwCount + 1
                If mRwCount > UBound(mRwKeys) Then
                    ReDim Preserve mRwKeys(1 To mRwCount * 2)
                    ReDim Preserve mRwValues(1 To mRwCount * 2)
                    ReDim Preserve mRwRegimes(1 To mRwCount * 2)
                End If
                mRwKeys(mRwCount) = Key
                mRwValues(mRwCount) = CDbl(Value)
                mRwRegimes(mRwCount) = Regimes
                KeyAdd mRwIndex, Key, mRwCount
            End If
            r = r + 1
        Loop
        If mRwCount = 0 And Problems = 0 Then
            LogMsg SEV_ERROR, SH_PARAMS, NM_RW_TABLE, "Risk-weight table is empty."
            Problems = 1
        End If
        LoadRiskWeightTable = (Problems = 0)

End Function


Private Function FXRate( _
    ByVal Ccy As String) _
    As Double
'
'==============================================================================
'                                    FXRate
'------------------------------------------------------------------------------
' PURPOSE
'   Units of reporting currency per unit of a currency; 0 when the currency
'   is not in the FX table.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Position   As Long    'Row in the FX table

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        Position = KeyIndex(mFxIndex, Ccy)
        If Position > 0 Then
            FXRate = mFxRate(Position)
        End If

End Function


Private Function TableRow( _
    ByVal Lookup As Collection, _
    ByRef Regimes() As String, _
    ByVal Key As String, _
    ByVal RegimeCode As String) _
    As Long
'
'==============================================================================
'                                   TableRow
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
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Position   As Long    'Row of the key

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
        If Len(Key) = 0 Then
            Exit Function
        End If
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
'==============================================================================
'                                 ClassifyRun
'------------------------------------------------------------------------------
' PURPOSE
'   Classify the run as REGULATORY, when every parameter compared (all but
'   the reporting date, reporting currency, default regime and DaysPerYear),
'   the haircut table and the risk-weight table equal the regulatory
'   parameter set REF_SET_ID and no risk weight is overridden (the overrides
'   are read with the netting sets, which are loaded first), or WHAT-IF,
'   listing each difference on Checks as INFO.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim i   As Long    'Netting set

'------------------------------------------------------------------------------
' COMPARE PARAMETERS
'------------------------------------------------------------------------------
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
        CompareTable NM_HC_TABLE, REF_HAIRCUTS, mHcIndex, mHcKeys, mHcValues, mHcRegimes, mHcCount
        CompareTable NM_RW_TABLE, REF_RISKWEIGHTS, mRwIndex, mRwKeys, mRwValues, mRwRegimes, mRwCount
        For i = 1 To mNsCount
            If mNS(i).HasRWOverride Then
                Deviation SH_NS, mNS(i).ID, "Risk weight override " & mNS(i).RWOverride & "."
            End If
        Next i
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
'==============================================================================
'                                 CompareParam
'------------------------------------------------------------------------------
' PURPOSE
'   Record a deviation when a parameter differs from the regulatory set.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If Abs(Actual - Reference) > 0.000000001 Then
            Deviation SH_PARAMS, Code, "Parameter " & Actual & " differs from " & Reference & " in " & REF_SET_ID & "."
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
'==============================================================================
'                                 CompareTable
'------------------------------------------------------------------------------
' PURPOSE
'   Record a deviation for every row of a Params table that differs from
'   the regulatory set: a value that differs, a row limited to one regime,
'   a row added, or a reference row missing.
'
' INPUTS
'   TableName: the workbook name of the table, for the message.
'   Reference: the set's rows as KEY=value pairs separated by ";".
'   Lookup, Keys, Values, Regimes, RowCount: the table as loaded.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Pairs()      As String        'Reference rows
    Dim Parts()      As String        'Key and value of one reference row
    Dim Seen         As Collection    'Reference keys
    Dim Position     As Long          'Row of a key in the table
    Dim i            As Long          'Reference row or table row

'------------------------------------------------------------------------------
' REFERENCE ROWS AGAINST THE TABLE
'------------------------------------------------------------------------------
        Set Seen = New Collection
        Pairs = Split(Reference, ";")
        For i = 0 To UBound(Pairs)
            Parts = Split(Pairs(i), "=")
            KeyAdd Seen, Parts(0), i + 1
            Position = KeyIndex(Lookup, Parts(0))
            If Position = 0 Then
                Deviation SH_PARAMS, Parts(0), TableName & ": reference row missing."
            ElseIf Abs(Values(Position) - Val(Parts(1))) > 0.000000001 Then
                Deviation SH_PARAMS, Parts(0), TableName & ": " & Values(Position) & " differs from " & _
                          Parts(1) & " in " & REF_SET_ID & "."
            ElseIf Regimes(Position) <> RG_BOTH Then
                Deviation SH_PARAMS, Parts(0), TableName & ": row limited to " & Regimes(Position) & "."
            End If
        Next i

'------------------------------------------------------------------------------
' ROWS ADDED TO THE TABLE
'------------------------------------------------------------------------------
        For i = 1 To RowCount
            If KeyIndex(Seen, Keys(i)) = 0 Then
                Deviation SH_PARAMS, Keys(i), TableName & ": row not in " & REF_SET_ID & "."
            End If
        Next i

End Sub


Private Sub Deviation( _
    ByVal SheetName As String, _
    ByVal Reference As String, _
    ByVal Message As String)
'
'==============================================================================
'                                  Deviation
'------------------------------------------------------------------------------
' PURPOSE
'   Count one difference from the regulatory set and list it as INFO.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
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
'==============================================================================
'                               LoadNettingSets
'------------------------------------------------------------------------------
' PURPOSE
'   Read every netting set row. A row whose inputs are not usable is kept
'   with its errors counted, so that its transactions are reported against
'   it and its exposure is withheld; a duplicate ID makes the first row
'   INVALID and the duplicate is not read.
'
' RETURNS
'   True when at least one netting set was read; otherwise False and the
'   run stops.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet       As Worksheet    'NettingSets sheet
    Dim FirstRow    As Long         'First data row
    Dim LastRow     As Long         'Last row with data
    Dim Block       As Variant      'Input block
    Dim ID          As String       'Netting-set ID of the row
    Dim Existing    As Long         'Index of an earlier row with the same ID
    Dim r           As Long         'Row in Block

'------------------------------------------------------------------------------
' READ THE BLOCK
'------------------------------------------------------------------------------
        Set Sheet = GetSheet(SH_NS)
        FirstRow = TableFirstRow(SH_NS)
        LastRow = LastDataRowAny(Sheet, FirstRow, NS_NCOLS, 0)
        If LastRow < FirstRow Then
            LogMsg SEV_ERROR, SH_NS, "", "No netting sets entered."
            Exit Function
        End If
        Block = TableRange(Sheet, FirstRow, 1, LastRow, NS_NCOLS).Value2
        ReDim mNS(1 To LastRow - FirstRow + 1)

'------------------------------------------------------------------------------
' READ EACH ROW
'------------------------------------------------------------------------------
        For r = 1 To UBound(Block, 1)
            If RowHasData(Block, r, NS_NCOLS, 0) Then
                ID = SafeStr(Block(r, NS_ID))
                Existing = KeyIndex(mNsIndex, ID)
                If Len(ID) = 0 Then
                    LogMsg SEV_ERROR, SH_NS, "", "Netting-set ID missing: row not read.", FirstRow + r - 1
                ElseIf Existing > 0 Then
                    NsError Existing, "Netting-set ID also entered on row " & (FirstRow + r - 1)
                Else
                    mNsCount = mNsCount + 1
                    KeyAdd mNsIndex, ID, mNsCount
                    ReadNettingSet mNsCount, Block, r, FirstRow + r - 1
                End If
            End If
        Next r
        If mNsCount = 0 Then
            LogMsg SEV_ERROR, SH_NS, "", "No usable netting set."
            Exit Function
        End If
        LoadNettingSets = True

End Function


Private Sub ReadNettingSet( _
    ByVal NsIndex As Long, _
    ByRef Block As Variant, _
    ByVal r As Long, _
    ByVal SheetRow As Long)
'
'==============================================================================
'                                ReadNettingSet
'------------------------------------------------------------------------------
' PURPOSE
'   Read and check the fields of one netting set row.
'
' INPUTS
'   NsIndex: index in mNS. Block, r: the input block and the row in it.
'   SheetRow: the sheet row, for the messages.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Flag      As Boolean    'A Y/N field read
    Dim Number    As Double     'A numeric field read
    Dim Found     As String     'Headers of the cells holding an error value

'------------------------------------------------------------------------------
' IDENTITY AND AGREEMENT
'------------------------------------------------------------------------------
        With mNS(NsIndex)
            .ID = SafeStr(Block(r, NS_ID))
            .Row = SheetRow
            .Counterparty = SafeStr(Block(r, NS_CPTY))
            If TryBool(Block(r, NS_MNA), False, Flag) And Not IsBlankCell(Block(r, NS_MNA)) Then
                .MNA = Flag
            Else
                NsError NsIndex, "Eligible master netting agreement must be Y or N"
            End If
            .SettleCcy = UTxt(Block(r, NS_SETTLE))
            If Len(.SettleCcy) = 0 Then
                .SettleCcy = pRepCcy
            ElseIf KeyIndex(mFxIndex, .SettleCcy) = 0 Then
                NsError NsIndex, "Settlement currency " & .SettleCcy & " not in the FX table"
            End If
            .Regime = UTxt(Block(r, NS_REGIME))
            If Len(.Regime) = 0 Then
                .Regime = pRegime
            ElseIf .Regime <> RG_CRR And .Regime <> RG_BCBS Then
                NsError NsIndex, "Regime override must be blank, CRR or BCBS"
                .Regime = pRegime
            End If

'------------------------------------------------------------------------------
' COUNTERPARTY
'------------------------------------------------------------------------------
            .CptyClass = UTxt(Block(r, NS_CLASS))
            Select Case .CptyClass
                Case CC_SOV, CC_INST, CC_CORP, CC_QCCP
                Case Else
                    NsError NsIndex, "Counterparty class must be SOV, INST, CORP or QCCP"
            End Select
            If Not IsBlankCell(Block(r, NS_CQS)) Then
                If IsNum(Block(r, NS_CQS)) Then
                    Number = CDbl(Block(r, NS_CQS))
                Else
                    Number = 0#
                End If
                If Number < 1# Or Number > 6# Or Number <> Int(Number) Then
                    NsError NsIndex, "Counterparty CQS must be blank or a whole number from 1 to 6"
                Else
                    .CQS = CLng(Number)
                End If
            End If
            .SCRA = UTxt(Block(r, NS_SCRA))
            If .CptyClass = CC_INST And .CQS = 0 Then
                If InStr(1, " " & SCRA_GRADES & " ", " " & .SCRA & " ") = 0 Or Len(.SCRA) = 0 Then
                    NsError NsIndex, "Unrated institution: SCRA grade must be A, A_ENH, B or C"
                End If
            ElseIf Len(.SCRA) > 0 Then
                LogMsg SEV_WARN, SH_NS, .ID, "SCRA grade ignored: it applies to unrated institutions only.", .Row
            End If

'------------------------------------------------------------------------------
' HOLDING PERIOD DRIVERS
'------------------------------------------------------------------------------
            .RemarginBD = 1#
            If TryDbl(Block(r, NS_FREQ), 1#, Number) Then
                If Number < 1# Then
                    NsError NsIndex, "Remargin frequency must be 1 business day or more"
                Else
                    .RemarginBD = Number
                End If
            Else
                NsError NsIndex, "Remargin frequency must be a number"
            End If
            If TryBool(Block(r, NS_LARGE), False, Flag) Then
                .LargeOrIlliquid = Flag
            Else
                NsError NsIndex, ">5,000 trades or illiquid collateral must be Y or N"
            End If
            If TryBool(Block(r, NS_DISPUTE), False, Flag) Then
                .Disputes = Flag
            Else
                NsError NsIndex, "Margin disputes must be Y or N"
            End If
            If TryDbl(Block(r, NS_HOLDOVR), 0#, Number) Then
                If Number < 0# Then
                    NsError NsIndex, "Holding period override must be 0 or more"
                Else
                    .HoldingOverride = Number
                End If
            Else
                NsError NsIndex, "Holding period override must be a number"
            End If
            If TryBool(Block(r, NS_ZEROHC), False, Flag) Then
                .ZeroHaircut = Flag
            Else
                NsError NsIndex, "0% haircut conditions met must be Y or N"
            End If

'------------------------------------------------------------------------------
' RISK WEIGHT OVERRIDE
'------------------------------------------------------------------------------
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
    'fields already reject it, so for them this is a second message.
        Found = ErrorColumns(Block, r, NS_HEADERS, 0)
        If Len(Found) > 0 Then
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
'==============================================================================
'                                   LoadLegs
'------------------------------------------------------------------------------
' PURPOSE
'   Read every row of the Transactions sheet as a leg, check its fields,
'   value it in the reporting currency and find its supervisory haircut.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet      As Worksheet    'Transactions sheet
    Dim FirstRow   As Long         'First data row
    Dim LastRow    As Long         'Last row with data
    Dim Block      As Variant      'Input block
    Dim r          As Long         'Row in Block

'------------------------------------------------------------------------------
' READ THE BLOCK
'------------------------------------------------------------------------------
        Set Sheet = GetSheet(SH_TX)
        FirstRow = TableFirstRow(SH_TX)
        LastRow = LastDataRowAny(Sheet, FirstRow, TX_NCOLS, TX_COMMENT)
        If LastRow < FirstRow Then
            LogMsg SEV_WARN, SH_TX, "", "No transactions entered."
            Exit Sub
        End If
        Block = TableRange(Sheet, FirstRow, 1, LastRow, TX_NCOLS).Value2
        ReDim mLegs(1 To LastRow - FirstRow + 1)

'------------------------------------------------------------------------------
' READ EACH ROW
'------------------------------------------------------------------------------
        For r = 1 To UBound(Block, 1)
            If RowHasData(Block, r, TX_NCOLS, TX_COMMENT) Then
                mLegCount = mLegCount + 1
                ReadLeg mLegCount, Block, r, FirstRow + r - 1
            End If
        Next r

End Sub


Private Sub ReadLeg( _
    ByVal LegIndex As Long, _
    ByRef Block As Variant, _
    ByVal r As Long, _
    ByVal SheetRow As Long)
'
'==============================================================================
'                                   ReadLeg
'------------------------------------------------------------------------------
' PURPOSE
'   Read and check the fields of one leg; every input error is an ERROR
'   that rejects the leg and, through BuildTransactions, its transaction.
'   A valid leg is then keyed to its haircut (FindHaircut).
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Found   As String    'Headers of the cells holding an error value

'------------------------------------------------------------------------------
' IDENTITY
'------------------------------------------------------------------------------
        With mLegs(LegIndex)
            .Row = SheetRow
            .Valid = True
            .LegID = SafeStr(Block(r, TX_LEG))
            .TxID = SafeStr(Block(r, TX_ID))
            .TxKey = UCase$(.TxID)
            .ResidualYears = -1#
            If Len(.LegID) = 0 Then
                .LegID = "(row " & SheetRow & ")"
                LegError LegIndex, "Leg ID missing"
            ElseIf KeyIndex(mLegIndex, .LegID) > 0 Then
                LegError LegIndex, "Leg ID entered twice"
            Else
                KeyAdd mLegIndex, .LegID, LegIndex
            End If
            If Len(.TxID) = 0 Then
                .TxID = "(row " & SheetRow & ")"
                .TxKey = UCase$(.TxID)
                LegError LegIndex, "Transaction ID missing"
            End If
            .NsIndex = KeyIndex(mNsIndex, SafeStr(Block(r, TX_NS)))
            If .NsIndex = 0 Then
                LegError LegIndex, "Netting set '" & SafeStr(Block(r, TX_NS)) & "' not on NettingSets"
            End If

'------------------------------------------------------------------------------
' TRANSACTION FIELDS
'------------------------------------------------------------------------------
            .TxType = UTxt(Block(r, TX_TYPE))
            Select Case .TxType
                Case TT_REPO, TT_REVREPO, TT_SECLEND, TT_SECBORROW, TT_MARGINLEND
                Case Else
                    LegError LegIndex, "Transaction type '" & .TxType & "' is not REPO, REVREPO, SECLEND, " & _
                                       "SECBORROW or MARGINLEND"
            End Select
            .StartSerial = ToSerial(Block(r, TX_START))
            If .StartSerial <= 0# Then
                LegError LegIndex, "Start date missing or not a date"
            End If
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
        ReadLegAsset LegIndex, Block, r
        With mLegs(LegIndex)
            If IsNum(Block(r, TX_MV)) Then
                .MarketValue = CDbl(Block(r, TX_MV))
            End If
            If .MarketValue <= 0# Then
                LegError LegIndex, "Market value must be a number above 0"
            End If
            .Ccy = UTxt(Block(r, TX_CCY))
            If FXRate(.Ccy) = 0# Then
                LegError LegIndex, "Currency '" & .Ccy & "' not in the FX table"
            Else
                .Value = .MarketValue * FXRate(.Ccy)
            End If
            If .Asset = AT_CASH Then
                .IssueKey = "CASH " & .Ccy
            End If
            .Signature = .Asset & "|" & .Issuer & "|" & .CQS & "|" & .SecMatSerial & "|" & .Listing & "|" & _
                         .FundH10 & "|" & .Ccy
        End With

'------------------------------------------------------------------------------
' WORKSHEET ERROR VALUES
'------------------------------------------------------------------------------
    'An error value would read as a blank cell and take the field's default.
        Found = ErrorColumns(Block, r, TX_HEADERS, TX_COMMENT)
        If Len(Found) > 0 Then
            LegError LegIndex, "Worksheet error value in: " & Found
        End If
        If mLegs(LegIndex).Valid Then
            FindHaircut LegIndex
        End If

End Sub


Private Function ErrorColumns( _
    ByRef Block As Variant, _
    ByVal r As Long, _
    ByVal Headers As String, _
    ByVal SkipColumn As Long) _
    As String
'
'==============================================================================
'                                 ErrorColumns
'------------------------------------------------------------------------------
' PURPOSE
'   List the headers of the cells of one input row that hold a worksheet
'   error value (#N/A, #VALUE! and others).
'
' INPUTS
'   Block, r: the input block and the row in it. Headers: the table's
'   headers, separated by |. SkipColumn: a column that is not read, or 0.
'
' RETURNS
'   The headers, separated by "; ", or an empty string.
'
' UPDATED
'   2026-10-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Names()   As String    'Headers, 0-based
    Dim Col       As Long      'Column in Block
    Dim Label     As String    'Header of the column

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
        Names = Split(Headers, "|")
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
'==============================================================================
'                                 ReadLegAsset
'------------------------------------------------------------------------------
' PURPOSE
'   Read and check the asset fields of one leg: asset type, issue ID and
'   the fields each asset type needs.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Number   As Double    'A numeric field read

'------------------------------------------------------------------------------
' ASSET TYPE AND ISSUE
'------------------------------------------------------------------------------
        With mLegs(LegIndex)
            .Asset = UTxt(Block(r, TX_ASSET))
            Select Case .Asset
                Case AT_CASH, AT_DEBT, AT_EQUITY, AT_GOLD, AT_FUND, AT_OTHER
                Case Else
                    LegError LegIndex, "Asset type '" & .Asset & "' is not CASH, DEBT, EQUITY, GOLD, FUND or OTHER"
                    Exit Sub
            End Select
            If .Asset <> AT_CASH Then
                .IssueKey = UTxt(Block(r, TX_ISSUE))
                If Len(.IssueKey) = 0 Then
                    LegError LegIndex, "Issue ID missing: every non-cash leg names its issue"
                End If
            End If

'------------------------------------------------------------------------------
' DEBT SECURITY
'------------------------------------------------------------------------------
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
                            .CQS = CLng(Number)
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
            End Select
        End With

End Sub


Private Sub FindHaircut( _
    ByVal LegIndex As Long)
'
'==============================================================================
'                                 FindHaircut
'------------------------------------------------------------------------------
' PURPOSE
'   Build the haircut-table key of a valid leg, and set its H10 and
'   eligibility. A leg other than a fund whose key is not in the table for
'   its regime is not eligible: lent, it takes NonEligibleHaircut; received,
'   it is not recognised as collateral. A fund is eligible with the H10
'   entered on its leg.
'
' REFERENCE
'   CRE22.49, 22.54; CRR Art. 197, 224(1), (4)-(6).
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Regime     As String    'Regime of the leg's netting set
    Dim Category   As String    'Haircut-table issuer category: SOV, OTH or SEC
    Dim Band       As String    'Credit-quality band
    Dim Position   As Long      'Row in the haircut table

'------------------------------------------------------------------------------
' BUILD THE KEY
'------------------------------------------------------------------------------
        With mLegs(LegIndex)
            Regime = mNS(.NsIndex).Regime
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
                    ElseIf Len(Band) > 0 Then
                        .HaircutKey = "DEBT_" & Category & "_" & Band & "_" & MaturityBand(.ResidualYears)
                    End If
            End Select

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
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
    ByVal Years As Double) _
    As String
'
'==============================================================================
'                                 MaturityBand
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
'==============================================================================
'
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
'==============================================================================
'                              BuildTransactions
'------------------------------------------------------------------------------
' PURPOSE
'   Group the legs into transactions by Transaction ID, check each
'   transaction as a whole, and reject it with all its legs when any leg is
'   invalid or the legs disagree. Count the transactions of each netting
'   set. A rejected transaction whose legs name several netting sets is
'   counted, as rejected, in each of them, so that none is reported VALID
'   without its legs.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LegInvalid()   As Boolean    'Per transaction: some leg is invalid
    Dim RejectedKeys   As Collection 'Transaction|netting set pairs counted as rejected
    Dim t              As Long       'Transaction
    Dim i              As Long       'Leg
    Dim Ns             As Long       'Netting set

'------------------------------------------------------------------------------
' GROUP
'------------------------------------------------------------------------------
        If mLegCount = 0 Then
            Exit Sub
        End If
        ReDim mTx(1 To mLegCount)
        ReDim LegInvalid(1 To mLegCount)
        For i = 1 To mLegCount
            t = KeyIndex(mTxIndex, mLegs(i).TxKey)
            If t = 0 Then
                mTxCount = mTxCount + 1
                t = mTxCount
                KeyAdd mTxIndex, mLegs(i).TxKey, t
                With mTx(t)
                    .ID = mLegs(i).TxID
                    .NsIndex = mLegs(i).NsIndex
                    .TxType = mLegs(i).TxType
                    .StartSerial = mLegs(i).StartSerial
                    .MatSerial = mLegs(i).MatSerial
                    .FirstRow = mLegs(i).Row
                End With
            Else
                CheckSameTransaction t, i
            End If
            mLegs(i).TxIndex = t
            AddLegToTransaction t, i
            If Not mLegs(i).Valid Then
                LegInvalid(t) = True
            End If
        Next i

'------------------------------------------------------------------------------
' CHECK EACH TRANSACTION
'------------------------------------------------------------------------------
        For t = 1 To mTxCount
            CheckTransaction t, LegInvalid(t)
            If mTx(t).NsIndex > 0 Then
                With mNS(mTx(t).NsIndex)
                    .TxCount = .TxCount + 1
                    If mTx(t).Rejected Then
                        .TxRejected = .TxRejected + 1
                    End If
                End With
            End If
        Next t

'------------------------------------------------------------------------------
' MARK THE LEGS OF REJECTED TRANSACTIONS AND COUNT THEM IN THEIR OTHER NETTING SETS
'------------------------------------------------------------------------------
        Set RejectedKeys = New Collection
        For i = 1 To mLegCount
            t = mLegs(i).TxIndex
            If mTx(t).Rejected Then
                mLegs(i).TxRejected = True
                mLegs(i).Recognised = False
                If mLegs(i).Valid Then
                    mLegs(i).Messages = AppendText(mLegs(i).Messages, "Transaction rejected")
                End If
                Ns = mLegs(i).NsIndex
                If Ns > 0 And Ns <> mTx(t).NsIndex Then
                    If KeyIndex(RejectedKeys, t & "|" & Ns) = 0 Then
                        KeyAdd RejectedKeys, t & "|" & Ns, 1
                        With mNS(Ns)
                            .TxCount = .TxCount + 1
                            .TxRejected = .TxRejected + 1
                        End With
                    End If
                End If
            End If
        Next i

End Sub


Private Sub CheckSameTransaction( _
    ByVal t As Long, _
    ByVal i As Long)
'
'==============================================================================
'                             CheckSameTransaction
'------------------------------------------------------------------------------
' PURPOSE
'   Check that a further leg of a transaction agrees with its first leg on
'   the netting set, type, start date and maturity date.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        With mLegs(i)
            If .NsIndex <> mTx(t).NsIndex Then
                LegError i, "Netting set differs from the transaction's first leg (row " & mTx(t).FirstRow & ")"
            ElseIf .TxType <> mTx(t).TxType Then
                LegError i, "Transaction type differs from the transaction's first leg (row " & mTx(t).FirstRow & ")"
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
'==============================================================================
'                             AddLegToTransaction
'------------------------------------------------------------------------------
' PURPOSE
'   Count a leg in its transaction: sides, non-cash legs and the currency
'   of the legs lent.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        With mTx(t)
            .LegCount = .LegCount + 1
            If mLegs(i).Sign > 0 Then
                .LentCount = .LentCount + 1
                If mLegs(i).Asset <> AT_CASH Then
                    .LentNonCash = .LentNonCash + 1
                End If
                If Len(.LentCcy) = 0 Then
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
'==============================================================================
'                               CheckTransaction
'------------------------------------------------------------------------------
' PURPOSE
'   Check a transaction as a whole: rejected when a leg is invalid, when it
'   has no leg lent or no leg received, or, without netting, when its legs
'   lent are in several currencies (the single-transaction formula needs
'   one exposure currency). Warn when its legs do not look like its type,
'   and decide whether it is short-term.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Reference   As String    'Transaction ID, for the messages

'------------------------------------------------------------------------------
' REJECT
'------------------------------------------------------------------------------
        With mTx(t)
            Reference = .ID
            If HasInvalidLeg Then
                .Rejected = True
                Exit Sub
            End If
            If .LentCount = 0 Or .ReceivedCount = 0 Then
                .Rejected = True
                LogMsg SEV_ERROR, SH_TX, Reference, "Transaction needs at least one leg LENT and one RECEIVED - " & _
                       "transaction rejected.", .FirstRow
                Exit Sub
            End If
            If Not mNS(.NsIndex).MNA And .LentCcy = "*" Then
                .Rejected = True
                LogMsg SEV_ERROR, SH_TX, Reference, "Without an eligible master netting agreement the legs lent " & _
                       "must share one currency - transaction rejected.", .FirstRow
                Exit Sub
            End If

'------------------------------------------------------------------------------
' WARN
'------------------------------------------------------------------------------
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
            End Select
            If .MatSerial > 0# And .MatSerial < pAsOf Then
                LogMsg SEV_WARN, SH_TX, Reference, "Maturity date before the reporting date: the transaction " & _
                       "is still calculated.", .FirstRow
            End If

'------------------------------------------------------------------------------
' SHORT-TERM
'------------------------------------------------------------------------------
    'Original maturity of three months or less; an open transaction is not
    'short-term.
            If .MatSerial > 0# Then
                .ShortTerm = (CDbl(DateAdd("m", SHORT_TERM_MONTHS, CDate(.StartSerial))) >= .MatSerial)
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
'==============================================================================
'                               CheckNettingSets
'------------------------------------------------------------------------------
' PURPOSE
'   For every netting set: its transaction types, whether every transaction
'   is short-term, the zero-haircut conditions, the consistency of each
'   security issue across its legs, the holding period override, the
'   risk-weight key, and then its status.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ns  As Long    'Netting set
    Dim t   As Long    'Transaction

'------------------------------------------------------------------------------
' TRANSACTION TYPES AND SHORT-TERM STATUS
'------------------------------------------------------------------------------
        For Ns = 1 To mNsCount
            mNS(Ns).AllShortTerm = True
        Next Ns
        For t = 1 To mTxCount
            If mTx(t).NsIndex > 0 And Not mTx(t).Rejected Then
                With mNS(mTx(t).NsIndex)
                    If mTx(t).TxType = TT_MARGINLEND Then
                        .HasCapitalMarket = True
                    Else
                        .HasRepoStyle = True
                    End If
                    If Not mTx(t).ShortTerm Then
                        .AllShortTerm = False
                    End If
                End With
            End If
        Next t

'------------------------------------------------------------------------------
' CONDITIONS, ISSUES, HOLDING PERIOD AND RISK WEIGHT
'------------------------------------------------------------------------------
        CheckZeroHaircut
        CheckIssues
        For Ns = 1 To mNsCount
            CheckHoldingOverride Ns
            FindRiskWeight Ns
            SetStatus Ns
        Next Ns

End Sub


Private Sub CheckZeroHaircut()
'
'==============================================================================
'                               CheckZeroHaircut
'------------------------------------------------------------------------------
' PURPOSE
'   For every netting set declared to meet the conditions of the 0%
'   haircut, check those the inputs show: repo-style transactions only,
'   every recognised leg cash or an eligible debt security of a central
'   government or central bank, one currency, daily remargining. The other
'   conditions (core market participant, documentation, settlement) are
'   declared by the Y on NettingSets and listed as INFO.
'
' REFERENCE
'   CRE22.60, 22.36-22.37; CRR Art. 227.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Failures()   As String    'Per netting set: conditions not met
    Dim FirstCcy()   As String    'Per netting set: currency the next leg is compared with
    Dim Ns           As Long      'Netting set
    Dim i            As Long      'Leg

'------------------------------------------------------------------------------
' NETTING-SET CONDITIONS
'------------------------------------------------------------------------------
        ReDim Failures(1 To mNsCount)
        ReDim FirstCcy(1 To mNsCount)
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
        For i = 1 To mLegCount
            Ns = mLegs(i).NsIndex
            If Ns > 0 Then
                If mNS(Ns).ZeroHaircut And mLegs(i).Recognised Then
                    If Len(FirstCcy(Ns)) = 0 Then
                        FirstCcy(Ns) = mLegs(i).Ccy
                    ElseIf mLegs(i).Ccy <> FirstCcy(Ns) Then
                        Failures(Ns) = AppendText(Failures(Ns), "legs in more than one currency (" & _
                                                 mLegs(i).LegID & ")")
                        FirstCcy(Ns) = mLegs(i).Ccy
                    End If
                    If Not (mLegs(i).Asset = AT_CASH Or (mLegs(i).Asset = AT_DEBT And mLegs(i).Issuer = IS_SOV)) Then
                        Failures(Ns) = AppendText(Failures(Ns), "leg " & mLegs(i).LegID & " is neither cash " & _
                                                 "nor a central government security")
                    ElseIf Not mLegs(i).Eligible Then
                        Failures(Ns) = AppendText(Failures(Ns), "leg " & mLegs(i).LegID & " is not eligible")
                    End If
                End If
            End If
        Next i

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
        For Ns = 1 To mNsCount
            If mNS(Ns).ZeroHaircut Then
                If Len(Failures(Ns)) > 0 Then
                    NsError Ns, "0% haircut declared but its conditions are not met: " & Failures(Ns)
                Else
                    LogMsg SEV_INFO, SH_NS, mNS(Ns).ID, "0% haircut applied: the conditions on the counterparty, " & _
                           "documentation and settlement are taken as declared.", mNS(Ns).Row
                End If
            End If
        Next Ns

End Sub


Private Sub CheckIssues()
'
'==============================================================================
'                                 CheckIssues
'------------------------------------------------------------------------------
' PURPOSE
'   Check that the recognised legs of one security issue in a netting set
'   describe it in the same way (asset type, issuer, CQS, maturity,
'   listing, fund haircut and currency), since the netting formula gives
'   each issue one haircut.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Seen    As Collection    'Netting set and issue key to the first leg
    Dim Key     As String        'Netting set and issue key
    Dim First   As Long          'First leg of an issue
    Dim i       As Long          'Leg

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        Set Seen = New Collection
        For i = 1 To mLegCount
            If mLegs(i).NsIndex > 0 And mLegs(i).Recognised And mLegs(i).Asset <> AT_CASH Then
                Key = mLegs(i).NsIndex & "|" & mLegs(i).IssueKey
                First = KeyIndex(Seen, Key)
                If First = 0 Then
                    KeyAdd Seen, Key, i
                ElseIf mLegs(First).Signature <> mLegs(i).Signature Then
                    NsError mLegs(i).NsIndex, "Issue " & mLegs(i).IssueKey & " is described differently on legs " & _
                                              mLegs(First).LegID & " and " & mLegs(i).LegID
                End If
            End If
        Next i

End Sub


Private Sub CheckHoldingOverride( _
    ByVal Ns As Long)
'
'==============================================================================
'                             CheckHoldingOverride
'------------------------------------------------------------------------------
' PURPOSE
'   Warn when a holding period override is shorter than the minimum of the
'   netting set's transaction types taken together: the override can only
'   lengthen it. Without netting the minimum is set per transaction, so an
'   override between the repo-style and the capital-market minimum still
'   lengthens the repo-style transactions.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Minimum   As Double    'Holding period without override or disputes

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        With mNS(Ns)
            If .HoldingOverride <= 0# Then
                Exit Sub
            End If
            Minimum = Checked(SFT_HoldingPeriod(MinimumHolding(.HasRepoStyle, .HasCapitalMarket), _
                                                .LargeOrIlliquid, False, 0#, pHoldLarge), "SFT_HoldingPeriod")
            If .HoldingOverride < Minimum Then
                LogMsg SEV_WARN, SH_NS, .ID, "Holding period override " & .HoldingOverride & _
                       " BD is below the minimum " & Minimum & " BD and is not used.", .Row
            End If
        End With

End Sub


Private Sub FindRiskWeight( _
    ByVal Ns As Long)
'
'==============================================================================
'                                FindRiskWeight
'------------------------------------------------------------------------------
' PURPOSE
'   Build the risk-weight key of a netting set from its counterparty
'   class, CQS, SCRA grade and short-term status, and find its risk weight;
'   an override replaces the table value. A netting set already INVALID is
'   not looked up, so that a missing input is not reported twice.
'
' REFERENCE
'   CRE20.7, 20.18-20.21, CRE54.7; CRR Art. 114, 120-122, 306(1)(a).
'
' UPDATED
'   2026-10-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ShortPart   As String    '"ST_" for a short-term institution exposure
    Dim Position    As Long      'Row in the risk-weight table

'------------------------------------------------------------------------------
' BUILD THE KEY
'------------------------------------------------------------------------------
        With mNS(Ns)
            If .AllShortTerm And .TxCount > .TxRejected Then
                ShortPart = "ST_"
            End If
            Select Case .CptyClass
                Case CC_QCCP
                    .RWKey = "QCCP"
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
            If .HasRWOverride Then
                .RiskWeight = .RWOverride
                LogMsg SEV_WARN, SH_NS, .ID, "Risk weight override " & Format$(.RWOverride, "0.0%") & _
                       " used instead of " & .RWKey & ".", .Row
                .RWKey = .RWKey & " (override)"
                Exit Sub
            End If
            If .InputErrors > 0 Then
                Exit Sub
            End If
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
'==============================================================================
'                                  SetStatus
'------------------------------------------------------------------------------
' PURPOSE
'   Status of a netting set: INVALID with input errors, INCOMPLETE with a
'   rejected transaction, NO TRANSACTIONS, or VALID. Only a VALID netting
'   set receives an exposure value.
'
' UPDATED
'   2026-10-25
'==============================================================================
'
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
'==============================================================================
'                              ComputeNettingSets
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate every netting set: the holding period of each netting set
'   with netting and of each transaction without, the haircut of every
'   leg in one pass, then E* of every VALID netting set, its RWA and own
'   funds requirement, and the totals per regime.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Ns  As Long    'Netting set
    Dim t   As Long    'Transaction
    Dim i   As Long    'Leg
    Dim r   As Long    'Regime index

'------------------------------------------------------------------------------
' HOLDING PERIODS
'------------------------------------------------------------------------------
        For Ns = 1 To mNsCount
            With mNS(Ns)
                If .MNA Then
                    .HoldingBD = HoldingFor(Ns, .HasRepoStyle, .HasCapitalMarket)
                    .HaircutScale = Sqr((.RemarginBD + .HoldingBD - 1#) / pHcBase)
                End If
            End With
        Next Ns
        For t = 1 To mTxCount
            Ns = mTx(t).NsIndex
            If Ns > 0 Then
                If Not mNS(Ns).MNA Then
                    mTx(t).HoldingBD = HoldingFor(Ns, mTx(t).TxType <> TT_MARGINLEND, mTx(t).TxType = TT_MARGINLEND)
                End If
            End If
        Next t

'------------------------------------------------------------------------------
' LEGS
'------------------------------------------------------------------------------
    'Every leg of a known netting set gets its haircut. LegCalc shows it for
    'the recognised legs of transactions not rejected; only recognised legs
    'of VALID netting sets enter the sums.
        For i = 1 To mLegCount
            Ns = mLegs(i).NsIndex
            If Ns > 0 Then
                If mNS(Ns).MNA Then
                    ApplyHaircut i, mNS(Ns).HoldingBD, ""
                    If mLegs(i).Recognised And mNS(Ns).Status = "VALID" Then
                        AddToPositions Ns, i
                    End If
                Else
                    t = mLegs(i).TxIndex
                    If Ns = mTx(t).NsIndex Then
                        ApplyHaircut i, mTx(t).HoldingBD, mTx(t).LentCcy
                    Else
                    'A leg filed under another netting set than its transaction's
                    'first leg: the transaction is rejected, and any holding period
                    'it has is the first leg's netting set's, so the leg takes its
                    'own. The haircut is neither shown on LegCalc nor summed.
                        ApplyHaircut i, HoldingFor(Ns, mTx(t).TxType <> TT_MARGINLEND, _
                                                   mTx(t).TxType = TT_MARGINLEND), mTx(t).LentCcy
                    End If
                    If mLegs(i).Recognised And mNS(Ns).Status = "VALID" Then
                        AddToTransaction t, i
                    End If
                End If
            End If
        Next i

'------------------------------------------------------------------------------
' EXPOSURE VALUE
'------------------------------------------------------------------------------
        FinishNetted
        FinishSingles

'------------------------------------------------------------------------------
' CAPITAL AND TOTALS
'------------------------------------------------------------------------------
        For Ns = 1 To mNsCount
            With mNS(Ns)
                r = RegimeIndex(.Regime)
                mRegSets(r) = mRegSets(r) + 1
                If .Status = "VALID" Then
                    .RWA = Checked(SFT_CapitalRequirement(.EStar, .RiskWeight, 1#), "SFT_CapitalRequirement")
                    .OwnFunds = Checked(SFT_CapitalRequirement(.EStar, .RiskWeight, pCapRatio), _
                                        "SFT_CapitalRequirement")
                    mRegValid(r) = mRegValid(r) + 1
                    mRegEStar(r) = mRegEStar(r) + .EStar
                    mRegRWA(r) = mRegRWA(r) + .RWA
                    mRegOwnFunds(r) = mRegOwnFunds(r) + .OwnFunds
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
    ByVal CapitalMarket As Boolean) _
    As Double
'
'==============================================================================
'                                MinimumHolding
'------------------------------------------------------------------------------
' PURPOSE
'   Minimum holding period of the transaction types covered: 5 business
'   days for repo-style transactions, 10 for other capital-market-driven
'   transactions, and the longer of the two when both are present.
'
' REFERENCE
'   CRE22.57-22.58; CRR Art. 224(2). The rule for both present is in
'   CRE22.58 only; the CRR does not state it (PROTOTYPE.md, assumption 1).
'
' UPDATED
'   2026-10-26
'==============================================================================
'
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
    ByVal CapitalMarket As Boolean) _
    As Double
'
'==============================================================================
'                                  HoldingFor
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
'==============================================================================
'
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
'==============================================================================
'                                 ApplyHaircut
'------------------------------------------------------------------------------
' PURPOSE
'   Scale the haircut of one leg for its holding period and its netting
'   set's remargining frequency, and set its adjusted value. With
'   ExposureCcy given (single-transaction formula), a leg received in
'   another currency also takes the currency haircut. A 0% haircut netting
'   set takes no haircut at all.
'
' INPUTS
'   i: the leg. HoldingBD: T_M. ExposureCcy: the currency of the legs lent
'   of its transaction, or "" under netting.
'
' REFERENCE
'   CRE22.46, 22.52, 22.59-22.60; CRR Art. 223, 224, 226, 227.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        With mLegs(i)
            .HoldingBD = HoldingBD
            .Scale = Sqr((mNS(.NsIndex).RemarginBD + HoldingBD - 1#) / pHcBase)
            .H = 0#
            .Hfx = 0#
            If Not mNS(.NsIndex).ZeroHaircut Then
                .H = Checked(SFT_ScaledHaircut(.H10, HoldingBD, mNS(.NsIndex).RemarginBD, pHcBase), _
                             "SFT_ScaledHaircut")
            End If
            If .Sign > 0 Then
                .Adjusted = .Value * (1# + .H)
            Else
                If Len(ExposureCcy) > 0 And .Ccy <> ExposureCcy And Not mNS(.NsIndex).ZeroHaircut Then
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
'==============================================================================
'                                AddToPositions
'------------------------------------------------------------------------------
' PURPOSE
'   Add a recognised leg of a netting set with netting to the netting
'   set's totals, to the position of its currency and, for a security, to
'   the position of its issue.
'
' UPDATED
'   2026-10-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim p   As Long    'Position

'------------------------------------------------------------------------------
' NETTING-SET TOTALS
'------------------------------------------------------------------------------
        With mLegs(i)
            mNS(Ns).LegsUsed = mNS(Ns).LegsUsed + 1
            If .Sign > 0 Then
                mNS(Ns).SumE = mNS(Ns).SumE + .Value
            Else
                mNS(Ns).SumC = mNS(Ns).SumC + .Value
            End If

'------------------------------------------------------------------------------
' CURRENCY AND ISSUE
'------------------------------------------------------------------------------
            p = PositionFor(Ns, PK_CURRENCY, .Ccy, "", .Ccy)
            AddLegValue p, i
            If .Asset <> AT_CASH Then
                p = PositionFor(Ns, PK_ISSUE, .IssueKey, .Asset, .Ccy)
                mPos(p).H = .H
                AddLegValue p, i
            End If
        End With

End Sub


Private Function PositionFor( _
    ByVal Ns As Long, _
    ByVal Kind As String, _
    ByVal Key As String, _
    ByVal Asset As String, _
    ByVal Ccy As String) _
    As Long
'
'==============================================================================
'                                 PositionFor
'------------------------------------------------------------------------------
' PURPOSE
'   Index of a position of a netting set, created when new.
'
' UPDATED
'   2026-10-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LookupKey   As String    'Netting set, kind and key
    Dim p           As Long      'Position

'------------------------------------------------------------------------------
' FIND OR CREATE
'------------------------------------------------------------------------------
        LookupKey = Ns & "|" & Kind & "|" & Key
        p = KeyIndex(mPosIndex, LookupKey)
        If p = 0 Then
            mPosCount = mPosCount + 1
            If mPosCount = 1 Then
                ReDim mPos(1 To 64)
            ElseIf mPosCount > UBound(mPos) Then
                ReDim Preserve mPos(1 To mPosCount * 2)
            End If
            p = mPosCount
            KeyAdd mPosIndex, LookupKey, p
            mPos(p).NsIndex = Ns
            mPos(p).Kind = Kind
            mPos(p).Key = Key
            mPos(p).Asset = Asset
            mPos(p).Ccy = Ccy
        End If
        PositionFor = p

End Function


Private Sub AddLegValue( _
    ByVal p As Long, _
    ByVal i As Long)
'
'==============================================================================
'                                 AddLegValue
'------------------------------------------------------------------------------
' PURPOSE
'   Add a leg's value to the lent or received side of a position.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
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
'==============================================================================
'                               AddToTransaction
'------------------------------------------------------------------------------
' PURPOSE
'   Add a recognised leg of a netting set without netting to its
'   transaction: the value lent or received and its adjusted value.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
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
'==============================================================================
'                                 FinishNetted
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate E* of every VALID netting set with an eligible master netting
'   agreement from its positions: the net position and haircut of each
'   security issue, the net and gross exposures, N, the FX term over the
'   currencies other than the settlement currency, and E*.
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
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Starts()   As Long      'First slot of each netting set in Order
    Dim Filled()   As Long      'Slots filled so far per netting set
    Dim Order()    As Long      'Position indices grouped by netting set
    Dim Ns         As Long      'Netting set
    Dim p          As Long      'Position or slot

'------------------------------------------------------------------------------
' GROUP THE POSITIONS BY NETTING SET
'------------------------------------------------------------------------------
        If mPosCount = 0 Then
            Exit Sub
        End If
        ReDim Starts(1 To mNsCount + 1)
        ReDim Filled(1 To mNsCount)
        ReDim Order(1 To mPosCount)
        For p = 1 To mPosCount
            Filled(mPos(p).NsIndex) = Filled(mPos(p).NsIndex) + 1
        Next p
        Starts(1) = 1
        For Ns = 1 To mNsCount
            Starts(Ns + 1) = Starts(Ns) + Filled(Ns)
            Filled(Ns) = 0
        Next Ns
        For p = 1 To mPosCount
            Ns = mPos(p).NsIndex
            Order(Starts(Ns) + Filled(Ns)) = p
            Filled(Ns) = Filled(Ns) + 1
        Next p

'------------------------------------------------------------------------------
' CALCULATE EACH NETTING SET
'------------------------------------------------------------------------------
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
'==============================================================================
'                               FinishOneNetted
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate E* of one netting set with netting from its positions,
'   Order(FirstSlot) to Order(LastSlot).
'
' REFERENCE
'   CRE22.64-22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Es()        As Double    'Net position of each issue
    Dim Hs()        As Double    'Haircut of each issue
    Dim IssuePos()  As Long      'Position index of each issue
    Dim Issues      As Long      'Issues
    Dim Hfx         As Double    'Scaled currency haircut
    Dim Largest     As Double    'Largest |E_s|
    Dim s           As Long      'Slot or issue
    Dim p           As Long      'Position

'------------------------------------------------------------------------------
' ISSUES
'------------------------------------------------------------------------------
        ReDim Es(1 To LastSlot - FirstSlot + 1)
        ReDim Hs(1 To LastSlot - FirstSlot + 1)
        ReDim IssuePos(1 To LastSlot - FirstSlot + 1)
        For s = FirstSlot To LastSlot
            p = Order(s)
            If mPos(p).Kind = PK_ISSUE Then
                Issues = Issues + 1
                Es(Issues) = mPos(p).Lent - mPos(p).Received
                Hs(Issues) = mPos(p).H
                IssuePos(Issues) = p
                If Abs(Es(Issues)) > Largest Then
                    Largest = Abs(Es(Issues))
                End If
            End If
        Next s
        For s = 1 To Issues
            mPos(IssuePos(s)).Counted = (Es(s) <> 0# And Abs(Es(s)) >= pThreshold * Largest)
        Next s

'------------------------------------------------------------------------------
' CURRENCIES
'------------------------------------------------------------------------------
        With mNS(Ns)
            If Not .ZeroHaircut Then
                Hfx = Checked(SFT_ScaledHaircut(pHfx, .HoldingBD, .RemarginBD, pHcBase), "SFT_ScaledHaircut")
            End If
            For s = FirstSlot To LastSlot
                p = Order(s)
                If mPos(p).Kind = PK_CURRENCY Then
                    mPos(p).H = Hfx
                    If mPos(p).Ccy <> .SettleCcy Then
                        .FXTerm = .FXTerm + Abs(mPos(p).Lent - mPos(p).Received) * Hfx
                    End If
                End If
            Next s

'------------------------------------------------------------------------------
' EXPOSURE VALUE
'------------------------------------------------------------------------------
            If Issues > 0 Then
                ReDim Preserve Es(1 To Issues)
                ReDim Preserve Hs(1 To Issues)
                .NetExposure = Checked(SFT_NetExposure(Es, Hs), "SFT_NetExposure")
                .GrossExposure = Checked(SFT_GrossExposure(Es, Hs), "SFT_GrossExposure")
                .IssueCount = CLng(Checked(SFT_IssueCount(Es, pThreshold), "SFT_IssueCount"))
            End If
            .SystematicTerm = pNetW * .NetExposure
            If .IssueCount > 0 Then
                .IdiosyncraticTerm = pGrossW * .GrossExposure / Sqr(.IssueCount)
            End If
            .EStar = Checked(SFT_ExposureNetted(.SumE, .SumC, .NetExposure, .GrossExposure, .IssueCount, _
                                                .FXTerm, pNetW, pGrossW), "SFT_ExposureNetted")
        End With

End Sub


Private Sub FinishSingles()
'
'==============================================================================
'                                FinishSingles
'------------------------------------------------------------------------------
' PURPOSE
'   Calculate E* of every transaction of a VALID netting set without an
'   eligible master netting agreement with the single-transaction formula,
'   and add it to its netting set: E* of the netting set is the sum over
'   its transactions, and its holding period the longest of theirs.
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
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeAvg    As Double    'Value-weighted haircut of the legs lent
    Dim HcAvg    As Double    'Value-weighted haircut of the legs received
    Dim HfxAvg   As Double    'Value-weighted currency haircut of the legs received
    Dim Ns       As Long      'Netting set
    Dim t        As Long      'Transaction

'------------------------------------------------------------------------------
' EACH TRANSACTION
'------------------------------------------------------------------------------
        For t = 1 To mTxCount
            Ns = mTx(t).NsIndex
            If Ns > 0 Then
                If Not mNS(Ns).MNA And mNS(Ns).Status = "VALID" Then
                    With mTx(t)
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
        Next t

End Sub


Private Function Checked( _
    ByVal Result As Variant, _
    ByVal FunctionName As String) _
    As Double
'
'==============================================================================
'                                   Checked
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
' UPDATED
'   2026-10-24
'==============================================================================
'
        If IsError(Result) Then
            Err.Raise ERR_FORMULA_DOMAIN, "CORE_Engine.Checked", FunctionName & " rejected the inputs the " & _
                      "engine passed: " & CStr(Result)
        End If
        Checked = CDbl(Result)

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
'==============================================================================
'                              ClearOutputSheets
'------------------------------------------------------------------------------
' PURPOSE
'   Clear the Results, Positions and LegCalc tables before a run. Checks is
'   cleared and rewritten by WriteChecks.
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

End Sub


Private Sub WriteResults()
'
'==============================================================================
'                                 WriteResults
'------------------------------------------------------------------------------
' PURPOSE
'   Write one Results row per netting set, then the TOTAL CRR and TOTAL
'   BCBS rows over the VALID netting sets of each regime. A netting set
'   that is not VALID shows its identity and status only.
'
' UPDATED
'   2026-10-25
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet      As Worksheet    'Results sheet
    Dim TopRow     As Long         'First data row
    Dim Block()    As Variant      'Output rows
    Dim RowCount   As Long         'Rows written
    Dim Ns         As Long         'Netting set
    Dim r          As Long         'Regime index

'------------------------------------------------------------------------------
' NETTING SETS
'------------------------------------------------------------------------------
        Set Sheet = GetSheet(SH_RESULTS)
        TopRow = OutputFirstRow(SH_RESULTS)
        ReDim Block(1 To mNsCount + 2, 1 To RS_NCOLS)
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
        For r = 1 To 2
            RowCount = RowCount + 1
            Block(RowCount, 1) = "TOTAL " & Choose(r, RG_CRR, RG_BCBS)
            Block(RowCount, 3) = Choose(r, RG_CRR, RG_BCBS)
            Block(RowCount, 18) = mRegEStar(r)
            Block(RowCount, 22) = mRegRWA(r)
            Block(RowCount, 23) = mRegOwnFunds(r)
            Block(RowCount, 24) = mRegValid(r) & " of " & mRegSets(r) & " netting set(s) VALID; " & _
                              (mRegSets(r) - mRegValid(r)) & " withheld"
        Next r

'------------------------------------------------------------------------------
' WRITE AND FORMAT
'------------------------------------------------------------------------------
        WriteBlock Sheet, TopRow, Block, RowCount, RS_NCOLS
        FormatColumns Sheet, TopRow, RowCount, "@|@|@|@|0|0|0|0|0.0000|#,##0|#,##0|#,##0|#,##0|0|#,##0|#,##0|" & _
                      "#,##0|#,##0|@|@|0.00%|#,##0|#,##0|@"
        TableRange(Sheet, TopRow + RowCount - 2, 1, TopRow + RowCount - 1, RS_NCOLS).Font.Bold = True

End Sub


Private Sub WritePositions()
'
'==============================================================================
'                                WritePositions
'------------------------------------------------------------------------------
' PURPOSE
'   Write the Positions sheet: for a VALID netting set with netting, one row
'   per security issue and per currency; without netting, one row per
'   transaction not rejected, filed under its first leg's netting set.
'
' UPDATED
'   2026-10-26
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet      As Worksheet    'Positions sheet
    Dim TopRow     As Long         'First data row
    Dim Block()    As Variant      'Output rows
    Dim RowCount   As Long         'Rows written
    Dim Ns         As Long         'Netting set
    Dim p          As Long         'Position
    Dim t          As Long         'Transaction
    Dim NetPos     As Double       'Lent less received

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
        Set Sheet = GetSheet(SH_POSITIONS)
        TopRow = OutputFirstRow(SH_POSITIONS)
        ReDim Block(1 To mPosCount + mTxCount + 1, 1 To PS_NCOLS)
        For Ns = 1 To mNsCount
            If mNS(Ns).MNA Then
                For p = 1 To mPosCount
                    If mPos(p).NsIndex = Ns Then
                        RowCount = RowCount + 1
                        With mPos(p)
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
                Next p
            Else
                For t = 1 To mTxCount
                    If mTx(t).NsIndex = Ns And Not mTx(t).Rejected Then
                        RowCount = RowCount + 1
                        With mTx(t)
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
                Next t
            End If
        Next Ns

'------------------------------------------------------------------------------
' WRITE AND FORMAT
'------------------------------------------------------------------------------
        If RowCount = 0 Then
            Exit Sub
        End If
        WriteBlock Sheet, TopRow, Block, RowCount, PS_NCOLS
        FormatColumns Sheet, TopRow, RowCount, "@|@|@|@|@|0|0|#,##0|#,##0|#,##0|0.0000%|#,##0|#,##0|@|#,##0|" & _
                      "#,##0|#,##0|#,##0"

End Sub


Private Sub WriteLegCalc()
'
'==============================================================================
'                                 WriteLegCalc
'------------------------------------------------------------------------------
' PURPOSE
'   Write one LegCalc row per leg, in input order, with every intermediate
'   figure of its haircut; an excluded leg shows why.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sheet      As Worksheet    'LegCalc sheet
    Dim TopRow     As Long         'First data row
    Dim Block()    As Variant      'Output rows
    Dim i          As Long         'Leg

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
        If mLegCount = 0 Then
            Exit Sub
        End If
        Set Sheet = GetSheet(SH_LEGCALC)
        TopRow = OutputFirstRow(SH_LEGCALC)
        ReDim Block(1 To mLegCount, 1 To LC_NCOLS)
        For i = 1 To mLegCount
            With mLegs(i)
                Block(i, 1) = .LegID
                Block(i, 2) = .TxID
                If .NsIndex > 0 Then
                    Block(i, 3) = mNS(.NsIndex).ID
                    Block(i, 19) = mNS(.NsIndex).Regime
                End If
                Block(i, 5) = .TxType
                Block(i, 6) = .Side
                Block(i, 7) = .Asset
                Block(i, 20) = .Messages
                If Not .Valid Or .TxRejected Then
                    Block(i, 4) = LS_EXCLUDED
                Else
                    If .Recognised Then
                        Block(i, 4) = LS_OK
                    Else
                        Block(i, 4) = LS_NOTRECOG
                    End If
                    Block(i, 8) = .HaircutKey
                    Block(i, 9) = IIf(.Eligible, "Y", "N")
                    If .ResidualYears >= 0# Then
                        Block(i, 10) = .ResidualYears
                    End If
                    Block(i, 11) = .Value
                    Block(i, 12) = .Sign * .Value
                    If .Recognised And .HoldingBD > 0# Then
                        Block(i, 13) = .H10
                        Block(i, 14) = .HoldingBD
                        Block(i, 15) = .Scale
                        Block(i, 16) = .H
                        Block(i, 17) = .Hfx
                        Block(i, 18) = .Adjusted
                    End If
                End If
            End With
        Next i

'------------------------------------------------------------------------------
' WRITE AND FORMAT
'------------------------------------------------------------------------------
        WriteBlock Sheet, TopRow, Block, mLegCount, LC_NCOLS
        FormatColumns Sheet, TopRow, mLegCount, "@|@|@|@|@|@|@|@|@|0.00|#,##0|#,##0|0.00%|0|0.0000|0.0000%|" & _
                      "0.0000%|#,##0|@|@"

End Sub


Private Function WriteChecks( _
    ByRef Failure As String) _
    As Boolean
'
'==============================================================================
'                                 WriteChecks
'------------------------------------------------------------------------------
' PURPOSE
'   Write the logged messages to the Checks sheet, errors in bold, or a
'   single "No issues found." line.
'
' RETURNS
'   True when the sheet was written. False when writing failed, with the
'   reason in Failure.
'
' ERROR POLICY
'   Contains any error and reports it through the result, never silently.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ChecksSheet     As Worksheet    'Checks sheet
    Dim TopRow          As Long         'First Checks data row
    Dim OutputBlock()   As Variant      'Output rows, (message, field)
    Dim i               As Long         'Message
    Dim j               As Long         'Field

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
        On Error GoTo Failed
        Failure = ""
        Set ChecksSheet = GetSheet(SH_CHECKS)
        TopRow = OutputFirstRow(SH_CHECKS)
        ClearOutputBlock ChecksSheet, TopRow, CK_NCOLS
        If mLogCount = 0 Then
            ReDim OutputBlock(1 To 1, 1 To CK_NCOLS)
            OutputBlock(1, 1) = SEV_INFO
            OutputBlock(1, 5) = "No issues found."
            WriteBlock ChecksSheet, TopRow, OutputBlock, 1, CK_NCOLS
            WriteChecks = True
            Exit Function
        End If
        ReDim OutputBlock(1 To mLogCount, 1 To CK_NCOLS)
        For i = 1 To mLogCount
            For j = 1 To CK_NCOLS
                OutputBlock(i, j) = mLog(j, i)
            Next j
        Next i
        WriteBlock ChecksSheet, TopRow, OutputBlock, mLogCount, CK_NCOLS
        For i = 1 To mLogCount
            If OutputBlock(i, 1) = SEV_ERROR Then
                TableCell(ChecksSheet, TopRow + i - 1, 1).Font.Bold = True
            End If
        Next i
        WriteChecks = True
        Exit Function

'------------------------------------------------------------------------------
' REPORT FAILURE
'------------------------------------------------------------------------------
Failed:
        Failure = "error " & Err.Number & ": " & Err.Description

End Function


Private Sub WriteRunInfo( _
    ByVal RunSeconds As Double, _
    ByVal Completed As Boolean, _
    ByVal WroteOutputs As Boolean, _
    ByVal Failure As String)
'
'==============================================================================
'                                 WriteRunInfo
'------------------------------------------------------------------------------
' PURPOSE
'   Write a one-line run summary to the Results cell named NM_RUNINFO: a
'   failure, a validation-only run, a completed run with its counts, run
'   class and duration, or a run stopped by input errors.
'
' ERROR POLICY
'   Never raises: the summary is informative. A protected or missing
'   Results sheet, or a missing name, keeps its earlier summary.
'
' UPDATED
'   2026-10-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Summary   As String    'Summary line
    Dim Stamp     As String    'Date and time of the run

'------------------------------------------------------------------------------
' COMPOSE
'------------------------------------------------------------------------------
        Stamp = Format$(Now, "yyyy-mm-dd hh:mm:ss")
        If Len(Failure) > 0 Then
            Summary = "Last run " & Stamp & " FAILED - " & Failure
        ElseIf Not WroteOutputs Then
            Summary = "Last validation " & Stamp & " | errors " & mErrCount & ", warnings " & mWarnCount & _
                      RunClassText() & " | outputs cleared: press Run SFT to calculate"
        ElseIf Completed Then
            Summary = "Last run " & Stamp & " | reporting date " & Format$(CDate(pAsOf), "yyyy-mm-dd") & _
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
        On Error Resume Next
        NamedCell(NM_RUNINFO).Value = Summary
        On Error GoTo 0

End Sub


Private Function RunClassText() As String
'
'==============================================================================
'                                 RunClassText
'------------------------------------------------------------------------------
' PURPOSE
'   The run class for the run summary: the regulatory set's identity, or
'   WHAT-IF with the number of differences; "" when not classified.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If mRunClass = RUN_REGULATORY Then
            RunClassText = " | regulatory parameters " & REF_SET_ID
        ElseIf mRunClass = RUN_WHATIF Then
            RunClassText = " | WHAT-IF: " & mDeviations & " difference(s) from " & REF_SET_ID
        End If

End Function
