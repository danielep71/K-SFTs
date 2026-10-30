Attribute VB_Name = "CORE_Config"
'==============================================================================
' MODULE: CORE_Config
'------------------------------------------------------------------------------
' PURPOSE
'   Hold every layout constant the SFT engine depends on: sheet names, the
'   workbook names that locate each table, the column of each input field,
'   the width of each output table, parameter codes, the regulatory
'   reference parameter set, regime, transaction, asset and counterparty
'   codes, the expected column headers, the severity labels written to the
'   Checks sheet and the error numbers raised by the workbook macros. When a
'   sheet layout changes, it changes here and nowhere else.
'
' PUBLIC SURFACE
'   None outside this VBA project. The constants are Public for in-project use
'   by CORE_Util, CORE_Engine and M_Main; Option Private Module keeps them
'   off the supported external surface.
'
' DEPENDENCIES
'   None. Every value is a literal.
'
' STATE OWNERSHIP
'   Constants only; no state.
'
' ERROR POLICY
'   Not applicable.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' UPDATED
'   2026-10-29
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    'Require explicit declarations; keep the constants inside this project.
    Option Explicit
    Option Private Module

'------------------------------------------------------------------------------
' SHEET NAMES
'------------------------------------------------------------------------------
    'Worksheet tab names, as the engine looks them up in ThisWorkbook.
        Public Const SH_README      As String = "README"          'Read-me; not used by the engine
        Public Const SH_PARAMS      As String = "Params"          'Parameters, haircut, risk-weight and FX tables
        Public Const SH_NS          As String = "NettingSets"     'Netting-set inputs
        Public Const SH_TX          As String = "Transactions"    'Transaction legs
        Public Const SH_RESULTS     As String = "Results"         'One row per netting set, and the totals
        Public Const SH_POSITIONS   As String = "Positions"       'Issues, currencies and single transactions
        Public Const SH_LEGCALC     As String = "LegCalc"         'One row per leg
        Public Const SH_CHECKS      As String = "Checks"          'Errors, warnings, information

'------------------------------------------------------------------------------
' COMMON LAYOUT
'------------------------------------------------------------------------------
    'Every sheet has two narrow margin columns (A:B); rows 1:2 hold the
    'title (K-SFTs in B1, the sheet title in B2). Every table starts in
    'column C and is located through the workbook name of its first header
    'cell, with data from the next row (CORE_Util.TableHeaderRow,
    'TableFirstRow and, for the output tables, OutputFirstRow), so its rows
    'can move. Column numbers in this module (NS_, TX_, RS_ and the others)
    'count from the first table column; CORE_Util.TableCell and TableRange
    'add COL_OFFSET. The run summary is written to the cell of the workbook
    'name NM_RUNINFO on Results.
        Public Const COL_OFFSET       As Long = 2         'Margin columns left of every table
        Public Const NM_RUNINFO       As String = "Results_RunInfo"        'Results cell receiving the run summary line
        Public Const NM_NS_HEADER     As String = "NettingSets_Header"     'NettingSets header cell "NettingSetID"
        Public Const NM_TX_HEADER     As String = "Transactions_Header"    'Transactions header cell "LegID"
        Public Const NM_RS_HEADER     As String = "Results_Header"         'Results header cell "NettingSetID"
        Public Const NM_PS_HEADER     As String = "Positions_Header"       'Positions header cell "NettingSetID"
        Public Const NM_LC_HEADER     As String = "LegCalc_Header"         'LegCalc header cell "LegID"
        Public Const NM_CK_HEADER     As String = "Checks_Header"          'Checks header cell "Severity"

'------------------------------------------------------------------------------
' PARAMS SHEET
'------------------------------------------------------------------------------
    'Each parameter is a row with its code under "Code" and its value under
    '"Value", read through the workbook name equal to its code
    '(CORE_Util.GetParam). The top-left cells of NM_HC_TABLE, NM_RW_TABLE
    'and NM_FX_TABLE are the header cells of the three lookup tables, each
    'running down to its first blank key. BD = business days; H10 = haircut
    'for a 10-business-day holding period with daily remargining.
        Public Const NM_PRM_CODE      As String = "Params_Code"               'Header cell "Code"
        Public Const NM_PRM_VALUE     As String = "Params_Value"              'Header cell "Value"
        Public Const NM_HC_TABLE      As String = "HC_Table"                  'Haircut table, from its header
        Public Const NM_RW_TABLE      As String = "RW_Table"                  'Risk-weight table, from its header
        Public Const NM_FX_TABLE      As String = "FX_Table"                  'FX table, from its header
        Public Const PRM_ASOF         As String = "AsOfDate"                  'Reporting date
        Public Const PRM_REPCCY       As String = "ReportingCcy"              'Reporting currency
        Public Const PRM_REGIME       As String = "Regime"                    'Default regime
        Public Const PRM_DAYSYEAR     As String = "DaysPerYear"               'Days per year, residual maturity
        Public Const PRM_HCBASE       As String = "HaircutBaseBD"             'Holding period of the H10 table, BD
        Public Const PRM_HOLD_REPO    As String = "HoldingRepoBD"             'Minimum holding period, repo-style, BD
        Public Const PRM_HOLD_CMT     As String = "HoldingCapitalMarketBD"    'Minimum holding period, other CMT, BD
        Public Const PRM_HOLD_LARGE   As String = "HoldingLargeBD"            'Holding period, large or illiquid, BD
        Public Const PRM_HFX          As String = "FXHaircut"                 'Currency-mismatch haircut, 10 BD
        Public Const PRM_NONELIG      As String = "NonEligibleHaircut"        'Haircut of non-eligible securities lent, 10 BD
        Public Const PRM_NETW         As String = "NetWeight"                 'Weight of the net exposure
        Public Const PRM_GROSSW       As String = "GrossWeight"               'Weight of the gross exposure
        Public Const PRM_NTHRESH      As String = "IssueCountThreshold"       'Share of the largest issue counted in N
        Public Const PRM_CAPRATIO     As String = "CapitalRatio"              'Own funds requirement per unit of RWA

'------------------------------------------------------------------------------
' REGULATORY PARAMETER SET
'------------------------------------------------------------------------------
    'The values a regulatory run uses, as listed in the methodology: CRE22
    '(effective 1 January 2023) and the consolidated CRR of 26.06.2026,
    'transcribed on 2026-10-24. Only the haircut table is checked against
    'the registered copies (#197; PROTOTYPE.md). A run whose parameters,
    'haircut table, risk-weight table and risk-weight overrides all agree
    'with this set is REGULATORY and names REF_SET_ID; any difference makes
    'it WHAT-IF (CORE_Engine.ClassifyRun). DaysPerYear is a convention and
    'does not change the class. REF_HAIRCUTS and REF_RISKWEIGHTS list each
    'table row as KEY=value, separated by ";".
        Public Const REF_SET_ID       As String = "CRE22-CRR-2026-10-24"    'Identity of the set
        Public Const RUN_REGULATORY   As String = "REGULATORY"              'Every parameter as the set
        Public Const RUN_WHATIF       As String = "WHAT-IF"                 'Some parameter differs
        Public Const REF_HCBASE       As Double = 10      'CRE22.49; CRR Art. 224 Table 1 (10-day column)
        Public Const REF_HOLD_REPO    As Double = 5       'CRE22.57; CRR Art. 224(2)
        Public Const REF_HOLD_CMT     As Double = 10      'CRE22.57; CRR Art. 224(2)
        Public Const REF_HOLD_LARGE   As Double = 20      'CRE22.58; CRR Art. 224(2), 285(3)
        Public Const REF_HFX          As Double = 0.08    'CRE22.52; CRR Art. 224 Table 4
        Public Const REF_NONELIG      As Double = 0.3     'CRE22.54; CRR Art. 224(4)
        Public Const REF_NETW         As Double = 0.4     'CRE22.64-22.65; CRR Art. 220(3)
        Public Const REF_GROSSW       As Double = 0.6     'CRE22.64-22.65; CRR Art. 220(3)
        Public Const REF_NTHRESH      As Double = 0.1     'CRE22.65; CRR Art. 220(3)
        Public Const REF_CAPRATIO     As Double = 0.08    'RBC20.1; CRR Art. 92(1)(c)
        Public Const REF_HAIRCUTS As String = "CASH=0;GOLD=0.2;EQUITY_MAIN=0.2;EQUITY_LISTED=0.3;" & _
            "DEBT_SOV_CQS1_LE1=0.005;DEBT_SOV_CQS1_1TO3=0.02;DEBT_SOV_CQS1_3TO5=0.02;" & _
            "DEBT_SOV_CQS1_5TO10=0.04;DEBT_SOV_CQS1_GT10=0.04;" & _
            "DEBT_OTH_CQS1_LE1=0.01;DEBT_OTH_CQS1_1TO3=0.03;DEBT_OTH_CQS1_3TO5=0.04;" & _
            "DEBT_OTH_CQS1_5TO10=0.06;DEBT_OTH_CQS1_GT10=0.12;" & _
            "DEBT_SEC_CQS1_LE1=0.02;DEBT_SEC_CQS1_1TO3=0.08;DEBT_SEC_CQS1_3TO5=0.08;" & _
            "DEBT_SEC_CQS1_5TO10=0.16;DEBT_SEC_CQS1_GT10=0.16;" & _
            "DEBT_SOV_CQS23_LE1=0.01;DEBT_SOV_CQS23_1TO3=0.03;DEBT_SOV_CQS23_3TO5=0.03;" & _
            "DEBT_SOV_CQS23_5TO10=0.06;DEBT_SOV_CQS23_GT10=0.06;" & _
            "DEBT_OTH_CQS23_LE1=0.02;DEBT_OTH_CQS23_1TO3=0.04;DEBT_OTH_CQS23_3TO5=0.06;" & _
            "DEBT_OTH_CQS23_5TO10=0.12;DEBT_OTH_CQS23_GT10=0.2;" & _
            "DEBT_SEC_CQS23_LE1=0.04;DEBT_SEC_CQS23_1TO3=0.12;DEBT_SEC_CQS23_3TO5=0.12;" & _
            "DEBT_SEC_CQS23_5TO10=0.24;DEBT_SEC_CQS23_GT10=0.24;" & _
            "DEBT_SOV_CQS4_ALL=0.15"
        Public Const REF_RISKWEIGHTS As String = "SOV_CQS1=0;SOV_CQS2=0.2;SOV_CQS3=0.5;SOV_CQS4=1;" & _
            "SOV_CQS5=1;SOV_CQS6=1.5;SOV_UNRATED=1;" & _
            "INST_CQS1=0.2;INST_CQS2=0.3;INST_CQS3=0.5;INST_CQS4=1;INST_CQS5=1;INST_CQS6=1.5;" & _
            "INST_ST_CQS1=0.2;INST_ST_CQS2=0.2;INST_ST_CQS3=0.2;INST_ST_CQS4=0.5;INST_ST_CQS5=0.5;" & _
            "INST_ST_CQS6=1.5;" & _
            "INST_SCRA_A=0.4;INST_SCRA_A_ENH=0.3;INST_SCRA_B=0.75;INST_SCRA_C=1.5;" & _
            "INST_ST_SCRA_A=0.2;INST_ST_SCRA_A_ENH=0.2;INST_ST_SCRA_B=0.5;INST_ST_SCRA_C=1.5;" & _
            "CORP_CQS1=0.2;CORP_CQS2=0.5;CORP_CQS3=0.75;CORP_CQS4=1;CORP_CQS5=1.5;CORP_CQS6=1.5;" & _
            "CORP_UNRATED=1;QCCP=0.02"

'------------------------------------------------------------------------------
' REGIMES
'------------------------------------------------------------------------------
    'Regime codes accepted on Params and as a netting-set override; a row
    'of the haircut or risk-weight table applies to BOTH or to one regime.
        Public Const RG_BCBS   As String = "BCBS"    'Basel framework, CRE22 and CRE20
        Public Const RG_CRR    As String = "CRR"     'EU Capital Requirements Regulation
        Public Const RG_BOTH   As String = "BOTH"    'Table row valid in both regimes

'------------------------------------------------------------------------------
' TRANSACTION CODES
'------------------------------------------------------------------------------
    'Transaction types. The first four are repo-style transactions, the
    'last an other capital-market-driven transaction [CRE22.56-22.57; CRR
    'Art. 224(2)]. Secured lending is out of scope.
        Public Const TT_REPO        As String = "REPO"          'Securities sold under repurchase, cash received
        Public Const TT_REVREPO     As String = "REVREPO"       'Cash lent, securities bought under resale
        Public Const TT_SECLEND     As String = "SECLEND"       'Securities lent against collateral
        Public Const TT_SECBORROW   As String = "SECBORROW"     'Securities borrowed against collateral posted
        Public Const TT_MARGINLEND  As String = "MARGINLEND"    'Margin lending against securities
    'Side of a leg: what the bank gives (E) or receives (C).
        Public Const SD_LENT        As String = "LENT"          'Lent, sold or posted by the bank: E
        Public Const SD_RECEIVED    As String = "RECEIVED"      'Borrowed, bought or held by the bank: C
    'Asset types of a leg.
        Public Const AT_CASH        As String = "CASH"          'Cash
        Public Const AT_DEBT        As String = "DEBT"          'Debt security
        Public Const AT_EQUITY      As String = "EQUITY"        'Equity or convertible bond
        Public Const AT_GOLD        As String = "GOLD"          'Gold bullion
        Public Const AT_FUND        As String = "FUND"          'Unit of a fund (CIU / UCITS)
        Public Const AT_OTHER       As String = "OTHER"         'Any other asset: not eligible
    'Issuer types of a debt security, and the haircut-table category each
    'maps to: SOV to SOV, INST and CORP to OTH, SEC to SEC.
        Public Const IS_SOV         As String = "SOV"           'Central government or central bank
        Public Const IS_INST        As String = "INST"          'Institution
        Public Const IS_CORP        As String = "CORP"          'Corporate and other issuers
        Public Const IS_SEC         As String = "SEC"           'Securitisation position
    'Equity listing.
        Public Const EL_MAIN        As String = "MAIN"          'In a main index
        Public Const EL_LISTED      As String = "LISTED"        'Listed on a recognised exchange
        Public Const EL_OTHER       As String = "OTHER"         'Neither: not eligible
    'Months within which a transaction is short-term for the institution
    'risk weights [CRE20.19, CRE20.21; CRR Art. 120(2), 121(3)(a)].
        Public Const SHORT_TERM_MONTHS As Long = 3

'------------------------------------------------------------------------------
' COUNTERPARTY CODES
'------------------------------------------------------------------------------
    'Exposure class of the counterparty, the first part of a risk-weight key.
        Public Const CC_SOV         As String = "SOV"           'Central government or central bank
        Public Const CC_INST        As String = "INST"          'Institution
        Public Const CC_CORP        As String = "CORP"          'Corporate
        Public Const CC_QCCP        As String = "QCCP"          'Qualifying CCP, trade exposure
    'SCRA grades of an unrated institution; A_ENH is grade A with a CET1
    'ratio of at least 14% and a leverage ratio of at least 5%.
        Public Const SCRA_GRADES    As String = "A A_ENH B C"

'------------------------------------------------------------------------------
' LEG AND POSITION STATUS
'------------------------------------------------------------------------------
        Public Const LS_OK          As String = "OK"                'Leg used
        Public Const LS_EXCLUDED    As String = "EXCLUDED"          'Leg or its transaction rejected
        Public Const LS_NOTRECOG    As String = "NOT RECOGNISED"    'Non-eligible collateral received
        Public Const PK_ISSUE       As String = "ISSUE"             'Positions: one security issue
        Public Const PK_CURRENCY    As String = "CURRENCY"          'Positions: one currency
        Public Const PK_TRANSACTION As String = "TRANSACTION"       'Positions: one transaction without netting

'------------------------------------------------------------------------------
' PARAMS TABLE HEADERS
'------------------------------------------------------------------------------
    'Column positions and header texts of each lookup table, from the
    'top-left cell of its workbook name.
        Public Const HC_KEY        As Long = 1    'Haircut key
        Public Const HC_H10        As Long = 6    'Haircut H10, decimal
        Public Const HC_REGIMES    As Long = 7    'BOTH, CRR or BCBS
        Public Const RW_KEY        As Long = 1    'Risk-weight key
        Public Const RW_WEIGHT     As Long = 4    'Risk weight, decimal
        Public Const RW_REGIMES    As Long = 5    'BOTH, CRR or BCBS
        Public Const HC_HEADERS As String = "HC_Key|Asset type|Issuer category|Credit quality|" & _
            "Residual maturity|Haircut H10|Regimes"
        Public Const RW_HEADERS As String = "RW_Key|Exposure class|Grade|Risk weight|Regimes"
        Public Const FX_HEADERS As String = "FX_Ccy||Units of reporting ccy per 1 unit"
        Public Const PRM_CODE_HEADER  As String = "Code"
        Public Const PRM_VALUE_HEADER As String = "Value"

'------------------------------------------------------------------------------
' NETTINGSETS INPUT COLUMNS
'------------------------------------------------------------------------------
    'Column number of each field on the NettingSets sheet.
        Public Const NS_ID         As Long = 1     'Netting-set ID
        Public Const NS_CPTY       As Long = 2     'Counterparty
        Public Const NS_MNA        As Long = 3     'Eligible master netting agreement (Y/N)
        Public Const NS_SETTLE     As Long = 4     'Settlement currency; blank: reporting currency
        Public Const NS_CLASS      As Long = 5     'Counterparty class: SOV, INST, CORP or QCCP
        Public Const NS_CQS        As Long = 6     'Counterparty credit quality step 1-6; blank: unrated
        Public Const NS_SCRA       As Long = 7     'SCRA grade of an unrated institution
        Public Const NS_FREQ       As Long = 8     'Remargining frequency N_R, business days
        Public Const NS_LARGE      As Long = 9     'Over 5,000 trades or illiquid collateral (Y/N)
        Public Const NS_DISPUTE    As Long = 10    'Margin disputes (Y/N)
        Public Const NS_HOLDOVR    As Long = 11    'Holding period override, business days
        Public Const NS_ZEROHC     As Long = 12    '0% haircut conditions met (Y/N)
        Public Const NS_RWOVR      As Long = 13    'Risk weight override, decimal
        Public Const NS_REGIME     As Long = 14    'Regime override; blank uses Params
        Public Const NS_NCOLS      As Long = 14    'Number of input columns
        Public Const NS_HEADERS As String = "NettingSetID|Counterparty|" & _
            "Eligible master netting agreement (Y/N)|Settlement ccy (blank = reporting)|" & _
            "Counterparty class (SOV / INST / CORP / QCCP)|Counterparty CQS (1-6, blank = unrated)|" & _
            "SCRA grade (A / A_ENH / B / C)|Remargin frequency N_R (BD)|" & _
            ">5,000 trades or illiquid collateral (Y/N)|Margin disputes (Y/N)|" & _
            "Holding period override (BD)|0% haircut conditions met (Y/N)|Risk weight override|" & _
            "Regime override (blank = Params)"

'------------------------------------------------------------------------------
' TRANSACTIONS INPUT COLUMNS
'------------------------------------------------------------------------------
    'Column number of each field on the Transactions sheet. Each row is a
    'leg: cash or one security issue given (LENT) or received (RECEIVED).
    'Legs sharing a Transaction ID form one transaction and must agree on
    'its netting set, type, start and maturity dates.
        Public Const TX_LEG        As Long = 1     'Leg ID, unique
        Public Const TX_ID         As Long = 2     'Transaction ID
        Public Const TX_NS         As Long = 3     'Netting-set ID
        Public Const TX_TYPE       As Long = 4     'Transaction type: REPO, REVREPO, SECLEND, SECBORROW, MARGINLEND
        Public Const TX_START      As Long = 5     'Start (trade) date
        Public Const TX_MAT        As Long = 6     'Maturity date; blank: open
        Public Const TX_SIDE       As Long = 7     'LENT or RECEIVED
        Public Const TX_ASSET      As Long = 8     'Asset type
        Public Const TX_ISSUE      As Long = 9     'Issue ID: ISIN or internal code; blank for cash
        Public Const TX_ISSUER     As Long = 10    'Issuer type of a debt security
        Public Const TX_CQS        As Long = 11    'Credit quality step of the issue, 1-6; blank: unrated
        Public Const TX_SECMAT     As Long = 12    'Maturity date of a debt security
        Public Const TX_LISTING    As Long = 13    'Equity listing: MAIN, LISTED or OTHER
        Public Const TX_FUNDH      As Long = 14    'Fund haircut H10, decimal
        Public Const TX_MV         As Long = 15    'Market value, in the leg currency
        Public Const TX_CCY        As Long = 16    'Currency of the leg
        Public Const TX_COMMENT    As Long = 17    'Free-text comment, not read
        Public Const TX_NCOLS      As Long = 17    'Number of input columns
        Public Const TX_HEADERS As String = "LegID|TransactionID|NettingSetID|" & _
            "Transaction type|Start date|Maturity date (blank = open)|Side (LENT / RECEIVED)|" & _
            "Asset type|Issue ID|Issuer type (SOV / INST / CORP / SEC)|Issue CQS (1-6, blank = unrated)|" & _
            "Security maturity date|Equity listing (MAIN / LISTED / OTHER)|Fund haircut H10|" & _
            "Market value|Currency|"

'------------------------------------------------------------------------------
' OUTPUT TABLE COLUMNS
'------------------------------------------------------------------------------
    'Number of columns the engine clears and writes on each output sheet,
    'and the position of the main Results columns (WriteResults writes them
    'by literal index; no procedure reads these constants).
        Public Const RS_NCOLS       As Long = 24    'Results
        Public Const RS_ID          As Long = 1     'Results: NettingSetID, or TOTAL CRR / TOTAL BCBS
        Public Const RS_REGIME      As Long = 3     'Results: BCBS or CRR
        Public Const RS_ESTAR       As Long = 18    'Results: E*, the exposure value
        Public Const RS_RWA         As Long = 22    'Results: risk-weighted exposure amount
        Public Const RS_OWNFUNDS    As Long = 23    'Results: own funds requirement
        Public Const RS_STATUS      As Long = 24    'Results: status; on a TOTAL row, the VALID and withheld count
        Public Const PS_NCOLS       As Long = 18    'Positions
        Public Const LC_NCOLS       As Long = 20    'LegCalc
        Public Const CK_NCOLS       As Long = 5     'Checks

'------------------------------------------------------------------------------
' ERROR NUMBERS
'------------------------------------------------------------------------------
    'Errors raised by the workbook macros, in the project range
    'vbObjectError + 2048 to + 4095.
        Public Const ERR_RUN_ACTIVE       As Long = vbObjectError + 2049    'An operation is already running
        Public Const ERR_CLEANUP_FAILED   As Long = vbObjectError + 2050    'Excel settings were not restored
        Public Const ERR_CHECKS_WRITE     As Long = vbObjectError + 2053    'Checks sheet could not be written
        Public Const ERR_FORMULA_DOMAIN   As Long = vbObjectError + 2055    'SFT_* function rejected engine inputs
        Public Const ERR_RECALC_PENDING   As Long = vbObjectError + 2056    'Inputs not recalculated before a run
        Public Const ERR_OUTPUT_LAYOUT    As Long = vbObjectError + 2058    'Output table header name missing

'------------------------------------------------------------------------------
' SEVERITY LABELS
'------------------------------------------------------------------------------
    'Severity written in the first column of the Checks sheet.
        Public Const SEV_ERROR   As String = "ERROR"      'Input rejected or run stopped
        Public Const SEV_WARN    As String = "WARNING"    'Input accepted with an assumption
        Public Const SEV_INFO    As String = "INFO"       'Information only
