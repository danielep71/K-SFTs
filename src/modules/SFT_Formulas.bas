Attribute VB_Name = "SFT_Formulas"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: SFT_Formulas
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the building blocks of the exposure value of securities
'   financing transactions (SFTs) under the financial collateral
'   comprehensive method, with supervisory haircuts, as public functions
'
' WHY THIS EXISTS
'   The engine calls these functions, and they can also be entered in
'   worksheet cells as user-defined functions, so every intermediate figure
'   can be checked independently
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - haircuts: the holding period (SFT_HoldingPeriod) and the haircut
'       scaled to it (SFT_ScaledHaircut)
'     - exposure value: one transaction (SFT_ExposureSingle) and a netting
'       set (SFT_NetExposure, SFT_GrossExposure, SFT_IssueCount,
'       SFT_ExposureNetted)
'     - capital: the own funds requirement (SFT_CapitalRequirement)
'     - private helpers that read ranges and arrays (ReadPairs, Flatten,
'       ElementIsNumber)
'
' ERROR POLICY
'   Each function checks the domain of its inputs, stated in its INPUTS
'   section, and returns a worksheet error value instead of raising or
'   returning a misleading number: #NUM! for a number outside its domain,
'   #VALUE! for a value that is not a number where one is needed, #REF! for
'   ranges of different sizes, and the error itself for an error value in a
'   range. A VBA caller receives the same error value in the Variant result
'   and tests it with IsError; CORE_Engine stops the run when it receives
'   one. An overflow beyond the range of a Double also gives #NUM!. A
'   Boolean is not a number: TRUE or FALSE in a range gives #VALUE!.
'
' DEPENDENCIES
'   VBA runtime and the Excel error constants (xlErrNum, xlErrValue,
'   xlErrRef); no WorksheetFunction, so results do not depend on the Excel
'   version or the locale.
'
' NOTES
'   This module is public (no Option Private Module): its functions are
'   available to worksheet cells.
'
' PUBLIC SURFACE
'   The eight SFT_* functions, listed in docs/PUBLIC_API.txt. Their names,
'   arguments, defaults and results are a contract: changing one is a public
'   API change.
'
' STATE OWNERSHIP
'   Stateless. Results depend only on the arguments.
'
' REFERENCE
'   BCBS CRE22 (paragraph numbers in brackets); EU CRR Art. 220, 223, 224,
'   226; risk weights per CRE20 and CRR Part Three Title II Chapter 2.
'
' COMPATIBILITY
'   Excel VBA; no references beyond the defaults.
'
' AUTHOR
'   Daniele Penza
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------


'
'------------------------------------------------------------------------------
'
'                                   HAIRCUTS
'
'------------------------------------------------------------------------------
'

Public Function SFT_HoldingPeriod( _
    ByVal MinimumBD As Double, _
    Optional ByVal LargeOrIlliquid As Boolean = False, _
    Optional ByVal Disputes As Boolean = False, _
    Optional ByVal OverrideBD As Double = 0#, _
    Optional ByVal LargeBD As Double = 20#) _
    As Variant
'
'------------------------------------------------------------------------------
'                            RESOLVE HOLDING PERIOD
'------------------------------------------------------------------------------
' PURPOSE
'   Holding period T_M of a transaction or netting set, in business days:
'     T = MinimumBD, at least LargeBD when LargeOrIlliquid,
'     at least OverrideBD when one is given, then doubled for Disputes.
'
' WHY THIS EXISTS
'   The supervisory haircuts are scaled to the holding period (see
'   SFT_ScaledHaircut), which depends on the transaction type, the size and
'   liquidity of the netting set, supervisory overrides and past disputes
'
' INPUTS
'   MinimumBD
'     Minimum holding period of the transaction type, business days, 1 or
'     more: 5 for repo-style transactions, 10 for other
'     capital-market-driven transactions; the longest when a netting set
'     holds both.
'
'   LargeOrIlliquid
'     True for a netting set with more than 5,000 transactions in the
'     previous quarter or with illiquid collateral.
'
'   Disputes
'     True after more than two margin-call disputes in the previous two
'     quarters lasting longer than the margin period of risk estimate.
'
'   OverrideBD
'     A longer holding period required by the supervisor or the bank,
'     business days, 0 or more; 0 for none. It never shortens T.
'
'   LargeBD
'     The holding period for LargeOrIlliquid, 1 or more; 20 by default.
'
' RETURNS
'   T_M in business days; #NUM! for an input outside its domain or a
'   doubling beyond the range of a Double.
'
' BEHAVIOR
'   Validates the inputs, starts from MinimumBD, raises it to LargeBD when
'   LargeOrIlliquid and to OverrideBD when that is longer, then doubles it
'   for Disputes. The floors are applied before the doubling, which applies
'   to the holding period that would otherwise be used.
'
' ERROR POLICY
'   Never raises: returns #NUM! for an input outside its domain (checked
'   before the overflow handler is enabled) and for an overflow
'
' DEPENDENCIES
'   VBA runtime (CVErr); Excel error constant xlErrNum
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.57-22.58; CRR Art. 224(2), Art. 285(3)-(4). The longest minimum
'   for a netting set holding both types is in CRE22.58 only.
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "SFT_Formulas.SFT_HoldingPeriod"

    Dim Holding                 As Double       'Holding period being built, business days

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Return #NUM! for an input outside its domain
        If MinimumBD < 1# Or OverrideBD < 0# Or LargeBD < 1# Then
            SFT_HoldingPeriod = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn an overflow into #NUM!
        On Error GoTo Overflowed

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Start from the minimum holding period of the transaction type
        Holding = MinimumBD
    'Apply the floor for a large or illiquid netting set
        If LargeOrIlliquid And LargeBD > Holding Then
            Holding = LargeBD
        End If
    'Apply a longer holding period required by the supervisor or the bank
        If OverrideBD > Holding Then
            Holding = OverrideBD
        End If
    'Double the floored holding period after margin-call disputes
        If Disputes Then
            Holding = 2# * Holding
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the holding period
        SFT_HoldingPeriod = Holding

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Overflowed:
    'Return #NUM! for an overflow
        SFT_HoldingPeriod = CVErr(xlErrNum)

End Function

Public Function SFT_ScaledHaircut( _
    ByVal H10 As Double, _
    ByVal HoldingBD As Double, _
    Optional ByVal RemarginBD As Double = 1#, _
    Optional ByVal BaseBD As Double = 10#) _
    As Variant
'
'------------------------------------------------------------------------------
'                          SCALE SUPERVISORY HAIRCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Haircut for the holding period and remargining frequency of the
'   transaction, from the supervisory haircut of the table:
'     H = H10 * sqrt((N_R + T_M - 1) / BaseBD)
'
' WHY THIS EXISTS
'   The supervisory haircuts of the table assume a holding period of BaseBD
'   business days with daily remargining; a transaction with another
'   holding period or remargining frequency needs the scaled haircut
'
' INPUTS
'   H10
'     Supervisory haircut for BaseBD business days with daily remargining,
'     decimal, 0 to 1.
'
'   HoldingBD
'     Holding period T_M, business days, 1 or more.
'
'   RemarginBD
'     Business days N_R between remargining (or revaluation), 1 or more; 1
'     (daily) by default.
'
'   BaseBD
'     Holding period of the table, business days, above 0; 10 by default.
'
' RETURNS
'   The scaled haircut, decimal; #NUM! for an input outside its domain or an
'   overflow.
'
' BEHAVIOR
'   Validates the inputs, then applies the square-root-of-time formula
'
' ERROR POLICY
'   Never raises: returns #NUM! for an input outside its domain and for an
'   overflow
'
' DEPENDENCIES
'   VBA runtime (CVErr, Sqr); Excel error constant xlErrNum
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.59; CRR Art. 226, which gives the same figure, to the tables'
'   rounding, from the column of Art. 224 Table 1 for the liquidation
'   period, since that column is H10 scaled by sqrt(T_M / 10). The CRR has
'   columns for 5, 10 and 20 days only; other periods extend the formula.
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "SFT_Formulas.SFT_ScaledHaircut"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn an overflow into #NUM!
        On Error GoTo Overflowed

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Return #NUM! for an input outside its domain
        If H10 < 0# Or H10 > 1# Or HoldingBD < 1# Or RemarginBD < 1# Or BaseBD <= 0# Then
            SFT_ScaledHaircut = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Scale the haircut to the holding period and remargining frequency
        SFT_ScaledHaircut = H10 * Sqr((RemarginBD + HoldingBD - 1#) / BaseBD)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Overflowed:
    'Return #NUM! for an overflow
        SFT_ScaledHaircut = CVErr(xlErrNum)

End Function


'
'------------------------------------------------------------------------------
'
'                               EXPOSURE VALUE
'
'------------------------------------------------------------------------------
'

Public Function SFT_ExposureSingle( _
    ByVal E As Double, _
    ByVal He As Double, _
    ByVal C As Double, _
    ByVal Hc As Double, _
    Optional ByVal Hfx As Double = 0#) _
    As Variant
'
'------------------------------------------------------------------------------
'                     COMPUTE EXPOSURE OF ONE TRANSACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Fully adjusted exposure value of one collateralised transaction:
'     E* = max(0, E * (1 + He) - C * (1 - Hc - Hfx))
'
' WHY THIS EXISTS
'   A transaction outside an eligible master netting agreement takes the
'   exposure value of the single-transaction formula
'
' INPUTS
'   E
'     Current value of what the bank lent, sold or posted, 0 or more.
'
'   He
'     Haircut of the exposure, scaled, 0 or more; for several items, the
'     value-weighted average of their haircuts [CRE22.55].
'
'   C
'     Current value of the collateral received, 0 or more.
'
'   Hc
'     Haircut of the collateral, scaled, 0 or more.
'
'   Hfx
'     Currency-mismatch haircut, scaled, 0 or more; 0 when the collateral
'     is in the currency of the exposure.
'
' RETURNS
'   E*, in the unit of E and C; #NUM! for an input outside its domain or
'   an overflow.
'
' BEHAVIOR
'   Validates the inputs, applies the formula and floors the result at 0
'
' ERROR POLICY
'   Never raises: returns #NUM! for an input outside its domain and for an
'   overflow
'
' DEPENDENCIES
'   VBA runtime (CVErr); Excel error constant xlErrNum
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.46; CRR Art. 223(2), (3), (5).
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "SFT_Formulas.SFT_ExposureSingle"

    Dim Adjusted                As Double       'E(1 + He) - C(1 - Hc - Hfx)

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn an overflow into #NUM!
        On Error GoTo Overflowed

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Return #NUM! for an input outside its domain
        If E < 0# Or He < 0# Or C < 0# Or Hc < 0# Or Hfx < 0# Then
            SFT_ExposureSingle = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Apply the haircuts to the exposure and the collateral
        Adjusted = E * (1# + He) - C * (1# - Hc - Hfx)
    'Floor the exposure value at 0
        If Adjusted < 0# Then
            Adjusted = 0#
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the fully adjusted exposure value
        SFT_ExposureSingle = Adjusted

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Overflowed:
    'Return #NUM! for an overflow
        SFT_ExposureSingle = CVErr(xlErrNum)

End Function

Public Function SFT_NetExposure( _
    ByVal Es As Variant, _
    ByVal Hs As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                    COMPUTE NET EXPOSURE OF A NETTING SET
'------------------------------------------------------------------------------
' PURPOSE
'   Net exposure of a netting set, the systematic component:
'     net = | sum over issues s of E_s * H_s |
'
' WHY THIS EXISTS
'   The net exposure is one of the terms of SFT_ExposureNetted
'
' INPUTS
'   Es
'     Range or array of the net position in each security issue, lent or
'     posted less borrowed or received, any sign.
'
'   Hs
'     Range or array of the scaled haircut of each issue, in the same order
'     and of the same size, 0 or more. Blank cells count as 0.
'
' RETURNS
'   The net exposure; #REF! for ranges of different sizes, #VALUE! for
'   text or a Boolean, #NUM! for a negative haircut, an overflow or an
'   empty array, or the first error value found.
'
' BEHAVIOR
'   Reads both ranges with ReadPairs, sums E_s * H_s over the issues and
'   returns the absolute value of the sum
'
' ERROR POLICY
'   Never raises: returns the worksheet error found by ReadPairs, and #NUM!
'   for an overflow or an empty array
'
' DEPENDENCIES
'   ReadPairs; VBA runtime (Abs, CVErr, IsEmpty); Excel error constant
'   xlErrNum
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.64-22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "SFT_Formulas.SFT_NetExposure"

    Dim Positions()             As Double       'E_s, flattened
    Dim Haircuts()              As Double       'H_s, flattened
    Dim Problem                 As Variant      'Error value found while reading; Empty for none
    Dim Total                   As Double       'Running sum of E_s * H_s
    Dim IssueIndex              As Long         'Issue

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn an overflow or an empty array into #NUM!
        On Error GoTo Overflowed

'------------------------------------------------------------------------------
' READ INPUTS
'------------------------------------------------------------------------------
    'Read both ranges and return the first problem found
        Problem = ReadPairs(Es, Hs, Positions, Haircuts)
        If Not IsEmpty(Problem) Then
            SFT_NetExposure = Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Sum the haircut positions with their signs
        For IssueIndex = 1 To UBound(Positions)
            Total = Total + Positions(IssueIndex) * Haircuts(IssueIndex)
        Next IssueIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the absolute value of the sum
        SFT_NetExposure = Abs(Total)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Overflowed:
    'Return #NUM! for an overflow or an empty array
        SFT_NetExposure = CVErr(xlErrNum)

End Function

Public Function SFT_GrossExposure( _
    ByVal Es As Variant, _
    ByVal Hs As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                   COMPUTE GROSS EXPOSURE OF A NETTING SET
'------------------------------------------------------------------------------
' PURPOSE
'   Gross exposure of a netting set, the idiosyncratic component:
'     gross = sum over issues s of |E_s| * H_s
'
' WHY THIS EXISTS
'   The gross exposure is one of the terms of SFT_ExposureNetted
'
' INPUTS
'   Es
'     As SFT_NetExposure.
'
'   Hs
'     As SFT_NetExposure.
'
' RETURNS
'   The gross exposure; the errors of SFT_NetExposure.
'
' BEHAVIOR
'   Reads both ranges with ReadPairs and sums |E_s| * H_s over the issues
'
' ERROR POLICY
'   Never raises: returns the worksheet error found by ReadPairs, and #NUM!
'   for an overflow or an empty array
'
' DEPENDENCIES
'   ReadPairs; VBA runtime (Abs, CVErr, IsEmpty); Excel error constant
'   xlErrNum
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.64-22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "SFT_Formulas.SFT_GrossExposure"

    Dim Positions()             As Double       'E_s, flattened
    Dim Haircuts()              As Double       'H_s, flattened
    Dim Problem                 As Variant      'Error value found while reading; Empty for none
    Dim Total                   As Double       'Running sum of |E_s| * H_s
    Dim IssueIndex              As Long         'Issue

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn an overflow or an empty array into #NUM!
        On Error GoTo Overflowed

'------------------------------------------------------------------------------
' READ INPUTS
'------------------------------------------------------------------------------
    'Read both ranges and return the first problem found
        Problem = ReadPairs(Es, Hs, Positions, Haircuts)
        If Not IsEmpty(Problem) Then
            SFT_GrossExposure = Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Sum the haircut absolute positions
        For IssueIndex = 1 To UBound(Positions)
            Total = Total + Abs(Positions(IssueIndex)) * Haircuts(IssueIndex)
        Next IssueIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the sum
        SFT_GrossExposure = Total

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Overflowed:
    'Return #NUM! for an overflow or an empty array
        SFT_GrossExposure = CVErr(xlErrNum)

End Function

Public Function SFT_IssueCount( _
    ByVal Es As Variant, _
    Optional ByVal Threshold As Double = 0.1) _
    As Variant
'
'------------------------------------------------------------------------------
'                    COUNT SECURITY ISSUES OF A NETTING SET
'------------------------------------------------------------------------------
' PURPOSE
'   Number N of security issues in a netting set that count in the gross
'   term: an issue counts when its net position is not 0 and at least
'   Threshold times the largest absolute net position of the netting set.
'
' WHY THIS EXISTS
'   SFT_ExposureNetted divides the gross term by sqrt(N)
'
' INPUTS
'   Es
'     Range or array of the net position in each security issue, any sign;
'     cash positions are not passed. Blank cells count as 0.
'
'   Threshold
'     Share of the largest position, 0 to 1; 0.1 (one tenth) by default.
'
' RETURNS
'   N, 0 when every position is 0; #NUM! for a threshold outside its
'   domain or an empty array, #VALUE! for text or a Boolean, or the first
'   error value found.
'
' BEHAVIOR
'   Validates Threshold, reads Es with ReadPairs (Es passed twice, any
'   sign), finds the largest absolute position, then counts the non-zero
'   positions of at least Threshold times it
'
' ERROR POLICY
'   Never raises: returns #NUM! for a threshold outside its domain (checked
'   before the overflow handler is enabled), the worksheet error found by
'   ReadPairs, and #NUM! for an overflow or an empty array
'
' DEPENDENCIES
'   ReadPairs; VBA runtime (Abs, CVErr, IsEmpty); Excel error constant
'   xlErrNum
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "SFT_Formulas.SFT_IssueCount"

    Dim Positions()             As Double       'E_s, flattened
    Dim Haircuts()              As Double       'Unused: Es is passed twice to ReadPairs
    Dim Problem                 As Variant      'Error value found while reading; Empty for none
    Dim Largest                 As Double       'Largest |E_s|
    Dim Count                   As Long         'Issues counted
    Dim IssueIndex              As Long         'Issue

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Return #NUM! for a threshold outside its domain
        If Threshold < 0# Or Threshold > 1# Then
            SFT_IssueCount = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn an overflow or an empty array into #NUM!
        On Error GoTo Overflowed

'------------------------------------------------------------------------------
' READ INPUTS
'------------------------------------------------------------------------------
    'Read the positions, of any sign, and return the first problem found
        Problem = ReadPairs(Es, Es, Positions, Haircuts, True)
        If Not IsEmpty(Problem) Then
            SFT_IssueCount = Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
    'Find the largest absolute position
        For IssueIndex = 1 To UBound(Positions)
            If Abs(Positions(IssueIndex)) > Largest Then
                Largest = Abs(Positions(IssueIndex))
            End If
        Next IssueIndex
    'Count the non-zero positions of at least Threshold times the largest
        For IssueIndex = 1 To UBound(Positions)
            If Positions(IssueIndex) <> 0# And _
                Abs(Positions(IssueIndex)) >= Threshold * Largest Then
                Count = Count + 1
            End If
        Next IssueIndex

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the number of issues counted
        SFT_IssueCount = Count

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Also reached for an empty array, which Flatten cannot size.
Overflowed:
    'Return #NUM! for an overflow or an empty array
        SFT_IssueCount = CVErr(xlErrNum)

End Function

Public Function SFT_ExposureNetted( _
    ByVal SumE As Double, _
    ByVal SumC As Double, _
    ByVal NetExposure As Double, _
    ByVal GrossExposure As Double, _
    ByVal N As Double, _
    Optional ByVal FXTerm As Double = 0#, _
    Optional ByVal NetWeight As Double = 0.4, _
    Optional ByVal GrossWeight As Double = 0.6) _
    As Variant
'
'------------------------------------------------------------------------------
'                      COMPUTE EXPOSURE OF A NETTING SET
'------------------------------------------------------------------------------
' PURPOSE
'   Fully adjusted exposure value of a netting set of SFTs covered by an
'   eligible master netting agreement:
'     E* = max(0, sum E - sum C + w_net * net + w_gross * gross / sqrt(N)
'                 + sum over currencies k of |E_k^fx| * Hfx)
'
' WHY THIS EXISTS
'   A netting set covered by an eligible master netting agreement takes the
'   netted formula instead of the single-transaction one
'
' INPUTS
'   SumE
'     Current value of all cash and securities lent, sold or posted, 0 or
'     more.
'
'   SumC
'     Current value of all cash and securities borrowed, bought or
'     received, 0 or more.
'
'   NetExposure
'     SFT_NetExposure, 0 or more.
'
'   GrossExposure
'     SFT_GrossExposure, 0 or more.
'
'   N
'     SFT_IssueCount, 0 or more and a whole number; with N = 0 the gross
'     term is 0, and GrossExposure must then be 0.
'
'   FXTerm
'     Sum over the currencies other than the settlement currency of the
'     absolute net position times the scaled currency haircut, 0 or more.
'
'   NetWeight
'     0.4 by default, 0 or more.
'
'   GrossWeight
'     0.6 by default, 0 or more.
'
' RETURNS
'   E*; #NUM! for an input outside its domain or an overflow.
'
' BEHAVIOR
'   Validates the inputs, computes the gross term when N is above 0, adds
'   the terms and floors the result at 0
'
' ERROR POLICY
'   Never raises: returns #NUM! for an input outside its domain and for an
'   overflow
'
' DEPENDENCIES
'   VBA runtime (CVErr, Int, Sqr); Excel error constant xlErrNum
'
' NOTES
'   None
'
' REFERENCE
'   CRE22.64-22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "SFT_Formulas.SFT_ExposureNetted"

    Dim GrossTerm               As Double       'GrossWeight * gross / sqrt(N)
    Dim Adjusted                As Double       'The value before the floor at 0

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn an overflow into #NUM!
        On Error GoTo Overflowed

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Return #NUM! for an input outside its domain
        If SumE < 0# Or SumC < 0# Or NetExposure < 0# Or GrossExposure < 0# Or N < 0# Or _
            N <> Int(N) Or FXTerm < 0# Or NetWeight < 0# Or GrossWeight < 0# Then
            SFT_ExposureNetted = CVErr(xlErrNum)
            Exit Function
        End If
    'Return #NUM! for a gross exposure without any issue to divide it by
        If N = 0# And GrossExposure > 0# Then
            SFT_ExposureNetted = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'Compute the gross term; it stays 0 when no issue counts
        If N > 0# Then
            GrossTerm = GrossWeight * GrossExposure / Sqr(N)
        End If
    'Add the terms
        Adjusted = SumE - SumC + NetWeight * NetExposure + GrossTerm + FXTerm
    'Floor the exposure value at 0
        If Adjusted < 0# Then
            Adjusted = 0#
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the fully adjusted exposure value
        SFT_ExposureNetted = Adjusted

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Overflowed:
    'Return #NUM! for an overflow
        SFT_ExposureNetted = CVErr(xlErrNum)

End Function


'
'------------------------------------------------------------------------------
'
'                                   CAPITAL
'
'------------------------------------------------------------------------------
'

Public Function SFT_CapitalRequirement( _
    ByVal EAD As Double, _
    ByVal RiskWeight As Double, _
    Optional ByVal CapitalRatio As Double = 0.08) _
    As Variant
'
'------------------------------------------------------------------------------
'                        COMPUTE OWN FUNDS REQUIREMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Own funds requirement of an exposure under the standardised approach:
'     K = EAD * RW * CapitalRatio
'
' WHY THIS EXISTS
'   The exposure value becomes a capital figure through the risk weight of
'   the counterparty and the own funds ratio
'
' INPUTS
'   EAD
'     Exposure value E*, 0 or more.
'
'   RiskWeight
'     Risk weight of the counterparty, decimal, 0 to 12.5.
'
'   CapitalRatio
'     Own funds requirement per unit of risk-weighted exposure, 0 to 1;
'     0.08 (8%) by default.
'
' RETURNS
'   The own funds requirement; #NUM! for an input outside its domain or an
'   overflow. The risk-weighted exposure amount is
'   SFT_CapitalRequirement(EAD, RW, 1).
'
' BEHAVIOR
'   Validates the inputs and returns their product
'
' ERROR POLICY
'   Never raises: returns #NUM! for an input outside its domain and for an
'   overflow
'
' DEPENDENCIES
'   VBA runtime (CVErr); Excel error constant xlErrNum
'
' NOTES
'   CORE_Engine calls it with CapitalRatio 1 for the risk-weighted exposure
'   amount and with the configured ratio for the own funds requirement
'
' REFERENCE
'   CRE22.48, RBC20.1; CRR Art. 92(1)(c), (4).
'
' UPDATED
'   2026-10-24
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "SFT_Formulas.SFT_CapitalRequirement"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Turn an overflow into #NUM!
        On Error GoTo Overflowed

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Return #NUM! for an input outside its domain
        If EAD < 0# Or RiskWeight < 0# Or RiskWeight > 12.5 Or _
            CapitalRatio < 0# Or CapitalRatio > 1# Then
            SFT_CapitalRequirement = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the own funds requirement
        SFT_CapitalRequirement = EAD * RiskWeight * CapitalRatio

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Overflowed:
    'Return #NUM! for an overflow
        SFT_CapitalRequirement = CVErr(xlErrNum)

End Function


'
'------------------------------------------------------------------------------
'
'                                   HELPERS
'
'------------------------------------------------------------------------------
'

Private Function ReadPairs( _
    ByVal Es As Variant, _
    ByVal Hs As Variant, _
    ByRef Positions() As Double, _
    ByRef Haircuts() As Double, _
    Optional ByVal AnySign As Boolean = False) _
    As Variant
'
'------------------------------------------------------------------------------
'                       READ TWO RANGES OF THE SAME SIZE
'------------------------------------------------------------------------------
' PURPOSE
'   Flatten two ranges or arrays of the same size into two 1-based Double
'   arrays, checking every element.
'
' WHY THIS EXISTS
'   SFT_NetExposure, SFT_GrossExposure and SFT_IssueCount read their ranges
'   the same way and return the same worksheet errors
'
' INPUTS
'   Es
'     Positions, any sign.
'
'   Hs
'     Haircuts, 0 or more unless AnySign.
'
'   Positions
'     Receives the flattened values of Es.
'
'   Haircuts
'     Receives the flattened values of Hs.
'
'   AnySign
'     True accepts a negative value in Hs (when Es is passed twice).
'
' RETURNS
'   Empty when both were read; otherwise the worksheet error the caller
'   returns: #REF! for different sizes, #VALUE! for text or a Boolean,
'   #NUM! for a negative haircut, or the error value found.
'
' BEHAVIOR
'   Flattens both inputs, compares their sizes, then checks and converts
'   each pair of elements in order, stopping at the first problem. Blank
'   cells and empty strings are left at 0.
'
' ERROR POLICY
'   Returns the problem as a worksheet error value; an error raised by
'   Flatten (an empty array) reaches the caller, whose handler returns
'   #NUM!
'
' DEPENDENCIES
'   Flatten, ElementIsNumber; VBA runtime; Excel error constants xlErrNum,
'   xlErrValue, xlErrRef
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
    Const PROC_NAME             As String = "SFT_Formulas.ReadPairs"

    Dim PositionValues          As Variant      'Es as a flat 1-based Variant array
    Dim HaircutValues           As Variant      'Hs as a flat 1-based Variant array
    Dim Count                   As Long         'Number of elements
    Dim ElementIndex            As Long         'Element

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set explicitly: LibreOffice, which runs tools/lo_regression.py, keeps a
    'Variant result from the previous call instead of starting with Empty.
        ReadPairs = Empty

'------------------------------------------------------------------------------
' FLATTEN
'------------------------------------------------------------------------------
    'Flatten both inputs
        PositionValues = Flatten(Es)
        HaircutValues = Flatten(Hs)
    'Return #REF! for inputs of different sizes
        Count = UBound(PositionValues)
        If UBound(HaircutValues) <> Count Then
            ReadPairs = CVErr(xlErrRef)
            Exit Function
        End If
    'Size the output arrays
        ReDim Positions(1 To Count)
        ReDim Haircuts(1 To Count)

'------------------------------------------------------------------------------
' CHECK EACH ELEMENT
'------------------------------------------------------------------------------
    'Check and convert each pair, stopping at the first problem
        For ElementIndex = 1 To Count
            'Return an error value found in the positions
            If IsError(PositionValues(ElementIndex)) Then
                ReadPairs = PositionValues(ElementIndex)
                Exit Function
            End If
            'Return an error value found in the haircuts
            If IsError(HaircutValues(ElementIndex)) Then
                ReadPairs = HaircutValues(ElementIndex)
                Exit Function
            End If
            'Return #VALUE! for an element that is not a number
            If Not ElementIsNumber(PositionValues(ElementIndex)) Or _
                Not ElementIsNumber(HaircutValues(ElementIndex)) Then
                ReadPairs = CVErr(xlErrValue)
                Exit Function
            End If
            'Convert the elements; blank cells and empty strings stay 0
            'CDbl stays unqualified: tools/lo_regression.py fails with VBA.CDbl
            If Not IsEmpty(PositionValues(ElementIndex)) And _
                VarType(PositionValues(ElementIndex)) <> vbString Then
                Positions(ElementIndex) = CDbl(PositionValues(ElementIndex))
            End If
            If Not IsEmpty(HaircutValues(ElementIndex)) And _
                VarType(HaircutValues(ElementIndex)) <> vbString Then
                Haircuts(ElementIndex) = CDbl(HaircutValues(ElementIndex))
            End If
            'Return #NUM! for a negative haircut unless any sign is accepted
            If Haircuts(ElementIndex) < 0# And Not AnySign Then
                ReadPairs = CVErr(xlErrNum)
                Exit Function
            End If
        Next ElementIndex

End Function

Private Function Flatten( _
    ByVal Source As Variant) _
    As Variant
'
'------------------------------------------------------------------------------
'                         FLATTEN A RANGE OR AN ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Turn a range, a 1-D or 2-D array or a single value into a 1-based 1-D
'   Variant array, row by row.
'
' WHY THIS EXISTS
'   A worksheet argument can be a Range, an array or a single value;
'   ReadPairs reads all of them the same way
'
' INPUTS
'   Source
'     A Range, an array or a scalar.
'
' RETURNS
'   The array; one element for a scalar.
'
' BEHAVIOR
'   Reads the value of a Range, wraps a scalar in a one-element array,
'   tells a 1-D array from a 2-D one, and copies the elements row by row
'
' ERROR POLICY
'   The probe of the second dimension suppresses its one expected error;
'   any other error (an empty array, which cannot be sized) is raised to
'   the caller
'
' DEPENDENCIES
'   VBA runtime; Range.Value for a Range argument
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
    Const PROC_NAME             As String = "SFT_Formulas.Flatten"

    Dim Values                  As Variant      'Source values
    Dim Flat()                  As Variant      'Result
    Dim Dimensions              As Long         'Number of array dimensions: 1 or 2
    Dim Probe                   As Long         'Upper bound of the second dimension, if any
    Dim Count                   As Long         'Elements written
    Dim SourceRow               As Long         'Row
    Dim Col                     As Long         'Column

'------------------------------------------------------------------------------
' READ THE SOURCE
'------------------------------------------------------------------------------
    'Read the values of a Range, or take the argument as it is
        If IsObject(Source) Then
            Values = Source.Value
        Else
            Values = Source
        End If
    'Wrap a single value in a one-element array
        If Not IsArray(Values) Then
            ReDim Flat(1 To 1)
            Flat(1) = Values
            Flatten = Flat
            Exit Function
        End If

'------------------------------------------------------------------------------
' COUNT THE DIMENSIONS
'------------------------------------------------------------------------------
    'UBound on a missing dimension raises; that single expected error tells
    'a 1-D array from a 2-D one.
        Dimensions = 2
        On Error Resume Next
        Probe = UBound(Values, 2)
        If Err.Number <> 0 Then
            Dimensions = 1
        End If
        Err.Clear
        On Error GoTo 0

'------------------------------------------------------------------------------
' FLATTEN
'------------------------------------------------------------------------------
    'Copy the elements of a 1-D array, or of a 2-D array row by row
        If Dimensions = 1 Then
            ReDim Flat(1 To UBound(Values) - LBound(Values) + 1)
            For SourceRow = LBound(Values) To UBound(Values)
                Count = Count + 1
                Flat(Count) = Values(SourceRow)
            Next SourceRow
        Else
            ReDim Flat(1 To (UBound(Values, 1) - LBound(Values, 1) + 1) * _
                (Probe - LBound(Values, 2) + 1))
            For SourceRow = LBound(Values, 1) To UBound(Values, 1)
                For Col = LBound(Values, 2) To Probe
                    Count = Count + 1
                    Flat(Count) = Values(SourceRow, Col)
                Next Col
            Next SourceRow
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the flat array
        Flatten = Flat

End Function

Private Function ElementIsNumber( _
    ByVal Content As Variant) _
    As Boolean
'
'------------------------------------------------------------------------------
'                     TELL WHETHER AN ELEMENT IS A NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether an element of a range is usable as a number: a number, a
'   blank cell, or an empty string, which both count as 0. A Boolean is
'   not, although IsNumeric accepts it.
'
' INPUTS
'   Content
'     One element of a flattened range or array.
'
' RETURNS
'   True for a number, a blank cell or an empty string; False otherwise
'
' UPDATED
'   2026-10-29
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Accept a blank cell, reject a Boolean, accept only an empty string,
    'and test any other value with IsNumeric
        If IsEmpty(Content) Then
            ElementIsNumber = True
        ElseIf VarType(Content) = vbBoolean Then
            ElementIsNumber = False
        ElseIf VarType(Content) = vbString Then
            ElementIsNumber = (VBA.Len(Content) = 0)
        Else
            ElementIsNumber = IsNumeric(Content)
        End If

End Function
