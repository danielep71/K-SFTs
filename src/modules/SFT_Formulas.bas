Attribute VB_Name = "SFT_Formulas"
'==============================================================================
' MODULE: SFT_Formulas
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the building blocks of the exposure value of securities
'   financing transactions (SFTs) under the financial collateral
'   comprehensive method, with supervisory haircuts, as public functions.
'   The engine calls them, and they can also be entered in worksheet cells
'   as user-defined functions, so every intermediate figure can be checked
'   independently.
'
' PUBLIC SURFACE
'   The eight SFT_* functions, listed in docs/PUBLIC_API.txt. Their names,
'   arguments, defaults and results are a contract: changing one is a public
'   API change.
'
' DEPENDENCIES
'   VBA runtime only; no WorksheetFunction, so results do not depend on the
'   Excel version or the locale.
'
' STATE OWNERSHIP
'   Stateless. Results depend only on the arguments.
'
' ERROR POLICY
'   Each function checks the domain of its inputs, stated in its INPUTS
'   section, and returns a worksheet error
'   value instead of raising or returning a misleading number: #NUM! for a
'   number outside its domain, #VALUE! for a value that is not a number
'   where one is needed, #REF! for ranges of different sizes, and the error
'   itself for an error value in a range. A VBA caller receives the same
'   error value in the Variant result and tests it with IsError;
'   CORE_Engine stops the run when it receives one. An overflow beyond the
'   range of a Double also gives #NUM!.
'
' REFERENCE
'   BCBS CRE22 (paragraph numbers in brackets); EU CRR Art. 220, 223, 224,
'   226; risk weights per CRE20 and CRR Part Three Title II Chapter 2.
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
    'Require explicit declarations. This module is public: its functions are
    'available to worksheet cells.
    Option Explicit


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
'==============================================================================
'                              SFT_HoldingPeriod
'------------------------------------------------------------------------------
' PURPOSE
'   Holding period T_M of a transaction or netting set, in business days:
'     T = MinimumBD, at least LargeBD when LargeOrIlliquid,
'     at least OverrideBD when one is given, then doubled for Disputes.
'
' INPUTS
'   MinimumBD: minimum holding period of the transaction type, business
'      days, 1 or more: 5 for repo-style transactions, 10 for other
'      capital-market-driven transactions; the longest when a netting set
'      holds both.
'   LargeOrIlliquid: True for a netting set with more than 5,000
'      transactions in the previous quarter or with illiquid collateral.
'   Disputes: True after more than two margin-call disputes lasting longer
'      than the holding period in the previous two quarters.
'   OverrideBD: a longer holding period required by the supervisor or the
'      bank, business days, 0 or more; 0 for none. It never shortens T.
'   LargeBD: the holding period for LargeOrIlliquid, 1 or more; 20 by
'      default.
'
' RETURNS
'   T_M in business days; #NUM! for an input outside its domain.
'
' REFERENCE
'   CRE22.57-22.58; CRR Art. 224(2), Art. 285(3)-(4).
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Holding   As Double    'Holding period being built, business days

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        If MinimumBD < 1# Or OverrideBD < 0# Or LargeBD < 1# Then
            SFT_HoldingPeriod = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
    'The floors are applied before the doubling, which applies to the
    'holding period that would otherwise be used.
        Holding = MinimumBD
        If LargeOrIlliquid And LargeBD > Holding Then
            Holding = LargeBD
        End If
        If OverrideBD > Holding Then
            Holding = OverrideBD
        End If
        If Disputes Then
            Holding = 2# * Holding
        End If
        SFT_HoldingPeriod = Holding

End Function


Public Function SFT_ScaledHaircut( _
    ByVal H10 As Double, _
    ByVal HoldingBD As Double, _
    Optional ByVal RemarginBD As Double = 1#, _
    Optional ByVal BaseBD As Double = 10#) _
    As Variant
'
'==============================================================================
'                              SFT_ScaledHaircut
'------------------------------------------------------------------------------
' PURPOSE
'   Haircut for the holding period and remargining frequency of the
'   transaction, from the supervisory haircut of the table:
'     H = H10 * sqrt((N_R + T_M - 1) / BaseBD)
'
' INPUTS
'   H10: supervisory haircut for BaseBD business days with daily
'      remargining, decimal, 0 to 1.
'   HoldingBD: holding period T_M, business days, 1 or more.
'   RemarginBD: business days N_R between remargining (or revaluation), 1
'      or more; 1 (daily) by default.
'   BaseBD: holding period of the table, business days, above 0; 10 by
'      default.
'
' RETURNS
'   The scaled haircut, decimal; #NUM! for an input outside its domain or an
'   overflow.
'
' REFERENCE
'   CRE22.59; CRR Art. 226, which gives the same figure from the column of
'   Art. 224 Table 1 for the liquidation period, since that column is H10
'   scaled by sqrt(T_M / 10).
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        If H10 < 0# Or H10 > 1# Or HoldingBD < 1# Or RemarginBD < 1# Or BaseBD <= 0# Then
            SFT_ScaledHaircut = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        SFT_ScaledHaircut = H10 * Sqr((RemarginBD + HoldingBD - 1#) / BaseBD)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
Overflowed:
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
'==============================================================================
'                              SFT_ExposureSingle
'------------------------------------------------------------------------------
' PURPOSE
'   Fully adjusted exposure value of one collateralised transaction:
'     E* = max(0, E * (1 + He) - C * (1 - Hc - Hfx))
'
' INPUTS
'   E: current value of what the bank lent, sold or posted, 0 or more.
'   He: haircut of the exposure, scaled, 0 or more; for several items, the
'      value-weighted average of their haircuts [CRE22.55].
'   C: current value of the collateral received, 0 or more.
'   Hc: haircut of the collateral, scaled, 0 or more.
'   Hfx: currency-mismatch haircut, scaled, 0 or more; 0 when the
'      collateral is in the currency of the exposure.
'
' RETURNS
'   E*, in the unit of E and C; #NUM! for an input outside its domain or
'   an overflow.
'
' REFERENCE
'   CRE22.46; CRR Art. 223(2), (5).
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Adjusted   As Double    'E(1 + He) - C(1 - Hc - Hfx)

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        If E < 0# Or He < 0# Or C < 0# Or Hc < 0# Or Hfx < 0# Then
            SFT_ExposureSingle = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        Adjusted = E * (1# + He) - C * (1# - Hc - Hfx)
        If Adjusted < 0# Then
            Adjusted = 0#
        End If
        SFT_ExposureSingle = Adjusted
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
Overflowed:
        SFT_ExposureSingle = CVErr(xlErrNum)

End Function


Public Function SFT_NetExposure( _
    ByVal Es As Variant, _
    ByVal Hs As Variant) _
    As Variant
'
'==============================================================================
'                               SFT_NetExposure
'------------------------------------------------------------------------------
' PURPOSE
'   Net exposure of a netting set, the systematic component:
'     net = | sum over issues s of E_s * H_s |
'
' INPUTS
'   Es: range or array of the net position in each security issue, lent
'      or posted less borrowed or received, any sign.
'   Hs: range or array of the scaled haircut of each issue, in the same
'      order and of the same size, 0 or more. Blank cells count as 0.
'
' RETURNS
'   The net exposure; #REF! for ranges of different sizes, #VALUE! for
'   text, #NUM! for a negative haircut, or the first error value found.
'
' REFERENCE
'   CRE22.64-22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Positions()   As Double     'E_s, flattened
    Dim Haircuts()    As Double     'H_s, flattened
    Dim Problem       As Variant    'Error value found while reading; Empty for none
    Dim Total         As Double     'Running sum of E_s * H_s
    Dim i             As Long       'Issue

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        Problem = ReadPairs(Es, Hs, Positions, Haircuts)
        If Not IsEmpty(Problem) Then
            SFT_NetExposure = Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        For i = 1 To UBound(Positions)
            Total = Total + Positions(i) * Haircuts(i)
        Next i
        SFT_NetExposure = Abs(Total)
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
Overflowed:
        SFT_NetExposure = CVErr(xlErrNum)

End Function


Public Function SFT_GrossExposure( _
    ByVal Es As Variant, _
    ByVal Hs As Variant) _
    As Variant
'
'==============================================================================
'                              SFT_GrossExposure
'------------------------------------------------------------------------------
' PURPOSE
'   Gross exposure of a netting set, the idiosyncratic component:
'     gross = sum over issues s of |E_s| * H_s
'
' INPUTS
'   As SFT_NetExposure.
'
' RETURNS
'   The gross exposure; the errors of SFT_NetExposure.
'
' REFERENCE
'   CRE22.64-22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Positions()   As Double     'E_s, flattened
    Dim Haircuts()    As Double     'H_s, flattened
    Dim Problem       As Variant    'Error value found while reading; Empty for none
    Dim Total         As Double     'Running sum of |E_s| * H_s
    Dim i             As Long       'Issue

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        Problem = ReadPairs(Es, Hs, Positions, Haircuts)
        If Not IsEmpty(Problem) Then
            SFT_GrossExposure = Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        For i = 1 To UBound(Positions)
            Total = Total + Abs(Positions(i)) * Haircuts(i)
        Next i
        SFT_GrossExposure = Total
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
Overflowed:
        SFT_GrossExposure = CVErr(xlErrNum)

End Function


Public Function SFT_IssueCount( _
    ByVal Es As Variant, _
    Optional ByVal Threshold As Double = 0.1) _
    As Variant
'
'==============================================================================
'                                SFT_IssueCount
'------------------------------------------------------------------------------
' PURPOSE
'   Number N of security issues in a netting set that count in the gross
'   term: an issue counts when its net position is not 0 and at least
'   Threshold times the largest absolute net position of the netting set.
'
' INPUTS
'   Es: range or array of the net position in each security issue, any
'      sign; cash positions are not passed. Blank cells count as 0.
'   Threshold: share of the largest position, 0 to 1; 0.1 (one tenth) by
'      default.
'
' RETURNS
'   N, 0 when every position is 0; #NUM! for a threshold outside its
'   domain, #VALUE! for text, or the first error value found.
'
' REFERENCE
'   CRE22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Positions()   As Double     'E_s, flattened
    Dim Haircuts()    As Double     'Unused: Es is passed twice to ReadPairs
    Dim Problem       As Variant    'Error value found while reading; Empty for none
    Dim Largest       As Double     'Largest |E_s|
    Dim Count         As Long       'Issues counted
    Dim i             As Long       'Issue

'------------------------------------------------------------------------------
' VALIDATE AND READ
'------------------------------------------------------------------------------
        If Threshold < 0# Or Threshold > 1# Then
            SFT_IssueCount = CVErr(xlErrNum)
            Exit Function
        End If
        Problem = ReadPairs(Es, Es, Positions, Haircuts, True)
        If Not IsEmpty(Problem) Then
            SFT_IssueCount = Problem
            Exit Function
        End If

'------------------------------------------------------------------------------
' COUNT
'------------------------------------------------------------------------------
        For i = 1 To UBound(Positions)
            If Abs(Positions(i)) > Largest Then
                Largest = Abs(Positions(i))
            End If
        Next i
        For i = 1 To UBound(Positions)
            If Positions(i) <> 0# And Abs(Positions(i)) >= Threshold * Largest Then
                Count = Count + 1
            End If
        Next i
        SFT_IssueCount = Count

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
'==============================================================================
'                              SFT_ExposureNetted
'------------------------------------------------------------------------------
' PURPOSE
'   Fully adjusted exposure value of a netting set of SFTs covered by an
'   eligible master netting agreement:
'     E* = max(0, sum E - sum C + w_net * net + w_gross * gross / sqrt(N)
'                 + sum over currencies k of |E_k^fx| * Hfx)
'
' INPUTS
'   SumE: current value of all cash and securities lent, sold or posted,
'      0 or more.
'   SumC: current value of all cash and securities borrowed, bought or
'      received, 0 or more.
'   NetExposure, GrossExposure: SFT_NetExposure and SFT_GrossExposure, 0 or
'      more.
'   N: SFT_IssueCount, 0 or more and a whole number; with N = 0 the gross
'      term is 0, and GrossExposure must then be 0.
'   FXTerm: sum over the currencies other than the settlement currency of
'      the absolute net position times the scaled currency haircut, 0 or
'      more.
'   NetWeight, GrossWeight: 0.4 and 0.6 by default, 0 or more.
'
' RETURNS
'   E*; #NUM! for an input outside its domain or an overflow.
'
' REFERENCE
'   CRE22.64-22.65; CRR Art. 220(3).
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GrossTerm   As Double    'GrossWeight * gross / sqrt(N)
    Dim Adjusted    As Double    'The value before the floor at 0

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        If SumE < 0# Or SumC < 0# Or NetExposure < 0# Or GrossExposure < 0# Or N < 0# Or _
           N <> Int(N) Or FXTerm < 0# Or NetWeight < 0# Or GrossWeight < 0# Then
            SFT_ExposureNetted = CVErr(xlErrNum)
            Exit Function
        End If
        If N = 0# And GrossExposure > 0# Then
            SFT_ExposureNetted = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        If N > 0# Then
            GrossTerm = GrossWeight * GrossExposure / Sqr(N)
        End If
        Adjusted = SumE - SumC + NetWeight * NetExposure + GrossTerm + FXTerm
        If Adjusted < 0# Then
            Adjusted = 0#
        End If
        SFT_ExposureNetted = Adjusted
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
Overflowed:
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
'==============================================================================
'                            SFT_CapitalRequirement
'------------------------------------------------------------------------------
' PURPOSE
'   Own funds requirement of an exposure under the standardised approach:
'     K = EAD * RW * CapitalRatio
'
' INPUTS
'   EAD: exposure value E*, 0 or more.
'   RiskWeight: risk weight of the counterparty, decimal, 0 to 12.5.
'   CapitalRatio: own funds requirement per unit of risk-weighted
'      exposure, 0 to 1; 0.08 (8%) by default.
'
' RETURNS
'   The own funds requirement; #NUM! for an input outside its domain or an
'   overflow. The risk-weighted exposure amount is
'   SFT_CapitalRequirement(EAD, RW, 1).
'
' REFERENCE
'   CRE22.48, RBC20.1; CRR Art. 92(1)(c), (4).
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
        On Error GoTo Overflowed
        If EAD < 0# Or RiskWeight < 0# Or RiskWeight > 12.5 Or CapitalRatio < 0# Or CapitalRatio > 1# Then
            SFT_CapitalRequirement = CVErr(xlErrNum)
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALCULATE
'------------------------------------------------------------------------------
        SFT_CapitalRequirement = EAD * RiskWeight * CapitalRatio
        Exit Function

'------------------------------------------------------------------------------
' HANDLE OVERFLOW
'------------------------------------------------------------------------------
Overflowed:
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
'==============================================================================
'                                  ReadPairs
'------------------------------------------------------------------------------
' PURPOSE
'   Flatten two ranges or arrays of the same size into two 1-based Double
'   arrays, checking every element.
'
' INPUTS
'   Es: positions, any sign; Hs: haircuts, 0 or more unless AnySign.
'   Positions, Haircuts: receive the flattened values.
'   AnySign: True accepts a negative value in Hs (when Es is passed twice).
'
' RETURNS
'   Empty when both were read; otherwise the worksheet error the caller
'   returns: #REF! for different sizes, #VALUE! for text, #NUM! for a
'   negative haircut, or the error value found.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PositionValues   As Variant    'Es as a flat 1-based Variant array
    Dim HaircutValues    As Variant    'Hs as a flat 1-based Variant array
    Dim Count            As Long       'Number of elements
    Dim i                As Long       'Element

'------------------------------------------------------------------------------
' FLATTEN
'------------------------------------------------------------------------------
        PositionValues = Flatten(Es)
        HaircutValues = Flatten(Hs)
        Count = UBound(PositionValues)
        If UBound(HaircutValues) <> Count Then
            ReadPairs = CVErr(xlErrRef)
            Exit Function
        End If
        ReDim Positions(1 To Count)
        ReDim Haircuts(1 To Count)

'------------------------------------------------------------------------------
' CHECK EACH ELEMENT
'------------------------------------------------------------------------------
        For i = 1 To Count
            If IsError(PositionValues(i)) Then
                ReadPairs = PositionValues(i)
                Exit Function
            End If
            If IsError(HaircutValues(i)) Then
                ReadPairs = HaircutValues(i)
                Exit Function
            End If
            If Not ElementIsNumber(PositionValues(i)) Or Not ElementIsNumber(HaircutValues(i)) Then
                ReadPairs = CVErr(xlErrValue)
                Exit Function
            End If
            If Not IsEmpty(PositionValues(i)) And VarType(PositionValues(i)) <> vbString Then
                Positions(i) = CDbl(PositionValues(i))
            End If
            If Not IsEmpty(HaircutValues(i)) And VarType(HaircutValues(i)) <> vbString Then
                Haircuts(i) = CDbl(HaircutValues(i))
            End If
            If Haircuts(i) < 0# And Not AnySign Then
                ReadPairs = CVErr(xlErrNum)
                Exit Function
            End If
        Next i

End Function


Private Function Flatten( _
    ByVal Source As Variant) _
    As Variant
'
'==============================================================================
'                                   Flatten
'------------------------------------------------------------------------------
' PURPOSE
'   Turn a range, a 1-D or 2-D array or a single value into a 1-based 1-D
'   Variant array, row by row.
'
' INPUTS
'   Source: a Range, an array or a scalar.
'
' RETURNS
'   The array; one element for a scalar.
'
' UPDATED
'   2026-10-24
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Values       As Variant    'Source values
    Dim Flat()       As Variant    'Result
    Dim Dimensions   As Long       'Number of array dimensions: 1 or 2
    Dim Probe        As Long       'Upper bound of the second dimension, if any
    Dim Count        As Long       'Elements written
    Dim r            As Long       'Row
    Dim Col          As Long       'Column

'------------------------------------------------------------------------------
' READ THE SOURCE
'------------------------------------------------------------------------------
        If IsObject(Source) Then
            Values = Source.Value
        Else
            Values = Source
        End If
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
        If Dimensions = 1 Then
            ReDim Flat(1 To UBound(Values) - LBound(Values) + 1)
            For r = LBound(Values) To UBound(Values)
                Count = Count + 1
                Flat(Count) = Values(r)
            Next r
        Else
            ReDim Flat(1 To (UBound(Values, 1) - LBound(Values, 1) + 1) * (Probe - LBound(Values, 2) + 1))
            For r = LBound(Values, 1) To UBound(Values, 1)
                For Col = LBound(Values, 2) To Probe
                    Count = Count + 1
                    Flat(Count) = Values(r, Col)
                Next Col
            Next r
        End If
        Flatten = Flat

End Function


Private Function ElementIsNumber( _
    ByVal Content As Variant) _
    As Boolean
'
'==============================================================================
'                               ElementIsNumber
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether an element of a range is usable as a number: a number, a
'   blank cell, or an empty string, which both count as 0.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        If IsEmpty(Content) Then
            ElementIsNumber = True
        ElseIf VarType(Content) = vbString Then
            ElementIsNumber = (Len(Content) = 0)
        Else
            ElementIsNumber = IsNumeric(Content)
        End If

End Function
