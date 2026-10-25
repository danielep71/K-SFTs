Attribute VB_Name = "CORE_Util"
'==============================================================================
' MODULE: CORE_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the small helpers the engine uses everywhere: tolerant conversion
'   of cell values to text, numbers, dates and Booleans; keyed lookups built
'   on a VBA Collection; and the sheet helpers that read parameters and input
'   columns and write output blocks. Taken from K-SACCR's CORE_Util with the
'   SA-CCR specific helpers left out, so that both workbooks read and write
'   their sheets the same way.
'
' PUBLIC SURFACE
'   None outside this VBA project. Every procedure is Public for in-project use
'   by CORE_Engine and M_Main; Option Private Module keeps them off the
'   supported external surface.
'
' DEPENDENCIES
'   CORE_Config for sheet names, layout constants, workbook names and error
'   numbers. Keyed lookups use a VBA Collection, so no Scripting.Dictionary
'   reference is needed.
'
' STATE OWNERSHIP
'   Stateless. The sheet helpers read and write the worksheets of ThisWorkbook
'   that their caller names; they change no Excel application setting.
'
' ERROR POLICY
'   Conversions never raise: an error value, blank or unusable input returns
'   the documented default. KeyIndex and NamedCell contain the expected
'   lookup errors and report "not found" instead. OutputFirstRow raises
'   ERR_OUTPUT_LAYOUT when an output table cannot be located. Every other
'   error, for example a missing worksheet, propagates to the caller.
'
' KNOWN DEVIATION
'   As in K-SACCR, this module lives in src/core but reads and writes
'   worksheets, which the repository structure reserves for src/workbook.
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
    'Require explicit declarations; keep the helpers inside this project.
    Option Explicit
    Option Private Module


'
'------------------------------------------------------------------------------
'
'                                 CONVERSIONS
'
'------------------------------------------------------------------------------
'

Public Function SafeStr( _
    ByVal Content As Variant) _
    As String
'
'==============================================================================
'                                   SafeStr
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to trimmed text without failing on error values.
'
' INPUTS
'   Content: any cell value, including #N/A, #VALUE! and other error values.
'
' RETURNS
'   The trimmed text of Content; an empty string for an error value, Empty
'   or Null.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'CStr raises on an error value, so error values, Empty and Null are
    'mapped to an empty string before converting.
        If IsError(Content) Then
            SafeStr = ""
        ElseIf IsEmpty(Content) Or IsNull(Content) Then
            SafeStr = ""
        Else
            SafeStr = Trim$(CStr(Content))
        End If

End Function


Public Function UTxt( _
    ByVal Content As Variant) _
    As String
'
'==============================================================================
'                                     UTxt
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to trimmed upper-case text, so that codes such as
'   "long", "Long" and " LONG " compare equal.
'
' INPUTS
'   Content: any cell value.
'
' RETURNS
'   UCase$ of SafeStr(Content).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
        UTxt = UCase$(SafeStr(Content))

End Function


Public Function IsBlankCell( _
    ByVal Content As Variant) _
    As Boolean
'
'==============================================================================
'                                 IsBlankCell
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a cell value is empty or contains only spaces.
'
' INPUTS
'   Content: any cell value. An error value counts as blank.
'
' RETURNS
'   True when SafeStr(Content) is empty.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
        IsBlankCell = (Len(SafeStr(Content)) = 0)

End Function


Public Function IsNum( _
    ByVal Content As Variant) _
    As Boolean
'
'==============================================================================
'                                    IsNum
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a cell value can be used as a number.
'
' INPUTS
'   Content: any cell value.
'
' RETURNS
'   False for an error value, Empty or a blank string; otherwise
'   IsNumeric(Content).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
    'IsNumeric alone treats Empty as numeric, so Empty and blank strings are
    'rejected first.
        If IsError(Content) Or IsEmpty(Content) Then
            IsNum = False
        ElseIf VarType(Content) = vbString Then
            IsNum = (Len(Trim$(Content)) > 0) And IsNumeric(Content)
        Else
            IsNum = IsNumeric(Content)
        End If

End Function


Public Function ToDbl( _
    ByVal Content As Variant, _
    Optional ByVal Fallback As Double = 0#) _
    As Double
'
'==============================================================================
'                                    ToDbl
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to a Double, falling back to a default.
'
' INPUTS
'   Content: any cell value.
'   Fallback: value returned when Content is not a usable number; 0 by
'   default.
'
' RETURNS
'   CDbl(Content) when IsNum(Content); otherwise Fallback.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
        If IsNum(Content) Then
            ToDbl = CDbl(Content)
        Else
            ToDbl = Fallback
        End If

End Function


Public Function ToSerial( _
    ByVal Content As Variant) _
    As Double
'
'==============================================================================
'                                   ToSerial
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to a date serial number.
'
' INPUTS
'   Content: a date, a number already holding a serial, or text in the ISO
'      form yyyy-mm-dd. Other text, such as 31/03/2027, 03/31/2027 or a serial
'      written as text, is not read: its meaning depends on the Windows
'      regional settings (K-SACCR #40).
'
' RETURNS
'   The date serial as a Double; -1 when Content is missing or not a date.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Trimmed        As String    'Trimmed text of Content
    Dim YearPart       As Long      'Year of an ISO text date
    Dim MonthPart      As Long      'Month of an ISO text date
    Dim DayPart        As Long      'Day of an ISO text date
    Dim CalendarDate   As Date      'The ISO text date, built from the three parts

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Start from the "missing" marker and replace it only when Content is
    'usable.
    'Text is tested before numbers: IsNum accepts numeric text such as
    '"47756", which is not a yyyy-mm-dd date.
        ToSerial = -1#
        If IsError(Content) Or IsEmpty(Content) Then
            Exit Function
        End If
        If VarType(Content) = vbDate Then
            ToSerial = CDbl(Content)
        ElseIf VarType(Content) = vbString Then
            Trimmed = Trim$(Content)
            If Not Trimmed Like "####-##-##" Then
                Exit Function
            End If

    'Only a calendar date is read: 2027-02-30 is not rolled into March.
            YearPart = CLng(Left$(Trimmed, 4))
            MonthPart = CLng(Mid$(Trimmed, 6, 2))
            DayPart = CLng(Right$(Trimmed, 2))
            If YearPart < 1900 Or MonthPart < 1 Or MonthPart > 12 Or DayPart < 1 Or DayPart > 31 Then
                Exit Function
            End If
            CalendarDate = DateSerial(YearPart, MonthPart, DayPart)
            If Month(CalendarDate) = MonthPart And Day(CalendarDate) = DayPart Then
                ToSerial = CDbl(CalendarDate)
            End If
        ElseIf IsNum(Content) Then
            ToSerial = CDbl(Content)
        End If

End Function


Public Function TryBool( _
    ByVal Content As Variant, _
    ByVal Fallback As Boolean, _
    ByRef Result As Boolean) _
    As Boolean
'
'==============================================================================
'                                   TryBool
'------------------------------------------------------------------------------
' PURPOSE
'   Read a Y/N style cell value strictly: unlike ToBool, an unrecognised
'   value is reported instead of silently taking the default (K-SACCR #35).
'
' INPUTS
'   Content: a Boolean, blank, or text such as Y, YES, TRUE, 1, N, NO,
'      FALSE, 0 (and the Italian VERO, SI, FALSO).
'   Fallback: value used when Content is blank.
'   Result: receives the value read.
'
' RETURNS
'   True with Result set when Content is blank or recognised; False for an
'   error value or unrecognised text, with Result left as Fallback.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Result = Fallback
        If IsError(Content) Then
            Exit Function
        End If
        If VarType(Content) = vbBoolean Then
            Result = Content
            TryBool = True
            Exit Function
        End If
        Select Case UTxt(Content)
            Case ""
                TryBool = True
            Case "Y", "YES", "TRUE", "1", "VERO", "SI"
                Result = True
                TryBool = True
            Case "N", "NO", "FALSE", "0", "FALSO"
                Result = False
                TryBool = True
        End Select

End Function


Public Function TryDbl( _
    ByVal Content As Variant, _
    ByVal Fallback As Double, _
    ByRef Result As Double) _
    As Boolean
'
'==============================================================================
'                                    TryDbl
'------------------------------------------------------------------------------
' PURPOSE
'   Read a numeric cell value strictly: unlike ToDbl, text that is not a
'   number is reported instead of silently taking the default (K-SACCR #35).
'
' INPUTS
'   Content: a number, blank, or anything else.
'   Fallback: value used when Content is blank.
'   Result: receives the value read.
'
' RETURNS
'   True with Result set when Content is blank or a usable number; False
'   for an error value or non-numeric text, with Result left as Fallback.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Result = Fallback
        If IsError(Content) Then
            Exit Function
        End If
        If IsBlankCell(Content) Then
            TryDbl = True
        ElseIf IsNum(Content) Then
            Result = CDbl(Content)
            TryDbl = True
        End If

End Function


Public Function Max2( _
    ByVal FirstValue As Double, _
    ByVal SecondValue As Double) _
    As Double
'
'==============================================================================
'                                     Max2
'------------------------------------------------------------------------------
' PURPOSE
'   Return the larger of two numbers.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If FirstValue > SecondValue Then
            Max2 = FirstValue
        Else
            Max2 = SecondValue
        End If

End Function


Public Function Min2( _
    ByVal FirstValue As Double, _
    ByVal SecondValue As Double) _
    As Double
'
'==============================================================================
'                                     Min2
'------------------------------------------------------------------------------
' PURPOSE
'   Return the smaller of two numbers.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
        If FirstValue < SecondValue Then
            Min2 = FirstValue
        Else
            Min2 = SecondValue
        End If

End Function


'
'------------------------------------------------------------------------------
'
'                                KEYED LOOKUPS
'
'------------------------------------------------------------------------------
'

Public Function KeyIndex( _
    ByVal Lookup As Collection, _
    ByVal Key As String) _
    As Long
'
'==============================================================================
'                                   KeyIndex
'------------------------------------------------------------------------------
' PURPOSE
'   Look up the array index stored under a key.
'
' INPUTS
'   Lookup: a Collection filled by KeyAdd.
'   Key: lookup key; compared case-insensitively.
'
' RETURNS
'   The stored 1-based index; 0 when the key is not present.
'
' ERROR POLICY
'   Collection.Item raises for a missing key. That single expected error is
'   cleared and reported as 0; an Empty result, which some VBA hosts return
'   instead of raising, is reported as 0 too (stored values are 1 or more).
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Stored   As Variant    'Value stored under the key

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'Keys are stored upper-case by KeyAdd, so the lookup upper-cases too.
        On Error Resume Next
        Stored = Lookup.Item(UCase$(Key))
        If Err.Number <> 0 Or IsEmpty(Stored) Then
            Err.Clear
            KeyIndex = 0
        Else
            KeyIndex = CLng(Stored)
        End If
        On Error GoTo 0

End Function


Public Sub KeyAdd( _
    ByVal Lookup As Collection, _
    ByVal Key As String, _
    ByVal Position As Long)
'
'==============================================================================
'                                    KeyAdd
'------------------------------------------------------------------------------
' PURPOSE
'   Store an array index under a key.
'
' INPUTS
'   Lookup: the Collection used as the index.
'   Key: lookup key; stored upper-case.
'   Position: 1-based position of the item in its array.
'
' ERROR POLICY
'   Raises the Collection's duplicate-key error if the key already exists;
'   callers check with KeyIndex first.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
        Lookup.Add Position, UCase$(Key)

End Sub


'
'------------------------------------------------------------------------------
'
'                                SHEET HELPERS
'
'------------------------------------------------------------------------------
'

Public Function GetSheet( _
    ByVal SheetName As String) _
    As Worksheet
'
'==============================================================================
'                                   GetSheet
'------------------------------------------------------------------------------
' PURPOSE
'   Return a worksheet of this workbook by tab name.
'
' INPUTS
'   SheetName: one of the SH_ constants in CORE_Config.
'
' ERROR POLICY
'   A missing sheet raises "Subscript out of range" to the caller.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' RETURN SHEET
'------------------------------------------------------------------------------
        Set GetSheet = ThisWorkbook.Worksheets(SheetName)

End Function


Public Function TableCell( _
    ByVal Sheet As Worksheet, _
    ByVal RowNumber As Long, _
    ByVal Column As Long) _
    As Range
'
'==============================================================================
'                                  TableCell
'------------------------------------------------------------------------------
' PURPOSE
'   Return the cell in a sheet row at a table column: column numbers count
'   from the first table column, right of the COL_OFFSET margin columns.
'
' INPUTS
'   Sheet: the worksheet.
'   RowNumber: the sheet row.
'   Column: the table column, 1 for the first.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        Set TableCell = Sheet.Cells(RowNumber, Column + COL_OFFSET)

End Function


Public Function NamedCell( _
    ByVal NameText As String) _
    As Range
'
'==============================================================================
'                                  NamedCell
'------------------------------------------------------------------------------
' PURPOSE
'   Return the top-left cell of a workbook name, or Nothing when the name
'   is missing or no longer refers to a cell.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target   As Range    'The name's range

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Names(NameText) raises when no such name exists; Target then stays
    'Nothing.
        On Error Resume Next
        Set Target = ThisWorkbook.Names(NameText).RefersToRange
        Err.Clear
        On Error GoTo 0
        If Not Target Is Nothing Then
            Set NamedCell = Target.Cells(1, 1)
        End If

End Function


Public Function NamedRow( _
    ByVal NameText As String) _
    As Long
'
'==============================================================================
'                                   NamedRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the sheet row of a workbook name's top-left cell; 0 when the
'   name is missing or broken.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cell   As Range    'The name's top-left cell

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        Set Cell = NamedCell(NameText)
        If Not Cell Is Nothing Then
            NamedRow = Cell.Row
        End If

End Function


Public Function TableHeaderName( _
    ByVal SheetName As String) _
    As String
'
'==============================================================================
'                               TableHeaderName
'------------------------------------------------------------------------------
' PURPOSE
'   Return the workbook name of the first header cell of a sheet's table:
'   NM_NS_HEADER and NM_TX_HEADER for the inputs, NM_RS_HEADER and the
'   others for the output tables.
'
' RETURNS
'   The name; "" for a sheet whose table is not located this way.
'
' UPDATED
'   2026-10-24
'==============================================================================
'
        Select Case SheetName
            Case SH_NS
                TableHeaderName = NM_NS_HEADER
            Case SH_TX
                TableHeaderName = NM_TX_HEADER
            Case SH_RESULTS
                TableHeaderName = NM_RS_HEADER
            Case SH_POSITIONS
                TableHeaderName = NM_PS_HEADER
            Case SH_LEGCALC
                TableHeaderName = NM_LC_HEADER
            Case SH_CHECKS
                TableHeaderName = NM_CK_HEADER
        End Select

End Function


Public Function TableHeaderRow( _
    ByVal SheetName As String) _
    As Long
'
'==============================================================================
'                                TableHeaderRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the header row of a sheet's table, located through the workbook
'   name of its first header cell (TableHeaderName), which must be in
'   column C of that sheet. The rows can therefore move: a row inserted or
'   deleted above a table moves the name with it.
'
' RETURNS
'   The row; 0 when the sheet has no such name, or the name is missing,
'   broken, on another sheet or not in column C.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderName   As String    'Workbook name of the header cell
    Dim Cell         As Range     'The named cell

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        HeaderName = TableHeaderName(SheetName)
        If Len(HeaderName) = 0 Then
            Exit Function
        End If
        Set Cell = NamedCell(HeaderName)
        If Not Cell Is Nothing Then
            If StrComp(Cell.Worksheet.Name, SheetName, vbTextCompare) = 0 And Cell.Column = COL_OFFSET + 1 Then
                TableHeaderRow = Cell.Row
            End If
        End If

End Function


Public Function TableFirstRow( _
    ByVal SheetName As String) _
    As Long
'
'==============================================================================
'                                TableFirstRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the first data row of a sheet's table: the row below its
'   headers (TableHeaderRow).
'
' RETURNS
'   The row; 0 when the header row cannot be located.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HeaderRow   As Long    'Header row

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        HeaderRow = TableHeaderRow(SheetName)
        If HeaderRow > 0 Then
            TableFirstRow = HeaderRow + 1
        End If

End Function


Public Function OutputFirstRow( _
    ByVal SheetName As String) _
    As Long
'
'==============================================================================
'                                OutputFirstRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the first data row of an output table (Results, Positions,
'   LegCalc, Checks): the row below the header cell its workbook name
'   refers to.
'
' ERROR POLICY
'   Raises ERR_OUTPUT_LAYOUT when the name is missing, broken, on another
'   sheet or not in column C: an output written at a guessed row could
'   leave rows of an earlier run under the headers. Like a missing sheet,
'   this needs the workbook rebuilt from the template.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FirstRow   As Long    'First data row; 0 when the name cannot locate it

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        FirstRow = TableFirstRow(SheetName)
        If FirstRow = 0 Then
            Err.Raise ERR_OUTPUT_LAYOUT, "CORE_Util.OutputFirstRow", _
                      "Workbook name '" & TableHeaderName(SheetName) & "' is missing or no longer refers to " & _
                      "the first header cell, in column C, of " & SheetName & _
                      " - rebuild the workbook from the template."
        End If
        OutputFirstRow = FirstRow

End Function


Public Function NamedTableCol( _
    ByVal NameText As String) _
    As Long
'
'==============================================================================
'                                NamedTableCol
'------------------------------------------------------------------------------
' PURPOSE
'   Return the column of a workbook name's top-left cell, counted as
'   TableCell counts table columns, so that the column can be passed to
'   TableCell and TableRange. Check NamedRow first: a missing name gives 0
'   here too.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cell   As Range    'The name's top-left cell

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
        Set Cell = NamedCell(NameText)
        If Not Cell Is Nothing Then
            NamedTableCol = Cell.Column - COL_OFFSET
        End If

End Function


Public Function TableRange( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal FirstColumn As Long, _
    ByVal LastRow As Long, _
    ByVal LastColumn As Long) _
    As Range
'
'==============================================================================
'                                  TableRange
'------------------------------------------------------------------------------
' PURPOSE
'   Return the block between two sheet rows and two table columns, as
'   TableCell counts them.
'
' UPDATED
'   2026-10-17
'==============================================================================
'
        Set TableRange = Sheet.Range(TableCell(Sheet, FirstRow, FirstColumn), TableCell(Sheet, LastRow, LastColumn))

End Function


Public Function UsedLastRow( _
    ByVal Sheet As Worksheet) _
    As Long
'
'==============================================================================
'                                 UsedLastRow
'------------------------------------------------------------------------------
' PURPOSE
'   Return the last row of a sheet's used range: the upper bound for every
'   data scan.
'
' INPUTS
'   Sheet: the worksheet to measure.
'
' RETURNS
'   Row number of the bottom edge of Sheet.UsedRange.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
        With Sheet.UsedRange
            UsedLastRow = .Row + .Rows.Count - 1
        End With

End Function


Public Function ColumnValues( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal LastRow As Long, _
    ByVal Column As Long) _
    As Variant
'
'==============================================================================
'                                 ColumnValues
'------------------------------------------------------------------------------
' PURPOSE
'   Read one column range in a single call, always as a 2-D array, also
'   for a single row, where Range.Value gives a scalar. Public because
'   the engine reads every input column with it.
'
' INPUTS
'   Sheet: the worksheet to read.
'   FirstRow, LastRow: rows to read, FirstRow <= LastRow.
'   Column: table column number.
'
' RETURNS
'   A 1-based array (1 To rows, 1 To 1). Element (r, 1) holds the value of
'   row FirstRow + r - 1.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Values      As Variant    'Value or array returned by Range.Value
    Dim Wrapped()   As Variant    'One-cell array built for a single-row range

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Range.Value returns a scalar, not an array, for a single cell. Wrap it
    'so that callers can always index the result as (r, 1).
        Values = TableRange(Sheet, FirstRow, Column, LastRow, Column).Value
        If IsArray(Values) Then
            ColumnValues = Values
        Else
            ReDim Wrapped(1 To 1, 1 To 1)
            Wrapped(1, 1) = Values
            ColumnValues = Wrapped
        End If

End Function


Public Function LastDataRowAny( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal ColumnCount As Long, _
    ByVal IgnoredColumn As Long) _
    As Long
'
'==============================================================================
'                                LastDataRowAny
'------------------------------------------------------------------------------
' PURPOSE
'   Find the last row with a value in any of the first ColumnCount columns,
'   so
'   that a row whose ID is blank is still found (K-SACCR #35).
'
' INPUTS
'   Sheet: the worksheet to scan.
'   FirstRow: first row of the data area.
'   ColumnCount: number of columns from the first table column.
'   IgnoredColumn: a column that does not count as data, such as a
'      free-text comment; 0 for none.
'
' RETURNS
'   Row number of the last such row; FirstRow - 1 when there is none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastUsedRow   As Long       'Last row of the used range
    Dim Values        As Variant    'Block values from FirstRow to LastUsedRow
    Dim r             As Long       'Row in Values, counted from the bottom
    Dim c             As Long       'Column in Values

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
        LastDataRowAny = FirstRow - 1
        LastUsedRow = UsedLastRow(Sheet)
        If LastUsedRow < FirstRow Then
            Exit Function
        End If
        Values = TableRange(Sheet, FirstRow, 1, LastUsedRow, ColumnCount).Value
        If Not IsArray(Values) Then
            If Not IsBlankCell(Values) Then
                LastDataRowAny = FirstRow
            End If
            Exit Function
        End If
        For r = UBound(Values, 1) To 1 Step -1
            For c = 1 To UBound(Values, 2)
                If c <> IgnoredColumn Then
                    If Not IsBlankCell(Values(r, c)) Then
                        LastDataRowAny = FirstRow + r - 1
                        Exit Function
                    End If
                End If
            Next c
        Next r

End Function


Public Function RowHasData( _
    ByRef Block As Variant, _
    ByVal RowIndex As Long, _
    ByVal ColumnCount As Long, _
    ByVal IgnoredColumn As Long) _
    As Boolean
'
'==============================================================================
'                                  RowHasData
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a row of an input block holds any value.
'
' INPUTS
'   Block: a 2-D input block; RowIndex: the row in it; ColumnCount: columns
'   to check; IgnoredColumn: a column that does not count as data; 0 for
'   none.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim c   As Long    'Column

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
        For c = 1 To ColumnCount
            If c <> IgnoredColumn Then
                If Not IsBlankCell(Block(RowIndex, c)) Then
                    RowHasData = True
                    Exit Function
                End If
            End If
        Next c

End Function


Public Function GetParam( _
    ByVal Code As String) _
    As Variant
'
'==============================================================================
'                                   GetParam
'------------------------------------------------------------------------------
' PURPOSE
'   Read one parameter value through the workbook name equal to its code.
'
' INPUTS
'   Code: a PRM_ code from CORE_Config, for example "Alpha".
'
' RETURNS
'   The stored value of the name's top-left cell (Value2, so that a
'   Currency or date format does not convert or round it, K-SACCR #196); Empty
'   when the name is missing or broken, which CORE_Engine.ValidateSchema
'   reports before any parameter is read.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Cell   As Range    'The name's top-left cell

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
        Set Cell = NamedCell(Code)
        If Not Cell Is Nothing Then
            GetParam = Cell.Value2
        End If

End Function


Public Sub ClearOutputBlock( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal ColumnCount As Long)
'
'==============================================================================
'                               ClearOutputBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Clear an output table before it is rewritten: values, bold and fill.
'
' INPUTS
'   Sheet: the output worksheet.
'   FirstRow: first data row to clear.
'   ColumnCount: number of columns from the first table column.
'
' STATE OWNERSHIP
'   Changes the cleared cells only. Number formats are left in place.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LastUsedRow   As Long    'Last row of the used range

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
        LastUsedRow = UsedLastRow(Sheet)
        If LastUsedRow < FirstRow Then
            Exit Sub
        End If
        With TableRange(Sheet, FirstRow, 1, LastUsedRow, ColumnCount)
            .ClearContents
            .Font.Bold = False
            .Interior.ColorIndex = xlNone
        End With

End Sub


Public Sub WriteBlock( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByRef Block() As Variant, _
    ByVal RowsToWrite As Long, _
    ByVal ColumnCount As Long)
'
'==============================================================================
'                                  WriteBlock
'------------------------------------------------------------------------------
' PURPOSE
'   Write the first rows of a 2-D array to a sheet in one call.
'
' INPUTS
'   Sheet: the output worksheet.
'   FirstRow: row receiving Block(1, *); the block starts in the first
'   table column.
'   Block: 1-based 2-D array; it may have more rows than RowsToWrite.
'   RowsToWrite: number of rows to write; nothing is written when it is 0.
'   ColumnCount: number of columns to write.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Assigning the array to a smaller range writes only its top-left part.
        If RowsToWrite <= 0 Then
            Exit Sub
        End If
        TableRange(Sheet, FirstRow, 1, FirstRow + RowsToWrite - 1, ColumnCount).Value = Block

End Sub


Public Sub FormatColumns( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal RowsToFormat As Long, _
    ByVal Formats As String)
'
'==============================================================================
'                                FormatColumns
'------------------------------------------------------------------------------
' PURPOSE
'   Apply a number format to each column of an output block.
'
' INPUTS
'   Sheet: the output worksheet.
'   FirstRow: first row of the block.
'   RowsToFormat: number of rows; nothing is formatted when it is 0.
'   Formats: one number format per table column from the first, separated
'      by "|". An empty entry leaves that column unchanged.
'
' UPDATED
'   2026-10-17
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Parts()   As String    'Number formats, one per column, 0-based
    Dim c         As Long      'Position in Parts; column c + 1

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
        If RowsToFormat <= 0 Then
            Exit Sub
        End If
        Parts = Split(Formats, "|")
        For c = 0 To UBound(Parts)
            If Len(Parts(c)) > 0 Then
                TableRange(Sheet, FirstRow, c + 1, FirstRow + RowsToFormat - 1, c + 1).NumberFormat = Parts(c)
            End If
        Next c

End Sub
