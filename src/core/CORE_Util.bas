Attribute VB_Name = "CORE_Util"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: CORE_Util
'------------------------------------------------------------------------------
' PURPOSE
'   Provide the small helpers the engine uses everywhere: tolerant conversion
'   of cell values to text, numbers, dates and Booleans; keyed lookups built
'   on a VBA Collection; and the sheet helpers that read parameters and input
'   columns and write output blocks.
'
' WHY THIS EXISTS
'   Cell values reach the engine as Variants that may hold error values,
'   blanks or loosely typed text, and every table is located through a
'   workbook name. One set of helpers reads, converts and locates them the
'   same way for every caller.
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides three families of helpers:
'     - conversions: SafeStr, UTxt, IsBlankCell, IsNum, ToDbl, ToSerial,
'       TryBool, TryDbl, Max2, Min2
'     - keyed lookups: KeyIndex, KeyAdd
'     - sheet helpers: GetSheet, TableCell, NamedCell, TableHeaderName,
'       TableHeaderRow, TableFirstRow, OutputFirstRow, TableRange,
'       UsedLastRow, LastDataRowAny, RowHasData, GetParam, ClearOutputBlock,
'       WriteBlock, FormatColumns
'
' ERROR POLICY
'   Conversions never raise: an error value, blank or unusable input returns
'   the documented default. KeyIndex and NamedCell contain the expected
'   lookup errors and report "not found" instead. OutputFirstRow raises
'   ERR_OUTPUT_LAYOUT when an output table cannot be located. Every other
'   error, for example a missing worksheet, propagates to the caller.
'
' DEPENDENCIES
'   CORE_Config for sheet names, layout constants, workbook names and error
'   numbers.
'
' NOTES
'   Keyed lookups use a VBA Collection, so no Scripting.Dictionary reference
'   is needed.
'
' PUBLIC SURFACE
'   None outside this VBA project. Every procedure is Public for in-project use
'   by CORE_Engine and M_Main; Option Private Module keeps them off the
'   supported external surface.
'
' STATE OWNERSHIP
'   Stateless. The sheet helpers read and write the worksheets of ThisWorkbook
'   that their caller names; they change no Excel application setting.
'
' KNOWN DEVIATION
'   This module lives in src/core but reads and writes worksheets, which
'   the repository structure reserves for src/workbook.
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
'                                 CONVERSIONS
'
'------------------------------------------------------------------------------
'

Public Function SafeStr( _
    ByVal Content As Variant) As String
'
'------------------------------------------------------------------------------
'                           CONVERT A VALUE TO TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to trimmed text without failing on error values.
'
' WHY THIS EXISTS
'   CStr raises on an error value, so raw cell values cannot be passed to it
'   directly.
'
' INPUTS
'   Content
'     Any cell value, including #N/A, #VALUE! and other error values.
'
' RETURNS
'   The trimmed text of Content; an empty string for an error value, Empty
'   or Null.
'
' BEHAVIOR
'   Maps an error value, Empty and Null to an empty string before
'   converting; any other value is converted with CStr and trimmed.
'
' ERROR POLICY
'   Does not raise for a cell value: error values, Empty and Null never
'   reach CStr.
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Map error values, Empty and Null to an empty string before converting
        If IsError(Content) Then
            SafeStr = vbNullString
        ElseIf IsEmpty(Content) Or IsNull(Content) Then
            SafeStr = vbNullString
        Else
            SafeStr = VBA.Trim$(VBA.CStr(Content))
        End If

End Function

Public Function UTxt( _
    ByVal Content As Variant) As String
'
'------------------------------------------------------------------------------
'                      CONVERT A VALUE TO UPPER-CASE TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to trimmed upper-case text, so that codes such as
'   "lent", "Lent" and " LENT " compare equal.
'
' WHY THIS EXISTS
'   Codes typed in input cells differ in case and surrounding spaces; the
'   callers compare them in one normalized form.
'
' INPUTS
'   Content
'     Any cell value.
'
' RETURNS
'   UCase$ of SafeStr(Content).
'
' BEHAVIOR
'   Converts Content with SafeStr and upper-cases the result.
'
' ERROR POLICY
'   Does not raise for a cell value (see SafeStr).
'
' DEPENDENCIES
'   SafeStr
'
' NOTES
'   None
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Upper-case the trimmed text
        UTxt = VBA.UCase$(SafeStr(Content))

End Function

Public Function IsBlankCell( _
    ByVal Content As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                            TEST A VALUE FOR BLANK
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a cell value is empty or contains only spaces.
'
' WHY THIS EXISTS
'   Gives every caller one definition of a blank cell, built on SafeStr.
'
' INPUTS
'   Content
'     Any cell value. An error value counts as blank.
'
' RETURNS
'   True when SafeStr(Content) is empty.
'
' BEHAVIOR
'   Converts Content with SafeStr and tests the length of the result.
'
' ERROR POLICY
'   Does not raise for a cell value (see SafeStr).
'
' DEPENDENCIES
'   SafeStr
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
    'Blank when the trimmed text is empty
        IsBlankCell = (VBA.Len(SafeStr(Content)) = 0)

End Function

Public Function IsNum( _
    ByVal Content As Variant) As Boolean
'
'------------------------------------------------------------------------------
'                          TEST A VALUE FOR A NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a cell value can be used as a number.
'
' WHY THIS EXISTS
'   IsNumeric alone treats Empty as numeric, so Empty and blank strings
'   have to be rejected first.
'
' INPUTS
'   Content
'     Any cell value.
'
' RETURNS
'   False for an error value, Empty or a blank string; otherwise
'   IsNumeric(Content).
'
' BEHAVIOR
'   Rejects an error value and Empty, requires a String to be non-blank,
'   and leaves every other value to IsNumeric.
'
' ERROR POLICY
'   Does not raise.
'
' DEPENDENCIES
'   VBA library only
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' TEST
'------------------------------------------------------------------------------
    'Reject error values, Empty and blank strings before IsNumeric
        If IsError(Content) Or IsEmpty(Content) Then
            IsNum = False
        ElseIf VarType(Content) = vbString Then
            IsNum = (VBA.Len(VBA.Trim$(Content)) > 0) And IsNumeric(Content)
        Else
            IsNum = IsNumeric(Content)
        End If

End Function

Public Function ToDbl( _
    ByVal Content As Variant, _
    Optional ByVal Fallback As Double = 0#) As Double
'
'------------------------------------------------------------------------------
'                         CONVERT A VALUE TO A DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to a Double, falling back to a default.
'
' WHY THIS EXISTS
'   Lets a caller read a numeric cell in one call, with the default of its
'   choice for a value that is not a usable number. TryDbl is the strict
'   reader that reports such a value instead.
'
' INPUTS
'   Content
'     Any cell value.
'
'   Fallback
'     Value returned when Content is not a usable number; 0 by default.
'
' RETURNS
'   CDbl(Content) when IsNum(Content); otherwise Fallback.
'
' BEHAVIOR
'   Tests Content with IsNum and converts it, or returns Fallback.
'
' ERROR POLICY
'   Does not raise for a cell value.
'
' DEPENDENCIES
'   IsNum
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Convert a usable number, otherwise return the fallback
        If IsNum(Content) Then
            ToDbl = VBA.CDbl(Content)
        Else
            ToDbl = Fallback
        End If

End Function

Public Function ToSerial( _
    ByVal Content As Variant) As Double
'
'------------------------------------------------------------------------------
'                       CONVERT A VALUE TO A DATE SERIAL
'------------------------------------------------------------------------------
' PURPOSE
'   Convert a cell value to a date serial number.
'
' WHY THIS EXISTS
'   Text such as 31/03/2027 or 03/31/2027 means different dates under
'   different Windows regional settings, so only an ISO text date is read.
'
' INPUTS
'   Content
'     A date, a number already holding a serial, or text in the ISO form
'     yyyy-mm-dd. Other text, such as 31/03/2027, 03/31/2027 or a serial
'     written as text, is not read: its meaning depends on the Windows
'     regional settings.
'
' RETURNS
'   The date serial as a Double; -1 when Content is missing or not a date.
'
' BEHAVIOR
'   Starts from the "missing" marker -1 and replaces it only when Content
'   is usable:
'     - an error value or Empty stays -1
'     - a Date is converted with CDbl
'     - text must match yyyy-mm-dd and name a calendar date: 2027-02-30
'       is not rolled into March
'     - any other numeric value is taken as a serial
'   Text is tested before numbers: IsNum accepts numeric text such as
'   "47756", which is not a yyyy-mm-dd date.
'
' ERROR POLICY
'   Does not raise for a cell value: unusable input returns -1.
'
' DEPENDENCIES
'   IsNum
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
    Const PROC_NAME             As String = "CORE_Util.ToSerial"

    Dim Trimmed                 As String       'Trimmed text of Content
    Dim YearPart                As Long         'Year of an ISO text date
    Dim MonthPart               As Long         'Month of an ISO text date
    Dim DayPart                 As Long         'Day of an ISO text date
    Dim CalendarDate            As Date         'The ISO text date, built from the three parts

'------------------------------------------------------------------------------
' CONVERT
'------------------------------------------------------------------------------
    'Start from the "missing" marker
        ToSerial = -1#
    'Leave an error value or Empty as missing
        If IsError(Content) Or IsEmpty(Content) Then
            Exit Function
        End If
    'Read a date, an ISO text date (a calendar date only) or a serial
        If VarType(Content) = vbDate Then
            ToSerial = VBA.CDbl(Content)
        ElseIf VarType(Content) = vbString Then
            Trimmed = VBA.Trim$(Content)
            If Not Trimmed Like "####-##-##" Then
                Exit Function
            End If
            YearPart = VBA.CLng(VBA.Left$(Trimmed, 4))
            MonthPart = VBA.CLng(VBA.Mid$(Trimmed, 6, 2))
            DayPart = VBA.CLng(VBA.Right$(Trimmed, 2))
            If YearPart < 1900 Or MonthPart < 1 Or MonthPart > 12 Or _
                DayPart < 1 Or DayPart > 31 Then
                Exit Function
            End If
            CalendarDate = DateSerial(YearPart, MonthPart, DayPart)
            If Month(CalendarDate) = MonthPart And Day(CalendarDate) = DayPart Then
                ToSerial = VBA.CDbl(CalendarDate)
            End If
        ElseIf IsNum(Content) Then
            ToSerial = VBA.CDbl(Content)
        End If

End Function

Public Function TryBool( _
    ByVal Content As Variant, _
    ByVal Fallback As Boolean, _
    ByRef Result As Boolean) As Boolean
'
'------------------------------------------------------------------------------
'                       READ A YES OR NO VALUE STRICTLY
'------------------------------------------------------------------------------
' PURPOSE
'   Read a Y/N style cell value strictly: an unrecognised value is
'   reported instead of silently taking the default.
'
' WHY THIS EXISTS
'   A mistyped flag must not silently take the default.
'
' INPUTS
'   Content
'     A Boolean, blank, or text such as Y, YES, TRUE, 1, N, NO, FALSE, 0
'     (and the Italian VERO, SI, FALSO).
'
'   Fallback
'     Value used when Content is blank.
'
'   Result
'     Receives the value read.
'
' RETURNS
'   True with Result set when Content is blank or recognised; False for an
'   error value or unrecognised text, with Result left as Fallback.
'
' BEHAVIOR
'   Sets Result to Fallback, rejects an error value, takes a Boolean as it
'   is, and otherwise matches UTxt(Content) against the recognised words.
'
' ERROR POLICY
'   Does not raise for a cell value: an unreadable value returns False.
'
' DEPENDENCIES
'   UTxt
'
' NOTES
'   The words are compared through UTxt, so case and surrounding spaces do
'   not matter.
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Start from the fallback
        Result = Fallback
    'Report an error value as unreadable
        If IsError(Content) Then
            Exit Function
        End If
    'Take a Boolean as it is
        If VarType(Content) = vbBoolean Then
            Result = Content
            TryBool = True
            Exit Function
        End If
    'Match blank text and the recognised words
        Select Case UTxt(Content)
            Case vbNullString
                TryBool = True
            Case "Y", "YES", "TRUE", "1", "VERO", "SI"
                Result = True
                TryBool = True
            Case "N", "NO", "FALSE", "0", "FALSO"
                Result = False
                TryBool = True
            Case Else
                'Leave unrecognised text reported as unreadable
        End Select

End Function

Public Function TryDbl( _
    ByVal Content As Variant, _
    ByVal Fallback As Double, _
    ByRef Result As Double) As Boolean
'
'------------------------------------------------------------------------------
'                            READ A NUMBER STRICTLY
'------------------------------------------------------------------------------
' PURPOSE
'   Read a numeric cell value strictly: unlike ToDbl, text that is not a
'   number is reported instead of silently taking the default.
'
' WHY THIS EXISTS
'   A mistyped number must not silently take the default.
'
' INPUTS
'   Content
'     A number, blank, or anything else.
'
'   Fallback
'     Value used when Content is blank.
'
'   Result
'     Receives the value read.
'
' RETURNS
'   True with Result set when Content is blank or a usable number; False
'   for an error value or non-numeric text, with Result left as Fallback.
'
' BEHAVIOR
'   Sets Result to Fallback, rejects an error value, accepts a blank value,
'   and converts a usable number with CDbl.
'
' ERROR POLICY
'   Does not raise for a cell value: an unreadable value returns False.
'
' DEPENDENCIES
'   IsBlankCell, IsNum
'
' NOTES
'   None
'
' UPDATED
'   2026-10-26
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Start from the fallback
        Result = Fallback
    'Report an error value as unreadable
        If IsError(Content) Then
            Exit Function
        End If
    'Accept a blank value and read a usable number
        If IsBlankCell(Content) Then
            TryDbl = True
        ElseIf IsNum(Content) Then
            Result = VBA.CDbl(Content)
            TryDbl = True
        End If

End Function

Public Function Max2( _
    ByVal FirstValue As Double, _
    ByVal SecondValue As Double) As Double
'
'------------------------------------------------------------------------------
'                       RETURN THE LARGER OF TWO NUMBERS
'------------------------------------------------------------------------------
' PURPOSE
'   Return the larger of two numbers.
'
' WHY THIS EXISTS
'   Gives callers a plain VBA maximum of two Doubles.
'
' INPUTS
'   FirstValue
'     First number.
'
'   SecondValue
'     Second number.
'
' RETURNS
'   The larger value; SecondValue when the two are equal.
'
' BEHAVIOR
'   Compares the two values.
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
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Return the larger value
        If FirstValue > SecondValue Then
            Max2 = FirstValue
        Else
            Max2 = SecondValue
        End If

End Function

Public Function Min2( _
    ByVal FirstValue As Double, _
    ByVal SecondValue As Double) As Double
'
'------------------------------------------------------------------------------
'                      RETURN THE SMALLER OF TWO NUMBERS
'------------------------------------------------------------------------------
' PURPOSE
'   Return the smaller of two numbers.
'
' WHY THIS EXISTS
'   Gives callers a plain VBA minimum of two Doubles.
'
' INPUTS
'   FirstValue
'     First number.
'
'   SecondValue
'     Second number.
'
' RETURNS
'   The smaller value; SecondValue when the two are equal.
'
' BEHAVIOR
'   Compares the two values.
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
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Return the smaller value
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
    ByVal Key As String) As Long
'
'------------------------------------------------------------------------------
'                            LOOK UP A KEYED INDEX
'------------------------------------------------------------------------------
' PURPOSE
'   Look up the array index stored under a key.
'
' WHY THIS EXISTS
'   A VBA Collection raises for a missing key; callers need a plain
'   "not found" result instead.
'
' INPUTS
'   Lookup
'     A Collection filled by KeyAdd.
'
'   Key
'     Lookup key; compared case-insensitively.
'
' RETURNS
'   The stored 1-based index; 0 when the key is not present.
'
' BEHAVIOR
'   Reads the upper-cased key, as KeyAdd stores it, with errors suppressed,
'   then reports a failed or Empty read as 0.
'
' ERROR POLICY
'   Collection.Item raises for a missing key. That single expected error is
'   cleared and reported as 0; an Empty result, which some VBA hosts return
'   instead of raising, is reported as 0 too (stored values are 1 or more).
'
' DEPENDENCIES
'   VBA Collection; KeyAdd stores the keys
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.KeyIndex"

    Dim Stored                  As Variant      'Value stored under the key

'------------------------------------------------------------------------------
' LOOK UP
'------------------------------------------------------------------------------
    'Read the key upper-cased, as KeyAdd stores it, trapping a missing key
        On Error Resume Next
        Stored = Lookup.Item(VBA.UCase$(Key))
    'Report a missing key or an Empty value as 0
        If Err.Number <> 0 Or IsEmpty(Stored) Then
            Err.Clear
            KeyIndex = 0
        Else
            KeyIndex = VBA.CLng(Stored)
        End If
    'Restore normal error handling
        On Error GoTo 0

End Function

Public Sub KeyAdd( _
    ByVal Lookup As Collection, _
    ByVal Key As String, _
    ByVal Position As Long)
'
'------------------------------------------------------------------------------
'                             STORE A KEYED INDEX
'------------------------------------------------------------------------------
' PURPOSE
'   Store an array index under a key.
'
' WHY THIS EXISTS
'   Keeps the upper-case key convention that KeyIndex relies on in one
'   place.
'
' INPUTS
'   Lookup
'     The Collection used as the index.
'
'   Key
'     Lookup key; stored upper-case.
'
'   Position
'     1-based position of the item in its array.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds Position to Lookup under the upper-cased key.
'
' ERROR POLICY
'   Raises the Collection's duplicate-key error if the key already exists;
'   callers check with KeyIndex first.
'
' DEPENDENCIES
'   VBA Collection
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
    'Add the position under the upper-cased key
        Lookup.Add Position, VBA.UCase$(Key)

End Sub

'
'------------------------------------------------------------------------------
'
'                                SHEET HELPERS
'
'------------------------------------------------------------------------------
'

Public Function GetSheet( _
    ByVal SheetName As String) As Worksheet
'
'------------------------------------------------------------------------------
'                          RETURN A WORKSHEET BY NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Return a worksheet of this workbook by tab name.
'
' WHY THIS EXISTS
'   Resolves sheet names against ThisWorkbook, whichever workbook is
'   active.
'
' INPUTS
'   SheetName
'     One of the SH_ constants in CORE_Config.
'
' RETURNS
'   The worksheet.
'
' BEHAVIOR
'   Looks the name up in ThisWorkbook.Worksheets.
'
' ERROR POLICY
'   A missing sheet raises "Subscript out of range" to the caller.
'
' DEPENDENCIES
'   ThisWorkbook; the SH_ constants of CORE_Config
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN SHEET
'------------------------------------------------------------------------------
    'Look the sheet up in this workbook
        Set GetSheet = ThisWorkbook.Worksheets(SheetName)

End Function

Public Function TableCell( _
    ByVal Sheet As Worksheet, _
    ByVal RowNumber As Long, _
    ByVal Column As Long) As Range
'
'------------------------------------------------------------------------------
'                             RETURN A TABLE CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Return the cell in a sheet row at a table column: column numbers count
'   from the first table column, right of the COL_OFFSET margin columns.
'
' WHY THIS EXISTS
'   Every table starts right of the margin columns, so callers count table
'   columns and this helper adds the offset.
'
' INPUTS
'   Sheet
'     The worksheet.
'
'   RowNumber
'     The sheet row.
'
'   Column
'     The table column, 1 for the first.
'
' RETURNS
'   The cell at RowNumber and sheet column Column + COL_OFFSET.
'
' BEHAVIOR
'   Adds COL_OFFSET to Column and returns Sheet.Cells at that position.
'
' ERROR POLICY
'   Errors from Excel propagate to the caller.
'
' DEPENDENCIES
'   COL_OFFSET in CORE_Config
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN CELL
'------------------------------------------------------------------------------
    'Shift the table column past the margin columns
        Set TableCell = Sheet.Cells(RowNumber, Column + COL_OFFSET)

End Function

Public Function NamedCell( _
    ByVal NameText As String) As Range
'
'------------------------------------------------------------------------------
'                      RESOLVE A WORKBOOK NAME TO A CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Return the top-left cell of a workbook name, or Nothing when the name
'   is missing or no longer refers to a cell.
'
' WHY THIS EXISTS
'   Tables and parameters are located through workbook names; callers need
'   "not found" as Nothing rather than an error.
'
' INPUTS
'   NameText
'     The workbook name.
'
' RETURNS
'   The top-left cell of the name's range; Nothing when the name is missing
'   or no longer refers to a range.
'
' BEHAVIOR
'   Reads the name's RefersToRange with errors suppressed and returns its
'   first cell when one was found.
'
' ERROR POLICY
'   Names(NameText) raises when no such name exists; that error is cleared
'   and Nothing is returned.
'
' DEPENDENCIES
'   ThisWorkbook.Names
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.NamedCell"

    Dim Target                  As Range        'The name's range

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Read the name's range; Target stays Nothing when the name is missing
        On Error Resume Next
        Set Target = ThisWorkbook.Names(NameText).RefersToRange
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0
    'Return the top-left cell of a resolved range
        If Not Target Is Nothing Then
            Set NamedCell = Target.Cells(1, 1)
        End If

End Function

Public Function TableHeaderName( _
    ByVal SheetName As String) As String
'
'------------------------------------------------------------------------------
'                       NAME THE HEADER CELL OF A TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Return the workbook name of the first header cell of a sheet's table:
'   NM_NS_HEADER and NM_TX_HEADER for the inputs, NM_RS_HEADER and the
'   others for the output tables.
'
' WHY THIS EXISTS
'   Tables are located through the workbook name of their first header
'   cell, so their rows can move.
'
' INPUTS
'   SheetName
'     One of the SH_ constants in CORE_Config.
'
' RETURNS
'   The name; an empty string for a sheet whose table is not located this
'   way.
'
' BEHAVIOR
'   Maps SH_NS, SH_TX, SH_RESULTS, SH_POSITIONS, SH_LEGCALC and SH_CHECKS
'   to their NM_*_HEADER name.
'
' ERROR POLICY
'   None
'
' DEPENDENCIES
'   The SH_ and NM_*_HEADER constants of CORE_Config
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
    'Map the sheet to the name of its first header cell
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
            Case Else
                'Leave the name empty for any other sheet
        End Select

End Function

Public Function TableHeaderRow( _
    ByVal SheetName As String) As Long
'
'------------------------------------------------------------------------------
'                       LOCATE THE HEADER ROW OF A TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Return the header row of a sheet's table, located through the workbook
'   name of its first header cell (TableHeaderName), which must be in
'   column C of that sheet. The rows can therefore move: a row inserted or
'   deleted above a table moves the name with it.
'
' WHY THIS EXISTS
'   A fixed header row would break when rows are inserted or deleted above
'   a table.
'
' INPUTS
'   SheetName
'     One of the SH_ constants in CORE_Config.
'
' RETURNS
'   The row; 0 when the sheet has no such name, or the name is missing,
'   broken, on another sheet or not in column C.
'
' BEHAVIOR
'   Looks up the header name, resolves it with NamedCell, and accepts the
'   cell only on the same sheet (compared case-insensitively) and in the
'   first table column.
'
' ERROR POLICY
'   Does not raise for a missing or broken name (see NamedCell).
'
' DEPENDENCIES
'   TableHeaderName, NamedCell, COL_OFFSET in CORE_Config
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.TableHeaderRow"

    Dim HeaderName              As String       'Workbook name of the header cell
    Dim Cell                    As Range        'The named cell

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Stop for a sheet whose table has no header name
        HeaderName = TableHeaderName(SheetName)
        If VBA.Len(HeaderName) = 0 Then
            Exit Function
        End If
    'Accept the named cell only on this sheet and in the first table column
        Set Cell = NamedCell(HeaderName)
        If Not Cell Is Nothing Then
            If VBA.StrComp(Cell.Worksheet.Name, SheetName, vbTextCompare) = 0 And _
                Cell.Column = COL_OFFSET + 1 Then
                TableHeaderRow = Cell.Row
            End If
        End If

End Function

Public Function TableFirstRow( _
    ByVal SheetName As String) As Long
'
'------------------------------------------------------------------------------
'                     LOCATE THE FIRST DATA ROW OF A TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Return the first data row of a sheet's table: the row below its
'   headers (TableHeaderRow).
'
' WHY THIS EXISTS
'   Callers read and write data rows, not the header row.
'
' INPUTS
'   SheetName
'     One of the SH_ constants in CORE_Config.
'
' RETURNS
'   The row; 0 when the header row cannot be located.
'
' BEHAVIOR
'   Adds one to the header row when it was found.
'
' ERROR POLICY
'   Does not raise for a missing or broken name (see TableHeaderRow).
'
' DEPENDENCIES
'   TableHeaderRow
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.TableFirstRow"

    Dim HeaderRow               As Long         'Header row

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Return the row below a located header row
        HeaderRow = TableHeaderRow(SheetName)
        If HeaderRow > 0 Then
            TableFirstRow = HeaderRow + 1
        End If

End Function

Public Function OutputFirstRow( _
    ByVal SheetName As String) As Long
'
'------------------------------------------------------------------------------
'                   LOCATE THE FIRST ROW OF AN OUTPUT TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Return the first data row of an output table (Results, Positions,
'   LegCalc, Checks): the row below the header cell its workbook name
'   refers to.
'
' WHY THIS EXISTS
'   An output written at a guessed row could leave rows of an earlier run
'   under the headers, so an output table that cannot be located stops
'   the run instead.
'
' INPUTS
'   SheetName
'     SH_RESULTS, SH_POSITIONS, SH_LEGCALC or SH_CHECKS.
'
' RETURNS
'   The first data row.
'
' BEHAVIOR
'   Locates the row with TableFirstRow and raises when it is 0.
'
' ERROR POLICY
'   Raises ERR_OUTPUT_LAYOUT when the name is missing, broken, on another
'   sheet or not in column C: an output written at a guessed row could
'   leave rows of an earlier run under the headers. Like a missing sheet,
'   this needs the workbook rebuilt from the template.
'
' DEPENDENCIES
'   TableFirstRow, TableHeaderName, ERR_OUTPUT_LAYOUT in CORE_Config
'
' NOTES
'   The error source is written as the literal "CORE_Util.OutputFirstRow",
'   the same value as PROC_NAME.
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.OutputFirstRow"

    Dim FirstRow                As Long         'First data row; 0 when the name cannot locate it

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Raise when the output table cannot be located
        FirstRow = TableFirstRow(SheetName)
        If FirstRow = 0 Then
            Err.Raise ERR_OUTPUT_LAYOUT, "CORE_Util.OutputFirstRow", _
                "Workbook name '" & TableHeaderName(SheetName) & _
                "' is missing or no longer refers to " & _
                "the first header cell, in column C, of " & SheetName & _
                " - rebuild the workbook from the template."
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the located row
        OutputFirstRow = FirstRow

End Function

Public Function TableRange( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal FirstColumn As Long, _
    ByVal LastRow As Long, _
    ByVal LastColumn As Long) As Range
'
'------------------------------------------------------------------------------
'                        RETURN A BLOCK OF TABLE CELLS
'------------------------------------------------------------------------------
' PURPOSE
'   Return the block between two sheet rows and two table columns, as
'   TableCell counts them.
'
' WHY THIS EXISTS
'   Callers address blocks in table columns, like single cells.
'
' INPUTS
'   Sheet
'     The worksheet.
'
'   FirstRow
'     First sheet row of the block.
'
'   FirstColumn
'     First table column of the block, 1 for the first.
'
'   LastRow
'     Last sheet row of the block.
'
'   LastColumn
'     Last table column of the block.
'
' RETURNS
'   The range between the two corner cells.
'
' BEHAVIOR
'   Builds the two corner cells with TableCell and returns the range
'   between them.
'
' ERROR POLICY
'   Errors from Excel propagate to the caller.
'
' DEPENDENCIES
'   TableCell
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN RANGE
'------------------------------------------------------------------------------
    'Join the two corner cells
        Set TableRange = Sheet.Range(TableCell(Sheet, FirstRow, FirstColumn), _
            TableCell(Sheet, LastRow, LastColumn))

End Function

Public Function UsedLastRow( _
    ByVal Sheet As Worksheet) As Long
'
'------------------------------------------------------------------------------
'                       MEASURE THE USED ROWS OF A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Return the last row of a sheet's used range: the upper bound for every
'   data scan.
'
' WHY THIS EXISTS
'   Scans stop at the used range instead of the last row of the sheet.
'
' INPUTS
'   Sheet
'     The worksheet to measure.
'
' RETURNS
'   Row number of the bottom edge of Sheet.UsedRange.
'
' BEHAVIOR
'   Adds the row count of the used range to its first row, less one.
'
' ERROR POLICY
'   Errors from Excel propagate to the caller.
'
' DEPENDENCIES
'   Worksheet.UsedRange
'
' NOTES
'   None
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
    'Bottom edge of the used range
        With Sheet.UsedRange
            UsedLastRow = .Row + .Rows.Count - 1
        End With

End Function

Public Function LastDataRowAny( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal ColumnCount As Long, _
    ByVal IgnoredColumn As Long) As Long
'
'------------------------------------------------------------------------------
'                         FIND THE LAST ROW WITH DATA
'------------------------------------------------------------------------------
' PURPOSE
'   Find the last row with a value in any of the first ColumnCount columns,
'   so that a row whose ID is blank is still found.
'
' WHY THIS EXISTS
'   A scan on the ID column alone would miss a row whose ID is blank.
'
' INPUTS
'   Sheet
'     The worksheet to scan.
'
'   FirstRow
'     First row of the data area.
'
'   ColumnCount
'     Number of columns from the first table column.
'
'   IgnoredColumn
'     A column that does not count as data, such as a free-text comment;
'     0 for none.
'
' RETURNS
'   Row number of the last such row; FirstRow - 1 when there is none.
'
' BEHAVIOR
'   Reads the block from FirstRow to the last used row in one call, then
'   scans it upward and stops at the first row holding a value outside
'   IgnoredColumn. A one-cell block is read as a single value.
'
' ERROR POLICY
'   Errors from Excel propagate to the caller.
'
' DEPENDENCIES
'   UsedLastRow, TableRange, IsBlankCell
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
    Const PROC_NAME             As String = "CORE_Util.LastDataRowAny"

    Dim LastUsedRow             As Long         'Last row of the used range
    Dim Values                  As Variant      'Block values from FirstRow to LastUsedRow
    Dim BlockRow                As Long         'Row in Values, counted from the bottom
    Dim Col                     As Long         'Column in Values

'------------------------------------------------------------------------------
' SCAN
'------------------------------------------------------------------------------
    'Start from "no data row"
        LastDataRowAny = FirstRow - 1
    'Stop when the used range ends above the data area
        LastUsedRow = UsedLastRow(Sheet)
        If LastUsedRow < FirstRow Then
            Exit Function
        End If
    'Read the block in one call
        Values = TableRange(Sheet, FirstRow, 1, LastUsedRow, ColumnCount).Value
    'Test a one-cell block, which is read as a single value
        If Not IsArray(Values) Then
            If Not IsBlankCell(Values) Then
                LastDataRowAny = FirstRow
            End If
            Exit Function
        End If
    'Scan upward for the first row with a value outside the ignored column
        For BlockRow = UBound(Values, 1) To 1 Step -1
            For Col = 1 To UBound(Values, 2)
                If Col <> IgnoredColumn Then
                    If Not IsBlankCell(Values(BlockRow, Col)) Then
                        LastDataRowAny = FirstRow + BlockRow - 1
                        Exit Function
                    End If
                End If
            Next Col
        Next BlockRow

End Function

Public Function RowHasData( _
    ByRef Block As Variant, _
    ByVal RowIndex As Long, _
    ByVal ColumnCount As Long, _
    ByVal IgnoredColumn As Long) As Boolean
'
'------------------------------------------------------------------------------
'                          TEST A BLOCK ROW FOR DATA
'------------------------------------------------------------------------------
' PURPOSE
'   Tell whether a row of an input block holds any value.
'
' WHY THIS EXISTS
'   Rows with no value are skipped, whichever column holds the data.
'
' INPUTS
'   Block
'     A 2-D input block.
'
'   RowIndex
'     The row in it.
'
'   ColumnCount
'     Columns to check.
'
'   IgnoredColumn
'     A column that does not count as data; 0 for none.
'
' RETURNS
'   True when a column other than IgnoredColumn holds a non-blank value;
'   False otherwise.
'
' BEHAVIOR
'   Tests columns 1 to ColumnCount with IsBlankCell and stops at the first
'   non-blank one.
'
' ERROR POLICY
'   Does not raise for a cell value. A RowIndex or ColumnCount outside
'   Block raises "Subscript out of range" to the caller.
'
' DEPENDENCIES
'   IsBlankCell
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
    Const PROC_NAME             As String = "CORE_Util.RowHasData"

    Dim Col                     As Long         'Column

'------------------------------------------------------------------------------
' CHECK
'------------------------------------------------------------------------------
    'Stop at the first non-blank column outside the ignored one
        For Col = 1 To ColumnCount
            If Col <> IgnoredColumn Then
                If Not IsBlankCell(Block(RowIndex, Col)) Then
                    RowHasData = True
                    Exit Function
                End If
            End If
        Next Col

End Function

Public Function GetParam( _
    ByVal Code As String) As Variant
'
'------------------------------------------------------------------------------
'                            READ A NAMED PARAMETER
'------------------------------------------------------------------------------
' PURPOSE
'   Read one parameter value through the workbook name equal to its code.
'
' WHY THIS EXISTS
'   Parameters are located by workbook name, so their cells can move.
'
' INPUTS
'   Code
'     A PRM_ code from CORE_Config, for example "AsOfDate" (PRM_ASOF).
'
' RETURNS
'   The stored value of the name's top-left cell (Value2, so that a
'   Currency or date format does not convert or round it); Empty
'   when the name is missing or broken, which CORE_Engine.ValidateSchema
'   reports before any parameter is read.
'
' BEHAVIOR
'   Resolves the name with NamedCell and reads Value2 of the cell found.
'
' ERROR POLICY
'   Does not raise for a missing or broken name (see NamedCell).
'
' DEPENDENCIES
'   NamedCell; the PRM_ constants of CORE_Config
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
    Const PROC_NAME             As String = "CORE_Util.GetParam"

    Dim Cell                    As Range        'The name's top-left cell

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Read the stored value of a resolved name
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
'------------------------------------------------------------------------------
'                            CLEAR AN OUTPUT TABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Clear an output table before it is rewritten: values, bold and fill.
'
' WHY THIS EXISTS
'   Rows of an earlier run must not remain below a shorter new output.
'
' INPUTS
'   Sheet
'     The output worksheet.
'
'   FirstRow
'     First data row to clear.
'
'   ColumnCount
'     Number of columns from the first table column.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears contents, bold and fill from FirstRow to the last used row;
'   does nothing when the used range ends above FirstRow.
'
' ERROR POLICY
'   Errors from Excel propagate to the caller.
'
' DEPENDENCIES
'   UsedLastRow, TableRange
'
' NOTES
'   None
'
' STATE OWNERSHIP
'   Changes the cleared cells only. Number formats are left in place.
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "CORE_Util.ClearOutputBlock"

    Dim LastUsedRow             As Long         'Last row of the used range

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Stop when the used range ends above the output rows
        LastUsedRow = UsedLastRow(Sheet)
        If LastUsedRow < FirstRow Then
            Exit Sub
        End If
    'Clear values, bold and fill; keep number formats
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
'------------------------------------------------------------------------------
'                          WRITE AN ARRAY TO A SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Write the first rows of a 2-D array to a sheet in one call.
'
' WHY THIS EXISTS
'   Callers build an output table in memory and write it in one
'   assignment instead of cell by cell.
'
' INPUTS
'   Sheet
'     The output worksheet.
'
'   FirstRow
'     Row receiving Block(1, *); the block starts in the first table
'     column.
'
'   Block
'     1-based 2-D array; it may have more rows than RowsToWrite.
'
'   RowsToWrite
'     Number of rows to write; nothing is written when it is 0.
'
'   ColumnCount
'     Number of columns to write.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Assigns Block to the range of RowsToWrite rows and ColumnCount columns
'   from FirstRow.
'
' ERROR POLICY
'   Errors from Excel propagate to the caller.
'
' DEPENDENCIES
'   TableRange
'
' NOTES
'   Assigning the array to a smaller range writes only its top-left part.
'
' UPDATED
'   2026-10-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Write nothing for an empty block
        If RowsToWrite <= 0 Then
            Exit Sub
        End If
    'Assign the top-left part of the array in one call
        TableRange(Sheet, FirstRow, 1, FirstRow + RowsToWrite - 1, ColumnCount).Value = Block

End Sub

Public Sub FormatColumns( _
    ByVal Sheet As Worksheet, _
    ByVal FirstRow As Long, _
    ByVal RowsToFormat As Long, _
    ByVal Formats As String)
'
'------------------------------------------------------------------------------
'                         APPLY COLUMN NUMBER FORMATS
'------------------------------------------------------------------------------
' PURPOSE
'   Apply a number format to each column of an output block.
'
' WHY THIS EXISTS
'   Each output table passes its column formats as one string.
'
' INPUTS
'   Sheet
'     The output worksheet.
'
'   FirstRow
'     First row of the block.
'
'   RowsToFormat
'     Number of rows; nothing is formatted when it is 0.
'
'   Formats
'     One number format per table column from the first, separated by "|".
'     An empty entry leaves that column unchanged.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Splits Formats on "|" and sets the NumberFormat of each column that has
'   a non-empty entry.
'
' ERROR POLICY
'   Errors from Excel propagate to the caller.
'
' DEPENDENCIES
'   TableRange
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
    Const PROC_NAME             As String = "CORE_Util.FormatColumns"

    Dim Parts()                 As String       'Number formats, one per column, 0-based
    Dim Col                     As Long         'Position in Parts; column Col + 1

'------------------------------------------------------------------------------
' FORMAT
'------------------------------------------------------------------------------
    'Format nothing for an empty block
        If RowsToFormat <= 0 Then
            Exit Sub
        End If
    'Apply each non-empty format to its column
        Parts = Split(Formats, "|")
        For Col = 0 To UBound(Parts)
            If VBA.Len(Parts(Col)) > 0 Then
                TableRange(Sheet, FirstRow, Col + 1, _
                    FirstRow + RowsToFormat - 1, Col + 1).NumberFormat = Parts(Col)
            End If
        Next Col

End Sub
