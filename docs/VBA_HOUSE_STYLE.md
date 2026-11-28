# ✍️ VBA house style

This document is authoritative for **how K-SFTs's VBA is written**:
naming, module rules, procedure contracts, errors, Excel-state handling and
source presentation. Where a component lives is owned by
[REPOSITORY_STRUCTURE.md](REPOSITORY_STRUCTURE.md); the supported hosts and
the import and export steps by [INSTALLATION.md](../INSTALLATION.md).

**Status on 2026-11-28:** the VBA components are the test modules in
`tests/modules/` and the five prototype modules in `src/`
([PROTOTYPE.md](methodology/PROTOTYPE.md)), which pass the gates below but
depart from the module rules where PROTOTYPE.md lists it: `src/core`
modules read and write sheets. These rules apply to every component. By my
decision of 2026-11-23 ([#307](https://github.com/danielep71/K-SFTs/issues/307)),
the naming, procedure contract, layout and statement rules below follow my
`vba-house-style` conventions, and every module was reformatted to them
without a change of behaviour; names that existed before that date keep
their spelling. On 2026-11-28 the modules and this guide were aligned with
the skill's updated checks
([#311](https://github.com/danielep71/K-SFTs/issues/311)). The `vba-*` gates of `python tools/check.py` enforce the
mechanical rules: component names and `VB_Name`, module options, ASCII,
tabs and trailing white space, `PtrSafe`, jump targets, explicit visibility
of whatever VBA would otherwise make public, and the public API inventory.
Everything else is enforced by review
([what the checks cannot see](../tools/README.md#vba-static-checks)).

<a id="naming"></a>

## 🏷️ Naming

The VBE lists every component in one flat project, so a component's name
carries its role in an upper-case prefix.

| Element | Convention | Example |
| --- | --- | --- |
| Core module (`src/core/`) | `CORE_` + subject | `CORE_Validation` |
| Public facade module (`src/modules/`) | `SFT_` + subject | `SFT_Functions` |
| Class (`src/classes/`) | `C` + noun | `CCollateralItem` |
| Workbook module (`src/workbook/`) | `ThisWorkbook`; sheets by code name `sh` + noun; standard modules `M_` + subject | `shInputs`, `M_Main` |
| Form (`src/forms/`) | `frm` + noun | `frmRunOptions` |
| Test module (`tests/modules/`) | `TEST_` + subject | `TEST_Harness` |
| Example module (`examples/modules/`) | `Example` + subject | `ExampleSingleNettingSet` |
| Worksheet function (`src/modules/`) | `SFT_` + quantity | `SFT_ExposureAmount` |
| Procedure in a standard module | role prefix of its module + `_` + area + `_` + verb and object | `CORE_Collateral_ValidateHaircut` |
| Class member, event handler | PascalCase; event handlers keep the name VBA requires | `MarketValue`, `Workbook_Open` |
| Parameter, local variable, loop counter | PascalCase noun, spelled out | `MarketValue`, `RowIndex` |
| Module-level variable | `m` + PascalCase | `mRunState` |
| Public variable | `gSFT_` + PascalCase | `gSFT_LastError` |
| Constant | `SFT_` + UPPER_SNAKE_CASE | `SFT_MAX_NETTING_SETS` |
| Error number constant | `SFT_ERR_` + UPPER_SNAKE_CASE | `SFT_ERR_INVALID_INPUT` |
| Enum or user-defined type / member | `SFT_` + PascalCase / `SFT_` + PascalCase | `SFT_ProductType` / `SFT_ProductRepo` |

The examples show the form of a name only; none of these components exists,
and no product, quantity or parameter is decided by them. The role prefix of
a module is the part before its first underscore (`CORE`, `M`, `TEST`);
`SFT_` is the project prefix. Procedures, constants and variables that
existed on 2026-11-23 keep their names: renaming a public name is a contract
change, and the reformatting of #307 renamed none.

- Names are English and spelled out. Abbreviations are limited to
  established terms of the documented method, written in full in the
  procedure's comment the first time they appear.
- No single-letter names and no type prefixes (`strName`, `lngCount`),
  loop counters included: every variable and parameter names what it holds
  (`RowIndex`, `Formatted`), also in a short helper.
- Worksheet function arguments use the terms of the documented method once
  the framework is chosen ([D1](REPOSITORY_STRUCTURE.md#-open-product-and-methodology-decisions)).
- VBA names are case-insensitive and the VBE shows one spelling per name
  across the project, so the source files, not the editor, carry the
  convention. The component name (`Attribute VB_Name`) must equal the file
  name without its extension.

## ⚙️ Module rules

- Every module declares `Option Explicit`.
- Every standard module in `src/core/` and `tests/modules/` declares
  `Option Private Module`, so its `Public` procedures are visible inside the
  project only.
- No `Option Base 1`: arrays are declared with explicit bounds
  (`Dim Values(1 To ValueCount)`).
- No `Variant` where a specific type works. Arguments are `ByVal` unless the
  procedure's contract is to write an output parameter.
- `src/core/` never touches the Excel object model (`Application`,
  workbooks, sheets, ranges) and holds no mutable module-level state: a core
  procedure depends only on its arguments
  ([dependency direction](REPOSITORY_STRUCTURE.md#-dependency-direction)).
- A `Declare` statement, if ever needed, uses `PtrSafe` and `LongPtr`,
  because 64-bit Excel is the
  [supported target](../INSTALLATION.md#supported-hosts).

<a id="procedure-contracts"></a>

## 📜 Procedure contracts

Every procedure documents its contract in a header immediately after its
signature, and every module in its component header. Both use the same nine
fields, in this order, none omitted: a field with nothing to say holds
`None`.

| Field | Explain |
| --- | --- |
| `PURPOSE` | The responsibility it owns, in one or two sentences |
| `WHY THIS EXISTS` | The problem it solves, or why it is shaped this way |
| `INPUTS` | Each argument on its own line, named as in the signature, then its meaning, **unit**, **valid domain** and ownership; `None` when there is none; exactly `None at module level` in a component header |
| `RETURNS` | Meaning and unit of the result; exactly `Nothing` for a `Sub`, whose output parameters are described under `INPUTS` |
| `BEHAVIOR` | What happens, in order |
| `ERROR POLICY` | Errors raised, propagated or contained, and cleanup |
| `DEPENDENCIES` | Object models, components and procedures it relies on |
| `NOTES` | Design decisions, caveats and call examples |
| `UPDATED` | Date of the last substantive change, `YYYY-MM-DD` |

- **Extra sections** that this repository needs come after `NOTES` and
  before `UPDATED`, which is always last: `STATE OWNERSHIP` (Excel or module
  state read, changed or restored), `REFERENCE` (the methodology entry, by
  source ID and locator, that the procedure implements), and in a component
  header `PUBLIC SURFACE`, `COMPATIBILITY` and `AUTHOR`.
- **Compact header.** A `Private` procedure of at most 15 code lines may
  keep only `PURPOSE`, `UPDATED` and the fields that say something. `Public`,
  `Friend` and event procedures always carry all nine.
- **Facts only.** A field states what the code and the methodology show;
  where the reason for a choice is unknown, it says so rather than guessing.
- A change of behaviour sets `UPDATED` to its date; a formatting-only change
  keeps the existing date.

### Units and domains

- Amounts are `Double` in a stated currency, never scaled (no thousands or
  millions). Which currency an amount is in (transaction, collateral or
  calculation currency) is part of its `INPUTS` entry.
- Currencies are three-letter ISO 4217 codes held as `String`.
- Rates, haircuts and other proportions are decimals: `0.05`, never `5`
  meaning 5%.
- Dates are VBA `Date`. A period is a `Double` in years unless its name
  states another unit (`HoldingPeriodBusinessDays`).
- Every numeric input states its valid domain, for example
  `MarketValue >= 0` or `0 <= Haircut < 1`.
- Day-count, holding-period, remargining and every other regulatory
  convention belong to the methodology
  ([methodology](methodology/README.md),
  [open decisions](REPOSITORY_STRUCTURE.md#-open-product-and-methodology-decisions)).
  A procedure cites
  the convention it implements in `REFERENCE`; it never chooses one silently.

<a id="errors"></a>

### Return values and errors

- **Error numbers** are named error constants (`SFT_ERR_` for new ones) in the range
  `vbObjectError + 2048` to `vbObjectError + 4095`, defined once in a core
  module and re-exposed by the facade when callers need them. The test
  modules use `vbObjectError + 4064` to `vbObjectError + 4095`; production
  code leaves those free.
- **Raise** with `Err.Raise number, source, description`. The source names
  the component and procedure (`"CORE_Validation.ValidateCollateral"`); the
  description states what was invalid and the offending value.
- **Core procedures and macros raise.** Worksheet functions in the facade do
  not: for an invalid or missing input they return a worksheet error value
  (`CVErr`). Each function's contract names the error value it returns for
  each condition. No procedure substitutes a default for an invalid input
  ([input and output contract](REPOSITORY_STRUCTURE.md#input-and-output-contract)).
- **`On Error Resume Next`** is allowed only around a single statement whose
  error is checked on the next line, followed by `On Error GoTo 0` or the
  procedure's handler.
- **No silent errors.** A procedure either handles an error completely or
  propagates it unchanged.

<a id="public-and-internal-procedures"></a>

### Public and internal procedures

- Procedures are `Private` unless another component calls them.
- A `Public` procedure in `src/core/` is internal: `Option Private Module`
  keeps it off the workbook's external surface.
- The supported API is only what
  [PUBLIC_API.txt](PUBLIC_API.txt) lists, from `src/modules/`
  ([public API boundary](REPOSITORY_STRUCTURE.md#-public-api-boundary)).
  Adding, removing or changing a listed declaration's name, arguments,
  defaults, units, return value or errors is a contract change: it needs its
  own issue, updates the inventory and the regression tests in the same pull
  request, and gets a changelog entry once releases exist.

<a id="operation-and-cleanup"></a>

## 🧹 Operation and cleanup

A procedure that changes Excel application state owns restoring it, and only
it.

1. **Capture** the values it will change before changing them
   (`Calculation`, `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`,
   `Cursor`, `StatusBar`).
2. **Change** only what the operation needs.
3. **Restore** in one cleanup block reached on both the success and the error
   path. It restores only the settings this procedure changed, to the values
   it captured. A setting the caller changed is the caller's to restore.
4. **Preserve the primary error.** Copy `Err.Number`, `Err.Source` and
   `Err.Description` into locals before cleanup runs, and re-raise those
   after cleanup. A failure inside cleanup never replaces the original error.
5. **Report cleanup failures.** Each restore is attempted even if another
   failed. A helper per setting contains its own error and returns what
   failed as text, empty on success. The failures are returned to the
   caller, also when the primary error is re-raised. When only cleanup
   failed, the procedure raises its cleanup-failure error (`SFT_ERR_CLEANUP_FAILED`
   in the sketch), never success.

The reference implementation is `RunExampleOperation` in
[tests/modules/TEST_Examples.bas](../tests/modules/TEST_Examples.bas), with
cases for the success path, a primary error with failed cleanup, and a
cleanup-only failure. The sketch below shows the pattern.

```vb
'In a standard module named M_Example (sketch only)
Private Const SFT_ERR_CLEANUP_FAILED     As Long = vbObjectError + 2049   'Only cleanup failed

Public Sub M_Example_RunOperation( _
    ByRef CleanupDetails As String)
'
'------------------------------------------------------------------------------
'                         RUN AN OPERATION AND RESTORE
'------------------------------------------------------------------------------
' PURPOSE ... (the nine fields) ... UPDATED
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Example.M_Example_RunOperation"

    Dim SavedCalculation        As XlCalculation    'Calculation mode before the operation
    Dim SavedScreen             As Boolean          'ScreenUpdating before the operation
    Dim StateCaptured           As Boolean          'True once both were read
    Dim ErrNumber               As Long             'Primary error number; 0 for none
    Dim ErrSource               As String           'Primary error source
    Dim ErrDescription          As String           'Primary error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start with no cleanup failure and route errors to the handler
        CleanupDetails = vbNullString
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CHANGE APPLICATION STATE
'------------------------------------------------------------------------------
    'Capture what this procedure will change, then change it
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

    '... work ...

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
CleanUp:
    'Restore only what this procedure changed, keeping each failure
        On Error GoTo 0
        If StateCaptured Then
            CleanupDetails = AppendDetail(RestoreCalculation(SavedCalculation), _
                RestoreScreenUpdating(SavedScreen))
        End If
    'Re-raise the primary error, else report a cleanup-only failure
        If ErrNumber <> 0 Then Err.Raise ErrNumber, ErrSource, ErrDescription
        If VBA.Len(CleanupDetails) > 0 Then
            Err.Raise SFT_ERR_CLEANUP_FAILED, PROC_NAME, CleanupDetails
        End If
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the primary error before cleanup can replace it
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
    'Restore state on the error path too
        Resume CleanUp

End Sub
```

`RestoreCalculation` and `RestoreScreenUpdating` each contain their own error
and return the failure as text, empty on success; `AppendDetail` joins two
failure texts with `"; "`, skipping empty ones.

Core procedures never change Excel state, so they need none of this.

## 🧱 Module layout

1. The `Attribute` lines of the export, as the VBE writes them.
2. One blank line, `Option Explicit` and, where required,
   `Option Private Module`; one blank line.
3. The component header: a lone `'` line, a banner, `' MODULE: <Name>`
   (`' CLASS: <Name>` for a class; forms and document modules use
   `MODULE:`), a banner, the nine fields with any extra sections, and a
   closing banner.
4. The declarations, each kind under its own section banner, in this order
   and only when present: `PRIVATE TYPES`, `WINDOWS API DECLARATIONS`,
   `PRIVATE CONSTANTS`, `PUBLIC CONSTANTS`, `PUBLIC ENUMS`, `PUBLIC TYPES`,
   `PUBLIC STATE`, `PRIVATE STATE`. A module may add a section named for
   what it holds (for example `SHEET NAMES`). Module state is never allowed
   in `src/core/`.
5. The procedures, families of them separated by group banners.

A banner is an apostrophe followed by 78 dashes, 79 columns in all. A test
module may draw its component header with `=` instead. Three kinds:

- a **section banner**: a banner, `' ` and an upper-case title, a banner;
- a **procedure title**: a banner, an upper-case verb phrase centred in the
  78 columns after the apostrophe, a banner;
- a **group banner**: a lone `'`, a banner, a lone `'`, a centred
  upper-case title, a lone `'`, a banner, a lone `'`.

Source decorations are ASCII only. A generated module is exempt from the
banner rules; its header says it is generated, names the generator and says
it must not be edited by hand.

## 📐 Procedure layout

- **Signature:** explicit `Public`, `Private` or `Friend`; `ByVal` or `ByRef`
  and an `As` type on every parameter; an `As` type on every function. A
  signature too long for one line puts each parameter on its own line,
  indented four, ending in `, _`; the closing parenthesis and the function's
  `As` type follow the last parameter.
- **Header:** after the signature, one blank line (none after a multi-line
  signature), a lone `'`, the procedure title, the contract fields, a
  closing banner, one blank line.
- **Body:** section banners in this order, each only when it has content:
  `DECLARE` (first `Const PROC_NAME As String = "<Module>.<Procedure>"`, used
  as the source of every error raised, then each `Dim` with `As` and a
  trailing comment aligned), `INITIALIZE` (`On Error GoTo ErrorHandler` and
  initial state), task sections named by a short upper-case verb phrase
  (`VALIDATE INPUTS`, `CALCULATE EXPOSURE`), `EXIT PROCEDURE` (cleanup and
  the clean `Exit Sub` or `Exit Function`), `ERROR HANDLER` (the label, then
  capture, cleanup and re-raise). One blank line goes above a section banner
  and none below it: the first line of the section follows the closing
  banner line directly.
- **Statements:** each statement, or small group of related statements, is
  preceded by a one-line intent comment at four spaces (sentence case, no
  space after the apostrophe, no final period); the code sits at eight
  spaces and each nested block adds four. `Else`, `ElseIf` and `End If` sit
  at the column of their `If`; `Next`, `Loop`, `End With` and `End Select`
  at the column of their opening statement; each `Case` is indented four
  from `Select Case` and its statements four more. A conditional-compilation
  directive (`#If`, `#Else`, `#End If`) opens no block: the statements it
  encloses stay at the level of the code around it. Labels
  (`ErrorHandler:`, `CleanUp:`) sit at column one.
- **Declarations:** everything under a module-level section banner is
  indented four; within a block, `As` and the trailing comment fall in
  aligned columns on multiples of four, and every declaration has a trailing
  comment stating its role or unit.
- **Spacing:** one blank line between `End Sub`, `End Function` or
  `End Property` and what follows, a group banner included; two between the
  declarations and the first group banner. Four-space indent unit, no tabs, no trailing white
  space.
- **Lines:** aim for 100 columns, never more than 120; a string literal's
  content and an aligned trailing comment do not count. A continuation line
  breaks after an operator or comma and is indented four more than its
  statement; `Err.Raise` keeps number and source on the first line and the
  message on the next.

## 🧾 Statements

- The `VBA.` qualifier on library calls (`VBA.Len`, `VBA.Trim$`,
  `VBA.CStr`) is optional. Write it where a reference may be lost (a bare
  `Len` then stops compiling with "Can't find project or library", while
  `VBA.Len` still compiles) or where a project name collides with a library
  function (`Format`, `Replace`); otherwise follow the file you edit. The
  modules of this repository use it. LibreOffice Basic does not know the
  `VBA` library name, so `tools/lo_regression.py` strips the qualifier from
  its in-memory copy; the source keeps it. Prefer the `$` string functions.
- Write `vbNullString` instead of `""`, and test emptiness with
  `VBA.Len(Text) = 0`.
- Give `InStr` and `StrComp` an explicit compare mode (`vbBinaryCompare` or
  `vbTextCompare`).
- Name the counter after every `Next`.
- One statement per line: no `:` separators. A one-line `If ... Then` is
  allowed for a guard or an `Exit`.
- Every `Select Case` has a `Case Else`, empty with a comment when no other
  value can occur.
- Release object variables with `Set ... = Nothing` on the exit path.

<a id="export-compatibility"></a>

## 💾 Export compatibility

- **Encoding:** the VBE reads and writes source in the Windows ANSI code
  page, which differs between hosts. Source is therefore **ASCII only**; a
  non-ASCII character is built at run time with `ChrW`. No byte-order mark.
  No working-tree encoding is configured in `.gitattributes`.
- **Line endings:** stored with LF in Git and checked out with CRLF by
  `.gitattributes`, whatever `core.autocrlf` says; the VBE expects CRLF.
  Never change these rules or the export header lines when reformatting.
- **Document modules:** `ThisWorkbook` and the sheet modules are exported as
  `.cls` files in `src/workbook/`. They are pasted into the workbook, never
  imported ([INSTALLATION.md](../INSTALLATION.md#importing-vba-into-excel)).
- **Forms:** a `.frm` and its `.frx` are one component, kept side by side in
  `src/forms/`. The `.frx` is binary and never edited as text.
- **Workbooks, Access databases and macro-enabled Office files** are build
  outputs, ignored by Git. One
  may be committed only at an exact path re-included in `.gitignore` and
  approved in an issue. The only one is the macro-free template
  `src/workbook/SFT_Template.xlsx`, approved in
  [#102](https://github.com/danielep71/K-SFTs/issues/102); the
  `vba-source` gate rejects any other workbook and a template that holds a
  VBA project.

## 🔎 Review and validation

- A formatting-only change leaves names, signatures, visibility, constants,
  string literals, calculations, tests and error behavior unchanged; it may
  qualify library calls with `VBA.`, write `vbNullString`, name a `Next`
  counter, add an empty `Case Else` or an explicit binary compare mode, and
  rename local variables. A defect found on the way goes in a separate
  change.
- Comments are kept true. A change updates every module header, procedure
  header, `INPUTS` entry and line comment it makes stale, and the `UPDATED`
  date of each procedure and module it touches.
- `python tools/check.py` must pass. It is a static check and never compiles
  or runs VBA: a static PASS proves neither that the project compiles nor
  that a result is right ([tools/README.md](../tools/README.md#vba-static-checks)).
- Since the first recorded host run, on 2026-10-24 at `1a6f0f7`
  ([#9](https://github.com/danielep71/K-SFTs/issues/9)), changed VBA is
  merged only after it has been imported, compiled and tested in Excel on a
  [supported host](../INSTALLATION.md#supported-hosts), and the run is
  recorded as in
  [recording an Excel run](../INSTALLATION.md#recording-an-excel-run).
  Before that run, VBA was merged on a static PASS with the pull request
  stating that it had not been compiled or run, as for the harness
  ([#76](https://github.com/danielep71/K-SFTs/pull/76)) and its evidence
  entry point ([#93](https://github.com/danielep71/K-SFTs/pull/93)).

---

**Style principle:** state the contract, name the unit, restore only what you
change, and never lose the original error.
