# ✍️ VBA house style

This document is authoritative for **how CCR-K-SFTs's VBA is written**:
naming, module rules, procedure contracts, errors, Excel-state handling and
source presentation. Where a component lives is owned by
[REPOSITORY_STRUCTURE.md](REPOSITORY_STRUCTURE.md); the supported hosts and
the import and export steps by [INSTALLATION.md](../INSTALLATION.md). Its
structure is adapted from K-SACCR's `docs/VBA_HOUSE_STYLE.md` (see
[NOTICE.md](../NOTICE.md)); none of K-SACCR's components, formulas or
parameters apply here.

**Status on 2026-10-24:** the only VBA components are the test modules in
`tests/modules/`; no production component exists yet. These rules apply to
every component. The `vba-*` gates of `python tools/check.py` enforce the
mechanical ones: component names and `VB_Name`, module options, ASCII,
tabs and trailing white space, `PtrSafe`, jump targets, explicit visibility
and the public API inventory. Everything else is enforced by review
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
| Procedure | PascalCase verb phrase | `ValidateCollateral` |
| Parameter, local variable | PascalCase noun, spelled out | `MarketValue` |
| Loop counter | one lower-case letter | `i`, `j`, `r`, `c` |
| Module-level variable | `m` + PascalCase | `mRunState` |
| Public variable | `g` + PascalCase | `gLastError` |
| Constant | UPPER_SNAKE_CASE | `MAX_NETTING_SETS` |
| Error number constant | `ERR_` + UPPER_SNAKE_CASE | `ERR_INVALID_INPUT` |
| Enum type / member | `Sft` + noun / short prefix + PascalCase | `SftProductType` / `ptRepo` |

The examples show the form of a name only; none of these components exists,
and no product, quantity or parameter is decided by them.

- Names are English and spelled out. Abbreviations are limited to
  established terms of the documented method, written in full in the
  procedure's comment the first time they appear.
- A single letter is reserved for a loop counter: the variable a `For` loop,
  or a `Do` loop stepping through positions, counts with. Every other
  variable and parameter names what it holds, also in a short helper
  (`Formatted`, not `s`).
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

Every procedure documents its contract in a banner immediately after its
signature. Omit sections that are empty.

| Section | Explain |
| --- | --- |
| `PURPOSE` | The responsibility the procedure owns |
| `INPUTS` | Meaning, **unit**, **valid domain** and ownership of each argument |
| `RETURNS` | Meaning and unit of the result and of any output parameter |
| `ERROR POLICY` | Errors raised, propagated or contained, and cleanup |
| `DEPENDENCIES` | Boundaries that matter to callers or maintainers |
| `STATE OWNERSHIP` | Excel or module state read, changed or restored |
| `REFERENCE` | The methodology entry, by source ID and locator, that the procedure implements |
| `UPDATED` | Date the procedure was last revised |

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

- **Error numbers** are named `ERR_` constants in the range
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
   failed, the procedure raises `ERR_CLEANUP_FAILED`, never success.

The reference implementation is `RunExampleOperation` in
[tests/modules/TEST_Examples.bas](../tests/modules/TEST_Examples.bas), with
cases for the success path, a primary error with failed cleanup, and a
cleanup-only failure. The sketch below shows the pattern.

```vb
'In a standard module named M_Example
Private Const ERR_CLEANUP_FAILED As Long = vbObjectError + 2049  'Sketch only


Public Sub RunOperation(ByRef CleanupDetails As String)
    Dim SavedCalculation   As XlCalculation    'Calculation mode before the operation
    Dim SavedScreen        As Boolean          'ScreenUpdating before the operation
    Dim StateCaptured      As Boolean          'True once both were read
    Dim ErrNumber          As Long             'Primary error number; 0 for none
    Dim ErrSource          As String           'Primary error source
    Dim ErrDescription     As String           'Primary error description

        CleanupDetails = vbNullString
        On Error GoTo Failed
        SavedCalculation = Application.Calculation
        SavedScreen = Application.ScreenUpdating
        StateCaptured = True
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        '... work ...

CleanUp:
        On Error GoTo 0
        If StateCaptured Then
            CleanupDetails = AppendDetail(RestoreCalculation(SavedCalculation), _
                                          RestoreScreenUpdating(SavedScreen))
        End If
        If ErrNumber <> 0 Then Err.Raise ErrNumber, ErrSource, ErrDescription
        If Len(CleanupDetails) > 0 Then
            Err.Raise ERR_CLEANUP_FAILED, "M_Example.RunOperation", CleanupDetails
        End If
        Exit Sub

Failed:
        ErrNumber = Err.Number
        ErrSource = Err.Source
        ErrDescription = Err.Description
        Resume CleanUp
End Sub


Private Function RestoreCalculation( _
    ByVal SavedCalculation As XlCalculation) _
    As String

        On Error GoTo Failed
        Application.Calculation = SavedCalculation
        Exit Function

Failed:
        RestoreCalculation = "Calculation not restored: " & Err.Description
End Function
```

`RestoreScreenUpdating` follows `RestoreCalculation`, and `AppendDetail` joins
two failure texts with `"; "`, skipping empty ones.

Core procedures never change Excel state, so they need none of this.

## 🧱 Module layout

1. `Attribute VB_Name` on the first line of a `.bas` export.
2. A full-width module banner: name, purpose, public surface, dependencies,
   state ownership, error policy, compatibility, updated date and author.
3. `MODULE SETTINGS`: `Option Explicit`, plus `Option Private Module` where
   required.
4. `MODULE CONSTANTS`, and `MODULE STATE` when needed (never in
   `src/core/`).
5. Centered, uppercase procedure-group headings, then the procedures.

Banners use an apostrophe and 78 `=`; section rules an apostrophe and 78 `-`.
Source decorations are ASCII only.

A generated module is exempt from the banner rules; its header says it is
generated, names the generator and says it must not be edited by hand.

## 📐 Indentation and wrapping

- Four-space levels; no tabs, no trailing whitespace.
- Local declarations at four spaces with `As` aligned in each group; body
  statements at eight spaces, nested blocks adding four.
- Procedure declarations, `End Sub` / `End Function`, labels and `#If`
  directives at column one.
- A signature with parameters puts one parameter per continuation line, and
  a function's return type on its own continuation line:

  ```vb
  Public Function ValidateCollateral( _
      ByVal MarketValue As Double, _
      ByVal Haircut As Double) _
      As Boolean
  ```

- Two empty lines between procedures. Aim for 79-column comments and about
  100 columns of code.
- Comments explain intent and invariants, not assignments. Each declaration
  carries a short trailing comment stating its role or unit.

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
- **Workbooks and other Office files** are build outputs, ignored by Git. One
  may be committed only at an exact path re-included in `.gitignore` and
  approved in an issue; the only intended one is the macro-free template in
  `src/workbook/`. No such path is approved yet.

## 🔎 Review and validation

- A formatting-only change leaves names, signatures, visibility, constants,
  calculations, tests and error behavior unchanged; a defect found on the way
  goes in a separate change.
- Comments are kept true. A change updates every module header, procedure
  header, `INPUTS` entry and line comment it makes stale, and the `UPDATED`
  date of each procedure and module it touches.
- `python tools/check.py` must pass. It is a static check and never compiles
  or runs VBA: a static PASS proves neither that the project compiles nor
  that a result is right ([tools/README.md](../tools/README.md#vba-static-checks)).
- Changed VBA is merged only after it has been imported, compiled and tested
  in Excel on a [supported host](../INSTALLATION.md#supported-hosts). How
  that run is recorded is defined in
  [#9](https://github.com/danielep71/CCR-K-SFTs/issues/9).

---

**Style principle:** state the contract, name the unit, restore only what you
change, and never lose the original error.
