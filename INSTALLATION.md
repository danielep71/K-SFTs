# 📦 Installation and developer setup

This document is authoritative for **the supported Excel hosts and how the
VBA source is imported into and exported from Excel**. How the source is
written is in [docs/VBA_HOUSE_STYLE.md](docs/VBA_HOUSE_STYLE.md); where each
component lives in [docs/REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md).
Its host and import sections are adapted from K-SACCR's `INSTALLATION.md`
(see [NOTICE.md](NOTICE.md)).

> [!IMPORTANT]
> **Status on 2026-10-24: repository setup only.** The only VBA components
> are the test modules in `tests/modules/`; no production component,
> template or workbook exists. The steps below have never been run on a
> host. The first real import, compile and run is recorded in
> [#9](https://github.com/danielep71/CCR-K-SFTs/issues/9). Prerequisites,
> local commands and contributor setup are completed in
> [#10](https://github.com/danielep71/CCR-K-SFTs/issues/10).

<a id="supported-hosts"></a>

## 🖥️ Supported hosts

Decided by the maintainer on 2026-10-24 in
[#4](https://github.com/danielep71/CCR-K-SFTs/issues/4).

| Host | Support level |
| --- | --- |
| Microsoft 365 or Excel 2016 and later, **Windows, 64-bit** | Supported target |
| Same versions, Windows, 32-bit | Best effort: kept compiling, not certified |
| Excel for Mac, Excel for the web, Excel 2013 and earlier | Not supported |

**References:** only the four defaults of a new workbook: *Visual Basic For
Applications*, *Microsoft Excel 16.0 Object Library*, *OLE Automation* and
*Microsoft Office 16.0 Object Library*. Adding any other reference needs an
issue and an update to this section.

**Intended support is not a test result.** The table is a commitment. Hosts
on which the source has actually been imported, compiled and run are listed
below, each with its evidence.

<a id="tested-environments"></a>

### Tested environments

| Host | Excel build | Windows | Bitness | Commit | Evidence |
| --- | --- | --- | --- | --- | --- |
| None yet | | | | | [#9](https://github.com/danielep71/CCR-K-SFTs/issues/9) |

<a id="get-the-source"></a>

## 📥 Get the source

Use a **Git clone**:

```sh
git clone https://github.com/danielep71/CCR-K-SFTs.git
cd CCR-K-SFTs
git switch release/0.1.0
```

`release/0.1.0` is the
[active release branch](docs/GOVERNANCE.md#-branches); `main` receives it
through an approved integration.

- A clone checks `.bas`, `.cls` and `.frm` files out with the CRLF line
  endings the VBE expects. `.gitattributes` sets them whatever
  `core.autocrlf` says.
- Do not use **Code → Download ZIP**: `.gitattributes` excludes repository
  plumbing such as `.gitattributes` and `.editorconfig` from archives, so a
  ZIP is not a complete checkout.
- Do not download single modules through GitHub's **Raw** view: it serves the
  LF form stored in Git, which the VBE does not read correctly.

<a id="importing-vba-into-excel"></a>

## 📂 Importing VBA into Excel

### Building a workbook from source

1. **Start from one exact commit** of a Git checkout. Never mix components
   from different commits.
2. **Create the workbook.** Once the macro-free template exists in
   `src/workbook/`, open it and save it with **File → Save As** as **Excel
   Macro-Enabled Workbook (`.xlsm`)**. Until then, start from a new blank
   workbook. Save it outside the checkout or in the ignored `test-results/`
   directory; built workbooks are never committed.
3. **Check the references.** Open the VBE (**Alt+F11**) and check that
   **Tools → References** shows only the four listed under
   [supported hosts](#supported-hosts).
4. **Import** with **File → Import File**, in this order, alphabetically by
   file name within each folder:
   1. every file in `src/core/`;
   2. every file in `src/classes/`;
   3. every file in `src/modules/`;
   4. every standard module (`.bas`) in `src/workbook/`; its `.cls` document
      modules are pasted, not imported (step 5);
   5. every `.frm` in `src/forms/`; its `.frx` loads with it, and a `.frx` is
      never imported on its own;
   6. for a **development** workbook only, every module in `tests/modules/`.

   A **production** workbook never contains test or example modules. Folders
   that do not exist yet are skipped.
5. **Paste the document modules.** `ThisWorkbook.cls` and the sheet modules in
   `src/workbook/` cannot be imported: the VBE would create a new class such
   as `ThisWorkbook1`. Open the `.cls` file in a text editor, copy the code
   below its `Attribute` lines, and paste it into the existing `ThisWorkbook`
   or sheet module. Each sheet's code name must equal its file name.
6. **Compile** with **Debug → Compile VBAProject**. It must complete with no
   error.
7. **Save**, and close and reopen when a clean session is needed.

Import by hand needs no change to Excel's security settings: neither *Trust
access to the VBA project object model* nor a lower macro-security level.
Enable macros in the built workbook only as your organization's policy
allows.

Do not paste source into arbitrarily named modules: the component name
(`VB_Name`) is part of the contract and must equal the file name.

<a id="exporting-changes-back"></a>

### Exporting changes back

1. In the VBE, right-click each changed component → **Export File**, and save
   it over its file in the checkout, in the same folder. For a document
   module, copy its code back into the matching `src/workbook/` file below
   the existing `Attribute` lines instead.
2. Review `git diff`. The VBE silently changes the capitalization of an
   identifier everywhere it appears when one declaration changes case; revert
   case-only changes you did not intend. Check that no non-ASCII character
   was introduced ([export compatibility](docs/VBA_HOUSE_STYLE.md#export-compatibility)).
3. Run `python tools/check.py`. Its `vba-*` gates check the exported files
   statically (names, options, ASCII, white space, jumps, visibility and the
   public API inventory); they do not replace compiling in Excel
   ([tools/README.md](tools/README.md#vba-static-checks)).
4. Commit. Git stores the files with LF; nothing is converted by hand.

Exported files keep the VBE's header lines (`VERSION`, `BEGIN … END`,
`Attribute`). Never edit or remove them when reformatting.

<a id="running-the-harness"></a>

## ▶️ Running the harness

In a **development** workbook built as above, open the Immediate window
(**Ctrl+G**) and type:

```text
TEST_Examples.RunExampleTests
```

It runs the example suite through `TEST_Harness` and writes one record per
line. A passing run ends with:

```text
RESULT=PASS; suite=examples; operation=PASS; completeness=COMPLETE; cases=9; assertions=18; failures=0; cleanup=PASS; recovery=NONE
```

Anything else, or the error *Suite examples did not pass*, is a failure. Two
more entry points show the failure paths and are expected to fail:

| Entry point | Expected end |
| --- | --- |
| `TEST_Examples.RunExampleTestsWithInjectedFailure` | `RESULT=FAIL`, `operation=FAIL`, `completeness=COMPLETE`, `failures=1`, then *Suite examples did not pass* |
| `TEST_Examples.RunExampleTestsWithSkippedCase` | `RESULT=FAIL`, `operation=PASS`, `completeness=INCOMPLETE`, `cases=8`, `assertions=16`, `failures=1`, then *Suite examples did not pass* |

- **Complete log.** The Immediate window keeps only its last 200 lines. To
  keep the whole run, type
  `TEST_Harness.SaveLog "<folder>\examples.log"` after it, with a folder
  outside the checkout or the ignored `test-results/` directory.
- **Recovery.** A run that stopped half way, or that changed Excel settings
  it did not restore, leaves `recovery=REQUIRED`, and the next run is
  refused. Check **File → Options → Formulas** (calculation) and the other
  settings the log names, then type `TEST_Harness.ResetTests`.
- **What it proves.** The suite exercises the harness itself: assertions,
  tolerances, error values, raised errors, completeness and the cleanup
  pattern. It contains no SFT calculation, and a passing run validates
  none. The log format is in [tests/README.md](tests/README.md#log-format);
  recording a run as evidence is defined in
  [#9](https://github.com/danielep71/CCR-K-SFTs/issues/9).

## 📚 Related documents

- [docs/VBA_HOUSE_STYLE.md](docs/VBA_HOUSE_STYLE.md): naming, contracts,
  errors, cleanup, encoding and line endings.
- [docs/REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md): layout,
  dependency direction, placement rules.
- [tools/README.md](tools/README.md): the repository checks.
- [SECURITY.md](SECURITY.md): synthetic data only; reporting a problem.
