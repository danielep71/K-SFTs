# 📦 Installation and developer setup

This document is authoritative for **the supported Excel hosts and how the
VBA source is imported into and exported from Excel**. How the source is
written is in [docs/VBA_HOUSE_STYLE.md](docs/VBA_HOUSE_STYLE.md); where each
component lives in [docs/REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md).

> [!IMPORTANT]
> **Status on 2026-10-25: working prototype.** The production components
> are the five prototype modules in `src/`, listed in
> `tools/vba-projects.json`, and the workbook template holds the prototype's
> 11 sheets with a demo portfolio
> ([PROTOTYPE.md](docs/methodology/PROTOTYPE.md)); the prototype is not the
> methodology of record. The import, compile and harness steps below were
> first run on a host on 2026-10-24
> ([#9](https://github.com/danielep71/K-SFTs/issues/9)), and with the
> prototype on 2026-10-25
> ([#135](https://github.com/danielep71/K-SFTs/pull/135)), both recorded in
> [tested environments](#tested-environments). The
> [developer setup](#developer-setup) has been run from a fresh clone on
> Linux. How changes are proposed is in [CONTRIBUTING.md](CONTRIBUTING.md).

<a id="supported-hosts"></a>

## 🖥️ Supported hosts

Decided by the maintainer on 2026-10-24 in
[#4](https://github.com/danielep71/K-SFTs/issues/4).

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
| Microsoft Excel 16.0, Windows 64-bit | 20430 (`Application.Build`) | Windows 11 Enterprise 25H2, OS build 26200.9711 | 64-bit Office, VBA7 | `1a6f0f7` | [#9](https://github.com/danielep71/K-SFTs/issues/9): import, compile and the three example runs; `check_excel_run.py` PASS, run on that Windows host |
| Microsoft Excel 16.0, Windows 64-bit | 20430 (`Application.Build`) | Windows 11 Enterprise 25H2, OS build 26200.9711 | 64-bit Office, VBA7 | `fbaee34` | [#135](https://github.com/danielep71/K-SFTs/pull/135): the prototype built with `Sync-KSFT-VBA.ps1` (Development), compiled and run on the demo portfolio (TestCatalogue 107 of 107, FormulaTest 25 of 25), and the three example runs; `check_excel_run.py` PASS, run on that Windows host |

<a id="get-the-source"></a>

## 📥 Get the source

Use a **Git clone**:

```sh
git clone https://github.com/danielep71/K-SFTs.git
cd K-SFTs
git switch release/0.2.0
```

`release/0.2.0` is the
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

<a id="developer-setup"></a>

## 🛠️ Developer setup

The repository checks run without Excel. They have been run on Linux, in CI
and from a fresh clone; on Windows they are untested, so use WSL or CI there.

| Prerequisite | Version | Needed for |
| --- | --- | --- |
| Git | any recent | the checkout; the checks read the tracked files through `git` |
| Python | 3.12 or later (CI uses 3.13) | `tools/check.py` and every Python gate |
| Node.js | 22 | the label synchronizer and drift-detector self-tests |
| Ruff, mypy | 0.16.6, 2.3.1 | the `ruff` and `mypy` gates |
| actionlint | 1.7.12 | the `actionlint` gate |
| shellcheck | optional | the shell scripts inside workflows, through actionlint; CI's runner has it |
| LibreOffice, its Python UNO bridge, openpyxl | optional; run with 24.2.7.2 and openpyxl 3.1.5 | `tools/lo_regression.py`, the prototype's headless regression; not a gate and not Excel evidence |
| xlsxwriter | optional; run with 3.2.9 | `tools/build_template.py`, which regenerates `src/workbook/SFT_Template.xlsx` |
| Excel | a [supported host](#supported-hosts) | importing, compiling and running the VBA only |

From the root of the checkout, on CPython 3.13 and Linux x86_64:

```sh
python3.13 -m venv .venv
. .venv/bin/activate
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
V=1.7.12
T=$(mktemp -d)
curl -fsSL -o "$T/actionlint.tar.gz" \
  "https://github.com/rhysd/actionlint/releases/download/v$V/actionlint_${V}_linux_amd64.tar.gz"
echo "8aca8db96f1b94770f1b0d72b6dddcb1ebb8123cb3712530b08cc387b349a3d8  $T/actionlint.tar.gz" \
  | sha256sum --check --strict
tar -xzf "$T/actionlint.tar.gz" -C .venv/bin actionlint
python tools/check.py
```

- `.venv/` is ignored by Git, so the tools installed in it leave the checkout
  clean. Install nothing else inside the checkout: `--ci` fails on any
  untracked file.
- The lock in `tools/requirements-quality-ci.txt` covers CPython 3.13 on
  Linux x86_64 only. Elsewhere, install the same versions without it
  (`pip install ruff==0.16.6 mypy==2.3.1`); CI remains the reference run.
- The checksum is the one the workflow verifies
  (`.github/workflows/repository-integrity.yml`).

| Command | Use |
| --- | --- |
| `python tools/check.py` | Before every push. Every gate must PASS; a missing tool is NOT RUN and fails the run. |
| `python tools/check.py --allow-missing` | While iterating without every tool. Ends INCOMPLETE, which is never a pass. |
| `python tools/check.py --ci` | What CI runs: also fails on a modified or untracked file, so run it after committing. |

The gates and modes are described in [tools/README.md](tools/README.md).

<a id="evidence-locations"></a>

### Where evidence goes

| Evidence | Location |
| --- | --- |
| Local check reports | `test-results/repository-integrity.md` and `.json`, and one `<gate>.log` per gate. The directory is ignored by Git. |
| Hosted check results | The `Repository integrity` run of the pull request: its job summary, and the artifact `repository-integrity-<sha>-<run>-<attempt>`, kept 30 days. |
| Harness log | The Immediate window; the complete log through `TEST_Harness.SaveLog`, saved outside the checkout or in `test-results/` ([running the harness](#running-the-harness)). |
| Evidence cited by an issue or pull request | Copied into its body: the run link, the SHA checked and the decisive lines, since logs and artifacts expire ([docs/GOVERNANCE.md](docs/GOVERNANCE.md#accepted-limitations)). Excel run evidence: [recording an Excel run](#recording-an-excel-run). |

<a id="importing-vba-into-excel"></a>

## 📂 Importing VBA into Excel

### Building a workbook from source

1. **Start from one exact commit** of a Git checkout. Never mix components
   from different commits.
2. **Create the workbook.** Open the macro-free template
   `src/workbook/SFT_Template.xlsx` and save it with **File → Save As** as
   **Excel Macro-Enabled Workbook (`.xlsm`)**, outside the checkout or in
   the ignored `test-results/` directory. Built workbooks are never
   committed, and nothing is ever saved over the template. The template
   holds the prototype's 11 sheets, with a demo portfolio, and is generated
   by `tools/build_template.py`; it is regenerated, never edited by hand.
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

<a id="synchronizing-vba"></a>

### Synchronizing VBA into a workbook (optional)

`tools/Sync-KSFT-VBA.ps1` does step 4 (import) and, on request, step 5
(document modules) for you, in a workbook built from the template. It is a
convenience, not a requirement:

- It needs Excel's **Trust access to the VBA project object model**
  (**File → Options → Trust Center → Trust Center Settings → Macro
  Settings**), which the manual steps above do not. Switch it on only as
  your organization's policy allows.
- Either route can [record an Excel run](#recording-an-excel-run).

With the workbook closed, drag the `.xlsm` onto `tools/Sync-KSFT-VBA.cmd`,
or run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-KSFT-VBA.ps1 -WorkbookPath "C:\path\K-SFT.xlsm" -Mode Development
```

| It does | Details |
| --- | --- |
| Imports the components of one workbook | `-Mode Production`: those under `src/`; `-Mode Development`: those and `tests/modules/`. Without `-Mode` it asks. The lists are in `tools/vba-projects.json`, which the `vba-manifest` gate keeps equal to the tracked components. |
| Replaces document-module code, if you agree | It asks, unless `-IncludeSheets` or `-SkipSheets` answers; the workbook's sheet objects are kept. |
| Imports from a folder | `-SourceDir` (or `tools/Sync-KSFT-VBA-Downloads.cmd`, which asks for one and defaults to `Downloads`) takes only the `.bas` and `.cls` files found there, by their `Attribute VB_Name`, after listing them for confirmation. |
| Prepares an Excel run's evidence | A `-Mode Development` synchronization, after saving, removes the files an earlier run left in `test-results\excel-run` (or `-EvidenceDir`): the three logs, `host.txt` and the modules in `exported\`. It then exports every standard, class and form module of the workbook into `exported\`. Production and `-SourceDir` synchronizations leave the folder alone. |
| Records the run | `-RecordRun` (or `tools/Record-KSFT-ExcelRun.cmd`) adds the run to a Development synchronization: it reopens the saved workbook with macros enabled in its own Excel instance, runs `SaveExampleEvidence` on the evidence folder, closes the workbook unsaved and runs `check_excel_run.py` against `HEAD` ([recording an Excel run](#recording-an-excel-run)). Exit code 3: synchronized, but the run did not pass the check. |
| Protects the workbook | A backup copy first, removed after a successful save. It writes no cells, formulas, names or formatting, and before saving it checks that no sheet was added, removed, renamed or reordered. The workbook is closed unsaved on any error. |

It never writes to the repository. After it, compile as in step 6.

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
- **Recovery.** A run that ended without `EndRun` while the VBA project kept
  its state (for example, a new run typed while stopped in break mode), or
  that changed Excel settings it did not restore, leaves `recovery=REQUIRED`,
  and the next run is refused. Check **File → Options → Formulas**
  (calculation) and the other settings the log names, then type
  `TEST_Harness.ResetTests`. **End** or **Reset** in the VBE clears the
  harness's state: the next run is then not refused and takes the settings
  it finds as its baseline, so check them by hand first.
- **What it proves.** The suite exercises the harness itself: assertions,
  tolerances, error values, raised errors, completeness and the cleanup
  pattern. It contains no SFT calculation, and a passing run validates
  none. The log format is in [tests/README.md](tests/README.md#log-format);
  recording a run as evidence is described below.

<a id="recording-an-excel-run"></a>

## 🧾 Recording an Excel run

The evidence that a commit's VBA was imported, compiled and run on a host,
for [#9](https://github.com/danielep71/K-SFTs/issues/9) and for every
later issue or pull request that cites an Excel run. It needs no
self-hosted runner and no change to the macro-security level. Built by
hand, it needs no *Trust access to the VBA project object model*; built with
[the synchronization script](#synchronizing-vba), it needs that setting and
takes fewer steps.

**In one step, with the script.** After step 1 below, drag the development
`.xlsm`, closed, onto `tools/Record-KSFT-ExcelRun.cmd`, or run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Sync-KSFT-VBA.ps1 -WorkbookPath "C:\path\K-SFT.xlsm" -RecordRun
```

It does steps 2 to 5: a Development synchronization, the run of
`SaveExampleEvidence` on `test-results\excel-run` (or `-EvidenceDir`), and
the check against `HEAD`, which writes `host.txt`. It needs Python 3.12 or
later and Git on `PATH`, and ends with `RECORDED` and the check's report,
or `RUN NOT RECORDED` and the reason. Macros are enabled only in the Excel
instance the script starts; the workbook is closed without saving after the
run. Then record as in step 6. Checks that need a person, such as how a
sheet looks, are done by opening the workbook afterwards.

**Step by step:**

1. **Fix the candidate.** In a clone with no local changes
   (`git status --short` prints nothing), type `git switch --detach <commit>`,
   then `git rev-parse HEAD` for the full SHA.
2. **Build a development workbook** from that checkout, with the modules of
   `tests/modules/`, and its evidence folder `test-results\excel-run\`
   (ignored by Git):
   - **With the script:** save the template as `.xlsm` and run a
     [`-Mode Development` synchronization](#synchronizing-vba). It also
     empties the evidence folder of an earlier run and exports the
     workbook's modules into its `exported\` folder.
   - **By hand:** build it as in
     [building a workbook from source](#building-a-workbook-from-source).
     After step 3, export each module into an empty `exported\` folder
     inside the evidence folder, under its own name (VBE: right-click →
     **Export File**), without editing it.
3. **Run.** In the Immediate window type, with the full path of the folder
   and no trailing backslash:

   ```text
   TEST_Examples.SaveExampleEvidence "C:\path\to\test-results\excel-run"
   ```

   It creates the folder when missing (its parent must exist) and removes
   the logs and `host.txt` of an earlier run, leaving `exported\` as it is.
   It then runs the suite normally, with the injected failure and with the
   skipped case, and saves each complete log as `examples.log`,
   `injected.log` and `skipped.log`. Each run prints a `SAVED=` line: the
   normal run with `error=0`, the two failure runs with the error *Suite
   examples did not pass*. The errors are printed, not raised, so that all
   three logs are written. No separate **Debug → Compile VBAProject** is
   needed: a run that reaches the three `SAVED=` lines is recorded as
   compiled.
4. **The host record, `host.txt`,** is written by the check in step 5 when
   it runs on the Excel host's Windows and the folder has none: the Windows
   build read from the machine, the four supported references and
   `compile=PASS`. Write it yourself only when the project reported a
   compile error, other references are ticked, or the check runs off
   Windows:

   ```text
   windows=Windows 11 Pro, version 24H2, OS build 26100.4061
   references=Visual Basic For Applications; Microsoft Excel 16.0 Object Library; OLE Automation; Microsoft Office 16.0 Object Library
   compile=PASS
   ```

   `windows` is what `winver` shows. `references` lists the boxes ticked
   in **Tools → References**, separated by `; `. `compile` is `PASS` only if
   no compile error was reported; otherwise it is the error.
5. **Check** the folder against the commit, with Python 3.12 or later and Git
   (run on Linux and, for #9, on Windows 11; from WSL or a Linux checkout if
   it fails on another host):

   ```sh
   python tools/check_excel_run.py test-results/excel-run --commit <commit>
   ```

   It must end with `excel-run: PASS`. It prints the commit, the host and
   the Git blob id of every imported component.
6. **Record.** Copy the report into the issue or pull request and add a row
   to [tested environments](#tested-environments). The logs stay outside Git.

`check_excel_run.py` reads only the folder and the commit:

| It fails when | Because |
| --- | --- |
| An exported module differs from the commit's file, or one is missing or extra | The workbook was not built from that commit |
| A log is missing or empty, does not end with `RESULT`, holds more than one run, or a line is not a `KEY=value` record | The log is incomplete |
| A record is missing, extra or out of the order in which the commit's suite writes it, or an unexpected error (`ERROR`) is logged | The run did not execute the suite as written |
| The normal run is not `RESULT=PASS`, or its cases and assertions differ from the commit's suite in name or order | The run did not pass as written |
| `CLEANUP` is not `PASS` or `RECOVERY` is not `NONE` in any of the three logs | Cleanup failed |
| The two failure runs do not end as [documented](#running-the-harness), or fail anything but their own failure: the injected run its one designated assertion, the skipped-case run only `harness.completeness` | The failure paths did not behave |
| `ENVIRONMENT` lacks a field, differs between the logs, or Office is not 64-bit | The host is unknown or not the supported target |
| `host.txt` is missing off Windows, lacks a field, the references are not exactly the four [supported](#supported-hosts) ones, or `compile` is not `PASS` | The host record is incomplete |

It cannot see Excel: the compile result, the references and the Windows
build are the runner's own record, written by hand or, on Windows, by
running the check. A PASS says the evidence is consistent
with the commit; it validates no SFT calculation.

## 📚 Related documents

- [docs/VBA_HOUSE_STYLE.md](docs/VBA_HOUSE_STYLE.md): naming, contracts,
  errors, cleanup, encoding and line endings.
- [docs/REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md): layout,
  dependency direction, placement rules.
- [tools/README.md](tools/README.md): the repository checks.
- [SECURITY.md](SECURITY.md): synthetic data only; reporting a problem.
