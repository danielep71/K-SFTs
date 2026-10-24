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
> template or workbook exists. The import, compile and harness steps below
> were first run on a host on 2026-10-24, recorded in
> [tested environments](#tested-environments) and
> [#9](https://github.com/danielep71/CCR-K-SFTs/issues/9). The
> [developer setup](#developer-setup) has been run from a fresh clone on
> Linux. How changes are proposed is in [CONTRIBUTING.md](CONTRIBUTING.md).

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
| Microsoft Excel 16.0, Windows 64-bit | 20430 (`Application.Build`) | Windows 11 Enterprise 25H2, OS build 26200.9711 | 64-bit Office, VBA7 | `1a6f0f7` | [#9](https://github.com/danielep71/CCR-K-SFTs/issues/9): import, compile and the three example runs; `check_excel_run.py` PASS, run on that Windows host |

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
for [#9](https://github.com/danielep71/CCR-K-SFTs/issues/9) and for every
later issue or pull request that cites an Excel run. It needs no
self-hosted runner, no change to the macro-security level and no *Trust
access to the VBA project object model*.

1. **Fix the candidate.** In a clone with no local changes
   (`git status --short` prints nothing), type `git switch --detach <commit>`,
   then `git rev-parse HEAD` for the full SHA.
2. **Make the evidence folder**, for example `test-results\excel-run\` in the
   checkout (ignored by Git), with an empty `exported\` folder inside.
3. **Build a development workbook** from that checkout as in
   [building a workbook from source](#building-a-workbook-from-source),
   with the modules of `tests/modules/`.
4. **Compile** with **Debug → Compile VBAProject**.
5. **Run.** In the Immediate window type, with the full path of the folder
   and no trailing backslash:

   ```text
   TEST_Examples.SaveExampleEvidence "C:\path\to\test-results\excel-run"
   ```

   It runs the suite normally, with the injected failure and with the
   skipped case, and saves each complete log as `examples.log`,
   `injected.log` and `skipped.log`. Each run prints a `SAVED=` line: the
   normal run with `error=0`, the two failure runs with the error *Suite
   examples did not pass*. The errors are printed, not raised, so that all
   three logs are written.
6. **Export the modules back.** In the VBE, right-click each module →
   **Export File**, into `exported\`, under its own name
   (`TEST_Examples.bas`, `TEST_Harness.bas`). Do not edit the exported files.
7. **Write `host.txt`** in the folder:

   ```text
   windows=Windows 11 Pro, version 24H2, OS build 26100.4061
   references=Visual Basic For Applications; Microsoft Excel 16.0 Object Library; OLE Automation; Microsoft Office 16.0 Object Library
   compile=PASS
   ```

   `windows` is what `winver` shows. `references` lists the boxes ticked
   in **Tools → References**, separated by `; `. Write `compile=PASS` only if
   step 4 completed with no error; otherwise write what it reported.
8. **Check** the folder against the commit, with Python 3.12 or later and Git
   (run on Linux and, for #9, on Windows 11; from WSL or a Linux checkout if
   it fails on another host):

   ```sh
   python tools/check_excel_run.py test-results/excel-run --commit <commit>
   ```

   It must end with `excel-run: PASS`. It prints the commit, the host and
   the Git blob id of every imported component.
9. **Record.** Copy the report into the issue or pull request and add a row
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
| `host.txt` lacks a field, the references are not exactly the four [supported](#supported-hosts) ones, or `compile` is not `PASS` | The host record is incomplete |

It cannot see Excel: the compile result, the references and the Windows
build are the runner's own record. A PASS says the evidence is consistent
with the commit; it validates no SFT calculation.

## 📚 Related documents

- [docs/VBA_HOUSE_STYLE.md](docs/VBA_HOUSE_STYLE.md): naming, contracts,
  errors, cleanup, encoding and line endings.
- [docs/REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md): layout,
  dependency direction, placement rules.
- [tools/README.md](tools/README.md): the repository checks.
- [SECURITY.md](SECURITY.md): synthetic data only; reporting a problem.
