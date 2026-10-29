# 🧰 Repository tooling

`check.py` is the single entry point for the repository's checks. The same
command runs locally and in the `Repository integrity` workflow. It checks the
repository tooling, the label and issue-form configuration and the workflows.
It also checks the format of the numerical test-case files and, statically,
the exported VBA source. It never compiles or runs VBA in Excel and never
validates an SFT calculation ([VBA static checks](#vba-static-checks)).

## ▶️ Running the checks

Requirements: Git, Python 3.12 or later and Node.js 22 on `PATH`, plus the
quality tools below. `check.py` itself needs only the standard library. The
step-by-step setup is in
[INSTALLATION.md](../INSTALLATION.md#developer-setup).

```sh
python tools/check.py
```

| Mode | Behaviour |
| --- | --- |
| *(default)* | Runs every gate. A gate whose tool is missing is **NOT RUN** and fails the run. |
| `--allow-missing` | Reports missing tools as **NOT RUN** and exits 0 with status **INCOMPLETE**. An incomplete run is never a pass; use it only to iterate locally. |
| `--ci` | Requires every tool and a clean checkout (no modified or untracked files). Used by the workflow. |

Exit code 0 means PASS (or INCOMPLETE with `--allow-missing`); 1 means a gate
failed, a tool was missing, or `--ci` found a dirty tree.

Reports go to the ignored `test-results/` directory:
`repository-integrity.json` and `repository-integrity.md` record the candidate
commit, the mode, whether the tree was dirty, each gate's result and the tool
versions; each gate also writes `<gate>.log`.

## ✅ Gates

| Gate | Checks |
| --- | --- |
| `tool-tests` | Unit tests of the tooling (`tools/test_*.py`), including this runner's status, report and missing-tool rules. |
| `labels-sync-self-test` | The label synchronizer's validation, overlay and reconciliation fixtures. |
| `labels-manifest` | `.github/labels.json` is valid under `.github/label-policy.json`. |
| `labels-drift-self-test` | The read-only drift detector's fixtures. |
| `issue-config` | The issue forms set the assignee and exactly their default priority (bug `P2`, enhancement and documentation `P3`), use only catalogue labels and have a dependencies field; blank issues are disabled; `blocked` is the catalogue's only dependency status label (`check_issues.py --check-config`). |
| `test-cases` | The numerical test-case files under `tests/fixtures/` and `tests/expected/` follow [docs/methodology/TEST_CASES.md](../docs/methodology/TEST_CASES.md): envelopes, identifiers and references, unit suffixes, output forms, rule IDs, the regime and calculation date against the version applicability table, reference classes, register sources and tolerances (`check_test_cases.py`). Format only: no value is computed or compared. |
| `changelog` | `CHANGELOG.md` follows its [date and version rules](../CHANGELOG.md#date-and-version-rules): `## [Unreleased]` first, release headings `## [X.Y.Z] - YYYY-MM-DD` with real dates, newest first by version and date, a link reference for each, a release's pointing to its tag; `VERSION` absent before the first dated release, then equal to the newest one (`check_changelog.py`). Nothing is tagged or published. |
| `vba-source` | VBA components only in their homes, named with their role's prefix, `VB_Name` equal to the file name; ASCII only; `Option Explicit`, and `Option Private Module` in `src/core/` and `tests/modules/`; no `Option Base 1`, tab, trailing white space or unclosed string literal; each form beside its `.frx`; stored with LF and checked out with CRLF; no committed workbook (`check_vba_source.py`). |
| `vba-conditionals` | Every `#If` block evaluates in each modeled host; a project `#Const`, an unknown symbol or an unbalanced block fails; a `Declare` reachable under VBA7 is `PtrSafe` (`check_vba_conditionals.py`). |
| `vba-jumps` | Every `GoTo`, `GoSub` and `Resume` target is a label of the same procedure, labels are unique per procedure and every procedure is closed, in each modeled host (`check_vba_jumps.py`). |
| `vba-manifest` | `tools/vba-projects.json` lists every tracked component of each workbook, once and in sorted order: the document modules of `src/workbook/`, the other `src/` components for production, and those plus `tests/modules/` for development; examples never (`check_vba_manifest.py`). |
| `vba-public-api` | Every procedure, and every module-level declaration that VBA would otherwise make public (`Enum`, `Type`, `Event`, `Declare`), states its visibility; the public declarations of `src/modules/` and [docs/PUBLIC_API.txt](../docs/PUBLIC_API.txt) match one to one with their signatures; no public name collides (`check_vba_public_api.py`). |
| `ruff` | Lint of `tools/` with the rules in `pyproject.toml`. |
| `ruff-complexity-probe` | The complexity ceiling is active: a synthetic complexity-15 function passes and one at 16 is rejected. |
| `mypy` | `mypy --strict` over `tools/`, tests included. |
| `mypy-strict-probe` | Strict mode is active: an annotated function passes and the same function without annotations is rejected. |
| `actionlint` | Syntax, expressions and shell scripts (through shellcheck, when installed) of every workflow. |

The label gates run on every pull request here, although the label workflows
run on a pull request only when it changes label files. Live label
reconciliation and drift detection stay in those workflows.

## 🧪 Prototype scripts

Not gates of `tools/check.py`:

| Script | Does |
| --- | --- |
| `build_template.py` | Writes `src/workbook/SFT_Template.xlsx` from `tests/derivations/prototype/`; needs `xlsxwriter` |
| `lo_regression.py` | Runs the exported VBA headless in LibreOffice (VBA compatibility mode), compares Results with the derivation and runs the `SFT_*` unit tests. Not Excel evidence; its harness-only shims are listed in its docstring |

<a id="vba-static-checks"></a>

## 🔤 VBA static checks

The four `vba-*` gates apply the mechanical rules of
[VBA_HOUSE_STYLE.md](../docs/VBA_HOUSE_STYLE.md) and the placement and API
rules of [REPOSITORY_STRUCTURE.md](../docs/REPOSITORY_STRUCTURE.md) to the
`.bas`, `.cls` and `.frm` files that Git tracks, staged ones included; a new
file is checked once it is added. They share the reading and lexing in
`vba_lexer.py`, and each prints how many components it examined: the five
prototype modules in `src/` and the test modules in `tests/modules/`. A PASS
says only that nothing tracked breaks a rule.

**Three different checks.** A static PASS is not a compilation and not a
validation:

| Check | Where | What it proves |
| --- | --- | --- |
| Static checks | `python tools/check.py`, locally and in CI | The source files follow the rules above |
| Compilation | Excel: **Debug → Compile VBAProject** on a [supported host](../INSTALLATION.md#supported-hosts) | The project compiles: names, calls, types, arguments |
| Numerical validation | The regression harness ([#7](https://github.com/danielep71/K-SFTs/issues/7)) against [published or independent cases](../docs/methodology/TEST_CASES.md#completeness-policy) | A rule computes the documented result |

**Modeled hosts.** `vba-conditionals`, `vba-jumps` and `vba-public-api` read
each component once per compilation environment, with the branches of every
`#If` that environment would compile: `vba7-win64`, the supported 64-bit
Excel, plus `vba7-win32` and the pre-VBA7 `vba6-win32`, so that a branch
written for them is still checked. As in VBA, `Win32` is true in all three
and `VBA6`, which means version-6 compatible, is true under VBA7 too.
The symbols are `VBA6`, `VBA7`, `Win32` and `Win64`; `Mac`, any other symbol
and a project `#Const` fail closed, because their value cannot be known.

**Not checked, left to review and to compiling in Excel:**

- module and procedure banners, contracts, units and `UPDATED` dates;
- names of procedures, variables, constants and enumerations, and the
  component prefixes' subject words;
- `ByVal`, `Variant` use, alignment, indentation depth and line length;
- that `src/core/` never touches the Excel object model and holds no
  module-level state, and that the facade holds no formulas;
- the error-handling and cleanup patterns, including the scope of
  `On Error Resume Next`;
- calls: whether a called procedure exists, is visible and gets the right
  arguments. Compiling in Excel checks direct calls.
- **Dynamic calls**: no check resolves a target named in a string or chosen
  at run time, such as `Application.Run`, `CallByName`, `Application.OnTime`,
  `Application.OnKey`, a control's `OnAction` or `Evaluate`, or one passed
  with `AddressOf`. Compiling does not check them either: a misspelled or
  missing target fails only when the call runs, so each needs a test that
  runs it. `Option Private Module` and `Private` do not stop
  `Application.Run` from reaching a procedure; the
  [public API boundary](../docs/REPOSITORY_STRUCTURE.md#-public-api-boundary)
  is a contract, not an access control.

**Design notes.**

- One shared lexer and one finding and exit format; the valid and invalid
  fixtures are in `test_check_vba.py`, run by `tool-tests`;
- ASCII only, not a Windows code page;
- this repository's homes, role prefixes (`CORE_`, `SFT_`, `C`, `ThisWorkbook`
  or `sh`, `M_`, `frm`, `TEST_`, `Example`), component kinds per home and the
  31-character name limit; `Option Private Module` in `tests/modules/` too;
  `Option Base 1`, tabs and trailing white space rejected; no workbook may be
  committed;
- signatures in the inventory without white space just inside parentheses,
  string literals kept as written;
- `Rem` recognized after a `:` separator and after a numeric line label as
  well as at the line start, and `vba-jumps` reading each statement of a
  line separated by `:`, so that `Private Sub F(): ...: End Sub` closes;
- a comment ending in ` _` read as continuing on the next line, as VBA reads
  it, in all four gates, so that a `GoTo`, a label, a directive or a
  declaration on that line is comment text;
- an apostrophe or a colon inside `[...]`, as in `['Sheet 1'!A1]` or
  `[A1:B2]`, read as part of the expression, and every target of
  `On <expression> GoTo` and `GoSub` checked, not only the first;
- `VBA6` true under VBA7, as Microsoft documents it, and no `Declare`
  keyword test on procedure lines, which would also match text inside string
  literals;
- no financial assertion.

Not checked by a gate: calls between procedures, since a static check cannot
resolve dynamic calls; and white space in files other than VBA source, which
`.editorconfig` covers. `check_excel_run.py` is written for this
repository's harness.

<a id="excel-run-evidence"></a>

## 🧾 Excel run evidence

`check_excel_run.py` checks the evidence folder of an Excel run against the
commit the workbook was built from: the exported modules against the
commit's components, the three harness logs against the commit's example
suite, and the host record, which it writes when the folder has none and
it runs on Windows. It is not a gate, since the evidence comes from
a host run; the procedure is in
[INSTALLATION.md](../INSTALLATION.md#recording-an-excel-run).

```sh
python tools/check_excel_run.py <evidence-folder> --commit <commit>
```

Exit code 0 means PASS, 1 a finding. It needs only the standard library and
Git. It has been run on Linux and, for #9's evidence, on Windows 11 in
PowerShell; the other Python tools are run on Linux only. If it fails on
another host, run it from WSL or a Linux checkout with the evidence folder
copied across. Its positive and negative fixtures, built from the commit's
own source, are in `test_check_excel_run.py`, run by
`tool-tests`. The fixture logs are synthetic: they follow the documented
log format, not a recorded run.

<a id="vba-synchronization"></a>

## 🔄 VBA synchronization

`Sync-KSFT-VBA.ps1` imports the repository's VBA into a workbook built from
`src/workbook/SFT_Template.xlsx`, by Excel's VBA object model, in the mode
and with the components that `vba-projects.json` lists. The `.cmd` files
run it for a workbook dragged onto them. A development synchronization
also prepares an Excel run's evidence folder: it removes an earlier run's
files and exports the workbook's modules. With `-RecordRun`, or
`Record-KSFT-ExcelRun.cmd`, it also runs `SaveExampleEvidence` in the saved
workbook and `check_excel_run.py` against `HEAD`, so one command records the
run. With `-FromTemplate`, or the launcher started without a workbook, it
first builds the workbook from the template. It is optional and needs *Trust
access to the VBA project object model*; the procedure and its safeguards
are in [INSTALLATION.md](../INSTALLATION.md#synchronizing-vba).

It uses this repository's prefixes and homes, skips homes that do not exist
yet, stops without saving when a mode lists nothing, imports forms from the
repository only, and keeps no retired-component list. No test runs it: it
needs Windows Excel. It is
validated by running it on a host: first on 2026-10-24, in development mode
in Windows PowerShell 5.1
([#102](https://github.com/danielep71/K-SFTs/issues/102)).

## 🗂️ Issue triage

`check_issues.py` applies the issue rules of `docs/GOVERNANCE.md` (Triage) to
every open issue and reports findings; it never changes an issue. The
`Issue triage` workflow runs it on issue events with a read-only token.
Pull requests are not audited, but their numbers are loaded so that a
dependency section may name one as a reference; it never counts as a
prerequisite. The token therefore needs `pull-requests: read` as well:
without it the issues API omits pull requests.

```sh
GITHUB_TOKEN=<token with issues:read and pull-requests:read> python tools/check_issues.py --repository danielep71/K-SFTs
python tools/check_issues.py --input issues.json   # issues exported as GitHub REST JSON
python tools/check_issues.py --check-config        # forms and label catalogue only
```

Exit code 0 means no findings, 1 means findings, 2 means the audit could not
run. Its rules are tested by `test_check_issues.py` in `tool-tests`.

## 📈 Traffic history

`traffic_history.py` merges one collection of GitHub traffic API responses
(`views.json`, `clones.json`, `referrers.json`, `paths.json`) into the CSV
history of the `traffic-history` branch. It needs only the standard library
and no token: the `Traffic history` workflow fetches the responses first.

```sh
python tools/traffic_history.py merge --raw <responses> --data <history>/data
python tools/traffic_history.py init --data <history>/data   # empty files with headers
```

Every response and every existing file is validated before anything is
written, so exit code 2 means the history was left unchanged. The merge and
data rules are in [docs/TRAFFIC.md](../docs/TRAFFIC.md); `test_traffic_history.py`
covers them in `tool-tests`.

## 🔧 Tool versions

| Tool | Version | Source |
| --- | --- | --- |
| Python | 3.13 in CI; 3.12 or later locally | `actions/setup-python` |
| Node.js | 22 | `actions/setup-node` |
| Ruff, mypy | 0.17.1, 2.4.0 | `tools/requirements-quality-ci.txt`, hash-locked for CPython 3.13 on Linux x86_64 |
| actionlint | 1.7.12 | Release archive verified by SHA-256 in the workflow |

Install the same quality tools locally, on CPython 3.13 and Linux x86_64:

```sh
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
```

On another platform, install the same versions without the lock
(`pip install ruff==0.17.1 mypy==2.4.0`); CI remains the reference run.
actionlint is available from its
[1.7.12 release](https://github.com/rhysd/actionlint/releases/tag/v1.7.12).

## ☁️ Continuous integration

`.github/workflows/repository-integrity.yml` runs `python tools/check.py --ci`
on every pull request into `main` or `release/**`, on every push to them and
on manual dispatch. It has no path filter, so its `Repository integrity` job
reports on every pull request and can be made a required check.

- The job checks out the exact commit under review without persisting
  credentials and verifies that `HEAD` is that commit. For a pull request this
  is GitHub's test merge of the head into the base.
- The token is read-only (`contents: read`); the job has no write permission,
  secret, approval, merge or alert-dismissal step.
- Every action is pinned to a full commit SHA.
- The Markdown report is added to the job summary, and `test-results/` is
  uploaded as the artifact `repository-integrity-<sha>-<run>-<attempt>`,
  **retained for 30 days**. Download it from the run page.
- The workflow uses no other repository at run time; its only downloads are
  the pinned actions, the locked Python wheels and the actionlint archive.

## 📦 Dependency updates

`.github/dependabot.yml` proposes weekly version updates each Monday at 05:00
Europe/Rome for pinned GitHub Actions and the Python requirements in `/tools`.
Updates are grouped separately as `github-actions` and `python-tools` against
the active release branch, with the
`ci` and `P3` labels and danielep71 as assignee, with no PR milestone.
During triage, link each PR to an appropriate issue carrying the milestone.
Dependabot reads this file from the default branch, `main`, where it is
active ([docs/GOVERNANCE.md](../docs/GOVERNANCE.md#configured-controls)). A
change to it takes effect once it is integrated into `main`. When a new release
branch is opened, update its `target-branch`. Milestones remain assigned only
to issues.

Nothing merges automatically. Before merging an update:

1. Read the release notes between the old and new versions of each action or
   Python package.
2. For Actions, check that each new pin is a full SHA that the upstream tag
   points to. For Python, review package provenance and compare every wheel
   SHA-256 with PyPI; retain the CPython 3.13 / Linux x86_64 lock contract.
3. Wait for `Repository integrity` to pass on the pull request; actionlint
   checks updated workflows, while Python updates must pass the existing
   hash-required binary installation and quality gates.
4. Review the final diff and head SHA yourself: the code-review bot does not
   review bot-authored pull requests (`docs/GOVERNANCE.md`, Review).

Python version updates to `tools/requirements-quality-ci.txt` use a 14-day
`cooldown.default-days` for all packages. Eligible versions are considered at
the next weekly check; this is not a promise of a PR exactly on day 14.
Cooldown applies to version updates, not Dependabot security updates. See the
[GitHub options reference](https://docs.github.com/en/code-security/reference/supply-chain-security/dependabot-options-reference#cooldown).

Dependabot proposals do not replace hash and platform review. If a proposal
cannot preserve the lock contract, regenerate the affected lock entries by
hand and verify the CPython 3.13 / Linux x86_64 wheel hashes against PyPI.
Dependabot's pip updater lists the hash of every published wheel: keep only
the CPython 3.13 / Linux x86_64 one for each package, as in
[#196](https://github.com/danielep71/K-SFTs/pull/196).
The actionlint checksum still requires manual updates: choose a release at
least two weeks old and compare the archive with its release checksums file.
