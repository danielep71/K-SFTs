# 🧰 Repository tooling

`check.py` is the single entry point for the repository's checks. The same
command runs locally and in the `Repository integrity` workflow. It checks the
repository tooling, the label and issue-form configuration and the workflows.
It does not check VBA source, compile anything in Excel or validate SFT
calculations; no such gates exist until the code they check exists.

## ▶️ Running the checks

Requirements: Git, Python 3.12 or later and Node.js 22 on `PATH`, plus the
quality tools below. `check.py` itself needs only the standard library.

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
| `ruff` | Lint of `tools/` with the rules in `pyproject.toml`. |
| `ruff-complexity-probe` | The complexity ceiling is active: a synthetic complexity-15 function passes and one at 16 is rejected. |
| `mypy` | `mypy --strict` over `tools/`, tests included. |
| `mypy-strict-probe` | Strict mode is active: an annotated function passes and the same function without annotations is rejected. |
| `actionlint` | Syntax, expressions and shell scripts (through shellcheck, when installed) of every workflow. |

The label gates run on every pull request here, although the label workflows
themselves only run when label files change. Live label reconciliation and
drift detection stay in those workflows.

## 🗂️ Issue triage

`check_issues.py` applies the issue rules of `docs/GOVERNANCE.md` (Triage) to
every open issue and reports findings; it never changes an issue. The
`Issue triage` workflow runs it on issue events with a read-only token.
Pull requests are not audited, but their numbers are loaded so that a
dependency section may name one as a reference; it never counts as a
prerequisite. The token therefore needs `pull-requests: read` as well:
without it the issues API omits pull requests.

```sh
GITHUB_TOKEN=<token with issues:read and pull-requests:read> python tools/check_issues.py --repository danielep71/CCR-K-SFTs
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
| Ruff, mypy | 0.16.6, 2.3.1 | `tools/requirements-quality-ci.txt`, hash-locked for CPython 3.13 on Linux x86_64 |
| actionlint | 1.7.12 | Release archive verified by SHA-256 in the workflow |

Install the same quality tools locally, on CPython 3.13 and Linux x86_64:

```sh
python -m pip install --only-binary=:all: --require-hashes -r tools/requirements-quality-ci.txt
```

On another platform, install the same versions without the lock
(`pip install ruff==0.16.6 mypy==2.3.1`); CI remains the reference run.
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

`.github/dependabot.yml` proposes weekly updates to the pinned GitHub Actions,
grouped into one pull request against the active release branch, with the
`ci` and `P3` labels and danielep71 as assignee, with no PR milestone.
During triage, link each PR to an appropriate issue carrying the milestone.
Dependabot reads this file from the default branch, so it takes effect once
it is integrated into `main`. When a new release branch is opened, update its
`target-branch`. Milestones remain assigned only to issues.

Nothing merges automatically. Before merging an update:

1. Read the release notes between the old and new versions of each action.
2. Check that each new pin is a full SHA that the upstream tag points to.
3. Wait for `Repository integrity` to pass on the pull request; actionlint
   checks the updated workflows.
4. Review the final diff and head SHA yourself: the code-review bot does not
   review bot-authored pull requests (`docs/GOVERNANCE.md`, Review).

Dependabot does not update `tools/requirements-quality-ci.txt` or the
actionlint checksum. Update them by hand: choose releases at least two weeks
old, download the wheels for CPython 3.13 on Linux x86_64, compare every
SHA-256 with PyPI, and for actionlint with the release's checksums file.
