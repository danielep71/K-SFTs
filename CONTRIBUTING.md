# 🤝 Contributing to CCR-K-SFTs

CCR-K-SFTs is an independent personal project with one maintainer, who
decides scope, priorities and what is integrated
([docs/GOVERNANCE.md](docs/GOVERNANCE.md)). The repository is private.
Every change starts from an issue and arrives through a pull request.

The repository is in **setup**: there is no calculation engine, workbook or
validated method yet ([README.md](README.md)). Contributions today are to
the setup itself: tooling, tests, documentation and governance.

Never disclose a suspected vulnerability in an issue or pull request; follow
[SECURITY.md](SECURITY.md).

<a id="reporting-a-problem"></a>

## 🐛 Reporting a problem

Open an issue with one of the forms; blank issues are disabled.

| Form | Use it for |
| --- | --- |
| **Bug report** | Something that behaves differently from its documentation. Give the full commit SHA, the environment (for Excel: build, bitness, Windows version), the steps, and the expected and observed behavior. |
| **Enhancement** | New or changed behavior. Define the observable contract, the non-goals and how it will be validated; a calculation change cites the regulatory provision it implements. |
| **Documentation** | A wrong, stale or missing statement. Name the document that is authoritative for it. |

**Synthetic data only.** Never attach or paste real trades, portfolios,
counterparty, collateral or client data, and never credentials, tokens or
internal paths. Reproduce the problem with invented values; a minimal
example is better than a real file. A workbook can hold data outside its
visible cells, so attach none: describe the steps or the exported source
instead. If confidential data has been posted, say so to the maintainer at
once rather than editing it out quietly; the [data rules](SECURITY.md#-data-and-secrets)
apply.

Each issue gets an assignee, exactly one priority and a milestone at triage
([docs/GOVERNANCE.md](docs/GOVERNANCE.md#-issues)).

<a id="development-workflow"></a>

## 🌿 Development workflow

1. **Set up** once, as in [developer setup](INSTALLATION.md#developer-setup).
2. **Branch** from the active release branch named in
   [docs/GOVERNANCE.md](docs/GOVERNANCE.md#-branches), never from `main` or
   GitHub's default branch by assumption. Name the branch
   `fix/<issue>-<slug>`, `docs/<slug>` or `test/<slug>`.
3. **Change** as little as the issue needs: no unrelated formatting,
   generated output or opportunistic refactoring. Update the documentation the
   change affects in the same pull request, and add a material change to the
   `[Unreleased]` section of [CHANGELOG.md](CHANGELOG.md).
4. **Check** with `python tools/check.py` before every push; every gate must
   pass. After committing, `python tools/check.py --ci` runs exactly what CI
   runs.
5. **Open a pull request** against the active release branch with the
   [template](.github/PULL_REQUEST_TEMPLATE.md). Link the issue with
   `Resolves #<issue>`; give the pull request no milestone.
6. **Review.** Every finding is fixed or answered, and the review covers the
   final head commit ([docs/GOVERNANCE.md](docs/GOVERNANCE.md#-review)).
7. **Merge.** The pull request is squash-merged once `Repository integrity`
   passes on its current head and the branch is up to date. Integration into
   `main` is a separate step, taken by the maintainer.

Never rebase, amend or force-push a commit that has been pushed; correct it
with a new commit.

### Commit messages

A subject line in the conventional form (`docs: ...`, `fix: ...`,
`test: ...`, `chore: ...`), then a body that states:

- **what** changed and **why**;
- **what was validated** at that commit, with the commands run and their
  results; and
- **what was not validated**, separating static checks from Excel runs.

Reference the issue. Do not put credentials, private links or claims that
were not checked in a commit message.

<a id="source-and-validation"></a>

## 🧪 Source and validation

- **Exported VBA is the source.** The `.bas`, `.cls` and `.frm` files are the
  reviewable record; a workbook never is, and none is committed. Import and
  export as in [INSTALLATION.md](INSTALLATION.md#importing-vba-into-excel);
  write as in [docs/VBA_HOUSE_STYLE.md](docs/VBA_HOUSE_STYLE.md); place
  components as in [docs/REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md).
- **A static PASS is not Excel evidence.** `tools/check.py` and CI read files;
  they do not compile VBA or run Excel
  ([three different checks](tools/README.md#vba-static-checks)). A pull
  request that changes VBA states whether it was compiled and run in Excel,
  on which host and at which commit, and never implies it was when it was
  not. Since the first host run on 2026-10-24
  ([#9](https://github.com/danielep71/CCR-K-SFTs/issues/9)), changed VBA
  merges only after an Excel run
  ([review and validation](docs/VBA_HOUSE_STYLE.md#-review-and-validation)).
- **A skipped or incomplete run is not a pass**, whether the checks end
  INCOMPLETE or the harness reports `completeness=INCOMPLETE`.
- **Expected values never come from the code under test.** They come from a
  published example, a registered external implementation, or a separate
  derivation, as the [reference classes](docs/methodology/TEST_CASES.md#reference-classes)
  define.

<a id="where-rules-live"></a>

## 📚 Where the rules live

| Need | Document |
| --- | --- |
| Scope, layout, dependency direction, public API | [docs/REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md) |
| Regulatory basis and test cases | [docs/methodology/README.md](docs/methodology/README.md), [docs/methodology/TEST_CASES.md](docs/methodology/TEST_CASES.md) |
| VBA naming, contracts, errors, encoding | [docs/VBA_HOUSE_STYLE.md](docs/VBA_HOUSE_STYLE.md) |
| Setup, Excel hosts, import and export, the harness | [INSTALLATION.md](INSTALLATION.md) |
| Test layout and the harness log | [tests/README.md](tests/README.md) |
| Checks and CI | [tools/README.md](tools/README.md) |
| Branches, review, protection, issues | [docs/GOVERNANCE.md](docs/GOVERNANCE.md) |
| Pull request checklist | [.github/PULL_REQUEST_TEMPLATE.md](.github/PULL_REQUEST_TEMPLATE.md) |
| Vulnerabilities and data | [SECURITY.md](SECURITY.md) |
| Versions, changelog and releases | [CHANGELOG.md](CHANGELOG.md), [RELEASING.md](RELEASING.md); nothing is published yet. |

## 📄 Licensing

The repository is covered by the [Mozilla Public License 2.0](LICENSE)
([NOTICE.md](NOTICE.md)). Contributions are accepted under the same
licence. Submit only what you have the right to submit: code, tests, data,
text and generated material alike. A file adapted from another repository is
recorded in the [provenance register](NOTICE.md#-provenance-register) in the
same change.
