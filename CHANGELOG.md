# 📜 Changelog

All notable changes to K-SFTs are recorded here, from the user's point of
view. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and versions follow
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).
It is not a commit log or a substitute for release evidence. **No version has
been released**: everything below is under `[Unreleased]`.

## 🧭 Maintenance policy

- Add a material change under `[Unreleased]` in the same pull request as the
  change it describes. A change with nothing for a user to notice (a typo, a
  refactor with no effect) needs no entry.
- Describe the observable result: what exists, what a user or contributor
  can now do, what changed in a contract, and what to do about it.
- Link the issue or pull request that holds the detail.
- Record only evidence that was produced. A static check is not an Excel run:
  say what was and was not run in Excel.
- Until a version is released, edit an existing entry rather than adding one
  that describes a change to an unreleased state: an entry describes what the
  release contains.
- Entries move into a dated section only when a release is prepared
  ([RELEASING.md](RELEASING.md)). A published section is not rewritten; a
  demonstrable error is corrected with a note.

<a id="date-and-version-rules"></a>

### Date and version rules

`python tools/check.py` (gate `changelog`) checks the headings, dates, order,
link references and `VERSION`. What a version number means (`0.y.z`, patch,
minor, major) is checked in review.

| Rule | Standard |
| --- | --- |
| Version | `MAJOR.MINOR.PATCH`, without a leading `v` in headings |
| Release heading | `## [X.Y.Z] - YYYY-MM-DD`, a real calendar date |
| Ordering | `[Unreleased]` first; releases newest to oldest, by version and date |
| Link references | `[Unreleased]` and every release; a release's points to `…/releases/tag/vX.Y.Z` |
| `0.y.z` | Initial development: no compatibility is promised between versions |
| Patch, minor, major | From `1.0.0`: a compatible correction, a compatible capability, an incompatible change to the public contract |

| Category | Use for |
| --- | --- |
| **Added** | New capabilities, files, checks or tests |
| **Changed** | Changes to existing behavior, contracts, tooling or documentation |
| **Fixed** | Corrected defects |
| **Security** | Safely disclosed security corrections |
| **Validation** | Evidence actually produced |
| **Known limitations** | Deliberate, unresolved boundaries |

Use only the categories an entry needs.

<a id="unreleased"></a>

## [Unreleased]

### Changed

- Dependabot proposes grouped weekly Python quality-tool updates from `/tools`
  after a 14-day cooldown, targeting the active release branch; package
  provenance, platform and hash review remain required (#187).

### Added

- Portable non-default settings register with glyphs, badges and documented
  GitHub default comparisons. Custom rulesets are recorded as additions;
  repository-specific observations remain in a separate evidence snapshot (#162).

No version is planned yet. This holds *v0.1.0 - Repo setup*, the repository
infrastructure, closed without a release (#13), and *v0.2.0 - Prototype
setup*, the SFT prototype as it is (#131). Later milestones lead to v1.0.0.

### Added

- Three SFT visual references in editable SVG and high-resolution PNG:
  netting-set aggregation, EU CRR uses and exposure calculation formulas
  with supervisory haircut tables. Embedded in the root README with
  methodological scope and prototype limitations (#142). The calculation
  reference also closes the workbook's README sheet, linked from its
  methodology heading. The README sheet is read-only, and its sections
  still expand and collapse (#138).

- A responsive K-SFTs project web page with prototype scope, calculation
  flow, framework artwork, setup resources and validation boundaries;
  Website followed by Wiki in the root README navigation (#155).

- A counterparty credit risk framework diagram in the root README, showing
  the calculator K-SFTs belongs beside, in editable SVG and PNG, shared with
  K-SACCR and K-Risk Engines (#151).

- A K-SFTs banner in the root README, with a reusable 1280 x 640 image
  in `docs/assets/` (#113, #117).

- Governance: the branch model, integration and sync pull requests, review
  of the final head, rulesets on `main`, `release/**` and `v*` tags, and the
  issue rules ([docs/GOVERNANCE.md](docs/GOVERNANCE.md); #1, #12).
- The project boundary, layout, dependency direction and public API boundary
  ([docs/REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md); #2).
- Licensing under MPL 2.0, the provenance register and the security policy,
  with private vulnerability reporting on GitHub now that the repository is
  public, and Dependabot security updates, secret scanning, push protection
  and CodeQL code scanning since 2026-10-29
  ([NOTICE.md](NOTICE.md), [SECURITY.md](SECURITY.md),
  [settings evidence](docs/REPOSITORY_SETTINGS_EVIDENCE.md#security-settings);
  #3, #123, #160, #162).
- Source conventions: exported VBA as the source, ASCII only, CRLF checkout,
  the VBA house style, supported hosts and the import and export steps
  ([docs/VBA_HOUSE_STYLE.md](docs/VBA_HOUSE_STYLE.md),
  [INSTALLATION.md](INSTALLATION.md); #4, #20).
- `python tools/check.py`, run locally and by the `Repository integrity`
  workflow on every pull request: tooling tests, label and issue-form
  configuration, the test-case format, static checks of the exported VBA
  source and the workflows ([tools/README.md](tools/README.md); #5, #6).
- A VBA regression harness with a neutral example suite in `tests/modules/`,
  and a procedure and checker for recording an Excel run against its commit.
  On Windows the checker writes the host record itself, and a development
  synchronization exports the modules into the evidence folder. With
  `-RecordRun`, or `tools/Record-KSFT-ExcelRun.cmd`, one command synchronizes,
  runs the suite and checks the run against `HEAD`; double-clicked, the
  launcher first builds a fresh development workbook from the template
  ([INSTALLATION.md](INSTALLATION.md#recording-an-excel-run); #7, #9, #147,
  #149, #168, #189).
- The macro-free workbook template `src/workbook/SFT_Template.xlsx` (a Readme
  sheet only), and an optional script that imports the repository's VBA into
  a workbook built from it, with its component manifest and the
  `vba-manifest` gate
  ([INSTALLATION.md](INSTALLATION.md#synchronizing-vba); #102).
- The regulatory basis (the EU CRR, with the Basel Framework as reference),
  registered copies of the texts relied on, and the test-case file format
  ([docs/methodology/](docs/methodology/README.md); #8).
- Onboarding: README, developer setup and contributing guide (#10, #120).
- Issue and pull request forms and an issue-triage check (#11).
- Repository traffic history kept beyond GitHub's 14 days
  ([docs/TRAFFIC.md](docs/TRAFFIC.md); #36).
- This changelog and the release procedure ([RELEASING.md](RELEASING.md); #13).
- The SFT prototype, as it is: E\* of each netting set under the Financial
  Collateral Comprehensive Method with supervisory haircuts, its risk weight,
  RWA and own funds, in an 11-sheet workbook template with a demo portfolio,
  whose every input cell explains its entry and refuses typed values outside
  the engine's limits; eight `SFT_*` worksheet functions; an independent Python
  reference and a LibreOffice regression. Tables, the run summary and the
  TestCatalogue lookups are reached through workbook names and header texts,
  not fixed rows and columns (#171). It is not the methodology of
  record, and departs from recorded decisions where
  [PROTOTYPE.md](docs/methodology/PROTOTYPE.md) lists it (#131, #135, #145).

### Validation

- The harness and the example suite at `1a6f0f7` were imported, compiled
  and run in Excel 16.0 (build 20430, 64-bit, VBA7) on Windows 11 Enterprise
  25H2. The normal run passed (9 cases, 18 assertions), both failure paths
  ended as documented, and `tools/check_excel_run.py` passed on the evidence
  ([tested environments](INSTALLATION.md#tested-environments); #9).
- On the same host, `tools/Sync-KSFT-VBA.ps1` imported the test modules into
  a workbook built from the template, in development mode, and the example
  suite passed in it (#102).
- The prototype at `fbaee34`, built on the same host with
  `tools/Sync-KSFT-VBA.ps1`, compiled and ran on the demo portfolio: the
  TestCatalogue's 107 checks against the reference and the 25 formula tests
  passed, the example suite passed, and `tools/check_excel_run.py` passed on
  the evidence (#135). The expected values are illustrative: an independent
  implementation of the specification, not published worked examples.

### Known limitations

- Excel has been run on one host only; 32-bit Office and other builds are
  untested.
- The calculation is a prototype, not the methodology of record. Its
  departures (a BCBS regime, a QCCP class, the open product, collateral and
  output decisions, formulas not yet checked against the registered copies,
  no test case in the format of TEST_CASES.md, `src/core` modules that touch
  Excel) are left to later milestones
  ([PROTOTYPE.md](docs/methodology/PROTOTYPE.md)).
- There is no production workbook.

[Unreleased]: https://github.com/danielep71/K-SFTs/commits/main
