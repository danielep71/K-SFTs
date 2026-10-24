# 🗂️ Repository structure

This document is authoritative for **the project boundary, the project
profile, where every durable artifact belongs and the boundary between the
calculation core, the public interface and the Excel host**. Its structure is
adapted from K-SACCR's `docs/REPOSITORY_STRUCTURE.md` (see
[NOTICE.md](../NOTICE.md)); none of K-SACCR's calculation content applies here.
Governance is in [GOVERNANCE.md](GOVERNANCE.md).

## 🎯 Scope and status

**Scope.** CCR-K-SFTs is intended to calculate counterparty credit risk
exposure for **securities financing transactions** (SFTs): repurchase and
reverse repurchase agreements, securities lending and borrowing, and margin
lending, together with the collateral exchanged under them. Which regulatory
framework, approach and products it supports is not decided yet; those
choices are recorded as [open decisions](#-open-product-and-methodology-decisions).

**Out of scope:** derivatives exposure, which K-SACCR covers with SA-CCR;
valuation adjustments such as CVA; and capital or risk-weighted assets beyond
the exposure measure, unless a later decision adds them.

**Status on 2026-10-24: repository setup only.** The milestone
*v0.1.0 - Repo setup* builds tooling, governance and conventions. There is no
calculation engine, public API, workbook or validated SFT method yet, and
nothing in this repository should be used to compute exposure. Production
calculations, numerical certification and workbook packaging belong to later
milestones.

<a id="profile-decision"></a>

## 🧭 Profile decision

**CCR-K-SFTs uses the `application` profile.** Decided by the maintainer on
2026-10-24 in [#2](https://github.com/danielep71/CCR-K-SFTs/issues/2).
`.github/label-policy.json` already selects the same profile for the label
catalogue.

| Question | Answer |
| --- | --- |
| Who calls it? | A user working in the CCR-K-SFTs workbook: entering transactions, collateral and netting sets, running the calculation and reading the results. |
| What does it own? | The workbook: input and result sheets, the calculation engine, and their startup and packaging. |
| Lifecycle | The workbook is the deliverable. It is built from the exported source in this repository, never edited as the source of truth. |
| Supported environments | Excel for Windows, to be specified in [#4](https://github.com/danielep71/CCR-K-SFTs/issues/4); intended and tested environments are recorded separately. |

**Why not `library`:** the end product is a workbook with its business user
interface, not only callable functions. Starting as a library would force a
profile change later.

**What stays library-like:** the calculation core is host-independent. It
never touches the Excel object model, workbooks, sheets or ranges, so it can
be tested without a workbook and reused if a library deliverable is ever
wanted.

### Python

Python is **tooling and reference testing only**:

- repository tooling in `tools/`: checks, automation and evidence scripts;
- independent reference derivations in `tests/derivations/`, which compute
  expected values for test fixtures from the documented method, separately
  from the VBA implementation.

Python is never shipped, never called by the workbook and never part of the
calculation product. A reference derivation is evidence for a test, not a
second implementation to be kept in step.

<a id="repository-layout"></a>

## 📁 Repository layout

```text
.github/            workflows, issue and pull-request templates, label catalogue
                    and its scripts, Dependabot
docs/               governance, architecture, contracts, methodology
  methodology/      regulatory basis, decisions, provenance, fixture contract
src/                production VBA source: the only input to the workbook
tests/              verification source, synthetic fixtures, expected values
tools/              repository checks and automation (Python)
```

The root holds the project documents (`README.md`, `SECURITY.md`,
`NOTICE.md`, `LICENSE`) and the repository configuration (`.editorconfig`,
`.gitattributes`, `.gitignore`, `pyproject.toml`).

| Directory | Owns | Must not own |
| --- | --- | --- |
| `src/` | Production components that go into the workbook, and the workbook template | Tests, examples, built workbooks |
| `tests/` | Test modules, synthetic fixtures, reviewed expected values and their reference derivations | Production entry points, run output |
| `examples/` | Examples that use only the public API | Tests, real data |
| `docs/` | Governance, architecture, contracts and [methodology](methodology/README.md) | Copies of root documents, generated output |
| `tools/` | Deterministic checks, build and evidence scripts | Calculation logic |

`src/`, `tests/`, `tools/` and `docs/methodology/` each have a `README.md`.
A subdirectory is created with its first real file, never empty:
`examples/` and every subdirectory listed below are reserved until then.
Run output goes to the ignored `test-results/` directory and is never
committed.

### Production source

| Location | Holds | Visibility |
| --- | --- | --- |
| `src/core/` | Calculation engine: validation, haircuts and exposure arithmetic, netting-set aggregation, shared helpers | In-project only. Every standard module declares `Option Private Module`. |
| `src/modules/` | Public facade: the supported entry points, constants and enumerations | Supported API, listed in [PUBLIC_API.txt](PUBLIC_API.txt) |
| `src/classes/` | Class modules, for example transaction, collateral or netting-set objects | Status stated in each class header |
| `src/workbook/` | Exported document modules (`ThisWorkbook`, sheet modules) and the macro-free workbook template | Host glue only; no calculation logic |
| `src/forms/` | UserForms, `.frm` beside its `.frx`, only if a form is needed | Host glue only |

### Verification and examples

| Location | Holds |
| --- | --- |
| `tests/modules/` | The VBA test harness, case runner and suites |
| `tests/fixtures/` | Synthetic inputs: transactions, collateral, netting sets |
| `tests/expected/` | Expected results with their provenance |
| `tests/derivations/` | Python reference derivations that compute expected results, cited by each case |
| `examples/modules/` | Example modules calling the public facade only |

The fixture schema, tolerances and provenance rules are defined in
[#8](https://github.com/danielep71/CCR-K-SFTs/issues/8).

## 🔀 Dependency direction

```text
src/workbook  ──▶  src/modules (facade)  ──▶  src/core
                         │                       ▲
src/classes  ◀───────────┘                       │
tests/       ────────────────────────────────────┘ (may test core directly)
examples/    ──▶  src/modules only
```

- `src/core` depends on nothing outside `src/core` (and `src/classes` once
  classes exist). No Excel object model, no workbook, sheet or range access,
  no user interface, no module-level run state.
- `src/modules` validates inputs at the public boundary and delegates to the
  core. It holds no formulas of its own.
- `src/workbook` and `src/forms` connect sheets, events and forms to the
  facade. They do not compute.
- `tests` may call core procedures directly; `examples` may not.

<a id="public-api-boundary"></a>

## 🧩 Public API boundary

A VBA `Public` declaration is not automatically supported API.

- **Supported:** declarations in `src/modules/` that are listed in
  [PUBLIC_API.txt](PUBLIC_API.txt). Changing one is a contract change: it
  needs its own issue, a documented reason and a changelog entry once
  releases exist.
- **In-project only:** everything in `src/core/`, kept off the external
  surface by `Option Private Module`, and anything in `src/modules/` not
  listed in the inventory.
- **Never supported:** test and example modules; they are not part of the
  workbook.

### Inventory convention

`docs/PUBLIC_API.txt` is the inventory of supported declarations. It has no
entries yet. Each supported declaration gets two lines:

```text
<component><TAB><kind><TAB><name>
# SIG<TAB><component><TAB><kind><TAB><name><TAB><normalized signature>
```

`kind` is `Function`, `Sub`, `Property Get`, `Property Let`, `Property Set`,
`Const`, `Enum` or `Type`. The signature is the declaration with whitespace
normalized to single spaces. Lines starting with `#` and blank lines are
comments, except `# SIG` lines. A check that compares the inventory with the
exported source is planned with the static checks in
[#6](https://github.com/danielep71/CCR-K-SFTs/issues/6).

### Input and output contract

The calculation will accept transactions, collateral and netting sets and
return exposure measures. Their contract is **not defined yet**. It is
written in `docs/contracts/` before the first facade procedure, and must fix
at least:

- the record types and identifiers: transaction, collateral item, netting
  set, counterparty;
- every field's type, unit and domain, including currencies, dates,
  day-count and sign conventions;
- validation rules and how an invalid input is reported (error codes or
  messages, never a silently substituted default);
- the output measures and their units, and how a partial or failed
  calculation is reported.

The contract follows the [open decisions](#-open-product-and-methodology-decisions):
it cannot be fixed before the regulatory framework and approach are chosen.

## ❓ Open product and methodology decisions

No regulatory formula, haircut, factor, bucket or netting assumption is a
default of this project until a decision below records it, with its source,
in `docs/methodology/`. In particular, K-SACCR's derivative SA-CCR formulas,
supervisory factors, asset-class buckets and netting-set assumptions **do not
apply to SFTs** and are never imported as defaults.

| # | Decision | Why it matters | Where it is made |
| --- | --- | --- | --- |
| D1 | Regulatory framework and jurisdiction. Candidates to confirm against the source text: the Basel Framework's credit risk mitigation standard (CRE22: comprehensive approach for collateralised transactions and SFTs under master netting agreements), and the EU CRR's financial collateral comprehensive method (Articles 220 to 223) | Fixes every formula, parameter and term, and the citation format | [#8](https://github.com/danielep71/CCR-K-SFTs/issues/8) |
| D2 | Rule version and effective date | Results must be reproducible against one dated text | #8 |
| D3 | Approach: supervisory haircuts, own-estimate haircuts, or a model-based approach | Determines inputs, parameters and validation | Later methodology issue |
| D4 | Products in scope: repo and reverse repo, securities lending and borrowing, margin lending | Fixes the transaction record and its validation | Later methodology issue |
| D5 | Netting: transactions under a master netting agreement versus single transactions | Fixes the netting-set record and aggregation | Later methodology issue |
| D6 | Collateral: eligibility, haircut source, currency mismatch, holding period and remargining frequency | Fixes the collateral record and the haircut arithmetic | Later methodology issue |
| D7 | Output measures: exposure only, or exposure with intermediate amounts for audit | Fixes the result sheets and the output contract | Later methodology issue |
| D8 | Numeric conventions: precision, rounding and tolerances | Fixes test tolerances and comparison rules | #8 |

A decision is closed only by a reviewed change that records it, with its
primary source, in `docs/methodology/` and updates this table.

## 🚦 Placement rules

1. VBA components (`.bas`, `.cls`, `.frm`, `.frx`) live only in `src/core/`,
   `src/modules/`, `src/classes/`, `src/workbook/`, `src/forms/`, `tests/`
   or `examples/`.
2. Every standard module in `src/core/` declares `Option Private Module`.
3. Test and example modules never go into the production import set.
4. Fixtures and examples are synthetic. Never commit real transactions,
   counterparties, collateral or portfolios ([SECURITY.md](../SECURITY.md)).
5. Workbooks are built from source and never committed, except at an exact
   path explicitly re-included in `.gitignore`; no such path exists yet.
6. A new location for VBA components, fixtures or the workbook becomes
   contractual only when this document, the directory README and the checks
   are updated together.

Rules 1 to 3 and 5 are to be enforced by the static checks of
[#6](https://github.com/danielep71/CCR-K-SFTs/issues/6); until then they are
enforced by review. Naming, encoding and import/export conventions are set in
[#4](https://github.com/danielep71/CCR-K-SFTs/issues/4).

## 📚 Directory guides

- [`../src/README.md`](../src/README.md)
- [`../tests/README.md`](../tests/README.md)
- [`methodology/README.md`](methodology/README.md)
- [`../tools/README.md`](../tools/README.md)

---

**Structure principle:** one home per responsibility, the engine never touches
Excel, only the facade is supported API, and no SFT rule exists until it is
decided and sourced.
