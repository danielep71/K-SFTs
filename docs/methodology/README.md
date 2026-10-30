# 📐 Methodology

This directory records **what K-SFTs will implement and on what
authority**: the controlling and reference texts with their dates, the
primary-source provenance of every rule, the scope decisions, assumptions,
and the trace from each rule to its formula, code and tests. The numerical
test-case contract is in [TEST_CASES.md](TEST_CASES.md), the supported scope
of v1.0.0 in [SUPPORT_MATRIX.md](SUPPORT_MATRIX.md); copies of the texts
relied on are in [sources/](sources/README.md). No derivative SA-CCR content
applies here.

> [!IMPORTANT]
> **Status on 2026-10-29: the methodology of record is not written yet.** The
> sources, their roles and the first scope are decided; no rule has entered
> the [traceability](#traceability) table. The prototype implements
> formulas, haircuts and parameters outside this methodology
> ([PROTOTYPE.md](PROTOTYPE.md)); they are not validated. On 2026-10-30 I
> decided the supported scope of v1.0.0 and closed D4, D6 and D7
> ([SUPPORT_MATRIX.md](SUPPORT_MATRIX.md), #172).

<a id="regulatory-basis"></a>

## ⚖️ Regulatory basis

Decided by me on 2026-10-24 in
[#8](https://github.com/danielep71/K-SFTs/issues/8) (D1).

| Role | Source | Use |
| --- | --- | --- |
| **Controlling** | EU Capital Requirements Regulation, Regulation (EU) No 575/2013, incorporating the amendments of Regulation (EU) 2024/1623 (CRR3) | The calculation specification. Every implemented rule cites a CRR article and paragraph. The original 2013 text is insufficient: CRR3 changed the relevant provisions |
| **Methodological reference** | Basel Framework, the in-force version of each chapter | Explains intent and structure. It never overrides the CRR and is not a calculation regime |
| **Interpretation** | EBA Single Rulebook Q&As | Edge cases and the line between calculation outputs and reporting measures. Each answer is checked against the regulatory version it was given under before it is relied on |

SFTs have their own collateral and netting methods: the Basel Framework
directs them to the credit risk mitigation framework, while SA-CCR covers
mainly derivatives and long-settlement transactions. No SA-CCR rule applies.

### CRR provisions

| Articles | Role in the SFT methodology |
| --- | --- |
| 220 | Supervisory volatility adjustments for transactions covered by eligible master netting agreements |
| 221 | Internal models approach for master netting agreements (out of scope, below) |
| 222 | Financial Collateral Simple Method (out of scope, below) |
| 223 | Financial Collateral Comprehensive Method: the core exposure and collateral adjustment |
| 224, 226, 227 | Supervisory volatility adjustments, their scaling, and the conditions for a 0% adjustment |
| 192 to 197, 205 to 207 | Definitions, collateral eligibility, and the requirements for master netting agreements and financial collateral that an implementation also needs |

Article 225 (own-estimate volatility adjustments) is deleted by Regulation
(EU) 2024/1623.

### Basel chapters

| Chapter | Use |
| --- | --- |
| CRE50, CRE51 | Definitions, transaction scope and the selection of the applicable approach |
| CRE22 | Financial collateral, supervisory haircuts, holding periods, remargining and master netting |
| CRE32 | IRB provisions, including the VaR-model approach for SFTs (out of scope, below) |
| CRE53 | Internal models method (out of scope, below) |
| CRE54 | Exposures to central counterparties (out of scope, below) |
| CRE56 | Minimum haircut floors for certain SFTs: a separate requirement whose jurisdictional applicability is to be assessed (out of scope, below) |

<a id="scope"></a>

## 🎯 First scope

Decided by me on 2026-10-24 in #8:

- **Approach (D3):** the Financial Collateral Comprehensive Method (CRR
  Article 223) with supervisory volatility adjustments (Articles 224, 226,
  227).
- **Netting (D5):** bilateral SFTs, as single transactions or under an
  eligible master netting agreement (Article 220, with the requirements of
  Article 206).

**Separate scope decisions**, each needing its own issue before any work:
the internal models approach (CRR Article 221, CRE32, CRE53), CCP-cleared
SFTs (CRE54), the CRE56 haircut floors, the Financial Collateral Simple
Method (Article 222), and other prudential measures such as large exposures,
leverage or reporting.

Products (D4), collateral (D6) and outputs (D7) were decided by me on
2026-10-30 in #172: all five products, the prototype's collateral rules, and
E\* with its intermediate amounts as the supported output, the risk weight,
RWA and own funds requirement being illustrative. CCP exposures are out of
scope. The [support matrix](SUPPORT_MATRIX.md) gives every combination.

<a id="version-applicability"></a>

## 📅 Version applicability

A calculation uses the version of the controlling text that applies on its
calculation date; for a test case, its `valuation_date`. Decided 2026-10-24
in #8 (D2). The table below is read by `tools/check_test_cases.py`: an
expected result is accepted only for a regime listed here, and only when its
case's `valuation_date` falls within that row's dates. A blank *To* means no
later version is registered yet.

| Regime | Source ID | From | To |
| --- | --- | --- | --- |
| `CRR` | `CRR` | 2025-01-01 | |

**Why 2025-01-01:** Regulation (EU) 2024/1623 applies from 1 January 2025
(its Article 2), and none of its points amending Articles 194 to 227 (points
(111) to (133) of its Article 1) is among those applying from 9 July 2024.
In the registered consolidated text, the only amendment markers within
Articles 192 to 228 are M5, M8, M9, M17 (Regulation (EU) 2024/1623) and the
corrigendum C2: none of the later acts M18 to M21 changes them. Calculation
dates before 1 January 2025 need the pre-CRR3 text, which is not registered,
and are not supported.

<a id="source-register"></a>

## 📚 Source register

Cite sources by ID in code `REFERENCE` sections, in the
[traceability](#traceability) table and in test cases. Reference values may
only come from a source in this register; add a row before citing a new one.
A later version is a new row, not an edit. Retrieval details, hashes and
reuse terms are in [sources/README.md](sources/README.md).

| ID | Role | Source | Version | Status |
| --- | --- | --- | --- | --- |
| `CRR` | Controlling | Regulation (EU) No 575/2013, consolidated text | 26.06.2026, version 021.001 | **Relied on.** Copy in `sources/`: page 1 and pages 281 to 323 (Chapter 4, Articles 192 to 227, and the start of Article 228). Read on 2026-10-24: Article 220(1)–(5) in full; Articles 221, 223(1) and 224(1) in part; the headings of Articles 222 to 228. Read on 2026-10-29 to check the workbook text ([#197](https://github.com/danielep71/K-SFTs/issues/197)): Articles 220(3), 223(6)–(7), 224 (Tables 1, 3 and 4, paragraphs 2 and 4 to 6), 226 and 227; the prototype's haircut table equals Tables 1, 3 and 4 (10-day column). Formulas are not yet transcribed or checked |
| `CRR3` | Controlling (amending act) | Regulation (EU) 2024/1623 | OJ L, 19.6.2024 | Read on 2026-10-24 for its date of application only (Article 2) and the list of its points amending Articles 194 to 227 |
| `CRE22` | Methodological reference | Basel Framework, CRE22 *Standardised approach: credit risk mitigation* | In force 1 January 2023, published 26 November 2020 | **Relied on.** Copy in `sources/` (the chapter PDF of this version). CRE22.40 to CRE22.65 read on 2026-10-24; on 2026-10-29 the prototype's haircut table was found equal to CRE22.49 (#197). Formulas are not yet transcribed |
| `CRE22-2028` | Methodological reference | CRE22, future version | In force 1 November 2028, published 28 October 2025 | Registered, not relied on |
| `CRE32` | Methodological reference | CRE32 *IRB approach: risk components* | In force 1 January 2023, published 27 March 2020 (future version in force 1 November 2028) | Copy in `sources/`; not read; out of scope |
| `CRE50` | Methodological reference | CRE50 *Counterparty credit risk definitions and terminology* | In force 15 December 2019, published 5 July 2024 | Copy in `sources/`; not read |
| `CRE51` | Methodological reference | CRE51 *Counterparty credit risk overview* | In force 1 January 2023, published 27 March 2020 (future version in force 1 November 2028) | Copy in `sources/`; CRE51.8 read on 2026-10-29 (#197) |
| `CRE53` | Methodological reference | CRE53 *Internal models method for counterparty credit risk* | In force 1 January 2023, published 5 June 2020 | Copy in `sources/`; not read; out of scope |
| `CRE54` | Methodological reference | CRE54 *Capital requirements for bank exposures to central counterparties* | In force 1 January 2023, published 27 March 2020 | Copy in `sources/`; CRE54.7 read on 2026-10-29 (#197); out of scope |
| `CRE56` | Methodological reference | CRE56 *Minimum haircut floors for securities financing transactions* | In force 1 January 2023, published 1 July 2021 | Copy in `sources/`; CRE56.1 and 56.2 read on 2026-10-29 (#197); out of scope until its applicability is assessed |
| `EBA-QA-2017_3375` | Interpretation | Maturity mismatch for exposures under master netting agreements (CRR Article 220) | Final, published 24/07/2020 | Copy in `sources/`. **Pre-CRR3:** it names the own-estimates approach (Article 225, now deleted) and Articles 223 to 226; to be re-checked against Article 220 as amended before reliance |
| `EBA-QA-2015_2010` | Interpretation | EAD of over-collateralised SFTs under the FCCM (COREP C 07.00) | Final, published 26/11/2021 | Copy in `sources/`. Reporting Q&A; pre-CRR3; to be re-checked before reliance |
| `EBA-QA-2017_3590` | Interpretation | Original exposure value of SFTs under Article 220 (large exposures, C 28.00) | Final, published 04/10/2019 | Copy in `sources/`. Reporting Q&A; pre-CRR3; to be re-checked before reliance |
| `EBA-QA-2022_6630` | Interpretation | Notional and current-market-value concepts for SFTs (COREP C 34.02) | Final, published 28/04/2023 | Copy in `sources/`. Reporting Q&A; pre-CRR3; to be re-checked before reliance |

An external implementation, the only source of an `independent` reference
value, is registered with an ID starting `IMPL-`
([reference classes](TEST_CASES.md#reference-classes)). None is registered.

The consolidated CRR has no legal effect, as its first page states; the
authentic text is the one published in the Official Journal. The Basel
Framework is a standard for jurisdictions to implement, not law. EBA Q&As
are non-binding.

<a id="assumptions-and-scope"></a>

## 🧭 Assumptions

| Topic | Current position |
| --- | --- |
| Units | Amounts in a stated currency, unscaled; rates and haircuts as decimals; periods in years unless the name says otherwise ([VBA_HOUSE_STYLE.md](../VBA_HOUSE_STYLE.md#units-and-domains)). Test cases follow the same rules ([TEST_CASES.md](TEST_CASES.md#fixture-format)) |
| Inputs | Transaction, collateral and netting-set data supplied by the caller. K-SFTs will not price securities or source market data, FX rates or ratings |
| Supervisory parameters | Taken only from a registered controlling text, recorded with its locator, and never typed in code without a `REFERENCE` |
| Numeric conventions | `Double` arithmetic, no intermediate rounding; tolerances are explicit on every expected value ([TEST_CASES.md](TEST_CASES.md#tolerance)) (D8) |
| Calculation outputs and reporting | A calculation output (such as the fully adjusted exposure value E*) is kept distinct from any reporting measure; the reporting Q&As in the register are used to draw that line, not as calculation rules |
| Derivative SA-CCR | Out of scope. No SA-CCR formula, supervisory factor, bucket, threshold or example is a default or a reference here |

<a id="open-decisions"></a>

## ❓ Open decisions

The decisions are listed in
[REPOSITORY_STRUCTURE.md](../REPOSITORY_STRUCTURE.md#-open-product-and-methodology-decisions).
D1, D2, D3, D5 and D8 are closed by this document and
[TEST_CASES.md](TEST_CASES.md); D4, D6 and D7 by
[SUPPORT_MATRIX.md](SUPPORT_MATRIX.md) on 2026-10-30 (#172). Each rule they
admit enters the [traceability](#traceability) table through #173.

<a id="regime-differences"></a>

## 🔀 CRR and Basel

There is one calculation regime, `CRR`; the prototype's `BCBS` regime is
rejected from v1.0.0 ([SUPPORT_MATRIX.md](SUPPORT_MATRIX.md), #248). Where the Basel reference differs
from the CRR, the CRR applies and the difference is recorded here before any
code depends on it, so that the method stays explainable. None is recorded
yet. Both texts define the exposure of SFTs under a master netting agreement
with a current-exposure term, a net-exposure and a gross-exposure term over
the security issues or groups, and currency-mismatch adjustments (CRR
Article 220(2)–(3); CRE22.64–22.65); whether every term, haircut and
threshold agrees has not been compared.

<a id="traceability"></a>

## 🔗 Traceability

Every implemented rule is traced **source → article and paragraph →
effective date → formula → test case**. Rule IDs are
`CRR.<article>[.<paragraph>]`, for example `CRR.220.3`. A rule enters this
table when code first implements it.

| Rule ID | Source | Article and paragraph | Effective from | Formula | Implemented in | Cases | Status |
| --- | --- | --- | --- | --- | --- | --- | --- |
| *None yet* | | | | | | | |

- **Formula** is transcribed in `docs/methodology/` from the registered
  copy, naming the page and the formula or table it comes from, and checked
  against it.
- **Cases** lists the case IDs; **Status** is the validation status of the
  [completeness policy](TEST_CASES.md#completeness-policy).
- A Basel paragraph that explains a rule may be added in **Source** next to
  the CRR, never instead of it.

## 📏 Rules

- Every formula, parameter or threshold used later cites its registered
  controlling source and locator here.
- No derivative SA-CCR content is a default for SFTs.
- Source documents are copied only under their publishers' terms, recorded
  in [sources/README.md](sources/README.md), and never edited.
