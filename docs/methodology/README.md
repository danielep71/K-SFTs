# 📐 Methodology

This directory records **what CCR-K-SFTs will implement and on what
authority**: the regulatory framework and its dated texts, the primary-source
provenance of every rule, assumptions, open decisions and the trace from each
rule to code and tests. The numerical test-case contract is in
[TEST_CASES.md](TEST_CASES.md). Its structure is adapted from K-SACCR's
`docs/methodology/` (see [NOTICE.md](../../NOTICE.md)); none of K-SACCR's
derivative SA-CCR content applies here.

> [!IMPORTANT]
> **Status on 2026-10-24: skeleton only.** The framework, jurisdiction and
> rule versions are decided (D1, D2) and the numeric conventions are defined
> (D8). No formula, haircut, parameter or netting treatment is implemented,
> transcribed or validated. D3 to D7 are open.

<a id="regulatory-basis"></a>

## ⚖️ Regulatory basis

Decided by the maintainer on 2026-10-24 in
[#8](https://github.com/danielep71/CCR-K-SFTs/issues/8) (D1).

| Regime | Identifier | Jurisdiction | Treatment relied on | Role |
| --- | --- | --- | --- | --- |
| EU Capital Requirements Regulation | `CRR` | European Union | Financial Collateral Comprehensive Method: Part Three, Title II, Chapter 4, Section 4, Articles 220 to 227 | **Baseline.** Default for every calculation and test case |
| Basel Framework, CRE22 | `BCBS` | Basel Committee standard, not law | Comprehensive approach for collateralised transactions, including SFTs under master netting agreements: CRE22.40 to CRE22.65 | Also supported; implemented as differences from the baseline |

- Every result will report the regime it was calculated under, and every
  expected result names its regime.
- Where the two texts agree, there is one implementation tagged with both
  rule IDs. Where they differ, the difference is recorded under
  [Regime differences](#regime-differences) before any code depends on it.
- The texts are pinned to the versions in the [source register](#source-register)
  (D2). A newer version is a new register row and a reviewed change, never a
  silent update.

<a id="source-register"></a>

## 📚 Source register

Cite sources by ID in code `REFERENCE` sections, in the
[traceability](#traceability) table and in test cases. Reference values may
only come from a source in this register; add a row before citing a new one.
A later version is a new row, not an edit.

| ID | Source | Scope relied on | Version / date | Status |
| --- | --- | --- | --- | --- |
| `CRR` | Regulation (EU) No 575/2013, consolidated text, as amended up to Regulation (EU) 2024/1623 and later acts included in the consolidation | Articles 220 to 227: master netting agreements for SFTs (220, 221), the Financial Collateral Comprehensive Method (223) and its volatility adjustments (224, 226, 227) | Consolidated text of 26.06.2026, version 021.001 | **Pinned (D2).** Read on 2026-10-24: Article 220(1)–(5) in full; Articles 221, 223(1) and 224(1) in part; the headings of Articles 222 to 228. Article 225 is shown as deleted by Regulation (EU) 2024/1623. Formulas and tables are images in the retrieved text and are not yet transcribed or checked |
| `BCBS` | Basel Framework, chapter CRE22, *Standardised approach: credit risk mitigation* | CRE22.40 to CRE22.65: comprehensive approach, haircuts, holding periods, SFTs under master netting agreements | Effective 1 January 2023, last updated 26 November 2020 (status *Current*) | **Pinned (D2).** CRE22.40 to CRE22.65 read on 2026-10-24. Formulas are images and are not yet transcribed or checked |
| `BCBS-2028` | Basel Framework, chapter CRE22, future version | Same paragraphs | Effective 1 November 2028, incorporating the technical amendment on hedging counterparty credit risk published 28 October 2025 | **Registered, not relied on.** To be reviewed before it takes effect; until then no case or code cites it |

An external implementation, the only source of an `independent` reference
value, is registered with an ID starting `IMPL-`
([reference classes](TEST_CASES.md#reference-classes)). None is registered.

<a id="retrieval"></a>

### Retrieval record

No copies of the texts are committed (decided 2026-10-24 in #8). Each row
records what was read, so that the same text can be found again. The hash is
that of the file as retrieved; the publishers generate these files on
request, so a new download need not give the same hash, and the version date
is the identifier to rely on.

| Source ID | Retrieved from | Retrieved | Format | SHA-256 of the retrieved file |
| --- | --- | --- | --- | --- |
| `CRR` | <https://publications.europa.eu/resource/celex/02013R0575-20260626> (Publications Office, content negotiated as XHTML; the same text as EUR-Lex CELEX `02013R0575-20260626`) | 2026-10-24 | XHTML, 13.2 MB | `2d041b5b32d4020fa1decb3bb64e835d6fcb747c5274a61f785b195a9dc9e6ad` |
| `BCBS` | <https://www.bis.org/committees/bcbs/basel-framework/standard/cre/22/inforce/2023-01-01/published/2020-11-26> | 2026-10-24 | HTML | `0d82986f3e9fca12d3b4a94f7f17cbc0cfbe12e704491f30704899f30bc699e1` |
| `BCBS-2028` | <https://www.bis.org/committees/bcbs/basel-framework/81416/chapter.pdf> (generated 24/10/2026 from the BIS data) | 2026-10-24 | PDF, 26 pages | `884aa3831dbb5317868b8c8ce2c153f0f04799ed0f76869ac424ba89b22cc6f5` |

The consolidated CRR has no legal effect, as its first page states; the
authentic text is the one published in the Official Journal. The Basel
Framework is a standard for jurisdictions to implement, not law.

<a id="assumptions-and-scope"></a>

## 🧭 Assumptions and scope

| Topic | Current position |
| --- | --- |
| Units | Amounts in a stated currency, unscaled; rates and haircuts as decimals; periods in years unless the name says otherwise ([VBA_HOUSE_STYLE.md](../VBA_HOUSE_STYLE.md#units-and-domains)). Test cases follow the same rules ([TEST_CASES.md](TEST_CASES.md#fixture-format)) |
| Inputs | Transaction, collateral and netting-set data supplied by the caller. CCR-K-SFTs will not price securities or source market data, FX rates or ratings |
| Supervisory parameters | Taken only from a registered text, recorded with its locator, and never typed in code without a `REFERENCE` |
| Numeric conventions | `Double` arithmetic, no intermediate rounding; tolerances are explicit on every expected value ([TEST_CASES.md](TEST_CASES.md#tolerance)) (D8) |
| Derivative SA-CCR | Out of scope. No K-SACCR formula, supervisory factor, bucket, threshold or example is a default or a reference here |
| Approach (D3) | Open. Under the registered `CRR` text, own-estimate volatility adjustments are no longer available (Article 225 deleted); the remaining treatments are the Supervisory Volatility Adjustments Approach (Articles 220, 223 to 227) and the internal models approach for IRB institutions with permission (Article 221) |

<a id="open-decisions"></a>

## ❓ Open decisions

The decisions are listed in
[REPOSITORY_STRUCTURE.md](../REPOSITORY_STRUCTURE.md#-open-product-and-methodology-decisions).
D1, D2 and D8 are closed by this document and [TEST_CASES.md](TEST_CASES.md);
D3 to D7 stay open, each to be closed by a reviewed change that records it
here with its primary source.

<a id="regime-differences"></a>

## 🔀 Regime differences

None recorded yet. Both registered texts define the exposure of SFTs under a
master netting agreement with a current-exposure term, a net-exposure and a
gross-exposure term over the security issues or groups, and currency-mismatch
adjustments (CRR Article 220(2)–(3); CRE22.64–22.65). Whether every term,
haircut and threshold agrees has not been compared; that comparison is due
before any code depends on either text.

<a id="traceability"></a>

## 🔗 Traceability

Rule IDs: `CRR.<article>[.<paragraph>]` for the CRR, for example `CRR.220.3`,
and `CRE22.<paragraph>` for Basel, for example `CRE22.65`. A rule enters this
table when code first implements it, with its source, the code that
implements it, the cases that test it and its validation status for each
regime ([completeness policy](TEST_CASES.md#completeness-policy)).

| Rule | Source and locator | Implemented in | Cases | Status |
| --- | --- | --- | --- | --- |
| *None yet* | | | | |

## 📏 Rules

- Every formula, parameter or threshold used later cites its registered
  source and locator here.
- No derivative SA-CCR content, from K-SACCR or elsewhere, is a default for
  SFTs.
- Source documents are cited, not copied, unless their licence allows it and
  a reviewed change adds them.
