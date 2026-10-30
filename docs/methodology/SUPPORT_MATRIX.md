# 🧾 Support matrix for v1.0.0

This document fixes **what K-SFTs v1.0.0 supports**: every product,
collateral, netting, counterparty, regime, date and output combination is
marked supported, rejected, illustrative or out of scope, with the behaviour
of the workbook and the reason. It records my decisions of 2026-10-30
([#172](https://github.com/danielep71/K-SFTs/issues/172)), which close the
open decisions D4, D6 and D7
([REPOSITORY_STRUCTURE.md](../REPOSITORY_STRUCTURE.md#-open-product-and-methodology-decisions)).

> [!IMPORTANT]
> A **supported** item is the scope that v0.3.0 traces to the CRR, tests and
> certifies; it is not validated yet. Where the 0.2.0 prototype still behaves
> differently, the *Prototype today* column says so, and
> [#248](https://github.com/danielep71/K-SFTs/issues/248) makes the workbook
> match before the numerical fixtures of
> [#174](https://github.com/danielep71/K-SFTs/issues/174) are written.

## Contents

- [Decisions](#decisions)
- [Statuses](#statuses)
- [Matrix](#matrix)
- [Prototype departures](#prototype-departures)
- [Architecture waiver](#architecture-waiver)
- [Deferred beyond v1.0.0](#deferred-beyond-v100)

<a id="decisions"></a>

## ⚖️ Decisions

Taken by me on 2026-10-30 in #172.

| Topic | Decision | Effect |
| --- | --- | --- |
| Regime | **CRR only**, as D1 (#8) decided. The Basel Framework stays the methodological reference | A netting set or default regime set to `BCBS` is rejected (#248) |
| Products (D4) | **All five:** repo, reverse repo, securities lending, securities borrowing, margin lending | D4 closed; no change to the prototype |
| Collateral (D6) | **As the prototype:** cash, debt, equities, gold and funds under the supervisory haircut table; non-eligible securities lent at 30%, non-eligible collateral received not recognised; the 0% haircut with its declared conditions; currency mismatch at 8%; holding periods and remargining as PROTOTYPE.md specifies | D6 closed; no change to the prototype |
| Outputs (D7) | **E\* and its intermediate amounts are supported.** The risk weight, RWA and own funds requirement stay as **illustrative** outputs | D7 closed; RW, RWA and own funds are labelled illustrative (#248) |
| CCP | **Bilateral SFTs only**, as D3 and D5 decided. Exposures to CCPs, qualifying or not, are out of scope | The `QCCP` counterparty class is rejected (#248) |
| Calculation date | **From 2025-01-01**, as D2 decided ([version applicability](README.md#version-applicability)) | An earlier reporting date is rejected (#248) |
| Architecture | **Waiver until v1.0.0** for the core/worksheet separation | [Below](#architecture-waiver); #249 |

<a id="statuses"></a>

## 🏷️ Statuses

| Status | Meaning |
| --- | --- |
| **Supported** | In the v1.0.0 scope: traced to the CRR (#173), covered by fixtures (#174) and regression tests (#175), and certified in Excel |
| **Entered** | Supported, but the value is a judgement or data the user enters and the workbook cannot check; the user is responsible for it |
| **Illustrative** | Calculated and shown, labelled illustrative; not traced, tested or certified as a regulatory result |
| **Rejected** | Refused by the workbook with a message: the netting set is INVALID, the transaction is rejected (INCOMPLETE), or the run stops |
| **Out of scope** | Not representable in the inputs, so it cannot be entered; no behaviour is defined |

<a id="matrix"></a>

## 🧮 Matrix

### Transactions and netting

| Item | Status | Behaviour | Prototype today | Reason |
| --- | --- | --- | --- | --- |
| `REPO`, `REVREPO`, `SECLEND`, `SECBORROW` | Supported | 5 BD holding period | Same | D4; CRR Art. 224(2) |
| `MARGINLEND` | Supported | 10 BD holding period | Same | D4; CRR Art. 224(2) |
| Any other transaction type | Rejected | The transaction is rejected; its netting sets are INCOMPLETE | Same | D4 |
| Single transactions (no eligible master netting agreement) | Supported | E\* per transaction, summed | Same | D5; CRR Art. 223 |
| Netting under an eligible master netting agreement | Supported | Netted E\* of the netting set | Same | D5; CRR Art. 220(3) |
| Eligibility of the master netting agreement | Entered | `MNA = Y` is taken as declared | Same | CRR Art. 206: legal conditions the inputs cannot show |
| Repo-style and margin lending in one netting set | Supported | 10 BD for the whole set (assumption 1 of PROTOTYPE.md) | Same | CRE22.58; the conservative reading of Art. 224(2) |
| More than 5,000 transactions, illiquid collateral, margin disputes | Supported | 20 BD minimum; doubled after disputes | Same | CRR Art. 224(2), 285(3)–(4) |
| Holding-period override, remargining frequency | Supported | An override can only lengthen the period | Same | CRR Art. 224(2), 226 |
| Cross-product netting with derivatives | Out of scope | Not representable | Same | D5; SA-CCR does not apply to SFTs |
| Commodity repos, secured lending (20 BD) | Out of scope | No product type for them | Same | D4 |

### Collateral and haircuts

| Item | Status | Behaviour | Prototype today | Reason |
| --- | --- | --- | --- | --- |
| Cash | Supported | Haircut 0 | Same | CRR Art. 224 Table 1 |
| Debt of sovereigns, institutions, corporates, securitisations, by CQS and residual maturity | Supported | 10-day haircut from the table, scaled | Same | D6; CRR Art. 224 Table 1 |
| Unrated debt of an institution | Supported | Treated as CQS 2–3 | Same | CRR Art. 197(4), 224(6) |
| Equities: main index, other listed | Supported | 20% / 30%, scaled | Same | CRR Art. 224 Table 3 |
| Gold | Supported | 20%, scaled | Same | CRR Art. 224 Table 3 |
| Funds | Entered | Look-through or highest haircut entered on the leg | Same | CRR Art. 224(5) |
| Non-eligible securities lent (`OTHER`, or no table row) | Supported | 30%, scaled | Same | CRR Art. 224(4) |
| Non-eligible collateral received | Supported | Not recognised: excluded from C and the positions | Same | CRR Art. 197, 223 |
| 0% haircut | Supported, conditions partly entered | Checked: repo-style, cash and eligible central government debt, one currency, daily remargining. Entered: the 0% risk weight of the securities, core market participant, documentation, settlement | Same | CRR Art. 227; my decision in #141 |
| Currency mismatch | Supported | 8%, scaled, on legs received in another currency; under a master netting agreement, on net positions per currency | Same | CRR Art. 224 Table 4, 220(3) |
| FX rates, market values, ratings (CQS) | Entered | Taken as given; K-SFTs prices nothing | Same | [Assumptions](README.md#assumptions-and-scope) |
| Short-term issue ratings (A-1 and similar) | Out of scope | No input for them | Same | Not in the prototype's tables |
| Own-estimate haircuts, VaR models, internal models | Out of scope | No input for them | Same | D3; CRR Art. 225 deleted by CRR3, Art. 221 out of scope |
| Haircut floors for non-centrally cleared SFTs (CRE56) | Out of scope | Not calculated | Same | D3: a separate scope decision |

### Counterparties, regime and dates

| Item | Status | Behaviour | Prototype today | Reason |
| --- | --- | --- | --- | --- |
| Counterparty classes `SOV`, `INST`, `CORP` | Supported for E\* | E\* does not depend on the class; it selects the illustrative risk weight | Same | D5 |
| `QCCP` counterparty class | Rejected | The netting set is INVALID | Computed with a 2% risk weight | D3, D5: exposures to CCPs are out of scope |
| Regime `CRR` | Supported | The only regime | One of two | D1 |
| Regime `BCBS` | Rejected | A netting set set to `BCBS` is INVALID; a default regime of `BCBS` on Params stops the run | Computed as a second regime | D1 |
| Reporting date from 2025-01-01 | Supported | The CRR consolidated text of 26.06.2026 applies | Same | D2 |
| Reporting date before 2025-01-01 | Rejected | The run stops | Computed | D2: the pre-CRR3 text is not registered |
| Reporting currency, settlement currency of the agreement | Entered | Any 3-letter code with an FX rate | Same | Inputs |

### Outputs

| Item | Status | Behaviour | Prototype today | Reason |
| --- | --- | --- | --- | --- |
| E\* per netting set | Supported | On Results, with the status | Same | D7 |
| Intermediate amounts: positions, scaled haircuts, holding period, ΣE, ΣC, net, gross, N, FX term | Supported | On Results and Positions | Same | D7: exposure with its intermediate amounts for audit |
| Status `VALID`, `INVALID`, `INCOMPLETE`, `NO TRANSACTIONS`; the Checks log | Supported | As PROTOTYPE.md specifies | Same | Run integrity |
| Run class `REGULATORY` / `WHAT-IF` | Supported | WHAT-IF when a parameter or table differs from the reference set | Same | Traceability of the parameter set |
| Risk weight, RWA, own funds requirement | Illustrative | Shown, labelled illustrative | Shown alike with E\* | D7; CRE20 is not registered and the risk-weight table is not checked |
| Risk-weight override | Illustrative | Replaces the table value and makes the run WHAT-IF | Same | Follows the risk weight |
| Maturity mismatch, the SCRA sovereign floor, CVA, buffers, Pillar 2, large exposures, leverage, reporting | Out of scope | Not calculated | Same | The [first scope](README.md#scope) |

<a id="prototype-departures"></a>

## 🔧 Prototype departures

Every departure that PROTOTYPE.md lists, and every assumption it states, is
resolved here or assigned to the issue that resolves it.

| Departure | Resolution |
| --- | --- |
| Two regimes, `CRR` and `BCBS` | CRR only; BCBS rejected (#248) |
| QCCP counterparty class | Rejected (#248) |
| Outputs E\*, RW, RWA, own funds | D7: E\* and intermediate amounts supported; RW, RWA and own funds illustrative and labelled (#248) |
| Products | D4 closed: all five supported |
| Haircut-table eligibility | D6 closed: as the prototype |
| Formulas and tables not checked against the registered copies | #173 traces and checks every supported rule; the risk-weight table stays unchecked while it is illustrative |
| Demo expected values outside the test-case format | #174 converts them, after #248 |
| `src/core/` reads and writes sheets; the formulas live in `src/modules/` | [Architecture waiver](#architecture-waiver) until v1.0.0; #249 |
| Assumption 1, mixed netting sets at 10 BD | Kept, supported |
| Assumption 2, one haircut per issue | Kept: differing descriptions make the netting set INVALID |
| Assumption 3, short-term risk weight only when every transaction qualifies | Kept, illustrative with the risk weight |
| Assumption 4, CRE20 locators by section | Illustrative with the risk weight; checked if the risk weight is ever supported |
| Assumption 5, expected values illustrative | #174 keeps the reference class honest |

<a id="architecture-waiver"></a>

## 🏗️ Architecture waiver

`src/core/` must never touch Excel, and the workbook depends on the core,
never the reverse ([REPOSITORY_STRUCTURE.md](../REPOSITORY_STRUCTURE.md#-dependency-direction)).
The prototype breaks both rules. By my decision of 2026-10-30:

| Item | Waiver |
| --- | --- |
| Scope | `CORE_Engine` and `CORE_Util` read and write sheets; `SFT_Formulas` in `src/modules/` holds the formulas and module-level run state, and `M_Main` calls it directly |
| Owner | Me |
| Mitigations | The facade stays the only entry point of the workbook; the LibreOffice regression and the Excel run cover the engine end to end; every v0.3.0 fixture fixes the results before the code moves |
| Ends | When [#249](https://github.com/danielep71/K-SFTs/issues/249) separates the core from the worksheet code, before the v1.0.0 contract is frozen ([#177](https://github.com/danielep71/K-SFTs/issues/177)) |
| Follow-up trigger | #177 cannot start while #249 is open |

<a id="deferred-beyond-v100"></a>

## 📦 Deferred beyond v1.0.0

None of these is planned; each needs its own issue and my decision:

- a ribbon or other custom user interface;
- advanced CSV import;
- a standalone Python engine;
- new prudential methods: internal models (CRR Art. 221, CRE32, CRE53), the
  Financial Collateral Simple Method (Art. 222), CCP-cleared SFTs (CRE54), the
  CRE56 haircut floors;
- a supported risk weight, RWA or own funds requirement, which needs CRE20
  and CRR Articles 114 to 122 registered and checked;
- comparison with an external implementation (an `IMPL-` reference,
  [TEST_CASES.md](TEST_CASES.md#reference-classes)).
