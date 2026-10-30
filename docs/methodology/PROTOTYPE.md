# 🧪 Prototype specification (0.2.0)

This document specifies **what the 0.2.0 prototype calculates**: formulas,
parameter set, support matrix and traceability of the code on this branch.
It is **not** the methodology of record: [README.md](README.md) and the
decisions of [#8](https://github.com/danielep71/K-SFTs/issues/8) govern, and
every rule below enters README.md's traceability table only through a
reviewed change.

**Where the prototype departs from the recorded decisions.** By my
decision of 2026-10-25, milestone v0.2.0 integrates the prototype as it is,
with these departures; each is resolved in a later milestone
([#131](https://github.com/danielep71/K-SFTs/issues/131)). On 2026-10-29 I
moved the robustness fixes of its audit into v0.2.0
([#141](https://github.com/danielep71/K-SFTs/issues/141)); they change no E\*
for valid inputs:

| Prototype | Recorded position | Needed |
| --- | --- | --- |
| Two calculation regimes, `CRR` and `BCBS`, per netting set | One regime, `CRR`; Basel is a reference only (D1) | Drop the BCBS regime, or my decision |
| QCCP counterparty class (2%, CRE54.7 / Art. 306(1)(a)) | CCP-cleared SFTs out of scope | Remove, or a scope issue |
| Outputs E\*, RW, RWA, own funds | Outputs open (D7) | Decision D7 |
| Products REPO, REVREPO, SECLEND, SECBORROW, MARGINLEND | Products open (D4) | Decision D4 |
| Haircut-table eligibility rules | Collateral eligibility open (D6) | Decision D6 |
| Formulas and tables transcribed from CRR Art. 220-227 and CRE22 | Source register: formulas and tables not yet transcribed or checked against the registered copies | Check against `sources/` with page locators. Done for the haircut table only: it equals CRE22.49 and CRR Art. 224 Tables 1, 3 and 4 ([#197](https://github.com/danielep71/K-SFTs/issues/197)) |
| Demo expected values in `tests/derivations/prototype/` | Test cases in the format of [TEST_CASES.md](TEST_CASES.md) | Convert to `tests/fixtures/` and `tests/expected/` |
| `CORE_Engine` and `CORE_Util` read and write sheets | `src/core/` never touches Excel | Split into core and host glue |
| `CORE_Engine` calls `SFT_Formulas` in `src/modules/`, which holds the formulas, keeps module-level run state, and `M_Main` calls it directly | The dependency direction of [REPOSITORY_STRUCTURE.md](../REPOSITORY_STRUCTURE.md#-dependency-direction) | Move the formulas to the core and route the workbook through the facade |

**Resolved on 2026-10-30.** By my decisions in
[#172](https://github.com/danielep71/K-SFTs/issues/172),
[SUPPORT_MATRIX.md](SUPPORT_MATRIX.md#prototype-departures) assigns each row
above, and each assumption below, a resolution: BCBS and QCCP rejected and
RW, RWA and own funds labelled illustrative
([#248](https://github.com/danielep71/K-SFTs/issues/248)); D4, D6 and D7
closed; the source checks in #173 and the fixtures in #174; the core
separation under a waiver until v1.0.0
([#249](https://github.com/danielep71/K-SFTs/issues/249)). The table stays as
the record of 0.2.0; this document describes the prototype until #248 changes
it.

## Contents

- [Scope](#scope)
- [Calculation](#calculation)
- [Regulatory parameter set](#regulatory-parameter-set)
- [Support matrix](#support-matrix)
- [Traceability](#traceability)
- [Assumptions and open points](#assumptions-and-open-points)
- [Sources](#sources)

<a id="scope"></a>

## 🎯 Scope

K-SFTs calculates, per netting set of securities financing transactions
(repos, reverse repos, securities lending and borrowing, margin lending):

1. the fully adjusted exposure value **E\*** under the financial collateral
   comprehensive method (FCCM) with **supervisory** haircuts;
2. the counterparty **risk weight** under the standardised approach;
3. the **risk-weighted exposure amount** (RWA = E\* × RW) and the **own funds
   requirement** (8% × RWA).

SA-CCR does not apply to SFTs (CRR Art. 271(1), 273; Basel CRE51.8): the
derivatives of the same counterparty are out of scope.
Both regimes are supported per netting set: **CRR** (Regulation (EU) No 575/2013,
consolidated text of 26.06.2026, i.e. with CRR3) and **BCBS** (CRE22 effective
1 January 2023, CRE20, CRE54).

<a id="calculation"></a>

## 🧮 Calculation

### Legs, transactions, netting sets

| Level | Input | Rule |
| --- | --- | --- |
| Leg | One row of Transactions: cash, or one security issue, `LENT` (given: E) or `RECEIVED` (C) | Valued in the reporting currency; keyed to a row of the haircut table |
| Transaction | Legs sharing a TransactionID | Needs one leg of each side; legs agree on netting set, type and dates; rejected as a whole, and counted as rejected in every netting set its legs name |
| Netting set | One row of NettingSets | With an eligible master netting agreement (MNA = Y) all transactions are netted; otherwise each is calculated alone and E\* is summed |

**Input checks.** The template's data-validation rules check typed entries
only ([#145](https://github.com/danielep71/K-SFTs/issues/145)); the engine checks
every value on Run.
- A pasted value bypasses them.
- Date cells take Excel dates, so a typed `2026-09-30` becomes one. The engine
  also reads text in the form `yyyy-mm-dd` (`CORE_Util.ToSerial`), for example
  pasted or imported; the rules do not accept it typed as text.
- Every sheet is protected without a password
  ([#197](https://github.com/danielep71/K-SFTs/issues/197)). Only the input
  cells can be edited: the yellow cells and the rows of the Params tables. The
  input sheets can still be filtered, sorted and extended by inserted rows.
  The macros protect every sheet again, for the user interface only, when the
  workbook opens and before each operation, so that they can write the
  outputs.
- Rows of a Params table keep its rules when they are inserted inside the
  table (right-click a row, Insert); the cells below a table are locked.
- Business-day inputs are whole numbers, in the rules and in the engine
  ([#138](https://github.com/danielep71/K-SFTs/issues/138)): `HaircutBaseBD`,
  the three holding periods on Params, and the remargining frequency and
  holding-period override on NettingSets. `DaysPerYear` takes decimals, for
  example 365.25.
- The demo's rejection case, Transactions row 48 (asset type `BOND`), fails
  its rule on purpose: it shows the rejection path.

### Haircuts

```text
H   = H10 × sqrt((N_R + T_M − 1) / 10)                         [CRE22.59; CRR Art. 226]
Hfx = 8%  × sqrt((N_R + T_M − 1) / 10)                         [CRE22.52; Art. 224 Table 4]
```

- `H10`: supervisory haircut for 10 business days with daily remargining, from
  the haircut table (debt by issuer category, CQS band and residual maturity;
  equities; gold; cash 0) [CRE22.49; Art. 224 Tables 1 and 3]. A fund takes the
  look-through or highest haircut entered on its leg [Art. 224(5)].
- `T_M`, holding period: 5 BD for repo-style transactions, 10 BD for margin
  lending, 10 BD when a netting set under an eligible master netting agreement
  holds both (without one, each transaction takes its own); at least 20 BD for
  a netting set with more than 5,000 transactions or illiquid collateral; an
  override can only lengthen it, and one that lengthens nothing is a WARNING;
  doubled after margin-call disputes [CRE22.57-22.58; Art. 224(2),
  285(3)-(4)].
- `N_R`: business days between remargining; 1 when blank.
- **Eligibility.** A leg whose key is not in the table for its regime is not
  eligible. Lent, it takes `NonEligibleHaircut` (30%, scaled) [CRE22.54;
  Art. 224(4)]; received, it is **not recognised**: excluded from C and from the
  positions. Unrated debt of an institution is treated as CQS 2-3
  [CRE22.34(3)(b), 22.49 (unrated bank securities); Art. 197(4), 224(6)].
- **0% haircut** (declared per netting set): H = Hfx = 0. The engine checks the
  conditions visible in the inputs — repo-style only, cash and eligible central
  government debt, one currency, daily remargining — and takes the others
  (the 0% risk weight of the securities, core market participant,
  documentation, settlement) as declared [CRE22.60, 22.36-22.37; Art. 227].
  The 0% risk weight under Chapter 2 [Art. 227(2)(a)] is not derived from the
  CQS: debt of a Member State's central government denominated and funded in
  its domestic currency takes 0% at any CQS [Art. 114(4)], and the inputs do
  not show that (my decision, #141).

### Exposure value

With an eligible master netting agreement [CRE22.64-22.65; CRR Art. 220(3)]:

```text
E* = max(0, ΣE_i − ΣC_j + 0.4 · net + 0.6 · gross / √N + Σ_k |E_k^fx| · Hfx)
net   = | Σ_s E_s · H_s |        gross = Σ_s |E_s| · H_s
E_s   = net position in security issue s (lent − received)
N     = issues with E_s ≠ 0 and |E_s| ≥ 0.1 × max_s |E_s|
E_k^fx = net position (cash and securities) in currency k ≠ settlement currency
```

Without one, per transaction [CRE22.46, 22.55; CRR Art. 223]:

```text
E*_t = max(0, E · (1 + He) − C · (1 − Hc − Hfx))
```

with value-weighted basket haircuts, so `E(1 + He) = Σ V_i (1 + H_i)` and
`C(1 − Hc − Hfx) = Σ V_j (1 − H_j − Hfx_j)`; Hfx applies to a leg received in a
currency other than that of the legs lent, which must share one currency.
E\* of the netting set is `Σ_t E*_t`.

With one issue the two formulas agree (0.4 + 0.6 = 1): the regression catalogue
and `SFT_ExposureNetted(100,102,2.04,2.04,1)` check this.

### Risk weight and capital

```text
RWA = E* × RW          own funds requirement = 8% × RWA        [CRE22.48, RBC20.1; Art. 92(1)(c)]
```

| Counterparty class | Key | Rule |
| --- | --- | --- |
| SOV | `SOV_CQSn`, `SOV_UNRATED` | 0/20/50/100/100/150, unrated 100 [CRE20.7; Art. 114(2)] |
| INST rated | `INST_CQSn`, `INST_ST_CQSn` | ECRA 20/30/50/100/100/150; short-term 20/20/20/50/50/150 [CRE20.18-20.19; Art. 120] |
| INST unrated | `INST_SCRA_g`, `INST_ST_SCRA_g` | SCRA A 40, A_ENH 30, B 75, C 150; short-term 20/20/50/150 [CRE20.21; Art. 121(3)] |
| CORP | `CORP_CQSn`, `CORP_UNRATED` | 20/50/75/100/150/150, unrated 100 [CRE20 corporates; Art. 122] |
| QCCP | `QCCP` | 2% on trade exposures [CRE54.7; Art. 306(1)(a)] |

`ST_` applies when **every** transaction of the netting set has an original
maturity of three months or less (start date + 3 months ≥ maturity date; an open
transaction is not short-term). A risk weight override replaces the table value
and makes the run WHAT-IF.

### Status

| Status | When | E\* |
| --- | --- | --- |
| `VALID` | No input error, no rejected transaction | Calculated |
| `INVALID: n input error(s)` | A netting-set input is unusable (a worksheet error value such as `#N/A` in any cell included), a duplicate ID, a 0% haircut whose conditions fail, an issue described differently on two legs, or no risk-weight row | Withheld |
| `INCOMPLETE: k of n transaction(s) rejected` | A transaction with a leg in the netting set was rejected; a leg holding a worksheet error value is rejected with it | Withheld |
| `NO TRANSACTIONS` | No leg refers to it | — |

A layout change (moved column, broken workbook name, missing sheet), an
unusable parameter or Params table, or no usable netting set stops the run; the
problem is listed on Checks. A netting set that is not VALID has no figures on
Results and no rows on Positions.

<a id="regulatory-parameter-set"></a>

## 📌 Regulatory parameter set

`CRE22-CRR-2026-10-24` (constants `REF_*` in `src/core/CORE_Config.bas`). A run
whose parameters, haircut table, risk-weight table and overrides equal it is
**REGULATORY**; otherwise **WHAT-IF**, with each difference listed on Checks.
`DaysPerYear` is a convention and does not change the class.

| Code | Value | Source |
| --- | ---: | --- |
| `HaircutBaseBD` | 10 | CRE22.49; Art. 224 Table 1 |
| `HoldingRepoBD` | 5 | CRE22.57; Art. 224(2) |
| `HoldingCapitalMarketBD` | 10 | CRE22.57; Art. 224(2) |
| `HoldingLargeBD` | 20 | CRE22.58; Art. 224(2), 285(3) |
| `FXHaircut` | 8% | CRE22.52; Art. 224 Table 4 |
| `NonEligibleHaircut` | 30% | CRE22.54; Art. 224(4) |
| `NetWeight` / `GrossWeight` | 0.4 / 0.6 | CRE22.64-22.65; Art. 220(3) |
| `IssueCountThreshold` | 0.1 | CRE22.65; Art. 220(3) |
| `CapitalRatio` | 8% | RBC20.1; Art. 92(1)(c) |
| Haircut table | 35 rows | CRE22.49; Art. 224 Tables 1, 3 (10-day column) |
| Risk-weight table | 35 rows | CRE20; Art. 114, 120-122, 306(1)(a) |

The 10-day haircuts of the consolidated CRR (Table 1: 0.5/2/2/4/4 for CQS 1
sovereign debt, …, 15 for CQS 4 sovereign debt; Table 3: 20/30/0/20) equal
CRE22.49, so one table serves both regimes; the `Regimes` column lets a row be
limited to one of them.

<a id="support-matrix"></a>

## 🧾 Support matrix

| Supported | Entered by the user | Out of scope |
| --- | --- | --- |
| Repo, reverse repo, securities lending and borrowing, margin lending | Eligibility of the master netting agreement | Secured lending (20 BD), commodity repos |
| Netting under an eligible MNA; single-transaction formula otherwise | Settlement currency of the agreement | Cross-product netting with derivatives |
| Cash, debt (SOV / INST / CORP / SEC), equities (main index / listed), gold, funds | Fund haircut (look-through or highest) | Short-term issue ratings (A-1 …) |
| Holding period: type, >5,000 / illiquid, disputes, override; remargining frequency | The 0% haircut conditions not visible in the inputs | Own-estimate haircuts, VaR, IMM, the simple approach |
| Non-eligible securities lent (30%), non-eligible collateral not recognised | SCRA grade of an unrated institution | Haircut floors for non-centrally cleared SFTs (CRE56) |
| Risk weights SOV / INST (ECRA, SCRA, short-term) / CORP / QCCP; override | CQS of counterparties and issues | Maturity mismatch, SCRA sovereign floor, CVA, buffers, Pillar 2 |

<a id="traceability"></a>

## 🔗 Traceability

| Rule | Code | CRR | Basel |
| --- | --- | --- | --- |
| Netted exposure value | `SFT_ExposureNetted`, `CORE_Engine.FinishOneNetted` | Art. 220(3) | CRE22.64-22.65 |
| Net and gross exposure, N | `SFT_NetExposure`, `SFT_GrossExposure`, `SFT_IssueCount` | Art. 220(3) | CRE22.65 |
| Single-transaction exposure value | `SFT_ExposureSingle`, `CORE_Engine.FinishSingles` | Art. 223(2), (3), (5) | CRE22.46, 22.55 |
| Haircut scaling | `SFT_ScaledHaircut`, `CORE_Engine.ApplyHaircut` | Art. 226 | CRE22.59 |
| Holding period | `SFT_HoldingPeriod`, `CORE_Engine.HoldingFor` | Art. 224(2), 285(3)-(4) | CRE22.57-22.58 |
| Haircut key and eligibility | `CORE_Engine.FindHaircut` | Art. 197, 224(1), (4)-(6) | CRE22.49, 22.54 |
| 0% haircut | `CORE_Engine.CheckZeroHaircut` | Art. 227 | CRE22.60, 22.36-22.37 |
| Risk weight | `CORE_Engine.FindRiskWeight` | Art. 114, 120-122, 306(1)(a) | CRE20.7, 20.18-20.21, CRE54.7 |
| Own funds requirement | `SFT_CapitalRequirement` | Art. 92(1)(c) | CRE22.48, RBC20.1 |
| Run class | `CORE_Engine.ClassifyRun` | — | — |

<a id="assumptions-and-open-points"></a>

## ⚖️ Assumptions and open points

1. **Mixed netting sets.** CRE22.58 sets 10 BD for a netting set mixing
   repo-style and other capital market transactions; the CRR text does not say
   so explicitly. The workbook uses the longer minimum (10) in both regimes.
2. **One haircut per issue.** Legs of one issue in a netting set must describe it
   identically (type, issuer, CQS, maturity, listing, fund haircut, currency);
   otherwise the netting set is INVALID.
3. **Short-term risk weight** only when every transaction qualifies, a
   conservative reading for netting sets that mix maturities.
4. **Paragraph locators** for the corporate risk weights (CRE20) are given by
   section, not paragraph, pending a check against the chapter PDF.
5. **Validation class.** The expected values are ILLUSTRATIVE: reproduced by an
   independent Python implementation of this document
   (`tests/derivations/sft_reference.py`), not by published worked examples —
   the Basel Committee publishes none for CRE22 netting.

<a id="sources"></a>

## 📚 Sources

| ID | Document | Version used |
| --- | --- | --- |
| `CRR` | Regulation (EU) No 575/2013, consolidated text (CELEX 02013R0575-20260626) | 26.06.2026 |
| `CRE22` | Basel Framework, CRE22 Standardised approach: credit risk mitigation | Effective 1 January 2023 |
| `CRE20` | Basel Framework, CRE20 Standardised approach: individual exposures | Effective 1 January 2023; not in the [source register](README.md#source-register) and not copied |
| `CRE54` | Basel Framework, CRE54 Capital requirements for bank exposures to CCPs | Effective 1 January 2023 |
| `CRE56` | Basel Framework, CRE56 Minimum haircut floors for SFTs | Effective 1 January 2023, the registered copy |

`CRR`, `CRE22`, `CRE54` and `CRE56` are registered, with copies and hashes in
[sources/](sources/README.md). `CRE20` is cited for the risk weights but is
not registered yet; the prototype's formulas and its risk-weight table are
not yet checked against the registered copies, and its haircut table only in
#197 (see the departures above).
