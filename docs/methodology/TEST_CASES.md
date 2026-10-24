# 🧪 Numerical test cases

This document defines how CCR-K-SFTs's numerical test cases are written:
files, inputs and their units, expected results and their provenance,
tolerances, case categories, invalid inputs and when a rule counts as
validated. Sources and rule IDs come from the [methodology](README.md). Its
structure is adapted from K-SACCR's `docs/methodology/TEST_CASES.md` (see
[NOTICE.md](../../NOTICE.md)); no K-SACCR case, value or threshold is reused.

> [!IMPORTANT]
> **Status on 2026-10-24: format only.** The one case,
> `illustrative-format-example`, shows the format and holds no regulatory
> value. The transaction, collateral and output fields of real cases are
> fixed with the input and output contract (`docs/contracts/`, not yet
> written) once D4, D6 and D7 are decided.
> `python tools/check_test_cases.py` enforces everything below that a file
> can show; it never runs a calculation.

## 📁 Files

A case is one input file and one expected-result file per calculation regime:

| File | Holds |
| --- | --- |
| `tests/fixtures/<case-id>.json` | Synthetic inputs: netting sets, transactions, collateral |
| `tests/expected/<case-id>.<regime>.json` | Expected results for one regime, each with its provenance. The only regime is `crr` ([version applicability](README.md#version-applicability)) |
| `tests/derivations/<name>.py` | Optional: a script that computes expected values from the documented method, separately from the code under test |

- Case IDs are lowercase kebab-case, equal to the fixture's file name and
  never reused.
- A case whose expected values are all `illustrative` starts with
  `illustrative-`, and only such a case may.
- Every fixture has at least one expected file, and every expected file a
  fixture.
- Files are UTF-8 JSON, stored with LF. All data is **synthetic**: never real
  transactions, counterparties, collateral or portfolio extracts
  ([SECURITY.md](../../SECURITY.md)).

<a id="fixture-format"></a>

## 📥 Fixture format

```json
{
  "schema_version": 1,
  "kind": "sft-fixture",
  "id": "illustrative-format-example",
  "synthetic": true,
  "description": "What the case exercises, in one sentence",
  "category": "nominal",
  "calculation_currency_code": "EUR",
  "valuation_date": "2026-01-02",
  "netting_sets": [
    {
      "id": "NS1",
      "counterparty_id": "CP1",
      "inputs": {"master_netting_agreement_flag": true},
      "transactions": [
        {
          "id": "T1",
          "product": "repo",
          "inputs": {
            "currency_code": "EUR",
            "cash_amount": 1000000,
            "start_date": "2026-01-02",
            "maturity_date": "2026-04-02"
          }
        }
      ],
      "collateral": [
        {
          "id": "C1",
          "transaction_id": "T1",
          "inputs": {"currency_code": "EUR", "market_value_amount": 1020000}
        }
      ]
    }
  ]
}
```

| Rule | Detail |
| --- | --- |
| Envelope | `schema_version` (1), `kind` (`sft-fixture`), `id`, `synthetic` (`true`), `description`, `category`, `calculation_currency_code` and `valuation_date` are required, and no other key is allowed |
| Identifiers | Each netting set has an `id`, unique in the fixture, and a `counterparty_id`. Each transaction and each collateral item has an `id`, unique in the fixture. A collateral item's `transaction_id` names a transaction of the same netting set, or is `null` for collateral held at netting-set level |
| Products | `product` is one of `repo`, `reverse_repo`, `securities_lending`, `securities_borrowing` or `margin_lending`, the products of D4. Naming one does not put it in scope |
| Inputs | Every quantity sits in an `inputs` object, whose keys are lower snake case and end in a unit suffix (below). Which keys exist is decided by the input contract; until it is written, only `illustrative-` cases may exist |
| Currencies | `calculation_currency_code` is the currency of results. An object whose `inputs` hold an `_amount` key also holds `currency_code`, the currency of all its amounts, so a currency mismatch is always visible |
| Signs | Amounts are unsigned unless the input contract defines a signed field; direction comes from `product` and the contract, never from a sign convention invented in a case |

**Unit suffixes.** A key without one of these suffixes is rejected:

| Suffix | Value |
| --- | --- |
| `_amount` | Number in the object's `currency_code`, unscaled |
| `_rate`, `_factor` | Decimal number: `0.05`, never `5` meaning 5% |
| `_years` | Number of years |
| `_business_days`, `_count` | Whole number, at least 0 |
| `_date` | Text, ISO `YYYY-MM-DD` |
| `_flag` | `true` or `false` |
| `_code` | Text code; one ending `currency_code` is three upper-case letters (ISO 4217) |
| `_id` | Text identifier |

`null` means not applicable. Numbers are finite; JSON cannot hold NaN or
infinity, so such inputs are tested in a VBA test module
([#7](https://github.com/danielep71/CCR-K-SFTs/issues/7)).

<a id="expected-format"></a>

## 📤 Expected-result format

```json
{
  "schema_version": 1,
  "kind": "sft-expected",
  "fixture": "illustrative-format-example",
  "regime": "CRR",
  "outputs": [
    {
      "quantity": "illustrative_total_cash_amount",
      "netting_set": "NS1",
      "transaction": null,
      "rule": null,
      "value": 1000000,
      "unit": "EUR",
      "tolerance": {"absolute": 0.000001, "relative": 0.000000001},
      "reference": {
        "class": "illustrative",
        "source": null,
        "locator": null,
        "derivation": "Why the value exists and how it was obtained",
        "script": null,
        "derived_by": "Name",
        "reviewed_by": null,
        "date": "2026-10-24"
      }
    }
  ]
}
```

| Rule | Detail |
| --- | --- |
| Envelope | `schema_version` (1), `kind` (`sft-expected`), `fixture` (an existing case ID), `regime` (a regime of the [version applicability](README.md#version-applicability) table, `CRR`, matching the lower-case file-name suffix) and a non-empty `outputs` list are required |
| Calculation date | The fixture's `valuation_date` selects the version of the controlling text: it must fall within the dates of a row of the version applicability table for the file's regime. A case dated before 1 January 2025 is rejected, because the pre-CRR3 text is not registered |
| Output scope | `netting_set` names a netting set of the fixture; `transaction`, when not `null`, names a transaction of that netting set. A netting-set output has `transaction: null` |
| Quantity | Lower snake case. The quantity names of real cases are fixed with the output contract (D7); an `illustrative-` case may use a name that starts with `illustrative_` for a format-only value |
| Number | `value`, `unit` and `tolerance`. `unit` is a currency code for an amount, `1` for a rate or factor, `years` or `business_days` |
| Text | `text`, compared exactly, with no `unit` or `tolerance`. A rejection is a text output whose quantity ends in `_status` |
| Rule | A rule ID under the [traceability](README.md#traceability) convention, `CRR.<article>[.<paragraph>]`. A Basel paragraph is never a rule ID: Basel is a methodological reference, not a calculation regime. It may be `null` only for an `illustrative` value |
| Reference | One per output; see [reference classes](#reference-classes) |

<a id="reference-classes"></a>

## 🏷️ Reference classes

Every expected value carries exactly one class. Only the first two are
numerical-validation references.

| Class | Meaning | Required fields |
| --- | --- | --- |
| `published` | Printed in a registered source, for example a worked example | `source` (a register ID), exact `locator` (paragraph, example, table) |
| `independent` | Computed by an external implementation by another author, registered as a source with an ID starting `IMPL-` | `source` (`IMPL-…`), `locator`, `derivation`, `derived_by` (the implementation and its version) |
| `illustrative` | Shows the format or a behaviour; **not** a validated SFT figure | `derivation` saying why it exists and how the value was obtained |

- A `published` value never cites an `IMPL-` source, and an `independent`
  value always does. A value derived within the project, even from a
  registered text, stays `illustrative`.
- `script`, when not `null`, names a file under `tests/derivations/` that
  computed the value; it must exist.
- `reviewed_by` may be `null`: the code under test and its expected values
  come from the same maintainer, and no independent reviewer is required.
- **Examples are never benchmarks.** Derivative SA-CCR examples or thresholds,
  from K-SACCR or elsewhere, are not SFT references; the register holds no
  such source, so no case can cite one.
- Test output reports `illustrative` results separately and never as
  validation.

<a id="tolerance"></a>

## 📏 Tolerance

An actual result `a` matches the expected value `e` when

```text
|a − e| ≤ max(absolute, relative × |e|)
```

Both bounds are stated on every numeric output, as finite numbers of at least
0; there is no hidden default. Both 0 means an exact comparison, for whole
numbers only.

| Reference | `absolute` | `relative` |
| --- | --- | --- |
| `independent` or `illustrative`, full precision | `1e-6` currency units (`1e-9` for a rate or factor) | `1e-9` |
| `published`, rounded | Half a unit of the last printed digit: printed in thousands → `500` | `0` |

- **Near zero** the absolute bound governs: as `|e|` shrinks, the relative
  term vanishes, so an expected `0` is matched within `absolute` only. An
  `absolute` of 0 with an expected 0 therefore demands an exact zero.
- **Domain rules apply before tolerance:** a quantity that must not be
  negative, such as an exposure value, fails if it is negative, however
  small.
- Results are compared unrounded; the engine does not round intermediate
  values (D8).

<a id="categories"></a>

## 🗂️ Case categories

Each fixture declares one `category`:

| Category | Covers |
| --- | --- |
| `nominal` | Typical inputs within the domain |
| `boundary` | Values on a domain edge: zero amounts, results that must be exactly zero, a haircut or holding period at a regulatory floor, a single transaction, a single security issue |
| `invalid` | Inputs that must be rejected: missing or unknown fields, negative amounts where not allowed, a maturity before the start, an unknown product or currency |

<a id="invalid-inputs"></a>

### Invalid inputs

- An `invalid` case is expected to be **rejected**, never calculated: each of
  its expected files holds at least one `_status` text output, and no numeric
  output for a transaction or netting set it rejects.
- In an `invalid` case the checker still enforces the structure (envelope,
  identifiers, keys, unit suffixes, links between items), but not the domain
  of the values: an unknown product, a value of the wrong type for its suffix,
  a malformed currency code or an amount without `currency_code` are allowed,
  because they are what the case tests. In every other category they are
  rejected.
- The status texts are fixed with the output contract. A rejection never
  substitutes a default value
  ([VBA_HOUSE_STYLE.md](../VBA_HOUSE_STYLE.md#errors)).
- Inputs JSON cannot express (non-finite numbers, wrong VBA types) are tested
  in a VBA test module instead of a fixture.

<a id="completeness-policy"></a>

## ✅ Completeness policy

- A rule in the [traceability](README.md#traceability) table is **validated**
  only when at least one `published` or `independent` case passes for it.
  `illustrative` cases never count.
- Each implemented rule also needs at least one `boundary` case, and each
  public input domain at least one `invalid` case, as a fixture or a VBA test.
- A run of the test harness that does not execute every expected case and
  assertion is **incomplete, not a pass**
  ([#7](https://github.com/danielep71/CCR-K-SFTs/issues/7)).
- A release states which rules are validated, under which version of the
  controlling text, and lists
  every rule that is implemented but not validated.

<a id="checking"></a>

## 🔎 Checking

`tools/check_test_cases.py` runs in `python tools/check.py` and checks the
files under `tests/fixtures/` and `tests/expected/`: envelopes, identifiers
and references between them, products, unit suffixes and value types,
currencies, output forms and units, rule IDs, the regime and the
calculation date against the [version applicability](README.md#version-applicability)
table, reference classes and their fields, sources against the
[register](README.md#source-register), scripts,
tolerances, the `illustrative-` naming rule and that each expected file of an
`invalid` case holds a `_status` text output.

It does not check:

- whether an input key or a quantity name belongs to the input and output
  contract, which does not exist yet;
- that an `invalid` case has no numeric output for what it rejects;
- any value: nothing is computed or compared. Running cases in Excel belongs
  to the regression harness
  ([#7](https://github.com/danielep71/CCR-K-SFTs/issues/7)).
