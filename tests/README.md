# 🧪 Tests

Verification source, synthetic fixtures and expected results. **No real case
is here yet**: the only case, `illustrative-format-example`, shows the file
format and validates nothing. There is no calculation to test until the
methodology decisions in
[docs/REPOSITORY_STRUCTURE.md](../docs/REPOSITORY_STRUCTURE.md) are taken.
The repository tooling has its own tests in [tools/](../tools/README.md).

| Location | Will hold |
| --- | --- |
| `modules/` | VBA test modules: the harness `TEST_Harness` and the neutral example suite `TEST_Examples`; the case runner and the SFT suites come with the calculation |
| `fixtures/` | Synthetic inputs: transactions, collateral, netting sets |
| `expected/` | Expected results, each with its provenance |
| `derivations/` | Python reference derivations that compute expected results from the documented method; evidence for tests, never shipped |

The fixture schema, tolerances, reference classes and completeness rules are
defined in [docs/methodology/TEST_CASES.md](../docs/methodology/TEST_CASES.md)
and checked by `python tools/check.py` (gate `test-cases`). Each subdirectory
is created with its first real file.

- Fixtures are **synthetic**. Never commit real transactions, counterparties,
  collateral or portfolios ([SECURITY.md](../SECURITY.md)).
- Derivative SA-CCR examples, factors or thresholds are not SFT references and
  are never reused as expected values.
- Test modules call the facade or the core directly; they never enter the
  production import set.

<a id="vba-harness"></a>

## 🧰 VBA harness

`TEST_Harness` runs one suite at a time. A suite calls `BeginRun` with its
name, its mode (`NORMAL` or a deliberate failure mode) and its expected case
and assertion counts, `BeginCase` with a lower-case kebab-case identifier for
each case, the assertions, and `EndRun` with the error that escaped its
runner (0 and empty strings for none). Each case has its own error handler that calls `RecordCaseError`,
so that one failing case does not stop the others.

| Assertion | Passes when |
| --- | --- |
| `AssertEqualDouble` | `Actual = Expected` exactly; only for results exact in binary |
| `AssertNear` | The difference is at most `max(absolute, relative × abs(Expected))`, both tolerances stated, as in [TEST_CASES.md](../docs/methodology/TEST_CASES.md#tolerance) |
| `AssertEqualText` | The texts are equal, compared binary |
| `AssertTrue` | The condition the suite evaluated is true |
| `AssertErrorValue` | The value is the given worksheet error, such as `#VALUE!` |
| `AssertRaised` | The `Err.Number` captured after the call is the expected one (0: nothing raised) |

A run passes only when every assertion passed, no unexpected error occurred,
exactly the expected cases and assertions ran, and Excel's `Calculation`,
`DisplayAlerts`, `EnableEvents` and `ScreenUpdating` are as they were at the
start. A skipped or extra case or assertion makes the run `INCOMPLETE`,
which is a failure. The harness never changes those settings itself; a
procedure that does restores them, as `RunExampleOperation` in
`TEST_Examples` shows ([cleanup pattern](../docs/VBA_HOUSE_STYLE.md#operation-and-cleanup)).
How to run it is in [INSTALLATION.md](../INSTALLATION.md#running-the-harness).

<a id="log-format"></a>

### Log format

One record per line, `KEY=value`, with further fields separated by `; `.
`RUN` and `ENVIRONMENT` come first; `CASE`, `ASSERT`, `ERROR` and `FAILURE`
follow as the run reaches them; the summary records come last, in the order
of the table:

| Record | Fields |
| --- | --- |
| `RUN=<suite>` | `mode` (`NORMAL` or a deliberate failure mode), `expected_cases`, `expected_assertions`, `started` |
| `ENVIRONMENT=` | `host`, `version`, `build` (Excel's build number), `os`, `office` (`64-bit` or `32-bit`), `runtime` (`VBA7` or `VBA6`) |
| `CASE=<case-id>` | One per case, in run order |
| `ASSERT=<case-id>/<name>` | `status` (`PASS` or `FAIL`); on failure, the assertion's detail: the expected and actual values (and the tolerances, for `AssertNear`), or the suite's text for `AssertTrue` |
| `FAILURE=<id>` | One per failure: an assertion, `<case-id>.unexpected-error`, `runner.unexpected-error`, `harness.case-id`, `harness.no-case` or `harness.completeness`, with its detail |
| `ERROR=<case-id>` or `ERROR=run` | `number`, `source`, `description` of an unexpected error |
| `CASES=<n>`, `ASSERTIONS=<n>` | `expected` |
| `FAILURES=<n>` | `errors` |
| `PRIMARY_FAILURE=<id>` | The first failure of the run, when there is one |
| `CLEANUP=PASS` or `FAIL` | `detail`: the settings compared, those that changed, or why they could not be compared |
| `RECOVERY=NONE` or `REQUIRED` | `reason`, `action`, when required |
| `RESULT=PASS` or `FAIL` | `suite`; `operation`, what ran: `PASS`, `FAIL` (an assertion or a case identifier failed) or `ERROR` (an unexpected error); `completeness` (`COMPLETE` or `INCOMPLETE`); `cases`, `assertions`, `failures`; `cleanup`; `recovery` |

`RESULT` is the last record of every run that started. A refused start
writes a `RESULT=FAIL` with `operation=NOT_RUN` and `cleanup=NOT_RUN`,
followed by a `RECOVERY=REQUIRED` record, both appended to the log kept from
the previous run; `ResetTests` writes `RESET=DONE`. Case identifiers and assertion names are
fixed in the suite's source, so two logs of the same commit can be compared
line by line, apart from `started` and `ENVIRONMENT`.
