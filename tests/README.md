# 🧪 Tests

Verification source, synthetic fixtures and expected results. **Nothing is
here yet.** There is no calculation to test until the methodology decisions in
[docs/REPOSITORY_STRUCTURE.md](../docs/REPOSITORY_STRUCTURE.md) are taken.
The repository tooling has its own tests in [tools/](../tools/README.md).

| Location | Will hold |
| --- | --- |
| `modules/` | VBA test harness, case runner and suites ([#7](https://github.com/danielep71/CCR-K-SFTs/issues/7)) |
| `fixtures/` | Synthetic inputs: transactions, collateral, netting sets |
| `expected/` | Expected results, each with its provenance |
| `derivations/` | Python reference derivations that compute expected results from the documented method; evidence for tests, never shipped |

The fixture schema, tolerances and provenance rules are defined in
[#8](https://github.com/danielep71/CCR-K-SFTs/issues/8). Each subdirectory is
created with its first real file.

- Fixtures are **synthetic**. Never commit real transactions, counterparties,
  collateral or portfolios ([SECURITY.md](../SECURITY.md)).
- Derivative SA-CCR examples, factors or thresholds are not SFT references and
  are never reused as expected values.
- Test modules call the facade or the core directly; they never enter the
  production import set.
