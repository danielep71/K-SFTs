# CCR-K-SFTs

Counterparty credit risk for securities financing transactions (SFTs):
repurchase agreements, securities lending and borrowing, and margin lending.
CCR-K-SFTs is intended to become an Excel/VBA application that calculates SFT
exposure from transactions, collateral and netting sets.

> [!IMPORTANT]
> **Status: repository setup only.** There is no calculation engine, public
> API, workbook or validated method yet. Nothing here should be used to
> compute exposure. The regulatory basis is chosen: the EU CRR controls,
> with the Basel Framework as the methodological reference. The first scope
> is bilateral SFTs under the Financial Collateral Comprehensive Method;
> products, collateral and outputs are still open decisions.

This is an independent personal project.

| Document | Purpose |
| --- | --- |
| [docs/REPOSITORY_STRUCTURE.md](docs/REPOSITORY_STRUCTURE.md) | Scope, profile, layout, dependency direction, public API boundary, open decisions |
| [docs/methodology/README.md](docs/methodology/README.md) | Regulatory basis, first scope, version applicability, source register, traceability |
| [docs/methodology/sources/README.md](docs/methodology/sources/README.md) | Copies of the regulatory texts relied on, with hashes and terms of use |
| [docs/methodology/TEST_CASES.md](docs/methodology/TEST_CASES.md) | Test-case files, tolerances, reference classes, completeness |
| [docs/VBA_HOUSE_STYLE.md](docs/VBA_HOUSE_STYLE.md) | VBA naming, contracts, errors, cleanup, encoding |
| [INSTALLATION.md](INSTALLATION.md) | Supported Excel hosts; importing and exporting the VBA source |
| [docs/GOVERNANCE.md](docs/GOVERNANCE.md) | Branches, integration, review, protection, issues, releases |
| [docs/TRAFFIC.md](docs/TRAFFIC.md) | Repository traffic history |
| [tools/README.md](tools/README.md) | Repository checks: `python tools/check.py` |
| [SECURITY.md](SECURITY.md) | Reporting vulnerabilities; synthetic data only |
| [NOTICE.md](NOTICE.md), [LICENSE](LICENSE) | Copyright, provenance and the MPL-2.0 licence |

Installation and contributor onboarding are completed in
[#10](https://github.com/danielep71/CCR-K-SFTs/issues/10).
