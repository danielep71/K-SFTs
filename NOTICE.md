# ⚖️ Licensing and provenance

Copyright 2026 Daniele Penza and contributors.

This is an independent personal project.

## 📝 Licence

The repository's code, documentation, workflows and configuration are covered
by the Mozilla Public License 2.0 in [LICENSE](LICENSE), unless a file carries
a separate notice.

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.

The repository is private. Its visibility, and any distribution of its
content, is decided by the maintainer; the licence does not make it public.

## 🔗 Provenance register

Files adapted from another repository keep a record of where they came from,
so that their notices and later changes can be traced. Each entry names the
source repository, an immutable revision and the local adaptations. The
source repositories are owned by the same maintainer and published under the
same licence, MPL-2.0.

| Files | Source | Revision | Local adaptations |
| --- | --- | --- | --- |
| `.github/labels.json`, `.github/label-policy.json` | [K-SACCR](https://github.com/danielep71/VBA-SACCR-Toolkit) | `7397382a638817b1f8bf47348205de7a1928cb78` | `P1` description narrowed to confirmed calculation bugs ([#11](https://github.com/danielep71/CCR-K-SFTs/issues/11)). |
| `.github/scripts/labels-sync.mjs`, `.github/scripts/labels-drift.mjs` | K-SACCR | `00bb72d1bf5ec9252b8ca5ba3adc567d948be15d` | Default policy path set to `.github/label-policy.json`; User-Agent strings made repository-neutral. |
| `.github/workflows/labels-sync.yml` | K-SACCR | `f4b4e0f6026236885050eb00ed564ce382fc7836` | None. |
| `.github/workflows/labels-drift.yml` | K-SACCR | `6d1f11301e12bb8512b41596f59bd42287e92773` | None. |
| `.editorconfig`, `.gitattributes`, `.gitignore` | K-SACCR | `00bb72d1bf5ec9252b8ca5ba3adc567d948be15d` | Repository name; workbook exception, R and CITATION.cff rules, documentation-asset paths and assistant-specific editor entries removed ([#20](https://github.com/danielep71/CCR-K-SFTs/issues/20)); comments record the ASCII-only encoding decision, the type-only classification and the intended template exception ([#4](https://github.com/danielep71/CCR-K-SFTs/issues/4)). |
| Issue forms in `.github/ISSUE_TEMPLATE/` | K-SACCR (structure) | `00bb72d1bf5ec9252b8ca5ba3adc567d948be15d` | Rewritten for this repository: areas, default priorities, dependency field, synthetic-data notice ([#11](https://github.com/danielep71/CCR-K-SFTs/issues/11)). |
| `SECURITY.md` | K-SACCR (structure) | `00bb72d1bf5ec9252b8ca5ba3adc567d948be15d` | Rewritten for this repository's scope, reporting route and security features ([#3](https://github.com/danielep71/CCR-K-SFTs/issues/3)). |
| `pyproject.toml`; the actionlint step of `.github/workflows/repository-integrity.yml` | K-SACCR (configuration) | `00bb72d1bf5ec9252b8ca5ba3adc567d948be15d` | Python 3.12 targets; new lock for CPython 3.13 ([#5](https://github.com/danielep71/CCR-K-SFTs/issues/5)). |
| `.github/workflows/traffic-history.yml`, `tools/traffic_history.py`, `docs/TRAFFIC.md` | K-SACCR `daily-traffic.yml` and K-Risk-Engines `docs/TRAFFIC.md` (design) | K-SACCR `00bb72d1bf5ec9252b8ca5ba3adc567d948be15d`; K-Risk-Engines `8942d070c80097762534ad6fb09e9c38b101da83` | Rewritten for this repository ([#36](https://github.com/danielep71/CCR-K-SFTs/issues/36)): merge logic moved into a tested script with validation and all-or-nothing writes; data branch created with a fixed layout; zero and unavailable days distinguished; no badges, alert issues or repository counts; token limited to the fetch step. |
| `docs/REPOSITORY_STRUCTURE.md` (structure: profile decision, layout, dependency direction, public API boundary, placement rules); `docs/PUBLIC_API.txt` (inventory format) | K-SACCR (structure) | `00bb72d1bf5ec9252b8ca5ba3adc567d948be15d` | Rewritten for this repository ([#2](https://github.com/danielep71/CCR-K-SFTs/issues/2)): SFT scope and setup status, open product and methodology decisions, Python's role; no calculation content, known deviations or K-SACCR components. |
| `docs/VBA_HOUSE_STYLE.md`; the *Supported hosts*, *Get the source* and *Importing VBA into Excel* sections of `INSTALLATION.md` | K-SACCR (structure) | `00bb72d1bf5ec9252b8ca5ba3adc567d948be15d` | Rewritten for this repository ([#4](https://github.com/danielep71/CCR-K-SFTs/issues/4)): `CORE_`/`SFT_`/`TEST_` prefixes, SFT units and domains with no regulatory convention fixed, ASCII-only source instead of cp1252, a cleanup sketch instead of a reference implementation, a separate tested-environments record; no K-SACCR components, known deviations, sync script or component list. |
| `.github/PULL_REQUEST_TEMPLATE.md`; parts of `docs/GOVERNANCE.md` (Branches, Integration, Review) | [K-Risk-Engines](https://github.com/danielep71/K-Risk-Engines) (structure) | `8942d070c80097762534ad6fb09e9c38b101da83` | Rewritten for this repository's branches, sync procedure and checks ([#1](https://github.com/danielep71/CCR-K-SFTs/issues/1), [#11](https://github.com/danielep71/CCR-K-SFTs/issues/11)). |

The other files were written for this repository. Revisions are full commit
SHAs on the source repository's `main` branch. A file adapted later from
another source is added here in the change that adapts it.

## 🧩 Third-party tools

The workflows and local checks use these tools at run time. None of them is
copied into the repository; each keeps its own licence.

| Tool | Version | Licence | Use |
| --- | --- | --- | --- |
| `actions/checkout`, `actions/setup-python`, `actions/setup-node`, `actions/upload-artifact` | Pinned by full SHA in the workflows | MIT | GitHub Actions steps |
| Ruff | 0.16.6 | MIT | Python lint (`tools/requirements-quality-ci.txt`) |
| mypy | 2.3.1 | MIT | Python type check (`tools/requirements-quality-ci.txt`) |
| actionlint | 1.7.12 | MIT | Workflow check, downloaded and verified by SHA-256 |

The licences were checked on 2026-10-24 against the upstream licence files and
package metadata.

## 🚫 Data

No third-party data, client data or real portfolio data is part of this
repository. Fixtures and examples are synthetic; see [SECURITY.md](SECURITY.md).
