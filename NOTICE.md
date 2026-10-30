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

The repository is public since 2026-10-25, by my decision
([#123](https://github.com/danielep71/K-SFTs/issues/123)). Its visibility,
and any other distribution of its content, is decided by me.

## 🔗 Provenance register

Files adapted from another repository keep a record of where they came from,
so that their notices and later changes can be traced. Each entry names the
source repository, an immutable revision and the local adaptations. The
source repositories are owned by me and published under the
same licence, MPL-2.0.

| Files | Source | Revision | Local adaptations |
| --- | --- | --- | --- |
| `.github/workflows/traffic-history.yml`, `tools/traffic_history.py`, `docs/TRAFFIC.md` | [K-Risk-Engines](https://github.com/danielep71/K-Risk-Engines) `docs/TRAFFIC.md` (design) | `8942d070c80097762534ad6fb09e9c38b101da83` | Rewritten for this repository ([#36](https://github.com/danielep71/K-SFTs/issues/36)): merge logic moved into a tested script with validation and all-or-nothing writes; data branch created with a fixed layout; zero and unavailable days distinguished; no badges, alert issues or repository counts; token limited to the fetch step. |
| `.github/PULL_REQUEST_TEMPLATE.md`; parts of `docs/GOVERNANCE.md` (Branches, Integration, Review) | K-Risk-Engines (structure) | `8942d070c80097762534ad6fb09e9c38b101da83` | Rewritten for this repository's branches, sync procedure and checks ([#1](https://github.com/danielep71/K-SFTs/issues/1), [#11](https://github.com/danielep71/K-SFTs/issues/11)). |

The other files were written for, or adapted by me into, this
repository. Revisions are full commit SHAs on the source repository's `main`
branch. A file adapted later from
another source is added here in the change that adapts it.

## 🧩 Third-party tools

The workflows and local checks use these tools at run time. None of them is
copied into the repository; each keeps its own licence.

| Tool | Version | Licence | Use |
| --- | --- | --- | --- |
| `actions/checkout`, `actions/setup-python`, `actions/setup-node`, `actions/upload-artifact` | Pinned by full SHA in the workflows | MIT | GitHub Actions steps |
| Ruff | 0.17.1 | MIT | Python lint (`tools/requirements-quality-ci.txt`) |
| mypy | 2.4.0 | MIT | Python type check (`tools/requirements-quality-ci.txt`) |
| actionlint | 1.7.12 | MIT | Workflow check, downloaded and verified by SHA-256 |

The licences were checked on 2026-10-24 against the upstream licence files and
package metadata, and again for Ruff 0.17.1 and mypy 2.4.0 on 2026-10-29
against their package metadata.

## 🚫 Data

No third-party data, client data or real portfolio data is part of this
repository. Fixtures and examples are synthetic; see [SECURITY.md](SECURITY.md).

The copies of regulatory texts in `docs/methodology/sources/` (EU legislation,
Basel Framework chapters, EBA Q&As) are third-party documents. They are not
covered by MPL-2.0 and stay under their publishers' terms, which
[docs/methodology/sources/README.md](docs/methodology/sources/README.md)
records.

## Artwork

`docs/assets/K-SFTs-social-preview.jpg` was generated with OpenAI image generation
on 2026-10-24 at my request. My own
social-preview images supplied the visual references (navy background,
mint accents and dashboard composition). The title and footer were corrected to `K-SFTs`
on the same date to match the repository name (#117). The panels are conceptual illustrations,
not screenshots of an implemented application or evidence of validated calculations.
The generated image was exported as a 1280 x 640 JPEG for README and
social-preview use (#113). On 2026-10-30, after the v0.2.0 release, I had
its pill changed from *Research & Development* to *Release 0.2.0*: only the
pill and the `K-SFTs` label beside it were repainted, the text set in Barlow
Condensed Bold (SIL Open Font License), which is not stored here (#235).

`docs/assets/SFT_aggregation.svg` and `docs/assets/SFT_aggregation.png` were
authored for K-SFTs on 2026-10-26 at my request (#142). The SVG
contains editable text and vector shapes; the PNG is its rendered export,
re-rendered on 2026-10-30 with Arial's metrics so that no title overruns its
box (#208). The diagram is a simplified methodological illustration, not a
screenshot or evidence of validated calculations.

The companion pairs `docs/assets/SFT_CRR_uses.svg` / `.png` and
`docs/assets/SFT_calculation_reference.svg` / `.png` were authored on
2026-10-26 for the same request (#142). Both are editable vector diagrams
with PNG exports. Regulatory-use labels distinguish separate frameworks;
calculation details identify prototype conventions and limitations.
`docs/assets/SFT_calculation_reference_workbook.png` is a half-size,
256-colour copy of the calculation reference PNG, made for the workbook's
README sheet (#167).

`docs/assets/CCR_framework.svg` and `docs/assets/CCR_framework.png` were
authored on 2026-10-26 at my request (#151), and the same
pair is used in K-SACCR and K-Risk Engines. The SVG holds editable text and
vector shapes; the PNG is its rendered export, at twice the SVG's size and
reduced to 256 colours. The diagram names the sister calculators K-SA-CCR,
K-CCP and K-CVA by my decision of 2026-10-26: an explicit
exception, for this diagram, to the rule of #137 that no file names
K-SACCR. It is an architecture illustration, not evidence of an implemented
or validated calculation.
