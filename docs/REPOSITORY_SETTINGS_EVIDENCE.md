# Repository settings evidence snapshot

This file preserves the dated repository-specific observations for **K-SFTs**.
The portable non-default settings reference is [REPOSITORY_SETTINGS.md](REPOSITORY_SETTINGS.md).
This snapshot is not a current certification and must not be copied as evidence
for another repository.
[Governance](GOVERNANCE.md) defines the working policy;
[SECURITY.md](../SECURITY.md) defines vulnerability reporting and data handling.
GitHub's live settings remain authoritative. This file records observations,
not automatic enforcement or authorization to change a setting.

Snapshot date: **2026-10-29 (Europe/Rome)**. Tracking: [#162](https://github.com/danielep71/K-SFTs/issues/162).

## Reading the register

- **Verified**: read from the live GitHub API or Settings UI on the stated date.
- **Run evidence**: a particular workflow completed; this is not proof of future coverage.
- **Historical**: retained earlier evidence, not a fresh confirmation.
- **Unverified**: no conclusive current read-back. Never interpret this as disabled.
- **Desired**: existing policy or a proposed baseline explicitly marked as such.

K-SFTs is a reference baseline, not a certification of ideal settings. No
settings were changed to prepare this document. An enabled scanner is not proof
that its findings block merges, and static checks are not Excel or numerical validation.

## General settings

Evidence: [repository API](https://api.github.com/repos/danielep71/K-SFTs), read 2026-10-29.
All rows below were verified on that date.

| Setting | Desired baseline | Observed setting | Exceptions / interpretation |
| --- | --- | --- | --- |
| Visibility / default branch | Public / main | Public / main | Development targets release/0.2.0, not main. |
| Issues / Wiki | Enabled | Enabled / enabled | Wiki editing access was not reverified. |
| Projects / Discussions | Retain current choice | Disabled / disabled | Repository feature flags; not an inventory of user-level projects. |
| Merge methods | Subject to branch rules | Merge, squash, rebase enabled | Protected branches restrict the methods below. |
| Auto-merge | Disabled | Disabled | Passing checks do not authorize integration into main. |
| Delete merged head branches | Enabled | Enabled | Release and traffic branches remain protected by rulesets. |
| Update-branch suggestions | Enabled | Enabled | Governance prohibits using this on protected integration branches. |
| Web commit sign-off | No change proposed | Not required | Separate from reviews and signed-commit requirements. |
| Pages / homepage | Enabled / project URL | Pages enabled; https://danielep71.github.io/K-SFTs/ | Deployment source and environment controls not freshly verified. |
| Topics | Relevant discovery terms | counterparty-credit-risk, excel, risk-management, securities-financing-transactions, vba | Topic count is not a security measure. |

## Branch and tag rulesets

All five rulesets were read in full through the API on **2026-10-29**:
active enforcement, empty bypass lists, and no ref exclusions. This confirms
configuration, not destructive enforcement tests.

| Rule / target | Desired and observed setting | Evidence | Exceptions |
| --- | --- | --- | --- |
| Protect main: default branch | PR required; merge commits only; 0 approvals; conversations resolved; strict up-to-date Repository integrity check from GitHub Actions (integration 15368); deletion and force pushes blocked | [25772744](https://github.com/danielep71/K-SFTs/rules/25772744) | No administrator bypass. Check enforced on creation. |
| Protect release branches: release/** | PR required; squash or merge; 0 approvals; conversations resolved; same required check and strict up-to-date rule; deletion and force pushes blocked | [25772772](https://github.com/danielep71/K-SFTs/rules/25772772) | Check not enforced on initial branch creation. No bypass. |
| Prevent release branch deletion: release/** | Deletion blocked independently | [25772781](https://github.com/danielep71/K-SFTs/rules/25772781) | No bypass. |
| Protect release tags: v* | Updates, deletion and force pushes blocked | [25772795](https://github.com/danielep71/K-SFTs/rules/25772795) | Creation is not blocked by this ruleset. |
| Protect traffic history: traffic-history | Deletion and force pushes blocked | [25775687](https://github.com/danielep71/K-SFTs/rules/25775687) | Direct automation writes allowed; no PR or CI requirement. |

Both PR rules disable stale-review dismissal, code-owner review and last-push
approval; no required reviewers are configured. Zero approvals accommodates
the sole-owner model; it does not establish independent approval. The rulesets
require only Repository integrity, not a separate CodeQL check or code-scanning rule.

Release synchronization uses a **main-to-release PR with a merge commit**,
not a direct fast-forward or an administrator exception. Follow the ordering
and divergence rules in [Integration](GOVERNANCE.md#-integration).
Historical enforcement demonstrations remain in
[Governance](GOVERNANCE.md#demonstrated-enforcement).

## Security settings

Evidence for UI rows: [Advanced Security settings](https://github.com/danielep71/K-SFTs/settings/security_analysis),
read **2026-10-29**. These observations supersede the disabled-feature snapshot
of 2026-10-25; the exact intervening activation time was not established.

| Setting | Desired baseline | Observed state / evidence | Exception or remaining verification |
| --- | --- | --- | --- |
| Private vulnerability reporting | Enabled | Verified enabled, UI | Reporting procedure remains in SECURITY.md. |
| Dependency graph | Enabled | Verified enabled, UI | Does not prove all dependencies are captured. |
| Automatic dependency submission | Retain current choice | Verified disabled, UI | Reassess when build ecosystems change. |
| Dependabot alerts | Enabled | Verified enabled, UI | Alert inventory was not audited here. |
| Dependabot security updates | Enabled | Verified enabled, UI | Distinct from scheduled version updates. |
| Grouped security updates | Retain current choice | Verified disabled, UI | Does not disable individual security updates. |
| Malware alerts | No change proposed | Verified disabled, UI | Optional control to evaluate separately. |
| Secret Protection | Enabled | Verified enabled, UI | No secret-injection test performed. |
| Push protection | Enabled | Verified enabled, UI | No blocked-push demonstration performed. |
| CodeQL | Configured; validate coverage | [Successful CodeQL Setup run](https://github.com/danielep71/K-SFTs/actions/runs/40837774975), 2026-10-29 | Run evidence only: language, trigger, branch coverage and merge blocking need confirmation. No claim of VBA analysis. |
| Code scanning thresholds | Retain current choice | Verified UI: security high or higher; standard errors only | A failure threshold alone does not make the scan a required merge gate. |
| Dependabot auto-triage | Review before standardizing | UI reports one enabled rule | Rule identity and dismissal behavior unverified; do not assume low-impact alerts stay visible. |
| AI scans / agentic autofix | No change proposed | Verified disabled, UI | Preview AI repository, PR and AI-application scans off; agentic autofix off. |
| Copilot Autofix | Retain current choice | Verified enabled, UI | Availability of suggestions does not prove fixes are safe or applied. |

## Actions and dependency automation

Repository-level Actions settings are **not fully verified**: the browser
comparison timed out. Workflow YAML is evidence of that workflow only, not of
the repository-wide defaults. Do not copy historical settings as fresh facts.

| Control | Desired baseline | Observed state / date | Evidence / remaining work |
| --- | --- | --- | --- |
| Full-SHA requirement at repository level | Proposed baseline: enforce | Unverified | Read Actions settings; distinguish policy from pins in YAML. |
| Default GITHUB_TOKEN permissions | Proposed baseline: read-only | Unverified | Verify repository default; jobs may explicitly request more permissions. |
| Actions may create/approve PRs | Proposed baseline: disabled | Unverified | Verify repository toggle. |
| Allowed Actions and reusable workflows | Explicit reviewed policy | Unverified | Record allowed publishers/patterns before declaring a standard. |
| External-contributor workflow approval | Explicit reviewed policy | Unverified | Record the selected contributor category. |
| Repository log/artifact retention | Explicit documented period | Unverified | Workflow artifacts use 30 days per tools/README.md; repository retention needs read-back. |
| Version updates | Weekly grouped GitHub Actions updates | File inspected 2026-10-29: Monday 05:00 Europe/Rome; target release/0.2.0; limit 5; ci/P3; assignee danielep71 | [.github/dependabot.yml](../.github/dependabot.yml); main's definition controls activation. |
| Integrity / labels / issue triage | Preserve installed gates and read-only audits | Historical configuration below; implementation described in tools/README.md | [Tooling](../tools/README.md); this documentation change does not audit all workflow executions. |
| Traffic analytics environment | main-only access with least-privilege credential | Historical maintainer read-back below | Reverify environment restrictions and credential scope without recording secret values. |

## Open verification and maintenance

Before promoting this baseline as fully verified, complete the Actions rows,
CodeQL coverage and gating, auto-triage rule identity, Pages deployment controls,
and analytics environment read-back. Collaborators, installed-app permissions,
webhooks, deploy keys and other environments are outside this snapshot; no
claim of parity across repositories is made for them.

After a settings change, update its observed value, date, evidence and exception
in this file through the normal documentation workflow. Keep desired policy
separate from observations; mark unavailable checks unverified. Link changes
to their issue/PR and preserve useful dated evidence. Do not reproduce current
settings tables in governance or security policy. A new benchmark version
requires explicit maintainer adoption, not just a green CI result.

## Historical configuration record

Historical snapshot, superseded by the dated rows above where they differ.
Do not treat pending items or disabled features below as current facts.

Recorded on **2026-10-23**, updated on **2026-10-24**. *Verified* means read
back from GitHub; *pending* means not yet configured or not yet verified. A
verified setting is not by itself demonstrated enforcement: what GitHub was
seen to enforce on a real pull request is listed under
[Demonstrated enforcement](GOVERNANCE.md#demonstrated-enforcement).

| Control | Status |
| --- | --- |
| Repository visibility | Public since 2026-10-25, by the maintainer's decision ([#123](https://github.com/danielep71/K-SFTs/issues/123)); read back from the API. It was private until then. |
| Account plan | GitHub Pro (maintainer-reported). Rulesets and branch protection apply on public repositories under any plan. |
| `main` and `release/0.1.0` branches | Verified on 2026-10-23: both exist; `release/0.1.0` was created from `main` at `85b2b3b`. |
| Rulesets for `main` and `release/**` | Verified on 2026-10-24 (read back through the API, [#12](https://github.com/danielep71/K-SFTs/issues/12)): active, no bypass. Both block deletion and force pushes and require a pull request with 0 approvals, resolved review threads and a passing, up-to-date `Repository integrity`. A separate ruleset also blocks deleting `release/**`, and another protects `v*` tags from deletion, force pushes and updates. Accepted exception: a new `release/**` branch may be created before its check runs. Required check, up-to-date and merge-method rules demonstrated on `release/**` ([#12](https://github.com/danielep71/K-SFTs/issues/12), [below](GOVERNANCE.md#demonstrated-enforcement)); the other rules are verified only. |
| `Repository integrity` check | Configured ([#5](https://github.com/danielep71/K-SFTs/issues/5)): runs `tools/check.py --ci` on every pull request and push to `main` and `release/**`, without a path filter. Verified on 2026-10-24 as the required check of both rulesets. |
| Merge methods | Verified on 2026-10-24: the repository allows merge commits, squash and rebase. The rulesets allow **merge commits only** into `main`, and **squash or merge commits** into `release/**`, matching [Integration](GOVERNANCE.md#-integration). Observed: task pull requests squash-merged into `release/0.1.0` (#40, #43), the release integrated into `main` with a merge commit (#41), and the sync back with a merge commit (#44). Head branches are deleted automatically after merge. |
| Label workflows | Verified as installed: `Validate label manifest` and `Label drift fixtures` run on pull requests into `main` and `release/**` that change label files; live label reconciliation runs on pushes to `main` that change them; live drift detection runs daily. |
| Codex code review | Enabled: observed on [#14](https://github.com/danielep71/K-SFTs/pull/14), on opening and on an `@codex review` request. |
| Dependency updates | Configured for GitHub Actions ([#5](https://github.com/danielep71/K-SFTs/issues/5)): weekly grouped pull requests against the active release branch. Dependabot reads its configuration from `main`, where it has been since [#15](https://github.com/danielep71/K-SFTs/pull/15). Active: verified on 2026-10-24 by its first pull request, [#18](https://github.com/danielep71/K-SFTs/pull/18). |
| Security features | Read on 2026-10-25 ([#123](https://github.com/danielep71/K-SFTs/issues/123)): private vulnerability reporting enabled; Dependabot alerts enabled (maintainer-reported); Dependabot security updates, secret scanning, push protection and code scanning not enabled, by the maintainer's decision. Recorded in [SECURITY.md](../SECURITY.md) ([#3](https://github.com/danielep71/K-SFTs/issues/3)). |
| Issue forms and PR template | Configured ([#11](https://github.com/danielep71/K-SFTs/issues/11)): Bug report, Enhancement and Documentation forms; blank issues disabled. |
| Issue triage workflow | Configured ([#11](https://github.com/danielep71/K-SFTs/issues/11)): read-only audit of open issues on issue events and manual dispatch. Issue events run it from `main`, where it has been since [#19](https://github.com/danielep71/K-SFTs/pull/19). Active: verified on 2026-10-24 by its runs on issue events, for example run [40104872758](https://github.com/danielep71/K-SFTs/actions/runs/40104872758). |
| Traffic history | Configured ([#36](https://github.com/danielep71/K-SFTs/issues/36)): `traffic-history` created as orphan commit `8560b31`; workflow on `main` since #41, scheduled daily. Verified on 2026-10-24: its ruleset is active, targets only `traffic-history`, blocks deletion and force pushes, and has no bypass. The `analytics` environment allows only `main` and holds `TRAFFIC_TOKEN` (read back in Settings by the maintainer). The first manual collection succeeded (run 40079797041). Pending: the first scheduled run. |

