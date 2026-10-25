# 🏛️ Repository governance

This document defines how changes to K-SFTs are proposed, reviewed and
integrated. It separates the **intended policy** from the **controls verified
as configured**; a rule written here is not evidence that GitHub enforces it.

## 👤 Ownership and scope

**Daniele Penza (@danielep71)** maintains K-SFTs and decides scope,
priorities, branch strategy, integration and publication. Contributors and
automation work within the scope authorized by the maintainer.

This is an independent personal project.

K-SFTs covers counterparty credit risk for securities financing
transactions. The current milestone, **v0.1.0 - Repo setup**, prepares the
repository infrastructure only. Production SFT calculations, numerical
certification and workbook packaging belong to later milestones; the project
boundary and architecture are defined in
[REPOSITORY_STRUCTURE.md](REPOSITORY_STRUCTURE.md) ([#2](https://github.com/danielep71/K-SFTs/issues/2)), and the
VBA conventions, supported hosts and import steps in
[VBA_HOUSE_STYLE.md](VBA_HOUSE_STYLE.md) and
[INSTALLATION.md](../INSTALLATION.md) ([#4](https://github.com/danielep71/K-SFTs/issues/4)).

Commits may be authored by the maintainer or by bots acting within authorized
repository workflows. Keep author attribution truthful: a bot-authored change
is not recreated under the maintainer's name only to change its authorship.
Bot-authored pull requests remain subject to the applicable checks and
maintainer review; bot authorship does not authorize automatic merging or
bypass branch rules.

## 🌿 Branches

| Branch | Role |
| --- | --- |
| `main` | Integrated source. Receives the active release branch through a merge commit after maintainer approval. No task branches or task pull requests target it. |
| `release/0.2.0` | **Active integration branch** for v0.2.0 - Prototype setup. Base for task branches and target of task and sync pull requests. Created from `main` at `3fc2c97` on 2026-10-25 by maintainer decision. |
| `release/0.1.0` | Integration branch of v0.1.0 - Repo setup, whose milestone was certified at `576bb58` and closed on 2026-10-25 ([#13](https://github.com/danielep71/K-SFTs/issues/13)). No longer active: no task branch is created from it and no pull request targets it. Created from `main` at `85b2b3b` on 2026-10-23. |
| Other `release/X.Y.Z` | Opened only by maintainer decision, without a `v` prefix. Do not infer a new target from a version number or from this document. |
| `fix/<issue>-<slug>`, `docs/<slug>`, `test/<slug>` | Short-lived task branches, created from the active release branch and deleted after merge. |
| `traffic-history` | Data-only branch written by the Traffic history workflow ([docs/TRAFFIC.md](TRAFFIC.md)). Unrelated to `main`; never merged, never a task base, never deleted. |

A release branch is an integration branch: its name does not imply that a
release will be published (see [Releases](#-releases)).

## 🔀 Integration

1. Create a task branch from the active release branch, never from `main` or
   GitHub's default branch by assumption.
2. Open a pull request against the active release branch. Link the issue it
   resolves and state what was validated and what was not.
3. Squash-merge once the required checks pass on the current pull request
   head and every review finding is fixed or answered. The squash commit
   subject ends with the pull request number, for example
   `docs: define repository governance (#<PR>)`, and its body ends with the
   issue reference.
4. Integrate the release branch into `main` through a pull request merged with
   a merge commit, only after maintainer approval.
5. After the integration commit passes its checks on `main`, bring the
   release branch level with `main` through a sync pull request from `main`
   into the release branch, merged with a merge commit once its checks pass.
   A sync pull request whose diff is empty needs CI only, no Codex review
   (see [Review](#-review)). No task pull request merges into the release
   branch between the integration and its sync: the release branch would
   gain a commit that `main` lacks, and the up-to-date rule would then refuse
   both the sync and the next integration. Task pull requests wait, or merge
   before the integration pull request is opened. If one merges in that
   window anyway, merge `main` into a branch cut from the release branch and
   merge that branch into the release branch with a merge commit, through a
   pull request whose diff is empty (CI only). Do not relax a rule to get
   past it ([#50](https://github.com/danielep71/K-SFTs/pull/50)).
6. `main` receives commits only through these integration pull requests. If
   it receives any other commit while the release branch has commits of its
   own, the sync pull request cannot meet the up-to-date rule; stop and
   record a maintainer decision before integrating further.

`main` and the release branches change only through pull requests, as the
[protection policy](#-protection-policy) requires. The release branch is
therefore not fast-forwarded by a direct push, and GitHub's *Update branch*
action, which pushes to the head branch, is not used when that branch is
`main` or a release branch. After an integration, the sync merge commit
changes no files: it gives the release branch the integration commit, so the
next release-to-`main` pull request meets the up-to-date rule.

Never rebase, amend or force-push a commit that has been pushed to a shared
branch. History is corrected with new commits.

## 🔎 Review

The maintainer is the sole owner and reviewer. GitHub does not allow an author
to approve their own pull request, so the policy requires a pull request but
no approving review. Review evidence consists of:

- **Maintainer review** of the final diff and head commit;
- **Codex code review**, which runs on pull requests opened by people;
- **CI**, the required checks on the current pull request revision.

These are separate forms of evidence: a green check is not a review, and a
clean review of an older revision does not cover later pushes.

Codex reviews a pull request when it is opened or marked ready, not after
later pushes. Before merging a pull request whose head changed after review:

1. Record the current head commit SHA.
2. Once any earlier review has finished, post `@codex review` as a new
   top-level conversation comment.
3. Wait until the review shows **Completed** for that head SHA.
4. Fix or explicitly answer every finding and resolve the applicable
   conversations. A further push requires this check again.
5. Record the reviewed SHA and the result (for example *no findings*, or
   the findings and how each was settled) in the pull request as text. A
   link is only a pointer: the Codex summary is one comment edited in
   place, so it later describes another commit or review.

Codex does not review bot-authored pull requests; for those, the maintainer's
review of the final diff and head SHA is the only review. Do not claim Codex
coverage when it was skipped. If review fails or is unavailable, leave the
pull request open and record the blocker.

**Sync pull requests.** A sync pull request from `main` into the release
branch (Integration, step 5) whose diff is empty changes no file: it only adds
the integration commit, whose content was reviewed on its own pull requests.
It needs CI only: `Repository integrity` must pass on its head, and no Codex
review is required. Before merging, confirm that the diff is empty (`git diff`
between the release branch and `main` shows nothing) and record that and the
CI result in the pull request. A sync pull request that changes any file is
reviewed like any other pull request, Codex included.

## 🔐 Protection policy

### Intended rules

Applies to `main` and every `release/**` branch:

| Rule | Policy |
| --- | --- |
| Pull request required | Yes; 0 approving reviews (sole owner). |
| Required status check | `Repository integrity`, the job of `.github/workflows/repository-integrity.yml` ([#5](https://github.com/danielep71/K-SFTs/issues/5)). |
| Branch up to date before merge | Required. |
| Conversations resolved before merge | Required. |
| Force pushes | Blocked. |
| Deletion | Blocked. |
| Bypass | None. |

`traffic-history` is written directly by automation and has its own rules:
deletion and force pushes blocked, no pull request or status check required,
no bypass ([docs/TRAFFIC.md](TRAFFIC.md), *Branch protection*).

A required check must report on every pull request into a protected branch.
Path-filtered workflows, such as the label workflows, must not be the only
required checks, or unrelated pull requests cannot merge.

### Configured controls

Recorded on **2026-10-23**, updated on **2026-10-24**. *Verified* means read
back from GitHub; *pending* means not yet configured or not yet verified. A
verified setting is not by itself demonstrated enforcement: what GitHub was
seen to enforce on a real pull request is listed under
[Demonstrated enforcement](#demonstrated-enforcement).

| Control | Status |
| --- | --- |
| Repository visibility | Public since 2026-10-25, by the maintainer's decision ([#123](https://github.com/danielep71/K-SFTs/issues/123)); read back from the API. It was private until then. |
| Account plan | GitHub Pro (maintainer-reported). Rulesets and branch protection apply on public repositories under any plan. |
| `main` and `release/0.1.0` branches | Verified on 2026-10-23: both exist; `release/0.1.0` was created from `main` at `85b2b3b`. |
| Rulesets for `main` and `release/**` | Verified on 2026-10-24 (read back through the API, [#12](https://github.com/danielep71/K-SFTs/issues/12)): active, no bypass. Both block deletion and force pushes and require a pull request with 0 approvals, resolved review threads and a passing, up-to-date `Repository integrity`. A separate ruleset also blocks deleting `release/**`, and another protects `v*` tags from deletion, force pushes and updates. Accepted exception: a new `release/**` branch may be created before its check runs. Required check, up-to-date and merge-method rules demonstrated on `release/**` ([#12](https://github.com/danielep71/K-SFTs/issues/12), [below](#demonstrated-enforcement)); the other rules are verified only. |
| `Repository integrity` check | Configured ([#5](https://github.com/danielep71/K-SFTs/issues/5)): runs `tools/check.py --ci` on every pull request and push to `main` and `release/**`, without a path filter. Verified on 2026-10-24 as the required check of both rulesets. |
| Merge methods | Verified on 2026-10-24: the repository allows merge commits, squash and rebase. The rulesets allow **merge commits only** into `main`, and **squash or merge commits** into `release/**`, matching [Integration](#-integration). Observed: task pull requests squash-merged into `release/0.1.0` (#40, #43), the release integrated into `main` with a merge commit (#41), and the sync back with a merge commit (#44). Head branches are deleted automatically after merge. |
| Label workflows | Verified as installed: `Validate label manifest` and `Label drift fixtures` run on pull requests into `main` and `release/**` that change label files; live label reconciliation runs on pushes to `main` that change them; live drift detection runs daily. |
| Codex code review | Enabled: observed on [#14](https://github.com/danielep71/K-SFTs/pull/14), on opening and on an `@codex review` request. |
| Dependency updates | Configured for GitHub Actions ([#5](https://github.com/danielep71/K-SFTs/issues/5)): weekly grouped pull requests against the active release branch. Dependabot reads its configuration from `main`, where it has been since [#15](https://github.com/danielep71/K-SFTs/pull/15). Active: verified on 2026-10-24 by its first pull request, [#18](https://github.com/danielep71/K-SFTs/pull/18). |
| Security features | Read on 2026-10-25 ([#123](https://github.com/danielep71/K-SFTs/issues/123)): private vulnerability reporting enabled; Dependabot alerts enabled (maintainer-reported); Dependabot security updates, secret scanning, push protection and code scanning not enabled, by the maintainer's decision. Recorded in [SECURITY.md](../SECURITY.md) ([#3](https://github.com/danielep71/K-SFTs/issues/3)). |
| Issue forms and PR template | Configured ([#11](https://github.com/danielep71/K-SFTs/issues/11)): Bug report, Enhancement and Documentation forms; blank issues disabled. |
| Issue triage workflow | Configured ([#11](https://github.com/danielep71/K-SFTs/issues/11)): read-only audit of open issues on issue events and manual dispatch. Issue events run it from `main`, where it has been since [#19](https://github.com/danielep71/K-SFTs/pull/19). Active: verified on 2026-10-24 by its runs on issue events, for example run [40104872758](https://github.com/danielep71/K-SFTs/actions/runs/40104872758). |
| Traffic history | Configured ([#36](https://github.com/danielep71/K-SFTs/issues/36)): `traffic-history` created as orphan commit `8560b31`; workflow on `main` since #41, scheduled daily. Verified on 2026-10-24: its ruleset is active, targets only `traffic-history`, blocks deletion and force pushes, and has no bypass. The `analytics` environment allows only `main` and holds `TRAFFIC_TOKEN` (read back in Settings by the maintainer). The first manual collection succeeded (run 40079797041). Pending: the first scheduled run. |

<a id="demonstrated-enforcement"></a>

### Demonstrated enforcement

Recorded on **2026-10-24** for
[#12](https://github.com/danielep71/K-SFTs/issues/12). Each row is a merge
GitHub refused, or a rule shown to hold, on a real pull request.

| Rule | Evidence |
| --- | --- |
| A failing required check blocks the merge | Controlled exercise [#79](https://github.com/danielep71/K-SFTs/pull/79) into `release/0.1.0`. Head `62c4518` added `tests/modules/TEST_DefectProbe.bas`, whose line 13 ends in trailing white space. Run [40102906705](https://github.com/danielep71/K-SFTs/actions/runs/40102906705) checked test merge `49a8403`, and `vba-source` failed with `tests/modules/TEST_DefectProbe.bas:13: trailing white space`. A squash merge pinned to that head was refused: *405 Repository rule violations found. Required status check "Repository integrity" is failing.* Correction `607a848` removed the module, and run [40103030691](https://github.com/danielep71/K-SFTs/actions/runs/40103030691) passed on test merge `808e576`. The pull request was squash-merged, so the defective revision never reached an integration branch. |
| A green check is not permission to merge | Sync [#49](https://github.com/danielep71/K-SFTs/pull/49) had an empty diff and a passing check, but was refused with *405 Base branch was modified* because the release branch had moved. [#29](https://github.com/danielep71/K-SFTs/pull/29), whose base had moved, was refused with *Required status check "Repository integrity" is expected*. |
| Only the allowed merge methods | While merge commits were not allowed on `release/**`, the merge of [#31](https://github.com/danielep71/K-SFTs/pull/31) was refused with HTTP 405. |

**What a pull request run checks.**

- **The candidate.** A run checks GitHub's test merge of the head into the
  current base, not the head alone. The run name carries the SHA of that test
  merge. The *Verify candidate identity* step fails the job if the checked-out
  commit differs from it.
- **The whole change.** Each gate reads every file in its scope in that
  commit, not only the changed files, and the workflow has no path filter,
  so every pull request is checked. Files outside every gate's scope, such as
  most Markdown documents, are checked by review only. The up-to-date rule
  makes the checked tree the tree that merges.
- **Not each commit.** Intermediate commits are not checked one by one. A
  squash merge lands only the checked tree.

**Not demonstrated.** These are verified by read-back only: the
resolved-conversation rule, the blocks on deletion and force pushes, the tag
ruleset, and every rule of the `main` ruleset, which has the same required
check but was not exercised with a failing one.

### Accepted limitations

- The sole owner cannot approve their own pull requests, so no approving
  review is required. The maintainer review, Codex review and CI above take
  its place.
- Live label reconciliation runs only from `main`: on pushes to `main` that
  change label files, and on manual dispatch on `main` (a dispatch on any
  other branch skips it). Label changes merged into a release branch take
  effect after integration into `main`.
- Run logs and report artifacts expire: the artifacts after 30 days, the logs
  after the repository's Actions retention period. The decisive lines of the
  enforcement evidence are therefore copied into
  [Demonstrated enforcement](#demonstrated-enforcement) and
  [#12](https://github.com/danielep71/K-SFTs/issues/12).
- A merge refusal can only be shown by attempting the merge. A rule that has
  not been triggered on a real pull request is recorded as verified, not as
  enforced.

Update this section, with its date, whenever a control is configured,
verified or changed.

## 🗂️ Issues

Open issues with a form: **Bug report**, **Enhancement** or **Documentation**
([.github/ISSUE_TEMPLATE](../.github/ISSUE_TEMPLATE)); blank issues are
disabled. Each form asks for the scope, the reproduction or requested
behavior, checkable acceptance criteria, the validation evidence required and
the dependency status. Pull requests follow
[.github/PULL_REQUEST_TEMPLATE.md](../.github/PULL_REQUEST_TEMPLATE.md).

Every issue has:

- the assignee **danielep71**;
- exactly one priority, `P1`, `P2` or `P3`, and the appropriate type labels
  from [.github/labels.json](../.github/labels.json);
- the milestone of the work it belongs to.

Milestones belong only to issues, so milestone progress counts tasks rather
than counting their implementation pull requests again. All pull requests
(task, integration, sync and bot-authored) have **no milestone** and keep the
labels appropriate to the work they deliver.

Task pull requests link the issues they resolve; integration and sync pull
requests reference the issues and preceding pull requests they carry. Bot
pull requests are linked to an existing relevant issue or a new tracking issue
during triage. Assign the milestone to that issue, never to the pull request.

### Priorities

| Priority | Use |
| --- | --- |
| `P1` | Reserved for a **confirmed calculation bug**: wrong results for valid inputs. Never for setup, documentation or tooling work. |
| `P2` | Material correctness, robustness or maintainability defect. Default of the Bug report form. |
| `P3` | Non-blocking quality, testing, documentation or setup work. Default of the Enhancement and Documentation forms. |

A form sets its default priority; triage confirms or changes it. Raise a bug
to `P1` only once the calculation defect is confirmed, and only with the `bug`
label.

### Dependency status

`blocked` is the only dependency status label. It means that a named,
unresolved prerequisite prevents work from starting. Ready issues carry no
status label.

- Name every blocker in the issue's *Dependencies and blocking status*
  section, as an issue link (`#12`) or on a line starting
  `External condition:`. Write `None` when nothing blocks the work.
- Add `blocked` when a named prerequisite is unresolved, and remove it when
  all of them are resolved or the maintainer explicitly resolves the
  dependency. When an issue closes, review the issues that name it.

### Triage

Issue forms cannot set a milestone or enforce the rules above, so the
read-only **Issue triage** workflow
([.github/workflows/issue-triage.yml](../.github/workflows/issue-triage.yml))
audits every open issue whenever an issue is opened, edited, closed,
reopened, labelled, assigned or given a milestone, and on manual dispatch.
It fails, and GitHub notifies the maintainer, when an open issue:

- is not assigned to danielep71, has no milestone, or does not have exactly
  one priority;
- has `P1` without `bug`;
- carries a label that is not in the catalogue;
- has no dependency section, or names a number that is neither an issue
  nor a pull request (a named pull request is a reference, never a
  prerequisite);
- is `blocked` without an open prerequisite or external condition, or names
  an open prerequisite or external condition without being `blocked`.

The workflow never changes an issue: it does not assign, label or set a
milestone. On each new issue, and whenever the audit fails, the maintainer:

1. sets the milestone of the work the issue belongs to;
2. confirms or changes the priority;
3. adds or removes `blocked` as the dependency section requires;
4. re-runs the workflow, or makes the next issue change, and checks that it
   passes.

The same rules run locally with
`python tools/check_issues.py --repository danielep71/K-SFTs` and a
`GITHUB_TOKEN`; `python tools/check.py` checks the forms and the label
catalogue offline.

### Closure

Close an issue as completed only when its acceptance criteria are met by
linked evidence. Check an acceptance box only when the linked evidence
demonstrates it. When the issue is closed, its body gains a **Resolution**
section with:

- links to the change (pull request and commit);
- the checks actually run, with their results;
- what was not validated, and any accepted limitation.

Close for another reason (duplicate, not planned, superseded) only with an
explicit explanation. Record material scope or policy decisions in the issue,
pull request or commit description.

## 📦 Releases

Versions follow Semantic Versioning; the changelog and its rules are in
[CHANGELOG.md](../CHANGELOG.md) and the procedure for preparing and
publishing a version is in [RELEASING.md](../RELEASING.md). No version has
been released and no tag exists. Closing a milestone or integrating a
release branch into `main` does not publish a release: preparation and
publication each happen only on an explicit maintainer decision.

## 🔐 Security and licensing

Never include tokens, credentials or confidential client data in issues,
examples or commits. Report suspected vulnerabilities privately as
[SECURITY.md](../SECURITY.md) describes; it also records the security features
in use.

The repository is licensed under the Mozilla Public License 2.0
([LICENSE](../LICENSE)). [NOTICE.md](../NOTICE.md) holds the copyright notice
and the provenance register of files adapted from other repositories: add an
entry in the change that adapts a file.
