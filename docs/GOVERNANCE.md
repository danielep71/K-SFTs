# 🏛️ Repository governance

This document defines how changes to CCR-K-SFTs are proposed, reviewed and
integrated. It separates the **intended policy** from the **controls verified
as configured**; a rule written here is not evidence that GitHub enforces it.

## 👤 Ownership and scope

**Daniele Penza (@danielep71)** maintains CCR-K-SFTs and decides scope,
priorities, branch strategy, integration and publication. Contributors and
automation work within the scope authorized by the maintainer.

This is an independent personal project.

CCR-K-SFTs covers counterparty credit risk for securities financing
transactions. The current milestone, **v0.1.0 - Repo setup**, prepares the
repository infrastructure only. Production SFT calculations, numerical
certification and workbook packaging belong to later milestones; the project
boundary and architecture are defined separately ([#2](https://github.com/danielep71/CCR-K-SFTs/issues/2)).

Commits may be authored by Daniele Penza or by bots acting within authorized
repository workflows. Keep author attribution truthful: a bot-authored change
is not recreated under the maintainer's name only to change its authorship.
Bot-authored pull requests remain subject to the applicable checks and
maintainer review; bot authorship does not authorize automatic merging or
bypass branch rules.

## 🌿 Branches

| Branch | Role |
| --- | --- |
| `main` | Integrated source. Receives the active release branch through a merge commit after maintainer approval. No task branches or task pull requests target it. |
| `release/0.1.0` | **Active setup integration branch** for v0.1.0 - Repo setup. Base for task branches and target of task and sync pull requests. Created from `main` at `85b2b3b` on 2026-10-23 by maintainer decision. |
| Other `release/X.Y.Z` | Opened only by maintainer decision, without a `v` prefix. Do not infer a new target from a version number or from this document. |
| `fix/<issue>-<slug>`, `docs/<slug>`, `test/<slug>` | Short-lived task branches, created from the active release branch and deleted after merge. |

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

Codex does not review bot-authored pull requests; for those, the maintainer's
review of the final diff and head SHA is the only review. Do not claim Codex
coverage when it was skipped. If review fails or is unavailable, leave the
pull request open and record the blocker.

## 🔐 Protection policy

### Intended rules

Applies to `main` and every `release/**` branch:

| Rule | Policy |
| --- | --- |
| Pull request required | Yes; 0 approving reviews (sole owner). |
| Required status check | `Repository integrity`, the job of `.github/workflows/repository-integrity.yml` ([#5](https://github.com/danielep71/CCR-K-SFTs/issues/5)). |
| Branch up to date before merge | Required. |
| Conversations resolved before merge | Required. |
| Force pushes | Blocked. |
| Deletion | Blocked. |
| Bypass | None. |

A required check must report on every pull request into a protected branch.
Path-filtered workflows, such as the label workflows, must not be the only
required checks, or unrelated pull requests cannot merge.

### Configured controls

Recorded on **2026-10-23**. *Verified* means read back from GitHub;
*pending* means not yet configured or not yet verified. Enforcement is
configured, read back and exercised in [#12](https://github.com/danielep71/CCR-K-SFTs/issues/12).

| Control | Status on 2026-10-23 |
| --- | --- |
| Repository visibility | Private. Not to be changed for setup. |
| Account plan | GitHub Pro (maintainer-reported). Rulesets and branch protection are available on private repositories under this plan; no upgrade is required. |
| `main` and `release/0.1.0` branches | Verified: both exist and point at `85b2b3b`. |
| Rulesets for `main` and `release/**` | Pending: not read back; configuration and verification in [#12](https://github.com/danielep71/CCR-K-SFTs/issues/12). |
| `Repository integrity` check | Configured ([#5](https://github.com/danielep71/CCR-K-SFTs/issues/5)): runs `tools/check.py --ci` on every pull request and push to `main` and `release/**`, without a path filter. Not yet required by a ruleset ([#12](https://github.com/danielep71/CCR-K-SFTs/issues/12)). |
| Squash merging available | Pending: not verified ([#12](https://github.com/danielep71/CCR-K-SFTs/issues/12)). |
| Label workflows | Verified as installed: `Validate label manifest` and `Label drift fixtures` run on pull requests into `main` and `release/**` that change label files; live label reconciliation runs on pushes to `main` that change them; live drift detection runs daily. |
| Codex code review | Enabled: observed on [#14](https://github.com/danielep71/CCR-K-SFTs/pull/14), on opening and on an `@codex review` request. |
| Dependency updates | Configured for GitHub Actions ([#5](https://github.com/danielep71/CCR-K-SFTs/issues/5)): weekly grouped pull requests against the active release branch. Dependabot reads its configuration from `main`, so it is active only after integration. |
| Security features | Dependabot alerts enabled (maintainer-reported); secret scanning, code scanning and private vulnerability reporting not available on this private repository ([#3](https://github.com/danielep71/CCR-K-SFTs/issues/3)). Recorded in [SECURITY.md](../SECURITY.md). |
| Issue forms and PR template | Configured ([#11](https://github.com/danielep71/CCR-K-SFTs/issues/11)): Bug report, Enhancement and Documentation forms; blank issues disabled. |
| Issue triage workflow | Configured ([#11](https://github.com/danielep71/CCR-K-SFTs/issues/11)): read-only audit of open issues on issue events and manual dispatch. Issue events run it from `main`, so it is active only after integration. |

### Accepted limitations

- The sole owner cannot approve their own pull requests, so no approving
  review is required. The maintainer review, Codex review and CI above take
  its place.
- Live label reconciliation runs only from `main`. Label changes merged into a
  release branch take effect after integration into `main`.

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

Pull requests into the active release branch carry the release milestone and
the labels of the work they deliver.

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
- has no dependency section, or names an issue that does not exist;
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
`python tools/check_issues.py --repository danielep71/CCR-K-SFTs` and a
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

The version, changelog and release procedure are defined in [#13](https://github.com/danielep71/CCR-K-SFTs/issues/13). Until then,
no release, tag or workbook is published. Closing a milestone or integrating a
release branch into `main` does not publish a release; publication happens
only on an explicit maintainer decision.

## 🔐 Security and licensing

Never include tokens, credentials or confidential client data in issues,
examples or commits. Report suspected vulnerabilities privately as
[SECURITY.md](../SECURITY.md) describes; it also records the security features
in use.

The repository is licensed under the Mozilla Public License 2.0
([LICENSE](../LICENSE)). [NOTICE.md](../NOTICE.md) holds the copyright notice
and the provenance register of files adapted from other repositories: add an
entry in the change that adapts a file.
