<!--
  Base: the active release branch named in docs/GOVERNANCE.md for task PRs.
  main only for a release integration approved by the maintainer; main into
  the release branch only for a sync PR. State which route applies below.
  Write NOT RUN or NOT APPLICABLE with a reason; never claim evidence that
  was not produced. Synthetic data only; no secrets or vulnerability details.
-->

## 🔗 Linked issues

<!-- Link task issues with Resolves #; integration and sync PRs use Refs #
for the issues and preceding PRs they carry. Milestones belong to issues only. -->

Resolves #

## 🧾 Changes

-

## 💡 Why

## ✅ Validated

<!-- Commands and checks actually run, with results. Hosted checks run on this pull request. -->

- [ ] `python tools/check.py` (all gates PASS; INCOMPLETE is not a pass)
- [ ] Other checks specific to this change:

## 🔎 Final-head review

<!-- After a push, wait for any earlier review to finish, then post
@codex review as a top-level comment and compare the completed review's SHA
with the current head (docs/GOVERNANCE.md, Review). For a bot-authored PR,
record the maintainer's review of the final head and that Codex was skipped.
A sync PR whose diff is empty needs CI only: record the empty diff instead of a
review. -->

- Final PR head SHA:
- Review type (Codex / maintainer review of a bot-authored PR / not required: empty-diff sync PR):
- Reviewed commit SHA and review permalink:
- `Repository integrity` result and run link:

- [ ] Review covers the current PR head; all findings are fixed or answered and applicable threads resolved.
- [ ] `Repository integrity` passes for the current revision and GitHub merge requirements are satisfied.
- [ ] Any push after this record is followed by renewed review and CI verification.

## ⚠️ Not validated

<!-- State plainly what was not checked and why. Write "Nothing" only if that is true. -->

## ☑️ Checklist

- [ ] The base follows docs/GOVERNANCE.md: active release branch for tasks; main only for an approved integration.
- [ ] Each linked issue has the assignee, exactly one priority label and a milestone; this PR has no milestone, links the relevant issues and carries appropriate labels.
- [ ] Scope is focused: no unrelated changes, local artifacts or generated output.
- [ ] Documentation affected by the change is updated.
- [ ] No tokens, credentials, real portfolio or client data.
