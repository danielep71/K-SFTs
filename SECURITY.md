# 🔒 Security policy

## 📌 Supported versions

The latest release, 0.2.0 ([RELEASING.md](RELEASING.md)), the active release
branch named in [docs/GOVERNANCE.md](docs/GOVERNANCE.md) and `main` are
supported on a best-effort basis. Reports must identify a full commit SHA.
Modified copies are unsupported unless the problem reproduces in this
repository.

## 📨 Reporting a vulnerability

Do **not** disclose a suspected vulnerability in an issue, pull request,
commit message, sample file or screenshot.

Report it privately in either of two ways:

- on GitHub, through **Security → Report a vulnerability**
  (<https://github.com/danielep71/K-SFTs/security/advisories/new>); or
- by email to **danielep71@gmail.com**, with the subject
  **Private security report — K-SFTs**.

Include:

- the full commit SHA and the affected file, procedure or workflow;
- the Excel/Office build, bitness and Windows version, where relevant;
- the impact and minimal reproduction steps, with synthetic data only; and
- any workaround you tested.

## ⏱️ Response targets

K-SFTs is maintained by one person; the targets are best effort, not a
contractual commitment:

- acknowledgement within 5 business days;
- an initial assessment within 10 business days of receiving enough evidence;
- an update at least every 14 days while the report is investigated.

Remediation and disclosure timing are agreed with the reporter according to
severity.

## 🎯 Scope

**In scope:** the repository tooling in [tools/](tools/README.md), the GitHub
Actions workflows and their configuration, and, once added, the VBA source,
workbook template and tests.

**Out of scope:**

- vulnerabilities in Excel, Office, Windows, GitHub, Python or Node.js
  themselves;
- organization-controlled macro and deployment policy;
- modified copies;
- ordinary defects without concrete security impact, such as an incorrect
  exposure figure. Report those as issues.

## ⚙️ Automation risk surface

The calculation prototype in VBA runs only in a workbook a user builds and
opens; no workflow runs it. The automation runs with the workflow token,
except the traffic collection, which also reads
`TRAFFIC_TOKEN`: a fine-grained token limited to this repository with
*Administration: read*, stored in the `analytics` environment, which only
`main` may use. Only the step that calls the traffic API receives it, and that
step runs no repository code ([docs/TRAFFIC.md](docs/TRAFFIC.md)).

| Workflow | Token permissions | Runs on |
| --- | --- | --- |
| `Repository integrity` | `contents: read` | Pull requests and pushes to `main` and `release/**`, manual dispatch |
| `Issue triage` | `contents: read`, `issues: read`, `pull-requests: read` | Issue events and manual dispatch, from `main` |
| `Detect issue-label drift` | `contents: read`, `issues: read` | Label-file pull requests, daily schedule, manual dispatch |
| `Sync issue labels` | `contents: read`; `issues: write` only to reconcile labels | Label-file pull requests (validation only); label-file pushes to `main` and manual dispatch on `main` (reconciliation) |
| `Traffic history` | Collection job: `contents: write`, to push `traffic-history`, plus `TRAFFIC_TOKEN` for the traffic and repository API. Alert job: `contents: read`, `issues: write`, to open a traffic-alert issue; no `TRAFFIC_TOKEN` | Daily schedule and manual dispatch, from `main` only |

Every checkout sets `persist-credentials: false`, and every action is pinned
to a full commit SHA; the traffic workflow passes its push credential to `git`
for that command only. No workflow approves, merges or dismisses alerts.

## 🛡️ Security features

Current security-feature observations, verification dates, evidence and
limitations are maintained in the [repository settings evidence](docs/REPOSITORY_SETTINGS_EVIDENCE.md#security-settings).
Its 2026-10-29 observations supersede the disabled-feature snapshot recorded
on 2026-10-25. This policy remains authoritative for reporting and handling
vulnerabilities; a settings record does not establish scan coverage or enforcement.

## 🔑 Data and secrets

- Fixtures and examples are **synthetic**. Never commit, attach or send real
  trades, portfolios, counterparty, collateral or client data.
- Never commit credentials, tokens, keys or internal paths. `.gitignore`
  excludes common secret and signing files (`.env`, `*.pem`, `*.key`,
  `*.pfx`); that is a safeguard, not a substitute for care.
- Excel files can hold sensitive material outside visible cells; inspect any
  workbook before sharing it.
- If a secret is exposed, revoke or rotate it at once, before reporting it.
- Renew `TRAFFIC_TOKEN` before it expires and at least every 90 days, and
  revoke it at once if the traffic workflow is changed by anyone but
  me ([docs/TRAFFIC.md](docs/TRAFFIC.md)).

## 🤝 Safe harbor

The project will not initiate or recommend legal action for good-faith
research that follows this policy, stays within this repository's own source,
avoids privacy violations and destructive actions, and is reported privately.
This does not authorize testing third-party systems; no bug bounty is offered.
