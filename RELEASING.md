# 🚀 Releasing

My procedure for preparing and publishing a version. The branch
model and integration are in [docs/GOVERNANCE.md](docs/GOVERNANCE.md); the
version and changelog rules are in
[CHANGELOG.md](CHANGELOG.md#date-and-version-rules).

> [!IMPORTANT]
> Nothing is versioned, tagged or published unless I explicitly
> decide it. Closing a milestone, or integrating a release branch into
> `main`, publishes nothing. Unreleased work stays under `## [Unreleased]`.

<a id="current-status"></a>

## 🧭 Current status

| Item | State |
| --- | --- |
| Released versions | [0.2.0](CHANGELOG.md#020) (2026-10-30), the first. *v0.1.0 - Repo setup* closed without a release ([#13](https://github.com/danielep71/K-SFTs/issues/13)) and is included in it |
| `VERSION` | `0.2.0` |
| Next version | Not decided. Later milestones lead to v1.0.0 |
| Active branch | `release/0.2.0` ([docs/GOVERNANCE.md](docs/GOVERNANCE.md#-branches)) |

<a id="versioning"></a>

## 🏷️ Versioning

| Item | Convention |
| --- | --- |
| Version | Semantic Versioning, `MAJOR.MINOR.PATCH`; `0.y.z` while in initial development ([rules](CHANGELOG.md#date-and-version-rules)) |
| Release branch | `release/X.Y.Z`, without `v`, opened only by my decision |
| Milestone | `vX.Y.Z - <title>`, on issues only |
| `VERSION` | One `X.Y.Z` line at the root. Created in the release-preparation pull request, never before; afterwards it always names the newest released version |
| Changelog | Work stays under `## [Unreleased]`; preparation moves it to `## [X.Y.Z] - YYYY-MM-DD` |
| Tag | Annotated `vX.Y.Z` on the candidate commit on `main`; the `v*` tag ruleset blocks its deletion, force pushes and updates |
| GitHub Release | Created from the tag. Until a production workbook is packaged: release notes only, with no asset |

The `changelog` gate of `python tools/check.py` enforces the link between the
two files: no `VERSION` before a dated release, and from then on `VERSION`
equal to the newest one ([tools/README.md](tools/README.md#-gates)).

<a id="preparation-and-publication"></a>

## 🧩 Preparation and publication

A release has two parts, each decided by me separately:

| Part | Steps | Effect |
| --- | --- | --- |
| **Preparation** | 1 to 4 | Changes only the repository: the dated changelog section, `VERSION`, the integration into `main` and the evidence. It can stop at any point and publishes nothing. |
| **Publication** | 5 and 6 | Creates the tag and the GitHub Release. A tag that may have been fetched is never moved or deleted. |

<a id="release-invariants"></a>

## 🔒 Release invariants

A version may be published only when:

1. one exact candidate commit is frozen: the merge commit that integrates the
   release branch into `main`;
2. the changelog section for the version is dated and matches what the
   candidate contains;
3. `python tools/check.py --ci` passes on the candidate, and the hosted
   `Repository integrity` check is green on it;
4. the candidate's VBA is imported, compiled and run in Excel on a
   [supported host](INSTALLATION.md#supported-hosts), and
   `tools/check_excel_run.py` passes on that evidence against the candidate
   ([recording an Excel run](INSTALLATION.md#recording-an-excel-run));
5. the annotated tag targets the candidate; and
6. the checks after publication pass.

If the source changes after steps 3 and 4, the evidence is stale and is
produced again for the new candidate.

## 📦 Preparing and publishing a version

### 1. Finalize the changelog and `VERSION`

In a task pull request into the release branch, move the `[Unreleased]`
entries into `## [X.Y.Z] - YYYY-MM-DD` and add its link reference,
`[X.Y.Z]: https://github.com/danielep71/K-SFTs/releases/tag/vX.Y.Z`.
Keep an empty `## [Unreleased]` section, with its link reference changed to
`https://github.com/danielep71/K-SFTs/compare/vX.Y.Z...HEAD`. Create or
update `VERSION` to `X.Y.Z` in the same pull request. Update the *Current
status* table above.

### 2. Integrate into `main` and freeze the candidate

Integrate the release branch into `main`, and sync it back, as in
[docs/GOVERNANCE.md](docs/GOVERNANCE.md#-integration). The integration merge
commit is the candidate: record its full SHA. Until the tag exists, the
release branch takes only fixes for a gate that failed on the candidate.
Each fix records its change in the dated section being released and is
integrated again; the new merge commit becomes the candidate.

### 3. Run the static checks

From a clean clone of the candidate:

```sh
git switch --detach FULL_SHA
python tools/check.py --ci
```

Confirm that `Repository integrity` is green on the same SHA on `main`.

### 4. Record the Excel run

Build a development workbook from the candidate and record the run as in
[recording an Excel run](INSTALLATION.md#recording-an-excel-run). Only
`excel-run: PASS` from

```sh
python tools/check_excel_run.py <evidence-folder> --commit FULL_SHA
```

qualifies the candidate. If it fails, fix through a task pull request and
start again at step 2.

### 5. Create the annotated tag

Only on my decision to publish, from the command line so that
the tag is annotated:

```sh
git fetch origin
git tag -a vX.Y.Z FULL_SHA -F ../tag-message-vX.Y.Z.txt
git cat-file -t vX.Y.Z        # must print: tag
git push origin vX.Y.Z
```

The tag message is the release's permanent record:

```text
vX.Y.Z - <milestone title>

Candidate <full 40-character SHA>.
Static checks: check.py --ci PASS; Repository integrity run <run id> PASS.
Excel: <host, build, bitness>; check_excel_run PASS.
Issues: #<n>, #<n>.
Deviations: none.
```

`Deviations:` is mandatory: write `none`, or list each invariant that was not
met. Copy results from the runs themselves.

### 6. Publish and verify

Create the GitHub Release from the tag, titled `vX.Y.Z - <milestone title>`,
with the changelog section and the evidence summary as notes. Until a
production workbook is packaged, attach nothing. Then check that a fresh clone of the tag passes
`python tools/check.py`, and that the release page names the tag and the
candidate SHA.

## 🧾 Evidence record

Keep with the release, in the release notes or the closing issue: the
candidate SHA, the integration pull request, the `check.py --ci` and
`Repository integrity` results, the `check_excel_run.py` report with the host
it names, the tag name and its target, and any deviation. Logs and artifacts
expire, so copy the decisive lines
([docs/GOVERNANCE.md](docs/GOVERNANCE.md#accepted-limitations)).

## 🧯 Recovery

- **Before the tag:** fix through a task pull request and repeat from step 2.
- **After the tag, before the Release:** stop and record the problem. If the
  tag may have been fetched, prefer a new patch version to deleting it.
- **After publication:** never move the tag or replace an asset. Publish a
  corrected patch version and record the problem in the changelog; report a
  vulnerability as [SECURITY.md](SECURITY.md) describes.

## ☑️ Checklist

| # | Gate | Done |
| ---: | --- | :---: |
| 1 | I decided to prepare the version | ☐ |
| 2 | Changelog section dated and `VERSION` created or updated | ☐ |
| 3 | Release branch integrated into `main`, and synced back | ☐ |
| 4 | Candidate frozen: the integration merge commit's full SHA | ☐ |
| 5 | `check.py --ci` and `Repository integrity` pass on the candidate | ☐ |
| 6 | `check_excel_run.py` passes on the candidate's Excel run | ☐ |
| 7 | I decided to publish | ☐ |
| 8 | Annotated tag on the candidate | ☐ |
| 9 | GitHub Release published from the tag and verified | ☐ |
