# 📈 Repository traffic history

GitHub keeps repository traffic for 14 days. The **Traffic history** workflow
([.github/workflows/traffic-history.yml](../.github/workflows/traffic-history.yml))
collects it once a day and merges it into CSV files on the data-only
`traffic-history` branch, so the history outlives that window. It is
repository analytics, not CI: it builds and tests nothing.

## ⏱️ Schedule

The workflow is scheduled daily at **06:17 UTC** and can be started from
**Actions → Traffic history → Run workflow** on `main`. GitHub may start
scheduled runs late; each collection records the time it was actually taken.
A manual run on any other branch is skipped.

## 🔑 Credential setup and renewal

The traffic API needs a token with **Administration: read** on the
repository. The workflow token cannot be given that permission, so a
fine-grained personal access token is used.

1. **Settings → Environments → New environment** `analytics`. Under
   *Deployment branches and tags*, select *Selected branches and tags* and
   allow `main` only. Do this before adding the secret: a run that names a
   missing environment creates it without that restriction.
2. **Settings → Developer settings → Fine-grained tokens → Generate new
   token**: resource owner `danielep71`, *Only select repositories*:
   `K-SFTs`, repository permission **Administration: Read-only**, nothing
   else. Set an expiry of at most 90 days.
3. In the `analytics` environment, add the secret **TRAFFIC_TOKEN** with that
   token. Never put the token in a file, issue, comment or chat.
4. Run the workflow manually on `main` and check that it passes and that
   `traffic-history` gains a commit.

**Renewal:** before the token expires, and at least every 90 days, generate a
new token with the same settings, replace the secret, and repeat step 4.
Revoke the token at once if it may have been exposed, or if the workflow is
changed by anyone but the maintainer. Collections missed while the token is
invalid are lost once they fall outside the 14-day window.

Only the fetch step receives `TRAFFIC_TOKEN`, and it runs `curl` only. The
merge script runs without it, and the push uses the workflow token with
`contents: write`.

## 🗄️ Stored data

The `traffic-history` branch holds `README.md` and these files under `data/`.
It was created on 2026-10-24 as the orphan commit `8560b31`, with the files
below empty apart from their headers, so it has no history in common with
`main`; `8915601` then set the current `daily.csv` header, before any data. It is never merged into `main` or a release branch, and its history is
not rewritten.

| File | Key | Contents |
| --- | --- | --- |
| `daily.csv` | `date` | Views and unique views, clones and unique clones per UTC day. `views_collected_at` and `clones_collected_at` give the collection that last changed each pair. |
| `snapshots.csv` | `snapshot_date` | One row per collection day: the first and last day the API reported, the 14-day totals, and `collected_at`, the collection that last changed that day's snapshot, referrers or paths. |
| `referrers.csv` | `snapshot_date`, `rank` | The top 10 referrers over the 14 days before each collection. |
| `paths.csv` | `snapshot_date`, `rank` | The top 10 repository paths over the same window. |

Referrers and paths are retained per collection day, because the API gives
them only as 14-day totals, never per day.

### Merge rules

`tools/traffic_history.py` merges each collection idempotently:

- **Overlapping windows:** for every date the collection reports, its values
  replace the stored ones. The newest collection wins, so a day that was
  incomplete when first collected is corrected later. Views and clones are
  merged separately, each with its own collection time, which stays unchanged
  when its values do not change.
- **Older history:** dates the collection does not report keep their values.
  Nothing is deleted.
- **No double counting:** values are replaced, never added. Unique counts are
  never summed across days or windows; a 14-day unique count is only in
  `snapshots.csv`.
- **Same-day runs:** a collection day's snapshot, referrers and paths are one
  unit. A second collection on the same UTC day replaces the unit only if any
  of the three changed. Running the workflow twice with the same data changes
  nothing.
- **Order:** "newest" means the latest collection time. Merging an older
  collection after a newer one, when a saved collection is replayed or a run
  is rebuilt on a newer branch tip, never replaces what the newer one stored;
  it can only fill a date or metric that nothing newer reported. A stored
  time is that of the collection that last *changed* the values; a day's
  counts only grow until the day ends, so an older collection replayed later
  cannot hold different values.

### Zero or unavailable

The API reports each day of its window, days without activity included, as
`0`. So:

- **`0`**: the day was collected and had no activity;
- **no row**: no collection covered the day, either because it is before
  tracking began or because collections stopped for more than 14 days;
- **a blank metric**: one of the two series was not reported for that day.

The step summary of each run lists the dates between the first and the last
stored day that have no row.

### Visibility

The repository is public since 2026-10-25, and so is the branch: anyone can
read its CSV files. By the maintainer's decision
([#123](https://github.com/danielep71/K-SFTs/issues/123)) it stays public. The
data is aggregate (counts, referrers and paths) and identifies no visitor. It
is not published through
GitHub Pages or badges. The counts include the maintainer's own visits and
every clone made by CI checkouts.

## 🛡️ Branch protection

The workflow pushes to `traffic-history` directly, so its rules differ from
those of `main` and `release/**` ([GOVERNANCE.md](GOVERNANCE.md#-protection-policy)):

| Rule | `traffic-history` |
| --- | --- |
| Deletion | Blocked. |
| Force pushes | Blocked. |
| Pull request required | No: the workflow commits directly. |
| Required status checks | None. |
| Bypass | None. |

## 🩺 Failure diagnosis and recovery

Each run fails before anything is written if a response is missing or
invalid, and leaves the history unchanged. Read the failing step:

| Step and message | Cause | Action |
| --- | --- | --- |
| *Fetch*: `TRAFFIC_TOKEN is not set` | Secret missing, or the run is not in the `analytics` environment. | Follow the credential setup. |
| *Fetch*: HTTP 401 | Token expired or revoked. | Renew the token. |
| *Fetch*: HTTP 403 or 404 | Token lacks *Administration: read* on this repository. | Regenerate it with the settings above. |
| *Fetch*: HTTP 429 or 5xx after retries | Rate limit or GitHub outage. | Run it again later; the next run collects the same 14 days. |
| *Merge*: `ERROR: ...; the history was not changed` | The API returned an unexpected shape, or a history file no longer has the expected layout. | Inspect the response or the named file; a file edited by hand must keep its header and columns. |
| *Commit and push*: `Could not push the history` | `traffic-history` kept moving, or a rule rejected the push. | Check the branch's rules and recent commits, then run again. |
| *Check out traffic-history* fails | The branch was deleted. | Restore it from the last known commit; never recreate it empty over existing history. |

A failed run can always be repeated: the merge is idempotent, and each
collection covers the last 14 days.

If the branch must be repaired, change it with a new commit made with
`python tools/traffic_history.py merge` or by hand, keeping every header.
`python tools/traffic_history.py init --data data` recreates missing files
with their headers and never overwrites an existing one.
