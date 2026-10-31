# 📈 Repository traffic history

GitHub keeps repository traffic for 14 days. The **Traffic history** workflow
([.github/workflows/traffic-history.yml](../.github/workflows/traffic-history.yml))
collects it once a day, with the repository's stars, forks, watchers and open
issues, and merges it into CSV files on the data-only `traffic-history`
branch, so the history outlives that window. It also writes the README's
traffic badges there, and opens an issue when traffic stands out
(*Traffic alert* below). It is repository analytics, not CI: it builds and
tests nothing.

## ⏱️ Schedule

The workflow is scheduled daily at **06:17 UTC** and can be started from
**Actions → Traffic history → Run workflow** on `main`. GitHub may start
scheduled runs late; each collection records the time it was actually taken.
A manual run on any other branch is skipped.

## 🔑 GitHub App setup

The traffic API needs **Administration: read** on this repository. The
workflow token cannot have that permission, so collection uses a short-lived
GitHub App installation token.

1. In **Settings → Environments → analytics**, restrict deployment branches
   to the default branch, `main`, before adding credentials. A run naming a
   missing environment creates it without that restriction.
2. Under **Settings → Developer settings → GitHub Apps** for `danielep71`,
   create the traffic App, or use the existing dedicated traffic App. Disable
   its webhook and grant only **Administration: Read-only** as a repository
   permission (GitHub adds **Metadata: Read-only**). Limit installation to
   this account.
3. Install the App on `K-SFTs` using **Only select repositories**. If the App
   already serves other repositories, add this repository without expanding
   its permissions. The workflow requests a token for this repository only.
4. In the `analytics` environment, set the variable **TRAFFIC_APP_CLIENT_ID**
   to the App's **Client ID**. Generate a dedicated private key and put the
   whole `.pem` content in the environment secret **TRAFFIC_APP_PRIVATE_KEY**.
   Delete the downloaded key file after saving it. Never commit the key or
   put it in an issue, comment or chat.
5. After configuration is confirmed and the workflow is integrated into the
   default branch, run **Actions → Traffic history → Run workflow** there.
   Check that collection succeeds and `traffic-history` gains a commit.

### Migration from the personal token

I install the App, generate its key and configure the environment variable
and secret. The migration pull request waits for my configuration confirmation
before merge. Only after the first successful App-based run and verified
history commit do I delete the old `TRAFFIC_TOKEN` secret from `analytics`
and revoke the personal token in my account settings. The workflow has no
fallback to that secret; its step-local `TRAFFIC_TOKEN` variable now holds
only the App installation token.

### Credential isolation and key rotation

Only **Create the traffic App token** receives the private key; the preceding
configuration check sees only whether it is set. The action requests an
installation token limited to this repository and **Administration: read**,
which expires after an hour and is revoked by the action when the job ends.
Only the fetch step receives it, and that step runs `curl`, not repository
code. The merge script receives neither credential; the push uses the
workflow token with `contents: write`.

Rotate the key if it may have been exposed or if anyone but me changes the
workflow. For a routine rotation:

1. Generate a new private key in the App's settings.
2. Replace **TRAFFIC_APP_PRIVATE_KEY** in **Settings → Environments →
   analytics**. The environment secret takes precedence over any repository
   secret with the same name.
3. Run the workflow on the default branch and verify successful API collection
   and a new `traffic-history` commit.
4. Delete the old key in the App's settings, checking its fingerprint first.
   If compromise is suspected, revoke the exposed key immediately instead
   of waiting for this verification.

Missed collections are lost once they fall outside the API's 14-day window.

## 🗄️ Stored data

The `traffic-history` branch holds `README.md` and these files under `data/`.
It was created on 2026-10-24 as the orphan commit `8560b31`, with the files
below empty apart from their headers, so it has no history in common with
`main`; `8915601` then set the 7-column `daily.csv` header, before any data.
The merge script rewrites a file still in an earlier layout with the current
header the next time it merges a collection. The branch is never merged into
`main` or a release branch, and its history is not rewritten.

| File | Key | Contents |
| --- | --- | --- |
| `daily.csv` | `date` | Views and unique views, clones and unique clones per UTC day. `views_collected_at` and `clones_collected_at` give the collection that last changed each pair; `views_observed_at` and `clones_observed_at`, the newest collection that reported it, whether or not it changed the values. |
| `snapshots.csv` | `snapshot_date` | One row per collection day: the first and last day the API reported, the 14-day totals, the repository's `stars`, `forks`, `watchers` and `open_issues` at collection time, `collected_at`, the collection that last changed that day's snapshot, referrers or paths, and `observed_at`, the newest collection of that day. |
| `referrers.csv` | `snapshot_date`, `rank` | The top 10 referrers over the 14 days before each collection. |
| `paths.csv` | `snapshot_date`, `rank` | The top 10 repository paths over the same window. |
| `badges/*.json` | — | The README badges, rewritten from the stored data after each merge (*Badges* below). |

Referrers and paths are retained per collection day, because the API gives
them only as 14-day totals, never per day.

The repository counts come from the repository API (`stargazers_count`,
`forks_count`, `subscribers_count`, `open_issues_count`), as in my other
engine's traffic history. `open_issues` is GitHub's count, which includes
open pull requests. Snapshots collected before this change reached `main`
(#236) have them blank: a blank count was not recorded, not zero.

### Badges

The README shows three badges read by shields.io from `data/badges/`, in its
*endpoint* format:

| Badge | Counts |
| --- | --- |
| `clones` | Every clone in `daily.csv`, summed since tracking began. |
| `visitors (14d)` | Unique visitors over the API's 14-day window, from the latest snapshot. Daily unique counts cannot be summed into a unique count. |
| `visits` | Daily unique visitors summed since tracking began: a visitor counts once per day they visit. |

Large numbers are shortened (`1.2k`, `11k`). A badge file changes only when
its value does, so a new collection with the same data leaves the badges
unchanged.

### Merge rules

`tools/traffic_history.py` merges each collection by date and metric:

- **Idempotent:** merging the same collection again, with the same collection
  time, changes nothing.
- **Newest wins:** for every date and metric, the values of the newest
  collection that reported them win, even when its counts are lower than the
  stored ones, so a day that was incomplete when first collected is corrected
  later. "Newest" means the latest collection time. Views and clones are
  merged separately, each with two times: `*_collected_at`, the collection
  that last changed the values, and `*_observed_at`, the newest collection
  that reported them, which also counts when it left them unchanged.
- **Order:** a collection older than the stored observation never replaces
  anything, whether it is a saved collection merged again by hand or a run
  rebuilt on a newer branch tip. It can only fill a date or a metric that no
  newer collection reported.
- **Older history:** dates the collection does not report keep their values.
  Nothing is deleted.
- **No double counting:** values are replaced, never added. Unique counts are
  never summed across days or windows; a 14-day unique count is only in
  `snapshots.csv`.
- **Same-day runs:** a collection day's snapshot, referrers and paths are one
  unit. A newer collection on the same UTC day replaces the unit only if any
  of the three changed; otherwise it updates the snapshot's `observed_at`
  alone. An older one never changes it.
- **New identical data:** a newer collection whose values match the stored
  ones still records its observation time in `daily.csv` and `snapshots.csv`,
  so every successful run adds a commit to `traffic-history`; the badges and
  the other files stay unchanged.
- **Earlier layouts:** a `daily.csv` without the `*_observed_at` columns, or a
  `snapshots.csv` without `observed_at`, is read with each observation time
  equal to its collection time, and rewritten in the current layout on the
  next merge, with every stored value and collection time unchanged.

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
read its CSV files. By my decision
([#123](https://github.com/danielep71/K-SFTs/issues/123)) it stays public. The
data is aggregate (counts, referrers and paths) and identifies no visitor. It
is not published through
GitHub Pages; the three badges show the totals above. The counts include my
own visits and every clone made by CI checkouts.

## 🚨 Traffic alert

After each collection the **Traffic alert** job compares the latest snapshot
with the one before (`traffic_history.py alert`) and opens one issue, titled
`Traffic alert: <snapshot date>`, when any of these holds:

- **Views or clones spike:** the 14-day total at least doubles and grows by
  at least 100 views or 10 clones;
- **Stars or forks increase;**
- **New referrer:** a referrer never stored before brings at least 5 views.

A count blank in the older snapshot raises nothing. At most one issue is
opened per snapshot date: the job skips a date whose issue exists, open or
closed. Like every issue it is assigned to me, labelled `P3` and
`repository`, and filed under the open milestone with the earliest due date,
so it passes *Issue triage*. Its *Schedule* starts on the snapshot date and
targets a week later. The thresholds are those of my other engine's traffic
workflow.

The job uses the workflow token with `issues: write` and `contents: read`. It
has no environment, so it receives neither the App's private key nor its
installation token, and it runs only
after a successful collection.

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
| *Check the traffic App configuration*: `TRAFFIC_APP_CLIENT_ID or TRAFFIC_APP_PRIVATE_KEY is not set` | Missing environment variable or secret, or the wrong environment. | Follow the App setup above. The check sees only a Boolean for the key. |
| *Create the traffic App token* fails | Wrong Client ID or key, deleted key, or App not installed on this repository. | Check installation and Client ID; replace or rotate the key as above. |
| *Fetch*: HTTP 401 | Installation token rejected, for example after App removal during a run. | Check installation and rerun. |
| *Fetch*: HTTP 403 or 404 | App lacks *Administration: read* or access to this repository. | Check the selected repository and permission; accept any approved permission change, then rerun. |
| *Fetch*: HTTP 429 or 5xx after retries | Rate limit or GitHub outage. | Run it again later; the next run collects the same 14 days. |
| *Merge*: `ERROR: ...; the history was not changed` | The API returned an unexpected shape, or a history file no longer has the expected layout. | Inspect the response or the named file; a file edited by hand must keep its header and columns. |
| *Fetch*: `API request for the repository failed` | Same causes as the traffic requests. | As above. |
| *Commit and push*: `Could not push the history` | `traffic-history` kept moving, or a rule rejected the push. | Check the branch's rules and recent commits, then run again. |
| *Check out traffic-history* fails | The branch was deleted. | Restore it from the last known commit; never recreate it empty over existing history. |
| *Traffic alert*: `No open milestone with a due date` | Every milestone is closed, or none has a due date. | Open the next milestone with a due date, then run the workflow again. |
| *Traffic alert*: `gh issue create` fails | A label, the assignee or the milestone no longer exists. | Restore it, then run the workflow again; the history is already saved. |

A failed run can always be repeated: an older collection never replaces a
newer one, and each collection covers the last 14 days.

If the branch must be repaired, change it with a new commit made with
`python tools/traffic_history.py merge` or by hand, keeping every header.
`python tools/traffic_history.py init --data data` recreates missing files
with their headers and never overwrites an existing one.
