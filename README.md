# traffic-history

Data-only branch of CCR-K-SFTs. The **Traffic history** workflow on `main`
merges GitHub traffic (views, clones, top referrers and popular paths) into
the CSV files under `data/` once a day, keeping it beyond GitHub's 14-day
window.

Do not merge this branch, base work on it or rewrite its history. The file
layout, the merge rules, the meaning of `0` and of a missing date, and the
recovery steps are in `docs/TRAFFIC.md` on `main`.

| File | Contents |
| --- | --- |
| `data/daily.csv` | Views and clones per UTC day |
| `data/snapshots.csv` | One row per collection day, with the 14-day totals |
| `data/referrers.csv` | Top referrers per collection day |
| `data/paths.csv` | Popular paths per collection day |
