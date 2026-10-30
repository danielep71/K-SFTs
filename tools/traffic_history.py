#!/usr/bin/env python3
"""Merge GitHub traffic API responses into the CSV history of the traffic-history branch.

The traffic API keeps 14 days. Each collection is merged by date and metric:
the newest collection wins for every date it reports, dates it does not report
keep their stored values, and nothing is ever summed across collections. A
date that no collection reported has no row: it is unavailable, not zero.

The merge reads only local files and needs no token (docs/TRAFFIC.md). Every
response and every existing file is validated before anything is written, so
an invalid or incomplete collection leaves the history unchanged. After each
merge it rewrites the README badge files under badges/ from the stored data.

The alert command compares the two latest snapshots and, when traffic spikes,
stars or forks increase, or a new referrer appears, writes the title and body
of one alert issue for the workflow to open.
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import sys
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any

DAILY = "daily.csv"
SNAPSHOTS = "snapshots.csv"
REFERRERS = "referrers.csv"
PATHS = "paths.csv"
HEADERS: dict[str, tuple[str, ...]] = {
    DAILY: ("date", "views", "unique_views", "views_collected_at",
            "clones", "unique_clones", "clones_collected_at"),
    SNAPSHOTS: ("snapshot_date", "collected_at", "first_date", "last_date",
                "views_14d", "unique_views_14d", "clones_14d", "unique_clones_14d",
                "stars", "forks", "watchers", "open_issues"),
    REFERRERS: ("snapshot_date", "rank", "referrer", "count", "uniques"),
    PATHS: ("snapshot_date", "rank", "path", "title", "count", "uniques"),
}
# A snapshots.csv written before the repository counts were added; read as
# having them blank, and rewritten with the current header on the next merge.
LEGACY_HEADERS: dict[str, tuple[str, ...]] = {SNAPSHOTS: HEADERS[SNAPSHOTS][:8]}
RESPONSES = ("views.json", "clones.json", "referrers.json", "paths.json", "repository.json")
# Snapshot column, and the repository API field it is read from.
REPOSITORY_COUNTS = {"stars": "stargazers_count", "forks": "forks_count",
                     "watchers": "subscribers_count", "open_issues": "open_issues_count"}
BADGES = "badges"
BADGE_COLOR = "217346"
MIDNIGHT = re.compile(r"^(\d{4}-\d{2}-\d{2})T00:00:00Z$")
ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
# Collection times are always written in this form, so they compare as strings.
STAMP = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$")
NUMERIC = {"views", "unique_views", "clones", "unique_clones", "views_14d", "unique_views_14d",
           "clones_14d", "unique_clones_14d", "rank", "count", "uniques", *REPOSITORY_COUNTS}
STAMPS = {"collected_at", "views_collected_at", "clones_collected_at"}
Row = dict[str, str]


class TrafficError(Exception):
    """An invalid response or history file; nothing has been written."""


@dataclass(frozen=True)
class Series:
    """One views or clones response: 14-day totals and per-day counts."""

    count: int
    uniques: int
    days: dict[str, tuple[int, int]]


@dataclass(frozen=True)
class Collection:
    collected_at: datetime
    views: Series
    clones: Series
    referrers: list[Row]
    paths: list[Row]
    repository: dict[str, int]

    @property
    def snapshot_date(self) -> str:
        return self.collected_at.date().isoformat()

    @property
    def stamp(self) -> str:
        return self.collected_at.isoformat(timespec="seconds").replace("+00:00", "Z")


@dataclass
class Report:
    changed: dict[str, int] = field(default_factory=dict)
    gaps: list[tuple[str, str]] = field(default_factory=list)


# --- Validation of API responses -------------------------------------------------


def count(value: Any, where: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise TrafficError(f"{where}: expected a non-negative integer, got {value!r}")
    return value


def text(value: Any, where: str) -> str:
    if not isinstance(value, str):
        raise TrafficError(f"{where}: expected a string, got {value!r}")
    return value


def parse_series(data: Any, key: str, collected: date) -> Series:
    if not isinstance(data, dict) or not isinstance(data.get(key), list):
        raise TrafficError(f"{key}: expected an object with a {key!r} list")
    days: dict[str, tuple[int, int]] = {}
    for index, item in enumerate(data[key]):
        where = f"{key}[{index}]"
        if not isinstance(item, dict):
            raise TrafficError(f"{where}: expected an object")
        match = MIDNIGHT.match(text(item.get("timestamp"), f"{where}.timestamp"))
        if not match:
            raise TrafficError(f"{where}.timestamp: not aligned to UTC midnight")
        day = match.group(1)
        if day in days:
            raise TrafficError(f"{where}: duplicate date {day}")
        if date.fromisoformat(day) > collected:
            raise TrafficError(f"{where}: date {day} is after the collection date")
        days[day] = (count(item.get("count"), f"{where}.count"),
                     count(item.get("uniques"), f"{where}.uniques"))
    return Series(count(data.get("count"), f"{key}.count"),
                  count(data.get("uniques"), f"{key}.uniques"), days)


def parse_ranked(data: Any, name: str, fields: Sequence[str]) -> list[Row]:
    if not isinstance(data, list):
        raise TrafficError(f"{name}: expected a list")
    rows = []
    for rank, item in enumerate(data, start=1):
        where = f"{name}[{rank - 1}]"
        if not isinstance(item, dict):
            raise TrafficError(f"{where}: expected an object")
        row = {"rank": str(rank)}
        row.update({key: text(item.get(key), f"{where}.{key}") for key in fields})
        row["count"] = str(count(item.get("count"), f"{where}.count"))
        row["uniques"] = str(count(item.get("uniques"), f"{where}.uniques"))
        rows.append(row)
    return rows


def parse_repository(data: Any) -> dict[str, int]:
    if not isinstance(data, dict):
        raise TrafficError("repository: expected an object")
    return {column: count(data.get(field), f"repository.{field}")
            for column, field in REPOSITORY_COUNTS.items()}


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as error:
        raise TrafficError(f"{path.name}: response missing") from error
    except (OSError, ValueError) as error:
        raise TrafficError(f"{path.name}: unreadable response ({error})") from error


def load_collection(raw: Path, collected_at: datetime) -> Collection:
    """Validate the five responses of one collection; all of them or none."""
    collected = collected_at.date()
    return Collection(
        collected_at=collected_at,
        views=parse_series(load_json(raw / "views.json"), "views", collected),
        clones=parse_series(load_json(raw / "clones.json"), "clones", collected),
        referrers=parse_ranked(load_json(raw / "referrers.json"), "referrers", ("referrer",)),
        paths=parse_ranked(load_json(raw / "paths.json"), "paths", ("path", "title")),
        repository=parse_repository(load_json(raw / "repository.json")),
    )


# --- History files ---------------------------------------------------------------


def read_history(path: Path) -> list[Row]:
    """Read one history file, refusing anything whose layout is not the expected one."""
    header = HEADERS[path.name]
    if not path.exists():
        return []
    try:
        content = path.read_text(encoding="utf-8")
    except (OSError, ValueError) as error:
        raise TrafficError(f"{path.name}: unreadable ({error})") from error
    reader = csv.DictReader(io.StringIO(content))
    found = tuple(reader.fieldnames or ())
    if found != header and found != LEGACY_HEADERS.get(path.name):
        raise TrafficError(f"{path.name}: header is {reader.fieldnames}, expected {list(header)}")
    rows = []
    for number, row in enumerate(reader, start=2):
        if None in row or any(value is None for value in row.values()):
            raise TrafficError(f"{path.name}:{number}: wrong number of fields")
        if not ISO_DATE.match(row[header[0]]):
            raise TrafficError(f"{path.name}:{number}: invalid date {row[header[0]]!r}")
        complete = {key: "" for key in header} | dict(row)
        check_fields(path.name, number, complete)
        rows.append(complete)
    return rows


def check_fields(name: str, number: int, row: Mapping[str, str]) -> None:
    """Counts are digits and collection times have their fixed form.

    A daily metric may be blank together with its collection time: that metric
    was never reported for the date. A snapshot's repository counts are blank
    when it was collected before they were recorded.
    """
    def blank_metric(key: str) -> bool:
        if name == SNAPSHOTS and key in REPOSITORY_COUNTS:
            return row[key] == ""
        metric = key.removesuffix("_collected_at").removeprefix("unique_")
        return (name == DAILY and row[metric] == "" and row[f"unique_{metric}"] == ""
                and row[f"{metric}_collected_at"] == "")

    bad = [key for key in NUMERIC & set(row)
           if not (row[key].isascii() and row[key].isdigit()) and not blank_metric(key)]
    if bad:
        raise TrafficError(f"{name}:{number}: non-numeric {', '.join(sorted(bad))}")
    bad = [key for key in STAMPS & set(row)
           if not STAMP.match(row[key]) and not blank_metric(key)]
    if bad:
        raise TrafficError(f"{name}:{number}: invalid {', '.join(sorted(bad))}")


def render(name: str, rows: Sequence[Mapping[str, str]]) -> str:
    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=HEADERS[name], lineterminator="\n")
    writer.writeheader()
    writer.writerows(rows)
    return buffer.getvalue()


# --- Merge rules -----------------------------------------------------------------


def merge_daily(rows: list[Row], collection: Collection) -> list[Row]:
    """Newest collection wins per date and metric; unreported dates keep their values.

    Views and clones each carry the time of the collection that last changed
    them. A metric stored by a later collection than this one is never
    replaced, so replaying an older collection cannot undo newer counts, yet it
    can still fill a metric that no newer collection reported.
    """
    by_date = {row["date"]: row for row in rows}
    for prefix, series in (("views", collection.views), ("clones", collection.clones)):
        stamp_key = f"{prefix}_collected_at"
        for day, (total, unique) in series.days.items():
            row = by_date.setdefault(day, {key: "" for key in HEADERS[DAILY]} | {"date": day})
            if row[stamp_key] > collection.stamp:
                continue
            values = {prefix: str(total), f"unique_{prefix}": str(unique)}
            if any(row[key] != value for key, value in values.items()):
                row.update(values)
                row[stamp_key] = collection.stamp
    return [by_date[day] for day in sorted(by_date)]


def snapshot_row(collection: Collection) -> Row:
    reported = sorted(set(collection.views.days) | set(collection.clones.days))
    return {
        "snapshot_date": collection.snapshot_date,
        "collected_at": collection.stamp,
        "first_date": reported[0] if reported else "",
        "last_date": reported[-1] if reported else "",
        "views_14d": str(collection.views.count),
        "unique_views_14d": str(collection.views.uniques),
        "clones_14d": str(collection.clones.count),
        "unique_clones_14d": str(collection.clones.uniques),
    } | {column: str(value) for column, value in collection.repository.items()}


def on_day(rows: Sequence[Row], day: str) -> list[Row]:
    return [row for row in rows if row["snapshot_date"] == day]


def replace_day(rows: Sequence[Row], day: str, new: Sequence[Row]) -> list[Row]:
    """Replace one snapshot date's rows; other dates are never touched."""
    kept = [row for row in rows if row["snapshot_date"] != day]
    return sorted([*kept, *new],
                  key=lambda row: (row["snapshot_date"], int(row.get("rank") or 0)))


def merge_day(current: Mapping[str, list[Row]], collection: Collection) -> dict[str, list[Row]]:
    """Merge a collection day's snapshot, referrers and paths as one unit.

    The snapshot's ``collected_at`` is the time of the collection that last
    changed any of the three. A later collection replaces the unit only if
    something in it differs; an earlier one replayed afterwards never does.
    """
    day = collection.snapshot_date
    unit = {name: current[name] for name in (SNAPSHOTS, REFERRERS, PATHS)}
    stored = on_day(current[SNAPSHOTS], day)
    if stored and stored[0]["collected_at"] > collection.stamp:
        return unit
    new = {
        SNAPSHOTS: [snapshot_row(collection)],
        REFERRERS: [{"snapshot_date": day} | row for row in collection.referrers],
        PATHS: [{"snapshot_date": day} | row for row in collection.paths],
    }
    ignore = {"collected_at"}
    unchanged = (
        bool(stored)
        and all(stored[0][key] == new[SNAPSHOTS][0][key]
                for key in HEADERS[SNAPSHOTS] if key not in ignore)
        and on_day(current[REFERRERS], day) == new[REFERRERS]
        and on_day(current[PATHS], day) == new[PATHS]
    )
    if unchanged:
        return unit
    return {name: replace_day(current[name], day, rows) for name, rows in new.items()}


def gaps(daily: Sequence[Mapping[str, str]]) -> list[tuple[str, str]]:
    """Date ranges between the first and last stored day that no collection reported."""
    days = sorted(date.fromisoformat(row["date"]) for row in daily)
    found = []
    for before, after in zip(days, days[1:], strict=False):
        if after - before > timedelta(days=1):
            found.append(((before + timedelta(days=1)).isoformat(),
                          (after - timedelta(days=1)).isoformat()))
    return found


def compact(value: int) -> str:
    """A badge number: 950, 1.2k, 11k, 3.4M."""
    for limit, suffix in ((1_000_000, "M"), (1_000, "k")):
        if value >= limit:
            return f"{value / limit:.1f}".removesuffix(".0") + suffix
    return str(value)


def total(rows: Sequence[Mapping[str, str]], key: str) -> int:
    return sum(int(row[key]) for row in rows if row[key])


def badges(daily: Sequence[Mapping[str, str]],
           snapshots: Sequence[Mapping[str, str]]) -> dict[str, str]:
    """shields.io endpoint files for the README, from the stored data only.

    clones: all clones stored since tracking began. visitors (14d): unique
    visitors over the API's window, from the latest snapshot; daily uniques
    cannot be summed into a unique count. visits: daily unique visitors summed
    since tracking began, so a visitor counts once per day of visit.
    """
    latest = max(snapshots, key=lambda row: row["snapshot_date"], default=None)
    values = {
        "clones": ("clones", total(daily, "clones")),
        "visitors": ("visitors (14d)", int(latest["unique_views_14d"]) if latest else 0),
        "visits": ("visits", total(daily, "unique_views")),
    }
    return {f"{BADGES}/{name}.json": json.dumps(
                {"schemaVersion": 1, "label": label, "message": compact(value),
                 "color": BADGE_COLOR}) + "\n"
            for name, (label, value) in values.items()}


def merge(raw: Path, data: Path, collected_at: datetime) -> Report:
    """Validate everything, then write only the files whose content changed."""
    collection = load_collection(raw, collected_at)
    current = {name: read_history(data / name) for name in HEADERS}
    merged = {DAILY: merge_daily([dict(row) for row in current[DAILY]], collection)}
    merged.update(merge_day(current, collection))
    report = Report(gaps=gaps(merged[DAILY]))
    contents = {name: render(name, rows) for name, rows in merged.items()}
    contents.update(badges(merged[DAILY], merged[SNAPSHOTS]))
    pending = {}
    for name, content in contents.items():
        path = data / name
        if not path.exists() or path.read_text(encoding="utf-8") != content:
            pending[name] = content
            report.changed[name] = len(merged[name]) if name in merged else 1
    write_all(data, pending)
    return report


def write_all(data: Path, contents: Mapping[str, str]) -> None:
    """Write every file to a temporary sibling first, then move them into place."""
    data.mkdir(parents=True, exist_ok=True)
    staged = []
    for name, content in contents.items():
        target = data / name
        target.parent.mkdir(parents=True, exist_ok=True)
        temporary = target.with_name(f".{target.name}.tmp")
        temporary.write_text(content, encoding="utf-8", newline="")
        staged.append((temporary, target))
    for temporary, target in staged:
        os.replace(temporary, target)


def initialize(data: Path) -> list[str]:
    """Create the empty history files; existing files are never overwritten."""
    created = [name for name in HEADERS if not (data / name).exists()]
    write_all(data, {name: render(name, []) for name in created})
    return created


# --- Traffic alert ---------------------------------------------------------------

# A 14-day total that at least doubles and grows by this much is a spike.
SPIKE_GROWTH = {"views_14d": 100, "clones_14d": 10}
NEW_REFERRER_VIEWS = 5
ALERT_DAYS = 7


def number(row: Mapping[str, str], key: str) -> int | None:
    return int(row[key]) if row[key] else None


def alert(data: Path, milestone_due: str) -> dict[str, str] | None:
    """The alert issue for the latest snapshot, or None when nothing stands out.

    It compares the latest snapshot with the one before: a views or clones
    spike, more stars or forks, or a referrer never seen before that brings at
    least NEW_REFERRER_VIEWS views. Counts missing from the older snapshot
    (collected before they were recorded) raise nothing.
    """
    snapshots = sorted(read_history(data / SNAPSHOTS), key=lambda row: row["snapshot_date"])
    if len(snapshots) < 2:
        return None
    previous, current = snapshots[-2], snapshots[-1]
    day = current["snapshot_date"]
    reasons = []
    for key, growth in SPIKE_GROWTH.items():
        before, now = int(previous[key]), int(current[key])
        if now >= max(before * 2, before + growth):
            reasons.append(f"{key.removesuffix('_14d').capitalize()} spike")
    for key in ("stars", "forks"):
        before_count, now_count = number(previous, key), number(current, key)
        if before_count is not None and now_count is not None and now_count > before_count:
            reasons.append(f"{key.capitalize()} increased")
    referrers = read_history(data / REFERRERS)
    known = {row["referrer"] for row in referrers if row["snapshot_date"] < day}
    today = on_day(referrers, day)
    new = [row for row in today
           if row["referrer"] not in known and int(row["count"]) >= NEW_REFERRER_VIEWS]
    if new:
        reasons.append("New referrer")
    if not reasons:
        return None

    def ranked(rows: Sequence[Row], label: str) -> list[str]:
        return [f"- {row[label]}: {row['count']} views, {row['uniques']} unique"
                for row in rows] or ["- None reported."]

    def compared(key: str, name: str) -> str:
        return (f"| {name} | {current[key] or 'not recorded'} "
                f"| {previous[key] or 'not recorded'} |")

    target = (date.fromisoformat(day) + timedelta(days=ALERT_DAYS)).isoformat()
    body = [
        "## Objective", "",
        f"Find out what drove the traffic recorded on {day}: "
        + ", ".join(reasons).lower() + ".", "",
        "## Context", "",
        "Opened by the *Traffic history* workflow (docs/TRAFFIC.md, *Traffic alert*).", "",
        f"| 14-day snapshot | {day} | {previous['snapshot_date']} |", "| --- | --- | --- |",
        compared("views_14d", "Views"), compared("unique_views_14d", "Unique visitors"),
        compared("clones_14d", "Clones"), compared("unique_clones_14d", "Unique cloners"),
        compared("stars", "Stars"), compared("forks", "Forks"), "",
        f"**New referrers** (at least {NEW_REFERRER_VIEWS} views):", "", *ranked(new, "referrer"), "",
        "**Top referrers:**", "", *ranked(today, "referrer"), "",
        "**Popular paths:**", "", *ranked(on_day(read_history(data / PATHS), day), "path"), "",
        "## Acceptance criteria", "",
        "- [ ] The sources are checked in **Insights → Traffic**: which referrer (a blog, "
        "forum, newsletter or another project) and which page drew the visits.",
        "- [ ] What was found, and any follow-up issue, is recorded in a Resolution.", "",
        "## Dependencies and blocking status", "", "None.", "",
        "## Schedule", "",
        f"- Planned start: **{day}**.", f"- Target completion: **{target}**.",
        f"- Milestone target: **{milestone_due}**.", "",
        "UTC collection date; dates are not evidence of completion.",
    ]
    return {"title": f"Traffic alert: {day}", "body": "\n".join(body) + "\n"}


# --- Command line ----------------------------------------------------------------


def summary(report: Report, collected_at: datetime) -> str:
    stamp = collected_at.isoformat(timespec="seconds").replace("+00:00", "Z")
    lines = ["# Traffic history", "", f"Collection: {stamp}", ""]
    if report.changed:
        lines += ["| File | Rows |", "| --- | --- |"]
        lines += [f"| `{name}` | {rows} |" for name, rows in sorted(report.changed.items())]
    else:
        lines.append("No change: the history already holds this collection.")
    if report.gaps:
        lines += ["", "Unavailable dates (never collected): "
                  + ", ".join(start if start == end else f"{start} to {end}"
                              for start, end in report.gaps)]
    return "\n".join(lines) + "\n"


def parse_time(value: str) -> datetime:
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        raise argparse.ArgumentTypeError(f"invalid timestamp {value!r}") from error
    if parsed.tzinfo is None:
        raise argparse.ArgumentTypeError("the timestamp needs a UTC offset")
    return parsed.astimezone(UTC)


def parse_args(argv: Sequence[str] | None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    commands = parser.add_subparsers(dest="command", required=True)
    merge_command = commands.add_parser("merge", help="merge one collection into the history")
    merge_command.add_argument("--raw", type=Path, required=True,
                               help=f"directory holding {', '.join(RESPONSES)}")
    merge_command.add_argument("--data", type=Path, required=True,
                               help="history directory (data/ on traffic-history)")
    merge_command.add_argument("--collected-at", type=parse_time,
                               default=datetime.now(UTC).replace(microsecond=0),
                               help="UTC time of the collection (default: now)")
    merge_command.add_argument("--summary", type=Path, help="append a Markdown report here")
    init_command = commands.add_parser("init", help="create empty history files")
    init_command.add_argument("--data", type=Path, required=True)
    alert_command = commands.add_parser(
        "alert", help="write the alert issue for the latest snapshot, if one is due")
    alert_command.add_argument("--data", type=Path, required=True)
    alert_command.add_argument("--milestone-due", type=date.fromisoformat, required=True,
                               help="due date of the milestone the issue goes into")
    alert_command.add_argument("--output", type=Path, required=True,
                               help="JSON file for the issue's title and body; "
                                    "not written when no alert is due")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    if args.command == "init":
        created = initialize(args.data)
        print(f"created: {', '.join(created) or 'nothing (all files exist)'}")
        return 0
    if args.command == "alert":
        try:
            issue = alert(args.data, args.milestone_due.isoformat())
        except TrafficError as error:
            print(f"ERROR: {error}", file=sys.stderr)
            return 2
        if issue is None:
            print("No alert: nothing stands out in the latest snapshot.")
            return 0
        args.output.write_text(json.dumps(issue), encoding="utf-8")
        print(issue["title"])
        return 0
    try:
        report = merge(args.raw, args.data, args.collected_at)
    except TrafficError as error:
        print(f"ERROR: {error}; the history was not changed", file=sys.stderr)
        return 2
    text_report = summary(report, args.collected_at)
    print(text_report, end="")
    if args.summary:
        with args.summary.open("a", encoding="utf-8") as handle:
            handle.write(text_report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
