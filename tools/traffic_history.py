#!/usr/bin/env python3
"""Merge GitHub traffic API responses into the CSV history of the traffic-history branch.

The traffic API keeps 14 days. Each collection is merged by date and metric:
the newest collection wins for every date it reports, dates it does not report
keep their stored values, and nothing is ever summed across collections. A
date that no collection reported has no row: it is unavailable, not zero.

The merge reads only local files and needs no token (docs/TRAFFIC.md). Every
response and every existing file is validated before anything is written, so
an invalid or incomplete collection leaves the history unchanged.
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
                "views_14d", "unique_views_14d", "clones_14d", "unique_clones_14d"),
    REFERRERS: ("snapshot_date", "rank", "referrer", "count", "uniques"),
    PATHS: ("snapshot_date", "rank", "path", "title", "count", "uniques"),
}
RESPONSES = ("views.json", "clones.json", "referrers.json", "paths.json")
MIDNIGHT = re.compile(r"^(\d{4}-\d{2}-\d{2})T00:00:00Z$")
ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
# Collection times are always written in this form, so they compare as strings.
STAMP = re.compile(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$")
NUMERIC = {"views", "unique_views", "clones", "unique_clones", "views_14d", "unique_views_14d",
           "clones_14d", "unique_clones_14d", "rank", "count", "uniques"}
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


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as error:
        raise TrafficError(f"{path.name}: response missing") from error
    except (OSError, ValueError) as error:
        raise TrafficError(f"{path.name}: unreadable response ({error})") from error


def load_collection(raw: Path, collected_at: datetime) -> Collection:
    """Validate the four responses of one collection; all of them or none."""
    collected = collected_at.date()
    return Collection(
        collected_at=collected_at,
        views=parse_series(load_json(raw / "views.json"), "views", collected),
        clones=parse_series(load_json(raw / "clones.json"), "clones", collected),
        referrers=parse_ranked(load_json(raw / "referrers.json"), "referrers", ("referrer",)),
        paths=parse_ranked(load_json(raw / "paths.json"), "paths", ("path", "title")),
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
    if tuple(reader.fieldnames or ()) != header:
        raise TrafficError(f"{path.name}: header is {reader.fieldnames}, expected {list(header)}")
    rows = []
    for number, row in enumerate(reader, start=2):
        if None in row or any(value is None for value in row.values()):
            raise TrafficError(f"{path.name}:{number}: wrong number of fields")
        if not ISO_DATE.match(row[header[0]]):
            raise TrafficError(f"{path.name}:{number}: invalid date {row[header[0]]!r}")
        check_fields(path.name, number, row)
        rows.append(dict(row))
    return rows


def check_fields(name: str, number: int, row: Mapping[str, str]) -> None:
    """Counts are digits and collection times have their fixed form.

    A daily metric may be blank together with its collection time: that metric
    was never reported for the date.
    """
    def blank_metric(key: str) -> bool:
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
    }


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


def merge(raw: Path, data: Path, collected_at: datetime) -> Report:
    """Validate everything, then write only the files whose content changed."""
    collection = load_collection(raw, collected_at)
    current = {name: read_history(data / name) for name in HEADERS}
    merged = {DAILY: merge_daily([dict(row) for row in current[DAILY]], collection)}
    merged.update(merge_day(current, collection))
    report = Report(gaps=gaps(merged[DAILY]))
    pending = {}
    for name, rows in merged.items():
        content = render(name, rows)
        path = data / name
        if not path.exists() or path.read_text(encoding="utf-8") != content:
            pending[name] = content
            report.changed[name] = len(rows)
    write_all(data, pending)
    return report


def write_all(data: Path, contents: Mapping[str, str]) -> None:
    """Write every file to a temporary sibling first, then move them into place."""
    data.mkdir(parents=True, exist_ok=True)
    staged = []
    for name, content in contents.items():
        temporary = data / f".{name}.tmp"
        temporary.write_text(content, encoding="utf-8", newline="")
        staged.append((temporary, data / name))
    for temporary, target in staged:
        os.replace(temporary, target)


def initialize(data: Path) -> list[str]:
    """Create the empty history files; existing files are never overwritten."""
    created = [name for name in HEADERS if not (data / name).exists()]
    write_all(data, {name: render(name, []) for name in created})
    return created


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
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    if args.command == "init":
        created = initialize(args.data)
        print(f"created: {', '.join(created) or 'nothing (all files exist)'}")
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
