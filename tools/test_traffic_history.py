"""Tests for the traffic history merge (tools/traffic_history.py)."""

from __future__ import annotations

import contextlib
import csv
import io
import json
import tempfile
import unittest
from collections.abc import Iterator, Sequence
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any

import traffic_history
from traffic_history import DAILY, PATHS, REFERRERS, SNAPSHOTS, TrafficError

Day = tuple[str, int, int]


def window(last: str, values: Sequence[tuple[int, int]]) -> list[Day]:
    """Days ending on ``last``, oldest first, one per value pair."""
    end = date.fromisoformat(last)
    return [((end - timedelta(days=len(values) - 1 - offset)).isoformat(), total, unique)
            for offset, (total, unique) in enumerate(values)]


def series(key: str, days: Sequence[Day]) -> dict[str, Any]:
    return {
        "count": sum(total for _, total, _ in days),
        "uniques": max((unique for _, _, unique in days), default=0),
        key: [{"timestamp": f"{day}T00:00:00Z", "count": total, "uniques": unique}
              for day, total, unique in days],
    }


def at(stamp: str) -> datetime:
    return datetime.fromisoformat(stamp).replace(tzinfo=UTC)


class Workspace:
    """A raw response directory and a history directory under one temporary root."""

    def __init__(self, root: Path) -> None:
        self.raw = root / "raw"
        self.data = root / "data"
        self.raw.mkdir()

    def collect(self, views: Sequence[Day], clones: Sequence[Day] | None = None,
                referrers: Any = None, paths: Any = None) -> None:
        responses: dict[str, Any] = {
            "views.json": series("views", views),
            "clones.json": series("clones", views if clones is None else clones),
            "referrers.json": [{"referrer": "github.com", "count": 3, "uniques": 1}]
            if referrers is None else referrers,
            "paths.json": [{"path": "/o/r", "title": "r", "count": 3, "uniques": 1}]
            if paths is None else paths,
        }
        for name, body in responses.items():
            (self.raw / name).write_text(json.dumps(body), encoding="utf-8")

    def merge(self, stamp: str) -> traffic_history.Report:
        return traffic_history.merge(self.raw, self.data, at(stamp))

    def rows(self, name: str) -> list[dict[str, str]]:
        with (self.data / name).open(encoding="utf-8", newline="") as handle:
            return list(csv.DictReader(handle))

    def daily(self) -> dict[str, tuple[str, str, str]]:
        return {row["date"]: (row["views"], row["clones"], row["collected_at"])
                for row in self.rows(DAILY)}

    def snapshot(self) -> dict[str, str]:
        return {name: (self.data / name).read_text(encoding="utf-8")
                for name in traffic_history.HEADERS if (self.data / name).exists()}


@contextlib.contextmanager
def workspace() -> Iterator[Workspace]:
    with tempfile.TemporaryDirectory() as directory:
        yield Workspace(Path(directory))


class MergeTests(unittest.TestCase):
    def test_first_collection_creates_every_file(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            report = ws.merge("2026-10-14T06:17:00")
            self.assertEqual(sorted(report.changed), sorted(traffic_history.HEADERS))
            self.assertEqual(len(ws.rows(DAILY)), 14)
            snapshot = ws.rows(SNAPSHOTS)[0]
            self.assertEqual((snapshot["first_date"], snapshot["last_date"]),
                             ("2026-10-01", "2026-10-14"))
            self.assertEqual(snapshot["collected_at"], "2026-10-14T06:17:00Z")
            self.assertEqual(ws.rows(REFERRERS)[0]["rank"], "1")

    def test_overlapping_windows_keep_older_history_and_take_newer_values(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 13 + [(2, 1)]))
            ws.merge("2026-10-14T06:00:00")
            # A week later: 2026-10-14 is now complete (5 views), and the
            # window has moved on by seven days.
            ws.collect(window("2026-10-21", [(1, 1)] * 6 + [(5, 2)] + [(3, 1)] * 7))
            ws.merge("2026-10-21T06:00:00")
            daily = ws.daily()
            self.assertEqual(len(daily), 21)
            self.assertEqual(daily["2026-10-01"], ("1", "1", "2026-10-14T06:00:00Z"))
            self.assertEqual(daily["2026-10-14"], ("5", "5", "2026-10-21T06:00:00Z"))
            # An overlapping day with unchanged values keeps its first collection time.
            self.assertEqual(daily["2026-10-08"][2], "2026-10-14T06:00:00Z")
            self.assertEqual(len(ws.rows(SNAPSHOTS)), 2)

    def test_repeated_run_changes_nothing(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(2, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            before = ws.snapshot()
            report = ws.merge("2026-10-14T09:30:00")
            self.assertEqual(report.changed, {})
            self.assertEqual(ws.snapshot(), before)

    def test_same_day_run_with_new_values_replaces_that_day_only(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-13", [(1, 1)] * 14))
            ws.merge("2026-10-13T06:00:00")
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            ws.collect(window("2026-10-14", [(1, 1)] * 13 + [(4, 2)]),
                       referrers=[{"referrer": "example.org", "count": 4, "uniques": 2}])
            ws.merge("2026-10-14T18:00:00")
            snapshots = ws.rows(SNAPSHOTS)
            self.assertEqual([row["snapshot_date"] for row in snapshots],
                             ["2026-10-13", "2026-10-14"])
            self.assertEqual(snapshots[1]["collected_at"], "2026-10-14T18:00:00Z")
            referrers = ws.rows(REFERRERS)
            self.assertEqual([(row["snapshot_date"], row["referrer"]) for row in referrers],
                             [("2026-10-13", "github.com"), ("2026-10-14", "example.org")])

    def test_zero_activity_is_stored_and_uncollected_dates_are_unavailable(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(0, 0)] * 14))
            ws.merge("2026-10-14T06:00:00")
            ws.collect(window("2026-11-10", [(0, 0)] * 13 + [(1, 1)]))
            report = ws.merge("2026-11-10T06:00:00")
            daily = ws.daily()
            self.assertEqual(daily["2026-10-01"][:2], ("0", "0"))
            self.assertNotIn("2026-10-20", daily)
            self.assertEqual(report.gaps, [("2026-10-15", "2026-10-27")])
            self.assertIn("2026-10-15 to 2026-10-27",
                          traffic_history.summary(report, at("2026-11-10T06:00:00")))

    def test_metric_reported_for_one_series_only_leaves_the_other_blank(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14),
                       clones=window("2026-10-13", [(1, 1)] * 13))
            ws.merge("2026-10-14T06:00:00")
            self.assertEqual(ws.daily()["2026-10-14"][:2], ("1", ""))


class FailureTests(unittest.TestCase):
    def assert_refused(self, ws: Workspace, message: str) -> None:
        before = ws.snapshot()
        with self.assertRaisesRegex(TrafficError, message):
            ws.merge("2026-10-15T06:00:00")
        self.assertEqual(ws.snapshot(), before)

    def test_invalid_or_partial_collection_leaves_history_unchanged(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            ws.merge("2026-10-14T06:00:00")
            ws.collect(window("2026-10-15", [(1, 1)] * 14), paths={"message": "Bad credentials"})
            self.assert_refused(ws, "paths: expected a list")
            ws.collect(window("2026-10-15", [(1, 1)] * 14))
            (ws.raw / "referrers.json").unlink()
            self.assert_refused(ws, "referrers.json: response missing")
            (ws.raw / "referrers.json").write_text("{not json", encoding="utf-8")
            self.assert_refused(ws, "unreadable response")

    def test_malformed_values_are_rejected(self) -> None:
        cases = [
            ([{"timestamp": "2026-10-14T05:00:00Z", "count": 1, "uniques": 1}], "UTC midnight"),
            ([{"timestamp": "2026-10-16T00:00:00Z", "count": 1, "uniques": 1}], "after"),
            ([{"timestamp": "2026-10-14T00:00:00Z", "count": -1, "uniques": 1}], "non-negative"),
            ([{"timestamp": "2026-10-14T00:00:00Z", "count": True, "uniques": 1}], "non-negative"),
            ([{"timestamp": "2026-10-14T00:00:00Z", "count": 1, "uniques": 1}] * 2, "duplicate"),
        ]
        for views, message in cases:
            with self.subTest(message=message), workspace() as ws:
                ws.collect(window("2026-10-14", [(1, 1)]))
                (ws.raw / "views.json").write_text(
                    json.dumps({"count": 1, "uniques": 1, "views": views}), encoding="utf-8")
                self.assert_refused(ws, message)

    def test_corrupt_history_file_is_never_overwritten(self) -> None:
        for content, message in [("date,views\n2026-10-01,1\n", "header"),
                                 (traffic_history.render(DAILY, []) + "2026-10-01,x,1,1,1,t\n",
                                  "non-numeric"),
                                 (traffic_history.render(DAILY, []) + "yesterday,1,1,1,1,t\n",
                                  "invalid date"),
                                 (traffic_history.render(DAILY, []) + "2026-10-01,1\n",
                                  "number of fields")]:
            with self.subTest(message=message), workspace() as ws:
                ws.data.mkdir()
                (ws.data / DAILY).write_text(content, encoding="utf-8")
                ws.collect(window("2026-10-15", [(1, 1)] * 14))
                self.assert_refused(ws, message)


class CommandTests(unittest.TestCase):
    def test_init_creates_headers_and_keeps_existing_files(self) -> None:
        with workspace() as ws:
            self.assertEqual(traffic_history.initialize(ws.data), list(traffic_history.HEADERS))
            self.assertEqual(ws.rows(PATHS), [])
            (ws.data / DAILY).write_text(traffic_history.render(DAILY, [
                {"date": "2026-10-01", "views": "1", "unique_views": "1", "clones": "",
                 "unique_clones": "", "collected_at": "2026-10-01T06:00:00Z"}]),
                encoding="utf-8")
            self.assertEqual(traffic_history.initialize(ws.data), [])
            self.assertEqual(len(ws.rows(DAILY)), 1)

    def test_merge_command_reports_and_exits_two_on_refusal(self) -> None:
        with workspace() as ws:
            ws.collect(window("2026-10-14", [(1, 1)] * 14))
            summary = ws.raw.parent / "summary.md"
            arguments = ["merge", "--raw", str(ws.raw), "--data", str(ws.data),
                         "--collected-at", "2026-10-14T06:00:00Z", "--summary", str(summary)]
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(traffic_history.main(arguments), 0)
            self.assertIn("| `daily.csv` | 14 |", summary.read_text(encoding="utf-8"))
            (ws.raw / "clones.json").write_text("[]", encoding="utf-8")
            with contextlib.redirect_stderr(io.StringIO()) as stderr:
                self.assertEqual(traffic_history.main(arguments), 2)
            self.assertIn("the history was not changed", stderr.getvalue())

    def test_timestamps_need_an_offset(self) -> None:
        self.assertEqual(traffic_history.parse_time("2026-10-14T08:00:00+02:00"),
                         at("2026-10-14T06:00:00"))
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit):
            traffic_history.main(["merge", "--raw", ".", "--data", ".",
                                  "--collected-at", "2026-10-14T06:00:00"])


if __name__ == "__main__":
    unittest.main()
