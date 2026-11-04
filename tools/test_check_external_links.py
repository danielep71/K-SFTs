"""Tests of check_external_links.py against a fake fetcher; no network is used."""

from __future__ import annotations

import contextlib
import io
import tempfile
import threading
import unittest
from collections.abc import Sequence
from pathlib import Path
from unittest import mock

import check_external_links as cel
from check_external_links import Reply


class FakeFetcher:
    """Answers each URL from a script of replies; the last one repeats."""

    def __init__(self, script: dict[str, Sequence[Reply]]) -> None:
        self.script = {url: list(replies) for url, replies in script.items()}
        self.calls: list[str] = []
        self.lock = threading.Lock()

    def __call__(self, url: str, timeout: float) -> Reply:
        with self.lock:
            self.calls.append(url)
            replies = self.script.get(url, [Reply(None, reason="unknown host")])
            return replies.pop(0) if len(replies) > 1 else replies[0]


class Sleeps:
    def __init__(self) -> None:
        self.delays: list[float] = []

    def __call__(self, delay: float) -> None:
        self.delays.append(delay)


def outcome(script: dict[str, Sequence[Reply]], url: str) -> tuple[cel.Outcome, FakeFetcher,
                                                                   Sleeps]:
    fetch, sleeps = FakeFetcher(script), Sleeps()
    return cel.check_url(url, fetch, sleeps), fetch, sleeps


class ClassificationTests(unittest.TestCase):
    def test_success_is_verified(self) -> None:
        for status in (200, 204):
            result, _, _ = outcome({"https://a.test/": [Reply(status, reason="OK")]},
                                   "https://a.test/")
            self.assertEqual(result.verdict, cel.VERIFIED)

    def test_not_verifiable_statuses(self) -> None:
        for status in (403, 429, 500, 503):
            result, _, _ = outcome({"https://a.test/": [Reply(status)]}, "https://a.test/")
            self.assertEqual(result.verdict, cel.UNVERIFIABLE, status)

    def test_network_failure_and_timeout_are_not_verifiable(self) -> None:
        for reason in ("URLError: Name or service not known", "TimeoutError: timed out"):
            result, _, _ = outcome({"https://a.test/": [Reply(None, reason=reason)]},
                                   "https://a.test/")
            self.assertEqual((result.verdict, result.detail), (cel.UNVERIFIABLE, reason))

    def test_other_statuses_are_broken(self) -> None:
        for status in (400, 404, 410, 451):
            result, fetch, _ = outcome({"https://a.test/": [Reply(status, reason="Gone")]},
                                       "https://a.test/")
            self.assertEqual(result.verdict, cel.BROKEN, status)
            self.assertEqual(len(fetch.calls), 1, "a broken answer is not retried")

    def test_redirect_without_location_is_broken(self) -> None:
        result, _, _ = outcome({"https://a.test/": [Reply(302)]}, "https://a.test/")
        self.assertEqual(result.verdict, cel.BROKEN)


class RetryTests(unittest.TestCase):
    def test_transient_failure_is_retried_with_backoff(self) -> None:
        result, fetch, sleeps = outcome(
            {"https://a.test/": [Reply(503), Reply(None, reason="reset"), Reply(200)]},
            "https://a.test/")
        self.assertEqual((result.verdict, result.attempts), (cel.VERIFIED, 3))
        self.assertEqual(sleeps.delays, [cel.BACKOFF, 2 * cel.BACKOFF])

    def test_at_most_two_retries(self) -> None:
        result, fetch, sleeps = outcome({"https://a.test/": [Reply(429)]}, "https://a.test/")
        self.assertEqual((result.verdict, result.attempts), (cel.UNVERIFIABLE, 3))
        self.assertEqual(len(fetch.calls), 1 + cel.RETRIES)
        self.assertEqual(len(sleeps.delays), cel.RETRIES)

    def test_forbidden_is_not_retried(self) -> None:
        result, fetch, _ = outcome({"https://a.test/": [Reply(403)]}, "https://a.test/")
        self.assertEqual((result.verdict, len(fetch.calls)), (cel.UNVERIFIABLE, 1))


class RedirectTests(unittest.TestCase):
    def test_redirects_are_followed(self) -> None:
        result, fetch, _ = outcome({
            "https://a.test/old": [Reply(301, "https://b.test/mid#frag")],
            "https://b.test/mid": [Reply(302, "../new")],
            "https://b.test/new": [Reply(200)],
        }, "https://a.test/old")
        self.assertEqual((result.verdict, result.final_url), (cel.VERIFIED, "https://b.test/new"))
        self.assertEqual(fetch.calls, ["https://a.test/old", "https://b.test/mid",
                                       "https://b.test/new"])

    def test_redirect_to_a_missing_page_is_broken(self) -> None:
        result, _, _ = outcome({"https://a.test/": [Reply(308, "https://a.test/gone")],
                                "https://a.test/gone": [Reply(404, reason="Not Found")]},
                               "https://a.test/")
        self.assertEqual((result.verdict, result.detail), (cel.BROKEN, "404 Not Found"))

    def test_redirect_loop_is_not_verifiable(self) -> None:
        result, fetch, _ = outcome({"https://a.test/": [Reply(302, "https://a.test/")]},
                                   "https://a.test/")
        self.assertEqual(result.verdict, cel.UNVERIFIABLE)
        self.assertIn("redirects", result.detail)
        self.assertEqual(len(fetch.calls), (cel.MAX_REDIRECTS + 1) * (1 + cel.RETRIES))

    def test_redirects_share_one_time_budget(self) -> None:
        now = [0.0]
        timeouts: list[float] = []

        def slow(url: str, timeout: float) -> Reply:
            timeouts.append(timeout)
            now[0] += 15.0
            return Reply(302, "https://a.test/next")

        reply, _ = cel.follow("https://a.test/", slow, lambda: now[0])
        self.assertEqual(timeouts, [cel.TIMEOUT, cel.TIMEOUT - 15.0])
        self.assertEqual((reply.status, reply.reason), (None, "timed out after 25 s"))


class CollectionTests(unittest.TestCase):
    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)

    def write(self, path: str, text: str) -> None:
        (self.root / path).parent.mkdir(parents=True, exist_ok=True)
        (self.root / path).write_text(text, encoding="utf-8")

    def test_self_urls(self) -> None:
        for url in ("https://github.com/danielep71/K-SFTs",
                    "https://github.com/danielep71/k-sfts/issues/7",
                    "https://github.com/danielep71/K-SFTs.git",
                    "https://raw.githubusercontent.com/danielep71/K-SFTs/main/README.md",
                    "https://github.com/danielep71/K-SFTs/blob/main/README.md#x",
                    "https://api.github.com/repos/danielep71/K-SFTs",
                    "https://danielep71.github.io/K-SFTs/"):
            self.assertTrue(cel.is_self(url), url)
        for url in ("https://github.com/danielep71", "https://github.com/danielep71/K-SFTs-x",
                    "https://img.shields.io/github/stars/danielep71/K-SFTs",
                    "https://danielep71.github.io/other/"):
            self.assertFalse(cel.is_self(url), url)

    def test_collect_skips_self_deduplicates_and_ignores_code(self) -> None:
        self.write("README.md", "[a](https://a.test/x#one) [b](https://a.test/x#two)\n"
                   "[self](https://github.com/danielep71/K-SFTs/issues/1)\n"
                   "`https://code.test/` [m](mailto:x@a.test) [r](docs/a.md)\n")
        self.write("docs/a.md", "See <https://b.test/> and https://c.test/page.\n")
        self.write("notes.txt", "https://ignored.test/\n")
        urls, skipped = cel.collect(self.root, ["README.md", "docs/a.md", "notes.txt"])
        self.assertEqual(urls, {"https://a.test/x": ["README.md:1", "README.md:1"],
                                "https://b.test/": ["docs/a.md:1"],
                                "https://c.test/page": ["docs/a.md:1"]})
        self.assertEqual(skipped, 1)

    def test_protocol_relative_urls_are_https(self) -> None:
        self.write("README.md", "[a](//a.test/x) [self](//github.com/danielep71/K-SFTs)\n")
        urls, skipped = cel.collect(self.root, ["README.md"])
        self.assertEqual(urls, {"https://a.test/x": ["README.md:1"]})
        self.assertEqual(skipped, 1)

    def test_report_and_exit_code(self) -> None:
        self.write("README.md", "[ok](https://ok.test/) [warn](https://warn.test/)\n")
        fetch = FakeFetcher({"https://ok.test/": [Reply(200, reason="OK")],
                             "https://warn.test/": [Reply(403, reason="Forbidden")],
                             "https://gone.test/": [Reply(410, reason="Gone")]})
        summary = self.root / "out" / "links.md"
        argv = ["--root", str(self.root), "--summary", str(summary)]
        with mock.patch("check_links.tracked_files", return_value=["README.md"]):
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                self.assertEqual(cel.main(argv, fetch, Sleeps()), 0)
            report = output.getvalue()
            self.assertEqual(summary.read_text(encoding="utf-8"), report)
            self.assertIn("Status: **WARN**", report)
            self.assertIn("| Not verifiable (warning) | 1 |", report)
            self.assertIn("- `https://warn.test/`: 403 Forbidden (README.md:1)", report)
            self.assertIn("## Verified (1)", report)

            self.write("README.md", "[gone](https://gone.test/) [ok](https://ok.test/)\n")
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(cel.main(argv, fetch, Sleeps()), 1)
            report = summary.read_text(encoding="utf-8")
        self.assertIn("Status: **FAIL**", report)
        self.assertIn("| Broken (error) | 1 |", report)
        self.assertIn("- `https://gone.test/`: 410 Gone (README.md:1)", report)
        self.assertIn("## Not verifiable (0)\n\nNone.", report)

    def test_user_agent_has_no_email_address(self) -> None:
        self.assertNotIn("@", cel.USER_AGENT)
        self.assertIn("K-SFTs", cel.USER_AGENT)


if __name__ == "__main__":
    unittest.main()
