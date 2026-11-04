#!/usr/bin/env python3
"""Check the external links of every Markdown file tracked by Git, over the network.

Not a gate of check.py: it needs the network, and the `External link check`
workflow runs it weekly and on manual dispatch (tools/README.md, External
links). It collects every http(s) URL with the extraction of check_links.py,
skips this repository's own URLs, removes fragments and duplicates, and
requests each URL with GET, following redirects itself, with a 25-second
budget per attempt, up to two retries with backoff and at most three
requests at a time. A 2xx answer is verified; 403, 429, 5xx, a timeout or a
network failure is not verifiable, a warning and never a pass; any other
answer, such as 404 or 410, is broken. The Markdown report lists the three
classes with their counts, on standard output and with --summary in a file.
Exit code 1 means at least one broken URL, 0 otherwise.
"""

from __future__ import annotations

import argparse
import http.client
import sys
import time
import urllib.error
import urllib.request
from collections.abc import Callable, Iterable, Mapping, Sequence
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urldefrag, urljoin, urlsplit

import check_links

ROOT = check_links.ROOT
USER_AGENT = ("K-SFTs-external-link-check/1.0 "
              "(+https://github.com/danielep71/K-SFTs; weekly documentation link check)")
ACCEPT = "text/html,application/xhtml+xml,application/pdf;q=0.9,*/*;q=0.8"
TIMEOUT = 25.0
RETRIES = 2
BACKOFF = 5.0
WORKERS = 3
MAX_REDIRECTS = 10
SELF_REPOSITORY = "/danielep71/k-sfts"
SELF_HOSTS = {
    "github.com": SELF_REPOSITORY,
    "www.github.com": SELF_REPOSITORY,
    "raw.githubusercontent.com": SELF_REPOSITORY,
    "api.github.com": "/repos" + SELF_REPOSITORY,
    "danielep71.github.io": "/k-sfts",
}

VERIFIED = "verified"
UNVERIFIABLE = "not verifiable"
BROKEN = "broken"
CLASSES = (BROKEN, UNVERIFIABLE, VERIFIED)


@dataclass(frozen=True)
class Reply:
    """One HTTP exchange: the status, or None when no response arrived."""

    status: int | None
    location: str | None = None
    reason: str = ""


Fetcher = Callable[[str, float], Reply]


@dataclass(frozen=True)
class Outcome:
    url: str
    verdict: str
    detail: str
    final_url: str
    attempts: int


class NoRedirect(urllib.request.HTTPRedirectHandler):
    """Return redirects as HTTPError, so that check_url follows them itself."""

    def redirect_request(self, req: urllib.request.Request, fp: object, code: int, msg: str,
                         headers: object, newurl: str) -> None:
        return None


def http_fetch(url: str, timeout: float) -> Reply:
    """One GET without following redirects; the timeout bounds each socket
    operation, as urllib allows. The body is not read."""
    opener = urllib.request.build_opener(NoRedirect)
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": ACCEPT})
    try:
        with opener.open(request, timeout=timeout) as response:
            return Reply(int(response.status), reason=str(response.reason))
    except urllib.error.HTTPError as error:
        location = error.headers.get("Location") if 300 <= error.code < 400 else None
        error.close()
        return Reply(error.code, location, str(error.reason))
    except (urllib.error.URLError, http.client.HTTPException, OSError, ValueError) as error:
        reason = getattr(error, "reason", None) or error
        return Reply(None, reason=f"{type(error).__name__}: {reason}")


def classify(reply: Reply) -> str:
    if reply.status is None:
        return UNVERIFIABLE
    if 200 <= reply.status < 300:
        return VERIFIED
    if reply.status in (403, 429) or reply.status >= 500:
        return UNVERIFIABLE
    return BROKEN


def transient(reply: Reply) -> bool:
    """Worth a retry: no response, rate limiting or a server error; not a 403."""
    return reply.status is None or reply.status == 429 or reply.status >= 500


def follow(url: str, fetch: Fetcher, clock: Callable[[], float]) -> tuple[Reply, str]:
    """Request url and its redirects within one TIMEOUT budget."""
    deadline = clock() + TIMEOUT
    current = url
    for _ in range(MAX_REDIRECTS + 1):
        remaining = deadline - clock()
        if remaining <= 0:
            return Reply(None, reason=f"timed out after {TIMEOUT:g} s"), current
        reply = fetch(current, remaining)
        if reply.status is None or not 300 <= reply.status < 400 or not reply.location:
            return reply, current
        current = urldefrag(urljoin(current, reply.location)).url
    return Reply(None, reason=f"more than {MAX_REDIRECTS} redirects"), current


def describe(reply: Reply) -> str:
    if reply.status is None:
        return reply.reason or "no response"
    return f"{reply.status} {reply.reason}".strip()


def check_url(url: str, fetch: Fetcher, sleep: Callable[[float], None] = time.sleep,
              clock: Callable[[], float] = time.monotonic) -> Outcome:
    """Classify one URL, retrying a transient failure up to RETRIES times."""
    attempt = 0
    while True:
        attempt += 1
        reply, final = follow(url, fetch, clock)
        if not transient(reply) or attempt > RETRIES:
            return Outcome(url, classify(reply), describe(reply), final, attempt)
        sleep(BACKOFF * 2 ** (attempt - 1))


def is_self(url: str) -> bool:
    """A URL of this repository, its raw files or its website."""
    parts = urlsplit(url)
    prefix = SELF_HOSTS.get((parts.hostname or "").lower())
    if prefix is None:
        return False
    path = parts.path.lower().removesuffix(".git")
    return path == prefix or path.startswith(prefix + "/")


def collect(root: Path, tracked: Iterable[str]) -> tuple[dict[str, list[str]], int]:
    """External URLs without fragment -> where they appear; and how many links
    to this repository were skipped."""
    urls: dict[str, list[str]] = {}
    skipped = 0
    for source in sorted(path for path in tracked if path.endswith(".md")):
        if not (root / source).is_file():
            continue
        text = (root / source).read_text(encoding="utf-8")
        for link in check_links.extract_links(text):
            # GitHub serves pages over HTTPS, so `//host/path` is HTTPS.
            target = "https:" + link.target if link.target.startswith("//") else link.target
            if urlsplit(target).scheme.lower() not in ("http", "https"):
                continue
            if is_self(target):
                skipped += 1
                continue
            urls.setdefault(urldefrag(target).url, []).append(f"{source}:{link.line}")
    return urls, skipped


def render(outcomes: Sequence[Outcome], locations: Mapping[str, Sequence[str]], files: int,
           skipped: int) -> str:
    groups = {verdict: [o for o in outcomes if o.verdict == verdict] for verdict in CLASSES}
    status = ("FAIL" if groups[BROKEN] else "WARN" if groups[UNVERIFIABLE] else "PASS")
    lines = [
        "# External link check",
        "",
        f"Status: **{status}**",
        "",
        f"{len(outcomes)} distinct external URLs from {files} tracked Markdown files; "
        f"{skipped} links to this repository skipped.",
        "",
        "| Class | URLs |",
        "| --- | --- |",
        f"| Broken (error) | {len(groups[BROKEN])} |",
        f"| Not verifiable (warning) | {len(groups[UNVERIFIABLE])} |",
        f"| Verified | {len(groups[VERIFIED])} |",
    ]
    titles = {BROKEN: "Broken", UNVERIFIABLE: "Not verifiable", VERIFIED: "Verified"}
    for verdict in CLASSES:
        lines += ["", f"## {titles[verdict]} ({len(groups[verdict])})", ""]
        if not groups[verdict]:
            lines.append("None.")
        for outcome in groups[verdict]:
            where = ", ".join(locations.get(outcome.url, []))
            moved = f"; redirected to `{outcome.final_url}`" if outcome.final_url != outcome.url \
                else ""
            tries = f"; {outcome.attempts} attempts" if outcome.attempts > 1 else ""
            lines.append(f"- `{outcome.url}`: {outcome.detail}{moved}{tries} ({where})")
    return "\n".join(lines) + "\n"


def run(root: Path, tracked: Sequence[str], fetch: Fetcher,
        sleep: Callable[[float], None] = time.sleep) -> tuple[list[Outcome], str]:
    locations, skipped = collect(root, tracked)
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        outcomes = list(pool.map(lambda url: check_url(url, fetch, sleep), sorted(locations)))
    files = sum(1 for path in tracked if path.endswith(".md"))
    return outcomes, render(outcomes, locations, files, skipped)


def main(argv: Sequence[str] | None = None, fetch: Fetcher = http_fetch,
         sleep: Callable[[float], None] = time.sleep) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--summary", type=Path, help="also write the Markdown report here")
    args = parser.parse_args(argv)
    root = args.root.resolve()
    outcomes, report = run(root, check_links.tracked_files(root), fetch, sleep)
    sys.stdout.write(report)
    if args.summary is not None:
        args.summary.parent.mkdir(parents=True, exist_ok=True)
        args.summary.write_text(report, encoding="utf-8", newline="\n")
    return 1 if any(outcome.verdict == BROKEN for outcome in outcomes) else 0


if __name__ == "__main__":
    sys.exit(main())
