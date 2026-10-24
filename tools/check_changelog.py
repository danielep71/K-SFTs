#!/usr/bin/env python3
"""Check CHANGELOG.md and VERSION against the rules of RELEASING.md.

Static: it reads the two files and never publishes, tags or compares a
release. CHANGELOG.md starts with an ``## [Unreleased]`` section; each release
heading is ``## [X.Y.Z] - YYYY-MM-DD`` with a real date, newest first by
version and date, with a link reference to its tag. VERSION does not exist
before the first dated release; from then on it holds one ``X.Y.Z`` line,
the newest release. Adapted from K-SACCR; see NOTICE.md.
"""

from __future__ import annotations

import argparse
import datetime
import re
import sys
from collections.abc import Sequence
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CHANGELOG = "CHANGELOG.md"
VERSION = "VERSION"
SEMVER = r"(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)"
VERSION_HEADING = re.compile(r"^## \[([^\]]+)\](.*)$", re.M)
RELEASE_SUFFIX = re.compile(r" - (\d{4}-\d{2}-\d{2})")


def valid_date(text: str) -> bool:
    try:
        datetime.date.fromisoformat(text)
    except ValueError:
        return False
    return True


def link_problems(text: str, label: str) -> list[str]:
    """The link reference of a release points at its tag's release page."""
    match = re.search(rf"^\[{re.escape(label)}\]: (\S+)$", text, re.M)
    if not match:
        return [f"{CHANGELOG}: missing link reference for [{label}]"]
    if label != "Unreleased" and not match.group(1).endswith(f"/releases/tag/v{label}"):
        return [f"{CHANGELOG}: the link reference for [{label}] must end with "
                f"/releases/tag/v{label}"]
    return []


def released_versions(text: str) -> list[str]:
    return [label for label, _ in VERSION_HEADING.findall(text) if label != "Unreleased"]


def check_changelog(text: str) -> list[str]:
    headings = VERSION_HEADING.findall(text)
    findings = []
    if not headings or headings[0] != ("Unreleased", ""):
        findings.append(f"{CHANGELOG}: the first version heading must be '## [Unreleased]'")
    seen: set[str] = set()
    releases: list[tuple[tuple[int, ...], str]] = []
    for label, suffix in headings:
        if label in seen:
            findings.append(f"{CHANGELOG}: duplicate heading [{label}]")
        seen.add(label)
        if label == "Unreleased":
            if suffix:
                findings.append(f"{CHANGELOG}: '## [Unreleased]' takes no date or suffix")
            continue
        dated = RELEASE_SUFFIX.fullmatch(suffix)
        if not re.fullmatch(SEMVER, label) or not dated:
            findings.append(f"{CHANGELOG}: release heading must be '## [X.Y.Z] - YYYY-MM-DD': "
                            f"[{label}]{suffix}")
            continue
        if not valid_date(dated.group(1)):
            findings.append(f"{CHANGELOG}: [{label}] date {dated.group(1)} is not a calendar date")
        releases.append((tuple(int(part) for part in label.split(".")), dated.group(1)))
        findings.extend(link_problems(text, label))
    for newer, older in zip(releases, releases[1:], strict=False):
        if newer[0] <= older[0] or newer[1] < older[1]:
            findings.append(f"{CHANGELOG}: releases must be listed newest first, by version "
                            "and date")
            break
    if "Unreleased" in seen:
        findings.extend(link_problems(text, "Unreleased"))
    return findings


def check_version(root: Path, released: list[str]) -> list[str]:
    """VERSION exists from the first release on and names the newest one."""
    path = root / VERSION
    if not path.is_file():
        return [f"{VERSION} is missing; {CHANGELOG} releases [{released[0]}]"] if released else []
    raw = path.read_text(encoding="utf-8")
    if not re.fullmatch(SEMVER + r"\n", raw):
        return [f"{VERSION} must hold one X.Y.Z line ending in a newline"]
    if not released:
        return [f"{VERSION} exists but {CHANGELOG} has no dated release; it is created "
                "when a release is prepared (RELEASING.md)"]
    if raw.strip() != released[0]:
        return [f"{VERSION} {raw.strip()} differs from the newest {CHANGELOG} release "
                f"[{released[0]}]"]
    return []


def check(root: Path) -> tuple[list[str], int]:
    path = root / CHANGELOG
    if not path.is_file():
        return [f"{CHANGELOG} is missing"], 0
    text = path.read_text(encoding="utf-8")
    released = released_versions(text)
    return check_changelog(text) + check_version(root, released), len(released)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    findings, released = check(args.root.resolve())
    for finding in findings:
        print(f"FAIL: {finding}")
    status = "FAIL" if findings else "PASS"
    print(f"changelog: {status} ({released} dated release(s); format and VERSION only, "
          "nothing published)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
