#!/usr/bin/env python3
"""Check the relative links and anchors of every Markdown file tracked by Git.

Offline: it reads files only and never opens a network connection
(tools/README.md, links). It extracts inline links and images
``[text](target)``, reference definitions ``[id]: target``, ``href`` and
``src`` attributes of ``<a>``, ``<img>`` and ``<source>`` tags, autolinks
``<scheme:...>`` and bare ``http(s)://`` URLs, outside fenced code blocks,
code spans and HTML comments; indented code blocks are not recognized. A relative target is resolved against the file's directory, or
against the repository root when it starts with ``/``. It fails when the
target leaves the repository, is not a tracked file or directory, is missing
from the working tree or is an empty file. A fragment on a Markdown target,
or a same-file ``#fragment``, must equal a GitHub heading slug of that file
(lowercase; letters, marks, digits, ``_``, ``-`` and spaces kept, spaces
turned into ``-``; a repeated slug suffixed ``-1``, ``-2``) or an explicit
``<a id>`` or ``<a name>`` anchor. Links with a scheme (``https:``,
``mailto:``) or starting with ``//`` are left to check_external_links.py.

Exit code 0 means no findings, 1 findings; each is printed as
``path:line: message``.
"""

from __future__ import annotations

import argparse
import bisect
import html
import posixpath
import re
import subprocess
import sys
import unicodedata
from collections.abc import Iterable, Iterator, Sequence
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import unquote

ROOT = Path(__file__).resolve().parents[1]

# A fence may follow blockquote and list-item markers on its line.
FENCE = re.compile(r"^(?:[ \t>]|[-*+](?=[ \t])|\d{1,9}[.)](?=[ \t]))*(`{3,}|~{3,})(.*)$")
COMMENT = re.compile(r"<!--.*?-->", re.S)
BACKTICKS = re.compile(r"`+")
PARAGRAPH_BREAK = re.compile(r"\n[ \t>]*\n")
ATX = re.compile(r"^ {0,3}(#{1,6})(?:[ \t]+(.*?))?(?:[ \t]+#+)?[ \t]*$")
# Blockquote markers and list-item markers before a heading on the same line.
CONTAINER = re.compile(r"^(?: {0,3}>[ \t]?| {0,3}(?:[-*+]|\d{1,9}[.)])[ \t]+)*")
SETEXT = re.compile(r"^ {0,3}(?:=+|-+)[ \t]*$")
NOT_PARAGRAPH = re.compile(r"^\s*(?:$|#|>|<|\||[-*+][ \t]|\d+[.)][ \t]|`{3}|~{3})")
REFERENCE = re.compile(
    r"^[ \t]*(?:>[ \t]*)*\[((?:[^\]\\\n]|\\.)+)\]:[ \t]*(<[^>\n]*>|\S+)", re.M)
HTML_LINK = re.compile(
    r"<(?:a|img|source)\b[^>]*?\s(?:href|src)\s*=\s*(?:\"([^\"]*)\"|'([^']*)')", re.I)
AUTOLINK = re.compile(r"<([A-Za-z][A-Za-z0-9+.-]{1,31}:[^\s<>]*)>")
BARE_URL = re.compile(r"(?<![\w/\[<\"'=])https?://[^\s<>]+", re.I)
ANCHOR = re.compile(r"<a\b[^>]*?\s(?:id|name)\s*=\s*(?:\"([^\"]*)\"|'([^']*)')", re.I)
SCHEME = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")
ESCAPED = re.compile(r"\\([!-/:-@\[-`{-~])")
PLACEHOLDER = 0xE000


@dataclass(frozen=True)
class Link:
    """One link target found in a Markdown file, with its 1-based line."""

    line: int
    target: str


def blank(text: str, start: int, end: int) -> str:
    """Replace text[start:end] with spaces, keeping newlines and offsets."""
    return text[:start] + re.sub(r"[^\n]", " ", text[start:end]) + text[end:]


def blank_fences(text: str) -> str:
    """Blank every fenced code block, fence lines included; an unclosed fence
    runs to the end of the file, as in CommonMark."""
    lines = text.splitlines(keepends=True)
    out: list[str] = []
    fence = ""
    for line in lines:
        match = FENCE.match(line.rstrip("\r\n"))
        if not fence:
            if match is None or (match.group(1)[0] == "`" and "`" in match.group(2)):
                out.append(line)
                continue
            fence = match.group(1)
        elif (match is not None and match.group(1)[0] == fence[0]
              and len(match.group(1)) >= len(fence) and not match.group(2).strip()):
            fence = ""
        out.append(re.sub(r"[^\n\r]", " ", line))
    return "".join(out)


def code_spans(text: str) -> Iterator[tuple[int, int, int, int]]:
    """(start, end, content start, content end) of each code span: a run of N
    backticks closed by the next run of exactly N in the same paragraph."""
    position = 0
    while (opening := BACKTICKS.search(text, position)) is not None:
        start = opening.start()
        if start > 0 and text[start - 1] == "\\":
            position = start + 1
            continue
        size = len(opening.group())
        limit = PARAGRAPH_BREAK.search(text, opening.end())
        stop = limit.start() if limit else len(text)
        closing = next((run for run in BACKTICKS.finditer(text, opening.end(), stop)
                        if len(run.group()) == size), None)
        if closing is None:
            position = opening.end()
            continue
        yield start, closing.end(), opening.end(), closing.start()
        position = closing.end()


def visible(text: str) -> str:
    """The text with code blocks, code spans and HTML comments blanked."""
    text = blank_fences(text)
    for start, end, _, _ in list(code_spans(text)):
        text = blank(text, start, end)
    for comment in list(COMMENT.finditer(text)):
        text = blank(text, comment.start(), comment.end())
    return text


def opens(text: str, close: int) -> bool:
    """Whether an unescaped `[` in the same paragraph matches the `]` at close."""
    depth, newline, content = 0, False, True
    for i in range(close - 1, -1, -1):
        char = text[i]
        if char == "\n":
            if newline and not content:
                return False
            newline, content = True, False
            continue
        content = content or not char.isspace()
        if i > 0 and text[i - 1] == "\\":
            continue
        if char == "]":
            depth += 1
        elif char == "[":
            if depth == 0:
                return True
            depth -= 1
    return False


def destination(text: str, start: int) -> tuple[int, str, int] | None:
    """The offset and destination of an inline link whose `(` ends at start,
    and the offset of its closing `)`; None when what follows is not a
    destination, optional title and `)`."""
    i = start
    while i < len(text) and text[i] in " \t\n":
        i += 1
    if text.startswith("<", i):
        end = text.find(">", i)
        if end < 0 or "\n" in text[i:end]:
            return None
        begin, target, i = i + 1, text[i + 1:end], end + 1
    else:
        begin, depth = i, 0
        while i < len(text) and not text[i].isspace():
            if text[i] == "\\":
                i += 1
            elif text[i] == "(":
                depth += 1
            elif text[i] == ")":
                if depth == 0:
                    break
                depth -= 1
            i += 1
        target = text[begin:i]
    while i < len(text) and text[i] in " \t\n":
        i += 1
    if i < len(text) and text[i] in "\"'(":
        closer = ")" if text[i] == "(" else text[i]
        end = text.find(closer, i + 1)
        if end < 0:
            return None
        i = end + 1
        while i < len(text) and text[i] in " \t\n":
            i += 1
    return (begin, target, i) if i < len(text) and text[i] == ")" else None


def line_of(newlines: Sequence[int], offset: int) -> int:
    return bisect.bisect_right(newlines, offset - 1) + 1


def extract_links(text: str) -> list[Link]:
    """Every link target of a Markdown text, in the order of the text."""
    clean = visible(text)
    newlines = [i for i, char in enumerate(clean) if char == "\n"]
    found: dict[int, str] = {}
    # Link titles are not links: no bare URL is taken from inside one.
    titles: list[tuple[int, int]] = []
    for match in re.finditer(r"\]\(", clean):
        if match.start() > 0 and clean[match.start() - 1] == "\\":
            continue
        inline = destination(clean, match.end())
        if inline is not None and inline[1] and opens(clean, match.start()):
            found[inline[0]] = rendered_destination(inline[1])
            titles.append((inline[0] + len(inline[1]), inline[2]))
    for match in REFERENCE.finditer(clean):
        if not match.group(1).startswith("^"):
            found[match.start(2)] = rendered_destination(match.group(2).strip("<>"))
            line_end = clean.find("\n", match.end())
            titles.append((match.end(), len(clean) if line_end < 0 else line_end))
    for pattern in (HTML_LINK, AUTOLINK):
        for match in pattern.finditer(clean):
            group = 1 if match.group(1) is not None else 2
            if match.group(group):
                found[match.start(group)] = html.unescape(match.group(group))
    for match in BARE_URL.finditer(clean):
        if not any(start <= match.start() < end for start, end in titles):
            found.setdefault(match.start(), trim_bare(match.group()))
    return [Link(line_of(newlines, offset), target) for offset, target in sorted(found.items())]


def trim_bare(url: str) -> str:
    """A bare URL without the trailing punctuation that GitHub leaves out of
    its automatic link: `?!.,:*_~`, and a `)` that closes nothing."""
    while url:
        if url[-1] in "?!.,:*_~'\"":
            url = url[:-1]
        elif url[-1] == ")" and url.count(")") > url.count("("):
            url = url[:-1]
        else:
            break
    return url


def rendered_destination(target: str) -> str:
    """A link destination as GitHub renders it: backslash escapes and
    character references decoded, `\\&amp;` kept as the literal `&amp;`."""
    return unprotect(html.unescape(protect(target)))


def protect(text: str) -> str:
    """Hide backslash-escaped punctuation from the markup rules below."""
    return ESCAPED.sub(lambda m: chr(PLACEHOLDER + ord(m.group(1))), text)


def unprotect(text: str) -> str:
    return "".join(chr(ord(c) - PLACEHOLDER) if PLACEHOLDER < ord(c) < PLACEHOLDER + 128
                   else c for c in text)


def inline_text(markup: str) -> str:
    """The rendered text of inline Markdown outside code spans."""
    text = protect(markup)
    text = re.sub(r"!\[[^\]]*\](?:\([^)]*\)|\[[^\]]*\])", "", text)
    text = re.sub(r"\[([^\]]*)\](?:\([^)]*\)|\[[^\]]*\])", r"\1", text)
    text = re.sub(r"<[^>]*>", "", text)
    text = re.sub(r"(?<![^\W_])_+|_+(?![^\W_])", "", text)
    return html.unescape(unprotect(text))


def heading_text(markup: str) -> str:
    """The rendered text of a heading: code span content kept as written."""
    parts: list[str] = []
    position = 0
    for start, end, content_start, content_end in code_spans(markup):
        parts.append(inline_text(markup[position:start]))
        content = markup[content_start:content_end].replace("\n", " ")
        if content.startswith(" ") and content.endswith(" ") and content.strip():
            content = content[1:-1]
        parts.append(content)
        position = end
    parts.append(inline_text(markup[position:]))
    return "".join(parts).strip()


def slugify(text: str) -> str:
    """GitHub's heading slug: lowercase; keep letters, marks, decimal digits,
    connector punctuation, `-` and spaces; turn each space into `-`."""
    kept = (char for char in text.lower()
            if char in "- " or unicodedata.category(char)[0] in "LM"
            or unicodedata.category(char) in ("Nd", "Pc"))
    return "".join(kept).replace(" ", "-")


def headings(text: str) -> Iterator[str]:
    """The text of each ATX and setext heading outside code and comments."""
    source = blank_fences(text)
    for comment in list(COMMENT.finditer(source)):
        source = blank(source, comment.start(), comment.end())
    lines = source.splitlines()
    for index, line in enumerate(lines):
        atx = ATX.match(CONTAINER.sub("", line, count=1))
        if atx:
            yield heading_text(atx.group(2) or "")
        elif (index > 0 and SETEXT.match(line) and not NOT_PARAGRAPH.match(lines[index - 1])
              and not ATX.match(lines[index - 1])):
            yield heading_text(lines[index - 1])


def anchors(text: str) -> set[str]:
    """Every fragment that reaches somewhere in a Markdown file on GitHub."""
    found: set[str] = set()
    counts: dict[str, int] = {}
    for title in headings(text):
        # GitHub's slugger: a taken slug gets the next free `-<n>` suffix of
        # its base, so `Foo`, `Foo`, `Foo-1` give `foo`, `foo-1`, `foo-1-1`.
        base = slug = slugify(title)
        while slug in found:
            counts[base] = counts.get(base, 0) + 1
            slug = f"{base}-{counts[base]}"
        found.add(slug)
    for match in ANCHOR.finditer(visible(text)):
        found.add(html.unescape(match.group(1) if match.group(1) is not None else match.group(2)))
    return found


def is_external(target: str) -> bool:
    return target.startswith("//") or SCHEME.match(target) is not None


class Repository:
    """The tracked files of a checkout, read lazily."""

    def __init__(self, root: Path, tracked: Iterable[str]) -> None:
        self.root = root
        self.files = set(tracked)
        self.directories = {"."} | {posixpath.dirname(path) or "." for path in self.files}
        for directory in list(self.directories):
            while directory not in (".", ""):
                directory = posixpath.dirname(directory) or "."
                self.directories.add(directory)
        self.cache: dict[str, set[str]] = {}

    def text(self, path: str) -> str:
        return (self.root / path).read_text(encoding="utf-8")

    def anchors(self, path: str) -> set[str]:
        if path not in self.cache:
            self.cache[path] = anchors(self.text(path))
        return self.cache[path]

    def target_problem(self, path: str) -> str | None:
        """Why a repository path cannot be a link target, or None."""
        if path in self.directories:
            return None if (self.root / path).is_dir() else "does not exist"
        disk = self.root / path
        if path not in self.files:
            return "is not tracked by Git" if disk.exists() else "does not exist"
        if not disk.is_file():
            return "does not exist"
        return "is an empty file" if disk.stat().st_size == 0 else None


def resolve(source: str, path: str) -> str | None:
    """The repository-relative path a link path names, or None if it escapes."""
    base = "" if path.startswith("/") else posixpath.dirname(source)
    joined = posixpath.normpath(posixpath.join(base, path.lstrip("/")))
    return None if joined == ".." or joined.startswith("../") else joined


def link_problem(repo: Repository, source: str, target: str) -> str | None:
    """Why a relative link of source is broken, or None."""
    reference, _, fragment = target.partition("#")
    path = unquote(reference.split("?", 1)[0])
    fragment = unquote(fragment)
    if path:
        resolved = resolve(source, path)
        if resolved is None:
            return f"`{target}` leaves the repository"
        problem = repo.target_problem(resolved)
        if problem:
            return f"`{target}`: {resolved} {problem}"
    else:
        resolved = source
    if fragment and resolved.lower().endswith(".md") and resolved in repo.files:
        if fragment not in repo.anchors(resolved):
            return f"`{target}`: no heading or anchor `{fragment}` in {resolved}"
    return None


def check(root: Path, tracked: Iterable[str]) -> list[str]:
    """Findings, as `path:line: message`, for the tracked Markdown files."""
    repo = Repository(root, tracked)
    findings = []
    for source in sorted(path for path in repo.files if path.endswith(".md")):
        if not (root / source).is_file():
            continue
        for link in extract_links(repo.text(source)):
            if is_external(link.target):
                continue
            problem = link_problem(repo, source, link.target)
            if problem:
                findings.append(f"{source}:{link.line}: {problem}")
    return findings


def tracked_files(root: Path) -> list[str]:
    listing = subprocess.run(["git", "-C", str(root), "ls-files", "-z"], capture_output=True,
                             check=True).stdout.decode("utf-8")
    return [path for path in listing.split("\0") if path]


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args(argv)
    root = args.root.resolve()
    tracked = tracked_files(root)
    findings = check(root, tracked)
    for finding in findings:
        print(finding)
    count = sum(1 for path in tracked if path.endswith(".md"))
    print(f"links: {'FAIL' if findings else 'PASS'} ({count} Markdown files, "
          f"{len(findings)} findings)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
