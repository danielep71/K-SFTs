"""Tests of check_links.py on fixture repositories in temporary directories."""

from __future__ import annotations

import contextlib
import io
import subprocess
import tempfile
import unittest
from pathlib import Path

import check_links as cl

ROOT = Path(__file__).resolve().parents[1]


class Fixture(unittest.TestCase):
    """A temporary repository whose files are all tracked unless stated."""

    def setUp(self) -> None:
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.untracked: set[str] = set()

    def write(self, path: str, text: str = "content\n", tracked: bool = True) -> None:
        file = self.root / path
        file.parent.mkdir(parents=True, exist_ok=True)
        file.write_text(text, encoding="utf-8")
        if not tracked:
            self.untracked.add(path)

    def tracked(self) -> list[str]:
        return sorted(path.relative_to(self.root).as_posix() for path in self.root.rglob("*")
                      if path.is_file() and path.relative_to(self.root).as_posix()
                      not in self.untracked)

    def findings(self) -> list[str]:
        return cl.check(self.root, self.tracked())

    def assertFindings(self, *fragments: str) -> None:
        findings = self.findings()
        self.assertEqual(len(findings), len(fragments), findings)
        for finding, fragment in zip(findings, fragments, strict=True):
            self.assertIn(fragment, finding)


class TargetTests(Fixture):
    def test_existing_targets_pass(self) -> None:
        self.write("docs/guide.md", "# Guide\n\n[up](../README.md) [root](/docs/guide.md)\n"
                   "[folder](../src/) [self](guide.md) [tool](../src/a.py#L3) [top](#)\n")
        self.write("README.md", "[guide](docs/guide.md \"title\") [spaced](<docs/guide.md>)\n")
        self.write("src/a.py")
        self.assertEqual(self.findings(), [])

    def test_missing_file(self) -> None:
        self.write("README.md", "Intro.\n\nSee [the guide](docs/missing.md).\n")
        self.assertFindings("README.md:3: `docs/missing.md`: docs/missing.md does not exist")

    def test_missing_directory(self) -> None:
        self.write("README.md", "[folder](nothing/)\n")
        self.assertFindings("nothing does not exist")

    def test_untracked_file(self) -> None:
        self.write("README.md", "[notes](notes.md)\n")
        self.write("notes.md", tracked=False)
        self.assertFindings("notes.md is not tracked by Git")

    def test_empty_file(self) -> None:
        self.write("README.md", "![logo](logo.png)\n")
        self.write("logo.png", "")
        self.assertFindings("README.md:1: `logo.png`: logo.png is an empty file")

    def test_escaping_path(self) -> None:
        self.write("docs/a.md", "[out](../../outside.md) [abs](/../x.md)\n")
        self.assertFindings("`../../outside.md` leaves the repository",
                            "`/../x.md` leaves the repository")

    def test_percent_encoded_path(self) -> None:
        self.write("README.md", "[file](my%20file.md)\n")
        self.write("my file.md")
        self.assertEqual(self.findings(), [])

    def test_external_links_are_ignored(self) -> None:
        self.write("README.md", "[a](https://example.invalid/x) [b](mailto:x@example.invalid) "
                   "[c](//example.invalid) <https://example.invalid/y> https://example.invalid/z\n")
        self.assertEqual(self.findings(), [])

    def test_html_attributes(self) -> None:
        self.write("README.md", '<img width="10"\n     src="missing.png">\n<a href="README.md">x</a>\n')
        self.assertFindings("README.md:2: `missing.png`")


class FragmentTests(Fixture):
    def test_missing_fragment(self) -> None:
        self.write("README.md", "# Title\n\n[a](docs/a.md#nowhere) [b](#title) [c](#absent)\n")
        self.write("docs/a.md", "# A\n")
        self.assertFindings("`docs/a.md#nowhere`: no heading or anchor `nowhere` in docs/a.md",
                            "`#absent`: no heading or anchor `absent` in README.md")

    def test_explicit_anchor(self) -> None:
        self.write("README.md", '<a id="quick-start"></a>\n\n## ⚡ Quick start\n\n'
                   "<a name='old'></a>\n[a](#quick-start) [b](#-quick-start) [c](#old)\n")
        self.assertEqual(self.findings(), [])

    def test_duplicate_headings_are_suffixed(self) -> None:
        self.write("README.md", "## Notes\n\n## Notes\n\n## Notes\n\n"
                   "[a](#notes) [b](#notes-1) [c](#notes-2) [d](#notes-3)\n")
        self.assertFindings("`#notes-3`")

    def test_headings_inside_blockquotes_and_lists(self) -> None:
        self.write("README.md", "> ## Note\n\n- ## Item\n\n1. > ### Deep\n\n"
                   "- plain item\n\n```\n> ## Fenced\n```\n\n"
                   "[a](#note) [b](#item) [c](#deep) [d](#plain-item) [e](#fenced)\n")
        self.assertFindings("`#plain-item`", "`#fenced`")

    def test_suffix_collisions_take_the_next_free_slug(self) -> None:
        self.write("README.md", "## Foo\n\n## Foo-1\n\n## Foo\n\n## Foo\n\n"
                   "[a](#foo) [b](#foo-1) [c](#foo-2) [d](#foo-3) [e](#foo-1-1)\n")
        self.assertFindings("`#foo-1-1`")
        self.write("README.md", "## Foo\n\n## Foo\n\n## Foo-1\n\n[a](#foo-1-1)\n")
        self.assertEqual(self.findings(), [])

    def test_emoji_and_punctuation_slug(self) -> None:
        self.write("README.md", "## 🧰 Repository tooling\n\n## ▶️ Running the checks\n\n"
                   "## A.B: c/d (e) — f_g, `check.py` & *em* [link](x.md)\n\n"
                   "Title é\n=======\n\n"
                   "[a](#-repository-tooling) [b](#%EF%B8%8F-running-the-checks) "
                   "[c](#ab-cd-e--f_g-checkpy--em-link) [d](#title-é) [e](#-running-the-checks)\n")
        self.write("x.md")
        self.assertFindings("`#-running-the-checks`")

    def test_slugify(self) -> None:
        self.assertEqual(cl.slugify("Hello, World!"), "hello-world")
        self.assertEqual(cl.slugify("⚙️ Settings"), "️-settings")
        self.assertEqual(cl.slugify("snake_case 2.0"), "snake_case-20")
        self.assertEqual(cl.heading_text("_em_ and __init__ and snake_case \\_x"),
                         "em and init and snake_case _x")

    def test_fragment_on_other_file_types_is_not_checked(self) -> None:
        self.write("README.md", "[a](tool.py#L10) [b](docs#x)\n")
        self.write("tool.py")
        self.write("docs/a.md")
        self.assertEqual(self.findings(), [])


class CodeTests(Fixture):
    def test_links_inside_code_are_ignored(self) -> None:
        self.write("README.md", "Inline `[a](missing1.md)` and ``[b](`missing2.md`)``.\n\n"
                   "```md\n[c](missing3.md)\n# Not a heading\n```\n\n"
                   "~~~\n[d](missing4.md)\n~~~\n\n"
                   "<!-- [e](missing5.md) -->\n\n"
                   "  ````\n  [f](missing6.md)\n  ```\n  ````\n"
                   "[g](#not-a-heading) [h](missing7.md)\n")
        self.assertFindings("README.md:18: `#not-a-heading`", "README.md:18: `missing7.md`")

    def test_fence_opened_on_a_list_item_line(self) -> None:
        self.write("README.md", "- ```md\n  [example](missing1.md)\n  ```\n\n"
                   "1. > ~~~\n   > [quoted](missing2.md)\n   > ~~~\n\n"
                   "[real](missing3.md)\n")
        self.assertFindings("README.md:9: `missing3.md`")

    def test_unclosed_backticks_do_not_hide_links(self) -> None:
        self.write("README.md", "A ` lone backtick [x](missing.md)\n")
        self.assertFindings("`missing.md`")

    def test_line_numbers_after_code(self) -> None:
        self.write("README.md", "```\n\n\n```\n`code`\n\n[x](missing.md)\n")
        self.assertFindings("README.md:7:")


class ReferenceTests(Fixture):
    def test_reference_style_links(self) -> None:
        self.write("README.md", "See [the guide][guide] and [notes].\n\n"
                   "[guide]: docs/guide.md#usage\n"
                   "[notes]: <docs/missing.md> \"Notes\"\n"
                   "  [anchor]: #intro\n"
                   "[^1]: A footnote, not a link.\n"
                   "[web]: https://example.invalid\n")
        self.write("docs/guide.md", "# Guide\n\n## Usage\n")
        self.assertFindings("README.md:4: `docs/missing.md`",
                            "README.md:5: `#intro`: no heading or anchor `intro`")


class EscapeTests(Fixture):
    def test_destinations_are_rendered_before_checking(self) -> None:
        self.write("README.md", "[a](a\\(b\\).md) [b](a&amp;b.md) [c](<a&#40;b).md>)\n\n"
                   "[d]: a\\&amp;b.md\n[e]: a&amp;b.md\n")
        self.write("a(b).md")
        self.write("a&b.md")
        self.assertFindings("README.md:3: `a&amp;b.md`")


class ExtractionTests(unittest.TestCase):
    def test_nested_and_titled_links(self) -> None:
        text = ("[![badge](https://img.example/b.svg)](#validation) "
                "[t](a_(b).md 'x') [u](c.md (y)) \\[not](d.md) [broken](e.md f g)\n")
        self.assertEqual([link.target for link in cl.extract_links(text)],
                         ["https://img.example/b.svg", "#validation", "a_(b).md", "c.md"])

    def test_link_titles_are_not_links(self) -> None:
        text = ('[a](https://a.test/ "See https://title.test/x") '
                "[b](<https://b.test/> 'https://title.test/y')\n"
                '[r]: https://r.test/ "https://title.test/z"\n'
                "After https://bare.test/.\n")
        self.assertEqual([link.target for link in cl.extract_links(text)],
                         ["https://a.test/", "https://b.test/", "https://r.test/",
                          "https://bare.test/"])

    def test_long_marker_lines_stay_linear(self) -> None:
        # Guards the fence and container patterns against backtracking.
        for line in (" " * 5000 + "x", " -" * 3000 + "x", "1. " * 3000 + "x", "> " * 3000):
            self.assertEqual(cl.extract_links(line + "\n"), [])
            self.assertEqual(list(cl.headings(line + "\n")), [])

    def test_bare_urls_are_trimmed(self) -> None:
        text = "See https://example.invalid/a. And (https://example.invalid/b).\n"
        self.assertEqual([link.target for link in cl.extract_links(text)],
                         ["https://example.invalid/a", "https://example.invalid/b"])


class MainTests(unittest.TestCase):
    def test_repository_passes(self) -> None:
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            code = cl.main(["--root", str(ROOT)])
        self.assertEqual(code, 0, output.getvalue())
        self.assertIn("links: PASS", output.getvalue())

    def test_git_repository_with_a_finding(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "README.md").write_text("[x](gone.md)\n", encoding="utf-8")
            (root / "gone.md").write_text("x\n", encoding="utf-8")
            subprocess.run(["git", "init", "-q", str(root)], check=True)
            subprocess.run(["git", "-C", str(root), "add", "README.md"], check=True)
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                code = cl.main(["--root", str(root)])
        self.assertEqual(code, 1)
        self.assertIn("README.md:1: `gone.md`: gone.md is not tracked by Git", output.getvalue())
        self.assertIn("links: FAIL (1 Markdown files, 1 findings)", output.getvalue())


if __name__ == "__main__":
    unittest.main()
