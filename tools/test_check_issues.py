"""Tests for the issue triage rules (tools/check_issues.py)."""

from __future__ import annotations

import contextlib
import io
import json
import shutil
import tempfile
import unittest
import unittest.mock
from pathlib import Path

import check_issues
from check_issues import Issue

ROOT = Path(__file__).resolve().parents[1]
REPOSITORY = "example-org/sample-project"
BLOCKED_LABEL = check_issues.BLOCKED
READY = "## Dependencies and blocking status\n\n**Ready to start**: none.\n"
SCHEDULE = "\n## Schedule\n\n- Planned start: **2026-11-04**.\n- Target completion: **2026-11-10**.\n"
DUE = "2026-12-20T00:00:00Z"


def issue(number: int, *labels: str, state: str = "open", body: str = READY,
          assignees: tuple[str, ...] = ("triage-person",), milestone: str | None = "v0.1.0",
          schedule: str = SCHEDULE, milestone_state: str | None = "open",
          due: str | None = DUE) -> Issue:
    """An issue that passes every rule unless an argument says otherwise; ``schedule``
    is appended to ``body``."""
    return Issue(number, state, labels or ("P3",), assignees, milestone, body + schedule,
                 milestone_state if milestone else None, due if milestone else None)


def schedule(target: str) -> str:
    return f"\n## Schedule\n\n- Target completion: **{target}**.\n"


def form_target(value: str) -> str:
    """The target end date field as GitHub renders an issue form answer."""
    return f"\n### 📅 Target end date\n\n{value}\n\n### 🔗 Other field\n\nx\n"


def blocked_by(*numbers: int, heading: str = "## Dependencies and blocking status") -> str:
    lines = "\n".join(f"- #{n} — prerequisite" for n in numbers)
    return f"## Objective\n\nx\n\n{heading}\n\n**blocked**:\n\n{lines}\n\n## Tracking\n\n#99\n"


def rules(issues: list[Issue], states: dict[int, str] | None = None,
          pulls: frozenset[int] = frozenset()) -> list[tuple[int, str]]:
    known = states if states is not None else {i.number: i.state for i in issues}
    return [(f.number, f.rule)
            for f in check_issues.audit(issues, known, REPOSITORY, pulls=pulls, assignee="triage-person")]


class MetadataTests(unittest.TestCase):
    def test_compliant_issue_has_no_findings(self) -> None:
        self.assertEqual(rules([issue(1)]), [])

    def test_assignee_priority_and_milestone_are_required(self) -> None:
        self.assertEqual(rules([issue(1, assignees=())]), [(1, "assignee")])
        self.assertEqual(rules([issue(1, "documentation")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, "P2", "P3")]), [(1, "priority")])
        self.assertEqual(rules([issue(1, milestone=None)]), [(1, "milestone")])

    def test_p1_is_reserved_for_bugs(self) -> None:
        self.assertEqual(rules([issue(1, "P1", "documentation")]), [(1, "p1-reserved")])
        self.assertEqual(rules([issue(1, "P1", "bug")]), [])

    def test_closed_issues_are_not_audited(self) -> None:
        self.assertEqual(rules([issue(1, "documentation", state="closed", assignees=())]), [])

    def test_unknown_labels_are_reported_against_the_catalogue(self) -> None:
        findings = check_issues.audit([issue(1, "P3", "ready")], {1: "open"}, REPOSITORY,
                                      catalogue={"P3"}, assignee="triage-person")
        self.assertEqual([(f.number, f.rule) for f in findings], [(1, "catalogue")])


class DependencyTests(unittest.TestCase):
    def test_blocked_by_an_open_prerequisite(self) -> None:
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), issue(1)]), [])

    def test_blocked_with_only_closed_prerequisites_must_be_unblocked(self) -> None:
        closed = issue(1, state="closed")
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), closed]),
                         [(2, "blocked")])

    def test_open_prerequisite_requires_blocked(self) -> None:
        self.assertEqual(rules([issue(2, "P3", body=blocked_by(1)), issue(1)]), [(2, "blocked")])

    def test_external_condition_counts_as_a_blocker(self) -> None:
        body = "### 🔗 Dependencies and blocking status\n\nExternal condition: Excel host access\n"
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=body)]), [])
        self.assertEqual(rules([issue(2, "P3", body=body)]), [(2, "blocked")])

    def test_references_outside_the_section_are_ignored(self) -> None:
        # blocked_by() also mentions #99 under Tracking; it must not count.
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(1)), issue(1)],
                               {1: "open", 2: "open"}), [])

    def test_issue_urls_and_form_headings_are_recognised(self) -> None:
        body = ("### 🔗 Dependencies and blocking status\n\n"
                f"- https://github.com/{REPOSITORY}/issues/1\n")
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=body), issue(1)]), [])

    def test_missing_section_and_unknown_reference(self) -> None:
        self.assertEqual(rules([issue(1, body="## Objective\n\nx\n")]), [(1, "dependencies")])
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(7))]),
                         [(2, "dependencies"), (2, "blocked")])

    def test_self_reference_is_ignored(self) -> None:
        self.assertEqual(rules([issue(2, "P3", body=blocked_by(2))]), [])

    def test_named_pull_request_is_a_reference_not_a_prerequisite(self) -> None:
        body = "## Dependencies and blocking status\n\nNone. #1 was delivered in #22.\n"
        closed = issue(1, state="closed")
        self.assertEqual(rules([issue(2, "P3", body=body), closed], pulls=frozenset({22})), [])
        # A pull request never justifies blocked, and unknown numbers are still reported.
        self.assertEqual(rules([issue(2, "P3", BLOCKED_LABEL, body=blocked_by(22, 7))],
                               pulls=frozenset({22})),
                         [(2, "dependencies"), (2, "blocked")])


class ScheduleTests(unittest.TestCase):
    def test_closed_milestone_is_reported(self) -> None:
        self.assertEqual(rules([issue(1, milestone_state="closed")]), [(1, "milestone-closed")])
        self.assertEqual(rules([issue(1, milestone_state="open")]), [])

    def test_missing_target_date(self) -> None:
        self.assertEqual(rules([issue(1, schedule="")]), [(1, "target-date-missing")])
        # A Schedule section without the line, an unanswered form field, or the line
        # outside the Schedule section (here in the scope) state no date either.
        self.assertEqual(rules([issue(1, schedule="\n## Schedule\n\nPlanned start: **2026-11-04**.\n")]),
                         [(1, "target-date-missing")])
        self.assertEqual(rules([issue(1, schedule=form_target("_No response_"))]),
                         [(1, "target-date-missing")])
        self.assertEqual(rules([issue(1, schedule="\n## Scope\n\n- Target completion: **2026-11-10**.\n")]),
                         [(1, "target-date-missing")])

    def test_invalid_target_date(self) -> None:
        for value in ("2026-02-30", "2026-13-01", "26-11-10", "2026-11-10T00:00", "TBD", ""):
            with self.subTest(value=value):
                self.assertEqual(rules([issue(1, schedule=schedule(value))]), [(1, "target-date-invalid")])
        self.assertEqual(rules([issue(1, schedule=form_target("10/11/2026"))]),
                         [(1, "target-date-invalid")])
        self.assertEqual(rules([issue(1, schedule=form_target("2026-02-30"))]),
                         [(1, "target-date-invalid")])

    def test_target_date_after_the_milestone_due_date(self) -> None:
        on_due = "\n## Schedule\n\nStart: **2026-12-14**. Target completion: **2026-12-20**. More.\n"
        self.assertEqual(rules([issue(1, schedule=schedule("2026-12-21"))]), [(1, "target-date-after-due")])
        self.assertEqual(rules([issue(1, schedule=on_due)]), [])
        # Only the date part of the due date counts.
        self.assertEqual(rules([issue(1, schedule=on_due, due="2026-12-20T23:59:59Z")]), [])
        self.assertEqual(rules([issue(1, schedule=on_due, due="2026-12-19T23:59:59Z")]),
                         [(1, "target-date-after-due")])
        self.assertEqual(rules([issue(1, schedule=form_target("2026-12-21"))]),
                         [(1, "target-date-after-due")])

    def test_without_a_due_date_only_presence_and_validity_are_checked(self) -> None:
        self.assertEqual(rules([issue(1, schedule=schedule("2099-01-01"), due=None)]), [])
        self.assertEqual(rules([issue(1, schedule="", due=None)]), [(1, "target-date-missing")])
        self.assertEqual(rules([issue(1, schedule=schedule("2026-02-30"), due=None)]),
                         [(1, "target-date-invalid")])
        self.assertEqual(rules([issue(1, schedule=schedule("2099-01-01"), milestone=None)]),
                         [(1, "milestone")])
        self.assertEqual(rules([issue(1, schedule=form_target("2026-02-30"), milestone=None)]),
                         [(1, "milestone"), (1, "target-date-invalid")])

    def test_schedule_heading_may_carry_a_qualifier(self) -> None:
        for heading in ("## Schedule — baseline 2026-10-29", "### 📅 Schedule", "## Schedule"):
            with self.subTest(heading=heading):
                body = f"\n{heading}\n\n- Target completion: **2026-12-21**.\n"
                self.assertEqual(check_issues.target_date(body), "2026-12-21")
        # Only the first word counts: these are not the Schedule section.
        for heading in ("## Dependency-schedule audit", "## Scheduled work"):
            with self.subTest(heading=heading):
                self.assertIsNone(check_issues.target_date(f"\n{heading}\n\n- Target completion: **2026-12-21**.\n"))

    def test_form_field_supplies_the_date_without_a_schedule_line(self) -> None:
        self.assertEqual(rules([issue(1, schedule=form_target("2026-11-10"))]), [])
        self.assertEqual(check_issues.target_date(form_target("  2026-11-10  ")), "2026-11-10")
        self.assertIsNone(check_issues.target_date("## Objective\n\nx\n"))

    def test_schedule_line_takes_precedence_over_the_form_field(self) -> None:
        self.assertEqual(rules([issue(1, schedule=form_target("2026-12-21") + SCHEDULE)]), [])
        self.assertEqual(rules([issue(1, schedule=SCHEDULE + form_target("not a date"))]), [])
        self.assertEqual(rules([issue(1, schedule=form_target("2026-11-10") + schedule("2026-12-21"))]),
                         [(1, "target-date-after-due")])
        self.assertEqual(rules([issue(1, schedule=form_target("2026-11-10") + schedule("2026-02-30"))]),
                         [(1, "target-date-invalid")])

    def test_closed_issues_are_not_checked(self) -> None:
        self.assertEqual(rules([issue(1, state="closed", schedule="", milestone_state="closed")]), [])

    def test_malformed_due_date_stops_the_audit(self) -> None:
        with self.assertRaisesRegex(ValueError, "not ISO 8601"):
            rules([issue(1, due="soon")])

    def test_issue_from_api_reads_the_milestone_state_and_due_date(self) -> None:
        item: dict[str, object] = {
            "number": 1, "state": "OPEN", "labels": [], "assignees": [], "body": None,
            "milestone": {"title": "v0.1.0", "state": "closed", "due_on": DUE}}
        parsed = check_issues.issue_from_api(item)
        self.assertEqual((parsed.milestone, parsed.milestone_state, parsed.milestone_due),
                         ("v0.1.0", "closed", DUE))
        item["milestone"] = {"title": "v0.1.0", "state": "open", "due_on": None}
        self.assertIsNone(check_issues.issue_from_api(item).milestone_due)
        item["milestone"] = None
        parsed = check_issues.issue_from_api(item)
        self.assertEqual((parsed.milestone, parsed.milestone_state, parsed.milestone_due),
                         (None, None, None))


class ConfigTests(unittest.TestCase):
    def test_repository_configuration_passes(self) -> None:
        self.assertEqual(check_issues.config_findings(ROOT), [])

    def test_unsupported_overlays_are_rejected_instead_of_ignored(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copytree(ROOT / ".github/scripts", root / ".github/scripts")
            manifest = json.loads((ROOT / ".github/labels.json").read_text())
            manifest["overlays"] = {"selected": [{"name": "extra"}]}
            (root / ".github/labels.json").write_text(json.dumps(manifest))
            with self.assertRaisesRegex(ValueError, "root must contain exactly"):
                check_issues.catalogue_labels(root)

    def test_assignee_comes_from_configuration_not_repository_identity(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            forms = root / ".github/ISSUE_TEMPLATE"
            forms.mkdir(parents=True)
            (forms / "bug.yml").write_text('assignees: [triage-person]\n')
            self.assertEqual(check_issues.configured_assignee(root), "triage-person")
            (forms / "bug.yml").write_text('assignees: []\n')
            with self.assertRaises(ValueError):
                check_issues.configured_assignee(root)

    @staticmethod
    def copy_configuration(root: Path) -> None:
        """The files config_findings reads, copied from the repository."""
        for folder in (".github/scripts", ".github/ISSUE_TEMPLATE"):
            shutil.copytree(ROOT / folder, root / folder)
        for name in (".github/labels.json", ".github/dependabot.yml", "docs/GOVERNANCE.md"):
            (root / name).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(ROOT / name, root / name)

    def test_label_guide_explains_every_catalogue_label_once(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.copy_configuration(root)
            guide = root / "docs/GOVERNANCE.md"
            text = guide.read_text(encoding="utf-8")
            row = next(line for line in text.splitlines() if line.startswith("| `wontfix` |"))
            guide.write_text(text.replace(row + "\n", ""), encoding="utf-8")
            self.assertIn("docs/GOVERNANCE.md must explain label 'wontfix' exactly once in the label tables "
                          "of Issues, found 0", check_issues.config_findings(root))
            guide.write_text(text.replace(row, row + "\n" + row), encoding="utf-8")
            self.assertIn("docs/GOVERNANCE.md must explain label 'wontfix' exactly once in the label tables "
                          "of Issues, found 2", check_issues.config_findings(root))
            guide.write_text(text.replace(row, row + "\n| `stale` | Not a label. |"), encoding="utf-8")
            self.assertIn("docs/GOVERNANCE.md explains label 'stale', which is not in the catalogue",
                          check_issues.config_findings(root))

    def test_documented_labels_reads_only_the_label_tables_of_issues(self) -> None:
        text = ("## Branches\n\n| Label | Use |\n| --- | --- |\n| `main` | x |\n\n"
                "## 🗂️ Issues\n\n| Priority | Use |\n| --- | --- |\n| `P1` | x |\n\nText `bug`.\n\n"
                "| Other | Use |\n| --- | --- |\n| `ci` | x |\n\n| Label | Use |\n| --- | --- |\n| `bug` | x |\n\n"
                "## Releases\n\n| Label | Use |\n| --- | --- |\n| `release` | x |\n")
        self.assertEqual(check_issues.documented_labels(text), ["P1", "bug"])

    def test_dependabot_labels_and_assignees_are_checked(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.copy_configuration(root)
            self.assertEqual(check_issues.config_findings(root), [])
            config = root / ".github/dependabot.yml"
            text = config.read_text(encoding="utf-8")
            config.write_text(text.replace("      - ci\n", "      - no-such-label\n", 1), encoding="utf-8")
            self.assertIn(".github/dependabot.yml sets unknown label 'no-such-label'",
                          check_issues.config_findings(root))
            config.write_text(text.replace("      - danielep71\n", "      - someone-else\n", 1), encoding="utf-8")
            self.assertIn(".github/dependabot.yml assigns 'someone-else', not 'danielep71'",
                          check_issues.config_findings(root))
            config.write_text(text.replace("    labels:\n      - ci\n      - P3\n",
                                           "    labels: [no-such-label] # custom labels\n", 1), encoding="utf-8")
            self.assertIn(".github/dependabot.yml sets unknown label 'no-such-label'",
                          check_issues.config_findings(root))

    def test_forms_require_the_target_end_date(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self.copy_configuration(root)
            form = root / ".github/ISSUE_TEMPLATE/enhancement.yml"
            text = form.read_text(encoding="utf-8")
            required = "      placeholder: YYYY-MM-DD\n    validations:\n      required: true\n"
            self.assertIn(required, text)
            prefix = "enhancement.yml: target_end_date must"
            cases = [
                (text.replace("id: target_end_date", "id: target_date"),
                 "enhancement.yml has no target_end_date field"),
                (text.replace(required, "      placeholder: YYYY-MM-DD\n"), f"{prefix} be required"),
                (text.replace(required, required.replace("true", "false")), f"{prefix} be required"),
                (text.replace("  - type: input\n    id: target_end_date",
                              "  - type: textarea\n    id: target_end_date"), f"{prefix} be an input"),
                (text.replace('label: "📅 Target end date"', 'label: "📅 Due"'),
                 f"{prefix} have the label 'Target end date'"),
            ]
            for changed, problem in cases:
                with self.subTest(problem=problem):
                    self.assertNotEqual(changed, text)
                    form.write_text(changed, encoding="utf-8")
                    self.assertIn(problem, check_issues.config_findings(root))
            form.write_text(text, encoding="utf-8")
            self.assertEqual(check_issues.config_findings(root), [])

    def test_form_field(self) -> None:
        text = ("body:\n  - type: input\n    id: a\n    attributes:\n      label: A\n"
                "  - type: textarea\n    id: b  # comment\n")
        self.assertEqual(check_issues.form_field(text, "a"),
                         ["  - type: input", "    id: a", "    attributes:", "      label: A"])
        self.assertEqual(check_issues.form_field(text, "b"), ["  - type: textarea", "    id: b  # comment"])
        self.assertIsNone(check_issues.form_field(text, "c"))

    def test_block_lists(self) -> None:
        text = ("updates:\n  - x: 1\n    labels:\n      - ci   # a comment\n\n      - \"P3\"\n"
                "    assignees: [triage-person]\n    other: y\n  - x: 2\n    labels:\n    - tests\n")
        self.assertEqual(check_issues.block_lists(text, "labels"), [["ci", "P3"], ["tests"]])
        self.assertEqual(check_issues.block_lists(text, "assignees"), [["triage-person"]])
        self.assertEqual(check_issues.block_lists(text, "missing"), [])
        commented = "labels: [no-such-label, P3]  # custom labels\nassignees: [someone-else] # owner\n"
        self.assertEqual(check_issues.block_lists(commented, "labels"), [["no-such-label", "P3"]])
        self.assertEqual(check_issues.block_lists(commented, "assignees"), [["someone-else"]])

    def test_flow_list(self) -> None:
        text = 'name: x\nlabels: ["bug", "P2"]\nassignees: [triage-person]\n'
        self.assertEqual(check_issues.flow_list(text, "labels"), ["bug", "P2"])
        self.assertEqual(check_issues.flow_list(text, "assignees"), ["triage-person"])
        self.assertEqual(check_issues.flow_list(text, "title"), [])


class MainTests(unittest.TestCase):
    def test_input_mode_reports_and_exits_nonzero_on_findings(self) -> None:
        items = [
            {"number": 1, "state": "open", "labels": [{"name": "P3"}],
             "assignees": [{"login": "triage-person"}], "milestone": None,
             "body": READY + "Delivered in #2, a pull request.\n" + SCHEDULE},
            {"number": 3, "state": "open", "labels": [{"name": "P3"}],
             "assignees": [{"login": "triage-person"}],
             "milestone": {"title": "v0.1.0", "state": "closed", "due_on": "2026-11-01T07:00:00Z"},
             "body": READY + SCHEDULE},
            {"number": 2, "state": "open", "labels": [], "assignees": [],
             "milestone": None, "body": READY, "pull_request": {}},
        ]
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "issues.json"
            source.write_text(json.dumps(items), encoding="utf-8")
            output = Path(directory) / "report.json"
            with contextlib.redirect_stdout(io.StringIO()) as stdout, \
                    unittest.mock.patch.object(check_issues, "configured_assignee",
                                               return_value="triage-person"):
                code = check_issues.main(["--root", str(ROOT), "--input", str(source),
                                          "--output", str(output), "--repo-name", REPOSITORY])
            self.assertEqual(code, 1)
            report = json.loads(output.read_text(encoding="utf-8"))
            self.assertEqual(report["status"], "fail")
            self.assertEqual([(f["number"], f["rule"]) for f in report["findings"]],
                             [(1, "milestone"), (3, "milestone-closed"), (3, "target-date-after-due")])
            self.assertIn("open issues audited: 2", stdout.getvalue())

    def test_offline_mode_requires_explicit_repository(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), self.assertRaises(SystemExit) as error:
            check_issues.parse_args(["--input", "issues.json"])
        self.assertEqual(error.exception.code, 2)

    def test_live_mode_requires_a_token(self) -> None:
        with contextlib.redirect_stderr(io.StringIO()), \
                unittest.mock.patch.dict("os.environ", {"GITHUB_TOKEN": ""}):
            self.assertEqual(check_issues.main(["--repository", REPOSITORY]), 2)


if __name__ == "__main__":
    unittest.main()
