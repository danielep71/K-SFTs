"""Tests of check_workflows.py, and of pr-milestones.yml's script against a fake API."""

from __future__ import annotations

import contextlib
import copy
import io
import os
import re
import unittest
from collections.abc import Callable
from pathlib import Path
from typing import Any
from unittest import mock

import check
import check_workflows as cw

ROOT = Path(__file__).resolve().parents[1]
YAML = cw.load_yaml()


def setUpModule() -> None:
    if YAML is None:
        if os.environ.get("GITHUB_ACTIONS"):
            raise RuntimeError("PyYAML is required on GitHub Actions "
                               "(tools/requirements-quality-ci.txt)")
        raise unittest.SkipTest("PyYAML is not installed; skipped locally only")


class Repository:
    """A mutable copy of the repository's workflows, Dependabot file and documents."""

    def __init__(self) -> None:
        workflows, dependabot, governance, settings = cw.load(ROOT, YAML)
        self.workflows: dict[str, Any] = copy.deepcopy(workflows)
        self.dependabot: Any = copy.deepcopy(dependabot)
        self.governance = governance
        self.settings = settings

    def job(self, workflow: str, job: str) -> Any:
        return self.workflows[workflow]["jobs"][job]

    def on(self, workflow: str) -> Any:
        return cw.triggers(self.workflows[workflow])

    def findings(self) -> list[str]:
        return cw.findings_for(self.workflows, self.dependabot, self.governance, self.settings)


class WorkflowInvariantTests(unittest.TestCase):
    def setUp(self) -> None:
        self.repo = Repository()

    def assertFinding(self, fragment: str) -> None:
        findings = self.repo.findings()
        self.assertTrue(any(fragment in finding for finding in findings),
                        f"no finding containing {fragment!r} in {findings}")

    def test_repository_passes(self) -> None:
        self.assertEqual(self.repo.findings(), [])

    def test_every_workflow_is_read(self) -> None:
        names = {path.name for path in (ROOT / cw.WORKFLOWS).glob("*.yml")}
        self.assertEqual(set(self.repo.workflows), names)
        self.assertIn(cw.MILESTONES_FILE, names)

    # Actions and checkout
    def test_unpinned_action(self) -> None:
        step = self.repo.job("repository-integrity.yml", "repository-integrity")["steps"][0]
        step["uses"] = "actions/checkout@v7"
        self.assertFinding("`actions/checkout@v7` is not pinned")

    def test_short_sha(self) -> None:
        self.repo.job("labels-drift.yml", "detect")["steps"][0]["uses"] = "actions/checkout@3d3c42e"
        self.assertFinding("not pinned to a 40-character SHA")

    def test_reusable_workflow_must_be_pinned(self) -> None:
        self.repo.workflows["labels-drift.yml"]["jobs"]["reuse"] = {
            "uses": "owner/repo/.github/workflows/x.yml@main", "permissions": {},
            "timeout-minutes": 5}
        self.assertFinding("job reuse: `owner/repo/.github/workflows/x.yml@main`")

    def test_checkout_persisting_credentials(self) -> None:
        step = self.repo.job("labels-sync.yml", "validate")["steps"][0]
        del step["with"]["persist-credentials"]
        self.assertFinding("actions/checkout without persist-credentials: false")
        step["with"]["persist-credentials"] = "false"
        self.assertFinding("actions/checkout without persist-credentials: false")

    # Permissions and timeouts
    def test_missing_workflow_permissions(self) -> None:
        del self.repo.workflows["issue-triage.yml"]["permissions"]
        self.assertFinding("issue-triage.yml: no workflow-level permissions")

    def test_workflow_level_write(self) -> None:
        self.repo.workflows["labels-sync.yml"]["permissions"] = {"contents": "write"}
        self.assertFinding("labels-sync.yml: workflow-level write permission ['contents']")

    def test_missing_timeout(self) -> None:
        del self.repo.job("traffic-history.yml", "alert")["timeout-minutes"]
        self.assertFinding("traffic-history.yml: job alert has no timeout-minutes")

    def test_extra_write_permission(self) -> None:
        job = self.repo.job("repository-integrity.yml", "repository-integrity")
        job["permissions"]["issues"] = "write"
        self.assertFinding("job repository-integrity has write permission ['issues']")

    def test_write_outside_the_allowed_scope(self) -> None:
        self.repo.job("traffic-history.yml", "alert")["permissions"]["contents"] = "write"
        self.assertFinding("job alert has write permission ['contents']")

    def test_write_all(self) -> None:
        self.repo.job("labels-sync.yml", "reconcile")["permissions"] = "write-all"
        self.assertFinding("job reconcile has write permission ['*']")

    def test_allowed_writes_pass(self) -> None:
        for (workflow, job), scopes in cw.ALLOWED_WRITES.items():
            granted = cw.writes(self.repo.job(workflow, job)["permissions"])
            self.assertEqual(granted, set(scopes), f"{workflow}/{job}")

    # The required check
    def test_required_job_with_if(self) -> None:
        self.repo.job("repository-integrity.yml", "repository-integrity")["if"] = "true"
        self.assertFinding("required job repository-integrity has `if`")

    def test_required_job_with_strategy(self) -> None:
        job = self.repo.job("repository-integrity.yml", "repository-integrity")
        job["strategy"] = {"matrix": {"python": ["3.13"]}}
        self.assertFinding("required job repository-integrity has `strategy`")

    def test_path_filter_on_pull_request(self) -> None:
        self.repo.on("repository-integrity.yml")["pull_request"]["paths"] = ["src/**"]
        self.assertFinding("`pull_request` filters are")

    def test_paths_ignore_on_push(self) -> None:
        self.repo.on("repository-integrity.yml")["push"]["paths-ignore"] = ["docs/**"]
        self.assertFinding("`push` filters are")

    def test_branch_filter_changed(self) -> None:
        self.repo.on("repository-integrity.yml")["pull_request"]["branches"] = ["main"]
        self.assertFinding("`pull_request` filters are")

    def test_required_check_renamed(self) -> None:
        self.repo.job("repository-integrity.yml", "repository-integrity")["name"] = "Integrity"
        self.assertFinding("0 jobs named `Repository integrity`")

    def test_rulesets_require_another_name(self) -> None:
        self.repo.settings = self.repo.settings.replace(
            "| Required status context | `Repository integrity` |",
            "| Required status context | `Integrity` |")
        self.assertFinding("the rulesets require ['Integrity', 'Repository integrity']")

    # labels-sync and pr-milestones
    def test_labels_sync_condition(self) -> None:
        self.repo.job("labels-sync.yml", "reconcile")["if"] = "github.event_name != 'pull_request'"
        self.assertFinding("reconcile's condition no longer checks the default branch")

    def test_milestones_events(self) -> None:
        self.repo.on("pr-milestones.yml")["pull_request"] = None
        self.assertFinding("pr-milestones.yml: events changed")

    def test_milestones_workflow_permissions(self) -> None:
        self.repo.workflows["pr-milestones.yml"]["permissions"] = {"contents": "read"}
        self.assertFinding("pr-milestones.yml: workflow-level permissions must be {}")

    def test_milestones_job_permissions(self) -> None:
        self.repo.job("pr-milestones.yml", "reconcile")["permissions"]["contents"] = "read"
        self.assertFinding("reconcile's permissions must be pull-requests: write")

    def test_milestones_condition(self) -> None:
        self.repo.job("pr-milestones.yml", "reconcile")["if"] = "always()"
        self.assertFinding("pr-milestones.yml: reconcile's condition must be")

    def test_milestones_second_step(self) -> None:
        steps = self.repo.job("pr-milestones.yml", "reconcile")["steps"]
        steps.insert(0, {"uses": "actions/checkout@" + "0" * 40,
                         "with": {"persist-credentials": False}})
        self.assertFinding("reconcile must have a single inline `run` step")

    def test_milestones_expression_in_code(self) -> None:
        step = self.repo.job("pr-milestones.yml", "reconcile")["steps"][0]
        step["run"] = step["run"].replace("import json", "import json  # ${{ github.head_ref }}")
        self.assertFinding("the inline code contains `${{`")

    # Dependabot
    def test_dependabot_duplicate_ecosystem(self) -> None:
        self.repo.dependabot["updates"].append(copy.deepcopy(self.repo.dependabot["updates"][1]))
        self.assertFinding("`pip` is covered 2 times")

    def test_dependabot_missing_ecosystem(self) -> None:
        del self.repo.dependabot["updates"][0]
        self.assertFinding("`github-actions` is covered 0 times")

    def test_dependabot_milestone(self) -> None:
        self.repo.dependabot["updates"][0]["milestone"] = 4
        self.assertFinding("`github-actions` sets a milestone")

    def test_dependabot_priorities(self) -> None:
        self.repo.dependabot["updates"][1]["labels"].append("P2")
        self.assertFinding("`pip` has priorities ['P2', 'P3']")
        self.repo.dependabot["updates"][1]["labels"] = ["ci"]
        self.assertFinding("`pip` has priorities []")


def quiet_main() -> int:
    with contextlib.redirect_stdout(io.StringIO()):
        return cw.main(["--root", str(ROOT)])


class MissingYamlTests(unittest.TestCase):
    def test_fails_on_github_actions(self) -> None:
        with mock.patch.object(cw, "load_yaml", return_value=None), \
                mock.patch.dict(os.environ, {"GITHUB_ACTIONS": "true"}):
            self.assertEqual(quiet_main(), 1)

    def test_skipped_locally(self) -> None:
        env = {key: value for key, value in os.environ.items() if key != "GITHUB_ACTIONS"}
        with mock.patch.object(cw, "load_yaml", return_value=None), \
                mock.patch.dict(os.environ, env, clear=True):
            self.assertEqual(quiet_main(), 2)

    def test_gate_runs_on_github_actions(self) -> None:
        with mock.patch("importlib.util.find_spec", return_value=None):
            with mock.patch.dict(os.environ, {"GITHUB_ACTIONS": "true"}):
                self.assertTrue(check.available("pyyaml"))
            env = {key: value for key, value in os.environ.items() if key != "GITHUB_ACTIONS"}
            with mock.patch.dict(os.environ, env, clear=True):
                self.assertFalse(check.available("pyyaml"))


REPOSITORY = "owner/repo"


class FakeGitHub:
    """The issues and pulls endpoints the milestone script uses, over live state."""

    def __init__(self, prs: dict[int, Any], issues: dict[int, Any]) -> None:
        self.items: dict[int, dict[str, Any]] = {}
        for number, milestone in prs.items():
            self.items[number] = {"number": number, "milestone": milestone, "pull_request": {}}
        for number, milestone in issues.items():
            self.items[number] = {"number": number, "milestone": milestone}
        self.patched: list[int] = []
        self.pages: list[int] = []

    def __call__(self, method: str, path: str, data: Any = None) -> Any:
        item = re.fullmatch(rf"/repos/{REPOSITORY}/issues/(\d+)", path)
        if item:
            number = int(item.group(1))
            if method == "PATCH":
                self.patched.append(number)
                self.items[number].update(data)
            return copy.deepcopy(self.items[number])
        if method == "GET" and path.startswith(f"/repos/{REPOSITORY}/pulls?"):
            # [?&]page= : a bare page= would match per_page= first.
            page = int(re.search(r"[?&]page=(\d+)", path).group(1))  # type: ignore[union-attr]
            size = int(re.search(r"[?&]per_page=(\d+)", path).group(1))  # type: ignore[union-attr]
            self.pages.append(page)
            prs = sorted(n for n, i in self.items.items() if "pull_request" in i)
            return [{"number": n, "milestone": self.items[n]["milestone"]}
                    for n in prs[(page - 1) * size:page * size]]
        raise AssertionError(f"unexpected request {method} {path}")

    def milestones(self, *, pulls: bool) -> dict[int, Any]:
        return {n: i["milestone"] for n, i in self.items.items()
                if ("pull_request" in i) == pulls}


class MilestoneScriptTests(unittest.TestCase):
    def setUp(self) -> None:
        workflow = Repository().workflows[cw.MILESTONES_FILE]
        source = cw.milestone_script(workflow)
        assert source is not None
        namespace: dict[str, Any] = {"__name__": "pr_milestones"}
        exec(compile(source, cw.MILESTONES_FILE, "exec"), namespace)
        self.script: Callable[[Any, str, Any], None] = namespace["reconcile"]

    def reconcile(self, api: FakeGitHub, repository: str, event: Any) -> None:
        with contextlib.redirect_stdout(io.StringIO()):
            self.script(api, repository, event)

    def test_pull_request_event_clears_its_milestone(self) -> None:
        api = FakeGitHub({5: {"number": 2}, 6: {"number": 2}}, {})
        self.reconcile(api, REPOSITORY, {"number": 5, "pull_request": {}})
        self.assertEqual(api.patched, [5])
        self.assertEqual(api.milestones(pulls=True), {5: None, 6: {"number": 2}})

    def test_pull_request_without_milestone_is_left_alone(self) -> None:
        api = FakeGitHub({5: None}, {})
        self.reconcile(api, REPOSITORY, {"number": 5, "pull_request": {}})
        self.assertEqual(api.patched, [])

    def test_scan_clears_every_page_and_no_issue(self) -> None:
        milestone = {"number": 4}
        prs = {n: (milestone if n % 40 == 7 else None) for n in range(1, 251)}
        issues = {n: milestone for n in range(251, 261)}
        api = FakeGitHub(prs, issues)
        self.reconcile(api, REPOSITORY, {"schedule": "17 3 * * *"})
        self.assertEqual(api.pages, [1, 2, 3])
        self.assertEqual(api.patched, [n for n in prs if prs[n] is not None])
        self.assertTrue(any(n > 200 for n in api.patched), "the third page was not cleared")
        self.assertTrue(all(value is None for value in api.milestones(pulls=True).values()))
        self.assertEqual(api.milestones(pulls=False), issues)

    def test_issue_number_is_refused(self) -> None:
        api = FakeGitHub({5: {"number": 2}}, {9: {"number": 2}})
        with self.assertRaisesRegex(ValueError, "Refusing to change an issue"):
            self.reconcile(api, REPOSITORY, {"number": 9, "pull_request": {}})
        self.assertEqual(api.patched, [])
        self.assertEqual(api.milestones(pulls=False), {9: {"number": 2}})


if __name__ == "__main__":
    unittest.main()
