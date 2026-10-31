#!/usr/bin/env python3
"""Check the security and governance invariants of the workflows and Dependabot.

Reads .github/workflows/*.yml and .github/dependabot.yml with PyYAML and
reports every finding (docs/GOVERNANCE.md; tools/README.md, workflow-invariants).
Read-only. Without PyYAML it fails when GITHUB_ACTIONS is set and exits 2
(skipped) otherwise, so that only a local run can skip it.
"""

from __future__ import annotations

import argparse
import importlib
import importlib.util
import os
import re
import sys
from collections import Counter
from collections.abc import Iterator, Mapping, Sequence
from pathlib import Path
from typing import Any

WORKFLOWS = Path(".github/workflows")
DEPENDABOT = Path(".github/dependabot.yml")
GOVERNANCE = Path("docs/GOVERNANCE.md")
SETTINGS = Path("docs/REPOSITORY_SETTINGS.md")

SHA_PINNED = re.compile(r"^[A-Za-z0-9_.-]+/[A-Za-z0-9_./-]+@[0-9a-f]{40}$")
DEFAULT_BRANCH_REF = "github.ref == format('refs/heads/{0}', github.event.repository.default_branch)"

# The only write permissions a job may hold: (workflow file, job id) -> scopes.
ALLOWED_WRITES: dict[tuple[str, str], frozenset[str]] = {
    ("labels-sync.yml", "reconcile"): frozenset({"issues"}),
    ("traffic-history.yml", "collect"): frozenset({"contents"}),
    ("traffic-history.yml", "alert"): frozenset({"issues"}),
    ("pr-milestones.yml", "reconcile"): frozenset({"pull-requests"}),
}

MILESTONES_FILE = "pr-milestones.yml"
MILESTONES_EVENTS: dict[str, Any] = {
    "push": {"branches": ["main"], "paths": [".github/workflows/pr-milestones.yml"]},
    "pull_request_target": {"types": ["opened", "reopened", "edited", "milestoned"]},
    "schedule": [{"cron": "17 3 * * *"}],
    "workflow_dispatch": None,
}
MILESTONES_IF = ("github.event_name == 'pull_request_target' || "
                 "github.event_name == 'schedule' || " + DEFAULT_BRANCH_REF)
SCRIPT = re.compile(r"<<'PY'\n(.*)\nPY\s*$", re.S)

# Ecosystems whose versions are pinned here, and where Dependabot must look.
PINNED_ECOSYSTEMS = {"github-actions": "/", "pip": "/tools"}
PRIORITY = re.compile(r"^P\d$")

Workflow = Mapping[Any, Any]


def load_yaml() -> Any | None:
    """PyYAML's module, or None when it is not installed."""
    if importlib.util.find_spec("yaml") is None:
        return None
    return importlib.import_module("yaml")


def strict_loader(yaml: Any) -> Any:
    """A SafeLoader that rejects a duplicate key, which PyYAML would silently
    resolve to its last value (a hidden `issues: write`, for example)."""

    class StrictLoader(yaml.SafeLoader):  # type: ignore[misc]
        def construct_mapping(self, node: Any, deep: bool = False) -> Any:
            keys = [self.construct_object(key, deep=deep) for key, _ in node.value]
            duplicates = sorted({str(key) for key in keys if keys.count(key) > 1})
            if duplicates:
                raise yaml.constructor.ConstructorError(
                    None, None, f"duplicate key(s) {duplicates}", node.start_mark)
            return super().construct_mapping(node, deep=deep)

    return StrictLoader


def triggers(workflow: Workflow) -> Any:
    """The `on:` mapping; YAML 1.1 reads an unquoted `on` key as True."""
    return workflow.get("on", workflow.get(True))


def squash(text: Any) -> str:
    return " ".join(str(text).split())


def jobs(workflow: Workflow) -> Iterator[tuple[str, Mapping[str, Any]]]:
    for job_id, job in (workflow.get("jobs") or {}).items():
        yield str(job_id), job if isinstance(job, Mapping) else {}


def writes(permissions: Any) -> set[str]:
    """The scopes a permissions value grants write access to; "*" for write-all."""
    if isinstance(permissions, str):
        return {"*"} if permissions == "write-all" else set()
    if isinstance(permissions, Mapping):
        return {str(scope) for scope, level in permissions.items() if level == "write"}
    return set()


def action_findings(name: str, workflow: Workflow) -> list[str]:
    findings = []
    for job_id, job in jobs(workflow):
        steps = job.get("steps") or []
        uses = [(f"job {job_id}", job["uses"])] if "uses" in job else []
        uses += [(f"job {job_id} step {i}", step["uses"]) for i, step in enumerate(steps, 1)
                 if isinstance(step, Mapping) and "uses" in step]
        for where, ref in uses:
            if not SHA_PINNED.match(str(ref)):
                findings.append(f"{name}: {where}: `{ref}` is not pinned to a 40-character SHA")
        for i, step in enumerate(steps, 1):
            if isinstance(step, Mapping) and str(step.get("uses", "")).startswith("actions/checkout@"):
                if (step.get("with") or {}).get("persist-credentials") is not False:
                    findings.append(f"{name}: job {job_id} step {i}: actions/checkout without "
                                    "persist-credentials: false")
    return findings


def permission_findings(name: str, workflow: Workflow) -> list[str]:
    findings = []
    if "permissions" not in workflow:
        findings.append(f"{name}: no workflow-level permissions")
    elif writes(workflow["permissions"]):
        findings.append(f"{name}: workflow-level write permission "
                        f"{sorted(writes(workflow['permissions']))}")
    for job_id, job in jobs(workflow):
        if "timeout-minutes" not in job:
            findings.append(f"{name}: job {job_id} has no timeout-minutes")
        extra = writes(job.get("permissions")) - ALLOWED_WRITES.get((name, job_id), frozenset())
        if extra:
            findings.append(f"{name}: job {job_id} has write permission {sorted(extra)} "
                            "outside the allowed list")
    return findings


def required_check(governance: str, settings: str) -> tuple[str, str, list[str], list[str]]:
    """The required check's name and workflow file, the protected branches, and findings."""
    findings = []
    match = re.search(r"^\| Required status check \| `([^`]+)`, the job of `([^`]+)`",
                      governance, re.M)
    name, path = (match.group(1), match.group(2)) if match else ("", "")
    if not match:
        findings.append(f"{GOVERNANCE}: no 'Required status check' row naming the job")
    contexts = re.findall(r"^\| Required status context \| `([^`]+)` \| `([^`]+)` \|", settings, re.M)
    names = {context for row in contexts for context in row}
    if not contexts:
        findings.append(f"{SETTINGS}: no 'Required status context' row")
    elif names != {name}:
        findings.append(f"{SETTINGS}: the rulesets require {sorted(names)}, "
                        f"{GOVERNANCE} names `{name}`")
    branches = re.search(r"^Applies to `([^`]+)` and every `([^`]+)` branch", governance, re.M)
    protected = [branches.group(1), branches.group(2)] if branches else []
    if not branches:
        findings.append(f"{GOVERNANCE}: no 'Applies to ... branch' line for the rulesets")
    return name, path, protected, findings


def integrity_findings(workflows: Mapping[str, Workflow], governance: str,
                       settings: str) -> list[str]:
    name, path, protected, findings = required_check(governance, settings)
    workflow = workflows.get(Path(path).name) if path else None
    if workflow is None:
        return findings + [f"required check workflow `{path}` not found"]
    file = Path(path).name
    matching = [(job_id, job) for job_id, job in jobs(workflow) if job.get("name") == name]
    if len(matching) != 1:
        findings.append(f"{file}: {len(matching)} jobs named `{name}`, the required check; "
                        "expected exactly 1")
    for job_id, job in matching:
        for key in ("if", "strategy"):
            if key in job:
                findings.append(f"{file}: required job {job_id} has `{key}`")
    on = triggers(workflow) or {}
    expected = {"branches": protected}
    for event in ("pull_request", "push"):
        if not isinstance(on, Mapping) or on.get(event) != expected:
            actual = on.get(event) if isinstance(on, Mapping) else on
            findings.append(f"{file}: `{event}` filters are {actual!r}; {GOVERNANCE} requires "
                            f"{expected!r} and no other filter")
    return findings


def labels_sync_findings(workflows: Mapping[str, Workflow]) -> list[str]:
    job = dict(jobs(workflows.get("labels-sync.yml", {}))).get("reconcile")
    if job is None:
        return ["labels-sync.yml: no reconcile job"]
    if DEFAULT_BRANCH_REF not in squash(job.get("if", "")):
        return ["labels-sync.yml: reconcile's condition no longer checks the default branch"]
    return []


def milestone_script(workflow: Workflow) -> str | None:
    """The inline Python of pr-milestones.yml's only step, or None."""
    reconcile = dict(jobs(workflow)).get("reconcile") or {}
    steps = reconcile.get("steps") or []
    if len(steps) != 1 or not isinstance(steps[0], Mapping):
        return None
    match = SCRIPT.search(str(steps[0].get("run", "")))
    return match.group(1) + "\n" if match else None


def milestone_findings(workflows: Mapping[str, Workflow]) -> list[str]:
    workflow = workflows.get(MILESTONES_FILE)
    if workflow is None:
        return [f"{MILESTONES_FILE}: not found"]
    findings = []
    if triggers(workflow) != MILESTONES_EVENTS:
        findings.append(f"{MILESTONES_FILE}: events changed: {triggers(workflow)!r}")
    if workflow.get("permissions") != {}:
        findings.append(f"{MILESTONES_FILE}: workflow-level permissions must be {{}}")
    all_jobs = dict(jobs(workflow))
    if list(all_jobs) != ["reconcile"]:
        findings.append(f"{MILESTONES_FILE}: jobs must be exactly [reconcile], not {list(all_jobs)}")
    job = all_jobs.get("reconcile", {})
    if job.get("permissions") != {"pull-requests": "write"}:
        findings.append(f"{MILESTONES_FILE}: reconcile's permissions must be pull-requests: write")
    if squash(job.get("if", "")) != MILESTONES_IF:
        findings.append(f"{MILESTONES_FILE}: reconcile's condition must be: {MILESTONES_IF}")
    steps = job.get("steps") or []
    step = steps[0] if len(steps) == 1 and isinstance(steps[0], Mapping) else {}
    if not step or "uses" in step or "run" not in step:
        findings.append(f"{MILESTONES_FILE}: reconcile must have a single inline `run` step")
    elif "${{" in str(step["run"]):
        findings.append(f"{MILESTONES_FILE}: the inline code contains `${{{{`")
    elif milestone_script(workflow) is None:
        findings.append(f"{MILESTONES_FILE}: no `<<'PY'` Python script in the step")
    return findings


def dependabot_findings(config: Any) -> list[str]:
    updates = config.get("updates") if isinstance(config, Mapping) else None
    if not isinstance(updates, list):
        return [f"{DEPENDABOT}: no updates list"]
    findings = []
    count = Counter(str(update.get("package-ecosystem")) for update in updates)
    for ecosystem, directory in PINNED_ECOSYSTEMS.items():
        if count[ecosystem] != 1:
            findings.append(f"{DEPENDABOT}: `{ecosystem}` is covered {count[ecosystem]} times; "
                            "expected once")
        for update in updates:
            if update.get("package-ecosystem") == ecosystem and update.get("directory") != directory:
                findings.append(f"{DEPENDABOT}: `{ecosystem}` must cover `{directory}`")
    for update in updates:
        ecosystem = update.get("package-ecosystem")
        if "milestone" in update:
            findings.append(f"{DEPENDABOT}: `{ecosystem}` sets a milestone")
        priorities = sorted(str(label) for label in update.get("labels") or []
                            if PRIORITY.match(str(label)))
        if priorities != ["P3"]:
            findings.append(f"{DEPENDABOT}: `{ecosystem}` has priorities {priorities}; "
                            "expected exactly ['P3']")
    return findings


def findings_for(workflows: Mapping[str, Workflow], dependabot: Any, governance: str,
                 settings: str) -> list[str]:
    findings = []
    for name, workflow in sorted(workflows.items()):
        findings += action_findings(name, workflow) + permission_findings(name, workflow)
    findings += integrity_findings(workflows, governance, settings)
    findings += labels_sync_findings(workflows) + milestone_findings(workflows)
    findings += dependabot_findings(dependabot)
    return findings


def load(root: Path, yaml: Any) -> tuple[dict[str, Workflow], Any, str, str]:
    loader = strict_loader(yaml)
    workflows = {path.name: yaml.load(path.read_text(encoding="utf-8"), Loader=loader)
                 for path in sorted((root / WORKFLOWS).glob("*.yml"))}
    dependabot = yaml.load((root / DEPENDABOT).read_text(encoding="utf-8"), Loader=loader)
    return (workflows, dependabot, (root / GOVERNANCE).read_text(encoding="utf-8"),
            (root / SETTINGS).read_text(encoding="utf-8"))


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args(argv)
    yaml = load_yaml()
    if yaml is None:
        if os.environ.get("GITHUB_ACTIONS"):
            print("workflow-invariants: FAIL: PyYAML is not installed "
                  "(tools/requirements-quality-ci.txt)")
            return 1
        print("workflow-invariants: SKIPPED locally: PyYAML is not installed")
        return 2
    try:
        workflows, dependabot, governance, settings = load(args.root, yaml)
    except yaml.YAMLError as error:
        print(f"workflow-invariants: FAIL: {error}")
        return 1
    findings = findings_for(workflows, dependabot, governance, settings)
    for finding in findings:
        print(finding)
    print(f"workflow-invariants: {'FAIL' if findings else 'PASS'} "
          f"({len(workflows)} workflows, {len(findings)} findings)")
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
