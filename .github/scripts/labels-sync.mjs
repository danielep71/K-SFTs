#!/usr/bin/env node

import assert from "node:assert/strict";
import { appendFile, readFile } from "node:fs/promises";
import process from "node:process";

const EXPECTED_SCHEMA_VERSION = 2;
const DEFAULT_MANIFEST = ".github/labels.json";
const API_VERSION = "2022-11-28";
const LABELS_PER_PAGE = 100;
const MAX_LABEL_PAGES = 100;
const MAX_RATE_LIMIT_RETRIES = 2;
const MAX_RETRY_DELAY_MS = 5000;
// The issues list filter takes a comma-separated list of label names.
const LABEL_FILTER_SEPARATOR = ",";

function compareNames(left, right) {
  const leftKey = left.toLowerCase();
  const rightKey = right.toLowerCase();
  if (leftKey < rightKey) return -1;
  if (leftKey > rightKey) return 1;
  return left < right ? -1 : left > right ? 1 : 0;
}

function sameKeys(value, expected) {
  const actual = Object.keys(value).sort();
  return JSON.stringify(actual) === JSON.stringify([...expected].sort());
}

function normalizedLiveLabel(label) {
  return {
    name: String(label.name ?? ""),
    color: String(label.color ?? "").toUpperCase(),
    description: label.description == null ? "" : String(label.description)
  };
}

function validateLabelArray(labels, location, seen) {
  const errors = [];
  if (!Array.isArray(labels)) return [`${location} must be an array`];

  labels.forEach((label, index) => {
    const item = `${location}[${index}]`;
    if (!label || typeof label !== "object" || Array.isArray(label)) {
      errors.push(`${item} must be an object`);
      return;
    }
    if (!sameKeys(label, ["name", "color", "description"])) {
      errors.push(`${item} must contain exactly name, color, and description`);
    }

    if (typeof label.name !== "string") {
      errors.push(`${item}.name must be a string`);
    } else {
      if (label.name.length === 0 || label.name !== label.name.trim()) {
        errors.push(`${item}.name must be non-empty with no surrounding whitespace`);
      }
      if (label.name.length > 50) errors.push(`${item}.name exceeds GitHub's 50-character limit`);
      if (/\r|\n/.test(label.name)) errors.push(`${item}.name must be single-line`);
      const key = label.name.toLowerCase();
      if (seen.has(key)) {
        errors.push(`${item}.name duplicates ${seen.get(key)} case-insensitively`);
      } else {
        seen.set(key, `${item}.name`);
      }
    }

    if (typeof label.color !== "string" || !/^[0-9A-F]{6}$/.test(label.color)) {
      errors.push(`${item}.color must be six uppercase hexadecimal characters without #`);
    }
    if (typeof label.description !== "string") {
      errors.push(`${item}.description must be a string`);
    } else {
      if (label.description.length === 0) errors.push(`${item}.description must not be empty`);
      if (label.description.length > 100) errors.push(`${item}.description exceeds GitHub's 100-character limit`);
      if (/\r|\n/.test(label.description)) errors.push(`${item}.description must be single-line`);
    }
  });

  for (let index = 1; index < labels.length; index += 1) {
    const previous = labels[index - 1]?.name;
    const current = labels[index]?.name;
    if (typeof previous === "string" && typeof current === "string" && compareNames(previous, current) >= 0) {
      errors.push(`${location} must be sorted case-insensitively by name: ${previous} precedes ${current}`);
      break;
    }
  }
  return errors;
}

export function validateManifest(document) {
  const errors = [];
  if (!document || typeof document !== "object" || Array.isArray(document)) {
    throw new Error("Manifest must be a JSON object.");
  }
  if (!sameKeys(document, ["schema_version", "prune", "labels"])) {
    errors.push("root must contain exactly schema_version, prune, and labels");
  }
  if (document.schema_version !== EXPECTED_SCHEMA_VERSION) {
    errors.push(`schema_version must be ${EXPECTED_SCHEMA_VERSION}`);
  }
  if (typeof document.prune !== "boolean") errors.push("prune must be a boolean");

  errors.push(...validateLabelArray(document.labels, "labels", new Map()));
  if (Array.isArray(document.labels) && document.labels.length === 0) errors.push("labels must not be empty");

  if (errors.length > 0) throw new Error(`Invalid label manifest:\n- ${errors.join("\n- ")}`);
  return structuredClone(document);
}

export function resolveDesired(document) {
  const labels = [...validateManifest(document).labels];
  labels.sort((left, right) => compareNames(left.name, right.name));
  return labels;
}

export function planChanges(desiredLabels, liveLabels, { prune = true } = {}) {
  const desired = new Map(desiredLabels.map(label => [label.name.toLowerCase(), label]));
  const live = new Map(liveLabels.map(label => {
    const normalized = normalizedLiveLabel(label);
    return [normalized.name.toLowerCase(), normalized];
  }));
  const changes = [];

  for (const [key, target] of desired) {
    const current = live.get(key);
    if (!current) {
      changes.push({ action: "create", name: target.name, target });
      continue;
    }
    const fields = [];
    if (current.name !== target.name) fields.push("name");
    if (current.color !== target.color) fields.push("color");
    if (current.description !== target.description) fields.push("description");
    if (fields.length > 0) {
      changes.push({ action: "update", name: target.name, currentName: current.name, fields, target });
    }
  }

  if (prune) {
    for (const [key, current] of live) {
      if (!desired.has(key)) changes.push({ action: "delete", name: current.name, currentName: current.name });
    }
  }

  const order = { update: 0, create: 1, delete: 2 };
  changes.sort((left, right) => order[left.action] - order[right.action] || compareNames(left.name, right.name));
  return changes;
}

function parseArguments(argv) {
  const parsed = {
    manifest: DEFAULT_MANIFEST,
    mode: null,
    live: null,
    repository: process.env.GITHUB_REPOSITORY || null,
    dryRun: false,
    selfTest: false
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === "--dry-run") {
      parsed.dryRun = true;
    } else if (argument === "--self-test") {
      parsed.selfTest = true;
    } else if (["--manifest", "--mode", "--live", "--repository"].includes(argument)) {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) throw new Error(`${argument} requires a value`);
      parsed[argument.slice(2)] = value;
      index += 1;
    } else {
      throw new Error(`Unknown argument: ${argument}`);
    }
  }
  return parsed;
}

async function loadJson(path) {
  let text;
  try {
    text = await readFile(path, "utf8");
  } catch (error) {
    throw new Error(`Cannot read ${path}: ${error.message}`);
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`Invalid JSON in ${path}: ${error.message}`);
  }
}

function splitRepository(repository) {
  const parts = String(repository ?? "").split("/");
  if (parts.length !== 2 || parts.some(part => part.length === 0)) {
    throw new Error("repository must be in owner/name form");
  }
  return parts.map(encodeURIComponent);
}

function apiUrl(repository, suffix) {
  const [owner, repo] = splitRepository(repository);
  return `https://api.github.com/repos/${owner}/${repo}${suffix}`;
}

function sleep(milliseconds) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

function isRateLimited(response) {
  return response.status === 429
    || (
      response.status === 403
      && (
        response.headers.get("x-ratelimit-remaining") === "0"
        || response.headers.has("retry-after")
      )
    );
}

function retryDelayMs(response, attempt) {
  const header = response.headers.get("retry-after");
  const retryAfter = header !== null && header.trim() !== "" ? Number(header) : NaN;
  if (Number.isFinite(retryAfter) && retryAfter >= 0) {
    return Math.min(retryAfter * 1000, MAX_RETRY_DELAY_MS);
  }
  return Math.min(1000 * (2 ** attempt), MAX_RETRY_DELAY_MS);
}

async function githubRequest(repository, token, method, suffix, body) {
  for (let attempt = 0; ; attempt += 1) {
    const response = await fetch(apiUrl(repository, suffix), {
      method,
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${token}`,
        "X-GitHub-Api-Version": API_VERSION,
        "User-Agent": "repository-labels-sync"
      },
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    if (response.ok) {
      if (response.status === 204) return null;
      return response.json();
    }

    const detail = await response.text();
    const retryable = method === "GET" && isRateLimited(response);
    if (!retryable || attempt >= MAX_RATE_LIMIT_RETRIES) {
      throw new Error(`${method} ${suffix} failed with HTTP ${response.status}: ${detail.slice(0, 500)}`);
    }
    await sleep(retryDelayMs(response, attempt));
  }
}

async function listLiveLabels(repository, token) {
  const labels = [];
  for (let page = 1; page <= MAX_LABEL_PAGES; page += 1) {
    const batch = await githubRequest(
      repository,
      token,
      "GET",
      `/labels?per_page=${LABELS_PER_PAGE}&page=${page}`
    );
    if (!Array.isArray(batch)) throw new Error("GitHub labels response is not an array");
    labels.push(...batch.map(normalizedLiveLabel));
    if (batch.length < LABELS_PER_PAGE) return labels;
  }
  throw new Error(`GitHub label pagination exceeded ${MAX_LABEL_PAGES} pages`);
}

// A name the issues list filter cannot match as exactly one label, or null.
function unqueryableReason(name) {
  if (name.length === 0) return "empty name";
  if (name.includes(LABEL_FILTER_SEPARATOR)) return "contains a comma, which the issues filter reads as a separator";
  if (name !== name.trim()) return "has surrounding whitespace";
  return null;
}

// Read-only: for every planned deletion, ask whether any issue or pull
// request, open or closed, carries the label. The issues list includes pull
// requests, so one result is enough to prove use; no pagination is needed.
async function findBlockedDeletions(repository, token, changes) {
  const inUse = [];
  const refused = [];
  for (const change of changes.filter(item => item.action === "delete")) {
    const name = change.currentName;
    const reason = unqueryableReason(name);
    if (reason) {
      refused.push({ name, reason });
      continue;
    }
    const batch = await githubRequest(
      repository,
      token,
      "GET",
      `/issues?labels=${encodeURIComponent(name)}&state=all&per_page=1`
    );
    if (!Array.isArray(batch)) throw new Error("GitHub issues response is not an array");
    if (batch.length > 0) {
      const item = batch[0];
      inUse.push({ name, example: `${item.pull_request ? "pull request" : "issue"} #${item.number}` });
    }
  }
  return { inUse, refused };
}

// Runs before the first create, update or delete: a label still in use, or
// one whose use cannot be checked, is never deleted, and nothing is changed.
async function assertDeletionsAllowed(repository, token, changes) {
  const { inUse, refused } = await findBlockedDeletions(repository, token, changes);
  if (inUse.length === 0 && refused.length === 0) return;
  const parts = [];
  if (inUse.length > 0) {
    parts.push(`still in use by an issue or pull request: ${inUse.map(item => `${JSON.stringify(item.name)} (${item.example})`).join(", ")}`);
  }
  if (refused.length > 0) {
    parts.push(`use cannot be queried reliably: ${refused.map(item => `${JSON.stringify(item.name)} (${item.reason})`).join(", ")}`);
  }
  throw new Error(
    `Refusing to delete labels outside the manifest; ${parts.join("; ")}. `
    + "No label was created, updated or deleted. Add each label to the manifest, "
    + "or remove it from every issue and pull request (or rename it) first."
  );
}

// The guard runs to completion before the state records that applying began.
async function guardAndApplyChanges(repository, token, changes, state) {
  await assertDeletionsAllowed(repository, token, changes);
  state.applying = true;
  await applyChanges(repository, token, changes);
}

async function applyChanges(repository, token, changes) {
  for (const change of changes) {
    if (change.action === "update") {
      await githubRequest(repository, token, "PATCH", `/labels/${encodeURIComponent(change.currentName)}`, {
        new_name: change.target.name,
        color: change.target.color,
        description: change.target.description
      });
    } else if (change.action === "create") {
      await githubRequest(repository, token, "POST", "/labels", change.target);
    } else if (change.action === "delete") {
      await githubRequest(repository, token, "DELETE", `/labels/${encodeURIComponent(change.currentName)}`);
    }
  }
}

// A Markdown code span for one dynamic value: the fence is longer than any
// backtick run inside it, a value at a backtick or both-ends space is padded,
// "|" is escaped for table cells and every line break becomes a space.
function markdownCode(value) {
  const text = String(value).replace(/\r\n|\r|\n/g, " ").replaceAll("|", "\\|");
  if (text.length === 0) return "` `";
  const longest = Math.max(0, ...Array.from(text.matchAll(/`+/g), match => match[0].length));
  const fence = "`".repeat(longest + 1);
  const pad = /^`|`$/.test(text) || (/^ .* $/s.test(text) && text.trim().length > 0) ? " " : "";
  return `${fence}${pad}${text}${pad}${fence}`;
}

function renderLabelNames(labels) {
  if (labels.length === 0) return "none";
  return labels.map(label => markdownCode(label.name)).join(", ");
}

function renderSummary({
  mode,
  manifest,
  manifestPath,
  desiredLabels,
  changes = [],
  verified = null,
  dryRun = false,
  error = null,
  partial = false
}) {
  const resolvedLabels = desiredLabels ?? [];
  const planned = changes !== null;
  const counts = Object.fromEntries(["create", "update", "delete"].map(action => [
    action,
    planned ? changes.filter(change => change.action === action).length : "n/a"
  ]));
  const lines = [
    "# Repository label synchronization",
    "",
    `- Mode: ${markdownCode(mode)}${dryRun ? " (dry run)" : ""}`,
    `- Label manifest: ${markdownCode(manifestPath ?? "unavailable")}`,
    `- Schema: **${manifest?.schema_version ?? "unavailable"}**`,
    `- Labels: **${resolvedLabels.length}**`,
    `- Prune: **${manifest ? (manifest.prune ? "enabled" : "disabled") : "unavailable"}**`,
    `- Planned changes: **${planned ? changes.length : "not computed (live labels not read)"}**`,
    `- Create / update / delete: **${counts.create} / ${counts.update} / ${counts.delete}**`
  ];
  if (verified !== null) lines.push(`- Post-run exact match: **${verified ? "yes" : "no"}**`);
  lines.push(
    "",
    "## Label catalogue",
    "",
    `- Labels (${resolvedLabels.length}): ${renderLabelNames(resolvedLabels)}`
  );
  if (error) {
    lines.push("", "## Failure", "", markdownCode(error));
    if (partial) {
      lines.push("", "The failure happened during or after reconciliation: some changes may already be applied. Compare the live labels with the manifest.");
    }
  } else if (!planned) {
    lines.push("", "Live labels were not read; no changes were planned.");
  } else if (changes.length === 0) {
    lines.push("", "No label changes are required.");
  } else {
    lines.push("", "| Action | Label | Detail |", "|---|---|---|");
    for (const change of changes) {
      const detail = change.action === "update" ? change.fields.join(", ") : change.action === "delete" ? "outside the manifest" : "missing live label";
      lines.push(`| ${change.action} | ${markdownCode(change.name)} | ${detail} |`);
    }
  }
  return `${lines.join("\n")}\n`;
}

async function publishSummary(summary) {
  process.stdout.write(summary);
  if (process.env.GITHUB_STEP_SUMMARY) await appendFile(process.env.GITHUB_STEP_SUMMARY, summary, "utf8");
}

function expectFailure(operation, pattern) {
  assert.throws(operation, pattern);
}

// Cells of one Markdown table row, split at unescaped pipes.
function tableCells(row) {
  return row.replace(/^\|/, "").replace(/\|$/, "").split(/(?<!\\)\|/).map(cell => cell.trim());
}

// Replaces fetch with a fake GitHub that answers the issues list filter (as
// GitHub does: exact label names, case-insensitive, open only unless state is
// given) and accepts every mutation; records every request it receives.
async function withFakeGitHub(issues, operation) {
  const original = globalThis.fetch;
  const requests = [];
  globalThis.fetch = async (url, init = {}) => {
    const parsed = new URL(url);
    const method = init.method ?? "GET";
    requests.push({ method, path: parsed.pathname, query: parsed.searchParams });
    if (method === "GET" && parsed.pathname.endsWith("/issues")) {
      const wanted = String(parsed.searchParams.get("labels") ?? "").split(",").map(name => name.toLowerCase());
      const state = parsed.searchParams.get("state") ?? "open";
      const perPage = Number(parsed.searchParams.get("per_page") ?? 30);
      const matching = issues.filter(issue => (
        (state === "all" || issue.state === state)
        && wanted.every(name => issue.labels.some(label => label.toLowerCase() === name))
      ));
      return new Response(JSON.stringify(matching.slice(0, perPage)), { status: 200 });
    }
    if (method === "DELETE") return new Response(null, { status: 204 });
    if (method === "GET") return new Response("unexpected request", { status: 404 });
    return new Response(init.body ?? "{}", { status: method === "POST" ? 201 : 200 });
  };
  try {
    let error = null;
    try {
      await operation();
    } catch (caught) {
      error = caught;
    }
    return { error, requests, mutations: requests.filter(request => request.method !== "GET") };
  } finally {
    globalThis.fetch = original;
  }
}

async function runSelfTest(manifestPath) {
  for (const [header, attempt, expected] of [
    [null, 0, 1000], [null, 2, 4000], ["", 1, 2000],
    ["invalid", 1, 2000], ["-1", 1, 2000], ["0", 2, 0],
    ["2", 0, 2000], ["999999", 0, MAX_RETRY_DELAY_MS],
    [null, 20, MAX_RETRY_DELAY_MS]
  ]) {
    const headers = new Headers(header === null ? {} : { "retry-after": header });
    assert.equal(retryDelayMs({ headers }, attempt), expected, `Retry-After: ${header}`);
  }

  const baseline = await loadJson(manifestPath);
  const manifest = validateManifest(baseline);
  const desired = resolveDesired(manifest);

  const missingRoot = structuredClone(baseline);
  delete missingRoot.prune;
  expectFailure(() => validateManifest(missingRoot), /root must contain exactly/);

  const extraRoot = structuredClone(baseline);
  extraRoot.overlays = {};
  expectFailure(() => validateManifest(extraRoot), /root must contain exactly/);

  const wrongSchema = structuredClone(baseline);
  wrongSchema.schema_version = 1;
  expectFailure(() => validateManifest(wrongSchema), /schema_version/);

  const invalidPrune = structuredClone(baseline);
  invalidPrune.prune = "true";
  expectFailure(() => validateManifest(invalidPrune), /prune must be a boolean/);

  const emptyLabels = structuredClone(baseline);
  emptyLabels.labels = [];
  expectFailure(() => validateManifest(emptyLabels), /must not be empty/);

  const missingField = structuredClone(baseline);
  delete missingField.labels[0].description;
  expectFailure(() => validateManifest(missingField), /name, color, and description/);

  const invalidColor = structuredClone(baseline);
  invalidColor.labels[0].color = "#c24e00";
  expectFailure(() => validateManifest(invalidColor), /uppercase hexadecimal/);

  const duplicateName = structuredClone(baseline);
  duplicateName.labels[1].name = baseline.labels[0].name.toUpperCase();
  expectFailure(() => validateManifest(duplicateName), /duplicates/);

  const longDescription = structuredClone(baseline);
  longDescription.labels[0].description = "x".repeat(101);
  expectFailure(() => validateManifest(longDescription), /100-character/);

  const unsorted = structuredClone(baseline);
  [unsorted.labels[0], unsorted.labels[1]] = [unsorted.labels[1], unsorted.labels[0]];
  expectFailure(() => validateManifest(unsorted), /sorted/);

  assert.equal(planChanges(desired, desired, { prune: manifest.prune }).length, 0, "idempotent plan must be empty");

  assert.deepEqual(planChanges(desired, desired.slice(1)).map(change => change.action), ["create"]);

  const changed = structuredClone(desired);
  changed[0].color = "000000";
  changed[0].description = "changed";
  const changedPlan = planChanges(desired, changed);
  assert.deepEqual(changedPlan.map(change => change.action), ["update"]);
  assert.deepEqual(changedPlan[0].fields, ["color", "description"]);

  const extra = [...structuredClone(desired), { name: "extra", color: "ABCDEF", description: "extra" }];
  assert.deepEqual(planChanges(desired, extra, { prune: true }).map(change => change.action), ["delete"]);
  assert.equal(planChanges(desired, extra, { prune: false }).length, 0);

  const caseDrift = structuredClone(desired);
  caseDrift[0].name = desired[0].name.toUpperCase();
  assert.deepEqual(planChanges(desired, caseDrift)[0].fields, ["name"]);

  const summary = renderSummary({
    mode: "plan",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    dryRun: true
  });
  assert.match(summary, /Label manifest: `fixture-labels\.json`/);
  assert.match(summary, new RegExp(`Labels: \\*\\*${desired.length}\\*\\*`));
  const validateSummary = renderSummary({
    mode: "validate",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    changes: null
  });
  assert.match(validateSummary, /Planned changes: \*\*not computed \(live labels not read\)\*\*/);
  assert.match(validateSummary, /Live labels were not read; no changes were planned\./);
  assert.doesNotMatch(validateSummary, /No label changes are required/);
  const failureSummary = renderSummary({
    mode: "failed",
    manifest: null,
    manifestPath: null,
    desiredLabels: [],
    changes: [{ action: "create", name: "new" }],
    partial: true,
    error: "write failed"
  });
  assert.match(failureSummary, /Prune: \*\*unavailable\*\*/);
  assert.match(failureSummary, /Planned changes: \*\*1\*\*/);
  assert.match(failureSummary, /some changes may already be applied/);
  assert.doesNotMatch(failureSummary, /Post-run exact match/);
  const unverifiedSummary = renderSummary({
    mode: "failed",
    manifest,
    manifestPath: "fixture-labels.json",
    desiredLabels: desired,
    changes: [{ action: "create", name: "new" }],
    verified: false,
    partial: true,
    error: "Post-run verification found 1 remaining difference(s)"
  });
  assert.match(unverifiedSummary, /Mode: `failed`/);
  assert.match(unverifiedSummary, /Post-run exact match: \*\*no\*\*/);
  assert.match(summary, new RegExp(`Labels \\(${desired.length}\\):`));
  for (const label of desired) assert.ok(summary.includes(markdownCode(label.name)), label.name);

  process.stdout.write("SELF-TEST PASS: retry delays, validation and reconciliation fixtures\n");

  // Deletion guard, against a fake GitHub. Each plan also updates and creates
  // a label, so a mutation before the guard would be recorded.
  const fakeIssues = [
    { number: 3, state: "open", labels: ["stale"] },
    { number: 9, state: "closed", pull_request: {}, labels: ["Old-PR"] },
    { number: 10, state: "closed", labels: ["bug"] }
  ];
  const mixedLive = structuredClone(desired).slice(1);
  mixedLive[0].color = mixedLive[0].color === "000000" ? "FFFFFF" : "000000";
  const guardPlan = extras => planChanges(desired, [...mixedLive, ...extras.map(name => (
    { name, color: "ABCDEF", description: "extra" }
  ))]);
  const guarded = async extras => {
    const state = { applying: false };
    const changes = guardPlan(extras);
    const outcome = await withFakeGitHub(fakeIssues, () => guardAndApplyChanges("fixture/repo", "token", changes, state));
    return { ...outcome, changes, state };
  };

  const openUse = await guarded(["stale"]);
  assert.ok(openUse.error, "a label on an open issue must not be deleted");
  assert.match(openUse.error.message, /still in use by an issue or pull request: "stale" \(issue #3\)/);
  assert.match(openUse.error.message, /No label was created, updated or deleted/);
  assert.deepEqual(openUse.mutations, [], "no mutation may precede a failed guard");
  assert.equal(openUse.state.applying, false, "a failed guard is not a partial run");
  const openQuery = openUse.requests.find(request => request.path.endsWith("/issues"));
  assert.equal(openQuery.query.get("labels"), "stale");
  assert.equal(openQuery.query.get("state"), "all");
  assert.equal(openQuery.query.get("per_page"), "1");
  process.stdout.write(`SELF-TEST GUARD (in use, open issue): ${openUse.error.message}\n`);

  const closedPull = await guarded(["old-pr"]);
  assert.ok(closedPull.error, "a label only on a closed pull request must not be deleted");
  assert.match(closedPull.error.message, /"old-pr" \(pull request #9\)/);
  assert.deepEqual(closedPull.mutations, []);

  const comma = await guarded(["a,b", "unused"]);
  assert.ok(comma.error, "a name the issues filter cannot express must be refused");
  assert.match(comma.error.message, /use cannot be queried reliably: "a,b" \(contains a comma/);
  assert.doesNotMatch(comma.error.message, /"unused"/);
  assert.deepEqual(comma.mutations, [], "a refused name blocks every mutation");
  assert.ok(!comma.requests.some(request => request.query.get("labels") === "a,b"));

  const both = await guarded(["stale", "old-pr", "x,y"]);
  assert.match(both.error.message, /"old-pr" \(pull request #9\), "stale" \(issue #3\); use cannot be queried reliably: "x,y"/);
  assert.deepEqual(both.mutations, []);

  const unused = await guarded(["unused: a/b é"]);
  assert.equal(unused.error, null, "an unused extra label is deleted");
  assert.equal(unused.state.applying, true);
  assert.deepEqual(unused.mutations.map(request => request.method), ["PATCH", "POST", "DELETE"]);
  assert.equal(decodeURIComponent(unused.mutations[2].path.split("/labels/")[1]), "unused: a/b é");
  const unusedQuery = unused.requests.findIndex(request => request.query.get("labels") === "unused: a/b é");
  assert.ok(unusedQuery >= 0 && unusedQuery < unused.requests.indexOf(unused.mutations[0]), "the query precedes every mutation");

  const noDeletion = await withFakeGitHub(fakeIssues, () => guardAndApplyChanges("fixture/repo", "token", guardPlan([]), { applying: false }));
  assert.equal(noDeletion.error, null);
  assert.equal(noDeletion.requests.filter(request => request.method === "GET").length, 0, "no deletion, no query");
  process.stdout.write("SELF-TEST PASS: deletion guard refuses labels in use (open issue, closed pull request) and unqueryable names before any mutation; an unused extra label is deleted\n");

  // Summary values: one code span per value, longer than its backtick runs.
  assert.equal(markdownCode("plain"), "`plain`");
  assert.equal(markdownCode("a`b"), "``a`b``");
  assert.equal(markdownCode("``x`"), "``` ``x` ```");
  assert.equal(markdownCode("a|b"), "`a\\|b`");
  assert.equal(markdownCode("one\r\ntwo\rthree\nfour"), "`one two three four`");
  assert.equal(markdownCode(" padded "), "`  padded  `");
  assert.equal(markdownCode(""), "` `");
  const hostile = "odd `x` | y\r\nz``";
  const hostileSummary = renderSummary({
    mode: "plan",
    manifest,
    manifestPath: "dir|`x`.json",
    desiredLabels: desired,
    changes: [{ action: "delete", name: hostile, currentName: hostile }]
  });
  const hostileRow = hostileSummary.split("\n").find(line => line.startsWith("| delete |"));
  assert.deepEqual(tableCells(hostileRow), ["delete", "``` odd `x` \\| y z`` ```", "outside the manifest"]);
  assert.match(hostileSummary, /Label manifest: ``dir\\\|`x`\.json``/);
  assert.doesNotMatch(hostileSummary, /\r/);
  const hostileFailure = renderSummary({ mode: "failed", manifest, manifestPath: "m", desiredLabels: desired, error: "HTTP 422: `bad` | line\nbreak" });
  assert.match(hostileFailure, /\n``HTTP 422: `bad` \\\| line break``\n/);
  process.stdout.write("SELF-TEST PASS: summary code spans for values with backticks, pipes and line breaks\n");
}

// What the run knew before a failure, so that the failure summary reports it.
const runState = { summaryBase: null, changes: null, applying: false, verified: null };

async function main() {
  const options = parseArguments(process.argv.slice(2));
  if (options.selfTest) {
    await runSelfTest(options.manifest);
    return;
  }

  const manifest = validateManifest(await loadJson(options.manifest));
  const desired = resolveDesired(manifest);
  const summaryBase = {
    manifest,
    manifestPath: options.manifest,
    desiredLabels: desired
  };
  runState.summaryBase = summaryBase;

  if (options.mode === "validate") {
    await publishSummary(renderSummary({ mode: "validate", ...summaryBase, changes: null }));
    return;
  }

  let live;
  if (options.mode === "plan") {
    if (!options.live) throw new Error("--mode plan requires --live");
    const liveDocument = await loadJson(options.live);
    live = Array.isArray(liveDocument) ? liveDocument : liveDocument.labels;
    if (!Array.isArray(live)) throw new Error("live fixture must be an array or contain a labels array");
  } else if (options.mode === "reconcile") {
    if (!options.repository) throw new Error("--mode reconcile requires --repository or GITHUB_REPOSITORY");
    if (!process.env.GITHUB_TOKEN) throw new Error("--mode reconcile requires GITHUB_TOKEN");
    live = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  } else {
    throw new Error("--mode must be validate, plan, or reconcile");
  }

  const changes = planChanges(desired, live, { prune: manifest.prune });
  runState.changes = changes;
  if (options.mode === "reconcile" && options.dryRun) {
    await assertDeletionsAllowed(options.repository, process.env.GITHUB_TOKEN, changes);
  }
  if (options.mode === "plan" || options.dryRun) {
    await publishSummary(renderSummary({ mode: options.mode, ...summaryBase, changes, dryRun: options.dryRun }));
    return;
  }

  await guardAndApplyChanges(options.repository, process.env.GITHUB_TOKEN, changes, runState);
  const after = await listLiveLabels(options.repository, process.env.GITHUB_TOKEN);
  const remaining = planChanges(desired, after, { prune: manifest.prune });
  const verified = remaining.length === 0;
  if (!verified) {
    runState.verified = false;
    throw new Error(`Post-run verification found ${remaining.length} remaining difference(s)`);
  }
  await publishSummary(renderSummary({ mode: "reconcile", ...summaryBase, changes, verified }));
}

main().catch(async error => {
  const message = error instanceof Error ? error.message : String(error);
  try {
    await publishSummary(renderSummary({
      mode: "failed",
      ...(runState.summaryBase ?? {
        manifest: null,
        manifestPath: null,
        desiredLabels: []
      }),
      changes: runState.changes,
      verified: runState.verified,
      partial: runState.applying,
      error: message
    }));
  } finally {
    process.stderr.write(`${message}\n`);
    process.exitCode = 1;
  }
});
