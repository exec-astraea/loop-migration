/**
 * Loop → Markdown Migration CLI
 *
 * Fetches workspace metadata from the Loop API, extracts the page hierarchy
 * from the Fluid snapshot, downloads each page as HTML, converts to Markdown,
 * and writes the result to disk.
 */

import "dotenv/config";
import { createInterface } from "node:readline/promises";
import { rm, stat } from "node:fs/promises";

import { getConfig } from "./lib/config.mts";
import { fetchLoopData } from "./lib/loop-api.mts";
import { fetchHierarchy } from "./lib/hierarchy.mts";
import { slugify } from "./lib/hierarchy.mts";
import { exportMarkdown } from "./lib/export.mts";
import type { Workspace } from "./lib/types.mts";
import type { ExportResult } from "./lib/types.mts";

// ---------------------------------------------------------------------------
// CLI helpers
// ---------------------------------------------------------------------------

const args = process.argv.slice(2);

function workspaceTitle(workspace: Workspace): string {
  return workspace.title || "(untitled)";
}

function workspaceDisplayLabel(workspace: Workspace): string {
  const title = workspaceTitle(workspace);
  return title.trim().toLowerCase() === "my workspace"
    ? "My workspace (Personal workspace)"
    : title;
}

function hasFlag(name: string): boolean {
  return args.includes(name);
}

function flagValue(name: string): string | undefined {
  const i = args.indexOf(name);
  if (i >= 0 && args[i + 1] && !args[i + 1].startsWith("-")) return args[i + 1];
  const eq = args.find((f) => f.startsWith(`${name}=`));
  return eq ? eq.slice(name.length + 1) : undefined;
}

// ---------------------------------------------------------------------------
// --help
// ---------------------------------------------------------------------------

if (hasFlag("-h") || hasFlag("--help")) {
  console.log(`
Usage: npm start -- [options]

Options:
  -w, --workspace NAME   Select workspace by name or ID
  -a, --all              Export all workspaces
  -p, --pick-workspace   Interactive workspace picker
  --page TITLE           Export a single page by title (substring match)
  --dump-html            Dump raw HTML alongside markdown (for debugging)
  -d, --delay MS         Delay between page requests in ms (default: 50)
  -n, --dry-run          Show what would be exported without fetching
  -h, --help             Show this help message

Examples:
  npm start                          # interactive picker  → workspaces/{slug}/
  npm start -- -w UFC                # export UFC workspace → workspaces/ufc/
  npm start -- --all                 # export all workspaces → workspaces/*/
  npm start -- -w UFC -n             # dry run
  npm start -- -w UFC -d 0           # no delay
  npm start -- --page "Client APIs"  # export single page
  npm start -- --page "Client APIs" --dump-html  # with raw HTML
`.trim());
  process.exit(0);
}

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------

const workspaceArg = flagValue("--workspace") || flagValue("-w");
const outputDir = "workspaces";
const rawDelay = flagValue("--delay") || flagValue("-d");
const delayMs = rawDelay !== undefined ? Number(rawDelay) : 50;
if (!Number.isFinite(delayMs) || delayMs < 0) {
  console.error(`Invalid --delay value: "${rawDelay}". Must be a non-negative number.`);
  process.exit(1);
}
const shouldPick = hasFlag("-p") || hasFlag("--pick-workspace");
const exportAll = hasFlag("-a") || hasFlag("--all");
const dryRun = hasFlag("-n") || hasFlag("--dry-run");
const pageFilter = flagValue("--page");
const dumpHtml = hasFlag("--dump-html");

// ---------------------------------------------------------------------------
// Workspace selection
// ---------------------------------------------------------------------------

/** Returns the selected workspace, or `null` to signal "export all". */
async function pickWorkspace(workspaces: Workspace[]): Promise<Workspace | null> {
  const rl = createInterface({ input: process.stdin, output: process.stdout });
  try {
    console.log("Select a workspace:");
    console.log(`  0. Export all workspaces`);
    workspaces.forEach((ws, i) =>
      console.log(`  ${i + 1}. ${workspaceDisplayLabel(ws)}`),
    );
    const answer = await rl.question("Enter number: ");
    const num = Number.parseInt(answer, 10);
    if (num === 0) return null;
    const idx = num - 1;
    if (Number.isNaN(idx) || idx < 0 || idx >= workspaces.length)
      throw new Error("Invalid selection");
    return workspaces[idx];
  } finally {
    rl.close();
  }
}

function findWorkspace(workspaces: Workspace[], nameOrId: string): Workspace {
  const exact = workspaces.find((ws) => ws.id === nameOrId);
  if (exact) return exact;
  const norm = nameOrId.trim().toLowerCase();
  const matches = workspaces.filter(
    (ws) => {
      const title = (ws.title || "").trim().toLowerCase();
      if (title === norm) return true;
      return norm === "personal workspace" && title === "my workspace";
    },
  );
  if (matches.length === 1) return matches[0];
  if (matches.length > 1)
    throw new Error(`Multiple workspaces match "${nameOrId}". Use -p to pick.`);
  throw new Error(`Workspace "${nameOrId}" not found. Use -p to pick.`);
}

// ---------------------------------------------------------------------------
// Export a single workspace
// ---------------------------------------------------------------------------

async function exportWorkspace(
  loopData: Awaited<ReturnType<typeof fetchLoopData>>,
  workspace: Workspace,
  dir: string,
  { forceClean = false }: { forceClean?: boolean } = {},
): Promise<ExportResult> {
  console.log(`\nWorkspace: ${workspaceDisplayLabel(workspace)}`);
  const flat = await fetchHierarchy(workspace);
  if (!dryRun && !pageFilter) {
    const dirExists = await stat(dir).then(() => true, () => false);
    if (forceClean || dirExists) await rm(dir, { recursive: true, force: true });
  }
  return exportMarkdown(loopData, workspace, flat, dir, { delayMs, dryRun, pageFilter, dumpHtml });
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  console.log("=== Loop → Markdown Migration ===\n");

  // Validate tokens early so the user gets a clear message before any API calls.
  getConfig();

  const loopData = await fetchLoopData();

  const allWorkspaces = loopData.workspaces || [];
  const workspaces = allWorkspaces.filter(
    (ws) => ws.mfs_info?.pod_id,
  );
  if (workspaces.length < allWorkspaces.length) {
    const skipped = allWorkspaces.filter((ws) => !ws.mfs_info?.pod_id);
    const names = skipped.map((ws) => ws.title || ws.id).join(", ");
    console.warn(
      `Note: ${skipped.length} workspace(s) hidden (no SharePoint pod_id): ${names}\n` +
      `  These may be empty or not yet synced to SharePoint.\n`,
    );
  }
  if (workspaces.length === 0) throw new Error("No workspaces with pod_id found");

  const results: ExportResult[] = [];

  if (exportAll) {
    // --all: export every workspace into its own subdirectory
    for (const ws of workspaces) {
      const wsDir = `${outputDir}/${slugify(ws.title || ws.id)}`;
      results.push(await exportWorkspace(loopData, ws, wsDir, { forceClean: true }));
    }
  } else {
    // Single workspace (or "all" chosen interactively)
    let workspace: Workspace | undefined;
    if (workspaceArg) {
      workspace = findWorkspace(workspaces, workspaceArg);
    } else if (shouldPick || workspaces.length > 1) {
      const picked = await pickWorkspace(workspaces);
      if (picked === null) {
        for (const ws of workspaces) {
          results.push(await exportWorkspace(loopData, ws, `${outputDir}/${slugify(ws.title || ws.id)}`, { forceClean: true }));
        }
      } else {
        workspace = picked;
      }
    } else {
      workspace = workspaces[0];
    }
    if (workspace) {
      results.push(await exportWorkspace(loopData, workspace, `${outputDir}/${slugify(workspace.title || workspace.id)}`));
    }
  }

  // Aggregate results
  const totalExported = results.reduce((s, r) => s + r.exported, 0);
  const totalSkipped = results.reduce((s, r) => s + r.skipped, 0);

  if (results.length > 1) {
    console.log(`\n=== Total: ${totalExported} exported, ${totalSkipped} skipped ===`);
  }

  // Exit 2 on partial failure so scripts can detect it
  if (totalSkipped > 0) process.exit(2);
}

main().catch((err: unknown) => {
  const msg = err instanceof Error ? err.message : String(err);
  console.error(`\nError: ${msg}`);
  process.exit(1);
});
