import type { LoopData, Workspace, LoopPage } from "./types.mts";
import { getConfig } from "./config.mts";

const LOOP_BASE =
  "https://substrate.office.com/recommended/api/v1.1/loop";
const MAX_PAGES = 50;

export function buildDedupKey(item: Record<string, unknown>, prefix: string) {
  for (const key of ["id", "workspace_id", "web_url", "url"]) {
    const v = item[key];
    if (typeof v === "string" && v.length > 0) return `${key}:${v}`;
  }
  return `${prefix}:${JSON.stringify(item)}`;
}

export function mergeArrays(
  a: Array<Record<string, unknown>> | undefined,
  b: Array<Record<string, unknown>> | undefined,
  prefix: string,
) {
  const merged = [...(a || [])];
  const seen = new Set(merged.map((i) => buildDedupKey(i, prefix)));
  for (const item of b || []) {
    const k = buildDedupKey(item, prefix);
    if (!seen.has(k)) {
      merged.push(item);
      seen.add(k);
    }
  }
  return merged;
}

export function mergeLoopData(base: LoopData, inc: LoopData): LoopData {
  return {
    ...base,
    ...inc,
    workspaces: mergeArrays(
      base.workspaces as unknown as Array<Record<string, unknown>>,
      inc.workspaces as unknown as Array<Record<string, unknown>>,
      "ws",
    ) as unknown as Workspace[],
    pages: mergeArrays(
      base.pages as unknown as Array<Record<string, unknown>>,
      inc.pages as unknown as Array<Record<string, unknown>>,
      "page",
    ) as unknown as LoopPage[],
    activities: mergeArrays(base.activities, inc.activities, "act"),
  };
}

async function loopGet(path: string, qs: string, loopToken: string): Promise<LoopData> {
  const url = `${LOOP_BASE}/${path}?${qs}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${loopToken}` },
  });
  if (res.status === 401 || res.status === 403) {
    throw new Error(
      `Loop API returned ${res.status} — your LOOP_BEARER_TOKEN has likely expired. ` +
      `Grab a fresh token and update your .env file.`,
    );
  }
  if (!res.ok) {
    const snippet = (await res.text()).slice(0, 300);
    throw new Error(`Loop API ${res.status} ${res.statusText} at ${url}${snippet ? ` :: ${snippet}` : ""}`);
  }
  return res.json() as Promise<LoopData>;
}

async function paginateAll(
  path: string,
  initialQs: string,
  loopToken: string,
): Promise<{ data: LoopData; requests: number }> {
  const visited = new Set<string>();
  let page = await loopGet(path, initialQs, loopToken);
  let merged = page;
  let requests = 1;

  while (page.next_page_link) {
    if (visited.has(page.next_page_link)) break;
    if (requests >= MAX_PAGES) throw new Error("Pagination limit reached");
    visited.add(page.next_page_link);
    page = await loopGet(path, page.next_page_link, loopToken);
    merged = mergeLoopData(merged, page);
    requests++;
  }

  return { data: merged, requests };
}

/**
 * Fetches all workspace + page metadata from the Loop API,
 * handling pagination automatically.
 *
 * Calls both the `/deltasync` and `/recent` endpoints because they return
 * different (overlapping) sets of workspaces.  Results are merged and
 * deduplicated by workspace/page id.
 */
export async function fetchLoopData(): Promise<LoopData> {
  console.log("Fetching workspace data from Loop API...");
  const { loopToken } = getConfig();

  // /recent surfaces workspaces ordered by activity — often includes ones
  // that deltasync misses (newly created, infrequently synced).
  const recent = await paginateAll(
    "recent",
    "top=30&settings=true&rs=en-us",
    loopToken,
  );
  const wsRecent = (recent.data.workspaces || []).length;

  // /deltasync returns the full component graph — may include workspaces
  // that haven't been touched recently.
  const delta = await paginateAll(
    "deltasync",
    "loopComponents=true&rs=en-us",
    loopToken,
  );
  const wsDelta = (delta.data.workspaces || []).length;

  const merged = mergeLoopData(recent.data, delta.data);
  const totalRequests = recent.requests + delta.requests;

  const ws = (merged.workspaces || []).length;
  const pg = (merged.pages || []).length;
  console.log(`  ${ws} workspaces, ${pg} pages (${totalRequests} API requests)`);
  if (wsRecent !== wsDelta) {
    console.log(`  (recent: ${wsRecent}, deltasync: ${wsDelta} — merged)`);
  }
  console.log();
  return merged;
}
