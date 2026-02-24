import type { LoopData, Workspace, LoopPage } from "./types.mts";
import { getConfig } from "./config.mts";

const DELTA_SYNC =
  "https://substrate.office.com/recommended/api/beta/loop/deltasync";
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

/**
 * Fetches all workspace + page metadata from the Loop delta-sync API,
 * handling pagination automatically.
 */
export async function fetchLoopData(): Promise<LoopData> {
  console.log("Fetching workspace data from Loop API...");
  const { loopToken } = getConfig();

  async function fetchPage(qs: string): Promise<LoopData> {
    const res = await fetch(`${DELTA_SYNC}?${qs}`, {
      headers: { Authorization: `Bearer ${loopToken}` },
    });
    if (res.status === 401 || res.status === 403) {
      throw new Error(
        `Loop API returned ${res.status} — your LOOP_BEARER_TOKEN has likely expired. ` +
        `Grab a fresh token and update your .env file.`,
      );
    }
    if (!res.ok) throw new Error(`Loop API ${res.status} ${res.statusText}`);
    return res.json() as Promise<LoopData>;
  }

  const visited = new Set<string>();
  let page = await fetchPage("loopComponents=true");
  let merged = page;
  let requests = 0;

  while (page.next_page_link && page.is_complete !== true) {
    if (visited.has(page.next_page_link)) throw new Error("Pagination loop");
    if (++requests >= MAX_PAGES) throw new Error("Pagination limit reached");
    visited.add(page.next_page_link);
    page = await fetchPage(page.next_page_link);
    merged = mergeLoopData(merged, page);
  }

  const ws = (merged.workspaces || []).length;
  const pg = (merged.pages || []).length;
  console.log(`  ${ws} workspaces, ${pg} pages\n`);
  return merged;
}
