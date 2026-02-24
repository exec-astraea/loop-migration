import { spGet } from "./sharepoint.mts";
import type { HierarchyNode, FlatEntry, Workspace } from "./types.mts";

// ---------------------------------------------------------------------------
// Snapshot fetching
// ---------------------------------------------------------------------------

export function decodePodId(podId: string) {
  const decoded = Buffer.from(podId, "base64").toString("utf8");
  const parts = decoded.split("|");
  if (parts.length < 4) throw new Error(`Invalid pod_id: expected >=4 parts, got ${parts.length}`);
  return { host: parts[1], driveId: parts[2], itemId: parts[3] };
}

async function fetchSnapshot(host: string, driveId: string, itemId: string) {
  const url = `https://${host}/_api/v2.1/drives/${driveId}/items/${itemId}/opStream/snapshots/trees/latest?ump=1`;
  console.log("Fetching Fluid snapshot...");
  const res = await spGet(url);
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`Snapshot fetch failed: ${res.status} — ${t.slice(0, 300)}`);
  }
  return res.json();
}

// ---------------------------------------------------------------------------
// SharedTree parsing
// ---------------------------------------------------------------------------

export function findSharedTree(snapshot: any): any {
  for (const blob of snapshot.blobs || []) {
    if (blob.size < 1000) continue;
    try {
      const d = JSON.parse(Buffer.from(blob.content, "base64").toString("utf8"));
      if (d.editHistory && d.internedStrings) return d;
    } catch { /* not a SharedTree blob */ }
  }
  throw new Error("No SharedTree blob found in snapshot");
}

export function findBuild0(sharedTree: any): any[] {
  for (const chunk of sharedTree.editHistory.editChunks)
    for (const edit of chunk.chunk)
      for (const change of edit.changes || [])
        if (change.type === 5 && change.source?.length > 1)
          return change.source;
  throw new Error("No multi-node Build change found");
}

// ---------------------------------------------------------------------------
// Page metadata extraction
// ---------------------------------------------------------------------------

export type PageMeta = Map<string, { title: string; emoji: string | null; itemId: string | null }>;

export function extractPageMeta(source: any[], IS: string[]): PageMeta {
  const meta: PageMeta = new Map();
  for (const node of source) {
    const data = node[2];
    if (!Array.isArray(data)) continue;
    const def = data[0];
    if (typeof def !== "object" || def?.label !== "LoopPage") continue;

    const entry = { title: "???", emoji: null as string | null, itemId: null as string | null };
    for (let k = 1; k < data.length; k += 2) {
      const tName = IS[data[k]];
      const inner = data[k + 1]?.[0];
      if (!inner) continue;
      const payload = Array.isArray(inner[2]) ? inner[2] : inner[1];

      if (tName === "displayText" && Array.isArray(payload) && typeof payload[0] === "string")
        entry.title = payload[0];
      if (tName === "icon") {
        const o = payload?.[0];
        if (o?.type === "emoji" && o.data) entry.emoji = o.data;
      }
      if (tName === "odspMetadata") {
        const o = payload?.[0];
        if (o?.itemId) entry.itemId = o.itemId;
      }
    }
    meta.set(def.id, entry);
  }
  return meta;
}

// ---------------------------------------------------------------------------
// Hierarchy tree extraction
// ---------------------------------------------------------------------------

export function extractHierarchy(source: any[], IS: string[], pageMeta: PageMeta): HierarchyNode[] {
  const valuesIdx = IS.indexOf("values");
  const wsData = source[0][2];
  if (wsData[0]?.label !== "LoopWorkspace")
    throw new Error(`Expected LoopWorkspace, got: ${wsData[0]?.label}`);

  let valuesTrait: any[] | null = null;
  for (let i = 1; i < wsData.length; i += 2) {
    if (wsData[i] === valuesIdx) { valuesTrait = wsData[i + 1]; break; }
  }
  if (!valuesTrait) throw new Error("No 'values' trait on workspace node");

  function parseNode(node: any): HierarchyNode {
    const data = node[2];
    const pageId: string = data[0];
    const meta = pageMeta.get(pageId);
    return {
      pageId,
      title: meta?.title || "???",
      emoji: meta?.emoji || null,
      spoItemId: meta?.itemId || null,
      children:
        data.length >= 3 && data[1] === valuesIdx && Array.isArray(data[2])
          ? data[2].map(parseNode)
          : [],
    };
  }

  return valuesTrait.map(parseNode);
}

// ---------------------------------------------------------------------------
// Helpers: slugify, flatten, count, print
// ---------------------------------------------------------------------------

export function slugify(text: string) {
  const slug = text.trim().toLowerCase()
    .replace(/[^\p{L}\p{N}\p{Emoji_Presentation}\w\s-]/gu, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");
  return slug || "untitled";
}

export function flattenWithPaths(nodes: HierarchyNode[], parentPath = ""): FlatEntry[] {
  const result: FlatEntry[] = [];
  const usedSlugs = new Set<string>();
  for (const node of nodes) {
    let slug = slugify(node.title);
    if (usedSlugs.has(slug)) {
      let counter = 2;
      while (usedSlugs.has(`${slug}-${counter}`)) counter++;
      slug = `${slug}-${counter}`;
    }
    usedSlugs.add(slug);
    const path = parentPath ? `${parentPath}/${slug}` : slug;
    result.push({
      pageId: node.pageId, title: node.title, emoji: node.emoji,
      spoItemId: node.spoItemId, path, hasChildren: node.children.length > 0,
    });
    result.push(...flattenWithPaths(node.children, path));
  }
  return result;
}

function countPages(nodes: HierarchyNode[]): number {
  return nodes.reduce((s, n) => s + 1 + countPages(n.children), 0);
}

export function printTree(nodes: HierarchyNode[], indent = "") {
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    const last = i === nodes.length - 1;
    const icon = n.emoji ? `${n.emoji} ` : "";
    const info = n.children.length > 0 ? ` (${n.children.length} children)` : "";
    console.log(`${indent}${last ? "└── " : "├── "}${icon}${n.title}${info}`);
    if (n.children.length > 0)
      printTree(n.children, indent + (last ? "    " : "│   "));
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Fetches the Fluid snapshot for a workspace and extracts the page hierarchy.
 * Returns a flat list of entries with filesystem paths.
 */
export async function fetchHierarchy(workspace: Workspace): Promise<FlatEntry[]> {
  const podId = workspace.mfs_info?.pod_id;
  if (!podId) throw new Error(`Workspace "${workspace.title}" has no pod_id`);

  const pod = decodePodId(podId);
  console.log(`Host: ${pod.host}  Drive: ${pod.driveId}\n`);

  const snapshot = await fetchSnapshot(pod.host, pod.driveId, pod.itemId);
  console.log(`  ${snapshot.blobs?.length} blobs, seq ${snapshot.latestSequenceNumber}`);

  const sharedTree = findSharedTree(snapshot);
  const build0 = findBuild0(sharedTree);
  const IS: string[] = sharedTree.internedStrings;

  const pageMeta = extractPageMeta(build0, IS);
  console.log(`  ${pageMeta.size} page metadata nodes`);

  const hierarchy = extractHierarchy(build0, IS, pageMeta);
  const total = countPages(hierarchy);
  const flat = flattenWithPaths(hierarchy);
  console.log(`  ${total} pages in hierarchy\n`);

  printTree(hierarchy);
  console.log();

  return flat;
}
