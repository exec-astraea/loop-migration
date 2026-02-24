import { writeFile, mkdir } from "node:fs/promises";
import TurndownService from "turndown";
// @ts-expect-error — no type declarations for this package
import { tables } from "turndown-plugin-gfm";

import { spGet } from "./sharepoint.mts";
import { slugify } from "./hierarchy.mts";
import type { LoopData, LoopPage, Workspace, FlatEntry, ExportResult } from "./types.mts";

// ---------------------------------------------------------------------------
// Page HTML fetching
// ---------------------------------------------------------------------------

export function extractSpHost(page: LoopPage): string | undefined {
  const url = page.sharepoint_info?.site_url;
  if (!url) return undefined;
  try {
    return new URL(url).origin;
  } catch {
    console.warn(`  [WARN] Malformed site_url for "${page.title?.trim()}": ${url}`);
    return undefined;
  }
}

export function extractItemId(pageId: string) {
  return pageId.split("_").pop();
}

async function fetchPageHtml(page: LoopPage): Promise<string | null> {
  const driveId = page.onedrive_info?.drive_id;
  const itemId = extractItemId(page.id);
  const spHost = extractSpHost(page);
  if (!driveId || !itemId || !spHost) return null;

  const url = `${spHost}/_api/v2.0/drives/${driveId}/items/${itemId}/content?format=html&ump=1`;
  const res = await spGet(url);
  if (!res.ok) {
    const t = await res.text();
    console.error(`  [SKIP] HTTP ${res.status} for "${page.title?.trim()}" — ${t.slice(0, 120)}`);
    return null;
  }
  return res.text();
}

// ---------------------------------------------------------------------------
// Turndown (HTML → Markdown)
// ---------------------------------------------------------------------------

function createTurndown() {
  const td = new TurndownService({
    headingStyle: "atx",
    codeBlockStyle: "fenced",
    bulletListMarker: "-",
  });

  // GFM tables
  td.use(tables as TurndownService.Plugin);

  // ---------- Table-cell helpers ----------
  // Loop wraps every cell's text in deeply nested divs (fluid-data-type,
  // lang/dir, data-docparser-context). Turndown treats each div as a block
  // element, adding newlines that break GFM table rows. The rules below
  // flatten block elements and <hr>s inside <td>/<th> to inline content.

  function isInTableCell(node: HTMLElement): boolean {
    let p = node.parentElement;
    while (p) {
      if (p.nodeName === "TD" || p.nodeName === "TH") return true;
      if (p.nodeName === "TABLE") return false;
      p = p.parentElement;
    }
    return false;
  }

  const CELL_BLOCKS = new Set(["DIV", "ARTICLE", "ADDRESS", "SECTION", "HEADER", "FOOTER"]);

  td.addRule("tableCellBlocks", {
    filter(node: HTMLElement) {
      return CELL_BLOCKS.has(node.nodeName) && isInTableCell(node);
    },
    replacement(content: string) {
      // Collapse block-level newlines to a single space so cells stay on one line.
      return " " + content.replace(/\n{1,}/g, " ").trim();
    },
  });

  td.addRule("tableCellHr", {
    filter(node: HTMLElement) {
      return node.nodeName === "HR" && isInTableCell(node);
    },
    replacement() {
      return " — ";
    },
  });

  // Strip mailto: links (Loop @mentions)
  td.addRule("mentions", {
    filter(node: HTMLElement) {
      return node.nodeName === "A" && (node.getAttribute("href") || "").startsWith("mailto:");
    },
    replacement(_content: string, node: HTMLElement) {
      return node.textContent || "";
    },
  });

  // Preserve code block language annotations
  td.addRule("fencedCodeBlock", {
    filter(node: HTMLElement) {
      return node.nodeName === "PRE" && !!node.querySelector("code");
    },
    replacement(_content: string, node: HTMLElement) {
      const code = node.querySelector("code")!;
      const lang = (code.getAttribute("class") || "").replace(/^language-/, "").trim();
      const text = code.textContent || "";
      return `\n\n\`\`\`${lang}\n${text.replace(/\n$/, "")}\n\`\`\`\n\n`;
    },
  });

  // Task list checkboxes
  td.addRule("taskListItem", {
    filter(node: HTMLElement) {
      return node.nodeName === "LI" && !!node.querySelector("input[type=checkbox]");
    },
    replacement(content: string, node: HTMLElement) {
      const checkbox = node.querySelector("input[type=checkbox]");
      const checked = checkbox?.hasAttribute("checked") ? "x" : " ";
      const text = content.replace(/^\s*\[[ x]\]\s*/i, "").trim();
      return `- [${checked}] ${text}\n`;
    },
  });

  // Strikethrough
  td.addRule("strikethrough", {
    filter: ["del", "s"],
    replacement(content: string) {
      return `~~${content}~~`;
    },
  });

  return td;
}

// ---------------------------------------------------------------------------
// Path mapping
// ---------------------------------------------------------------------------

export function fileSlug(input: string) {
  return slugify(input).slice(0, 80);
}

export function buildPathMap(entries: FlatEntry[]) {
  const map = new Map<string, { dir: string; isSection: boolean }>();
  for (const entry of entries) {
    if (!entry.spoItemId) continue;
    if (entry.hasChildren) {
      map.set(entry.spoItemId, { dir: entry.path, isSection: true });
    } else {
      const lastSlash = entry.path.lastIndexOf("/");
      map.set(entry.spoItemId, {
        dir: lastSlash >= 0 ? entry.path.slice(0, lastSlash) : "",
        isSection: false,
      });
    }
  }
  return map;
}

// ---------------------------------------------------------------------------
// Filename deduplication
// ---------------------------------------------------------------------------

/**
 * Returns a unique filename within the given directory tracker.
 * If "my-page.md" is taken, tries "my-page-2.md", "my-page-3.md", etc.
 */
export function dedupeFilename(dir: string, filename: string, usedNames: Map<string, Set<string>>): string {
  if (!usedNames.has(dir)) usedNames.set(dir, new Set());
  const dirSet = usedNames.get(dir)!;

  if (!dirSet.has(filename)) {
    dirSet.add(filename);
    return filename;
  }

  const ext = filename.endsWith(".md") ? ".md" : "";
  const base = ext ? filename.slice(0, -ext.length) : filename;
  let counter = 2;
  while (dirSet.has(`${base}-${counter}${ext}`)) counter++;
  const deduped = `${base}-${counter}${ext}`;
  dirSet.add(deduped);
  return deduped;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export interface ExportOptions {
  delayMs?: number;
  dryRun?: boolean;
  pageFilter?: string;
  dumpHtml?: boolean;
}

/**
 * Downloads each page as HTML, converts to Markdown, and writes to `outputDir/`.
 * Uses the hierarchy's flat entries to determine directory structure.
 * Returns an ExportResult with counts and skipped page titles.
 */
export async function exportMarkdown(
  loopData: LoopData,
  workspace: Workspace,
  flat: FlatEntry[],
  outputDir: string,
  opts: ExportOptions = {},
): Promise<ExportResult> {
  const { delayMs = 50, dryRun = false, pageFilter, dumpHtml = false } = opts;

  let pages = (loopData.pages || []).filter(
    (p) =>
      p.type === "Fluid" &&
      p.onedrive_info?.drive_id &&
      !p.is_deleted &&
      p.workspace_id === workspace.id,
  );

  if (pageFilter) {
    const norm = pageFilter.toLowerCase();
    pages = pages.filter((p) =>
      (p.title || "").toLowerCase().includes(norm),
    );
    if (pages.length === 0) {
      console.error(`No pages matching "${pageFilter}" in workspace "${workspace.title}".`);
      return { exported: 0, skipped: 0, skippedPages: [] };
    }
  }

  const pathMap = buildPathMap(flat);
  const td = dryRun ? null : createTurndown();
  const usedNames = new Map<string, Set<string>>();

  const mode = dryRun ? "Dry run" : "Exporting";
  console.log(`${mode}: ${pages.length} pages → ${outputDir}/\n`);
  if (!dryRun) await mkdir(outputDir, { recursive: true });

  let exported = 0;
  let skipped = 0;
  const skippedPages: string[] = [];

  for (let i = 0; i < pages.length; i++) {
    const page = pages[i];
    const title = (page.title || "untitled").trim();
    const spoItemId = extractItemId(page.id) || "";

    const treeEntry = pathMap.get(spoItemId);
    let dir: string;
    let filename: string;

    if (treeEntry) {
      dir = treeEntry.dir ? `${outputDir}/${treeEntry.dir}` : outputDir;
      filename = treeEntry.isSection ? "_index.md" : `${fileSlug(title)}.md`;
    } else {
      dir = outputDir;
      filename = `${fileSlug(title || "untitled")}.md`;
    }

    filename = dedupeFilename(dir, filename, usedNames);

    const relPath =
      dir === outputDir ? filename : `${dir.slice(outputDir.length + 1)}/${filename}`;

    if (dryRun) {
      console.log(`  ${relPath}`);
      exported++;
      continue;
    }

    await mkdir(dir, { recursive: true });
    process.stdout.write(`[${i + 1}/${pages.length}] ${title}...`);

    const html = await fetchPageHtml(page);
    if (!html) {
      skipped++;
      skippedPages.push(title);
      console.log(" skipped");
      continue;
    }

    const md = `# ${title}\n\n${td!.turndown(html)}`;
    await writeFile(`${dir}/${filename}`, md, "utf8");
    if (dumpHtml) {
      const htmlFile = filename.replace(/\.md$/, ".html");
      await writeFile(`${dir}/${htmlFile}`, html, "utf8");
    }
    exported++;
    console.log(` → ${relPath} (${md.length} chars)`);

    if (delayMs > 0 && i < pages.length - 1)
      await new Promise((r) => setTimeout(r, delayMs));
  }

  console.log(`\nDone: ${exported} exported, ${skipped} skipped`);

  if (skippedPages.length > 0) {
    console.log(`\nSkipped pages:`);
    for (const name of skippedPages) console.log(`  - ${name}`);
  }

  return { exported, skipped, skippedPages };
}
