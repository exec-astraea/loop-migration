import { describe, it, expect, vi } from "vitest";
import {
  extractItemId,
  extractSpHost,
  fileSlug,
  buildPathMap,
  dedupeFilename,
} from "../lib/export.mts";
import type { LoopPage, FlatEntry } from "../lib/types.mts";

// ---------------------------------------------------------------------------
// extractItemId
// ---------------------------------------------------------------------------

describe("extractItemId", () => {
  it("extracts the segment after the last underscore", () => {
    expect(extractItemId("ws_drive_item123")).toBe("item123");
  });

  it("returns the full string when there is no underscore", () => {
    expect(extractItemId("nounderscores")).toBe("nounderscores");
  });

  it("handles multiple underscores", () => {
    expect(extractItemId("a_b_c_d")).toBe("d");
  });

  it("returns empty string for trailing underscore", () => {
    expect(extractItemId("trailing_")).toBe("");
  });
});

// ---------------------------------------------------------------------------
// extractSpHost
// ---------------------------------------------------------------------------

describe("extractSpHost", () => {
  it("extracts origin from a valid site_url", () => {
    const page = { id: "x", sharepoint_info: { site_url: "https://contoso.sharepoint.com/sites/team" } } as LoopPage;
    expect(extractSpHost(page)).toBe("https://contoso.sharepoint.com");
  });

  it("returns undefined when site_url is missing", () => {
    const page = { id: "x" } as LoopPage;
    expect(extractSpHost(page)).toBeUndefined();
  });

  it("returns undefined and warns for malformed URL", () => {
    const spy = vi.spyOn(console, "warn").mockImplementation(() => {});
    const page = { id: "x", title: "Bad", sharepoint_info: { site_url: "not-a-url" } } as LoopPage;
    expect(extractSpHost(page)).toBeUndefined();
    expect(spy).toHaveBeenCalledOnce();
    spy.mockRestore();
  });
});

// ---------------------------------------------------------------------------
// fileSlug
// ---------------------------------------------------------------------------

describe("fileSlug", () => {
  it("slugifies and truncates to 80 chars", () => {
    const long = "A".repeat(100);
    const result = fileSlug(long);
    expect(result.length).toBeLessThanOrEqual(80);
    expect(result).toBe("a".repeat(80));
  });

  it("produces the same result as slugify for short input", () => {
    expect(fileSlug("Hello World")).toBe("hello-world");
  });
});

// ---------------------------------------------------------------------------
// buildPathMap
// ---------------------------------------------------------------------------

describe("buildPathMap", () => {
  const entry = (spoItemId: string | null, path: string, hasChildren: boolean): FlatEntry => ({
    pageId: "p", title: "T", emoji: null, spoItemId, path, hasChildren,
  });

  it("maps section pages (hasChildren) with dir = full path, isSection = true", () => {
    const map = buildPathMap([entry("item1", "design/components", true)]);
    expect(map.get("item1")).toEqual({ dir: "design/components", isSection: true });
  });

  it("maps leaf pages with dir = parent path, isSection = false", () => {
    const map = buildPathMap([entry("item2", "design/components/button", false)]);
    expect(map.get("item2")).toEqual({ dir: "design/components", isSection: false });
  });

  it("uses empty dir for top-level leaf pages", () => {
    const map = buildPathMap([entry("item3", "readme", false)]);
    expect(map.get("item3")).toEqual({ dir: "", isSection: false });
  });

  it("skips entries without spoItemId", () => {
    const map = buildPathMap([entry(null, "orphan", false)]);
    expect(map.size).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// dedupeFilename
// ---------------------------------------------------------------------------

describe("dedupeFilename", () => {
  it("returns the original filename when no collision", () => {
    const used = new Map<string, Set<string>>();
    expect(dedupeFilename("/out", "page.md", used)).toBe("page.md");
  });

  it("appends -2 on first collision", () => {
    const used = new Map<string, Set<string>>();
    dedupeFilename("/out", "page.md", used);
    expect(dedupeFilename("/out", "page.md", used)).toBe("page-2.md");
  });

  it("increments counter for repeated collisions", () => {
    const used = new Map<string, Set<string>>();
    dedupeFilename("/out", "page.md", used);
    dedupeFilename("/out", "page.md", used); // page-2.md
    expect(dedupeFilename("/out", "page.md", used)).toBe("page-3.md");
  });

  it("tracks directories independently", () => {
    const used = new Map<string, Set<string>>();
    dedupeFilename("/out/a", "page.md", used);
    expect(dedupeFilename("/out/b", "page.md", used)).toBe("page.md");
  });

  it("handles _index.md collisions", () => {
    const used = new Map<string, Set<string>>();
    dedupeFilename("/out", "_index.md", used);
    expect(dedupeFilename("/out", "_index.md", used)).toBe("_index-2.md");
  });
});
