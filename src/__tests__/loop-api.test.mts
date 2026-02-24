import { describe, it, expect } from "vitest";
import { buildDedupKey, mergeArrays, mergeLoopData } from "../lib/loop-api.mts";
import type { LoopData } from "../lib/types.mts";

// ---------------------------------------------------------------------------
// buildDedupKey
// ---------------------------------------------------------------------------

describe("buildDedupKey", () => {
  it("uses the 'id' field when present", () => {
    expect(buildDedupKey({ id: "abc" }, "ws")).toBe("id:abc");
  });

  it("prefers 'id' over other key fields", () => {
    expect(buildDedupKey({ id: "1", workspace_id: "2" }, "p")).toBe("id:1");
  });

  it("falls back to workspace_id", () => {
    expect(buildDedupKey({ workspace_id: "ws1" }, "p")).toBe("workspace_id:ws1");
  });

  it("falls back to web_url", () => {
    expect(buildDedupKey({ web_url: "https://example.com" }, "p")).toBe("web_url:https://example.com");
  });

  it("falls back to JSON stringification when no known key exists", () => {
    const item = { foo: "bar" };
    expect(buildDedupKey(item, "pfx")).toBe(`pfx:${JSON.stringify(item)}`);
  });

  it("skips empty string values for known keys", () => {
    expect(buildDedupKey({ id: "", url: "http://x" }, "p")).toBe("url:http://x");
  });
});

// ---------------------------------------------------------------------------
// mergeArrays
// ---------------------------------------------------------------------------

describe("mergeArrays", () => {
  it("merges two arrays, deduplicating by id", () => {
    const a = [{ id: "1", name: "A" }];
    const b = [{ id: "1", name: "A-dup" }, { id: "2", name: "B" }];
    const result = mergeArrays(a, b, "test");
    expect(result).toHaveLength(2);
    expect(result[0]).toEqual({ id: "1", name: "A" }); // keeps first
    expect(result[1]).toEqual({ id: "2", name: "B" });
  });

  it("handles undefined inputs", () => {
    expect(mergeArrays(undefined, [{ id: "1" }], "t")).toHaveLength(1);
    expect(mergeArrays([{ id: "1" }], undefined, "t")).toHaveLength(1);
    expect(mergeArrays(undefined, undefined, "t")).toHaveLength(0);
  });

  it("does not mutate the original arrays", () => {
    const a = [{ id: "1" }];
    const b = [{ id: "2" }];
    const result = mergeArrays(a, b, "t");
    expect(result).toHaveLength(2);
    expect(a).toHaveLength(1);
    expect(b).toHaveLength(1);
  });
});

// ---------------------------------------------------------------------------
// mergeLoopData
// ---------------------------------------------------------------------------

describe("mergeLoopData", () => {
  const base: LoopData = {
    workspaces: [{ id: "ws1", title: "WS1" }],
    pages: [{ id: "p1", title: "Page1" }],
    activities: [{ id: "a1" }],
    is_complete: false,
  };

  const inc: LoopData = {
    workspaces: [{ id: "ws1", title: "WS1" }, { id: "ws2", title: "WS2" }],
    pages: [{ id: "p2", title: "Page2" }],
    activities: [{ id: "a2" }],
    is_complete: true,
    next_page_link: undefined,
  };

  it("deduplicates workspaces across pages", () => {
    const merged = mergeLoopData(base, inc);
    expect(merged.workspaces).toHaveLength(2);
  });

  it("merges pages from both responses", () => {
    const merged = mergeLoopData(base, inc);
    expect(merged.pages).toHaveLength(2);
  });

  it("merges activities", () => {
    const merged = mergeLoopData(base, inc);
    expect(merged.activities).toHaveLength(2);
  });

  it("takes scalar fields from the incremental response", () => {
    const merged = mergeLoopData(base, inc);
    expect(merged.is_complete).toBe(true);
  });
});
