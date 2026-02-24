import { describe, it, expect } from "vitest";
import {
  slugify,
  decodePodId,
  findSharedTree,
  findBuild0,
  extractPageMeta,
  extractHierarchy,
  flattenWithPaths,
} from "../lib/hierarchy.mts";
import type { HierarchyNode } from "../lib/types.mts";

// ---------------------------------------------------------------------------
// slugify
// ---------------------------------------------------------------------------

describe("slugify", () => {
  it("lowercases and replaces spaces with hyphens", () => {
    expect(slugify("Hello World")).toBe("hello-world");
  });

  it("strips non-word characters except hyphens", () => {
    expect(slugify("What's up? (2025)")).toBe("whats-up-2025");
  });

  it("collapses multiple hyphens", () => {
    expect(slugify("a --- b")).toBe("a-b");
  });

  it("trims leading/trailing hyphens", () => {
    expect(slugify("  --hello--  ")).toBe("hello");
  });

  it("preserves underscores", () => {
    expect(slugify("my_page_title")).toBe("my_page_title");
  });

  it("handles empty string", () => {
    expect(slugify("")).toBe("untitled");
  });

  it("preserves emoji", () => {
    expect(slugify("🚀🔥")).toBe("🚀🔥");
  });

  it("preserves emoji with text", () => {
    expect(slugify("🎨 Design")).toBe("🎨-design");
  });
});

// ---------------------------------------------------------------------------
// decodePodId
// ---------------------------------------------------------------------------

describe("decodePodId", () => {
  it("decodes a valid base64 pod_id into host, driveId, itemId", () => {
    const raw = "prefix|contoso.sharepoint.com|driveABC|item123";
    const encoded = Buffer.from(raw).toString("base64");
    const result = decodePodId(encoded);
    expect(result).toEqual({
      host: "contoso.sharepoint.com",
      driveId: "driveABC",
      itemId: "item123",
    });
  });

  it("throws on fewer than 4 pipe-separated parts", () => {
    const encoded = Buffer.from("only|two").toString("base64");
    expect(() => decodePodId(encoded)).toThrow("expected >=4 parts");
  });

  it("handles extra pipe-separated parts gracefully", () => {
    const raw = "a|host|drive|item|extra|stuff";
    const encoded = Buffer.from(raw).toString("base64");
    const result = decodePodId(encoded);
    expect(result.host).toBe("host");
    expect(result.driveId).toBe("drive");
    expect(result.itemId).toBe("item");
  });
});

// ---------------------------------------------------------------------------
// findSharedTree
// ---------------------------------------------------------------------------

describe("findSharedTree", () => {
  function makeBlob(data: object, size = 5000) {
    return {
      size,
      content: Buffer.from(JSON.stringify(data)).toString("base64"),
    };
  }

  it("finds the blob with editHistory and internedStrings", () => {
    const tree = { editHistory: {}, internedStrings: ["a"] };
    const snapshot = {
      blobs: [
        makeBlob({ irrelevant: true }),
        makeBlob(tree),
      ],
    };
    expect(findSharedTree(snapshot)).toEqual(tree);
  });

  it("skips blobs smaller than 1000 bytes", () => {
    const tree = { editHistory: {}, internedStrings: [] };
    const snapshot = { blobs: [makeBlob(tree, 500)] };
    expect(() => findSharedTree(snapshot)).toThrow("No SharedTree blob found");
  });

  it("throws when no matching blob exists", () => {
    const snapshot = { blobs: [makeBlob({ nope: true })] };
    expect(() => findSharedTree(snapshot)).toThrow("No SharedTree blob found");
  });
});

// ---------------------------------------------------------------------------
// findBuild0
// ---------------------------------------------------------------------------

describe("findBuild0", () => {
  it("finds the first Build change (type 5) with multiple source nodes", () => {
    const source = [{}, {}]; // 2 nodes
    const sharedTree = {
      editHistory: {
        editChunks: [
          {
            chunk: [
              { changes: [{ type: 3 }, { type: 5, source: source }] },
            ],
          },
        ],
      },
    };
    expect(findBuild0(sharedTree)).toBe(source);
  });

  it("skips Build changes with only 1 source node", () => {
    const sharedTree = {
      editHistory: {
        editChunks: [
          { chunk: [{ changes: [{ type: 5, source: [{}] }] }] },
        ],
      },
    };
    expect(() => findBuild0(sharedTree)).toThrow("No multi-node Build change");
  });

  it("throws when no Build change exists", () => {
    const sharedTree = {
      editHistory: { editChunks: [{ chunk: [{ changes: [{ type: 1 }] }] }] },
    };
    expect(() => findBuild0(sharedTree)).toThrow("No multi-node Build change");
  });
});

// ---------------------------------------------------------------------------
// extractPageMeta
// ---------------------------------------------------------------------------

describe("extractPageMeta", () => {
  // Interned strings: index 0 = "displayText", 1 = "icon", 2 = "odspMetadata"
  const IS = ["displayText", "icon", "odspMetadata"];

  function makePageNode(id: string, title: string, emoji: string | null, itemId: string | null) {
    const traits: any[] = [];
    // displayText trait (IS index 0)
    traits.push(0, [[null, null, [title]]]);
    // icon trait (IS index 1)
    if (emoji) traits.push(1, [[null, null, [{ type: "emoji", data: emoji }]]]);
    // odspMetadata trait (IS index 2)
    if (itemId) traits.push(2, [[null, null, [{ itemId }]]]);

    return [null, null, [{ label: "LoopPage", id }, ...traits]];
  }

  it("extracts title, emoji, and itemId from LoopPage nodes", () => {
    const source = [
      makePageNode("page-1", "Meeting Notes", "📝", "sp-item-1"),
      makePageNode("page-2", "Design Doc", "🎨", "sp-item-2"),
    ];
    const meta = extractPageMeta(source, IS);
    expect(meta.size).toBe(2);
    expect(meta.get("page-1")).toEqual({ title: "Meeting Notes", emoji: "📝", itemId: "sp-item-1" });
    expect(meta.get("page-2")).toEqual({ title: "Design Doc", emoji: "🎨", itemId: "sp-item-2" });
  });

  it("defaults to '???' title when displayText is missing", () => {
    const node = [null, null, [{ label: "LoopPage", id: "p1" }]];
    const meta = extractPageMeta([node], IS);
    expect(meta.get("p1")?.title).toBe("???");
  });

  it("skips non-LoopPage nodes", () => {
    const node = [null, null, [{ label: "LoopWorkspace", id: "ws1" }]];
    const meta = extractPageMeta([node], IS);
    expect(meta.size).toBe(0);
  });

  it("skips nodes without array data", () => {
    const node = [null, null, "not-an-array"];
    const meta = extractPageMeta([node], IS);
    expect(meta.size).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// extractHierarchy
// ---------------------------------------------------------------------------

describe("extractHierarchy", () => {
  // IS: 0 = "values"
  const IS = ["values"];
  const valuesIdx = 0;

  function makeWorkspaceSource(children: any[]) {
    // source[0] = workspace node
    const wsNode = [null, null, [{ label: "LoopWorkspace" }, valuesIdx, children]];
    return [wsNode];
  }

  it("builds a flat hierarchy from workspace values", () => {
    const pageMeta = new Map([
      ["p1", { title: "Page One", emoji: null, itemId: "item1" }],
      ["p2", { title: "Page Two", emoji: "🔥", itemId: "item2" }],
    ]);
    // Each child: [ref, ref, [pageId]]
    const children = [
      [null, null, ["p1"]],
      [null, null, ["p2"]],
    ];
    const source = makeWorkspaceSource(children);
    const tree = extractHierarchy(source, IS, pageMeta);

    expect(tree).toHaveLength(2);
    expect(tree[0]).toMatchObject({ pageId: "p1", title: "Page One", children: [] });
    expect(tree[1]).toMatchObject({ pageId: "p2", title: "Page Two", emoji: "🔥" });
  });

  it("builds nested children using the values trait", () => {
    const pageMeta = new Map([
      ["parent", { title: "Parent", emoji: null, itemId: "ip" }],
      ["child", { title: "Child", emoji: null, itemId: "ic" }],
    ]);
    const childNode = [null, null, ["child"]];
    const parentNode = [null, null, ["parent", valuesIdx, [childNode]]];
    const source = makeWorkspaceSource([parentNode]);
    const tree = extractHierarchy(source, IS, pageMeta);

    expect(tree).toHaveLength(1);
    expect(tree[0].children).toHaveLength(1);
    expect(tree[0].children[0].pageId).toBe("child");
  });

  it("throws when first node is not LoopWorkspace", () => {
    const source = [[null, null, [{ label: "SomethingElse" }]]];
    expect(() => extractHierarchy(source, IS, new Map())).toThrow("Expected LoopWorkspace");
  });
});

// ---------------------------------------------------------------------------
// flattenWithPaths
// ---------------------------------------------------------------------------

describe("flattenWithPaths", () => {
  const leaf = (id: string, title: string, itemId: string): HierarchyNode => ({
    pageId: id, title, emoji: null, spoItemId: itemId, children: [],
  });

  it("produces slugified paths for leaf nodes", () => {
    const nodes = [leaf("1", "Hello World", "i1")];
    const flat = flattenWithPaths(nodes);
    expect(flat).toHaveLength(1);
    expect(flat[0].path).toBe("hello-world");
    expect(flat[0].hasChildren).toBe(false);
  });

  it("nests paths for children", () => {
    const parent: HierarchyNode = {
      pageId: "p", title: "Parent", emoji: null, spoItemId: "ip",
      children: [leaf("c", "Child Page", "ic")],
    };
    const flat = flattenWithPaths([parent]);
    expect(flat).toHaveLength(2);
    expect(flat[0].path).toBe("parent");
    expect(flat[0].hasChildren).toBe(true);
    expect(flat[1].path).toBe("parent/child-page");
  });

  it("handles deeply nested hierarchy", () => {
    const deep: HierarchyNode = {
      pageId: "a", title: "A", emoji: null, spoItemId: "ia",
      children: [{
        pageId: "b", title: "B", emoji: null, spoItemId: "ib",
        children: [leaf("c", "C", "ic")],
      }],
    };
    const flat = flattenWithPaths([deep]);
    expect(flat).toHaveLength(3);
    expect(flat[2].path).toBe("a/b/c");
  });

  it("returns empty array for empty input", () => {
    expect(flattenWithPaths([])).toEqual([]);
  });

  it("deduplicates sibling slugs", () => {
    const nodes: HierarchyNode[] = [
      { pageId: "1", title: "Q&A", emoji: null, spoItemId: "a", children: [] },
      { pageId: "2", title: "QA", emoji: null, spoItemId: "b", children: [] },
      { pageId: "3", title: "QA", emoji: null, spoItemId: "c", children: [] },
    ];
    const flat = flattenWithPaths(nodes);
    expect(flat[0].path).toBe("qa");
    expect(flat[1].path).toBe("qa-2");
    expect(flat[2].path).toBe("qa-3");
  });
});
