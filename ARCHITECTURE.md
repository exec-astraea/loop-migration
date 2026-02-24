# Architecture: Fetching Data from Loop & SharePoint

## Overview

The tool exports Microsoft Loop workspaces to local Markdown files. It talks to **two separate services** using two different authentication tokens:

| Service | Base URL | Token env var | Purpose |
|---|---|---|---|
| **Loop API** | `substrate.office.com` | `LOOP_BEARER_TOKEN` | Workspace & page metadata |
| **SharePoint API** | `*.sharepoint.com` | `SHAREPOINT_BEARER_TOKEN` | Fluid snapshots (hierarchy) and page HTML content |

Both tokens are validated at startup by `getConfig()` (in `config.mts`), which reads from `.env` via dotenv and memoizes the result.

## Data flow

```
Loop API (v1.1)           SharePoint API (v2.1)          SharePoint API (v2.0)
───────────────           ─────────────────────          ─────────────────────
  │                              │                              │
  │  1a. /workspaces             │                              │
  │  ───────────────►            │                              │
  │  canonical workspace list    │                              │
  │  (incl. Personal workspace)  │                              │
  │                              │                              │
  │  1b. /recent                 │                              │
  │  ───────────────►            │                              │
  │  recently-active workspaces  │                              │
  │  + pages                     │                              │
  │                              │                              │
  │  1c. /deltasync              │                              │
  │  ───────────────►            │                              │
  │  full component graph        │                              │
  │  + pages                     │                              │
  │                              │                              │
  │  ── merge & deduplicate ──   │                              │
  │                              │                              │
  │                     2. opStream/snapshots                   │
  │                     ────────────────────►                   │
  │                     Fluid snapshot (JSON)                   │
  │                     → page hierarchy tree                   │
  │                                                             │
  │                                          3. items/{id}/content
  │                                          ─────────────────────►
  │                                          Page HTML (per page)
  │                                          → converted to Markdown
```

## Step 1 — Loop workspace & page discovery

No single Loop API endpoint returns all workspaces and pages. The tool queries **three** endpoints and merges the results, deduplicating by `id`:

### 1a. `/workspaces` — canonical workspace list

**Endpoint:** `GET https://substrate.office.com/recommended/api/v1.1/loop/workspaces?rs=en-us`

Returns the user's workspace list. This is the only endpoint that reliably includes the **Personal workspace** (titled "My workspace" by the API). It returns workspace metadata only — no pages.

### 1b. `/recent` — recently-active workspaces + pages

**Endpoint:** `GET https://substrate.office.com/recommended/api/v1.1/loop/recent?top=30&settings=true&rs=en-us`

Returns workspaces and pages ordered by recent activity. Newly created workspaces typically appear here first. Capped at 30 items per request.

### 1c. `/deltasync` — full component graph

**Endpoint:** `GET https://substrate.office.com/recommended/api/v1.1/loop/deltasync?loopComponents=true&rs=en-us`

Returns the full workspace + page + component graph. May include workspaces not touched recently. This is where the bulk of the page metadata comes from.

### Merge logic

All three responses share the same `LoopData` shape:
- `workspaces[]` — each with an `id`, `title`, and `mfs_info.pod_id` (base64-encoded pointer to the SharePoint backing store)
- `pages[]` — each with an `id`, `title`, `type`, `workspace_id`, `is_deleted`, `onedrive_info.drive_id`, and `sharepoint_info.site_url`
- `activities[]` — not used by this tool

**Auth:** Standard `Authorization: Bearer {token}` header.

**Pagination:** Each response may include `next_page_link` (a query string). The tool follows these links, merging and deduplicating results by `id`, until no more links are returned.

**Key detail:** This step gives us page *metadata* but not the page *content* or *hierarchy* (folder structure). Those come from SharePoint. Workspaces without `mfs_info.pod_id` are filtered out with a warning (they are typically empty or not yet synced to SharePoint).

## Step 2 — Fluid snapshot (hierarchy)

Each Loop workspace is backed by a Fluid Framework container stored in SharePoint. The workspace's `mfs_info.pod_id` is a base64-encoded string of the format `…|{host}|{driveId}|{itemId}` that points to this container.

**Endpoint:** `GET https://{host}/_api/v2.1/drives/{driveId}/items/{itemId}/opStream/snapshots/trees/latest?ump=1`

**Auth:** SharePoint's multipart POST convention (see below).

**What it returns:** A JSON snapshot of the Fluid container, containing an array of `blobs` with base64-encoded content.

### Extracting the hierarchy from the snapshot

1. **Find the SharedTree blob** — scan the blobs for one whose decoded JSON has both `editHistory` and `internedStrings` properties.

2. **Find Build 0** — walk `editHistory.editChunks[*].chunk[*].changes[*]` looking for a change with `type: 5` (a Build change) whose `source` array has more than one node. This is the initial state of the workspace tree.

3. **Extract page metadata** — Build 0's `source` array contains nodes. Each node with a `LoopPage` label has traits:
   - `displayText` → page title
   - `icon` → emoji (e.g. `{ type: "emoji", data: "🚀" }`)
   - `odspMetadata` → `{ itemId }` (the SharePoint item ID, used to match against Loop page IDs)

4. **Extract the tree** — The first node in the source array is a `LoopWorkspace` node. Its `values` trait contains the ordered list of top-level pages. Each page entry may recursively contain child pages in its own `values` trait, forming the full hierarchy.

5. **Flatten** — The tree is flattened into a list of `{ pageId, title, emoji, spoItemId, path }` entries where `path` is a slugified filesystem path like `design/components/button-styles`.

## Step 3 — Page HTML export

For each non-deleted, non-shared page in the selected workspace:

**Endpoint:** `GET https://{spHost}/_api/v2.0/drives/{driveId}/items/{itemId}/content?format=html&ump=1`

- `spHost` is derived from the page's `sharepoint_info.site_url`
- `driveId` comes from `onedrive_info.drive_id`
- `itemId` is extracted from the Loop page ID (the segment after the last `_`)

**Auth:** Same SharePoint multipart POST convention.

**What it returns:** The full page content as HTML.

The HTML is then converted to Markdown using [Turndown](https://github.com/mixmark-io/turndown) and written to disk at the path determined by the hierarchy from Step 2.

## SharePoint authentication

SharePoint endpoints don't accept a normal `Authorization` header from browser-origin requests. Instead, the tool uses a **multipart form POST** that smuggles the credentials:

```http
POST {url}
Content-Type: multipart/form-data;boundary={uuid}
Origin: https://loop.cloud.microsoft
Referer: https://loop.cloud.microsoft/

--{uuid}
Authorization: Bearer {token}
X-HTTP-Method-Override: GET
_post: 1

--{uuid}--
```

This is the same mechanism the Loop web client uses. The `X-HTTP-Method-Override: GET` header inside the multipart body tells SharePoint to treat this POST as a GET. Both the snapshot (v2.1) and content (v2.0) endpoints use this pattern.

## Module structure

```
src/
├── main.mts              CLI entry point, arg parsing, orchestration
└── lib/
    ├── types.mts          Shared interfaces (Workspace, LoopPage, LoopData, …)
    ├── config.mts         Token validation & memoization (getConfig)
    ├── sharepoint.mts     SP multipart auth helper (spGet) with retry + backoff
    ├── loop-api.mts       Step 1: /workspaces + /recent + /deltasync with merge
    ├── hierarchy.mts      Step 2: Fluid snapshot → page tree → flat entries
    └── export.mts         Step 3: HTML fetch → Turndown → .md files
```

## CLI flags

| Flag | Description |
|---|---|
| `-w, --workspace NAME` | Select workspace by name or ID (also accepts "Personal workspace") |
| `-a, --all` | Export all workspaces into subdirectories |
| `-p, --pick-workspace` | Interactive numbered picker |
| `--page TITLE` | Export a single page by substring match |
| `--dump-html` | Save raw HTML alongside markdown |
| `-d, --delay MS` | Delay between page requests (default: 50) |
| `-n, --dry-run` | Show what would be exported without fetching |
| `-h, --help` | Show help |
