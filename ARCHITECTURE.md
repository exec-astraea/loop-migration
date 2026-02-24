# Architecture: Fetching Data from Loop & SharePoint

## Overview

The tool exports Microsoft Loop workspaces to local Markdown files. It talks to **two separate services** using two different authentication tokens:

| Service | Base URL | Token env var | Purpose |
|---|---|---|---|
| **Loop API** | `substrate.office.com` | `LOOP_BEARER_TOKEN` | Workspace & page metadata |
| **SharePoint API** | `*.sharepoint.com` | `SHAREPOINT_BEARER_TOKEN` | Fluid snapshots (hierarchy) and page HTML content |

## Data flow

```
Loop API                  SharePoint API (v2.1)          SharePoint API (v2.0)
─────────                 ─────────────────────          ─────────────────────
  │                              │                              │
  │  1. deltasync                │                              │
  │  ──────────►                 │                              │
  │  workspaces + pages          │                              │
  │  (metadata only)             │                              │
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

## Step 1 — Loop delta-sync (metadata)

**Endpoint:** `GET https://substrate.office.com/recommended/api/beta/loop/deltasync?loopComponents=true`

**Auth:** Standard `Authorization: Bearer {token}` header.

**What it returns:** A JSON payload containing:
- `workspaces[]` — each with an `id`, `title`, and `mfs_info.pod_id` (base64-encoded pointer to the SharePoint backing store)
- `pages[]` — each with an `id`, `title`, `type`, `workspace_id`, `is_deleted`, `onedrive_info.drive_id`, and `sharepoint_info.site_url`
- `activities[]` — not used by this tool

**Pagination:** The response may include `next_page_link` (a query string) and `is_complete: false`. The tool follows these links, merging and deduplicating results by `id`, until the full dataset is collected.

**Key detail:** This step gives us page *metadata* but not the page *content* or *hierarchy* (folder structure). Those come from SharePoint.

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
    ├── sharepoint.mts     SP multipart auth helper (spGet)
    ├── loop-api.mts       Step 1: Loop delta-sync with pagination
    ├── hierarchy.mts      Step 2: Fluid snapshot → page tree → flat entries
    └── export.mts         Step 3: HTML fetch → Turndown → .md files
```
