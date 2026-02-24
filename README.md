# Loop Migration

Export Microsoft Loop workspaces to Markdown files, preserving the sidebar page hierarchy (nested folders, titles, emoji).

## Prerequisites

- Node.js 22+
- An active Microsoft 365 account with access to the Loop workspaces you want to export

## Setup

```bash
npm install
```

Create a `.env` file with two tokens (see [Getting tokens](#getting-tokens) below):

```env
LOOP_BEARER_TOKEN=Bearer eyJ0eXAi...
SHAREPOINT_BEARER_TOKEN=Bearer eyJ0eXAi...
```

## Usage

```bash
npm start                          # interactive picker → choose workspace or export all
npm start -- -w UFC                # select workspace by name → workspaces/ufc/
npm start -- --all                 # export all workspaces   → workspaces/*/
npm start -- -w UFC -d 0           # no delay between requests
npm start -- -n                    # dry run (shows what would be exported)
```

This runs the full pipeline in a single command with no intermediate files:

1. Fetches workspace & page metadata from the Loop API
2. Fetches the Fluid snapshot and extracts the sidebar page hierarchy
3. Downloads each page as HTML, converts to Markdown, writes to `workspaces/`

Deleted pages and shared-with-me pages are automatically skipped.

| Flag | Long form | Description | Default |
|------|-----------|-------------|---------|
| `-w` | `--workspace` | Select workspace by name or ID | interactive picker |
| `-a` | `--all` | Export all workspaces | off |
| `-p` | `--pick-workspace` | Force interactive workspace picker | off |
| `-d` | `--delay` | Delay (ms) between page requests | `50` |
| `-n` | `--dry-run` | Show what would be exported without fetching | off |

The interactive picker (shown when no flags are given, or with `-p`) includes an **"Export all workspaces"** option at position `0`.

## Getting tokens

Both tokens are short-lived OAuth Bearer tokens that you capture from your browser's dev tools. They typically expire after ~1 hour, so grab fresh ones right before running the pipeline.

### Loop Bearer Token (`LOOP_BEARER_TOKEN`)

1. Open [loop.cloud.microsoft](https://loop.cloud.microsoft) and sign in
2. Open DevTools → Network tab
3. Filter requests by `substrate.office.com`
4. Look for requests to `deltasync` — click one
5. In the request headers, copy the full `Authorization` header value (starts with `Bearer eyJ...`)

### SharePoint Bearer Token (`SHAREPOINT_BEARER_TOKEN`)

1. Open any Loop page in your browser
2. Open DevTools → Network tab
3. Filter requests by your SharePoint domain (e.g. `yourtenant.sharepoint.com`)
4. Look for requests to `opStream` or `content` endpoints
5. In the request headers (or multipart body), copy the `Authorization: Bearer eyJ...` value

> **Tip:** Both tokens may sometimes be the same token, but they target different APIs (Substrate vs SharePoint) so they may differ depending on your tenant configuration.

## Output structure

All output goes under `workspaces/`, with each workspace in its own subdirectory — even when only one workspace is exported.

```
workspaces/
└── my-workspace/          ← slugified workspace name
    ├── team-life/
    │   ├── _index.md
    │   ├── daily-responsibilities.md
    ├── meeting-notes/
    │   ├── _index.md
    │   ├── 2025-q2/
    │   │   ├── _index.md
    │   │   └── 2025-04-23.md
    │   └── ...
    └── ...
```

- Section pages (with children) → `folder/_index.md`
- Leaf pages → `folder/slugified-title.md`

Re-running an export clears the workspace subdirectory before writing. Exporting a workspace for the first time into a fresh directory skips the clear.
