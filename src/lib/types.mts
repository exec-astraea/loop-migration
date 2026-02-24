export interface Workspace {
  id: string;
  title?: string;
  mfs_info?: { pod_id?: string };
  [key: string]: unknown;
}

export interface LoopPage {
  id: string;
  title?: string;
  type?: string;
  workspace_id?: string;
  is_deleted?: boolean;
  onedrive_info?: { drive_id?: string };
  sharepoint_info?: { site_url?: string };
  [key: string]: unknown;
}

export interface LoopData {
  workspaces?: Workspace[];
  pages?: LoopPage[];
  activities?: Array<Record<string, unknown>>;
  next_page_link?: string;
  is_complete?: boolean;
  [key: string]: unknown;
}

export interface HierarchyNode {
  pageId: string;
  title: string;
  emoji: string | null;
  spoItemId: string | null;
  children: HierarchyNode[];
}

export interface FlatEntry {
  pageId: string;
  title: string;
  emoji: string | null;
  spoItemId: string | null;
  path: string;
  hasChildren: boolean;
}

export interface ExportResult {
  exported: number;
  skipped: number;
  skippedPages: string[];
}
