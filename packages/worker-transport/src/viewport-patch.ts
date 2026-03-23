import type { CellSnapshot, RecalcMetrics, Viewport } from "@bilig/protocol";

export interface ViewportPatchSubscription extends Viewport {
  sheetName: string;
}

export interface ViewportPatchedCell {
  row: number;
  col: number;
  snapshot: CellSnapshot;
  displayText: string;
  copyText: string;
  editorText: string;
  formatId: number;
  styleId: number;
}

export interface ViewportAxisPatch {
  index: number;
  size: number;
  hidden: boolean;
}

export interface ViewportPatch {
  version: number;
  full: boolean;
  viewport: ViewportPatchSubscription;
  metrics: RecalcMetrics;
  cells: ViewportPatchedCell[];
  columns: ViewportAxisPatch[];
  rows: ViewportAxisPatch[];
}

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isViewportPatch(value: unknown): value is ViewportPatch {
  if (!isRecord(value) || !isRecord(value["viewport"]) || !isRecord(value["metrics"])) {
    return false;
  }
  return Array.isArray(value["cells"]) && Array.isArray(value["columns"]) && Array.isArray(value["rows"]);
}

export function encodeViewportPatch(patch: ViewportPatch): Uint8Array {
  return encoder.encode(JSON.stringify(patch));
}

export function decodeViewportPatch(bytes: Uint8Array): ViewportPatch {
  const parsed = JSON.parse(decoder.decode(bytes)) as unknown;
  if (!isViewportPatch(parsed)) {
    throw new Error("Invalid viewport patch payload");
  }
  return parsed;
}
