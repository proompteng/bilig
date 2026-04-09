import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStyleField,
  CellStylePatch,
  LiteralInput,
  WorkbookSnapshot,
} from "@bilig/protocol";
import { isCellRangeRef, isWorkbookSnapshot } from "@bilig/protocol";
import { parseCellAddress } from "@bilig/formula";
import {
  applyWorkbookAgentCommandBundle,
  isWorkbookAgentCommandBundle,
  type WorkbookAgentCommandBundle,
} from "@bilig/agent-api";
import { isCommitOps, type CommitOp, type SpreadsheetEngine } from "@bilig/core";
import {
  isEngineOpBatch,
  isEngineOps,
  type EngineOp,
  type EngineOpBatch,
} from "@bilig/workbook-domain";

export type WorkbookChangeUndoBundle =
  | {
      kind: "engineOps";
      ops: EngineOp[];
    }
  | {
      kind: "snapshot";
      snapshot: WorkbookSnapshot;
    };

export interface DirtyRegion {
  sheetName: string;
  rowStart: number;
  rowEnd: number;
  colStart: number;
  colEnd: number;
}

export type WorkbookEventPayload =
  | {
      kind: "applyBatch";
      batch: EngineOpBatch;
    }
  | {
      kind: "applyAgentCommandBundle";
      bundle: WorkbookAgentCommandBundle;
    }
  | {
      kind: "setCellValue";
      sheetName: string;
      address: string;
      value: LiteralInput;
    }
  | {
      kind: "setCellFormula";
      sheetName: string;
      address: string;
      formula: string;
    }
  | {
      kind: "clearCell";
      sheetName: string;
      address: string;
    }
  | {
      kind: "clearRange";
      range: CellRangeRef;
    }
  | {
      kind: "renderCommit";
      ops: CommitOp[];
    }
  | {
      kind: "fillRange";
      source: CellRangeRef;
      target: CellRangeRef;
    }
  | {
      kind: "copyRange";
      source: CellRangeRef;
      target: CellRangeRef;
    }
  | {
      kind: "moveRange";
      source: CellRangeRef;
      target: CellRangeRef;
    }
  | {
      kind: "updateRowMetadata";
      sheetName: string;
      startRow: number;
      count: number;
      height: number | null;
      hidden: boolean | null;
    }
  | {
      kind: "updateColumnMetadata";
      sheetName: string;
      startCol: number;
      count: number;
      width: number | null;
      hidden: boolean | null;
    }
  | {
      kind: "updateColumnWidth";
      sheetName: string;
      columnIndex: number;
      width: number;
    }
  | {
      kind: "setFreezePane";
      sheetName: string;
      rows: number;
      cols: number;
    }
  | {
      kind: "setRangeStyle";
      range: CellRangeRef;
      patch: CellStylePatch;
    }
  | {
      kind: "clearRangeStyle";
      range: CellRangeRef;
      fields?: readonly CellStyleField[];
    }
  | {
      kind: "setRangeNumberFormat";
      range: CellRangeRef;
      format: CellNumberFormatInput;
    }
  | {
      kind: "clearRangeNumberFormat";
      range: CellRangeRef;
    }
  | {
      kind: "restoreVersion";
      versionId: string;
      versionName: string;
      sheetName?: string;
      address?: string;
      snapshot: WorkbookSnapshot;
    }
  | {
      kind: "revertChange";
      targetRevision: number;
      targetSummary: string;
      sheetName?: string;
      address?: string;
      range?: CellRangeRef;
      appliedBundle: WorkbookChangeUndoBundle;
    }
  | {
      kind: "redoChange";
      targetRevision: number;
      targetSummary: string;
      sheetName?: string;
      address?: string;
      range?: CellRangeRef;
      appliedBundle: WorkbookChangeUndoBundle;
    };

export interface WorkbookEventRecord {
  workbookId: string;
  revision: number;
  actorUserId: string;
  clientMutationId: string | null;
  payload: WorkbookEventPayload;
  createdAt: string;
}

export interface AuthoritativeWorkbookEventRecord {
  revision: number;
  clientMutationId: string | null;
  payload: WorkbookEventPayload;
}

export interface AuthoritativeWorkbookEventBatch {
  afterRevision: number;
  headRevision: number;
  calculatedRevision: number;
  events: readonly AuthoritativeWorkbookEventRecord[];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === "string");
}

export function isWorkbookChangeUndoBundle(value: unknown): value is WorkbookChangeUndoBundle {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }
  switch (value["kind"]) {
    case "engineOps":
      return isEngineOps(value["ops"]);
    case "snapshot":
      return isWorkbookSnapshot(value["snapshot"]);
    default:
      return false;
  }
}

export function isWorkbookEventPayload(value: unknown): value is WorkbookEventPayload {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }

  switch (value["kind"]) {
    case "applyBatch":
      return isEngineOpBatch(value["batch"]);
    case "applyAgentCommandBundle":
      return isWorkbookAgentCommandBundle(value["bundle"]);
    case "setCellValue":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["address"] === "string" &&
        value["value"] !== undefined
      );
    case "setCellFormula":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["address"] === "string" &&
        typeof value["formula"] === "string"
      );
    case "clearCell":
      return typeof value["sheetName"] === "string" && typeof value["address"] === "string";
    case "clearRange":
      return isCellRangeRef(value["range"]);
    case "renderCommit":
      return isCommitOps(value["ops"]);
    case "fillRange":
    case "copyRange":
    case "moveRange":
      return isCellRangeRef(value["source"]) && isCellRangeRef(value["target"]);
    case "updateRowMetadata":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["startRow"] === "number" &&
        typeof value["count"] === "number" &&
        (typeof value["height"] === "number" || value["height"] === null) &&
        (typeof value["hidden"] === "boolean" || value["hidden"] === null)
      );
    case "updateColumnMetadata":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["startCol"] === "number" &&
        typeof value["count"] === "number" &&
        (typeof value["width"] === "number" || value["width"] === null) &&
        (typeof value["hidden"] === "boolean" || value["hidden"] === null)
      );
    case "updateColumnWidth":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["columnIndex"] === "number" &&
        typeof value["width"] === "number"
      );
    case "setFreezePane":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["rows"] === "number" &&
        typeof value["cols"] === "number"
      );
    case "setRangeStyle":
      return isCellRangeRef(value["range"]) && typeof value["patch"] === "object";
    case "clearRangeStyle":
      return (
        isCellRangeRef(value["range"]) &&
        (value["fields"] === undefined || isStringArray(value["fields"]))
      );
    case "setRangeNumberFormat":
      return (
        isCellRangeRef(value["range"]) &&
        (typeof value["format"] === "string" ||
          (typeof value["format"] === "object" && value["format"] !== null))
      );
    case "clearRangeNumberFormat":
      return isCellRangeRef(value["range"]);
    case "restoreVersion":
      return (
        typeof value["versionId"] === "string" &&
        typeof value["versionName"] === "string" &&
        (value["sheetName"] === undefined || typeof value["sheetName"] === "string") &&
        (value["address"] === undefined || typeof value["address"] === "string") &&
        isWorkbookSnapshot(value["snapshot"])
      );
    case "revertChange":
    case "redoChange":
      return (
        typeof value["targetRevision"] === "number" &&
        typeof value["targetSummary"] === "string" &&
        (value["sheetName"] === undefined || typeof value["sheetName"] === "string") &&
        (value["address"] === undefined || typeof value["address"] === "string") &&
        (value["range"] === undefined || isCellRangeRef(value["range"])) &&
        isWorkbookChangeUndoBundle(value["appliedBundle"])
      );
    default:
      return false;
  }
}

export function isAuthoritativeWorkbookEventRecord(
  value: unknown,
): value is AuthoritativeWorkbookEventRecord {
  return (
    isRecord(value) &&
    typeof value["revision"] === "number" &&
    (typeof value["clientMutationId"] === "string" || value["clientMutationId"] === null) &&
    isWorkbookEventPayload(value["payload"])
  );
}

export function isAuthoritativeWorkbookEventBatch(
  value: unknown,
): value is AuthoritativeWorkbookEventBatch {
  return (
    isRecord(value) &&
    typeof value["afterRevision"] === "number" &&
    typeof value["headRevision"] === "number" &&
    typeof value["calculatedRevision"] === "number" &&
    Array.isArray(value["events"]) &&
    value["events"].every((event) => isAuthoritativeWorkbookEventRecord(event))
  );
}

function singleCellRegion(sheetName: string, address: string): DirtyRegion {
  const parsed = parseCellAddress(address, sheetName);
  return {
    sheetName,
    rowStart: parsed.row,
    rowEnd: parsed.row,
    colStart: parsed.col,
    colEnd: parsed.col,
  };
}

function rangeRegion(range: CellRangeRef): DirtyRegion {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    rowStart: Math.min(start.row, end.row),
    rowEnd: Math.max(start.row, end.row),
    colStart: Math.min(start.col, end.col),
    colEnd: Math.max(start.col, end.col),
  };
}

export function deriveDirtyRegions(payload: WorkbookEventPayload): DirtyRegion[] | null {
  switch (payload.kind) {
    case "setCellValue":
    case "setCellFormula":
    case "clearCell":
      return [singleCellRegion(payload.sheetName, payload.address)];
    case "clearRange":
      return [rangeRegion(payload.range)];
    case "fillRange":
    case "copyRange":
    case "moveRange":
      return [rangeRegion(payload.source), rangeRegion(payload.target)];
    case "setRangeStyle":
    case "clearRangeStyle":
    case "setRangeNumberFormat":
    case "clearRangeNumberFormat":
      return [rangeRegion(payload.range)];
    case "applyBatch":
    case "applyAgentCommandBundle":
    case "renderCommit":
    case "updateRowMetadata":
    case "updateColumnMetadata":
    case "updateColumnWidth":
    case "setFreezePane":
    case "restoreVersion":
    case "revertChange":
    case "redoChange":
      return null;
    default: {
      const exhaustive: never = payload;
      return exhaustive;
    }
  }
}

export function applyWorkbookEvent(engine: SpreadsheetEngine, payload: WorkbookEventPayload): void {
  switch (payload.kind) {
    case "applyBatch":
      engine.applyRemoteBatch(payload.batch);
      return;
    case "applyAgentCommandBundle":
      applyWorkbookAgentCommandBundle(engine, payload.bundle);
      return;
    case "setCellValue":
      engine.setCellValue(payload.sheetName, payload.address, payload.value);
      return;
    case "setCellFormula":
      engine.setCellFormula(payload.sheetName, payload.address, payload.formula);
      return;
    case "clearCell":
      engine.clearCell(payload.sheetName, payload.address);
      return;
    case "clearRange":
      engine.clearRange(payload.range);
      return;
    case "renderCommit":
      engine.renderCommit(payload.ops);
      return;
    case "fillRange":
      engine.fillRange(payload.source, payload.target);
      return;
    case "copyRange":
      engine.copyRange(payload.source, payload.target);
      return;
    case "moveRange":
      engine.moveRange(payload.source, payload.target);
      return;
    case "updateRowMetadata":
      engine.updateRowMetadata(
        payload.sheetName,
        payload.startRow,
        payload.count,
        payload.height,
        payload.hidden,
      );
      return;
    case "updateColumnMetadata":
      engine.updateColumnMetadata(
        payload.sheetName,
        payload.startCol,
        payload.count,
        payload.width,
        payload.hidden,
      );
      return;
    case "updateColumnWidth":
      engine.updateColumnMetadata(payload.sheetName, payload.columnIndex, 1, payload.width, null);
      return;
    case "setFreezePane":
      engine.setFreezePane(payload.sheetName, payload.rows, payload.cols);
      return;
    case "setRangeStyle":
      engine.setRangeStyle(payload.range, payload.patch);
      return;
    case "clearRangeStyle":
      engine.clearRangeStyle(payload.range, payload.fields);
      return;
    case "setRangeNumberFormat":
      engine.setRangeNumberFormat(payload.range, payload.format);
      return;
    case "clearRangeNumberFormat":
      engine.clearRangeNumberFormat(payload.range);
      return;
    case "restoreVersion":
      engine.importSnapshot(payload.snapshot);
      return;
    case "revertChange":
    case "redoChange":
      if (payload.appliedBundle.kind === "engineOps") {
        engine.applyOps(payload.appliedBundle.ops);
        return;
      }
      engine.importSnapshot(payload.appliedBundle.snapshot);
      return;
    default: {
      const exhaustive: never = payload;
      throw new Error(`Unhandled workbook event: ${JSON.stringify(exhaustive)}`);
    }
  }
}
