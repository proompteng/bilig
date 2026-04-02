import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStyleField,
  CellStylePatch,
  LiteralInput,
} from "@bilig/protocol";
import { applyBatchArgsSchema } from "@bilig/zero-sync";
import type { CommitOp, SpreadsheetEngine } from "@bilig/core";
import { parseCellAddress } from "@bilig/formula";
import { z } from "zod";

type EngineOpBatch = z.infer<typeof applyBatchArgsSchema>["batch"];

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
      kind: "updateColumnWidth";
      sheetName: string;
      columnIndex: number;
      width: number;
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
    };

export interface WorkbookEventRecord {
  workbookId: string;
  revision: number;
  actorUserId: string;
  clientMutationId: string | null;
  payload: WorkbookEventPayload;
  createdAt: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === "string");
}

function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["startAddress"] === "string" &&
    typeof value["endAddress"] === "string"
  );
}

export function isWorkbookEventPayload(value: unknown): value is WorkbookEventPayload {
  if (!isRecord(value) || typeof value["kind"] !== "string") {
    return false;
  }

  switch (value["kind"]) {
    case "applyBatch":
      return Array.isArray(value["batch"]);
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
      return Array.isArray(value["ops"]);
    case "fillRange":
    case "copyRange":
      return isCellRangeRef(value["source"]) && isCellRangeRef(value["target"]);
    case "updateColumnWidth":
      return (
        typeof value["sheetName"] === "string" &&
        typeof value["columnIndex"] === "number" &&
        typeof value["width"] === "number"
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
    default:
      return false;
  }
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
      return [rangeRegion(payload.source), rangeRegion(payload.target)];
    case "setRangeStyle":
    case "clearRangeStyle":
    case "setRangeNumberFormat":
    case "clearRangeNumberFormat":
      return [rangeRegion(payload.range)];
    case "applyBatch":
    case "renderCommit":
    case "updateColumnWidth":
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
    case "updateColumnWidth":
      engine.updateColumnMetadata(payload.sheetName, payload.columnIndex, 1, payload.width, null);
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
    default: {
      const exhaustive: never = payload;
      throw new Error(`Unhandled workbook event: ${JSON.stringify(exhaustive)}`);
    }
  }
}
