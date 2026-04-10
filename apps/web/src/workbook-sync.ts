import { isCommitOps } from "@bilig/core";
import type { Zero } from "@rocicorp/zero";
import { createRenderCommitArgs, mutators } from "@bilig/zero-sync";
import type { CellNumberFormatInput, CellStyleField, CellStylePatch } from "@bilig/protocol";
import { isCellRangeRef, isLiteralInput } from "@bilig/protocol";

export type WorkbookMutationMethod =
  | "setCellValue"
  | "setCellFormula"
  | "clearCell"
  | "clearRange"
  | "renderCommit"
  | "fillRange"
  | "copyRange"
  | "moveRange"
  | "insertRows"
  | "deleteRows"
  | "insertColumns"
  | "deleteColumns"
  | "updateRowMetadata"
  | "updateColumnMetadata"
  | "setFreezePane"
  | "setRangeStyle"
  | "clearRangeStyle"
  | "setRangeNumberFormat"
  | "clearRangeNumberFormat";

type KnownPendingWorkbookMutationMethod = WorkbookMutationMethod | "updateColumnWidth";

export interface PendingWorkbookMutationInput {
  readonly method: KnownPendingWorkbookMutationMethod;
  readonly args: unknown[];
}

export interface PendingWorkbookMutation extends PendingWorkbookMutationInput {
  readonly id: string;
  readonly localSeq: number;
  readonly baseRevision: number;
  readonly enqueuedAtUnixMs: number;
  readonly submittedAtUnixMs: number | null;
  readonly lastAttemptedAtUnixMs: number | null;
  readonly ackedAtUnixMs: number | null;
  readonly rebasedAtUnixMs: number | null;
  readonly failedAtUnixMs: number | null;
  readonly attemptCount: number;
  readonly failureMessage: string | null;
  readonly status: "local" | "submitted" | "acked" | "rebased" | "failed";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function isCellStyleFieldList(value: unknown): value is CellStyleField[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === "string");
}

export function isCellStylePatchValue(value: unknown): value is CellStylePatch {
  return isRecord(value);
}

export function isCellNumberFormatInputValue(value: unknown): value is CellNumberFormatInput {
  return typeof value === "string" || isRecord(value);
}

export { isCellRangeRef, isCommitOps, isLiteralInput };

export function isWorkbookMutationMethod(value: unknown): value is WorkbookMutationMethod {
  return (
    value === "setCellValue" ||
    value === "setCellFormula" ||
    value === "clearCell" ||
    value === "clearRange" ||
    value === "renderCommit" ||
    value === "fillRange" ||
    value === "copyRange" ||
    value === "moveRange" ||
    value === "insertRows" ||
    value === "deleteRows" ||
    value === "insertColumns" ||
    value === "deleteColumns" ||
    value === "updateRowMetadata" ||
    value === "updateColumnMetadata" ||
    value === "setFreezePane" ||
    value === "setRangeStyle" ||
    value === "clearRangeStyle" ||
    value === "setRangeNumberFormat" ||
    value === "clearRangeNumberFormat"
  );
}

function isKnownPendingWorkbookMutationMethod(
  value: unknown,
): value is KnownPendingWorkbookMutationMethod {
  return value === "updateColumnWidth" || isWorkbookMutationMethod(value);
}

export function isPendingWorkbookMutationInput(
  value: unknown,
): value is PendingWorkbookMutationInput {
  return (
    isRecord(value) &&
    isKnownPendingWorkbookMutationMethod(value["method"]) &&
    Array.isArray(value["args"])
  );
}

export function isPendingWorkbookMutation(value: unknown): value is PendingWorkbookMutation {
  return (
    isRecord(value) &&
    typeof value["id"] === "string" &&
    typeof value["localSeq"] === "number" &&
    typeof value["baseRevision"] === "number" &&
    typeof value["enqueuedAtUnixMs"] === "number" &&
    (value["submittedAtUnixMs"] === null || typeof value["submittedAtUnixMs"] === "number") &&
    (value["lastAttemptedAtUnixMs"] === null ||
      typeof value["lastAttemptedAtUnixMs"] === "number") &&
    (value["ackedAtUnixMs"] === null || typeof value["ackedAtUnixMs"] === "number") &&
    (value["rebasedAtUnixMs"] === null || typeof value["rebasedAtUnixMs"] === "number") &&
    (value["failedAtUnixMs"] === null || typeof value["failedAtUnixMs"] === "number") &&
    typeof value["attemptCount"] === "number" &&
    (value["failureMessage"] === null || typeof value["failureMessage"] === "string") &&
    (value["status"] === "local" ||
      value["status"] === "submitted" ||
      value["status"] === "acked" ||
      value["status"] === "rebased" ||
      value["status"] === "failed") &&
    isPendingWorkbookMutationInput(value)
  );
}

export function isPendingWorkbookMutationList(
  value: unknown,
): value is readonly PendingWorkbookMutation[] {
  return Array.isArray(value) && value.every((entry) => isPendingWorkbookMutation(entry));
}

export function buildZeroWorkbookMutation(
  documentId: string,
  mutation: PendingWorkbookMutationInput | PendingWorkbookMutation,
): Parameters<Zero["mutate"]>[0] {
  const { method, args } = mutation;
  const clientMutationId = "id" in mutation ? mutation.id : undefined;
  switch (method) {
    case "setCellValue": {
      const [sheetName, address, value] = args;
      if (typeof sheetName !== "string" || typeof address !== "string" || !isLiteralInput(value)) {
        throw new Error("Invalid setCellValue args");
      }
      return mutators.workbook.setCellValue({
        documentId,
        clientMutationId,
        sheetName,
        address,
        value,
      });
    }
    case "setCellFormula": {
      const [sheetName, address, formula] = args;
      if (
        typeof sheetName !== "string" ||
        typeof address !== "string" ||
        typeof formula !== "string"
      ) {
        throw new Error("Invalid setCellFormula args");
      }
      return mutators.workbook.setCellFormula({
        documentId,
        clientMutationId,
        sheetName,
        address,
        formula,
      });
    }
    case "clearCell": {
      const [sheetName, address] = args;
      if (typeof sheetName !== "string" || typeof address !== "string") {
        throw new Error("Invalid clearCell args");
      }
      return mutators.workbook.clearCell({ documentId, clientMutationId, sheetName, address });
    }
    case "clearRange": {
      const [range] = args;
      if (!isCellRangeRef(range)) {
        throw new Error("Invalid clearRange args");
      }
      return mutators.workbook.clearRange({ documentId, clientMutationId, range });
    }
    case "renderCommit": {
      const [ops] = args;
      if (!isCommitOps(ops)) {
        throw new Error("Invalid renderCommit args");
      }
      return mutators.workbook.renderCommit(
        createRenderCommitArgs({ documentId, clientMutationId, ops }),
      );
    }
    case "fillRange":
    case "copyRange":
    case "moveRange": {
      const [source, target] = args;
      if (!isCellRangeRef(source) || !isCellRangeRef(target)) {
        throw new Error(`Invalid ${method} args`);
      }
      return mutators.workbook[method]({ documentId, clientMutationId, source, target });
    }
    case "insertRows":
    case "deleteRows":
    case "insertColumns":
    case "deleteColumns": {
      const [sheetName, start, count] = args;
      if (typeof sheetName !== "string" || typeof start !== "number" || typeof count !== "number") {
        throw new Error(`Invalid ${method} args`);
      }
      if (method === "insertRows") {
        return mutators.workbook.insertRows({
          documentId,
          clientMutationId,
          sheetName,
          start,
          count,
        });
      }
      if (method === "deleteRows") {
        return mutators.workbook.deleteRows({
          documentId,
          clientMutationId,
          sheetName,
          start,
          count,
        });
      }
      if (method === "insertColumns") {
        return mutators.workbook.insertColumns({
          documentId,
          clientMutationId,
          sheetName,
          start,
          count,
        });
      }
      return mutators.workbook.deleteColumns({
        documentId,
        clientMutationId,
        sheetName,
        start,
        count,
      });
    }
    case "updateRowMetadata": {
      const [sheetName, startRow, count, height, hidden] = args;
      if (
        typeof sheetName !== "string" ||
        typeof startRow !== "number" ||
        typeof count !== "number" ||
        (height !== null && typeof height !== "number") ||
        (hidden !== null && typeof hidden !== "boolean")
      ) {
        throw new Error("Invalid updateRowMetadata args");
      }
      return mutators.workbook.updateRowMetadata({
        documentId,
        clientMutationId,
        sheetName,
        startRow,
        count,
        height,
        hidden,
      });
    }
    case "updateColumnMetadata": {
      const [sheetName, startCol, count, width, hidden] = args;
      if (
        typeof sheetName !== "string" ||
        typeof startCol !== "number" ||
        typeof count !== "number" ||
        (width !== null && typeof width !== "number") ||
        (hidden !== null && typeof hidden !== "boolean")
      ) {
        throw new Error("Invalid updateColumnMetadata args");
      }
      return mutators.workbook.updateColumnMetadata({
        documentId,
        clientMutationId,
        sheetName,
        startCol,
        count,
        width,
        hidden,
      });
    }
    case "updateColumnWidth": {
      const [sheetName, columnIndex, width] = args;
      if (
        typeof sheetName !== "string" ||
        typeof columnIndex !== "number" ||
        typeof width !== "number"
      ) {
        throw new Error("Invalid updateColumnWidth args");
      }
      return mutators.workbook.updateColumnMetadata({
        documentId,
        clientMutationId,
        sheetName,
        startCol: columnIndex,
        count: 1,
        width,
        hidden: null,
      });
    }
    case "setFreezePane": {
      const [sheetName, rows, cols] = args;
      if (typeof sheetName !== "string" || typeof rows !== "number" || typeof cols !== "number") {
        throw new Error("Invalid setFreezePane args");
      }
      return mutators.workbook.setFreezePane({
        documentId,
        clientMutationId,
        sheetName,
        rows,
        cols,
      });
    }
    case "setRangeStyle": {
      const [range, patch] = args;
      if (!isCellRangeRef(range) || !isCellStylePatchValue(patch)) {
        throw new Error("Invalid setRangeStyle args");
      }
      return mutators.workbook.setRangeStyle({ documentId, clientMutationId, range, patch });
    }
    case "clearRangeStyle": {
      const [range, fields] = args;
      if (!isCellRangeRef(range) || (fields !== undefined && !isCellStyleFieldList(fields))) {
        throw new Error("Invalid clearRangeStyle args");
      }
      return mutators.workbook.clearRangeStyle({ documentId, clientMutationId, range, fields });
    }
    case "setRangeNumberFormat": {
      const [range, format] = args;
      if (!isCellRangeRef(range) || !isCellNumberFormatInputValue(format)) {
        throw new Error("Invalid setRangeNumberFormat args");
      }
      return mutators.workbook.setRangeNumberFormat({
        documentId,
        clientMutationId,
        range,
        format,
      });
    }
    case "clearRangeNumberFormat": {
      const [range] = args;
      if (!isCellRangeRef(range)) {
        throw new Error("Invalid clearRangeNumberFormat args");
      }
      return mutators.workbook.clearRangeNumberFormat({ documentId, clientMutationId, range });
    }
  }
}
