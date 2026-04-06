import type { CommitOp } from "@bilig/core";
import type { Zero } from "@rocicorp/zero";
import { mutators } from "@bilig/zero-sync";
import type {
  CellNumberFormatInput,
  CellRangeRef,
  CellStyleField,
  CellStylePatch,
  LiteralInput,
} from "@bilig/protocol";

export type WorkbookMutationMethod =
  | "setCellValue"
  | "setCellFormula"
  | "clearCell"
  | "clearRange"
  | "renderCommit"
  | "fillRange"
  | "copyRange"
  | "moveRange"
  | "updateColumnWidth"
  | "setRangeStyle"
  | "clearRangeStyle"
  | "setRangeNumberFormat"
  | "clearRangeNumberFormat";

export interface PendingWorkbookMutationInput {
  readonly method: WorkbookMutationMethod;
  readonly args: unknown[];
}

export interface PendingWorkbookMutation extends PendingWorkbookMutationInput {
  readonly id: string;
  readonly localSeq: number;
  readonly baseRevision: number;
  readonly enqueuedAtUnixMs: number;
  readonly status: "pending";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

export function isLiteralInput(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  );
}

export function isCellRangeRef(value: unknown): value is CellRangeRef {
  return (
    isRecord(value) &&
    typeof value["sheetName"] === "string" &&
    typeof value["startAddress"] === "string" &&
    typeof value["endAddress"] === "string"
  );
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

export function isCommitOps(value: unknown): value is CommitOp[] {
  return Array.isArray(value);
}

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
    value === "updateColumnWidth" ||
    value === "setRangeStyle" ||
    value === "clearRangeStyle" ||
    value === "setRangeNumberFormat" ||
    value === "clearRangeNumberFormat"
  );
}

export function isPendingWorkbookMutationInput(
  value: unknown,
): value is PendingWorkbookMutationInput {
  return (
    isRecord(value) && isWorkbookMutationMethod(value["method"]) && Array.isArray(value["args"])
  );
}

export function isPendingWorkbookMutation(value: unknown): value is PendingWorkbookMutation {
  return (
    isRecord(value) &&
    typeof value["id"] === "string" &&
    typeof value["localSeq"] === "number" &&
    typeof value["baseRevision"] === "number" &&
    typeof value["enqueuedAtUnixMs"] === "number" &&
    value["status"] === "pending" &&
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
  mutation: PendingWorkbookMutationInput,
): Parameters<Zero["mutate"]>[0] {
  const { method, args } = mutation;
  switch (method) {
    case "setCellValue": {
      const [sheetName, address, value] = args;
      if (typeof sheetName !== "string" || typeof address !== "string" || !isLiteralInput(value)) {
        throw new Error("Invalid setCellValue args");
      }
      return mutators.workbook.setCellValue({ documentId, sheetName, address, value });
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
      return mutators.workbook.setCellFormula({ documentId, sheetName, address, formula });
    }
    case "clearCell": {
      const [sheetName, address] = args;
      if (typeof sheetName !== "string" || typeof address !== "string") {
        throw new Error("Invalid clearCell args");
      }
      return mutators.workbook.clearCell({ documentId, sheetName, address });
    }
    case "clearRange": {
      const [range] = args;
      if (!isCellRangeRef(range)) {
        throw new Error("Invalid clearRange args");
      }
      return mutators.workbook.clearRange({ documentId, range });
    }
    case "renderCommit": {
      const [ops] = args;
      if (!isCommitOps(ops)) {
        throw new Error("Invalid renderCommit args");
      }
      return mutators.workbook.renderCommit({ documentId, ops });
    }
    case "fillRange":
    case "copyRange":
    case "moveRange": {
      const [source, target] = args;
      if (!isCellRangeRef(source) || !isCellRangeRef(target)) {
        throw new Error(`Invalid ${method} args`);
      }
      return mutators.workbook[method]({ documentId, source, target });
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
      return mutators.workbook.updateColumnWidth({
        documentId,
        sheetName,
        columnIndex,
        width,
      });
    }
    case "setRangeStyle": {
      const [range, patch] = args;
      if (!isCellRangeRef(range) || !isCellStylePatchValue(patch)) {
        throw new Error("Invalid setRangeStyle args");
      }
      return mutators.workbook.setRangeStyle({ documentId, range, patch });
    }
    case "clearRangeStyle": {
      const [range, fields] = args;
      if (!isCellRangeRef(range) || (fields !== undefined && !isCellStyleFieldList(fields))) {
        throw new Error("Invalid clearRangeStyle args");
      }
      return mutators.workbook.clearRangeStyle({ documentId, range, fields });
    }
    case "setRangeNumberFormat": {
      const [range, format] = args;
      if (!isCellRangeRef(range) || !isCellNumberFormatInputValue(format)) {
        throw new Error("Invalid setRangeNumberFormat args");
      }
      return mutators.workbook.setRangeNumberFormat({ documentId, range, format });
    }
    case "clearRangeNumberFormat": {
      const [range] = args;
      if (!isCellRangeRef(range)) {
        throw new Error("Invalid clearRangeNumberFormat args");
      }
      return mutators.workbook.clearRangeNumberFormat({ documentId, range });
    }
  }
}
