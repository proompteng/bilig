import { ValueTag, type CellSnapshot } from "@bilig/protocol";
import type { EngineOp } from "@bilig/workbook-domain";
import { makeCellKey, type WorkbookStore } from "../../workbook-store.js";
import type { PreparedCellAddress, TransactionRecord } from "../runtime-state.js";

type FastHistoryOp = Extract<
  EngineOp,
  | { kind: "upsertWorkbook" }
  | { kind: "upsertSheet" }
  | { kind: "renameSheet" }
  | { kind: "setCellValue" }
  | { kind: "setCellFormula" }
  | { kind: "clearCell" }
  | { kind: "setCellFormat" }
>;

type FastHistoryCloneOp = FastHistoryOp | Extract<EngineOp, { kind: "deleteSheet" }>;

export interface FastMutationHistoryResult {
  forward: TransactionRecord;
  inverse: TransactionRecord;
  undoOps: EngineOp[] | null;
}

interface FastMutationHistoryArgs {
  readonly workbook: WorkbookStore;
  readonly getCellByIndex: (cellIndex: number) => CellSnapshot;
  readonly ops: readonly EngineOp[];
  readonly potentialNewCells?: number;
  readonly includeUndoOps?: boolean;
  readonly cloneForwardOps?: boolean;
  readonly preparedCellAddressesByOpIndex?: readonly (PreparedCellAddress | null)[];
}

function isFastHistoryOp(op: EngineOp): op is FastHistoryOp {
  return (
    op.kind === "upsertWorkbook" ||
    op.kind === "upsertSheet" ||
    op.kind === "renameSheet" ||
    op.kind === "setCellValue" ||
    op.kind === "setCellFormula" ||
    op.kind === "clearCell" ||
    op.kind === "setCellFormat"
  );
}

function cloneFastHistoryForwardOp(op: FastHistoryOp): FastHistoryOp {
  switch (op.kind) {
    case "upsertWorkbook":
      return {
        kind: "upsertWorkbook",
        name: op.name,
      };
    case "upsertSheet":
      return {
        kind: "upsertSheet",
        name: op.name,
        order: op.order,
      };
    case "renameSheet":
      return {
        kind: "renameSheet",
        oldName: op.oldName,
        newName: op.newName,
      };
    case "setCellValue":
      return {
        kind: "setCellValue",
        sheetName: op.sheetName,
        address: op.address,
        value: op.value,
      };
    case "setCellFormula":
      return {
        kind: "setCellFormula",
        sheetName: op.sheetName,
        address: op.address,
        formula: op.formula,
      };
    case "clearCell":
      return {
        kind: "clearCell",
        sheetName: op.sheetName,
        address: op.address,
      };
    case "setCellFormat":
      return {
        kind: "setCellFormat",
        sheetName: op.sheetName,
        address: op.address,
        format: op.format,
      };
  }
}

function cloneFastHistoryUndoOp(op: FastHistoryCloneOp): FastHistoryCloneOp {
  if (op.kind === "deleteSheet") {
    return {
      kind: "deleteSheet",
      name: op.name,
    };
  }
  return cloneFastHistoryForwardOp(op);
}

function restoreCellOpFromSnapshot(
  workbook: WorkbookStore,
  getCellByIndex: (cellIndex: number) => CellSnapshot,
  sheetName: string,
  address: string,
  preparedCellAddress: PreparedCellAddress | null,
): Extract<FastHistoryOp, { kind: "setCellValue" | "setCellFormula" | "clearCell" }> {
  const sheet = workbook.getSheet(sheetName);
  const cellIndex =
    preparedCellAddress && sheet
      ? workbook.cellKeyToIndex.get(
          makeCellKey(sheet.id, preparedCellAddress.row, preparedCellAddress.col),
        )
      : workbook.getCellIndex(sheetName, address);
  if (cellIndex === undefined) {
    return { kind: "clearCell", sheetName, address };
  }

  const snapshot = getCellByIndex(cellIndex);
  if (snapshot.formula !== undefined) {
    return {
      kind: "setCellFormula",
      sheetName,
      address,
      formula: snapshot.formula,
    };
  }

  switch (snapshot.value.tag) {
    case ValueTag.Empty:
    case ValueTag.Error:
      return { kind: "clearCell", sheetName, address };
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return {
        kind: "setCellValue",
        sheetName,
        address,
        value: snapshot.value.value,
      };
  }
}

function buildFastInverseOp(
  workbook: WorkbookStore,
  getCellByIndex: (cellIndex: number) => CellSnapshot,
  op: FastHistoryOp,
  preparedCellAddress: PreparedCellAddress | null,
): FastHistoryCloneOp | null {
  switch (op.kind) {
    case "upsertWorkbook":
      return {
        kind: "upsertWorkbook",
        name: workbook.workbookName,
      };
    case "upsertSheet": {
      const existing = workbook.getSheet(op.name);
      if (!existing) {
        return { kind: "deleteSheet", name: op.name };
      }
      return {
        kind: "upsertSheet",
        name: existing.name,
        order: existing.order,
      };
    }
    case "renameSheet": {
      const existing = workbook.getSheet(op.newName);
      if (!existing) {
        return null;
      }
      return {
        kind: "renameSheet",
        oldName: op.newName,
        newName: op.oldName,
      };
    }
    case "setCellValue":
    case "setCellFormula":
    case "clearCell":
      return restoreCellOpFromSnapshot(
        workbook,
        getCellByIndex,
        op.sheetName,
        op.address,
        preparedCellAddress,
      );
    case "setCellFormat": {
      const sheet = workbook.getSheet(op.sheetName);
      const cellIndex =
        preparedCellAddress && sheet
          ? workbook.cellKeyToIndex.get(
              makeCellKey(sheet.id, preparedCellAddress.row, preparedCellAddress.col),
            )
          : workbook.getCellIndex(op.sheetName, op.address);
      return {
        kind: "setCellFormat",
        sheetName: op.sheetName,
        address: op.address,
        format: cellIndex === undefined ? null : (workbook.getCellFormat(cellIndex) ?? null),
      };
    }
    default:
      return null;
  }
}

function collectCreatedSheetNames(
  workbook: WorkbookStore,
  ops: readonly FastHistoryOp[],
): ReadonlySet<string> {
  const knownSheetNames = new Set(workbook.sheetsByName.keys());
  const liveCreatedSheetNames = new Set<string>();
  const createdSheetNames = new Set<string>();

  for (let index = 0; index < ops.length; index += 1) {
    const op = ops[index];
    if (!op) {
      continue;
    }
    if (op.kind === "upsertSheet") {
      if (!knownSheetNames.has(op.name)) {
        liveCreatedSheetNames.add(op.name);
        createdSheetNames.add(op.name);
      }
      knownSheetNames.add(op.name);
      continue;
    }
    if (op.kind !== "renameSheet") {
      continue;
    }
    if (liveCreatedSheetNames.delete(op.oldName)) {
      liveCreatedSheetNames.add(op.newName);
      createdSheetNames.add(op.newName);
    }
    if (knownSheetNames.delete(op.oldName)) {
      knownSheetNames.add(op.newName);
    }
  }

  return createdSheetNames;
}

export function tryBuildFastMutationHistory(
  args: FastMutationHistoryArgs,
): FastMutationHistoryResult | null {
  const forwardOps = Array<FastHistoryOp>(args.ops.length);
  if (
    args.preparedCellAddressesByOpIndex &&
    args.preparedCellAddressesByOpIndex.length !== args.ops.length
  ) {
    throw new Error("Prepared cell addresses must align with fast-history operations");
  }
  for (let index = 0; index < args.ops.length; index += 1) {
    const op = args.ops[index];
    if (op === undefined || !isFastHistoryOp(op)) {
      return null;
    }
    forwardOps[index] = args.cloneForwardOps === false ? op : cloneFastHistoryForwardOp(op);
  }

  const createdSheetNames = collectCreatedSheetNames(args.workbook, forwardOps);

  const inverseOps: EngineOp[] = [];
  for (let index = forwardOps.length - 1; index >= 0; index -= 1) {
    const op = forwardOps[index];
    if (op === undefined) {
      return null;
    }
    if (
      (op.kind === "setCellValue" ||
        op.kind === "setCellFormula" ||
        op.kind === "clearCell" ||
        op.kind === "setCellFormat") &&
      createdSheetNames.has(op.sheetName)
    ) {
      continue;
    }
    const inverse = buildFastInverseOp(
      args.workbook,
      args.getCellByIndex,
      op,
      args.preparedCellAddressesByOpIndex?.[index] ?? null,
    );
    if (!inverse) {
      return null;
    }
    inverseOps.push(inverse);
  }

  return {
    forward:
      args.potentialNewCells === undefined
        ? { ops: forwardOps }
        : { ops: forwardOps, potentialNewCells: args.potentialNewCells },
    inverse: {
      ops: inverseOps,
      potentialNewCells: args.ops.length,
    },
    undoOps:
      args.includeUndoOps === false
        ? null
        : inverseOps.map((op) => {
            if (op.kind === "deleteSheet") {
              return cloneFastHistoryUndoOp(op);
            }
            if (!isFastHistoryOp(op)) {
              throw new TypeError(`Unsupported fast-path undo op: ${op.kind}`);
            }
            return cloneFastHistoryUndoOp(op);
          }),
  };
}
