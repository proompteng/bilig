import { SpreadsheetEngine } from "@bilig/core";
import { formatAddress } from "@bilig/formula";
import type { WorkbookSnapshot } from "@bilig/protocol";
import {
  clearCellArgsSchema,
  rangeMutationArgsSchema,
  replaceSnapshotArgsSchema,
  renderCommitArgsSchema,
  setCellFormulaArgsSchema,
  setCellValueArgsSchema,
  updateColumnWidthArgsSchema,
} from "@bilig/zero-sync";
import {
  loadWorkbookSnapshot,
  persistWorkbookProjection,
  type MaterializedComputedCell,
  type Queryable,
} from "./store.js";

interface ServerTransactionLike {
  dbTransaction: {
    wrappedTransaction: Queryable;
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isQueryable(value: unknown): value is Queryable {
  return isRecord(value) && typeof value["query"] === "function";
}

function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  return (
    isRecord(value) &&
    value["version"] === 1 &&
    isRecord(value["workbook"]) &&
    typeof value["workbook"]["name"] === "string" &&
    Array.isArray(value["sheets"])
  );
}

function requireServerTransaction(tx: unknown): ServerTransactionLike {
  if (
    !isRecord(tx) ||
    !isRecord(tx["dbTransaction"]) ||
    !isQueryable(tx["dbTransaction"]["wrappedTransaction"])
  ) {
    throw new Error("Expected a server-side Zero transaction");
  }

  return {
    dbTransaction: {
      wrappedTransaction: tx["dbTransaction"]["wrappedTransaction"],
    },
  };
}

async function createWorkbookEngine(documentId: string, snapshot: WorkbookSnapshot) {
  const engine = new SpreadsheetEngine({
    workbookName: documentId,
    replicaId: `server:${documentId}`,
  });
  await engine.ready();
  engine.importSnapshot(snapshot);
  return engine;
}

function materializeComputedCells(engine: SpreadsheetEngine): MaterializedComputedCell[] {
  const entries: MaterializedComputedCell[] = [];

  for (const sheet of engine.workbook.sheetsByName.values()) {
    sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
      const address = formatAddress(row, col);
      const cell = engine.getCell(sheet.name, address);
      entries.push({
        sheetName: sheet.name,
        address,
        value: cell.value,
        flags: cell.flags,
        version: cell.version,
      });
    });
  }

  return entries;
}

async function applyWorkbookMutation(
  documentId: string,
  tx: ServerTransactionLike,
  apply: (engine: SpreadsheetEngine) => void,
) {
  const db = tx.dbTransaction.wrappedTransaction;
  const snapshot = await loadWorkbookSnapshot(db, documentId);
  const engine = await createWorkbookEngine(documentId, snapshot);
  apply(engine);
  const nextSnapshot = engine.exportSnapshot();
  const computedCells = materializeComputedCells(engine);
  await persistWorkbookProjection(db, documentId, nextSnapshot, computedCells);
  return {
    documentId,
    updatedAt: new Date().toISOString(),
  };
}

async function replaceWorkbookSnapshot(
  documentId: string,
  tx: ServerTransactionLike,
  snapshot: WorkbookSnapshot,
) {
  const db = tx.dbTransaction.wrappedTransaction;
  const engine = await createWorkbookEngine(documentId, snapshot);
  const nextSnapshot = engine.exportSnapshot();
  const computedCells = materializeComputedCells(engine);
  await persistWorkbookProjection(db, documentId, nextSnapshot, computedCells);
  return {
    documentId,
    updatedAt: new Date().toISOString(),
  };
}

export async function handleServerMutator(tx: unknown, name: string, args: unknown): Promise<void> {
  const serverTx = requireServerTransaction(tx);

  switch (name) {
    case "workbook.setCellValue": {
      const parsed = setCellValueArgsSchema.parse(args);
      await applyWorkbookMutation(parsed.documentId, serverTx, (engine) => {
        engine.setCellValue(parsed.sheetName, parsed.address, parsed.value);
      });
      return;
    }

    case "workbook.setCellFormula": {
      const parsed = setCellFormulaArgsSchema.parse(args);
      await applyWorkbookMutation(parsed.documentId, serverTx, (engine) => {
        engine.setCellFormula(parsed.sheetName, parsed.address, parsed.formula);
      });
      return;
    }

    case "workbook.clearCell": {
      const parsed = clearCellArgsSchema.parse(args);
      await applyWorkbookMutation(parsed.documentId, serverTx, (engine) => {
        engine.clearCell(parsed.sheetName, parsed.address);
      });
      return;
    }

    case "workbook.renderCommit": {
      const parsed = renderCommitArgsSchema.parse(args);
      await applyWorkbookMutation(parsed.documentId, serverTx, (engine) => {
        engine.renderCommit(parsed.ops);
      });
      return;
    }

    case "workbook.fillRange": {
      const parsed = rangeMutationArgsSchema.parse(args);
      await applyWorkbookMutation(parsed.documentId, serverTx, (engine) => {
        engine.fillRange(parsed.source, parsed.target);
      });
      return;
    }

    case "workbook.copyRange": {
      const parsed = rangeMutationArgsSchema.parse(args);
      await applyWorkbookMutation(parsed.documentId, serverTx, (engine) => {
        engine.copyRange(parsed.source, parsed.target);
      });
      return;
    }

    case "workbook.updateColumnWidth": {
      const parsed = updateColumnWidthArgsSchema.parse(args);
      await applyWorkbookMutation(parsed.documentId, serverTx, (engine) => {
        engine.updateColumnMetadata(parsed.sheetName, parsed.columnIndex, 1, parsed.width, null);
      });
      return;
    }

    case "workbook.replaceSnapshot": {
      const parsed = replaceSnapshotArgsSchema.parse(args);
      if (!isWorkbookSnapshot(parsed.snapshot)) {
        throw new Error("Invalid workbook snapshot payload");
      }
      await replaceWorkbookSnapshot(parsed.documentId, serverTx, parsed.snapshot);
      return;
    }

    default:
      throw new Error(`Unknown Zero mutator: ${name}`);
  }
}
