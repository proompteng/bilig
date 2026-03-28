import { SpreadsheetEngine } from "@bilig/core";
import { formatAddress } from "@bilig/formula";
import type {
  CellBorderSidePatch,
  CellNumberFormatInput,
  CellNumberFormatPreset,
  CellStyleAlignmentPatch,
  CellStyleBordersPatch,
  CellStyleFillPatch,
  CellStyleFontPatch,
  CellStylePatch,
  WorkbookSnapshot,
} from "@bilig/protocol";
import {
  applyBatchArgsSchema,
  clearRangeNumberFormatArgsSchema,
  clearRangeStyleArgsSchema,
  clearCellArgsSchema,
  rangeMutationArgsSchema,
  replaceSnapshotArgsSchema,
  renderCommitArgsSchema,
  setCellFormulaArgsSchema,
  setCellValueArgsSchema,
  setRangeNumberFormatArgsSchema,
  setRangeStyleArgsSchema,
  updateColumnWidthArgsSchema,
} from "@bilig/zero-sync";
import { z } from "zod";
import {
  loadWorkbookState,
  persistWorkbookProjection,
  type WorkbookRuntimeState,
  type MaterializedComputedCell,
  type Queryable,
} from "./store.js";
import type { SessionIdentity } from "../session.js";

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

function normalizeBorderSidePatch(
  patch: {
    style?: CellBorderSidePatch["style"] | undefined;
    weight?: CellBorderSidePatch["weight"] | undefined;
    color?: CellBorderSidePatch["color"] | undefined;
  } | null,
): CellBorderSidePatch | null {
  if (patch === null) {
    return null;
  }
  const normalized: CellBorderSidePatch = {};
  if (patch.style !== undefined) {
    normalized.style = patch.style;
  }
  if (patch.weight !== undefined) {
    normalized.weight = patch.weight;
  }
  if (patch.color !== undefined) {
    normalized.color = patch.color;
  }
  return normalized;
}

function normalizeStylePatch(
  patch: z.infer<typeof setRangeStyleArgsSchema>["patch"],
): CellStylePatch {
  const normalized: CellStylePatch = {};

  if (patch.fill !== undefined) {
    if (patch.fill === null) {
      normalized.fill = null;
    } else {
      const fill: CellStyleFillPatch = {};
      if (patch.fill.backgroundColor !== undefined) {
        fill.backgroundColor = patch.fill.backgroundColor;
      }
      normalized.fill = fill;
    }
  }
  if (patch.font !== undefined) {
    if (patch.font === null) {
      normalized.font = null;
    } else {
      const font: CellStyleFontPatch = {};
      if (patch.font.family !== undefined) {
        font.family = patch.font.family;
      }
      if (patch.font.size !== undefined) {
        font.size = patch.font.size;
      }
      if (patch.font.bold !== undefined) {
        font.bold = patch.font.bold;
      }
      if (patch.font.italic !== undefined) {
        font.italic = patch.font.italic;
      }
      if (patch.font.underline !== undefined) {
        font.underline = patch.font.underline;
      }
      if (patch.font.color !== undefined) {
        font.color = patch.font.color;
      }
      normalized.font = font;
    }
  }
  if (patch.alignment !== undefined) {
    if (patch.alignment === null) {
      normalized.alignment = null;
    } else {
      const alignment: CellStyleAlignmentPatch = {};
      if (patch.alignment.horizontal !== undefined) {
        alignment.horizontal = patch.alignment.horizontal;
      }
      if (patch.alignment.vertical !== undefined) {
        alignment.vertical = patch.alignment.vertical;
      }
      if (patch.alignment.wrap !== undefined) {
        alignment.wrap = patch.alignment.wrap;
      }
      if (patch.alignment.indent !== undefined) {
        alignment.indent = patch.alignment.indent;
      }
      normalized.alignment = alignment;
    }
  }
  if (patch.borders !== undefined) {
    if (patch.borders === null) {
      normalized.borders = null;
    } else {
      const borders: CellStyleBordersPatch = {};
      if (patch.borders.top !== undefined) {
        borders.top = normalizeBorderSidePatch(patch.borders.top);
      }
      if (patch.borders.right !== undefined) {
        borders.right = normalizeBorderSidePatch(patch.borders.right);
      }
      if (patch.borders.bottom !== undefined) {
        borders.bottom = normalizeBorderSidePatch(patch.borders.bottom);
      }
      if (patch.borders.left !== undefined) {
        borders.left = normalizeBorderSidePatch(patch.borders.left);
      }
      normalized.borders = borders;
    }
  }

  return normalized;
}

function normalizeNumberFormatInput(
  format: z.infer<typeof setRangeNumberFormatArgsSchema>["format"],
): CellNumberFormatInput {
  if (typeof format === "string") {
    return format;
  }

  const normalized: CellNumberFormatPreset = {
    kind: format.kind,
  };
  if (format.currency !== undefined) {
    normalized.currency = format.currency;
  }
  if (format.decimals !== undefined) {
    normalized.decimals = format.decimals;
  }
  if (format.useGrouping !== undefined) {
    normalized.useGrouping = format.useGrouping;
  }
  if (format.negativeStyle !== undefined) {
    normalized.negativeStyle = format.negativeStyle;
  }
  if (format.zeroStyle !== undefined) {
    normalized.zeroStyle = format.zeroStyle;
  }
  if (format.dateStyle !== undefined) {
    normalized.dateStyle = format.dateStyle;
  }
  return normalized;
}

async function createWorkbookEngine(
  documentId: string,
  snapshot: WorkbookSnapshot,
  replicaSnapshot?: ReturnType<SpreadsheetEngine["exportReplicaSnapshot"]> | null,
) {
  const engine = new SpreadsheetEngine({
    workbookName: documentId,
    replicaId: `server:${documentId}`,
  });
  await engine.ready();
  engine.importSnapshot(snapshot);
  if (replicaSnapshot) {
    engine.importReplicaSnapshot(replicaSnapshot);
  }
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
        rowNum: row,
        colNum: col,
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
  session?: SessionIdentity,
) {
  const db = tx.dbTransaction.wrappedTransaction;
  const state = await loadWorkbookState(db, documentId);
  const engine = await createWorkbookEngine(documentId, state.snapshot, state.replicaSnapshot);
  apply(engine);
  const nextSnapshot = engine.exportSnapshot();
  const computedCells = materializeComputedCells(engine);
  await persistWorkbookProjection(db, documentId, nextSnapshot, computedCells, {
    replicaSnapshot: engine.exportReplicaSnapshot(),
    revision: state.headRevision + 1,
    ownerUserId: resolveOwnerUserId(state, session),
    updatedBy: session?.userID ?? "system",
  });
  return {
    documentId,
    updatedAt: new Date().toISOString(),
  };
}

async function replaceWorkbookSnapshot(
  documentId: string,
  tx: ServerTransactionLike,
  snapshot: WorkbookSnapshot,
  session?: SessionIdentity,
) {
  const db = tx.dbTransaction.wrappedTransaction;
  const state = await loadWorkbookState(db, documentId);
  const engine = await createWorkbookEngine(documentId, snapshot, state.replicaSnapshot);
  const nextSnapshot = engine.exportSnapshot();
  const computedCells = materializeComputedCells(engine);
  await persistWorkbookProjection(db, documentId, nextSnapshot, computedCells, {
    replicaSnapshot: engine.exportReplicaSnapshot(),
    revision: state.headRevision + 1,
    ownerUserId: resolveOwnerUserId(state, session),
    updatedBy: session?.userID ?? "system",
  });
  return {
    documentId,
    updatedAt: new Date().toISOString(),
  };
}

async function applyWorkbookBatch(
  documentId: string,
  tx: ServerTransactionLike,
  batch: unknown,
  session?: SessionIdentity,
) {
  const db = tx.dbTransaction.wrappedTransaction;
  const state = await loadWorkbookState(db, documentId);
  const engine = await createWorkbookEngine(documentId, state.snapshot, state.replicaSnapshot);
  engine.applyRemoteBatch(applyBatchArgsSchema.parse({ documentId, batch }).batch);
  const nextSnapshot = engine.exportSnapshot();
  const computedCells = materializeComputedCells(engine);
  await persistWorkbookProjection(db, documentId, nextSnapshot, computedCells, {
    replicaSnapshot: engine.exportReplicaSnapshot(),
    revision: state.headRevision + 1,
    ownerUserId: resolveOwnerUserId(state, session),
    updatedBy:
      session?.userID ??
      (isRecord(batch) && typeof batch["replicaId"] === "string" ? batch["replicaId"] : "system"),
  });
}

function resolveOwnerUserId(state: WorkbookRuntimeState, session?: SessionIdentity): string {
  if (state.ownerUserId !== "system" || !session?.userID) {
    return state.ownerUserId;
  }
  return session.userID;
}

export async function handleServerMutator(
  tx: unknown,
  name: string,
  args: unknown,
  session?: SessionIdentity,
): Promise<void> {
  const serverTx = requireServerTransaction(tx);

  switch (name) {
    case "workbook.applyBatch": {
      const parsed = applyBatchArgsSchema.parse(args);
      await applyWorkbookBatch(parsed.documentId, serverTx, parsed.batch, session);
      return;
    }

    case "workbook.setCellValue": {
      const parsed = setCellValueArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.setCellValue(parsed.sheetName, parsed.address, parsed.value);
        },
        session,
      );
      return;
    }

    case "workbook.setCellFormula": {
      const parsed = setCellFormulaArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.setCellFormula(parsed.sheetName, parsed.address, parsed.formula);
        },
        session,
      );
      return;
    }

    case "workbook.clearCell": {
      const parsed = clearCellArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.clearCell(parsed.sheetName, parsed.address);
        },
        session,
      );
      return;
    }

    case "workbook.renderCommit": {
      const parsed = renderCommitArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.renderCommit(parsed.ops);
        },
        session,
      );
      return;
    }

    case "workbook.fillRange": {
      const parsed = rangeMutationArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.fillRange(parsed.source, parsed.target);
        },
        session,
      );
      return;
    }

    case "workbook.copyRange": {
      const parsed = rangeMutationArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.copyRange(parsed.source, parsed.target);
        },
        session,
      );
      return;
    }

    case "workbook.updateColumnWidth": {
      const parsed = updateColumnWidthArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.updateColumnMetadata(parsed.sheetName, parsed.columnIndex, 1, parsed.width, null);
        },
        session,
      );
      return;
    }

    case "workbook.setRangeStyle": {
      const parsed = setRangeStyleArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.setRangeStyle(parsed.range, normalizeStylePatch(parsed.patch));
        },
        session,
      );
      return;
    }

    case "workbook.clearRangeStyle": {
      const parsed = clearRangeStyleArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.clearRangeStyle(parsed.range, parsed.fields);
        },
        session,
      );
      return;
    }

    case "workbook.setRangeNumberFormat": {
      const parsed = setRangeNumberFormatArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.setRangeNumberFormat(parsed.range, normalizeNumberFormatInput(parsed.format));
        },
        session,
      );
      return;
    }

    case "workbook.clearRangeNumberFormat": {
      const parsed = clearRangeNumberFormatArgsSchema.parse(args);
      await applyWorkbookMutation(
        parsed.documentId,
        serverTx,
        (engine) => {
          engine.clearRangeNumberFormat(parsed.range);
        },
        session,
      );
      return;
    }

    case "workbook.replaceSnapshot": {
      const parsed = replaceSnapshotArgsSchema.parse(args);
      if (!isWorkbookSnapshot(parsed.snapshot)) {
        throw new Error("Invalid workbook snapshot payload");
      }
      await replaceWorkbookSnapshot(parsed.documentId, serverTx, parsed.snapshot, session);
      return;
    }

    default:
      throw new Error(`Unknown Zero mutator: ${name}`);
  }
}
