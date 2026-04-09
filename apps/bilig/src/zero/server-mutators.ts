import { SpreadsheetEngine } from "@bilig/core";
import { applyWorkbookAgentCommandBundle } from "@bilig/agent-api";
import type { EngineOp } from "@bilig/workbook-domain";
import type {
  CellBorderSidePatch,
  CellNumberFormatInput,
  CellNumberFormatPreset,
  CellStyleAlignmentPatch,
  CellStyleBordersPatch,
  CellStyleFillPatch,
  CellStyleFontPatch,
  CellStylePatch,
} from "@bilig/protocol";
import {
  applyAgentCommandBundleArgsSchema,
  clearRangeArgsSchema,
  clearRangeNumberFormatArgsSchema,
  clearRangeStyleArgsSchema,
  clearCellArgsSchema,
  parseApplyBatchArgs,
  parseRenderCommitArgs,
  redoLatestWorkbookChangeArgsSchema,
  rangeMutationArgsSchema,
  revertWorkbookChangeArgsSchema,
  setCellFormulaArgsSchema,
  setCellValueArgsSchema,
  setFreezePaneArgsSchema,
  setRangeNumberFormatArgsSchema,
  setRangeStyleArgsSchema,
  undoLatestWorkbookChangeArgsSchema,
  updateColumnMetadataArgsSchema,
  updatePresenceArgsSchema,
  updateColumnWidthArgsSchema,
  updateRowMetadataArgsSchema,
  type WorkbookChangeUndoBundle,
  type WorkbookEventPayload,
} from "@bilig/zero-sync";
import { z } from "zod";
import type { SessionIdentity } from "../http/session.js";
import { WorkbookRuntimeManager } from "../workbook-runtime/runtime-manager.js";
import type { Queryable } from "./store.js";
import { acquireWorkbookMutationLock } from "./workbook-runtime-store.js";
import { ensureWorkbookDocumentExists } from "./workbook-migration-store.js";
import { persistWorkbookMutation } from "./workbook-mutation-store.js";
import { upsertWorkbookPresence } from "./presence-store.js";
import {
  loadLatestRedoableWorkbookChange,
  loadLatestUndoableWorkbookChange,
  loadWorkbookChange,
} from "./workbook-change-store.js";

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

function resolveOwnerUserId(state: { ownerUserId: string }, session?: SessionIdentity): string {
  if (state.ownerUserId !== "system" || !session?.userID) {
    return state.ownerUserId;
  }
  return session.userID;
}

function toEngineUndoBundle(undoOps: readonly EngineOp[] | null): WorkbookChangeUndoBundle | null {
  if (!undoOps || undoOps.length === 0) {
    return null;
  }
  return {
    kind: "engineOps",
    ops: structuredClone([...undoOps]),
  };
}

function captureEngineUndoBundle(
  engine: SpreadsheetEngine,
  mutate: (engine: SpreadsheetEngine) => void,
): WorkbookChangeUndoBundle | null {
  return toEngineUndoBundle(
    engine.captureUndoOps(() => {
      mutate(engine);
    }).undoOps,
  );
}

function applyWorkbookChangeUndoBundle(
  engine: SpreadsheetEngine,
  undoBundle: WorkbookChangeUndoBundle,
): WorkbookChangeUndoBundle | null {
  switch (undoBundle.kind) {
    case "engineOps":
      return toEngineUndoBundle(engine.applyOps(undoBundle.ops, { captureUndo: true }));
    case "snapshot": {
      const redoSnapshot = engine.exportSnapshot();
      engine.importSnapshot(undoBundle.snapshot);
      return {
        kind: "snapshot",
        snapshot: redoSnapshot,
      };
    }
    default: {
      const exhaustive: never = undoBundle;
      return exhaustive;
    }
  }
}

async function commitWorkbookMutation(
  documentId: string,
  tx: ServerTransactionLike,
  eventPayload: WorkbookEventPayload,
  runtimeManager: WorkbookRuntimeManager,
  mutate: (engine: SpreadsheetEngine) => WorkbookChangeUndoBundle | null,
  clientMutationId?: string,
  session?: SessionIdentity,
  updatedBy = session?.userID ?? "system",
) {
  return await runtimeManager.runExclusive(documentId, async () => {
    const db = tx.dbTransaction.wrappedTransaction;
    await acquireWorkbookMutationLock(db, documentId);
    const state = await runtimeManager.loadRuntime(db, documentId);
    try {
      const undoBundle = mutate(state.engine);
      const ownerUserId = resolveOwnerUserId(state, session);
      const result = await persistWorkbookMutation(db, documentId, {
        previousState: state,
        nextEngine: state.engine,
        updatedBy,
        ownerUserId,
        eventPayload,
        undoBundle,
        ...(clientMutationId !== undefined ? { clientMutationId } : {}),
      });
      runtimeManager.commitMutation(documentId, {
        projectionCommit: result.projectionCommit,
        headRevision: result.revision,
        calculatedRevision: result.calculatedRevision,
        ownerUserId,
      });
      return {
        documentId,
        revision: result.revision,
        updatedAt: result.updatedAt,
      };
    } catch (error) {
      runtimeManager.invalidate(documentId);
      throw error;
    }
  });
}

async function commitWorkbookHistoryMutation(input: {
  documentId: string;
  serverTx: ServerTransactionLike;
  runtimeManager: WorkbookRuntimeManager;
  session?: SessionIdentity;
  clientMutationId?: string;
  targetChange: {
    revision: number;
    summary: string;
    sheetName: string | null;
    anchorAddress: string | null;
    range: import("./workbook-change-store.js").WorkbookChangeRange | null;
    undoBundle: WorkbookChangeUndoBundle;
  };
  eventKind: "revertChange" | "redoChange";
}): Promise<void> {
  const {
    clientMutationId,
    documentId,
    eventKind,
    runtimeManager,
    serverTx,
    session,
    targetChange,
  } = input;
  await runtimeManager.runExclusive(documentId, async () => {
    const db = serverTx.dbTransaction.wrappedTransaction;
    await acquireWorkbookMutationLock(db, documentId);
    const state = await runtimeManager.loadRuntime(db, documentId);
    const eventPayload: WorkbookEventPayload = {
      kind: eventKind,
      targetRevision: targetChange.revision,
      targetSummary: targetChange.summary,
      ...(targetChange.sheetName ? { sheetName: targetChange.sheetName } : {}),
      ...(targetChange.anchorAddress ? { address: targetChange.anchorAddress } : {}),
      ...(targetChange.range ? { range: targetChange.range } : {}),
      appliedBundle: targetChange.undoBundle,
    };
    try {
      const undoBundle = applyWorkbookChangeUndoBundle(state.engine, targetChange.undoBundle);
      const ownerUserId = resolveOwnerUserId(state, session);
      const result = await persistWorkbookMutation(db, documentId, {
        previousState: state,
        nextEngine: state.engine,
        updatedBy: session?.userID ?? "system",
        ownerUserId,
        eventPayload,
        undoBundle,
        ...(clientMutationId !== undefined ? { clientMutationId } : {}),
      });
      runtimeManager.commitMutation(documentId, {
        projectionCommit: result.projectionCommit,
        headRevision: result.revision,
        calculatedRevision: result.calculatedRevision,
        ownerUserId,
      });
    } catch (error) {
      runtimeManager.invalidate(documentId);
      throw error;
    }
  });
}

export async function handleServerMutator(
  tx: unknown,
  name: string,
  args: unknown,
  runtimeManager: WorkbookRuntimeManager,
  session?: SessionIdentity,
): Promise<void> {
  const serverTx = requireServerTransaction(tx);

  switch (name) {
    case "workbook.applyBatch": {
      const parsed = parseApplyBatchArgs(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "applyBatch",
          batch: parsed.batch,
        },
        runtimeManager,
        (engine) => {
          return toEngineUndoBundle(engine.applyOps(parsed.batch.ops, { captureUndo: true }));
        },
        parsed.clientMutationId,
        session,
        session?.userID ??
          (isRecord(parsed.batch) && typeof parsed.batch["replicaId"] === "string"
            ? parsed.batch["replicaId"]
            : "system"),
      );
      return;
    }

    case "workbook.applyAgentCommandBundle": {
      const parsed = applyAgentCommandBundleArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "applyAgentCommandBundle",
          bundle: parsed.bundle,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            applyWorkbookAgentCommandBundle(draft, parsed.bundle);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.setCellValue": {
      const parsed = setCellValueArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "setCellValue",
          sheetName: parsed.sheetName,
          address: parsed.address,
          value: parsed.value,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setCellValue(parsed.sheetName, parsed.address, parsed.value);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.setCellFormula": {
      const parsed = setCellFormulaArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "setCellFormula",
          sheetName: parsed.sheetName,
          address: parsed.address,
          formula: parsed.formula,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setCellFormula(parsed.sheetName, parsed.address, parsed.formula);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.clearCell": {
      const parsed = clearCellArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "clearCell",
          sheetName: parsed.sheetName,
          address: parsed.address,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.clearCell(parsed.sheetName, parsed.address);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.clearRange": {
      const parsed = clearRangeArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "clearRange",
          range: parsed.range,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.clearRange(parsed.range);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.renderCommit": {
      const parsed = parseRenderCommitArgs(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "renderCommit",
          ops: parsed.ops,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.renderCommit(parsed.ops);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.fillRange": {
      const parsed = rangeMutationArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "fillRange",
          source: parsed.source,
          target: parsed.target,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.fillRange(parsed.source, parsed.target);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.copyRange": {
      const parsed = rangeMutationArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "copyRange",
          source: parsed.source,
          target: parsed.target,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.copyRange(parsed.source, parsed.target);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.moveRange": {
      const parsed = rangeMutationArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "moveRange",
          source: parsed.source,
          target: parsed.target,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.moveRange(parsed.source, parsed.target);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.updateRowMetadata": {
      const parsed = updateRowMetadataArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "updateRowMetadata",
          sheetName: parsed.sheetName,
          startRow: parsed.startRow,
          count: parsed.count,
          height: parsed.height,
          hidden: parsed.hidden,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.updateRowMetadata(
              parsed.sheetName,
              parsed.startRow,
              parsed.count,
              parsed.height,
              parsed.hidden,
            );
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.updateColumnMetadata": {
      const parsed = updateColumnMetadataArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "updateColumnMetadata",
          sheetName: parsed.sheetName,
          startCol: parsed.startCol,
          count: parsed.count,
          width: parsed.width,
          hidden: parsed.hidden,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.updateColumnMetadata(
              parsed.sheetName,
              parsed.startCol,
              parsed.count,
              parsed.width,
              parsed.hidden,
            );
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.updateColumnWidth": {
      const parsed = updateColumnWidthArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "updateColumnWidth",
          sheetName: parsed.sheetName,
          columnIndex: parsed.columnIndex,
          width: parsed.width,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.updateColumnMetadata(parsed.sheetName, parsed.columnIndex, 1, parsed.width, null);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.setFreezePane": {
      const parsed = setFreezePaneArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "setFreezePane",
          sheetName: parsed.sheetName,
          rows: parsed.rows,
          cols: parsed.cols,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setFreezePane(parsed.sheetName, parsed.rows, parsed.cols);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.setRangeStyle": {
      const parsed = setRangeStyleArgsSchema.parse(args);
      const patch = normalizeStylePatch(parsed.patch);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "setRangeStyle",
          range: parsed.range,
          patch,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setRangeStyle(parsed.range, patch);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.clearRangeStyle": {
      const parsed = clearRangeStyleArgsSchema.parse(args);
      const eventPayload =
        parsed.fields === undefined
          ? {
              kind: "clearRangeStyle" as const,
              range: parsed.range,
            }
          : {
              kind: "clearRangeStyle" as const,
              range: parsed.range,
              fields: parsed.fields,
            };
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        eventPayload,
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.clearRangeStyle(parsed.range, parsed.fields);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.setRangeNumberFormat": {
      const parsed = setRangeNumberFormatArgsSchema.parse(args);
      const format = normalizeNumberFormatInput(parsed.format);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "setRangeNumberFormat",
          range: parsed.range,
          format,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.setRangeNumberFormat(parsed.range, format);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.clearRangeNumberFormat": {
      const parsed = clearRangeNumberFormatArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "clearRangeNumberFormat",
          range: parsed.range,
        },
        runtimeManager,
        (engine) => {
          return captureEngineUndoBundle(engine, (draft) => {
            draft.clearRangeNumberFormat(parsed.range);
          });
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.updatePresence": {
      const parsed = updatePresenceArgsSchema.parse(args);
      await ensureWorkbookDocumentExists(
        serverTx.dbTransaction.wrappedTransaction,
        parsed.documentId,
        session?.userID ?? "system",
      );
      await upsertWorkbookPresence(serverTx.dbTransaction.wrappedTransaction, {
        documentId: parsed.documentId,
        sessionId: parsed.sessionId,
        userId: session?.userID ?? "system",
        sheetId: parsed.sheetId ?? null,
        sheetName: parsed.sheetName ?? null,
        address: parsed.address ?? null,
        selection: parsed.selection,
      });
      return;
    }

    case "workbook.revertChange": {
      const parsed = revertWorkbookChangeArgsSchema.parse(args);
      const db = serverTx.dbTransaction.wrappedTransaction;
      const targetChange = await loadWorkbookChange(db, parsed.documentId, parsed.revision);
      if (!targetChange) {
        throw new Error("Workbook change was not found");
      }
      if (!targetChange.undoBundle) {
        throw new Error("Workbook change is not revertible");
      }
      if (targetChange.revertedByRevision !== null) {
        throw new Error(
          `Workbook change was already reverted in r${targetChange.revertedByRevision}`,
        );
      }
      if (targetChange.eventKind === "revertChange" || targetChange.revertsRevision !== null) {
        throw new Error("Reverting a revert change is not supported");
      }
      await commitWorkbookHistoryMutation({
        documentId: parsed.documentId,
        serverTx,
        runtimeManager,
        session,
        ...(parsed.clientMutationId !== undefined
          ? { clientMutationId: parsed.clientMutationId }
          : {}),
        targetChange: {
          revision: targetChange.revision,
          summary: targetChange.summary,
          sheetName: targetChange.sheetName,
          anchorAddress: targetChange.anchorAddress,
          range: targetChange.range,
          undoBundle: targetChange.undoBundle,
        },
        eventKind: "revertChange",
      });
      return;
    }

    case "workbook.undoLatestChange": {
      const parsed = undoLatestWorkbookChangeArgsSchema.parse(args);
      const targetChange = await loadLatestUndoableWorkbookChange(
        serverTx.dbTransaction.wrappedTransaction,
        {
          documentId: parsed.documentId,
          actorUserId: session?.userID ?? "system",
        },
      );
      if (!targetChange?.undoBundle) {
        throw new Error("No undoable workbook change was found");
      }
      await commitWorkbookHistoryMutation({
        documentId: parsed.documentId,
        serverTx,
        runtimeManager,
        session,
        ...(parsed.clientMutationId !== undefined
          ? { clientMutationId: parsed.clientMutationId }
          : {}),
        targetChange: {
          revision: targetChange.revision,
          summary: targetChange.summary,
          sheetName: targetChange.sheetName,
          anchorAddress: targetChange.anchorAddress,
          range: targetChange.range,
          undoBundle: targetChange.undoBundle,
        },
        eventKind: "revertChange",
      });
      return;
    }

    case "workbook.redoLatestChange": {
      const parsed = redoLatestWorkbookChangeArgsSchema.parse(args);
      const targetChange = await loadLatestRedoableWorkbookChange(
        serverTx.dbTransaction.wrappedTransaction,
        {
          documentId: parsed.documentId,
          actorUserId: session?.userID ?? "system",
        },
      );
      if (!targetChange?.undoBundle) {
        throw new Error("No redoable workbook change was found");
      }
      await commitWorkbookHistoryMutation({
        documentId: parsed.documentId,
        serverTx,
        runtimeManager,
        session,
        ...(parsed.clientMutationId !== undefined
          ? { clientMutationId: parsed.clientMutationId }
          : {}),
        targetChange: {
          revision: targetChange.revision,
          summary: targetChange.summary,
          sheetName: targetChange.sheetName,
          anchorAddress: targetChange.anchorAddress,
          range: targetChange.range,
          undoBundle: targetChange.undoBundle,
        },
        eventKind: "redoChange",
      });
      return;
    }

    default:
      throw new Error(`Unknown Zero mutator: ${name}`);
  }
}
