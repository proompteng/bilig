import { SpreadsheetEngine } from "@bilig/core";
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
  applyBatchArgsSchema,
  clearRangeArgsSchema,
  clearRangeNumberFormatArgsSchema,
  clearRangeStyleArgsSchema,
  clearCellArgsSchema,
  rangeMutationArgsSchema,
  renderCommitArgsSchema,
  setCellFormulaArgsSchema,
  setCellValueArgsSchema,
  setRangeNumberFormatArgsSchema,
  setRangeStyleArgsSchema,
  updatePresenceArgsSchema,
  updateColumnWidthArgsSchema,
  type WorkbookEventPayload,
} from "@bilig/zero-sync";
import { z } from "zod";
import type { SessionIdentity } from "../http/session.js";
import { WorkbookRuntimeManager } from "../workbook-runtime/runtime-manager.js";
import { acquireWorkbookMutationLock, persistWorkbookMutation, type Queryable } from "./store.js";
import { upsertWorkbookPresence } from "./presence-store.js";

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

async function commitWorkbookMutation(
  documentId: string,
  tx: ServerTransactionLike,
  eventPayload: WorkbookEventPayload,
  runtimeManager: WorkbookRuntimeManager,
  mutate: (engine: SpreadsheetEngine) => void,
  clientMutationId?: string,
  session?: SessionIdentity,
  updatedBy = session?.userID ?? "system",
) {
  return await runtimeManager.runExclusive(documentId, async () => {
    const db = tx.dbTransaction.wrappedTransaction;
    await acquireWorkbookMutationLock(db, documentId);
    const state = await runtimeManager.loadRuntime(db, documentId);
    try {
      mutate(state.engine);
      const ownerUserId = resolveOwnerUserId(state, session);
      const result = await persistWorkbookMutation(db, documentId, {
        previousState: state,
        nextEngine: state.engine,
        updatedBy,
        ownerUserId,
        eventPayload,
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
      const parsed = applyBatchArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "applyBatch",
          batch: parsed.batch,
        },
        runtimeManager,
        (engine) => {
          engine.applyRemoteBatch(parsed.batch);
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
          engine.setCellValue(parsed.sheetName, parsed.address, parsed.value);
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
          engine.setCellFormula(parsed.sheetName, parsed.address, parsed.formula);
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
          engine.clearCell(parsed.sheetName, parsed.address);
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
          engine.clearRange(parsed.range);
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.renderCommit": {
      const parsed = renderCommitArgsSchema.parse(args);
      await commitWorkbookMutation(
        parsed.documentId,
        serverTx,
        {
          kind: "renderCommit",
          ops: parsed.ops,
        },
        runtimeManager,
        (engine) => {
          engine.renderCommit(parsed.ops);
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
          engine.fillRange(parsed.source, parsed.target);
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
          engine.copyRange(parsed.source, parsed.target);
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
          engine.moveRange(parsed.source, parsed.target);
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
          engine.updateColumnMetadata(parsed.sheetName, parsed.columnIndex, 1, parsed.width, null);
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
          engine.setRangeStyle(parsed.range, patch);
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
          engine.clearRangeStyle(parsed.range, parsed.fields);
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
          engine.setRangeNumberFormat(parsed.range, format);
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
          engine.clearRangeNumberFormat(parsed.range);
        },
        parsed.clientMutationId,
        session,
      );
      return;
    }

    case "workbook.updatePresence": {
      const parsed = updatePresenceArgsSchema.parse(args);
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

    default:
      throw new Error(`Unknown Zero mutator: ${name}`);
  }
}
