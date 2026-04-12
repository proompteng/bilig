import { defineMutator, defineMutatorsWithType } from "@rocicorp/zero";
import { isWorkbookAgentCommandBundle, type WorkbookAgentCommandBundle } from "@bilig/agent-api";
import { isCommitOps, type CommitOp } from "@bilig/core";
import { z } from "zod";
import { isEngineOpBatch, type EngineOpBatch } from "@bilig/workbook-domain";
import type { CellRangeRef } from "@bilig/protocol";
import { isLiteralInput } from "@bilig/protocol";
import { schema } from "./schema.js";

const literalInputSchema = z.union([z.number(), z.string(), z.boolean(), z.null()]);

const cellRangeRefSchema = z.object({
  sheetName: z.string().min(1),
  startAddress: z.string().min(1),
  endAddress: z.string().min(1),
}) satisfies z.ZodType<CellRangeRef>;

const cellStylePatchSchema = z.object({
  fill: z
    .object({
      backgroundColor: z.string().nullable().optional(),
    })
    .nullable()
    .optional(),
  font: z
    .object({
      family: z.string().nullable().optional(),
      size: z.number().nullable().optional(),
      bold: z.boolean().nullable().optional(),
      italic: z.boolean().nullable().optional(),
      underline: z.boolean().nullable().optional(),
      color: z.string().nullable().optional(),
    })
    .nullable()
    .optional(),
  alignment: z
    .object({
      horizontal: z.enum(["general", "left", "center", "right"]).nullable().optional(),
      vertical: z.enum(["top", "middle", "bottom"]).nullable().optional(),
      wrap: z.boolean().nullable().optional(),
      indent: z.number().nullable().optional(),
    })
    .nullable()
    .optional(),
  borders: z
    .object({
      top: z
        .object({
          style: z.enum(["solid", "dashed", "dotted", "double"]).nullable().optional(),
          weight: z.enum(["thin", "medium", "thick"]).nullable().optional(),
          color: z.string().nullable().optional(),
        })
        .nullable()
        .optional(),
      right: z
        .object({
          style: z.enum(["solid", "dashed", "dotted", "double"]).nullable().optional(),
          weight: z.enum(["thin", "medium", "thick"]).nullable().optional(),
          color: z.string().nullable().optional(),
        })
        .nullable()
        .optional(),
      bottom: z
        .object({
          style: z.enum(["solid", "dashed", "dotted", "double"]).nullable().optional(),
          weight: z.enum(["thin", "medium", "thick"]).nullable().optional(),
          color: z.string().nullable().optional(),
        })
        .nullable()
        .optional(),
      left: z
        .object({
          style: z.enum(["solid", "dashed", "dotted", "double"]).nullable().optional(),
          weight: z.enum(["thin", "medium", "thick"]).nullable().optional(),
          color: z.string().nullable().optional(),
        })
        .nullable()
        .optional(),
    })
    .nullable()
    .optional(),
});

const cellStyleFieldSchema = z.enum([
  "backgroundColor",
  "fontFamily",
  "fontSize",
  "fontBold",
  "fontItalic",
  "fontUnderline",
  "fontColor",
  "alignmentHorizontal",
  "alignmentVertical",
  "alignmentWrap",
  "alignmentIndent",
  "borderTop",
  "borderRight",
  "borderBottom",
  "borderLeft",
]);

const cellNumberFormatPresetSchema = z.object({
  kind: z.enum([
    "general",
    "number",
    "currency",
    "accounting",
    "percent",
    "date",
    "time",
    "datetime",
    "text",
  ]),
  currency: z.string().optional(),
  decimals: z.number().int().nonnegative().optional(),
  useGrouping: z.boolean().optional(),
  negativeStyle: z.enum(["minus", "parentheses"]).optional(),
  zeroStyle: z.enum(["zero", "dash"]).optional(),
  dateStyle: z.enum(["short", "iso"]).optional(),
});

const cellNumberFormatInputSchema = z.union([z.string(), cellNumberFormatPresetSchema]);

const presenceSelectionSchema = z.object({
  sheetName: z.string().min(1),
  address: z.string().min(1),
});

const defineMutators = defineMutatorsWithType<typeof schema>();
type ZeroMutatorSchema = Parameters<typeof defineMutator>[0];

type JsonValue = string | number | boolean | null | JsonValue[] | { [key: string]: JsonValue };

const jsonValueSchema: z.ZodType<JsonValue> = z.lazy(() =>
  z.union([
    z.string(),
    z.number(),
    z.boolean(),
    z.null(),
    z.array(jsonValueSchema),
    z.object({}).catchall(jsonValueSchema),
  ]),
);

const engineOpBatchSchema = z
  .object({
    id: z.string().min(1),
    replicaId: z.string().min(1),
    clock: z.object({
      counter: z.number().int().nonnegative(),
    }),
    ops: z.array(jsonValueSchema),
  })
  .refine((value): boolean => isEngineOpBatch(value), {
    message: "Invalid engine op batch",
  });

const baseMutationArgsSchema = z.object({
  documentId: z.string().min(1),
  clientMutationId: z.string().min(1).optional(),
});

export const applyBatchArgsSchema = baseMutationArgsSchema.extend({
  batch: engineOpBatchSchema,
});

export const setCellValueArgsSchema = baseMutationArgsSchema.extend({
  sheetName: z.string().min(1),
  address: z.string().min(1),
  value: literalInputSchema,
});

export const setCellFormulaArgsSchema = baseMutationArgsSchema.extend({
  sheetName: z.string().min(1),
  address: z.string().min(1),
  formula: z.string(),
});

export const clearCellArgsSchema = baseMutationArgsSchema.extend({
  sheetName: z.string().min(1),
  address: z.string().min(1),
});

export const renderCommitArgsSchema = baseMutationArgsSchema.extend({
  ops: z.array(jsonValueSchema).refine((value): boolean => isCommitOps(value), {
    message: "Invalid render commit ops",
  }),
});

export const rangeMutationArgsSchema = baseMutationArgsSchema.extend({
  source: cellRangeRefSchema,
  target: cellRangeRefSchema,
});

export const updateRowMetadataArgsSchema = baseMutationArgsSchema.extend({
  sheetName: z.string().min(1),
  startRow: z.number().int().nonnegative(),
  count: z.number().int().positive(),
  height: z.number().int().positive().nullable(),
  hidden: z.boolean().nullable(),
});

export const updateColumnMetadataArgsSchema = baseMutationArgsSchema.extend({
  sheetName: z.string().min(1),
  startCol: z.number().int().nonnegative(),
  count: z.number().int().positive(),
  width: z.number().int().positive().nullable(),
  hidden: z.boolean().nullable(),
});

export const updateColumnWidthArgsSchema = baseMutationArgsSchema.extend({
  sheetName: z.string().min(1),
  columnIndex: z.number().int().nonnegative(),
  width: z.number().int().positive(),
});

export const structuralAxisMutationArgsSchema = baseMutationArgsSchema.extend({
  sheetName: z.string().min(1),
  start: z.number().int().nonnegative(),
  count: z.number().int().positive(),
});

export const setFreezePaneArgsSchema = baseMutationArgsSchema.extend({
  sheetName: z.string().min(1),
  rows: z.number().int().nonnegative(),
  cols: z.number().int().nonnegative(),
});

export const setRangeStyleArgsSchema = baseMutationArgsSchema.extend({
  range: cellRangeRefSchema,
  patch: cellStylePatchSchema,
});

export const clearRangeStyleArgsSchema = baseMutationArgsSchema.extend({
  range: cellRangeRefSchema,
  fields: z.array(cellStyleFieldSchema).optional(),
});

export const setRangeNumberFormatArgsSchema = baseMutationArgsSchema.extend({
  range: cellRangeRefSchema,
  format: cellNumberFormatInputSchema,
});

export const clearRangeNumberFormatArgsSchema = baseMutationArgsSchema.extend({
  range: cellRangeRefSchema,
});

export const clearRangeArgsSchema = baseMutationArgsSchema.extend({
  range: cellRangeRefSchema,
});

export const updatePresenceArgsSchema = baseMutationArgsSchema.extend({
  sessionId: z.string().min(1),
  presenceClientId: z.string().min(1).optional(),
  sheetId: z.number().int().positive().optional(),
  sheetName: z.string().min(1).optional(),
  address: z.string().optional(),
  selection: presenceSelectionSchema.optional(),
});

export const revertWorkbookChangeArgsSchema = baseMutationArgsSchema.extend({
  revision: z.number().int().positive(),
});

export const undoLatestWorkbookChangeArgsSchema = baseMutationArgsSchema;

export const redoLatestWorkbookChangeArgsSchema = baseMutationArgsSchema;

export const applyAgentCommandBundleArgsSchema = baseMutationArgsSchema.extend({
  bundle: z.custom<WorkbookAgentCommandBundle>(
    isWorkbookAgentCommandBundle,
    "Invalid workbook agent command bundle",
  ),
});

const zeroApplyBatchArgsSchema: ZeroMutatorSchema = {
  "~standard": {
    version: applyBatchArgsSchema["~standard"].version,
    vendor: applyBatchArgsSchema["~standard"].vendor,
    validate: (value) => applyBatchArgsSchema["~standard"].validate(value),
  },
};

const zeroRenderCommitArgsSchema: ZeroMutatorSchema = {
  "~standard": {
    version: renderCommitArgsSchema["~standard"].version,
    vendor: renderCommitArgsSchema["~standard"].vendor,
    validate: (value) => renderCommitArgsSchema["~standard"].validate(value),
  },
};

export function parseApplyBatchArgs(args: unknown): {
  documentId: string;
  clientMutationId?: string;
  batch: EngineOpBatch;
} {
  const parsed = applyBatchArgsSchema.parse(args);
  if (!isEngineOpBatch(parsed.batch)) {
    throw new Error("Invalid engine op batch");
  }
  return parsed.clientMutationId === undefined
    ? { documentId: parsed.documentId, batch: parsed.batch }
    : {
        documentId: parsed.documentId,
        clientMutationId: parsed.clientMutationId,
        batch: parsed.batch,
      };
}

export function parseRenderCommitArgs(args: unknown): {
  documentId: string;
  clientMutationId?: string;
  ops: CommitOp[];
} {
  const parsed = renderCommitArgsSchema.parse(args);
  if (!isCommitOps(parsed.ops)) {
    throw new Error("Invalid render commit ops");
  }
  return parsed.clientMutationId === undefined
    ? { documentId: parsed.documentId, ops: parsed.ops }
    : {
        documentId: parsed.documentId,
        clientMutationId: parsed.clientMutationId,
        ops: parsed.ops,
      };
}

async function noop(): Promise<void> {}

export const mutators = defineMutators({
  workbook: {
    applyBatch: defineMutator(zeroApplyBatchArgsSchema, noop),
    setCellValue: defineMutator(setCellValueArgsSchema, noop),
    setCellFormula: defineMutator(setCellFormulaArgsSchema, noop),
    clearCell: defineMutator(clearCellArgsSchema, noop),
    clearRange: defineMutator(clearRangeArgsSchema, noop),
    renderCommit: defineMutator(zeroRenderCommitArgsSchema, noop),
    fillRange: defineMutator(rangeMutationArgsSchema, noop),
    copyRange: defineMutator(rangeMutationArgsSchema, noop),
    moveRange: defineMutator(rangeMutationArgsSchema, noop),
    updateRowMetadata: defineMutator(updateRowMetadataArgsSchema, noop),
    updateColumnMetadata: defineMutator(updateColumnMetadataArgsSchema, noop),
    updateColumnWidth: defineMutator(updateColumnWidthArgsSchema, noop),
    insertRows: defineMutator(structuralAxisMutationArgsSchema, noop),
    deleteRows: defineMutator(structuralAxisMutationArgsSchema, noop),
    insertColumns: defineMutator(structuralAxisMutationArgsSchema, noop),
    deleteColumns: defineMutator(structuralAxisMutationArgsSchema, noop),
    setFreezePane: defineMutator(setFreezePaneArgsSchema, noop),
    setRangeStyle: defineMutator(setRangeStyleArgsSchema, noop),
    clearRangeStyle: defineMutator(clearRangeStyleArgsSchema, noop),
    setRangeNumberFormat: defineMutator(setRangeNumberFormatArgsSchema, noop),
    clearRangeNumberFormat: defineMutator(clearRangeNumberFormatArgsSchema, noop),
    updatePresence: defineMutator(updatePresenceArgsSchema, noop),
    revertChange: defineMutator(revertWorkbookChangeArgsSchema, noop),
    undoLatestChange: defineMutator(undoLatestWorkbookChangeArgsSchema, noop),
    redoLatestChange: defineMutator(redoLatestWorkbookChangeArgsSchema, noop),
  },
});

function toJsonCommitOp(op: CommitOp): JsonValue {
  switch (op.kind) {
    case "upsertWorkbook":
      if (typeof op.name !== "string") {
        throw new Error("Invalid commit op: missing workbook name");
      }
      return { kind: op.kind, name: op.name };
    case "upsertSheet":
      if (typeof op.name !== "string" || typeof op.order !== "number") {
        throw new Error("Invalid commit op: missing sheet payload");
      }
      return { kind: op.kind, name: op.name, order: op.order };
    case "renameSheet":
      if (typeof op.oldName !== "string" || typeof op.newName !== "string") {
        throw new Error("Invalid commit op: missing rename payload");
      }
      return { kind: op.kind, oldName: op.oldName, newName: op.newName };
    case "deleteSheet":
      if (typeof op.name !== "string") {
        throw new Error("Invalid commit op: missing sheet name");
      }
      return { kind: op.kind, name: op.name };
    case "upsertCell":
      if (typeof op.sheetName !== "string" || typeof op.addr !== "string") {
        throw new Error("Invalid commit op: missing cell coordinates");
      }
      if (isLiteralInput(op.value)) {
        return { kind: op.kind, sheetName: op.sheetName, addr: op.addr, value: op.value };
      }
      if (typeof op.formula === "string") {
        return { kind: op.kind, sheetName: op.sheetName, addr: op.addr, formula: op.formula };
      }
      if (typeof op.format === "string") {
        return { kind: op.kind, sheetName: op.sheetName, addr: op.addr, format: op.format };
      }
      throw new Error("Invalid commit op: missing upsertCell payload");
    case "deleteCell":
      if (typeof op.sheetName !== "string" || typeof op.addr !== "string") {
        throw new Error("Invalid commit op: missing deleteCell payload");
      }
      return { kind: op.kind, sheetName: op.sheetName, addr: op.addr };
    default:
      throw new Error(`Unsupported commit op: ${String(op.kind)}`);
  }
}

export function createRenderCommitArgs(args: {
  documentId: string;
  clientMutationId?: string | undefined;
  ops: CommitOp[];
}): { documentId: string; clientMutationId?: string; ops: JsonValue[] } {
  const jsonOps = args.ops.map((op) => toJsonCommitOp(op));
  return args.clientMutationId === undefined
    ? { documentId: args.documentId, ops: jsonOps }
    : {
        documentId: args.documentId,
        clientMutationId: args.clientMutationId,
        ops: jsonOps,
      };
}

export { isLiteralInput };
