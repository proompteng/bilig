import { defineMutator, defineMutatorsWithType } from "@rocicorp/zero";
import { z } from "zod";
import type { EngineOpBatch } from "@bilig/workbook-domain";
import type { CellRangeRef, LiteralInput } from "@bilig/protocol";
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

const defineMutators = defineMutatorsWithType<typeof schema>();

const engineOpBatchSchema = z.object({
  id: z.string().min(1),
  replicaId: z.string().min(1),
  clock: z.object({
    counter: z.number().int().nonnegative(),
  }),
  ops: z.array(z.any()),
}) satisfies z.ZodType<EngineOpBatch>;

const baseMutationArgsSchema = z.object({
  documentId: z.string().min(1),
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
  ops: z.array(z.any()),
});

export const rangeMutationArgsSchema = baseMutationArgsSchema.extend({
  source: cellRangeRefSchema,
  target: cellRangeRefSchema,
});

export const updateColumnWidthArgsSchema = baseMutationArgsSchema.extend({
  sheetName: z.string().min(1),
  columnIndex: z.number().int().nonnegative(),
  width: z.number().int().positive(),
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

export const replaceSnapshotArgsSchema = baseMutationArgsSchema.extend({
  snapshot: z.any(),
});

async function noop(): Promise<void> {}

export const mutators = defineMutators({
  workbook: {
    applyBatch: defineMutator(applyBatchArgsSchema, noop),
    setCellValue: defineMutator(setCellValueArgsSchema, noop),
    setCellFormula: defineMutator(setCellFormulaArgsSchema, noop),
    clearCell: defineMutator(clearCellArgsSchema, noop),
    clearRange: defineMutator(clearRangeArgsSchema, noop),
    renderCommit: defineMutator(renderCommitArgsSchema, noop),
    fillRange: defineMutator(rangeMutationArgsSchema, noop),
    copyRange: defineMutator(rangeMutationArgsSchema, noop),
    updateColumnWidth: defineMutator(updateColumnWidthArgsSchema, noop),
    setRangeStyle: defineMutator(setRangeStyleArgsSchema, noop),
    clearRangeStyle: defineMutator(clearRangeStyleArgsSchema, noop),
    setRangeNumberFormat: defineMutator(setRangeNumberFormatArgsSchema, noop),
    clearRangeNumberFormat: defineMutator(clearRangeNumberFormatArgsSchema, noop),
    replaceSnapshot: defineMutator(replaceSnapshotArgsSchema, noop),
  },
});

export function isLiteralInput(value: unknown): value is LiteralInput {
  return (
    value === null ||
    typeof value === "number" ||
    typeof value === "string" ||
    typeof value === "boolean"
  );
}
