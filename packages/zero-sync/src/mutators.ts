import { defineMutator, defineMutatorsWithType } from "@rocicorp/zero";
import { z } from "zod";
import type { CellRangeRef, LiteralInput } from "@bilig/protocol";
import { schema } from "./schema.js";

const literalInputSchema = z.union([z.number(), z.string(), z.boolean(), z.null()]);

const cellRangeRefSchema = z.object({
  sheetName: z.string().min(1),
  startAddress: z.string().min(1),
  endAddress: z.string().min(1),
}) satisfies z.ZodType<CellRangeRef>;

const defineMutators = defineMutatorsWithType<typeof schema>();

export const setCellValueArgsSchema = z.object({
  documentId: z.string().min(1),
  sheetName: z.string().min(1),
  address: z.string().min(1),
  value: literalInputSchema,
});

export const setCellFormulaArgsSchema = z.object({
  documentId: z.string().min(1),
  sheetName: z.string().min(1),
  address: z.string().min(1),
  formula: z.string(),
});

export const clearCellArgsSchema = z.object({
  documentId: z.string().min(1),
  sheetName: z.string().min(1),
  address: z.string().min(1),
});

export const renderCommitArgsSchema = z.object({
  documentId: z.string().min(1),
  ops: z.array(z.any()),
});

export const rangeMutationArgsSchema = z.object({
  documentId: z.string().min(1),
  source: cellRangeRefSchema,
  target: cellRangeRefSchema,
});

export const updateColumnWidthArgsSchema = z.object({
  documentId: z.string().min(1),
  sheetName: z.string().min(1),
  columnIndex: z.number().int().nonnegative(),
  width: z.number().int().positive(),
});

export const replaceSnapshotArgsSchema = z.object({
  documentId: z.string().min(1),
  snapshot: z.any(),
});

async function noop(): Promise<void> {}

export const mutators = defineMutators({
  workbook: {
    setCellValue: defineMutator(setCellValueArgsSchema, noop),
    setCellFormula: defineMutator(setCellFormulaArgsSchema, noop),
    clearCell: defineMutator(clearCellArgsSchema, noop),
    renderCommit: defineMutator(renderCommitArgsSchema, noop),
    fillRange: defineMutator(rangeMutationArgsSchema, noop),
    copyRange: defineMutator(rangeMutationArgsSchema, noop),
    updateColumnWidth: defineMutator(updateColumnWidthArgsSchema, noop),
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
