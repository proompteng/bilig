import { formatAddress, parseCellAddress } from "@bilig/formula";
import type { CellRangeRef } from "@bilig/protocol";
import type { WorkbookAgentUiContext } from "@bilig/contracts";
import { z } from "zod";
import {
  resolveWorkbookSelectorToSingleRange,
  workbookSemanticSelectorSchema,
  type ResolvedWorkbookSelector,
} from "./workbook-selector-resolver.js";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";

export const cellRangeRefSchema = z.object({
  sheetName: z.string().trim().min(1),
  startAddress: z.string().trim().min(1),
  endAddress: z.string().trim().min(1),
});

export const rangeOrSelectorSchema = z
  .object({
    range: cellRangeRefSchema.optional(),
    selector: workbookSemanticSelectorSchema.optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: "Provide exactly one of range or selector",
  });

export const readRangeToolArgsSchema = z.union([
  z.object({
    sheetName: z.string().trim().min(1),
    startAddress: z.string().trim().min(1),
    endAddress: z.string().trim().min(1),
  }),
  z.object({
    selector: workbookSemanticSelectorSchema,
  }),
]);

const writeCellInputSchema = z.union([
  z.string(),
  z.number(),
  z.boolean(),
  z.null(),
  z.object({ value: z.union([z.string(), z.number(), z.boolean(), z.null()]) }),
  z.object({ formula: z.string().min(1) }),
]);

export const writeRangeToolArgsSchema = z.union([
  z.object({
    sheetName: z.string().trim().min(1),
    startAddress: z.string().trim().min(1),
    values: z.array(z.array(writeCellInputSchema).min(1)).min(1),
  }),
  z.object({
    selector: workbookSemanticSelectorSchema,
    values: z.array(z.array(writeCellInputSchema).min(1)).min(1),
  }),
]);

export const transferRangeTargetSchema = z.union([
  cellRangeRefSchema,
  workbookSemanticSelectorSchema,
]);

export const transferRangeToolArgsSchema = z.object({
  source: transferRangeTargetSchema,
  target: transferRangeTargetSchema,
});

export type ReadRangeToolArgs = z.infer<typeof readRangeToolArgsSchema>;
export type WriteRangeToolArgs = z.infer<typeof writeRangeToolArgsSchema>;
export type RangeOrSelectorArgs = z.infer<typeof rangeOrSelectorSchema>;
export type TransferRangeToolArgs = z.infer<typeof transferRangeToolArgsSchema>;

export const workbookSemanticSelectorJsonSchema = {
  oneOf: [
    {
      type: "object",
      additionalProperties: false,
      required: ["kind", "sheet", "start", "end"],
      properties: {
        kind: { type: "string", const: "a1Range" },
        sheet: { type: "string" },
        start: { type: "string" },
        end: { type: "string" },
        revision: { type: "number" },
      },
    },
    {
      type: "object",
      additionalProperties: false,
      required: ["kind", "name"],
      properties: {
        kind: { type: "string", const: "namedRange" },
        name: { type: "string" },
        revision: { type: "number" },
      },
    },
    {
      type: "object",
      additionalProperties: false,
      required: ["kind", "table"],
      properties: {
        kind: { type: "string", const: "table" },
        table: { type: "string" },
        sheet: { type: "string" },
        revision: { type: "number" },
      },
    },
    {
      type: "object",
      additionalProperties: false,
      required: ["kind", "table", "column"],
      properties: {
        kind: { type: "string", const: "tableColumn" },
        table: { type: "string" },
        column: { type: "string" },
        sheet: { type: "string" },
        revision: { type: "number" },
      },
    },
    {
      type: "object",
      additionalProperties: false,
      required: ["kind"],
      properties: {
        kind: { type: "string", const: "currentSelection" },
        revision: { type: "number" },
      },
    },
    {
      type: "object",
      additionalProperties: false,
      required: ["kind"],
      properties: {
        kind: { type: "string", const: "currentRegion" },
        anchor: {
          type: "object",
          additionalProperties: false,
          required: ["sheet", "address"],
          properties: {
            sheet: { type: "string" },
            address: { type: "string" },
          },
        },
        revision: { type: "number" },
      },
    },
    {
      type: "object",
      additionalProperties: false,
      required: ["kind"],
      properties: {
        kind: { type: "string", const: "visibleRows" },
        sheet: { type: "string" },
        revision: { type: "number" },
      },
    },
  ],
};

export const cellRangeRefJsonSchema = {
  type: "object",
  additionalProperties: false,
  required: ["sheetName", "startAddress", "endAddress"],
  properties: {
    sheetName: { type: "string" },
    startAddress: { type: "string" },
    endAddress: { type: "string" },
  },
};

export const rangeOrSelectorJsonSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    range: cellRangeRefJsonSchema,
    selector: workbookSemanticSelectorJsonSchema,
  },
};

function normalizeRange(range: CellRangeRef): CellRangeRef {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(Math.min(start.row, end.row), Math.min(start.col, end.col)),
    endAddress: formatAddress(Math.max(start.row, end.row), Math.max(start.col, end.col)),
  };
}

function resolveSelectorRange(input: {
  runtime: WorkbookRuntime;
  selector: z.infer<typeof workbookSemanticSelectorSchema>;
  uiContext: WorkbookAgentUiContext | null;
}): {
  readonly range: CellRangeRef;
  readonly resolution: ResolvedWorkbookSelector;
} {
  const resolved = resolveWorkbookSelectorToSingleRange(input);
  return {
    range: normalizeRange(resolved.range),
    resolution: resolved.resolution,
  };
}

export function resolveReadRangeRequest(input: {
  runtime: WorkbookRuntime;
  args: ReadRangeToolArgs;
  uiContext: WorkbookAgentUiContext | null;
}): {
  readonly range: CellRangeRef;
  readonly resolution: ResolvedWorkbookSelector | null;
} {
  if ("selector" in input.args) {
    return resolveSelectorRange({
      runtime: input.runtime,
      selector: input.args.selector,
      uiContext: input.uiContext,
    });
  }
  return {
    range: normalizeRange({
      sheetName: input.args.sheetName,
      startAddress: input.args.startAddress,
      endAddress: input.args.endAddress,
    }),
    resolution: null,
  };
}

export function resolveRangeOrSelectorRequest(input: {
  runtime: WorkbookRuntime;
  args: RangeOrSelectorArgs;
  uiContext: WorkbookAgentUiContext | null;
}): {
  readonly range: CellRangeRef;
  readonly resolution: ResolvedWorkbookSelector | null;
} {
  if (input.args.selector) {
    return resolveSelectorRange({
      runtime: input.runtime,
      selector: input.args.selector,
      uiContext: input.uiContext,
    });
  }
  return {
    range: normalizeRange(input.args.range!),
    resolution: null,
  };
}

export function resolveTransferRangeRequest(input: {
  runtime: WorkbookRuntime;
  args: TransferRangeToolArgs;
  uiContext: WorkbookAgentUiContext | null;
}): {
  readonly source: CellRangeRef;
  readonly target: CellRangeRef;
  readonly sourceResolution: ResolvedWorkbookSelector | null;
  readonly targetResolution: ResolvedWorkbookSelector | null;
} {
  const resolveTarget = (
    target: TransferRangeToolArgs["source"],
  ): {
    readonly range: CellRangeRef;
    readonly resolution: ResolvedWorkbookSelector | null;
  } =>
    "kind" in target
      ? resolveSelectorRange({
          runtime: input.runtime,
          selector: target,
          uiContext: input.uiContext,
        })
      : { range: normalizeRange(target), resolution: null };

  const source = resolveTarget(input.args.source);
  const target = resolveTarget(input.args.target);
  return {
    source: source.range,
    target: target.range,
    sourceResolution: source.resolution,
    targetResolution: target.resolution,
  };
}

function countRangeRows(range: CellRangeRef): number {
  const normalized = normalizeRange(range);
  const start = parseCellAddress(normalized.startAddress, normalized.sheetName);
  const end = parseCellAddress(normalized.endAddress, normalized.sheetName);
  return end.row - start.row + 1;
}

function countRangeColumns(range: CellRangeRef): number {
  const normalized = normalizeRange(range);
  const start = parseCellAddress(normalized.startAddress, normalized.sheetName);
  const end = parseCellAddress(normalized.endAddress, normalized.sheetName);
  return end.col - start.col + 1;
}

export function resolveWriteRangeRequest(input: {
  runtime: WorkbookRuntime;
  args: WriteRangeToolArgs;
  uiContext: WorkbookAgentUiContext | null;
}): {
  readonly sheetName: string;
  readonly startAddress: string;
  readonly resolution: ResolvedWorkbookSelector | null;
} {
  if (!("selector" in input.args)) {
    return {
      sheetName: input.args.sheetName,
      startAddress: input.args.startAddress,
      resolution: null,
    };
  }
  const resolved = resolveSelectorRange({
    runtime: input.runtime,
    selector: input.args.selector,
    uiContext: input.uiContext,
  });
  const rowCount = input.args.values.length;
  const columnCount = input.args.values.reduce((max, row) => Math.max(max, row.length), 0);
  const rangeRows = countRangeRows(resolved.range);
  const rangeColumns = countRangeColumns(resolved.range);
  if (
    !(rangeRows === 1 && rangeColumns === 1) &&
    (rangeRows !== rowCount || rangeColumns !== columnCount)
  ) {
    throw new Error(
      `Selector ${resolved.resolution.displayLabel} resolves to ${String(rangeRows)}x${String(rangeColumns)}, but write_range received ${String(rowCount)}x${String(columnCount)} values`,
    );
  }
  return {
    sheetName: resolved.range.sheetName,
    startAddress: resolved.range.startAddress,
    resolution: resolved.resolution,
  };
}
