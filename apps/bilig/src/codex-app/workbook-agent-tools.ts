import { formatAddress, parseCellAddress } from "@bilig/formula";
import {
  formatErrorCode,
  ValueTag,
  type CellRangeRef,
  type CellNumberFormatInput,
  type CellNumberFormatPreset,
  type CellStylePatch,
  normalizeCellNumberFormatPreset,
} from "@bilig/protocol";
import { WORKBOOK_AGENT_TOOL_NAMES, normalizeWorkbookAgentToolName } from "@bilig/agent-api";
import type {
  CodexDynamicToolCallRequest,
  CodexDynamicToolCallResult,
  CodexDynamicToolSpec,
  JsonValue,
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
} from "@bilig/agent-api";
import {
  clearRangeArgsSchema,
  rangeMutationArgsSchema,
  setRangeNumberFormatArgsSchema,
  setRangeStyleArgsSchema,
  updateColumnMetadataArgsSchema,
  updateRowMetadataArgsSchema,
} from "@bilig/zero-sync";
import type {
  WorkbookAgentUiContext,
  WorkbookAgentWorkflowRun,
  WorkbookViewport,
} from "@bilig/contracts";
import { z } from "zod";
import type { SessionIdentity } from "../http/session.js";
import type { ZeroSyncService } from "../zero/service.js";
import {
  findWorkbookFormulaIssues,
  searchWorkbook,
  summarizeWorkbookStructure,
  traceWorkbookDependencies,
} from "./workbook-agent-comprehension.js";

const MAX_MUTATION_RANGE_CELLS = 400;
const MAX_READ_RANGE_CELLS = 4000;

const writeCellInputSchema = z.union([
  z.string(),
  z.number(),
  z.boolean(),
  z.null(),
  z.object({ value: z.union([z.string(), z.number(), z.boolean(), z.null()]) }),
  z.object({ formula: z.string().min(1) }),
]);

const readRangeToolArgsSchema = z.object({
  sheetName: z.string().min(1),
  startAddress: z.string().min(1),
  endAddress: z.string().min(1),
});

const inspectCellToolArgsSchema = z.object({
  sheetName: z.string().min(1).optional(),
  address: z.string().min(1).optional(),
});
const formulaIssueToolArgsSchema = z.object({
  sheetName: z.string().min(1).optional(),
  limit: z.number().int().positive().max(200).optional(),
});
const readRecentChangesToolArgsSchema = z.object({
  limit: z.number().int().positive().max(50).optional(),
});
const startWorkflowToolArgsSchema = z.discriminatedUnion("workflowTemplate", [
  z.object({
    workflowTemplate: z.literal("summarizeWorkbook"),
  }),
  z.object({
    workflowTemplate: z.literal("summarizeCurrentSheet"),
  }),
  z.object({
    workflowTemplate: z.literal("describeRecentChanges"),
  }),
  z.object({
    workflowTemplate: z.literal("findFormulaIssues"),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(200).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("highlightFormulaIssues"),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(200).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("highlightCurrentSheetOutliers"),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(100).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("styleCurrentSheetHeaders"),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("normalizeCurrentSheetHeaders"),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("normalizeCurrentSheetNumberFormats"),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("normalizeCurrentSheetWhitespace"),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("fillCurrentSheetFormulasDown"),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("traceSelectionDependencies"),
  }),
  z.object({
    workflowTemplate: z.literal("explainSelectionCell"),
  }),
  z.object({
    workflowTemplate: z.literal("searchWorkbookQuery"),
    query: z.string().trim().min(1),
    sheetName: z.string().min(1).optional(),
    limit: z.number().int().positive().max(50).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("createCurrentSheetRollup"),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("createCurrentSheetReviewTab"),
    sheetName: z.string().min(1).optional(),
  }),
  z.object({
    workflowTemplate: z.literal("createSheet"),
    name: z.string().trim().min(1),
  }),
  z.object({
    workflowTemplate: z.literal("renameCurrentSheet"),
    name: z.string().trim().min(1),
  }),
  z.object({
    workflowTemplate: z.literal("hideCurrentRow"),
  }),
  z.object({
    workflowTemplate: z.literal("hideCurrentColumn"),
  }),
  z.object({
    workflowTemplate: z.literal("unhideCurrentRow"),
  }),
  z.object({
    workflowTemplate: z.literal("unhideCurrentColumn"),
  }),
]);
export type WorkbookAgentStartWorkflowRequest = z.infer<typeof startWorkflowToolArgsSchema>;
const searchWorkbookToolArgsSchema = z.object({
  query: z.string().trim().min(1),
  sheetName: z.string().min(1).optional(),
  limit: z.number().int().positive().max(50).optional(),
});
const traceDependenciesToolArgsSchema = z.object({
  sheetName: z.string().min(1).optional(),
  address: z.string().min(1).optional(),
  direction: z.enum(["precedents", "dependents", "both"]).optional(),
  depth: z.number().int().positive().max(4).optional(),
});

const writeRangeToolArgsSchema = z.object({
  sheetName: z.string().min(1),
  startAddress: z.string().min(1),
  values: z.array(z.array(writeCellInputSchema).min(1)).min(1),
});

const sheetMutationToolArgsSchema = z.object({
  name: z.string().trim().min(1),
});

const renameSheetToolArgsSchema = z.object({
  currentName: z.string().trim().min(1),
  nextName: z.string().trim().min(1),
});

const rowMetadataToolArgsSchema = z
  .object({
    sheetName: updateRowMetadataArgsSchema.shape.sheetName,
    startRow: updateRowMetadataArgsSchema.shape.startRow,
    count: updateRowMetadataArgsSchema.shape.count,
    height: updateRowMetadataArgsSchema.shape.height.optional(),
    hidden: updateRowMetadataArgsSchema.shape.hidden.optional(),
  })
  .refine((value) => value.height !== undefined || value.hidden !== undefined, {
    message: "height or hidden is required",
  });

const columnMetadataToolArgsSchema = z
  .object({
    sheetName: updateColumnMetadataArgsSchema.shape.sheetName,
    startCol: updateColumnMetadataArgsSchema.shape.startCol,
    count: updateColumnMetadataArgsSchema.shape.count,
    width: updateColumnMetadataArgsSchema.shape.width.optional(),
    hidden: updateColumnMetadataArgsSchema.shape.hidden.optional(),
  })
  .refine((value) => value.width !== undefined || value.hidden !== undefined, {
    message: "width or hidden is required",
  });

const clearRangeToolArgsSchema = clearRangeArgsSchema.pick({ range: true });
const transferRangeToolArgsSchema = rangeMutationArgsSchema.pick({ source: true, target: true });
const formatRangeToolArgsSchema = z
  .object({
    range: clearRangeToolArgsSchema.shape.range,
    patch: setRangeStyleArgsSchema.shape.patch.optional(),
    numberFormat: setRangeNumberFormatArgsSchema.shape.format.optional(),
  })
  .refine((value) => value.patch !== undefined || value.numberFormat !== undefined, {
    message: "patch or numberFormat is required",
  });

type FormatRangePatchInput = NonNullable<z.infer<typeof setRangeStyleArgsSchema.shape.patch>>;
type FormatRangeNumberFormatInput = NonNullable<
  z.infer<typeof setRangeNumberFormatArgsSchema.shape.format>
>;

function textToolResult(text: string, success = true): CodexDynamicToolCallResult {
  return {
    success,
    contentItems: [
      {
        type: "inputText",
        text,
      },
    ],
  };
}

function stringifyJson(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

function summarizeWorkbookChangeRecord(
  record: Awaited<ReturnType<ZeroSyncService["listWorkbookChanges"]>>[number],
) {
  return {
    revision: record.revision,
    actorUserId: record.actorUserId,
    eventKind: record.eventKind,
    summary: record.summary,
    sheetName: record.sheetName,
    anchorAddress: record.anchorAddress,
    range: record.range,
    createdAtUnixMs: record.createdAtUnixMs,
    revertedByRevision: record.revertedByRevision,
    revertsRevision: record.revertsRevision,
  };
}

function normalizeFormula(formula: string): string {
  return formula.startsWith("=") ? formula.slice(1) : formula;
}

function normalizeStylePatch(patch: FormatRangePatchInput): CellStylePatch {
  const normalized: CellStylePatch = {};
  if (patch.fill === null) {
    normalized.fill = null;
  } else if (patch.fill !== undefined) {
    const fill: NonNullable<CellStylePatch["fill"]> = {};
    if (patch.fill.backgroundColor !== undefined) {
      fill.backgroundColor = patch.fill.backgroundColor;
    }
    normalized.fill = fill;
  }
  if (patch.font === null) {
    normalized.font = null;
  } else if (patch.font !== undefined) {
    const font: NonNullable<CellStylePatch["font"]> = {};
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
  if (patch.alignment === null) {
    normalized.alignment = null;
  } else if (patch.alignment !== undefined) {
    const alignment: NonNullable<CellStylePatch["alignment"]> = {};
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
  if (patch.borders === null) {
    normalized.borders = null;
  } else if (patch.borders !== undefined) {
    const borders: NonNullable<CellStylePatch["borders"]> = {};
    const sides = [
      ["top", patch.borders.top],
      ["right", patch.borders.right],
      ["bottom", patch.borders.bottom],
      ["left", patch.borders.left],
    ] as const;
    for (const [sideName, sideValue] of sides) {
      if (sideValue === undefined) {
        continue;
      }
      if (sideValue === null) {
        borders[sideName] = null;
        continue;
      }
      const side: NonNullable<NonNullable<CellStylePatch["borders"]>[typeof sideName]> = {};
      if (sideValue.style !== undefined) {
        side.style = sideValue.style;
      }
      if (sideValue.weight !== undefined) {
        side.weight = sideValue.weight;
      }
      if (sideValue.color !== undefined) {
        side.color = sideValue.color;
      }
      borders[sideName] = side;
    }
    normalized.borders = borders;
  }
  return normalized;
}

function normalizeNumberFormatInput(input: FormatRangeNumberFormatInput): CellNumberFormatInput {
  if (typeof input === "string") {
    return input;
  }

  const preset: CellNumberFormatPreset = {
    kind: input.kind,
  };
  if (typeof input.currency === "string") {
    preset.currency = input.currency;
  }
  if (typeof input.decimals === "number") {
    preset.decimals = input.decimals;
  }
  if (typeof input.useGrouping === "boolean") {
    preset.useGrouping = input.useGrouping;
  }
  if (input.negativeStyle === "minus" || input.negativeStyle === "parentheses") {
    preset.negativeStyle = input.negativeStyle;
  }
  if (input.zeroStyle === "zero" || input.zeroStyle === "dash") {
    preset.zeroStyle = input.zeroStyle;
  }
  if (input.dateStyle === "short" || input.dateStyle === "iso") {
    preset.dateStyle = input.dateStyle;
  }

  return normalizeCellNumberFormatPreset(preset);
}

function normalizeRange(range: CellRangeRef): CellRangeRef & {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
} {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);
  return {
    ...range,
    startAddress: formatAddress(startRow, startCol),
    endAddress: formatAddress(endRow, endCol),
    startRow,
    endRow,
    startCol,
    endCol,
  };
}

function countRangeCells(range: CellRangeRef): number {
  const bounds = normalizeRange(range);
  return (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1);
}

function ensureRangeLimit(range: CellRangeRef, limit: number): void {
  const count = countRangeCells(range);
  if (count > limit) {
    throw new Error(
      `Range ${range.sheetName}!${range.startAddress}:${range.endAddress} has ${String(count)} cells; tool limit is ${String(limit)} cells per call`,
    );
  }
}

function resolveSelectionRange(context: WorkbookAgentUiContext | null): CellRangeRef {
  if (!context) {
    throw new Error("No browser workbook context is attached to this chat session");
  }
  return {
    sheetName: context.selection.sheetName,
    startAddress: context.selection.address,
    endAddress: context.selection.address,
  };
}

function resolveVisibleRange(context: WorkbookAgentUiContext | null): CellRangeRef {
  if (!context) {
    throw new Error("No browser workbook context is attached to this chat session");
  }
  return viewportToRange(context.selection.sheetName, context.viewport);
}

function resolveInspectionTarget(
  context: WorkbookAgentUiContext | null,
  args: z.infer<typeof inspectCellToolArgsSchema>,
): {
  sheetName: string;
  address: string;
} {
  if (args.sheetName && args.address) {
    return {
      sheetName: args.sheetName,
      address: args.address,
    };
  }
  if (!context) {
    throw new Error("sheetName and address are required when no browser workbook context exists");
  }
  return context.selection;
}

function viewportToRange(sheetName: string, viewport: WorkbookViewport): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(viewport.rowStart, viewport.colStart),
    endAddress: formatAddress(viewport.rowEnd, viewport.colEnd),
  };
}

function serializeCellValue(value: {
  tag: ValueTag;
  value?: number | boolean | string;
  code?: number;
}): JsonValue {
  switch (value.tag) {
    case ValueTag.Empty:
      return null;
    case ValueTag.Number:
    case ValueTag.Boolean:
    case ValueTag.String:
      return value.value ?? null;
    case ValueTag.Error:
      return typeof value.code === "number" ? formatErrorCode(value.code) : "#ERROR!";
    default:
      return null;
  }
}

function workbookToolContextText(context: WorkbookAgentUiContext | null): string {
  if (!context) {
    return "No browser view context is attached to this chat session yet.";
  }
  const visibleRange = viewportToRange(context.selection.sheetName, context.viewport);
  return stringifyJson({
    selection: context.selection,
    visibleRange: {
      sheetName: visibleRange.sheetName,
      startAddress: visibleRange.startAddress,
      endAddress: visibleRange.endAddress,
    },
  });
}

async function inspectWorkbookRange(
  context: WorkbookAgentToolContext,
  range: CellRangeRef,
): Promise<CodexDynamicToolCallResult> {
  const normalizedRange = normalizeRange(range);
  ensureRangeLimit(normalizedRange, MAX_READ_RANGE_CELLS);
  const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
    const rows: JsonValue[] = [];
    for (let row = normalizedRange.startRow; row <= normalizedRange.endRow; row += 1) {
      const rowEntries: JsonValue[] = [];
      for (let col = normalizedRange.startCol; col <= normalizedRange.endCol; col += 1) {
        const cell = runtime.engine.getCell(normalizedRange.sheetName, formatAddress(row, col));
        rowEntries.push({
          address: cell.address,
          value: serializeCellValue(cell.value),
          ...(cell.formula !== undefined ? { formula: `=${cell.formula}` } : {}),
          ...(cell.format !== undefined ? { format: cell.format } : {}),
        });
      }
      rows.push(rowEntries);
    }
    return {
      range: {
        sheetName: normalizedRange.sheetName,
        startAddress: normalizedRange.startAddress,
        endAddress: normalizedRange.endAddress,
      },
      rows,
    };
  });
  return textToolResult(stringifyJson(result));
}

async function inspectWorkbookCell(
  context: WorkbookAgentToolContext,
  target: {
    sheetName: string;
    address: string;
  },
): Promise<CodexDynamicToolCallResult> {
  const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
    const cell = runtime.engine.explainCell(target.sheetName, target.address);
    return {
      sheetName: cell.sheetName,
      address: cell.address,
      value: serializeCellValue(cell.value),
      formula: cell.formula !== undefined ? `=${cell.formula}` : null,
      format: cell.format ?? null,
      version: cell.version,
      inCycle: cell.inCycle,
      mode: cell.mode ?? null,
      topoRank: cell.topoRank ?? null,
      directPrecedents: [...cell.directPrecedents],
      directDependents: [...cell.directDependents],
    };
  });
  return textToolResult(stringifyJson(result));
}

export interface WorkbookAgentToolContext {
  readonly documentId: string;
  readonly session: SessionIdentity;
  readonly uiContext: WorkbookAgentUiContext | null;
  readonly zeroSyncService: ZeroSyncService;
  readonly stageCommand: (command: WorkbookAgentCommand) => Promise<WorkbookAgentCommandBundle>;
  readonly startWorkflow?: (
    input: WorkbookAgentStartWorkflowRequest,
  ) => Promise<WorkbookAgentWorkflowRun>;
}

function createDynamicToolSpecs(): readonly CodexDynamicToolSpec[] {
  return [
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.getContext,
      description:
        "Read the current browser workbook context, including the active cell selection and visible viewport.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readWorkbook,
      description:
        "Read a workbook summary with sheet names, populated cell counts, and used ranges.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRange,
      description:
        "Read a rectangular cell range. Use this before editing a region you have not inspected yet.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["sheetName", "startAddress", "endAddress"],
        properties: {
          sheetName: { type: "string" },
          startAddress: { type: "string" },
          endAddress: { type: "string" },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readSelection,
      description: "Read the currently selected cell from the attached browser workbook context.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange,
      description:
        "Read the currently visible viewport range from the attached browser workbook context.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges,
      description:
        "Read the most recent durable workbook changes, including revisions, summaries, and affected ranges.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {
          limit: { type: "number" },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.startWorkflow,
      description:
        "Start a built-in durable workbook workflow for saved workbook summaries, formula review/highlight tasks, formatting-cleanup tasks like numeric outlier highlighting or consistent header styling, import-cleanup tasks like header, number-format, whitespace normalization, or formula fill-down cleanup, search/report tasks, rollup previews, review-tab creation, or safe structural preview workflows like create-sheet, rename-sheet, and row/column visibility changes.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["workflowTemplate"],
        properties: {
          workflowTemplate: {
            type: "string",
            enum: [
              "summarizeWorkbook",
              "summarizeCurrentSheet",
              "describeRecentChanges",
              "findFormulaIssues",
              "highlightFormulaIssues",
              "highlightCurrentSheetOutliers",
              "styleCurrentSheetHeaders",
              "normalizeCurrentSheetHeaders",
              "normalizeCurrentSheetNumberFormats",
              "normalizeCurrentSheetWhitespace",
              "fillCurrentSheetFormulasDown",
              "traceSelectionDependencies",
              "explainSelectionCell",
              "searchWorkbookQuery",
              "createCurrentSheetRollup",
              "createCurrentSheetReviewTab",
              "createSheet",
              "renameCurrentSheet",
              "hideCurrentRow",
              "hideCurrentColumn",
              "unhideCurrentRow",
              "unhideCurrentColumn",
            ],
          },
          query: { type: "string" },
          sheetName: { type: "string" },
          limit: { type: "number" },
          name: { type: "string" },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.inspectCell,
      description:
        "Explain one cell, including its current value, formula, version, cycle status, and direct precedents/dependents. Defaults to the current selection when no address is provided.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {
          sheetName: { type: "string" },
          address: { type: "string" },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues,
      description:
        "Scan the workbook for broken formulas, error cells, cycles, and formulas still running through the JS fallback path.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {
          sheetName: { type: "string" },
          limit: { type: "number" },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook,
      description:
        "Search workbook sheet names, addresses, formulas, inputs, and displayed values through the warm local runtime.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["query"],
        properties: {
          query: { type: "string" },
          sheetName: { type: "string" },
          limit: { type: "number" },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.traceDependencies,
      description:
        "Trace workbook precedents and dependents from one cell for multiple hops. Defaults to the current selection when no address is provided.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {
          sheetName: { type: "string" },
          address: { type: "string" },
          direction: {
            type: "string",
            enum: ["precedents", "dependents", "both"],
          },
          depth: { type: "number" },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.writeRange,
      description:
        "Write a rectangular matrix of spreadsheet inputs starting at a top-left address. Use primitives for literals, {formula} for formulas, and null to clear a cell.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["sheetName", "startAddress", "values"],
        properties: {
          sheetName: { type: "string" },
          startAddress: { type: "string" },
          values: {
            type: "array",
            items: {
              type: "array",
              items: {
                oneOf: [
                  { type: "string" },
                  { type: "number" },
                  { type: "boolean" },
                  { type: "null" },
                  {
                    type: "object",
                    additionalProperties: false,
                    required: ["value"],
                    properties: {
                      value: {
                        oneOf: [
                          { type: "string" },
                          { type: "number" },
                          { type: "boolean" },
                          { type: "null" },
                        ],
                      },
                    },
                  },
                  {
                    type: "object",
                    additionalProperties: false,
                    required: ["formula"],
                    properties: {
                      formula: { type: "string" },
                    },
                  },
                ],
              },
            },
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.clearRange,
      description: "Clear a rectangular range of cells.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["range"],
        properties: {
          range: {
            type: "object",
            additionalProperties: false,
            required: ["sheetName", "startAddress", "endAddress"],
            properties: {
              sheetName: { type: "string" },
              startAddress: { type: "string" },
              endAddress: { type: "string" },
            },
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.formatRange,
      description:
        "Apply style and/or number-format changes to a range. Use patch for style properties and numberFormat for number formatting.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["range"],
        properties: {
          range: {
            type: "object",
            additionalProperties: false,
            required: ["sheetName", "startAddress", "endAddress"],
            properties: {
              sheetName: { type: "string" },
              startAddress: { type: "string" },
              endAddress: { type: "string" },
            },
          },
          patch: { type: "object" },
          numberFormat: {
            oneOf: [{ type: "string" }, { type: "object" }],
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.fillRange,
      description: "Fill a target range from a source range using spreadsheet fill semantics.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["source", "target"],
        properties: {
          source: {
            type: "object",
            additionalProperties: false,
            required: ["sheetName", "startAddress", "endAddress"],
            properties: {
              sheetName: { type: "string" },
              startAddress: { type: "string" },
              endAddress: { type: "string" },
            },
          },
          target: {
            type: "object",
            additionalProperties: false,
            required: ["sheetName", "startAddress", "endAddress"],
            properties: {
              sheetName: { type: "string" },
              startAddress: { type: "string" },
              endAddress: { type: "string" },
            },
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.copyRange,
      description: "Copy a source range into a target range.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["source", "target"],
        properties: {
          source: {
            type: "object",
            additionalProperties: false,
            required: ["sheetName", "startAddress", "endAddress"],
            properties: {
              sheetName: { type: "string" },
              startAddress: { type: "string" },
              endAddress: { type: "string" },
            },
          },
          target: {
            type: "object",
            additionalProperties: false,
            required: ["sheetName", "startAddress", "endAddress"],
            properties: {
              sheetName: { type: "string" },
              startAddress: { type: "string" },
              endAddress: { type: "string" },
            },
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.moveRange,
      description: "Move a source range into a target range.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["source", "target"],
        properties: {
          source: {
            type: "object",
            additionalProperties: false,
            required: ["sheetName", "startAddress", "endAddress"],
            properties: {
              sheetName: { type: "string" },
              startAddress: { type: "string" },
              endAddress: { type: "string" },
            },
          },
          target: {
            type: "object",
            additionalProperties: false,
            required: ["sheetName", "startAddress", "endAddress"],
            properties: {
              sheetName: { type: "string" },
              startAddress: { type: "string" },
              endAddress: { type: "string" },
            },
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.createSheet,
      description: "Create a new worksheet.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["name"],
        properties: {
          name: { type: "string" },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.renameSheet,
      description: "Rename an existing worksheet.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["currentName", "nextName"],
        properties: {
          currentName: { type: "string" },
          nextName: { type: "string" },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.updateRowMetadata,
      description:
        "Hide, unhide, resize, or reset row metadata across a bounded row span on one sheet.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["sheetName", "startRow", "count"],
        properties: {
          sheetName: { type: "string" },
          startRow: { type: "number" },
          count: { type: "number" },
          height: {
            oneOf: [{ type: "number" }, { type: "null" }],
          },
          hidden: {
            oneOf: [{ type: "boolean" }, { type: "null" }],
          },
        },
      },
    },
    {
      name: WORKBOOK_AGENT_TOOL_NAMES.updateColumnMetadata,
      description:
        "Hide, unhide, resize, or reset column metadata across a bounded column span on one sheet.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        required: ["sheetName", "startCol", "count"],
        properties: {
          sheetName: { type: "string" },
          startCol: { type: "number" },
          count: { type: "number" },
          width: {
            oneOf: [{ type: "number" }, { type: "null" }],
          },
          hidden: {
            oneOf: [{ type: "boolean" }, { type: "null" }],
          },
        },
      },
    },
  ] satisfies readonly CodexDynamicToolSpec[];
}

export const workbookAgentDynamicToolSpecs = createDynamicToolSpecs();

async function stageCommandResult(
  context: WorkbookAgentToolContext,
  command: WorkbookAgentCommand,
): Promise<CodexDynamicToolCallResult> {
  const bundle = await context.stageCommand(command);
  return textToolResult(
    stringifyJson({
      staged: true,
      bundleId: bundle.id,
      summary: bundle.summary,
      scope: bundle.scope,
      riskClass: bundle.riskClass,
      estimatedAffectedCells: bundle.estimatedAffectedCells,
      affectedRanges: bundle.affectedRanges,
    }),
  );
}

function workflowToolResult(run: WorkbookAgentWorkflowRun): CodexDynamicToolCallResult {
  return textToolResult(
    stringifyJson({
      workflowRun: {
        runId: run.runId,
        workflowTemplate: run.workflowTemplate,
        title: run.title,
        summary: run.summary,
        status: run.status,
        completedAtUnixMs: run.completedAtUnixMs,
        errorMessage: run.errorMessage,
      },
      artifact: run.artifact,
    }),
  );
}

export async function handleWorkbookAgentToolCall(
  context: WorkbookAgentToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult> {
  try {
    switch (normalizeWorkbookAgentToolName(request.tool)) {
      case WORKBOOK_AGENT_TOOL_NAMES.getContext: {
        return textToolResult(workbookToolContextText(context.uiContext));
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readWorkbook: {
        const summary = await context.zeroSyncService.inspectWorkbook(
          context.documentId,
          (runtime) => ({
            documentId: context.documentId,
            context: context.uiContext,
            ...summarizeWorkbookStructure(runtime),
          }),
        );
        return textToolResult(stringifyJson(summary));
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRange: {
        const args = readRangeToolArgsSchema.parse(request.arguments);
        return await inspectWorkbookRange(context, {
          sheetName: args.sheetName,
          startAddress: args.startAddress,
          endAddress: args.endAddress,
        });
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readSelection: {
        return await inspectWorkbookRange(context, resolveSelectionRange(context.uiContext));
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange: {
        return await inspectWorkbookRange(context, resolveVisibleRange(context.uiContext));
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges: {
        const args = readRecentChangesToolArgsSchema.parse(request.arguments);
        const changes = await context.zeroSyncService.listWorkbookChanges(
          context.documentId,
          args.limit,
        );
        return textToolResult(
          stringifyJson({
            documentId: context.documentId,
            changeCount: changes.length,
            changes: changes.map((record) => summarizeWorkbookChangeRecord(record)),
          }),
        );
      }
      case WORKBOOK_AGENT_TOOL_NAMES.startWorkflow: {
        const args = startWorkflowToolArgsSchema.parse(request.arguments);
        if (!context.startWorkflow) {
          throw new Error("Built-in workflow execution is not available in this session");
        }
        return workflowToolResult(await context.startWorkflow(args));
      }
      case WORKBOOK_AGENT_TOOL_NAMES.inspectCell: {
        const args = inspectCellToolArgsSchema.parse(request.arguments);
        return await inspectWorkbookCell(context, resolveInspectionTarget(context.uiContext, args));
      }
      case WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues: {
        const args = formulaIssueToolArgsSchema.parse(request.arguments);
        const report = await context.zeroSyncService.inspectWorkbook(
          context.documentId,
          (runtime) =>
            findWorkbookFormulaIssues(runtime, {
              ...(args.sheetName ? { sheetName: args.sheetName } : {}),
              ...(args.limit !== undefined ? { limit: args.limit } : {}),
            }),
        );
        return textToolResult(stringifyJson(report));
      }
      case WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook: {
        const args = searchWorkbookToolArgsSchema.parse(request.arguments);
        const report = await context.zeroSyncService.inspectWorkbook(
          context.documentId,
          (runtime) =>
            searchWorkbook(runtime, {
              query: args.query,
              ...(args.sheetName ? { sheetName: args.sheetName } : {}),
              ...(args.limit !== undefined ? { limit: args.limit } : {}),
            }),
        );
        return textToolResult(stringifyJson(report));
      }
      case WORKBOOK_AGENT_TOOL_NAMES.traceDependencies: {
        const args = traceDependenciesToolArgsSchema.parse(request.arguments);
        const target = resolveInspectionTarget(context.uiContext, args);
        const report = await context.zeroSyncService.inspectWorkbook(
          context.documentId,
          (runtime) =>
            traceWorkbookDependencies(runtime, {
              sheetName: target.sheetName,
              address: target.address,
              ...(args.direction ? { direction: args.direction } : {}),
              ...(args.depth !== undefined ? { depth: args.depth } : {}),
            }),
        );
        return textToolResult(stringifyJson(report));
      }
      case WORKBOOK_AGENT_TOOL_NAMES.writeRange: {
        const args = writeRangeToolArgsSchema.parse(request.arguments);
        const start = parseCellAddress(args.startAddress, args.sheetName);
        const maxWidth = args.values.reduce(
          (width, rowValues) => Math.max(width, rowValues.length),
          0,
        );
        const endAddress = formatAddress(
          start.row + args.values.length - 1,
          start.col + maxWidth - 1,
        );
        ensureRangeLimit(
          {
            sheetName: args.sheetName,
            startAddress: args.startAddress,
            endAddress,
          },
          MAX_MUTATION_RANGE_CELLS,
        );
        return await stageCommandResult(context, {
          kind: "writeRange",
          sheetName: args.sheetName,
          startAddress: args.startAddress,
          values: args.values.map((rowValues) =>
            rowValues.map((cellInput) => {
              if (
                cellInput === null ||
                typeof cellInput === "string" ||
                typeof cellInput === "number" ||
                typeof cellInput === "boolean"
              ) {
                return cellInput;
              }
              if ("formula" in cellInput) {
                return {
                  formula: `=${normalizeFormula(cellInput.formula)}`,
                };
              }
              return {
                value: cellInput.value,
              };
            }),
          ),
        });
      }
      case WORKBOOK_AGENT_TOOL_NAMES.clearRange: {
        const args = clearRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.range, MAX_MUTATION_RANGE_CELLS);
        return await stageCommandResult(context, {
          kind: "clearRange",
          range: args.range,
        });
      }
      case WORKBOOK_AGENT_TOOL_NAMES.formatRange: {
        const args = formatRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.range, MAX_MUTATION_RANGE_CELLS);
        const formatCommand: Extract<WorkbookAgentCommand, { kind: "formatRange" }> = {
          kind: "formatRange",
          range: args.range,
        };
        if (args.patch !== undefined) {
          formatCommand.patch = normalizeStylePatch(args.patch);
        }
        if (args.numberFormat !== undefined) {
          formatCommand.numberFormat = normalizeNumberFormatInput(args.numberFormat);
        }
        return await stageCommandResult(context, formatCommand);
      }
      case WORKBOOK_AGENT_TOOL_NAMES.fillRange: {
        const args = transferRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.source, MAX_MUTATION_RANGE_CELLS);
        ensureRangeLimit(args.target, MAX_MUTATION_RANGE_CELLS);
        return await stageCommandResult(context, {
          kind: "fillRange",
          source: args.source,
          target: args.target,
        });
      }
      case WORKBOOK_AGENT_TOOL_NAMES.copyRange: {
        const args = transferRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.source, MAX_MUTATION_RANGE_CELLS);
        ensureRangeLimit(args.target, MAX_MUTATION_RANGE_CELLS);
        return await stageCommandResult(context, {
          kind: "copyRange",
          source: args.source,
          target: args.target,
        });
      }
      case WORKBOOK_AGENT_TOOL_NAMES.moveRange: {
        const args = transferRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.source, MAX_MUTATION_RANGE_CELLS);
        ensureRangeLimit(args.target, MAX_MUTATION_RANGE_CELLS);
        return await stageCommandResult(context, {
          kind: "moveRange",
          source: args.source,
          target: args.target,
        });
      }
      case WORKBOOK_AGENT_TOOL_NAMES.createSheet: {
        const args = sheetMutationToolArgsSchema.parse(request.arguments);
        return await stageCommandResult(context, {
          kind: "createSheet",
          name: args.name,
        });
      }
      case WORKBOOK_AGENT_TOOL_NAMES.renameSheet: {
        const args = renameSheetToolArgsSchema.parse(request.arguments);
        return await stageCommandResult(context, {
          kind: "renameSheet",
          currentName: args.currentName,
          nextName: args.nextName,
        });
      }
      case WORKBOOK_AGENT_TOOL_NAMES.updateRowMetadata: {
        const args = rowMetadataToolArgsSchema.parse(request.arguments);
        return await stageCommandResult(context, {
          kind: "updateRowMetadata",
          sheetName: args.sheetName,
          startRow: args.startRow,
          count: args.count,
          ...(args.height !== undefined
            ? {
                height: args.height === null ? null : Math.max(1, Math.round(args.height)),
              }
            : {}),
          ...(args.hidden !== undefined ? { hidden: args.hidden } : {}),
        });
      }
      case WORKBOOK_AGENT_TOOL_NAMES.updateColumnMetadata: {
        const args = columnMetadataToolArgsSchema.parse(request.arguments);
        return await stageCommandResult(context, {
          kind: "updateColumnMetadata",
          sheetName: args.sheetName,
          startCol: args.startCol,
          count: args.count,
          ...(args.width !== undefined
            ? {
                width: args.width === null ? null : Math.max(1, Math.round(args.width)),
              }
            : {}),
          ...(args.hidden !== undefined ? { hidden: args.hidden } : {}),
        });
      }
      default:
        return textToolResult(`Unknown bilig tool: ${request.tool}`, false);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return textToolResult(`Tool ${request.tool} failed: ${message}`, false);
  }
}
