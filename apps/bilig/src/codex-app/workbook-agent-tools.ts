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
} from "@bilig/zero-sync";
import type { WorkbookAgentUiContext, WorkbookViewport } from "@bilig/contracts";
import { z } from "zod";
import type { SessionIdentity } from "../http/session.js";
import type { ZeroSyncService } from "../zero/service.js";
import {
  findWorkbookFormulaIssues,
  searchWorkbook,
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
}

function createDynamicToolSpecs(): readonly CodexDynamicToolSpec[] {
  return [
    {
      name: "bilig.get_context",
      description:
        "Read the current browser workbook context, including the active cell selection and visible viewport.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: "bilig.read_workbook",
      description:
        "Read a workbook summary with sheet names, populated cell counts, and used ranges.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: "bilig.read_range",
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
      name: "bilig.read_selection",
      description: "Read the currently selected cell from the attached browser workbook context.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: "bilig.read_visible_range",
      description:
        "Read the currently visible viewport range from the attached browser workbook context.",
      inputSchema: {
        type: "object",
        additionalProperties: false,
        properties: {},
      },
    },
    {
      name: "bilig.inspect_cell",
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
      name: "bilig.find_formula_issues",
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
      name: "bilig.search_workbook",
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
      name: "bilig.trace_dependencies",
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
      name: "bilig.write_range",
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
      name: "bilig.clear_range",
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
      name: "bilig.format_range",
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
      name: "bilig.fill_range",
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
      name: "bilig.copy_range",
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
      name: "bilig.move_range",
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
      name: "bilig.create_sheet",
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
      name: "bilig.rename_sheet",
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

export async function handleWorkbookAgentToolCall(
  context: WorkbookAgentToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult> {
  try {
    switch (request.tool) {
      case "bilig.get_context": {
        return textToolResult(workbookToolContextText(context.uiContext));
      }
      case "bilig.read_workbook": {
        const summary = await context.zeroSyncService.inspectWorkbook(
          context.documentId,
          (runtime) => {
            const snapshot = runtime.engine.exportSnapshot();
            return {
              documentId: context.documentId,
              context: context.uiContext,
              sheets: [...snapshot.sheets]
                .toSorted((left, right) => left.order - right.order)
                .map((sheet) => {
                  let minRow = Number.POSITIVE_INFINITY;
                  let maxRow = Number.NEGATIVE_INFINITY;
                  let minCol = Number.POSITIVE_INFINITY;
                  let maxCol = Number.NEGATIVE_INFINITY;
                  for (const cell of sheet.cells) {
                    const parsed = parseCellAddress(cell.address, sheet.name);
                    minRow = Math.min(minRow, parsed.row);
                    maxRow = Math.max(maxRow, parsed.row);
                    minCol = Math.min(minCol, parsed.col);
                    maxCol = Math.max(maxCol, parsed.col);
                  }
                  return {
                    name: sheet.name,
                    order: sheet.order,
                    cellCount: sheet.cells.length,
                    usedRange:
                      sheet.cells.length === 0
                        ? null
                        : {
                            startAddress: formatAddress(minRow, minCol),
                            endAddress: formatAddress(maxRow, maxCol),
                          },
                  };
                }),
            };
          },
        );
        return textToolResult(stringifyJson(summary));
      }
      case "bilig.read_range": {
        const args = readRangeToolArgsSchema.parse(request.arguments);
        return await inspectWorkbookRange(context, {
          sheetName: args.sheetName,
          startAddress: args.startAddress,
          endAddress: args.endAddress,
        });
      }
      case "bilig.read_selection": {
        return await inspectWorkbookRange(context, resolveSelectionRange(context.uiContext));
      }
      case "bilig.read_visible_range": {
        return await inspectWorkbookRange(context, resolveVisibleRange(context.uiContext));
      }
      case "bilig.inspect_cell": {
        const args = inspectCellToolArgsSchema.parse(request.arguments);
        return await inspectWorkbookCell(context, resolveInspectionTarget(context.uiContext, args));
      }
      case "bilig.find_formula_issues": {
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
      case "bilig.search_workbook": {
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
      case "bilig.trace_dependencies": {
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
      case "bilig.write_range": {
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
      case "bilig.clear_range": {
        const args = clearRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.range, MAX_MUTATION_RANGE_CELLS);
        return await stageCommandResult(context, {
          kind: "clearRange",
          range: args.range,
        });
      }
      case "bilig.format_range": {
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
      case "bilig.fill_range": {
        const args = transferRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.source, MAX_MUTATION_RANGE_CELLS);
        ensureRangeLimit(args.target, MAX_MUTATION_RANGE_CELLS);
        return await stageCommandResult(context, {
          kind: "fillRange",
          source: args.source,
          target: args.target,
        });
      }
      case "bilig.copy_range": {
        const args = transferRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.source, MAX_MUTATION_RANGE_CELLS);
        ensureRangeLimit(args.target, MAX_MUTATION_RANGE_CELLS);
        return await stageCommandResult(context, {
          kind: "copyRange",
          source: args.source,
          target: args.target,
        });
      }
      case "bilig.move_range": {
        const args = transferRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.source, MAX_MUTATION_RANGE_CELLS);
        ensureRangeLimit(args.target, MAX_MUTATION_RANGE_CELLS);
        return await stageCommandResult(context, {
          kind: "moveRange",
          source: args.source,
          target: args.target,
        });
      }
      case "bilig.create_sheet": {
        const args = sheetMutationToolArgsSchema.parse(request.arguments);
        return await stageCommandResult(context, {
          kind: "createSheet",
          name: args.name,
        });
      }
      case "bilig.rename_sheet": {
        const args = renameSheetToolArgsSchema.parse(request.arguments);
        return await stageCommandResult(context, {
          kind: "renameSheet",
          currentName: args.currentName,
          nextName: args.nextName,
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
