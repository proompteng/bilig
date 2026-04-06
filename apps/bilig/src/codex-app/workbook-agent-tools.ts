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
import type { WorkbookAgentCommand, WorkbookAgentCommandBundle } from "@bilig/agent-api";
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
import type {
  CodexDynamicToolCallRequest,
  CodexDynamicToolCallResult,
  CodexDynamicToolSpec,
  JsonValue,
} from "./codex-app-server-types.js";

const MAX_TOOL_RANGE_CELLS = 400;

const writeCellInputSchema = z.union([
  z.string(),
  z.number(),
  z.boolean(),
  z.null(),
  z.object({
    value: z.union([z.string(), z.number(), z.boolean(), z.null()]),
  }),
  z.object({
    formula: z.string().min(1),
  }),
]);

const readRangeToolArgsSchema = z.object({
  sheetName: z.string().min(1),
  startAddress: z.string().min(1),
  endAddress: z.string().min(1),
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
const transferRangeToolArgsSchema = rangeMutationArgsSchema.pick({
  source: true,
  target: true,
});
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

function ensureRangeLimit(range: CellRangeRef): void {
  const count = countRangeCells(range);
  if (count > MAX_TOOL_RANGE_CELLS) {
    throw new Error(
      `Range ${range.sheetName}!${range.startAddress}:${range.endAddress} has ${String(count)} cells; tool limit is ${String(MAX_TOOL_RANGE_CELLS)} cells per call`,
    );
  }
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
        const range = normalizeRange({
          sheetName: args.sheetName,
          startAddress: args.startAddress,
          endAddress: args.endAddress,
        });
        ensureRangeLimit(range);
        const result = await context.zeroSyncService.inspectWorkbook(
          context.documentId,
          (runtime) => {
            const rows: JsonValue[] = [];
            for (let row = range.startRow; row <= range.endRow; row += 1) {
              const rowEntries: JsonValue[] = [];
              for (let col = range.startCol; col <= range.endCol; col += 1) {
                const cell = runtime.engine.getCell(range.sheetName, formatAddress(row, col));
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
                sheetName: range.sheetName,
                startAddress: range.startAddress,
                endAddress: range.endAddress,
              },
              rows,
            };
          },
        );
        return textToolResult(stringifyJson(result));
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
        ensureRangeLimit({
          sheetName: args.sheetName,
          startAddress: args.startAddress,
          endAddress,
        });
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
        ensureRangeLimit(args.range);
        return await stageCommandResult(context, {
          kind: "clearRange",
          range: args.range,
        });
      }
      case "bilig.format_range": {
        const args = formatRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.range);
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
        ensureRangeLimit(args.source);
        ensureRangeLimit(args.target);
        return await stageCommandResult(context, {
          kind: "fillRange",
          source: args.source,
          target: args.target,
        });
      }
      case "bilig.copy_range": {
        const args = transferRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.source);
        ensureRangeLimit(args.target);
        return await stageCommandResult(context, {
          kind: "copyRange",
          source: args.source,
          target: args.target,
        });
      }
      case "bilig.move_range": {
        const args = transferRangeToolArgsSchema.parse(request.arguments);
        ensureRangeLimit(args.source);
        ensureRangeLimit(args.target);
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
