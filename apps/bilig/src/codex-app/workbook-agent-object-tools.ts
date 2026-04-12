import { formatAddress, parseCellAddress } from "@bilig/formula";
import {
  WORKBOOK_AGENT_TOOL_NAMES,
  normalizeWorkbookAgentToolName,
  type CodexDynamicToolCallRequest,
  type CodexDynamicToolCallResult,
  type CodexDynamicToolSpec,
  type WorkbookAgentCommand,
  type WorkbookAgentCommandBundle,
  type WorkbookAgentExecutionRecord,
} from "@bilig/agent-api";
import {
  ValueTag,
  type WorkbookDefinedNameValueSnapshot,
  type WorkbookPivotSnapshot,
  type WorkbookTableSnapshot,
} from "@bilig/protocol";
import type { WorkbookAgentUiContext } from "@bilig/contracts";
import { z } from "zod";
import type { SessionIdentity } from "../http/session.js";
import type { ZeroSyncService } from "../zero/service.js";
import {
  rangeOrSelectorJsonSchema,
  rangeOrSelectorSchema,
  resolveRangeOrSelectorRequest,
  workbookSemanticSelectorJsonSchema,
} from "./workbook-agent-selector-tooling.js";
import {
  resolveWorkbookSelector,
  workbookSemanticSelectorSchema,
  type WorkbookSemanticSelector,
} from "./workbook-selector-resolver.js";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";

const literalInputSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

const definedNameValueSchema: z.ZodType<WorkbookDefinedNameValueSnapshot> = z.union([
  literalInputSchema,
  z.object({
    kind: z.literal("scalar"),
    value: literalInputSchema,
  }),
  z.object({
    kind: z.literal("cell-ref"),
    sheetName: z.string().trim().min(1),
    address: z.string().trim().min(1),
  }),
  z.object({
    kind: z.literal("range-ref"),
    sheetName: z.string().trim().min(1),
    startAddress: z.string().trim().min(1),
    endAddress: z.string().trim().min(1),
  }),
  z.object({
    kind: z.literal("structured-ref"),
    tableName: z.string().trim().min(1),
    columnName: z.string().trim().min(1),
  }),
  z.object({
    kind: z.literal("formula"),
    formula: z.string().trim().min(1),
  }),
]);

const namedRangeMutationArgsSchema = z
  .object({
    name: z.string().trim().min(1),
    value: definedNameValueSchema.optional(),
    selector: workbookSemanticSelectorSchema.optional(),
  })
  .refine((value) => (value.value ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: "Provide exactly one of value or selector",
  });

const deleteNamedRangeArgsSchema = z.object({
  name: z.string().trim().min(1),
});

const tableMutationArgsSchema = z
  .object({
    name: z.string().trim().min(1),
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: workbookSemanticSelectorSchema.optional(),
    headerRow: z.boolean().optional(),
    totalsRow: z.boolean().optional(),
    columnNames: z.array(z.string().trim().min(1)).optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: "Provide exactly one of range or selector",
  });

const deleteTableArgsSchema = z.object({
  name: z.string().trim().min(1),
});

const pivotValueSchema = z.object({
  sourceColumn: z.string().trim().min(1),
  summarizeBy: z.enum(["sum", "count"]),
  outputLabel: z.string().trim().min(1).optional(),
});
const pivotValuesSchema = z.array(pivotValueSchema).min(1);

const pivotMutationArgsSchema = z
  .object({
    name: z.string().trim().min(1),
    sheetName: z.string().trim().min(1),
    address: z.string().trim().min(1),
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: workbookSemanticSelectorSchema.optional(),
    groupBy: z.array(z.string().trim().min(1)),
    values: pivotValuesSchema,
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: "Provide exactly one of range or selector",
  });

const deletePivotArgsSchema = z.union([
  z.object({
    name: z.string().trim().min(1),
  }),
  z.object({
    sheetName: z.string().trim().min(1),
    address: z.string().trim().min(1),
  }),
]);

export const workbookAgentObjectToolSpecs = [
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.listPivots,
    description: "List workbook pivot tables with source ranges, grouping fields, and output size.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {},
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.createNamedRange,
    description:
      "Create a workbook named range or named reference from either an explicit value payload or a semantic selector.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["name"],
      properties: {
        name: { type: "string" },
        value: { type: "object" },
        selector: workbookSemanticSelectorJsonSchema,
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.updateNamedRange,
    description:
      "Update a workbook named range or named reference from either an explicit value payload or a semantic selector.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["name"],
      properties: {
        name: { type: "string" },
        value: { type: "object" },
        selector: workbookSemanticSelectorJsonSchema,
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.deleteNamedRange,
    description: "Delete a workbook named range or named reference.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.createTable,
    description:
      "Create a workbook table from an explicit range or semantic selector. Column names default from the header row when available.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["name"],
      properties: {
        name: { type: "string" },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: workbookSemanticSelectorJsonSchema,
        headerRow: { type: "boolean" },
        totalsRow: { type: "boolean" },
        columnNames: {
          type: "array",
          items: { type: "string" },
        },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.resizeTable,
    description:
      "Resize or update a workbook table from an explicit range or semantic selector. Column names default from the header row when available.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["name"],
      properties: {
        name: { type: "string" },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: workbookSemanticSelectorJsonSchema,
        headerRow: { type: "boolean" },
        totalsRow: { type: "boolean" },
        columnNames: {
          type: "array",
          items: { type: "string" },
        },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.deleteTable,
    description: "Delete a workbook table by name.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.createPivotTable,
    description:
      "Create a pivot table from an explicit source range or semantic selector, with grouping fields and aggregate value definitions.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["name", "sheetName", "address", "groupBy", "values"],
      properties: {
        name: { type: "string" },
        sheetName: { type: "string" },
        address: { type: "string" },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: workbookSemanticSelectorJsonSchema,
        groupBy: {
          type: "array",
          items: { type: "string" },
        },
        values: {
          type: "array",
          items: {
            type: "object",
            additionalProperties: false,
            required: ["sourceColumn", "summarizeBy"],
            properties: {
              sourceColumn: { type: "string" },
              summarizeBy: { type: "string", enum: ["sum", "count"] },
              outputLabel: { type: "string" },
            },
          },
        },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.updatePivotTable,
    description:
      "Update a pivot table from an explicit source range or semantic selector, with grouping fields and aggregate value definitions.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["name", "sheetName", "address", "groupBy", "values"],
      properties: {
        name: { type: "string" },
        sheetName: { type: "string" },
        address: { type: "string" },
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: workbookSemanticSelectorJsonSchema,
        groupBy: {
          type: "array",
          items: { type: "string" },
        },
        values: {
          type: "array",
          items: {
            type: "object",
            additionalProperties: false,
            required: ["sourceColumn", "summarizeBy"],
            properties: {
              sourceColumn: { type: "string" },
              summarizeBy: { type: "string", enum: ["sum", "count"] },
              outputLabel: { type: "string" },
            },
          },
        },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.deletePivotTable,
    description: "Delete a pivot table by either name or its anchor sheet/address.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        name: { type: "string" },
        sheetName: { type: "string" },
        address: { type: "string" },
      },
    },
  },
] satisfies readonly CodexDynamicToolSpec[];

export interface WorkbookAgentObjectToolContext {
  readonly documentId: string;
  readonly session: SessionIdentity;
  readonly uiContext: WorkbookAgentUiContext | null;
  readonly zeroSyncService: ZeroSyncService;
  readonly stageCommand: (command: WorkbookAgentCommand) => Promise<
    | WorkbookAgentCommandBundle
    | {
        readonly bundle: WorkbookAgentCommandBundle;
        readonly executionRecord: WorkbookAgentExecutionRecord | null;
        readonly disposition?: "queuedForTurnApply" | "reviewQueued";
      }
  >;
}

function stringifyJson(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

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

async function stageCommandResult(
  context: WorkbookAgentObjectToolContext,
  command: WorkbookAgentCommand,
): Promise<CodexDynamicToolCallResult> {
  const result = await context.stageCommand(command);
  const normalized =
    "bundle" in result
      ? result
      : { bundle: result, executionRecord: null, disposition: "reviewQueued" };
  const bundle = normalized.bundle;
  if (normalized.executionRecord) {
    return textToolResult(
      stringifyJson({
        applied: true,
        staged: false,
        reviewQueued: false,
        bundleId: bundle.id,
        summary: `Applied workbook change set at revision r${String(normalized.executionRecord.appliedRevision)}: ${normalized.executionRecord.summary}`,
        revision: normalized.executionRecord.appliedRevision,
        scope: normalized.executionRecord.scope,
        riskClass: normalized.executionRecord.riskClass,
        estimatedAffectedCells: bundle.estimatedAffectedCells,
        affectedRanges: bundle.affectedRanges,
      }),
    );
  }
  if (normalized.disposition === "queuedForTurnApply") {
    return textToolResult(
      stringifyJson({
        applied: false,
        staged: true,
        reviewQueued: false,
        queuedForTurnApply: true,
        bundleId: bundle.id,
        summary: `Queued workbook change set for turn apply: ${bundle.summary}`,
        scope: bundle.scope,
        riskClass: bundle.riskClass,
        estimatedAffectedCells: bundle.estimatedAffectedCells,
        affectedRanges: bundle.affectedRanges,
      }),
    );
  }
  return textToolResult(
    stringifyJson({
      applied: false,
      staged: true,
      reviewQueued: true,
      bundleId: bundle.id,
      summary: `Prepared workbook review item: ${bundle.summary}`,
      scope: bundle.scope,
      riskClass: bundle.riskClass,
      estimatedAffectedCells: bundle.estimatedAffectedCells,
      affectedRanges: bundle.affectedRanges,
    }),
  );
}

function normalizeFormulaText(formula: string): string {
  return formula.startsWith("=") ? formula : `=${formula}`;
}

function definedNameValueFromSelector(
  selector: WorkbookSemanticSelector,
  runtime: WorkbookRuntime,
  uiContext: WorkbookAgentUiContext | null,
): WorkbookDefinedNameValueSnapshot {
  const resolution = resolveWorkbookSelector({
    runtime,
    selector,
    uiContext,
  });
  switch (selector.kind) {
    case "tableColumn":
      return {
        kind: "structured-ref",
        tableName: selector.table,
        columnName: selector.column,
      };
    case "namedRange":
      if (resolution.namedRange) {
        return structuredClone(resolution.namedRange.value);
      }
      break;
    case "a1Range":
    case "currentRegion":
    case "currentSelection":
    case "columnQuery":
    case "rowQuery":
    case "table":
    case "visibleRows":
      break;
    default: {
      const exhaustive: never = selector;
      return exhaustive;
    }
  }
  const range = resolution.derivedA1Ranges[0];
  if (!range) {
    throw new Error(`Selector ${resolution.displayLabel} does not resolve to a workbook range`);
  }
  if (range.startAddress === range.endAddress) {
    return {
      kind: "cell-ref",
      sheetName: range.sheetName,
      address: range.startAddress,
    };
  }
  return {
    kind: "range-ref",
    sheetName: range.sheetName,
    startAddress: range.startAddress,
    endAddress: range.endAddress,
  };
}

function formatColumnLabel(index: number): string {
  return formatAddress(0, index).replace(/\d+/gu, "");
}

function defaultColumnNames(input: {
  runtime: WorkbookRuntime;
  range: { sheetName: string; startAddress: string; endAddress: string };
  headerRow: boolean;
}): string[] {
  const start = parseCellAddress(input.range.startAddress, input.range.sheetName);
  const end = parseCellAddress(input.range.endAddress, input.range.sheetName);
  const names: string[] = [];
  for (let col = start.col; col <= end.col; col += 1) {
    if (input.headerRow) {
      const cell = input.runtime.engine.getCell(
        input.range.sheetName,
        formatAddress(start.row, col),
      );
      const candidate =
        typeof cell.input === "string"
          ? cell.input.trim()
          : typeof cell.formula === "string"
            ? normalizeFormulaText(cell.formula)
            : cell.value.tag === ValueTag.Number && typeof cell.value.value === "number"
              ? String(cell.value.value)
              : cell.value.tag === ValueTag.Boolean && typeof cell.value.value === "boolean"
                ? String(cell.value.value)
                : cell.value.tag === ValueTag.String && typeof cell.value.value === "string"
                  ? cell.value.value.trim()
                  : "";
      names.push(candidate.length > 0 ? candidate : `Column ${String(col - start.col + 1)}`);
      continue;
    }
    names.push(`Column ${formatColumnLabel(col)}`);
  }
  return names;
}

function normalizeTableCommand(input: {
  name: string;
  range: { sheetName: string; startAddress: string; endAddress: string };
  headerRow?: boolean;
  totalsRow?: boolean;
  columnNames?: string[];
  existingTable?: WorkbookTableSnapshot | null;
  runtime: WorkbookRuntime;
}): WorkbookTableSnapshot {
  const start = parseCellAddress(input.range.startAddress, input.range.sheetName);
  const end = parseCellAddress(input.range.endAddress, input.range.sheetName);
  const width = end.col - start.col + 1;
  const columnNames =
    input.columnNames ??
    input.existingTable?.columnNames ??
    defaultColumnNames({
      runtime: input.runtime,
      range: input.range,
      headerRow: input.headerRow ?? input.existingTable?.headerRow ?? true,
    });
  if (columnNames.length !== width) {
    throw new Error(
      `Table ${input.name} spans ${String(width)} columns, but received ${String(columnNames.length)} column names`,
    );
  }
  return {
    name: input.name,
    sheetName: input.range.sheetName,
    startAddress: input.range.startAddress,
    endAddress: input.range.endAddress,
    columnNames,
    headerRow: input.headerRow ?? input.existingTable?.headerRow ?? true,
    totalsRow: input.totalsRow ?? input.existingTable?.totalsRow ?? false,
  };
}

async function resolvePivotDeleteTarget(
  runtime: WorkbookRuntime,
  args: z.infer<typeof deletePivotArgsSchema>,
): Promise<{ sheetName: string; address: string }> {
  if ("sheetName" in args) {
    return {
      sheetName: args.sheetName,
      address: args.address,
    };
  }
  const matches = runtime.engine
    .getPivotTables()
    .filter((pivot) => pivot.name.trim().toUpperCase() === args.name.trim().toUpperCase());
  if (matches.length === 0) {
    throw new Error(`Pivot ${args.name} does not exist`);
  }
  if (matches.length > 1) {
    throw new Error(`Pivot ${args.name} is ambiguous; specify sheetName and address`);
  }
  return {
    sheetName: matches[0]!.sheetName,
    address: matches[0]!.address,
  };
}

function listWorkbookPivots(runtime: WorkbookRuntime): readonly WorkbookPivotSnapshot[] {
  return runtime.engine.getPivotTables().map((pivot) => structuredClone(pivot));
}

export async function handleWorkbookAgentObjectToolCall(
  context: WorkbookAgentObjectToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult | null> {
  const normalizedTool = normalizeWorkbookAgentToolName(request.tool);
  switch (normalizedTool) {
    case WORKBOOK_AGENT_TOOL_NAMES.listPivots: {
      const payload = await context.zeroSyncService.inspectWorkbook(
        context.documentId,
        (runtime) => ({
          documentId: context.documentId,
          pivotCount: runtime.engine.getPivotTables().length,
          pivots: listWorkbookPivots(runtime),
        }),
      );
      return textToolResult(stringifyJson(payload));
    }
    case WORKBOOK_AGENT_TOOL_NAMES.createNamedRange:
    case WORKBOOK_AGENT_TOOL_NAMES.updateNamedRange: {
      const args = namedRangeMutationArgsSchema.parse(request.arguments);
      const value = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
        args.selector
          ? definedNameValueFromSelector(args.selector, runtime, context.uiContext)
          : args.value!,
      );
      return await stageCommandResult(context, {
        kind: "upsertDefinedName",
        name: args.name,
        value:
          typeof value === "object" && value !== null && "kind" in value && value.kind === "formula"
            ? { ...value, formula: normalizeFormulaText(value.formula) }
            : typeof value === "string" && value.startsWith("=")
              ? normalizeFormulaText(value)
              : value,
      });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deleteNamedRange: {
      const args = deleteNamedRangeArgsSchema.parse(request.arguments);
      return await stageCommandResult(context, {
        kind: "deleteDefinedName",
        name: args.name,
      });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.createTable:
    case WORKBOOK_AGENT_TOOL_NAMES.resizeTable: {
      const args = tableMutationArgsSchema.parse(request.arguments);
      const table = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const existingTable =
          args.selector?.kind === "table"
            ? resolveWorkbookSelector({
                runtime,
                selector: args.selector,
                uiContext: context.uiContext,
              }).table
            : null;
        const resolved = resolveRangeOrSelectorRequest({
          runtime,
          args: {
            ...(args.range ? { range: args.range } : {}),
            ...(args.selector ? { selector: args.selector } : {}),
          },
          uiContext: context.uiContext,
        });
        return normalizeTableCommand({
          name: args.name,
          range: resolved.range,
          ...(args.headerRow !== undefined ? { headerRow: args.headerRow } : {}),
          ...(args.totalsRow !== undefined ? { totalsRow: args.totalsRow } : {}),
          ...(args.columnNames ? { columnNames: args.columnNames } : {}),
          ...(existingTable ? { existingTable } : {}),
          runtime,
        });
      });
      return await stageCommandResult(context, {
        kind: "upsertTable",
        table,
      });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deleteTable: {
      const args = deleteTableArgsSchema.parse(request.arguments);
      return await stageCommandResult(context, {
        kind: "deleteTable",
        name: args.name,
      });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.createPivotTable:
    case WORKBOOK_AGENT_TOOL_NAMES.updatePivotTable: {
      const args = pivotMutationArgsSchema.parse(request.arguments);
      const pivot = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const resolved = resolveRangeOrSelectorRequest({
          runtime,
          args: {
            ...(args.range ? { range: args.range } : {}),
            ...(args.selector ? { selector: args.selector } : {}),
          },
          uiContext: context.uiContext,
        });
        const normalizedRange = resolved.range;
        return {
          name: args.name,
          sheetName: args.sheetName,
          address: args.address,
          source: normalizedRange,
          groupBy: [...args.groupBy],
          values: args.values.map((value) => ({
            sourceColumn: value.sourceColumn,
            summarizeBy: value.summarizeBy,
            ...(value.outputLabel ? { outputLabel: value.outputLabel } : {}),
          })),
          rows: 1,
          cols: Math.max(args.groupBy.length + args.values.length, 1),
        } satisfies WorkbookPivotSnapshot;
      });
      return await stageCommandResult(context, {
        kind: "upsertPivotTable",
        pivot,
      });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deletePivotTable: {
      const args = deletePivotArgsSchema.parse(request.arguments);
      const target = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
        resolvePivotDeleteTarget(runtime, args),
      );
      return await stageCommandResult(context, {
        kind: "deletePivotTable",
        sheetName: target.sheetName,
        address: target.address,
      });
    }
    default:
      return null;
  }
}
