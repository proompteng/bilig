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
import type {
  CellRangeRef,
  LiteralInput,
  WorkbookDataValidationRuleSnapshot,
  WorkbookDataValidationSnapshot,
} from "@bilig/protocol";
import type { WorkbookAgentUiContext } from "@bilig/contracts";
import { z } from "zod";
import type { SessionIdentity } from "../http/session.js";
import type { ZeroSyncService } from "../zero/service.js";
import {
  rangeOrSelectorJsonSchema,
  rangeOrSelectorSchema,
  resolveRangeOrSelectorRequest,
} from "./workbook-agent-selector-tooling.js";
import type { WorkbookRuntime } from "../workbook-runtime/runtime-manager.js";

const validationListSourceSchema = z.union([
  z.object({
    kind: z.literal("named-range"),
    name: z.string().trim().min(1),
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
]);

const literalValueSchema = z.union([z.string(), z.number(), z.boolean(), z.null()]);

const validationRuleSchema = z.union([
  z
    .object({
      kind: z.literal("list"),
      values: z.array(literalValueSchema).min(1).optional(),
      source: validationListSourceSchema.optional(),
    })
    .refine((value) => (value.values ? 1 : 0) + (value.source ? 1 : 0) === 1, {
      message: "Provide exactly one of values or source for list validation rules",
    }),
  z.object({
    kind: z.literal("checkbox"),
    checkedValue: literalValueSchema.optional(),
    uncheckedValue: literalValueSchema.optional(),
  }),
  z
    .object({
      kind: z.enum(["whole", "decimal", "date", "time", "textLength"]),
      operator: z.enum([
        "between",
        "notBetween",
        "equal",
        "notEqual",
        "greaterThan",
        "greaterThanOrEqual",
        "lessThan",
        "lessThanOrEqual",
      ]),
      values: z.array(literalValueSchema).min(1).max(2),
    })
    .refine(
      (value) =>
        ["between", "notBetween"].includes(value.operator)
          ? value.values.length === 2
          : value.values.length === 1,
      {
        message: "between and notBetween require two values; other operators require one value",
      },
    ),
]);

const listDataValidationRulesArgsSchema = z
  .object({
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) <= 1, {
    message: "Provide at most one of range or selector",
  });

const dataValidationMutationArgsSchema = z
  .object({
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
    rule: validationRuleSchema,
    allowBlank: z.boolean().optional(),
    showDropdown: z.boolean().optional(),
    promptTitle: z.string().trim().min(1).optional(),
    promptMessage: z.string().trim().min(1).optional(),
    errorStyle: z.enum(["stop", "warning", "information"]).optional(),
    errorTitle: z.string().trim().min(1).optional(),
    errorMessage: z.string().trim().min(1).optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: "Provide exactly one of range or selector",
  });

const removeDataValidationArgsSchema = rangeOrSelectorSchema;

export const workbookAgentValidationToolSpecs = [
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.listDataValidationRules,
    description:
      "List workbook data validation rules. Optionally filter the result to a specific explicit range or semantic selector.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.createDataValidation,
    description:
      "Create a data validation rule on an explicit range or semantic selector, including dropdown, checkbox, or scalar validation constraints.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["rule"],
      properties: {
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        rule: { type: "object" },
        allowBlank: { type: "boolean" },
        showDropdown: { type: "boolean" },
        promptTitle: { type: "string" },
        promptMessage: { type: "string" },
        errorStyle: { type: "string", enum: ["stop", "warning", "information"] },
        errorTitle: { type: "string" },
        errorMessage: { type: "string" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.updateDataValidation,
    description:
      "Update a data validation rule on an explicit range or semantic selector. This replaces the rule for that exact target range.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["rule"],
      properties: {
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        rule: { type: "object" },
        allowBlank: { type: "boolean" },
        showDropdown: { type: "boolean" },
        promptTitle: { type: "string" },
        promptMessage: { type: "string" },
        errorStyle: { type: "string", enum: ["stop", "warning", "information"] },
        errorTitle: { type: "string" },
        errorMessage: { type: "string" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.removeDataValidation,
    description: "Remove a data validation rule from an explicit range or semantic selector.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
      },
    },
  },
] satisfies readonly CodexDynamicToolSpec[];

export interface WorkbookAgentValidationToolContext {
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
    contentItems: [{ type: "inputText", text }],
  };
}

async function stageCommandResult(
  context: WorkbookAgentValidationToolContext,
  command: WorkbookAgentCommand,
): Promise<CodexDynamicToolCallResult> {
  const result = await context.stageCommand(command);
  const normalized =
    "bundle" in result
      ? result
      : { bundle: result, executionRecord: null, disposition: "reviewQueued" as const };
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

function normalizeRange(range: CellRangeRef): CellRangeRef {
  const start = parseCellAddress(range.startAddress, range.sheetName);
  const end = parseCellAddress(range.endAddress, range.sheetName);
  return {
    sheetName: range.sheetName,
    startAddress: formatAddress(Math.min(start.row, end.row), Math.min(start.col, end.col)),
    endAddress: formatAddress(Math.max(start.row, end.row), Math.max(start.col, end.col)),
  };
}

function rangesIntersect(left: CellRangeRef, right: CellRangeRef): boolean {
  const leftBounds = normalizeRange(left);
  const rightBounds = normalizeRange(right);
  const leftStart = parseCellAddress(leftBounds.startAddress, leftBounds.sheetName);
  const leftEnd = parseCellAddress(leftBounds.endAddress, leftBounds.sheetName);
  const rightStart = parseCellAddress(rightBounds.startAddress, rightBounds.sheetName);
  const rightEnd = parseCellAddress(rightBounds.endAddress, rightBounds.sheetName);
  return !(
    leftBounds.sheetName !== rightBounds.sheetName ||
    leftEnd.row < rightStart.row ||
    rightEnd.row < leftStart.row ||
    leftEnd.col < rightStart.col ||
    rightEnd.col < leftStart.col
  );
}

function listWorkbookDataValidations(runtime: WorkbookRuntime): WorkbookDataValidationSnapshot[] {
  return runtime.engine
    .exportSnapshot()
    .sheets.flatMap((sheet) => runtime.engine.getDataValidations(sheet.name))
    .map((validation) => structuredClone(validation));
}

function normalizeValidationRule(
  rule: z.infer<typeof validationRuleSchema>,
): WorkbookDataValidationRuleSnapshot {
  switch (rule.kind) {
    case "list": {
      const normalized: Extract<WorkbookDataValidationRuleSnapshot, { kind: "list" }> = {
        kind: "list",
      };
      if (rule.values) {
        normalized.values = [...rule.values] as LiteralInput[];
      }
      if (rule.source) {
        normalized.source = structuredClone(rule.source);
      }
      return normalized;
    }
    case "checkbox": {
      const normalized: Extract<WorkbookDataValidationRuleSnapshot, { kind: "checkbox" }> = {
        kind: "checkbox",
      };
      if (rule.checkedValue !== undefined) {
        normalized.checkedValue = rule.checkedValue;
      }
      if (rule.uncheckedValue !== undefined) {
        normalized.uncheckedValue = rule.uncheckedValue;
      }
      return normalized;
    }
    case "whole":
    case "decimal":
    case "date":
    case "time":
    case "textLength":
      return {
        kind: rule.kind,
        operator: rule.operator,
        values: [...rule.values] as LiteralInput[],
      };
  }
}

function toValidationRecord(
  args: z.infer<typeof dataValidationMutationArgsSchema> & {
    range: CellRangeRef;
  },
): WorkbookDataValidationSnapshot {
  return {
    range: normalizeRange(args.range),
    rule: normalizeValidationRule(args.rule),
    ...(args.allowBlank !== undefined ? { allowBlank: args.allowBlank } : {}),
    ...(args.showDropdown !== undefined ? { showDropdown: args.showDropdown } : {}),
    ...(args.promptTitle ? { promptTitle: args.promptTitle } : {}),
    ...(args.promptMessage ? { promptMessage: args.promptMessage } : {}),
    ...(args.errorStyle ? { errorStyle: args.errorStyle } : {}),
    ...(args.errorTitle ? { errorTitle: args.errorTitle } : {}),
    ...(args.errorMessage ? { errorMessage: args.errorMessage } : {}),
  };
}

export async function handleWorkbookAgentValidationToolCall(
  context: WorkbookAgentValidationToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult | null> {
  const normalizedTool = normalizeWorkbookAgentToolName(request.tool);
  switch (normalizedTool) {
    case WORKBOOK_AGENT_TOOL_NAMES.listDataValidationRules: {
      const args = listDataValidationRulesArgsSchema.parse(request.arguments);
      const payload = await context.zeroSyncService.inspectWorkbook(
        context.documentId,
        (runtime) => {
          const validations = listWorkbookDataValidations(runtime);
          const filtered =
            args.range || args.selector
              ? (() => {
                  const resolved = resolveRangeOrSelectorRequest({
                    runtime,
                    args: {
                      ...(args.range ? { range: args.range } : {}),
                      ...(args.selector ? { selector: args.selector } : {}),
                    },
                    uiContext: context.uiContext,
                  });
                  return validations.filter((validation) =>
                    rangesIntersect(validation.range, resolved.range),
                  );
                })()
              : validations;
          return {
            documentId: context.documentId,
            validationCount: filtered.length,
            validations: filtered,
          };
        },
      );
      return textToolResult(stringifyJson(payload));
    }
    case WORKBOOK_AGENT_TOOL_NAMES.createDataValidation:
    case WORKBOOK_AGENT_TOOL_NAMES.updateDataValidation: {
      const args = dataValidationMutationArgsSchema.parse(request.arguments);
      const validation = await context.zeroSyncService.inspectWorkbook(
        context.documentId,
        (runtime) => {
          const resolved = resolveRangeOrSelectorRequest({
            runtime,
            args: {
              ...(args.range ? { range: args.range } : {}),
              ...(args.selector ? { selector: args.selector } : {}),
            },
            uiContext: context.uiContext,
          });
          return toValidationRecord({ ...args, range: resolved.range });
        },
      );
      return await stageCommandResult(context, {
        kind: "setDataValidation",
        validation,
      });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.removeDataValidation: {
      const args = removeDataValidationArgsSchema.parse(request.arguments);
      const range = await context.zeroSyncService.inspectWorkbook(
        context.documentId,
        (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args,
            uiContext: context.uiContext,
          }).range,
      );
      return await stageCommandResult(context, {
        kind: "clearDataValidation",
        range: normalizeRange(range),
      });
    }
    default:
      return null;
  }
}
