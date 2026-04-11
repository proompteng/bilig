import type { CodexThread, CodexThreadItem } from "@bilig/agent-api";
import { WORKBOOK_AGENT_TOOL_NAMES, renderWorkbookAgentSkillInstructions } from "@bilig/agent-api";
import type {
  WorkbookAgentSessionSnapshot,
  WorkbookAgentTimelineCitation,
  WorkbookAgentTimelineEntry,
} from "@bilig/contracts";
import { z } from "zod";

const workbookAgentUiContextSchema = z.object({
  selection: z.object({
    sheetName: z.string().min(1),
    address: z.string().min(1),
  }),
  viewport: z.object({
    rowStart: z.number().int().nonnegative(),
    rowEnd: z.number().int().nonnegative(),
    colStart: z.number().int().nonnegative(),
    colEnd: z.number().int().nonnegative(),
  }),
});

export const createSessionBodySchema = z.object({
  sessionId: z.string().min(1).optional(),
  threadId: z.string().min(1).optional(),
  scope: z.enum(["private", "shared"]).optional(),
  context: workbookAgentUiContextSchema.optional(),
});

export const updateContextBodySchema = z.object({
  context: workbookAgentUiContextSchema,
});

export const startTurnBodySchema = z.object({
  prompt: z.string().trim().min(1),
  context: workbookAgentUiContextSchema.optional(),
});

export const startWorkflowBodySchema = z
  .discriminatedUnion("workflowTemplate", [
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
  ])
  .and(
    z.object({
      context: workbookAgentUiContextSchema.optional(),
    }),
  );

export const reviewPendingBundleBodySchema = z.object({
  decision: z.enum(["approved", "rejected"]),
});

export function createWorkbookAgentBaseInstructions(): string {
  return [
    "You are the bilig workbook assistant embedded inside a spreadsheet product.",
    "Stay narrowly focused on inspecting and editing the active workbook.",
    "Use the provided bilig workbook tools and dynamic tools for spreadsheet work.",
    "Do not use filesystem, shell, web, connector, or unrelated tools.",
    renderWorkbookAgentSkillInstructions(),
  ].join(" ");
}

export function createWorkbookAgentDeveloperInstructions(): string {
  return [
    "Before changing cells you have not inspected, read the relevant workbook range first.",
    `Use ${WORKBOOK_AGENT_TOOL_NAMES.startWorkflow} with summarizeWorkbook, summarizeCurrentSheet, describeRecentChanges, findFormulaIssues, highlightFormulaIssues, highlightCurrentSheetOutliers, styleCurrentSheetHeaders, normalizeCurrentSheetHeaders, normalizeCurrentSheetNumberFormats, normalizeCurrentSheetWhitespace, fillCurrentSheetFormulasDown, traceSelectionDependencies, explainSelectionCell, searchWorkbookQuery, createCurrentSheetRollup, createCurrentSheetReviewTab, createSheet, renameCurrentSheet, hideCurrentRow, hideCurrentColumn, unhideCurrentRow, or unhideCurrentColumn when the request matches those built-in durable workflows and you want the result saved in the thread.`,
    `Use ${WORKBOOK_AGENT_TOOL_NAMES.readWorkbook} first when the user asks for workbook-wide structure, important sheets, or a starting summary and the built-in workflow is not the best fit.`,
    `When the user refers to the current cell, selection, or visible area, call ${WORKBOOK_AGENT_TOOL_NAMES.getContext}.`,
    `Prefer ${WORKBOOK_AGENT_TOOL_NAMES.readSelection}, ${WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange}, and ${WORKBOOK_AGENT_TOOL_NAMES.inspectCell} for context-native workbook analysis.`,
    `Use ${WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues}, ${WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook}, ${WORKBOOK_AGENT_TOOL_NAMES.traceDependencies}, and ${WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges} for warm-runtime workbook comprehension instead of broad guesswork.`,
    "All workbook writes must stage semantic preview bundles instead of applying immediately.",
    "Use the bundle-staging workbook tools to assemble one coherent preview per turn when the task is related.",
    "After staging workbook changes, summarize the preview and tell the user to review and apply it from the rail.",
    "If the requested action is outside the available bilig workbook tools, say exactly which workbook capability is missing instead of improvising.",
  ].join(" ");
}

function stringifyJson(value: unknown): string {
  return JSON.stringify(value, null, 2);
}

function formatToolContentItems(
  contentItems:
    | Array<
        | {
            type: "inputText";
            text: string;
          }
        | {
            type: "inputImage";
            imageUrl: string;
          }
      >
    | null
    | undefined,
): string | null {
  if (!contentItems || contentItems.length === 0) {
    return null;
  }
  return contentItems
    .map((item) => (item.type === "inputText" ? item.text : `[image] ${item.imageUrl}`))
    .join("\n");
}

function textFromUserContent(
  content: readonly {
    type: "text";
    text: string;
  }[],
): string {
  return content.map((item) => item.text).join("\n");
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isUserTextContentItem(value: unknown): value is { type: "text"; text: string } {
  return isRecord(value) && value["type"] === "text" && typeof value["text"] === "string";
}

function isUserMessageItem(
  item: CodexThreadItem,
): item is Extract<CodexThreadItem, { type: "userMessage" }> {
  return (
    item.type === "userMessage" &&
    Array.isArray(item.content) &&
    item.content.every((entry) => isUserTextContentItem(entry))
  );
}

function isAgentMessageItem(
  item: CodexThreadItem,
): item is Extract<CodexThreadItem, { type: "agentMessage" }> {
  return (
    item.type === "agentMessage" &&
    typeof item.text === "string" &&
    (item.phase === null || typeof item.phase === "string")
  );
}

function isPlanItem(item: CodexThreadItem): item is Extract<CodexThreadItem, { type: "plan" }> {
  return item.type === "plan" && typeof item.text === "string";
}

function isToolContentItem(item: unknown): item is
  | {
      type: "inputText";
      text: string;
    }
  | {
      type: "inputImage";
      imageUrl: string;
    } {
  return (
    typeof item === "object" &&
    item !== null &&
    (("type" in item &&
      item.type === "inputText" &&
      "text" in item &&
      typeof item.text === "string") ||
      ("type" in item &&
        item.type === "inputImage" &&
        "imageUrl" in item &&
        typeof item.imageUrl === "string"))
  );
}

function isDynamicToolCallItem(
  item: CodexThreadItem,
): item is Extract<CodexThreadItem, { type: "dynamicToolCall" }> {
  return (
    item.type === "dynamicToolCall" &&
    typeof item.tool === "string" &&
    (item.status === "inProgress" || item.status === "completed" || item.status === "failed") &&
    (item.contentItems === null ||
      (Array.isArray(item.contentItems) &&
        item.contentItems.every((entry) => isToolContentItem(entry)))) &&
    (item.success === null || typeof item.success === "boolean")
  );
}

export function createSystemEntry(
  id: string,
  turnId: string | null,
  text: string,
  citations: readonly WorkbookAgentTimelineCitation[] = [],
): WorkbookAgentTimelineEntry {
  return {
    id,
    kind: "system",
    turnId,
    text,
    phase: null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [...citations],
  };
}

export function mapThreadItemToEntry(
  item: CodexThreadItem,
  turnId: string | null,
): WorkbookAgentTimelineEntry {
  if (isUserMessageItem(item)) {
    return {
      id: item.id,
      kind: "user",
      turnId,
      text: textFromUserContent(item.content),
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
      citations: [],
    };
  }

  if (isAgentMessageItem(item)) {
    return {
      id: item.id,
      kind: "assistant",
      turnId,
      text: item.text,
      phase: item.phase,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
      citations: [],
    };
  }

  if (isPlanItem(item)) {
    return {
      id: item.id,
      kind: "plan",
      turnId,
      text: item.text,
      phase: null,
      toolName: null,
      toolStatus: null,
      argumentsText: null,
      outputText: null,
      success: null,
      citations: [],
    };
  }

  if (isDynamicToolCallItem(item)) {
    return {
      id: item.id,
      kind: "tool",
      turnId,
      text: null,
      phase: null,
      toolName: item.tool,
      toolStatus: item.status,
      argumentsText: stringifyJson(item.arguments),
      outputText: formatToolContentItems(item.contentItems),
      success: item.success,
      citations: [],
    };
  }

  return createSystemEntry(item.id, turnId, `Codex emitted ${item.type}.`);
}

type MutableWorkbookAgentSessionSnapshot = {
  -readonly [Key in keyof WorkbookAgentSessionSnapshot]: Key extends "entries"
    ? WorkbookAgentTimelineEntry[]
    : WorkbookAgentSessionSnapshot[Key];
};

export function cloneSnapshot(
  snapshot: MutableWorkbookAgentSessionSnapshot,
): WorkbookAgentSessionSnapshot {
  return {
    ...snapshot,
    entries: snapshot.entries.map((entry) => ({ ...entry })),
    ...(snapshot.context ? { context: structuredClone(snapshot.context) } : { context: null }),
    pendingBundle: snapshot.pendingBundle ? structuredClone(snapshot.pendingBundle) : null,
    executionRecords: snapshot.executionRecords.map((record) => structuredClone(record)),
    workflowRuns: snapshot.workflowRuns.map((run) => structuredClone(run)),
  };
}

export function buildEntriesFromThread(thread: CodexThread): WorkbookAgentTimelineEntry[] {
  const entries: WorkbookAgentTimelineEntry[] = [];
  for (const turn of thread.turns) {
    for (const item of turn.items) {
      entries.push(mapThreadItemToEntry(item, turn.id));
    }
  }
  return entries;
}
