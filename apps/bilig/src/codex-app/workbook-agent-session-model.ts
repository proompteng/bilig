import {
  normalizeWorkbookAgentToolName,
  type CodexThread,
  type CodexThreadItem,
} from "@bilig/agent-api";
import type {
  WorkbookAgentSessionSnapshot,
  WorkbookAgentTextEntryKind,
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
  executionPolicy: z.enum(["autoApplySafe", "autoApplyAll", "ownerReview"]).optional(),
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
      workflowTemplate: z.literal("repairFormulaIssues"),
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
    "You are bilig's workbook assistant inside a spreadsheet product.",
    "Help with the active workbook only.",
    "Use only the provided workbook tools.",
    "If the request needs a capability the tools do not provide, say what is missing.",
  ].join(" ");
}

export function createWorkbookAgentDeveloperInstructions(): string {
  return [
    "Inspect before you edit unfamiliar cells or ranges.",
    "Prefer the smallest workbook tool that matches the request.",
    "When the request refers to the current cell, selection, or visible area, use the browser workbook context tools first.",
    "Use workbook or range reads for workbook-wide structure or unseen regions instead of guessing.",
    "Range and cell reads expose formatting metadata, including fill/background, font, alignment, borders, and number format when present.",
    "Use the workflow tool only for built-in multi-step or durable tasks.",
    "Use direct structural sheet tools for one-step sheet edits that should happen immediately.",
    "Apply workbook changes directly when the session policy allows it.",
    "When the session policy routes a change set to owner review, summarize the prepared review item in workbook terms.",
    "Do not use non-workbook tools or invent unsupported capabilities.",
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

function extractReasoningFragments(value: unknown): string[] {
  if (typeof value === "string" && value.trim().length > 0) {
    return [value];
  }
  if (Array.isArray(value)) {
    return value.flatMap((entry) => extractReasoningFragments(entry));
  }
  if (!isRecord(value)) {
    return [];
  }
  return [
    ...extractReasoningFragments(value["text"]),
    ...extractReasoningFragments(value["summary"]),
    ...extractReasoningFragments(value["content"]),
  ];
}

function extractReasoningText(item: CodexThreadItem): string {
  if (!isRecord(item) || item["type"] !== "reasoning") {
    return "";
  }
  const text = extractReasoningFragments(item).join("\n").trim();
  return text;
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

export function createTextTimelineEntry(input: {
  id: string;
  kind: WorkbookAgentTextEntryKind;
  turnId: string | null;
  text: string;
  phase?: string | null;
  citations?: readonly WorkbookAgentTimelineCitation[];
}): WorkbookAgentTimelineEntry {
  return {
    id: input.id,
    kind: input.kind,
    turnId: input.turnId,
    text: input.text,
    phase: input.phase ?? null,
    toolName: null,
    toolStatus: null,
    argumentsText: null,
    outputText: null,
    success: null,
    citations: [...(input.citations ?? [])],
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
    return createTextTimelineEntry({
      id: item.id,
      kind: "assistant",
      turnId,
      text: item.text,
      phase: item.phase,
    });
  }

  if (isPlanItem(item)) {
    return createTextTimelineEntry({
      id: item.id,
      kind: "plan",
      turnId,
      text: item.text,
    });
  }

  if (item.type === "reasoning") {
    return createTextTimelineEntry({
      id: item.id,
      kind: "reasoning",
      turnId,
      text: extractReasoningText(item),
    });
  }

  if (isDynamicToolCallItem(item)) {
    return {
      id: item.id,
      kind: "tool",
      turnId,
      text: null,
      phase: null,
      toolName: normalizeWorkbookAgentToolName(item.tool),
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
