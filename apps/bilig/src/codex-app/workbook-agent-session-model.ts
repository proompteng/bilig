import { renderWorkbookAgentSkillInstructions } from "@bilig/agent-api";
import type { WorkbookAgentSessionSnapshot, WorkbookAgentTimelineEntry } from "@bilig/contracts";
import { z } from "zod";
import type { CodexThread, CodexThreadItem } from "./codex-app-server-types.js";

export const createSessionBodySchema = z.object({
  sessionId: z.string().min(1).optional(),
  threadId: z.string().min(1).optional(),
  context: z
    .object({
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
    })
    .optional(),
});

export const updateContextBodySchema = z.object({
  context: z.object({
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
  }),
});

export const startTurnBodySchema = z.object({
  prompt: z.string().trim().min(1),
  context: z
    .object({
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
    })
    .optional(),
});

export function createWorkbookAgentBaseInstructions(): string {
  return [
    "You are the bilig workbook assistant embedded inside a spreadsheet product.",
    "Stay narrowly focused on inspecting and editing the active workbook.",
    "Use the provided bilig.* local workbook skills and dynamic tools for spreadsheet work.",
    "Do not use filesystem, shell, web, connector, or unrelated tools.",
    renderWorkbookAgentSkillInstructions(),
  ].join(" ");
}

export function createWorkbookAgentDeveloperInstructions(): string {
  return [
    "Before changing cells you have not inspected, read the relevant workbook range first.",
    "When the user refers to the current cell, selection, or visible area, call bilig.get_context.",
    "Prefer bilig.read_selection, bilig.read_visible_range, and bilig.inspect_cell for context-native workbook analysis.",
    "All workbook writes must stage semantic preview bundles instead of applying immediately.",
    "Use the bundle-staging workbook tools to assemble one coherent preview per turn when the task is related.",
    "After staging workbook changes, summarize the preview and tell the user to review and apply it from the rail.",
    "If the requested action is outside the available bilig.* tools, say exactly which workbook capability is missing instead of improvising.",
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
