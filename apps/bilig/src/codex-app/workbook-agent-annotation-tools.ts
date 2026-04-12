import { parseCellAddress } from "@bilig/formula";
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
import type { WorkbookCommentThreadSnapshot, WorkbookNoteSnapshot } from "@bilig/protocol";
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

const listCommentsArgsSchema = z
  .object({
    range: rangeOrSelectorSchema.shape.range.optional(),
    selector: rangeOrSelectorSchema.shape.selector.optional(),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) <= 1, {
    message: "Provide at most one of range or selector",
  });

const singleTargetArgsSchema = z.object({
  range: rangeOrSelectorSchema.shape.range.optional(),
  selector: rangeOrSelectorSchema.shape.selector.optional(),
  text: z.string().trim().min(1).optional(),
});

const replyCommentArgsSchema = z.object({
  range: rangeOrSelectorSchema.shape.range.optional(),
  selector: rangeOrSelectorSchema.shape.selector.optional(),
  text: z.string().trim().min(1),
});

export const workbookAgentAnnotationToolSpecs = [
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.getComments,
    description:
      "List workbook comment threads and notes. Optionally filter to an explicit range or semantic selector.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.addComment,
    description: "Add a new comment thread to a single-cell range or semantic selector target.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["text"],
      properties: {
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        text: { type: "string" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.replyComment,
    description: "Reply to an existing comment thread on a single-cell range or selector target.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["text"],
      properties: {
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        text: { type: "string" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.resolveComment,
    description: "Mark an existing comment thread as resolved on a single target cell.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.deleteComment,
    description: "Delete a comment thread from a single target cell.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.addNote,
    description: "Add a note to a single target cell.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["text"],
      properties: {
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        text: { type: "string" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.updateNote,
    description: "Update an existing note on a single target cell.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["text"],
      properties: {
        range: rangeOrSelectorJsonSchema.properties.range,
        selector: rangeOrSelectorJsonSchema.properties.selector,
        text: { type: "string" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.deleteNote,
    description: "Delete a note from a single target cell.",
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

export interface WorkbookAgentAnnotationToolContext {
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
  context: WorkbookAgentAnnotationToolContext,
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

function assertSingleCellRange(range: {
  readonly sheetName: string;
  readonly startAddress: string;
  readonly endAddress: string;
}): { sheetName: string; address: string } {
  if (range.startAddress !== range.endAddress) {
    throw new Error("Comments and notes require a single-cell target");
  }
  return {
    sheetName: range.sheetName,
    address: range.startAddress,
  };
}

function listWorkbookAnnotations(runtime: WorkbookRuntime) {
  const sheets = runtime.engine.exportSnapshot().sheets.map((sheet) => sheet.name);
  return {
    commentThreads: sheets.flatMap((sheetName) => runtime.engine.getCommentThreads(sheetName)),
    notes: sheets.flatMap((sheetName) => runtime.engine.getNotes(sheetName)),
  };
}

function annotationsIntersect(
  target: { sheetName: string; startAddress: string; endAddress: string },
  address: { sheetName: string; address: string },
): boolean {
  if (target.sheetName !== address.sheetName) {
    return false;
  }
  const start = parseCellAddress(target.startAddress, target.sheetName);
  const end = parseCellAddress(target.endAddress, target.sheetName);
  const cell = parseCellAddress(address.address, address.sheetName);
  const startRow = Math.min(start.row, end.row);
  const endRow = Math.max(start.row, end.row);
  const startCol = Math.min(start.col, end.col);
  const endCol = Math.max(start.col, end.col);
  return cell.row >= startRow && cell.row <= endRow && cell.col >= startCol && cell.col <= endCol;
}

export async function handleWorkbookAgentAnnotationToolCall(
  context: WorkbookAgentAnnotationToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult | null> {
  const normalizedTool = normalizeWorkbookAgentToolName(request.tool);
  switch (normalizedTool) {
    case WORKBOOK_AGENT_TOOL_NAMES.getComments: {
      const args = listCommentsArgsSchema.parse(request.arguments);
      const payload = await context.zeroSyncService.inspectWorkbook(
        context.documentId,
        (runtime) => {
          const annotations = listWorkbookAnnotations(runtime);
          if (!args.range && !args.selector) {
            return {
              documentId: context.documentId,
              commentThreadCount: annotations.commentThreads.length,
              noteCount: annotations.notes.length,
              commentThreads: annotations.commentThreads,
              notes: annotations.notes,
            };
          }
          const resolved = resolveRangeOrSelectorRequest({
            runtime,
            args: {
              ...(args.range ? { range: args.range } : {}),
              ...(args.selector ? { selector: args.selector } : {}),
            },
            uiContext: context.uiContext,
          });
          return {
            documentId: context.documentId,
            commentThreadCount: annotations.commentThreads.filter((thread) =>
              annotationsIntersect(resolved.range, thread),
            ).length,
            noteCount: annotations.notes.filter((note) =>
              annotationsIntersect(resolved.range, note),
            ).length,
            commentThreads: annotations.commentThreads.filter((thread) =>
              annotationsIntersect(resolved.range, thread),
            ),
            notes: annotations.notes.filter((note) => annotationsIntersect(resolved.range, note)),
          };
        },
      );
      return textToolResult(stringifyJson(payload));
    }
    case WORKBOOK_AGENT_TOOL_NAMES.addComment: {
      const args = singleTargetArgsSchema.parse(request.arguments);
      if (!args.text) {
        throw new Error("text is required");
      }
      const text = args.text;
      const thread = await context.zeroSyncService.inspectWorkbook(
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
          const target = assertSingleCellRange(resolved.range);
          return {
            threadId: crypto.randomUUID(),
            sheetName: target.sheetName,
            address: target.address,
            comments: [{ id: crypto.randomUUID(), body: text }],
          } satisfies WorkbookCommentThreadSnapshot;
        },
      );
      return await stageCommandResult(context, { kind: "upsertCommentThread", thread });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.replyComment: {
      const args = replyCommentArgsSchema.parse(request.arguments);
      const thread = await context.zeroSyncService.inspectWorkbook(
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
          const target = assertSingleCellRange(resolved.range);
          const existing = runtime.engine.getCommentThread(target.sheetName, target.address);
          if (!existing) {
            throw new Error("Comment thread does not exist at the target cell");
          }
          return {
            ...existing,
            comments: [...existing.comments, { id: crypto.randomUUID(), body: args.text }],
          } satisfies WorkbookCommentThreadSnapshot;
        },
      );
      return await stageCommandResult(context, { kind: "upsertCommentThread", thread });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.resolveComment: {
      const args = singleTargetArgsSchema.parse(request.arguments);
      const thread = await context.zeroSyncService.inspectWorkbook(
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
          const target = assertSingleCellRange(resolved.range);
          const existing = runtime.engine.getCommentThread(target.sheetName, target.address);
          if (!existing) {
            throw new Error("Comment thread does not exist at the target cell");
          }
          return {
            ...existing,
            resolved: true,
          } satisfies WorkbookCommentThreadSnapshot;
        },
      );
      return await stageCommandResult(context, { kind: "upsertCommentThread", thread });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deleteComment: {
      const args = singleTargetArgsSchema.parse(request.arguments);
      const target = await context.zeroSyncService.inspectWorkbook(
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
          return assertSingleCellRange(resolved.range);
        },
      );
      return await stageCommandResult(context, {
        kind: "deleteCommentThread",
        sheetName: target.sheetName,
        address: target.address,
      });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.addNote:
    case WORKBOOK_AGENT_TOOL_NAMES.updateNote: {
      const args = singleTargetArgsSchema.parse(request.arguments);
      if (!args.text) {
        throw new Error("text is required");
      }
      const text = args.text;
      const note = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
        const resolved = resolveRangeOrSelectorRequest({
          runtime,
          args: {
            ...(args.range ? { range: args.range } : {}),
            ...(args.selector ? { selector: args.selector } : {}),
          },
          uiContext: context.uiContext,
        });
        const target = assertSingleCellRange(resolved.range);
        return {
          sheetName: target.sheetName,
          address: target.address,
          text,
        } satisfies WorkbookNoteSnapshot;
      });
      return await stageCommandResult(context, { kind: "upsertNote", note });
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deleteNote: {
      const args = singleTargetArgsSchema.parse(request.arguments);
      const target = await context.zeroSyncService.inspectWorkbook(
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
          return assertSingleCellRange(resolved.range);
        },
      );
      return await stageCommandResult(context, {
        kind: "deleteNote",
        sheetName: target.sheetName,
        address: target.address,
      });
    }
    default:
      return null;
  }
}
