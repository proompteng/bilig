import {
  WORKBOOK_AGENT_TOOL_NAMES,
  normalizeWorkbookAgentToolName,
  type CodexDynamicToolCallRequest,
  type CodexDynamicToolCallResult,
  type CodexDynamicToolSpec,
} from "@bilig/agent-api";
import { z } from "zod";
import type { ZeroSyncService } from "../zero/service.js";
import {
  scanWorkbookBrokenReferences,
  scanWorkbookHiddenRowsAffectingResults,
  scanWorkbookInconsistentFormulas,
  scanWorkbookPerformanceHotspots,
  scanWorkbookUsedRangeBloat,
  verifyWorkbookInvariants,
} from "./workbook-agent-audit.js";

const auditArgsSchema = z.object({
  sheetName: z.string().trim().min(1).optional(),
  limit: z.number().int().positive().max(200).optional(),
});

const hiddenRowsAuditArgsSchema = auditArgsSchema.extend({
  depth: z.number().int().positive().max(6).optional(),
});

const verifyInvariantArgsSchema = z.object({
  roundTrip: z.boolean().optional(),
});

export const workbookAgentAuditToolSpecs = [
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.scanBrokenReferences,
    description: "Scan formulas for broken #REF! references, optionally scoped to one sheet.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.scanHiddenRowsAffectingResults,
    description:
      "Find formulas whose precedent chains traverse hidden rows, optionally scoped to one sheet.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        sheetName: { type: "string" },
        limit: { type: "number" },
        depth: { type: "number" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.scanInconsistentFormulas,
    description:
      "Detect copied-formula runs with outliers in contiguous rows or columns, optionally scoped to one sheet.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.scanUsedRangeBloat,
    description:
      "Report sheets where metadata extends the effective used range well beyond populated cells.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.scanPerformanceHotspots,
    description:
      "Rank workbook sheets that are likely to be recalc or maintenance hotspots based on formulas, pivots, spills, and JS-only formulas.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.verifyInvariants,
    description:
      "Verify workbook structural invariants and optionally run a snapshot round-trip stability check.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        roundTrip: { type: "boolean" },
      },
    },
  },
] satisfies readonly CodexDynamicToolSpec[];

export interface WorkbookAgentAuditToolContext {
  readonly documentId: string;
  readonly zeroSyncService: ZeroSyncService;
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

export async function handleWorkbookAgentAuditToolCall(
  context: WorkbookAgentAuditToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult | null> {
  const normalizedTool = normalizeWorkbookAgentToolName(request.tool);
  switch (normalizedTool) {
    case WORKBOOK_AGENT_TOOL_NAMES.scanBrokenReferences: {
      const args = auditArgsSchema.parse(request.arguments);
      const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
        scanWorkbookBrokenReferences(runtime, args),
      );
      return textToolResult(stringifyJson(report));
    }
    case WORKBOOK_AGENT_TOOL_NAMES.scanHiddenRowsAffectingResults: {
      const args = hiddenRowsAuditArgsSchema.parse(request.arguments);
      const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
        scanWorkbookHiddenRowsAffectingResults(runtime, args),
      );
      return textToolResult(stringifyJson(report));
    }
    case WORKBOOK_AGENT_TOOL_NAMES.scanInconsistentFormulas: {
      const args = auditArgsSchema.parse(request.arguments);
      const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
        scanWorkbookInconsistentFormulas(runtime, args),
      );
      return textToolResult(stringifyJson(report));
    }
    case WORKBOOK_AGENT_TOOL_NAMES.scanUsedRangeBloat: {
      const args = auditArgsSchema.parse(request.arguments);
      const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
        scanWorkbookUsedRangeBloat(runtime, args),
      );
      return textToolResult(stringifyJson(report));
    }
    case WORKBOOK_AGENT_TOOL_NAMES.scanPerformanceHotspots: {
      const args = auditArgsSchema.parse(request.arguments);
      const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
        scanWorkbookPerformanceHotspots(runtime, args),
      );
      return textToolResult(stringifyJson(report));
    }
    case WORKBOOK_AGENT_TOOL_NAMES.verifyInvariants: {
      const args = verifyInvariantArgsSchema.parse(request.arguments);
      const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
        verifyWorkbookInvariants(runtime, args),
      );
      return textToolResult(stringifyJson(report));
    }
    default:
      return null;
  }
}
