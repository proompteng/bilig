import {
  WORKBOOK_AGENT_TOOL_NAMES,
  normalizeWorkbookAgentToolName,
  type CodexDynamicToolCallRequest,
  type CodexDynamicToolSpec,
  type WorkbookAgentCommand,
} from "@bilig/agent-api";
import {
  setFreezePaneArgsSchema,
  structuralAxisMutationArgsSchema,
  updateColumnMetadataArgsSchema,
  updateRowMetadataArgsSchema,
} from "@bilig/zero-sync";
import { z } from "zod";
import {
  cellRangeRefJsonSchema,
  workbookSemanticSelectorJsonSchema,
} from "./workbook-agent-selector-tooling.js";
import { workbookSemanticSelectorSchema } from "./workbook-selector-resolver.js";

const sheetMutationToolArgsSchema = z.object({
  name: z.string().trim().min(1),
});

const renameSheetToolArgsSchema = z.object({
  currentName: z.string().trim().min(1),
  nextName: z.string().trim().min(1),
});

const structuralAxisToolArgsSchema = z.object({
  sheetName: structuralAxisMutationArgsSchema.shape.sheetName,
  start: structuralAxisMutationArgsSchema.shape.start,
  count: structuralAxisMutationArgsSchema.shape.count,
});

const freezePaneToolArgsSchema = z.object({
  sheetName: setFreezePaneArgsSchema.shape.sheetName,
  rows: setFreezePaneArgsSchema.shape.rows,
  cols: setFreezePaneArgsSchema.shape.cols,
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

const rangeToolArgsSchema = z.object({
  range: z.object({
    sheetName: z.string().min(1),
    startAddress: z.string().min(1),
    endAddress: z.string().min(1),
  }),
});

export const sortToolArgsSchema = z
  .object({
    range: rangeToolArgsSchema.shape.range.optional(),
    selector: workbookSemanticSelectorSchema.optional(),
    keys: z
      .array(
        z.object({
          keyAddress: z.string().trim().min(1),
          direction: z.enum(["asc", "desc"]),
        }),
      )
      .min(1),
  })
  .refine((value) => (value.range ? 1 : 0) + (value.selector ? 1 : 0) === 1, {
    message: "Provide exactly one of range or selector",
  });

export const workbookAgentStructuralToolSpecs = [
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
    name: WORKBOOK_AGENT_TOOL_NAMES.deleteSheet,
    description: "Delete an existing worksheet.",
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
    name: WORKBOOK_AGENT_TOOL_NAMES.insertRows,
    description: "Insert one or more rows at a zero-based row index on a sheet.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["sheetName", "start", "count"],
      properties: {
        sheetName: { type: "string" },
        start: { type: "number" },
        count: { type: "number" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.deleteRows,
    description: "Delete one or more rows at a zero-based row index on a sheet.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["sheetName", "start", "count"],
      properties: {
        sheetName: { type: "string" },
        start: { type: "number" },
        count: { type: "number" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.insertColumns,
    description: "Insert one or more columns at a zero-based column index on a sheet.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["sheetName", "start", "count"],
      properties: {
        sheetName: { type: "string" },
        start: { type: "number" },
        count: { type: "number" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.deleteColumns,
    description: "Delete one or more columns at a zero-based column index on a sheet.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["sheetName", "start", "count"],
      properties: {
        sheetName: { type: "string" },
        start: { type: "number" },
        count: { type: "number" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.setFreezePane,
    description:
      "Set frozen top rows and left columns on a sheet. Use rows=0 and cols=0 to clear frozen panes.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["sheetName", "rows", "cols"],
      properties: {
        sheetName: { type: "string" },
        rows: { type: "number" },
        cols: { type: "number" },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.setFilter,
    description: "Set a filter range on one sheet.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        range: cellRangeRefJsonSchema,
        selector: workbookSemanticSelectorJsonSchema,
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.clearFilter,
    description: "Clear a filter range on one sheet.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        range: cellRangeRefJsonSchema,
        selector: workbookSemanticSelectorJsonSchema,
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.setSort,
    description: "Set sort metadata for a range using one or more key addresses.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["keys"],
      properties: {
        range: cellRangeRefJsonSchema,
        selector: workbookSemanticSelectorJsonSchema,
        keys: {
          type: "array",
          minItems: 1,
          items: {
            type: "object",
            additionalProperties: false,
            required: ["keyAddress", "direction"],
            properties: {
              keyAddress: { type: "string" },
              direction: { type: "string", enum: ["asc", "desc"] },
            },
          },
        },
      },
    },
  },
  {
    name: WORKBOOK_AGENT_TOOL_NAMES.clearSort,
    description: "Clear sort metadata for a range on one sheet.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        range: cellRangeRefJsonSchema,
        selector: workbookSemanticSelectorJsonSchema,
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

export function parseWorkbookAgentStructuralToolCommand(
  request: CodexDynamicToolCallRequest,
): WorkbookAgentCommand | null {
  switch (normalizeWorkbookAgentToolName(request.tool)) {
    case WORKBOOK_AGENT_TOOL_NAMES.createSheet: {
      const args = sheetMutationToolArgsSchema.parse(request.arguments);
      return {
        kind: "createSheet",
        name: args.name,
      };
    }
    case WORKBOOK_AGENT_TOOL_NAMES.renameSheet: {
      const args = renameSheetToolArgsSchema.parse(request.arguments);
      return {
        kind: "renameSheet",
        currentName: args.currentName,
        nextName: args.nextName,
      };
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deleteSheet: {
      const args = sheetMutationToolArgsSchema.parse(request.arguments);
      return {
        kind: "deleteSheet",
        name: args.name,
      };
    }
    case WORKBOOK_AGENT_TOOL_NAMES.insertRows: {
      const args = structuralAxisToolArgsSchema.parse(request.arguments);
      return {
        kind: "insertRows",
        sheetName: args.sheetName,
        start: args.start,
        count: args.count,
      };
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deleteRows: {
      const args = structuralAxisToolArgsSchema.parse(request.arguments);
      return {
        kind: "deleteRows",
        sheetName: args.sheetName,
        start: args.start,
        count: args.count,
      };
    }
    case WORKBOOK_AGENT_TOOL_NAMES.insertColumns: {
      const args = structuralAxisToolArgsSchema.parse(request.arguments);
      return {
        kind: "insertColumns",
        sheetName: args.sheetName,
        start: args.start,
        count: args.count,
      };
    }
    case WORKBOOK_AGENT_TOOL_NAMES.deleteColumns: {
      const args = structuralAxisToolArgsSchema.parse(request.arguments);
      return {
        kind: "deleteColumns",
        sheetName: args.sheetName,
        start: args.start,
        count: args.count,
      };
    }
    case WORKBOOK_AGENT_TOOL_NAMES.setFreezePane: {
      const args = freezePaneToolArgsSchema.parse(request.arguments);
      return {
        kind: "setFreezePane",
        sheetName: args.sheetName,
        rows: Math.max(0, Math.round(args.rows)),
        cols: Math.max(0, Math.round(args.cols)),
      };
    }
    case WORKBOOK_AGENT_TOOL_NAMES.updateRowMetadata: {
      const args = rowMetadataToolArgsSchema.parse(request.arguments);
      return {
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
      };
    }
    case WORKBOOK_AGENT_TOOL_NAMES.updateColumnMetadata: {
      const args = columnMetadataToolArgsSchema.parse(request.arguments);
      return {
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
      };
    }
    default:
      return null;
  }
}
