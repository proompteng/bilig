import { defineQueriesWithType, defineQuery } from "@rocicorp/zero";
import { z } from "zod";
import { schema } from "./schema.js";
import { zql } from "./zql.js";

const defineQueries = defineQueriesWithType<typeof schema>();

export const workbookQueryArgsSchema = z.object({
  documentId: z.string().min(1),
});

const workbookSheetArgsSchema = workbookQueryArgsSchema
  .extend({
    sheetId: z.string().min(1).optional(),
    sheetName: z.string().min(1).optional(),
  })
  .refine((args) => args.sheetId !== undefined || args.sheetName !== undefined, {
    message: "sheetId or sheetName is required",
  });

function resolveSheetId(args: z.infer<typeof workbookSheetArgsSchema>): string {
  return args.sheetId ?? args.sheetName ?? "";
}

export const workbookCellArgsSchema = workbookSheetArgsSchema.extend({
  address: z.string().min(1),
});

export const workbookTileArgsSchema = workbookSheetArgsSchema.extend({
  rowStart: z.number().int().nonnegative(),
  rowEnd: z.number().int().nonnegative(),
  colStart: z.number().int().nonnegative(),
  colEnd: z.number().int().nonnegative(),
});

export const workbookRowTileArgsSchema = workbookSheetArgsSchema.extend({
  rowStart: z.number().int().nonnegative(),
  rowEnd: z.number().int().nonnegative(),
});

export const workbookColumnTileArgsSchema = workbookSheetArgsSchema.extend({
  colStart: z.number().int().nonnegative(),
  colEnd: z.number().int().nonnegative(),
});

const workbookGet = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.workbooks.where("id", documentId).one(),
);

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function resolveQueryUserId(ctx: unknown): string {
  return isRecord(ctx) && typeof ctx["userID"] === "string" ? ctx["userID"] : "";
}

const sheetByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.sheets.where("workbookId", documentId).orderBy("sortOrder", "asc"),
);

const cellInputOne = defineQuery(workbookCellArgsSchema, ({ args }) =>
  zql.cells
    .where("workbookId", args.documentId)
    .where("sheetName", resolveSheetId(args))
    .where("address", args.address)
    .one(),
);

const cellInputTile = defineQuery(workbookTileArgsSchema, ({ args }) =>
  zql.cells
    .where("workbookId", args.documentId)
    .where("sheetName", resolveSheetId(args))
    .where("rowNum", ">=", args.rowStart)
    .where("rowNum", "<=", args.rowEnd)
    .where("colNum", ">=", args.colStart)
    .where("colNum", "<=", args.colEnd)
    .orderBy("rowNum", "asc")
    .orderBy("colNum", "asc"),
);

const cellEvalOne = defineQuery(workbookCellArgsSchema, ({ args }) =>
  zql.cell_eval
    .where("workbookId", args.documentId)
    .where("sheetName", resolveSheetId(args))
    .where("address", args.address)
    .one(),
);

const cellEvalTile = defineQuery(workbookTileArgsSchema, ({ args }) =>
  zql.cell_eval
    .where("workbookId", args.documentId)
    .where("sheetName", resolveSheetId(args))
    .where("rowNum", ">=", args.rowStart)
    .where("rowNum", "<=", args.rowEnd)
    .where("colNum", ">=", args.colStart)
    .where("colNum", "<=", args.colEnd)
    .orderBy("rowNum", "asc")
    .orderBy("colNum", "asc"),
);

const sheetRowTile = defineQuery(workbookRowTileArgsSchema, ({ args }) =>
  zql.row_metadata
    .where("workbookId", args.documentId)
    .where("sheetName", resolveSheetId(args))
    .where("startIndex", ">=", args.rowStart)
    .where("startIndex", "<=", args.rowEnd)
    .orderBy("startIndex", "asc"),
);

const sheetColTile = defineQuery(workbookColumnTileArgsSchema, ({ args }) =>
  zql.column_metadata
    .where("workbookId", args.documentId)
    .where("sheetName", resolveSheetId(args))
    .where("startIndex", ">=", args.colStart)
    .where("startIndex", "<=", args.colEnd)
    .orderBy("startIndex", "asc"),
);

const cellStyleByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.cell_styles.where("workbookId", documentId).orderBy("styleId", "asc"),
);

const numberFormatByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.cell_number_formats.where("workbookId", documentId).orderBy("formatId", "asc"),
);

const presenceCoarseByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.presence_coarse.where("workbookId", documentId).orderBy("updatedAt", "desc"),
);

const sheetViewByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId }, ctx }) =>
  zql.sheet_view
    .where((eb) =>
      eb.and(
        eb.cmp("workbookId", documentId),
        eb.or(eb.cmp("visibility", "shared"), eb.cmp("ownerUserId", resolveQueryUserId(ctx))),
      ),
    )
    .orderBy("updatedAt", "desc")
    .orderBy("name", "asc"),
);

const workbookChangeByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.workbook_change
    .where("workbookId", documentId)
    .orderBy("createdAt", "desc")
    .orderBy("revision", "desc"),
);

const workbookVersionByWorkbook = defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
  zql.workbook_version
    .where("workbookId", documentId)
    .orderBy("updatedAt", "desc")
    .orderBy("createdAt", "desc")
    .orderBy("name", "asc"),
);

const workbookScenarioByWorkbook = defineQuery(
  workbookQueryArgsSchema,
  ({ args: { documentId }, ctx }) =>
    zql.workbook_scenario
      .where("workbookId", documentId)
      .where("ownerUserId", resolveQueryUserId(ctx))
      .orderBy("updatedAt", "desc")
      .orderBy("createdAt", "desc")
      .orderBy("name", "asc"),
);

const workbookScenarioByDocument = defineQuery(
  workbookQueryArgsSchema,
  ({ args: { documentId }, ctx }) =>
    zql.workbook_scenario
      .where("documentId", documentId)
      .where("ownerUserId", resolveQueryUserId(ctx))
      .one(),
);

export const queries = defineQueries({
  workbook: {
    get: workbookGet,
  },
  workbooks: {
    get: workbookGet,
  },
  sheet: {
    byWorkbook: sheetByWorkbook,
  },
  sheets: {
    byWorkbook: sheetByWorkbook,
  },
  cellInput: {
    one: cellInputOne,
    tile: cellInputTile,
  },
  cells: {
    one: cellInputOne,
    tile: cellInputTile,
  },
  cellEval: {
    one: cellEvalOne,
    tile: cellEvalTile,
  },
  cellRender: {
    one: cellEvalOne,
    tile: cellEvalTile,
  },
  sheetRow: {
    tile: sheetRowTile,
  },
  rowMetadata: {
    tile: sheetRowTile,
  },
  sheetCol: {
    tile: sheetColTile,
  },
  columnMetadata: {
    tile: sheetColTile,
  },
  cellStyle: {
    byWorkbook: cellStyleByWorkbook,
  },
  numberFormat: {
    byWorkbook: numberFormatByWorkbook,
  },
  presenceCoarse: {
    byWorkbook: presenceCoarseByWorkbook,
  },
  presence: {
    byWorkbook: presenceCoarseByWorkbook,
  },
  sheetView: {
    byWorkbook: sheetViewByWorkbook,
  },
  sheetViews: {
    byWorkbook: sheetViewByWorkbook,
  },
  workbookChange: {
    byWorkbook: workbookChangeByWorkbook,
  },
  workbookChanges: {
    byWorkbook: workbookChangeByWorkbook,
  },
  workbookVersion: {
    byWorkbook: workbookVersionByWorkbook,
  },
  workbookVersions: {
    byWorkbook: workbookVersionByWorkbook,
  },
  workbookScenario: {
    byWorkbook: workbookScenarioByWorkbook,
    byDocument: workbookScenarioByDocument,
  },
  workbookScenarios: {
    byWorkbook: workbookScenarioByWorkbook,
    byDocument: workbookScenarioByDocument,
  },
});
