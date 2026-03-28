import { defineQueriesWithType, defineQuery } from "@rocicorp/zero";
import { z } from "zod";
import { schema } from "./schema.js";
import { zql } from "./zql.js";

const defineQueries = defineQueriesWithType<typeof schema>();

export const workbookQueryArgsSchema = z.object({
  documentId: z.string().min(1),
});

export const workbookCellArgsSchema = workbookQueryArgsSchema.extend({
  sheetName: z.string().min(1),
  address: z.string().min(1),
});

export const workbookTileArgsSchema = workbookQueryArgsSchema.extend({
  sheetName: z.string().min(1),
  rowStart: z.number().int().nonnegative(),
  rowEnd: z.number().int().nonnegative(),
  colStart: z.number().int().nonnegative(),
  colEnd: z.number().int().nonnegative(),
});

export const workbookRowTileArgsSchema = workbookQueryArgsSchema.extend({
  sheetName: z.string().min(1),
  rowStart: z.number().int().nonnegative(),
  rowEnd: z.number().int().nonnegative(),
});

export const workbookColumnTileArgsSchema = workbookQueryArgsSchema.extend({
  sheetName: z.string().min(1),
  colStart: z.number().int().nonnegative(),
  colEnd: z.number().int().nonnegative(),
});

export const queries = defineQueries({
  workbooks: {
    get: defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
      zql.workbooks.where("id", documentId).one(),
    ),
  },
  sheets: {
    byWorkbook: defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
      zql.sheets.where("workbookId", documentId).orderBy("sortOrder", "asc"),
    ),
  },
  cells: {
    one: defineQuery(workbookCellArgsSchema, ({ args }) =>
      zql.cells
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("address", args.address)
        .one(),
    ),
    tile: defineQuery(workbookTileArgsSchema, ({ args }) =>
      zql.cells
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("rowNum", ">=", args.rowStart)
        .where("rowNum", "<=", args.rowEnd)
        .where("colNum", ">=", args.colStart)
        .where("colNum", "<=", args.colEnd)
        .orderBy("rowNum", "asc")
        .orderBy("colNum", "asc"),
    ),
  },
  cellEval: {
    one: defineQuery(workbookCellArgsSchema, ({ args }) =>
      zql.cell_eval
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("address", args.address)
        .one(),
    ),
    tile: defineQuery(workbookTileArgsSchema, ({ args }) =>
      zql.cell_eval
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("rowNum", ">=", args.rowStart)
        .where("rowNum", "<=", args.rowEnd)
        .where("colNum", ">=", args.colStart)
        .where("colNum", "<=", args.colEnd)
        .orderBy("rowNum", "asc")
        .orderBy("colNum", "asc"),
    ),
  },
  computedCells: {
    one: defineQuery(workbookCellArgsSchema, ({ args }) =>
      zql.cell_eval
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("address", args.address)
        .one(),
    ),
    tile: defineQuery(workbookTileArgsSchema, ({ args }) =>
      zql.cell_eval
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("rowNum", ">=", args.rowStart)
        .where("rowNum", "<=", args.rowEnd)
        .where("colNum", ">=", args.colStart)
        .where("colNum", "<=", args.colEnd)
        .orderBy("rowNum", "asc")
        .orderBy("colNum", "asc"),
    ),
  },
  rowMetadata: {
    tile: defineQuery(workbookRowTileArgsSchema, ({ args }) =>
      zql.row_metadata
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("startIndex", "<=", args.rowEnd)
        .orderBy("startIndex", "asc"),
    ),
  },
  columnMetadata: {
    tile: defineQuery(workbookColumnTileArgsSchema, ({ args }) =>
      zql.column_metadata
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("startIndex", "<=", args.colEnd)
        .orderBy("startIndex", "asc"),
    ),
  },
  styleRanges: {
    intersectTile: defineQuery(workbookTileArgsSchema, ({ args }) =>
      zql.sheet_style_ranges
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("endRow", ">=", args.rowStart)
        .where("startRow", "<=", args.rowEnd)
        .where("endCol", ">=", args.colStart)
        .where("startCol", "<=", args.colEnd)
        .orderBy("startRow", "asc")
        .orderBy("startCol", "asc"),
    ),
  },
  formatRanges: {
    intersectTile: defineQuery(workbookTileArgsSchema, ({ args }) =>
      zql.sheet_format_ranges
        .where("workbookId", args.documentId)
        .where("sheetName", args.sheetName)
        .where("endRow", ">=", args.rowStart)
        .where("startRow", "<=", args.rowEnd)
        .where("endCol", ">=", args.colStart)
        .where("startCol", "<=", args.colEnd)
        .orderBy("startRow", "asc")
        .orderBy("startCol", "asc"),
    ),
  },
  styles: {
    byWorkbook: defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
      zql.cell_styles.where("workbookId", documentId).orderBy("id", "asc"),
    ),
  },
  numberFormats: {
    byWorkbook: defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
      zql.cell_number_formats.where("workbookId", documentId).orderBy("id", "asc"),
    ),
  },
});
