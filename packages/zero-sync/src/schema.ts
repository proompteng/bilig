import { boolean, createSchema, json, number, relationships, string, table } from "@rocicorp/zero";

const workbooks = table("workbooks")
  .columns({
    id: string(),
    name: string(),
    ownerUserId: string().from("owner_user_id"),
    headRevision: number().from("head_revision"),
    calculatedRevision: number().from("calculated_revision"),
    calcMode: string<"automatic" | "manual">().from("calc_mode"),
    compatibilityMode: string<"excel-modern" | "odf-1.4">().from("compatibility_mode"),
    recalcEpoch: number().from("recalc_epoch"),
    createdAt: number().from("created_at"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("id");

const sheets = table("sheets")
  .columns({
    workbookId: string().from("workbook_id"),
    name: string(),
    sortOrder: number().from("sort_order"),
    freezeRows: number().from("freeze_rows"),
    freezeCols: number().from("freeze_cols"),
    createdAt: number().from("created_at"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("workbookId", "name");

const cells = table("cells")
  .columns({
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    address: string(),
    rowNum: number().from("row_num").optional(),
    colNum: number().from("col_num").optional(),
    inputValue: json().from("input_value").optional(),
    formula: string().optional(),
    format: string().optional(),
    explicitFormatId: string().from("explicit_format_id").optional(),
    sourceRevision: number().from("source_revision"),
    updatedBy: string().from("updated_by"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("workbookId", "sheetName", "address");

const cellEval = table("cell_eval")
  .columns({
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    address: string(),
    rowNum: number().from("row_num").optional(),
    colNum: number().from("col_num").optional(),
    value: json(),
    flags: number(),
    version: number(),
    styleId: string().from("style_id").optional(),
    formatId: string().from("format_id").optional(),
    formatCode: string().from("format_code").optional(),
    calcRevision: number().from("calc_revision"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("workbookId", "sheetName", "address");

const rowMetadata = table("row_metadata")
  .columns({
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    startIndex: number().from("start_index"),
    count: number(),
    size: number().optional(),
    hidden: boolean().optional(),
    sourceRevision: number().from("source_revision"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("workbookId", "sheetName", "startIndex");

const columnMetadata = table("column_metadata")
  .columns({
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    startIndex: number().from("start_index"),
    count: number(),
    size: number().optional(),
    hidden: boolean().optional(),
    sourceRevision: number().from("source_revision"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("workbookId", "sheetName", "startIndex");

const definedNames = table("defined_names")
  .columns({
    workbookId: string().from("workbook_id"),
    name: string(),
    value: json(),
  })
  .primaryKey("workbookId", "name");

const styles = table("cell_styles")
  .columns({
    workbookId: string().from("workbook_id"),
    id: string().from("style_id"),
    recordJSON: json().from("record_json"),
    hash: string(),
    createdAt: number().from("created_at"),
  })
  .primaryKey("workbookId", "id");

const numberFormats = table("cell_number_formats")
  .columns({
    workbookId: string().from("workbook_id"),
    id: string().from("format_id"),
    code: string(),
    kind: string(),
    createdAt: number().from("created_at"),
  })
  .primaryKey("workbookId", "id");

const styleRanges = table("sheet_style_ranges")
  .columns({
    id: string(),
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    startRow: number().from("start_row"),
    endRow: number().from("end_row"),
    startCol: number().from("start_col"),
    endCol: number().from("end_col"),
    styleId: string().from("style_id"),
    sourceRevision: number().from("source_revision"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("id");

const formatRanges = table("sheet_format_ranges")
  .columns({
    id: string(),
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    startRow: number().from("start_row"),
    endRow: number().from("end_row"),
    startCol: number().from("start_col"),
    endCol: number().from("end_col"),
    formatId: string().from("format_id"),
    sourceRevision: number().from("source_revision"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("id");

export const schema = createSchema({
  tables: [
    workbooks,
    sheets,
    cells,
    cellEval,
    rowMetadata,
    columnMetadata,
    definedNames,
    styles,
    numberFormats,
    styleRanges,
    formatRanges,
  ],
  relationships: [
    relationships(workbooks, ({ many }) => ({
      sheets: many({
        sourceField: ["id"],
        destField: ["workbookId"],
        destSchema: sheets,
      }),
      definedNames: many({
        sourceField: ["id"],
        destField: ["workbookId"],
        destSchema: definedNames,
      }),
      styles: many({
        sourceField: ["id"],
        destField: ["workbookId"],
        destSchema: styles,
      }),
      numberFormats: many({
        sourceField: ["id"],
        destField: ["workbookId"],
        destSchema: numberFormats,
      }),
    })),
    relationships(sheets, ({ many, one }) => ({
      workbook: one({
        sourceField: ["workbookId"],
        destField: ["id"],
        destSchema: workbooks,
      }),
      cells: many({
        sourceField: ["workbookId", "name"],
        destField: ["workbookId", "sheetName"],
        destSchema: cells,
      }),
      cellEval: many({
        sourceField: ["workbookId", "name"],
        destField: ["workbookId", "sheetName"],
        destSchema: cellEval,
      }),
      rowMetadata: many({
        sourceField: ["workbookId", "name"],
        destField: ["workbookId", "sheetName"],
        destSchema: rowMetadata,
      }),
      columnMetadata: many({
        sourceField: ["workbookId", "name"],
        destField: ["workbookId", "sheetName"],
        destSchema: columnMetadata,
      }),
      styleRanges: many({
        sourceField: ["workbookId", "name"],
        destField: ["workbookId", "sheetName"],
        destSchema: styleRanges,
      }),
      formatRanges: many({
        sourceField: ["workbookId", "name"],
        destField: ["workbookId", "sheetName"],
        destSchema: formatRanges,
      }),
    })),
    relationships(cells, ({ one }) => ({
      sheet: one({
        sourceField: ["workbookId", "sheetName"],
        destField: ["workbookId", "name"],
        destSchema: sheets,
      }),
    })),
    relationships(cellEval, ({ one }) => ({
      sheet: one({
        sourceField: ["workbookId", "sheetName"],
        destField: ["workbookId", "name"],
        destSchema: sheets,
      }),
    })),
    relationships(rowMetadata, ({ one }) => ({
      sheet: one({
        sourceField: ["workbookId", "sheetName"],
        destField: ["workbookId", "name"],
        destSchema: sheets,
      }),
    })),
    relationships(columnMetadata, ({ one }) => ({
      sheet: one({
        sourceField: ["workbookId", "sheetName"],
        destField: ["workbookId", "name"],
        destSchema: sheets,
      }),
    })),
    relationships(styleRanges, ({ one }) => ({
      sheet: one({
        sourceField: ["workbookId", "sheetName"],
        destField: ["workbookId", "name"],
        destSchema: sheets,
      }),
    })),
    relationships(formatRanges, ({ one }) => ({
      sheet: one({
        sourceField: ["workbookId", "sheetName"],
        destField: ["workbookId", "name"],
        destSchema: sheets,
      }),
    })),
    relationships(definedNames, ({ one }) => ({
      workbook: one({
        sourceField: ["workbookId"],
        destField: ["id"],
        destSchema: workbooks,
      }),
    })),
    relationships(styles, ({ one }) => ({
      workbook: one({
        sourceField: ["workbookId"],
        destField: ["id"],
        destSchema: workbooks,
      }),
    })),
    relationships(numberFormats, ({ one }) => ({
      workbook: one({
        sourceField: ["workbookId"],
        destField: ["id"],
        destSchema: workbooks,
      }),
    })),
  ],
});

declare module "@rocicorp/zero" {
  interface DefaultTypes {
    schema: typeof schema;
  }
}
