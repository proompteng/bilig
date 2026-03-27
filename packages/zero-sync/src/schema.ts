import { boolean, createSchema, json, number, relationships, string, table } from "@rocicorp/zero";

const workbooks = table("workbooks")
  .columns({
    id: string(),
    name: string(),
    snapshot: json(),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("id");

const sheets = table("sheets")
  .columns({
    workbookId: string().from("workbook_id"),
    name: string(),
    sortOrder: number().from("sort_order"),
  })
  .primaryKey("workbookId", "name");

const cells = table("cells")
  .columns({
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    address: string(),
    inputValue: json().from("input_value").optional(),
    formula: string().optional(),
    format: string().optional(),
  })
  .primaryKey("workbookId", "sheetName", "address");

const computedCells = table("computed_cells")
  .columns({
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    address: string(),
    value: json(),
    flags: number(),
    version: number(),
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
  })
  .primaryKey("workbookId", "sheetName", "startIndex");

const definedNames = table("defined_names")
  .columns({
    workbookId: string().from("workbook_id"),
    name: string(),
    value: json(),
  })
  .primaryKey("workbookId", "name");

const workbookMetadata = table("workbook_metadata")
  .columns({
    workbookId: string().from("workbook_id"),
    key: string(),
    value: json(),
  })
  .primaryKey("workbookId", "key");

const calculationSettings = table("calculation_settings")
  .columns({
    workbookId: string().from("workbook_id"),
    mode: string<"automatic" | "manual">(),
    recalcEpoch: number().from("recalc_epoch"),
  })
  .primaryKey("workbookId");

export const schema = createSchema({
  tables: [
    workbooks,
    sheets,
    cells,
    computedCells,
    rowMetadata,
    columnMetadata,
    definedNames,
    workbookMetadata,
    calculationSettings,
  ],
  relationships: [
    relationships(workbooks, ({ many, one }) => ({
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
      workbookMetadataEntries: many({
        sourceField: ["id"],
        destField: ["workbookId"],
        destSchema: workbookMetadata,
      }),
      calculationSettings: one({
        sourceField: ["id"],
        destField: ["workbookId"],
        destSchema: calculationSettings,
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
      computedCells: many({
        sourceField: ["workbookId", "name"],
        destField: ["workbookId", "sheetName"],
        destSchema: computedCells,
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
    })),
    relationships(cells, ({ one }) => ({
      sheet: one({
        sourceField: ["workbookId", "sheetName"],
        destField: ["workbookId", "name"],
        destSchema: sheets,
      }),
    })),
    relationships(computedCells, ({ one }) => ({
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
    relationships(definedNames, ({ one }) => ({
      workbook: one({
        sourceField: ["workbookId"],
        destField: ["id"],
        destSchema: workbooks,
      }),
    })),
    relationships(workbookMetadata, ({ one }) => ({
      workbook: one({
        sourceField: ["workbookId"],
        destField: ["id"],
        destSchema: workbooks,
      }),
    })),
    relationships(calculationSettings, ({ one }) => ({
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
