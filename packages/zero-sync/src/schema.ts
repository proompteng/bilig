import { boolean, createSchema, json, number, string, table } from "@rocicorp/zero";

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

const cellStyles = table("cell_styles")
  .columns({
    workbookId: string().from("workbook_id"),
    styleId: string().from("style_id"),
    styleJson: json().from("record_json"),
    hash: string(),
    createdAt: number().from("created_at"),
  })
  .primaryKey("workbookId", "styleId");

const numberFormats = table("cell_number_formats")
  .columns({
    workbookId: string().from("workbook_id"),
    formatId: string().from("format_id"),
    kind: string(),
    code: string(),
    createdAt: number().from("created_at"),
  })
  .primaryKey("workbookId", "formatId");

const cells = table("cells")
  .columns({
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    rowNum: number().from("row_num"),
    colNum: number().from("col_num"),
    address: string(),
    inputValue: json().from("input_value").optional(),
    formula: string().optional(),
    format: string().optional(),
    styleId: string().from("style_id").optional(),
    explicitFormatId: string().from("explicit_format_id").optional(),
    sourceRevision: number().from("source_revision"),
    updatedBy: string().from("updated_by"),
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

const cellEval = table("cell_eval")
  .columns({
    workbookId: string().from("workbook_id"),
    sheetName: string().from("sheet_name"),
    rowNum: number().from("row_num"),
    colNum: number().from("col_num"),
    address: string(),
    value: json(),
    styleId: string().from("style_id").optional(),
    formatId: string().from("format_id").optional(),
    styleJson: json().from("style_json").optional(),
    formatCode: string().from("format_code").optional(),
    flags: number(),
    version: number(),
    calcRevision: number().from("calc_revision"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("workbookId", "sheetName", "address");

const definedNames = table("defined_names")
  .columns({
    workbookId: string().from("workbook_id"),
    name: string(),
    value: json(),
  })
  .primaryKey("workbookId", "name");

const presenceCoarse = table("presence_coarse")
  .columns({
    workbookId: string().from("workbook_id"),
    sessionId: string().from("session_id"),
    userId: string().from("user_id"),
    sheetId: number().from("sheet_id").optional(),
    sheetName: string().from("sheet_name").optional(),
    address: string().optional(),
    selectionJson: json().from("selection_json").optional(),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("workbookId", "sessionId");

const sheetView = table("sheet_view")
  .columns({
    workbookId: string().from("workbook_id"),
    id: string(),
    ownerUserId: string().from("owner_user_id"),
    name: string(),
    visibility: string<"private" | "shared">(),
    sheetId: number().from("sheet_id").optional(),
    sheetName: string().from("sheet_name").optional(),
    address: string(),
    viewportJson: json().from("viewport_json"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("workbookId", "id");

const workbookChange = table("workbook_change")
  .columns({
    workbookId: string().from("workbook_id"),
    revision: number(),
    actorUserId: string().from("actor_user_id"),
    clientMutationId: string().from("client_mutation_id").optional(),
    eventKind: string().from("event_kind"),
    summary: string(),
    sheetId: number().from("sheet_id").optional(),
    sheetName: string().from("sheet_name").optional(),
    anchorAddress: string().from("anchor_address").optional(),
    rangeJson: json().from("range_json").optional(),
    undoBundleJson: json().from("undo_bundle_json").optional(),
    revertedByRevision: number().from("reverted_by_revision").optional(),
    revertsRevision: number().from("reverts_revision").optional(),
    createdAt: number().from("created_at"),
  })
  .primaryKey("workbookId", "revision");

const workbookVersion = table("workbook_version")
  .columns({
    workbookId: string().from("workbook_id"),
    id: string(),
    ownerUserId: string().from("owner_user_id"),
    name: string(),
    revision: number(),
    sheetId: number().from("sheet_id").optional(),
    sheetName: string().from("sheet_name").optional(),
    address: string().optional(),
    viewportJson: json().from("viewport_json").optional(),
    createdAt: number().from("created_at"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("workbookId", "id");

const workbookScenario = table("workbook_scenario")
  .columns({
    documentId: string().from("document_id"),
    workbookId: string().from("workbook_id"),
    ownerUserId: string().from("owner_user_id"),
    name: string(),
    baseRevision: number().from("base_revision"),
    sheetId: number().from("sheet_id").optional(),
    sheetName: string().from("sheet_name").optional(),
    address: string().optional(),
    viewportJson: json().from("viewport_json").optional(),
    createdAt: number().from("created_at"),
    updatedAt: number().from("updated_at"),
  })
  .primaryKey("documentId");

export const schema = createSchema({
  tables: [
    workbooks,
    sheets,
    cellStyles,
    numberFormats,
    cells,
    rowMetadata,
    columnMetadata,
    cellEval,
    definedNames,
    presenceCoarse,
    sheetView,
    workbookChange,
    workbookVersion,
    workbookScenario,
  ],
  relationships: [],
});

declare module "@rocicorp/zero" {
  interface DefaultTypes {
    schema: typeof schema;
  }
}
