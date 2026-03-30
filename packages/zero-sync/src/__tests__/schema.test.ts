import { describe, expect, it } from "vitest";
import { schema } from "../schema";

describe("zero sync schema", () => {
  it("maps workbooks.updated_at as a numeric timestamp", () => {
    expect(schema.tables.workbooks.columns.updatedAt.type).toBe("number");
    expect(schema.tables.workbooks.columns.updatedAt.serverName).toBe("updated_at");
    expect("snapshot" in schema.tables.workbooks.columns).toBe(false);
    expect("replicaSnapshot" in schema.tables.workbooks.columns).toBe(false);
  });

  it("exposes the normalized workbook projection tables", () => {
    expect(schema.tables.cells.columns.rowNum.serverName).toBe("row_num");
    expect(schema.tables.cells.columns.styleId.serverName).toBe("style_id");
    expect(schema.tables.cell_eval.columns.calcRevision.serverName).toBe("calc_revision");
    expect(schema.tables.cell_eval.columns.styleId.serverName).toBe("style_id");
    expect(schema.tables.cell_eval.columns.styleJson.serverName).toBe("style_json");
    expect(schema.tables.cell_eval.columns.formatId.serverName).toBe("format_id");
    expect(schema.tables.cell_eval.columns.formatCode.serverName).toBe("format_code");
    expect("cell_styles" in schema.tables).toBe(false);
    expect("cell_number_formats" in schema.tables).toBe(false);
    expect("sheet_style_ranges" in schema.tables).toBe(false);
    expect("sheet_format_ranges" in schema.tables).toBe(false);
    expect("workbook_metadata" in schema.tables).toBe(false);
    expect("calculation_settings" in schema.tables).toBe(false);
  });
});
