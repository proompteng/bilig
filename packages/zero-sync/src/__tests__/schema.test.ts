import { describe, expect, it } from "vitest";
import { schema } from "../schema";

describe("zero sync schema", () => {
  it("maps workbooks.updated_at as a numeric timestamp", () => {
    expect(schema.tables.workbooks.columns.updatedAt.type).toBe("number");
    expect(schema.tables.workbooks.columns.updatedAt.serverName).toBe("updated_at");
  });

  it("exposes the normalized workbook projection tables", () => {
    expect(schema.tables.cells.columns.rowNum.serverName).toBe("row_num");
    expect(schema.tables.computed_cells.columns.calcRevision.serverName).toBe("calc_revision");
    expect(schema.tables.cell_styles.name).toBe("cell_styles");
    expect(schema.tables.cell_number_formats.name).toBe("cell_number_formats");
    expect(schema.tables.sheet_style_ranges.name).toBe("sheet_style_ranges");
    expect(schema.tables.sheet_format_ranges.name).toBe("sheet_format_ranges");
  });
});
