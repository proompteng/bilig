import { describe, expect, it } from "vitest";
import { schema } from "../schema";

describe("zero sync schema", () => {
  it("maps workbooks.updated_at as a numeric timestamp", () => {
    expect(schema.tables.workbooks.columns.updatedAt.type).toBe("number");
    expect(schema.tables.workbooks.columns.updatedAt.serverName).toBe("updated_at");
  });
});
