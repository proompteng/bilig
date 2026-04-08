import { describe, expect, it } from "vitest";
import { updatePresenceArgsSchema } from "../mutators.js";

describe("zero sync mutator schemas", () => {
  it("accepts workbook presence updates with the current selection payload", () => {
    const result = updatePresenceArgsSchema.safeParse({
      documentId: "doc-1",
      sessionId: "session-1",
      sheetName: "Sheet1",
      address: "B2",
      selection: {
        sheetName: "Sheet1",
        address: "B2",
      },
    });

    expect(result.success).toBe(true);
  });

  it("rejects workbook presence updates with malformed selection payloads", () => {
    const result = updatePresenceArgsSchema.safeParse({
      documentId: "doc-1",
      sessionId: "session-1",
      selection: {
        sheetName: "Sheet1",
        address: 42,
      },
    });

    expect(result.success).toBe(false);
  });
});
