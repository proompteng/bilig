import { describe, expect, it } from "vitest";
import {
  decodeUnknownSync,
  DocumentStateSummarySchema,
  ErrorEnvelopeSchema,
  RuntimeSessionSchema,
} from "../index.js";

describe("@bilig/contracts", () => {
  it("decodes a v2 runtime session payload", () => {
    const decoded = decodeUnknownSync(RuntimeSessionSchema, {
      authToken: "user-123",
      userId: "user-123",
      roles: ["editor"],
      isAuthenticated: true,
      authSource: "header",
    });

    expect(decoded.authToken).toBe("user-123");
    expect(decoded.authSource).toBe("header");
  });

  it("decodes a v2 document state payload", () => {
    const decoded = decodeUnknownSync(DocumentStateSummarySchema, {
      documentId: "book-1",
      cursor: 4,
      owner: null,
      sessions: ["browser:1"],
      latestSnapshotCursor: 3,
    });

    expect(decoded.documentId).toBe("book-1");
    expect(decoded.latestSnapshotCursor).toBe(3);
  });

  it("decodes a v2 error envelope", () => {
    const decoded = decodeUnknownSync(ErrorEnvelopeSchema, {
      error: "TEST_FAILURE",
      message: "boom",
      retryable: false,
    });

    expect(decoded.retryable).toBe(false);
  });
});
