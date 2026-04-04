import { describe, expect, it, vi } from "vitest";
import { XLSX_CONTENT_TYPE } from "@bilig/agent-api";
import * as XLSX from "xlsx";
import {
  createCloseWorkbookSessionResponse,
  createOpenWorkbookSessionResponse,
  documentIdFromSessionId,
  loadWorkbookIntoRuntime,
} from "./workbook-session-shared.js";

describe("workbook-session-shared", () => {
  it("derives a document id from a session id", () => {
    expect(documentIdFromSessionId("doc-1:replica-1")).toBe("doc-1");
    expect(documentIdFromSessionId("doc-2")).toBe("doc-2");
  });

  it("creates standard open and close session responses", () => {
    expect(createOpenWorkbookSessionResponse("open-1", "doc-1:replica-1")).toEqual({
      kind: "ok",
      id: "open-1",
      sessionId: "doc-1:replica-1",
    });
    expect(createCloseWorkbookSessionResponse("close-1")).toEqual({
      kind: "ok",
      id: "close-1",
    });
  });

  it("prepares workbook imports once and delegates registration and publish hooks", async () => {
    const registerPreparedSession = vi.fn();
    const publishImportedSnapshot = vi.fn();
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([["hello"]]), "Sheet1");
    const encodedWorkbook: unknown = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
    if (!(encodedWorkbook instanceof Uint8Array)) {
      throw new Error("Expected xlsx writer to return workbook bytes");
    }

    const response = await loadWorkbookIntoRuntime(
      {
        kind: "loadWorkbookFile",
        id: "load-1",
        replicaId: "replica-1",
        fileName: "tiny.xlsx",
        contentType: XLSX_CONTENT_TYPE,
        openMode: "create",
        bytesBase64: Buffer.from(encodedWorkbook).toString("base64"),
      },
      {
        serverUrl: "http://127.0.0.1:4321",
        browserAppBaseUrl: "http://127.0.0.1:3000",
      },
      {
        registerPreparedSession,
        publishImportedSnapshot,
      },
    );

    expect(registerPreparedSession).toHaveBeenCalledTimes(1);
    expect(publishImportedSnapshot).toHaveBeenCalledTimes(1);
    expect(response).toEqual(
      expect.objectContaining({
        kind: "workbookLoaded",
        id: "load-1",
        sessionId: expect.stringContaining(":replica-1"),
        serverUrl: "http://127.0.0.1:4321",
      }),
    );
  });
});
