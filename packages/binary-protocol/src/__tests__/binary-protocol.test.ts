import { describe, expect, it } from "vitest";

import {
  WORKBOOK_SNAPSHOT_CONTENT_TYPE,
  createSnapshotChunkFrames,
  decodeFrame,
  encodeFrame,
} from "../index.js";

describe("binary protocol", () => {
  it("roundtrips hello frames", () => {
    const decoded = decodeFrame(
      encodeFrame({
        kind: "hello",
        documentId: "book-1",
        replicaId: "browser-a",
        sessionId: "sess-1",
        protocolVersion: 1,
        lastServerCursor: 42,
        capabilities: ["sync", "agent"],
      }),
    );

    expect(decoded).toEqual({
      kind: "hello",
      documentId: "book-1",
      replicaId: "browser-a",
      sessionId: "sess-1",
      protocolVersion: 1,
      lastServerCursor: 42,
      capabilities: ["sync", "agent"],
    });
  });

  it("roundtrips batch append frames with mixed ops", () => {
    const decoded = decodeFrame(
      encodeFrame({
        kind: "appendBatch",
        documentId: "book-1",
        cursor: 7,
        batch: {
          id: "replica:1",
          replicaId: "replica",
          clock: { counter: 1 },
          ops: [
            { kind: "upsertWorkbook", name: "spec" },
            { kind: "upsertSheet", name: "Sheet1", order: 0 },
            { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 7 },
            { kind: "setCellFormula", sheetName: "Sheet1", address: "A2", formula: "A1*2" },
            { kind: "setCellFormat", sheetName: "Sheet1", address: "A1", format: "0.00" },
            {
              kind: "upsertCellStyle",
              style: {
                id: "style-a",
                fill: { backgroundColor: "#ff0000" },
                font: { family: "IBM Plex Sans" },
              },
            },
            {
              kind: "setStyleRange",
              range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
              styleId: "style-a",
            },
            { kind: "clearCell", sheetName: "Sheet1", address: "B1" },
          ],
        },
      }),
    );

    expect(decoded.kind).toBe("appendBatch");
    if (decoded.kind !== "appendBatch") {
      return;
    }

    expect(decoded.documentId).toBe("book-1");
    expect(decoded.cursor).toBe(7);
    expect(decoded.batch.ops).toEqual([
      { kind: "upsertWorkbook", name: "spec" },
      { kind: "upsertSheet", name: "Sheet1", order: 0 },
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 7 },
      { kind: "setCellFormula", sheetName: "Sheet1", address: "A2", formula: "A1*2" },
      { kind: "setCellFormat", sheetName: "Sheet1", address: "A1", format: "0.00" },
      {
        kind: "upsertCellStyle",
        style: {
          id: "style-a",
          fill: { backgroundColor: "#ff0000" },
          font: { family: "IBM Plex Sans" },
        },
      },
      {
        kind: "setStyleRange",
        range: { sheetName: "Sheet1", startAddress: "B2", endAddress: "C3" },
        styleId: "style-a",
      },
      { kind: "clearCell", sheetName: "Sheet1", address: "B1" },
    ]);
  });

  it("roundtrips expanded workbook metadata and structural ops", () => {
    const decoded = decodeFrame(
      encodeFrame({
        kind: "appendBatch",
        documentId: "book-2",
        cursor: 9,
        batch: {
          id: "replica:2",
          replicaId: "replica",
          clock: { counter: 2 },
          ops: [
            { kind: "setWorkbookMetadata", key: "locale", value: "en-US" },
            {
              kind: "updateRowMetadata",
              sheetName: "Sheet1",
              start: 1,
              count: 2,
              size: 24,
              hidden: null,
            },
            { kind: "setFreezePane", sheetName: "Sheet1", rows: 1, cols: 2 },
            {
              kind: "setSort",
              sheetName: "Sheet1",
              range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "C10" },
              keys: [{ keyAddress: "C1", direction: "desc" }],
            },
            { kind: "upsertDefinedName", name: "TaxRate", value: 0.085 },
            {
              kind: "upsertTable",
              table: {
                name: "Sales",
                sheetName: "Sheet1",
                startAddress: "A1",
                endAddress: "C10",
                columnNames: ["Region", "Product", "Amount"],
                headerRow: true,
                totalsRow: false,
              },
            },
            {
              kind: "upsertPivotTable",
              name: "SalesPivot",
              sheetName: "Sheet1",
              address: "F2",
              source: { sheetName: "Sheet1", startAddress: "A1", endAddress: "C10" },
              groupBy: ["Region"],
              values: [{ sourceColumn: "Amount", summarizeBy: "sum" }],
              rows: 3,
              cols: 2,
            },
            { kind: "deletePivotTable", sheetName: "Sheet1", address: "F2" },
          ],
        },
      }),
    );

    expect(decoded.kind).toBe("appendBatch");
    if (decoded.kind !== "appendBatch") {
      return;
    }

    expect(decoded.batch.ops).toEqual([
      { kind: "setWorkbookMetadata", key: "locale", value: "en-US" },
      {
        kind: "updateRowMetadata",
        sheetName: "Sheet1",
        start: 1,
        count: 2,
        size: 24,
        hidden: null,
      },
      { kind: "setFreezePane", sheetName: "Sheet1", rows: 1, cols: 2 },
      {
        kind: "setSort",
        sheetName: "Sheet1",
        range: { sheetName: "Sheet1", startAddress: "A1", endAddress: "C10" },
        keys: [{ keyAddress: "C1", direction: "desc" }],
      },
      { kind: "upsertDefinedName", name: "TaxRate", value: 0.085 },
      {
        kind: "upsertTable",
        table: {
          name: "Sales",
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "C10",
          columnNames: ["Region", "Product", "Amount"],
          headerRow: true,
          totalsRow: false,
        },
      },
      {
        kind: "upsertPivotTable",
        name: "SalesPivot",
        sheetName: "Sheet1",
        address: "F2",
        source: { sheetName: "Sheet1", startAddress: "A1", endAddress: "C10" },
        groupBy: ["Region"],
        values: [{ sourceColumn: "Amount", summarizeBy: "sum" }],
        rows: 3,
        cols: 2,
      },
      { kind: "deletePivotTable", sheetName: "Sheet1", address: "F2" },
    ]);
  });

  it("roundtrips compatibility mode and explicit defined-name value payloads", () => {
    const decoded = decodeFrame(
      encodeFrame({
        kind: "appendBatch",
        documentId: "book-3",
        cursor: 10,
        batch: {
          id: "replica:3",
          replicaId: "replica",
          clock: { counter: 3 },
          ops: [
            {
              kind: "setCalculationSettings",
              settings: { mode: "manual", compatibilityMode: "odf-1.4" },
            },
            {
              kind: "upsertDefinedName",
              name: "SalesRange",
              value: {
                kind: "range-ref",
                sheetName: "Sheet1",
                startAddress: "A1",
                endAddress: "A3",
              },
            },
            {
              kind: "upsertDefinedName",
              name: "TaxExpr",
              value: { kind: "formula", formula: "=B1*0.1" },
            },
          ],
        },
      }),
    );

    expect(decoded.kind).toBe("appendBatch");
    if (decoded.kind !== "appendBatch") {
      return;
    }

    expect(decoded.batch.ops).toEqual([
      {
        kind: "setCalculationSettings",
        settings: { mode: "manual", compatibilityMode: "odf-1.4" },
      },
      {
        kind: "upsertDefinedName",
        name: "SalesRange",
        value: {
          kind: "range-ref",
          sheetName: "Sheet1",
          startAddress: "A1",
          endAddress: "A3",
        },
      },
      {
        kind: "upsertDefinedName",
        name: "TaxExpr",
        value: { kind: "formula", formula: "=B1*0.1" },
      },
    ]);
  });

  it("rejects malformed payload lengths", () => {
    const encoded = encodeFrame({
      kind: "heartbeat",
      documentId: "book-1",
      cursor: 9,
      sentAtUnixMs: 12,
    });
    encoded[7] = 0;
    encoded[8] = 0;
    encoded[9] = 0;
    encoded[10] = 0;

    expect(() => decodeFrame(encoded)).toThrow(/length mismatch/i);
  });

  it("chunks snapshot payloads into ordered snapshotChunk frames", () => {
    const frames = createSnapshotChunkFrames({
      documentId: "book-1",
      snapshotId: "snap-1",
      cursor: 3,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes: new TextEncoder().encode("abcdef"),
      chunkSize: 2,
    });

    expect(frames).toHaveLength(3);
    expect(frames.map((frame) => decodeFrame(encodeFrame(frame)))).toEqual([
      {
        kind: "snapshotChunk",
        documentId: "book-1",
        snapshotId: "snap-1",
        cursor: 3,
        chunkIndex: 0,
        chunkCount: 3,
        contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
        bytes: new TextEncoder().encode("ab"),
      },
      {
        kind: "snapshotChunk",
        documentId: "book-1",
        snapshotId: "snap-1",
        cursor: 3,
        chunkIndex: 1,
        chunkCount: 3,
        contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
        bytes: new TextEncoder().encode("cd"),
      },
      {
        kind: "snapshotChunk",
        documentId: "book-1",
        snapshotId: "snap-1",
        cursor: 3,
        chunkIndex: 2,
        chunkCount: 3,
        contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
        bytes: new TextEncoder().encode("ef"),
      },
    ]);
  });
});
