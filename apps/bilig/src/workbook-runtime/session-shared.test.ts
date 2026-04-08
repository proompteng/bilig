import { describe, expect, it } from "vitest";
import { WORKBOOK_SNAPSHOT_CONTENT_TYPE, createSnapshotChunkFrames } from "@bilig/binary-protocol";
import type { WorkbookSnapshot } from "@bilig/protocol";
import {
  SNAPSHOT_ASSEMBLY_MAX_AGE_MS,
  acceptSnapshotChunk,
  createSnapshotPublication,
  decodeWorkbookSnapshotBytes,
  encodeWorkbookSnapshot,
} from "./session-shared.js";

function createSnapshot(): WorkbookSnapshot {
  return {
    version: 1,
    workbook: { name: "doc-1" },
    sheets: [
      {
        name: "Sheet1",
        order: 0,
        cells: [{ address: "A1", value: 1 }],
      },
    ],
  };
}

describe("session-shared", () => {
  it("reuses encoded bytes for the same snapshot object", () => {
    const snapshot = createSnapshot();

    const first = encodeWorkbookSnapshot(snapshot);
    const second = encodeWorkbookSnapshot(snapshot);
    const publication = createSnapshotPublication("doc-1", 1, snapshot);

    expect(second).toBe(first);
    expect(publication.bytes).toBe(first);
  });

  it("evicts stale snapshot assemblies before accepting a new chunk", () => {
    const registry = new Map();
    registry.set("stale", {
      documentId: "doc-stale",
      snapshotId: "stale",
      cursor: 1,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      chunkCount: 2,
      chunks: [new Uint8Array([1]), undefined],
      updatedAtUnixMs: 0,
    });

    const bytes = encodeWorkbookSnapshot(createSnapshot());
    const frames = createSnapshotChunkFrames({
      documentId: "doc-1",
      snapshotId: "fresh",
      cursor: 2,
      contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
      bytes,
    });

    const result = acceptSnapshotChunk(registry, frames[0]!, {
      nowUnixMs: SNAPSHOT_ASSEMBLY_MAX_AGE_MS + 1,
    });

    expect(result?.snapshotId).toBe("fresh");
    expect(registry.has("stale")).toBe(false);
    expect(registry.has("fresh")).toBe(false);
  });

  it("decodes assembled workbook snapshots through the shared protocol guard", () => {
    const snapshot = createSnapshot();

    expect(
      decodeWorkbookSnapshotBytes({
        documentId: "doc-1",
        snapshotId: "snapshot-1",
        cursor: 1,
        contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
        bytes: encodeWorkbookSnapshot(snapshot),
      }),
    ).toEqual(snapshot);
  });

  it("rejects assembled workbook payloads that are missing sheet cells", () => {
    expect(() =>
      decodeWorkbookSnapshotBytes({
        documentId: "doc-1",
        snapshotId: "snapshot-1",
        cursor: 1,
        contentType: WORKBOOK_SNAPSHOT_CONTENT_TYPE,
        bytes: new TextEncoder().encode(
          JSON.stringify({
            version: 1,
            workbook: { name: "doc-1" },
            sheets: [{ name: "Sheet1", order: 0 }],
          }),
        ),
      }),
    ).toThrow("Workbook snapshot payload does not match the expected schema");
  });
});
