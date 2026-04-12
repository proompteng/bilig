import { describe, expect, it } from "vitest";
import { SpreadsheetEngine } from "../engine.js";

describe("SpreadsheetEngine comments and notes", () => {
  it("roundtrips comment threads and notes through snapshots", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "annotation-roundtrip" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCommentThread({
      threadId: "thread-1",
      sheetName: "Sheet1",
      address: "B2",
      comments: [{ id: "comment-1", body: "Check this total." }],
    });
    engine.setNote({
      sheetName: "Sheet1",
      address: "C3",
      text: "Manual override",
    });

    const snapshot = engine.exportSnapshot();
    expect(
      snapshot.sheets.find((sheet) => sheet.name === "Sheet1")?.metadata?.commentThreads,
    ).toEqual([
      {
        threadId: "thread-1",
        sheetName: "Sheet1",
        address: "B2",
        comments: [{ id: "comment-1", body: "Check this total." }],
      },
    ]);
    expect(snapshot.sheets.find((sheet) => sheet.name === "Sheet1")?.metadata?.notes).toEqual([
      {
        sheetName: "Sheet1",
        address: "C3",
        text: "Manual override",
      },
    ]);

    const restored = new SpreadsheetEngine({ workbookName: "annotation-roundtrip-restored" });
    await restored.ready();
    restored.importSnapshot(snapshot);

    expect(restored.getCommentThreads("Sheet1")).toEqual([
      {
        threadId: "thread-1",
        sheetName: "Sheet1",
        address: "B2",
        comments: [{ id: "comment-1", body: "Check this total." }],
      },
    ]);
    expect(restored.getNotes("Sheet1")).toEqual([
      {
        sheetName: "Sheet1",
        address: "C3",
        text: "Manual override",
      },
    ]);
  });

  it("rewrites comment and note anchors across structural edits", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "annotation-structural" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCommentThread({
      threadId: "thread-1",
      sheetName: "Sheet1",
      address: "B2",
      comments: [{ id: "comment-1", body: "Check this total." }],
    });
    engine.setNote({
      sheetName: "Sheet1",
      address: "C3",
      text: "Manual override",
    });

    engine.insertRows("Sheet1", 1, 1);
    expect(engine.getCommentThreads("Sheet1")).toEqual([
      {
        threadId: "thread-1",
        sheetName: "Sheet1",
        address: "B3",
        comments: [{ id: "comment-1", body: "Check this total." }],
      },
    ]);
    expect(engine.getNotes("Sheet1")).toEqual([
      {
        sheetName: "Sheet1",
        address: "C4",
        text: "Manual override",
      },
    ]);

    engine.deleteColumns("Sheet1", 0, 1);
    expect(engine.getCommentThreads("Sheet1")).toEqual([
      {
        threadId: "thread-1",
        sheetName: "Sheet1",
        address: "A3",
        comments: [{ id: "comment-1", body: "Check this total." }],
      },
    ]);
    expect(engine.getNotes("Sheet1")).toEqual([
      {
        sheetName: "Sheet1",
        address: "B4",
        text: "Manual override",
      },
    ]);
  });
});
