import { describe, expect, test } from "vitest";
import { buildInternalClipboardRange, matchesInternalClipboardPaste } from "../gridInternalClipboard.js";

describe("gridInternalClipboard", () => {
  test("builds a clipboard signature and address range from copied values", () => {
    expect(
      buildInternalClipboardRange(
        { x: 1, y: 2, width: 2, height: 2 },
        [
          ["1", "2"],
          ["3", "4"]
        ]
      )
    ).toEqual({
      sourceStartAddress: "B3",
      sourceEndAddress: "C4",
      signature: "1\u001f2\u001e3\u001f4",
      plainText: "1\t2\n3\t4",
      rowCount: 2,
      colCount: 2
    });
  });

  test("matches internal clipboard pastes by signature and rectangular shape", () => {
    const clipboard = buildInternalClipboardRange(
      { x: 1, y: 2, width: 2, height: 2 },
      [
        ["1", "2"],
        ["3", "4"]
      ]
    );

    expect(
      matchesInternalClipboardPaste(clipboard, [
        ["1", "2"],
        ["3", "4"]
      ])
    ).toBe(true);

    expect(
      matchesInternalClipboardPaste(clipboard, [
        ["1", "2", "3"],
        ["4", "5", "6"]
      ])
    ).toBe(false);
  });
});
