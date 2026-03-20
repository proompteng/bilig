import { describe, expect, test } from "vitest";
import {
  parseClipboardPlainText,
  serializeClipboardMatrix,
  serializeClipboardPlainText
} from "../gridClipboard.js";

describe("gridClipboard", () => {
  test("serializes matrix signatures and plain text", () => {
    const values = [
      ["A", "B"],
      ["1", "2"]
    ] as const;

    expect(serializeClipboardMatrix(values)).toBe("A\u001fB\u001e1\u001f2");
    expect(serializeClipboardPlainText(values)).toBe("A\tB\n1\t2");
  });

  test("parses clipboard plain text with mixed line endings", () => {
    expect(parseClipboardPlainText("A\tB\r\n1\t2\r3\t4")).toEqual([
      ["A", "B"],
      ["1", "2"],
      ["3", "4"]
    ]);
    expect(parseClipboardPlainText("")).toEqual([]);
  });
});
