import { describe, expect, test } from "vitest";
import { resolveGridKeyAction } from "../gridKeyActions.js";

describe("gridKeyActions", () => {
  test("appends printable characters during edit mode when the editor input is not focused", () => {
    expect(
      resolveGridKeyAction({
        event: { key: "x", ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: true,
        editorValue: "abc",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [0, 0],
        currentSelectionCell: [0, 0],
        currentRangeAnchor: [0, 0],
      }),
    ).toEqual({ kind: "edit-append", value: "abcx" });
  });

  test("returns begin-edit and movement actions for core keys", () => {
    expect(
      resolveGridKeyAction({
        event: { key: "F2", ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: "12",
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
      }),
    ).toEqual({ kind: "begin-edit", selectionBehavior: "caret-end", pendingTypeSeed: null });

    expect(
      resolveGridKeyAction({
        event: { key: "Enter", ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
      }),
    ).toEqual({ kind: "move-selection", cell: [2, 5] });

    expect(
      resolveGridKeyAction({
        event: { key: "ArrowRight", ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [1, 4],
      }),
    ).toEqual({ kind: "extend-selection", anchor: [1, 4], target: [3, 4] });
  });

  test("returns clipboard and typed-entry actions", () => {
    expect(
      resolveGridKeyAction({
        event: { key: "c", ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [3, 2],
        currentRangeAnchor: [1, 1],
      }),
    ).toEqual({ kind: "clipboard-copy" });

    expect(
      resolveGridKeyAction({
        event: { key: "v", ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [3, 2],
        currentRangeAnchor: [1, 1],
      }),
    ).toEqual({ kind: "clipboard-paste", target: [3, 2] });

    expect(
      resolveGridKeyAction({
        event: { key: "7", ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: "1",
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
      }),
    ).toEqual({
      kind: "begin-edit",
      seed: "17",
      selectionBehavior: "caret-end",
      pendingTypeSeed: "17",
    });
  });
});
