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

  test("commits and cancels edit mode keys before the overlay input takes focus", () => {
    expect(
      resolveGridKeyAction({
        event: { key: "Enter", ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: true,
        editorValue: "123",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [0, 0],
        currentSelectionCell: [0, 0],
        currentRangeAnchor: [0, 0],
      }),
    ).toEqual({ kind: "commit-edit", movement: [0, 1] });

    expect(
      resolveGridKeyAction({
        event: { key: "Tab", ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: true,
        editorValue: "123",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [0, 0],
        currentSelectionCell: [0, 0],
        currentRangeAnchor: [0, 0],
      }),
    ).toEqual({ kind: "commit-edit", movement: [-1, 0] });

    expect(
      resolveGridKeyAction({
        event: { key: "Escape", ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: true,
        editorValue: "123",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [0, 0],
        currentSelectionCell: [0, 0],
        currentRangeAnchor: [0, 0],
      }),
    ).toEqual({ kind: "cancel-edit" });
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

  test("supports sheet-style navigation keys and selection shortcuts", () => {
    expect(
      resolveGridKeyAction({
        event: { key: "Home", ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [5, 4],
        currentSelectionCell: [5, 4],
        currentRangeAnchor: [5, 4],
      }),
    ).toEqual({ kind: "move-selection", cell: [0, 4] });

    expect(
      resolveGridKeyAction({
        event: { key: "End", ctrlKey: true, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [5, 4],
        currentSelectionCell: [5, 4],
        currentRangeAnchor: [2, 1],
      }),
    ).toEqual({ kind: "extend-selection", anchor: [2, 1], target: [16383, 1048575] });

    expect(
      resolveGridKeyAction({
        event: { key: "PageDown", ctrlKey: false, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
      }),
    ).toEqual({ kind: "move-selection", cell: [2, 24] });

    expect(
      resolveGridKeyAction({
        event: { key: " ", ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [3, 9],
        currentSelectionCell: [3, 9],
        currentRangeAnchor: [3, 9],
      }),
    ).toEqual({ kind: "select-row", col: 3, row: 9 });

    expect(
      resolveGridKeyAction({
        event: { key: " ", ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [3, 9],
        currentSelectionCell: [3, 9],
        currentRangeAnchor: [3, 9],
      }),
    ).toEqual({ kind: "select-column", col: 3, row: 9 });

    expect(
      resolveGridKeyAction({
        event: {
          key: " ",
          ctrlKey: true,
          metaKey: false,
          altKey: false,
          shiftKey: true,
        },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [3, 9],
        currentSelectionCell: [3, 9],
        currentRangeAnchor: [3, 9],
      }),
    ).toEqual({ kind: "select-all" });

    expect(
      resolveGridKeyAction({
        event: { key: "a", ctrlKey: true, metaKey: false, altKey: false },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [1, 1],
        currentSelectionCell: [1, 1],
        currentRangeAnchor: [1, 1],
      }),
    ).toEqual({ kind: "select-all" });
  });

  test("extends an existing keyboard range from the active edge", () => {
    expect(
      resolveGridKeyAction({
        event: { key: "ArrowRight", ctrlKey: false, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        currentSelectionRange: { x: 2, y: 4, width: 2, height: 1 },
      }),
    ).toEqual({ kind: "extend-selection", anchor: [2, 4], target: [4, 4] });

    expect(
      resolveGridKeyAction({
        event: { key: "ArrowDown", ctrlKey: true, metaKey: false, altKey: false, shiftKey: true },
        isEditingCell: false,
        editorValue: "",
        editorInputFocused: false,
        pendingTypeSeed: null,
        selectedCell: [2, 4],
        currentSelectionCell: [2, 4],
        currentRangeAnchor: [2, 4],
        currentSelectionRange: { x: 2, y: 4, width: 3, height: 2 },
      }),
    ).toEqual({ kind: "extend-selection", anchor: [2, 4], target: [4, 1048575] });
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
