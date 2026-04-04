// @vitest-environment jsdom
import { ValueTag, type CellSnapshot } from "@bilig/protocol";
import { createGridSelection } from "../gridSelection.js";
import {
  applyGridClipboardValues,
  captureGridClipboardSelection,
  handleGridKey,
  handleGridPasteCapture,
  shouldHandleGridSurfaceKey,
  shouldHandleGridWindowKey,
} from "../gridClipboardKeyboardController.js";
import { describe, expect, test, vi } from "vitest";
import type { GridEngineLike } from "../grid-engine.js";

function createCellSnapshot(address: string, input: string): CellSnapshot {
  return {
    sheetName: "Sheet1",
    address,
    input,
    value: { tag: ValueTag.String, value: input, stringId: 0 },
    flags: 0,
    version: 0,
  };
}

function createEngine(cells: Record<string, string>): GridEngineLike {
  return {
    getCell: (_sheetName, address) => createCellSnapshot(address, cells[address] ?? ""),
    getCellStyle: () => undefined,
    subscribeCells: () => () => {},
    workbook: {
      getSheet: () => undefined,
    },
  };
}

describe("gridClipboardKeyboardController", () => {
  test("routes external clipboard data through paste operations", () => {
    const onCopyRange = vi.fn();
    const onPaste = vi.fn();
    const internalClipboardRef = { current: null };

    applyGridClipboardValues({
      internalClipboardRef,
      onCopyRange,
      onPaste,
      sheetName: "Sheet1",
      target: [2, 3],
      values: [["A", "B"]],
    });

    expect(onCopyRange).not.toHaveBeenCalled();
    expect(onPaste).toHaveBeenCalledWith("Sheet1", "C4", [["A", "B"]]);
  });

  test("routes matching internal clipboard data through copy-range operations", () => {
    const onCopyRange = vi.fn();
    const onPaste = vi.fn();
    const internalClipboardRef = {
      current: {
        sourceStartAddress: "A1",
        sourceEndAddress: "B2",
        signature: "A\u001fB\u001eC\u001fD",
        plainText: "A\tB\nC\tD",
        rowCount: 2,
        colCount: 2,
      },
    };

    applyGridClipboardValues({
      internalClipboardRef,
      onCopyRange,
      onPaste,
      sheetName: "Sheet1",
      target: [3, 4],
      values: [
        ["A", "B"],
        ["C", "D"],
      ],
    });

    expect(onCopyRange).toHaveBeenCalledWith("A1", "B2", "D5", "E6");
    expect(onPaste).not.toHaveBeenCalled();
  });

  test("captures the selected grid range into an internal clipboard payload", () => {
    const internalClipboardRef = { current: null };

    const clipboard = captureGridClipboardSelection({
      engine: createEngine({
        A1: "alpha",
        B1: "beta",
        A2: "gamma",
        B2: "delta",
      }),
      gridSelection: {
        ...createGridSelection(0, 0),
        current: {
          cell: [0, 0],
          range: { x: 0, y: 0, width: 2, height: 2 },
          rangeStack: [],
        },
      },
      internalClipboardRef,
      sheetName: "Sheet1",
    });

    expect(clipboard).toEqual({
      sourceStartAddress: "A1",
      sourceEndAddress: "B2",
      signature: "alpha\u001fbeta\u001egamma\u001fdelta",
      plainText: "alpha\tbeta\ngamma\tdelta",
      rowCount: 2,
      colCount: 2,
    });
    expect(internalClipboardRef.current).toEqual(clipboard);
  });

  test("applies parsed paste payloads to the active selection and clears pending keyboard paste state", () => {
    const applyClipboardValues = vi.fn();
    const event = {
      clipboardData: {
        getData: (type: string) =>
          type === "text/html" ? "<table><tr><td>A</td><td>B</td></tr></table>" : "ignored",
        setData: vi.fn(),
      },
      preventDefault: vi.fn(),
      stopPropagation: vi.fn(),
    };

    handleGridPasteCapture({
      applyClipboardValues,
      event,
      gridSelection: createGridSelection(1, 2),
      pendingKeyboardPasteSequenceRef: { current: 3 },
      selectedCell: { col: 1, row: 2 },
      suppressNextNativePasteRef: { current: false },
    });

    expect(applyClipboardValues).toHaveBeenCalledWith([1, 2], [["A", "B"]]);
    expect(event.preventDefault).toHaveBeenCalled();
    expect(event.stopPropagation).toHaveBeenCalled();
  });

  test("maps keyboard actions into selection updates", () => {
    const setGridSelection = vi.fn();
    const onSelect = vi.fn();

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: "",
      event: {
        key: "Enter",
        ctrlKey: false,
        metaKey: false,
        altKey: false,
        preventDefault: vi.fn(),
      },
      gridSelection: createGridSelection(2, 4),
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onSelect,
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 2, row: 4 },
      setGridSelection,
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    });

    expect(setGridSelection.mock.calls[0]?.[0]?.current?.cell).toEqual([2, 5]);
    expect(onSelect).toHaveBeenCalledWith("C6");
  });

  test("toggles boolean cells with space instead of entering text edit mode", () => {
    const toggleSelectedBooleanCell = vi.fn();
    const preventDefault = vi.fn();

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: "",
      event: {
        key: " ",
        ctrlKey: false,
        metaKey: false,
        altKey: false,
        preventDefault,
      },
      gridSelection: createGridSelection(1, 1),
      isSelectedCellBoolean: () => true,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onSelect: vi.fn(),
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 1, row: 1 },
      setGridSelection: vi.fn(),
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell,
    });

    expect(preventDefault).toHaveBeenCalled();
    expect(toggleSelectedBooleanCell).toHaveBeenCalledTimes(1);
  });

  test("select-all updates the active address to A1", () => {
    const setGridSelection = vi.fn();
    const onSelect = vi.fn();

    handleGridKey({
      applyClipboardValues: vi.fn(),
      beginSelectedEdit: vi.fn(),
      captureInternalClipboardSelection: vi.fn(),
      editorValue: "",
      event: {
        key: "a",
        ctrlKey: true,
        metaKey: false,
        altKey: false,
        preventDefault: vi.fn(),
      },
      gridSelection: createGridSelection(3, 7),
      isSelectedCellBoolean: () => false,
      isEditingCell: false,
      onCancelEdit: vi.fn(),
      onClearCell: vi.fn(),
      onCommitEdit: vi.fn(),
      onEditorChange: vi.fn(),
      onSelect,
      pendingKeyboardPasteSequenceRef: { current: 0 },
      pendingTypeSeedRef: { current: null },
      selectedCell: { col: 3, row: 7 },
      setGridSelection,
      suppressNextNativePasteRef: { current: false },
      toggleSelectedBooleanCell: vi.fn(),
    });

    expect(setGridSelection).toHaveBeenCalledTimes(1);
    expect(onSelect).toHaveBeenCalledWith("A1");
  });

  test("only claims global grid shortcuts when focus is on the document body", () => {
    const host = document.createElement("div");
    document.body.append(host);

    expect(
      shouldHandleGridWindowKey(
        { altKey: false, ctrlKey: false, key: "Enter", metaKey: false, shiftKey: false },
        document.body,
        host,
      ),
    ).toBe(true);

    const input = document.createElement("input");
    document.body.append(input);
    input.focus();
    expect(
      shouldHandleGridWindowKey(
        { altKey: false, ctrlKey: false, key: "Enter", metaKey: false, shiftKey: false },
        input,
        host,
      ),
    ).toBe(false);
  });

  test("filters grid-surface key handling to grid-relevant keys", () => {
    expect(
      shouldHandleGridSurfaceKey({
        altKey: false,
        ctrlKey: false,
        key: "Enter",
        metaKey: false,
      }),
    ).toBe(true);

    expect(
      shouldHandleGridSurfaceKey({
        altKey: false,
        ctrlKey: false,
        key: "Shift",
        metaKey: false,
      }),
    ).toBe(false);
  });
});
