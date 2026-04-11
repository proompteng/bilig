import { describe, expect, it, vi } from "vitest";
import { CompactSelection } from "../gridTypes.js";
import { getGridMetrics } from "../gridMetrics.js";
import {
  beginWorkbookGridEdit,
  openWorkbookGridHeaderContextMenuFromKeyboard,
  selectEntireWorkbookSheet,
  toggleWorkbookGridBooleanCell,
} from "../gridInteractionCommands.js";

describe("gridInteractionCommands", () => {
  it("should begin editing with the explicit seed or current cell seed", () => {
    // Arrange
    const onBeginEdit = vi.fn();
    const engine = {
      getCell: vi.fn(() => ({
        sheetName: "Sheet1",
        address: "A1",
        value: { tag: 3, value: "from-cell", stringId: 1 },
        flags: 0,
        version: 1,
      })),
    };

    // Act
    beginWorkbookGridEdit({
      engine,
      onBeginEdit,
      sheetName: "Sheet1",
      address: "A1",
      seed: "typed",
    });
    beginWorkbookGridEdit({
      engine,
      onBeginEdit,
      sheetName: "Sheet1",
      address: "A1",
    });

    // Assert
    expect(onBeginEdit).toHaveBeenNthCalledWith(1, "typed", "caret-end");
    expect(onBeginEdit).toHaveBeenNthCalledWith(2, "from-cell", "caret-end");
  });

  it("should toggle boolean cells and ignore non-boolean cells", () => {
    // Arrange
    const onToggleBooleanCell = vi.fn();
    const engine = {
      getCell: vi
        .fn()
        .mockReturnValueOnce({
          sheetName: "Sheet1",
          address: "B2",
          value: { tag: 2, value: true },
          flags: 0,
          version: 1,
        })
        .mockReturnValueOnce({
          sheetName: "Sheet1",
          address: "C3",
          value: { tag: 3, value: "text", stringId: 1 },
          flags: 0,
          version: 1,
        }),
    };

    // Act
    expect(
      toggleWorkbookGridBooleanCell({
        engine,
        onToggleBooleanCell,
        sheetName: "Sheet1",
        col: 1,
        row: 1,
      }),
    ).toBe(true);
    expect(
      toggleWorkbookGridBooleanCell({
        engine,
        onToggleBooleanCell,
        sheetName: "Sheet1",
        col: 2,
        row: 2,
      }),
    ).toBe(false);

    // Assert
    expect(onToggleBooleanCell).toHaveBeenCalledTimes(1);
    expect(onToggleBooleanCell).toHaveBeenCalledWith("Sheet1", "B2", false);
  });

  it("should open the keyboard header context menu against the resolved header target", () => {
    // Arrange
    const openContextMenuForTarget = vi.fn(() => true);

    // Act
    const opened = openWorkbookGridHeaderContextMenuFromKeyboard({
      hostBounds: { left: 100, top: 200 },
      gridSelection: {
        columns: CompactSelection.fromSingleSelection(2),
        rows: CompactSelection.empty(),
        current: undefined,
      },
      selectedCell: [2, 3],
      getCellScreenBounds: () => ({ x: 300, y: 400, width: 104, height: 22 }),
      gridMetrics: getGridMetrics(),
      openContextMenuForTarget,
    });

    // Assert
    expect(opened).toBe(true);
    expect(openContextMenuForTarget).toHaveBeenCalledWith({
      target: { kind: "column", index: 2 },
      x: 352,
      y: 212,
    });
  });

  it("should select the entire sheet and commit editing first when needed", () => {
    // Arrange
    const onCommitEdit = vi.fn();
    const setGridSelection = vi.fn();
    const onSelect = vi.fn();
    const focusGrid = vi.fn();

    // Act
    selectEntireWorkbookSheet({
      isEditingCell: true,
      onCommitEdit,
      setGridSelection,
      onSelect,
      focusGrid,
    });

    // Assert
    expect(onCommitEdit).toHaveBeenCalledTimes(1);
    expect(setGridSelection).toHaveBeenCalledTimes(1);
    expect(onSelect).toHaveBeenCalledWith("A1");
    expect(focusGrid).toHaveBeenCalledTimes(1);
  });
});
